import customtkinter as ctk
import threading
import sys
import urllib.parse
import requests
import re
import time
import json
import os
import io
import win32com.client
from pypresence import Presence
from PIL import Image, ImageDraw
import pystray 

# --- CONFIGURATION DEFAULTS ---
DEFAULT_CONFIG = {
    "client_id": "1462375131782447321",
    "update_interval": 5,
    "show_buttons": True,
    "minimize_to_tray": True,
    "auto_connect": True,
    "start_minimized": True
}

# --- THEME SETUP ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class ModernBridgeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 1. Load Configuration
        self.config = self.load_config()

        # Window Setup
        self.title("MediaMonkey 5 Bridge (v1.4)")
        self.geometry("500x700") 
        self.resizable(False, False)

        # 2. LOAD CUSTOM ICON (Window & Tray)
        self.icon_path = resource_path("logo.ico")
        
        # A. Set Window Icon (Taskbar/Titlebar)
        if os.path.exists(self.icon_path):
            try:
                self.iconbitmap(self.icon_path)
            except: pass

        # B. Prepare Tray Icon Image
        try:
            if os.path.exists(self.icon_path):
                self.tray_image = Image.open(self.icon_path)
            else:
                raise FileNotFoundError
        except:
            # Fallback: Create an orange square if no logo found
            self.tray_image = Image.new('RGB', (64, 64), color=(255, 165, 0))
            d = ImageDraw.Draw(self.tray_image)
            d.rectangle((16, 16, 48, 48), fill=(255, 255, 255))

        # State Variables
        self.rpc = None
        self.mm = None
        self.last_track = ""
        self.is_running = False
        self.current_art_url = "logo"
        self.tray_icon = None

        # Default Art Placeholder
        self.default_art = ctk.CTkImage(
            light_image=Image.new("RGB", (200, 200), (30, 30, 30)),
            dark_image=Image.new("RGB", (200, 200), (30, 30, 30)),
            size=(200, 200)
        )

        # --- GRID LAYOUT ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # HEADER CARD
        self.header_frame = ctk.CTkFrame(self, corner_radius=15, fg_color="#1e1e2e")
        self.header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        self.status_indicator = ctk.CTkLabel(
            self.header_frame, 
            text="● DISCONNECTED", 
            text_color="#ED4245", 
            font=("Roboto Medium", 12)
        )
        self.status_indicator.pack(pady=(15, 10))

        # ALBUM ART
        self.art_label = ctk.CTkLabel(self.header_frame, text="", image=self.default_art)
        self.art_label.pack(pady=(0, 15))

        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="Ready to Connect", 
            font=("Roboto", 20, "bold"),
            text_color="white",
            wraplength=400
        )
        self.title_label.pack(pady=5)

        self.artist_label = ctk.CTkLabel(
            self.header_frame, 
            text="Open MediaMonkey to begin", 
            font=("Roboto", 14), 
            text_color="#a6a6a6"
        )
        self.artist_label.pack(pady=(0, 20))

        # CONTROLS
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.grid(row=1, column=0, padx=20, pady=0, sticky="ew")
        
        self.start_btn = ctk.CTkButton(
            self.btn_frame, 
            text="START BRIDGE", 
            command=self.start_bridge,
            font=("Roboto", 14, "bold"),
            fg_color="#5865F2", 
            hover_color="#4752C4",
            height=40,
            corner_radius=20
        )
        self.start_btn.pack(side="left", expand=True, padx=5, fill="x")

        self.stop_btn = ctk.CTkButton(
            self.btn_frame, 
            text="STOP", 
            command=self.stop_bridge,
            font=("Roboto", 14, "bold"),
            fg_color="#2b2d31",
            hover_color="#ed4245",
            state="disabled",
            height=40,
            corner_radius=20
        )
        self.stop_btn.pack(side="right", expand=True, padx=5, fill="x")

        # LOGGING
        self.log_label = ctk.CTkLabel(self, text="Connection Log", font=("Roboto", 12, "bold"), text_color="#a6a6a6")
        self.log_label.grid(row=2, column=0, padx=25, pady=(20, 0), sticky="w")

        self.log_area = ctk.CTkTextbox(
            self, 
            fg_color="#11111b", 
            text_color="#00ff9d", 
            font=("Consolas", 11),
            corner_radius=10
        )
        self.log_area.grid(row=3, column=0, padx=20, pady=(5, 20), sticky="nsew")

        self.protocol("WM_DELETE_WINDOW", self.on_close_attempt)

        # --- AUTO START LOGIC ---
        if self.config.get("start_minimized", False):
            self.withdraw()
            threading.Thread(target=self.create_tray_icon, daemon=True).start()
            self.log("Started minimized in System Tray.")
        
        if self.config.get("auto_connect", False):
            self.after(1000, self.start_bridge)

    # --- CONFIGURATION ---
    def load_config(self):
        config = DEFAULT_CONFIG.copy()
        if os.path.exists("config.json"):
            try:
                with open("config.json", "r") as f:
                    config.update(json.load(f))
            except: pass
        
        with open("config.json", "w") as f:
            json.dump(config, f, indent=4)
        return config

    # --- TRAY ---
    def create_tray_icon(self):
        menu = pystray.Menu(pystray.MenuItem("Open", self.restore_window), pystray.MenuItem("Quit", self.quit_app))
        self.tray_icon = pystray.Icon("MM5Bridge", self.tray_image, "MM5 Bridge", menu)
        self.tray_icon.run()

    def on_close_attempt(self):
        if self.config["minimize_to_tray"]:
            self.withdraw()
            threading.Thread(target=self.create_tray_icon, daemon=True).start()
            self.log("Minimized to System Tray.")
        else:
            self.quit_app()

    def restore_window(self, icon, item):
        self.tray_icon.stop()
        self.after(0, self.deiconify)

    def quit_app(self, icon=None, item=None):
        if self.tray_icon: self.tray_icon.stop()
        self.is_running = False
        self.quit()
        sys.exit()

    # --- CORE LOGIC ---
    def log(self, message):
        self.log_area.insert("end", f">> {message}\n")
        self.log_area.see("end")

    def set_status(self, status, color):
        self.status_indicator.configure(text=f"● {status}", text_color=color)

    def start_bridge(self):
        if self.is_running: return
        self.log("Initializing Bridge...")
        try:
            self.rpc = Presence(self.config["client_id"])
            self.rpc.connect()
            self.is_running = True
            
            self.start_btn.configure(state="disabled", fg_color="#2b2d31")
            self.stop_btn.configure(state="normal", fg_color="#ED4245")
            self.set_status("CONNECTED", "#57F287")
            self.log("Discord Connected. Listening for MediaMonkey...")
            self.poll_mediamonkey() 
        except Exception as e:
            self.log(f"Connection Error: {e}")
            self.set_status("ERROR", "#ED4245")

    def stop_bridge(self):
        self.is_running = False
        self.start_btn.configure(state="normal", fg_color="#5865F2")
        self.stop_btn.configure(state="disabled", fg_color="#2b2d31")
        self.set_status("DISCONNECTED", "#ED4245")
        self.title_label.configure(text="Ready to Connect")
        self.artist_label.configure(text="---")
        self.art_label.configure(image=self.default_art)
        if self.rpc:
            try: self.rpc.clear()
            except: pass
        self.log("Bridge stopped.")

    def clean_string(self, text):
        return re.sub(r"[\(\[].*?[\)\]]", "", text).strip()

    def fetch_album_art(self, artist, album):
        try:
            clean_artist = self.clean_string(artist)
            clean_album = self.clean_string(album)
            query = urllib.parse.quote(f"{clean_artist} {clean_album}")
            url = f"https://itunes.apple.com/search?term={query}&media=music&entity=album&limit=1"
            response = requests.get(url, timeout=2)
            data = response.json()
            
            if data['resultCount'] > 0:
                art_url = data['results'][0]['artworkUrl100'].replace("100x100", "512x512")
                self.current_art_url = art_url
                self.log(f"Art Found: {clean_album}")

                img_data = requests.get(art_url).content
                pil_img = Image.open(io.BytesIO(img_data))
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(200, 200))
                self.after(0, lambda: self.art_label.configure(image=ctk_img))
            else:
                self.current_art_url = "logo"
                self.after(0, lambda: self.art_label.configure(image=self.default_art))
        except:
            self.current_art_url = "logo"
            self.after(0, lambda: self.art_label.configure(image=self.default_art))

    def update_discord(self, artist, title, album, start_time=None):
        if not self.rpc: return
        try:
            btns = None
            if self.config["show_buttons"]:
                yt_query = urllib.parse.quote(f"{artist} - {title}")
                btns = [
                    {"label": "Listen on YouTube", "url": f"https://www.youtube.com/results?search_query={yt_query}"},
                    {"label": "Search Apple Music", "url": f"https://music.apple.com/us/search?term={yt_query}"}
                ]

            self.rpc.update(
                state=f"by {artist}",
                details=f"{title}",
                large_image=self.current_art_url,
                large_text=album,
                small_image="play",
                small_text="Playing",
                start=start_time,
                buttons=btns
            )
        except: pass

    def poll_mediamonkey(self):
        if not self.is_running: return
        try:
            if self.mm is None:
                self.mm = win32com.client.Dispatch("SongsDB5.SDBApplication")
        except:
            self.mm = None 
            self.set_status("SEARCHING FOR MM5...", "#FEE75C")

        try:
            if self.mm and self.mm.Player.IsPlaying:
                song = self.mm.Player.CurrentSong
                if song:
                    current_pos_ms = self.mm.Player.PlaybackTime
                    current_pos_sec = current_pos_ms / 1000
                    start_timestamp = int(time.time() - current_pos_sec)
                    track_key = f"{song.ArtistName} - {song.Title}"
                    
                    if track_key != self.last_track:
                        self.last_track = track_key
                        self.log(f"Now Playing: {track_key}")
                        self.set_status("BROADCASTING", "#57F287")
                        self.title_label.configure(text=song.Title)
                        self.artist_label.configure(text=song.ArtistName)
                        self.current_art_url = "logo"
                        threading.Thread(target=self.fetch_album_art, args=(song.ArtistName, song.AlbumName), daemon=True).start()
                    
                    self.update_discord(song.ArtistName, song.Title, song.AlbumName, start_time=start_timestamp)
            else:
                if self.last_track != "PAUSED":
                    try: self.rpc.clear()
                    except: pass
                    self.last_track = "PAUSED"
                    self.set_status("IDLE", "#FEE75C")
                    self.title_label.configure(text="Paused")
                    self.art_label.configure(image=self.default_art)
        except:
            self.mm = None
        self.after(self.config["update_interval"] * 1000, self.poll_mediamonkey)

if __name__ == "__main__":
    app = ModernBridgeApp()
    app.mainloop()