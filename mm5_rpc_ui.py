import tkinter as tk
import pythoncom  # <--- ADD THIS
from tkinter import scrolledtext
# ... rest of imports
import tkinter as tk
from tkinter import scrolledtext
import threading
import time
import sys
import win32com.client
from pypresence import Presence
import requests

# --- CONFIGURATION ---
CLIENT_ID = '1462375131782447321' # <--- PASTE ID HERE
USER_AGENT = 'MM5-DiscordRPC/1.0 ( contact@example.com )'

class MM5RPCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MM5 Discord Bridge")
        self.root.geometry("450x350")
        self.root.configure(bg="#2C2F33") # Discord Dark Grey

        # Threading Flags
        self.running = False
        self.thread = None

        # --- GUI ELEMENTS ---
        
        # Header / Status
        self.status_label = tk.Label(root, text="Status: STOPPED", fg="#ED4245", bg="#2C2F33", font=("Segoe UI", 12, "bold"))
        self.status_label.pack(pady=10)

        # Current Track Info
        self.track_label = tk.Label(root, text="Waiting for playback...", fg="#FFFFFF", bg="#2C2F33", font=("Segoe UI", 10))
        self.track_label.pack(pady=5)

        # Buttons Frame
        btn_frame = tk.Frame(root, bg="#2C2F33")
        btn_frame.pack(pady=10)

        self.start_btn = tk.Button(btn_frame, text="Start Bridge", command=self.start_rpc, bg="#5865F2", fg="white", width=15, font=("Segoe UI", 9, "bold"), relief="flat")
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_rpc, bg="#ED4245", fg="white", width=15, font=("Segoe UI", 9, "bold"), relief="flat", state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # Console Log
        self.log_area = scrolledtext.ScrolledText(root, height=10, bg="#23272A", fg="#00FF00", font=("Consolas", 9))
        self.log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.log_area.insert(tk.END, ">> Ready to connect...\n")

        # Handle Window Close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, message):
        """Thread-safe logging to the text box"""
        self.root.after(0, self._log_internal, message)

    def _log_internal(self, message):
        self.log_area.insert(tk.END, f">> {message}\n")
        self.log_area.see(tk.END)

    def update_ui_status(self, text, color, song_info=""):
        """Thread-safe UI updates"""
        self.root.after(0, lambda: self.status_label.config(text=text, fg=color))
        if song_info:
            self.root.after(0, lambda: self.track_label.config(text=song_info))

    def start_rpc(self):
        if not self.running:
            self.running = True
            self.start_btn.config(state=tk.DISABLED, bg="#4f545c")
            self.stop_btn.config(state=tk.NORMAL, bg="#ED4245")
            self.thread = threading.Thread(target=self.rpc_worker, daemon=True)
            self.thread.start()

    def stop_rpc(self):
        if self.running:
            self.log("Stopping background thread...")
            self.running = False
            self.start_btn.config(state=tk.NORMAL, bg="#5865F2")
            self.stop_btn.config(state=tk.DISABLED, bg="#4f545c")
            self.update_ui_status("Status: STOPPED", "#ED4245", "Bridge halted.")

    def on_close(self):
        """Clean shutdown when X is clicked"""
        self.running = False
        self.root.destroy()
        sys.exit()

    # --- WORKER THREAD LOGIC ---
    def rpc_worker(self):
        pythoncom.CoInitialize() # <--- THIS IS THE CRITICAL FIX
        self.log("Initializing RPC handshake...")
        try:
            rpc = Presence(CLIENT_ID)
            rpc.connect()
            self.log("Connected to Discord IPC.")
            self.update_ui_status("Status: ACTIVE", "#57F287")
        except Exception as e:
            self.log(f"Discord Error: {e}")
            self.stop_rpc()
            return

        mm = None
        last_track = ""
        art_cache = {}

        while self.running:
            try:
                # 1. Connect to MM5
                if mm is None:
                    try:
                        mm = win32com.client.Dispatch("SongsDB5.SDBApplication")
                        _ = mm.Player.IsPlaying # Test access
                    except:
                        mm = None
                        # Don't spam logs, just wait
                        time.sleep(3)
                        continue

                player = mm.Player
                
                if player.IsPlaying:
                    song = player.CurrentSong
                    if not song: continue

                    track_key = f"{song.ArtistName} - {song.Title}"
                    
                    # Update Logic
                    if track_key != last_track:
                        last_track = track_key
                        self.log(f"Now Playing: {track_key}")
                        self.update_ui_status("Status: BROADCASTING", "#57F287", track_key)

                        # Quick iTunes/MusicBrainz Art Fetch (Simplified for brevity)
                        # [Insert your art fetching function here if desired]
                        # For UI version, we'll default to 'logo' to keep it responsive
                        current_art = "logo"

                        rpc.update(
                            state=f"by {song.ArtistName}",
                            details=f"{song.Title}",
                            large_image=current_art,
                            small_image="play",
                            small_text="Playing"
                        )
                else:
                    rpc.clear()
                    self.update_ui_status("Status: IDLE", "#FEE75C", "Paused / Stopped")

                time.sleep(5) # Poll rate

            except Exception as e:
                self.log(f"Error: {e}")
                mm = None
                time.sleep(5)

# --- ENTRY POINT ---
if __name__ == "__main__":
    root = tk.Tk()
    app = MM5RPCApp(root)
    root.mainloop()