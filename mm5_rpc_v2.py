import tkinter as tk
from tkinter import scrolledtext
import sys
import time
import win32com.client
from pypresence import Presence

# --- CONFIGURATION ---
CLIENT_ID = '1462375131782447321' # Your App ID

class MM5RPCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MM5 Discord Bridge (v2)")
        self.root.geometry("450x350")
        self.root.configure(bg="#2C2F33")

        # State Variables
        self.rpc = None
        self.mm = None
        self.last_track = ""
        self.is_running = False

        # --- UI ELEMENTS ---
        self.status_label = tk.Label(root, text="Status: READY", fg="#99AAB5", bg="#2C2F33", font=("Segoe UI", 12, "bold"))
        self.status_label.pack(pady=10)

        self.track_label = tk.Label(root, text="Click Start to connect", fg="#FFFFFF", bg="#2C2F33", font=("Segoe UI", 10))
        self.track_label.pack(pady=5)

        btn_frame = tk.Frame(root, bg="#2C2F33")
        btn_frame.pack(pady=10)

        self.start_btn = tk.Button(btn_frame, text="Start Bridge", command=self.start_bridge, bg="#5865F2", fg="white", width=15, font=("Segoe UI", 9, "bold"), relief="flat")
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = tk.Button(btn_frame, text="Stop", command=self.stop_bridge, bg="#ED4245", fg="white", width=15, font=("Segoe UI", 9, "bold"), relief="flat", state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.log_area = scrolledtext.ScrolledText(root, height=10, bg="#23272A", fg="#00FF00", font=("Consolas", 9))
        self.log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Handle Window Close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log(self, message):
        self.log_area.insert(tk.END, f">> {message}\n")
        self.log_area.see(tk.END)

    def start_bridge(self):
        if self.is_running: return
        
        self.log("Connecting to Discord...")
        try:
            self.rpc = Presence(CLIENT_ID)
            self.rpc.connect()
            self.is_running = True
            self.start_btn.config(state=tk.DISABLED, bg="#4f545c")
            self.stop_btn.config(state=tk.NORMAL, bg="#ED4245")
            self.status_label.config(text="Status: ACTIVE", fg="#57F287")
            self.log("Connected! Polling MediaMonkey...")
            
            # START THE POLLING LOOP
            self.poll_mediamonkey() 
            
        except Exception as e:
            self.log(f"Discord Connection Failed: {e}")

    def stop_bridge(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL, bg="#5865F2")
        self.stop_btn.config(state=tk.DISABLED, bg="#4f545c")
        self.status_label.config(text="Status: STOPPED", fg="#ED4245")
        if self.rpc:
            try: self.rpc.clear()
            except: pass
        self.log("Bridge stopped.")

    def poll_mediamonkey(self):
        """This runs every 5 seconds on the MAIN THREAD"""
        if not self.is_running: return

        # 1. Connect/Reconnect to MediaMonkey
        try:
            if self.mm is None:
                # --- CRITICAL FIX: Use Dispatch instead of GetActiveObject ---
                self.mm = win32com.client.Dispatch("SongsDB5.SDBApplication")
        except Exception:
            self.mm = None 
            self.status_label.config(text="Status: SEARCHING...", fg="#FEE75C")
            self.track_label.config(text="Open MediaMonkey...")

        # 2. Check Playback
        try:
            if self.mm:
                if self.mm.Player.IsPlaying:
                    song = self.mm.Player.CurrentSong
                    if song:
                        track_key = f"{song.ArtistName} - {song.Title}"
                        
                        # Only update Discord if song changed
                        if track_key != self.last_track:
                            self.last_track = track_key
                            self.log(f"Playing: {track_key}")
                            self.status_label.config(text="Status: BROADCASTING", fg="#57F287")
                            self.track_label.config(text=track_key)
                            
                            self.rpc.update(
                                state=f"by {song.ArtistName}",
                                details=f"{song.Title}",
                                large_image="logo",
                                small_image="play",
                                small_text="Playing"
                            )
                else:
                    # Paused logic
                    if self.last_track != "PAUSED":
                        self.rpc.clear()
                        self.last_track = "PAUSED"
                        self.status_label.config(text="Status: IDLE", fg="#FEE75C")
                        self.track_label.config(text="Paused / Stopped")

        except Exception as e:
            self.log(f"Error reading MM5: {e}")
            self.mm = None # Force reconnect next loop

        # Schedule this function to run again in 5000ms (5 seconds)
        self.root.after(5000, self.poll_mediamonkey)

    def on_close(self):
        self.is_running = False
        self.root.destroy()
        sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = MM5RPCApp(root)
    root.mainloop()