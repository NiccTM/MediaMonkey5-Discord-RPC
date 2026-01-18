import time
import requests
import win32com.client
from pypresence import Presence

# --- CONFIGURATION ---
CLIENT_ID = '1462375131782447321'  # <--- PASTE YOUR ID HERE
USER_AGENT = 'MM5-DiscordRPC/1.0 ( myemail@example.com )' # MusicBrainz STRICTLY requires a contact info UA

# --- CACHE & STATE ---
art_cache = {} 

def get_cover_musicbrainz(artist, album):
    """
    Queries MusicBrainz for the Release Group ID, then fetches the front cover
    from the Cover Art Archive. Returns a direct HTTPS URL.
    """
    try:
        # 1. Search for the release
        search_url = "https://musicbrainz.org/ws/2/release/"
        headers = {'User-Agent': USER_AGENT}
        params = {
            'query': f'artist:"{artist}" AND release:"{album}"',
            'fmt': 'json',
            'limit': 1
        }
        
        resp = requests.get(search_url, headers=headers, params=params, timeout=5)
        if resp.status_code != 200: return None
        
        data = resp.json()
        if not data.get('releases'): return None
        
        # 2. Extract MBID and construct the image URL
        mbid = data['releases'][0]['id']
        return f"https://coverartarchive.org/release/{mbid}/front-500"
        
    except Exception as e:
        # Fail silently to keep the loop running; return None to trigger fallback
        return None

def main():
    print("--- MediaMonkey 5 Discord RPC Started ---")
    
    # Initialize Discord Connection
    rpc = Presence(CLIENT_ID)
    try:
        rpc.connect()
        print("Connected to Discord IPC.")
    except Exception:
        print("Error: Could not connect to Discord. Is the app running?")
        return

    mm = None
    last_track_key = ""
    current_art_url = "logo" # Default to uploaded asset key
    start_time = None

    while True:
        try:
            # 1. Hook into MediaMonkey COM Object
            if mm is None:
                try:
                    mm = win32com.client.Dispatch("SongsDB5.SDBApplication")
                    # Simple check to see if the object is alive
                    _ = mm.Player.IsPlaying
                except Exception:
                    mm = None
                    time.sleep(5)
                    continue

            player = mm.Player
            
            # 2. Check Playback State
            if player.IsPlaying:
                song = player.CurrentSong
                if not song: continue

                # Create unique ID for current track to manage state
                track_key = f"{song.ArtistName}_{song.AlbumName}"
                
                # 3. Handle Metadata Changes
                if track_key != last_track_key:
                    last_track_key = track_key
                    start_time = time.time()
                    print(f"Now Playing: {song.Title} - {song.ArtistName}")

                    # Check RAM Cache first
                    if track_key in art_cache:
                        current_art_url = art_cache[track_key]
                    else:
                        # Fetch from API
                        fetched_url = get_cover_musicbrainz(song.ArtistName, song.AlbumName)
                        if fetched_url:
                            current_art_url = fetched_url
                            art_cache[track_key] = fetched_url # Store in RAM
                        else:
                            current_art_url = "logo" # Fallback

                # 4. Push Update to Discord
                rpc.update(
                    state=f"by {song.ArtistName}",
                    details=f"{song.Title}",
                    large_image=current_art_url,
                    large_text=song.AlbumName,
                    start=start_time,
                    small_image="play",
                    small_text="Playing"
                )
            else:
                # Clear status if paused/stopped
                rpc.clear()
                last_track_key = "" # Reset so resuming triggers an update

            # Poll every 15s (Discord's rate limit is 15s for status updates anyway)
            time.sleep(15)

        except KeyboardInterrupt:
            print("Stopping...")
            break
        except Exception as e:
            print(f"Connection lost: {e}")
            mm = None
            time.sleep(5)

if __name__ == "__main__":
    main()