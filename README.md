# MediaMonkey 5 Discord RPC Bridge

A standalone middleware application that integrates MediaMonkey 5 playback state with the Discord Rich Presence (RPC) API.

## Technical Overview
MediaMonkey 5 operates in a sandboxed Chromium environment, which limits the stability of internal network requests. This project decouples the RPC logic into an external process, ensuring stability and zero overhead on the media player itself.

### Architecture
The application acts as a bridge between two Inter-Process Communication (IPC) protocols:
1.  **Input (COM):** Polls the `SongsDB5.SDBApplication` Windows COM interface to retrieve real-time telemetry (Track, Artist, Playback State) from the active MediaMonkey process.
2.  **Output (IPC):** Formats the telemetry into a JSON payload and transmits it to the local Discord client via Unix Domain Sockets (or Named Pipes on Windows) using the `pypresence` library.

### Implementation Details
* **Language:** Python 3.14
* **Concurrency:** Single-threaded event loop (Tkinter) with optimized polling (5000ms interval) to prevent race conditions and minimize CPU usage.
* **Distribution:** Compiled to PE format (Portable Executable) via PyInstaller, bundling the Python runtime for zero-dependency execution.

## Installation

### Option 1: Standalone Binary (Recommended)
1.  Download `MM5_Bridge_V1.3_Stable.exe` from the [Releases Page](https://github.com/NiccTM/MediaMonkey5-Discord-RPC/releases).
2.  Run the executable while MediaMonkey 5 is open.
3.  The bridge will automatically attach to the active process PID.

### Option 2: Run from Source
```bash
git clone [https://github.com/NiccTM/MediaMonkey5-Discord-RPC.git](https://github.com/NiccTM/MediaMonkey5-Discord-RPC.git)
pip install pypresence pywin32
python mm5_rpc_v1.3.py
