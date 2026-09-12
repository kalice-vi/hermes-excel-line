
import os, socket, subprocess, time
def ensure_brain():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.5)
        s.connect(("127.0.0.1", 8766))
        s.close()
    except Exception:
        # Restart brain_server.py
        script_dir = os.path.dirname(os.path.abspath(__file__))
        server = os.path.join(script_dir, "scripts", "brain_server.py")
        subprocess.Popen(["python", server], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
if __name__ == "__main__":
    ensure_brain()
