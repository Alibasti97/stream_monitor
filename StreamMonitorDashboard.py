import tkinter as tk
from tkinter import scrolledtext
import threading
import sys


class StreamMonitorDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Stream Monitor Dashboard")
        self.root.geometry("450x500")

        self.monitor_thread = None
        self.stop_event = threading.Event()

        # --- Stream Tester URL ---
        tk.Label(root, text="Stream Tester URL:").pack(anchor='w', padx=10)
        self.tester_url = tk.Entry(root, width=80)
        self.tester_url.pack(padx=10, pady=5)
        self.tester_url.insert(0, "https://www.radiantmediaplayer.com/stream-tester.html")

        # --- Streaming URL ---
        tk.Label(root, text="Stream URL:").pack(anchor='w', padx=10)
        self.stream_url = tk.Entry(root, width=80)
        self.stream_url.pack(padx=10, pady=5)
        self.stream_url.insert(0, "https://cdn01khi.tamashaweb.com:8087/YlUHeDQb7a/city41News-abr/playlist.m3u8")

        # --- Buttons ---
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        self.start_button = tk.Button(btn_frame, text="Start", command=self.start_monitoring, bg='green', fg='white', width=10)
        self.start_button.pack(side='left', padx=10)

        self.stop_button = tk.Button(btn_frame, text="Stop", command=self.stop_monitoring, bg='red', fg='white', width=10)
        self.stop_button.pack(side='left', padx=10)

        # --- Log Output ---
        tk.Label(root, text="Log Output:").pack(anchor='w', padx=10)
        self.log_area = scrolledtext.ScrolledText(root, width=85, height=20, state='disabled')
        self.log_area.pack(padx=10, pady=5)

        sys.stdout = TextRedirector(self.log_area, "stdout")

    def start_monitoring(self):
        if self.monitor_thread and self.monitor_thread.is_alive():
            print("[INFO] Monitoring already running.")
            return

        self.stop_event.clear()

        tester_link = self.tester_url.get()
        stream_link = self.stream_url.get()

        print(f"[INFO] Starting monitor for: {stream_link}")
        self.monitor_thread = threading.Thread(target=self.run_monitor, args=(tester_link, stream_link), daemon=True)
        self.monitor_thread.start()

    def run_monitor(self, tester_link, stream_link):
        import main  # Import your main.py logic
        main.main(tester_link, stream_link, self.stop_event)

    def stop_monitoring(self):
        print("[INFO] Sending stop signal to monitoring...")
        self.stop_event.set()


class TextRedirector:
    def __init__(self, widget, tag):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state='normal')
        self.widget.insert("end", str)
        self.widget.configure(state='disabled')
        self.widget.see("end")

    def flush(self):
        pass


if __name__ == "__main__":
    root = tk.Tk()
    app = StreamMonitorDashboard(root)
    root.mainloop()
