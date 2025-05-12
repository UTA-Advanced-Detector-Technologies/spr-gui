import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyvisa
import time
import threading
import csv
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

class PowerLoggerApp:
    def __init__(self, master):
        self.master = master
        master.title("LED Pulse and Power Logger")
        master.geometry("800x900")

        self.running = False

        self.create_widgets()

        self.rm = pyvisa.ResourceManager()
        self.dc2200 = None
        self.pm101 = None
        self.device_map = {}
        self.init_devices()

        self.power_log = []
        self.time_log = []
        self.led_transitions = []

    def create_widgets(self):
        settings_frame = ttk.LabelFrame(self.master, text="Pulse Settings")
        settings_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(settings_frame, text="LED ON Duration (s):").grid(row=0, column=0)
        self.on_entry = ttk.Entry(settings_frame, width=10)
        self.on_entry.insert(0, "1.0")
        self.on_entry.grid(row=0, column=1, padx=5)

        ttk.Label(settings_frame, text="LED OFF Duration (s):").grid(row=0, column=2)
        self.off_entry = ttk.Entry(settings_frame, width=10)
        self.off_entry.insert(0, "1.0")
        self.off_entry.grid(row=0, column=3, padx=5)

        ttk.Label(settings_frame, text="Cycles:").grid(row=0, column=4)
        self.cycles_entry = ttk.Entry(settings_frame, width=10)
        self.cycles_entry.insert(0, "10")
        self.cycles_entry.grid(row=0, column=5, padx=5)

        filename_frame = ttk.LabelFrame(self.master, text="CSV File Name")
        filename_frame.pack(fill="x", padx=10, pady=5)
        self.filename_entry = ttk.Entry(filename_frame)
        self.filename_entry.insert(0, "gui_power_log")
        self.filename_entry.pack(fill="x", padx=10, pady=5)

        intensity_frame = ttk.LabelFrame(self.master, text="LED Intensity (%)")
        intensity_frame.pack(fill="x", padx=10, pady=5)
        self.intensity_slider = ttk.Scale(intensity_frame, from_=0, to=100, orient=tk.HORIZONTAL)
        self.intensity_slider.set(50)
        self.intensity_slider.pack(fill="x", padx=10, pady=5)
        self.intensity_slider.bind("<ButtonRelease-1>", self.apply_intensity)

        self.intensity_label = ttk.Label(intensity_frame, text="Current Intensity: 50.0%")
        self.intensity_label.pack(pady=2)

        rate_frame = ttk.LabelFrame(self.master, text="PM101 Sampling Rate")
        rate_frame.pack(fill="x", padx=10, pady=5)
        self.rate_choice = tk.StringVar()
        self.rate_choice.set("Fast")
        ttk.Label(rate_frame, text="Rate:").pack(side=tk.LEFT, padx=5)
        rate_menu = ttk.OptionMenu(rate_frame, self.rate_choice, "Fast", "Fast", "Medium", "Slow")
        rate_menu.pack(side=tk.LEFT, padx=5)

        ttk.Label(self.master, text="Power (W):").pack(pady=5)
        self.power_label = ttk.Label(self.master, text="N/A", font=("Helvetica", 16))
        self.power_label.pack()

        self.led_state_label = ttk.Label(self.master, text="LED State: OFF", font=("Helvetica", 12))
        self.led_state_label.pack(pady=5)

        btn_frame = ttk.Frame(self.master)
        btn_frame.pack(pady=10)

        self.start_button = ttk.Button(btn_frame, text="Start", command=self.start_logging)
        self.start_button.grid(row=0, column=0, padx=10)

        self.stop_button = ttk.Button(btn_frame, text="Stop", command=self.stop_logging, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=10)

        self.refresh_button = ttk.Button(btn_frame, text="Refresh Devices", command=self.init_devices)
        self.refresh_button.grid(row=0, column=2, padx=10)

        self.save_button = ttk.Button(btn_frame, text="Save Plot", command=self.save_plot)
        self.save_button.grid(row=0, column=3, padx=10)

        self.status_label = ttk.Label(self.master, text="Status: Ready")
        self.status_label.pack(pady=5)

        self.device_listbox = tk.Listbox(self.master, height=4)
        self.device_listbox.pack(fill="x", padx=10)

        self.figure, self.ax = plt.subplots(figsize=(6, 3))
        self.ax.set_title("Power vs Time")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Power (W)")
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.master)
        self.canvas.get_tk_widget().pack(pady=10)

    def apply_intensity(self, event):
        intensity = float(self.intensity_slider.get())
        self.intensity_label.config(text=f"Current Intensity: {intensity:.1f}%")
        try:
            if self.dc2200:
                self.dc2200.write(f"SOUR:CURR {intensity:.1f}")
                self.status_label.config(text=f"LED Intensity Set To: {intensity:.1f}%")
        except Exception as e:
            self.status_label.config(text=f"Failed to set intensity: {e}")

    def pulse_and_record(self):
        try:
            on_duration = float(self.on_entry.get())
            off_duration = float(self.off_entry.get())
            cycles = int(self.cycles_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values for pulse settings.")
            self.stop_logging()
            return

        base_filename = self.filename_entry.get().strip()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        script_dir = os.path.dirname(os.path.abspath(__file__))
        full_filename = os.path.join(script_dir, f"{base_filename}_{timestamp}.csv")

        with open(full_filename, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Timestamp", "Power (W)", "LED State"])

            try:
                for _ in range(cycles):
                    if not self.running:
                        break

                    self.dc2200.write("OUTP:STAT ON")
                    timestamp = datetime.now()
                    self.led_state_label.config(text="LED State: ON")
                    writer.writerow([timestamp.isoformat(), "", "ON"])
                    self.led_transitions.append((timestamp, 'ON'))

                    start = time.time()
                    while time.time() - start < on_duration:
                        if not self.running: break
                        power = float(self.pm101.query("MEAS:SCAL:POW?"))
                        timestamp = datetime.now()
                        self.power_label.config(text=f"{power:.4e}")
                        writer.writerow([timestamp.isoformat(), power, ""])
                        self.time_log.append(timestamp)
                        self.power_log.append(power)
                        self.update_plot()

                        rate = self.rate_choice.get()
                        if rate == "Medium":
                            time.sleep(0.01)
                        elif rate == "Slow":
                            time.sleep(0.1)

                    self.dc2200.write("OUTP:STAT OFF")
                    timestamp = datetime.now()
                    self.led_state_label.config(text="LED State: OFF")
                    writer.writerow([timestamp.isoformat(), "", "OFF"])
                    self.led_transitions.append((timestamp, 'OFF'))

                    start = time.time()
                    while time.time() - start < off_duration:
                        if not self.running: break
                        power = float(self.pm101.query("MEAS:SCAL:POW?"))
                        timestamp = datetime.now()
                        self.power_label.config(text=f"{power:.4e}")
                        writer.writerow([timestamp.isoformat(), power, ""])
                        self.time_log.append(timestamp)
                        self.power_log.append(power)
                        self.update_plot()

                        rate = self.rate_choice.get()
                        if rate == "Medium":
                            time.sleep(0.01)
                        elif rate == "Slow":
                            time.sleep(0.1)

                self.status_label.config(text=f"Status: Completed. Saved as {full_filename}")

            except Exception as e:
                messagebox.showerror("Error", f"Logging failed:\n{e}")
                self.status_label.config(text="Status: Error")

        self.stop_logging()

    def update_plot(self):
        self.ax.clear()
        self.ax.plot(self.time_log, self.power_log, linestyle="-", marker=".")

        for t, state in self.led_transitions:
            color = 'green' if state == 'ON' else 'red'
            self.ax.axvline(x=t, color=color, linestyle='--')

        self.ax.set_title("Power vs Time")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Power (W)")
        self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
        self.figure.autofmt_xdate()
        self.canvas.draw()

    def save_plot(self):
        filename = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if filename:
            self.figure.savefig(filename)
            messagebox.showinfo("Saved", f"Plot saved as {filename}")

    def start_logging(self):
        self.running = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.time_log.clear()
        self.power_log.clear()
        self.led_transitions.clear()
        self.ax.clear()
        self.ax.set_title("Power vs Time")
        self.ax.set_xlabel("Time")
        self.ax.set_ylabel("Power (W)")
        self.canvas.draw()

        self.thread = threading.Thread(target=self.pulse_and_record)
        self.thread.start()

    def stop_logging(self):
        self.running = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="Status: Stopped")

    def init_devices(self):
        self.device_listbox.delete(0, tk.END)
        self.device_map.clear()
        try:
            resources = self.rm.list_resources()
            found_dc2200 = None
            found_pm101 = None

            for res in resources:
                try:
                    dev = self.rm.open_resource(res)
                    idn = dev.query("*IDN?").strip().upper()
                    self.device_listbox.insert(tk.END, f"{res} → {idn}")
                    self.device_map[res] = idn
                    if "DC2200" in idn:
                        found_dc2200 = dev
                    elif "PM101" in idn:
                        found_pm101 = dev
                except Exception:
                    self.device_listbox.insert(tk.END, f"{res} → <no response>")

            if not found_dc2200 or not found_pm101:
                raise RuntimeError("Could not identify both DC2200 and PM101 from connected devices.")

            self.dc2200 = found_dc2200
            self.pm101 = found_pm101

            self.status_label.config(text="Status: Devices connected")

        except Exception as e:
            messagebox.showerror("Error", f"Device initialization failed:\n{e}")
            self.status_label.config(text="Status: Connection failed")

    def close(self):
        self.running = False
        if self.dc2200:
            self.dc2200.close()
        if self.pm101:
            self.pm101.close()
        self.rm.close()
        self.master.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = PowerLoggerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()

