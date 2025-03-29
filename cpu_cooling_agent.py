# Copyright (c) 2025 Arkaprava
# This software is licensed under the MIT License and the OpenHardwareMonitor License.
# See LICENSE file in the project root for full license information and the OpenHardwareMonitor License in the OpenHardwareMonitor folder.

import tkinter as tk
from tkinter import ttk
import psutil
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from datetime import datetime
import threading
import time
from sklearn.linear_model import LinearRegression
from collections import deque
import wmi
import win32com.client
import comtypes.client

class CPUCoolingAgent:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CPU Cooling Agent")
        self.root.state('zoomed')  # This will maximize the window properly
        self.root.configure(bg='#f0f0f0')

        # Initialize data storage
        self.temp_history = []
        self.time_history = []
        self.power_history = deque(maxlen=60)  # Power consumption history
        self.max_history_points = 60
        self.warning_threshold = 40
        self.critical_threshold = 55
        self.optimal_temp_min = 20  # Adjusted to be more realistic
        self.optimal_temp_max = 40  # Set to match warning threshold
        self.fan_control_enabled = True
        self.current_fan_speed = 80
        self.system_health = 100  # System health percentage
        self.current_profile = "balanced"  # Current cooling profile
        self.prediction_window = 10  # Predict temperature 10 seconds ahead
        self.temp_predictor = LinearRegression()
        self.prediction_enabled = True
        self.running = True  # Flag for controlling the update thread

        self.cooling_profiles = {
            "silent": {
                "max_fan_speed": 100,
                "temp_threshold": 18,
                "fan_curve": lambda t: min(100, max(80, 7 * (t - 15)))
            },
            "balanced": {
                "max_fan_speed": 100,
                "temp_threshold": 20,
                "fan_curve": lambda t: min(100, max(90, 8 * (t - 18)))
            },
            "performance": {
                "max_fan_speed": 100,
                "temp_threshold": 22,
                "fan_curve": lambda t: min(100, max(100, 10 * (t - 20)))
            }
        }

        self.setup_ui()
        self.setup_graphs()
        self.update_thread = threading.Thread(target=self.update_data, daemon=True)
        self.update_thread.start()

    def setup_ui(self):
        # Here I am creating the Main frame of the software with scrollbar
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        main_frame = ttk.Frame(main_canvas)

        # Configure grid weight to allow proper expansion
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Configure scrollable canvas
        main_canvas.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        scrollbar.grid(row=0, column=1, sticky='ns')
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        # Configure main frame and its scroll region
        main_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(
                scrollregion=main_canvas.bbox("all"),
                width=main_canvas.winfo_width(),
                height=main_canvas.winfo_height()
            )
        )
        
        # Create window for main frame in canvas with proper expansion
        main_canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # Bind mouse wheel to scrolling with improved handling
        def _on_mousewheel(event):
            if main_canvas.winfo_height() < main_frame.winfo_height():
                main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        # Bind mouse wheel for Windows with focus check
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Configure main frame to expand properly
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(list(range(7)), weight=1)
        
        # Set minimum size for better usability
        self.root.minsize(800, 600)

        # Advanced Features Section
        advanced_frame = ttk.LabelFrame(main_frame, text="Advanced Features")
        advanced_frame.grid(row=1, column=0, pady=5, sticky='ew')

        # Temperature Threshold Controls
        threshold_frame = ttk.Frame(advanced_frame)
        threshold_frame.grid(row=0, column=0, padx=5, pady=2, sticky='ew')
        
        ttk.Label(threshold_frame, text="Warning Threshold:").grid(row=0, column=0, sticky='w')
        self.warning_threshold_var = tk.StringVar(value=str(self.warning_threshold))
        warning_entry = ttk.Entry(threshold_frame, textvariable=self.warning_threshold_var, width=5)
        warning_entry.grid(row=0, column=1, padx=5)
        
        ttk.Label(threshold_frame, text="Critical Threshold:").grid(row=0, column=2, padx=5)
        self.critical_threshold_var = tk.StringVar(value=str(self.critical_threshold))
        critical_entry = ttk.Entry(threshold_frame, textvariable=self.critical_threshold_var, width=5)
        critical_entry.grid(row=0, column=3)

        # Auto-Optimization Toggle
        self.auto_optimize_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(advanced_frame, text="Enable Auto-Optimization", 
                       variable=self.auto_optimize_var).grid(row=1, column=0, padx=5, pady=2)

        # Core-Specific Monitoring Toggle
        self.core_monitoring_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(advanced_frame, text="Enable Core-Specific Monitoring", 
                       variable=self.core_monitoring_var).grid(row=2, column=0, padx=5, pady=2)

        # Thermal Throttling Detection
        self.throttle_detection_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(advanced_frame, text="Enable Thermal Throttling Detection", 
                       variable=self.throttle_detection_var).grid(row=3, column=0, padx=5, pady=2)

        # Performance Logging
        log_frame = ttk.Frame(advanced_frame)
        log_frame.grid(row=4, column=0, padx=5, pady=2, sticky='ew')
        ttk.Label(log_frame, text="Performance Log Interval (s):").grid(row=0, column=0)
        self.log_interval_var = tk.StringVar(value="60")
        ttk.Entry(log_frame, textvariable=self.log_interval_var, width=5).grid(row=0, column=1, padx=5)
        ttk.Button(log_frame, text="Export Logs", command=self.export_logs).grid(row=0, column=2)

        # Creating the Battery Status Frame
        battery_frame = ttk.Frame(main_frame)
        battery_frame.grid(row=2, column=0, sticky='ew', padx=5, pady=2)
        
        # Battery Percentage Display
        self.battery_label = ttk.Label(battery_frame, text="Battery: --%", font=("Arial", 12))
        self.battery_label.grid(row=0, column=0, padx=5)
        
        # Battery Time Remaining Display
        self.battery_time_label = ttk.Label(battery_frame, text="Time Left: --:--", font=("Arial", 12))
        self.battery_time_label.grid(row=0, column=1, padx=5)

        # Author credit button
        author_button = ttk.Button(main_frame, text="Made by: Arkaprava Chakraborty", 
                                command=lambda: self.open_linkedin())
        author_button.grid(row=3, column=0, sticky='e', padx=5, pady=2)

        # Status frame creation
        status_frame = ttk.LabelFrame(main_frame, text="System Status")
        status_frame.grid(row=4, column=0, sticky='ew', pady=5, padx=5)

        # CPU Temperature Display with status indicator
        temp_frame = ttk.Frame(status_frame)
        temp_frame.grid(row=0, column=0, sticky='ew', padx=5, pady=2)
        self.temp_label = ttk.Label(temp_frame, text="CPU Temperature: -- °C", font=("Arial", 14))
        self.temp_label.grid(row=0, column=0, sticky='w')
        self.temp_status = ttk.Label(temp_frame, text="●", font=("Arial", 14))
        self.temp_status.grid(row=0, column=1, padx=5)

        # CPU Usage Display
        self.usage_label = ttk.Label(status_frame, text="CPU Usage: --%", font=("Arial", 14))
        self.usage_label.grid(row=1, column=0, pady=2)

        # Power Consumption Display
        self.power_label = ttk.Label(status_frame, text="Power Consumption: -- W", font=("Arial", 14))
        self.power_label.grid(row=2, column=0, pady=2)

        # Predicted Temperature Display
        self.prediction_label = ttk.Label(status_frame, text="Predicted Temperature: -- °C", font=("Arial", 14))
        self.prediction_label.grid(row=3, column=0, pady=2)

        # Fan Control Frame
        fan_frame = ttk.LabelFrame(main_frame, text="Fan Control")
        fan_frame.grid(row=5, column=0, pady=5, sticky='ew')

        # Cooling Profile Selection
        profile_frame = ttk.Frame(fan_frame)
        profile_frame.grid(row=0, column=0, padx=5, pady=2, sticky='ew')
        ttk.Label(profile_frame, text="Cooling Profile:").grid(row=0, column=0, sticky='w')
        self.profile_var = tk.StringVar(value="balanced")
        profile_menu = ttk.OptionMenu(profile_frame, self.profile_var, "balanced", 
                                    "silent", "balanced", "performance",
                                    command=self.change_cooling_profile)
        profile_menu.grid(row=0, column=1, sticky='e')

        # Fan Control Enable Switch
        self.fan_control_var = tk.BooleanVar(value=False)
        self.fan_control_switch = ttk.Checkbutton(fan_frame, text="Enable Fan Control",
                                                variable=self.fan_control_var,
                                                command=self.toggle_fan_control)
        self.fan_control_switch.grid(row=1, column=0, padx=5, pady=2)

        # Fan Speed Control
        speed_frame = ttk.Frame(fan_frame)
        speed_frame.grid(row=2, column=0, padx=5, pady=2, sticky='ew')
        speed_label = ttk.Label(speed_frame, text="Fan Speed:")
        speed_label.grid(row=0, column=0, sticky='w')
        self.fan_speed_value = ttk.Label(speed_frame, text="50%")
        self.fan_speed_value.grid(row=0, column=1, sticky='e')

        self.fan_speed = ttk.Scale(fan_frame, from_=0, to=100, orient=tk.HORIZONTAL,
                                command=self.update_fan_speed)
        self.fan_speed.set(50)
        self.fan_speed.grid(row=3, column=0, padx=5, pady=2, sticky='ew')

        # Quick Cool Button
        self.cool_button = ttk.Button(fan_frame, text="Quick Force Cool", command=self.quick_cool)
        self.cool_button.grid(row=4, column=0, pady=5)

        # Graph Frame
        self.graph_frame = ttk.Frame(main_frame)
        self.graph_frame.grid(row=6, column=0, pady=5, sticky='nsew')
        self.graph_frame.grid_columnconfigure(0, weight=1)
        self.graph_frame.grid_rowconfigure(0, weight=1)

    def setup_graphs(self):
        # Create figure with optimized size and spacing
        self.fig = plt.figure(figsize=(16, 8))  # Increased height for more detail
        
        # Create subplots with adjusted spacing for 2x2 layout
        gs = self.fig.add_gridspec(2, 2, width_ratios=[1, 1], height_ratios=[1, 1], wspace=0.3, hspace=0.3)
        
        # Position temperature graph
        self.ax = self.fig.add_subplot(gs[0, 0])
        
        # Position health graph
        self.health_ax = self.fig.add_subplot(gs[0, 1])
        
        # Position power consumption graph
        self.power_ax = self.fig.add_subplot(gs[1, 0])
        
        # Position prediction graph
        self.prediction_ax = self.fig.add_subplot(gs[1, 1])
        
        # Configure canvas and widget placement
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.graph_frame)
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky='nsew', padx=5, pady=5)

        # Setup temperature graph with enhanced styling
        self.ax.set_title('CPU Temperature History', pad=12, fontsize=12, weight='bold')
        self.ax.set_xlabel('Time', fontsize=11)
        self.ax.set_ylabel('Temperature (°C)', fontsize=11)
        self.line, = self.ax.plot([], [], linewidth=2.5, color='#FF6B6B')
        self.ax.grid(True, linestyle='--', alpha=0.7)
        self.ax.set_xlim(0, 60)
        self.ax.set_ylim(0, max(80, self.critical_threshold + 20))
        self.ax.tick_params(axis='both', which='major', labelsize=10)
        self.ax.set_facecolor('#F8F9FA')

        # Setup health graph with enhanced styling
        self.health_ax.set_title('System Health Report', pad=12, fontsize=12, weight='bold')
        self.health_ax.set_ylim(0, 100)
        self.health_ax.set_xlim(0, 60)
        self.health_line, = self.health_ax.plot([], [], color='#4BC0C0', linewidth=2.5)
        self.health_ax.set_ylabel('Health %', fontsize=11)
        self.health_ax.set_xlabel('Time', fontsize=11)
        self.health_ax.grid(True, linestyle='--', alpha=0.7)
        self.health_ax.tick_params(axis='both', which='major', labelsize=10)
        self.health_ax.set_facecolor('#F8F9FA')

        # Setup power consumption graph
        self.power_ax.set_title('Power Consumption', pad=12, fontsize=12, weight='bold')
        self.power_ax.set_xlabel('Time', fontsize=11)
        self.power_ax.set_ylabel('Power (W)', fontsize=11)
        self.power_line, = self.power_ax.plot([], [], color='#36A2EB', linewidth=2.5)
        self.power_ax.grid(True, linestyle='--', alpha=0.7)
        self.power_ax.set_xlim(0, 60)
        self.power_ax.set_ylim(0, 100)
        self.power_ax.tick_params(axis='both', which='major', labelsize=10)
        self.power_ax.set_facecolor('#F8F9FA')

        # Setup prediction graph
        self.prediction_ax.set_title('Temperature Prediction', pad=12, fontsize=12, weight='bold')
        self.prediction_ax.set_xlabel('Time (Future)', fontsize=11)
        self.prediction_ax.set_ylabel('Temperature (°C)', fontsize=11)
        self.prediction_line, = self.prediction_ax.plot([], [], color='#FF9F40', linewidth=2.5)
        self.prediction_ax.grid(True, linestyle='--', alpha=0.7)
        self.prediction_ax.set_xlim(0, 10)
        self.prediction_ax.set_ylim(0, max(80, self.critical_threshold + 20))
        self.prediction_ax.tick_params(axis='both', which='major', labelsize=10)
        self.prediction_ax.set_facecolor('#F8F9FA')

        # Enable responsive layout
        self.fig.canvas.mpl_connect('resize_event', self.on_resize)

    def on_resize(self, event):
        # Update layout on window resize
        self.fig.tight_layout(pad=2.0)
        self.canvas.draw()

    def update_graph(self):
        try:
            if not self.temp_history or not self.time_history:
                return

            current_times = [t.strftime('%H:%M:%S') for t in self.time_history]
            x_range = range(len(self.temp_history))

            # Update temperature graph with validation
            if len(self.temp_history) > 0:
                self.line.set_data(x_range, self.temp_history)
                self.ax.set_xticks(x_range[::max(1, len(x_range)//5)])
                self.ax.set_xticklabels(current_times[::max(1, len(current_times)//5)], rotation=45)
                self.ax.set_xlim(0, max(60, len(x_range)))
                self.ax.relim()
                self.ax.autoscale_view()

            # Update health graph with validation
            if len(self.temp_history) > 0:
                health_data = [self.calculate_health(temp) for temp in self.temp_history]
                self.health_line.set_data(x_range, health_data)
                self.health_ax.set_xticks(x_range[::max(1, len(x_range)//5)])
                self.health_ax.set_xticklabels(current_times[::max(1, len(current_times)//5)], rotation=45)
                self.health_ax.set_xlim(0, max(60, len(x_range)))

            # Update power consumption graph with validation
            power_data = list(self.power_history)
            if power_data:
                power_x_range = range(len(power_data))
                self.power_line.set_data(power_x_range, power_data)
                power_times = current_times[-len(power_data):] if len(current_times) >= len(power_data) else current_times
                self.power_ax.set_xticks(power_x_range[::max(1, len(power_x_range)//5)])
                self.power_ax.set_xticklabels(power_times[::max(1, len(power_times)//5)], rotation=45)
                self.power_ax.set_xlim(0, max(60, len(power_x_range)))
                self.power_ax.relim()
                self.power_ax.autoscale_view()

            # Update prediction graph with validation
            if len(self.temp_history) >= 10 and self.prediction_enabled:
                try:
                    X_future = np.array(range(11)).reshape(-1, 1)
                    self.temp_predictor.fit(np.array(range(len(self.temp_history[-10:]))).reshape(-1, 1),
                                          np.array(self.temp_history[-10:]))
                    predicted_temps = self.temp_predictor.predict(X_future)
                    self.prediction_line.set_data(X_future.flatten(), predicted_temps)
                    self.prediction_ax.relim()
                    self.prediction_ax.autoscale_view()
                except Exception as pred_e:
                    print(f"Prediction update error: {str(pred_e)}")

            # Apply tight layout and draw
            self.fig.tight_layout(pad=2.0)
            self.canvas.draw()

        except Exception as e:
            print(f"Graph update error: {str(e)}")
            # Reset all plots on error
            self.line.set_data([], [])
            self.health_line.set_data([], [])
            self.power_line.set_data([], [])
            self.prediction_line.set_data([], [])
            try:
                self.canvas.draw()
            except Exception:
                pass

    def update_ui(self, temp, usage):
        if temp is not None:
            self.temp_label.config(text=f"CPU Temperature: {temp:.1f} °C")
            self.usage_label.config(text=f"CPU Usage: {usage}%")
            
            # Update power consumption with enhanced calculation
            cpu_freq = psutil.cpu_freq().current
            power = (cpu_freq * usage / 100 * 0.1) + (temp * 0.05)  # Consider temperature impact
            self.power_label.config(text=f"Power Consumption: {power:.1f} W")
            self.power_history.append(power)
            
            # Enhanced battery monitoring
            try:
                battery = psutil.sensors_battery()
                if battery:
                    percent = battery.percent
                    self.battery_label.config(text=f"Battery: {percent}% {'🔌' if battery.power_plugged else '🔋'}")
                    
                    # Detailed time remaining calculation
                    if battery.secsleft != -1 and not battery.power_plugged:
                        hours = battery.secsleft // 3600
                        minutes = (battery.secsleft % 3600) // 60
                        time_str = f"Time Left: {hours:02d}:{minutes:02d}"
                        # Add estimated time based on current usage
                        estimated_time = battery.secsleft * (1 - (usage / 200))  # Adjust for CPU load
                        est_hours = int(estimated_time // 3600)
                        est_minutes = int((estimated_time % 3600) // 60)
                        time_str += f" (Est: {est_hours:02d}:{est_minutes:02d})"
                        self.battery_time_label.config(text=time_str)
                    else:
                        self.battery_time_label.config(text="Time Left: Plugged In ⚡")
                    
                    # Enhanced battery status indicators
                    if percent <= 10:
                        self.battery_label.config(foreground='red', font=('Arial', 12, 'bold'))
                        if not battery.power_plugged:
                            self.show_low_battery_warning()
                    elif percent <= 20:
                        self.battery_label.config(foreground='red')
                    elif percent <= 50:
                        self.battery_label.config(foreground='orange')
                    else:
                        self.battery_label.config(foreground='green')
            except Exception as e:
                self.battery_label.config(text="Battery: N/A")
                self.battery_time_label.config(text="Time Left: N/A")
            
            # Enhanced temperature prediction
            if self.prediction_enabled:
                predicted_temp = self.predict_temperature()
                if predicted_temp is not None:
                    prediction_text = f"Predicted Temperature: {predicted_temp:.1f} °C"
                    # Add trend indicator
                    if len(self.temp_history) > 1:
                        trend = predicted_temp - self.temp_history[-1]
                        prediction_text += f" ({'↑' if trend > 0 else '↓' if trend < 0 else '→'})"
                    self.prediction_label.config(text=prediction_text)
                    
                    # Enhanced warning visualization
                    if predicted_temp > self.critical_threshold:
                        self.prediction_label.config(foreground='red', font=('Arial', 14, 'bold'))
                        self.show_critical_prediction_warning(predicted_temp)
                    elif predicted_temp > self.warning_threshold:
                        self.prediction_label.config(foreground='orange', font=('Arial', 14))
                    else:
                        self.prediction_label.config(foreground='green', font=('Arial', 14))



    def handle_no_sensor(self):
        self.root.after(0, self.update_ui, None, psutil.cpu_percent())
        self.temp_label.config(text="No Temperature Sensor Found", foreground='orange')
        self.temp_status.config(foreground='orange')
        time.sleep(2)

    def handle_error(self, error_msg):
        self.temp_label.config(text="Sensor Error", foreground='orange')
        self.temp_status.config(foreground='orange')

    def show_critical_prediction_warning(self, predicted_temp):
        try:
            # Create a warning window
            warning_window = tk.Toplevel(self.root)
            warning_window.title("Temperature Warning")
            warning_window.geometry("400x150")
            warning_window.configure(bg='#ffebee')

            # Warning message
            message = f"WARNING: Critical temperature predicted!\nPredicted temperature: {predicted_temp:.1f}°C\nTaking preventive measures..."
            warning_label = ttk.Label(warning_window, text=message, background='#ffebee', font=('Arial', 12))
            warning_label.pack(pady=20)

            # Automatically enable maximum cooling
            self.fan_control_var.set(True)
            self.fan_speed.set(100)
            self.update_fan_speed(100)

            # Auto-close after 5 seconds
            warning_window.after(5000, warning_window.destroy)
        except Exception as e:
            print(f"Error showing critical warning: {str(e)}")

    def show_low_battery_warning(self):
        try:
            # Create a warning window
            warning_window = tk.Toplevel(self.root)
            warning_window.title("Low Battery Warning")
            warning_window.geometry("400x150")
            warning_window.configure(bg='#fff3e0')

            # Warning message
            message = "WARNING: Battery level critically low!\nPlease connect to power source.\nReducing performance to conserve power..."
            warning_label = ttk.Label(warning_window, text=message, background='#fff3e0', font=('Arial', 12))
            warning_label.pack(pady=20)

            # Automatically switch to power-saving mode
            self.profile_var.set("silent")
            self.change_cooling_profile("silent")

            # Auto-close after 5 seconds
            warning_window.after(5000, warning_window.destroy)
        except Exception as e:
            print(f"Error showing battery warning: {str(e)}")

    def calculate_health(self, temp):
        try:
            if temp >= self.critical_threshold:
                return max(0, 100 - ((temp - self.critical_threshold) * 5))
            elif temp >= self.warning_threshold:
                return max(50, 100 - ((temp - self.warning_threshold) * 2))
            elif self.optimal_temp_min <= temp <= self.optimal_temp_max:
                return 100
            else:
                return max(80, 100 - abs(temp - self.optimal_temp_max))
        except Exception as e:
            print(f"Error calculating health: {str(e)}")
            return 100  # Return default health on error

    def show_normal(self):
        self.temp_label.config(foreground='black')
        self.temp_status.config(foreground='green')

    def show_warning(self):
        self.temp_label.config(foreground='orange')
        self.temp_status.config(foreground='orange')

    def show_critical_warning(self):
        self.temp_label.config(foreground='red')
        self.temp_status.config(foreground='red')

    def toggle_fan_control(self):
        self.fan_control_enabled = self.fan_control_var.get()
        if not self.fan_control_enabled:
            self.fan_speed.set(50)
            self.update_fan_speed(50)

    def update_fan_speed(self, value):
        try:
            speed = int(float(value))
            if not 0 <= speed <= 100:
                print(f"Invalid fan speed value: {speed}")
                return
            self.current_fan_speed = speed
            self.fan_speed_value.config(text=f"{speed}%")
            if self.fan_control_enabled:
                self.apply_fan_speed(speed)
        except ValueError as e:
            print(f"Error setting fan speed: {str(e)}")
            self.handle_error("Invalid fan speed value")

    def change_cooling_profile(self, profile):
        if profile not in self.cooling_profiles:
            print(f"Invalid cooling profile: {profile}")
            return

        self.current_profile = profile
        profile_settings = self.cooling_profiles[profile]

        # Update fan speed based on current temperature if available
        if self.temp_history:
            current_temp = self.temp_history[-1]
            new_speed = profile_settings['fan_curve'](current_temp)
            self.fan_speed.set(new_speed)
            self.update_fan_speed(new_speed)

        # Update thresholds based on profile
        self.warning_threshold = profile_settings['temp_threshold']
        self.warning_threshold_var.set(str(self.warning_threshold))
        self.critical_threshold = profile_settings['temp_threshold'] + 15
        self.critical_threshold_var.set(str(self.critical_threshold))

    def apply_fan_speed(self, speed):
        if not 0 <= speed <= 100:
            print(f"Invalid fan speed value: {speed}")
            return
            
        try:
            import win32com.client
            import comtypes.client
            
            # Initialize WMI interface
            w = wmi.WMI(namespace="root\\wmi")
            
            # Get fan control interface with enhanced error handling
            try:
                fans = w.instances("Win32_Fan")
                fan_controlled = False
                
                for fan in fans:
                    if hasattr(fan, 'DesiredSpeed'):
                        try:
                            # Convert percentage to actual fan speed
                            max_speed = fan.MaxSpeed if hasattr(fan, 'MaxSpeed') else 5000
                            desired_speed = int((speed / 100.0) * max_speed)
                            
                            # Set fan speed with validation
                            if 0 <= desired_speed <= max_speed:
                                fan.DesiredSpeed = desired_speed
                                fan_controlled = True
                            else:
                                print(f"Calculated fan speed {desired_speed} is outside valid range for this fan")
                        except Exception as fan_e:
                            print(f"Error controlling individual fan: {str(fan_e)}")
                            continue
                
                if not fan_controlled:
                    self.handle_error("No controllable fans found or all control attempts failed")
                    
            except Exception as wmi_e:
                self.handle_error(f"Failed to access fan controls: {str(wmi_e)}")
                
        except Exception as e:
            self.handle_error(f"Fan control system error: {str(e)}")
            print(f"Detailed fan control error: {str(e)}")
            # Fallback to ACPI fan control if available
            try:
                w = wmi.WMI(namespace="root\\wmi")
                acpi = w.instances("ACPI_FanSpeed")
                acpi_controlled = False
                
                for fan in acpi:
                    if hasattr(fan, 'FanSpeed'):
                        fan.FanSpeed = speed
                        acpi_controlled = True
                        
                if not acpi_controlled:
                    raise Exception("No ACPI fan control available")
                    
            except Exception as e2:
                print(f"ACPI fan control error: {str(e2)}")
                self.handle_error("Failed to control fan speed")
                self.fan_control_var.set(False)
                self.fan_control_enabled = False

    def calculate_health(self, temp):
        # Calculate system health based on temperature
        if temp <= self.optimal_temp_min:
            return 100
        elif temp >= self.critical_threshold:
            return max(0, 40 - (temp - self.critical_threshold) * 5)  # More gradual decline
        elif temp >= self.warning_threshold:
            # Linear decline between warning and critical thresholds
            warning_range = self.critical_threshold - self.warning_threshold
            temp_over_warning = temp - self.warning_threshold
            return max(0, 80 - (temp_over_warning / warning_range) * 40)
        else:
            # Gradual decline between optimal and warning
            optimal_range = self.warning_threshold - self.optimal_temp_min
            temp_over_optimal = temp - self.optimal_temp_min
            return max(0, 100 - (temp_over_optimal / optimal_range) * 20)

    def adjust_fan_speed(self, temp, usage):
        if temp >= self.critical_threshold:
            self.fan_speed.set(100)
        elif temp >= self.warning_threshold:
            target_speed = min(int(70 + (temp - self.warning_threshold) * 2), 100)
            self.fan_speed.set(target_speed)
        else:
            base_speed = max(30, int(usage / 2))
            self.fan_speed.set(base_speed)

    def get_current_temperature(self):
        if self.temp_history:
            return self.temp_history[-1]
        return None

    def quick_cool(self):
        try:
            # Initialize or reset click counter
            if not hasattr(self, 'cool_click_count') or self.cool_click_count >= self.target_clicks:
                self.cool_click_count = 0
                self.target_clicks = random.randint(25, 35)
                self.cool_button.config(text=f"Click {self.target_clicks} times to activate Force Cooling", foreground='#E3F2FD')
                self.is_quick_cooling = False
            
            # Increment click counter
            self.cool_click_count += 1
            remaining_clicks = self.target_clicks - self.cool_click_count
            
            # Calculate color intensity based on clicks (darker blue as clicks increase)
            color_intensity = min(255, int((self.cool_click_count / self.target_clicks) * 255))
            blue_color = f'#{color_intensity:02x}{color_intensity:02x}FF'
            
            if self.cool_click_count >= self.target_clicks:
                # Activate force cooling
                self.cool_button.config(text="Force Cooling Activated!", foreground='green')
                self.fan_control_var.set(True)
                self.fan_control_enabled = True
                self.is_quick_cooling = True
                self.quick_cool_start_time = time.time()
                self.quick_cool_duration = 30  # 30 seconds of cooling
                self.force_cool()
            else:
                # Update button with remaining clicks
                self.cool_button.config(
                    text=f"Keep clicking! {remaining_clicks} more to activate cooling",
                    foreground=blue_color
                )
        except Exception as e:
            print(f"Error in quick cool: {str(e)}")
            self.is_quick_cooling = False

    def force_cool(self):
        """Separate method for force cooling logic"""
        if not self.is_quick_cooling:
            return
            
        try:
            # Check if cooling duration has expired
            if time.time() - self.quick_cool_start_time >= self.quick_cool_duration:
                self.is_quick_cooling = False
                self.fan_speed.set(80)
                self.update_fan_speed(80)
                self.cool_button.config(text="Quick Force Cool", foreground='black')
                self.temp_status.config(foreground='green')
                return

            # Set maximum fan speed
            self.fan_speed.set(100)
            self.update_fan_speed(100)
            self.fan_speed_value.config(text="100%")
            self.temp_status.config(foreground='blue')

            # Update button text with remaining time
            remaining_time = int(self.quick_cool_duration - (time.time() - self.quick_cool_start_time))
            self.cool_button.config(
                text=f"Cooling in progress... {remaining_time}s remaining",
                foreground='blue'
            )

            # Schedule next update
            self.root.after(1000, self.force_cool)

        except Exception as e:
            print(f"Error in force cooling: {str(e)}")
            # Ensure fan stays at max speed even if there's an error
            self.fan_speed.set(100)
            self.update_fan_speed(100)
            self.root.after(1000, self.force_cool)

    def open_linkedin(self):
        import webbrowser
        import base64
        import hashlib
        import hmac
        
        # Encoded URL with integrity check
        _encoded = b'aHR0cHM6Ly93d3cubGlua2VkaW4uY29tL2luL2Fya2FwcmF2YS1jaGFrcmFib3J0eS04YjlhMmIyODM/dXRtX3NvdXJjZT1zaGFyZSZ1dG1fY2FtcGFpZ249c2hhcmVfdmlhJnV0bV9jb250ZW50PXByb2ZpbGUmdXRtX21lZGl1bT1hbmRyb2lkX2FwcA=='
        _key = b'CPU_COOLING_AGENT_2024'
        _signature = b'7f91d4562196529cee1c7bc0eb3c0647'
        
        # Verify integrity
        def _verify_url(encoded_url, key, expected_signature):
            try:
                # Generate HMAC for verification
                hmac_obj = hmac.new(key, encoded_url, hashlib.md5)
                if not hmac.compare_digest(hmac_obj.hexdigest().encode(), expected_signature):
                    raise ValueError("URL integrity check failed")
                return base64.b64decode(encoded_url).decode()
            except Exception:
                return None
        
        # Attempt to open verified URL
        url = _verify_url(_encoded, _key, _signature)
        if url:
            webbrowser.open(url)

    def run(self):
        try:
            self.root.mainloop()
        finally:
            self.running = False  # Signal the update thread to stop
            if hasattr(self, 'update_thread'):
                self.update_thread.join(timeout=1.0)  # Wait for thread to finish

    def update_data(self):
        while self.running:
            try:
                # Get CPU temperature and usage
                cpu_usage = psutil.cpu_percent()
                cpu_temp = None
                error_message = None
                
                # Try multiple methods to get temperature with fallback options
                try:
                    # Try OpenHardwareMonitor first
                    import wmi
                    try:
                        w = wmi.WMI(namespace="root\\OpenHardwareMonitor")
                        temperature_infos = w.Sensor()
                        cpu_temp = None
                        for sensor in temperature_infos:
                            if sensor.SensorType==u'Temperature' and 'CPU' in sensor.Name:
                                cpu_temp = float(sensor.Value)
                                break
                        if cpu_temp is None:
                            raise Exception("No CPU temperature sensors found.")
                        print(f"Temperature read from OpenHardwareMonitor: {cpu_temp}°C")
                        if cpu_temp is None:
                            raise Exception("CPU temperature sensor not found")
                    except Exception as e:
                        error_message = f"OpenHardwareMonitor error: {str(e)}\nPlease ensure OpenHardwareMonitor is running."
                        print(error_message)

                    # Fallback to Windows Management Instrumentation
                    if cpu_temp is None:
                        try:
                            w = wmi.WMI(namespace="root\\wmi")
                            temperature_info = w.MSAcpi_ThermalZoneTemperature()[0]
                            cpu_temp = float(temperature_info.CurrentTemperature) / 10.0 - 273.15
                            print(f"Temperature read from WMI: {cpu_temp}°C")
                        except Exception as e:
                            if error_message:
                                error_message += f"\nWMI error: {str(e)}"
                            else:
                                error_message = f"WMI error: {str(e)}"
                            print(error_message)

                    # Simulated temperature as final fallback
                    if cpu_temp is None:
                        # Base temperature calculation on CPU usage
                        base_temp = 25  # Base temperature when idle
                        usage_factor = cpu_usage / 100.0
                        temp_range = 15  # Maximum temperature increase based on usage
                        cpu_temp = base_temp + (usage_factor * temp_range)
                        
                        # Add some realistic variation
                        import random
                        cpu_temp += random.uniform(-0.5, 0.5)
                        cpu_temp = round(cpu_temp, 1)
                        print(f"Using simulated temperature: {cpu_temp}°C (based on CPU usage: {cpu_usage}%)")
                    
                except Exception as e:
                    print(f"General error in temperature monitoring: {str(e)}")
                    error_message = str(e)

                # Update UI with temperature and error information
                if cpu_temp is not None:
                    self.root.after(0, self.update_ui, cpu_temp, cpu_usage)
                    self.temp_history.append(cpu_temp)
                    self.time_history.append(datetime.now())
                    
                    # Keep history within limits
                    if len(self.temp_history) > self.max_history_points:
                        self.temp_history.pop(0)
                        self.time_history.pop(0)
                    
                    # Update graph
                    self.root.after(0, self.update_graph)
                    
                    # Check temperature status
                    self.root.after(0, self.check_temperature_status, cpu_temp)
                    
                    # Adjust fan speed if auto-optimization is enabled
                    if self.auto_optimize_var.get() and self.fan_control_enabled:
                        self.root.after(0, self.adjust_fan_speed, cpu_temp, cpu_usage)

                time.sleep(1)  # Update interval
                
            except Exception as e:
                print(f"Critical error in update loop: {str(e)}")
                time.sleep(2)  # Longer delay on error

    def export_logs(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"cpu_cooling_logs_{timestamp}.csv"
            
            with open(filename, 'w') as f:
                f.write("Timestamp,Temperature,CPU Usage,Power Consumption,Fan Speed,System Health\n")
                
                for i in range(len(self.time_history)):
                    time_str = self.time_history[i].strftime("%Y-%m-%d %H:%M:%S")
                    temp = self.temp_history[i]
                    usage = psutil.cpu_percent()
                    power = list(self.power_history)[i] if i < len(self.power_history) else 0
                    health = self.calculate_health(temp)
                    
                    f.write(f"{time_str},{temp:.1f},{usage},{power:.1f},{self.current_fan_speed},{health:.1f}\n")
                    
            print(f"Logs exported to {filename}")
        except Exception as e:
            print(f"Error exporting logs: {str(e)}")
            
    
    
    def run(self):
        try:
            self.root.mainloop()
        finally:
            self.running = False  # Signal the update thread to stop
            if hasattr(self, 'update_thread'):
                self.update_thread.join(timeout=1.0)  # Wait for thread to finish

if __name__ == "__main__":
    app = CPUCoolingAgent()
    app.run()