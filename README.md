# CPU Cooling Agent
## Patent Documentation and Technical Manual

### Technical Specifications

1. **System Architecture**
   - Real-time monitoring of CPU temperature and usage
   - Predictive temperature modeling using linear regression
   - Adaptive fan control with multiple cooling profiles
   - System health assessment based on thermal conditions

2. **Monitoring Algorithms**
   - Temperature prediction using last 10 data points (Linear Regression)
   - Dynamic threshold adjustment based on cooling profile
   - Power consumption estimation (CPU frequency * usage)
   - Battery life impact analysis

3. **Fan Control Mechanisms**
   - Direct WMI interface for fan speed control
   - Fallback to ACPI if WMI unavailable
   - Profile-based fan curves (silent/balanced/performance)
   - Emergency quick-cool function (100% fan speed for 30s)

4. **Hardware Integration**
   - OpenHardwareMonitor for sensor data collection
   - Supports standard WMI/ACPI fan control interfaces
   - Automatic detection of controllable fans
   - Graceful degradation when hardware access fails

5. **Data Visualization**
   - Real-time temperature and health graphs
   - Color-coded status indicators (normal/warning/critical)
   - Scrollable interface for small screens
   - Responsive layout that adapts to window size

### Installation

1. **Prerequisites**
   - Python 3.8 or higher installed
   - .NET Framework 4.5 or higher (required for OpenHardwareMonitor)
   - Administrative privileges (for hardware monitoring)

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **OpenHardwareMonitor Setup**
   - Extract all files from OpenHardwareMonitor directory
   - Run OpenHardwareMonitor.exe once to generate configuration
   - Ensure it has proper permissions to monitor hardware

### Usage

1. **Running the Agent**
   ```bash
   python cpu_cooling_agent.py
   ```

2. **Command Line Options**
   ```bash
   python cpu_cooling_agent.py --interval 5 --threshold 75 --log logs.csv
   ```
   - `--interval`: Monitoring interval in seconds (default: 3)
   - `--threshold`: Temperature threshold in °C (default: 70)
   - `--log`: Path to log file (default: cpu_cooling_logs_[timestamp].csv)

3. **Monitoring Interface**
   - The agent will display real-time CPU temperature and cooling status
   - OpenHardwareMonitor provides detailed hardware statistics

### Configuration

1. **OpenHardwareMonitor Configuration**
   - Edit `OpenHardwareMonitor.config` to customize:
     - Sensor monitoring intervals
     - Temperature thresholds
     - Plot display settings

2. **Agent Configuration**
   - Modify `cpu_cooling_agent.py` to adjust:
     - Cooling algorithm parameters
     - Fan control logic
     - Notification settings

### Building Executable

1. **Create standalone executable**
   ```bash
   pyinstaller --onefile cpu_cooling_agent.py
   ```
   - Output will be in `dist/` directory

### Troubleshooting

1. **Common Issues**
   - "Permission denied" errors: Run as Administrator
   - Missing dependencies: Verify all Python packages are installed
   - Hardware not detected: Check OpenHardwareMonitor compatibility

2. **Logging**
   - Detailed logs are saved in timestamped CSV files
   - Includes temperature readings and cooling actions

### License
- This project uses OpenHardwareMonitor under its license terms
- See `License.html` in OpenHardwareMonitor directory

### Support
For technical support or feature requests, please contact the project maintainers.