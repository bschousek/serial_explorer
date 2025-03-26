# Windows Serial Port Inspector

## Overview

The WindowsSerialPortInspector is a Python library for comprehensive serial port discovery and information retrieval on Windows systems. It provides detailed insights into serial ports, including both currently connected and previously connected devices.

## Prerequisites

- Python 3.7+
- Required Libraries:
  ```bash
  pip install wmi pywin32
  ```

## Installation

```bash
# Clone the repository
git clone https://github.com/bschousek/serial_explorer

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Basic Example

```python
from serial_port_inspector import WindowsSerialPortInspector

# Create an inspector
inspector = WindowsSerialPortInspector()

# Get all serial ports
ports = inspector.list_serial_ports()

# Iterate and display detailed information
for port in ports:
    print(f"Port: {port}")
    
    # Get comprehensive port details
    details = port.get_detailed_info()
    
    # Access specific information
    print(f"Manufacturer: {details['manufacturer']}")
    print(f"USB Vendor ID: {details['usb_details']['vendor_id']}")
```

## Classes

### `SerialPort`

Represents a single serial port with comprehensive system information.

#### Attributes

- `port_name`: COM port name (e.g., 'COM3')
- `description`: Detailed device description
- `manufacturer`: Device manufacturer
- `device_id`: Unique device identifier
- `hardware_ids`: List of hardware identifiers
- `compatible_ids`: List of compatible device identifiers
- `registry_path`: Windows Registry path for the device
- `friendly_name`: User-friendly device name
- `status`: Current device status

#### Methods

- `get_detailed_info()`: Returns a comprehensive dictionary of all port information
- `__str__()`: String representation of the serial port

### `WindowsSerialPortInspector`

Discovers and lists serial ports on a Windows system.

#### Methods

- `list_serial_ports()`: Returns a list of SerialPort objects

## Detailed Information Retrieval

The library uses multiple techniques to gather serial port information:

- Windows Management Instrumentation (WMI) queries
- Windows Registry examination
- Subprocess command-line enumeration

## Retrievable Information

Each serial port object can provide:

- Detailed device description
- Manufacturer information
- Driver details
  - Provider
  - Installation date
  - Version
- USB-specific information
  - Vendor ID
  - Product ID
- System-level details
  - Class GUID
  - Configuration flags
  - Problem codes

## Limitations

- Windows-specific implementation
- Requires administrative privileges for full information retrieval
- Information may vary based on device type and system configuration

## Error Handling

- Gracefully handles missing or inaccessible information
- Prints error messages for failed information retrieval
- Continues processing other ports if one fails

## Troubleshooting

- Ensure WMI and pywin32 are correctly installed
- Run with administrator privileges for complete information
- Check Windows event logs for device-specific issues

## Performance Considerations

- Information retrieval may take a few seconds for multiple ports
- No built-in caching mechanisms

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

Your Name - your.email@example.com

Project Link: [https://github.com/yourusername/windows-serial-port-inspector](https://github.com/yourusername/windows-serial-port-inspector)

## Version

- Current Version: 1.1.0
- Last Updated: March 2024
- Compatibility: Windows 10/11, Python 3.7+

## Disclaimer

This library is provided "as-is" without warranties. Always verify critical system information through official Windows tools.