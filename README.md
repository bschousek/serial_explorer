# serial_explorer

## Overview
A comprehensive Python utility for deep serial port exploration on Windows, providing unprecedented insights into system serial interfaces beyond traditional Device Manager limitations.

## Installation
```bash
# Ensure you have the required dependencies
pip install wmi
```

## Core Classes and Usage

### WindowsSerialPortInspector
The primary class for serial port discovery and information retrieval.

#### Initialization
```python
from serial_explorer import WindowsSerialPortInspector

# Create an inspector instance
inspector = WindowsSerialPortInspector()
```

#### Key Methods

1. **Get Present Ports Summary**
```python
# Retrieve information for currently connected ports
present_ports = inspector.get_present_ports_summary()
for port in present_ports:
    print(f"Port: {port['port_name']}")
    print(f"Description: {port['description']}")
    print(f"Manufacturer: {port['manufacturer']}")
```

2. **Get Absent Ports Summary**
```python
# Discover ports that are not currently connected
absent_ports = inspector.get_absent_ports_summary()
for port in absent_ports:
    print(f"Disconnected Port: {port['port_name']}")
```

3. **Get Comprehensive Port Details**
```python
# Retrieve full details for all ports
detailed_ports = inspector.get_all_ports_detailed_info()
for port in detailed_ports:
    print("Full Port Information:")
    print(f"Port Name: {port['port_name']}")
    print(f"Device ID: {port['device_id']}")
    print(f"USB Details: {port['usb_details']}")
    print(f"Driver Info: {port['driver_info']}")
```

### Progress Tracking (Optional)
```python
# Optional progress callback for long-running port scans
def progress_callback(progress):
    print(f"Scanning Progress: {progress:.2f}%")

# Use progress tracking with any method
inspector.get_all_ports_detailed_info(progress_callback=progress_callback)
```

### SerialPort Class
For advanced users, the `SerialPort` class provides granular port information:

```python
from serial_explorer import SerialPort

# Create a port object for a specific port
port = SerialPort('COM3')
port._populate_summary_info()
port._populate_detailed_info()

# Access detailed information
summary = port.get_summary_info()
details = port.get_detailed_info()
```

## Information Retrieved
- Port identification
- Connection status
- Hardware details
- USB information
- Driver metadata
- System configuration flags

## System Requirements
- Python 3.x
- Windows operating system
- Administrative privileges recommended
- Required libraries: 
  - `wmi`
  - `winreg`

## Potential Use Cases
- Hardware diagnostics
- Device driver investigation
- IoT device management
- Embedded system debugging

## Limitations
- Windows-specific implementation
- Requires system-level information access

## Future Enhancements
- Cross-platform port discovery
- Advanced filtering mechanisms
- Expanded port state tracking