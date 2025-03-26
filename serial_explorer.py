import winreg
import wmi
import re
import win32com.client
import subprocess
import os

class SerialPort:
    def __init__(self, port_name=None):
        """
        Initialize a SerialPort object with comprehensive information
        """
        self.wmi = wmi.WMI()
        self.port_name = port_name
        
        # Comprehensive port attributes
        self.description = None
        self.manufacturer = None
        self.device_id = None
        self.hardware_ids = []
        self.compatible_ids = []
        self.registry_path = None
        self.friendly_name = None
        self.status = None
        self.physical_location = None
        
        # Advanced port characteristics
        self.driver_info = {
            'provider': None,
            'date': None,
            'version': None
        }
        
        # Connection details
        self.connection_details = {
            'bus_type': None,
            'port_type': None,
            'speed_capabilities': []
        }
        
        # USB-specific details
        self.usb_details = {
            'vendor_id': None,
            'product_id': None,
            'serial_number': None,
            'manufacturer': None,
            'product_name': None
        }
        
        # Additional system information
        self.system_details = {
            'class_guid': None,
            'config_flags': None,
            'problem_code': None
        }
        
        # Always populate information if port name is provided
        if port_name:
            self._populate_all_info()

    def _populate_all_info(self):
        """
        Comprehensive method to populate all port information
        """
        try:
            # WMI Query for base information
            wmi_query = f"SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%{self.port_name}%'"
            wmi_ports = self.wmi.query(wmi_query)
            
            if wmi_ports:
                port = wmi_ports[0]
                self.description = port.Description
                self.manufacturer = port.Manufacturer
                self.device_id = port.DeviceID
                self.status = port.Status
                
                # Try to extract hardware and compatible IDs
                try:
                    self.hardware_ids = port.HardwareID or []
                    self.compatible_ids = port.CompatibleID or []
                except:
                    pass
            
            # Additional WMI queries for more detailed information
            self._populate_detailed_wmi_info()
            
            # Registry information
            self._populate_registry_details()
            
            # USB-specific details extraction
            self._extract_usb_details()
        
        except Exception as e:
            print(f"Error populating info for {self.port_name}: {e}")

    def _populate_detailed_wmi_info(self):
        """
        Retrieve additional detailed WMI information
        """
        try:
            # Query for driver information
            drivers = self.wmi.Win32_SystemDriver(Name=f"%{self.port_name}%")
            if drivers:
                driver = drivers[0]
                self.driver_info = {
                    'provider': getattr(driver, 'ServiceType', None),
                    'date': getattr(driver, 'InstallDate', None),
                    'version': getattr(driver, 'Version', None)
                }
            
            # Query for signed driver details
            signed_drivers = self.wmi.Win32_PnPSignedDriver(DeviceName=f"%{self.port_name}%")
            if signed_drivers:
                driver = signed_drivers[0]
                self.system_details = {
                    'class_guid': getattr(driver, 'ClassGuid', None),
                    'config_flags': getattr(driver, 'ConfigManagerUserConfig', None),
                    'problem_code': getattr(driver, 'ConfigManagerErrorCode', None)
                }
        
        except Exception as e:
            print(f"Detailed WMI info retrieval error for {self.port_name}: {e}")

    def _populate_registry_details(self):
        """
        Extract detailed information from Windows Registry
        """
        try:
            # Construct potential registry paths
            if self.device_id:
                # Split device ID to construct registry path
                parts = self.device_id.split('\\')
                if len(parts) >= 2:
                    base_path = f"SYSTEM\\CurrentControlSet\\Enum\\{parts[0]}\\{parts[1]}"
                    
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path)
                        
                        # Try to get friendly name
                        try:
                            self.friendly_name = winreg.QueryValueEx(key, "FriendlyName")[0]
                        except:
                            pass
                        
                        self.registry_path = base_path
                    
                    except FileNotFoundError:
                        pass
        
        except Exception as e:
            print(f"Registry details retrieval error for {self.port_name}: {e}")

    def _extract_usb_details(self):
        """
        Extract USB-specific details from hardware IDs
        """
        try:
            for hw_id in self.hardware_ids:
                # Regular expression to extract Vendor and Product IDs
                usb_match = re.search(r'VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})', hw_id)
                if usb_match:
                    self.usb_details = {
                        'vendor_id': usb_match.group(1),
                        'product_id': usb_match.group(2)
                    }
                    break
        except Exception as e:
            print(f"USB details extraction error for {self.port_name}: {e}")

    def get_detailed_info(self):
        """
        Return a dictionary of all collected port information
        """
        return {
            'port_name': self.port_name,
            'description': self.description,
            'manufacturer': self.manufacturer,
            'device_id': self.device_id,
            'friendly_name': self.friendly_name,
            'status': self.status,
            'registry_path': self.registry_path,
            'hardware_ids': self.hardware_ids,
            'compatible_ids': self.compatible_ids,
            'driver_info': self.driver_info,
            'usb_details': self.usb_details,
            'system_details': self.system_details
        }

    def __str__(self):
        """
        String representation of the SerialPort
        """
        return f"SerialPort: {self.port_name} - {self.description or 'Unknown'}"

class WindowsSerialPortInspector:
    def __init__(self):
        """
        Initialize the SerialPortInspector
        """
        self.wmi = wmi.WMI()

    def list_serial_ports(self):
        """
        Retrieve a list of SerialPort objects
        
        Returns:
        list: A list of SerialPort objects
        """
        serial_ports = []
        
        try:
            # Comprehensive WMI query for serial ports
            wmi_query = "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%COM%'"
            ports = self.wmi.query(wmi_query)
            
            # Extract unique COM port names
            seen_ports = set()
            for port in ports:
                # Extract COM port name
                com_match = re.search(r'(COM\d+)', port.Name)
                if com_match:
                    com_port = com_match.group(1)
                    if com_port not in seen_ports:
                        seen_ports.add(com_port)
                        serial_port = SerialPort(com_port)
                        serial_ports.append(serial_port)
            
            # Fallback method using subprocess
            try:
                output = subprocess.check_output('wmic path win32_serialport get deviceid', shell=True).decode()
                com_ports = re.findall(r'(COM\d+)', output)
                
                # Add any unique ports not already in the list
                for port in com_ports:
                    if port not in seen_ports:
                        serial_ports.append(SerialPort(port))
            
            except Exception as enum_error:
                print(f"Alternative port enumeration error: {enum_error}")
        
        except Exception as e:
            print(f"Serial port listing error: {e}")
        
        return serial_ports

# Example usage
if __name__ == "__main__":
    inspector = WindowsSerialPortInspector()
    ports = inspector.list_serial_ports()
    
    for port in ports:
        print("\nDetailed Port Information:")
        detailed_info = port.get_detailed_info()
        for key, value in detailed_info.items():
            print(f"{key}: {value}")