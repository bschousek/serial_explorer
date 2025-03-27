import winreg
import wmi
import re
import subprocess
import threading
import queue
import time

class SerialPortProgressTracker:
    def __init__(self):
        """
        Initialize a progress tracker for serial port information retrieval
        """
        self.total_ports = 0
        self.processed_ports = 0
        self.current_port = None
        self.is_complete = False
        self._lock = threading.Lock()

    def update(self, current_port=None):
        """
        Update progress tracking
        """
        with self._lock:
            self.processed_ports += 1
            self.current_port = current_port

    def get_progress(self):
        """
        Get current progress
        """
        with self._lock:
            if self.total_ports == 0:
                return 0.0
            return (self.processed_ports / self.total_ports) * 100

class SerialPort:
    def __init__(self, port_name, progress_tracker=None):
        """
        Initialize a SerialPort object with comprehensive information
        """
        self.wmi = wmi.WMI()
        self.port_name = port_name
        self.progress_tracker = progress_tracker
        
        # Basic information
        self.is_currently_present = False
        
        # Comprehensive port attributes
        self.description = None
        self.manufacturer = None
        self.device_id = None
        self.hardware_ids = []
        self.compatible_ids = []
        self.registry_path = None
        self.friendly_name = None
        self.status = None
        
        # Advanced details
        self.driver_info = {
            'provider': None,
            'date': None,
            'version': None
        }
        
        self.usb_details = {
            'vendor_id': None,
            'product_id': None
        }
        
        self.system_details = {
            'class_guid': None,
            'config_flags': None,
            'problem_code': None
        }

    def _populate_summary_info(self):
        """
        Populate basic summary information quickly
        """
        try:
            # Quick WMI query to check if port is currently present
            wmi_query = f"SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%{self.port_name}%'"
            wmi_ports = self.wmi.query(wmi_query)
            
            if wmi_ports:
                self.is_currently_present = True
                port = wmi_ports[0]
                self.description = port.Description
                self.manufacturer = port.Manufacturer
                self.status = port.Status
            
            # Update progress tracker if provided
            if self.progress_tracker:
                self.progress_tracker.update(self.port_name)
        
        except Exception as e:
            print(f"Summary info error for {self.port_name}: {e}")

    def _populate_detailed_info(self):
        """
        Populate comprehensive port information
        """
        try:
            # Populate additional details if summary indicates presence
            wmi_query = f"SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%{self.port_name}%'"
            wmi_ports = self.wmi.query(wmi_query)
            
            if wmi_ports:
                port = wmi_ports[0]
                
                # Detailed identification
                self.device_id = port.DeviceID
                self.hardware_ids = getattr(port, 'HardwareID', []) or []
                self.compatible_ids = getattr(port, 'CompatibleID', []) or []
                
                # Registry details
                self._populate_registry_details()
                
                # USB and driver details
                self._extract_usb_details()
                self._populate_driver_details()
            
            # Update progress tracker if provided
            if self.progress_tracker:
                self.progress_tracker.update(self.port_name)
        
        except Exception as e:
            print(f"Detailed info error for {self.port_name}: {e}")

    def _populate_registry_details(self):
        """
        Extract detailed information from Windows Registry
        """
        try:
            if self.device_id:
                parts = self.device_id.split('\\')
                if len(parts) >= 2:
                    base_path = f"SYSTEM\\CurrentControlSet\\Enum\\{parts[0]}\\{parts[1]}"
                    
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path)
                        
                        try:
                            self.friendly_name = winreg.QueryValueEx(key, "FriendlyName")[0]
                        except:
                            pass
                        
                        self.registry_path = base_path
                    
                    except FileNotFoundError:
                        pass
        
        except Exception as e:
            print(f"Registry details error for {self.port_name}: {e}")

    def _extract_usb_details(self):
        """
        Extract USB-specific details from hardware IDs
        """
        try:
            for hw_id in self.hardware_ids:
                usb_match = re.search(r'VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})', hw_id)
                if usb_match:
                    self.usb_details = {
                        'vendor_id': usb_match.group(1),
                        'product_id': usb_match.group(2)
                    }
                    break
        except Exception as e:
            print(f"USB details error for {self.port_name}: {e}")

    def _populate_driver_details(self):
        """
        Retrieve driver information
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
            
            # Query for system details
            signed_drivers = self.wmi.Win32_PnPSignedDriver(DeviceName=f"%{self.port_name}%")
            if signed_drivers:
                driver = signed_drivers[0]
                self.system_details = {
                    'class_guid': getattr(driver, 'ClassGuid', None),
                    'config_flags': getattr(driver, 'ConfigManagerUserConfig', None),
                    'problem_code': getattr(driver, 'ConfigManagerErrorCode', None)
                }
        
        except Exception as e:
            print(f"Driver details error for {self.port_name}: {e}")

    def get_summary_info(self):
        """
        Return summary information for the port
        """
        return {
            'port_name': self.port_name,
            'is_currently_present': self.is_currently_present,
            'description': self.description,
            'manufacturer': self.manufacturer,
            'status': self.status
        }

    def get_detailed_info(self):
        """
        Return comprehensive port information
        """
        return {
            'port_name': self.port_name,
            'is_currently_present': self.is_currently_present,
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

class WindowsSerialPortInspector:
    def __init__(self):
        """
        Initialize the SerialPortInspector
        """
        self.wmi = wmi.WMI()

    def _discover_ports(self):
        """
        Discover all potential serial ports
        """
        ports = set()
        
        # WMI Query
        try:
            wmi_query = "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%COM%'"
            for port in self.wmi.query(wmi_query):
                com_match = re.search(r'(COM\d+)', port.Name)
                if com_match:
                    ports.add(com_match.group(1))
        except Exception as e:
            print(f"WMI port discovery error: {e}")
        
        # Subprocess fallback
        try:
            output = subprocess.check_output('wmic path win32_serialport get deviceid', shell=True).decode()
            com_ports = re.findall(r'(COM\d+)', output)
            ports.update(com_ports)
        except Exception as e:
            print(f"Subprocess port discovery error: {e}")
        
        return list(ports)

    def get_present_ports_summary(self, progress_callback=None):
        """
        Get summary information for currently present ports
        
        Args:
            progress_callback (function, optional): Callback to track progress
        
        Returns:
            list: Summary information for present ports
        """
        # Discover ports
        all_ports = self._discover_ports()
        
        # Setup progress tracking
        progress_tracker = SerialPortProgressTracker()
        progress_tracker.total_ports = len(all_ports)
        
        # Store results
        present_ports = []
        
        # Process ports
        for port_name in all_ports:
            port = SerialPort(port_name, progress_tracker)
            port._populate_summary_info()
            
            # Only add currently present ports
            if port.is_currently_present:
                present_ports.append(port.get_summary_info())
            
            # Call progress callback if provided
            if progress_callback:
                progress_callback(progress_tracker.get_progress())
        
        return present_ports

    def get_absent_ports_summary(self, progress_callback=None):
        """
        Get summary information for ports not currently present
        
        Args:
            progress_callback (function, optional): Callback to track progress
        
        Returns:
            list: Summary information for absent ports
        """
        # Discover ports
        all_ports = self._discover_ports()
        
        # Setup progress tracking
        progress_tracker = SerialPortProgressTracker()
        progress_tracker.total_ports = len(all_ports)
        
        # Store results
        absent_ports = []
        
        # Process ports
        for port_name in all_ports:
            port = SerialPort(port_name, progress_tracker)
            port._populate_summary_info()
            
            # Only add absent ports
            if not port.is_currently_present:
                absent_ports.append(port.get_summary_info())
            
            # Call progress callback if provided
            if progress_callback:
                progress_callback(progress_tracker.get_progress())
        
        return absent_ports

    def get_all_ports_detailed_info(self, progress_callback=None):
        """
        Get comprehensive information for all ports
        
        Args:
            progress_callback (function, optional): Callback to track progress
        
        Returns:
            list: Detailed information for all ports
        """
        # Discover ports
        all_ports = self._discover_ports()
        
        # Setup progress tracking
        progress_tracker = SerialPortProgressTracker()
        progress_tracker.total_ports = len(all_ports)
        
        # Store results
        detailed_ports = []
        
        # Process ports
        for port_name in all_ports:
            port = SerialPort(port_name, progress_tracker)
            port._populate_summary_info()
            port._populate_detailed_info()
            
            detailed_ports.append(port.get_detailed_info())
            
            # Call progress callback if provided
            if progress_callback:
                progress_callback(progress_tracker.get_progress())
        
        return detailed_ports

# Example usage
if __name__ == "__main__":
    def print_progress(progress):
        print(f"Progress: {progress:.2f}%")
    
    inspector = WindowsSerialPortInspector()
    
    # Get summary of present ports
    print("Present Ports:")
    present_ports = inspector.get_present_ports_summary(print_progress)
    print(present_ports)
    
    # Get summary of absent ports
    print("\nAbsent Ports:")
    absent_ports = inspector.get_absent_ports_summary(print_progress)
    print(absent_ports)
    
    # Get detailed info for all ports
    print("\nDetailed Port Information:")
    detailed_ports = inspector.get_all_ports_detailed_info(print_progress)
    print(detailed_ports)