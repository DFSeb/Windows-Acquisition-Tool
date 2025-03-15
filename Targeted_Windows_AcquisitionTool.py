import os
import sys
import time
import platform
import subprocess
import logging
import shutil
from datetime import datetime
import winreg
import ctypes
from ctypes import wintypes

# Set up logging
logging.basicConfig(
    filename='vhd_operation.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def get_system_info():
    """Gather system information and return as a string"""
    info = []
    info.append(f"System: {platform.system()} {platform.version()}")
    info.append(f"Machine: {platform.machine()}")
    info.append(f"Processor: {platform.processor()}")
    info.append(f"Python Version: {platform.python_version()}")
    
    # Get RAM info
    try:
        class MEMORYSTATUSEX(ctypes.Structure):
            _fields_ = [
                ("dwLength", wintypes.DWORD),
                ("dwMemoryLoad", wintypes.DWORD),
                ("ullTotalPhys", ctypes.c_ulonglong),
                ("ullAvailPhys", ctypes.c_ulonglong),
                ("ullTotalPageFile", ctypes.c_ulonglong),
                ("ullAvailPageFile", ctypes.c_ulonglong),
                ("ullTotalVirtual", ctypes.c_ulonglong),
                ("ullAvailVirtual", ctypes.c_ulonglong),
                ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
            ]
        
        memory_status = MEMORYSTATUSEX()
        memory_status.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory_status))
        total_ram = memory_status.ullTotalPhys / (1024 ** 3)
        available_ram = memory_status.ullAvailPhys / (1024 ** 3)
        info.append(f"RAM: {total_ram:.2f} GB total, {available_ram:.2f} GB available")
    except:
        info.append("RAM: Unable to determine")
    
    # Get disk information
    try:
        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        for letter in range(ord('A'), ord('Z') + 1):
            if bitmask & (1 << (letter - ord('A'))):
                drive = f"{chr(letter)}:\\"
                try:
                    total, free = ctypes.c_ulonglong(), ctypes.c_ulonglong()
                    ret = ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                        drive,
                        None,
                        ctypes.byref(total),
                        ctypes.byref(free)
                    )
                    if ret:
                        drives.append(f"Drive {drive} - Total: {total.value / (1024**3):.2f} GB, Free: {free.value / (1024**3):.2f} GB")
                except:
                    drives.append(f"Drive {drive} - Unable to get space information")
        info.append("Drives:")
        for drive in drives:
            info.append(f"  {drive}")
    except:
        info.append("Drives: Unable to determine")
    
    return "\n".join(info)

def is_admin():
    """Check if the script is running with administrator privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def create_vhd(vhd_path, size_mb):
    """Create a VHD file using diskpart"""
    logging.info(f"Creating VHD at {vhd_path} with size {size_mb}MB")
    
    # Create a diskpart script
    script_content = f"""create vdisk file="{vhd_path}" maximum={size_mb} type=expandable
attach vdisk
create partition primary
format fs=ntfs quick label="VHD Storage"
assign letter=V
exit
"""
    
    # Write the script to a temporary file
    script_path = os.path.join(os.environ['TEMP'], 'vhd_script.txt')
    with open(script_path, 'w') as f:
        f.write(script_content)
    
    # Run diskpart with the script
    try:
        result = subprocess.run(['diskpart', '/s', script_path], 
                               capture_output=True, 
                               text=True, 
                               check=True)
        logging.info("VHD created successfully")
        logging.debug(result.stdout)
    except subprocess.CalledProcessError as e:
        logging.error(f"Error creating VHD: {e}")
        logging.error(f"Diskpart output: {e.stdout}")
        logging.error(f"Diskpart error: {e.stderr}")
        raise
    finally:
        # Clean up the temporary script
        if os.path.exists(script_path):
            os.remove(script_path)
    
    return "V:\\"

def detach_vhd(vhd_path):
    """Detach the VHD file using diskpart"""
    logging.info(f"Detaching VHD at {vhd_path}")
    
    # Create a diskpart script
    script_content = f"""select vdisk file="{vhd_path}"
detach vdisk
exit
"""
    
    # Write the script to a temporary file
    script_path = os.path.join(os.environ['TEMP'], 'vhd_detach_script.txt')
    with open(script_path, 'w') as f:
        f.write(script_content)
    
    # Run diskpart with the script
    try:
        result = subprocess.run(['diskpart', '/s', script_path], 
                               capture_output=True, 
                               text=True, 
                               check=True)
        logging.info("VHD detached successfully")
        logging.debug(result.stdout)
    except subprocess.CalledProcessError as e:
        logging.error(f"Error detaching VHD: {e}")
        logging.error(f"Diskpart output: {e.stdout}")
        logging.error(f"Diskpart error: {e.stderr}")
        raise
    finally:
        # Clean up the temporary script
        if os.path.exists(script_path):
            os.remove(script_path)

def calculate_directory_size(path):
    """Calculate the size of a directory in MB"""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if os.path.exists(fp):
                total_size += os.path.getsize(fp)
    
    # Convert to MB and add 20% buffer
    size_mb = int((total_size / (1024 * 1024)) * 1.2)
    return max(100, size_mb)  # Minimum 100MB

def calculate_file_size(path):
    """Calculate the size of a file in MB"""
    size_mb = int((os.path.getsize(path) / (1024 * 1024)) * 1.2)
    return max(100, size_mb)  # Minimum 100MB

def copy_with_metadata(src, dst):
    """Copy a file preserving metadata"""
    logging.info(f"Copying {src} to {dst}")
    
    # First, copy the file
    shutil.copy2(src, dst)
    
    # Try to copy extended attributes and ACLs
    try:
        subprocess.run(['icacls', dst, '/setintegritylevel', 'medium'], 
                      capture_output=True, check=False)
        
        # Copy ACLs
        subprocess.run(['icacls', src, '/save', 'acls.txt'], 
                      capture_output=True, check=False)
        subprocess.run(['icacls', dst, '/restore', 'acls.txt'], 
                      capture_output=True, check=False)
        
        # Clean up
        if os.path.exists('acls.txt'):
            os.remove('acls.txt')
    except Exception as e:
        logging.warning(f"Could not copy extended attributes: {e}")

def main():
    # Check if running as administrator
    if not is_admin():
        logging.error("This script requires administrator privileges")
        print("This script requires administrator privileges. Please run as administrator.")
        input("Press Enter to exit...")
        sys.exit(1)
    
    print("=" * 60)
    print("Windows Folder & File Collector")
    print("=" * 60)
    
    # Get system information
    system_info = get_system_info()
    print("\nGathering system information...")
    
    # Save system information to file
    with open('system_info.txt', 'w') as f:
        f.write(system_info)
    
    print("\nSystem information saved to system_info.txt")
    
    # Prompt for file or directory
    print("\nPlease enter the path to the file or directory you want to store in a VHD:")
    source_path = input("> ").strip()
    
    # Validate path
    if not os.path.exists(source_path):
        logging.error(f"Path not found: {source_path}")
        print(f"Error: Path not found: {source_path}")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Determine if it's a file or directory
    is_directory = os.path.isdir(source_path)
    
    # Calculate required size
    if is_directory:
        size_mb = calculate_directory_size(source_path)
        print(f"\nCalculated size: {size_mb} MB for directory")
    else:
        size_mb = calculate_file_size(source_path)
        print(f"\nCalculated size: {size_mb} MB for file")
    
    # Prompt for VHD location
    print("\nPlease enter the path where you want to save the VHD file:")
    vhd_path = input("> ").strip()
    
    # Create a timestamp for the VHD filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create the VHD filename
    vhd_filename = f"storage_{timestamp}.vhd"
    full_vhd_path = os.path.join(vhd_path, vhd_filename)
    
    print(f"\nCreating VHD: {full_vhd_path}")
    print("This may take a few minutes...")
    
    try:
        # Create the VHD
        vhd_mount_point = create_vhd(full_vhd_path, size_mb)
        
        print(f"\nVHD created and mounted at {vhd_mount_point}")
        
        # Copy files to the VHD
        if is_directory:
            # Create a directory with the same name in the VHD
            base_name = os.path.basename(source_path)
            vhd_target = os.path.join(vhd_mount_point, base_name)
            os.makedirs(vhd_target, exist_ok=True)
            
            # Copy the directory contents
            print(f"\nCopying directory {source_path} to {vhd_target}...")
            
            for root, dirs, files in os.walk(source_path):
                # Create the relative path in the VHD
                rel_path = os.path.relpath(root, source_path)
                vhd_dir = os.path.join(vhd_target, rel_path)
                
                # Create the directory in the VHD
                os.makedirs(vhd_dir, exist_ok=True)
                
                # Copy files in this directory
                for file in files:
                    src_file = os.path.join(root, file)
                    dst_file = os.path.join(vhd_dir, file)
                    copy_with_metadata(src_file, dst_file)
        else:
            # Copy the file to the VHD
            file_name = os.path.basename(source_path)
            vhd_target = os.path.join(vhd_mount_point, file_name)
            
            print(f"\nCopying file {source_path} to {vhd_target}...")
            copy_with_metadata(source_path, vhd_target)
        
        print("\nAll files copied successfully!")
        
        # Create a manifest file
        manifest_path = os.path.join(vhd_mount_point, "manifest.txt")
        with open(manifest_path, 'w') as f:
            f.write(f"Backup created on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Source: {source_path}\n")
            f.write(f"Is Directory: {is_directory}\n\n")
            f.write(system_info)
        
        print("\nManifest file created.")
        
        # Detach the VHD
        print("\nDetaching VHD...")
        detach_vhd(full_vhd_path)
        
        print(f"\nVHD detached. Your data is now stored in: {full_vhd_path}")
        print("\nTo access the VHD, right-click it in File Explorer and select 'Mount'.")
        
    except Exception as e:
        logging.error(f"Error: {e}")
        print(f"\nAn error occurred: {e}")
        
        # Try to detach the VHD if it exists
        try:
            detach_vhd(full_vhd_path)
        except:
            pass
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()