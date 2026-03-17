import os
import time
import subprocess
import psutil
import winreg
import logging

logger = logging.getLogger('ole_tester')

def get_word_path() -> str:
    """Attempt to find the MS Word executable path from the registry."""
    try:
        # Try to find Word in App Paths
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe") as key:
            path, _ = winreg.QueryValueEx(key, "")
            return path
    except FileNotFoundError:
        pass
    
    # Fallback to common locations
    common_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
            
    return ""

def get_dll_for_clsid(clsid_str: str) -> str:
    """
    Attempts to look up the InprocServer32 DLL path for a given CLSID.
    This helps us know exactly what DLL to look for in the memory map.
    """
    clsid_str = f"{{{clsid_str.upper()}}}"
    dll_path = None
    
    # Check HKLM
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Classes\CLSID\{clsid_str}\InprocServer32") as key:
            dll_path, _ = winreg.QueryValueEx(key, "")
            if dll_path:
                return dll_path.lower()
    except FileNotFoundError:
        pass
        
    # Check HKCU
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\CLSID\{clsid_str}\InprocServer32") as key:
            dll_path, _ = winreg.QueryValueEx(key, "")
            if dll_path:
                return dll_path.lower()
    except FileNotFoundError:
        pass
        
    # Sometimes it's a LocalServer32 (EXE) instead of InprocServer32 (DLL), but for OLE it's usually DLL or OCX
    return ""

def kill_word_processes():
    """Kills all existing WINWORD.EXE processes."""
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == 'winword.exe':
                logger.info(f"Killing existing Word process (PID: {proc.info['pid']})")
                proc.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

def monitor_process(rtf_path: str, clsid_str: str, timeout: int, word_path: str = None) -> dict:
    """
    Launches Word with the RTF file and monitors if the target DLL is loaded.
    """
    kill_word_processes()
    time.sleep(1) # Give it a moment to fully close
    
    if not word_path:
        word_path = get_word_path()
        if not word_path:
            return {"status": "Error", "message": "Could not find MS Word executable."}
            
    target_dll = get_dll_for_clsid(clsid_str)
    dll_basename = os.path.basename(target_dll).lower() if target_dll else None
    
    logger.info(f"Target DLL for CLSID: {target_dll if target_dll else 'Unknown (Will check generic loading)'}")
    
    # Launch Word
    logger.info(f"Launching Word: {word_path} \"{os.path.abspath(rtf_path)}\"")
    try:
        # We start Word and keep the Popen object
        process = subprocess.Popen([word_path, os.path.abspath(rtf_path)])
    except Exception as e:
        return {"status": "Error", "message": f"Failed to start Word: {e}"}

    start_time = time.time()
    word_pid = process.pid
    
    result = {
        "status": "Timeout/Not Found",
        "dll_loaded": None,
        "time_taken": timeout,
        "clsid": clsid_str
    }
    
    try:
        ps_proc = psutil.Process(word_pid)
        
        while time.time() - start_time < timeout:
            if not ps_proc.is_running():
                result["status"] = "Process Terminated"
                break
                
            try:
                # Iterate through loaded memory maps (DLLs)
                maps = ps_proc.memory_maps()
                for m in maps:
                    path = m.path.lower()
                    
                    # If we know the exact DLL name from Registry
                    if dll_basename and dll_basename in path:
                        result["status"] = "Success"
                        result["dll_loaded"] = path
                        result["time_taken"] = round(time.time() - start_time, 2)
                        return result
                    
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                # Might happen briefly or if process dies
                pass
                
            time.sleep(0.5) # Poll every 0.5 seconds
            
    except psutil.NoSuchProcess:
        result["status"] = "Process Terminated Early"
        
    finally:
        # Cleanup
        logger.info("Cleaning up Word process...")
        try:
            if psutil.pid_exists(word_pid):
                p = psutil.Process(word_pid)
                p.kill()
        except:
            pass

    # If we didn't find specific DLL, but maybe we didn't know its name (Not in InprocServer32)
    if result["status"] == "Timeout/Not Found" and not dll_basename:
        # This is a grey area. We might just say Not Found.
        result["message"] = "CLSID InprocServer32 not found in registry, couldn't determine target DLL."
        
    return result
