import winreg
import logging
from typing import List

logger = logging.getLogger('ole_tester')

def get_all_clsids() -> List[str]:
    """
    Enumerates HKEY_CLASSES_ROOT\\CLSID and returns a list of CLSID strings.
    """
    clsids = []
    logger.info("Scanning registry for installed CLSIDs...")
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "CLSID") as key:
            index = 0
            while True:
                try:
                    clsid = winreg.EnumKey(key, index)
                    # A standard CLSID format is {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
                    if clsid.startswith("{") and clsid.endswith("}") and len(clsid) == 38:
                        clsids.append(clsid[1:-1])  # remove braces
                    index += 1
                except OSError:
                    # No more items in the registry key
                    break
    except Exception as e:
        logger.error(f"Failed to enumerate CLSIDs: {e}")
        
    logger.info(f"Found {len(clsids)} CLSIDs in total.")
    return clsids
