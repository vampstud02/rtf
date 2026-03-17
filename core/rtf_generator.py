import struct
import uuid
import binascii

def clsid_to_bytes(clsid_str: str) -> bytes:
    """
    Converts a CLSID string (e.g., '0002CE02-0000-0000-C000-000000000046')
    to its 16-byte little-endian binary representation.
    """
    try:
        u = uuid.UUID(clsid_str)
        return u.bytes_le
    except ValueError as e:
        raise ValueError(f"Invalid CLSID format: {clsid_str}") from e

def build_rtf_payload(clsid_str: str) -> str:
    """
    Generates an RTF document containing the specified CLSID as an embedded OLE object.
    
    Rather than building a full OLE structured storage, we will use an older technique
    that Word understands: embedding an explicit \objclass [CLSID] with minimal objdata.
    Or, using a standard OLE stream wrapper that triggers object instantiation.
    """
    try:
        # Validate CLSID format
        uuid.UUID(clsid_str)
    except ValueError:
        raise ValueError(f"Invalid CLSID format: {clsid_str}. Must be like 0002CE02-0000-0000-C000-000000000046")

    # Word allows specifying the class directly in RTF via \objclass [CLSID].
    # Then we supply minimal OLE1.0 data.
    
    # OLE Version (4 bytes, 0x00000002)
    # Format ID (4 bytes, 0x00000002 for Embedded)
    # Classname String (Null terminated, let's use the curly brace format for CLSID)
    class_str = f"{{{clsid_str.upper()}}}"
    class_bytes = class_str.encode('ascii') + b"\x00"
    
    ole_ver = struct.pack('<I', 2)
    format_id = struct.pack('<I', 2)
    class_len = struct.pack('<I', len(class_bytes))
    
    # TopicName (0 length), ItemName (0 length), NativeDataSize (0 bytes)
    topic_len = struct.pack('<I', 0)
    item_len  = struct.pack('<I', 0)
    data_size = struct.pack('<I', 0)
    
    ole_data = ole_ver + format_id + class_len + class_bytes + topic_len + item_len + data_size
    hex_data = binascii.hexlify(ole_data).decode('ascii')
    
    # RTF Wrapper
    rtf_template = r"""{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
{\*\generator OLETester;}
\pard\sa200\sl276\slmult1\f0\fs22\lang9 OLE Object Test Document\par
{\object\objemb{\*\objclass %s}\objw1\objh1{\*\objdata %s}}
}"""
    
    return rtf_template % (class_str, hex_data)

def generate_rtf_file(clsid_str: str, output_path: str):
    """
    Saves the generated RTF payload to a file.
    """
    rtf_content = build_rtf_payload(clsid_str)
    with open(output_path, 'w') as f:
        f.write(rtf_content)
