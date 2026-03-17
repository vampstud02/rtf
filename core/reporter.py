import json
import csv
import logging
import os
from typing import Dict, Any

logger = logging.getLogger('ole_tester')

def report_console(result: Dict[str, Any]):
    """Prints the result to the console in a readable format."""
    print("="*40)
    print("OLE Test Result")
    print("="*40)
    print(f"CLSID:       {result.get('clsid')}")
    print(f"Status:      {result.get('status')}")
    
    if result.get('dll_loaded'):
        print(f"Loaded DLL:  {result.get('dll_loaded')}")
        print(f"Time Taken:  {result.get('time_taken')} seconds")
    
    if result.get('message'):
        print(f"Message:     {result.get('message')}")
    print("="*40)

def report_csv(result: Dict[str, Any], output_path: str):
    """Appends the result to a CSV file."""
    file_exists = os.path.isfile(output_path)
    
    with open(output_path, mode='a', newline='', encoding='utf-8') as f:
        fieldnames = ['clsid', 'status', 'dll_loaded', 'time_taken', 'message']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writeheader()
            
        writer.writerow({
            'clsid': result.get('clsid', ''),
            'status': result.get('status', ''),
            'dll_loaded': result.get('dll_loaded', ''),
            'time_taken': result.get('time_taken', ''),
            'message': result.get('message', '')
        })
    logger.info(f"Report appended to {output_path}")

def report_json(result: Dict[str, Any], output_path: str):
    """Appends the result to a JSON array file."""
    data = []
    if os.path.isfile(output_path):
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if content:
                    data = json.loads(content)
        except json.JSONDecodeError:
            logger.warning(f"Could not read existing JSON from {output_path}. Starting fresh.")
            
    data.append(result)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4)
        
    logger.info(f"Report saved to {output_path}")

def generate_report(result: Dict[str, Any], format_type: str, base_filename: str):
    """Routes the result to the appropriate reporter based on chosen format."""
    if format_type == 'console':
        report_console(result)
    elif format_type == 'csv':
        report_csv(result, f"{base_filename}.csv")
    elif format_type == 'json':
        report_json(result, f"{base_filename}.json")
    else:
        logger.error(f"Unknown format type: {format_type}")
