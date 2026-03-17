import argparse
import logging
import sys
import os

from core.rtf_generator import generate_rtf_file
from core.monitor import monitor_process
from core.reporter import generate_report
from core.registry_scanner import get_all_clsids

def setup_logger():
    """Configure basic logging."""
    logger = logging.getLogger('ole_tester')
    logger.setLevel(logging.INFO)
    
    # Create console handler
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    
    # Create formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    ch.setFormatter(formatter)
    
    # Add the handlers to the logger
    logger.addHandler(ch)
    return logger

def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="OLE Object Load & Activation Tester",
        formatter_class=argparse.RawTextHelpFormatter
    )
    
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--clsid', 
        help="A specific CLSID to test (e.g., 0002CE02-0000-0000-C000-000000000046)"
    )
    group.add_argument(
        '--all-clsids', 
        action='store_true',
        help="Scan the local registry and test all installed CLSIDs."
    )
    
    parser.add_argument(
        '--limit',
        type=int,
        default=0,
        help="Limit the number of CLSIDs to test when using --all-clsids (default: 0 = no limit)"
    )

    parser.add_argument(
        '--word-path', 
        type=str, 
        default=None,
        help="Optional: Path to WINWORD.EXE. If not provided, it will attempt to use the default system association."
    )
    
    parser.add_argument(
        '--timeout', 
        type=int, 
        default=10, 
        help="Timeout in seconds to wait for the DLL to load (default: 10)"
    )
    
    parser.add_argument(
        '--format', 
        choices=['console', 'csv', 'json'], 
        default='console',
        help="Output format: console (default), csv, or json"
    )
    
    parser.add_argument(
        '--output',
        type=str,
        default='report',
        help="Base file name for output (e.g., 'report' creates 'report.csv' or 'report.json'). Ignored if format is 'console'."
    )

    return parser.parse_args()

def run_test(clsid: str, args, logger):
    """Runs the test for a single CLSID."""
    rtf_filename = f"test_{clsid}.rtf"
    try:
        # 1. Generate RTF Payload
        logger.info(f"Generating RTF payload: {rtf_filename}")
        generate_rtf_file(clsid, rtf_filename)
        
        # 2. Execute and Monitor
        logger.info("Starting Word and monitoring process...")
        result = monitor_process(rtf_filename, clsid, args.timeout, args.word_path)
        
        # 3. Report Results
        if args.format != 'console':
            generate_report(result, 'console', args.output) # Always show console output as well
        generate_report(result, args.format, args.output)
        
    except Exception as e:
        logger.error(f"An error occurred while testing {clsid}: {e}")
    finally:
        # Cleanup the generated RTF if desired
        if os.path.exists(rtf_filename):
            try:
                os.remove(rtf_filename)
            except OSError:
                pass


def main():
    logger = setup_logger()
    args = parse_args()
    
    logger.info("Starting OLE Object Tester...")
    logger.info(f"Timeout: {args.timeout}s")
    logger.info(f"Output Format: {args.format}")
    
    if args.clsid:
        logger.info(f"Testing single CLSID: {args.clsid}")
        run_test(args.clsid, args, logger)
    elif args.all_clsids:
        logger.info("Testing all CLSIDs found in local registry...")
        clsids = get_all_clsids()
        if args.limit > 0:
            clsids = clsids[:args.limit]
            logger.info(f"Limited testing to first {args.limit} CLSIDs.")
        
        for i, clsid in enumerate(clsids, 1):
            logger.info(f"=== Testing CLSID {i}/{len(clsids)}: {clsid} ===")
            run_test(clsid, args, logger)

if __name__ == "__main__":
    main()
