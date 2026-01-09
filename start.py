#!/usr/bin/env python3
"""
WordTrack Startup Script

This script starts both the proxy server and the add-in debugging session.

Author: Perry Radau
Date: 2024-12-19
Dependencies: Python 3.6+, subprocess, pathlib, signal, time, socket, tempfile
Usage: python3 start.py
       or: ./start.py (after chmod +x start.py)
"""

import os
import sys
import subprocess
import signal
import time
import socket
import tempfile
import shutil
from pathlib import Path
from typing import Optional, Tuple


# Configuration
PROXY_PORT = 3001
DEV_SERVER_PORT = 3000
MAX_PROXY_WAIT = 10
MAX_DEV_SERVER_WAIT = 30
DEFAULT_DOC = Path.home() / "Downloads" / "Default.docx"


class WordTrackStarter:
    """Manages starting and monitoring WordTrack services."""
    
    def __init__(self):
        """Initialize the starter with script directory and process tracking."""
        self.script_dir = Path(__file__).parent.absolute()
        self.proxy_process: Optional[subprocess.Popen] = None
        self.dev_server_process: Optional[subprocess.Popen] = None
        self.addin_process: Optional[subprocess.Popen] = None
        self.proxy_running = False
        self.proxy_log: Optional[Path] = None
        self.dev_server_log: Optional[Path] = None
        
        # Change to script directory
        os.chdir(self.script_dir)
        
        # Set up signal handlers
        signal.signal(signal.SIGINT, self._cleanup_handler)
        signal.signal(signal.SIGTERM, self._cleanup_handler)
    
    def _cleanup_handler(self, signum, frame):
        """Handle cleanup on Ctrl+C or termination signal."""
        print("\nShutting down...")
        self.cleanup()
        sys.exit(0)
    
    def is_port_listening(self, port: int) -> bool:
        """
        Check if a port is listening.
        
        Args:
            port: Port number to check
            
        Returns:
            True if port is listening, False otherwise
        """
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(0.5)
                result = sock.connect_ex(('localhost', port))
                return result == 0
        except Exception:
            return False
    
    def is_process_running(self, process: Optional[subprocess.Popen]) -> bool:
        """
        Check if a process is still running.
        
        Args:
            process: Process object to check
            
        Returns:
            True if process is running, False otherwise
        """
        if process is None:
            return False
        return process.poll() is None
    
    def start_proxy_server(self) -> bool:
        """
        Start the proxy server if not already running.
        
        Returns:
            True if proxy started successfully or was already running, False otherwise
        """
        print("Checking if proxy server is already running...")
        
        if self.is_port_listening(PROXY_PORT):
            print(f"Proxy server already running on port {PROXY_PORT}, skipping...")
            self.proxy_running = True
            return True
        
        print("Starting proxy server in background...")
        
        # Create temporary log file
        self.proxy_log = Path(tempfile.mkstemp(suffix='.log', prefix='wordtrack-proxy-', dir='/tmp')[1])
        print(f"Proxy server output will be logged to: {self.proxy_log}")
        
        try:
            # Start proxy server
            with open(self.proxy_log, 'w') as log_file:
                self.proxy_process = subprocess.Popen(
                    ['npm', 'run', 'proxy'],
                    stdout=log_file,
                    stderr=subprocess.STDOUT,
                    cwd=self.script_dir
                )
            
            self.proxy_running = False
            
            # Wait for proxy to start
            print("Waiting for proxy server to start...")
            for wait_count in range(MAX_PROXY_WAIT):
                time.sleep(1)
                
                # Check if process exited
                if not self.is_process_running(self.proxy_process):
                    print("\nERROR: Proxy server process exited unexpectedly!")
                    print(f"Check the proxy log for errors: {self.proxy_log}")
                    print("\nProxy server output:")
                    if self.proxy_log.exists():
                        print(self.proxy_log.read_text())
                    print("\nCommon issues:")
                    print("  - Port 3001 may be in use: lsof -i :3001")
                    print("  - Node.js version too old (needs v16+): node -v")
                    print("  - Missing dependencies: npm install")
                    return False
                
                # Check if port is listening
                if self.is_port_listening(PROXY_PORT):
                    self.proxy_running = True
                    print(f"Proxy server started successfully on port {PROXY_PORT}")
                    # Show initial proxy output
                    if self.proxy_log.exists() and self.proxy_log.stat().st_size > 0:
                        print("Proxy server output:")
                        print(self.proxy_log.read_text())
                    # Clean up log file after showing output
                    self.proxy_log.unlink()
                    self.proxy_log = None
                    return True
            
            # Timeout
            print(f"\nERROR: Proxy server failed to start after {MAX_PROXY_WAIT} seconds")
            print(f"Check the proxy log for errors: {self.proxy_log}")
            print("\nProxy server output:")
            if self.proxy_log and self.proxy_log.exists():
                print(self.proxy_log.read_text())
            print("\nCommon issues:")
            print("  - Port 3001 may be in use: lsof -i :3001")
            print("  - Node.js version too old (needs v16+): node -v")
            print("  - Missing dependencies: npm install")
            print("  - Certificate issues: npx office-addin-dev-certs install")
            
            # Kill the failed process
            if self.proxy_process:
                self.proxy_process.kill()
            if self.proxy_log and self.proxy_log.exists():
                self.proxy_log.unlink()
            return False
            
        except Exception as e:
            print(f"\nERROR: Failed to start proxy server: {e}")
            if self.proxy_process:
                self.proxy_process.kill()
            return False
    
    def start_addin_debugging(self) -> bool:
        """
        Start the add-in debugging session.
        
        Returns:
            True if started successfully, False otherwise
        """
        print("Starting add-in debugging...")
        print("Starting add-in (will create temp document with WordTrack)...")
        
        try:
            self.addin_process = subprocess.Popen(
                ['npx', 'office-addin-debugging', 'start', 'manifest.xml'],
                cwd=self.script_dir
            )
            return True
        except Exception as e:
            print(f"ERROR: Failed to start add-in debugging: {e}")
            return False
    
    def wait_for_dev_server(self) -> bool:
        """
        Wait for dev server to start on port 3000.
        
        Returns:
            True if dev server started, False otherwise
        """
        print("Waiting for dev server to start on port 3000...")
        
        for wait_count in range(MAX_DEV_SERVER_WAIT):
            time.sleep(1)
            
            if self.is_port_listening(DEV_SERVER_PORT):
                print("Dev server started successfully on port 3000")
                return True
        
        # Dev server didn't start automatically, try manual start
        print(f"\nWARNING: Dev server (port {DEV_SERVER_PORT}) did not start automatically after {MAX_DEV_SERVER_WAIT} seconds")
        print("Attempting to start dev server manually...")
        
        try:
            self.dev_server_log = Path('/tmp/wordtrack-dev-server.log')
            with open(self.dev_server_log, 'w') as log_file:
                self.dev_server_process = subprocess.Popen(
                    ['npm', 'run', 'dev-server'],
                    stdout=log_file,
                    stderr=subprocess.STDOUT,
                    cwd=self.script_dir
                )
            
            # Wait a bit for it to start
            time.sleep(5)
            
            if self.is_port_listening(DEV_SERVER_PORT):
                print("Dev server started successfully on port 3000 (manual start)")
                return True
            else:
                print("\nERROR: Dev server failed to start even manually")
                print(f"Check the dev server log: {self.dev_server_log}")
                print("\nDev server output:")
                if self.dev_server_log.exists():
                    print(self.dev_server_log.read_text())
                else:
                    print("(log file not found)")
                print("\nCommon issues:")
                print("  - Port 3000 may be in use: lsof -i :3000")
                print("  - Node.js version too old (needs v16+): node -v")
                print("  - Missing dependencies: npm install")
                print("  - Certificate issues: npx office-addin-dev-certs install")
                print("\nTry starting manually in a separate terminal:")
                print("  npm run dev-server")
                return False
        except Exception as e:
            print(f"ERROR: Failed to start dev server manually: {e}")
            return False
    
    def find_temp_document_file(self) -> Optional[Path]:
        """
        Find the temp document file created by office-addin-debugging.
        
        Returns:
            Path to the temp file if found, None otherwise
        """
        # Search for files with "Word add-in" in the name in temp directories
        search_dirs = [
            Path("/var/folders"),
            Path("/private/var/folders"),
            Path(tempfile.gettempdir()),
        ]
        
        # Look for files matching the pattern "Word add-in *.docx"
        pattern = "Word add-in*.docx"
        print(f"Searching for temp document file...")
        
        for search_dir in search_dirs:
            if not search_dir.exists():
                continue
            try:
                for found_file in search_dir.rglob(pattern):
                    if found_file.is_file() and found_file.suffix == ".docx":
                        # Check if it was recently modified (within last hour)
                        file_age = time.time() - found_file.stat().st_mtime
                        if file_age < 3600:  # Modified within last hour
                            print(f"Found temp document: {found_file}")
                            return found_file
            except (PermissionError, OSError) as e:
                # Skip directories we can't access
                continue
        
        return None
    
    def monitor_word_and_save_on_close(self):
        """
        Monitor Word process and copy temp document to Default.docx when Word closes.
        """
        print("Monitoring Word process...")
        print("Your work will be saved to Default.docx in ~/Downloads when Word closes.")
        
        # Wait a bit for Word to start
        time.sleep(3)
        
        # Monitor for Word process
        word_running = True
        last_check = time.time()
        
        while word_running:
            time.sleep(2)  # Check every 2 seconds
            
            # Check if Word is still running
            try:
                result = subprocess.run(
                    ['pgrep', '-f', 'Microsoft Word'],
                    capture_output=True,
                    timeout=2
                )
                word_running = result.returncode == 0
            except Exception:
                # If pgrep fails, assume Word might still be running
                pass
            
            # If Word just closed, wait a moment for file system to sync
            if not word_running:
                print("\nWord has closed. Saving your work...")
                time.sleep(1)  # Give file system a moment
                
                # Find and copy the temp document
                temp_file = self.find_temp_document_file()
                
                if temp_file and temp_file.exists():
                    try:
                        print(f"Found your document: {temp_file.name}")
                        print(f"Copying to Default.docx in ~/Downloads...")
                        shutil.copy2(temp_file, DEFAULT_DOC)
                        print(f"\n" + "="*60)
                        print("SUCCESS: Your work has been saved!")
                        print("="*60)
                        print(f"Location: {DEFAULT_DOC}")
                        print(f"File: Default.docx")
                        print("="*60)
                    except Exception as e:
                        print(f"\nError copying file: {e}")
                        print(f"Your original file is still at: {temp_file}")
                else:
                    print("\nWarning: Could not find the temp document file.")
                    print("Your work may have been saved elsewhere.")
                    print("Check Word's recent documents or temp folders.")
                
                break
    
    def monitor_servers(self):
        """
        Monitor both servers and keep the script running.
        
        Periodically checks if servers are still running and reports status.
        """
        first_check = True
        
        while True:
            time.sleep(5)
            
            # Check proxy server
            if not self.proxy_running:
                if self.proxy_process and not self.is_process_running(self.proxy_process):
                    print("\nWARNING: Proxy server process has stopped!")
                    print("The add-in will not be able to connect to Claude API.")
                    print("Restart the script to start the proxy again.")
                    self.proxy_running = True  # Prevent cleanup from trying to kill it
                elif not self.is_port_listening(PROXY_PORT):
                    print("\nWARNING: Proxy server process is running but port 3001 is not listening!")
                    print("The proxy may have crashed or failed to bind to the port.")
                    print("Check for errors and restart the script.")
            
            # Check dev server on first check
            if first_check:
                first_check = False
                if not self.is_port_listening(DEV_SERVER_PORT):
                    print("\nWARNING: Dev server (port 3000) is not running!")
                    print("The add-in cannot load without the dev server.")
                    print("Check the output above for webpack errors.")
                    print("Try running manually: npm run dev-server")
                else:
                    print("Status check: Dev server (port 3000) and proxy server (port 3001) are running.")
    
    def cleanup(self):
        """Clean up all started processes."""
        if not self.proxy_running and self.proxy_process:
            print(f"Stopping proxy server (PID: {self.proxy_process.pid})...")
            self.proxy_process.terminate()
            time.sleep(1)
            if self.is_process_running(self.proxy_process):
                self.proxy_process.kill()
        
        if self.dev_server_process:
            print(f"Stopping dev server (PID: {self.dev_server_process.pid})...")
            self.dev_server_process.terminate()
            time.sleep(1)
            if self.is_process_running(self.dev_server_process):
                self.dev_server_process.kill()
        
        if self.addin_process:
            print(f"Stopping add-in process (PID: {self.addin_process.pid})...")
            self.addin_process.terminate()
            time.sleep(1)
            if self.is_process_running(self.addin_process):
                self.addin_process.kill()
        
        # Clean up log files
        if self.proxy_log and self.proxy_log.exists():
            self.proxy_log.unlink()
        if self.dev_server_log and self.dev_server_log.exists():
            self.dev_server_log.unlink()
    
    def run(self):
        """
        Main execution method.
        
        Starts all services and monitors them.
        """
        print(f"Starting WordTrack from: {self.script_dir}")
        
        # Start proxy server
        if not self.start_proxy_server():
            sys.exit(1)
        
        # Start add-in debugging
        if not self.start_addin_debugging():
            self.cleanup()
            sys.exit(1)
        
        # Wait for dev server
        self.wait_for_dev_server()
        
        # Start background thread to monitor Word and save on close
        import threading
        monitor_thread = threading.Thread(target=self.monitor_word_and_save_on_close, daemon=True)
        monitor_thread.start()
        
        # Wait briefly for add-in process (may exit after starting dev server)
        # Note: wait() with timeout requires Python 3.3+, but we'll use a simple check
        if self.addin_process:
            time.sleep(1)
            # Check if it already exited
            if self.addin_process.poll() is not None:
                pass  # Process already exited
        
        print("\nAdd-in debugging command completed. Dev server may still be running.")
        print(f"Proxy server is still running on port {PROXY_PORT}.")
        print("Press Ctrl+C to stop the proxy server and exit.")
        print()
        
        # Monitor servers
        self.monitor_servers()


def main():
    """Main entry point."""
    starter = WordTrackStarter()
    starter.run()


if __name__ == '__main__':
    main()
