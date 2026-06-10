# -*- coding: utf-8 -*-
"""
logger.py
---------
Centralized logging system for DocLink.

Logs to both console (stdout) and a persistent log file.
When running from C# tool, console output may not be visible,
so this ensures all diagnostic information is captured to disk.

Public API
----------
LogManager          – singleton logger instance
LogManager.info()   – log info message
LogManager.debug()  – log debug message
LogManager.warning() – log warning message
LogManager.error()  – log error message
LogManager.section() – log section header
LogManager.get_log_path() – get path to current log file
"""

import os
import sys
import datetime
import traceback

__all__ = ["LogManager"]


class _LoggerSingleton(object):
    """Centralized logging that writes to both console and file."""

    def __init__(self):
        self._log_dir = None
        self._log_file = None
        self._initialized = False
        self._init_attempted = False
        self._pyrevit_output = None  # For real-time output window
        try:
            self._init_log()
        except Exception as e:
            # Absolute last resort - at least output something
            try:
                sys.stderr.write("LOGGER INIT FAILED: {}\n".format(str(e)))
                sys.stderr.flush()
            except:
                pass

    def _init_log(self):
        """Initialize the log file in a persistent temp location."""
        if self._initialized or self._init_attempted:
            return

        self._init_attempted = True

        try:
            # Determine log directory - try multiple fallbacks
            log_root = None
            
            # Try 1: APPDATA (most common)
            appdata = os.environ.get("APPDATA")
            if appdata:
                log_root = os.path.join(appdata, "DocLink_Logs")
            
            # Try 2: TEMP as fallback
            if not log_root:
                temp_dir = os.environ.get("TEMP")
                if temp_dir:
                    log_root = os.path.join(temp_dir, "DocLink_Logs")
            
            # Try 3: Current users home
            if not log_root:
                home = os.path.expanduser("~")
                if home and home != "~":
                    log_root = os.path.join(home, "AppData", "Roaming", "DocLink_Logs")
            
            # Try 4: Fallback to relative path
            if not log_root:
                log_root = os.path.abspath("./DocLink_Logs")

            # Ensure directory exists (compatible with older Python)
            if not os.path.exists(log_root):
                try:
                    os.makedirs(log_root)
                except Exception as e:
                    sys.stderr.write("Cannot create log dir {}: {}\n".format(log_root, e))
                    sys.stderr.flush()
                    return

            # Create dated log file
            today = datetime.datetime.now().strftime("%Y%m%d")
            log_filename = "DocLink_Log_{}.txt".format(today)
            self._log_file = os.path.join(log_root, log_filename)
            self._log_dir = log_root

            # Try to write header to verify access
            try:
                with open(self._log_file, "a") as f:
                    f.write("\n" + "=" * 80 + "\n")
                    f.write("DocLink Session Started: {}\n".format(
                        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                    f.write("Python: {}\n".format(sys.version))
                    f.write("Log Directory: {}\n".format(log_root))
                    f.write("=" * 80 + "\n\n")
                    f.flush()
                
                self._initialized = True
                # SUCCESS - write to stderr so we know logger is working
                sys.stderr.write("[LogManager] Initialized: {}\n".format(self._log_file))
                sys.stderr.flush()
                
            except Exception as write_err:
                sys.stderr.write("[LogManager] Cannot write to log file: {}\n".format(write_err))
                sys.stderr.flush()

        except Exception as ex:
            sys.stderr.write("[LogManager INIT ERROR] {}\n".format(str(ex)))
            sys.stderr.flush()

    def _write(self, level, message):
        """Write message to console, file, and optionally pyRevit output window."""
        try:
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            formatted = "[{} {}] {}".format(timestamp, level, message)

            # Always write to console (stderr for visibility)
            try:
                sys.stderr.write(formatted + "\n")
                sys.stderr.flush()
            except Exception:
                try:
                    print(formatted)
                except Exception:
                    pass

            # Write to log file if initialized
            if self._initialized and self._log_file:
                try:
                    with open(self._log_file, "a") as f:
                        f.write(formatted + "\n")
                        f.flush()
                except Exception as ex:
                    try:
                        sys.stderr.write("[LogManager FILE ERROR] {}\n".format(str(ex)))
                        sys.stderr.flush()
                    except Exception:
                        pass

            # Write to pyRevit output window if available
            if self._pyrevit_output:
                try:
                    # Format for pyRevit output with visual indicators
                    if "ERROR" in level or "✗" in message:
                        self._pyrevit_output.print_md("**❌ {}** `{}`".format(level, message))
                    elif "WARN" in level or "⚠" in message:
                        self._pyrevit_output.print_md("**⚠️ {}** `{}`".format(level, message))
                    elif "✓" in message:
                        self._pyrevit_output.print_md("**✅ {}** `{}`".format(level, message))
                    else:
                        self._pyrevit_output.print_md("`[{}] {}`".format(level, message))
                except Exception:
                    # Don't let pyRevit output errors break logging
                    pass

        except Exception as e:
            # Silently fail - don't let logging break the tool
            pass

    def debug(self, message):
        """Log a debug message."""
        self._write("DEBUG", message)

    def info(self, message):
        """Log an info message."""
        self._write("INFO", message)

    def warning(self, message):
        """Log a warning message."""
        self._write("WARN", message)

    def error(self, message):
        """Log an error message."""
        self._write("ERROR", message)

    def section(self, title):
        """Log a section header."""
        sep = "-" * (len(title) + 4)
        self._write("INFO", sep)
        self._write("INFO", "  " + title)
        self._write("INFO", sep)
        
        # Also output to pyRevit window with special formatting
        if self._pyrevit_output:
            try:
                self._pyrevit_output.print_md("---")
                self._pyrevit_output.print_md("## **{}**".format(title))
                self._pyrevit_output.print_md("---")
            except Exception:
                pass

    def set_pyrevit_output(self, output_window):
        """Set the pyRevit output window for real-time logging display.
        
        Args:
            output_window: The pyrevit script output object from get_output()
        """
        self._pyrevit_output = output_window
        self.info("PyRevit output window connected for real-time logging")

    def exception(self, message):
        """Log an exception with traceback."""
        self.error(message)
        tb = traceback.format_exc()
        for line in tb.split("\n"):
            if line.strip():
                self.error(line)

    def get_log_path(self):
        """Return the path to the current log file."""
        return self._log_file if self._initialized else "Log not initialized"

    def get_log_dir(self):
        """Return the directory containing log files."""
        return self._log_dir if self._initialized else "Log not initialized"


# Singleton instance - created on import
LogManager = _LoggerSingleton()
