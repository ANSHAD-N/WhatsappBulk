"""
Small reusable utility helpers.
"""

import sys
import time
import random
import subprocess


def log(msg, end="\n"):
    """Print with flush."""
    print(msg, end=end, flush=True)


def fatal(msg):
    """Print error and exit."""
    log(f"ERROR: {msg}")
    sys.exit(1)


def human_delay(min_s=1.0, max_s=3.0):
    """Sleep a random human-like delay."""
    time.sleep(random.uniform(min_s, max_s))


def check_internet():
    """Check internet connectivity."""
    try:
        result = subprocess.run(
            ["ping", "-n", "1", "-w", "3000", "8.8.8.8"],
            capture_output=True, timeout=5,
        )
        return result.returncode == 0
    except Exception:
        return False
