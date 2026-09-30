"""Inno Setup signing callback. Private keys stay in the configured provider."""
import sys
from signing import sign_windows

if __name__ == "__main__":
    sign_windows(sys.argv[1])
