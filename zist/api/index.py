import os
import sys

# Ensure parent directory is in path so main and processor can be imported
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import app
