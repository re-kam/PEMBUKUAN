import sys, os
project = os.path.dirname(os.path.abspath(__file__))
if project not in sys.path:
    sys.path.insert(0, project)
from app import app as application
