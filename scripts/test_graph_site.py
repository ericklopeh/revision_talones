import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from services.graph_storage import get_site_id, get_drive_id


site_id = get_site_id()
print("SITE ID:", site_id)

drive_id = get_drive_id(site_id)
print("DRIVE ID:", drive_id)
