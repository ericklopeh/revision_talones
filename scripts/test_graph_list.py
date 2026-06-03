import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import requests
from services.graph_storage import graph_headers


url = "https://graph.microsoft.com/v1.0/sites?search=*"

response = requests.get(url, headers=graph_headers(), timeout=30)

print("STATUS:", response.status_code)
print(response.text)
