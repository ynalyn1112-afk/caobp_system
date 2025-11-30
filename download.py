import os
import requests
from pathlib import Path

# ----------------------------------------
# Auto-download all required front-end libs
# ----------------------------------------

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

LIBRARIES = {
    "bootstrap": [
        ("https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css", STATIC_DIR / "css/bootstrap.min.css"),
        ("https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js", STATIC_DIR / "js/bootstrap.bundle.min.js"),
    ],
    "bootstrap-icons": [
        ("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css", STATIC_DIR / "css/bootstrap-icons.css"),
        ("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/fonts/bootstrap-icons.woff2", STATIC_DIR / "css/fonts/bootstrap-icons.woff2"),
        ("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/fonts/bootstrap-icons.woff", STATIC_DIR / "css/fonts/bootstrap-icons.woff"),
    ],
    "chartjs": [
        ("https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js", STATIC_DIR / "js/chart.umd.min.js"),
    ]
}


def download_file(url, dest_path):
    """Download and save a file into the given static directory."""
    os.makedirs(dest_path.parent, exist_ok=True)
    print(f"⬇️  Downloading: {url}")
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(dest_path, "wb") as f:
            f.write(response.content)
        print(f"✅ Saved: {dest_path}")
    except requests.RequestException as e:
        print(f"❌ Error downloading {url}: {e}")


def main():
    print("\n🚀 Starting to download all required libraries...\n")
    for lib_name, files in LIBRARIES.items():
        print(f"📦 {lib_name}")
        for url, dest in files:
            download_file(url, dest)
    print("\n✅ All libraries downloaded successfully!\n")
    print("📂 Saved under:", STATIC_DIR)


if __name__ == "__main__":
    main()
