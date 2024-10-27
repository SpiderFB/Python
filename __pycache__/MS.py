# import subprocess
# import time
# subprocess.run(["start", "ms-windows-store://search/?query=eFootball"], shell=True)


import subprocess

# Replace '9NT1ZBBV6WH6' with the product ID of the app you want to install
product_id = "9NT1ZBBV6WH6"
subprocess.run(["powershell", "-Command", f"winget install --id {product_id}"], shell=True)
