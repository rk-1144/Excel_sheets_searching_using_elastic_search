

# # pip install pandas openpyxl

# import msal
# import requests
# import io
# import pandas as pd

# # ==========================
# # CONFIG
# # ==========================
# CLIENT_ID = ""
# AUTHORITY = "https://login.microsoftonline.com/common"
# SCOPES = ["Files.Read.All", "User.Read"]

# ROOT_FOLDER = "Excel"        # OneDrive folder name
# SEARCH_VALUE = "TARGET_TEXT" # value to search inside Excel

# # ==========================
# # AUTHENTICATION
# # ==========================
# app = msal.PublicClientApplication(
#     CLIENT_ID,
#     authority=AUTHORITY
# )

# result = app.acquire_token_interactive(
#     scopes=SCOPES,
#     prompt="select_account"
# )

# if "access_token" not in result:
#     raise Exception("Authentication failed:", result)

# access_token = result["access_token"]
# HEADERS = {"Authorization": f"Bearer {access_token}"}

# # ==========================
# # HELPERS
# # ==========================
# def graph_get(url):
#     r = requests.get(url, headers=HEADERS)
#     if r.status_code != 200:
#         raise Exception(r.json())
#     return r.json()

# def download_file(file_id):
#     url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
#     r = requests.get(url, headers=HEADERS)
#     r.raise_for_status()
#     return r.content

# # ==========================
# # SEARCH LOGIC
# # ==========================
# def search_excel(file_name, file_id):
#     try:
#         excel_bytes = download_file(file_id)
#         df = pd.read_excel(io.BytesIO(excel_bytes))

#         mask = df.astype(str).apply(
#             lambda row: SEARCH_VALUE.lower() in " ".join(row).lower(),
#             axis=1
#         )

#         matches = df[mask]

#         if not matches.empty:
#             print("\n" + "=" * 60)
#             print(f"📄 FILE: {file_name}")
#             print(matches)
#             print("=" * 60)

#     except Exception as e:
#         print(f"❌ Failed reading {file_name}: {e}")

# def traverse_folder(folder_path):
#     url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}:/children"
#     data = graph_get(url)

#     for item in data.get("value", []):
#         if "folder" in item:
#             traverse_folder(f"{folder_path}/{item['name']}")
#         else:
#             name = item["name"]
#             if name.lower().endswith(".xlsx"):
#                 search_excel(name, item["id"])

# # ==========================
# # RUN
# # ==========================
# if __name__ == "__main__":
#     print("🔍 Searching Excel files in OneDrive...")
#     traverse_folder(ROOT_FOLDER)
#     print("\n✅ Done")

# import msal

# CLIENT_ID = "54019c01-0139-4eca-a6ae-a73b24d55f3d"
# AUTHORITY = "https://login.microsoftonline.com/common"
# SCOPES = ["User.Read"]

# app = msal.PublicClientApplication(
#     CLIENT_ID,
#     authority=AUTHORITY
# )

# result = app.acquire_token_interactive(
#     scopes=SCOPES,
#     # redirect_uri="https://login.microsoftonline.com/common/oauth2/nativeclient",
#     prompt="select_account"
# )

# print(result)



import msal
import requests
import io
import pandas as pd

CLIENT_ID = ""
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.Read.All", "User.Read"]

app = msal.PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

# -------- DEVICE CODE FLOW --------
flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise Exception("Failed to create device flow")

print(flow["message"])  
# Example:
# To sign in, use a web browser to open https://microsoft.com/devicelogin
# and enter the code ABCD-EFGH

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise Exception(result)

access_token = result["access_token"]
HEADERS = {"Authorization": f"Bearer {access_token}"}

def list_folder(folder_path):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}:/children"
    r = requests.get(url, headers=HEADERS)
    data = r.json()

    for item in data.get("value", []):
        if "folder" in item:
            list_folder(f"{folder_path}/{item['name']}")
        else:
            name = item["name"]
            if name.lower().endswith(".xlsx"):
                file_id = item["id"]
                content_url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
                file = requests.get(content_url, headers=HEADERS)

                df = pd.read_excel(io.BytesIO(file.content))
                print(f"\n📄 {name}")
                print(df)

list_folder("Excel")

