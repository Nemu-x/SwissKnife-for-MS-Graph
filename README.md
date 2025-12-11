# 🗡️ SwissKnife for Microsoft Graph

**SwissKnife** is a lightweight, offline, cross-platform desktop client for **Microsoft Graph API**, designed primarily for IT administrators who prefer buttons over endless PowerShell scripts.

The tool wraps dozens of common Graph operations into a clean GUI:  
Teams, Channels, Groups, OneDrive, SharePoint, Intune, Admin, Audit & Raw requests — all in one place.

---

## 🚀 Features

### 🛠️ Core Capabilities
- **Authentication via App Registration**  
  Tenant ID · Client ID · Client Secret · `.default` permissions.

### 👥 Microsoft Teams & Groups
- List user’s Teams  
- List channels in a Team  
- Create Standard / Private / Shared channels  
- Add members & owners to Teams and Channels  
- Create Microsoft 365 Groups  
- Add group members / owners  
- Convert Microsoft 365 Group to a Team (Teamify)

### 📁 OneDrive
- List root folder  
- Download files  
- Upload files  
- Work with any user’s OneDrive (delegated via application permissions)

### 🏢 SharePoint
- List all sites / search by keyword  
- List drive root  
- Upload / download files  
- Work with any site by ID

### 👤 Admin Console
- Get user info  
- Block / Unblock user accounts

### 📱 Intune (Device Management)
- List managed devices  
- Device info  
- Wipe  
- Retire  
*(requires appropriate permissions; can’t be fully tested without Intune license)*

### 📊 Audit Logs
- Sign-in logs  
- Directory audit logs  
*(requires appropriate permissions)*

### 🧪 Raw Graph Explorer
- Full manual request tool  
- Supports GET / POST / PATCH / PUT / DELETE  
- Supports JSON bodies  
- Preloaded example queries

---

## 🎨 GUI Highlights

- **Dark & Light themes**
- **Four result views:**  
  - Table  
  - Details (pretty JSON)  
  - Tree  
  - Raw JSON (with syntax highlighting)
- Modern clean design based on `#2E2E2E` dark grey palette  
- Fully cross-platform (Windows / macOS / Linux)

---

## 📦 Downloads

See the **Releases** section for pre-built binaries:

### ✔ Windows — `.exe`  
### ✔ macOS — `.app`  
### ✔ Linux — standalone binary

> If macOS warns that the developer is unknown, right-click → “Open”, or run:  
> `xattr -dr com.apple.quarantine SwissKnifeGraph.app`

---

## 🧰 Installation (from source)

```bash
git clone https://github.com/Nemu-x/SwissKnife-for-MS-Graph
cd SwissKnife-for-MS-Graph
pip install -r requirements.txt
python gui_qt.py
