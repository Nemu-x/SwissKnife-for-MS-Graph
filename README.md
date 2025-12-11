# 🗡️ SwissKnife for Microsoft Graph

SwissKnife is a lightweight, cross-platform desktop client for **Microsoft Graph API**, built for IT administrators who prefer clean buttons instead of bulky PowerShell scripts.

Supports Teams, OneDrive, SharePoint, Groups, Admin, Intune, Audit logs & raw Graph queries — all in one place.

---

## 🚀 Features

### 🛠 Core
- Client Credentials authentication (App Registration)
- Dark/Light theme switching
- Multiple result views: **Table / Details / Tree / Raw JSON**
- JSON syntax highlighting

### 👥 Teams & Groups
- List user joined Teams  
- List Team channels  
- Create Standard / Private / Shared channels  
- Add members & owners  
- Create Microsoft 365 Groups  
- Add owners & members  
- Convert Group → Team ("Teamify")

### 📁 OneDrive
- List root folder  
- Upload / download files  
- Works with any user’s OneDrive

### 🏢 SharePoint
- List sites  
- Search sites  
- List site drive  
- Upload / download files

### 👤 Admin
- User info  
- Block / unblock users  

### 📱 Intune
- List managed devices  
- Device info  
- Wipe & retire devices  

### 📊 Audit
- Sign-in logs  
- Directory audit logs  

### 🧪 Raw Graph Editor
- Manual GET / POST / PATCH / PUT / DELETE  
- JSON body support  
- Preloaded example queries  

---

# 🔐 Azure App Registration Setup

Full guide: **[SETUP_AZURE_APP.md](https://github.com/Nemu-x/SwissKnife-for-MS-Graph/blob/main/SETUP_AZURE_APP.md)**

### Quick version:

1. Go to **Azure Portal → Azure Active Directory → App registrations → New registration**
2. Name:  
   `SwissKnife Graph`
3. Supported account type:  
   ✔ Single tenant
4. Redirect URI:  
   _(not required for client credentials)_

### Add a client secret:
- Certificates & secrets → New client secret  
- Copy the value — you’ll need it.

### Add API permissions (Application permissions):

| Area | Permissions |
|-----|-------------|
| Teams & Groups | `Group.ReadWrite.All`, `Directory.ReadWrite.All`, `Team.ReadBasic.All`, `Channel.Create`, `Channel.ReadWrite.All` |
| OneDrive | `Files.ReadWrite.All` |
| SharePoint | `Sites.ReadWrite.All` |
| Mail | `Mail.ReadWrite`, `Mail.Send` |
| Admin | `Directory.ReadWrite.All` |
| Audit | `AuditLog.Read.All` |
| Intune | `DeviceManagementManagedDevices.ReadWrite.All` |

📌 **Click “Grant admin consent”** (important!)

---

# 📦 Installation

### Run from source:
```bash
pip install -r requirements.txt
python gui_qt.py
```

🧡 Support the Project
If the tool helps you automate your work, you can support development:
USDT (TRC20): 0xD9333e859Fb74D885d22E27568589de61E4433b5
BTC:          bc1qkkcgpqym967k2x73al6f7fpvkx52q4rzkut3we
ETH:          0xD9333e859Fb74D885d22E27568589de61E4433b5

🙌 Contributing
Pull requests, feature suggestions, and issue reports are welcome.
Feel free to contribute new modules (e.g., Licensing, Entra configuration, Intune apps, SharePoint sites management, etc.).


✨ Author
Developed with ❤️ by Nemu
GitHub: https://github.com/Nemu-x
Project: https://github.com/Nemu-x/SwissKnife-for-MS-Graph
