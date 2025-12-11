
---

# 📄 2. `SETUP_AZURE_APP.md`


# ☁️ Azure App Registration Setup

SwissKnife uses **Client Credentials Flow**, so it requires an App Registration configured with proper **Application permissions**.

---

## 1️⃣ Create an App Registration

1. Go to **Azure Portal**  
2. Open **Azure Active Directory**  
3. Select **App registrations** → **New registration**
4. Fill in:
   - **Name:** SwissKnife Graph
   - **Supported account types:** *Accounts in this organizational directory only*
   - **Redirect URI:** *(leave empty)*

Click **Register**.

---

## 2️⃣ Create a Client Secret

1. Open your app  
2. Go to **Certificates & secrets**  
3. Click **New client secret**  
4. Copy the **VALUE** — you will not see it again.

---

## 3️⃣ Assign API Permissions

Open:

**API permissions → Add a permission → Microsoft Graph → Application permissions**

### Add all permissions needed:

#### Teams & Groups
- `Directory.ReadWrite.All`
- `Group.ReadWrite.All`
- `Team.ReadBasic.All`
- `Channel.ReadWrite.All`

#### OneDrive
- `Files.ReadWrite.All`

#### SharePoint
- `Sites.ReadWrite.All`

#### Admin
- `Directory.ReadWrite.All`

#### Mail
- `Mail.ReadWrite`
- `Mail.Send`

#### Audit
- `AuditLog.Read.All`

#### Intune
- `DeviceManagementManagedDevices.ReadWrite.All`

---

## 4️⃣ Grant Admin Consent

Click:
**Grant admin consent for <tenant>**

All permissions must show ✔ **Granted**.

---

## 5️⃣ Put the values into SwissKnife

Open the app → Auth section:

Tenant ID: <Directory ID>
Client ID: <Application (client) ID>
Client Secret: <your secret>


Press **Connect** → should show **Connected**.

---

You’re ready to go 🎉
