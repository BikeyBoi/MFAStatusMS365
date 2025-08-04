# MFA Status Report for Microsoft 365 Users

This PowerShell script connects to Microsoft Graph and generates an Excel report showing the Multi-Factor Authentication (MFA) status for all enabled users in an Azure AD tenant.

## 📋 Features

- Retrieves all **enabled**, **non-system** users with valid UPNs.
- Detects various authentication methods:
  - Microsoft Authenticator App
  - Phone-based MFA
  - FIDO2 security keys
  - Windows Hello for Business
  - Software OATH tokens
  - Temporary Access Pass
  - Email MFA
- Classifies users as `Enabled`, `Disabled`, or `Error` based on available MFA methods.
- Exports results to a **formatted Excel file** with conditional formatting.
- Cross-platform support (macOS, Windows, Linux).

---

## 🚀 Getting Started

### 1. **Install Required Modules**

The script will automatically install these if missing:

- [Microsoft.Graph.Users](https://www.powershellgallery.com/packages/Microsoft.Graph.Users)
- [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel)

You can also install them manually:

```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module ImportExcel -Scope CurrentUser

