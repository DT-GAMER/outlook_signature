# Outlook Signature Deployment System

This system automatically creates and installs Outlook email signatures for all users using a central template and CSV.

## Setup (IT/Admins)

1. Clone or download this repo:
   - Public: `git clone https://github.com/DT-GAMER/outlook_signature.git`
   - Or download ZIP and extract

2. Update `signatures/users.csv` with your user data.

3. Customize `signatures/signature-template.html` if needed.

4. Run the PowerShell script:

```powershell
cd .\scripts
.\install-signature.ps1
