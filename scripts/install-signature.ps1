$templatePath = Join-Path $PSScriptRoot "..\signatures\signature-template.html"
$userCsvPath = Join-Path $PSScriptRoot "..\signatures\users.csv"
$signatureFolder = Join-Path $env:APPDATA "Microsoft\Signatures"

# Create signature folder if missing
if (!(Test-Path $signatureFolder)) {
    New-Item -ItemType Directory -Path $signatureFolder | Out-Null
}

# Ensure registry keys exist before setting signatures
$registryPathBase = "HKCU:\Software\Microsoft\Office\16.0"
$commonPath = Join-Path $registryPathBase "Common"
$mailSettingsPath = Join-Path $commonPath "MailSettings"

if (-not (Test-Path $commonPath)) {
    New-Item -Path $registryPathBase -Name "Common" | Out-Null
}

if (-not (Test-Path $mailSettingsPath)) {
    New-Item -Path $commonPath -Name "MailSettings" | Out-Null
}

# Read template and user data
$template = Get-Content -Path $templatePath -Raw
$users = Import-Csv -Path $userCsvPath

foreach ($user in $users) {
    # Replace placeholders in template with user info
    $signatureHtml = $template
    foreach ($property in $user.PSObject.Properties) {
        $placeholder = "{" + $property.Name + "}"
        $signatureHtml = $signatureHtml -replace [regex]::Escape($placeholder), $property.Value
    }

    # Normalize file name for signature (remove non-word chars)
    $filename = ($user.FULL_NAME -replace '[^\w]', '_')

    # Save HTML signature file
    $pathHtml = Join-Path $signatureFolder "$filename.htm"
    Set-Content -Path $pathHtml -Value $signatureHtml -Encoding UTF8

    # Create plain text version by stripping HTML tags & decoding entities
    $signatureTxt = $signatureHtml -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
    $pathTxt = Join-Path $signatureFolder "$filename.txt"
    Set-Content -Path $pathTxt -Value $signatureTxt -Encoding UTF8

    # Create basic RTF version (very simple)
    $rtfHeader = "{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}"
    $rtfContent = $signatureTxt -replace '\r?\n', '\\par ' -replace '\\', '\\\\'
    $rtfFooter = "}"
    $signatureRtf = "$rtfHeader\par $rtfContent $rtfFooter"
    $pathRtf = Join-Path $signatureFolder "$filename.rtf"
    Set-Content -Path $pathRtf -Value $signatureRtf -Encoding ASCII

    # Set this signature as default in Outlook for current user
    Set-ItemProperty -Path $mailSettingsPath -Name "NewSignature" -Value $filename
    Set-ItemProperty -Path $mailSettingsPath -Name "ReplySignature" -Value $filename

    Write-Host "Deployed signature for $($user.FULL_NAME) as $filename"
}

Write-Host "✅ All signatures deployed to $signatureFolder"
