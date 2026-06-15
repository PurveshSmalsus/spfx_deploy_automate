#!/bin/bash

# Exit immediately if a command exits with a non-zero status
set -e

echo "===================================================="
echo "SPFx Azure AD Authentication Certificate Generator"
echo "===================================================="

# Output filenames
PRIVATE_KEY="spfx-private.key"
CSR_FILE="spfx.csr"
PUBLIC_CERT="spfx-public.crt"
PFX_FILE="spfx-certificate.pfx"
BASE64_FILE="spfx-base64.txt"

# 1. Generate Private Key
echo "1. Generating Private Key ($PRIVATE_KEY)..."
openssl genrsa -out "$PRIVATE_KEY" 2048

# 2. Generate Certificate Signing Request (CSR)
echo "2. Generating Certificate Signing Request ($CSR_FILE)..."
echo "Please fill in the certificate subject information below (or press enter for defaults):"
openssl req -new -key "$PRIVATE_KEY" -out "$CSR_FILE" -subj "/CN=SPFxScaffoldingApp/O=Smalsusinfolabs/C=US"

# 3. Generate Public Certificate
echo "3. Generating Public Certificate ($PUBLIC_CERT)..."
openssl x509 -req -days 365 -in "$CSR_FILE" -signkey "$PRIVATE_KEY" -out "$PUBLIC_CERT"

# 4. Export to PFX (Password protected)
echo "4. Exporting to PFX certificate ($PFX_FILE)..."
echo "===================================================="
echo "IMPORTANT: You will now be prompted to enter an export password."
echo "Please REMEMBER this password. You will need it for:"
echo "  - Uploading the certificate to Azure"
echo "  - Adding to GitHub Secrets as 'SHAREPOINT_CERTIFICATE_PASSWORD'"
echo "===================================================="
openssl pkcs12 -export -out "$PFX_FILE" -inkey "$PRIVATE_KEY" -in "$PUBLIC_CERT"

# 5. Convert PFX to Base64
echo "5. Converting PFX certificate to Base64 ($BASE64_FILE)..."
# On macOS, base64 -i works. We can also use cat | base64 for compatibility.
cat "$PFX_FILE" | base64 > "$BASE64_FILE"

echo "===================================================="
echo "SUCCESS! The following files have been created:"
echo "  - $PRIVATE_KEY (Private key - KEEP SAFE & DO NOT COMMIT)"
echo "  - $PUBLIC_CERT (Upload this to your Azure App Registration)"
echo "  - $PFX_FILE (PFX Certificate)"
echo "  - $BASE64_FILE (Copy the content of this file and paste it into"
echo "    your GitHub repository secret named 'SHAREPOINT_CERTIFICATE_BASE64')"
echo "===================================================="
