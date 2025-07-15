CA Certificate Creation Completed!

Files created:
- root-ca.cer: Root CA certificate
- codesign.cer: Code signing certificate (public key)
- codesign.pfx: Code signing certificate (private key)
- deploy-ca.bat: Deployment script
- inno-setup-config.txt: Inno Setup configuration

Usage:
1. Run deploy-ca.bat as administrator
2. Use inno-setup-config.txt in your .iss file
3. Test: signtool sign /f "codesign.pfx" /p "YourSecurePassword123!" /fd sha256 "test.exe"

Certificate Thumbprints:
- Root CA: 462E6CCBC6FCB849853DCB64C062521A03ACBE53
- Code Signing: 72FC4845C79C0C46EB6F5B7D1967679E6B627BB6
- Created: 06/24/2025 14:15:16
