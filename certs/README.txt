CA Certificate Creation Completed!

Files created:
- root-ca.cer: Root CA certificate
- codesign.cer: Code signing certificate (public key)
- codesign.pfx: Code signing certificate (private key)
- deploy-ca.bat: Deployment script
- inno-setup-config.txt: Inno Setup configuration

PFX password:
  The .pfx password is NEVER stored in this repository. Supply it via the
  CODESIGN_PASSWORD environment variable before building:

    cmd :  set CODESIGN_PASSWORD=<pfx password>
    ps  :  $env:CODESIGN_PASSWORD = '<pfx password>'
    permanent (either shell):  setx CODESIGN_PASSWORD "<pfx password>"

  build_and_package.bat / .ps1 skip signing (with a warning) when it is unset.

Usage:
1. Run deploy-ca.bat as administrator
2. Set CODESIGN_PASSWORD (see above)
3. Use inno-setup-config.txt in your .iss file
4. Test: signtool sign /f "codesign.pfx" /p "%CODESIGN_PASSWORD%" /fd sha256 "test.exe"

Certificate Thumbprints:
- Root CA: 462E6CCBC6FCB849853DCB64C062521A03ACBE53
- Code Signing: 72FC4845C79C0C46EB6F5B7D1967679E6B627BB6
- Created: 06/24/2025 14:15:16
