Internal Code Signing CA
========================

Files:
- root-ca.cer:  Root CA certificate (public) - deploy to every target machine
- codesign.cer: Code signing certificate (public)
- codesign.pfx: Code signing certificate (private key) - NOT in git
- deploy-ca.bat: Installs root-ca.cer into the Trusted Root store
- inno-setup-config.txt: Inno Setup SignTool reference

Certificate thumbprints (re-issued 2026-07-31):
- Root CA      : 22A3522069358892EC70E14C148811B8FC5BB4AA  (valid to 2036-07-31)
- Code Signing : 0A272F4D67B48757FFD37CC4D63AE63C33F5C7BE  (valid to 2029-07-31)

Why they were re-issued
-----------------------
The previous root CA carried EKU = Client Authentication + Server Authentication
and did NOT include Code Signing (1.3.6.1.5.5.7.3.3). Under EKU nesting rules a
root that lacks an EKU makes the whole chain invalid for that purpose, so every
signature verified as:

    NotValidForUsage - 要求的使用方式的憑證不正確

even though the leaf certificate itself had the correct Code Signing EKU.

The new root is issued with NO EKU extension (unrestricted, which is correct for
a CA); the leaf carries Code Signing only. Signatures now verify as:

    Status: Valid - Signature verified.

Note: Windows' New-SelfSignedCertificate adds Client+Server Auth EKU by default.
When re-issuing, use -Type Custom and do not pass an EKU text extension for the
root, otherwise the same problem returns.

PFX password
------------
Never stored in this repository. Supply it via the CODESIGN_PASSWORD
environment variable before building:

    cmd :  set CODESIGN_PASSWORD=<pfx password>
    ps  :  $env:CODESIGN_PASSWORD = '<pfx password>'
    permanent (either shell):  setx CODESIGN_PASSWORD "<pfx password>"

build_and_package.bat / .ps1 sign both the inner EXE and the installer when it
is set; when unset they skip signing with a warning (the build still succeeds).

The PFX must contain the end-entity certificate ONLY. Exporting with the full
chain makes signtool fail with "Multiple certificates were found that meet all
the given criteria".

    Export-PfxCertificate -Cert $leaf -FilePath codesign.pfx `
        -Password $pw -ChainOption EndEntityCertOnly

Deploying to other machines
---------------------------
1. Run deploy-ca.bat as administrator (installs root-ca.cer to Trusted Root)
2. Machines that trusted the OLD root should remove it - see the thumbprints in
   any backup-* folder here. Without the new root installed, signatures show as
   "unknown publisher" (the signature itself is still intact and timestamped).

Usage
-----
1. Run deploy-ca.bat as administrator
2. Set CODESIGN_PASSWORD (see above)
3. Build: build_and_package.ps1
4. Verify: signtool verify /pa /v <file>
