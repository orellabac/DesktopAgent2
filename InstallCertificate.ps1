# This script must be run as an administrator
$cert = New-SelfSignedCertificate -Subject 127.0.0.1 -certstorelocation cert:\LocalMachine\My -dnsname "127.0.0.1","localhost", "xps-mrojas","xps-mrojas.artinsoft.com" -FriendlyName DesktopAgent2 -NotAfter (Get-Date).AddMonths(120)
$pwd = ConvertTo-SecureString -String 'password1234' -Force -AsPlainText
Export-certificate -cert $cert -FilePath c:\junk\certificate\cert.crt
Export-PfxCertificate -cert $cert -FilePath c:\junk\certificate\powershellcert.pfx -Password $pwd
$InstalledCertificate = Import-certificate -FilePath c:\junk\certificate\cert.crt -certstorelocation cert:\LocalMachine\Root
$InstalledCertificate.FriendlyName = "DesktopAgent2"
Set-Content -Path 'c:\junk\certificate\uninstall.ps1' -Value "Get-ChildItem Cert:\LocalMachine\Root\$($InstalledCertificate.ThumbPrint) | Remove-Item"
$InstalledCertificate
