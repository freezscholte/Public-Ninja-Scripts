$DaysToAlert = 21
$Today = Get-Date
$version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentVersion
if ($Version -lt "6.3") {
    write-host "Unsupported OS. Only Server 2012R2 and up are supported."
    exit 1
}
$CertsBound = get-webbinding | where-object { $_.Protocol -eq "https" }
$Diag = foreach ($Cert in $CertsBound) {
    $CertFile = Get-ChildItem -path "CERT:LocalMachine\$($Cert.CertificateStoreName)" | Where-Object -Property ThumbPrint -eq $cert.certificateHash
    $Diff = (New-TimeSpan -Start $Today -End $CertFile.NotAfter).Days


    if ($diff -lt $DaysToAlert -and $certfile.notbefore -ne $null -and $diff -gt 0) {
        [PSCustomObject]@{
            Friendlyname = $certfile.FriendlyName
            SubjectName  = $Certfile.subject
            CreationDate = $Certfile.NotBefore
            ExpireDate   = $Certfile.NotAfter
            DaysToExp    = $Diff
        }

    }

}

if ($null -eq $diag) {
    $Customfield = "Healthy - No expiring certificates found."
    $Customfield
    Ninja-Property-Set sslCertificates  $Customfield

}
else {
    $i = "Unhealthy - Please check if certificate needs to be renewed"
    $Customfield = $Diag + $i | Format-List | Out-String
    $Customfield
    Ninja-Property-Set sslCertificates  $Customfield

}
