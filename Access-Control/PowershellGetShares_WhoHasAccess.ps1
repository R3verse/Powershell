Import-Module ActiveDirectory
Set-ExecutionPolicy RemoteSigned

$mypath = "U:\mytest.csv"

if((Test-Path -Path "U:\mytest.csv" -pathtype Container) -eq $True){
  Write-Host -ForegroundColor red "We found a file: $mypath --> REMOVING!`r`n"
  Remove-Item $mypath -Force
}

#$ShareName = "\\FILESERVER01\D$"
$ShareName = "\\FILESERVER01\Skole"
#$ShareName = "\\FILESERVER02\repair"

$shares = Get-Childitem -path $ShareName |
          Where-Object {$_.PSIsContainer} |
          Get-ACL |
          Select-Object Path -ExpandProperty Access |
          Select Path, FileSystemRights,AccessControlType,IdentityReference
$shares | export-csv $mypath -Delimiter ';' -NoTypeInformation -Encoding UTF8
Add-Content $mypath "" 
Add-Content $mypath "" 
Add-Content $mypath "Disse personer har adgang til share navnet"
Add-Content $mypath "" 
Add-Content $mypath ""

$shares | select -Expand IdentityReference |
  select -Expand Value |
  % {
    $name = $_ -replace '^domain\\'  # <-- replace with actual domain name
    Get-ADObject -filter { Name -eq $name }
  } |
  ? { $_.ObjectClass -eq 'group' } |
  % {
    $_
    Get-ADGroupMember -Identity $_ |
      Get-ADUser -Properties * |
      select Name,SamAccountName,Mail
  } |
  Out-File $mypath -Append -encoding utf8


Write-Host -ForegroundColor Cyan "$($Group.name) is now exported to: $mypath`n"