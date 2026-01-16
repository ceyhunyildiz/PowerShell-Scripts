<#
.SYNOPSIS
Exports Active Directory user attributes to an Excel file.

.DESCRIPTION
Retrieves a wide range of Active Directory user attributes and exports
the collected data to an Excel file for reporting and analysis purposes.

This script is typically used to generate comprehensive user reports
from Active Directory environments.

.AUTHOR
Ceyhun Yıldız

.DATE
2026-01-16
#>


Import-Module ActiveDirectory; $out="C:\AD_PS_Ops\Reports\AD_Kullanicilar_Attribute.csv"; New-Item -ItemType Directory -Path (Split-Path $out) -Force | Out-Null; $props=@("CanonicalName","City","Company","Created","Department","Description","DisplayName","DistinguishedName","EmailAddress","GivenName","LastBadPasswordAttempt","LastLogonDate","Manager","MobilePhone","Office","OfficePhone","POBox","SamAccountName","State","StreetAddress","Surname","Title","ObjectClass","Name"); [System.IO.File]::WriteAllText($out,(Get-ADUser -Filter * -Properties $props | Select-Object SamAccountName,Description,POBox,ObjectClass,DisplayName,GivenName,Surname,Name,CanonicalName,DistinguishedName,Company,Department,Title,Manager,EmailAddress,MobilePhone,OfficePhone,Office,StreetAddress,City,State,Created,LastLogonDate,LastBadPasswordAttempt | ConvertTo-Csv -NoTypeInformation -Delimiter ';' | Out-String),[System.Text.UTF8Encoding]::new($true))
