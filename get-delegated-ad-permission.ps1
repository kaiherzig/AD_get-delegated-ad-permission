$ViewMode = "grid" #Grid oder csv
$DomainController = "mydc.mydomain.com" #FQDN Domain Controller
$DN = "DC=intra,DC=mydomain,DC=com" #If using CSV specify FilePath
$FilePath = "C:\Users\Administrator\Desktop\query.csv" #FilePath for CSV
$UseFilter = $false #Bool - $true or $false
$Verbose = $true #$true or $false
$ObjectFilter = "(|(objectClass=domain)(objectClass=organizationalUnit)(objectClass=group)(sAMAccountType=805306368)(objectCategory=Computer))"

$baseSearch = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainController/$DN")
$dnSearch = New-Object System.DirectoryServices.DirectorySearcher($bSearch)
$dnSearch.SearchRoot = $baseSearch
$dnSearch.PageSize = 1000
if ($UseFilter -eq $true) {
        $dnSearch.Filter = $ObjectFilter
}
$dnSearch.SearchScope = "Subtree"

$extPerms = `
        '00299570-246d-11d0-a768-00aa006e0529',
'ab721a54-1e2f-11d0-9819-00aa0040529b',
'0'

$result = @()

foreach ($objResult in $dnSearch.FindAll()) {
        $obj = $objResult.GetDirectoryEntry()

        if ($Verbose -eq $true) {
                Write-Host "Searching... " $obj.distinguishedName
        }
        $permissions = $obj.PsBase.ObjectSecurity.GetAccessRules($true, $false, [Security.Principal.NTAccount])
    
        $result += $permissions | Where-Object { `
                        $_.AccessControlType -eq 'Allow' -and ($_.ObjectType -in $extPerms) -and $_.IdentityReference -notin ('NT AUTHORITY\SELF', 'NT AUTHORITY\SYSTEM', 'S-1-5-32-548') `
        } | Select-Object `
        @{n = 'Object'; e = { $obj.distinguishedName } }, 
        @{n = 'Account'; e = { $_.IdentityReference } },
        @{n = 'Permission'; e = { $_.ActiveDirectoryRights } }

}
if ($ViewMode -eq "grid") {
        $result | Out-GridView
}
if ($ViewMode -eq "csv") {
        $result | Export-CSV $FilePath
}