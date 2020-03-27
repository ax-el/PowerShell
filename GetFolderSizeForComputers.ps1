# Script qui liste le contenu d'un dossier spécifique sur tous les postes d'une OU
myDN="OU=myOU,DC=contoso,DC=local"
myDCServer="myServer.contoso.local"
myComputerPath="C:\Users"

foreach ($Computer in get-ADComputer -server $myDCServer -searchbase $myDN -Filter *)
    {
    if (test-connection.DNSHostName -Quiet)
    {
        $Result = invoke-command -computername.DNSHostName {
            $FolderSize = "{0:N2}" -f ((Get-ChildItem myComputerPath -Recurse |Measure-Object -property length -sum).sum / 1MB)
            return $FolderSize
            }
        Write-Host $Computer.name ";" $Result
        }
    }
