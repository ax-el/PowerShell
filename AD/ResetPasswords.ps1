
Import-Module ActiveDirectory
$BaseDir="c:\TEMP\" #basedir des chemins relatifs
#Format du fichier : 
#
#SamAccountName;Password; <<= Les colonnes
#SamAccounNames;Passwords; <<= Les données

$UsersFile = $BaseDir+"NewPasswords.csv"
$Users = Import-Csv -Path $UsersFile -Delimiter ";"
foreach ($User in $Users){
  If(Get-ADUser -identity $User.SamAccountname){
    Set-adaccountpassword -identity $User.SamAccountname -reset -newpassword (ConvertTo-SecureString -AsPlainText $User.Password -Force) 
  }
}
