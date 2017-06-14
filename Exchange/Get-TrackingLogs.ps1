$Type = "Receive"
$Expediteur = "username@domain.tld"
$Serveur = "ServerName"
$Debut = "01/26/2017 10:00:00"
$Fin = "01/27/2017 11:00:00"
 
Get-MessageTrackingLog -Server $Serveur -Start $Debut -End $Fin -Sender $Expediteur -EventID $Type -Resultsize 3000 
