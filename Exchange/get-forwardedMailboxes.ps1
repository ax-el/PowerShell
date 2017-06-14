# Export des adresses pour lesquelles des regles "administrateur" ou "utilisateur" generent un transfert de mail
# @DEEL-V1-2017/06/14

$inboxes = get-mailbox -ResultSize unlimited
$ExportFile="c:\temp\redirection.csv"
foreach ($inbox in $inboxes)
{
    $ExportItems= @()
#Cas des regles "Administrateur" positionnees sur la BAL
    if($inbox.ForwardingAddress -ne $null){
        foreach($Address in $inbox.ForwardingAddress){
        write-host($Address)
            $ExportItems+=[pscustomobject]@{Name=$inbox.name;RuleName="Transfert Administrateur";ForwardTo=$Address;RedirectTo="";ForwardAsAttachmentTo=""}
        }
    }
#Cas des regles "Utilisateur" positionnees sur la BAL
    $rules = get-inboxrule $inbox|select-Object -Property Name, Enabled, ForwardTo,RedirectTo,ForwardAsAttachmentTo
    foreach ($rule in $rules){
        if($rule.ForwardTo -and $rule.Enabled){
            for($i=0; $i -lt $rule.ForwardTo.count;$i++){
                $ExportItems+=[pscustomobject]@{Name=$inbox.name;RuleName=$rule.Name;ForwardTo=$rule.ForwardTo[$i];RedirectTo="";ForwardAsAttachmentTo=""}
            }
        }
        if($rule.RedirectTo -and $rule.Enabled){
           for($i=0; $i -lt $rule.RedirectTo.count;$i++){
                $ExportItems+=[pscustomobject]@{Name=$inbox.name;RuleName=$rule.Name;ForwardTo="";RedirectTo=$rule.RedirectTo[$i];ForwardAsAttachmentTo=""}
            }
        }
        if($rule.ForwardAsAttachmentTo -and $rule.Enabled){
           for($i=0; $i -lt $rule.ForwardAsAttachmentTo.count;$i++){
                $ExportItems+=[pscustomobject]@{Name=$inbox.name;RuleName=$rule.Name;ForwardTo="";RedirectTo="";ForwardAsAttachmentTo=$rule.ForwardAsAttachmentTo[$i]}
            }
        }
    }
$ExportItems | Export-CSV -Path $ExportFile -Delimiter ";" -Append -NoTypeInformation -Encoding UTF8
}
