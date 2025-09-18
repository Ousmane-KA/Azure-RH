# Nettoyage & Connexion Graph
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users.Actions
Import-Module ActiveDirectory

$scopes = "Sites.ReadWrite.All","User.ReadWrite.All","Mail.Send"
Connect-MgGraph -Scopes $scopes -NoWelcome

# IDs SharePoint
$siteId = "73850331-3700-4243-b4af-33d50034b755"
$listId = "2e190d11-8b95-4d97-824e-77267ee879d0"  # JOINER list

# Récupération des utilisateurs SharePoint
$usersFromSharePoint = Get-MgSiteListItem -SiteId $siteId -ListId $listId -ExpandProperty Fields

foreach ($user in $usersFromSharePoint) {
    $fields = $user.Fields.AdditionalProperties

    # Vérifie si TRAITEPOWERSHELL = 1 si non on ignore
    if ($fields.TRAITEPOWERSHELL -ne "1") {
        continue
    }

    $mailUser = $fields.MAILUSER
    if (-not $mailUser) {
        Write-Warning " Utilisateur sans MAILUSER, ignoré."
        continue
    }

    # Extraire SamAccountName depuis l'email
    $samAccountName = $mailUser.Split("@")[0]

    # Recherche dans AD
    $adUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -Properties * -ErrorAction SilentlyContinue

    if ($adUser) {
        try {
            # Mise à jour AD
            Set-ADUser -Identity $adUser.DistinguishedName `
                -Title $fields.INTITULEPOSTE `
                -Department $fields.SERVICE `
                -City $fields.VILLE `
                -OfficePhone $fields.TELEPHONE `
                -MobilePhone $fields.MOBILE `
                -DisplayName "$($fields.PRENOM) $($fields.NOM)"

            Write-Host " Utilisateur '$samAccountName' mis à jour dans AD." -ForegroundColor Green
            

        } catch {
            Write-Error " Erreur mise à jour AD pour '$samAccountName' : $_"
        }
    } else {
        Write-Warning " Utilisateur '$samAccountName' introuvable dans AD."
    }
}

# Déconnexion Graph API
Disconnect-MgGraph
