# Nettoyage de la session
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

# Importer les modules Microsoft Graph nécessaires
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users.Actions

# Définir les permissions requises et se connecter à Microsoft Graph
$scopes = "Sites.ReadWrite.All","User.ReadWrite.All","Mail.Send"
Connect-MgGraph -Scopes $scopes -NoWelcome

# Définir les IDs de votre site SharePoint et de la liste Leaver
$siteId = "73850331-3700-4243-b4af-33d50034b755"   # ID du site SharePoint
$listId = "eda7d0cc-0e09-48cb-82a7-4525efc5868f"     # ID de la liste Leaver

# Récupérer les utilisateurs depuis la liste SharePoint (extension des champs)
$usersFromSharePoint = Get-MgSiteListItem -SiteId $siteId -ListId $listId -ExpandProperty Fields

# Importer le module Active Directory
Import-Module ActiveDirectory

# Fonction pour envoyer un e-mail de notification
function Send-NotificationEmail {
    param (
        [string]$userName,
        [string]$action  # Par exemple "désactivé"
    )

    # Définir les paramètres du message
    $params = @{
        Message = @{
            Subject = "Compte utilisateur $action : $userName"
            Body    = @{
                ContentType = "HTML"
                Content     = "
                Bonjour,<br><br>
                Le compte de l'utilisateur <strong>$userName</strong> a été <strong>$action</strong>.<br><br>
                Cordialement,<br>
                L'équipe informatique"
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "lorenzo.giguel@labefrei.fr"  # Remplacez par l'adresse souhaitée
                    }
                }
            )
        }
    }

    # Envoyer l'e-mail via Microsoft Graph
    Send-MgUserMail -UserId "lorenzo.giguel@labefrei.fr" -BodyParameter $params
}

# Boucle pour traiter chaque utilisateur de la liste Leaver
foreach ($user in $usersFromSharePoint) {
    # Récupérer les champs depuis la liste SharePoint (utilisation de AdditionalProperties)
    $fields = $user.Fields.AdditionalProperties

    if ($fields.TRAITEPOWERSHELL -eq "1") {
        Write-Host " Utilisateur '$($fields.NOM)' déjà traité. Ignoré."
        continue
    }
    # Récupérer les valeurs attendues depuis la liste
    $givenName    = $fields.NOM
    $surname      = $fields.PRENOM
    # Utilisation de la colonne MAILDEPART pour l'email de l'utilisateur
    $emailAddress = $fields.MAILDEPART

    # Vérifier que le champ email est renseigné
    if (-not $emailAddress) {
        Write-Warning "Aucun email renseigné pour l'utilisateur '$givenName $surname'. Cet item sera ignoré."
        continue
    }
    
    # Définir le SamAccountName en extrayant la partie avant "@" si l'email en contient une
    if ($emailAddress.Contains("@")) {
        $samAccountName = $emailAddress.Split("@")[0]
    } else {
        $samAccountName = $emailAddress
    }
    
    # Recherche de l'utilisateur dans AD par SamAccountName
    $adUser = Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue

    if ($adUser) {
        try {
            # Désactiver le compte dans AD
            Disable-ADAccount -Identity $adUser.SamAccountName -ErrorAction Stop
            Write-Host "Le compte de l'utilisateur '$givenName $surname' (SamAccountName: $samAccountName) a été désactivé."
            # Envoyer une notification par e-mail indiquant que le compte a été désactivé
            Send-NotificationEmail -userName "$givenName $surname" -action "désactivé"
        } catch {
            Write-Error "Erreur lors de la désactivation de l'utilisateur '$givenName $surname' : $_"
        }

        try {
            Update-MgSiteListItemField -SiteId $siteId -ListId $listId -ListItemId $user.Id -BodyParameter @{
                "TRAITEPOWERSHELL" = "1"
            }
            Write-Host " TRAITEPOWERSHELL mis à jour pour '$givenName $surname'."
        } catch {
            Write-Warning " Erreur mise à jour TRAITEPOWERSHELL pour '$givenName $surname': $_"
        }
    }
    else {
        Write-Output "L'utilisateur '$givenName $surname' (SamAccountName: $samAccountName) n'a pas été trouvé dans AD. Aucun traitement."
    }
}
