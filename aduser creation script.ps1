Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear(); Clear-Host

# Importer les modules Microsoft Graph nécessaires
Import-Module Microsoft.Graph.Sites
Import-Module Microsoft.Graph.Users.Actions

# Définir les permissions requises
$scopes = "Sites.ReadWrite.All","User.ReadWrite.All","Mail.Send"
Connect-MgGraph -Scopes $scopes -NoWelcome

# Définir les IDs de votre site SharePoint et de votre liste SharePoint contenant les utilisateurs
$siteId = "73850331-3700-4243-b4af-33d50034b755"   # ID du site SharePoint
$listId = "2e190d11-8b95-4d97-824e-77267ee879d0"   # ID de la liste SharePoint contenant les utilisateurs

# Récupérer les utilisateurs depuis la liste SharePoint
$usersFromSharePoint = Get-MgSiteListItem -SiteId $siteId -ListId $listId -ExpandProperty Fields

# Importer le module Active Directory
Import-Module ActiveDirectory

# Définir le DN de base correspondant à votre domaine
$baseDN = "DC=labefrei,DC=fr"

# Fonction pour obtenir ou créer une OU dans un conteneur donné
function Get-OrCreate-OU {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OUName,
        [Parameter(Mandatory = $true)]
        [string]$ParentDN
    )

    if ([string]::IsNullOrEmpty($ParentDN)) {
        Write-Error "Le paramètre ParentDN ne peut être vide."
        return $null
    }

    # Rechercher uniquement l'OU, sans la créer
    $ou = Get-ADOrganizationalUnit -Filter "Name -eq '$OUName'" -SearchBase $ParentDN -ErrorAction SilentlyContinue

    if ($ou) {
        Write-Host "OU '$OUName' trouvée dans '$ParentDN'."

        # Tentative de désactiver la protection contre la suppression accidentelle
        try {
            Set-ADObject -Identity $ou.DistinguishedName -ProtectedFromAccidentalDeletion $false -ErrorAction Stop
            Write-Host "Protection désactivée pour OU '$OUName'."
        } catch {
            Write-Warning "Impossible de désactiver la protection pour OU '$OUName' : $_"
        }

        return $ou
    } else {
        Write-Warning "OU '$OUName' non trouvée dans '$ParentDN'. Aucune création effectuée."
        return $null
    }
}

# Obtenir ou créer l'OU "Production" à la racine du domaine
$prodOUObj = Get-OrCreate-OU -OUName "Production" -ParentDN $baseDN
if (-not $prodOUObj) {
    Write-Error "Impossible de créer ou récupérer l'OU 'Production'. Arrêt du script."
    exit
}

# Définir un mot de passe par défaut pour les utilisateurs
$defaultPassword = "DefaultP@ssw0rd"

# Fonction pour envoyer un e-mail de notification
function Send-NotificationEmail {
    param (
        [string]$userName,
        [string]$deptOU
    )

    # Définir les paramètres du message
    $params = @{
        Message = @{
            Subject = "Utilisateur créé : $userName"
            Body = @{
                ContentType = "HTML"
                Content = "
                Bonjour, <br><br>
                L'utilisateur $userName a été créé dans l'OU '$deptOU'.<br><br>
                Cordialement,<br>
                L'équipe informatique"
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = "lorenzo.giguel@labefrei.fr"  # Remplacez par l'adresse de votre choix
                    }
                }
            )
        }
    }

    # Envoyer l'email via Send-MgUserMail
    Send-MgUserMail -UserId "lorenzo.giguel@labefrei.fr" -BodyParameter $params
}

# Boucle pour créer les utilisateurs depuis SharePoint
foreach ($user in $usersFromSharePoint) {
    # Récupération des champs depuis la liste SharePoint
    $fields = $user.Fields.AdditionalProperties
    
    # Vérifie si l'utilisateur a déjà été traité
if ($fields.TRAITEPOWERSHELL -eq "1") {
    Write-Host "Utilisateur '$($fields.NOM)' déjà traité. Ignoré."
    continue
}

    $givenName = $fields.PRENOM
    $surname = $fields.NOM
    $title = $fields.INTITULEPOSTE
    $department = $fields.SERVICE
    $city = $fields.VILLE
    $emailAddress = $fields.MAILUSER
    $userPrincipalName = $fields.MAILUSER
    $samAccountName = "$($fields.PRENOM).$($fields.NOM)"
    $officePhone = $fields.TELEPHONE
    $mobile = $fields.MOBILE
    $managerEmail = $fields.MAILDUSUPERIEUR

    # Nettoyage du nom de département
    if ($department -eq $null) {
        Write-Warning "⚠️ L'utilisateur '$givenName' n'a pas de département défini. Il sera ignoré."
        continue
    }
    
    # Vérifier et récupérer le champ "Département"
    if ([string]::IsNullOrEmpty($department)) {
        Write-Warning "⚠️ L'utilisateur '$($user.Fields.NOM)' n'a pas de département défini. Il sera ignoré."
        continue
    }

    # Créer ou obtenir l'OU correspondant au département sous "Production"
    $deptOUObj = Get-OrCreate-OU -OUName $department -ParentDN $prodOUObj.DistinguishedName

    if ($deptOUObj -and $deptOUObj.DistinguishedName) {
        try {
            # Création de l'utilisateur dans Active Directory
            New-ADUser -Name "$givenName $surname" `
                -GivenName $givenName `
                -Surname $surname `
                -UserPrincipalName $userPrincipalName `
                -SamAccountName $samAccountName `
                -EmailAddress $emailAddress `
                -OfficePhone $officePhone `
                -MobilePhone $mobile `
                -Title $title `
                -Department $department `
                -Company $company `
                -City $city `
                -Path $deptOUObj.DistinguishedName `
                -AccountPassword (ConvertTo-SecureString $defaultPassword -AsPlainText -Force) `
                -Enabled $true `
                -ChangePasswordAtLogon $true -ErrorAction Stop

            Write-Host "Utilisateur '$givenName' créé dans l'OU '$($deptOUObj.DistinguishedName)'."

            # Envoi de l'e-mail de notification après la création de l'utilisateur
            Send-NotificationEmail -userName "$givenName $surname" -deptOU $deptOUObj.Name

            # Mise à jour de la liste SharePoint : TRAITEPOWERSHELL = 1
            try {
                Update-MgSiteListItemField -SiteId $siteId -ListId $listId -ListItemId $user.Id -BodyParameter @{
                    "TRAITEPOWERSHELL" = "1"
                }
                Write-Host " TRAITEPOWERSHELL mis à jour pour l'utilisateur '$($fields.NOM)'."
            } catch {
                Write-Warning " Impossible de mettre à jour TRAITEPOWERSHELL pour '$($fields.NOM)': $_"
            }
        

        } catch {
            Write-Error "Erreur lors de la création de l'utilisateur '$givenName' : $_"
        }
    } else {
        Write-Error "Échec de création ou récupération de l'OU pour le département '$department'. L'utilisateur '$givenName' n'a pas été créé."
    }

}
