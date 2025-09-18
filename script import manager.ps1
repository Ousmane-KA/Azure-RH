
Import-Module ActiveDirectory


$users = Import-CSV -Path ".\EFREI.CSV"

foreach ($user in $users) {
    # Vérifier que la colonne Manager n'est pas vide
    if ($user.Manager -and $user.Manager.Trim() -ne "") {
        # Construire le filtre pour rechercher le manager (en supposant que la valeur correspond au SamAccountName)
        $managerFilter = "SamAccountName -eq '$($user.Manager.Trim())'"
        $manager = Get-ADUser -Filter $managerFilter -ErrorAction SilentlyContinue

        if ($manager) {
            # Construire le filtre pour rechercher l'utilisateur concerné (en utilisant son SamAccountName)
            $adUserFilter = "SamAccountName -eq '$($user.SamAccountName.Trim())'"
            $adUser = Get-ADUser -Filter $adUserFilter -ErrorAction SilentlyContinue

            if ($adUser) {
                try {
                    # Mettre à jour l'attribut Manager de l'utilisateur avec le DN du manager
                    Set-ADUser -Identity $adUser -Manager $manager.DistinguishedName
                    Write-Host "Le manager de l'utilisateur '$($user.DisplayName)' a été défini sur '$($manager.SamAccountName)'."
                } catch {
                    Write-Error "Erreur lors de la mise à jour de l'utilisateur '$($user.DisplayName)': $_"
                }
            } else {
                Write-Warning "L'utilisateur '$($user.SamAccountName)' n'a pas été trouvé dans Active Directory."
            }
        } else {
            Write-Warning "Le manager '$($user.Manager)' n'a pas été trouvé pour l'utilisateur '$($user.DisplayName)'."
        }
    } else {
        Write-Host "Aucun manager défini pour l'utilisateur '$($user.DisplayName)'."
    }
}