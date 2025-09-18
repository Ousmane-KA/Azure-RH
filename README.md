# Azure-RH


<img width="1209" height="667" alt="image" src="https://github.com/user-attachments/assets/6df82022-783e-40aa-abdb-742803e0231d" />


---

**automatisation du cycle de vie utilisateur** dans un environnement d’entreprise à travers l’intégration de Microsoft Forms, SharePoint, PowerShell, Active Directory (AD DS) et l’envoi de mails pour la gestion des arrivées (JOINER) et des départs (LEAVER) des employés. 

---
## Objectif du projet

Automatiser la **création et la désactivation des comptes utilisateurs** dans l'Active Directory à partir de formulaires RH, tout en tenant informé le service informatique par mail.

---

## Composants utilisés

| Outil / Système                          | Rôle                                                                                      |
| ---------------------------------------- | ----------------------------------------------------------------------------------------- |
| **Microsoft Forms**                      | Collecte des données RH via formulaires JOINER et LEAVER                                  |
| **SharePoint**                           | Stockage des données issues des formulaires dans des listes JOINER/LEAVER                 |
| **PowerShell (JOINER.ps1 / LEAVER.ps1)** | Scripts d’automatisation pour créer/désactiver les comptes AD et mettre à jour SharePoint |
| **Active Directory (AD DS)**             | Système de gestion des identités utilisateur                                              |
| **Exchange (Outlook)**                   | Envoi d’emails au support informatique (préparation de poste ou restitution)              |
| **Service IT (Support utilisateur)**     | Exécutant les tâches physiques ou manuelles (préparation, récupération du matériel, etc.) |
| **VMware / Windows Server**              | Environnement serveur pour héberger AD et exécuter les scripts                            |

---

## Description du processus

### Partie 1 : Arrivée d’un employé (JOINER)

1. **Saisie du formulaire JOINER** par le service RH via Microsoft Forms.
2. **Stockage automatique** dans une **liste SharePoint JOINER**.
3. Un **script PowerShell JOINER.ps1** est déclenché :

   * Crée l’utilisateur dans **Active Directory**.
   * Met à jour la liste SharePoint avec l’état `TRAITEPOWERSHELL = 1`.
4. Envoie d’un **mail automatique de "préparation de poste"** au support IT.

### Partie 2 : Départ d’un employé (LEAVER)

4. Le service RH saisit un formulaire LEAVER.
5. Stockage dans la **liste SharePoint LEAVER**.
6. Un **script PowerShell LEAVER.ps1** est exécuté :

   * Désactive le compte utilisateur dans AD.
   * Met à jour SharePoint (`TRAITEPOWERSHELL = 1`).
7. Envoi d’un **mail automatique de "restitution de poste"** au support IT.

---

