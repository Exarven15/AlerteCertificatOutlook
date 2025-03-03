<#
.SYNOPSIS
    Copyright (c) 2024 RUELLET Mathieu

.DESCRIPTION
    Ce script est protégé par les lois sur le copyright et est la propriété de RUELLET Mathieu.
    Toute utilisation non autorisée de ce script sans l'autorisation expresse de RUELLET Mathieu est interdite.
    Ce script permet l'importation des clients radius vers un autre radius avec comme paramètre le nom du client, 
    l'adresse IP, le nom du vendeur et le mot de passe secret (Bien lire les commentaire pour comprendre ce qu'il faut changer).

.NOTES
    Nom du fichier : alerteCertificatOutlook.ps1
    Auteur        : RUELLET Mathieu
    Version       : 1.0
    Date          : 29/01/2025

.EXAMPLE
    .\alerteCertificatOutlook.ps1
    Il faut juste lancer le script et faire attention aux variables qu'il faut modifier.

.LINK
    __
#>

# Fonction pour calculer les jours restants jusqu'à la date d'expiration d'un certificat
function Get-DaysUntilExpiry {
    param (
        [datetime]$CertExpiryDate
    )
    
    $today = Get-Date
    $daysLeft = ($CertExpiryDate - $today).Days
    return $daysLeft
}

# Fonction pour ajouter un événement à l'agenda Outlook
function Add-OutlookEvent {
    param (
        [string]$Subject,
        [string]$Body,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    $appointment = $calendar.Items.Add()

    $appointment.Subject = $Subject
    $appointment.Body = $Body
    $appointment.Start = $StartDate
    $appointment.End = $EndDate
    $appointment.ReminderSet = $true
    $appointment.ReminderMinutesBeforeStart = 60 # Rappel 1 heure avant l'événement
    $appointment.Save()
}

# Liste de certificats avec leurs dates d'expiration (à remplacer par les données réelles/nouvelles)
$certificates = @{
    "NPS-2" = Get-Date "2025-04-10"
    "S-NPS1-P" = Get-Date "2025-09-16"
    "S-NPS2-P" = Get-Date "2025-09-16"
}

# Nombre de jours avant l'expiration pour planifier les alertes
$alertIntervals = @(60, 30, 15, 7)

# Parcourir la liste de certificats et vérifier les dates d'expiration
foreach ($certName in $certificates.Keys) {
    $certExpiryDate = $certificates[$certName]

    foreach ($interval in $alertIntervals) {
        $alertDate = $certExpiryDate.AddDays(-$interval)
        if ($alertDate -gt (Get-Date)) {
            # Ajouter un événement à l'agenda Outlook, attention les caractères spéciaux peuvent renvoyé un output différent.
            $subject = "Rappel : Expiration du certificat $certName dans $interval jours"
            $body = "Le certificat '$certName' expire le $($certExpiryDate.ToString('dd/MM/yyyy')). Ce rappel est planifier $interval jours avant la date d'expiration."
            $startDate = $alertDate.Date.AddHours(9) # Par exemple, début à 9h le jour de l'alerte
            $endDate = $startDate.AddHours(1) # Durée de l'événement : 1 heure

            Add-OutlookEvent -Subject $subject -Body $body -StartDate $startDate -EndDate $endDate
        }
    }
}
