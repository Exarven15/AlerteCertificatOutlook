# Fonction pour supprimer les événements liés aux certificats
function Remove-CertificateAlerts {
    param (
        [string]$Keyword
    )

    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    $items = $calendar.Items

    foreach ($item in $items) {
        if ($item.Subject -like "*$Keyword*") {
            $item.Delete()
        }
    }
}

# Supprimer les alertes existantes
Remove-CertificateAlerts -Keyword "Rappel : Expiration du certificat"