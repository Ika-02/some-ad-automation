# Définir les caractères possibles pour le mot de passe
$characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()_+[]{}|;:,.<>?"
$lowercase = "abcdefghijklmnopqrstuvwxyz"
$uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
$digits = "0123456789"
$special = "!@#$%&*()_+;:,.<>?"

# Fonction pour générer un mot de passe sécurisé
function Generate-SecurePassword {
    param (
        [int]$length = 7
    )

    if ($length -lt 7) {
        throw "Le mot de passe doit avoir une longueur d'au moins 7 caractères."
    }

    # Initialiser une chaîne vide pour le mot de passe
    $password = ""

    # Ajouter au moins un caractère de chaque type
    $password += $lowercase[(Get-Random -Maximum $lowercase.Length)]
    $password += $uppercase[(Get-Random -Maximum $uppercase.Length)]
    $password += $digits[(Get-Random -Maximum $digits.Length)]
    $password += $special[(Get-Random -Maximum $special.Length)]

    # Générer le reste du mot de passe de la longueur spécifiée
    for ($i = 4; $i -lt $length; $i++) {
        $randomIndex = Get-Random -Maximum $characters.Length
        $password += $characters[$randomIndex]
    }

    # Mélanger les caractères du mot de passe pour plus de sécurité
    $password = [string]::Join("", ($password.ToCharArray() | Sort-Object {Get-Random}))

    return $password
}

# Fonction pour normaliser les chaînes de caractères
function Normalize-String {
    param (
        [Parameter(Mandatory = $true)]
        [string]$textInput
    )
    return ([Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($textInput.ToLower().Replace(" ", "-"))))
}


# Fonction pour générer l'identifiant
function Generate-Identifiant {
    param (
        [string]$firstName,
        [string]$lastName
    )

    $identifiant = "$firstName.$lastName" -replace ' ', ''
    if ($identifiant.Length -gt 20) {
        if ($firstName.Length -ge 1) {
            $identifiant = "$($firstName.Substring(0,1)).$lastName" -replace ' ', ''
        } else {
            $identifiant = $lastName
        }
        if ($identifiant.Length -gt 20) {
            Add-Content -Path "log.txt" -Value "Identifiant trop long: $firstName $lastName"
            return $null
        }
    }

    return $identifiant
}