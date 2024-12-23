##### !!!! CHANGE VALUES HERE BEFORE USING !!!! #####
New-Variable -Name "domainName" -Value "test.lan" -Option Constant
New-Variable -Name "path" -Value "C:\shared" -Option Constant

# Installer le module ImportExcel si nécessaire
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

Add-Type -AssemblyName System.Windows.Forms

# Importer la fonction Generate-SecurePassword depuis PasswordGenerator.ps1
. ./AGDLP-utils.ps1
. ./AGDLP-modules.ps1

# Créer une boîte de dialogue pour sélectionner un fichier
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "CSV files (*.csv)|*.csv"
$OpenFileDialog.Title = "Sélectionnez un fichier CSV"

# Afficher la boîte de dialogue et obtenir le chemin du fichier sélectionné
if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $filePath = $OpenFileDialog.FileName
    $fileExtension = [System.IO.Path]::GetExtension($filePath)

    # Lire le fichier CSV et traiter les noms
    $content = (Get-Content $filePath -Encoding Default | ConvertFrom-Csv -Delimiter ';')
    $newContent = @()
    $faultContent = @()

    $domain = $domainName -split "\."
    $domain = "DC=$($domain[0]),DC=$($domain[1])"

    $totalItems = $content.Count
    $progress = 0

    foreach ($row in $content) {
        # Create UO, GG, GL and folder
        $dep = Normalize-String -textInput $row.Departement
        $dep = $dep -split "/"
        if ($dep.Length -gt 1) {
            $departement = $dep[1]
            $subDepartement = $dep[0]
            $group = $subDepartement
            $position = "OU=$subDepartement,OU=$departement,$domain"
            Add-OU -path $domain -name $departement
            Add-UPNSuffix -path $domainName -name "$departement.$(($domainName -split "\.")[1])"
            Add-GG -path "OU=$departement,$domain" -name $departement
            Add-GL -path "OU=$departement,$domain" -name $departement
            Add-Folder -path $path -name $departement
            
            Add-OU -path "OU=$departement,$domain" -name $subDepartement
            Add-GG -path "OU=$subDepartement,OU=$departement,$domain" -name $subDepartement
            Add-GL -path "OU=$subDepartement,OU=$departement,$domain" -name $subDepartement
            Add-Folder -path "$path\$departement" -name $subDepartement
            Add-ADGroupMember -Identity "GG - $departement" -Members "GG - $subDepartement"
        } else {
            $departement = $dep[0]
            $group = $departement
            $position = "OU=$departement,$domain"
            Add-OU -path $domain -name $departement
            Add-UPNSuffix -path $domainName -name "$departement.$(($domainName -split "\.")[1])"
            Add-GG -path "OU=$departement,$domain" -name $departement
            Add-GL -path "OU=$departement,$domain" -name $departement
            Add-Folder -path $path -name $departement
        }

        # Create user
        $nom = Normalize-String -textInput $row.Nom
        $prenom = Normalize-String -textInput $row.Prenom
        $bureau = Normalize-String -textInput $row.Bureau
        $numero = Normalize-String -textInput $row.NInterne
        $description = Normalize-String -textInput $row.Description
        $suffix = "$departement.$(($domainName -split "\.")[1])"
        $identifiant = Generate-Identifiant -firstName $prenom -lastName $nom
        if ($identifiant -eq $null -or ($newContent | Where-Object { $_.Identifiant -eq $identifiant })) {
            $faultContent += $row
            Add-Content -Path "log.txt" -Value "Doublon ou identifiant trop long: $prenom $nom"
        } else {
            $passwordLength = if ($row.Departement -eq "Direction") { 15 } else { 7 }
            $motDePasse = Generate-SecurePassword -length $passwordLength
            $newContent += [PSCustomObject]@{
                Nom = $nom
                Prenom = $prenom
                Identifiant = $identifiant
                "Mot de passe" = $motDePasse
            }
            Add-User -id $identifiant -firstname $prenom -lastname $nom -password $motDePasse -office $bureau -description $description -number $numero -suffix $suffix -group $group -path $position
        }
        $progress++
        Write-Progress -PercentComplete (($progress / $totalItems) * 100) -Activity "Adding items to Active Directory" 
    }
    Write-Host "All OU, GG, GL and users have been created"
    
    # Définir le chemin du nouveau fichier .xlsx
    $newFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($filePath), "EmployesAvecIdentifiants.xlsx")
    $faultFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($filePath), "fault.csv")
    
    # Vérifier si les fichiers existent déjà et les supprimer si nécessaire
    if (Test-Path -Path $newFilePath) {Remove-Item -Path $newFilePath -Force}
    if (Test-Path -Path $faultFilePath) {Remove-Item -Path $faultFilePath -Force}

    # Exporter le contenu fautif dans fault.csv
    $faultContent | Export-Csv -Path $faultFilePath -Delimiter ';' -NoTypeInformation
    Write-Host "Les lignes avec erreurs ont ete exportées vers $faultFilePath"

    try {
        # Exporter les nouvelles données dans un fichier .xlsx
        $newContent | Export-Excel -Path $newFilePath -WorksheetName "Utilisateurs" -AutoSize
        Write-Host "Les noms, prenoms, identifiants et mots de passe ont ete exports vers $newFilePath"
    } catch {
        Write-Host "Erreur lors de la sauvegarde du fichier. Tentative de sauvegarde dans un autre emplacement."
        $alternativePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "EmployesAvecIdentifiants.xlsx")
        $newContent | Export-Excel -Path $alternativePath -WorksheetName "Utilisateurs" -AutoSize
        Write-Host "Les noms, prenoms, identifiants et mots de passe ont ete exportés vers $alternativePath"
    }
    
    # Added common shared folder
    Add-Folder -path $path -name "commun"
    # Add OU for Computers account and move them
    Add-OU -path $domain -name "terminals"
    try {
        redircmp "OU=terminals,$domain"
        Get-ADComputer -SearchBase "CN=Computers,$domain" -Filter * | Move-ADObject -TargetPath "OU=terminals,$domain"
        Write-Host "Computers accounts moved to new destination"
    } catch {Write-Error "An error occured while trying to move the computers -> $_"}
} else {Write-Host "Aucun fichier selectionné."}