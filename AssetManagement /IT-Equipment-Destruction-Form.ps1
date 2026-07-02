<#
.SYNOPSIS
    Application PowerShell avec interface graphique permettant de générer automatiquement un rapport de destruction de matériel informatique au format PDF, incluant les informations de l'équipement et des preuves photographiques avant et après destruction.
.DESCRIPTION
    
IT Equipment Destruction Form est un outil destiné aux équipes IT pour documenter et tracer les opérations de destruction de matériel informatique.

L'application permet de saisir les informations d'identification d'un équipement, de sélectionner la méthode de destruction utilisée, d'ajouter des photos avant et après intervention, puis de générer automatiquement un rapport PDF standardisé.

Le rapport produit facilite le suivi du cycle de vie des actifs informatiques, la conservation des preuves de destruction et la conformité aux procédures internes de gestion du matériel.

#>
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.Office.Interop.Word

function Upload-Photo {
    param([string]$initialDirectory)
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $initialDirectory
    $dialog.Filter = "Images (*.jpg;*.png;*.jpeg;*.bmp)|*.jpg;*.png;*.jpeg;*.bmp|All files (*.*)|*.*"
    $dialog.Title = "Select a photo"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    }
    return $null
}

# Création du formulaire principal
$form = New-Object System.Windows.Forms.Form
$form.Text = "IT Equipment Destruction Form"
$form.Size = New-Object System.Drawing.Size(800, 950)  # Augmenté pour s'assurer que tout est visible
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Charger l'icône depuis le même dossier que le script
$iconPath = Join-Path $PSScriptRoot "Novares.ico"
if (Test-Path $iconPath) {
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)
}

# Header
$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size = New-Object System.Drawing.Size(800, 80)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$form.Controls.Add($headerPanel)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "IT EQUIPMENT DESTRUCTION FORM"
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(500, 40)
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$headerPanel.Controls.Add($titleLabel)

# === LOGO + NOVARES dans un sous-panel aligné à droite ===

# Chemin vers le logo
$logoPath = Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "logo.png"

# Sous-panel pour logo + texte
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Size = New-Object System.Drawing.Size(200, 60)
$rightPanel.Location = New-Object System.Drawing.Point(580, 10)
$rightPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)  # même fond que le header

# PictureBox (logo à gauche)
$LogoPictureBox = New-Object System.Windows.Forms.PictureBox
$LogoPictureBox.Size = New-Object System.Drawing.Size(30, 30)
$LogoPictureBox.Location = New-Object System.Drawing.Point(0, 15)
$LogoPictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$LogoPictureBox.Image = [System.Drawing.Image]::FromFile($logoPath)

# Label "NOVARES" (à droite du logo)
$NovaresLabel = New-Object System.Windows.Forms.Label
$NovaresLabel.Text = "NOVARES"
$NovaresLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$NovaresLabel.ForeColor = [System.Drawing.Color]::White
$NovaresLabel.AutoSize = $true
$NovaresLabel.Location = New-Object System.Drawing.Point(40, 17)

# Ajouter logo + texte dans le mini-panel
$rightPanel.Controls.Add($LogoPictureBox)
$rightPanel.Controls.Add($NovaresLabel)

# Ajouter ce panel au header principal
$headerPanel.Controls.Add($rightPanel)


# Logo NOVARES (à droite du bandeau)
$logoLabel = New-Object System.Windows.Forms.Label
$logoLabel.Text = "NOVARES"
$logoLabel.Location = New-Object System.Drawing.Point(650, 25)
$logoLabel.Size = New-Object System.Drawing.Size(130, 30)
try {
    $logoLabel.Font = New-Object System.Drawing.Font("Pirulen", 16, [System.Drawing.FontStyle]::Regular)
} catch {
    # Fallback si Pirulen n'est pas disponible
    $logoLabel.Font = New-Object System.Drawing.Font("Arial Black", 14, [System.Drawing.FontStyle]::Bold)
}
$logoLabel.ForeColor = [System.Drawing.Color]::White
$logoLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$headerPanel.Controls.Add($logoLabel)

# Fonction pour créer des champs de formulaire
function Add-FormField {
    param($labelText, $yPos, $isCombo = $false, $items = $null)
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(20, $yPos)
    $label.Size = New-Object System.Drawing.Size(200, 20)
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($label)
    
    if ($isCombo) {
        $combo = New-Object System.Windows.Forms.ComboBox
        $combo.Location = New-Object System.Drawing.Point(250, $yPos)
        $combo.Size = New-Object System.Drawing.Size(500, 25)
        $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        if ($items) { $combo.Items.AddRange($items) }
        $combo.SelectedIndex = 0
        $form.Controls.Add($combo)
        return $combo
    } else {
        $textbox = New-Object System.Windows.Forms.TextBox
        $textbox.Location = New-Object System.Drawing.Point(250, $yPos)
        $textbox.Size = New-Object System.Drawing.Size(500, 25)
        $form.Controls.Add($textbox)
        return $textbox
    }
}

# Positionnement initial
$yPos = 100

# Champs du formulaire
$textModel = Add-FormField "PC Model*" $yPos; $yPos += 40
$textBrand = Add-FormField "PC Brand*" $yPos; $yPos += 40
$textTag = Add-FormField "Service TAG*" $yPos; $yPos += 40
$textSerial = Add-FormField "HDD Serial*" $yPos; $yPos += 40
$textUser = Add-FormField "User*" $yPos; $yPos += 40

$comboMethod = Add-FormField "Destruction Method*" $yPos $true @("Degaussing", "Shredding", "Drilling", "Crushing", "Other")
$yPos += 40

# Champ spécifique pour "Other"
$labelOtherMethod = New-Object System.Windows.Forms.Label
$labelOtherMethod.Text = "Specify Method*"
$labelOtherMethod.Location = New-Object System.Drawing.Point(20, $yPos)
$labelOtherMethod.Size = New-Object System.Drawing.Size(200, 20)
$labelOtherMethod.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$labelOtherMethod.Visible = $false
$form.Controls.Add($labelOtherMethod)

$textOtherMethod = New-Object System.Windows.Forms.TextBox
$textOtherMethod.Location = New-Object System.Drawing.Point(250, $yPos)
$textOtherMethod.Size = New-Object System.Drawing.Size(500, 25)
$textOtherMethod.Visible = $false
$form.Controls.Add($textOtherMethod)
$yPos += 40

# DatePicker
$dateLabel = New-Object System.Windows.Forms.Label
$dateLabel.Text = "Decommission Date*"
$dateLabel.Location = New-Object System.Drawing.Point(20, $yPos)
$dateLabel.Size = New-Object System.Drawing.Size(200, 20)
$dateLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.Controls.Add($dateLabel)

$datePicker = New-Object System.Windows.Forms.DateTimePicker
$datePicker.Location = New-Object System.Drawing.Point(250, $yPos)
$datePicker.Size = New-Object System.Drawing.Size(200, 25)
$datePicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$datePicker.CustomFormat = "dd/MM/yyyy"
$form.Controls.Add($datePicker)
$yPos += 60

# Photos
$btnBefore = New-Object System.Windows.Forms.Button
$btnBefore.Text = "Upload BEFORE Photo*"
$btnBefore.Location = New-Object System.Drawing.Point(20, $yPos)
$btnBefore.Size = New-Object System.Drawing.Size(200, 35)
$btnBefore.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$btnBefore.ForeColor = [System.Drawing.Color]::White
$btnBefore.FlatStyle = 'Flat'
$form.Controls.Add($btnBefore)

$picBefore = New-Object System.Windows.Forms.PictureBox
$picBefore.Location = New-Object System.Drawing.Point(250, $yPos)
$picBefore.Size = New-Object System.Drawing.Size(500, 150)
$picBefore.BorderStyle = "FixedSingle"
$picBefore.SizeMode = "Zoom"
$form.Controls.Add($picBefore)
$yPos += 170

$btnAfter = New-Object System.Windows.Forms.Button
$btnAfter.Text = "Upload AFTER Photo*"
$btnAfter.Location = New-Object System.Drawing.Point(20, $yPos)
$btnAfter.Size = New-Object System.Drawing.Size(200, 35)
$btnAfter.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$btnAfter.ForeColor = [System.Drawing.Color]::White
$btnAfter.FlatStyle = 'Flat'
$form.Controls.Add($btnAfter)

$picAfter = New-Object System.Windows.Forms.PictureBox
$picAfter.Location = New-Object System.Drawing.Point(250, $yPos)
$picAfter.Size = New-Object System.Drawing.Size(500, 150)
$picAfter.BorderStyle = "FixedSingle"
$picAfter.SizeMode = "Zoom"
$form.Controls.Add($picAfter)
$yPos += 180

# Options
$checkShowPDF = New-Object System.Windows.Forms.CheckBox
$checkShowPDF.Text = "Open PDF after generation"
$checkShowPDF.Location = New-Object System.Drawing.Point(20, $yPos)
$checkShowPDF.Size = New-Object System.Drawing.Size(200, 20)
$checkShowPDF.Checked = $true
$form.Controls.Add($checkShowPDF)
$yPos += 30

# Message requis
$requiredLabel = New-Object System.Windows.Forms.Label
$requiredLabel.Text = "* All fields are required"
$requiredLabel.Location = New-Object System.Drawing.Point(20, $yPos)
$requiredLabel.Size = New-Object System.Drawing.Size(200, 20)
$requiredLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$requiredLabel.ForeColor = [System.Drawing.Color]::Black
$form.Controls.Add($requiredLabel)
$yPos += 30

# Boutons (positionnés pour être toujours visibles)
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "GENERATE PDF REPORT"
$btnGenerate.Location = New-Object System.Drawing.Point(150, $yPos)
$btnGenerate.Size = New-Object System.Drawing.Size(200, 45)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnGenerate)

$btnQuit = New-Object System.Windows.Forms.Button
$btnQuit.Text = "EXIT"
$btnQuit.Location = New-Object System.Drawing.Point(400, $yPos)
$btnQuit.Size = New-Object System.Drawing.Size(200, 45)
$btnQuit.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnQuit.ForeColor = [System.Drawing.Color]::White
$btnQuit.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnQuit)

# Version en bas à droite - Toujours visible
$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Text = "Version A0P01"
$versionLabel.Location = New-Object System.Drawing.Point(680, 870)
$versionLabel.Size = New-Object System.Drawing.Size(100, 20)
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$versionLabel.ForeColor = [System.Drawing.Color]::Black
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::BottomRight
$versionLabel.BringToFront()
$form.Controls.Add($versionLabel)

# Variables pour les photos
$script:beforePhotoPath = $null
$script:afterPhotoPath = $null

# Événements
$btnBefore.Add_Click({
    $path = Upload-Photo -initialDirectory "$env:USERPROFILE\Pictures"
    if ($path) {
        $script:beforePhotoPath = $path
        try {
            $picBefore.Image = [System.Drawing.Image]::FromFile($path)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading image: $_", "Error", "OK", "Error")
            $script:beforePhotoPath = $null
            $picBefore.Image = $null
        }
    }
})

$btnAfter.Add_Click({
    $path = Upload-Photo -initialDirectory "$env:USERPROFILE\Pictures"
    if ($path) {
        $script:afterPhotoPath = $path
        try {
            $picAfter.Image = [System.Drawing.Image]::FromFile($path)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading image: $_", "Error", "OK", "Error")
            $script:afterPhotoPath = $null
            $picAfter.Image = $null
        }
    }
})

$comboMethod.Add_SelectedIndexChanged({
    if ($comboMethod.SelectedItem -eq "Other") {
        $labelOtherMethod.Visible = $true
        $textOtherMethod.Visible = $true
        $form.Height = 980
        $versionLabel.Location = New-Object System.Drawing.Point(680, 900)  # Ajuste la position de la version
    } else {
        $labelOtherMethod.Visible = $false
        $textOtherMethod.Visible = $false
        $form.Height = 950
        $versionLabel.Location = New-Object System.Drawing.Point(680, 870)  # Position normale de la version
    }
    # S'assurer que la version reste au premier plan
    $versionLabel.BringToFront()
})

$btnGenerate.Add_Click({
    # Validation
    $errorMessage = @()
    
    if ([string]::IsNullOrWhiteSpace($textModel.Text)) { $errorMessage += "PC Model" }
    if ([string]::IsNullOrWhiteSpace($textBrand.Text)) { $errorMessage += "PC Brand" }
    if ([string]::IsNullOrWhiteSpace($textTag.Text)) { $errorMessage += "Service TAG" }
    if ([string]::IsNullOrWhiteSpace($textSerial.Text)) { $errorMessage += "HDD Serial" }
    if ([string]::IsNullOrWhiteSpace($textUser.Text)) { $errorMessage += "User" }
    if (-not $script:beforePhotoPath) { $errorMessage += "BEFORE Photo" }
    if (-not $script:afterPhotoPath) { $errorMessage += "AFTER Photo" }
    
    if ($comboMethod.SelectedItem -eq "Other" -and [string]::IsNullOrWhiteSpace($textOtherMethod.Text)) {
        $errorMessage += "Specify Method (when Other is selected)"
    }
    
    if ($errorMessage.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please fill all required fields:`n`n- " + ($errorMessage -join "`n- "), "Warning", "OK", "Warning")
        return
    }

    # Désactiver le bouton et changer le texte pour feedback utilisateur
    $btnGenerate.Enabled = $false
    $originalText = $btnGenerate.Text
    $btnGenerate.Text = "GENERATING PDF..."
    $btnGenerate.BackColor = [System.Drawing.Color]::Gray
    
    # Forcer la mise à jour de l'interface
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    # Génération PDF
    $word = $null
    $doc = $null
    try {
        # Créer Word en mode invisible pour éviter les blocages
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
		$word.DisplayAlerts = [Microsoft.Office.Interop.Word.WdAlertLevel]::wdAlertsNone
        $doc = $word.Documents.Add()
        
        # Forcer le refresh pendant l'opération
        [System.Windows.Forms.Application]::DoEvents()
        
        # Contenu du document
        $selection = $word.Selection
        $selection.TypeText("IT EQUIPMENT DESTRUCTION REPORT`n`n")
        
        # Tableau
        $table = $doc.Tables.Add($selection.Range, 7, 2)
        $table.Borders.Enable = $true
        
        $table.Cell(1,1).Range.Text = "PC Model:"
        $table.Cell(1,2).Range.Text = $textModel.Text
        
        $table.Cell(2,1).Range.Text = "PC Brand:"
        $table.Cell(2,2).Range.Text = $textBrand.Text
        
        $table.Cell(3,1).Range.Text = "Service TAG:"
        $table.Cell(3,2).Range.Text = $textTag.Text
        
        $table.Cell(4,1).Range.Text = "HDD Serial:"
        $table.Cell(4,2).Range.Text = $textSerial.Text
        
        $table.Cell(5,1).Range.Text = "User:"
        $table.Cell(5,2).Range.Text = $textUser.Text
        
        $table.Cell(6,1).Range.Text = "Destruction Method:"
        $methodText = $comboMethod.SelectedItem
        if ($comboMethod.SelectedItem -eq "Other") {
            $methodText += " (" + $textOtherMethod.Text + ")"
        }
        $table.Cell(6,2).Range.Text = $methodText
        
        $table.Cell(7,1).Range.Text = "Decommission Date:"
        $table.Cell(7,2).Range.Text = $datePicker.Value.ToString("dd/MM/yyyy")
        
        # Forcer le refresh après le tableau
        [System.Windows.Forms.Application]::DoEvents()

        # Photos - Tableau 2x2 pour les photos côte à côte
        $selection.EndKey(6) # Aller à la fin du document
        $selection.TypeParagraph()
        $selection.TypeParagraph()
        
        # Créer un tableau 2 lignes x 2 colonnes pour les photos
        $photoTable = $doc.Tables.Add($selection.Range, 2, 2)
        $photoTable.Borders.Enable = $false
        $photoTable.Columns(1).Width = 250
        $photoTable.Columns(2).Width = 250
        
        # Titres des photos
        $photoTable.Cell(1,1).Range.Text = "BEFORE Destruction"
        $photoTable.Cell(1,1).Range.Font.Bold = $true
        $photoTable.Cell(1,1).Range.ParagraphFormat.Alignment = 1  # Centre
        
        $photoTable.Cell(1,2).Range.Text = "AFTER Destruction"
        $photoTable.Cell(1,2).Range.Font.Bold = $true
        $photoTable.Cell(1,2).Range.ParagraphFormat.Alignment = 1  # Centre
        
        # Forcer le refresh avant les images (opération lente)
        $btnGenerate.Text = "INSERTING IMAGES..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Insérer les photos dans les cellules
        $beforeRange = $photoTable.Cell(2,1).Range
        $beforeRange.Text = ""
        $beforeImage = $beforeRange.InlineShapes.AddPicture($script:beforePhotoPath)
        $beforeImage.LockAspectRatio = 1
        $beforeImage.Width = 220
        
        # Refresh entre les deux images
        [System.Windows.Forms.Application]::DoEvents()
        
        $afterRange = $photoTable.Cell(2,2).Range
        $afterRange.Text = ""
        $afterImage = $afterRange.InlineShapes.AddPicture($script:afterPhotoPath)
        $afterImage.LockAspectRatio = 1
        $afterImage.Width = 220

        # Forcer le refresh avant la sauvegarde
        $btnGenerate.Text = "SAVING PDF..."
        [System.Windows.Forms.Application]::DoEvents()

        # Enregistrement - Solution simple et efficace
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "PDF Files (*.pdf)|*.pdf"
        $saveDialog.Title = "Save Destruction Report"
        $saveDialog.FileName = "Destruction_Report_" + (Get-Date -Format "yyyyMMdd") + ".pdf"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Solution directe : Créer d'abord un fichier Word temporaire, puis l'ouvrir et l'exporter
            $tempWordFile = [System.IO.Path]::GetTempFileName() + ".docx"
            
            try {
                # Sauvegarder d'abord en Word
                $doc.SaveAs2($tempWordFile)
                
                # Fermer et rouvrir pour nettoyer les références
                $doc.Close()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
                
                # Rouvrir le document
                $doc = $word.Documents.Open($tempWordFile)
                
                # Export PDF simple
                $doc.SaveAs2($saveDialog.FileName, 17)
                
                # Nettoyer le fichier temporaire
                if (Test-Path $tempWordFile) {
                    Remove-Item $tempWordFile -Force -ErrorAction SilentlyContinue
                }
                
                if ($checkShowPDF.Checked) {
                    Start-Process $saveDialog.FileName
                }
                
                [System.Windows.Forms.MessageBox]::Show("PDF generated successfully!", "Success", "OK", "Information")
                
            } catch {
                # Si ça échoue encore, au moins on sauve en Word
                try {
                    $docxPath = $saveDialog.FileName -replace "\.pdf$", ".docx"
                    $doc.SaveAs2($docxPath)
                    [System.Windows.Forms.MessageBox]::Show("PDF conversion failed, but document saved as:`n$docxPath`n`nYou can manually save it as PDF from Word.", "Partial Success", "OK", "Information")
                    if ($checkShowPDF.Checked) {
                        Start-Process $docxPath
                    }
                } catch {
                    [System.Windows.Forms.MessageBox]::Show("Error saving document: $($_.Exception.Message)", "Error", "OK", "Error")
                }
            }
        }
        
        # Nettoyage COM - Mise à jour pour la nouvelle variable $doc
        if ($doc) {
            $doc.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            $doc = $null
        }
        if ($word) {
            $word.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
            $word = $null
        }
        
        # Forcer le nettoyage mémoire
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # Tuer les processus Word orphelins éventuels
        Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error generating PDF: $($_.Exception.Message)", "Error", "OK", "Error")
        
        # Nettoyage d'urgence en cas d'erreur
        try {
            if ($doc) { $doc.Close(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null }
            if ($word) { $word.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null }
            Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        } catch { }
    }
    finally {
        # Restaurer le bouton dans tous les cas
        $btnGenerate.Enabled = $true
        $btnGenerate.Text = $originalText
        $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        $form.Refresh()
    }
})

$btnQuit.Add_Click({
    try {
        # Libérer les images si elles existent
        if ($picBefore.Image -ne $null) {
            $picBefore.Image.Dispose()
        }
        if ($picAfter.Image -ne $null) {
            $picAfter.Image.Dispose()
        }
        
        # Fermer le formulaire
        $form.Close()
        
        # Forcer la collecte des déchets
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Forcer la sortie de PowerShell
        [System.Environment]::Exit(0)
    }
    catch {
        # En cas d'erreur, forcer la sortie
        [System.Environment]::Exit(0)
    }
})

# Affichage du formulaire
[void]$form.ShowDialog()  je peux le le mettre ou pas
