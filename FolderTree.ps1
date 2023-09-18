Add-Type -AssemblyName PresentationFramework, System.Windows.Forms

Set-ExecutionPolicy Unrestricted -Scope Process -Force

function Get-FolderDetails {
    param(
        [string]$FolderPath,
        [switch]$Recurse
    )
  
    $Output = @()

    try {
        $FolderList = if ($Recurse) {
            Get-ChildItem -Recurse -Directory -Path $FolderPath -Force
        } else {
            Get-ChildItem -Directory -Path $FolderPath
        }

        $totalFolders = $FolderList.Count
        $counter = 0

        ForEach ($Folder in $FolderList) {
            $counter++
            $percentage = ($counter / $totalFolders) * 100

            Write-Progress -PercentComplete $percentage -Status "Processing folder $($Folder.FullName)" -Activity "Folder $counter of $totalFolders"

            $Acl = Get-Acl -Path $Folder.FullName

            ForEach ($Access in $Acl.Access) {
                $Properties = [ordered]@{
                    'Folder Name'  = $Folder.FullName
                    'Group/User'   = $Access.IdentityReference
                    'Permissions'  = $Access.FileSystemRights
                    'Inherited'    = $Access.IsInherited
                    'Size(MB)'     = (Get-ChildItem -Path $Folder.FullName -Recurse -File -Force -ErrorAction SilentlyContinue |
                                      Measure-Object -Property Length -Sum).Sum / 1MB
                }
                $Output += New-Object -TypeName PSObject -Property $Properties
            }
        }

    } catch {
        Write-Error "Error processing folders: $_"
    }

    return $Output
}

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
  
$form = New-Object Windows.Forms.Form
$form.Text = 'Select a directory'
$form.Size = New-Object Drawing.Size(450,320)  # Increased form height
$form.StartPosition = 'CenterScreen'
# Checkbox for Size Retrieval
$chkSize = New-Object Windows.Forms.CheckBox
$chkSize.Location = New-Object Drawing.Point(10, 80)
$chkSize.Size = New-Object Drawing.Size(200, 20)
$chkSize.Text = 'Retrieve Folder Sizes'
$chkSize.Checked = $true
$form.Controls.Add($chkSize)

# Input for Depth of Recursion
$labelDepth = New-Object Windows.Forms.Label
$labelDepth.Location = New-Object Drawing.Point(10, 105)
$labelDepth.Size = New-Object Drawing.Size(100, 20)
$labelDepth.Text = 'Recursion Depth:'
$form.Controls.Add($labelDepth)

$txtDepth = New-Object Windows.Forms.TextBox
$txtDepth.Location = New-Object Drawing.Point(110, 105)
$txtDepth.Size = New-Object Drawing.Size(50, 20)
$txtDepth.Text = '0'
$form.Controls.Add($txtDepth)

# Input for Folder Filter
$labelFilter = New-Object Windows.Forms.Label
$labelFilter.Location = New-Object Drawing.Point(10, 125)
$labelFilter.Size = New-Object Drawing.Size(100, 20)
$labelFilter.Text = 'Folder Filter:'
$form.Controls.Add($labelFilter)

$txtFilter = New-Object Windows.Forms.TextBox
$txtFilter.Location = New-Object Drawing.Point(110, 125)
$txtFilter.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($txtFilter)
$folderBrowser = New-Object Windows.Forms.FolderBrowserDialog

# Input for CSV File Path
$labelCsvPath = New-Object Windows.Forms.Label
$labelCsvPath.Location = New-Object Drawing.Point(10, 150)
$labelCsvPath.Size = New-Object Drawing.Size(100, 20)
$labelCsvPath.Text = 'CSV File Path:'
$form.Controls.Add($labelCsvPath)

$txtCsvPath = New-Object Windows.Forms.TextBox
$txtCsvPath.Location = New-Object Drawing.Point(110, 150)
$txtCsvPath.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($txtCsvPath)

$browseCsvButton = New-Object Windows.Forms.Button
$browseCsvButton.Location = New-Object Drawing.Point(320, 150)
$browseCsvButton.Size = New-Object Drawing.Size(35, 20)
$browseCsvButton.Text = '...'
$browseCsvButton.Add_Click({
    $saveFileDialog = New-Object Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = 'CSV Files (*.csv)|*.csv'
    $saveFileDialog.Title = 'Save CSV File'
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtCsvPath.Text = $saveFileDialog.FileName
    }
})
$form.Controls.Add($browseCsvButton)


# Create a ComboBox for selecting search profiles
$profileComboBox = New-Object Windows.Forms.ComboBox
$profileComboBox.Location = New-Object Drawing.Point(10, 250)
$profileComboBox.Size = New-Object Drawing.Size(200, 20)
$form.Controls.Add($profileComboBox)

# Create a Save Profile button
$saveProfileButton = New-Object Windows.Forms.Button
$saveProfileButton.Location = New-Object Drawing.Point(220, 250)
$saveProfileButton.Size = New-Object Drawing.Size(100, 23)
$saveProfileButton.Text = 'Save Profile'
$form.Controls.Add($saveProfileButton)

# Create a Load Profile button
$loadProfileButton = New-Object Windows.Forms.Button
$loadProfileButton.Location = New-Object Drawing.Point(320, 250)
$loadProfileButton.Size = New-Object Drawing.Size(100, 23)
$loadProfileButton.Text = 'Load Profile'
$form.Controls.Add($loadProfileButton)
$profileComboBox.Items.Clear()
$profileFiles = Get-ChildItem -Path "profiles" -Filter "*.xml" -File
$profileComboBox.Items.AddRange($profileFiles.Name)
# Create a Profile Data object to store the profile information
$profileData = New-Object PSObject -Property @{
    'FolderPath' = ''
    'IncludeSubfolders' = $true
    'RetrieveSizes' = $true
    'RecursionDepth' = 0
    'FolderFilter' = ''
    'CSVFilePath' = ''
    'ProfileName' = ''
}

# Function to display a custom input box dialog
function Show-InputBoxDialog {
    param (
        [string]$prompt,
        [string]$title,
        [string]$default
    )
    
    $form = New-Object Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object Drawing.Size(330, 130)
    #$form.StartPosition = 'CenterScreen'
    
    $label = New-Object Windows.Forms.Label
    $label.Location = New-Object Drawing.Point(10, 10)
    $label.Size = New-Object Drawing.Size(280, 20)
    $label.Text = $prompt
    $form.Controls.Add($label)
    
    $textBox = New-Object Windows.Forms.TextBox
    $textBox.Location = New-Object Drawing.Point(10, 40)
    $textBox.Size = New-Object Drawing.Size(260, 20)
    $textBox.Text = $default
    $form.Controls.Add($textBox)
    
    $okButton = New-Object Windows.Forms.Button
    $okButton.Location = New-Object Drawing.Point(210, 70)
    $okButton.Size = New-Object Drawing.Size(75, 23)
    $okButton.Text = 'OK'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object Windows.Forms.Button
    $cancelButton.Location = New-Object Drawing.Point(130, 70)
    $cancelButton.Size = New-Object Drawing.Size(75, 23)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text
    } else {
        $null
    }
}

$saveProfileButton.Add_Click({
    $profileName = Show-InputBoxDialog -prompt 'Enter a profile name:' -title 'Save Profile' -default ''
    if ($profileName) {
        $profileData.FolderPath = $textBox.Text
        $profileData.IncludeSubfolders = $checkBox.Checked
        $profileData.RetrieveSizes = $chkSize.Checked
        $profileData.RecursionDepth = [int]$txtDepth.Text
        $profileData.FolderFilter = $txtFilter.Text
        $profileData.CSVFilePath = $txtCsvPath.Text
        $profileData.ProfileName = $profileName

        # Save the profile data to a text file
        $profileData | Export-Clixml -Path "profiles\$profileName.xml"
        [System.Windows.Forms.MessageBox]::Show('Profile saved successfully!', 'Save Profile', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Clear and reload the profiles in the ComboBox
        $profileComboBox.Items.Clear()
        $profileFiles = Get-ChildItem -Path "profiles" -Filter "*.xml" -File
        $profileComboBox.Items.AddRange($profileFiles.Name)
    }
})

# Function to load a saved profile
$loadProfileButton.Add_Click({
    $selectedProfile = $profileComboBox.SelectedItem
    if ($selectedProfile) {
        try {
            $profileName = $selectedProfile.ToString()
            $profileData = Import-Clixml -Path "profiles\$profileName"
            
            # Update the form controls with the loaded profile data
            $pathPrompt = $profileData.FolderPath
            $includeSubfolders = $profileData.IncludeSubfolders
            $retrieveSizes = $profileData.RetrieveSizes
            $recursionDepth = $profileData.RecursionDepth
            $folderFilter = $profileData.FolderFilter
            $csvFilePath = $profileData.CSVFilePath

            # Update the form controls
            $textBox.Text = $pathPrompt
            $checkBox.Checked = $includeSubfolders
            $chkSize.Checked = $retrieveSizes
            $txtDepth.Text = $recursionDepth
            $txtFilter.Text = $folderFilter
            $txtCsvPath.Text = $csvFilePath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error loading profile: $_", 'Load Profile', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

$okButton = New-Object Windows.Forms.Button
$okButton.Location = New-Object Drawing.Point(250, 190)
$okButton.Size = New-Object Drawing.Size(75, 23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button
$cancelButton.Location = New-Object Drawing.Point(150, 190)
$cancelButton.Size = New-Object Drawing.Size(75, 23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.Controls.Add($cancelButton)

$label = New-Object Windows.Forms.Label
$label.Location = New-Object Drawing.Point(10, 15)
$label.Size = New-Object Drawing.Size(280, 20)
$label.Text = 'Please select a directory:'
$form.Controls.Add($label)

$textBox = New-Object Windows.Forms.TextBox
$textBox.Location = New-Object Drawing.Point(10, 40)
$textBox.Size = New-Object Drawing.Size(260, 20)
$form.Controls.Add($textBox)

$browseButton = New-Object Windows.Forms.Button
$browseButton.Location = New-Object Drawing.Point(280, 38)
$browseButton.Size = New-Object Drawing.Size(75, 23)
$browseButton.Text = 'Browse'
$browseButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBox.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($browseButton)

$checkBox = New-Object Windows.Forms.CheckBox
$checkBox.Location = New-Object Drawing.Point(10, 190)
$checkBox.Size = New-Object Drawing.Size(300, 20)
$checkBox.Text = 'Include Sub Folders?'
$form.Controls.Add($checkBox)
# Disable Depth text box if the 'Include Sub Folders' is checked
$checkBox.Add_CheckStateChanged({
    if ($checkBox.Checked) {
        $txtDepth.Enabled = $false
    } else {
        $txtDepth.Enabled = $true
    }
})
$form.Topmost = $true

$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $pathPrompt = $textBox.Text
    $includeSubfolders = $checkBox.Checked
    $retrieveSizes = $chkSize.Checked
    $recursionDepth = [int]$txtDepth.Text
    $folderFilter = $txtFilter.Text
    $csvFilePath = $txtCsvPath.Text  # Get the CSV file path from the input field

    $params = @{
        'FolderPath' = $pathPrompt
    }

    if ($includeSubfolders) {
        $params['Recurse'] = $true
    }

    if ($recursionDepth -gt 0) {
        $params['Depth'] = $recursionDepth
    }

    if ($folderFilter) {
        $params['Filter'] = $folderFilter
    }

    $Output = Get-FolderDetails @params

    if (-not $retrieveSizes) {
        $Output = $Output | Select-Object -Property * -ExcludeProperty 'Size(MB)'
    }

    if ($includeSubfolders) {
        $Output | Export-Csv -Path $csvFilePath -NoTypeInformation
    }

    $Output | Out-GridView -Title Zahavi-FS -PassThru
}
