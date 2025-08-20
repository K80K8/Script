# Requires ImportExcel - Open Powershell as Administrator:
# Install-Module -Name ImportExcel -Scope CurrentUser

# To allow local scripts and signed remote scripts to run in Powershell:
# Get-ExecutionPolicy
# Set-ExecutionPolicy RemoteSigned

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Data Destruction Report Generator"
$form.Size = New-Object System.Drawing.Size(500,170)
# $form.BackColor = [System.Drawing.Color]::FromArgb(255, 87, 210)
$form.StartPosition = "CenterScreen"

# Input folder label and textbox
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Input Folder:"
$lblInput.Location = New-Object System.Drawing.Point(10,20)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(100,15)
$txtInput.Size = New-Object System.Drawing.Size(280,20)
$form.Controls.Add($txtInput)

$btnInput = New-Object System.Windows.Forms.Button
$btnInput.Text = "Browse"
$btnInput.Location = New-Object System.Drawing.Point(390,13)
# $btnInput.BackColor = [System.Drawing.Color]::FromArgb(216,145,239)
$btnInput.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select the folder containing XML files"
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtInput.Text = $folderDialog.SelectedPath
    }
})
$form.Controls.Add($btnInput)

# Generate button
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Report"
$btnGenerate.Location = New-Object System.Drawing.Point(180,70)
$btnGenerate.Size = New-Object System.Drawing.Size(120,40)
$btnGenerate.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
# $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(216,145,239)
$btnGenerate.Add_Click({

    $inputFolder = $txtInput.Text

    if (-not (Test-Path $inputFolder)) {
        [System.Windows.Forms.MessageBox]::Show("Please select a valid input folder.","Error","OK",[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $xmlFiles = Get-ChildItem -Path $inputFolder -Filter "*.xml"
    if ($xmlFiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No XML files found in the selected input folder.","Error","OK",[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Ask user where to save file
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Title = "Save Report As"
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $saveDialog.FileName = "DataDestructionReport_$timestamp.xlsx"  # Suggested default
    $saveDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"

    if ($saveDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    $savePath = $saveDialog.FileName

    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Sheets.Item(1)

    # Title row
    $worksheet.Range("A1:Q1").Merge()
    $worksheet.Range("A1").Value2 = "Sam's Spares Data Destruction Report"
    $worksheet.Range("A1").Font.Size = 42

    # Headers
    $headers = @("Model", "Asset ID", "Blancco Wiped / Destroyed Drive","Disk Serial")
    for ($i=0; $i -lt $headers.Count; $i++) {
        $worksheet.Cells.Item(2, $i+1).NumberFormat = "@"  # Force cells to be text format
        $worksheet.Cells.Item(2, $i+1).Value2 = $headers[$i]
    }

    $row = 3

    # Iterate through xml files get relevant data and save to Excel file
    foreach ($file in $xmlFiles) {
        [xml]$xml = Get-Content $file.FullName

        $customNode = $xml.SelectSingleNode("//user_data/entries[@name='fields']/entry[@name='Custom 1']")
        $customValue = if ($customNode) { $customNode.InnerText } else { "" }

        $systemNode = $xml.SelectSingleNode("//entries[@name='system']")
        $version = $systemNode.SelectSingleNode("entry[@name='version']").InnerText

        $diskNode = $xml.SelectNodes("//entries[@name='disk']") |
                    Where-Object { $_.SelectSingleNode("entry[@name='interface_type']").InnerText -ne "USB" } |
                    Select-Object -First 1
        $diskSerial = $diskNode.SelectSingleNode("entry[@name='serial']").InnerText

        $erasureNode = $xml.SelectSingleNode("//blancco_erasure_report/entries[@name='erasures']/entries[@name='erasure']")
        $state = $erasureNode.SelectSingleNode("entry[@name='state']").InnerText
        $blanccoStatus = if ($state -eq "Successful") { "Green" } else { "Red" }

        for ($col=1; $col -le $headers.Count; $col++) {
            $worksheet.Cells.Item($row,$col).NumberFormat = "@"
        }

        $worksheet.Cells.Item($row,1).Value2 = $version
        $worksheet.Cells.Item($row,2).Value2 = $customValue
        $worksheet.Cells.Item($row,3).Value2 = "" 
        $worksheet.Cells.Item($row,4).Value2 = $diskSerial

        $cell = $worksheet.Cells.Item($row,3)
        $cell.Interior.ColorIndex = if ($blanccoStatus -eq "Green") {4} else {3}

        $row++
    }

    $worksheet.Columns.AutoFit()

    $workbook.SaveAs($savePath)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

    [System.Windows.Forms.MessageBox]::Show("Report saved to:`n$savePath","Success","OK",[System.Windows.Forms.MessageBoxIcon]::Information)
})

$form.Controls.Add($btnGenerate)

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()