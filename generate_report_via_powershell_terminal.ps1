# Requires ImportExcel - Open Powershell as Administrator:
# Install-Module -Name ImportExcel -Scope CurrentUser

# Might need execution policy permissions. To allow local scripts and signed remote scripts to run:
# Get-ExecutionPolicy
# Set-ExecutionPolicy RemoteSigned

# Ask user for input and output folders
$inputFolder = Read-Host "Enter the path to the folder containing XML files"
$outputFolder = Read-Host "Enter the path to the folder where Excel report will be saved"

# Validate folders
if (-not (Test-Path $inputFolder)) {
    Write-Host "Input folder does not exist." -ForegroundColor Red
    exit
}
if (-not (Test-Path $outputFolder)) {
    Write-Host "Output folder does not exist. Creating it..."
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# Load XML files
$xmlFiles = Get-ChildItem -Path $inputFolder -Filter "*.xml"

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
    $worksheet.Cells.Item(2, $i+1).NumberFormat = "@"
    $worksheet.Cells.Item(2, $i+1).Value2 = $headers[$i]
}

$row = 3

# Get the values from the xml files
foreach ($file in $xmlFiles) {
    [xml]$xml = Get-Content $file.FullName

    # Custom 1
    $customNode = $xml.SelectSingleNode("//user_data/entries[@name='fields']/entry[@name='Custom 1']")
    $customValue = if ($customNode) { $customNode.InnerText } else { "" }

    # System version
    $systemNode = $xml.SelectSingleNode("//entries[@name='system']")
    $version = $systemNode.SelectSingleNode("entry[@name='version']").InnerText

    # First SATA/SSD disk (skip USB)
    $diskNode = $xml.SelectNodes("//entries[@name='disk']") |
                Where-Object { $_.SelectSingleNode("entry[@name='interface_type']").InnerText -ne "USB" } |
                Select-Object -First 1
    $diskSerial = $diskNode.SelectSingleNode("entry[@name='serial']").InnerText

    # Blancco status
    $erasureNode = $xml.SelectSingleNode("//blancco_erasure_report/entries[@name='erasures']/entries[@name='erasure']")
    $state = $erasureNode.SelectSingleNode("entry[@name='state']").InnerText
    $blanccoStatus = if ($state -eq "Successful") { "Green" } else { "Red" }

    # Force cells to text
    for ($col=1; $col -le $headers.Count; $col++) {
        $worksheet.Cells.Item($row,$col).NumberFormat = "@"
    }

    # Fill row
    $worksheet.Cells.Item($row,1).Value2 = $version
    $worksheet.Cells.Item($row,2).Value2 = $customValue
    $worksheet.Cells.Item($row,3).Value2 = "" # Blancco status
    $worksheet.Cells.Item($row,4).Value2 = $diskSerial

    # Apply color
    $cell = $worksheet.Cells.Item($row,3)
    $cell.Interior.ColorIndex = if ($blanccoStatus -eq "Green") {4} else {3}

    $row++
}

# Auto-fit columns
$worksheet.Columns.AutoFit()

# Save with timestamp in output folder
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$savePath = Join-Path $outputFolder "DataDestructionReport_$timestamp.xlsx"
$workbook.SaveAs($savePath)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Report saved to $savePath" -ForegroundColor Green