# https://github.com/EvotecIT/PSWriteWord/tree/master
# https://perfectdoc.studio/inspiration/generate-word-documents-from-excel-data/
# https://www.youtube.com/watch?app=desktop&v=yO8vszeccQE&ab_channel=LearnAccessByCrystal

# Import required modules

Import-Module ImportExcel

Import-Module PSWriteWord

# Load Excel data

$excelData = Import-Excel -Path 'C:\Users\Shadow\source\Report Generation PS Script\test_data.xlsx' -WorksheetName 'Sheet1'

# Specify the output directory

$outputDirectory = 'C:\Users\Shadow\source\Report Generation PS Script'

# Create the output directory if it doesn't exist

if (-not (Test-Path -Path $outputDirectory -PathType Container)) {

    New-Item -Path $outputDirectory -ItemType Directory

}

# Loop through each customer record in Excel data

foreach ($customer in $excelData) {

    # Create a new Word document for each customer

    $doc = New-WordDocument

    # Add customer information to the document

    Add-WordText -WordDocument $doc -Text "Customer Name: $($customer.Name)"

    Add-WordText -WordDocument $doc -Text "Email: $($customer.Email)"

    Add-WordText -WordDocument $doc -Text "Address: $($customer.Address)"

    # Generate a unique file name for each document

    $fileName = Join-Path -Path $outputDirectory -ChildPath "$($customer.Name)_Report.docx"

   # Save and close the Word document

    Save-WordDocument $doc -Path $fileName

    # $doc.Close()

}

Write-Host "Generated $($excelData.Count) customer reports in $outputDirectory."