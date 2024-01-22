param(
    [string] $inputNameString,
    [string] $templateFilePath,
    [string] $outputFolderPath,
    [string] $placeHolder,
    [switch] $help
)

# Check if required parameters are provided or if help is on
if ($help -or -not $inputNameString -or -not $templateFilePath){
    Write-Host "Usage: PowerShellScript.ps1 -inputNameString 'Name Surname' -templateFilePath C:\Path\To\Your\Template.docx -outputFolderPath C:\Path\To\Your\Folder\ -placeHolder 'placeholder'`nWhere:`n- inputNameString  is the name to replacete`n- templateFilePath is the path to a .docx file containing placeholder sequences.`n- outputFolderPath is the output location (default: same as template file)`n- placeholder      is the placeholder present in the template (default: '<<<ReplaceThis>>>')`n`nAdding -help to the command shows this text`n"
    Exit
}

if(-not $placeHolder){
    $placeHolder = "<<<ReplaceThis>>>"
}

if(-not $outputFolderPath){
    $outputFolderPath = $templateFilePath -replace '\\\w*\.\w*', "\"
}

# Set the output PDF path based on the template file path
$outputPdfPath = "${outputFolderPath}${inputNameString}_replaced.pdf"


# Create a new Word application
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false  # Set to $true if you want to see Word in action

# Open the Word template file
$document = $wordApp.Documents.Open($templateFilePath)

# Find and replace a placeholder in the Word document with the input string
$range = $document.Content
$range.Find.Execute($placeHolder, $false, $false, $false, $false, $false, $true, 1, $false, $inputNameString, 2)

# Save the modified document as a PDF
$document.SaveAs([ref]$outputPdfPath, [ref]17)  # 17 represents the PDF format

# Close the Word document and application
$document.Close($false)
$wordApp.Quit()

# Release the COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null

Write-Output "$outputPdfPath"
