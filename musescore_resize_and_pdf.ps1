# Powershell script
# Applies a style file to a musecore mscz file then converts it to a pdf.
# SETUP: double check and amend your path to the musescore executable in the commands below
# See https://musescore.org/en/handbook/3/command-line-options for help with musescore cli usage
# The original mscz file is modified with the new style as well as the output pdf
# Style files: https://musescore.org/en/handbook/3/layout-and-formatting#save-and-load-style
# Scaling: https://musescore.org/en/handbook/3/page-settings#scaling

param(
        [Parameter(Mandatory, HelpMessage="Path to musescore file")]
        [string]$msczFile,

        [Parameter(HelpMessage="Output file path")]
        [string]$outputPdf="$((Get-Item $msczFile).Basename).pdf",

        [Parameter(HelpMessage="Set the score's page width")]
        [string]$pageWidth=8.5,
        [Parameter(HelpMessage="Set the score's page height")]
        [string]$pageHeight=11,
    

        [Parameter(HelpMessage="Set the score's page scaling space")]
        [string]$Spatium=2.23
    )
# SETUP: double check and amend your path to the musescore executable in the commands below
    $path="C:\Program Files\MuseScore 3\bin\Musescore3.exe" 


# Save style definition to a temp file.
# To see available style settings, from musescore, save a style as a .mss file and inspect the contents
$styleFile = New-TemporaryFile
Set-Content -Path $styleFile.FullName -Value @"
<?xml version="1.0" encoding="UTF-8"?>
<museScore version="3.02">
<Style>
    <pageWidth>${pageWidth}</pageWidth>
    <pageHeight>${pageHeight}</pageHeight>
    <pageEvenLeftMargin>0.393701</pageEvenLeftMargin>
    <pageOddLeftMargin>0.390157</pageOddLeftMargin>
    <pageEvenTopMargin>0.390157</pageEvenTopMargin>
    <pageEvenBottomMargin>0.790157</pageEvenBottomMargin>
    <pageOddTopMargin>0.390157</pageOddTopMargin>
    <pageOddBottomMargin>0.790157</pageOddBottomMargin>
    
    <Spatium>${Spatium}</Spatium>
</Style>
</museScore>
"@

# Apply the style to the score
# Pipe to out-null because we want to wait for Musescore to finish update the mscz file before converting it to pdf
Write-Host "Updating score style..."
& $path -S $styleFile.FullName -o $msczFile $msczFile | Out-Null

# Convert the score to pdf
Write-Host "Converting to pdf..."
& $path -o $outputPdf $msczFile

# Clean up temp files
Remove-Item -Path $styleFile.FullName
Write-Host "Done"

