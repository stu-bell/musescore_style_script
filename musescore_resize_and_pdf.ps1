# Powershell script
# Applies style settings to a musecore mscz file and optionally converts format (eg to a PDF).
# SETUP: double check and amend your path to the musescore executable in the musescorePath parameter
# The original mscz file is modified with the new style as well as the output file
# The style template in this script modifies just the styles I'm iterested in. 
# Modify the xml template and powershell parameters to add your own styles
# Style files (*.mss): https://musescore.org/en/handbook/3/layout-and-formatting#save-and-load-style
# See https://musescore.org/en/handbook/3/command-line-options for help with musescore cli usage

param(
        [Parameter(Mandatory, HelpMessage="Musescore input file path")]
        [string]$msczFile,

        [Parameter(HelpMessage="Output file path")]
        [string]$outputFile="$((Get-Item $msczFile).Basename).pdf",

        [Parameter(HelpMessage="Musescore Executable path")]
        [string]$musescorePath="C:\Program Files\MuseScore 3\bin\Musescore3.exe",

        # Page size parameters
        # Scaling: https://musescore.org/en/handbook/3/page-settings#scaling
        [Parameter(HelpMessage="Page scaling space")]
        [string]$Spatium=2.23,
        [Parameter(HelpMessage="Page width")]
        [string]$PageWidth=8.5,
        [Parameter(HelpMessage="Page height")]
        [string]$PageHeight=11,

        # Page margin params
        [Parameter(HelpMessage="Page Top and Bottom margins")]
        [string]$TopBottomMargin=0.35,
        # "Right Margin" = page width - left margin - printable width
        [Parameter(HelpMessage="Page Left and Right margins")]
        [string]$LeftRightMargin=0.35,
        [Parameter(HelpMessage="Page PrintableWidth")]
        [string]$PrintableWidth=($PageWidth - 2*$LeftRightMargin)
)

# Save style definition to a temp file.
$styleFile = New-TemporaryFile
# Modify the mss style xml template below to add your own styles
# To see available style settings, from musescore, save a style as a .mss file and inspect the contents
Set-Content -Path $styleFile.FullName -Value @"
<?xml version="1.0" encoding="UTF-8"?>
<museScore version="3.02">
<Style>
    <Spatium>${Spatium}</Spatium>
    <pageWidth>${PageWidth}</pageWidth>
    <pageHeight>${PageHeight}</pageHeight>
    <!-- 'Right margin' is determined by page width, printable width and left margins -->
    <pagePrintableWidth>${PrintableWidth}</pagePrintableWidth>
    <pageEvenLeftMargin>${LeftRightMargin}</pageEvenLeftMargin>
    <pageOddLeftMargin>${LeftRightMargin}</pageOddLeftMargin>
    <pageEvenTopMargin>${TopBottomMargin}</pageEvenTopMargin>
    <pageEvenBottomMargin>${TopBottomMargin}</pageEvenBottomMargin>
    <pageOddTopMargin>${TopBottomMargin}</pageOddTopMargin>
    <pageOddBottomMargin>${TopBottomMargin}</pageOddBottomMargin>
</Style>
</museScore>
"@

# Apply the style to the score
# Pipe to out-null because we want to wait for Musescore to finish update the mscz file before converting it to pdf
Write-Host "Updating score style..."
& $musescorePath -S $styleFile.FullName -o $msczFile $msczFile | Out-Null

# Convert the score to output format
Write-Host "Converting to output format..."
& $musescorePath -o $outputFile $msczFile | Out-Null

# Clean up temp files
Remove-Item -Path $styleFile.FullName 
Write-Host "Done"
