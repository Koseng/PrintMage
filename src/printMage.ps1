param (
    $pathWorkDir = "C:\printmage",
    $pdfFullPath = "C:\printmage\test.pdf",
    $pdfFileBaseName = "test"
)

try {

Set-Location $pathWorkDir

$logFile = "log.txt"
if([System.IO.File]::Exists($logFile)){ Remove-Item $logFile }

Function LogWrite {
  Param ([string]$logstring)
  $dateString = [System.DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")
  Write-Host "$dateString $logstring"
  Add-content $logFile -value "$dateString $logstring"
}

LogWrite "#### -------- INPUT $pathWorkDir # $pdfFullPath # $pdfFileBaseName --------------"

# READ CONFIGURATION
# --------------------------
[XML]$conf                   = Get-Content "configuration.xml"
[string]$pathPDF24           = $conf.Config.PathPDF24
[string]$defaultPrinterName  = """$($conf.Config.DefaultPrinterName)"""
LogWrite "#### CONFIG $pathPDF24 # $defaultPrinterName # $($conf.Config.ConditionOnlyOnFirstPage) # $($conf.Config.DeleteOriginalPdf) # $($conf.Config.DeleteTemporaryFiles)"

$atLeastOneConditionMatched = $false;
$pdfQuoteFullPath       = """$pdfFullPath"""
$docToolQuoteFullPath   = """$pathPDF24\pdf24-DocTool.exe"""
$readerQuoteFullPath    = """$pathPDF24\pdf24-Reader.exe"""
$textFileFullPath       = "$pathWorkDir\$pdfFileBaseName" + ".txt"
$tempPdfFullPath        = "$pathWorkDir\$pdfFileBaseName" + "_TEMP.pdf"
$tempPdfQuoteFullPath   = """$tempPdfFullPath"""
$textFileQuoteFullPath  = """$textFileFullPath"""

# EXTRACT TEXT FROM PDF
# --------------------------
if ([int32]$conf.Config.ConditionOnlyOnFirstPage) { $pagesCommand = "-l 1" } else { $pagesCommand = "-f 1" } 
LogWrite "#### START text extraction from $pdfQuoteFullPath to $textFileQuoteFullPath"
Start-Process "pdftotext.exe" -Wait -ArgumentList "-raw", "-nopgbrk", $pagesCommand, $pdfQuoteFullPath, $textFileQuoteFullPath
$extractedText = [System.IO.File]::ReadAllText($textFileFullPath)

# PROCESS ALL PRINT CONDITIONS
# ------------------------------
foreach($cond in $conf.Config.PrintConditions.PrintCondition) 
{
    $isIncluded = $true
    $isExcluded = $true
    if ([string]$cond.IncludeRegEx) { $isIncluded = $extractedText -match    $cond.IncludeRegEx }
    if ([string]$cond.ExcludeRegEx) { $isExcluded = $extractedText -notmatch $cond.ExcludeRegEx }

    LogWrite "#### [$($cond.Name)] IncludeRegEx=$isIncluded ExcludRegEx=$isExcluded"

    # GENERATE DOCUMENT AND PRINT IT
    # ------------------------------
    if ($isIncluded -and $isExcluded)
    {
        $atLeastOneConditionMatched = $true; 
        $printerName = """$($cond.PrinterName)"""     
        $joinString = ""
        for ($i=1; $i -le $cond.Copies; $i++) {$joinstring += "$pdfQuoteFullPath "}

        LogWrite "#### START document generation for [$($cond.Name)] to $tempPdfQuoteFullPath"
        Start-Process -FilePath $docToolQuoteFullPath -Wait -ArgumentList "-noProgress", "-join", "-profile default/good", "-outputFile $tempPdfQuoteFullPath", $joinString
        LogWrite "#### START printing $tempPdfQuoteFullPath for [$($cond.Name)]"
        Start-Process -FilePath $readerQuoteFullPath  -Wait -WindowStyle Minimized -ArgumentList "/printTo $printerName", $tempPdfQuoteFullPath 
    }
} # foreach

# PRINT WITH DEFAULT PRINTER
#--------------------------------
if (!$atLeastOneConditionMatched)
{
    LogWrite "#### START printing $pdfQuoteFullPath with default printer"
    Start-Process -FilePath $readerQuoteFullPath  -Wait -WindowStyle Minimized -ArgumentList "/printTo $defaultPrinterName", $pdfQuoteFullPath
}

# DELETE FILES
#--------------------
if ([int32]$conf.Config.DeleteTemporaryFiles) 
{ 
    if([System.IO.File]::Exists($textFileFullPath)){ Remove-Item $textFileFullPath }
    if([System.IO.File]::Exists($tempPdfFullPath)){ Remove-Item $tempPdfFullPath } 
}
   
if ([int32]$conf.Config.DeleteOriginalPdf) 
{ 
    if([System.IO.File]::Exists($pdfFullPath)){ Remove-Item $pdfFullPath }
}

} # try
catch
{
    LogWrite "#### Error: $($_.Exception.Message)"
}
finally
{
    LogWrite "#### Script ended"
    exit 0
}



