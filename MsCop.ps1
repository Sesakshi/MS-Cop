param (
  [string]$manuscriptRoot = "C:\Test",
  [string]$xslRoot = "C:\temp\MsCop\source\MsCop",
  [string]$outputRoot = "C:\temp",
  [switch]$asXml = $true, # Only apply a default to one of the "as" switch values...or leave it off entirely and pass it with the command
  [switch]$asHtml,
  [string]$MSCatalogPath = "C:\Manuscripts\ManuScriptCatalog.xml",
  [switch]$includeDCT,
  [switch]$help
)
# An array used to exclude manuscripts by name. Adjust this as-needed.
$excludelist =  @(
    "DCT_*","DuckCreekTech_*","Billing_*","BasicTransaction_*","BookPolicy_Rules_*","Cancel_Rules_*","CancelPending_Rules_*",
    "DCTPortfolio_*","DuckStoreTutorial*","Endorse_Rules_*","FinalAudit_Rules_*","FinalReport_Rules_*","Information_Rules_*",
    "InterimAudit_Rules_*","IssuePolicy_Rules_*","MidPoint_Rules_*","New_Rules_*","NonRenew_Rules_*","ProductStudio_*","Reinstate_Rules_*","Renew_Rules_*","RenewalChange_Rules_*",
    "Reporting_Rules_*","Rescind_Rules_*","RescindCancelPending_Rules_*","RescindNonRenew_Rules_*","Resume_Rules_*","Revise_Rules_*","RevisedFinalAudit_Rules_*",
    "RevisedFinalReport_Rules_*","RevisedInterimAudit_Rules_*","RevisedReporting_Rules_*","Rewrite_Rules_*","Suite_Events_*",
    "SuiteEvents_*","Suspend_Rules_*","VoidFinalAudit_Rules_*","VoidFinalReport_Rules_*","VoidInterimAudit_Rules*","VoidReporting_Rules*"
    )

if ($includeDCT -eq $true){
    Clear-Variable -Name "excludeList"
}

function review($manuscriptRoot, $xslFile, $outputRoot, $extension) {
  if (-not (Test-Path $outputRoot)) {
    md -Force $outputRoot
  }

  $xsltSettings = New-Object System.Xml.Xsl.XsltSettings
  $xsltSettings.EnableScript = 1
  $xsltSettings.EnableDocumentFunction = 1
  $xmlUrlResolver = New-Object System.Xml.XmlUrlResolver

  $transform = New-Object System.Xml.Xsl.XslCompiledTransform;
  $transform.Load($xslFile, $xsltSettings, $xmlUrlResolver)

  # Transform all the .xml files in the target folder, but only if they have a matching file name in the ManuScriptCatalog.xml file, and they aren't specifically excluded.
  Get-ChildItem $manuscriptRoot -Recurse -Filter *.xml -Exclude $excludelist | ForEach-Object {
    $name = $_.Name
    $xml = $_.FullName
    $basename = $_.BaseName
    if (Get-Content -Path $MSCatalogPath | Select-String -Pattern $name -SimpleMatch -Quiet){  
        $review = [System.IO.Path]::Combine($outputRoot, "$($basename).$($extension)")

        Write-Host "$xml -> $review"
        $transform.Transform($xml, $review)
    }
  }
}

function removeEmptyReviews($outputRoot, $xpath) {
  Get-ChildItem $outputRoot | Foreach-Object {
    $doc = [xml](Get-Content $_.FullName)

    if ($doc.DocumentElement -eq $null) {
      Remove-Item $_.FullName
    } else {
      $nodes = $doc.DocumentElement.SelectNodes($xpath)

      if ($nodes.Count -eq 0) {
        Remove-Item $_.FullName
      }
    }
  }
}

if (-not $help){
	if ($manuscriptRoot -eq $null) {
		Write-Host "Please specify ManuScript root path with -manuscriptRoot parameter. Use -help for more information."
	}

	if ((-not $asXml) -and (-not $asHtml)) {
	  Write-Host "Please specify output type.  Options: -asXml or -asHtml. Use -help for more information."
	  return
	}

	if ($MSCatalogPath -eq $null) {
		Write-Host "Please specify the path to the ManuScriptCatalog.xml file with the -MSCatalogPath parameter. Use -help for more information."
	}

	if ($asXml) {
	  $xslFile = [System.IO.Path]::Combine($xslRoot, "MsCop_XML.xsl")
	  $output = [System.IO.Path]::Combine($outputRoot, "XML")
	  review $manuscriptRoot $xslFile $output "xml"
	  removeEmptyReviews $output "modules/module/items/item"
	}

	if ($asHtml) {
	  $xslFile = [System.IO.Path]::Combine($xslRoot, "MsCop_HTML.xsl")
	  $output = [System.IO.Path]::Combine($outputRoot, "HTML")
	  review $manuscriptRoot $xslFile $output "html"
	  removeEmptyReviews $output "BODY/div/@id"
	} 
} else {
	Write-Host @"
===============================================================================
DESCRIPTION: Used to initiate MsCop, a ManuScript static code analysis tool. 
This will run ManuScript xml files through a series of transforms to attempt to
find problems. The script should be invoked from a PowerShell prompt. Defaults
can be set for parameters by modifying the MsCop.ps1 file.
-------------------------------------------------------------------------------
USAGE: .\MsCop.ps1 -manuscriptRoot "C:\DCT\ManuScripts" -xslRoot "C:\MsCop" -outputRoot "C:\Temp" -MSCatalogPath "C:\DCT\Shared\ManuScriptCatalog.xml" -asXml

-------------------------------------------------------------------------------
PARAMETERS:
-manuscriptRoot     Path to the ManuScript folder. We look recursively into all
                    sub-folders.
					
-xslRoot            Path to the MsCop folder

-outputRoot         Path to the folder to drop the results. We will create
                    either an XML, or HTML subfolder at this location.
					
-MSCatalogPath      Path to the environment ManuScriptCatalog.xml file

-asXML              Use this to indicate XML output

-asHTML             Use this to indicate HTML output

-includeDCT         Use this to generate reports against base Duck Creek
                    ManuScripts
					
-help               Description, usage and parameters
===============================================================================
"@
}
