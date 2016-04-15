#requires -Version 5.0
<#
	.SYNOPSIS
		Downloads Softpaqs from HP Inc for each operating system and model specified.
	
	.DESCRIPTION
		Downloads Softpaqs from HP Inc for each operating system and model specified.
		
		+Requries Powershell 5.0
		+Requries PSFTP module (https://gallery.technet.microsoft.com/scriptcenter/PowerShell-FTP-Client-db6fe0cb)
	
	.PARAMETER cfgFile
		JSON configuration file
	
	.PARAMETER outDir
		Folder to place the downloaded softpaqs. There will be duplicates. DeDupe strongly encouraged.
	
	.EXAMPLE
				PS C:\> HpDriverDownloader.ps1 -cfgFile 'Value1' -outDir 'Value2'
	
	.NOTES
		Additional information about the file.
#>
[CmdletBinding()]
param
(
	[Parameter(ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[String]$cfgFile = "$PSScriptRoot\config.xml",
	[Parameter(ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[String]$outDir = "$PSScriptRoot\Softpaqs"
)

Import-Module PSFTP

$processedFile = "$PSScriptRoot\AlreadyProcessed.txt"
$procModelFile = "$PSScriptRoot\ProcessedModels.txt"

$secpasswd = ConvertTo-SecureString "password" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("anonymous", $secpasswd)

function main {
	
	if (!(Test-Path "$outDir\ProductCatalog\modelcatalog.xml")) {
		DownloadCatalog -DownloadDir $outDir
		#TEST CATALOG VERSION HERE
	}
	
	if (!(Test-Path $cfgFile)) {
		BuildConfig -Catalog "$outDir\ProductCatalog\modelcatalog.xml" -OutXML $cfgFile
		Write-Output "You must enable Language, Model and OS in the config file for this script to do anything.`n""$cfgFile"""
	}
	
	# Load configuration
	$config = ([xml](gc "$PSScriptRoot\config.xml")).config
	
	# Load modelcatalog.xml into a hastable^2
	$modelCatalog = Get-ModelCatalog -CatalogDir "$outDir\ProductCatalog"
	
	# Process the catalog against our configuration to build a processing queue
	$queue = ProcessCatalog -ModelCatalog $modelCatalog -Config $config -CatalogDir "$outDir\ProductCatalog"
	
	# Download the requisite Softpaqs
	DownloadandExtractSoftpaqs -Queue ([ref]$queue) -OutDir $outDir
	
	$queue | export-csv c:\Downloads\queue.csv -Force -NoTypeInformation
	
	
	
	<#
	#load softpaq that have already been processed
	if (Test-Path $processedFile) {
		$processedSp = gc $processedFile
	}
	
	if (Test-Path $procModelFile) {
		$processedModels = gc $procModelFile
	}
	#>
	#Set config items
	#$myOs = $config.operatingsystems.os | where { $_.enabled -eq 'true' }
	#$myLang = $config.languages.language | where { $_.enabled -eq 'true' }
	#$myModel = $config.models.model| where { $_.enabled -eq 'true' }
	
	# Import catalog
	############################DownloadCatalog ("$outDir")
	<#
	$data = @{ }
	foreach ($model in ($config.models.model | where { $_.enabled -eq 'true' })) {
		#$data.Add('ModelId', $model.Id)
		foreach ($l in (([xml](gc "$outDir\ProductCatalog\$($model.Id).xml")).ProductCatalog.ProductModel)) {
			$
		}
		
		
	}
	Write-Host "end"
	#>
	#Load chosen options
	#[array]$selectOS = $catalog.OperatingSystem | where {$myOs.name -contains $_.Name}
	#[array]$selectLang = $catalog.Language | where { $myLang.name -contains $_.Name }
	#[array]$selectModels = $catalog.ProductLine.ProductFamily.ProductModel | where { $myLang.name -contains $_.Na }
	
	<#
	foreach ($model in $catalog.ProductLine.ProductFamily.ProductModel){
		if ( ($myModel.name -contains $model.Name) -and ($processedModels -notcontains $model.Name) ) {
			$currentModels = (($myModel.name | ? {$_.HP -eq $model.Name}).altname)
			write-verbose "Checking model ""$($model.Name)"""

			[xml]$modelXml = gc "$outDir\ProductCatalog\$($model.Id).xml"
			# For each OS
			foreach ($o in $modelXml.NewDataSet.ProductCatalog.ProductModel.OS) {
				write-debug "`tOS $($o.Id) > $($selectOS.Id)"
				if ($selectOS.Id -contains $o.Id){
					
					# For each Language				
					foreach ($l in $o.Lang){
						write-debug "`t`t Lang $($selectLang.Id) > $($l.id)"
						if ($selectLang.Id -contains $l.id){
							
							#For each service paq  
							foreach ( $sp in ($l.SP | ? { ($_.S -eq '0') }) ){
                                Write-Debug "`t`t`t SP: $($sp.Id)"

								foreach ($ssp in ($catalog.Softpaq | ? { ($_.Id -eq $sp.Id) -and ($_.Category -like "Driver*") })){
                                    Write-Debug "`t`t`t`t SSP: $($ssp.Id)"
									$spFile = "$outDir\softpaq\sp$($sp.id).exe"
									$spDst = "$outDir\softpaq\$($sp.id)"
									

									# Did we already download this?
									if ( !(Test-Path -PathType Container -Path $spDst) ) {
										
										FtpDownload -RemoteFile $ssp.URL -Destination "$outDir\Softpaq" -Credentials $mycreds
										ExtractSoftpaq -SoftpaqPath $spFile -Destination $spDst
										if (Test-Path $spDst){
											Remove-Item $spFile -Force
										}
									} else {
										#already downloaded
									}
									
									#Create Path and copy drivers
									foreach ($sos in ($selectOS | ? {$_.id -eq $o.id})){
										foreach ($m in $currentModels){
											$drvPath = "$outDir\$($sos.Name)\$m\$($ssp.Name)"
											if ( !(Test-Path $drvPath) ){
												New-Item -Path $drvPath -ItemType directory -ErrorAction Continue
											
												Copy-Item "$spDst\*" $drvPath -Recurse -ErrorAction Continue
                                                write-verbose "Copying ""$spDst"" to ""$drvPath"""
											}
										}
									}
                                    
                                    $ssp.id | Out-File $processedFile -Append
                                    #$processedSp += $ssp.id
                                    
								}#Select SP
                                    
							} #SP
						}
					} #Lang
				}
			} #OS
		    $model.Name | out-file $processedModels -Append
        } 
	} #model
	#>
} #main

function DownloadandExtractSoftpaqs {
	param
	(
		[Parameter(Mandatory = $true, Position = 1)]
		[ref]$Queue,
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$OutDir
	)
	
	foreach ($sp in ($Queue.Value | select -Unique SoftpaqId, SoftpaqUrl, SoftpaqName)) {
		$spFile = "$OutDir\Softpaq\sp$($sp.SoftpaqId).exe"
		$spFldr = "$OutDir\Softpaq\sp$($sp.SoftpaqId)"
		
		#$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { Write-Verbose "$($_.SoftpaqId)"}
		
		# Download the SP if not already present
		if (Test-Path -PathType Container -Path $spFldr) {
			$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Extracted' }		
		} elseif (Test-Path -Path $spFile) {
			if (ExtractSoftpaq -SoftpaqPath $spFile -Destination $spFldr) {
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Extracted' }
			} else {
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Error: Cannot extract' }
			}
		} else {

			if (FtpDownload -RemoteFile $sp.SoftpaqUrl -Destination "$outDir\Softpaq" -Credentials $mycreds) {
				#$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Downloaded' }
				if (ExtractSoftpaq -SoftpaqPath $spFile -Destination $spFldr) {
					$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Extracted' }
					Remove-Item $spFile -Force
				} else {
					$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Error: Cannot extract' }
				}
			} else {
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Download Error' }
			}
			
			
			
		}
		
	} #foreach
}

function DownloadCatalog {
	param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$DownloadDir
	)
	
	Write-Verbose "Downloading catalog to $DownloadDir"
	$ProductCatalogUpdateUrl = 'ftp://ftp.hp.com/pub/caps-softpaq/ProductCatalogUpdate.xml'
	$ProductCatalogUrl = 'ftp://ftp.hp.com/pub/caps-softpaq/ProductCatalog.zip'
	
	FtpDownload -Credentials $mycreds -Destination $DownloadDir -RemoteFile $ProductCatalogUpdateUrl
	FtpDownload -Credentials $mycreds -Destination $DownloadDir -RemoteFile $ProductCatalogUrl
	
	Write-Verbose "Extracting catalog to $DownloadDir"
	Expand-Archive "$DownloadDir\ProductCatalog.zip"  $DownloadDir -Force
	
	Remove-Item "$DownloadDir\ProductCatalog.zip"
	
}


function ConnectFtp {
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$SessionName,
		[Parameter(Mandatory = $true, Position = 1)]
		[ref]$Session,
		[Parameter(Mandatory = $true, Position = 2)]
		$Credentials,
		[Parameter(Mandatory = $true, Position = 3)]
		$Site
	)
	$Session.value = Get-FTPConnection -Session MyTestSession
	if ($Session.Value -eq $null) {
		$Session.value = Set-FTPConnection -Credentials $Credentials -Server $ftpServer -Session $SessionName -UsePassive -ErrorAction Stop
	}
}

function FtpDownload {
	Param (
		
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$RemoteFile,
		[Parameter(Mandatory = $true, Position = 1)]
		[string]$Destination,
		[Parameter(Mandatory = $true, Position = 2)]
		$Credentials
	)
	
	$ftpServer = ($RemoteFile -split ('/'))[2]
	$ftpPath = $RemoteFile -replace ("ftp://$ftpServer", "")
	#write-verbose "$ftpServer $ftpPath"
	
	Try {
		#Create folder to hold CVA files
		if (!(Test-Path $Destination)) {
			New-Item -ItemType directory -Path $Destination -ErrorAction Stop
		}
		
		#Connect to FTP server
		$Session = Get-FTPConnection -Session 'HP'
		if ($Session -eq $null) {
			$Session = Set-FTPConnection -Credentials $Credentials -Server $ftpServer -Session 'HP' -UsePassive -ErrorAction Stop
		}
		
		#Download File
		Get-FTPItem -Session $session -LocalPath $Destination -ErrorAction Stop -Path $ftpPath -Overwrite
		
		
	} Catch [System.IO.IOException] {
		Write-Error "Unable to write file"
	} Catch {
		Write-Error $_.Exception
	}
}

Function ExtractSoftpaq {
	param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$SoftpaqPath,
		[Parameter(Mandatory = $true, Position = 1)]
		[string]$Destination
	)
	Write-Verbose "Extracting: ""$SoftpaqPath"""
	$args = @('-e', '-s', '-f', $Destination)
	$ret = start-process $SoftpaqPath -ArgumentList $args -Wait -NoNewWindow
	return ($ret.ExitCode)
}

<#
	.SYNOPSIS
		Creates a default config based on current HP ProductCatalog.xml
#>
function BuildConfig {
	param
	(
		[Parameter(Mandatory = $true, Position = 0)]
		$Catalog,
		[Parameter(Mandatory = $true, Position = 1)]
		$OutXML
	)
	
	$template = @"
<?xml version="1.0"?>
<config>
	<languages>
	$(foreach ($l in ($catalog.Language)) { "`t`t<language id=""$($l.Id)"" name=""$(Escape($l.Name))"" enabled=""false"" />`n" })
	</languages>
	<operatingsystems>
	$(foreach ($o in ($catalog.OperatingSystem)) { "`t`t<os id=""$($o.Id)"" name=""$(Escape($o.Name))"" enabled=""false"" />`n" })
	</operatingsystems>
	<models>
	$(foreach ($m in ($catalog.ProductLine.ProductFamily.ProductModel | where { $_ -ne $null })) {
		"`t`t<model id=""$(Escape($m.Id))"" name=""$(Escape($m.Name))"" altname="""" enabled=""false"" />`n"
	})
	</models>
	<categories>
	$(foreach ($c in ($catalog.Softpaq.Category | select -Unique | sort)) { "`t`t<category name=""$(Escape($c))"" enabled=""false"" />`n" })
	</categories>
</config>
"@
	
	$template | Out-File $OutXML -force
}

function Get-ModelCatalog {
	[OutputType([hashtable])]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[string]$CatalogDir
	)
	
	$modelCatalogXml = ([xml](gc "$CatalogDir\modelcatalog.xml")).NewDataSet.ProductCatalog
	$modelCatalog = @{ }
	
	Write-Verbose "Processing Softpaqs in ModelCatlog.xml"
	$modelCatalog.Add('Softpaqs', (New-XmlNodeHash -XmlNode ($modelCatalogXml.Softpaq) -Key 'Id'))
	
	Write-Verbose "Processing Languages in ModelCatlog.xml"
	$modelCatalog.Add('Languages', (New-XmlNodeHash -XmlNode ($modelCatalogXml.Language) -Key 'Id'))
	
	Write-Verbose "Processing Models in ModelCatlog.xml"
	$modelCatalog.Add('Models', (New-XmlNodeHash -XmlNode ($modelCatalogXml.ProductLine.ProductFamily.ProductModel) -Key 'Id'))
	
	Write-Verbose "Processing Operating Systems in ModelCatlog.xml"
	$modelCatalog.Add('OperatingSystems', (New-XmlNodeHash -XmlNode ($modelCatalogXml.OperatingSystem) -Key 'Id'))
	
	return $modelCatalog
}

function New-XmlNodeHash {
	param
	(
		[Parameter(Mandatory = $true, Position = 1)]
		[array]$XmlNode,
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$Key
	)
	$outHash = @{ }
	
	foreach ($sub in $xmlNode) {
		try {
			$outHash.Add(($sub.$Key), $sub)
		} catch {
			#Write-Warning "Duplicate key $Key for $XmlNode"
		}
	}
	
	return $outHash
}

<#
	.SYNOPSIS
		A brief description of the ProcessCatalog function.
	
	.DESCRIPTION
		A detailed description of the ProcessCatalog function.
	
	.PARAMETER Config
		Configuration data in XML
	
	.PARAMETER ModelCatalog
		Hastable contain modelcatalog information
	
	.NOTES
		Additional information about the function.
#>
function ProcessCatalog {
	param
	(
		[Parameter(Mandatory = $true)]
		[System.Xml.XmlElement]$Config,
		[Parameter(Mandatory = $true)]
		[hashtable]$ModelCatalog,
		[Parameter(Mandatory = $true)]
		[String]$CatalogDir
	)
	
	$processedData = @()
	
	$cfgOs = $Config.operatingsystems.os | where { $_.enabled -eq 'true' }
	$cfgLng = $Config.languages.language | where { $_.enabled -eq 'true' }
	$cfgCat = ($Config.categories.category | where { $_.enabled -eq 'true' })
	$cfgModel = $Config.models.model | where { $_.enabled -eq 'true' }
	
	foreach ($file in (gci $catalogDir | where { $_.Name -ne 'modelcatalog.xml' } | where { $cfgModel.Id -contains $_.BaseName })) {
		Write-Verbose "Processing catalog file $($file.Name)"
		
		$sps = ((([xml](gc $file.FullName)).NewDataSet.ProductCatalog.ProductModel.OS |
		where { $cfgOs.id -contains $_.Id }).Lang |
		where { $cfgLng.id -contains $_.Id }).SP |
		where { $_.S -eq '0' }
		
		foreach ($sp in $SPs) {
			if ( ($cfgCat.name) -contains (($ModelCatalog.Softpaqs)[$sp.Id].Category) ) {
				$processedData += New-Object -TypeName PSObject -Property ([Ordered]@{
					'LangId' = $sp.ParentNode.Id;
					'LangName' = (($ModelCatalog.Languages)[$sp.ParentNode.Id]).Name
					'OsId' = $sp.ParentNode.ParentNode.Id
					'OsName' = ($ModelCatalog.OperatingSystems)[$sp.ParentNode.ParentNode.Id].Name
					'OsShortName' = ($ModelCatalog.OperatingSystems)[$sp.ParentNode.ParentNode.Id].ssmname
					'SoftpaqId' = $sp.Id
					'SoftpaqName' = ($ModelCatalog.Softpaqs)[$sp.Id].Name
					'SoftpaqUrl' = ($ModelCatalog.Softpaqs)[$sp.Id].Url
					'SoftpaqCategory' = ($ModelCatalog.Softpaqs)[$sp.Id].Category
					'SoftpaqVersion' = ($ModelCatalog.Softpaqs)[$sp.Id].Version
					'ModelId' = $sp.ParentNode.ParentNode.ParentNode.Id
					'ModelName' = ($ModelCatalog.Models)[$sp.ParentNode.ParentNode.ParentNode.Id].Name
					'ModelCustomName' = ($ModelCatalog.Models)[$sp.ParentNode.ParentNode.ParentNode.Id].altname
					'FileStatus' = ''
					'Status' = ''
				});
			}
		}
	}
	
	$processedData
}

main