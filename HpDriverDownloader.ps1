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
	[String]$cfgFile = "$PSScriptRoot\config.json",
	
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
	# Load config file
	$config = gc $cfgFile | ConvertFrom-Json
	
	#load softpaq that have already been processed
	if (Test-Path $processedFile)
	{
		$processedSp = gc $processedFile
	}
	
	if (Test-Path $procModelFile)
	{
		$processedModels = gc $procModelFile
	}
	
	#Set config items
	$myOs = $config.Os
	$myLang = $config.Language
	$myModel = $config.Model
	
	
	# Import catalog
	DownloadCatalog ("$outDir")
	$catalog = ([xml](gc "$outDir\ProductCatalog\modelcatalog.xml")).DocumentElement.ProductCatalog
	
	#TEST BUILD CATALOG
	BuildConfig -Catalog $catalog -OutXML "$PSScriptRoot\config.xml"
	
	#Load chosen options
	[array]$selectOS = $catalog.OperatingSystem | where {$myOs -contains $_.Name}
	[array]$selectLang = $catalog.Language | where {$myLang -contains $_.Name}
	
	foreach ($model in $catalog.ProductLine.ProductFamily.ProductModel){
		if ( ($myModel.HP -contains $model.Name) -and ($processedModels -notcontains $model.Name) ) {
			$currentModels = (($myModel | ? {$_.HP -eq $model.Name}).bios)
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

} #main

function DownloadCatalog {
	param (
		[Parameter(Mandatory=$true, Position=0)]
		[string] $DownloadDir
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
    	[Parameter(Mandatory=$true, Position=0)]
		[string] $SessionName,
		
		[Parameter(Mandatory=$true, Position=1)]
		[ref] $Session,

		[Parameter(Mandatory=$true, Position=2)]
		$Credentials,

        [Parameter(Mandatory=$true, Position=3)]
		$Site
    )
    $Session.value = Get-FTPConnection -Session MyTestSession
    if ($Session.Value -eq $null){
        $Session.value = Set-FTPConnection -Credentials $Credentials -Server $ftpServer -Session $SessionName -UsePassive -ErrorAction Stop 
    }
}

function FtpDownload {
	Param (	
	
		[Parameter(Mandatory=$true, Position=0)]
		[string] $RemoteFile,
		
		[Parameter(Mandatory=$true, Position=1)]
		[string] $Destination, 
		
		[Parameter(Mandatory=$true, Position=2)]
		$Credentials
	)
	
	$ftpServer = ($RemoteFile -split('/'))[2]
	$ftpPath = $RemoteFile -replace("ftp://$ftpServer", "")
	write-verbose "$ftpServer $ftpPath"
	
	Try {
		#Create folder to hold CVA files
		if ( !(Test-Path $Destination) ){
			New-Item -ItemType directory -Path $Destination -ErrorAction Stop
		}
		
		#Connect to FTP server
        $Session = Get-FTPConnection -Session 'HP'
        if ($Session -eq $null){
            $Session = Set-FTPConnection -Credentials $Credentials -Server $ftpServer -Session 'HP' -UsePassive -ErrorAction Stop 
        }
		
		#Download File
		Get-FTPItem -Session $session -LocalPath $Destination -ErrorAction Stop -Path $ftpPath -Overwrite

		
	} Catch [System.IO.IOException] {
		Write-Error "Unable to create CVA folder ""$CvaDir"""
	} Catch {
		Write-Output $_.Exception
	}
}

Function ExtractSoftpaq {
	param (
		[Parameter(Mandatory=$true, Position=0)]
		[string] $SoftpaqPath,
		
		[Parameter(Mandatory=$true, Position=1)]
		[string] $Destination
	)
	Write-Verbose "Extracting: ""$SoftpaqPath"" to ""$Destination"""
	$args = @('-e', '-s', '-f', $Destination)
	start-process $SoftpaqPath -ArgumentList $args -Wait -NoNewWindow
}

<#
	.SYNOPSIS
		Creates a default config based on current HP ProductCatalog.xml
#>
function BuildConfig
{
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
		$(foreach ($l in ($catalog.Language.Name)) { "`t`t<language name=""$l"" enabled=""false"" />`n" })
		</languages>
		<operatingsystems>
		$(foreach ($o in ($catalog.OperatingSystem.Name)) { "`t`t<os name=""$o"" enabled=""false"" />`n" })
		</operatingsystems>
		<models>
		$(foreach ($m in ($catalog.ProductLine.ProductFamily.ProductModel.Name | where { $_.length -gt 0 })) { "`t`t<model name=""$m"" altname="""" enabled=""false"" />`n" })
		</models>
	</config>
"@
	
	$template | Out-File $OutXML -force
}

main