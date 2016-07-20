#Requires -Version 5.0 -RunAsAdministrator

<#
	.SYNOPSIS
		Downloads Softpaqs from HP Inc for each operating system and model specified.
	
	.DESCRIPTION
		Downloads Softpaqs from HP Inc for each operating system and model specified.
		
		+Requries Powershell 5.0
		+Requries PSFTP module (https://gallery.technet.microsoft.com/scriptcenter/PowerShell-FTP-Client-db6fe0cb)
	
	.PARAMETER cfgFile
		Specify a configuration file. (Defaults to script folder)
	
	.PARAMETER outDir
		Specify a directory to store Softpaqs. (Defaults to script folder)

	.PARAMETER Update
		Downloads catalog and updates config file. Generates config file if one is not present.

	.PARAMETER UseSymLinks
		Uses symbolic links for each OS/Model combination instead of creating duplicates. If not enabling this make sure you have a large drive or deduplication enabled.

	.EXAMPLE 
		Initial Usage - You must run this to generate the config file.
		Don't forget to edit config file after running this
				PS C:\> HPDriverDownloader.ps1 -OutDir D:\Drivers -Update
	.EXAMPLE
		Downloads softpaq file to 'D:\Drivers'. Creates symlinks instead of copying files for each model.
				PS C:\> HPDriverDownloader.ps1 -OutDir D:\Drivers -UseSymLink
	
	.NOTES
		Additional information about the file.
#>
[CmdletBinding()]
param
(
	[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
	[String]$cfgFile,

	#[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
	#[String]$outDir,

	#[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
	#[switch]$UseSymLinks,

	[Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
	[switch]$Update
)

Import-Module PSFTP

$processedFile = "$PSScriptRoot\AlreadyProcessed.txt"
$procModelFile = "$PSScriptRoot\ProcessedModels.txt"

$secpasswd = ConvertTo-SecureString "password" -AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ("anonymous", $secpasswd)

#Set Globals
if ($Script:cfgFile -eq ""){$Script:cfgFile = "$PSScriptRoot\config.xml"}


function main {
	cls
	
	if ($Update) {
		if ((!(Test-Path "$outDir\ProductCatalog\modelcatalog.xml"))) {
			Write-Verbose "No catalog found Downloading and extracting."
			DownloadCatalog -DownloadDir "$outDir"
			$modelCatalog = Get-ModelCatalog -CatalogDir "$outDir\ProductCatalog"
			BuildConfig -ConfigTemplatePath "$PSScriptRoot\config.template.xml" -Catalog $modelCatalog -ConfigFilePath $cfgFile
		} else {
			Write-Verbose "Checking for new catalog from HP."
			DownloadCatalog -DownloadDir "$outDir\temp"
			$newModelCatalog = Get-ModelCatalog -CatalogDir "$outDir\temp\ProductCatalog"
			$modelCatalog = Get-ModelCatalog -CatalogDir "$outDir\ProductCatalog"
			Write-Verbose "Found new catalog version $($newModelCatalog.CatalogVersion). ($($newModelCatalog.CatalogVersion)>$($modelCatalog.CatalogVersion))"
			if ( ([version]($newModelCatalog.CatalogVersion)) -gt ([version]($modelCatalog.CatalogVersion)) ) {
				Write-Verbose "Found new catalog version $($newModelCatalog.CatalogVersion). ($($newModelCatalog.CatalogVersion)>$($modelCatalog.CatalogVersion))"
				Try {
					Remove-Item Remove-Item -Recurse -Force "$outDir\ProductCatalog"
					Copy-Item "$outDir\temp\ProductCatalog" "$outDir\ProductCatalog" -Recurse
					$modelCatalog = $newModelCatalog
				} Catch {
					Throw $Error.Exception
				}
				BuildConfig -ConfigTemplatePath "$PSScriptRoot\config.template.xml" -Catalog $modelCatalog -ConfigFilePath $cfgFile -Update
				
			}
		}
		Write-Output "You must enable Language, Model, Category and OS in the config file for this script to do anything.`n""$cfgFile"""
	} else {
		<#
		if (!(Test-Path $Script:cfgFile)) {
			Throw "ERROR: ""$Script:cfgFile"" not found. Use -Update switch to generate"
		} elseif (!(Test-Path "$outDir\ProductCatalog\modelcatalog.xml")) {
			Throw "ERROR: ""$outDir\ProductCatalog\modelcatalog.xml"" not found. Use -Update switch to generate"
		} else {
		#>
			
			Write-Output "Loading configuration from $cfgFile..."
			$config = ([xml](gc $cfgFile)).config
			
			if ($config.OutDir.length -gt 0){
				$Script:outDir = $config.OutDir
			} else {
				$Script:outDir = "$PSScriptRoot\Softpaqs"
			}
			
			if ($config.OutputStructure.length -gt 0){
				$Script:outStruct = $config.OutputStructure
			} else {
				#Output hardlinks by default
				$Script:outStruct = '3'
			}
			
			Write-Output "Loading catalog..."
			$modelCatalog = Get-ModelCatalog -CatalogDir "$outDir\ProductCatalog"
			
			Write-Output "Process the catalog against our configuration to build a processing queue..."
			$queue = (ProcessCatalog -ModelCatalog $modelCatalog -Config $config -CatalogDir "$outDir\ProductCatalog") |
			group 'LangId', 'SoftpaqName', 'OsId', 'ModelId' |
			Foreach-Object { $_.Group | Sort-Object SoftpaqVersionR | Select-Object -Last 1 }
			
			Write-Output "Download the requisite Softpaqs..."
			DownloadandExtractSoftpaqs -Queue ([ref]$queue) -OutDir $outDir
			
			Write-Output "Building driver structure..."
			Switch ($Script:ourStruct){
				1 { Create-PathsAndCopyFiles -Queue ([ref]$queue) -OutDir $outDir }
				2 { Create-DriverStructure -Queue ([ref]$queue) -OutDir $outDir }
				3 { Create-DriverStructure -Queue ([ref]$queue) -OutDir $outDir }
				default { Create-DriverStructure -Queue ([ref]$queue) -OutDir $outDir }
			}
		#}
	}
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
			try {
				FtpDownload -RemoteFile $sp.SoftpaqUrl -Destination "$outDir\Softpaq" -Credentials $mycreds
				ExtractSoftpaq -SoftpaqPath $spFile -Destination $spFldr
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Extracted' }
				Remove-Item $spFile -Force
			} catch {
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Error: Unable to download' }
				$Queue.Value | where { $_.SoftpaqId -eq $sp.SoftpaqId } | foreach { $_.FileStatus = 'Error: Cannot extract' }
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

function String2Version {
	param (
		[string]$String
	)
	Write-Verbose "VERSION |$String|"
	$stringAry = $String.split('.')
	$revAry = ($stringAry[3]).Split(' ')
	$revStr = "$($revAry[0])".PadLeft(6, '0') + "$([int][char]("$($revAry[1])"))".PadLeft(2, '0') + "$($revAry[2])"
	
	return [version]"$($stringAry[0]).$($stringAry[1]).$($stringAry[2]).$revStr"
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
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[string]$RemoteFile,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[string]$Destination,
		[Parameter(Mandatory = $true,
				   Position = 2)]
		$Credentials
	)
	
	$ftpServer = ($RemoteFile -split ('/'))[2]
	$ftpPath = $RemoteFile -replace ("ftp://$ftpServer", "")
	
	Try {
		#Create folder to hold files
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
		 "Unable to write file"
	} Catch {
		$_.Exception
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


function Escape {
	param (
		[string]$string
	)
	return [System.Security.SecurityElement]::Escape($string)
}

function BuildConfig {
	[CmdletBinding(DefaultParameterSetName = 'New')]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[hashtable]$Catalog,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[string]$ConfigFilePath,
		[Parameter(Mandatory = $true,
				   Position = 2)]
		[string]$ConfigTemplatePath,
		[Parameter(Position = 3)]
		[switch]$Update
	)
	
	if ($Update) {
		# Backup config
		Copy-Item -Path $ConfigFilePath -Destination "$ConfigFilePath.old" -Force
		
		# Read currennt config
		$curConfig = ([xml](gc "C:\Users\e139068.HOUTX\Documents\Scripts\HP\HpDriverDownloader\testconfig.xml")).config
		
		# Find enabled settings in the current catalog
		$langDiff = @($curConfig.languages.language | where { $_.enabled -eq 'true' })
		$osDiff = @($curConfig.operatingsystems.os | where { $_.enabled -eq 'true' })
		$modDiff = @($curConfig.models.model | where { $_.enabled -eq 'true' })
		$catDiff = @($curConfig.categories.category | where { $_.enabled -eq 'true' })
	}
	
	# Build xml for new config
	$languages = @()
	foreach ($entry in ($catalog.Languages.Values)) {
		if ($langDiff.Id -contains $entry.Id) {
			$languages += "`t`t<language id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" enabled=""true"" />`n"
		} else {
			$languages += "`t`t<language id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" enabled=""false"" />`n"
		}
	}
	
	$oses = @()
	foreach ($entry in ($Catalog.OperatingSystems.Values)) {
		if ($osDiff.Id -contains $entry.Id) {
			$oses += "`t`t<os id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" enabled=""true"" />`n"
		} else {
			$oses += "`t`t<os id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" enabled=""false"" />`n"
		}
	}
	
	$models = @()
	foreach ($entry in ($catalog.Models.Values)) {
		if ($modDiff.Id -contains $entry.Id) {
			$altName = $curConfig.models.model | ? { $_.id -eq $entry.Id } | select -expandproperty altname
			$models += "`t`t<model id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" altname=""$altName"" enabled=""true"" />`n"
		} else {
			$models += "`t`t<model id=""$($entry.Id)"" name=""$(Escape($entry.Name))"" altname="""" enabled=""false"" />`n"
		}
	}
	
	$categories = @()
	foreach ($entry in ($modelCatalog.Softpaqs.Values.Category | select -Unique)) {
		if ($catDiff.Name -contains $entry) {
			$categories += "`t`t<category name=""$entry"" enabled=""true"" />`n"
		} else {
			$categories += "`t`t<category name=""$entry"" enabled=""false"" />`n"
		}
	}
	
	# Transform template into new config
	$cfgData = (gc $ConfigTemplatePath)
	$cfgData = $cfgData.replace('<!languages!>', $languages)
	$cfgData = $cfgData.replace('<!operatingsystems!>', $oses)
	$cfgData = $cfgData.replace('<!models!>', $models)
	$cfgData = $cfgData.replace('<!categories!>', $categories)
	$cfgData = $cfgData.replace('<!catver!>', $catalog.CatalogVersion)
	
	# Write new config file
	$cfgData | Out-File $ConfigFilePath -Force -Encoding utf8
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
	
	$modelCatalog.Add('CatalogVersion', $modelCatalogXml.CatalogVersion)
	
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
	$outHash = [Ordered]@{ }
	
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
			if (($cfgCat.name) -contains (($ModelCatalog.Softpaqs)[$sp.Id].Category)) {
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
					#'SoftpaqVersionR' = String2Version(($ModelCatalog.Softpaqs)[$sp.Id].Version)
					'SoftpaqVersionR' = ($ModelCatalog.Softpaqs)[$sp.Id].Version.split('.').split(' ')
					'SoftpaqDir' = ''
					'ModelId' = $sp.ParentNode.ParentNode.ParentNode.Id
					'ModelName' = ($ModelCatalog.Models)[$sp.ParentNode.ParentNode.ParentNode.Id].Name
					#'ModelCustomName' = ($ModelCatalog.Models)[$sp.ParentNode.ParentNode.ParentNode.Id].altname
					'ModelCustomName' = ($Config.models.model | ? {$_.name -eq (($ModelCatalog.Models)[$sp.ParentNode.ParentNode.ParentNode.Id].Name)} | select -expandproperty altname)
					'FileStatus' = ''
					'Status' = ''
				});
			}
		}
	}
	
	$processedData
}

function Create-Hardlinks {
	param (
		[string] $outDir,
		[string] $spDir
	)
	

	
	$tree = (gci $spDir -Recurse)
		
	#Create folder Structure
	foreach ($item in ($tree | where { $_.PSIsContainer -eq $true })){
			$relDir = $item.FullName -ireplace ([regex]::Escape($spDir), "")
			New-Item -ItemType Directory -Path "$outDir$relDir" -ErrorAction SilentlyContinue
	
	}
	
	#Create hardlinks
	foreach ($item in ($tree | where { $_.PSIsContainer -eq $false })){
		$target = $item.FullName
		$link = $item.FullName -ireplace ([regex]::Escape($spDir), "$outDir")
        if (!(test-path $link)){
		    New-Hardlink -Link $link -Target $target
        } else {
            write-host "Link ""$link"" already exists"
        }
	}
}

function Create-DriverStructure {
	param
	(
		[Parameter(Mandatory = $true, Position = 1)]
		[ref]$Queue,
		
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$OutDir
	)
	
	foreach ($entry in ($Queue.Value | where { $_.FileStatus -eq 'Extracted' })) {
		$drvTitle = "$($entry.SoftpaqName) ($($entry.SoftpaqVersion))"
		if ($entry.ModelCustomName.length -gt 1) {
			$symDir = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelCustomName)"
		} else {
			$symDir = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelName)"
		}
		$spDir = "$outDir\softpaq\sp$($entry.SoftpaqId)"
		
		if (!(Test-Path $symDir)) {
			New-Item -Path $symDir -ItemType directory -ErrorAction Continue | Out-Null
		}
		
		#Junction points
		<#
		if (!(Test-Path "$symDir\$drvTitle")) {
			New-SymLink -Path $spDir -SymName "$symDir\$drvTitle" -Directory | Out-Null
		}
		#>
		
		#New-HardLink -Link "$symDir\$drvTitle" -Target $spDir
		Create-Hardlinks -outDir "$symDir\$drvTitle" -spDir $spDir
		
	}
	
}

function Create-JuntionPoints {
	param
	(
		[Parameter(Mandatory = $true, Position = 1)]
		[ref]$Queue,
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$OutDir
	)
	
	foreach ($entry in ($Queue.Value | where { $_.FileStatus -eq 'Extracted' })) {
		$drvTitle = "$($entry.SoftpaqName) ($($entry.SoftpaqVersion))"
		if ($entry.ModelCustomName.length -gt 1) {
			$symDir = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelCustomName)"
		} else {
			$symDir = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelName)"
		}
		$spDir = "$outDir\softpaq\sp$($entry.SoftpaqId)"
		
		if (!(Test-Path $symDir)) {
			New-Item -Path $symDir -ItemType directory -ErrorAction Continue | Out-Null
		}
		
		#Junction points
		<#
		if (!(Test-Path "$symDir\$drvTitle")) {
			New-SymLink -Path $spDir -SymName "$symDir\$drvTitle" -Directory | Out-Null
		}
		#>
		
		New-HardLink -Link "$symDir\$drvTitle" -Target $spDir
		
	}
	
}

function Create-PathsAndCopyFiles {
	param
	(
		[Parameter(Mandatory = $true, Position = 1)]
		[ref]$Queue,
		[Parameter(Mandatory = $true, Position = 2)]
		[string]$OutDir
	)
	
	#Create Path and copy drivers
	foreach ($entry in ($Queue.Value | where { $_.FileStatus -eq 'Extracted' })) {
		
		$drvTitle = "$($entry.SoftpaqName) ($($entry.SoftpaqVersion))"
		$drvPath = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelName)\$drvTitle)"
		$spDir = "$outDir\softpaq\sp$($entry.SoftpaqId)"
		
		if (!(Test-Path $drvPath)) {
			New-Item -Path $drvPath -ItemType directory -ErrorAction Continue
			
			Copy-Item -Path "$spDir\*" -Destination $drvPath -Recurse -ErrorAction Continue
		}
		
	}
}

function Supersede {
	param (
		[string]$ReferenceString,
		[string]$DifferenceString
	)
	
	$refAry = $ReferenceString.split('.').split(' ')
	$difAry = $DifferenceString.split('.').split(' ')
	
	$max = ($refAry.length) - 1
	
	foreach ($i in 0..$max) {
		if ($difAry[$i] -gt $refAry[$i]) {
			$result = $true
			break;
		} else {
			$result = $false
		}
	}
	return $result
}

Function New-SymLink {
    <#
        .SYNOPSIS
            Creates a Symbolic link to a file or directory

        .DESCRIPTION
            Creates a Symbolic link to a file or directory as an alternative to mklink.exe

        .PARAMETER Path
            Name of the path that you will reference with a symbolic link.

        .PARAMETER SymName
            Name of the symbolic link to create. Can be a full path/unc or just the name.
            If only a name is given, the symbolic link will be created on the current directory that the
            function is being run on.

        .PARAMETER File
            Create a file symbolic link

        .PARAMETER Directory
            Create a directory symbolic link

        .NOTES
            Name: New-SymLink
            Author: Boe Prox
            Created: 15 Jul 2013


        .EXAMPLE
            New-SymLink -Path "C:\users\admin\downloads" -SymName "C:\users\admin\desktop\downloads" -Directory

            SymLink                          Target                   Type
            -------                          ------                   ----
            C:\Users\admin\Desktop\Downloads C:\Users\admin\Downloads Directory

            Description
            -----------
            Creates a symbolic link to downloads folder that resides on C:\users\admin\desktop.

        .EXAMPLE
            New-SymLink -Path "C:\users\admin\downloads\document.txt" -SymName "SomeDocument" -File

            SymLink                             Target                                Type
            -------                             ------                                ----
            C:\users\admin\desktop\SomeDocument C:\users\admin\downloads\document.txt File

            Description
            -----------
            Creates a symbolic link to document.txt file under the current directory called SomeDocument.
    #>
	[cmdletbinding(
				   DefaultParameterSetName = 'Directory',
				   SupportsShouldProcess = $True
				   )]
	Param (
		[parameter(Position = 0, ParameterSetName = 'Directory', ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
		[parameter(Position = 0, ParameterSetName = 'File', ValueFromPipeline = $True,
				   ValueFromPipelineByPropertyName = $True, Mandatory = $True)]
		[ValidateScript({
			If (Test-Path $_) { $True } Else {
				Throw "`'$_`' doesn't exist!"
			}
		})]
		[string]$Path,
		[parameter(Position = 1, ParameterSetName = 'Directory')]
		[parameter(Position = 1, ParameterSetName = 'File')]
		[string]$SymName,
		[parameter(Position = 2, ParameterSetName = 'File')]
		[switch]$File,
		[parameter(Position = 2, ParameterSetName = 'Directory')]
		[switch]$Directory
	)
	Begin {
		Try {
			$null = [mklink.symlink]
		} Catch {
			Add-Type @"
            using System;
            using System.Runtime.InteropServices;
 
            namespace mklink
            {
                public class symlink
                {
                    [DllImport("kernel32.dll")]
                    public static extern bool CreateSymbolicLink(string lpSymlinkFileName, string lpTargetFileName, int dwFlags);
                }
            }
"@
		}
	}
	Process {
		#Assume target Symlink is on current directory if not giving full path or UNC
		If ($SymName -notmatch "^(?:[a-z]:\\)|(?:\\\\\w+\\[a-z]\$)") {
			$SymName = "{0}\{1}" -f $pwd, $SymName
		}
		$Flag = @{
			File = 0
			Directory = 1
		}
		If ($PScmdlet.ShouldProcess($Path, 'Create Symbolic Link')) {
			Try {
				$return = [mklink.symlink]::CreateSymbolicLink($SymName, $Path, $Flag[$PScmdlet.ParameterSetName])
				If ($return) {
					$object = New-Object PSObject -Property @{
						SymLink = $SymName
						Target = $Path
						Type = $PScmdlet.ParameterSetName
					}
					$object.pstypenames.insert(0, 'System.File.SymbolicLink')
					$object
				} Else {
					Throw "Unable to create symbolic link!"
				}
			} Catch {
				Write-warning ("{0}: {1}" -f $path, $_.Exception.Message)
			}
		}
	}
}
####################################################
# http://zduck.com/2013/mklink-powershell-module/
# Joshua Poehls
####################################################
function New-Hardlink {
    <#
    .SYNOPSIS
        Creates a hard link.
    #>
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [string] $Link,
        [Parameter(Position=1, Mandatory=$true)]
        [string] $Target,
        [Parameter(Position=2)]
        [switch] $Force
    )

    Try {
		Invoke-MKLINK -Link $Link -Target $Target -HardLink -Force $Force
	} Catch {
		Write-Error "Unable to create link $Link"
	}
}

function Invoke-MKLINK {
    <#
    .SYNOPSIS
        Creates a symbolic link, hard link, or directory junction.
    #>
    [CmdletBinding(DefaultParameterSetName = "Symlink")]
    param (
        [Parameter(Position=0, Mandatory=$true)]
        [string] $Link,
        [Parameter(Position=1, Mandatory=$true)]
        [string] $Target,

        [Parameter(ParameterSetName = "Symlink")]
        [switch] $Symlink = $true,
        [Parameter(ParameterSetName = "HardLink")]
        [switch] $HardLink,
        [Parameter(ParameterSetName = "Junction")]
        [switch] $Junction,
        [Parameter()]
        [bool] $Force
    )
    
    # Resolve the paths incase a relative path was passed in.
    $Link = (Force-Resolve-Path $Link)
    $Target = (Force-Resolve-Path $Target)

    # Ensure target exists.
    if (-not(Test-Path $Target)) {
        throw "Target does not exist.`nTarget: $Target"
    }

    # Ensure link does not exist.
    if (Test-Path $Link) {
        if ($Force) {
            Remove-Item $Link -Recurse -Force
        }
        else {
            throw "A file or directory already exists at the link path.`nLink: $Link"
        }
    }

    $isDirectory = (Get-Item $Target).PSIsContainer
    $mklinkArg = ""

    if ($Symlink -and $isDirectory) {
        $mkLinkArg = "/D"
    }

    if ($Junction) {
        # Ensure we are linking a directory. (Junctions don't work for files.)
        if (-not($isDirectory)) {
            throw "The target is a file. Junctions cannot be created for files.`nTarget: $Target"
        }

        $mklinkArg = "/J"
    }

    if ($HardLink) {
        # Ensure we are linking a file. (Hard links don't work for directories.)
        if ($isDirectory) {
            throw "The target is a directory. Hard links cannot be created for directories.`nTarget: $Target"
        }

        $mkLinkArg = "/H"
    }

    # Capture the MKLINK output so we can return it properly.
    # Includes a redirect of STDERR to STDOUT so we can capture it as well.
    $output = cmd /c mklink $mkLinkArg `"$Link`" `"$Target`" 2>&1

    if ($lastExitCode -ne 0) {
        throw "MKLINK failed. Exit code: $lastExitCode`n$output"
    }
    else {
        Write-Output $output
    }
}

function Force-Resolve-Path {
    <#
    .SYNOPSIS
        Calls Resolve-Path but works for files that don't exist.
    .REMARKS
        From http://devhawk.net/2010/01/21/fixing-powershells-busted-resolve-path-cmdlet/
    #>
    param (
        [string] $FileName
    )

    $FileName = Resolve-Path $FileName -ErrorAction SilentlyContinue `
                                       -ErrorVariable _frperror
    if (-not($FileName)) {
        $FileName = $_frperror[0].TargetObject
    }

    return $FileName
}

main