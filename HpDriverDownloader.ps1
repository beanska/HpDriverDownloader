#requires -Version 5.0
<#
	.SYNOPSIS
		Downloads Softpaqs from HP Inc for each operating system and model specified.
	
	.DESCRIPTION
		Downloads Softpaqs from HP Inc for each operating system and model specified.
		
		+Requries Powershell 5.0
		+Requries PSFTP module (https://gallery.technet.microsoft.com/scriptcenter/PowerShell-FTP-Client-db6fe0cb)
	
	.PARAMETER cfgFile
		XML configuration file
	
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
	[String]$outDir = "$PSScriptRoot\Softpaqs",
	[Parameter(ValueFromPipeline = $true,
			   ValueFromPipelineByPropertyName = $true)]
	[switch]$UseSymLinks = $false
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
	$queue = (ProcessCatalog -ModelCatalog $modelCatalog -Config $config -CatalogDir "$outDir\ProductCatalog") |
		group 'LangId', 'SoftpaqName', 'OsId', 'ModelId' |
		Foreach-Object { $_.Group | Sort-Object SoftpaqVersionR | Select-Object -Last 1 }
	
	# Download the requisite Softpaqs
	DownloadandExtractSoftpaqs -Queue ([ref]$queue) -OutDir $outDir
	
	if ($UseSymLinks) {
		Create-JuntionPoints -Queue ([ref]$queue) -OutDir $outDir
	} else {
		Create-PathsAndCopyFiles -Queue ([ref]$queue) -OutDir $outDir
	}
		
	#$queue | group 'LangId', 'SoftpaqName', 'OsId', 'ModelId' |
	#Foreach-Object { $_.Group | Sort-Object SoftpaqVersionR | Select-Object -Last 1 }
	#| export-csv c:\Downloads\queue.csv -Force -NoTypeInformation
	
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
		throw "Unable to write file"
	} Catch {
		throw $_.Exception
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
					'SoftpaqVersionR' = String2Version(($ModelCatalog.Softpaqs)[$sp.Id].Version)
					'SoftpaqDir' = ''
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
		$symDir = "$outDir\$($entry.LangName)\$($entry.OsShortName)\$($entry.ModelName)"
		$spDir = "$outDir\softpaq\sp$($entry.SoftpaqId)"
		
		if (!(Test-Path $symDir)) {
			New-Item -Path $symDir -ItemType directory -ErrorAction Continue | Out-Null
		}
		
		if (!(Test-Path "$symDir\$drvTitle" )) {
			New-SymLink -Path $spDir -SymName "$symDir\$drvTitle" -Directory | Out-Null
		}
		
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
	foreach ($entry in ($Queue.Value | where {$_.FileStatus -eq 'Extracted' }) ) {
		
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
main