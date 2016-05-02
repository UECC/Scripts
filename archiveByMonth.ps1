<# 
.SYNOPSIS 
   ZIP old files by monthly .ZIP files
.DESCRIPTION
   Created a .ZIP files in a old\ folder containing files per monthly which is older than first day of the previous month
.NOTES 
   Version: 0.29
   Author: UECC, Henrik Næss
#> 

#$zipPath = "C:\work\PortableApps\PortableApps\7-ZipPortable\App\7-Zip64\7z.exe"		#LOCAL
$zipPath = "C:\Program Files (x86)\PortableApps\7-ZipPortable\App\7-Zip64\7z.exe"
$archivemonths = 1

$tempFolder = gc env:temp 
$cutoffdate = (Get-Date -Hour 0 -Minute 0 -Second 0 -Day 1).AddMonths(-$archivemonths)


Function CheckRequiredSoftware
{
	if(!(Test-Path $zipPath))
	{
		Write-Host "7-Zip program $($zipPath) is missing. Halted!"
		return
	}
}

Function ArchiveFilesByMonth
{
    param(
        [Parameter(Mandatory=$true)][string] $archivefolder,    
		[Parameter(Mandatory=$true)][string] $includePattern
        )

	if(!(Test-Path $archiveFolder))
	{
		Write-Host "Folder $($archiveFolder) doesn't exists"
		return
	}
	Set-Location $archivefolder
	
	$oldFolder = Join-Path $archivefolder "_old\"
	if(!(Test-Path $archiveFolder)) { New-Item $oldFolder -itemtype directory }

	$randomNo = Get-Random
	$filelist = $tempFolder + "filelist$($randomNo).txt"
    $myarray=@()

	# TODO: Improve to include only relevant files
	gci $archivefolder -recurse -include $includePattern -exclude *.zip | ?{$_.LastWriteTime -lt $cutoffdate} | %{
		#record the month & year of the files in an array
		$temp='' | Select Filenamex, Period
        $temp.Filenamex = ($_ | Resolve-Path -Relative).TrimStart(".\")
        $temp.Period = '{0:MMyyyy}' -f $_.LastWriteTime
        $myarray += $temp
	}
	
	#Get each unique month/year from the array
	foreach($period in ($myarray | Select Period -unique)){
		$zipFileYYYYmm = $oldFolder + ($period.Period).Substring(2, 4) + ($period.Period).Substring(0, 2) + ".zip"
		
		$myarray | ?{$_.Period -eq $period.period} | Select -expand Filenamex | out-file $filelist -encoding ASCII
		$argumentlist = "a -tzip $($zipFileYYYYmm) @$($filelist)"
		
        $result = start-process "$($zipPath)" -argumentlist $argumentlist -NoNewWindow -wait
        
		# If .ZIP created then delete the files
		# can later be changed to move "if using 7-Zip 9.22 beta or later"
		if(Test-Path $zipFileYYYYmm) { gc $filelist | %{gci $_} | %{$_.Delete() } }
	}
	Write-Host "Completed $($archivefolder)"
}

Function HorizonInOut
{
	$rootPath = "d:\softship\IN_OUT\"
	ArchiveFilesByMonth ($rootPath + "APERAK_ECS\ARCHIVE") IC01.*
	ArchiveFilesByMonth ($rootPath + "APERAK_ECS\LOGS") APERAK?CS_*.log
		
	ArchiveFilesByMonth ($rootPath + "APERAK_ICS\ARCHIVE") IC01.*
	ArchiveFilesByMonth ($rootPath + "APERAK_ICS\LOGS") APERAK?CS_*.log	
	
	ArchiveFilesByMonth ($rootPath + "CustomerIN\Archive") *_*.xml
	ArchiveFilesByMonth ($rootPath + "CustomerIN\Logs\Backup") XMLCSMSG*.log

	ArchiveFilesByMonth ($rootPath + "CustomerOUT\Archive") *_*.xml
	ArchiveFilesByMonth ($rootPath + "CustomerOUT\Logs\Backup") XMLCSMSG*.log

	ArchiveFilesByMonth ($rootPath + "dRoE2Line\Archive") Exchange_Rates_*.xml
	ArchiveFilesByMonth ($rootPath + "dRoE2Line\Log\Backup") logfile_*.log
	
	ArchiveFilesByMonth ($rootPath + "IFCCUS_IMP\ARCHIVE") *Info*.xml
	ArchiveFilesByMonth ($rootPath + "IFCCUS_IMP\ERROR") *Info*.xml
	ArchiveFilesByMonth ($rootPath + "IFCCUS_IMP\LOG") log*.txt
	
	ArchiveFilesByMonth ($rootPath + "ifspurord\IFSPUROUT\archive") IFSPUR*.XML
	ArchiveFilesByMonth ($rootPath + "ifsrevvou\IFSREVOUT\archive") IFSREV*.XML
	ArchiveFilesByMonth ($rootPath + "ifssupvou\IFSSUPOUT\archive") IFSSUP*.XML
	
	ArchiveFilesByMonth ($rootPath + "IFTMCS_ECS\ARCHIVE") *_*_IFTMCS_ECS*.TXT

	ArchiveFilesByMonth ($rootPath + "IFTMCS_ICS\ARCHIVE") *_IFTMCS_ICS*.TXT

	ArchiveFilesByMonth ($rootPath + "IFTMCS_IMP\OUT\Backup") IMP.BIP.IFTMCS.*.TXT

	ArchiveFilesByMonth ($rootPath + "IFTMCS_MRN_ECS\ARCHIVE\MailFormatFiles") IC01.GS.*.mailFormat
	ArchiveFilesByMonth ($rootPath + "IFTMCS_MRN_ECS\ARCHIVE\MRNFiles") IC01.GS.*

	ArchiveFilesByMonth ($rootPath + "IFTMCS_MRN_IMP\ARCHIVE\MailFormatFiles") IC01.CK.*.mailFormat
	ArchiveFilesByMonth ($rootPath + "IFTMCS_MRN_IMP\ARCHIVE\MRNFiles") IC01.CK.*
	
	ArchiveFilesByMonth ($rootPath + "IFTMCSIMP_APERAK\IN\Backup") BIP.UECC*.edi

	ArchiveFilesByMonth ($rootPath + "IFTSTAAN\ARCHIVE") *IFTSTAAN*.TXT

	ArchiveFilesByMonth ($rootPath + "LoadingInstruction\archive") *_*.xml
	ArchiveFilesByMonth ($rootPath + "LoadingInstruction\log") LoadingInstruction_*.log

	ArchiveFilesByMonth ($rootPath + "LoadingResult\archive") *_*.xml
	ArchiveFilesByMonth ($rootPath + "LoadingResult\log\archive") *.log
	ArchiveFilesByMonth ($rootPath + "LoadingResult\log\archive") ~*.tmp			# Temporary (since error in interface)
	
	ArchiveFilesByMonth ($rootPath + "OB10XML\ARCHIVE") *_*_UECC*.XML

	ArchiveFilesByMonth ($rootPath + "PrintSpool\Archive") *-*-*.pdf
	ArchiveFilesByMonth ($rootPath + "PrintSpool\MailFolder") *-*-*

	ArchiveFilesByMonth ($rootPath + "RAOE_AUTOUPD\Logs\Backup") AutoVoyROElog.*.txt

	ArchiveFilesByMonth ($rootPath + "Reports\Backup") *
	ArchiveFilesByMonth ($rootPath + "Reports\PDF") *.pdf
	ArchiveFilesByMonth ($rootPath + "Reports\temp") *.psr

	ArchiveFilesByMonth ($rootPath + "ScheduleChangeNotification\Out") *.txt

	ArchiveFilesByMonth ($rootPath + "Standard\IN\Voucher\HQ") *.pdf

	ArchiveFilesByMonth ($rootPath + "TariffNotification") *-*.xml					# Interface not in use?
	
	ArchiveFilesByMonth ($rootPath + "VendorIN\Archive") *_*.xml
	ArchiveFilesByMonth ($rootPath + "VendorIN\Logs\Backup") XMLCSMSG*.log
	
	ArchiveFilesByMonth ($rootPath + "XMLCargoStatus\Logs\Backup") XMLCSMSG*.log
	
	# PRINTSPOOL 					# at root (obs! recursve)
	#Save_as_Excel # IGNORE (few files)
}

CheckRequiredSoftware

HorizonInOut
