function global:Out-TFSProjectsToCSV
{
	<# 
	.SYNOPSIS 
		Parse the output from TFSProjects utility and output as a CSV file
	.DESCRIPTION 
		Take the RTF file generated from TFSProjects (https://tfsprojects.codeplex.com/)
		and convert it to a usable CSV file
		
		##############################################################################################
		##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
		##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
		##############################################################################################
	.NOTES 
		File Name  : Out-TFSProjectsToCSV.ps1
		Author     : Larry G. Wapnitsky (larry@wapnitsky.com)
		
		##############################################################################################
		##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
		##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
		##############################################################################################

	.LINK 
		https://tfsprojects.codeplex.com/
	.EXAMPLE 
		Out-TFSProjectsToCSV -TFSProjectsFile tfsp.rtf -CSVOutputFileFolder c:\temp
		
		This will take the tsfp.rtf file, created by TFSProjects, and output
		a CSV file called tsfp.csv to the c:\temp folder
	.EXAMPLE
		Out-TFSProjectsToCSV
		
		Without specifying any options, you will be presented with a file-dialog to
		choose the input file (RTF), and a folder-dialog to choose the destination
		for the CSV file
	.PARAMETER TFSProjectsFile
		(Optional) The name of the RTF file that you saved from TFSProjects.
		
		If not speficied, you will be presented with a file-dialog to choose the input file
	.PARAMETER CSVOutputFileFolder
		(Optional) The folder where the CSV output will be place.  If not specified,
		you will be presented with a folder-dialog to choose the destination
	#>
	
	Param
	(
		$TFSProjectsFile = $(
			Add-Type -AssemblyName System.Windows.Forms
			$dlg = New-Object -TypeName System.Windows.Forms.OpenFileDialog
			$dlg.Filter = "RTF Files (*.rtf)|*.rtf"
			
			if ($dlg.ShowDialog() -eq 'OK')
				{ $dlg.FileName }
			else
				{ throw "No TFSProjects filename submitted"}
		),
		
		$CSVOutputFileFolder = $(
			Add-Type -AssemblyName System.Windows.Forms
			$dlg = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
			$dlg.Description = "Location for CSV Output"
			$dlg.SelectedPath = [io.path]::GetDirectoryName($TFSProjectsFile)
			if ($dlg.ShowDialog() -eq 'OK')
				{ $dlg.SelectedPath }
			else
				{ throw "No CSV Output path submitted"}
		)
	)
	
	function Convert-RTFtoTXT
	{
		Write-Progress -id 1 -Activity "Converting RTF to Text" -Status $TFSProjectsFile
		write-verbose "opening $TFSProjectsFile in Microsoft Word"
		Add-Type -AssemblyName Microsoft.Office.Interop.Word
		
		$wordApp.visible = $false
		$wordapp.DisplayAlerts = [microsoft.office.interop.word.wdalertlevel]::wdAlertsNone
		
		$txtFile = $TFSProjectsFile.Replace("rtf","txt")
		
		$wordDoc = $wordApp.Documents.Open($TFSProjectsFile)
		$wordDoc.SaveAs([ref]$txtFile, [ref][Microsoft.Office.Interop.Word.WDSaveFormat]::wdFormatText)

		$wordDoc.Close()
		
		$wordApp.Quit()
		
		return $txtFile
	}
	
	function TFSProjectsToCSV
	{
		$rx_NameEmail        =                 [regex]"\t*Name\ \:\ (?<Name>[\w]+\,*\ *[\w]*)\t+Email\ \:\ (?<Email>.*)"
		$rx_ProjectGroupName =                 [regex]"\t*Project\ Group\ Name[\ \:]+(?<ProjectGroupName>.*)"
		$rx_ProjectName      =                 [regex]"Project\ Name[\ \:]+(?<ProjectName>.*)"
	
		$TFSProjects = Get-Content $txtFile
		
		$ProjectName = $null
		$GroupName = $null
		
		$p_count = 0
		
		$CSV = ForEach ($line in $TFSProjects)
		{
			$p_count+=1
			
			$Project = $rx_ProjectName.Match($line)
			$GN = $rx_ProjectGroupName.Match($line)
			$NE = $rx_NameEmail.Match($line)
			
			if ($Project.Success)
				{ $ProjectName = $Project.Groups["ProjectName"].Value }
			
			if ($GN.Success)
				{ $GroupName = $GN.Groups["ProjectGroupName"].Value }
			
			if ($NE.Success)
				{ 
					$percentComplete = (($($p_count)/$($TFSProjects.count))*100)
					$percentComplete = "{0:N2}" -f $percentComplete
		
					$Name = $NE.Groups["Name"].Value
					$Email = $NE.Groups["Email"].Value
					
					$NTAccountName = (get-qaduser -Name $Name).NTAccountName
					if ($NTAccountName -eq "") 
					{
						$NTAccountName = (get-qaduser -ProxyAddress "smtp:$($Email)").NTAccountName
					}
					
				
					New-Object -type PSObject -Property @{
						Name = $Name
						Email = $Email
						"Project Name" = $ProjectName
						"Group Name" = $GroupName
						NTAccountName = $NTAccountName
					}
					
					write-progress -id 1 -Activity "$ProjectName | $GroupName | $Name | $Email | $NTAccountName" -Status "$($percentComplete)% Complete" -PercentComplete $percentComplete
				}
		}
		
		$CSVfilename = $TFSProjectsFile.Replace("rtf", "csv")
		$CSV | Select "Project Name","Group Name",Name,NTAccountName,Email | Export-CSV -NoTypeInformation -Path $CSVfilename
	}
	
	##############################################################################################
	##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
	##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
	##############################################################################################

	Add-PSSnapin Quest.ActiveRoles.ADManagement -ErrorAction SilentlyContinue
	$ErrorActionPreference = "SilentlyContinue"
	
	$txtFile = Convert-RTFtoTXT
	TFSProjectsToCSV
}

Write-Host "New Commands added:" -ForegroundColor Green
Write-Host "`tOut-TFSProjectsToCSV" 
Write-Host "`nFor command usage, please use the Get-Help command." -ForegroundColor Yellow

	
	