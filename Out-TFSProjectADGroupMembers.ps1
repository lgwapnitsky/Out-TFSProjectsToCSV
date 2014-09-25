function global:Out-TFSProjectADGroupMembers
{
	<# 
	.SYNOPSIS 
		Takes the output from Out-TFSProjectADGroupMembers.ps1
		and reports the members of any AD groups
	.DESCRIPTION 
		Takes the output from Out-TFSProjectADGroupMembers.ps1
		and reports the members of any AD groups
		##############################################################################################
		##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
		##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
		##############################################################################################
	.NOTES 
		File Name  : Out-TFSProjectsADGroupMembers.ps1
		Author     : Larry G. Wapnitsky (larry@wapnitsky.com)
		
		##############################################################################################
		##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
		##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
		##############################################################################################

	.LINK 
		https://tfsprojects.codeplex.com/
	.EXAMPLE 
		Out-TFSProjectADGroupMembers
		
		Without specifying any options, you will be presented with a file-dialog to
		choose the input file (CSV).  The output will be written in the same folder
		as the input file
	.EXAMPLE 
		Out-TFSProjectADGroupMembers -CSVOutputFile tsfp.csv
		
		This will take the tsfp.csv file, created by Out-TFSProjectsToCSV, and output
		a CSV file called tsfp_ADMembers.csv to the same folder as the CSVOutputFile
	.PARAMETER CSVOutputFile
		
		(optional) If not speficied, you will be presented with a file-dialog to choose the input file
	#> 
	 

	Param
	(
		$CSVOutputFile = $(
			Add-Type -AssemblyName System.Windows.Forms
			$dlg = New-Object -TypeName System.Windows.Forms.OpenFileDialog
			$dlg.Title = "Converted Projects CSV Output"
			$dlg.Filter = "CSV Files (*.csv)|*.csv"
			$dlg.InitialDirectory = $pwd
			if ($dlg.ShowDialog() -eq 'OK')
				{ $dlg.FileName }
			else
				{ throw "No CSV Output file selected"}
		)
	)
	
	$data = import-csv $CSVOutputFile
	$outputFile = $CSVOutputFile.Replace(".csv", "_ADMembers.csv")
	
	$p_count = 0
	$groupList = @{}
	$userList = @{}
	
	$data | foreach-object {
		$p_count+=1
		$percentComplete = (($($p_count)/$($($data.count)))*100)
		$percentComplete = "{0:N2}" -f $percentComplete
	
		$grpName = $_.name
		
		write-progress -id 1 -Activity $grpName -Status "$($percentComplete)% Complete" -PercentComplete $percentComplete
		
		switch ($userList.containsKey($grpName))
		{
			$false {
				$grp = get-qadgroup $grpName
				if (($grp -ne $null) -and ($groupList.containsKey($grp.name) -eq $false))
				{
					$grpMembers = $grp.members | foreach {
						$user = get-qaduser $_
						New-Object PSObject -Property @{
							Group = $grpName
							User = $user.name
							Account = $user.samaccountname
						}
					}
					
					try {
						$groupList.Add($grp.name, $grpMembers) 
						$grpMembers | select Group,User,Account | export-csv -NoTypeInformation -Append $outputFile
					}
					catch {}
				}
			}
		
			$true {
				try {$userList.add($grpName, $null)}
				catch{}
			}
		}
	}
}

Write-Host "New Commands added:" -ForegroundColor Green
Write-Host "`tOut-TFSProjectADGroupMembers" 
Write-Host "`nFor command usage, please use the Get-Help command." -ForegroundColor Yellow