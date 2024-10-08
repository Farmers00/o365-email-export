#Prompts inline
$Admin = Read-Host -Prompt 'Enter your E-Mail Address'
$Description = Read-Host -Prompt 'Description of Search'

#Specify CSV file to use and imports it
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	InitialDirectory = [Environment]::GetFolderPath('Desktop')
	Filter = 'csv (*.csv)|*.csv'
	}
$FileBrowser.ShowDialog() | Out-Null
$CSVPath = $FileBrowser.FileName
$Emails = (Import-Csv $CSVPath).UPN

#Connect to Security and Compliance Center
Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName $Admin

#Start Compliance Search
$Locations = @()
Write-Host "Creating and running search: " $Description
ForEach ($Email in $Emails){
	$UserSearch = Get-User -Identity $Email -ErrorAction SilentlyContinue | Select Name
	If ($UserSearch -ne $Null){
		#Removed because missing UPN in o365 - $Locations += ($UserSearch.UserPrincipalName).Where({$_.Trim()})
		$Locations += ($Email)
	}
	Else {
		Write-Host "$Email not found, probably already removed."
	}
}

$Search = New-ComplianceSearch -Name $Description -ExchangeLocation $Locations | Start-ComplianceSearch
While ((Get-ComplianceSearch $Search.Name).Status -ne "Completed"){
	Write-Host " ." -NoNewline
	Start-Sleep -s 3
}

Write-Host "Search Complete!"

#Start Export Action
Write-Host "Begining Export: " $Description -NoNewline
New-ComplianceSearchAction -SearchName $Description -Export -Format Fxstream
Write-Host "Check Security and Compliance Center for Status"

#Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$False
