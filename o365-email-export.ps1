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
$emails = (Import-Csv $CSVPath).UPN

#Connect to Security and Compliance Center
Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName $Admin

#Start Compliance Search
$Locations = @()
Write-Host "Creating and running search: " $Description
ForEach ($email in $emails)
{
	$UserSearch = Get-User -Identity $email
	If ($UserSearch -ne $null)
	{
		$Locations += $UserSearch
	}
	Else {
		Write-Host "$email not found."
	}
}
$search = New-ComplianceSearch -Name $Description -ExchangeLocation $Locations | Start-ComplianceSearch
While ((Get-ComplianceSearch $search.Name).Status -ne "Completed")
	{
    Write-Host " ." -NoNewline
    Start-Sleep -s 3
	}
Write-Host "Search Complete!"

#Start Export Action
Write-Host "Begining Export: " $Description -NoNewline
New-ComplianceSearchAction -SearchName $Description -Export -Format Fxstream
Write-Host "Check Security and Compliance Center for Status"

#Disconnect from Exchange Online
Disconnect-ExchangeOnline