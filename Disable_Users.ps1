#Prompts inline
$UPN = Read-Host -Prompt 'Enter your E-Mail Address'
$Description = Read-Host -Prompt 'Description of Search'

#Specify CSV file to use and imports it
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
	InitialDirectory = [Environment]::GetFolderPath('Desktop')
	Filter = 'csv (*.csv)|*.csv'
	}
$FileBrowser.ShowDialog() | Out-Null
$CSVPath = $FileBrowser.FileName
$email = (Import-Csv $CSVPath).email

#Connect to Security and Compliance Center
Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName $UPN

#Start Compliance Search
Write-Host "Creating and running search: " $Description
$search = New-ComplianceSearch -Name $Description -ExchangeLocation $email | Start-ComplianceSearch
while ((Get-ComplianceSearch $search.Name).Status -ne "Completed")
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