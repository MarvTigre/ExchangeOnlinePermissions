$title = "Dodawanie/Odbieranie uprawnie� do skrzynki"
$message = "Chcesz nada� czy odebra� uprawnienia?"

$Nadac = New-Object System.Management.Automation.Host.ChoiceDescription "&Nada�", `
    "Nadanie uprawnie� do skrzynki."

$Odebrac = New-Object System.Management.Automation.Host.ChoiceDescription "&Odebra�", `
    "Odebranie uprawnie� do skrzynki."

$options = [System.Management.Automation.Host.ChoiceDescription[]]($Nadac, $Odebrac)

$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

write-host
write-host "==========================================================="
write-host "              Logowanie do Exchange Online"
write-host "==========================================================="
write-host	

$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

write-host
write-host "==========================================================="
write-host "              ��czenie z Exchange Online"
write-host "==========================================================="
write-host	

Import-PSSession $Session

$ErrorActionPreference = "Stop"

	write-host
	write-host "==========================================================="
	write-host
	$User1_email = Read-Host -Prompt 'Komu dajemy / odbieramy uprawnienia? [adres e-mail]:' 
    	$User1 = (Get-Recipient -Identity $User1_email).DisplayName
	write-host
	write-host "==========================================================="

	write-Host
	$User2_email = Read-Host -Prompt 'Do czyjego konta? [adres e-mail]:'
	$User2 = (Get-Recipient -Identity $User2_email).DisplayName
	write-host 
	write-host "==========================================================="

switch ($result)
    {
        0 {	
		Add-MailboxPermission -Identity $User2_email -User $User1_email -AccessRights FullAccess -InheritanceType All -AutoMapping $false
		write-host
		write-host "===================================================================================================="
		write-host
		Write-Host "Nada�e� uprawnienia do konta '$User2' u�ytkownikowi '$User1'"
		write-host
		write-host "===================================================================================================="
		write-host		
		}
        1 {
		Remove-MailboxPermission -Identity $User2_email -User $User1_email -AccessRights FullAccess -InheritanceType All
		write-host
		write-host "===================================================================================================="
		write-host
		Write-Host "Odebra�e� uprawnienia do konta '$User2' u�ytkownikowi '$User1'"
		write-host
		write-host "===================================================================================================="
		write-host		
		}
    }

Write-Host 
Write-Host "Aby zako�czy�, wci�nij dowolny klawisz."
Write-Host

$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

write-host
write-host "==========================================================="
write-host "          Zamykanie po��czenia z Exchange Online"
write-host "==========================================================="
write-host

Remove-PSSession $Session