<#
	Author: 	Keith Fenech
	Date: 		29/04/2020
	Version:	1.0

	Description:
	Output files with specific extention
#>

# Get aliases of all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited | where{$_.emailaddresses -like "*@domain.com"}
 
# Email address to remove
$ExcludedSMTPAddress = "*@domain.com"
 
# Run through list of users attempting to remove smtp address
$Counter = 0
ForEach ($User in $Mailboxes) {
    Try {
        $VarMailbox = Get-Mailbox -Identity "$User"
        $AllAddresses = $VarMailbox | Select -ExpandProperty EmailAddresses
        $ProxyAddresses = $AllAddresses | Where {$_ -is "Microsoft.Exchange.Data.SmtpProxyAddress"} | Select -ExpandProperty SmtpAddress
        # If the mailbox has the excluded address and it is not the primary SMTP address remove it
        If (($ProxyAddresses -like "$ExcludedSMTPAddress") -and ($VarMailbox.PrimarySmtpAddress.ToString() -notlike "$ExcludedSMTPAddress")) {
            $Counter++
            $AddressesToKeep = $AllAddresses | Where {$_.SmtpAddress -notlike "$ExcludedSMTPAddress"}
            Set-Mailbox -Identity "$User" -EmailAddresses $AddressesToKeep
        }
    }
    Catch {Write-Warning -Message "A problem occurred for $User"}
}
 
Write-Host -Object "Finished! The address '$ExcludedSMTPAddress' was found and removed from $Counter mailboxes" -ForegroundColor Green
