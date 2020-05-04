<#
	Author: 	Keith Fenech
	Date: 		29/04/2020
	Version:	1.0

	Description:
	Output files with specific extention
#>

# Get aliases of all mailboxes
$Mailboxes = Get-Mailbox -ResultSize Unlimited | where{$_.emailaddresses -like "*@midimalta.com"}
 
# Email address to remove
$IncludeSMTPAddress = "*@midimaltaPLC.mail.onmicrosoft.com"
 
# Run through list of users attempting to remove smtp address
$Counter = 0
ForEach ($User in $Mailboxes) {
    Try {
        $VarMailbox = Get-Mailbox -Identity "$User"
        $AllAddresses = $VarMailbox | Select -ExpandProperty EmailAddresses
        $ProxyAddresses = $AllAddresses | Where {$_ -is "Microsoft.Exchange.Data.SmtpProxyAddress"} | Select -ExpandProperty SmtpAddress
        # If the mailbox does not have the included address add it
        If ($ProxyAddresses -notlike "$IncludeSMTPAddress") {
            $Counter++
	Set-Mailbox -Identity "$User" -EmailAddresses @{Add="$($User.Alias)@midimaltaPLC.mail.onmicrosoft.com"}
        }
    }
    Catch {Write-Warning -Message "A problem occurred for $User"}
}
 
Write-Host -Object "Finished! The address '$IncludeSMTPAddress' was added to $Counter mailboxes" -ForegroundColor Green
