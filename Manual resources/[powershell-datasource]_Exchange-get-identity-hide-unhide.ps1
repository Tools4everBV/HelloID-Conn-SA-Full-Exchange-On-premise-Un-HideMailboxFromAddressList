$name = $datasource.name

# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -Authentication Basic -AllowRedirection -SessionOption $sessionOption
    Import-PSSession -Session $exchangeSession -AllowClobber
    Write-Information "Successfully connected to Exchange '$ExchangeConnectionUri'"
} catch {
    Write-Error "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'"
}

try {
    $searchQuery = "*$Name*"
    $mailboxes = Get-Mailbox -ResultSize:Unlimited -Filter "{Alias -like '$searchQuery' -or name -like '$searchQuery'}"
    $mailboxes = $mailboxes | Sort-Object -Property DisplayName, HiddenFromAddressListsEnabled
    $resultCount = @($mailboxes).Count
    Write-Information "Result count: $resultCount"
    if($resultCount -gt 0)
    {
        foreach($mailbox in $mailboxes){
            $returnObject = @{
                name=$mailbox.DisplayName  + " [" + $mailbox.SamAccountName + "]"; 
                sAMAccountName=$mailbox.SamAccountName
                HiddenFromAddressLists = $mailbox.HiddenFromAddressListsEnabled
            }
        }
    }
    Write-output $returnObject
} catch {
    Write-Error "Error generating name, Message '$($_.Exception.Message)'"
}

# Disconnect from Exchange
Remove-PsSession -Session $exchangeSession
Write-Information "Successfully disconnected from Exchange"a
