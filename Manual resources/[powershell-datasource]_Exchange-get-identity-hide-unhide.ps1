$name = $datasource.name

# Connect to Exchange
try{
    $adminSecurePassword = ConvertTo-SecureString -String "$ExchangeAdminPassword" -AsPlainText -Force
    $adminCredential = $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername,$adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -ErrorAction Stop 
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber
    Write-Information "Successfully connected to Exchange using the URI [$exchangeConnectionUri]" 
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
            Write-output $returnObject
        }
    }    
} catch {
    Write-Error "Error generating name, Message '$($_.Exception.Message)'"
}

# Disconnect from Exchange
try{
    Remove-PsSession -Session $exchangeSession -Confirm:$false -ErrorAction Stop
    Write-Information "Successfully disconnected from Exchange using the URI [$exchangeConnectionUri]"     
    
} catch {
    Write-Error "Error disconnecting from Exchange.  Error: $($_.Exception.Message)"    
}
