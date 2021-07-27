if ($hidefromGal = "True"){
    $hide = $true
} elseif ($hidefromGal = "False"){
    $hide = $false
}

try{
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeAdminUsername,$adminSecurePassword
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -Authentication Basic -AllowRedirection -SessionOption $sessionOption
    Import-PSSession -Session $exchangeSession -AllowClobber
    HID-Write-Status -Message "Successfully connected to Exchange '$ExchangeConnectionUri'" -Event Success
} catch {
    HID-Write-Status -Message "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'" -Event Error
    HID-Write-Summary -Message "Error connecting to Exchange using the URI '$exchangeConnectionUri', Message '$($_.Exception.Message)'" -Event Failed
}

try {
    Set-Mailbox -Identity $identity -HiddenFromAddressListsEnabled $hide
    HID-Write-Status -Message "Mailbox '$identity' HiddenFromAddressListsEnabled set to '$hide'" -Event Success
    HID-Write-Summary -Message "Mailbox '$identity' HiddenFromAddressListsEnabled set to '$hide'" -Event Success
} catch {
    HID-Write-Status -Message "Could not hide/unhide mailbox with identity '$identity' from Global Address List, Message '$($_.Exception.Message)'" -Event Error
    HID-Write-Summary -Message "Could not hide/unhide mailbox with identity '$identity' from Global Address List, Message '$($_.Exception.Message)'" -Event failed
}

# Disconnect from Exchange
Remove-PsSession -Session $exchangeSession
HID-Write-Status -Message "Successfully disconnected from Exchange" -Event Success
HID-Write-Summary -Message "Successfully disconnected from Exchange" -Event Success
