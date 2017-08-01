#Lord Hagen / olehag04@nfk.no

$RegKeyWaste = "HKCU:\SOFTWARE\Microsoft\Office\15.0\Outlook\Prefrences"
New-Item -Path $RegKeyWaste -Name DelegateSentItemsStyle -Force
New-ItemProperty -Path $RegKeyWaste -Name DelegateSentItemsStyle -Value 1
Get-ItemProperty $RegKeyWaste

$RegKeySent = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\General"
New-Item -Path $RegKeySent -Name DelegateWastebasketStyle -Force
New-ItemProperty -Path $RegKeySent -Name DelegateWastebasketStyle -Value 4
Get-ItemProperty $RegKeySent
