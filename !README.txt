
====================================================================================
REMOVING PREVIOUS APP VERSION DURING MSI INSTALL
====================================================================================

In order to automatically uninstall the previous version of any app using an MSI, 
the value in [Setup Project > Properties > UpgradeCode] for the MSI must be 
identical to the UpgradeCode of the currently installed application. 

In order to obtain the UpgradeCode of an already installed application, use 
the following Powershell: 

	$products = Get-CimInstance -Class Win32_product
	$properties = Get-CimInstance -Query `
		"SELECT ProductCode,Value FROM Win32_Property WHERE Property='UpgradeCode'"
	$myApp = $products | where-object {$_.name -match '^<MY_APP_NAME>$'}
	$upgradeCodeProperty = $properties | Where-Object {$_.ProductCode -eq $myApp.IdentifyingNumber}
	$upgradeCode.Value

Make sure the UpgradeCode in your MSI Setup project is identical to that value 
and the MSI should uninstall the prior version of the app before installing the 
new one. 

Currently [2022.11.11] the UpgradeCode for BMS is: {4851483F-B07B-4C84-A71A-57A5CF7BE695}

====================================================================================


