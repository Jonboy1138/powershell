#############################################################################
#
# Check-HPWarranty.ps1
# Version: 1.0
# Author: Jon Bennett
# Date: 7/18/2018
#
# This script uses the Selenium webdriver to check the serial numbers on the
# HP Warranty Check website. The serial numbers are passed in via an Excel
# spreadsheet.
#
# *Set-ExecutionPolicy Unrestricted must be run prior to running script*
#
#############################################################################

# Needed to open new tabs when opening from Powershell not ISE
Add-Type -AssemblyName System.Windows.Forms

# Selenium setup
$seleniumOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$seleniumOptions.AddArgument(@('--start-maximized', '--disable-infobars', '-enable-automation', "--lang=$language"))
$seleniumOptions.AddExtension('C:\Selenium\BrowserProfile\5.14.0_0.crx')
$seleniumOptions.AddUserProfilePreference("credentials_enable_service", $false)
$seleniumOptions.AddUserProfilePreference("profile.password_manager_enabled", $false)
$seleniumDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($seleniumOptions)

$computers = Import-Excel 'H:\powershell\Check-HPWarranty\computers.xlsx'
$i = 0

$computers | ForEach-Object {
  While ($i -lt $computers.Count) {
    $serial = $computers.SerialNumber[$i]
    $seleniumDriver.Url = 'https://support.hp.com/us-en/checkwarranty'
    $seleniumDriver.FindElementByXPath('//*[@id="wFormSerialNumber"]').SendKeys("$serial")
    $i = $i + 1
    $seleniumDriver.FindElementByXPath('//*[@id="btnWFormSubmit"]').Click()
    $seleniumDriver.FindElementByCssSelector("body").SendKeys([System.Windows.Forms.SendKeys]::SendWait("^t"))
    $tab = $seleniumDriver.WindowHandles[$i]
    $seleniumDriver.SwitchTo().Window($tab)
  }#END While ($i -le $computers.Count)
}#END ForEach-Object