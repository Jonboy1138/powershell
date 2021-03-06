Import-Csv -Path C:\test\Employees.csv | foreach {
  $password = [System.Web.Security.Membership]::GeneratePassword((Get-Random -Minimum 20 -Maximum 32), 3)
  $secPw = ConvertTo-SecureString -String $password -AsPlainText -Force

  $userName = '{0}{1}' -f $_.FirstName.Substring(0,1),$_.LastName
  $NewUserParameters = @{
    GivenName = $_.FirstName
    Surname = $_.LastName
    Name = $userName
    AccountPassword = $secPw
  }
  New-ADUser @NewUserParameters
  Add-ADGroupMember -Identity 'Accounting' -Members $userName
}