#############################################
#AD
#############################################

#Import AD moudle
Get-Module ActiveDirectory | Import-Module

#Define connection variables
$ADuserName            = "domain\admin"
$ADsecurePassword      = "Password" | ConvertTo-SecureString -AsPlainText -Force
$ADCredentials         = New-Object System.Management.Automation.PSCredential($ADuserName, $ADsecurePassword)
$ADOutputFile          = "AD_LastLogonDate.csv"

#Define retrived date range (older than 60 days)
$lastdate              = (Get-Date).AddDays(-60)

#Retrive domains list in the forest
$Domains               = (Get-ADForest).Domains

    Foreach ($domain in $Domains)
    {
        #Get AD Users Criteria: (Get all users from all domains in the forest) 
        $ADobjUsers = get-aduser -filter * -properties * -Credential $AdCredentials -Server $domain | `
        Select Name, UserPrincipalName, lastlogondate, Enabled | `
        where {($_.lastlogondate -eq $null) -or ($_.lastlogondate -le $lastdate) -AND ($_.enabled -eq $true)}
        
        Foreach ($adobjUser in $adobjUsers)
        {
            #Prepare UPN & Name variables
            $strUserName = $adobjUser.Name
            $strUserPrincipalName = $adobjUser.UserPrincipalName

            #Check if they have a last logon time.
        if ($ADobjUser.LastLogondate -eq $null) 
        {
            #Never logged in, update Last Logon Variable
            $strLastLogonTime = "Never Logged In"
        }
        else 
        {
            #Update last logon variable with data
            $strLastLogonTime = $ADobjUser.LastLogonDate
        }
            
        #Prepare the user details in CSV format for writing to file
        $strUserDetails = "$strUserName,$strUserPrincipalName,$strLastLogonTime"
            
        #Append the data to file
        Out-File -FilePath $ADOutputFile -InputObject $strUserDetails -Encoding UTF8 -append
        }
    }

#############################################
#Exchange Online
#############################################
$EXuserName            = "admin@domain.com"
$EXsecurePassword      = "admin Password" | ConvertTo-SecureString -AsPlainText -Force
$EOLOutputFile         = "EOL_LastLogonDate.csv"

    #Remove all existing PS sessions
    Get-PSSession | Remove-PSSession
        
    #O365 creds
    $Office365Credentials  = New-Object System.Management.Automation.PSCredential ($EXuserName, $EXsecurePassword)
    
    #Create remote Powershell session
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri `
    "https://outlook.office365.com/powershell-liveid/" -Credential $Office365credentials -Authentication Basic –AllowRedirection
    
    Import-PSSession $Session -CommandName Get-Mailbox,Get-MailboxStatistics -FormatTypeName * -AllowClobber
    
    #Gather mailboxes from O365
    $objUsers = get-mailbox -ResultSize 50 | Where {($_.DisplayName -notlike "Discovery Search Mailbox")} | ` 
    Get-MailboxStatistics | Where {($_.lastlogonTime -eq $null) -or ($_.lastlogonTime -le $lastdate)} `
    | select lastlogontime, Displayname, UserPrincipalName

    Foreach ($objUser in $objUsers)
    {   
        #Prepare UPN variable
        $strDispalyName         = $objUser.DisplayName
        $strUserPrincipalName   = $objUser.UserPrincipalName
                
        #Check if they have a last logon time.
        if ($objUser.LastLogonTime -eq $null)
        {
        #Never logged in, update Last Logon Variable
        $strLastLogonTime = "Never Logged In"
        }
        else
        {
        #Update last logon variable with data from Office 365
        $strLastLogonTime = $objUser.LastLogonTime
        }
        
        #Prepare the user details in CSV format for writing to file
        $strUserDetails = "$strDispalyName,$strUserPrincipalName,$strLastLogonTime"
        
        #Append the data to file
        Out-File -FilePath $EOLOutputFile -InputObject $strUserDetails -Encoding UTF8 -append
    }
    #Clean up session
    Get-PSSession | Remove-PSSession

#############################################
#Send Email
#############################################
$userName            = "email@domain.com"
$securePassword      = "email password" | ConvertTo-SecureString -AsPlainText -Force
$EmailCredentials    = New-Object System.Management.Automation.PSCredential($userName, $securePassword)
$EmailSubject        = "Stale Accounts Report - $(Get-Date -Format g)"

$mailParams = @{
    SmtpServer                 = "smtp.office365.com"
    Port                       = "587"
    UseSSL                     = $true
    Credential                 = $EmailCredentials
    From                       = "email@domain.com"
    To                         = "email@domain.com"
    Subject                    = $EmailSubject
    attachments                = $EOLoutputFile, $ADOutputFile
    body                       = "Report for accounts that have not been accessed for the past 60 days on On-prem AD & O365"
    DeliveryNotificationOption = 'OnFailure'
}
Send-MailMessage @mailParams   
