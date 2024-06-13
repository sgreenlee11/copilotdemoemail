# Description: This script is used to generate a fictional email thread between a group of users. The script uses the OpenAI API to generate content for the email thread. The script uses the Microsoft Graph PowerShell module to send emails to the users involved in the thread. The script is designed to simulate a fictional email thread for demonstration purposes.
#This example is customized to run in an Azure Automation PowerShell Runbook.

# Import required modules
#Import-Module Microsoft.Graph.Authentication
#Import-Module Microsoft.Graph.Users.Actions
#Import-Module Microsoft.Graph.Mail
#Import-Module PSOpenai


# Define functions to interact with Microsoft Graph API using the Microsoft Graph PowerShell module
function Send-FirstMessage {
    [CmdletBinding()]
    param (
        [Parameter()]
        [System.String]
        $Sender,
        [Parameter()]
        $Subject,
        [Parameter()]
        $Body,
        [Parameter()]
        [Array]
        $Recipients
    )
    #Build graph message parameters
    $messageparams = @{
        subject      = $($subject)
        body         = @{
            contentType = "HTML"
            content     = $($body)
        }
        toRecipients = $recipientarray
    }
    $message = New-MgUserMessage -UserId $sender -BodyParameter $messageparams
    Send-MgUserMessage -UserId $sender -MessageId $message.Id
    return $message
}

function Send-ReplyAllMessage {

    [CmdletBinding()]
    param (
        [Parameter()]
        [System.String]
        $Sender,
        [Parameter()]
        $ConversationId,
        [Parameter()]
        $Subject,
        [Parameter()]
        $Body
    )

    #Send reply all email from next user
    $bodyparam = @{
        message = @{
            subject = "RE: " + $($Subject)
            body    = @{
                contentType = "HTML"
                content     = $($Body)
            }
        }
    }

    $msg = Get-MgUserMessage -UserId $Sender -Filter "ConversationId eq '$($ConversationId)'" -Top 100
    if ($msg.count -gt 1) {
        $msg = $msg | Where-Object { $_.Subject -notmatch "Undeliverable" } | Sort-Object -Property createdDateTime -Descending | Select-Object -First 1
    }
    $message = New-MgUserMessageReplyAll -UserId $Sender -MessageId $msg.Id -BodyParameter $bodyparam
    Send-MgUserMessage -UserId $Sender -MessageId $message.Id
    return $message
}

#Connect to Graph PowerShell with Client Secret
$tenantId = "<Tenant ID>"
$clientId = "<Client ID>"
$Certificate = Get-AutomationCertificate -Name '<certificate name>'
Connect-MgGraph -TenantId $tenantId -ClientId $clientId -Certificate $Certificate

#Connect to OpenAI API with psopenai module
$AuthType = 'azure'
Connect-AzAccount -Identity
$secret = Get-AzKeyVaultSecret -VaultName "<AKV name>" -Name "<Secret Name>" -AsPlainText
Write-Output $secret
$env:OPENAI_API_KEY = $secret
$global:OPENAI_API_BASE = '<Open API Base URI'


#Create static array with values from $users
$users = @(
    [PSCustomObject]@{
        UserPrincipalName = "GradyA@domain.com"
        DisplayName = "Grady Archie"
        Mail = "GradyA@domain.com"
        GivenName = "Grady"
        Surname = "Archie"
    },
    [PSCustomObject]@{
        UserPrincipalName = "DebraB@domain.com"
        DisplayName = "Debra Berger"
        Mail = "DebraB@domain.com"
        GivenName = "Debra"
        Surname = "Berger"
    },
    [PSCustomObject]@{
        UserPrincipalName = "MeganB@domain.com"
        DisplayName = "Megan Bowen"
        Mail = "MeganB@domain.com"
        GivenName = "Megan"
        Surname = "Bowen"
    },
    [PSCustomObject]@{
        UserPrincipalName = "ChristieC@domain.com"
        DisplayName = "Christie Cline"
        Mail = "ChristieC@domain.com"
        GivenName = "Christie"
        Surname = "Cline"
    },
    [PSCustomObject]@{
        UserPrincipalName = "AllanD@domain.com"
        DisplayName = "Allan Deyoung"
        Mail = "AllanD@domain.com"
        GivenName = "Allan"
        Surname = "Deyoung"
    },
    [PSCustomObject]@{
        UserPrincipalName = "AlexW@domain.com"
        DisplayName = "Alex Wilber"
        Mail = "AlexW@domain.com"
        GivenName = "Alex"
        Surname = "Wilber"
    },
    [PSCustomObject]@{
        UserPrincipalName = "NestorW@domain.com"
        DisplayName = "Nestor Wilke"
        Mail = "NestorW@domain.com"
        GivenName = "Nestor"
        Surname = "Wilke"
    }
)

#Get list of unique email subjects to avoid duplication. Select user from list to be used as the mailbox to search
$subjectlist = Get-MgUserMessage -UserId "<user@domain.com>" -Top 100 -Property Subject -Filter "not(startswith(subject, 'PIM')) and not(startswith(subject, 'RE:')) and not(startswith(subject, 'Test')) and not(startswith(subject, 'Undeliverable'))" | Select-Object -ExpandProperty Subject -Unique
$starterprompt = "You are an Executive at Contoso, a large beverage company. Please come up with a random fictional topic to build an email thread around. The list of potential topics should be broad and can span both business and fun morale events. For example, financials, business plans, corporate real estate, happy hours and personnel recognition. This is not a full list, just an example. Use aspects of the current time to randomize the topic. Please start by telling me the email subject line. Please only reply with the subject line and no other content. Do not include the words `"Subject Line:`" in your response. The subject should not be related to any of the following subjects: $($subjectlist -join ', ')."

$start = Get-Random -Minimum 0 -Maximum 6
$initiator = $users[$start]
$response = Request-AzureChatCompletion -Deployment "<OAI Deployment Name>" -AuthType $AuthType -Message $starterprompt

$subject = $response.Answer 
$subject = $subject -replace "\`"",""
#Send an email to all users from initiator with subject line

$response = Request-AzureChatCompletion -Deployment "<OAI Deployment Name>" -AuthType $AuthType -Message "You are an Executive at Contoso, a large beverage company. Please draft an email to send to the team with the subject line $($subject) from $($initiator.displayname). This is a fictional email thread so please be creative and do not worry about being accurate. Please only reply with the email content and no other content. Do not include the subject in the response."

#Replace new line with <p> tag for HTML formatting
$response.Answer = $response.Answer -replace "`n","<p>"

#Build array of recipients
$recipientarray = @()
foreach ($user in ($users | Where-Object { $_.mail -ne $initiator.mail })) {
    $recipientarray += @{"emailAddress" = @{"address" = $user.mail } }
}

$message1 = Send-FirstMessage -Sender $initiator.UserPrincipalName -Subject $subject -Body $response.Answer -Recipients $recipientarray

#$previousmessage = Get-MgUserMessage -UserId $initiator.UserPrincipalName -Filter "ConversationId eq '$($message1.ConversationId)'"
$conversationid = $message1.ConversationId
$previoussender = $initiator
$previousresponse = $response

sleep -Seconds 30

$replycount = 0
Do{
#Determine next user to respond
$nextuser = $users | Get-Random

#Generate next response in thread
$replyresponse = $previousresponse | Request-AzureChatCompletion -Deployment "<OAI Deployment Name>" -AuthType $AuthType -Message "You are an email Ghost Writer for a Team of Executives at Contoso, a large beverage company. Please draft a reply to the email thread with the subject line $($subject) from $($nextuser.displayname). This is a fictional email thread so please be creative and do not worry about being accurate. Please only reply with the email content and no other content. Do not include the subject in the response."

#Replace new line with <p> tag for HTML formatting
$replyresponse.Answer = $replyresponse.Answer -replace "`n","<p>"

$replymessage = Send-ReplyAllMessage -Sender $nextuser.UserPrincipalName -ConversationId $conversationid -Subject $subject -Body $replyresponse.Answer

#Wait 30 seconds before next reply to allow for delivery
sleep -Seconds 30

$previoussender = $nextuser
$previousresponse = $replyresponse
$conversationid = $replymessage.ConversationId
#Increment Reply Count
$replycount++
} while ($replycount -lt 5)
