
function Get-AppToken{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsAutomationClientID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsAutomationClientSecret,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $tenantId

    )

    # Azure AD OAuth Application Token for Graph API
    # Get OAuth token for a AAD Application (returned as $token)

    # Application (client) ID, tenant ID and secret

    $clientId = $UniIDTeamsAutomationClientID
    $clientSecret =  $UniIDTeamsAutomationClientSecret

    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    # Get OAuth 2.0 Token
    $tokenRequestArgs = @{
        Method = "Post"
        Uri = $uri 
        ContentType = "application/x-www-form-urlencoded" 
        Body = $body 
        UseBasicParsing = $true
    }
    $tokenRequest = Invoke-WebRequest @tokenRequestArgs

    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token

    # Specify the URI to call and method
    # $uri = "https://graph.microsoft.com/v1.0/groups/"
    # $method = "POST"

    return $token
}

function Get-UserToken{
    $clientId = ""
    $tenantId = ""

    $resource = "https://graph.microsoft.com/"
    $scope = "Group.ReadWrite.All"

    $codeBody = @{ 

        resource  = $resource
        client_id = $clientId
        scope     = $scope

    }

    # Get OAuth Code
    $codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/devicecode" -Body $codeBody

    # Print Code to console
    Write-Output "`n$($codeRequest.message)"

    $tokenBody = @{

        grant_type = "urn:ietf:params:oauth:grant-type:device_code"
        code       = $codeRequest.device_code
        client_id  = $clientId

    }

    # Get OAuth Token
    while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {

        $tokenRequest = try {

            Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/token" -Body $tokenBody

        }
        catch {

            $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json

            # If not waiting for auth, throw error
            if ($errorMessage.error -ne "authorization_pending") {

                throw

            }

        }

    }

    return $tokenRequest.access_token
}

function Get-UserDelegateToken{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsMeetingClientID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsMeetingClientSecret,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $tenantId,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsMeetingUsername,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $UniIDTeamsMeetingPassword

    )
    $clientID = $UniIDTeamsMeetingClientID
    $tenantName = $tenantId
    $ClientSecret = $UniIDTeamsMeetingClientSecret
    $Username = $UniIDTeamsMeetingUsername
    $Password = $UniIDTeamsMeetingPassword
    
    
    $ReqTokenBody = @{
        Grant_Type    = "Password"
        client_Id     = $clientID
        Client_Secret = $clientSecret
        Username      = $Username
        Password      = $Password
        Scope         = "Group.ReadWrite.All"
    } 
    
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

    Return $TokenResponse.access_token
}

function Create-Offic365Group{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Token,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Description,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Year,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $DisplayName,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $MailEnabled,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $MailNickname,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $SecurityEnabled,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $OwnerID,
        [Parameter(Mandatory=$true, Position=1)]
        [string] $Visibility

    )
    $uri = "https://graph.microsoft.com/v1.0/groups/"
    $method = "POST"

<#
Values for resourceBehaviorOptions:

AllowOnlyMembersToPost
CalendarMemberReadOnly
ConnectorsEnabled
HideGroupInOutlook
NotebookForLearningCommunitiesEnabled
ReportToOriginator
SharePointReadonlyForMembers
SubscriptionEnabled
SubscribeMembersToCalendarEvents
SubscribeMembersToCalendarEventsDisabled
SubscribeNewGroupMembers
WelcomeEmailDisabled
WelcomeEmailEnabled
#>

    $body = "{
        ""description"": ""Automated Team: $Description Year: $Year"",
        ""displayName"": ""$DisplayName"",
        ""groupTypes"": [
          ""Unified""
        ],
        ""mailEnabled"":  $MailEnabled,
        ""mailNickname"": ""$MailNickname"",
        ""securityEnabled"": $SecurityEnabled,
        ""visibility"": ""$Visibility"",
        ""owners@odata.bind"": [
            ""https://graph.microsoft.com/beta/users/$OwnerID""
        ],    
        ""resourceBehaviorOptions"": [
            ""WelcomeEmailDisabled"",
            ""HideGroupInOutlook""    
            ]
        }"

    $groupResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Token"} -Body $body -ErrorAction Stop

    return $groupResult    
}

function Add-UsersToGroup{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $UserID
    )

    #add user to group
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/members/`$ref"
    $method = "POST"

    $body = "{
        ""@odata.id"": ""https://graph.microsoft.com/v1.0/users/$UserID""
      }"

    $AddMemberResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop

    return $AddMemberResult
}

function Remove-UserFromGroup{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $UserID
    )

    #add user to group
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/members/$UserID/`$ref"
    $method = "DELETE"

    $AddMemberResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop

    return $AddMemberResult
}

function Add-OwnerToGroup{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $UserID
    )

    #set Owner
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/owners/`$ref"
    $method = "POST"

    $body = "{
        ""@odata.id"": ""https://graph.microsoft.com/v1.0/users/$UserID""
      }"

    $AddOwnerResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop

    return $AddOwnerResult
}

function Create-Team{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $allowCreateUpdateChannels,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $allowUserEditMessages,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $allowUserDeleteMessages,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $allowGiphy,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $giphyContentRating
    )

    #create Team
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/team"
    $method = "PUT"

    $body = "{  
        ""memberSettings"": {
          ""allowCreateUpdateChannels"": ""$allowCreateUpdateChannels""
        },
        ""messagingSettings"": {
          ""allowUserEditMessages"": ""$allowUserEditMessages"",
          ""allowUserDeleteMessages"": ""$allowUserDeleteMessages""
        },
        ""funSettings"": {
          ""allowGiphy"": ""$allowGiphy"",
          ""giphyContentRating"": ""$giphyContentRating""
        }
      }"

    $createTeamResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
}

function Create-ChannelPrivate{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $UserID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $displayName,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $description,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $membershipType
    )

    if($membershipType -eq $null){
        $membershipType = "Null"
    }

    #create Channel
    $uri = "https://graph.microsoft.com/beta/teams/$GroupID/channels"
    $method = "POST"
 
    $body = "{
        ""@odata.type"": ""#Microsoft.Teams.Core.channel"",
        ""membershipType"": ""$membershipType"",
        ""displayName"": ""$displayName"",
        ""description"": ""$description"",
        ""members"":
            [
                {
                ""@odata.type"":""#microsoft.graph.aadUserConversationMember"",
                ""user@odata.bind"":""https://graph.microsoft.com/beta/users('$UserID')"",
                ""roles"":[""owner""]
                }
            ]
       }"
       $body
    $createChannelResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $createChannelResult
}

function Create-Channel{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $displayName,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $description
    )

    if($membershipType -eq $null){
        $membershipType = "Null"
    }

    #create Channel
    $uri = "https://graph.microsoft.com/beta/teams/$GroupID/channels"
    $method = "POST"
 
    $body = "{
        ""@odata.type"": ""#Microsoft.Teams.Core.channel"",
        ""displayName"": ""$displayName"",
        ""description"": ""$description""
       }"
       $body
    $createChannelResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $createChannelResult
}

function Create-ChannelEvent{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Subject,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $StartDateTime,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $EndDateTime,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $ThreadID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $GroupID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/beta/me/onlineMeetings"
    $method = "POST"
 
    $body = "{
        ""startDateTime"":""$StartDateTime+08:00"",
        ""endDateTime"":""$EndDateTime+08:00"",
        ""subject"":""$Subject"",
        ""chatInfo"": {
          ""threadId"":""$ThreadID""
        },
        ""participants"": {
          ""organizer"": {
            ""identity"": {
              ""user"": {
                ""id"": ""eef4ca8b-facf-41fc-acff-06a01736b9c3""
              }
            }
          }
        }
      }"
       $body
    $createChannelEventResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $createChannelEventResult
}

function Send-ChatMessage{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $ChannelID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $TeamID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $MessageID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Message
    )

    #create Channel
    $uri = "https://graph.microsoft.com/beta/teams/$TeamID/channels/$ChannelID/messages/$MessageID/replies"
    $method = "POST"
 
    $body = "{
        ""body"": {
            ""contentType"": ""html"",
            ""content"": ""$Message""
          }
      }"
       $body
    $chatResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $chatResult
}

function Create-OnlineMeeting{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Subject,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Content,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $StartDateTime,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $DisplayName
    )

    if($membershipType -eq $null){
        $membershipType = "Null"
    }

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/events"
    $method = "POST"
 
    $body = "{
        ""subject"": ""$Subjec"",
        ""body"": {
          ""contentType"": ""HTML"",
          ""content"": ""$Content""
        },
        ""start"": {
            ""dateTime"": ""$StartDateTime"",
            ""timeZone"": ""Australian Western Standard Time""
        },
        ""end"": {
            ""dateTime"": ""$StartDateTime"",
            ""timeZone"": ""Australian Western Standard Time""
        },
        ""location"":{
            ""displayName"":""$DisplayName""
        },
        ""attendees"": [
          {
            ""emailAddress"": {
              ""address"":""0123123123@email.com"",
              ""name"": ""Mason Torres""
            },
            ""type"": ""required""
          }
        ]
      }"
       $body
    $createChannelResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $createChannelResult
}

function Get-App{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/installedApps?`$expand=teamsAppDefinition"
    $method = "GET"

    $appsResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $appsResult
}

function Remove-App{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $AppID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/installedApps/$AppID"
    $method = "DELETE"

    $appsResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $appsResult

}

function Get-Channels{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/channels"
    $method = "GET"

    $channelsResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $channelsResult
}

function Get-Tabs{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Token,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $GroupID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $ChannelID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/channels/$ChannelID/tabs"
    $method = "GET"

    $tabsResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $tabsResult
}

function Remove-Tab{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Token,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $GroupID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $ChannelID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $TabID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/channels/$ChannelID/tabs/$tabID"
    $method = "DELETE"

    $tabsResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $tabsResult
}

function Add-TabWeb{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $ChannelID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $displayName,
         [Parameter(Mandatory=$false, Position=0)]
         [string] $entityId,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $contentUrl,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $websiteUrl,
         [Parameter(Mandatory=$false, Position=0)]
         [string] $removeUrl
    )

    if($removeUrl -eq $null){
        $removeUrl = "Null"
    }

    if($entityId -eq $null){
        $entityId = "Null"
    }

    $body = "{
        ""displayName"": ""$displayName"",
        ""teamsApp@odata.bind"" : ""https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web/"",
        ""configuration"": {
          ""entityId"": ""$entityId"",
          ""contentUrl"": ""$contentUrl"",
          ""websiteUrl"": ""$websiteUrl"",
          ""removeUrl"": ""$removeUrl""
        }
      }"

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/channels/$ChannelID/tabs"
    $method = "POST"

    $tabResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
    
    return $tabResult
}

function Add-TabOneNote{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $ChannelID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $displayName,
         [Parameter(Mandatory=$false, Position=0)]
         [string] $entityId,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $contentUrl,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $websiteUrl,
         [Parameter(Mandatory=$false, Position=0)]
         [string] $removeUrl
    )

    if($removeUrl -eq $null){
        $removeUrl = "Null"
    }

    if($entityId -eq $null){
        $entityId = "Null"
    }

    #TeamsAppID https://docs.microsoft.com/en-us/graph/teams-configuring-builtin-tabs#onenote-tabs
    $body = "{
        ""displayName"": ""$displayName"",
        ""teamsApp@odata.bind"" : ""https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78/"",
        ""configuration"": {
          ""entityId"": ""$entityId"",
          ""contentUrl"": ""$contentUrl"",
          ""websiteUrl"": ""$websiteUrl"",
          ""removeUrl"": ""$removeUrl""
        }
      }"

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/teams/$GroupID/channels/$ChannelID/tabs"
    $method = "POST"

    $tabResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
    
    return $tabResult
}

function Get-Onenotes{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Token,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $GroupID
    )

    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/onenote/notebooks"
    $method = "GET"

    $notebooksResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $notebooksResult
}

function Create-Notebook{
    Param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string] $Token,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $GroupID,
        [Parameter(Mandatory=$true, Position=0)]
        [string] $displayName
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/onenote/notebooks"
    $method = "POST"
 
    $body = "{
        ""displayName"": ""$displayName""
      }"
 
    $notebookResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
     
    return $notebookResult
}

function Get-User{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $UserPrincipalName
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName"
    $method = "GET"

    $userResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $userResult
}

function Get-GroupMembers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/members"
    $method = "GET"

    $GroupMembersResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    $GroupMembers = $GroupMembersResult.value

    $GroupMembersResultNextLink = $GroupMembersResult."@odata.nextLink"
    while($GroupMembersResultNextLink -ne $null){
        $GroupMembersResult = (Invoke-RestMethod -Uri $GroupMembersResultNextLink –Headers @{Authorization = "Bearer $token"} –Method Get -Verbose -ErrorAction Stop) 
        $GroupMembersResultNextLink = $GroupMembersResult."@odata.nextLink"
        $GroupMembers += $GroupMembersResult.value
    }

    return $GroupMembers
}

function Get-GroupOwners{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupID/owners"
    $method = "GET"

    $GroupMembersResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $GroupMembersResult
}

function Add-ChannelOwner{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $ChannelID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $UserID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/beta/teams/$GroupID/channels/$ChannelID/members/$UserID"
    $method = "PATCH"

    $body = "{
            ""@odata.type"": ""#microsoft.graph.aadUserConversationMember"",
            ""roles"": [""owner""]
            }"
 
    $ChannelOwnerResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -Body $body -ErrorAction Stop
    
    return $ChannelOwnerResult
}

function Get-TeamsUsers{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $Token,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $GroupID,
         [Parameter(Mandatory=$true, Position=0)]
         [string] $ChannelID
    )

    #create Channel
    $uri = "https://graph.microsoft.com/beta/teams/$GroupID/channels/$ChannelID/members"
    $method = "GET"

    $userResult = Invoke-RestMethod -Method $method -uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
    
    return $userResult
}

function Generate-ChannelName{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $SourceName,
         [Parameter(Mandatory=$true, Position=0)]
         [PSCustomObject] $Team
    )

    foreach($Schedule in $team.Schedules){
        if($SourceName -like $Schedule.Channel){
            $startTime = (get-Date $Schedule.Start_Time  -UFormat '%I:%M %p').replace(":",".")
            return "$SourceName $($Schedule.Day) $startTime"
        }
    }
    
}