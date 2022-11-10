# auto nag autonag
# This is an auto-nag script for SCSM incidents
# version 1.0.2

#region DEFINITIONS
#Environment Constants
$SCSMInbox = "scsminbox@company.com"
$SMTPServer = "smtp.company.com"
$Port = 25
$PendingUserFeedbackStatus = Get-SCSMEnumeration -name "Enum.036fda4e1e164c3caa5f91abc46d86c2$"

# Relevant Classes
$irClass = get-scsmclass -name "system.workitem.incident$"
$affectedUserRelClass = get-scsmrelationshipclass -name "System.WorkItemAffectedUser$"
$assignedToUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItemAssignedToUser$"
$managersUserRelClass = Get-SCSMRelationshipClass -name "System.UserManagesUser$"
$UserManagesUserRelClass = get-scsmrelationshipclass "System.UserManagesUser$"
$Incidents = get-scsmobject -class $irClass -filter "status -eq $($PendingUserFeedbackStatus.id)"
$CommentToAdd = "No User Feedback Recieved"
$ClosureDayValue = "90"
$AdminEmailList = @()
# Get-SCSMEnumeration IncidentTierQueuesEnum$ | Get-SCSMChildEnumeration # List all incident support groups, in case there are more to add later
#endregion DEFINITIONS

#region ADD ACTION LOG FUNCTION
# This function was borrowed from AdhocAdam and tweaked to work here 
# https://github.com/AdhocAdam/smletsexchangeconnector

function Add-ActionLogEntry {
    param (
        [parameter(Mandatory=$true, Position=0)]
        $WIObject,
        [parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Assign","AnalystComment","Closed","Escalated","EmailSent","EndUserComment","FileAttached","FileDeleted","Reactivate","Resolved","TemplateApplied")]
        [string] $Action,
        [parameter(Mandatory=$false, Position=2)]
        [string] $Comment,
        [parameter(Mandatory=$true, Position=3)]
        [string] $EnteredBy,
        [parameter(Mandatory=$false, Position=4)]
        [Nullable[boolean]] $IsPrivate = $false
    )

    #Choose the Action Log Entry to be created. Depending on the Action Log being used, the $propDescriptionComment Property could be either Comment or Description.
    switch ($Action)
    {
        Assign {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordAssigned"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        AnalystComment {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $propDescriptionComment = "Comment"}
        Closed {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordClosed"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        Escalated {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordEscalated"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        EmailSent {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.EmailSent"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        EndUserComment {$CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"; $propDescriptionComment = "Comment"}
        FileAttached {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileAttached"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        FileDeleted {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileDeleted"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        Reactivate {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordReopened"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        Resolved {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordResolved"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
        TemplateApplied {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.TemplateApplied"; $ActionEnum = get-scsmenumeration $ActionType ; $propDescriptionComment = "Description"}
    }

    #Alias on Type Projection for Service Requests and Problem and are singular, whereas Incident and Change Request are plural. Update $CommentClassName
    if (($WIObject.ClassName -eq "System.WorkItem.Problem") -or ($WIObject.ClassName -eq "System.WorkItem.ServiceRequest")) {$CommentClassName = "ActionLog"} else {$CommentClassName = "ActionLogs"}

    #Analyst and End User Comments Classes have different Names based on the Work Item class
    if ($Action -eq "AnalystComment")
    {
        switch ($WIObject.ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "AnalystComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "AnalystCommentLog"}
            "System.WorkItem.Problem" {$CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "AnalystComments"}
        }
    }
    if ($Action -eq "EndUserComment")
    {
        switch ($WIObject.ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "UserComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "EndUserCommentLog"}
            "System.WorkItem.Problem" {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "UserComments"}
        }
    }

    # Generate a new GUID for the entry
    $NewGUID = ([guid]::NewGuid()).ToString()

    # Create the object projection with properties
    $Projection = @{__CLASS = "$($WIObject.ClassName)";
                    __SEED = $WIObject;
                    $CommentClassName = @{__CLASS = $CommentClass;
                                        __OBJECT = @{Id = $NewGUID;
                                            DisplayName = $NewGUID;
                                            ActionType = $ActionType;
                                            $propDescriptionComment = $Comment;
                                            Title = "$($ActionEnum.DisplayName)";
                                            EnteredBy  = $EnteredBy;
                                            EnteredDate = (Get-Date).ToUniversalTime();
                                            IsPrivate = $IsPrivate;
                                        }
                    }
    }

    #create the projection based on the work item class
    switch ($WIObject.ClassName)
    {
        "System.WorkItem.Incident" {New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection$" -Projection $Projection }
        "System.WorkItem.ServiceRequest" {New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection$" -Projection $Projection }
        "System.WorkItem.Problem" {New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType$" -Projection $Projection }
        "System.WorkItem.ChangeRequest" {New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel$" -Projection $Projection }
    }
}
#endregion ADD ACTION LOG FUNCTION

foreach ($incident in $incidents) {
    #get the affected user's details
    $AffectedUser = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
    $AffectedUserEmail = Get-SCSMRelatedObject -smobject $affectedUser | Where-Object { ($_.ClassName -eq "System.Notification.Endpoint") -and ($_.DisplayName -like "*SMTP") }

    #get the assigned user's details
    $AssignedUser = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
    $AssignedUserEmail = Get-SCSMRelatedObject -smobject $assignedUser | Where-Object { ($_.ClassName -eq "System.Notification.Endpoint") -and ($_.DisplayName -like "*SMTP") }
    
    # Get the user/analyst action log, sort them based on last modified
    $MostRecentActionLogEntry = Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc | Select-Object -First 1
    $MostRecentComment = $MostRecentActionLogEntry.Comment
    $DayDiff = ((get-date) - ($MostRecentActionLogEntry.LastModified)).Days

    # Email Properties
    $Subject = "Following up on [$($Incident.id)] - $($incident.Title)"
    $To = $AffectedUserEmail.TargetAddress
    $Cc1 = $AssignedUserEmail.TargetAddress
    $Cc2 = @("$($AssignedUserEmail.TargetAddress); anyotherrecipients@company.com") # Assigned User and Others

    # First Nag at the 5th day
    if ((((get-date) - $MostRecentActionLogEntry.LastModified).Days -eq 5) -and ($MostRecentActionLogEntry.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $NagBody = @"
        <html>
            <body>
                <table width=`"100%`" border=`"0`">
                    <tr style=`"height: 100px;`">
                        <td style=`"background-color: #000000;width: 100%;`" align=`"center`">
                            <img src=`"AddYourImageHere.png`" />
                        </td>
                    </tr>
                    <tr>
                        <td style=`"background-color: lightgrey; font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;`">
                            <br><h2 style=`"text-align: center;`">An analyst is waiting for your response on $($Incident.id).</h2>
                            <p style=`"text-align: center;`"><b>
                            To view this work item in the portal, click <a href=`"https://SCSM.company.com/CustomSpace/EditWorkItem.html?id=$($Incident.id)`">here</a></b></p><br>
                            <br><br>
                            Hi $($affectedUser.firstname), have you had a chance to look at my previous comment? I will post it here, just in case you need it.<br>
                            Please reply to this email if you are in need of further assistance.<br>
                            <br>
                            Thanks!<br>
                            $($assignedUser.FirstName) $($assignedUser.LastName)<br>
                            <br>
                            <u><b>Most Recent Comment:</b></u><br>
                            $MostRecentComment<br>
                            <br>
                            To speak to an analyst, please call the appropriate service desk and reference ticket $($Incident.id):
                            <br><br>
                            Service Desk 1: 555-555-5555<br>
                            Service Desk 2: 555-555-5555<br><br>
                            <br>                  
                        </td>
                    </tr>
                </table>
            </body>
        </html>
"@
        [string]$NagBody = ($NagBody -split '(</html>)')[0..1]
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc1 -Subject $Subject -Body $NagBody -BodyAsHtml

        # Add to Admin Report Email
        $AdminEmailList += [PSCustomObject]@{
            "ID" = $Incident.ID
            "Affected User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Assigned User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Support Group" = $Incident.TierQueue.DisplayName
            "Last Comment Date" = (Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc | Select-Object -First 1).LastModified
        }
    }
    
    # Second Nag at the 7th day
    if ((((get-date) - $MostRecentActionLogEntry.LastModified).Days -eq 7) -and ($MostRecentActionLogEntry.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $NagBody = @"
        <html>
            <body>
                <table width=`"100%`" border=`"0`">
                    <tr style=`"height: 100px;`">
                        <td style=`"background-color: #000000;width: 100%;`" align=`"center`">
                            <img src=`"AddYourImageHere.png`" />
                        </td>
                    </tr>
                    <tr>
                        <td style=`"background-color: lightgrey; font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;`">
                            <br><h2 style=`"text-align: center;`">An analyst is waiting for your response on $($Incident.id).</h2>
                            <p style=`"text-align: center;`"><b>
                            To view this work item in the portal, click <a href=`"https://SCSM.company.com/CustomSpace/EditWorkItem.html?id=$($Incident.id)`">here</a></b></p><br>
                            <br><br>
                            Hello $($affectedUser.firstname), I haven't heard from you since my previous attempt. I am including my most recent comment, sent in the previous email as well. Please reply to this email if you are in need of further assistance.<br>
                            <br>
                            Thanks!<br>
                            $($assignedUser.FirstName) $($assignedUser.LastName)<br>
                            <br>
                            <u><b>Most Recent Comment:</b></u><br>
                            $MostRecentComment<br>
                            <br>
                            To speak to an analyst, please call the appropriate service desk and reference ticket $($Incident.id):
                            <br><br>
                            Service Desk 1: 555-555-5555<br>
                            Service Desk 2: 555-555-5555<br><br>
                            <br>                        
                        </td>
                    </tr>
                </table>
            </body>
        </html>
"@
        [string]$NagBody = ($NagBody -split '(</html>)')[0..1]
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc1 -Subject $Subject -Body $NagBody -BodyAsHtml

        # Add to Admin Report Email
        $AdminEmailList += [PSCustomObject]@{
            "ID" = $Incident.ID
            "Affected User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Assigned User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Support Group" = $Incident.TierQueue.DisplayName
            "Last Comment Date" = (Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc | Select-Object -First 1).LastModified
        }
    }

    # Third Nag at the 10th day
    if ((((get-date) - $MostRecentActionLogEntry.LastModified).Days -eq 10) -and ($MostRecentActionLogEntry.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $NagBody = @"
        <html>
            <body>
                <table width=`"100%`" border=`"0`">
                    <tr style=`"height: 100px;`">
                        <td style=`"background-color: #000000;width: 100%;`" align=`"center`">
                            <img src=`"AddYourImageHere.png`" />
                        </td>
                    </tr>
                    <tr>
                        <td style=`"background-color: lightgrey; font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;`">
                            <br><h2 style=`"text-align: center;`">An analyst is waiting for your response on $($Incident.id).</h2>
                            <p style=`"text-align: center;`"><b>
                            To view this work item in the portal, click <a href=`"https://SCSM.company.com/CustomSpace/EditWorkItem.html?id=$($Incident.id)`">here</a></b></p><br>
                            <br><br>
                            Hi $($affectedUser.firstname), I'm just checking again to see if there's anything else you need done on this ticket. If there is no response by this time tomorrow, this notification will enter an automated state and will ask for a response once a day until the $ClosureDayValue day mark, at which time it will automatically resolve and cease notifications.<br><br>Please reply to this email if you are in need of further assistance.<br>
                            <br>
                            Thanks!<br>
                            $($assignedUser.FirstName) $($assignedUser.LastName)<br>
                            <br>
                            <u><b>Most Recent Comment:</b></u><br>
                            $MostRecentComment<br>
                            <br>
                            To speak to an analyst, please call the appropriate service desk and reference ticket $($Incident.id):
                            <br><br>
                            Service Desk 1: 555-555-5555<br>
                            Service Desk 2: 555-555-5555<br><br>
                            <br>                    
                        </td>
                    </tr>
                </table>
            </body>
        </html>
"@
        [string]$FirstNagBody = ($FirstNagBody -split '(</html>)')[0..1]
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc2 -Subject $Subject -Body $FirstNagBody -BodyAsHtml

        # Add to Admin Report Email
        $AdminEmailList += [PSCustomObject]@{
            "ID" = $Incident.ID
            "Affected User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Assigned User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Support Group" = $Incident.TierQueue.DisplayName
            "Last Comment Date" = (Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc | Select-Object -First 1).LastModified
        }
    }

    # Fourth Nag at the 11th day and onward until day 89
    if ((((get-date) - $MostRecentActionLogEntry.LastModified).Days -ge 11) -and (((get-date) - $MostRecentActionLogEntry.LastModified).Days -le 89) -and ($MostRecentActionLogEntry.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $NagBody = @"
        <html>
            <body>
                <table width=`"100%`" border=`"0`">
                    <tr style=`"height: 100px;`">
                        <td style=`"background-color: #000000;width: 100%;`" align=`"center`">
                            <img src=`"AddYourImageHere.png`" />
                        </td>
                    </tr>
                    <tr>
                        <td style=`"background-color: lightgrey; font-family: 'Gill Sans', 'Gill Sans MT', Calibri, 'Trebuchet MS', sans-serif;`">
                            <br><h2 style=`"text-align: center;`">An analyst is waiting for your response on $($Incident.id).</h2>
                            <p style=`"text-align: center;`"><b>
                            To view this work item in the portal, click <a href=`"https://SCSM.company.com/CustomSpace/EditWorkItem.html?id=$($Incident.id)`">here</a></b></p><br>
                            <br><br>
                            The analyst for this ticket has been awaiting additional information for $DayDiff days. At $ClosureDayValue days, this ticket will automatically resolve.<br>Please reply to this email to continue keeping this ticket open, or if you would like the analyst to resolve it. Thank you.<br>
                            <br>
                            <u><b>Most Recent Comment:</b></u><br>
                            $MostRecentComment<br>
                            <br>
                            To speak to an analyst, please call the appropriate service desk and reference ticket $($Incident.id):
                            <br><br>
                            Service Desk 1: 555-555-5555<br>
                            Service Desk 2: 555-555-5555<br><br>
                            <br>                   
                        </td>
                    </tr>
                </table>
            </body>
        </html>
"@
        [string]$NagBody = ($NagBody -split '(</html>)')[0..1]
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc2 -Subject $Subject -Body $NagBody -BodyAsHtml

        # Add to Admin Report Email
        $AdminEmailList += [PSCustomObject]@{
            "ID" = $Incident.ID
            "Affected User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Assigned User" = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
            "Support Group" = $Incident.TierQueue.DisplayName
            "Last Comment Date" = (Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc | Select-Object -First 1).LastModified
        }
    }

    # Incident Closure
    if ((((get-date) - $MostRecentActionLogEntry.LastModified).Days -ge $ClosureDayValue) -and ($MostRecentActionLogEntry.ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        Add-ActionLogEntry -WIObject $incident -Action Resolved -Comment $CommentToAdd -EnteredBy $AssignedUser
        Set-SCSMObject -SMObject $incident -PropertyHashtable @{"ResolvedDate" = (Get-Date); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionCategory" = "Enum.66ea772bb1d74f53a4c9537cf599d67b"; "ResolutionDescription" = $CommentToAdd}
    }        
}

#region ADMIN REPORT
$Style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$Style = $Style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$Style = $Style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$Style = $Style + "TD{border: 1px solid black; padding: 5px; }"
$Style = $Style + "</style>"

$Subject = "SCSM Auto Nag Report"
$To = "admin@company.com"
$Intro = "<P>The following work items have met the Auto Nag criteria and have been sent an email</P>"
$Body = ($AdminEmailList | ConvertTo-Html -Head $Style -PreContent $Intro | Out-String)
Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Subject $Subject -Body $Body -BodyAsHtml
#endregion ADMIN REPORT
