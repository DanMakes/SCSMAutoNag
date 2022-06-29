# This is an auto-nag script for SCSM incidents
# version 1.0

#Environment Constants
$SCSMInbox = "mailbox@company.com"
$SMTPServer = "mail.company.com"
$Port = 25
$PendingUserFeedbackStatus = Get-SCSMEnumeration -name "Enum.036fda4e1e164c3caa5f91abc46d86c2$"

# Relevant Classes
$irClass = get-scsmclass -name "system.workitem.incident$"
$affectedUserRelClass = get-scsmrelationshipclass -name "System.WorkItemAffectedUser$"
$assignedToUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItemAssignedToUser$"
$managersUserRelClass = Get-SCSMRelationshipClass -name "System.UserManagesUser$"
$Incidents = get-scsmobject -class $irClass -filter "status -eq $($PendingUserFeedbackStatus.id)"

# This function was borrowed from AdhocAdam and tweaked to work here 
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

$CommentToAdd = "No User Feedback Recieved"

foreach ($incident in $incidents) {
    #get the affected user's details
    $AffectedUser = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($affectedUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
    $AffectedUserEmail = Get-SCSMRelatedObject -smobject $affectedUser | Where-Object { ($_.ClassName -eq "System.Notification.Endpoint") -and ($_.DisplayName -like "*SMTP") }

    #get the assigned user's details
    $AssignedUser = Get-SCSMRelationshipObject -BySource $incident -Filter "RelationshipId -eq '$($assignedToUserRelClass.Id)'" | Where-Object { $_.IsDeleted -eq $false } | Select-Object -expandproperty TargetObject | Foreach-Object { Get-scsmobject -id $_.Id  }
    $AssignedUserEmail = Get-SCSMRelatedObject -smobject $assignedUser | Where-Object { ($_.ClassName -eq "System.Notification.Endpoint") -and ($_.DisplayName -like "*SMTP") }
    
    # Get the user/analyst action log, sort them based on last modified
    $actionLog = Get-SCSMRelatedObject -smobject $incident | Where-Object { ($_.ClassName -eq ("System.WorkItem.TroubleTicket.AnalystCommentLog") -or $_.ClassName -eq ("System.WorkItem.TroubleTicket.UserCommentLog")) } | Sort-Object LastModified -desc
    $MostRecentComment = ($actionlog[0]).comment

    # Email Properties
    $Subject = "[$($Incident.id)] - $($incident.Title)"
    $To = $AffectedUserEmail.TargetAddress
    $Cc = $AssignedUserEmail.TargetAddress

    # First Nag
    if ((((get-date) - $actionLog[0].LastModified).Days -eq 3) -and ($actionLog[0].ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $FirstNagBody = "Hi $($affectedUser.firstname), have you had a chance to look at my previous comment?<br>I'll post it here, just in case you need it:<br><br>$MostRecentComment<br><br>Please reply to this email if you are in need of further assistance.<br><br>Thanks!<br><br>$($assignedUser.FirstName) $($assignedUser.LastName)"
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc -Subject $Subject -Body $FirstNagBody -BodyAsHtml
    }
    
    # Second Nag
    if ((((get-date) - $actionLog[0].LastModified).Days -eq 4) -and ($actionLog[0].ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $SecondNagBody = "Hello $($affectedUser.firstname), I haven't heard from you since my previous attempt yesterday. I'll make another attempt tomorrow before resolving this incident.<br><br>Here is my most recent comment, sent in yesterday's email as well:<br><br>$MostRecentComment<br><br>Please reply to this email if you are in need of further assistance.<br><br>Thanks!<br><br>$($assignedUser.FirstName) $($assignedUser.LastName)"
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc -Subject $Subject -Body $SecondNagBody -BodyAsHtml
    }

    # Third Nag
    if ((((get-date) - $actionLog[0].LastModified).Days -eq 5) -and ($actionLog[0].ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        $ThirdNagBody = "Hi $($affectedUser.firstname), please let me know if you still require assistance with this incident. If there is no response by this time tomorrow, this incident will auto-resolve.<br><br>Here is my most recent comment, sent in the previous two email communications:<br><br>$MostRecentComment<br><br>Please reply to this email if you are in need of further assistance.<br><br>Thanks!<br><br>$($assignedUser.FirstName) $($assignedUser.LastName)"
        Send-MailMessage -SmtpServer $SMTPServer -Port $Port -From $SCSMInbox -To $To -Cc $Cc -Subject $Subject -Body $ThirdNagBody -BodyAsHtml
    }

    # Incident Closure
    if ((((get-date) - $actionLog[0].LastModified).Days -ge 6) -and ($actionLog[0].ClassName -eq "System.WorkItem.TroubleTicket.AnalystCommentLog")) {
        Add-ActionLogEntry -WIObject $incident -Action Resolved -Comment $CommentToAdd -EnteredBy $AssignedUser
        Set-SCSMObject -SMObject $incident -PropertyHashtable @{"ResolvedDate" = (Get-Date); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionCategory" = "Enum.66ea772bb1d74f53a4c9537cf599d67b"; "ResolutionDescription" = $CommentToAdd}
    }
}
