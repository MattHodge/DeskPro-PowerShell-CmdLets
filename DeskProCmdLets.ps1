<#
.Synopsis
   Send Tickets to DeskPro HelpDesk
.DESCRIPTION
   Allows you to send tickets to a hosted DeskPro HelpDesk.
.EXAMPLE
   Add-DeskProTicket -APIKey 2:SDDDFGDF88901JF81 -Message "Test message from PowerShell" -SubDomain mySubdomain -Subject "Test subject from PowerShell" -DepartmentID 5

   Creates a new ticket in DepartmentID5.
.NOTES
       NAME:      Get-PagerDutyEvent
       AUTHOR:    Matthew Hodgkins
       WEBSITE:   http://www.hodgkins.net.au
       WEBSITE:   https://github.com/MattHodge
#>

function Add-DeskProTicket
{
    [CmdletBinding(ConfirmImpact='Low')]
    Param
    (
        # The DeskPro API Key
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$APIKey,

        # Subject of the ticket
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        # First message of the ticket
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        # DeskPro Subdomain
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$SubDomain,

        # Email Address To Create The Ticket As From
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$PersonEmail,

        # Ticket Label
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Label,

        # Department the ticket is in. If not specified, uses the default ticket department.
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [int]$DepartmentID,

        # The message is considered to be written by the API agent rather than the ticket owner.
        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$MessageAsAgent
    )

    $body = @{
                subject = $Subject
                message = $Message
                person_email = $PersonEmail
            }

    if ($Label -ne $null)
    {
        Add-Member -NotePropertyName label -NotePropertyValue $Label -InputObject $body
    }

    if ($DepartmentID -ne $null)
    {
        Add-Member -NotePropertyName department_id -NotePropertyValue $DepartmentID -InputObject $body
    }

    if ($MessageAsAgent)
    {
        Add-Member -NotePropertyName message_as_agent -NotePropertyValue 1 -InputObject $body
    }



    try
    {  
        $body = $body | ConvertTo-Json
        $result = Invoke-RestMethod -Uri ('https://' + $SubDomain + '.deskpro.com/api/tickets') -Headers @{"X-DeskPRO-API-Key"=$APIKey} -method Post -Body $body -ContentType "application/json"
        Write-Output $result
    }

    catch
    {
        $failureMessage = $_.Exception.Message
    }

} # End Function

