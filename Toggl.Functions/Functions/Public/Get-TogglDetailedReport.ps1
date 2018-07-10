<#
        .Synopsis
        Returns report data for the specified date range.
        .Description
        This cmdlet uses Toggl's tag functionality to classify time entries.
        The following tags are used:
            Billable
            Holiday
            Non-Billable
            PTO
            Training
            Utilized
        .Example
        Get-TogglDetailedReport -From 1/1 -Client Acme
        Generate a detailed report, for client Acme, from 1/1 of the current year to today
        .Example
        Get-TogglDetailedReport -From 1/1 -To 3/1 -Ticket 232641
        Generate a detailed report, for ticket # 232641, from 1/1 of the current year to 3/1
        .Example
        Get-TogglDetailedReport -From 1/1 -To (Get-Date 1/1).AddMonths(1) -WorkType Billable | measure 'Duration(hrs)' -Sum
        Calculate the number of billable hours for a particular period
        .Example
        Get-TogglDetailedReport -Ticket 232641 -WorkType Billable | measure 'Duration(hrs)' -Sum
        Calculate the number of billable hours for a particular project
        .Link
        https://github.com/Dapacruz/Toggl.Functions
#>
function Get-TogglDetailedReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)][Alias('Since')]
        [datetime]$From = (Get-Date).AddMonths(-12),
        [Parameter(Position=1)][Alias('Until')]
        [datetime]$To = (Get-Date),
        [string]$Client,
        [string]$Ticket,
        [string]$Project,
        [string]$Description,
        [string]$WorkType,
        [string]$User,
        [string]$ApiToken,
        [string]$WorkspaceID
    )
    Begin {
        $toggl_dir = "$HOME\.toggl.functions\"
        if (-not (Test-Path $toggl_dir)) {
            New-Item -Path $toggl_dir -ItemType Directory | Out-Null
        }
        
        $user_agent = "$toggl_dir\user_agent"
        if (-not $User) {
            if (Test-Path $user_agent) {
                $User = Get-Content $user_agent
            } else {
                $User = Read-Host -Prompt 'User Name (email address)'
                Set-Content -Value $User -Path $user_agent
            }
        }
        
        $api_token = "$toggl_dir\api_token"
        if (-not $ApiToken) {
            if (Test-Path $api_token) {
                $ApiToken = Get-Content $api_token
            } else {
                $ApiToken = Read-Host -Prompt 'API Token'
                Set-Content -Value $ApiToken -Path $api_token
            }
        }
        
        $workspace_id = "$toggl_dir\workspace_id"
        if (-not $WorkspaceID) {
            if (Test-Path $workspace_id) {
                $WorkspaceID = Get-Content $workspace_id
            } else {
                $WorkspaceID = Read-Host -Prompt 'Workspace ID'
                Set-Content -Value $WorkspaceID -Path $workspace_id
            }
        }
    }
    Process {
        $pass = "api_token"
        $pair = "$($ApiToken):$($pass)"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $base64 = [System.Convert]::ToBase64String($bytes)
        $basic_auth_value = "Basic $base64"
        $headers = @{Authorization=$basic_auth_value}
        $content_type = "application/json"
        $since = $From.ToString('yyyy-MM-dd')
        $until = $To.ToString('yyyy-MM-dd')
        
        # Query Toggl API for report details
        $uri_report = "https://toggl.com/reports/api/v2/details?since=$since&until=$until&display_hours=decimal&rounding=on&user_agent=$User&workspace_id=$WorkspaceID"
        $toggl_response = Invoke-RestMethod -Uri $uri_report -Headers $headers -ContentType $content_type
        $response_total = $toggl_response.total_count
        $page_num = 1
        $report = @()
        
        while ($response_total -gt 0) {
            $toggl_response = Invoke-RestMethod -Uri ($uri_report + '&page=' + $page_num) -Headers $headers -ContentType $content_type
            $TogglResponseData = $toggl_response.data
            $report += $TogglResponseData
            $response_total = $response_total - $toggl_response.per_page

            $page_num++
        }
        
        $report = $report | Select-Object @{n='Date';e={Get-Date $_.start -Format yyyy-MM-dd}},
                                        @{n='Client';e={$_.client}},
                                        @{n='Ticket';e={($_.project -replace '(.*?\b(\d+)\b.*|.*)','$2') -replace '$^','n/a'}},
                                        @{n='Description';e={$_.description}},
                                        @{n='Project';e={$_.project -replace '^Ticket\s?#\s?\d+\s\((.+)\)','$1'}},
                                        @{n='Duration(hrs)';e={'{0:n2}' -f ($_.dur/1000/60/60)}},
                                        @{n='WorkType';e={$_.tags -as [string]}} | Sort-Object Date
        
        if ($Client) {
            $report = $report.Where{$_.Client -like $Client}
        }
        if ($Ticket) {
            $report = $report.Where{$_.Ticket -like $Ticket}
        }
        if ($Project) {
            $report = $report.Where{$_.Project -like $Project}
        }
        if ($Description) {
            $report = $report.Where{$_.Description -like $Description}
        }
        if ($WorkType) {
            $report = $report.Where{$_.WorkType -like $WorkType}
        }

        # Insert custom TypeName (defined in $PSScriptRoot\format.ps1xml) to control default display
        foreach ($obj in $report) {
            $obj.PSTypeNames.Insert(0,'Toggl.Report.Detailed')
        }
            
        Write-Output $report
    }
    End {}
}
