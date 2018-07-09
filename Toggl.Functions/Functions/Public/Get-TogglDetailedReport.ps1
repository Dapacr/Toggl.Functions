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

        .Link
        https://github.com/Dapacruz/VMware.VimAutomation.Custom
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
        [string]$UserAgent = $(if (Test-Path $PSScriptRoot\user_agent) { Get-Content $PSScriptRoot\user_agent } else { Set-Content -Value (Read-Host -Prompt 'Email Address') -Path $PSScriptRoot\user_agent -PassThru | Out-String }),
        [string]$User = $(if (Test-Path $PSScriptRoot\api_token) { Get-Content $PSScriptRoot\api_token } else { $User = Set-Content -Value (Read-Host -Prompt 'API Token') -Path $PSScriptRoot\api_token -PassThru | Out-String }),
        [string]$WorkspaceID = $(if (Test-Path $PSScriptRoot\workspace_id) { Get-Content $PSScriptRoot\workspace_id } else { Set-Content -Value (Read-Host -Prompt 'Workspace ID') -Path $PSScriptRoot\workspace_id -PassThru | Out-String })
    )

    $pass = "api_token"
    $pair = "$($User):$($pass)"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $base64 = [System.Convert]::ToBase64String($bytes)
    $basic_auth_value = "Basic $base64"
    $headers = @{Authorization=$basic_auth_value}
    $content_type = "application/json"
    $since = $From.ToString('yyyy-MM-dd')
    $until = $To.ToString('yyyy-MM-dd')
    
    # Query Toggl API for report details
    $uri_report = "https://toggl.com/reports/api/v2/details?since=$since&until=$until&display_hours=decimal&rounding=on&user_agent=$UserAgent&workspace_id=$WorkspaceID"
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
        $report = $report.Where{$_.Client -match $Client}
    }
    if ($Ticket) {
        $report = $report.Where{$_.Ticket -match $Ticket}
    }
    if ($Project) {
        $report = $report.Where{$_.Project -match $Project}
    }
    if ($Description) {
        $report = $report.Where{$_.Description -match $Description}
    }
    if ($WorkType) {
        $report = $report.Where{$_.WorkType -match $WorkType}
    }

    # Insert custom TypeName (defined in $PSScriptRoot\format.ps1xml) to control default display
    foreach ($obj in $report) {
        $obj.PSTypeNames.Insert(0,'Toggl.Report.Detailed')
    }
        
    Write-Output $report
}
