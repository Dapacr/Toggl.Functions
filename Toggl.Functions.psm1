<#
#>
function Get-BillableTimeReport {
  [CmdletBinding()] 
  Param (
    [Parameter(Position=0)][Alias('Day')]
    $Date
  )

  $UserAgent = "dcruz@dsatechnologies.com"
  $WorkspaceID = "789619"
  $user = "982b4538c4ac97feff249d0c54463164" # <-- enter your API token here
  $pass = "api_token"
  $pair = "$($user):$($pass)"
  $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
  $base64 = [System.Convert]::ToBase64String($bytes)
  $basicAuthValue = "Basic $base64"
  $headers = @{ Authorization = $basicAuthValue }
  $contentType = "application/json"
  $BillableHours = 0
  $OvertimeHours = 0
  $TotalHours = 0

  if ($Date) {
    $Date = $(Get-Date -Date $Date)
  } else {
    $Date = Get-Date
  }
  $CurrentDay = $Date.Day
  $CurrentMonth = $Date.Month
  $CurrentYear = $Date.Year

  if ($CurrentDay -ge 1 -and $CurrentDay -le 15) {
    $PayPeriodStartDay = 1
    $Since = "$CurrentYear-$CurrentMonth-01"
    $Until = "$CurrentYear-$CurrentMonth-$CurrentDay"
    $NumberOfWorkDays = 0

    for ($x = $PayPeriodStartDay; $x -le $CurrentDay; $x += 1) {
      switch ((get-date $CurrentYear-$CurrentMonth-$x).DayOfWeek) {
        'Monday' { $NumberOfWorkDays += 1 }
        'Tuesday' { $NumberOfWorkDays += 1 }
        'Wednesday' { $NumberOfWorkDays += 1 }
        'Thursday' { $NumberOfWorkDays += 1 }
        'Friday' { $NumberOfWorkDays += 1 }
      }
    }
    $WorkingHours = $NumberOfWorkDays * 8
  } elseif ($CurrentDay -ge 16 -and $CurrentDay -le 31) {
    $PayPeriodStartDay = 16
    $Since = "$CurrentYear-$CurrentMonth-16"
    $Until = "$CurrentYear-$CurrentMonth-$CurrentDay"
    $NumberOfWorkDays = 0

    for ($x = $PayPeriodStartDay; $x -le $CurrentDay; $x += 1) {
      switch ((get-date $CurrentYear-$CurrentMonth-$x).DayOfWeek) {
        'Monday' { $NumberOfWorkDays += 1 }
        'Tuesday' { $NumberOfWorkDays += 1 }
        'Wednesday' { $NumberOfWorkDays += 1 }
        'Thursday' { $NumberOfWorkDays += 1 }
        'Friday' { $NumberOfWorkDays += 1 }
      }
    }
    $WorkingHours = $NumberOfWorkDays * 8
  }

  #Authorization
  ##############
  # Invoke-RestMethod -Uri https://www.toggl.com/api/v8/me -Headers $headers -ContentType $contentType

  #Get Workspaces
  ###############
  # Invoke-RestMethod -Uri "https://www.toggl.com/api/v8/workspaces" -Headers $headers -ContentType $contentType

  #Get Users
  ##########
  # Invoke-RestMethod -Uri "https://www.toggl.com/api/v8/workspaces/$WorkspaceID/users" -Headers $headers -ContentType $contentType

  #Reports Request Parameters
  ###########################
  # user_agent: string, required, the name of your application or your email address so we can get in touch in case you're doing something wrong.
  # workspace_id: integer, required. The workspace whose data you want to access.
  # since: string, ISO 8601 date (YYYY-MM-DD), by default until - 6 days.
  # until: string, ISO 8601 date (YYYY-MM-DD), by default today
  # billable: possible values: yes/no/both, default both
  # client_ids: client ids separated by a comma, 0 if you want to filter out time entries without a client
  # project_ids: project ids separated by a comma, 0 if you want to filter out time entries without a project
  # user_ids: user ids separated by a comma
  # tag_ids: tag ids separated by a comma, 0 if you want to filter out time entries without a tag
  # task_ids: task ids separated by a comma, 0 if you want to filter out time entries without a task
  # time_entry_ids: time entry ids separated by a comma
  # description: string, time entry description
  # without_description: true/false, filters out the time entries which do not have a description ('(no description)')
  # order_field:
  # - date/description/duration/user in detailed reports
  # - title/duration/amount in summary reports
  # - title/day1/day2/day3/day4/day5/day6/day7/week_total in weekly report
  # order_desc: on/off, on for descending and off for ascending order
  # distinct_rates: on/off, default off
  # rounding: on/off, default off, rounds time according to workspace settings
  # display_hours: decimal/minutes, display hours with minutes or as a decimal number, default minutes

  #Detail Report
  #Last 6 days
  ##############
  $uriReport = "https://toggl.com/reports/api/v2/details?since=$Since&until=$Until&display_hours=decimal&rounding=on&user_agent=$UserAgent&workspace_id=$WorkspaceID"
  $TogglResponse = Invoke-RestMethod -Uri $uriReport -Headers $headers -ContentType $contentType
  $responseTotal = $TogglResponse.total_count
  $pageNum = 1
  $DetailReport = @()

  while ($responseTotal -gt 0) { 
    $TogglResponse = Invoke-RestMethod -Uri ($uriReport + '&page=' + $pageNum) -Headers $headers -ContentType $contentType
    $TogglResponseData = $TogglResponse.data
    $DetailReport += $TogglResponseData
    $responseTotal = $responseTotal - $TogglResponse.per_page 
  
    $pageNum++
  }

  if ($PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent) {
    Write-Output -InputObject "`nBillable Time Report"
    Write-Output -InputObject $DetailReport | Where-Object -FilterScript { $_.project_color -eq 3 } |
    Sort-Object -Property start |
      Format-List -Property @{N='Date';E={get-date $_.start -Format d}},
                                        Client,
                                        Project,
                                        Description,
                                        @{N='Duration'; E={"{0:N2}" -f ($_.Dur/1000/60/60)}}
  }

    foreach ($page in $DetailReport) {
      if ($page.project_color -eq 3) {
        $BillableHours += $page.Dur/1000/60/60
      }
  
      $TotalHours += $page.Dur/1000/60/60
    }


  Write-Output -InputObject ('Normal Hours: {0:N2}' -f $WorkingHours)

  $OvertimeHours = $TotalHours-$WorkingHours
  if ($OvertimeHours -lt 0) {
    Write-Output -InputObject ('Overtime Hours: {0:N2}' -f (0))
  } else {
    Write-Output -InputObject ('Overtime Hours: {0:N2}' -f ($OvertimeHours))
  }
  Write-Output -InputObject ("Total Hours: {0:N2}`n" -f $TotalHours)
  Write-Output -InputObject ('Billable Hours: {0:N2}' -f $BillableHours)
  Write-Output -InputObject ("Percent Billable: {0:P0}`n" -f ($BillableHours/$WorkingHours))
}