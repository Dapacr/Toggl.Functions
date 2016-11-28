function Get-BillableTimeReport {
  [CmdletBinding()] 
  Param (
    [Parameter(Position=0)]
    [Alias('Day')]
    [datetime]$Date
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
  }
  else {
    $Date = Get-Date
  }
  $CurrentDay = $Date.Day
  $CurrentMonth = $Date.Month
  $CurrentYear = $Date.Year

  # Calculate normal working hours for the specified pay period
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
  }
  elseif ($CurrentDay -ge 16 -and $CurrentDay -le 31) {
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

  # Query Toggl API for report details
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

  # Output billable time entries if verbose is enabled (Orange projects are considered billable; all other colors are considered non-billable)
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

  # Calculate billable and total hours (Orange projects are considered billable; all other colors are considered non-billable)
  foreach ($page in $DetailReport) {
    if ($page.project_color -eq 3) {
      $BillableHours += $page.Dur/1000/60/60
    }
  
    $TotalHours += $page.Dur/1000/60/60
  }
  
  # Output summary totals
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