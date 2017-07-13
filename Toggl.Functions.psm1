function Get-BillableTimeReport {
    [CmdletBinding()] 
    Param (
        [Parameter(Position=0)]
        [Alias('Day')]
        [datetime]$Date = (Get-Date)
    )

    $UserAgent = "dcruz@dsatechnologies.com"
    $WorkspaceID = "789619"
    $User = "982b4538c4ac97feff249d0c54463164" # <-- enter your API token here
    $Pass = "api_token"
    $Pair = "$($User):$($Pass)"
    $Bytes = [System.Text.Encoding]::ASCII.GetBytes($Pair)
    $Base64 = [System.Convert]::ToBase64String($Bytes)
    $BasicAuthValue = "Basic $Base64"
    $Headers = @{ Authorization = $BasicAuthValue }
    $contentType = "application/json"
    $BillableHours = 0
    $UtilizedHours = 0
    $PtoHours = 0
    $HolidayHours = 0
    $OvertimeHours = 0
    $TotalHours = 0
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
    $TogglResponse = Invoke-RestMethod -Uri $uriReport -Headers $Headers -ContentType $contentType
    $responseTotal = $TogglResponse.total_count
    $pageNum = 1
    $DetailReport = @()

    while ($responseTotal -gt 0) { 
        $TogglResponse = Invoke-RestMethod -Uri ($uriReport + '&page=' + $pageNum) -Headers $Headers -ContentType $contentType
        $TogglResponseData = $TogglResponse.data
        $DetailReport += $TogglResponseData
        $responseTotal = $responseTotal - $TogglResponse.per_page 
  
        $pageNum++
    }

    # Output billable time entries if verbose is enabled
    if ($PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent) {
        Write-Output -InputObject "`nBillable Time Report"
        Write-Output -InputObject $DetailReport | Where-Object -FilterScript { $_.tags -eq 'Billable' } |
            Sort-Object -Property start |
                Format-List -Property @{N='Date';E={get-date $_.start -Format d}},
                    Client,
                    Project,
                    Description,
                    @{N='Duration'; E={"{0:N2}" -f ($_.Dur/1000/60/60)}}
    }

    # Calculate billable, utilized, PTO, holiday and total hours
    foreach ($page in $DetailReport) {
        if ($page.tags -eq 'Billable') {
            $BillableHours += $page.Dur/1000/60/60
        }
  
        if ($page.tags -eq 'Utilized') {
            $UtilizedHours += $page.Dur/1000/60/60
        }
      
        if ($page.tags -eq 'PTO') {
            $PtoHours += $page.Dur/1000/60/60
        }

        if ($page.tags -eq 'Holiday') {
            $HolidayHours += $page.Dur/1000/60/60
        }
      
        $TotalHours += $page.Dur/1000/60/60
    }
  
    # Output summary totals
    Write-Output -InputObject ("`nTotal Hours: {0:N2}`n" -f $TotalHours)
    
    Write-Output -InputObject ("`tNormal: {0:N2}" -f $WorkingHours)
    $OvertimeHours = $TotalHours-$WorkingHours
    if ($OvertimeHours -lt 0) {
        Write-Output -InputObject ("`tOvertime: {0:N2}" -f (0))
    }
    else {
        Write-Output -InputObject ("`tOvertime: {0:N2}" -f ($OvertimeHours))
    }
    Write-Output -InputObject ("`tPTO: {0:N2}" -f $PTOHours)
    Write-Output -InputObject ("`tHoliday: {0:N2}" -f $HolidayHours)
    Write-Output -InputObject ("`tBillable: {0:N2}" -f $BillableHours)
    Write-Output -InputObject ("`tUtilized: {0:N2}`n`n" -f $UtilizedHours)
    
    
    Write-Output -InputObject ("Percent Billable: {0:P0}`n" -f ($BillableHours/($WorkingHours-$PtoHours-$HolidayHours)))
    Write-Output -InputObject ("Percent Utilized: {0:P0}`n" -f (($BillableHours+$UtilizedHours)/($WorkingHours-$PtoHours-$HolidayHours)))
}