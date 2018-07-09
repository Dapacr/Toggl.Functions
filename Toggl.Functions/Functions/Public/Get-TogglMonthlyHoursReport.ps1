<#
        .Synopsis
        Calculate a monthly billable hours report in the specified date range.
        .Example
        Get-TogglMonthlyHoursReport | Export-Csv -NoTypeInformation 'c:\temp\TogglMonthlyBillableHoursReport.csv'
        Calculate monthly billable hours for the past year, covert to CSV and copy to the clipboard. Can be pasted into Excel directly
        .Example
        Get-TogglMonthlyHoursReport | Out-GridView
        Calculate monthly billable hours for the past year and output to grid view
        .Example
        Get-TogglMonthlyHoursReport | ConvertTo-Csv -NoTypeInformation -Delimiter "`t" | clip
        Calculate monthly billable hours, covert to CSV and copy to the clipboard. Can be pasted directly into Excel
        .Link
        https://github.com/Dapacruz/Toggl.Functions
#>
function Get-TogglMonthlyHoursReport {
    [CmdletBinding()]
    Param (
        [Parameter(Position=0)]
        [datetime]$From = (Get-Date).AddMonths(-12),
        [Parameter(Position=1)]
        [datetime]$To = (Get-Date)
    )

    $utilization_report = Get-TogglUtilizationReport -From $From -To $To -ExcludeCurrentPeriod | Group-Object -Property {(Get-Date $_.PeriodStart).Year}, {(Get-Date $_.PeriodStart).Month}
    $report = @()

    foreach ($period in $utilization_report) {
        $obj = New-Object -TypeName PSObject

        # Insert custom TypeName (defined in $PSScriptRoot\format.ps1xml) to control default display
        $obj.PSTypeNames.Insert(0,'Toggl.Report.Monthly.Hours')

        $count = $period.Group.Count
        if ($count -eq 2) {
            $period_start = Get-Date $period.Group[0].PeriodStart
            $month = '{0:d2}' -f $period_start.Month
            $year = $period_start.Year
            $billable_hrs = $period.Group.'Billable(hrs)' | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            $pto_hrs = $period.Group.'PTO(hrs)' | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            $training_hrs = $period.Group.'Training(hrs)' | Measure-Object -Sum | Select-Object -ExpandProperty Sum
            
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Month'  -Value "$month-$year"
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Billable(hrs)' -Value ('{0:N2}' -f $billable_hrs)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'PTO(hrs)' -Value ('{0:N2}' -f $pto_hrs)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Training(hrs)' -Value ('{0:N2}' -f $training_hrs)
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Total(hrs)' -Value ('{0:N2}' -f ($billable_hrs + $pto_hrs + $training_hrs))
            
            $report += $obj
        }
    }

    Write-Output $report
}
