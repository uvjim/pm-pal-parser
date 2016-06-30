Param(
    [Parameter(Mandatory=$true)]
    [string]$WithoutPMFolder,

    [Parameter(Mandatory=$true)]
    [string]$WithPMFolder,

    [Parameter(Mandatory=$true)]
    [string]$Filename,

    [Parameter(Mandatory=$false)]
    [string[]]$Application,

    [Parameter(Mandatory=$false, ParameterSetName="TreatAsOne")]
    [string]$Label = "All Endpoints",

    [Parameter(Mandatory=$false, ParameterSetName="TreatAsOne")]
    [switch]$TreatAsOne
)

$script:objExcel = $null
$script:objWorkbook = $null

function Measure-WorkingProcess {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$File,

        [Parameter(Mandatory=$false)]
        [string[]]$Application
    )

    $ret = $null
    if (Test-Path -Path $File) {
        $html = New-Object -ComObject "HTMLFile"
        $source = Get-Content -Path $file -Raw
        $html.IHTMLDocument2_write($source)

        Write-Host "  Determining working process columns for extraction...    " -NoNewLine
        $table = $html.getElementsByTagName("TABLE") | Where id -eq "table3"
        $colAvg = ($table.rows()[0].cells() | Where InnerText -eq "Avg").cellIndex
        $colProcess = ($table.rows()[0].cells() | Where InnerText -eq "\Process(*)\Working Set").cellIndex
        Write-Host "(Process name: $colProcess,  Avg: $colAvg)"
        $numRows = $table.rows().Length
        $msg = "Parsing applications"
        Write-Host "  $msg...    " -NoNewLine
        $apps = @()
        foreach ($r in $table.rows()) {
            if ($r.rowIndex -ne 0) {
                $a = ($r.cells() | Where { $_.cellIndex -eq $colProcess}).InnerText
                $ws = [long]($r.cells() | Where { $_.cellIndex -eq $colAvg}).InnerText
                $apps += New-Object -Type PSObject -Property @{'Application'=$a.ToLower();'WorkingSet'=$ws}
            }
            Write-Progress -Activity "$msg in $File..." -CurrentOperation "Row $($r.rowIndex)/$numRows" -PercentComplete (($r.rowIndex / $numRows) * 100)
        }
        Write-Progress -Activity $msg -Completed
        Write-Host "Done"
        if (-not $Application) {
            $Application = @()
            $msg = "Determining applications..."
            Write-Host "  $msg    " -NoNewLine
            $Application = $apps | % { $_.Application.Split("/#")[1] } | Select -Unique
            Write-Host "Done"
        }
        $ret = @()
        $idxApplication = 0
        if ($Application -notcontains '_total') { $Application += '_total' }
        foreach ($a in $Application) {
            $idxApplication += 1
            $msg = "Retrieving details for application $a ($idxApplication/$($Application.Length))..."
            Write-Host "  $msg    " -NoNewLine
            Write-Progress -Activity $msg -PercentComplete (($idxApplication / $Application.Length) * 100)
            $instances = $apps | Where { $_.Application.Contains("/$($a.ToLower())") }
            $i = $instances | Measure-Object WorkingSet -Sum -Average
            $ret += New-Object -Type PSObject -Property @{'Application'=$a; 'NumberOfInstances'=$i.Count; 'Sum'=$i.Sum; 'Avg'=[string]$i.Average; 'Server'=[System.IO.Path]::GetFilenameWithoutExtension($File)}
            Write-Host "Done"
        }
        Write-Progress -Activity $msg -Completed
    }
    return $ret
}

function Measure-Sessions {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$File
    )

    $ret = $null
    if (Test-Path -Path $File) {
        $html = New-Object -ComObject "HTMLFile"
        $source = Get-Content -Path $file -Raw
        $html.IHTMLDocument2_write($source)

        Write-Host "  Determining session columns for extraction...    " -NoNewLine
        $table = $html.getElementsByTagName("TABLE") | Where id -eq "table7"
        $colAvg = ($table.rows()[0].cells() | Where InnerText -eq "Avg").cellIndex
        $colMax = ($table.rows()[0].cells() | Where InnerText -eq "Max").cellIndex
        $colMin = ($table.rows()[0].cells() | Where InnerText -eq "Min").cellIndex
        Write-Host "(Avg: $colAvg, Min: $colMin, Max: $colMax)"
        Write-Host "  Extracting session details...    " -NoNewLine
        $props = @{'Server'=[System.IO.Path]::GetFilenameWithoutExtension($File)}
        $props.Add('Avg', [long]($table.rows()[1].cells() | Where cellIndex -eq $colAvg).InnerText)
        $props.Add('Min', [long]($table.rows()[1].cells() | Where cellIndex -eq $colMin).InnerText)
        $props.Add('Max', [long]($table.rows()[1].cells() | Where cellIndex -eq $colMax).InnerText)
        $ret = New-Object -Type PSObject -Property $props
        Write-Host "Done"
    }
    return $ret
}

function Measure-ProcessorTime {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$File
    )

    $ret = $null
    if (Test-Path -Path $File) {
        $html = New-Object -ComObject "HTMLFile"
        $source = Get-Content -Path $file -Raw
        $html.IHTMLDocument2_write($source)

        Write-Host "  Determining processor time columns for extraction...    " -NoNewLine
        $table = $html.getElementsByTagName("TABLE") | Where id -eq "table4"
        $colAvg = ($table.rows()[0].cells() | Where InnerText -eq "Avg").cellIndex
        $colMax = ($table.rows()[0].cells() | Where InnerText -eq "Max").cellIndex
        $colMin = ($table.rows()[0].cells() | Where InnerText -eq "Min").cellIndex
        Write-Host "(Avg: $colAvg, Min: $colMin, Max: $colMax)"
        Write-Host "  Extracting session details...    " -NoNewLine
        $props = @{'Server'=[System.IO.Path]::GetFilenameWithoutExtension($File)}
        $props.Add('Avg', [double]($table.rows()[1].cells() | Where cellIndex -eq $colAvg).InnerText)
        $props.Add('Min', [double]($table.rows()[1].cells() | Where cellIndex -eq $colMin).InnerText)
        $props.Add('Max', [double]($table.rows()[1].cells() | Where cellIndex -eq $colMax).InnerText)
        $ret = New-Object -Type PSObject -Property $props
        Write-Host "Done"
    }
    return $ret
}

function Measure-ProcessorQueueLength {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$File
    )

    $ret = $null
    if (Test-Path -Path $File) {
        $html = New-Object -ComObject "HTMLFile"
        $source = Get-Content -Path $file -Raw
        $html.IHTMLDocument2_write($source)

        Write-Host "  Determining processor queue length columns for extraction...    " -NoNewLine
        $table = $html.getElementsByTagName("TABLE") | Where id -eq "table5"
        $colAvg = ($table.rows()[0].cells() | Where InnerText -eq "Avg").cellIndex
        $colMax = ($table.rows()[0].cells() | Where InnerText -eq "Max").cellIndex
        $colMin = ($table.rows()[0].cells() | Where InnerText -eq "Min").cellIndex
        Write-Host "(Avg: $colAvg, Min: $colMin, Max: $colMax)"
        Write-Host "  Extracting queue length details...    " -NoNewLine
        $props = @{'Server'=[System.IO.Path]::GetFilenameWithoutExtension($File)}
        $props.Add('Avg', [double]($table.rows()[1].cells() | Where cellIndex -eq $colAvg).InnerText)
        $props.Add('Min', [double]($table.rows()[1].cells() | Where cellIndex -eq $colMin).InnerText)
        $props.Add('Max', [double]($table.rows()[1].cells() | Where cellIndex -eq $colMax).InnerText)
        $ret = New-Object -Type PSObject -Property $props
        Write-Host "Done"
    }
    return $ret
}

function New-ExcelFile {
    Param(
        [Parameter(Mandatory=$true)]
        [string[]]$Sheets
    )

    Write-Host "Creating Excel file...    " -NoNewLine
    $script:objExcel = New-Object -ComObject Excel.Application
    # $script:objExcel.Visible = $true
    $script:objWorkbook = $script:objExcel.Workbooks.Add()
    foreach($s in $Sheets) { $script:objWorkbook.Sheets.Add([System.Reflection.Missing]::Value, $script:objWorkbook.Sheets.Item($script:objWorkbook.Sheets.Count)).Name = $s }
    $script:objExcel.Application.DisplayAlerts = $false
    $script:objWorkbook.Sheets(1).Delete() #remove the worksheet that we don't need
    $script:objExcel.Application.DisplayAlerts = $true
    Write-Host "Done"
}

function Write-ExcelFile {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$Filename
    )

    Write-Host "Saving Excel file to $Filename...    " -NoNewLine
    $script:objWorkbook.Close($true, $Filename)
    $script:objExcel.Quit()
    Write-Host "Done"
}

function New-ExcelApplicationMemorySheet {
    Param(
        [Parameter(Mandatory=$true)]
        $WithoutPM,

        [Parameter(Mandatory=$true)]
        $WithPM,

        [Parameter(Mandatory=$true)]
        [int]$SheetIndex,

        [Parameter(Mandatory=$false)]
        [switch]$AllServers
    )

    #-- Build the Application Memory Sheet --#
    Write-Host "  Building the Application Memory sheet...    " -NoNewLine
    $sheet = $script:objWorkbook.Sheets($SheetIndex)
    $sheet.Range("A1") = "Server Name"
    $sheet.Range("B1") = "Process Name"
    $sheet.Range("C1") = "Working Set (without PM)"
    $sheet.Range("D1") = "Working Set (with PM)"
    $sheet.Range("E1") = "%age decrease"
    $sheet.Rows(1).Font.Bold = $true

    #-- Get the without PM rows in --#
    $row = 2
    foreach($a in $WithoutPM) {
        if ($a.Application -ne '_total') {
            $sheet.Cells($row, 1) = $a.Server
            $sheet.Cells($row, 2) = $a.Application
            $sheet.Cells($row, 3) = [Math]::Round($a.Avg / 1MB)
            $row += 1
        }
    }
    #-- Get the with PM rows in --#
    foreach($a in $WithPM) {
        $objCell = $sheet.Columns(2).Find($a.Application, [System.Reflection.Missing]::Value, -4163)
        if ($objCell) {
            $addrStart = $objCell.Address()
            $break = $false
            do {
                if ($objCell.Offset(0, -1).Value() -eq $a.Server) {
                    $sheet.Cells($objCell.Row, 4) = [Math]::Round($a.Avg / 1MB)
                    $break = $true
                } else {
                    $objCell = $sheet.Columns(2).FindNext($objCell)
                }
            } until ($objCell.Address() -eq $addrStart -or -not $objCell -or $break)
        }
    }

    #-- workout %age decrease --#
    foreach($c in $sheet.Columns(1).Cells()) {
        if ($c.Row() -gt 1) {
            if (-not $c.Value()) { break }
            if ($c.Offset(0, 2).Value() -ne 0) {
                $c.Offset(0, 4).Formula = "=($($c.Offset(0, 2).Address())-$($c.Offset(0, 3).Address()))/$($c.Offset(0, 2).Address())"
                $c.Offset(0, 4).NumberFormat = "0.00%"
            } else {
                $c.Offset(0, 4) = "N/A"
                $c.Offset(0, 4).HorizontalAlignment = -4152
            }
        }
    }

    #-- Build the All Servers chart if needed --#
    if ($AllServers) {
        $objShape = $sheet.Shapes.AddChart2(286, 54)
        $objShape.Chart.ChartTitle.Text = "All Endpoints"
        $objCell = $sheet.Columns(1).Find('', [System.Reflection.Missing]::Value, -4163)
        $rngData = $sheet.Range($sheet.Cells(1, 1).Address(), $objCell.Offset(-1, 3).Address())
        $objShape.Chart.SetSourceData($rngData)
        $objShape.Chart.SetElement(104)
    }

    #-- Build the individual charts --#
    foreach($s in ($WithoutPM.Server | Select -Unique)) {
        $objCell = $sheet.Columns(1).Find($s, [System.Reflection.Missing]::Value, -4163)
        $addrStart = $objCell.Address()
        do {
            $addrEnd = $objCell.Address()
            $objCell = $sheet.Columns(1).FindNext($objCell)
        } until ($objCell.Address() -eq $addrStart -or -not $objCell)
        $rngData = $sheet.Range($sheet.Range($addrStart).Offset(0, 1), $sheet.Range($addrEnd).Offset(0, 3))
        $objShape = $sheet.Shapes.AddChart2(286, 54)
        $objShape.Chart.ChartTitle.Text = $s
        $objShape.Chart.SetSourceData($rngData)
        $objShape.Chart.PlotBy = 2
        $objShape.Chart.FullSeriesCollection(1).Name = $sheet.Cells(1, 3)
        $objShape.Chart.FullSeriesCollection(2).Name = $sheet.Cells(1, 4)
        $objShape.Chart.SetElement(104)
    }

    #-- Clean Up --#
    if ($rngData) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($rngData) }
    Remove-Variable rngData
    if ($objCell) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objCell) }
    Remove-Variable objCell
    if ($objShape) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($objShape) }
    Remove-Variable objShape
    if ($sheet) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) }
    Remove-Variable sheet
    Write-Host "Done"
}

function New-ExcelAvgCPUAndMemorySheet {
    Param(
        [Parameter(Mandatory=$true)]
        $WithoutPM,

        [Parameter(Mandatory=$true)]
        $WithPM,

        [Parameter(Mandatory=$true)]
        $WithoutPMSessions,

        [Parameter(Mandatory=$true)]
        $WithPMSessions,

        [Parameter(Mandatory=$true)]
        $WithoutPMProcessorTime,

        [Parameter(Mandatory=$true)]
        $WithPMProcessorTime,

        [Parameter(Mandatory=$true)]
        [int]$SheetIndex
    )

    Write-Host "  Building the Average CPU and Memory sheet...    " -NoNewLine
    $sheet = $script:objWorkbook.Sheets($SheetIndex)
    $sheet.Range("A1") = "Server Name"
    $sheet.Range("B1") = "Average Sessions (without PM)"
    $sheet.Range("C1") = "Average Sessions (with PM)"
    $sheet.Range("D1") = "Average Total Processor Time (without PM)"
    $sheet.Range("E1") = "Average Total Processor Time (with PM)"
    $sheet.Range("F1") = "Average Total Working Set (without PM)"
    $sheet.Range("G1") = "Average Total Working Set (with PM)"
    $sheet.Rows(1).Font.Bold = $true

    #-- Get the without PM rows in --#
    $row = 2
    $WithoutPM = $WithoutPM | Where Application -eq '_total'
    foreach($a in $WithoutPM) {
        $sheet.Cells($row, 1) = $a.Server
        $sheet.Cells($row, 2) = ($WithoutPMSessions | Where Server -eq $a.Server).Avg
        $sheet.Cells($row, 4) = ($WithoutPMProcessorTime | Where Server -eq $a.Server).Avg
        $sheet.Cells($row, 6) = [Math]::Round($a.Avg / 1MB)
        $row += 1
    }
    #-- Get the with PM rows in --#
    $WithPM = $WithPM | Where Application -eq '_total'
    foreach($a in $WithPM) {
        $objCell = $sheet.Columns(1).Find($a.Server, [System.Reflection.Missing]::Value, -4163)
        if ($objCell) {
            $sheet.Cells($objCell.Row, 3) = ($WithPMSessions | Where Server -eq $a.Server).Avg
            $sheet.Cells($objCell.Row, 5) = ($WithPMProcessorTime | Where Server -eq $a.Server).Avg
            $sheet.Cells($objCell.Row, 7) = [Math]::Round($a.Avg / 1MB)
        }
    }

    #-- Need to do the summary tables now (per server and session and overall) --#
    $objCell = $sheet.Columns(1).Find('', [System.Reflection.Missing]::Value, -4163)
    $row = $objCell.Row + 2
    $titleRow = $row
    $sheet.Rows($row).Font.Bold = $true
    $sheet.Cells($row, 1) = "Server Name"
    $sheet.Cells($row, 2) = "Avg Processor Time per Session (without PM)"
    $sheet.Cells($row, 3) = "Avg Processor Time per Session (with PM)"
    $sheet.Cells($row, 4) = "%age decrease"
    $sheet.Cells($row, 6) = "Server Name"
    $sheet.Cells($row, 7) = "Avg Working Set per Session (without PM)"
    $sheet.Cells($row, 8) = "Avg Working Set per Session (with PM)"
    $sheet.Cells($row, 9) = "%age decrease"
    $rngProcessorTimeData = $sheet.Cells($row, 1)
    $rngWorkingSetData = $sheet.Cells($row, 6)
    $row += 1
    for ($r=2; $r -le $objCell.Row - 1; $r++) {
        $sheet.Cells($row, 1) = $sheet.Cells($r, 1).Text
        $sheet.Cells($row, 2).Formula = "=ROUND($($sheet.Cells($r, 4).Address())/$($sheet.Cells($r, 2).Address()), 2)"
        $sheet.Cells($row, 3).Formula = "=ROUND($($sheet.Cells($r, 5).Address())/$($sheet.Cells($r, 3).Address()), 2)"
        $sheet.Cells($row, 4).Formula = "=($($sheet.Cells($row, 2).Address())-$($sheet.Cells($row, 3).Address()))/$($sheet.Cells($row, 2).Address())"
        $sheet.Cells($row, 4).NumberFormat = "0.00%"
        $sheet.Cells($row, 6) = $sheet.Cells($r, 1).Text
        $sheet.Cells($row, 7).Formula = "=ROUND($($sheet.Cells($r, 6).Address())/$($sheet.Cells($r, 2).Address()), 2)"
        $sheet.Cells($row, 8).Formula = "=ROUND($($sheet.Cells($r, 7).Address())/$($sheet.Cells($r, 3).Address()), 2)"
        $sheet.Cells($row, 9).Formula = "=($($sheet.Cells($row, 7).Address())-$($sheet.Cells($row, 8).Address()))/$($sheet.Cells($row, 7).Address())"
        $sheet.Cells($row, 9).NumberFormat = "0.00%"
        $row += 1
    }
    $rngProcessorTimeData = $sheet.Range($rngProcessorTimeData.Address(), $sheet.Cells($row - 1, 3))
    $rngWorkingSetData = $sheet.Range($rngWorkingSetData.Address(), $sheet.Cells($row - 1, 8))

    #-- Draw the charts --#
    $objShape = $sheet.Shapes.AddChart2(286, 54)
    $objShape.Chart.ChartTitle.Text = "Processor Time per Session"
    $objShape.Chart.SetSourceData($rngProcessorTimeData)
    $objShape.Chart.PlotBy = 2
    $objShape = $sheet.Shapes.AddChart2(286, 54)
    $objShape.Chart.ChartTitle.Text = "Working Set per Session"
    $objShape.Chart.SetSourceData($rngWorkingSetData)
    $objShape.Chart.PlotBy = 2

    Write-Host "Done"
}

function New-ExcelProcessorQueueSheet {
    Param(
        [Parameter(Mandatory=$true)]
        $WithoutPMSessions,

        [Parameter(Mandatory=$true)]
        $WithPMSessions,

        [Parameter(Mandatory=$true)]
        $WithoutPMProcessorQueueLength,

        [Parameter(Mandatory=$true)]
        $WithPMProcessorQueueLength,

        [Parameter(Mandatory=$true)]
        [int]$SheetIndex
    )

    Write-Host "  Building the Processor Queue sheet...    " -NoNewLine
    $sheet = $script:objWorkbook.Sheets($SheetIndex)
    $sheet.Range("A1") = "Processor Queue Length"
    $sheet.Range("A2") = "Without PM"
    $sheet.Range("A3") = "Server"
    $sheet.Range("B3") = "Sessions"
    $sheet.Range("C3") = "Queue Length"
    $row = 4
    foreach ($i in $WithoutPMSessions) {
        $sheet.Cells($row, 1) = $i.Server
        $sheet.Cells($row, 2) = $i.Avg
        $sheet.Cells($row, 3) = ($WithoutPMProcessorQueueLength | Where Server -eq $i.Server).Avg
        $row += 1
    }
    $sheet.Cells($row, 1) = "With PM"
    $row += 1
    $sheet.Cells($row, 1) = "Server"
    $sheet.Cells($row, 2) = "Sessions"
    $sheet.Cells($row, 3) = "Queue Length"
    $row += 1
    foreach ($i in $WithPMSessions) {
        $sheet.Cells($row, 1) = $i.Server
        $sheet.Cells($row, 2) = $i.Avg
        $sheet.Cells($row, 3) = ($WithPMProcessorQueueLength | Where Server -eq $i.Server).Avg
        $row += 1
    }

    #-- add the scatter chart --#
    $objChart = $sheet.Shapes.AddChart2(240, -4169).Chart
    $objChart.ChartTitle.Text = $sheet.Range("A1").Value()
    $objChart.SetElement(301)
    $objChart.SetElement(104)
    $objChart.Axes(2).HasTitle = $true
    $objChart.Axes(2).AxisTitle.Text = $sheet.Range("C3").Value()
    $objChart.Axes(1).HasTitle = $true
    $objChart.Axes(1).AxisTitle.Text = $sheet.Range("B3").Value()
    foreach ($c in $objChart.FullSeriesCollection()) { $c.Delete() | Out-Null }
    $s = $objChart.SeriesCollection().NewSeries.Invoke()
    $s.Name = $sheet.Range("A2").Value()
    $s.XValues = $sheet.Range($sheet.Cells(4, 2).Address(), $sheet.Cells($WithoutPMSessions.Count - 1 + 4, 2).Address())
    $s.Values = $sheet.Range($sheet.Cells(4, 3).Address(), $sheet.Cells($WithoutPMSessions.Count - 1 + 4, 3).Address())
    $s = $objChart.SeriesCollection().NewSeries.Invoke()
    $s.Name = $sheet.Cells($row - $WithPMSessions.Count - 2, 1).Value()
    $s.XValues =  $sheet.Range($sheet.Cells($row - $WithPMSessions.Count, 2).Address(), $sheet.Cells($row - 1, 2).Address())
    $s.Values = $sheet.Range($sheet.Cells($row - $WithPMSessions.Count, 3).Address(), $sheet.Cells($row - 1, 3).Address())
    Write-Host "Done"
}

function Get-TreatAsOneObject {
    Param(
        [Parameter(Mandatory=$true)]
        $InputObject,

        [Parameter(Mandatory=$false)]
        [string]$Key,

        [Parameter(Mandatory=$true)]
        [string]$Server
    )

    $ret = @()
    if (($InputObject | Get-Member) | Where Name -eq Group) {
        foreach ($app in $InputObject) {
            $props = @{'Server'=$Server}
            $props.Add($Key, $app.Name)
            $NoteProps = $app.Group | Get-Member -MemberType NoteProperty | Where { -not $_.Definition.StartsWith('string') }
            foreach ($np in $NoteProps) {
                $m = $app.Group | Measure-Object $np.Name -Sum
                $props.Add($np.Name, $m.Sum)
            }
            $props.Add('Avg', $props['Sum'] / $props['NumberOfInstances'])
            $ret += New-Object PSObject -Property $props
        }
    } else {
        $props = @{'Server'=$Server}
        $props.Add('Avg', ($InputObject | Measure-Object -Property 'Avg' -Average).Average)
        $props.Add('Max', ($InputObject | Measure-Object -Property 'Max' -Maximum).Maximum)
        $props.Add('Min', ($InputObject | Measure-Object -Property 'Min' -Minimum).Minimum)
        $ret += New-Object PSObject -Property $props
    }
    return $ret
}

if ((Test-Path -Path $WithoutPMFolder) -and (Test-Path -Path $WithPMFolder)) {
    $WithoutPMWorkingProcess = @()
    $WithoutPMSessions = @()
    $WithoutPMProcessorTime = @()
    $WithoutPMProcessorQueueLength = @()
    foreach ($f in (Get-ChildItem -Path "$WithoutPMFolder\*.htm")) {
        Write-Host "Parsing $($f.FullName)"
        $d = Measure-WorkingProcess -File $f.FullName -Application $Application
        if ($d) { $WithoutPMWorkingProcess += $d }
        $s = Measure-Sessions -File $f.FullName
        if ($s) { $WithoutPMSessions += $s }
        $p = Measure-ProcessorTime -File $f.FullName
        if ($p) { $WithoutPMProcessorTime += $p }
        $l = Measure-ProcessorQueueLength -File $f.FullName
        if ($l) { $WithoutPMProcessorQueueLength += $l }
    }
    $WithPMWorkingProcess = @()
    $WithPMSessions = @()
    $WithPMProcessorTime = @()
    $WithPMProcessorQueueLength = @()
    foreach ($f in (Get-ChildItem -Path "$WithPMFolder\*.htm")) {
        Write-Host "Parsing $($f.FullName)"
        $d = Measure-WorkingProcess -File $f.FullName -Application $Application
        if ($d) { $WithPMWorkingProcess += $d }
        $s = Measure-Sessions -File $f.FullName
        if ($s) { $WithPMSessions += $s }
        $p = Measure-ProcessorTime -File $f.FullName
        if ($p) { $WithPMProcessorTime += $p }
        $l = Measure-ProcessorQueueLength -File $f.FullName
        if ($l) { $WithPMProcessorQueueLength += $l }
    }
    if ($PSCmdlet.ParameterSetName -eq 'TreatAsOne') {
        Write-Host "Aggregating data...    " -NoNewLine
        $WithoutPMWorkingProcess = Get-TreatAsOneObject -InputObject ($WithoutPMWorkingProcess | Group-Object -Property Application) -Key 'Application' -Server $Label
        $WithPMWorkingProcess = Get-TreatAsOneObject -InputObject ($WithPMWorkingProcess | Group-Object -Property Application) -Key 'Application' -Server $Label
        $WithoutPMSessionsSep = $WithoutPMSessions
        $WithPMSessionsSep = $WithPMSessions
        $WithoutPMSessions = Get-TreatAsOneObject -InputObject $WithoutPMSessions -Server $Label
        $WithPMSessions = Get-TreatAsOneObject -InputObject $WithPMSessions -Server $Label
        $WithoutPMProcessorTime = Get-TreatAsOneObject -InputObject $WithoutPMProcessorTime -Server $Label
        $WithPMProcessorTime = Get-TreatAsOneObject -InputObject $WithPMProcessorTime -Server $Label
        Write-Host "Done"
    }
}
if ($WithoutPMWorkingProcess -and $WithPMWorkingProcess) {
    $WithoutPMWorkingProcess = $WithoutPMWorkingProcess | Sort Server,Application
    $WithPMWorkingProcess = $WithPMWorkingProcess | Sort Server,Application
    $sheets = @('Average CPU & Memory', 'Application Memory', 'Processor Queue Length')
    New-ExcelFile -Sheets $sheets
    New-ExcelApplicationMemorySheet -WithoutPM $WithoutPMWorkingProcess -WithPM $WithPMWorkingProcess -SheetIndex $([array]::indexOf($sheets, 'Application Memory') + 1) -AllServers:($PSCmdlet.ParameterSetName -ne 'TreatAsOne')
    New-ExcelAvgCPUAndMemorySheet -WithoutPM $WithoutPMWorkingProcess -WithPM $WithPMWorkingProcess -WithoutPMSessions $WithoutPMSessions -WithPMSessions $WithPMSessions -WithoutPMProcessorTime $WithoutPMProcessorTime -WithPMProcessorTime $WithPMProcessorTime -SheetIndex $([array]::indexOf($sheets, 'Average CPU & Memory') + 1)
    if ($PSCmdlet.ParameterSetName -eq 'TreatAsOne') {
        New-ExcelProcessorQueueSheet -WithoutPMSessions $WithoutPMSessionsSep -WithPMSessions $WithPMSessionsSep -WithoutPMProcessorQueueLength $WithoutPMProcessorQueueLength -WithPMProcessorQueueLength $WithPMProcessorQueueLength -SheetIndex $([array]::indexOf($sheets, 'Processor Queue Length') + 1)
    } else {
        New-ExcelProcessorQueueSheet -WithoutPMSessions $WithoutPMSessions -WithPMSessions $WithPMSessions -WithoutPMProcessorQueueLength $WithoutPMProcessorQueueLength -WithPMProcessorQueueLength $WithPMProcessorQueueLength -SheetIndex $([array]::indexOf($sheets, 'Processor Queue Length') + 1)
    }
    Write-ExcelFile -Filename $Filename
    if ($script:objWorkbook) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:objWorkbook) }
    Remove-Variable objWorkbook -Scope Script
    if ($script:objExcel) { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:objExcel) }
    Remove-Variable objExcel -Scope Script
}
