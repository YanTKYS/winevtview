Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# =========================================
# EventPeek - 簡易イベントビューア
# Windows PowerShell 5.1 対応版
# =========================================

$form = New-Object System.Windows.Forms.Form
$form.Text = "EventPeek - 簡易イベントビューア"
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)

# 対象ログ
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "対象ログ:"
$lblLog.Location = New-Object System.Drawing.Point(10, 14)
$lblLog.Width = 60

$comboLog = New-Object System.Windows.Forms.ComboBox
$comboLog.Items.AddRange(@("System + Application", "System", "Application"))
$comboLog.SelectedIndex = 0
$comboLog.Location = New-Object System.Drawing.Point(75, 10)
$comboLog.Width = 170
$comboLog.DropDownStyle = "DropDownList"

# 期間
$lblPeriod = New-Object System.Windows.Forms.Label
$lblPeriod.Text = "期間:"
$lblPeriod.Location = New-Object System.Drawing.Point(260, 14)
$lblPeriod.Width = 40

$comboPeriod = New-Object System.Windows.Forms.ComboBox
$comboPeriod.Items.AddRange(@("直近1時間", "直近24時間", "直近7日"))
$comboPeriod.SelectedIndex = 1
$comboPeriod.Location = New-Object System.Drawing.Point(300, 10)
$comboPeriod.Width = 120
$comboPeriod.DropDownStyle = "DropDownList"

# イベントID
$lblEventId = New-Object System.Windows.Forms.Label
$lblEventId.Text = "イベントID:"
$lblEventId.Location = New-Object System.Drawing.Point(435, 14)
$lblEventId.Width = 70

$txtEventId = New-Object System.Windows.Forms.TextBox
$txtEventId.Location = New-Object System.Drawing.Point(505, 10)
$txtEventId.Width = 80

# ボタン
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "検索"
$btnSearch.Location = New-Object System.Drawing.Point(600, 8)
$btnSearch.Width = 80

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "CSV出力"
$btnExport.Location = New-Object System.Drawing.Point(690, 8)
$btnExport.Width = 90

$btnViewer = New-Object System.Windows.Forms.Button
$btnViewer.Text = "イベントビューア"
$btnViewer.Location = New-Object System.Drawing.Point(790, 8)
$btnViewer.Width = 120

# 一覧グリッド
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 45)
$grid.Size = New-Object System.Drawing.Size(1060, 430)
$grid.Anchor = "Top, Bottom, Left, Right"
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.RowHeadersVisible = $false

# 詳細欄
$detail = New-Object System.Windows.Forms.TextBox
$detail.Location = New-Object System.Drawing.Point(10, 485)
$detail.Size = New-Object System.Drawing.Size(1060, 165)
$detail.Anchor = "Bottom, Left, Right"
$detail.Multiline = $true
$detail.ScrollBars = "Vertical"
$detail.ReadOnly = $true

# ステータス表示
$status = New-Object System.Windows.Forms.Label
$status.Text = ""
$status.Location = New-Object System.Drawing.Point(920, 14)
$status.Width = 150
$status.Anchor = "Top, Right"

$script:CurrentRows = @()

function Get-TargetLogs {
    switch ($comboLog.SelectedItem) {
        "System" {
            return @("System")
        }
        "Application" {
            return @("Application")
        }
        default {
            return @("System", "Application")
        }
    }
}

function Get-StartTime {
    switch ($comboPeriod.SelectedItem) {
        "直近1時間" {
            return (Get-Date).AddHours(-1)
        }
        "直近7日" {
            return (Get-Date).AddDays(-7)
        }
        default {
            return (Get-Date).AddDays(-1)
        }
    }
}

function Search-Events {
    $detail.Text = ""
    $status.Text = "検索中..."
    $form.Refresh()

    $logs = Get-TargetLogs
    $start = Get-StartTime

    $filter = @{
        LogName   = $logs
        Level     = 1, 2, 3   # 1=Critical, 2=Error, 3=Warning
        StartTime = $start
    }

    $eventIdText = $txtEventId.Text.Trim()

    if ($eventIdText -ne "") {
        if ($eventIdText -notmatch '^\d+$') {
            [System.Windows.Forms.MessageBox]::Show(
                "イベントIDは数字で入力してください。",
                "入力エラー",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $status.Text = ""
            return
        }

        $filter.Id = [int]$eventIdText
    }

    try {
        $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 500 -ErrorAction Stop)

        $script:CurrentRows = @(
            $events | ForEach-Object {
                $message = $_.Message
                if ($null -eq $message) {
                    $message = ""
                }

                $summary = (($message -replace "`r?`n", " ") -replace "\s+", " ").Trim()

                if ($summary.Length -gt 300) {
                    $summary = $summary.Substring(0, 300) + "..."
                }

                [PSCustomObject]@{
                    TimeCreated    = $_.TimeCreated
                    LogName        = $_.LogName
                    Level          = $_.LevelDisplayName
                    Id             = $_.Id
                    ProviderName   = $_.ProviderName
                    MessageSummary = $summary
                    FullMessage    = $message
                }
            }
        )

        $grid.DataSource = $null
        $grid.DataSource = @(
            $script:CurrentRows |
            Select-Object TimeCreated, LogName, Level, Id, ProviderName, MessageSummary
        )

        if ($grid.Columns.Count -gt 0) {
            $grid.Columns["TimeCreated"].HeaderText = "時刻"
            $grid.Columns["LogName"].HeaderText = "ログ"
            $grid.Columns["Level"].HeaderText = "レベル"
            $grid.Columns["Id"].HeaderText = "ID"
            $grid.Columns["ProviderName"].HeaderText = "ソース"
            $grid.Columns["MessageSummary"].HeaderText = "メッセージ"

            $grid.Columns["TimeCreated"].Width = 140
            $grid.Columns["LogName"].Width = 90
            $grid.Columns["Level"].Width = 80
            $grid.Columns["Id"].Width = 70
            $grid.Columns["ProviderName"].Width = 180
        }

        $status.Text = "$($script:CurrentRows.Count) 件"
    }
    catch {
        $grid.DataSource = $null
        $script:CurrentRows = @()
        $status.Text = "エラー"

        [System.Windows.Forms.MessageBox]::Show(
            "イベントログの取得に失敗しました。`r`n`r`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

$btnSearch.Add_Click({
    Search-Events
})

$btnExport.Add_Click({
    if (-not $script:CurrentRows -or $script:CurrentRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "出力する結果がありません。",
            "確認",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }

    $fileName = "EventPeek_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $path = Join-Path ([Environment]::GetFolderPath("Desktop")) $fileName

    try {
        $script:CurrentRows |
            Select-Object TimeCreated, LogName, Level, Id, ProviderName, MessageSummary |
            Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "CSVを出力しました。`r`n`r`n$path",
            "出力完了",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "CSV出力に失敗しました。`r`n`r`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$btnViewer.Add_Click({
    try {
        Start-Process eventvwr.msc
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "イベントビューアを開けませんでした。`r`n`r`n$($_.Exception.Message)",
            "エラー",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -gt 0) {
        $index = $grid.SelectedRows[0].Index

        if ($index -ge 0 -and $index -lt $script:CurrentRows.Count) {
            $row = $script:CurrentRows[$index]

            $detail.Text = @"
時刻: $($row.TimeCreated)
ログ: $($row.LogName)
レベル: $($row.Level)
イベントID: $($row.Id)
ソース: $($row.ProviderName)

$($row.FullMessage)
"@
        }
    }
})

$form.Controls.AddRange(@(
    $lblLog,
    $comboLog,
    $lblPeriod,
    $comboPeriod,
    $lblEventId,
    $txtEventId,
    $btnSearch,
    $btnExport,
    $btnViewer,
    $status,
    $grid,
    $detail
))

# 起動時に初回検索
Search-Events

[void]$form.ShowDialog()
