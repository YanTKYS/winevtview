#Requires -Version 5.1

# =========================================
# WinEvtView - 簡易イベントビューア
# System / Application の主要なイベントをすぐ確認するための軽量ツール
# Windows 11 / Windows PowerShell 5.1 対応
# =========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# 起動時および検索時の共通設定
$script:MaxEvents    = 500          # 1回の検索で取得する最大件数
$script:EventLevels  = @(1, 2, 3)   # 1=Critical, 2=Error, 3=Warning
$script:CurrentRows  = @()          # 一覧に表示中のイベント（行番号と同じ並び）

# -----------------------------------------
# 共通処理
# -----------------------------------------

function Show-Dialog {
    param(
        [string]$Text,
        [string]$Caption,
        [System.Windows.Forms.MessageBoxIcon]$Icon
    )

    [void][System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Caption,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    )
}

function Test-NoMatchingEventsError {
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) {
        return $false
    }

    # Get-WinEvent は「条件に一致するイベントが無い」場合もエラーを返すため、
    # 表示言語に依存しないエラーIDで通常状態かどうかを判定する。
    return ($ErrorRecord.FullyQualifiedErrorId -like 'NoMatchingEventsFound*')
}

function Add-GridColumn {
    param(
        [System.Windows.Forms.DataGridView]$Grid,
        [string]$Name,
        [string]$HeaderText,
        [int]$Width,
        [switch]$Fill
    )

    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $column.Name = $Name
    $column.HeaderText = $HeaderText
    # 行番号と取得結果の並びを一致させるため、並べ替えは行わない
    $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable

    if ($Fill) {
        # メッセージ列だけを残り幅いっぱいに広げる
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
        $column.MinimumWidth = $Width
    }
    else {
        $column.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::None
        $column.Width = $Width
    }

    [void]$Grid.Columns.Add($column)
}

# -----------------------------------------
# 画面
# -----------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "WinEvtView - 簡易イベントビューア"
$form.ClientSize = New-Object System.Drawing.Size(1084, 661)
$form.MinimumSize = New-Object System.Drawing.Size(900, 560)
$form.StartPosition = "CenterScreen"

# 対象ログ
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "対象ログ:"
$lblLog.Location = New-Object System.Drawing.Point(12, 9)
$lblLog.Size = New-Object System.Drawing.Size(58, 21)
$lblLog.TextAlign = "MiddleLeft"

$comboLog = New-Object System.Windows.Forms.ComboBox
$comboLog.Items.AddRange(@("System + Application", "System", "Application"))
$comboLog.SelectedIndex = 0
$comboLog.Location = New-Object System.Drawing.Point(70, 9)
$comboLog.Width = 160
$comboLog.DropDownStyle = "DropDownList"

# 期間
$lblPeriod = New-Object System.Windows.Forms.Label
$lblPeriod.Text = "期間:"
$lblPeriod.Location = New-Object System.Drawing.Point(240, 9)
$lblPeriod.Size = New-Object System.Drawing.Size(38, 21)
$lblPeriod.TextAlign = "MiddleLeft"

$comboPeriod = New-Object System.Windows.Forms.ComboBox
$comboPeriod.Items.AddRange(@("直近1時間", "直近24時間", "直近7日"))
$comboPeriod.SelectedIndex = 0
$comboPeriod.Location = New-Object System.Drawing.Point(278, 9)
$comboPeriod.Width = 110
$comboPeriod.DropDownStyle = "DropDownList"

# イベントID
$lblEventId = New-Object System.Windows.Forms.Label
$lblEventId.Text = "イベントID:"
$lblEventId.Location = New-Object System.Drawing.Point(398, 9)
$lblEventId.Size = New-Object System.Drawing.Size(68, 21)
$lblEventId.TextAlign = "MiddleLeft"

$txtEventId = New-Object System.Windows.Forms.TextBox
$txtEventId.Location = New-Object System.Drawing.Point(466, 9)
$txtEventId.Width = 70
$txtEventId.MaxLength = 6

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($txtEventId, "半角数字で入力します。空欄の場合はすべてのイベントIDが対象です。")

# ボタン
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "検索"
$btnSearch.Location = New-Object System.Drawing.Point(548, 8)
$btnSearch.Size = New-Object System.Drawing.Size(80, 24)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "CSV出力"
$btnExport.Location = New-Object System.Drawing.Point(636, 8)
$btnExport.Size = New-Object System.Drawing.Size(90, 24)

$btnViewer = New-Object System.Windows.Forms.Button
$btnViewer.Text = "イベントビューア"
$btnViewer.Location = New-Object System.Drawing.Point(734, 8)
$btnViewer.Size = New-Object System.Drawing.Size(140, 24)

# 一覧グリッド
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 44)
$grid.Size = New-Object System.Drawing.Size(1060, 400)
$grid.Anchor = "Top, Bottom, Left, Right"
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.AutoSizeColumnsMode = "None"

Add-GridColumn -Grid $grid -Name "TimeCreated"  -HeaderText "時刻"       -Width 150
Add-GridColumn -Grid $grid -Name "LogName"      -HeaderText "ログ"       -Width 100
Add-GridColumn -Grid $grid -Name "Level"        -HeaderText "レベル"     -Width 80
Add-GridColumn -Grid $grid -Name "Id"           -HeaderText "ID"         -Width 70
Add-GridColumn -Grid $grid -Name "ProviderName" -HeaderText "ソース"     -Width 180
Add-GridColumn -Grid $grid -Name "Message"      -HeaderText "メッセージ" -Width 220 -Fill

# 詳細欄
$detail = New-Object System.Windows.Forms.TextBox
$detail.Location = New-Object System.Drawing.Point(12, 452)
$detail.Size = New-Object System.Drawing.Size(1060, 179)
$detail.Anchor = "Bottom, Left, Right"
$detail.Multiline = $true
$detail.ScrollBars = "Vertical"
$detail.ReadOnly = $true

# ステータス表示
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = ""
[void]$statusStrip.Items.Add($statusLabel)

# -----------------------------------------
# 検索
# -----------------------------------------

function New-EventRow {
    param($EventRecord)

    # メッセージやレベル表示名はイベント提供元のメッセージリソースに依存するため、
    # 取得できないイベントが混ざっていても一覧全体が壊れないようイベント単位で受け止める。
    # PowerShell はプロパティ取得時の例外を握りつぶして $null を返すことがあるため、
    # 例外と値の両方で判定する。
    $message = $null

    try {
        $message = $EventRecord.Message
    }
    catch {
        $message = $null
    }

    if ([string]::IsNullOrWhiteSpace($message)) {
        $message = "（このイベントのメッセージを取得できませんでした）"
    }

    $level = $null

    try {
        $level = $EventRecord.LevelDisplayName
    }
    catch {
        $level = $null
    }

    if ($null -eq $level) {
        $level = ""
    }

    # 一覧・CSV用に改行や連続する空白を1行にまとめる
    $singleLine = ($message -replace '\s+', ' ').Trim()
    $summary = $singleLine

    if ($summary.Length -gt 200) {
        $summary = $summary.Substring(0, 200) + "..."
    }

    [PSCustomObject]@{
        TimeText       = "{0:yyyy/MM/dd HH:mm:ss}" -f $EventRecord.TimeCreated
        LogName        = $EventRecord.LogName
        Level          = $level
        Id             = $EventRecord.Id
        ProviderName   = $EventRecord.ProviderName
        MessageSummary = $summary
        MessageLine    = $singleLine
        FullMessage    = $message
    }
}

function Get-TargetLogs {
    switch ($comboLog.SelectedItem) {
        "System"      { return @("System") }
        "Application" { return @("Application") }
        default       { return @("System", "Application") }
    }
}

function Get-StartTime {
    switch ($comboPeriod.SelectedItem) {
        "直近24時間" { return (Get-Date).AddDays(-1) }
        "直近7日"    { return (Get-Date).AddDays(-7) }
        default      { return (Get-Date).AddHours(-1) }
    }
}

function Show-EventDetail {
    if ($grid.SelectedRows.Count -eq 0) {
        $detail.Text = ""
        return
    }

    $index = $grid.SelectedRows[0].Index

    if ($index -lt 0 -or $index -ge $script:CurrentRows.Count) {
        $detail.Text = ""
        return
    }

    $row = $script:CurrentRows[$index]

    # テキストボックスは CRLF でないと改行として表示しないため、改行コードをそろえる
    $body = ($row.FullMessage -replace "`r`n", "`n") -replace "`n", "`r`n"

    $detail.Text = @(
        "時刻: $($row.TimeText)",
        "ログ: $($row.LogName)",
        "レベル: $($row.Level)",
        "イベントID: $($row.Id)",
        "ソース: $($row.ProviderName)",
        "",
        $body
    ) -join "`r`n"
}

function Clear-Result {
    $script:CurrentRows = @()
    $grid.Rows.Clear()
    $detail.Text = ""
}

function Search-Events {
    $eventIdText = $txtEventId.Text.Trim()
    $eventId = 0

    if ($eventIdText -ne "") {
        if ($eventIdText -notmatch '^\d+$' -or -not [int]::TryParse($eventIdText, [ref]$eventId)) {
            Show-Dialog `
                -Text "イベントIDは半角数字で入力してください。`r`n（すべてのイベントIDを対象にする場合は空欄にします）" `
                -Caption "入力エラー" `
                -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
            $txtEventId.Focus()
            $txtEventId.SelectAll()
            return
        }
    }

    Clear-Result
    $statusLabel.Text = "検索中..."
    $statusStrip.Refresh()

    $filter = @{
        LogName   = Get-TargetLogs
        Level     = $script:EventLevels
        StartTime = Get-StartTime
    }

    if ($eventIdText -ne "") {
        $filter.Id = $eventId
    }

    $queryErrors = $null
    $events = @()

    try {
        # 該当イベントが無い場合もエラーになるため、いったん受け取ってから内容で判定する
        $events = @(
            Get-WinEvent -FilterHashtable $filter -MaxEvents $script:MaxEvents `
                -ErrorAction SilentlyContinue -ErrorVariable queryErrors
        )
    }
    catch {
        $statusLabel.Text = "取得に失敗しました"
        Show-Dialog `
            -Text "イベントログの取得に失敗しました。`r`n`r`n$($_.Exception.Message)" `
            -Caption "エラー" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $realErrors = @($queryErrors | Where-Object { -not (Test-NoMatchingEventsError $_) })

    if ($realErrors.Count -gt 0) {
        $statusLabel.Text = "取得に失敗しました"
        Show-Dialog `
            -Text "イベントログの取得に失敗しました。`r`n`r`n$($realErrors[0].Exception.Message)" `
            -Caption "エラー" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $script:CurrentRows = @(
        $events | ForEach-Object { New-EventRow -EventRecord $_ }
    )

    $grid.SuspendLayout()
    try {
        foreach ($row in $script:CurrentRows) {
            [void]$grid.Rows.Add(
                $row.TimeText,
                $row.LogName,
                $row.Level,
                $row.Id,
                $row.ProviderName,
                $row.MessageSummary
            )
        }
    }
    finally {
        $grid.ResumeLayout()
    }

    Show-EventDetail

    if ($script:CurrentRows.Count -eq 0) {
        # 条件に一致するイベントが無いだけなので、エラーとして扱わない
        $statusLabel.Text = "0 件（該当するイベントはありません）"
    }
    else {
        $statusLabel.Text = "$($script:CurrentRows.Count) 件（最大 $($script:MaxEvents) 件）"
    }
}

# -----------------------------------------
# イベントハンドラ
# -----------------------------------------

$btnSearch.Add_Click({
    Search-Events
})

$btnExport.Add_Click({
    if ($script:CurrentRows.Count -eq 0) {
        Show-Dialog `
            -Text "出力する結果がありません。" `
            -Caption "確認" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $fileName = "WinEvtView_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    $path = Join-Path ([Environment]::GetFolderPath("Desktop")) $fileName

    try {
        $script:CurrentRows |
            Select-Object @(
                @{ Name = "時刻";       Expression = { $_.TimeText } },
                @{ Name = "ログ";       Expression = { $_.LogName } },
                @{ Name = "レベル";     Expression = { $_.Level } },
                @{ Name = "ID";         Expression = { $_.Id } },
                @{ Name = "ソース";     Expression = { $_.ProviderName } },
                @{ Name = "メッセージ"; Expression = { $_.MessageLine } }
            ) |
            Export-Csv -Path $path -NoTypeInformation -Encoding UTF8

        Show-Dialog `
            -Text "CSVを出力しました。`r`n`r`n$path" `
            -Caption "出力完了" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Show-Dialog `
            -Text "CSV出力に失敗しました。`r`n`r`n$($_.Exception.Message)" `
            -Caption "エラー" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnViewer.Add_Click({
    try {
        Start-Process eventvwr.msc
    }
    catch {
        Show-Dialog `
            -Text "イベントビューアを開けませんでした。`r`n`r`n$($_.Exception.Message)" `
            -Caption "エラー" `
            -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$grid.Add_SelectionChanged({
    Show-EventDetail
})

$comboLog.Add_SelectedIndexChanged({
    Search-Events
})

$comboPeriod.Add_SelectedIndexChanged({
    Search-Events
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
    $grid,
    $detail,
    $statusStrip
))

# Enterキーでも検索できるようにする
$form.AcceptButton = $btnSearch

# 起動時に初期条件（System + Application / 直近1時間）で検索する
Search-Events

[void]$form.ShowDialog()
