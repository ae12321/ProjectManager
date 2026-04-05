using module .\Database.psm1
# ---------------------------------------------------------------------------
# ProjectManager.ps1
# ---------------------------------------------------------------------------

$script:CurrentDirPath = Split-Path -parent $MyInvocation.MyCommand.Path
$script:DatabaseName   = (Split-Path -leaf $MyInvocation.MyCommand.Path) -replace '\.ps1$','.db'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
$script:DatabasePath = Join-Path $script:CurrentDirPath $script:DatabaseName
$script:AttachmentPath = Join-Path $script:CurrentDirPath "attachment"
if (-not (Test-Path $script:AttachmentPath)) {
    New-Item $script:AttachmentPath -ItemType Directory -Force | Out-Null
}

$db = [Database]::new($script:DatabasePath)

function New-Id  { return [System.Guid]::NewGuid().ToString() }

function Get-Now { return (Get-Date).ToString("o") }

$db.executeNonQuery(@"
CREATE TABLE IF NOT EXISTS Project (
    id    TEXT PRIMARY KEY,
    title TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS Task (
    id        TEXT PRIMARY KEY,
    projectId TEXT NOT NULL,
    title     TEXT NOT NULL,
    priority  TEXT NOT NULL DEFAULT '中',
    deadline  TEXT,
    detail    TEXT,
    updatedAt TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS Memo (
    id        TEXT PRIMARY KEY,
    projectId TEXT NOT NULL,
    title     TEXT NOT NULL,
    detail    TEXT,
    pinned    INTEGER NOT NULL DEFAULT 0,
    updatedAt TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS Shortcut (
    id        TEXT PRIMARY KEY,
    projectId TEXT NOT NULL,
    name      TEXT NOT NULL,
    path      TEXT NOT NULL,
    updatedAt TEXT NOT NULL
);
"@)

# 既存DBにpinnedカラムを追加（安全に）
try { $db.executeNonQuery("ALTER TABLE Memo ADD COLUMN pinned INTEGER NOT NULL DEFAULT 0") } catch {}

# 設定ファイル（JSON）による永続化
$script:ConfigPath = Join-Path $script:CurrentDirPath ($script:DatabaseName -replace '\.db$', '.config.json')

function Get-Config ($key) {
    if (-not (Test-Path $script:ConfigPath)) { return $null }
    try {
        $json = Get-Content $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        return $json.$key
    } catch { return $null }
}

function Set-Config ($key, $value) {
    $json = if (Test-Path $script:ConfigPath) {
        try { Get-Content $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json }
        catch { [PSCustomObject]@{} }
    } else { [PSCustomObject]@{} }
    $json | Add-Member -MemberType NoteProperty -Name $key -Value $value -Force
    $json | ConvertTo-Json -Compress | Set-Content $script:ConfigPath -Encoding UTF8
}

# プロジェクトの保存・変更
function Save-Project ($proj) {
    $db.executeNonQuery("INSERT OR REPLACE INTO Project(id,title) VALUES(@id,@title)", @{
        "@id" = $proj.id; "@title" = $proj.title
    })
}

# プロジェクトの削除
function Delete-Project ($projId) {
    $memos = $db.executeQuery("SELECT id FROM Memo WHERE projectId=@pid", @{ "@pid" = $projId })
    $memoDirs = $memos | ForEach-Object { Join-Path $script:AttachmentPath $_.id }
    $db.BeginTransaction()
    try {
        $db.executeNonQuery("DELETE FROM Shortcut WHERE projectId=@pid", @{ "@pid" = $projId })
        $db.executeNonQuery("DELETE FROM Memo    WHERE projectId=@pid", @{ "@pid" = $projId })
        $db.executeNonQuery("DELETE FROM Task    WHERE projectId=@pid", @{ "@pid" = $projId })
        $db.executeNonQuery("DELETE FROM Project WHERE id=@id",        @{ "@id"  = $projId })
        $db.Commit()
        foreach ($dir in $memoDirs) {
            if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
        }
    } catch {
        $db.Rollback()
        throw
    }
}

# タスクの保存・変更
function Save-Task ($task, $projectId) {
    $db.executeNonQuery(@"
INSERT OR REPLACE INTO Task(id,projectId,title,priority,deadline,detail,updatedAt)
VALUES(@id,@pid,@title,@priority,@deadline,@detail,@updatedAt)
"@, @{
        "@id"        = $task.id
        "@pid"       = $projectId
        "@title"     = $task.title
        "@priority"  = if ($task.priority) { $task.priority } else { "中" }
        "@deadline"  = if ($task.deadline) { $task.deadline } else { "" }
        "@detail"    = if ($task.detail)   { $task.detail   } else { "" }
        "@updatedAt" = $task.updatedAt
    })
}

# タスクの削除
function Delete-Task ($taskId) {
    $db.executeNonQuery("DELETE FROM Task WHERE id=@id", @{ "@id" = $taskId })
}

# メモの保存・変更
function Save-Memo ($memo, $projectId) {
    $db.executeNonQuery(@"
INSERT OR REPLACE INTO Memo(id,projectId,title,detail,pinned,updatedAt)
VALUES(@id,@pid,@title,@detail,@pinned,@updatedAt)
"@, @{
        "@id"        = $memo.id
        "@pid"       = $projectId
        "@title"     = $memo.title
        "@detail"    = if ($memo.detail) { $memo.detail } else { "" }
        "@pinned"    = if ($memo.pinned) { 1 } else { 0 }
        "@updatedAt" = $memo.updatedAt
    })
}

# メモの削除
function Delete-Memo ($memoId) {
    $db.executeNonQuery("DELETE FROM Memo WHERE id=@id", @{ "@id" = $memoId })
    $attDir = Join-Path $script:AttachmentPath $memoId
    if (Test-Path $attDir) { Remove-Item $attDir -Recurse -Force }
}

# ショートカットの保存・変更
function Save-Shortcut ($sc, $projectId) {
    $db.executeNonQuery(@"
INSERT OR REPLACE INTO Shortcut(id,projectId,name,path,updatedAt)
VALUES(@id,@pid,@name,@path,@updatedAt)
"@, @{
        "@id"        = $sc.id
        "@pid"       = $projectId
        "@name"      = $sc.name
        "@path"      = $sc.path
        "@updatedAt" = $sc.updatedAt
    })
}

# ショートカットの削除
function Delete-Shortcut ($scId) {
    $db.executeNonQuery("DELETE FROM Shortcut WHERE id=@id", @{ "@id" = $scId })
}

# 1プロジェクト分をDBから読み込む
function Load-Project ($projId) {
    $rows = $db.executeQuery("SELECT id,title FROM Project WHERE id=@id", @{ "@id" = $projId })
    if ($rows.Count -eq 0) { return $null }
    $p = $rows[0]

    $tasks = [System.Collections.Generic.List[object]]::new()
    $taskRows = $db.executeQuery("SELECT * FROM Task WHERE projectId=@pid", @{ "@pid" = $projId })
    foreach ($t in $taskRows) {
        $tasks.Add([ordered]@{
            id        = $t.id
            title     = $t.title
            priority  = if ($t.priority) { $t.priority } else { "中" }
            deadline  = if ($t.deadline)  { $t.deadline  } else { "" }
            detail    = if ($t.detail)    { $t.detail    } else { "" }
            updatedAt = $t.updatedAt
        }) | Out-Null
    }

    $memos = [System.Collections.Generic.List[object]]::new()
    $memoRows = $db.executeQuery("SELECT * FROM Memo WHERE projectId=@pid", @{ "@pid" = $projId })
    foreach ($m in $memoRows) {
        $memos.Add([ordered]@{
            id        = $m.id
            title     = $m.title
            detail    = $m.detail
            pinned    = if ($m.pinned) { [int]$m.pinned } else { 0 }
            updatedAt = $m.updatedAt
        }) | Out-Null
    }

    $shortcuts = [System.Collections.Generic.List[object]]::new()
    $scRows = $db.executeQuery("SELECT * FROM Shortcut WHERE projectId=@pid", @{ "@pid" = $projId })
    foreach ($sc in $scRows) {
        $shortcuts.Add([ordered]@{
            id        = $sc.id
            name      = $sc.name
            path      = $sc.path
            updatedAt = $sc.updatedAt
        }) | Out-Null
    }

    return [ordered]@{ id = $p.id; title = $p.title; tasks = $tasks; memos = $memos; shortcuts = $shortcuts }
}

# $script:Projects の該当エントリをDBから再読み込みして更新する
function Reload-Project ($projId) {
    for ($i = 0; $i -lt $script:Projects.Count; $i++) {
        if ($script:Projects[$i].id -eq $projId) {
            $updated = Load-Project $projId
            if ($null -ne $updated) { $script:Projects[$i] = $updated }
            return
        }
    }
}

# 全プロジェクトをDBから読み込む
function Load-AllData {
    $projects = [System.Collections.Generic.List[object]]::new()
    foreach ($p in $db.executeQuery("SELECT id FROM Project")) {
        $loaded = Load-Project $p.id
        if ($null -ne $loaded) { $projects.Add($loaded) | Out-Null }
    }
    return $projects
}

$script:Projects         = [System.Collections.Generic.List[object]]::new([object[]]@(Load-AllData))
$script:CurrentProjectId = ""

# 画面要素の定義
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Project Manager" Height="660" Width="960"
        WindowStartupLocation="CenterScreen"
        Background="#F0F2F5"
        FontFamily="BIZ UIGothic">
  <Window.Resources>
    <Style TargetType="Button" x:Key="PrimaryBtn">
      <Setter Property="Background" Value="#4A90D9"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="13"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#357ABD"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="Button" x:Key="DangerBtn">
      <Setter Property="Background" Value="#E05252"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="13"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#B94040"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="Button" x:Key="SecondaryBtn">
      <Setter Property="Background" Value="#7C8591"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,4"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="13"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#5A626B"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <DockPanel Margin="4">

    <!-- ===== TOP: Project bar ===== -->
    <Border DockPanel.Dock="Top" Background="White"
            BorderBrush="#D0D5DD" BorderThickness="1" Margin="0,0,0,8" Padding="8,8">
      <DockPanel>
        <TextBlock DockPanel.Dock="Left" Text="プロジェクト："
                   FontSize="13" FontWeight="Bold" Foreground="#1A3A5C"
                   VerticalAlignment="Center" Margin="0,0,8,0"/>
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="BtnAddProject"    Content="＋ 追加" Style="{StaticResource PrimaryBtn}"   Margin="0,0,4,0"/>
          <Button x:Name="BtnRenameProject" Content="編集"    Style="{StaticResource SecondaryBtn}" Margin="0,0,4,0"/>
          <Button x:Name="BtnDeleteProject" Content="削除"    Style="{StaticResource DangerBtn}"/>
        </StackPanel>
        <ComboBox x:Name="LstProjects" FontSize="13" Padding="6,4"
                  VerticalAlignment="Center" MinWidth="200" HorizontalAlignment="Left"/>
      </DockPanel>
    </Border>

    <!-- ===== MAIN content ===== -->
    <Border Background="White" BorderBrush="#D0D5DD" BorderThickness="1">
      <Grid x:Name="RightPanel">
        <TextBlock x:Name="PlaceholderText"
                   Text="&#x2191; プロジェクトを選択してください"
                   HorizontalAlignment="Center" VerticalAlignment="Center"
                   FontSize="14" Foreground="#999" Visibility="Visible"/>

        <DockPanel x:Name="ContentPanel" Visibility="Collapsed">
          <TabControl x:Name="TabMain" Background="Transparent"
                      BorderThickness="0,1,0,0" BorderBrush="#D0D5DD">

            <!-- ══ MEMO TAB ══ -->
            <TabItem FontSize="13" Width="80">
              <TabItem.Header>
                <TextBlock Text="メモ" Width="70" TextAlignment="Center"/>
              </TabItem.Header>
              <DockPanel Margin="8">

                <!-- Search bar -->
                <Border DockPanel.Dock="Top" Margin="0,0,0,8">
                  <DockPanel>

                    <Button x:Name="BtnAddMemo" DockPanel.Dock="Left"
                            Content="＋ 追加" Style="{StaticResource PrimaryBtn}" Margin="0,0,10,0"/>
                    <TextBlock DockPanel.Dock="Left" Text="検索："
                               FontSize="13" VerticalAlignment="Center" Margin="0,0,6,0"/>
                    <Button x:Name="BtnClearSearch" DockPanel.Dock="Right"
                            Content="✕" Width="28" FontSize="12"
                            Background="#DDDDDD" Foreground="#444444"
                            BorderThickness="1" BorderBrush="#BBBBBB"
                            Cursor="Hand" VerticalAlignment="Center" Margin="4,0,0,0"/>
                    <TextBox x:Name="TxtSearch" FontSize="13" Padding="4,3"
                             VerticalAlignment="Center"
                             BorderBrush="#B0B8C4" BorderThickness="1"/>
                  </DockPanel>
                </Border>

                <!-- Memo list | detail -->
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="200" MinWidth="80"/>
                    <ColumnDefinition Width="5"/>
                    <ColumnDefinition Width="*" MinWidth="200"/>
                  </Grid.ColumnDefinitions>

                  <!-- メモ一覧 + ボタン -->
                  <DockPanel Grid.Column="0">

                    <ListBox x:Name="LstMemos"
                             BorderThickness="1" BorderBrush="#D0D5DD"
                             FontSize="13" DisplayMemberPath="display"/>
                  </DockPanel>

                  <!-- GridSplitter: メモ一覧と詳細の境界 -->
                  <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch"
                                VerticalAlignment="Stretch" Background="#D0D5DD"
                                Cursor="SizeWE"/>

                  <!-- メモ詳細 -->
                  <DockPanel Grid.Column="2">
                    <!-- 詳細ヘッダー: タイトル + 操作ボタン -->
                    <Border DockPanel.Dock="Top" Margin="0,0,0,4" Padding="0,0,0,0">
                      <DockPanel LastChildFill="True" Margin="2,0,0,0">
                        <StackPanel x:Name="SpMemoReadButtons" DockPanel.Dock="Right"
                                    Orientation="Horizontal" Visibility="Collapsed">
                          <Button x:Name="BtnMemoPin"    Content="📌" Width="36" FontSize="12"
                                  Background="#A0A8B0" Foreground="White" BorderThickness="0"
                                  Padding="0,4" Cursor="Hand" Margin="0,0,4,0"/>
                          <Button x:Name="BtnMemoEdit"   Content="編集" Width="52" FontSize="12"
                                  Background="#4A90D9" Foreground="White" BorderThickness="0"
                                  Padding="0,4" Cursor="Hand" Margin="0,0,4,0"/>
                          <Button x:Name="BtnMemoDelete" Content="削除" Width="52" FontSize="12"
                                  Background="#E05252" Foreground="White" BorderThickness="0"
                                  Padding="0,4" Cursor="Hand"/>
                        </StackPanel>
                        <StackPanel x:Name="SpMemoEditButtons" DockPanel.Dock="Right"
                                    Orientation="Horizontal" Visibility="Collapsed">
                          <Button x:Name="BtnMemoSave"   Content="保存"       Width="60" FontSize="12"
                                  Background="#4A90D9" Foreground="White" BorderThickness="0"
                                  Padding="0,4" Cursor="Hand" Margin="0,0,4,0"/>
                          <Button x:Name="BtnMemoCancel" Content="キャンセル" Width="76" FontSize="12"
                                  Background="#7C8591" Foreground="White" BorderThickness="0"
                                  Padding="0,4" Cursor="Hand"/>
                        </StackPanel>
                        <TextBlock x:Name="TxtMemoTitleDisp" FontSize="13" FontWeight="SemiBold"
                                   TextTrimming="CharacterEllipsis"
                                   Visibility="Collapsed" Margin="5,2,8,0"/>
                        <TextBox x:Name="TxtMemoTitleEdit" FontSize="13" Padding="4,3"
                                 BorderBrush="#B0B8C4" BorderThickness="1"
                                 Margin="0,0,8,0" Visibility="Collapsed"/>
                      </DockPanel>
                    </Border>
                    <!-- 添付ファイルパネル（下部固定） -->
                    <Border x:Name="PnlAttachments" DockPanel.Dock="Bottom"
                            BorderBrush="#D0D5DD" BorderThickness="1" CornerRadius="4"
                            Margin="0,6,0,0" Padding="6,4" Visibility="Collapsed">
                      <StackPanel x:Name="SpAttachments" Orientation="Vertical"/>
                    </Border>
                    <!-- コンテンツエリア: Read / Write 切替 -->
                    <Grid Margin="2,0,0,0">
                      <!-- Read mode: 本文表示 -->
                      <TextBox x:Name="RtbMemoDetail" IsReadOnly="True"
                               BorderBrush="#D0D5DD" BorderThickness="1"
                               Padding="2" FontSize="13"
                               Background="Transparent"
                               TextWrapping="Wrap"
                               VerticalScrollBarVisibility="Auto"
                               HorizontalScrollBarVisibility="Disabled"/>
                      <!-- Write mode: 詳細編集 -->
                      <TextBox x:Name="TxtMemoDetailEdit" Visibility="Collapsed"
                               FontSize="13" Padding="2" AcceptsReturn="True"
                               TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                               BorderBrush="#4A90D9" BorderThickness="2"/>
                    </Grid>
                  </DockPanel>
                </Grid>

              </DockPanel>
            </TabItem>

            <!-- ══ TASK TAB ══ -->
            <TabItem FontSize="13" Width="80">
              <TabItem.Header>
                <TextBlock Text="タスク" Width="70" TextAlignment="Center"/>
              </TabItem.Header>
              <DockPanel Margin="8">
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,8">
                  <Button x:Name="BtnAddTask" Content="＋ 追加" Style="{StaticResource PrimaryBtn}"/>
                </StackPanel>
                <ListView x:Name="LstTasks" BorderThickness="1" BorderBrush="#D0D5DD"
                          FontSize="13" SelectionMode="Single">
                  <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                      <Style.Triggers>
                        <DataTrigger Binding="{Binding deadlineColor}" Value="Red">
                          <Setter Property="Background" Value="#FFCCCC"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding deadlineColor}" Value="Yellow">
                          <Setter Property="Background" Value="#FFF8CC"/>
                        </DataTrigger>
                      </Style.Triggers>
                    </Style>
                  </ListView.ItemContainerStyle>
                  <ListView.View>
                    <GridView>
                      <GridViewColumn Header="削除" Width="40">
                        <GridViewColumn.CellTemplate>
                          <DataTemplate>
                            <Button Content="🗑" Tag="{Binding id}"
                                    Width="28" Height="22" FontSize="12"
                                    Background="Transparent" BorderThickness="0"
                                    Cursor="Hand" Name="BtnRowDelete" ToolTip="削除"/>
                          </DataTemplate>
                        </GridViewColumn.CellTemplate>
                      </GridViewColumn>

                      <GridViewColumn Header="優先度" Width="60"
                                      DisplayMemberBinding="{Binding priority}"/>
                      <GridViewColumn Header="タスク名" Width="200"
                                      DisplayMemberBinding="{Binding title}"/>
                      <GridViewColumn Header="詳細" Width="160">
                        <GridViewColumn.CellTemplate>
                          <DataTemplate>
                            <TextBlock Text="{Binding detailOneLine}"
                                       TextTrimming="CharacterEllipsis"
                                       TextWrapping="NoWrap"
                                       ToolTip="{Binding detail}"
                                       MaxWidth="150"/>
                          </DataTemplate>
                        </GridViewColumn.CellTemplate>
                      </GridViewColumn>
                      <GridViewColumn Header="期限" Width="90"
                                      DisplayMemberBinding="{Binding deadlineDisp}"/>
                    </GridView>
                  </ListView.View>
                </ListView>
              </DockPanel>
            </TabItem>

            <!-- ══ SHORTCUT TAB ══ -->
            <TabItem FontSize="13" Width="120">
              <TabItem.Header>
                <TextBlock Text="ショートカット" Width="110" TextAlignment="Center"/>
              </TabItem.Header>
              <DockPanel Margin="8">
                <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,8">
                  <Button x:Name="BtnAddShortcut" Content="＋ 追加" Style="{StaticResource PrimaryBtn}"/>
                </StackPanel>
                <ListView x:Name="LstShortcuts" BorderThickness="1" BorderBrush="#D0D5DD"
                          FontSize="13" SelectionMode="Single">
                  <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem">
                      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                    </Style>
                  </ListView.ItemContainerStyle>
                  <ListView.View>
                    <GridView>
                      <GridViewColumn Header="削除" Width="40">
                        <GridViewColumn.CellTemplate>
                          <DataTemplate>
                            <Button Content="🗑" Tag="{Binding id}"
                                    Width="28" Height="22" FontSize="12"
                                    Background="Transparent" BorderThickness="0"
                                    Cursor="Hand" Name="BtnScDelete" ToolTip="削除"/>
                          </DataTemplate>
                        </GridViewColumn.CellTemplate>
                      </GridViewColumn>
                      <GridViewColumn Header="編集" Width="40">
                        <GridViewColumn.CellTemplate>
                          <DataTemplate>
                            <Button Content="🖊" Tag="{Binding id}"
                                    Width="28" Height="22" FontSize="12"
                                    Background="Transparent" BorderThickness="0"
                                    Cursor="Hand" Name="BtnScEdit" ToolTip="編集"/>
                          </DataTemplate>
                        </GridViewColumn.CellTemplate>
                      </GridViewColumn>
                      <GridViewColumn Header="名称" Width="160"
                                      DisplayMemberBinding="{Binding name}"/>
                      <GridViewColumn Header="パス / URL" Width="300"
                                      DisplayMemberBinding="{Binding path}"/>
                    </GridView>
                  </ListView.View>
                </ListView>
              </DockPanel>
            </TabItem>

          </TabControl>
        </DockPanel>
      </Grid>
    </Border>
  </DockPanel>
</Window>
'@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

function Ctrl($n) { $window.FindName($n) }

$LstProjects      = Ctrl "LstProjects"
$BtnAddProject    = Ctrl "BtnAddProject"
$BtnRenameProject = Ctrl "BtnRenameProject"
$BtnDeleteProject = Ctrl "BtnDeleteProject"
$PlaceholderText  = Ctrl "PlaceholderText"
$ContentPanel     = Ctrl "ContentPanel"
$TabMain          = Ctrl "TabMain"
$LstTasks         = Ctrl "LstTasks"
$BtnAddTask       = Ctrl "BtnAddTask"
$LstMemos         = Ctrl "LstMemos"
$BtnAddMemo       = Ctrl "BtnAddMemo"
$LstShortcuts     = Ctrl "LstShortcuts"
$BtnAddShortcut   = Ctrl "BtnAddShortcut"

$RtbMemoDetail    = Ctrl "RtbMemoDetail"
$PnlAttachments   = Ctrl "PnlAttachments"
$SpAttachments    = Ctrl "SpAttachments"
$TxtSearch        = Ctrl "TxtSearch"
$BtnClearSearch   = Ctrl "BtnClearSearch"

$SpMemoReadButtons  = Ctrl "SpMemoReadButtons"
$SpMemoEditButtons  = Ctrl "SpMemoEditButtons"
$BtnMemoPin         = Ctrl "BtnMemoPin"
$BtnMemoEdit        = Ctrl "BtnMemoEdit"
$BtnMemoDelete      = Ctrl "BtnMemoDelete"
$BtnMemoSave        = Ctrl "BtnMemoSave"
$BtnMemoCancel      = Ctrl "BtnMemoCancel"
$TxtMemoTitleDisp   = Ctrl "TxtMemoTitleDisp"
$TxtMemoTitleEdit   = Ctrl "TxtMemoTitleEdit"
$TxtMemoDetailEdit  = Ctrl "TxtMemoDetailEdit"

# プロジェクト用ダイアログ
function Show-InputDialog ($title, $label, $default = "") {
    $dlg = New-Object System.Windows.Window
    $dlg.Title = $title; $dlg.Width = 420; $dlg.Height = 155
    $dlg.WindowStartupLocation = "CenterOwner"; $dlg.Owner = $window
    $dlg.ResizeMode = "NoResize"
    $dlg.Background = [System.Windows.Media.Brushes]::WhiteSmoke

    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = "14"

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $label; $lbl.FontSize = 13; $lbl.Margin = "0,0,0,6"
    $sp.Children.Add($lbl) | Out-Null

    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Text = $default; $tb.FontSize = 13; $tb.Padding = "4"; $tb.Margin = "0,0,0,12"
    $sp.Children.Add($tb) | Out-Null

    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = "Horizontal"; $row.HorizontalAlignment = "Right"

    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = "OK"; $ok.Width = 70; $ok.Margin = "0,0,8,0"; $ok.IsDefault = $true
    $ok.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $row.Children.Add($ok) | Out-Null

    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "キャンセル"; $cancel.Width = 90; $cancel.IsCancel = $true
    $cancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $row.Children.Add($cancel) | Out-Null

    $sp.Children.Add($row) | Out-Null
    $dlg.Content = $sp
    $tb.Focus() | Out-Null; $tb.SelectAll()

    if ($dlg.ShowDialog() -eq $true -and $tb.Text.Trim() -ne "") { return $tb.Text.Trim() }
    return $null
}

# タスク用ダイアログ
function Show-TaskDialog ($title, $initTitle = "", $initPriority = "中", $initDeadline = "", $initDetail = "") {
    $dlg = New-Object System.Windows.Window
    $dlg.Title = $title; $dlg.Width = 420; $dlg.Height = 490
    $dlg.WindowStartupLocation = "CenterOwner"; $dlg.Owner = $window
    $dlg.ResizeMode = "NoResize"
    $dlg.Background = [System.Windows.Media.Brushes]::WhiteSmoke

    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = "14"

    # タスク名
    $lblN = New-Object System.Windows.Controls.TextBlock
    $lblN.Text = "タスク名"; $lblN.FontSize = 13; $lblN.Margin = "0,0,0,4"
    $sp.Children.Add($lblN) | Out-Null
    $tbTitle = New-Object System.Windows.Controls.TextBox
    $tbTitle.Text = $initTitle; $tbTitle.FontSize = 13; $tbTitle.Padding = "4"; $tbTitle.Margin = "0,0,0,10"
    $sp.Children.Add($tbTitle) | Out-Null

    # 優先度
    $rowP = New-Object System.Windows.Controls.StackPanel
    $rowP.Orientation = "Horizontal"; $rowP.Margin = "0,0,0,10"
    $lblP = New-Object System.Windows.Controls.TextBlock
    $lblP.Text = "優先度："; $lblP.FontSize = 13; $lblP.VerticalAlignment = "Center"; $lblP.Margin = "0,0,6,0"
    $rowP.Children.Add($lblP) | Out-Null
    $cbPriority = New-Object System.Windows.Controls.ComboBox
    $cbPriority.FontSize = 13; $cbPriority.Width = 70
    @("大","中","小") | ForEach-Object { $cbPriority.Items.Add($_) | Out-Null }
    $cbPriority.SelectedItem = $initPriority
    $rowP.Children.Add($cbPriority) | Out-Null
    $sp.Children.Add($rowP) | Out-Null

    # 期限ラベル + クリアボタン
    $rowDL = New-Object System.Windows.Controls.DockPanel; $rowDL.Margin = "0,0,0,4"
    $lblD = New-Object System.Windows.Controls.TextBlock
    $lblD.Text = "期限（任意）："; $lblD.FontSize = 13; $lblD.VerticalAlignment = "Center"
    $lblD.SetValue([System.Windows.Controls.DockPanel]::DockProperty, [System.Windows.Controls.Dock]::Left)
    $rowDL.Children.Add($lblD) | Out-Null

    $btnClearDL = New-Object System.Windows.Controls.Button
    $btnClearDL.Content = "クリア"; $btnClearDL.FontSize = 12; $btnClearDL.Padding = "6,2"
    $btnClearDL.Background = [System.Windows.Media.Brushes]::LightGray
    $btnClearDL.BorderThickness = 1; $btnClearDL.Cursor = "Hand"; $btnClearDL.Margin = "8,0,0,0"
    $btnClearDL.SetValue([System.Windows.Controls.DockPanel]::DockProperty, [System.Windows.Controls.Dock]::Left)
    $rowDL.Children.Add($btnClearDL) | Out-Null
    $sp.Children.Add($rowDL) | Out-Null

    # カレンダー
    $cal = New-Object System.Windows.Controls.Calendar
    $cal.SelectionMode = "SingleDate"
    $cal.Margin = "0,0,0,10"
    $cal.FontSize = 12
    # 初期値セット
    if ($initDeadline -ne "") {
        try {
            $initDt = [datetime]::ParseExact($initDeadline, "yyyy/MM/dd", $null)
            $cal.SelectedDate = $initDt
            $cal.DisplayDate  = $initDt
        } catch {}
    }
    $sp.Children.Add($cal) | Out-Null

    # クリアボタン動作
    $btnClearDL.Add_Click({ $cal.SelectedDate = $null })

    # 詳細
    $lblDet = New-Object System.Windows.Controls.TextBlock
    $lblDet.Text = "詳細（任意）"; $lblDet.FontSize = 13; $lblDet.Margin = "0,0,0,4"
    $sp.Children.Add($lblDet) | Out-Null
    $tbDetail = New-Object System.Windows.Controls.TextBox
    $tbDetail.Text = $initDetail; $tbDetail.FontSize = 13; $tbDetail.Padding = "4"
    $tbDetail.TextWrapping = "Wrap"; $tbDetail.AcceptsReturn = $true
    $tbDetail.Height = 70; $tbDetail.VerticalScrollBarVisibility = "Auto"
    $tbDetail.Margin = "0,0,0,10"
    $sp.Children.Add($tbDetail) | Out-Null

    # OK / キャンセル
    $rowBtn = New-Object System.Windows.Controls.StackPanel
    $rowBtn.Orientation = "Horizontal"; $rowBtn.HorizontalAlignment = "Right"
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = "OK"; $ok.Width = 70; $ok.Margin = "0,0,8,0"; $ok.IsDefault = $true
    $ok.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $rowBtn.Children.Add($ok) | Out-Null
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "キャンセル"; $cancel.Width = 90; $cancel.IsCancel = $true
    $cancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $rowBtn.Children.Add($cancel) | Out-Null
    $sp.Children.Add($rowBtn) | Out-Null

    $dlg.Content = $sp; $tbTitle.Focus() | Out-Null; $tbTitle.SelectAll()

    if ($dlg.ShowDialog() -eq $true -and $tbTitle.Text.Trim() -ne "") {
        $dlStr = ""
        if ($null -ne $cal.SelectedDate) {
            $dlStr = ([datetime]$cal.SelectedDate).ToString("yyyy/MM/dd")
        }
        return @{
            title    = $tbTitle.Text.Trim()
            priority = if ($cbPriority.SelectedItem) { $cbPriority.SelectedItem } else { "中" }
            deadline = $dlStr
            detail   = $tbDetail.Text
        }
    }
    return $null
}

# ショートカット用ダイアログ
function Show-ShortcutDialog ($title, $initName = "", $initPath = "") {
    $dlg = New-Object System.Windows.Window
    $dlg.Title = $title; $dlg.Width = 460; $dlg.Height = 220
    $dlg.WindowStartupLocation = "CenterOwner"; $dlg.Owner = $window
    $dlg.ResizeMode = "NoResize"
    $dlg.Background = [System.Windows.Media.Brushes]::WhiteSmoke

    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = "14"

    # 名称
    $lblN = New-Object System.Windows.Controls.TextBlock
    $lblN.Text = "名称"; $lblN.FontSize = 13; $lblN.Margin = "0,0,0,4"
    $sp.Children.Add($lblN) | Out-Null
    $tbName = New-Object System.Windows.Controls.TextBox
    $tbName.Text = $initName; $tbName.FontSize = 13; $tbName.Padding = "4"; $tbName.Margin = "0,0,0,10"
    $sp.Children.Add($tbName) | Out-Null

    # パス / URL
    $lblP = New-Object System.Windows.Controls.TextBlock
    $lblP.Text = "パス / URL"; $lblP.FontSize = 13; $lblP.Margin = "0,0,0,4"
    $sp.Children.Add($lblP) | Out-Null

    $rowPath = New-Object System.Windows.Controls.DockPanel; $rowPath.Margin = "0,0,0,14"
    $btnsBrowse = New-Object System.Windows.Controls.StackPanel
    $btnsBrowse.Orientation = "Horizontal"
    $btnsBrowse.SetValue([System.Windows.Controls.DockPanel]::DockProperty, [System.Windows.Controls.Dock]::Right)

    $btnFolder = New-Object System.Windows.Controls.Button
    $btnFolder.Content = "フォルダ…"; $btnFolder.FontSize = 12; $btnFolder.Padding = "6,4"
    $btnFolder.Margin = "4,0,0,0"
    $btnFile = New-Object System.Windows.Controls.Button
    $btnFile.Content = "ファイル…"; $btnFile.FontSize = 12; $btnFile.Padding = "6,4"
    $btnFile.Margin = "4,0,0,0"
    $btnsBrowse.Children.Add($btnFolder) | Out-Null
    $btnsBrowse.Children.Add($btnFile)   | Out-Null

    $tbPath = New-Object System.Windows.Controls.TextBox
    $tbPath.Text = $initPath; $tbPath.FontSize = 13; $tbPath.Padding = "4"
    $rowPath.Children.Add($btnsBrowse) | Out-Null
    $rowPath.Children.Add($tbPath)     | Out-Null
    $sp.Children.Add($rowPath) | Out-Null

    # フォルダ選択
    $btnFolder.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = "フォルダを選択してください"
        if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbPath.Text = $folderDlg.SelectedPath
            if ($tbName.Text.Trim() -eq "") { $tbName.Text = Split-Path $folderDlg.SelectedPath -Leaf }
        }
    }.GetNewClosure())

    # ファイル選択
    $btnFile.Add_Click({
        $fileDlg = New-Object System.Windows.Forms.OpenFileDialog
        $fileDlg.Title = "ファイルを選択してください"
        if ($fileDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbPath.Text = $fileDlg.FileName
            if ($tbName.Text.Trim() -eq "") {
                $tbName.Text = [System.IO.Path]::GetFileNameWithoutExtension($fileDlg.FileName)
            }
        }
    }.GetNewClosure())

    # OK / キャンセル
    $rowBtn = New-Object System.Windows.Controls.StackPanel
    $rowBtn.Orientation = "Horizontal"; $rowBtn.HorizontalAlignment = "Right"
    $ok = New-Object System.Windows.Controls.Button
    $ok.Content = "OK"; $ok.Width = 70; $ok.Margin = "0,0,8,0"; $ok.IsDefault = $true
    $ok.Add_Click({ $dlg.DialogResult = $true; $dlg.Close() })
    $rowBtn.Children.Add($ok) | Out-Null
    $cancel = New-Object System.Windows.Controls.Button
    $cancel.Content = "キャンセル"; $cancel.Width = 90; $cancel.IsCancel = $true
    $cancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    $rowBtn.Children.Add($cancel) | Out-Null
    $sp.Children.Add($rowBtn) | Out-Null

    $dlg.Content = $sp; $tbName.Focus() | Out-Null

    if ($dlg.ShowDialog() -eq $true -and $tbName.Text.Trim() -ne "" -and $tbPath.Text.Trim() -ne "") {
        return @{ name = $tbName.Text.Trim(); path = $tbPath.Text.Trim() }
    }
    return $null
}


# メモの詳細を表示（添付ファイルは下部にリスト表示）
function Render-MemoDetail ($text, $memoId = "") {
    $RtbMemoDetail.Text = if ($text) { $text } else { "" }
    $SpAttachments.Children.Clear()
    $PnlAttachments.Visibility = "Collapsed"

    # 添付ファイル（下部パネルに表示）
    $attDir = Join-Path $script:AttachmentPath $memoId
    $atts = if (-not [string]::IsNullOrEmpty($memoId) -and (Test-Path $attDir)) {
        @(Get-ChildItem $attDir -File | Select-Object -ExpandProperty Name)
    } else { @() }

    if (-not [string]::IsNullOrEmpty($memoId)) {
        $PnlAttachments.Visibility = "Visible"
        $header = New-Object System.Windows.Controls.TextBlock
        $header.Text = "添付ファイル"; $header.FontSize = 11
        $header.Foreground = [System.Windows.Media.Brushes]::Gray
        $header.Margin = [System.Windows.Thickness]::new(0,0,0,4)
        $SpAttachments.Children.Add($header) | Out-Null

        if ($atts.Count -gt 0) {
            $attFolderPath = Join-Path $script:AttachmentPath $memoId

            $attBorder = New-Object System.Windows.Controls.Border
            $attBorder.Cursor       = [System.Windows.Input.Cursors]::Hand
            $attBorder.Padding      = [System.Windows.Thickness]::new(6, 4, 6, 4)
            $attBorder.CornerRadius = [System.Windows.CornerRadius]::new(3)
            $attBorder.Background   = [System.Windows.Media.Brushes]::Transparent

            $attBorder.Add_MouseEnter({
                $attBorder.Background = [System.Windows.Media.SolidColorBrush](
                    [System.Windows.Media.Color]::FromArgb(20, 0, 0, 0))
            }.GetNewClosure())
            $attBorder.Add_MouseLeave({
                $attBorder.Background = [System.Windows.Media.Brushes]::Transparent
            }.GetNewClosure())
            $attBorder.Add_MouseLeftButtonUp({
                param($s, $e); $e.Handled = $true
                if (Test-Path $attFolderPath) {
                    Start-Process "explorer.exe" $attFolderPath
                } else {
                    [System.Windows.MessageBox]::Show("フォルダが見つかりません：$attFolderPath")
                }
            }.GetNewClosure())

            $sv = New-Object System.Windows.Controls.ScrollViewer
            $sv.VerticalScrollBarVisibility   = "Auto"
            $sv.HorizontalScrollBarVisibility = "Disabled"
            if ($atts.Count -ge 4) { $sv.MaxHeight = 66 }

            $sp = New-Object System.Windows.Controls.StackPanel
            foreach ($fileName in $atts) {
                $tb = New-Object System.Windows.Controls.TextBlock
                $tb.Text = "📎 $fileName"; $tb.FontSize = 12
                $sp.Children.Add($tb) | Out-Null
            }
            $sv.Content = $sp
            $attBorder.Child = $sv
            $SpAttachments.Children.Add($attBorder) | Out-Null
        } else {
            # 添付フォルダを開くボタン領域
            $emptyBorder = New-Object System.Windows.Controls.Border
            $emptyBorder.Cursor  = [System.Windows.Input.Cursors]::Hand
            $emptyBorder.Padding = [System.Windows.Thickness]::new(6, 4, 6, 4)
            $emptyBorder.CornerRadius = [System.Windows.CornerRadius]::new(3)
            $emptyBorder.Background = [System.Windows.Media.Brushes]::Transparent

            $emptyBorder.Add_MouseEnter({
                $emptyBorder.Background = [System.Windows.Media.SolidColorBrush](
                    [System.Windows.Media.Color]::FromArgb(20, 0, 0, 0))
            }.GetNewClosure())
            $emptyBorder.Add_MouseLeave({
                $emptyBorder.Background = [System.Windows.Media.Brushes]::Transparent
            }.GetNewClosure())

            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = "📎 フォルダを開く"
            $lbl.FontSize = 12
            $lbl.Foreground = [System.Windows.Media.Brushes]::LightGray
            $lbl.HorizontalAlignment = "Center"
            $lbl.VerticalAlignment = "Center"
            $emptyBorder.Child = $lbl

            $capMemoId = $memoId
            $capAttPath = $script:AttachmentPath
            $emptyBorder.Add_MouseLeftButtonUp({
                param($s, $e); $e.Handled = $true
                $folderPath = Join-Path $capAttPath $capMemoId
                if (-not (Test-Path $folderPath)) { New-Item -ItemType Directory -Path $folderPath | Out-Null }
                Start-Process "explorer.exe" $folderPath
            }.GetNewClosure())

            $SpAttachments.Children.Add($emptyBorder) | Out-Null
        }
    }
}

# ---------------------------------------------------------------------------
# メモ詳細パネルのモード制御
# ---------------------------------------------------------------------------

function Set-MemoEditMode {
    $SpMemoReadButtons.Visibility = "Collapsed"
    $SpMemoEditButtons.Visibility = "Visible"
    $TxtMemoTitleDisp.Visibility  = "Collapsed"
    $TxtMemoTitleEdit.Visibility  = "Visible"
    $RtbMemoDetail.Visibility     = "Collapsed"
    $TxtMemoDetailEdit.Visibility = "Visible"
}

# item が null → ヘッダー非表示（未選択状態）
# item が非null → Read モードでヘッダー表示
function Set-MemoDetailHeader ($item) {
    if ($null -eq $item) {
        $SpMemoReadButtons.Visibility = "Collapsed"
        $SpMemoEditButtons.Visibility = "Collapsed"
        $TxtMemoTitleDisp.Visibility  = "Collapsed"
        $TxtMemoTitleEdit.Visibility  = "Collapsed"
        $RtbMemoDetail.Visibility     = "Visible"
        $TxtMemoDetailEdit.Visibility = "Collapsed"
        $SpAttachments.Children.Clear()
        $PnlAttachments.Visibility    = "Collapsed"
    } else {
        $TxtMemoTitleDisp.Text        = $item.title
        $SpMemoReadButtons.Visibility = "Visible"
        $SpMemoEditButtons.Visibility = "Collapsed"
        $TxtMemoTitleDisp.Visibility  = "Visible"
        $TxtMemoTitleEdit.Visibility  = "Collapsed"
        $RtbMemoDetail.Visibility     = "Visible"
        $TxtMemoDetailEdit.Visibility = "Collapsed"
        # ピン状態でボタン色を切り替え
        if ($item.pinned -eq 1) {
            $BtnMemoPin.Background = "#D4870A"  # ピン済み：濃いオレンジ
        } else {
            $BtnMemoPin.Background = "#A0A8B0"  # 未ピン：グレー
        }
    }
}

function Enter-MemoEditMode {
    $item = $LstMemos.SelectedItem
    if ($null -eq $item) { return }
    $TxtMemoTitleEdit.Text  = $item.title
    $TxtMemoDetailEdit.Text = $item.detail
    Set-MemoEditMode
    $TxtMemoDetailEdit.Focus() | Out-Null
}

function Save-MemoEdit {
    $cpid = $script:CurrentProjectId; if ([string]::IsNullOrEmpty($cpid)) { return }
    $newTitle = $TxtMemoTitleEdit.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($newTitle)) {
        [System.Windows.MessageBox]::Show("タイトルを入力してください。", "入力エラー", "OK", "Warning") | Out-Null
        $TxtMemoTitleEdit.Focus() | Out-Null
        return
    }
    if ($null -ne $script:MemoNewId) {
        $newId = $script:MemoNewId; $script:MemoNewId = $null
        Save-Memo @{ id = $newId; title = $newTitle; detail = $TxtMemoDetailEdit.Text; updatedAt = Get-Now } $cpid
        Reload-Project $cpid; Refresh-Memos
        for ($i = 0; $i -lt $LstMemos.Items.Count; $i++) {
            if ($LstMemos.Items[$i].id -eq $newId) { $LstMemos.SelectedIndex = $i; break }
        }
    } else {
        # 既存メモ更新
        $item = $LstMemos.SelectedItem; if ($null -eq $item) { return }
        $proj = Get-SelectedProject; if ($null -eq $proj) { return }
        $memo = $proj.memos | Where-Object { $_.id -eq $item.id } | Select-Object -First 1
        if ($null -eq $memo) { return }
        $memo.title     = $newTitle
        $memo.detail    = $TxtMemoDetailEdit.Text
        $memo.updatedAt = Get-Now
        Save-Memo $memo $cpid
        Reload-Project $cpid
        Refresh-Memos
    }
    # SelectionChanged で Set-MemoDetailHeader が呼ばれ Read モードに戻る
}

# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------
$script:MemoNewId = $null   # null=編集中, 非null=新規作成中

function Cancel-MemoEdit {
    $script:MemoNewId = $null
    $item = $LstMemos.SelectedItem
    if ($null -ne $item) { Set-MemoDetailHeader $item; Render-MemoDetail $item.detail $item.id }
    else { Set-MemoDetailHeader $null; $RtbMemoDetail.Text = "" }
}

function Enter-MemoNewMode {
    $cpid = $script:CurrentProjectId; if ([string]::IsNullOrEmpty($cpid)) { return }
    $LstMemos.SelectedIndex = -1   # 先に解除（SelectionChanged が $MemoNewId をリセットする前）
    $script:MemoNewId = New-Id     # 解除後に新IDをセット
    $TxtMemoTitleEdit.Text  = ""
    $TxtMemoDetailEdit.Text = ""
    Set-MemoEditMode
    $TxtMemoTitleEdit.Focus() | Out-Null
}

function Get-SelectedProject {
    $idx = $LstProjects.SelectedIndex
    if ($idx -lt 0) { return $null }
    return $script:Projects[$idx]
}

function Refresh-Projects {
    $sel = $LstProjects.SelectedIndex
    $LstProjects.Items.Clear()
    foreach ($p in $script:Projects) { $LstProjects.Items.Add($p.title) | Out-Null }
    if ($sel -ge 0 -and $sel -lt $LstProjects.Items.Count) { $LstProjects.SelectedIndex = $sel }
}

function Refresh-Memos {
    $proj = Get-SelectedProject
    $LstMemos.Items.Clear()
    $RtbMemoDetail.Text = ""
    if ($null -eq $proj) { return }

    $memos = if ($null -ne $proj.memos) { @($proj.memos) } else { @() }

    # キーワードをスペース区切りで分割（AND検索）
    $rawQuery = $TxtSearch.Text.Trim()
    $keywords = if ($rawQuery -ne "") {
        $rawQuery -split '\s+' | Where-Object { $_ -ne "" }
    } else { @() }

    # ソート: pinned 降順 → updatedAt 降順
    $memos = $memos | Sort-Object @(
        @{ Expression = { [int]$_.pinned }; Descending = $true },
        @{ Expression = { [datetime]$_.updatedAt }; Descending = $true }
    )

    foreach ($m in $memos) {
        # すべてのキーワードがタイトルまたは詳細に含まれる場合のみ表示
        $match = $true
        foreach ($kw in $keywords) {
            $escaped = [regex]::Escape($kw)
            if (($m.title -notmatch $escaped) -and ($m.detail -notmatch $escaped)) {
                $match = $false; break
            }
        }
        if (-not $match) { continue }

        $LstMemos.Items.Add([PSCustomObject]@{
            id      = $m.id
            title   = $m.title
            detail  = $m.detail
            pinned  = $m.pinned
            display = if ([int]$m.pinned -eq 1) { "📌 $($m.title)" } else { $m.title }
        }) | Out-Null
    }
}

function Open-Shortcut ($path) {
    if ([string]::IsNullOrWhiteSpace($path)) { return }
    try { Start-Process $path } catch {
        [System.Windows.MessageBox]::Show("開けませんでした：$path", "エラー", "OK", "Error") | Out-Null
    }
}

function Refresh-Shortcuts {
    $LstShortcuts.Items.Clear()
    $proj = Get-SelectedProject; if ($null -eq $proj) { return }
    $shortcuts = if ($null -ne $proj.shortcuts) { @($proj.shortcuts) } else { @() }
    foreach ($sc in $shortcuts) {
        $LstShortcuts.Items.Add([PSCustomObject]@{
            id   = $sc.id
            name = $sc.name
            path = $sc.path
        }) | Out-Null
    }
}

function Refresh-Tasks {
    $proj = Get-SelectedProject
    $LstTasks.Items.Clear()
    if ($null -eq $proj) { return }
    $tasks = if ($null -ne $proj.tasks) { @($proj.tasks) } else { @() }
    $now = Get-Date
    $priorityOrder = @{ "大" = 0; "中" = 1; "小" = 2 }
    $sorted = $tasks | Sort-Object @(
        @{ Expression = { $priorityOrder[$_.priority] }; Ascending = $true },
        @{ Expression = { [datetime]$_.updatedAt }; Ascending = $false }
    )
    foreach ($t in $sorted) {
        # 期限の表示と色判定
        $dlColor = "None"
        $dlDisp  = ""
        if (-not [string]::IsNullOrEmpty($t.deadline)) {
            try {
                $dlDate = [datetime]::ParseExact($t.deadline, "yyyy/MM/dd", $null)
                $dlDisp = $t.deadline
                $diff   = ($dlDate.Date - $now.Date).TotalDays
                if   ($diff -lt 0)  { $dlColor = "Red"    }
                elseif ($diff -le 3) { $dlColor = "Yellow" }
            } catch {}
        }
        $LstTasks.Items.Add([PSCustomObject]@{
            id            = $t.id
            title         = $t.title
            priority      = $t.priority
            deadline      = $t.deadline
            deadlineDisp  = $dlDisp
            deadlineColor = $dlColor
            detail        = $t.detail
            detailOneLine = ($t.detail -replace "`r`n|`r|`n", " ").Trim()
            updatedAt     = $t.updatedAt
        }) | Out-Null
    }
}

function Show-ProjectContent {
    $proj = Get-SelectedProject
    if ($null -eq $proj) {
        $PlaceholderText.Visibility = "Visible"
        $ContentPanel.Visibility   = "Collapsed"
        return
    }
    $PlaceholderText.Visibility = "Collapsed"
    $ContentPanel.Visibility   = "Visible"
    Refresh-Tasks
    Refresh-Memos
    Refresh-Shortcuts
}

# ---------------------------------------------------------------------------
# イベント系
# ---------------------------------------------------------------------------
$LstProjects.Add_SelectionChanged({
    $TxtSearch.Text = ""   # プロジェクト切替時に検索をリセット
    Show-ProjectContent
    $proj = Get-SelectedProject
    if ($null -ne $proj) {
        $script:CurrentProjectId = [string]$proj.id
        try { Set-Config "lastProjectId" $proj.id } catch {}
    } else {
        $script:CurrentProjectId = ""
    }
})

$BtnAddProject.Add_Click({
    $name = Show-InputDialog "プロジェクト追加" "プロジェクト名を入力してください："
    if ($null -ne $name) {
        $newProj = [ordered]@{
            id        = New-Id; title = $name
            tasks     = [System.Collections.Generic.List[object]]::new()
            memos     = [System.Collections.Generic.List[object]]::new()
            shortcuts = [System.Collections.Generic.List[object]]::new()
        }
        $script:Projects.Add($newProj)
        Save-Project $newProj; Refresh-Projects
        $LstProjects.SelectedIndex = $script:Projects.Count - 1
    }
})

$BtnRenameProject.Add_Click({
    $proj = Get-SelectedProject
    if ($null -eq $proj) { [System.Windows.MessageBox]::Show("プロジェクトを選択してください。"); return }
    $name = Show-InputDialog "プロジェクト編集" "新しいプロジェクト名：" $proj.title
    if ($null -ne $name) {
        $proj.title = $name; Save-Project $proj
        $sel = $LstProjects.SelectedIndex; Refresh-Projects; $LstProjects.SelectedIndex = $sel
    }
})

$BtnDeleteProject.Add_Click({
    $proj = Get-SelectedProject
    if ($null -eq $proj) { [System.Windows.MessageBox]::Show("プロジェクトを選択してください。"); return }
    $confirm = [System.Windows.MessageBox]::Show(
        "「$($proj.title)」を削除しますか？（タスク・メモも全て削除されます）",
        "確認", "YesNo", "Warning")
    if ($confirm -eq "Yes") {
        $idx = $LstProjects.SelectedIndex
        Delete-Project $proj.id
        $script:Projects.RemoveAt($idx)
        Refresh-Projects; Show-ProjectContent
    }
})

$BtnAddTask.Add_Click({
    $proj = Get-SelectedProject; if ($null -eq $proj) { return }
    $res = Show-TaskDialog "タスク追加"
    if ($null -ne $res -and $res.title -ne "") {
        $newTask = [ordered]@{
            id        = New-Id
            title     = $res.title
            priority  = $res.priority
            deadline  = $res.deadline
            detail    = $res.detail
            updatedAt = Get-Now
        }
        Save-Task $newTask $proj.id
        Reload-Project $proj.id; Refresh-Tasks
    }
})

# 行内ボタンのクリックを Button.Click イベントとして ListView でキャッチして編集か削除を判定
$LstTasks.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($s, $e)
        $btn = $e.OriginalSource
        if ($btn -isnot [System.Windows.Controls.Button]) { return }
        $taskId = $btn.Tag
        if ([string]::IsNullOrEmpty($taskId)) { return }

        $proj = Get-SelectedProject; if ($null -eq $proj) { return }
        $task = $proj.tasks | Where-Object { $_.id -eq $taskId } | Select-Object -First 1
        if ($null -eq $task) { return }

        if ($btn.Name -eq "BtnRowDelete") {
            $confirm = [System.Windows.MessageBox]::Show(
                "「$($task.title)」を削除しますか？", "確認", "YesNo", "Warning")
            if ($confirm -eq "Yes") {
                Delete-Task $task.id
                Reload-Project $proj.id; Refresh-Tasks
            }
        }
        $e.Handled = $true
    }
)

# ダブルクリックで編集（タスク行あり）または新規追加（空白行）
$LstTasks.Add_MouseDoubleClick({
    param($s, $e)
    $proj = Get-SelectedProject; if ($null -eq $proj) { return }

    # クリック位置の ListViewItem を取得
    $hit = $LstTasks.InputHitTest($e.GetPosition($LstTasks))
    $item = $null
    $current = $hit
    while ($null -ne $current) {
        if ($current -is [System.Windows.Controls.ListViewItem]) { $item = $current; break }
        $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
    }

    if ($null -ne $item -and $null -ne $item.DataContext) {
        # タスク行でダブルクリック → 編集
        $sel = $item.DataContext
        $dl  = $sel.deadline
        $pr  = $sel.priority
        $det = $sel.detail
        $res = Show-TaskDialog "タスク編集" $sel.title $pr $dl $det
        if ($null -ne $res) {
            $task = $proj.tasks | Where-Object { $_.id -eq $sel.id } | Select-Object -First 1
            if ($task) {
                $task.title     = $res.title
                $task.priority  = $res.priority
                $task.deadline  = $res.deadline
                $task.detail    = $res.detail
                $task.updatedAt = Get-Now
                Save-Task $task $proj.id
                Reload-Project $proj.id; Refresh-Tasks
            }
        }
    } else {
        # 空白行でダブルクリック → 新規追加
        $res = Show-TaskDialog "タスク追加"
        if ($null -ne $res -and $res.title -ne "") {
            $newTask2 = [ordered]@{
                id        = New-Id
                title     = $res.title
                priority  = $res.priority
                deadline  = $res.deadline
                detail    = $res.detail
                updatedAt = Get-Now
            }
            Save-Task $newTask2 $proj.id
            Reload-Project $proj.id; Refresh-Tasks
        }
    }
})

# タスク列幅をウィンドウリサイズに合わせて動的調整（タスク名60%・詳細40%）
$script:ColTitle  = ($LstTasks.View.Columns)[2]  # タスク名
$script:ColDetail = ($LstTasks.View.Columns)[3]  # 詳細
function Resize-TaskColumns {
    $fixed = 40 + 60 + 90 + 18  # 削除 + 優先度 + 期限 + スクロールバー
    $avail = $LstTasks.ActualWidth - $fixed
    if ($avail -lt 60) { return }
    $script:ColTitle.Width  = [Math]::Floor($avail * 0.6)
    $script:ColDetail.Width = [Math]::Floor($avail * 0.4)
}
$LstTasks.Add_SizeChanged({ Resize-TaskColumns })

# ショートカット列幅をウィンドウリサイズに合わせて動的調整（名称35%・パス65%）
$script:ColScName = ($LstShortcuts.View.Columns)[2]  # 名称
$script:ColScPath = ($LstShortcuts.View.Columns)[3]  # パス / URL
function Resize-ShortcutColumns {
    $fixed = 40 + 40 + 18  # 削除 + 編集 + スクロールバー
    $avail = $LstShortcuts.ActualWidth - $fixed
    if ($avail -lt 60) { return }
    $script:ColScName.Width = [Math]::Floor($avail * 0.35)
    $script:ColScPath.Width = [Math]::Floor($avail * 0.65)
}
$LstShortcuts.Add_SizeChanged({ Resize-ShortcutColumns })

$BtnAddShortcut.Add_Click({
    $proj = Get-SelectedProject; if ($null -eq $proj) { return }
    $res = Show-ShortcutDialog "ショートカットを追加"
    if ($null -ne $res) {
        $newSc = [ordered]@{
            id        = New-Id
            name      = $res.name
            path      = $res.path
            updatedAt = Get-Now
        }
        Save-Shortcut $newSc $proj.id
        Reload-Project $proj.id; Refresh-Shortcuts
    }
})

# 行内ボタン（削除・編集）のクリックをキャッチ
$LstShortcuts.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($s, $e)
        $btn = $e.OriginalSource
        if ($btn -isnot [System.Windows.Controls.Button]) { return }
        $scId = $btn.Tag
        if ([string]::IsNullOrEmpty($scId)) { return }

        $proj = Get-SelectedProject; if ($null -eq $proj) { return }
        $sc = $proj.shortcuts | Where-Object { $_.id -eq $scId } | Select-Object -First 1
        if ($null -eq $sc) { return }

        if ($btn.Name -eq "BtnScDelete") {
            Delete-Shortcut $sc.id
            Reload-Project $proj.id; Refresh-Shortcuts
        } elseif ($btn.Name -eq "BtnScEdit") {
            $res = Show-ShortcutDialog "ショートカットを編集" $sc.name $sc.path
            if ($null -ne $res) {
                $sc.name      = $res.name
                $sc.path      = $res.path
                $sc.updatedAt = Get-Now
                Save-Shortcut $sc $proj.id
                Reload-Project $proj.id; Refresh-Shortcuts
            }
        }
        $e.Handled = $true
    }
)

# ダブルクリックでショートカットを開く
$LstShortcuts.Add_MouseDoubleClick({
    param($s, $e)
    $hit = $LstShortcuts.InputHitTest($e.GetPosition($LstShortcuts))
    $current = $hit
    $item = $null
    while ($null -ne $current) {
        if ($current -is [System.Windows.Controls.ListViewItem]) { $item = $current; break }
        $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
    }
    if ($null -ne $item -and $null -ne $item.DataContext) {
        Open-Shortcut $item.DataContext.path
    }
})

$LstMemos.Add_MouseDoubleClick({
    # クリック位置がListBoxItem上かどうかをビジュアルツリーで判定
    $e = $args[1]
    $current = $e.OriginalSource
    $hitItem = $false
    while ($null -ne $current -and $current -isnot [System.Windows.Controls.ListBox]) {
        if ($current -is [System.Windows.Controls.ListBoxItem]) { $hitItem = $true; break }
        $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
    }

    if (-not $hitItem) {
        # 空領域をダブルクリック → 新規メモ入力モード
        Enter-MemoNewMode
        return
    }

    # 既存メモをダブルクリック → 編集モードへ
    $sel = $LstMemos.SelectedItem
    if ($null -ne $sel) { Enter-MemoEditMode }
})

$LstMemos.Add_SelectionChanged({
    $script:MemoNewId = $null
    $item = $LstMemos.SelectedItem
    if ($null -ne $item) {
        Set-MemoDetailHeader $item
        Render-MemoDetail $item.detail $item.id
    } else {
        Set-MemoDetailHeader $null
        $RtbMemoDetail.Text = ""
    }
})

$BtnAddMemo.Add_Click({ Enter-MemoNewMode })


$TxtSearch.Add_TextChanged({ Refresh-Memos })

$BtnClearSearch.Add_Click({
    $TxtSearch.Text = ""
    $TxtSearch.Focus() | Out-Null
})

$BtnMemoPin.Add_Click({
    $item = $LstMemos.SelectedItem; if ($null -eq $item) { return }
    $proj = Get-SelectedProject; if ($null -eq $proj) { return }
    $memo = $proj.memos | Where-Object { $_.id -eq $item.id } | Select-Object -First 1
    if ($null -eq $memo) { return }
    $cpid = $script:CurrentProjectId
    $memo.pinned = if ($memo.pinned -eq 1) { 0 } else { 1 }
    Save-Memo $memo $cpid
    Reload-Project $cpid
    Refresh-Memos
    # 操作したメモを再選択
    for ($i = 0; $i -lt $LstMemos.Items.Count; $i++) {
        if ($LstMemos.Items[$i].id -eq $item.id) { $LstMemos.SelectedIndex = $i; break }
    }
})

$BtnMemoEdit.Add_Click({ Enter-MemoEditMode })

$BtnMemoCancel.Add_Click({ Cancel-MemoEdit })

$BtnMemoSave.Add_Click({ Save-MemoEdit })

$BtnMemoDelete.Add_Click({
    $item = $LstMemos.SelectedItem
    if ($null -eq $item) { return }
    $cpid = $script:CurrentProjectId; if ([string]::IsNullOrEmpty($cpid)) { return }
    $confirm = [System.Windows.MessageBox]::Show(
        "「$($item.title)」を削除しますか？", "確認", "YesNo", "Warning")
    if ($confirm -eq "Yes") {
        $proj = Get-SelectedProject; if ($null -eq $proj) { return }
        $memo = $proj.memos | Where-Object { $_.id -eq $item.id } | Select-Object -First 1
        if ($memo) { Delete-Memo $memo.id; Reload-Project $cpid; Refresh-Memos }
    }
})

# 編集 TextBox: Ctrl+S で保存、Escape でキャンセル
$TxtMemoDetailEdit.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq "S" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        Save-MemoEdit; $e.Handled = $true
    } elseif ($e.Key -eq "Escape") {
        Cancel-MemoEdit; $e.Handled = $true
    }
})
$TxtMemoTitleEdit.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq "Escape") { Cancel-MemoEdit; $e.Handled = $true }
})

# ---------------------------------------------------------------------------
# 描画開始
Refresh-Projects
if ($script:Projects.Count -gt 0) {
    $lastId  = Get-Config "lastProjectId"
    $lastIdx = -1
    if ($lastId) {
        for ($i = 0; $i -lt $script:Projects.Count; $i++) {
            if ($script:Projects[$i].id -eq $lastId) { $lastIdx = $i; break }
        }
    }
    $LstProjects.SelectedIndex = if ($lastIdx -ge 0) { $lastIdx } else { 0 }
}
Show-ProjectContent
$TabMain.SelectedIndex = 0

$window.ShowDialog() | Out-Null
