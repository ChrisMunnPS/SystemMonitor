#Requires -Version 5.1
<#
.SYNOPSIS  WPF System Monitor - CPU, RAM, Disk, Network, Alerts, CSV/HTML export.
.EXAMPLE   .\SystemMonitor.ps1
.EXAMPLE   .\SystemMonitor.ps1 -RefreshSeconds 3 -CpuThreshold 80
#>
[CmdletBinding()]
param(
    [ValidateRange(1,60)][int]$RefreshSeconds = 2,
    [ValidateRange(1,100)][int]$CpuThreshold  = 75,
    [ValidateRange(1,100)][int]$RamThreshold  = 80,
    [ValidateRange(1,100)][int]$DiskThreshold = 90
)

if (-not $env:WPF_STA_CHILD) {
    $env:WPF_STA_CHILD = '1'
    $a = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath,
           '-RefreshSeconds',$RefreshSeconds,'-CpuThreshold',$CpuThreshold,
           '-RamThreshold',$RamThreshold,'-DiskThreshold',$DiskThreshold)
    Start-Process powershell.exe -ArgumentList $a -Wait
    $env:WPF_STA_CHILD = $null; exit
}
$env:WPF_STA_CHILD = $null

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ── Friendly OS name from registry ────────────────────────────────────────────
function Get-FriendlyOS {
    try {
        $r   = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -EA Stop
        $nm  = $r.ProductName      # May say "Windows 10" even on Win 11
        $dv  = $r.DisplayVersion   # e.g. "24H2"
        $bn  = [int]$r.CurrentBuildNumber
        # Build 22000+ is Windows 11; registry ProductName is wrong on some systems
        if ($nm -match 'Windows 10' -and $bn -ge 22000) {
            $nm = $nm -replace 'Windows 10','Windows 11'
        }
        # Also catch cases where edition is missing from ProductName
        if ($nm -notmatch 'Windows 1[01]') {
            $caption = (Get-CimInstance Win32_OperatingSystem -EA SilentlyContinue).Caption
            if ($caption) { $nm = $caption }
        }
        if ($dv) { return "$nm $dv (Build $bn)" }
        return "$nm (Build $bn)"
    } catch { return [Environment]::OSVersion.VersionString }
}

# ── XAML ───────────────────────────────────────────────────────────────────────
[xml]$Xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="System Monitor"
    Height="800" Width="1200" MinHeight="650" MinWidth="950"
    WindowStartupLocation="CenterScreen"
    Background="#0F1117">
  <Grid Margin="8">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- HEADER -->
    <Grid Grid.Row="0" Margin="5,0,5,6">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
        <TextBlock Text="SYSTEM MONITOR" Foreground="#E2E8F0"
                   FontSize="15" FontWeight="Bold" FontFamily="Segoe UI" VerticalAlignment="Center"/>
        <TextBlock x:Name="txtHostname" Foreground="#6C63FF" FontSize="11"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="12,0,0,0"/>
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
        <CheckBox x:Name="chkAutoRefresh" Content="Auto Refresh" IsChecked="True"
                  Foreground="#CBD5E1" FontFamily="Segoe UI" FontSize="11"
                  VerticalAlignment="Center" Margin="3,0,10,0"/>
        <Button x:Name="btnRefresh" Content="Refresh Now" Margin="3,0" Padding="12,7"/>
        <Button x:Name="btnExportCsv"  Content="Export CSV"      Margin="3,0" Padding="12,7"/>
        <Button x:Name="btnExportMd"   Content="Export Markdown" Margin="3,0" Padding="12,7"/>
        <Button x:Name="btnExportHtml" Content="Export HTML"     Margin="3,0" Padding="12,7"/>
        <Button x:Name="btnSettings"   Content="Thresholds"  Margin="3,0" Padding="12,7"/>
      </StackPanel>
    </Grid>

    <!-- REMOTE HOST BAR -->
    <Border Grid.Row="2" Background="#12151F" CornerRadius="8" Padding="8,6"
            Margin="5,0,5,4" BorderBrush="#2D3148" BorderThickness="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="200"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" Text="Remote Host:" Foreground="#8B8FA8" FontSize="11"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="0,0,8,0"/>
        <ComboBox x:Name="cmbRemoteHost" Grid.Column="1" IsEditable="True"
                  FontFamily="Consolas" FontSize="11" Height="26"
                  Text="" ToolTip="Enter hostname or IP. Leave blank for local machine."/>
        <Button x:Name="btnConnect" Grid.Column="2" Content="Connect"
                Margin="6,0,0,0" Padding="10,4" FontSize="11"/>
        <Button x:Name="btnDisconnect" Grid.Column="3" Content="Disconnect"
                Margin="4,0,0,0" Padding="10,4" FontSize="11" IsEnabled="False"/>
        <Ellipse x:Name="ellStatus" Grid.Column="4" Width="10" Height="10"
                 Fill="#4B5563" Margin="8,0,0,0" VerticalAlignment="Center"
                 ToolTip="Connection status"/>
        <TextBlock x:Name="txtRemoteStatus" Grid.Column="5" Foreground="#8B8FA8"
                   FontSize="11" FontFamily="Segoe UI" VerticalAlignment="Center" Margin="6,0,0,0"
                   Text="Local machine"/>
        <TextBlock Grid.Column="6" Text="Filter processes:" Foreground="#8B8FA8" FontSize="11"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="0,0,6,0"/>
      </Grid>
    </Border>

    <!-- TOOL LAUNCH STRIP -->
    <Border Grid.Row="2" Background="#1A1D27" CornerRadius="8" Padding="10,7"
            Margin="5,0,5,6" BorderBrush="#2D3148" BorderThickness="1">
      <DockPanel>
        <TextBlock Text="TOOLS" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI"
                   FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,12,0"
                   DockPanel.Dock="Left"/>
        <WrapPanel x:Name="pnlTools" Orientation="Horizontal"/>
      </DockPanel>
    </Border>

    <!-- KPI CARDS -->
    <Grid Grid.Row="3">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
              BorderBrush="#2D3148" BorderThickness="1">
        <StackPanel>
          <TextBlock Text="CPU USAGE" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" FontWeight="SemiBold"/>
          <Grid Margin="0,3,0,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock x:Name="txtCpu" Text="0%" Foreground="#E2E8F0" FontSize="30" FontFamily="Segoe UI" FontWeight="Bold"/>
            <TextBlock x:Name="txtCpuAlert" Grid.Column="1" Text="[!]" FontSize="14" Foreground="#F97316" VerticalAlignment="Top" Visibility="Collapsed"/>
          </Grid>
          <ProgressBar x:Name="barCpu" Maximum="100" Height="6" Margin="0,5,0,0" BorderThickness="0" Background="#2D3148" Foreground="#6C63FF"/>
          <TextBlock x:Name="txtCpuDetail" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>

      <Border Grid.Column="1" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
              BorderBrush="#2D3148" BorderThickness="1">
        <StackPanel>
          <TextBlock Text="MEMORY" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" FontWeight="SemiBold"/>
          <Grid Margin="0,3,0,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock x:Name="txtRam" Text="0%" Foreground="#E2E8F0" FontSize="30" FontFamily="Segoe UI" FontWeight="Bold"/>
            <TextBlock x:Name="txtRamAlert" Grid.Column="1" Text="[!]" FontSize="14" Foreground="#F97316" VerticalAlignment="Top" Visibility="Collapsed"/>
          </Grid>
          <ProgressBar x:Name="barRam" Maximum="100" Height="6" Margin="0,5,0,0" BorderThickness="0" Background="#2D3148" Foreground="#6C63FF"/>
          <TextBlock x:Name="txtRamDetail" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>

      <Border Grid.Column="2" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
              BorderBrush="#2D3148" BorderThickness="1">
        <StackPanel>
          <TextBlock Text="PRIMARY DISK" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" FontWeight="SemiBold"/>
          <Grid Margin="0,3,0,0">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock x:Name="txtDisk" Text="0%" Foreground="#E2E8F0" FontSize="30" FontFamily="Segoe UI" FontWeight="Bold"/>
            <TextBlock x:Name="txtDiskAlert" Grid.Column="1" Text="[!]" FontSize="14" Foreground="#F97316" VerticalAlignment="Top" Visibility="Collapsed"/>
          </Grid>
          <ProgressBar x:Name="barDisk" Maximum="100" Height="6" Margin="0,5,0,0" BorderThickness="0" Background="#2D3148" Foreground="#6C63FF"/>
          <TextBlock x:Name="txtDiskDetail" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" Margin="0,5,0,0"/>
        </StackPanel>
      </Border>

      <Border Grid.Column="3" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
              BorderBrush="#2D3148" BorderThickness="1">
        <StackPanel>
          <TextBlock Text="NETWORK" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" FontWeight="SemiBold"/>
          <TextBlock x:Name="txtNet" Text="OUT: 0  IN: 0" Foreground="#E2E8F0" FontSize="18" FontFamily="Segoe UI" FontWeight="Bold" Margin="0,3,0,0"/>
          <ProgressBar x:Name="barNet" Maximum="100" Height="6" Margin="0,5,0,0" BorderThickness="0" Background="#2D3148" Foreground="#6C63FF"/>
          <TextBlock x:Name="txtNetDetail" Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" Margin="0,5,0,0"/>
          <TextBlock x:Name="txtPing"      Foreground="#8B8FA8" FontSize="10" FontFamily="Segoe UI" Margin="0,2,0,0"/>
          <TextBlock x:Name="txtAdapter"   Foreground="#6B7080" FontSize="9"  FontFamily="Segoe UI" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
        </StackPanel>
      </Border>
    </Grid>

    <!-- BATTERY / POWER ROW (shown only when battery detected) -->
    <Border x:Name="borderBattery" Grid.Row="3" Background="#1A1D27" CornerRadius="8"
            Padding="10,8" Margin="5,0,5,4" BorderBrush="#2D3148" BorderThickness="1"
            Visibility="Collapsed">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Grid.Column="0" x:Name="txtBattIcon" Text="BAT" Foreground="#6C63FF"
                   FontSize="13" FontWeight="Bold" FontFamily="Segoe UI" VerticalAlignment="Center" Margin="0,0,10,0"/>
        <ProgressBar x:Name="barBatt" Grid.Column="1" Maximum="100" Height="8"
                     Background="#2D3148" Foreground="#6C63FF" BorderThickness="0" VerticalAlignment="Center"/>
        <TextBlock x:Name="txtBattPct"    Grid.Column="2" Foreground="#E2E8F0" FontSize="13" FontWeight="Bold"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="10,0,0,0"/>
        <TextBlock x:Name="txtBattStatus" Grid.Column="3" Foreground="#8B8FA8" FontSize="11"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="10,0,0,0"/>
        <TextBlock x:Name="txtBattTime"   Grid.Column="4" Foreground="#8B8FA8" FontSize="11"
                   FontFamily="Segoe UI" VerticalAlignment="Center" Margin="10,0,0,0"/>
      </Grid>
    </Border>

    <!-- DETAILS -->
    <Grid Grid.Row="5">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="1.4*"/>
      </Grid.ColumnDefinitions>

      <!-- LEFT: Drive pie charts + Alert log -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="*"/>
          <RowDefinition Height="160"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
                BorderBrush="#2D3148" BorderThickness="1">
          <DockPanel>
            <TextBlock DockPanel.Dock="Top" Text="DRIVE USAGE" Foreground="#CBD5E1" FontSize="12"
                       FontFamily="Segoe UI" FontWeight="SemiBold" Margin="0,0,0,8"/>
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <WrapPanel x:Name="wrapDrives" Orientation="Horizontal"/>
            </ScrollViewer>
          </DockPanel>
        </Border>

        <Border Grid.Row="2" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
                BorderBrush="#2D3148" BorderThickness="1">
          <DockPanel>
            <Grid DockPanel.Dock="Top" Margin="0,0,0,6">
              <TextBlock Text="ALERT LOG" Foreground="#CBD5E1" FontSize="12" FontFamily="Segoe UI" FontWeight="SemiBold"/>
              <Button x:Name="btnClearAlerts" Content="Clear" HorizontalAlignment="Right" Padding="8,3" FontSize="10"/>
            </Grid>
            <ListBox x:Name="lstAlerts" Background="Transparent" BorderThickness="0"
                     Foreground="#F97316" FontFamily="Consolas" FontSize="11"
                     ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
          </DockPanel>
        </Border>
      </Grid>

      <!-- RIGHT: Process list -->
      <Border Grid.Column="1" Background="#1A1D27" CornerRadius="8" Padding="14" Margin="5"
              BorderBrush="#2D3148" BorderThickness="1">
        <DockPanel>
          <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="TOP PROCESSES" Foreground="#CBD5E1" FontSize="12" FontFamily="Segoe UI" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
              <Button x:Name="btnEndTask" Content="End Task" Padding="10,4" FontSize="10" Margin="0,0,8,0"/>
              <TextBlock Text="Sort:" Foreground="#8B8FA8" FontSize="11"
                         FontFamily="Segoe UI" VerticalAlignment="Center" Margin="0,0,5,0"/>
              <ComboBox x:Name="cmbSort" FontSize="11" Width="75">
                <ComboBoxItem Content="CPU"    IsSelected="True"/>
                <ComboBoxItem Content="Memory"/>
                <ComboBoxItem Content="Name"/>
              </ComboBox>
            </StackPanel>
          </Grid>
          <Grid DockPanel.Dock="Top" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="txtProcFilter" Background="#12151F" Foreground="#CBD5E1"
                     BorderBrush="#3D4268" BorderThickness="1" Padding="6,4"
                     FontFamily="Consolas" FontSize="11"
                     Tag="Filter processes... (name or PID)"/>
            <Button x:Name="btnProcFilterClear" Grid.Column="1" Content="X"
                    Margin="4,0,0,0" Padding="8,4" FontSize="10"/>
          </Grid>
          <DataGrid x:Name="gridProcs" AutoGenerateColumns="False" IsReadOnly="True"
                    CanUserSortColumns="True" HeadersVisibility="Column" SelectionMode="Single"
                    Background="Transparent" Foreground="#CBD5E1" BorderThickness="0"
                    GridLinesVisibility="Horizontal" FontFamily="Consolas" FontSize="11"
                    RowBackground="Transparent" AlternatingRowBackground="#1E2130">
            <DataGrid.Columns>
              <DataGridTextColumn Header="PID"     Binding="{Binding PID}"     Width="55"/>
              <DataGridTextColumn Header="Process" Binding="{Binding Name}"    Width="*"/>
              <DataGridTextColumn Header="CPU%"    Binding="{Binding CPU}"     Width="60"/>
              <DataGridTextColumn Header="RAM MB"  Binding="{Binding RAM}"     Width="68"/>
              <DataGridTextColumn Header="Threads" Binding="{Binding Threads}" Width="62"/>
              <DataGridTextColumn Header="Status"  Binding="{Binding Status}"  Width="58"/>
            </DataGrid.Columns>
          </DataGrid>
        </DockPanel>
      </Border>
    </Grid>

    <!-- STATUS BAR -->
    <Border Grid.Row="6" Background="#1A1D27" CornerRadius="5" Margin="5,3,5,0"
            Padding="10,5" BorderBrush="#2D3148" BorderThickness="1">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="txtStatus" Foreground="#8B8FA8" FontSize="11" FontFamily="Segoe UI" VerticalAlignment="Center"/>
        <TextBlock x:Name="txtUptime" Grid.Column="1" Foreground="#6C63FF" FontSize="11" FontFamily="Segoe UI" VerticalAlignment="Center" Margin="18,0"/>
        <TextBlock x:Name="txtTime"   Grid.Column="2" Foreground="#8B8FA8" FontSize="11" FontFamily="Consolas" VerticalAlignment="Center"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

# ── Load window ────────────────────────────────────────────────────────────────
$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
$Window = [System.Windows.Markup.XamlReader]::Load($Reader)

$C = @{}
foreach ($n in @(
    'txtHostname','txtCpu','txtCpuAlert','barCpu','txtCpuDetail',
    'txtRam','txtRamAlert','barRam','txtRamDetail',
    'txtDisk','txtDiskAlert','barDisk','txtDiskDetail',
    'txtNet','barNet','txtNetDetail',
    'wrapDrives','gridProcs','lstAlerts','cmbSort','pnlTools',
    'cmbRemoteHost','btnConnect','btnDisconnect','ellStatus','txtRemoteStatus',
    'txtPing','txtAdapter','borderBattery','barBatt','txtBattIcon','txtBattPct','txtBattStatus','txtBattTime',
    'txtProcFilter','btnProcFilterClear',
    'txtStatus','txtUptime','txtTime',
    'chkAutoRefresh','btnRefresh','btnExportCsv','btnExportMd','btnExportHtml','btnSettings','btnClearAlerts','btnEndTask'
)) { $C[$n] = $Window.FindName($n) }

# ── Shared brush converter ─────────────────────────────────────────────────────
$conv = [System.Windows.Media.BrushConverter]::new()

# ── Style helper ──────────────────────────────────────────────────────────────
function Set-ButtonStyle {
    param($b, [string]$Bg = '#2D3148', [string]$Fg = '#CBD5E1')
    $b.Background      = $conv.ConvertFromString($Bg)
    $b.Foreground      = $conv.ConvertFromString($Fg)
    $b.BorderBrush     = $conv.ConvertFromString('#3D4268')
    $b.BorderThickness = [System.Windows.Thickness]::new(1)
    $b.FontFamily      = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $b.FontSize        = 11
    $b.Cursor          = [System.Windows.Input.Cursors]::Hand
}

foreach ($n in @('btnRefresh','btnExportCsv','btnExportMd','btnExportHtml','btnSettings','btnClearAlerts')) {
    Set-ButtonStyle $C[$n]
}
Set-ButtonStyle $C['btnEndTask'] '#3D1A1A' '#EF4444'

# Style checkbox
$C['chkAutoRefresh'].Foreground = $conv.ConvertFromString('#CBD5E1')

# ── Placeholder text for filter box ─────────────────────────────────────────
$C['txtProcFilter'].Add_GotFocus({
    if ($C['txtProcFilter'].Text -eq $C['txtProcFilter'].Tag) {
        $C['txtProcFilter'].Text = ''
        $C['txtProcFilter'].Foreground = $conv.ConvertFromString('#CBD5E1')
    }
})
$C['txtProcFilter'].Add_LostFocus({
    if ($C['txtProcFilter'].Text -eq '') {
        $C['txtProcFilter'].Text = $C['txtProcFilter'].Tag
        $C['txtProcFilter'].Foreground = $conv.ConvertFromString('#4B5563')
    }
})
$C['txtProcFilter'].Text = $C['txtProcFilter'].Tag
$C['txtProcFilter'].Foreground = $conv.ConvertFromString('#4B5563')

# ── Remote host state ─────────────────────────────────────────────────────────
$script:RemoteHost   = ''     # blank = local
$script:CimSession   = $null  # active CIM session when connected remotely

# ── DataGrid column header style ───────────────────────────────────────────────
$hdrStyle = [System.Windows.Style]::new([System.Windows.Controls.Primitives.DataGridColumnHeader])
$T = [System.Windows.Controls.Primitives.DataGridColumnHeader]
foreach ($item in @(
    @{ DP=$T::BackgroundProperty;      Val=$conv.ConvertFromString('#0F1117') },
    @{ DP=$T::ForegroundProperty;      Val=$conv.ConvertFromString('#8B8FA8') },
    @{ DP=$T::FontWeightProperty;      Val=[System.Windows.FontWeights]::SemiBold },
    @{ DP=$T::FontSizeProperty;        Val=[double]11 },
    @{ DP=$T::PaddingProperty;         Val=[System.Windows.Thickness]::new(8,5,8,5) },
    @{ DP=$T::BorderThicknessProperty; Val=[System.Windows.Thickness]::new(0) }
)) { $hdrStyle.Setters.Add([System.Windows.Setter]::new($item.DP, $item.Val)) }
$C['gridProcs'].ColumnHeaderStyle = $hdrStyle

# ── Tool launch strip ──────────────────────────────────────────────────────────
$tools = @(
    @{ Label='Disk Cleanup';    Tip='Free up disk space (cleanmgr)';           Cmd={ Start-Process 'cleanmgr' } },
    @{ Label='Reliability';     Tip='Windows Reliability Monitor';              Cmd={ Start-Process 'perfmon' -ArgumentList '/rel' } },
    @{ Label='WiFi Report';     Tip='Generate & open WiFi diagnostics report';  Cmd={
        $C['txtStatus'].Text = 'Generating WiFi report...'
        Start-Process 'netsh' -ArgumentList 'wlan show wlanreport' -WindowStyle Hidden -Wait
        $rpt = "$env:ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html"
        if (Test-Path $rpt) { Start-Process $rpt }
        else { $C['txtStatus'].Text = '[!] WiFi report not found - adapter may not support this' }
    }},
    @{ Label='Services';        Tip='Windows Services manager (services.msc)';  Cmd={ Start-Process 'services.msc' } },
    @{ Label='Resource Mon';    Tip='Detailed resource usage (resmon)';          Cmd={ Start-Process 'resmon' } },
    @{ Label='Event Viewer';    Tip='System & application event logs';           Cmd={ Start-Process 'eventvwr' } },
    @{ Label='Device Manager';  Tip='Manage hardware & drivers';                 Cmd={ Start-Process 'devmgmt.msc' } },
    @{ Label='Disk Management'; Tip='Partition & format drives (diskmgmt.msc)';  Cmd={ Start-Process 'diskmgmt.msc' } },
    @{ Label='System Info';     Tip='Full hardware & software report (msinfo32)';Cmd={ Start-Process 'msinfo32' } },
    @{ Label='Perf Monitor';    Tip='Performance counters & logs (perfmon)';     Cmd={ Start-Process 'perfmon' } },
    @{ Label='Task Manager';    Tip='Windows Task Manager';                      Cmd={ Start-Process 'taskmgr' } },
    @{ Label='System Config';   Tip='Startup & boot options (msconfig)';         Cmd={ Start-Process 'msconfig' } }
)
foreach ($t in $tools) {
    $tb = [System.Windows.Controls.Button]::new()
    $tb.Content         = $t.Label
    $tb.ToolTip         = $t.Tip
    $tb.Margin          = [System.Windows.Thickness]::new(0,0,5,0)
    $tb.Padding         = [System.Windows.Thickness]::new(10,4,10,4)
    $tb.Background      = $conv.ConvertFromString('#2D3148')
    $tb.Foreground      = $conv.ConvertFromString('#CBD5E1')
    $tb.BorderBrush     = $conv.ConvertFromString('#3D4268')
    $tb.BorderThickness = [System.Windows.Thickness]::new(1)
    $tb.FontFamily      = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $tb.FontSize        = 11
    $tb.Cursor          = [System.Windows.Input.Cursors]::Hand
    $cmd = $t.Cmd
    $handler = { try { & $cmd } catch { $C['txtStatus'].Text = "[!] $($_.Exception.Message)" } }.GetNewClosure()
    $tb.Add_Click([System.Windows.RoutedEventHandler]$handler)
    $C['pnlTools'].Children.Add($tb) | Out-Null
}

# ── Right-click context menu on process grid ───────────────────────────────────
# Fix 1: right-click must select the row under the cursor before the menu opens.
# WPF DataGrid does not auto-select on right-click - handle PreviewMouseRightButtonDown.
$C['gridProcs'].Add_PreviewMouseRightButtonDown({
    param($src, $e)
    $hit = $src.InputHitTest($e.GetPosition($src))
    if ($null -ne $hit) {
        $row = [System.Windows.Media.VisualTreeHelper]::GetParent($hit)
        while ($null -ne $row -and $row -isnot [System.Windows.Controls.DataGridRow]) {
            $row = [System.Windows.Media.VisualTreeHelper]::GetParent($row)
        }
        if ($null -ne $row) { $row.IsSelected = $true }
    }
})

$ctx = [System.Windows.Controls.ContextMenu]::new()

# Fix 2: capture $C and $conv into each handler via GetNewClosure()
$miEnd = [System.Windows.Controls.MenuItem]::new()
$miEnd.Header     = 'End Task'
$miEnd.Foreground = $conv.ConvertFromString('#EF4444')
$miEnd.FontWeight = 'SemiBold'
$miEnd.Add_Click(({
    $row = $C['gridProcs'].SelectedItem
    if ($null -eq $row) { $C['txtStatus'].Text = '[!] No process selected'; return }
    try {
        $p = Get-Process -Id $row.PID -ErrorAction Stop
        $p.Kill()
        $C['txtStatus'].Text = "OK  Ended: $($row.Name) (PID $($row.PID))"
    } catch { $C['txtStatus'].Text = "[!] Could not end $($row.Name): $($_.Exception.Message)" }
}).GetNewClosure())

$miInfo = [System.Windows.Controls.MenuItem]::new()
$miInfo.Header = 'Open File Location'
$miInfo.Add_Click(({
    $row = $C['gridProcs'].SelectedItem
    if ($null -eq $row) { $C['txtStatus'].Text = '[!] No process selected'; return }
    try {
        $p    = Get-Process -Id $row.PID -ErrorAction Stop
        $path = $p.MainModule.FileName
        if ([string]::IsNullOrEmpty($path)) { throw 'No module path available' }
        Start-Process explorer.exe -ArgumentList "/select,`"$path`""
        $C['txtStatus'].Text = "OK  Opening location for $($row.Name)"
    } catch { $C['txtStatus'].Text = "[!] $($row.Name): $($_.Exception.Message)" }
}).GetNewClosure())

$ctx.Items.Add($miEnd)  | Out-Null
$ctx.Items.Add([System.Windows.Controls.Separator]::new()) | Out-Null
$ctx.Items.Add($miInfo) | Out-Null
$C['gridProcs'].ContextMenu = $ctx

# ── Donut pie chart builder ────────────────────────────────────────────────────
function New-DriveCard {
    param([PSCustomObject]$Drive)
    $sz    = 80.0   # chart diameter
    $cx    = $sz / 2
    $cy    = $sz / 2
    $outer = $sz / 2 - 2
    $inner = $sz * 0.30
    $pct   = $Drive._Pct
    $color = if ($pct -ge 90) {'#EF4444'} elseif ($pct -ge 75) {'#F97316'} else {'#6C63FF'}

    # Canvas with pie
    $canvas = [System.Windows.Controls.Canvas]::new()
    $canvas.Width  = $sz
    $canvas.Height = $sz

    # Background track
    $bg = [System.Windows.Shapes.Ellipse]::new()
    $bg.Width  = $sz - 4; $bg.Height = $sz - 4
    $bg.Fill   = $conv.ConvertFromString('#2D3148')
    [System.Windows.Controls.Canvas]::SetLeft($bg, 2)
    [System.Windows.Controls.Canvas]::SetTop($bg,  2)
    $canvas.Children.Add($bg) | Out-Null

    # Pie slice
    if ($pct -gt 0 -and $pct -lt 100) {
        $a   = $pct / 100 * 360
        $rad = $a * [Math]::PI / 180
        $sx  = $cx
        $sy  = $cy - $outer
        $ex  = $cx + $outer * [Math]::Sin($rad)
        $ey  = $cy - $outer * [Math]::Cos($rad)
        $lg  = $a -gt 180

        $geo = [System.Windows.Media.PathGeometry]::new()
        $fig = [System.Windows.Media.PathFigure]::new()
        $fig.StartPoint = [System.Windows.Point]::new($cx, $cy)
        $fig.Segments.Add([System.Windows.Media.LineSegment]::new(
            [System.Windows.Point]::new($sx, $sy), $true))
        $arc = [System.Windows.Media.ArcSegment]::new()
        $arc.Point          = [System.Windows.Point]::new($ex, $ey)
        $arc.Size           = [System.Windows.Size]::new($outer, $outer)
        $arc.IsLargeArc     = $lg
        $arc.SweepDirection = 'Clockwise'
        $fig.Segments.Add($arc)
        $fig.IsClosed = $true
        $geo.Figures.Add($fig)

        $slice      = [System.Windows.Shapes.Path]::new()
        $slice.Data = $geo
        $slice.Fill = $conv.ConvertFromString($color)
        $canvas.Children.Add($slice) | Out-Null
    } elseif ($pct -ge 100) {
        $full = [System.Windows.Shapes.Ellipse]::new()
        $full.Width  = $sz - 4; $full.Height = $sz - 4
        $full.Fill   = $conv.ConvertFromString('#EF4444')
        [System.Windows.Controls.Canvas]::SetLeft($full, 2)
        [System.Windows.Controls.Canvas]::SetTop($full,  2)
        $canvas.Children.Add($full) | Out-Null
    }

    # Donut hole
    $hole = [System.Windows.Shapes.Ellipse]::new()
    $hole.Width  = $inner * 2; $hole.Height = $inner * 2
    $hole.Fill   = $conv.ConvertFromString('#1A1D27')
    [System.Windows.Controls.Canvas]::SetLeft($hole, $cx - $inner)
    [System.Windows.Controls.Canvas]::SetTop($hole,  $cy - $inner)
    $canvas.Children.Add($hole) | Out-Null

    # Percent label inside donut
    $lbl = [System.Windows.Controls.TextBlock]::new()
    $lbl.Text          = "$([int]$pct)%"
    $lbl.Foreground    = $conv.ConvertFromString('#E2E8F0')
    $lbl.FontSize      = 11
    $lbl.FontWeight    = 'Bold'
    $lbl.FontFamily    = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $lbl.TextAlignment = 'Center'
    $lbl.Width         = $inner * 2
    [System.Windows.Controls.Canvas]::SetLeft($lbl, $cx - $inner)
    [System.Windows.Controls.Canvas]::SetTop($lbl,  $cy - 8)
    $canvas.Children.Add($lbl) | Out-Null

    # Outer card
    $card = [System.Windows.Controls.Border]::new()
    $card.Background      = $conv.ConvertFromString('#12151F')
    $card.CornerRadius    = [System.Windows.CornerRadius]::new(6)
    $card.Padding         = [System.Windows.Thickness]::new(8)
    $card.Margin          = [System.Windows.Thickness]::new(0,0,8,8)
    $card.BorderBrush     = $conv.ConvertFromString('#2D3148')
    $card.BorderThickness = [System.Windows.Thickness]::new(1)

    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.HorizontalAlignment = 'Center'

    # Drive letter
    $drv = [System.Windows.Controls.TextBlock]::new()
    $drv.Text          = $Drive.Drive
    $drv.Foreground    = $conv.ConvertFromString($color)
    $drv.FontSize      = 16; $drv.FontWeight = 'Bold'
    $drv.FontFamily    = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $drv.TextAlignment = 'Center'
    $drv.Margin        = [System.Windows.Thickness]::new(0,0,0,4)
    $sp.Children.Add($drv) | Out-Null

    $sp.Children.Add($canvas) | Out-Null

    # Free / Total
    $info = [System.Windows.Controls.TextBlock]::new()
    $info.Text          = "$($Drive.Free) free"
    $info.Foreground    = $conv.ConvertFromString('#8B8FA8')
    $info.FontSize      = 9
    $info.FontFamily    = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $info.TextAlignment = 'Center'
    $info.Margin        = [System.Windows.Thickness]::new(0,4,0,0)
    $sp.Children.Add($info) | Out-Null

    $tot = [System.Windows.Controls.TextBlock]::new()
    $tot.Text          = "of $($Drive.Total)"
    $tot.Foreground    = $conv.ConvertFromString('#6B7080')
    $tot.FontSize      = 9
    $tot.FontFamily    = [System.Windows.Media.FontFamily]::new('Segoe UI')
    $tot.TextAlignment = 'Center'
    $sp.Children.Add($tot) | Out-Null

    $card.Child = $sp
    return $card
}

# ── Helpers ────────────────────────────────────────────────────────────────────
function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return '{0:N1} GB' -f ($Bytes/1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes/1MB) }
    if ($Bytes -ge 1KB) { return '{0:N1} KB' -f ($Bytes/1KB) }
    return "$Bytes B"
}
function Set-BarColour {
    param([System.Windows.Controls.ProgressBar]$Bar,
          [double]$Value, [int]$WarnAt, [int]$CritAt = 90)
    $Bar.Value = $Value
    $hex = if ($Value -ge $CritAt) {'#EF4444'} elseif ($Value -ge $WarnAt) {'#F97316'} else {'#6C63FF'}
    $Bar.Foreground = $conv.ConvertFromString($hex)
}

# ── State ──────────────────────────────────────────────────────────────────────
$script:AlertLog   = [System.Collections.Generic.List[string]]::new()
$script:History    = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:PrevNetIn  = 0L
$script:PrevNetOut = 0L
$script:StartTime  = [datetime]::Now
$script:PrevProcCpu      = @{}
$script:LastProcTime     = $null
$script:AlertLogPath     = [System.IO.Path]::Combine($env:TEMP, "SystemMonitor_alerts.log")
$script:CachedAdapter    = $null   # refresh every 30s
$script:CachedAdapterAge = 0
$script:CachedBattery    = $null
$script:CachedBattAge    = 0
$script:CachedGateway    = $null
$script:LastDriveKey     = $null
$script:TickCount        = 0       # increment each refresh
$script:AutoExportMins   = 0      # 0 = disabled
$script:LastAutoExport   = [datetime]::Now
$script:CpuCounter = [System.Diagnostics.PerformanceCounter]::new('Processor','% Processor Time','_Total')
$null = $script:CpuCounter.NextValue()

# ── Collect metrics (local or remote via CIM) ────────────────────────────────
function Get-CimParam {
    # Returns hashtable of CIM params - uses session for remote, nothing for local
    if ($script:CimSession) { return @{ CimSession = $script:CimSession } }
    return @{}
}

function Get-Metrics {
    $cimP = Get-CimParam
    $isRemote = $null -ne $script:CimSession

    # CPU: local uses fast PerformanceCounter; remote uses WMI LoadPercentage
    if ($isRemote) {
        $cpuWmi = Get-CimInstance Win32_Processor @cimP -ErrorAction Stop
        $cpu = [Math]::Round(($cpuWmi | Measure-Object LoadPercentage -Average).Average, 1)
    } else {
        $cpu = [Math]::Round($script:CpuCounter.NextValue(), 1)
    }

    $os      = Get-CimInstance Win32_OperatingSystem @cimP -ErrorAction Stop
    $ramTot  = $os.TotalVisibleMemorySize * 1KB
    $ramFree = $os.FreePhysicalMemory    * 1KB
    $ramUsed = $ramTot - $ramFree
    $ramPct  = [Math]::Round(($ramUsed/$ramTot)*100, 1)

    # Disks: use Win32_LogicalDisk for both local and remote (consistent)
    $ldisks = Get-CimInstance Win32_LogicalDisk @cimP -Filter "DriveType=3" -ErrorAction SilentlyContinue
    $disks  = $ldisks | ForEach-Object {
        $tot  = $_.Size
        $free = $_.FreeSpace
        $used = $tot - $free
        $pct  = if ($tot -gt 0) { [Math]::Round(($used/$tot)*100,1) } else { 0 }
        [PSCustomObject]@{
            Drive   = $_.DeviceID
            Label   = if ($_.VolumeName) { $_.VolumeName } else { $_.DeviceID }
            Total   = Format-Bytes $tot
            Free    = Format-Bytes $free
            UsedPct = "$pct%"
            _Pct=$pct; _Free=$free; _Total=$tot
        }
    }
    $primary = $disks | Where-Object { $_.Drive -eq 'C:' } | Select-Object -First 1
    if (-not $primary) { $primary = $disks | Select-Object -First 1 }

    # Network (local only - raw perf data not reliably available over CIM remote)
    $netIn = $netOut = 0L
    if (-not $isRemote) {
        $adapters = Get-CimInstance Win32_PerfRawData_Tcpip_NetworkInterface -EA SilentlyContinue
        if ($adapters) {
            $netIn  = [long]($adapters | Measure-Object BytesReceivedPerSec -Sum).Sum
            $netOut = [long]($adapters | Measure-Object BytesSentPerSec     -Sum).Sum
        }
    }
    $dIn  = [Math]::Max(0L, $netIn  - $script:PrevNetIn)
    $dOut = [Math]::Max(0L, $netOut - $script:PrevNetOut)
    $script:PrevNetIn  = $netIn
    $script:PrevNetOut = $netOut

    # Processes: CIM-based for both so remote works; calculate real CPU% via delta
    $rawProcs = Get-CimInstance Win32_Process @cimP -ErrorAction SilentlyContinue

    # Bulk fetch Responding status in one call (far faster than per-process Get-Process)
    $respondingMap = @{}
    if (-not $isRemote) {
        Get-Process -ErrorAction SilentlyContinue | ForEach-Object {
            $respondingMap[$_.Id] = $_.Responding
        }
    }

    # CPU% delta from last sample
    $now = [datetime]::Now
    $elapsed = if ($script:LastProcTime) { ($now - $script:LastProcTime).TotalSeconds } else { $RefreshSeconds }
    $script:LastProcTime = $now

    $procs = $rawProcs | ForEach-Object {
        $pid_ = $_.ProcessId
        $prevCpu = if ($script:PrevProcCpu.ContainsKey($pid_)) { $script:PrevProcCpu[$pid_] } else { $null }
        $currCpu = $_.KernelModeTime + $_.UserModeTime   # 100-nanosecond units
        $cpuPct  = if ($null -ne $prevCpu -and $elapsed -gt 0) {
            [Math]::Round(($currCpu - $prevCpu) / ($elapsed * 1e7 * [Environment]::ProcessorCount) * 100, 1)
        } else { 0 }
        $cpuPct = [Math]::Max(0, [Math]::Min(100, $cpuPct))
        $script:PrevProcCpu[$pid_] = $currCpu

        $responding = if ($isRemote) { $true } else {
            # Use pre-fetched hashtable instead of per-process Get-Process call
            if ($respondingMap.ContainsKey($pid_)) { $respondingMap[$pid_] } else { $true }
        }
        [PSCustomObject]@{
            PID=$pid_; Name=$_.Name; CPU=$cpuPct
            RAM=[Math]::Round($_.WorkingSetSize/1MB,1)
            Threads=$_.ThreadCount
            Status=if($responding){'OK'}else{'NR'}
            _Responding=$responding
        }
    }

    # Sort and filter top 30
    $procs = switch ($C['cmbSort'].SelectedIndex) {
        1 { $procs | Sort-Object RAM     -Descending | Select-Object -First 30 }
        2 { $procs | Sort-Object Name               | Select-Object -First 30 }
        default { $procs | Sort-Object CPU -Descending | Select-Object -First 30 }
    }

    [PSCustomObject]@{
        Timestamp=[datetime]::Now; CPU=$cpu
        RamPct=$ramPct
        RamUsedGB=[Math]::Round($ramUsed/1GB,2); RamTotalGB=[Math]::Round($ramTot/1GB,2)
        DiskPct=if($primary){$primary._Pct}else{0}
        DiskFreeGB=if($primary){[Math]::Round($primary._Free/1GB,2)}else{0}
        DiskTotalGB=if($primary){[Math]::Round($primary._Total/1GB,2)}else{0}
        NetInKb=[Math]::Round($dIn/1KB,1); NetOutKb=[Math]::Round($dOut/1KB,1)
        Disks=$disks; Processes=$procs
    }
}

# ── Update UI ──────────────────────────────────────────────────────────────────
function Update-UI {
    param([PSCustomObject]$m)

    $C['txtCpu'].Text        = "$($m.CPU)%"
    Set-BarColour $C['barCpu']  $m.CPU     $CpuThreshold
    $C['txtCpuDetail'].Text  = "Cores: $([Environment]::ProcessorCount)"

    $C['txtRam'].Text        = "$($m.RamPct)%"
    Set-BarColour $C['barRam']  $m.RamPct  $RamThreshold
    $C['txtRamDetail'].Text  = "$($m.RamUsedGB) GB / $($m.RamTotalGB) GB"

    $C['txtDisk'].Text       = "$($m.DiskPct)%"
    Set-BarColour $C['barDisk'] $m.DiskPct $DiskThreshold
    $C['txtDiskDetail'].Text = "$($m.DiskFreeGB) GB free of $($m.DiskTotalGB) GB"

    $outStr = Format-Bytes ([long]($m.NetOutKb * 1KB))
    $inStr  = Format-Bytes ([long]($m.NetInKb  * 1KB))
    $C['txtNet'].Text        = "OUT: $outStr/s   IN: $inStr/s"
    $C['barNet'].Value       = [Math]::Min(100.0, ($m.NetInKb + $m.NetOutKb) / 10)
    $C['txtNetDetail'].Text  = "Live network I/O"

    # Rebuild drive pie charts only when data changes (compare key string)
    $driveKey = ($m.Disks | ForEach-Object { "$($_.Drive)$($_.UsedPct)" }) -join '|'
    if ($driveKey -ne $script:LastDriveKey) {
        $script:LastDriveKey = $driveKey
        $C['wrapDrives'].Children.Clear()
        foreach ($d in $m.Disks) {
            $C['wrapDrives'].Children.Add((New-DriveCard $d)) | Out-Null
        }
    }

    $C['gridProcs'].ItemsSource = $m.Processes

    $ts = $m.Timestamp.ToString('HH:mm:ss')
    $newAlert = $false
    foreach ($chk in @(
        @{Val=$m.CPU;    Thr=$CpuThreshold;  Label='CPU';  Ctrl=$C['txtCpuAlert']},
        @{Val=$m.RamPct; Thr=$RamThreshold;  Label='RAM';  Ctrl=$C['txtRamAlert']},
        @{Val=$m.DiskPct;Thr=$DiskThreshold; Label='Disk'; Ctrl=$C['txtDiskAlert']}
    )) {
        if ($chk.Val -ge $chk.Thr) {
            $chk.Ctrl.Visibility = 'Visible'
            $msg = "[$ts] ALERT: $($chk.Label) at $($chk.Val)% (limit $($chk.Thr)%)"
            if ($script:AlertLog.Count -eq 0 -or $script:AlertLog[-1] -ne $msg) {
                $script:AlertLog.Add($msg)
                $C['lstAlerts'].Items.Add($msg)
                $C['lstAlerts'].ScrollIntoView($msg)
                # Persist to rolling log file
                try { Add-Content -Path $script:AlertLogPath -Value $msg -Encoding UTF8 } catch {}
                $newAlert = $true
            }
        } else { $chk.Ctrl.Visibility = 'Collapsed' }
    }

    # Battery - refresh every 15 ticks (~30s at 2s interval)
    $script:TickCount++
    if ($script:TickCount % 15 -eq 1 -or $null -eq $script:CachedBattery) {
        $script:CachedBattery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
    }
    try {
        $batt = $script:CachedBattery
        if ($batt) {
            $C['borderBattery'].Visibility = 'Visible'
            $pct = $batt.EstimatedChargeRemaining
            $C['barBatt'].Value   = $pct
            $C['txtBattPct'].Text = "$pct%"
            $hex = if ($pct -le 15) {'#EF4444'} elseif ($pct -le 30) {'#F97316'} else {'#6C63FF'}
            $C['barBatt'].Foreground = $conv.ConvertFromString($hex)
            $statusStr = switch ($batt.BatteryStatus) {
                1 {'Discharging'} 2 {'AC (full)'} 3 {'Fully Charged'} 4 {'Low'} 5 {'Critical'}
                6 {'Charging'} 7 {'Charging+High'} 8 {'Charging+Low'} 9 {'Charging+Crit'} 11 {'Partial AC'}
                default {"Status $($batt.BatteryStatus)"}
            }
            $C['txtBattStatus'].Text = $statusStr
            $C['txtBattIcon'].Foreground = $conv.ConvertFromString($hex)
            $mins = $batt.EstimatedRunTime
            $C['txtBattTime'].Text = if ($mins -and $mins -lt 65535) {
                "$([int]($mins/60))h $($mins%60)m remaining"
            } else { '' }
        } else { $C['borderBattery'].Visibility = 'Collapsed' }
    } catch { $C['borderBattery'].Visibility = 'Collapsed' }

    # Ping - async so it never blocks the UI thread
    try {
        if ($script:TickCount % 5 -eq 1 -or $null -eq $script:CachedGateway) {
            $script:CachedGateway = if ($script:RemoteHost) { $script:RemoteHost } else {
                (Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction Stop |
                 Sort-Object RouteMetric | Select-Object -First 1).NextHop
            }
        }
        if ($script:CachedGateway) {
            $pingObj   = [System.Net.NetworkInformation.Ping]::new()
            $pingAsync = $pingObj.SendPingAsync($script:CachedGateway, 800)
            $pingAsync.ContinueWith([System.Action[System.Threading.Tasks.Task[System.Net.NetworkInformation.PingReply]]]{
                param($t)
                $r = $t.Result
                $txt = if ($r.Status -eq 'Success') { "Ping $($script:CachedGateway): $($r.RoundtripTime) ms" } else { "Ping $($script:CachedGateway): timeout" }
                $C['txtPing'].Dispatcher.InvokeAsync({ $C['txtPing'].Text = $txt }) | Out-Null
                $pingObj.Dispose()
            }) | Out-Null
        }
    } catch { $C['txtPing'].Text = '' }

    # Adapter info - expensive, refresh every 30 ticks (~60s)
    if ($script:TickCount % 30 -eq 1 -or $null -eq $script:CachedAdapter) {
        try {
            $nic = Get-NetAdapter -ErrorAction Stop |
                   Where-Object { $_.Status -eq 'Up' -and $_.InterfaceDescription -notlike '*Loopback*' } |
                   Sort-Object LinkSpeed -Descending | Select-Object -First 1
            if ($nic) {
                $ip = (Get-NetIPAddress -InterfaceIndex $nic.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                       Select-Object -First 1).IPAddress
                $script:CachedAdapter = "$($nic.Name) | $ip | $([Math]::Round($nic.LinkSpeed/1MB,0)) Mbps"
            }
        } catch { $script:CachedAdapter = '' }
    }
    $C['txtAdapter'].Text = $script:CachedAdapter

    # Scheduled auto-export
    if ($script:AutoExportMins -gt 0) {
        $minsSinceLast = ([datetime]::Now - $script:LastAutoExport).TotalMinutes
        if ($minsSinceLast -ge $script:AutoExportMins) {
            $script:LastAutoExport = [datetime]::Now
            $autoPath = [System.IO.Path]::Combine(
                $env:TEMP, "SystemMonitor_auto_$(Get-Date -Format 'yyyyMMdd_HHmmss').html")
            try {
                Export-ToHtmlFile $autoPath
                $C['txtStatus'].Text = "Auto-exported -> $autoPath"
            } catch {}
        }
    }

    $up = [datetime]::Now - $script:StartTime
    $C['txtUptime'].Text = "Up: $([int]$up.TotalHours)h $($up.Minutes)m $($up.Seconds)s"
    $C['txtTime'].Text   = $m.Timestamp.ToString('ddd dd MMM yyyy  HH:mm:ss')

    # 10-second rolling average for status bar
    $script:History.Add($m)
    if ($script:History.Count -gt 1000) { $script:History.RemoveAt(0) }

    $window10 = $script:History | Where-Object { ($m.Timestamp - $_.Timestamp).TotalSeconds -le 10 }
    if ($window10.Count -gt 0) {
        $avgCpu  = [Math]::Round(($window10 | Measure-Object CPU    -Average).Average, 1)
        $avgRam  = [Math]::Round(($window10 | Measure-Object RamPct -Average).Average, 1)
        $avgDisk = [Math]::Round(($window10 | Measure-Object DiskPct -Average).Average, 1)
        $totNetIn  = Format-Bytes ([long](($window10 | Measure-Object NetInKb  -Sum).Sum * 1KB))
        $totNetOut = Format-Bytes ([long](($window10 | Measure-Object NetOutKb -Sum).Sum * 1KB))
        $peakCpu = [Math]::Round(($window10 | Measure-Object CPU -Maximum).Maximum, 1)
        $statStr = "10s avg: CPU ${avgCpu}% (peak ${peakCpu}%)  RAM ${avgRam}%  Disk ${avgDisk}%  Net IN $totNetIn OUT $totNetOut"
    } else { $statStr = "OK  Refresh every ${RefreshSeconds}s" }

    $C['txtStatus'].Text = if ($newAlert) { "[!] New alert -- see Alert Log  |  $statStr" } else { $statStr }
}

# ── Export CSV ─────────────────────────────────────────────────────────────────
function Export-ToCsv {
    $dlg = [System.Windows.Forms.SaveFileDialog]@{
        Title='Export CSV'; Filter='CSV files (*.csv)|*.csv'
        FileName="SystemMonitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    }
    if ($dlg.ShowDialog() -eq 'OK') {
        $script:History | Select-Object Timestamp,CPU,RamPct,RamUsedGB,RamTotalGB,
            DiskPct,DiskFreeGB,DiskTotalGB,NetInKb,NetOutKb |
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $C['txtStatus'].Text = "DONE  Exported $($script:History.Count) rows -> $($dlg.FileName)"
    }
}

# ── Export Markdown ───────────────────────────────────────────────────────────
function Export-ToMarkdown {
    $dlg = [System.Windows.Forms.SaveFileDialog]@{
        Title='Export Markdown'; Filter='Markdown files (*.md)|*.md'
        FileName="SystemMonitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').md"
    }
    if ($dlg.ShowDialog() -ne 'OK') { return }

    $data = $script:History | Select-Object -Last 200
    if ($data.Count -eq 0) { $C['txtStatus'].Text = '[!] No data to export'; return }

    $avgCpu   = [Math]::Round(($data | Measure-Object CPU     -Average).Average,1)
    $avgRam   = [Math]::Round(($data | Measure-Object RamPct  -Average).Average,1)
    $avgDisk  = [Math]::Round(($data | Measure-Object DiskPct -Average).Average,1)
    $peakCpu  = [Math]::Round(($data | Measure-Object CPU     -Maximum).Maximum,1)
    $peakRam  = [Math]::Round(($data | Measure-Object RamPct  -Maximum).Maximum,1)
    $peakDisk = [Math]::Round(($data | Measure-Object DiskPct -Maximum).Maximum,1)
    $totIn    = Format-Bytes ([long](($data | Measure-Object NetInKb  -Sum).Sum * 1KB))
    $totOut   = Format-Bytes ([long](($data | Measure-Object NetOutKb -Sum).Sum * 1KB))
    $duration = ($data[-1].Timestamp - $data[0].Timestamp).ToString('hh\:mm\:ss')
    $alerts   = if ($script:AlertLog.Count -gt 0) {
        ($script:AlertLog | ForEach-Object { "- $_" }) -join "`n"
    } else { "_No alerts triggered during this session._" }

    $rows = $data | ForEach-Object {
        "| $($_.Timestamp.ToString('HH:mm:ss')) | $($_.CPU)% | $($_.RamPct)% | $($_.RamUsedGB) GB | $($_.DiskPct)% | $($_.NetInKb) KB/s | $($_.NetOutKb) KB/s |"
    }

    $md = @"
# System Monitor Report
**Host:** $([Environment]::MachineName)
**OS:** $(Get-FriendlyOS)
**Exported:** $(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm:ss')
**Session duration:** $duration  |  **Samples:** $($data.Count)
**Thresholds:** CPU >= ${CpuThreshold}%  RAM >= ${RamThreshold}%  Disk >= ${DiskThreshold}%

---

## Summary Statistics

| Metric | Average | Peak |
|--------|---------|------|
| CPU | ${avgCpu}% | ${peakCpu}% |
| RAM | ${avgRam}% | ${peakRam}% |
| Disk (C:) | ${avgDisk}% | ${peakDisk}% |
| Net In (total) | $totIn | - |
| Net Out (total) | $totOut | - |

---

## Alerts

$alerts

---

## Sample Data (last $($data.Count) samples)

| Time | CPU | RAM% | RAM Used | Disk% | Net In | Net Out |
|------|-----|------|----------|-------|--------|---------|
$($rows -join "`n")
"@
    $md | Set-Content -Path $dlg.FileName -Encoding UTF8
    $C['txtStatus'].Text = "DONE  Markdown exported -> $($dlg.FileName)"
    Start-Process $dlg.FileName
}

# ── Export HTML to path (used by auto-export and interactive export) ────────────
function Export-ToHtmlFile {
    param([string]$Path)
    $data = $script:History | Select-Object -Last 300
    if ($data.Count -eq 0) { return }
    $labels   = ($data | ForEach-Object { "`"$($_.Timestamp.ToString('HH:mm:ss'))`"" }) -join ','
    $cpuArr   = ($data | ForEach-Object { $_.CPU })     -join ','
    $ramArr   = ($data | ForEach-Object { $_.RamPct })  -join ','
    $diskArr  = ($data | ForEach-Object { $_.DiskPct }) -join ','
    $netInArr = ($data | ForEach-Object { $_.NetInKb })  -join ','
    $netOutArr= ($data | ForEach-Object { $_.NetOutKb }) -join ','
    $scatterCpuNet = ($data | ForEach-Object { "{x:$($_.CPU),y:$($_.NetInKb)}" }) -join ','
    $scatterCpuRam = ($data | ForEach-Object { "{x:$($_.CPU),y:$($_.RamPct)}" })  -join ','
    $step = [Math]::Max(1, [int]($data.Count / 60))
    $hmData   = $data | Where-Object { ($data.IndexOf($_) % $step) -eq 0 }
    $hmLabels = ($hmData | ForEach-Object { "`"$($_.Timestamp.ToString('HH:mm:ss'))`"" }) -join ','
    $hmCpu    = ($hmData | ForEach-Object { $_.CPU })     -join ','
    $hmRam    = ($hmData | ForEach-Object { $_.RamPct })  -join ','
    $hmDisk   = ($hmData | ForEach-Object { $_.DiskPct }) -join ','
    $avgCpu   = [Math]::Round(($data | Measure-Object CPU     -Average).Average,1)
    $avgRam   = [Math]::Round(($data | Measure-Object RamPct  -Average).Average,1)
    $avgDisk  = [Math]::Round(($data | Measure-Object DiskPct -Average).Average,1)
    $peakCpu  = [Math]::Round(($data | Measure-Object CPU     -Maximum).Maximum,1)
    $peakRam  = [Math]::Round(($data | Measure-Object RamPct  -Maximum).Maximum,1)
    $peakDisk = [Math]::Round(($data | Measure-Object DiskPct -Maximum).Maximum,1)
    $totIn    = Format-Bytes ([long](($data | Measure-Object NetInKb  -Sum).Sum * 1KB))
    $totOut   = Format-Bytes ([long](($data | Measure-Object NetOutKb -Sum).Sum * 1KB))
    $rows = $data | Select-Object -Last 100 | ForEach-Object {
        $cc=if($_.CPU     -ge $CpuThreshold)  {' class="w"'} else {''}
        $rc=if($_.RamPct  -ge $RamThreshold)  {' class="w"'} else {''}
        $dc=if($_.DiskPct -ge $DiskThreshold) {' class="w"'} else {''}
        "<tr><td>$($_.Timestamp.ToString('HH:mm:ss'))</td><td$cc>$($_.CPU)%</td>" +
        "<td$rc>$($_.RamPct)%</td><td>$($_.RamUsedGB)/$($_.RamTotalGB) GB</td>" +
        "<td$dc>$($_.DiskPct)%</td><td>$($_.NetInKb) KB/s</td><td>$($_.NetOutKb) KB/s</td></tr>"
    }
    $alerts = if ($script:AlertLog.Count -gt 0) {
        ($script:AlertLog | ForEach-Object { "<li>$_</li>" }) -join ''
    } else { '<li><em>No alerts triggered</em></li>' }
@"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>System Monitor Report - $([Environment]::MachineName)</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#0F1117;color:#CBD5E1;padding:24px}
h1{color:#6C63FF;margin-bottom:4px;font-size:22px}
h2{color:#CBD5E1;font-size:15px;margin:28px 0 12px;border-bottom:1px solid #2D3148;padding-bottom:6px}
.meta{color:#8B8FA8;font-size:12px;margin-bottom:20px}
.stat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin-bottom:8px}
.stat-card{background:#1A1D27;border:1px solid #2D3148;border-radius:8px;padding:14px}
.stat-card .val{font-size:26px;font-weight:700;color:#E2E8F0}
.stat-card .peak{font-size:11px;color:#8B8FA8;margin-top:2px}
.stat-card .lbl{font-size:10px;color:#8B8FA8;font-weight:600;text-transform:uppercase;margin-bottom:6px}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}
.chart-box,.chart-full{background:#1A1D27;border:1px solid #2D3148;border-radius:8px;padding:16px;margin-bottom:16px}
.chart-box h3,.chart-full h3{font-size:12px;color:#8B8FA8;font-weight:600;text-transform:uppercase;margin-bottom:12px}
table{border-collapse:collapse;width:100%;font-size:12px}
th{background:#1A1D27;color:#8B8FA8;padding:8px 12px;text-align:left;font-size:11px}
td{padding:6px 12px;border-bottom:1px solid #2D3148}
tr:hover td{background:#1E2130}.w{color:#F97316;font-weight:bold}
canvas{max-height:280px}
ul{padding-left:20px;color:#F97316;font-size:12px;font-family:Consolas,monospace}
li{margin:3px 0}
</style>
</head>
<body>
<h1>System Monitor &mdash; $([Environment]::MachineName)</h1>
<div class="meta">
  <strong>OS:</strong> $(Get-FriendlyOS) &nbsp;|&nbsp;
  <strong>Exported:</strong> $(Get-Date -Format 'dddd, dd MMMM yyyy HH:mm:ss') &nbsp;|&nbsp;
  <strong>Samples:</strong> $($data.Count) &nbsp;|&nbsp;
  <strong>Thresholds:</strong> CPU&ge;${CpuThreshold}% &nbsp; RAM&ge;${RamThreshold}% &nbsp; Disk&ge;${DiskThreshold}%
</div>
<h2>Summary</h2>
<div class="stat-grid">
  <div class="stat-card"><div class="lbl">CPU Avg</div><div class="val">${avgCpu}%</div><div class="peak">Peak ${peakCpu}%</div></div>
  <div class="stat-card"><div class="lbl">RAM Avg</div><div class="val">${avgRam}%</div><div class="peak">Peak ${peakRam}%</div></div>
  <div class="stat-card"><div class="lbl">Disk Avg</div><div class="val">${avgDisk}%</div><div class="peak">Peak ${peakDisk}%</div></div>
  <div class="stat-card"><div class="lbl">Net In Total</div><div class="val" style="font-size:18px">$totIn</div><div class="peak">cumulative</div></div>
  <div class="stat-card"><div class="lbl">Net Out Total</div><div class="val" style="font-size:18px">$totOut</div><div class="peak">cumulative</div></div>
</div>
<h2>Alerts</h2><ul>$alerts</ul>
<h2>Charts</h2>
<div class="chart-grid">
  <div class="chart-box"><h3>CPU / RAM / Disk Over Time</h3><canvas id="lineChart"></canvas></div>
  <div class="chart-box"><h3>Network I/O Over Time (stacked area)</h3><canvas id="areaChart"></canvas></div>
</div>
<div class="chart-grid">
  <div class="chart-box"><h3>Resource Balance - last 40 samples</h3><canvas id="barChart"></canvas></div>
  <div class="chart-box"><h3>Net In vs Net Out - last 40 samples</h3><canvas id="netBar"></canvas></div>
</div>
<div class="chart-grid">
  <div class="chart-box"><h3>CPU % vs Net In KB/s (scatter)</h3><canvas id="scatter1"></canvas></div>
  <div class="chart-box"><h3>CPU % vs RAM % (scatter)</h3><canvas id="scatter2"></canvas></div>
</div>
<div class="chart-full"><h3>Heatmap - Resource intensity over time</h3><canvas id="heatmap" style="max-height:160px"></canvas></div>
<h2>Sample Data (last 100)</h2>
<table><thead><tr><th>Time</th><th>CPU</th><th>RAM%</th><th>RAM Used</th><th>Disk%</th><th>Net In</th><th>Net Out</th></tr></thead>
<tbody>$($rows -join '')</tbody></table>
<script>
const LABELS=[$labels],CPU=[$cpuArr],RAM=[$ramArr],DISK=[$diskArr],NETIN=[$netInArr],NETOUT=[$netOutArr];
const opt=()=>({responsive:true,animation:false,plugins:{legend:{labels:{color:'#8B8FA8',font:{size:11}}},tooltip:{mode:'index',intersect:false}},scales:{x:{ticks:{color:'#6B7080',maxTicksLimit:8,font:{size:10}},grid:{color:'#2D3148'}},y:{ticks:{color:'#6B7080',font:{size:10}},grid:{color:'#2D3148'}}}});
new Chart(document.getElementById('lineChart'),{type:'line',data:{labels:LABELS,datasets:[{label:'CPU %',data:CPU,borderColor:'#6C63FF',backgroundColor:'transparent',pointRadius:0,borderWidth:2,tension:0.3},{label:'RAM %',data:RAM,borderColor:'#06B6D4',backgroundColor:'transparent',pointRadius:0,borderWidth:2,tension:0.3},{label:'Disk %',data:DISK,borderColor:'#10B981',backgroundColor:'transparent',pointRadius:0,borderWidth:2,tension:0.3}]},options:{...opt(),scales:{x:{ticks:{color:'#6B7080',maxTicksLimit:8,font:{size:10}},grid:{color:'#2D3148'}},y:{min:0,max:100,ticks:{color:'#6B7080',font:{size:10}},grid:{color:'#2D3148'}}}}});
new Chart(document.getElementById('areaChart'),{type:'line',data:{labels:LABELS,datasets:[{label:'Net Out',data:NETOUT,borderColor:'#F97316',backgroundColor:'rgba(249,115,22,0.3)',fill:true,pointRadius:0,tension:0.3},{label:'Net In',data:NETIN,borderColor:'#6C63FF',backgroundColor:'rgba(108,99,255,0.3)',fill:true,pointRadius:0,tension:0.3}]},options:opt()});
const N=40,SL=Math.max(0,LABELS.length-N),bl=LABELS.slice(SL),bc=CPU.slice(SL),br=RAM.slice(SL),bd=DISK.slice(SL),bi=NETIN.slice(SL),bo=NETOUT.slice(SL);
new Chart(document.getElementById('barChart'),{type:'bar',data:{labels:bl,datasets:[{label:'CPU %',data:bc,backgroundColor:'rgba(108,99,255,0.8)'},{label:'RAM %',data:br,backgroundColor:'rgba(6,182,212,0.8)'},{label:'Disk %',data:bd,backgroundColor:'rgba(16,185,129,0.8)'}]},options:{...opt(),scales:{x:{ticks:{color:'#6B7080',maxTicksLimit:10,font:{size:9}},grid:{color:'#2D3148'}},y:{min:0,max:100,ticks:{color:'#6B7080',font:{size:10}},grid:{color:'#2D3148'}}}}});
new Chart(document.getElementById('netBar'),{type:'bar',data:{labels:bl,datasets:[{label:'Net In',data:bi,backgroundColor:'rgba(108,99,255,0.8)'},{label:'Net Out',data:bo,backgroundColor:'rgba(249,115,22,0.8)'}]},options:opt()});
new Chart(document.getElementById('scatter1'),{type:'scatter',data:{datasets:[{label:'CPU vs Net In',data:[$scatterCpuNet],backgroundColor:'rgba(108,99,255,0.6)',pointRadius:3}]},options:{...opt(),scales:{x:{title:{display:true,text:'CPU %',color:'#8B8FA8'},ticks:{color:'#6B7080'},grid:{color:'#2D3148'}},y:{title:{display:true,text:'Net In KB/s',color:'#8B8FA8'},ticks:{color:'#6B7080'},grid:{color:'#2D3148'}}}}});
new Chart(document.getElementById('scatter2'),{type:'scatter',data:{datasets:[{label:'CPU vs RAM',data:[$scatterCpuRam],backgroundColor:'rgba(6,182,212,0.6)',pointRadius:3}]},options:{...opt(),scales:{x:{title:{display:true,text:'CPU %',color:'#8B8FA8'},ticks:{color:'#6B7080'},grid:{color:'#2D3148'}},y:{title:{display:true,text:'RAM %',color:'#8B8FA8'},min:0,max:100,ticks:{color:'#6B7080'},grid:{color:'#2D3148'}}}}});
const HML=[$hmLabels],HMC=[$hmCpu],HMR=[$hmRam],HMD=[$hmDisk];
new Chart(document.getElementById('heatmap'),{type:'bar',data:{labels:HML,datasets:[{label:'CPU',data:HMC,backgroundColor:HMC.map(v=>'rgba(108,99,255,'+Math.min(1,(v/100)*1.2)+')'),borderWidth:0},{label:'RAM',data:HMR,backgroundColor:HMR.map(v=>'rgba(6,182,212,'+Math.min(1,(v/100)*1.2)+')'),borderWidth:0},{label:'Disk',data:HMD,backgroundColor:HMD.map(v=>'rgba(16,185,129,'+Math.min(1,(v/100)*1.2)+')'),borderWidth:0}]},options:{responsive:true,animation:false,plugins:{legend:{labels:{color:'#8B8FA8',font:{size:11}}}},scales:{x:{stacked:false,ticks:{color:'#6B7080',maxTicksLimit:12,font:{size:9}},grid:{color:'#2D3148'}},y:{display:false}}}});
</script>
</body>
</html>
"@ | Set-Content -Path $Path -Encoding UTF8
}

# ── Export HTML ────────────────────────────────────────────────────────────────
function Export-ToHtml {
    $dlg = [System.Windows.Forms.SaveFileDialog]@{
        Title='Export HTML Report'; Filter='HTML files (*.html)|*.html'
        FileName="SystemMonitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }
    if ($dlg.ShowDialog() -ne 'OK') { return }
    if ($script:History.Count -eq 0) { $C['txtStatus'].Text = '[!] No data to export'; return }
    Export-ToHtmlFile $dlg.FileName
    $C['txtStatus'].Text = "DONE  HTML exported -> $($dlg.FileName)"
    Start-Process $dlg.FileName
}

# ── Thresholds dialog ──────────────────────────────────────────────────────────
function Show-SettingsDialog {
    $sd = [System.Windows.Window]::new()
    $sd.Title='Alert Thresholds'; $sd.Width=320; $sd.Height=300
    $sd.WindowStartupLocation='CenterOwner'; $sd.Owner=$Window; $sd.ResizeMode='NoResize'
    $sd.Background=$conv.ConvertFromString('#1A1D27')
    $sp=[System.Windows.Controls.StackPanel]::new()
    $sp.Margin=[System.Windows.Thickness]::new(20)
    $hdr=[System.Windows.Controls.TextBlock]::new()
    $hdr.Text='Set alert thresholds (%)'
    $hdr.Foreground=$conv.ConvertFromString('#CBD5E1')
    $hdr.FontSize=13; $hdr.FontWeight='SemiBold'
    $hdr.Margin=[System.Windows.Thickness]::new(0,0,0,12)
    $sp.Children.Add($hdr)|Out-Null
    $inputs=@{}
    foreach ($row in @(
        @{L='CPU';V=$CpuThreshold},@{L='RAM';V=$RamThreshold},@{L='Disk';V=$DiskThreshold}
    )) {
        $g=[System.Windows.Controls.Grid]::new()
        $g.Margin=[System.Windows.Thickness]::new(0,0,0,6)
        $c1=[System.Windows.Controls.ColumnDefinition]::new()
        $c1.Width=[System.Windows.GridLength]::new(60)
        $c2=[System.Windows.Controls.ColumnDefinition]::new()
        $c2.Width=[System.Windows.GridLength]::new(1,'Star')
        $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2)
        $lbl=[System.Windows.Controls.TextBlock]::new()
        $lbl.Text=$row.L; $lbl.Foreground=$conv.ConvertFromString('#8B8FA8')
        $lbl.VerticalAlignment='Center'
        [System.Windows.Controls.Grid]::SetColumn($lbl,0); $g.Children.Add($lbl)|Out-Null
        $tb=[System.Windows.Controls.TextBox]::new()
        $tb.Text=$row.V; $tb.Background=$conv.ConvertFromString('#2D3148')
        $tb.Foreground=$conv.ConvertFromString('#E2E8F0')
        $tb.Padding=[System.Windows.Thickness]::new(7,4,7,4)
        [System.Windows.Controls.Grid]::SetColumn($tb,1); $g.Children.Add($tb)|Out-Null
        $inputs[$row.L]=$tb; $sp.Children.Add($g)|Out-Null
    }
    $sBtn=[System.Windows.Controls.Button]::new()
    $sBtn.Content='Save'; $sBtn.Margin=[System.Windows.Thickness]::new(0,12,0,0)
    $sBtn.Background=$conv.ConvertFromString('#6C63FF')
    $sBtn.Foreground=[System.Windows.Media.Brushes]::White
    $sBtn.Padding=[System.Windows.Thickness]::new(0,9,0,9)
    $sBtn.BorderThickness=[System.Windows.Thickness]::new(0)
    # Auto-export row
    $gae=[System.Windows.Controls.Grid]::new()
    $gae.Margin=[System.Windows.Thickness]::new(0,0,0,6)
    $cae1=[System.Windows.Controls.ColumnDefinition]::new(); $cae1.Width=[System.Windows.GridLength]::new(60)
    $cae2=[System.Windows.Controls.ColumnDefinition]::new(); $cae2.Width=[System.Windows.GridLength]::new(1,'Star')
    $gae.ColumnDefinitions.Add($cae1); $gae.ColumnDefinitions.Add($cae2)
    $lae=[System.Windows.Controls.TextBlock]::new(); $lae.Text='Auto HTML'
    $lae.Foreground=$conv.ConvertFromString('#8B8FA8'); $lae.VerticalAlignment='Center'
    [System.Windows.Controls.Grid]::SetColumn($lae,0); $gae.Children.Add($lae)|Out-Null
    $tbae=[System.Windows.Controls.TextBox]::new(); $tbae.Text=$script:AutoExportMins
    $tbae.Background=$conv.ConvertFromString('#2D3148'); $tbae.Foreground=$conv.ConvertFromString('#E2E8F0')
    $tbae.Padding=[System.Windows.Thickness]::new(7,4,7,4); $tbae.ToolTip='Auto-export HTML every N minutes (0=off)'
    [System.Windows.Controls.Grid]::SetColumn($tbae,1); $gae.Children.Add($tbae)|Out-Null
    $sp.Children.Add($gae)|Out-Null
    $laeNote=[System.Windows.Controls.TextBlock]::new()
    $laeNote.Text='(Auto HTML mins, 0=off)'
    $laeNote.Foreground=$conv.ConvertFromString('#6B7080'); $laeNote.FontSize=9
    $laeNote.Margin=[System.Windows.Thickness]::new(0,0,0,8)
    $sp.Children.Add($laeNote)|Out-Null

    $sBtn.Add_Click({
        $cv=[int]$inputs['CPU'].Text; $rv=[int]$inputs['RAM'].Text; $dv=[int]$inputs['Disk'].Text
        if ($cv -in 1..100) { $script:CpuThreshold=$cv }
        if ($rv -in 1..100) { $script:RamThreshold=$rv }
        if ($dv -in 1..100) { $script:DiskThreshold=$dv }
        $ae=[int]$tbae.Text
        if ($ae -ge 0) { $script:AutoExportMins=$ae; $script:LastAutoExport=[datetime]::Now }
        $sd.Close()
    })
    $sp.Children.Add($sBtn)|Out-Null
    $sd.Content=$sp; [void]$sd.ShowDialog()
}

# ── Remote host connect / disconnect ─────────────────────────────────────────
function Connect-RemoteHost {
    $host_ = $C['cmbRemoteHost'].Text.Trim()
    if ([string]::IsNullOrEmpty($host_) -or $host_ -eq 'localhost' -or $host_ -eq '.') {
        Disconnect-RemoteHost; return
    }
    $C['txtStatus'].Text      = "Connecting to $host_..."
    $C['ellStatus'].Fill      = $conv.ConvertFromString('#F97316')
    $C['btnConnect'].IsEnabled = $false
    try {
        if ($script:CimSession) { Remove-CimSession $script:CimSession -EA SilentlyContinue }
        $opt = New-CimSessionOption -Protocol Wsman -ErrorAction Stop
        $script:CimSession = New-CimSession -ComputerName $host_ -SessionOption $opt -ErrorAction Stop
        $script:RemoteHost = $host_

        # Fetch remote OS for title
        $remoteOs = (Get-CimInstance Win32_OperatingSystem -CimSession $script:CimSession -EA SilentlyContinue).Caption
        $C['txtHostname'].Text = "$host_  |  $remoteOs"

        $C['ellStatus'].Fill      = $conv.ConvertFromString('#10B981')
        $C['txtRemoteStatus'].Text = "Connected: $host_"
        $C['txtRemoteStatus'].Foreground = $conv.ConvertFromString('#10B981')
        $C['btnDisconnect'].IsEnabled    = $true
        $C['btnConnect'].IsEnabled       = $false

        # Remember hostname in dropdown
        $existing = @($C['cmbRemoteHost'].Items | ForEach-Object { "$_" })
        if ($existing -notcontains $host_) { $C['cmbRemoteHost'].Items.Add($host_) | Out-Null }

        # Reset CPU delta cache on new connection
        $script:PrevProcCpu = @{}
        $script:CachedAdapter = $null
        $script:CachedBattery = $null
        $script:CachedGateway = $host_

        Update-UI (Get-Metrics)
    } catch {
        $C['ellStatus'].Fill      = $conv.ConvertFromString('#EF4444')
        $C['txtRemoteStatus'].Text = "Failed: $($_.Exception.Message)"
        $C['txtRemoteStatus'].Foreground = $conv.ConvertFromString('#EF4444')
        $C['btnConnect'].IsEnabled = $true
        $script:CimSession = $null; $script:RemoteHost = ''
    }
}

function Disconnect-RemoteHost {
    if ($script:CimSession) { Remove-CimSession $script:CimSession -EA SilentlyContinue }
    $script:CimSession = $null; $script:RemoteHost = ''
    $C['ellStatus'].Fill      = $conv.ConvertFromString('#4B5563')
    $C['txtRemoteStatus'].Text = 'Local machine'
    $C['txtRemoteStatus'].Foreground = $conv.ConvertFromString('#8B8FA8')
    $C['btnDisconnect'].IsEnabled = $false
    $C['btnConnect'].IsEnabled    = $true
    $C['cmbRemoteHost'].Text      = ''
    $C['txtHostname'].Text        = "$([Environment]::MachineName)  |  $osName"
    $script:PrevProcCpu   = @{}
    $script:CachedAdapter = $null
    $script:CachedBattery = $null
    $script:CachedGateway = $null
    Update-UI (Get-Metrics)
}

# ── End task helper (shared by button and context menu) ────────────────────────
function Invoke-EndTask {
    $row = $C['gridProcs'].SelectedItem
    if ($null -eq $row) { $C['txtStatus'].Text = '[!] Select a process first'; return }
    try {
        if ($script:CimSession) {
            # Remote: use CIM to terminate
            $proc = Get-CimInstance Win32_Process -CimSession $script:CimSession `
                        -Filter "ProcessId=$($row.PID)" -ErrorAction Stop
            $result = Invoke-CimMethod -InputObject $proc -MethodName Terminate -ErrorAction Stop
            if ($result.ReturnValue -ne 0) { throw "WMI returned code $($result.ReturnValue)" }
        } else {
            (Get-Process -Id $row.PID -ErrorAction Stop).Kill()
        }
        $C['txtStatus'].Text = "OK  Ended: $($row.Name) (PID $($row.PID))"
    } catch { $C['txtStatus'].Text = "[!] Could not end $($row.Name): $($_.Exception.Message)" }
}

# ── Timer ──────────────────────────────────────────────────────────────────────
$Timer=[System.Windows.Threading.DispatcherTimer]::new()
$Timer.Interval=[TimeSpan]::FromSeconds($RefreshSeconds)
$Timer.Add_Tick({
    if (-not $C['chkAutoRefresh'].IsChecked) { return }
    try   { Update-UI (Get-Metrics) }
    catch { $C['txtStatus'].Text="[!] Error: $($_.Exception.Message)" }
})

# ── Wire controls ──────────────────────────────────────────────────────────────
$osName = Get-FriendlyOS
$C['txtHostname'].Text = "$([Environment]::MachineName)  |  $osName"
$C['ellStatus'].Fill = $conv.ConvertFromString('#4B5563')

$C['btnRefresh'].Add_Click({
    try   { Update-UI (Get-Metrics) }
    catch { $C['txtStatus'].Text="[!] Error: $($_.Exception.Message)" }
})
$C['chkAutoRefresh'].Add_Checked({   $C['txtStatus'].Text = 'Auto-refresh ON' })
$C['chkAutoRefresh'].Add_Unchecked({ $C['txtStatus'].Text = 'Auto-refresh OFF -- grid is frozen for selection' })
$C['btnConnect'].Add_Click({    Connect-RemoteHost })
$C['btnDisconnect'].Add_Click({ Disconnect-RemoteHost })
$C['cmbRemoteHost'].Add_KeyDown({
    param($s,$e)
    if ($e.Key -eq 'Return') { Connect-RemoteHost }
})

# Process filter
$C['txtProcFilter'].Add_TextChanged({
    if ($C['txtProcFilter'].Text -ne $C['txtProcFilter'].Tag) {
        $filterText = $C['txtProcFilter'].Text.Trim()
        if ([string]::IsNullOrEmpty($filterText)) { return }
        $current = $C['gridProcs'].ItemsSource
        if ($null -ne $current) {
            $filtered = @($current) | Where-Object {
                $_.Name -like "*$filterText*" -or "$($_.PID)" -like "*$filterText*"
            }
            $C['gridProcs'].ItemsSource = $filtered
        }
    }
})
$C['btnProcFilterClear'].Add_Click({
    $C['txtProcFilter'].Text = $C['txtProcFilter'].Tag
    $C['txtProcFilter'].Foreground = $conv.ConvertFromString('#4B5563')
})

$C['btnExportCsv'].Add_Click({  Export-ToCsv })
$C['btnExportMd'].Add_Click({   Export-ToMarkdown })
$C['btnExportHtml'].Add_Click({ Export-ToHtml })
$C['btnSettings'].Add_Click({   Show-SettingsDialog })
$C['btnEndTask'].Add_Click({    Invoke-EndTask })
$C['btnClearAlerts'].Add_Click({ $script:AlertLog.Clear(); $C['lstAlerts'].Items.Clear() })

$Window.Add_Loaded({
    $Timer.Start()
    try { Update-UI (Get-Metrics) } catch {}
})
$Window.Add_Closed({
    $Timer.Stop()
    $script:CpuCounter.Dispose()
})

$app = [System.Windows.Application]::new()
$app.Run($Window) | Out-Null