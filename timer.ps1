# 研修タイマー：PowerPoint などの画面の最前面に浮かぶタイマー（ロボットのイラスト付き）
# 使い方：クリック＝スタート／一時停止、ドラッグ＝移動、右クリック＝メニュー、ホイール＝±1分
# 起動例：powershell -STA -File timer.ps1 -Minutes 5 -Label "ワーク①"
param(
  [double]$Minutes = 5,
  [string]$Label = ""
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, Microsoft.VisualBasic

# 座標はすべてイラスト（1536×1024）の上の位置。Viewbox でまとめて縮小して表示する
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="研修タイマー" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="True" ResizeMode="NoResize" SizeToContent="WidthAndHeight"
        FontFamily="BIZ UDPGothic, Meiryo UI, Yu Gothic UI" UseLayoutRounding="True"
        ToolTip="クリック：スタート／一時停止　ドラッグ：移動　右クリック：メニュー　ホイール：±1分">
  <Viewbox x:Name="Box" Width="420" Margin="14">
    <Viewbox.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="3" Opacity="0.3" Color="#1B2F5E"/></Viewbox.Effect>
    <Grid Width="1536" Height="1024">
      <Grid.Clip><RectangleGeometry Rect="0,0,1536,1024" RadiusX="96" RadiusY="96"/></Grid.Clip>
      <Image x:Name="Bg" Stretch="Fill"/>
      <Canvas>
        <!-- 数字 -->
        <Viewbox Canvas.Left="262" Canvas.Top="528" Width="1018" Height="318">
          <TextBlock x:Name="TimeText" Text="05:00" FontFamily="Arial Rounded MT Bold, Segoe UI Black" FontWeight="Bold" Foreground="#1E2A4A"/>
        </Viewbox>
        <!-- 進み具合のバー -->
        <Rectangle Canvas.Left="300" Canvas.Top="850" Width="936" Height="18" RadiusX="9" RadiusY="9" Fill="#E2EAF5"/>
        <Rectangle x:Name="Bar" Canvas.Left="300" Canvas.Top="850" Width="936" Height="18" RadiusX="9" RadiusY="9" Fill="#3C8DEB"/>
        <!-- 名前 -->
        <Border x:Name="LabelChip" Canvas.Left="40" Canvas.Top="34" CornerRadius="40" Background="#E6FFFFFF" Padding="34,10,34,12">
          <TextBlock x:Name="LabelText" FontSize="54" FontWeight="Bold" Foreground="#1E2A4A"/>
        </Border>
        <!-- ロボットのふきだし（電球の上に重ねる） -->
        <Canvas x:Name="Bubble" Canvas.Left="1060" Canvas.Top="48" Visibility="Hidden">
          <Canvas.RenderTransform><ScaleTransform x:Name="BubbleScale" CenterX="210" CenterY="120"/></Canvas.RenderTransform>
          <Path x:Name="BubbleTail" Data="M60,190 L10,262 L130,206 Z" Fill="#FFFFFF" Stroke="#3C8DEB" StrokeThickness="8" StrokeLineJoin="Round"/>
          <Border x:Name="BubbleBody" Width="440" Height="230" CornerRadius="80" Background="#FFFFFF" BorderBrush="#3C8DEB" BorderThickness="8">
            <TextBlock x:Name="BubbleText" FontSize="76" FontWeight="Bold" Foreground="#1E2A4A" TextAlignment="Center"
                       HorizontalAlignment="Center" VerticalAlignment="Center" LineHeight="84"/>
          </Border>
          <Rectangle Canvas.Left="58" Canvas.Top="196" Width="64" Height="14" Fill="#FFFFFF"/>
        </Canvas>
      </Canvas>
    </Grid>
  </Viewbox>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Parse($xaml)
$ui = @{}
'Box','Bg','TimeText','Bar','LabelChip','LabelText','Bubble','BubbleScale','BubbleTail','BubbleBody','BubbleText' |
  ForEach-Object { $ui[$_] = $win.FindName($_) }

$imgPath = Join-Path $PSScriptRoot 'robot_bg.png'
if (Test-Path $imgPath) {
  $img = New-Object Windows.Media.Imaging.BitmapImage
  $img.BeginInit(); $img.CacheOption = 'OnLoad'; $img.UriSource = New-Object Uri $imgPath; $img.EndInit()
  $ui.Bg.Source = $img
} else {
  [void][Windows.MessageBox]::Show("イラストの画像（robot_bg.png）が見つかりません。`ntimer.ps1 と同じフォルダに置いてください。", '研修タイマー')
  exit
}

$bc = New-Object Windows.Media.BrushConverter
function Brush([string]$hex) { $bc.ConvertFromString($hex) }
$C = @{ blue = Brush '#3C8DEB'; amber = Brush '#E09A2D'; coral = Brush '#E0675A'; navy = Brush '#1E2A4A' }
$BAR_W = 936.0

# ---------- state ----------
$script:total = [math]::Max(1.0, [double]$Minutes) * 60000   # ms
$script:base = $script:total                                  # remaining at last start/pause
$script:state = 'stopped'                                     # stopped | running | paused | done
$script:sw = New-Object Diagnostics.Stopwatch
$script:warned = $false
$script:soundOn = $true

function Set-Label([string]$text) {
  $ui.LabelText.Text = $text
  $ui.LabelChip.Visibility = if ($text.Trim()) { 'Visible' } else { 'Collapsed' }
}
Set-Label $Label

function Remaining { if ($script:state -in 'running', 'done') { $script:base - $script:sw.Elapsed.TotalMilliseconds } else { $script:base } }
function Fmt([double]$ms) {
  $s = [math]::Ceiling([math]::Abs($ms) / 1000)
  if ($s -ge 3600) { '{0}:{1:00}:{2:00}' -f [math]::Floor($s / 3600), [math]::Floor(($s % 3600) / 60), ($s % 60) }
  else { '{0:00}:{1:00}' -f [math]::Floor($s / 60), ($s % 60) }
}

# ---------- sounds (generated, no files needed) ----------
function New-Wav($notes) {
  $rate = 22050
  $len = [int](($notes | ForEach-Object { $_[1] + $_[2] } | Measure-Object -Maximum).Maximum * $rate) + 2000
  $buf = New-Object double[] $len
  foreach ($n in $notes) {
    $f = $n[0]; $s0 = [int]($n[1] * $rate); $cnt = [int]($n[2] * $rate); $vol = $n[3]
    $w = 2 * [math]::PI * $f / $rate; $k = 5.0 / $cnt
    for ($i = 0; $i -lt $cnt; $i++) {
      $env = [math]::Exp(-$i * $k); if ($i -lt 200) { $env *= $i / 200 }
      $buf[$s0 + $i] += $vol * $env * [math]::Sin($w * $i)
    }
  }
  $ms = New-Object IO.MemoryStream
  $bw = New-Object IO.BinaryWriter $ms
  $bw.Write([Text.Encoding]::ASCII.GetBytes('RIFF')); $bw.Write([int](36 + $len * 2)); $bw.Write([Text.Encoding]::ASCII.GetBytes('WAVEfmt '))
  $bw.Write([int]16); $bw.Write([int16]1); $bw.Write([int16]1); $bw.Write([int]$rate); $bw.Write([int]($rate * 2)); $bw.Write([int16]2); $bw.Write([int16]16)
  $bw.Write([Text.Encoding]::ASCII.GetBytes('data')); $bw.Write([int]($len * 2))
  foreach ($v in $buf) { $bw.Write([int16]([math]::Max(-1.0, [math]::Min(1.0, [double]$v)) * 32000)) }
  $bw.Flush(); $ms.Position = 0
  $p = New-Object Media.SoundPlayer $ms; $p.Load(); $p
}
$sndDing = New-Wav @(, @(880, 0, 0.5, 0.35))
$sndChime = New-Wav @(@(784, 0, 0.9, 0.3), @(988, 0.18, 0.9, 0.3), @(1175, 0.36, 0.9, 0.3), @(1568, 0.54, 1.2, 0.3),
                      @(784, 1.3, 0.9, 0.2), @(988, 1.48, 0.9, 0.2), @(1175, 1.66, 0.9, 0.2), @(1568, 1.84, 1.4, 0.2))
function Play($snd) { if ($script:soundOn) { try { $snd.Play() } catch {} } }

# ---------- drawing ----------
function Render {
  $rem = Remaining
  if ($script:state -eq 'running' -and $rem -le 0) { Finish; $rem = Remaining }
  $done = $script:state -eq 'done'
  $prog = if ($done) { 1.0 } else { [math]::Min(1.0, [math]::Max(0.0, 1.0 - $rem / $script:total)) }
  $warn = (-not $done) -and $script:state -ne 'stopped' -and $rem -le 60000

  if ($warn -and -not $script:warned -and $script:state -eq 'running') { $script:warned = $true; Play $sndDing }

  $accent = if ($done) { $C.coral } elseif ($warn) { $C.amber } else { $C.blue }
  $ui.TimeText.Text = if ($done) { '+' + (Fmt (-$rem)) } else { Fmt $rem }
  $ui.TimeText.Foreground = if ($done) { $C.coral } else { $C.navy }
  $ui.Bar.Width = [math]::Max(0.0, $BAR_W * (1.0 - $prog))   # bar shrinks as time passes
  $ui.Bar.Fill = $accent

  $say = switch ($script:state) {
    'stopped' { "クリックで`nスタート" }
    'paused'  { "一時停止中" }
    'done'    { "おしまい！" }
    default   { if ($warn) { "あと少し！" } else { "" } }
  }
  $ui.Bubble.Visibility = if ($say) { 'Visible' } else { 'Hidden' }
  $ui.BubbleText.Text = $say
  $ui.BubbleText.FontSize = if ($script:state -eq 'stopped') { 62 } else { 76 }
  $ui.BubbleBody.BorderBrush = $accent; $ui.BubbleTail.Stroke = $accent
  $ui.BubbleText.Foreground = if ($done) { $C.coral } else { $C.navy }
}

function Start-Pulse {
  $anim = New-Object Windows.Media.Animation.DoubleAnimation(1.0, 1.08, [TimeSpan]::FromMilliseconds(420))
  $anim.AutoReverse = $true; $anim.RepeatBehavior = [Windows.Media.Animation.RepeatBehavior]::Forever
  $ui.BubbleScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleXProperty, $anim)
  $ui.BubbleScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleYProperty, $anim)
}
function Stop-Pulse {
  $ui.BubbleScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleXProperty, $null)
  $ui.BubbleScale.BeginAnimation([Windows.Media.ScaleTransform]::ScaleYProperty, $null)
}

# ---------- actions ----------
function Finish { $script:state = 'done'; Play $sndChime; Start-Pulse }
function Toggle {
  switch ($script:state) {
    'running' { $script:base = Remaining; $script:sw.Reset(); $script:state = 'paused' }
    'done'    { Reset-Timer; return }
    default   { $script:sw.Restart(); $script:state = 'running' }
  }
  Render
}
function Set-Duration([double]$sec) {
  $sec = [math]::Max(1.0, [math]::Min(10800.0, [math]::Round($sec)))
  $script:total = $sec * 1000; $script:base = $script:total
  $script:state = 'stopped'; $script:warned = $false; $script:sw.Reset(); Stop-Pulse; Render
}
function Reset-Timer { Set-Duration ($script:total / 1000) }
function Adjust([double]$deltaSec) {
  if ($script:state -eq 'done') { return }
  if ($script:state -eq 'stopped') { Set-Duration ([math]::Max(60.0, $script:total / 1000 + $deltaSec)); return }
  $d = $deltaSec * 1000
  $newRem = [math]::Max(1000.0, (Remaining) + $d)
  $script:base += ($newRem - (Remaining))
  $script:total = [math]::Max([double]($script:total + $d), [double]$newRem)
  if ($newRem -gt 60000) { $script:warned = $false }
  Render
}

# ---------- right-click menu ----------
function Add-Item($parent, $header, [scriptblock]$action) {
  $mi = New-Object Windows.Controls.MenuItem; $mi.Header = $header
  if ($action) { $mi.Add_Click($action) }
  [void]$parent.Items.Add($mi); $mi
}
$menu = New-Object Windows.Controls.ContextMenu
$mTime = Add-Item $menu '時間を選ぶ' $null
foreach ($m in 1, 3, 5, 6, 8, 10, 15, 20, 30) { (Add-Item $mTime "$m 分" { Set-Duration ($this.Tag * 60) }).Tag = $m }
[void](Add-Item $menu '時間を入力…' {
  $v = [Microsoft.VisualBasic.Interaction]::InputBox('時間を入れてください（例：7　または　2:30）', '研修タイマー', [string]([math]::Round($script:total / 60000, 2)))
  if ($v -match '^\s*(\d+)\s*[:：]\s*(\d{1,2})\s*$') { Set-Duration ([int]$Matches[1] * 60 + [int]$Matches[2]) }
  elseif ($v -match '^\s*\d+(\.\d+)?\s*$') { Set-Duration ([double]$v * 60) }
})
[void](Add-Item $menu '名前を付ける…' {
  $v = [Microsoft.VisualBasic.Interaction]::InputBox('タイマーの名前（空にすると名前を消します）', '研修タイマー', $ui.LabelText.Text)
  Set-Label $v
})
[void]$menu.Items.Add((New-Object Windows.Controls.Separator))
[void](Add-Item $menu 'スタート／一時停止（スペース）' { Toggle })
[void](Add-Item $menu 'リセット（R）' { Reset-Timer })
[void](Add-Item $menu '＋1分（↑）' { Adjust 60 })
[void](Add-Item $menu '−1分（↓）' { Adjust -60 })
[void]$menu.Items.Add((New-Object Windows.Controls.Separator))
$mSound = Add-Item $menu '音を鳴らす' { $script:soundOn = -not $script:soundOn; $mSound.IsChecked = $script:soundOn }
$mSound.IsChecked = $true
$mSize = Add-Item $menu '大きさ' $null
foreach ($s in @(@('小', 300), @('ふつう', 420), @('大', 600), @('特大', 840))) {
  (Add-Item $mSize $s[0] { $ui.Box.Width = $this.Tag }).Tag = [double]$s[1]
}
[void]$menu.Items.Add((New-Object Windows.Controls.Separator))
[void](Add-Item $menu '閉じる' { $win.Close() })
$win.ContextMenu = $menu

# ---------- mouse & keyboard ----------
$win.Add_MouseLeftButtonDown({
  $l = $win.Left; $t = $win.Top
  try { $win.DragMove() } catch {}
  if ([math]::Abs($win.Left - $l) -lt 3 -and [math]::Abs($win.Top - $t) -lt 3) { Toggle }
})
$win.Add_MouseWheel({ if ($_.Delta -gt 0) { Adjust 60 } else { Adjust -60 } })
$win.Add_KeyDown({
  switch ($_.Key) {
    'Space' { Toggle }
    'R'     { Reset-Timer }
    'Up'    { Adjust 60 }
    'Down'  { Adjust -60 }
  }
})
# keep it above the slide show even after clicking back into PowerPoint
$win.Add_Deactivated({ $win.Topmost = $false; $win.Topmost = $true })

$win.Add_Loaded({
  $wa = [Windows.SystemParameters]::WorkArea
  $win.Left = $wa.Right - $win.ActualWidth - 8
  $win.Top = $wa.Bottom - $win.ActualHeight - 8
})

$tick = New-Object Windows.Threading.DispatcherTimer
$tick.Interval = [TimeSpan]::FromMilliseconds(100)
$tick.Add_Tick({ Render })
$tick.Start()

Render
[void]$win.ShowDialog()
