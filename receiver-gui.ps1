# CDC File Receiver GUI + QR送信 (PowerShell WinForms)
# 実行: powershell -Command "iex (Get-Content $env:USERPROFILE\Desktop\receiver-gui.ps1 -Raw)"
# バッファ受信 : 'f' 送信 -> [CDCF][nameLen2][name][dataLen4][data][sum4]
# ストリーミング: 待機      -> [CDCS][nameLen2][name] 繰返し[chunkLen4][data] 終端[0][totalLen4]
# QR送信      : テキスト/ファイルをQRコード列にしてスマホ(qr-reader.html)へ渡す
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'CDC File Receiver'
$form.Size = New-Object System.Drawing.Size(1060, 700)
$form.StartPosition = 'CenterScreen'

$script:recvStop = $false

$lblPort = New-Object System.Windows.Forms.Label
$lblPort.Text = 'COM:'; $lblPort.Location = '12,16'; $lblPort.AutoSize = $true
$form.Controls.Add($lblPort)

$cmbPort = New-Object System.Windows.Forms.ComboBox
$cmbPort.Location = '52,12'; $cmbPort.Width = 110; $cmbPort.DropDownStyle = 'DropDownList'
$form.Controls.Add($cmbPort)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = '更新'; $btnRefresh.Location = '170,10'; $btnRefresh.Width = 55
$form.Controls.Add($btnRefresh)

$rbBuf = New-Object System.Windows.Forms.RadioButton
$rbBuf.Text = 'バッファ受信(f)'; $rbBuf.Location = '240,4'; $rbBuf.AutoSize = $true
$form.Controls.Add($rbBuf)

$rbStream = New-Object System.Windows.Forms.RadioButton
$rbStream.Text = 'ストリーミング受信'; $rbStream.Location = '240,26'; $rbStream.AutoSize = $true; $rbStream.Checked = $true
$form.Controls.Add($rbStream)

$chkLoop = New-Object System.Windows.Forms.CheckBox
$chkLoop.Text = '連続受信(複数)'; $chkLoop.Location = '400,15'; $chkLoop.AutoSize = $true
$form.Controls.Add($chkLoop)

$btnRecv = New-Object System.Windows.Forms.Button
$btnRecv.Text = '受信開始'; $btnRecv.Location = '12,52'; $btnRecv.Width = 130; $btnRecv.Height = 40
$form.Controls.Add($btnRecv)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '停止'; $btnStop.Location = '150,52'; $btnStop.Width = 80; $btnStop.Height = 40; $btnStop.Enabled = $false
$form.Controls.Add($btnStop)

$pb = New-Object System.Windows.Forms.ProgressBar
$pb.Location = '12,104'; $pb.Width = 522; $pb.Height = 22
$form.Controls.Add($pb)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = '待機中'; $lblStatus.Location = '12,132'; $lblStatus.Width = 522; $lblStatus.AutoSize = $false
$form.Controls.Add($lblStatus)

$log = New-Object System.Windows.Forms.TextBox
$log.Multiline = $true; $log.ScrollBars = 'Vertical'; $log.Location = '12,158'; $log.Width = 522; $log.Height = 220; $log.ReadOnly = $true
$form.Controls.Add($log)

function Log($m) { $log.AppendText($m + "`r`n") }

function RefreshPorts {
  $cmbPort.Items.Clear()
  foreach ($p in [System.IO.Ports.SerialPort]::GetPortNames()) { [void]$cmbPort.Items.Add($p) }
  if ($cmbPort.Items.Count -gt 0) { $cmbPort.SelectedIndex = 0 }
}
$btnRefresh.Add_Click({ RefreshPorts })

# n バイトちょうど読む。UI を止めない。data 無しが deadlineSec 続いたら timeout。
function ReadExact($sp, $n, $deadlineSec) {
  $b = New-Object byte[] $n; $off = 0; $dl = (Get-Date).AddSeconds($deadlineSec)
  while ($off -lt $n) {
    try { $r = $sp.Read($b, $off, $n - $off); if ($r -gt 0) { $off += $r; $dl = (Get-Date).AddSeconds($deadlineSec) } }
    catch [System.TimeoutException] { }
    [System.Windows.Forms.Application]::DoEvents()
    if ($script:recvStop) { throw 'stopped' }
    if ((Get-Date) -gt $dl) { throw 'timeout' }
  }
  return ,$b
}

function SyncMagic($sp, $m0, $m1, $m2, $m3, $deadlineSec) {
  $w = New-Object System.Collections.Generic.List[byte]
  while ($true) {
    $x = (ReadExact $sp 1 $deadlineSec)[0]
    $w.Add($x); if ($w.Count -gt 4) { $w.RemoveAt(0) }
    if ($w.Count -eq 4 -and $w[0] -eq $m0 -and $w[1] -eq $m1 -and $w[2] -eq $m2 -and $w[3] -eq $m3) { return }
  }
}

# 保存先の重複回避: Desktop\name が存在すれば "base (2).ext" 等にする
function Get-UniqueSavePath([string]$dir, [string]$name) {
  $path = [System.IO.Path]::Combine($dir, $name)
  if (-not (Test-Path -LiteralPath $path)) { return $path }
  $base = [System.IO.Path]::GetFileNameWithoutExtension($name)
  $ext = [System.IO.Path]::GetExtension($name)
  $i = 2
  while ($true) {
    $path = [System.IO.Path]::Combine($dir, ('{0} ({1}){2}' -f $base, $i, $ext))
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $i++
  }
}

$btnStop.Add_Click({ $script:recvStop = $true; Log '停止要求...' })

$btnRecv.Add_Click({
  if ($cmbPort.SelectedItem -eq $null) { Log 'COMポートを選択してください'; return }
  $script:recvStop = $false
  $btnRecv.Enabled = $false; $btnStop.Enabled = $true; $pb.Value = 0; $pb.Style = 'Blocks'; $log.Clear()
  $sp = $null; $fileCount = 0
  try {
    $sp = New-Object System.IO.Ports.SerialPort $cmbPort.SelectedItem,115200,'None',8,'One'
    $sp.DtrEnable = $true; $sp.ReadTimeout = 200; $sp.Open()
    Start-Sleep -Milliseconds 200
    do {
      $pb.Style = 'Blocks'; $pb.Value = 0
      try {
        if ($rbBuf.Checked) {
          $sp.Write('f'); Log 'f 送信、受信中...'
          SyncMagic $sp 67 68 67 70 8
          $nl = [BitConverter]::ToUInt16((ReadExact $sp 2 5), 0)
          $name = [System.Text.Encoding]::UTF8.GetString((ReadExact $sp $nl 5), 0, $nl)
          $len = [BitConverter]::ToUInt32((ReadExact $sp 4 5), 0)
          Log "file=$name  len=$len"
          $path = Get-UniqueSavePath ($env:USERPROFILE + '\Desktop') $name
          $fs = [System.IO.File]::Open($path, 'Create'); $sum = 0; $got = 0
          try {
            while ($got -lt $len) {
              $c = [Math]::Min(65536, $len - $got)
              $chunk = ReadExact $sp $c 8; $fs.Write($chunk, 0, $c)
              for ($i = 0; $i -lt $c; $i++) { $sum = ($sum + $chunk[$i]) -band 0xFFFFFFFF }
              $got += $c; $pb.Value = [int](100 * $got / $len); $lblStatus.Text = "$got / $len bytes"; [System.Windows.Forms.Application]::DoEvents()
            }
            $sumR = [BitConverter]::ToUInt32((ReadExact $sp 4 5), 0)
          } finally { $fs.Close() }
          if ($sum -eq $sumR) { $lblStatus.Text = "完了: $name"; Log "*** SAVED: $path ***" } else { $lblStatus.Text = '照合NG'; Log "FAIL: sum $sum vs $sumR" }
        } else {
          Log 'ストリーミング待機... スマホで「ストリーミング送信」してください'
          $pb.Style = 'Marquee'
          SyncMagic $sp 67 68 67 83 180
          $nl = [BitConverter]::ToUInt16((ReadExact $sp 2 5), 0)
          $name = [System.Text.Encoding]::UTF8.GetString((ReadExact $sp $nl 5), 0, $nl)
          Log "file=$name  (streaming)"
          $path = Get-UniqueSavePath ($env:USERPROFILE + '\Desktop') $name
          $fs = [System.IO.File]::Open($path, 'Create'); $total = 0
          try {
            while ($true) {
              $cl = [BitConverter]::ToUInt32((ReadExact $sp 4 10), 0)
              if ($cl -eq 0) { break }
              $chunk = ReadExact $sp $cl 10; $fs.Write($chunk, 0, $cl)
              $total += $cl; $lblStatus.Text = "$total bytes"; [System.Windows.Forms.Application]::DoEvents()
            }
            $tl = [BitConverter]::ToUInt32((ReadExact $sp 4 5), 0)
          } finally { $fs.Close() }
          $pb.Style = 'Blocks'; $pb.Value = 100
          if ($total -eq $tl) { $lblStatus.Text = "完了: $name ($total bytes)"; Log "*** SAVED: $path ***" } else { $lblStatus.Text = '長さNG'; Log "FAIL: length $total vs $tl" }
        }
        $fileCount++
        Log ('受信 ' + $fileCount + ' 件目 完了')
      } catch {
        if ($_.Exception.Message -eq 'stopped') { Log '停止しました'; break }
        throw
      }
    } while ($chkLoop.Checked -and -not $script:recvStop)
  } catch {
    Log ('エラー: ' + $_.Exception.Message)
  } finally {
    if ($sp -ne $null -and $sp.IsOpen) { $sp.Close() }
    $pb.Style = 'Blocks'
    $btnRecv.Enabled = $true; $btnStop.Enabled = $false; $script:recvStop = $false
    Log ('受信終了 (計 ' + $fileCount + ' 件)')
  }
})


# ================= カルテ機→スマホ (QR送信) =================
# コンテナ: ASCIIヘッダ "CDCQ1 <T|F> <0|1> <nameB64> <origSize>" + 0x0A + データ(生 or gzip)
# b64all = Base64(コンテナ), id8 = SHA256(b64all文字列のUTF8バイト) 先頭8桁hex(小文字)
# QRテキスト: Q1|<id8>|<idx>|<total>|<b64allの120文字チャンク>  (idxは0始まり, ECC=M)
# エンコーダ: QRCoder v1.4.2 をC#5互換に改変したもの (qrencoder/ の3ファイルを結合埋め込み)

$script:qrDropFile = $null

$qrEncoderSource = @'
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

// ==== QRCodeGenerator.cs ====
namespace QRCoder
{
    public class QRCodeGenerator : IDisposable
    {
        private static readonly char[] alphanumEncTable = { ' ', '$', '%', '*', '+', '-', '.', '/', ':' };
        private static readonly int[] capacityBaseValues = { 41, 25, 17, 10, 34, 20, 14, 8, 27, 16, 11, 7, 17, 10, 7, 4, 77, 47, 32, 20, 63, 38, 26, 16, 48, 29, 20, 12, 34, 20, 14, 8, 127, 77, 53, 32, 101, 61, 42, 26, 77, 47, 32, 20, 58, 35, 24, 15, 187, 114, 78, 48, 149, 90, 62, 38, 111, 67, 46, 28, 82, 50, 34, 21, 255, 154, 106, 65, 202, 122, 84, 52, 144, 87, 60, 37, 106, 64, 44, 27, 322, 195, 134, 82, 255, 154, 106, 65, 178, 108, 74, 45, 139, 84, 58, 36, 370, 224, 154, 95, 293, 178, 122, 75, 207, 125, 86, 53, 154, 93, 64, 39, 461, 279, 192, 118, 365, 221, 152, 93, 259, 157, 108, 66, 202, 122, 84, 52, 552, 335, 230, 141, 432, 262, 180, 111, 312, 189, 130, 80, 235, 143, 98, 60, 652, 395, 271, 167, 513, 311, 213, 131, 364, 221, 151, 93, 288, 174, 119, 74, 772, 468, 321, 198, 604, 366, 251, 155, 427, 259, 177, 109, 331, 200, 137, 85, 883, 535, 367, 226, 691, 419, 287, 177, 489, 296, 203, 125, 374, 227, 155, 96, 1022, 619, 425, 262, 796, 483, 331, 204, 580, 352, 241, 149, 427, 259, 177, 109, 1101, 667, 458, 282, 871, 528, 362, 223, 621, 376, 258, 159, 468, 283, 194, 120, 1250, 758, 520, 320, 991, 600, 412, 254, 703, 426, 292, 180, 530, 321, 220, 136, 1408, 854, 586, 361, 1082, 656, 450, 277, 775, 470, 322, 198, 602, 365, 250, 154, 1548, 938, 644, 397, 1212, 734, 504, 310, 876, 531, 364, 224, 674, 408, 280, 173, 1725, 1046, 718, 442, 1346, 816, 560, 345, 948, 574, 394, 243, 746, 452, 310, 191, 1903, 1153, 792, 488, 1500, 909, 624, 384, 1063, 644, 442, 272, 813, 493, 338, 208, 2061, 1249, 858, 528, 1600, 970, 666, 410, 1159, 702, 482, 297, 919, 557, 382, 235, 2232, 1352, 929, 572, 1708, 1035, 711, 438, 1224, 742, 509, 314, 969, 587, 403, 248, 2409, 1460, 1003, 618, 1872, 1134, 779, 480, 1358, 823, 565, 348, 1056, 640, 439, 270, 2620, 1588, 1091, 672, 2059, 1248, 857, 528, 1468, 890, 611, 376, 1108, 672, 461, 284, 2812, 1704, 1171, 721, 2188, 1326, 911, 561, 1588, 963, 661, 407, 1228, 744, 511, 315, 3057, 1853, 1273, 784, 2395, 1451, 997, 614, 1718, 1041, 715, 440, 1286, 779, 535, 330, 3283, 1990, 1367, 842, 2544, 1542, 1059, 652, 1804, 1094, 751, 462, 1425, 864, 593, 365, 3517, 2132, 1465, 902, 2701, 1637, 1125, 692, 1933, 1172, 805, 496, 1501, 910, 625, 385, 3669, 2223, 1528, 940, 2857, 1732, 1190, 732, 2085, 1263, 868, 534, 1581, 958, 658, 405, 3909, 2369, 1628, 1002, 3035, 1839, 1264, 778, 2181, 1322, 908, 559, 1677, 1016, 698, 430, 4158, 2520, 1732, 1066, 3289, 1994, 1370, 843, 2358, 1429, 982, 604, 1782, 1080, 742, 457, 4417, 2677, 1840, 1132, 3486, 2113, 1452, 894, 2473, 1499, 1030, 634, 1897, 1150, 790, 486, 4686, 2840, 1952, 1201, 3693, 2238, 1538, 947, 2670, 1618, 1112, 684, 2022, 1226, 842, 518, 4965, 3009, 2068, 1273, 3909, 2369, 1628, 1002, 2805, 1700, 1168, 719, 2157, 1307, 898, 553, 5253, 3183, 2188, 1347, 4134, 2506, 1722, 1060, 2949, 1787, 1228, 756, 2301, 1394, 958, 590, 5529, 3351, 2303, 1417, 4343, 2632, 1809, 1113, 3081, 1867, 1283, 790, 2361, 1431, 983, 605, 5836, 3537, 2431, 1496, 4588, 2780, 1911, 1176, 3244, 1966, 1351, 832, 2524, 1530, 1051, 647, 6153, 3729, 2563, 1577, 4775, 2894, 1989, 1224, 3417, 2071, 1423, 876, 2625, 1591, 1093, 673, 6479, 3927, 2699, 1661, 5039, 3054, 2099, 1292, 3599, 2181, 1499, 923, 2735, 1658, 1139, 701, 6743, 4087, 2809, 1729, 5313, 3220, 2213, 1362, 3791, 2298, 1579, 972, 2927, 1774, 1219, 750, 7089, 4296, 2953, 1817, 5596, 3391, 2331, 1435, 3993, 2420, 1663, 1024, 3057, 1852, 1273, 784 };
        private static readonly int[] capacityECCBaseValues = { 19, 7, 1, 19, 0, 0, 16, 10, 1, 16, 0, 0, 13, 13, 1, 13, 0, 0, 9, 17, 1, 9, 0, 0, 34, 10, 1, 34, 0, 0, 28, 16, 1, 28, 0, 0, 22, 22, 1, 22, 0, 0, 16, 28, 1, 16, 0, 0, 55, 15, 1, 55, 0, 0, 44, 26, 1, 44, 0, 0, 34, 18, 2, 17, 0, 0, 26, 22, 2, 13, 0, 0, 80, 20, 1, 80, 0, 0, 64, 18, 2, 32, 0, 0, 48, 26, 2, 24, 0, 0, 36, 16, 4, 9, 0, 0, 108, 26, 1, 108, 0, 0, 86, 24, 2, 43, 0, 0, 62, 18, 2, 15, 2, 16, 46, 22, 2, 11, 2, 12, 136, 18, 2, 68, 0, 0, 108, 16, 4, 27, 0, 0, 76, 24, 4, 19, 0, 0, 60, 28, 4, 15, 0, 0, 156, 20, 2, 78, 0, 0, 124, 18, 4, 31, 0, 0, 88, 18, 2, 14, 4, 15, 66, 26, 4, 13, 1, 14, 194, 24, 2, 97, 0, 0, 154, 22, 2, 38, 2, 39, 110, 22, 4, 18, 2, 19, 86, 26, 4, 14, 2, 15, 232, 30, 2, 116, 0, 0, 182, 22, 3, 36, 2, 37, 132, 20, 4, 16, 4, 17, 100, 24, 4, 12, 4, 13, 274, 18, 2, 68, 2, 69, 216, 26, 4, 43, 1, 44, 154, 24, 6, 19, 2, 20, 122, 28, 6, 15, 2, 16, 324, 20, 4, 81, 0, 0, 254, 30, 1, 50, 4, 51, 180, 28, 4, 22, 4, 23, 140, 24, 3, 12, 8, 13, 370, 24, 2, 92, 2, 93, 290, 22, 6, 36, 2, 37, 206, 26, 4, 20, 6, 21, 158, 28, 7, 14, 4, 15, 428, 26, 4, 107, 0, 0, 334, 22, 8, 37, 1, 38, 244, 24, 8, 20, 4, 21, 180, 22, 12, 11, 4, 12, 461, 30, 3, 115, 1, 116, 365, 24, 4, 40, 5, 41, 261, 20, 11, 16, 5, 17, 197, 24, 11, 12, 5, 13, 523, 22, 5, 87, 1, 88, 415, 24, 5, 41, 5, 42, 295, 30, 5, 24, 7, 25, 223, 24, 11, 12, 7, 13, 589, 24, 5, 98, 1, 99, 453, 28, 7, 45, 3, 46, 325, 24, 15, 19, 2, 20, 253, 30, 3, 15, 13, 16, 647, 28, 1, 107, 5, 108, 507, 28, 10, 46, 1, 47, 367, 28, 1, 22, 15, 23, 283, 28, 2, 14, 17, 15, 721, 30, 5, 120, 1, 121, 563, 26, 9, 43, 4, 44, 397, 28, 17, 22, 1, 23, 313, 28, 2, 14, 19, 15, 795, 28, 3, 113, 4, 114, 627, 26, 3, 44, 11, 45, 445, 26, 17, 21, 4, 22, 341, 26, 9, 13, 16, 14, 861, 28, 3, 107, 5, 108, 669, 26, 3, 41, 13, 42, 485, 30, 15, 24, 5, 25, 385, 28, 15, 15, 10, 16, 932, 28, 4, 116, 4, 117, 714, 26, 17, 42, 0, 0, 512, 28, 17, 22, 6, 23, 406, 30, 19, 16, 6, 17, 1006, 28, 2, 111, 7, 112, 782, 28, 17, 46, 0, 0, 568, 30, 7, 24, 16, 25, 442, 24, 34, 13, 0, 0, 1094, 30, 4, 121, 5, 122, 860, 28, 4, 47, 14, 48, 614, 30, 11, 24, 14, 25, 464, 30, 16, 15, 14, 16, 1174, 30, 6, 117, 4, 118, 914, 28, 6, 45, 14, 46, 664, 30, 11, 24, 16, 25, 514, 30, 30, 16, 2, 17, 1276, 26, 8, 106, 4, 107, 1000, 28, 8, 47, 13, 48, 718, 30, 7, 24, 22, 25, 538, 30, 22, 15, 13, 16, 1370, 28, 10, 114, 2, 115, 1062, 28, 19, 46, 4, 47, 754, 28, 28, 22, 6, 23, 596, 30, 33, 16, 4, 17, 1468, 30, 8, 122, 4, 123, 1128, 28, 22, 45, 3, 46, 808, 30, 8, 23, 26, 24, 628, 30, 12, 15, 28, 16, 1531, 30, 3, 117, 10, 118, 1193, 28, 3, 45, 23, 46, 871, 30, 4, 24, 31, 25, 661, 30, 11, 15, 31, 16, 1631, 30, 7, 116, 7, 117, 1267, 28, 21, 45, 7, 46, 911, 30, 1, 23, 37, 24, 701, 30, 19, 15, 26, 16, 1735, 30, 5, 115, 10, 116, 1373, 28, 19, 47, 10, 48, 985, 30, 15, 24, 25, 25, 745, 30, 23, 15, 25, 16, 1843, 30, 13, 115, 3, 116, 1455, 28, 2, 46, 29, 47, 1033, 30, 42, 24, 1, 25, 793, 30, 23, 15, 28, 16, 1955, 30, 17, 115, 0, 0, 1541, 28, 10, 46, 23, 47, 1115, 30, 10, 24, 35, 25, 845, 30, 19, 15, 35, 16, 2071, 30, 17, 115, 1, 116, 1631, 28, 14, 46, 21, 47, 1171, 30, 29, 24, 19, 25, 901, 30, 11, 15, 46, 16, 2191, 30, 13, 115, 6, 116, 1725, 28, 14, 46, 23, 47, 1231, 30, 44, 24, 7, 25, 961, 30, 59, 16, 1, 17, 2306, 30, 12, 121, 7, 122, 1812, 28, 12, 47, 26, 48, 1286, 30, 39, 24, 14, 25, 986, 30, 22, 15, 41, 16, 2434, 30, 6, 121, 14, 122, 1914, 28, 6, 47, 34, 48, 1354, 30, 46, 24, 10, 25, 1054, 30, 2, 15, 64, 16, 2566, 30, 17, 122, 4, 123, 1992, 28, 29, 46, 14, 47, 1426, 30, 49, 24, 10, 25, 1096, 30, 24, 15, 46, 16, 2702, 30, 4, 122, 18, 123, 2102, 28, 13, 46, 32, 47, 1502, 30, 48, 24, 14, 25, 1142, 30, 42, 15, 32, 16, 2812, 30, 20, 117, 4, 118, 2216, 28, 40, 47, 7, 48, 1582, 30, 43, 24, 22, 25, 1222, 30, 10, 15, 67, 16, 2956, 30, 19, 118, 6, 119, 2334, 28, 18, 47, 31, 48, 1666, 30, 34, 24, 34, 25, 1276, 30, 20, 15, 61, 16 };
        private static readonly int[] alignmentPatternBaseValues = { 0, 0, 0, 0, 0, 0, 0, 6, 18, 0, 0, 0, 0, 0, 6, 22, 0, 0, 0, 0, 0, 6, 26, 0, 0, 0, 0, 0, 6, 30, 0, 0, 0, 0, 0, 6, 34, 0, 0, 0, 0, 0, 6, 22, 38, 0, 0, 0, 0, 6, 24, 42, 0, 0, 0, 0, 6, 26, 46, 0, 0, 0, 0, 6, 28, 50, 0, 0, 0, 0, 6, 30, 54, 0, 0, 0, 0, 6, 32, 58, 0, 0, 0, 0, 6, 34, 62, 0, 0, 0, 0, 6, 26, 46, 66, 0, 0, 0, 6, 26, 48, 70, 0, 0, 0, 6, 26, 50, 74, 0, 0, 0, 6, 30, 54, 78, 0, 0, 0, 6, 30, 56, 82, 0, 0, 0, 6, 30, 58, 86, 0, 0, 0, 6, 34, 62, 90, 0, 0, 0, 6, 28, 50, 72, 94, 0, 0, 6, 26, 50, 74, 98, 0, 0, 6, 30, 54, 78, 102, 0, 0, 6, 28, 54, 80, 106, 0, 0, 6, 32, 58, 84, 110, 0, 0, 6, 30, 58, 86, 114, 0, 0, 6, 34, 62, 90, 118, 0, 0, 6, 26, 50, 74, 98, 122, 0, 6, 30, 54, 78, 102, 126, 0, 6, 26, 52, 78, 104, 130, 0, 6, 30, 56, 82, 108, 134, 0, 6, 34, 60, 86, 112, 138, 0, 6, 30, 58, 86, 114, 142, 0, 6, 34, 62, 90, 118, 146, 0, 6, 30, 54, 78, 102, 126, 150, 6, 24, 50, 76, 102, 128, 154, 6, 28, 54, 80, 106, 132, 158, 6, 32, 58, 84, 110, 136, 162, 6, 26, 54, 82, 110, 138, 166, 6, 30, 58, 86, 114, 142, 170 };
        private static readonly int[] remainderBits = { 0, 7, 7, 7, 7, 7, 0, 0, 0, 0, 0, 0, 0, 3, 3, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 4, 4, 3, 3, 3, 3, 3, 3, 3, 0, 0, 0, 0, 0, 0 };

        private static readonly List<AlignmentPattern> alignmentPatternTable = CreateAlignmentPatternTable();
        private static readonly List<ECCInfo> capacityECCTable = CreateCapacityECCTable();
        private static readonly List<VersionInfo> capacityTable = CreateCapacityTable();
        private static readonly List<Antilog> galoisField = CreateAntilogTable();
        private static readonly Dictionary<char, int> alphanumEncDict = CreateAlphanumEncDict();

        public enum EciMode
        {
            Default = 0,
            Iso8859_1 = 3,
            Iso8859_2 = 4,
            Utf8 = 26
        }

        /// <summary>
        /// Initializes the QR code generator
        /// </summary>
        public QRCodeGenerator()
        {
        }

        /// <summary>
        /// Calculates the QR code data which than can be used in one of the rendering classes to generate a graphical representation.
        /// </summary>
        /// <param name="plainText">The payload which shall be encoded in the QR code</param>
        /// <param name="eccLevel">The level of error correction data</param>
        /// <param name="forceUtf8">Shall the generator be forced to work in UTF-8 mode?</param>
        /// <param name="utf8BOM">Should the byte-order-mark be used?</param>
        /// <param name="eciMode">Which ECI mode shall be used?</param>
        /// <param name="requestedVersion">Set fixed QR code target version.</param>
        /// <exception cref="QRCoder.Exceptions.DataTooLongException">Thrown when the payload is too big to be encoded in a QR code.</exception>
        /// <returns>Returns the raw QR code data which can be used for rendering.</returns>
        public QRCodeData CreateQrCode(string plainText, ECCLevel eccLevel, bool forceUtf8 = false, bool utf8BOM = false, EciMode eciMode = EciMode.Default, int requestedVersion = -1)
        {
            return GenerateQrCode(plainText, eccLevel, forceUtf8, utf8BOM, eciMode, requestedVersion);
        }

        /// <summary>
        /// Calculates the QR code data which than can be used in one of the rendering classes to generate a graphical representation.
        /// </summary>
        /// <param name="binaryData">A byte array which shall be encoded/stored in the QR code</param>
        /// <param name="eccLevel">The level of error correction data</param>
        /// <exception cref="QRCoder.Exceptions.DataTooLongException">Thrown when the payload is too big to be encoded in a QR code.</exception>
        /// <returns>Returns the raw QR code data which can be used for rendering.</returns>
        public QRCodeData CreateQrCode(byte[] binaryData, ECCLevel eccLevel)
        {
            return GenerateQrCode(binaryData, eccLevel);
        }

        /// <summary>
        /// Calculates the QR code data which than can be used in one of the rendering classes to generate a graphical representation.
        /// </summary>
        /// <param name="plainText">The payload which shall be encoded in the QR code</param>
        /// <param name="eccLevel">The level of error correction data</param>
        /// <param name="forceUtf8">Shall the generator be forced to work in UTF-8 mode?</param>
        /// <param name="utf8BOM">Should the byte-order-mark be used?</param>
        /// <param name="eciMode">Which ECI mode shall be used?</param>
        /// <param name="requestedVersion">Set fixed QR code target version.</param>
        /// <exception cref="QRCoder.Exceptions.DataTooLongException">Thrown when the payload is too big to be encoded in a QR code.</exception>
        /// <returns>Returns the raw QR code data which can be used for rendering.</returns>
        public static QRCodeData GenerateQrCode(string plainText, ECCLevel eccLevel, bool forceUtf8 = false, bool utf8BOM = false, EciMode eciMode = EciMode.Default, int requestedVersion = -1)
        {
            EncodingMode encoding = GetEncodingFromPlaintext(plainText, forceUtf8);
            var codedText = PlainTextToBinary(plainText, encoding, eciMode, utf8BOM, forceUtf8);
            var dataInputLength = GetDataLength(encoding, plainText, codedText, forceUtf8);
            int version = requestedVersion;
            if (version == -1)
            {
                version = GetVersion(dataInputLength+(eciMode != EciMode.Default?2:0), encoding, eccLevel);
            }
            else
            {
                //Version was passed as fixed version via parameter. Thus let's check if chosen version is valid.
                var minVersion = GetVersion(dataInputLength + (eciMode != EciMode.Default ? 2 : 0), encoding, eccLevel);
                if (minVersion > version)
                {
                    var maxSizeByte = capacityTable[version - 1].Details.First(x => x.ErrorCorrectionLevel == eccLevel).CapacityDict[encoding];
                    throw new QRCoder.Exceptions.DataTooLongException(eccLevel.ToString(), encoding.ToString(), version, maxSizeByte);
                }                    
            }

            string modeIndicator = String.Empty;
            if (eciMode != EciMode.Default)
            {
                modeIndicator = DecToBin((int)EncodingMode.ECI, 4);
                modeIndicator += DecToBin((int)eciMode, 8);
            }
            modeIndicator += DecToBin((int)encoding, 4);
            var countIndicator = DecToBin(dataInputLength, GetCountIndicatorLength(version, encoding));
            var bitString = modeIndicator + countIndicator;

            bitString += codedText;

            return GenerateQrCode(bitString, eccLevel, version);
        }


        /// <summary>
        /// Calculates the QR code data which than can be used in one of the rendering classes to generate a graphical representation.
        /// </summary>
        /// <param name="binaryData">A byte array which shall be encoded/stored in the QR code</param>
        /// <param name="eccLevel">The level of error correction data</param>
        /// <exception cref="QRCoder.Exceptions.DataTooLongException">Thrown when the payload is too big to be encoded in a QR code.</exception>
        /// <returns>Returns the raw QR code data which can be used for rendering.</returns>
        public static QRCodeData GenerateQrCode(byte[] binaryData, ECCLevel eccLevel)
        {
            int version = GetVersion(binaryData.Length, EncodingMode.Byte, eccLevel);

            string modeIndicator = DecToBin((int)EncodingMode.Byte, 4);
            string countIndicator = DecToBin(binaryData.Length, GetCountIndicatorLength(version, EncodingMode.Byte));

            string bitString = modeIndicator + countIndicator;
            foreach (byte b in binaryData)
            {
                bitString += DecToBin(b, 8);
            }

            return GenerateQrCode(bitString, eccLevel, version);
        }

        private static QRCodeData GenerateQrCode(string bitString, ECCLevel eccLevel, int version)
        {
            //Fill up data code word
            var eccInfo = capacityECCTable.Single(x => x.Version == version && x.ErrorCorrectionLevel == eccLevel);
            var dataLength = eccInfo.TotalDataCodewords * 8;
            var lengthDiff = dataLength - bitString.Length;
            if (lengthDiff > 0)
                bitString += new string('0', Math.Min(lengthDiff, 4));
            if ((bitString.Length % 8) != 0)
                bitString += new string('0', 8 - (bitString.Length % 8));
            while (bitString.Length < dataLength)
                bitString += "1110110000010001";
            if (bitString.Length > dataLength)
                bitString = bitString.Substring(0, dataLength);

            //Calculate error correction words
            var codeWordWithECC = new List<CodewordBlock>(eccInfo.BlocksInGroup1 + eccInfo.BlocksInGroup2);
            for (var i = 0; i < eccInfo.BlocksInGroup1; i++)
            {
                var bitStr = bitString.Substring(i * eccInfo.CodewordsInGroup1 * 8, eccInfo.CodewordsInGroup1 * 8);
                var bitBlockList = BinaryStringToBitBlockList(bitStr);
                var bitBlockListDec = BinaryStringListToDecList(bitBlockList);
                var eccWordList = CalculateECCWords(bitStr, eccInfo);
                var eccWordListDec = BinaryStringListToDecList(eccWordList);
                codeWordWithECC.Add(
                    new CodewordBlock(1,
                                      i + 1,
                                      bitStr,
                                      bitBlockList,
                                      eccWordList,
                                      bitBlockListDec,
                                      eccWordListDec)
                                );
            }
            bitString = bitString.Substring(eccInfo.BlocksInGroup1 * eccInfo.CodewordsInGroup1 * 8);
            for (var i = 0; i < eccInfo.BlocksInGroup2; i++)
            {
                var bitStr = bitString.Substring(i * eccInfo.CodewordsInGroup2 * 8, eccInfo.CodewordsInGroup2 * 8);
                var bitBlockList = BinaryStringToBitBlockList(bitStr);
                var bitBlockListDec = BinaryStringListToDecList(bitBlockList);
                var eccWordList = CalculateECCWords(bitStr, eccInfo);
                var eccWordListDec = BinaryStringListToDecList(eccWordList);
                codeWordWithECC.Add(new CodewordBlock(2,
                                      i + 1,
                                      bitStr,
                                      bitBlockList,
                                      eccWordList,
                                      bitBlockListDec,
                                      eccWordListDec)
                                );
            }


            //Interleave code words
            var interleavedWordsSb = new StringBuilder();
            for (var i = 0; i < Math.Max(eccInfo.CodewordsInGroup1, eccInfo.CodewordsInGroup2); i++)
            {
                foreach (var codeBlock in codeWordWithECC)
                    if (codeBlock.CodeWords.Count > i)
                        interleavedWordsSb.Append(codeBlock.CodeWords[i]);
            }


            for (var i = 0; i < eccInfo.ECCPerBlock; i++)
            {
                foreach (var codeBlock in codeWordWithECC)
                    if (codeBlock.ECCWords.Count > i)
                        interleavedWordsSb.Append(codeBlock.ECCWords[i]);
            }
            interleavedWordsSb.Append(new string('0', remainderBits[version - 1]));
            var interleavedData = interleavedWordsSb.ToString();


            //Place interleaved data on module matrix
            var qr = new QRCodeData(version);
            var blockedModules = new List<Rectangle>();
            ModulePlacer.PlaceFinderPatterns(ref qr, ref blockedModules);
            ModulePlacer.ReserveSeperatorAreas(qr.ModuleMatrix.Count, ref blockedModules);
            ModulePlacer.PlaceAlignmentPatterns(ref qr, alignmentPatternTable.Where(x => x.Version == version).Select(x => x.PatternPositions).First(), ref blockedModules);
            ModulePlacer.PlaceTimingPatterns(ref qr, ref blockedModules);
            ModulePlacer.PlaceDarkModule(ref qr, version, ref blockedModules);
            ModulePlacer.ReserveVersionAreas(qr.ModuleMatrix.Count, version, ref blockedModules);
            ModulePlacer.PlaceDataWords(ref qr, interleavedData, ref blockedModules);
            var maskVersion = ModulePlacer.MaskCode(ref qr, version, ref blockedModules, eccLevel);
            var formatStr = GetFormatString(eccLevel, maskVersion);

            ModulePlacer.PlaceFormat(ref qr, formatStr);
            if (version >= 7)
            {
                var versionString = GetVersionString(version);
                ModulePlacer.PlaceVersion(ref qr, versionString);
            }


            ModulePlacer.AddQuietZone(ref qr);
            return qr;
        }

        private static string GetFormatString(ECCLevel level, int maskVersion)
        {
            var generator = "10100110111";
            var fStrMask = "101010000010010";

            var fStr = (level == ECCLevel.L) ? "01" : (level == ECCLevel.M) ? "00" : (level == ECCLevel.Q) ? "11" : "10";
            fStr += DecToBin(maskVersion, 3);
            var fStrEcc = fStr.PadRight(15, '0').TrimStart('0');
            while (fStrEcc.Length > 10)
            {
                var sb = new StringBuilder();
                generator = generator.PadRight(fStrEcc.Length, '0');
                for (var i = 0; i < fStrEcc.Length; i++)
                    sb.Append((Convert.ToInt32(fStrEcc[i]) ^ Convert.ToInt32(generator[i])).ToString());
                fStrEcc = sb.ToString().TrimStart('0');
            }
            fStrEcc = fStrEcc.PadLeft(10, '0');
            fStr += fStrEcc;

            var sbMask = new StringBuilder();
            for (var i = 0; i < fStr.Length; i++)
                sbMask.Append((Convert.ToInt32(fStr[i]) ^ Convert.ToInt32(fStrMask[i])).ToString());
            return sbMask.ToString();
        }

        private static string GetVersionString(int version)
        {
            var generator = "1111100100101";

            var vStr = DecToBin(version, 6);
            var vStrEcc = vStr.PadRight(18, '0').TrimStart('0');
            while (vStrEcc.Length > 12)
            {
                var sb = new StringBuilder();
                generator = generator.PadRight(vStrEcc.Length, '0');
                for (var i = 0; i < vStrEcc.Length; i++)
                    sb.Append((Convert.ToInt32(vStrEcc[i]) ^ Convert.ToInt32(generator[i])).ToString());
                vStrEcc = sb.ToString().TrimStart('0');
            }
            vStrEcc = vStrEcc.PadLeft(12, '0');
            vStr += vStrEcc;

            return vStr;
        }

        private static class ModulePlacer
        {
            public static void AddQuietZone(ref QRCodeData qrCode)
            {
                var quietLine = new bool[qrCode.ModuleMatrix.Count + 8];
                for (var i = 0; i < quietLine.Length; i++)
                    quietLine[i] = false;
                for (var i = 0; i < 4; i++)
                    qrCode.ModuleMatrix.Insert(0, new BitArray(quietLine));
                for (var i = 0; i < 4; i++)
                    qrCode.ModuleMatrix.Add(new BitArray(quietLine));
                for (var i = 4; i < qrCode.ModuleMatrix.Count - 4; i++)
                {
                    bool[] quietPart = { false, false, false, false };
                    var tmpLine = new List<bool>(quietPart);
                    tmpLine.AddRange(qrCode.ModuleMatrix[i].Cast<bool>());
                    tmpLine.AddRange(quietPart);
                    qrCode.ModuleMatrix[i] = new BitArray(tmpLine.ToArray());
                }
            }

            private static string ReverseString(string inp)
            {
                string newStr = string.Empty;
                if (inp.Length > 0)
                {
                    for (int i = inp.Length - 1; i >= 0; i--)
                        newStr += inp[i];
                }
                return newStr;
            }

            public static void PlaceVersion(ref QRCodeData qrCode, string versionStr)
            {
                var size = qrCode.ModuleMatrix.Count;

                var vStr = ReverseString(versionStr);

                for (var x = 0; x < 6; x++)
                {
                    for (var y = 0; y < 3; y++)
                    {
                        qrCode.ModuleMatrix[y + size - 11][x] = vStr[x * 3 + y] == '1';
                        qrCode.ModuleMatrix[x][y + size - 11] = vStr[x * 3 + y] == '1';
                    }
                }
            }

            public static void PlaceFormat(ref QRCodeData qrCode, string formatStr)
            {
                var size = qrCode.ModuleMatrix.Count;
                var fStr = ReverseString(formatStr);
                var modules = new[,] {
                    { 8, 0, size - 1, 8 },
                    { 8, 1, size - 2, 8 },
                    { 8, 2, size - 3, 8 },
                    { 8, 3, size - 4, 8 },
                    { 8, 4, size - 5, 8 },
                    { 8, 5, size - 6, 8 },
                    { 8, 7, size - 7, 8 },
                    { 8, 8, size - 8, 8 },
                    { 7, 8, 8, size - 7 },
                    { 5, 8, 8, size - 6 },
                    { 4, 8, 8, size - 5 },
                    { 3, 8, 8, size - 4 },
                    { 2, 8, 8, size - 3 },
                    { 1, 8, 8, size - 2 },
                    { 0, 8, 8, size - 1 } };
                for (var i = 0; i < 15; i++)
                {
                    var p1 = new Point(modules[i, 0], modules[i, 1]);
                    var p2 = new Point(modules[i, 2], modules[i, 3]);
                    qrCode.ModuleMatrix[p1.Y][p1.X] = fStr[i] == '1';
                    qrCode.ModuleMatrix[p2.Y][p2.X] = fStr[i] == '1';
                }
            }


            public static int MaskCode(ref QRCodeData qrCode, int version, ref List<Rectangle> blockedModules, ECCLevel eccLevel)
            {
                int? selectedPattern = null;
                var patternScore = 0;

                var size = qrCode.ModuleMatrix.Count;

                var methods = new Dictionary<int, Func<int, int, bool>>(8) {
                    { 1, MaskPattern.Pattern1 }, {2, MaskPattern.Pattern2 }, {3, MaskPattern.Pattern3 }, {4, MaskPattern.Pattern4 },
                    {5,  MaskPattern.Pattern5 }, {6, MaskPattern.Pattern6 }, {7, MaskPattern.Pattern7 }, {8, MaskPattern.Pattern8 }
                };

                foreach (var pattern in methods)
                {
                    var qrTemp = new QRCodeData(version);
                    for (var y = 0; y < size; y++)
                    {
                        for (var x = 0; x < size; x++)
                        {
                            qrTemp.ModuleMatrix[y][x] = qrCode.ModuleMatrix[y][x];
                        }

                    }

                    var formatStr = GetFormatString(eccLevel, pattern.Key - 1);
                    ModulePlacer.PlaceFormat(ref qrTemp, formatStr);
                    if (version >= 7)
                    {
                        var versionString = GetVersionString(version);
                        ModulePlacer.PlaceVersion(ref qrTemp, versionString);
                    }

                    for (var x = 0; x < size; x++)
                    {
                        for (var y = 0; y < x; y++)
                        {
                            if (!IsBlocked(new Rectangle(x, y, 1, 1), blockedModules))
                            {
                                qrTemp.ModuleMatrix[y][x] ^= pattern.Value(x, y);
                                qrTemp.ModuleMatrix[x][y] ^= pattern.Value(y, x);
                            }
                        }

                        if (!IsBlocked(new Rectangle(x, x, 1, 1), blockedModules))
                        {
                            qrTemp.ModuleMatrix[x][x] ^= pattern.Value(x, x);
                        }
                    }

                    var score = MaskPattern.Score(ref qrTemp);
                    if (!selectedPattern.HasValue || patternScore > score)
                    {
                        selectedPattern = pattern.Key;
                        patternScore = score;
                    }
                }

                for (var x = 0; x < size; x++)
                {
                    for (var y = 0; y < x; y++)
                    {
                        if (!IsBlocked(new Rectangle(x, y, 1, 1), blockedModules))
                        {
                            qrCode.ModuleMatrix[y][x] ^= methods[selectedPattern.Value](x, y);
                            qrCode.ModuleMatrix[x][y] ^= methods[selectedPattern.Value](y, x);
                        }
                    }

                    if (!IsBlocked(new Rectangle(x, x, 1, 1), blockedModules))
                    {
                        qrCode.ModuleMatrix[x][x] ^= methods[selectedPattern.Value](x, x);
                    }
                }
                return selectedPattern.Value - 1;
            }


            public static void PlaceDataWords(ref QRCodeData qrCode, string data, ref List<Rectangle> blockedModules)
            {
                var size = qrCode.ModuleMatrix.Count;
                var up = true;
                var datawords = new Queue<bool>();
                for (int i = 0; i< data.Length; i++)
                {
                    datawords.Enqueue(data[i] != '0');
                }
                for (var x = size - 1; x >= 0; x = x - 2)
                {
                    if (x == 6)
                        x = 5;
                    for (var yMod = 1; yMod <= size; yMod++)
                    {
                        int y;
                        if (up)
                        {
                            y = size - yMod;
                            if (datawords.Count > 0 && !IsBlocked(new Rectangle(x, y, 1, 1), blockedModules))
                                qrCode.ModuleMatrix[y][x] = datawords.Dequeue();
                            if (datawords.Count > 0 && x > 0 && !IsBlocked(new Rectangle(x - 1, y, 1, 1), blockedModules))
                                qrCode.ModuleMatrix[y][x - 1] = datawords.Dequeue();
                        }
                        else
                        {
                            y = yMod - 1;
                            if (datawords.Count > 0 && !IsBlocked(new Rectangle(x, y, 1, 1), blockedModules))
                                qrCode.ModuleMatrix[y][x] = datawords.Dequeue();
                            if (datawords.Count > 0 && x > 0 && !IsBlocked(new Rectangle(x - 1, y, 1, 1), blockedModules))
                                qrCode.ModuleMatrix[y][x - 1] = datawords.Dequeue();
                        }
                    }
                    up = !up;
                }
            }

            public static void ReserveSeperatorAreas(int size, ref List<Rectangle> blockedModules)
            {
                blockedModules.AddRange(new[]{
                    new Rectangle(7, 0, 1, 8),
                    new Rectangle(0, 7, 7, 1),
                    new Rectangle(0, size-8, 8, 1),
                    new Rectangle(7, size-7, 1, 7),
                    new Rectangle(size-8, 0, 1, 8),
                    new Rectangle(size-7, 7, 7, 1)
                });
            }

            public static void ReserveVersionAreas(int size, int version, ref List<Rectangle> blockedModules)
            {
                blockedModules.AddRange(new[]{
                    new Rectangle(8, 0, 1, 6),
                    new Rectangle(8, 7, 1, 1),
                    new Rectangle(0, 8, 6, 1),
                    new Rectangle(7, 8, 2, 1),
                    new Rectangle(size-8, 8, 8, 1),
                    new Rectangle(8, size-7, 1, 7)
                });

                if (version >= 7)
                {
                    blockedModules.AddRange(new[]{
                    new Rectangle(size-11, 0, 3, 6),
                    new Rectangle(0, size-11, 6, 3)
                });
                }
            }
            public static void PlaceDarkModule(ref QRCodeData qrCode, int version, ref List<Rectangle> blockedModules)
            {
                qrCode.ModuleMatrix[4 * version + 9][8] = true;
                blockedModules.Add(new Rectangle(8, 4 * version + 9, 1, 1));
            }

            public static void PlaceFinderPatterns(ref QRCodeData qrCode, ref List<Rectangle> blockedModules)
            {
                var size = qrCode.ModuleMatrix.Count;
                int[] locations = { 0, 0, size - 7, 0, 0, size - 7 };

                for (var i = 0; i < 6; i = i + 2)
                {
                    for (var x = 0; x < 7; x++)
                    {
                        for (var y = 0; y < 7; y++)
                        {
                            if (!(((x == 1 || x == 5) && y > 0 && y < 6) || (x > 0 && x < 6 && (y == 1 || y == 5))))
                            {
                                qrCode.ModuleMatrix[y + locations[i + 1]][x + locations[i]] = true;
                            }
                        }
                    }
                    blockedModules.Add(new Rectangle(locations[i], locations[i + 1], 7, 7));
                }
            }

            public static void PlaceAlignmentPatterns(ref QRCodeData qrCode, List<Point> alignmentPatternLocations, ref List<Rectangle> blockedModules)
            {
                foreach (var loc in alignmentPatternLocations)
                {
                    var alignmentPatternRect = new Rectangle(loc.X, loc.Y, 5, 5);
                    var blocked = false;
                    foreach (var blockedRect in blockedModules)
                    {
                        if (Intersects(alignmentPatternRect, blockedRect))
                        {
                            blocked = true;
                            break;
                        }
                    }
                    if (blocked)
                        continue;

                    for (var x = 0; x < 5; x++)
                    {
                        for (var y = 0; y < 5; y++)
                        {
                            if (y == 0 || y == 4 || x == 0 || x == 4 || (x == 2 && y == 2))
                            {
                                qrCode.ModuleMatrix[loc.Y + y][loc.X + x] = true;
                            }
                        }
                    }
                    blockedModules.Add(new Rectangle(loc.X, loc.Y, 5, 5));
                }
            }

            public static void PlaceTimingPatterns(ref QRCodeData qrCode, ref List<Rectangle> blockedModules)
            {
                var size = qrCode.ModuleMatrix.Count;
                for (var i = 8; i < size - 8; i++)
                {
                    if (i % 2 == 0)
                    {
                        qrCode.ModuleMatrix[6][i] = true;
                        qrCode.ModuleMatrix[i][6] = true;
                    }
                }
                blockedModules.AddRange(new[]{
                    new Rectangle(6, 8, 1, size-16),
                    new Rectangle(8, 6, size-16, 1)
                });
            }

            private static bool Intersects(Rectangle r1, Rectangle r2)
            {
                return r2.X < r1.X + r1.Width && r1.X < r2.X + r2.Width && r2.Y < r1.Y + r1.Height && r1.Y < r2.Y + r2.Height;
            }

            private static bool IsBlocked(Rectangle r1, List<Rectangle> blockedModules)
            {
                foreach (var blockedMod in blockedModules)
                {
                    if (Intersects(blockedMod, r1))
                        return true;
                }
                return false;
            }

            private static class MaskPattern
            {
                public static bool Pattern1(int x, int y)
                {
                    return (x + y) % 2 == 0;
                }

                public static bool Pattern2(int x, int y)
                {
                    return y % 2 == 0;
                }

                public static bool Pattern3(int x, int y)
                {
                    return x % 3 == 0;
                }

                public static bool Pattern4(int x, int y)
                {
                    return (x + y) % 3 == 0;
                }

                public static bool Pattern5(int x, int y)
                {
                    return ((int)(Math.Floor(y / 2d) + Math.Floor(x / 3d)) % 2) == 0;
                }

                public static bool Pattern6(int x, int y)
                {
                    return ((x * y) % 2) + ((x * y) % 3) == 0;
                }

                public static bool Pattern7(int x, int y)
                {
                    return (((x * y) % 2) + ((x * y) % 3)) % 2 == 0;
                }

                public static bool Pattern8(int x, int y)
                {
                    return (((x + y) % 2) + ((x * y) % 3)) % 2 == 0;
                }

                public static int Score(ref QRCodeData qrCode)
                {
                    int score1 = 0,
                        score2 = 0,
                        score3 = 0,
                        score4 = 0;
                    var size = qrCode.ModuleMatrix.Count;

                    //Penalty 1
                    for (var y = 0; y < size; y++)
                    {
                        var modInRow = 0;
                        var modInColumn = 0;
                        var lastValRow = qrCode.ModuleMatrix[y][0];
                        var lastValColumn = qrCode.ModuleMatrix[0][y];
                        for (var x = 0; x < size; x++)
                        {
                            if (qrCode.ModuleMatrix[y][x] == lastValRow)
                                modInRow++;
                            else
                                modInRow = 1;
                            if (modInRow == 5)
                                score1 += 3;
                            else if (modInRow > 5)
                                score1++;
                            lastValRow = qrCode.ModuleMatrix[y][x];


                            if (qrCode.ModuleMatrix[x][y] == lastValColumn)
                                modInColumn++;
                            else
                                modInColumn = 1;
                            if (modInColumn == 5)
                                score1 += 3;
                            else if (modInColumn > 5)
                                score1++;
                            lastValColumn = qrCode.ModuleMatrix[x][y];
                        }
                    }


                    //Penalty 2
                    for (var y = 0; y < size - 1; y++)
                    {
                        for (var x = 0; x < size - 1; x++)
                        {
                            if (qrCode.ModuleMatrix[y][x] == qrCode.ModuleMatrix[y][x + 1] &&
                                qrCode.ModuleMatrix[y][x] == qrCode.ModuleMatrix[y + 1][x] &&
                                qrCode.ModuleMatrix[y][x] == qrCode.ModuleMatrix[y + 1][x + 1])
                                score2 += 3;
                        }
                    }

                    //Penalty 3
                    for (var y = 0; y < size; y++)
                    {
                        for (var x = 0; x < size - 10; x++)
                        {
                            if ((qrCode.ModuleMatrix[y][x] &&
                                !qrCode.ModuleMatrix[y][x + 1] &&
                                qrCode.ModuleMatrix[y][x + 2] &&
                                qrCode.ModuleMatrix[y][x + 3] &&
                                qrCode.ModuleMatrix[y][x + 4] &&
                                !qrCode.ModuleMatrix[y][x + 5] &&
                                qrCode.ModuleMatrix[y][x + 6] &&
                                !qrCode.ModuleMatrix[y][x + 7] &&
                                !qrCode.ModuleMatrix[y][x + 8] &&
                                !qrCode.ModuleMatrix[y][x + 9] &&
                                !qrCode.ModuleMatrix[y][x + 10]) ||
                                (!qrCode.ModuleMatrix[y][x] &&
                                !qrCode.ModuleMatrix[y][x + 1] &&
                                !qrCode.ModuleMatrix[y][x + 2] &&
                                !qrCode.ModuleMatrix[y][x + 3] &&
                                qrCode.ModuleMatrix[y][x + 4] &&
                                !qrCode.ModuleMatrix[y][x + 5] &&
                                qrCode.ModuleMatrix[y][x + 6] &&
                                qrCode.ModuleMatrix[y][x + 7] &&
                                qrCode.ModuleMatrix[y][x + 8] &&
                                !qrCode.ModuleMatrix[y][x + 9] &&
                                qrCode.ModuleMatrix[y][x + 10]))
                            {
                                score3 += 40;
                            }

                            if ((qrCode.ModuleMatrix[x][y] &&
                                !qrCode.ModuleMatrix[x + 1][y] &&
                                qrCode.ModuleMatrix[x + 2][y] &&
                                qrCode.ModuleMatrix[x + 3][y] &&
                                qrCode.ModuleMatrix[x + 4][y] &&
                                !qrCode.ModuleMatrix[x + 5][y] &&
                                qrCode.ModuleMatrix[x + 6][y] &&
                                !qrCode.ModuleMatrix[x + 7][y] &&
                                !qrCode.ModuleMatrix[x + 8][y] &&
                                !qrCode.ModuleMatrix[x + 9][y] &&
                                !qrCode.ModuleMatrix[x + 10][y]) ||
                                (!qrCode.ModuleMatrix[x][y] &&
                                !qrCode.ModuleMatrix[x + 1][y] &&
                                !qrCode.ModuleMatrix[x + 2][y] &&
                                !qrCode.ModuleMatrix[x + 3][y] &&
                                qrCode.ModuleMatrix[x + 4][y] &&
                                !qrCode.ModuleMatrix[x + 5][y] &&
                                qrCode.ModuleMatrix[x + 6][y] &&
                                qrCode.ModuleMatrix[x + 7][y] &&
                                qrCode.ModuleMatrix[x + 8][y] &&
                                !qrCode.ModuleMatrix[x + 9][y] &&
                                qrCode.ModuleMatrix[x + 10][y]))
                            {
                                score3 += 40;
                            }
                        }
                    }

                    //Penalty 4
                    double blackModules = 0;
                    foreach (var row in qrCode.ModuleMatrix)
                        foreach (bool bit in row)
                            if (bit)
                                blackModules++;

                    var percent = (blackModules / (qrCode.ModuleMatrix.Count * qrCode.ModuleMatrix.Count)) * 100;
                    var prevMultipleOf5 = Math.Abs((int) Math.Floor(percent/5)*5 - 50)/5;
                    var nextMultipleOf5 = Math.Abs((int)Math.Floor(percent / 5) * 5 -45)/5;
                    score4 = Math.Min(prevMultipleOf5, nextMultipleOf5)*10;

                    return score1 + score2 + score3 + score4;
                }
            }

        }

        private static List<string> CalculateECCWords(string bitString, ECCInfo eccInfo)
        {
            var eccWords = eccInfo.ECCPerBlock;
            var messagePolynom = CalculateMessagePolynom(bitString);
            var generatorPolynom = CalculateGeneratorPolynom(eccWords);

            for (var i = 0; i < messagePolynom.PolyItems.Count; i++)
                messagePolynom.PolyItems[i] = new PolynomItem(messagePolynom.PolyItems[i].Coefficient,
                    messagePolynom.PolyItems[i].Exponent + eccWords);

            for (var i = 0; i < generatorPolynom.PolyItems.Count; i++)
                generatorPolynom.PolyItems[i] = new PolynomItem(generatorPolynom.PolyItems[i].Coefficient,
                    generatorPolynom.PolyItems[i].Exponent + (messagePolynom.PolyItems.Count-1));

            var leadTermSource = messagePolynom;
            for (var i = 0; (leadTermSource.PolyItems.Count > 0 && leadTermSource.PolyItems[leadTermSource.PolyItems.Count - 1].Exponent > 0); i++)
            {
                if (leadTermSource.PolyItems[0].Coefficient == 0)
                {
                    leadTermSource.PolyItems.RemoveAt(0);
                    leadTermSource.PolyItems.Add(new PolynomItem(0, leadTermSource.PolyItems[leadTermSource.PolyItems.Count - 1].Exponent - 1));
                }
                else
                {
                    var resPoly = MultiplyGeneratorPolynomByLeadterm(generatorPolynom, ConvertToAlphaNotation(leadTermSource).PolyItems[0], i);
                    resPoly = ConvertToDecNotation(resPoly);
                    resPoly = XORPolynoms(leadTermSource, resPoly);
                    leadTermSource = resPoly;
                }
            }
            return leadTermSource.PolyItems.Select(x => DecToBin(x.Coefficient, 8)).ToList();
        }

        private static Polynom ConvertToAlphaNotation(Polynom poly)
        {
            var newPoly = new Polynom();
            for (var i = 0; i < poly.PolyItems.Count; i++)
                newPoly.PolyItems.Add(
                    new PolynomItem(
                        (poly.PolyItems[i].Coefficient != 0
                            ? GetAlphaExpFromIntVal(poly.PolyItems[i].Coefficient)
                            : 0), poly.PolyItems[i].Exponent));
            return newPoly;
        }

        private static Polynom ConvertToDecNotation(Polynom poly)
        {
            var newPoly = new Polynom();
            for (var i = 0; i < poly.PolyItems.Count; i++)
                newPoly.PolyItems.Add(new PolynomItem(GetIntValFromAlphaExp(poly.PolyItems[i].Coefficient), poly.PolyItems[i].Exponent));
            return newPoly;
        }

        private static int GetVersion(int length, EncodingMode encMode, ECCLevel eccLevel)
        {

            var fittingVersions = capacityTable.Where(
                x => x.Details.Any(
                    y => (y.ErrorCorrectionLevel == eccLevel
                          && y.CapacityDict[encMode] >= Convert.ToInt32(length)
                          )
                    )
              ).Select(x => new
              {
                  version = x.Version,
                  capacity = x.Details.Single(y => y.ErrorCorrectionLevel == eccLevel)
                                            .CapacityDict[encMode]
              });

            if (fittingVersions.Any())
                return fittingVersions.Min(x => x.version);

            var maxSizeByte = capacityTable.Where(
                x => x.Details.Any(
                    y => (y.ErrorCorrectionLevel == eccLevel))
                ).Max(x => x.Details.Single(y => y.ErrorCorrectionLevel == eccLevel).CapacityDict[encMode]);
            throw new QRCoder.Exceptions.DataTooLongException(eccLevel.ToString(), encMode.ToString(), maxSizeByte);
        }

        private static EncodingMode GetEncodingFromPlaintext(string plainText, bool forceUtf8)
        {
            if (forceUtf8) return EncodingMode.Byte;
            EncodingMode result = EncodingMode.Numeric; // assume numeric
            foreach (char c in plainText)
            {
                if (IsInRange(c, '0', '9')) continue;   // numeric - char.IsDigit() for Latin1
                result = EncodingMode.Alphanumeric;     // not numeric, assume alphanumeric
                if (IsInRange(c, 'A', 'Z') || alphanumEncTable.Contains(c)) continue; // alphanumeric
                return EncodingMode.Byte;               // not numeric or alphanumeric, assume byte
            }
            return result;                              // either numeric or alphanumeric
        }

        private static bool IsInRange(char c, char min, char max)
        {
            return (uint)(c - min) <= (uint)(max - min);
        }

        private static Polynom CalculateMessagePolynom(string bitString)
        {
            var messagePol = new Polynom();
            for (var i = bitString.Length / 8 - 1; i >= 0; i--)
            {
                messagePol.PolyItems.Add(new PolynomItem(BinToDec(bitString.Substring(0, 8)), i));
                bitString = bitString.Remove(0, 8);
            }
            return messagePol;
        }


        private static Polynom CalculateGeneratorPolynom(int numEccWords)
        {
            var generatorPolynom = new Polynom();
            generatorPolynom.PolyItems.AddRange(new[]{
                new PolynomItem(0,1),
                new PolynomItem(0,0)
            });
            for (var i = 1; i <= numEccWords - 1; i++)
            {
                var multiplierPolynom = new Polynom();
                multiplierPolynom.PolyItems.AddRange(new[]{
                   new PolynomItem(0,1),
                new PolynomItem(i,0)
                });

                generatorPolynom = MultiplyAlphaPolynoms(generatorPolynom, multiplierPolynom);
            }

            return generatorPolynom;
        }

        private static List<string> BinaryStringToBitBlockList(string bitString)
        {
            const int blockSize = 8;
            var numberOfBlocks = (int)Math.Ceiling(bitString.Length / (double)blockSize);
            var blocklist = new List<string>(numberOfBlocks);

            for (int i = 0; i < bitString.Length; i += blockSize)
            {
                blocklist.Add(bitString.Substring(i, blockSize));
            }

            return blocklist;
        }

        private static List<int> BinaryStringListToDecList(List<string> binaryStringList)
        {
            return binaryStringList.Select(binaryString => BinToDec(binaryString)).ToList();
        }

        private static int BinToDec(string binStr)
        {
            return Convert.ToInt32(binStr, 2);
        }

        private static string DecToBin(int decNum)
        {
            return Convert.ToString(decNum, 2);
        }

        private static string DecToBin(int decNum, int padLeftUpTo)
        {
            var binStr = DecToBin(decNum);
            return binStr.PadLeft(padLeftUpTo, '0');
        }

        private static int GetCountIndicatorLength(int version, EncodingMode encMode)
        {
            if (version < 10)
            {
                if (encMode == EncodingMode.Numeric)
                    return 10;
                else if (encMode == EncodingMode.Alphanumeric)
                    return 9;
                else
                    return 8;
            }
            else if (version < 27)
            {
                if (encMode == EncodingMode.Numeric)
                    return 12;
                else if (encMode == EncodingMode.Alphanumeric)
                    return 11;
                else if (encMode == EncodingMode.Byte)
                    return 16;
                else
                    return 10;
            }
            else
            {
                if (encMode == EncodingMode.Numeric)
                    return 14;
                else if (encMode == EncodingMode.Alphanumeric)
                    return 13;
                else if (encMode == EncodingMode.Byte)
                    return 16;
                else
                    return 12;
            }
        }

        private static int GetDataLength(EncodingMode encoding, string plainText, string codedText, bool forceUtf8)
        {
            return forceUtf8 || IsUtf8(encoding, plainText, forceUtf8) ? (codedText.Length / 8) : plainText.Length;
        }

        private static bool IsUtf8(EncodingMode encoding, string plainText, bool forceUtf8)
        {
            return (encoding == EncodingMode.Byte && (!IsValidISO(plainText) || forceUtf8));
        }

        private static bool IsValidISO(string input)
        {
            var bytes = Encoding.GetEncoding("ISO-8859-1").GetBytes(input);
            //var result = Encoding.GetEncoding("ISO-8859-1").GetString(bytes);
            var result = Encoding.GetEncoding("ISO-8859-1").GetString(bytes,0,bytes.Length);
            return String.Equals(input, result);
        }

        private static string PlainTextToBinary(string plainText, EncodingMode encMode, EciMode eciMode, bool utf8BOM, bool forceUtf8)
        {
            switch(encMode)
            {
                case EncodingMode.Alphanumeric:
                    return PlainTextToBinaryAlphanumeric(plainText);
                case EncodingMode.Numeric:
                    return PlainTextToBinaryNumeric(plainText);
                case EncodingMode.Byte:
                    return PlainTextToBinaryByte(plainText, eciMode, utf8BOM, forceUtf8);
                case EncodingMode.Kanji:
                    return string.Empty;
                case EncodingMode.ECI:
                default:
                    return string.Empty;
            }
        }

        private static string PlainTextToBinaryNumeric(string plainText)
        {
            var codeText = string.Empty;
            while (plainText.Length >= 3)
            {
                var dec = Convert.ToInt32(plainText.Substring(0, 3));
                codeText += DecToBin(dec, 10);
                plainText = plainText.Substring(3);

            }
            if (plainText.Length == 2)
            {
                var dec = Convert.ToInt32(plainText);
                codeText += DecToBin(dec, 7);
            }
            else if (plainText.Length == 1)
            {
                var dec = Convert.ToInt32(plainText);
                codeText += DecToBin(dec, 4);
            }
            return codeText;
        }

        private static string PlainTextToBinaryAlphanumeric(string plainText)
        {
            var codeText = string.Empty;
            while (plainText.Length >= 2)
            {
                var token = plainText.Substring(0, 2);
                var dec = alphanumEncDict[token[0]] * 45 + alphanumEncDict[token[1]];
                codeText += DecToBin(dec, 11);
                plainText = plainText.Substring(2);

            }
            if (plainText.Length > 0)
            {
                codeText += DecToBin(alphanumEncDict[plainText[0]], 6);
            }
            return codeText;
        }

        private string PlainTextToBinaryECI(string plainText)
        {
            var codeText = string.Empty;
            byte[] _bytes = Encoding.GetEncoding("ascii").GetBytes(plainText);
            foreach(byte _byte in _bytes)
            {
                codeText += DecToBin(_byte, 8);
            }
            return codeText;
        }

        private static string ConvertToIso8859(string value, string Iso = "ISO-8859-2")
        {
            Encoding iso = Encoding.GetEncoding(Iso);
            Encoding utf8 = Encoding.UTF8;
            byte[] utfBytes = utf8.GetBytes(value);
            byte[] isoBytes = Encoding.Convert(utf8, iso, utfBytes);
#if NETFRAMEWORK || NETSTANDARD2_0 || NET5_0
            return iso.GetString(isoBytes);
#else
            return iso.GetString(isoBytes, 0, isoBytes.Length);
#endif
        }

        private static string PlainTextToBinaryByte(string plainText, EciMode eciMode, bool utf8BOM, bool forceUtf8)
        {
            byte[] codeBytes;
            var codeText = string.Empty;

            if (IsValidISO(plainText) && !forceUtf8)
                codeBytes = Encoding.GetEncoding("ISO-8859-1").GetBytes(plainText);
            else
            {
                switch(eciMode)
                {
                    case EciMode.Iso8859_1:
                        codeBytes = Encoding.GetEncoding("ISO-8859-1").GetBytes(ConvertToIso8859(plainText, "ISO-8859-1"));
                        break;
                    case EciMode.Iso8859_2:
                        codeBytes = Encoding.GetEncoding("ISO-8859-2").GetBytes(ConvertToIso8859(plainText, "ISO-8859-2"));
                        break;
                    case EciMode.Default:
                    case EciMode.Utf8:
                    default:
                        codeBytes = utf8BOM ? Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(plainText)).ToArray() : Encoding.UTF8.GetBytes(plainText);
                        break;
                }
            }

            foreach (var b in codeBytes)
                codeText += DecToBin(b, 8);

            return codeText;
        }


        private static Polynom XORPolynoms(Polynom messagePolynom, Polynom resPolynom)
        {
            var resultPolynom = new Polynom();
            Polynom longPoly, shortPoly;
            if (messagePolynom.PolyItems.Count >= resPolynom.PolyItems.Count)
            {
                longPoly = messagePolynom;
                shortPoly = resPolynom;
            }
            else
            {
                longPoly = resPolynom;
                shortPoly = messagePolynom;
            }

            for (var i = 0; i < longPoly.PolyItems.Count; i++)
            {
                var polItemRes = new PolynomItem
                (

                        longPoly.PolyItems[i].Coefficient ^
                        (shortPoly.PolyItems.Count > i ? shortPoly.PolyItems[i].Coefficient : 0),
                    messagePolynom.PolyItems[0].Exponent - i
                );
                resultPolynom.PolyItems.Add(polItemRes);
            }
            resultPolynom.PolyItems.RemoveAt(0);
            return resultPolynom;
        }


        private static Polynom MultiplyGeneratorPolynomByLeadterm(Polynom genPolynom, PolynomItem leadTerm, int lowerExponentBy)
        {
            var resultPolynom = new Polynom();
            foreach (var polItemBase in genPolynom.PolyItems)
            {
                var polItemRes = new PolynomItem(

                    (polItemBase.Coefficient + leadTerm.Coefficient) % 255,
                    polItemBase.Exponent - lowerExponentBy
                );
                resultPolynom.PolyItems.Add(polItemRes);
            }
            return resultPolynom;
        }


        private static Polynom MultiplyAlphaPolynoms(Polynom polynomBase, Polynom polynomMultiplier)
        {
            var resultPolynom = new Polynom();
            foreach (var polItemBase in polynomMultiplier.PolyItems)
            {
                foreach (var polItemMulti in polynomBase.PolyItems)
                {
                    var polItemRes = new PolynomItem
                    (
                        ShrinkAlphaExp(polItemBase.Coefficient + polItemMulti.Coefficient),
                        (polItemBase.Exponent + polItemMulti.Exponent)
                    );
                    resultPolynom.PolyItems.Add(polItemRes);
                }
            }
            var exponentsToGlue = resultPolynom.PolyItems.GroupBy(x => x.Exponent).Where(x => x.Count() > 1).Select(x => x.First().Exponent);
            var toGlue = exponentsToGlue as IList<int> ?? exponentsToGlue.ToList();
            var gluedPolynoms = new List<PolynomItem>(toGlue.Count);
            foreach (var exponent in toGlue)
            {
                var coefficient = resultPolynom.PolyItems.Where(x => x.Exponent == exponent).Aggregate(0, (current, polynomOld)
                    => current ^ GetIntValFromAlphaExp(polynomOld.Coefficient));
                var polynomFixed = new PolynomItem(GetAlphaExpFromIntVal(coefficient), exponent);
                gluedPolynoms.Add(polynomFixed);
            }
            resultPolynom.PolyItems.RemoveAll(x => toGlue.Contains(x.Exponent));
            resultPolynom.PolyItems.AddRange(gluedPolynoms);
            resultPolynom.PolyItems.Sort((x, y) => -x.Exponent.CompareTo(y.Exponent));
            return resultPolynom;
        }

        private static int GetIntValFromAlphaExp(int exp)
        {
            return galoisField.Find(alog => alog.ExponentAlpha == exp).IntegerValue;
        }

        private static int GetAlphaExpFromIntVal(int intVal)
        {
            return galoisField.Find(alog => alog.IntegerValue == intVal).ExponentAlpha;
        }

        private static int ShrinkAlphaExp(int alphaExp)
        {
            // ReSharper disable once PossibleLossOfFraction
            return (int)((alphaExp % 256) + Math.Floor((double)(alphaExp / 256)));
        }

        private static Dictionary<char, int> CreateAlphanumEncDict()
        {
            var localAlphanumEncDict = new Dictionary<char, int>(45);
            //Add numbers
            for (int i = 0; i < 10; i++)
                localAlphanumEncDict.Add((char)('0' + i), i);
            //Add chars
            for (char c = 'A'; c <= 'Z'; c++)
                localAlphanumEncDict.Add(c, localAlphanumEncDict.Count());
            //Add special chars
            for (int i = 0; i < alphanumEncTable.Length; i++)
                localAlphanumEncDict.Add(alphanumEncTable[i], localAlphanumEncDict.Count());
            return localAlphanumEncDict;
        }

        private static List<AlignmentPattern> CreateAlignmentPatternTable()
        {
            var localAlignmentPatternTable = new List<AlignmentPattern>(40);

            for (var i = 0; i < (7 * 40); i = i + 7)
            {
                var points = new List<Point>();
                for (var x = 0; x < 7; x++)
                {
                    if (alignmentPatternBaseValues[i + x] != 0)
                    {
                        for (var y = 0; y < 7; y++)
                        {
                            if (alignmentPatternBaseValues[i + y] != 0)
                            {
                                var p = new Point(alignmentPatternBaseValues[i + x] - 2, alignmentPatternBaseValues[i + y] - 2);
                                if (!points.Contains(p))
                                    points.Add(p);
                            }
                        }
                    }
                }

                localAlignmentPatternTable.Add(new AlignmentPattern()
                {
                    Version = (i + 7) / 7,
                    PatternPositions = points
                }
                );
            }
            return localAlignmentPatternTable;
        }


        private static List<ECCInfo> CreateCapacityECCTable()
        {
            var localCapacityECCTable = new List<ECCInfo>(160);
            for (var i = 0; i < (4 * 6 * 40); i = i + (4 * 6))
            {
                localCapacityECCTable.AddRange(
                new[]
                {
                    new ECCInfo(
                        (i+24) / 24,
                        ECCLevel.L,
                        capacityECCBaseValues[i],
                        capacityECCBaseValues[i+1],
                        capacityECCBaseValues[i+2],
                        capacityECCBaseValues[i+3],
                        capacityECCBaseValues[i+4],
                        capacityECCBaseValues[i+5]),
                    new ECCInfo
                    (
                        version: (i + 24) / 24,
                        errorCorrectionLevel: ECCLevel.M,
                        totalDataCodewords: capacityECCBaseValues[i+6],
                        eccPerBlock: capacityECCBaseValues[i+7],
                        blocksInGroup1: capacityECCBaseValues[i+8],
                        codewordsInGroup1: capacityECCBaseValues[i+9],
                        blocksInGroup2: capacityECCBaseValues[i+10],
                        codewordsInGroup2: capacityECCBaseValues[i+11]
                    ),
                    new ECCInfo
                    (
                        version: (i + 24) / 24,
                        errorCorrectionLevel: ECCLevel.Q,
                        totalDataCodewords: capacityECCBaseValues[i+12],
                        eccPerBlock: capacityECCBaseValues[i+13],
                        blocksInGroup1: capacityECCBaseValues[i+14],
                        codewordsInGroup1: capacityECCBaseValues[i+15],
                        blocksInGroup2: capacityECCBaseValues[i+16],
                        codewordsInGroup2: capacityECCBaseValues[i+17]
                    ),
                    new ECCInfo
                    (
                        version: (i + 24) / 24,
                        errorCorrectionLevel: ECCLevel.H,
                        totalDataCodewords: capacityECCBaseValues[i+18],
                        eccPerBlock: capacityECCBaseValues[i+19],
                        blocksInGroup1: capacityECCBaseValues[i+20],
                        codewordsInGroup1: capacityECCBaseValues[i+21],
                        blocksInGroup2: capacityECCBaseValues[i+22],
                        codewordsInGroup2: capacityECCBaseValues[i+23]
                    )
                });
            }
            return localCapacityECCTable;
        }

        private static List<VersionInfo> CreateCapacityTable()
        {
            var localCapacityTable = new List<VersionInfo>(40);
            for (var i = 0; i < (16 * 40); i = i + 16)
            {
                localCapacityTable.Add(new VersionInfo(

                    (i + 16) / 16,
                    new List<VersionInfoDetails>(4)
                    {
                        new VersionInfoDetails(
                             ECCLevel.L,
                             new Dictionary<EncodingMode,int>(){
                                 { EncodingMode.Numeric, capacityBaseValues[i] },
                                 { EncodingMode.Alphanumeric, capacityBaseValues[i+1] },
                                 { EncodingMode.Byte, capacityBaseValues[i+2] },
                                 { EncodingMode.Kanji, capacityBaseValues[i+3] },
                            }
                        ),
                        new VersionInfoDetails(
                             ECCLevel.M,
                             new Dictionary<EncodingMode,int>(){
                                 { EncodingMode.Numeric, capacityBaseValues[i+4] },
                                 { EncodingMode.Alphanumeric, capacityBaseValues[i+5] },
                                 { EncodingMode.Byte, capacityBaseValues[i+6] },
                                 { EncodingMode.Kanji, capacityBaseValues[i+7] },
                             }
                        ),
                        new VersionInfoDetails(
                             ECCLevel.Q,
                             new Dictionary<EncodingMode,int>(){
                                 { EncodingMode.Numeric, capacityBaseValues[i+8] },
                                 { EncodingMode.Alphanumeric, capacityBaseValues[i+9] },
                                 { EncodingMode.Byte, capacityBaseValues[i+10] },
                                 { EncodingMode.Kanji, capacityBaseValues[i+11] },
                             }
                        ),
                        new VersionInfoDetails(
                             ECCLevel.H,
                             new Dictionary<EncodingMode,int>(){
                                 { EncodingMode.Numeric, capacityBaseValues[i+12] },
                                 { EncodingMode.Alphanumeric, capacityBaseValues[i+13] },
                                 { EncodingMode.Byte, capacityBaseValues[i+14] },
                                 { EncodingMode.Kanji, capacityBaseValues[i+15] },
                             }
                        )
                    }
                ));
            }
            return localCapacityTable;
        }

        private static List<Antilog> CreateAntilogTable()
        {
            var localGaloisField = new List<Antilog>(256);

            int gfItem = 1;
            for (var i = 0; i < 256; i++)
            {
                localGaloisField.Add(new Antilog(i, gfItem));
                gfItem *= 2;
                if (gfItem > 255)
                    gfItem ^= 285;
            }
            return localGaloisField;
        }

        /// <summary>
        /// Error correction level. These define the tolerance levels for how much of the code can be lost before the code cannot be recovered.
        /// </summary>
        public enum ECCLevel
        {
            /// <summary>
            /// 7% may be lost before recovery is not possible
            /// </summary>
            L,
            /// <summary>
            /// 15% may be lost before recovery is not possible
            /// </summary>
            M,
            /// <summary>
            /// 25% may be lost before recovery is not possible
            /// </summary>
            Q,
            /// <summary>
            /// 30% may be lost before recovery is not possible
            /// </summary>
            H
        }

        private enum EncodingMode
        {
            Numeric = 1,
            Alphanumeric = 2,
            Byte = 4,
            Kanji = 8,
            ECI = 7
        }

        private struct AlignmentPattern
        {
            public int Version;
            public List<Point> PatternPositions;
        }

        private struct CodewordBlock
        {
            public CodewordBlock(int groupNumber, int blockNumber, string bitString, List<string> codeWords,
                List<string> eccWords, List<int> codeWordsInt, List<int> eccWordsInt)
            {
                this.GroupNumber = groupNumber;
                this.BlockNumber = blockNumber;
                this.BitString = bitString;
                this.CodeWords = codeWords;
                this.ECCWords = eccWords;
                this.CodeWordsInt = codeWordsInt;
                this.ECCWordsInt = eccWordsInt;
            }

            public readonly int GroupNumber;
            public readonly int BlockNumber;
            public readonly string BitString;
            public readonly List<string> CodeWords;
            public readonly List<int> CodeWordsInt;
            public readonly List<string> ECCWords;
            public readonly List<int> ECCWordsInt;
        }

        private struct ECCInfo
        {
            public ECCInfo(int version, ECCLevel errorCorrectionLevel, int totalDataCodewords, int eccPerBlock, int blocksInGroup1,
                int codewordsInGroup1, int blocksInGroup2, int codewordsInGroup2)
            {
                this.Version = version;
                this.ErrorCorrectionLevel = errorCorrectionLevel;
                this.TotalDataCodewords = totalDataCodewords;
                this.ECCPerBlock = eccPerBlock;
                this.BlocksInGroup1 = blocksInGroup1;
                this.CodewordsInGroup1 = codewordsInGroup1;
                this.BlocksInGroup2 = blocksInGroup2;
                this.CodewordsInGroup2 = codewordsInGroup2;
            }
            public readonly int Version;
            public readonly ECCLevel ErrorCorrectionLevel;
            public readonly int TotalDataCodewords;
            public readonly int ECCPerBlock;
            public readonly int BlocksInGroup1;
            public readonly int CodewordsInGroup1;
            public readonly int BlocksInGroup2;
            public readonly int CodewordsInGroup2;
        }

        private struct VersionInfo
        {
            public VersionInfo(int version, List<VersionInfoDetails> versionInfoDetails)
            {
                this.Version = version;
                this.Details = versionInfoDetails;
            }
            public readonly int Version;
            public readonly List<VersionInfoDetails> Details;
        }

        private struct VersionInfoDetails
        {
            public VersionInfoDetails(ECCLevel errorCorrectionLevel, Dictionary<EncodingMode, int> capacityDict)
            {
                this.ErrorCorrectionLevel = errorCorrectionLevel;
                this.CapacityDict = capacityDict;
            }

            public readonly ECCLevel ErrorCorrectionLevel;
            public readonly Dictionary<EncodingMode, int> CapacityDict;
        }

        private struct Antilog
        {
            public Antilog(int exponentAlpha, int integerValue)
            {
                this.ExponentAlpha = exponentAlpha;
                this.IntegerValue = integerValue;
            }
            public readonly int ExponentAlpha;
            public readonly int IntegerValue;
        }

        private struct PolynomItem
        {
            public PolynomItem(int coefficient, int exponent)
            {
                this.Coefficient = coefficient;
                this.Exponent = exponent;
            }

            public readonly int Coefficient;
            public readonly int Exponent;
        }

        private class Polynom
        {
            public Polynom()
            {
                this.PolyItems = new List<PolynomItem>();
            }

            public List<PolynomItem> PolyItems { get; set; }

            public override string ToString()
            {
                var sb = new StringBuilder();
                //this.PolyItems.ForEach(x => sb.Append("a^" + x.Coefficient + "*x^" + x.Exponent + " + "));
                foreach (var polyItem in this.PolyItems)
                {
                    sb.Append("a^" + polyItem.Coefficient + "*x^" + polyItem.Exponent + " + ");
                }

                return sb.ToString().TrimEnd(new[] { ' ', '+' });
            }
        }

        private class Point
        {
            public readonly int X;
            public readonly int Y;
            public Point(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }

        private class Rectangle
        {
            public readonly int X;
            public readonly int Y;
            public readonly int Width;
            public readonly int Height;

            public Rectangle(int x, int y, int w, int h)
            {
                this.X = x;
                this.Y = y;
                this.Width = w;
                this.Height = h;
            }
        }

        public void Dispose()
        {
            // left for back-compat
        }
    }
}

// ==== QRCodeData.cs ====
namespace QRCoder
{
    public class QRCodeData
    {
        public List<BitArray> ModuleMatrix { get; set; }
        public int Version { get; private set; }
        public QRCodeData(int version)
        {
            this.Version = version;
            int size = ModulesPerSideFromVersion(version);
            this.ModuleMatrix = new List<BitArray>();
            for (int i = 0; i < size; i++)
                this.ModuleMatrix.Add(new BitArray(size));
        }
        private static int ModulesPerSideFromVersion(int version)
        {
            return 21 + (version - 1) * 4;
        }
    }
}

// ==== Exceptions.cs ====
namespace QRCoder.Exceptions
{
    public class DataTooLongException : Exception
    {
        public DataTooLongException(string eccLevel, string encodingMode, int maxSizeByte)
            : base("Payload too long for QR (ECC=" + eccLevel + ", mode=" + encodingMode + ", max=" + maxSizeByte + "B)") { }
        public DataTooLongException(string eccLevel, string encodingMode, int version, int maxSizeByte)
            : base("Payload too long for QR (ECC=" + eccLevel + ", mode=" + encodingMode + ", ver=" + version + ", max=" + maxSizeByte + "B)") { }
    }
}
'@

function Ensure-QREncoder {
  if (-not ([System.Management.Automation.PSTypeName]'QRCoder.QRCodeGenerator').Type) {
    Add-Type -TypeDefinition $qrEncoderSource -ReferencedAssemblies @('System.dll', 'System.Core.dll', 'System.Drawing.dll')
  }
}

# ---- QR送信 UI ----
$grpQR = New-Object System.Windows.Forms.GroupBox
$grpQR.Text = 'カルテ機→スマホ (QR送信)'; $grpQR.Location = '12,388'; $grpQR.Size = New-Object System.Drawing.Size(522, 220)
$form.Controls.Add($grpQR)

$txtQRText = New-Object System.Windows.Forms.TextBox
$txtQRText.Multiline = $true; $txtQRText.ScrollBars = 'Vertical'; $txtQRText.Location = '12,22'; $txtQRText.Size = New-Object System.Drawing.Size(498, 118)
$grpQR.Controls.Add($txtQRText)

$pnlDrop = New-Object System.Windows.Forms.Panel
$pnlDrop.Location = '12,148'; $pnlDrop.Size = New-Object System.Drawing.Size(290, 58); $pnlDrop.BorderStyle = 'FixedSingle'; $pnlDrop.AllowDrop = $true
$grpQR.Controls.Add($pnlDrop)

$lblDropFile = New-Object System.Windows.Forms.Label
$lblDropFile.Text = 'ここにファイルをドロップ'; $lblDropFile.Dock = 'Fill'; $lblDropFile.TextAlign = 'MiddleCenter'; $lblDropFile.AllowDrop = $true
$pnlDrop.Controls.Add($lblDropFile)

$btnQRClear = New-Object System.Windows.Forms.Button
$btnQRClear.Text = 'クリア'; $btnQRClear.Location = '314,148'; $btnQRClear.Size = New-Object System.Drawing.Size(88, 26)
$grpQR.Controls.Add($btnQRClear)

$chkGzip = New-Object System.Windows.Forms.CheckBox
$chkGzip.Text = 'gzip圧縮'; $chkGzip.Location = '314,182'; $chkGzip.AutoSize = $true
$grpQR.Controls.Add($chkGzip)

$btnShowQR = New-Object System.Windows.Forms.Button
$btnShowQR.Text = 'QR表示'; $btnShowQR.Location = '412,148'; $btnShowQR.Size = New-Object System.Drawing.Size(98, 58)
$grpQR.Controls.Add($btnShowQR)

# ---- ドラッグ&ドロップ ----
$qrDragEnter = {
  if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
  else { $_.Effect = [System.Windows.Forms.DragDropEffects]::None }
}
$qrDragDrop = {
  $files = $_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
  if ($files -ne $null -and $files.Length -gt 0) {
    $script:qrDropFile = [string]$files[0]
    $lblDropFile.Text = 'ファイル: ' + [System.IO.Path]::GetFileName($script:qrDropFile)
  }
}
$pnlDrop.Add_DragEnter($qrDragEnter); $pnlDrop.Add_DragDrop($qrDragDrop)
$lblDropFile.Add_DragEnter($qrDragEnter); $lblDropFile.Add_DragDrop($qrDragDrop)

$btnQRClear.Add_Click({
  $script:qrDropFile = $null
  $lblDropFile.Text = 'ここにファイルをドロップ'
})

# ---- プロトコル ----
function New-QRContainer([byte[]]$data, [bool]$isFile, [string]$name, [bool]$compress) {
  $payload = $data
  if ($compress) {
    $ms = New-Object System.IO.MemoryStream
    $gz = New-Object System.IO.Compression.GZipStream($ms, [System.IO.Compression.CompressionMode]::Compress)
    $gz.Write($data, 0, $data.Length); $gz.Close()
    $payload = $ms.ToArray(); $ms.Close()
  }
  $nameB64 = '-'
  if ($name -ne $null -and $name.Length -gt 0) { $nameB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($name)) }
  $typ = 'T'; if ($isFile) { $typ = 'F' }
  $cf = '0'; if ($compress) { $cf = '1' }
  $header = 'CDCQ1 ' + $typ + ' ' + $cf + ' ' + $nameB64 + ' ' + $data.Length
  $hb = [System.Text.Encoding]::ASCII.GetBytes($header)
  $out = New-Object byte[] ($hb.Length + 1 + $payload.Length)
  [Array]::Copy($hb, 0, $out, 0, $hb.Length)
  $out[$hb.Length] = 0x0A
  [Array]::Copy($payload, 0, $out, $hb.Length + 1, $payload.Length)
  return ,$out
}

function Get-QRTexts([byte[]]$container) {
  $b64 = [Convert]::ToBase64String($container)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($b64))
  $sha.Dispose()
  $id8 = (-join ($hash | ForEach-Object { $_.ToString('x2') })).Substring(0, 8)
  $CH = 120
  $total = [int][Math]::Ceiling($b64.Length / [double]$CH)
  $texts = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $total; $i++) {
    $start = $i * $CH
    $len = [Math]::Min($CH, $b64.Length - $start)
    $texts.Add('Q1|' + $id8 + '|' + $i + '|' + $total + '|' + $b64.Substring($start, $len))
  }
  return @{ Texts = $texts; Id8 = $id8; Total = $total }
}

# ---- QRビットマップ生成 (ModuleMatrix描画, 4モジュール白枠) ----
function New-QRBitmap([string]$text) {
  $gen = New-Object QRCoder.QRCodeGenerator
  $qd = $gen.CreateQrCode($text, [QRCoder.QRCodeGenerator+ECCLevel]::M, $false, $false, [QRCoder.QRCodeGenerator+EciMode]::Default, -1)
  $mm = $qd.ModuleMatrix
  $n = $mm.Count
  $border = 4
  $scale = [int][Math]::Floor(250 / ($n + 2 * $border))
  if ($scale -lt 3) { $scale = 3 }
  if ($scale -gt 8) { $scale = 8 }
  $px = ($n + 2 * $border) * $scale
  $bmp = New-Object System.Drawing.Bitmap($px, $px)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.Clear([System.Drawing.Color]::White)
  for ($r = 0; $r -lt $n; $r++) {
    $row = $mm[$r]
    for ($c = 0; $c -lt $n; $c++) {
      if ($row[$c]) { $g.FillRectangle([System.Drawing.Brushes]::Black, ($c + $border) * $scale, ($r + $border) * $scale, $scale, $scale) }
    }
  }
  $g.Dispose()
  $gen.Dispose()
  return $bmp
}

# ---- QR表示 ----
$btnShowQR.Add_Click({
  $btnShowQR.Enabled = $false
  try {
    Ensure-QREncoder
    $isFile = $false; $name = ''; $data = $null
    if ($script:qrDropFile) {
      $isFile = $true
      $name = [System.IO.Path]::GetFileName($script:qrDropFile)
      $data = [System.IO.File]::ReadAllBytes($script:qrDropFile)
    } else {
      if ([string]::IsNullOrEmpty($txtQRText.Text)) { [void][System.Windows.Forms.MessageBox]::Show('テキストを入力するかファイルをドロップしてください'); return }
      $data = [System.Text.Encoding]::UTF8.GetBytes($txtQRText.Text)
    }
    $container = New-QRContainer $data $isFile $name $chkGzip.Checked
    $qr = Get-QRTexts $container
    Log ('QR送信: id=' + $qr.Id8 + '  ' + $qr.Total + '枚  (container ' + $container.Length + 'B)')

    $qf = New-Object System.Windows.Forms.Form
    $qf.Text = 'QR送信 id=' + $qr.Id8 + ' (' + $qr.Total + '枚)'
    $qf.FormBorderStyle = 'Sizable'; $qf.WindowState = 'Maximized'; $qf.BackColor = [System.Drawing.Color]::White; $qf.StartPosition = 'CenterScreen'

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = 'Fill'; $flow.AutoScroll = $true; $flow.BackColor = [System.Drawing.Color]::White
    $qf.Controls.Add($flow)

    $hdr = New-Object System.Windows.Forms.Label
    $hdr.Text = 'id=' + $qr.Id8 + '  全' + $qr.Total + '枚  スマホ(qr-reader)で全て読み取ってください'
    $hdr.Dock = 'Top'; $hdr.Height = 30; $hdr.TextAlign = 'MiddleLeft'
    $hdr.Font = New-Object System.Drawing.Font('Meiryo UI', 12, [System.Drawing.FontStyle]::Bold)
    $qf.Controls.Add($hdr)

    for ($qi = 0; $qi -lt $qr.Texts.Count; $qi++) {
      $bmp = New-QRBitmap $qr.Texts[$qi]
      $cell = New-Object System.Windows.Forms.Panel
      $cell.Width = $bmp.Width; $cell.Height = $bmp.Height + 22
      $cell.Margin = New-Object System.Windows.Forms.Padding(10)
      $pbx = New-Object System.Windows.Forms.PictureBox
      $pbx.Image = $bmp; $pbx.SizeMode = 'AutoSize'; $pbx.Location = New-Object System.Drawing.Point(0, 0)
      $cap = New-Object System.Windows.Forms.Label
      $cap.Text = '' + ($qi + 1) + ' / ' + $qr.Total
      $cap.Location = New-Object System.Drawing.Point(0, $bmp.Height); $cap.Width = $bmp.Width; $cap.Height = 20; $cap.TextAlign = 'TopCenter'
      $cell.Controls.Add($pbx); $cell.Controls.Add($cap)
      $flow.Controls.Add($cell)
    }
    [void]$qf.Show()
  } catch {
    [void][System.Windows.Forms.MessageBox]::Show('QRエラー: ' + $_.Exception.Message)
    Log ('QRエラー: ' + $_.Exception.Message)
  } finally {
    $btnShowQR.Enabled = $true
  }
})

# ---- Web App 案内QR (スマホでこのQRを読むと受信Web Appが開く) ----
$lblAppQR = New-Object System.Windows.Forms.Label
$lblAppQR.Text = "スマホでこのQRを読むと" + [char]10 + "受信Web Appが開きます"
$lblAppQR.Location = '572,10'; $lblAppQR.Size = New-Object System.Drawing.Size(216, 34); $lblAppQR.TextAlign = 'MiddleCenter'
$form.Controls.Add($lblAppQR)
$appQrB64 = 'iVBORw0KGgoAAAANSUhEUgAAANIAAADSCAIAAACw+wkVAAAEKElEQVR4nO3dQW7kNhRAQSvI/a+sLLLsjSCQfBRRtXa77fEDG39ISdd933+w1j+L3w9kR8NqR0B2BGRHQHYEZEdAdgRkR0B2BGRHQHYEZEdAdgRkR0B2BGRH4N93L7uu668z6kT0k9/i973mveqJed955d/CakdAdgRkR0B2BGTHdybZX/Out30yqb2bLufNib+vGjUR7/+3eMJqR0B2BGRHQHYEZMeXJ9lRU8+o2e333d9Nl6O+5p3rwXu9+z7tRGy1IyA7ArIjIDsCsuOsSXZ/8+bEUTPgfeidpa12BGRHQHYEZEdAdgQOmWTnndQdpb2adTdWOwKyIyA7ArIjIDvOmmRXTo7zThe/M+quUNeg3d7d9natdgRkR0B2BGRHQHZ8eZJt9xznTakrr6W9p91vajdWOwKyIyA7ArIjIDsC1267dfOs3JOd5z7i72W1IyA7ArIjIDsCsuPLz5Nt9y6f/ISjjHr3e+FMOurfcNTPbLUjIDsCsiMgOwKy4/Q92XmT0bt3H2XU9L3yubTtbGu1IyA7ArIjIDsCsuP058nOu1b0yTzVniVe+aSeK73b1RNWOwKyIyA7ArIjIDu+syc7b0fvi9/5id3uN/WE08UcxYcsAdkRkB0B2XH6nuyTr3k3l+02j49ybTahO13Mh/mQJSA7ArIjIDtOfwrPyms8R52VbXczr/ReyvP+Va12BGRHQHYEZEdAdnx5kp13P+FRe7srv/PKE7+/Rs2b8+Zoqx0B2RGQHQHZEZAd35lk5+2uvnvVvPtE7T8n3gvv3eQpPHyYD1kCsiMgOwKy48vPk/21/5Nh503fK6fCa+Fv4TpZPsyHLAHZEZAdAdnx5T3Zd68atXO6crJeuZN7pT+h62Q5ig9ZArIjIDsCsuOsp/Dsf9545T1+n7gX7lDP+/+BJ6x2BGRHQHYEZEdAdnxnkl058bUz1ygrZ9s7fQLvE1Y7ArIjIDsCsiMgO07fk135PJ2V94Da7ezuE+1cb7UjIDsCsiMgOwKy46zrZFd+zbyTzLvdpere7Nk9ThfzGT5kCciOgOwIyI6z7l08z+/0NG8H9t0+8sodz2vh6eJRrHYEZEdAdgRkR0B2fGeSbc/lPrFymht1jvre7ET0vPey2hGQHQHZEZAdAdnx5Um2Pa067/5FK++D9MTKexfP+72sdgRkR0B2BGRHQHacNcmuNG/eHPUc2HnuQdftuk6Ww/mQJSA7ArIjIDsCh0yyu82bTybrK33mbHsnK6sdAdkRkB0B2RGQHWdNsrtdvzlvbm3vf3VNO+3sebIcxYcsAdkRkB0B2fHlSXa3aW7/E8ijzLuyeNSrflntCMiOgOwIyI6A7AhcK3dO4X9WOwKyIyA7ArIjIDsCsiMgOwKyIyA7ArIjIDsCsiMgOwKyIyA7ArIjIDv+1vsP93Iy6cvgIn0AAAAASUVORK5CYII='
$appQrMs = New-Object System.IO.MemoryStream(,[Convert]::FromBase64String($appQrB64))
$picAppQR = New-Object System.Windows.Forms.PictureBox
$picAppQR.Image = [System.Drawing.Image]::FromStream($appQrMs)
$picAppQR.SizeMode = 'AutoSize'; $picAppQR.Location = '575,46'
$form.Controls.Add($picAppQR)
$lnkApp = New-Object System.Windows.Forms.LinkLabel
$lnkApp.Text = 'https://yuya847.github.io/cdc-qr-reader/'
$lnkApp.Location = '560,262'; $lnkApp.Size = New-Object System.Drawing.Size(238, 34); $lnkApp.TextAlign = 'MiddleCenter'
$lnkApp.Add_LinkClicked({ Start-Process 'https://yuya847.github.io/cdc-qr-reader/' })
$form.Controls.Add($lnkApp)

# ---- セットアップ / 接続ガイド (WiFi接続QR + 192.168.4.1 QR) ----
$grpSetup = New-Object System.Windows.Forms.GroupBox
$grpSetup.Text = 'セットアップ / 接続'; $grpSetup.Location = '546,306'; $grpSetup.Size = New-Object System.Drawing.Size(490, 345)
$form.Controls.Add($grpSetup)

$lblSetupInfo = New-Object System.Windows.Forms.Label
$lblSetupInfo.Text = '①XIAO ESP32-S3をUSBに挿す' + [char]10 + '②スマホをWiFi接続（QR①をカメラで読むと自動接続）' + [char]10 + '③ブラウザで 192.168.4.1 を開く（QR②）'
$lblSetupInfo.Location = '12,20'; $lblSetupInfo.Size = New-Object System.Drawing.Size(466, 54)
$grpSetup.Controls.Add($lblSetupInfo)

$lblWifiQR = New-Object System.Windows.Forms.Label
$lblWifiQR.Text = '① WiFi接続 (XIAO-KB-Bridge)'
$lblWifiQR.Location = '12,80'; $lblWifiQR.Size = New-Object System.Drawing.Size(234, 18)
$grpSetup.Controls.Add($lblWifiQR)

$wifiQrB64 = 'iVBORw0KGgoAAAANSUhEUgAAAOoAAADqCAIAAAAj9k4ZAAAFHUlEQVR4nO3dUY6bOgBA0fLU/W+ZtwD6gZCNfcM5303ITK+s4LHNcZ7nH2j6b/UHgOfkS5h8CZMvYfIlTL6EyZcw+RImX8LkS5h8CZMvYfIlTL6EyZcw+RImX8L+PnvZcRx/1rmzQ+TOJ3z2PqOufsf1Wnc+z7H0Mz/zbNeP0Zcw+RImX8LkS5h8+d7Mw9W88yLu3BGPmmd481Xz5hDOae8zyqhZDqMvYfIlTL6EyZcw+RI2bOZh1N3lqLvdtXfN8+7r165MOJb+n14ZfQmTL2HyJUy+hMmXsIkzD2s9mx94c43BqLv4Y+lcxFpGX8LkS5h8CZMvYfIl7EdmHuadYzBqr8ebn/n8zJNSjb6EyZcw+RImX8LkS9jEmYdfvf9du8Lh2edZe+rFPEZfwuRLmHwJky9h8iVs2MzDbucPzDsvYtTJkGv/zR377+Mw+hImX8LkS5h8CZMvYcduf8Xe/1SHeVe/Y96TQIuMvoTJlzD5EiZfwuTL99Y8vHlv+xunPq5dvXA1akfG2pMojL6EyZcw+RImX8Lky/fWPDybDZh3/sCda827ix+1DmH/3Q1Xa1dTGH0Jky9h8iVMvoTJl7BXZx7uvM8z8+Yr5nlzjmW334Y1D+DLA2W++xImX8LkS9jEcx5G7S949s6jrN0lsdtzNI7NZoqMvoTJlzD5EiZfwuRL2OJnW8x7IuT+qxd2221x3JjTmPfUjGe/Q6MvYfIlTL6EyZcw+RI27ITJefe2z95nlLVP55x3rd84+cHoS5h8CZMvYfIlTL6EDVvzMOr+d956/nnnW867r3/zt3EOWrlxNW+OxehLmHwJky9h8iVMvnxv5mG3MwHmrab48mkMx/bP2jD6EiZfwuRLmHwJky9hE9c8jLL2/vfOjMGztRx35hDePKnyat6JoKMYfQmTL2HyJUy+hMmXsMVP1fyNXQnzntDx5hMzz2k7Teb9XEZfwuRLmHwJky9h8iVs2G6LtXeyu92P33nnZ1fff4/G6YRJuMOXB8LkS5h8CZMv31vz8I83Wrp6f+2a/+LPdU7bf3HnVaMYfQmTL2HyJUy+hMmXsGEzD/9460H3pLudSLD/tY7td1LYbQG+PFDmuy9h8iVMvoQNO+dhnt12Ury5nuGZ48VVEKM454HP8eWBMPkSJl/C5EvYxBMm31wJMOoT3rnW/isK3rT2nEyjL2HyJUy+hMmXMPkStnjmYd7ZAs+8+VNcvfnbOJeejHFlzQOf48sDYfIlTL6EyZewibstdpuLKO6teHNOYxRrHuAWXx4Iky9h8iVMvoRNPGHyTfufonDHr87DXDlhEnx5oMx3X8LkS5h8Cfv77GVv/g396nq3++Y8w5v7He686hj0f7H2mZ7PGH0Jky9h8iVMvoTJl+/NPFwV1xi8uaLg+qo7d/r7r4I4lz4Rw+hLmHwJky9h8iVMvoQNm3nY7a/hb65eeHatO6+a51j6k45i9CVMvoTJlzD5EiZfwibOPKz15tMedzu78hg05/PsWm/+5o2+hMmXMPkSJl/C5EvYz848zFvPMOoO/dm/2W1NyPloLmIUoy9h8iVMvoTJlzD5EjZx5mHtUzPmPffzzvvMu9bVMe0Mh/2vbvQlTL6EyZcw+RImX773VM3dnm0xag/CqNMY5s0z7P8k0Kt5P7vRlzD5EiZfwuRLmHz53swD7MDoS5h8CZMvYfIlTL6EyZcw+RImX8LkS5h8CZMvYfIlTL6EyZcw+RImX8Lky5+u/wHq7/7K33SLDwAAAABJRU5ErkJggg=='
$wifiQrMs = New-Object System.IO.MemoryStream(,[Convert]::FromBase64String($wifiQrB64))
$picWifiQR = New-Object System.Windows.Forms.PictureBox
$picWifiQR.Image = [System.Drawing.Image]::FromStream($wifiQrMs)
$picWifiQR.SizeMode = 'AutoSize'; $picWifiQR.Location = '12,100'
$grpSetup.Controls.Add($picWifiQR)

$lblUrlQR = New-Object System.Windows.Forms.Label
$lblUrlQR.Text = '② 192.168.4.1 を開く'
$lblUrlQR.Location = '256,80'; $lblUrlQR.Size = New-Object System.Drawing.Size(226, 18)
$grpSetup.Controls.Add($lblUrlQR)

$url192QrB64 = 'iVBORw0KGgoAAAANSUhEUgAAALoAAAC6CAIAAACWbMCmAAADY0lEQVR4nO3dS27bQBQAQTPI/a+sHIAbNjE/xVVrx1aExkgPMySvz+fzA8/8efhzIBcaqwuBXAjkQiAXArkQyIVALgRyIZALgVwI5EIgFwK5EMiFQC4Ef39eua7rZ5/7CcB3r+fd7/k8OH94/z2jXvMo705RWl0I5EIgFwK5EMiF+ZPR3bzrlVZOEE/ml+vB1PPu3Tj/PbS6EMiFQC4EciGQCzsmo7tR+zijvNvr2es67D20uhDIhUAuBHIhkAtnTEYrvTurNurM2zXoDN75rC4EciGQC4FcCOTC75uM9l7X8/kvpp4nrC4EciGQC4FcCOTCGZPRynnhyRz07hqid64Hf/0bZy6rC4FcCORCIBcCubBjMtp7hc67OWjUz4xy2lVOd1YXArkQyIVALgRyIbhO25VYadSsNG/P6DRWFwK5EMiFQC4EcmHH84xOu/Pbk9czalq5tk5YK+9fZ3UhkAuBXAjkQiAXTr0Dw8rn9aycVu5OO3E36p23uhDIhUAuBHIhkAvzJ6OVOxfv7om98k5019a7Paw83Wd1IZALgVwI5EIgF3acpjtt6nn3zf+0837X1v2pO6sLgVwI5EIgFwK5cMYdGEbtIo2anuZNaqedebuzZ8QGPowI5EIgFwK5MH8yGvVNe+/Uc9qMc02by1xnxAY+jAjkQiAXArmwYzJaad6+0ry/frf3nXedEdP5MCKQC4FcCOTCjnvT7Z0XRv2eUf/qmrantvfptFYXArkQyIVALgRy4Yy7do/6fj7vxN3Kk4TXwl02p+k4gg8j5MIcVhcCuXDGHRhW2nu674m9U5jrjNjAhxGBXAjkQiAXzrhr9zynnXkbdbftUVPPvP+p1YVALgRyIZALgVw44zqjUc5/Zutn6zVNd/aMOIIPIwK5EMiFQC6ccZ3Rac/0eWfUs1avQfOdPSO+hg8jArkQyIVALpwxGX2jlTsy17Q7jc/bZbO6EMiFQC4EciGQC79vMpp3fdC8PazPoCls5TVfVhcCuRDIhUAuBHLh1Lt2n2/lU4e+kdWFQC4EciGQC4Fc2DEZ7f3mP+/OCaOewfSZ9npWPk3J6kIgFwK5EMiFQC78vucZsYbVhUAuBHIhkAuBXAjkQiAXArkQyIVALgRyIZALgVwI5EIgFwK5EMiFn+f+AU+le5aIdhggAAAAAElFTkSuQmCC'
$url192QrMs = New-Object System.IO.MemoryStream(,[Convert]::FromBase64String($url192QrB64))
$picUrlQR = New-Object System.Windows.Forms.PictureBox
$picUrlQR.Image = [System.Drawing.Image]::FromStream($url192QrMs)
$picUrlQR.SizeMode = 'AutoSize'; $picUrlQR.Location = '256,100'
$grpSetup.Controls.Add($picUrlQR)

RefreshPorts
[void]$form.ShowDialog()
