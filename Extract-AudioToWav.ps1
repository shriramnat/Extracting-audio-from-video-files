param(
  [string]$Folder = "",
  [string]$File = "",

  [string]$OutFolder = "",

  [ValidateSet("pcm_s16le","pcm_s24le","pcm_s32le")]
  [string]$PcmCodec = "pcm_s24le",

  [int]$ForceSampleRate = 0,
  [int]$ForceChannels = 0,

  [switch]$Overwrite,
  [switch]$UseW64,
  [switch]$VerboseDiag
)

function Log {
  param(
    [string]$Message,
    [ValidateSet("INFO","WARN","ERROR","OK","DIAG")]
    [string]$Level = "INFO"
  )
  $ts = (Get-Date).ToString("HH:mm:ss")
  switch ($Level) {
    "INFO"  { Write-Host "[$ts] [INFO ] $Message" -ForegroundColor Cyan }
    "WARN"  { Write-Host "[$ts] [WARN ] $Message" -ForegroundColor Yellow }
    "ERROR" { Write-Host "[$ts] [ERROR] $Message" -ForegroundColor Red }
    "OK"    { Write-Host "[$ts] [ OK  ] $Message" -ForegroundColor Green }
    "DIAG"  { Write-Host "[$ts] [DIAG ] $Message" -ForegroundColor DarkGray }
  }
}

function Ensure-Tool {
  param([string]$Name)
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Missing dependency: '$Name' not found in PATH."
  }
}

function Get-DefaultOutFolder {
  param([string]$ModePath)
  if (Test-Path $ModePath -PathType Container) {
    Join-Path $ModePath "_WAV"
  } else {
    Join-Path (Split-Path -Parent $ModePath) "_WAV"
  }
}

function Sanitize-Name {
  param([string]$Name)
  if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
  $bad = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($c in $bad) { $Name = $Name.Replace($c, "_") }
  ($Name -replace "\s+", " ").Trim()
}

function Invoke-FFProbeJson {
  param([string[]]$ProbeArgs)

  if ($VerboseDiag) { Log ("ffprobe " + ($ProbeArgs -join " ")) "DIAG" }

  $out = & ffprobe @ProbeArgs 2>&1
  if (-not $out) { throw "ffprobe returned no output" }

  try { return ($out | ConvertFrom-Json) }
  catch {
    Log "ffprobe output was not valid JSON:" "ERROR"
    Log ($out -join "`n") "ERROR"
    throw
  }
}

function Get-AudioStreams {
  param([string]$Path)

  $json = Invoke-FFProbeJson @(
    "-v","error",
    "-print_format","json",
    "-show_entries","stream=index,codec_name,sample_rate,channels,channel_layout,tags",
    "-select_streams","a",
    $Path
  )

  # Force array shape ALWAYS
  $streams = @()
  if ($null -ne $json.streams) { $streams = @($json.streams) }
  return $streams
}

function Build-OutPath {
  param(
    [System.IO.FileInfo]$InputFile,
    [int]$GlobalStreamIndex,
    [hashtable]$Tags,
    [string]$OutputDir,
    [string]$Ext
  )

  $base = [System.IO.Path]::GetFileNameWithoutExtension($InputFile.Name)

  $lang  = ""
  $title = ""
  if ($Tags) {
    if ($Tags.ContainsKey("language")) { $lang = Sanitize-Name $Tags["language"] }
    if ($Tags.ContainsKey("title")) { $title = Sanitize-Name $Tags["title"] }
    if (-not $title -and $Tags.ContainsKey("handler_name")) { $title = Sanitize-Name $Tags["handler_name"] }
  }

  $suffixParts = @()
  if ($lang)  { $suffixParts += $lang }
  if ($title) { $suffixParts += $title }

  $suffix = ""
  if ($suffixParts.Count -gt 0) { $suffix = "__" + ($suffixParts -join "_") }

  $name = "{0}__idx{1}{2}{3}" -f $base, $GlobalStreamIndex, $suffix, $Ext
  Join-Path $OutputDir $name
}

function Extract-OneStream {
  param(
    [System.IO.FileInfo]$InputFile,
    [int]$GlobalStreamIndex,
    [string]$OutPath,
    [string]$Container
  )

  if ((Test-Path $OutPath) -and (-not $Overwrite)) {
    Log "Skipping (exists): $OutPath" "WARN"
    return [pscustomobject]@{
      FileName = $InputFile.Name; FullPath = $InputFile.FullName
      StreamIndex = $GlobalStreamIndex; OutPath = $OutPath
      Status = "SkippedExists"; Error = $null
    }
  }

  $args = @()
  if ($Overwrite) { $args += "-y" } else { $args += "-n" }

  $args += @(
    "-hide_banner",
    "-v","error",
    "-i", $InputFile.FullName,

    # IMPORTANT: map by GLOBAL stream index, not a:0 ordering
    "-map", ("0:{0}" -f $GlobalStreamIndex),

    "-vn","-sn","-dn",
    "-c:a", $PcmCodec
  )

  if ($ForceSampleRate -gt 0) { $args += @("-ar", "$ForceSampleRate") }
  if ($ForceChannels -gt 0)   { $args += @("-ac", "$ForceChannels") }

  $args += @("-f", $Container, $OutPath)

  if ($VerboseDiag) { Log ("ffmpeg " + ($args -join " ")) "DIAG" }

  # Capture stderr so failures are visible
  $ffOut = & ffmpeg @args 2>&1
  if ($LASTEXITCODE -ne 0) {
    Log ("ffmpeg failed for {0} stream {1}. Output:`n{2}" -f $InputFile.Name, $GlobalStreamIndex, ($ffOut -join "`n")) "ERROR"
    return [pscustomobject]@{
      FileName = $InputFile.Name; FullPath = $InputFile.FullName
      StreamIndex = $GlobalStreamIndex; OutPath = $OutPath
      Status = "Failed"; Error = "ffmpeg exit code $LASTEXITCODE"
    }
  }

  # Confirm file exists (belt + suspenders)
  if (-not (Test-Path $OutPath)) {
    Log "ffmpeg returned success but output file missing: $OutPath" "ERROR"
    return [pscustomobject]@{
      FileName = $InputFile.Name; FullPath = $InputFile.FullName
      StreamIndex = $GlobalStreamIndex; OutPath = $OutPath
      Status = "Failed"; Error = "OutputMissing"
    }
  }

  Log "Created: $([IO.Path]::GetFileName($OutPath))" "OK"
  return [pscustomobject]@{
    FileName = $InputFile.Name; FullPath = $InputFile.FullName
    StreamIndex = $GlobalStreamIndex; OutPath = $OutPath
    Status = "OK"; Error = $null
  }
}

# ---------------- Setup ----------------
Ensure-Tool "ffmpeg"
Ensure-Tool "ffprobe"

$hasFolder = -not [string]::IsNullOrWhiteSpace($Folder)
$hasFile   = -not [string]::IsNullOrWhiteSpace($File)
if (($hasFolder -and $hasFile) -or (-not $hasFolder -and -not $hasFile)) {
  throw "Provide exactly one: -Folder <path> OR -File <path>"
}

if ($hasFolder) {
  if (-not (Test-Path $Folder -PathType Container)) { throw "Folder not found: $Folder" }
  if ([string]::IsNullOrWhiteSpace($OutFolder)) { $OutFolder = Get-DefaultOutFolder $Folder }
} else {
  if (-not (Test-Path $File -PathType Leaf)) { throw "File not found: $File" }
  if ([string]::IsNullOrWhiteSpace($OutFolder)) { $OutFolder = Get-DefaultOutFolder $File }
}

New-Item -ItemType Directory -Force -Path $OutFolder | Out-Null

$container = "wav"; $ext = ".wav"
if ($UseW64) { $container = "w64"; $ext = ".w64" }

Log "Output folder: $OutFolder"
Log ("PCM codec: {0} | ForceSampleRate: {1} | ForceChannels: {2} | Container: {3}" -f $PcmCodec, $ForceSampleRate, $ForceChannels, $container)

# ---------------- Inputs ----------------
$inputs = @()
if ($hasFolder) {
  $inputs = Get-ChildItem -Path $Folder -File -Recurse
  Log ("Folder mode: scanning {0} files (no extension filtering)" -f $inputs.Count) "INFO"
} else {
  $inputs = @(Get-Item $File)
  Log "Single-file mode." "INFO"
}

# ---------------- Process ----------------
$report = @()
$index = 0

foreach ($item in $inputs) {
  $index++
  if ($inputs.Count -gt 1) {
    Write-Progress -Activity "Extracting audio streams to PCM WAV" -Status "$index / $($inputs.Count): $($item.Name)" -PercentComplete (($index / [double]$inputs.Count) * 100)
  }

  Log "---- File $index/$($inputs.Count): $($item.FullName)" "INFO"

  $streams = @()
  try { $streams = Get-AudioStreams -Path $item.FullName }
  catch {
    Log "Skipping (not probe-able as media)" "WARN"
    $report += [pscustomobject]@{
      FileName=$item.Name; FullPath=$item.FullName; StreamIndex=$null; OutPath=$null; Status="SkippedNotMedia"; Error=$_.Exception.Message
    }
    Write-Host ""
    continue
  }

  if (-not $streams -or $streams.Count -eq 0) {
    Log "No audio streams found. Skipping." "WARN"
    $report += [pscustomobject]@{
      FileName=$item.Name; FullPath=$item.FullName; StreamIndex=$null; OutPath=$null; Status="SkippedNoAudio"; Error=$null
    }
    Write-Host ""
    continue
  }

  Log ("Found {0} audio stream(s)." -f $streams.Count) "OK"

  foreach ($st in $streams) {
    $globalIndex = [int]$st.index

    $tags = @{}
    if ($st.tags) { $st.tags.psobject.Properties | ForEach-Object { $tags[$_.Name] = $_.Value } }

    Log ("Stream idx={0} | codec={1} | sr={2} | ch={3} | layout={4}" -f $globalIndex, $st.codec_name, $st.sample_rate, $st.channels, $st.channel_layout) "DIAG"

    $outPath = Build-OutPath -InputFile $item -GlobalStreamIndex $globalIndex -Tags $tags -OutputDir $OutFolder -Ext $ext
    Log ("Output: {0}" -f $outPath) "DIAG"

    $report += Extract-OneStream -InputFile $item -GlobalStreamIndex $globalIndex -OutPath $outPath -Container $container
  }

  Write-Host ""
}

if ($inputs.Count -gt 1) {
  Write-Progress -Activity "Extracting audio streams to PCM WAV" -Completed
}

# ---------------- Report ----------------
$csv = Join-Path $OutFolder "wav_extraction_report.csv"
$report | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csv
Log "Report written: $csv" "OK"

$ok = ($report | Where-Object Status -eq "OK").Count
$sk = ($report | Where-Object Status -like "Skipped*").Count
$fl = ($report | Where-Object Status -eq "Failed").Count
Log ("Summary: OK={0}  Skipped={1}  Failed={2}" -f $ok, $sk, $fl) "INFO"
