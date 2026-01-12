<#
.SYNOPSIS
    Extracts audio streams from video/media files to PCM WAV or W64 format.

.DESCRIPTION
    This PowerShell script uses FFmpeg to extract all audio streams from video or media files
    and converts them to uncompressed PCM WAV (or W64) format. It can process a single file
    or recursively scan an entire folder. Each audio stream is extracted separately with
    metadata-based naming (language, title, stream index).

.PARAMETER Folder
    Path to a folder to scan recursively for media files. Cannot be used with -File.
    Example: -Folder "C:\Videos"

.PARAMETER File
    Path to a single media file to process. Cannot be used with -Folder.
    Example: -File "C:\Videos\movie.mp4"

.PARAMETER OutFolder
    Output directory for extracted audio files. If not specified, defaults to "_WAV" 
    subdirectory relative to the input folder or file location.
    Example: -OutFolder "C:\Output"

.PARAMETER PcmCodec
    PCM codec to use for audio encoding. Valid values: "pcm_s16le" (16-bit), 
    "pcm_s24le" (24-bit), "pcm_s32le" (32-bit). Default: "pcm_s24le"
    Example: -PcmCodec "pcm_s16le"

.PARAMETER ForceSampleRate
    Force a specific sample rate (Hz) for output audio. 0 = keep original. Default: 0
    Example: -ForceSampleRate 48000

.PARAMETER ForceChannels
    Force a specific number of audio channels. 0 = keep original. Default: 0
    Example: -ForceChannels 2

.PARAMETER Overwrite
    If set, overwrites existing output files. Without this flag, existing files are skipped.

.PARAMETER UseW64
    If set, outputs W64 format instead of WAV. W64 supports larger file sizes (>4GB).

.PARAMETER VerboseDiag
    If set, enables verbose diagnostic output showing ffmpeg/ffprobe commands.

.EXAMPLE
    .\Extract-AudioToWav.ps1 -Folder "C:\Videos" -PcmCodec "pcm_s16le"
    Extracts audio from all files in C:\Videos to 16-bit WAV format.

.EXAMPLE
    .\Extract-AudioToWav.ps1 -File "movie.mp4" -ForceSampleRate 48000 -ForceChannels 2
    Extracts audio from movie.mp4, resampling to 48kHz stereo.

.EXAMPLE
    .\Extract-AudioToWav.ps1 -Folder "C:\Videos" -UseW64 -Overwrite
    Extracts audio to W64 format, overwriting existing files.

.NOTES
    Prerequisites:
    - FFmpeg and FFprobe must be installed and available in PATH
    - Tested with PowerShell 5.1+
    
    Output Naming Convention:
    - Format: {basename}__idx{stream_index}__{language}_{title}.{ext}
    - Example: movie__idx1__eng_Commentary.wav

.LINK
    https://github.com/shriramnat/Extracting-audio-from-video-files
#>

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

<#
.SYNOPSIS
    Outputs a timestamped, color-coded log message to the console.

.PARAMETER Message
    The message text to display.

.PARAMETER Level
    Log level: INFO (cyan), WARN (yellow), ERROR (red), OK (green), DIAG (dark gray).
#>
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

<#
.SYNOPSIS
    Verifies that a required tool/command is available in the system PATH.

.PARAMETER Name
    The name of the command/tool to check (e.g., "ffmpeg", "ffprobe").

.NOTES
    Throws an exception if the tool is not found.
#>
function Ensure-Tool {
  param([string]$Name)
  if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
    throw "Missing dependency: '$Name' not found in PATH."
  }
}

<#
.SYNOPSIS
    Determines the default output folder path based on the input path.

.PARAMETER ModePath
    The input file or folder path.

.OUTPUTS
    Returns the path to a "_WAV" subdirectory relative to the input location.

.NOTES
    If input is a folder, creates "_WAV" inside it.
    If input is a file, creates "_WAV" in the parent directory.
#>
function Get-DefaultOutFolder {
  param([string]$ModePath)
  if (Test-Path $ModePath -PathType Container) {
    Join-Path $ModePath "_WAV"
  } else {
    Join-Path (Split-Path -Parent $ModePath) "_WAV"
  }
}

<#
.SYNOPSIS
    Sanitizes a string for use as a safe filename component.

.PARAMETER Name
    The string to sanitize.

.OUTPUTS
    Returns a sanitized string with invalid filename characters replaced by underscores,
    and multiple spaces collapsed to single spaces.

.NOTES
    Removes all characters that are invalid in Windows/Unix filenames.
#>
function Sanitize-Name {
  param([string]$Name)
  if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
  $bad = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($c in $bad) { $Name = $Name.Replace($c, "_") }
  ($Name -replace "\s+", " ").Trim()
}

<#
.SYNOPSIS
    Executes ffprobe and returns its output as a PowerShell object (parsed from JSON).

.PARAMETER ProbeArgs
    Array of arguments to pass to ffprobe.

.OUTPUTS
    Returns the parsed JSON output from ffprobe as a PowerShell object.

.NOTES
    Throws an exception if ffprobe fails or returns invalid JSON.
    Displays diagnostic information when -VerboseDiag is enabled.
#>
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

<#
.SYNOPSIS
    Retrieves information about all audio streams in a media file.

.PARAMETER Path
    Full path to the media file to probe.

.OUTPUTS
    Returns an array of audio stream objects containing metadata like codec, sample rate,
    channels, and tags.

.NOTES
    Uses ffprobe to analyze the file. Returns an empty array if no audio streams are found.
#>
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

<#
.SYNOPSIS
    Constructs the output file path for an extracted audio stream.

.PARAMETER InputFile
    The source media file.

.PARAMETER GlobalStreamIndex
    The global stream index (0-based) from the media container.

.PARAMETER Tags
    Hashtable of stream tags (language, title, handler_name, etc.).

.PARAMETER OutputDir
    The output directory path.

.PARAMETER Ext
    The file extension to use (.wav or .w64).

.OUTPUTS
    Returns the full output path with naming convention:
    {basename}__idx{stream_index}__{language}_{title}.{ext}

.NOTES
    Sanitizes language and title metadata to create safe, descriptive filenames.
#>
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

<#
.SYNOPSIS
    Extracts a single audio stream from a media file to PCM WAV/W64 format.

.PARAMETER InputFile
    The source media file.

.PARAMETER GlobalStreamIndex
    The global stream index (0-based) to extract.

.PARAMETER OutPath
    The output file path for the extracted audio.

.PARAMETER Container
    The output container format ("wav" or "w64").

.OUTPUTS
    Returns a custom object with extraction status: FileName, FullPath, StreamIndex, 
    OutPath, Status (OK/Failed/SkippedExists), and Error message.

.NOTES
    Uses ffmpeg to perform the extraction. Respects -Overwrite, -ForceSampleRate, 
    and -ForceChannels parameters from the parent scope.
#>
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

# ================================================================================
# INITIALIZATION - Verify prerequisites and validate parameters
# ================================================================================

Ensure-Tool "ffmpeg"
Ensure-Tool "ffprobe"

# Validate that exactly one input mode is specified
$hasFolder = -not [string]::IsNullOrWhiteSpace($Folder)
$hasFile   = -not [string]::IsNullOrWhiteSpace($File)
if (($hasFolder -and $hasFile) -or (-not $hasFolder -and -not $hasFile)) {
  throw "Provide exactly one: -Folder <path> OR -File <path>"
}

# Determine and create output directory
if ($hasFolder) {
  if (-not (Test-Path $Folder -PathType Container)) { throw "Folder not found: $Folder" }
  if ([string]::IsNullOrWhiteSpace($OutFolder)) { $OutFolder = Get-DefaultOutFolder $Folder }
} else {
  if (-not (Test-Path $File -PathType Leaf)) { throw "File not found: $File" }
  if ([string]::IsNullOrWhiteSpace($OutFolder)) { $OutFolder = Get-DefaultOutFolder $File }
}

New-Item -ItemType Directory -Force -Path $OutFolder | Out-Null

# Configure output format (WAV or W64)
$container = "wav"; $ext = ".wav"
if ($UseW64) { $container = "w64"; $ext = ".w64" }

# Display configuration summary
Log "Output folder: $OutFolder"
Log ("PCM codec: {0} | ForceSampleRate: {1} | ForceChannels: {2} | Container: {3}" -f $PcmCodec, $ForceSampleRate, $ForceChannels, $container)

# ================================================================================
# INPUT COLLECTION - Build list of files to process
# ================================================================================

$inputs = @()
if ($hasFolder) {
  $inputs = Get-ChildItem -Path $Folder -File -Recurse
  Log ("Folder mode: scanning {0} files (no extension filtering)" -f $inputs.Count) "INFO"
} else {
  $inputs = @(Get-Item $File)
  Log "Single-file mode." "INFO"
}

# ================================================================================
# MAIN PROCESSING LOOP - Extract audio streams from each file
# ================================================================================

$report = @()
$index = 0

foreach ($item in $inputs) {
  $index++
  if ($inputs.Count -gt 1) {
    Write-Progress -Activity "Extracting audio streams to PCM WAV" -Status "$index / $($inputs.Count): $($item.Name)" -PercentComplete (($index / [double]$inputs.Count) * 100)
  }

  Log "---- File $index/$($inputs.Count): $($item.FullName)" "INFO"

  # Probe the file for audio streams
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

  # Skip files with no audio streams
  if (-not $streams -or $streams.Count -eq 0) {
    Log "No audio streams found. Skipping." "WARN"
    $report += [pscustomobject]@{
      FileName=$item.Name; FullPath=$item.FullName; StreamIndex=$null; OutPath=$null; Status="SkippedNoAudio"; Error=$null
    }
    Write-Host ""
    continue
  }

  Log ("Found {0} audio stream(s)." -f $streams.Count) "OK"

  # Process each audio stream
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

# Clear progress indicator
if ($inputs.Count -gt 1) {
  Write-Progress -Activity "Extracting audio streams to PCM WAV" -Completed
}

# ================================================================================
# REPORTING - Generate CSV report and summary
# ================================================================================

$csv = Join-Path $OutFolder "wav_extraction_report.csv"
$report | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csv
Log "Report written: $csv" "OK"

$ok = ($report | Where-Object Status -eq "OK").Count
$sk = ($report | Where-Object Status -like "Skipped*").Count
$fl = ($report | Where-Object Status -eq "Failed").Count
Log ("Summary: OK={0}  Skipped={1}  Failed={2}" -f $ok, $sk, $fl) "INFO"
