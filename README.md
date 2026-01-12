# Extract-AudioToWav

A PowerShell script that extracts audio streams from video and media files, converting them to uncompressed PCM WAV or W64 format using FFmpeg.

## Overview

This script uses FFmpeg to extract all audio streams from video/media files and converts them to high-quality PCM WAV format. It's designed for batch processing and handles multiple audio streams per file, preserving stream metadata in the output filenames.

### Key Features

- **Batch Processing**: Process a single file or recursively scan entire folders
- **Multiple Audio Streams**: Extracts all audio streams from files with multiple audio tracks
- **Metadata-Based Naming**: Output files include language, title, and stream index information
- **Format Options**: Supports both WAV and W64 (for files >4GB)
- **PCM Quality Control**: Choose between 16-bit, 24-bit, or 32-bit PCM encoding
- **Audio Resampling**: Optional sample rate and channel conversion
- **Progress Tracking**: Real-time progress indicator and detailed logging
- **Comprehensive Reporting**: Generates a CSV report of all extraction operations

## Prerequisites

Before using this script, ensure you have the following installed:

- **PowerShell**: Version 5.1 or higher (pre-installed on Windows 10+)
- **FFmpeg**: Must be installed and available in your system PATH
  - Download from: https://ffmpeg.org/download.html
  - Verify installation: `ffmpeg -version`
- **FFprobe**: Included with FFmpeg installation
  - Verify installation: `ffprobe -version`

### Installing FFmpeg on Windows

1. Download FFmpeg from the official website
2. Extract the archive to a location (e.g., `C:\ffmpeg`)
3. Add the `bin` folder to your system PATH environment variable
4. Restart PowerShell and verify with `ffmpeg -version`

## Parameters

### Input Parameters (mutually exclusive)

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-Folder` | String | Yes* | Path to a folder to scan recursively for media files |
| `-File` | String | Yes* | Path to a single media file to process |

*Exactly one of `-Folder` or `-File` must be specified.

### Output Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-OutFolder` | String | `_WAV/` | Output directory for extracted audio files. Defaults to `_WAV` subdirectory relative to input location |

### Audio Encoding Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-PcmCodec` | String | `pcm_s24le` | PCM codec: `pcm_s16le` (16-bit), `pcm_s24le` (24-bit), or `pcm_s32le` (32-bit) |
| `-ForceSampleRate` | Integer | 0 | Force specific sample rate in Hz (e.g., 48000). 0 = keep original |
| `-ForceChannels` | Integer | 0 | Force specific number of channels. 0 = keep original |

### Behavioral Flags

| Parameter | Type | Description |
|-----------|------|-------------|
| `-Overwrite` | Switch | Overwrite existing output files. Without this, existing files are skipped |
| `-UseW64` | Switch | Output W64 format instead of WAV. W64 supports files larger than 4GB |
| `-VerboseDiag` | Switch | Enable verbose diagnostic output showing all ffmpeg/ffprobe commands |

## Usage Examples

### Basic Usage

Extract audio from a single video file:
```powershell
.\Extract-AudioToWav.ps1 -File "C:\Videos\movie.mp4"
```

Extract audio from all files in a folder:
```powershell
.\Extract-AudioToWav.ps1 -Folder "C:\Videos"
```

### Advanced Usage

Extract with custom output directory and 16-bit PCM:
```powershell
.\Extract-AudioToWav.ps1 -Folder "C:\Videos" -OutFolder "C:\Audio\Output" -PcmCodec "pcm_s16le"
```

Force 48kHz stereo output with overwrite:
```powershell
.\Extract-AudioToWav.ps1 -File "movie.mkv" -ForceSampleRate 48000 -ForceChannels 2 -Overwrite
```

Use W64 format for large files:
```powershell
.\Extract-AudioToWav.ps1 -Folder "C:\Videos" -UseW64
```

Enable verbose diagnostics for troubleshooting:
```powershell
.\Extract-AudioToWav.ps1 -File "movie.mp4" -VerboseDiag
```

## Output

### File Naming Convention

Output files follow this naming pattern:
```
{basename}__idx{stream_index}__{language}_{title}.{ext}
```

**Examples:**
- `movie__idx1.wav` - Single audio stream with no metadata
- `movie__idx1__eng_Director Commentary.wav` - Stream with language and title
- `movie__idx2__spa.wav` - Stream with language only

### Directory Structure

By default, extracted audio files are saved to a `_WAV` folder:

**Folder mode:**
```
C:\Videos\
├── movie1.mp4
├── movie2.mkv
└── _WAV\
    ├── movie1__idx1__eng.wav
    ├── movie2__idx1__eng.wav
    ├── movie2__idx2__spa.wav
    └── wav_extraction_report.csv
```

**File mode:**
```
C:\Videos\
├── movie.mp4
└── _WAV\
    ├── movie__idx1__eng.wav
    └── wav_extraction_report.csv
```

### Extraction Report

A CSV report (`wav_extraction_report.csv`) is automatically generated in the output folder with the following columns:

- **FileName**: Original media filename
- **FullPath**: Full path to source file
- **StreamIndex**: Global stream index
- **OutPath**: Path to extracted audio file
- **Status**: OK, Failed, SkippedExists, SkippedNoAudio, or SkippedNotMedia
- **Error**: Error message if extraction failed

### Console Output

The script provides color-coded, timestamped logging:
- **Cyan (INFO)**: General information and progress
- **Green (OK)**: Successful operations
- **Yellow (WARN)**: Warnings and skipped files
- **Red (ERROR)**: Errors and failures
- **Dark Gray (DIAG)**: Diagnostic information (requires `-VerboseDiag`)

## Troubleshooting

### Common Issues

**"Missing dependency: 'ffmpeg' not found in PATH"**
- Solution: Install FFmpeg and ensure it's in your system PATH

**"Folder not found" or "File not found"**
- Solution: Verify the path is correct and use absolute paths or proper relative paths

**"ffprobe returned no output"**
- Solution: The file may be corrupted or not a valid media file

**Files are being skipped**
- Check the console output for warnings
- Existing files are skipped unless `-Overwrite` is specified
- Files without audio streams are skipped automatically

### Getting Help

View built-in help documentation:
```powershell
Get-Help .\Extract-AudioToWav.ps1 -Full
```

View parameter details:
```powershell
Get-Help .\Extract-AudioToWav.ps1 -Parameter *
```

View examples:
```powershell
Get-Help .\Extract-AudioToWav.ps1 -Examples
```

## Technical Details

### Audio Processing

- **Extraction Method**: Uses FFmpeg's `-map` option with global stream indexing
- **Codec Conversion**: Converts all audio to uncompressed PCM (Pulse Code Modulation)
- **Stream Mapping**: Preserves original stream order using global stream indices
- **Metadata Preservation**: Extracts language and title tags from stream metadata

### Supported Input Formats

The script supports any media format that FFmpeg can read, including:
- Video: MP4, MKV, AVI, MOV, WMV, FLV, WEBM
- Audio: MP3, AAC, FLAC, OGG, WMA
- Containers with multiple audio tracks

### Performance Considerations

- Processing time depends on file size and audio complexity
- PCM conversion creates larger files (uncompressed audio)
- W64 format recommended for source files with very long audio (>4GB output)
- Batch processing displays progress indicator for visual feedback

## License

This script is provided as-is for free use and modification.

## Contributing

Contributions, issues, and feature requests are welcome!

## Author

Created for extracting audio from video files efficiently and reliably.