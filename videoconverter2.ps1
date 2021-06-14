## Videoconverter
## v 3.0
## Powershell script

param (

	[Parameter(Mandatory=$true)][String]$file,
	[int]$crf = 20,
	[String]$outfile,
	[switch]$nocopy,
	[switch]$seccond_audio,
	[switch]$hard_subtitles,
	[int[]]$parts

)

### General definitions ###
# Path to ffmpeg
[string]$ffmpeg = "ffmpeg"

# Path to ffprobe
[string]$ffprobe = "ffprobe.exe"

# Directory to download local copy of source file to
[string]$tempFolder = "W:\"

# Directory to store final file into
[string]$targetFolder = "X:\recorder\konv\"

# Addition for seccond audio source
$suffix = "English"

# Set level of video compression ( ultrafast | faster | fast | medium | slow | slower | placebo )
$video_encoding_preset="slower"

### Variables ###
[int]$fehler = 0
[string]$source_File_1 = ""
[string]$source_File_2 = ""
[string]$target_File = ""
[string]$input_parameters = ""
[string]$filter_parameters = ""
[string]$output_parameters = ""

### Functions ###
function get-FormatedTime-from-Frame ( [int]$frame, [int]$FramesPerSec ) {

	([double]$timeCalculated = $frame / $FramesPerSec) | Out-Null
	([string]$timeFormatted = "{0:HH:mm:ss.fff}" -f ([datetime]([timespan]::fromseconds($timeCalculated)).Ticks)) | Out-Null
	return $timeFormatted

}

function system_disable_sleep {

	# Request system availability, no sleep idle timeout while executing
	$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
  public static extern void SetThreadExecutionState(uint esFlags);
'@
	$ste = Add-Type -memberDefinition $code -name System -namespace Win32 -passThru 
	$ste::SetThreadExecutionState([uint32]"0x80000000" -bor [uint32]"0x00000001")

}

function system_allow_sleep {

	# Return to default sleep timeout
	$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
  public static extern void SetThreadExecutionState(uint esFlags);
'@
	$ste = Add-Type -memberDefinition $code -name System -namespace Win32 -passThru 
	$ste::SetThreadExecutionState([uint32]"0x80000000")

}

function get_base_file_name ( [string]$filename ) {

	return [System.IO.Path]::GetFileNameWithoutExtension($filename)

}

function get_extension ( [string]$filename ) {

	return [System.IO.Path]::GetExtension($filename)

}

function get_path_to_file ( [string]$filename ) {

	return [System.IO.Path]::GetDirectoryName($filename)

}

function get_temp_file_name ( [string]$filename ) {

	[string]$tempFileName = ""

	if ( $nocopy -eq $true ) {

		$tempFileName = $filename

	} else {

		$tempFileName = $tempFolder + ([System.IO.Path]::GetFileName($filename))

	}

	return $tempFileName
}

function get_video_resolution ( [string]$filename ) {

	### find the resolution of a video source with ffprobe

	# Read all streams with ffprobe
	$rawstreams = (& "$ffprobe" '-hide_banner' '-loglevel' 'fatal' '-show_entries' 'stream="index,codec_type,codec_name,height"' '-print_format' 'json' '-i' "$filename" | ConvertFrom-Json)

	# Create an empty array and parse all streams into it
	$streams = @()

	for ($i = 0; $i -lt $rawstreams.streams.Length; $i++) {

		$line = New-Object -TypeName PSObject
		$line | Add-Member -MemberType Noteproperty -Name 'index' -Value $rawstreams.streams[$i].index
		$line | Add-Member -MemberType Noteproperty -Name 'codec_type' -Value $rawstreams.streams[$i].codec_type
		$line | Add-Member -MemberType Noteproperty -Name 'codec_name' -Value $rawstreams.streams[$i].codec_name
		$line | Add-Member -MemberType Noteproperty -Name 'height' -Value $rawstreams.streams[$i].height
		$streams += $line

	}

	# Search for a video stream and return the resolution height
	$videotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "video" }

	if ($videotrack -ne $null) {

		return [int]($streams[$videotrack[0].index].height)

	} else {

		# If no video stream was found, return -1
		Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
		Write-Host "Es konnte keine Videoauflösung gefunden werden." -ForegroundColor White -BackgroundColor DarkRed

		$fehler = 1
		return -1
	}
}

function append_to_filename ( [string]$filename, [string]$suffix ) {

	return ((get_path_to_file $filename) + "\" + (get_base_file_name $filename) + " - " + $suffix + (get_extension $filename))

}

function file_exists ( [string]$filename ) {

	if (Test-Path $filename) {

		return $true

	} else {

		return $false

	}
}

function get_cut_source_video {

	if ($parts.length -gt 0) {
		return $true
	} else {
		return $false
	}

}

function prepare_input_parameters {

	return "-hide_banner -hwaccel d3d11va -forced_subs_only 1"

}

function build_video_filter ( [string]$source_file ) {

	### Construct the filter command for burning in subtitles from $source_file
	# Get video resolution for correct formatting
	[int]$resolutionY = get_video_resolution $source_file

	$filter_String = ""

	if ( $resolutionY -ne -1 ) {

		$filter_String = "-filter_complex `"[0:v:0]subtitles=original_size="

		if ( $resolutionY -eq 1080 ) {
			$filter_String += "hd1080:force_style='Alignment=2,Fontsize=24,OutlineColour=&H00000000,Outline=1,BackColour=&H80000000,Shadow=1'"
		} elseif ( $resolutionY -eq 720 ) {
			$filter_String += "hd720:force_style='Alignment=2,Fontsize=24,OutlineColour=&H00000000,Outline=1,BackColour=&H80000000,Shadow=1'"
		} elseif ( $resolutionY -eq 480 ) {
			$filter_String += "hd480:force_style='Alignment=2,Fontsize=20,OutlineColour=&H00000000,Outline=1,BackColour=&H80000000,Shadow=1'"
		} else {
			$filter_String += "pal:force_style='Alignment=2,Fontsize=16,OutlineColour=&H00000000,Outline=1,BackColour=&H80000000,Shadow=1'"
		}

		$filter_String += ":filename=" + (($source_file -replace "\\", "\\\\") -replace ":", "\\\:") + "[vout]`""

		return $filter_String

	} else {

		return -1

	}

}



### Main script ###

system_disable_sleep


# Prepare parameters
$input_parameters = prepare_input_parameters

# Check first source file and copy to cache
if ( $file -ne "" ) {

	# Check if file exists
	if ( Test-Path $file ) {

		# Set target file full path
		$target_File = $tempFolder + (get_base_file_name $file) + ".final.mkv"
		$source_File_1 = get_temp_file_name $file
		$input_parameters += " -i `"" + $source_File_1 + "`""
		if ( $nocopy -eq $false ) {

			Copy-Item "$file" -Destination "$tempFolder"
			if ( !(Test-Path $source_File_1 )) {

				Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
				Write-Host "1. Quelldatei konnte nicht in den Zwischenspeicher kopiert werden." -ForegroundColor White -BackgroundColor DarkRed
				$fehler = 1

			}
		}

		# Prepare filter graph for subtitles and map filter
		if ( $hard_subtitles -eq $true ) {

			$filter_parameters = build_video_filter $source_File_1
			$output_parameters = " -map [vout]"
			if ( $filter_parameters -eq -1 ) {

				$fehler = 1

			}

		# Without filter, directly map first video
		} else {

			$output_parameters = " -map 0:v:0"

		}

		# Check, if 2. source shalll be used and copy to cache, then map 1. and 2. file for audio
		if ( $seccond_audio -eq $true ) {

			if ( Test-Path (append_to_filename $file $suffix) ) {
				$source_File_2 = get_temp_file_name (append_to_filename $file $suffix)
				$input_parameters += " -i `"" + $source_File_2 + "`""
				$output_parameters += " -map 0:a:0 -map 1:a:0"
				if ( $nocopy -eq $false ) {

					Copy-Item (append_to_filename $file $suffix) -Destination "$tempFolder"
					if ( !(Test-Path (append_to_filename $file $suffix)) ) {

						Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
						Write-Host "2. Quelldatei konnte nicht in den Zwischenspeicher kopiert werden." -ForegroundColor White -BackgroundColor DarkRed
						$fehler = 1

					}

				}

			} else {

				Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
				Write-Host "Auf 2. Quelldatei kann nicht zugegriffen werden." -ForegroundColor White -BackgroundColor DarkRed
				$fehler = 1

			}

		# else just map 1. file for audio
		} else {

			$output_parameters += " -map 0:a:0"

		}

		if ( get_cut_source_video -eq $true ) {

			$output_parameters += " -c:v libx264 -crf 0 -preset:v ultrafast -c:a flac -f matroska `"" + (append_to_filename $target_File "lossless") + "`""
			# Find number of channels and set up mapping
			################### TBD ####################

		} else {

			$output_parameters += " -c:v libx265 -crf " + $crf + " -preset:v " + $video_encoding_preset + " -c:a copy -f matroska `"" + $target_File + "`""

		}
	} else {

	Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
	Write-Host "Auf Quelldatei kann nicht zugegriffen werden." -ForegroundColor White -BackgroundColor DarkRed
	$fehler = 1

	}

} else {

	Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
	Write-Host "Keine Quelldatei angegeben." -ForegroundColor White -BackgroundColor DarkRed
	$fehler = 1

}

if ( $fehler -eq 0 ) {

	# Convert!
	Write-Host
	Write-Host "$ffmpeg" $input_parameters $filter_parameters $output_parameters -ForegroundColor Yellow
	Write-Host
	Write-Host ""
	Invoke-Expression "$ffmpeg --% $input_parameters $filter_parameters $output_parameters"

	if ( Test-Path $target_File ) {

		Copy-Item "$target_File" -Destination "$targetFolder"
		Write-Host "Converted file $target_File moved to $targetFolder" -ForegroundColor Green
		Remove-Item "$target_File"
		if ( !(Test-Path $target_File ) ) {

			Write-Host "Temporary target file successfully removed." -ForegroundColor Green

		}
		Write-Host
		Remove-Item "$source_File_1"
		if ( $seccond_audio -eq $true ) {

			Remove-Item "$source_File_2"

		}
		if ( !(Test-Path $source_File_1 ) -and !(Test-Path $source_File_2) ) {

			Write-Host "Temporary source files successfully removed." -ForegroundColor Green

		}
		Write-Host

	} else {

		Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
		Write-Host "Conversion failed!" -ForegroundColor White -BackgroundColor DarkRed
		Write-Host

	}
}


system_allow_sleep
