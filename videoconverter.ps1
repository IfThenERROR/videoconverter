## Videoconverter
## v 2.1
## Powershell script

param (

[Parameter(Mandatory=$true)][String]$file,

[String]$extension = "mkv",

[int]$fps = 25,

[int]$crf = 20,

[switch]$lossless,

[switch]$deint,

[switch]$copyaudio,

[switch]$hightempcompression,

[string]$audiocodec,

[String]$outfile,

[switch]$nocopy,

[switch]$removefile,

[switch]$removetemp,

[switch]$deletechapters,

[switch]$finalize = $false,

[string]$aspectratio = "16:9",

[single]$gainaudio = 1.0,

[int[]]$parts

)

### General definitions ###
# Directory to download local copy of source file to
[string]$tempFolder = "W:\"

# Directory to store final file into
[string]$targetFolder = "X:\recorder\konv\"

# Path to ffprobe
[string]$ffprobe = "C:\portable apps\ffmpeg\bin\ffprobe.exe"

# Path to ffmpeg
[string]$ffmpeg = "C:\`"portable apps\ffmpeg\bin\ffmpeg.exe`""

# Path to leading black video file for finalizing
[string]$finalizePreStereo = "C:\konv\black sekunde 169 stereo.lossy.mkv"
[string]$finalizePre51 = "C:\konv\black sekunde 169 5.1.lossy.mkv"
[string]$finalizePre43 = "C:\konv\black sekunde 43 stereo.lossy.mkv"


### Variables ###
[int]$fehler = 0
[string]$tempFile = ""
[string]$audioTempFiles = ""
[string]$extension = ""
[string]$videoMapping = ""
[string]$quellenCommand = ""
[int]$segmente = 0
[string[]]$startzeiten
[string[]]$endzeiten
[string]$quellen
[string]$videoFilterCommand = ""
[string]$videoFilter = ""
[string]$audioFilterCommand = ""
[string]$audioFilter = ""
[string]$filters = ""
[string]$concatFilter = ""
[string]$concatStreams = ""
[int]$concatAnzahl = 0
[string]$channelsCommand = ""
[string]$bitrateCommand = ""
[string]$chapCommand = ""
[string]$chapMapping = ""
[string]$parameters


### Functions ###
function get-FormatedTime-from-Frame([int]$frame, [int]$FramesPerSec) {

	([double]$timeCalculated = $frame / $FramesPerSec) | Out-Null
	([string]$timeFormatted = "{0:HH:mm:ss.fff}" -f ([datetime]([timespan]::fromseconds($timeCalculated)).Ticks)) | Out-Null
	return $timeFormatted

}


### Main script ###
#Request system availabiliry, no sleep idle timeout while executing
$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
public static extern void SetThreadExecutionState(uint esFlags);
'@
$ste = Add-Type -memberDefinition $code -name System -namespace Win32 -passThru
$ste::SetThreadExecutionState([uint32]"0x80000000" -bor [uint32]"0x00000001")


# If finalize, than also remove temporary file
if ($finalize -eq $true) {

	$removetemp = $true

}

# Set working directory
if ( $nocopy -eq $false ) {

	$tempFile = $tempFolder + ([System.IO.Path]::GetFileName($file))

} else {

	# Working directory is same as source file
	$tempFile = $file

}

# Set up video filters
$videoFilterCommand = "-vf"
if ($deint -eq $true) {

	$filter = "bwdif,fps=25"

} else {

	$filter = "fps=25"

}

# Set up audio filters
if (!($gainaudio -eq 1.0)) {

	$audioFilterCommand = "-filter:a"
	$audioFilter = "volume=" + $gainaudio

} else {

	$audioFilterCommand = ""
	$audioFilter = ""

}

# Check if the source file exists, then start converting
if ( Test-Path $file) {

	# Check if parts were defined
	if ($parts.length -gt 0) {

		if ($parts.length % 2 -ne 0) {

			Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
			Write-Host "Bei Abschnitten fehlt eine Start- oder Endzeit." -ForegroundColor White -BackgroundColor DarkRed
			$fehler = 1
	
		} else {

			# Find best audio stream
			$rawstreams = (& "$ffprobe" '-hide_banner' '-loglevel' 'fatal' '-show_entries' 'stream="index,codec_type,codec_name,channels,bit_rate:stream_tags=language"' '-print_format' 'json' '-i' "$file" | ConvertFrom-Json)
			$streams = @()

			for ($i = 0; $i -lt $rawstreams.streams.Length; $i++) {

				$line = New-Object -TypeName PSObject
				$line | Add-Member -MemberType Noteproperty -Name 'index' -Value $rawstreams.streams[$i].index
				$line | Add-Member -MemberType Noteproperty -Name 'codec_type' -Value $rawstreams.streams[$i].codec_type
				$line | Add-Member -MemberType Noteproperty -Name 'codec_name' -Value $rawstreams.streams[$i].codec_name
				$line | Add-Member -MemberType Noteproperty -Name 'channels' -Value $rawstreams.streams[$i].channels
				$line | Add-Member -MemberType Noteproperty -Name 'bit_rate' -Value $rawstreams.streams[$i].bit_rate
				$line | Add-Member -MemberType Noteproperty -Name 'language' -Value $rawstreams.streams[$i].tags.language
				$streams += $line

			}

			# 5.1 + deu
			$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.channels -eq "6" -and $_.language -like "deu" }

			if ($audiotrack -ne $null) {

				$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
				$channels = 6

			} else {

				# 5.1 + ger
				$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.channels -eq "6" -and $_.language -like "ger" }

				if ($audiotrack -ne $null) {

					$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
					$channels = 6

				} else {

					# ac3 + deu
					$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.codec_name -eq "ac3" -and $_.language -like "deu" }

					if ($audiotrack -ne $null) {

						$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
						$channels = 2

					} else {

						# ac3 + ger
						$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.codec_name -eq "ac3" -and $_.language -like "ger" }

						if ($audiotrack -ne $null) {

							$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
							$channels = 2

						} else {

							# deu
							$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.language -like "deu" }

							if ($audiotrack -ne $null) {

								$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
								$channels = 2

							} else {

								# ger
								$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.language -like "ger" }

								if ($audiotrack -ne $null) {

									$bitrate = [int]($streams[$audiotrack[0].index].bit_rate)
									$channels = 2

								} else {

									Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
									Write-Host "Keine passende Audiospur gefunden" -ForegroundColor White -BackgroundColor DarkRed
									$fehler = 1

								}

							}

						}

					}

				}

			}

			# Set up parameters for audio codec
			# Default is ac3 for 5.1, aac for stereo
			if ( $audiocodec -eq "" ) {

				if ( [int]($streams[$audiotrack[0].index].channels) -eq 6 ) {

					$audiocodec = "ac3"
					$channelsCommand = "-ac"
					$bitrateCommand = "-b:a"
					#Limit bitrate to 448k for 5.1
					if ( $bitrate -ge 448000 ) { $bitrate = 448000 }

				} else {

					$audiocodec = "aac"
					$channelsCommand = "-ac"
					$bitrateCommand = "-b:a"
					#Limit bitrate to 128k for stereo
					$bitrate = 128000

				}

			} elseif ( $audiocodec -eq "" ) {

				$channelsCommand = "-ac"
				$bitrate = ""

			}

			# Set up audio mapping
			$audiomapping = "0:" + $($audiotrack[0].index)

			# Check, if aspect ratio is defined
			if ( !($aspectratio -eq "" ) {

				[string]$arCommand = "-aspect"

			} else {

				[string]$arCommand = ""

			}

			# Check, if chapters shall be deleted
			if ( $deletechapters -eq $true ) {

				[string]$chapCommand = "-map_chapters"
				[string]$chapMapping = "-1"

			}

			# If no parameter for outfile is entered, create a file with same name in temp directory, else use the set namen
			if ( $outfile -eq "" ) {

				$outfile = ([System.IO.Path]::GetFileNameWithoutExtension($file))

			} else {

				$outfile = $outfile

			}

			# Prepare filters
			$segmente = $parts.length / 2
			$filterCommand = "-vf"

			# Copy the source file into the temporary folder if not disabled
			if ( $nocopy -eq $false ) {

				if (!(Test-Path $tempfile)) {
					Write-Host Copying source to temporary folder
					Copy-Item "$file" -Destination "$tempfile"
				}

			}

			$quellen = "-i `"" + $tempFile + "`""


			# Directly encode, if not cutting more than one part
			if (( $segmente -eq 1 ) -and ( $finalize -eq $false )) {

				$extension = ".final.mkv"
				$outfile = $tempFolder + $outfile + $extension

				# Compose parameters in a string
				if ( $lossless -eq $true ) {

					$parameters = "-hide_banner -hwaccel d3d11va -ss " + (get-FormatedTime-from-Frame -frame ($parts[0] + 1) -FramesPerSec $fps) + " $quellen  -to " + (get-FormatedTime-from-Frame -frame ($parts[1] - $parts[0] - 2) -FramesPerSec $fps) + " $arCommand $aspectratio $videoFilterCommand $videoFilter $audioFilterCommand $audioFilter -map 0:v:0 -c:v:0 libx265 -preset:v:0 slow -x265-params 'lossless=1' -map $audiomapping -c:a flac $channelsCommand $channels $chapCommand $chapMapping -f matroska -r 25 `"$outfile`""

				} else {

					$parameters = "-hide_banner -hwaccel d3d11va -ss " + (get-FormatedTime-from-Frame -frame ($parts[0] + 1) -FramesPerSec $fps) + " $quellen  -to " + (get-FormatedTime-from-Frame -frame ($parts[1] - $parts[0] - 2) -FramesPerSec $fps) + " $arCommand $aspectratio $videoFilterCommand $videoFilter $audioFilterCommand $audioFilter -map 0:v:0 -c:v:0 libx265 -preset:v:0 slow -crf $crf -map $audiomapping -c:a $audiocodec $channelsCommand $channels $bitrateCommand $bitrate $chapCommand $chapMapping -f matroska -r 25 `"$outfile`""

				}

				# Show the ffmpeg command that will be applied
				Write-Host "$ffmpeg" $parameters -ForegroundColor Yellow
				Write-Host ""

				# Launch ffmpeg
				Invoke-Expression "$ffmpeg --% $parameters"

				# Delete temp file if requested
				if ( ($removetemp) -or ($removefile) ) { Remove-Item "$tempfile" }

			# For more than one part, cache parts first, than merge
			} else {

				# Loop through all frame pairs and cut segments
				for ( $i = 1; $i -le $segmente; $i++ ) {

					# Compose parameters in a string
					if ( $hightempcompression -eq $true ) {

						$parameters = "-hide_banner -hwaccel d3d11va -ss " + (get-FormatedTime-from-Frame -frame $parts[(2 * $i) - 2] -FramesPerSec $fps) + " $quellen -to " + (get-FormatedTime-from-Frame -frame ($parts[(2 * $i) - 1] - $parts[(2 * $i) - 2] - 1) -FramesPerSec $fps) + " $arCommand $aspectratio $videoFilterCommand $videoFilter $audioFilterCommand $audioFilter -map 0:v:0 -c:v:0 libx265 -preset:v:0 medium -x265-params 'lossless=1' -map $audiomapping -c:a flac $channelsCommand $channels $chapCommand $chapMapping -f matroska -r 25 `"$tempFolder$outfile`.$i`.konv.mkv`""

					} else {

						$parameters = "-hide_banner -hwaccel d3d11va -ss " + (get-FormatedTime-from-Frame -frame $parts[(2 * $i) - 2] -FramesPerSec $fps) + " $quellen -to " + (get-FormatedTime-from-Frame -frame ($parts[(2 * $i) - 1] - $parts[(2 * $i) - 2] - 1) -FramesPerSec $fps) + " $arCommand $aspectratio $filterCommand $filter -map 0:v:0 -c:v:0 libx264 -preset:v:0 superfast -crf 0 -map $audiomapping -c:a flac $channelsCommand $channels $chapCommand $chapMapping -f matroska -r 25 `"$tempFolder$outfile`.$i`.konv.mkv`""

					}

					# Show the ffmpeg command that will be applied
					Write-Host "$ffmpeg" $parameters -ForegroundColor Yellow
					Write-Host ""

					# Launch ffmpeg
					Invoke-Expression "$ffmpeg --% $parameters"

				}

				# Prepare file name. Append ".konv", if the video is just transcoded, append ".final", if the video is finalized
				if ( $finalize ) {

					$extension = ".final.mkv"

				} else {

					$extension = ".konv.mkv"

				}

				# Merge all segments and encode

				# Remove temporary source to free space
				if ( ($removetemp) -or ($removefile) ) { Remove-Item "$tempfile" }

				# Set up sources
				# prepand blank segment if finalizing
				if ($finalize -eq $true) {

					if ($aspectratio -eq "4:3" ) {

						$quellen = "-i `"$finalizePre43`""

					} elseif ($channels -eq 6) {

						$quellen = "-i `"$finalizePre51`""

					} elseif ($channels -eq 2) {

						$quellen = "-i `"$finalizePreStereo`""

					} else {

						# Kann nicht finalisieren, da kein Intro mit passenden Kanälen existiert
						Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
						Write-Host "Unknown setting for audio channels found. Neither 5.1 nor stereo!" -ForegroundColor White -BackgroundColor DarkRed
						Write-Host
						$fehler = 1

					}

					$concatAnzahl = 1

				}

				# Add all segments as source
				for ( $i = 1; $i -le $segmente; $i++ ) {

					$quellen = "$quellen -i `"$tempFolder$outfile`.$i`.konv.mkv`""

				}

				$concatAnzahl = $concatAnzahl + $segmente

				# Set up concat filter
				for ( $i = 0; $i -lt $concatAnzahl; $i++ ) {

					$concatStreams = "$concatStreams" + "[" + "$i" + ":v:0][" + "$i" + ":a:0]"

				}

				###############concat filter vorbereiten
				$concatFilter = "$concatStreams" + "concat=n=" + "$concatAnzahl" + ":v=1:a=1[videoOut][audioOut]"

				# Compose parameters in a string
				$parameters = "-hide_banner -hwaccel d3d11va $quellen -filter_complex `"$concatFilter`" -map [videoOut] -c:v:0 libx265 -preset:v:0 slow -crf $crf -map [audioOut] -c:a $audiocodec $channelsCommand $channels $bitrateCommand $bitrate $chapCommand $chapMapping -f matroska -r 25 `"$tempFolder$outfile$extension`""

				# Show the ffmpeg command that will be applied
				Write-Host "$ffmpeg" $parameters -ForegroundColor Yellow
				Write-Host ""

				# Launch ffmpeg
				Invoke-Expression "$ffmpeg --% $parameters"


				# Clear temp files
				for ( $i = 1; $i -le $segmente; $i++ ) {

					Remove-Item ( "$tempFolder" + "$outfile" + "`.$i`.konv.mkv")

				}
				$outfile = $tempFolder + $outfile + $extension

			}


		}

		# Convert
		if ( $fehler -eq 0 ) {

			# Move converted file to final directory and remove local copy
			if (Test-Path $outfile) {

				Copy-Item "$outfile" -Destination "$targetFolder"
				Write-Host "Converted file $outfile moved to $targetFolder" -ForegroundColor Green
				Remove-Item "$outfile"
				Write-Host "Temporary file successfully removed."
				Write-Host

				# Delete source file if requested
				if ( $removefile ) { Remove-Item "$file" }

			} else {

				Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
				Write-Host "Conversion failed!" -ForegroundColor White -BackgroundColor DarkRed
				Write-Host

			}

		}

	} else {

		Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
		Write-Host "Start- und Endzeit fehlen." -ForegroundColor White -BackgroundColor DarkRed
		$fehler = 1

	}

} else {

	Write-Host "ERROR!" -ForegroundColor White -BackgroundColor DarkRed
	Write-Host "Source file not found!" -ForegroundColor White -BackgroundColor DarkRed
	Write-Host
	$fehler = 1

}

#Return to default sleep timeout
$ste::SetThreadExecutionState([uint32]"0x80000000")
