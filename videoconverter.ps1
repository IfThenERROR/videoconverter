# Parameter übergeben

param (

#[Parameter(Mandatory=$true)]

[String]$file,

[String]$extension = "mkv",

[int]$fps = 25,

[int]$startframe,

[int]$endframe,

[int]$crf = 20,

[switch]$deint,

[switch]$copyaudio,

[switch]$copysubtitles,

[switch]$hardsubtitles,

[String]$outfile,

[switch]$nocopy,

[switch]$keepfile,

[string]$aspectratio

)

# Verzeichnis für lokale Kopie
[string]$tempFolder = "C:\konv\"

[int]$fehler = 0

[string]$videoMapping = "0:v:0"

# Startzeit aus Startframe berechnen
if ($startframe -gt 0) {

	[string]$starttimeCommand = "-ss"
	[double]$starttime = $startframe / $fps
	[string]$starttimeFormatted = "{0:HH:mm:ss.fff}" -f ([datetime]([timespan]::fromseconds($starttime)).Ticks)

} else {

	[string]$starttimeCommand = ""
	[string]$starttimeFormatted = ""

}

# Endzeit aus Endframe berechnen
if ($endframe -gt 0) {

	[string]$endtimeCommand = "-to"
	[double]$endtime = $endframe / $fps
	[string]$endtimeFormatted = "{0:HH:mm:ss.fff}" -f ([datetime]([timespan]::fromseconds($endtime)).Ticks)

} else {

	[string]$endtimeCommand = ""
	[string]$endtimeFormatted = ""

}

# Filter setzen
[string]$filterCommand = "-filter:v"
if ($deint -eq $true) {

	[string]$filter = "bwdif,fps=25"

} else {

	[string]$filter = "fps=25"

}

# Audiospur finden
$rawstreams = c:\"portable apps\ffmpeg\bin\ffprobe.exe" -hide_banner -show_entries stream="index,codec_type,codec_name,channels,bit_rate:stream_tags=language" -print_format json -i $file | ConvertFrom-Json
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

# dts behalten
$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.codec_name -eq "dts" }

if ($audiotrack -ne $null) {

	$copyaudio = $true

} else {

	# aac behalten
	$audiotrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "audio" -and $_.codec_name -eq "aac" }

	if ($audiotrack -ne $null) {

		$copyaudio = $true

	} else {

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

								$fehler = 1

							}

						}

					}

				}

			}

		}

	}

}

if ( $copyaudio -eq $true ) {

	$audioMapping = "0:a"
	$audioCommand = "copy"
	$channelsCommand = ""
	$channels = ""
	$bitrateCommand = ""
	$bitrate = ""

} else {

	$audiomapping = "0:" + $($audiotrack[0].index)
	$audiocommand = "ac3"
	$channelsCommand = "-ac"
	$bitrateCommand = "-b:a"


	if ( $bitrate -ge 384000 ) {

		if ( [int]($streams[$audiotrack[0].index].channels) -eq 6) {

			$bitrate = 384000

		} else {

			$bitrate = 256000

		}

	}

}

# Prüfen, ob Untertitel kopiert werden sollen
$subtitleTrack = $streams | Where-Object -FilterScript { $_.codec_type -eq "subtitle" }

if ($subtitleTrack -ne $null) {

	#Prüfer ob Untertitel eingebrannt werden sollen
	if ( $hardsubtitles -eq $true ) {

		$filterCommand = "-filter_complex"
		$filter = "[0:v:0]$filter[v]; [v][0:s:0]overlay[out]"
		$videoMapping = "[out]"

	} elseif ( $copysubtitles -eq $true ) {

		[string]$subtitleMappingCommand = "-map"
		[string]$subtitleMapping = "0:s"
		[string]$subtitleCodecCommand = "-c:s"
		[string]$subtitleCodec = "copy"

	} else {

		[string]$subtitleMappingCommand = ""
		[string]$subtitleMapping = ""
		[string]$subtitleCodecCommand = ""
		[string]$subtitleCodec = ""

	}

}

# Prüfen, ob aspect ratio vorgegeben ist
if ( !($aspectratio -eq "" ) {

	[string]$arCommand = "-aspect"

} else {

	[string]$arCommand = ""

}

# Dateiname bestimmen
if ( $nocopy -eq $false ) {

	[string]$tempFile = $tempFolder + ([System.IO.Path]::GetFileName($file))

} else {

	# Arbeitsverzeichnis ist Speicherort der Quelldatei
	[string]$tempFile = $file

}

if ( $outfile -eq "" ) {

	$outfile = $tempFolder + ([System.IO.Path]::GetFileNameWithoutExtension($file)) + ".konv.mkv"

}

else{

	$outfile = $tempFolder + $outfile + ".konv.mkv"

}

# Konvertieren
if ( $fehler -eq 0 ) {

	if ( $nocopy -eq $false ) {
		if (!(Test-Path $tempfile)) { Copy-Item "$file" -Destination "$tempfile" }
	}
	& c:\"portable apps\ffmpeg\bin\ffmpeg.exe" "-hide_banner", "-hwaccel", "dxva2", "$starttimeCommand", "$starttimeFormatted", "$endtimeCommand", "$endtimeFormatted", "-n", "-i", "$tempfile", "$arCommand", "$aspectratio", "$filterCommand", "$filter", "-map", "$videoMapping", "-c:v:0", "libx265", "-preset:v:0", "slow", "-crf", "$crf", "-map", "$audiomapping", "-c:a", "$audiocommand", "$channelsCommand", "$channels", "$bitrateCommand", "$bitrate", "$subtitleMappingCommand", "$subtitleMapping", "$subtitleCodecCommand", "$subtitleCodec", "-f", "matroska", "-r", "25", "$outfile"
	if ( !$keepfile ) { Remove-Item "$tempfile" }

}


Write-Host c:\"portable apps\ffmpeg\bin\ffmpeg.exe" "-hide_banner", "-hwaccel", "dxva2", "$starttimeCommand", "$starttimeFormatted", "$endtimeCommand", "$endtimeFormatted", "-n", "-i", "$tempfile", "$arCommand", "$aspectratio", "$filterCommand", "$filter", "-map", "$videoMapping", "-c:v:0", "libx265", "-preset:v:0", "slow", "-crf", "$crf", "-map", "$audiomapping", "-c:a", "$audiocommand", "$channelsCommand", "$channels", "$bitrateCommand", "$bitrate", "$subtitleMappingCommand", "$subtitleMapping", "$subtitleCodecCommand", "$subtitleCodec", "-f", "matroska", "-r", "25", "$outfile"
