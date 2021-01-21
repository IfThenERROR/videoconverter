## Video merger without reencoding (but sometimes with small glitches at transitions)
## v 1.0
## Powershell script

# Parameter übergeben

param (

[String[]]$parts,

[String]$outfile

)

# Directory to download local copy of source file to
[string]$tempFolder = "W:\"

# Directory to store final file into
[string]$targetFolder = "X:\recorder\konv\"

[int]$fehler = 0

if ( $outfile -eq "" ) {

	$outfile = $tempFolder + ([System.IO.Path]::GetFileNameWithoutExtension($parts[0])) + ".merged"

}

else{

	$outfile = $tempFolder + $outfile

}

foreach ($filename in $parts) {
	[String]$MyFile = ($outfile + ".txt")
	[IO.File]::AppendAllLines(($outfile + ".txt"), [string[]]("file '" + $filename + "'"))
}

c:\"portable apps\ffmpeg\bin\ffmpeg.exe" -hide_banner -f concat -safe 0 -i ($outfile + ".txt") -c copy ($outfile + ".mkv")
Copy-Item ($outfile + ".mkv") -Destination "$targetFolder"
Remove-Item ($outfile + ".txt")
Remove-Item ($outfile + ".mkv")
