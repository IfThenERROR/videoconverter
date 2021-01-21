## Video merger seemless with reencoding
## v 1.0
## Powershell script

# Parameter übergeben

param (

[String[]]$parts,

[int]$channels = 6,

[String]$outfile

)

# Directory to download local copy of source file to
[string]$tempFolder = "W:\"

# Directory to store final file into
[string]$targetFolder = "X:\recorder\konv\"

[string]$ffmpeg = "C:\`"portable apps\ffmpeg\bin\ffmpeg.exe`""
[int]$fehler = 0
[int]$counter = 0
[string]$inputString = ""
[string]$filterComplexString = ""
[string]$audiocodec = "ac3"
[string]$audiobitrate = "-b:a 384000"
[string]$parameters = ""

# Request system availability, no sleep idle timeout while executing
$code=@' 
[DllImport("kernel32.dll", CharSet = CharSet.Auto,SetLastError = true)]
  public static extern void SetThreadExecutionState(uint esFlags);
'@
$ste = Add-Type -memberDefinition $code -name System -namespace Win32 -passThru 
$ste::SetThreadExecutionState([uint32]"0x80000000" -bor [uint32]"0x00000001")



if ( $outfile -eq "" ) {

	$outfile = $tempFolder + ([System.IO.Path]::GetFileNameWithoutExtension($parts[0])) + ".merged"

}

else{

	$outfile = $tempFolder + $outfile

}

foreach ($filename in $parts) {

	Copy-Item "$filename" -Destination "$tempFolder"
	$inputString = $inputString + " -i `"$tempfolder" + (Split-Path $filename -leaf) + "`""
	$filterComplexString = "$filterComplexString" + "[$counter" + ":v:0][$counter" + ":a:0]"
	$counter = $counter + 1

}

$filterComplexString = "$filterComplexString" + "concat=n=" + $counter + ":v=1:a=1[videoOut][audioOut]"
$parameters = "-hide_banner -hwaccel d3d11va $inputString -filter_complex $filterComplexString -map [videoOut] -c:v:0 libx265 -preset:v:0 slow -crf 20 -map [audioOut] -c:a $audiocodec -ac $channels $audiobitrate -map_chapters -1 -f matroska -r 25 `"" + $outfile + ".mkv`""
write-host $parameters
write-host

Write-Host "$ffmpeg --% $parameters" -ForegroundColor Yellow
Invoke-Expression "$ffmpeg --% $parameters"

Copy-Item ($outfile + ".mkv") -Destination "$targetFolder"

foreach ($filename in $parts) {

	Remove-Item ("$tempFolder" + (Split-Path $filename -leaf))

}

Remove-Item ($outfile + ".mkv")

# Return to default sleep timeout
$ste::SetThreadExecutionState([uint32]"0x80000000")
