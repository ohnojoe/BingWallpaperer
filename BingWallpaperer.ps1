#   Updates Windows desktop wallpaper from Bing's Image of the Day web service
#   ###########################################################################
#   Created by:     Joe Blumenow (2014)
#   Last Updated:   2019-03-10
#   #########################################


#   Parameters
#   #########################################
#   $size - ideal image dimensions in WxH format, not all dimensions supported
#   $idx  - image index, defaults to 0 (latest image), max value is 7 (at time of writing)
#   $mkt  - the market/region of images to uses
#   $savePath - local save path of the image. Defaults to My Pictures folder if left blank.
param (
    [string]$size = "",
    [string]$idx = "0",
    [string]$mkt = "en-GB",
    [string]$savePath = "",
    [switch]$RegisterSchedule = $false,
    [switch]$UnregisterSchedule = $false,
    
    [ValidateSet("NoChange", "Center", "Tile", "Stretch", "Fit", "Fill")]
    [string]$wallpaperStyle = "NoChange"

)
 

#   .NET class to correctly set wallpaper
#   Adapted from: https://social.technet.microsoft.com/Forums/en-US/9af1769e-197f-4ef3-933f-83cb8f065afb/background-change
#   #########################################
Add-Type @"
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
namespace Wallpaper
{
    public enum Style : int {
        NoChange, Center, Tile, Stretch, Fit, Fill
    }
    public class Setter {
        public const int SetDesktopWallpaper = 20;
        public const int UpdateIniFile = 0x01;
        public const int SendWinIniChange = 0x02;
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo (int uAction, int uParam, string lpvParam, int fuWinIni);
        public static void SetWallpaper ( string path, Wallpaper.Style style ) {	  	  
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
            switch (style) {
                case Style.Center:
                    key.SetValue(@"WallpaperStyle", "1"); 
                    key.SetValue(@"TileWallpaper", "0"); 
                    break;
                case Style.Tile:
                    key.SetValue(@"WallpaperStyle", "1"); 
                    key.SetValue(@"TileWallpaper", "1");
                    break;
                case Style.Stretch:
                    key.SetValue(@"WallpaperStyle", "2"); 
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case Style.Fit:
                    key.SetValue(@"WallpaperStyle", "6"); 
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case Style.Fill:
                    key.SetValue(@"WallpaperStyle", "10"); 
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case Style.NoChange:
                    break;
            }
            key.Close();

            // Set wallpaper after style change to ensure the change is picked up
            SystemParametersInfo( SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange );
        }
    }
}
"@


#   Get an array of possible image dimensions
#   with the ideal size first in the array
#   #########################################
Function Get-IdealImageDimensionsArray($idealSize) {
    $validImageDimensionArray = New-Object System.Collections.ArrayList
    
    # This should be a list of all known possible
    # image dimensions.
    # There are likely many more than this, but
    # this should cover the main ones.

    $validImageDimensionArray.Add("1920x1200") | Out-Null
    $validImageDimensionArray.Add("1920x1080") | Out-Null   # 1920x1080 tends to exclude bing branding
    $validImageDimensionArray.Add("1366x768") | Out-Null
    $validImageDimensionArray.Add("1024x768") | Out-Null

    # Remove the ideal size, in case it's in the list
    $validImageDimensionArray.Remove($idealSize)

    # Add the ideal size as the first
    $validImageDimensionArray.Insert(0, $idealSize)

    $validImageDimensionArray
}


#   Downloads the specified image from
#   #########################################
Function Get-Image($imageSize, $idx, $mkt) {

    # Base Bing url, comes in handy later
    $urlBing = "http://www.bing.com"

    # Bing Image feed containing xml of specified image (always defaults to latest if idx or mkt params not valid)
    $urlBingImageFeed = "{0}/HPImageArchive.aspx?format=xml&idx={1}&n=1&mkt={2}" -f $urlBing, $idx, $mkt

    $urlImageBasePath = ""
    $urlImage = ""
    $savelocation = ""

    Write-Debug ("Feed URL: " + $urlBingImageFeed)

    # Initialise WebClient, proxy
    $webClient = New-Object System.Net.WebClient

    # Get image base path
    $page = $webClient.DownloadString($urlBingImageFeed)
    $regex = [regex] '<urlBase>(.*)</urlBase>'
    $match = $regex.Match($page)
    $urlImageBasePath = $match.Groups[1].Value

    Write-Debug ("Image base path: $urlImageBasePath") -Verbose

    # Build  image url
    $urlImage = "{0}{1}_{2}.jpg" -f $urlBing, $urlImageBasePath, $imageSize
    
    Write-Debug "Image URL: $urlImage"

    if ($savePath -eq "") {
        $myPicturesFolder = [Environment]::GetFolderPath("MyPictures")
    } else {
        $myPicturesFolder = $savePath
    }

    $savelocation = [io.path]::combine($myPicturesFolder, 'bingimageoftheday.jpg')

    Try {
        Write-Debug ("Downloading image: " + $urlImage) -Verbose

        $webClient.DownloadFile($urlImage, $savelocation)

        ## Assume something is wrong if file doesn't exist or length is less than a 1KB
        if (!(Test-Path $savelocation) -or ((Get-Item $savelocation).length -lt 1kb)) { 
            Write-Warning "There was a problem downloading the image $urlImage"
            $savelocation = ""
        }
    }
    Catch [System.Net.WebException] {
        ## Assume it fails due to 404
        Write-Warning "Unable to download image $urlImage"
    
        $savelocation = ""
    }

    return $savelocation
}


#   Gets the current screen DPI factor
#   #########################################
function Get-DpiFactor {
    $DPISetting = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI).AppliedDPI
    switch ($DPISetting)
    {
        96 {return 1}
        120 {return 1.25}
        144 {return 1.5}
        192 {return 2}
    }
    return 1
}

#   Gets the current screen resolution for
#   the primary monitor
#   #########################################
function Get-ScreenResolution {       
    [void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $s = [System.Windows.Forms.SystemInformation]::PrimaryMonitorSize
    $dpi = Get-DpiFactor
    [decimal]$w = [math]::Floor($s.Width * $dpi)
    [decimal]$h = [math]::Floor($s.Height * $dpi)    
    return "${w}x${h}"
}


function Register-Schedule($taskPath, $taskName) { 
    # Make sure task doesn't already exist
    Unregister-Schedule $taskPath $taskName

    $argument = $script:MyInvocation.MyCommand.Path
    if ($size -ne "") { $argument += " -size $size" }
    if ($idx -ne 0) { $argument += " -idx $idx" }
    if ($mkt -ne "en-GB") { $argument += " -mkt $mkt" }
    if ($wallpaperStyle -ne "NoChange") { $argument += " -wallpaperStyle $wallpaperStyle" }
    $trigger = New-ScheduledTaskTrigger -At 00:10 -Daily
    $action = New-ScheduledTaskAction -Execute "Powershell.exe" -Argument $argument
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable
    $task = Register-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings
    Write-Output "    Scheduled task registered."
}


function Unregister-Schedule($taskPath, $taskName) {
    $existingTask = Get-ScheduledTask | where { $_.TaskName -eq $taskName -and $_.TaskPath -eq $taskPath}
    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Output "    Scheduled task unregistered"
    } else {
        Write-Output "    Scheduled task doesn't currently exist"
    }
}


#   Main Program Entry Point
#   #########################################

$taskPath = "\"
$taskName = "Bing Wallpaperer Daily Update"

Write-Output ""
Write-Output ">>> Bing Image of the Day Wallpaper Updater <<<"

if ($UnregisterSchedule) {
    Write-Output "... Unregistering daily schedule task"
    Unregister-Schedule $taskPath $taskName
}

if ($RegisterSchedule) {
    Write-Output "... Registering daily schedule task"
    Register-Schedule $taskPath $taskName
}


if ($size -eq "") {
    $size = Get-ScreenResolution
}


# Get a list of the main image sizes, with the
# specified or ideal size as the first item
$imageSizes = Get-IdealImageDimensionsArray $size

# Loop through all our sizes here because not
# all sizes are available for all images
# particularly HD images, so we're doing it this
# way in case the specified image isn't available
# it will at least fallback to an alternative.
foreach ($imageSize in $imageSizes) {

    # Let's try and download the image
    Write-Output "... Trying to get image - size: $imageSize, idx: $idx, mkt: $mkt"
    $saveLocation = Get-Image $imageSize $idx $mkt

    if ($saveLocation -ne "") {
        # All's good, an image was saved...
        Write-Output "... Image successfully saved to: $saveLocation" 

        # Set the image as the desktop wallpaper
        [Wallpaper.Setter]::SetWallpaper($saveLocation, $wallpaperStyle)         
        Write-Output "... Wallpaper set with style of $wallpaperStyle"
        
        # We can break out of our $imageSizes loop now
        break
    }
    
    # No image saved, there must have been a problem
    Write-Warning "Unable to download image at this size: $imageSize"
}

Write-Output ">>> All done, goodbye! <<<"