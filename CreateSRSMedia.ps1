<#
.SYNOPSIS

    Create SRSv2 media appropriate for setting up an SRSv2 device.


.DESCRIPTION

    This script automates some sanity checks and copying operations that are
    necessary to create bootable SRSv2 media. Booting an SRSv2 device using the
    media created from this process will result in the SRSv2 shutting down. The
    SRSv2 can then either be put into service, or booted with separate WinPE
    media for image capture.

    To use this script, you will need:

    1. An Internet connection
    2. A USB drive with sufficient space (16GB+), inserted into this computer
    3. Windows 10 Enterprise or Windows 10 Enterprise IoT media, which must be
       accessible from this computer (you will be prompted for a path). The
       Windows media build number must match the build required by the SRSv2
       deployment kit.
    4. If using Windows 10 Enterprise IoT, an ePKEA (volume license key)

.EXAMPLE
    .\CreateSrsMedia

    Prompt for required information, validate provided inputs, and (if all
    validations pass) create media on the specified USB device.

.NOTES

    This script requires that you provide Windows media targeted for the x64
    architecture.

    Only one driver pack can be used at a time. Each unique supported SKU of
    SRSv2 computer hardware must have its own, separate image.

    The build number of the Windows media being used *must* match the build
    required by the SRSv2 deployment kit.

#>

<#
Revision history
    1.0.0  - Initial release
    1.0.1  - Support source media with WIM >4GB
    1.1.0  - Switch Out-Null to Write-Debug for troubleshooting
             Record transcripts for troubleshooting
             Require the script be run from a path without spaces
             Require the script be run from an NTFS filesystem
             Soft check for sufficient scratch space
             Warn that the target USB drive will be wiped
             Rethrow exceptions after cleanup on main path
    1.2.0  - Indicate where to get Enterprise media
             Improve error handling for non-Enterprise media
             Report and exit on copy errors
             Work with spaces in the script's path
             Explicitly reject Windows 10 Media Creation Tool media
             Fix OEM media regression caused by splitting WIMs
    1.3.1  - Read config information from MSI
             Added infrastructure for downloading files
             Support for automatically downloading Windows updates
             Support for automatically downloading the deployment kit MSI
             Support for self-updating
             Added menu-driven driver selection/downloading
    1.3.2  - Fix OEM media regression caused by splitting WIMs
    1.4.0  - Support BIOS booting
    1.4.1  - BIOS booting controlled by metadata
    1.4.2  - Fix driver pack informative output
             Add 64-bit check to prevent 32-bit accidents
             Add debugging cross-check
             Add checks to prevent the script being run in weird ways
             Add warning about image cleanup taking a long time
             Fix space handling in self-update
    1.4.3  - Add non-terminating disk initialization logic
             Delete "system volume information" to prevent Windows Setup issues
             Add return code checking for native commands
    1.4.4  - Improve rejection of non-LP CABs
    1.4.5  - Add host OS check to prevent older DISM etc. mangling newer images
    1.5.0  - Add support for mismatched OS build number vs. feature build number
    1.5.1  - Change OEM default key.

#>
[CmdletBinding()]
param(
    [Switch]$ShowVersion, <# If set, output the script version number and exit. #>
    [Switch]$Manufacturing <# Internal use. #>
)

$ErrorActionPreference = "Stop"
$DebugPreference = if($PSCmdlet.MyInvocation.BoundParameters["Debug"]) { "Continue" } else { "SilentlyContinue" }
Set-StrictMode -Version Latest

$CreateSrsMediaScriptVersion = "1.5.1"

[version]$DevHostOs = "10.0.17763.503"
[version]$HostOs = "0.0.0.0"

$robocopy_success = {$_ -lt 8 -and $_ -ge 0}

if ($ShowVersion) {
    Write-Output $CreateSrsMediaScriptVersion
    exit
}

function Remove-Directory {
  <#
    .SYNOPSIS
        
        Recursively remove a directory and all its children.

    .DESCRIPTION

        Powershell can't handle 260+ character paths, but robocopy can. This
        function allows us to safely remove a directory, even if the files
        inside exceed Powershell's usual 260 character limit.
  #>
param(
    [parameter(Mandatory=$true)]
    [string]$path <# The path to recursively remove #>
)

    # Make an empty reference directory
    $cleanup = Join-Path $PSScriptRoot "empty-temp"
    if (Test-Path $cleanup) {
        Remove-Item -Path $cleanup -Recurse -Force
    }
    New-Item -ItemType Directory $cleanup | Write-Debug

    # Use robocopy to clear out the guts of the victim path
    (Invoke-Native "& robocopy '$cleanup' '$path' /mir" $robocopy_success) | Write-Debug

    # Remove the folders, now that they're empty.
    Remove-Item $path -Force
    Remove-Item $cleanup -Force
}

function Test-OsPath {
  <#
    .SYNOPSIS
        
        Test if $OsPath contains valid Windows setup media for SRSv2.

    .DESCRIPTION

        Tests if the provided path leads to Windows setup media that meets
        requirements for SRSv2 installation. Specifically, the media must:

          - Appear to have the normal file layout for Windows media
          - Be Enterprise (not Pro, Home, etc.)
          - Be Enterprise IoT (not plain Enterprise)
          - Be targeted for x64
          - Match the version number the SRSv2 deployment kit requires

    .OUTPUTS bool

        $true if $OsPath refers to valid Windows installation media, $false otherwise.
  #>
param(
  [parameter(Mandatory=$true)]
  $OsPath, <# The path that to verify contains appropriate Windows installation media. #>
  [parameter(Mandatory=$true)]
  $KitOsRequired, <# The Windows version the deployment kit requires. #>
  [parameter(Mandatory=$false)]
  [switch]$IsOem <# Whether this an OEM customer or (otherwise) Enterprise customer. #>
)

    if (!(Test-Path $OsPath)) {
        Write-Host "The path provided does not exist. Please specify a path that is the root of the Windows installation media you wish to use."
        return $false
    }

    # Save some paths we'll re-use.
    $OsSources = (Join-Path $OsPath "sources")
    $InstallEsd = (Join-Path $OsSources "install.esd")
    $InstallWim = (Join-Path $OsSources "install.wim")

    if (Test-Path($InstallEsd)) {
        Write-Host "This appears to be media generated by the Windows 10 media creation tool, which creates only non-Enterprise media."
        PrintWhereToGetMedia -IsOem:$IsOem
        return $false
    }

    # Basic sanity check -- does this look like Windows install media?
    if (
        (!(Test-Path (Join-Path $OsPath "boot"))) -or
        (!(Test-Path (Join-Path $OsPath "efi"))) -or
        (!(Test-Path $OsSources)) -or
        (!(Test-Path (Join-Path $OsPath "support" ))) -or
        (!(Test-Path (Join-Path $OsPath "autorun.inf"))) -or
        (!(Test-Path (Join-Path $OsPath "bootmgr"))) -or
        (!(Test-Path (Join-Path $OsPath "bootmgr.efi"))) -or
        (!(Test-Path (Join-Path $OsPath "setup.exe"))) -or
        (!(Test-Path $InstallWim))
        ) {
        Write-Host "The path provided does not seem to point to valid Windows installation media. Please specify a path that is the root of the Windows installation media you wish to use."
        PrintWhereToGetMedia -IsOem:$IsOem
        return $false
    }

    $img = $null

    # The source media must have an image named "Windows 10 Enterprise" in its install.wim
    try {
        $img = Get-WindowsImage -ImagePath $InstallWim -Name "Windows 10 Enterprise"
    } catch {
        Write-Host ("The media specified does not appear to contain a Windows 10 Enterprise image.")
        Write-Host ("Double-check that:")
        Write-Host ("  - You are using Windows 10 Enterprise media")
        Write-Host ("  - You have the latest version of this script, and")
        Write-Host ("  - The media is not corrupt.")
        PrintWhereToGetMedia -IsOem:$IsOem
        Write-Host ("")
	    Write-Host ("Images present on this media are:")
        Get-WindowsImage -ImagePath $InstallWim |% {
            Write-Host ("  - {0}" -f $_.ImageName)
        }
        return $false
    }

    # Windows 10 Enterprise has EI.CFG, but Windows 10 Enterprise IoT does not.
    $IsIoT = !(Test-Path (Join-Path $OsSources "EI.CFG"))

    # OEMs need to use IoT.
    if ($IsOem -and !$IsIoT) {
        Write-Host "You appear to have specified a path to Windows 10 Enterprise media. However, you need to use Windows 10 Enterprise IoT media."
        PrintWhereToGetMedia -IsOem:$IsOem
        return $false
    }

    # Non-OEMs need to use normal Windows 10 Enterprise.
    if (!$IsOem -and $IsIoT) {
        Write-Host "You appear to have specified a path to Windows 10 Enterprise IoT media. However, you need to use Windows 10 Enterprise media."
        PrintWhereToGetMedia -IsOem:$IsOem
        return $false
    }

    # We only accept x64-arch media. 9 == x64
    if ($img.Architecture -ne 9) {
        Write-Host ("Your Windows installation media is targeted for the {0} architecture. SRSv2 requires x64 targeted installation media." -f $img.Architecture)
        return $false
    }

    # We only accept media with a version number matching the version required
    # by the SRSv2 kit.
    if ($img.Version -ne $KitOsRequired) {
        Write-Host ("Your Windows installation media is version {0}. Your SRSv2 kit requires version {1}." -f $img.Version, $KitOsRequired)
        return $false
    }

    # We only accept "Enterprise" media (not, e.g., Pro, Home, or Eval)
    if ($img.EditionId -ne "Enterprise") {
        Write-Host ("You need to acquire Enterprise edition installation media. The installation media provided is '{0}', not Enterprise." -f $img.EditionId)
        PrintWhereToGetMedia -IsOem:$IsOem
        return $false
    }

    return $true
}

function Test-ePKEA {
  <#
    .SYNOPSIS

        Determine if $key is in the correct format of a Windows license key.

    .DESCRIPTION

        Determines if $key follows the basic Windows license key format of five
        dash-separated groups of five alphanumeric characters each. This
        function does *not* test if the key is a valid key -- only that it
        follows the correct formatting.

    .OUTPUTS bool

        $true if $key is in the correct format to be a Windows license key, $false otherwise.
  #>
param(
    [parameter(Mandatory=$true)]
    [AllowEmptyString()]
    [string]$key <# The Windows license key to check. #>
)

    # Could make the regex case insensitive, but this is fast and easy.
    $key = $key.ToUpperInvariant()

    $result = ($key -match '^([A-Z0-9]{5}-){4}[A-Z0-9]{5}$')
    if (!$result) {
        Write-Host ""
        Write-Host "Not a valid ePKEA. Please enter a value in the form of:"
        Write-Host "xxxxx-xxxxx-xxxxx-xxxxx-xxxxx"
        Write-Host "Where each digit represented by an 'x' is an alphanumeric (A-Z or 0-9) value."
        Write-Host ""
    }
    return $result
}

function Test-Unattend-Compat {
    <#
        .SYNOPSIS
        
            Test to see if this script is compatible with a given SRSv2 Unattend.xml file.

        .DESCRIPTION

            Looks for metadata in the $xml parameter indicating the lowest version of
            the CreateSrsMedia script the XML file will work with.

        .OUTPUTS bool
            
            Return $true if CreateSrsMedia is compatible with the SRSv2
            Unattend.xml file in $xml, $false otherwise.
    #>
param(
    [parameter(Mandatory=$true)]
    [System.Xml.XmlDocument]$Xml, <# The SRSv2 AutoUnattend to check compatibility with. #>
    [parameter(Mandatory=$true)]
    [int]$Rev <# The maximum compatibility revision this script supports. #>
)
    $nodes = $Xml.SelectNodes("//comment()[starts-with(normalize-space(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')), 'srsv2-compat-rev:')]")

    # If the file has no srsv2-compat-rev value, assume rev 0, which all scripts work with.
    if ($nodes -eq $null -or $nodes.Count -eq 0) {
        return $true
    }

    $URev = 0

    # If there is more than one value, be conservative: take the biggest value
    $nodes | 
    ForEach-Object {
        $current = $_.InnerText.Split(":")[1]
        if ($URev -lt $current) {
            $URev = $current
        }
    }

    return $Rev -ge $URev

}

function Remove-Xml-Comments {
  <#
    .SYNOPSIS
        
        Remove all comments that are direct children of $node.

    .DESCRIPTION
        
        Remove all the comment children nodes (non-recursively) from the specified $node.
  #>
param(
    [parameter(Mandatory=$true)]
    [System.Xml.XmlNode]$node <# The XML node to strip comments from. #>
)
    $node.SelectNodes("comment()") |
    ForEach-Object {
        $node.RemoveChild($_) | Write-Debug
    }
}

function Add-AutoUnattend-Key {
  <#
    .SYNOPSIS
        
        Inject $key as a product key into the AutoUnattend XML $xml.

    .DESCRIPTION
        
        Injects the $key value as a product key in $xml, where $xml is an
        AutoUnattend file already containing a Microsoft-Windows-Setup UserData
        node. Any comments in the UserData node are stripped.

        If a ProductKey node already exists, this function does *not* remove or
        replace it.
  #>
param(
    [parameter(Mandatory=$true)]
    [System.Xml.XmlDocument]$xml, <# The SRSv2 AutoUnattend to modify. #>
    [parameter(Mandatory=$true)]
    [string]$key <# The Windows license key to inject. #>
)

    $XmlNs = @{"u"="urn:schemas-microsoft-com:unattend"}
    $node = (Select-Xml -Namespace $XmlNs -Xml $xml -XPath "//u:settings[@pass='specialize']").Node
    $NShellSetup = $xml.CreateElement("", "component", $XmlNs["u"])
    $NShellSetup.SetAttribute("name", "Microsoft-Windows-Shell-Setup") | Write-Debug
    $NShellSetup.SetAttribute("processorArchitecture", "amd64") | Write-Debug
    $NShellSetup.SetAttribute("publicKeyToken", "31bf3856ad364e35") | Write-Debug
    $NShellSetup.SetAttribute("language", "neutral") | Write-Debug
    $NShellSetup.SetAttribute("versionScope", "nonSxS") | Write-Debug
    $NProductKey = $xml.CreateElement("", "ProductKey", $XmlNs["u"])
    $NProductKey.InnerText = $key
    $NShellSetup.AppendChild($NProductKey) | Write-Debug
    $node.PrependChild($NShellSetup) | Write-Debug
}

function Set-AutoUnattend-Partitions {
  <#
    .SYNOPSIS

        Set up the AutoUnattend file for use with BIOS based systems, if requested.

    .DESCRIPTION

        If -BIOS is specified, reconfigure a (nominally UEFI) AutoUnattend
        partition configuration to be compatible with BIOS-based systems
        instead. Otherwise, do nothing.
  #>
param(
    [parameter(Mandatory=$true)]
    [System.Xml.XmlDocument]$xml, <# The SRSv2 AutoUnattend to modify. #>
    [parameter(Mandatory=$true)]
    [switch]$BIOS <# If True, assume UEFI input and reconfigure for BIOS. #>
)

    # for UEFI, do nothing.
    if (!$BIOS) {
        return
    }

    # BIOS logic...
    $XmlNs = @{"u"="urn:schemas-microsoft-com:unattend"}
    $node = (Select-Xml -Namespace $XmlNs -Xml $xml -XPath "//u:settings[@pass='windowsPE']/u:component[@name='Microsoft-Windows-Setup']").Node

    # Remove the first partition (EFI)
    $node.DiskConfiguration.Disk.CreatePartitions.RemoveChild($node.DiskConfiguration.Disk.CreatePartitions.CreatePartition[0]) | Write-Debug

    # Re-number the remaining partition as 1
    $node.DiskConfiguration.Disk.CreatePartitions.CreatePartition.Order = "1"

    # Install to partition 1
    $node.ImageInstall.OSImage.InstallTo.PartitionID = "1"
}

function Set-AutoUnattend-Sysprep-Mode {
  <#
    .SYNOPSIS
        
        Set the SRSv2 sysprep mode to "reboot" or "shutdown" in the AutoUnattend file $xml.

    .DESCRIPTION
        
        Sets the SRSv2 AutoUnattend represented by $xml to either reboot (if
        -Reboot is used), or shut down (if -shutdown is used). Any comments
        under the containing RunSynchronousCommand node are stripped.

        This function assumes that a singular sysprep command is specified in
        $xml with /generalize and /oobe flags, in the auditUser pass,
        Microsoft-Windows-Deployment component. It further assumes that the
        sysprep command has the /reboot option specified by default.
  #>
param(
    [parameter(Mandatory=$true)]
    [System.Xml.XmlDocument]$Xml, <# The SRSv2 AutoUnattend to modify. #>
    [parameter(Mandatory=$true,ParameterSetName='reboot')]
    [switch]$Reboot, <# Whether sysprep should perform a reboot or a shutdown. #>
    [parameter(Mandatory=$true,ParameterSetName='shutdown')]
    [switch]$Shutdown <# Whether sysprep should perform a shutdown or a reboot. #>
)
    $XmlNs = @{"u"="urn:schemas-microsoft-com:unattend"}
    $node = (Select-Xml -Namespace $XmlNs -Xml $Xml -XPath "//u:settings[@pass='auditUser']/u:component[@name='Microsoft-Windows-Deployment']/u:RunSynchronous/u:RunSynchronousCommand/u:Path[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sysprep') and contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'generalize') and contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'oobe')]").Node
    Remove-Xml-Comments $node.ParentNode
    if ($Shutdown -or !$Reboot) {
        $node.InnerText = $node.InnerText.ToLowerInvariant() -replace ("/reboot", "/shutdown")
    }
}

function Get-TextListSelection {
  <#
    .SYNOPSIS

        Prompt the user to pick an item from an array.


    .DESCRIPTION

        Given an array of items, presents the user with a text-based, numbered
        list of the array items. The user must then select one item from the
        array (by index). That index is then returned.

        Invalid selections cause the user to be re-prompted for input.


    .OUTPUTS int

        The index of the item the user selected from the array.
  #>
  param(
    [parameter(Mandatory=$true)]<# The list of objects to select from #>
    $Options,
    [parameter(Mandatory=$false)]<# The property of the objects to use for the list #>
    $Property = $null,
    [parameter(Mandatory=$false)]<# The prompt to display to the user #>
    $Prompt = "Selection",
    [parameter(Mandatory=$false)]<# Whether to allow a blank entry to make the default selection #>
    [switch]
    $AllowDefault = $true,
    [parameter(Mandatory=$false)]<# Whether to automatically select the default value, without prompting #>
    [switch]
    $AutoDefault = $false
  )

  $index = 0
  $response = -1
  $DefaultValue = $null
  $DefaultIndex = -1

  if ($AllowDefault) {
    $DefaultIndex = 0
    if ($AutoDefault) {
      return $DefaultIndex
    }
  }

  $Options | Foreach-Object -Process {
    $value = $_
    if ($Property -ne $null) {
      $value = $_.$Property
    }
    if ($DefaultValue -eq $null) {
      $DefaultValue = $value
    }
    Write-Host("[{0,2}] {1}" -f $index, $value)
    $index++
  } -End {
    if ($AllowDefault) {
      Write-Host("(Default: {0})" -f $DefaultValue)
    }
    while ($response -lt 0 -or $response -ge $Options.Count) {
      try {
        $response = Read-Host -Prompt $Prompt -ErrorAction SilentlyContinue
        if ([string]::IsNullOrEmpty($response)) {
          [int]$response = $DefaultIndex
        } else {
          [int]$response = $response
        }
      } catch {}
    }
  }

  # Write this out for transcript purposes.
  Write-Host ("Selected option {0}." -f $response)

  return $response
}

function SyncDirectory {
  <#
    .SYNOPSIS
        Sync a source directory to a destination.

    .DESCRIPTION
        Given a source and destination directories, make the destination
        directory's contents match the source's, recursively.
  #>
  param(
    [parameter(Mandatory=$true)] <# The source directory containing the subirectory to sync. #>
    $Src,
    [parameter(Mandatory=$true)] <# The destination directory that may or may not yet contain the subdirectory being synchronized #>
    $Dst,
    [parameter(Mandatory=$false)] <# Any additional flags to pass to robocopy #>
    $Flags
  )

  (Invoke-Native "& robocopy /mir '$Src' '$Dst' /R:0 $Flags" $robocopy_success) | Write-Debug
  if ($LASTEXITCODE -gt 7) {
    Write-Error ("Copy failed. Try re-running with -Debug to see more details.{0}Source: {1}{0}Destination: {2}{0}Flags: {3}{0}Error code: {4}" -f "`n`t", $Src, $Dst, ($Flags -Join " "), $LASTEXITCODE)
  }
}

function SyncSubdirectory {
  <#
    .SYNOPSIS
        Sync a single subdirectory from a source directory to a destination.

    .DESCRIPTION
        Given a source directory Src with a subdirectory Subdir, recreate Subdir
        as a subdirectory under Dst.
  #>
  param(
    [parameter(Mandatory=$true)] <# The source directory containing the subirectory to sync. #>
    $Src,
    [parameter(Mandatory=$true)] <# The destination directory that may or may not yet contain the subdirectory being synchronized #>
    $Dst,
    [parameter(Mandatory=$true)] <# The name of the subdirectory to synchronize #>
    $Subdir,
    [parameter(Mandatory=$false)] <# Any additional flags to pass to robocopy #>
    $Flags
  )

  $Paths = Join-Path -Path @($Src, $Dst) -ChildPath $Subdir
  SyncDirectory $Paths[0] $Paths[1] $Flags
}

function SyncSubdirectories {
  <#
    .SYNOPSIS
        Recreate each subdirectory from the source in the destination.

    .DESCRIPTION
        For each subdirectory contained in the source, synchronize with a
        corresponding subdirectory in the destination. This does not synchronize
        non-directory files from the source to the destination, nor does it
        purge "extra" subdirectories in the destination where the source does
        not contain such directories.
  #>
  param(
    [parameter(Mandatory=$true)] <# The source directory #>
    $Src,
    [parameter(Mandatory=$true)] <# The destination directory #>
    $Dst,
    [parameter(Mandatory=$false)] <# Any additional flags to pass to robocopy #>
    $Flags
  )

  Get-ChildItem $Src -Directory | ForEach-Object { SyncSubdirectory $Src $Dst $_.Name $Flags }
}

function ConvertFrom-PSCustomObject {
<#
    .SYNOPSIS
        Recursively convert a PSCustomObject to a hashtable.

    .DESCRIPTION
        Converts a set of (potentially nested) PSCustomObjects into an easier-to-
        manipulate set of (potentially nested) hashtables. This operation does not
        recurse into arrays; any PSCustomObjects embedded in arrays will be left
        as-is.

    .OUTPUT hashtable
#>
param(
    [parameter(Mandatory=$true)]
    [PSCustomObject]$object <# The PSCustomeObject to recursively convert to a hashtable #>
)

    $retval = @{}

    $object.PSObject.Properties |% {
        $value = $null

        if ($_.Value -ne $null -and $_.Value.GetType().Name -eq "PSCustomObject") {
            $value = ConvertFrom-PSCustomObject $_.Value
        } else {
            $value = $_.Value
        }
        $retval.Add($_.Name, $value)
    }
    return $retval
}

function Resolve-Url {
<#
    .SYNOPSIS
        Recursively follow URL redirections until a non-redirecting URL is reached.

    .DESCRIPTION
        Chase URL redirections (e.g., FWLinks, safe links, URL-shortener links)
        until a non-redirection URL is found, or the redirection chain is deemed
        to be too long.

    .OUTPUT System.Uri
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$url <# The URL to (recursively) resolve to a concrete target. #>
)
    $orig = $url
    $result = $null
    $depth = 0
    $maxdepth = 10

    do {
        if ($depth -ge $maxdepth) {
            Write-Error "Unable to resolve $orig after $maxdepth redirects."
        }
        $depth++
        $resolve = [Net.WebRequest]::Create($url)
        $resolve.Method = "HEAD"
        $resolve.AllowAutoRedirect = $false
        $result = $resolve.GetResponse()
        $url = $result.GetResponseHeader("Location")
    } while ($result.StatusCode -eq "Redirect")

    if ($result.StatusCode -ne "OK") {
        Write-Error ("Unable to resolve {0} due to status code {1}" -f $orig, $result.StatusCode)
    }

    return $result.ResponseUri
}

function Save-Url {
<#
    .SYNOPSIS
        Given a URL, download the target file to the same path as the currently-
        running script.

    .DESCRIPTION
        Download a file referenced by a URL, with some added niceties:

          - Tell the user the file is being downloaded
          - Skip the download if the file already exists
          - Keep track of partial downloads, and don't count them as "already
            downloaded" if they're interrupted

        Optionally, an output file name can be specified, and it will be used. If
        none is specified, then the file name is determined from the (fully
        resolved) URL that was provided.

    .OUTPUT string
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$url, <# URL to download #>
    [Parameter(Mandatory=$true)]
    [String]$name, <# A friendly name describing what (functionally) is being downloaded; for the user. #>
    [Parameter(Mandatory=$false)]
    [String]$output = $null <# An optional file name to download the file as. Just a file name -- not a path! #>
)

    $res = (Resolve-Url $url)

    # If the filename is not specified, use the filename in the URL.
    if ([string]::IsNullOrEmpty($output)) {
        $output = (Split-Path $res.LocalPath -Leaf)
    }

    $File = Join-Path $PSScriptRoot $output
    if (!(Test-Path $File)) {
        Write-Host "Downloading $name... " -NoNewline
        $TmpFile = "${File}.downloading"

        # Clean up any existing (unfinished, previous) download.
        Remove-Item $TmpFile -Force -ErrorAction SilentlyContinue

        # Download to the temp file, then rename when the download is complete
        (New-Object System.Net.WebClient).DownloadFile($res, $TmpFile)
        Rename-Item $TmpFile $File -Force

        Write-Host "done"
    } else {
        Write-Host "Found $name already downloaded."
    }

    return $File
}

function Test-Signature {
<#
    .SYNOPSIS
        Verify the AuthentiCode signature of a file, deleting the file and writing
        an error if it fails verification.

    .DESCRIPTION
        Given a path, check that the target file has a valid AuthentiCode signature.
        If it does not, delete the file, and write an error to the error stream.
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$Path <# The path of the file to verify the Authenticode signature of. #>
)
    if (!(Test-Path $Path)) {
        Write-Error ("File does not exist: {0}" -f $Path)
    }

    $name = (Get-Item $Path).Name
    Write-Host ("Validating signature for {0}... " -f $name) -NoNewline

    switch ((Get-AuthenticodeSignature $Path).Status) {
        ("Valid") {
            Write-Host "success."
        }

        default {
            Write-Host "failed."

            # Invalid files should not remain where they could do harm.
            Remove-Item $Path | Write-Debug
            Write-Error ("File {0} failed signature validation." -f $name)
        }
    }
}

function PrintWhereToGetLangpacks {
param(
    [parameter(Mandatory=$false)]
    [switch]$IsOem
)
    if ($IsOem) {
        Write-Host ("   OEMs:            http://go.microsoft.com/fwlink/?LinkId=131359")
        Write-Host ("   System builders: http://go.microsoft.com/fwlink/?LinkId=131358")
    } else {
        Write-Host ("   MPSA customers:         http://go.microsoft.com/fwlink/?LinkId=125893")
        Write-Host ("   Other volume licensees: http://www.microsoft.com/licensing/servicecenter")
    }
}

function PrintWhereToGetMedia {
param(
    [parameter(Mandatory=$false)]
    [switch]$IsOem
)

    if ($IsOem) {
        Write-Host ("   OEMs must order physical Windows 10 Enterprise IoT media.")
    } else {
        Write-Host ("   Enterprise customers can access Windows 10 Enterprise media from the Volume Licensing Service Center:")
        Write-Host ("   http://www.microsoft.com/licensing/servicecenter")
    }
}

function Render-Menu {
<#
    .SYNOPSIS
      Present a data-driven menu system to the user.

    .DESCRIPTION
      Render a data-driven menu system to guide the user through more complicated
      decision-making processes.

    .NOTES
      Right now, the menu system is used only for selecting which driver pack to
      download.

      Action: Download
      Parameters:
        - Targets: an array of strings (URLs)
      Description:
        Chases redirects and downloads each URL listed in the "Targets" array.
        Verifies the downloaded file's AuthentiCode signature.
      Returns:
        a string (file path) for each downloaded file.

      Action: Menu
      Parameters:
        - Targets: an array of other MenuItem names (each must be a key in $MenuItems)
        - Message: Optional. The prompt text to use when asking for the user's
                   selection.
      Description:
        Presents a menu, composed of the names listed in "Targets," to the user. The
        menu item that is selected by the user is then recursively passed to
        Render-Menu for processing.

      Action: Redirect
      Parameters:
        - Target: A MenuItem name (must be a key in $MenuItems)
      Description:
        The menu item indicated by "Target" is recursively passed to Render-Menu
        for processing.

      Action: Warn
      Parameters:
        - Message: The warning to display to the user
      Description:
        Displays a warning consisting of the "Message" text to the user.

    .OUTPUT string
      One or more strings, each representing a downloaded file.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    $MenuItem, <# The initial menu item to process #>
    [parameter(Mandatory=$true)]
    $MenuItems, <# The menu items (recursively) referenced by the initial menu item #>
    [parameter(Mandatory=$true)]
    [hashtable]$Variables
)
    if ($MenuItem.ContainsKey("Variables")) {
        foreach ($Key in $MenuItem["Variables"].Keys) {
            if ($Variables.ContainsKey($Key)) {
                $Variables[$Key] = $MenuItem["Variables"][$Key]
            } else {
                $Variables.Add($Key, $MenuItem["Variables"][$Key])
            }
        }
    }
    Switch ($MenuItem.Action) {
        "Download" {
            Write-Verbose "Processing download menu entry."
            ForEach ($URL in $MenuItem["Targets"]) {
                $file = (Save-Url $URL "driver")
                Test-Signature $file
                Write-Output $file
            }
        }

        "Menu" {
            Write-Verbose "Processing nested menu entry."
            $Options = $MenuItem["Targets"]
            $Prompt = @{}
            if ($MenuItem.ContainsKey("Message")) {
                $Prompt = @{ "Prompt"=($MenuItem["Message"]) }
            }
            $Selection = $MenuItem["Targets"][(Get-TextListSelection -Options $Options -AllowDefault:$false @Prompt)]
            Render-Menu -MenuItem $MenuItems[$Selection] -MenuItems $MenuItems -Variables $Variables
        }

        "Redirect" {
            Write-Verbose ("Redirecting to {0}" -f $MenuItem["Target"])
            Render-Menu -MenuItem $MenuItems[$MenuItem["Target"]] -MenuItems $MenuItems -Variables $Variables
        }

        "Warn" {
            Write-Warning $MenuItem["Message"]
        }
    }
}

function Invoke-Native {
<#
    .SYNOPSIS
        Run a native command and process its exit code.

    .DESCRIPTION
        Invoke a command line specified in $command, and check the resulting $LASTEXITCODE against
        $success to determine if the command succeeded or failed. If the command failed, error out.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [string]$command, <# The native command to execute. #>
    [parameter(Mandatory=$false)]
    [ScriptBlock]$success = {$_ -eq 0} <# Test of $_ (last exit code) that returns $true if $command was successful, $false otherwise. #>
)

    Invoke-Expression $command
    $result = $LASTEXITCODE
    if (!($result |% $success)) {
        Write-Error "Command '$command' failed test '$success' with code '$result'."
        exit 1
    }
}

function Expand-Archive {
<#
    .SYNOPSIS
        Extract files from supported archives.

    .NOTES
        Supported file types are .msi and .cab.
#>
[CmdletBinding()]
param(
    [parameter(Mandatory=$true)]
    [string]$source, <# The archive file to expand. #>
    [parameter(Mandatory=$true)]
    [string]$destination <# The directory to place the extracted archive files in. #>
)

    if (!(Test-Path $destination)) {
        mkdir $destination | Write-Debug
    }

    switch ([IO.Path]::GetExtension($source)) {
        ".msi" {
            Start-Process "msiexec" -ArgumentList ('/a "{0}" /qn TARGETDIR="{1}"' -f $source, $destination) -NoNewWindow -Wait
        }
        ".cab" {
            (& expand.exe "$source" -F:* "$destination") | Write-Debug
        }
        default {
            Write-Error "Unsupported archive type."
            exit 1
        }
    }
}


####
## Start of main script
####

Start-Transcript

try {
    $AutoUnattendCompatLevel = 2

    # Just creating a lower scope for the temp vars.
    if ($true) {
        # Build a complete version string for the current host OS.
        $a = [System.Environment]::OSVersion.Version
        $b = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name UBR).UBR
        $HostOs = [version]::New($a.Major, $a.Minor, $a.Build, $b)
    }

    # Warn about versions of Windows the script may not be tested with.
    if ($HostOs -lt $DevHostOs) {
        Write-Warning "This version of Windows may not be new enough to run this script."
        Write-Warning "If you encounter problems, please update to the latest widely-available version of Windows."
    }

    Write-Host "Host OS version $HostOs"

    # We have to do the copy-paste check first, as an "exit" from a copy-paste context will
    # close the PowerShell instance (even PowerShell ISE), and prevent other exit-inducing
    # errors from being seen.
    if ([string]::IsNullOrEmpty($PSCommandPath)) {
        Write-Host "This script must be saved to a file, and run as a script."
        Write-Host "It cannot be copy-pasted into a PowerShell prompt."

        # PowerShell ISE doesn't allow reading a key, so just wait a day...
        if (Test-Path Variable:psISE) {
            Start-Sleep -Seconds (60*60*24)
            exit
        }

        # Wait for the user to see the error and acknowledge before closing the shell.
        Write-Host -NoNewLine 'Press any key to continue...'
        $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') | Out-Null
        exit
    }

    # DISM commands don't work in 32-bit PowerShell.
    try {
        if (!([Environment]::Is64BitProcess)) {
            Write-Host "This script must be run from 64-bit PowerShell."
            exit
        }
    } catch {
        Write-Host "Please make sure you have the latest version of PowerShell and the .NET runtime installed."
        exit
    }

    # Dot-sourcing is unecessary for this script, and has weird behaviors/side-effects.
    # Don't permit it.
    if ($MyInvocation.InvocationName -eq ".") {
        Write-Host "This script does not support being 'dot sourced.'"
        Write-Host "Please call the script using only its full or relative path, without a preceding dot/period."
        exit
    }

    # Like dot-sourcing, PowerShell ISE executes stuff in a way that causes weird behaviors/side-effects,
    # and is generally a hassle (and unecessary) to support.
    if (Test-Path Variable:psISE) {
        Write-Host "This script does not support being run in Powershell ISE."
        Write-Host "Please call this script using the normal PowerShell prompt, or by passing the script name directly to the PowerShell.exe executable."
        exit
    }

    # Have to be admin to do things like DISM commands.
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
        [Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Host "This script must be run from an elevated console."
        exit
    }

    Write-Host ("Script version {0}" -f $CreateSrsMediaScriptVersion)
    $UpdatedScript = Save-Url "https://go.microsoft.com/fwlink/?linkid=867842" "CreateSrsMedia" "update.ps1"
    Test-Signature $UpdatedScript
    Unblock-File $UpdatedScript
    [Version]$UpdatedScriptVersion = (& powershell -executionpolicy unrestricted ($UpdatedScript.Replace(" ", '` ')) -ShowVersion)
    if ($UpdatedScriptVersion -gt [Version]$CreateSrsMediaScriptVersion) {
        Write-Host ("Newer script found, version {0}" -f $UpdatedScriptVersion)
        Remove-Item $PSCommandPath
        Rename-Item $UpdatedScript $PSCommandPath
        $Arguments = ""
        $ScriptPart = 0

        # Find the first non-escaped space. This separates the script filename from the rest of the arguments.
        do {
            # If we find an escape character, jump over the character it's escaping.
            if($MyInvocation.Line[$ScriptPart] -eq "``") { $ScriptPart++ }
            $ScriptPart++
        } while($ScriptPart -lt $MyInvocation.Line.Length -and $MyInvocation.Line[$ScriptPart] -ne " ")

        # If we found an unescaped space, there are arguments -- extract them.
        if($ScriptPart -lt $MyInvocation.Line.Length) {
            $Arguments = $MyInvocation.Line.Substring($ScriptPart)
        }

        # Convert the script from a potentially relative path to a known-absolute path.
        # PSCommandPath does not escape spaces, so we need to do that.
        $Script = $PSCommandPath.Replace(" ", "`` ")

        Write-Host "Running the updated script."
        # Reconstruct a new, well-escaped, absolute-pathed, unrestricted call to PowerShell
        Start-Process "$psHome\powershell.exe" -ArgumentList ("-executionpolicy unrestricted " + $Script + $Arguments)
        Exit
    } else {
        Remove-Item $UpdatedScript
    }
    Write-Host ""

    # Script stats for debugging
    Write-Host (Get-FileHash -Algorithm SHA512 $PSCommandPath).Hash
    Write-Host (Get-Item $PSCommandPath).Length
    Write-Host ""

    # Initial sanity checks

    $ScriptDrive = [System.IO.DriveInfo]::GetDrives() |? { (Split-Path -Path $_.Name -Qualifier) -eq (Split-Path -Path $PSScriptRoot -Qualifier) }

    if ($ScriptDrive.DriveFormat -ne "NTFS") {
        Write-Host "This script must be run from an NTFS filesystem, as it can potentially cache very large files."
        exit
    }

    # Perform an advisory space check
    $EstimatedCacheSpace =  (1024*1024*1024*1.5) + # Estimated unpacked driver size
                            (1024*1024*1024*16) +  # Estimated exported WIM size
                            (1024*1024*100)        # Estimated unpacked SRSv2 kit size
    if ($ScriptDrive.AvailableFreeSpace -lt $EstimatedCacheSpace) {
        Write-Warning "The drive this script is running from may not have enough free space for the script to complete successfully."
        Write-Warning ("You should ensure at least {0:F2}GiB are available before continuing." -f ($EstimatedCacheSpace / (1024*1024*1024)) )
        Write-Warning "Would you like to proceed anyway?"
        do {
            $confirmation = (Read-Host -Prompt "YES or NO")
            if ($confirmation -eq "YES") {
                Write-Warning "Proceeding despite potentially insufficient scratch space."
                break
            }

            if ($confirmation -eq "NO") {
                Write-Host "Please re-run the script after you make more space available on the current drive, or move the script to a drive with more available space."
                exit
            }

            Write-Host "Invalid option."
        } while ($true)
    }

    # Determine OEM status
    $IsOem = $null
    if ($Manufacturing) {
        $IsOem = $true
    }
    while ($IsOem -eq $null) {
        Write-Host "What type of customer are you?"
        switch (Read-Host -Prompt "OEM or Enterprise") {
            "OEM" {
                $IsOem = $true
            }

            "Enterprise" {
                $IsOem = $false
            }

            Default {
                $IsOem = $null
            }
        }
    }


    if ($true) {
        $i = 1

        Write-Host ("Please make sure you have all of the following available:")
        Write-Host ("")
        Write-Host ("{0}. A USB drive with sufficient space (16GB+)." -f $i++)
        Write-Host ("   The contents of this drive WILL BE LOST!")
    if ($IsOem) {
        Write-Host ("{0}. Windows 10 Enterprise IoT media that matches your SRSv2 deployment kit." -f $i++)
    } else {
        Write-Host ("{0}. Windows 10 Enterprise media that matches your SRSv2 deployment kit." -f $i++)
    }
        PrintWhereToGetMedia -IsOem:$IsOem
    if ($IsOem) {
        Write-Host ("{0}. your ePKEA license key." -f $i++)
    }
        Write-Host ("{0}. Any language pack (LP and/or LIP) files to be included." -f $i++)
        PrintWhereToGetLangpacks -IsOem:$IsOem
        Write-Host ("")
        Write-Host ("Please do not continue until you have all these items in order.")
        Write-Host ("")
    }


    # Acquire the SRS deployment kit
    $SRSDK = Save-Url "https://go.microsoft.com/fwlink/?linkid=851168" "deployment kit"
    Test-Signature $SRSDK


    ## Extract the deployment kit.
    $RigelMedia = Join-Path $PSScriptRoot "SRSv2"

    if (Test-Path $RigelMedia) {
      Remove-Directory $RigelMedia
    }

    Write-Host "Extracting the deployment kit... " -NoNewline
    Expand-Archive $SRSDK $RigelMedia
    Write-Host "done."


    ## Pull relevant values from the deployment kit
    $RigelMedia = Join-Path $RigelMedia "Skype Room System Deployment Kit"

    $UnattendFile = Join-Path $RigelMedia "AutoUnattend.xml"

    $xml = New-Object System.Xml.XmlDocument
    $XmlNs = @{"u"="urn:schemas-microsoft-com:unattend"}
    $xml.Load($UnattendFile)

    # Check if we're compatible with this kit.
    if (!(Test-Unattend-Compat -Xml $xml -Rev $AutoUnattendCompatLevel)) {
        Write-Host "This version of CreateSrsMedia is not compatible with your deployment kit."
        Write-Host "Re-run the script and allow it to self-update."
        exit
    }

    $UnattendConfigFile = ([io.path]::Combine($RigelMedia, '$oem$', '$1', 'Rigel', 'x64', 'Scripts', 'Provisioning', 'config.json'))
    $UnattendConfig = @{}

    if ((Test-Path $UnattendConfigFile)) {
        $UnattendConfig = ConvertFrom-PSCustomObject (Get-Content $UnattendConfigFile | ConvertFrom-Json)
    }

    $SrsKitOs = (Select-Xml -Namespace $XmlNs -Xml $xml -XPath "//u:assemblyIdentity/@version").Node.Value

    # The language pack version should match the unattend version.
    $LangPackVersion = $SrsKitOs

    # In some cases, AutoUnattend does not/can not match the required media's
    # reported version number. In those cases, the correct media version is
    # explicitly specified in the config file.
    if ($UnattendConfig.ContainsKey("MediaVersion")) {
        $SrsKitOs = $UnattendConfig["MediaVersion"]
    }

    $DriverDest = ((Select-Xml -Namespace $XmlNs -Xml $xml -XPath "//u:DriverPaths/u:PathAndCredentials/u:Path/text()").ToString())

    Write-Host "This deployment kit is built for Windows build $SrsKitOs."
    Write-Host "This deployment kit draws its drivers from $DriverDest."

    # Prevent old tools (e.g., DISM) from messing up images that are newer than the tool itself,
    # and creating difficult-to-debug images that look like they work right up until you can't
    # actually install them.
    if ($HostOs -lt $SrsKitOs) {
        Write-Host ""
        Write-Host "The host OS this script is running from must be at least as new as the target"
        Write-Host "OS required by the deployment kit. Please update this machine to at least"
        Write-Host "Windows version $SrsKitOs and then re-run this script."
        Write-Host ""
        Write-Error "Current host OS is older than target OS version."
        exit
    }

    $DriverDest = $DriverDest.Replace("%configsetroot%", $RigelMedia)

    # Acquire the driver pack
    Write-Host ""
    Write-Host "Please indicate what drivers you wish to use with this installation."
    $Variables = @{}
    $DriverPacks = Render-Menu -MenuItem $UnattendConfig["Drivers"]["RootItem"] -MenuItems $UnattendConfig["Drivers"]["MenuItems"] -Variables $Variables

    $BIOS = $false

    if ($Variables.ContainsKey("BIOS")) {
        $BIOS = $Variables["BIOS"]
    }

    ## Extract the driver pack
    $DriverMedia = Join-Path $PSScriptRoot "Drivers"

    if (Test-Path $DriverMedia) {
      Remove-Directory $DriverMedia
    }

    New-Item -ItemType Directory $DriverMedia | Write-Debug

    ForEach ($DriverPack in $DriverPacks) {
        $Target = Join-Path $DriverMedia (Get-Item $DriverPack).BaseName
        Write-Host ("Extracting {0}... " -f (Split-Path $DriverPack -Leaf)) -NoNewline
        Expand-Archive $DriverPack $Target
        Write-Host "done."
    }

    # Acquire the language packs
    $LanguagePacks = @(Get-Item -Path (Join-Path $PSScriptRoot "*.cab"))
    $InstallLP = New-Object System.Collections.ArrayList
    $InstallLIP = New-Object System.Collections.ArrayList

    Write-Host "Identifying language packs... "
    ForEach ($LanguagePack in $LanguagePacks) {
        $package = $null
        try {
            $package = (Get-WindowsPackage -Online -PackagePath $LanguagePack)
        } catch {
            Write-Warning "$LanguagePack is not a language pack."
            continue
        }
        if ($package.ReleaseType -ine "LanguagePack") {
            Write-Warning "$LanguagePack is not a language pack."
            continue
        }
        $parts = $package.PackageName.Split("~")
        if ($parts[2] -ine "amd64") {
            Write-Warning "$LanguagePack is not for the right architecture."
            continue
        }
        if ($parts[4] -ine $LangPackVersion) {
            Write-Warning "$LanguagePack is not for the right OS version."
            continue
        }
        $type = ($package.CustomProperties |? {$_.Name -ieq "LPType"}).Value
        if ($type -ieq "LIP") {
            $InstallLIP.Add($LanguagePack) | Write-Debug
        } elseif ($type -ieq "Client") {
            $InstallLP.Add($LanguagePack) | Write-Debug
        } else {
            Write-Warning "$LanguagePack is of unknown type."
        }
    }
    Write-Host "... done identifying language packs."


    # Acquire the updates
    $InstallUpdates = New-Object System.Collections.ArrayList

    # Only get updates if the MSI indicates they're necessary.
    if ($UnattendConfig.ContainsKey("RequiredUpdates")) {
        $UnattendConfig["RequiredUpdates"].Keys |% {
            $URL = $UnattendConfig["RequiredUpdates"][$_]
            $File = Save-Url $URL "update $_"
            $InstallUpdates.Add($File) | Write-Debug
        }
    }

    # Verify signatures on whatever updates were aquired.
    foreach ($update in $InstallUpdates) {
        Test-Signature $update
    }

    if ($InstallLP.Count -eq 0 -and $InstallLIP.Count -eq 0 -and $InstallUpdates -ne $null) {
        Write-Warning "THIS IS YOUR ONLY CHANCE TO PRE-INSTALL LANGUAGE PACKS."
        Write-Host "Because you are pre-installing an update, you will NOT be able to pre-install language packs to the image at a later point."
        Write-Host "You are currently building an image with NO pre-installed language packs."
        Write-Host "Are you ABSOLUTELY SURE this is what you intended?"

        do {
            $confirmation = (Read-Host -Prompt "YES or NO")
            if ($confirmation -eq "YES") {
                Write-Warning "PROCEEDING TO GENERATE SLIPSTREAM IMAGE WITH NO PRE-INSTALLED LANGUAGE PACKS."
                break
            }

            if ($confirmation -eq "NO") {
                Write-Host "Please place the LP and LIP cab files you wish to use in this directory, and run the script again."
                Write-Host ""
                Write-Host "You can download language packs from the following locations:"
                PrintWhereToGetLangpacks -IsOem:$IsOem
                exit
            }

            Write-Host "Invalid option."
        } while ($true)
    }

    # Discover and prompt for selection of a reasonable target drive
    $TargetDrive = $null

    $TargetType = "USB"
    if ($Manufacturing) {
        $TargetType = "File Backed Virtual"
    }
    $ValidTargetDisks = @((Get-Disk) |? {$_.BusType -eq $TargetType})

    if ($ValidTargetDisks.Count -eq 0) {
        Write-Host "You do not have any valid media plugged in. Ensure that you have a removable drive inserted into the computer."
        exit
    }

    Write-Host ""
    Write-Host "Reminder: all data on the drive you select will be lost!"
    Write-Host ""

    $TargetDisk = ($ValidTargetDisks[(Get-TextListSelection -Options $ValidTargetDisks -Property "FriendlyName" -Prompt "Please select a target drive" -AllowDefault:$false)])

    # Acquire the Windows install media root
    do {
        $WindowsMedia = Read-Host -Prompt "Please enter the path to the root of your Windows install media"
    } while (!(Test-OsPath -OsPath $WindowsMedia -KitOsRequired $SrsKitOs -IsOem:$IsOem))

    # Acquire the ePKEA
    $LicenseKey = ""

    if ($Manufacturing) {
        $LicenseKey = "XQQYW-NFFMW-XJPBH-K8732-CKFFD"
    } elseif ($IsOem) {
        do {
            $LicenseKey = Read-Host -Prompt "Please enter your ePKEA"
        } while (!(Test-ePKEA $LicenseKey))
    }

    ###
    ## Let the user know what we've discovered
    ###

    Write-Host ""
    if ($IsOem) {
        Write-Host "Creating OEM media."
    } else {
        Write-Host "Creating Enterprise media."
    }
    Write-Host ""
    if ($BIOS) {
        Write-Host "Creating BIOS-compatible media."
    } else {
        Write-Host "Creating UEFI-compatible media."
    }
    Write-Host ""
    Write-Host "Using SRSv2 kit:      " $SRSDK
    Write-Host "Using driver packs:   "
    ForEach ($pack in $DriverPacks) {
        Write-Host "    $pack"
    }
    Write-Host "Using Windows media:  " $WindowsMedia

    Write-Host "Using language packs: "
    ForEach ($pack in $InstallLP) {
        Write-Host "    $pack"
    }
    ForEach ($pack in $InstallLIP) {
        Write-Host "    $pack"
    }

    Write-Host "Using updates:        "
    ForEach ($update in $InstallUpdates) {
        Write-Host "    $update"
    }
    Write-Host "Writing stick:        " $TargetDisk.FriendlyName
    Write-Host ""


    ###
    ## Make the stick.
    ###


    # Partition & format
    Write-Host "Formatting and partitioning the target drive... " -NoNewline
    Get-Disk $TargetDisk.DiskNumber | Initialize-Disk -PartitionStyle MBR -ErrorAction SilentlyContinue
    Clear-Disk -Number $TargetDisk.DiskNumber -RemoveData -RemoveOEM -Confirm:$false
    Get-Disk $TargetDisk.DiskNumber | Initialize-Disk -PartitionStyle MBR -ErrorAction SilentlyContinue
    Get-Disk $TargetDisk.DiskNumber | Set-Disk -PartitionStyle MBR

    ## Windows refuses to quick format FAT32 over 32GB in size.
    $part = $null
    try {
        ## For disks >= 32GB
        $part = New-Partition -DiskNumber $TargetDisk.DiskNumber -Size 32GB -AssignDriveLetter -IsActive -ErrorAction Stop
    } catch {
        ## For disks < 32GB
        $part = New-Partition -DiskNumber $TargetDisk.DiskNumber -UseMaximumSize -AssignDriveLetter -IsActive -ErrorAction Stop
    }

    $part | Format-Volume -FileSystem FAT32 -NewFileSystemLabel "SRSV2" -Confirm:$false | Write-Debug

    $TargetDrive = ("{0}:\" -f $part.DriveLetter)
    Write-Host "done."

    # Windows
    Write-Host "Copying Windows... " -NoNewline
    ## Exclude install.wim, since apparently some Windows source media are not USB EFI compatible (?) and have WIMs >4GB in size.
    SyncDirectory -Src $WindowsMedia -Dst $TargetDrive -Flags @("/xf", "install.wim")
    Write-Host "done."

    $NewInstallWim = (Join-Path $PSScriptRoot "install.wim")
    $InstallWimMnt = (Join-Path $PSScriptRoot "com-mnt")

    try {
        mkdir $InstallWimMnt | Write-Debug

        Write-Host "Copying the installation image... " -NoNewline
        Export-WindowsImage -DestinationImagePath "$NewInstallWim" -SourceImagePath (Join-Path (Join-Path $WindowsMedia "sources") "install.wim") -SourceName "Windows 10 Enterprise" | Write-Debug
        Write-Host "done."

        # Image update
        if ($InstallLP.Count -gt 0 -or $InstallLIP.Count -gt 0 -or $InstallUpdates -ne $null) {
            Write-Host "Mounting the installation image... " -NoNewline
            Mount-WindowsImage -ImagePath "$NewInstallWim" -Path "$InstallWimMnt" -Name "Windows 10 Enterprise" | Write-Debug
            Write-Host "done."

            Write-Host "Applying language packs... " -NoNewline
            ForEach ($pack in $InstallLP) {
                Add-WindowsPackage -Path "$InstallWimMnt" -PackagePath "$pack" -ErrorAction Stop | Write-Debug
            }
            ForEach ($pack in $InstallLIP) {
                Add-WindowsPackage -Path "$InstallWimMnt" -PackagePath "$pack" -ErrorAction Stop | Write-Debug
            }
            Write-Host "done."

            Write-Host "Applying updates... " -NoNewline
            ForEach ($update in $InstallUpdates) {
                Add-WindowsPackage -Path "$InstallWimMnt" -PackagePath "$update" -ErrorAction Stop | Write-Debug
            }
            Write-Host "done."

            Write-Host ""
            Write-Warning "PLEASE WAIT PATIENTLY"
            Write-Host "This next part can, on some hardware, take multiple hours to complete."
            Write-Host "Aborting at this point will result in NON-FUNCTIONAL MEDIA."
            Write-Host "To minimize wait time, consider hardware improvements:"
            Write-Host "  - Use a higher (single-core) performance CPU"
            Write-Host "  - Use a fast SSD, connected by a fast bus (6Gbps SATA, 8Gbps NVMe, etc.)"
            Write-Host ""

            Write-Host "Cleaning up the installation image... " -NoNewline
            Set-ItemProperty (Join-Path (Join-Path $TargetDrive "sources") "lang.ini") -name IsReadOnly -value $false
            Invoke-Native "& dism /quiet /image:$InstallWimMnt /gen-langini /distribution:$TargetDrive"
            Invoke-Native "& dism /quiet /image:$InstallWimMnt /cleanup-image /startcomponentcleanup /resetbase"
            Write-Host "done."

            Write-Host "Unmounting the installation image... " -NoNewline
            Dismount-WindowsImage -Path $InstallWimMnt -Save | Write-Debug
            rmdir $InstallWimMnt
            Write-Host "done."
        }

        Write-Host "Splitting the installation image... " -NoNewline
        Split-WindowsImage -ImagePath "$NewInstallWim" -SplitImagePath (Join-Path (Join-Path $TargetDrive "sources") "install.swm") -FileSize 2047 | Write-Debug
        del $NewInstallWim
        Write-Host "done."
    } catch {
        try { Dismount-WindowsImage -Path $InstallWimMnt -Discard -ErrorAction SilentlyContinue } catch {}
        del $InstallWimMnt -Force -ErrorAction SilentlyContinue
        del $NewInstallWim -Force -ErrorAction SilentlyContinue
        throw
    }

    # Drivers
    Write-Host "Injecting drivers... " -NoNewline
    SyncSubdirectories -Src $DriverMedia -Dst $DriverDest
    Write-Host "done."

    # Rigel
    Write-Host "Copying Rigel build... " -NoNewline
    SyncSubdirectories -Src $RigelMedia -Dst $TargetDrive
    Copy-Item (Join-Path $RigelMedia "*.*") $TargetDrive | Write-Debug
    Write-Host "done."

    # Snag and update the unattend
    Write-Host "Configuring unattend files... " -NoNewline

    $RootUnattendFile = ([io.path]::Combine($TargetDrive, 'AutoUnattend.xml'))
    $InnerUnattendFile = ([io.path]::Combine($TargetDrive, '$oem$', '$1', 'Rigel', 'x64', 'Scripts', 'Provisioning', 'AutoUnattend.xml'))

    ## Handle the root unattend
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($RootUnattendFile)
    if ($IsOem) {
        Add-AutoUnattend-Key $xml $LicenseKey
    }
    Set-AutoUnattend-Sysprep-Mode -Xml $xml -Shutdown
    Set-AutoUnattend-Partitions -Xml $xml -BIOS:$BIOS
    $xml.Save($RootUnattendFile)

    ## Handle the inner unattend
    $xml = New-Object System.Xml.XmlDocument
    $xml.Load($InnerUnattendFile)
    if ($IsOem) {
        Add-AutoUnattend-Key $xml "XQQYW-NFFMW-XJPBH-K8732-CKFFD"
    }
    Set-AutoUnattend-Sysprep-Mode -Xml $xml -Reboot
    Set-AutoUnattend-Partitions -Xml $xml -BIOS:$BIOS
    $xml.Save($InnerUnattendFile)

    Write-Host "done."

    # Let Windows setup know what kind of license key to check for.
    Write-Host "Selecting image... " -NoNewline
    $TargetEICfg = (Join-Path (Join-Path $TargetDrive "sources") "EI.cfg")
    $OEMEICfg = @"
[EditionID]
Enterprise
[Channel]
OEM
[VL]
0
"@
    $EnterpriseEICfg = @"
[EditionID]
Enterprise
[Channel]
Retail
[VL]
1
"@
    if ($IsOem) {
        $OEMEICfg | Out-File -FilePath $TargetEICfg -Force
    } else {
        $EnterpriseEICfg | Out-File -FilePath $TargetEICfg -Force
    }
    Write-Host "done."


    Write-Host "Cleaning up... " -NoNewline

    Remove-Directory $DriverMedia
    Remove-Directory $RigelMedia

    # This folder can sometimes cause copy errors during Windows Setup, specifically when Setup is creating the ConfigSet folder.
    Remove-Item (Join-Path $TargetDrive "System Volume Information") -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "done."


    Write-Host ""
    Write-Host "Please safely eject your USB stick before removing it."

    if ($InstallUpdates -ne $null) {
        Write-Warning "DO NOT PRE-INSTALL LANGUAGE PACKS AFTER THIS POINT"
        Write-Warning "You have applied a Windows Update to this media. Any pre-installed language packs must be added BEFORE Windows updates."
    }
} finally {
    Stop-Transcript
}
# SIG # Begin signature block
# MIIdjgYJKoZIhvcNAQcCoIIdfzCCHXsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUh9BPWz4hVxS/70EcfmUYeXCf
# 7MugghhqMIIE2jCCA8KgAwIBAgITMwAAATooqWKENAQ6aAAAAAABOjANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTkxMDIzMjMxNzEx
# WhcNMjEwMTIxMjMxNzExWjCByjELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAldBMRAw
# DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
# LTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlvbnMgTGltaXRlZDEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046QUI0MS00QjI3LUYwMjYxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0GCSqGSIb3DQEBAQUA
# A4IBDwAwggEKAoIBAQD4sio6vNOH8bDBok2LoiN4D73CYRK6DEo/NwIA1CHxN8Mg
# 67v4GW2Gg3o0lik5j5OKmWkGav1NN3Cvy6guDuvdsswW1MSAIH8HZhjYld0AgSYY
# YTtfbjerKfnCeHYz8yuS2M0rhwxhzPUp9zh1OW6KSw1Pq+NOhDc8/7kYyps3I2Vr
# T/JEshi/mrE33XHn/2QfA19MN+OxUjmPySL1OO4S5GFvDjErxZAz5XrrQMMX65/l
# GvdQw6f5hu8KuKix8RQ9gbaBIU680s40eNx5AOTLkp5weN4YpIY+IxMXp41sUfCb
# qJcoz6UlI2Nyl19mUo3wbnwQGkTdEgD6HW/tC7qRAgMBAAGjggEJMIIBBTAdBgNV
# HQ4EFgQUwE/e9+gnSO9OeN9eOgsjg8cAy9cwHwYDVR0jBBgwFoAUIzT42VJGcArt
# QPt2+7MrsMM1sw8wVAYDVR0fBE0wSzBJoEegRYZDaHR0cDovL2NybC5taWNyb3Nv
# ZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNy
# bDBYBggrBgEFBQcBAQRMMEowSAYIKwYBBQUHMAKGPGh0dHA6Ly93d3cubWljcm9z
# b2Z0LmNvbS9wa2kvY2VydHMvTWljcm9zb2Z0VGltZVN0YW1wUENBLmNydDATBgNV
# HSUEDDAKBggrBgEFBQcDCDANBgkqhkiG9w0BAQUFAAOCAQEAcfwWy/jkHd+5oHBh
# hosLBNq6tbMxeoL1zEl9LnrKgc8CT0PSq1dfueehYbXOhpRgGwR2JRqm8FzBUbKq
# vhFOKekbiajTwpcQmPYdQ/lUBSXXw2vMdvj8Qzon+quHlqISLUMG/DrZN+qoxQmO
# X6vOMjIXaa43p2+d7OycYYcq+5S+slpRufgu2ghNKdUD5GGiuXRIaqSAghxtgfWS
# +6fBK2+PUlbcozYAvCT+lbnatIy7ZBrIlD3CHlGFMk37Ng7mmCkJYLylifuqxHQr
# 7vR7jQmC8ykHBYrz95JE4nz24OPDe8MwKZbOp1ek40plnok8sw2u2xPsfQeOYLPY
# kOJsNjCCBf8wggPnoAMCAQICEzMAAAFRno2PQHGjDkEAAAAAAVEwDQYJKoZIhvcN
# AQELBQAwfjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYG
# A1UEAxMfTWljcm9zb2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMTAeFw0xOTA1MDIy
# MTM3NDZaFw0yMDA1MDIyMTM3NDZaMHQxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xHjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJVaxoZpRx00HvFVw2Z19mJUGFgU
# ZyfwoyrGA0i85lY0f0lhAu6EeGYnlFYhLLWh7LfNO7GotuQcB2Zt5Tw0Uyjj0+/v
# UyAhL0gb8S2rA4fu6lqf6Uiro05zDl87o6z7XZHRDbwzMaf7fLsXaYoOeilW7SwS
# 5/LjneDHPXozxsDDj5Be6/v59H1bNEnYKlTrbBApiIVAx97DpWHl+4+heWg3eTr5
# CXPvOBxPhhGbHPHuMxWk/+68rqxlwHFDdaAH9aTJceDFpjX0gDMurZCI+JfZivKJ
# HkSxgGrfkE/tTXkOVm2lKzbAhhOSQMHGE8kgMmCjBm7kbKEd2quy3c6ORJECAwEA
# AaOCAX4wggF6MB8GA1UdJQQYMBYGCisGAQQBgjdMCAEGCCsGAQUFBwMDMB0GA1Ud
# DgQWBBRXghquSrnt6xqC7oVQFvbvRmKNzzBQBgNVHREESTBHpEUwQzEpMCcGA1UE
# CxMgTWljcm9zb2Z0IE9wZXJhdGlvbnMgUHVlcnRvIFJpY28xFjAUBgNVBAUTDTIz
# MDAxMis0NTQxMzUwHwYDVR0jBBgwFoAUSG5k5VAF04KqFzc3IrVtqMp1ApUwVAYD
# VR0fBE0wSzBJoEegRYZDaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraW9wcy9j
# cmwvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNybDBhBggrBgEFBQcBAQRV
# MFMwUQYIKwYBBQUHMAKGRWh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2lvcHMv
# Y2VydHMvTWljQ29kU2lnUENBMjAxMV8yMDExLTA3LTA4LmNydDAMBgNVHRMBAf8E
# AjAAMA0GCSqGSIb3DQEBCwUAA4ICAQBaD4CtLgCersquiCyUhCegwdJdQ+v9Go4i
# Elf7fY5u5jcwW92VESVtKxInGtHL84IJl1Kx75/YCpD4X/ZpjAEOZRBt4wHyfSlg
# tmc4+J+p7vxEEfZ9Vmy9fHJ+LNse5tZahR81b8UmVmUtfAmYXcGgvwTanT0reFqD
# DP+i1wq1DX5Dj4No5hdaV6omslSycez1SItytUXSV4v9DVXluyGhvY5OVmrSrNJ2
# swMtZ2HKtQ7Gdn6iNntR1NjhWcK6iBtn1mz2zIluDtlRL1JWBiSjBGxa/mNXiVup
# MP60bgXOE7BxFDB1voDzOnY2d36ztV0K5gWwaAjjW5wPyjFV9wAyMX1hfk3aziaW
# 2SqdR7f+G1WufEooMDBJiWJq7HYvuArD5sPWQRn/mjMtGcneOMOSiZOs9y2iRj8p
# pnWq5vQ1SeY4of7fFQr+mVYkrwE5Bi5TuApgftjL1ZIo2U/ukqPqLjXv7c1r9+si
# eOcGQpEIn95hO8Ef6zmC57Ol9Ba1Ths2j+PxDDa+lND3Dt+WEfvxGbB3fX35hOaG
# /tNzENtaXK15qPhErbCTeljWhLPYk8Tk8242Z30aZ/qh49mDLsiL0ksurxKdQtXt
# v4g/RRdFj2r4Z1GMzYARfqaxm+88IigbRpgdC73BmwoQraOq9aLz/F1555Ij0U3o
# rXDihVAzgzCCBgcwggPvoAMCAQICCmEWaDQAAAAAABwwDQYJKoZIhvcNAQEFBQAw
# XzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcGCgmSJomT8ixkARkWCW1pY3Jvc29m
# dDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# MB4XDTA3MDQwMzEyNTMwOVoXDTIxMDQwMzEzMDMwOVowdzELMAkGA1UEBhMCVVMx
# EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoT
# FU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEhMB8GA1UEAxMYTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgUENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAn6Fssd/b
# SJIqfGsuGeG94uPFmVEjUK3O3RhOJA/u0afRTK10MCAR6wfVVJUVSZQbQpKumFww
# JtoAa+h7veyJBw/3DgSY8InMH8szJIed8vRnHCz8e+eIHernTqOhwSNTyo36Rc8J
# 0F6v0LBCBKL5pmyTZ9co3EZTsIbQ5ShGLieshk9VUgzkAyz7apCQMG6H81kwnfp+
# 1pez6CGXfvjSE/MIt1NtUrRFkJ9IAEpHZhEnKWaol+TTBoFKovmEpxFHFAmCn4Tt
# VXj+AZodUAiFABAwRu233iNGu8QtVJ+vHnhBMXfMm987g5OhYQK1HQ2x/PebsgHO
# IktU//kFw8IgCwIDAQABo4IBqzCCAacwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4E
# FgQUIzT42VJGcArtQPt2+7MrsMM1sw8wCwYDVR0PBAQDAgGGMBAGCSsGAQQBgjcV
# AQQDAgEAMIGYBgNVHSMEgZAwgY2AFA6sgmBAVieX5SUT/CrhClOVWeSkoWOkYTBf
# MRMwEQYKCZImiZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0
# MS0wKwYDVQQDEyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHmC
# EHmtFqFKoKWtTHNY9AcTLmUwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL2NybC5t
# aWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvbWljcm9zb2Z0cm9vdGNlcnQu
# Y3JsMFQGCCsGAQUFBwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3d3dy5taWNy
# b3NvZnQuY29tL3BraS9jZXJ0cy9NaWNyb3NvZnRSb290Q2VydC5jcnQwEwYDVR0l
# BAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEFBQADggIBABCXisNcA0Q23em0rXfb
# znlRTQGxLnRxW20ME6vOvnuPuC7UEqKMbWK4VwLLTiATUJndekDiV7uvWJoc4R0B
# hqy7ePKL0Ow7Ae7ivo8KBciNSOLwUxXdT6uS5OeNatWAweaU8gYvhQPpkSokInD7
# 9vzkeJkuDfcH4nC8GE6djmsKcpW4oTmcZy3FUQ7qYlw/FpiLID/iBxoy+cwxSnYx
# PStyC8jqcD3/hQoT38IKYY7w17gX606Lf8U1K16jv+u8fQtCe9RTciHuMMq7eGVc
# WwEXChQO0toUmPU8uWZYsy0v5/mFhsxRVuidcJRsrDlM1PZ5v6oYemIp76KbKTQG
# dxpiyT0ebR+C8AvHLLvPQ7Pl+ex9teOkqHQ1uE7FcSMSJnYLPFKMcVpGQxS8s7Ow
# TWfIn0L/gHkhgJ4VMGboQhJeGsieIiHQQ+kr6bv0SMws1NgygEwmKkgkX1rqVu+m
# 3pmdyjpvvYEndAYR7nYhv5uCwSdUtrFqPYmhdmG0bqETpr+qR/ASb/2KMmyy/t9R
# yIwjyWa9nR2HEmQCPS2vWY+45CHltbDKY7R4VAXUQS5QrJSwpXirs6CWdRrZkocT
# dSIvMqgIbqBbjCW/oO+EyiHW6x5PyZruSeD3AWVviQt9yGnI5m7qp5fOMSn/DsVb
# XNhNG6HY+i+ePy5VFmvJE6P9MIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCBI4wggSKAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAFRno2PQHGjDkEAAAAAAVEwCQYFKw4DAhoFAKCB
# ojAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUeqafbnBqFbOh6/ZrLH7lUcN/6IYw
# QgYKKwYBBAGCNwIBDDE0MDKgFIASAE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbTANBgkqhkiG9w0BAQEFAASCAQANRlRrXAtHIhl5
# q1vjSWPyeuMQjQfPMie7lMxSoL1KLARsNRPd0JvEQtxZD+MaUgvtEhx2lniqSDnF
# spgWmlqZLvil6GPQrFSLVLbkonGn08RZWvbmqBSLqhlPYMJHykSYfyXjSReMUeJA
# lD7A+OT7nzDT0E9IstM6/+xIgm9UEzTEYh3DpM1B988lPGTHpNsAACqAtdc4x0D+
# fcReLKyr87A0PfeMOyR0wHr/uyFRKe7NmpVppAAcITqvzurg6+M/jy8/FPuFvMBm
# OVhs9ycyQcQhsFNjyjVeE1Nbai7CRrJzCR2zJXaXRifgKUAUPWRUw/kdAPVrQGN/
# TJjKxR0OoYICKDCCAiQGCSqGSIb3DQEJBjGCAhUwggIRAgEBMIGOMHcxCzAJBgNV
# BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
# HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAfBgNVBAMTGE1pY3Jvc29m
# dCBUaW1lLVN0YW1wIFBDQQITMwAAATooqWKENAQ6aAAAAAABOjAJBgUrDgMCGgUA
# oF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMTkx
# MjA1MTgyODAzWjAjBgkqhkiG9w0BCQQxFgQUDqYVLl3++6JeMm3a6/duzRJGcwkw
# DQYJKoZIhvcNAQEFBQAEggEAZyeAlOQF6t3W76vE9zgN7xLqEHTUznM6KFLl4VhG
# 0Cx6dhGyMMWVn4zS2HdI/h+ReciQSUK5/9zxb1zX2rCST25oaLUOhd/DzXPL6qDZ
# SPfRNx8xLGiA+lJ9ik8ZF4y2VoYQgyzyMPbY+u/NQMyveNuMUVw5vpJI6BHd5aWD
# cG4mRNZafq9juLyoGrFCqHGeaC6RCGUIllV73SZNjBHUvC7nCr6pCVujpbHuwMQ7
# Wx5Q1BOKpZ8tCSYqsF7iNtsFlHdJEL2mtml4h5lYfI6F4o1gNAycp6R0ZFotw1yX
# tgy0uzFDRgIuVHDONAUzoJ4MVexvMRFk1aIQC1aWed//Lw==
# SIG # End signature block
