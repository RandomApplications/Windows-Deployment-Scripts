#
# MIT License
#
# Copyright (c) 2021 Free Geek
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),
# to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,
# and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#

$Host.UI.RawUI.WindowTitle = 'Download Network Drivers for WinRE USB Install from Cache'

$basePath = "$Env:PUBLIC\Windows Deployment"
if (Test-Path "$Env:PUBLIC\Windows Deployment.lnk") {
	$basePath = (New-Object -ComObject WScript.Shell).CreateShortcut("$Env:PUBLIC\Windows Deployment.lnk").TargetPath
}

if (-not (Test-Path $basePath)) {
	New-Item -ItemType 'Directory' -Path $basePath -ErrorAction Stop | Out-Null
}

$winREnetDriversOutputPath = "$basePath\WinRE Network Drivers for USB Install"


$driversCacheBasePath = '\\FG-WindowsNAS\FG-Windows-Drivers\Cache' # SMB share credentials SHOULD BE SAVED in "Credential Manager" app so that it will auto-connect when the path is specified.

if (-not (Test-Path $driversCacheBasePath)) {
	Write-Host "`n  ERROR: Failed to connect to local Free Geek SMB share `"$driversCacheBasePath`"." -ForegroundColor Red

	$Host.UI.RawUI.FlushInputBuffer() # So that key presses before this point are ignored.
	Read-Host "`n`n  FAILED TO DOWNLOAD NETWORK DRIVERS FROM CACHE" | Out-Null

	exit 1
}


Write-Output "`n  Checking Drivers Cache for Stray (No Longer Referenced) Drivers to Delete..."

try {
	$allReferencedCachedDrivers = @()

	Get-ChildItem "$driversCacheBasePath\*" -File -Include '*.txt' -ErrorAction Stop | ForEach-Object {
		$allReferencedCachedDrivers += Get-Content $_.FullName -ErrorAction Stop
	}

	$allReferencedCachedDrivers = $allReferencedCachedDrivers | Sort-Object -Unique

	$deletedStrayCachedDriversCount = 0

	$currentEpochTime = [int64](Get-Date -UFormat '%s') # ALSO check for any LOCK files over a day old and delete them and their associated drivers (assuming if they exist something went wrong and the driver may be incomplete).
	Get-ChildItem "$driversCacheBasePath\Unique Drivers\*" -File -Include '*-CACHING.lock' -ErrorAction Stop | ForEach-Object {
		$thisCachedDriverLockFilePathAge = ($currentEpochTime - [int64](Get-Content -Raw $_.FullName -ErrorAction Stop))

		if (($null -eq $thisCachedDriverLockFilePathAge) -or ($thisCachedDriverLockFilePathAge -ge 86400) -or ($thisCachedDriverLockFilePathAge -lt 0)) {
			$thisCachedDriverDirectoryPath = $_.FullName.Replace('-CACHING.lock', '')
			if (Test-Path $thisCachedDriverDirectoryPath) {
				Remove-Item $thisCachedDriverDirectoryPath -Recurse -Force -ErrorAction Stop
			}

			Remove-Item $_.FullName -ErrorAction Stop

			$deletedStrayCachedDriversCount ++
		}
	}

	Get-ChildItem "$driversCacheBasePath\Unique Drivers" -Directory -ErrorAction Stop | ForEach-Object {
		if ((-not $allReferencedCachedDrivers.Contains($_.Name)) -and (-not (Test-Path "$($_.FullName)-CACHING.lock"))) {
			Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop
			$deletedStrayCachedDriversCount ++
		}
	}

	if ($deletedStrayCachedDriversCount -eq 0) {
		Write-Host "`n  Successfully Checked Drivers Cache and Found No Strays to Delete" -ForegroundColor Green
	} else {
		Write-Host "`n  Successfully Deleted $deletedStrayCachedDriversCount Strays From Drivers Cache" -ForegroundColor Green
	}
} catch {
	Write-Host "`n  ERROR DELETED STRAY CACHED DRIVERS: $_" -ForegroundColor Red
	Write-Host "`n  ERROR: Failed to check for or delete stray cached drivers." -ForegroundColor Red
}

if (-not (Test-Path $winREnetDriversOutputPath)) {
	New-Item -ItemType 'Directory' -Path $winREnetDriversOutputPath -ErrorAction Stop | Out-Null
}


Write-Output "`n`n  Downloading WinRE Network Drivers for USB Install from Driver Cache..."

$allCachedDriverPaths = (Get-ChildItem "$driversCacheBasePath\Unique Drivers" -Directory).FullName

$netDriverFolderNames = @()

# This Driver .inf parsing code is based on code written for "Install Windows.ps1"
$thisDriverIndex = 0
$downloadedNetDriversCount = 0
foreach ($thisDriverFolderPath in $allCachedDriverPaths) {
	$thisDriverFolderName = (Split-Path $thisDriverFolderPath -Leaf)

	if ($thisDriverFolderName.Contains('.inf_amd64_')) {
		$thisDriverInfContents = Get-Content "$thisDriverFolderPath\$($thisDriverFolderName.Substring(0, $thisDriverFolderName.LastIndexOf('.'))).inf"

		foreach ($thisDriverInfLine in $thisDriverInfContents) {
			if (($lineCommentIndex = $thisDriverInfLine.IndexOf(';')) -gt -1) { # Remove .inf comments from each line before any parsing to avoid matching any text within comments.
				$thisDriverInfLine = $thisDriverInfLine.Substring(0, $lineCommentIndex)
			}

			$thisDriverInfLine = $thisDriverInfLine.Trim()

			if ($thisDriverInfLine -ne '') {
				$thisDriverInfLineUPPER = $thisDriverInfLine.ToUpper()

				if ($thisDriverInfLine.StartsWith('[')) {
					# https://docs.microsoft.com/en-us/windows-hardware/drivers/install/inf-version-section
					$wasInfVersionSection = $isInfVersionSection
					$isInfVersionSection = ($thisDriverInfLineUPPER -eq '[VERSION]')

					if ($wasInfVersionSection -and (-not $isInfVersionSection)) {
						# If passed Version section and didn't already break from getting a NET class, then we can stop reading lines because we don't want this driver.
						break
					}
				} elseif ($isInfVersionSection -and (($lineEqualsIndex = $thisDriverInfLine.IndexOf('=')) -gt -1) -and $thisDriverInfLineUPPER.Contains('CLASS') -and (-not $thisDriverInfLineUPPER.Contains('CLASSGUID'))) {
					$thisDriverClass = $thisDriverInfLine.Substring($lineEqualsIndex + 1).Trim().ToUpper() # It appears that the Class Names will never be in quotes or be variables that need to be translated.

					if ($thisDriverClass -eq 'NET') {
						$thisDriverIndex ++
						$thisDriverSizeMB = $([math]::Round(((Get-ChildItem -Path $thisDriverFolderPath -Recurse | Measure-Object -Property 'Length' -Sum).Sum / 1MB), 2))
						if ($thisDriverSizeMB -lt 20) {
							$netDriverFolderNames += $thisDriverFolderName
							if (-not (Test-Path "$winREnetDriversOutputPath\$thisDriverFolderName")) {
								try {
									Write-Output "    $thisDriverIndex) Downloading Network Driver: $thisDriverFolderName ($thisDriverSizeMB MB)..."
									Copy-Item $thisDriverFolderPath $winREnetDriversOutputPath -Recurse -Force -ErrorAction Stop
									$downloadedNetDriversCount ++
								} catch {
									Write-Host "      ERROR DOWNLOADING NETWORK DRIVER `"$thisDriverFolderName`": $_" -ForegroundColor Red
								}
							} else {
								Write-Output "    $thisDriverIndex) ALREADY DOWNLOADED NETWORK DRIVER: $thisDriverFolderName ($thisDriverSizeMB MB)"
							}
						} else {
							Write-Output "    $thisDriverIndex) SKIPPING LARGE NETWORK DRIVER: $thisDriverFolderName ($thisDriverSizeMB MB)"
						}
					}

					break
				}
			}
		}
	}
}

Write-Host "`n  Downloaded $downloadedNetDriversCount Network Drivers from Cache" -ForegroundColor Green


if ($thisDriverIndex -gt 0) {
	Write-Output "`n`n  Checking Downloaded Network Drivers for Stray (No Longer Cached) Drivers to Delete..."

	try {
		$deletedStrayDownloadedNetDriversCount = 0

		Get-ChildItem "$winREnetDriversOutputPath" -ErrorAction Stop | ForEach-Object {
			if (-not $netDriverFolderNames.Contains($_.Name)) {
				$deletedStrayDownloadedNetDriversCount ++
				Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop
			}
		}

		if ($deletedStrayDownloadedNetDriversCount -eq 0) {
			Write-Host "`n  Checked Downloaded Network Drivers and Found No Strays to Delete" -ForegroundColor Green
		} else {
			Write-Host "`n  Deleted $deletedStrayDownloadedNetDriversCount Strays From Downloaded Network Drivers" -ForegroundColor Green
		}
	} catch {
		Write-Host "`n  ERROR DELETED STRAY DOWNLOADED NETWORK DRIVERS: $_" -ForegroundColor Red
		Write-Host "`n  ERROR: Failed to check for or delete stray downloaded network drivers." -ForegroundColor Red
	}
}


$Host.UI.RawUI.FlushInputBuffer() # So that key presses before this point are ignored.
Read-Host "`n`n  DONE DOWNLOADING NETWORK DRIVERS FROM CACHE" | Out-Null
