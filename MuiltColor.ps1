# Function to get the input file
function Get-InputFile {
    param (
        [string]$prompt = "Select the input G-code file"
    )
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = $prompt
    $fileDialog.Filter = "G-code files (*.gcode)|*.gcode|All files (*.*)|*.*"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    } else {
        throw "No file selected."
    }
}

# Function to get the output file
function Get-OutputFile {
    param (
        [string]$prompt = "Select location and name for the output G-code file"
    )
    Add-Type -AssemblyName System.Windows.Forms
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Title = $prompt
    $fileDialog.Filter = "G-code files (*.gcode)|*.gcode|All files (*.*)|*.*"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    } else {
        throw "No file selected."
    }
}
function Process-GCode {
    param (
        [string]$inputFile,
        [string]$outputFile,
        [int]$safeZLift = 15,
        [array]$toolNames
    )

    Write-Progress -Activity "Processing G-code" -Status "Processing..." -PercentComplete 0 -id 1

    # Initialize variables for the last coordinates, layer number, and tool number
    $global:toolNumber = 0
    $insertedM117 = $false
   # $extrusionAmount = 0
   # $insideToolChangeLoad = $false

    # Create a new list to store the modified lines
    $modifiedLines = New-Object System.Collections.Generic.List[System.String]

    # Read the file line by line and process in reverse
    $fileContent = Get-Content $inputFile
    $maxline = $fileContent.Length
    $ChangeCount = 0
     $totalChanges = 1
    for ($i = $maxline - 1; $i -ge 0; $i--) {
        Write-Progress -Activity "Processing G-code:   $($line) " -Status "Processing...  $( $ChangeCount) of $totalChanges" -PercentComplete (($ChangeCount / $totalChanges) * 100) -id 2
         Write-Progress -Activity "Reading Line $($i) || $((($maxline - $i) / $maxline) * 100)% || "  -PercentComplete ((($maxline - $i) / $maxline) * 100) -Id  1
        $line = $fileContent[$i]
        if($line -match "; total filament change = (\d+)") {
          $totalChanges = $matches[1]
        }
        # Check if the line indicates a manual tool change
        if ($line -match "; MANUAL_TOOL_CHANGE T(\d+)") {
            $global:toolNumber = $matches[1]
            $ChangeCount++
            $insertedM117 = $false
        }

        # Check if inside tool change load section
       # if ($line -match "; CP TOOLCHANGE WIPE") {
       #     $insideToolChangeLoad = $true
       # } elseif ($line -match "; CP TOOLCHANGE LOAD") {
       #     $insideToolChangeLoad = $false
       # }

        # Remove E# and save extrusion amount
       # if ($insideToolChangeLoad -and $line -match "G1 .*E([0-9.+-]+)") {
            #$extrusionAmount += [double]$matches[1]
          #  $line = $line -replace " E[0-9.+-]+", ""
       # }

        # Insert M117 and M106 S0 commands before M600 command if not already inserted
        if ($line.StartsWith("M600") -and -not $insertedM117) {       
          #  if ($extrusionAmount -ne 0) {
           #     $modifiedLines.Insert(0, "G1 E$extrusionAmount F180 ; Restore extrusion amount")
           #     $extrusionAmount = 0
           # }
            $modifiedLines.Insert(0, "M600 -B2 -L5 -U10 ;Wait for Tool Change")
            $modifiedLines.Insert(0, "M117 Change to $($toolNames[$global:toolNumber]) Changes left $( $ChangeCount); Display the tool number in the terminal")
            $modifiedLines.Insert(0, "M106 S0; Turn off fan for tool change")
            $line=''
            $insertedM117 = $true

        }

        # Add the current line to the front of the modified lines list
        $modifiedLines.Insert(0, $line)
    }

      $modifiedLines.Insert(3, '; Modify by using MuiltColor.ps1 © Carter Cordingley')

    # Write the modified lines to the output file
    $modifiedLines | Set-Content -Path $outputFile -Encoding UTF8

    Write-Output "G-code modifications complete!"
}
function Clean-Memory {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Write-Output "Memory cleanup complete!"
    }
# Main script execution
$inputFile = Get-InputFile
$outputFile = Get-OutputFile
$toolNames = @("Yellow", "Black", "LightGreen", "DarkGreen")
Clear-Host
Clean-Memory
 Get-Date
Process-GCode -inputFile $inputFile -outputFile $outputFile -safeZLift 15 -toolNames $toolNames
 Get-Date