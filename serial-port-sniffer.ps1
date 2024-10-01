<# 
Author: Nathan Winslow
Date: 10/1/2024

.SYNOPSIS
The purpose of this script is to send data to and receive data from a serial COM port.
The script allows the user to choose a port and open it, then type commands in to the host device.

TODO:
- Allow user to enter in Serial port parameters (baud rate, parity, etc.)
- 
#>

function Validate {
    param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern("COM\d")] # Validate the use selected a valid port.
        [String[]]$port
    )
}

function print_port_info {
  # Initialize an array to hold the port status objects
  $portStatuses = @()

  # https://superuser.com/a/1835725
  $all_ports = Get-CimInstance -Class Win32_SerialPort | Select-Object Name, Description
  foreach ($port in $all_ports) {
    try {
      # Try to initialize the port
      $testPort = new-Object System.IO.Ports.SerialPort $port,115200,None,8,one
      $testPort.Open() # Try to open the port
      $testPort.Close() # Close the port if it was successfully opened
      
      # Create a custom object with port details and status
      $portStatus = New-Object PSObject -Property @{
          PortName = $port.Name
          Status = 'Available'
      }
    }
    catch {
      # If an exception is caught, the port is busy
      $portStatus = New-Object PSObject -Property @{
          PortName = $port.Name
          Status = 'Busy'
      }
    }
  # Add the custom object to the array
  $portStatuses += $portStatus
  }
  $portStatuses | Format-Table -AutoSize
}

function print_help {
  $menuOptions = @()
  $helpOption = New-Object PSObject -Property @{
    command = "-h"
    description = "Prints the help menu."
    flags = ""
  }
  $deleteOption = New-Object PSObject -Property @{
    command = "-q"
    description = "Exits the program."
    flags = ""
  }
  $menuOptions += $helpOption
  $menuOptions += $deleteOption
  # Clear-Host
  $menuOptions | Format-Table -Autosize
}

# Clear-Host
print_port_info
$portname = Read-Host -Prompt "Select a COM port"
Validate $portname
# TODO: loop back if user selected invalid port
Write-Host "You selected:" $portname

$port = New-Object System.IO.Ports.SerialPort
$port.PortName = $portname
$port.BaudRate = "115200"
$port.Parity = "None"
$port.DataBits = 8
$port.StopBits = 1
$port.ReadTimeout = 10000 # 10 seconds
$port.DtrEnable = "true"
if ($port.IsOpen) {
  $port.Close()
}

Start-Sleep -Milliseconds 1000 # wait 1000ms
$port.Open()

# Clear-Host
print_help

do {
$command = Read-Host -Prompt ">"
switch ($command) {
  '-q'{ $port.Close(); break }
  '-h'{ print_help }
  Default {
    $port.WriteLine($command)
    while ($port.BytesToRead) {
      $host_out = $port.ReadExisting()
    }
    Write-Host $host_out
  }
}
} while ($port.IsOpen)

Write-Host "Exiting Program..."
if ($port.IsOpen) {
  Write-Host "Port Not closed."
  $port.Close()
  $port.Dispose()
}
