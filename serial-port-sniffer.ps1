$port = New-Object System.IO.Ports.SerialPort
$port.PortName = “COM36”  # change as needed
$port.BaudRate = “57600”
$port.Parity = “None”
$port.DataBits = 8
$port.StopBits = 1
$port.ReadTimeout = 9000  # 9 seconds
$port.DtrEnable = “true”
$port.open() # opens serial connection

Start-Sleep -m 100 # wait 100ms seconds until device is ready

try {
    Start-Sleep -m 100

    while($myinput = $port.ReadExisting()) {
        $myinput
    }

    $port.Write(“s `r`n”) #writes command to display status
    Start-Sleep -m 100

    while($myinput = $port.ReadExisting()) {
        $myinput
    }
}

catch [TimeoutException] {
    # Error handling code here
}

finally {
    # Any cleanup code goes here
}

$port.Close()  # closes serial connection
