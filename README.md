# Terminal based serial port shell for Windows
For all my fellow terminal hermits (Termits). 

Use this module if:
- You don't want/need to use PuTTY/Tera Term
- Need a native Windows solution.

## Installation

To install, simply copy the `.psm1` file to a known location and add the following to your `$PROFILE`:
**NOTE:** Be sure to change the file path as needed.
```pwsh
Import-Module "C:\path_to\the_module\serial_shell.psm1"
```

## Usage

You can invoke the module with:
```pwsh
SerialShell -COMPort # -BaudRate 115200
```

- To Exit, `<Ctrl+a>` then `z`
- To list all serial ports currently connected to your PC `<Ctrl+a>` then `l`
