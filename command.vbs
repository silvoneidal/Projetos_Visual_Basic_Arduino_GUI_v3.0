Set WshShell = CreateObject("WScript.Shell")

strCommand = Wscript.arguments(0)

Set WshShell = CreateObject("WScript.Shell")
Set objCmd = WshShell.Exec("cmd /c " & strCommand)

strOutput = ""
Do While Not objCmd.StdOut.AtEndOfStream
    strOutput = strOutput & objCmd.StdOut.ReadLine() & vbCrLf
Loop

WScript.StdOut.WriteLine strOutput
WScript.Quit


