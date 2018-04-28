' ______ ____    ____ .______    _______ .______           _______. __  .___  ___.   ______   .__   __. 
'/      |\   \  /   / |   _  \  |   ____||   _  \         /       ||  | |   \/   |  /  __  \  |  \ |  | 
'|  ,----' \   \/   /  |  |_)  | |  |__   |  |_) |       |   (----`|  | |  \  /  | |  |  |  | |   \|  | 
'|  |       \_    _/   |   _  <  |   __|  |      /         \   \   |  | |  |\/|  | |  |  |  | |  . `  | 
'|  `----.    |  |     |  |_)  | |  |____ |  |\  \----..----)   |  |  | |  |  |  | |  `--'  | |  |\   | 
'\_______|    |__|     |______/  |_______|| _| `._____||_______/   |__| |__|  |__|  \______/  |__| \__| 
'
'Copyright 2018          Simon Durkee          Simon@SimonDurkee.com         http://www.SimonDurkee.com                                                                                                     
'         _______________________________________
'________|                                      |_______
'\       |       DownloadURLs-FFMpeg.vbs        |      /
' \      |                                      |     /
' /      |______________________________________|     \
'/__________)                                (_________\
'
'   INPUT FILE FORMAT
'       Episode#,URL | One per line, Episode# is optional
'   PARAMETERS
'       Filename - Filename to process | REQUIRED
'       ASyncronous - Download List Asyncronously | OPTIONAL Default=False
'   CHANGE LOG
'       03-18-2018 - 1.0 - Release Version
'       03-18-2018 - 1.1 - Added Episode # in Input File
'       03-19-2018 - 1.2 - Set FFMpeg Window to Minimized
'                          Overwrite file if exists
'	04-28-2018 - 1.3 - Change Episode # in Input File to Include Full Season: eg S00E03,{URL}
'
'
set fs=CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
strPath = Wscript.ScriptFullName
arrayPath = Split(strPath,"\")
strShow = arrayPath(UBound(arrayPath)-1)
strSeason = Replace(GetArgs(1),".txt","")
intEpisode = 0
strFilename = GetArgs(1)
boolASync = GetArgs(2)
If LCase(boolASync)="true" or boolASync="1" Then
    boolASync = True
Else
    boolASync = False
End If
Set objArgs = Wscript.Arguments
'Wscript.Echo "Reading: " & strFilename
Set inFile = fs.OpenTextFile(strFilename)
While Not inFile.AtEndOfStream
    inLine = inFile.ReadLine()
    If InStr(inLine,",") Then
        arrayLine = Split(inLine,",")
	strEpisode = arrayLine(0)
        URL = arrayLine(1)
    Else
        URL = inLine
        intEpisode = intEpisode + 1
	strEpisode = strSeason & "E" & PadLeft(intEpisode,"0",2)
    End If
    strVideoFilename = strShow & " - " & strEpisode & ".mp4"
    Wscript.Echo "Downloading " & strVideoFilename
    objShell.Run "FFMPEG -i """ & URL & """ -c copy -y """ & strVideoFilename & """",2,Not boolASync
    
Wend

Function GetArgs(intArgPosition)
    Dim ArgValue,objArgs
    ArgValue=""
    If intArgPosition=0 Then
        ArgValue = Wscript.ScriptName
    Else
        Set objArgs = Wscript.Arguments
        If objArgs.Count => intArgPosition Then
            ArgValue = objArgs(intArgPosition-1)
        End If
    End If
    GetArgs = ArgValue
End Function
Function PadLeft(strInput,strPadChar,intLength)
    While Len(strInput)<>intLength
        strInput = strPadChar & strInput
    Wend
    PadLeft = strInput
End Function
Function RunProgram(prog)
    Dim objExec,s,line
    Set objExec = objShell.Exec(prog)
    Do
        line = objExec.StdOut.ReadLine()
        s = s & line & vbcrlf
    Loop While Not objExec.Stdout.atEndOfStream
    RunProgram=s
End Function
