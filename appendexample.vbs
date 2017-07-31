Option Explicit

Const ForReading = 1, ForAppending = 8

Dim inputFileName, outputFileName
    inputFileName  = "C:\Users\Dimitar\Desktop\BPSDRC\PAYBOTH.dat"
    outputFileName = "C:\Users\Dimitar\Desktop\BPSDRC\PAYIMP.dat"

Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists( inputFileName ) Then
        WScript.Quit
		WScript.echo "No File"
    End If

    If fso.GetFile( inputFileName ).Size < 1 Then 
        WScript.Quit
    End If 

Dim newFile, inputFile, outputFile 
    newFile = Not fso.FileExists( outputFileName )

    Set inputFile = fso.OpenTextFile( inputFileName, ForReading )
    Set outputFile = fso.OpenTextFile( outputFileName, ForAppending, True )

Dim lineBuffer

    lineBuffer = inputFile.ReadLine()
    If newFile Then 
        outputFile.WriteLine lineBuffer
    End If

    Do While Not inputFile.AtEndOfStream  
        lineBuffer = inputFile.ReadLine
        lineBuffer = mid(lineBuffer,1,47) & "000000000000000000" & mid(lineBuffer,66,121) & "CDF" & mid(lineBuffer,190)
        outputFile.WriteLine lineBuffer
    Loop

    inputFile.Close
    outputFile.Close

    fso.DeleteFile inputFileName 