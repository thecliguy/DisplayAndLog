Option Explicit

Sub DisplayAndLog(s_TextToWrite, b_DisplayText, s_LogFile, b_DisplayDateTime)
    '***************************************************************************
    ' DETAILS
    '   Copyright (C) 2019
    '   Adam Russell <adam[at]thecliguy[dot]co[dot]uk> 
    '   https://www.thecliguy.co.uk
    '   
    '   Licensed under the MIT License.
    '
    ' PURPOSE
    '   Outputs a given string to the console window (stdout) and/or to a log 
    '   file.
    '
    ' PARAMETERS
    '   All parameters are mandatory.
    '
    '   s_TextToWrite        A string denoting the text to be displayed and/or 
    '                        logged.
    '
    '   b_DisplayText        A boolean denoting whether the output is to be 
    '                        displayed to the console.
    '
    '   s_LogFile            A string containing the path to the log file.
    '                        If logging to a file is not required, then supply
    '                        an empty string.
    '
    '   b_DisplayDateTime    A boolean denoting whether the date/time should be 
    '                        included in any console output. Has no effect
    '                        unless used in conjunction with b_DisplayText.
    '
    ' EXAMPLE USAGE
    '   DisplayAndLog "Hello World", True, "C:\Bitbucket\LogFile.txt", True
    '
    '***************************************************************************
    ' DEVELOPMENT LOG
    '
    ' 0.1.0, 2019-12-21, Adam Russell
    '   * First release
    '
    '***************************************************************************
    
    ' OpenTextFile Method needs a Const value
    Const ForAppending = 8
    Const ForReading = 1
    Const ForWriting = 2
    Const TristateUseDefault = -2   ' Opens the file using the system default.
    Const TristateTrue = -1         ' Opens the file as Unicode.
    Const TristateFalse = 0         ' Opens the file as ASCII.
    
    Dim dtsNow, strDateIso8601, objFSO, objTextFile
    
    dtsNow = Now()

    ' The value returned by date/time component functions (Second, Minute, 
    ' Hour, Day, Month and Year) where less than 10 is not padded with a leading 
    ' zero. As a workaround, a zero is prepended to all values but is discarded
    ' (using the Right function) if the value already consists of two digits.
    strDateIso8601 = "[" & DatePart("yyyy",dtsNow) _
                     & "-" & Right("0" & DatePart("m",dtsNow), 2) _
                     & "-" & Right("0" & DatePart("d",dtsNow), 2) _
                     & " " & Right("0" & DatePart("h",dtsNow), 2) _
                     & ":" & Right("0" & DatePart("n",dtsNow), 2) _
                     & ":" & Right("0" & DatePart("s",dtsNow), 2) & "]"
    
    If b_DisplayText Then
        If b_DisplayDateTime Then
            WScript.Echo strDateIso8601 & " " & s_TextToWrite
        Else
            WScript.Echo s_TextToWrite
        End If
    End If
    
    If s_LogFile <> "" Then 
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objTextFile = objFSO.OpenTextFile(s_LogFile, ForAppending, True, TristateTrue)
        objTextFile.WriteLine strDateIso8601 & " " & s_TextToWrite
        objTextFile.Close
        Set objFSO = Nothing
        Set objTextFile = Nothing
    End If
End Sub
