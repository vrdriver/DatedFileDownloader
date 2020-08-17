' Stephen Monro 20200817
'
' Project: DatedFile Downloader
'
' This so that you can download files with the data format of YYYYMMDD from a specific starting date.
' This is a vbs script, designed to be run with wscript in Windows.
'
' Download curl from here for windows: https://curl.haxx.se/windows/
' Get the binary version and the dependancies
'
' Copy the contents of the /bin folder from both zip files to the root path of this script.
'
'
'
'
'
' The Batch file will look like this:
'
' cURL_Command + locationsave + GENERATED_DATE + APPEND_URL + options + URL + GENERATED_DATE + APPEND_URL
'
' curl  -o "files\DOG20200815.mp3" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:79.0) Gecko/20100101 Firefox/79.0" -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Connection: keep-alive" -H "Upgrade-Insecure-Requests: 1" -H "Pragma: no-cache" -H "Cache-Control: no-cache" "https://s1.amazonaws.com/somethingtest/CAT/DOG20200815.mp3"
'
'
' Configuration


strStartPath = 		               "C:\temp\script\"
strStartDrive = 		           "C:"	

StartDate = "20200815"
FinishDate = GetToday() ' "20200815"

cURL_SaveLocation = "files\DOG" ' make sure there is a path called 'files'. DOG is the prefix of the dated file.
' Path to harvest the files from.
URL = "https://s1.amazonaws.com/somethingtest/CAT/DOG"

' The end of the file name'
APPEND_URL = ".mp3"
 

' These are curl settings that will probably be need to let the server think it's got
' a normal web browser downloading files from it, rather than a script.

cURL_Command = "curl "
cURL_compiledlocation = " -o " & chr(34) & cURL_SaveLocation ' imagine the date goes after that
cURL_Options = cURL_Options & " -H " & chr(34) & "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:79.0) Gecko/20100101 Firefox/79.0" & chr(34)
cURL_Options = cURL_Options & " -H " & chr(34) & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" & chr(34) 
cURL_Options = cURL_Options & " -H " & chr(34) & "Accept-Language: en-US,en;q=0.5" & chr(34) 
cURL_Options = cURL_Options & " --compressed"
cURL_Options = cURL_Options & " -H " & chr(34) & "Connection: keep-alive" & chr(34)
cURL_Options = cURL_Options & " -H " & chr(34) & "Upgrade-Insecure-Requests: 1" & chr(34) 
cURL_Options = cURL_Options & " -H " & chr(34) & "Pragma: no-cache" & chr(34) 
cURL_Options = cURL_Options & " -H " & chr(34) & "Cache-Control: no-cache" & chr(34)



' NOTHING ELSE TO CONFIGURE

Const ForWriting = 2    
Set objFso = CreateObject("Scripting.FileSystemObject") 
Set objBatFile = objFso.OpenTextFile(strStartPath & "process.bat", ForWriting, True)
objBatFile.WriteLine "@echo Starting conversion"            
set args = WScript.Arguments

if StartDate = "" then
    StartDate = InputBox("Enter the start Date YYYYMMDD: ")     
end if

dim sStrDate 
aStrDate = StartDate

' Create the batch file contents
Do
    ' Add a line to the script
    objBatFile.WriteLine cURL_Command + cURL_compiledlocation + aStrDate + APPEND_URL + chr(34) + cURL_Options + " " +  chr(34) & URL + aStrDate + APPEND_URL & chr(34)
    ' Add another day
    aStrDate = addaday(aStrDate)
Loop Until aStrDate = FinishDate

' Execute the batch file and start the downloads 
    
objBatFile.WriteLine "@echo Finished..."
objBatFile.WriteLine "pause"



objBatFile.Close	
Set WshShell = WScript.CreateObject("WScript.Shell")
'intReturn = WshShell.Run(strStartPath & "\process.bat", 4, TRUE)
 


 ' **** functions 

  Function GetToday()
    Dim t 
    t = Now
    timeStamp = Year(t) & "" & _
    Right("0" & Month(t),2)  & "" & _
    Right("0" & Day(t),2)
    GetToday = timeStamp
End Function

function addaday(sStrDate) 
' VBScript has weird date formatting for it's dateadd function.
    yearPart = Mid(sStrDate, 1, 4)
    monthPart= MonthName(Mid(sStrDate, 5, 2), True) 
    dayPart = Mid(sStrDate, 7, 2) 
    sStrDate = DateAdd("d", 1, dayPart & "-" & monthPart & "-" & yearPart)

    if (len(DatePart("m", sStrDate)) < 2) then
        monthPart = "0" & CStr(DatePart("m", sStrDate))
    else
        monthPart = CStr(DatePart("m", sStrDate))
    end if

    if (len(DatePart("d", sStrDate)) < 2) then
        dayPart = "0" & CStr(DatePart("d", sStrDate))
    else
        dayPart = CStr(DatePart("d", sStrDate))
    end if

    yearPart = CStr(DatePart("yyyy", sStrDate))
    addaday = yearPart + monthPart + dayPart  
end function
