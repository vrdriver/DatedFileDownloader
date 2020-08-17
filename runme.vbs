' Stephen Monro 20200817
'
' Project: DatedFile Downloader
'
' This so that you can download files with the data format of YYYYMMDD from a specific starting date.
' This is a vbs script, designed to be run with wscript in Windows.
'
' Download wget from here for windows: http://gnuwin32.sourceforge.net/packages/wget.htm
' Get the binary version and the dependancies
'
' Binaries	        Zip	 		850448	 	31 December 2008		 	254b95bd96564eb6db590f2b51f8fd8b
' http://downloads.sourceforge.net/gnuwin32/wget-1.11.4-1-bin.zip
'
' Dependencies      Zip	 		1443871	 	4 September 2008		 	b23ce98bc2d03f852f8516183c619d42
' http://downloads.sourceforge.net/gnuwin32/wget-1.11.4-1-dep.zip
'
' Copy the contents of the /bin folder from both zip files to the root path of this script.
'
'
'
'
'
' Examples:
'
' URL + GENERATED_DATE + APPEND_URL'
' http://domain.com/20200817.mp3
' 
'
' The Batch file will look like this:
'
' WGET_Command + " " + URL + GENERATED_DATE + APPEND_URL
' wget http://domain.com/20200817.mp3
'
'
'

' Configuration
WGET_Command = "wget "
URL = "http://domain.com/path/"
APPEND_URL = ".mp3"

strStartPath = 		               "C:\temp\date\ravi\"
strStartDrive = 		           "C:"	


' NOTHING ELSE TO CONFIGURE

Const ForWriting = 2    
Set objFso = CreateObject("Scripting.FileSystemObject") 
Set objBatFile = objFso.OpenTextFile(strStartPath & "process.bat", ForWriting, True)
objBatFile.WriteLine "@echo Starting conversion"            
set args = WScript.Arguments

select case args.Count
case 0
    'FirstDate = GetDateString()	 
    FirstDate = InputBox("Enter the start Date YYYYMMDD: ") 
end select

    
dim sStrDate 
aStrDate = FirstDate

' Create the batch file contents
Do
    objBatFile.WriteLine WGET_Command + " " + URL + aStrDate + APPEND_URL
    aStrDate = addaday(aStrDate)
Loop Until  aStrDate = GetToday()

' Execute the batch file and start the downloads 
    
objBatFile.WriteLine "@echo Finished..."
'   objBatFile.WriteLine "pause"
objBatFile.Close	
Set WshShell = WScript.CreateObject("WScript.Shell")
intReturn = WshShell.Run(strStartPath & "\process.bat", 4, TRUE)



 
  

  Function GetToday()
    Dim t 
    t = Now
    timeStamp = Year(t) & "" & _
    Right("0" & Month(t),2)  & "" & _
    Right("0" & Day(t),2)

    GetToday = timeStamp
End Function




function addaday(sStrDate) 

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
