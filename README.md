# Dated File Downloader
Batch download files that have the date format of YYYYMMDD in them, from a specific starting date.

I needed something to download a lot of files by knowing the location and date of the file, rather than manually downloading them, or downloading via a browser.
This script let's you specify the start and finish date, the web address and the location of where the files should be.

Basically, if you have a thousand files like this:

```
/ server.com / path /
20200801.mp3
20200802.mp3
20200803.mp3
20200804.mp3
20200805.mp3
```

It uses Windows VBScrip, so wscript or cscript should work fine.
 
This so that you can download files with the data format of YYYYMMDD from a specific starting date.

## Instructions:

* Download curl from here for windows: https://curl.haxx.se/windows/
Get the binary version and the dependancies
Copy the contents of the /bin folder from both zip files to the root path of this script.
* Congifure the settings below.
* Run the file runme.vbs
* Execute process.bat



### The Batch file will look like this:

cURL_Command + locationsave + GENERATED_DATE + APPEND_URL + options + URL + GENERATED_DATE + APPEND_URL

```
curl  -o "files\DOG20200815.mp3" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:79.0) Gecko/20100101 Firefox/79.0" -H "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8" -H "Accept-Language: en-US,en;q=0.5" --compressed -H "Connection: keep-alive" -H "Upgrade-Insecure-Requests: 1" -H "Pragma: no-cache" -H "Cache-Control: no-cache" "https://s1.amazonaws.com/somethingtest/CAT/DOG20200815.mp3"
````


