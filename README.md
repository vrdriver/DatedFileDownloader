# DatedFileDownloader
Batch download files that have the date format of YYYYMMDD in them, from a specific starting date.

I needed something to download a lot of files by knowing the date of the file, rather than manually downloading them, or downloading via a browser.
This let's you specify the start and finish date, the web address and the location of where the files should be.


Basically, if you have a thousand files like this:

/ server.com / path /
20200801
20200802
20200803
20200804
20200805

You can save them on your computer.

It uses Windows VBScript. 

See the runme.vbs file for set up configuration and create an empty file called process.bat to run, which contains all of the download commands.
