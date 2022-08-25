One VBS script collaborating with Outlook.
- A VBS script to dump appointments into a text file. The number of days before or after today can be specified as command parameter.

Usage: short_schedule.vbs [days] [outFilename]
  E.g.    cscript.exe /NoLogo short_schedule.vbs -5 -

The VBS scripts can be executed from the task scheduler.

Depending libraries:
- BASP21 http://www.hi-ho.ne.jp/~babaq/basp21.html
  MidB is used to truncate Kanji string into a fixed length.
  If any problem, comment out all lines in MyLeftB function except for
     MyLeftB = pS_String
