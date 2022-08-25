Option Explicit

' Usage: $0 [days] [outFilename]
'    cscript.exe /NoLogo short_schedule.vbs -5 -
' デプロイ先は ~/bin/ 

Const olFolderCalendar = 9
Const olFolderIndex = 6
Const olFolderManagedEmail = 29
Const olMeetingReceivedAndCanceled = 7

'コマンドライン引数（パラメータ）の取得

Dim oParam
Set oParam = WScript.Arguments

'Dim idx
'For idx = 0 To oParam.Count - 1 
'  WScript.echo oParam(idx)
'Next

' 第一引数（あれば）は、日数
Dim days
Dim stmCSVFile 	'As TextStream

' デフォルトの日数は 今から14日
days = 14

If oParam.Count > 0 Then
   days = oParam(0)
End If

' 14- と書いても -14 と解釈する
If Right(days, 1) = "-" Then
   days = -1 * Left(days, Len(days)-1)
End If


' 第二引数（あれば）は、書き出しファイル名。 標準出力を意味する - も指定可能
Dim strFileName
strFileName = "D:\tmp\sche.txt"

If oParam.Count > 1 Then
	strFileName = oParam(1)
End If

If strFileName = "-" Then
    Set stmCSVFile = WScript.StdOut

    'スクリプト・ホストのファイル名を取得
    Dim strHostName
    strHostName = LCase(Mid(WScript.FullName,  InStrRev(WScript.FullName,"\") + 1))

    'もしホストがwscript.exeなら
    If strHostName = "wscript.exe" Then
	WScript.Echo "Usage: cscript.exe $0 [days] [file]; file can be '-' only when executed from cscript.exe"
	WScript.Quit()
    End if

Else
    Dim objFSO 		'As FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set stmCSVFile = objFSO.CreateTextFile(strFileName, True)
End If

'stmCSVFile.WriteLine days


' Main を呼び出す
Main days, stmCSVFile


'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
'  Main routine
'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
' Outlookオブジェクトを作って、処理を開始

Public Sub Main(days, outStream)
    Dim OTApp 	'As Outlook.Application
'    If Process.GetProcessesByName("OUTLOOK").Count() > 0 Then
'        ' If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
'         Set OTApp = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application)
'    Else
        ' If not, create a new instance of Outlook and log on to the default profile.
	Set OTApp = CreateObject("Outlook.Application")
'        Dim ns	 'As Outlook.NameSpace
'        Set ns = OTApp.GetNamespace("MAPI")
'        ns.Logon "", "", Missing.Value, Missing.Value
'        Set ns = Nothing
'    End If

'	outStream.WriteLine days

	ShortScheduleDaysAP OTApp, days, outStream
End Sub

'''''''''''''''
' Calendar オブジェクトを取得して処理を開始
'''''''''''''''

Public Sub ShortScheduleDaysAP(ap, days, outStream)
    Dim myNamespace 	'As Outlook.NameSpace
    
    Set myNamespace = ap.GetNamespace("MAPI")
    
    Dim myCalendar 	'As Outlook.Folder
    
    Set myCalendar = myNamespace.GetDefaultFolder(olFolderCalendar)
    
    ShortScheduleDays myCalendar, days, outStream

End Sub


'''''''''''''''
' 時刻を 2359のような 数字4文字に変換
'''''''''''''''
Private Function Formathhnn(d)
    Dim hh, nn
    hh = Right("0" & Hour(d), 2)
    nn = Right("0" & Minute(d), 2)
    Formathhnn = hh & nn '  Format(d, "hhnn")
End Function

'''''''''''''''
' 日付を 04/01火 のような文字列に変換
'''''''''''''''
Private Function Formatmmdd(d)
    Dim mm, dd, aaa
    mm = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d), 2)
    aaa = Mid("日月火水木金土", Weekday(d), 1)
    Formatmmdd = mm & "/" & dd & aaa ' Format(d, "mm/ddaaa")
End Function

'''''''''''''''
' 処理の本体
'''''''''''''''
' 指定の日数（正の時は未来、負のときは過去）のスケジュールを取得
' MeetingReceivedAndCanceled はスキップ
' 時刻が0000はスキップ
' 
' 日付が変わったら日付を出力
' 開始時刻-終了時刻 タイトル@場所 を出力
' ただし、場所は、冗長なテキストを削除
' 最後に出力先をclose
' 

Public Sub ShortScheduleDays(fldCalendar, days, outStream)
    Dim strStart	'As String
    Dim strEnd 		'As String
    Dim dtExport 	'As Date
    Dim colAppts 	'As Items
    Dim objAppt 	'As AppointmentItem
    Dim strLine 	'As String
    '
    dtExport = Now ' 来月の予定をエクスポートする場合は Now の代わりに DateAdd("m",1,Now) を使用します。
    ' 月単位ではなく任意の単位にする場合は以下の記述を変更します。
    If days > 0 Then
       strStart = Year(Now) & "/" & Month(Now) & "/" & Day(Now) & " 00:00"
       strEnd = DateAdd("d", days, CDate(strStart)) & " 00:00"
    Else
       strEnd = Year(Now) & "/" & Month(Now) & "/" & Day(Now) & " 00:00"
       strStart = DateAdd("d", days, CDate(strEnd)) & " 00:00"
    End If

'    outStream.WriteLine strEnd
'    outStream.WriteLine strStart

    '
     Set colAppts = fldCalendar.Items
    colAppts.Sort "[Start]"
    colAppts.IncludeRecurrences = True
    Set objAppt = colAppts.Find("[Start] < """ & strEnd & """ AND [End] >= """ & strStart & """")
    Dim mmdd 		'As String
    Dim mmddOld 	'As String
    mmddOld = ""
    'While Not objAppt Is Nothing
    While TypeName(objAppt) <> "Nothing"
        If objAppt.MeetingStatus = olMeetingReceivedAndCanceled Or _
            Formathhnn(objAppt.Start) = "0000" _
        Then
        Else
        
            mmdd = Formatmmdd(objAppt.Start)
            If mmdd <> mmddOld Then
                outStream.WriteLine mmdd
                mmddOld = mmdd
            End If
        
            strLine = " "
            strLine = strLine & Formathhnn(objAppt.Start)
            strLine = strLine & "-"
            strLine = strLine & Formathhnn(objAppt.End)
            strLine = strLine & " "
            strLine = strLine & objAppt.Subject

	    Dim shortLocation
            shortLocation = objAppt.Location
	    shortLocation = shortLocationTruncate(shortLocation)

            If shortLocation <> "" Then
                strLine = strLine & "@"
                strLine = strLine & shortLocation
            End If

	    Dim fromTo
	    With objAppt
'	       fromTo = .Organizer & "⇒" & .RequiredAttendees
	       fromTo = .RequiredAttendees
	    End With
	    
	    Dim RE
	    Set RE = CreateObject("VBScript.RegExp")
            With RE
	       .Global = True
               .Pattern = "[a-zA-Z, ]*/"
               fromTo = .Replace(fromTo, "")
	    End With
	    strLine = strLine & ":"
	    strLine = strLine & fromTo

	    strLine = strLine & messageBody(objAppt)

	    strLine = MyLeftB(strLine, 160)
            outStream.WriteLine strLine
        End If
        Set objAppt = colAppts.FindNext
    Wend
    outStream.Close
End Sub

Function shortLocationTruncate(shortLocation)
   Dim RE
   Set RE = CreateObject("VBScript.RegExp")
   With RE
      .Pattern = "川崎）ICC棟5F）"
      shortLocation = .Replace(shortLocation, "")

      .Pattern = "川崎）"
      shortLocation = .Replace(shortLocation, "")
   End With
   shortLocationTruncate = shortLocation
End Function

Function shortLocationTruncateOld(shortLocation)
   Dim RE
   Set RE = CreateObject("VBScript.RegExp")
   With RE
'      .Pattern = "(.* |^)([^ ]+)(会議室|会議コーナー|会議卓|クラスタ).*"
'      shortLocation = .Replace(shortLocation, "$2")
      .Pattern = "(会議室|会議コーナー|会議卓|クラスタ)"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "(.* |^)([^ ]+)(応接室)"
'      shortLocation = .Replace(shortLocation, "$2応接")
      
      .Pattern = "応接室"
      shortLocation = .Replace(shortLocation, "応接")

      .Pattern = " [^ ]*可$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " [^ ]*不可$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " [^ ]*可 "
      shortLocation = .Replace(shortLocation, " ")
      
      .Pattern = " [^ ]*不可 "
      shortLocation = .Replace(shortLocation, " ")
      
'      .Pattern = "TV会議利用不可"
'      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " 顧客対応用$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*研究所\)"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*研究所\）"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*研\)"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*研\）"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "研究所 "
'      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "川崎"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "共通"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " .階"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " \d+人"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "^[^ ]*研究部"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^[^ ]*研究センター"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^[^ ]*研究所"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^(| )研究所"
'      shortLocation = .Replace(shortLocation, "")

'      .Pattern = " 研2"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = " 2号館"
'      shortLocation = .Replace(shortLocation, "")
      

      .Pattern = "[<(][-a-zA-Z0-9@\.]+[)>]"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^ +"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " +$"
      shortLocation = .Replace(shortLocation, "")
      
   End With

   shortLocationTruncate = shortLocation
End Function

Function messageBody(mes)
    Dim strLine     'As String

'    On Error GoTo ErrMessageLineHandler
'    WScript.echo strLine
    Dim RE
    Set RE = CreateObject("VBScript.RegExp")

    Dim shortBody

'    body の head 30行を取り出す。
'    body の挨拶、名乗りを除く（最初から五行以内の）お疲れ様です。（いつも）（大変に）お世話に（なっております｜なります）。^.*上田です。
'    body の空行、改行を除く

'    WScript.echo mes.Body
    shortBody = MyLeftB(mes.Body, 200)
'    WScript.echo shortBody

    With RE
        
' 改行や空白を除去
        .Global = True
        
        .Pattern = "-----*"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "=====*"
        shortBody = .Replace(shortBody, "")

        .Pattern = "_____*"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "\n"
        shortBody = .Replace(shortBody, "")
        
        .Pattern = "\r"
        shortBody = .Replace(shortBody, "")
        
'        .Pattern = " *"
'        shortBody = .Replace(shortBody, "")

        .Pattern = "　*" ' 全角空白は無条件で除く
        shortBody = .Replace(shortBody, "")

' 前後どちらかが英数字でないなら空白を除く
        .Pattern = "([^0-z]) *"
        shortBody = .Replace(shortBody, "$1")
        
        .Pattern = " *([^0-z])"
        shortBody = .Replace(shortBody, "$1")
        
        .Pattern = "\t"
        shortBody = .Replace(shortBody, "")
                
    End With
    
'    WScript.echo shortBody
    If shortBody <> "" Then
        strLine = strLine & "|"
        strLine = strLine & shortBody
    End If

    ' shortBody =~ s/.* (.*)会議.*/$1/
'    messageLine =  leftB(strLine, 145)
'ErrMessageLineHandler:
'    On Error GoTo 0
	
'    WScript.echo strLine
    messageBody = strLine
'    WScript.echo strLine
End Function

''**引数説明*******************************
 ''* pS_String:元の対象文字列
''* pI_Len 　:先頭から抜き出すバイト数
''*****************************************
'Public Function MyLeftB(pS_String As String, pI_Len As Integer) As String
Public Function MyLeftB(pS_String, pI_Len )
 Dim S_Wkstring 'As String '作業用文字列エリア

'On Error GoTo ErrHandler

 ''SJISに変換し、LeftB関数を行う
 Dim bobj
 Set bobj = CreateObject("basp21")
' S_Wkstring = LeftB(pS_String, pI_Len)
' MyLeftB = S_Wkstring

 Dim kcode
 kcode = bobj.KConv(pS_String, 0)
 MyLeftB = pS_String
 ' WScript.echo kcode  ' UNICODE UCS2 = 4
 MyLeftB = bobj.KConv(MyLeftB, 1, kcode)
 MyLeftB = bobj.MidB(MyLeftB, 0, pI_Len)
 MyLeftB = bobj.KConv(MyLeftB, kcode, 1)
 
' S_Wkstring = LeftB(bobj.KConv(pS_String, 1), pI_Len)
' S_Wkstring = bobj.MidB(pS_String, 1, pI_Len)

'  MyLeftB = bobj.KConv(S_Wkstring, kcode, 1)
'  MyLeftB = CStr(bobj.KConv(S_Wkstring, kcode, 1))
 
'S_Wkstring = LeftB(StrConv(pS_String, vbFromUnicode), pI_Len)
' MyLeftB = StrConv(S_Wkstring, vbUnicode)

' On Error GoTo 0
' Exit Function

'ErrHandler:
' MyLeftB = ""
' On Error GoTo 0
End Function

