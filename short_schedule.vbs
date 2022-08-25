Option Explicit

' Usage: $0 [days] [outFilename]
'    cscript.exe /NoLogo short_schedule.vbs -5 -
' �f�v���C��� ~/bin/ 

Const olFolderCalendar = 9
Const olFolderIndex = 6
Const olFolderManagedEmail = 29
Const olMeetingReceivedAndCanceled = 7

'�R�}���h���C�������i�p�����[�^�j�̎擾

Dim oParam
Set oParam = WScript.Arguments

'Dim idx
'For idx = 0 To oParam.Count - 1 
'  WScript.echo oParam(idx)
'Next

' �������i����΁j�́A����
Dim days
Dim stmCSVFile 	'As TextStream

' �f�t�H���g�̓����� ������14��
days = 14

If oParam.Count > 0 Then
   days = oParam(0)
End If

' 14- �Ə����Ă� -14 �Ɖ��߂���
If Right(days, 1) = "-" Then
   days = -1 * Left(days, Len(days)-1)
End If


' �������i����΁j�́A�����o���t�@�C�����B �W���o�͂��Ӗ����� - ���w��\
Dim strFileName
strFileName = "D:\tmp\sche.txt"

If oParam.Count > 1 Then
	strFileName = oParam(1)
End If

If strFileName = "-" Then
    Set stmCSVFile = WScript.StdOut

    '�X�N���v�g�E�z�X�g�̃t�@�C�������擾
    Dim strHostName
    strHostName = LCase(Mid(WScript.FullName,  InStrRev(WScript.FullName,"\") + 1))

    '�����z�X�g��wscript.exe�Ȃ�
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


' Main ���Ăяo��
Main days, stmCSVFile


'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
'  Main routine
'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
' Outlook�I�u�W�F�N�g������āA�������J�n

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
' Calendar �I�u�W�F�N�g���擾���ď������J�n
'''''''''''''''

Public Sub ShortScheduleDaysAP(ap, days, outStream)
    Dim myNamespace 	'As Outlook.NameSpace
    
    Set myNamespace = ap.GetNamespace("MAPI")
    
    Dim myCalendar 	'As Outlook.Folder
    
    Set myCalendar = myNamespace.GetDefaultFolder(olFolderCalendar)
    
    ShortScheduleDays myCalendar, days, outStream

End Sub


'''''''''''''''
' ������ 2359�̂悤�� ����4�����ɕϊ�
'''''''''''''''
Private Function Formathhnn(d)
    Dim hh, nn
    hh = Right("0" & Hour(d), 2)
    nn = Right("0" & Minute(d), 2)
    Formathhnn = hh & nn '  Format(d, "hhnn")
End Function

'''''''''''''''
' ���t�� 04/01�� �̂悤�ȕ�����ɕϊ�
'''''''''''''''
Private Function Formatmmdd(d)
    Dim mm, dd, aaa
    mm = Right("0" & Month(d), 2)
    dd = Right("0" & Day(d), 2)
    aaa = Mid("�����ΐ��؋��y", Weekday(d), 1)
    Formatmmdd = mm & "/" & dd & aaa ' Format(d, "mm/ddaaa")
End Function

'''''''''''''''
' �����̖{��
'''''''''''''''
' �w��̓����i���̎��͖����A���̂Ƃ��͉ߋ��j�̃X�P�W���[�����擾
' MeetingReceivedAndCanceled �̓X�L�b�v
' ������0000�̓X�L�b�v
' 
' ���t���ς��������t���o��
' �J�n����-�I������ �^�C�g��@�ꏊ ���o��
' �������A�ꏊ�́A�璷�ȃe�L�X�g���폜
' �Ō�ɏo�͐��close
' 

Public Sub ShortScheduleDays(fldCalendar, days, outStream)
    Dim strStart	'As String
    Dim strEnd 		'As String
    Dim dtExport 	'As Date
    Dim colAppts 	'As Items
    Dim objAppt 	'As AppointmentItem
    Dim strLine 	'As String
    '
    dtExport = Now ' �����̗\����G�N�X�|�[�g����ꍇ�� Now �̑���� DateAdd("m",1,Now) ���g�p���܂��B
    ' ���P�ʂł͂Ȃ��C�ӂ̒P�ʂɂ���ꍇ�͈ȉ��̋L�q��ύX���܂��B
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
'	       fromTo = .Organizer & "��" & .RequiredAttendees
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
      .Pattern = "���jICC��5F�j"
      shortLocation = .Replace(shortLocation, "")

      .Pattern = "���j"
      shortLocation = .Replace(shortLocation, "")
   End With
   shortLocationTruncate = shortLocation
End Function

Function shortLocationTruncateOld(shortLocation)
   Dim RE
   Set RE = CreateObject("VBScript.RegExp")
   With RE
'      .Pattern = "(.* |^)([^ ]+)(��c��|��c�R�[�i�[|��c��|�N���X�^).*"
'      shortLocation = .Replace(shortLocation, "$2")
      .Pattern = "(��c��|��c�R�[�i�[|��c��|�N���X�^)"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "(.* |^)([^ ]+)(���ڎ�)"
'      shortLocation = .Replace(shortLocation, "$2����")
      
      .Pattern = "���ڎ�"
      shortLocation = .Replace(shortLocation, "����")

      .Pattern = " [^ ]*��$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " [^ ]*�s��$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " [^ ]*�� "
      shortLocation = .Replace(shortLocation, " ")
      
      .Pattern = " [^ ]*�s�� "
      shortLocation = .Replace(shortLocation, " ")
      
'      .Pattern = "TV��c���p�s��"
'      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " �ڋq�Ή��p$"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*������\)"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*������\�j"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*��\)"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "^.*��\�j"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "������ "
'      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "���"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = "����"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " .�K"
      shortLocation = .Replace(shortLocation, "")
      
      .Pattern = " \d+�l"
      shortLocation = .Replace(shortLocation, "")
      
'      .Pattern = "^[^ ]*������"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^[^ ]*�����Z���^�["
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^[^ ]*������"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = "^(| )������"
'      shortLocation = .Replace(shortLocation, "")

'      .Pattern = " ��2"
'      shortLocation = .Replace(shortLocation, "")
'      
'      .Pattern = " 2����"
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

'    body �� head 30�s�����o���B
'    body �̈��A�A�����������i�ŏ�����܍s�ȓ��́j�����l�ł��B�i�����j�i��ςɁj�����b�Ɂi�Ȃ��Ă���܂��b�Ȃ�܂��j�B^.*��c�ł��B
'    body �̋�s�A���s������

'    WScript.echo mes.Body
    shortBody = MyLeftB(mes.Body, 200)
'    WScript.echo shortBody

    With RE
        
' ���s��󔒂�����
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

        .Pattern = "�@*" ' �S�p�󔒂͖������ŏ���
        shortBody = .Replace(shortBody, "")

' �O��ǂ��炩���p�����łȂ��Ȃ�󔒂�����
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

    ' shortBody =~ s/.* (.*)��c.*/$1/
'    messageLine =  leftB(strLine, 145)
'ErrMessageLineHandler:
'    On Error GoTo 0
	
'    WScript.echo strLine
    messageBody = strLine
'    WScript.echo strLine
End Function

''**��������*******************************
 ''* pS_String:���̑Ώە�����
''* pI_Len �@:�擪���甲���o���o�C�g��
''*****************************************
'Public Function MyLeftB(pS_String As String, pI_Len As Integer) As String
Public Function MyLeftB(pS_String, pI_Len )
 Dim S_Wkstring 'As String '��Ɨp������G���A

'On Error GoTo ErrHandler

 ''SJIS�ɕϊ����ALeftB�֐����s��
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

