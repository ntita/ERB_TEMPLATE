Attribute VB_Name = "Main"
Private pInSheet As Worksheet
Private pOutSheet As Worksheet
Private pTLItemList As Collection
Public CurDate As Date

Public Property Get mainUTC() As Integer
    If pmainUTC = 0 Then
        Set EMChatHeader = InSheet.Cells.Find(What:="Timeline from CSDP", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
        If EMChatHeader Is Nothing Then
            MsgBox "Timeline from CSDP not found. The program expects a column with a three-line header, the first of which contains 'Timeline from CSDP'.", vbCritical, "Error"
            End
        End If
        pmainUTC = Conversion.CInt(Right(InSheet.Cells(EMChatHeader.Row + 2, EMChatHeader.Column).Value, 3))
    End If
    mainUTC = pmainUTC
End Property
Public Property Get TLItemList() As Collection
    If pTLItemList Is Nothing Then
        Set pTLItemList = New Collection
    End If
    Set TLItemList = pTLItemList
End Property
Public Property Get InSheet() As Worksheet
    If pInSheet Is Nothing Then
        For Each sh In ActiveWorkbook.Worksheets
            If sh.Name = "Input list" Then
                Set pInSheet = sh
                Exit For
            End If
        Next sh
        If pInSheet Is Nothing Then
            MsgBox "Input list not found. It should be called 'Input list'", vbCritical, "Error"
            End
        End If
    End If
    Set InSheet = pInSheet
End Property
Public Property Get OutSheet() As Worksheet
    If pOutSheet Is Nothing Then
        For Each sh In ActiveWorkbook.Worksheets
            If sh.Name = "Prepared timeline output" Then
                Set pOutSheet = sh
                Exit For
            End If
        Next sh
        If pOutSheet Is Nothing Then
            MsgBox "Output list not found. It should be called 'Prepared timeline output'", vbCritical, "Error"
            End
        End If
    End If
    Set OutSheet = pOutSheet
End Property
Sub readCSDP()
    Application.StatusBar = "reading CSDP: processing..."
    Dim CSDPHeader As Range
    Dim DateOfEvent As Range
    Dim TLItem As TimelineItem
    Dim CurCell As Range
    Dim timeStamp As Date
    Dim timeSubStrLen As Date
    Set CSDPHeader = InSheet.Cells.Find(What:="Timeline from CSDP", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If CSDPHeader Is Nothing Then
        MsgBox "Timeline from CSDP not found. The program expects a column with a three-line header, the first of which contains 'Timeline from CSDP'.", vbCritical, "Error"
        End
    End If
    Set DateOfEvent = InSheet.Cells.Find(What:="Date of event", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If DateOfEvent Is Nothing Then
        MsgBox "Date of event not found. The program expects a table with a three-line header and one row. That row have to contain date.", vbCritical, "Error"
        End
    End If
    CurDate = InSheet.Cells(DateOfEvent.Row + 3, DateOfEvent.Column).Value
    For ind = CSDPHeader.Row + 3 To 1000000
        Set CurCell = InSheet.Cells(ind, CSDPHeader.Column)
        If CurCell.Value = Constants.vbNullString Then
            ind = CurCell.End(xlDown).Row - 1
        Else
            If (CurCell.Value Like "##:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                Or CurCell.Value Like "#:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                Or CurCell.Value Like "?[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                Or CurCell.Value Like "##:## [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *" _
                Or CurCell.Value Like "#:## [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *" _
                Or CurCell.Value Like "? [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *") _
            Then
                If CurCell.Value Like "##:## *" _
                   Or CurCell.Value Like "#:## *" _
                   Or CurCell.Value Like "? *" _
                Then
                    Let SpaceLen = 1
                Else
                    Let SpaceLen = 0
                End If
                If Not (TLItem Is Nothing) Then
                    timeStamp = TLItem.timeStamp
                End If
                Set TLItem = New TimelineItem
                If (CurCell.Value Like "##:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                    Or CurCell.Value Like "##:## [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *") _
                Then
                    timeSubStrLen = 5
                    TLItem.timeStamp = Left(CurCell.Value, timeSubStrLen)
                    TLItem.timeStamp = TLItem.timeStamp + CurDate
                ElseIf (CurCell.Value Like "#:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                        Or CurCell.Value Like "#:## [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *") _
                Then
                    timeSubStrLen = 4
                    TLItem.timeStamp = Left(CurCell.Value, timeSubStrLen)
                    TLItem.timeStamp = TLItem.timeStamp + CurDate
                ElseIf (CurCell.Value Like "?[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                        Or CurCell.Value Like "? [A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z] *") _
                Then
                    timeSubStrLen = 1
                    TLItem.timeStamp = timeStamp
                Else
                    decision = MsgBox("Line found that does not match any of the CSDP-patterns and will be skipped. Abort script (ignore warning, keep execute)?", vbYesNo + vbDefaultButton2, "Error")
                    If decision = vbYes Then
                        End
                    End If
                End If
                TLItem.authorName = Mid(CurCell.Value, timeSubStrLen + 1 + SpaceLen, 7)
                TLItem.cellAddress = CurCell.Row & "," & CurCell.Column
                TLItem.chatType = 1
                TLItem.mvalue = Mid(CurCell.Value, timeSubStrLen + 1 + 7 + 2 * SpaceLen)
                If TLItem.timeStamp < timeStamp Then
                    TLItem.timeStamp = TLItem.timeStamp + 1
                    CurDate = CurDate + 1
                End If
                TLItemList.Add TLItem
            End If
        End If
    Next ind
    Application.StatusBar = "reading CSDP: done"
End Sub
Private Sub readEMChat(chatName As String, chatType As Integer, afterCell As Range)
    Dim EMChatHeader As Range
    Dim CurCell As Range
    Dim timeLen As String
    Dim TLItem As TimelineItem
    Set EMChatHeader = InSheet.Cells.Find(What:=chatName, After:=afterCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If EMChatHeader Is Nothing Then
        MsgBox "Chat '" & chatName & "' not found.", vbCritical, "Error"
        'End Sub
    End If
    If IsDate(InSheet.Cells(EMChatHeader.Row + 1, EMChatHeader.Column).Value) Then
        CurDate = InSheet.Cells(EMChatHeader.Row + 1, EMChatHeader.Column).Value
    End If
    'EM chat #
    EMUTC = -Right(InSheet.Cells(EMChatHeader.Row + 2, EMChatHeader.Column).Value, 3) + mainUTC
    For ind = EMChatHeader.Row + 3 To 1000000
        Set CurCell = InSheet.Cells(ind, EMChatHeader.Column)
        If CurCell.Value = Constants.vbNullString Then
            ind = CurCell.End(xlDown).Row - 1
        Else
            If (Trim(CurCell.Value) Like "*##:## ??:" _
                Or Trim(CurCell.Value) Like "*#:## ??:" _
                Or Trim(CurCell.Value) Like "*##:##:" _
                Or Trim(CurCell.Value) Like "*#:##:") _
            Then
                If (Trim(CurCell.Value) Like "*##:## ??:") _
                Then
                    timeLen = 8
                ElseIf (Trim(CurCell.Value) Like "*#:## ??:") _
                Then
                    timeLen = 7
                ElseIf (Trim(CurCell.Value) Like "*##:##:") _
                Then
                    timeLen = 5
                ElseIf (Trim(CurCell.Value) Like "*#:##:") _
                Then
                    timeLen = 4
                End If
                Set TLItem = New TimelineItem
                TLItem.authorName = Trim(Left(Trim(CurCell.Value), Len(Trim(CurCell.Value)) - 1 - timeLen))
                TLItem.cellAddress = CurCell.Row & "," & CurCell.Column
                TLItem.chatType = chatType
                TLItem.mvalue = ""
                TLItem.timeStamp = Left(Right(Trim(CurCell.Value), 1 + timeLen), timeLen)
                TLItem.timeStamp = TLItem.timeStamp + CurDate
                TLItem.timeStamp = DateAdd("h", EMUTC, TLItem.timeStamp)
                If TLItem.timeStamp < timeStamp Then
                    TLItem.timeStamp = TLItem.timeStamp + 1
                    CurDate = CurDate + 1
                End If
                TLItemList.Add TLItem
                timeStamp = TLItem.timeStamp
            End If
        End If
    Next ind
End Sub
Sub readMainEMChats()
    Application.StatusBar = "reading main chats: processing..."
    Dim MainEMChatHeader As Range
    Dim CurMainChatCell As Range
    Dim EMChatHeader As Range
    Dim CurCell As Range
    Dim timeLen As String
    Dim TLItem As TimelineItem
    Set MainEMChatHeader = InSheet.Cells.Find(What:="Main EM chat", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If MainEMChatHeader Is Nothing Then
        MsgBox "Table of main chats not found. The program expects a table with a three-line header, the first of which equals 'Main EM chat'.", vbCritical, "Error"
        End
    End If
    For ind = MainEMChatHeader.Row + 3 To MainEMChatHeader.End(xlDown).End(xlDown).End(xlDown).Row
        Set CurMainChatCell = InSheet.Cells(ind, MainEMChatHeader.Column)
        If CurMainChatCell.Value = Constants.vbNullString Then
            Exit For
        End If
        readEMChat CurMainChatCell.Value, 2, CurMainChatCell
    Next ind
    Application.StatusBar = "reading main chats: done"
End Sub
Sub readAdditionalEMChats()
    Application.StatusBar = "reading additional chat: processing..."
    Dim MainEMChatHeader As Range
    Dim CurMainChatCell As Range
    Dim EMChatHeader As Range
    Dim CSDPHeader As Range
    Dim CurCell As Range
    Dim timeLen As String
    Dim TLItem As TimelineItem
    Dim n As Integer
    Set CSDPHeader = InSheet.Cells.Find(What:="Timeline from CSDP", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    n = 2
    Set MainEMChatHeader = InSheet.Cells.Find(What:="Main EM chat", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If MainEMChatHeader Is Nothing Then
        MsgBox "Table of main chats not found. The program expects a table with a three-line header, the first of which equals 'Main EM chat'.", vbCritical, "Error"
        End
    End If
    For jnd = 1 To 1000
        Set CurCell = InSheet.Cells(CSDPHeader.Row, jnd)
        If CurCell.Value Like "EM chat*" Then
            isMain = False
            For ind = MainEMChatHeader.Row + 3 To MainEMChatHeader.End(xlDown).End(xlDown).End(xlDown).Row
                Set CurMainChatCell = InSheet.Cells(ind, MainEMChatHeader.Column)
                If CurMainChatCell.Value = Constants.vbNullString Then
                    Exit For
                End If
                If CurMainChatCell.Value = CurCell.Value Then
                    isMain = True
                    Exit For
                End If
            Next ind
            If Not isMain Then
                n = n + 1
                readEMChat CurCell.Value, n, CurCell
            End If
        End If
    Next jnd
    Application.StatusBar = "reading additional chat: done"
End Sub
Private Sub sortTimeLine(CSDPTimeline As Range)
    ' sorting by time
    Application.StatusBar = "sorting timeline: processing..."
    OutSheet.Sort.SortFields.Clear
    OutSheet.Sort.SortFields.Add Key:=Range(OutSheet.Cells(CSDPTimeline.Row + 2, CSDPTimeline.Column), _
        OutSheet.Cells(CSDPTimeline.Row + 2, CSDPTimeline.Column).End(xlDown)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With OutSheet.Sort
        .SetRange Range(OutSheet.Cells(CSDPTimeline.Row + 2, CSDPTimeline.Column - 1), OutSheet.Cells(CSDPTimeline.Row + 2, CSDPTimeline.Column).End(xlDown))
        .Apply
    End With
    OutSheet.Sort.SortFields.Clear
    Application.StatusBar = "sorting timeline: done"
End Sub
Private Sub renderCSDP()
    ActiveCell.MergeArea.Cells.Count
End Sub
Private Sub multyplySampleRow(CSDPTimeline As Range)
    ' prepare space by multiplying sample string
    Dim sum As Long
    Dim TLItem As TimelineItem
    Dim AdditionalEMChatHeader As Range
    Application.StatusBar = "preparing space by multiplying sample string: processing..."
    sum = 0
    For ind = 1 To TLItemList.Count
        OutSheet.Rows((CSDPTimeline.Row + 3) & ":" & (CSDPTimeline.Row + 2 + 2 ^ (ind - 1))).Copy
        OutSheet.Cells(CSDPTimeline.Row + 3, 1).Insert shift:=xlDown
        sum = sum + 2 ^ (ind - 1)
        If sum >= TLItemList.Count Then
            Exit For
        End If
    Next ind
    ' delete excess strings
    If sum - TLItemList.Count > 3 Then
        OutSheet.Rows((CSDPTimeline.Row + TLItemList.Count + 2) & ":" & (CSDPTimeline.Row + 1 + sum)).Delete shift:=xlUp
    End If
    'For Each TLItem In TLItemList
    '    If TLItem.chatType > 4 Then
    '        Set AdditionalEMChatHeader = OutSheet.Cells.Find(What:="Additional EM chat #2", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
    '            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    '            MatchCase:=False, SearchFormat:=False)
    '        OutSheet.Columns((AdditionalEMChatHeader.Column) & ":" & (AdditionalEMChatHeader.Column + AdditionalEMChatHeader.MergeArea.Columns.Count - 1)).Copy
    '        OutSheet.Cells(1, AdditionalEMChatHeader.Column + AdditionalEMChatHeader.MergeArea.Columns.Count).Insert shift:=xlRight
    '    End If
    'Next TLItem
    Application.StatusBar = "preparing space by multiplying sample string: done"
End Sub
Private Sub printUnsortedTimeline(CSDPTimeline As Range)
    ' print unsorted info
    Application.StatusBar = "printing unsorted info: processing..."
    For ind = 1 To TLItemList.Count
        'Exit For
        OutSheet.Cells(CSDPTimeline.Row + 2 + ind - 1, CSDPTimeline.Column).Value = TLItemList(ind).timeStamp
        OutSheet.Cells(CSDPTimeline.Row + 2 + ind - 1, CSDPTimeline.Column - 1).Value = ind
    Next ind
    Application.StatusBar = "printing unsorted info: done"
End Sub
Private Sub deleteOldTimeline(CSDPTimeline As Range)
    Application.StatusBar = "deleting old timeline: processing..."
    OutSheet.Rows((CSDPTimeline.Row + 2) & ":" & OutSheet.Cells(1000000, CSDPTimeline.Column).End(xlUp).Row).Delete shift:=xlUp
    Application.StatusBar = "deleting old timeline: done"
End Sub
Sub reprintTimeLine()
    Dim myCell As Range
    Dim CSDPTimeline As Range
    Dim CSDPReportedBy As Range
    Dim CSDPMessage As Range
    Dim EMChats() As Range
    Dim MainEMChatTimeline As Range
    Dim MainEMChatMessage As Range
    Dim AdditionalCount As Long
    AdditionalCount = 0
    Set CSDPTimeline = OutSheet.Cells.Find(What:="CSDP timeline", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If CSDPTimeline Is Nothing Then
        MsgBox "Header cell for CSDP timeline not found. The program expects a cell equals 'CSDP timeline'.", vbCritical, "Error"
        End
    End If
    CDSPWide = CSDPTimeline.MergeArea.Cells.Columns.Count
    For Each myCell In OutSheet.Range(OutSheet.Cells(CSDPTimeline.Row + 1, CSDPTimeline.Column), OutSheet.Cells(CSDPTimeline.Row + 1, CSDPTimeline.Column + CDSPWide - 1))
        If UCase(myCell.Value) Like "*REPORT*" Then
            Set CSDPReportedBy = myCell
            Exit For
        End If
    Next myCell
    If CSDPReportedBy Is Nothing Then
        MsgBox "Header cell for CSDP ReportedBy not found. The program expects a cell containing something with 'Report' on the next line after the header cell CSDP timeline.", vbCritical, "Error"
        End
    End If
    For Each myCell In OutSheet.Range(OutSheet.Cells(CSDPTimeline.Row + 1, CSDPTimeline.Column), OutSheet.Cells(CSDPTimeline.Row + 1, CSDPTimeline.Column + CDSPWide - 1))
        If UCase(myCell.Value) Like "*MESSAGE*" Then
            Set CSDPMessage = myCell
            Exit For
        End If
    Next myCell
    If CSDPMessage Is Nothing Then
        MsgBox "Header cell for CSDP Message not found. The program expects a cell containing something with 'Message' on the next line after the header cell CSDP timeline.", vbCritical, "Error"
        End
    End If
    If OutSheet.Cells(CSDPTimeline.Row + 2, CSDPTimeline.Column).Value <> Constants.vbNullString Then
        deleteOldTimeline CSDPTimeline
    End If
    Set MainEMChatTimeline = OutSheet.Cells.Find(What:="Main EM chat", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If MainEMChatTimeline Is Nothing Then
        MsgBox "Header cell for Main EM chat timeline not found. The program expects a cell equals 'Main EM chat'.", vbCritical, "Error"
        End
    End If
    MainEMChatWide = MainEMChatTimeline.MergeArea.Cells.Columns.Count
    For Each myCell In OutSheet.Range(OutSheet.Cells(MainEMChatTimeline.Row + 1, MainEMChatTimeline.Column), OutSheet.Cells(MainEMChatTimeline.Row + 1, MainEMChatTimeline.Column + MainEMChatWide - 1))
        If UCase(myCell.Value) Like "*MESSAGE*" Then
            Set MainEMChatMessage = myCell
            Exit For
        End If
    Next myCell
    If MainEMChatMessage Is Nothing Then
        MsgBox "Header cell for Main EM chat Message not found. The program expects a cell containing something with 'Message' on the next line after the header cell Main EM chat.", vbCritical, "Error"
        End
    End If
    ReDim Preserve EMChats(2 To 256, 0 To 1) As Range
    Set EMChats(2, 0) = MainEMChatTimeline
    Set EMChats(2, 1) = MainEMChatMessage
    For ind = EMChats(2, 0).Column + 1 To 10000
        If UCase(OutSheet.Cells(EMChats(2, 0).Row, ind)) Like "*ADDITIONAL EM CHAT*" Then
            AdditionalCount = AdditionalCount + 1
            Set EMChats((2 + AdditionalCount), 0) = OutSheet.Cells(EMChats(2, 0).Row, ind)
            AdditionalEMChatWide = EMChats((2 + AdditionalCount), 0).MergeArea.Cells.Columns.Count
            For Each myCell In OutSheet.Range(OutSheet.Cells(EMChats((2 + AdditionalCount), 0).Row + 1, EMChats((2 + AdditionalCount), 0).Column), OutSheet.Cells(EMChats((2 + AdditionalCount), 0).Row + 1, EMChats((2 + AdditionalCount), 0).Column + AdditionalEMChatWide - 1))
                If UCase(myCell.Value) Like "*MESSAGE*" Then
                    Set EMChats((2 + AdditionalCount), 1) = myCell
                    Exit For
                End If
            Next myCell
        End If
    Next ind
    Application.ScreenUpdating = False
    multyplySampleRow CSDPTimeline
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    printUnsortedTimeline CSDPTimeline
    Application.ScreenUpdating = True
    Application.ScreenUpdating = False
    sortTimeLine CSDPTimeline
    Application.ScreenUpdating = True
    Application.StatusBar = "rendering timeline: processing..."
    Application.ScreenUpdating = False
    ' rendering
    Dim a() As String
    Dim tvalue As Variant
    ' fill chat timeline
    For ind = CSDPTimeline.Row + 2 To 1000000
        Set myCell = OutSheet.Cells(ind, CSDPTimeline.Column - 1)
        If myCell.Value = Constants.vbNullString Then
            Exit For
        End If
        Set TLItem = TLItemList(myCell.Value)
        If TLItem.chatType = 1 Then
            OutSheet.Cells(myCell.Row, CSDPReportedBy.Column) = TLItem.authorName
            OutSheet.Cells(myCell.Row, CSDPMessage.Column) = TLItem.mvalue
            myCell.Value = Null
        ElseIf TLItem.chatType > 1 Then
            OutSheet.Cells(myCell.Row, CSDPTimeline.Column).Value = Null
            a = Split(TLItem.cellAddress, ",", 2)
            OutSheet.Cells(myCell.Row, EMChats(TLItem.chatType, 0).Column).Value = InSheet.Cells(Conversion.CInt(a(0)), Conversion.CInt(a(1))).Value
            myCell.Value = Null
            OutSheet.Cells(myCell.Row, EMChats(TLItem.chatType, 1).Column).Value = InSheet.Cells(Conversion.CInt(a(0)) + 1, Conversion.CInt(a(1))).Value
            j = 0
            For i = 2 To 50
                tvalue = InSheet.Cells(Conversion.CInt(a(0)) + i, Conversion.CInt(a(1))).Value
                If (Trim(tvalue) Like "*##:## ??:" _
                    Or Trim(tvalue) Like "*#:## ??:" _
                    Or Trim(tvalue) Like "*##:##:" _
                    Or Trim(tvalue) Like "*#:##:") Then
                    Exit For
                End If
                If (Trim(tvalue) = Constants.vbNullString) Then
                    j = j - 1
                Else
                    OutSheet.Cells(myCell.Row + i + j - 1, EMChats(TLItem.chatType, 1).Column).EntireRow.Insert
                    ind = ind + 1
                    OutSheet.Cells(myCell.Row + i + j - 1, EMChats(TLItem.chatType, 1).Column).Value = InSheet.Cells(Conversion.CInt(a(0)) + i, Conversion.CInt(a(1))).Value
                End If
            Next i
        End If
    Next ind
    Application.ScreenUpdating = True
    Application.StatusBar = "rendering timeline: done"
End Sub
Sub Main()
    Dim myRange As Range
    Dim TLItem As TimelineItem
    Dim myCell As Range
    Dim timeStamp As Variant
    Dim EMChat As Range
    Dim EMChatRange As Range
    Let dbgmode = 0
    Set TLItem = New TimelineItem
    
    readCSDP
    readMainEMChats
    readAdditionalEMChats
    OutSheet.Activate
    reprintTimeLine
    
    For ind = 1 To TLItemList.Count
        If TLItemList(ind).chatType <> 1 Then
            For j = 1 To 1000
                If j + 3 = OutSheet.Cells(3, 9).End(xlDown).Row Then
                    OutSheet.Cells(OutSheet.Cells(3, 9).End(xlDown).Row + 1, 1).EntireRow.Insert
                    OutSheet.Range(OutSheet.Cells(3 + j, 3), OutSheet.Cells(3 + j, 9)).Copy
                    OutSheet.Cells(OutSheet.Cells(3, 9).End(xlDown).Row + 1, 3).Select
                    OutSheet.Paste
                    OutSheet.Range(OutSheet.Cells(2 + j, 3), OutSheet.Cells(2 + j, 8)).Copy
                    OutSheet.Cells(3 + j, 3).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                        SkipBlanks:=False, Transpose:=False
                End If
                If OutSheet.Cells(3 + j, 4).Value = Constants.vbNullString Then
                    OutSheet.Cells(3 + j, 4).Value = TLItemList(ind).authorName
                    Exit For
                End If
                If OutSheet.Cells(3 + j, 4).Value = TLItemList(ind).authorName Then
                    Exit For
                End If
            Next j
        End If
    Next ind
    recolor
End Sub

Sub recolor()
    If OutSheet.Cells(4, 4).Value = Constants.vbNullString Then
        Exit Sub
    End If
    Dim man As Range
    Dim TeamListHeader As Range
    Dim CSDPTimeline As Range
    Dim MainEMChatTimeline As Range
    readCSDP
    readMainEMChats
    readAdditionalEMChats
    Set TeamListHeader = OutSheet.Cells.Find(What:="Team list", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    Set CSDPTimeline = OutSheet.Cells.Find(What:="CSDP timeline", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    Set MainEMChatTimeline = OutSheet.Cells.Find(What:="Main EM chat", After:=OutSheet.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    For Each man In Range(OutSheet.Cells(TeamListHeader.Row + 2, TeamListHeader.Column + 1), OutSheet.Cells(TeamListHeader.Row + 2, TeamListHeader.Column + 1).End(xlDown))
        For ind = CSDPTimeline.Row + 2 To 1000000
            If OutSheet.Cells(ind, MainEMChatTimeline.Column).Value = Constants.vbNullString Then
                ind = OutSheet.Cells(ind, MainEMChatTimeline.Column).End(xlDown).Row - 1
            ElseIf OutSheet.Cells(ind, MainEMChatTimeline.Column).Value Like "*" & man.Value & "*" Then
                With OutSheet.Cells(ind, MainEMChatTimeline.Column)
                    .Interior.Color = OutSheet.Cells(man.Row, 9).Interior.Color
                    .Font.Color = OutSheet.Cells(man.Row, 9).Font.Color
                End With
            End If
        Next ind
    Next man
    For i = MainEMChatTimeline.Column + 1 To 10000
        If OutSheet.Cells(MainEMChatTimeline.Row, i).Value = Constants.vbNullString Then
            i = OutSheet.Cells(MainEMChatTimeline.Row, i).End(xlToRight).Column - 1
        Else
            If UCase(OutSheet.Cells(MainEMChatTimeline.Row, i).Value) Like "*ADDITIONAL EM CHAT*" Then
                Set AdditionalEMChatTimeline = OutSheet.Cells(MainEMChatTimeline.Row, i)
                For Each man In Range(OutSheet.Cells(TeamListHeader.Row + 2, TeamListHeader.Column + 1), OutSheet.Cells(TeamListHeader.Row + 2, TeamListHeader.Column + 1).End(xlDown))
                    For ind = CSDPTimeline.Row + 2 To 1000000
                        If OutSheet.Cells(ind, AdditionalEMChatTimeline.Column).Value = Constants.vbNullString Then
                            ind = OutSheet.Cells(ind, AdditionalEMChatTimeline.Column).End(xlDown).Row - 1
                        ElseIf OutSheet.Cells(ind, AdditionalEMChatTimeline.Column).Value Like "*" & man.Value & "*" Then
                            With OutSheet.Cells(ind, AdditionalEMChatTimeline.Column)
                                .Interior.Color = OutSheet.Cells(man.Row, 9).Interior.Color
                                .Font.Color = OutSheet.Cells(man.Row, 9).Font.Color
                            End With
                        End If
                    Next ind
                Next man
            End If
        End If
    Next i
End Sub
