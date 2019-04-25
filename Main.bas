Attribute VB_Name = "Main"
Private pInSheet As Worksheet
Private pOutSheet As Worksheet
Private pTLItemList As Collection
Private pSettings As ERB_Settings
Public CurDate As Date
Public Property Get MainUTC() As Integer
    If pMainUTC = 0 Then
        Set EMChatHeader = InSheet.Range(Settings.CSDP_HeaderAdress)
        If EMChatHeader Is Nothing Then
            MsgBox "Timeline from CSDP not found. See initializator of ERB_Settings class, property pCSDP_HeaderAdress.", vbCritical, "Error"
            End
        End If
        pMainUTC = Conversion.CInt(Right(InSheet.Cells(EMChatHeader.Row + 2, EMChatHeader.Column).Value, 3))
    End If
    MainUTC = pMainUTC
End Property
Public Property Get Settings() As ERB_Settings
    If pSettings Is Nothing Then
        Set pSettings = New ERB_Settings
    End If
    Set Settings = pSettings
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
    Set CSDPHeader = InSheet.Range(Settings.CSDP_HeaderAdress)
    If CSDPHeader Is Nothing Then
        MsgBox "Timeline from CSDP not found. See initializator of ERB_Settings class, property pCSDP_HeaderAdress.", vbCritical, "Error"
        End
    End If
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    For ind = CSDPHeader.Row + Settings.CSDP_HeaderHeigth To Settings.MaxString
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
                TLItem.ChatType = 1
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
Private Sub readEMChat(chatName As String, ChatType As Integer, afterCell As Range)
    Dim EMChatHeader As Range
    Dim CurCell As Range
    Dim timeLen As String
    Dim TLItem As TimelineItem
    Set EMChatHeader = InSheet.Cells.Find(What:=chatName, After:=afterCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    If EMChatHeader Is Nothing Then
        MsgBox "Chat '" & chatName & "' not found.", vbCritical, "Error"
    End If
    If IsDate(InSheet.Cells(EMChatHeader.Row + 1, EMChatHeader.Column).Value) Then
        CurDate = InSheet.Cells(EMChatHeader.Row + 1, EMChatHeader.Column).Value
    End If
    EMUTC = -Right(InSheet.Cells(EMChatHeader.Row + 2, EMChatHeader.Column).Value, 3) + MainUTC
    For ind = EMChatHeader.Row + Settings.CSDP_HeaderHeigth To Settings.MaxString
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
                TLItem.ChatType = ChatType
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
    Dim MainEMChats As Range
    Dim CurMainChatCell As Range
    Dim EMChatHeader As Range
    Dim CurCell As Range
    Dim timeLen As String
    Dim TLItem As TimelineItem
    Set MainEMChats = InSheet.Range(Settings.MainChats_FirstDataAdress)
    If MainEMChats Is Nothing Then
        MsgBox "Table of main chats not found. See initializator of ERB_Settings class, property pMainChats_FirstDataAdress.", vbCritical, "Error"
        End
    End If
    For ind = MainEMChats.Row To Settings.MaxString
        Set CurMainChatCell = InSheet.Cells(ind, MainEMChats.Column)
        If CurMainChatCell.Value = Constants.vbNullString _
        Then
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
    Set CSDPHeader = InSheet.Range(Settings.CSDP_HeaderAdress)
    n = 2
    Set MainEMChatHeader = InSheet.Range(Settings.MainChats_FirstDataAdress)
    For jnd = CSDPHeader.Column + 1 To Settings.MaxColumn
        Set CurCell = InSheet.Cells(CSDPHeader.Row, jnd)
        If CurCell.Value Like Settings.MainChat_HeaderMask _
        Then
            isMain = False
            For ind = MainEMChatHeader.Row To Settings.MaxString
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
    InSheet.Columns(10).Cells(xlLastCell).Activate
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
        If TLItem.ChatType = 1 Then
            OutSheet.Cells(myCell.Row, CSDPReportedBy.Column) = TLItem.authorName
            OutSheet.Cells(myCell.Row, CSDPMessage.Column) = TLItem.mvalue
            myCell.Value = Null
        ElseIf TLItem.ChatType > 1 Then
            OutSheet.Cells(myCell.Row, CSDPTimeline.Column).Value = Null
            a = Split(TLItem.cellAddress, ",", 2)
            OutSheet.Cells(myCell.Row, EMChats(TLItem.ChatType, 0).Column).Value = InSheet.Cells(Conversion.CInt(a(0)), Conversion.CInt(a(1))).Value
            myCell.Value = Null
            OutSheet.Cells(myCell.Row, EMChats(TLItem.ChatType, 1).Column).Value = InSheet.Cells(Conversion.CInt(a(0)) + 1, Conversion.CInt(a(1))).Value
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
                    OutSheet.Cells(myCell.Row + i + j - 1, EMChats(TLItem.ChatType, 1).Column).EntireRow.Insert
                    ind = ind + 1
                    OutSheet.Cells(myCell.Row + i + j - 1, EMChats(TLItem.ChatType, 1).Column).Value = InSheet.Cells(Conversion.CInt(a(0)) + i, Conversion.CInt(a(1))).Value
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
        If TLItemList(ind).ChatType <> 1 Then
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

Sub createERB_Template_Timeline()
    'longInt, cell amount
    Dim lCA As Long
    'Address of cell on auxiliary sheet to copy
    Dim CellAdrAux As String
    CellAdrAux = "A2"
    'string, auxiliary sheet name
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheet_Prefix & "Timeline"
    'norm vars
    '****************************
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim CurDate As Date
    '****************************
    'end norm vars
    Application.ScreenUpdating = False
    'copy chat to new list
    With InSheet
        If IsEmpty(.[Settings.TimeLine_CellAdrTrg]) Then MsgBox "Cannot find CSDP timeline, check the settings", vbCritical, "Error"
        If Not IsSheetHereByName(AuxSheetName) Then
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = AuxSheetName
        Else
            Sheets(AuxSheetName).Cells.ClearContents
        End If
        Sheets(AuxSheetName).[A1:D1] = Array("Id", "DateOf", "AutorOf", "Note")
        lCA = .Cells(.Rows.Count, .Range(Settings.TimeLine_CellAdrTrg).Column).End(xlUp).Row - .Range(Settings.TimeLine_CellAdrTrg).Row + 1
        .Range(Settings.TimeLine_CellAdrTrg).Resize(lCA).Copy Sheets(AuxSheetName).Range(CellAdrAux)
    End With
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = Settings.TimeLine_RegExp
    End With
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    With Sheets(AuxSheetName)
        For Each cl In .Range(CellAdrAux).Resize(lCA)
            If Int(100 * cl.Row / lCA) <> prevPrc Then
                Application.StatusBar = "Выполнено: " & Format(Int(100 * (cl.Row - 1) / lCA), "##0") & "%" & String(CLng(20 * (cl.Row - 1) / lCA), ChrW(9632))
            End If
            prevPrc = Int(100 * (cl.Row - InSheet.Range(Settings.TimeLine_CellAdrTrg).Row + 1) / lCA)
            For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответствиям)
                If Not IsDate(M.SubMatches(0)) _
                Then
                    .Cells(cl.Row, 2) = .Cells(cl.Row - 1, 2).Value ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                Else
                    .Cells(cl.Row, 2) = CurDate + CDate(M.SubMatches(0)) ' время записи (первая группа соответствия)
                End If
                If .Cells(cl.Row, 2) < .Cells(cl.Row - 1, 2) Then
                    CurDate = CurDate + 1
                    .Cells(cl.Row, 2) = .Cells(cl.Row, 2) + 1
                End If
                .Cells(cl.Row, 3) = M.SubMatches(1) ' автор записи (вторая группа соответствия)
                .Cells(cl.Row, 4) = M.SubMatches(2) ' описание события (третья группа соответствия)
            Next M
            .Cells(cl.Row, 1) = cl.Row ' id записи (совпадает с номером строки)
        Next cl
        .Columns.AutoFit
    End With
    Application.StatusBar = "done"
    Application.ScreenUpdating = True
End Sub
Sub createERB_Template_Chat(CellAdrTrg As String)
    'longint, cells amount
    Dim lCA As Long
    'longint, counter
    Dim i As Long
    'Address of cell on auxiliary sheet to copy
    Dim CellAdrAux As String
    CellAdrAux = "A2"
    'string, prefix for technical lists
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheet_Prefix & "Chat_" & InSheet.Range(CellAdrTrg).Column
    Dim Feature As String
    Feature = IsThisChatMain(CellAdrTrg)
    'norm vars
    '****************************
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim CurDate As Date
    '****************************
    'end norm vars
    'copy chat to new list
    With InSheet
        If IsEmpty(.[CellAdrTrg]) Then MsgBox "Cannot find chat, check the settings", vbCritical, "Error"
        If Not IsSheetHereByName(AuxSheetName) Then
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = AuxSheetName
        Else
            Sheets(AuxSheetName).Cells.ClearContents
        End If
        Sheets(AuxSheetName).[A1:E1] = Array("Id", "DateOf", "AutorOf", "Note", "Feature")
        lCA = .Cells(.Rows.Count, .Range(CellAdrTrg).Column).End(xlUp).Row - .Range(CellAdrTrg).Row + 1
    End With
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = Settings.Chat_RegExp
    End With
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    With Sheets(AuxSheetName)
        i = .Range(CellAdrAux).Row - 1
        For Each cl In InSheet.Range(CellAdrTrg).Resize(lCA)
            If Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA) <> prevPrc Then
                Application.StatusBar = "Выполнено: " & Format(Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), "##0") & "%" & String(CLng(20 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), ChrW(9632)) '& String(20 - CLng(20 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), ChrW(9633))
            End If
            prevPrc = Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA)
            If re.Test(cl.Value) Then
                i = i + 1
                For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответствиям)
                    .Cells(i, 3) = M.SubMatches(0) ' автор записи (первая группа соответствия)
                    If Not IsDate(M.SubMatches(1)) _
                    Then
                        .Cells(i, 2) = .Cells(i - 1, 2) ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                    Else
                        .Cells(i, 2) = CurDate + CDate(M.SubMatches(1)) ' время записи (вторая группа соответствия)
                    End If
                    If .Cells(i, 2) < .Cells(i - 1, 2) Then
                        CurDate = CurDate + 1
                        .Cells(i, 2) = .Cells(i, 2) + 1
                    End If
                Next M
                .Cells(i, 5) = Feature
            Else
                If IsEmpty(.Cells(i, 4)) Then
                    .Cells(i, 4) = cl
                Else
                    .Cells(i, 4) = .Cells(i, 4) & vbLf & cl
                End If
            End If
            .Cells(i, 1) = i ' id записи (совпадает с номером строки)
        Next cl
        .Columns.AutoFit
    End With
End Sub
Sub createERB_Template_All_Chats()
    Application.StatusBar = "loading..."
    With InSheet
        For Each cl In .Range(.Range(Settings.Chat_CellAdrTrg), .Cells(.Range(Settings.Chat_CellAdrTrg).Row, .Columns.Count).End(xlToLeft))
            Application.ScreenUpdating = False
            createERB_Template_Chat (cl.Address)
            Application.ScreenUpdating = True
        Next cl
    End With
    Application.StatusBar = "done"
End Sub
Sub normalize()
    createERB_Template_Timeline
    createERB_Template_All_Chats
End Sub
Function IsSheetHereByName(aName As String) As Boolean
    IsSheetHereByName = False
    For Each sh In Sheets
        If sh.Name = aName Then
            IsSheetHereByName = True
            Exit Function
        End If
    Next
End Function
Function IsThisChatMain(aCellAdr As String) As String
    IsThisChatMain = "Additional"
    With InSheet
        For Each cl In .Range(Settings.MainChats_FirstDataAdress).CurrentRegion
            If cl.Value = .Cells(.Range(aCellAdr).Row - 3, .Range(aCellAdr).Column).Value Then
                IsThisChatMain = "Main"
                Exit Function
            End If
        Next cl
    End With
End Function
