VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pChatType As Long
Private pHeader As Range
Private pUTC As Long
Private pChatLen As Long
Private pDateOf As Date
Private pInSheet As Worksheet
Private pOutSheet As Worksheet
Public CurDate As Date
' геттеры
Public Property Get InSheet() As Worksheet
    If pInSheet Is Nothing Then
        Dim sh As Worksheet
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
        Dim sh As Worksheet
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
Public Property Get Name() As String
    Name = pName
End Property
Public Property Get ChatType() As Long
    ChatType = pChatType
End Property
Public Property Get Header() As Range
    If (pHeader Is Nothing) Then
        Dim HeaderCell As Range
        Set HeaderCell = InSheet.Cells.Find(What:=Name, After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
        If HeaderCell Is Nothing Then
            MsgBox Name & " not found. The program expects a column with a three-line header, the first of which contains '" & Name & "', second is Date, third - time zone.", vbCritical, "Error"
            End
        End If
        Set pHeader = HeaderCell
    End If
    Set Header = pHeader
End Property
Public Property Get UTC() As Long
    If Not pUTC < -12 And pUTC > 12 Then
        pUTC = Right(InSheet.Cells(Header.Row + 2, Header.Column).Value, 3)
    End If
    UTC = pUTC
End Property
Public Property Get ChatLen() As Long
    If pChatLen <= 0 Then
        pChatLen = InSheet.Cells(1048576, Header.Column).End(xlUp).Row - Header.Row - 2
    End If
    ChatLen = pChatLen
End Property
Public Property Get DateOf() As Date
    If (pDateOf = 0) Then
        If ChatType = -1 Then
            MsgBox "Chat type is not defined.", vbCritical, "Error"
            End
        ElseIf ChatType <> 1 Then
            pDateOf = InSheet.Cells(Header.Row + 1, Header.Column).Value
        End If
        If ChatType = 1 Or pDateOf = 0 Then
            Dim DateOfEvent As Range
            Set DateOfEvent = InSheet.Cells.Find(What:="Date of event", After:=InSheet.Cells(1, 1), LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                MatchCase:=False, SearchFormat:=False)
            If DateOfEvent Is Nothing Then
                MsgBox "Date of event not found. The program expects a table with a three-line header and one row. That row have to contain date.", vbCritical, "Error"
                End
            End If
            pDateOf = InSheet.Cells(DateOfEvent.Row + 3, DateOfEvent.Column).Value
        End If
    End If
    DateOf = pDateOf
End Property
' сеттеры
' "инициализатор", так как нормальный конcтруктор не предусмотрен в vba
Public Sub ChatLine_Init(newName As String, newChatType As Long)
    Dim status As String
    Application.StatusBar = "initializing new ChatLine: processing..."
    pName = newName
    pChatType = newChatType
    status = "initializing new ChatLine '" & Name & "'(" & ChatType & "): processing..."
    Application.StatusBar = status & " header is " & Header.Value
    Application.StatusBar = status & " UTC is " & UTC
    Application.StatusBar = status & " ChatLen is " & ChatLen
    Application.StatusBar = status & " DateOf is " & DateOf
    Application.StatusBar = ""
End Sub
Private Sub readChat(ByRef TLItemList As Collection)
    Application.StatusBar = "reading " & Name & ": processing..."
    Dim TLItem As TimelineItem
    Dim CurCell As Range
    Dim timeStamp As Date
    Dim timeSubStrLen As Date
    CurDate = DateOf
    For ind = Header.Row + 3 To ChatLen + Header.Row + 4
        Set CurCell = InSheet.Cells(ind, Header.Column)
        If CurCell.Value = Constants.vbNullString Then
            ind = CurCell.End(xlDown).Row - 1
        Else
            If (ChatType = 1) Then
                If (CurCell.Value Like "##:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                    Or CurCell.Value Like "#:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                    Or CurCell.Value Like "?[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*") _
                Then
                    If Not (TLItem Is Nothing) Then
                        timeStamp = TLItem.timeStamp
                    End If
                    Set TLItem = New TimelineItem
                    If CurCell.Value Like "##:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*" _
                    Then
                        timeSubStrLen = 5
                        TLItem.timeStamp = Left(CurCell.Value, timeSubStrLen)
                        TLItem.timeStamp = TLItem.timeStamp + CurDate
                        TLItem.timeStamp = DateAdd("h", -UTC, TLItem.timeStamp)
                    ElseIf (CurCell.Value Like "#:##[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*") _
                    Then
                        timeSubStrLen = 4
                        TLItem.timeStamp = Left(CurCell.Value, timeSubStrLen)
                        TLItem.timeStamp = TLItem.timeStamp + CurDate
                        TLItem.timeStamp = DateAdd("h", -UTC, TLItem.timeStamp)
                    ElseIf (CurCell.Value Like "?[A-Z][A-Z][A-Z][A-Z][A-Z][A-Z][A-Z]*") _
                    Then
                        timeSubStrLen = 1
                        TLItem.timeStamp = timeStamp
                    Else
                        decision = MsgBox("Line found that does not match any of the CSDP-patterns and will be skipped. Abort script (ignore warning, keep execute)?", vbYesNo + vbDefaultButton2, "Error")
                        If decision = vbYes Then
                            End
                        End If
                    End If
                    TLItem.authorName = Mid(CurCell.Value, timeSubStrLen + 1, 7)
                    TLItem.cellAddress = CurCell.Row & "," & CurCell.Column
                    TLItem.ChatType = ChatType
                    TLItem.mvalue = Mid(CurCell.Value, timeSubStrLen + 1 + 7)
                    If TLItem.timeStamp < timeStamp Then
                        TLItem.timeStamp = TLItem.timeStamp + 1
                        CurDate = CurDate + 1
                    End If
                    TLItemList.Add TLItem
                End If
            ElseIf (ChatType > 1) Then
                If (Trim(CurCell.Value) Like "*##:## ??:" _
                    Or Trim(CurCell.Value) Like "*#:## ??:" _
                    Or Trim(CurCell.Value) Like "*##:##:" _
                    Or Trim(CurCell.Value) Like "*#:##:") Then
                    If (Trim(CurCell.Value) Like "*##:## ??:") Then
                        timeLen = 8
                    ElseIf (Trim(CurCell.Value) Like "*#:## ??:") Then
                        timeLen = 7
                    ElseIf (Trim(CurCell.Value) Like "*##:##:") Then
                        timeLen = 5
                    ElseIf (Trim(CurCell.Value) Like "*#:##:") Then
                        timeLen = 4
                    End If
                    Set TLItem = New TimelineItem
                    TLItem.authorName = Trim(Left(Trim(CurCell.Value), Len(Trim(CurCell.Value)) - 1 - timeLen))
                    TLItem.cellAddress = CurCell.Row & "," & CurCell.Column
                    TLItem.ChatType = ChatType
                    TLItem.mvalue = ""
                    TLItem.timeStamp = Left(Right(Trim(CurCell.Value), 1 + timeLen), timeLen)
                    TLItem.timeStamp = TLItem.timeStamp + CurDate
                    TLItem.timeStamp = DateAdd("h", -UTC, TLItem.timeStamp)
                    If TLItem.timeStamp < timeStamp Then
                        TLItem.timeStamp = TLItem.timeStamp + 1
                        CurDate = CurDate + 1
                    End If
                    TLItemList.Add TLItem
                    timeStamp = TLItem.timeStamp
                End If
            End If
        End If
    Next ind
    Application.StatusBar = "reading " & Name & ": done"
End Sub

Private Sub Class_Initialize()
    pName = ""
    pChatType = -1
    Set pHeader = Nothing
    pUTC = 13
    pChatLen = -1
    pDateOf = 0
End Sub

