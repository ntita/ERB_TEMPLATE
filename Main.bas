Attribute VB_Name = "Main"
Private pInSheet As Worksheet
Private pOutSheet As Worksheet
Private pTechSheet As Worksheet
Private pAuxSheet As Worksheet
Private pSettings As ERB_Settings
Private pEvents As Object
Private pMainUTC As Integer
Private pTeam As Object
Public CurDate As Date
Public AdditionalCount As Integer
Private pOutputBegin As Integer
Public Property Get MainUTC() As Integer
    If pMainUTC = 0 Then
        Set EMChatHeader = InSheet.Range(Settings.CSDP_HeaderAdress)
        If EMChatHeader Is Nothing Then
            MsgBox "Timeline from CSDP not found. See initializator of ERB_Settings class, property pCSDP_HeaderAdress.", vbCritical, "Error"
            End
        End If
        pMainUTC = Conversion.CInt(Right(InSheet.Range(Settings.CSDP_HeaderAdress).Offset(2).Value, 3))
    End If
    MainUTC = pMainUTC
End Property
Public Property Get Settings() As ERB_Settings
    If pSettings Is Nothing Then
        Set pSettings = New ERB_Settings
    End If
    Set Settings = pSettings
End Property
Public Property Get Events() As Object
    If pEvents Is Nothing Then
        Set pEvents = CreateObject("Scripting.Dictionary")
        With TechSheet.[K2].CurrentRegion
            For Row = 2 To .Rows.Count
                Key = CStr(.Cells(Row, 1).Value)
                Item = CStr(.Cells(Row, 2).Value)
                pEvents.Add Key, Item
            Next
        End With
    End If
    Set Events = pEvents
End Property
Public Property Get Team() As Object
    If pTeam Is Nothing Then
        Set pTeam = CreateObject("Scripting.Dictionary")
        With OutSheet.[C2].CurrentRegion
            For Row = 3 To .Rows.Count
                Key = CStr(.Cells(Row, 2).Value)
                Item = Array(CStr(.Cells(Row, 1).Value), Row, False)
                If Key <> "" Then
                    pTeam.Add Key, Item
                End If
            Next
        End With
    End If
    Set Team = pTeam
End Property
Public Property Get InSheet() As Worksheet
    Dim SheetName As String
    SheetName = Settings.InSheetName
    If pInSheet Is Nothing Then
        Set pInSheet = GetSheet(SheetName)
    End If
    Set InSheet = pInSheet
End Property
Public Property Get OutSheet() As Worksheet
    Dim SheetName As String
    SheetName = Settings.OutSheetName
    If pOutSheet Is Nothing Then
        Set pOutSheet = GetSheet(SheetName)
    End If
    Set OutSheet = pOutSheet
End Property
Public Property Get TechSheet() As Worksheet
    Dim SheetName As String
    SheetName = Settings.TechSheetName
    If pTechSheet Is Nothing Then
        Set pTechSheet = GetSheet(SheetName)
    End If
    Set TechSheet = pTechSheet
End Property
Public Property Get AuxSheet() As Worksheet
    Dim SheetName As String
    SheetName = Settings.AuxSheetName
    If pAuxSheet Is Nothing Then
        Set pAuxSheet = GetSheet(SheetName)
    End If
    Set AuxSheet = pAuxSheet
End Property
Public Property Get OutputBegin() As Integer
    If pOutputBegin = 0 Then
        With OutSheet
            For Each cl In .Range(.Cells(1, OutColumns.EV), .Cells(.Rows.Count, OutColumns.EV))
                If cl.Value = "EV" Then
                    pOutputBegin = cl.Row + 1
                    Exit For
                End If
            Next cl
        End With
    End If
    OutputBegin = pOutputBegin
End Property
Public Property Set OutputBegin(aOutputBegin)
    With OutSheet
        For Each cl In .Range(.Cells(1, OutColumns.EV), .Cells(.Rows.Count, OutColumns.EV))
            If cl.Value = "EV" Then
                pOutputBegin = cl.Row + 1
                Exit For
            End If
        Next cl
    End With
    OutputBegin = pOutputBegin
End Property
Sub createERB_Template_Timeline()
    'longInt, cell amount
    Dim lCA As Long
    'Address of cell on auxiliary sheet to copy
    Dim CellAdrAux As String
    CellAdrAux = "A2"
    'string, auxiliary sheet name
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheetName
    'norm vars
    '****************************
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    Dim CurDate As Date
    Dim UTC As Integer
    UTC = Conversion.CInt(Right(InSheet.Range(Settings.TimeLine_CellAdrTrg).Offset(-1).Value, 3))
    '****************************
    'end norm vars
    Application.ScreenUpdating = False
    'copy chat to new list
    With InSheet
        If IsEmpty(.[Settings.TimeLine_CellAdrTrg]) Then MsgBox "Cannot find CSDP timeline, check the settings", vbCritical, "Error"
        AuxSheet.Cells.ClearContents
        '  EV - 1, min from start - 2, Highlights - 3, UTC - 10;4, Reported by - 15;5, Message - 24;6
        ', Name/time stamp - 80;7, message - 92;8, Feature - 93;9, Name - 94;10, Id - 95;11, DateOf - 96;12
        AuxSheet.[A1:M1] = Array("Id", "ChatKind", "ChatName", "Employee", "EmployeeSurname", "UTC", "DateOf" _
          , "EV", "min from start", "Highlights", "Time", "Reported by", "Message")
        lCA = .Cells(.Rows.Count, .Range(Settings.TimeLine_CellAdrTrg).Column).End(xlUp).Row - .Range(Settings.TimeLine_CellAdrTrg).Row + 1
        .Range(Settings.TimeLine_CellAdrTrg).Resize(lCA).Copy Sheets(AuxSheetName).Range(CellAdrAux)
    End With
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = Settings.TimeLine_RegExp
    End With
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    Application.StatusBar = "CSDP: " & Format(0, "##0") & "%"
    With Sheets(AuxSheetName)
        For Each cl In .Range(CellAdrAux).Resize(lCA)
            ' вывод состо€ни€ дл€ нагл€дности, если состо€ние помен€лось
            If Int(100 * cl.Row / lCA) <> prevPrc Then
                Application.StatusBar = "CSDP: " & Format(Int(100 * (cl.Row - 1) / lCA), "##0") & "%" & String(CLng(20 * (cl.Row - 1) / lCA), ChrW(9632))
            End If
            ' запонимаем состо€ние
            prevPrc = Int(100 * (cl.Row - InSheet.Range(Settings.TimeLine_CellAdrTrg).Row + 1) / lCA)
            .Cells(cl.Row, AuxColumns.ChatKind) = "CSDP" ' Feature
            .Cells(cl.Row, AuxColumns.ChatName) = "CSDP" ' Name
            For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответстви€м)
                .Cells(cl.Row, AuxColumns.EV) = EventShortName(M.SubMatches(2))
                .Cells(cl.Row, AuxColumns.Highlights) = Events.Item(.Cells(cl.Row, AuxColumns.EV).Value)
                If Not IsDate(M.SubMatches(0)) _
                Then
                    .Cells(cl.Row, AuxColumns.DateOf) = .Cells(cl.Row - 1, AuxColumns.DateOf).Value ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                Else
                    .Cells(cl.Row, AuxColumns.DateOf) = CurDate + CDate(M.SubMatches(0)) ' врем€ записи (перва€ группа соответстви€)
                End If
                ' если в предыдущей строке в этом же столбце записана дата
                ' и она больше той, которую записали в текущей строке
                ' и в предыдущей строке этот же чат, то, по-видимому, перешли за границу дн€, надо увеличить день
                If IsDate(.Cells(cl.Row - 1, AuxColumns.DateOf)) _
                  And .Cells(cl.Row, AuxColumns.DateOf) < .Cells(cl.Row - 1, AuxColumns.DateOf) _
                  And .Cells(cl.Row, AuxColumns.ChatName) = .Cells(cl.Row - 1, AuxColumns.ChatName) _
                Then
                    CurDate = CurDate + 1
                    .Cells(cl.Row, AuxColumns.DateOf) = .Cells(cl.Row, AuxColumns.DateOf) + 1
                End If
                .Cells(cl.Row, AuxColumns.Time) = Format(.Cells(cl.Row, AuxColumns.DateOf), "hh:mm")
                .Cells(cl.Row, AuxColumns.Reported_by) = M.SubMatches(1) ' автор записи (втора€ группа соответстви€)
                .Cells(cl.Row, AuxColumns.Message) = M.SubMatches(2) ' описание событи€ (треть€ группа соответстви€)
            Next M
            .Cells(cl.Row, AuxColumns.ID) = cl.Row - 1 ' id записи
        Next cl
        .Columns.AutoFit
    End With
    Application.StatusBar = ""
    Application.ScreenUpdating = True
End Sub
Sub createERB_Template_Chat(CellAdrTrg As String)
    'longint, cells amount
    Dim lCA As Long
    'longint, counter
    Dim i As Long
    'string, prefix for technical lists
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheetName
    ' Ќомера колонок дл€ чата с учЄтом признака Additional
    Dim ccChatName_timestamp As Integer
    ccChatName_timestamp = AuxColumns.ChatName_timestamp
    Dim ccChatmessage As Integer
    ccChatmessage = AuxColumns.Chatmessage
    Dim Feature As String
    Feature = IsThisChatMain(CellAdrTrg)
    If Feature <> "Main" Then
        AdditionalCount = AdditionalCount + 1
        ccChatmessage = AuxColumns.Chatmessage + 2 * AdditionalCount
        ccChatName_timestamp = AuxColumns.ChatName_timestamp + 2 * AdditionalCount
    End If
    Dim Name As String
    Name = InSheet.Range(CellAdrTrg).Offset(-3).Value
    Dim UTC As Integer
    UTC = Conversion.CInt(Right(InSheet.Range(CellAdrTrg).Offset(-1).Value, 3))
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
            Sheets(Sheets.Count).Cells.Interior.PatternColorIndex = xlAutomatic
        End If
        lCA = .Cells(.Rows.Count, .Range(CellAdrTrg).Column).End(xlUp).Row - .Range(CellAdrTrg).Row + 1
    End With
    With Sheets(AuxSheetName)
        .[A1:F1] = Array("Id", "ChatKind", "ChatName", "Employee", "EmployeeSurname", "UTC", "DateOf")
        .Range(.Cells(1, ccChatName_timestamp), .Cells(1, ccChatmessage)) = Array("Name/time stamp #" & AdditionalCount, "message #" & AdditionalCount)
    End With
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = Settings.Chat_RegExp
    End With
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    Application.StatusBar = Name & ": " & Format(0, "##0") & "%"
    With Sheets(AuxSheetName)
        i = .[A1].CurrentRegion.Rows.Count
        For Each cl In InSheet.Range(CellAdrTrg).Resize(lCA)
            If Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA) <> prevPrc Then
                Application.StatusBar = Name & ": " & Format(Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), "##0") & "%" & String(CLng(20 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), ChrW(9632))
            End If
            prevPrc = Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA)
            If re.Test(cl.Value) Then
                i = i + 1
                .Cells(i, AuxColumns.ChatKind) = Feature
                .Cells(i, AuxColumns.ChatName) = Name
                .Cells(i, AuxColumns.UTC) = UTC
                For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответстви€м)
                    .Cells(i, ccChatName_timestamp) = M.SubMatches(0) ' автор записи и метка времени (перва€ группа соответстви€)
                    .Cells(i, AuxColumns.Employee) = M.SubMatches(1) ' автор записи (втора€ группа соответстви€)
                    .Cells(i, AuxColumns.EmployeeSurname) = M.SubMatches(2) ' фамили€ автора записи (треть€ группа соответстви€)
                    If Not IsDate(M.SubMatches(3)) _
                    Then
                        .Cells(i, AuxColumns.DateOf) = .Cells(i - 1, AuxColumns.DateOf) ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                    Else
                        .Cells(i, AuxColumns.DateOf) = CurDate + CDate(M.SubMatches(3)) ' врем€ записи (четвЄрта€ группа соответстви€)
                        If MainUTC <> UTC Then .Cells(i, AuxColumns.DateOf) = DateAdd("h", UTC - MainUTC, .Cells(i, AuxColumns.DateOf).Value)
                    End If
                    If IsDate(.Cells(i - 1, AuxColumns.DateOf)) _
                      And .Cells(i, AuxColumns.DateOf) < .Cells(i - 1, AuxColumns.DateOf) _
                      And .Cells(i - 1, AuxColumns.ChatName) = .Cells(i, AuxColumns.ChatName) Then
                        CurDate = CurDate + 1
                        .Cells(i, AuxColumns.DateOf) = .Cells(i, AuxColumns.DateOf) + 1
                    End If
                Next M
            Else
                If IsEmpty(.Cells(i, ccChatmessage)) Then
                    .Cells(i, ccChatmessage) = cl
                Else
                    .Cells(i, ccChatmessage) = .Cells(i, ccChatmessage) & vbLf & cl
                End If
            End If
            .Cells(i, AuxColumns.ID) = i - 1 ' id записи (совпадает с номером строки)
        Next cl
        .Columns.AutoFit
    End With
End Sub
Sub createERB_Template_All_Chats()
    'string, prefix for technical lists
    Application.StatusBar = "loading..."
    With InSheet
        AdditionalCount = 0
        For Each cl In .Range(.Range(Settings.Chat_CellAdrTrg), .Cells(.Range(Settings.Chat_CellAdrTrg).Row, .Columns.Count).End(xlToLeft))
            Application.ScreenUpdating = False
            createERB_Template_Chat (cl.Address)
            Application.ScreenUpdating = True
        Next cl
    End With
    Application.StatusBar = "done"
End Sub
Sub normalize()
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheetName
    With InSheet
        If Not IsSheetHereByName(AuxSheetName) Then
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = AuxSheetName
            Sheets(Sheets.Count).Cells.Interior.PatternColorIndex = xlAutomatic
        Else
            Sheets(AuxSheetName).Cells.ClearContents
        End If
    End With
    createERB_Template_Timeline
    createERB_Template_All_Chats
    Application.StatusBar = "Sorting columns on auxiliary sheet..."
    With Sheets(AuxSheetName).[A1].CurrentRegion
        .Sort key1:=.Cells(2, AuxColumns.DateOf), order1:=xlAscending, Header:=xlYes
    End With
    Application.StatusBar = ""
End Sub
Sub outputCSDPnChat()
    Application.StatusBar = "Delete old output... in progress"
    With OutSheet.Rows(OutputBegin & ":" & OutSheet.Rows.Count)
        .Clear
        .Interior.PatternColorIndex = xlAutomatic
    End With
    Application.StatusBar = ""
    ' скопировать столбцы со вспомогательного листа в соответствующие столбцы на output листе
    With AuxSheet
        ' подсчитываем количество строк в таблице на вспомогательном листе
        amount = .Range("A1").CurrentRegion.Rows.Count
        Application.StatusBar = "Copying data from aux to out sheet... in progress"
        ' копирование по столбцам
        For Each Item In Array(Array(AuxColumns.EV, OutColumns.EV, 0) _
            , Array(AuxColumns.min_from_start, OutColumns.min_from_start, 0) _
            , Array(AuxColumns.Highlights, OutColumns.Highlights, 0) _
            , Array(AuxColumns.Time, OutColumns.Time, 0) _
            , Array(AuxColumns.Reported_by, OutColumns.Reported_by, 0) _
            , Array(AuxColumns.Message, OutColumns.Message, 0) _
            , Array(AuxColumns.ChatName_timestamp, OutColumns.ChatName_timestamp, AdditionalCount) _
            , Array(AuxColumns.Chatmessage, OutColumns.Chatmessage, AdditionalCount))
            For i = 0 To Item(2)
                Application.StatusBar = "Copying " & amount & " rows from aux column " & Item(0) & " to out column " & Item(1) & "..."
                .Cells(2, Item(0) + 2 * i).Resize(amount).Copy OutSheet.Cells(OutputBegin, Item(1) + 45 * i)
            Next i
        Next Item
    End With
    Application.StatusBar = ""
End Sub
Sub RenderCSDPnChat()
    amount = Sheets(Settings.AuxSheetName).Range("A1").CurrentRegion.Rows.Count
    Application.StatusBar = "Merge cells: " & Format(0, "##0") & "%"
    With OutSheet
        For i = 1 To amount
            Application.StatusBar = "Merge cells: " & Format(Int(100 * i / amount), "##0") & "%" & String(CLng(20 * i / amount), ChrW(9632))
            For Each Item In Array(Array(4, 6), Array(11, 5), Array(16, 9), Array(25, 54))
                If .Cells(OutputBegin - 1 + i, Item(0)).Value <> "" Then
                    .Cells(OutputBegin - 1 + i, Item(0)).Resize(, Item(1)).Merge
                End If
            Next Item
            For j = 0 To AdditionalCount
                If .Cells(OutputBegin - 1 + i, 81 + 45 * j).Value <> "" Then
                    .Cells(OutputBegin - 1 + i, 81 + 45 * j).Resize(, 12).Merge
                End If
                If .Cells(OutputBegin - 1 + i, 93 + 45 * j).Value <> "" Then
                    .Cells(OutputBegin - 1 + i, 93 + 45 * j).Resize(, 33).Merge
                End If
            Next j
        Next i
    End With
    Application.StatusBar = ""
End Sub
Sub RenewTeamList()
    With AuxSheet.[D1].CurrentRegion
        For Row = 2 To .Rows.Count
            Key = CStr(.Cells(Row, 4).Value)
            If Key <> "" Then
                If Not Team.Exists(Key) Then
                    Item = Array("", -1, True)
                    Team.Add Key, Item
                Else
                    Team(Key) = Array(Team(Key)(0), Team(Key)(1), True)
                End If
            End If
        Next
    End With
End Sub
Sub RecolorChat()
    With OutSheet
        For i = 0 To 20
            amount = .Range(.Cells(OutputBegin, OutColumns.ChatName_timestamp + i * 45), .Cells(.Rows.Count, OutColumns.ChatName_timestamp + i * 45).End(xlUp)).Rows.Count
            For Each cl In .Range(.Cells(OutputBegin, OutColumns.ChatName_timestamp + i * 45), .Cells(.Rows.Count, OutColumns.ChatName_timestamp + i * 45).End(xlUp))
                If cl.Row - OutputBegin + 1 > 0 Then Application.StatusBar = "Recolor chat #" & i & ": " & Format(Int(100 * (cl.Row - OutputBegin + 1) / amount), "##0") & "%" & String(CLng(20 * (cl.Row - OutputBegin + 1) / amount), ChrW(9632))
                If cl.Value <> "" Then
                    For Each emp In Team.Keys()
                        If InStr(cl.Value, emp) <> 0 Then
                            cl.Interior.Color = OutSheet.Cells(Team(emp)(1) + 1, 9).Interior.Color
                            cl.Font.Color = OutSheet.Cells(Team(emp)(1) + 1, 9).Font.Color
                            Exit For
                        End If
                    Next emp
                End If
            Next cl
            Application.StatusBar = ""
        Next i
    End With
    Application.StatusBar = ""
End Sub
Sub reprintTimeline()
    With OutSheet
        chain_len = .[K5].End(xlDown).Row - .[L5].Row
        timeline_width = .Cells(4, .Columns.Count).End(xlToLeft).Column - .[L4].Column + .Cells(4, .Columns.Count).End(xlToLeft).MergeArea.Cells.Count
        With .[L5].Resize(chain_len, timeline_width)
            .Clear
            .Interior.PatternColorIndex = xlAutomatic
        End With
        cr = .[L5].Row
        cc = .[L5].Column
        exam_cr = cr
        exam_cc = cc
        amount = .Range(.Cells(OutputBegin, OutColumns.EV), .Cells(.Rows.Count, OutColumns.EV).End(xlUp)).Rows.Count
        For Each cl In .Range(.Cells(OutputBegin, OutColumns.EV), .Cells(.Rows.Count, OutColumns.EV).End(xlUp))
            Application.StatusBar = "Reprint timeline: " & Format(Int(100 * (cl.Row - OutputBegin + 1) / amount), "##0") & "%" & String(CLng(20 * (cl.Row - OutputBegin + 1) / amount), ChrW(9632))
            If cl.Value = "RP" Or cl.Value = "RI" Or cl.Value = "ESt" Then
                exam_cc = exam_cc + 26
                cr = exam_cr
                cc = exam_cc
            End If
            If cl.Value = "*" Then
                If cr - .[L5].Row >= chain_len Then
                    .Cells(cr - 1, cc).EntireRow.Insert
                    .Cells(cr, "L").Resize(, timeline_width).Cut OutSheet.Cells(cr - 1, "L")
                    .Cells(cr, "L").Resize(, timeline_width).Interior.PatternColorIndex = xlAutomatic
                    '.Cells(cr, "C").Resize(.Cells(cr, "D").End(xlDown).Row - cr + 1, 7).Cut OutSheet.Cells(cr - 1, "C")
                End If
                With .Cells(cr, cc)
                    .Value = Format(cl.Offset(, 9).Value, "h:mm")
                    .Font.Size = 8
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Font.Bold = True
                    .Resize(, 3).Merge
                End With
                With .Cells(cr, cc + 3)
                    .Value = cl.Offset(, 1).Value
                    .Font.Size = 8
                    If .Offset(-1, -1) <> "" Then
                        .Interior.Color = .Offset(-1, -1).Interior.Color
                        .Font.Color = .Offset(-1, -1).Font.Color
                    ElseIf .Offset(-1, 0) <> "" Then
                        .Interior.Color = .Offset(-1, 0).Interior.Color
                        .Font.Color = .Offset(-1, 0).Font.Color
                    Else
                        .Interior.Color = .Offset(-1, -3).Interior.Color
                        .Font.Color = .Offset(-1, -3).Font.Color
                    End If
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .Resize(, 3).Merge
                End With
                cl.Offset(, 2).Copy .Cells(cr, cc + 6)
                .Cells(cr, cc + 6).Font.Size = 8
                .Cells(cr, cc + 6).Resize(, 20).Merge
                cr = cr + 1
                cc = cc + 1
            End If
        Next cl
        Application.StatusBar = ""
    End With
End Sub
Sub Main()
    normalize
    RenewTeamList
    outputCSDPnChat
    RenderCSDPnChat
    RecolorChat
    reprintTimeline
    OutSheet.Activate
End Sub
Sub ClearAll()
    Application.StatusBar = "Delete old output... in progress"
    With OutSheet
        chain_len = .[K5].End(xlDown).Row - .[L5].Row
        timeline_width = .Cells(4, .Columns.Count).End(xlToLeft).Column - .[L4].Column + .Cells(4, .Columns.Count).End(xlToLeft).MergeArea.Cells.Count
        With .[L5].Resize(chain_len, timeline_width)
            .Clear
            .Interior.PatternColorIndex = xlAutomatic
        End With
    End With
    With OutSheet.Rows(OutputBegin & ":" & OutSheet.Rows.Count)
        .Clear
        .Interior.PatternColorIndex = xlAutomatic
    End With
    Application.StatusBar = "Delete old input... in progress"
    With InSheet.Range(Settings.TimeLine_CellAdrTrg).Resize(InSheet.Rows.Count - InSheet.Range(Settings.TimeLine_CellAdrTrg).Row, InSheet.Cells(InSheet.Range(Settings.CSDP_HeaderAdress).Row, InSheet.Columns.Count).End(xlToLeft).Column)
        .Clear
    End With
    With ActiveWorkbook
        For Each sh In .Worksheets
            If sh.Name = Settings.AuxSheetName Then
                Application.DisplayAlerts = False
                sh.Delete
                Application.DisplayAlerts = True
                Exit For
            End If
        Next sh
    End With
    Application.StatusBar = ""
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
            If cl.Value = .Range(aCellAdr).Offset(-3).Value Then
                IsThisChatMain = "Main"
                Exit Function
            End If
        Next cl
    End With
End Function
Function EventShortName(aEventName As String) As String
    EventShortName = ""
    For Each Key In Events.Keys()
        If InStr(aEventName, Events.Item(Key)) <> 0 Or InStr(Events.Item(Key), aEventName) <> 0 Then
            EventShortName = Key
            Exit Function
        End If
    Next Key
End Function
Function GetSheet(aSheetName As String) As Worksheet
    With ActiveWorkbook
        For Each sh In .Worksheets
            If sh.Name = aSheetName Then
                Set GetSheet = sh
                Exit Function
            End If
        Next sh
        If aSheetName = Settings.AuxSheetName Then
            .Sheets.Add After:=Sheets(Sheets.Count)
            With .Sheets(Sheets.Count)
                .Name = aSheetName
                .Cells.Interior.PatternColorIndex = xlAutomatic
            End With
        Else
            MsgBox "Sheet called '" & aSheetName & "' not found. Please, check the ERB_Settings", vbCritical, "Error"
        End If
    End With
End Function

