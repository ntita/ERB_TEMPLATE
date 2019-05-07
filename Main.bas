Attribute VB_Name = "Main"
Private pInSheet As Worksheet
Private pOutSheet As Worksheet
Private pTechSheet As Worksheet
Private pTLItemList As Collection
Private pSettings As ERB_Settings
Private pEvents As Object
Private pMainUTC As Integer
'AuxSheet column const
Const cID As Integer = 1
Const cFeature As Integer = 2
Const cName As Integer = 3
Const cEmployee As Integer = 4
Const cUTC As Integer = 5
Const cDateOf As Integer = 6
Const cEV As Integer = 7
Const cmin_from_start As Integer = 8
Const cHighlights As Integer = 9
Const cTime As Integer = 10
Const cReported_by As Integer = 11
Const cmessage As Integer = 12
Const cChatName_timestamp As Integer = 13
Const cChatmessage As Integer = 14
'end AuxSheet column const
Public CurDate As Date
Public AdditionalCount As Integer
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
Public Property Get Events() As Object
    If pEvents Is Nothing Then
        Set pEvents = CreateObject("Scripting.Dictionary")
        With TechSheet.Range("Events")
            For Row = 2 To .Rows.Count
                Key = CStr(.Cells(Row, 1).Value)
                Item = CStr(.Cells(Row, 2).Value)
                pEvents.Add Key, Item
            Next
        End With
    End If
    Set Events = pEvents
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
Sub createERB_Template_Timeline()
    'longInt, cell amount
    Dim lCA As Long
    'Address of cell on auxiliary sheet to copy
    Dim CellAdrAux As String
    CellAdrAux = "A2"
    'string, auxiliary sheet name
    Dim AuxSheetName As String
    AuxSheetName = Settings.AuxSheet_Prefix
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
        '  EV - 1, min from start - 2, Highlights - 3, UTC - 10;4, Reported by - 15;5, Message - 24;6
        ', Name/time stamp - 80;7, message - 92;8, Feature - 93;9, Name - 94;10, Id - 95;11, DateOf - 96;12
        Sheets(AuxSheetName).[A1:L1] = Array("Id", "Feature", "Name", "Employee", "UTC", "DateOf" _
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
    Application.StatusBar = "CSDP: " & Format(0, "000") & "%"
    With Sheets(AuxSheetName)
        For Each cl In .Range(CellAdrAux).Resize(lCA)
            ' вывод состояния для наглядности, если состояние поменялось
            If Int(100 * cl.Row / lCA) <> prevPrc Then
                Application.StatusBar = "CSDP: " & Format(Int(100 * (cl.Row - 1) / lCA), "000") & "%" & String(CLng(20 * (cl.Row - 1) / lCA), ChrW(9632))
            End If
            ' запонимаем состояние
            prevPrc = Int(100 * (cl.Row - InSheet.Range(Settings.TimeLine_CellAdrTrg).Row + 1) / lCA)
            .Cells(cl.Row, cFeature) = "CSDP" ' Feature
            .Cells(cl.Row, cName) = "CSDP" ' Name
            For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответствиям)
                .Cells(cl.Row, cEV) = EventShortName(M.SubMatches(2))
                .Cells(cl.Row, cHighlights) = Events.Item(.Cells(cl.Row, cEV).Value)
                If Not IsDate(M.SubMatches(0)) _
                Then
                    .Cells(cl.Row, cDateOf) = .Cells(cl.Row - 1, cDateOf).Value ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                Else
                    .Cells(cl.Row, cDateOf) = CurDate + CDate(M.SubMatches(0)) ' время записи (первая группа соответствия)
                End If
                ' если в предыдущей строке в этом же столбце записана дата
                ' и она больше той, которую записали в текущей строке
                ' и в предыдущей строке этот же чат, то, по-видимому, перешли за границу дня, надо увеличить день
                If IsDate(.Cells(cl.Row - 1, cDateOf)) _
                  And .Cells(cl.Row, cDateOf) < .Cells(cl.Row - 1, cDateOf) _
                  And .Cells(cl.Row, cName) = .Cells(cl.Row - 1, cName) _
                Then
                    CurDate = CurDate + 1
                    .Cells(cl.Row, cDateOf) = .Cells(cl.Row, cDateOf) + 1
                End If
                .Cells(cl.Row, cTime) = Format(.Cells(cl.Row, cDateOf), "hh:mm")
                .Cells(cl.Row, cReported_by) = M.SubMatches(1) ' автор записи (вторая группа соответствия)
                .Cells(cl.Row, cmessage) = M.SubMatches(2) ' описание события (третья группа соответствия)
            Next M
            .Cells(cl.Row, cID) = cl.Row - 1 ' id записи
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
    AuxSheetName = Settings.AuxSheet_Prefix
    ' Номера колонок для чата с учётом признака Additional
    Dim ccChatName_timestamp As Integer
    ccChatName_timestamp = cChatName_timestamp
    Dim ccChatmessage As Integer
    ccChatmessage = cChatmessage
    Dim Feature As String
    Feature = IsThisChatMain(CellAdrTrg)
    If Feature <> "Main" Then
        AdditionalCount = AdditionalCount + 1
        ccChatmessage = cChatmessage + 2 * AdditionalCount
        ccChatName_timestamp = cChatName_timestamp + 2 * AdditionalCount
    End If
    Dim Name As String
    Name = InSheet.Cells(InSheet.Range(CellAdrTrg).Row - 3, InSheet.Range(CellAdrTrg).Column).Value
    Dim UTC As Integer
    UTC = Conversion.CInt(Right(InSheet.Cells(InSheet.Range(CellAdrTrg).Row - 1, InSheet.Range(CellAdrTrg).Column).Value, 3))
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
        End If
        lCA = .Cells(.Rows.Count, .Range(CellAdrTrg).Column).End(xlUp).Row - .Range(CellAdrTrg).Row + 1
    End With
    With Sheets(AuxSheetName)
        .[A1:F1] = Array("Id", "Feature", "Name", "Employee", "UTC", "DateOf")
        .Range(.Cells(1, ccChatName_timestamp), .Cells(1, ccChatmessage)) = Array("Name/time stamp #" & AdditionalCount, "message #" & AdditionalCount)
    End With
    With re
        .Global = False
        .IgnoreCase = True
        .Pattern = Settings.Chat_RegExp
    End With
    CurDate = InSheet.Range(Settings.DateOfEvent_Address).Value
    Application.StatusBar = Name & ": " & Format(0, "000") & "%"
    With Sheets(AuxSheetName)
        i = .[A1].CurrentRegion.Rows.Count
        For Each cl In InSheet.Range(CellAdrTrg).Resize(lCA)
            If Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA) <> prevPrc Then
                Application.StatusBar = Name & ": " & Format(Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), "000") & "%" & String(CLng(20 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA), ChrW(9632))
            End If
            prevPrc = Int(100 * (cl.Row - .Range(CellAdrTrg).Row + 1) / lCA)
            If re.Test(cl.Value) Then
                i = i + 1
                .Cells(i, cFeature) = Feature
                .Cells(i, cName) = Name
                .Cells(i, cUTC) = UTC
                For Each M In re.Execute(cl.Value) ' цикл по MatchColection (по всем найденным соответствиям)
                    .Cells(i, ccChatName_timestamp) = M.SubMatches(0) ' автор записи и метка времени (первая группа соответствия)
                    .Cells(i, cEmployee) = M.SubMatches(1) ' автор записи (первая группа соответствия)
                    If Not IsDate(M.SubMatches(2)) _
                    Then
                        .Cells(i, cDateOf) = .Cells(i - 1, cDateOf) ' если нет временной метки, то считаем, что она совпадает с предыдущей строкой
                    Else
                        .Cells(i, cDateOf) = CurDate + CDate(M.SubMatches(2)) - MainUTC + UTC ' время записи (вторая группа соответствия)
                    End If
                    If IsDate(.Cells(i - 1, cDateOf)) _
                      And .Cells(i, cDateOf) < .Cells(i - 1, cDateOf) _
                      And .Cells(i - 1, cName) = .Cells(i, cName) Then
                        CurDate = CurDate + 1
                        .Cells(i, cDateOf) = .Cells(i, cDateOf) + 1
                    End If
                Next M
            Else
                If IsEmpty(.Cells(i, ccChatmessage)) Then
                    .Cells(i, ccChatmessage) = cl
                Else
                    .Cells(i, ccChatmessage) = .Cells(i, ccChatmessage) & vbLf & cl
                End If
            End If
            .Cells(i, cID) = i - 1 ' id записи (совпадает с номером строки)
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
    AuxSheetName = Settings.AuxSheet_Prefix
    With InSheet
        If Not IsSheetHereByName(AuxSheetName) Then
            Sheets.Add After:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = AuxSheetName
        Else
            Sheets(AuxSheetName).Cells.ClearContents
        End If
    End With
    createERB_Template_Timeline
    createERB_Template_All_Chats
    Application.StatusBar = "Sorting columns(""A:N"")..."
    Sheets(AuxSheetName).Columns("A:N").Sort key1:=Sheets(AuxSheetName).Range("F2"), _
      order1:=xlAscending, Header:=xlYes
    Application.StatusBar = ""
End Sub
Sub outputCSDPnChat()
    Application.ScreenUpdating = False
    Application.StatusBar = "Delete old output... in progress"
    OutSheet.Rows("25:" & OutSheet.Rows.Count).Delete shift:=xlUp
    Application.ScreenUpdating = True
    ' скопировать столбцы со вспомогательного листа в соответствующие столбцы на output листе
    With Sheets(Settings.AuxSheet_Prefix)
        ' подсчитываем количество строк в таблице на вспомогательном листе
        amount = .Range("A1").CurrentRegion.Rows.Count
        Application.StatusBar = ""
        ' первые 12 столбцов до чатов
        For Each Item In Array(Array("G2", "B25"), Array("H2", "C25"), Array("I2", "D25") _
          , Array("J2", "K25"), Array("K2", "P25"), Array("L2", "Y25"))
            Application.StatusBar = "Copying " & amount & " rows from aux column " & Item(0) & " to out column " & Item(1) & "..."
            .Range(Item(0)).Resize(amount).Copy OutSheet.Range(Item(1))
        Next Item
        For i = 0 To AdditionalCount
            Application.StatusBar = "Copying " & amount & " rows from aux column " & .Cells(2, cChatName_timestamp + 2 * i).Address & " to out column " & OutSheet.Cells(25, 81 + 45 * i).Address & "..."
            .Cells(2, cChatName_timestamp + 2 * i).Resize(amount).Copy OutSheet.Cells(25, 81 + 45 * i)
            Application.StatusBar = "Copying " & amount & " rows from aux column " & .Cells(2, cChatmessage + 2 * i).Address & " to out column " & OutSheet.Cells(25, 93 + 45 * i).Address & "..."
            .Cells(2, cChatmessage + 2 * i).Resize(amount).Copy OutSheet.Cells(25, 93 + 45 * i)
        Next i
    End With
    Application.StatusBar = ""
End Sub
Sub RenderCSDPnChat()
    amount = Sheets(Settings.AuxSheet_Prefix).Range("A1").CurrentRegion.Rows.Count
    Application.StatusBar = "Merge cells: " & Format(0, "##0") & "%"
    With OutSheet
        For i = 1 To amount
            Application.StatusBar = "Merge cells: " & Format(Int(100 * i / amount), "000") & "%" & String(CLng(20 * i / amount), ChrW(9632))
            For Each Item In Array(Array(4, 6), Array(11, 5), Array(16, 9), Array(25, 54))
                If .Cells(24 + i, Item(0)).Value <> "" Then
                    .Cells(24 + i, Item(0)).Resize(, Item(1)).Merge
                End If
            Next Item
            For j = 0 To AdditionalCount
                If .Cells(24 + i, 81 + 45 * j).Value <> "" Then
                    .Cells(24 + i, 81 + 45 * j).Resize(, 12).Merge
                End If
                If .Cells(24 + i, 93 + 45 * j).Value <> "" Then
                    .Cells(24 + i, 93 + 45 * j).Resize(, 33).Merge
                End If
            Next j
        Next i
    End With
    Application.StatusBar = ""
End Sub
Sub RenewTeamList()
    With OutSheet
        amount = .Range("C2").CurrentRegion.Rows.Count - 2
        For Each cl In .Range("D4").Resize(amount)
            If cl.Value = "" Then
                'continue
            End If
        Next cl
    End With
End Sub
Sub Main()
    normalize
    outputCSDPnChat
    RenderCSDPnChat
    RenewTeamList
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
Function EventShortName(aEventName As String) As String
    EventShortName = ""
    For Each Key In Events.Keys()
        If InStr(aEventName, Events.Item(Key)) <> 0 Then
            EventShortName = Key
            Exit Function
        End If
    Next Key
End Function
Function GetSheet(aSheetName As String) As Worksheet
    For Each sh In ActiveWorkbook.Worksheets
        If sh.Name = aSheetName Then
            Set GetSheet = sh
            Exit Function
        End If
    Next sh
    MsgBox "Sheet called '" & aSheetName & "' not found. Please, check the ERB_Settings", vbCritical, "Error"
End Function
