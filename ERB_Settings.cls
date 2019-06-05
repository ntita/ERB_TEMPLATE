VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ERB_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pCSDP_HeaderAdress As String
Private pCSDP_HeaderHeigth As Long
Private pMainChats_FirstDataAdress As String
Private pMainChat_HeaderMask As String
Private pDateOfEvent_Address As String
Private pMaxString As Long
Private pMaxColumn As Long
Private pTimeLine_RegExp As String
Private pTimeLine_CellAdrTrg As String
Private pChat_RegExp As String
Private pChat_CellAdrTrg As String
Private pAuxSheetName As String
Private pInSheetName As String
Private pOutSheetName As String
Private pTechSheetName As String
Enum AuxColumns
    ID = 1
    ChatKind
    ChatName
    Employee
    EmployeeSurname
    UTC
    DateOf
    EV
    min_from_start
    Highlights
    Time
    Reported_by
    Message
    ChatName_timestamp
    Chatmessage
End Enum
Enum OutColumns
    EV = 2
    min_from_start = 3
    Highlights = 4
    Time = 11
    Reported_by = 16
    Message = 25
    ChatName_timestamp = 81
    Chatmessage = 93
End Enum
' ???????
Public Property Get CSDP_HeaderAdress() As String
    CSDP_HeaderAdress = pCSDP_HeaderAdress
End Property
Public Property Get CSDP_HeaderHeigth() As Long
    CSDP_HeaderHeigth = pCSDP_HeaderHeigth
End Property
Public Property Get DateOfEvent_Address() As String
    DateOfEvent_Address = pDateOfEvent_Address
End Property
Public Property Get MainChats_FirstDataAdress() As String
    MainChats_FirstDataAdress = pMainChats_FirstDataAdress
End Property
Public Property Get MainChat_HeaderMask() As String
    MainChat_HeaderMask = pMainChat_HeaderMask
End Property
Public Property Get TimeLine_RegExp() As String
    TimeLine_RegExp = pTimeLine_RegExp
End Property
Public Property Get TimeLine_CellAdrTrg() As String
    TimeLine_CellAdrTrg = pTimeLine_CellAdrTrg
End Property
Public Property Get AuxSheetName() As String
    AuxSheetName = pAuxSheetName
End Property
Public Property Get Chat_RegExp() As String
    Chat_RegExp = pChat_RegExp
End Property
Public Property Get Chat_CellAdrTrg() As String
    Chat_CellAdrTrg = pChat_CellAdrTrg
End Property
Public Property Get TechSheetName() As String
    TechSheetName = pTechSheetName
End Property
Public Property Get InSheetName() As String
    InSheetName = pInSheetName
End Property
Public Property Get OutSheetName() As String
    OutSheetName = pOutSheetName
End Property
Private Sub Class_Initialize()
    pCSDP_HeaderAdress = "J2"
    pCSDP_HeaderHeigth = 3
    pDateOfEvent_Address = "D5"
    pMainChats_FirstDataAdress = "D10"
    ' sheets names
    pInSheetName = "Input list"
    pOutSheetName = "Prepared timeline output"
    pTechSheetName = "backgound_tech_list"
    pAuxSheetName = "ERB_Template_Chat"
    ' CSDP Timeline normilize
    ' RegExp: group #1 - timestamp, #2 - author, #3 - event full name
    ' before 190606 pTimeLine_RegExp = "(\d?\d:\d\d)? ?(\w{7})(.*)"
    pTimeLine_RegExp = "(\d?\d:\d\d)? ()(.*)"
    pTimeLine_CellAdrTrg = "J5"
    ' chat normilize
    ' RegExp: group #1 - raw header, #2 - author, #3 - author's surname, #4 - timestamp
    ' before 190606 pChat_RegExp = "^((\w* (.*)) (\d?\d:\d\d):)\s*$"
    pChat_RegExp = "^((\w* (.*)) (\d?\d:\d\d[^:]*):)\s*$"
    pChat_CellAdrTrg = "L5"
    ' end normilize
End Sub
