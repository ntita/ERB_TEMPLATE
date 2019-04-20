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
' геттеры
Public Property Get CSDP_HeaderAdress() As String
    CSDP_HeaderAdress = pCSDP_HeaderAdress
End Property
Public Property Get CSDP_HeaderHeigth() As Long
    CSDP_HeaderHeigth = pCSDP_HeaderHeigth
End Property
Public Property Get DateOfEvent_Address() As String
    DateOfEvent_Address = pDateOfEvent_Address
End Property
Public Property Get MaxString() As Long
    MaxString = pMaxString
End Property
Public Property Get MaxColumn() As Long
    MaxColumn = pMaxColumn
End Property
Public Property Get MainChats_FirstDataAdress() As String
    MainChats_FirstDataAdress = pMainChats_FirstDataAdress
End Property
Public Property Get MainChat_HeaderMask() As String
    MainChat_HeaderMask = pMainChat_HeaderMask
End Property
Private Sub Class_Initialize()
    pCSDP_HeaderAdress = "J2"
    pCSDP_HeaderHeigth = 3
    pDateOfEvent_Address = "D5"
    pMainChat_HeaderMask = "EM chat*"
    pMainChats_FirstDataAdress = "D10"
    pMaxString = 1048575
    pMaxColumn = 16382
End Sub
