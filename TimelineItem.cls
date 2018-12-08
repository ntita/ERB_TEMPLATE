VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimelineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pChatType As Integer
Private ptimeStamp As Date
Private pauthorName As String
Private pcellAddress As String
Private pmvalue As String

' геттеры
Public Property Get ChatType() As Integer
    ChatType = pChatType
End Property
Public Property Get timeStamp() As Date
    timeStamp = ptimeStamp
End Property
Public Property Get authorName() As String
    authorName = pauthorName
End Property
Public Property Get cellAddress() As String
    cellAddress = pcellAddress
End Property
Public Property Get mvalue() As String
    mvalue = pmvalue
End Property

' сеттеры
Public Property Let ChatType(ByVal newChatType As Integer)
    pChatType = newChatType
End Property
Public Property Let timeStamp(ByVal newtimeStamp As Date)
    ptimeStamp = newtimeStamp
End Property
Public Property Let authorName(ByVal newauthorName As String)
    pauthorName = newauthorName
End Property
Public Property Let cellAddress(ByVal newcellAddress As String)
    pcellAddress = newcellAddress
End Property
Public Property Let mvalue(ByVal newmvalue As String)
    pmvalue = newmvalue
End Property

' конструктор
Private Sub Class_Initialize()
    pChatType = 0
    ptimeStamp = DateTime.Now
    pauthorName = ""
    pcellAddress = ""
    pmvalue = ""
End Sub
