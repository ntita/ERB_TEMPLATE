VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimelineItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pchatType As Integer
Private ptimeStamp As Date
Private pauthorName As String
Private pcellAddress As String
Private pmvalue As String

' геттеры
Public Property Get chatType() As Integer
    chatType = pchatType
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
Public Property Let chatType(ByVal newchatType As Integer)
    pchatType = newchatType
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
Public Sub TimelineItem_Initialize()
    chatType = 0
    timeStamp = Null
    authorName = ""
    cellAddress = ""
    pmvalue = ""
End Sub

'
