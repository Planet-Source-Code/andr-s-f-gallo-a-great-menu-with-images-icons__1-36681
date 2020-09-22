VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMyItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mvarsMenuText As String
Private mvarfType As Long
Private mvariButton As Integer
Private mvarbMainMenu As Boolean
Public Property Let bMainMenu(ByVal vData As Boolean)
mvarbMainMenu = vData
End Property
Public Property Get bMainMenu() As Boolean
bMainMenu = mvarbMainMenu
End Property
Public Property Let iButton(ByVal vData As Integer)
mvariButton = vData
End Property
Public Property Get iButton() As Integer
iButton = mvariButton
End Property
Public Property Let fType(ByVal vData As Long)
mvarfType = vData
End Property
Public Property Get fType() As Long
fType = mvarfType
End Property
Public Property Let sMenuText(ByVal vData As String)
mvarsMenuText = vData
End Property
Public Property Get sMenuText() As String
sMenuText = mvarsMenuText
End Property

