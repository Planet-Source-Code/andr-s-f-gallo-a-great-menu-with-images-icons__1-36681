VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMyItemDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mCol As Collection
Public Function Add(Optional sMenuText As String, Optional fType As Long, Optional iButton As Integer, Optional sKey As String) As CMyItemData
Dim objNewMember As CMyItemData
Set objNewMember = New CMyItemData
objNewMember.sMenuText = sMenuText
objNewMember.fType = fType
objNewMember.iButton = iButton
If Len(sKey) = 0 Then
  mCol.Add objNewMember
Else
  mCol.Add objNewMember, sKey
End If
Set Add = objNewMember
Set objNewMember = Nothing
End Function
Public Property Get Item(vntIndexKey As Variant) As CMyItemData
On Error GoTo h:
Set Item = mCol(vntIndexKey)
h:
End Property
Public Property Get Count() As Long
Count = mCol.Count
End Property
Public Sub Remove(vntIndexKey As Variant)
mCol.Remove vntIndexKey
End Sub
Private Sub Class_Initialize()
Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
Set mCol = Nothing
End Sub

