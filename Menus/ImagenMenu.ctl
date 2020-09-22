VERSION 5.00
Begin VB.UserControl ImagenMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ImagenMenu.ctx":0000
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "ImagenMenu.ctx":0312
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "ImagenMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Public Property Let Form(NewVal As Variant)
Set SelForm = NewVal
End Property
Public Property Let ImageList(NewVal As Variant)
Set ImageL = NewVal
End Property
Public Sub Init()
Set CoolMenuObj = New CoolMenu
CoolMenuObj.Install SelForm.hWnd, ImageL, False, False
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 40 * Screen.TwipsPerPixelX
UserControl.Height = 40 * Screen.TwipsPerPixelY
End Sub
