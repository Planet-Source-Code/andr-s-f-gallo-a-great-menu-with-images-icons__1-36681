VERSION 5.00
Object = "{6A50FD6B-E019-4299-A499-21B50F8563CD}#3.0#0"; "GraphicalMenu.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3045
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList I 
      Left            =   2640
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "SubMenu1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0112
            Key             =   "SubMenu2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0224
            Key             =   "Sub-SubMenu1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0336
            Key             =   "Sub-SubMenu2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0448
            Key             =   "SubMenu3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":055A
            Key             =   "SubMenu4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":066C
            Key             =   "SubMenu5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":077E
            Key             =   "SubMenu6"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0890
            Key             =   "SubMenu7"
         EndProperty
      EndProperty
   End
   Begin GraphicalMenu.ImagenMenu ImagenMenu1 
      Height          =   600
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1058
   End
   Begin VB.Menu M1 
      Caption         =   "Menu1"
      Begin VB.Menu S 
         Caption         =   "SubMenu1"
         Checked         =   -1  'True
      End
      Begin VB.Menu S2 
         Caption         =   "SubMenu2"
         Begin VB.Menu ss 
            Caption         =   "Sub-SubMenu1"
            Index           =   1
            Shortcut        =   ^E
         End
         Begin VB.Menu ss 
            Caption         =   "Sub-SubMenu2"
            Index           =   2
         End
      End
      Begin VB.Menu S3 
         Caption         =   "SubMenu3"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu S4 
         Caption         =   "SubMenu4"
      End
      Begin VB.Menu S5 
         Caption         =   "SubMenu5"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu M2 
      Caption         =   "Menu2"
      Begin VB.Menu Sm 
         Caption         =   "SubMenu6"
      End
      Begin VB.Menu Sm2 
         Caption         =   "SubMenu7"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'Sets the Imagelist to the menu control
ImagenMenu1.ImageList = I
'Sets in which form it will work
ImagenMenu1.Form = Me
'Start
ImagenMenu1.Init
End Sub
