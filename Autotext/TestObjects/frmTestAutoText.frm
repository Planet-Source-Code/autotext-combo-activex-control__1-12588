VERSION 5.00
Object = "{49A45244-C3B7-40BF-A91E-FF1306814D71}#1.0#0"; "autotext.ocx"
Begin VB.Form frmTestAutoText 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1035
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1305
      Width           =   1635
   End
   Begin vbAutoText.AutoText cbo 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Top             =   720
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      BeginProperty Font0 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowContextMenu=   0   'False
      CustomMenu      =   "Hello"
   End
End
Attribute VB_Name = "frmTestAutoText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   Dim ebo As IAutoText


Private Sub Form_Load()
   Dim i As Long
   Dim s As Single
   
   Set ebo = cbo.Object
   's = Timer
   For i = 0 To 10000
      With ebo
         .AddItem CStr(i) & "Chuck U"
         .ItemData(.NewIndex) = i
      End With
   Next i
   'Debug.Print Format$(Timer - s, "####.##")
End Sub
