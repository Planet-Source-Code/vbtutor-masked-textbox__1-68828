VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   533
      TabIndex        =   1
      Text            =   "Regular TextBox"
      Top             =   1080
      Width           =   3615
   End
   Begin ResInput.RInput UserControl11 
      Height          =   495
      Left            =   533
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Mask            =   "+   -  -   -    "
      ValidChars      =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Text1_GotFocus()
  Text1.BackColor = &H80FFFF                              ' I prefer to use this approach to get user attention
End Sub

Private Sub Text1_LostFocus()
  Text1.BackColor = &H80000005                            ' Switch background color to white
End Sub

