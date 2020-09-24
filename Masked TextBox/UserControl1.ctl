VERSION 5.00
Begin VB.UserControl RInput 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ScaleHeight     =   675
   ScaleWidth      =   3675
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "RInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Const def_LFStyle = 1               ' Default LostFocusStyle = ShowRedColor

Public Enum LostFocusStyles
    None
    ShowRedColor                    ' If the user dosn't fill the textbox, the text will turn red
    MustFill                        ' The user must fill the textbox
End Enum

Private OldColor As Long
Private LFStyle As LostFocusStyles ' Current LostFocusStyle
Private sFormat As String          ' A format of your choice, Spaces in this format indicates where the ValidChars goes
Private sValidChars As String      ' The Chars that can be entered into the format
Private TP As Integer              ' Text position in the textbox

Private Sub UserControl_Initialize()
  OldColor = Text1.BackColor
  Text1 = "" ' Clear the textbox
  LFStyle = def_LFStyle ' Set Default LostFocusStyle = None
End Sub

Private Sub UserControl_Resize()
  ' Set the size so that the textbox is as big as the UserControl
  Text1.Height = UserControl.Height
  Text1.Width = UserControl.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ' Read the Properties from the PropBag
  sFormat = PropBag.ReadProperty("Mask")
  sValidChars = PropBag.ReadProperty("ValidChars")
  LFStyle = PropBag.ReadProperty("LFStyle", def_LFStyle)
  Text1 = sFormat ' Show the new format in the textbox
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  ' Write the Properties to the PropBag
  Call PropBag.WriteProperty("Mask", sFormat)
  Call PropBag.WriteProperty("ValidChars", sValidChars)
  Call PropBag.WriteProperty("LFStyle", LFStyle, def_LFStyle)
End Sub

Public Property Get Text() As String
  Text = Text1 ' Return the text from the textbox
End Property

Public Property Get Mask() As String
  Mask = sFormat ' Return the current Mask format
End Property

Public Property Let Mask(s_Format As String)
  sFormat = s_Format ' Set new Mask format
  Text1 = s_Format   ' Show it in the textbox
  ' Notify the container that the Mask's value has been changed.
  PropertyChanged "Mask"
End Property

Public Property Get ValidChars() As String
  ValidChars = sValidChars ' Return the current Valid Chars
End Property

Public Property Let ValidChars(s_ValidChars As String)
  sValidChars = s_ValidChars ' Set the new Valid Chars
  PropertyChanged "ValidChars"
End Property

Public Property Get LostFocusStyle() As LostFocusStyles
    LostFocusStyle = LFStyle
End Property

Public Property Let LostFocusStyle(ByVal New_LFStyle As LostFocusStyles)
    LFStyle = New_LFStyle
    PropertyChanged "LFStyle"
End Property

Private Sub Text1_GotFocus()
  Text1.BackColor = &H80FFFF                              ' I prefer to use this approach to get user attention
End Sub

Private Sub Text1_LostFocus()
  If Len(sFormat) <> TP Then       ' The textbox haven't been filled so
    Select Case LFStyle
      Case ShowRedColor
        Text1.BackColor = vbRed     ' Set the new text color to red
      Case MustFill
        Text1.SetFocus             ' Keep the user focused to this textbox
    End Select
  Else
    Text1.BackColor = OldColor
  End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    ' Delete all chars in textbox if and only if all is selected
    If Text1.SelLength = Len(sFormat) Then
      Text1 = sFormat ' Set the textbox to show the "empty" format
      TP = 0 ' Reset the Text position
      Text1.SelStart = 0 ' Clear selection
    End If
    KeyCode = 0 ' Ignore delete key
  End If
  If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim S As String ' String used to temporally hold the text in text1
  Dim iSepLenS As Integer
  Dim iSepLenTP As Integer
  Dim iSepLenCP As Integer
  Dim iSpaceInStr As Integer
  Dim CP As Integer ' Cursor position in the textbox
  Dim N As Integer  ' Counter
  Dim Key As Integer ' The Key Pressed
 
 Key = KeyAscii ' Save pressed key
 KeyAscii = 0 ' Ignore key, I'll handle what does and dosn't go in the Textbox
 S = Text1
 CP = Text1.SelStart ' Get the cursor position
 If (Key <> vbKeyBack) Then
   If Key = 13 Then Exit Sub
   If Key < 48 Or Key > 57 Then
     Beep
     Exit Sub ' Ignore char
   End If
   TP = TP + 1 'Assume the text is getting a char longer
   
   ' Check for separators at the text position or at the start in text1
   If Trim$(Mid$(sFormat, TP, 1)) <> "" Then
     iSpaceInStr = InStr(TP + 1, sFormat, " ")
     If iSpaceInStr > 0 Then
       iSepLenS = iSpaceInStr - TP ' Get length of separators
     Else ' There are only separators in sFormat
       Beep
       Exit Sub ' Ignore char -> sFormat = kuk-ku
     End If
     TP = TP + iSepLenS ' Add length of separators in the Start of the sFormat to TP if any
   End If
     
   ' Check for separators one position ahead of the last char in text
   If Trim$(Mid$(sFormat, TP + 1, 1)) <> "" Then
     iSpaceInStr = InStr(TP + 1, sFormat, " ")
     If iSpaceInStr > 0 Then
       iSepLenTP = iSpaceInStr - (TP + 1) ' Get length of separators
     Else ' There are only separators left in sFormat
       iSepLenTP = Len(sFormat) - TP ' Get length of separators at end of sFormat
     End If
   ' Check for separators one position ahead of the cursor position in text
   ElseIf Trim$(Mid$(sFormat, CP + 1, 1)) <> "" Then
     iSpaceInStr = InStr(CP + 1, sFormat, " ")
     If iSpaceInStr > 0 Then
       iSepLenCP = iSpaceInStr - (CP + 1)
     Else ' There are only separators left in sFormat
       iSepLenCP = -1 ' Mark that overwrite isn't possible
     End If
   End If
   If (TP > Len(sFormat)) Then ' Overwrite char at cursor position
     TP = TP - 1 ' The text didn't get longer as we are overwriting
     If iSepLenCP = -1 Then ' Don't overwrite separators at end of sFormat
       Text1.SelStart = Len(sFormat)
       Beep
       Exit Sub ' Ignore char
     End If
     CP = CP + iSepLenCP ' If there isn't a separator, 0 (nada) will get added to CP
     Text1 = S
     Text1.SelStart = CP ' Set new cursor position
     Exit Sub
   End If
   
   Mid(S, TP, 1) = Chr(Key) ' Put in the char
 
   Text1 = S
   TP = TP + iSepLenTP ' if there isn't a separator, 0 (nada) will get added to TP
   Text1.SelStart = TP
 
 Else '***** Handle the backspace key *****
   
   If Text1.SelLength = Len(sFormat) Then ' Delete all chars in textbox
     Text1 = sFormat
     TP = 0
     Text1.SelStart = 0
     Exit Sub
   End If
   If Text1.SelLength <> 0 Then Exit Sub ' Ignore deletion if only parts of the text is selected
   If CP = 0 Then Exit Sub   ' Nothing to delete at this cursor position
   
   If CP <> TP Then Exit Sub ' Allow deletion from the back of the text only
      
   ' Check for separators at the cursor position in text
   If Trim$(Mid$(sFormat, CP, 1)) <> "" Then
     For N = CP To 1 Step -1 ' Compute the length of the separator(s)
       If Mid$(sFormat, N, 1) = " " Then Exit For
     Next N
     TP = TP - (CP - N) ' Subtract length of separators from TP
     CP = N ' Set cursor to start of separator.
   End If

   ' Only check for separators one position behind the CP in the text if there is any
   If CP > 1 Then
     ' Check for separator(s) one position behind the cursor position in the Text
     If Trim$(Mid$(sFormat, CP - 1, 1)) <> "" Then
       For N = (CP - 1) To 1 Step -1 ' Compute the length of the separator(s)
         If Mid$(sFormat, N, 1) = " " Then Exit For
       Next N
       CP = N + 1 ' Set cursor to start of separator.
     End If
   End If
 
   Mid(S, TP, 1) = " " ' Replace char at TP with " "
   Text1 = S
   TP = CP - 1
   Text1.SelStart = TP
 End If
End Sub
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' This will grey (disable) all Edit menu items in the right click popup menu
  If Button = vbRightButton Then Text1.Locked = True
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Unlock it as the Edit menu items has been disabled
  If Button = vbRightButton Then Text1.Locked = False
End Sub


