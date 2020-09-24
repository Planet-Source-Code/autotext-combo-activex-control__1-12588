VERSION 5.00
Begin VB.UserControl AutoText 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ClipControls    =   0   'False
   DataBindingBehavior=   1  'vbSimpleBound
   KeyPreview      =   -1  'True
   PropertyPages   =   "AutoText.ctx":0000
   ScaleHeight     =   345
   ScaleWidth      =   2940
   ToolboxBitmap   =   "AutoText.ctx":0013
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2580
   End
   Begin VB.ComboBox cbo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   0
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Begin VB.Menu mnuSub 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "AutoText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Option Base 1

Implements IAutoText
   
   Enum enuBorderStyleConstants
      [None] = 0
      [Fixed_Single] = 1
   End Enum
   
   Const LF_FACESIZE = 32
   
   Type LOGFONT
      lfHeight                   As Long
      lfWidth                    As Long
      lfEscapement               As Long
      lfOrientation              As Long
      lfWeight                   As Long
      lfItalic                   As Byte
      lfUnderline                As Byte
      lfStrikeOut                As Byte
      lfCharSet                  As Byte
      lfOutPrecision             As Byte
      lfClipPrecision            As Byte
      lfQuality                  As Byte
      lfPitchAndFamily           As Byte
      lfFaceName(LF_FACESIZE)    As Byte
   End Type

   Type RECT
      Left                       As Long
      Top                        As Long
      Right                      As Long
      Bottom                     As Long
   End Type
   
   Private Type TEXTMETRIC
      tmHeight                   As Long
      tmAscent                   As Long
      tmDescent                  As Long
      tmInternalLeading          As Long
      tmExternalLeading          As Long
      tmAveCharWidth             As Long
      tmMaxCharWidth             As Long
      tmWeight                   As Long
      tmOverhang                 As Long
      tmDigitizedAspectX         As Long
      tmDigitizedAspectY         As Long
      tmFirstChar                As Byte
      tmLastChar                 As Byte
      tmDefaultChar              As Byte
      tmBreakChar                As Byte
      tmItalic                   As Byte
      tmUnderlined               As Byte
      tmStruckOut                As Byte
      tmPitchAndFamily           As Byte
      tmCharSet                  As Byte
   End Type
   
   Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
   Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
   Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
   Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
   Private Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
   Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
   Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
   Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
   
   Private Const WM_CONTEXTMENU = &H7B
   
   Private Const GW_CHILD = 5
   Private Const GW_HWNDFIRST = 0
   Private Const GW_HWNDLAST = 1
   
' Virtual Keycode Constants
   Private Const VK_MENU = &H12     ' The Alt key both left and right
   Private Const VK_SHIFT = &H10    ' The Shift key both left and right
   Private Const VK_CONTROL = &H11  ' The Ctrl key both left and right
   
' GetSystemMetrics API Constants
   Private Const SM_CXSCREEN = 0
   Private Const SM_CYSCREEN = 1
   
' Custom styles for this control
   Public Enum Styles
      ShowCombo = 0  ' Auto-type ComboBox Feature
      ShowText = 1   ' Auto-type TextBox Feature
   End Enum
   
' Used to set the ComboBox API messages
' By not using the "As Any" argument option, we shall insure
' proper type-casting is enforced
   Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
   
' ComboBox API messages
   Private Const CB_GETITEMDATA              As Long = &H150
   Private Const CB_SETITEMDATA              As Long = &H151
   Private Const CB_FINDSTRINGEXACT          As Long = &H158
   Private Const CB_FINDSTRING               As Long = &H14C
   Private Const CB_RESETCONTENT             As Long = &H14B
   Private Const CB_ADDSTRING                As Long = &H143
   Private Const CB_DELETESTRING             As Long = &H144
   Private Const CB_GETDROPPEDSTATE          As Long = &H157
   Private Const CB_SHOWDROPDOWN             As Long = &H14F
   Private Const CB_SETEXTENDEDUI            As Long = &H155
   Private Const CB_GETLBTEXTLEN             As Long = &H149
   Private Const CB_SETDROPPEDWIDTH          As Long = &H160
   Private Const CB_GETCURSEL                As Long = &H147
   Private Const CB_SETCURSEL                As Long = &H14E
   Private Const CB_GETDROPPEDCONTROLRECT    As Long = &H152
   Private Const CBN_DROPDOWN                As Long = 7
   Private Const CBN_CLOSEUP                 As Long = 8
   
' Combo Box SendMessage API return values.
   Private Const IS_OKAY                     As Long = &H0
   Private Const IS_ERR                      As Long = (-1)
   Private Const IS_ERRSPACE                 As Long = (-2)
   Private Const ZERO                        As Long = &H0
   Private Const WM_UNDO                     As Long = &H304
   
'local variable(s) to hold property value(s)
   Private m_MousePointer                    As Integer
   Private m_AllowContextMenu                As Boolean
   Private m_MouseIcon                       As IPictureDisp
   Private m_ContextInstance                 As Integer
   Private HasHook                           As Boolean
   Private WithEvents tmr                    As Timer
Attribute tmr.VB_VarHelpID = -1
   
'Private flags
   Private m_bEditFromCode                   As Boolean
   Private m_Style                           As Integer
   Private m_NewIndex                        As Integer
   Private m_ListSize                        As Integer
   Private m_Text                            As String
   Private m_StartText                       As String
   Private IsLocked                          As Boolean
   Private IsRunTime                         As Boolean
   Private m_MyInstance                      As Integer
   Private m_LastIndex                       As Long
   Private m_ExtendedKeys                    As Integer
   Private LastKeyPressed                    As Integer
   Private m_CustomMenu                      As String
   
   Private Const m_def_Text = ""
   Private Const m_def_Style = 0
   Private Const m_def_NewIndex = -1
   Private Const m_def_ListSize = 32767
   
' Event Declarations:
   Public Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
   Public Event DblClick(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
   Public Event MenuClick(Index As Integer, sCaption As String)
   Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
   Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
   Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
   Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
   Public Event DropDown()
   Public Event OnClose()
   
'Private Sub cbo_Change(Index As Integer)
'   RaiseEvent Change
'End Sub

Private Sub cbo_Click(Index As Integer)
   RaiseEvent Click
   RaiseEvent Change
   DoEvents
End Sub

Private Sub cbo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim i          As Long
   Dim tmpCbo     As ComboBox
   
   If KeyCode = vbKeyReturn Then
      If m_Style = CInt(ShowCombo) Then
         Set tmpCbo = cbo(0)
      Else
         Set tmpCbo = cbo(1)
      End If
      With tmpCbo
         i = IsInListExact(.Text)
         If .ListIndex <> i Then
            .ListIndex = i
         End If
      End With
   End If
   RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub cbo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      Select Case m_Style
         Case ShowCombo
            With cbo(0)
               .SelStart = Len(.Text)
            End With
         Case ShowText
            With cbo(1)
               .SelStart = Len(.Text)
            End With
      End Select
   End If
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub mnuSub_Click(Index As Integer)
   RaiseEvent MenuClick(Index, mnuSub(Index).Caption)
End Sub

Private Sub tmr_Timer()
   If HasHook Then
      tmr.Interval = 10
      If m_Style = CInt(ShowCombo) Then
         gHW = GetWindow(cbo(0).hWnd, GW_CHILD)
      Else
         gHW = GetWindow(GetWindow(cbo(1).hWnd, GW_CHILD), GW_HWNDLAST)
      End If
      UnhookContextMenu
      HasHook = False
   Else
      If cbo(m_Style).ListIndex = -1 Then Exit Sub
   End If
   tmr.Enabled = False
   Set tmr = Nothing
End Sub

Private Sub UserControl_Initialize()
   m_Style = m_def_Style
   m_ListSize = m_def_ListSize
   m_Text = m_def_Text
   m_NewIndex = m_def_NewIndex
   IsLocked = False
   m_AllowContextMenu = True
   m_ContextInstance = -1
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   LastKeyPressed = KeyCode
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyShift Then        ' The Shift key
      m_ExtendedKeys = (m_ExtendedKeys And Not vbShiftMask)
   ElseIf KeyCode = vbKeyMenu Then     ' The Alt key
      m_ExtendedKeys = (m_ExtendedKeys And Not vbAltMask)
   ElseIf KeyCode = vbKeyControl Then  ' The Ctrl key
      m_ExtendedKeys = (m_ExtendedKeys And Not vbCtrlMask)
   End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Dim Instance_Scan As Integer
   
   With PropBag
      Style = .ReadProperty("Style", m_def_Style)
      ListSize = .ReadProperty("ListSize", m_def_ListSize)
      cbo(m_Style).Enabled = .ReadProperty("Enabled", True)
      m_Text = .ReadProperty("Text", m_def_Text)
      BackColor = PropBag.ReadProperty("BackColor", &H80000005)
      Set cbo(0).Font = PropBag.ReadProperty("Font0", UserControl.Ambient.Font)
      Set cbo(1).Font = PropBag.ReadProperty("Font1", UserControl.Ambient.Font)
      Set UserControl.Font = cbo(m_Style).Font
      Locked = PropBag.ReadProperty("Locked", False)
      Set m_MouseIcon = .ReadProperty("MouseIcon", Nothing)
      m_MousePointer = .ReadProperty("MousePointer", vbDefault)
      m_AllowContextMenu = .ReadProperty("AllowContextMenu", True)
      CustomMenu = .ReadProperty("CustomMenu", "")
      cbo(0).Appearance = .ReadProperty("BorderStyle", 1)
      cbo(1).Appearance = .ReadProperty("BorderStyle", 1)
   End With
   
   If UserControl.Ambient.UserMode Then
      IsRunTime = True
' Set up the subclassing hook routines
      For Instance_Scan = MIN_INSTANCES To MAX_INSTANCES
         If Instances(Instance_Scan).In_Use = False Then
            m_MyInstance = Instance_Scan
            
            Instances(Instance_Scan).In_Use = True
            Instances(Instance_Scan).ClassAddr = ObjPtr(Me)
            
            instanceChk(Instance_Scan).In_Use = True
            instanceChk(Instance_Scan).ClassAddr = ObjPtr(Me)
            
            Hook_Window UserControl.hWnd, m_MyInstance, cbo(m_Style).hWnd
            If m_Style = CInt(ShowCombo) Then
               Hook_Mouse UserControl.hWnd, m_MyInstance, GetWindow(cbo(m_Style).hWnd, GW_CHILD)
            Else
               Hook_Mouse UserControl.hWnd, m_MyInstance, GetWindow(GetWindow(cbo(m_Style).hWnd, GW_CHILD), GW_HWNDLAST)
            End If
            Exit For
         End If
         
      Next Instance_Scan
' Have the down arrow key dropping the list portion of the control instead of the default of F4.
      Call SendMessageByNum(cbo(0).hWnd, CB_SETEXTENDEDUI, CLng(1), ZERO)
   End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Style", m_Style, m_def_Style
      .WriteProperty "ListSize", m_ListSize, m_def_ListSize
      .WriteProperty "Enabled", cbo(m_Style).Enabled, True
      .WriteProperty "Text", m_Text, m_def_Text
      .WriteProperty "BackColor", cbo(m_Style).BackColor, &H80000005
      .WriteProperty "Font0", cbo(0).Font, UserControl.Ambient.Font
      .WriteProperty "Font1", cbo(1).Font, UserControl.Ambient.Font
      .WriteProperty "Locked", IsLocked, False
      .WriteProperty "MouseIcon", m_MouseIcon, Nothing
      .WriteProperty "MousePointer", m_MousePointer, vbDefault
      .WriteProperty "AllowContextMenu", m_AllowContextMenu, True
      .WriteProperty "CustomMenu", m_CustomMenu, ""
      .WriteProperty "BorderStyle", cbo(0).Appearance, 1
   End With
End Sub

Private Sub UserControl_Terminate()
   If IsRunTime Then
      UnHookMouse m_MyInstance
      UnHookWindow m_MyInstance
   End If
End Sub

Private Sub UserControl_Resize()
   Dim ctl           As ComboBox
   Static IsBusy     As Boolean
   
   If IsBusy Then Exit Sub Else IsBusy = True
   With UserControl
      DoEvents
      If .Height <> cbo(m_Style).Height + 5 Then .Height = cbo(m_Style).Height + 5
      For Each ctl In cbo
         With ctl
            .Left = 9
            .Top = 7
            .Width = UserControl.Width - 20
         End With
      Next ctl
   End With
   IsBusy = False
   RaiseEvent Resize
End Sub

' Sub is used by subclass routine.
Friend Sub ParentChanged(ByVal wMsg As Long)
   Dim i                   As Long
   Dim iLen                As Integer
   Dim tmpCbo              As ComboBox
   Dim pt                  As POINTAPI
   Dim vKeys(0 To 255)     As Byte
   Dim lRtn                As Long
   Dim hWnd                As Long
   Dim screenX             As Long
   Dim screenY             As Long
   Dim curX                As Single
   Dim curY                As Single
   Dim iShift              As Integer
   Dim iButton             As Integer
   
On Error GoTo ErrorHandler
   If LastKeyPressed = vbKeyDelete Or LastKeyPressed = vbKeyBack Or m_bEditFromCode Then
      Exit Sub
   Else
      m_bEditFromCode = True
   End If
   
   If m_Style = CInt(ShowCombo) Then
      Set tmpCbo = cbo(0)
   Else
      Set tmpCbo = cbo(1)
   End If
   
   Select Case wMsg
      Case CBN_DROPDOWN
         RaiseEvent DropDown
      Case CBN_CLOSEUP
         RaiseEvent OnClose
      Case CBN_EDITCHANGE
         With tmpCbo
            iLen = Len(.Text)
            i = SendMessageByString(.hWnd, CB_FINDSTRING, -1, ByVal .Text)
            If i <> -1 Then
               If i <> m_LastIndex And m_LastIndex <> -1 Then
                  If .Text <> Mid(.List(m_LastIndex), 1, iLen) Then
                     m_LastIndex = i
                  End If
                  SendMessageByNum cbo(m_Style).hWnd, CB_SETCURSEL, m_LastIndex, ZERO
                  'DoEvents
               Else
                  SendMessageByNum cbo(m_Style).hWnd, CB_SETCURSEL, i, ZERO
                  'DoEvents
                  m_LastIndex = i
               End If
               .Text = .List(m_LastIndex)
               .SelStart = iLen
               .SelLength = Len(.Text) '- iLen
               
            Else
               If m_LastIndex <> -1 And .ListCount > 0 Then
                  If Len(.Text) = 1 Then
                     .Text = ""
                  ElseIf .Text <> "" Then
                     SendMessageByNum cbo(m_Style).hWnd, CB_SETCURSEL, m_LastIndex, ZERO
                     'DoEvents
                     .Text = .List(m_LastIndex)
                     .SelStart = iLen - 1
                     .SelLength = Len(.Text)
                  End If
               Else
                  If Len(.Text) > 0 Then
                     .Text = Mid(.Text, 1, iLen - 1)
                  Else
                     .Text = ""
                  End If
                  .SelStart = Len(.Text)
               End If
            End If
         End With
         PropertyChanged "Text"
         RaiseEvent Change
      Case CBN_KILLFOCUS
         With tmpCbo
            If m_LastIndex <> -1 And .ListIndex <> m_LastIndex And .ListCount > 0 Then
               i = IsInListExact(.Text)
               SendMessageByNum cbo(m_Style).hWnd, CB_SETCURSEL, i, ZERO
            End If
         End With
         PropertyChanged "Text"
         'RaiseEvent Change
         
      Case CBN_SELENDOK
         m_LastIndex = cbo(0).ListIndex
         
         With tmpCbo
            If .ListIndex <> -1 Then
               m_Text = .List(.ListIndex)
            Else
               m_Text = ""
            End If
         End With
         PropertyChanged "Text"
   End Select
   m_bEditFromCode = False
   Exit Sub
ErrorHandler:
   m_bEditFromCode = False
   Err.Clear
End Sub

Friend Sub MouseDown(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
   If Button = vbRightButton And Not m_AllowContextMenu And HasHook = False Then
      If m_Style = CInt(ShowCombo) Then
         gHW = GetWindow(cbo(m_Style).hWnd, GW_CHILD)
      Else
         gHW = GetWindow(GetWindow(cbo(m_Style).hWnd, GW_CHILD), GW_HWNDLAST)
      End If
      HookContextMenu
      HasHook = True
   End If
   
   If Button = vbRightButton And m_CustomMenu <> "" And m_AllowContextMenu Then
      UserControl.PopupMenu mnu
   End If
   m_ExtendedKeys = GetShiftState
   RaiseEvent MouseDown(Button, m_ExtendedKeys, X, Y)
   
End Sub

Friend Sub MouseMove(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
   RaiseEvent MouseMove(Button, m_ExtendedKeys, X, Y)
   If cbo(m_Style).ToolTipText <> UserControl.Extender.ToolTipText Then
      cbo(0).ToolTipText = UserControl.Extender.ToolTipText & ""
      cbo(1).ToolTipText = UserControl.Extender.ToolTipText & ""
   End If
End Sub

Friend Sub MouseUp(ByVal Button As Integer, ByVal X As Single, ByVal Y As Single)
   RaiseEvent MouseUp(Button, m_ExtendedKeys, X, Y)
   If HasHook Then
      Set tmr = New Timer
      tmr.Interval = 5
      tmr.Enabled = True
   End If
End Sub

Public Function IsInListExact(strItem As String) As Integer
Attribute IsInListExact.VB_Description = "Returns a ListIndex position if an entry is found, -1  if no match is found. Will only return index for an  exact match."
   IsInListExact = SendMessageByString(cbo(m_Style).hWnd, CB_FINDSTRINGEXACT, -1, strItem)
End Function

Public Function IsInList(strItem As String) As Integer
Attribute IsInList.VB_Description = "Returns a ListIndex position if an entry is found, -1  if no match is found. Will accept partial matches."
   IsInList = SendMessageByString(cbo(m_Style).hWnd, CB_FINDSTRING, -1, strItem)
End Function

Public Sub ScrollDown()
Attribute ScrollDown.VB_Description = "Increases the list index property by one if able."
On Error Resume Next
   If IsLocked Then Exit Sub
   m_bEditFromCode = True
   
   With cbo(m_Style)
      If .ListIndex < (.ListCount - 1) And .ListCount > 0 Then
         .ListIndex = .ListIndex + 1
         m_LastIndex = .ListIndex
         m_Text = .List(.ListIndex)
         .SelStart = 0
         .SelLength = Len(.Text)
         PropertyChanged "Text"
      End If
   End With
   m_bEditFromCode = False
   If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ScrollUp()
Attribute ScrollUp.VB_Description = "Decreases the list index property by one if able."
On Error Resume Next
   If IsLocked Then Exit Sub
   m_bEditFromCode = True
   
   With cbo(m_Style)
      If .ListIndex > 0 Then
         .ListIndex = .ListIndex - 1
         m_LastIndex = .ListIndex
         m_Text = .List(.ListIndex)
         .SelStart = 0
         .SelLength = Len(.Text)
         PropertyChanged "Text"
      End If
   End With
   m_bEditFromCode = False
   If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub AddItem(ByVal sText As String)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
   If cbo(m_Style).ListCount < m_ListSize Then
      m_NewIndex = SendMessageByString(cbo(m_Style).hWnd, CB_ADDSTRING, ZERO, sText)
   End If
End Sub

Public Sub ClearContents()
Attribute ClearContents.VB_Description = "Clears the contents of a control or the system Clipboard."
   Call SendMessageByNum(cbo(m_Style).hWnd, CB_RESETCONTENT, ZERO, ZERO)
   cbo(m_Style).Text = ""
   m_Text = ""
   PropertyChanged "Text"
End Sub

Public Property Get ListSize() As Integer
Attribute ListSize.VB_Description = "Returns/sets the maximum number of list entries allowed in the control."
Attribute ListSize.VB_ProcData.VB_Invoke_Property = ";Scale"
   ListSize = m_ListSize
End Property

Public Property Let ListSize(ByVal newValue As Integer)
   m_ListSize = newValue
End Property

Public Property Get Style() As Styles
Attribute Style.VB_Description = "Returns/Sets whether the control behaves as a text box or combo box."
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Misc"
   Style = m_Style
End Property

Public Property Let Style(ByVal newValue As Styles)
   Dim Instance_Scan As Integer
   
On Error Resume Next
   m_Style = CInt(newValue)
   Select Case m_Style
      Case ShowCombo
         If cbo(0).ListCount > 0 Then
            If cbo(0).ListIndex <> cbo(1).ListIndex And cbo(1).ListIndex <> -1 Then cbo(0).ListIndex = cbo(1).ListIndex
         End If
         cbo(0).Visible = True
         cbo(1).Visible = False
         If Err.Number <> 0 Then Err.Clear
      Case ShowText
         If cbo(1).ListCount > 0 Then
            If cbo(1).ListIndex <> cbo(0).ListIndex And cbo(0).ListIndex <> -1 Then cbo(1).ListIndex = cbo(0).ListIndex
         End If
         cbo(1).Visible = True
         cbo(0).Visible = False
         If Err.Number <> 0 Then Err.Clear
   End Select
   If m_MyInstance <> 0 Then
      UnHookWindow m_MyInstance
      UnHookMouse m_MyInstance
      
      For Instance_Scan = MIN_INSTANCES To MAX_INSTANCES
         If Instances(Instance_Scan).In_Use = False Then
            m_MyInstance = Instance_Scan
            Instances(Instance_Scan).In_Use = True
            Instances(Instance_Scan).ClassAddr = ObjPtr(Me)
            
            instanceChk(Instance_Scan).In_Use = True
            instanceChk(Instance_Scan).ClassAddr = ObjPtr(Me)
            
            Hook_Window UserControl.hWnd, m_MyInstance, cbo(m_Style).hWnd
            If m_Style = CInt(ShowCombo) Then
               Hook_Mouse UserControl.hWnd, m_MyInstance, GetWindow(cbo(m_Style).hWnd, GW_CHILD)
            Else
               Hook_Mouse UserControl.hWnd, m_MyInstance, GetWindow(cbo(m_Style).hWnd, GW_HWNDLAST)
            End If
            Exit For
         End If
      Next Instance_Scan
   End If
   
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute ListIndex.VB_MemberFlags = "400"
   ListIndex = SendMessageByNum(cbo(m_Style).hWnd, CB_GETCURSEL, ZERO, ZERO)
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Long)
   If cbo(m_Style).ListIndex = New_ListIndex Then Exit Property
   If New_ListIndex < -1 Then New_ListIndex = -1
   SendMessageByNum cbo(m_Style).hWnd, CB_SETCURSEL, New_ListIndex, ZERO
   With cbo(m_Style)
      If New_ListIndex <> -1 Then
         m_Text = .List(New_ListIndex)
      Else
         m_Text = ""
      End If
   End With
   PropertyChanged "ListIndex"
   PropertyChanged "Text"
   cbo_Click m_Style
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Locked = cbo(m_Style).Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   cbo(m_Style).Locked() = New_Locked
   IsLocked = New_Locked
   PropertyChanged "Locked"
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific four byte Integer for each item in a ComboBox or ListBox control."
Attribute ItemData.VB_ProcData.VB_Invoke_Property = ";List"
   ItemData = SendMessageByNum(cbo(m_Style).hWnd, CB_GETITEMDATA, Index, ZERO)
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal New_ItemData As Long)
   Dim lRtn    As Long
   lRtn = SendMessageByNum(cbo(m_Style).hWnd, CB_SETITEMDATA, Index, New_ItemData)
   If lRtn <> -1 Then
      PropertyChanged "ItemData"
   Else
      Err.Raise vbObjectError + 362, "AutoText.ItemData Property", "The Index entry of (" & Index & ") does not exist."
   End If
End Property

Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
Attribute List.VB_ProcData.VB_Invoke_Property = ";List"
   List = cbo(m_Style).List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
On Error GoTo ErrorHandler
   With cbo(m_Style)
      .List(Index) = New_List
      If .ListIndex <> -1 Then
         m_Text = .List(.ListIndex)
      Else
         m_Text = ""
      End If
   End With
   PropertyChanged "List"
   PropertyChanged "Text"
   Exit Property
ErrorHandler:
   Err.Raise Err.Number, "AutoText.List Property", Err.Description
End Property

Public Sub CboAdjustAuto(Optional CharCount As Long = 0)
Attribute CboAdjustAuto.VB_Description = "Automatically resizes the drop-list width to the largest list value."
   Dim r                      As Long
   Dim i                      As Long
   Dim NumOfChars             As Long
   Dim LongestComboItem       As Long
   Dim avgCharWidth           As Long
   Dim NewDropDownWidth       As Long
   Dim cnt                    As Integer
   Dim tm                     As TEXTMETRIC
   Dim oldFont                As Long
   Dim theFont                As Long
   Dim di                     As Long
   Dim lf                     As LOGFONT
   Dim TempByteArray()        As Byte
   Dim X                      As Integer
   Dim ByteArrayLimit         As Long
   
   Const OUT_DEFAULT_PRECIS = 0
   Const DEFAULT_QUALITY = 0
   Const DEFAULT_PITCH = 0
   Const FF_DONTCARE = 0
   Const DEFAULT_CHARSET = 1
   Const TMPF_FIXED_PITCH = 1
   
   Const CHARS = "A"
   If m_Style = CInt(ShowText) Then Exit Sub
   If UserControl.Font <> cbo(m_Style).Font Then Set UserControl.Font = cbo(m_Style).Font
' Loop through the combo entries, using SendMessage
' with CB_GETLBTEXTLEN to determine the longest item
' in the dropdown portion of the combo
   If CharCount = 0 Then
      cnt = ListCount - 1
      For i = 0 To cnt
         NumOfChars = SendMessageByNum(cbo(0).hWnd, CB_GETLBTEXTLEN, i, 0)
         If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars
      Next i
   Else
      LongestComboItem = CharCount
   End If
' Get the average size of the characters using the GetFontDialogUnits API.
' Because a dummy string is used in GetFontDialogUnits, avgCharWidth is
' an approximation based on that string.
   i = ScaleMode
   ScaleMode = vbPixels
   'avgCharWidth = TextWidth(CHARS) '/ Len(CHARS)
   
   With lf
      .lfHeight = TextHeight(CHARS)
      .lfWidth = TextWidth(CHARS)
      .lfEscapement = 0
      .lfWeight = UserControl.Font.Weight
      If (UserControl.Font.Bold) Then lf.lfItalic = 1
      If (UserControl.Font.Underline) Then lf.lfUnderline = 1
      If (UserControl.Font.Strikethrough) Then lf.lfStrikeOut = 1
      lf.lfOutPrecision = OUT_DEFAULT_PRECIS
      lf.lfClipPrecision = OUT_DEFAULT_PRECIS
      lf.lfQuality = DEFAULT_QUALITY
      lf.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
      lf.lfCharSet = DEFAULT_CHARSET
      TempByteArray = StrConv(UserControl.Font.Name & Chr$(0), vbFromUnicode)
      ByteArrayLimit = UBound(TempByteArray)
      For X = 0 To ByteArrayLimit
         lf.lfFaceName(X) = TempByteArray(X)
      Next X
   End With
   
   theFont = CreateFontIndirect(lf)
   oldFont = SelectObject(UserControl.hdc, theFont)
   di = GetTextMetrics(UserControl.hdc, tm)
   
   If (tm.tmPitchAndFamily And TMPF_FIXED_PITCH) = 0 Then
      avgCharWidth = tm.tmMaxCharWidth
      NewDropDownWidth = (LongestComboItem * avgCharWidth) + 2
   Else
      avgCharWidth = tm.tmAveCharWidth
      NewDropDownWidth = ((LongestComboItem - 2) * avgCharWidth) + 18
      
   End If
   
   di = SelectObject(UserControl.hdc, oldFont)
   DeleteObject (theFont)
   
' Resize the dropdown portion of the combo box
   r = SendMessageByNum(cbo(0).hWnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, ZERO)
   DoEvents
   ScaleMode = i
End Sub

Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control.\n"
Attribute ListCount.VB_ProcData.VB_Invoke_Property = ";Scale"
   Select Case m_Style
      Case ShowCombo
         ListCount = cbo(0).ListCount
      Case ShowText
         ListCount = cbo(1).ListCount
   End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=0,0,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
   Enabled = cbo(m_Style).Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   cbo(m_Style).Enabled() = New_Enabled
   PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
   UserControl.Refresh
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "122c"
   Text = cbo(m_Style).Text
End Property

Public Property Let Text(ByVal New_Text As String)
   If CanPropertyChange("Text") Then
      If m_StartText <> "" Then
         m_Text = New_Text
         m_bEditFromCode = True
         With cbo(m_Style)
            m_LastIndex = IsInList(m_Text)
            .ListIndex = m_LastIndex
            .Text = New_Text
         End With
         m_bEditFromCode = False
         PropertyChanged "Text"
      Else
         If New_Text = "" Then Exit Property
         m_StartText = New_Text
         m_Text = New_Text
         cbo(m_Style).Text = New_Text
         Set tmr = New Timer
         tmr.Enabled = True
      End If
   End If
End Property

Public Sub RemoveItem(ByVal Index As Integer)
Attribute RemoveItem.VB_Description = "Removes an item from a ListBox or ComboBox control or a row from a Grid control."
   Dim r    As Long
   
   r = SendMessageByNum(cbo(m_Style).hWnd, CB_DELETESTRING, Index, ZERO)
   With cbo(m_Style)
      If .ListIndex <> -1 Then
         m_Text = .List(.ListIndex)
      Else
         m_Text = ""
      End If
   End With
   PropertyChanged "Text"
   If r = IS_ERR Then Err.Raise vbObjectError + 84, "SUB: RemoveItem(Index As Integer)", "Index for removal is invalid"
   
End Sub

Public Property Get SelStart() As Integer
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute SelStart.VB_MemberFlags = "400"
   SelStart = cbo(m_Style).SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Integer)
   cbo(m_Style).SelStart = New_SelStart
   PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Integer
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_ProcData.VB_Invoke_Property = ";Scale"
Attribute SelLength.VB_MemberFlags = "400"
   SelLength = cbo(m_Style).SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Integer)
   cbo(m_Style).SelLength = New_SelLength
   PropertyChanged "SelLength"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute SelText.VB_MemberFlags = "400"
   SelText = cbo(m_Style).SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
   cbo(m_Style).SelText() = New_SelText
   PropertyChanged "SelText"
   m_Text = cbo(m_Style).Text
   PropertyChanged "Text"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = " Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
   BackColor = cbo(m_Style).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   cbo(m_Style).BackColor = New_BackColor
   PropertyChanged "BackColor"
End Property

Public Sub SetFocus()
Attribute SetFocus.VB_Description = "Moves the focus to the specified object."
On Error Resume Next
   With cbo(m_Style)
      If .Visible Then .SetFocus
   End With
   If Err.Number <> 0 Then Err.Clear
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = " Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
   Set Font = cbo(m_Style).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set cbo(m_Style).Font = New_Font
   Set UserControl.Font = New_Font
   PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
   FontBold = cbo(m_Style).Font.Bold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
   cbo(m_Style).Font.Bold() = New_FontBold
   UserControl_Resize
   PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
   FontItalic = cbo(m_Style).Font.Italic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
   cbo(m_Style).Font.Italic() = New_FontItalic
   UserControl_Resize
   PropertyChanged "FontItalic"
End Property

Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
   FontName = cbo(m_Style).Font.Name
End Property

Public Property Let FontName(ByVal New_FontName As String)
   cbo(m_Style).Font.Name() = New_FontName
   UserControl_Resize
   PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
   FontSize = cbo(m_Style).Font.Size
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
   cbo(m_Style).Font.Size() = New_FontSize
   UserControl_Resize
   PropertyChanged "FontSize"
End Property

Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = " Returns/sets strikethrough font styles."
Attribute FontStrikethru.VB_ProcData.VB_Invoke_Property = ";Font"
   FontStrikethru = cbo(m_Style).Font.Strikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
   cbo(m_Style).Font.Strikethru() = New_FontStrikethru
   UserControl_Resize
   PropertyChanged "FontStrikethru"
End Property

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles"
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
   FontUnderline = cbo(m_Style).Font.Underline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
   cbo(m_Style).Font.Underline() = New_FontUnderline
   UserControl_Resize
   PropertyChanged "FontUnderline"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute hWnd.VB_UserMemId = -515
Attribute hWnd.VB_MemberFlags = "400"
   hWnd = UserControl.hWnd
End Property

Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = " Returns the index of the item most recently added to a control."
Attribute NewIndex.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute NewIndex.VB_MemberFlags = "400"
   NewIndex = m_NewIndex
End Property

Public Sub FillWithRecSet(rs As ADODB.Recordset, Optional rsOrdPos As Integer = 1)
Attribute FillWithRecSet.VB_Description = "Fills control with the contents of an ADODB.Recordset object. The first field in the recordset should contain a four byte Integer. If the recordset has only one field, 0 should be specified for OrdPos argument."
   Dim lngRtn                 As Long
   Dim cnt                    As Integer
   Dim hWnd                   As Long
   Dim New_Text               As String
   Dim NumOfChars             As Long
   Dim LongestComboItem       As Long
   
   Const RECSET_NOT_READY = 91
   
On Error GoTo ErrorHandler
   If rs.State <> adStateOpen Then
      Call SendMessageByNum(hWnd, CB_RESETCONTENT, ZERO, ZERO)
      Exit Sub
   End If
   
   m_bEditFromCode = True
   If rs.Supports(adMovePrevious) Then
      If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
   End If
   
   hWnd = cbo(m_Style).hWnd
' Clear the current contents
   Call SendMessageByNum(hWnd, CB_RESETCONTENT, ZERO, ZERO)
   hWnd = cbo(m_Style).hWnd
   
   Do Until rs.EOF
      cnt = cnt + 1
      If ListSize < cnt Then ListSize = cnt
' Fill Combo Box
      If Not IsNull(rs(rsOrdPos).Value) Then
         NumOfChars = Len(rs(rsOrdPos).Value)
         lngRtn = SendMessageByString(hWnd, CB_ADDSTRING, ZERO, rs(rsOrdPos).Value)
      Else
         lngRtn = SendMessageByString(hWnd, CB_ADDSTRING, ZERO, "")
      End If
      If lngRtn >= IS_OKAY Then
         If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars
         If rsOrdPos <> 0 Then
' Fill the ItemData
            If Not IsNull(rs(0).Value) Then
               lngRtn = SendMessageByNum(hWnd, CB_SETITEMDATA, lngRtn, rs(0))
            Else
               lngRtn = SendMessageByNum(hWnd, CB_SETITEMDATA, lngRtn, 0)
            End If
            If lngRtn = IS_ERR Then
               MsgBox "Couldn't add all the specified entries to the drop-down box." _
                    & vbCrLf & "Cancelling the current procedure.!!", _
                    vbCritical + vbOKOnly, "ERROR IN FILLING COMBOBOX!!!"
               Exit Do
            End If
         End If
      ElseIf lngRtn = IS_ERR Then
         MsgBox "Couldn't add all the specified entries to the drop-down box." _
                 & vbCrLf & "Cancelling the current procedure.!!", _
                 vbCritical + vbOKOnly, "ERROR IN FILLING COMBOBOX!!!"
         Exit Do
      ElseIf lngRtn = IS_ERRSPACE Then
         MsgBox "Not enouph memory to add all entries to the current drop-down box." _
                 & vbCrLf & "Try closing another application if you have one running!!" _
                 & vbCrLf & "Cancelling the current procedure.!!", _
                 vbCritical + vbOKOnly, "ERROR IN FILLING COMBOBOX!!!"
         Exit Do
      Else
         Exit Do
      End If
      rs.MoveNext
   Loop
   
   CboAdjustAuto LongestComboItem
   DoEvents
   With cbo(m_Style)
      .ListIndex = IsInList(m_Text)
      If .ListIndex <> -1 Then
         .Text = m_Text
      Else
         .Text = ""
      End If
   End With
   m_bEditFromCode = False
   Exit Sub
ErrorHandler:
   m_Text = cbo(m_Style).Text
   PropertyChanged "Text"
   m_bEditFromCode = False
   If Err.Number = RECSET_NOT_READY Then
      Err.Clear
   Else
      Err.Raise Err.Number, "FillComboWithRecSet--Sub", Err.Description
   End If
End Sub

Public Sub SyncByItemData(lData As Long)
Attribute SyncByItemData.VB_Description = "Selects an list entry if an ItemData match is found."
   Dim i       As Integer
   Dim r       As Long
   Dim hWnd    As Long
   Dim cnt     As Integer

On Error GoTo ErrorHandler
   m_bEditFromCode = True
   With cbo(m_Style)
      hWnd = .hWnd
      cnt = .ListCount - 1
      For i = 0 To cnt
         r = SendMessageByNum(hWnd, CB_GETITEMDATA, i, ZERO)
         If r = lData Then
            .ListIndex = i
            m_LastIndex = i
            m_Text = .List(i)
            PropertyChanged "Text"
            Exit For
         End If
      Next i
   End With
   m_bEditFromCode = False
   Exit Sub
ErrorHandler:
   Err.Clear
   m_bEditFromCode = False
End Sub

Public Function IsListDown() As Boolean
Attribute IsListDown.VB_Description = "Returns True if the drop-list portion of the control is visible."
   IsListDown = (SendMessageByString(cbo(m_Style).hWnd, CB_GETDROPPEDSTATE, ZERO, ZERO) <> False)
End Function

Public Sub ShowDropDown(Optional ShowList As Boolean = True)
Attribute ShowDropDown.VB_Description = "Opens or closes the dop-list portion of the control. This is only available when the Style property is set to ShowCombo."
   SendMessageByNum cbo(m_Style).hWnd, CB_SHOWDROPDOWN, ShowList, False
End Sub

Public Property Set MouseIcon(ByVal vData As IPictureDisp)
   Set m_MouseIcon = vData
   Set cbo(0).MouseIcon = m_MouseIcon
   Set cbo(1).MouseIcon = m_MouseIcon
   PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
   Set MouseIcon = m_MouseIcon
End Property

Public Property Let MousePointer(ByVal vData As MousePointerConstants)
   m_MousePointer = CInt(vData)
   cbo(0).MousePointer = m_MousePointer
   cbo(1).MousePointer = m_MousePointer
   PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
   MousePointer = m_MousePointer
End Property

Public Property Get AllowContextMenu() As Boolean
Attribute AllowContextMenu.VB_ProcData.VB_Invoke_Property = "pgCustomMenu"
   AllowContextMenu = m_AllowContextMenu
End Property

Public Property Let AllowContextMenu(ByVal bNewValue As Boolean)
   m_AllowContextMenu = bNewValue
   PropertyChanged "AllowContextMenu"
End Property

Friend Sub DblClick(Button As Integer, X As Single, Y As Single)
   Static IsBusy As Byte

   Select Case IsBusy
      Case 0
         IsBusy = 1
         Exit Sub
      Case Is = 1
         IsBusy = 2
         Exit Sub
   End Select
   IsBusy = 0
   RaiseEvent DblClick(Button, m_ExtendedKeys, X, Y)
End Sub

Public Property Get CustomMenu() As String
Attribute CustomMenu.VB_ProcData.VB_Invoke_Property = "pgCustomMenu"
   CustomMenu = m_CustomMenu
End Property

Public Property Let CustomMenu(sNewValue As String)
   m_CustomMenu = sNewValue
   If InStr(m_CustomMenu, "~") = 0 Then
      AddSubMnuItem Split(m_CustomMenu)
   Else
      AddSubMnuItem Split(m_CustomMenu, "~")
   End If
   PropertyChanged "CustomMenu"
End Property

Private Sub AddSubMnuItem(pArray As Variant)
   Dim X             As Long
   Dim Y             As Long
   Dim i             As Integer
   
   If IsArray(pArray) Then
      i = UBound(pArray)
   ElseIf pArray <> "" Then
      i = 0
   Else
      i = -1
   End If
   
   If mnuSub(0).Caption <> "" Then
      Y = mnuSub.Count - 1
      For X = Y To 1 Step -1
         Unload mnuSub(X)
      Next
      mnuSub(0).Caption = ""
   End If
   
   If i <> -1 Then
      With mnuSub
         For X = 0 To i
            If Len(mnuSub(0).Caption) = 0 Then
               .Item(0).Caption = pArray(X)
            Else
               Load mnuSub(.UBound + 1)
               .Item(.UBound).Caption = pArray(X)
            End If
         Next X
      End With
   End If
End Sub

Private Function GetShiftState() As Integer
   Dim curShift               As Integer
   Dim bytShift(0 To 255)     As Byte
   
   If GetKeyboardState(bytShift(0)) > 0 Then
      If bytShift(VK_SHIFT) >= 128 Then
         curShift = (curShift Or vbShiftMask)
      End If
      
      If bytShift(VK_CONTROL) >= 128 Then
         curShift = (curShift Or vbCtrlMask)
      End If
      
      If bytShift(VK_MENU) >= 128 Then
         curShift = (curShift Or vbAltMask)
      End If
      
   End If
   GetShiftState = curShift
End Function

Public Property Get BorderStyle() As enuBorderStyleConstants
   BorderStyle = cbo(0).Appearance
End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As enuBorderStyleConstants)
   cbo(0).Appearance = CInt(NewBorderStyle)
   cbo(1).Appearance = CInt(NewBorderStyle)
End Property

'********************************************************
'~~~~~~~~~~~IAutoText Interface definitions~~~~~~~~~~~~~~
'********************************************************
Private Function IAutoText_IsListDown() As Boolean
   IAutoText_IsListDown = IsListDown
End Function

Private Sub IAutoText_ShowDropDown(Optional ShowList As Boolean = True)
   ShowDropDown ShowList
End Sub

Private Sub IAutoText_SyncByItemData(lData As Long)
   SyncByItemData lData
End Sub

Private Sub IAutoText_AddItem(ByVal sText As String)
   AddItem sText
End Sub

Private Property Let IAutoText_BackColor(ByVal RHS As stdole.OLE_COLOR)
   BackColor = RHS
End Property

Private Property Get IAutoText_BackColor() As stdole.OLE_COLOR
   IAutoText_BackColor = BackColor
End Property

Private Sub IAutoText_ClearContents()
   ClearContents
End Sub

Private Property Let IAutoText_Enabled(ByVal RHS As Boolean)
   Enabled = RHS
End Property

Private Property Get IAutoText_Enabled() As Boolean
   IAutoText_Enabled = Enabled
End Property

Private Property Set IAutoText_Font(ByVal RHS As stdole.Font)
   Set Font = RHS
End Property

Private Property Get IAutoText_Font() As stdole.Font
   Set IAutoText_Font = Font
End Property

Private Property Let IAutoText_ListSize(ByVal RHS As Integer)
   ListSize = RHS
End Property

Private Property Get IAutoText_ListSize() As Integer
   IAutoText_ListSize = ListSize
End Property

Private Property Get IAutoText_hWnd() As Long
   IAutoText_hWnd = hWnd
End Property

Private Function IAutoText_IsInList(strItem As String) As Integer
   IAutoText_IsInList = IsInList(strItem)
End Function

Private Function IAutoText_IsInListExact(strItem As String) As Integer
   IAutoText_IsInListExact = IsInListExact(strItem)
End Function

Private Property Let IAutoText_ItemData(ByVal Index As Integer, ByVal RHS As Long)
   ItemData(Index) = RHS
End Property

Private Property Get IAutoText_ItemData(ByVal Index As Integer) As Long
   IAutoText_ItemData = ItemData(Index)
End Property

Private Property Let IAutoText_List(ByVal Index As Integer, ByVal RHS As String)
   List(Index) = RHS
End Property

Private Property Get IAutoText_List(ByVal Index As Integer) As String
   IAutoText_List = List(Index)
End Property

Private Property Get IAutoText_ListCount() As Integer
   IAutoText_ListCount = ListCount
End Property

Private Sub IAutoText_ScrollDown()
   ScrollDown
End Sub

Private Property Let IAutoText_ListIndex(ByVal RHS As Integer)
   ListIndex = RHS
End Property

Private Property Get IAutoText_ListIndex() As Integer
   IAutoText_ListIndex = ListIndex
End Property

Private Sub IAutoText_ScrollUp()
   ScrollUp
End Sub

Private Property Let IAutoText_Locked(ByVal RHS As Boolean)
   Locked = RHS
End Property

Private Property Get IAutoText_Locked() As Boolean
   IAutoText_Locked = Locked
End Property

Private Property Get IAutoText_NewIndex() As Integer
   IAutoText_NewIndex = NewIndex
End Property

Private Sub IAutoText_Refresh()
   Refresh
End Sub

Private Sub IAutoText_RemoveItem(ByVal Index As Integer)
   RemoveItem Index
End Sub

Private Property Let IAutoText_SelLength(ByVal RHS As Integer)
   SelLength = RHS
End Property

Private Property Get IAutoText_SelLength() As Integer
   IAutoText_SelLength = SelLength
End Property

Private Property Let IAutoText_SelStart(ByVal RHS As Integer)
   SelStart = RHS
End Property

Private Property Get IAutoText_SelStart() As Integer
   IAutoText_SelStart = SelStart
End Property

Private Property Let IAutoText_SelText(ByVal RHS As String)
   Me.SelText = RHS
End Property

Private Property Get IAutoText_SelText() As String
   IAutoText_SelText = SelText
End Property

Private Sub IAutoText_SetFocus()
   SetFocus
End Sub

Private Property Let IAutoText_Style(ByVal RHS As IStyles)
   Style = RHS
End Property

Private Property Get IAutoText_Style() As IStyles
   IAutoText_Style = Style
End Property

Private Property Let IAutoText_Text(ByVal RHS As String)
   Text = RHS
End Property

Private Property Get IAutoText_Text() As String
   IAutoText_Text = Text
End Property

Private Property Let IAutoText_ToolTipText(ByVal RHS As String)
   UserControl.Extender.ToolTipText = RHS
End Property

Private Property Get IAutoText_ToolTipText() As String
   IAutoText_ToolTipText = UserControl.Extender.ToolTipText & ""
End Property

Private Property Let IAutoText_FontBold(ByVal RHS As Boolean)
   Me.FontBold = RHS
End Property

Private Property Get IAutoText_FontBold() As Boolean
   IAutoText_FontBold = Me.FontBold
End Property

Private Property Let IAutoText_FontItalic(ByVal RHS As Boolean)
   Me.FontItalic = RHS
End Property

Private Property Get IAutoText_FontItalic() As Boolean
   IAutoText_FontItalic = Me.FontItalic
End Property

Private Property Let IAutoText_FontName(ByVal RHS As String)
   Me.FontName = RHS
End Property

Private Property Get IAutoText_FontName() As String
   IAutoText_FontName = Me.FontName
End Property

Private Property Let IAutoText_FontSize(ByVal RHS As Single)
   Me.FontSize = RHS
End Property

Private Property Get IAutoText_FontSize() As Single
   IAutoText_FontSize = Me.FontSize
End Property

Private Property Let IAutoText_FontStrikethru(ByVal RHS As Boolean)
   Me.FontStrikethru = RHS
End Property

Private Property Get IAutoText_FontStrikethru() As Boolean
   IAutoText_FontStrikethru = Me.FontStrikethru
End Property

Private Property Let IAutoText_FontUnderline(ByVal RHS As Boolean)
   Me.FontUnderline = RHS
End Property

Private Property Get IAutoText_FontUnderline() As Boolean
   IAutoText_FontUnderline = Me.FontUnderline
End Property

Private Property Let IAutoText_BorderStyle(ByVal RHS As Integer)
   Me.BorderStyle = RHS
End Property

Private Property Get IAutoText_BorderStyle() As Integer
   IAutoText_BorderStyle = CInt(Me.BorderStyle)
End Property
