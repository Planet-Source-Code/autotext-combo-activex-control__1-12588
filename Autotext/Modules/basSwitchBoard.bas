Attribute VB_Name = "basSubClassSwitchBoard"
'***************************************************************
' Name: control subclassing switchboard
' Description:The Switchboard:A method for handling subclassing in
' ActiveX controls if you develop ActiveX controls and intend to subclass or hook
' a window, you'll very quickly discover a problem when you attempt
' to site multiple instances of your control. The subclassing,which worked
' fine With a Single instance of your control, now no longer works and is, in
' fact, most likely is causing a GPF. Why is this happening? The AddressOf operator
' requires you to place the callback routine in a module. This module is shared
' between all instances of your control and the variables and subroutines that the
' module provide are not unique to each instance. The easiest way to visualize the
' problem is to imagine a shared phoneline where multiple parties are trying to dial
' a number, talk, and hangup, all at the same time. What's needed is an operator, a
' routine that controls the dialing (hooking), the talking (the callback routine),
' and who routes information to the instance of the control that requested it. The
' Switchboard subroutine (see below) and it's supporting code provides
' method For subclassing from multiple instances of your ActiveX control. It is not
' memory intensive, nor is it slow. It's biggest weakness is that it is hardcoded to
' intercept particular messages (in this case, ComboBox messages to trap events).
'
'Inputs: None
'
'Returns: None
'
' Assumes:  Replace each instance of this with a reference to your user control.
' It is very important that your code detect and respond to a subclassed window
' when it either closes(WM_CLOSE) or is destroyed (WM_DESTROY).
' When this message is received, you should immediately unhook
' the window in question. The example code provided here does this, but knowing why it
' does it will hopefully save you some grief. Code Starts Here Because this codes hooks
' into the windows messaging system, you should not use the IDE's STOP button to
' terminate the execution of your code. Closing the form normally is mandatory.
' Debugging will become difficult once you have subclassed a window, so I recommend
' adding instancing support after the bulk of your programming work has been completed.
' As with any serious API programming tasks, you should save your project before execution.

' Side Effects: None
'***************************************************************
Option Explicit
'***************************************************************
' Windows API/Global Declarations for control subclassing
' switchboard, Mouse switchboard and context menu switchboard.
'***************************************************************
   Public Const GWL_WNDPROC = (-4&)
   Public Const WH_MOUSE = 7
   Public Const HC_ACTION     As Long = 0
   Public Const HC_GETNEXT    As Long = 1
   Public Const WM_COMMAND = &H111
   Public Const WM_SYSCOMMAND = &H112
   Public Const WM_CONTEXTMENU = &H7B
   Public Const WM_CLOSE = &H10
   Public Const WM_DESTROY = &H2
   
   Public Const MIN_INSTANCES = 1
   Public Const MAX_INSTANCES = 400
   
   'Public Const CBN_CLOSEUP = 8
   'Public Const CBN_DROPDOWN = 7
   Public Const CBN_EDITCHANGE = 5
   Public Const CBN_KILLFOCUS = 4
   Public Const CBN_SELENDOK = 9
   
   Public Type Instances
      In_Use         As Boolean  ' This instance is alive.
      ClassAddr      As Long     ' Pointer to self.
      hwnd           As Long     ' hWnd being hooked.
      PrevWndProc    As Long     ' Stored For unhooking.
      Child_hWnd     As Long     ' To return notification messages from the AutoText
   End Type                      ' combo boxes.
   
' Hooking Related Declares
'-------------
'API Declares.
'=============
   Public Declare Function CallNextHookEx Lib "user32" (ByVal hHooks As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHooks As Long) As Long
   Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
   Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
   Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
   Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
   Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
   Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
      
   Public Instances(MIN_INSTANCES To MAX_INSTANCES) As Instances
   
Public Function SwitchBoard(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim instance_check   As Integer
   Dim p_AutoText       As AutoText
   Dim PrevWndProc      As Long
   
On Error Resume Next
'Do this early as we may unhook
   PrevWndProc = Is_Hooked(hwnd)
   If Msg = WM_CLOSE Or Msg = WM_DESTROY Then
      For instance_check = MIN_INSTANCES To MAX_INSTANCES
         If Instances(instance_check).hwnd = hwnd Then
On Error Resume Next
            CopyMemory p_AutoText, Instances(instance_check).ClassAddr, 4
            p_AutoText.ParentChanged Msg
            CopyMemory p_AutoText, 0, 4
         End If
      Next instance_check
      
   ElseIf Msg = WM_COMMAND Then
      For instance_check = MIN_INSTANCES To MAX_INSTANCES
         If Instances(instance_check).hwnd = hwnd Then
            If lParam = Instances(instance_check).Child_hWnd Then
               If IsWindow(Instances(instance_check).Child_hWnd) Then
                  CopyMemory p_AutoText, Instances(instance_check).ClassAddr, 4
                  p_AutoText.ParentChanged ByVal WordHi(ByVal wParam)
                  CopyMemory p_AutoText, 0, 4
               End If
            End If
            Exit For
         End If
      Next instance_check
   
   End If
   SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
End Function

' Hooks a window or acts as if it does if the window is
' already hooked by a previous instance of myUC.
Public Sub Hook_Window(ByVal hwnd As Long, ByVal instance_ndx As Integer, ByVal Child_hWnd As Long)
   Instances(instance_ndx).PrevWndProc = Is_Hooked(hwnd)
   
   If Instances(instance_ndx).PrevWndProc = 0& Then
      Instances(instance_ndx).PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SwitchBoard)
      Instances(instance_ndx).hwnd = hwnd
      Instances(instance_ndx).Child_hWnd = Child_hWnd
   End If
   
End Sub

' Unhooks only if no other instances need the hWnd
Public Sub UnHookWindow(ByVal instance_ndx As Integer)
   If TimesHooked(Instances(instance_ndx).hwnd) = 1 Then
      SetWindowLong Instances(instance_ndx).hwnd, GWL_WNDPROC, Instances(instance_ndx).PrevWndProc
      Instances(instance_ndx).In_Use = False
      Instances(instance_ndx).PrevWndProc = 0&
      Instances(instance_ndx).hwnd = 0&
      Instances(instance_ndx).ClassAddr = 0&
      Instances(instance_ndx).Child_hWnd = 0&
   End If
   
End Sub

' Determine if we have already hooked a window,
' and returns the PrevWndProc if true, 0 if false
Private Function Is_Hooked(ByVal hwnd As Long) As Long
   Dim ndx As Integer
   
   Is_Hooked = 0
   For ndx = MIN_INSTANCES To MAX_INSTANCES
      If Instances(ndx).hwnd = hwnd Then
         Is_Hooked = Instances(ndx).PrevWndProc
         Exit For
      End If
   Next ndx
End Function

' Returns a count of the number of times a given
' window has been hooked by instances of myUC.
Private Function TimesHooked(ByVal hwnd As Long) As Long
   Dim ndx As Integer
   Dim cnt As Integer
   
   For ndx = MIN_INSTANCES To MAX_INSTANCES
      If Instances(ndx).hwnd = hwnd Then
         cnt = cnt + 1
         Exit For
      End If
   Next ndx
   TimesHooked = cnt
End Function

Private Function WordHi(ByVal LongIn As Long) As Integer
   Call CopyMemory(WordHi, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Private Function WordLo(ByVal LongIn As Long) As Integer
   Call CopyMemory(WordLo, ByVal VarPtr(LongIn), 2)
End Function

