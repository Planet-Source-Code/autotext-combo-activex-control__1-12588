Attribute VB_Name = "basMouseHook"
'***************************************************************
' Name: control mouse hook
' Description:The MouseProc: A method for handling mouse subclassing in
' ActiveX controls. If you develop ActiveX controls and intend to subclass or hook
' a window, you'll very quickly discover a problem when you attempt
' to site multiple instances of your control. The subclassing,which worked
' fine With a Single instance of your control, now no longer works and is, in
' fact, most likely is causing a GPF. Why is this happening?
' The AddressOf operator requires you to place the callback routine in a module.
' This module is shared between all instances of your control and the variables and
' subroutines that the module provide are not unique to each instance.
' The easiest way to visualize the problem is to imagine a shared phoneline where
' multiple parties are trying to dial a number, talk, and hangup,
' all at the same time. What's needed is an operator, a routine that
' controls the dialing (hooking), the talking (the callback routine),
' and who tracks mouse information to the instance of the control that requires it. The
' mouse subroutine (see below) and it's supporting code provides
' method For subclassing from multiple instances of your ActiveX control. It is not
' memory intensive, nor does it slow down your program.
'
'Inputs: None
'
'Returns: None
'
' Assumes:  This codes hooks
' into the windows messaging system, you should not use the IDE's STOP button to
' terminate the execution of your code. Closing the form normally is mandatory.
' Debugging will become difficult once you have subclassed a window, so I recommend
' adding instancing support after the bulk of your programming work has been completed.
' As with any serious API programming tasks, you should save your project before execution.
Option Explicit
Option Base 0
'-------------------------------------------------------------------------------
'Timer APIs:
   Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
   Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
   Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

'-------------------------------------------------------------------------------
'A list of pointers to timer objects. The list uses timer IDs as the keys.

   Public gcTimerObjects As SortedList

'-------------------------------------------------------------------------------

' Windows Point API structure for tracking cursor locations.
   Public Type POINTAPI
      X     As Long
      Y     As Long
   End Type
   
''--------------
' API Constants.
''==============
' Used for Hooking into the Mouse.
   Private Const WM_MOUSEMOVE          As Long = &H200
   Private Const WM_LBUTTONDOWN        As Long = &H201
   Private Const WM_LBUTTONUP          As Long = &H202
   Private Const WM_LBUTTONDBLCLK      As Long = &H203
   Private Const WM_RBUTTONUP          As Long = &H205
   Private Const WM_RBUTTONDOWN        As Long = &H204
   Private Const WM_RBUTTONDBLCLK      As Long = &H206
   Private Const WM_MBUTTONDOWN        As Long = &H207
   Private Const WM_MBUTTONUP          As Long = &H208
   Private Const WM_MBUTTONDBLCLK      As Long = &H209
   
' Private API Declarations for tracking mouse coordinates.
   Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal X As Long, ByVal Y As Long) As Long
   Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
   
' This API will turn screen X, Y coordinates into the control's internal
' coordinates to mimick Microsofts default control behavior when mapping
' X, Y coordinate arguments of mouse events.
   Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
   
' Used to track the currently pressed mouse buttons
   Private iButtons                    As Integer
   Private hWnd                        As Long
   Private pt                          As POINTAPI
   Private ptX                         As Single
   Private ptY                         As Single
   Private instance_check              As Integer
   Private p_AutoText                  As AutoText
   Private m_AutoText                  As AutoText
   Private PrevWndProc                 As Long
   
' This array will keep track of the number of controls currently
' sub-classing the mouse
   Public instanceChk(MIN_INSTANCES To MAX_INSTANCES)    As Instances
   
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
   If nCode > HC_ACTION Then
      MouseProc = HC_GETNEXT
      
      If GetCursorPos(pt) <> 0 Then
         hWnd = WindowFromPointXY(pt.X, pt.Y) + 0
      End If
      
      If hWnd <> 0 Then
         
         For instance_check = MIN_INSTANCES To MAX_INSTANCES
            If instanceChk(instance_check).Child_hWnd = hWnd Then
               PrevWndProc = Is_Hooked(instanceChk(instance_check).hWnd)
               Exit For
            End If
         Next instance_check
      Else
         instance_check = MAX_INSTANCES + 1
      End If
      
      If instance_check <= MAX_INSTANCES And PrevWndProc <> 0 Then
         If IsWindow(instanceChk(instance_check).Child_hWnd) Then
            ScreenToClient instanceChk(instance_check).Child_hWnd, pt
            ptX = pt.X * Screen.TwipsPerPixelX
            ptY = pt.Y * Screen.TwipsPerPixelY
            
            Select Case wParam
               Case WM_MOUSEMOVE
                  CopyMemory p_AutoText, instanceChk(instance_check).ClassAddr, 4
                  Set m_AutoText = p_AutoText
                  CopyMemory p_AutoText, 0&, 4
                  Call m_AutoText.MouseMove(iButtons, ptX, ptY)
                  Set m_AutoText = Nothing
                  
               Case WM_LBUTTONDBLCLK, WM_MBUTTONDBLCLK, WM_RBUTTONDBLCLK
                  Select Case wParam
                     Case WM_LBUTTONDBLCLK
                        iButtons = (iButtons Or vbLeftButton)
                     Case WM_MBUTTONDBLCLK
                        iButtons = (iButtons Or vbMiddleButton)
                     Case WM_RBUTTONDBLCLK
                        iButtons = (iButtons Or vbRightButton)
                  End Select
                  CopyMemory p_AutoText, instanceChk(instance_check).ClassAddr, 4
                  Set m_AutoText = p_AutoText
                  CopyMemory p_AutoText, 0&, 4
                  Call m_AutoText.DblClick(iButtons, ptX, ptY)
                  Set m_AutoText = Nothing
                  
               Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
                  Select Case wParam
                     Case WM_LBUTTONDOWN
                        iButtons = (iButtons Or vbLeftButton)
                     Case WM_MBUTTONDOWN
                        iButtons = (iButtons Or vbMiddleButton)
                     Case WM_RBUTTONDOWN
                        iButtons = (iButtons Or vbRightButton)
                  End Select
                  CopyMemory p_AutoText, instanceChk(instance_check).ClassAddr, 4
                  Set m_AutoText = p_AutoText
                  CopyMemory p_AutoText, 0&, 4
                  Call m_AutoText.MouseDown(iButtons, ptX, ptY)
                  Set m_AutoText = Nothing
                  
               Case WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
                  Select Case wParam
                     Case WM_LBUTTONUP
                        iButtons = (iButtons Or vbLeftButton)
                     Case WM_MBUTTONUP
                        iButtons = (iButtons Or vbMiddleButton)
                     Case WM_RBUTTONUP
                        iButtons = (iButtons Or vbRightButton)
                  End Select
                  CopyMemory p_AutoText, instanceChk(instance_check).ClassAddr, 4
                  Set m_AutoText = p_AutoText
                  CopyMemory p_AutoText, 0&, 4
                  Call m_AutoText.MouseUp(iButtons, ptX, ptY)
                  Set m_AutoText = Nothing
                  Select Case wParam
                     Case WM_LBUTTONUP
                        iButtons = (iButtons And Not vbLeftButton)
                     Case WM_MBUTTONUP
                        iButtons = (iButtons And Not vbMiddleButton)
                     Case WM_RBUTTONUP
                        iButtons = (iButtons And Not vbRightButton)
                  End Select
            End Select
            
         End If
      End If
      If Err.Number <> 0 Then Err.Clear
   End If
   
   MouseProc = CallNextHookEx(PrevWndProc, nCode, wParam, lParam)
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Hooks a Mouse or acts as if it does if the Mouse is
' already hooked by a previous instance of myUC.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub Hook_Mouse(ByVal hWnd As Long, ByVal instance_ndx As Integer, ByVal Child_hWnd As Long)
   instanceChk(instance_ndx).PrevWndProc = Is_Hooked(hWnd)
   
   If instanceChk(instance_ndx).PrevWndProc = 0& Then
      instanceChk(instance_ndx).PrevWndProc = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0&, App.ThreadID)
      instanceChk(instance_ndx).hWnd = hWnd
      instanceChk(instance_ndx).Child_hWnd = Child_hWnd
   End If
   
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Unhooks only if no other instanceChk need the subclassed procedure.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub UnHookMouse(ByVal instance_ndx As Integer)
   Dim prevWnd    As Long
   
   prevWnd = Is_Hooked(instanceChk(instance_ndx).hWnd)
   If prevWnd <> 0 Then
      Call UnhookWindowsHookEx(prevWnd)
   End If
   instanceChk(instance_ndx).In_Use = False
   instanceChk(instance_ndx).PrevWndProc = 0&
   instanceChk(instance_ndx).hWnd = 0&
   instanceChk(instance_ndx).ClassAddr = 0&
   instanceChk(instance_ndx).Child_hWnd = 0&
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Determine if we have already hooked a Mouse,
' and returns the PrevWndProc if true, 0 if false
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function Is_Hooked(ByVal hWnd As Long) As Long
   Dim ndx As Integer
   
   Is_Hooked = 0
   For ndx = MIN_INSTANCES To MAX_INSTANCES
      If instanceChk(ndx).hWnd = hWnd Then
         Is_Hooked = instanceChk(ndx).PrevWndProc
         Exit For
      End If
   Next ndx
End Function

'-------------------------------------------------------------------------------
'The timer code:
'-------------------------------------------------------------------------------
Private Sub TimerProc(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal lTimerID As Long, ByVal lTime As Long)
   Dim nPtr             As Long
   Dim TimerObject     As Timer
   
'Create a Timer object from the pointer
   nPtr = gcTimerObjects.ItemByKey(lTimerID)
   CopyMemory TimerObject, nPtr, 4
'Call a method which will fire the Timer event
   TimerObject.Tick
'Get rid of the Timer object so that VB will not try to release it
   CopyMemory TimerObject, 0&, 4
End Sub

Public Function StartTimer(lInterval As Long) As Long
   StartTimer = SetTimer(0, 0, lInterval, AddressOf TimerProc)
End Function

Public Sub StopTimer(lTimerID As Long)
   KillTimer 0, lTimerID
End Sub

Public Sub SetInterval(lInterval As Long, lTimerID As Long)
   SetTimer 0, lTimerID, lInterval, AddressOf TimerProc
End Sub

