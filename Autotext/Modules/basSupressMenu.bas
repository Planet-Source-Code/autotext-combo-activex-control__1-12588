Attribute VB_Name = "basSupressMenu"
Option Explicit

   Public lpPrevWndProc    As Long
   Public gHW              As Long
'--- End Of Hiding the Popup

Public Function gWindowProc(ByVal hwnd As Long, ByVal Msg As Long, _
                 ByVal wParam As Long, ByVal lParam As Long) As Long
   
   If Msg = WM_CONTEXTMENU Then
      gWindowProc = True ' Do not let Windows pass on the Popup Menu message to the requesting control.
   Else ' Send all other messages to the default message handler
      gWindowProc = CallWindowProc(lpPrevWndProc, hwnd, Msg, wParam, lParam)
   End If
End Function

Public Sub HookContextMenu()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf gWindowProc)
End Sub

Public Sub UnhookContextMenu()
   Dim Temp As Long
   Temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

