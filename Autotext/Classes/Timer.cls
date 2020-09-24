VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-------------------------------------------------------------------------------
' Timer object code has been demonstrated before. This implementation contains
' one slight improvement over some other implementations: It's global collection
' of timer objects is a keyed list. The items are pointers to the timer objects.
' The keys are timer IDs (returned by StartTimer). This enables the TimerProc in
' modTimer to very quickly find the appropriate timer object, given a timer ID.
'-------------------------------------------------------------------------------
   Public Event Timer()
   
   Private Const mnDefaultInterval As Long = 1
   
   Private mnTimerID             As Long
   Private mnInterval            As Long
   Private mfEnabled             As Boolean
   
Public Property Get Interval() As Long
   Interval = mnInterval
End Property

Public Property Let Interval(Value As Long)
   If mnInterval <> Value Then
      mnInterval = Value
      If mfEnabled Then
         SetInterval mnInterval, mnTimerID
      End If
   End If
End Property

Public Property Get Enabled() As Boolean
   Enabled = mfEnabled
End Property

Public Property Let Enabled(Value As Boolean)
   If mfEnabled <> Value Then
      If Value Then
         mnTimerID = StartTimer(mnInterval)
         If mnTimerID <> 0 Then
            mfEnabled = True
' Storing Me in the global would add a reference to Me, which
' would prevent Me from being released, which in turn would
' prevent my Class_Terminate code from running. To prevent
' this, I store a "soft reference" - the collection holds a
' pointer to me without incrementing my reference count.
            gcTimerObjects.Add ObjPtr(Me), mnTimerID
         End If
      Else
         StopTimer mnTimerID
         mfEnabled = False
         gcTimerObjects.Remove mnTimerID
      End If
   End If
End Property

Private Sub Class_Initialize()
    If gcTimerObjects Is Nothing Then Set gcTimerObjects = New SortedList
    mnInterval = mnDefaultInterval
End Sub

Private Sub Class_Terminate()
    Enabled = False
End Sub

Friend Sub Tick()
    RaiseEvent Timer
End Sub

