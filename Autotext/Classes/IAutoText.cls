VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAutoText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
   Public Enum IStyles
      IShowCombo = 0
      IShowText = 1
   End Enum
   
Public Function IsInListExact(strItem As String) As Integer
   ' VTable Interface Definition
End Function

Public Function IsInList(strItem As String) As Integer
   ' VTable Interface Definition
End Function

Public Sub ScrollDown()
   ' VTable Interface Definition
End Sub

Public Sub ScrollUp()
   ' VTable Interface Definition
End Sub

Public Sub AddItem(ByVal sText As String)
   ' VTable Interface Definition
End Sub

Public Sub ClearContents()
   ' VTable Interface Definition
End Sub

Public Property Get ListSize() As Integer
   ' VTable Interface Definition
End Property

Public Property Let ListSize(ByVal newValue As Integer)
   ' VTable Interface Definition
End Property

Public Property Get Style() As IStyles
   ' VTable Interface Definition
End Property

Public Property Let Style(ByVal newValue As IStyles)
   ' VTable Interface Definition
End Property

Public Property Get Enabled() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get ListCount() As Integer
   ' VTable Interface Definition
End Property

Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
Attribute Text.VB_MemberFlags = "101c"
   ' VTable Interface Definition
End Property

Public Property Let Text(ByVal New_Text As String)
   ' VTable Interface Definition
End Property

Public Property Get ToolTipText() As String
   ' VTable Interface Definition
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
   ' VTable Interface Definition
End Property

Public Sub RemoveItem(ByVal Index As Integer)
   ' VTable Interface Definition
End Sub

Public Property Get SelStart() As Integer
   ' VTable Interface Definition
End Property

Public Property Let SelStart(ByVal New_SelStart As Integer)
   ' VTable Interface Definition
End Property

Public Property Get SelLength() As Integer
   ' VTable Interface Definition
End Property

Public Property Let SelLength(ByVal New_SelLength As Integer)
   ' VTable Interface Definition
End Property

Public Property Get BackColor() As OLE_COLOR
   ' VTable Interface Definition
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   ' VTable Interface Definition
End Property

Public Sub SetFocus()
   ' VTable Interface Definition
End Sub

Public Property Get Font() As Font
   ' VTable Interface Definition
End Property

Public Property Set Font(ByVal New_Font As Font)
   ' VTable Interface Definition
End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515
   ' VTable Interface Definition
End Property

Public Property Get ItemData(ByVal Index As Integer) As Long
   ' VTable Interface Definition
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
   ' VTable Interface Definition
End Property

Public Property Get List(ByVal Index As Integer) As String
   ' VTable Interface Definition
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
   ' VTable Interface Definition
End Property

Public Property Get ListIndex() As Integer
   ' VTable Interface Definition
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
   ' VTable Interface Definition
End Property

Public Property Get Locked() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get NewIndex() As Integer
   ' VTable Interface Definition
End Property

Public Sub Refresh()
   ' VTable Interface Definition
End Sub

Public Property Get SelText() As String
   ' VTable Interface Definition
End Property

Public Property Let SelText(ByVal New_SelText As String)
   ' VTable Interface Definition
End Property

Public Sub SyncByItemData(lData As Long)
   ' VTable Interface Definition
End Sub

Public Sub ShowDropDown(Optional ShowList As Boolean = True)
   ' VTable Interface Definition
End Sub

Public Function IsListDown() As Boolean
   ' VTable Interface Definition
End Function

Public Property Get FontBold() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get FontItalic() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get FontName() As String
   ' VTable Interface Definition
End Property

Public Property Let FontName(ByVal New_FontName As String)
   ' VTable Interface Definition
End Property

Public Property Get FontSize() As Single
   ' VTable Interface Definition
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
   ' VTable Interface Definition
End Property

Public Property Get FontStrikethru() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get FontUnderline() As Boolean
   ' VTable Interface Definition
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
   ' VTable Interface Definition
End Property

Public Property Get BorderStyle() As Integer
   ' VTable Interface Definition
End Property

Public Property Let BorderStyle(ByVal NewBorderStyle As Integer)
   ' VTable Interface Definition
End Property
