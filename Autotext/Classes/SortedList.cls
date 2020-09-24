VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
' Andrew J. Esser
' Create Date: Aug-15-2000
'-------------------------------------------------------------------------------
' Purpose
'-------------------------------------------------------------------------------
' This class stores items using sorted keys for fast retrieval. The preferred
' retrieval method is by key, but retrieval by index is allowed; the client
' should realize that an item's index will probably change.
'-------------------------------------------------------------------------------
Option Explicit
   Private Type Item
      Item  As Long
      Key   As Long
   End Type
   
'The actual data:
   Private mudtItems()  As Item
   Private mcItems      As Long
   
Public Property Get Count() As Long
   Count = mcItems
End Property

' Handle the error ourselves if Index is too big. We handle this ourselves
' because when the array shrinks, we will not clean up the newly deleted
' items. Therefore, an Index greater than Count might still return an
' item, but this item would be invalid.
Public Property Get ItemByIndex(Index As Long) As Long
   If Index > mcItems Then Err.Raise 9
' Let VB handle the error if Index is too small
   ItemByIndex = mudtItems(Index).Item
End Property

Public Property Get ItemByKey(Key As Long) As Long
   Dim nIndex As Long
   
   If FindItem(Key, nIndex) Then
      ItemByKey = mudtItems(nIndex).Item
   Else
      Err.Raise 5, , "There is no item with the key " & Key & "."
   End If
End Property

' Handle the error ourselves if Index is too big. We handle this ourselves
' because when the array shrinks, we will not clean up the newly deleted
' items. Therefore, and Index greater than Count might still return an
' item, but this item would be invalid.
Public Property Get KeyByIndex(Index As Long) As Long
    If Index > mcItems Then Err.Raise 9
' Let VB handle the error  if Index is too small
    KeyByIndex = mudtItems(Index).Key
End Property

Public Sub Add(Item As Long, Key As Long)
   Dim nIndex As Long
   
   If FindItem(Key, nIndex) Then
      Err.Raise 457
   Else
' Add the item at nIndex
' and increase the array.
      mcItems = mcItems + 1
      ReDim Preserve mudtItems(1 To mcItems)
' Move the items from this position to the former end.
' Only move if there are items to move.
      If nIndex < mcItems Then
' Each item takes up 8 bytes
         CopyMemory mudtItems(nIndex + 1), mudtItems(nIndex), 8 * (mcItems - nIndex)
      End If
      mudtItems(nIndex).Item = Item
      mudtItems(nIndex).Key = Key
   End If
End Sub

Public Function Remove(Key As Long) As Long
   Dim nIndex As Long
   
   If FindItem(Key, nIndex) Then
' Return the item
      Remove = mudtItems(nIndex).Item
' Move the items from this position + 1 to the end.
' Only move if there are items to move.
      If nIndex < mcItems Then
' Each item takes up 8 bytes.
         CopyMemory mudtItems(nIndex), mudtItems(nIndex + 1), 8 * (mcItems - nIndex)
      End If
' Decrease the array.
      mcItems = mcItems - 1
   Else
      Err.Raise 5, , "There is no item with the key " & Key & "."
   End If
End Function

Public Sub Clear()
   mcItems = 0
   Erase mudtItems
End Sub

' Action:  Performs a binary search on the items array
' given a key, return an index indicating either:
'    1.  The location where the item was found,
'    2.  or the location where the item should be added.
' Return True to indicate case 1 and False to indicate case 2.
Public Function FindItem(Key As Long, Optional Index As Long) As Boolean
   Dim fFound       As Boolean
   Dim nSearchPos   As Long
   Dim nLBound      As Long
   Dim nUBound      As Long
   
' Start with a field of search which includes all the items.
   nUBound = mcItems
   nLBound = 1
   
' Continue until the field of search is invalid.
   Do Until nLBound > nUBound
' Look at the item in the middle of the array
' using integer division.
      nSearchPos = (nLBound + nUBound) \ 2
      Select Case mudtItems(nSearchPos).Key
         Case Key
' Found it! Stop the loop and remember nSearchPos.
            fFound = True
            Exit Do
         Case Is < Key
' Narrow the search to the items above nSearchPos.
            nLBound = nSearchPos + 1
         Case Is > Key
' Narrow the search to the items below nSearchPos.
            nUBound = nSearchPos - 1
      End Select
   Loop
   If fFound Then
      Index = nSearchPos
      FindItem = True
   Else
' There are three cases above which would lead to this code:
'    1.  The loop never ran because mcItems = 0.
'    2.  The loop stopped because Case Is > Key raised nLBound above
'        nUBound (and above nSearchPos).
'    3.  The loop stopped because Case Is < Key lowered nUBound below
'        nLBound.
' In each case, we can determine where the new item should be added
'    1.  The new item should be added at position 1.
'    2.  The new item should be added at position nLBound.
'    3.  The new item should be added at position nSearchPos.
      If nLBound > nSearchPos Then
' This means that either case 1 or 2 occurred. Therefore the new item
' should be added at position 1 or nLBound. In case 1, nLBound = 1
' so we can simply use nLBound.
          Index = nLBound
      Else
' This means that case 3 occurred. Therefore the new item should be
' added at position nSearchPos.
          Index = nSearchPos
      End If
' FindItem = False (implicit).
   End If
End Function

