<div align="center">

## AutoText/Combo ActiveX Control


</div>

### Description

Autotext is an extended API implementation of Visual Basic's Combo Box. The control can be used as an autotype combobox or an autotype textbox. This code contains a proper implementation of an API timer class as well, using soft references to avoid GPFs. Allows custom context menus and allows the default context menus to be disabled. Raises a Menu_Click event which specifies which menu item was clicked. Has mouse down, mouse move and mouse up events. Code shows how to create a mouse hook. Switchboard code insures that the proper control recieves the right notifications. Has a SyncByItemData function to select an entry by an ItemData entry. Has a FillWithRecSet function to enable it to be filled directly with a recordset. You can specify the ordinal postion within the recordset that you want the drop-down list to be filled with. If you do not require the ItemData to be set, you must specify the optional default ordinal position of one to be zero. Implements an extended interface to make the control faster than the default implementation.

eg. Dim eboName As IAutoText

Set eboName = cboName.Object

I have borrowed many snippets of code from this site an thought that it was my turn to share one of my gems with the people of Planet Source Code. This ActiveX is being used by many programs and has been thouroughly tested for flaws. My thanks to all the people who have contributed to Planet Source Code.

PS. Vote if you want to.
 
### More Info
 
This code requires the ADO 2.1 data objects. (This can be set to other ADO libraries if required.)


<span>             |<span>
---                |---
**Submitted On**   |2000-11-06 21:38:14
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD114031162000\.zip](https://github.com/Planet-Source-Code/autotext-combo-activex-control__1-12588/archive/master.zip)








