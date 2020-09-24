<div align="center">

## Close All MDI Children Simply


</div>

### Description

This code allows you to close all the MDI child forms in an MDI form at once.

Goto http://www.vbgreatone.com/ to learn like you've never learned before. Get over 500 api functions with example for each, and tips like this and, much much more. Just visit it, i'm sure you will like it.

First, create a menu item in the MDI form name mnuCloseAll, then paste in this code:
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VBGreatone \(http://www\.vbg1\.com\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vbgreatone-http-www-vbg1-com.md)
**Level**          |Intermediate
**User Rating**    |4.3 (176 globes from 41 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vbgreatone-http-www-vbg1-com-close-all-mdi-children-simply__1-5593/archive/master.zip)





### Source Code

```
Private Sub mnuCloseAll_Click()
 Screen.MousePointer = vbHourglass
 Do While Not (Me.ActiveForm Is Nothing)
  Unload Me.ActiveForm
 Loop
 Screen.MousePointer = vbDefault
End Sub
'Once the user clicks on that menu item, the MDI child forms will close.
```

