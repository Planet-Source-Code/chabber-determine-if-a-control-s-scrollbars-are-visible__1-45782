<div align="center">

## Determine if a control's scrollbars are visible


</div>

### Description

Use this function to determine if the scrollbars on a control are visible.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[chabber](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chabber.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chabber-determine-if-a-control-s-scrollbars-are-visible__1-45782/archive/master.zip)

### API Declarations

```
'API Constants
Private Const GWL_STYLE = (-16)
Private Const WS_HSCROLL = &H100000
Private Const WS_VSCROLL = &H200000
'API Declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
```


### Source Code

```
Private Function IsScrollBarVisible(ControlHwnd As Long) As Boolean
Dim blnResult As Boolean
Dim wndStyle As Long
  'Retrieve the window style of the control.
  wndStyle = GetWindowLong(ControlHwnd, GWL_STYLE)
  'Test if the vertical scroll bar style is present
  'in the window style, indicating that a vertical
  'scroll bar is visible.
  If (wndStyle And WS_VSCROLL) <> 0 Then
    blnResult = True
  End If
  ' Test if the horizontal scroll bar style is present
  ' in the window style, indicating that a horizontal
  ' scroll bar is visible.
  If (wndStyle And WS_HSCROLL) <> 0 Then
    blnResult = True
  End If
  IsScrollBarVisible = blnResult
End Function
```

