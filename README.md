<div align="center">

## A Way to take a screenshot


</div>

### Description

This code simply takes a picture of your desktop or a screenshot. I guess if you wanted to you could use this to view the desktop of a computer over a network, although that would be a bit slow. I hope this code is of some interest to some people Hope this helps.
 
### More Info
 
'Check out my website at http://www.vbtutor.com

'Thanks

Set the Form properties to the following:

AutoRedraw True

BorderStyle 0 - None

WindowState 2 - Maximized


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve.md)
**Level**          |Unknown
**User Rating**    |3.8 (23 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-a-way-to-take-a-screenshot__1-3841/archive/master.zip)

### API Declarations

```
'Check out my site at http://www.vbtutor.com
'Thanks
Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Integer, ByVal x As Integer, _
ByVal y As Integer, ByVal nWidth As Integer, _
ByVal nHeight As Integer, ByVal _
hSrcDC As Integer, ByVal xSrc As Integer, _
ByVal ySrc As Integer, ByVal dwRop As _
Long) As Integer
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetDC Lib "user32" _
(ByVal hwnd As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
```


### Source Code

```
'Add this code to the form_load event
'or whatever you want to make it occur
'Get the hWnd of the desktop
DeskhWnd& = GetDesktopWindow()
'BitBlt needs the DC to copy the image. So, we
'need the GetDC API.
DeskDC& = GetDC(DeskhWnd&)
BitBlt Form1.hDC, 0&, 0&, _
Screen.Width, Screen.Height, DeskDC&, _
0&, 0&, SRCCOPY
'This code was requested by 1 visitor to my site
'Check out my site at http://www.vbtutor.com
'Thanks
```

