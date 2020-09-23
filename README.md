<div align="center">

## SetTransparent


</div>

### Description

This code makes the form Tranparent but keep the OBJECTS visible.

Orignally Written by Kalani COM

Modified by Satin Katiyar
 
### More Info
 
Refrence to form & a array containing the objects to be kept Visible

User Must Know how to pass objects.The Example of use is as give below:

Here We are assuming a command button having name command1 & Text box Having name Text1

Private Sub Form_Load()

Dim obj(1) As Object 'Use no. of controls -1 instead of 1

Set obj(0) = Command1

Set obj(1) = Text1

SetTransparent frm, obj

End Sub


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Satin Katiyar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/satin-katiyar.md)
**Level**          |Unknown
**User Rating**    |4.2 (67 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/satin-katiyar-settransparent__1-1658/archive/master.zip)

### API Declarations

```
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
 Public Const RGN_AND = 1
 Public Const RGN_COPY = 5
 Public Const RGN_DIFF = 4
 Public Const RGN_OR = 2
 Public Const RGN_XOR = 3
Type POINTAPI
 x As Long
 Y As Long
End Type
Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
```


### Source Code

```
Public Sub SetTransparent(frm As Form, obj() As Object)
 'This code was takin from a AOL Visual Basic
 'Message Board. It was submited by: SOOPRcow
 'Modified By Satin Katiyar
 Dim rctClient As RECT, rctFrame As RECT
 Dim hClient As Long, hFrame As Long, hObj As Long
 Dim Start As Integer, Finish As Integer, I As Integer
 '// Grab client area and frame area
 GetWindowRect frm.hWnd, rctFrame
 GetClientRect frm.hWnd, rctClient
 '// Convert client coordinates to screen coordinates
 Dim lpTL As POINTAPI, lpBR As POINTAPI
 lpTL.x = rctFrame.Left
 lpTL.Y = rctFrame.Top
 lpBR.x = rctFrame.Right
 lpBR.Y = rctFrame.Bottom
 ScreenToClient frm.hWnd, lpTL
 ScreenToClient frm.hWnd, lpBR
 rctFrame.Left = lpTL.x
 rctFrame.Top = lpTL.Y
 rctFrame.Right = lpBR.x
 rctFrame.Bottom = lpBR.Y
 rctClient.Left = Abs(rctFrame.Left)
 rctClient.Top = Abs(rctFrame.Top)
 rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
 rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
 rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
 rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
 rctFrame.Top = 0
 rctFrame.Left = 0
 '// Convert RECT structures to region handles
 hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
 hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
 '//Set the Scale mode of form to pixels
 Dim mode As Integer
 mode = frm.ScaleMode
 frm.ScaleMode = 3
 '// Create the new "Transparent" boundry & Add the control regions to it
 CombineRgn hFrame, hClient, hFrame, RGN_XOR
 Start = LBound(obj)
 Finish = UBound(obj)
 For I = Start To Finish
 hObj = CreateRectRgn(obj(I).Left + 4, obj(I).Top + 23, obj(I).Left + obj(I).Width + 4, obj(I).Top + obj(I).Height + 23)
 CombineRgn hFrame, hObj, hFrame, RGN_OR
 Next
 '// Now lock the window's area to this created region
 SetWindowRgn frm.hWnd, hFrame, True
 '//Restores the scale mode
 frm.ScaleMode = mode
End Sub
```

