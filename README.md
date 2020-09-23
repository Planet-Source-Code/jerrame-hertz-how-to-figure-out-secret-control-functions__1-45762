<div align="center">

## How to figure out secret control functions\.


</div>

### Description

You can find these hidden setting with a little work,
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jerrame Hertz](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jerrame-hertz.md)
**Level**          |Intermediate
**User Rating**    |4.7 (66 globes from 14 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jerrame-hertz-how-to-figure-out-secret-control-functions__1-45762/archive/master.zip)

### API Declarations

```
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
```


### Source Code

```
There is a code to change the background /
foreground color of progressbar here on PSC
by Juha Söderqvist , Very cool that you could
do something like this. but, How do you figure
this out? Well I can tell you how.
 These controls are written in C++ and there
properties are stored in header files, Files with
the ".h" extension. If you do a little work you
will find the information you need to pull off
these hidden features. The main thing you need to
know is the windows message constants. How did we
get to them? Here is the method I used. You have
to have C++ installed.
First I did a search in the visual studio include
folder (C:\Program Files\Microsoft Visual Studio\VC98\Include)
for "*.h" files that contained the text
"PROGRESS", I got back 90 files. The progress bar
is add with the Common Controls, So I looked for
a file that seemed to be related and found
"COMMCTRL.H". I opened it up and searched the
text for "Progress" and found the following line:
//====== PROGRESS CONTROL =====================================================
This looked like the right place, so I looked for
the naming prefix for the control. The first thing
listed was (#define PBS_SMOOTH    0x01)
PBS, So I then looked for something like back
ground color and found
(#define PBM_SETBARCOLOR   (WM_USER+9)		// lParam = bar color).
 The important part here is the definition
 (WM_USER+9), We now need to find out what is the
definition of WM_USER. I searched the text but it
was not there, so it must have been in an include
file. I then did another search in the same
directory for "define WM_USER" and only got back
one file called "WINUSER.H", In this file I
searched for the text "#define WM_USER" and found
the line
(#define WM_USER       0x0400).
Well now we have all we need to set the bar color
(ForegroundColor). 0x0400 is a hex number so we
use a calculator to get the decimal value and
come back with 1024, Now add the 9 from the
WM_USER+9 and you get the number 1033. this is
the number constant we need for the controls bar
color. Just put it into the SendMessage API
{Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long}
and send it an RGB color for the lParam variable,
and the progressbar's handle(progressbar1.hwnd)
for the hwnd variable and run the program
{lngRet = SendMessage(progressbar1.hwnd, 1033, 0, ByVal RGB(100, 255, 0))}.
Your code might look like this,
{
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
 lParam As Any) As Long
Enum PB_Colors
 PB_SetBarColor = 1033
 PB_SetBackColor = 8193
End Enum
Private Sub Form_Load()
 Dim lngRet As Long
 With ProgressBar1
  ' Set the bar color
  lngRet = SendMessage(.hwnd, PB_SetBarColor, 0, ByVal RGB(0, 255, 0))
  ' Set the back color
  lngRet = SendMessage(.hwnd, PB_SetBackColor, 0, ByVal RGB(0, 0, 0))
  .Value = 50
 End With
End Sub
}
There it is the forground color now show up as
the rgb color you sent it. A little more research
and I found the information to set the background
color. The variables I found where on lines like
these ones here.
#define PBM_SETBKCOLOR   CCM_SETBKCOLOR // lParam = bkColor
#define CCM_SETBKCOLOR   (CCM_FIRST + 1) // lParam is bkColor
#define CCM_FIRST    0x2000  // Common control shared messages
```

