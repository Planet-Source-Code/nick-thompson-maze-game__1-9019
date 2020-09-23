Attribute VB_Name = "main"
Global levelpath As String
Global fruitx() As Integer  'Fruit x position
Global fruity() As Integer  'Fruit y position
Global fruitz() As Integer  'Fruit z position (which floor)
Global ghostx() As Integer  'Ghost x position
Global ghosty() As Integer  'Ghost y position
Global ghostz() As Integer  'Ghost z position
Global cfloors As Integer   'Cycle through floors when loading
Global a As Long            'Used for For...Next loops e.g. mapheight
Global b As Long            'Used for For...Next loops often within others e.g. mapwidth within mapheight
Global ftype As Integer
Global wtype As Integer
Global tftype As Integer
Global twtype As Integer
Global nlevels As Integer
Global buggydir As Integer
Global viewxpos As Integer
Global viewypos As Integer
Global viewzpos As Integer
Global nfruit As Integer
Global cfruit As Integer
Global mapwidth As Long
Global mapheight As Long
Global mapinfo() As Long
Global mapfloors As Long
Global password As String
Global mapname As String
Global startxpos As Integer
Global startypos As Integer
Global startzpos As Integer
Global level As Long
Global lives As Integer
Global nghosts As Integer
Global success As Integer
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008

