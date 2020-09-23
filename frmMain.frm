VERSION 5.00
Begin VB.Form frmMain1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "System Information"
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4140
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Computer name"
      Height          =   255
      Left            =   180
      TabIndex        =   19
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   18
      Top             =   180
      Width           =   1905
   End
   Begin VB.Label lblIP 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   17
      Top             =   2700
      Width           =   1905
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address"
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   2700
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RAM (Available)"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "RAM (Total)"
      Height          =   255
      Left            =   180
      TabIndex        =   14
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Processor"
      Height          =   255
      Left            =   180
      TabIndex        =   13
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Processor Vendor"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label lblRamAvail 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   11
      Top             =   2340
      Width           =   1905
   End
   Begin VB.Label lblRamTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   10
      Top             =   2100
      Width           =   1905
   End
   Begin VB.Label lblNormSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   9
      Top             =   1860
      Width           =   1905
   End
   Begin VB.Label lblRawSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   8
      Top             =   1620
      Width           =   1905
   End
   Begin VB.Label lblProcessor 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   7
      Top             =   1380
      Width           =   2295
   End
   Begin VB.Label lblVendor 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   6
      Top             =   1140
      Width           =   2295
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   5
      Top             =   900
      Width           =   2295
   End
   Begin VB.Label lblOS 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   2220
      TabIndex        =   4
      Top             =   660
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Processor Speed(Normal)"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   1860
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Processor Speed(Raw)"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Service Pack"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   900
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OS"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   660
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pPidl As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, ByVal dwItem1 As Long, ByVal dwItem2 As Long)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const CSIDL_DESKTOP = &H0
Private Const SHCNF_IDLIST = &H0
Private Const SHCNE_ALLEVENTS = &H7FFFFFFF
Private Const RDW_ERASENOW = &H200
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
Private Const REG_SZ = 1  'Unicode nul terminated string
Private Const REG_BINARY = 3  'Free form binary
Private Const REG_DWORD = 4  '32-bit number
Private Const ERROR_SUCCESS = 0&

Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Public Sub GetInfo()

Dim memoryInfo As MEMORYSTATUS
Dim sCpu As String, sVendor As String
Dim sL2Cache As String
Dim sRawSpeed As String
Dim sNormSpeed As String
Dim dl&, s$
Dim mySys As SYSTEM_INFO

  GlobalMemoryStatus memoryInfo
  lblRamTotal.Caption = Round(memoryInfo.dwTotalPhys / 1043321, 0)
  lblRamAvail.Caption = Round(memoryInfo.dwAvailPhys / 1043321, 0)
  
  sCpu = String(255, 0)
  sRawSpeed = String(255, 0)
  sNormSpeed = String(255, 0)
  sVendor = String(255, 0)
  sL2Cache = String(255, 0)
  
  GetProcessor sCpu, sVendor, sL2Cache
  GetProcessorRawSpeed sRawSpeed
  GetProcessorNormSpeed sNormSpeed

  lblProcessor.Caption = StripZero(sCpu)
  lblVendor.Caption = StripZero(sVendor)
  lblRawSpeed.Caption = StripZero(sRawSpeed)
  lblNormSpeed.Caption = StripZero(sNormSpeed)
  
  myVer.dwOSVersionInfoSize = 148
  dl& = GetVersionEx&(myVer)
  s$ = LPSTRToVBString(myVer.szCSDVersion)
  
  lblSP.Caption = s$
  
  If myVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    s$ = "Windows95 "
  ElseIf myVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
    s$ = "Windows NT "
  End If
  
  lblOS.Caption = s$ & myVer.dwMajorVersion & "." & myVer.dwMinorVersion & " Build " & (myVer.dwBuildNumber And &HFFFF&)
  
  'lblIP = Winsock1.LocalIP
  lblIP = GetIPAddress

End Sub
Sub ChangeWallpaper(BMPfilename As String, UpdateRegistry As Boolean)
    If UpdateRegistry Then
        SystemParametersInfo 20, 0, BMPfilename, 1
    Else
        SystemParametersInfo 20, 0, BMPfilename, 0
    End If
End Sub

Public Function StripZero(sInput As String) As String

Dim nPos As Integer
Dim x As New clsWinAPI

  nPos = InStr(1, sInput, Chr(0))
  
  If nPos <> 0 Then
    StripZero = Left$(sInput, nPos - 1)
  Else
    StripZero = sInput
  End If
  
  Label6 = x.GetSysComputerName
  
End Function

Public Function LPSTRToVBString$(ByVal s$)
  
Dim nullpos&

  nullpos& = InStr(s$, Chr$(0))
  
  If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
  Else
      LPSTRToVBString = ""
  End If
  
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Click()
    Unload Me
    End
End Sub
Public Function CaptureForm(frmSrc As Form) As Picture
   Set CaptureForm = CaptureWindow(frmSrc.hwnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function
Private Sub Form_Load()
Dim lRet As Long
Dim sSaveName As String

  GetInfo
  sSaveName = "c:\SysInfo.bmp"
  Me.Show: DoEvents
  
  'light
  Me.Line (0, 0)-(Me.Width, 0), vbWhite:  Me.Line (0, 0)-(0, Me.Height), vbWhite
  'black
  Me.Line (0, Me.Height - 10)-(Me.Width, Me.Height - 10), vbBlack:  Me.Line (Me.Width - 10, 0)-(Me.Width - 10, Me.Height - 10), vbBlack
  'shadow
  Me.Line (10, Me.Height - 30)-(Me.Width - 10, Me.Height - 30), RGB(90, 90, 90):  Me.Line (Me.Width - 30, 10)-(Me.Width - 30, Me.Height - 20), RGB(90, 90, 90)
  
  DoEvents
  Picture1.Picture = CaptureForm(Me)
  DoEvents
  SavePicture Picture1.Picture, sSaveName
  DoEvents
  SaveSettingString HKEY_CURRENT_USER, "Control Panel\Desktop", "TileWallPaper", "0"
  SaveSettingString HKEY_CURRENT_USER, "Control Panel\Desktop", "WallPaper", sSaveName
  DoEvents
  ChangeWallpaper sSaveName, False
  DoEvents
  RefreshDesktop
  Unload Me
  End
End Sub
Private Sub RefreshDesktop()
Dim lPidl As Long
Dim lRet As Long

 ' Get handle for the desktop
 SHGetSpecialFolderLocation Me.hwnd, CSIDL_DESKTOP, lPidl

 ' Get the system to refresh the desktop
 SHChangeNotify SHCNE_ALLEVENTS, SHCNF_IDLIST, lPidl, 0
 lRet = RedrawWindow(lPidl, ByVal 0&, 0&, RDW_ERASENOW)
End Sub

Public Sub SaveSettingString(hKey As Long, strPath _
As String, strValue As String, strData As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(hKey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
ByVal strData, Len(strData))

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If

lRegResult = RegCloseKey(hCurKey)
End Sub
  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture

  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE

   ' Depending on the value of Client get the proper device context.
   If Client Then
      hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
   Else
      hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire
                                    ' window.
   End If

   ' Create a memory device context for the copy process.
   hDCMemory = CreateCompatibleDC(hDCSrc)
   ' Create a bitmap and place it in the memory DC.
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)

   ' Get screen properties.
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster
                                                      ' capabilities.
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette
                                                        ' support.
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of
                                                        ' palette.

   ' If the screen has a palette make a copy and realize it.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      ' Create a copy of the system palette.
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      ' Select the new palette into the memory DC and realize it.
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If

   ' Copy the on-screen image into the memory DC.
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

' Remove the new copy of the  on-screen image.
   hBmp = SelectObject(hDCMemory, hBmpPrev)

   ' If the screen has a palette get back the palette that was
   ' selected in previously.
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If

   ' Release the device context resources back to the system.
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)

   ' Call CreateBitmapPicture to create a picture object from the
   ' bitmap and palette handles. Then return the resulting picture
   ' object.
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long

   Dim Pic As PicBmp
   ' IPicture requires a reference to "Standard OLE Types."
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID

   ' Fill in with IDispatch Interface ID.
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With

   ' Fill Pic with necessary parts.
   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With

   ' Create Picture object.
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

   ' Return the new Picture object.
   Set CreateBitmapPicture = IPic
End Function



