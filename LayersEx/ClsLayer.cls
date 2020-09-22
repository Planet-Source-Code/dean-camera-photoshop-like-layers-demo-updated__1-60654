VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BitmapHandle As Long
Public Index As Integer

Public Alpha As Integer
Public Visible As Boolean

Public OffsetX As Long
Public OffsetY As Long

Private Sub Class_Initialize()
    Dim ScreenDC As Long
    
    ScreenDC = GetDC(0) ' Get the DC of the screen
    BitmapHandle = CreateCompatibleBitmap(ScreenDC, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height) ' Create a new bitmap compatible with the screen
    ReleaseDC 0&, ScreenDC ' Release control of the Screen's DC; prevents a GDI leak
End Sub

Private Sub Class_Terminate()
    DeleteObject BitmapHandle ' Delete the created tempoary bitmap
End Sub

Sub LoadFile(FileName As String)
    If (Dir(FileName) = "") Then Exit Sub ' File not found, exit sub

    DeleteObject BitmapHandle ' Delete the created bitmap
    
    BitmapHandle = LoadImage(App.hInstance, FileName, IMAGE_BITMAP, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, LR_LOADFROMFILE) ' Load and stretch the file into a new bitmap

    Alpha = 255 ' Set a default Alpha value (no transparency)
End Sub
