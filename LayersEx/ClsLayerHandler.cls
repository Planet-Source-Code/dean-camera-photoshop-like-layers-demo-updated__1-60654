VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsLayerHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TempDC As Long
Private TempBitmap As Long
Private TempLayerDC As Long

Public Layers As Collection
Public BGColour As Long

Public Event StartRender()
Public Event FinishRender()

Public Event RenderingLayer(Layer As Integer)
Public Event RenderedLayer(LayerNum As Integer)

Private Sub Class_Initialize()
    Dim ScreenDC As Long

    ScreenDC = GetDC(0) ' Get the DC for the screen
    TempLayerDC = CreateCompatibleDC(ScreenDC) ' Create a tempoary DC for the layer's bitmaps
    TempDC = CreateCompatibleDC(ScreenDC) ' Create a tempoary DC for the rendered layers
    
    TempBitmap = CreateCompatibleBitmap(ScreenDC, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height) ' Creat a bitmap for the rendered layers DC
    SelectObject TempDC, TempBitmap ' Select the created Bitmap into the tempoary DC's context
    
    ReleaseDC 0&, ScreenDC ' Release the control of the screen's DC; prevents a GDI leak
        
    BGColour = vbWhite ' Set the default background colour to white
    
    Set Layers = New Collection ' Create a new layers collection
End Sub

Private Sub Class_Terminate()
    DeleteDC TempDC ' Delete the rendered layers DC
    DeleteDC TempLayerDC ' Delete the layer DC
    
    DeleteObject TempBitmap ' Delete the rendered layers bitmap

    Set Layers = Nothing ' Delete the layers collection
End Sub

Public Function CreateLayer() As ClsLayer
    Dim cLayer As ClsLayer
    
    Set cLayer = New ClsLayer ' Set the tempoary layer to a new layer
    Layers.Add cLayer ' Add the new layer to the layers collection
    cLayer.Index = Layers.Count ' Set the index for the new layer
    Set CreateLayer = cLayer ' Set the function return value to the created layer object
End Function

Public Sub DeleteLayer(Index As Long)
    If Index > Layers.Count Or Index < 0 Then Exit Sub ' Prevents an error if the index number is invalid

    Layers.Remove Index ' Delete the collection entry (which is now nothing)
End Sub

Public Sub RenderSingleLayer(DestDC As Long, Index As Integer, Width As Long, Height As Long, Optional NoOffset As Boolean)
    Dim cLayer As ClsLayer
    Dim gBrush As Long, rRect As RECT

    gBrush = CreateSolidBrush(BGColour) ' Create a new solid brush with the user-selected background colour
    rRect.Bottom = FrmTestLayers.picLayerPreview.Height ' \ Set the rectangle values to the
    rRect.Right = FrmTestLayers.picLayerPreview.Width '   / dimentions of the layer picturebox
    FillRect TempDC, rRect, gBrush ' Fill the DestDC with solid white brush
    DeleteObject gBrush ' Delete the solid background brush
    
    Set cLayer = Layers(Index) ' Set the tempoary layer to the passed layer index
    SelectObject TempLayerDC, cLayer.BitmapHandle ' Select the layer's bitmap into the tempoary layer DC
    
    StretchBlt DestDC, IIf(NoOffset, 0, cLayer.OffsetX), IIf(NoOffset, 0, cLayer.OffsetY), Width, Height, TempLayerDC, 0, 0, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, vbSrcCopy ' Stretch the layer bitmap onto the destination DC
    
    Set cLayer = Nothing
End Sub

Public Sub DrawLayers(DestDC As Long)
    Dim CurrLayer As ClsLayer
    Dim BF As BLENDFUNCTION, lBF As Long
    Dim gBrush As Long, rRect As RECT
    
    RaiseEvent StartRender

    gBrush = CreateSolidBrush(BGColour) ' Create a new solid brush with the user-selected background colour
    rRect.Bottom = FrmTestLayers.picImage.Height ' \ Set the rectangle values to the
    rRect.Right = FrmTestLayers.picImage.Width '   / dimentions of the layer picturebox
    FillRect TempDC, rRect, gBrush ' Fill the DestDC with solid white brush
    DeleteObject gBrush ' Delete the solid background brush
    
    For Each CurrLayer In Layers
        RaiseEvent RenderingLayer(CurrLayer.Index)
    
        If CurrLayer.Visible Then
            SelectObject TempLayerDC, CurrLayer.BitmapHandle ' Select the current layer's bitmap into the tempoary layer DC
    
            With BF
                .BlendOp = AC_SRC_OVER
                .BlendFlags = 0
                .SourceConstantAlpha = CurrLayer.Alpha ' Set the blending Alpha value to the Alpha value of the current layer
                .AlphaFormat = 0
            End With
            RtlMoveMemory lBF, BF, 4 ' Copy the blending structure above to a long
        
            AlphaBlend TempDC, CurrLayer.OffsetX, CurrLayer.OffsetY, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, TempLayerDC, 0, 0, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, lBF ' Blend the layer onto the temp DC
        End If
        
        RaiseEvent RenderedLayer(CurrLayer.Index)
    Next
    
    BitBlt DestDC, 0, 0, FrmTestLayers.picImage.Width, FrmTestLayers.picImage.Height, TempDC, 0, 0, vbSrcCopy ' Copy the rendered layer DC onto the Dest DC
    
    RaiseEvent FinishRender
End Sub
