VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTestLayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layer Test  Program (Updated) - By Dean Camera"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7080
   Icon            =   "FrmTestLayers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeleteLayer 
      Caption         =   "Delete Layer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame frmLayer 
      Caption         =   "Layer"
      Height          =   2295
      Left            =   4800
      TabIndex        =   6
      Top             =   1200
      Width           =   2175
      Begin VB.ComboBox cmbLayerAlpha 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmbSelLayer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   405
         Width           =   1935
      End
      Begin VB.PictureBox picLayerPreview 
         Height          =   735
         Left            =   650
         ScaleHeight     =   675
         ScaleWidth      =   795
         TabIndex        =   8
         Top             =   765
         Width           =   855
      End
      Begin VB.CheckBox chkVisible 
         Alignment       =   1  'Right Justify
         Caption         =   "Visible:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1965
         Width           =   1935
      End
      Begin VB.Label lblSelectedLayer 
         Caption         =   "Selected Layer:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLayerAlpha 
         Caption         =   "Layer Alpha:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1650
         Width           =   855
      End
   End
   Begin VB.PictureBox picBGColourPreview 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdChooseBG 
      Caption         =   "Set BG Colour"
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CDialogs 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pgbRenderProgress 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3930
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdAddLayer 
      Caption         =   "Add Layer"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.PictureBox picImage 
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblLayerscap 
      Alignment       =   2  'Center
      Caption         =   "Layers: 0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "FrmTestLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                   PHOTOSHOP-LIKE ALPHABLENDED LAYERS DEMO
'                             BY DEAN CAMERA, 2005

' This project will emulate the alpha-blended (transparent) layers, as used in the popular image editing program
' "Photoshop". For simplicity, loaded images are strected to the height and width of the picturebox on this form.
'
' SUPPORTED OS: Windows 2000, 98, ME, XP
'
' When you add a layer, a new ClsLayer blank Bitmap is created. This allows you to immediatly manipulate the layer's
' bitmap, although in this case the OpenFile routine is called to load a pre-created bitmap picture. Once called,
' the OpenFile sub will re-create the layer's bitmap and load the selected image. When a layer image is loaded,
' this demo calls the DrawLayers subroutine of the ClsLayerHandler object to refresh the screen. When this function
' is called, a tempoary DC is cleared (filled with a white brush) and each layer's bitmap is loaded into a created
' LayerDC, where it is alphablended with the main tempoary DC. Once all layers have been rendered, this tempoary DC
' is BitBlted - copied - into yet another DC (called "FinalDC" on this form).
'
' This FinalDC is itself BitBlted onto the picturebox when the picturebox's Paint method is called. The reason
' why two tempoary DC's are used is simple; it reduces flicker when a layer is moved. At the moment, all the layers
' are rendered before they are put into the FinalDC for BitBlting to the screen, thus a redraw only occurs after all
' layers are rendered. If each layer is alphablended onto the FinalDC directly, a Paint command will occur each time
' the user moves their mouse, whether or not the layers have finished rendering (causing flicker).
'
' I think this implementation is clever because only a Bitmap is created for each layer (as opposed to both a Bitmap
' and a DC), saving memory. The current layer that is being rendered has its Bitmap placed into the context of a
' single tempoary DC created in the LayerHandler class. On the downside, the flickerless system i'm using means that
' the speed of the program (when moving a layer) drops with each layer added. With more effort a system could be
' imposed so that all the layers benieth the selected layer (and possibly above the selected layer) are rendered onto
' yet another tempoary DC, so that only a single AlphaBlend call and two BitBlt's are needed each time a layer is
' moved. This is definetly possible but would increase the complexity of this demonstartion somewhat. If enough
' requests are made I shall add this function and resubmit.
'
' Because the equivelent GDIAlphaBlend (which uses the GDI.dll library) function is used instead of the normal
' AlphaBlend function (which uses the msimg32.dll library), only two DLLs are used in this program, User32.dll and
' GDI.dll, both of which are part of the standard Windows OS.
'
' Let me know what you think by emailing me: dean_camera@hotmail.com
'
'  %%%%%%%%%%%                           ------------                           ------
' % Layer(s)  % ==   SelectObject   ==> | Layer Temp | ==    AlphaBlend    ==> | Temp |
' % Bitmap(s) % == (For each layer) ==> |     DC     | == (For each layer) ==> |  DC  |
'  %%%%%%%%%%%                           ------------                           ------
'                                                                                 ||
'  ------------                               -------                             ||
' | Picturebox | <==        BitBlt        == | Final | <==         BitBlt       =='|
' |    hDC     | <== (Upon Paint Request) == |  DC   | <== (When redering done) ==='
'  ------------                               -------
'
'
' UPDATE: Fixed many GDI memory leaks (due to not releasing GetDC(0)), added some new features
'

Option Explicit

Public WithEvents LayerHandler As ClsLayerHandler
Attribute LayerHandler.VB_VarHelpID = -1

Private FinalDC As Long
Private FinalBitmap As Long

Private PrevOffsetX As Long
Private PrevOffsetY As Long

Private Sub cmbLayerAlpha_Click()
    ShowLayerChanges ' Drop-down value clicked, show the changes
End Sub

Private Sub cmbLayerAlpha_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbEnter Then ShowLayerChanges ' Enter key pressed, update the changes
End Sub

Private Sub cmdChooseBG_Click()
    CDialogs.ShowColor ' Show the choose colour dialog
    LayerHandler.BGColour = CDialogs.Color ' Set the background layer colour
    picBGColourPreview.BackColor = CDialogs.Color ' Show the background colour preview

    LayerHandler.DrawLayers FinalDC ' Re-render the layers with the new background colour
    picImage_Paint ' Refresh the picturebox
End Sub

Private Sub cmdDeleteLayer_Click()
    Dim LayerIndex As Long
    
    LayerHandler.DeleteLayer Int(cmbSelLayer.Text) 'Delete the selcted layer
    
    cmbSelLayer.Clear ' Remove all the layer numbers from the select layer combobox
    
    For LayerIndex = 1 To LayerHandler.Layers.Count
        cmbSelLayer.AddItem LayerIndex ' Add each remaining layer number back the the select layer combobox
    Next
    
    picLayerPreview.Cls ' Clear the preview picturebox since the layer has been deleted
    cmbLayerAlpha.Enabled = False ' Disable the select layer alpha combobox
    chkVisible.Enabled = False ' Disable the visible checkbox
    cmdDeleteLayer.Enabled = False ' Disable the Delete Layer button
    
    If LayerHandler.Layers.Count = 0 Then ' No layers left
        cmbSelLayer.Enabled = False ' Disable the select layer combobox
    End If
    
    LayerHandler.DrawLayers FinalDC ' Render all the remaining layers onto the tempoary DC "FinalDC"
    picImage_Paint ' Force a refresh of the picturebox
End Sub

Private Sub Form_Load()
    Dim ScreenDC As Long

    Me.ScaleMode = vbPixels
    
    cmbLayerAlpha.AddItem "100%" ' \
    cmbLayerAlpha.AddItem "75%" '  | Add the percentage sample
    cmbLayerAlpha.AddItem "50%" '  | values to the alpha combobox
    cmbLayerAlpha.AddItem "25%" '  /
    
    Set LayerHandler = New ClsLayerHandler ' Create a new LayerHandler
    
    ScreenDC = GetDC(0) ' Get the DC of the Screen
    
    FinalDC = CreateCompatibleDC(ScreenDC) ' Create a new DC, compatible with the screen
    FinalBitmap = CreateCompatibleBitmap(ScreenDC, picImage.Width, picImage.Height) ' Create a bitmap for the created DC, compatible with the screen
    SelectObject FinalDC, FinalBitmap ' Select the created Bitmap into the created DC's context

    ReleaseDC 0&, ScreenDC ' Release the DC of the Screen; prevents a GDI leak

    LayerHandler.DrawLayers FinalDC ' Render the white background (no layers are present yet) to the temp DC
    picImage_Paint ' Force a picturebox paint to show the white background
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set LayerHandler = Nothing
    
    DeleteDC FinalDC ' Delete the tempoary DC
    DeleteObject FinalBitmap ' Delete the tempoary Bitmap
    
    Set FrmTestLayers = Nothing ' Completly destroy this form
End Sub

Private Sub cmbSelLayer_Click()
    Dim cLayer As ClsLayer
    
    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempoary layer object to the selected layer

    LayerHandler.RenderSingleLayer picLayerPreview.hdc, Int(cmbSelLayer.Text), picLayerPreview.Width / Screen.TwipsPerPixelX, picLayerPreview.Height / Screen.TwipsPerPixelY, True ' Render the selected layer onto the preview picturebox
    cmbLayerAlpha.Text = cLayer.Alpha ' Set the alpha textbox to the selected layer's alpha amount
    chkVisible.Value = IIf(cLayer.Visible, 1, 0) ' Set the Visible checkbox value to that of the layer
    
    cmdDeleteLayer.Enabled = True ' Enable the delete layer button
    
    Set cLayer = Nothing
End Sub

Private Sub cmdAddLayer_Click()
    Dim cLayer As ClsLayer
    
    CDialogs.Filter = "Bitmap Files (*.bmp)|*.bmp"
    CDialogs.ShowOpen ' Show the Open file dialogue
    
    DoEvents ' Redraw the "Add Layer" button to prevent visual glitch
    
    If CDialogs.FileName <> vbNullString Then
        Set cLayer = LayerHandler.CreateLayer ' Create a new layer
        cLayer.LoadFile CDialogs.FileName ' Load the chosen bitmap into the new layer's bitmap
        cLayer.Alpha = 128 ' Set a default Alpha value
        cLayer.Visible = True ' Set the default of Visible to True
    
        LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary "FinalDC"
        picImage_Paint ' Force a refresh of the picturebox - FinalDC is BitBlted onto the picturebox via the Paint method

        lblLayerscap.Caption = "Layers: " & LayerHandler.Layers.Count ' Show the total layers
    
        cmbSelLayer.Enabled = True ' Enable the Select Layer combobox
        chkVisible.Enabled = True ' Enable the Layer Visible checkbox
        
        cmbSelLayer.AddItem LayerHandler.Layers.Count ' Add the new layer to the Select Layer combobox
    End If
    
    CDialogs.FileName = vbNullString
End Sub

Private Sub LayerHandler_RenderedLayer(LayerNum As Integer)
    If LayerNum <= pgbRenderProgress.Max Then ' Only set the progressbar's value if it is less than or equal to the the maximum (this would occur when a layer other than the maximum is deleted)
        pgbRenderProgress.Value = LayerNum
    End If
End Sub

Private Sub LayerHandler_StartRender()
    If LayerHandler.Layers.Count > 0 Then ' Only set the progressbar's maximum if at least one layer present
        pgbRenderProgress.Max = LayerHandler.Layers.Count
    End If
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PrevOffsetX = X ' \ Save the current mouse position
    PrevOffsetY = Y ' / into tempoary variables
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cLayer As ClsLayer
    
    If Button <> 1 Then Exit Sub ' Not left button pressed, exit sub
    If cmbSelLayer.Text = vbNullString Then Exit Sub ' No layer selected, exit sub

    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempoary layer to the selected layer
    
    cLayer.OffsetX = cLayer.OffsetX + ((X - PrevOffsetX) / Screen.TwipsPerPixelX) ' Move the layer by the X amount that the mouse has been moved
    cLayer.OffsetY = cLayer.OffsetY + ((Y - PrevOffsetY) / Screen.TwipsPerPixelY) ' Move the layer by the Y amount that the mouse has been moved
        
    LayerHandler.DrawLayers FinalDC ' Draw the layers
    picImage_Paint ' Refresh the picturebox
    
    PrevOffsetX = X ' \ Save the new current mouse position
    PrevOffsetY = Y ' / into tempoary variables
    
    Set cLayer = Nothing
End Sub

Private Sub picImage_Paint()
    ' Painting of the picturebox is very quick because it just BitBlt's previously rendered data from a
    ' the "FinalDC" DC to the picturebox. This method stops redrawing flicker.

    BitBlt picImage.hdc, 0, 0, picImage.Width, picImage.Height, FinalDC, 0, 0, vbSrcCopy ' Copy the FinalDC to the picturebox
End Sub

Private Sub chkVisible_Click()
    ShowLayerChanges ' Checkbox value changed, update the layer properties and re-render
End Sub

Sub ShowLayerChanges() ' This sub updates the peoperties of the selected layer and re-renders all layers to reflect the changes
    Dim cLayer As ClsLayer
    Dim NewAlphaValue As Integer
    
    If cmbSelLayer.Text = vbNullString Then Exit Sub ' If no layer selected, exit sub
    
    Set cLayer = LayerHandler.Layers(Int(cmbSelLayer.Text)) ' Set the tempary layer to the selected layer

    If cmbLayerAlpha.Text <> vbNullString Then ' If an alpha value specified
        If Right$(cmbLayerAlpha.Text, 1) = "%" Then ' Percentage value specified
            NewAlphaValue = (255 / 100) * Int(Left$(cmbLayerAlpha.Text, Len(cmbLayerAlpha.Text) - 1))  ' Calculate the alpha value from the entered percentage
        Else ' Straight value specified
            NewAlphaValue = Int(cmbLayerAlpha.Text) ' Set the variable to the entered value
        End If
    
        If NewAlphaValue >= 0 And NewAlphaValue <= 255 Then ' Alpha value valid (0-255)
            cLayer.Alpha = Int(NewAlphaValue) ' Set the alpha value of the selected layer to the new alpha value
        End If
    End If
    
    cLayer.Visible = chkVisible.Value ' Set the layer visible attribute
    
    LayerHandler.DrawLayers FinalDC ' Render all the layers onto the tempoary DC "FinalDC"
    picImage_Paint ' Force a refresh of the picturebox
    
    Set cLayer = Nothing
End Sub
