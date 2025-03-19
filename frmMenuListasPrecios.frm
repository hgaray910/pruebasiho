VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuListasPrecios 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de listas de precios"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmMenuListasPrecios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   1620
      Left            =   120
      Top             =   120
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   2858
      BackColor       =   16777215
      ForeColor       =   11682635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackgroundAlignment=   4
      BorderColor     =   8677689
      Caption         =   ""
      CaptionAlignment=   4
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   13284230
      HeaderColorTopRight=   13284230
      HeaderColorBottomLeft=   16777215
      HeaderColorBottomRight=   16777215
      Begin MyCommandButton.MyButton cmdAsignaListaPrecios 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         ToolTipText     =   "Asignación de listas de precios"
         Top             =   1080
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Asignación de listas de precios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoListasPrecios 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         ToolTipText     =   "Mantenimiento de listas de precios"
         Top             =   285
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Listas de precios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdListasPreciosCargo 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         ToolTipText     =   "Listas de precios por cargo"
         Top             =   680
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Listas de precios por cargo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdListaPreciosPemex 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Listas de precios Pemex"
         Top             =   1450
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Listas de precios Pemex"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuListasPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAsignaListaPrecios_Click()
    frmAsignaLista.vllngNumeroOpcion = 308
    frmAsignaLista.HelpContextID = 25
    frmAsignaLista.Show vbModal
End Sub

Private Sub cmdListaPreciosPemex_Click()
    frmListasPreciosPemex.Show vbModal
End Sub

Private Sub cmdListasPreciosCargo_Click()
    frmListaPreciosCargo.Show vbModal
End Sub

Private Sub cmdMantoListasPrecios_Click()
    frmMantoListasPrecios.HelpContextID = 25
    frmMantoListasPrecios.Show vbModal
End Sub

Private Sub Form_Activate()
    fblnHabilitaObjetos Me
    
    Dim rsTemp As New ADODB.Recordset
    
    '*****************caso proceso pemex
    Set rsTemp = frsRegresaRs("SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR " & _
                            "FROM SIPARAMETRO WHERE SIPARAMETRO.VCHNOMBRE = 'BITACTIVAPRECIOSPEMEX' AND SIPARAMETRO.CHRMODULO='PV'")
    lblbActiva = False
    If rsTemp.RecordCount <> 0 Then
        lblbActiva = rsTemp!valor
    End If
    If lblbActiva Then
        cmdListaPreciosPemex.Enabled = True
        cmdListaPreciosPemex.Visible = True
        frmMenuListasPrecios.Height = 2700
        Frame1.Height = 1990
        cmdListaPreciosPemex.Left = 180
        cmdListaPreciosPemex.Top = 1500
    Else
        cmdListaPreciosPemex.Enabled = False
        cmdListaPreciosPemex.Visible = False
    End If
     '*************************************
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

