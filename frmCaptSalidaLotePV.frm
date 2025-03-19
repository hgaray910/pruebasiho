VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmCaptSalidaLotePV 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salida por caducidad"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraLotes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6360
      Left            =   -26
      TabIndex        =   0
      Top             =   -130
      Width           =   12100
      Begin VB.CheckBox chkCodigoBarras 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modo código de barras"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   23
         ToolTipText     =   "Modo código de barras"
         Top             =   15000
         Width           =   1935
      End
      Begin VB.TextBox txtCodigoBarras 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Código de barras"
         Top             =   15000
         Width           =   3465
      End
      Begin VB.OptionButton OptUnidadMinima 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Unidad mínima"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Unidad mínima"
         Top             =   2200
         Width           =   3735
      End
      Begin VB.OptionButton OptUnidadAlterna 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Unidad alterna"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   1320
         TabIndex        =   1
         ToolTipText     =   "Unidad alterna"
         Top             =   1930
         Width           =   3735
      End
      Begin VB.Frame FraCaptura 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1550
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   5680
         Begin VB.TextBox txtCantidadEntUM 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Cantidad actual de entradas para este lote-artículo en unidad mínima"
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox txtCantidadEntUV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Cantidad actual de entradas para este lote-artículo en unidad alterna"
            Top             =   630
            Width           =   930
         End
         Begin VB.TextBox txtLote 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            MaxLength       =   12
            TabIndex        =   3
            ToolTipText     =   "Lote"
            Top             =   230
            Width           =   1245
         End
         Begin VB.TextBox txtFechaCaduce 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Fecha de caducidad"
            Top             =   230
            Width           =   1455
         End
         Begin VB.TextBox txtCantDevol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            TabIndex        =   4
            ToolTipText     =   "Cantidad de artículos a rebajar"
            Top             =   1040
            Width           =   930
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Cantidad en entradas"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   150
            TabIndex        =   18
            Top             =   690
            Width           =   2160
         End
         Begin VB.Label lblLote 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "No. de lote"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   150
            TabIndex        =   16
            Top             =   290
            Width           =   1095
         End
         Begin VB.Label lblCantOrden 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "Cantidad a rebajar"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   250
            Left            =   150
            TabIndex        =   12
            Top             =   1100
            Width           =   1890
         End
      End
      Begin VB.Frame frmGrabar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   2400
         TabIndex        =   10
         Top             =   5460
         Width           =   720
         Begin MyCommandButton.MyButton cmdGrabarRegistro 
            Height          =   600
            Left            =   60
            TabIndex        =   24
            ToolTipText     =   "Confirmar "
            Top             =   130
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   1058
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            MaskColor       =   16777215
            Picture         =   "frmCaptSalidaLotePV.frx":0000
            BackColorOver   =   -2147483633
            BackColorFocus  =   -2147483633
            BackColorDisabled=   -2147483633
            BorderColor     =   -2147483627
            TransparentColor=   16777215
            Caption         =   ""
            DepthEvent      =   1
            PictureDisabled =   "frmCaptSalidaLotePV.frx":0984
            PictureAlignment=   4
            PictureDisabledEffect=   0
            ShowFocus       =   -1  'True
         End
      End
      Begin VB.TextBox txtClaveArticulo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Clave del artículo"
         Top             =   240
         Width           =   4825
      End
      Begin VB.TextBox txtDescripcionLgaArt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Descripción del artículo"
         Top             =   650
         Width           =   4825
      End
      Begin VB.TextBox txtTotalADevolver 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Total de cantidad de salida"
         Top             =   1410
         Width           =   2890
      End
      Begin VSFlex7LCtl.VSFlexGrid VSFArticuloLote 
         Height          =   1500
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   5680
         _cx             =   10019
         _cy             =   2646
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfArticulos 
         Height          =   6000
         Left            =   5920
         TabIndex        =   5
         Top             =   240
         Width           =   5750
         _cx             =   10142
         _cy             =   10583
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Código de barras"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   15000
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Contenido"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   1050
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Artículo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Total de cantidad de salida"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1470
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmCaptSalidaLotePV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjInventario
'| Nombre del Formulario    : frmCaptSalidaLote
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar la rebaja de lotes de un articulo determinado
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : ESanchez
'| Autor                    : ESanchez
'| Fecha de Creación        : 30/Oct/2005
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA

'''''''''''''''''''''''''''''''''''''''''''''''
'Parametros generales
'''''''''''''''''''''''''''''''''''''''''''''''
Public vlstrTablaReferencias As String
Public vlstrChrCveArticulo As String
Public vlintIdArticulo As Long
Public vlstrNoMovimiento As Long
Public vllngNumHojaConsumo As Long
Public vlstrNoMovimientoInicial As Long     'MOVIMIENTO INICIAL
Public vlstrTablaRefInicial As String       'SI SE QUIERE TRAER INFORMACION DE UN TIPO DE MOVIMIENTO EN EXCLUSIVO
Public vlblnSoloMoviento As Boolean         'Bit para saber si muestra todos los lotes del articulo o solo todos los lotes de un articulo/movimiento
Public vlstrTipoMovimiento As String        'UM o UV, cual es que se esta solicitando de salida
Public vlintContenidoArt As Long            'Contenido en caso de UV o 1 en UM
Public vlblnMuestraDevueltos As Boolean     'Muestra los articulos devueltos para vlstrnomovimientoincial
Public vlintNoDepartamento As Integer
Public vlStrTitUM As String                 'Que descripcion de unidad minima
Public vlStrTitUV As String                 'Que descripcion de unidad alterna
Public vlintTipoAccion As Integer           'Se usa unicamente para la captura fisica de inventario: 1 = Entrada UV, 2 = Entrada UM, 3 = Salida UV y 4 = Salida UM

Public vlblnPermiteCambioUnidad As Boolean  'Variable que indicará si se podra cambiar la unidad de medida mientras se seleccionan los lotes
Public vgPantallaAnterior As String

'''''''''''''''''''''''''''''''''''''''''''''''
' Variable locales
'''''''''''''''''''''''''''''''''''''''''''''''
Dim vlrsLotes As New ADODB.Recordset
Dim vlstrsql As String
Dim vlIntCont As Long
Dim vllngTotalDescontados As Integer
Dim vllngTotalLotesEnArticulo As Long
Dim alLotexArt() As varLotes
Dim alArticulos() As varLotes
Dim vlintcont2 As Integer
Dim vlblnEsta As Boolean

Dim vlblnEstatusAgregar As Boolean
Dim vllngLotesAcum As Long   'Cantidad de unidades de lotes acumulados (en Minima)
Dim vllngLotesDisp As Long   'Cantidad de unidades de lotes disponibles para ese articulo (en Minima)
Public vlblnCargada As Boolean
Public vlblnGuardarTraza As Boolean 'Indica si guardara automaticamente, si viene desde el codigo de trazabilidad o desde el grid al darle enter
Public vlblnEncontrado As Boolean 'Indica si encontre el lote de la etiqueta de trazabilidad
Public vgLngTotalLotesxArt As Long

Private Sub chkCodigoBarras_Click()
    If chkCodigoBarras.Value = 0 Then
        txtCodigoBarras.Enabled = False
        txtCodigoBarras.Text = ""
    Else
        txtCodigoBarras.Enabled = True
        'Variable que ayuda para saber si viene desde cargo de requisiciones a paciente
        If vlblnCargada = True Then
            txtCodigoBarras.SetFocus
        End If
    End If
End Sub

Public Sub cmdGrabarRegistro_Click()
On Error GoTo NotificaError
    
    For vlIntCont = 1 To vgLngTotalLotesxMov
        If Trim(vlstrTablaReferencias) = "SAJ" Or Trim(vlstrTablaReferencias) = "EAJ" Then
            If Trim(agLotes(vlIntCont).Articulo) = Trim(txtClaveArticulo.Text) And agLotes(vlIntCont).TipoAccion = vlintTipoAccion Then
                agLotes(vlIntCont).Articulo = ""
                agLotes(vlIntCont).Borrado = "*"
                agLotes(vlIntCont).CantidadUM = 0
                agLotes(vlIntCont).CantidadUV = 0
                agLotes(vlIntCont).CantidadUMInicio = 0
                agLotes(vlIntCont).CantidadUVInicio = 0
                agLotes(vlIntCont).lote = ""
                agLotes(vlIntCont).Movimiento = 0
                agLotes(vlIntCont).Devolucion = 0
            End If
        Else
            If Trim(agLotes(vlIntCont).Articulo) = Trim(txtClaveArticulo.Text) Then
                agLotes(vlIntCont).Articulo = ""
                agLotes(vlIntCont).Borrado = "*"
                agLotes(vlIntCont).CantidadUM = 0
                agLotes(vlIntCont).CantidadUV = 0
                agLotes(vlIntCont).CantidadUMInicio = 0
                agLotes(vlIntCont).CantidadUVInicio = 0
                agLotes(vlIntCont).lote = ""
                agLotes(vlIntCont).Movimiento = 0
                agLotes(vlIntCont).Devolucion = 0
            End If
        End If
    Next vlIntCont
    
    For vlIntCont = 1 To vgLngTotalLotesxArt
        vgLngTotalLotesxMov = vgLngTotalLotesxMov + 1
        ReDim Preserve agLotes(vgLngTotalLotesxMov)
        agLotes(vgLngTotalLotesxMov).Articulo = alLotexArt(vlIntCont).Articulo
        agLotes(vgLngTotalLotesxMov).Borrado = alLotexArt(vlIntCont).Borrado
        
        agLotes(vgLngTotalLotesxMov).CantidadUV = alLotexArt(vlIntCont).CantidadUV
        vllngCantidadUV = vllngCantidadUV + alLotexArt(vlIntCont).CantidadUV ' caso 20417
        
        agLotes(vgLngTotalLotesxMov).CantidadUM = alLotexArt(vlIntCont).CantidadUM
        vllngCantidadUM = vllngCantidadUM + alLotexArt(vlIntCont).CantidadUM ' caso 20417
        
        agLotes(vgLngTotalLotesxMov).CantidadUMInicio = alLotexArt(vlIntCont).CantidadUMInicio
        agLotes(vgLngTotalLotesxMov).CantidadUVInicio = alLotexArt(vlIntCont).CantidadUVInicio
        agLotes(vgLngTotalLotesxMov).fechaCaducidad = alLotexArt(vlIntCont).fechaCaducidad
        agLotes(vgLngTotalLotesxMov).lote = alLotexArt(vlIntCont).lote
        agLotes(vgLngTotalLotesxMov).Movimiento = alLotexArt(vlIntCont).Movimiento
        agLotes(vgLngTotalLotesxMov).Devolucion = alLotexArt(vlIntCont).Devolucion
        agLotes(vgLngTotalLotesxMov).TipoAccion = alLotexArt(vlIntCont).TipoAccion
    Next vlIntCont
    
    vgblnCapturoLoteYCaduc = True
    Unload Me
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
End Sub

Private Sub cmdGrabarRegistro2_Click()

End Sub

Private Sub Form_Activate()
Dim vldblDescontados As Double
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlblnBitTrazabilidad As Boolean
    '--------------Trazabilidad--------------------
    Set rs = frsRegresaRs("select NVL(VCHVALOR,' ') as VALOR FROM SIPARAMETRO where VCHNOMBRE = 'BITTRAZABILIDAD' AND INTCVEEMPRESACONTABLE = " & CStr(vgintClaveEmpresaContable))
    If rs.RecordCount <> 0 Then vlblnBitTrazabilidad = Trim(rs!Valor)
    rs.Close
    If vlblnBitTrazabilidad Then
        frmCaptSalidaLotePV.Height = 7015
        FraLotes.Height = 6760
        frmGrabar.Top = 5860
        VSFArticuloLote.Top = 4360
        FraCaptura.Top = 2800
        vsfArticulos.Top = 640
        OptUnidadMinima.Top = 2600
        OptUnidadAlterna.Top = 2330
        Label4.Top = 2320
        Label2.Top = 1870
        txtTotalADevolver.Top = 1810
        txtDescripcionLgaArt.Top = 1050
        txtClaveArticulo.Top = 640
        Label1.Top = 700
        Label5.Top = 300
        Label5.Left = 120
        txtCodigoBarras.Top = 240
        txtCodigoBarras.Left = 1940
        txtCodigoBarras.Width = 3860
        chkCodigoBarras.Top = 240
        chkCodigoBarras.Left = 5930
        chkCodigoBarras.Enabled = True
        
    End If
    vlblnEncontrado = False
    vlblnCargada = True
    txtClaveArticulo = vlstrChrCveArticulo
    OptUnidadAlterna.Caption = vlStrTitUV
    OptUnidadMinima.Caption = vlStrTitUM
    OptUnidadMinima.Value = IIf(vlstrTipoMovimiento = "UM", True, False)
    OptUnidadAlterna.Value = IIf(vlstrTipoMovimiento = "UV", True, False)
    OptUnidadMinima.Enabled = IIf(vlblnPermiteCambioUnidad, True, False)
    OptUnidadAlterna.Enabled = IIf(vlblnPermiteCambioUnidad, True, False)
    OptUnidadMinima.Visible = IIf(vlintContenidoArt = 1, False, True)

    pLlenaArregloArticulos              'Busca todos los articulos de entradas
    
    If vllngTotalLotesEnArticulo > 0 Then

        'Llena grid con todo lo que se habia grabado ya de este articulo
        'If vgPantallaAnterior <> "frmKardexPaciente" Then
           pLlenaArregloLocal
        'End If
        vllngTotalDescontados = 0
        pVaciaArregloGrid
        
        If vlblnPermiteCambioUnidad Then
            If vglngTotalCantUMLotexArt Mod vlintContenidoArt Then
                OptUnidadMinima.Value = True
                OptUnidadAlterna.Enabled = False
            End If
        End If
        
        If vlstrTipoMovimiento = "UV" Then
            vldblDescontados = vglngTotalCantUVLotexArt + (vglngTotalCantUMLotexArt / vlintContenidoArt)
        Else
            vldblDescontados = vglngTotalCantUMLotexArt + (vglngTotalCantUVLotexArt * vlintContenidoArt)
        End If
        
        vllngTotalDescontados = vldblDescontados

        If vllngTotalDescontados <= Val(txtTotalADevolver) Then
            txtCantDevol.Text = Val(txtTotalADevolver) - vllngTotalDescontados
        Else
            If vllngTotalDescontados > Val(txtTotalADevolver) Then
                txtCantDevol.Text = Val(txtTotalADevolver) - vllngTotalDescontados
            End If
        End If
        
        pCalculaLotesDisponibles
        
        If Val(txtCantDevol.Text) = 0 Then
            FraCaptura.Enabled = False
            cmdGrabarRegistro.Enabled = True
            cmdGrabarRegistro.SetFocus
        Else
            If vgblnForzarLoteYCaduc Then
                pCalculaLotesAcumulados
                If vllngLotesAcum = vllngLotesDisp Then
                    FraCaptura.Enabled = False
                    cmdGrabarRegistro.Enabled = True
                    cmdGrabarRegistro.SetFocus
                Else
                    FraCaptura.Enabled = True
                    pEnfocaTextBox txtLote
                End If
            Else
                FraCaptura.Enabled = True
                pEnfocaTextBox txtLote
            End If
        End If
    Else
        Unload Me
    End If
    'Esta parte sirve para cuando es un cargo a requisicion a paciente y usaron la funcion del codigo de barras desde la pantalla de surtir
'    If vlblnBitTrazabilidad Then
'        'Si viene con un codigo de barras al activar la pantalla quiere decir que es funcion de cargo a requisicion
'        If txtCodigoBarras.Text <> "" Then
'            txtCodigoBarras_KeyDown 13, 0
'        End If
'        If vlblnGuardarTraza Then
'            Call cmdGrabarRegistro_Click
'        End If
'    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError             'Manejo del error
    
    If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
    If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
    With VSFArticuloLote
        .Clear
        .Redraw = False
        .Visible = False
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .FixedAlignment(1) = flexAlignCenterCenter
        .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad"
        .ColWidth(0) = 0        'Fix
        .ColWidth(1) = 1200      'Cantidaduv
        .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1200) 'Cantidadum
        .ColWidth(3) = 1200     'Lote
        .ColAlignment(3) = flexAlignLeftBottom
        .ColWidth(4) = 1350     'Fecha de caducidad
        .Redraw = True
        .Visible = True
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If Val(txtCantDevol.Text) <> 0 And txtCantDevol.Text <> "" Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg("17"), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                vgblnCapturoLoteYCaduc = False
                Unload Me
            Else
                pEnfocaTextBox txtLote
            End If
        Else
            vgblnCapturoLoteYCaduc = False
            Unload Me
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    cmdGrabarRegistro.Enabled = False
    vlblnCargada = True
End Sub

Private Sub OptUnidadAlterna_Click()
    If vllngTotalLotesEnArticulo > 0 And vlstrTipoMovimiento = "UM" Then
    
        vlstrTipoMovimiento = "UV"
        txtTotalADevolver.Text = Trim(str(Val(txtTotalADevolver.Text) / vlintContenidoArt))
        txtCantDevol.Text = IIf(Trim(txtCantDevol.Text) <> "", Trim(str(Val(txtCantDevol.Text) / vlintContenidoArt)), txtCantDevol.Text)
        
        vgPantallaAnterior = "NA"

        If vllngTotalDescontados <> 0 Then
            vllngTotalDescontados = vllngTotalDescontados / vlintContenidoArt
        End If
        
        pCalculaLotesDisponibles
        
        If Val(txtCantDevol.Text) = 0 And Trim(txtCantDevol.Text) <> "" Then
'            FraCaptura.Enabled = False
'            cmdGrabarRegistro.Enabled = True
'            cmdGrabarRegistro.SetFocus
        Else
            If vgblnForzarLoteYCaduc Then
                pCalculaLotesAcumulados
                If vllngLotesAcum = vllngLotesDisp Then
                    FraCaptura.Enabled = False
                    cmdGrabarRegistro.Enabled = True
                    cmdGrabarRegistro.SetFocus
                Else
                    FraCaptura.Enabled = True
                    pEnfocaTextBox txtLote
                End If
            Else
                FraCaptura.Enabled = True
                pEnfocaTextBox txtLote
            End If
        End If
    End If
End Sub

Private Sub OptUnidadMinima_Click()
    If vllngTotalLotesEnArticulo > 0 And vlstrTipoMovimiento = "UV" Then
    
        vlstrTipoMovimiento = "UM"
        txtTotalADevolver.Text = Trim(str(Val(txtTotalADevolver.Text) * vlintContenidoArt))
        txtCantDevol.Text = IIf(Trim(txtCantDevol.Text) <> "", Trim(str(Val(txtCantDevol.Text) * vlintContenidoArt)), txtCantDevol.Text)
        
        If vllngTotalDescontados <> 0 Then
            vllngTotalDescontados = vllngTotalDescontados * vlintContenidoArt
        End If
        
        pCalculaLotesDisponibles
        
        If Val(txtCantDevol.Text) = 0 And Trim(txtCantDevol.Text) <> "" Then
'            FraCaptura.Enabled = False
'            cmdGrabarRegistro.Enabled = True
'            cmdGrabarRegistro.SetFocus
        Else
            If vgblnForzarLoteYCaduc Then
                pCalculaLotesAcumulados
                If vllngLotesAcum = vllngLotesDisp Then
                    FraCaptura.Enabled = False
                    cmdGrabarRegistro.Enabled = True
                    cmdGrabarRegistro.SetFocus
                Else
                    FraCaptura.Enabled = True
                    pEnfocaTextBox txtLote
                End If
            Else
                FraCaptura.Enabled = True
                pEnfocaTextBox txtLote
            End If
        End If
    End If
End Sub

Private Sub txtCantDevol_GotFocus()
    pSelTextBox txtCantDevol
End Sub

Private Sub txtCantDevol_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim vlblnRebaja As Boolean

    If KeyAscii = 13 Then
        If Trim(txtLote.Text) <> "" Then
            If Val(txtCantDevol.Text) <= 0 Then
                txtCantDevol.SetFocus
            Else
                vlblnRebaja = fblnrebaja(Val(txtCantidadEntUV), Val(txtCantidadEntUM), Val(txtCantDevol.Text))
                If vlblnRebaja Then
                    pResta
                    If Not vgblnForzarLoteYCaduc Then
                        
                        'cmdGrabarRegistro.Enabled = True ' se comenta en el caso 20417
                        
                        If Val(txtCantDevol.Text) = 0 Then
                            FraCaptura.Enabled = False
                            cmdGrabarRegistro.SetFocus
                        Else
                            FraCaptura.Enabled = True
                            pEnfocaTextBox txtLote
                        End If
                    End If
                Else
                    pEnfocaTextBox txtCantDevol
                    Exit Sub
                End If
            End If
        Else
            pEnfocaTextBox txtLote
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantDevol_KeyPress"))
End Sub

Public Sub txtCodigoBarras_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    Dim rsEtiqueta As ADODB.Recordset
    Dim vlstrSentencia As String
    Dim intcontador As Integer
    
    If txtCodigoBarras.Text = "" Then
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyReturn
            'Siempre cantidad en 1
            txtCantDevol.Text = 1
            'Ahora traemos los datos de la etiqueta
            vlstrSentencia = "select * from ivetiqueta where intidetiqueta = " & txtCodigoBarras.Text & " and chrcvearticulo = " & txtClaveArticulo.Text
            
            Set rsEtiqueta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rsEtiqueta.RecordCount > 0 Then
                With rsEtiqueta
                    For intcontador = 1 To vsfArticulos.Rows - 1
                        If Trim(vsfArticulos.TextMatrix(intcontador, 3)) = Trim(!chrlote) Then
                            vsfArticulos.Row = intcontador
                            vsfArticulos_DblClick
                            vlblnEncontrado = True
                        End If
                    Next intcontador
                    'Ahora agregamos los datos
                    txtCantDevol_KeyPress vbKeyReturn
                    'Variable que ayuda para saber si viene desde cargo de requisiciones a paciente
                    If vlblnCargada = True Then
                        txtCodigoBarras.SetFocus
                    End If
                End With
                rsEtiqueta.Close
            Else
                MsgBox "El código de barras proporcionado no coincide." & vbCrLf & "Por favor, inténtelo nuevamente.", vbCritical, "Mensaje"
            End If
    End Select
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCodigoBarras_KeyDown"))
End Sub

Private Sub txtCodigoBarras_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtLote_GotFocus()
    pSelTextBox txtLote
End Sub

Private Sub txtLote_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------------
' Procedimiento para validar que exista el lote  y la cantidad no exceda
'--------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintResul As Integer
    Dim vlstrMensaje As String

    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtLote.Text) <> "" Then
                '''''''''''''''''''''''''''''''''''''''''''
                'busca si esta en el grid para traerse los datos
                vlblnEsta = False
                For vlintcont2 = 1 To vllngTotalLotesEnArticulo
                    If Trim(txtLote.Text) = Trim(alArticulos(vlintcont2).lote) Then
                        vlblnEsta = True
                        txtCantidadEntUV = alArticulos(vlintcont2).CantidadUV
                        txtCantidadEntUM = alArticulos(vlintcont2).CantidadUM
                        txtFechaCaduce = Trim(Format(alArticulos(vlintcont2).fechaCaducidad, "DD/MMM/YYYY"))

                        'txtCantDevol.Text = Str(vlCantR)
                        pEnfocaTextBox txtCantDevol
                        Exit Sub
                    End If
                Next vlintcont2
                If vlblnEsta = False Then
                    MsgBox "Lote " & Trim(txtLote.Text) & " no existe para este artículo", vbCritical, "Mensaje"
                    pEnfocaTextBox txtLote
                    Exit Sub
                End If
            Else
                pEnfocaTextBox txtLote
            End If
        Case vbKeyEscape
            vgblnCapturoLoteYCaduc = False
            Unload Me
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtLote_KeyPress"))

End Sub

Private Function fblnValidaLote() As Boolean
On Error GoTo NotificaError
      
    fblnValidaLote = False
    'Valida si ya ha sido seleccionado
    For vlIntCont = 1 To vgLngTotalLotesxArt
        If Trim(txtLote.Text) = Trim(alLotexArt(vlIntCont).lote) Then
            MsgBox "Lote " & Trim(txtLote.Text) & " ya seleccionado", vbCritical, "Mensaje"
            'Exit Function
        End If
    Next vlIntCont
                    
    'Valida si este articulo tiene ya capturado algun lote
    If vllngTotalLotesEnArticulo > 0 Then
        For vlIntCont = 1 To vllngTotalLotesEnArticulo
            If Trim(txtLote.Text) = Trim(alArticulos(vlIntCont).lote) Then
                fblnValidaLote = True
                txtCantidadEntUM = alArticulos(vlIntCont).CantidadUM
                txtCantidadEntUV = alArticulos(vlIntCont).CantidadUV
                txtFechaCaduce.Text = str(alArticulos(vlIntCont).fechaCaducidad)
                Exit For
            End If
        Next vlIntCont
    End If

    If Not (fblnValidaLote) Then MsgBox "Lote " & Trim(txtLote.Text) & " no existe para este artículo", vbCritical, "Mensaje"

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaLote"))
End Function

Private Sub pResta()
On Error GoTo NotificaError
Dim vlblnSeleccionada As Boolean
Dim vlblnExisteEnArticulos As Boolean

    If (vllngTotalDescontados + Val(txtCantDevol.Text)) <= Val(txtTotalADevolver.Text) Then
        vlblnSeleccionada = False
        vlblnExisteEnArticulos = False
        
        If vllngTotalLotesEnArticulo > 0 Then
            For vlIntCont = 1 To vllngTotalLotesEnArticulo
                If Trim(txtLote.Text) = Trim(alArticulos(vlIntCont).lote) Then
                    vlblnExisteEnArticulos = True
                    ' Ver si existe solo aumenta la cantidad
                    For vlintcont2 = 1 To vgLngTotalLotesxArt
                        If Trim(txtLote.Text) = Trim(alLotexArt(vlintcont2).lote) And Trim(alLotexArt(vlintcont2).Borrado) <> "*" Then
                            If vlstrTipoMovimiento = "UV" Then
                                alLotexArt(vlintcont2).CantidadUV = alLotexArt(vlintcont2).CantidadUV + Val(txtCantDevol.Text)
                            Else
                                alLotexArt(vlintcont2).CantidadUM = alLotexArt(vlintcont2).CantidadUM + Val(txtCantDevol.Text)
                            End If
                            vlblnSeleccionada = True
                            Exit For
                        End If
                    Next vlintcont2
                    'si no existe lo da de alta en el arreglo
                    If vlblnSeleccionada = False Then
                        vgLngTotalLotesxArt = vgLngTotalLotesxArt + 1
                        ReDim Preserve alLotexArt(vgLngTotalLotesxArt)
                        alLotexArt(vgLngTotalLotesxArt).Articulo = vlstrChrCveArticulo
                        alLotexArt(vgLngTotalLotesxArt).Borrado = ""
                        If vlstrTipoMovimiento = "UV" Then
                            alLotexArt(vgLngTotalLotesxArt).CantidadUV = txtCantDevol.Text
                            alLotexArt(vgLngTotalLotesxArt).CantidadUM = 0
                        Else
                            alLotexArt(vgLngTotalLotesxArt).CantidadUM = txtCantDevol.Text
                            alLotexArt(vgLngTotalLotesxArt).CantidadUV = 0
                        End If
                        alLotexArt(vgLngTotalLotesxArt).fechaCaducidad = txtFechaCaduce.Text
                        alLotexArt(vgLngTotalLotesxArt).lote = txtLote.Text
                        alLotexArt(vgLngTotalLotesxArt).TipoAccion = vlintTipoAccion
                    End If
                    pVaciaArregloGrid
                    vllngTotalDescontados = vllngTotalDescontados + Val(txtCantDevol.Text)
                    txtCantDevol.Text = Val(txtTotalADevolver.Text) - vllngTotalDescontados
                    txtLote.Text = ""
                    txtFechaCaduce.Text = ""
                    If Val(txtCantDevol.Text) = 0 Then
                        FraCaptura.Enabled = False
                        cmdGrabarRegistro.Enabled = True
                        cmdGrabarRegistro.SetFocus
                    Else
                        If vgblnForzarLoteYCaduc Then
                            pCalculaLotesAcumulados
                            If vllngLotesAcum = vllngLotesDisp Then
                                FraCaptura.Enabled = False
                                cmdGrabarRegistro.Enabled = True
                                cmdGrabarRegistro.SetFocus
                            Else
                                pEnfocaTextBox txtLote
                            End If
                        Else
                            pEnfocaTextBox txtLote
                        End If
                    End If
                    Exit For
                End If
            Next vlIntCont
        End If
        
        If Not (vlblnExisteEnArticulos) Then
            MsgBox "Lote " & Trim(txtLote.Text) & " no existe para este artículo", vbCritical, "Mensaje"
            pEnfocaTextBox txtLote
            Exit Sub
        End If
        
        If vlblnPermiteCambioUnidad Then
            If OptUnidadAlterna.Value Then
                OptUnidadAlterna.Enabled = True
                OptUnidadMinima.Enabled = True
            Else
                If Val(txtCantDevol.Text) Mod vlintContenidoArt > 0 Then
                    OptUnidadAlterna.Enabled = False
                Else
                    OptUnidadAlterna.Enabled = True
                End If
            End If
        End If
    Else
        MsgBox SIHOMsg(36) & " menor igual a " & str(Val(txtTotalADevolver.Text) - vllngTotalDescontados), vbCritical, "Mensaje"
        pEnfocaTextBox txtCantDevol
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pResta"))
End Sub

Private Sub txtLote_LostFocus()
    If Trim(txtLote.Text) <> "" Then
        '''''''''''''''''''''''''''''''''''''''''''
        'busca si esta en el grid para traerse los datos
        vlblnEsta = False
        For vlintcont2 = 1 To vllngTotalLotesEnArticulo
            If Trim(txtLote.Text) = Trim(alArticulos(vlintcont2).lote) Then
                vlblnEsta = True
                txtCantidadEntUV = alArticulos(vlintcont2).CantidadUV
                txtCantidadEntUM = alArticulos(vlintcont2).CantidadUM
                txtFechaCaduce = Trim(Format(alArticulos(vlintcont2).fechaCaducidad, "DD/MMM/YYYY"))
                Exit Sub
            End If
        Next vlintcont2
        If vlblnEsta = False Then
            MsgBox "Lote " & Trim(txtLote.Text) & " no existe para este artículo", vbCritical, "Mensaje"
            pEnfocaTextBox txtLote
            Exit Sub
        End If
    End If
End Sub

Private Sub VSFArticuloLote_DblClick()
On Error GoTo NotificaError
    Dim vlstrStyle, vlstrResponse, MyString
    Dim vlblnAvanzara As Boolean
    
    If VSFArticuloLote.Row > -1 Then
        If VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3) <> "" Then
            If vgLngTotalLotesxArt > 0 Then
                For vlIntCont = 1 To vgLngTotalLotesxArt
                    If Trim(VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3)) = Trim(alLotexArt(vlIntCont).lote) And Trim(alLotexArt(vlIntCont).Borrado) <> "*" Then
                        vlblnAvanzara = True
                        
                        If vlstrTipoMovimiento = "UV" Then
                            If alLotexArt(vlIntCont).CantidadUV > 0 Then
                                vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                                vlstrResponse = MsgBox("Seguro de eliminar lote " & Trim(VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3)) & " ?", vlstrStyle, "Mensaje")
                                vlblnAvanzara = IIf(vlstrResponse = vbYes, True, False)
                            Else
                                ' El lote no tiene seleccionadas unidades alternas que eliminar.
                                MsgBox SIHOMsg(1164), vbInformation, "Mensaje"
                                If OptUnidadAlterna.Enabled Then
                                    OptUnidadAlterna.SetFocus
                                Else
                                    VSFArticuloLote.SetFocus
                                End If
                                vlblnAvanzara = False
                            End If
                        Else
                            If alLotexArt(vlIntCont).CantidadUM > 0 Then
                                vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                                vlstrResponse = MsgBox("Seguro de eliminar lote " & Trim(VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3)) & " ?", vlstrStyle, "Mensaje")
                                vlblnAvanzara = IIf(vlstrResponse = vbYes, True, False)
                            Else
                                ' El lote no tiene seleccionadas unidades mínimas que eliminar.
                                MsgBox SIHOMsg(1165), vbInformation, "Mensaje"
                                If OptUnidadMinima.Enabled Then
                                    OptUnidadMinima.SetFocus
                                Else
                                    VSFArticuloLote.SetFocus
                                End If
                                vlblnAvanzara = False
                            End If
                        End If
                            
                        If vlblnAvanzara Then
                            vllngTotalDescontados = vllngTotalDescontados - IIf(vlstrTipoMovimiento = "UV", alLotexArt(vlIntCont).CantidadUV, alLotexArt(vlIntCont).CantidadUM)
                            txtCantDevol.Text = Val(txtTotalADevolver.Text) - vllngTotalDescontados
                            
                            alLotexArt(vlIntCont).CantidadUM = IIf(vlstrTipoMovimiento = "UM", 0, alLotexArt(vlIntCont).CantidadUM)
                            alLotexArt(vlIntCont).CantidadUV = IIf(vlstrTipoMovimiento = "UV", 0, alLotexArt(vlIntCont).CantidadUV)
                            If (Val(alLotexArt(vlIntCont).CantidadUV) + Val(alLotexArt(vlIntCont).CantidadUM)) = 0 Then
                                alLotexArt(vlIntCont).Borrado = "*"
                            End If
                            
                            If Not vgblnForzarLoteYCaduc Then
                                If Val(txtCantDevol.Text) = 0 Then
                                    FraCaptura.Enabled = False
                                    cmdGrabarRegistro.Enabled = True
                                    cmdGrabarRegistro.SetFocus
                                Else
                                    If vgblnForzarLoteYCaduc Then
                                        pVaciaArregloGrid
                                        pCalculaLotesAcumulados
                                        If vllngLotesAcum = vllngLotesDisp Then
                                            FraCaptura.Enabled = False
                                            cmdGrabarRegistro.Enabled = True
                                            cmdGrabarRegistro.SetFocus
                                        Else
                                            FraCaptura.Enabled = True
                                            pEnfocaTextBox txtLote
                                        End If
                                    Else
                                        FraCaptura.Enabled = True
                                        pEnfocaTextBox txtLote
                                    End If
                                End If
                            Else
                                If Val(txtCantDevol.Text) = 0 Then
                                    FraCaptura.Enabled = False
                                    cmdGrabarRegistro.Enabled = True
                                    cmdGrabarRegistro.SetFocus
                                Else
                                    If vgblnForzarLoteYCaduc Then
                                        pVaciaArregloGrid
                                        pCalculaLotesAcumulados
                                        If vllngLotesAcum = vllngLotesDisp Then
                                            FraCaptura.Enabled = False
                                            cmdGrabarRegistro.Enabled = True
                                            cmdGrabarRegistro.SetFocus
                                        Else
                                            FraCaptura.Enabled = True
                                            cmdGrabarRegistro.Enabled = False
                                            pEnfocaTextBox txtLote
                                        End If
                                    Else
                                        FraCaptura.Enabled = True
                                        cmdGrabarRegistro.Enabled = False
                                        pEnfocaTextBox txtLote
                                    End If
                                End If
                            End If
                            Exit For
                        Else
                            Exit For
                        End If
                        pEnfocaTextBox txtLote
                    End If
                Next vlIntCont
            End If
            
            pVaciaArregloGrid
            
            If vlblnPermiteCambioUnidad Then
                If OptUnidadAlterna.Value Then
                    OptUnidadAlterna.Enabled = True
                    OptUnidadMinima.Enabled = True
                Else
                    OptUnidadAlterna.Enabled = IIf(Val(txtCantDevol.Text) Mod vlintContenidoArt > 0, False, True)
                End If
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":VSFArticuloLote_DblClick"))
End Sub

Private Sub pVaciaArregloGrid(Optional vlstrIniciaCuenta As String)
On Error GoTo NotificaError
Dim vlintTotalSeries As Long

    With VSFArticuloLote

        .Clear
        .Cols = 5
        .Rows = 1
        
        For vlIntCont = 1 To vgLngTotalLotesxArt
            If Trim(alLotexArt(vlIntCont).lote) <> "" And Trim(alLotexArt(vlIntCont).Borrado) <> "*" Then
                vlintTotalSeries = vlintTotalSeries + 1
                .Rows = .Rows + 1

                .TextMatrix(vlintTotalSeries, 1) = Trim(str(alLotexArt(vlIntCont).CantidadUV))
                .TextMatrix(vlintTotalSeries, 2) = Trim(str(alLotexArt(vlIntCont).CantidadUM))

                .TextMatrix(vlintTotalSeries, 3) = Trim(alLotexArt(vlIntCont).lote)
                .TextMatrix(vlintTotalSeries, 4) = Trim(Format(alLotexArt(vlIntCont).fechaCaducidad, "DD/MMM/YYYY"))
            End If
        Next vlIntCont
        
        If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
        If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
        .FixedCols = 1
        .FixedRows = 1
        .FixedAlignment(1) = flexAlignCenterCenter
        .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad"
        .ColWidth(0) = 0        'Fix
        .ColWidth(1) = 1200      'Cantidaduv
        .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1200) 'Cantidadum
        .ColWidth(3) = 1200     'Lote
        .ColWidth(4) = 2050     'Fecha de caducidad
        
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftBottom
        'Alineacion de la fecha de caducidad porque no se dejaba
        For vlIntCont = 1 To VSFArticuloLote.Rows - 1
            VSFArticuloLote.Row = vlIntCont
            VSFArticuloLote.Col = 4
            VSFArticuloLote.CellAlignment = flexAlignRightBottom
        Next vlIntCont
        End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":VSFArticuloLote_DblClick"))
End Sub
Private Sub pLlenaArregloLocal()
On Error GoTo NotificaError

    If (Not alLotexArt) = -1 Then
       ReDim Preserve alLotexArt(1)
    End If

    vglngTotalCantLotexArt = 0
    vglngTotalCantUVLotexArt = 0
    vglngTotalCantUMLotexArt = 0
    
    vgLngTotalLotesxArt = 0
    
    For vlIntCont = 1 To vgLngTotalLotesxMov
        If Trim(agLotes(vlIntCont).Articulo) = Trim(txtClaveArticulo.Text) And agLotes(vlIntCont).Borrado <> "*" Then
            If Trim(vlstrTablaReferencias) = "SAJ" Or Trim(vlstrTablaReferencias) = "EAJ" Then 'Si es captura fisica
                If agLotes(vlIntCont).TipoAccion = vlintTipoAccion Then
                    vgLngTotalLotesxArt = vgLngTotalLotesxArt + 1
                    vglngTotalCantUVLotexArt = vglngTotalCantUVLotexArt + agLotes(vlIntCont).CantidadUV
                    vglngTotalCantUMLotexArt = vglngTotalCantUMLotexArt + agLotes(vlIntCont).CantidadUM
                    ReDim Preserve alLotexArt(vgLngTotalLotesxArt)
                    alLotexArt(vgLngTotalLotesxArt).Articulo = agLotes(vlIntCont).Articulo
                    alLotexArt(vgLngTotalLotesxArt).Borrado = ""
                    alLotexArt(vgLngTotalLotesxArt).CantidadUV = agLotes(vlIntCont).CantidadUV
                    alLotexArt(vgLngTotalLotesxArt).CantidadUM = agLotes(vlIntCont).CantidadUM
                    alLotexArt(vgLngTotalLotesxArt).CantidadUVInicio = agLotes(vlIntCont).CantidadUVInicio
                    alLotexArt(vgLngTotalLotesxArt).CantidadUMInicio = agLotes(vlIntCont).CantidadUMInicio
                    alLotexArt(vgLngTotalLotesxArt).fechaCaducidad = agLotes(vlIntCont).fechaCaducidad
                    alLotexArt(vgLngTotalLotesxArt).lote = agLotes(vlIntCont).lote
                    alLotexArt(vgLngTotalLotesxArt).TipoAccion = agLotes(vlIntCont).TipoAccion
                End If
            Else
                If Trim(vlstrTablaReferencias) <> "SOS" And Trim(vlstrTablaReferencias) <> "EOE" Then
                    vgLngTotalLotesxArt = vgLngTotalLotesxArt + 1
                    vglngTotalCantUVLotexArt = vglngTotalCantUVLotexArt + agLotes(vlIntCont).CantidadUV
                    vglngTotalCantUMLotexArt = vglngTotalCantUMLotexArt + agLotes(vlIntCont).CantidadUM
                    ReDim Preserve alLotexArt(vgLngTotalLotesxArt)
                    alLotexArt(vgLngTotalLotesxArt).Articulo = agLotes(vlIntCont).Articulo
                    alLotexArt(vgLngTotalLotesxArt).Borrado = ""
                    alLotexArt(vgLngTotalLotesxArt).CantidadUV = agLotes(vlIntCont).CantidadUV
                    alLotexArt(vgLngTotalLotesxArt).CantidadUM = agLotes(vlIntCont).CantidadUM
                    alLotexArt(vgLngTotalLotesxArt).CantidadUVInicio = agLotes(vlIntCont).CantidadUVInicio
                    alLotexArt(vgLngTotalLotesxArt).CantidadUMInicio = agLotes(vlIntCont).CantidadUMInicio
                    alLotexArt(vgLngTotalLotesxArt).fechaCaducidad = agLotes(vlIntCont).fechaCaducidad
                    alLotexArt(vgLngTotalLotesxArt).lote = agLotes(vlIntCont).lote
                    'alLotexArt(vgLngTotalLotesxArt).TipoAccion = agLotes(vlintCont).TipoAccion
                Else
                    If (vlstrTipoMovimiento = "UV" And agLotes(vlIntCont).CantidadUV <> 0) Or (vlstrTipoMovimiento = "UM" And agLotes(vlIntCont).CantidadUM <> 0) Then
                        vgLngTotalLotesxArt = vgLngTotalLotesxArt + 1
                        vglngTotalCantUVLotexArt = vglngTotalCantUVLotexArt + agLotes(vlIntCont).CantidadUV
                        vglngTotalCantUMLotexArt = vglngTotalCantUMLotexArt + agLotes(vlIntCont).CantidadUM
                        ReDim Preserve alLotexArt(vgLngTotalLotesxArt)
                        alLotexArt(vgLngTotalLotesxArt).Articulo = agLotes(vlIntCont).Articulo
                        alLotexArt(vgLngTotalLotesxArt).Borrado = IIf(vgPantallaAnterior = "frmKardexPaciente", "*", "")
                        alLotexArt(vgLngTotalLotesxArt).CantidadUV = agLotes(vlIntCont).CantidadUV
                        alLotexArt(vgLngTotalLotesxArt).CantidadUM = agLotes(vlIntCont).CantidadUM
                        alLotexArt(vgLngTotalLotesxArt).CantidadUVInicio = agLotes(vlIntCont).CantidadUVInicio
                        alLotexArt(vgLngTotalLotesxArt).CantidadUMInicio = agLotes(vlIntCont).CantidadUMInicio
                        alLotexArt(vgLngTotalLotesxArt).fechaCaducidad = agLotes(vlIntCont).fechaCaducidad
                        alLotexArt(vgLngTotalLotesxArt).lote = agLotes(vlIntCont).lote
                        'alLotexArt(vgLngTotalLotesxArt).TipoAccion = agLotes(vlintCont).TipoAccion
                    End If
                End If
            End If
        End If
    Next vlIntCont

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaArregloLocal"))
End Sub

Private Sub pLlenaArregloArticulos()
On Error GoTo NotificaError
    
    vllngTotalLotesEnArticulo = 0
    If vllngNumHojaConsumo <> 0 Then
        Set vlrsLotes = frsEjecuta_SP(Trim(vlstrChrCveArticulo) & "|" & vllngNumHojaConsumo & "|" & vlintNoDepartamento, "SP_IVSELARTICULOSLOTESCONSUMO", , , , True)
    Else
         Set vlrsLotes = frsEjecuta_SP(Trim(vlstrChrCveArticulo) & "|'*'|" & IIf(vlblnSoloMoviento, Val(vlstrNoMovimientoInicial), "-1") & "|" & vlintNoDepartamento, "SP_IVSELARTICULOSLOTES", , , , True)
       ' Set vlrsLotes = frsEjecuta_SP(Trim(vlstrChrCveArticulo) & "|'*'|" & IIf(vlblnSoloMoviento, Val(vlstrNoMovimientoInicial), "-1") & "|" & 3, "SP_IVSELARTICULOSLOTES", , , , True)
    End If

    Do While Not vlrsLotes.EOF
        vllngTotalLotesEnArticulo = vllngTotalLotesEnArticulo + 1
        ReDim Preserve alArticulos(vllngTotalLotesEnArticulo)
        alArticulos(vllngTotalLotesEnArticulo).Articulo = Trim(vlrsLotes!Articulo)
        alArticulos(vllngTotalLotesEnArticulo).CantidadUV = vlrsLotes!intExistenciaDeptouv
        alArticulos(vllngTotalLotesEnArticulo).CantidadUM = vlrsLotes!intexistenciadeptoum
        alArticulos(vllngTotalLotesEnArticulo).fechaCaducidad = vlrsLotes!dtmFechaCaducidad
        alArticulos(vllngTotalLotesEnArticulo).lote = Trim(vlrsLotes!lote)
        vlrsLotes.MoveNext
    Loop
    vlrsLotes.Close
    'If vlrsLotes.RecordCount <> 0 Then
    '    vlrsLotes.MoveFirst
        'For vlIntCont = 1 To vlrsLotes.RecordCount
        '    vllngTotalLotesEnArticulo = vllngTotalLotesEnArticulo + 1
        '    ReDim Preserve alArticulos(vllngTotalLotesEnArticulo)
        '    'alArticulos(vllngTotalLotesEnArticulo).Articulo = Trim(vlrsLotes!chrcvearticulo)
        '    alArticulos(vllngTotalLotesEnArticulo).Articulo = Trim(vlrsLotes!Articulo)
        '    alArticulos(vllngTotalLotesEnArticulo).CantidadUV = vlrsLotes!intExistenciaDeptouv
        '    alArticulos(vllngTotalLotesEnArticulo).CantidadUM = vlrsLotes!intexistenciadeptoum
        '    alArticulos(vllngTotalLotesEnArticulo).fechaCaducidad = vlrsLotes!dtmFechaCaducidad
        '    'alArticulos(vllngTotalLotesEnArticulo).Lote = Trim(vlrsLotes!chrlote)
        '    alArticulos(vllngTotalLotesEnArticulo).Lote = Trim(vlrsLotes!Lote)
        '    vlrsLotes.MoveNext
        'Next vlIntCont
    'End If

    pLimpiaConfArregloArt
    
    'VACIO LO ENCONTRADO EN UN GRID PARA QUE MUESTRE LOS LOTES DISPONIBLES PARA REBAJAR
    'If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
    'If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
    'With vsfArticulos
    '    .Clear
    '    .Cols = 5
    '    .Rows = 1
    '    For vlIntCont = 1 To vllngTotalLotesEnArticulo
    '    'Do While X <= vllngTotalLotesEnArticulo
    '        .Rows = .Rows + 1
    '        .TextMatrix(vlIntCont, 1) = Trim(Str(alArticulos(vlIntCont).CantidadUV))
    '        .TextMatrix(vlIntCont, 2) = Trim(Str(alArticulos(vlIntCont).CantidadUM))
    '        .TextMatrix(vlIntCont, 3) = Trim(alArticulos(vlIntCont).Lote)
    '        .TextMatrix(vlIntCont, 4) = Trim(Format(alArticulos(vlIntCont).fechaCaducidad, "DD/MMM/YYYY"))
    '    Next vlIntCont
        'X = X + 1
        'Loop
    '    .FixedCols = 1
    '    .FixedRows = 1
    '    .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad"
    '    .ColWidth(0) = 0                'Fix
    '    .ColWidth(1) = 1200              'Cantidaduv
    '    .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1200) 'Cantidadum
    '    .ColWidth(3) = 1200             'Lote
    '    .ColWidth(4) = 1350             'Fecha de caducidad
        
    '    .ColAlignment(4) = flexAlignRightBottom
    '    .ColAlignment(3) = flexAlignLeftBottom
        'If .Rows < 20 Then .Rows = 21   'Para que no se vea tan solita
    'End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaArregloArticulos"))
End Sub

Private Sub vsfArticulos_DblClick()
On Error GoTo NotificaError
Dim vlstrStyle, vlstrResponse, MyString

    If vsfArticulos.Row > -1 Then
        If vsfArticulos.TextMatrix(vsfArticulos.Row, 3) <> "" Then
            txtLote.Text = Trim(vsfArticulos.TextMatrix(vsfArticulos.Row, 3))
            txtCantidadEntUV = Val(vsfArticulos.TextMatrix(vsfArticulos.Row, 1))
            txtCantidadEntUM = Val(vsfArticulos.TextMatrix(vsfArticulos.Row, 2))
            txtFechaCaduce = vsfArticulos.TextMatrix(vsfArticulos.Row, 4)
            pEnfocaTextBox txtCantDevol
            Exit Sub
        Else
            pEnfocaTextBox txtLote
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":vsfArticulos_DblClick"))
End Sub

'Private Sub prebaja(vlCantUVtmp As Long, vlCantUMtmp As Long, vlCantRTmp As Long)
Private Function fblnrebaja(vlCantUVtmp As Long, vlCantUMtmp As Long, vlCantRTmp As Long) As Boolean
Dim vSuficiente As Integer
Dim vlblnSuficiente As Boolean
Dim vlCantUV As Long
Dim vlCantUM As Long
Dim vlCantR As Long

    fblnrebaja = False

    vlCantUV = vlCantUVtmp
    vlCantUM = vlCantUMtmp
    vlCantR = vlCantRTmp
    
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    ' VER SI EXISTE PARA AGREGARLE LO QUE YA TENIA SELECCIONADO
    For vlintcont2 = 1 To vgLngTotalLotesxArt
        If Trim(txtLote.Text) = Trim(alLotexArt(vlintcont2).lote) And Trim(alLotexArt(vlintcont2).Borrado) <> "*" Then
            If vlstrTipoMovimiento = "UV" Then
                vlCantR = vlCantR + alLotexArt(vlintcont2).CantidadUV + (alLotexArt(vlintcont2).CantidadUM / vlintContenidoArt)
            Else
                vlCantR = vlCantR + alLotexArt(vlintcont2).CantidadUM + (alLotexArt(vlintcont2).CantidadUV * vlintContenidoArt)
            End If
            Exit For
        End If
    Next vlintcont2

    vlblnSuficiente = True

    If vlstrTipoMovimiento = "UV" Then     'Descuento en unidad ALTERNA
        If vlCantUV >= vlCantR Then
            vlCantUV = vlCantUV - vlCantR
        Else
            If vlCantUV > 0 Then
                vlCantR = vlCantR - vlCantUV
                vlCantUV = 0
            End If
            If vlCantUM / vlintContenidoArt >= vlCantR Then
                vlCantUM = vlCantUM - vlCantR * vlintContenidoArt
            Else
                vlblnSuficiente = False
            End If
        End If
    Else                                    'Descuento en unidad MINIMA
        If vlCantUM >= vlCantR Then
            vlCantUM = vlCantUM - vlCantR
        Else
            If vlCantUM > 0 Then
                vlCantR = vlCantR - vlCantUM
                vlCantUM = 0
            End If
            If vlCantUV * vlintContenidoArt >= vlCantR Then
                vlCantUM = Round(((vlCantUV * vlintContenidoArt - vlCantR) / vlintContenidoArt - Fix((vlCantUV * vlintContenidoArt - vlCantR) / vlintContenidoArt)) * vlintContenidoArt, 0)
                vlCantUV = Fix((vlCantUV * vlintContenidoArt - vlCantR) / vlintContenidoArt)
            Else
                vlblnSuficiente = False
            End If
        End If
    End If
    If vlblnSuficiente = True Then
        fblnrebaja = True
    Else
        MsgBox "Lote " & Trim(txtLote.Text) & " no tiene la cantidad para rebajar", vbCritical, "Mensaje"
    End If
End Function

Private Sub vsfArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = 13 Then vsfArticulos_DblClick
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":vsfArticulos_KeyDown"))
End Sub

Private Sub pCalculaLotesDisponibles()
    Dim vlIntCont As Integer
    On Error GoTo NotificaError
    
    vllngLotesDisp = 0
    
    If vsfArticulos.Rows > 0 Then
        For vlIntCont = 1 To vsfArticulos.Rows - 1
            vllngLotesDisp = vllngLotesDisp + (Val(vsfArticulos.TextMatrix(vlIntCont, 1)) * vlintContenidoArt)
            vllngLotesDisp = vllngLotesDisp + Val(vsfArticulos.TextMatrix(vlIntCont, 2))
        Next vlIntCont
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCalculaLotesDisponibles"))
End Sub

Private Sub pCalculaLotesAcumulados()
    Dim vlIntCont As Integer
    On Error GoTo NotificaError
    
    vllngLotesAcum = 0
    
    If VSFArticuloLote.Rows > 0 Then
        For vlIntCont = 1 To VSFArticuloLote.Rows - 1
            vllngLotesAcum = vllngLotesAcum + (Val(VSFArticuloLote.TextMatrix(vlIntCont, 1)) * vlintContenidoArt)
            vllngLotesAcum = vllngLotesAcum + Val(VSFArticuloLote.TextMatrix(vlIntCont, 2))
        Next vlIntCont
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCalculaLotesAcumulados"))
End Sub

Private Sub pLimpiaConfArregloArt()
 On Error GoTo NotificaError

'VACIO LO ENCONTRADO EN UN GRID PARA QUE MUESTRE LOS LOTES DISPONIBLES PARA REBAJAR
    If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
    If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
    With vsfArticulos
        .Clear
        .Cols = 5
        .Rows = 1
        
        For vlIntCont = 1 To vllngTotalLotesEnArticulo
        
            .Rows = .Rows + 1
            .TextMatrix(vlIntCont, 1) = Trim(str(alArticulos(vlIntCont).CantidadUV))
            .TextMatrix(vlIntCont, 2) = Trim(str(alArticulos(vlIntCont).CantidadUM))
            .TextMatrix(vlIntCont, 3) = Trim(alArticulos(vlIntCont).lote)
            .TextMatrix(vlIntCont, 4) = Trim(Format(alArticulos(vlIntCont).fechaCaducidad, "DD/MMM/YYYY"))
        Next vlIntCont
        
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad"
        .ColWidth(0) = 0                'Fix
        .ColWidth(1) = 1100              'Cantidaduv 1200
        .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1100) 'Cantidadum 1200
        .ColWidth(3) = 1400             'Lote 1200
        .ColWidth(4) = 2110             'Fecha de caducidad 2110
        
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftBottom
        'Alineacion de la fecha de caducidad porque no se dejaba
        For vlIntCont = 1 To vsfArticulos.Rows - 1
            vsfArticulos.Row = vlIntCont
            vsfArticulos.Col = 4
            vsfArticulos.CellAlignment = flexAlignRightBottom
        Next vlIntCont
        'If .Rows < 20 Then .Rows = 21   'Para que no se vea tan solita
    End With
        

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCalculaLotesAcumulados"))
End Sub

'caso 20417 GIRM
Public Function BorrarLote()
Dim vllngContad As Long

  'Caso 20453 GIRM
  If (Not agLotes) = -1 Then
    Exit Function
  End If
  'Caso 20453
  
    For vllngContad = 1 To vgLngTotalLotesxMov
        agLotes(vllngContad).Borrado = "*"
        agLotes(vllngContad).CantidadUM = 0
        agLotes(vllngContad).CantidadUV = 0
        agLotes(vllngContad).Movimiento = 0
        agLotes(vllngContad).Devolucion = 0
    Next vllngContad
    'vgLngTotalLotesxMov = 1
    vgLngTotalLotesxMov = 0 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
    If UBound(agLotes) > 1 Then
        ReDim Preserve agLotes(1)
    End If
End Function
'Caso 20417
