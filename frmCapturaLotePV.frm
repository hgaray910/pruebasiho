VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmCapturaLotePV 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de caducidades"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraLotes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   9730
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Clave del artículo"
         Top             =   195
         Visible         =   0   'False
         Width           =   8370
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
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Descripción del artículo"
         Top             =   195
         Width           =   8370
      End
      Begin VB.CheckBox ChkModoCBarra 
         BackColor       =   &H80000005&
         Caption         =   "Modo código de barras"
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
         Left            =   6855
         TabIndex        =   12
         Top             =   -1560
         Width           =   2655
      End
      Begin VB.TextBox TxtCBarra 
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
         Left            =   1980
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Código de barras"
         Top             =   -1560
         Width           =   4770
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
         Height          =   250
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   15
         Top             =   -1500
         Width           =   1725
      End
   End
   Begin VB.TextBox txtTotalARecibir 
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Total de unidades recibidas"
      Top             =   810
      Width           =   1230
   End
   Begin VB.Frame FraCaptura 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3810
      Begin VB.TextBox txtCantRecep 
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
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   4
         ToolTipText     =   "Cantidad de artículos "
         Top             =   240
         Width           =   1360
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
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   3
         ToolTipText     =   "Lote"
         Top             =   660
         Width           =   1365
      End
      Begin MSMask.MaskEdBox txtFechaCaduce 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "Fecha de caducidad"
         Top             =   1080
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblCantOrden 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cantidad"
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
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Width           =   945
      End
      Begin VB.Label lblLote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Número de lote"
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
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   1530
      End
      Begin VB.Label lblFechaCad 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Fecha de caducidad"
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
         Left            =   180
         TabIndex        =   6
         Top             =   1155
         Width           =   2040
      End
   End
   Begin VB.Frame frmGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   4635
      TabIndex        =   0
      Top             =   2350
      Width           =   720
      Begin MyCommandButton.MyButton cmdGrabarRegistro 
         Height          =   600
         Left            =   60
         TabIndex        =   1
         ToolTipText     =   "Confirmar"
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
         Picture         =   "frmCapturaLotePV.frx":0000
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmCapturaLotePV.frx":0984
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFArticuloLote 
      Height          =   1125
      Left            =   3960
      TabIndex        =   17
      ToolTipText     =   "Lotes y caducidades"
      Top             =   1215
      Width           =   5895
      _cx             =   10398
      _cy             =   1984
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
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
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
      SelectionMode   =   0
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
      FormatString    =   $"frmCapturaLotePV.frx":1308
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Total recibido"
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
      Left            =   4080
      TabIndex        =   18
      Top             =   870
      Width           =   1320
   End
End
Attribute VB_Name = "frmCapturaLotePV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjInventario
'| Nombre del Formulario    : frmCapturaLote
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar la captura de lotes de un articulo determinado
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : ESanchez
'| Autor                    : ESanchez
'| Fecha de Creación        : 30/Oct/2005
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
Option Explicit

Private Declare Function shellexecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const SW_NORMAL = 1


'Parametros iniciales de la forma
Public vlstrTablaReferencias As String
Public vlstrChrCveArticulo As String
Public vlintIdArticulo As Long
Public vlstrNoMovimiento As Long
Public vlstrNoDevolucion As Long
Public vlstrTipoMovimiento As String        'UM o UV, cual se esta solicitando para captura
Public vlintNoDepartamento As Integer
Public vlintContenidoArt As Long            'Contenido en caso de UV o 1 en UM
Public vlStrTitUM As String                 'Que descripcion de unidad minima
Public vlStrTitUV As String                 'Que descripcion de unidad alterna
Public vlblnValidaMovimiento As Boolean     'Indica si se valida que el articulo pueda estar en el arreglo de lotes para dos o mas movimientos
Public vlintTipoAccion As Integer           'Se usa unicamente para la captura fisica de inventario: 1 = Entrada UV, 2 = Entrada UM, 3 = Salida UV y 4 = Salida UM
Public vllngNoRecepcion As Long             'Id de la recepcion
Public LoteB, UrlB As String

' Variable locales
Dim vlrsBoletines As New ADODB.Recordset

Dim vlstrsql As String
Dim vlstrFecCadDBO As String
Dim vlIntCont As Long
Dim vllngTotalAsignados As Long
Dim vlblnExiste As Boolean
Dim alLotexArt() As varLotes
Dim vlstrStyle, vlstrResponse, MyString
Dim vldtmfechaServer As Date
Dim vldtmfechaNoCaduca As Date
Dim vlblnEstatusAgregar As Boolean
Dim vgLngTotalLotesxArt As Long
Dim vglngTotalCantLotexArt As Long
Dim vglngTotalCantUVLotexArt As Long
Dim vglngTotalCantUMLotexArt As Long
Dim vgblnManejaLote As Boolean

'Public vgblnConfirmaCargo As Boolean 'Confirma que se van a realizar los cargos
'Public vgintCuentaFlete As Long

Dim strCveArticulo As String
Dim strLote As String
Dim strCBarra As String

Private Sub ChkModoCBarra_Click()
    TxtCBarra.Enabled = (ChkModoCBarra.Value = vbChecked)
    txtCantRecep.Enabled = Not (ChkModoCBarra.Value = vbChecked)
    txtLote.Enabled = Not (ChkModoCBarra.Value = vbChecked)
    txtFechaCaduce.Enabled = Not (ChkModoCBarra.Value = vbChecked)
    txtTotalARecibir.Enabled = Not (ChkModoCBarra.Value = vbChecked)
End Sub

Private Sub cmdGrabarRegistro_Click()

Dim strVariosLotes As String
Dim strRenglon  As String
Dim vlIntCols As Integer

On Error GoTo NotificaError

    If fblnContinuar() Then
        For vlIntCont = 1 To vgLngTotalLotesxMov
            If Trim(vlstrTablaReferencias) = "SAJ" Or Trim(vlstrTablaReferencias) = "EAJ" Then ' Si es Captura fisica unicamente limpia los lotes que sean del tipo de accion: 1 = Entrada UV, 2 = Entrada UM, 3 = Salida UV y 4 = Salida UM
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
                    agLotes(vlIntCont).TablaRelacion = ""
                    'agLotes(vlintCont).TipoAccion = 0
                End If
            Else
                If Trim(agLotes(vlIntCont).Articulo) = Trim(txtClaveArticulo.Text) _
                    And (Not vlblnValidaMovimiento Or (vlblnValidaMovimiento And agLotes(vlIntCont).Movimiento = vlstrNoMovimiento And agLotes(vlIntCont).Devolucion = vlstrNoDevolucion) Or (vlblnValidaMovimiento And Trim(agLotes(vlIntCont).TablaRelacion) = "EBO" And Trim(vlstrTablaReferencias) <> "ERE")) Then
                    agLotes(vlIntCont).Articulo = ""
                    agLotes(vlIntCont).Borrado = "*"
                    agLotes(vlIntCont).CantidadUM = 0
                    agLotes(vlIntCont).CantidadUV = 0
                    agLotes(vlIntCont).CantidadUMInicio = 0
                    agLotes(vlIntCont).CantidadUVInicio = 0
                    agLotes(vlIntCont).lote = ""
                    agLotes(vlIntCont).Movimiento = 0
                    agLotes(vlIntCont).Devolucion = 0
                    agLotes(vlIntCont).TablaRelacion = ""
                    'agLotes(vlintCont).TipoAccion = 0
                End If
            End If
        Next vlIntCont

        For vlIntCont = 1 To vgLngTotalLotesxArt
            vgLngTotalLotesxMov = vgLngTotalLotesxMov + 1
            ReDim Preserve agLotes(vgLngTotalLotesxMov)
            agLotes(vgLngTotalLotesxMov).Articulo = alLotexArt(vlIntCont).Articulo
            agLotes(vgLngTotalLotesxMov).Borrado = alLotexArt(vlIntCont).Borrado
            agLotes(vgLngTotalLotesxMov).CantidadUM = alLotexArt(vlIntCont).CantidadUM
            agLotes(vgLngTotalLotesxMov).CantidadUV = alLotexArt(vlIntCont).CantidadUV
            agLotes(vgLngTotalLotesxMov).CantidadUMInicio = alLotexArt(vlIntCont).CantidadUMInicio
            agLotes(vgLngTotalLotesxMov).CantidadUVInicio = alLotexArt(vlIntCont).CantidadUVInicio
            agLotes(vgLngTotalLotesxMov).fechaCaducidad = alLotexArt(vlIntCont).fechaCaducidad
            agLotes(vgLngTotalLotesxMov).lote = alLotexArt(vlIntCont).lote
            agLotes(vgLngTotalLotesxMov).Movimiento = alLotexArt(vlIntCont).Movimiento
            agLotes(vgLngTotalLotesxMov).Devolucion = alLotexArt(vlIntCont).Devolucion
            agLotes(vgLngTotalLotesxMov).TablaRelacion = alLotexArt(vlIntCont).TablaRelacion
            agLotes(vgLngTotalLotesxMov).TipoAccion = alLotexArt(vlIntCont).TipoAccion
        Next vlIntCont

        If vgblnTrazabilidad Then
            If VSFArticuloLote.Rows <= 2 Then
                frmPOS.lngCantidad = Val(txtTotalARecibir.Text)
                frmPOS.strCveArticulo = strCveArticulo
                frmPOS.strLote = strLote
                agLotes(1).CantidadUM = 0 ' Limpia  cantidades para no buscar ubicaciones
                agLotes(1).CantidadUV = 0 ' Limpia  cantidades para no buscar ubicaciones
            Else
                For vlIntCont = 1 To VSFArticuloLote.Rows - 1
                    strRenglon = VSFArticuloLote.TextMatrix(vlIntCont, IIf(vlstrTipoMovimiento = "UV", 1, 2)) & ";" & _
                                 VSFArticuloLote.TextMatrix(vlIntCont, 5) & ";" & _
                                 VSFArticuloLote.TextMatrix(vlIntCont, 3) & ";" & _
                                 VSFArticuloLote.TextMatrix(vlIntCont, 6)
                    strVariosLotes = strVariosLotes & IIf(vlIntCont = 1, "", "|") & strRenglon
                    
                    agLotes(vlIntCont).CantidadUM = 0 ' Limpia  cantidades para no buscar ubicaciones
                    agLotes(vlIntCont).CantidadUV = 0 ' Limpia  cantidades para no buscar ubicaciones
                Next vlIntCont
                
                frmPOS.lngCantidad = Val(txtTotalARecibir.Text)
                frmPOS.strVariosLotes = strVariosLotes
            End If
            
        End If

        vgblnCapturoLoteYCaduc = True
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
End Sub


Private Sub Form_Activate()
On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon

    txtClaveArticulo = vlstrChrCveArticulo
    pLlenaArregloLocal

    vllngTotalAsignados = 0
    pVaciaArregloGrid

    vllngTotalAsignados = IIf(vlstrTipoMovimiento = "UV", vglngTotalCantUVLotexArt, vglngTotalCantUMLotexArt)
    If vllngTotalAsignados <= Val(txtTotalARecibir) Then
        txtCantRecep.Text = Val(txtTotalARecibir) - vllngTotalAsignados
    Else
        If vllngTotalAsignados > Val(txtTotalARecibir) Then
            txtCantRecep.Text = "" 'vllngTotalAsignados - Val(txtTotalARecibir)
        End If
    End If

    If Val(txtCantRecep.Text) = 0 And vllngTotalAsignados <= Val(txtTotalARecibir) Then
        FraCaptura.Enabled = False
        cmdGrabarRegistro.Enabled = True
        cmdGrabarRegistro.SetFocus
    Else
        FraCaptura.Enabled = True
        pEnfocaTextBox txtCantRecep
    End If

    CfgPantallaTrazabilidad

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 27 Then
        If Val(txtCantRecep.Text) <> 0 And txtCantRecep.Text <> "" Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg("17"), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                vgblnCapturoLoteYCaduc = False
                Unload Me
            Else
                pEnfocaTextBox txtCantRecep
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
    cmdGrabarRegistro.Enabled = False
    vldtmfechaServer = fdtmServerFecha
    vldtmfechaNoCaduca = DateAdd("d", 1, vldtmfechaServer)

    If ChkModoCBarra.Value = False Then
        TxtCBarra.Enabled = False
    End If
End Sub

Private Sub txtCantRecep_GotFocus()
    pSelTextBox txtCantRecep
End Sub

Private Sub txtCantRecep_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If Val(txtCantRecep.Text) = 0 Then
            txtCantRecep.SetFocus
        Else
            Call pEnfocaTextBox(txtLote)
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantRecep_KeyPress"))
End Sub

Private Sub TxtCBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnEncontrado As Boolean
    Dim vlstrSentencia As String
    Dim vglngRows As Long
    Dim rs As New ADODB.Recordset

 On Error GoTo NotificaError

    blnEncontrado = False

    If TxtCBarra.Text = "" Then
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyReturn

            'Ahora traemos los datos de la etiqueta
            vlstrSentencia = "Select * from ivetiqueta where rtrim(ltrim(INTIDETIQUETA)) = '" & Replace(Trim(TxtCBarra.Text), "'", "''") & "' And chrcvearticulo = " & txtClaveArticulo.Text

            'Se realiza la consulta de la forma estandar
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)

            If rs.RecordCount > 0 Then
                With rs
                    vglngRows = VSFArticuloLote.Rows - 1
                    txtTotalARecibir.Text = Val(IIf(vglngRows = 0, 0, IIf(Val(VSFArticuloLote.TextMatrix(vglngRows, 1)) > 0, Val(VSFArticuloLote.TextMatrix(vglngRows, 1)), Val(VSFArticuloLote.TextMatrix(vglngRows, 2))))) + 1
                    txtCantRecep.Text = 1
                    txtLote.Text = Trim(!chrlote)
                    txtFechaCaduce.Text = Trim(!dtmFechaCaducidad)
                    txtClaveArticulo.Text = Trim(!chrcvearticulo)
                    txtDescripcionLgaArt.Text = Trim(TraerDescripcionArticulo())
                    Call txtFechaCaduce_KeyDown(vbKeyReturn, 0)
                    TxtCBarra.Text = ""
                    blnEncontrado = True
                End With
                rs.Close
                TxtCBarra.SetFocus
            End If

            If Not blnEncontrado Then
                MsgBox "El código de barras proporcionado no coincide." & vbCrLf & "Por favor, inténtelo nuevamente.", vbCritical, "Mensaje"
                TxtCBarra.Text = ""
                TxtCBarra.SetFocus
            End If

    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCBarra_KeyDown"))
End Sub

Private Sub TxtCBarra_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtFechaCaduce_GotFocus()
    pSelMkTexto txtFechaCaduce
End Sub

Private Sub txtFechaCaduce_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------------
' Procedimiento para validar cuando se presiona la tecla <Esc> para salir
'--------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintResul As Integer
    Dim vlintArray As Integer
    Dim vlstrMensaje As String
    Dim strSQL As String
    Dim ProxNum As String
    Dim FechaCaduce As String
    Dim vlLote As String
    Dim VlArrayUrl
    Dim vlDescripcion As String
    Dim vlBand As Boolean



    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtCantRecep.Text) = 0 Then
                pEnfocaTextBox txtCantRecep
            Else
                If Not (fblnValidaFechaCaduce) Then
                    pEnfocaMkTexto txtFechaCaduce
                    Exit Sub
                End If

                vlblnEstatusAgregar = True
                FechaCaduce = Format(txtFechaCaduce.Text, "DD/MM/YYYY")
                vlLote = UCase(txtLote.Text)

                If vgblnTrazabilidad Then
                    Dim rsTrazabilidad As ADODB.Recordset

                    strLote = vlLote
                    strCBarra = Trim(TxtCBarra.Text)

                    Set rsTrazabilidad = frsRegresaRs("SELECT CHRCVEARTICULO FROM ivetiqueta WHERE intidetiqueta = '" & TxtCBarra.Text & "'", adLockOptimistic, adOpenDynamic)

                    If rsTrazabilidad.RecordCount > 0 Then
                        strCveArticulo = Trim(rsTrazabilidad!chrcvearticulo)
                    End If

                    rsTrazabilidad.Close

                End If
                

                pAgregaenGrid

                If Not vgblnForzarLoteYCaduc Then
                   cmdGrabarRegistro.Enabled = True
                End If

                If vlblnEstatusAgregar = True Then
                    If Val(txtCantRecep.Text) = 0 Then
                        cmdGrabarRegistro.SetFocus
                    Else
                        pEnfocaTextBox txtCantRecep
                    End If


                     If LoteB <> "" And UrlB <> "" Then

                        vlBand = True
                        VlArrayUrl = Split(UrlB, "/")
                        vlDescripcion = VlArrayUrl(UBound(VlArrayUrl))

                        'If frmRecepciones.intIndexBol > 0 Then
                         For vlintArray = 0 To UBound(aBoletines)
                             If aBoletines(vlintArray).lote = LoteB Then
                                 vlBand = False
                                 Exit For
                             End If

                         Next vlintArray
                        'End If

                        If vlBand = True Then
                           'ReDim Preserve aBoletines(frmRecepciones.intIndexBol)

                           'aBoletines(frmRecepciones.intIndexBol).lote = CStr(LoteB)
                           'aBoletines(frmRecepciones.intIndexBol).Url = CStr(UrlB)
                           'aBoletines(frmRecepciones.intIndexBol).Descripcion = CStr(vlDescripcion)
                           'aBoletines(frmRecepciones.intIndexBol).FechaCaduca = FechaCaduce


                           'frmRecepciones.intIndexBol = frmRecepciones.intIndexBol + 1
                           'frmOtrasEntSal.intIndexBol = frmOtrasEntSal.intIndexBol + 1
                        End If

                    End If

                End If

            End If

        Case vbKeyEscape
            vgblnCapturoLoteYCaduc = False
            Unload Me
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaCaduce_KeyDown"))
End Sub

Private Sub txtLote_GotFocus()
    pSelTextBox txtLote
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim rsFechaCad As New ADODB.Recordset
    Dim X
    Dim Resultado As String

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = 13 Then
        If txtLote.Text <> "" Then
            vlblnExiste = False
            vlstrFecCadDBO = ""


     'Verifica si esxite licencia para alerta sanitaria
     If ObtenerLicenciaBoletin Then

               '-------------------------------------------------------
               'Verifica si esxite alerta sanitaria
               LoteB = ""
               UrlB = ""
               vlstrResponse = 0
            If Mid(txtClaveArticulo.Text, 1, 1) = "1" Then
              Resultado = Replace(enviarGet("https://boletinados.hospisoft.mx/api/boletin?LoteBuscar=" & txtLote.Text), Chr(34), "'")

                   If Len(Resultado) > 10 Then
                    Dim ContCad As Integer

                       'Numero de Lote
                       ContCad = InStr(1, Resultado, "':'") + 3
                       LoteB = Mid(Resultado, ContCad, (InStr(ContCad, Resultado, "','") - ContCad))

                       'Url de la notificación
                       ContCad = InStr(ContCad, Resultado, "':'") + 3
                       UrlB = Mid(Resultado, ContCad, ((InStr(ContCad, Resultado, ".pdf") + 4) - ContCad))

                         MsgBox ("El lote " & LoteB & " tiene emitida una alerta sanitaria relacionada." & Chr(13)), vbOKOnly + vbExclamation, "Mensaje"
                         X = shellexecute(Me.hwnd, "Open", UrlB, &O0, &O0, SW_NORMAL)


                        'Fue notificado
                        vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                        vlstrResponse = MsgBox("Fue notificado de una alerta sanitaria relacionada con el lote  " & LoteB & "." & Chr(13) & "¿Desea continuar con la recepción?", vlstrStyle, "Mensaje")




                   End If
            End If
            '--------------------------Fin Alerta sanitaria
        End If



            '-------------------------------------------------------
            'CHECA FECHAS CON RESPECTO A LO QUE YA SE HAYA CAPTURADO
            For vlIntCont = 1 To vgLngTotalLotesxArt
                If Trim(txtLote.Text) = Trim(alLotexArt(vlIntCont).lote) And Trim(alLotexArt(vlIntCont).Borrado) <> "*" Then
                    txtFechaCaduce.Text = Format(CStr(alLotexArt(vlIntCont).fechaCaducidad), "DD/MM/YYYY")
                    vlblnExiste = True
                    Exit For
                End If
            Next vlIntCont
            
            ' si no existe como capturado
            If vlblnExiste = False Then
                'Checa si existe en la base de datos para traerse la fecha de caducidad
                Set rsFechaCad = frsRegresaRs("Select dtmfechacaducidad from IVLOTES where Rtrim(LTRIM(chrlote)) = '" & Trim(txtLote.Text) & "' AND CHRCVEARTICULO = '" & Trim(txtClaveArticulo) & "'")
                If rsFechaCad.RecordCount > 0 Then
                    txtFechaCaduce.Text = Format(CStr(rsFechaCad!dtmFechaCaducidad), "DD/MM/YYYY")
                    vlstrFecCadDBO = Format(CStr(rsFechaCad!dtmFechaCaducidad), "DD/MM/YYYY")
                Else
                    vlstrFecCadDBO = ""
                End If
            End If


             If vlstrResponse <> vbYes And LoteB <> "" Then
                 pEnfocaTextBox txtLote
             Else
                 pEnfocaMkTexto txtFechaCaduce
             End If

        Else
            pEnfocaTextBox txtLote
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 45 And Not KeyAscii = 47 And Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123) And Not (KeyAscii > 39 And KeyAscii < 44) And Not KeyAscii = 35 And Not KeyAscii = 46 And Not KeyAscii = 58 Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtLote_KeyPress"))
End Sub

Private Sub VSFArticuloLote_DblClick()
On Error GoTo NotificaError

    If VSFArticuloLote.Row > -1 Then
        If VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3) <> "" Then
            vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
            vlstrResponse = MsgBox("Seguro de eliminar lote " & Trim(VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3)) & " ?", vlstrStyle, "Mensaje")
            If vlstrResponse = vbYes Then

                If vgLngTotalLotesxArt > 0 Then
                    For vlIntCont = 1 To vgLngTotalLotesxArt
                        If Trim(VSFArticuloLote.TextMatrix(VSFArticuloLote.Row, 3)) = Trim(alLotexArt(vlIntCont).lote) Then

                            vllngTotalAsignados = vllngTotalAsignados - IIf(vlstrTipoMovimiento = "UV", alLotexArt(vlIntCont).CantidadUV, alLotexArt(vlIntCont).CantidadUM)
                            txtCantRecep.Text = Val(txtTotalARecibir.Text) - vllngTotalAsignados

                            alLotexArt(vlIntCont).CantidadUM = IIf(vlstrTipoMovimiento = "UM", 0, alLotexArt(vlIntCont).CantidadUM)
                            alLotexArt(vlIntCont).CantidadUV = IIf(vlstrTipoMovimiento = "UV", 0, alLotexArt(vlIntCont).CantidadUV)
                            If (Val(alLotexArt(vlIntCont).CantidadUV) + Val(alLotexArt(vlIntCont).CantidadUM)) = 0 Then
                                alLotexArt(vlIntCont).Borrado = "*"
                            End If

                            If Not vgblnForzarLoteYCaduc Then
                                If Val(txtCantRecep.Text) = 0 Then
                                    FraCaptura.Enabled = False
                                    cmdGrabarRegistro.Enabled = True
                                    cmdGrabarRegistro.SetFocus
                                Else
                                    FraCaptura.Enabled = True
                                    pEnfocaTextBox txtCantRecep
                                End If
                            Else
                                If Val(txtCantRecep.Text) = 0 Then
                                    FraCaptura.Enabled = False
                                    cmdGrabarRegistro.Enabled = True
                                    cmdGrabarRegistro.SetFocus
                                Else
                                    FraCaptura.Enabled = True
                                    cmdGrabarRegistro.Enabled = False
                                    pEnfocaTextBox txtCantRecep
                                End If
                            End If
                            Exit For
                        End If
                    Next vlIntCont
                End If
                pVaciaArregloGrid
            End If
            pEnfocaTextBox txtCantRecep
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":VSFArticuloLote_DblClick"))
End Sub

'Validacion para no permitir borrar un lote que ya tiene movimientos con salidas
Private Function fblnContinuar() As Boolean
On Error GoTo NotificaError
    Dim lintContador As Integer
    Dim rs As New ADODB.Recordset
    Dim lintContenido As Integer
    Dim llngCantidad As Long
    Dim lblnFueBorrado As Boolean
    Dim llngReferencia As Long

    fblnContinuar = True
    llngReferencia = IIf(Trim(vlstrTablaReferencias) = "EOE" Or Trim(vlstrTablaReferencias) = "SOS", Val(vlstrNoMovimiento), vllngNoRecepcion)
    If Trim(vlstrTablaReferencias) = "EOE" Or Trim(vlstrTablaReferencias) = "SOS" Or Trim(vlstrTablaReferencias) = "ERE" Then
        For lintContador = 1 To vgLngTotalLotesxArt
            lblnFueBorrado = fblnExisteBorrado(Trim(alLotexArt(lintContador).lote), Trim(alLotexArt(lintContador).Articulo))
            If alLotexArt(lintContador).CantidadUMInicio > 0 Or alLotexArt(lintContador).CantidadUVInicio > 0 Or (alLotexArt(lintContador).Borrado = "" And lblnFueBorrado) Then
                lintContenido = 1
                Set rs = frsRegresaRs("select intContenido from ivArticulo where chrcveArticulo = '" & Trim(alLotexArt(lintContador).Articulo) & "'")
                If rs.RecordCount <> 0 Then lintContenido = rs!intContenido
                rs.Close
                'si el lote fue borrado o se modificó a una cantidad menor, se valida la cantidad minima permitida
                If alLotexArt(lintContador).Borrado = "*" Or (alLotexArt(lintContador).CantidadUVInicio * lintContenido + alLotexArt(lintContador).CantidadUMInicio) > (alLotexArt(lintContador).CantidadUV * lintContenido + alLotexArt(lintContador).CantidadUM Or lblnFueBorrado) Then
                    'se verifica el saldo mínimo de los movimientos siguientes en el kardex de lotes
                    vgstrParametrosSP = "SELECT MIN((KARDEXLOTE.RELEXISTENCIAUV * ARTICULO.INTCONTENIDO) + KARDEXLOTE.RELEXISTENCIAUM) minimo " & _
                                        " FROM IVKARDEXINVENTARIOLOTE KARDEXLOTE INNER JOIN IVARTICULO ARTICULO ON TRIM(KARDEXLOTE.CHRCVEARTICULO) = ARTICULO.CHRCVEARTICULO " & _
                                        " LEFT JOIN IVKARDEXINVENTARIO KARDEX ON KARDEXLOTE.INTNUMMOVIMIENTOKARDEX = KARDEX.INTNUMMOVIMIENTO " & _
                                        " WHERE Trim(KARDEXLOTE.CHRCVEARTICULO) = '" & Trim(alLotexArt(lintContador).Articulo) & "' And KARDEXLOTE.SMICVEDEPARTAMENTO = " & vlintNoDepartamento & _
                                        " And Trim(KARDEXLOTE.CHRLOTE) = '" & Trim(alLotexArt(lintContador).lote) & "' " & _
                                        " AND CASE WHEN TRIM(KARDEXLOTE.VCHTABLARELACION) IN ('EAJ','SAJ') THEN KARDEXLOTE.DTMFECHAINVENTARIO ELSE KARDEX.DTMFECHAHORAMOV END " & _
                                        " > (SELECT MAX(KI.DTMFECHAHORAMOV) FROM IVKARDEXINVENTARIO KI " & _
                                        "   WHERE KI.CHRCVEARTICULO = '" & Trim(alLotexArt(lintContador).Articulo) & "' AND KI.SMICVEDEPARTAMENTO = " & vlintNoDepartamento & _
                                        "   AND TRIM(KI.VCHTABLARELACION) = '" & Trim(vlstrTablaReferencias) & "' AND KI.NUMNUMREFERENCIA = " & llngReferencia & ")"
                    Set rs = frsRegresaRs(vgstrParametrosSP)
                    If rs.RecordCount <> 0 Then
                        ' si es nulo no hubo movimientos para el lote
                        If Not IsNull(rs!Minimo) Then
                            llngCantidad = flngCantidadInicio(Trim(alLotexArt(lintContador).lote), Trim(alLotexArt(lintContador).Articulo), lintContenido)
                            If rs!Minimo <= llngCantidad Then
                                If alLotexArt(lintContador).Borrado = "*" Then
                                    'fue borrado, se revisa si se capturo de nuevo con otra cantidad, si la nueva cantidad es mayor a cero significa que existe en el arreglo con nueva cantidad
                                    If fblnNuevaCantidad(Trim(alLotexArt(lintContador).lote), Trim(alLotexArt(lintContador).Articulo)) = False Then
                                        If ((alLotexArt(lintContador).CantidadUV * lintContenido + alLotexArt(lintContador).CantidadUM) < rs!Minimo) _
                                                Or (alLotexArt(lintContador).CantidadUVInicio * lintContenido + alLotexArt(lintContador).CantidadUMInicio > rs!Minimo) Then
                                            MsgBox "No se puede eliminar el lote " & Trim(alLotexArt(lintContador).lote) & ", existen movimientos de salida del lote.", vbExclamation, "Mensaje"
                                            fblnContinuar = False
                                        End If
                                    End If
                                ElseIf (alLotexArt(lintContador).CantidadUV * lintContenido + alLotexArt(lintContador).CantidadUM) < (llngCantidad - rs!Minimo) Then
                                    MsgBox "La cantidad del lote " & Trim(alLotexArt(lintContador).lote) & " debe ser mayor o igual a " & _
                                            IIf(alLotexArt(lintContador).CantidadUM > 0, (llngCantidad - rs!Minimo), Int((llngCantidad - rs!Minimo) / lintContenido)) & _
                                           " ya que existen movimientos de salida del lote.", vbExclamation, "Mensaje"
                                    fblnContinuar = False
                                End If
                            End If
                        End If
                    End If
                    rs.Close
                End If
            End If
        Next lintContador
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnContinuar"))
End Function

Private Function flngCantidadInicio(strLote As String, strCveArticulo As String, intContenido As Integer) As Long
    Dim lintContador As Integer
    flngCantidadInicio = 0
    For lintContador = 1 To vgLngTotalLotesxArt
        If alLotexArt(lintContador).CantidadUMInicio > 0 Or alLotexArt(lintContador).CantidadUVInicio > 0 And Trim(alLotexArt(lintContador).lote) = strLote And Trim(alLotexArt(lintContador).Articulo) = strCveArticulo Then
            flngCantidadInicio = IIf(alLotexArt(lintContador).CantidadUMInicio > 0, alLotexArt(lintContador).CantidadUMInicio, (alLotexArt(lintContador).CantidadUVInicio * intContenido))
        End If
    Next lintContador
End Function

Private Function fblnNuevaCantidad(strLote As String, strCveArticulo As String) As Boolean
On Error GoTo NotificaError
    Dim lintContador As Integer
    fblnNuevaCantidad = False
    For lintContador = 1 To vgLngTotalLotesxArt
        If alLotexArt(lintContador).Borrado <> "*" And Trim(alLotexArt(lintContador).lote) = strLote And Trim(alLotexArt(lintContador).Articulo) = strCveArticulo Then
            fblnNuevaCantidad = True
        End If
    Next lintContador

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnNuevaCantidad"))
End Function

Private Function fblnExisteBorrado(strLote As String, strCveArticulo As String) As Boolean
    Dim lintContador As Integer
    fblnExisteBorrado = False
    For lintContador = 1 To vgLngTotalLotesxArt
        If alLotexArt(lintContador).Borrado = "*" And Trim(alLotexArt(lintContador).lote) = strLote And Trim(alLotexArt(lintContador).Articulo) = strCveArticulo Then
            fblnExisteBorrado = True
        End If
    Next lintContador
End Function

Private Sub pConfiguraGrid()
On Error GoTo NotificaError 'Manejo del error

    If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
    If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
    With VSFArticuloLote
        .Clear
        .Redraw = False
        .Visible = False
        '.Cols = 5
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .FixedAlignment(1) = flexAlignCenterCenter
        .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad|Cve Artículo"
        .ColWidth(0) = 0        'Fix
        .ColWidth(1) = 1200     'CantidadUV
        .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1200) 'Cantidadum '700      'CantidadUM
        .ColWidth(3) = 1200     'Lote
        .ColWidth(4) = 1350     'Fecha de caducidad
        .ColWidth(5) = 0      'Clave del artículo de trazabilidad
        .ColWidth(6) = 0     'Clave de código de barras

        .ColAlignment(4) = flexAlignRightBottom
        .ColAlignment(3) = flexAlignLeftBottom
        .Redraw = True
        .Visible = True
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub pAgregaenGrid()
Dim vllngQueRenglonExiste As Long
Dim vllngCantidadRenglon As Long
Dim vllngRenglon As Long
Dim vllngResultado As Long

On Error GoTo NotificaError

    If Trim(txtCantRecep.Text) = "" Then
        MsgBox "Necesita capturar una cantidad", vbCritical, "Mensaje"
        vlblnEstatusAgregar = False
        pEnfocaTextBox txtCantRecep
        Exit Sub
    End If

    If Trim(txtLote.Text) = "" Then
        MsgBox "Necesita capturar un lote", vbCritical, "Mensaje"
        vlblnEstatusAgregar = False
        pEnfocaTextBox txtLote
        Exit Sub
    End If
    
    If VSFArticuloLote.Rows > 2 Then
        For vllngRenglon = 1 To VSFArticuloLote.Rows - 1
            If vlstrTipoMovimiento = "UV" Then
                vllngResultado = vllngResultado + CLng(VSFArticuloLote.TextMatrix(vllngRenglon, 1))
            Else
                vllngResultado = vllngResultado + CLng(VSFArticuloLote.TextMatrix(vllngRenglon, 2))
            End If
        Next vllngRenglon
        
        txtTotalARecibir.Text = vllngResultado + Val(txtCantRecep.Text)
    End If
    

    If (vllngTotalAsignados + Val(txtCantRecep.Text)) <= Val(txtTotalARecibir.Text) Then

        vlblnExiste = False

        If vgLngTotalLotesxArt > 0 Then
            For vlIntCont = 1 To vgLngTotalLotesxArt
                If Trim(txtLote.Text) = Trim(alLotexArt(vlIntCont).lote) And Trim(alLotexArt(vlIntCont).Borrado) <> "*" Then
                    '---------------------------------------------
                    'CHECA FECHAS CON RESPECTO A LO QUE YA SE HAYA CAPTURADO
                    If Format(Trim(alLotexArt(vlIntCont).fechaCaducidad), "DD/MM/YYYY") <> Format(txtFechaCaduce.Text, "DD/MM/YYYY") Then
                        vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                        vlstrResponse = MsgBox("Fecha de caducidad es diferente a la registrada en el lote, ¿Seguro de cambiar?", vlstrStyle, "Mensaje")
                        If Not (vlstrResponse = vbYes) Then
                            txtFechaCaduce.Text = Trim(alLotexArt(vlIntCont).fechaCaducidad)
                        Else
                            alLotexArt(vlIntCont).fechaCaducidad = txtFechaCaduce
                        End If

                    End If
                    '---------------------------------------------
                    If vlstrTipoMovimiento = "UV" Then
                        alLotexArt(vlIntCont).CantidadUV = alLotexArt(vlIntCont).CantidadUV + Val(txtCantRecep.Text)
                    Else
                        alLotexArt(vlIntCont).CantidadUM = alLotexArt(vlIntCont).CantidadUM + Val(txtCantRecep.Text)
                    End If

                    pVaciaArregloGrid
                    vllngTotalAsignados = vllngTotalAsignados + Val(txtCantRecep.Text)
                    txtCantRecep.Text = Val(txtTotalARecibir.Text) - vllngTotalAsignados
                    txtLote.Text = ""
                    txtFechaCaduce = "  /  /    "

                    If Val(txtCantRecep.Text) = 0 Then
                        FraCaptura.Enabled = False
                        cmdGrabarRegistro.Enabled = True
                        cmdGrabarRegistro.SetFocus
                    Else
                        pEnfocaTextBox txtCantRecep
                    End If
                    vlblnExiste = True
                    Exit For
                End If
            Next vlIntCont
        End If

        If Not (vlblnExiste) Then

            vgLngTotalLotesxArt = vgLngTotalLotesxArt + 1
            ReDim Preserve alLotexArt(vgLngTotalLotesxArt)
            alLotexArt(vgLngTotalLotesxArt).Articulo = vlstrChrCveArticulo
            alLotexArt(vgLngTotalLotesxArt).Borrado = ""
            If vlstrTipoMovimiento = "UV" Then
                alLotexArt(vgLngTotalLotesxArt).CantidadUV = Val(txtCantRecep.Text)
                alLotexArt(vgLngTotalLotesxArt).CantidadUM = 0
            Else
                alLotexArt(vgLngTotalLotesxArt).CantidadUM = Val(txtCantRecep.Text)
                alLotexArt(vgLngTotalLotesxArt).CantidadUV = 0
            End If

            alLotexArt(vgLngTotalLotesxArt).CantidadUMInicio = 0
            alLotexArt(vgLngTotalLotesxArt).CantidadUVInicio = 0
            alLotexArt(vgLngTotalLotesxArt).fechaCaducidad = txtFechaCaduce.Text
            alLotexArt(vgLngTotalLotesxArt).lote = txtLote.Text
            alLotexArt(vgLngTotalLotesxArt).Movimiento = vlstrNoMovimiento
            alLotexArt(vgLngTotalLotesxArt).Devolucion = vlstrNoDevolucion
            alLotexArt(vgLngTotalLotesxArt).TablaRelacion = vlstrTablaReferencias
            alLotexArt(vgLngTotalLotesxArt).TipoAccion = vlintTipoAccion
            alLotexArt(vgLngTotalLotesxArt).CodBar = Trim(TxtCBarra.Text)

            pVaciaArregloGrid

            vllngTotalAsignados = vllngTotalAsignados + Val(txtCantRecep.Text)
            txtCantRecep.Text = Val(txtTotalARecibir.Text) - vllngTotalAsignados
            txtLote.Text = ""
            txtFechaCaduce = "  /  /    "

            If Val(txtCantRecep.Text) = 0 Then
                FraCaptura.Enabled = False
                cmdGrabarRegistro.Enabled = True
                cmdGrabarRegistro.SetFocus
            Else
                pEnfocaTextBox txtCantRecep
            End If

        End If

    Else
        If (Val(txtTotalARecibir.Text) - vllngTotalAsignados) < 0 Then
            MsgBox SIHOMsg(452) & " Verifique el total asignado en lotes.", vbCritical, "Mensaje"
        Else
            MsgBox SIHOMsg(36) & " menor igual a " & Str(Val(txtTotalARecibir.Text) - vllngTotalAsignados), vbCritical, "Mensaje"
        End If
        vlblnEstatusAgregar = False
        pEnfocaTextBox txtCantRecep
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregaenGrid"))
End Sub



Private Sub pVaciaArregloGrid(Optional vlstrIniciaCuenta As String)
On Error GoTo NotificaError
Dim vlintTotalSeries As Long

    If vlStrTitUM = "" Then vlStrTitUM = "Mínima"
    If vlStrTitUV = "" Then vlStrTitUV = "Alterna"
    
    'If VSFArticuloLote.Rows > 1 Then
        With VSFArticuloLote
    
            .Clear
            '.Cols = 5
            .Cols = 7
            .Rows = 1
    
            For vlIntCont = 1 To vgLngTotalLotesxArt
                If Trim(alLotexArt(vlIntCont).Borrado) <> "*" Then
                    vlintTotalSeries = vlintTotalSeries + 1
                    .Rows = .Rows + 1
                    .TextMatrix(vlintTotalSeries, 1) = Trim(Str(alLotexArt(vlIntCont).CantidadUV))
                    .TextMatrix(vlintTotalSeries, 2) = Trim(Str(alLotexArt(vlIntCont).CantidadUM))
                    .TextMatrix(vlintTotalSeries, 3) = Trim(alLotexArt(vlIntCont).lote)
                    .TextMatrix(vlintTotalSeries, 4) = Trim(Format(alLotexArt(vlIntCont).fechaCaducidad, "DD/MMM/YYYY"))
                    .TextMatrix(vlintTotalSeries, 5) = Trim(strCveArticulo)
                    .TextMatrix(vlintTotalSeries, 6) = Trim(alLotexArt(vlIntCont).CodBar)
                End If
            Next vlIntCont
    
            .FixedCols = 1
            .FixedRows = 1
            .FixedAlignment(1) = flexAlignCenterCenter
            .FormatString = "|" & Trim(vlStrTitUV) & "|" & Trim(vlStrTitUM) & "|Lote|Fecha caducidad|Cve artículo|Código Barra"
            .ColWidth(0) = 0        'Fix
            .ColWidth(1) = 1200     'CantidadUV
            .ColWidth(2) = IIf(vlintContenidoArt = 1, 0, 1200) 'Cantidadum
            .ColWidth(3) = 1200     'Lote
            '.ColWidth(4) = 1350      'Fecha de caducidad
            .ColWidth(4) = 2295     'Fecha de caducidad
            .ColWidth(5) = 0      'Clave de artículo trazabilidad
            .ColWidth(6) = 0      'Clave de código de barras
    
            .ColAlignment(4) = flexAlignRightBottom
            .ColAlignment(3) = flexAlignLeftBottom
    
        End With
    'End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":VSFArticuloLote_DblClick"))
End Sub

Private Function fblnValidaFechaCaduce() As Boolean
'--------------------------------------------------------------------------
' Función que valida el ingreso de la fecha de caducidad
'--------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlstrMensaje As String

    fblnValidaFechaCaduce = True

    If txtFechaCaduce = "  /  /    " Then
        vlstrMensaje = SIHOMsg("2") & Chr(13) & "Dato: " & txtFechaCaduce.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
        fblnValidaFechaCaduce = False
        txtFechaCaduce.Text = Format(CDate(vldtmfechaNoCaduca), "dd/mm/yyyy")
        Exit Function
    Else
        If Not fblnValidaFecha(txtFechaCaduce) Then
            vlstrMensaje = SIHOMsg("29") & Chr(13) '"!Fecha no válida!, formato de fecha dd/mm/aaaa"
            Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
            fblnValidaFechaCaduce = False
            txtFechaCaduce.Text = Format(CDate(vldtmfechaNoCaduca), "dd/mm/yyyy")
            Exit Function
        End If

        If Year(CDate(txtFechaCaduce.Text)) < 1900 Then
            '¡Fecha no válida!
            MsgBox SIHOMsg(254), vbOKOnly + vbExclamation, "Mensaje"
            fblnValidaFechaCaduce = False
            txtFechaCaduce.Text = Format(CDate(vldtmfechaNoCaduca), "dd/mm/yyyy")
            Exit Function
        End If

        If vlstrFecCadDBO <> "" Then
            If vlstrFecCadDBO <> txtFechaCaduce.Text Then
                If CDate(txtFechaCaduce.Text) <= vldtmfechaServer Then
                    vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                    vlstrResponse = MsgBox(SIHOMsg(1157), vlstrStyle, "Mensaje")
                    ' La fecha de caducidad es menor o igual a la del sistema, ¿Desea continuar?
                    If Not (vlstrResponse = vbYes) Then
                        fblnValidaFechaCaduce = False
                        txtFechaCaduce.Text = vlstrFecCadDBO
                        Exit Function
                    Else
                        vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                        vlstrResponse = MsgBox(SIHOMsg(1158), vlstrStyle, "Mensaje")
                        ' Fecha de caducidad registrada en inventarios es diferente a la capturada, ¿Desea cambiarla?
                        If Not (vlstrResponse = vbYes) Then
                            txtFechaCaduce.Text = vlstrFecCadDBO
                            Exit Function
                        End If
                    End If
                Else
                    vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                    vlstrResponse = MsgBox(SIHOMsg(1158), vlstrStyle, "Mensaje")
                    ' Fecha de caducidad registrada en inventarios es diferente a la capturada, ¿Desea cambiarla?
                    If Not (vlstrResponse = vbYes) Then
                        txtFechaCaduce.Text = vlstrFecCadDBO
                        Exit Function
                    End If
                End If
            Else
                If CDate(txtFechaCaduce.Text) <= vldtmfechaServer Then
                    vlstrStyle = vbYesNo + vbQuestion + vbDefaultButton2
                    vlstrResponse = MsgBox(SIHOMsg(1157), vlstrStyle, "Mensaje")
                    ' La fecha de caducidad es menor o igual a la del sistema, ¿Desea continuar?
                    If Not (vlstrResponse = vbYes) Then
                        fblnValidaFechaCaduce = False
                        txtFechaCaduce.Text = vlstrFecCadDBO
                        Exit Function
                    End If
                End If
            End If
        Else
            If CDate(txtFechaCaduce.Text) <= vldtmfechaServer Then
                Call MsgBox(SIHOMsg(1159), vbExclamation, "Mensaje")
                ' La fecha de caducidad debe ser mayor a la fecha del sistema.
                fblnValidaFechaCaduce = False
                txtFechaCaduce.Text = Format(CDate(vldtmfechaNoCaduca), "dd/mm/yyyy")
                Exit Function
            End If
        End If
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Function
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaFechaCaduce"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Function
        End If
    End If
End Function

Private Sub pLlenaArregloLocal()
On Error GoTo NotificaError

    vglngTotalCantLotexArt = 0
    vglngTotalCantUMLotexArt = 0
    vglngTotalCantUVLotexArt = 0
    vgLngTotalLotesxArt = 0

    For vlIntCont = 1 To vgLngTotalLotesxMov
        If Trim(agLotes(vlIntCont).Articulo) = Trim(txtClaveArticulo.Text) And Trim(agLotes(vlIntCont).Borrado) <> "*" _
            And (Not vlblnValidaMovimiento Or (vlblnValidaMovimiento And agLotes(vlIntCont).Movimiento = vlstrNoMovimiento And agLotes(vlIntCont).Devolucion = vlstrNoDevolucion) Or (vlblnValidaMovimiento And Trim(agLotes(vlIntCont).TablaRelacion) = "EBO" And Trim(vlstrTablaReferencias) <> "ERE")) Then
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
                        alLotexArt(vgLngTotalLotesxArt).Movimiento = agLotes(vlIntCont).Movimiento
                        alLotexArt(vgLngTotalLotesxArt).Devolucion = agLotes(vlIntCont).Devolucion
                        alLotexArt(vgLngTotalLotesxArt).TablaRelacion = agLotes(vlIntCont).TablaRelacion
                        alLotexArt(vgLngTotalLotesxArt).TipoAccion = agLotes(vlIntCont).TipoAccion
                    End If
                Else
                    If Trim(vlstrTablaReferencias) <> "SOS" And Trim(vlstrTablaReferencias) <> "EOE" Then 'Si no es Otras entradas / salidas
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
                        alLotexArt(vgLngTotalLotesxArt).Movimiento = agLotes(vlIntCont).Movimiento
                        alLotexArt(vgLngTotalLotesxArt).Devolucion = agLotes(vlIntCont).Devolucion
                        alLotexArt(vgLngTotalLotesxArt).TablaRelacion = agLotes(vlIntCont).TablaRelacion
                        'alLotexArt(vgLngTotalLotesxArt).TipoAccion = agLotes(vlintCont).TipoAccion
                    Else
                        If (vlstrTipoMovimiento = "UV" And agLotes(vlIntCont).CantidadUV <> 0) Or (vlstrTipoMovimiento = "UM" And agLotes(vlIntCont).CantidadUM <> 0) Then
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
                            alLotexArt(vgLngTotalLotesxArt).Movimiento = agLotes(vlIntCont).Movimiento
                            alLotexArt(vgLngTotalLotesxArt).Devolucion = agLotes(vlIntCont).Devolucion
                            alLotexArt(vgLngTotalLotesxArt).TablaRelacion = agLotes(vlIntCont).TablaRelacion
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

Private Function enviarGet(ByVal Url As String) As String
    Dim http As Object

    enviarGet = ""
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    http.Open "GET", Url, False
    http.send

    If http.Status = 200 Then
        Dim json As Object

        enviarGet = http.responseText
        'MsgBox resultado
    Else
        enviarGet = "[E]"
        'MsgBox "Error en la solicitud: " & http.Status & " - " & http.statusText
    End If

    Set http = Nothing

End Function

Private Function ObtenerLicenciaBoletin() As Boolean

  Dim rs As New ADODB.Recordset
  Dim RfcEmpresa As String
  Dim strRfcEncriptado As String
  Dim LicVchValor As String

  RfcEmpresa = ""
  LicVchValor = ""
  ObtenerLicenciaBoletin = False
  'obtener Rfc Empresa
   Set rs = frsRegresaRs("select NVL(VCHRFC,' ') as VCHRFC FROM CNEMPRESACONTABLE where TNYCLAVEEMPRESA=" & CStr(vgintClaveEmpresaContable))
    If rs.RecordCount <> 0 Then RfcEmpresa = Trim(rs!vchRFC)
    rs.Close

    If RfcEmpresa <> "" Then strRfcEncriptado = fstrEncrypt(RfcEmpresa, "CONTASIHO041099ELECTRO")

  'obtener VCHLICENCIABOLETINES
    Set rs = frsRegresaRs("select NVL(VCHVALOR,' ') as VALOR FROM SIPARAMETRO where VCHNOMBRE = 'VCHLICENCIABOLETINES' AND INTCVEEMPRESACONTABLE = " & CStr(vgintClaveEmpresaContable))
    If rs.RecordCount <> 0 Then LicVchValor = Trim(rs!valor)
    rs.Close

    If RfcEmpresa <> "" And LicVchValor <> "" Then ObtenerLicenciaBoletin = IIf(LicVchValor = strRfcEncriptado, True, False)

End Function

'Se configura la pantalla para la trazabilidad deacuerdo con el bit
Private Sub CfgPantallaTrazabilidad()

    With frmCapturaLotePV
        .Width = 10065
        .Height = IIf(vgblnTrazabilidad = True, 4080, 3600)
    End With

    With frmGrabar
        .Top = IIf(vgblnTrazabilidad = True, 2835, 2350)
        .Left = 4630
    End With

    With FraCaptura
        .Top = IIf(vgblnTrazabilidad = True, 1200, 720)
        .Left = 120
    End With

    With VSFArticuloLote
        .Top = IIf(vgblnTrazabilidad = True, 1695, 1220)
        .Left = 3960
    End With

    With Label2
        .Top = IIf(vgblnTrazabilidad = True, 1320, 870)
        .Left = 4080
    End With

    With txtTotalARecibir
        .Top = IIf(vgblnTrazabilidad = True, 1290, 810)
        .Left = 5520
    End With

    With FraLotes
        .Top = 0
        .Left = 120
        .Width = 9730
        .Height = IIf(vgblnTrazabilidad = True, 1185, 705)
    End With

    With Label1
        .Top = IIf(vgblnTrazabilidad = True, 735, 240)
        .Left = 180
    End With

    With txtDescripcionLgaArt
        .Top = IIf(vgblnTrazabilidad = True, 675, 195)
        .Left = 1200
    End With

    With txtClaveArticulo
        .Top = IIf(vgblnTrazabilidad = True, 675, 195)
        .Left = 1200
    End With

    With Label3
        .Top = IIf(vgblnTrazabilidad = True, 300, -1500)
        .Left = 180
    End With

    With TxtCBarra
        .Top = IIf(vgblnTrazabilidad = True, 240, -1560)
        .Left = 2040
    End With

    With ChkModoCBarra
        .Top = IIf(vgblnTrazabilidad = True, 240, -1560)
        .Left = 6915
    End With
End Sub

Private Function TraerDescripcionArticulo() As String
    Dim vlstrSentencia  As String
    Dim rs As New ADODB.Recordset

    TraerDescripcionArticulo = ""

    vlstrSentencia = "select vchNombreComercial Descripcion from ivArticulo " & _
                             " Where chrCveArticulo = " & _
                             "(Select CAST(CHRCVEARTICULO as CHAR(10)) from ivetiqueta where rtrim(ltrim(INTIDETIQUETA)) = '" & Replace(Trim(TxtCBarra.Text), "'", "''") & "'and rownum < 2)"

    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)

    If rs.RecordCount > 0 Then
        TraerDescripcionArticulo = rs!Descripcion
    End If
End Function
