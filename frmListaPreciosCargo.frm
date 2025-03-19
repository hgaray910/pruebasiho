VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaPreciosCargo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listas de precio por cargos"
   ClientHeight    =   10230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10230
   ScaleMode       =   0  'User
   ScaleWidth      =   13020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Height          =   810
      Left            =   310
      TabIndex        =   44
      Top             =   10320
      Visible         =   0   'False
      Width           =   11385
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   300
         Left            =   45
         TabIndex        =   45
         Top             =   480
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000002&
         Caption         =   "Actualizando listas de precios, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   47
         Top             =   180
         Width           =   11250
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   11325
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Cargando datos, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   90
         TabIndex        =   46
         Top             =   180
         Width           =   11250
      End
   End
   Begin VB.TextBox txtCargoSel 
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "Descripción del cargo seleccionado"
      Top             =   3960
      Width           =   11095
   End
   Begin VB.Frame Frame4 
      Height          =   1935
      Left            =   120
      TabIndex        =   42
      Top             =   1800
      Width           =   12775
      Begin VSFlex7LCtl.VSFlexGrid grdCargos 
         Height          =   1575
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Listado de cargos"
         Top             =   240
         Width           =   12535
         _cx             =   22110
         _cy             =   2778
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
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
         AutoResize      =   0   'False
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame fraFiltros 
      Height          =   1815
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   12775
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   315
         Left            =   11305
         TabIndex        =   8
         ToolTipText     =   "Filtrar cargos"
         Top             =   1305
         Width           =   1300
      End
      Begin VB.ComboBox cboTipoCargo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Tipo de cargo"
         Top             =   240
         Width           =   4925
      End
      Begin VB.ComboBox cboConceptoFacturacion 
         Height          =   315
         Left            =   7745
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Concepto de factura"
         Top             =   600
         Width           =   4910
      End
      Begin VB.ComboBox cboSubFamilia 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Subfamilia"
         Top             =   1305
         Width           =   4925
      End
      Begin VB.ComboBox cboFamilia 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Familia"
         Top             =   960
         Width           =   4925
      End
      Begin VB.ComboBox cboClasificacionSA 
         Height          =   315
         Left            =   7745
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Clasificación"
         Top             =   240
         Width           =   4910
      End
      Begin VB.TextBox txtIniciales 
         Height          =   315
         Left            =   7745
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Descripción del cargo"
         Top             =   960
         Width           =   4885
      End
      Begin VB.TextBox txtClaveArticulo 
         Height          =   315
         Left            =   4615
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Clave del artículo"
         Top             =   600
         Width           =   1485
      End
      Begin VB.ComboBox cboArtMed 
         Height          =   315
         ItemData        =   "frmListaPreciosCargo.frx":0000
         Left            =   1200
         List            =   "frmListaPreciosCargo.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Tipo de artículo"
         Top             =   600
         Width           =   2270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Concepto de factura"
         Height          =   195
         Left            =   6260
         TabIndex        =   41
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblTipoCargo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de cargo"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   300
         Width           =   990
      End
      Begin VB.Label lblTipoArticulo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo artículo"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   660
         Width           =   900
      End
      Begin VB.Label lblSubFamilia 
         AutoSize        =   -1  'True
         Caption         =   "Subfamilia"
         Height          =   195
         Left            =   135
         TabIndex        =   38
         Top             =   1365
         Width           =   720
      End
      Begin VB.Label lblFamilia 
         AutoSize        =   -1  'True
         Caption         =   "Familia"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1020
         Width           =   480
      End
      Begin VB.Label lblClasificacion 
         AutoSize        =   -1  'True
         Caption         =   "Clasificación "
         Height          =   195
         Left            =   6260
         TabIndex        =   36
         Top             =   300
         Width           =   930
      End
      Begin VB.Label lblDescripcionCargo 
         AutoSize        =   -1  'True
         Caption         =   "Descripción cargo"
         Height          =   195
         Left            =   6260
         TabIndex        =   35
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label lblClaveArticulo 
         AutoSize        =   -1  'True
         Caption         =   "Clave artículo"
         Height          =   195
         Left            =   3570
         TabIndex        =   34
         Top             =   660
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   6248
      TabIndex        =   32
      Top             =   9480
      Width           =   615
      Begin VB.CommandButton cmdGrabar 
         Height          =   495
         Left            =   60
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmListaPreciosCargo.frx":0033
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar"
         Top             =   130
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listas de precios"
      Height          =   5015
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   12775
      Begin VB.TextBox txtEditCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ListBox UpDown1 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "frmListaPreciosCargo.frx":0375
         Left            =   240
         List            =   "frmListaPreciosCargo.frx":0382
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VSFlex7LCtl.VSFlexGrid grdPrecios 
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Listas de precios"
         Top             =   240
         Width           =   12535
         _cx             =   22110
         _cy             =   5953
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
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
         AutoResize      =   0   'False
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
         Begin VB.ListBox lstPesos 
            Appearance      =   0  'Flat
            Height          =   420
            ItemData        =   "frmListaPreciosCargo.frx":03C0
            Left            =   5620
            List            =   "frmListaPreciosCargo.frx":03CA
            TabIndex        =   48
            Top             =   500
            Visible         =   0   'False
            Width           =   1000
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Modificar listas"
         Height          =   1185
         Left            =   120
         TabIndex        =   25
         Top             =   3690
         Width           =   10375
         Begin VB.ComboBox cboTipoIncremento 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Tipo de incremento"
            Top             =   480
            Width           =   3655
         End
         Begin VB.TextBox txtMargenUtilidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3880
            TabIndex        =   13
            ToolTipText     =   "Margen de utilidad"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtCostoBase 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5320
            TabIndex        =   14
            ToolTipText     =   "Costo base"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6760
            TabIndex        =   15
            ToolTipText     =   "Precio"
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkUsarTabulador 
            Caption         =   "Usar tabulador"
            Height          =   255
            Left            =   3880
            TabIndex        =   17
            ToolTipText     =   "Usar tabulador"
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdAplicar 
            Caption         =   "Aplicar"
            Height          =   350
            Left            =   8680
            TabIndex        =   18
            ToolTipText     =   "Aplicar cambios a las listas de precios seleccionadas"
            Top             =   705
            Width           =   1560
         End
         Begin VB.CheckBox chkIncrementoAutomatico 
            Caption         =   "Incremento automático"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Incremento automático"
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de incremento"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Margen de utilidad"
            Height          =   195
            Left            =   3880
            TabIndex        =   28
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Costo base"
            Height          =   195
            Left            =   5320
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Precio"
            Height          =   195
            Left            =   6760
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdInvertirSeleccion 
         Caption         =   "Invertir selección"
         Height          =   350
         Left            =   10795
         TabIndex        =   19
         ToolTipText     =   "Invertir selección"
         Top             =   3690
         Width           =   1830
      End
      Begin VB.CommandButton cmdRec 
         Caption         =   "Recalcular"
         Height          =   350
         Left            =   10795
         TabIndex        =   20
         ToolTipText     =   "Recalcular el precio de las listas seleccionadas"
         Top             =   4080
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Lista predeterminada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   10795
         TabIndex        =   31
         Top             =   4470
         Width           =   1830
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Precios no asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   10795
         TabIndex        =   30
         Top             =   4680
         Width           =   1830
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descripción del cargo"
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   4020
      Width           =   1545
   End
End
Attribute VB_Name = "frmListaPreciosCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Index del cbo cboTipoCargo
Const cintIndexTodos = 0
Const cintIndexArticulo = 1
Const cintIndexEstudio = 2
Const cintIndexExamen = 3
Const cintIndexExamenGrupo = 4
Const cintIndexGrupo = 5
Const cintIndexOtro = 6
Const cintIndexPaquete = 7

'Columnas del grid de cargos
Const cintColSelCargo = 0
Const cintColClaveCargo = 1
Const cintColDescCargo = 2
Const cintColDescTipoCargo = 3
Const cintColTipoCargo = 4
Const cintColPrecioEspecifico = 5

'Columnas del grid de precios:
Const cintColSeleccion = 0
Const cintColClaveLista = 1
Const cintColLista = 2
Const cintColIncremetoAutomatico = 3
Const cintColTipoIncremento = 4
Const cintColUtilidad = 5
Const cintColUtilidadSubrogado = 6
Const cintColTabulador = 7
Const cintColCostoBase = 8
Const cintColPrecio = 9
Const cintColTipoMoneda = 10
Const cintColPrecioMaximo = 11
Const cintColCostoMasAlto = 12
Const cintColPrecioUltimaEntrada = 13
Const cintColListaPredeterminada = 14
Const cintColNuevoEnLista = 15
Const cintColModificado = 16

Const clngRojo = &HC0&
Const clngAzul = &HC00000

Const cstrUltimaCompra = "ULTIMA COMPRA"
Const cstrCompraMasAlta = "COMPRA MAS ALTA"
Const cstrPrecioMaximoPublico = "PRECIO MAXIMO AL PUBLICO"

Dim llngMarcados As Long
Dim lstrTipoCargoSel As String
Dim lstrCveCargoSel As String
Dim lblnPermisoCosto As Boolean

Private Sub cboArtMed_Click()
1         On Error GoTo NotificaError
          Dim rs As ADODB.Recordset
          
2         txtIniciales.Text = ""
3         pConfiguraGridCargos
4         lstrTipoCargoSel = ""
5         lstrCveCargoSel = ""
6         txtCargoSel.Text = ""
7         pConfiguraGrid
8         cmdGrabar.Enabled = False
9         llngMarcados = 0
10        pHabilitaModificar
          
11        cboFamilia.Clear
12        If cboArtMed.ListIndex > 0 Then
13            Set rs = frsEjecuta_SP(cboArtMed.ItemData(cboArtMed.ListIndex), "SP_IVSELFAMILIA")
14            pLlenarCboRs cboFamilia, rs, 0, 1
15            rs.Close
16        End If
17        cboFamilia.AddItem "<TODOS>", 0
18        cboFamilia.ItemData(cboFamilia.newIndex) = 0
19        cboFamilia.ListIndex = 0
          
20    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboArtMed_Click" & " Linea:" & Erl()))
End Sub

Private Sub cboClasificacionSA_Click()
    On Error GoTo NotificaError
    
    txtIniciales.Text = ""
    pConfiguraGridCargos
    lstrTipoCargoSel = ""
    lstrCveCargoSel = ""
    txtCargoSel.Text = ""
    pConfiguraGrid
    cmdGrabar.Enabled = False
    llngMarcados = 0
    pHabilitaModificar

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboClasificacionSA_Click"))
End Sub

Private Sub cboConceptoFacturacion_Click()
    On Error GoTo NotificaError

    txtIniciales.Text = ""
    pConfiguraGridCargos
    lstrTipoCargoSel = ""
    lstrCveCargoSel = ""
    txtCargoSel.Text = ""
    pConfiguraGrid
    cmdGrabar.Enabled = False
    llngMarcados = 0
    pHabilitaModificar

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboConceptoFacturacion_Click"))
End Sub

Private Sub cboFamilia_Click()
1      On Error GoTo NotificaError
          Dim rs As ADODB.Recordset
          
2         txtIniciales.Text = ""
3         pConfiguraGridCargos
4         lstrTipoCargoSel = ""
5         lstrCveCargoSel = ""
6         txtCargoSel.Text = ""
7         pConfiguraGrid
8         cmdGrabar.Enabled = False
9         llngMarcados = 0
10        pHabilitaModificar
          
11        cboSubFamilia.Clear
12        If cboFamilia.ListIndex > 0 Then
13            Set rs = frsEjecuta_SP(cboFamilia.ItemData(cboFamilia.ListIndex) & "|" & cboArtMed.ItemData(cboArtMed.ListIndex), "SP_IVSELSUBFAMILIAXFAMILIA")
14            pLlenarCboRs cboSubFamilia, rs, 2, 3
15            rs.Close
16        End If
17        cboSubFamilia.AddItem "<TODOS>", 0
18        cboSubFamilia.ItemData(cboSubFamilia.newIndex) = 0
19        cboSubFamilia.ListIndex = 0
          
20        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFamilia_Click" & " Linea:" & Erl()))
End Sub

Private Sub cboSubFamilia_Click()
    On Error GoTo NotificaError

    txtIniciales.Text = ""
    pConfiguraGridCargos
    lstrTipoCargoSel = ""
    lstrCveCargoSel = ""
    txtCargoSel.Text = ""
    pConfiguraGrid
    cmdGrabar.Enabled = False
    llngMarcados = 0
    pHabilitaModificar

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboSubfamilia_Click"))
End Sub

Private Sub cboTipoCargo_Click()
1     On Error GoTo NotificaError
2         txtIniciales.Text = ""
3         pConfiguraGridCargos
4         lstrTipoCargoSel = ""
5         lstrCveCargoSel = ""
6         txtCargoSel.Text = ""
7         pConfiguraGrid
8         cmdGrabar.Enabled = False
9         llngMarcados = 0
10        pHabilitaModificar
          
11        If cboTipoCargo.Text = "ARTICULOS" Then
12            pCargaCbosArticulo
13        End If
14        If cboTipoCargo.Text = "ESTUDIOS" Then
15            pCargaCboClasificacionSA "IM"
16        End If
17        If cboTipoCargo.Text = "EXAMENES" Or cboTipoCargo.Text = "EXAMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXAMENES" Then
18            pCargaCboClasificacionSA "LA"
19        End If
          
20        lblTipoArticulo.Enabled = cboTipoCargo.Text = "ARTICULOS"
21        cboArtMed.Enabled = cboTipoCargo.Text = "ARTICULOS"
22        txtClaveArticulo.Enabled = cboTipoCargo.Text = "ARTICULOS"
23        If txtClaveArticulo.Enabled = False Then txtClaveArticulo.Text = ""
24        lblClaveArticulo.Enabled = cboTipoCargo.Text = "ARTICULOS"
          
25        lblFamilia.Enabled = cboTipoCargo.Text = "ARTICULOS"
26        cboFamilia.Enabled = cboTipoCargo.Text = "ARTICULOS"
          
27        lblSubFamilia.Enabled = cboTipoCargo.Text = "ARTICULOS"
28        cboSubFamilia.Enabled = cboTipoCargo.Text = "ARTICULOS"
          
29        lblClasificacion.Enabled = cboTipoCargo.Text = "EXAMENES" Or cboTipoCargo.Text = "EXAMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXAMENES" Or cboTipoCargo.Text = "ESTUDIOS"
30        cboClasificacionSA.Enabled = cboTipoCargo.Text = "EXAMENES" Or cboTipoCargo.Text = "EXAMENES Y GRUPOS" Or cboTipoCargo.Text = "GRUPOS DE EXAMENES" Or cboTipoCargo.Text = "ESTUDIOS"
          
31        If cboArtMed.ListCount <> 0 Then
32            cboArtMed.ListIndex = 0
33        End If
          
34    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoCargo_Click" & " Linea:" & Erl()))
End Sub

Private Sub pCargaCboClasificacionSA(strTipo As String)
1         On Error GoTo NotificaError

          Dim rs As ADODB.Recordset
2         cboClasificacionSA.Clear
3         If strTipo = "IM" Then
4             Set rs = frsEjecuta_SP("1|-1|-1", "SP_IMSELCLASIFICACIONESTUDIO")
5             pLlenarCboRs cboClasificacionSA, rs, 0, 1
6             rs.Close
7             cboClasificacionSA.AddItem "<TODOS>", 0
8             cboClasificacionSA.ItemData(cboClasificacionSA.newIndex) = 0
9             cboClasificacionSA.ListIndex = 0
10        End If
11        If strTipo = "LA" Then
12            Set rs = frsEjecuta_SP("", "SP_LASELCLASIFICACIONEXAMEN")
13            pLlenarCboRs cboClasificacionSA, rs, 0, 1
14            rs.Close
15            cboClasificacionSA.AddItem "<TODOS>", 0
16            cboClasificacionSA.ItemData(cboClasificacionSA.newIndex) = 0
17            cboClasificacionSA.ListIndex = 0
18        End If
          
19        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCboClasificacionSA" & " Linea:" & Erl()))
End Sub

Private Sub pCargaCbosArticulo()
    cboArtMed.ListIndex = 0
    cboArtMed_Click
End Sub

Private Sub pConfiguraGridCargos()
On Error GoTo NotificaError
    With grdCargos
        .Clear
        .Cols = 5
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Clave|Descripción cargo|Tipo||"
        .ColWidth(cintColSelCargo) = 150
        .ColWidth(cintColClaveCargo) = 1500
        .ColWidth(cintColDescCargo) = 8020
        .ColWidth(cintColDescTipoCargo) = 2550
        .ColWidth(cintColTipoCargo) = 0
        .ColWidth(cintColPrecioEspecifico) = 0
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .AddItem ""
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridCargos"))
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    With grdPrecios
        .Clear
        .Rows = 1
        .Cols = 16
        .FixedCols = 1
        .FixedRows = 1
        
                        '    00|   01|              02|                   03|             04|             05|            06|        07|    08|09|10|11|12|13|
        .FormatString = "|Clave|Lista de precios|Incremento" & Chr(13) & "automático|Tipo" & Chr(13) & "incremento|Margen" & Chr(13) & "utilidad|Margen" & Chr(13) & "utilidad" & Chr(13) & "subrogado|Usar" & Chr(13) & "tabulador|Costo base|Precio|Moneda||||||"
        .RowHeight(0) = 850
        .ColWidth(cintColSeleccion) = 150               ' seleccion
        .ColWidth(cintColClaveLista) = 600              ' clave de la lista de precios
        .ColWidth(cintColLista) = 4200                  ' lista de precios
        .ColWidth(cintColIncremetoAutomatico) = 900     ' incremento automatico
        .ColWidth(cintColTipoIncremento) = 1150         ' tipo incremento
        .ColWidth(cintColUtilidad) = 1050               ' margen de utilidad
        .ColWidth(cintColUtilidadSubrogado) = 0         ' margen de utilidad subrogado, solo se mostrara si es un articulo o medicamente 1050
        .ColWidth(cintColTabulador) = 800               ' usar tabulador
        .ColWidth(cintColCostoBase) = 1185              ' costo base
        .ColWidth(cintColPrecio) = 1185                 ' precio
        .ColWidth(cintColTipoMoneda) = 1000            'Tipo de moneda (dólares o pesos)
        .ColWidth(cintColPrecioMaximo) = 0              ' precio maximo al publico
        .ColWidth(cintColCostoMasAlto) = 0              ' precio costo mas alto
        .ColWidth(cintColPrecioUltimaEntrada) = 0       ' precio ultima entrada
        .ColWidth(cintColListaPredeterminada) = 0       ' lista predeterminada
        .ColWidth(cintColNuevoEnLista) = 0              ' nuevo en la lista de precios
        .ColWidth(cintColModificado) = 0                ' Indicar si fue modificado el cargo en la lista
                
        .FixedAlignment(cintColSeleccion) = flexAlignCenterCenter
        .FixedAlignment(cintColClaveLista) = flexAlignCenterCenter
        .FixedAlignment(cintColLista) = flexAlignCenterCenter
        .FixedAlignment(cintColIncremetoAutomatico) = flexAlignCenterCenter
        .FixedAlignment(cintColTipoIncremento) = flexAlignCenterCenter
        .FixedAlignment(cintColUtilidad) = flexAlignCenterCenter
        .FixedAlignment(cintColUtilidadSubrogado) = flexAlignCenterCenter
        .FixedAlignment(cintColTabulador) = flexAlignCenterCenter
        .FixedAlignment(cintColCostoBase) = flexAlignCenterCenter
        .FixedAlignment(cintColPrecio) = flexAlignCenterCenter
        .FixedAlignment(cintColTipoMoneda) = flexAlignCenterCenter
    
        .ColAlignment(cintColClaveLista) = flexAlignLeftCenter
        .ColAlignment(cintColUtilidad) = flexAlignRightCenter
        .ColAlignment(cintColUtilidadSubrogado) = flexAlignRightCenter
        .ColAlignment(cintColSeleccion) = flexAlignCenterCenter
        .ColAlignment(cintColLista) = flexAlignLeftCenter
        .ColAlignment(cintColIncremetoAutomatico) = flexAlignCenterCenter
        .ColAlignment(cintColCostoBase) = flexAlignRightCenter
        .ColAlignment(cintColTabulador) = flexAlignCenterCenter
        .ColAlignment(cintColPrecio) = flexAlignRightCenter
        .ColAlignment(cintColTipoMoneda) = flexAlignCenterCenter
        
        .AddItem ""
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub

Private Sub cboTipoIncremento_Click()
    Dim rs As New ADODB.Recordset
    
    If grdPrecios.Rows > 1 Then
        If Trim(grdPrecios.TextMatrix(1, cintColClaveLista)) <> "" Then
            If cboTipoIncremento.ListIndex = 2 Then
                txtCostoBase.Enabled = lblnPermisoCosto
                Set rs = frsRegresaRs("SELECT NVL(MNYPRECIOMAXIMOPUBLICO,0) precio FROM IVARTICULOEMPRESAS WHERE TRIM(CHRCVEARTICULO) = '" & Trim(lstrCveCargoSel) & "' AND TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable)
                If rs.RecordCount <> 0 Then
                    txtCostoBase.Text = Format(IIf(IsNull(rs!precio), 0, rs!precio), "$###,###,###,##0.0000##")
                Else
                    txtCostoBase.Text = Format(0, "$###,###,###,##0.0000##")
                End If
            Else
                txtCostoBase.Enabled = False
                txtCostoBase.Text = Format(IIf(cboTipoIncremento.ListIndex = 0, grdPrecios.TextMatrix(1, cintColPrecioUltimaEntrada), grdPrecios.TextMatrix(1, cintColCostoMasAlto)), "$###,###,###,##0.0000##")
            End If
        End If
    End If
    
End Sub

Private Sub cmdAplicar_Click()
1     On Error GoTo NotificaError
          Dim lngContador As Long
       
2         If lstrTipoCargoSel = "AR" And cboTipoIncremento.ListIndex = 2 Then
3             If MsgBox(SIHOMsg(1220), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
4                 pPoliticaPrecioMaximo (Replace(txtCostoBase.Text, "$", ""))
5             Else
6                 Exit Sub
7             End If
8         End If
       
9         For lngContador = 1 To grdPrecios.Rows - 1
10            If Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*" Then
11                grdPrecios.TextMatrix(lngContador, cintColIncremetoAutomatico) = IIf(chkIncrementoAutomatico.Value = 1, "*", "")
12                grdPrecios.TextMatrix(lngContador, cintColPrecio) = Format(Replace(txtPrecio.Text, "$", ""), "$###,###,###,##0.00####")
13                grdPrecios.TextMatrix(lngContador, cintColUtilidad) = Format(Replace(txtMargenUtilidad.Text, "%", ""), "0.0000") & "%"
                   
14                If lstrTipoCargoSel = "AR" Then
15                    grdPrecios.TextMatrix(lngContador, cintColTabulador) = IIf(chkUsarTabulador.Value = 1, "*", "")
16                    grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = IIf(cboTipoIncremento.ListIndex = 0, "ÚLTIMA", IIf(cboTipoIncremento.ListIndex = 1, "COMPRA", "PRECIO"))
17                    If cboTipoIncremento.ListIndex = 0 Then ' Ultima compra
18                        grdPrecios.TextMatrix(lngContador, cintColCostoBase) = Format(grdPrecios.TextMatrix(lngContador, cintColPrecioUltimaEntrada), "$###,###,###,##0.0000##")
19                    ElseIf cboTipoIncremento.ListIndex = 1 Then ' Compra mas alta
20                        grdPrecios.TextMatrix(lngContador, cintColCostoBase) = Format(grdPrecios.TextMatrix(lngContador, cintColCostoMasAlto), "$###,###,###,##0.0000##")
21                    Else
22                        grdPrecios.TextMatrix(lngContador, cintColCostoBase) = Format(Replace(txtCostoBase.Text, "$", ""), "$###,###,###,##0.0000##")
23                    End If
24                Else
25                    If lblnPermisoCosto Then
26                        grdPrecios.TextMatrix(lngContador, cintColCostoBase) = Format(Replace(txtCostoBase.Text, "$", ""), "$###,###,###,##0.0000##")
27                    End If
28                End If
29                grdPrecios.TextMatrix(lngContador, cintColModificado) = "*"
30            End If
31        Next lngContador
         
32    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAplicar_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdFiltrar_Click()
1     On Error GoTo NotificaError
          Dim strTipoCargo As String
          Dim rs As New ADODB.Recordset
          
2         If cboTipoCargo.ListIndex = -1 Or cboTipoCargo.ListIndex = cintIndexTodos Then
3             strTipoCargo = "*"
4         ElseIf cboTipoCargo.ListIndex = cintIndexArticulo Then
5             strTipoCargo = "AR"
6         ElseIf cboTipoCargo.ListIndex = cintIndexExamen Then
7             strTipoCargo = "EX"
8         ElseIf cboTipoCargo.ListIndex = cintIndexExamenGrupo Then
9             strTipoCargo = "EG"
10        ElseIf cboTipoCargo.ListIndex = cintIndexGrupo Then
11            strTipoCargo = "GE"
12        ElseIf cboTipoCargo.ListIndex = cintIndexEstudio Then
13            strTipoCargo = "ES"
14        ElseIf cboTipoCargo.ListIndex = cintIndexOtro Then
15            strTipoCargo = "OC"
16        ElseIf cboTipoCargo.ListIndex = cintIndexPaquete Then
17            strTipoCargo = "PA"
18        End If
                  
19        pConfiguraGridCargos
20        lstrTipoCargoSel = ""
21        lstrCveCargoSel = ""
22        txtCargoSel.Text = ""
23        pConfiguraGrid
24        cmdGrabar.Enabled = False
25        llngMarcados = 0
26        pHabilitaModificar
          
27        vgstrParametrosSP = Str(vgintNumeroDepartamento) & "|" & strTipoCargo & "|" & cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex)
28        If lblTipoArticulo.Enabled Then
29            vgstrParametrosSP = vgstrParametrosSP & "|" & cboArtMed.ItemData(cboArtMed.ListIndex) & "|" & cboFamilia.ItemData(cboFamilia.ListIndex) & "|" & cboSubFamilia.ItemData(cboSubFamilia.ListIndex)
30        Else
31            vgstrParametrosSP = vgstrParametrosSP & "|3|0|0"
32        End If
33        If lblClasificacion.Enabled Then
34            vgstrParametrosSP = vgstrParametrosSP & "|" & cboClasificacionSA.ItemData(cboClasificacionSA.ListIndex)
35        Else
36            vgstrParametrosSP = vgstrParametrosSP & "|0"
37        End If
38        vgstrParametrosSP = vgstrParametrosSP & "|" & Trim(txtClaveArticulo.Text) & "|" & Trim(txtIniciales.Text)
39        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCARGOS")
40        If rs.RecordCount <> 0 Then
41            grdCargos.Visible = False
42            grdCargos.Rows = 1
43            Do While Not rs.EOF
44                grdCargos.Rows = grdCargos.Rows + 1
45                grdCargos.TextMatrix(grdCargos.Rows - 1, cintColClaveCargo) = rs!clave
46                grdCargos.TextMatrix(grdCargos.Rows - 1, cintColDescCargo) = rs!Descripcion
47                grdCargos.TextMatrix(grdCargos.Rows - 1, cintColDescTipoCargo) = rs!TipoDescripcion
48                grdCargos.TextMatrix(grdCargos.Rows - 1, cintColTipoCargo) = rs!tipo
49                If IsNull(rs!precioespecifico) Then
50                    grdCargos.TextMatrix(grdCargos.Rows - 1, cintColPrecioEspecifico) = "0"
51                Else
52                    grdCargos.TextMatrix(grdCargos.Rows - 1, cintColPrecioEspecifico) = Val(rs!precioespecifico)
53                End If
                  
54                rs.MoveNext
55            Loop
56            grdCargos.Visible = True
57            grdCargos.Row = 1
58        Else
59            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
60        End If
          
61    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdFiltrar_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdGrabar_Click()
1     On Error GoTo NotificaError
          Dim llngPersonaGraba As Long
          Dim lngContador As Long
          Dim llngRow As Long
          Dim rs As New ADODB.Recordset
          
2         If fblnRevisaPermiso(vglngNumeroLogin, 4054, "E") Then
3             llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
4             If llngPersonaGraba <> 0 Then
      '            freBarra.Visible = True
      '            pgbBarra.Value = 0
      '            freBarra.Refresh
                  
5                 EntornoSIHO.ConeccionSIHO.BeginTrans
6                     For lngContador = 1 To grdPrecios.Rows - 1
      '                    If ((lngContador / grdPrecios.Rows) * 100) Mod 2 Then
      '                        pgbBarra.Value = (lngContador / grdPrecios.Rows) * 100
      '                    End If
                              
7                         If Trim(grdPrecios.TextMatrix(lngContador, cintColModificado)) = "*" Then
8                             vgstrParametrosSP = grdPrecios.TextMatrix(lngContador, cintColClaveLista) & "|" & Trim(lstrCveCargoSel) & "|" & Trim(lstrTipoCargoSel) & "|" & CStr(Val(Format(grdPrecios.TextMatrix(lngContador, cintColPrecio), "###########.######"))) & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "PRECIO", "M", IIf(grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "COMPRA", "A", "C")) & "|" & Replace(grdPrecios.TextMatrix(lngContador, cintColUtilidad), "%", "") & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColTabulador) = "*", "1", "0") & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColIncremetoAutomatico) = "*", "1", "0") & "|" & IIf(grdPrecios.TextMatrix(lngContador, cintColTipoMoneda) = "PESOS", "1", "0") & "|" & Replace(grdPrecios.TextMatrix(lngContador, cintColUtilidadSubrogado), "%", "")
9                             If Val(grdPrecios.TextMatrix(lngContador, cintColNuevoEnLista)) = 1 Then
10                                Set rs = frsRegresaRs("select count(*) total from pvdetallelista where pvdetallelista.INTCVELISTA = " & grdPrecios.TextMatrix(lngContador, cintColClaveLista) & _
                                                    " and pvdetallelista.CHRCVECARGO = '" & Trim(lstrCveCargoSel) & "' and pvdetallelista.CHRTIPOCARGO ='" & Trim(lstrTipoCargoSel) & "'")
11                                If rs!Total = 0 Then
12                                    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSDETALLELISTAPRECIO"
13                                Else
14                                    frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDDETALLELISTAPRECIO"
15                                End If
16                            Else
17                                frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDDETALLELISTAPRECIO"
18                            End If
                              
                              'Actualiza costos base
19                            If (grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "PRECIO" Or grdPrecios.TextMatrix(lngContador, cintColTipoIncremento) = "NA") Then
20                                vgstrParametrosSP = vgintClaveEmpresaContable & "|" & Trim(lstrTipoCargoSel) & "|" & Trim(lstrCveCargoSel) & "|" & Format(grdPrecios.TextMatrix(lngContador, cintColCostoBase), "0.0000##") & "|" & IIf(Format(grdPrecios.TextMatrix(lngContador, cintColCostoBase), "0.000000") = "0.000000", "1", "0")
21                                frsEjecuta_SP vgstrParametrosSP, "sp_PVActualizaCosto"
22                            End If
23                        End If
24                    Next lngContador
                  
25                    pGuardarLogTransaccion Me.Name, EnmCambiar, llngPersonaGraba, "LISTA DE PRECIOS POR CARGOS", lstrCveCargoSel
                  
26                EntornoSIHO.ConeccionSIHO.CommitTrans
                  
                  'La información se actualizó satisfactoriamente.
27                MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
                  
28                For llngRow = 1 To grdCargos.Rows - 1
29                    grdCargos.TextMatrix(llngRow, cintColSelCargo) = ""
30                Next llngRow
31                lstrTipoCargoSel = ""
32                lstrCveCargoSel = ""
33                txtCargoSel.Text = ""
34                pConfiguraGrid
35                cmdGrabar.Enabled = False
36                pIniciaModificarListas
37                llngMarcados = 0
38                pHabilitaModificar
                  
39            End If
40        End If
41    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabar_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdInvertirSeleccion_Click()
1         On Error GoTo NotificaError
          Dim lngContador As Long
          
2         If grdPrecios.Rows > 1 Then
3             If Trim(grdPrecios.TextMatrix(1, cintColClaveLista)) <> "" Then
4                 For lngContador = 1 To grdPrecios.Rows - 1
5                     grdPrecios.TextMatrix(lngContador, 0) = IIf(Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*", "", "*")
                      
6                     If Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*" Then
7                         llngMarcados = llngMarcados + 1
8                     Else
9                         llngMarcados = llngMarcados - 1
10                    End If
11                Next lngContador
12                pHabilitaModificar
13            End If
14        End If
          
15        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdInvertirSeleccion_Click" & " Linea:" & Erl()))
End Sub

Private Sub cmdRec_Click()
    On Error GoTo NotificaError

    Dim lngContador As Long

    '¿Está seguro que desea recalcular los precios de la lista?
    If MsgBox(SIHOMsg(990), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
        With grdPrecios
            For lngContador = 1 To .Rows - 1
                If Trim(.TextMatrix(lngContador, 0)) = "*" Then
                    pCalcularPrecio lngContador
                    .TextMatrix(lngContador, 0) = ""
                    .TextMatrix(lngContador, cintColModificado) = "*"
                End If
            Next
        End With
        
        llngMarcados = 0
        pHabilitaModificar
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdRec_Click"))
End Sub

Private Sub pCalcularPrecio(lngRow As Long)
1         On Error GoTo NotificaError
          
          Dim dblPrecio As Double
          Dim dblAumentoTabulador As Double
          Dim rs As ADODB.Recordset
          Dim strParametros As String
          Dim dblCosto As Double
          Dim dblUtilidad As Double
          Dim dblUtilidadSubrogado As Double
          Dim strSQL As String
          
2         If grdPrecios.TextMatrix(lngRow, cintColIncremetoAutomatico) = "*" Then
3             dblCosto = CDbl(grdPrecios.TextMatrix(lngRow, cintColCostoBase))
4             dblUtilidad = CDbl(Replace(grdPrecios.TextMatrix(lngRow, cintColUtilidad), "%", ""))
              dblUtilidadSubrogado = CDbl(Replace(grdPrecios.TextMatrix(lngRow, cintColUtilidadSubrogado), "%", ""))
5             dblAumentoTabulador = 0
6             If grdPrecios.TextMatrix(lngRow, cintColTabulador) = "*" Then
                  'Set rs = frsRegresaRs("SELECT SP_IVSELTABULADORCATARTICULO(" & dblCosto & "," & frmMantenimientoArticulo.lintTipoArticulo & "," & frmMantenimientoArticulo.llngContenidoArticulo & ", " & fintTabuladorListaPrecio(lngCveLista) & ") AUMENTO FROM DUAL")
7                 strSQL = "select sp_IVSelTabulador(" & dblCosto & ", '" & Trim(lstrCveCargoSel) & "', " & fintTabuladorListaPrecio(grdPrecios.TextMatrix(lngRow, cintColClaveLista)) & ") aumento from dual"
8                 Set rs = frsRegresaRs(strSQL)
9                 If Not rs.EOF Then
10                    dblAumentoTabulador = rs!aumento
11                End If
12                rs.Close
13            End If
14            If grdPrecios.ColWidth(cintColUtilidadSubrogado) = 0 Then
                 dblCosto = dblCosto * (1 + (dblUtilidad / 100))
              Else
                 dblCosto = dblCosto * (1 + (dblUtilidadSubrogado / 100))
              End If
15            dblCosto = dblCosto * (1 + (dblAumentoTabulador / 100))
16            grdPrecios.TextMatrix(lngRow, 9) = Format(dblCosto, "$###,###,###,##0.00####")
17        End If
          
18        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalcularPrecio" & " Linea:" & Erl()))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
          Dim llngRow As Long

2         If KeyCode = vbKeyEscape Then
3             If ActiveControl.Name <> "txtEditCol" And ActiveControl.Name <> "UpDown1" And ActiveControl.Name <> "lstPesos" Then
4                 KeyCode = 0
5                 If Trim(txtCargoSel.Text) <> "" Then
6                     If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
7                         lstrTipoCargoSel = ""
8                         lstrCveCargoSel = ""
9                         txtCargoSel.Text = ""
10                        For llngRow = 1 To grdCargos.Rows - 1
11                            grdCargos.TextMatrix(llngRow, cintColSelCargo) = ""
12                        Next llngRow
13                        pConfiguraGrid
14                        cmdGrabar.Enabled = False
15                        pIniciaModificarListas
16                        llngMarcados = 0
17                        pHabilitaModificar
18                    End If
19                ElseIf grdCargos.Rows > 1 Then
20                    If Trim(grdCargos.TextMatrix(1, cintColClaveCargo)) <> "" Then
21                        cboTipoCargo.ListIndex = 0
22                        cboTipoCargo_Click
23                        txtIniciales.Text = ""
24                        If cboClasificacionSA.ListCount > 0 Then cboClasificacionSA.ListIndex = 0
25                        If cboConceptoFacturacion.ListCount > 0 Then cboConceptoFacturacion.ListIndex = 0
26                        cboTipoCargo.SetFocus
27                    Else
28                        Unload Me
29                    End If
30                Else
31                    Unload Me
32                End If
33            End If
34        ElseIf KeyCode = vbKeyReturn Then
35             If ActiveControl.Name <> "grdPrecios" And ActiveControl.Name <> "txtEditCol" And ActiveControl.Name <> "UpDown1" And ActiveControl.Name <> "grdCargos" And ActiveControl.Name <> "lstPesos" Then
36                SendKeys vbTab
37             End If
38        End If

39    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown" & " Linea:" & Erl()))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    lblnPermisoCosto = fblnRevisaPermiso(vglngNumeroLogin, 4054, "C")
    pLlenaCombos
    
    pIniciaModificarListas
    llngMarcados = 0
    pHabilitaModificar
    
    cmdGrabar.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pIniciaModificarListas()
    txtMargenUtilidad.Text = "0.0000%"
    txtPrecio.Text = "$0.00"
    chkIncrementoAutomatico.Value = 0
    chkUsarTabulador.Value = 0
    If cboTipoIncremento.ListCount > 0 Then
        cboTipoIncremento.ListIndex = 0
    End If
    txtCostoBase.Text = ""
End Sub

Private Sub pLlenaCombos()
1     On Error GoTo NotificaError
          Dim rs As New ADODB.Recordset

          'tipo de cargo
2         cboTipoCargo.AddItem "<TODOS>", cintIndexTodos
3         cboTipoCargo.AddItem "ARTICULOS", cintIndexArticulo
4         cboTipoCargo.AddItem "ESTUDIOS", cintIndexEstudio
5         cboTipoCargo.AddItem "EXAMENES", cintIndexExamen
6         cboTipoCargo.AddItem "EXAMENES Y GRUPOS", cintIndexExamenGrupo
7         cboTipoCargo.AddItem "GRUPOS DE EXAMENES", cintIndexGrupo
8         cboTipoCargo.AddItem "OTROS CONCEPTOS", cintIndexOtro
9         cboTipoCargo.AddItem "PAQUETES", cintIndexPaquete
10        cboTipoCargo.ListIndex = 0
          
          'Conceptos de faturación
11        Set rs = frsEjecuta_SP("0|1|-1", "sp_PvSelConceptoFactura")
12        If rs.RecordCount > 0 Then
13            pLlenarCboRs cboConceptoFacturacion, rs, 0, 1, 3
14            cboConceptoFacturacion.ListIndex = 0
15        End If
          
16        cboTipoIncremento.Clear
17        cboTipoIncremento.AddItem "ÚLTIMA COMPRA", 0
18        cboTipoIncremento.AddItem "COMPRA MÁS ALTA", 1
19        cboTipoIncremento.AddItem "PRECIO MÁXIMO AL PÚBLICO", 2
20        cboTipoIncremento.ListIndex = 0
               
21    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaCombos" & " Linea:" & Erl()))
End Sub


Private Sub grdCargos_Click()
On Error GoTo NotificaError
    Dim llngRowSel As Long
    Dim llngRow As Long
    If grdCargos.MouseCol = 0 And grdCargos.MouseRow > 0 Then
        If grdCargos.TextMatrix(grdCargos.Row, cintColClaveCargo) <> "" Then
            grdCargos_KeyDown 13, 0
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCargos_Click"))
End Sub

Private Sub grdCargos_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
          Dim llngRowSel As Long
          Dim llngRow As Long
2         If KeyCode = vbKeyReturn And grdCargos.Row > 0 Then
3             If grdCargos.TextMatrix(grdCargos.Row, cintColClaveCargo) <> "" Then
4                 lstrTipoCargoSel = ""
5                 lstrCveCargoSel = ""
6                 If grdCargos.TextMatrix(grdCargos.Row, cintColSelCargo) = "" Then
7                     llngRowSel = grdCargos.Row
8                     For llngRow = 1 To grdCargos.Rows - 1
9                         grdCargos.TextMatrix(llngRow, cintColSelCargo) = ""
10                    Next llngRow
11                    grdCargos.TextMatrix(llngRowSel, cintColSelCargo) = "*"
12                    txtCargoSel.Text = Trim(grdCargos.TextMatrix(llngRowSel, cintColDescCargo))
13                    lstrTipoCargoSel = Trim(grdCargos.TextMatrix(llngRowSel, cintColTipoCargo))
14                    lstrCveCargoSel = Trim(grdCargos.TextMatrix(llngRowSel, cintColClaveCargo))
15                    pConfiguraGrid
16                    pLlenaGridListas llngRowSel
17                    pIniciaModificarListas
18                    llngMarcados = 0
19                    pHabilitaModificar
20                Else
21                    grdCargos.TextMatrix(grdCargos.Row, cintColSelCargo) = ""
22                    txtCargoSel.Text = ""
23                    pConfiguraGrid
24                    cmdGrabar.Enabled = False
25                    llngMarcados = 0
26                    pHabilitaModificar
27                End If
28            End If
29        End If
30    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdCargos_KeyDown" & " Linea:" & Erl()))
End Sub

Private Sub pLlenaGridListas(lngRow As Long)
1     On Error GoTo NotificaError
          Dim rs As New ADODB.Recordset
          Dim rsSubrogado As New ADODB.Recordset 'para saber si el Articulo/Medicamento es subrogado
          Dim objIntGrid As Integer
          Dim objIntColsGrid As Integer
          Dim strSentencia As String
          Dim rsParametroMargen As New ADODB.Recordset
          Dim vlstrBitCapturaMargen As Boolean
          
2         vgstrParametrosSP = "'" & Trim(grdCargos.TextMatrix(lngRow, cintColClaveCargo)) & "'|'" & grdCargos.TextMatrix(lngRow, cintColTipoCargo) & "'|" & vgintNumeroDepartamento
3         Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelListasPrecios")
               
4         If rs.RecordCount = 0 Then ' no hay listas de precios dadas de alta o activas
             '¡No se encontraron listas de precios configuradas o activas!
5            MsgBox SIHOMsg(1218), vbExclamation + vbOKOnly, "Mensaje"
6         Else
              If grdCargos.TextMatrix(lngRow, cintColTipoCargo) = "AR" Then
                
                Set rsParametroMargen = frsRegresaRs("SELECT BITCAPTURAMARGENSUBROGADO FROM PvParametro")
                If rsParametroMargen.RecordCount > 0 Then
                    vlstrBitCapturaMargen = rsParametroMargen!BITCAPTURAMARGENSUBROGADO
                End If
                rsParametroMargen.Close
                If vlstrBitCapturaMargen Then
              
              
                    strSentencia = "Select count(*)total from ivarticulo " & _
                                   "inner join IVArticulosSubrogados on ivarticulo.INTIDARTICULO = IVArticulosSubrogados.INTIDARTICULO " & _
                                   "where ivarticulo.CHRCVEARTICULO = " & Trim(grdCargos.TextMatrix(lngRow, cintColClaveCargo))
                    Set rsSubrogado = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
                    If rsSubrogado!Total > 0 Then
                       grdPrecios.ColWidth(cintColUtilidadSubrogado) = 1050
                    Else
                        grdPrecios.ColWidth(cintColUtilidadSubrogado) = 0
                    End If
                End If
              Else
                grdPrecios.ColWidth(cintColUtilidadSubrogado) = 0
              End If
7             grdPrecios.Visible = False
8             grdPrecios.Rows = 1
9             Do While Not rs.EOF
10                grdPrecios.Rows = grdPrecios.Rows + 1
11                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColClaveLista) = rs!claveLista
12                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColLista) = Trim(rs!Descripcion)
13                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColIncremetoAutomatico) = IIf(rs!bitIncremento = 1, "*", "")
14                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoIncremento) = IIf(lstrTipoCargoSel = "AR", rs!tipoIncremento, "NA")
15                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidad) = Format(rs!margenUtilidad, "0.0000") & "%"
                  grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColUtilidadSubrogado) = Format(rs!margenUtilidadSubrogado, "0.0000") & "%"
16                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTabulador) = IIf(rs!bitTabulador = 1, "*", "")
17                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoBase) = Format(rs!costo, "$###,###,###,##0.0000##")
18                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecio) = Format(rs!precio, "$###,###,###,##0.00####")
19                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColTipoMoneda) = rs!TipoMoneda
20                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioMaximo) = rs!precioMaximo
21                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColCostoMasAlto) = rs!CostoMasAlto
22                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColPrecioUltimaEntrada) = rs!costoUltimaEntrada
23                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColListaPredeterminada) = rs!bitPredeterminada
24                grdPrecios.TextMatrix(grdPrecios.Rows - 1, cintColNuevoEnLista) = rs!nuevo
25                rs.MoveNext
26            Loop
              
27            For objIntGrid = 1 To grdPrecios.Rows - 1
28                If Trim(grdPrecios.TextMatrix(objIntGrid, cintColNuevoEnLista)) = "1" Then
29                    For objIntColsGrid = 3 To grdPrecios.Cols - 1
30                       grdPrecios.Col = objIntColsGrid
31                       grdPrecios.Row = objIntGrid
32                       grdPrecios.CellForeColor = &HC0&
33                       grdPrecios.CellFontBold = True
34                    Next objIntColsGrid
35                End If
36                If Trim(grdPrecios.TextMatrix(objIntGrid, cintColListaPredeterminada)) = "1" Then
37                    grdPrecios.Col = 1
38                    grdPrecios.Row = objIntGrid
39                    grdPrecios.CellForeColor = &HC00000
40                    grdPrecios.CellFontBold = True
41                    grdPrecios.Col = 2
42                    grdPrecios.Row = objIntGrid
43                    grdPrecios.CellForeColor = &HC00000
44                    grdPrecios.CellFontBold = True
45                End If
46            Next objIntGrid
47            grdPrecios.Visible = True
48            grdPrecios.Row = 1
49            cmdGrabar.Enabled = True
50        End If
          
51    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaGridListas" & " Linea:" & Erl()))
End Sub

Private Sub grdPrecios_Click()
    On Error GoTo NotificaError
                
1    If grdPrecios.MouseRow <> 0 And Trim(grdPrecios.TextMatrix(1, 1)) <> "" Then
2        If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClaveLista)) <> "" Then
3            If grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidad Then
4                 If Val(grdCargos.TextMatrix(grdCargos.Row, cintColPrecioEspecifico)) <> 1 Then
5                      pEditarColumna 32, txtEditCol, grdPrecios.Row, grdPrecios.Col
6                 Else
                     '¡No se pueden modificar precios relacionados con cargos de paquetes!
7                      MsgBox SIHOMsg(1592) & ":" & Chr(13) & Trim(grdCargos.TextMatrix(grdCargos.Row, cintColDescCargo)), vbOKOnly + vbInformation, "Mensaje"
8                      Exit Sub
9                 End If
              ElseIf grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidadSubrogado Then
                  If Val(grdCargos.TextMatrix(grdCargos.Row, cintColPrecioEspecifico)) <> 1 Then
                      pEditarColumna 32, txtEditCol, grdPrecios.Row, grdPrecios.Col
                 Else
                     '¡No se pueden modificar precios relacionados con cargos de paquetes!
                      MsgBox SIHOMsg(1592) & ":" & Chr(13) & Trim(grdCargos.TextMatrix(grdCargos.Row, cintColDescCargo)), vbOKOnly + vbInformation, "Mensaje"
                      Exit Sub
                 End If
10            ElseIf grdPrecios.Col = cintColTipoIncremento And lstrTipoCargoSel = "AR" Then
11                pPonerUpDown grdPrecios.Row
12            ElseIf grdPrecios.Col = cintColCostoBase Then
                'SOLO SE PUEDE MODIFICAR SI SE TIENE EL PERMISO Y SI SE TRATA DE POLITICA DE PRECIO MAXIMO AL PUBLICO
13                If lblnPermisoCosto And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO MAXIMO AL PUBLICO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA") Then
14                    pEditarColumna 32, txtEditCol, grdPrecios.Row, grdPrecios.Col
15                End If
16            ElseIf grdPrecios.Col = cintColTipoMoneda And lstrTipoCargoSel = "PA" Then
17                pMostrarlstPesos grdPrecios
18            ElseIf grdPrecios.Col = cintColTabulador And lstrTipoCargoSel = "AR" Then
19                grdPrecios.TextMatrix(grdPrecios.Row, cintColTabulador) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColTabulador)) = "*", "", "*")
20                grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
21            ElseIf grdPrecios.Col = cintColIncremetoAutomatico Then
22                grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColIncremetoAutomatico)) = "*", "", "*")
23                grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
24            End If
                   
25            If grdPrecios.MouseCol = cintColSeleccion Then
26                grdPrecios.TextMatrix(grdPrecios.Row, cintColSeleccion) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColSeleccion)) = "*", "", "*")
27                If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColSeleccion)) = "*" Then
28                    llngMarcados = llngMarcados + 1
29                Else
30                    llngMarcados = llngMarcados - 1
31                End If
32                pHabilitaModificar
33            End If
34        End If
35    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_Click" & " Linea:" & Erl()))
End Sub

Private Sub pPonerUpDown(llngRow As Long)
1     On Error GoTo NotificaError
          Dim intIndex As Integer
          
2         If lstrTipoCargoSel = "AR" Then
3             UpDown1.ListIndex = -1
4             For intIndex = 0 To UpDown1.ListCount - 1
5                 If UpDown1.List(intIndex) = IIf(grdPrecios.TextMatrix(llngRow, cintColTipoIncremento) = "ÚLTIMA", "ÚLTIMA COMPRA", IIf(grdPrecios.TextMatrix(llngRow, cintColTipoIncremento) = "COMPRA", "COMPRA MÁS ALTA", "PRECIO MÁXIMO AL PÚBLICO")) Then
6                     UpDown1.ListIndex = intIndex
7                     Exit For
8                 End If
9             Next
10            With grdPrecios
11                UpDown1.Move .Left + .CellLeft, .Top + .CellTop, UpDown1.Width, UpDown1.Height
12            End With
13            UpDown1.Visible = True
14            UpDown1.SetFocus
15        End If
          
16    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPonerUpDown" & " Linea:" & Erl()))
End Sub

Private Sub pMostrarlstPesos(grd As VSFlexGrid)
    With grd
        lstPesos.Move .CellLeft, .CellTop, lstPesos.Width, lstPesos.Height
        
        If .TextMatrix(.Row, cintColTipoMoneda) = "PESOS" Then
            lstPesos.ListIndex = 1
        Else
            lstPesos.ListIndex = 0
    End If
    End With
    
    lstPesos.Visible = True
    lstPesos.SetFocus
End Sub

Private Sub grdPrecios_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
          
2         If grdPrecios.Row > 0 Then
3             If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClaveLista)) <> "" Then
4                 If grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidad Or grdPrecios.Col = cintColUtilidadSubrogado Or (lblnPermisoCosto And grdPrecios.Col = cintColCostoBase And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA")) Then
5                     If KeyCode = vbKeyF2 And grdPrecios.Row <> 0 Then pEditarColumna 13, txtEditCol, grdPrecios.Row, grdPrecios.Col
6                 ElseIf grdPrecios.Col = cintColTipoIncremento Then
7                 ElseIf grdPrecios.Col = cintColTipoMoneda Then
8                 Else
9                     If KeyCode = vbKeyReturn Then
10                        If grdPrecios.Row - 1 < grdPrecios.Rows Then
11                            If grdPrecios.Row = grdPrecios.Rows - 1 Then
12                                grdPrecios.Row = 1
13                            Else
14                                grdPrecios.Row = grdPrecios.Row + 1
15                            End If
16                        End If
17                    End If
18                End If
19            End If
20        End If
21    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyDown" & " Linea:" & Erl()))
End Sub

Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, lngRow As Long, lngCol As Long)
1         On Error GoTo NotificaError
          Dim vlintTexto As Integer

2         With txtEdit
3             .Text = Replace(grdPrecios.TextMatrix(lngRow, lngCol), "%", "") 'Inicialización del Textbox
4             Select Case KeyAscii
                  Case 0 To 32
                      'Edita el texto de la celda en la que está posicionado
5                     .SelStart = 0
6                     .SelLength = 1000
7                 Case 8, 48 To 57
                      ' Reemplaza el texto actual solo si se teclean números
8                     vlintTexto = Chr(KeyAscii)
9                     .Text = vlintTexto
10                    .SelStart = 1
11                Case 46
                      ' Reemplaza el texto actual solo si se teclean números
12                    .Text = "."
13                    .SelStart = 1
14            End Select
15        End With
                  
          ' Muestra el textbox en el lugar indicado
16        With grdPrecios
17            If .CellWidth < 0 Then Exit Sub
18            txtEdit.Move .Left + .CellLeft, .Top + .CellTop + 30, .CellWidth, .CellHeight - 10
19        End With
          
20        txtEdit.Visible = True
21        txtEdit.SetFocus
          
22        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna" & " Linea:" & Erl()))
End Sub

Private Sub grdPrecios_KeyPress(KeyAscii As Integer)
1         On Error GoTo NotificaError

2         If grdPrecios.MouseRow <> 0 And grdPrecios.Row > 0 Then
3             If Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintColClaveLista)) <> "" Then
4                 If grdPrecios.Col = cintColPrecio Or grdPrecios.Col = cintColUtilidad Or grdPrecios.Col = cintColUtilidadSubrogado Or (grdPrecios.Col = cintColCostoBase And lblnPermisoCosto And (grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO" Or grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "NA")) Then
5                     pEditarColumna KeyAscii, txtEditCol, grdPrecios.Row, grdPrecios.Col
6                 ElseIf grdPrecios.Col = cintColIncremetoAutomatico Or grdPrecios.Col = cintColTabulador Then
7                     If KeyAscii = 32 Then grdPrecios_Click
8                 ElseIf grdPrecios.Col = cintColTipoIncremento Then
9                     If KeyAscii = 13 Then pPonerUpDown grdPrecios.Row
10                ElseIf grdPrecios.Col = cintColTipoMoneda And lstrTipoCargoSel = "PA" Then
11                    If KeyAscii = 13 Then pMostrarlstPesos grdPrecios
12                Else
13                    Exit Sub
14                End If
15            End If
16        End If
          
17        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyPress" & " Linea:" & Erl()))
End Sub

Private Sub lstPesos_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError
2     If KeyAscii = 13 Then
3             lstPesos_MouseUp 0, 0, 0, 0
4             If grdPrecios.Row < grdPrecios.Rows - 1 Then
5                 grdPrecios.Row = grdPrecios.Row + 1
6             Else
7                 grdPrecios.SetFocus
8             End If
9         End If
10        If KeyAscii = 27 Then
11            grdPrecios.SetFocus
12            lstPesos.Visible = False
13        End If
14    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_KeyPress" & " Linea:" & Erl()))
End Sub

Private Sub lstPesos_LostFocus()
On Error GoTo NotificaError
    lstPesos.Visible = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_LostFocus"))
End Sub

Private Sub lstPesos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo NotificaError
    grdPrecios.Text = lstPesos.Text
    grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
    grdPrecios.SetFocus
    lstPesos.Visible = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_MouseUp"))
End Sub

Private Sub lstPesos_Validate(Cancel As Boolean)
On Error GoTo NotificaError
    lstPesos_MouseUp 0, 0, 0, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstPesos_Validate"))
End Sub

Private Sub txtClaveArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Else
        If KeyAscii = 46 Then
            KeyAscii = 7
        Else
            pValidaNumero KeyAscii
        End If
    End If
End Sub

Private Sub txtCostoBase_GotFocus()
    Replace txtCostoBase, "$", ""
    pSelTextBox txtCostoBase
End Sub

Private Sub txtCostoBase_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 6, txtCostoBase) Then KeyAscii = 0
End Sub

Private Sub txtCostoBase_LostFocus()
    If txtCostoBase.Text <> "" Then
        txtCostoBase.Text = Format(txtCostoBase.Text, "$###,###,###,##0.0000##")
    Else
        txtCostoBase.Text = "$0.0000"
    End If
End Sub

Private Sub txtEditCol_KeyDown(KeyCode As Integer, Shift As Integer)
1      On Error GoTo NotificaError
          'Para verificar que tecla fue presionada en el textbox
2         With grdPrecios
3             Select Case KeyCode
                  Case 27   'ESC
4                     .SetFocus
5                     txtEditCol.Visible = False
6                     KeyCode = 0
                      'vlblnEscTxtEditCOl = True
7                 Case 38   'Flecha para arriba
8                     .SetFocus
9                     If grdPrecios.Col = cintColPrecio Then
10                        Call pSetCellValueCol(grdPrecios, txtEditCol)
11                    End If
12                    If grdPrecios.Col = cintColUtilidad Then
13                        Call pSetCellValueCol(grdPrecios, txtEditCol)
14                    End If
                      If grdPrecios.Col = cintColUtilidadSubrogado Then
                          Call pSetCellValueCol(grdPrecios, txtEditCol)
                      End If
15                    If grdPrecios.Col = cintColCostoBase Then
16                        Call pSetCellValueCol(grdPrecios, txtEditCol)
17                    End If
18                    DoEvents
19                    If .Row > .FixedRows Then
                          'vgblnNoEditar = True
20                        .Row = .Row - 1
                          'vgblnNoEditar = False
21                    End If
                      'vlblnEscTxtEditCOl = False
22                Case 40, 13
23                    .SetFocus
24                    If grdPrecios.Col = cintColPrecio Then
25                        Call pSetCellValueCol(grdPrecios, txtEditCol)
26                    End If
27                    If grdPrecios.Col = cintColUtilidad Then
28                        Call pSetCellValueCol(grdPrecios, txtEditCol)
29                    End If
                      If grdPrecios.Col = cintColUtilidadSubrogado Then
                          Call pSetCellValueCol(grdPrecios, txtEditCol)
                      End If
30                    If grdPrecios.Col = cintColCostoBase Then
31                        Call pSetCellValueCol(grdPrecios, txtEditCol)
32                    End If
33                    DoEvents
34                    If .Row < .Rows - 1 Then
                          'vgblnNoEditar = True
35                        .Row = .Row + 1
                          'vgblnNoEditar = False
36                    End If
                      'vlblnEscTxtEditCOl = False
37            End Select
38        End With
39    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtEditCol_KeyDown" & " Linea:" & Erl()))
End Sub

Private Sub pSetCellValueCol(grid As VSFlexGrid, txtEdit As TextBox)
1         On Error GoTo NotificaError

          ' NOTA:
          ' Este código debe ser  llamado cada vez que el grid pierde el foco y su contenido puede cambiar.
          ' De otra manera, el nuevo valor de la celda se perdería.
          
2         If grid.Col = cintColPrecio Then
3             If txtEdit.Visible Then
4                 If txtEdit.Text <> "" Then
5                     If IsNumeric(txtEdit.Text) Then
6                         grid.Text = Format(txtEdit.Text, "$###,###,###,##0.00####")
7                         grid.TextMatrix(grid.Row, cintColModificado) = "*"
8                     End If
9                 End If
10                txtEdit.Visible = False
11            End If
12        End If
          
13        If grid.Col = cintColCostoBase Then
14            If txtEdit.Visible Then
15                If txtEdit.Text <> "" Then
16                    If IsNumeric(txtEdit.Text) Then
17                        If lstrTipoCargoSel = "AR" And Trim(grid.TextMatrix(grid.Row, cintColTipoIncremento)) = "PRECIO" Then
18                            If MsgBox(SIHOMsg(1220), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
19                                grid.Text = Format(txtEdit.Text, "$###,###,###,##0.0000##")
20                                grid.TextMatrix(grid.Row, cintColModificado) = "*"
21                                pPoliticaPrecioMaximo (txtEditCol.Text)
22                            Else
23                                txtEdit.Visible = False
24                            End If
25                        Else
26                            grid.Text = Format(txtEdit.Text, "$###,###,###,##0.0000##")
27                            grid.TextMatrix(grid.Row, cintColModificado) = "*"
28                        End If
29                    End If
30                End If
31                txtEdit.Visible = False
32            End If
33        End If
          
34        If grid.Col = cintColUtilidad Then
35            If txtEdit.Visible Then
36                If txtEdit.Text <> "" Then
37                    txtEdit.Text = Replace(txtEdit.Text, "%", "")
38                    If IsNumeric(txtEdit.Text) Then
39                        grid.Text = Format(txtEdit.Text, "0.0000") & "%"
40                        grid.TextMatrix(grid.Row, cintColModificado) = "*"
41                    End If
42                End If
43                txtEdit.Visible = False
44            End If
45        End If
          If grid.Col = cintColUtilidadSubrogado Then
             If txtEdit.Visible Then
                If txtEdit.Text <> "" Then
                    txtEdit.Text = Replace(txtEdit.Text, "%", "")
                    If IsNumeric(txtEdit.Text) Then
                        grid.Text = Format(txtEdit.Text, "0.0000") & "%"
                        grid.TextMatrix(grid.Row, cintColModificado) = "*"
                    End If
                End If
                txtEdit.Visible = False
              End If
           End If
          
46        Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol" & " Linea:" & Erl()))
End Sub

Private Sub txtEditCol_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    Dim bytNumDecimales As Byte
    
    If grdPrecios.Col = cintColUtilidad Or grdPrecios.Col = cintColUtilidadSubrogado Or grdPrecios.Col = cintColCostoBase Then
        bytNumDecimales = 6
    Else
        bytNumDecimales = 6 ' precio
    End If
    If Not fblnFormatoCantidad(txtEditCol, KeyAscii, bytNumDecimales) Then
        KeyAscii = 7
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtEditCol_KeyPress"))
End Sub

Private Sub txtEditCol_LostFocus()
    On Error GoTo NotificaError
    
    If txtEditCol.Visible Then
        txtEditCol.Visible = False
    End If
'    If grdPrecios.Col = 8 Then
'        grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
'    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtEditCol_LostFocus"))
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub pHabilitaModificar()
    On Error GoTo NotificaError

    cboTipoIncremento.Enabled = llngMarcados <> 0 And lstrTipoCargoSel = "AR"
    txtCostoBase.Enabled = llngMarcados <> 0 And lblnPermisoCosto And ((lstrTipoCargoSel = "AR" And cboTipoIncremento.ListIndex = 2) Or lstrTipoCargoSel <> "AR")
    txtMargenUtilidad.Enabled = llngMarcados <> 0
    txtPrecio.Enabled = llngMarcados <> 0
    chkIncrementoAutomatico.Enabled = llngMarcados <> 0
    chkUsarTabulador.Enabled = llngMarcados <> 0 And lstrTipoCargoSel = "AR"
    cmdAplicar.Enabled = llngMarcados <> 0
    cmdRec.Enabled = llngMarcados <> 0
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaModificar"))
End Sub

Private Sub txtMargenUtilidad_GotFocus()
    txtMargenUtilidad.Text = Replace(txtMargenUtilidad.Text, "%", "")
    pSelTextBox txtMargenUtilidad
End Sub

Private Sub txtMargenUtilidad_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 4, txtMargenUtilidad) Then KeyAscii = 0
End Sub

Private Function fValidaCantidad(vlintCaracter As Integer, vlintDecimales As Integer, CajaText As TextBox) As Boolean ' procedimiento para validar la cantidad que se introduce a un textbox
    Dim vlintPosicionCursor As Integer
    Dim vlintPosiciones As Integer
    Dim vlintPosicionPunto As Integer
    Dim vlintNumeroDecimales As Integer
    
    fValidaCantidad = True
    If Not IsNumeric(Chr(vlintCaracter)) Then 'no es numero
        If Not vlintCaracter = vbKeyBack Then 'no es retroceso
            If Not vlintCaracter = vbKeyReturn Then 'no es Enter
                If Not vlintCaracter = 46 Then ' no es el punto
                    fValidaCantidad = False ' se anula, estos son los unicos caracteres que se pueden ingresar al texbox
                Else 'es un punto debemos veriricar si se tiene un punto ya en el text
                    If fblnValidaPunto(CajaText.Text) Then ' ya hay un punto
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    Else ' se intenta ingresar un caracter numerico, revisar decimales, revisar si se tiene seleccionado el textbox
        If CajaText.SelText <> CajaText.Text Then
            vlintPosicionCursor = CajaText.SelStart
            vlintPosicionPunto = InStr(1, CajaText.Text, ".")
            If vlintPosicionPunto > 0 Then ' si hay punto
                If vlintPosicionCursor > vlintPosicionPunto Then ' si la poscion es mayor entonces debemos de revisar los decimales
                    'contamos la cantidad de decimales
                    For vlintPosiciones = vlintPosicionPunto + 1 To Len(CajaText.Text)
                    vlintNumeroDecimales = vlintNumeroDecimales + 1
                    Next vlintPosiciones
                    'si ya son tantos como vlinDecimales entonces no permite la insercion
                    If vlintNumeroDecimales >= vlintDecimales Then
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub txtMargenUtilidad_LostFocus()
    If txtMargenUtilidad.Text <> "" Then
        txtMargenUtilidad.Text = Format(txtMargenUtilidad.Text, "0.0000") & "%"
    Else
        txtMargenUtilidad.Text = "0.0000%"
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    txtPrecio.Text = Replace(txtPrecio.Text, "$", "")
    pSelTextBox txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 6, txtPrecio) Then KeyAscii = 0
End Sub

Private Sub txtPrecio_LostFocus()
    If txtPrecio.Text <> "" Then
        txtPrecio.Text = Format(txtPrecio.Text, "$###,###,###,##0.00####")
    Else
        txtPrecio.Text = "$0.00"
    End If
End Sub

Private Sub UpDown1_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 13 Then
        UpDown1_MouseUp 0, 0, 0, 0
        If grdPrecios.Row < grdPrecios.Rows - 1 Then
            grdPrecios.Row = grdPrecios.Row + 1
        End If
    End If
    If KeyAscii = 27 Then
        grdPrecios.SetFocus
        UpDown1.Visible = False
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_KeyPress"))
End Sub

Private Sub UpDown1_LostFocus()
On Error GoTo NotificaError
    UpDown1.Visible = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_LostFocus"))
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         On Error GoTo NotificaError
          
2         grdPrecios.TextMatrix(grdPrecios.Row, cintColModificado) = "*"
3         If UpDown1.Text = "PRECIO MÁXIMO AL PÚBLICO" And grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) <> "PRECIO" Then
4             If MsgBox(SIHOMsg(1220), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
5                 pPoliticaPrecioMaximo (Replace(FormatCurrency(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioMaximo), cintColTipoIncremento), "$", ""))
6             Else
7                 grdPrecios.SetFocus
8                 UpDown1.Visible = False
9                 Exit Sub
10            End If
11        End If
             
12        Select Case UpDown1.Text
              Case "ÚLTIMA COMPRA"
13                grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoBase) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioUltimaEntrada), "$###,###,###,##0.0000##")
14                grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "ÚLTIMA"
15            Case "COMPRA MÁS ALTA"
16                grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoBase) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoMasAlto), "$###,###,###,##0.0000##")
17                grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "COMPRA"
18            Case "PRECIO MÁXIMO AL PÚBLICO"
19                grdPrecios.TextMatrix(grdPrecios.Row, cintColCostoBase) = Format(grdPrecios.TextMatrix(grdPrecios.Row, cintColPrecioMaximo), "$###,###,###,##0.0000##")
20                grdPrecios.TextMatrix(grdPrecios.Row, cintColTipoIncremento) = "PRECIO"
21        End Select
22        grdPrecios.SetFocus
23        UpDown1.Visible = False
          
24    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_MouseUp" & " Linea:" & Erl()))
End Sub

Private Sub pPoliticaPrecioMaximo(vlstrPrecio As String)
    On Error GoTo NotificaError

    Dim objIntGrid As Integer
    
    For objIntGrid = 1 To grdPrecios.Rows - 1
        If grdPrecios.TextMatrix(objIntGrid, cintColTipoIncremento) = "PRECIO" Then
            grdPrecios.TextMatrix(objIntGrid, cintColCostoBase) = Format(vlstrPrecio, "$###,###,###,##0.0000##")
            grdPrecios.TextMatrix(objIntGrid, cintColModificado) = "*"
        End If
    Next objIntGrid
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPoliticaPrecioMaximo"))
End Sub

Private Sub UpDown1_Validate(Cancel As Boolean)
    On Error GoTo NotificaError
    UpDown1_MouseUp 0, 0, 0, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_Validate"))
End Sub
