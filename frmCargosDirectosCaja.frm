VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargosDirectosCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos directos"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   Icon            =   "frmCargosDirectosCaja.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1255
      Left            =   45
      TabIndex        =   56
      Top             =   8160
      Width           =   1730
      Begin VB.CommandButton cmdSeleccionarTodo 
         Caption         =   "Seleccionar todos"
         Height          =   300
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Seleccionar todos los cargos"
         Top             =   210
         Width           =   1455
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   300
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Seleccionar los cargos"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdQuitarTodo 
         Caption         =   "Quitar todos"
         Height          =   300
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Remover la selección en todos los cargos"
         Top             =   520
         Width           =   1455
      End
   End
   Begin VB.PictureBox PB 
      Height          =   135
      Left            =   6960
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   43
      Top             =   8220
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Frame FreAplicar 
      Enabled         =   0   'False
      Height          =   1485
      Left            =   45
      TabIndex        =   39
      Top             =   2340
      Width           =   5280
      Begin VB.CommandButton cmdImprimirEstado 
         Enabled         =   0   'False
         Height          =   495
         Left            =   553
         Picture         =   "frmCargosDirectosCaja.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Imprimir el estado de cuenta"
         Top             =   495
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   2680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Departamento"
         Top             =   1060
         Width           =   2535
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2680
         MaxLength       =   5
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Cantidad a cargar"
         Top             =   195
         Width           =   1100
      End
      Begin VB.Frame Frame2 
         Height          =   1275
         Left            =   1575
         TabIndex        =   40
         Top             =   100
         Width           =   30
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Aplicar cargo"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3990
         TabIndex        =   6
         Top             =   210
         Width           =   1215
      End
      Begin VB.CheckBox chkMedicamentoAplicado 
         Caption         =   "Medicamento aplicado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2680
         TabIndex        =   7
         Top             =   695
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.Label lblEstadoCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Estado de cuenta"
         Height          =   195
         Left            =   150
         TabIndex        =   52
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label lblDepartamento 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   1640
         TabIndex        =   44
         Top             =   1120
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1640
         TabIndex        =   41
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Descripción completa del cargo "
      Height          =   900
      Left            =   1845
      TabIndex        =   33
      Top             =   8520
      Width           =   9690
      Begin VB.TextBox txtNombreComercial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   570
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   240
         Width           =   9435
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1335
      Left            =   1680
      TabIndex        =   26
      Top             =   9480
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   360
         Left            =   165
         TabIndex        =   27
         Top             =   675
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
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
         Height          =   240
         Left            =   75
         TabIndex        =   28
         Top             =   135
         Width           =   7875
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   285
         Left            =   30
         Top             =   120
         Width           =   8145
      End
   End
   Begin VB.Frame FrePaciente 
      Height          =   2265
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   5280
      Begin VB.TextBox txtCuarto 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1855
         Width           =   3600
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   15
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton OptTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   0
         Left            =   3105
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   615
         Width           =   3600
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   0
         Top             =   270
         Width           =   1300
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   500
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1300
         Width           =   3600
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   960
         Width           =   3600
      End
      Begin VB.Label Label8 
         Caption         =   "Cuarto"
         Height          =   255
         Left            =   150
         TabIndex        =   30
         Top             =   1855
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Número de cuenta"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   640
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa"
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   1290
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de paciente"
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   965
         Width           =   1335
      End
   End
   Begin VB.Frame FreDetalle 
      Caption         =   "Cargos del paciente"
      Enabled         =   0   'False
      Height          =   4290
      Left            =   45
      TabIndex        =   17
      Top             =   3885
      Width           =   11490
      Begin VB.TextBox txtBusqueda 
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Descripción del cargo a buscar"
         Top             =   270
         Width           =   3400
      End
      Begin MSMask.MaskEdBox MskFecha 
         Height          =   300
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         Appearance      =   0
         BackColor       =   12640511
         MaxLength       =   14
         Format          =   "dd/mmm/yyyy HH:mm"
         Mask            =   "##/##/## ##:##"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   32
         Top             =   750
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.CheckBox chkIncluyeFacturados 
         Caption         =   "Incluir facturados"
         Height          =   195
         Left            =   2925
         TabIndex        =   31
         Top             =   15
         Width           =   1635
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargos 
         Height          =   3555
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Cargos en la cuenta del paciente"
         Top             =   600
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   6271
         _Version        =   393216
         Cols            =   8
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FormatString    =   "|Descricpion|Precio|Cantidad|Subtotal|Descuento|Monto|Tipo"
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   5
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label13 
         Caption         =   "<Supr> - Para eliminar los cargos seleccionados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7200
         TabIndex        =   55
         Top             =   315
         Width           =   4095
      End
      Begin VB.Label Label11 
         Caption         =   "Descripción del cargo"
         Height          =   255
         Left            =   150
         TabIndex        =   53
         Top             =   315
         Width           =   2055
      End
   End
   Begin VB.Frame FreElementos 
      Caption         =   "Elementos a incluir"
      Enabled         =   0   'False
      Height          =   3780
      Left            =   5400
      TabIndex        =   16
      Top             =   45
      Width           =   6135
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   240
         Left            =   960
         TabIndex        =   47
         Top             =   200
         Width           =   4995
         Begin VB.OptionButton optElementos 
            Caption         =   "Todos"
            Height          =   255
            Index           =   3
            Left            =   3330
            TabIndex        =   51
            Top             =   50
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optElementos 
            Caption         =   "Otros"
            Height          =   195
            Index           =   2
            Left            =   2540
            TabIndex        =   50
            Top             =   50
            Width           =   1050
         End
         Begin VB.OptionButton optElementos 
            Caption         =   "Artículos"
            Height          =   195
            Index           =   1
            Left            =   1500
            TabIndex        =   49
            Top             =   50
            Width           =   1020
         End
         Begin VB.OptionButton optElementos 
            Caption         =   "Medicamentos"
            Height          =   195
            Index           =   0
            Left            =   50
            TabIndex        =   48
            Top             =   50
            Width           =   1440
         End
      End
      Begin VB.OptionButton optClaveDescripcion 
         Caption         =   "Clave"
         Height          =   195
         Index           =   1
         Left            =   4755
         TabIndex        =   4
         Top             =   740
         Width           =   1140
      End
      Begin VB.OptionButton optClaveDescripcion 
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   4755
         TabIndex        =   3
         Top             =   490
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.TextBox txtBuscaElemento 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   4515
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdElementos 
         Height          =   2720
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Selección del elemento a cargar"
         Top             =   945
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4789
         _Version        =   393216
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         SelectionMode   =   1
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Tipo cargo"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   220
         Width           =   765
      End
   End
   Begin VB.Label Label12 
      Caption         =   "<supr> - Para eliminar los cargos seleccionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   54
      Top             =   4200
      Width           =   6495
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Precio modificado"
      Height          =   195
      Left            =   8565
      TabIndex        =   38
      Top             =   8220
      Width           =   1260
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0080C0FF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   8265
      Top             =   8220
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7320
      TabIndex        =   37
      Top             =   8220
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   7245
      Top             =   8220
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Left            =   9915
      Top             =   8220
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha modificada"
      Height          =   195
      Left            =   10215
      TabIndex        =   36
      Top             =   8235
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Excluido"
      Height          =   195
      Left            =   7560
      TabIndex        =   35
      Top             =   8220
      Width           =   600
   End
End
Attribute VB_Name = "frmCargosDirectosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmCargosDirectos                                      -
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar los cargos directos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 5/Feb/2001
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : 08/Mar/2001
'| Fecha última modificación: 30/Sep/2002
'-------------------------------------------------------------------------------------
'| Fecha última modificación: 07/Oct/2003   (Contraseña para cambiar el precio del cargo)
'-------------------------------------------------------------------------------------
'| Fecha última modificación: 25/Nov/2003   Permitir el acceso a la cuenta del paciente
'| solo cuando tiene estado de "abierta" (AdAdmision.bitCuentaCerrada = 0 RegistroExterno.bitCuentaCerrada = 0)
'-------------------------------------------------------------------------------------
'| Fecha última modificación: 16/Ene/2004   Se muestren solamente artículos de costo o ambos IvArticulo.chrCostoGasto = 'C' or
'| IvArticulo.chrCostoGasto = 'A'
'-------------------------------------------------------------------------------------


Option Explicit

Public vllngNumeroOpcion As Long
Public llngNumOpcionHabilitaCambioFecha As Long
Public llngNumOpcionHabilitaMedicamentoAplicado As Long
Public vlblnLlamadoPCE As Boolean

Dim vgbitParametros As Boolean
Dim vgintEmpresa As Integer
Dim vgintTipoPaciente As Integer
Dim vgstrEstadoManto As String
Dim vgblnNoEditarPagos As Boolean
Dim vgblnNoEditarFecha As Boolean
Dim vgintColumnaCurrency As Integer 'Para la columna que se va a editar
Dim vgblnEditaPago As Boolean 'Para saber si se esta editando una cantidad
Dim vllngPersonaGraba As Long
Dim vlblnEditando As Boolean
Dim lblnExcluir As Boolean
Dim lstrCodigo As String
Dim vlblnHabilitaCambioFecha As Boolean
Dim lblnSeleccionarDepto As Boolean

Private Type InformacionHonorarioMedico
    intCveMedico        As Integer ' Clave del médico al que se le pagará el honorario
    strProcedimiento    As String  ' Procedimiento que realizará médico
    dblImporteCargo     As Double  ' Importe con el que se realizará el cargo
    dblImporteHonorario As Double  ' Importe con el que se generará la cuenta por pagar para el médico
End Type

Dim inmInfoHonorario As InformacionHonorarioMedico

Dim vlblnTransaccionenCurso As Boolean
Dim vllngMarcados As Long
Dim Modo As Boolean
Dim vlblnMantenerAutorizado As Boolean
Dim vlintUsuarioAutorizado As Long
Dim vlblnMensajeAviso As Boolean
Dim vlblnBotonNo As Boolean
Dim vlblnavisonoexistente As Boolean

Private Sub pConfiguraGridCargos()
    On Error GoTo NotificaError

    With grdCargos
        .Cols = 16
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Tipo|Descripción del cargo|Cantidad|Precio|Fecha/Hora|Tipo|Referencia|Concepto de facturación|Departamento|Factura"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 400  'Tipo de cargo
        .ColWidth(2) = 4000 'Descripción
        .ColWidth(3) = 700  'Cantidad
        .ColWidth(4) = 1430 'Precio
        .ColWidth(5) = 1630 'Fecha
        .ColWidth(6) = 1200 'Tipo de documento
        .ColWidth(7) = 1000 'Numero de documento
        .ColWidth(8) = 4000 'Concepto de facturación
        .ColWidth(9) = 4000 'Departamento
        .ColWidth(10) = 800 'Factura
        .ColWidth(11) = 0   'bitDescuentaInventario
        .ColWidth(12) = 0   'CveConceptoFacturacion
        .ColWidth(13) = 0   'Clave del Cargo
        .ColWidth(14) = 0   'Estatus de aplicado, no aplicado
        .ColWidth(15) = 0   'Para el ordenamiento por fechas
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignLeftBottom
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignLeftBottom
        .ColAlignment(9) = flexAlignLeftBottom
        .ColAlignment(10) = flexAlignLeftBottom
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    Unload Me
End Sub

Private Sub pCargaElementos()
    On Error GoTo NotificaError

    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim rsElementos As New ADODB.Recordset
    Dim vlstrSentenciaOtroConcepto As String
    Dim rsOtroConcepto As New ADODB.Recordset
    Dim vlsmicvedepartamentoOtroConcepto As Integer
    Dim llngExcluido As Long
    Dim rsSelCargosExcluidos As ADODB.Recordset
    Dim vlstrTipoCargo As String
    
    '------------------------
    ' Limpieza del grid
    '------------------------
    grdElementos.Clear
    grdElementos.Rows = 2
    grdElementos.Cols = 0
    '------------------------
    ' Configurar el grid
    '------------------------
    pConfiguraGrid
    grdElementos.RowData(1) = -1
    
    
    '--------------------------------------
    ' Obtiene el departamento del usuario
    '--------------------------------------
    vlsmicvedepartamentoOtroConcepto = 0
    vlstrSentenciaOtroConcepto = "select smicvedepartamento from login where intNumeroLogin = " & vglngNumeroLogin
    Set rsOtroConcepto = frsRegresaRs(vlstrSentenciaOtroConcepto, adLockOptimistic, adOpenDynamic)
    If rsOtroConcepto.RecordCount > 0 Then
        vlsmicvedepartamentoOtroConcepto = rsOtroConcepto!smicvedepartamento
    End If
    
    If optElementos(3).Value Then
        vlstrTipoCargo = "TO"
    ElseIf optElementos(0).Value Or optElementos(1).Value Then
        vlstrTipoCargo = "AR"
    ElseIf optElementos(2).Value Then
        vlstrTipoCargo = "OC"
    End If
    
    vgstrParametrosSP = CStr(Val(txtMovimientoPaciente.Text)) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vlstrTipoCargo
        
    ' Sp que regresa cargos excluidos
    
    Set rsSelCargosExcluidos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCARGOSEXCLUIDOS")
    
    
    '------------------------
    ' Armar la Inscrucción
    '------------------------
    If Trim(txtBuscaElemento.Text) <> "" Then
        vlstrSentencia = ""
        '-----------------------------------
        ' ARTICULOS
        '-----------------------------------
        If optElementos(0).Value Or optElementos(3).Value Or optElementos(1).Value Then 'Medicamentos o Todos o Articulos
            vlstrSentencia = "select ivArticulo.intIDArticulo Clave, ivArticulo.vchNombreComercial Descripcion, " & _
                "'AR' Tipo, ivArticulo.intContenido Contenido, 1 ExistenciaUM, 1 ExistenciaUV, " & _
                " ivarticulo.chrCveArticulo chrArticulo " & _
                " from ivarticulo " & _
                " Where rtrim(ltrim(chrCostoGasto)) <> 'G' and  rtrim(ltrim(vchEstatus)) = 'ACTIVO' " & _
                " and IVArticulo." & IIf(optClaveDescripcion(0).Value, "vchNombreComercial", "chrCveArticulo") & " Like '" & fstrParseo(Trim(txtBuscaElemento.Text)) & "%'" & _
                IIf(optElementos(0).Value, " and ivArticulo.chrCveArtMedicamen = '1' ", IIf(optElementos(1).Value, " and ivArticulo.chrCveArtMedicamen = '0' ", "")) & _
                " And  (0 = (select count(*) From CcCargoExcluido " & _
                                 " inner join CcCargoExcluidoArticulo on CcCargoExcluidoArticulo.INTCONSECUTIVO = CcCargoExcluido.INTCONSECUTIVO " & _
                                 "where CcCargoExcluido.INTCVEEMPRESA = " & vgintEmpresa & " and ivArticulo.intIDArticulo = CcCargoExcluidoArticulo.INTIDARTICULO  " & _
                                 " and CcCargoExcluido.INTAUTORIZACION = 2)) "
        End If
        
        If optElementos(3).Value Then 'Todos
            vlstrSentencia = vlstrSentencia + "    Union "
        End If
        
        '-----------------------------------
        ' OTROS CONCEPTOS
        '-----------------------------------
        If optElementos(2).Value Or optElementos(3).Value Then 'Otros Conceptos o Todos
           vlstrSentencia = vlstrSentencia + "select intCveConcepto Clave, chrDescripcion Descripcion, 'OC' Tipo, " & _
                        " 1 Contenido, 1 ExistenciaUM, 1 ExistenciaUV, '' chrArticulo " & _
                        "from PvOtroConcepto " & _
                        "     left join PvOtroConceptoDepto on PvOtroConcepto.intcveconcepto = PvOtroConceptoDepto.intcveotroconcepto " & _
                        " where bitEstatus = 1 " & _
                        " and " & IIf(optClaveDescripcion(0).Value, "chrDescripcion", "rtrim(ltrim(CAST(intCveConcepto as CHAR(30))))") & " Like '" & fstrParseo(Trim(txtBuscaElemento.Text)) & "%'" & _
                        " and PvOtroConceptoDepto.smicvedepartamento = " & vlsmicvedepartamentoOtroConcepto & _
                        " and  (0 = (select count(*) From CcCargoExcluido " & _
                                 " inner join CcCargoExcluidoOtroConcepto on CcCargoExcluidoOtroConcepto.INTCONSECUTIVO = CcCargoExcluido.INTCONSECUTIVO " & _
                                 "where CcCargoExcluido.INTCVEEMPRESA = " & vgintEmpresa & " and pvOtroConcepto.intCveConcepto = CcCargoExcluidoOtroConcepto.INTCVEOTROCONCEPTO  " & _
                                 " and CcCargoExcluido.INTAUTORIZACION = 2)) "
        End If
        
        If optElementos(3).Value Then 'Todos
            vlstrSentencia = vlstrSentencia + "    Union "
        End If
        
        '-----------------------------------
        ' EXAMENES
        '-----------------------------------
        If optElementos(3).Value Then 'Todos
            vlstrSentencia = vlstrSentencia + "select intCveExamen Clave, chrNombre Descripcion, 'EX' Tipo, " & _
                        " 1 Contenido, 1 ExistenciaUM, 1 ExistenciaUV, '' chrArticulo " & _
                        "from LaExamen " & _
                        " where bitEstatusActivo = 1 and bitCaracteristica = 0" & _
                        " and " & IIf(optClaveDescripcion(0).Value, "chrNombre", "rtrim(ltrim(CAST(intCveExamen as char(30))))") & " Like '" & fstrParseo(Trim(txtBuscaElemento.Text)) & "%'" & _
                        " And  (0 = (select count(*) From CcCargoExcluido " & _
                                 " inner join CcCargoExcluidoExamen on CcCargoExcluidoExamen.INTCONSECUTIVO = CcCargoExcluido.INTCONSECUTIVO " & _
                                 "where CcCargoExcluido.INTCVEEMPRESA = " & vgintEmpresa & " and LaExamen.INTCVEEXAMEN = CcCargoExcluidoExamen.INTCVEEXAMEN  " & _
                                 " and CcCargoExcluido.INTAUTORIZACION = 2)) "
            
            vlstrSentencia = vlstrSentencia + "    Union "
        
        '-----------------------------------
        ' GRUPO DE EXAMENES
        '-----------------------------------
            vlstrSentencia = vlstrSentencia + "select intCveGrupo Clave, chrNombre Descripcion, 'GE' Tipo, " & _
                        " 1 Contenido, 1 ExistenciaUM, 1 ExistenciaUV, '' chrArticulo " & _
                        "from LaGrupoExamen " & _
                        " where bitEstatusActivo = 1 " & _
                        " and " & IIf(optClaveDescripcion(0).Value, "chrNombre", "rtrim(ltrim(CAST(intCveGrupo as char (30))))") & " Like '" & fstrParseo(Trim(txtBuscaElemento.Text)) & "%'" & _
                        " And  (0 = (select count(*) From CcCargoExcluido " & _
                                 " inner join CcCargoExcluidoGrupoExamen on CcCargoExcluidoGrupoExamen.INTCONSECUTIVO = CcCargoExcluido.INTCONSECUTIVO " & _
                                 "where CcCargoExcluido.INTCVEEMPRESA = " & vgintEmpresa & " and LaGrupoExamen.INTCVEGRUPO = CcCargoExcluidoGrupoExamen.INTCVEGRUPO  " & _
                                 " and CcCargoExcluido.INTAUTORIZACION = 2)) "
            
            vlstrSentencia = vlstrSentencia + "    Union "
        
        '-----------------------------------
        ' ESTUDIOS
        '-----------------------------------
            vlstrSentencia = vlstrSentencia + "select intCveEstudio Clave, vchNombre Descripcion, 'ES' Tipo, " & _
                        " 1 Contenido, 1 ExistenciaUM, 1 ExistenciaUV, '' chrArticulo " & _
                        "from imEstudio " & _
                        " where bitStatusActivo = 1 " & _
                        " and " & IIf(optClaveDescripcion(0).Value, "vchNombre", "rtrim(ltrim(CAST(intCveEstudio as char(30))))") & " Like '" & fstrParseo(Trim(txtBuscaElemento.Text)) & "%'" & _
                        " And  (0 = (select count(*) From CcCargoExcluido " & _
                                 " inner join CcCargoExcluidoEstudio on CcCargoExcluidoEstudio.INTCONSECUTIVO = CcCargoExcluido.INTCONSECUTIVO " & _
                                 "where CcCargoExcluido.INTCVEEMPRESA = " & vgintEmpresa & " and imEstudio.INTCVEESTUDIO = CcCargoExcluidoEstudio.INTCVEESTUDIO " & _
                                 " and CcCargoExcluido.INTAUTORIZACION = 2)) "
        
        End If
        
        vlstrSentencia = vlstrSentencia & " order by tipo, descripcion "
       
        Set rsElementos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        
        With grdElementos
            .Redraw = False
            .Visible = False
                           
            Do While Not rsElementos.EOF
            
                If .RowData(1) <> -1 Then
                     .Rows = .Rows + 1
                     .Row = .Rows - 1
                 End If
                .RowData(.Row) = rsElementos!clave
                .TextMatrix(.Row, 1) = rsElementos!Descripcion
                .TextMatrix(.Row, 2) = rsElementos!tipo
                .TextMatrix(.Row, 3) = rsElementos!Contenido
                .TextMatrix(.Row, 4) = rsElementos!ExistenciaUM
                .TextMatrix(.Row, 5) = rsElementos!ExistenciaUV
                .TextMatrix(.Row, 6) = IIf(IsNull(rsElementos!chrArticulo), "", rsElementos!chrArticulo)
                If rsElementos!ExistenciaUM = 0 And rsElementos!ExistenciaUV = 0 Then
                    For vlintContador = 1 To 2
                        .Col = vlintContador
                        .CellForeColor = &H8000000B
                    Next
                End If

                Do While Not rsSelCargosExcluidos.EOF

                    If rsSelCargosExcluidos!clave = .RowData(.Row) And rsSelCargosExcluidos!TipoCargo = rsElementos!tipo Then
                        For vlintContador = 1 To 2
                            .Col = vlintContador
                            .CellBackColor = &H99FFFF
                        Next
                        rsSelCargosExcluidos.MoveLast
                    End If

                    rsSelCargosExcluidos.MoveNext

                Loop

                If rsSelCargosExcluidos.RecordCount > 0 Then rsSelCargosExcluidos.MoveFirst
                                
                rsElementos.MoveNext
            Loop
             
            .Col = 1
            .Redraw = True
            .Visible = True
            .Row = 1
            
        End With
        
        rsSelCargosExcluidos.Close
        
        rsElementos.Close
        
    End If
    If grdElementos.RowData(1) = -1 Then 'Significa que esta vacia
        grdElementos.Enabled = False
    Else
        grdElementos.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaElementos"))
    Unload Me
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError

    With grdElementos
        .Cols = 7
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción del cargo|Tipo"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 4900 'Descripción del cargo
        .ColWidth(2) = 600  'Tipo
        .ColWidth(3) = 0    'Contenido
        .ColWidth(4) = 0    'ExistenciaUM
        .ColWidth(5) = 0    'EsistenciaUV
        .ColWidth(6) = 0    'Cve articulo en Char
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarVertical
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
    Unload Me
End Sub
Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    On Error GoTo NotificaError
    
    Dim vlbytColumnas As Byte
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        .Clear
'        For vlbytColumnas = 1 To .Cols - 1
'            .TextMatrix(1, vlbytColumnas) = ""
'            .Col = vlbytColumnas
'            .BackColor = &H80000005
'            .ForeColor = &H80000008
'        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
    Unload Me
End Sub

Private Sub pCancelar()
    On Error GoTo NotificaError
    
    FreElementos.Enabled = False
    FreDetalle.Enabled = False
    FreAplicar.Enabled = False
    FrePaciente.Enabled = True
    txtPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    txtCuarto.Text = ""
    txtBuscaElemento.Text = ""
    vgstrEstadoManto = ""
    txtMovimientoPaciente.Locked = False
    OptTipoPaciente(0).Enabled = True
    OptTipoPaciente(1).Enabled = True
    cmdImprimirEstado.Enabled = False
    lblEstadoCuenta.Enabled = False
    Label11.Enabled = False
    txtBusqueda.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pcancelar"))
    Unload Me
End Sub


Private Sub chkIncluyeFacturados_Click()
    On Error GoTo NotificaError
    
    pLlenaCargos

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkIncluyeFacturados_Click"))
    Unload Me
End Sub
Private Sub chkMedicamentoAplicado_Click()
    On Error GoTo NotificaError
    
    'If grdElementos.TextMatrix(grdElementos.Row, 2) <> "AR" Then
    '    chkMedicamentoAplicado.Value = 0
    'End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkMedicamentoAplicado_Click"))
    Unload Me
End Sub

Private Sub chkMedicamentoAplicado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(cboDepartamento) Then cboDepartamento.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------
' Esta función verifica si un cargo pertenece al concepto de facturación configurado para los honorarios médicos
' en los Parámetros del módulo de caja.
'-----------------------------------------------------------------------------------------------------------------
Private Function fblnCargoPerteneceConceptoHonorarios(vlngCveCargo As Long) As Boolean
    Dim vlngConceptoFacturacionHonorario As Long
    Dim vlngConceptoFacturacionCargo As Long
    Dim rs As New ADODB.Recordset

    
    fblnCargoPerteneceConceptoHonorarios = False
    vlngConceptoFacturacionHonorario = fintDameConceptoFacturacionHonorario
    '| Valida si existe configurado un concepto de facturación para los honorarios médicos y se está cargando un "Otro Concepto"
    If vlngConceptoFacturacionHonorario <> -1 Then
        '| Identifica el concepto de facturación del cargo
        Set rs = frsRegresaRs("select PvOtroConcepto.SMICONCEPTOFACT From PvOtroConcepto Where PvOtroConcepto.INTCVECONCEPTO = " & vlngCveCargo)
        If Not rs.EOF Then vlngConceptoFacturacionCargo = rs!SMICONCEPTOFACT
        'Si el cargo que se está realizando es el mismo que está configurado para los honorarios médicos
        If vlngConceptoFacturacionCargo = vlngConceptoFacturacionHonorario Then
            fblnCargoPerteneceConceptoHonorarios = True
        End If
    End If
End Function



Private Sub pActualizaPrecioCargo(vlngCveCargo As Long, vlngPersonaGraba As Long, vlngRow As Long)
    Dim strSentencia As String
    Dim vldblTotDescuento As Double
    Dim vldblIVA As Double
    Dim vlaryParametrosSalida() As String
    Dim rs As New ADODB.Recordset
    Dim intCantidadCargo As Integer
    
    
    
    strSentencia = "Select PvCargo.MNYCANTIDAD from PvCargo Where PvCargo.INTNUMCARGO = " & vlngCveCargo
    Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenStatic)
    If rs.RecordCount <> 0 Then
        intCantidadCargo = rs!MNYCantidad
    End If
    
    strSentencia = "SELECT PVBASEHONORARIOMEDICO.* FROM PVBASEHONORARIOMEDICO INNER JOIN PVCARGO ON (PVCARGO.INTNUMCARGO = PVBASEHONORARIOMEDICO.INTNUMCARGO) WHERE PVBASEHONORARIOMEDICO.INTNUMCARGO = " & vlngCveCargo
    Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenStatic)
    If rs.RecordCount <> 0 Then
        strSentencia = "Update PVBASEHONORARIOMEDICO " & _
                         " Set INTCVEMEDICO = " & inmInfoHonorario.intCveMedico & _
                             " , NUMIMPORTEAFACTURAR = " & inmInfoHonorario.dblImporteCargo * intCantidadCargo & _
                             " , NUMIMPORTEHONORARIO = " & inmInfoHonorario.dblImporteHonorario * intCantidadCargo & _
                             " , VCHPROCEDIMIENTO = '" & Trim(inmInfoHonorario.strProcedimiento) & "'" & _
                       " Where INTNUMCARGO = " & vlngCveCargo
        pEjecutaSentencia strSentencia
    Else
        '------------------------------------------------------------
        '|  Inserta la información del honorario médico
        '------------------------------------------------------------
        strSentencia = "Insert Into PVBASEHONORARIOMEDICO ( INTNUMCARGO " & _
                                                        " , INTCVECUENTA " & _
                                                        " , INTCVEMEDICO " & _
                                                        " , NUMIMPORTEAFACTURAR " & _
                                                        " , NUMIMPORTEHONORARIO " & _
                                                        " , VCHPROCEDIMIENTO) " & _
                                                 " Values ( " & vlngCveCargo & _
                                                        " , " & txtMovimientoPaciente.Text & _
                                                        " , " & inmInfoHonorario.intCveMedico & _
                                                        " , " & inmInfoHonorario.dblImporteCargo * intCantidadCargo & _
                                                        " , " & inmInfoHonorario.dblImporteHonorario * intCantidadCargo & _
                                                        " , '" & Trim(inmInfoHonorario.strProcedimiento) & "')"
        pEjecutaSentencia strSentencia
    End If
    '---------------------------------------------------------------------------
    '|  Actualiza el precio, descuento e IVA del cargo por el honorario
    '---------------------------------------------------------------------------
    
    grdCargos.TextMatrix(vlngRow, 4) = FormatCurrency(inmInfoHonorario.dblImporteCargo, 2)
    
    '-----------------------
    'Descuentos
    '-----------------------
    vldblTotDescuento = 0
    'vlchrTipoDescuento = " "
    
    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
    frsEjecuta_SP IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintTipoPaciente & "|" & _
                    vgintEmpresa & "|" & CLng(Val(txtMovimientoPaciente.Text)) & "|" & _
                    grdCargos.TextMatrix(vlngRow, 1) & "|" & _
                    CLng(Val(grdCargos.TextMatrix(vlngRow, 13))) & "|" & _
                    Val(Format(inmInfoHonorario.dblImporteCargo, "###########0.00")) & "|" & _
                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                    grdCargos.TextMatrix(vlngRow, 14) & "|" & _
                    1 & "|" & _
                    Format(grdCargos.TextMatrix(vlngRow, 3), "") & "|" & _
                    grdCargos.TextMatrix(vlngRow, 11), _
                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
    pObtieneValores vlaryParametrosSalida, vldblTotDescuento
    '-----------------------
    'IVA
    '-----------------------
    strSentencia = "Select smyIva/100 IVA from pvConceptoFacturacion " & _
                    " where smiCveConcepto = " & Trim(grdCargos.TextMatrix(vlngRow, 12))
    Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    vldblIVA = rs!IVA * _
            (Val(Format(grdCargos.TextMatrix(vlngRow, 4), "############.##")) * _
            Val(Format(grdCargos.TextMatrix(vlngRow, 3), "############.##")) - _
            vldblTotDescuento)
    rs.Close
    '-----------------------------------------------------------------------
    'Actualiza el precio capturado en la información del honorario
    '-----------------------------------------------------------------------
    strSentencia = "Update pvCargo set mnyPrecio = " & Format(inmInfoHonorario.dblImporteCargo, "###########0.00") & _
                    ", mnyIVA = " & Trim(str(Round(vldblIVA, 2))) & _
                    ", mnyDescuento = " & Trim(str(Round(vldblTotDescuento, 6))) & _
                    ", bitPrecioManual = 1 " & _
                    ", intEmpleado = " & str(vlngPersonaGraba) & _
                    " where intNumCargo = " & vlngCveCargo
    pEjecutaSentencia strSentencia
    '---------------------------------------------
    'grdCargos.TextMatrix(vlngRow, 4) = FormatCurrency(inmInfoHonorario.dblImporteCargo)
End Sub



Private Sub cmdCargar_Click()
    On Error GoTo NotificaError
    Dim vlblnError As Boolean
    Dim vllngResultado As Long
    Dim SQL As String
    Dim strCodigo As String
    Dim blnExcluido As Boolean
    Dim blnEdit As Boolean
    Dim blnAutoriza As Boolean
    Dim vlstrSentencia As String
    Dim vllngProveedor As Long
    Dim rsServSub As ADODB.Recordset
    Dim blnCargoPerteneceConceptoHonorarios As Boolean
    

    vlblnError = False
    
    If vlblnTransaccionenCurso Then Exit Sub
    
    If Val(txtCantidad.Text) <> 0 Then
        If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
            If vllngPersonaGraba = 0 Then
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            End If
            If vllngPersonaGraba <> 0 Then
                    'Se revisa si es excluido antes de cargarse:
                If Not fblnCargoExcluidoContinuar(Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E"), grdElementos.RowData(grdElementos.Row), grdElementos.TextMatrix(grdElementos.Row, 2), blnAutoriza) Then
                    Exit Sub
                Else
                    If blnAutoriza Then
                            lstrCodigo = ""
                            blnExcluido = lblnExcluir
                            strCodigo = lstrCodigo
                            If fblnAceptarCargoExcluido(blnExcluido, strCodigo) Then
                                blnEdit = True
                            Else
                                Exit Sub
                            End If
                    End If
                End If
                blnCargoPerteneceConceptoHonorarios = False
                '|  Si el elemento que se está agregando es de tipo OC se validará si pertenece al concepto de facturación configurado para el honorario médico
                If grdElementos.TextMatrix(grdElementos.Row, 2) = "OC" Then
                    blnCargoPerteneceConceptoHonorarios = fblnCargoPerteneceConceptoHonorarios(grdElementos.RowData(grdElementos.Row))
                    If blnCargoPerteneceConceptoHonorarios Then
                        '| Solicita la información para generar el honorario
                        If Not fblnSolicitaInformacionHonorario(-1, Val(txtCantidad.Text)) Then Exit Sub
                    End If
                End If
                
                'Se revisa si es excluido antes de cargarse:
                'Revisar si es otro concepto examen grupo de examenes o estudio y si es que si es un cargo subrogado y de quien
                If grdElementos.TextMatrix(grdElementos.Row, 2) = "ES" Or _
                        grdElementos.TextMatrix(grdElementos.Row, 2) = "OC" Or _
                            grdElementos.TextMatrix(grdElementos.Row, 2) = "EX" Or _
                                grdElementos.TextMatrix(grdElementos.Row, 2) = "GE" Then
                    
'                    vlstrSentencia = "SELECT css.INTCVESERVICIOSUB, css.INTCVEPROVEEDOR,cop.VCHNOMBRECOMERCIAL " & _
'                     " FROM coserviciosubrogado css left join " & _
'                     " COPROVEEDOR cop on cop.INTCVEPROVEEDOR = css.INTCVEPROVEEDOR " & _
'                     " WHERE chrtiposervicio = '" & grdElementos.TextMatrix(grdElementos.Row, 2) & "'" & _
'                     " AND intcvetiposervicio = " & grdElementos.RowData(grdElementos.Row) & _
'                     " AND TNYCLAVEEMPRESA =  " & vgintClaveEmpresaContable & _
'                     " AND (INTCVEEMPRESA =  " & vgintEmpresa & " OR INTCVEEMPRESA = 0)" & _
'                     " AND TRUNC (SYSDATE) <= dtmfechafinal "
'                    Set rsServSub = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    
                    vlstrSentencia = grdElementos.TextMatrix(grdElementos.Row, 2) & "|" & grdElementos.RowData(grdElementos.Row) & "|" & vgintEmpresa & "|" & vgintClaveEmpresaContable
                    Set rsServSub = frsEjecuta_SP(vlstrSentencia, "sp_coselprovsubvarios")
                    
                    If rsServSub.RecordCount > 0 Then
                        frmSeleccionProveedorSubrogado.lblCargo.Caption = grdElementos.TextMatrix(grdElementos.Row, 1)
                        'pLlenarCboRs_new frmSeleccionProveedorSubrogado.cboProveedores, rsServSub, 0, 2, 4
                        pLlenarCboRs_new frmSeleccionProveedorSubrogado.cboProveedores, rsServSub, 0, 2
                        frmSeleccionProveedorSubrogado.cboProveedores.ListIndex = 0
                        frmSeleccionProveedorSubrogado.Show vbModal
                        If frmSeleccionProveedorSubrogado.cboProveedores.ListIndex >= 0 Then
                            vllngProveedor = frmSeleccionProveedorSubrogado.cboProveedores.ItemData(frmSeleccionProveedorSubrogado.cboProveedores.ListIndex)
                        Else
                            If frmSeleccionProveedorSubrogado.cboProveedores.ListIndex = -1 Then Exit Sub
                        End If
                        Unload frmSeleccionProveedorSubrogado
                        'If vllngProveedor = 0 Then Exit Sub
                    'Else
                    '    If rsServSub.RecordCount > 0 Then
                    '        vllngProveedor = rsServSub!INTCVESERVICIOSUB
                    '    End If
                    End If
                End If
                
                With EntornoSIHO
                    .ConeccionSIHO.BeginTrans
                    vlblnTransaccionenCurso = True
                    vllngResultado = 1
                    vgstrParametrosSP = grdElementos.RowData(grdElementos.Row) & "|" & IIf(lblnSeleccionarDepto, cboDepartamento.ItemData(cboDepartamento.ListIndex), vgintNumeroDepartamento) & "|" & "D" & "|" & 0 & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & grdElementos.TextMatrix(grdElementos.Row, 2) & "|" & chkMedicamentoAplicado.Value & "|" & Val(txtCantidad.Text) & "|" & vllngPersonaGraba & "|" & 0 & "|" & "" & "|" & 0 & "|" & 2
                    frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCARGOS", True, vllngResultado
                    If vllngResultado > 0 Then
                        If blnEdit Then
                            frsEjecuta_SP vllngResultado & "|" & IIf(blnExcluido, "1", "0") & "|" & strCodigo, "sp_PVUpdCodigoAut", True
                            If Not blnExcluido Then
                                frsEjecuta_SP txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & vllngResultado & "|1|0|0|0|0|0|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|0", "sp_PVUpdTrasladoCargos", True
                            End If
                        End If
                        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "CARGO DIRECTO EN CAJA", CStr(grdElementos.RowData(grdElementos.Row)))
                        pLlenaCargos
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
                        grdCargos.Row = grdCargos.Rows - 1
                        'Si el cargo que se está realizando es el mismo que está configurado para los honorarios médicos
                        If blnCargoPerteneceConceptoHonorarios Then
                            pActualizaPrecioCargo vllngResultado, vllngPersonaGraba, grdCargos.Row
                        End If
                        txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)) = "", Trim(grdCargos.TextMatrix(grdCargos.Row, 2)), fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)))
'****************************************************************************************************************************************************************************************
'****************************************************************************************************************************************************************************************
                        
                        'El cargo se realizó satisfactoriamente
                        SQL = "delete from PVTIPOPACIENTEPROCESO where PVTIPOPACIENTEPROCESO.intnumerologin = " & vglngNumeroLogin & _
                            "and PVTIPOPACIENTEPROCESO.intproceso = " & enmTipoProceso.Cargos
                        pEjecutaSentencia SQL
                        
                        SQL = "insert into PVTIPOPACIENTEPROCESO (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.Cargos & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
                        pEjecutaSentencia SQL
                        If vllngProveedor <> 0 Then
                            vgstrParametrosSP = vllngResultado & "|" & vllngProveedor & "|" & vgintClaveEmpresaContable & "|0"
                            frsEjecuta_SP vgstrParametrosSP, "sp_pvinscargoservsub"
                        End If
                        .ConeccionSIHO.CommitTrans
                        vlblnTransaccionenCurso = False
                        'agregar a tabla nueva iddepvcargo y proveedor seleccionado
                        MsgBox SIHOMsg(316), vbInformation, "Mensaje"
                    Else
                        .ConeccionSIHO.RollbackTrans
                        vlblnTransaccionenCurso = False
                        MsgBox SIHOMsg(vllngResultado * -1), vbExclamation, "Mensaje"
                        vlblnError = True
                    End If
                End With
                pLlenaCargos
                pEnfocaTextBox txtBuscaElemento
                
            End If
            cmdCargar.Enabled = False
        Else
            MsgBox SIHOMsg(65), vbExclamation, "Mensaje"
        End If
    Else
        MsgBox "Imposible realizar un cargo con cantidad cero(0)", vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargar_Click"))
    Unload Me
End Sub



Private Function fblnSolicitaInformacionHonorario(vlngCveCargo As Long, vintCantidadCargo As Integer) As Boolean
    fblnSolicitaInformacionHonorario = True
    frmInformacionHonorarioMedico.vglngCveCargo = vlngCveCargo
    frmInformacionHonorarioMedico.vgintCantidadCargo = vintCantidadCargo
    frmInformacionHonorarioMedico.Show vbModal, Me
    '| Si no se capturó la información requerida se sale y no realiza el cargo
    If frmInformacionHonorarioMedico.vgblnInformacionValida Then
        'Carga la informacion del honorario
        With inmInfoHonorario
            .intCveMedico = frmInformacionHonorarioMedico.vgintCveMedico
            .strProcedimiento = frmInformacionHonorarioMedico.vgstrProcedimiento
            .dblImporteCargo = Val(Format(frmInformacionHonorarioMedico.vgdblImporteCargo, "############.##"))
            .dblImporteHonorario = Val(Format(frmInformacionHonorarioMedico.vgdblImporteHonorario, "############.##"))
        End With
    Else
        'Limpia la informacion del honorario
        With inmInfoHonorario
            .intCveMedico = -1
            .strProcedimiento = ""
            .dblImporteCargo = -1
            .dblImporteHonorario = -1
        End With
        fblnSolicitaInformacionHonorario = False
    End If

End Function


'---------------------------------------------------------------------------------------------
'  Regresa la clave del concepto de facturación configurado para los honorarios médicos,
'  si no está configurado regresa un -1
'---------------------------------------------------------------------------------------------
Private Function fintDameConceptoFacturacionHonorario() As Long
    Dim strSentencia As String
    Dim rsConceptoFacturacionHonorario As New ADODB.Recordset

    fintDameConceptoFacturacionHonorario = -1
    strSentencia = "Select INTCVECONCEPTOHONORARIOMEDICO From PvParametro Where PvParametro.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rsConceptoFacturacionHonorario = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    If rsConceptoFacturacionHonorario.RecordCount > 0 Then
        fintDameConceptoFacturacionHonorario = IIf(IsNull(rsConceptoFacturacionHonorario!intCveConceptoHonorarioMedico), -1, rsConceptoFacturacionHonorario!intCveConceptoHonorarioMedico)
    End If


End Function

Private Sub cmdImprimirEstado_Click()
    frmReporteEstadoCuenta.llngNumeroCuenta = Val(txtMovimientoPaciente.Text)
    frmReporteEstadoCuenta.lstrTipoPaciente = IIf(OptTipoPaciente(0).Value, "I", "E")
    frmReporteEstadoCuenta.Show vbModal
End Sub

Private Sub cmdQuitarTodo_Click()
    If grdCargos.TextMatrix(1, 1) <> "" Then
        Dim vlintContador As Integer
        For vlintContador = 1 To Me.grdCargos.Rows - 1
        
            If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "F" Then
            
            Else
                grdCargos.TextMatrix(vlintContador, 0) = ""
            End If
            
        Next
    Else
    
    End If
    
End Sub

Private Sub cmdSeleccionar_Click()

    Dim vllngColumnaActual As Integer

    If grdCargos.TextMatrix(1, 1) <> "" Then
        
        Dim i As Integer
        For i = 0 To grdCargos.RowSel - grdCargos.Row
            
            vllngColumnaActual = grdCargos.Col
            grdCargos.Col = 0
            grdCargos.TextMatrix((grdCargos.Row + i), 0) = "*"
            vllngMarcados = vllngMarcados + 1
            grdCargos.Col = vllngColumnaActual
        
        Next
    Else
    
    End If

    If fblnCanFocus(grdCargos) Then
        grdCargos.SetFocus
    End If

End Sub

Private Sub cmdSeleccionarTodo_Click()
    
    If grdCargos.TextMatrix(1, 1) <> "" Then
        Dim vlintContador2 As Integer
        For vlintContador2 = 1 To Me.grdCargos.Rows - 1
        
            If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "F" Then
            
            Else
                grdCargos.TextMatrix(vlintContador2, 0) = "*"
            End If
        
        Next
    Else
    
    End If
    
    If fblnCanFocus(grdCargos) Then
        grdCargos.SetFocus
    End If
    

End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    'Seguridad para deshabilitar el de los medicamentos aplicados y no aplicados
    '-------------------------------------------------------
    'Revisamos si tiene permiso para modificar el medicamento aplicado
    '-------------------------------------------------------
    If cgstrModulo = "PV" Then
        chkMedicamentoAplicado.Enabled = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionHabilitaMedicamentoAplicado, "E", True)
    Else
        chkMedicamentoAplicado.Enabled = True
    End If
   ' If txtMovimientoPaciente <> "" Then
   '    If Not Me.vgblnDesdePV Then vlblnLlamadoPCE = True
   ' Else
   '    If Not Me.vgblnDesdePV Then vlblnLlamadoPCE = False
   ' End If
      
    frmCargosDirectosCaja.Refresh
    
    vgintColumnaCurrency = 4
    pConfiguraGridCargos

    vgstrNombreForm = Me.Name

    vlblnTransaccionenCurso = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = vbKeyEscape Then
        If vgstrEstadoManto <> "CE" Then
           txtMovimientoPaciente.Text = ""
           txtNombreComercial.Text = ""
           vlblnMantenerAutorizado = False
           txtBusqueda.Text = ""
           
        End If
        KeyAscii = 0
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon

    vllngPersonaGraba = 0
    vgstrEstadoManto = ""
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Cargos) > 0 Then
      If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Cargos) = 1 Then
        OptTipoPaciente(0).Value = True
      Else
        OptTipoPaciente(1).Value = True
      End If
    End If
    
    vlblnMantenerAutorizado = False
    vgintEmpresa = 0
    vgintTipoPaciente = 0
    pDepartamentoIngreso
     
    vlblnBotonNo = False
    Label11.Enabled = False
    txtBusqueda.Enabled = False
     
End Sub
Private Sub pDepartamentoIngreso()
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim llngIndex As Long
    
    lblnSeleccionarDepto = False
    If cgstrModulo = "PV" Or cgstrModulo = "CC" Then
        Set rs = frsSelParametros("PV", vgintClaveEmpresaContable, "BITSELECCIONARDEPTOCARGODIRECTO")
        If Not rs.EOF Then
            lblnSeleccionarDepto = rs!Valor = "1"
        End If
        rs.Close
    End If
    
    cboDepartamento.Clear
    cboDepartamento.Enabled = lblnSeleccionarDepto
    lblDepartamento.Enabled = lblnSeleccionarDepto
    Set rs = frsRegresaRs("SELECT smiCveDepartamento, trim(vchDescripcion) FROM NoDepartamento WHERE bitAtiendePacientes = 1 AND bitEstatus = 1 and tnyClaveEmpresa = " & vgintClaveEmpresaContable)
    pLlenarCboRs cboDepartamento, rs, 0, 1, -1
    rs.Close
    
    llngIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    If llngIndex = -1 Then
        cboDepartamento.AddItem vgstrNombreDepartamento
        cboDepartamento.ItemData(cboDepartamento.newIndex) = vgintNumeroDepartamento
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    Else
        cboDepartamento.ListIndex = llngIndex
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDepartamentoIngreso"))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    txtBusqueda.Text = ""
    txtMovimientoPaciente.Text = ""
    
    If vlblnLlamadoPCE Then Exit Sub
    
    If vgstrEstadoManto = "C" Then
        Cancel = 1
        vgstrEstadoManto = ""
        pLimpiaGrid grdCargos
        pCancelar
        pEnfocaTextBox txtMovimientoPaciente
    Else
        If vgstrEstadoManto = "CE" Then
            Cancel = 1
            vgstrEstadoManto = "C"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub pEliminaInfoHonorarioMedico(vlngNumCargo As Long)
    Dim strSentencia As String
    
    strSentencia = "Delete From PVBASEHONORARIOMEDICO Where INTNUMCARGO = " & vlngNumCargo
    pEjecutaSentencia strSentencia
End Sub

Private Sub GrdCargos_DblClick()
    On Error GoTo NotificaError
    
    Dim vllngColumnaActual As Long
    
    
    'Sirve para ordenar las columnas, 2 = Descripción del cargo, 5 = Fecha, 8 = Descripcion del cargo, 9 = Departamento
    If (grdCargos.MouseRow = 0 And (grdCargos.Col = 2 Or grdCargos.Col = 8 Or grdCargos.Col = 9)) Then
          
        ' Ordena en forma ascendente
        If Modo Then
            grdCargos.Col = grdCargos.Col
            grdCargos.Sort = 2
            Modo = False
        ' Ordena en forma descendente
        Else
            grdCargos.Col = grdCargos.Col
            grdCargos.Sort = 1
            Modo = True
        End If
        
    ElseIf (grdCargos.MouseRow = 0 And grdCargos.Col = 5) Then
        If Modo Then
            grdCargos.Col = 15
            grdCargos.Sort = 2
            Modo = False
        ' Ordena en forma descendente
        Else
            grdCargos.Col = 15
            grdCargos.Sort = 1
            Modo = True
        End If
    
    ElseIf grdCargos.MouseRow = 0 And grdCargos.MouseCol = 0 Then
        
        If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "*" Then
            Dim vlintContador As Integer
            For vlintContador = 1 To Me.grdCargos.Rows - 1
            
                If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "F" Then
                
                Else
                    grdCargos.TextMatrix(vlintContador, 0) = ""
                End If
                
            Next
        Else
            Dim vlintContador2 As Integer
            For vlintContador2 = 1 To Me.grdCargos.Rows - 1
            
                If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "F" Then
                
                Else
                    grdCargos.TextMatrix(vlintContador2, 0) = "*"
                End If
            
            Next
        End If
        
    ElseIf grdCargos.MouseRow <> 0 Then
    
    
        If Trim(grdCargos.TextMatrix(grdCargos.Row, 1)) <> "" Then
            If Trim(grdCargos.TextMatrix(grdCargos.Row, 0)) = "*" Then
                grdCargos.TextMatrix(grdCargos.Row, 0) = ""
                vllngMarcados = vllngMarcados - 1
                
            Else
                vllngColumnaActual = grdCargos.Col
                grdCargos.Col = 0
                grdCargos.TextMatrix(grdCargos.Row, 0) = "*"
                vllngMarcados = vllngMarcados + 1
                grdCargos.Col = vllngColumnaActual
                
            End If
        End If
        
    Else
        
    End If
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_DblClick"))
    Unload Me
End Sub

Private Sub pBorrarCargo()
    On Error GoTo NotificaError
    
    Dim vllngResultado As Long
    Dim vllngPersonaGraba As Long
    Dim rsAfectaKardex As ADODB.Recordset
    Dim vlstrTemp As String
    Dim vlintMensaje As Integer
    Dim rsFacturaSubCxP As ADODB.Recordset
    Dim rsFacturaSubHono As ADODB.Recordset
    Dim rsCargo As ADODB.Recordset
    
    '-------------------------
    'No se pueden borrar los cargos FACTURADOS
    '-------------------------
    If grdCargos.TextMatrix(grdCargos.Row, 0) = "F" Then Exit Sub
    '-------------------------
    'No se pueden borrar los cargos que sean de VALES
    '-------------------------
    'If grdCargos.TextMatrix(grdCargos.Row, 6) = "Vale" Then Exit Sub

    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "C") Then
        If grdCargos.RowData(1) <> -1 Then
        
        
            vlstrTemp = "select intnumcargo from pvcargo where intnumcargo = " & grdCargos.RowData(grdCargos.Row)
            Set rsCargo = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
            If rsCargo.RecordCount = 0 Then
                vlblnavisonoexistente = True
                Exit Sub
            End If
                
            vlstrTemp = "SELECT CPCUENTAPAGAR.VCHNUMEROFACTURA, COPROVEEDOR.VCHNOMBRE FROM CPCUENTAPAGAR INNER JOIN CPCARGOSSUBCONTRARECIBO ON CPCARGOSSUBCONTRARECIBO.INTNUMEROCXP = CPCUENTAPAGAR.INTNUMEROCXP " & _
                                "INNER JOIN COPROVEEDOR ON CPCUENTAPAGAR.INTCVEPROVEEDOR = COPROVEEDOR.INTCVEPROVEEDOR " & _
                                "WHERE CPCARGOSSUBCONTRARECIBO.INTCVECARGOSERVSUB = " & grdCargos.RowData(grdCargos.Row)
            Set rsFacturaSubCxP = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
            vlstrTemp = "SELECT CPHONORARIODIRECTO.VCHRDESCRIPCIONHONORARIO, COPROVEEDOR.VCHNOMBRE FROM CPHONORARIODIRECTO INNER JOIN CPCARGOSSUBHONORARIODIRECTO ON CPCARGOSSUBHONORARIODIRECTO.INTIDHONORARIO = CPHONORARIODIRECTO.INTIDHONORARIO " & _
                                "INNER JOIN COPROVEEDOR ON CPHONORARIODIRECTO.INTCVEPROVEEDOR = COPROVEEDOR.INTCVEPROVEEDOR " & _
                                "WHERE CPCARGOSSUBHONORARIODIRECTO.INTCVECARGOSERVSUB = " & grdCargos.RowData(grdCargos.Row)
            Set rsFacturaSubHono = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
            If rsFacturaSubCxP.RecordCount > 0 Then
                MsgBox SIHOMsg(1584) & " con una factura del proveedor. " & vbCrLf & vbCrLf & "Factura del proveedor " & rsFacturaSubCxP!vchNombre & ":" & vbCrLf & rsFacturaSubCxP!VCHNUMEROFACTURA, vbExclamation, "Mensaje"
                Exit Sub
            ElseIf rsFacturaSubHono.RecordCount > 0 Then
                MsgBox SIHOMsg(1584) & " con un recibo de honorarios del proveedor. " & vbCrLf & vbCrLf & "Recibo del proveedor " & rsFacturaSubHono!vchNombre & ":" & vbCrLf & rsFacturaSubHono!VCHRDESCRIPCIONHONORARIO, vbExclamation, "Mensaje"
                Exit Sub
            End If
            vlstrTemp = "select intnumkardex from pvcargo where intnumcargo = " & grdCargos.RowData(grdCargos.Row)
            Set rsAfectaKardex = frsRegresaRs(vlstrTemp, adLockReadOnly, adOpenForwardOnly)
            If rsAfectaKardex!INTNUMKARDEX > 0 Then
                'El movimiento no afectará las existencias y podría generar cantidades negativas en el recálculo del artículo ¿Desea continuar?
                vlintMensaje = 1098
            Else
                'Esta seguro de borrar el cargo
                vlintMensaje = 1627
            End If
            
            If vlblnMensajeAviso Then
            
                        If vlblnMantenerAutorizado = False Then
                            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                        Else
                        
                        End If
                        If vllngPersonaGraba <> 0 Or vlblnMantenerAutorizado = True Then
                            vlblnMantenerAutorizado = True
                            vllngPersonaGraba = vlintUsuarioAutorizado
                            ' Elimina la información de los honorarios, relacionada con el cargo
                            If grdCargos.TextMatrix(grdCargos.Row, 1) = "OC" Then
'                                If fblnCargoPerteneceConceptoHonorarios(grdCargos.TextMatrix(grdCargos.Row, 13)) Then
                                    pEliminaInfoHonorarioMedico (grdCargos.RowData(grdCargos.Row))
'                                End If
                            End If
                            vllngResultado = 1
                            vgstrParametrosSP = grdCargos.RowData(grdCargos.Row) & "|" & "ECCD" & "|" & vlintUsuarioAutorizado & "|" & vgintNumeroDepartamento & "|" & "" & "|" & 0 & "|" & 0 & "|" & "" & "|" & 2
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDBORRACARGO", False, vllngResultado
                            If vllngResultado = 0 Then
                                vgstrParametrosSP = grdCargos.RowData(grdCargos.Row) & "|" & vgintClaveEmpresaContable
                                frsEjecuta_SP vgstrParametrosSP, "sp_pvdelcargoservsub"
                                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vlintUsuarioAutorizado, "CARGO DIRECTO EN CAJA", CStr(grdCargos.RowData(grdCargos.Row)))
                                
                            Else
                                MsgBox SIHOMsg(CInt(vllngResultado)), vbExclamation, "Mensaje"
                            End If
                        End If

            
            
            Else
            
                If Not vlblnBotonNo Then
            
                    If MsgBox(SIHOMsg(vlintMensaje), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        ' Persona que graba
                        
                        vllngPersonaGraba = vlintUsuarioAutorizado
                        
                        If vlblnMantenerAutorizado = False Then
                            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                            If vllngPersonaGraba <> 0 Then
                                vlblnMantenerAutorizado = True
                                vlblnMensajeAviso = True
                                
                            Else
                                vlblnMensajeAviso = False
                                vlblnBotonNo = True
                            End If
                        ElseIf vlblnMantenerAutorizado = True Then
                            vlblnMensajeAviso = True
                        Else
                        
                        End If
                        If vllngPersonaGraba <> 0 Or vlblnMantenerAutorizado = True Then
                            vlblnMantenerAutorizado = True
                            vlintUsuarioAutorizado = vllngPersonaGraba
                            ' Elimina la información de los honorarios, relacionada con el cargo
                            If grdCargos.TextMatrix(grdCargos.Row, 1) = "OC" Then
'                                If fblnCargoPerteneceConceptoHonorarios(grdCargos.TextMatrix(grdCargos.Row, 13)) Then
                                    pEliminaInfoHonorarioMedico (grdCargos.RowData(grdCargos.Row))
'                                End If
                            End If
                            vllngResultado = 1
                            vgstrParametrosSP = grdCargos.RowData(grdCargos.Row) & "|" & "ECCD" & "|" & vlintUsuarioAutorizado & "|" & vgintNumeroDepartamento & "|" & "" & "|" & 0 & "|" & 0 & "|" & "" & "|" & 2
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDBORRACARGO", False, vllngResultado
                            If vllngResultado = 0 Then
                                vgstrParametrosSP = grdCargos.RowData(grdCargos.Row) & "|" & vgintClaveEmpresaContable
                                frsEjecuta_SP vgstrParametrosSP, "sp_pvdelcargoservsub"
                                Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vlintUsuarioAutorizado, "CARGO DIRECTO EN CAJA", CStr(grdCargos.RowData(grdCargos.Row)))
                                
                            Else
                                MsgBox SIHOMsg(CInt(vllngResultado)), vbExclamation, "Mensaje"
                            End If
                        End If
                    Else
                    
                        vlblnBotonNo = True
                    
                    End If
                Else
                
                End If
            End If
        End If
    Else
        '257 antiguo mensaje ¡No se pueden borrar los datos!
        MsgBox SIHOMsg(810), vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_DblClick"))
    Unload Me
End Sub


Private Sub grdCargos_RowColChange()
    
    If grdCargos.Row > 0 Then
        '  Agrega la información del honorario médico
        
        txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)) = "", Trim(grdCargos.TextMatrix(grdCargos.Row, 2)), fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)))
    End If
End Sub

Private Function fstrInfoMedicoProcedimiento(v_lngCveCargo As Long) As String
    Dim rsInfoHonorarios As New ADODB.Recordset
    Dim strSentencia As String
    Dim strInfoHonorario As String

    strInfoHonorario = ""
    strSentencia = "Select FN_PVHONORARIOMEDICO(" & v_lngCveCargo & ") InfoHonorario From Dual"
    Set rsInfoHonorarios = frsRegresaRs(strSentencia, adLockReadOnly, adOpenStatic)
    If rsInfoHonorarios.RecordCount <> 0 Then
        fstrInfoMedicoProcedimiento = Trim(IIf(IsNull(rsInfoHonorarios!InfoHonorario), "", rsInfoHonorarios!InfoHonorario))
        DoEvents
    End If

End Function

Private Sub grdElementos_DblClick()
    On Error GoTo NotificaError
    
    pEnfocaTextBox txtCantidad
    cmdCargar.Enabled = True
    txtBusqueda.Text = ""

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdElementos_DblClick"))
    Unload Me
End Sub
Private Sub grdElementos_GotFocus()
    On Error GoTo NotificaError
    
    If grdElementos.RowData(1) = -1 Then
        grdElementos.Enabled = False
        FreElementos.Enabled = False
    End If
    If grdElementos.Row > 0 Then
      txtNombreComercial.Text = Trim(grdElementos.TextMatrix(grdElementos.Row, 1))
      grdElementos.Col = 1
      grdElementos.CellForeColor = &HFF0000 '&H80000008
      grdElementos.CellFontBold = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdElementos_GotFocus"))
    Unload Me
End Sub
Private Sub grdElementos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtCantidad
        cmdCargar.Enabled = True
        txtBusqueda.Text = ""
    End If
    
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        txtNombreComercial.Text = Trim(grdElementos.TextMatrix(grdElementos.Row, 1))
      grdElementos.Col = 1
      grdElementos.CellForeColor = &HFF0000 '&H80000008
      grdElementos.CellFontBold = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdElementos_KeyDown"))
    Unload Me
End Sub

Private Sub grdElementos_LeaveCell()

      grdElementos.Col = 1
      grdElementos.CellForeColor = &H80000008
      grdElementos.CellFontBold = False

End Sub

Private Sub grdElementos_RowColChange()
  If grdElementos.Row > 0 Then
    txtNombreComercial.Text = Trim(grdElementos.TextMatrix(grdElementos.Row, 1))
    grdElementos.Col = 1
      grdElementos.CellForeColor = &HFF0000 '&H80000008
    grdElementos.CellFontBold = True
    
  End If
End Sub

Private Sub MskFecha_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdCargos
        Select Case KeyCode
            Case 27   'ESC
                 MskFecha.Visible = False
                .SetFocus
                'vgstrEstadoManto = "C"
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    vgblnNoEditarFecha = True
                    .Row = .Row - 1
                    vgblnNoEditarFecha = False
                End If
                txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(.RowData(.Row)) = "", Trim(.TextMatrix(.Row, 2)), fstrInfoMedicoProcedimiento(.RowData(.Row)))
            Case 13
               If MskFecha.Text <> Format(.TextMatrix(.Row, .Col), "dd/mm/yyyy HH:mm") Then ' cambio el valor del campo?
                  If IsDate(MskFecha.Text) Then ' el nuevo valor es una fecha?
                     If fblnFechaenRango Then  ' la fecha no puede ser menor a la fecha de ingreso ni mayor a la de egreso(en caso de que existe)
                        pSetCellValueColFecha grdCargos, MskFecha
                     Else
                        pEnfocaMkTexto MskFecha
                     End If
                  Else ' no es una fecha, para atras los filders
                  '¡Fecha no válida!, formato de fecha dd/mm/yyyy HH:mm
                  MsgBox SIHOMsg(1232), vbOKOnly + vbExclamation, "Mensaje"
                  pEnfocaMkTexto MskFecha
                  End If
               Else ' no cambio el valor del campo fecha
                  MskFecha.Visible = False
                  .SetFocus
               End If
            Case 40
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    vgblnNoEditarFecha = True
                    .Row = .Row + 1
                    vgblnNoEditarFecha = False
                Else
                    .Row = 1
                End If
                txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(.RowData(.Row)) = "", Trim(.TextMatrix(.Row, 2)), fstrInfoMedicoProcedimiento(.RowData(.Row)))
        End Select
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":MskFecha_KeyDown"))
    Unload Me
End Sub
Private Function fblnFechaenRango() As Boolean
'Valida que la fecha que se intenta colocar en un cargo no sea menor a la fecha de ingreso del paciente
'ni mayor a la fecha de egreso del paciente
Dim objSTR As String
Dim ObjRS As New ADODB.Recordset
Dim objStrFecha As String

fblnFechaenRango = True
objStrFecha = Format(MskFecha.Text, "dd/mm/yyyy HH:mm")

objSTR = "Select DTMFECHAHORAINGRESO,DTMFECHAHORAEGRESO from EXPACIENTEINGRESO where INTNUMCUENTA = " & Me.txtMovimientoPaciente & _
          " and CHRTIPOINGRESO ='" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'"
Set ObjRS = frsRegresaRs(objSTR, adLockOptimistic)

If ObjRS.RecordCount = 0 Then ' no hay registro, no se hace la modificación(ALGO ESTA PASANDO)
   fblnFechaenRango = False
Else ' si existe , ahora validamos la fecha
   'la fecha no puede ser menor a la fecha de ingreso
   If Format(ObjRS!DTMFECHAHORAINGRESO, "dd/mm/yyyy HH:mm") > CDate(objStrFecha) Then '
   fblnFechaenRango = False
   '¡La fecha del cargo no puede ser menor a la fecha de ingreso del paciente!
    MsgBox SIHOMsg(1233), vbExclamation, "Mensaje"
   
   Else ' si la fecha no es menor a la fecha de ingreso
      If Not IsNull(ObjRS!DTMFECHAHORAEGRESO) Then ' buscamos si existe fecha de egreso
        'la fecha no puede ser mayor a la fecha de egreso
         If Format(ObjRS!DTMFECHAHORAEGRESO, "dd/mm/yyyy HH:mm") < CDate(objStrFecha) Then
         fblnFechaenRango = False
         '¡La fecha del cargo no puede ser mayor a la fecha de egreso del paciente!
         MsgBox SIHOMsg(1234), vbExclamation, "Mensaje"
'        Else 'la fecha no es mayor a la fecha de egreso
         End If
      Else ' no hay fecha de egreso
         'la fecha no puede ser mayor a la fecha actual
         If Format(fdtmServerFechaHora(), "dd/mm/yyyy HH:mm") < CDate(objStrFecha) Then
         fblnFechaenRango = False
         '¡La fecha debe ser menor o igual a la del sistema!
         MsgBox SIHOMsg(1235), vbExclamation, "Mensaje"
'        Else la fecha no es mayor a la fecha actual
         End If
      End If
    End If
End If
End Function
Private Sub mskFecha_LostFocus()
  If Not vlblnEditando Then
     MskFecha.Visible = False
  End If
End Sub
Private Sub optClaveDescripcion_Click(Index As Integer)
    On Error GoTo NotificaError
    
    txtBuscaElemento.Text = ""
    pEnfocaTextBox txtBuscaElemento
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optClaveDescripcion_Click"))
    Unload Me
End Sub

Private Sub optElementos_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo NotificaError
    
    pCargaElementos
    If grdElementos.Enabled Then
        grdElementos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optElementos_MouseUp"))
    Unload Me
End Sub

Private Sub txtBuscaElemento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyDown Then
        If grdElementos.Enabled Then
            grdElementos.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBuscaElemento_KeyDown"))
    Unload Me
End Sub
Private Sub txtBuscaElemento_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        If grdElementos.RowData(1) <> -1 Then
            grdElementos.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBuscaElemento_KeyPress"))
    Unload Me
End Sub
Private Sub txtBuscaElemento_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    pCargaElementos

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBuscaElemento_KeyUp"))
    Unload Me
End Sub

Private Sub txtBusqueda_Change()

Dim vlintMiRow As Integer
Dim X As Integer

    grdCargos.Col = 2
    grdCargos.Sort = 1

    If txtBusqueda.Text = "" Then
    
    Else

        For X = 1 To grdCargos.Rows - 1 '''iterate through each row to find the search string
            grdCargos.Row = X
            If Mid(grdCargos.Text, 1, Len(txtBusqueda.Text)) = txtBusqueda.Text Then
                
                With Me.grdCargos
                    .Row = X 'sets the current row
                    .Col = 0  'sets the current col and start of range
                    .ColSel = .Cols - 1 'sets the end range
                    .TopRow = X
                End With
                
                Exit Sub
            End If
        Next X
    
    End If

End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    If ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 241 Or KeyAscii = 209 Or KeyAscii = 32) Or (KeyAscii > 96 And KeyAscii < 123) Then
        txtBusqueda = txtBusqueda
    Else
        KeyAscii = 0
        Beep
    End If
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdCargar.Enabled And cmdCargar.Visible Then
          cmdCargar.SetFocus
        Else
          If txtBuscaElemento.Enabled And txtBuscaElemento.Visible Then txtBuscaElemento.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_KeyDown"))
    Unload Me
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

'Private Sub txtMovimientoPaciente_GotFocus()
'    If txtMovimientoPaciente.Text <> "" Then txtMovimientoPaciente_KeyDown vbKeyReturn, 0
'    cmdCargar.Enabled = False
'End Sub
Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsPostergado As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                If OptTipoPaciente(1).Value Then 'Externos
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = False
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    '.vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                Else
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = False
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha ing."", TO_CHAR(ExPacienteIngreso.dtmFechaHoraEgreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    '.vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                
                If txtMovimientoPaciente <> -1 Then
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
        Else
            '------------------------
            ' Progress Bar : "Cargando Datos..."
            '------------------------
            frmCargosDirectosCaja.Refresh
            freBarra.Top = 3000
            freBarra.MousePointer = ssHourglass
            pgbBarra.Value = 10
            freBarra.Visible = True
            freBarra.Refresh
            '------------------------
            If OptTipoPaciente(0).Value Then 'Internos
                vlstrSentencia = "SELECT rtrim(AdPaciente.vchApellidoPaterno)||' '||rtrim(AdPaciente.vchApellidoMaterno)||' '||rtrim(AdPaciente.vchNombre) as Nombre, " & _
                        "AdAdmision.intCveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "AdAdmision.tnyCveTipoPaciente cveTipoPaciente, AdTipoPaciente.vchDescripcion as Tipo,  " & _
                        "AdAdmision.vchNumCuarto Cuarto," & _
                        "AdAdmision.bitCuentaCerrada CuentaCerrada " & _
                        "FROM AdAdmision " & _
                        "INNER JOIN AdPaciente ON AdAdmision.numCvePaciente = AdPaciente.numCvePaciente " & _
                        "INNER JOIN AdTipoPaciente ON AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON ADADMISION.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where AdAdmision.numNumCuenta = " & txtMovimientoPaciente.Text & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            Else 'Externos
                vlstrSentencia = "SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as Nombre, " & _
                        "RegistroExterno.intClaveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "RegistroExterno.tnyCveTipoPaciente as cveTipoPaciente, AdTipoPaciente.vchDescripcion  as Tipo, '' as Cuarto, RegistroExterno.bitCuentaCerrada CuentaCerrada " & _
                        "FROM RegistroExterno " & _
                        "INNER JOIN Externo ON RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
                        "INNER JOIN AdTipoPaciente ON RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON REGISTROEXTERNO.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where RegistroExterno.intNumCuenta = " & txtMovimientoPaciente.Text & " And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            End If
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            
            Set rsPostergado = frsRegresaRs("SELECT BITPOSTERGADA FROM EXPACIENTEINGRESO WHERE INTNUMCUENTA = " & txtMovimientoPaciente.Text, adLockOptimistic, adOpenDynamic)
            
            If rs.RecordCount <> 0 Then
                If rs!CuentaCerrada = 1 And rsPostergado!BITPOSTERGADA = 1 Then
                    'La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
                    MsgBox "La cuenta se encuentra postergada, no es posible realizar cambios.", vbExclamation, "Mensaje"
                ElseIf rs!CuentaCerrada = 1 And rsPostergado!BITPOSTERGADA = 0 Then
                    'La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
                    MsgBox SIHOMsg(596), vbExclamation, "Mensaje"
                Else
                    pLlenaCargos
                    vgstrEstadoManto = "C" 'Cargando
                    txtMovimientoPaciente.Locked = True
                    OptTipoPaciente(0).Enabled = False
                    OptTipoPaciente(1).Enabled = False
                    txtPaciente.Locked = True
                    txtTipoPaciente.Locked = True
                    txtCuarto.Locked = True
                    txtEmpresaPaciente.Locked = True
                    FreElementos.Enabled = True
                    grdElementos.Enabled = True
                    txtPaciente.Text = rs!Nombre
                    
                    vgintEmpresa = IIf(IsNull(rs!cveEmpresa), 0, rs!cveEmpresa)
                    txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                    
                    vgintTipoPaciente = rs!cveTipoPaciente
                    
                    txtTipoPaciente.Text = rs!tipo
                    txtCuarto = IIf(IsNull(rs!Cuarto), 0, rs!Cuarto)
                    
                    chkMedicamentoAplicado.Value = 1
                    If fblnRevisaPermiso(vglngNumeroLogin, 305, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 305, "L", True) Or fblnRevisaPermiso(vglngNumeroLogin, 305, "E", True) Then
                        cmdImprimirEstado.Enabled = True
                        lblEstadoCuenta.Enabled = True
                    End If
                    FreDetalle.Enabled = True
                    FreAplicar.Enabled = True
                    grdElementos.Row = 1
                    grdElementos.ColSel = 1
                    pEnfocaTextBox txtBuscaElemento
                End If
            Else
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                pCancelar
            End If
            freBarra.Visible = False
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub
Public Sub pLlenaCargos()
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    Dim rsSeleccionaCargos As New ADODB.Recordset
    Dim lngAux As Long
    Dim lngAncho As Long
    
    
    '-------------------------------------------------------------------
    ' Este Procedure lo usan las pantallas de Cargos directos de Caja y de????
    '-------------------------------------------------------------------
    grdCargos.Redraw = False
    pLimpiaGrid grdCargos
    pConfiguraGridCargos
    lngAncho = 800
    lngAux = 0
    vgstrParametrosSP = txtMovimientoPaciente & "|" & IIf(OptTipoPaciente(0), "I", "E") & "|" & IIf(chkIncluyeFacturados.Value = 0, 0, 2) & "|" & "-1|C|N|0"
    Set rsSeleccionaCargos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCARGOSPACIENTE", 0)
    
    With rsSeleccionaCargos
        Do While Not .EOF
            If grdCargos.RowData(1) <> -1 Then
                 grdCargos.Rows = grdCargos.Rows + 1
                 grdCargos.Row = grdCargos.Rows - 1
            End If
            pgbBarra.Value = (.Bookmark / .RecordCount) * 100
            grdCargos.RowData(grdCargos.Row) = !IntNumCargo
            grdCargos.TextMatrix(grdCargos.Row, 0) = IIf(IsNull(!FolioFactura), "", "F")
            grdCargos.TextMatrix(grdCargos.Row, 1) = !chrTipoCargo
            grdCargos.TextMatrix(grdCargos.Row, 2) = IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo)
            grdCargos.TextMatrix(grdCargos.Row, 3) = !MNYCantidad
            grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(!mnyPrecio, 2)
            grdCargos.TextMatrix(grdCargos.Row, 5) = Format(!dtmFechahora, "dd/MMM/YYYY HH:mm")
            grdCargos.TextMatrix(grdCargos.Row, 6) = !TipoDocumento
            grdCargos.TextMatrix(grdCargos.Row, 7) = !intFolioDocumento
            grdCargos.TextMatrix(grdCargos.Row, 8) = !Concepto
            grdCargos.TextMatrix(grdCargos.Row, 9) = !nombreDepartamento
            grdCargos.TextMatrix(grdCargos.Row, 10) = IIf(IsNull(!FolioFactura), "", !FolioFactura)
            grdCargos.TextMatrix(grdCargos.Row, 15) = Format(!dtmFechahora, "yyyymmddhhmmss")
           
            PB.Font = grdCargos.CellFontName
            PB.FontSize = grdCargos.CellFontSize
            lngAux = PB.TextWidth(grdCargos.TextMatrix(grdCargos.Row, 10))
            If lngAux > lngAncho Then
               lngAncho = lngAux
            End If
            
            grdCargos.TextMatrix(grdCargos.Row, 11) = !intDescuentaInventario
            grdCargos.TextMatrix(grdCargos.Row, 12) = !CveConceptoFacturacion
            grdCargos.TextMatrix(grdCargos.Row, 13) = !CHRCVECARGO
            grdCargos.TextMatrix(grdCargos.Row, 14) = !Aplicado
                        
            If !PrecioManual Then
               grdCargos.Col = 4
               grdCargos.CellBackColor = &HC0FFFF
            End If '--------------------------------------------------------------------------
            If !FechaManual Then
               grdCargos.Col = 5
               grdCargos.CellBackColor = &HC0E0FF
            End If '---------------------------------------------------------------------------
            If !Excluido = "X" Then
               For vlintContador = 1 To grdCargos.Cols - 1
                   grdCargos.Col = vlintContador
                   grdCargos.CellForeColor = &HFF0000
               Next
            End If '---------------------------------------------------------------------------
            .MoveNext
        Loop
        If lngAncho > 800 Then grdCargos.ColWidth(10) = lngAncho + 100
    .Close
    End With
    grdCargos.Redraw = True

    If grdCargos.TextMatrix(1, 1) <> "" Then
        
        
        With grdCargos
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
        End With
        Label11.Enabled = True
        txtBusqueda.Enabled = True
        grdCargos.Redraw = True
    
    Else
    
        Label11.Enabled = False
        txtBusqueda.Enabled = False
        txtBusqueda.Text = ""
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCargos"))
    Unload Me
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    pEnfocaTextBox txtMovimientoPaciente

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            OptTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            OptTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
    Unload Me
End Sub

Private Sub GrdCargos_GotFocus()
    On Error GoTo NotificaError
    
    If grdCargos.Row > 0 Then
      txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)) = "", Trim(grdCargos.TextMatrix(grdCargos.Row, 2)), fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)))
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_GotFocus"))
    Unload Me
End Sub

Private Sub grdCargos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    vlblnavisonoexistente = False
    If KeyCode = vbKeyDelete Then
        
        Dim vlintMiRow As Integer
        Dim X As Integer

        grdCargos.Col = 0

            For X = 1 To grdCargos.Rows - 1
                grdCargos.Row = X
                If Mid(grdCargos.Text, 1, 1) = "*" Then
                    
                    With Me.grdCargos
                        .Row = X
                        .Col = 0
                        .ColSel = .Cols - 1
                    End With
                    
                    pBorrarCargo
                    
                    If vlblnBotonNo Then
                        vlblnMensajeAviso = False
                        vlblnBotonNo = False
                        Exit Sub
                    End If
                    
                    If grdCargos.TextMatrix(1, 1) <> "" Then
    
                        Label11.Enabled = True
                        txtBusqueda.Enabled = True
                    
                    Else
                    
                        Label11.Enabled = False
                        txtBusqueda.Enabled = False
                        txtBusqueda.Text = ""
                    
                    End If
                    
                End If
            Next X
            
        vlblnBotonNo = False
        pLimpiaGrid grdCargos
        pLlenaCargos
        If vlblnMensajeAviso Or vlblnavisonoexistente Then
            MsgBox SIHOMsg(1628), vbInformation, "Mensaje"
        Else
        
        End If
        vlblnMensajeAviso = False
    End If
    
    
    If grdCargos.Col = vgintColumnaCurrency Then
        If KeyCode = vbKeyF2 Then 'para que se edite el contenido de la celda como en excel
           pEditarColumna 13, txtPrecio, grdCargos
        End If
    ElseIf grdCargos.Col = 5 Then ' edición de la fecha
        If KeyCode = vbKeyF2 Then
           pEditarColumnaFecha 13, MskFecha, grdCargos
        End If
    Else
        If KeyCode = vbKeyReturn Then
            grdCargos.Col = 0
            grdCargos.CellFontBold = True
            grdCargos.Col = 1
            If grdCargos.Row - 1 < grdCargos.Rows Then
                If grdCargos.Row = grdCargos.Rows - 1 Then
                    grdCargos.Row = 1
                Else
                    grdCargos.Row = grdCargos.Row + 1
                    If grdCargos.Row = grdCargos.Rows - 1 Then
                        grdCargos.Row = 1
                    Else
                        grdCargos.Row = grdCargos.Row + 1
                    End If
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_KeyDown"))
    Unload Me
End Sub
Private Sub grdCargos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If grdCargos.Col = vgintColumnaCurrency Then 'Columna que puede ser editada
        pEditarColumna KeyAscii, txtPrecio, grdCargos
    ElseIf grdCargos.Col = 5 And KeyAscii = 13 Then  ' se preciona "Enter" y se esta en la columna fecha
        pEditarColumnaFecha KeyAscii, Me.MskFecha, grdCargos
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_KeyPress"))
    Unload Me
End Sub

Private Sub grdCargos_Scroll()
    On Error GoTo NotificaError
    
    If txtPrecio.Visible Then
      If grdCargos.Col = vgintColumnaCurrency Then
        If txtPrecio.Left <> grdCargos.Left + grdCargos.CellLeft Or _
           txtPrecio.Top <> grdCargos.Top + grdCargos.CellTop Or _
           txtPrecio.Width <> grdCargos.CellWidth - 8 Or _
           txtPrecio.Height <> grdCargos.CellHeight - 8 Then
          txtPrecio.Move grdCargos.Left + grdCargos.CellLeft, grdCargos.Top + grdCargos.CellTop, grdCargos.CellWidth - 8, grdCargos.CellHeight - 8
        End If
      End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_Scroll"))
    Unload Me
End Sub
Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vldblIVA As Double
    Dim vldblTotDescuento As Double
    Dim vlstrSentencia As String
    Dim vlchrTipoDescuento As String
    Dim vldblPorcentajeDescuento As Double
    Dim vllngPersonaGraba As Long
    Dim vlaryParametrosSalida() As String
    Dim rsContenido As New ADODB.Recordset
    Dim intContenido As Integer
    
    '---------------------------------------------------------
    ' NOTA:
    '       Este código debe ser llamado cada vez que
    '       el grid pierde el foco y su contenido puede cambiar.
    '       De otra manera, el nuevo valor de la celda se perdería.
    If grid.MouseCol = vgintColumnaCurrency Then
        grid.Col = vgintColumnaCurrency
    End If
    If grid.Col = vgintColumnaCurrency Then
        If txtPrecio.Visible Then
            If IsNumeric(Trim(txtPrecio.Text)) Then
                If Not Val(Trim(txtPrecio.Text)) > 0 Then
                  MsgBox SIHOMsg(452), vbExclamation, "Precio del cargo"
                  Exit Sub
                End If
                '¿Está seguro de que desea cambiar el precio del cargo?
                vlblnEditando = True
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                vlblnEditando = False
                If vllngPersonaGraba <> 0 Then
                
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    grid.Text = FormatCurrency(txtPrecio.Text, 2)
                    
                    intContenido = 1
                    If grdCargos.TextMatrix(grdCargos.Row, 2) = "AR" Then
                        Set rsContenido = frsEjecuta_SP(grdCargos.TextMatrix(grdCargos.Row, 13), "sp_GnSelArticulo")
                        If Not rsContenido.EOF Then intContenido = rs!intContenido
                    End If
                    
                    '-----------------------
                    'Descuentos
                    '-----------------------
                    vldblTotDescuento = 0
                    vlchrTipoDescuento = " "
                    
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    frsEjecuta_SP IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintTipoPaciente & "|" & _
                                    vgintEmpresa & "|" & CLng(Val(txtMovimientoPaciente.Text)) & "|" & _
                                    grdCargos.TextMatrix(grdCargos.Row, 1) & "|" & _
                                    CLng(Val(grdCargos.TextMatrix(grdCargos.Row, 13))) & "|" & _
                                    Val(Format(txtPrecio.Text, "###########0.00")) & "|" & _
                                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                    grdCargos.TextMatrix(grdCargos.Row, 14) & "|" & _
                                    intContenido & "|" & _
                                    Format(grdCargos.TextMatrix(grdCargos.Row, 3), "") & "|" & _
                                    grdCargos.TextMatrix(grdCargos.Row, 11), _
                                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblTotDescuento
                    '-----------------------
                    'IVA
                    '-----------------------
                    vlstrSentencia = "Select smyIva/100 IVA from pvConceptoFacturacion " & _
                                    " where smiCveConcepto = " & Trim(grdCargos.TextMatrix(grdCargos.Row, 12))
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    vldblIVA = rs!IVA * _
                            (Val(Format(grdCargos.TextMatrix(grdCargos.Row, 4), "############.##")) * _
                            Val(Format(grdCargos.TextMatrix(grdCargos.Row, 3), "############.##")) - _
                            vldblTotDescuento)
                    rs.Close
                    vlstrSentencia = "Update pvCargo set mnyPrecio = " & Format(txtPrecio.Text, "###########0.00") & _
                                    ", mnyIVA = " & Trim(str(Round(vldblIVA, 6))) & _
                                    ", mnyDescuento = " & Trim(str(Round(vldblTotDescuento, 6))) & _
                                    ", bitPrecioManual = 1 " & _
                                    ", intEmpleado = " & str(vllngPersonaGraba) & _
                                    " where intNumCargo = " & Trim(str(grdCargos.RowData(grdCargos.Row)))
                    pEjecutaSentencia vlstrSentencia
                    '---------------------------------------------
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "CARGO DIRECTO EN CAJA (CAMBIO DE PRECIO)", CStr(grdCargos.RowData(grdCargos.Row)))
                    
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                End If
                vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 1)
            End If
            txtPrecio.Visible = False
            pLlenaCargos
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSetCellValueCol"))
    Unload Me
End Sub
Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    
    Dim vlintTexto As Integer
    
    '-------------------------
    'Que se salga cuando si no tiene permiso
    '-------------------------
    If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3068, 3071), "E") Then Exit Sub
    '-------------------------
    'Que se salga cuando ya esta facturado
    '-------------------------
    If grid.TextMatrix(grid.Row, 0) = "F" Then Exit Sub
    
    '|  Si el elemento que se está agregando es de tipo OC se validará si pertenece al concepto de facturación configurado para el honorario médico
    If grid.TextMatrix(grid.Row, 1) = "OC" And fblnCargoPerteneceConceptoHonorarios(grid.TextMatrix(grid.Row, 13)) Then
        If vllngPersonaGraba = 0 Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        End If
        If vllngPersonaGraba <> 0 Then
            '| Solicita la información para generar el honorario
            If fblnSolicitaInformacionHonorario(grid.RowData(grid.Row), grid.TextMatrix(grid.Row, 3)) Then
                pActualizaPrecioCargo grid.RowData(grid.Row), vllngPersonaGraba, grid.Row
            End If
            txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grid.RowData(grid.Row)) = "", Trim(grid.TextMatrix(grid.Row, 2)), fstrInfoMedicoProcedimiento(grid.RowData(grid.Row)))
        End If
        Exit Sub
    End If
        
        
    With txtEdit
       If Val(Format(grid.Text, "############.##")) = 0 Then
            .Text = FormatCurrency((Val(Format(grid.TextMatrix(grid.Row, 3), "############.##")) - Val(Format(grid.TextMatrix(grid.Row, 4), "############.##"))), 2)
       Else
            .Text = Replace(grid, "$", "") 'Inicialización del Textbox
       End If
       
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                    .SelStart = 0
                    .SelLength = 1000
            Case 8, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
            Case 46
                ' Reemplaza el texto actual solo si se teclean números
                .Text = "."
                .SelStart = 1
        End Select
    End With
    
    ' Muestra el textbox en el lugar indicado
    With grid
    
        If .CellWidth < 0 Then Exit Sub
            txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    If vgstrEstadoManto = "C" Then vgstrEstadoManto = vgstrEstadoManto & "E"
    'vgstrEstadoManto = vgstrEstadoManto & "E"
    txtEdit.Visible = True
    txtEdit.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumna"))
    Unload Me
End Sub
Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdCargos
        Select Case KeyCode
            Case 27   'ESC
                 txtPrecio.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    vgblnNoEditarPagos = True
                    .Row = .Row - 1
                    vgblnNoEditarPagos = False
                End If
                txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)) = "", Trim(.TextMatrix(.Row, 2)), fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)))
                vgblnEditaPago = False
                vgstrEstadoManto = "C"
            Case 13
                If txtPrecio.Text <> Replace(.TextMatrix(.Row, .Col), "$", "") Then
                   Call pSetCellValueCol(grdCargos, txtPrecio)
                Else
                   txtPrecio.Visible = False
                  .SetFocus
                End If
            Case 40
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    vgblnNoEditarPagos = True
                    .Row = .Row + 1
                    vgblnNoEditarPagos = False
                Else
                    .Row = 1
                End If
                txtNombreComercial.Text = IIf(fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)) = "", Trim(.TextMatrix(.Row, 2)), fstrInfoMedicoProcedimiento(grdCargos.RowData(grdCargos.Row)))
                vgblnEditaPago = False
                vgstrEstadoManto = "C"
        End Select
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyDown"))
    Unload Me
End Sub
Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    ' Solo permite números
    If Not fblnFormatoCantidad(txtPrecio, KeyAscii, 6) Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPrecio_KeyPress"))
    Unload Me
End Sub
Sub pMuestraNomComercialCompleto(pgrdGrid As MSHFlexGrid)
    Dim vlintColumna As Integer

    vlintColumna = pgrdGrid.Col
    pgrdGrid.Col = 1
    If Not pgrdGrid.CellFontBold Then
        txtNombreComercial.Text = IIf(pgrdGrid.Name = "grdCargos", IIf(fstrInfoMedicoProcedimiento(pgrdGrid.RowData(pgrdGrid.Row)) = "", Trim(pgrdGrid.TextMatrix(pgrdGrid.Row, 1)), fstrInfoMedicoProcedimiento(pgrdGrid.RowData(pgrdGrid.Row))), Trim(pgrdGrid.TextMatrix(pgrdGrid.Row, 1)))
    Else
        txtNombreComercial.Text = ""
    End If
    pgrdGrid.Col = vlintColumna
    
End Sub
Private Sub txtPrecio_LostFocus()
  If Not vlblnEditando Then
    txtPrecio.Visible = False
  End If
End Sub
Private Function fblnAceptarCargoExcluido(ByRef blnExcluido As Boolean, ByRef strCodigo As String) As Boolean
    Dim frmConf As New frmAutorizacion
    fblnAceptarCargoExcluido = frmConf.fblnAceptarCargoExcluido(blnExcluido, strCodigo)
    lblnExcluir = blnExcluido
    lstrCodigo = strCodigo
End Function
Public Sub pEditarColumnaFecha(KeyAscii As Integer, MkdEdit As MaskEdBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    '-------------------------------------------------------
    'Revisamos si tiene permiso para modificar la fecha
    '-------------------------------------------------------
    If Not fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionHabilitaCambioFecha, "E") Then Exit Sub
    '-------------------------------------------------------
    'Que se salga cuando ya esta facturado
    '-------------------------------------------------------
    If grid.TextMatrix(grid.Row, 0) = "F" Then Exit Sub

    '-------------------------------------------------------
    'Asignamos el valor de la celda
    '-------------------------------------------------------
     MkdEdit.Mask = ""
     MkdEdit.Text = Format(grid.TextMatrix(grid.Row, grid.Col), "dd/mm/yyyy HH:mm")
     MkdEdit.Mask = "##/##/#### ##:##"
    '-------------------------------------------------------
    ' Muestra el maskeditbox en el lugar indicado
    '-------------------------------------------------------
    With grid
        If .CellWidth < 0 Then Exit Sub
            MkdEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    If vgstrEstadoManto = "C" Then vgstrEstadoManto = vgstrEstadoManto & "E"
    MkdEdit.Visible = True
     pEnfocaMkTexto MkdEdit
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumnaFecha"))
    Unload Me
End Sub
Private Sub pSetCellValueColFecha(grid As MSHFlexGrid, MkdEdit As MaskEdBox)
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vlstrFecha As String
    Dim vlstrHora As String
    Dim vllngPersonaGraba As Long
    
    If grid.Col = 5 Then
        If MkdEdit.Visible Then
           'If IsDate(MkdEdit.Text) Then
              vlblnEditando = True
              vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
              vlblnEditando = False
              If vllngPersonaGraba <> 0 Then
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    vlstrFecha = Trim(Mid(MkdEdit.Text, 1, 10))
                    vlstrHora = Trim(Mid(MkdEdit.Text, 11, 6))
              
                    vlstrSentencia = "UPDATE PVCARGO SET DTMFECHAHORA = " & fstrFechaSQL(vlstrFecha, vlstrHora) & ", BITFECHAMANUAL = 1" & _
                                     ", INTCVEEMPLEADOMODIFICOFECHA = " & vllngPersonaGraba & " WHERE INTNUMCARGO = " & grdCargos.RowData(grdCargos.Row)
              
                    pEjecutaSentencia vlstrSentencia
              
                    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "CARGO DIRECTO EN CAJA (CAMBIO DE FECHA)", CStr(grdCargos.RowData(grdCargos.Row)))
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 1)
                    MkdEdit.Visible = False
                    pLlenaCargos
              End If
'           Else
'             '¡Fecha no válida!, formato de fecha dd/mm/yyyy HH:mm
'              MsgBox SIHOMsg(1232), vbOKOnly + vbExclamation, "Mensaje"
'              pEnfocaMkTexto MkdEdit
'           End If
       End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSetCellValueColFecha"))
    Unload Me
    
End Sub





