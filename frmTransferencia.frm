VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransferencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencias"
   ClientHeight    =   7425
   ClientLeft      =   3465
   ClientTop       =   2730
   ClientWidth     =   12045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   12045
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabTransfencia 
      Height          =   8625
      Left            =   -45
      TabIndex        =   23
      Top             =   -345
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   15214
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmTransferencia.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraMaestro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetalle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraNumero"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmTransferencia.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "grdTransferencia"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraNumero 
         Height          =   405
         Left            =   210
         TabIndex        =   42
         Top             =   480
         Width           =   2745
         Begin VB.TextBox txtNumero 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1740
            TabIndex        =   0
            ToolTipText     =   "Consecutivo"
            Top             =   90
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   0
            TabIndex        =   43
            Top             =   150
            Width           =   555
         End
      End
      Begin VB.Frame fraDetalle 
         Height          =   4750
         Left            =   105
         TabIndex        =   41
         Top             =   2080
         Width           =   11925
         Begin VB.CommandButton cmdBorrar 
            Height          =   495
            Left            =   11310
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmTransferencia.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Eliminar de la lista"
            Top             =   4150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VSFlex7LCtl.VSFlexGrid vsfTransferencia 
            Height          =   3900
            Left            =   105
            TabIndex        =   12
            ToolTipText     =   "Formas de pago"
            Top             =   200
            Width           =   11700
            _cx             =   20637
            _cy             =   6879
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12632256
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   12
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTransferencia.frx":01DA
            ScrollTrack     =   -1  'True
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
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   2
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTransferencia 
         Height          =   6450
         Left            =   -74865
         TabIndex        =   28
         Top             =   1170
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   11377
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame3 
         Height          =   780
         Left            =   -74865
         TabIndex        =   37
         Top             =   360
         Width           =   11865
         Begin VB.CommandButton cmdCargar 
            Caption         =   "&Cargar datos"
            Height          =   330
            Left            =   10005
            TabIndex        =   27
            ToolTipText     =   "Cargar la información"
            Top             =   255
            Width           =   1620
         End
         Begin VB.ComboBox cboTipoBus 
            Height          =   315
            Left            =   5925
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Filtro del tipo de movimiento"
            Top             =   270
            Width           =   3870
         End
         Begin MSMask.MaskEdBox mskFechaInicio 
            Height          =   315
            Left            =   705
            TabIndex        =   24
            ToolTipText     =   "Fecha de inicio de la consulta"
            Top             =   255
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   315
            Left            =   2805
            TabIndex        =   25
            ToolTipText     =   "Fecha final de la consulta"
            Top             =   255
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tipo transferencia"
            Height          =   195
            Left            =   4575
            TabIndex        =   40
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2295
            TabIndex        =   39
            Top             =   315
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   135
            TabIndex        =   38
            Top             =   315
            Width           =   465
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2932
         TabIndex        =   36
         Top             =   6900
         Width           =   6270
         Begin VB.CommandButton cmdPrintPoliza 
            Caption         =   "Póliza"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4100
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmTransferencia.frx":02AE
            TabIndex        =   22
            ToolTipText     =   "Imprimir póliza"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   2070
         End
         Begin VB.CommandButton cmdPrint 
            Enabled         =   0   'False
            Height          =   495
            Left            =   3585
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmTransferencia.frx":066C
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Impresión del reporte"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   90
            Picture         =   "frmTransferencia.frx":0D6E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   585
            Picture         =   "frmTransferencia.frx":1170
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Anterior registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1080
            Picture         =   "frmTransferencia.frx":12E2
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Búsqueda"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1590
            Picture         =   "frmTransferencia.frx":1454
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2100
            Picture         =   "frmTransferencia.frx":15C6
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmTransferencia.frx":1AB8
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Grabar"
            Top             =   165
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3090
            Picture         =   "frmTransferencia.frx":1DFA
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancelar"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fraMaestro 
         Height          =   1750
         Left            =   120
         TabIndex        =   29
         Top             =   330
         Width           =   11925
         Begin VB.ComboBox cboDeptoTransfiere 
            Height          =   315
            ItemData        =   "frmTransferencia.frx":22EC
            Left            =   1830
            List            =   "frmTransferencia.frx":22EE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Departamento donde se hizo el movimiento"
            Top             =   600
            Width           =   4185
         End
         Begin VB.ComboBox cboDatoCorteRecibe 
            Height          =   315
            Left            =   7620
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Selección del corte"
            Top             =   1320
            Width           =   4185
         End
         Begin VB.TextBox txtDatoPersonaRecibe 
            Height          =   315
            Left            =   7620
            TabIndex        =   10
            ToolTipText     =   "Consecutivo"
            Top             =   960
            Width           =   4185
         End
         Begin VB.ComboBox cboCorteTransfiere 
            Height          =   315
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   8
            ToolTipText     =   "Selección del corte"
            Top             =   1320
            Width           =   4185
         End
         Begin VB.ComboBox cboDeptoRecibe 
            Height          =   315
            Left            =   7620
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Destino del dinero"
            Top             =   600
            Width           =   4185
         End
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   7620
            TabIndex        =   2
            ToolTipText     =   "Fecha del movimiento"
            Top             =   255
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   3225
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Tipo"
            Top             =   240
            Width           =   2790
         End
         Begin MSMask.MaskEdBox mskFechaIni 
            Height          =   315
            Left            =   1830
            TabIndex        =   6
            ToolTipText     =   "Fecha del movimiento"
            Top             =   945
            Visible         =   0   'False
            Width           =   1550
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaFinal 
            Height          =   315
            Left            =   4440
            TabIndex        =   7
            ToolTipText     =   "Fecha del movimiento"
            Top             =   945
            Visible         =   0   'False
            Width           =   1550
            _ExtentX        =   2725
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lblHasta 
            Caption         =   "Hasta"
            Height          =   180
            Left            =   3700
            TabIndex        =   48
            Top             =   1005
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblDatoPersonaTransfiere 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1830
            TabIndex        =   5
            ToolTipText     =   "Empleado que hizo el movimiento"
            Top             =   945
            Width           =   4185
         End
         Begin VB.Label lblCorteRecibe 
            AutoSize        =   -1  'True
            Caption         =   "Corte "
            Height          =   195
            Left            =   6105
            TabIndex        =   47
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label lblPersonaRecibe 
            AutoSize        =   -1  'True
            Caption         =   "Persona"
            Height          =   195
            Left            =   6105
            TabIndex        =   46
            Top             =   1005
            Width           =   585
         End
         Begin VB.Label lblCorteTransfiere 
            AutoSize        =   -1  'True
            Caption         =   "Corte "
            Height          =   195
            Left            =   105
            TabIndex        =   45
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label lblEstado 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   9435
            TabIndex        =   3
            ToolTipText     =   "Estado de la transferencia"
            Top             =   255
            Width           =   2370
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   8880
            TabIndex        =   35
            Top             =   315
            Width           =   495
         End
         Begin VB.Label lblPersonaTransfiere 
            AutoSize        =   -1  'True
            Caption         =   "Persona"
            Height          =   195
            Left            =   105
            TabIndex        =   34
            Top             =   1005
            Width           =   1695
         End
         Begin VB.Label lblDeptoRecibe 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   6105
            TabIndex        =   33
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label lblDeptoTransfiere 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   660
            Width           =   1005
         End
         Begin VB.Label lblDatoDeptoTransfiere 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1830
            TabIndex        =   44
            ToolTipText     =   "Departamento donde se hizo el movimiento"
            Top             =   600
            Width           =   4185
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   6105
            TabIndex        =   31
            Top             =   315
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   2850
            TabIndex        =   30
            Top             =   315
            Width           =   315
         End
      End
   End
End
Attribute VB_Name = "frmTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
' Transferencias de fondo fijo, a departamentos y a bancos
' Fecha de desarrollo: Viernes, 15 de Junio de 2007
'------------------------------------------------------------------------------------

Option Explicit

'Id's de los tipos de transferencias
Const cintIdFondoFijo = 0
Const cintIdTransDepto = 1
Const cintIdTransBanco = 2
Const cintIdTransFormas = 3

'Grid vsfTransferencia:
Const cintColvsfFecha = 1           'Agregado (CR)
Const cintColvsfTipoDocumento = 2   'Agregado (CR)
Const cintColvsfFolioDocumento = 3  'Agregado (CR)
Const cintColvsfIdForma = 4
Const cintColvsfFormaPago = 5
Const cintColvsfMoneda = 6
Const cintColvsfCtaFuente = 7
Const cintColvsfCtaDestino = 8
Const cintColvsfCantidadFondo = 9
Const cintColvsfTipoCambio = 10
Const cintColvsfReferencia = 11
Const cintColvsfCantidadDisponible = 12
Const cintColvsfTransferir = 13
Const cintColvsfCantidadTransferir = 14
Const cintColvsfCveFormaDestino = 15
Const cintColvsfCantidadTransferida = 16
Const cintColvsfCantidadRecibida = 17
'AAT
Const cintColvsfBanco = 18
Const cintColvsfIdBanco = 19
Const cintColvsfIdRow = 20
Const cintvsfTransferenciaCols = 21

'AAT
Const cstrFormatvsf = "|Fecha|Tipo|Folio|IdForma|Forma de pago|Moneda|CtaFuente|CtaDestino|Cantidad fondo|TipoCambio|Referencia|Cantidad disponible|Transferir|Cantidad transferir|CveFormaDestino|Cantidad transferida|Cantidad recibida|Cuenta bancaria"

'Grid grdTransferencia:
Const cintColgrdFecha = 1
Const cintColgrdId = 2
Const cintColgrdTipo = 3
Const cintColgrdEmpleado = 4
Const cintColgrdEmpleadoCancela = 5
Const cintColgrdEstado = 6
Const cintgrdTransferenciaCols = 7
Const cstrFormatgrd = "|Fecha|Número|Tipo|Empleado registró|Empleado que canceló|Estado"

Const cstrNumero = "###########.##"

Const llngColorCanceladas = &HC0&
Const llngColorActivas = &H80000012

Dim rsDepartamentos As New ADODB.Recordset  'Para cargar combo destino de departamentos
Dim rsBancos As New ADODB.Recordset         'Para cargar combo destino de bancos
Dim rsFormas As New ADODB.Recordset         'Para cargar las formas de pago para el registro de fondo fijo
Dim rsFormasPago As New ADODB.Recordset     'Para cargar las formas de pago para el registro de intercambio entre formas de pago

Dim rsTransferencia As New ADODB.Recordset  'Cuando se consulta una transferencia
Dim rs As New ADODB.Recordset               'Varios usos
Dim rsFormaEquivalente As New ADODB.Recordset

Dim llngForma As Integer                    'Para identificar  la forma de pago a excluir
Dim lstrFormasPago As String                'Cadena de las formas de pago

Dim lblnConsulta As Boolean                 'Para saber cuando se esta haciendo una consulta de una transferencia
Dim llngIdTransferencia As Long             'Id de las transferencia guardada
Dim llngPersonaGraba As Long                'Id persona guarda
Dim ldblTipoCambioVenta As Double

Dim lblnPermisoFondo As Boolean             'Para saber si puede guardar fondo fijo
Dim lblnPermisoTranDepto As Boolean         'Para saber si puede guardar transferencias a departamento
Dim lblnPermisoTranBanco As Boolean         'Para saber si puede guardar transferencias a bancos
Dim lblnPermisoCambioformas As Boolean      'Para saber si puede guardar cambio entre formas de pago

Dim llngNumCorte As Long                    'Para saber en que corte se está guardando
Dim llngEstadoCorte As Long                 'Para el estado del corte
Dim lblnCorteValido As Boolean              'Para saber si se continua al entrar a la pantalla

Dim llngCtaFormaFuente As Long              'Cta. contable de la forma de pago fuente
Dim llngCtaFormaDestino As Long             'Cta. contable de la forma de pago destino
Dim ldblMontoBanco As Double                'Cantidad que se transferirá al banco
Dim ldblMontoTotal As Double                'Cantidad total que afecta una forma de pago

'---------- (CR) - AGREGADOS PARA PERMITIR REALIZAR CAMBIOS DE FORMAS DE CORTES CERRADOS ----------'
Dim llngNumOpcionSelDeptoCorte  As Long     'Indica el número de la opción para seleccionar el departamento y el corte de caja
Dim lintNumeroDepartamento As Integer       'Indica el departamento de dónde se realiza la transferencia
Dim lblnPermisoCambioFormasDepto As Boolean 'Para saber si puede seleccionar el departamento y el corte para el cambio de formas de pago
Dim llngNumPoliza As Long                   'Indica el número de la póliza generada al guardar la transferencia
Dim lstrSentencia As String                 'Para formar las instrucciones SQL

Private Type RegistroPoliza
    vllngNumeroCuenta As Long
    vldblCantidadMovimiento As Double
    vlintTipoMovimiento As Integer
End Type
Dim apoliza() As RegistroPoliza             'Para el registro de las pólizas de movimientos generados en cortes cerrados

Dim vllngNumPoliza As Long
Dim vlintTipoCorte As Integer

'AAT
Dim itemsCombo As String
'--------------------------------------------------------------------------------------------------'

Private Sub cboCorteTransfiere_Click()
On Error GoTo NotificaError

    '(CR) Caso 7442 - Modificado para que tome en cuenta la opción "Cambio de formas de pago"
    'If cboCorteTransfiere.ListIndex <> -1 And cboTipo.ItemData(cboTipo.ListIndex) <> cintIdTransFormas And cboTipo.ItemData(cboTipo.ListIndex) <> cintIdFondoFijo Then
    If cboCorteTransfiere.ListIndex <> -1 Then
        If cboTipo.ItemData(cboTipo.ListIndex) <> cintIdTransFormas And cboTipo.ItemData(cboTipo.ListIndex) <> cintIdFondoFijo Then
            pConfiguraVsf "TD"
            If Not lblnConsulta Then
                pCargaDinero cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex), -1, 2, 0
                    
                If Trim(vsfTransferencia.TextMatrix(1, cintColvsfFormaPago)) = "" Then
                    'No existe dinero disponible para realizar una transferencia.
                    MsgBox SIHOMsg(785), vbOKOnly + vbExclamation, "Mensaje"
                End If
    
                vsfTransferencia.Col = cintColvsfTransferir
                vsfTransferencia.Row = 1
            End If
            
        ElseIf lblnPermisoCambioFormasDepto And cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
            pConfiguraVsf "FP"
            If Not lblnConsulta Then
                pCargaFormasPago cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex), CLng(lintNumeroDepartamento), -1
            
                If Trim(vsfTransferencia.TextMatrix(1, cintColvsfFormaPago)) = "" Then
                    'No existe dinero disponible para realizar una transferencia.
                    MsgBox SIHOMsg(785), vbOKOnly + vbExclamation, "Mensaje"
                End If
            
                vsfTransferencia.Col = cintColvsfTransferir
                vsfTransferencia.Row = 1
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCorteTransfiere_Click"))
End Sub

Private Sub cboCorteTransfiere_GotFocus()
On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboCorteTransfiere_GotFocus"))
End Sub

Private Sub cboDeptoRecibe_Click()
On Error GoTo NotificaError

    If cboDeptoRecibe.ListIndex <> -1 Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And Not lblnConsulta Then
            pCargaEquivalencia CLng(lintNumeroDepartamento), cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)
        End If
        
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas And Not lblnConsulta Then
            pCargaReferencia cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDeptoRecibe_Click"))
End Sub

Private Sub pCargaEquivalencia(lngCveDeptoFuente As Long, lngCveDeptoDestino As Long)
On Error GoTo NotificaError

    vgstrParametrosSP = CStr(lngCveDeptoFuente) & "|" & CStr(lngCveDeptoDestino)
    Set rsFormaEquivalente = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormaEquivalente")

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaEquivalencia"))
End Sub

Private Sub pCargaReferencia(lngCveforma As Long)
On Error GoTo NotificaError

    Dim rsFormaDatos As New ADODB.Recordset
    Dim rsBancosForma As New ADODB.Recordset

    vgstrParametrosSP = CStr(lngCveforma) & "|-1|-1|-1|1|" & "*"
    Set rsFormaDatos = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormapago")
    If rsFormaDatos.RecordCount > 0 Then
        If rsFormaDatos!bitpreguntafolio = 1 Then
            txtDatoPersonaRecibe.Enabled = True
            lblPersonaRecibe.Enabled = True
        End If
        
        'Si es de tipo transferencia a banco, llenar combo de bancos
        If rsFormaDatos!chrTipo = "B" Then
            cboDatoCorteRecibe.Enabled = True
            lblCorteRecibe.Enabled = True
            If rsBancos.RecordCount <> 0 Then
                pLlenarCboRs cboDatoCorteRecibe, rsBancos, 4, 5
            Else
                'No existen bancos activos, selecciones otra forma de pago.
                MsgBox SIHOMsg(203) & "" & SIHOMsg(934), vbOKOnly + vbExclamation, "Mensaje"
            End If
        End If
        
        pConfiguraVsf "FP"
        llngForma = rsFormaDatos!intFormaPago
        
        If cboCorteTransfiere.ListIndex = -1 Then
            MsgBox "No se ha seleccionado un corte.", vbOKOnly + vbExclamation, "Mensaje"
            cboCorteTransfiere.SetFocus
        Else
            pCargaFormasPago cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex), CLng(lintNumeroDepartamento), rsFormaDatos!intFormaPago
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaReferencia"))
End Sub

Private Sub cboDeptoRecibe_GotFocus()
On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboDeptoRecibe_GotFocus"))
End Sub

Private Sub cboDeptoTransfiere_Click()
    If cboDeptoTransfiere.ListIndex <> -1 And cboTipo.ListIndex <> -1 Then
        If Not lblnConsulta Then
            lintNumeroDepartamento = cboDeptoTransfiere.ItemData(cboDeptoTransfiere.ListIndex)
            If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
                pConfiguraVsf "FP"
                pCargaCortesDepto True
                
                '- Cargar formas de pago del departamento seleccionado -'
                vgstrParametrosSP = CStr(-1) & "|" & "-1" & "|" & CStr(-1) & "|" & CStr(lintNumeroDepartamento) & "|" & "1" & "|" & "*"
                Set rsFormasPago = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormapago")
                pCargaFormas
            End If
        End If
    End If
End Sub
'AAT
Private Sub cboTipo_Change()
    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        cboDeptoRecibe.Enabled = False
    Else
        cboDeptoRecibe.Enabled = True
    End If
End Sub

Private Sub cboTipo_Click()
On Error GoTo NotificaError

    Dim llngNumeroCorte As Long
    Dim rsCorte As New ADODB.Recordset
    
    'cboDeptoRecibe.Visible = True

    If cboTipo.ListIndex <> -1 Then
        'AAT
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
            cboDeptoRecibe.Enabled = False
        Else
            cboDeptoRecibe.Enabled = True
        End If
    
        pCargaDeptoTransfiere False 'Busca el departamento del usuario que ingresó
        pHabilitaControles (lblnPermisoCambioFormasDepto And cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas) 'Deshabilitar controles
        pOcultaControles (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas) 'Ocultar el rango de fecha
    
        '----- FONDO FIJO -----'
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
            pTitulo cintIdFondoFijo
            If Not lblnConsulta Then
                pConfiguraVsf "FF"
                cboCorteTransfiere.Enabled = False
                pHabilitaTitulo 0, 0, 0, 0, 0, 0
                
                mskFecha.Mask = ""
                mskFecha.Text = fdtmServerFecha
                mskFecha.Mask = "##/##/####"
                
                'lblDatoDeptoTransfiere.Caption = vgstrNombreDepartamento
                lblDatoDeptoTransfiere.Caption = cboDeptoTransfiere.Text
                lblDatoPersonaTransfiere.Caption = ""
                
                cboCorteTransfiere.Clear
                
                cboDeptoRecibe.Clear
                cboDeptoRecibe.Enabled = False
                
                txtDatoPersonaRecibe.Text = ""
                txtDatoPersonaRecibe.Enabled = False
                
                pCargaCorte
                
                cboDatoCorteRecibe.Clear
                cboDatoCorteRecibe.Enabled = False
                vsfTransferencia.Col = cintColvsfFormaPago
                vsfTransferencia.Row = 1
            End If
            
        '----- TRANSFERENCIA A DEPARTAMENTO -----'
        ElseIf cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Then
            pTitulo cintIdTransDepto
            If Not lblnConsulta Then
                pConfiguraVsf "TD"
                pHabilitaTitulo 0, 0, 1, 1, 0, 0
                
                mskFecha.Mask = ""
                mskFecha.Text = fdtmServerFecha
                mskFecha.Mask = "##/##/####"
                
                cboCorteTransfiere.Enabled = True
                cboDeptoRecibe.Enabled = True
                
                'lblDatoDeptoTransfiere.Caption = vgstrNombreDepartamento
                lblDatoDeptoTransfiere.Caption = cboDeptoTransfiere.Text
                lblDatoPersonaTransfiere.Caption = ""
                
                pCargaCorte
                pCargaDeptos
                
                txtDatoPersonaRecibe.Text = ""
                txtDatoPersonaRecibe.Enabled = False
                cboDatoCorteRecibe.Clear
                cboDatoCorteRecibe.Enabled = False
                lblCorteTransfiere.Enabled = False
            End If
            
        '----- TRANSFERENCIA A BANCO -----'
        'AAT
        ElseIf cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
            
            pTitulo cintIdTransBanco
            
            If Not lblnConsulta Then
                'AAT
                pCargaBancosVsf
                pConfiguraVsf "TB"
                
                pHabilitaTitulo 0, 0, 0, 1, 0, 0
                                
                cboCorteTransfiere.Clear
                cboCorteTransfiere.Enabled = False
                
                'lblDatoDeptoTransfiere.Caption = vgstrNombreDepartamento
                lblDatoDeptoTransfiere.Caption = cboDeptoTransfiere.Text
                lblDatoPersonaTransfiere.Caption = ""
                
                'AAT
                'pCargaBancos
                
                txtDatoPersonaRecibe.Text = ""
                txtDatoPersonaRecibe.Enabled = False
                cboDatoCorteRecibe.Clear
                cboDatoCorteRecibe.Enabled = False
                'AAT
                cboDeptoRecibe.Enabled = False
                
                pCargaDinero -1, CLng(lintNumeroDepartamento), 1, 1
                
                If Trim(vsfTransferencia.TextMatrix(1, cintColvsfFormaPago)) = "" Then
                    'No existe dinero disponible para realizar una transferencia.
                    MsgBox SIHOMsg(785), vbOKOnly + vbExclamation, "Mensaje"
                End If
                
                vsfTransferencia.Col = cintColvsfTransferir
                vsfTransferencia.Row = 1
            End If
            
        '----- CAMBIO DE FORMAS DE PAGO -----'
        ElseIf cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
            pTitulo cintIdTransFormas
            If Not lblnConsulta Then
                pConfiguraVsf "FP" 'Se utilizan las mismas columnas que para la transferencia a bancos
                pHabilitaTitulo 0, 0, 0, 1, 0, 0

                'lblDatoDeptoTransfiere.Caption = vgstrNombreDepartamento
                lblDatoDeptoTransfiere.Caption = cboDeptoTransfiere.Text
                lblDatoPersonaTransfiere.Caption = ""

                pCargaFormas
                
                txtDatoPersonaRecibe.Text = ""
                txtDatoPersonaRecibe.Enabled = False
                cboDatoCorteRecibe.Clear
                cboDatoCorteRecibe.Enabled = False
                
                '-------------------------------------------- CAMBIOS PARA CASO 7442 --------------------------------------------'
                '-- Validación para verificar que el usuario tenga permisos, si los tiene, permitir seleccionar depto. y corte --'
                If lblnPermisoCambioFormasDepto Then
                    pHabilitaControles True
                    pCargaCortesDepto True
                Else
                '-- Si no tiene permisos, inhabilitar selección de departamento y de corte y mostrar solo el corte activo --'
                    llngNumeroCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
                    pCargaFormasPago llngNumeroCorte, CLng(lintNumeroDepartamento), -1
                    
                    cboCorteTransfiere.Clear
                    vgstrParametrosSP = llngNumeroCorte & "|" & "0|' '|' '" & "|" & -1 & "|" & -1
                    Set rsCorte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelCorte")
                    If rsCorte.RecordCount > 0 Then
                        cboCorteTransfiere.AddItem CStr(rsCorte!IdCorte) & " - " & Format(rsCorte!FechaAbre, "dd/mmm/yyyy hh:mm") & " - " & UCase(rsCorte!Estado)
                        cboCorteTransfiere.ItemData(cboCorteTransfiere.newIndex) = rsCorte!IdCorte
                    End If
                    cboCorteTransfiere.ListIndex = 0
                    cboCorteTransfiere.Enabled = False
                    
                    If Trim(vsfTransferencia.TextMatrix(1, cintColvsfFormaPago)) = "" Then
                        'No existe dinero disponible para realizar una transferencia.
                        MsgBox SIHOMsg(785), vbOKOnly + vbExclamation, "Mensaje"
                    End If
                End If
                '----------------------------------------------------------------------------------------------------------------'
                
                vsfTransferencia.Col = cintColvsfTransferir
                vsfTransferencia.Row = 1
            End If
        End If
        
        lblFecha.Enabled = Not (cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto)
        mskFecha.Enabled = Not (cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto)
    Else
        pHabilitaTitulo 1, 1, 1, 1, 1, 1
    End If

    fraDetalle.Enabled = cboTipo.ListIndex <> -1

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipo_Click"))
End Sub

Private Sub pTitulo(intTipo As Integer)
On Error GoTo NotificaError

    lblDeptoTransfiere.Caption = IIf(intTipo = cintIdTransDepto, "Departamento transfiere", "Departamento registra")
    'lblPersonaTransfiere.Caption = IIf(intTipo = cintIdTransDepto, "Persona transfiere", "Persona registra")
    Select Case intTipo
        Case cintIdTransDepto: lblPersonaTransfiere.Caption = "Persona transfiere"
        Case cintIdTransFormas: lblPersonaTransfiere.Caption = "Rango de cortes"
        Case Else: lblPersonaTransfiere.Caption = "Persona registra"
    End Select
    lblCorteTransfiere.Caption = IIf(intTipo = cintIdTransDepto, "Corte transfiere", IIf(intTipo = cintIdFondoFijo, "Corte recibe", "Corte"))
        
    lblDeptoRecibe.Caption = IIf(intTipo = cintIdFondoFijo, "Departamento", IIf(intTipo = cintIdTransDepto, "Departamento recibe", IIf(intTipo = cintIdTransFormas, "Forma de pago", "Banco recibe")))
    lblPersonaRecibe.Caption = IIf(intTipo = cintIdTransDepto, "Persona recibe", IIf(intTipo = cintIdTransFormas, "Referencia", "Persona"))
    lblCorteRecibe.Caption = IIf(intTipo = cintIdTransDepto, "Corte recibe", IIf(intTipo = cintIdTransFormas, "Banco", "Corte"))

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pTitulo"))
End Sub

Private Sub pCargaCorte()
On Error GoTo NotificaError

    Dim vlstrCargarRs As String
    Dim rsCorteCerrado As New ADODB.Recordset

    cboCorteTransfiere.Clear
    Set rs = frsEjecuta_SP(CStr(lintNumeroDepartamento) & "|" & CStr(vglngNumeroEmpleado), "Sp_PvSelCorteDepto")
    If rs.RecordCount <> 0 Then
        vlintTipoCorte = rs!TipoCorte
        Do While Not rs.EOF
            vlstrCargarRs = "SELECT * FROM PvCorteCerrado WHERE PvCorteCerrado.intNumCorte =  " & rs!intnumcorte
            Set rsCorteCerrado = frsRegresaRs(vlstrCargarRs, adLockReadOnly, adOpenForwardOnly)
            If rsCorteCerrado.EOF Then
                cboCorteTransfiere.AddItem CStr(rs!intnumcorte) & " - " & Format(rs!dtmFechahora, "dd/mmm/yyyy hh:mm") & " - " & IIf(IsNull(rs!dtmFechaRegistro), "ABIERTO", "CERRADO")
                cboCorteTransfiere.ItemData(cboCorteTransfiere.newIndex) = rs!intnumcorte
            End If
            rs.MoveNext
        Loop
        cboCorteTransfiere.ListIndex = 0
        cboCorteTransfiere.Enabled = False
        'cboCorteTransfiere.ListIndex = 0
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCorte"))
End Sub

Private Sub pCargaFormasPago(lngNumCorte As Long, lngNumDepto As Long, lngformaexcluir As Integer)
On Error GoTo NotificaError

    vgstrParametrosSP = CStr(lngNumCorte) & "|" & CStr(lngNumDepto) & "|" & lngformaexcluir
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormasPagoCorte")
    If rs.RecordCount > 0 Then
        With vsfTransferencia
            Do While Not rs.EOF
                '-- Verificar que NO se haya realizado una transferencia previa con la forma de pago --'
                If Not fblnTransferida(lngNumCorte, rs!CveForma, rs!cantidad) Then
                    .TextMatrix(.Rows - 1, cintColvsfIdForma) = rs!CveForma
                    .TextMatrix(.Rows - 1, cintColvsfFormaPago) = rs!formapago
                    .TextMatrix(.Rows - 1, cintColvsfMoneda) = rs!PESOS
                    .TextMatrix(.Rows - 1, cintColvsfCtaFuente) = rs!IdCtaContable
                    .TextMatrix(.Rows - 1, cintColvsfReferencia) = IIf(Trim(rs!Referencia) = "0", " ", rs!Referencia)
                    .TextMatrix(.Rows - 1, cintColvsfCantidadDisponible) = FormatCurrency(rs!cantidad, 2)
                    .TextMatrix(.Rows - 1, cintColvsfCveFormaDestino) = rs!chrTipo
                    '-------------------------- (CR) Agregados Caso 7442 --------------------------'
                    .TextMatrix(.Rows - 1, cintColvsfFecha) = Format(rs!fecha, "dd/mmm/yyyy hh:nn")
                    .TextMatrix(.Rows - 1, cintColvsfTipoDocumento) = rs!TipoDocumento
                    .TextMatrix(.Rows - 1, cintColvsfFolioDocumento) = rs!folio
                    '------------------------------------------------------------------------------'
                    .Rows = .Rows + 1
                End If
                rs.MoveNext
            Loop

            If .Rows > 2 Then .Rows = .Rows - 1 'Borra el último renglón vacío
        End With
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaFormasPago"))
End Sub

Private Sub pCargaDinero(lngNumCorte As Long, lngNumDepto As Long, intincluirtarjeta As Integer, intTRANSFUERACORTE As Integer)
On Error GoTo NotificaError

    vgstrParametrosSP = CStr(lngNumCorte) & "|" & "-1" & "|" & CStr(lngNumDepto) & "|" & "'-1'" & "|" & intincluirtarjeta & "|" & intTRANSFUERACORTE
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelDineroCorte")
    If rs.RecordCount Then
        With vsfTransferencia
            Do While Not rs.EOF
                '- CASO 7723: Revisar que la cuenta contable de la forma de pago NO pertenezca a un banco -'
                If flngEsCuentaBanco(rs!IdCtaContable) = 0 Then
                    If FormatCurrency(rs!cantidad, 2) <> "$0.00" Then
                        .TextMatrix(.Rows - 1, cintColvsfIdForma) = rs!CveForma
                        .TextMatrix(.Rows - 1, cintColvsfFormaPago) = rs!formapago
                        .TextMatrix(.Rows - 1, cintColvsfMoneda) = rs!PESOS
                        .TextMatrix(.Rows - 1, cintColvsfCtaFuente) = rs!IdCtaContable
                        .TextMatrix(.Rows - 1, cintColvsfReferencia) = IIf(Trim(rs!Referencia) = "0", " ", rs!Referencia)
                        .TextMatrix(.Rows - 1, cintColvsfCantidadDisponible) = FormatCurrency(rs!cantidad, 2)
                        .Rows = .Rows + 1
                    End If
                End If
                rs.MoveNext
            Loop
            If .Rows > 2 Then .Rows = .Rows - 1
        End With
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDinero"))
End Sub

Private Sub pCargaDeptos()
On Error GoTo NotificaError

    cboDeptoRecibe.Clear
    If rsDepartamentos.RecordCount <> 0 Then
        rsDepartamentos.MoveFirst
        Do While Not rsDepartamentos.EOF
            If rsDepartamentos!clave <> lintNumeroDepartamento Then
                cboDeptoRecibe.AddItem rsDepartamentos!Descripcion
                cboDeptoRecibe.ItemData(cboDeptoRecibe.newIndex) = rsDepartamentos!clave
            End If
            rsDepartamentos.MoveNext
        Loop
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDeptos"))
End Sub

Private Sub pCargaFormas()
On Error GoTo NotificaError
    Dim strTipo As String 'Se usa para ver si se van a excluir las formas de pago transferencia a banco
   
    If rsBancos.RecordCount > 0 Then
        strTipo = "*"
    Else
        strTipo = "B"
    End If

    cboDeptoRecibe.Clear
    If rsFormasPago.RecordCount <> 0 Then
        rsFormasPago.MoveFirst
        Do While Not rsFormasPago.EOF
            If rsFormasPago!bitestatusactivo = 1 And rsFormasPago!chrTipo <> "C" And rsFormasPago!chrTipo <> strTipo Then
                cboDeptoRecibe.AddItem rsFormasPago!chrdescripcion
                cboDeptoRecibe.ItemData(cboDeptoRecibe.newIndex) = rsFormasPago!intFormaPago
            End If
            rsFormasPago.MoveNext
        Loop
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaFormas"))
End Sub

Private Sub pCargaBancos()
On Error GoTo NotificaError

    cboDeptoRecibe.Clear
    If rsBancos.RecordCount <> 0 Then
        rsBancos.MoveFirst
        Do While Not rsBancos.EOF
            If rsBancos!BITESTATUS = 1 Then
                cboDeptoRecibe.AddItem rsBancos!VCHNOMBREBANCO
                cboDeptoRecibe.ItemData(cboDeptoRecibe.newIndex) = rsBancos!tnynumerobanco
            End If
            rsBancos.MoveNext
        Loop
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaBancos"))
End Sub

    
Private Sub pConfiguraVsf(strMovto As String)
On Error GoTo NotificaError
' strMovto = II = Cuando INICIA la pantalla
' strMovto = FF = Fondo fijo
' strMovto = TD = Transferencia departamento
' strMovto = TB = Transferencia banco
' strMovto = FP = Cambio de formas de pago
' strMovto = RD = Recepciones de dinero
' strMovto = CF = Consulta fondo fijo
' strMovto = CT = Consulta transferencia departamento
' strMovto = CB = Consulta transferencia banco
' strMovto = CF = Consulta cambio de formas de pago
' strMovto = CR = Consulta recepciones

    With vsfTransferencia
        .Visible = False
        
        .Clear
        .Rows = 2
        .Cols = cintvsfTransferenciaCols
        .FormatString = cstrFormatvsf
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(cintColvsfFecha) = IIf(strMovto = "FP" And Not lblnConsulta, 1530, 0)  'Agregado (CR)
        .ColWidth(cintColvsfTipoDocumento) = IIf(strMovto = "FP", 1250, 0)  'Agregado (CR)
        .ColWidth(cintColvsfFolioDocumento) = IIf(strMovto = "FP", 1150, 0) 'Agregado (CR)
        .ColWidth(cintColvsfIdForma) = 0
        'AAT
        '.ColWidth(cintColvsfFormaPago) = IIf(strMovto = "FP", 3700, 5500)
        .ColWidth(cintColvsfFormaPago) = IIf(strMovto = "FP", 3700, 3700)
        .ColWidth(cintColvsfMoneda) = 0
        .ColWidth(cintColvsfCtaFuente) = 0
        .ColWidth(cintColvsfCtaDestino) = 0
        .ColWidth(cintColvsfCantidadFondo) = IIf(strMovto = "FF", 1400, 0)
        .ColWidth(cintColvsfTipoCambio) = 0
        .ColWidth(cintColvsfReferencia) = IIf(strMovto <> "FF" And strMovto <> "II", 1200, 0)
        'AAT
        .ColWidth(cintColvsfCantidadDisponible) = IIf(strMovto = "TD" Or strMovto = "TB" Or strMovto = "FP", 1500, 0)
        .ColWidth(cintColvsfTransferir) = IIf(strMovto = "TD" Or strMovto = "TB" Or strMovto = "FP", 950, 0)
        '----- MODIFICADO (CR) -----'
        '.ColWidth(cintColvsfCantidadTransferir) = IIf(strMovto = "TD" Or strMovto = "TB" Or strMovto = "FP", 1400, 0)
        .ColWidth(cintColvsfCantidadTransferir) = IIf(strMovto = "TD" Or strMovto = "TB", 1400, 0)
        '---------------------------'
        .ColWidth(cintColvsfCveFormaDestino) = 0
        .ColWidth(cintColvsfCantidadTransferida) = IIf(strMovto = "CT" Or strMovto = "CB" Or strMovto = "CF", 1500, 0)
        .ColWidth(cintColvsfCantidadRecibida) = IIf(strMovto = "CR", 1400, 0)
        'AAT
        .ColWidth(cintColvsfBanco) = IIf(strMovto = "TB" Or strMovto = "CB", 3900, 0)
        .ColWidth(cintColvsfIdBanco) = 0 'Se oculta columna, se usa para obtener ID banco al generar poliza en Guarda
        .ColWidth(cintColvsfIdRow) = 0 'Se oculta columna, se usa para guardar el ID de la fila PADRE para las transferencias en EFECTIVO
        
        .ColAlignment(cintColvsfFecha) = flexAlignLeftCenter                'Agregado (CR)
        .ColAlignment(cintColvsfTipoDocumento) = flexAlignLeftCenter        'Agregado (CR)
        .ColAlignment(cintColvsfFolioDocumento) = flexAlignLeftCenter       'Agregado (CR)
        .ColAlignment(cintColvsfFormaPago) = flexAlignLeftCenter
        .ColAlignment(cintColvsfCantidadFondo) = flexAlignRightCenter
        .ColAlignment(cintColvsfReferencia) = flexAlignRightCenter
        .ColAlignment(cintColvsfCantidadDisponible) = flexAlignRightCenter
        .ColAlignment(cintColvsfTransferir) = flexAlignCenterCenter
        .ColAlignment(cintColvsfCantidadTransferir) = flexAlignRightCenter
        .ColAlignment(cintColvsfCantidadTransferida) = flexAlignRightCenter
        .ColAlignment(cintColvsfCantidadRecibida) = flexAlignRightCenter
        'AAT
        .ColAlignment(cintColvsfBanco) = flexAlignLeftCenter
        
        .FixedAlignment(cintColvsfFecha) = flexAlignCenterCenter            'Agregado (CR)
        .FixedAlignment(cintColvsfTipoDocumento) = flexAlignCenterCenter    'Agregado (CR)
        .FixedAlignment(cintColvsfFolioDocumento) = flexAlignCenterCenter   'Agregado (CR)
        .FixedAlignment(cintColvsfFormaPago) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfCantidadFondo) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfReferencia) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfCantidadDisponible) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfTransferir) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfCantidadTransferir) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfCantidadTransferida) = flexAlignCenterCenter
        .FixedAlignment(cintColvsfCantidadRecibida) = flexAlignCenterCenter
        'AAT
        .FixedAlignment(cintColvsfBanco) = flexAlignCenterCenter
        
        .ColDataType(cintColvsfTransferir) = flexDTBoolean
        .ColDataType(cintColvsfFecha) = flexDTDate 'Agregado (CR)
        'AAT
        .ColDataType(cintColvsfBanco) = flexDTString
        
        .EditMaxLength = 10 'Agregado (CR)
        If strMovto = "FP" Then .TextMatrix(0, cintColvsfTransferir) = "Cambiar" 'Agregado (CR)
        
        .Visible = True
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraVsf"))
End Sub

Private Sub cboTipo_GotFocus()
On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipo_GotFocus"))
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError

    If grdTransferencia.Row > 1 Then grdTransferencia.Row = grdTransferencia.Row - 1
    pMuestra grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)
    pHabilita 1, 1, 1, 1, 1, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf(rsTransferencia!Estado = "R", 1, 0)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub

Private Sub cmdBorrar_Click()
On Error GoTo NotificaError

    vsfTransferencia.RemoveItem vsfTransferencia.Row
    vsfTransferencia_RowColChange

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBorrar_Click"))
End Sub

Private Sub cmdCargar_Click()
On Error GoTo NotificaError

    If Not IsDate(mskFechaInicio.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        mskFechaInicio.SetFocus
    Else
        If Not IsDate(mskFechaFin.Text) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            mskFechaFin.SetFocus
        Else
            If CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
                '¡Rango de fechas no válido!
                MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
                mskFechaFin.SetFocus
            Else
                pCarga True
                If Val(grdTransferencia.TextMatrix(1, cintColgrdId)) = 0 Then
                    'No encontró datos:
                    mskFechaInicio.SetFocus
                Else
                    grdTransferencia.SetFocus
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargar_Click"))
End Sub

Private Sub pCarga(blnMensaje As Boolean)
On Error GoTo NotificaError

    Dim intcontador As Integer

    pConfiguraGrd
    
    vgstrParametrosSP = CStr(-1) & "|" & fstrFechaSQL(mskFechaInicio.Text) & "|" & fstrFechaSQL(mskFechaFin.Text) & "|" & "1" & "|" & CStr(lintNumeroDepartamento) & "|" & CStr(cboTipoBus.ItemData(cboTipoBus.ListIndex))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelTransferencia")
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
            If Not fblnEsta(rs!IdTransferencia) Then
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdFecha) = Format(rs!fecha, "dd/mmm/yyyy")
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdId) = rs!IdTransferencia
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdTipo) = rs!tipo
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdEmpleado) = rs!EmpleadoRegistra
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdEmpleadoCancela) = rs!EmpleadoCancela
                grdTransferencia.TextMatrix(grdTransferencia.Rows - 1, cintColgrdEstado) = rs!TipoEstado
                If Trim(rs!Estado) = "C" Then
                    For intcontador = 1 To grdTransferencia.Cols - 1
                        grdTransferencia.Row = grdTransferencia.Rows - 1
                        grdTransferencia.Col = intcontador
                        grdTransferencia.CellForeColor = llngColorCanceladas
                    Next intcontador
                End If
                
                grdTransferencia.Rows = grdTransferencia.Rows + 1
            End If
            rs.MoveNext
        Loop
        grdTransferencia.Rows = grdTransferencia.Rows - 1
        grdTransferencia.Col = cintColgrdFecha
        grdTransferencia.Row = 1
    Else
        If blnMensaje Then
            'No existe información en ese rango de fechas.
            MsgBox SIHOMsg(719), vbInformation + vbOKOnly, "Mensaje"
        End If
    End If
    rs.Close

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCarga"))
End Sub

Private Function fblnEsta(lngIdTransferencia As Long) As Boolean
On Error GoTo NotificaError
    Dim intcontador As Integer
    
    fblnEsta = False
    intcontador = 1
    Do While intcontador <= grdTransferencia.Rows - 1 And Not fblnEsta
        fblnEsta = Val(grdTransferencia.TextMatrix(intcontador, cintColgrdId)) = lngIdTransferencia
        intcontador = intcontador + 1
    Loop

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnEsta"))
End Function

Private Sub cmdDelete_Click()
On Error GoTo NotificaError

    If fblnDatosCancelar Then
        If rsTransferencia!IdTipo = cintIdFondoFijo Then
            pCancelaFondo
        End If
        If rsTransferencia!IdTipo = cintIdTransDepto Then
            pCancelaTransDepto
        End If
        If rsTransferencia!IdTipo = cintIdTransBanco Then
            pCancelaTransBanco
        End If
        If rsTransferencia!IdTipo = cintIdTransFormas Then
            pCancelaTransFormas
        End If
        
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub pCancelaTransFormas()
On Error GoTo NotificaError

    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    Dim intcontador As Integer

    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If llngNumCorte = rsTransferencia!IdCorte Then
        If fblnCorteLibre(llngNumCorte) Then
            If fblnTransActiva Then
                 pPosicionaForma
                 
                '1.- Cancelar en maestro y poner en la tabla de cancelados
                vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & CStr(llngPersonaGraba)
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaTransferencia"
                
                '2.- Insertar las formas de pago en el corte en forma positiva:
                ldblMontoBanco = 0
                With vsfTransferencia
                    For intcontador = 1 To .Rows - 1
                        dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferida), cstrNumero))
                        dblTipoCambio = Val(Format(.TextMatrix(intcontador, cintColvsfTipoCambio), cstrNumero))
                        ldblMontoBanco = ldblMontoBanco + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                        
                        vgstrParametrosSP = CStr(llngNumCorte) _
                                            & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                            & "|" & CStr(rsTransferencia!IdTransferencia) _
                                            & "|" & "TR" & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                            & "|" & CStr(dblCantidad) _
                                            & "|" & CStr(dblTipoCambio) _
                                            & "|" & .TextMatrix(intcontador, cintColvsfReferencia) & "|" & CStr(rsTransferencia!IdCorte)
                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                    
                        'CARGO al a cuenta de la forma de pago FUENTE:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                                            & "|" & CStr(rsTransferencia!IdTransferencia) _
                                            & "|" & "TR" _
                                            & "|" & .TextMatrix(intcontador, cintColvsfCtaFuente) _
                                            & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
                                            & "|" & "1"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                    Next intcontador
                End With
                
                '3.- Insertar la forma de pago maestro en el corte, en forma negativa:
                vgstrParametrosSP = CStr(llngNumCorte) _
                                    & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                    & "|" & CStr(rsTransferencia!IdTransferencia) _
                                    & "|" & "TR" & "|" & rsTransferencia!idformapago _
                                    & "|" & CStr(rsTransferencia!cantidad2 * -1) _
                                    & "|" & CStr(rsTransferencia!tipocambio2) _
                                    & "|" & CStr(rsTransferencia!Referencia) & "|" & CStr(rsTransferencia!IdCorte)
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                    
                'ABONO al a cuenta de la forma de pago maestro:
                vgstrParametrosSP = CStr(llngNumCorte) _
                                    & "|" & CStr(rsTransferencia!IdTransferencia) _
                                    & "|" & "TR" _
                                    & "|" & CStr(rsTransferencia!cuenta) _
                                    & "|" & CStr(rsTransferencia!cantidad2 * IIf(rsTransferencia!tipocambio2 = 0, 1, rsTransferencia!tipocambio2)) _
                                    & "|" & "0"
                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                
                '4.- Afectar el libro de bancos en su moneda, en caso de que la forma de pago maestro sea de tipo Transferencia bancaria
                'Esta variable <ldblMontoBanco> tomo valor en <fblnGuardaForma> al recorrer las formas de pago:
                If rsTransferencia!tipoforma = "B" Then
                    ldblMontoBanco = ldblMontoBanco / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
                    vgstrParametrosSP = CStr(cboDatoCorteRecibe.ItemData(cboDatoCorteRecibe.ListIndex)) & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & "CDE" & "|" & "0" & "|" & CStr(ldblMontoBanco) & "|" & CStr(rsTransferencia!IdTransferencia)
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
                End If
                
                '5.- Liberar el corte:
                pLiberaCorte llngNumCorte
                
                '6.- Afectar registro de transacciones
                pGuardarLogTransaccion Me.Name, EnmBorrar, llngPersonaGraba, "TRANSFERENCIA BANCO", CStr(rsTransferencia!IdTransferencia)
               
                '7.- Cancela los movimientos de la forma de pago de la transferencia
                pCancelaMovimiento rsTransferencia!IdTransferencia, rsTransferencia!IdCorte, llngNumCorte, llngPersonaGraba
               
                EntornoSIHO.ConeccionSIHO.CommitTrans
                'La operación se realizó satisfactoriamente.
                MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
                txtNumero.SetFocus
            End If
        End If
    Else
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        MsgBox SIHOMsg(940), vbInformation + vbOKOnly, "Mensaje"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaTransBanco"))
End Sub

Private Sub pCancelaTransBanco()
On Error GoTo NotificaError

    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    Dim intcontador As Integer
    Dim lngNumDetalle As Long
    'AAT
    Dim strIdBanco As String
    Dim vl_refTR As String
    Dim vl_conTR As String
    Dim vl_RefConcepto As Integer

    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(mskFecha.Text)), Month(CDate(mskFecha.Text))) Then
        'El periodo contable esta cerrado
        MsgBox SIHOMsg(209), vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub
    Else
        If fblnTransActiva Then
            'AAT
            'pPosicionaBanco
        
            '1.- Cancelar en maestro y poner en la tabla de cancelados
            vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & CStr(llngPersonaGraba)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaTransferencia"
            vllngNumPoliza = 0
            lstrConceptoPoliza = ""
            If fblnObtenerConfigPolizanva(0, "1", frmTransferencia.Name) = True Then
                For vg_intIndiceConceptoPoliza = 0 To UBound(ConceptosPoliza, 1)
                    Select Case ConceptosPoliza(vg_intIndiceConceptoPoliza).tipo
                        Case "1"
                            lstrConceptoPoliza = lstrConceptoPoliza & ConceptosPoliza(vg_intIndiceConceptoPoliza).Texto
                        Case "3"
                            lstrConceptoPoliza = lstrConceptoPoliza & txtNumero
                        Case "25"
                            lstrConceptoPoliza = lstrConceptoPoliza & UCase(cboTipo.Text)
                    End Select
                Next vg_intIndiceConceptoPoliza
            Else
                Me.MousePointer = 0
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                Exit Sub
            End If
            vllngNumPoliza = flngInsertarPoliza(CDate(mskFecha.Text), "D", "CANCELACIÓN " & lstrConceptoPoliza, llngPersonaGraba)
            
            '2.- Insertar las formas de pago en el corte en forma positiva:
            ldblMontoBanco = 0
            With vsfTransferencia
                For intcontador = 1 To .Rows - 1
                    For vl_RefConcepto = 1 To 2
                        lstrConceptoDetallePoliza = " "
                        If fblnObtenerConfigPolizanva(vl_RefConcepto, "1", frmTransferencia.Name) = True Then
                            For vg_intIndiceConceptoPoliza = 0 To UBound(ConceptosPoliza, 1)
                                Select Case ConceptosPoliza(vg_intIndiceConceptoPoliza).tipo
                                    Case "1"
                                        lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & ConceptosPoliza(vg_intIndiceConceptoPoliza).Texto
                                    Case "26"
                                            lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & vsfTransferencia.TextMatrix(intcontador, cintColvsfFormaPago)
                                    Case "27"
                                            lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & vsfTransferencia.TextMatrix(intcontador, cintColvsfBanco)
                                End Select
                            Next vg_intIndiceConceptoPoliza
                        End If
                        If vl_RefConcepto = 2 Then vl_refTR = Trim(lstrConceptoDetallePoliza)
                        If vl_RefConcepto = 1 Then vl_conTR = Trim(lstrConceptoDetallePoliza)
                    Next vl_RefConcepto
                    'AAT
                    strIdBanco = Val(.TextMatrix(intcontador, cintColvsfIdBanco))
                    'se busca info del banco para obtener la cuenta y/o bitestatusmoneda
                    pPosicionaBancoVsf (strIdBanco)
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferida), cstrNumero))
                    dblTipoCambio = Val(Format(.TextMatrix(intcontador, cintColvsfTipoCambio), cstrNumero))
                    'AAT
                    'ldblMontoBanco = ldblMontoBanco + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                    ldblMontoBanco = dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                
'                    vgstrParametrosSP = CStr(llngNumCorte) _
'                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
'                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
'                                        & "|" & "TR" & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
'                                        & "|" & CStr(dblCantidad) _
'                                        & "|" & CStr(dblTipoCambio) _
'                                        & "|" & .TextMatrix(intcontador, cintColvsfReferencia) & "|" & CStr(rsTransferencia!IdCorte)
'                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
'
'                    'CARGO al a cuenta de la forma de pago FUENTE:
'                    vgstrParametrosSP = CStr(llngNumCorte) _
'                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
'                                        & "|" & "TR" _
'                                        & "|" & .TextMatrix(intcontador, cintColvsfCtaFuente) _
'                                        & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
'                                        & "|" & "1"
'                    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                    
                    'AAT
                    vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "TR" _
                                        & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                        & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                    frsEjecuta_SP vgstrParametrosSP, "SP_PVDETALLETRANSFBANCO"
                    lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, CLng(.TextMatrix(intcontador, cintColvsfCtaFuente)), dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio), "1", vl_refTR, vl_conTR)
                                
            '3.- Afectar el ABONO a la cuenta del banco:
'            vgstrParametrosSP = CStr(llngNumCorte) _
'                                & "|" & CStr(rsTransferencia!IdTransferencia) _
'                                & "|" & "TR" _
'                                & "|" & CStr(rsBancos!intNumeroCuenta) _
'                                & "|" & CStr(ldblMontoBanco) _
'                                & "|" & "0"
'            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"

                    'AAT se metieron estas lineas dentro del ciclo para que lea cada renglon de la grid
                    'Se cambia cboDeptoRecibe x el dato de banco de cada linea
                    '3.- Afectar el ABONO a la cuenta del banco:
                    lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, rsBancos!intNumeroCuenta, ldblMontoBanco, "0", vl_refTR, vl_conTR)
                      
                    '5.- Afectar el libro de bancos en su moneda
                    'Esta variable <ldblMontoBanco> tomo valor en <fblnGuardaForma> al recorrer las formas de pago:
                    ldblMontoBanco = ldblMontoBanco / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
                    
                    'vgstrParametrosSP = CStr(cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)) & "|" & fstrFechaSQL(CDate(mskFecha.Text), "00:00:00") _
                    '            & "|" & "CDE" & "|" & "0" & "|" & CStr(ldblMontoBanco) & "|" & CStr(rsTransferencia!IdTransferencia)
                    vgstrParametrosSP = strIdBanco & "|" & fstrFechaSQL(CDate(mskFecha.Text), "00:00:00") _
                                & "|" & "CDE" & "|" & "0" & "|" & CStr(ldblMontoBanco) & "|" & CStr(rsTransferencia!IdTransferencia)
                                
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
                    
                Next intcontador
            End With
            
            '6.- Liberar el corte:
'            pLiberaCorte llngNumCorte
            
            '7.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmBorrar, llngPersonaGraba, "TRANSFERENCIA BANCO", CStr(rsTransferencia!IdTransferencia)
        
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            txtNumero.SetFocus
        End If
    End If
        
'    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaTransBanco"))
End Sub

Private Sub pCancelaTransDepto()
On Error GoTo NotificaError

    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    Dim intcontador As Integer

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    If fblnCorteLibre(rsTransferencia!IdCorte) Then
        If fblnTransActiva Then
            '1.- Cancelar en maestro y poner en la tabla de cancelados
            vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & CStr(llngPersonaGraba)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaTransferencia"
        
            '2.- Insertar las formas de pago en el corte en forma positiva:
            With vsfTransferencia
                For intcontador = 1 To .Rows - 1
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferida), cstrNumero))
                    dblTipoCambio = Val(Format(.TextMatrix(intcontador, cintColvsfTipoCambio), cstrNumero))
                
                    vgstrParametrosSP = CStr(rsTransferencia!IdCorte) _
                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "TR" & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & .TextMatrix(intcontador, cintColvsfReferencia) & "|" & CStr(rsTransferencia!IdCorte)
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                Next intcontador
            End With
        
            '3.- Liberar el corte:
            pLiberaCorte rsTransferencia!IdCorte
        
            '4.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmBorrar, llngPersonaGraba, "TRANSFERENCIA DEPARTAMENTO", CStr(rsTransferencia!IdTransferencia)
        
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            txtNumero.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaTransDepto"))
End Sub

Private Sub pCancelaFondo()
On Error GoTo NotificaError

    Dim intcontador As Integer
    Dim dblCantidad As Double
    Dim dblTipoCambio As Double

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If fblnCorteLibre(llngNumCorte) Then
        If fblnTransActiva Then
            '1.- Cancelar en maestro y poner en la tabla de cancelados
            vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & CStr(llngPersonaGraba)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaTransferencia"
    
            '2.- Insertar las formas de pago en el corte actual en negativos:
            With vsfTransferencia
                For intcontador = 1 To .Rows - 1
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadFondo), cstrNumero)) * -1
                    dblTipoCambio = Val(Format(.TextMatrix(intcontador, cintColvsfTipoCambio), cstrNumero))
                
                    vgstrParametrosSP = CStr(llngNumCorte) _
                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "RI" & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & "0" & "|" & CStr(rsTransferencia!IdCorte)
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                Next intcontador
            End With
            
            '3.- Liberar el corte:
            pLiberaCorte llngNumCorte
    
            '4.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmBorrar, llngPersonaGraba, "FONDO FIJO", CStr(rsTransferencia!IdTransferencia)
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            txtNumero.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaFondo"))
End Sub

Private Function fblnDatosCancelar() As Boolean
On Error GoTo NotificaError
    
    fblnDatosCancelar = True
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     Que se introduzca la contraseña correcta    -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    fblnDatosCancelar = llngPersonaGraba <> 0
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     Que se tenga permisos de escritura, segun el proceso seleccionado   -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosCancelar Then
        If (cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo And Not lblnPermisoFondo) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And Not lblnPermisoTranDepto) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco And Not lblnPermisoTranBanco) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas And Not lblnPermisoCambioformas) Then
            fblnDatosCancelar = False
            '¡El usuario no tiene permiso para grabar datos!
            MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosCancelar"))
End Function

Private Sub cmdEnd_Click()
On Error GoTo NotificaError

    grdTransferencia.Row = grdTransferencia.Rows - 1
    pMuestra grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)
    pHabilita 1, 1, 1, 1, 1, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf(rsTransferencia!Estado = "R", 1, 0)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
On Error GoTo NotificaError

    SSTabTransfencia.Tab = 1
    mskFechaInicio.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
On Error GoTo NotificaError

    If grdTransferencia.Row < grdTransferencia.Rows - 1 Then grdTransferencia.Row = grdTransferencia.Row + 1
    pMuestra grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)
    pHabilita 1, 1, 1, 1, 1, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf(rsTransferencia!Estado = "R", 1, 0)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Private Sub cmdPrint_Click()
On Error GoTo NotificaError

    Dim rptReporte As CRAXDRT.Report
    Dim alstrParametros(0) As String
    
    vgstrParametrosSP = txtNumero.Text & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "0" & "|" & "-1" & "|" & "-1"
    Set rsTransferencia = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelTransferencia")
    
    pInstanciaReporte rptReporte, "rptTransferenciaCredito.rpt"
    rptReporte.DiscardSavedData
    
    alstrParametros(0) = "NombreHospital" & ";" & Trim(vgstrNombreHospitalCH) & ";TRUE"
    
    pCargaParameterFields alstrParametros, rptReporte
    
    pImprimeReporte rptReporte, rsTransferencia, "P", "Transferencia a departamento"

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub

Private Sub cmdPrintPoliza_Click()
On Error GoTo NotificaError
    
    If llngNumPoliza = 0 Then Exit Sub
    
    pImprimePoliza Str(llngNumPoliza), "P"
    
    'If fblnCanFocus(txtNumero) Then txtNumero.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrintPoliza_Click"))
End Sub


Private Function fblnValidaCuenta() As Boolean
    
    fblnValidaCuenta = True
    'AAT
    Dim dblCantidad     As Double
    Dim intcontador     As Long
    Dim strCuentaBanco  As String
    
    With vsfTransferencia
        For intcontador = 1 To .Rows - 1
            .Row = intcontador
            .Col = cintColvsfTransferir
            If .CellChecked = flexChecked And cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
                dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferida), cstrNumero))
                strCuentaBanco = .TextMatrix(intcontador, cintColvsfBanco)
                If strCuentaBanco = "" Then
                  'error cuenta debe estar seleccionada
                  fblnValidaCuenta = False
                  Exit Function
                End If
            End If
        Next intcontador
    End With
End Function


Private Sub cmdSave_Click()
On Error GoTo NotificaError

    If Not fblnValidaCuenta Then
        MsgBox "Debe seleccionar la cuenta destino en cada una de las transferencias a realizar", vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    If fblnDatosGrabar() Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
            If vlintTipoCorte = 2 And llngPersonaGraba <> vglngNumeroEmpleado Then
                'La persona que inició el proceso no corresponde a la persona que quiere grabar.
                MsgBox SIHOMsg(593), vbInformation + vbOKOnly, "Mensaje"
                Exit Sub
            Else
                pGrabaFondo
            End If
        End If
        
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And Not lblnConsulta Then
            If vlintTipoCorte = 2 And llngPersonaGraba <> vglngNumeroEmpleado Then
                'La persona que inició el proceso no corresponde a la persona que quiere grabar.
                MsgBox SIHOMsg(593), vbInformation + vbOKOnly, "Mensaje"
                Exit Sub
            Else
                pGrabaTransDepto
            End If
        End If
        
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And lblnConsulta Then
            pGrabaRecepcion
        End If
        'AAT
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
            If fblnValidaCuenta Then
                pGrabaTransBanco
            Else
                MsgBox "Debe seleccionar la cuenta destino en cada una de las transferencias a realizar", vbExclamation + vbOKOnly, "Mensaje"
                Exit Sub
            End If
        End If
        
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
            pGrabaTransFormas
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
End Sub

Private Sub pGrabaTransFormas()
On Error GoTo NotificaError

    Dim blnGuardaForma As Boolean
    Dim vllngConsecutivoDetalle As Long
    Dim blnerror As Boolean
    Dim vllngBanco As Long
    '---- (CR) Agregado - Caso 7442 ----'
    Dim lblnEntraCorte As Boolean        'Indica si el corte está cerrado y los movimientos deben generarse con datos del corte cerrado
    Dim ldtmFechaCorte As Date           'Especifica la fecha del corte con el que se está trabajando
    Dim lstrFechaTransferencia As String 'Fecha de la transferencia: si el corte está cerrado, trabaja con la fecha del corte, si no, con la fecha actual
    '-----------------------------------'
    
    blnerror = False
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    '*-     INTERCAMBIO ENTRE FORMAS DE PAGO     *-'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    
    pPosicionaForma 'Posicionarse en la forma de pago que va a recibir la transferencia
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    '*- Que se haya capturado folio o banco si se requiere en los cambios de forma de pago *-'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    If rsFormasPago!bitpreguntafolio = 1 Then
        If txtDatoPersonaRecibe.Text = "" Then
            blnerror = True
            'No ha ingresado datos
            MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
            If fblnCanFocus(txtDatoPersonaRecibe) Then txtDatoPersonaRecibe.SetFocus
        End If
    End If
    
    If rsFormasPago!chrTipo = "B" And Not blnerror Then
        If cboDatoCorteRecibe.ListIndex = -1 Then
            blnerror = True
            'No ha ingresado datos
            MsgBox SIHOMsg(2), vbOKOnly + vbExclamation, "Mensaje"
            If fblnCanFocus(cboDatoCorteRecibe) Then cboDatoCorteRecibe.SetFocus
        End If
    End If
    
    If blnerror Then
        Exit Sub
    Else
        EntornoSIHO.ConeccionSIHO.BeginTrans
    End If
     
    '-------- MODIFICADO PARA CASO 7442 (CR) --------'
    If lblnPermisoCambioFormasDepto Then
        llngNumCorte = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)
    Else
        llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    End If
        
    If fblnCorteLibre(llngNumCorte) Then
        '1.- Traer los datos del corte: Si está abierto, traer fecha actual. Si está cerrado traer fecha del corte
        ldtmFechaCorte = CDate(frsEjecuta_SP(Str(llngNumCorte), "Sp_GnSelDatosCorte")!dtmFechahora)
        'lblnEntraCorte = (Format(ldtmFechaCorte, "dd/mm/yyyy") <= CDate(mskFecha.Text)) And (llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P"))
        lblnEntraCorte = fblnCorteAbierto(llngNumCorte)
        If lblnEntraCorte Then
            lstrFechaTransferencia = fstrFechaSQL(fdtmServerFecha, fdtmServerHora)
        Else
            lstrFechaTransferencia = fstrFechaSQL(Format(ldtmFechaCorte, "dd/mm/yyyy"), Format(fdtmServerHora, "hh:mm:ss"))
        End If
        '--------------------------------------------------------------------------------------------------------------------------------'
    
        ReDim apoliza(0)
        apoliza(0).vllngNumeroCuenta = 0
        llngNumPoliza = 0
    
        '2.- Guardar transferencia maestro:
        llngIdTransferencia = flngTransferencia(llngNumCorte)
    
        '3.- Guardar las formas de pago a transferir
        blnGuardaForma = fblnGuardaFormas(llngNumCorte, lstrFechaTransferencia, lblnEntraCorte)
        If Not blnGuardaForma Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'La información ha cambiado, consulte de nuevo.
            MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
        Else
            '4.- Insertar detalle del corte de la forma maestro
            vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & lstrFechaTransferencia _
                                & "|" & CStr(llngIdTransferencia) _
                                & "|" & "TR" _
                                & "|" & CStr(cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)) _
                                & "|" & IIf(rsFormasPago!BITPESOS = 0, CStr(ldblMontoTotal / ldblTipoCambioVenta), CStr(ldblMontoTotal)) _
                                & "|" & IIf(rsFormasPago!BITPESOS = 1, 0, CStr(ldblTipoCambioVenta)) _
                                & "|" & IIf(CStr(txtDatoPersonaRecibe.Text) = "", 0, CStr(txtDatoPersonaRecibe.Text)) _
                                & "|" & CStr(llngNumCorte)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
            vllngConsecutivoDetalle = flngObtieneIdentity("Sec_PvDetalleCorte", 1)
            
            '5.- Guardar intercambio entre formas:
            If cboDatoCorteRecibe.ListIndex = -1 Then
                vllngBanco = 0
            Else
                vllngBanco = cboDatoCorteRecibe.ItemData(cboDatoCorteRecibe.ListIndex)
            End If
            
            vgstrParametrosSP = CStr(llngIdTransferencia) _
                                & "|" & vllngConsecutivoDetalle _
                                & "|" & vllngBanco _
                                & "|" & IIf(rsFormasPago!INTCUENTACONTABLE = 0, rsBancos!intNumeroCuenta, rsFormasPago!INTCUENTACONTABLE)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaFormas"
                        
            '6.- Afectar la poliza del corte, CARGO a la forma de pago que recibe:
            vgstrParametrosSP = CStr(llngNumCorte) _
                                & "|" & CStr(llngIdTransferencia) _
                                & "|" & "TR" _
                                & "|" & IIf(CStr(rsFormasPago!INTCUENTACONTABLE) = 0, rsBancos!intNumeroCuenta, CStr(rsFormasPago!INTCUENTACONTABLE)) _
                                & "|" & CStr(ldblMontoTotal) _
                                & "|" & "1"
            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
        
            '7.- Afectar el libro de bancos en su moneda, si es que aplica
            'Esta variable <ldblMontoBanco> toma valor en <fblnGuardaForma> al recorrer las formas de pago:
            ldblMontoBanco = ldblMontoBanco / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
            'If rsFormaspago!chrTipo = "B" Then
            '    ldblMontoTotal = ldblMontoTotal / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
            '    vgstrParametrosSP = CStr(cboDatoCorteRecibe.ItemData(cboDatoCorteRecibe.ListIndex)) & "|" & lstrFechaTransferencia _
            '                        & "|" & "DEP" & "|" & CStr(ldblMontoTotal) & "|" & "0" & "|" & CStr(llngIdTransferencia)
            '    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
            'End If
            vllngBanco = flngEsCuentaBanco(rsFormasPago!INTCUENTACONTABLE)
            If vllngBanco <> 0 Then
                ldblMontoTotal = ldblMontoTotal / IIf(rsFormasPago!BITPESOS = 0, ldblTipoCambioVenta, 1)
                If Not lblnEntraCorte Then
                    'Si el corte está cerrado, realizar el movimiento en el kardex del banco
                    vgstrParametrosSP = vllngBanco & "|" & lstrFechaTransferencia & "|" & "DEP" & "|" & CStr(ldblMontoTotal) & "|" & "0" & "|" & CStr(llngIdTransferencia)
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
                End If
                '-- Guardar información de la forma de pago en tabla intermedia. Se marca como pendiente (1) si el corte está abierto o registrada (2) si el corte está cerrado --'
                vgstrParametrosSP = llngNumCorte & "|" & lstrFechaTransferencia & "|" & rsFormasPago!intFormaPago & "|" & vllngBanco & "|" & _
                                    ldblMontoTotal & "|" & rsFormasPago!BITPESOS & "|" & ldblTipoCambioVenta & "|" & "DEP" & "|" & "TR" & "|" & llngIdTransferencia & "|" & _
                                    llngPersonaGraba & "|" & lintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & IIf(lblnEntraCorte, "1", "2") & "|" & cgstrModulo
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
            End If

            '8.- Liberar el corte:
            pLiberaCorte llngNumCorte
            
            '9.- Si el corte está cerrado revisar que existan movimientos de pólizas
            If Not lblnEntraCorte Then
                If apoliza(0).vllngNumeroCuenta <> 0 Then
                    'Si existen pólizas, revisar si el periodo contable NO está cerrado
                    If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(ldtmFechaCorte), Month(ldtmFechaCorte)) Then
                        MsgBox SIHOMsg(209), vbExclamation + vbOKOnly, "Mensaje"
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        Exit Sub
                    Else
                        'Guardar pólizas maestro y detalle de las diferentes cuentas contables
                        llngNumPoliza = flngInsertarPolizaMaestro(CDate(Format(ldtmFechaCorte, "dd/mmm/yyyy")), "D", "TRANSFERENCIA " & CStr(llngIdTransferencia) & " POR CAMBIO DE FORMA DE PAGO", lintNumeroDepartamento, llngPersonaGraba)
                        pGuardarDetallePoliza llngNumPoliza
                        pGuardarPolizaTransferencia llngIdTransferencia, llngNumPoliza 'Guardar relación de transferencia y número de póliza
                    End If
                End If
            End If
            
            '10.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "CAMBIO DE FORMAS DE PAGO", CStr(llngIdTransferencia)
            
            '11.- Terminar la transacción
            EntornoSIHO.ConeccionSIHO.CommitTrans
                
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            
            If llngNumPoliza <> 0 Then
                MsgBox "Se ha generado una póliza para este movimiento, presione el botón de póliza para consultarla", vbInformation + vbOKOnly, "Mensaje"
                pMuestra CStr(llngIdTransferencia)
            Else
                pLimpia
                pConfiguraVsf "II"
            End If
            pHabilita 0, 0, 1, 0, 0, 0, 0, 0
            If Not cmdPrintPoliza.Enabled Then txtNumero.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaTransFormas"))
End Sub

'AAT - se modifica incluyendo parametro de entrada
Private Sub pGrabaTransBanco()
On Error GoTo NotificaError

    Dim blnGuardaForma As Boolean
    Dim lngNumDetalle As Long
    Dim intcontador As Long
    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    Dim strIdBanco As String
    'AAT
    Dim vlstrCargarRs As String
    Dim rsTranferencia As New ADODB.Recordset
    Dim vl_refTR As String
    Dim vl_conTR As String
    Dim vl_RefConcepto As Integer
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     TRANSFERENCIA A BANCO     -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*      -*-*-*-*-*'
    EntornoSIHO.ConeccionSIHO.BeginTrans
        
'    llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
'    If fblnCorteLibre(llngNumCorte) Then
        
'        '1.- Guardar transferencia maestro:
'        llngIdTransferencia = flngTransferencia(llngNumCorte)
        
    '1.- Guardar transferencia maestro:
    
    'Se crea Poliza asociada a la trasferencia
    vllngNumPoliza = 0
    lstrConceptoPoliza = ""
    If fblnObtenerConfigPolizanva(0, "1", frmTransferencia.Name) = True Then
        For vg_intIndiceConceptoPoliza = 0 To UBound(ConceptosPoliza, 1)
            Select Case ConceptosPoliza(vg_intIndiceConceptoPoliza).tipo
                Case "1"
                    lstrConceptoPoliza = lstrConceptoPoliza & ConceptosPoliza(vg_intIndiceConceptoPoliza).Texto
                Case "3"
                    lstrConceptoPoliza = lstrConceptoPoliza & txtNumero
                Case "25"
                    lstrConceptoPoliza = lstrConceptoPoliza & UCase(cboTipo.Text)
            End Select
        Next vg_intIndiceConceptoPoliza
    Else
        Me.MousePointer = 0
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Sub
    End If
    vllngNumPoliza = flngInsertarPoliza(CDate(mskFecha.Text), "D", lstrConceptoPoliza, llngPersonaGraba)
    llngIdTransferencia = flngTransferencia(0)
    pEjecutaSentencia "INSERT INTO PVTRANSFERENCIABANCOPOLIZA VALUES(" & llngIdTransferencia & ", " & vllngNumPoliza & ")"
                
    'AAT - Insercion de trasferencia para un unico banco - Codigo Original
    'If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        
        '2.- Guardar transferencia a banco:
    '    vgstrParametrosSP = CStr(llngIdTransferencia) & "|" & cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)
    '    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaBanco"
        
    'Else
    
    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        'AAT - Insercion de transferencia para varios bancos segun el FlexGrid
        '2.- Guardar cada transferencia a banco del flexgrid por cada linea:
        With vsfTransferencia
            intcontador = 1
            Do While intcontador <= .Rows - 1
                'Si hay importe a transferir
                If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                    'obtener el valor del id del banco
                    strIdBanco = Val(.TextMatrix(intcontador, cintColvsfIdBanco))
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                    
                    'Validar si ya hay registro de transferencia vs banco
                    vlstrCargarRs = "SELECT COUNT(INTIDTRANSFERENCIA) AS CONTEO FROM PVTRANSFERENCIABANCO WHERE INTIDTRANSFERENCIA =  " & CStr(llngIdTransferencia) & " AND INTCVEBANCO = " & strIdBanco
                    Set rsTranferencia = frsRegresaRs(vlstrCargarRs, adLockReadOnly, adOpenForwardOnly)
                    
                    If rsTranferencia!conteo = 0 Then
                        If strIdBanco <> 0 Then
                            vgstrParametrosSP = CStr(llngIdTransferencia) & "|" & strIdBanco
                            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaBanco"
                        End If
                    End If
                End If
                intcontador = intcontador + 1
            Loop
        End With
    End If
        
    '        '3.- Guardar las formas de pago a transferir
    '        blnGuardaForma = fblnGuardaForma(llngNumCorte)
            
    'AAT hace commit para guadar encabezado poliza
'    EntornoSIHO.ConeccionSIHO.CommitTrans
'    EntornoSIHO.ConeccionSIHO.BeginTrans
    '3.- Guardar las formas de pago a transferir
    With vsfTransferencia
        ldblMontoBanco = 0
        intcontador = 1
        Do While intcontador <= .Rows - 1
            If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                dblTipoCambio = Val(.TextMatrix(intcontador, cintColvsfTipoCambio))
            
                'Detalle de las formas de pago de la transferencia:
                'AAT
                vgstrParametrosSP = CStr(llngIdTransferencia) _
                                    & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                    & "|" & CStr(dblCantidad) _
                                    & "|" & CStr(dblTipoCambio) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaForma"
                
                'AAT - si es para diferentes bancos no acumula el monto - se registra individualmente
                ldblMontoBanco = dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                'Linea Original
                'ldblMontoBanco = ldblMontoBanco + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                        
                For vl_RefConcepto = 1 To 2
                    lstrConceptoDetallePoliza = " "
                    If fblnObtenerConfigPolizanva(vl_RefConcepto, "1", frmTransferencia.Name) = True Then
                        For vg_intIndiceConceptoPoliza = 0 To UBound(ConceptosPoliza, 1)
                            Select Case ConceptosPoliza(vg_intIndiceConceptoPoliza).tipo
                                Case "1"
                                    lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & ConceptosPoliza(vg_intIndiceConceptoPoliza).Texto
                                Case "26"
                                    If vsfTransferencia.TextMatrix(intcontador, cintColvsfCantidadTransferir) <> "" Then
                                        lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & Trim(vsfTransferencia.TextMatrix(intcontador, cintColvsfFormaPago))
                                    End If
                                Case "27"
                                    If vsfTransferencia.TextMatrix(intcontador, cintColvsfCantidadTransferir) <> "" Then
                                        lstrConceptoDetallePoliza = lstrConceptoDetallePoliza & Trim(vsfTransferencia.TextMatrix(intcontador, cintColvsfBanco))
                                    End If
                            End Select
                        Next vg_intIndiceConceptoPoliza
                    End If
                    If vl_RefConcepto = 2 Then vl_refTR = Trim(lstrConceptoDetallePoliza)
                    If vl_RefConcepto = 1 Then vl_conTR = Trim(lstrConceptoDetallePoliza)
                Next vl_RefConcepto
                
                If dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio) <> 0 Then
                    If Val(.TextMatrix(intcontador, cintColvsfCtaFuente)) <> 0 Then
                        lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(intcontador, cintColvsfCtaFuente)), dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio), "0", vl_refTR, vl_conTR)
                    Else
                        lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(intcontador, cintColvsfIdBanco)), dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio), "0", vl_refTR, vl_conTR)
                    End If
                End If
                
                'Corte simulado negativo:
                dblCantidad = dblCantidad * -1
                'AAT
                vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                    & "|" & CStr(llngIdTransferencia) _
                                    & "|" & "TR" _
                                    & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                    & "|" & CStr(dblCantidad) _
                                    & "|" & CStr(dblTipoCambio) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                frsEjecuta_SP vgstrParametrosSP, "SP_PVDETALLETRANSFBANCO"
                
                intcontador = intcontador + 1
            Else
                intcontador = intcontador + 1
            End If
        Loop
    End With
        
    'AAT - se registra detalle poliza segun tipo de trasferencia
    'If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        
    '   pPosicionaBanco
        
'        '4.- Afectar la poliza del corte, CARGO a la cuenta del banco:
'        vgstrParametrosSP = CStr(llngNumCorte) _
'                            & "|" & CStr(llngIdTransferencia) _
'                            & "|" & "TR" _
'                            & "|" & CStr(rsBancos!intNumeroCuenta) _
'                            & "|" & CStr(ldblMontoBanco) _
'                            & "|" & "1"
'        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
        
    '    If ldblMontoBanco <> 0 Then
            '4.- Afectar la poliza del corte, CARGO a la cuenta del banco:
    '        lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, rsBancos!intNumeroCuenta, ldblMontoBanco, "1")
    '    End If
    'Else
    
    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        'Para bancos diferentes se lee liena a linea del flexgrid
        With vsfTransferencia
            ldblMontoBanco = 0
            intcontador = 1

            Do While intcontador <= .Rows - 1
                'si el importe a trasferir en diferente de cero
                If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                    'se obtiene id del banco en la fila
                    strIdBanco = Val(.TextMatrix(intcontador, cintColvsfIdBanco))
                    'se busca info del banco para obtener la cuenta
                    pPosicionaBancoVsf (strIdBanco)

                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                    dblTipoCambio = Val(.TextMatrix(intcontador, cintColvsfTipoCambio))
                    ldblMontoBanco = dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)

                    If ldblMontoBanco <> 0 Then
                        '4.- Afectar la poliza del corte, CARGO a la cuenta del banco:
                        lngNumDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, rsBancos!intNumeroCuenta, ldblMontoBanco, "1", vl_refTR, vl_conTR)
                    End If
                End If
                intcontador = intcontador + 1
            Loop
        End With
    End If
    vl_refTR = ""
    vl_conTR = ""
    
    'AAT - se registra detalle poliza segun tipo de trasferencia
    'If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then 'Transferencia original todo a un solo banco
        '5.- Afectar el libro de bancos en su moneda
        'Esta variable <ldblMontoBanco> tomo valor en <fblnGuardaForma> al recorrer las formas de pago:
    '    ldblMontoBanco = ldblMontoBanco / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
        
    '    vgstrParametrosSP = CStr(cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)) & "|" & fstrFechaSQL(CDate(mskFecha.Text), "00:00:00") _
    '                        & "|" & "DEP" & "|" & CStr(ldblMontoBanco) & "|" & "0" & "|" & CStr(llngIdTransferencia)
    '    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
    'Else
    
    'AAT - se registra detalle poliza segun tipo de trasferencia
    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        'Para bancos diferentes se lee liena a linea del flexgrid
        With vsfTransferencia
            ldblMontoBanco = 0
            intcontador = 1
            Do While intcontador <= .Rows - 1
                'si el importe a trasferir en diferente de cero
                If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                    'se obtiene id del banco en la fila
                    strIdBanco = Val(.TextMatrix(intcontador, cintColvsfIdBanco))
                    'se busca info del banco para obtener la cuenta y bitstatusmoneda
                    pPosicionaBancoVsf (strIdBanco)
                    
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                    dblTipoCambio = Val(.TextMatrix(intcontador, cintColvsfTipoCambio))
                    ldblMontoBanco = dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                    
                    ldblMontoBanco = ldblMontoBanco / IIf(rsBancos!bitestatusmoneda = 0, ldblTipoCambioVenta, 1)
        
                    vgstrParametrosSP = strIdBanco & "|" & fstrFechaSQL(CDate(mskFecha.Text), "00:00:00") _
                            & "|" & "DEP" & "|" & CStr(ldblMontoBanco) & "|" & "0" & "|" & CStr(llngIdTransferencia)
                    frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
                End If
                intcontador = intcontador + 1
            Loop
        End With
    
    End If
    
'        '6.- Liberar el corte:
'        pLiberaCorte llngNumCorte
    
    '7.- Afectar registro de transacciones
    pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "TRANSFERENCIA BANCO", CStr(llngIdTransferencia)
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
        
    'La operación se realizó satisfactoriamente.
    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
    
    txtNumero.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaTransBanco"))
End Sub

Private Sub pPosicionaBanco()
On Error GoTo NotificaError

    Dim blnTermina As Boolean

    rsBancos.MoveFirst
    Do While Not rsBancos.EOF And Not blnTermina
        blnTermina = rsBancos!tnynumerobanco = cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)
        If Not blnTermina Then
            rsBancos.MoveNext
        End If
    Loop

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPosicionaBanco"))
End Sub

Private Sub pPosicionaBancoVsf(strNumBanco As String)
On Error GoTo NotificaError

    Dim blnTermina As Boolean

    rsBancos.MoveFirst
    Do While Not rsBancos.EOF And Not blnTermina
        blnTermina = rsBancos!tnynumerobanco = strNumBanco
        If Not blnTermina Then
            rsBancos.MoveNext
        End If
    Loop

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPosicionaBanco"))
End Sub



Private Sub pPosicionaForma()
On Error GoTo NotificaError

    Dim blnTermina As Boolean

    rsFormasPago.MoveFirst
    Do While Not rsFormasPago.EOF And Not blnTermina
        blnTermina = rsFormasPago!intFormaPago = cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex)
        If Not blnTermina Then
            rsFormasPago.MoveNext
        End If
    Loop
    
    blnTermina = False
    If rsFormasPago!chrTipo = "B" Then
        rsBancos.MoveFirst
        Do While Not rsBancos.EOF Or Not blnTermina
            blnTermina = rsBancos!tnynumerobanco = cboDatoCorteRecibe.ItemData(cboDatoCorteRecibe.ListIndex)
            If Not blnTermina Then
                rsBancos.MoveNext
            End If
        Loop
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPosicionaBanco"))
End Sub

Private Function fblnTransActiva() As Boolean
On Error GoTo NotificaError

    fblnTransActiva = False
    
    'Revisar que no haya sido cancelada:
    vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "0" & "|" & "-1" & "|" & "-1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelTransferencia")
    If rs.RecordCount <> 0 Then
        fblnTransActiva = Trim(rs!Estado) = "A"
    Else
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'La información ha cambiado, consulte de nuevo.
        MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
        txtNumero.SetFocus
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnTransActiva"))
End Function

Private Sub pGrabaRecepcion()
On Error GoTo NotificaError

    Dim intcontador As Long
    'Dim intcontador As Integer
    Dim dblCantidad As Double
    Dim dblTipoCambio As Double

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     RECEPCION DEL DINERO QUE SE TRANSFIERE      -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If fblnCorteLibre(llngNumCorte) Then
        If fblnTransActiva Then
            '1.- Actualizar el estado de la transferencia recibida
            frsEjecuta_SP CStr(rsTransferencia!IdTransferencia), "Sp_PvUpdTransferenciaRecibida"
            
            '2.- Insertar en la tabla de transferencias recibidas
            vgstrParametrosSP = CStr(rsTransferencia!IdTransferencia) & "|" & CStr(llngNumCorte) & "|" & CStr(llngPersonaGraba)
            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaRecibida"

            '3.- Afectar el corte con cantidades positivas y la poliza del corte
            With vsfTransferencia
                For intcontador = 1 To .Rows - 1
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferida), cstrNumero))
                    dblTipoCambio = Val(Format(.TextMatrix(intcontador, cintColvsfTipoCambio), cstrNumero))
                
                    vgstrParametrosSP = CStr(llngNumCorte) _
                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "TR" & "|" & .TextMatrix(intcontador, cintColvsfCveFormaDestino) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & .TextMatrix(intcontador, cintColvsfReferencia) & "|" & CStr(rsTransferencia!IdCorte)
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                    
                    'CARGO al a cuenta de la forma de pago DESTINO:
                    vgstrParametrosSP = CStr(llngNumCorte) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "TR" _
                                        & "|" & .TextMatrix(intcontador, cintColvsfCtaDestino) _
                                        & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
                                        & "|" & "1"
                    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                    
                    'ABONO al a cuenta de la forma de pago FUENTE:
                    vgstrParametrosSP = CStr(llngNumCorte) _
                                        & "|" & CStr(rsTransferencia!IdTransferencia) _
                                        & "|" & "TR" _
                                        & "|" & .TextMatrix(intcontador, cintColvsfCtaFuente) _
                                        & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
                                        & "|" & "0"
                    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                Next intcontador
            End With

            '4.- Liberar el corte:
            pLiberaCorte llngNumCorte
            
            '5.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "RECEPCION DINERO", CStr(rsTransferencia!IdTransferencia)
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            pMuestra rsTransferencia!IdTransferencia 'Para que refresque los datos de "Persona recibe" y "Corte recibe"
            
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            
            pHabilita 0, 0, 1, 0, 0, 0, 0, 1
            
            cmdPrint.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaRecepcion"))
End Sub

Private Sub pGrabaTransDepto()
On Error GoTo NotificaError

    Dim blnGuardaForma As Boolean

    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     TRANSFERENCIA A DEPARTAMENTO    -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    If fblnCorteLibre(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)) Then
        '1.- Guardar transferencia maestro:
        llngIdTransferencia = flngTransferencia(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex))
        
        '2.- Guardar transferencia a departamento:
        vgstrParametrosSP = CStr(llngIdTransferencia) & "|" & CStr(cboDeptoRecibe.ItemData(cboDeptoRecibe.ListIndex))
        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaDepto"
        
        '3.- Guardas las formas de pago a transferir:
        blnGuardaForma = fblnGuardaForma(cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex))
        
        If Not blnGuardaForma Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'La información ha cambiado, consulte de nuevo.
            MsgBox SIHOMsg(381), vbExclamation + vbOKOnly, "Mensaje"
        Else
            '4.- Liberar el corte:
            pLiberaCorte cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex)
        
            '5.- Afectar registro de transacciones
            pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "TRANSFERENCIA DEPARTAMENTO", CStr(llngIdTransferencia)
    
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            
            txtNumero.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaTransDepto"))
End Sub

Private Function fblnCorteLibre(lngNumCorte As Long) As Boolean
On Error GoTo NotificaError

    llngEstadoCorte = 1
    frsEjecuta_SP CStr(lngNumCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, llngEstadoCorte
    
    fblnCorteLibre = llngEstadoCorte = 2
    
    If llngEstadoCorte <> 2 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'En este momento se está afectando el corte, espere un momento e intente de nuevo.
        MsgBox SIHOMsg(779), vbExclamation + vbOKOnly, "Mensaje"
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnCorteLibre"))
End Function

Private Sub pGrabaFondo()
On Error GoTo NotificaError

    Dim blnGuardaForma As Boolean
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*'
    '*-     FONDO FIJO      -*'
    '*-*-*-*-*-*-*-*-*-*-*-*-*'
    llngNumCorte = flngNumeroCorte(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If fblnCorteLibre(llngNumCorte) Then
        '1.- Guardar transferencia maestro:
        llngIdTransferencia = flngTransferencia(llngNumCorte)
        
        '2.- Guardar las formas de pago del fondo fijo y afectar corte:
        blnGuardaForma = fblnGuardaForma(llngNumCorte)
    
        '3.- Liberar el corte:
        pLiberaCorte llngNumCorte
        
        '4.- Afectar registro de transacciones
        pGuardarLogTransaccion Me.Name, EnmGrabar, llngPersonaGraba, "FONDO FIJO", CStr(llngIdTransferencia)
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'La operación se realizó satisfactoriamente.
        MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
        
        txtNumero.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaFondo"))
End Sub

Private Function fblnGuardaFormas(lngNumCorte As Long, dtmFechaPoliza As String, Optional lblnEntraCorte As Boolean = True) As Boolean
On Error GoTo NotificaError
    
    Dim intcontador As Long
    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    Dim lngBanco As Long
    
    fblnGuardaFormas = True
        
    With vsfTransferencia
        ldblMontoBanco = 0
        ldblMontoTotal = 0
        intcontador = 1
        Do While intcontador <= .Rows - 1 And fblnGuardaFormas
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
            '*-     CAMBIO DE FORMAS DE PAGO        -*'
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
            If (cboTipo.ItemData(cboTipo.ListIndex)) = cintIdTransFormas And Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                vgstrParametrosSP = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex) & "|" & CLng(lintNumeroDepartamento) & "|" & llngForma
                Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormasPagoCorte")
                If rs.RecordCount <> 0 And llngNumCorte = cboCorteTransfiere.ItemData(cboCorteTransfiere.ListIndex) Then
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                    dblTipoCambio = Val(.TextMatrix(intcontador, cintColvsfTipoCambio))
                
                    'Detalle de las formas de pago de la transferencia:
                    'AAT
                    vgstrParametrosSP = CStr(llngIdTransferencia) _
                                        & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & IIf(.TextMatrix(intcontador, cintColvsfReferencia) = " ", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                        & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaForma"
                    
                    ldblMontoTotal = ldblMontoTotal + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                    
                    If Trim(CStr(.TextMatrix(intcontador, cintColvsfCveFormaDestino))) = "B" Then
                        ldblMontoBanco = ldblMontoBanco + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                    End If
                    
                    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
                        'ABONO a la cuenta de la forma de pago que transfiere:
                        vgstrParametrosSP = CStr(llngNumCorte) _
                                            & "|" & CStr(llngIdTransferencia) _
                                            & "|" & "TR" _
                                            & "|" & .TextMatrix(intcontador, cintColvsfCtaFuente) _
                                            & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
                                            & "|" & "0"
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                    End If
                    
                    'Corte negativo:
                    dblCantidad = dblCantidad * -1
                    
                    vgstrParametrosSP = CStr(lngNumCorte) _
                                        & "|" & IIf(lblnEntraCorte, fstrFechaSQL(fdtmServerFecha, fdtmServerHora), dtmFechaPoliza) _
                                        & "|" & CStr(llngIdTransferencia) _
                                        & "|" & "TR" _
                                        & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                        & "|" & CStr(dblCantidad) _
                                        & "|" & CStr(dblTipoCambio) _
                                        & "|" & .TextMatrix(intcontador, cintColvsfReferencia) _
                                        & "|" & CStr(lngNumCorte)
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                    
                    '----------------------------------------------------------------------------------------------------------------'
                    '- (CR) Agregado Caso 7442: Movimientos para la generación de la póliza si las cuentas contables son diferentes -'
                    dblCantidad = dblCantidad * -1 'Para devolver la cantidad a positiva
                    If Val(.TextMatrix(intcontador, cintColvsfCtaFuente)) <> rsFormasPago!INTCUENTACONTABLE Then
                        pIncluyeMovimiento rsFormasPago!INTCUENTACONTABLE, dblCantidad, 1 'Cargo a la cuenta que recibe
                        pIncluyeMovimiento CLng(.TextMatrix(intcontador, cintColvsfCtaFuente)), dblCantidad, 0 'Abono a la cuenta que transfiere
                    End If
                    
                    lngBanco = flngEsCuentaBanco(CLng(.TextMatrix(intcontador, cintColvsfCtaFuente)))
                    If Not lblnEntraCorte And lngBanco <> 0 Then
                        'Si el corte está cerrado, realizar el movimiento en el kardex del banco de la cuenta que transfiere
                        vgstrParametrosSP = lngBanco & "|" & dtmFechaPoliza & "|" & "CDE" & "|" & "0" & "|" & CStr(dblCantidad) & "|" & CStr(llngIdTransferencia)
                        frsEjecuta_SP vgstrParametrosSP, "SP_CPINSKARDEXBANCO"
                    End If
                    '----------------------------------------------------------------------------------------------------------------'
                Else
                    fblnGuardaFormas = False
                End If
            End If
        
            intcontador = intcontador + 1
        Loop
    End With

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnGuardaFormas"))
End Function

Private Function fblnGuardaForma(lngNumCorte As Long) As Boolean
On Error GoTo NotificaError
    'Esta función regresa falso cuando la transferencia a departamento no es válida
    Dim intcontador As Long
    Dim dblCantidad As Double
    Dim dblTipoCambio As Double
    
    fblnGuardaForma = True
    
    With vsfTransferencia
        ldblMontoBanco = 0
        intcontador = 1
        Do While intcontador <= .Rows - 1 And fblnGuardaForma
            '*-*-*-*-*-*-*-*-*'
            '*- FONDO FIJO  -*'
            '*-*-*-*-*-*-*-*-*'
            If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo And Val(.TextMatrix(intcontador, cintColvsfIdForma)) <> 0 Then
            'Si es fondo fijo y se seleccionó forma de pago y se puso:
                
                'Detalle de las formas de pago de la transferencia:
                'AAT
                vgstrParametrosSP = CStr(llngIdTransferencia) _
                                    & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                    & "|" & Format(.TextMatrix(intcontador, cintColvsfCantidadFondo), cstrNumero) _
                                    & "|" & .TextMatrix(intcontador, cintColvsfTipoCambio) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                    & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaForma"
                
                'Corte:
                vgstrParametrosSP = CStr(lngNumCorte) _
                                    & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                    & "|" & CStr(llngIdTransferencia) _
                                    & "|" & "RI" _
                                    & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                    & "|" & Format(.TextMatrix(intcontador, cintColvsfCantidadFondo), cstrNumero) _
                                    & "|" & .TextMatrix(intcontador, cintColvsfTipoCambio) _
                                    & "|" & "0" _
                                    & "|" & CStr(lngNumCorte)
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
            End If
            
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
            '*-  TRANSFERENCIA A DEPARTAMENTO O BANCO   -*'
            '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
            If (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco) And Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) <> 0 Then
                If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Then
                    'Dinero del corte:
                    vgstrParametrosSP = CStr(lngNumCorte) & "|" & .TextMatrix(intcontador, cintColvsfIdForma) & "|" & "-1" & "|" & "'-1'" & "|" & 2
                Else
                    'Dinero del depto:
                    vgstrParametrosSP = "-1" & "|" & .TextMatrix(intcontador, cintColvsfIdForma) & "|" & CStr(lintNumeroDepartamento) & "|'" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "-1", .TextMatrix(intcontador, cintColvsfReferencia)) & "'|" & 1
                End If
                
                If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
                    Set rs = frsEjecuta_SP(vgstrParametrosSP & "|" & 1, "Sp_PvSelDineroCorte")
                Else
                    Set rs = frsEjecuta_SP(vgstrParametrosSP & "|" & 0, "Sp_PvSelDineroCorte")
                End If
                
                If rs.RecordCount <> 0 Then
                    dblCantidad = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero))
                    If Round(rs!cantidad, 2) >= dblCantidad Then
                        dblTipoCambio = Val(.TextMatrix(intcontador, cintColvsfTipoCambio))
                    
                        'Detalle de las formas de pago de la transferencia:
                        'AAT
                        vgstrParametrosSP = CStr(llngIdTransferencia) _
                                            & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                            & "|" & CStr(dblCantidad) _
                                            & "|" & CStr(dblTipoCambio) _
                                            & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfReferencia)) = "", "0", .TextMatrix(intcontador, cintColvsfReferencia)) _
                                            & "|" & IIf(Trim(.TextMatrix(intcontador, cintColvsfIdBanco)) = "", "0", .TextMatrix(intcontador, cintColvsfIdBanco))
                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferenciaForma"
                        
                        ldblMontoBanco = ldblMontoBanco + dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)
                        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
                            'ABONO a la cuenta de la forma de pago:
                            vgstrParametrosSP = CStr(llngNumCorte) _
                                                & "|" & CStr(llngIdTransferencia) _
                                                & "|" & "TR" _
                                                & "|" & .TextMatrix(intcontador, cintColvsfCtaFuente) _
                                                & "|" & CStr(dblCantidad * IIf(dblTipoCambio = 0, 1, dblTipoCambio)) _
                                                & "|" & "0"
                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                        End If
                        
                        'Corte negativo:
                        dblCantidad = dblCantidad * -1
                        
                        vgstrParametrosSP = CStr(lngNumCorte) _
                                            & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                            & "|" & CStr(llngIdTransferencia) _
                                            & "|" & "TR" _
                                            & "|" & .TextMatrix(intcontador, cintColvsfIdForma) _
                                            & "|" & CStr(dblCantidad) _
                                            & "|" & CStr(dblTipoCambio) _
                                            & "|" & .TextMatrix(intcontador, cintColvsfReferencia) _
                                            & "|" & CStr(lngNumCorte)
                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsDetalleCorte"
                    Else
                        fblnGuardaForma = False
                    End If
                Else
                    fblnGuardaForma = False
                End If
            End If
            intcontador = intcontador + 1
        Loop
    End With

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnGuardaForma"))
End Function

Private Function flngTransferencia(lngNumCorte As Long) As Long
On Error GoTo NotificaError

    vgstrParametrosSP = fstrFechaSQL(mskFecha.Text) & "|" & CStr(llngPersonaGraba) & "|" & CStr(lintNumeroDepartamento) & "|" & CStr(cboTipo.ItemData(cboTipo.ListIndex)) & "|" & CStr(lngNumCorte)
    flngTransferencia = 1
    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsTransferencia", True, flngTransferencia

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngTransferencia"))
End Function

Private Sub cmdTop_Click()
On Error GoTo NotificaError
    
    grdTransferencia.Row = 1
    pMuestra grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)
    pHabilita 1, 1, 1, 1, 1, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf(rsTransferencia!Estado = "R", 1, 0)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError

    Dim lngMensaje As Long

    lngMensaje = flngCorteValido(lintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If lngMensaje <> 0 Then
        lblnCorteValido = False
        'Cierre el corte actual.
        MsgBox SIHOMsg(Str(lngMensaje)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    Else
        lblnCorteValido = True
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 27 Then
        If SSTabTransfencia.Tab <> 0 Then
            SSTabTransfencia.Tab = 0
            txtNumero.SetFocus
        Else
            If SSTabTransfencia.Tab = 0 Then
                If Me.ActiveControl.Name <> "txtNumero" Then
                    If (lblnConsulta Or cmdSave.Enabled) And lblnCorteValido Then
                        '¿Desea abandonar la operación?
                        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                            KeyAscii = 0
                            
                            cmdSave.Enabled = False
                            txtNumero.SetFocus
                        Else
                            KeyAscii = 0
                        End If
                    End If
                Else
                    Unload Me
                End If
            End If
        End If
    Else
        If KeyAscii = 13 Then
            If Me.ActiveControl.Name = "txtNumero" Then
                If Val(txtNumero.Text) = 0 Then
                    txtNumero.Text = CStr(flngSigTransferencia)
                    SendKeys vbTab
                Else
                    pMuestra txtNumero
                    If lblnConsulta Then
                        'Revisar si es una transferencia que registró el departamento o una que recibió:
                        If rsTransferencia!IdDepto = lintNumeroDepartamento Or (rsTransferencia!IdDeptoRecibe And Trim(rsTransferencia!Estado) <> "C") Then
                            pHabilita 0, 0, 0, 0, 0, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf((rsTransferencia!Estado = "R" Or rsTransferencia!Estado = "A"), 1, 0)
                            If cmdSave.Enabled Then
                                cmdSave.SetFocus
                            ElseIf cmdDelete.Enabled Then
                                cmdDelete.SetFocus
                            ElseIf cmdPrint.Enabled Then
                                cmdPrint.SetFocus
                            Else
                                pEnfocaTextBox txtNumero
                            End If
                        Else
                            'La información no pertenece a este departamento.
                            MsgBox SIHOMsg(782), vbOKOnly + vbExclamation, "Mensaje"
                            txtNumero_GotFocus
                            pEnfocaTextBox txtNumero
                        End If
                    Else
                        txtNumero_GotFocus
                        txtNumero.Text = CStr(flngSigTransferencia)
                        SendKeys vbTab
                    End If
                End If
            Else
                If Me.ActiveControl.Name = "vsfTransferencia" Then
                    vsfTransferencia.SetFocus
                Else
                    SendKeys vbTab
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub


Private Sub pHabilitaTitulo(intDepto As Integer, intPersona As Integer, intCorte As Integer, intDeptoRecibe As Integer, intPersonaRecibe As Integer, intCorteRecibe As Integer)
On Error GoTo NotificaError

    lblDeptoTransfiere.Enabled = intDepto = 1
    lblPersonaTransfiere.Enabled = intPersona = 1
    lblCorteTransfiere.Enabled = intCorte = 1
    lblDeptoRecibe.Enabled = intDeptoRecibe = 1
    lblPersonaRecibe.Enabled = intPersonaRecibe = 1
    lblCorteRecibe.Enabled = intCorteRecibe = 1

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaTitulo"))
End Sub

Private Sub pMuestra(strIdTransferencia As String)
On Error GoTo NotificaError

    vgstrParametrosSP = strIdTransferencia & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "0" & "|" & "-1" & "|" & "-1"
    Set rsTransferencia = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelTransferencia")
    With rsTransferencia
        lblnConsulta = .RecordCount <> 0
        'If .RecordCount <> 0 Then
        If lblnConsulta Then
            txtNumero.Text = !IdTransferencia
            cboTipo.ListIndex = flngLocalizaCbo(cboTipo, CStr(!IdTipo))
            
            '----- (CR) Agregado para caso 7442 -----'
            cboDeptoTransfiere.ListIndex = flngLocalizaCboTxt(cboDeptoTransfiere, !DepartamentoRegistra)
            cboDeptoTransfiere.Enabled = True
            pMkTextAsignaValor mskFechaIni, Format(!FechaCorte, "dd/mm/yyyy")
            pMkTextAsignaValor mskFechaFinal, Format(!FechaCorte, "dd/mm/yyyy")
            pOcultaControles False
            '----------------------------------------'
            
            lblFecha.Enabled = True
            mskFecha.Enabled = True
            mskFecha.Mask = ""
            mskFecha.Text = !fecha
            mskFecha.Mask = "##/##/####"
            
            lblEstado.Caption = !TipoEstado
            lblEstado.ForeColor = IIf(Trim(!Estado) = "C", llngColorCanceladas, llngColorActivas)
            
            lblDatoDeptoTransfiere.Caption = !DepartamentoRegistra
            lblDatoPersonaTransfiere.Caption = !EmpleadoRegistra
            
            cboCorteTransfiere.Clear
            If !IdCorte <> 0 Then
                cboCorteTransfiere.AddItem CStr(!IdCorte) & " - " & Format(!FechaCorte, "dd/mmm/yyyy hh:nn") & " - " & IIf(fblnCorteAbierto(!IdCorte), "ABIERTO", "CERRADO")
                cboCorteTransfiere.ListIndex = 0
            End If
            
            cboDeptoRecibe.Clear
            txtDatoPersonaRecibe.Text = ""
            txtDatoPersonaRecibe.Enabled = False
            cboDatoCorteRecibe.Clear
            cboDatoCorteRecibe.Enabled = False
            If !IdTipo = cintIdFondoFijo Then
                'Fondo fijo
                pHabilitaTitulo 1, 1, 1, 0, 0, 0
                pConfiguraVsf "FF"
                
                Do While Not .EOF
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfIdForma) = !idForma
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfFormaPago) = !formapago
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCantidadFondo) = FormatCurrency(!cantidad, 2)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfTipoCambio) = FormatCurrency(!TipoCambio, 2)
                    vsfTransferencia.Rows = vsfTransferencia.Rows + 1
                    .MoveNext
                Loop
                vsfTransferencia.Rows = vsfTransferencia.Rows - 1
                
                .MoveFirst
            End If
            
            If !IdTipo = cintIdTransDepto Then
                'Transferencia a departamento
                pHabilitaTitulo 1, 1, 1, 1, 1, 1
                
                If !IdDeptoRecibe = lintNumeroDepartamento Then
                    pCargaEquivalencia !IdDepto, !IdDeptoRecibe
                End If
            
                cboDeptoRecibe.AddItem !DepartamentoRecibe
                cboDeptoRecibe.ListIndex = 0
                
                txtDatoPersonaRecibe.Text = !EmpleadoRecibe
                txtDatoPersonaRecibe.Enabled = False
                If !Estado = "R" Then
                    cboDatoCorteRecibe.AddItem CStr(!IdCorteRecibe) & " - " & Format(!FechaCorteRecibe, "dd/mmm/yyyy hh:mm")
                    cboDatoCorteRecibe.ItemData(cboCorteTransfiere.newIndex) = 1
                    'cboDatoCorteRecibe.ListIndex = 0
                    cboDatoCorteRecibe.Enabled = False
                End If
                
                pConfiguraVsf "CT"
                
                Do While Not .EOF
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfIdForma) = !idForma
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfFormaPago) = !formapago
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfReferencia) = IIf(Trim(!folio) = "0", " ", !folio)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCantidadTransferida) = FormatCurrency(!cantidad, 2)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfTipoCambio) = FormatCurrency(!TipoCambio, 2)
                    vsfTransferencia.Rows = vsfTransferencia.Rows + 1
                    .MoveNext
                Loop
                vsfTransferencia.Rows = vsfTransferencia.Rows - 1
                
                .MoveFirst
            End If
            'AAT
                If !IdTipo = cintIdTransBanco Then
                'Transferencia a banco
                pHabilitaTitulo 1, 1, 1, 1, 0, 0
                
                'cboDeptoRecibe.AddItem !Banco
                'cboDeptoRecibe.ItemData(cboDeptoRecibe.newIndex) = !IdBanco
                'cboDeptoRecibe.ListIndex = 0
                
                pConfiguraVsf "CB"
                
                Do While Not .EOF
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfIdForma) = !idForma
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfFormaPago) = !formapago
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCtaFuente) = !ctacontable
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfReferencia) = IIf(Trim(!folio) = "0", " ", !folio)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCantidadTransferida) = FormatCurrency(!cantidad, 2)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfTipoCambio) = FormatCurrency(!TipoCambio, 2)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfIdBanco) = !IdBanco
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfBanco) = !Banco
                    vsfTransferencia.Rows = vsfTransferencia.Rows + 1
                    .MoveNext
                Loop
                vsfTransferencia.Rows = vsfTransferencia.Rows - 1
                
                .MoveFirst
            End If
            
            If !IdTipo = cintIdTransFormas Then
                'Cambio de formas de pago
                pHabilitaTitulo 1, 1, 1, 1, 0, 0
                
                cboDeptoRecibe.AddItem !Formapago2
                cboDeptoRecibe.ItemData(cboDeptoRecibe.newIndex) = !idformapago
                cboDeptoRecibe.ListIndex = 0
                cboDatoCorteRecibe.AddItem !Banco2
                cboDatoCorteRecibe.ItemData(cboDeptoRecibe.newIndex) = !IdBanco2
                cboDatoCorteRecibe.ListIndex = 0
                txtDatoPersonaRecibe = IIf(!Referencia = 0, "", !Referencia)
                txtDatoPersonaRecibe.Enabled = True
                cboDatoCorteRecibe.Enabled = True
                lblPersonaRecibe.Enabled = True
                lblCorteRecibe.Enabled = True
                
                pConfiguraVsf "CF"
                
                Do While Not .EOF
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfIdForma) = !idForma
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfFormaPago) = !formapago
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCtaFuente) = !ctacontable
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfReferencia) = IIf(Trim(!folio) = 0, " ", !folio)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfCantidadTransferida) = FormatCurrency(!cantidad, 2)
                    vsfTransferencia.TextMatrix(vsfTransferencia.Rows - 1, cintColvsfTipoCambio) = FormatCurrency(!TipoCambio, 2)
                    vsfTransferencia.Rows = vsfTransferencia.Rows + 1
                    .MoveNext
                Loop
                vsfTransferencia.Rows = vsfTransferencia.Rows - 1
                
                llngNumPoliza = flngPolizaTransferencia(CLng(strIdTransferencia))
                cmdPrintPoliza.Enabled = (llngNumPoliza <> 0)
                
                .MoveFirst
            End If
            
            fraMaestro.Enabled = False
            fraDetalle.Enabled = True
        End If
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestra"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    lintNumeroDepartamento = vgintNumeroDepartamento '(CR) Asignar el valor del departamento a la variable local
    
    '----- Permisos para las transferencias -----'
    lblnPermisoFondo = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionFondoFijo, "E", True)
    lblnPermisoTranDepto = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionTransDepto, "E", True)
    lblnPermisoTranBanco = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionTransBanco, "E", True)
    lblnPermisoCambioformas = fblnRevisaPermiso(vglngNumeroLogin, cintNumOpcionCambioformas, "E", True)
    
    '----- Agregar tipos de transferencia -----'
    cboTipo.AddItem "Fondo fijo", 0
    cboTipo.ItemData(cboTipo.newIndex) = cintIdFondoFijo
    cboTipo.AddItem "Transferencia a departamento", 1
    cboTipo.ItemData(cboTipo.newIndex) = cintIdTransDepto
    cboTipo.AddItem "Transferencia a banco", 2
    cboTipo.ItemData(cboTipo.newIndex) = cintIdTransBanco
    cboTipo.AddItem "Cambio de formas de pago", 3
    cboTipo.ItemData(cboTipo.newIndex) = cintIdTransFormas
    
    '----- Agregar tipos de búsqueda -----'
    cboTipoBus.AddItem "<Todas>", 0
    cboTipoBus.ItemData(cboTipoBus.newIndex) = -1
    cboTipoBus.AddItem "Fondo fijo", 1
    cboTipoBus.ItemData(cboTipoBus.newIndex) = cintIdFondoFijo
    cboTipoBus.AddItem "Transferencia a departamento", 2
    cboTipoBus.ItemData(cboTipoBus.newIndex) = cintIdTransDepto
    cboTipoBus.AddItem "Transferencia a banco", 3
    cboTipoBus.ItemData(cboTipoBus.newIndex) = cintIdTransBanco
    cboTipoBus.AddItem "Cambio de formas de pago", 4
    cboTipoBus.ItemData(cboTipoBus.newIndex) = cintIdTransFormas
    cboTipoBus.AddItem "Transferencia varios bancos", 5
    
    vgstrParametrosSP = CStr(-1) & "|" & "1" & "|" & "*" & "|" & vgintClaveEmpresaContable
    Set rsDepartamentos = frsEjecuta_SP(vgstrParametrosSP, "Sp_GnSelDepartamento")
    
    Set rsBancos = frsEjecuta_SP("-1|" & CStr(vgintClaveEmpresaContable), "Sp_CpSelBanco")
    
    vgstrParametrosSP = CStr(-1) & "|" & "-1" & "|" & CStr(-1) & "|" & CStr(lintNumeroDepartamento) & "|" & "1" & "|" & "*"
    Set rsFormasPago = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormapago")
    
    lstrFormasPago = fstrFormasPago
    
    ldblTipoCambioVenta = fdblTipoCambio(fdtmServerFecha, "V")
    
    SSTabTransfencia.Tab = 0

    fraNumero.BorderStyle = 0
    
    pConfiguraVsf "II"
    
    '- (CR) Caso 7442: Agregado para permitir seleccionar el departamento del cual se realizará la transferencia -'
    If cgstrModulo = "PV" Then
        llngNumOpcionSelDeptoCorte = 2510
    ElseIf cgstrModulo = "CC" Then
        llngNumOpcionSelDeptoCorte = 2511
    End If
    lblnPermisoCambioFormasDepto = fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcionSelDeptoCorte, "E", True)
    pCargaDeptoTransfiere True
    cboDeptoTransfiere.Enabled = False 'Por defecto inhabilitado
    pObtenerPrimerUltimoDia fdtmServerFecha 'Valores iniciales del rango de búsqueda de los cortes
    
    '-------------------------------------------------------------------------------------------------------------'

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Function fblnDatosGrabar()
On Error GoTo NotificaError

    Dim intcontador As Long
    Dim intTotal As Long
    Dim lngIdFormaDestino As Long

    fblnDatosGrabar = True
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se haya seleccionado un tipo de transferencia: (fondo fijo, a departamento o banco) '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If cboTipo.ListIndex = -1 Then
        fblnDatosGrabar = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
        cboTipo.SetFocus
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se tenga permisos de escritura, segun el proceso seleccionado '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If (cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo And Not lblnPermisoFondo) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And Not lblnPermisoTranDepto) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco And Not lblnPermisoTranBanco) _
        Or (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas And Not lblnPermisoCambioformas) Then
            fblnDatosGrabar = False
            '¡El usuario no tiene permiso para grabar datos!
            MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que tenga una fecha válida para transferencias a banco o departamento '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If Not IsDate(mskFecha.Text) Then
            fblnDatosGrabar = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
            mskFecha.SetFocus
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que tenga una fecha menor o igual a la actual para transferencias a banco '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas) And CDate(mskFecha.Text) > fdtmServerFecha Then
            fblnDatosGrabar = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
            mskFecha.SetFocus
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que el periodo contable donde generará la transferencia a banco no esté cerrado '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
            If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(mskFecha.Text)), Month(CDate(mskFecha.Text))) Then
                'El periodo contable esta cerrado
                MsgBox SIHOMsg(209), vbExclamation + vbOKOnly, "Mensaje"
                
                fblnDatosGrabar = False
                mskFecha.SetFocus
            End If
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se haya capturado algo en el grid para el fondo fijo: '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo And vsfTransferencia.Rows - 1 = 1 Then
            fblnDatosGrabar = False
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
            vsfTransferencia.Col = cintColvsfFormaPago
            vsfTransferencia.Row = 1
            vsfTransferencia.SetFocus
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se haya capturado algo en el grid para el fondo fijo: '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
            With vsfTransferencia
                intcontador = 1
                Do While intcontador <= .Rows - 2 And fblnDatosGrabar
                    fblnDatosGrabar = Val(Format(.TextMatrix(intcontador, cintColvsfCantidadFondo), cstrNumero)) <> 0
                    intcontador = intcontador + 1
                Loop
                
                If Not fblnDatosGrabar Then
                    '¡No ha ingresado datos!
                    MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                    .Col = cintColvsfFormaPago
                    .Row = 1
                    .SetFocus
                End If
            End With
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se haya seleccionado el departamento o banco o forma de pago al que se le transferirá '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        'AAT
        'If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
            If cboDeptoRecibe.ListIndex = -1 Then
                fblnDatosGrabar = False
                '¡Dato no válido, seleccione un valor de la lista!
                MsgBox SIHOMsg(3), vbOKOnly + vbExclamation, "Mensaje"
                cboDeptoRecibe.SetFocus
            End If
        End If
    End If

    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que se haya capturado algo para transferencia a departamentoo o banco '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If (cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas) And Not lblnConsulta Then
            With vsfTransferencia
                .Col = cintColvsfTransferir
                intTotal = 0
                intcontador = 1
                Do While intcontador <= .Rows - 1 And fblnDatosGrabar
                    .Row = intcontador
                    If .CellChecked = flexChecked Then
                        If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) > 0 Then
                            If Val(Format(.TextMatrix(intcontador, cintColvsfCantidadDisponible), cstrNumero)) >= Val(Format(.TextMatrix(intcontador, cintColvsfCantidadTransferir), cstrNumero)) Then
                                'Si es dólares, pongo el tipo de cambio
                                If Val(.TextMatrix(intcontador, cintColvsfMoneda)) = 0 Then
                                    .TextMatrix(intcontador, cintColvsfTipoCambio) = ldblTipoCambioVenta
                                Else
                                    .TextMatrix(intcontador, cintColvsfTipoCambio) = 0
                                End If
                                intTotal = intTotal + 1
                            Else
                                fblnDatosGrabar = False
                            End If
                        Else
                            fblnDatosGrabar = False
                        End If
                    End If
                    intcontador = intcontador + 1
                Loop
                
                If Not fblnDatosGrabar Then
                    'Dato incorrecto.
                    MsgBox SIHOMsg(406), vbExclamation + vbOKOnly, "Mensaje"
                    .Col = cintColvsfCantidadTransferir
                    .SetFocus
                Else
                    If intTotal = 0 Then
                        fblnDatosGrabar = False
                        '¡No ha ingresado datos!
                        MsgBox SIHOMsg(2), vbExclamation + vbOKOnly, "Mensaje"
                        .Col = cintColvsfTransferir
                        .Row = 1
                        .SetFocus
                    End If
                End If
            End With
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    ' Que existan formas de pago equivalentes para transferencia a departamento para el departamento que recibe '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
    If fblnDatosGrabar Then
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto And lblnConsulta Then
            With vsfTransferencia
                intcontador = 1
                intTotal = 0
                Do While intcontador <= .Rows - 1
                    If rsFormaEquivalente.RecordCount = 0 Then
                        pPintaRojo intcontador
                        intTotal = intTotal + 1
                    Else
                        lngIdFormaDestino = flngIdDestino(.TextMatrix(intcontador, cintColvsfIdForma))
                        If lngIdFormaDestino = 0 Then
                            pPintaRojo intcontador
                            intTotal = intTotal + 1
                        Else
                            .TextMatrix(intcontador, cintColvsfCveFormaDestino) = lngIdFormaDestino
                            'Estas dos variables toman valor en <flngIdDestino>:
                            .TextMatrix(intcontador, cintColvsfCtaFuente) = llngCtaFormaFuente
                            .TextMatrix(intcontador, cintColvsfCtaDestino) = llngCtaFormaDestino
                        End If
                    End If
                    intcontador = intcontador + 1
                Loop
                
                If intTotal <> 0 Then
                    fblnDatosGrabar = False
                    'No se encontró la forma de pago equivalente para el departamento seleccionado.
                    MsgBox SIHOMsg(783), vbExclamation + vbOKOnly, "Mensaje"
                End If
            End With
        End If
    End If
    
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    ' Que se introduzca la contraseña correcta '
    '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-'
    If fblnDatosGrabar Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        fblnDatosGrabar = llngPersonaGraba <> 0
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosGrabar"))
End Function

Private Function flngIdDestino(lngCveforma As Long) As Long
On Error GoTo NotificaError

    flngIdDestino = 0
    rsFormaEquivalente.MoveFirst
    Do While Not rsFormaEquivalente.EOF And flngIdDestino = 0
        If rsFormaEquivalente!CveFuente = lngCveforma Then
            flngIdDestino = rsFormaEquivalente!CveDestino
            llngCtaFormaFuente = rsFormaEquivalente!CtaFuente
            llngCtaFormaDestino = rsFormaEquivalente!CtaDestino
        End If
        rsFormaEquivalente.MoveNext
    Loop

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngIdDestino"))
End Function

Private Sub pPintaRojo(intRengon As Long)
On Error GoTo NotificaError

    Dim intColumna As Integer
    
    vsfTransferencia.Row = intRengon
    For intColumna = 1 To vsfTransferencia.Cols - 1
        vsfTransferencia.Col = intColumna
        vsfTransferencia.CellForeColor = llngColorCanceladas
    Next intColumna

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPintaRojo"))
End Sub

Private Function fstrFormasPago() As String
On Error GoTo NotificaError
    fstrFormasPago = ""

    vgstrParametrosSP = CStr(-1) & "|" & "0" & "|" & CStr(-1) & "|" & CStr(lintNumeroDepartamento) & "|" & "1" & "|" & "E"
    Set rsFormas = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFormaPago")
    If rsFormas.RecordCount <> 0 Then
        Do While Not rsFormas.EOF
            fstrFormasPago = fstrFormasPago & "|#" & CStr(rsFormas!intFormaPago) & ";" & Trim(rsFormas!chrdescripcion)
            rsFormas.MoveNext
        Loop
    End If

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrFormasPago"))
End Function

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer, intPrint As Integer)
On Error GoTo NotificaError

    cmdTop.Enabled = intTop = 1
    cmdBack.Enabled = intBack = 1
    cmdLocate.Enabled = intlocate = 1
    cmdNext.Enabled = intNext = 1
    cmdEnd.Enabled = intEnd = 1
    cmdSave.Enabled = intSave = 1
    cmdDelete.Enabled = intDelete = 1
    cmdPrint.Enabled = intPrint = 1

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError

    If SSTabTransfencia.Tab <> 0 Then
        Cancel = 1
        SSTabTransfencia.Tab = 0
        txtNumero.SetFocus
    Else
        If SSTabTransfencia.Tab = 0 Then
            If Me.ActiveControl.Name <> "txtNumero" Then
                If (lblnConsulta Or cmdSave.Enabled) And lblnCorteValido Then
                    '¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        Cancel = 1
                        txtNumero.SetFocus
                    Else
                        Cancel = 1
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub grdTransferencia_DblClick()
On Error GoTo NotificaError

    If Val(grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)) <> 0 Then
        pMuestra grdTransferencia.TextMatrix(grdTransferencia.Row, cintColgrdId)
        pHabilita 1, 1, 1, 1, 1, IIf(rsTransferencia!IdTipo = cintIdTransDepto And rsTransferencia!IdDeptoRecibe = lintNumeroDepartamento And rsTransferencia!Estado = "A", 1, 0), IIf(Trim(rsTransferencia!Estado) = "A" And rsTransferencia!IdDepto = lintNumeroDepartamento, 1, 0), IIf((rsTransferencia!Estado = "R" Or rsTransferencia!Estado = "A"), 1, 0)
        SSTabTransfencia.Tab = 0
        cmdLocate.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdTransferencia_DblClick"))
End Sub

Private Sub grdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        grdTransferencia_DblClick
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdTransferencia_KeyDown"))
End Sub

Private Sub mskFecha_GotFocus()
On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
    pSelMkTexto mskFecha

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecha_GotFocus"))
End Sub

Private Sub mskFechaFin_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFechaFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_LostFocus()
On Error GoTo NotificaError

    If Trim(mskFechaFin.ClipText) = "" Then
        mskFechaFin.Mask = ""
        mskFechaFin.Text = fdtmServerFecha
        mskFechaFin.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaFinal_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFechaFinal

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFinal_GotFocus"))
End Sub

Private Sub mskFechaFinal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsDate(mskFechaIni.Text) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            mskFechaIni.SetFocus
        Else
            If Not IsDate(mskFechaFinal.Text) Then
                '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                mskFechaFinal.SetFocus
            Else
                If CDate(mskFechaIni.Text) > CDate(mskFechaFinal.Text) Then
                    '¡Rango de fechas no válido!
                    MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
                    mskFechaFinal.SetFocus
                Else
                    pCargaCortesDepto True
                End If
            End If
        End If
    End If
End Sub

Private Sub mskFechaFinal_LostFocus()
On Error GoTo NotificaError

    If Trim(mskFechaFinal.ClipText) = "" Then
        mskFechaFinal.Mask = ""
        mskFechaFinal.Text = fdtmServerFecha
        mskFechaFinal.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFinal_LostFocus"))
End Sub

Private Sub mskFechaIni_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFechaIni

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_GotFocus"))
End Sub

Private Sub mskFechaIni_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskFechaIni.ClipText) = "" Then
        mskFechaIni.Mask = ""
        mskFechaIni.Text = fdtmServerFecha
        mskFechaIni.Mask = "##/##/####"
    End If
    
    pConfiguraVsf "FP"
    If cboCorteTransfiere.ListCount > 0 Then cboCorteTransfiere.ListIndex = -1
    'pCargaCortesDepto True

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaIni_LostFocus"))
End Sub

Private Sub mskFechaInicio_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFechaInicio

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskFechaInicio.ClipText) = "" Then
        mskFechaInicio.Mask = ""
        mskFechaInicio.Text = fdtmServerFecha
        mskFechaInicio.Mask = "##/##/####"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_LostFocus"))
End Sub

Private Sub txtDatoPersonaRecibe_GotFocus()
    pSelTextBox txtDatoPersonaRecibe
End Sub

Private Sub txtDatoPersonaRecibe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If fblnCanFocus(CboDatoCorteRecibe) Then CboDatoCorteRecibe.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtNumero_GotFocus()
On Error GoTo NotificaError

    pLimpia
    pConfiguraVsf "II"
    pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    
    pSelTextBox txtNumero
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumero_GotFocus"))
End Sub

Private Sub pConfiguraGrd()
On Error GoTo NotificaError

    With grdTransferencia
        .Visible = False
        
        .Clear
        .Cols = cintgrdTransferenciaCols
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = cstrFormatgrd
        .ColWidth(0) = 100
        .ColWidth(cintColgrdFecha) = 1100
        .ColWidth(cintColgrdTipo) = 2100
        .ColWidth(cintColgrdId) = 900
        .ColWidth(cintColgrdEmpleado) = 3000
        .ColWidth(cintColgrdEmpleadoCancela) = 3000
        .ColWidth(cintColgrdEstado) = 1300

        .ColAlignment(cintColgrdFecha) = flexAlignLeftCenter
        .ColAlignment(cintColgrdTipo) = flexAlignLeftCenter
        .ColAlignment(cintColgrdId) = flexAlignRightCenter
        .ColAlignment(cintColgrdEmpleado) = flexAlignLeftCenter
        .ColAlignment(cintColgrdEmpleadoCancela) = flexAlignLeftCenter
        .ColAlignment(cintColgrdEstado) = flexAlignLeftCenter

        .ColAlignmentFixed(cintColgrdFecha) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColgrdTipo) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColgrdId) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColgrdEmpleado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColgrdEmpleadoCancela) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColgrdEstado) = flexAlignCenterCenter

        .Visible = True
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrd"))
End Sub

Private Sub pLimpia()
On Error GoTo NotificaError

    lblnConsulta = False
    
    fraMaestro.Enabled = True
    fraDetalle.Enabled = False
    
    txtNumero.Text = CStr(flngSigTransferencia)
    
    cboTipo.ListIndex = -1
    
    lblFecha.Enabled = True
    mskFecha.Enabled = True
    mskFecha.Mask = ""
    mskFecha.Text = fdtmServerFecha
    mskFecha.Mask = "##/##/####"
    
    lblEstado.Caption = ""
    
    lblDeptoTransfiere.Caption = "Departamento"
    lblPersonaTransfiere.Caption = "Persona"
    lblCorteTransfiere.Caption = "Corte"

    lblDatoDeptoTransfiere.Caption = vgstrNombreDepartamento
    lblDatoPersonaTransfiere.Caption = ""
    
    cboCorteTransfiere.Clear
    
    lblDeptoRecibe.Caption = "Departamento"
    lblPersonaRecibe.Caption = "Persona"
    lblCorteRecibe.Caption = "Corte"
    
    cboDeptoRecibe.Clear
    txtDatoPersonaRecibe.Text = ""
    txtDatoPersonaRecibe.Enabled = False
    cboDatoCorteRecibe.Clear
    cboDatoCorteRecibe.Enabled = False
    
    vsfTransferencia.Clear
    cmdBorrar.Enabled = False
    
    'Datos de la consulta de transferencias:
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = fdtmServerFecha
    mskFechaInicio.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    
    cboTipoBus.ListIndex = 0
    
    '------ (CR) AGREGADOS CASO 7442 ------'
    pCargaDeptoTransfiere False
    pObtenerPrimerUltimoDia fdtmServerFecha
    pOcultaControles False
    cmdPrintPoliza.Enabled = False
    '--------------------------------------'
    
    pCarga False
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Public Function flngSigTransferencia() As Long
On Error GoTo NotificaError
    
    flngSigTransferencia = 1
    frsEjecuta_SP "", "Sp_PvSelIdTransferencia", False, flngSigTransferencia

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngSigTransferencia"))
End Function

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtNumero_KeyPress"))
End Sub

'AAT
Public Sub pCargaBancosVsf()
    'Realizamos la carga lista bancos para llenar el combo que sera utiliza en la grid
    
    itemsCombo = ""
    If rsBancos.RecordCount <> 0 Then
        rsBancos.MoveFirst
        Do While Not rsBancos.EOF
            If rsBancos!BITESTATUS = 1 Then
                itemsCombo = itemsCombo & "|#" & rsBancos!tnynumerobanco & ";" & rsBancos!VCHNOMBREBANCO & ""
            End If
            rsBancos.MoveNext
        Loop
    End If
End Sub

Private Sub vsfTransferencia_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo NotificaError

    Dim lngRenglonForma As Long
    'Dim intcontador As Integer
    Dim intcontador As Long
    Dim dblTipoCambio As Double
    Dim intmsgvalue  As Integer
    'AAT
    Dim intOldRow As Long
    Dim intRowDel As Long
    

    With vsfTransferencia
        'AAT
        intOldRow = .Row
        intRowDel = 0
        '*-*-*-*-*-*-*-*-*-*-*-*-*'
        '*-     FONDO FIJO      -*'
        '*-*-*-*-*-*-*-*-*-*-*-*-*'
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
            If Col = cintColvsfFormaPago And .ComboIndex <> -1 Then
                'Buscar si ya existe esta forma de pago en el grid:
                lngRenglonForma = flngExisteDatoColVsf(vsfTransferencia, cintColvsfIdForma, .ComboData(.ComboIndex))
                If lngRenglonForma = -1 Then
                    dblTipoCambio = IIf(fblnDolares(.ComboData(.ComboIndex)), ldblTipoCambioVenta, 0)
                    
                    .TextMatrix(Row, cintColvsfFormaPago) = .ComboItem
                    .TextMatrix(Row, cintColvsfIdForma) = .ComboData(.ComboIndex)
                    .TextMatrix(Row, cintColvsfTipoCambio) = dblTipoCambio
                    .Col = cintColvsfCantidadFondo
                    
                    If Row = .Rows - 1 Then
                        .Rows = .Rows + 1
                    End If
                    cmdBorrar.Enabled = True
                Else
                    If lngRenglonForma <> Row Then
                        'Este dato ya está registrado.
                        MsgBox SIHOMsg(404), vbOKOnly + vbInformation, "Mensaje"
                    End If
                    
                    If Val(.TextMatrix(Row, cintColvsfIdForma)) <> 0 Then
                        'Si no está en el último renglón, posicionar en el dato que estaba:
                        intcontador = 0
                        Do While intcontador <= .ComboCount - 1
                            If .ComboData(intcontador) = Val(.TextMatrix(Row, cintColvsfIdForma)) Then
                                .TextMatrix(Row, cintColvsfFormaPago) = .ComboItem(intcontador)
                                .Col = cintColvsfCantidadFondo
                            End If
                            intcontador = intcontador + 1
                        Loop
                    End If
                End If
            End If
            
            If Col = cintColvsfCantidadFondo Then
                .TextMatrix(Row, cintColvsfCantidadFondo) = FormatCurrency(CStr(Val(Format(.TextMatrix(Row, cintColvsfCantidadFondo), cstrNumero))), 2)
                If Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                .Col = cintColvsfFormaPago
            End If
        End If
        
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
        '*-     TRANSFERENCIA A DEPARTAMENTO Ó BANCO, CAMBIO DE FORMAS DE PAGO    -*'
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*'
        If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransDepto Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
            'AAT - SELECCIONA CUENTA BANCARIA DE LA TRANSFERENCIA
            If Col = cintColvsfBanco Then
                'AAT -  asigno id del banco seleccionado en la fila a columna oculta
                vsfTransferencia.TextMatrix(.Row, cintColvsfIdBanco) = vsfTransferencia.ComboData(vsfTransferencia.ComboIndex)
                'Si es transferencia de efectivo muevo a la siguinete fila que se creo si se parte la transferencia
                'AAT - si es efectivo la transferencia elimina los hijos asociados a la fila padre
                lstrSentencia = "SELECT * FROM PVFORMAPAGO WHERE PVFORMAPAGO.INTFORMAPAGO = " & vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfIdForma)
                Set rs = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rs!chrTipo = "E" Then  'FORMA DE PAGO EFECTIVO
                    'Se posiciona en la siguiente fila
                    If Row < .Rows - 1 Then
                        .Row = .Row + 1
                        .Col = cintColvsfTransferir
                        .SetFocus
                    End If
                End If
            End If
        
            If Col = cintColvsfTransferir Then
                If .CellChecked = flexChecked Then
                    If Trim(.TextMatrix(.Row, cintColvsfReferencia)) <> "" Or cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas Then
                        .TextMatrix(.Row, cintColvsfCantidadTransferir) = .TextMatrix(.Row, cintColvsfCantidadDisponible)
                        If Row < .Rows - 1 Then
                            .Row = .Row + 1
                        End If
                    Else
                        .Col = cintColvsfCantidadTransferir 'SE POSICIONA EN COLUMNA IMPORTE A TRANSFERIR
                        .SetFocus
                    End If
    
                ElseIf vsfTransferencia.TextMatrix(Row, cintColvsfIdForma) <> "" Then
                'Desmarca casilla Transferencia
                    'AAT - si es efectivo la transferencia elimina los hijos asociados a la fila padre
                    lstrSentencia = "SELECT * FROM PVFORMAPAGO WHERE PVFORMAPAGO.INTFORMAPAGO = " & vsfTransferencia.TextMatrix(Row, cintColvsfIdForma)
                    Set rs = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rs!chrTipo = "E" Then
                        For intRowDel = (.Rows - 1) To (Row + 1) Step -1
                            'SI es fila hija de la transferencia en efectivo borrela
                            If Val(Trim(.TextMatrix(intRowDel, cintColvsfIdRow))) >= (Row) Then
                                .RemoveItem (intRowDel)
                            End If
                        Next intRowDel
                    End If
                    .TextMatrix(Row, cintColvsfCantidadTransferir) = ""
                    .TextMatrix(Row, cintColvsfBanco) = ""
                    .TextMatrix(Row, cintColvsfIdBanco) = ""
                End If
            Else
                If Col = cintColvsfCantidadTransferir Then
                    .Col = cintColvsfTransferir
                    If Val(Format(.TextMatrix(Row, cintColvsfCantidadTransferir), cstrNumero)) = 0 Then
                        .TextMatrix(Row, cintColvsfCantidadTransferir) = ""
                        .CellChecked = flexUnchecked
                    Else
                        '----- Agregado (CR): Revisar que la cantidad a transferir no sea mayor a la disponible -----'
                        If Val(Format(.TextMatrix(Row, cintColvsfCantidadTransferir), cstrNumero)) > Val(Format(.TextMatrix(Row, cintColvsfCantidadDisponible), cstrNumero)) Then
                            'La cantidad no puede ser mayor que el monto disponible
                            MsgBox SIHOMsg(919) & "el monto disponible.", vbOKOnly + vbExclamation, "Mensaje"
                            .Col = cintColvsfCantidadTransferir
                            .SetFocus
                            Exit Sub
                        Else
                            .TextMatrix(Row, cintColvsfCantidadTransferir) = FormatCurrency(CStr(Val(Format(.TextMatrix(Row, cintColvsfCantidadTransferir), cstrNumero))), 2)
                            .CellChecked = flexChecked
                            
                            'AAT - si es efectivo preguntar con ventana si va a dividir la transferencia en efectivo
                            lstrSentencia = "SELECT * FROM PVFORMAPAGO WHERE PVFORMAPAGO.INTFORMAPAGO = " & vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfIdForma)
                            Set rs = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
                            If rs!chrTipo = "E" Then
                                intOldRow = .Row
                               If Val(Format(.TextMatrix(Row, cintColvsfCantidadTransferir), cstrNumero)) < Val(Format(.TextMatrix(Row, cintColvsfCantidadDisponible), cstrNumero)) And cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
                                    'YES = 6, NO = 7
                                    intmsgvalue = MsgBox("¿Desea agregar nueva fila con la diferencia a transferir?", vbYesNo + vbExclamation, "Mensaje")
                                    If intmsgvalue = 6 Then
                                        'CREA NUEVA FILA PARA LA DIFERENCIA DE LA TRANSFERENCIA
                                        .Rows = .Rows + 1
                                        .TextMatrix(.Rows - 1, cintColvsfIdForma) = .TextMatrix(Row, cintColvsfIdForma)
                                        .TextMatrix(.Rows - 1, cintColvsfFormaPago) = .TextMatrix(Row, cintColvsfFormaPago)
                                        .TextMatrix(.Rows - 1, cintColvsfMoneda) = .TextMatrix(Row, cintColvsfMoneda)
                                        .TextMatrix(.Rows - 1, cintColvsfCtaFuente) = .TextMatrix(Row, cintColvsfCtaFuente)
                                        .TextMatrix(.Rows - 1, cintColvsfCantidadDisponible) = FormatCurrency((.TextMatrix(Row, cintColvsfCantidadDisponible) - .TextMatrix(Row, cintColvsfCantidadTransferir)), 2)
                                        .TextMatrix(.Rows - 1, cintColvsfIdRow) = .Row  'Asigna el ID de la fila anterior
                                        
                                        .Col = cintColvsfBanco 'SE POSICIONA EN COLUMNA IMPORTE A TRANSFERIR
                                        .Row = intOldRow
                                        .SetFocus
                                    Else
                                        'Se posiciona en la siguiente fila
                                        If Row < .Rows - 1 Then
                                            .Row = .Row + 1
                                            .Col = cintColvsfTransferir
                                            .SetFocus
                                        End If
                                    End If
                               Else
                                    'Se posiciona en la siguiente fila
                                    If Row < .Rows - 1 Then
                                        .Row = .Row + 1
                                        .Col = cintColvsfTransferir
                                        .SetFocus
                                    End If
                               End If
                            End If
                            
                            'AAT - Se posiciona en columna cuenta para seleccionar cuenta bancaria a transferir
                            .Col = cintColvsfBanco
                            .SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_AfterEdit"))
End Sub

Private Function fblnDolares(lngCveforma As Long) As Boolean
On Error GoTo NotificaError

    rsFormas.MoveFirst
    Do While Not rsFormas.EOF
        If rsFormas!intFormaPago = lngCveforma Then
            fblnDolares = rsFormas!BITPESOS = 0
        End If
        rsFormas.MoveNext
    Loop

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDolares"))
End Function

Private Sub vsfTransferencia_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo NotificaError

    '- Si el tipo es Cambio de formas de pago no permitir cambiar los datos del grid, únicamente marcar la columna Cambiar -'
    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransFormas And Col <> cintColvsfTransferir Then
        Cancel = True
    Else
        If Col = cintColvsfFormaPago And cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
            If Trim(lstrFormasPago) = "" Then
                'No existen formas de pago.
                MsgBox SIHOMsg(293), vbOKOnly + vbExclamation, "Mensaje"
            Else
                vsfTransferencia.ComboList = lstrFormasPago
                vsfTransferencia.ComboIndex = 0
            End If
        Else
            vsfTransferencia.ComboList = ""
        End If
    End If
    
    'AAT
    'Si es la columna de bancos, cargar el combo
    If Col = cintColvsfBanco And cboTipo.ItemData(cboTipo.ListIndex) = cintIdTransBanco Then
        vsfTransferencia.ComboList = itemsCombo
        vsfTransferencia.ComboIndex = 1
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_BeforeEdit"))
End Sub

Private Sub vsfTransferencia_ChangeEdit()
    'AAT
    If vsfTransferencia.Col = cintColvsfBanco Then
        vsfTransferencia_AfterEdit vsfTransferencia.Row, vsfTransferencia.Col
    End If
End Sub

Private Sub vsfTransferencia_Click()
On Error GoTo NotificaError
        If vsfTransferencia.RowSel <> -1 Then
            If vsfTransferencia.Col = cintColvsfTransferir Then
                vsfTransferencia.Row = vsfTransferencia.RowSel
                vsfTransferencia_AfterEdit vsfTransferencia.Row, vsfTransferencia.Col
            End If
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_Click"))
End Sub

Private Sub vsfTransferencia_GotFocus()
On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_GotFocus"))
End Sub

Private Sub vsfTransferencia_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlintPosicionPunto As Integer
    
    If Col = cintColvsfCantidadTransferir Or Col = cintColvsfCantidadFondo Then
        If Val(vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfIdForma)) <> 0 Then
            If Col = cintColvsfCantidadTransferir Then
                lstrSentencia = "SELECT * FROM PVFORMAPAGO WHERE  PVFORMAPAGO.INTFORMAPAGO = " & vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfIdForma)
                Set rs = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rs!chrTipo = "E" Then
'                If Val(vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfReferencia)) = 0 Then
                    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And _
                       Not KeyAscii = vbKeyReturn And Not KeyAscii = Asc(".") Then
                        KeyAscii = 7
                        
                    '- Validación de decimales: Solamente se permite un punto -'
                    ElseIf KeyAscii = Asc(".") Then
                        If vsfTransferencia.EditText <> "" Then
                            vlintPosicionPunto = InStr(1, vsfTransferencia.EditText, ".")
                            If vlintPosicionPunto > 0 And vsfTransferencia.EditSelText = "" Then
                                KeyAscii = 0
                            End If
                        End If
                    End If
                Else
                    KeyAscii = 7
                End If
            Else
                If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And _
                   Not KeyAscii = vbKeyReturn And Not KeyAscii = Asc(".") Then
                    KeyAscii = 7
                End If
            End If
        Else
            KeyAscii = 7
        End If
    Else
         KeyAscii = 7
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_KeyPressEdit"))
End Sub

Private Sub vsfTransferencia_RowColChange()
On Error GoTo NotificaError

    If cboTipo.ItemData(cboTipo.ListIndex) = cintIdFondoFijo Then
        cmdBorrar.Enabled = Val(vsfTransferencia.TextMatrix(vsfTransferencia.Row, cintColvsfIdForma)) <> 0
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":vsfTransferencia_RowColChange"))
End Sub

'************************* (CR) AGREGADOS: FUNCIONES Y PROCEDIMIENTOS PARA EL CAMBIO DE FORMAS DE PAGO *************************'
'- Verificar si la cuenta de la forma de pago pertenece a un banco activo, si lo es regresa el número de banco -'
Private Function flngEsCuentaBanco(llngNumeroCuenta As Long) As Long
On Error GoTo NotificaError

    Dim rsCuentaBanco As New ADODB.Recordset
    
    flngEsCuentaBanco = 0
    lstrSentencia = "SELECT tnyNumeroBanco FROM CpBanco WHERE bitEstatus = 1 AND intNumeroCuenta = " & llngNumeroCuenta
    Set rsCuentaBanco = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsCuentaBanco.RecordCount <> 0 Then
        flngEsCuentaBanco = rsCuentaBanco!tnynumerobanco
    End If
    rsCuentaBanco.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngEsCuentaBanco"))
End Function

'- Verificar si un corte está abierto -'
Private Function fblnCorteAbierto(llngNumCorte As Long) As Boolean
    fblnCorteAbierto = False
    vgstrParametrosSP = CStr(llngNumCorte) & "|0|" & fstrFechaSQL(mskFechaIni.Text) & "|" & fstrFechaSQL(mskFechaFinal.Text) & "|" & lintNumeroDepartamento & "|" & "-1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCORTE")
    If rs.RecordCount <> 0 Then
        fblnCorteAbierto = IsNull(rs!FechaCierra)
    End If
End Function

'- Verificar si se ha realizado una transferencia del mismo corte con la misma forma de pago -'
Private Function fblnTransferida(llngNumCorte As Long, llngCveForma As Long, ldblCantidad As Double) As Boolean
On Error GoTo NotificaError

    Dim rsTransferencia As New ADODB.Recordset
    
    lstrSentencia = "SELECT T.intNumCorte, T.intIdTransferencia, T.dtmFecha, F.intCveFormaPago, F.mnyCantidad " & _
                    " FROM PvTransferencia T INNER JOIN PvTransferenciaFormaPago F ON T.intIdTransferencia = F.intIdTransferencia " & _
                    " WHERE T.intTipo = 3 AND T.intNumCorte = " & llngNumCorte & " AND T.intCveDepartamento = " & lintNumeroDepartamento & _
                    " AND F.intCveFormaPago = " & llngCveForma & " AND ROUND(F.mnyCantidad, 2) = ROUND(" & ldblCantidad & ", 2)"
    Set rsTransferencia = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    fblnTransferida = rsTransferencia.RecordCount <> 0
    rsTransferencia.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnTransferida"))
End Function

'- Llena los departamentos para especificar de cuál se hará el movimiento de transferencia -'
Private Sub pCargaDeptoTransfiere(lblnInicia As Boolean)
On Error GoTo NotificaError

    If lblnInicia Then
        cboDeptoTransfiere.Clear
        If rsDepartamentos.RecordCount <> 0 Then
            pLlenarCboRs cboDeptoTransfiere, rsDepartamentos, 0, 1
        End If
    End If
    
    'Inicializa en el departamento del usuario que ingresó
    cboDeptoTransfiere.ListIndex = flngLocalizaCbo(cboDeptoTransfiere, CStr(vgintNumeroDepartamento))
    cboDeptoTransfiere.Enabled = lblnInicia
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaDeptoTransfiere"))
End Sub

'- Llena los cortes del departamento especificado para los cambios de formas de pago -'
Private Sub pCargaCortesDepto(Optional lblnMensaje As Boolean)
On Error GoTo NotificaError

    Dim lblnMostrarAbiertos As Boolean
    Dim lblnCortes As Boolean
    Dim lstrFechaIni As String
    Dim lstrFechaFin As String
    
    cboCorteTransfiere.Clear
    lblnCortes = False
    
    lblnMostrarAbiertos = (lintNumeroDepartamento = vgintNumeroDepartamento)
    lstrFechaIni = fstrFechaSQL(mskFechaIni.Text)
    lstrFechaFin = fstrFechaSQL(mskFechaFinal.Text)
    
    vgstrParametrosSP = -1 & "|1|" & lstrFechaIni & "|" & lstrFechaFin & "|" & lintNumeroDepartamento & "|" & -1
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelCorte")
    If rs.RecordCount <> 0 Then
        lblnCortes = True
        
        Do While Not rs.EOF
            'Si el corte está cerrado o el usuario es del mismo departamento y el periodo contable NO ha sido cerrado'
            If (Not IsNull(rs!FechaCierra) Or lblnMostrarAbiertos) And Not fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(rs!FechaAbre), Month(rs!FechaAbre)) Then
                'Agregar y ordenar elemento descendentemente (Sp_PvSelCorte regresa los cortes ordenados ascendentemente)
                pOrdenarLstCboBoxItem cboCorteTransfiere, CStr(rs!IdCorte) & " - " & Format(rs!FechaAbre, "dd/mmm/yyyy hh:mm") & " - " & UCase(rs!Estado), rs!IdCorte, False, True
            End If
            rs.MoveNext
        Loop
        
        If cboCorteTransfiere.ListCount > 0 Then
            'cboCorteTransfiere.ListIndex = 0
            
            cboCorteTransfiere.Enabled = True
        Else
            lblnCortes = False
        End If
    End If
    
    If Not lblnCortes And lblnMensaje Then
        MsgBox "No se encontraron cortes cerrados en este departamento.", vbOKOnly + vbExclamation, "Mensaje"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaCortesDepto"))
End Sub

'- Procedimiento para obtener el primer y último día de la fecha especificada -'
Private Sub pObtenerPrimerUltimoDia(ldtmFecha As Date)
    Dim ldtmPrimer As Date
    Dim ldtmUltimo As Date
      
    If lblnPermisoCambioFormasDepto Then
        ldtmPrimer = DateSerial(Year(ldtmFecha), Month(ldtmFecha) + 0, 1)
        'ldtmUltimo = DateSerial(Year(ldtmFecha), Month(ldtmFecha) + 1, 0)
    Else
        ldtmPrimer = ldtmFecha
    End If
    
    ldtmUltimo = ldtmFecha
    
    mskFechaIni.Mask = ""
    mskFechaIni.Text = ldtmPrimer
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFinal.Mask = ""
    mskFechaFinal.Text = ldtmUltimo
    mskFechaFinal.Mask = "##/##/####"
End Sub

'- Procedimiento para habilitar o deshabilitar controles según la forma -'
Private Sub pHabilitaControles(lblnHabilitar As Boolean)
    cboDeptoTransfiere.Enabled = lblnHabilitar
    lblDeptoTransfiere.Enabled = lblnHabilitar
    lblCorteTransfiere.Enabled = lblnHabilitar
    mskFechaIni.Enabled = lblnHabilitar
    mskFechaFinal.Enabled = lblnHabilitar
    lblHasta.Enabled = lblnHabilitar
    lblPersonaTransfiere.Enabled = lblnHabilitar
End Sub

'- Procedimiento para ocultar o mostrar controles según la forma -'
Private Sub pOcultaControles(lblnOcultar As Boolean)
    lblDatoPersonaTransfiere.Visible = Not lblnOcultar
    cboDeptoTransfiere.Visible = lblnOcultar
    mskFechaIni.Visible = lblnOcultar
    mskFechaFinal.Visible = lblnOcultar
    lblHasta.Visible = lblnOcultar
End Sub

'- Procedimiento para ordenar un combobox o un listbox de manera ascendente/descendente -'
Private Sub pOrdenarLstCboBoxItem(objControl As Object, ByVal Item As String, _
                                  Optional ByVal ItemData As Long = 0&, _
                                  Optional ByVal Ascending As Boolean = True, _
                                  Optional ByVal byItemData As Boolean = False, _
                                  Optional ByVal CaseSensitive As Boolean = False)

' objControl = Objeto en donde se agregarán los elementos. Puede ser un Combo o un Listbox
' Item = La cadena a ser agregada al objeto
' ItemData = La propiedad ItemData del nuevo elemento
' Ascending = True para ordenar ascendentemente; False para ordenar descendentemente
' byItemData = True para ordenar por la propiedad .ItemData, False para ordener por Item
' CaseSensitive = Si byItemData es True es ignorado

    Dim UB As Long, LB As Long, newIndex As Long
    Dim lComp As Long, lSortOrder As Long, lSortType As Long
    Dim Count As Long, lTestValue As Long, Index As Long
    
    Count = objControl.ListCount
    If Count = 0& Then ' Vacío, agregar como primer elemento
        objControl.AddItem Item
        objControl.ItemData(Count) = ItemData
        Exit Sub
    ElseIf Count < 0& Then ' No se pueden usar indices negativos
        Exit Sub
    Else
        Index = Count
    End If
    
    If Ascending Then lSortOrder = -1& Else lSortOrder = 1&
    If CaseSensitive Then lSortType = vbBinaryCompare Else lSortType = vbTextCompare
    
    For Index = Index To Count
        UB = Index
        LB = 1&
    
        If byItemData Then
            Do Until LB > UB
                newIndex = LB + ((UB - LB) \ 2&) - 1&
                lTestValue = objControl.ItemData(newIndex)
                If ItemData = lTestValue Then
                    lComp = 0&
                    Exit Do
                Else
                    If ItemData < lTestValue Then lComp = -1& Else lComp = 1&
                    If lComp = lSortOrder Then UB = newIndex Else LB = newIndex + 2&
                End If
            Loop
        Else
            Do Until LB > UB
                newIndex = LB + ((UB - LB) \ 2&) - 1&
                lComp = StrComp(Item, objControl.List(newIndex), lSortType)
                If lComp = 0& Then Exit Do
                If lComp = lSortOrder Then UB = newIndex Else LB = newIndex + 2&
            Loop
        End If

        If lComp = -lSortOrder Then newIndex = newIndex + 1&
        
        objControl.AddItem Item, newIndex
        objControl.ItemData(newIndex) = ItemData
    Next
End Sub

'----------------------------------------- PÓLIZAS -----------------------------------------'
'- Función modificada de >>flngInsertarPoliza<< para cambiar el departamento que inserta la póliza -'
Public Function flngInsertarPolizaMaestro(vldtmFecha As Date, vlstrTipoPoliza As String, vlStrConcepto As String, vlintDepartamento As Integer, vllngEmpleado As Long) As Long
On Error GoTo NotificaError
    
    Dim rsCnPoliza As New ADODB.Recordset
    
    lstrSentencia = "SELECT * FROM CnPoliza WHERE intNumeroPoliza = -1"
    Set rsCnPoliza = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
    With rsCnPoliza
        .AddNew
        !tnyclaveempresa = vgintClaveEmpresaContable
        !smiEjercicio = Year(CDate(vldtmFecha))
        !tnyMes = Month(CDate(vldtmFecha))
        !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, vlstrTipoPoliza, Year(vldtmFecha), Month(vldtmFecha), False)
        !dtmFechaPoliza = vldtmFecha
        !chrTipoPoliza = vlstrTipoPoliza
        !vchConceptoPoliza = IIf(Len(Trim(vlStrConcepto)) > 250, Mid(Trim(vlStrConcepto), 1, 250), Trim(vlStrConcepto))
        !smicvedepartamento = vlintDepartamento
        !INTCVEEMPLEADO = vllngEmpleado
        !vchNumero = " "
        !bitAsentada = 0
        .Update
    End With
    flngInsertarPolizaMaestro = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!intNumeroPoliza)
    rsCnPoliza.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngInsertarPolizaMaestro"))
End Function

Private Sub pIncluyeMovimiento(vllngxNumeroCuenta As Long, vldblxCantidad As Double, vlintxTipoMovto As Integer)
'================================='
' Valores:                        '
'   vlintxTipoMovto = 1 = cargo   '
'   vlintxTipoMovto = 0 = abono   '
'================================='
On Error GoTo NotificaError

    Dim vlblnEstaCuenta As Boolean
    Dim vllngContador As Long
        
    If apoliza(0).vllngNumeroCuenta = 0 Then
        apoliza(0).vllngNumeroCuenta = vllngxNumeroCuenta
        apoliza(0).vldblCantidadMovimiento = vldblxCantidad
        apoliza(0).vlintTipoMovimiento = vlintxTipoMovto
    Else
        vlblnEstaCuenta = False
        vllngContador = 0
        Do While vllngContador <= UBound(apoliza, 1) And Not vlblnEstaCuenta
            If apoliza(vllngContador).vllngNumeroCuenta = vllngxNumeroCuenta And apoliza(vllngContador).vlintTipoMovimiento = vlintxTipoMovto Then
                vlblnEstaCuenta = True
            End If
            If Not vlblnEstaCuenta Then
                vllngContador = vllngContador + 1
            End If
        Loop
        
        If vlblnEstaCuenta Then
            apoliza(vllngContador).vldblCantidadMovimiento = apoliza(vllngContador).vldblCantidadMovimiento + vldblxCantidad
        Else
            ReDim Preserve apoliza(UBound(apoliza, 1) + 1)
            apoliza(UBound(apoliza, 1)).vllngNumeroCuenta = vllngxNumeroCuenta
            apoliza(UBound(apoliza, 1)).vldblCantidadMovimiento = vldblxCantidad
            apoliza(UBound(apoliza, 1)).vlintTipoMovimiento = vlintxTipoMovto
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIncluyeMovimiento"))
End Sub

'- Guardar en CnDetallePoliza la póliza que corresponde a la transferencia, cuando la cuenta contable que recibe es diferente a la que transfiere -'
Private Sub pGuardarDetallePoliza(lngNumPoliza As Long)
On Error GoTo NotificaError
    
    Dim lngContador As Long
    Dim lngNumDetalle As Long

    lngContador = 0
    Do While lngContador <= UBound(apoliza(), 1)
        If apoliza(lngContador).vldblCantidadMovimiento <> 0 Then
            lngNumDetalle = flngInsertarPolizaDetalle(lngNumPoliza, apoliza(lngContador).vllngNumeroCuenta, apoliza(lngContador).vldblCantidadMovimiento, apoliza(lngContador).vlintTipoMovimiento)
        End If
        lngContador = lngContador + 1
    Loop

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGuardarDetallePoliza"))
    Unload Me
End Sub

'- Guardar la relación de la transferencia con la póliza generada -'
Private Sub pGuardarPolizaTransferencia(llngIdTransferencia As Long, llngNumPoliza As Long)
    lstrSentencia = "INSERT INTO PVTRANSFERENCIAPOLIZA VALUES(" & llngIdTransferencia & ", " & llngNumPoliza & ")"
    pEjecutaSentencia lstrSentencia
End Sub

'- Regresa el número de la póliza relacionada a la transferencia -'
Private Function flngPolizaTransferencia(llngIdTransferencia As Long) As Long
On Error GoTo NotificaError

    Dim rsTransPoliza As New ADODB.Recordset
    
    flngPolizaTransferencia = 0
    lstrSentencia = "SELECT INTNUMPOLIZA FROM PVTRANSFERENCIAPOLIZA WHERE intIDTransferencia = " & llngIdTransferencia
    Set rsTransPoliza = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsTransPoliza.RecordCount <> 0 Then
        flngPolizaTransferencia = rsTransPoliza!intNumPoliza
    End If
    rsTransPoliza.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngPolizaTransferencia"))
End Function
'-------------------------------------------------------------------------------------------'

'- CASO 7442: Cancelar el movimiento de las formas de pago de la transferencia -'
Private Sub pCancelaMovimiento(vllngNumTransferencia As Long, vllngCorteTransferencia As Long, vllngCorteActual As Long, vllngPersonaGraba As Long)
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim lstrFecha As String
    Dim ldblCantidad As Double
                 
    lstrSentencia = "SELECT MB.intFormaPago, MB.mnyCantidad, MB.mnyTipoCambio, FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco) AS IdBanco " & _
                    " FROM PvMovimientoBancoForma MB " & _
                    " INNER JOIN PvFormaPago FP ON MB.intFormaPago = FP.intFormaPago " & _
                    " LEFT  JOIN CpBanco B ON B.intNumeroCuenta = FP.intCuentaContable " & _
                    " WHERE TRIM(MB.chrTipoDocumento) = 'TR' AND MB.intNumDocumento = " & vllngNumTransferencia & _
                    " AND MB.intNumCorte = " & vllngCorteTransferencia & " AND MB.mnyCantidad > 0"
    Set rs = frsRegresaRs(lstrSentencia)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!chrTipo <> "C" Then
                lstrFecha = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) '- Fecha y hora del movimiento -'
                ldblCantidad = rs!MNYCantidad * (-1) 'Cantidad negativa para que se tome como abono
    
                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rs!intFormaPago & "|" & rs!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rs!MNYTIPOCAMBIO = 0, 1, 0) & "|" & rs!MNYTIPOCAMBIO & "|" & "CDE" & "|" & "TR" & "|" & vllngNumTransferencia & "|" & _
                                    vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & lstrFecha & "|" & "1" & "|" & cgstrModulo
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaMovimiento"))
End Sub
