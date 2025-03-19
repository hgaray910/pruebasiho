VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresupuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9795
   ScaleMode       =   0  'User
   ScaleWidth      =   13335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Descripción completa del cargo"
      Height          =   900
      Left            =   90
      TabIndex        =   125
      Top             =   8880
      Width           =   13155
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
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   240
         Width           =   12915
      End
   End
   Begin VB.Frame freBusqueda 
      Height          =   6885
      Left            =   3120
      TabIndex        =   57
      Top             =   10000
      Width           =   7155
      Begin VB.TextBox txtBuscaNombre 
         Height          =   300
         Left            =   870
         TabIndex        =   59
         Top             =   570
         Width           =   6165
      End
      Begin VB.CommandButton cmdAceptoBuscar 
         Caption         =   "&Aceptar"
         Height          =   450
         Left            =   2820
         TabIndex        =   61
         ToolTipText     =   "Guardar los cambios."
         Top             =   6285
         Width           =   1530
      End
      Begin VB.ListBox lstBuscaPresupuesto 
         Height          =   5130
         Left            =   135
         TabIndex        =   60
         Top             =   1050
         Width           =   6900
      End
      Begin VB.Label Label7 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   135
         TabIndex        =   62
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda de presupuestos"
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
         Height          =   195
         Left            =   105
         TabIndex        =   58
         Top             =   160
         Width           =   2355
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   30
         Top             =   120
         Width           =   7085
      End
   End
   Begin VB.Frame FraDuplicar 
      Height          =   6600
      Left            =   120
      TabIndex        =   84
      Top             =   10000
      Visible         =   0   'False
      Width           =   13155
      Begin VB.ComboBox cboTipoConvenio 
         Height          =   315
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         ToolTipText     =   "Seleccione el tipo de convenio"
         Top             =   875
         Width           =   4350
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   1
         Left            =   120
         MaxLength       =   5
         TabIndex        =   114
         Top             =   6120
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   90
         ToolTipText     =   "Seleccione la empresa de donde viene el paciente"
         Top             =   1200
         Width           =   4350
      End
      Begin VB.CommandButton cmdActualizaPrecios 
         Caption         =   "Actualizar precios"
         Height          =   375
         Index           =   1
         Left            =   11440
         TabIndex        =   95
         Top             =   1635
         Width           =   1575
      End
      Begin VB.CommandButton cmdAplicaDescuento 
         Caption         =   "Aplicar"
         Height          =   345
         Index           =   1
         Left            =   6360
         TabIndex        =   94
         Top             =   1635
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txtPorcentajeDescuento 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   5100
         MaxLength       =   3
         TabIndex        =   93
         Top             =   1635
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Index           =   1
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   88
         ToolTipText     =   "Seleccione el tipo de paciente"
         Top             =   555
         Width           =   4350
      End
      Begin VB.CheckBox chkTomarDescuentos 
         Caption         =   "Tomar en cuenta descuentos actuales asignados por tipo de paciente"
         Height          =   550
         Index           =   1
         Left            =   120
         TabIndex        =   92
         ToolTipText     =   "Si se toma en cuenta o no los descuentos."
         Top             =   1530
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CommandButton cmdSalirDuplicado 
         Caption         =   "Salir"
         Height          =   375
         Left            =   6480
         TabIndex        =   100
         ToolTipText     =   "Salir"
         Top             =   6085
         Width           =   940
      End
      Begin VB.CommandButton cmdCrearDuplicado 
         Caption         =   "Duplicar"
         Height          =   375
         Left            =   5535
         TabIndex        =   98
         ToolTipText     =   "Duplicar el presupuesto"
         Top             =   6085
         Width           =   940
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPresupuesto 
         Height          =   3915
         Index           =   1
         Left            =   120
         TabIndex        =   96
         ToolTipText     =   "Cargos que incluye el presupuesto"
         Top             =   2085
         Width           =   12900
         _ExtentX        =   22754
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   8
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FormatString    =   "|Descripción|Precio|Cantidad|Subtotal|Descuento|Monto|Tipo"
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   5
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblTipoConvenio 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de convenio"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   119
         Top             =   925
         Width           =   1245
      End
      Begin VB.Label lstEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   101
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   99
         Top             =   1680
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblDescuento 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   97
         Top             =   1680
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblTipoPaciente 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Duplicar presupuesto"
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
         Height          =   195
         Left            =   120
         TabIndex        =   87
         Top             =   180
         Width           =   4155
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   45
         Top             =   120
         Width           =   13075
      End
   End
   Begin VB.Frame fraDatosPaquete 
      Caption         =   "Datos para generar paquete"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   90
      TabIndex        =   105
      Top             =   6735
      Width           =   5010
      Begin VB.ComboBox cboTratamiento 
         Height          =   315
         ItemData        =   "frmPresupuestos.frx":0000
         Left            =   1650
         List            =   "frmPresupuestos.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   108
         ToolTipText     =   "Tipo de tratamiento"
         Top             =   585
         Width           =   3255
      End
      Begin VB.ComboBox cboConceptoFactura 
         Height          =   315
         Left            =   1650
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   107
         ToolTipText     =   "Concepto de facturación utilizado para el paquete"
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtNumPaquete 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   106
         ToolTipText     =   "Número de paquete"
         Top             =   930
         Width           =   1230
      End
      Begin VB.Label lblTratamiento 
         AutoSize        =   -1  'True
         Caption         =   "Tratamiento"
         Height          =   195
         Left            =   120
         TabIndex        =   111
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lblConceptoFactura 
         AutoSize        =   -1  'True
         Caption         =   "Concepto de factura"
         Height          =   195
         Left            =   120
         TabIndex        =   110
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label lblNumPaquete 
         AutoSize        =   -1  'True
         Caption         =   "Número de paquete"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   109
         Top             =   975
         Width           =   1410
      End
   End
   Begin VB.Frame fraMotivoNoAutorizado 
      Caption         =   "Motivo"
      Enabled         =   0   'False
      Height          =   1275
      Left            =   5160
      TabIndex        =   104
      Top             =   6735
      Width           =   4245
      Begin VB.TextBox txtMotivoNoAutorizado 
         Height          =   945
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   117
         ToolTipText     =   "Motivo por el cual se cambia el estado del presupuesto a no autorizado"
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   90
      TabIndex        =   76
      Top             =   0
      Width           =   13160
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Estado del presupuesto"
         Top             =   230
         Width           =   2080
      End
      Begin VB.TextBox txtDiasVencimiento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7845
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Número de días de vigencia del presupuesto a partir de la fecha del mismo"
         Top             =   230
         Width           =   1200
      End
      Begin VB.TextBox txtClave 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   880
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Número del presupuesto"
         Top             =   230
         Width           =   1200
      End
      Begin MSMask.MaskEdBox mskFechaPresupuesto 
         Height          =   315
         Left            =   4695
         TabIndex        =   1
         ToolTipText     =   "Fecha del presupuesto"
         Top             =   210
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblEstado 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
         Height          =   195
         Left            =   10250
         TabIndex        =   85
         Top             =   240
         Width           =   495
      End
      Begin VB.Label DiasVencimiento 
         AutoSize        =   -1  'True
         Caption         =   "Días de vencimiento"
         Height          =   195
         Left            =   6285
         TabIndex        =   79
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label lblFechaPresupuesto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha del presupuesto"
         Height          =   195
         Left            =   2940
         TabIndex        =   78
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   140
         TabIndex        =   77
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame fraAccionPresupuesto 
      Height          =   855
      Left            =   3770
      TabIndex        =   71
      Top             =   7965
      Width           =   5630
      Begin VB.CommandButton cmdNoAutorizar 
         Caption         =   "Cambiar a no autorizado"
         Height          =   540
         Left            =   4460
         TabIndex        =   83
         ToolTipText     =   "Cambiar el estado del presupuesto a ""No autorizado"""
         Top             =   200
         Width           =   1090
      End
      Begin VB.CommandButton cmdVerPagos 
         Caption         =   "Ver pagos"
         Height          =   540
         Left            =   3475
         TabIndex        =   82
         ToolTipText     =   "Ver los pagos y/o devoluciones de dinero del paciente"
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton cmdCrearPaquete 
         Caption         =   "Generar paquete"
         Height          =   540
         Left            =   2500
         TabIndex        =   81
         ToolTipText     =   "Generar paquete "
         Top             =   200
         Width           =   960
      End
      Begin VB.CommandButton cmdDuplicar 
         Caption         =   "Duplicar presupuesto"
         Height          =   540
         Left            =   1455
         TabIndex        =   80
         ToolTipText     =   "Duplicar presupuesto"
         Top             =   200
         Width           =   1030
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame3"
         Height          =   645
         Left            =   1340
         TabIndex        =   75
         Top             =   100
         Width           =   60
      End
      Begin VB.OptionButton optCarta 
         Caption         =   "Carta"
         Height          =   255
         Left            =   550
         TabIndex        =   74
         ToolTipText     =   "Selccione carta si desea el presupuesto extendido."
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optTicket 
         Caption         =   "Ticket"
         Height          =   255
         Left            =   550
         TabIndex        =   73
         ToolTipText     =   "Seleccione ticket si desea imprimirlo de manera comprimida"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   480
         Left            =   50
         MaskColor       =   &H80000014&
         Picture         =   "frmPresupuestos.frx":0022
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Permite imprimir el presupuesto"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame freParametros 
      Height          =   4110
      Left            =   2308
      TabIndex        =   40
      Top             =   10005
      Visible         =   0   'False
      Width           =   8805
      Begin VB.Frame Frame9 
         Caption         =   "Mensaje del presupuesto"
         Height          =   1230
         Left            =   225
         TabIndex        =   42
         Top             =   585
         Width           =   8355
         Begin VB.TextBox txtMensaje 
            Height          =   700
            Left            =   285
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   43
            ToolTipText     =   "Este texto aparecerá en la parte superior del presupuesto (sólo en la forma extendida)."
            Top             =   315
            Width           =   7845
         End
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   450
         Left            =   3720
         TabIndex        =   46
         ToolTipText     =   "Guardar los cambios."
         Top             =   3555
         Width           =   1530
      End
      Begin VB.Frame Frame10 
         Caption         =   "Notas aclaratorias"
         Height          =   1230
         Left            =   225
         TabIndex        =   44
         Top             =   2145
         Width           =   8355
         Begin VB.TextBox txtNotas 
            Height          =   700
            Left            =   285
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   45
            ToolTipText     =   "Este texto aparecerá en la parte inferior del presupuesto."
            Top             =   330
            Width           =   7845
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000002&
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   45
         Top             =   120
         Width           =   8715
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Parámetros de los presupuestos"
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
         Height          =   300
         Left            =   90
         TabIndex        =   41
         Top             =   180
         Width           =   4200
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1335
      Left            =   2608
      TabIndex        =   35
      Top             =   9120
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   360
         Left            =   165
         TabIndex        =   36
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
         TabIndex        =   37
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
   Begin VB.Frame freBotones 
      Height          =   855
      Left            =   90
      TabIndex        =   34
      Top             =   7965
      Width           =   3615
      Begin VB.CommandButton cmdDelete 
         Height          =   480
         Left            =   3050
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Borrar presupuesto"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   480
         Left            =   2550
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0836
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Guardar el registro"
         Top             =   220
         Width           =   495
      End
      Begin VB.CommandButton cmdUltimoRegistro 
         Height          =   480
         Left            =   2050
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Último registro"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguienteRegistro 
         Height          =   480
         Left            =   1550
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0CEA
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Siguiente registro"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   480
         Left            =   1050
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0E5C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Búsqueda"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAnteriorRegistro 
         Height          =   480
         Left            =   550
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":0FCE
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Anterior registro"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrimerRegistro 
         Height          =   480
         Left            =   50
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPresupuestos.frx":1140
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Primer registro"
         Top             =   220
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame FreTotales 
      Enabled         =   0   'False
      Height          =   2090
      Left            =   9470
      TabIndex        =   25
      Top             =   6735
      Width           =   3780
      Begin VB.TextBox txtImporte 
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
         Height          =   340
         Left            =   1440
         TabIndex        =   112
         ToolTipText     =   "Subtotal del presupuesto"
         Top             =   195
         Width           =   2235
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   1440
         TabIndex        =   33
         ToolTipText     =   "Total del presupuesto"
         Top             =   1620
         Width           =   2235
      End
      Begin VB.TextBox txtIva 
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
         Left            =   1440
         TabIndex        =   32
         ToolTipText     =   "Iva del presupuesto"
         Top             =   1260
         Width           =   2235
      End
      Begin VB.TextBox txtDescuentos 
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
         Left            =   1440
         TabIndex        =   31
         ToolTipText     =   "Total de descuento"
         Top             =   540
         Width           =   2235
      End
      Begin VB.TextBox txtSubtotal 
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
         Left            =   1440
         TabIndex        =   30
         ToolTipText     =   "Subtotal del presupuesto"
         Top             =   900
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   113
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   1650
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Top             =   1290
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   26
         Top             =   215
         Width           =   1215
      End
   End
   Begin VB.Frame FreElementos 
      Caption         =   "Elementos disponibles"
      Height          =   2455
      Left            =   6235
      TabIndex        =   20
      Top             =   680
      Width           =   7015
      Begin VB.OptionButton optCodigoBarras 
         Caption         =   "Código barras"
         Height          =   225
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   1290
      End
      Begin VB.OptionButton optDescripcion 
         Caption         =   "Descripción"
         Height          =   225
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optClave 
         Caption         =   "Clave"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtSeleArticulo 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Teclee la clave o la descripción del cargo"
         Top             =   540
         Width           =   6715
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdElementos 
         Height          =   1455
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Seleccione el cargo de esta lista"
         Top             =   900
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   2566
         _Version        =   393216
         GridColor       =   12632256
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame fraElementoFijo 
      Caption         =   "Elementos fijos"
      Height          =   1275
      Left            =   6235
      TabIndex        =   63
      Top             =   3105
      Width           =   7015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdElementoFijo 
         Height          =   930
         Left            =   120
         TabIndex        =   64
         Top             =   225
         Width           =   6715
         _ExtentX        =   11853
         _ExtentY        =   1640
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame freDatos 
      Caption         =   "Datos del paciente"
      Height          =   2455
      Left            =   90
      TabIndex        =   18
      Top             =   680
      Width           =   6075
      Begin VB.TextBox txtNumCuenta 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   102
         ToolTipText     =   "Número de cuenta"
         Top             =   180
         Width           =   1230
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Urgencias"
         Height          =   195
         Index           =   2
         Left            =   2080
         TabIndex        =   6
         Top             =   225
         Width           =   1020
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externo"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   835
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Interno"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   225
         Width           =   800
      End
      Begin VB.Frame freDatos2 
         BorderStyle     =   0  'None
         Height          =   1905
         Left            =   90
         TabIndex        =   53
         Top             =   480
         Width           =   5880
         Begin VB.ComboBox cboTipoConvenio 
            Height          =   315
            Index           =   0
            Left            =   1470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Seleccione el tipo de convenio"
            Top             =   1260
            Width           =   4350
         End
         Begin VB.ComboBox cboProcedimiento 
            Height          =   315
            ItemData        =   "frmPresupuestos.frx":12B2
            Left            =   1470
            List            =   "frmPresupuestos.frx":12B4
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Seleccione el procedimiento quirúrgico"
            Top             =   600
            Width           =   4350
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Index           =   0
            Left            =   1470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            ToolTipText     =   "Seleccione la empresa de donde viene el paciente"
            Top             =   1575
            Width           =   4350
         End
         Begin VB.ComboBox cboTipoPaciente 
            Height          =   315
            Index           =   0
            Left            =   1470
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Seleccione el tipo de paciente"
            Top             =   945
            Width           =   4350
         End
         Begin VB.TextBox txtNombre 
            Height          =   285
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   7
            ToolTipText     =   "Nombre del paciente con que se imprimirá la cotización"
            Top             =   0
            Width           =   4350
         End
         Begin VB.TextBox txtDireccion 
            Height          =   285
            Left            =   1470
            MaxLength       =   100
            TabIndex        =   8
            ToolTipText     =   "Dirección del paciente (opcional)"
            Top             =   300
            Width           =   4350
         End
         Begin VB.Label lblTipoConvenio 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de convenio"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   118
            Top             =   1320
            Width           =   1245
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Procedimiento quirúrgico"
            Height          =   435
            Left            =   45
            TabIndex        =   70
            Top             =   570
            Width           =   1260
            WordWrap        =   -1  'True
         End
         Begin VB.Label lstEmpresa 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   69
            Top             =   1575
            Width           =   615
         End
         Begin VB.Label lblTipoPaciente 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Index           =   0
            Left            =   50
            TabIndex        =   56
            Top             =   1055
            Width           =   1200
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   50
            TabIndex        =   55
            Top             =   15
            Width           =   555
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Left            =   50
            TabIndex        =   54
            Top             =   285
            Width           =   675
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número cuenta"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3460
         TabIndex        =   103
         Top             =   220
         Width           =   1095
      End
   End
   Begin VB.Frame FreIncluir 
      Height          =   1275
      Left            =   90
      TabIndex        =   21
      Top             =   3105
      Width           =   6075
      Begin VB.CommandButton cmdParametros 
         Caption         =   "Parámetros"
         Height          =   480
         Left            =   300
         TabIndex        =   86
         ToolTipText     =   "Permite configurar los mensajes mostrados en el presupuesto"
         Top             =   435
         Width           =   940
      End
      Begin VB.CommandButton cmdAplicaDescuento 
         Caption         =   "Aplicar"
         Height          =   345
         Index           =   0
         Left            =   3720
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txtPorcentajeDescuento 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   66
         Top             =   850
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Frame Frame7 
         Height          =   1050
         Left            =   1350
         TabIndex        =   39
         Top             =   120
         Width           =   75
      End
      Begin VB.CheckBox chkTomarDescuentos 
         Caption         =   "Tomar en cuenta descuentos actuales asignados por tipo de paciente"
         Height          =   570
         Index           =   0
         Left            =   1560
         TabIndex        =   38
         ToolTipText     =   "Si se toma en cuenta o no los descuentos."
         Top             =   240
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Incluir"
         Height          =   615
         Index           =   0
         Left            =   4800
         MaskColor       =   &H80000014&
         Picture         =   "frmPresupuestos.frx":12B6
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Incluir un cargo al presupuesto"
         Top             =   390
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Excluir"
         Height          =   615
         Index           =   1
         Left            =   5400
         MaskColor       =   &H80000014&
         Picture         =   "frmPresupuestos.frx":1410
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Excluir un cargo al presupuesto"
         Top             =   390
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.Label lblPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Index           =   0
         Left            =   3360
         TabIndex        =   67
         Top             =   910
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblDescuento 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   65
         Top             =   910
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.Frame FreDetalle 
      BorderStyle     =   0  'None
      Height          =   1900
      Left            =   90
      TabIndex        =   24
      Top             =   4395
      Width           =   13170
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   0
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   115
         Top             =   1095
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPresupuesto 
         Height          =   1800
         Index           =   0
         Left            =   0
         TabIndex        =   116
         ToolTipText     =   "Cargos que incluye el presupuesto"
         Top             =   60
         Width           =   13155
         _ExtentX        =   23204
         _ExtentY        =   3175
         _Version        =   393216
         Cols            =   8
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         FormatString    =   "|Descripción|Precio|Cantidad|Subtotal|Descuento|Monto|Tipo"
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   5
         _Band(0).Cols   =   8
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame freVarios 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   120
      TabIndex        =   120
      Top             =   6340
      Width           =   13170
      Begin VB.CommandButton cmdActualizaPrecios 
         Caption         =   "Actualizar precios y descuentos"
         Height          =   375
         Index           =   0
         Left            =   10680
         TabIndex        =   123
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtMargenUtilidadTotal 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   5535
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   122
         ToolTipText     =   "Margen de utilidad total del presupuesto"
         Top             =   45
         Width           =   1200
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
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
         Index           =   2
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   121
         Top             =   1095
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblUtilidadTotal 
         AutoSize        =   -1  'True
         Caption         =   "Margen de utilidad total"
         Height          =   195
         Left            =   3720
         TabIndex        =   124
         Top             =   75
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmPresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmPresupuestos
'-------------------------------------------------------------------------------------
'| Objetivo: Realizar los presupuesto para futuros clientes
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 13/Enero/2001
'| Ultimas modificaciones, especificar:
'-------------------------------------------------------------------------------------

Option Explicit
Const vgintColumnaCurrency = -1 'Para la columna que se va a editar
Const intTotalColsgrdPresupuesto = 16
Const strFormatgrdPresupuesto = "|Descripción|Costo|Margen utilidad|Precio|Cantidad|Importe|Descuento|Subtotal|IVA|Tipo|"
Const ColDescripciongrdPresupuesto = 1      'Descripción
Const ColCostogrdPresupuesto = 2            'Costo
Const ColMargenUtilidadgrdPresupuesto = 3   'Margen utilidad
Const ColPreciogrdPresupuesto = 4           'Precio
Const ColCantidadgrdPresupuesto = 5         'Cantidad
Const ColSubtotalgrdPresupuesto = 6         'Subtotal
Const ColDescuentogrdPresupuesto = 7        'Descuento en cantidad
Const ColMontogrdPresupuesto = 8            'Monto
Const ColMontoIVAgrdPresupuesto = 9         'Monto IVA
Const ColTipogrdPresupuesto = 10            'Tipo
Const ColIVAgrdPresupuesto = 11             'Porcentaje IVA
Const ColPorcentajeDescgrdPresupuesto = 12  'Porcentaje de descuento
Const ColContenidogrdPresupuesto = 13       'Contenido
Const ColModoDescuentogrdPresupuesto = 14   'Modo de Descuento Inventario
Const ColRowDatagrdPresupuesto = 15         'RowData

Private vgrptReporte As CRAXDRT.Report

Dim rsPresupuesto As New ADODB.Recordset
Dim rsPvSelDescuento As New ADODB.Recordset

Dim vgbitParametros As Boolean
Dim vgblnNoEditarPagos As Boolean
Dim vgblnEditaPago As Boolean 'Para saber si se esta editando una cantidad
Dim vlblnCargando As Boolean
Dim vgblnEsNuevo As Boolean 'Para saber si es un presupuesto nuevo

Dim vgstrEstadoManto As String 'Vacio = Sin movimientos, "C" = Consulta o creando nuevo presupuesto, "BC" = Ventana de busqueda abierta, "E" = Editando valores en grid
Dim vlaryResultados() As String

Dim vglngClaveSiguiente As Long
Dim vlTipoPacienteSocio As Long

Dim vldblImporte As Double
Dim vldblSubtotal As Double
Dim vldblDescuento As Double
Dim vldblIVA As Double
Dim vldbltotal As Double

Dim vgintTipoPaciente As Integer 'Guarda el index el optionButton

Dim rsPaquetes As New ADODB.Recordset
Dim vllngCvePaquete As Long
Dim vllngCtaPaciente As Long
Dim vllngPersonaGraba As Long
Dim vlintCvePresupuesto As Integer
Dim vlstrTipoTipoPaciente As String                     'Clasificación del tipo de paciente seleccionado
                                                        '"PA" = particulares
                                                        '"CO" = convenios
                                                        '"EM" = empleados
                                                        '"ME" = médicos

Dim vgdblPrecioPredGrupo As Double
Dim vgstrTipoPredGrupo As String
Dim vglngClavePredGrupo As Long

Dim vgdblPrecioConvGrupo As Double
Dim vgstrTipoConvGrupo As String
Dim vglngClaveConvGrupo As Long
Dim vglngTipoParticular As Long

Dim vldblCostoTotal As Double
Dim vldblMargen As Double
Dim vldblmargenutilidad As Double
Dim vgstrTipoPaciente As String, vgstrTipoConvenio As String, vgstrEmpresa As String

Private Sub pElementoGrupoConv(vlIntClave As Long)
' Procedimiento que carga las variables con los datos del elemento mas caro del grupo en base a la lista del convenio
    Dim vlstrSentencia As String
    Dim vlintcontador As Integer
    Dim vllngEstaEnPaquete As Integer
    Dim vldblPrecio As Double
    Dim rsTemp As ADODB.Recordset

    vgdblPrecioConvGrupo = 0
    vgstrTipoConvGrupo = "GC"
    vglngClaveConvGrupo = 0

    vlstrSentencia = "SELECT chrTipoCargo Tipo, intCveCargo Clave FROM PVDETALLEGRUPOCARGO WHERE intCveGrupo = " & vlIntClave
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        For vlintcontador = 1 To rsTemp.RecordCount
            vllngEstaEnPaquete = FintBuscaEnRowData(grdPresupuesto(0), rsTemp!clave, IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo))
            If vllngEstaEnPaquete = -1 Then
                vldblPrecio = 0
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                vgstrParametrosSP = rsTemp!clave & "|" & IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo) & "|" & IIf(cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) <> 0, cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex), CStr(vglngTipoParticular)) & "|" & cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) & "|" & IIf(cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) <> 0, IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")), "E") & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                pObtieneValores vlaryResultados, vldblPrecio
                If vldblPrecio > vgdblPrecioConvGrupo Then
                    vgdblPrecioConvGrupo = vldblPrecio
                    vgstrTipoConvGrupo = IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo)
                    vglngClaveConvGrupo = rsTemp!clave
                End If
            End If
            rsTemp.MoveNext
        Next vlintcontador
    End If
    
    rsTemp.Close
    
End Sub


Private Sub pElementoGrupoPred(vlIntClave As Long, Index As Integer)
' Procedimiento que carga las variables con los datos del elemento mas caro del grupo en base a las listas predeterminadas
    Dim vlstrSentencia As String
    Dim vlintcontador As Integer
    Dim vllngEstaEnPaquete As Integer
    Dim vldblPrecio As Double
    Dim rsTemp As ADODB.Recordset

    vgdblPrecioPredGrupo = 0
    vgstrTipoPredGrupo = "GC"
    vglngClavePredGrupo = 0

    vlstrSentencia = "SELECT chrTipoCargo Tipo, intCveCargo Clave FROM PVDETALLEGRUPOCARGO WHERE intCveGrupo = " & vlIntClave
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        For vlintcontador = 1 To rsTemp.RecordCount
            vllngEstaEnPaquete = FintBuscaEnRowData(grdPresupuesto(Index), rsTemp!clave, IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo))
            If vllngEstaEnPaquete = -1 Then
                vldblPrecio = 0
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                vgstrParametrosSP = rsTemp!clave & "|" & IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo) & "|" & CStr(vglngTipoParticular) & "|0|E|0|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                pObtieneValores vlaryResultados, vldblPrecio
                If vldblPrecio > vgdblPrecioPredGrupo Then
                    vgdblPrecioPredGrupo = vldblPrecio
                    vgstrTipoPredGrupo = IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo)
                    vglngClavePredGrupo = rsTemp!clave
                End If
            End If
            rsTemp.MoveNext
        Next vlintcontador
    End If
    
    rsTemp.Close
    
End Sub

Private Function FintBuscaEnRowData(grdHBusca As MSHFlexGrid, vllngCriterio As Long, vlstrTipoElemento)
    On Error GoTo NotificaError

    Dim vlintcontador As Long
    FintBuscaEnRowData = -1
    With grdHBusca
    For vlintcontador = 1 To .Rows - 1
        If CLng(IIf(.TextMatrix(vlintcontador, ColRowDatagrdPresupuesto) = "", "-1", .TextMatrix(vlintcontador, ColRowDatagrdPresupuesto))) > -1 Then
            If CLng(.TextMatrix(vlintcontador, ColRowDatagrdPresupuesto)) = vllngCriterio And vlstrTipoElemento = .TextMatrix(vlintcontador, ColTipogrdPresupuesto) Then
                FintBuscaEnRowData = vlintcontador
                Exit For
            End If
        End If
    Next
    End With
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":FintBuscaEnRowData"))
End Function
Private Sub pHabilitaDatosPaquete(vlblnEstado As Boolean, vlblnConsulta As Boolean, Optional vlblnNuevo As Boolean = False)
    If vlblnEstado = False And vlblnConsulta = False Then
        cboConceptoFactura.ListIndex = -1
        cboTratamiento.ListIndex = -1
        lblNumPaquete.Enabled = vlblnEstado
        If Not vlblnNuevo Then
            pHabilitaBotones 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1
            vgstrEstadoManto = "C"
        End If
    End If
    cboConceptoFactura.Enabled = vlblnEstado
    cboTratamiento.Enabled = vlblnEstado
    fraDatosPaquete.Enabled = vlblnEstado
    lblConceptoFactura.Enabled = vlblnEstado
    lblTratamiento.Enabled = vlblnEstado
End Sub

Private Sub pHabilitaFramesUpdate(vlblnEstado As Boolean)
    On Error GoTo NotificaError
        FreElementos.Enabled = vlblnEstado
        fraElementoFijo.Enabled = vlblnEstado
        FreDetalle.Enabled = vlblnEstado
        FreIncluir.Enabled = vlblnEstado
        freDatos2.Enabled = vlblnEstado
        cmdActualizaPrecios(0).Enabled = vlblnEstado
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaFramesNoUpdate"))
End Sub

Private Sub pGrabaNoAutorizado()
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    vlstrSentencia = "UPDATE PVPRESUPUESTO SET CHRESTADO = 'N' WHERE INTCVEPRESUPUESTO = " & txtClave.Text
    pEjecutaSentencia vlstrSentencia

    Call pGuardarLogPresupuesto("N", vllngPersonaGraba, CLng(txtClave.Text), txtMotivoNoAutorizado.Text)
    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PRESUPUESTO (Cambio de estado a NO AUTORIZADO)", txtClave.Text)
    EntornoSIHO.ConeccionSIHO.CommitTrans

    'La operación se realizó satisfactoriamente.
    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"

    'fraMotivoNoAutorizado.Visible = False
    'freDatos.Enabled = True
    'freBotones.Enabled = True
    'Frame3.Enabled = True
    'pHabilitaFrames True
    'lblMotivoNoAutorizado.Visible = False
    txtMotivoNoAutorizado.Enabled = False
    txtClave_KeyDown vbKeyReturn, 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaNoAutorizado"))
End Sub


Private Sub pGuardarLogPresupuesto(vlstrEstado As String, vlintIDEmpleado As Long, vllngCvePresupuesto As Long, vlstrComentarios As String)
    On Error GoTo NotificaError
   
    Dim vlstrsql As String
            
    vlstrsql = "INSERT INTO PvPresupuestoLog (intCvePresupuesto, chrEstado, dtmFecha, chrComentarios, intCveEmpleado) " & _
               "VALUES (" & vllngCvePresupuesto & ", '" & vlstrEstado & "', " & fstrFechaSQL(fdtmServerFechaHora, fdtmServerHora) & ", '" & vlstrComentarios & "', " & vlintIDEmpleado & ")"
    EntornoSIHO.ConeccionSIHO.Execute vlstrsql
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " :modProcedimientos " & ":pGuardarLogPresupuesto"))
End Sub

Private Sub pAsignaPaquete(lngnumCuenta As Long, lngCvePaquete As Long, dblPrecio As Double, strTipoPaciente As String, Optional intbitpesos As Integer)
1         On Error GoTo NotificaError

          Dim strSentencia As String
          Dim dblIVA As Double
          Dim dblDescuento As Double
          Dim rsDescuento As New ADODB.Recordset
          Dim vllngCantPaquetesSinFacturar As Long
          Dim intBitValidaPaquetes As Long
          Dim strParametrosPaquetes As String

2         strSentencia = "UPDATE PvPaquetePaciente SET BITPAQUETEDEFAULT = 0 where chrTipoPaciente = '" & strTipoPaciente & "' and intMovPaciente = " & str(lngnumCuenta)
3         pEjecutaSentencia strSentencia
          
          'Obtener el IVA del Paquete
4         strSentencia = " Select isnull(pvConceptoFacturacion.SMYIVA,0) smyIVA " & _
              " From PvPaquete " & _
              " Inner Join PvConceptoFacturacion On (pvPaquete.SMICONCEPTOFACTURA = pvConceptoFacturacion.SMICVECONCEPTO)" & _
              " Where pvPaquete.INTNUMPAQUETE = " & Trim(str(lngCvePaquete))
              
5         dblIVA = frsRegresaRs(strSentencia, adLockOptimistic, adOpenForwardOnly)!smyIVA
          
          'Obtener descuento del paquete
6         dblDescuento = 0
          
7         strSentencia = "SELECT SP_pvseldescuentopaquete('" & strTipoPaciente & "', " & str(lngnumCuenta) & ", " & str(lngCvePaquete) & _
                                              ", " & str(dblPrecio) & ", " & vgintNumeroDepartamento & ", '" & fdtmServerFecha & "') As Descuento " & _
                           "FROM DUAL"
              
8         Set rsDescuento = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
9         If rsDescuento.RecordCount > 0 Then
10            dblDescuento = rsDescuento!Descuento
11        End If
         
12        vgstrParametrosSP = str(lngnumCuenta) & "|" & strTipoPaciente & "|" & str(lngCvePaquete) & "|" & str(dblPrecio) & "|" & "1" _
          & "|" & (dblPrecio - dblDescuento) * (dblIVA / 100) & "|" & str(vllngPersonaGraba) & "|" & str(dblDescuento) & "|1" & "|" & Trim(str(intbitpesos))
          
13        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPAQUETEPACIENTE"
              
          '' SECCION AGREGADA PARA QUE SE ACTUALICEN LOS CARGOS CON LA ASIGNACIÓN DEL PAQUETE
              ' Valida que el paquete no esté facturado
14            vllngCantPaquetesSinFacturar = 1
15            strParametrosPaquetes = str(lngnumCuenta) & "|" & strTipoPaciente & "|" & str(lngCvePaquete)
16            frsEjecuta_SP strParametrosPaquetes, "FN_PVSELPAQUETESINFACTURAR", True, vllngCantPaquetesSinFacturar
              
              ' Regresa bit para validar paquetes
17            intBitValidaPaquetes = 1
18            frsEjecuta_SP str(lngCvePaquete), "FN_PVSELVALIDACARGOSPAQUETE", True, intBitValidaPaquetes
              
19            vgstrParametrosSP = str(lngnumCuenta) & "|" & strTipoPaciente & "|" & str(lngCvePaquete) & "|" & vllngCantPaquetesSinFacturar & "|" & intBitValidaPaquetes & "|" & -1
20            frsEjecuta_SP vgstrParametrosSP, "sp_pvupdcargospaquete"
          '' SECCION AGREGADA PARA QUE SE ACTUALICEN LOS CARGOS CON LA ASIGNACIÓN DEL PAQUETE
          
    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsignaPaquete" & " Linea:" & Erl()))
        Unload Me
End Sub


Private Sub pVerPagos(vlstrDestino As String)
    On Error GoTo NotificaError

    Dim vllngCuenta As Long
    Dim alstrParametros(11) As String
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    Dim vlstrTipoIngreso  As String
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlstrFechaInicio As String
    Dim vlstrFechaFin As String
    
    vllngCuenta = 0

    If Val(txtNumCuenta.Text) <> 0 Then
        vllngCuenta = CLng(txtNumCuenta.Text)
    End If
    
    vlstrTipoIngreso = 0
    vlstrSentencia = "select pi.intcvetipoingreso, pi.intcvetipopaciente, pi.chrTipoIngreso from ExPacienteIngreso pi where pi.intNumCuenta = " & txtNumCuenta.Text
    Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
    If rs.RecordCount > 0 Then vlstrTipoIngreso = rs!chrTipoIngreso
    
    vlstrFechaInicio = Format((fdtmServerFecha - 60), "dd/mm/yyyy")
    vlstrFechaFin = Format(fdtmServerFecha, "dd/mm/yyyy")
    
    vlstrx = fstrFechaSQL(vlstrFechaInicio)
    vlstrx = vlstrx & "|" & fstrFechaSQL(vlstrFechaFin)
    vlstrx = vlstrx & "|" & vllngCuenta
    vlstrx = vlstrx & "|" & vlstrTipoIngreso
    vlstrx = vlstrx & "|" & vgintNumeroDepartamento
    vlstrx = vlstrx & "|" & vgintClaveEmpresaContable
    vlstrx = vlstrx & "|" & "0"

    Set rsReporte = frsEjecuta_SP(vlstrx, "Sp_PVAnticipoPaciente")
    If rsReporte.RecordCount > 0 Then

      pInstanciaReporte vgrptReporte, "rptPVAnticipoPaciente.rpt"
      vgrptReporte.DiscardSavedData

      alstrParametros(0) = "p_empresa;" & Trim(vgstrNombreHospitalCH)
      alstrParametros(4) = "p_finicio;" & UCase(Format(vlstrFechaInicio, "dd/mmm/yyyy"))   'fstrFechaSQL(txtFechaInicio.Text, "")
      alstrParametros(5) = "p_ffin;" & UCase(Format(vlstrFechaFin, "dd/mmm/yyyy"))    'fstrFechaSQL(txtFechaFin.Text, "")
      If optTipoPaciente(0).Value Then
          alstrParametros(6) = "p_tipopaciente;" & "Internos"
      Else
          If optTipoPaciente(1).Value Then
              alstrParametros(6) = "p_tipopaciente;" & "Externos"
          Else
              alstrParametros(6) = "p_tipopaciente;" & "<TODOS>"
          End If
      End If
      alstrParametros(7) = "p_paciente;" & Trim("")
      alstrParametros(11) = "p_tiporpt;" & "ENTRADAS Y SALIDAS DE DINERO"

      pCargaParameterFields alstrParametros, vgrptReporte
      pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Entradas y salidas de dinero"

    Else
      'No existe información con esos parámetros.
      MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If

    If rsReporte.State <> adStateClosed Then rsReporte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pVerPagos"))
End Sub


Private Function flngGuardaPaciente(lngCvePaquete As Long) As Long
On Error GoTo NotificaError
    Dim vlstrTipoIngreso  As String
    Dim vlintBitPesos As Integer
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsPaquete  As New ADODB.Recordset
    Dim vldblPrecio As Double
    
    flngGuardaPaciente = 0
    
    frmAdmisionPaciente.vlblnMostrarTabGenerales = True
    frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
    frmAdmisionPaciente.vlblnMostrarTabInternos = False
    frmAdmisionPaciente.vlblnMostrarTabPrepagos = True
    frmAdmisionPaciente.vlblnMostrarTabIngresosPrevios = False
    frmAdmisionPaciente.vlblnMostrarTabEgresados = False
    frmAdmisionPaciente.vlblnMostrarTabExternos = False
    frmAdmisionPaciente.vlintPestañaInicial = 0
    
    frmAdmisionPaciente.blnAbrirCuenta = False
    frmAdmisionPaciente.blnActivar = False
    frmAdmisionPaciente.blnHabilitarAbrirCuenta = False
    frmAdmisionPaciente.blnHabilitarActivar = False
    frmAdmisionPaciente.blnHabilitarReporte = False
    frmAdmisionPaciente.blnConsulta = False
    frmAdmisionPaciente.vglngExpedienteConsulta = 0
    frmAdmisionPaciente.blnHonorariosCC = False
    frmAdmisionPaciente.vgstrForma = "frmPresupuestos"
    frmAdmisionPaciente.vglngCvePaquete = lngCvePaquete
    frmAdmisionPaciente.vglngTipoPaciente = CLng(Val(cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex)))
    If cboTipoConvenio(0).ListIndex > -1 Then
        frmAdmisionPaciente.vglngTipoConvenio = CLng(Val(cboTipoConvenio(0).ItemData(cboTipoConvenio(0).ListIndex)))
    End If
    If cboEmpresa(0).ListIndex > -1 Then
        frmAdmisionPaciente.vglngEmpresaPaciente = CLng(Val(cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex)))
    End If
    frmAdmisionPaciente.vllngNumeroOpcionExterno = 352
    frmAdmisionPaciente.Show vbModal, Me
    
    vglngNumeroPaciente = frmAdmisionPaciente.vglngExpediente
    vglngNumeroCuenta = frmAdmisionPaciente.vglngCuenta
    
    If vglngNumeroCuenta > 0 Then
    
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        vlstrSentencia = "Update pvPaquetePaciente set mnyDescuento = " & Val(Format(txtDescuentos.Text, "##########0.00####")) & ", mnyIvaPaquete = " & Val(Format(txtIva.Text, "##########0.00####")) & " where intMovPaciente = " & vglngNumeroCuenta & " and intNumPaquete = " & lngCvePaquete
        pEjecutaSentencia vlstrSentencia
        
        vlstrSentencia = "Update pvPresupuesto set chrEstado = 'A', intNumPaquete = " & lngCvePaquete & " where intCvePresupuesto = " & Trim(txtClave.Text)
        pEjecutaSentencia vlstrSentencia

        Call pGuardarLogPresupuesto("A", vllngPersonaGraba, CLng(txtClave.Text), "Presupuesto autorizado, creación del paquete " & lngCvePaquete & " del paciente " & vglngNumeroCuenta)
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        flngGuardaPaciente = vglngNumeroCuenta
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngGuardaPaciente"))
End Function


Private Function flngGuardaPaquete() As Long
On Error GoTo NotificaError

    Dim rsDetallePaquete As New ADODB.Recordset
    Dim rsDescuentoInventario As New ADODB.Recordset
    Dim vlintcontador As Integer
    Dim vlstrSentenciaDescuentoInventario As String
    Dim vlstrSentencia As String
    Dim vlintRow As Integer

    Dim rsReporte As New ADODB.Recordset
    Dim rsCostoPaquete As New ADODB.Recordset
    Dim vlintsubtotal As Double
    Dim vlIntCont As Integer
    Dim vlblnDatosValidos As Boolean
    Dim lngCvePaquete As Long
    Dim vldblTotalCostoBasePaquete As Double
    Dim rs As New ADODB.Recordset
    Dim rsListaPrecios As New ADODB.Recordset

    flngGuardaPaquete = 0
    vlblnDatosValidos = True
    
    pHabilitaDatosPaquete True, False
        
    For vlintcontador = 1 To grdPresupuesto(0).Rows - 1
        If grdPresupuesto(0).TextMatrix(vlintcontador, ColTipogrdPresupuesto) = "EF" Then
            MsgBox "No es posible generar el paquete debido a que el presupuesto contiene elementos fijos.", vbOKOnly + vbInformation, "Mensaje"
            grdPresupuesto(0).SetFocus
            cboConceptoFactura.ListIndex = -1
            cboTratamiento.ListIndex = -1
            pHabilitaDatosPaquete False, False
            Exit Function
        End If
        If grdPresupuesto(0).TextMatrix(vlintcontador, ColTipogrdPresupuesto) = "PA" Then
            MsgBox "No es posible generar el paquete debido a que uno de los elementos del presupuesto es un paquete.", vbOKOnly + vbInformation, "Mensaje"
            grdPresupuesto(0).SetFocus
            cboConceptoFactura.ListIndex = -1
            cboTratamiento.ListIndex = -1
            pHabilitaDatosPaquete False, False
            Exit Function
        End If
    Next
    vlstrSentencia = "SELECT * From pvconceptofacturacion " & _
                     "INNER JOIN pvconceptofacturacionempresa ON pvconceptofacturacion.smicveconcepto = pvconceptofacturacionempresa.intcveconceptofactura " & _
                     "WHERE pvconceptofacturacionempresa.intCveDepartamento = " & vgintNumeroDepartamento & _
                     " AND pvconceptofacturacion.bitpaquetepresupuesto = 1"
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        cboConceptoFactura.ListIndex = fintLocalizaCbo(cboConceptoFactura, rs!smicveconcepto)
    End If
    
    If vlblnDatosValidos And cboConceptoFactura.ListIndex = -1 Then
        vlblnDatosValidos = False
        'Seleccione el dato.
        'MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
        cboConceptoFactura.SetFocus
        Exit Function
    End If
    If vlblnDatosValidos And cboTratamiento.ListIndex = -1 Then
        vlblnDatosValidos = False
        'Seleccione el dato.
        'MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
        cboTratamiento.SetFocus
        Exit Function
    End If

    If vlblnDatosValidos Then
            vlstrSentencia = "SELECT * FROM PVLISTAPRECIO " & _
                             "WHERE PVLISTAPRECIO.bitestatusactivo = 1 " & _
                             "AND PVLISTAPRECIO.smidepartamento = (SELECT intcvedepartamento " & _
                                                                  "FROM PVCONCEPTOFACTURACIONEMPRESA " & _
                                                                  "WHERE INTCVECONCEPTOFACTURA = " & cboConceptoFactura.ItemData(cboConceptoFactura.ListIndex) & _
                                                                  " AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable & ")"
                                                                  
        Set rsListaPrecios = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsListaPrecios.RecordCount = 0 Then
            vlblnDatosValidos = False
            MsgBox "No se encontró configurada alguna lista de precios para el concepto de facturación", vbOKOnly + vbInformation, "Mensaje"
            Exit Function
        End If
    End If
    
    vlstrSentencia = "select * from PvPaquete"
    Set rsPaquetes = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)

    With rsPaquetes
        '--------------------------------------------------------
        ' Persona que graba
        '--------------------------------------------------------
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Function

        '--------------------------------------
        EntornoSIHO.ConeccionSIHO.BeginTrans
        '--------------------------------------
        ' Grabar el Concepto de Facturación
        '--------------------------------------
        .AddNew
        '!chrDescripcion = txtClave.Text & " PAQUETE DE PRESUPUESTO"
        !chrDescripcion = "PRESUPUESTO NUM. " & txtClave.Text
        !bitactivo = 1
        !bitValidaCargosPaquete = 1
        !SMICONCEPTOFACTURA = cboConceptoFactura.ItemData(cboConceptoFactura.ListIndex)
        !chrTratamiento = Trim(cboTratamiento.List(cboTratamiento.ListIndex))
        !chrTipo = "PAQUETE"
        !mnyAnticipoSugerido = Val(Format("0", "###############.##"))
        !dtmFechaActualizacion = mskFechaPresupuesto
        !bitcostobase = 0   'Costa mas alto
        !chrTipoIngresoDescuento = "T"  'Todos
        !bitincrementoautomatico = 0
        !intOrigen = 1
        .Update

        lngCvePaquete = flngObtieneIdentity("SEC_PVPAQUETE", rsPaquetes!intnumpaquete)
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "PAQUETE", CStr(lngCvePaquete))

        'Call PBorraCargosAsignados(txtClave.Text)
        vlstrSentencia = "select * from PvDetallePaquete where intNumPaquete = -5" ' Paque no regrese nada
        Set rsDetallePaquete = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        If IIf(grdPresupuesto(0).TextMatrix(1, ColRowDatagrdPresupuesto) = "", -1, grdPresupuesto(0).TextMatrix(1, ColRowDatagrdPresupuesto)) <> -1 Then
            For vlintcontador = 1 To grdPresupuesto(0).Rows - 1
                rsDetallePaquete.AddNew
                rsDetallePaquete!intnumpaquete = IIf(CLng(lngCvePaquete) < 0, CLng(lngCvePaquete) * -1, CLng(lngCvePaquete))
                rsDetallePaquete!intCveCargo = grdPresupuesto(0).TextMatrix(vlintcontador, ColRowDatagrdPresupuesto)
                rsDetallePaquete!chrTipoCargo = Trim(grdPresupuesto(0).TextMatrix(vlintcontador, ColTipogrdPresupuesto))
                rsDetallePaquete!SMICANTIDAD = grdPresupuesto(0).TextMatrix(vlintcontador, ColCantidadgrdPresupuesto)
                rsDetallePaquete!INTDESCUENTOINVENTARIO = CInt(Val(grdPresupuesto(0).TextMatrix(vlintcontador, ColModoDescuentogrdPresupuesto)))
                rsDetallePaquete!mnyMontoLimite = Val(Format("0", "############.##"))
                rsDetallePaquete!mnycosto = Val(Format(grdPresupuesto(0).TextMatrix(vlintcontador, ColCostogrdPresupuesto), "############.##"))
                rsDetallePaquete!mnyPrecio = Val(Format(grdPresupuesto(0).TextMatrix(vlintcontador, ColPreciogrdPresupuesto), "############.##"))
                rsDetallePaquete!MNYDESCUENTO = Val(Format(grdPresupuesto(0).TextMatrix(vlintcontador, ColDescuentogrdPresupuesto), "############.##"))
                rsDetallePaquete!MNYIVA = Val(Format(grdPresupuesto(0).TextMatrix(vlintcontador, ColMontoIVAgrdPresupuesto), "############.##"))
                rsDetallePaquete!smicveconcepto = IIf(grdPresupuesto(0).TextMatrix(vlintcontador, ColTipogrdPresupuesto) = "GC", flngBuscaConcepto(grdPresupuesto(0).TextMatrix(vlintcontador, ColRowDatagrdPresupuesto)), 0)
                rsDetallePaquete!MNYPRECIOESPECIFICO = 0
                rsDetallePaquete.Update
            Next
        End If
        rsDetallePaquete.Close
        'End If

        vldblTotalCostoBasePaquete = 0
        For vlintcontador = 1 To grdPresupuesto(0).Rows - 1
            vldblTotalCostoBasePaquete = vldblTotalCostoBasePaquete + Val(Format(grdPresupuesto(0).TextMatrix(vlintcontador, ColCostogrdPresupuesto), ""))
        Next vlintcontador
        
        pEjecutaSentencia ("INSERT INTO PVCOSTOCARGOS VALUES(" & vgintClaveEmpresaContable & "," & IIf(CLng(lngCvePaquete) < 0, CLng(lngCvePaquete) * -1, CLng(lngCvePaquete)) & ",'PA'," & vldblTotalCostoBasePaquete & ")")

        'Se alimentan el departamento para el paquete con el departamento que lo crea
        vlstrSentencia = "INSERT INTO PVPAQUETEDEPARTAMENTO (INTNUMPAQUETE, SMICVEDEPARTAMENTO) VALUES(" & IIf(CLng(lngCvePaquete) < 0, CLng(lngCvePaquete) * -1, CLng(lngCvePaquete)) & ", " & vgintNumeroDepartamento & ")"
        pEjecutaSentencia vlstrSentencia
        
        'Se alimenta del detalle de las listas de precios del concepto de facturación
        If rsListaPrecios.RecordCount > 0 Then
            rsListaPrecios.MoveFirst
            For vlintcontador = 1 To rsListaPrecios.RecordCount
                vgstrParametrosSP = CStr(rsListaPrecios!intcvelista) & "|" & lngCvePaquete & "|" & "PA" & "|" & CStr(Val(Format(txtImporte.Text, "##########0.00####"))) & "|" & "C" & "|" & "0.0000" & "|" & "0" & "|" & "0" & "|" & IIf("PESOS" = "PESOS", "1", "0")
                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSDETALLELISTAPRECIO"
                rsListaPrecios.MoveNext
            Next vlintcontador
            'pGuardarLogTransaccion Me.Name, EnmGrabar, lngGraba, "LISTA DE PRECIOS DESDE PRESUPUESTOS", txtCveConcepto.Text
        End If
                        
        EntornoSIHO.ConeccionSIHO.CommitTrans
    End With
    
    flngGuardaPaquete = lngCvePaquete
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngGuardaPaquete"))
End Function


Private Sub pConceptosFactura()
    Dim rsConcepto As New ADODB.Recordset
    Dim rsConceptoHosp As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "SELECT smiCveConcepto, chrDescripcion " & _
                     "FROM PvConceptoFacturacion " & _
                     "WHERE bitActivo = 1 AND intTipo = 0 " & _
                     "ORDER BY chrDescripcion"
    Set rsConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboConceptoFactura, rsConcepto, 0, 1
    rsConcepto.Close
    
'    vlstrSentencia = "SELECT smiCveConcepto, TRIM(chrDescripcion) " & _
'                     "FROM PvConceptoFacturacion " & _
'                     "WHERE bitActivo = 1 " & _
'                     "AND intTipo = 0"
'    Set rsConceptoHosp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'
'    pLlenarCboRs cboMovConceptoFactura, rsConceptoHosp, 0, 1
'    cboMovConceptoFactura.ListIndex = 0
'    rsConceptoHosp.Close
End Sub
Private Function flngBuscaConcepto(vlintClaveGrupo As Long) As Long
' Funcion que busca el grupo en el grid y regresa su numero de concepto de facturacion
    Dim vlintcontador As Integer
    
    flngBuscaConcepto = 0
    
'    With grdGrupos
'        For vlintContador = 1 To .Rows - 1
'            flngBuscaConcepto = IIf(.RowData(vlintContador) = vlintClaveGrupo, .TextMatrix(vlintContador, 2), flngBuscaConcepto)
'        Next
'    End With

End Function
Private Sub pCargaElementos()
On Error GoTo NotificaError
    Dim vlintcontador As Integer
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    cboConceptoFactura.ListIndex = -1
    cboTratamiento.ListIndex = -1
    txtNumPaquete.Text = ""
    txtMotivoNoAutorizado.Text = ""
    
    '------------------------
    ' Limpieza del grid
    '------------------------
    pLimpiaRenglonMSHFlexGrid grdElementos, 1
    grdElementos.Rows = 2
    grdElementos.Cols = 0
    '------------------------
    ' Configurar el grid
    '------------------------
    pConfiguraGrid
    grdElementos.RowData(1) = -1
    
    If Trim(txtSeleArticulo.Text) = "" Then Exit Sub
    
    '-- Select para cargar los elementos ---
    '------------------------------------------------------------------------
    ' Nota
    ' Le quite el FILTRO del DEPARTAMENTO, porque Sergio me dijo que así les servía mas
    ' 06-Junio-2002
    '-------------------------------------------------------------------------
    vlstrSentencia = "Select ivarticulo.intIDArticulo Clave, vchNombreComercial Descripcion, 'AR' as Tipo " & _
        " FROM ivarticulo where ivarticulo.vchEstatus = 'ACTIVO' AND chrCveArtMedicamen <> 2 " & _
        " and " & IIf(Not optCodigoBarras.Value, IIf(optDescripcion.Value, "vchNombreComercial", "chrCveArticulo") & " like '" & Trim(txtSeleArticulo.Text) & "%'", _
        " ivarticulo.chrCveArticulo = (select chrCveArticulo from IvCodigoBarrasArticulo where vchCodigoBarras = '" & Trim(txtSeleArticulo.Text) & "')")
                                                                '" AND ivarticulo.smiCveConceptFact in (select ConFac.smiCveConcepto from pvConceptoFacturacion ConFac where ConFac.smiDepartamento = " & Trim(Str(vgintNumeroDepartamento)) & ")"
        ' Sergio me pidio que quitara el filtro 07-Jun-02
    If Not optCodigoBarras.Value Then 'Osea que si es por Codigo de barras, pues que no busque otra cosa que no sean articulos
        vlstrSentencia = vlstrSentencia & "    Union " & _
            " Select intNumPaquete as Clave, pvpaquete.chrDescripcion as Descripcion, 'PA' as Tipo " & _
            " FROM pvpaquete where pvpaquete.bitActivo = 1 " & _
            " and " & IIf(optDescripcion.Value, "chrDescripcion", "cast(intNumPaquete as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'" & _
            "    Union " & _
            " Select intCveExamen as Clave, laExamen.chrNombre as Descripcion, 'EX' as Tipo " & _
            " FROM laExamen where laExamen.bitEstatusActivo = 1 " & _
            " and " & IIf(optDescripcion.Value, "chrNombre", "cast(intCveExamen as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'"
         vlstrSentencia = vlstrSentencia & "    Union " & _
            " Select intCveGrupo as Clave, LaGrupoExamen.chrNombre as Descripcion, 'GE' as Tipo " & _
            " FROM LaGrupoExamen where LaGrupoExamen.BitEstatusActivo = 1 " & _
            " and " & IIf(optDescripcion.Value, "chrNombre", "cast(intCveGrupo as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'" & _
            "    Union " & _
            " Select intCveEstudio as Clave, ImEstudio.vchNombre as Descripcion, 'ES' as Tipo " & _
            " FROM ImEstudio where ImEstudio.BitStatusActivo = 1 " & _
            " and " & IIf(optDescripcion.Value, "vchNombre", "cast(intCveEstudio as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'" & _
            "    Union " & _
            " Select IntCveConcepto as Clave, PvOtroConcepto.chrDescripcion as Descripcion, 'OC' as Tipo " & _
            " FROM PvOtroConcepto where PvOtroConcepto.BitEstatus = 1 " & _
            " and " & IIf(optDescripcion.Value, "chrDescripcion", "cast(intCveConcepto as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'" & _
            "    Union " & _
            " Select intCveGrupo as Clave, PvGrupoCargo.vchNombre as Descripcion, 'GC' as Tipo " & _
            " FROM PvGrupoCargo where PvGrupoCargo.BitActivo = 1 " & _
            " and " & IIf(optDescripcion.Value, "vchNombre", "cast(intCveGrupo as char(10))") & " like '" & Trim(txtSeleArticulo.Text) & "%'" & _
            " ORDER BY Descripcion "
    End If
    
    '------------------------
    ' Abre el RS
    '------------------------
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly, 100)
    
    '------------------------
    ' Cargar el grid
    '------------------------
    With grdElementos
        .Redraw = False
        .Visible = False
        Do While Not rs.EOF
            If .RowData(1) <> -1 Then
                 .Rows = .Rows + 1
                 .Row = .Rows - 1
             End If
            .RowData(.Row) = rs!clave
            .TextMatrix(.Row, 1) = rs!Descripcion
            .TextMatrix(.Row, 2) = rs!tipo
            rs.MoveNext
        Loop
        .Redraw = True
        .Visible = True
        .Row = 1
        .Col = 1
    End With
    
    rs.Close
    
    If grdElementos.RowData(1) = -1 Then 'Significa que esta vacia
        grdElementos.Enabled = False
    Else
        grdElementos.Enabled = True
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaElementos"))
End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError

    With grdElementos
        .Cols = 3
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción del cargo|Tipo"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 4900 'Descripción del cargo
        .ColWidth(2) = 600  'Tipo
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarVertical
    End With
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGrid"))
End Sub

Private Sub cboConceptoFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub


Private Sub cboEmpresa_Click(Index As Integer)
'    *********************************************************************************************
'    Valida los cambios de parametros (*Tipo de Convenio, para actualizar el precio y no guardar datos de otros parametros)
If grdPresupuesto(0).TextMatrix(1, 1) <> "" Then
    If cboEmpresa(0).Enabled And vgstrEmpresa <> Trim(cboEmpresa(0).Text) And vgstrEmpresa <> "" And Trim(cboEmpresa(0).Text) <> "" Then
        MsgBox "Al cambiar la Empresa se deben actualizar precios", vbOKOnly + vbInformation, "Mensaje"
        cmdActualizaPrecios(0).SetFocus
        vgstrEmpresa = Trim(cboEmpresa(0).Text)
    Else
        vgstrEmpresa = Trim(cboEmpresa(0).Text)
    End If
Else
    vgstrEmpresa = Trim(cboEmpresa(0).Text)
End If
'    *********************************************************************************************
End Sub

Private Sub CboEmpresa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            optDescripcion.SetFocus
        Else
            chkTomarDescuentos(1).SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboEmpresa_KeyDown"))
End Sub

Private Sub cboProcedimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboProcedimiento_KeyDown"))
End Sub


Private Sub cboTipoConvenio_Click(Index As Integer)
    Dim rsEmpresa As New ADODB.Recordset
    Dim rsempresapaciente As New ADODB.Recordset
    Dim vlstrSentencia As String
         
    On Error GoTo NotificaError
         
    cboEmpresa(Index).Clear
    cboEmpresa(Index).ListIndex = -1
    cboEmpresa(Index).Enabled = False
    lstEmpresa(Index).Enabled = False
         
    If cboTipoConvenio(Index).ListIndex <> -1 Then
       vlstrSentencia = "select distinct ccEmpresa.intCveEmpresa, ccEmpresa.vchDescripcion from CcEmpresa " & _
                        "where CcEmpresa.tnyCveTipoConvenio =" & str(cboTipoConvenio(Index).ItemData(cboTipoConvenio(Index).ListIndex)) & " and bitActivo = 1"
       
       Set rsEmpresa = frsRegresaRs(vlstrSentencia)
       If rsEmpresa.RecordCount <> 0 Then
            pLlenarCboRs cboEmpresa(Index), rsEmpresa, 0, 1
            cboEmpresa(Index).ListIndex = 0
            cboEmpresa(Index).Enabled = True
            lstEmpresa(Index).Enabled = True
       End If
    End If
'    *********************************************************************************************
'    Valida los cambios de parametros (*Tipo de Convenio, para actualizar el precio y no guardar datos de otros parametros)
If grdPresupuesto(0).TextMatrix(1, 1) <> "" Then
    If cboTipoConvenio(0).Enabled And vgstrTipoConvenio <> Trim(cboTipoConvenio(0).Text) And vgstrTipoConvenio <> "" And Trim(cboTipoConvenio(0).Text) <> "" Then
        MsgBox "Al cambiar el Tipo de Convenio se deben actualizar precios", vbOKOnly + vbInformation, "Mensaje"
        cmdActualizaPrecios(0).SetFocus
        vgstrTipoConvenio = Trim(cboTipoConvenio(0).Text)
    Else
        vgstrTipoConvenio = Trim(cboTipoConvenio(0).Text)
    End If
Else
    vgstrTipoConvenio = Trim(cboTipoConvenio(0).Text)
End If
'    *********************************************************************************************
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoConvenio_Click"))
    Unload Me
End Sub


Private Sub cboTipoConvenio_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If cboEmpresa(Index).Enabled Then
            cboEmpresa(Index).SetFocus
        Else
            If Index = 0 Then
                optDescripcion.SetFocus
            Else
                chkTomarDescuentos(1).SetFocus
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoConvenio_KeyDown"))
End Sub

Private Sub cboTipoPaciente_Click(Index As Integer)
On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    cboTipoConvenio(Index).Clear
    cboTipoConvenio(Index).ListIndex = -1
    cboTipoConvenio(Index).Enabled = False
    lblTipoConvenio(Index).Enabled = False
    
    cboEmpresa(Index).Clear
    cboEmpresa(Index).ListIndex = -1
    cboEmpresa(Index).Enabled = False
    lstEmpresa(Index).Enabled = False

    If cboTipoPaciente(Index).ListIndex > -1 Then
        vlstrSentencia = "Select bitUtilizaConvenio from AdTipoPaciente where tnyCveTipoPaciente = " & cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex)
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rs.RecordCount <> 0 Then
            If rs.Fields(0) = 0 Then
                cboEmpresa(Index).ListIndex = -1
                cboEmpresa(Index).Enabled = False
                lstEmpresa(Index).Enabled = False
            ElseIf vlTipoPacienteSocio = cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex) Then
                cboEmpresa(Index).ListIndex = -1
                cboEmpresa(Index).Enabled = False
                lstEmpresa(Index).Enabled = False
            Else
                cboEmpresa(Index).ListIndex = -1
                cboEmpresa(Index).Enabled = True
                lstEmpresa(Index).Enabled = True
            End If
        End If
        rs.Close
    
        vlstrSentencia = "Select chrTipo, bitDesconocido, bitUtilizaConvenio, bitFamiliar From AdTipoPaciente Where tnyCveTipoPaciente = " & str(cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex))
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount <> 0 Then
            vlstrTipoTipoPaciente = Trim(rs!chrTipo)
            'vlblnUtilizaConvenio = IIf(rs!bitUtilizaConvenio = 1, True, False)
        End If
        rs.Close

        'Tipos de convenio
        If vlstrTipoTipoPaciente = "CO" Then
            vlstrSentencia = "Select tnyCveTipoConvenio, vchDescripcion From CcTipoConvenio"
            Set rs = frsRegresaRs(vlstrSentencia)
            If rs.RecordCount <> 0 Then
                pLlenarCboRs cboTipoConvenio(Index), rs, 0, 1
                cboTipoConvenio(Index).ListIndex = 0
                cboTipoConvenio(Index).Enabled = True
                lblTipoConvenio(Index).Enabled = True
            End If
            rs.Close
        End If
    End If
'    *********************************************************************************************
'    Valida los cambios de parametros (*TipoPaciente para actualizar el precio y no guardar datos de otros parametros)
If grdPresupuesto(0).TextMatrix(1, 1) <> "" Then
    If cboTipoPaciente(0).Enabled And vgstrTipoPaciente <> Trim(cboTipoPaciente(0).Text) And vgstrTipoPaciente <> "" And Trim(cboTipoPaciente(0).Text) <> "" Then
        MsgBox "Al cambiar el Tipo de Paciente se deben actualizar precios", vbOKOnly + vbInformation, "Mensaje"
        cmdActualizaPrecios(0).SetFocus
        vgstrTipoPaciente = Trim(cboTipoPaciente(0).Text)
    Else
        vgstrTipoPaciente = Trim(cboTipoPaciente(0).Text)
    End If
Else
    vgstrTipoPaciente = Trim(cboTipoPaciente(0).Text)
End If
'    *********************************************************************************************
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoPaciente_Click"))
End Sub

Private Sub cboTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If cboTipoConvenio(Index).Enabled Then
            cboTipoConvenio(Index).SetFocus
        Else
            If Index = 0 Then
                optDescripcion.SetFocus
            Else
                chkTomarDescuentos(1).SetFocus
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTipoPaciente_KeyDown"))
End Sub

Private Sub cboTratamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdCrearPaquete.SetFocus
End Sub


Private Sub chkTomarDescuentos_Click(Index As Integer)
On Error GoTo NotificaError

    If chkTomarDescuentos(Index).Value = 1 Then
        lblDescuento(Index).Visible = False
        txtPorcentajeDescuento(Index).Visible = False
        lblPorcentaje(Index).Visible = False
        cmdAplicaDescuento(Index).Visible = False
    Else
        lblDescuento(Index).Visible = True
        txtPorcentajeDescuento(Index).Visible = True
        lblPorcentaje(Index).Visible = True
        cmdAplicaDescuento(Index).Visible = True
        'If Not vlblnCargando Then txtPorcentajeDescuento.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkTomarDescuentos_Click"))
End Sub

Private Sub chkTomarDescuentos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
            If fblnCanFocus(txtPorcentajeDescuento(Index)) Then
                txtPorcentajeDescuento(Index).SetFocus
            Else
                cmdActualizaPrecios(Index).SetFocus
            End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkTomarDescuentos_KeyDown"))
End Sub


Private Sub cmdAceptar_Click()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "Select * from pvMensajePresupuesto"
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount = 0 Then
        rs.AddNew
    End If
    
    rs!chrMensaje = Trim(txtMensaje.Text) & " "
    rs!chrNotas = Trim(txtNotas.Text) & " "
    rs.Update
    rs.Close
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    vgbitParametros = False
    freParametros.Visible = False

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAceptar_Click"))
End Sub

Private Sub pCargaTiposPaciente()
    On Error GoTo NotificaError

    Dim vlintcontador As Integer
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select tnyCveTipoPaciente, vchDescripcion from AdTipoPaciente order by vchDescripcion"
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    Do While Not rs.EOF
        cboTipoPaciente(0).AddItem rs!VCHDESCRIPCION
        cboTipoPaciente(0).ItemData(cboTipoPaciente(0).newIndex) = rs!tnyCveTipoPaciente
        cboTipoPaciente(1).AddItem rs!VCHDESCRIPCION
        cboTipoPaciente(1).ItemData(cboTipoPaciente(1).newIndex) = rs!tnyCveTipoPaciente
        rs.MoveNext
    Loop
    
    If cboTipoPaciente(0).ListCount > 0 Then
        cboTipoPaciente(0).ListIndex = 0
    End If
    rs.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaTiposPaciente"))
End Sub

Private Sub pCargaEmpresas()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim bitactivo As Integer
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select intCveEmpresa, vchDescripcion, bitactivo from ccEmpresa order by vchDescripcion"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    
    Do While Not rs.EOF
        bitactivo = rs.Fields("bitactivo")
        If bitactivo = 1 Then
            cboEmpresa(0).AddItem rs!VCHDESCRIPCION
            cboEmpresa(0).ItemData(cboEmpresa(0).newIndex) = rs!intcveempresa
            cboEmpresa(1).AddItem rs!VCHDESCRIPCION
            cboEmpresa(1).ItemData(cboEmpresa(1).newIndex) = rs!intcveempresa
            rs.MoveNext
        Else
            rs.MoveNext
        End If
    Loop
    If cboEmpresa(0).ListCount > 0 Then
        cboEmpresa(0).ListIndex = 0
    End If
    rs.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaEmpresas"))
End Sub

Private Sub pConfiguraGridCargos(vlIndex As Integer)
    On Error GoTo NotificaError

    With grdPresupuesto(vlIndex)
        .Cols = intTotalColsgrdPresupuesto
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = strFormatgrdPresupuesto
        .ColWidth(0) = 100 'Fix
        .ColWidth(ColDescripciongrdPresupuesto) = 4250 'Descripción
        .ColWidth(ColCostogrdPresupuesto) = 1200 'Costo
        .ColWidth(ColMargenUtilidadgrdPresupuesto) = 1200 'Margen utilidad
        .ColWidth(ColPreciogrdPresupuesto) = 1200 'Precio
        .ColWidth(ColCantidadgrdPresupuesto) = 750  'Cantidad
        .ColWidth(ColSubtotalgrdPresupuesto) = 1300 'Subtotal
        .ColWidth(ColDescuentogrdPresupuesto) = 1250 'Descuento en cantidad
        .ColWidth(ColMontogrdPresupuesto) = 1300 'Monto
        .ColWidth(ColMontoIVAgrdPresupuesto) = 1250    'Monto IVA
        .ColWidth(ColTipogrdPresupuesto) = 500  'Tipo
        .ColWidth(ColIVAgrdPresupuesto) = 0    'Iva
        .ColWidth(ColPorcentajeDescgrdPresupuesto) = 0    'Porcentaje de descuento
        .ColWidth(ColContenidogrdPresupuesto) = 0   'Contenido
        .ColWidth(ColModoDescuentogrdPresupuesto) = 0   'Modo de Descuento Inventario
        .ColWidth(ColRowDatagrdPresupuesto) = 0   'RowData
        .ColAlignment(ColDescripciongrdPresupuesto) = flexAlignLeftBottom
        .ColAlignment(ColCostogrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColMargenUtilidadgrdPresupuesto) = flexAlignCenterCenter
        .ColAlignment(ColPreciogrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColCantidadgrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColSubtotalgrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColDescuentogrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColMontogrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColMontoIVAgrdPresupuesto) = flexAlignRightCenter
        .ColAlignment(ColTipogrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColDescripciongrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColCostogrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColMargenUtilidadgrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColPreciogrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColCantidadgrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColSubtotalgrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColDescuentogrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColMontogrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColMontoIVAgrdPresupuesto) = flexAlignCenterCenter
        .ColAlignmentFixed(ColTipogrdPresupuesto) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, ColPreciogrdPresupuesto) = ""
        .TextMatrix(1, ColRowDatagrdPresupuesto) = "-1"
    End With
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridCargos"))
End Sub

Private Sub pSeleccionaElemento(vlintCantidad As Integer)
    On Error GoTo NotificaError

   Dim rs As New ADODB.Recordset
   Dim vlstrSentencia As String
   Dim vlstrCualLista As String
   Dim vllngClaveElemento As Long
   Dim vlintPosicion As Integer
   Dim vlintcontador As Integer
   Dim lstListas As ListBox
   Dim vldblPrecio As Double
   Dim vldblIncrementoTarifa As Double 'Para el incremento en la tarifa
   Dim vldblSubtotal As Double
   Dim vldblDescuento As Double
   Dim vldblIVA As Double
   Dim vldblDescUnitario As Double
   Dim vldblTotDescuento As Double
   Dim vlstrTipoDescuento As String
   Dim vllngEmpresa As String
   Dim vldtmFecha As Date
   Dim vllngContenido As Long
   Dim vlintModoDescuentoInventario As Integer
   Dim vgintTipoPaciente As Integer
   Dim vlintClaveArticulo As String 'Clave del articulo como INTEGER para mandarlo al SP de Precios
   Dim vldblCosto As Double
   Dim rsCosto As ADODB.Recordset
   Dim vllngPosicion As Long
    
    If cboTipoPaciente(0).ListIndex = -1 Then
        MsgBox "Es necesario seleccionar el tipo de paciente.", vbOKOnly + vbInformation, "Mensaje"
        cboTipoPaciente(0).SetFocus
        Exit Sub
    End If
    If grdElementos.Cols = 1 Then
        Exit Sub
    End If
    If Val(Format(vlintCantidad, "#########.##")) > 0 And grdElementos.TextMatrix(1, 1) <> "" Then
        With grdPresupuesto(0)
          'Bit mostrar columnas costo,margen de utilidad y precio solo para personas que tengan control total
          'Usaremos la misma seguridad que usan para ver el pago
          If Not fblnRevisaPermiso(vglngNumeroLogin, 7037, "C", True) Then
              'El usuario no tiene permiso para ver las columnas
              .ColWidth(ColMargenUtilidadgrdPresupuesto) = 0
              .ColWidth(ColPreciogrdPresupuesto) = 0
              .ColWidth(ColCostogrdPresupuesto) = 0
          Else
              'El usuario tiene permiso para ver las columnas
              .ColWidth(ColMargenUtilidadgrdPresupuesto) = 1200
              .ColWidth(ColPreciogrdPresupuesto) = 1200
              .ColWidth(ColCostogrdPresupuesto) = 1200
          End If
      
         vlstrCualLista = grdElementos.TextMatrix(grdElementos.Row, 2)
         vllngClaveElemento = grdElementos.RowData(grdElementos.Row)
            
         vllngPosicion = FintBuscaEnRowData(grdPresupuesto(0), CLng(grdElementos.RowData(grdElementos.Row)), vlstrCualLista)
         If vllngPosicion = -1 Then
             'Cuando no esta en la lista
         Else
             .Row = vllngPosicion
             MsgBox SIHOMsg(1632), vbOKOnly + vbInformation, "Mensaje"
             Exit Sub
         End If
         
         If vllngClaveElemento <> -1 Then
            If vlstrCualLista <> "PA" Then
               '--------------------------------------------
               ' PRECIOS
               '--------------------------------------------
               If vlstrCualLista = "CI" Then
                  'Cuando es una cirugia no se agarra de ExCirugia sino de PVOtroConcepto
                  vlstrSentencia = "Select INTCVECIRUGIA  Cargo from ExCirugia where intCveCirugia = " & RTrim(str(vllngClaveElemento))
                  Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                  
                  pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                  vgstrParametrosSP = CLng(rs!Cargo) & "|" & "OC" & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                  frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                  pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                  
                  rs.Close
               Else
                  If (cboEmpresa(0).ListIndex < 0 Or cboEmpresa(0).Enabled = False) And cboEmpresa(0).List(cboEmpresa(0).ListIndex) = "" Then
                      vllngEmpresa = 0
                  Else
                      vllngEmpresa = cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex)
                  End If
                  
                  '-------------------
                  'Para los precios de los articulo se maneja con la Clave del articulo
                  '-------------------
                  If vlstrCualLista <> "GC" Then
                        If vlstrCualLista = "AR" Then 'Nomas para los articulos
                            vlstrSentencia = "select chrCveArticulo ClaveArticulo, intIdArticulo from ivArticulo where intIDArticulo = " & str(vllngClaveElemento)
                            vlintClaveArticulo = CLng(Val(frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)!intIdArticulo))
                            pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                            vgstrParametrosSP = vlintClaveArticulo & "|" & vlstrCualLista & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & vllngEmpresa & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                        Else
                            pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                            vgstrParametrosSP = vllngClaveElemento & "|" & vlstrCualLista & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & vllngEmpresa & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                        End If
                        frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                        pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                  Else
                        pElementoGrupoPred vllngClaveElemento, 0
                        vldblPrecio = vgdblPrecioPredGrupo
                        'pElementoGrupoConv (vllngClaveElemento)
                        'vldblPrecioConvenio = vgdblPrecioConvGrupo
                  End If
                                    
                  vgstrParametrosSP = IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo) & "|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & vldblPrecio & "|" & vgintClaveEmpresaContable & "|" & vllngEmpresa
                  Set rsCosto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELOBTENERCOSTO")
                  vldblCosto = rsCosto!costo
                  
               End If
                   
               vgintTipoPaciente = cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex)
               If vlstrCualLista = "AR" Or (vlstrCualLista = "GC" And (vgstrTipoPredGrupo = "AR" Or vgstrTipoPredGrupo = "ME")) Then 'Nomas para los articulos
                  '-------------------
                  ' Tipo de descuento de Inventario
                  '-------------------
                  vlstrSentencia = "select intContenido Contenido, substring(vchNombreComercial,1,50) Articulo,  ivUA.vchDescripcion UnidadAlterna,  ivUM.vchDescripcion UnidadMinima" & _
                                  " From ivArticulo " & _
                                  " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                  " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                  " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", vllngClaveElemento, vglngClavePredGrupo) '>><<
                  
                  Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                  
                  vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
                  vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                  If vllngContenido > 1 Then
                    If vlstrCualLista <> "GC" Then
                      If MsgBox("¿Desea realizar la venta de " & Trim(rs!Articulo) & " por " & Trim(rs!UnidadAlterna) & "?" & Chr(13) & "Si selecciona NO, se venderá por " & Trim(rs!UnidadMinima) & ".", vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                          vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
                      End If
                    Else
                        vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
                    End If
                  End If
                  rs.Close
               End If
               
               If vldblPrecio = -1 Or vldblPrecio = 0 Then
                  MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                  Exit Sub
               End If
               
               '-----------------------
               'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
               '-----------------------
               If vlintModoDescuentoInventario = 1 Then
                  vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                  vldblCosto = vldblCosto / CDbl(vllngContenido)
               End If
               
               If vldblPrecio = -1 Or vldblPrecio = 0 Then
                  MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                  Exit Sub
               End If
               
               vldblSubtotal = vldblPrecio * CInt(vlintCantidad)
               vldblTotDescuento = 0
               If chkTomarDescuentos(0).Value Then
                   'Descuentos
                   vgstrParametrosSP = IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & vllngEmpresa & "|" & 0 & "|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo) & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, , True)
                   Set rsPvSelDescuento = frsEjecuta_SP(vgstrParametrosSP, "SP_PvSelDescuento", False)

                  'Por Tipo de Paciente
                  '-------------------------------------
                   vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 1, 1)
                   If vlstrTipoDescuento = "%" Then 'Porcentaje
                       vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2)) / 100
                       vldblDescUnitario = (vldblPrecio * vldblDescuento)
                       vldblTotDescuento = vldblDescUnitario * CInt(vlintCantidad)
                       vldblSubtotal = (vldblPrecio - vldblDescUnitario) * CInt(vlintCantidad)
                   Else  'Cantidad
                       vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2))
                       vldblDescUnitario = vldblDescuento
                       vldblTotDescuento = vldblDescuento
                       vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                   End If

                  
                  'Por Empresa
                  '-------------------------------------
                   vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 1, 1)
                   If vlstrTipoDescuento = "%" Then 'Porcentaje
                       vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2)) / 100
                       vldblTotDescuento = vldblTotDescuento + (((vldblSubtotal / CInt(vlintCantidad)) * vldblDescuento) * CInt(vlintCantidad))
                       vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                   Else
                        If vlstrTipoDescuento = "S" Then 'Costo
                            vldblDescuento = (vldblPrecio - vldblCosto) * CInt(vlintCantidad)
                            vldblTotDescuento = vldblTotDescuento + vldblDescuento
                            vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                        Else  'Cantidad
                            vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2))
                            vldblTotDescuento = vldblTotDescuento + vldblDescuento
                            vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                        End If
                   End If
                   rsPvSelDescuento.Close
               Else
                  If Val(txtPorcentajeDescuento(0).Text) <> 0 Then
                      vldblDescuento = Val(txtPorcentajeDescuento(0).Text)
                      vldblTotDescuento = (vldblPrecio * CInt(vlintCantidad)) * vldblDescuento / 100
                      vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                  End If
               End If
               '----------------------------------------------
               ' Procedimiento para obtener el IVA
               '----------------------------------------------
               vldblIVA = fdblObtenerIva(IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo), IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo)) / 100
               vlintPosicion = FintBuscaEnRowData(grdPresupuesto(0), vllngClaveElemento, vlstrCualLista)
               If vlintPosicion = -1 Then        'Cuando no esta en la lista
                  If CLng(.TextMatrix(1, ColRowDatagrdPresupuesto)) <> -1 Then
                      .Rows = .Rows + 1
                      .Row = .Rows - 1
                  End If
               Else
                  .Row = vlintPosicion 'Funciona como modificación
               End If
                  
               .TextMatrix(.Row, ColDescripciongrdPresupuesto) = grdElementos.TextMatrix(grdElementos.Row, 1)
               .TextMatrix(.Row, ColCostogrdPresupuesto) = Format(vldblCosto, "$###,###,###,##0.00")
               '.TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = CInt(((vldblPrecio - vldblCosto) / vldblPrecio) * 100)
               '.TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = Format((((vldblPrecio - vldblCosto) / vldblCosto) * 100), "00.00") & "%"
                If vldblCosto > 0 Then
                    '.TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = Format(((vldblPrecio / vldblCosto) - 1) * 100, "0.00") & "%"
                    
                    'Costo total = costo * cantidad
                    'Margen = Subtotal - Costo total
                    'Margen de utilidad = (Margen / Subtotal) * 100
                    
                    'Costo total
                    vldblCostoTotal = vldblCosto * CInt(vlintCantidad)
                    'Margen
                    vldblMargen = vldblSubtotal - vldblCostoTotal
                    'Margen de utilidad
                    If vldblSubtotal > 0 Then
                        vldblmargenutilidad = (vldblMargen / vldblSubtotal) * 100
                    Else
                        vldblmargenutilidad = 0
                    End If
                    .TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = Format(vldblmargenutilidad, "0.00") & "%"
                Else
                    .TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = Format(0, "0.00") & "%"
                End If
               .TextMatrix(.Row, ColPreciogrdPresupuesto) = Format(vldblPrecio, "$###,###,###,###.00")
               .TextMatrix(.Row, ColCantidadgrdPresupuesto) = vlintCantidad
               .TextMatrix(.Row, ColSubtotalgrdPresupuesto) = Format(vldblPrecio * CInt(vlintCantidad), "$###,###,###,###.00")
               .TextMatrix(.Row, ColDescuentogrdPresupuesto) = Format(vldblTotDescuento, "$###,###,###,###.00")
               .TextMatrix(.Row, ColMontogrdPresupuesto) = Format(vldblSubtotal, "$###,###,###,###.00")
               .TextMatrix(.Row, ColTipogrdPresupuesto) = vlstrCualLista
               .TextMatrix(.Row, ColIVAgrdPresupuesto) = vldblIVA
               .TextMatrix(.Row, ColMontoIVAgrdPresupuesto) = ((vldblPrecio * vlintCantidad) - vldblTotDescuento) * vldblIVA
               .TextMatrix(.Row, ColPorcentajeDescgrdPresupuesto) = vldblTotDescuento / Val(Format(.TextMatrix(.Row, ColSubtotalgrdPresupuesto), "###########.00"))
               .TextMatrix(.Row, ColContenidogrdPresupuesto) = vllngContenido
               .TextMatrix(.Row, ColModoDescuentogrdPresupuesto) = vlintModoDescuentoInventario
               .TextMatrix(.Row, ColRowDatagrdPresupuesto) = vllngClaveElemento
               .Redraw = True
               .Refresh
               pRecalcula 0
               
            '----------------
            'P A Q U E T E S
            '----------------
            Else    'Este es para paquetes
                    If cboEmpresa(0).ListIndex < 0 Or cboEmpresa(0).Enabled = False Then
                        vllngEmpresa = 0
                    Else
                        vllngEmpresa = cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex)
                    End If
                    
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = CLng(vllngClaveElemento) & "|" & vlstrCualLista & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & vllngEmpresa & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                     
                  vgintTipoPaciente = cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex)
                  
                  If vldblPrecio = -1 Or vldblPrecio = 0 Then
                     MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                     Exit Sub
                  End If
                  
                  vldblSubtotal = vldblPrecio * CInt(vlintCantidad)
                  '----------------------------------------------
                  'Descuentos
                  '----------------------------------------------
                  vldblTotDescuento = 0
                   If chkTomarDescuentos(0).Value Then
                       'Descuentos
                       vgstrParametrosSP = IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & vllngEmpresa & "|" & 0 & "|" & IIf(vlstrCualLista = "PA", "CF", vlstrCualLista) & "|" & CLng(vllngClaveElemento) & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, , True)
                       Set rsPvSelDescuento = frsEjecuta_SP(vgstrParametrosSP, "SP_PvSelDescuento", False)
    
                      'Por Tipo de Paciente
                      '-------------------------------------
                       vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 1, 1)
                       If vlstrTipoDescuento = "%" Then 'Porcentaje
                           vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2)) / 100
                           vldblDescUnitario = (vldblPrecio * vldblDescuento)
                           vldblTotDescuento = vldblDescUnitario * CInt(vlintCantidad)
                           vldblSubtotal = (vldblPrecio - vldblDescUnitario) * CInt(vlintCantidad)
                       Else  'Cantidad
                           vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2))
                           vldblDescUnitario = vldblDescuento
                           vldblTotDescuento = vldblDescuento
                           vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                       End If
    
                      'Por Empresa
                      '-------------------------------------
                       vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 1, 1)
                       If vlstrTipoDescuento = "%" Then 'Porcentaje
                           vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2)) / 100
                           vldblTotDescuento = vldblTotDescuento + (((vldblSubtotal / CInt(vlintCantidad)) * vldblDescuento) * CInt(vlintCantidad))
                           vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                       Else  'Cantidad
                           vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2))
                           vldblTotDescuento = vldblTotDescuento + vldblDescuento
                           vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                       End If
                       rsPvSelDescuento.Close
                   Else
                      If Val(txtPorcentajeDescuento(0).Text) <> 0 Then
                          vldblDescuento = Val(txtPorcentajeDescuento(0).Text)
                          vldblTotDescuento = (vldblPrecio * CInt(vlintCantidad)) * vldblDescuento / 100
                          vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                      End If
                   End If
                   
                  
                  '----------------------------------------------
                  ' Procedimiento para obtener el IVA
                  '----------------------------------------------
                  vldblIVA = fdblObtenerIva(vllngClaveElemento, vlstrCualLista) / 100
                  vlintPosicion = FintBuscaEnRowData(grdPresupuesto(0), vllngClaveElemento, vlstrCualLista)
                  If vlintPosicion = -1 Then        'Cuando no esta en la lista
                      If CLng(.TextMatrix(1, ColRowDatagrdPresupuesto)) <> -1 Then
                          .Rows = .Rows + 1
                          .Row = .Rows - 1
                      End If
                  Else
                      .Row = vlintPosicion 'Funciona como modificación
                  End If
                  
                  .TextMatrix(.Row, ColDescripciongrdPresupuesto) = grdElementos.TextMatrix(grdElementos.Row, 1) 'Trim(rsRS!Descripcion)
                  .TextMatrix(.Row, ColPreciogrdPresupuesto) = Format(vldblPrecio, "$###,###,###,###.00")
                  .TextMatrix(.Row, ColCantidadgrdPresupuesto) = vlintCantidad
                  .TextMatrix(.Row, ColSubtotalgrdPresupuesto) = Format(vldblPrecio * CInt(vlintCantidad), "$###,###,###,###.00")
                  .TextMatrix(.Row, ColDescuentogrdPresupuesto) = Format(vldblTotDescuento, "$###,###,###,###.00")
                  .TextMatrix(.Row, ColMontogrdPresupuesto) = Format(vldblSubtotal, "$###,###,###,###.00")
                  .TextMatrix(.Row, ColTipogrdPresupuesto) = vlstrCualLista
                  .TextMatrix(.Row, ColIVAgrdPresupuesto) = vldblIVA
                  .TextMatrix(.Row, ColPorcentajeDescgrdPresupuesto) = vldblTotDescuento / Val(Format(.TextMatrix(.Row, ColSubtotalgrdPresupuesto), "###########.00"))
                  .TextMatrix(.Row, ColContenidogrdPresupuesto) = vllngContenido
                  .TextMatrix(.Row, ColModoDescuentogrdPresupuesto) = vlintModoDescuentoInventario
                  .TextMatrix(.Row, ColRowDatagrdPresupuesto) = vllngClaveElemento
                  .Redraw = True
                  .Refresh
                  pRecalcula 0
            End If
         Else
             MsgBox SIHOMsg(3), vbCritical, "Mensaje"
         End If
         vlintCantidad = 1
         pEnfocaTextBox txtSeleArticulo
      End With
   End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSeleccionaElemento"))
End Sub

Private Sub pCalculaGrid()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlstrCualLista As String
    Dim vllngClaveElemento As Long
    Dim vlintPosicion As Integer
    Dim vlintcontador As Integer
    Dim lstListas As ListBox
    Dim vldblPrecio As Double
    Dim vldblIncrementoTarifa As Double
    Dim vldblSubtotal As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim vldblDescUnitario As Double
    Dim vldblTotDescuento As Double
    Dim vlstrTipoDescuento As String
    Dim vlintCantidad As Integer
    
        With grdPresupuesto(0)
            vlintCantidad = .TextMatrix(.Row, ColCantidadgrdPresupuesto)
            vlstrCualLista = grdPresupuesto(0).TextMatrix(grdPresupuesto(0).Row, ColTipogrdPresupuesto)
            vllngClaveElemento = grdPresupuesto(0).TextMatrix(grdPresupuesto(0).Row, ColRowDatagrdPresupuesto)
            
            If vllngClaveElemento <> -1 Then

                'Precio unitario
                If vlstrCualLista = "CI" Then
                    'Cuando es una cirugia no se agarra de ExCirugia sino de PVOtroConcepto
                    vlstrSentencia = "Select intCveCargo  Cargo from ExCirugia where intCveCirugia = " & RTrim(str(vllngClaveElemento))
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)

                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = CLng(rs!Cargo) & "|" & "OC" & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 1 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                    
                    rs.Close
                Else
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = CLng(vllngClaveElemento) & "|" & vlstrCualLista & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) & "|" & IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 1 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                    
                End If
                
                If vldblPrecio = -1 Or vldblPrecio = 0 Then
                    MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                vldblSubtotal = vldblPrecio * CInt(vlintCantidad)
                vldblTotDescuento = 0
                If chkTomarDescuentos(0).Value Then
                     'Descuentos
                     vgstrParametrosSP = "I" & "|" & cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex) & "|" & cboEmpresa(0).ItemData(cboEmpresa(0).ListIndex) & "|" & 0 & "|" & vlstrCualLista & "|" & CLng(vllngClaveElemento) & "|" & vgintNumeroDepartamento & "|" & fdtmServerFecha
                     Set rsPvSelDescuento = frsEjecuta_SP(vgstrParametrosSP, "SP_PvSelDescuento", False)
                    
                    'Por Tipo de Paciente
                    '-------------------------------------
                     vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 1, 1)
                     If vlstrTipoDescuento = "%" Then 'Porcentaje
                         vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2)) / 100
                         vldblDescUnitario = (vldblPrecio * vldblDescuento)
                         vldblTotDescuento = vldblDescUnitario * CInt(vlintCantidad)
                         vldblSubtotal = (vldblPrecio - vldblDescUnitario) * CInt(vlintCantidad)
                     Else  'Cantidad
                         vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2))
                         vldblDescUnitario = vldblDescuento
                         vldblTotDescuento = vldblDescuento
                         vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                     End If
                    
                    'Por Empresa
                    '-------------------------------------
                     vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 1, 1)
                     If vlstrTipoDescuento = "%" Then 'Porcentaje
                         vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2)) / 100
                         vldblTotDescuento = vldblTotDescuento + (((vldblSubtotal / CInt(vlintCantidad)) * vldblDescuento) * CInt(vlintCantidad))
                         vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                     Else  'Cantidad
                         vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2))
                         vldblTotDescuento = vldblTotDescuento + vldblDescuento
                         vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                     End If
                     rsPvSelDescuento.Close
                End If
                
                '----------------------------------------------
                ' Procedimiento para obtener el IVA
                '----------------------------------------------
                vldblIVA = fdblObtenerIva(vllngClaveElemento, vlstrCualLista) / 100
                
                .TextMatrix(.Row, ColPreciogrdPresupuesto) = Format(vldblPrecio, "$###,###,###,###.00")
                .TextMatrix(.Row, ColCantidadgrdPresupuesto) = vlintCantidad
                .TextMatrix(.Row, ColSubtotalgrdPresupuesto) = Format(vldblPrecio * CInt(vlintCantidad), "$###,###,###,###.00")
                .TextMatrix(.Row, ColDescuentogrdPresupuesto) = Format(vldblTotDescuento, "$###,###,###,###.00")
                .TextMatrix(.Row, ColMontogrdPresupuesto) = Format(vldblSubtotal, "$###,###,###,###.00")
                .TextMatrix(.Row, ColTipogrdPresupuesto) = vlstrCualLista
                .TextMatrix(.Row, ColIVAgrdPresupuesto) = vldblSubtotal * vldblIVA
                .Redraw = True
                .Refresh
            Else
                MsgBox SIHOMsg(3), vbCritical, "Mensaje"
            End If
            vlintCantidad = 1
            grdPresupuesto(0).SetFocus
        End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaGrid"))
End Sub


Private Function fdblObtenerIva(vllngCveCargo As Long, vlstrTipoCargo As String) As Variant
    On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
        
        Select Case vlstrTipoCargo
        Case "PA"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM PvPaquete INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "PvPaquete.smiConceptoFactura = PvConceptoFacturacion.smiCveConcepto " & _
                        "where PvPaquete.intNumPaquete = " & RTrim(str(vllngCveCargo))
                        
        Case "CI"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM ExCirugia INNER JOIN " & _
                        "PvOtroConcepto ON " & _
                        "ExCirugia.intCveCargo = PvOtroConcepto.intCveConcepto INNER " & _
                        "Join " & _
                        "PvConceptoFacturacion ON " & _
                        "PvOtroConcepto.smiConceptoFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where ExCirugia.intCveCirugia = " & RTrim(str(vllngCveCargo))
        Case "OC"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM PvOtroConcepto INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "PvOtroConcepto.smiConceptoFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where PvOtroConcepto.intCveConcepto = " & RTrim(str(vllngCveCargo))
        Case "GE"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM LaGrupoExamen INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "LaGrupoExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where LaGrupoExamen.intCveGrupo = " & RTrim(str(vllngCveCargo))
        Case "EX"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM LaExamen INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "LaExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where LaExamen.intCveExamen = " & RTrim(str(vllngCveCargo))
        Case "ES"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM ImEstudio INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "ImEstudio.smiConFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where ImEstudio.intCveEstudio = " & RTrim(str(vllngCveCargo))
        Case "AR"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM IvArticulo INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "IvArticulo.smiCveConceptFact2 = PvConceptoFacturacion.smiCveConcepto " & _
                        "where IvArticulo.intIDArticulo = " & vllngCveCargo
        End Select
    
    If Trim(vlstrSentencia) = "" Then
        fdblObtenerIva = 0
        Exit Function
    End If
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        fdblObtenerIva = rs!IVA
    Else
        fdblObtenerIva = 0
    End If
    rs.Close
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fdblObtenerIva"))
End Function

Private Sub cmdAceptarNoAutorizado_Click()
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub

    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    vlstrSentencia = "UPDATE PVPRESUPUESTO SET CHRESTADO = 'N' WHERE INTCVEPRESUPUESTO = " & txtClave.Text
    pEjecutaSentencia vlstrSentencia

    Call pGuardarLogPresupuesto("N", vllngPersonaGraba, CLng(txtClave.Text), txtMotivoNoAutorizado.Text)
    Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PRESUPUESTO (Cambio de estado a NO AUTORIZADO)", txtClave.Text)
    EntornoSIHO.ConeccionSIHO.CommitTrans

    'La operación se realizó satisfactoriamente.
    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"

    fraMotivoNoAutorizado.Visible = False
    freDatos.Enabled = True
    freBotones.Enabled = True
    Frame3.Enabled = True
    pHabilitaFrames True
    txtClave_KeyDown vbKeyReturn, 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAceptarNoAutorizado_Click"))
End Sub

Private Sub cmdAceptoBuscar_Click()
    On Error GoTo NotificaError

    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    cboEmpresa(0).Enabled = False
    cboTipoPaciente(0).Enabled = False

    txtSeleArticulo = ""
    pCargaElementos

    cboEmpresa(0).ListIndex = -1
    cboProcedimiento.ListIndex = -1
    cboConceptoFactura.ListIndex = -1
    cboTratamiento.ListIndex = -1
    txtNumPaquete.Text = ""
    
    freDatos.Enabled = True
    freBotones.Enabled = True
    freVarios.Enabled = True
    If lstBuscaPresupuesto.ListIndex <> -1 And lstBuscaPresupuesto.ListCount > 0 Then
        txtClave.Text = lstBuscaPresupuesto.ItemData(lstBuscaPresupuesto.ListIndex)
        txtClave_KeyDown vbKeyReturn, 0
    Else
        pEnfocaTextBox txtNombre
        'vgstrEstadoManto = "C"
    End If
    freBusqueda.Visible = False
    Exit Sub
        
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAceptoBuscar_Click"))
End Sub

Private Sub cmdActualizaPrecios_Click(Index As Integer)
On Error GoTo NotificaError
    Dim vlintIndice As Integer
    Dim vldblPrecio As Double
    Dim vldblIncrementoTarifa As Double
    Dim vllngEmpresa As Long
    Dim vldblIVA As Double
    Dim vldblTotDescuento As Double
    Dim vldblDescuento As Double
    Dim vldblDescUnitario As Double
    Dim vldblSubtotal As Double
    Dim vlstrTipoDescuento As String
    Dim vlintCantidad As Integer
    Dim vlblnEntro As Boolean
    Dim rsPvSelDescuento As ADODB.Recordset
    Dim vldblCosto As Double
    Dim vlstrCualLista As String
    Dim vllngClaveElemento As Long
    Dim rsCosto As ADODB.Recordset
    Dim vldblPrecioPres As Double
   Dim rs As New ADODB.Recordset
   Dim vlstrSentencia As String
   Dim vllngContenido As Long
    
    vlblnEntro = False
    If (cboEmpresa(Index).ListIndex < 0 Or cboEmpresa(Index).Enabled = False) And cboEmpresa(0).List(cboEmpresa(0).ListIndex) = "" Then
        vllngEmpresa = 0
    Else
        vllngEmpresa = cboEmpresa(Index).ItemData(cboEmpresa(Index).ListIndex)
    End If

    For vlintIndice = 1 To Me.grdPresupuesto(Index).Rows - 1
        If CLng(IIf(grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto) = "", "-1", grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto))) > -1 Then
            vllngClaveElemento = grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto)
            vlstrCualLista = grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto)
            If grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto) <> "EF" Then
                vlblnEntro = True
                If vlstrCualLista <> "GC" Then
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto) & "|" & grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto) & "|" & cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex) & "|" & vllngEmpresa & "|" & IIf(optTipoPaciente(Index), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio, vldblIncrementoTarifa
                    
                    If grdPresupuesto(Index).TextMatrix(vlintIndice, ColModoDescuentogrdPresupuesto) = 1 Then
                        vldblPrecioPres = vldblPrecio
                         vldblPrecio = vldblPrecio / grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto)
                    Else
                        vldblPrecioPres = vldblPrecio
                    End If
                Else
                    pElementoGrupoPred vllngClaveElemento, Index
                    vldblPrecio = vgdblPrecioPredGrupo
                    If grdPresupuesto(Index).TextMatrix(vlintIndice, ColModoDescuentogrdPresupuesto) = 1 Then
                        If grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto) = 0 Or IsNull(grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto)) Then
                            vlstrSentencia = "select intContenido Contenido, substring(vchNombreComercial,1,50) Articulo,  ivUA.vchDescripcion UnidadAlterna,  ivUM.vchDescripcion UnidadMinima" & _
                                            " From ivArticulo " & _
                                            " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                            " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                            " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", vllngClaveElemento, vglngClavePredGrupo) '>><<
                              Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                              vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                              rs.Close
                              vldblPrecioPres = vldblPrecio
                              vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                        Else
                            vldblPrecioPres = vldblPrecio
                            vldblPrecio = vldblPrecio / grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto)
                        End If
                    Else
                        vldblPrecioPres = vldblPrecio
                    End If
                End If
                
                'vldblIVA = fdblObtenerIva(CLng(grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto)), grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto)) / 100
                vldblIVA = fdblObtenerIva(IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo), IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo)) / 100

                vldblTotDescuento = 0
                vlintCantidad = 1
                If chkTomarDescuentos(Index).Value Then
                    'vgstrParametrosSP = IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex) & "|" & vllngEmpresa & "|" & 0 & "|" & IIf(grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto) = "PA", "CF", grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto)) & "|" & grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto) & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, , True)
                    vgstrParametrosSP = IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U")) & "|" & cboTipoPaciente(Index).ItemData(cboTipoPaciente(Index).ListIndex) & "|" & vllngEmpresa & "|" & 0 & "|" & IIf(vlstrCualLista = "PA", "CF", IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo)) & "|" & IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo) & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, , True)
                    
                    Set rsPvSelDescuento = frsEjecuta_SP(vgstrParametrosSP, "SP_PvSelDescuento", False)
                    vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 1, 1)
                    If vlstrTipoDescuento = "%" Then 'Porcentaje
                        vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2)) / 100
                        vldblDescUnitario = (vldblPrecio * vldblDescuento)
                        vldblTotDescuento = vldblDescUnitario * CInt(vlintCantidad)
                        vldblSubtotal = (vldblPrecio - vldblDescUnitario) * CInt(vlintCantidad)
                    Else  'Cantidad
                        vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOTP)), 2))
                        vldblDescUnitario = vldblDescuento
                        vldblTotDescuento = vldblDescuento
                        vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                    End If
                    vlstrTipoDescuento = Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 1, 1)
                    If vlstrTipoDescuento = "%" Then 'Porcentaje
                        vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2)) / 100
                        vldblTotDescuento = vldblTotDescuento + (((vldblSubtotal / CInt(vlintCantidad)) * vldblDescuento) * CInt(vlintCantidad))
                        vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                    Else  'Cantidad
                        vldblDescuento = Val(Mid(RTrim(LTrim(rsPvSelDescuento!DESCUENTOEM)), 2))
                        vldblTotDescuento = vldblTotDescuento + vldblDescuento
                        vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                    End If
                    rsPvSelDescuento.Close
                Else
                   If Val(txtPorcentajeDescuento(Index).Text) <> 0 Then
                       vldblDescuento = Val(txtPorcentajeDescuento(Index).Text)
                       vldblTotDescuento = (vldblPrecio * CInt(vlintCantidad)) * vldblDescuento / 100
                       vldblSubtotal = (vldblPrecio * CInt(vlintCantidad)) - vldblTotDescuento
                   End If
                End If
        
                grdPresupuesto(Index).TextMatrix(vlintIndice, ColPreciogrdPresupuesto) = FormatCurrency(vldblPrecio, 2)
                grdPresupuesto(Index).TextMatrix(vlintIndice, ColDescuentogrdPresupuesto) = FormatCurrency(vldblTotDescuento, 2)
                grdPresupuesto(Index).TextMatrix(vlintIndice, ColIVAgrdPresupuesto) = vldblIVA
                grdPresupuesto(Index).TextMatrix(vlintIndice, ColPorcentajeDescgrdPresupuesto) = vldblTotDescuento / IIf((vldblPrecio * CInt(vlintCantidad)) = 0, 1, (vldblPrecio * CInt(vlintCantidad)))
                'vldblCosto = CDbl(Val(Format(grdPresupuesto(Index).TextMatrix(vlintIndice, ColCostogrdPresupuesto), "############.##")))
                
                'Actualiza el costo
                vlstrCualLista = grdPresupuesto(Index).TextMatrix(vlintIndice, ColTipogrdPresupuesto)
                vllngClaveElemento = grdPresupuesto(Index).TextMatrix(vlintIndice, ColRowDatagrdPresupuesto)
'                vgstrParametrosSP = IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo) & "|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & vldblPrecio & "|" & vgintClaveEmpresaContable & "|" & vllngEmpresa
                vgstrParametrosSP = IIf(vlstrCualLista <> "GC", vllngClaveElemento, vglngClavePredGrupo) & "|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & vldblPrecioPres & "|" & vgintClaveEmpresaContable & "|" & vllngEmpresa
                Set rsCosto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELOBTENERCOSTO")
                vldblCosto = rsCosto!costo
                If grdPresupuesto(Index).TextMatrix(vlintIndice, ColModoDescuentogrdPresupuesto) = 1 Then
                    If grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto) = 0 Or IsNull(grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto)) Then
                        vldblCosto = vldblCosto / CDbl(vllngContenido)
                    Else
                        vldblCosto = vldblCosto / grdPresupuesto(Index).TextMatrix(vlintIndice, ColContenidogrdPresupuesto)
                    End If
                End If
                grdPresupuesto(Index).TextMatrix(vlintIndice, ColCostogrdPresupuesto) = FormatCurrency(vldblCosto, 2)
                If vldblCosto > 0 Then
                    'grdPresupuesto(Index).TextMatrix(vlintIndice, ColMargenUtilidadgrdPresupuesto) = Format(((vldblPrecio / vldblCosto) - 1) * 100, "0.00") & "%"
                    vldblCostoTotal = vldblCosto * CInt(vlintCantidad)
                    vldblMargen = vldblSubtotal - vldblCostoTotal
                    If vldblSubtotal > 0 Then
                        vldblmargenutilidad = (vldblMargen / vldblSubtotal) * 100
                    Else
                        vldblmargenutilidad = 0
                    End If
                    
                    grdPresupuesto(Index).TextMatrix(vlintIndice, ColMargenUtilidadgrdPresupuesto) = Format(vldblmargenutilidad, "0.00") & "%"
                End If
            End If
        End If
    Next
    If vlblnEntro Then
        pRecalcula Index
        MsgBox "Precios actualizados satisfactoriamente.", vbOKOnly + vbInformation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdActualizaPrecios_Click"))
End Sub

Private Sub cmdAplicaDescuento_Click(Index As Integer)
On Error GoTo NotificaError

    Dim X As Long
    
    If Val(txtPorcentajeDescuento(Index).Text) > 100 Then
        MsgBox SIHOMsg(35), vbCritical, "Mensaje"
        txtPorcentajeDescuento(Index).SetFocus
        Exit Sub
    End If
    
    If Trim(grdPresupuesto(Index).TextMatrix(1, ColDescripciongrdPresupuesto)) <> "" Then
    
        For X = 1 To grdPresupuesto(Index).Rows - 1
            If CLng(IIf(grdPresupuesto(Index).TextMatrix(X, ColRowDatagrdPresupuesto) = "", "-1", grdPresupuesto(Index).TextMatrix(X, ColRowDatagrdPresupuesto))) > -1 Then
                grdPresupuesto(Index).TextMatrix(X, ColPorcentajeDescgrdPresupuesto) = Val(txtPorcentajeDescuento(Index).Text) / 100
            End If
        Next X
        
        pRecalcula Index
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAplicaDescuento_Click"))
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo NotificaError

    vgstrEstadoManto = "BC"
    freBusqueda.Top = ((Me.Height - freBusqueda.Height) / 2.3)
    freBusqueda.Left = ((Me.Width - freBusqueda.Width) / 2)
    freBusqueda.Visible = True
    freDatos.Enabled = False
    freBotones.Enabled = False
    pEnfocaTextBox txtBuscaNombre
    pHabilitaFrames False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBuscar_Click"))
End Sub

Private Sub cmdCrearDuplicado_Click()
On Error GoTo NotificaError

    Dim rsNuevoPresupuesto As ADODB.Recordset
    Dim vlblnHayCargos As Boolean
    Dim X As Integer
    
    vlblnHayCargos = False
    For X = 1 To Me.grdPresupuesto(1).Rows - 1
        If CLng(IIf(grdPresupuesto(1).TextMatrix(X, ColRowDatagrdPresupuesto) = "", "-1", grdPresupuesto(1).TextMatrix(X, ColRowDatagrdPresupuesto))) > -1 Then
            vlblnHayCargos = True
        End If
    Next
    If vlblnHayCargos = False Then
        MsgBox "El presupuesto no tiene elementos para duplicar.", vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    For X = 1 To grdPresupuesto(1).Rows - 1
        If CLng(IIf(grdPresupuesto(1).TextMatrix(X, ColRowDatagrdPresupuesto) = "", "-1", grdPresupuesto(1).TextMatrix(X, ColRowDatagrdPresupuesto))) > -1 Then
            If Val(Format(grdPresupuesto(1).TextMatrix(X, ColPreciogrdPresupuesto), "###########.00")) = 0 Then
                'El elemento seleccionado no cuenta con un precio capturado.
                MsgBox Replace(SIHOMsg(301), "El elemento seleccionado no cuenta", "Existen elementos que no cuentan"), vbOKOnly + vbExclamation, "Mensaje"
                grdPresupuesto(1).Row = X
                grdPresupuesto(1).Col = ColPreciogrdPresupuesto
                grdPresupuesto(1).SetFocus
                Exit Sub
            End If
        End If
    Next
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    Set rsNuevoPresupuesto = frsRegresaRs("Select * from pvPresupuesto where intcvepresupuesto = -1", adLockOptimistic)
    rsNuevoPresupuesto.AddNew
    
    'Inicio de Grabada en tabla maestro y detalle
    If Not fnGrabaPresupuesto(rsNuevoPresupuesto, 1, True) Then
        FraDuplicar.Visible = False
        pHabilitaFrames True
        Frame3.Enabled = True
        cmdActualizaPrecios(0).Enabled = True
        txtClave.SetFocus
    Else
        cmdSalirDuplicado_Click
        txtClave.Text = vlintCvePresupuesto
        txtClave_KeyDown vbKeyReturn, 0
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCrearDuplicado_Click"))
End Sub

Private Sub cmdCrearPaquete_Click()
On Error GoTo NotificaError
    vgstrEstadoManto = "PA"
    pHabilitaBotones 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
    vllngCvePaquete = flngGuardaPaquete()
    If vllngCvePaquete > 0 Then
        vllngCtaPaciente = flngGuardaPaciente(vllngCvePaquete)
        If vllngCtaPaciente > 0 Then
            pHabilitaBotones 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0
            txtNumCuenta.Text = vllngCtaPaciente
            txtClave_KeyDown vbKeyReturn, 0
            vgstrEstadoManto = "C"
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCrearPaquete_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
            vlstrSentencia = "Delete from pvDetallePresupuesto where intCvePresupuesto = " & Trim(txtClave.Text)
            pEjecutaSentencia vlstrSentencia
            vlstrSentencia = "Delete from pvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text)
            pEjecutaSentencia vlstrSentencia
        EntornoSIHO.ConeccionSIHO.CommitTrans
        pEnfocaTextBox txtClave
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub cmdDuplicar_Click()
On Error GoTo NotificaError
    'pHabilitaBotones 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0
    vgstrEstadoManto = "DU"
    pHabilitaFrames False
    Frame3.Enabled = False
    cmdActualizaPrecios(0).Enabled = False
    FraDuplicar.Top = 695
    FraDuplicar.Left = 95
    FraDuplicar.Width = 13155
    cboTipoPaciente(1).ListIndex = 0
    cboEmpresa(1).ListIndex = -1
    pLimpiaGrid grdPresupuesto(1)
    pConfiguraGridCargos 1
    pMoverInformacion grdPresupuesto(0), grdPresupuesto(1)
    FraDuplicar.Visible = True
    cboTipoPaciente(1).SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDuplicar_Click"))
End Sub

Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsDetalle As New ADODB.Recordset
    Dim vlintcontador As Integer
    Dim vlblnFueAlta As Boolean
    Dim vlintCvePresupuesto As Integer
    Dim SQL As String
    
    If cboTipoPaciente(0).ListIndex = -1 Then
        MsgBox "Seleccionar el tipo de paciente.", vbOKOnly + vbInformation, "Mensaje"
        cboTipoPaciente(0).SetFocus
        Exit Sub
    End If
    If CLng(grdPresupuesto(0).TextMatrix(1, ColRowDatagrdPresupuesto)) = -1 Then
        MsgBox "No es posible guardar el presupuesto ya que no contiene ningún elemento.", vbCritical, "Mensaje"
        Exit Sub
    End If
    
    If fblnDatosValidos() Then
        'If MsgBox(SIHOMsg(4), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            '--------------------------------------------------------
            ' Persona que graba
            '--------------------------------------------------------
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba = 0 Then Exit Sub
            
            vlblnFueAlta = False
            If vgblnEsNuevo = True Then
                vglngClaveSiguiente = frsRegresaRs("SELECT LAST_NUMBER MAXIMO FROM USER_SEQUENCES WHERE SEQUENCE_NAME = 'SEC_PVPRESUPUESTO'").Fields(0)
                If txtClave.Text <> vglngClaveSiguiente Then
                    txtClave.Text = vglngClaveSiguiente
                End If
                vlstrSentencia = "select * from PvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text)
                Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                If rs.RecordCount = 0 Then 'Pos no existe
                    rs.AddNew
                    vlblnFueAlta = True
                Else
                    MsgBox SIHOMsg(1), vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
            Else
                vlstrSentencia = "select * from PvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text)
                Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                If rs.RecordCount = 0 Then 'Pos no existe
                    MsgBox SIHOMsg(1), vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
            End If
                                    
            'vlblnFueAlta = False
            'vlstrSentencia = "select * from PvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text)
            'Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            'If rs.RecordCount = 0 Then 'Pos no existe
            '    rs.AddNew
            '    vlblnFueAlta = True
            'End If
            
            'Inicio de Grabada en tabla maestro y detalle
            If Not fnGrabaPresupuesto(rs, 0, vlblnFueAlta) Then
                cmdImprimir.Enabled = True
                optTicket.Enabled = True
                optCarta.Enabled = True
                optTicket.Value = 1
                If fblnCanFocus(cmdImprimir) Then cmdImprimir.SetFocus
            Else
                txtClave_KeyDown vbKeyReturn, 0
                Exit Sub
            End If
        'End If
    End If
    vgstrTipoPaciente = vgstrTipoConvenio = vgstrEmpresa = Empty
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabarRegistro_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    Dim X As Long
    
    fblnDatosValidos = True
        
    If fblnDatosValidos And Not IsDate(mskFechaPresupuesto.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        mskFechaPresupuesto.SetFocus
    End If
    If fblnDatosValidos And CDate(mskFechaPresupuesto.Text) > fdtmServerFecha Then
        fblnDatosValidos = False
        '¡La fecha debe ser menor o igual a la del sistema!
        MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
        mskFechaPresupuesto.SetFocus
    End If
    If fblnDatosValidos And Val(txtDiasVencimiento.Text) = 0 Then
        fblnDatosValidos = False
        '¡Dato no válido!
        MsgBox "¡Dato no válido!, debe ser mayor de cero", vbOKOnly + vbInformation, "Mensaje"
        txtDiasVencimiento.SetFocus
    End If

    If fblnDatosValidos Then
        For X = 1 To grdPresupuesto(0).Rows - 1
            If Val(Format(grdPresupuesto(0).TextMatrix(X, ColPreciogrdPresupuesto), "###########.00")) = 0 Then
                fblnDatosValidos = False
             
                'El elemento seleccionado no cuenta con un precio capturado.
                MsgBox Replace(SIHOMsg(301), "El elemento seleccionado no cuenta", "Existen elementos que no cuentan"), vbOKOnly + vbExclamation, "Mensaje"
               
                Exit For
            End If
        Next
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlintcontador As Integer
    Dim vlstrNombreHospital As String
    Dim vlstrRegistro As String
    Dim vlstrDireccionHospital As String
    Dim vlstrTelefonoHospital As String
    Dim alstrParametros(8) As String
    Dim alstrParametros1(8) As String
    Dim rsReporte As New ADODB.Recordset
    
    ' Validación de cuando esta vacia la cotizacion
    If CLng(grdPresupuesto(0).TextMatrix(1, ColRowDatagrdPresupuesto)) = -1 Then Exit Sub
    
    '-------Borrar todo ---------------------------------
    vlstrSentencia = "Delete from PVImprimePresupuesto"
    pEjecutaSentencia vlstrSentencia
    
    '-------Traer datos generales del Hospital-----------
  
    vlstrNombreHospital = IIf(IsNull(vgstrNombreHospitalCH), "", Trim(vgstrNombreHospitalCH))
    vlstrRegistro = IIf(IsNull(vgstrRfCCH), "", "R.SSA " & RTrim(vgstrSSACH) & " RFC " & Trim(vgstrRfCCH))
    vlstrDireccionHospital = IIf(IsNull(vgstrDireccionCH), "", Trim(vgstrDireccionCH) & " CP " & Trim(vgstrCodPostalCH))
    vlstrTelefonoHospital = IIf(IsNull(vgstrTelefonoCH), "", Trim(vgstrTelefonoCH))
  
    
    '-------Crear el RS para grabar los datos -----------
   
    vlstrSentencia = "Select * from PVImprimePresupuesto"
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    With rs
        For vlintcontador = 1 To grdPresupuesto(0).Rows - 1
            .AddNew
            If cboEmpresa(0).Enabled Then
                !chrEmpresa = cboEmpresa(0).List(cboEmpresa(0).ListIndex)
            Else
                !chrEmpresa = cboTipoPaciente(0).List(cboTipoPaciente(0).ListIndex)
            End If
            !CHRNOMBRE = IIf(txtNombre.Text = "", " ", txtNombre.Text)
            !chrDireccion = IIf(txtDireccion.Text = "", " ", txtDireccion.Text)
            !VCHPROCEDIMIENTO = IIf(cboProcedimiento.Text = "", " ", Left(cboProcedimiento.Text, 50))
            !chrCargo = Trim(grdPresupuesto(0).TextMatrix(vlintcontador, ColDescripciongrdPresupuesto))
            !mnyPrecio = CDec(grdPresupuesto(0).TextMatrix(vlintcontador, ColPreciogrdPresupuesto))
            !MNYCantidad = grdPresupuesto(0).TextMatrix(vlintcontador, ColCantidadgrdPresupuesto)
            !MNYSUBTOTAL = CDec(grdPresupuesto(0).TextMatrix(vlintcontador, ColSubtotalgrdPresupuesto))
            !MNYDESCUENTO = "-" & CDec(grdPresupuesto(0).TextMatrix(vlintcontador, ColDescuentogrdPresupuesto))
            !mnyMonto = CDec(grdPresupuesto(0).TextMatrix(vlintcontador, ColMontogrdPresupuesto))
            '!mnyTImporte = CDec(txtSubtotal.Text)
            !mnyTSubtotal = CDec(txtImporte.Text)
            !mnyTDescuento = CDec(txtDescuentos.Text)
            !mnyTIva = CDec(txtIva.Text)
            !mnyTTotal = CDec(txtTotal.Text)
            'En este campose guarda un consecutivo para controlar el orden con que debe ser impreso
            !intConsecutivo = vlintcontador
            !CHRCVECARGO = Trim(grdPresupuesto(0).TextMatrix(vlintcontador, ColRowDatagrdPresupuesto))
            !chrTipoCargo = Trim(grdPresupuesto(0).TextMatrix(vlintcontador, ColTipogrdPresupuesto))
            'Se agregó la columna para el caso 20434
            !mnycosto = CDec(grdPresupuesto(0).TextMatrix(vlintcontador, ColCostogrdPresupuesto))
            
            .Update
        Next
    End With
    
    rs.Close
    
    If optCarta.Value Then
    
      Set rsReporte = frsEjecuta_SP("", "SP_PVSELIMPRIMEPRESUPUESTO")
    
      If rsReporte.RecordCount > 0 Then
        pInstanciaReporte vgrptReporte, "rptPresupuestoCarta.rpt"
        vgrptReporte.DiscardSavedData
    
        alstrParametros(0) = "Mensaje;" & IIf(txtMensaje.Text = "", " ", txtMensaje.Text)
        alstrParametros(1) = "NombreEmpresa;" & Trim(vlstrNombreHospital)
        alstrParametros(2) = "NombreReporte;" & "PRESUPUESTO" & " " & Trim(cboProcedimiento.Text)
        alstrParametros(3) = "Nota;" & IIf(txtNotas.Text = "", " ", txtNotas.Text)
        alstrParametros(4) = "Estado;" & txtEstado.Text
        alstrParametros(5) = "Vigencia;" & Format(CDate(mskFechaPresupuesto) + Val(txtDiasVencimiento.Text), "dd/mmm/yyyy")
        alstrParametros(6) = "FechaPresupuesto;" & Format(CDate(mskFechaPresupuesto), "dd/mmm/yyyy")
        alstrParametros(7) = "NumeroPresupuesto;" & txtClave.Text
        If Not fblnRevisaPermiso(vglngNumeroLogin, 7037, "C", True) Then
            'El usuario no tiene permiso para ver las columnas
            alstrParametros(8) = "Mostrar;" & "S"
        Else
            alstrParametros(8) = "Mostrar;" & "C"
        End If
        
        pCargaParameterFields alstrParametros, vgrptReporte
      
        pImprimeReporte vgrptReporte, rsReporte, "P", "Carta del presupuesto"
      Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
      End If
      If rsReporte.State <> adStateClosed Then rsReporte.Close
        
    Else
      Set rsReporte = frsRegresaRs("SELECT * FROM PVIMPRIMEPRESUPUESTO order by chrcargo")
    
      If rsReporte.RecordCount > 0 Then
        pInstanciaReporte vgrptReporte, "rptPresupuestoTicket.rpt"
        vgrptReporte.DiscardSavedData
    
        alstrParametros1(0) = "NombreEmpresa;" & Trim(vlstrNombreHospital)
        alstrParametros1(1) = "Direccion;" & vlstrDireccionHospital
        alstrParametros1(2) = "Telefono;" & "TEL. " & Format(RTrim(vlstrTelefonoHospital), "###-##-##")
        alstrParametros1(3) = "Registro;" & vlstrRegistro
        alstrParametros1(4) = "FechaActual;" & UCase(Format(fdtmServerFecha, "dd/mmm/yyyy"))
        alstrParametros1(5) = "NombreReporte;" & "PRESUPUESTO" & " " & Trim(cboProcedimiento.Text)
        alstrParametros1(6) = "Nota;" & txtNotas.Text
        alstrParametros1(7) = "Estado;" & txtEstado.Text
        alstrParametros1(8) = "Vigencia;" & CDate(mskFechaPresupuesto) + Val(txtDiasVencimiento.Text)
        
        pCargaParameterFields alstrParametros1, vgrptReporte
      
        pImprimeReporte vgrptReporte, rsReporte, "P", "Ticket del presupuesto"
      Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
      End If
      If rsReporte.State <> adStateClosed Then rsReporte.Close
        
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub

Private Sub pNuevo()
    On Error GoTo NotificaError

    vgstrEstadoManto = ""
    
    pLimpiaGrid grdPresupuesto(0)
    pLimpiaGrid grdElementos
    pConfiguraGridCargos 0
    pHabilitaBotones 0, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0
    
    rsPresupuesto.Requery
    
    mskFechaPresupuesto.Mask = ""
    mskFechaPresupuesto.Text = fdtmServerFecha
    mskFechaPresupuesto.Mask = "##/##/####"
    mskFechaPresupuesto.Enabled = False
    
    txtNumCuenta.Text = ""
    txtDiasVencimiento.Text = 0
    txtDiasVencimiento.Enabled = False
    txtEstado.Text = "CREADO"
    txtImporte.Text = ""
    txtSubtotal.Text = ""
    txtDescuentos.Text = ""
    txtIva.Text = ""
    txtTotal.Text = ""
    txtNombre.Text = ""
    txtDireccion.Text = ""
    txtSeleArticulo = ""
    txtPorcentajeDescuento(0).Text = ""
    
    chkTomarDescuentos(0).Value = 1
    
    optDescripcion.Value = True
    optTipoPaciente(vgintTipoPaciente).Value = True
    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    
    cboEmpresa(0).ListIndex = -1
    cboTipoPaciente(0).ListIndex = -1
    cboProcedimiento.ListIndex = -1
    cboConceptoFactura.ListIndex = -1
    cboTratamiento.ListIndex = -1
    txtNumPaquete.Text = ""
    txtMargenUtilidadTotal = ""
    
    'pEnfocaTextBox txtClave
    'lblMotivoNoAutorizado.Visible = False
    txtMotivoNoAutorizado.Enabled = False
    txtMotivoNoAutorizado.Text = ""
'    lblConceptoFactura.Visible = False
'    cboConceptoFactura.Visible = False
'    lblTratamiento.Visible = False
'    cboTratamiento.Visible = False
'    lblNumPaquete.Visible = False
'    txtNumPaquete.Visible = False
'    fraDatosPaquete.Visible = False
    pHabilitaDatosPaquete False, False, True
    
    pHabilitaFrames False
    freBotones.Enabled = True
    
    vgblnEsNuevo = True 'Se agrego para saber que es nuevo
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pNuevo"))
End Sub

Private Sub cmdNoAutorizar_Click()

    pHabilitaBotones 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1
    If txtMotivoNoAutorizado.Enabled Then
        If Trim(txtMotivoNoAutorizado.Text) = "" Then
            MsgBox "No ha ingresado el motivo por el cual se cambia el estado del presupuesto a no autorizado.", vbOKOnly + vbInformation, "Mensaje"
            If fblnCanFocus(txtMotivoNoAutorizado) Then txtMotivoNoAutorizado.SetFocus
        Else
            pGrabaNoAutorizado
            vgstrEstadoManto = "C"
        End If
    Else
        vgstrEstadoManto = "NA"
        txtMotivoNoAutorizado.Text = ""
        fraMotivoNoAutorizado.Enabled = True
        txtMotivoNoAutorizado.Enabled = True
        pEnfocaTextBox txtMotivoNoAutorizado
    End If
End Sub

Private Sub cmdParametros_Click()
    On Error GoTo NotificaError
    
    vgbitParametros = True
    pHabilitaFrames False
    freParametros.Top = 875
    freParametros.Visible = True
    freDatos.Enabled = False
    freBotones.Enabled = False
    pEnfocaTextBox txtMensaje
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdParametros_Click"))
End Sub

Private Sub pHabilitaFrames(vlblnEstado As Boolean)
    On Error GoTo NotificaError

    FreElementos.Enabled = vlblnEstado
    fraElementoFijo.Enabled = vlblnEstado
    FreDetalle.Enabled = vlblnEstado
    FreIncluir.Enabled = vlblnEstado
    freDatos2.Enabled = vlblnEstado
    fraAccionPresupuesto.Enabled = vlblnEstado
    freBotones.Enabled = vlblnEstado
    freVarios.Enabled = vlblnEstado
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaFrames"))
End Sub

Private Sub cmdSalirDuplicado_Click()
On Error GoTo NotificaError
    pHabilitaFrames True
    Frame3.Enabled = True
    cmdActualizaPrecios(0).Enabled = True
    FraDuplicar.Visible = False
    vgstrEstadoManto = "C"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDuplicar_Click"))
End Sub

Private Sub cmdSelecciona_Click(Index As Integer)
On Error GoTo NotificaError

    If Index = 0 Then
        pSeleccionaElemento 1
    Else
        grdPresupuesto_DblClick 0
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSelecciona_Click"))
End Sub

Private Sub cmdVerPagos_Click()
On Error GoTo NotificaError
    
    Dim vllngContador As Long
    Dim vllngContador2 As Long
    Dim intSeleccionados As Integer
    Dim blnSiHaySinSeleccion As Boolean
    
    pVerPagos "P"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If vgbitParametros Then
        Cancel = 1
        If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            vgbitParametros = False
            freParametros.Visible = False
            freDatos.Enabled = True
            freBotones.Enabled = True
            If txtEstado.Text = "CREADO" Then
                pHabilitaFrames True
            Else
                pHabilitaFrames False
            End If
        End If
    Else
        If vgstrEstadoManto <> "E" Then
            If vgstrEstadoManto = "C" Then
                Cancel = 1
                If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pEnfocaTextBox txtClave
                End If
            Else
                If vgstrEstadoManto = "BC" Then
                    Cancel = 1
                    vgstrEstadoManto = "C"
                    freDatos.Enabled = True
                    freBotones.Enabled = True
                    freBusqueda.Visible = False
                    If txtEstado.Text = "CREADO" Then
                        pHabilitaFrames True
                    Else
                        pHabilitaFrames False
                    End If
                    pEnfocaTextBox txtNombre
                End If
                If vgstrEstadoManto = "PA" Then
                    Cancel = 1
                    vgstrEstadoManto = "C"
                    pHabilitaDatosPaquete False, False
                    If fblnCanFocus(cmdImprimir) Then cmdImprimir.SetFocus
                    pHabilitaDatosPaquete False, False
                End If
                If vgstrEstadoManto = "NA" Then
                    Cancel = 1
                    vgstrEstadoManto = "C"
                    pHabilitaBotones 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1
                    If fblnCanFocus(cmdImprimir) Then cmdImprimir.SetFocus
                    txtMotivoNoAutorizado.Text = ""
                    fraMotivoNoAutorizado.Enabled = False
                End If
                If vgstrEstadoManto = "DU" Then
                    Cancel = 1
                    vgstrEstadoManto = "C"
                    cmdSalirDuplicado_Click
                    pEnfocaTextBox txtNombre
                End If
            End If
        Else
            If FraDuplicar.Visible Then
                vgstrEstadoManto = "DU"
                Cancel = 1
            Else
                vgstrEstadoManto = "C"
                Cancel = 1
            End If
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub grdElementoFijo_DblClick()
    On Error GoTo NotificaError

    If Trim(grdElementoFijo.TextMatrix(1, 1)) <> "" Then
        pAgregaElementoFijo
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdElementoFijo_DblClick"))
End Sub

Private Sub grdElementoFijo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If Trim(grdElementoFijo.TextMatrix(1, 1)) <> "" Then
            pAgregaElementoFijo
        End If
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdElementoFijo_KeyPress"))
End Sub

Private Sub pAgregaElementoFijo()
    On Error GoTo NotificaError

    Dim vllngRenglon As Long
    
    If Trim(grdPresupuesto(0).TextMatrix(1, ColDescripciongrdPresupuesto)) = "" Then
        vllngRenglon = 1
    Else
        grdPresupuesto(0).Rows = grdPresupuesto(0).Rows + 1
        vllngRenglon = grdPresupuesto(0).Rows - 1
    End If
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColDescripciongrdPresupuesto) = grdElementoFijo.TextMatrix(grdElementoFijo.Row, 1)
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColPreciogrdPresupuesto) = FormatCurrency("0")
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColCantidadgrdPresupuesto) = 1
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColSubtotalgrdPresupuesto) = FormatCurrency("0")
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColDescuentogrdPresupuesto) = FormatCurrency("0")
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColMontogrdPresupuesto) = FormatCurrency("0")
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColTipogrdPresupuesto) = "EF"
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColIVAgrdPresupuesto) = Val(grdElementoFijo.TextMatrix(grdElementoFijo.Row, 3))
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColPorcentajeDescgrdPresupuesto) = "0"
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColContenidogrdPresupuesto) = 1 'Contenido de IVArticulo, Oviamente nomas para articulos jala
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColModoDescuentogrdPresupuesto) = 0 'Tipo de descuento de inventario
    grdPresupuesto(0).TextMatrix(vllngRenglon, ColRowDatagrdPresupuesto) = Val(grdElementoFijo.TextMatrix(grdElementoFijo.Row, 2))

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAgregaElementoFijo"))
End Sub

Private Sub grdElementos_DblClick()
    On Error GoTo NotificaError

    pSeleccionaElemento 1
    
    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdElementos_DblClick"))
End Sub

Private Sub grdElementos_EnterCell()
    txtNombreComercial.Text = Trim(grdElementos.TextMatrix(grdElementos.Row, 1))
End Sub

Private Sub grdElementos_GotFocus()
    txtNombreComercial.Text = Trim(grdElementos.TextMatrix(grdElementos.Row, 1))
End Sub

Private Sub grdElementos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = vbKeyReturn Then
        grdElementos_DblClick
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdElementos_KeyDown"))
End Sub

Private Sub grdElementos_LostFocus()
    txtNombreComercial.Text = ""
End Sub

Private Sub grdPresupuesto_Click(Index As Integer)
On Error GoTo NotificaError

    Dim vldblValor As Double
    
    If CLng(IIf(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColRowDatagrdPresupuesto) = "", "-1", grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColRowDatagrdPresupuesto))) = -1 Then Exit Sub
        
    If Trim(grdPresupuesto(Index).TextMatrix(1, ColDescripciongrdPresupuesto)) <> "" Then
        If ( _
            grdPresupuesto(Index).Col = vgintColumnaCurrency And _
            Trim(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColTipogrdPresupuesto)) <> "EF" _
            ) Or _
            ( _
            grdPresupuesto(Index).Col = ColPreciogrdPresupuesto And _
            Trim(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColTipogrdPresupuesto)) = "EF" _
            ) Or _
            ( _
            grdPresupuesto(Index).Col = ColDescuentogrdPresupuesto And _
            Val(Format(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColSubtotalgrdPresupuesto), "############.00")) <> 0 _
            ) Or _
            ( _
            grdPresupuesto(Index).Col = ColCantidadgrdPresupuesto _
            ) _
            Then
                        
            If grdPresupuesto(Index).Col = ColDescuentogrdPresupuesto Then
                vldblValor = Round(Val(Format(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, grdPresupuesto(Index).Col), "############.00")) / Val(Format(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColSubtotalgrdPresupuesto), "############.00")), 2) * 100
            Else
                vldblValor = Val(Format(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, grdPresupuesto(Index).Col), "############.00"))
'                If Val(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColCostogrdPresupuesto)) = 0 Then
'                    MsgBox "El elemento seleccionado no tiene costo.", vbOKOnly + vbInformation, "Mensaje"
'                    grdPresupuesto(Index).RemoveItem (grdPresupuesto(Index).Row)
'                    Exit Sub
'                End If
            End If
                            
            pEditarColumna vldblValor, txtCantidad(Index), grdPresupuesto(Index)
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPresupuesto_Click"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vgstrEstadoManto = ""
    'txtSubtotal.Text = Format(0, "$###,###,###,##0.00")
    vlTipoPacienteSocio = flngTipoPacienteSocio 'Verifica si hay tipos de socio configurados
    vlblnCargando = False
    
    pInstanciaReporte vgrptReporte, "rptPresupuestoCarta.rpt"
    pCargaElementoFijo
    pCargaProcedimientos
    pCargaEmpresas
    pCargaTiposPaciente
    'pConfiguraGridCargos 0
    
    'pCargaElementos
    
    vlstrSentencia = "Select * from pvMensajePresupuesto"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        txtMensaje.Text = Trim(rs!chrMensaje)
        txtNotas.Text = Trim(rs!chrNotas)
    End If
    rs.Close
    
    'El RS general
    vlstrSentencia = "select intCvePresupuesto from pvPresupuesto where smiCveDepartamento = " & vgintNumeroDepartamento & _
         " order by intCvePresupuesto"
    Set rsPresupuesto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenDynamic)
    
    frmPresupuestos.Refresh
    
    'optTipoPaciente(0).Value = True
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Presupuesto) > 0 Then
        If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Presupuesto) = 1 Then
           vgintTipoPaciente = 1
        ElseIf fintEsInterno(vglngNumeroLogin, enmTipoProceso.Presupuesto) = 2 Then
            vgintTipoPaciente = 0
        Else
            vgintTipoPaciente = 2
        End If
    Else
        vgintTipoPaciente = 1
    End If

    pConceptosFactura
    
    ' Para ver cual es el Tipo de Paciente Particular
    vlstrSentencia = 0
    Set rs = frsSelParametros("SI", -1, "INTTIPOPARTICULAR")
    If rs.RecordCount > 0 Then vglngTipoParticular = IIf(IsNull(rs!Valor), 0, rs!Valor)
    rs.Close
    
    pEnfocaTextBox txtClave
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pCargaElementoFijo()
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset

    With grdElementoFijo
        .Rows = 2
        .Cols = 4
    End With

    Set rs = frsRegresaRs("select vchDescripcion Descripcion, intCveElemento Clave, smyIva/100 IVA From PvPresupuestoElementoFijo Where bitActivo = 1")
    
    If rs.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdElementoFijo, rs
    End If
    
    With grdElementoFijo
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción"
        .ColWidth(0) = 100
        .ColWidth(1) = 4900
        .ColWidth(2) = 0
        .ColWidth(3) = 0
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaElementoFijo"))
End Sub

Private Sub cmdAnteriorRegistro_Click()
    On Error GoTo NotificaError

    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    cboEmpresa(0).Enabled = False
    cboTipoPaciente(0).Enabled = False


    txtSeleArticulo = ""
    pCargaElementos

    Call pPosicionaRegRs(rsPresupuesto, "A")
    pModificaRegistro
    Exit Sub
        
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAnteriorRegistro_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    On Error GoTo NotificaError

    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    cboEmpresa(0).Enabled = False
    cboTipoPaciente(0).Enabled = False


    txtSeleArticulo = ""
    pCargaElementos

    Call pPosicionaRegRs(rsPresupuesto, "I")
    pModificaRegistro
    Exit Sub
        
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrimerRegistro_Click"))
End Sub

Private Sub cmdSiguienteRegistro_Click()
    On Error GoTo NotificaError

    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    cboEmpresa(0).Enabled = False
    cboTipoPaciente(0).Enabled = False

    txtSeleArticulo = ""
    pCargaElementos

    Call pPosicionaRegRs(rsPresupuesto, "S")
    pModificaRegistro
    Exit Sub
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSiguienteRegistro_Click"))
End Sub

Private Sub cmdUltimoRegistro_Click()
    On Error GoTo NotificaError
    
    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    cboEmpresa(0).Enabled = False
    cboTipoPaciente(0).Enabled = False

    txtSeleArticulo = ""
    pCargaElementos

    Call pPosicionaRegRs(rsPresupuesto, "U")
    pModificaRegistro
    Exit Sub
        
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdUltimoRegistro_Click"))
End Sub

Private Sub pModificaRegistro()
    On Error GoTo NotificaError

    With rsPresupuesto
        If Not .EOF And Not .BOF Then
            txtClave.Text = rsPresupuesto!INTCVEPRESUPUESTO
            txtClave_KeyDown vbKeyReturn, 0
        End If
    End With
                
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pModificaRegistro"))
End Sub

Private Function fintLocalizaEnLista(lstLista As ListBox, intClave As Integer) As Integer
    On Error GoTo NotificaError
    Dim vlintcontador As Integer
    fintLocalizaEnLista = -1   'Regresa un -1 si no lo encuentra
    For vlintcontador = 0 To lstLista.ListCount
        If lstLista.ItemData(vlintcontador) = intClave Then
            fintLocalizaEnLista = vlintcontador
            vlintcontador = lstLista.ListCount + 1
        End If
    Next vlintcontador
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintLocalizaEnLista"))
End Function



Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = vbKeyEscape Then
        txtCantidad(0).Visible = False
        txtCantidad(1).Visible = False
        KeyAscii = 0
        Unload Me
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub grdPresupuesto_DblClick(Index As Integer)
On Error GoTo NotificaError

    With grdPresupuesto(Index)
        If .Rows > 2 Then
            If CLng(IIf(.TextMatrix(.Row, ColRowDatagrdPresupuesto) = "", "-1", .TextMatrix(.Row, ColRowDatagrdPresupuesto))) > -1 Then
                pBorrarRegMshFGrdData grdPresupuesto(Index).Row, grdPresupuesto(Index)
            End If
        Else
            pLimpiaMshFGrid grdPresupuesto(Index)
            .Rows = 2
            pConfiguraGridCargos Index
        End If
        pRecalcula Index
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPresupuesto_DblClick"))
End Sub

Private Sub grdPresupuesto_EnterCell(Index As Integer)
    txtNombreComercial.Text = Trim(grdPresupuesto(0).TextMatrix(grdPresupuesto(0).Row, 1))
End Sub

Private Sub grdPresupuesto_GotFocus(Index As Integer)
    txtNombreComercial.Text = Trim(grdPresupuesto(0).TextMatrix(grdPresupuesto(0).Row, 1))
End Sub

Private Sub grdPresupuesto_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    If grdPresupuesto(Index).Col = vgintColumnaCurrency Then
        If KeyCode = vbKeyF2 Then 'para que se edite el contenido de la celda como en excel
            Call pEditarColumna(13, txtCantidad(Index), grdPresupuesto(Index))
        End If
    Else
        If KeyCode = vbKeyReturn Then
            If txtCantidad(Index).Visible Then
                grdPresupuesto(Index).Col = 0
                grdPresupuesto(Index).CellFontBold = True
                grdPresupuesto(Index).Col = 1
                If grdPresupuesto(Index).Row - 1 < grdPresupuesto(Index).Rows Then
                    If grdPresupuesto(Index).Row = grdPresupuesto(Index).Rows - 1 Then
                        grdPresupuesto(Index).Row = 1
                    Else
                        grdPresupuesto(Index).Row = grdPresupuesto(Index).Row + 1
                        If grdPresupuesto(Index).Row = grdPresupuesto(Index).Rows - 1 Then
                            grdPresupuesto(Index).Row = 1
                        Else
                            grdPresupuesto(Index).Row = grdPresupuesto(Index).Row + 1
                        End If
                    End If
                End If
            Else
                grdPresupuesto_Click (Index)
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPresupuesto_KeyDown"))
End Sub

Private Sub grdPresupuesto_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError

    If Trim(grdPresupuesto(Index).TextMatrix(1, ColDescripciongrdPresupuesto)) <> "" Then
        If ( _
            grdPresupuesto(Index).Col = vgintColumnaCurrency And _
            Trim(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColTipogrdPresupuesto)) <> "EF" _
            ) Or _
            ( _
            grdPresupuesto(Index).Col = ColPreciogrdPresupuesto And _
            Trim(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColTipogrdPresupuesto)) = "EF" _
            ) Or _
            ( _
            grdPresupuesto(Index).Col = ColDescuentogrdPresupuesto And _
            Val(Format(grdPresupuesto(Index).TextMatrix(grdPresupuesto(Index).Row, ColSubtotalgrdPresupuesto), "############.00")) <> 0 _
            ) _
            Then
            
            If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
                KeyAscii = 7
            Else
                If IsNumeric(Chr(KeyAscii)) Then
                    pEditarColumna Val(Chr(KeyAscii)), txtCantidad(Index), grdPresupuesto(Index)
                End If
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPresupuesto_KeyPress"))
End Sub

Private Sub grdPresupuesto_LostFocus(Index As Integer)
    txtNombreComercial.Text = ""
End Sub

Private Sub grdPresupuesto_Scroll(Index As Integer)
    vgstrEstadoManto = "C"
    txtCantidad(Index).Visible = False
End Sub

Private Sub lstBuscaPresupuesto_DblClick()
    On Error GoTo NotificaError
    vgblnEsNuevo = False 'Se agrego para saber que no es nuevo
    
    cmdAceptoBuscar_Click
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstBuscaPresupuesto_DblClick"))
End Sub

Private Sub lstBuscaPresupuesto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cmdAceptoBuscar_Click
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstBuscaPresupuesto_KeyDown"))
End Sub

Private Sub mskFechaPresupuesto_GotFocus()
On Error GoTo NotificaError

    pSelMkTexto mskFechaPresupuesto
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaPresupuesto_GotFocus"))
End Sub

Private Sub mskFechaPresupuesto_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        txtDiasVencimiento.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaPresupuesto_KeyPress"))
End Sub

Private Sub optClave_Click()
    On Error GoTo NotificaError

    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 11
    txtSeleArticulo.SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optClave_Click"))
End Sub

Private Sub optClave_KeyDown(KeyCode As Integer, Shift As Integer)
    optClave_Click
End Sub


Private Sub optCodigoBarras_Click()
    On Error GoTo NotificaError

    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 30
    txtSeleArticulo.SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optCodigoBarras_Click"))
End Sub

Private Sub optCodigoBarras_KeyDown(KeyCode As Integer, Shift As Integer)
    optCodigoBarras_Click
End Sub


Private Sub optDescripcion_Click()
    On Error GoTo NotificaError

    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 30
    txtSeleArticulo.SetFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optDescripcion_Click"))
End Sub

Private Sub optDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    optDescripcion_Click
End Sub


Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If txtClave <> "" Then
            txtNombre.SetFocus
        Else
            pEnfocaTextBox txtClave
        End If
    End If
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub


Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoPaciente_KeyPress"))

End Sub


Private Sub txtBuscaNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstBuscaPresupuesto.Enabled Then
        lstBuscaPresupuesto.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBuscaNombre_KeyDown"))
End Sub

Private Sub txtBuscaNombre_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBuscaNombre_KeyPress"))
End Sub

Private Sub txtBuscaNombre_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vlstrSentencia2 As String
    vlstrSentencia = "select * from pvPresupuesto "
    vlstrSentencia2 = "and smicvedepartamento = " & vgintNumeroDepartamento
    PSuperBusqueda txtBuscaNombre, vlstrSentencia, lstBuscaPresupuesto, "chrNombre", 100, vlstrSentencia2
'    PSuperBusqueda txtBuscaNombre, vlstrSentencia, lstBuscaPresupuesto, "intCvePresupuesto", 100, vlstrSentencia2
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtBuscaNombre_KeyUp"))
End Sub

Private Sub txtCantidad_LostFocus(Index As Integer)
    On Error GoTo NotificaError

    vgstrEstadoManto = "C"
    txtCantidad(Index).Visible = False
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidad_LostFocus"))
End Sub


Private Sub txtClave_LostFocus()

'    Dim vlstrSentencia As String
'    Dim rs As New ADODB.Recordset
'    Dim vllngNada As Long
'
'        If txtClave <> "" Then
'            vlstrSentencia = "select * from pvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text) & _
'                   " and smiCveDepartamento = " & vgintNumeroDepartamento
'            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'            If rs.RecordCount > 0 Then
'
'                optTipoPaciente(0).Enabled = False
'                optTipoPaciente(1).Enabled = False
'                optTipoPaciente(2).Enabled = False
'
'                'vlblnCargando = True
'                If rs!bitAplicarDescuento = 1 Then
'                    Me.chkTomarDescuentos(0).Value = vbUnchecked
'                    Me.txtPorcentajeDescuento(0).Text = IIf(IsNull(rs!mnyAplicarDescuento), "", rs!mnyAplicarDescuento)
'                Else
'                    Me.txtPorcentajeDescuento(0).Text = 0
'                    Me.chkTomarDescuentos(0).Value = vbChecked
'                End If
'                If rs!CHRTIPOPACIENTE = "I" Then
'                    Me.optTipoPaciente(1).Value = True
'                ElseIf rs!CHRTIPOPACIENTE = "E" Then
'                    Me.optTipoPaciente(0).Value = True
'                Else
'                    Me.optTipoPaciente(2).Value = True
'                End If
'                'vlblnCargando = False
'                txtNombre.Text = Trim(rs!CHRNOMBRE)
'                txtDireccion.Text = Trim(rs!chrDireccion)
'                'txtProcedimiento.Text = IIf(IsNull(rs!VCHPROCEDIMIENTO), "", rs!VCHPROCEDIMIENTO)
'                cboTipoPaciente(0).ListIndex = fintLocalizaCbo(cboTipoPaciente(0), rs!intTipoPaciente)
'                If rs!intEmpresa = 0 Then
'                    cboEmpresa.ListIndex = 0
'                Else
'                    cboEmpresa.ListIndex = fintLocalizaCbo(cboEmpresa, rs!intEmpresa)
'                End If
'
'                cmdDelete.Enabled = True
'                cmdImprimir.Enabled = True
'                vllngNada = fintLocalizaPkRs(rsPresupuesto, 0, txtClave.Text)
'
'            ElseIf CLng(Val(txtClave.Text)) >= vglngClaveSiguiente Then
'                txtClave.Text = vglngClaveSiguiente
'            Else
'                MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
'                pNuevo
'                Exit Sub
'            End If
'            pCargaGridConsulta
'            rs.Close
'            vgstrEstadoManto = "C" 'Estatus de Cambios
'            pEnfocaTextBox txtNombre

'            If optTipoPaciente(0).Value = True Then
'                optTipoPaciente(0).SetFocus
'            ElseIf optTipoPaciente(1).Value = True Then
'                optTipoPaciente(1).SetFocus
'            ElseIf optTipoPaciente(2).Value = True Then
'                optTipoPaciente(2).SetFocus
'            End If

'        Else
'            pEnfocaTextBox txtClave
'        End If
        
        
End Sub

Private Sub txtDiasVencimiento_GotFocus()
On Error GoTo NotificaError
    pSelTextBox txtDiasVencimiento
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDiasVencimiento_GotFocus"))
End Sub

Private Sub txtDiasVencimiento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

    If KeyAscii = 13 Then
        If optTipoPaciente(0).Value Then
            If fblnCanFocus(optTipoPaciente(0)) Then optTipoPaciente(0).SetFocus
        ElseIf optTipoPaciente(1).Value Then
            If fblnCanFocus(optTipoPaciente(1)) Then optTipoPaciente(1).SetFocus
        Else
            If fblnCanFocus(optTipoPaciente(1)) Then optTipoPaciente(1).SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDiasVencimiento_KeyPress"))
End Sub

Private Sub txtDiasVencimiento_LostFocus()
    If txtDiasVencimiento.Text = "" Then txtDiasVencimiento.Text = 0
End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

End Sub



Private Sub txtMotivoNoAutorizado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        'cmdAceptarNoAutorizado.SetFocus
        cmdNoAutorizar.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMotivoNoAutorizado_KeyDown"))
End Sub


Private Sub txtMotivoNoAutorizado_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMotivoNoAutorizado_KeyPress"))
End Sub


Private Sub txtNotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
End Sub

Private Sub txtPorcentajeDescuento_GotFocus(Index As Integer)
On Error GoTo NotificaError

    pSelTextBox txtPorcentajeDescuento(Index)
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorcentajeDescuento_GotFocus"))
End Sub

Private Sub txtPorcentajeDescuento_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = Asc(".") And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            cmdAplicaDescuento(Index).SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPorcentajeDescuento_KeyPress"))
End Sub

Private Sub txtSeleArticulo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If optClave.Value Then
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtSeleArticulo_KeyPress"))
End Sub

Private Sub txtClave_GotFocus()
    On Error GoTo NotificaError

'    Dim vlstrSentencia As String
'    Dim rs As New ADODB.Recordset
    
    If vgstrEstadoManto <> "DU" Then
        vglngClaveSiguiente = frsRegresaRs("SELECT LAST_NUMBER MAXIMO FROM USER_SEQUENCES WHERE SEQUENCE_NAME = 'SEC_PVPRESUPUESTO'").Fields(0)
        txtClave.Text = vglngClaveSiguiente
        pSelTextBox txtClave
        pNuevo
    End If
    
'    If IsNull(rs!Maximo) Then
'        vglngClaveSiguiente = 1
'    Else
'        vglngClaveSiguiente = rs!Maximo + 1
'    End If
'    txtClave.Text = vglngClaveSiguiente
'    rs.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_GotFocus"))
End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vllngNada As Long

    If KeyCode = vbKeyReturn Then
        If txtClave <> "" Then
            vlstrSentencia = "select pvPresupuesto.*, to_char(DTMFECHA, 'DD/MM/YYYY') as fecha, to_char(DTMFECHAPRESUPUESTO, 'DD/MM/YYYY') as fechaPresupuesto from pvPresupuesto where intCvePresupuesto = " & Trim(txtClave.Text) & _
                   " and smiCveDepartamento = " & vgintNumeroDepartamento
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rs.RecordCount > 0 Then
                pConsulta rs
                pCargaGridConsulta
            ElseIf CLng(Val(txtClave.Text)) >= vglngClaveSiguiente Then
                txtClave.Text = vglngClaveSiguiente
                pHabilitaFrames True
                pHabilitaBotones 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0
                optTipoPaciente(0).Enabled = True
                optTipoPaciente(1).Enabled = True
                optTipoPaciente(2).Enabled = True
                txtDiasVencimiento.Enabled = True
                mskFechaPresupuesto.Enabled = True
                cboTipoPaciente(0).Enabled = True
                vgblnEsNuevo = True
            Else
                MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
                'pNuevo
                Exit Sub
            End If
            rs.Close
            If vgstrEstadoManto <> "PA" Then vgstrEstadoManto = "C" 'Estatus de Cambios
            
            If mskFechaPresupuesto.Enabled Then
                mskFechaPresupuesto.SetFocus
            Else
                If fblnCanFocus(cmdImprimir) Then cmdImprimir.SetFocus
            End If
        Else
            pEnfocaTextBox txtClave
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyDown"))
End Sub

Private Sub pCargaGridConsulta()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP(txtClave.Text, "sp_PVSelPresupuesto")
    
    'Habilitar frames
    'pHabilitaFrames True
    
    'Limpieza del Grid y configuración del grid
    'pLimpiaGrid grdPresupuesto(0)
    'pConfiguraGridCargos 0
    '--------------------------------------------
    With grdPresupuesto(0)
            If Not fblnRevisaPermiso(vglngNumeroLogin, 7037, "C", True) Then
                'El usuario no tiene permiso para ver las columnas
                .ColWidth(ColMargenUtilidadgrdPresupuesto) = 0
                .ColWidth(ColPreciogrdPresupuesto) = 0
                .ColWidth(ColCostogrdPresupuesto) = 0
            Else
                'El usuario tiene permiso para ver las columnas
                .ColWidth(ColMargenUtilidadgrdPresupuesto) = 1200
                .ColWidth(ColPreciogrdPresupuesto) = 1200
                .ColWidth(ColCostogrdPresupuesto) = 1200
            End If
    End With
    'Cargada del Grid
    Do While Not rs.EOF
        With grdPresupuesto(0)
            If CLng(.TextMatrix(1, ColRowDatagrdPresupuesto)) <> -1 Then
                .Rows = .Rows + 1
            End If
            .Row = .Rows - 1
            .TextMatrix(.Row, ColRowDatagrdPresupuesto) = IIf(IsNull(rs!clave), 0, rs!clave)
            .TextMatrix(.Row, ColDescripciongrdPresupuesto) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
            .TextMatrix(.Row, ColMargenUtilidadgrdPresupuesto) = Format(rs!margenUtilidad, "#0.00") & "%"
            .TextMatrix(.Row, ColCostogrdPresupuesto) = Format(rs!costo, "$###,###,###,##0.00")
            .TextMatrix(.Row, ColPreciogrdPresupuesto) = Format(rs!precio, "$###,###,###,###.00")
            .TextMatrix(.Row, ColCantidadgrdPresupuesto) = rs!cantidad
            .TextMatrix(.Row, ColSubtotalgrdPresupuesto) = Format(rs!precio * CLng(rs!cantidad), "$###,###,###,###.00")
            .TextMatrix(.Row, ColDescuentogrdPresupuesto) = Format(rs!Descuento, "$###,###,###,###.00")
            .TextMatrix(.Row, ColMontogrdPresupuesto) = Format((rs!precio * CLng(rs!cantidad)) - rs!Descuento, "$###,###,###,###.00")
            .TextMatrix(.Row, ColTipogrdPresupuesto) = rs!tipo  'Tipo de cargo
            If rs!IVA > 0 Then
                .TextMatrix(.Row, ColIVAgrdPresupuesto) = rs!IVA / (rs!precio * CLng(rs!cantidad) - rs!Descuento)
            End If
            .TextMatrix(.Row, ColMontoIVAgrdPresupuesto) = rs!IVA
            .TextMatrix(.Row, ColPorcentajeDescgrdPresupuesto) = rs!Descuento / (rs!precio * CLng(rs!cantidad))
            .TextMatrix(.Row, ColContenidogrdPresupuesto) = IIf(IsNull(rs!Contenido), 0, rs!Contenido)
            .TextMatrix(.Row, ColModoDescuentogrdPresupuesto) = IIf(IsNull(rs!INTDESCUENTOINVENTARIO), 0, rs!INTDESCUENTOINVENTARIO)
        End With
        rs.MoveNext
    Loop
    rs.Close
    
    pRecalcula 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaGridConsulta"))
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtClave_KeyPress"))
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        cboProcedimiento.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDireccion_KeyDown"))
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDireccion_KeyPress"))
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtDireccion
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":TxtNombre_KeyDown"))
End Sub
Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":TxtNombre_KeyPress"))
End Sub

Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox, Index As Integer)
    On Error GoTo NotificaError

    Dim vldblDescuento As Double
    Dim vlstrCantidad As String
    
    If grid.Col = ColCantidadgrdPresupuesto Then 'Cantidad
        vlstrCantidad = grid.Text
        grid.Text = Format(txtCantidad(Index).Text, "######")
        
        If Val(Format(txtCantidad(Index).Text, "######")) = 0 Then
            With grid
                If .Rows > 2 Then
                    'pBorrarRegMshFGrdData grid.Row, grid
                    grid.Text = vlstrCantidad
                Else
                    'pLimpiaMshFGrid grid
                    '.Rows = 2
                    'pConfiguraGridCargos grid.Index
                    grid.Text = vlstrCantidad
                End If
            End With
        End If
    Else
        If grid.Col = ColPreciogrdPresupuesto Then 'Precio
            grid.Text = FormatCurrency(Val(txtCantidad(Index).Text), 2)
        Else
            If grid.Col = ColDescuentogrdPresupuesto Then 'Descuento
                vldblDescuento = Val(Format(grid.TextMatrix(grid.Row, ColSubtotalgrdPresupuesto), "########.##")) * Val(txtCantidad(Index).Text) / 100
                'Se valida que el descuento no sea mayor que 100%
                If Val(txtCantidad(Index).Text) >= 100 Then
                    MsgBox SIHOMsg(35), vbCritical, "Mensaje"
                    Exit Sub
                End If
                grid.Text = FormatCurrency(str(vldblDescuento))
                grid.TextMatrix(grid.Row, ColPorcentajeDescgrdPresupuesto) = vldblDescuento / Val(Format(grid.TextMatrix(grid.Row, ColSubtotalgrdPresupuesto), "###########.00"))
            End If
        End If
    End If
    
    If Index = 1 Then
        vgstrEstadoManto = "DU"
    Else
        vgstrEstadoManto = "C"
    End If
    pRecalcula Index

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol"))
End Sub

Private Sub pRecalcula(vlIndex As Integer)
    On Error GoTo NotificaError

    Dim X As Long
    Dim vldblTotalDescuentos As Double
    Dim vldblTotalIVA As Double
    Dim vldblTotalSubtotal As Double
    Dim vldblTotalImporte As Double
    
    Dim vldblPrecioElemento As Double
    Dim vldblCantidadElemento As Double
    Dim vldblDescuentoElemento As Double
    Dim vldblIVAElemento As Double
    Dim vldblMontoIVAElemento As Double
    Dim vldblMargenTotal As Double
    Dim vldblMargenUtilidadTotal As Double
    Dim vldblPrecioCantidadTotal As Double
    
    
    vldblTotalDescuentos = 0
    vldblTotalIVA = 0
    vldblTotalSubtotal = 0
    vldblCostoTotal = 0
    vldblMargen = 0
    vldblmargenutilidad = 0
    vldblMargenTotal = 0
    vldblMargenUtilidadTotal = 0
    vldblPrecioCantidadTotal = 0
    
    For X = 1 To grdPresupuesto(vlIndex).Rows - 1
        If grdPresupuesto(vlIndex).TextMatrix(X, ColPreciogrdPresupuesto) <> "" Then
            vldblPrecioElemento = Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColPreciogrdPresupuesto), "############.00"))
            vldblCantidadElemento = Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColCantidadgrdPresupuesto), "############.00"))
        
            grdPresupuesto(vlIndex).TextMatrix(X, ColSubtotalgrdPresupuesto) = FormatCurrency(vldblPrecioElemento * vldblCantidadElemento, 2)

            vldblDescuentoElemento = Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColSubtotalgrdPresupuesto), "############.00")) * Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColPorcentajeDescgrdPresupuesto), "############.0000"))
            vldblIVAElemento = Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColIVAgrdPresupuesto), "############.00"))
            vldblMontoIVAElemento = ((vldblPrecioElemento * vldblCantidadElemento - vldblDescuentoElemento) * vldblIVAElemento)
        
            grdPresupuesto(vlIndex).TextMatrix(X, ColMontogrdPresupuesto) = FormatCurrency((vldblPrecioElemento * vldblCantidadElemento) - vldblDescuentoElemento, 2)
            grdPresupuesto(vlIndex).TextMatrix(X, ColDescuentogrdPresupuesto) = FormatCurrency(vldblDescuentoElemento, 2)
            grdPresupuesto(vlIndex).TextMatrix(X, ColMontoIVAgrdPresupuesto) = FormatCurrency(vldblMontoIVAElemento, 2)
            
            vldblTotalImporte = vldblTotalImporte + (vldblPrecioElemento * vldblCantidadElemento)
            vldblTotalDescuentos = vldblTotalDescuentos + vldblDescuentoElemento
            vldblTotalSubtotal = vldblTotalSubtotal + ((vldblPrecioElemento * vldblCantidadElemento) - vldblDescuentoElemento)
            vldblTotalIVA = vldblTotalIVA + ((vldblPrecioElemento * vldblCantidadElemento - vldblDescuentoElemento) * vldblIVAElemento)
            
            vldblCostoTotal = CDbl(Val(Format(grdPresupuesto(vlIndex).TextMatrix(X, ColCostogrdPresupuesto), "############.##"))) * vldblCantidadElemento
            vldblMargen = ((vldblPrecioElemento * vldblCantidadElemento) - vldblDescuentoElemento) - vldblCostoTotal
            
            vldblPrecioCantidadTotal = (vldblPrecioElemento * vldblCantidadElemento)
            If (vldblMargen = 0) Then
                vldblmargenutilidad = 0
            End If
            
            If (vldblPrecioCantidadTotal <> vldblDescuentoElemento And vldblMargen > 0) Then 'calculo normal para cantidades positivas
                vldblmargenutilidad = (vldblMargen / ((vldblPrecioElemento * vldblCantidadElemento) - vldblDescuentoElemento)) * 100
                End If
            If (vldblMargen < 0 And vldblCostoTotal > 0) Then 'calculo margen utilidad negativas
                 If (vldblPrecioCantidadTotal = vldblDescuentoElemento And vldblCostoTotal = vldblDescuentoElemento) Then 'cuando importe=descuento=costo entonces es perdida -%100
                        vldblmargenutilidad = -100
                            Else
                              vldblmargenutilidad = ((vldblMargen * 100) / vldblCostoTotal)
                 End If
            End If
            grdPresupuesto(vlIndex).TextMatrix(X, ColMargenUtilidadgrdPresupuesto) = Format(vldblmargenutilidad, "0.00") & "%"
            vldblMargenTotal = vldblMargenTotal + vldblMargen
        End If
    Next X
     If vldblMargenTotal < 0 And vldblTotalSubtotal > 0 Then
        vldblMargenUtilidadTotal = (vldblMargenTotal * 100 / vldblTotalSubtotal)
    End If
    If vldblMargenTotal > 0 And vldblTotalSubtotal > 0 Then
        vldblMargenUtilidadTotal = (vldblMargenTotal / vldblTotalSubtotal) * 100
    End If
    txtMargenUtilidadTotal.Text = Format(vldblMargenUtilidadTotal, "0.00") & "%"
    
    'Totales
    If vlIndex = 0 Then
        txtImporte.Text = FormatCurrency(vldblTotalImporte, 2)
        txtDescuentos.Text = FormatCurrency(vldblTotalDescuentos, 2)
        txtSubtotal.Text = FormatCurrency((Round(vldblTotalImporte, 2) - Round(vldblTotalDescuentos, 2)), 2)
        txtIva.Text = FormatCurrency(vldblTotalIVA, 2)
        txtTotal.Text = FormatCurrency(Round(vldblTotalImporte, 2) - Round(vldblTotalDescuentos, 2) + Round(vldblTotalIVA, 2), 2)
    Else
        vldblImporte = vldblTotalImporte
        vldblSubtotal = Round(vldblTotalImporte, 2) - Round(vldblTotalDescuentos, 2)
        vldblDescuento = vldblTotalDescuentos
        vldblIVA = vldblTotalIVA
        vldbltotal = (Round(vldblImporte, 2) - Round(vldblTotalDescuentos, 2)) + Round(vldblTotalIVA, 2)
        grdPresupuesto(vlIndex).TextMatrix(grdPresupuesto(vlIndex).Rows - 5, ColMontogrdPresupuesto) = FormatCurrency(vldblImporte, 2)
        grdPresupuesto(vlIndex).Row = grdPresupuesto(vlIndex).Rows - 5
        grdPresupuesto(vlIndex).Col = ColSubtotalgrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).Col = ColMontogrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).TextMatrix(grdPresupuesto(vlIndex).Rows - 4, ColMontogrdPresupuesto) = FormatCurrency(vldblDescuento, 2)
        grdPresupuesto(vlIndex).Row = grdPresupuesto(vlIndex).Rows - 4
        grdPresupuesto(vlIndex).Col = ColSubtotalgrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).Col = ColMontogrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).TextMatrix(grdPresupuesto(vlIndex).Rows - 3, ColMontogrdPresupuesto) = FormatCurrency(vldblSubtotal, 2)
        grdPresupuesto(vlIndex).Row = grdPresupuesto(vlIndex).Rows - 3
        grdPresupuesto(vlIndex).Col = ColSubtotalgrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).Col = ColMontogrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).TextMatrix(grdPresupuesto(vlIndex).Rows - 2, ColMontogrdPresupuesto) = FormatCurrency(vldblIVA, 2)
        grdPresupuesto(vlIndex).Row = grdPresupuesto(vlIndex).Rows - 2
        grdPresupuesto(vlIndex).Col = ColSubtotalgrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).Col = ColMontogrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).TextMatrix(grdPresupuesto(vlIndex).Rows - 1, ColMontogrdPresupuesto) = FormatCurrency(vldbltotal, 2)
        grdPresupuesto(vlIndex).Row = grdPresupuesto(vlIndex).Rows - 1
        grdPresupuesto(vlIndex).Col = ColSubtotalgrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
        grdPresupuesto(vlIndex).Col = ColMontogrdPresupuesto
        grdPresupuesto(vlIndex).CellFontBold = True
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pRecalcula"))
End Sub

Public Sub pEditarColumna(vldblValor As Double, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError


    Dim vlintTexto As Integer

    txtEdit.Text = IIf(vldblValor = 0, "", vldblValor)
    txtEdit.SelStart = Len(txtEdit.Text)

    With grid
        If .CellWidth < 0 Then Exit Sub
            txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    
    vgstrEstadoManto = "E"
    txtEdit.Visible = True
    txtEdit.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna"))
End Sub

Private Sub txtCantidad_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pSetCellValueCol grdPresupuesto(Index), txtCantidad(Index), Index
        grdPresupuesto(Index).SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidad_KeyDown"))
End Sub

Private Sub txtCantidad_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
   
    ' Solo permite números
    If grdPresupuesto(Index).Col = ColCantidadgrdPresupuesto Then
        If KeyAscii = 46 Or Len(txtCantidad(Index)) > 2 Then
            KeyAscii = 7
        End If
    End If
    If grdPresupuesto(Index).Col = ColDescuentogrdPresupuesto Then
        If Len(txtCantidad(Index)) > 4 Then
            KeyAscii = 7
        End If
    End If
    If grdPresupuesto(Index).Col = ColPreciogrdPresupuesto Then
        If Len(txtCantidad(Index)) > 9 Then
            KeyAscii = 7
        End If
    End If
    If Not fblnFormatoCantidad(txtCantidad(Index), KeyAscii, 0) Or KeyAscii = 7 Then
        KeyAscii = 7
    End If
    'End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidad_KeyPress"))
End Sub

Private Sub txtSeleArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    Dim vlintcontador As Integer
    Dim vlstrSentencia As String
    Dim rsFiltroCargos As ADODB.Recordset
    Dim vlIntCont As Integer
    Dim vlstrNewcadena As String
    
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And grdElementos.Enabled Then
        pCargaElementos
        If grdElementos.Enabled Then
            grdElementos.SetFocus
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtSeleArticulo_KeyDown"))
End Sub

Private Sub txtSeleArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If optDescripcion.Value Then
        pCargaElementos
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtSeleArticulo_KeyUp"))
End Sub

Private Sub pLimpiaRenglonMSHFlexGrid(grdGrid As MSHFlexGrid, intRow As Integer)
    Dim intIndex As Integer
    For intIndex = 0 To grdGrid.Cols - 1
        grdGrid.TextMatrix(intRow, intIndex) = ""
    Next
End Sub

Private Sub pLimpiaGrid(grdGrid As MSHFlexGrid)
On Error GoTo NotificaError

    grdGrid.Clear
    grdGrid.Rows = 2
    grdGrid.Cols = 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaGrid"))
End Sub

Private Sub pCargaProcedimientos()
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    
    Set rs = frsRegresaRs("Select intcvecirugia, vchDescripcion from excirugia order by vchDescripcion", adLockReadOnly, adOpenForwardOnly)
    
    Do While Not rs.EOF
        cboProcedimiento.AddItem rs!VCHDESCRIPCION
        cboProcedimiento.ItemData(cboProcedimiento.newIndex) = rs!intCveCirugia
        rs.MoveNext
    Loop
    rs.Close
    
    cboProcedimiento.AddItem "", 0
    cboProcedimiento.ItemData(cboProcedimiento.newIndex) = 0
    cboProcedimiento.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaProcedimientos"))
End Sub

Private Sub pConsulta(rsConsultaPresupuesto As ADODB.Recordset)
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset

    pLimpiaGrid grdPresupuesto(0)
    pConfiguraGridCargos 0
    
    pHabilitaDatosPaquete False, True
    txtNumCuenta.Text = ""
    
    pHabilitaFrames True
    
    chkTomarDescuentos(0).Value = IIf(rsConsultaPresupuesto!mnyAplicarDescuento = 1, vbUnchecked, vbChecked)
    
    mskFechaPresupuesto.Text = IIf(IsNull(rsConsultaPresupuesto!fechapresupuesto), rsConsultaPresupuesto!fecha, rsConsultaPresupuesto!fechapresupuesto)
    mskFechaPresupuesto.Enabled = False
    
    txtDiasVencimiento.Text = rsConsultaPresupuesto!intDiasVencimiento
    txtDiasVencimiento.Enabled = False
    txtPorcentajeDescuento(0).Text = IIf(rsConsultaPresupuesto!mnyAplicarDescuento = 0, 0, IIf(IsNull(rsConsultaPresupuesto!mnyAplicarDescuento), "", rsConsultaPresupuesto!mnyAplicarDescuento))
    txtNombre.Text = Trim(rsConsultaPresupuesto!CHRNOMBRE)
    txtDireccion.Text = Trim(rsConsultaPresupuesto!chrDireccion)
    If rsConsultaPresupuesto!chrestado = "C" Then
        txtEstado.Text = "CREADO"
    ElseIf rsConsultaPresupuesto!chrestado = "A" Then
        txtEstado.Text = "AUTORIZADO"
        vlstrSentencia = "Select pvpaquetepaciente.*, pvpaquete.smiConceptoFactura, pvpaquete.chrTratamiento " & _
                         "From pvpaquetepaciente inner join pvpresupuesto on pvpresupuesto.intnumpaquete = pvpaquetepaciente.intnumpaquete " & _
                                                "inner join pvpaquete on pvpaquetepaciente.intnumpaquete = pvpaquete.intnumpaquete " & _
                         "Where pvpresupuesto.chrestado = 'A' and pvpresupuesto.intcvepresupuesto  = " & txtClave.Text
        Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
        If rs.RecordCount > 0 Then
            txtNumCuenta.Text = rs!INTMOVPACIENTE
            txtNumPaquete.Text = rs!intnumpaquete
            cboConceptoFactura.ListIndex = fintLocalizaCbo(cboConceptoFactura, rs!SMICONCEPTOFACTURA)
            cboTratamiento.ListIndex = fintLocalizaCritCbo(cboTratamiento, Trim(rs!chrTratamiento))
            pHabilitaDatosPaquete False, True
        End If
    ElseIf rsConsultaPresupuesto!chrestado = "N" Then
        txtEstado.Text = "NO AUTORIZADO"
        vlstrSentencia = "Select * from pvpresupuestolog " & _
                         "Where chrestado = 'N' and intcvepresupuesto  = " & txtClave.Text

        Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
        If rs.RecordCount > 0 Then
            txtMotivoNoAutorizado.Text = IIf(IsNull(rs!chrComentarios), "", rs!chrComentarios)
        End If
    Else
        txtEstado.Text = "SIN RESPUESTA"
    End If
                
    If rsConsultaPresupuesto!CHRTIPOPACIENTE = "I" Then
        optTipoPaciente(1).Value = True
    ElseIf rsConsultaPresupuesto!CHRTIPOPACIENTE = "E" Then
        optTipoPaciente(0).Value = True
    Else
        optTipoPaciente(2).Value = True
    End If
    optTipoPaciente(0).Enabled = False
    optTipoPaciente(1).Enabled = False
    optTipoPaciente(2).Enabled = False
    
    'vlblnCargando = False
    If IsNull(rsConsultaPresupuesto!INTCVEPROCEDIMIENTO) Then
        cboProcedimiento.ListIndex = -1
        'cboProcedimiento.Text = IIf(IsNull(rsConsultaPresupuesto!VCHPROCEDIMIENTO), "", rsConsultaPresupuesto!VCHPROCEDIMIENTO)
    Else
        cboProcedimiento.ListIndex = fintLocalizaCbo(cboProcedimiento, rsConsultaPresupuesto!INTCVEPROCEDIMIENTO)
    End If
    cboTipoPaciente(0).ListIndex = fintLocalizaCbo(cboTipoPaciente(0), rsConsultaPresupuesto!intTipoPaciente)
    
    vlstrSentencia = "Select chrTipo, bitDesconocido, bitUtilizaConvenio, bitFamiliar From AdTipoPaciente Where tnyCveTipoPaciente = " & str(cboTipoPaciente(0).ItemData(cboTipoPaciente(0).ListIndex))
    Set rs = frsRegresaRs(vlstrSentencia)
    If rs.RecordCount <> 0 Then
        vlstrTipoTipoPaciente = Trim(rs!chrTipo)
        'vlblnUtilizaConvenio = IIf(rs!bitUtilizaConvenio = 1, True, False)
    End If
    rs.Close
    
    If vlstrTipoTipoPaciente = "CO" Then
        'Primero el tipo de convenio y después la empresa
        vgstrParametrosSP = CStr(IIf(IsNull(rsConsultaPresupuesto!intEmpresa), 0, rsConsultaPresupuesto!intEmpresa)) & "|-1|-1"
        Set rs2 = frsEjecuta_SP(vgstrParametrosSP, "Sp_CcSelEmpresa")
        If Not rs2.EOF Then
            cboTipoConvenio(0).ListIndex = flngLocalizaCbo(cboTipoConvenio(0), str(IIf(IsNull(rs2!tnyCveTipoConvenio), 0, rs2!tnyCveTipoConvenio)))
        Else
            cboTipoConvenio(0).ListIndex = flngLocalizaCbo(cboTipoConvenio(0), 0)
        End If
        rs2.Close
        
        If rsConsultaPresupuesto!intEmpresa = 0 Then
            cboEmpresa(0).ListIndex = -1
        Else
            cboEmpresa(0).ListIndex = fintLocalizaCbo(cboEmpresa(0), rsConsultaPresupuesto!intEmpresa)
        End If
        cboEmpresa(0).Enabled = False
    End If
    
    If txtEstado.Text = "CREADO" Then
        pHabilitaBotones 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 0, 1
    ElseIf txtEstado.Text = "AUTORIZADO" Then
        pHabilitaBotones 0, 1, 1, 1, 1, 1, 0, 0, 1, 1, 0, IIf(Val(txtNumCuenta.Text) = 0, 0, 1), 0
        pHabilitaFramesUpdate False
    ElseIf txtEstado.Text = "NO AUTORIZADO" Then
        pHabilitaBotones 0, 1, 1, 1, 1, 1, 0, 0, 1, 1, 0, 0, 0
        pHabilitaFramesUpdate False
    ElseIf txtEstado.Text = "SIN RESPUESTA" Then
        pHabilitaBotones 0, 1, 1, 1, 1, 1, 0, 0, 1, 1, 0, 0, 1
        pHabilitaFramesUpdate False
    End If
    
    vgblnEsNuevo = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConsulta"))
End Sub

Private Sub pHabilitaBotones(blnActualiza As Boolean, blnPrimer As Boolean, blnAntes As Boolean, blnBusca As Boolean, blnSiguiente As Boolean, blnUltimo As Boolean, blnGraba As Boolean, blnBorra As Boolean, blnImprime As Boolean, blnDuplica As Boolean, blnPaquete As Boolean, blnPagos As Boolean, blnNoAutoriza As Boolean)
On Error GoTo NotificaError

    cmdActualizaPrecios(0).Enabled = blnActualiza
    cmdPrimerRegistro.Enabled = blnPrimer
    cmdAnteriorRegistro.Enabled = blnAntes
    cmdBuscar.Enabled = blnBusca
    cmdSiguienteRegistro.Enabled = blnSiguiente
    cmdUltimoRegistro.Enabled = blnUltimo
    cmdGrabarRegistro.Enabled = blnGraba
    cmdDelete.Enabled = blnBorra
    cmdImprimir.Enabled = blnImprime
    optTicket.Enabled = blnImprime
    optCarta.Enabled = blnImprime
    cmdDuplicar.Enabled = blnDuplica
    cmdCrearPaquete.Enabled = blnPaquete
    If blnPagos = False Then
        cmdVerPagos.Enabled = blnPagos
    Else
        'Verifica si el usuario tiene permiso para ver los pagos del paciente que ya tiene un paquete cotizado
        If fblnRevisaPermiso(vglngNumeroLogin, 4149, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4149, "L", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4149, "E", True) Then cmdVerPagos.Enabled = blnPagos
    End If
    cmdNoAutorizar.Enabled = blnNoAutoriza
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pNuevo"))
End Sub

Public Function fnGrabaPresupuesto(rs As ADODB.Recordset, vlintIndex As Integer, vlblnAlta As Boolean) As Boolean
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rsDetalle As ADODB.Recordset
    Dim vlintcontador As Integer
        
    EntornoSIHO.ConeccionSIHO.BeginTrans
       
    'Inicio de Grabada en tabla maestro
    fnGrabaPresupuesto = False
    If chkTomarDescuentos(vlintIndex).Value = vbChecked Then
        rs!bitAplicarDescuento = 0
        rs!mnyAplicarDescuento = 0
    Else
        rs!bitAplicarDescuento = 1
        rs!mnyAplicarDescuento = Val(Me.txtPorcentajeDescuento(vlintIndex).Text)
    End If
    rs!CHRTIPOPACIENTE = IIf(optTipoPaciente(0), "E", IIf(optTipoPaciente(1), "I", "U"))
    rs!smicvedepartamento = vgintNumeroDepartamento
    rs!dtmfecha = fdtmServerFechaHora
    rs!CHRNOMBRE = txtNombre.Text
    rs!chrDireccion = txtDireccion.Text
    rs!VCHPROCEDIMIENTO = Left(cboProcedimiento.List(cboProcedimiento.ListIndex), 50)
    If cboProcedimiento.ListIndex >= 0 Then
        rs!INTCVEPROCEDIMIENTO = cboProcedimiento.ItemData(cboProcedimiento.ListIndex)
    Else
        rs!INTCVEPROCEDIMIENTO = 0
    End If
    rs!intTipoPaciente = cboTipoPaciente(vlintIndex).ItemData(cboTipoPaciente(vlintIndex).ListIndex)
    If cboEmpresa(vlintIndex).ListIndex < 0 Then
        rs!intEmpresa = 0
    Else
        rs!intEmpresa = cboEmpresa(vlintIndex).ItemData(cboEmpresa(vlintIndex).ListIndex)
    End If
    If vlintIndex = 0 Then
        rs!MNYIMPORTE = CDec(txtImporte.Text)
        rs!MNYSUBTOTAL = CDec(txtSubtotal.Text)
        rs!mnyDescuentos = CDec(txtDescuentos.Text)
        rs!MNYIVA = CDec(txtIva.Text)
        rs!MNYTOTAL = CDec(txtTotal.Text)
    Else
        rs!MNYSUBTOTAL = vldblSubtotal
        rs!mnyDescuentos = vldblDescuento
        rs!MNYIVA = vldblIVA
        rs!MNYTOTAL = vldbltotal
    End If
    rs!chrMensaje1 = Trim(txtMensaje.Text)
    rs!chrMensaje2 = Trim(txtNotas.Text)
    
    rs!intDiasVencimiento = txtDiasVencimiento
    If vlintIndex = 1 Then
        rs!chrestado = "C"
        rs!dtmfechapresupuesto = fdtmServerFechaHora
    Else
        rs!dtmfechapresupuesto = CDate(mskFechaPresupuesto)
        If txtEstado.Text = "CREADO" Then
            rs!chrestado = "C"
        ElseIf txtEstado.Text = "AUTORIZADO" Then
            rs!chrestado = "A"
        ElseIf txtEstado.Text = "NO AUTORIZADO" Then
            rs!chrestado = "N"
        Else
            rs!chrestado = "S"
        End If
    End If
    rs.Update
            
    If vlblnAlta Then
        vlintCvePresupuesto = flngObtieneIdentity("SEC_PVPRESUPUESTO", rs!INTCVEPRESUPUESTO)
    Else
        vlintCvePresupuesto = rs!INTCVEPRESUPUESTO
    End If
    
    ' Borrado del detalle si es que existe
    vlstrSentencia = "Delete from PvDetallePresupuesto where intCvePresupuesto = " & vlintCvePresupuesto
    pEjecutaSentencia vlstrSentencia
    
    'Inicio de grabada en el detalle
    vlstrSentencia = "Select * from pvDetallePresupuesto where intCvePresupuesto = " & vlintCvePresupuesto
    Set rsDetalle = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    With grdPresupuesto(vlintIndex)
        For vlintcontador = 1 To grdPresupuesto(0).Rows - 1
            If .TextMatrix(vlintcontador, ColDescripciongrdPresupuesto) <> "" Then
                rsDetalle.AddNew
                rsDetalle!INTCVEPRESUPUESTO = vlintCvePresupuesto
                rsDetalle!intCveCargo = CLng(.TextMatrix(vlintcontador, ColRowDatagrdPresupuesto))
                rsDetalle!mnyPrecio = Val(Format(.TextMatrix(vlintcontador, ColPreciogrdPresupuesto), "############.##"))
                rsDetalle!intCantidad = CLng(Val(.TextMatrix(vlintcontador, ColCantidadgrdPresupuesto)))
                rsDetalle!MNYDESCUENTO = Val(Format(.TextMatrix(vlintcontador, ColDescuentogrdPresupuesto), "############.##"))
                'rsDetalle!MNYIVA = ((rsDetalle!MNYPRECIO * rsDetalle!intCantidad) - rsDetalle!MNYDESCUENTO) * Val(.TextMatrix(vlintContador, 8))
                rsDetalle!MNYIVA = Val(Format(.TextMatrix(vlintcontador, ColMontoIVAgrdPresupuesto), "############.##"))
                rsDetalle!chrTipoCargo = .TextMatrix(vlintcontador, ColTipogrdPresupuesto)
                rsDetalle!INTDESCUENTOINVENTARIO = .TextMatrix(vlintcontador, ColModoDescuentogrdPresupuesto)
                rsDetalle!mnycosto = Val(Format(.TextMatrix(vlintcontador, ColCostogrdPresupuesto), "############.##"))
                rsDetalle!NUMMARGENUTILIDAD = Val(Format(Replace(.TextMatrix(vlintcontador, ColMargenUtilidadgrdPresupuesto), "%", ""), "##.##"))
                rsDetalle.Update
            End If
        Next
    End With
    rsDetalle.Close
    rs.Close
    
    vlstrSentencia = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
      "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.Presupuesto
    pEjecutaSentencia vlstrSentencia
    
    vlstrSentencia = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.Presupuesto & "," & IIf(optTipoPaciente(0), "'E'", IIf(optTipoPaciente(1), "'I'", "'U'")) & ")"
    pEjecutaSentencia vlstrSentencia
    
    If vlblnAlta Then Call pGuardarLogPresupuesto("C", vllngPersonaGraba, CLng(vlintCvePresupuesto), "")
    Call pGuardarLogTransaccion(Me.Name, IIf(vlblnAlta, EnmGrabar, EnmCambiar), vglngNumeroLogin, "PRESUPUESTO", CStr(vlintCvePresupuesto))
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    'La operación se realizó satisfactoriamente.
    MsgBox SIHOMsg(420) & IIf(vlblnAlta, Chr(13) & "Presupuesto generado: " & vlintCvePresupuesto & ".", ""), vbInformation + vbOKOnly, "Mensaje"
    fnGrabaPresupuesto = True
    
Exit Function
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fnGrabaPresupuesto"))
End Function

Private Sub pMoverInformacion(grdOrigen As MSHFlexGrid, grdDestino As MSHFlexGrid)
On Error GoTo NotificaError
    Dim vlIntCont As Integer

    cboTipoPaciente(1).ListIndex = cboTipoPaciente(0).ListIndex
    cboTipoConvenio(1).ListIndex = cboTipoConvenio(0).ListIndex
    cboEmpresa(1).ListIndex = cboEmpresa(0).ListIndex
    vlIntCont = 1
    With grdDestino
        Do While vlIntCont <= grdOrigen.Rows - 1
            .RowData(grdDestino.Rows - 1) = grdOrigen.RowData(vlIntCont)
            .TextMatrix(.Rows - 1, ColRowDatagrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColRowDatagrdPresupuesto)
            .TextMatrix(.Rows - 1, ColDescripciongrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColDescripciongrdPresupuesto)
            .TextMatrix(.Rows - 1, ColMargenUtilidadgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColMargenUtilidadgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColCostogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColCostogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColPreciogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColPreciogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColCantidadgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColCantidadgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColSubtotalgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColSubtotalgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColDescuentogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColDescuentogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColMontogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColMontogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColTipogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColTipogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColIVAgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColIVAgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColMontoIVAgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColMontoIVAgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColPorcentajeDescgrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColPorcentajeDescgrdPresupuesto)
            .TextMatrix(.Rows - 1, ColContenidogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColContenidogrdPresupuesto)
            .TextMatrix(.Rows - 1, ColModoDescuentogrdPresupuesto) = grdOrigen.TextMatrix(vlIntCont, ColModoDescuentogrdPresupuesto)

            grdDestino.Rows = grdDestino.Rows + 1
            vlIntCont = vlIntCont + 1
        Loop
        .Rows = grdDestino.Rows + 5
        .TextMatrix(.Rows - 5, ColRowDatagrdPresupuesto) = -1
        .TextMatrix(.Rows - 5, ColSubtotalgrdPresupuesto) = "Importe"
        .TextMatrix(.Rows - 5, ColMontogrdPresupuesto) = 0
        .TextMatrix(.Rows - 4, ColRowDatagrdPresupuesto) = -1
        .TextMatrix(.Rows - 4, ColSubtotalgrdPresupuesto) = "Descuento"
        .TextMatrix(.Rows - 4, ColMontogrdPresupuesto) = 0
        .TextMatrix(.Rows - 4, ColRowDatagrdPresupuesto) = -1
        
        .TextMatrix(.Rows - 3, ColSubtotalgrdPresupuesto) = "Subtotal"
        .TextMatrix(.Rows - 3, ColMontogrdPresupuesto) = 0
        .TextMatrix(.Rows - 3, ColRowDatagrdPresupuesto) = -1
        
        .TextMatrix(.Rows - 2, ColSubtotalgrdPresupuesto) = "IVA"
        .TextMatrix(.Rows - 2, ColMontogrdPresupuesto) = 0
        .TextMatrix(.Rows - 2, ColRowDatagrdPresupuesto) = -1
        .TextMatrix(.Rows - 1, ColSubtotalgrdPresupuesto) = "Total"
        .TextMatrix(.Rows - 1, ColMontogrdPresupuesto) = 0
    End With
    
    pRecalcula 1
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMoverInformacion"))
End Sub
