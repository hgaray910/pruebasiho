VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEntradaSalidaDinero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entradas y salidas de dinero"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   1080
      TabIndex        =   77
      Top             =   1920
      Visible         =   0   'False
      Width           =   8760
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   78
         Top             =   600
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarraCFD 
         BackColor       =   &H80000002&
         Caption         =   "Generando el Comprobante Fiscal Digital, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   120
         Width           =   8610
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   0
         Left            =   30
         Top             =   120
         Width           =   8700
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000002&
         Caption         =   "Generando el Comprobante Fiscal Digital, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   79
         Top             =   180
         Width           =   8610
      End
   End
   Begin TabDlg.SSTab SSTabPagos 
      Height          =   8190
      Left            =   -105
      TabIndex        =   28
      Top             =   -540
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   14446
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      OLEDropMode     =   1
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmEntradaSalidaDinero.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPaciente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraRecibo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraImprimir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmEntradaSalidaDinero.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "chkSocios"
      Tab(1).Control(2)=   "cmdCargar"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame8"
      Tab(1).Control(5)=   "Frame6"
      Tab(1).Control(6)=   "fraRangoFechas"
      Tab(1).Control(7)=   "grdPagos"
      Tab(1).Control(8)=   "Label57(9)"
      Tab(1).Control(9)=   "Label57(13)"
      Tab(1).Control(10)=   "Label57(6)"
      Tab(1).Control(11)=   "Label57(12)"
      Tab(1).Control(12)=   "Label57(15)"
      Tab(1).Control(13)=   "Label57(7)"
      Tab(1).Control(14)=   "Label57(10)"
      Tab(1).Control(15)=   "Label1"
      Tab(1).ControlCount=   16
      Begin VB.Frame Frame1 
         Height          =   1725
         Left            =   -70855
         TabIndex        =   83
         Top             =   585
         Width           =   4575
         Begin VB.OptionButton optMostrarSolo 
            Caption         =   "Mostrar todo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Value           =   -1  'True
            Width           =   4305
         End
         Begin VB.OptionButton optMostrarSolo 
            Caption         =   "Mostrar sólo cancelación rechazada"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   91
            Top             =   1320
            Width           =   4305
         End
         Begin VB.OptionButton optMostrarSolo 
            Caption         =   "Mostrar sólo pendientes de autorización de cancelación"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   89
            Top             =   1000
            Width           =   4305
         End
         Begin VB.OptionButton optMostrarSolo 
            Caption         =   "Mostrar sólo pendientes de timbre fiscal"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   680
            Width           =   4305
         End
      End
      Begin VB.CheckBox chkSocios 
         Caption         =   "Socios"
         Height          =   255
         Left            =   -66180
         TabIndex        =   49
         Top             =   1600
         Width           =   1335
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   395
         Left            =   -66180
         TabIndex        =   50
         Top             =   1915
         Width           =   2085
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo movimiento"
         Height          =   960
         Left            =   -73040
         TabIndex        =   73
         Top             =   1350
         Width           =   2085
         Begin VB.OptionButton optTipoMovimiento 
            Caption         =   "Salidas de dinero"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   46
            Top             =   660
            Width           =   1650
         End
         Begin VB.OptionButton optTipoMovimiento 
            Caption         =   "Entradas de dinero"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   45
            Top             =   450
            Width           =   1725
         End
         Begin VB.OptionButton optTipoMovimiento 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   44
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Frame fraImprimir 
         Height          =   705
         Left            =   8400
         TabIndex        =   71
         Top             =   6360
         Width           =   2580
         Begin VB.ComboBox cboFormato 
            Height          =   315
            Left            =   100
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   240
            Width           =   2370
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Búsqueda por"
         Height          =   740
         Left            =   -73040
         TabIndex        =   70
         Top             =   585
         Width           =   2085
         Begin VB.OptionButton optTipoBusqueda 
            Caption         =   "Nombre del &paciente"
            Height          =   195
            Index           =   1
            Left            =   110
            TabIndex        =   43
            Top             =   450
            Width           =   1800
         End
         Begin VB.OptionButton optTipoBusqueda 
            Caption         =   "Rango de &fechas"
            Height          =   195
            Index           =   0
            Left            =   110
            TabIndex        =   42
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipo de paciente"
         Height          =   1725
         Left            =   -74745
         TabIndex        =   59
         Top             =   585
         Width           =   1620
         Begin VB.OptionButton optTipo 
            Caption         =   "&Deudor diverso"
            Height          =   210
            Index           =   3
            Left            =   90
            TabIndex        =   41
            ToolTipText     =   "Solo deudores diversos"
            Top             =   1320
            Width           =   1380
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Todos"
            Height          =   210
            Index           =   2
            Left            =   90
            TabIndex        =   38
            ToolTipText     =   "Todos"
            Top             =   360
            Width           =   900
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Externo"
            Height          =   210
            Index           =   1
            Left            =   90
            TabIndex        =   40
            ToolTipText     =   "Solo pacientes externos"
            Top             =   1000
            Width           =   900
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Interno"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   39
            ToolTipText     =   "Solo pacientes internos"
            Top             =   680
            Width           =   900
         End
      End
      Begin VB.Frame fraRangoFechas 
         Height          =   980
         Left            =   -66180
         TabIndex        =   58
         Top             =   585
         Width           =   2085
         Begin MSMask.MaskEdBox mskFechaFinal 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   48
            ToolTipText     =   "Fecha final de la búsqueda"
            Top             =   560
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaInicial 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   720
            TabIndex        =   47
            ToolTipText     =   "Fecha inicial de la búsqueda"
            Top             =   200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy "
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label12 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Desde"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraRecibo 
         Height          =   3345
         Left            =   160
         TabIndex        =   52
         Top             =   2970
         Width           =   10815
         Begin VB.ListBox lstFacturaASustituirDFP 
            Height          =   755
            IntegralHeight  =   0   'False
            ItemData        =   "frmEntradaSalidaDinero.frx":0038
            Left            =   7380
            List            =   "frmEntradaSalidaDinero.frx":003F
            TabIndex        =   18
            ToolTipText     =   "Pagos a los cuales sustituye"
            Top             =   1920
            Width           =   3315
         End
         Begin VB.CheckBox chkFacturaSustitutaDFP 
            Caption         =   "CFDI sustituto"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   6390
            TabIndex        =   17
            ToolTipText     =   "Indicar que el pago que se generará es sustituto de otro previamente cancelado"
            Top             =   1980
            Width           =   1005
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptos 
            Height          =   2010
            Left            =   1395
            TabIndex        =   12
            Top             =   675
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   3545
            _Version        =   393216
            GridColor       =   12632256
            FocusRect       =   0
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox txtComentario 
            Height          =   480
            Left            =   1400
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   19
            ToolTipText     =   "Comentario adicional"
            Top             =   2775
            Width           =   9300
         End
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   8565
            TabIndex        =   56
            Top             =   660
            Width           =   1995
            Begin VB.OptionButton optPesos 
               Caption         =   "Pesos"
               Height          =   195
               Left            =   60
               TabIndex        =   14
               ToolTipText     =   "Cantidad en pesos"
               Top             =   15
               Width           =   825
            End
            Begin VB.OptionButton optDolares 
               Caption         =   "Dólares"
               Height          =   195
               Left            =   960
               TabIndex        =   15
               ToolTipText     =   "Cantidad en dólares"
               Top             =   0
               Width           =   915
            End
         End
         Begin VB.TextBox txtPersona 
            Height          =   315
            Left            =   1395
            MaxLength       =   200
            TabIndex        =   11
            ToolTipText     =   "Persona que paga"
            Top             =   315
            Width           =   4905
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7080
            MaxLength       =   15
            TabIndex        =   13
            ToolTipText     =   "Cantidad del pago"
            Top             =   600
            Width           =   1380
         End
         Begin VB.TextBox txtCantidadenLetras 
            Height          =   830
            Left            =   6390
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Cantidad en letras"
            Top             =   990
            Width           =   4305
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Comentario"
            Height          =   195
            Left            =   90
            TabIndex        =   57
            Top             =   2790
            Width           =   795
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Persona"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblCantidad 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   6390
            TabIndex        =   54
            Top             =   660
            Width           =   630
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   195
            Left            =   90
            TabIndex        =   53
            Top             =   675
            Width           =   690
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPagos 
         Height          =   4080
         Left            =   -74760
         TabIndex        =   51
         Top             =   2400
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         GridColor       =   12632256
         FocusRect       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   2633
         TabIndex        =   37
         Top             =   6360
         Width           =   5655
         Begin VB.CommandButton cmdConfirmarTimbre 
            Caption         =   "Confirmar timbre fiscal"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4020
            Picture         =   "frmEntradaSalidaDinero.frx":005C
            TabIndex        =   76
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1080
         End
         Begin VB.CommandButton cmdCFD 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5100
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmEntradaSalidaDinero.frx":054E
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   3525
            Picture         =   "frmEntradaSalidaDinero.frx":0E6C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Imprimir recibo"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3030
            Picture         =   "frmEntradaSalidaDinero.frx":156E
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Borrar pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2535
            MaskColor       =   &H80000000&
            Picture         =   "frmEntradaSalidaDinero.frx":1A60
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Grabar pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2040
            Picture         =   "frmEntradaSalidaDinero.frx":1DA2
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Ultimo pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1550
            Picture         =   "frmEntradaSalidaDinero.frx":2294
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Siguiente pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1050
            Picture         =   "frmEntradaSalidaDinero.frx":2786
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Consulta de pagos"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   555
            Picture         =   "frmEntradaSalidaDinero.frx":2C78
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Anterior pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmEntradaSalidaDinero.frx":316A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Primer pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame fraPaciente 
         Height          =   2430
         Left            =   160
         TabIndex        =   29
         Top             =   525
         Width           =   10800
         Begin VB.Frame freEntradaSalida 
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   390
            Left            =   1560
            TabIndex        =   68
            Top             =   2010
            Width           =   4305
            Begin VB.OptionButton optEntradaDinero 
               Caption         =   "Entrada de dinero"
               Height          =   315
               Left            =   60
               TabIndex        =   9
               Top             =   30
               Value           =   -1  'True
               Width           =   1620
            End
            Begin VB.OptionButton optSalidaDinero 
               Caption         =   "Salida de dinero"
               Height          =   330
               Left            =   1710
               TabIndex        =   10
               Top             =   30
               Width           =   1890
            End
         End
         Begin VB.Frame Frame7 
            Height          =   2235
            Left            =   7440
            TabIndex        =   67
            Top             =   120
            Width           =   75
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1455
            Left            =   7590
            TabIndex        =   63
            Top             =   195
            Width           =   3165
            Begin VB.TextBox txtFolioRecibo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1275
               TabIndex        =   61
               Top             =   390
               Width           =   1845
            End
            Begin VB.TextBox txtDocumento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1275
               Locked          =   -1  'True
               TabIndex        =   60
               ToolTipText     =   "Documento en el cual se descontó el pago"
               Top             =   720
               Width           =   1845
            End
            Begin MSMask.MaskEdBox mskFecha 
               Height          =   315
               Left            =   1275
               TabIndex        =   62
               Top             =   45
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   556
               _Version        =   393216
               Format          =   "dd/mmm/yyyy"
               PromptChar      =   " "
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Folio del recibo"
               Height          =   195
               Left            =   75
               TabIndex        =   66
               Top             =   450
               Width           =   1185
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Left            =   90
               TabIndex        =   65
               Top             =   105
               Width           =   1185
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Factura"
               Height          =   195
               Left            =   75
               TabIndex        =   64
               Top             =   780
               Width           =   540
            End
         End
         Begin VB.TextBox txtMovimientoPaciente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "Número de cuenta del paciente"
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox txtPaciente 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Nombre del paciente"
            Top             =   585
            Width           =   5700
         End
         Begin VB.TextBox txtEmpresaPaciente 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Nombre de la empresa del paciente"
            Top             =   915
            Width           =   5700
         End
         Begin VB.TextBox txtTipoPaciente 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "Tipo de paciente"
            Top             =   1260
            Width           =   4035
         End
         Begin VB.TextBox txtFechaInicial 
            Height          =   315
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Fecha de inicio de atención"
            Top             =   1605
            Width           =   1890
         End
         Begin VB.TextBox txtFechaFinal 
            Height          =   315
            Left            =   3765
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Fecha final de atención"
            Top             =   1605
            Width           =   1890
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   510
            Left            =   3310
            TabIndex        =   30
            Top             =   180
            Width           =   4350
            Begin VB.OptionButton optDeudor 
               Caption         =   "Deudor diverso"
               Height          =   255
               Left            =   2630
               TabIndex        =   74
               Top             =   120
               Width           =   1380
            End
            Begin VB.OptionButton optSocio 
               Caption         =   "Socio"
               Height          =   195
               Left            =   1830
               TabIndex        =   3
               Top             =   135
               Width           =   700
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Externo"
               Height          =   195
               Index           =   1
               Left            =   880
               TabIndex        =   2
               Top             =   135
               Width           =   840
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Interno"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   1
               Top             =   135
               Value           =   -1  'True
               Width           =   800
            End
         End
         Begin VB.Label lblPagoCancelado 
            Alignment       =   2  'Center
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Recibo cancelado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   7640
            TabIndex        =   96
            Top             =   1845
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de movimiento"
            Height          =   255
            Left            =   105
            TabIndex        =   69
            Top             =   2085
            Width           =   1485
         End
         Begin VB.Label lbCuenta 
            AutoSize        =   -1  'True
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   105
            TabIndex        =   36
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   35
            Top             =   645
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   105
            TabIndex        =   34
            Top             =   975
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   105
            TabIndex        =   33
            Top             =   1320
            Width           =   1200
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de atención"
            Height          =   195
            Left            =   105
            TabIndex        =   32
            Top             =   1665
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   3570
            TabIndex        =   31
            Top             =   1665
            Width           =   90
         End
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A "
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   9
         Left            =   -71160
         TabIndex        =   95
         Top             =   6560
         Width           =   255
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A "
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   13
         Left            =   -74760
         TabIndex        =   94
         Top             =   6830
         Width           =   255
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A "
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   6
         Left            =   -74760
         TabIndex        =   93
         Top             =   6560
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de autorización de cancelación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -74400
         TabIndex        =   92
         Top             =   6845
         Width           =   3060
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación rechazada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -70800
         TabIndex        =   90
         Top             =   6575
         Width           =   1680
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74400
         TabIndex        =   88
         Top             =   6575
         Width           =   855
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   -71160
         TabIndex        =   82
         Top             =   6830
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de timbre fiscal"
         Height          =   195
         Left            =   -70800
         TabIndex        =   81
         Top             =   6845
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmEntradaSalidaDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------
' Programa para registro, consulta y cancelación de pagos por los diferentes conceptos.
' Ejem: Anticipo internamiento, liquidación de cuenta, abono a cuenta, pago deducible, etc.
' Fecha de programación: Martes 23 de Enero de 2001
'-------------------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'       01/Oct/2002 : Inclusión del parámetro para Concepto de Deducible y Coaseguro
'       28/Nov/2012 : Inclusión de la forma de pago Transferencias para las Entradas de dinero.
'-------------------------------------------------------------------------------------------------
' Fecha:        05/Julio/2003
' Por:          Rodolfo Ramos Garcia
' Descripción:  Que valide el Deducible, coaseguro y copago por separado para poder ser facturado.
'-------------------------------------------------------------------------------------------------

Option Explicit

'Columnas de la consulta de entradas y salidas:
Const cintColIdEntradaSalida = 1
Const cintColFecha = 2
Const cintColTipo = 3
Const cIntColFolio = 4
Const cintColCuenta = 5
Const cintColPaciente = 6
Const cintColCantidad = 7
Const cintColMoneda = 8
Const cintColEmpleado = 9
Const cintColEmpleadoCancela = 10
Const cintColDepartamento = 11
Const cintColbitCancelado = 12
Const cintColTipoMovimiento = 13
Const cintColPendienteCancelarSAT = 14

Const cintColumnas = 15

'Columnas de la lista de conceptos:
Public lintColNombreConcepto As Integer
Public lintColIdConcepto As Integer
Public lintColTipoConcepto As Integer
Public lintColNumeroCuenta As Integer

'--------------------------------------------'
' Propiedades para el control de la pantalla '
'--------------------------------------------'
Public lblnManipulacion As Boolean
Public lblnGrabo As Boolean

Public vglngMovPaciente As Long
Public vgstrTipoPago As String
Public vgstrTipoPaciente As String
Public vgstrTipoEntradaSalida As String

Dim vlstrSentencia As String

Dim vlblnLimpiar As Boolean
Dim vlblnConsulta As Boolean
Dim vlblnPacienteSeleccionado As Boolean

Dim rsEntradaSalida As New ADODB.Recordset

Dim vllngFormatoaUsar As Long
Dim vllngPersonaGraba As Long

Dim vldblTipoCambioVenta As Double

Dim aFormasPago() As FormasPago

Private vgrptReporte As CRAXDRT.Report

Dim alstrParametrosSalida() As String

Dim lintDeducible As Integer    'Para saber si al paciente se le puede registrar un pago por deducible
Dim lintCoaseguro As Integer    'Para saber si al paciente se le puede registrar un pago por coaseguro
Dim lintCoaseguroAdicional As Integer 'Para saber si al paciente se le puede registrar un pago por coaseguro adicional
Dim lintCopago As Integer       'Para saber si al paciente se le puede registrar un pago por copago

Dim rsControlSeguro As New ADODB.Recordset
Dim lintCveEmpresa As Integer   'Clave de la empresa del paciente, se usa para afectar el control de seguros en caso de pagos por deducible, coaseguro, copago, etc.
Dim llngNumCuenta As Long       'Número de cuenta del paciente seleccionado en la consulta
Dim ldtmFecha As Date           'Fecha actual
Dim vlblnValidarDesglosarIVA As Boolean   'Valida si está activo el parametro "Desglosar IVA en pagos"

Public vgblnCuentaFacturada As Boolean    '(CR) - Agregado para caso no. 6863 - Valida si la cuenta ya está facturada
Public vgblnLiquidicacion As Boolean      '(CR) - Agregado para caso no. 6894 - Valida si se manda llamar desde liquidación de cuenta en Facturación

Dim vlrfcPaciente As String

Dim vlintIncluidoenFactura As Integer
Dim llngCveConceptoDeudor As Long       'Clave del concepto de salida del deudor diverso
Dim lstrDescConceptoDeudor As String    'Descripción del concepto de salida para deudor diverso
Dim lblnConsultaDeudor As Boolean       'Indica si se esta consultando a un deudor diverso

'Datos fiscales
Dim vlstrDFRFC As String
Dim vlstrDFNombre As String
Dim vlstrDFDireccion As String
Dim vlstrDFNumExterior As String
Dim vlstrDFNumInterior As String
Dim vlBitDFExtranjero As Integer
Dim lngCveCiudad As Long                        'Clave de la ciudad del domicilio fiscal
Dim vlstrDFTelefono As String
Dim vlstrDFColonia As String
Dim vlstrDFCodigoPostal As String
Dim vlintUsoCFDI As Long
Dim vlstrRegimenFiscal As String
Dim vlstrCodigoPostal As String
Dim vlstrsql As String
Dim vlnblnLocate As Boolean

Dim vlblnCancelarRecibosOtroDepto As Boolean
Dim vlblnActivaMotivo As Boolean
Dim vlblnCancelaErrorTim As Boolean 'Si hubo error de timbre luego de quedar pendiente, damos click internamente al botón de cancelar y con esto nos saltamos las validaciones de usuario

Public Sub pInicializaReporte()
On Error GoTo NotificaError

    Dim strSentencia As String
    Dim rsFormatos As New ADODB.Recordset
        
    If cboFormato.ListCount = 0 Or optSalidaDinero.Value Then
        pInstanciaReporte vgrptReporte, "rptPVReciboPago.rpt"
    Else
        strSentencia = "SELECT vchReporte FROM PvFormatoRecibo WHERE intIdFormato = " & cboFormato.ItemData(cboFormato.ListIndex)
        Set rsFormatos = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsFormatos.RecordCount > 0 Then
            pInstanciaReporte vgrptReporte, rsFormatos!vchReporte
        Else
            pInstanciaReporte vgrptReporte, "rptPVReciboPago.rpt"
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pInicializaReporte"))
End Sub

Public Sub pFormatosRecibo()
On Error GoTo NotificaError

    Dim strSentencia As String
    Dim rsFormatos As New ADODB.Recordset
    
    strSentencia = "SELECT * FROM PvFormatoRecibo ORDER BY intIdFormato"
    Set rsFormatos = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    With cboFormato
        .Clear
        Do While Not rsFormatos.EOF
            .AddItem rsFormatos!VCHDESCRIPCION
            .ItemData(.newIndex) = rsFormatos!intIdFormato
            rsFormatos.MoveNext
        Loop
        If .ListCount <> 0 Then
            .ListIndex = 0
        End If
    End With
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pFormatosRecibo"))
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intPrint As Integer, intDelete As Integer)
On Error GoTo NotificaError
Dim rs As ADODB.Recordset
    
    cmdTop.Enabled = intTop = 1
    cmdBack.Enabled = intBack = 1
    cmdLocate.Enabled = intlocate = 1
    cmdNext.Enabled = intNext
    cmdEnd.Enabled = intEnd = 1
    cmdSave.Enabled = intSave = 1
    cmdPrint.Enabled = IIf(cmdCFD.Enabled, 0, IIf(cmdConfirmartimbre.Enabled, 0, intPrint = 1))
    cmdDelete.Enabled = IIf(lblPagoCancelado.Caption = "Pendiente timbre", 0, intDelete = 1)
    
'    If rsEntradaSalida.State <> 0 Then
'        'Si está activado el parámetro “Utilizar cuenta puente en sustitución de cuentas de banco al cancelar y/o
'        're facturar”, los pagos automáticos generados al cancelar facturas no se podrán cancelar debido a que
'        'el importe de este tipo de pago corresponde al importe de la factura que afectó a bancos, el cual
'        'deberá ser solicitado por el sistema al momento de facturar nuevamente la cuenta
'        vlstrSentencia = "SELECT PvConceptoPago.* FROM PvConceptoPago WHERE intNumConcepto = " & Str(rsEntradaSalida!intNumConcepto)
'        Set rs = frsRegresaRs(vlstrSentencia)
'        If rs.RecordCount <> 0 Then
'            If rs!bitpagocancelafactura = 1 Then
'                If cmdDelete.Visible Then cmdDelete.Enabled = False
'            End If
'        End If
'    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub pLimpia()
On Error GoTo NotificaError
    
    fraPaciente.Enabled = True
    
    vlblnConsulta = False
    lblnConsultaDeudor = False
    vlblnPacienteSeleccionado = False
    txtMovimientoPaciente.Text = ""
    txtPaciente.Text = ""
    vlrfcPaciente = ""
    txtEmpresaPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtFechaFinal.Text = ""
    txtFechaInicial.Text = ""

    mskFecha.Mask = ""
    mskFecha.Text = ldtmFecha
    mskFecha.Mask = "##/##/####"
    txtPersona.Text = ""
    txtCantidad.Text = ""
    txtCantidadenLetras.Text = ""
    optPesos.Value = False
    optDolares.Value = False
    txtDocumento.Text = ""
    txtComentario.Text = ""
    
    optEntradaDinero.Value = False
    optSalidaDinero.Value = False
    
    pCargaFolio IIf(optEntradaDinero.Value, "RE", IIf(optSalidaDinero.Value, "SD", ""))
    
    pLimpiaConceptos
    
    fraRecibo.Enabled = False
    freEntradaSalida.Enabled = True
    
    'Reiniciar los filtros de la búsqueda
    optTipoBusqueda(0).Value = True
    optTipo(2).Value = True
    optTipoMovimiento(0).Value = True
    
    mskFechaInicial.Mask = ""
    mskFechaInicial.Text = ldtmFecha
    mskFechaInicial.Mask = "##/##/####"
    
    mskFechaFinal.Mask = ""
    mskFechaFinal.Text = ldtmFecha
    mskFechaFinal.Mask = "##/##/####"
    
    pLimpiaGridConsulta
    pConfiguraGridConsulta

    lblPagoCancelado.Visible = False
    
    fraImprimir.Visible = False
    
    cmdSave.Enabled = False
    cmdLocate.Enabled = True
    cmdCFD.Enabled = False
    cmdConfirmartimbre.Enabled = False
    
    chkFacturaSustitutaDFP.Enabled = False
    chkFacturaSustitutaDFP.Value = 0
    lstFacturaASustituirDFP.Enabled = False
    
    
    lstFacturaASustituirDFP.Clear
    vlnblnLocate = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Public Sub pLimpiaConceptos()
On Error GoTo NotificaError

    With grdConceptos
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .ColWidth(0) = 100
        .ColWidth(lintColNombreConcepto) = 4450
        .ColWidth(lintColIdConcepto) = 0
        .ColWidth(lintColTipoConcepto) = 0
        .ColWidth(lintColNumeroCuenta) = 0
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaConceptos"))
End Sub

Private Sub cboFormato_Change()
On Error GoTo NotificaError

    pInicializaReporte
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFormato_Change"))
End Sub

Private Sub chkFacturaSustitutaDFP_Click()
Dim i As Integer

 If Not vlnblnLocate Then
    If chkFacturaSustitutaDFP.Value = 0 Then
        lstFacturaASustituirDFP.Clear
        ReDim aFoliosPrevios(0)
    Else
        If chkFacturaSustitutaDFP.Value = 1 Then
            frmBusquedaFacturasPreviasVP.vlstrsql = vlstrsql
            frmBusquedaFacturasPreviasVP.Caption = "Comprobantes previos"
            frmBusquedaFacturasPreviasVP.grdFacturas.ToolTipText = "Comprobante previamente cancelado al cual se sustituye"
            frmBusquedaFacturasPreviasVP.Show vbModal, Me
            
            lstFacturaASustituirDFP.Clear
            For i = 0 To UBound(aFoliosPrevios())
                If aFoliosPrevios(i).chrfoliofactura <> "" Then
                    lstFacturaASustituirDFP.AddItem aFoliosPrevios(i).chrfoliofactura
                End If
            Next i
            
            If frmBusquedaFacturasPreviasVP.vlchrfoliofactura = "" Then
                chkFacturaSustitutaDFP.Value = 0
            End If
        End If
    End If
  End If
End Sub

Private Sub chkFacturaSustitutaDFP_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
   
    If KeyAscii = 13 Then
        txtComentario.SetFocus
    Else
        SendKeys vbTab
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkFacturaSustitutaDFP_KeyPress"))
End Sub

Private Sub chkSocios_Click()
On Error GoTo NotificaError

    If chkSocios.Value = 1 Then
        Frame6.Enabled = False
        Frame8.Enabled = False
        Frame4.Enabled = False
        optTipo(0).Enabled = False
        optTipo(1).Enabled = False
        optTipo(2).Enabled = False
        optTipo(3).Enabled = False
        optTipoBusqueda(0).Value = True
        optTipoBusqueda(1).Enabled = False
        optTipoMovimiento(1).Value = True
        optTipoMovimiento(0).Enabled = False
        optTipoMovimiento(1).Enabled = True
        optTipoMovimiento(2).Enabled = False
    Else
        Frame6.Enabled = True
        Frame8.Enabled = True
        Frame4.Enabled = True
        optTipo(0).Enabled = True
        optTipo(1).Enabled = True
        optTipo(2).Enabled = True
        optTipo(3).Enabled = True
        optTipoBusqueda(0).Enabled = True
        optTipoBusqueda(1).Enabled = True
        optTipoMovimiento(0).Value = True
        optTipoMovimiento(0).Enabled = True
        optTipoMovimiento(1).Enabled = True
        optTipoMovimiento(2).Enabled = True
    End If
    
    cmdCargar_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkSocios_Click"))
End Sub

Private Sub cmdCargar_Click()
On Error GoTo NotificaError
    Dim intCol As Integer
    Dim intcontador As Long
    Dim rs As New ADODB.Recordset
    Dim vlForeColor As Variant
    Dim vlBackColor As Variant
    
    If Not IsDate(mskFechaFinal.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        pEnfocaMkTexto mskFechaFinal
    Else
        If CDate(mskFechaInicial.Text) > CDate(mskFechaFinal.Text) Then
            '¡Rango de fechas no valido!
            MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaInicial.SetFocus
        Else
            pLimpiaGridConsulta
            
            Me.MousePointer = 11
            
            Set rs = frsEjecuta_SP(IIf(optTipoMovimiento(0).Value, "*", IIf(optTipoMovimiento(1).Value, "E", "S")) & "|" & _
            IIf(fraRangoFechas.Enabled, "F", "P") & "|" & _
            IIf(optTipo(2).Value, "*", IIf(optTipo(0).Value, "I", IIf(optTipo(3).Value, "D", "E"))) & "|" & _
            fstrFechaSQL(mskFechaInicial.Text) & "|" & _
            fstrFechaSQL(mskFechaFinal.Text) & "|" & _
            llngNumCuenta & "|" & vgintClaveEmpresaContable & "|" & chkSocios.Value & "|" & IIf(Not optMostrarSolo(1).Value, 0, 1) & "|" & IIf(Not optMostrarSolo(3).Value, 0, 1) & "|" & IIf(Not optMostrarSolo(4).Value, 0, 1), "SP_PVSELENTSALDINEROFEC_NE")
            
            With rs
                If .RecordCount <> 0 Then
                    .MoveFirst
                    Do While Not .EOF
                        If Trim(grdPagos.TextMatrix(1, 1)) <> "" Then
                            grdPagos.Rows = grdPagos.Rows + 1
                        End If
                        
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColFecha) = Format(!FechaDocumento, "dd/mmm/yyyy")
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColTipo) = !tipo
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cIntColFolio) = !FolioDocumento
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColCuenta) = IIf(IsNull(!cuenta), "", !cuenta)
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColPaciente) = !NOMBREPACIENTE
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColCantidad) = FormatCurrency(!cantidad, 2)
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColMoneda) = !Moneda
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColEmpleado) = !Empleado
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColEmpleadoCancela) = !EmpleadoCancela
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColDepartamento) = !departamento
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColIdEntradaSalida) = !NumEntradaSalida
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColbitCancelado) = !Cancelado
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColTipoMovimiento) = !TipoMovimiento
                        grdPagos.TextMatrix(grdPagos.Rows - 1, cintColPendienteCancelarSAT) = !PendienteCancelarSAT_NE
                        
                        If !Cancelado Then
                            For intcontador = 1 To grdPagos.Cols - 1
                                grdPagos.Col = intcontador
                                grdPagos.Row = grdPagos.Rows - 1
                                grdPagos.CellForeColor = &HFF&
                            Next intcontador
                        Else
                            vlForeColor = vbBlack  '| Negro
                            vlBackColor = &HFFFFFF '| Blanco
                            If fblnPendienteTimbre(!NumEntradaSalida) Then
                                grdPagos.Redraw = False
                                grdPagos.Row = grdPagos.Rows - 1
                                For intCol = 1 To grdPagos.Cols - 1
                                    grdPagos.Col = intCol
                                    grdPagos.CellBackColor = &H80FFFF 'amarillo
                                Next
                            Else
                                grdPagos.Redraw = False
                                grdPagos.Row = grdPagos.Rows - 1
                    
                                Select Case rs!PendienteCancelarSAT_NE
                                    Case "PC", "XX"
                                        vlForeColor = &HFF&    '| Rojo
                                        vlBackColor = &HC0E0FF '| Naranja
                                    Case "PA"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &H80FF&  '| Naranja fuerte
                                    Case "CR"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &HFF&    '| Rojo
                                End Select
                            
                                For intCol = 1 To grdPagos.Cols - 1
                                    grdPagos.Col = intCol
                                    grdPagos.CellForeColor = vlForeColor
                                    grdPagos.CellBackColor = vlBackColor
                                Next
                            End If
                        End If
                            
                        .MoveNext
                    Loop
                End If
                .Close
            End With
            
            Me.MousePointer = 0
            
            pConfiguraGridConsulta
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargar_Click"))
End Sub

Private Sub cmdCFD_Click()
    frmComprobanteFiscalDigitalInternet.lngComprobante = rsEntradaSalida!intNumPago
    frmComprobanteFiscalDigitalInternet.strTipoComprobante = "AN"
    'frmComprobanteFiscalDigitalInternet.blnCancelado = txtCanceladada.Visible
    frmComprobanteFiscalDigitalInternet.blnFacturaSinComprobante = False
    frmComprobanteFiscalDigitalInternet.Show vbModal, Me
End Sub

Private Sub cmdConfirmartimbre_Click()
    Dim lngCveFormato As Long
    Dim lngNumPagoSalida As Long
    Dim intResultado As Integer
    EntornoSIHO.ConeccionSIHO.BeginTrans
    intResultado = pGeneraCFDI(False, rsEntradaSalida!intNumPago, "0", "0", "", "")
    If intResultado = 0 Then
        lngNumPagoSalida = rsEntradaSalida!intNumPago
        pEliminaPendientesTimbre rsEntradaSalida!intNumPago, "AN"
    ElseIf intResultado = 2 Then
        pEliminaPendientesTimbre rsEntradaSalida!intNumPago, "AN"
        vlblnCancelaErrorTim = True
        cmdDelete_Click
        pMuestraPago rsEntradaSalida!intNumPago, IIf(optEntradaDinero.Value, "E", "S")
    End If
    If intResultado < 2 Then
        EntornoSIHO.ConeccionSIHO.CommitTrans
        lngCveFormato = 1
        frsEjecuta_SP vgintNumeroDepartamento & "|0|0|T", "fn_PVSelFormatoFactura2", True, lngCveFormato
        fblnImprimeComprobanteDigital rsEntradaSalida!intNumPago, "AN", "I", lngCveFormato, 0
        If intResultado = 0 Then
            If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
                '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pEnviarCFD "AN", lngNumPagoSalida, CLng(vgintClaveEmpresaContable), Trim(vlstrDFRFC), vglngNumeroEmpleado, Me
                End If
            End If
        End If
        txtMovimientoPaciente.SetFocus
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    If lblnManipulacion Then Exit Sub
    
    If SSTabPagos.Tab = 1 Then
        optDeudor.Value = False
        optSocio.Value = False
        OptTipoPaciente(0).Value = True
        pLimpia
        SSTabPagos.Tab = 0
        txtMovimientoPaciente.SetFocus
        Cancel = 1
    Else
        If cmdSave.Enabled Or cmdPrint.Enabled Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                optDeudor.Value = False
                optSocio.Value = False
                OptTipoPaciente(0).Value = True
                pLimpia
                pEnfocaTextBox txtMovimientoPaciente
            End If
            Cancel = 1
        Else
            vglngMovPaciente = 0
            vgstrTipoPaciente = ""
            vgstrTipoPago = ""
            vgstrTipoEntradaSalida = ""
            Unload Me
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub grdConceptos_Click()
On Error GoTo NotificaError

    pCargarCantidad
    
    If optSocio.Value = True Then
        optDolares.Enabled = False
        optPesos.Value = True
    End If
    If fblnGenerarComprobante And txtMovimientoPaciente.Text <> "" Then
        pFacturasDirectasAnteriores
    Else
        chkFacturaSustitutaDFP.Enabled = False
        lstFacturaASustituirDFP.Enabled = False
        
        lstFacturaASustituirDFP.Clear
        chkFacturaSustitutaDFP.Value = 0
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdConceptos_Click"))
End Sub

Private Sub grdConceptos_GotFocus()
On Error GoTo NotificaError

    pCargarCantidad
    
    If optSocio.Value = True Then
        optDolares.Enabled = False
        optPesos.Value = True
    End If
 
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdConceptos_GotFocus"))
End Sub

Private Sub grdConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If lblCantidad.Enabled Then
            txtCantidad.SetFocus
        Else
            txtComentario.SetFocus
        End If
    Else
        If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            pCargarCantidad
            
            If optSocio.Value = True Then
                optDolares.Enabled = False
                optPesos.Value = True
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdConceptos_KeyDown"))
End Sub

Private Sub pCargarCantidad()
On Error GoTo NotificaError

    If vlblnConsulta Or lblnManipulacion Then Exit Sub

    With grdConceptos
        lblCantidad.Enabled = .TextMatrix(.Row, lintColTipoConcepto) = "NO" Or .TextMatrix(.Row, lintColTipoConcepto) = "SD"
        txtCantidad.Enabled = .TextMatrix(.Row, lintColTipoConcepto) = "NO" Or .TextMatrix(.Row, lintColTipoConcepto) = "SD"
        optPesos.Enabled = .TextMatrix(.Row, lintColTipoConcepto) = "NO" Or .TextMatrix(.Row, lintColTipoConcepto) = "SD"
        optDolares.Enabled = .TextMatrix(.Row, lintColTipoConcepto) = "NO" Or .TextMatrix(.Row, lintColTipoConcepto) = "SD"
        
        If .TextMatrix(.Row, lintColTipoConcepto) = "DE" Then 'Deducible
            txtCantidad.Text = FormatCurrency(rsControlSeguro!MNYCANTIDADDEDUCIBLE, 2)
            optPesos.Value = True
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        ElseIf .TextMatrix(.Row, lintColTipoConcepto) = "CO" Then 'Coaseguro
            txtCantidad.Text = FormatCurrency(rsControlSeguro!MNYCANTIDADCOASEGURO, 2)
            optPesos.Value = True
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        ElseIf .TextMatrix(.Row, lintColTipoConcepto) = "CA" Then 'Coaseguro adicional
            txtCantidad.Text = FormatCurrency(rsControlSeguro!MNYCANTIDADCOASEGUROADICIONAL, 2)
            optPesos.Value = True
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        ElseIf .TextMatrix(.Row, lintColTipoConcepto) = "CP" Then 'Copago
            txtCantidad.Text = FormatCurrency(rsControlSeguro!MNYCANTIDADCOPAGO, 2)
            optPesos.Value = True
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        Else
            txtCantidad.Text = ""
            optPesos.Value = False
            optDolares.Value = False
            txtCantidadenLetras.Text = ""
        End If
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargarCantidad"))
End Sub

Private Sub grdConceptos_RowColChange()
    grdConceptos_Click
End Sub

Private Sub grdPagos_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        grdPagos_DblClick
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPagos_KeyPress"))
End Sub

Private Sub cmdBack_Click()
On Error GoTo NotificaError
    
    If grdPagos.Row > 1 Then
        grdPagos.Row = grdPagos.Row - 1
    End If
    pMuestraPago CLng(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)), grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento)
    pHabilita 1, 1, 1, 1, 1, 0, IIf(fblnRevisaPermiso(vglngNumeroLogin, 302, "C"), 1, 0), IIf((IsNull(rsEntradaSalida!chrfoliofactura) Or Trim(rsEntradaSalida!chrfoliofactura) = "") And rsEntradaSalida!bitcancelado = 0, 1, 0)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
On Error GoTo NotificaError
    
    Dim vllngNumeroCorte As Long
    Dim vllngCorteGrabando As Long
    Dim rsPvDocumentoCancelado As New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim vltipopago As String
    Dim vllngIDComprobante As Long
    Dim rsGnCom As ADODB.Recordset
    Dim blnCFDI As Boolean
    Dim lngCveFormato As Long
    Dim lngNumPagoSalida As Long
    
    blnCFDI = False
    '-------------------------------------------------------'
    '   Osea que la entrada o salida no este ya facturado   '
    '-------------------------------------------------------'
    If Not IsNull(rsEntradaSalida!chrfoliofactura) And Trim(rsEntradaSalida!chrfoliofactura) <> "" Then Exit Sub
    
      If grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento) = "S" And optEntradaDinero.Value Then
        MsgBox "No es posible realizar la cancelación, debido a que el Tipo de Movimiento es 'Entrada de Dinero'.", vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub ''
      End If
    
      If grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento) = "E" And optSalidaDinero.Value Then
        MsgBox "No es posible realizar la cancelación, debido a que el Tipo de Movimiento es 'Salida de Dinero'.", vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub ''
      End If
       
    If optSalidaDinero.Value Then
        vlstrSentencia = "select * from pvPagoDevolucionPaciente where intnumsalida = " & rsEntradaSalida!intNumSalida & " and bitpagado = 1 and chrestado = 'P'"
        Set rs = frsRegresaRs(vlstrSentencia)
       If rs.RecordCount > 0 Then
            MsgBox "No es posible realizar la cancelación, debido a que ya se generó un cheque o transferencia para devolver el dinero.", vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
       End If
     End If

    
    If optEntradaDinero.Value And optDeudor.Value Then
        vlstrSentencia = "select * from cpComprobacion where intnumpago = " & rsEntradaSalida!intNumPago & " and trim(vchestatus) = 'ACTIVO'"
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount > 0 Then
            MsgBox "No es posible cancelar la entrada de dinero, está relacionada con una comprobación de gastos.", vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
        End If
    End If
    '-----------------------'
    '   Persona que Graba   '
    '-----------------------'
    If vlblnCancelaErrorTim = False Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    Else
        vllngPersonaGraba = rsEntradaSalida!intCveEmpleado
    End If
    If vllngPersonaGraba = 0 And vlblnCancelaErrorTim = False Then Exit Sub
        If vlblnActivaMotivo Then
            frmMotivosCancelacion.blnActivaUUID = False
            frmMotivosCancelacion.Show vbModal, Me
            If vgMotivoCancelacion = "" Then Exit Sub
        End If
    
    '-----------------------------------------------------'
    '   Verificar que el documento sea del Departamento   '
    '-----------------------------------------------------'
    If rsEntradaSalida!SMIDEPARTAMENTO <> vgintNumeroDepartamento Then
        If vlblnCancelarRecibosOtroDepto = True Then
            'El documento no pertenece a este departamento, ¿Desea continuar?
            If MsgBox(SIHOMsg(303), vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "No es posible cancelar el documento, debido a que no pertenece a este departamento.", vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
        End If
    End If
        
    
    
    
    '----- NUEVO CFDI PAGOS
    If optEntradaDinero.Value Then
        If Not IsNull(rsEntradaSalida!INTIDCOMPROBANTE) Then
            vllngIDComprobante = 0
            lngNumPagoSalida = rsEntradaSalida!intNumPago
            Set rsGnCom = frsRegresaRs("SELECT GNCOMPROBANTEFISCALDIGITAL.INTIDCOMPROBANTE From GNCOMPROBANTEFISCALDIGITAL INNER JOIN PVPAGO ON GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVPAGO.INTNUMPAGO WHERE GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'AN' AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = " & lngNumPagoSalida)
            If Not rsGnCom.EOF Then
                vllngIDComprobante = rsGnCom!INTIDCOMPROBANTE
                If vllngIDComprobante = rsEntradaSalida!INTIDCOMPROBANTE Then
                    '------------------------------------'
                    ' Cancelar el CFDi por medio del PAC '
                    '------------------------------------'
                    blnCFDI = True
                    If Not fblnCancelaCFDi(lngNumPagoSalida, "AN") Then
                       If vlstrMensajeErrorCancelacionCFDi <> "" Then
                            MsgBox vlstrMensajeErrorCancelacionCFDi, vbOKOnly + vbCritical, "Mensaje"
                            txtMovimientoPaciente.SetFocus
                            'EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                       End If
                    End If
                        If vlblnActivaMotivo Then
                            frsEjecuta_SP CStr(vllngIDComprobante) & "|'" & vgMotivoCancelacion & "'", "SP_CNUPDCANCELACOMPROBPAGO"
                        Else
                            frsEjecuta_SP CStr(vllngIDComprobante) & "|", "SP_CNUPDCANCELACOMPROBPAGO"
                        End If
                End If
            End If
            rsGnCom.Close
        End If
    End If
    '----- NUEVO CFDI PAGOS
    
    '---------------------------'
    '   Inicio de TRANSACCIÓN   '
    '---------------------------'
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '---------------------------------------'
    '   Bloqueo de la cuenta del paciente   '
    '---------------------------------------'
    If optSocio.Value = False And optDeudor.Value = False Then If Not fblnBloqueoCuenta() Then Exit Sub
    
    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    
    
    '------------------------------------------'
    '   Estatus del Corte y numero del corte   '
    '------------------------------------------'
    vllngCorteGrabando = 1
    frsEjecuta_SP CStr(vllngNumeroCorte) & "|" & "Grabando", "Sp_PvUpdEstatusCorte", True, vllngCorteGrabando
    If vllngCorteGrabando <> 2 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Sub
    End If
    
    '1.- Se cancela el pago
    If optEntradaDinero.Value Then
        vlstrSentencia = "UPDATE PvPago SET bitCancelado = 1 WHERE intNumPago = " & rsEntradaSalida!intNumPago
    Else
        vlstrSentencia = "UPDATE PvSalidaDinero SET bitCancelado = 1 WHERE intNumSalida = " & rsEntradaSalida!intNumSalida
    End If
    pEjecutaSentencia vlstrSentencia
    
    '2.- Póliza de cancelacion y registro en corte
    
'    If rsEntradaSalida.State <> 0 Then
        vlstrSentencia = "SELECT PvConceptoPago.* FROM PvConceptoPago WHERE intNumConcepto = " & str(rsEntradaSalida!intNumConcepto)
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount <> 0 Then
            If rs!bitpagocancelafactura = 1 Then
                vltipopago = "PA"
                vgstrParametrosSP = Trim(rsEntradaSalida!CHRFOLIORECIBO) & "|" & vllngPersonaGraba & "|" & "PAF" & "|" & CStr(rsEntradaSalida!intNumCorte) & "|" & CStr(vllngNumeroCorte)
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaDoctoCorte"
            Else
                vltipopago = "NO"
                vgstrParametrosSP = Trim(rsEntradaSalida!CHRFOLIORECIBO) & "|" & vllngPersonaGraba & "|" & IIf(optEntradaDinero.Value, IIf(vlintIncluidoenFactura = 1, "REP", "REF"), IIf(vlintIncluidoenFactura = 1, "SDP", "SDF")) & "|" & CStr(rsEntradaSalida!intNumCorte) & "|" & CStr(vllngNumeroCorte)
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaDoctoCorte"
            End If
        End If
'    End If
    
    '3.- Limpiar el control de aseguradora
    If grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "DE" _
    Or grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CO" _
    Or grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CP" _
    Or grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CA" Then
        vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & str(lintCveEmpresa) & "|" & grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) & "|" & " "
        frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCONTROLASEGURADORA"
    End If
    
    '------------------------------------------------------------------------------------'
    '   Recordset tipo tabla para guardar documentos cancelados (PvDocumentoCancelado)   '
    '------------------------------------------------------------------------------------'
    vlstrSentencia = "SELECT * FROM PvDocumentoCancelado WHERE chrFolioDocumento = ''"
    Set rsPvDocumentoCancelado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    '4.- Se registra en documentos cancelados
    With rsPvDocumentoCancelado
        .AddNew
        !chrFolioDocumento = rsEntradaSalida!CHRFOLIORECIBO
        !chrTipoDocumento = IIf(optEntradaDinero.Value, "RE", "SD")
        !SMIDEPARTAMENTO = vgintNumeroDepartamento
        !intEmpleado = vllngPersonaGraba
        !dtmfecha = fdtmServerFecha
        .Update
    End With
    
    '5.- Se cancela el movimiento de la forma de pago - (CR) Agregado para caso 6894 (Modificado)-
    If optEntradaDinero.Value Then
        pCancelaMovimiento rsEntradaSalida!intNumPago, Trim(rsEntradaSalida!CHRFOLIORECIBO), IIf(vltipopago = "PA", "PA", "RE"), rsEntradaSalida!intNumCorte, vllngNumeroCorte
    Else
        pCancelaMovimiento rsEntradaSalida!intNumSalida, Trim(rsEntradaSalida!CHRFOLIORECIBO), "SD", rsEntradaSalida!intNumCorte, vllngNumeroCorte
    End If
      
    If optSalidaDinero.Value Then
        'Se cambia el estado a cancelado del pago por el cual se generaría un cheque o transferencia
        'en caso de haber utilizado la forma de pago "DEVOLUCIONES A PACIENTE POR CUENTAS POR PAGAR"
        vlstrSentencia = "Update pvPagoDevolucionPaciente set chrEstado = 'C' where intnumpago = " & rsEntradaSalida!intNumSalida
        vgstrParametrosSP = rsEntradaSalida!intNumSalida & "|" & 0 & "|" & "C"
        frsEjecuta_SP vgstrParametrosSP, "sp_cpUpdDevolucionPacBanco", True
    End If
    
    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "ENTRADAS Y SALIDAS DE DINERO", CStr(rsEntradaSalida!CHRFOLIORECIBO))
    
    pLiberaCorte vllngNumeroCorte
    
    If optSocio.Value = False Then pLiberaCuenta
    
    EntornoSIHO.ConeccionSIHO.CommitTrans
    If vlblnCancelaErrorTim = False Then
        'Asegúrese de que la impresora esté lista y  presione aceptar.
        MsgBox SIHOMsg(343), vbInformation, "Mensaje"
        If blnCFDI Then
            lngCveFormato = 1
            frsEjecuta_SP vgintNumeroDepartamento & "|0|0|T", "fn_PVSelFormatoFactura2", True, lngCveFormato
            fblnImprimeComprobanteDigital lngNumPagoSalida, "AN", "I", lngCveFormato, 0
        Else
            cmdPrint_Click
        End If
        txtMovimientoPaciente.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete_Click"))
End Sub

Private Sub cmdEnd_Click()
On Error GoTo NotificaError
    
    grdPagos.Row = grdPagos.Rows - 1
    pMuestraPago CLng(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)), grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento)
    pHabilita 1, 1, 1, 1, 1, 0, IIf(fblnRevisaPermiso(vglngNumeroLogin, 302, "C"), 1, 0), IIf((IsNull(rsEntradaSalida!chrfoliofactura) Or Trim(rsEntradaSalida!chrfoliofactura) = "") And rsEntradaSalida!bitcancelado = 0, 1, 0)
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
On Error GoTo NotificaError

    SSTabPagos.Tab = 1
    cmdCargar.SetFocus
    
    If optSocio.Value = True Then
        chkSocios.Value = 1
    Else
        chkSocios.Value = 0
    End If
    
    vlnblnLocate = True
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub pLimpiaGridConsulta()
On Error GoTo NotificaError

    With grdPagos
        .Redraw = False
        .Visible = False
        .Rows = 2
        .Clear
        .Cols = cintColumnas
        .FixedCols = 1
        .FixedRows = 1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaGridConsulta"))
End Sub

Private Sub pConfiguraGridConsulta()
On Error GoTo NotificaError

    With grdPagos
        .FormatString = "|Id|Fecha|Tipo|Folio|Cuenta|Paciente|Cantidad|Moneda|Empleado que registró|Empleado que canceló|Departamento|bitCancelado|TipoMovimiento"
        
        .ColWidth(0) = 100
        .ColWidth(cintColIdEntradaSalida) = 0
        .ColWidth(cintColFecha) = 1100
        .ColWidth(cintColTipo) = 1000
        .ColWidth(cIntColFolio) = 1000
        .ColWidth(cintColCuenta) = 1600
        .ColWidth(cintColPaciente) = 3000
        .ColWidth(cintColCantidad) = 1000
        .ColWidth(cintColMoneda) = 850
        .ColWidth(cintColEmpleado) = 2500
        .ColWidth(cintColEmpleadoCancela) = 2500
        .ColWidth(cintColDepartamento) = 2000
        .ColWidth(cintColbitCancelado) = 0
        .ColWidth(cintColTipoMovimiento) = 0
        .ColWidth(cintColPendienteCancelarSAT) = 0
        
        .ColAlignment(cintColFecha) = flexAlignLeftCenter
        .ColAlignment(cintColTipo) = flexAlignLeftCenter
        .ColAlignment(cIntColFolio) = flexAlignLeftCenter
        .ColAlignment(cintColCuenta) = flexAlignRightCenter
        .ColAlignment(cintColPaciente) = flexAlignLeftCenter
        .ColAlignment(cintColCantidad) = flexAlignRightCenter
        .ColAlignment(cintColMoneda) = flexAlignLeftCenter
        .ColAlignment(cintColEmpleado) = flexAlignLeftCenter
        .ColAlignment(cintColEmpleadoCancela) = flexAlignLeftCenter
        .ColAlignment(cintColDepartamento) = flexAlignLeftCenter
        
        .ColAlignmentFixed(cintColFecha) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTipo) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColFolio) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCuenta) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColPaciente) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCantidad) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColMoneda) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColEmpleado) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColEmpleadoCancela) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDepartamento) = flexAlignCenterCenter
        
        .Redraw = True
        .Visible = True
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridConsulta"))
End Sub

Private Sub cmdNext_Click()
On Error GoTo NotificaError
    
    If grdPagos.Row < grdPagos.Rows - 1 Then
        grdPagos.Row = grdPagos.Row + 1
    End If
    pMuestraPago CLng(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)), grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento)
    pHabilita 1, 1, 1, 1, 1, 0, IIf(fblnRevisaPermiso(vglngNumeroLogin, 302, "C"), 1, 0), IIf((IsNull(rsEntradaSalida!chrfoliofactura) Or Trim(rsEntradaSalida!chrfoliofactura) = "") And rsEntradaSalida!bitcancelado = 0, 1, 0)

    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Public Sub cmdPrint_Click()
On Error GoTo NotificaError

    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    Dim alstrParametros(0) As String
    Dim strFechaHoy As String
    Dim ObjRs As New ADODB.Recordset
    Dim objNombreHospital As String
    Dim objRFCHospital As String
    Dim objDireccionHospital As String
    Dim objColoniaHospital As String
    
    Set ObjRs = frsRegresaRs("select * from CNEmpresaContable where tnyClaveEmpresa = " & vgintClaveEmpresaContable)
    
    If Not ObjRs.EOF Then
       objNombreHospital = Trim(IIf(IsNull(ObjRs!vchNombre), "", ObjRs!vchNombre))
       objRFCHospital = Trim(IIf(IsNull(ObjRs!vchRFC), "", ObjRs!vchRFC))
       objDireccionHospital = IIf(IsNull(ObjRs!vchCalle), "", Trim(ObjRs!vchCalle)) & IIf(IsNull(ObjRs!VCHNUMEROEXTERIOR), "", " No. " & Trim(ObjRs!VCHNUMEROEXTERIOR)) & IIf(IsNull(ObjRs!VCHNUMEROINTERIOR), "", " Int. " & Trim(ObjRs!VCHNUMEROINTERIOR))
       objColoniaHospital = Trim(IIf(IsNull(ObjRs!VCHCOLONIA), "", ObjRs!VCHCOLONIA))
    End If
    ObjRs.Close
          
    If txtCantidadenLetras.Text = "" Then
        txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), IIf(optPesos.Value, "pesos", "dólares"), "")
    End If
          
    vlstrx = txtFolioRecibo.Text
    vlstrx = vlstrx & "|" & txtCantidadenLetras.Text
    vlstrx = vlstrx & "|" & IIf(optEntradaDinero.Value, "E", "S")
    vgstrParametrosSP = txtFolioRecibo.Text & "|" & txtCantidadenLetras.Text & "|" & IIf(optEntradaDinero.Value, "E", "S") & "|" & Trim(objNombreHospital) & "|" & Trim(objRFCHospital) & "|" & Trim(objDireccionHospital) & "|" & Trim(objColoniaHospital)
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptReciboPago")
    
    If rsReporte.RecordCount > 0 Then
        strFechaHoy = Trim(vgstrCiudadCH) & " " & Trim(vgstrEstadoCH) & ", A " & str(Day(rsReporte!dtmfecha)) & " DE " & UCase(fstrMesLetra(Month(rsReporte!dtmfecha), True)) & " DEL " & str(Year(rsReporte!dtmfecha))
    
        'Instancia el reporte de acuerdo a los formatos capturados
        pInicializaReporte
        
        vgrptReporte.DiscardSavedData
        
        alstrParametros(0) = "FechaHoy;" & strFechaHoy
        
        pCargaParameterFields alstrParametros, vgrptReporte
        
        pImprimeReporte vgrptReporte, rsReporte, "I", "Recibo de pago", True, 2
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    If rsReporte.State <> adStateClosed Then rsReporte.Close
    
    If Not lblnManipulacion Then
        OptTipoPaciente(0).Value = True
        pLimpia
        txtMovimientoPaciente.SetFocus
    Else
        Me.Visible = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    
    Dim vlstrFolioDocumento As String
    Dim vllngNumeroCorte As Long
    Dim vlintNumeroFormas As Integer
    Dim vldblTipoCambio As Double
    Dim vllngFoliosFaltantes As Long
    Dim vldblCantidadPagar As Double
    Dim vllngCorteGrabando As Long
    Dim SQL As String
    Dim lngCantidadValida As Long
    Dim vlrsCalculaPagos As New ADODB.Recordset
    Dim rsPvDetalleCorte As New ADODB.Recordset
    Dim dblCantidadConcepto As Double
    Dim dblCantidadIVA As Double
    Dim lngNumPagoSalida As Long
    Dim intcontador As Integer
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    Dim strParametrosSP As String
    Dim rsFormaPago As ADODB.Recordset
    Dim rsSocio As ADODB.Recordset
    Dim vllngPagoUno As Boolean
    Dim ObjRs As New ADODB.Recordset
    
    Dim strSql As String
    Dim vllngNumDetalleCorte As Long
    
    Dim dblTipoCambioCompra As Double 'Tipo cuando se hizo una transferencia a un banco con moneda dlls
    Dim vllngIdKardex As Long 'Id. del registro insertado en el kardex
    Dim vldblComisionIvaBancaria As Double
    
    Dim blnGenerarComprobante As Boolean
    Dim lngCveFormato As Long
    Dim intResultado As Integer
    
    If Not fblnRevisaPermiso(vglngNumeroLogin, 302, "E") Then Exit Sub
    
    '-------------------------------------------------------------------'
    '   Valida si la cuenta se encuentra bloqueada por trabajo social   '
    '-------------------------------------------------------------------'
    If optSocio.Value = False And optDeudor.Value = False Then
        If fblnCuentaBloqueada(Trim(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")) Then
            'No se puede realizar ésta operación. La cuenta se encuentra bloqueada por trabajo social.
            MsgBox SIHOMsg(662), vbCritical, "Mensaje"
            Exit Sub
        End If
    End If
    
    '---------------------------------------'
    '   Datos validos para grabar el PAGO   '
    '---------------------------------------'
    If Not fblnDatosValidos() Then Exit Sub
    
    If fblnGenerarComprobante And Not optDeudor.Value Then
        If MsgBox("¿Desea comprobante fiscal?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            blnGenerarComprobante = True
            pDatosFiscales vlrfcPaciente
            vlstrDFRFC = frmDatosFiscales.vgstrRFC
            vlstrDFNombre = Replace(frmDatosFiscales.vgstrNombre, "'", "''")
            vlstrDFDireccion = frmDatosFiscales.vgstrDireccion
            vlstrDFNumExterior = frmDatosFiscales.vgstrNumExterior
            vlstrDFNumInterior = frmDatosFiscales.vgstrNumInterior
            vlBitDFExtranjero = frmDatosFiscales.vgBitExtranjero
            lngCveCiudad = frmDatosFiscales.llngCveCiudad
            vlstrDFTelefono = frmDatosFiscales.vgstrTelefono
            vlstrDFColonia = frmDatosFiscales.vgstrColonia
            vlstrDFCodigoPostal = frmDatosFiscales.vgstrCP
            vlintUsoCFDI = frmDatosFiscales.vgintUsoCFDI
            vlstrRegimenFiscal = frmDatosFiscales.vlstrRegimenFiscal
            Unload frmDatosFiscales
            If vlstrDFRFC = "" Then
                Exit Sub
            End If
        Else
            blnGenerarComprobante = False
        End If
    Else
        blnGenerarComprobante = False
    End If
    
    '-----------------------'
    '   Persona que graba   '
    '-----------------------'
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    '---------------------------------------'
    '   Tipo de cambio y cantidad a pagar   '
    '---------------------------------------'
    vldblTipoCambio = vldblTipoCambioVenta
    If optDolares.Value Then
        vldblCantidadPagar = Val(Format(txtCantidad.Text, "")) * vldblTipoCambio
    Else
        vldblCantidadPagar = Val(Format(txtCantidad.Text, ""))
    End If
    
    '-----------------------------------------------------------------------------------------------'
    '   Si es una Devolución y este paciente tiene o no pagos, como para que se regrese el dinero   '
    '-----------------------------------------------------------------------------------------------'
    If optSalidaDinero.Value Then
        lngCantidadValida = 1
        vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & "|" & str(vldblCantidadPagar)
        frsEjecuta_SP vgstrParametrosSP, "SP_PVSELCANTIDADVALIDASALIDA", True, lngCantidadValida
        
        If lngCantidadValida <> 1 Then
            'No se puede realizar la salida, la cantidad excede los pagos realizados por el paciente.
            MsgBox SIHOMsg(CInt(lngCantidadValida)), vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
        End If
    End If
    
    '--------------------'
    '   Formas de pago   '
    '--------------------'
    If optEntradaDinero.Value Then
        ' (CR) - Modificado para caso 6894: Que se muestren las formas de pago por transferencias SOLO cuando es entrada de dinero '
        If Not fblnFormasPagoPos(aFormasPago(), vldblCantidadPagar, True, vldblTipoCambio, False, 0, "", vlrfcPaciente, False, False, True, "frmEntradaSalidaDinero") Then Exit Sub
    Else
        If Not fblnFormasPagoPos(aFormasPago(), vldblCantidadPagar, True, vldblTipoCambio, False, 0, "", vlrfcPaciente, False, False, False, "frmEntradaSalidaDinero-S") Then Exit Sub
    End If
        
    '--------------------------------------------------------------------------------'
    '   Recordset tipo tabla para guardar movimientos en el corte (PvDetalleCorte)   '
    '--------------------------------------------------------------------------------'
    vlstrSentencia = "SELECT * FROM PvDetalleCorte WHERE intConsecutivo = -1"
    Set rsPvDetalleCorte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        
    '---------------------------'
    '   Inicio de TRANSACCIÓN   '
    '---------------------------'
    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    '-----------------------'
    '   Bloqueo de cuenta   '
    '-----------------------'
    If optSocio.Value = False And optDeudor.Value = False Then
        If Not fblnBloqueoCuenta() Then
            Exit Sub
        End If
    End If
    
    '---------------------'
    '   Número de CORTE   '
    '---------------------'
    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    
    '---------------------------------------'
    '   Estatus de "GRABANDO" en el Corte   '
    '---------------------------------------'
    vllngCorteGrabando = 1
    frsEjecuta_SP vllngNumeroCorte & "|Grabando", "Sp_PvUpdEstatusCorte", True, vllngCorteGrabando
    If vllngCorteGrabando <> 2 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Sub
    End If
    
    '-----------------------'
    '   Control de FOLIOS   '
    '-----------------------'
    vllngFoliosFaltantes = 0
    
    pCargaArreglo alstrParametrosSalida, vllngFoliosFaltantes & "|" & adInteger & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
    frsEjecuta_SP IIf(optEntradaDinero.Value, "RE", "SD") & "|" & vgintNumeroDepartamento & "|1", "Sp_GnFolios", , , alstrParametrosSalida
    pObtieneValores alstrParametrosSalida, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
    '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
    strSerie = Trim(strSerie)
    vlstrFolioDocumento = strSerie & strFolio
    
    If Trim(vlstrFolioDocumento) = "0" Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        Exit Sub
    End If
    txtFolioRecibo.Text = vlstrFolioDocumento
        
    '------------------------------------------------------------------'
    '   Graba la Entrada(PvPago) o Salida de Dinero (PvSalidaDinero)   '
    '------------------------------------------------------------------'
    lblnGrabo = True
    If optEntradaDinero.Value Then
        If optSocio.Value Then
            strParametrosSP = txtMovimientoPaciente & "|" & -1
            Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_PVSELSOCIOS")
            If rsSocio.RecordCount = 0 Then txtMovimientoPaciente.SetFocus
                        
            vgstrParametrosSP = vlstrFolioDocumento & "|" & _
                                CStr(rsSocio!intcvesocio) & "|" & _
                                "S" & "|" & _
                                fstrFechaSQL(fdtmServerFecha, , True) & "|" & _
                                txtPersona.Text & "|" & _
                                Format(txtCantidad.Text, "############.00") & "|" & _
                                IIf(optPesos.Value, 1, 0) & "|" & _
                                CStr(IIf(optPesos.Value, 0, vldblTipoCambio)) & "|" & _
                                CStr(vgintNumeroDepartamento) & "|" & _
                                CStr(vllngPersonaGraba) & "|" & _
                                CStr(vllngNumeroCorte) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto) & "|" & _
                                Trim(Replace(txtComentario.Text, vbCrLf, "")) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) & "|" & _
                                IIf(vlblnValidarDesglosarIVA, 1, 0)
            rsSocio.Close
        ElseIf optDeudor.Value Then
            vgstrParametrosSP = vlstrFolioDocumento & "|0|" & _
                                "D" & "|" & _
                                fstrFechaSQL(fdtmServerFecha, , True) & "|" & _
                                txtPersona.Text & "|" & _
                                Format(txtCantidad.Text, "############.00") & "|" & _
                                IIf(optPesos.Value, 1, 0) & "|" & _
                                CStr(IIf(optPesos.Value, 0, vldblTipoCambio)) & "|" & _
                                CStr(vgintNumeroDepartamento) & "|" & _
                                CStr(vllngPersonaGraba) & "|" & _
                                CStr(vllngNumeroCorte) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto) & "|" & _
                                Trim(Replace(txtComentario.Text, vbCrLf, "")) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) & "|" & _
                                IIf(vlblnValidarDesglosarIVA, 1, 0)
        Else
            vgstrParametrosSP = vlstrFolioDocumento & "|" & _
                                txtMovimientoPaciente.Text & "|" & _
                                IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & _
                                fstrFechaSQL(fdtmServerFecha, , True) & "|" & _
                                txtPersona.Text & "|" & _
                                Format(txtCantidad.Text, "############.00") & "|" & _
                                IIf(optPesos.Value, 1, 0) & "|" & _
                                CStr(IIf(optPesos.Value, 0, vldblTipoCambio)) & "|" & _
                                CStr(vgintNumeroDepartamento) & "|" & _
                                CStr(vllngPersonaGraba) & "|" & _
                                CStr(vllngNumeroCorte) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto) & "|" & _
                                Trim(Replace(txtComentario.Text, vbCrLf, "")) & "|" & _
                                grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) & "|" & _
                                IIf(vlblnValidarDesglosarIVA, 1, 0)
        End If
       
        lngNumPagoSalida = 1
        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPAGO", True, lngNumPagoSalida, , , True
        
        If lngNumPagoSalida = 1 Then ' validamos si se tiene 1 por que trono o por si inserto el registro 1
            Set ObjRs = frsRegresaRs("Select max(intnumpago) from pvpago", adLockOptimistic)
                If ObjRs.RecordCount = 0 Then
                   lngNumPagoSalida = -1
                Else
                   If ObjRs.Fields(0) <> 1 Then lngNumPagoSalida = -1
                End If
        End If
        If blnGenerarComprobante Then
            pEjecutaSentencia "update PVPago set chrRFC = '" & vlstrDFRFC & "'" & _
            ", vchNombre = '" & vlstrDFNombre & "'" & _
            ", vchCalle = '" & vlstrDFDireccion & "'" & _
            ", vchNumeroExterior = '" & vlstrDFNumExterior & "'" & _
            ", vchNumeroInterior = '" & vlstrDFNumInterior & "'" & _
            ", vchTelefono = '" & vlstrDFTelefono & "'" & _
            ", vchColonia = '" & vlstrDFColonia & "'" & _
            ", vchCodigoPostal = '" & vlstrDFCodigoPostal & "'" & _
            ", intCveCiudad = " & lngCveCiudad & _
            ", intCveUsoCFDI = " & vlintUsoCFDI & _
            " where intNumPago = " & lngNumPagoSalida
        End If
    Else
        vgstrParametrosSP = vlstrFolioDocumento & "|" & _
                            txtMovimientoPaciente.Text & "|" & _
                            IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & _
                            fstrFechaSQL(fdtmServerFecha, fdtmServerHora, True) & "|" & _
                            txtPersona.Text & "|" & _
                            Format(txtCantidad.Text, "############.00") & "|" & _
                            IIf(optPesos.Value, 1, 0) & "|" & _
                            CStr(IIf(optPesos.Value, 0, vldblTipoCambio)) & "|" & _
                            CStr(vgintNumeroDepartamento) & "|" & _
                            CStr(vllngPersonaGraba) & "|" & _
                            CStr(vllngNumeroCorte) & "|" & _
                            grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto) & "|" & _
                            Trim(Replace(txtComentario.Text, vbCrLf, ""))
        lngNumPagoSalida = 1
        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSSALIDADINERO", True, lngNumPagoSalida, , , True
        
        
        If lngNumPagoSalida = 1 Then ' validamos si se tiene 1 por que trono o por si inserto el registro 1
            Set ObjRs = frsRegresaRs("Select max(intnumsalida) from pvSalidadinero", adLockOptimistic)
                If ObjRs.RecordCount = 0 Then
                   lngNumPagoSalida = -1
                Else
                   If ObjRs.Fields(0) <> 1 Then lngNumPagoSalida = -1
                End If
        End If
        
    End If
    
    If lngNumPagoSalida > -1 Then ' que continue con el proceso si se inserto el pago o la salida de dinero.
            '---------------------------------------------------------'
            ' Si es entrada de dinero abono a la cuenta del concepto, '
            ' si es salida, cargo a la cuenta del concepto            '
            '---------------------------------------------------------'
            
            ' Para validar si el parametro "Desglosar IVA en pagos" está activo '
            If vlblnValidarDesglosarIVA Then
            '********** >>SI<< desglosar IVA **********'
                dblCantidadConcepto = Val(Format(txtCantidad.Text, "############.00"))
                dblCantidadIVA = dblCantidadConcepto - dblCantidadConcepto / (1 + (vgdblCantidadIvaGeneral / 100))
                dblCantidadConcepto = dblCantidadConcepto - dblCantidadIVA
                
                frsEjecuta_SP CStr(vllngNumeroCorte) & "|" & vlstrFolioDocumento & "|" & _
                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                              grdConceptos.TextMatrix(grdConceptos.Row, lintColNumeroCuenta) & "|" & _
                              IIf(optPesos.Value, dblCantidadConcepto, dblCantidadConcepto * vldblTipoCambio) & "|" & _
                              IIf(optEntradaDinero.Value, 0, 1) & "|" & "REC", "Sp_PvInsPvCortePoliza", True
        
                frsEjecuta_SP CStr(vllngNumeroCorte) & "|" & vlstrFolioDocumento & "|" & _
                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                              glngCtaIVACobrado & "|" & _
                              IIf(optPesos.Value, dblCantidadIVA, dblCantidadIVA * vldblTipoCambio) & "|" & _
                              IIf(optEntradaDinero.Value, 0, 1) & "|" & "REC", "Sp_PvInsPvCortePoliza", True
            Else
            '********** >>NO<< desglosar IVA **********'
                dblCantidadConcepto = Val(Format(txtCantidad.Text, "############.00"))
                dblCantidadIVA = dblCantidadConcepto - dblCantidadConcepto / (1 + (vgdblCantidadIvaGeneral / 100))
                
                frsEjecuta_SP CStr(vllngNumeroCorte) & "|" & vlstrFolioDocumento & "|" & _
                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                              grdConceptos.TextMatrix(grdConceptos.Row, lintColNumeroCuenta) & "|" & _
                              IIf(optPesos.Value, dblCantidadConcepto, dblCantidadConcepto * vldblTipoCambio) & "|" & _
                              IIf(optEntradaDinero.Value, 0, 1) & "|" & "REC", "Sp_PvInsPvCortePoliza", True
            End If
            
            If optEntradaDinero.Value Then
                'se actualiza clave de impuesto y cuenta del concepto de entrada
                vgstrParametrosSP = lngNumPagoSalida & "|" & grdConceptos.TextMatrix(grdConceptos.Row, lintColNumeroCuenta)
                frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDPAGO"
            End If
            
            '-----------------------------------'
            '   Afecta corte (PvDetalleCorte)   '
            '-----------------------------------'
            If optSocio.Value = False Then
                vlintNumeroFormas = UBound(aFormasPago(), 1)
                
                For intcontador = 0 To vlintNumeroFormas
                    With rsPvDetalleCorte
                        If aFormasPago(intcontador).vlintNumFormaPago <> -9 Then
                            .AddNew
                            !intNumCorte = vllngNumeroCorte
                            !dtmFechahora = fdtmServerFecha + fdtmServerHora
                            !chrFolioDocumento = vlstrFolioDocumento
                            !chrTipoDocumento = IIf(optEntradaDinero.Value, "RE", "SD")
                            !intFormaPago = aFormasPago(intcontador).vlintNumFormaPago
                            If aFormasPago(intcontador).vldblTipoCambio = 0 Then
                                !mnyCantidadPagada = aFormasPago(intcontador).vldblCantidad
                            Else
                                !mnyCantidadPagada = aFormasPago(intcontador).vldblDolares
                            End If
                            !mnytipocambio = aFormasPago(intcontador).vldblTipoCambio
                            !intfoliocheque = IIf(Trim(aFormasPago(intcontador).vlstrFolio) = "", "0", Trim(aFormasPago(intcontador).vlstrFolio))
                            !intNumCorteDocumento = vllngNumeroCorte
                            .Update
                                                
                            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", rsPvDetalleCorte!intConsecutivo)
                        
                            If Not aFormasPago(intcontador).vlbolEsCredito Then
                                If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                    frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                End If
                            End If
                        'ElseIf aFormasPago(intcontador).vlintNumFormaPago = -9 Then
                        '    vllngCuentaDevolucionPaciente = aFormasPago(intcontador).vllngCuentaContable
                        '    vlstrRFCDevolucionPaciente = aFormasPago(intcontador).vlstrRFC
                        End If
                        
                        'Si es entrada de dinero CARGO a la cuenta de la forma de pago
                        'Si es salida, ABONO a la cuenta de la forma de pago
                        frsEjecuta_SP vllngNumeroCorte & "|" & _
                                      vlstrFolioDocumento & "|" & _
                                      IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                      aFormasPago(intcontador).vllngCuentaContable & "|" & _
                                      IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, aFormasPago(intcontador).vldblDolares * aFormasPago(intcontador).vldblTipoCambio) & "|" & _
                                      IIf(optEntradaDinero.Value, 1, 0) & "|" & "REC", "Sp_PvInsPvCortePoliza", True
                    
                        ' Agregado para caso 8741
                        ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
                        If optEntradaDinero.Value Then
                            If aFormasPago(intcontador).vllngCuentaComisionBancaria <> 0 And aFormasPago(intcontador).vldblCantidadComisionBancaria <> 0 Then
                                 ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                                frsEjecuta_SP vllngNumeroCorte & "|" & _
                                              vlstrFolioDocumento & "|" & _
                                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                              aFormasPago(intcontador).vllngCuentaComisionBancaria & "|" & _
                                              aFormasPago(intcontador).vldblCantidadComisionBancaria & "|" & _
                                              1 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                                If aFormasPago(intcontador).vldblIvaComisionBancaria <> 0 Then
                                    ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                                    frsEjecuta_SP vllngNumeroCorte & "|" & _
                                                  vlstrFolioDocumento & "|" & _
                                                  IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                                  glngCtaIVAPagado & "|" & _
                                                  aFormasPago(intcontador).vldblIvaComisionBancaria & "|" & _
                                                  1 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                                End If
                                ' Se genera un abono por la cantidad de la comisión bancaria y su iva a la cuenta de la forma de pago
                                frsEjecuta_SP vllngNumeroCorte & "|" & _
                                              vlstrFolioDocumento & "|" & _
                                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                              aFormasPago(intcontador).vllngCuentaContable & "|" & _
                                              (aFormasPago(intcontador).vldblCantidadComisionBancaria + aFormasPago(intcontador).vldblIvaComisionBancaria) & "|" & _
                                              0 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                            End If
                        End If
                    End With
                    
                    '- (CR) Agregado para caso 6894: Pago con Transferencia bancaria -'
                    'If optEntradaDinero.Value Then
                        'If aFormasPago(intContador).lngIdBanco <> 0 Then
                        '    dblTipoCambioCompra = fdblTipoCambio(CDate(mskFecha.Text), "C")
                        '    '- Registra en el libro de bancos la transferencia a cuenta del paciente -'
                        '    vgstrParametrosSP = CStr(aFormasPago(intContador).lngIdBanco) & "|" & fstrFechaSQL(mskFecha.Text, "00:00") & "|" & "TPA" & "|" & aFormasPago(intContador).vldblCantidad / IIf(aFormasPago(intContador).intMoneda = 0, dblTipoCambioCompra, 1) & "|" & "0" & "|" & CStr(lngNumPagoSalida)
                        '    frsEjecuta_SP vgstrParametrosSP, "Sp_CpInsKardexBanco"
                        '    vllngIdKardex = flngObtieneIdentity("Sec_CpKardexBanco", 0)
                        '
                        '    '- Registra en la tabla kárdex del banco de caja -'
                        '    vgstrParametrosSP = CStr(lngNumPagoSalida) & "|" & CStr(vllngIdKardex) & "|" & "RE"
                        '    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsPagoKardexBanco"
                        'End If
                    'End If
                    If Not aFormasPago(intcontador).vlbolEsCredito Then
                        If aFormasPago(intcontador).vlintNumFormaPago <> -9 Then
                            '----- Guardar información de la forma de pago en tabla intermedia -----'
                            vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(mskFecha.Text, Format(fdtmServerHora, "hh:mm:ss")) & "|" & aFormasPago(intcontador).vlintNumFormaPago & "|" & aFormasPago(intcontador).lngIdBanco & "|" & _
                                                IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, aFormasPago(intcontador).vldblDolares) * IIf(optEntradaDinero.Value, 1, -1) & "|" & IIf(aFormasPago(intcontador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(intcontador).vldblTipoCambio & "|" & _
                                                fstrTipoMovimientoForma(aFormasPago(intcontador).vlintNumFormaPago) & "|" & IIf(optEntradaDinero.Value, "RE", "SD") & "|" & lngNumPagoSalida & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(mskFecha.Text, fdtmServerHora) & "|" & "1" & "|" & cgstrModulo
                            frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                        ElseIf aFormasPago(intcontador).vlintNumFormaPago = -9 Then     'Forma de pago "DEVOLUCIONES A PACIENTE POR CUENTAS POR PAGAR"
                            ' Guarda los datos para el pago de la devolución a paciente, que se realizará en cuentas por pagar
                            SQL = "INSERT INTO pvPagoDevolucionPaciente (intNumSalida, mnyImporte, intNumCuentaDevolucion, chrRFCBeneficiario) VALUES(" & lngNumPagoSalida & "," & aFormasPago(intcontador).vldblCantidad & "," & aFormasPago(intcontador).vllngCuentaContable & ",'" & aFormasPago(intcontador).vlstrRFC & "')"
                            pEjecutaSentencia SQL
                        End If
                        
                        ' Agregado para caso 8741
                        ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                        vldblComisionIvaBancaria = 0
                        If optEntradaDinero.Value Then
                            If aFormasPago(intcontador).vllngCuentaComisionBancaria <> 0 And aFormasPago(intcontador).vldblCantidadComisionBancaria <> 0 Then
                                If aFormasPago(intcontador).vldblTipoCambio = 0 Then
                                     vldblComisionIvaBancaria = (aFormasPago(intcontador).vldblCantidadComisionBancaria + aFormasPago(intcontador).vldblIvaComisionBancaria) * -1
                                Else
                                     vldblComisionIvaBancaria = (aFormasPago(intcontador).vldblCantidadComisionBancaria + aFormasPago(intcontador).vldblIvaComisionBancaria) / aFormasPago(intcontador).vldblTipoCambio * -1
                                End If
                                vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(mskFecha.Text, Format(fdtmServerHora, "hh:mm:ss")) & "|" & aFormasPago(intcontador).vlintNumFormaPago & "|" & aFormasPago(intcontador).lngIdBanco & "|" & _
                                                    vldblComisionIvaBancaria & "|" & IIf(aFormasPago(intcontador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(intcontador).vldblTipoCambio & "|" & _
                                                    "CBA" & "|" & IIf(optEntradaDinero.Value, "RE", "SD") & "|" & lngNumPagoSalida & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(mskFecha.Text, fdtmServerHora) & "|" & "1" & "|" & cgstrModulo
                                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                            End If
                        End If
                    End If
                Next intcontador
                
                'Afectar el control de aseguradora:
                If optEntradaDinero.Value And _
                   (grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "DE" Or _
                    grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CO" Or _
                    grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CA" Or _
                    grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) = "CP") Then
                    vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & str(lintCveEmpresa) & "|" & grdConceptos.TextMatrix(grdConceptos.Row, lintColTipoConcepto) & "|" & Trim(vlstrFolioDocumento)
                    frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCONTROLASEGURADORA"
                End If
                
                SQL = "DELETE FROM PvTipoPacienteProceso WHERE PvTipoPacienteProceso.intNumeroLogin = " & vglngNumeroLogin & " AND PvTipoPacienteProceso.intProceso = " & enmTipoProceso.pagos
                pEjecutaSentencia SQL
                
                SQL = "INSERT INTO PvTipoPacienteProceso (intNumeroLogin, intProceso, chrTipoPaciente) VALUES(" & vglngNumeroLogin & "," & enmTipoProceso.pagos & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
                pEjecutaSentencia SQL
            Else
                vlintNumeroFormas = UBound(aFormasPago(), 1)
                        
                For intcontador = 0 To vlintNumeroFormas
                    If aFormasPago(intcontador).vlintNumFormaPago <> -9 Then
                        strParametrosSP = CStr(vllngNumeroCorte) _
                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & vlstrFolioDocumento _
                                        & "|" & "RE" _
                                        & "|" & CStr(aFormasPago(intcontador).vlintNumFormaPago) _
                                        & "|" & CStr(aFormasPago(intcontador).vldblCantidad) _
                                        & "|" & CStr(0) _
                                        & "|" & IIf(Trim(aFormasPago(intcontador).vlstrFolio) = "", "0", Trim(aFormasPago(intcontador).vlstrFolio)) _
                                        & "|" & CStr(vllngNumeroCorte)
                        frsEjecuta_SP strParametrosSP, "Sp_PvInsDetalleCorte"
                        
                        vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                        
                        If Not aFormasPago(intcontador).vlbolEsCredito Then
                            If Trim(aFormasPago(intcontador).vlstrRFC) <> "" And Trim(aFormasPago(intcontador).vlstrBancoSAT) <> "" Then
                                frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(intcontador).vlstrRFC) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(intcontador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(intcontador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(intcontador).vldtmFecha))) & "'|'" & Trim(aFormasPago(intcontador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                            End If
                        End If
                    End If
                    
                    ' Cargo a la cuenta de la forma de pago
                    frsEjecuta_SP vllngNumeroCorte & "|" & _
                                  vlstrFolioDocumento & "|" & _
                                  "RE" & "|" & _
                                  aFormasPago(intcontador).vllngCuentaContable & "|" & _
                                  IIf(aFormasPago(intcontador).vldblTipoCambio = 0, aFormasPago(intcontador).vldblCantidad, aFormasPago(intcontador).vldblDolares * aFormasPago(intcontador).vldblTipoCambio) & "|" & _
                                  1 & "|" & "REC", "Sp_PvInsPvCortePoliza", True
                                  
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
                    If optEntradaDinero.Value Then
                        If aFormasPago(intcontador).vllngCuentaComisionBancaria <> 0 And aFormasPago(intcontador).vldblCantidadComisionBancaria <> 0 Then
                             ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                            frsEjecuta_SP vllngNumeroCorte & "|" & _
                                          vlstrFolioDocumento & "|" & _
                                          IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                          aFormasPago(intcontador).vllngCuentaComisionBancaria & "|" & _
                                          aFormasPago(intcontador).vldblCantidadComisionBancaria & "|" & _
                                          1 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                            If aFormasPago(intcontador).vldblIvaComisionBancaria <> 0 Then
                                ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                                frsEjecuta_SP vllngNumeroCorte & "|" & _
                                              vlstrFolioDocumento & "|" & _
                                              IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                              glngCtaIVAPagado & "|" & _
                                              aFormasPago(intcontador).vldblIvaComisionBancaria & "|" & _
                                              1 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                            End If
                            ' Se genera un abono por la cantidad de la comisión bancaria y su iva a la cuenta de la forma de pago
                            frsEjecuta_SP vllngNumeroCorte & "|" & _
                                          vlstrFolioDocumento & "|" & _
                                          IIf(optEntradaDinero.Value, "RE", "SD") & "|" & _
                                          aFormasPago(intcontador).vllngCuentaContable & "|" & _
                                          (aFormasPago(intcontador).vldblCantidadComisionBancaria + aFormasPago(intcontador).vldblIvaComisionBancaria) & "|" & _
                                          0 & "|" & "CBA", "Sp_PvInsPvCortePoliza", True
                        End If
                    End If
                Next intcontador
            End If
            
             pRefacturacionAnteriores
            
            '----- NUEVO CFDI ANTICIPOS
            If blnGenerarComprobante Then
                intResultado = pGeneraCFDI(True, lngNumPagoSalida, strAnoAprobacion, strNumeroAprobacion, strSerie, strFolio)
                If intResultado = 2 Then
                    Exit Sub
                End If
            End If
            '----- NUEVO CFDI ANTICIPOS (FIN)
            
            
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "ENTRADAS Y SALIDAS DE DINERO", vlstrFolioDocumento)
            
            pLiberaCorte vllngNumeroCorte
            
            If optSocio.Value = False Then pLiberaCuenta
                
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            If blnGenerarComprobante Then
                lngCveFormato = 1
                frsEjecuta_SP vgintNumeroDepartamento & "|0|0|T", "fn_PVSelFormatoFactura2", True, lngCveFormato
                fblnImprimeComprobanteDigital lngNumPagoSalida, "AN", "I", lngCveFormato, 0
                If intResultado = 0 Then
                    If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
                    '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                        If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                            pEnviarCFD "AN", lngNumPagoSalida, CLng(vgintClaveEmpresaContable), Trim(vlstrDFRFC), vllngPersonaGraba, Me
                        End If
                    End If
                End If
                txtMovimientoPaciente.SetFocus
            Else
                fraPaciente.Enabled = False
                pHabilita 0, 0, 0, 0, 0, 0, 1, 0
                fraRecibo.Enabled = False
                fraImprimir.Visible = cboFormato.ListCount <> 0 And Not optSalidaDinero.Value
                cmdPrint.SetFocus
            End If
    Else 'por X o Y NO se tiene número de pago o salida de $$$ y se dió RollBack(PUMMMMM!!!!)
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            cmdSave.Enabled = False
            cmdPrint.Enabled = False
            Unload Me
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave_Click"))
    cmdSave.Enabled = False
    cmdPrint.Enabled = False
    Unload Me
End Sub

Private Sub pLiberaCuenta()
On Error GoTo NotificaError
    
    frsEjecuta_SP IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & txtMovimientoPaciente.Text & "|0", "SP_EXUPDCUENTAOCUPADA"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLiberaCuenta"))
End Sub

Private Function fblnBloqueoCuenta() As Boolean
On Error GoTo NotificaError
    
    Dim X As Integer
    Dim vlblnTermina As Boolean
    Dim vlstrBloqueo As String
                
    fblnBloqueoCuenta = False
    vlblnTermina = False
    
    X = 1
    Do While X <= cgintIntentoBloqueoCuenta And Not vlblnTermina
        vlstrBloqueo = fstrBloqueaCuenta(Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E"))
        If vlstrBloqueo = "F" Then
            vlblnTermina = True
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            'La cuenta ya ha sido facturada, no se pudo realizar ningún movimiento.
            MsgBox SIHOMsg(299), vbOKOnly + vbInformation, "Mensaje"
        Else
            If vlstrBloqueo = "O" Then
                If X = cgintIntentoBloqueoCuenta Then
                    vlblnTermina = True
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    'La cuenta esta siendo usada por otra persona, intente de nuevo.
                    MsgBox SIHOMsg(300), vbOKOnly + vbInformation, "Mensaje"
                End If
            Else
                If vlstrBloqueo = "L" Then
                    vlblnTermina = True
                    fblnBloqueoCuenta = True
                End If
            End If
        End If
        X = X + 1
    Loop

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnBloqueoCuenta"))
End Function

Private Function fblnDatosValidos() As Boolean
On Error GoTo NotificaError
    
    Dim vllngCtaContableMedico As Long
    Dim rsCuenta As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    
    fblnDatosValidos = True
    If Trim(txtFolioRecibo.Text) = "" Then
        fblnDatosValidos = False
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        
        If optEntradaDinero.Value = True Then
            optEntradaDinero.SetFocus
        Else
            If optSalidaDinero.Value = True Then
                optSalidaDinero.SetFocus
            End If
        End If
    End If
    
    If fblnDatosValidos And Trim(Me.txtPersona.Text) = "" Then ' el nombre de la persona no puede ser null en PVPAGO
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtPersona.SetFocus
    End If
    
    If fblnDatosValidos And Trim(txtCantidad.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtCantidad.SetFocus
    End If
    
    If fblnDatosValidos And (Not optPesos.Value And Not optDolares.Value) Then
        fblnDatosValidos = False
        'Seleccione el tipo de moneda.
        MsgBox SIHOMsg(296), vbOKOnly + vbInformation, "Mensaje"
        optPesos.SetFocus
    End If
    
    If fblnDatosValidos And optDeudor.Value Then
        Set rsCuenta = frsRegresaRs("select * from pvConceptoPagoEmpresa where intnumconcepto = " & llngCveConceptoDeudor & " and intcveEmpresa = " & vgintClaveEmpresaContable)
        If rsCuenta.RecordCount <> 0 Then
            If Not fblnCuentaAfectable(fstrCuentaContable(rsCuenta!intNumeroCuenta), vgintClaveEmpresaContable) Then
                fblnDatosValidos = False
                'La cuenta seleccionada no acepta movimientos.
                MsgBox SIHOMsg(375) & " " & fstrCuentaContable(rsCuenta!intNumeroCuenta) & " " & fstrDescripcionCuenta(fstrCuentaContable(rsCuenta!intNumeroCuenta), vgintClaveEmpresaContable) & Chr(13) & "Cuenta del concepto.", vbExclamation + vbOKOnly, "Mensaje"
            Else
                grdConceptos.TextMatrix(grdConceptos.Row, lintColNumeroCuenta) = rsCuenta!intNumeroCuenta
            End If
        Else
            fblnDatosValidos = False
            'No se encuentra registrada la cuenta contable del concepto
            MsgBox SIHOMsg(1484), vbOKOnly + vbInformation, "Mensaje"
        End If
    End If
    
    'Para validar si el parametro "Desglosar IVA en pagos" está activo
    strSql = "SELECT BITDESGLOSAIVA FROM PVCONCEPTOPAGO WHERE INTNUMCONCEPTO = " & grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto)
    Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    vlblnValidarDesglosarIVA = IIf(rs!bitdesglosaiva = 0, False, True)
    If fblnDatosValidos And optDeudor.Value And vlblnValidarDesglosarIVA Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(1486), vbOKOnly + vbInformation, "Mensaje"
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub cmdTop_Click()
On Error GoTo NotificaError
    
    grdPagos.Row = 1
    pMuestraPago CLng(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)), grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento)
    pHabilita 1, 1, 1, 1, 1, 0, IIf(fblnRevisaPermiso(vglngNumeroLogin, 302, "C"), 1, 0), IIf((IsNull(rsEntradaSalida!chrfoliofactura) Or Trim(rsEntradaSalida!chrfoliofactura) = "") And rsEntradaSalida!bitcancelado = 0, 1, 0)
            
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub

Private Sub pMuestraPago(lngNumPagoSalida As Long, strTipoMovimiento As String)
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim strParametrosSP As String
    Dim rsSocio As ADODB.Recordset
    Dim strClaveSocio As String
    Dim strSentencia As String
    Dim vlForeColor As Variant
    Dim vlBackColor As Variant
    Dim rsSustituto As ADODB.Recordset
                                         
    vlblnConsulta = True
    lblnConsultaDeudor = False
    lblPagoCancelado.Visible = False

    If strTipoMovimiento = "E" Then
        strSentencia = "SELECT PvPago.* " & _
                       "     , Case " & _
                       "            When PVPENDIENTESCANCELARSAT.INTCOMPROBANTE is null then 'NP' " & _
                       "            Else PVPENDIENTESCANCELARSAT.CHRESTADO " & _
                       "        End PendienteCancelarSAT_NE " & _
                       "  FROM PvPago " & _
                       "       LEFT JOIN PVPENDIENTESCANCELARSAT ON PVPENDIENTESCANCELARSAT.INTCOMPROBANTE = PvPago.INTNUMPAGO AND PVPENDIENTESCANCELARSAT.CHRTIPOCOMPROBANTE = 'AN' " & _
                       " WHERE PvPago.intNumPago = " & str(lngNumPagoSalida)
        Set rsEntradaSalida = frsRegresaRs(strSentencia)
    Else
        Set rsEntradaSalida = frsRegresaRs("SELECT PvSalidaDinero.* FROM PvSalidaDinero WHERE PvSalidaDinero.intNumSalida = " & str(lngNumPagoSalida))
    End If

    If rsEntradaSalida!CHRTIPOPACIENTE <> "S" Then
        If Not fblnDatosPaciente(rsEntradaSalida!INTMOVPACIENTE, rsEntradaSalida!CHRTIPOPACIENTE) Then Exit Sub
    Else
        strParametrosSP = 0 & "|" & CStr(rsEntradaSalida!INTMOVPACIENTE)
        Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_PVSELSOCIOS")
        If rsSocio.RecordCount > 0 Then
            txtPaciente.Text = Trim(rsSocio!vchApellidoPaterno) & " " & Trim(rsSocio!vchApellidoMaterno) & " " & Trim(rsSocio!vchNombre)
            vlrfcPaciente = Trim(Replace(Replace(Replace(rsSocio!vchRFC, "-", ""), "_", ""), " ", ""))
            txtTipoPaciente.Text = "SOCIO"
            strClaveSocio = rsSocio!VCHCLAVESOCIO
        End If
        
        rsSocio.Close
    End If
    
    If rsEntradaSalida!CHRTIPOPACIENTE = "S" Then
        txtMovimientoPaciente.MaxLength = 20
        txtMovimientoPaciente.Text = strClaveSocio
    ElseIf rsEntradaSalida!CHRTIPOPACIENTE = "D" Then
        txtMovimientoPaciente.Text = ""
        lblnConsultaDeudor = True
    Else
        txtMovimientoPaciente.Text = rsEntradaSalida!INTMOVPACIENTE
    End If
    
    '        txtMovimientoPaciente.Text = Str(rsEntradaSalida!intMovPaciente)
    OptTipoPaciente(0).Value = rsEntradaSalida!CHRTIPOPACIENTE = "I"
    OptTipoPaciente(1).Value = rsEntradaSalida!CHRTIPOPACIENTE = "E"
    optSocio.Value = rsEntradaSalida!CHRTIPOPACIENTE = "S"
    optDeudor.Value = rsEntradaSalida!CHRTIPOPACIENTE = "D"
    
    optEntradaDinero.Value = strTipoMovimiento = "E"
    optSalidaDinero.Value = strTipoMovimiento = "S"
    
    txtFolioRecibo.Text = Trim(rsEntradaSalida!CHRFOLIORECIBO)
    mskFecha.Mask = ""
    mskFecha.Text = Format(rsEntradaSalida!dtmfecha, "dd/mm/yyyy")
    mskFecha.Mask = "##/##/####"
    txtPersona.Text = Trim(rsEntradaSalida!chrPersona)
    txtCantidad.Text = FormatCurrency(rsEntradaSalida!MNYCantidad, 2)
    If (rsEntradaSalida!BITPESOS = 1) Then
        optPesos.Value = True
        optDolares.Value = False
    Else
        optPesos.Value = False
        optDolares.Value = True
    End If
    txtCantidadenLetras.Text = fstrNumeroenLetras(rsEntradaSalida!MNYCantidad, IIf(rsEntradaSalida!BITPESOS = 1, "pesos", "dólares"), IIf(rsEntradaSalida!BITPESOS = 1, "M.N.", ""))
    txtComentario.Text = IIf(IsNull(Trim(rsEntradaSalida!vchComentario)), " ", Trim(rsEntradaSalida!vchComentario))
    txtDocumento.Text = Trim(IIf(IsNull(rsEntradaSalida!chrfoliofactura), "", rsEntradaSalida!chrfoliofactura))
    vlintIncluidoenFactura = IIf(IsNull(rsEntradaSalida!bitIncluidoenFactura), 0, rsEntradaSalida!bitIncluidoenFactura)
    
    Set rsSustituto = frsRegresaRs("Select PVPAGOREFACTURACION.CHRFOLIOPAGOCANCELADA from PVPAGOREFACTURACION where PVPAGOREFACTURACION.CHRFOLIOPAGOACTIVADA = '" & Trim(rsEntradaSalida!CHRFOLIORECIBO) & "'")
    lstFacturaASustituirDFP.Clear
    If rsSustituto.RecordCount > 0 Then
        Do While Not rsSustituto.EOF
            lstFacturaASustituirDFP.AddItem Trim(rsSustituto!chrfoliopagocancelada)
            rsSustituto.MoveNext
        Loop
        chkFacturaSustitutaDFP.Value = 1
    Else
        chkFacturaSustitutaDFP.Value = 0
    End If
   
    If rsEntradaSalida!bitcancelado Then
        vlForeColor = vbRed
        vlBackColor = vbWhite
        lblPagoCancelado.Caption = "Recibo cancelado"
        lblPagoCancelado.Visible = True
    Else
        '|  El recibo está activo
        vlForeColor = vbBlack  '| Negro
        vlBackColor = vbWhite  '| Blanco
        lblPagoCancelado.Visible = False
        If optEntradaDinero.Value Then
            If fblnPendienteTimbre(rsEntradaSalida!intNumPago) Then
                vlForeColor = vbBlack  '| Negro
                vlBackColor = &H80FFFF '| Amarillo
                lblPagoCancelado.Caption = "Pendiente timbre"
                lblPagoCancelado.Visible = True
            Else
                Select Case rsEntradaSalida!PendienteCancelarSAT_NE
                    Case "PC", "XX"
                        vlForeColor = &HFF&    '| Rojo
                        vlBackColor = &HC0E0FF '| Naranja
                        lblPagoCancelado.Caption = "Pendiente de cancelar"
                        lblPagoCancelado.Visible = True
                    Case "PA"
                        vlForeColor = &HFFFFFF '| Blanco
                        vlBackColor = &H80FF&  '| Naranja fuerte
                        lblPagoCancelado.Caption = "Pendiente de autorización de cancelación"
                        lblPagoCancelado.Visible = True
                    Case "CR"
                        vlForeColor = &HFFFFFF '| Blanco
                        vlBackColor = &HFF&    '| Rojo
                        lblPagoCancelado.Caption = "Cancelación rechazada"
                        lblPagoCancelado.Visible = True
                End Select
            End If
        End If
    End If
    lblPagoCancelado.ForeColor = vlForeColor
    lblPagoCancelado.BackColor = vlBackColor
            
            
            
            
''''''    lblPagoCancelado.Visible = IIf(rsEntradaSalida!bitcancelado = 1, True, False)
    
    pLimpiaConceptos
    
    grdConceptos.Row = 1
    vlstrSentencia = "SELECT PvConceptoPago.* FROM PvConceptoPago WHERE intNumConcepto = " & str(rsEntradaSalida!intNumConcepto)
    Set rs = frsRegresaRs(vlstrSentencia)
    If rs.RecordCount <> 0 Then
        grdConceptos.TextMatrix(1, lintColNombreConcepto) = rs!chrDescripcion
        grdConceptos.TextMatrix(1, lintColTipoConcepto) = rs!chrTipo
    End If
    
    cmdConfirmartimbre.Enabled = False
    cmdCFD.Enabled = False
    vlblnActivaMotivo = False
    
    If strTipoMovimiento = "E" Then
        If Not IsNull(rsEntradaSalida!INTIDCOMPROBANTE) Then
            If fblnPendienteTimbre(rsEntradaSalida!intNumPago) Then
                lblPagoCancelado.Caption = "Pendiente timbre"
                lblPagoCancelado.ForeColor = vbBlack
                lblPagoCancelado.BackColor = &H80FFFF
                cmdConfirmartimbre.Enabled = True
                lblPagoCancelado.Visible = True
            Else
                cmdCFD.Enabled = True
                vlblnActivaMotivo = True
            End If
           
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraPago"))
End Sub

Private Function fblnPendienteTimbre(lngNumPago As Long)
    Dim rs As ADODB.Recordset
    Dim strSql As String
    strSql = "select intNumPago from PVPago" & _
    " inner join GNPendientestimbreFiscal on GNPendientestimbreFiscal.INTCOMPROBANTE = PVPago.intNumPago and GNPendientestimbreFiscal.CHRTIPOCOMPROBANTE = 'AN'" & _
    " where PVPago.intNumPago = " & lngNumPago & " and PVPago.INTIDCOMPROBANTE is not null"
    Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        fblnPendienteTimbre = True
    Else
        fblnPendienteTimbre = False
    End If
    rs.Close
End Function

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    Dim vllngMensaje As Long
    
    pSocios
    
    If vldblTipoCambioVenta = 0 Then
        'Registre el tipo de cambio del día.
        MsgBox SIHOMsg(335), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
    
    vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If vllngMensaje <> 0 Then
        'Cierre el corte actual antes de registrar este documento.
        'No existe corte abierto.
        MsgBox SIHOMsg(str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje"
        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
        Unload Me
    End If

    vgblnCuentaFacturada = False '(CR) - Agregado para caso no. 6863

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        optEntradaDinero.Enabled = True
        optSalidaDinero.Enabled = True
        If SSTabPagos.Tab = 1 Then
            OptTipoPaciente(0).Value = True
            pLimpia
            SSTabPagos.Tab = 0
            txtMovimientoPaciente.SetFocus
        Else
            If lblnManipulacion Then 'Cuando la pantalla se mando llamar desde la facturación
                Unload Me
            Else
                If cmdSave.Enabled Or cmdPrint.Enabled Then
                    '¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        OptTipoPaciente(0).Value = True
                        pLimpia
                        txtMovimientoPaciente.SetFocus
                    End If
                    KeyAscii = 0
                Else
                    Unload Me
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rsPar As New ADODB.Recordset
    Dim rsConcepto As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vgstrNombreForm = Me.Name
    Me.Icon = frmMenuPrincipal.Icon
        
    lintColNombreConcepto = 1
    lintColIdConcepto = 2
    lintColTipoConcepto = 3
    lintColNumeroCuenta = 4
    
    ldtmFecha = fdtmServerFecha
    
    'Se realiza el instanciamiento del reporte
    pInicializaReporte
    
    If Not lblnManipulacion Then
        OptTipoPaciente(0).Value = True
        If fintEsInterno(vglngNumeroLogin, enmTipoProceso.pagos) > 0 Then
            If fintEsInterno(vglngNumeroLogin, enmTipoProceso.pagos) = 1 Then
                OptTipoPaciente(0).Value = True
            Else
                OptTipoPaciente(1).Value = True
            End If
        End If
    End If
    
    vldblTipoCambioVenta = fdblTipoCambio(fdtmServerFecha, "V")
    
    'Recibos configurados:
    pFormatosRecibo
    
    vlblnLimpiar = True
    
    SSTabPagos.Tab = 0
    
    'Parametro para
    llngCveConceptoDeudor = -1
    lstrDescConceptoDeudor = ""
    Set rsPar = frsSelParametros("PV", vgintClaveEmpresaContable, "INTCONCEPTOENTRADACOMPROBACION")
    If rsPar.RecordCount <> 0 Then
        If Not IsNull(rsPar!valor) Then
            Set rsConcepto = frsRegresaRs("select * from PvConceptoPago where intNumConcepto = " & rsPar!valor)
            If rsConcepto.RecordCount <> 0 Then
                llngCveConceptoDeudor = rsConcepto!intNumConcepto
                lstrDescConceptoDeudor = Trim(rsConcepto!chrDescripcion)
            End If
        End If
    End If
    
    '--------------------------------------------------------------------------
    ' Parámetro que indica si se permite cancelar recibos de otro departamento
    '--------------------------------------------------------------------------
    vlblnCancelarRecibosOtroDepto = False
    vlstrSentencia = "select bitCancelarRecibosOtroDepto from pvparametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsTemp = frsRegresaRs(vlstrSentencia)
    vlblnCancelarRecibosOtroDepto = IIf(IsNull(rsTemp!bitCancelarRecibosOtroDepto), False, rsTemp!bitCancelarRecibosOtroDepto)
    rsTemp.Close
    
    'Regimen fiscal
    pCargaRegimenFiscal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub grdPagos_DblClick()
On Error GoTo NotificaError
Dim rs As ADODB.Recordset

    
    If Val(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)) <> 0 Then
        pMuestraPago CLng(grdPagos.TextMatrix(grdPagos.Row, cintColIdEntradaSalida)), grdPagos.TextMatrix(grdPagos.Row, cintColTipoMovimiento)
        pHabilita 1, 1, 1, 1, 1, 0, IIf(fblnRevisaPermiso(vglngNumeroLogin, 302, "C"), 1, 0), IIf((IsNull(rsEntradaSalida!chrfoliofactura) Or Trim(rsEntradaSalida!chrfoliofactura) = "") And rsEntradaSalida!bitcancelado = 0, 1, 0)
            
        SSTabPagos.Tab = 0
        cmdLocate.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPagos_DblClick"))
End Sub



Private Sub lstFacturaASustituirDFP_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If txtComentario.Enabled Then
            txtComentario.SetFocus
        End If
    Else
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstFacturaASustituirDFP_KeyPress"))
End Sub

Private Sub mskFecha_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado Then
        txtMovimientoPaciente.SetFocus
    Else
        pSelMkTexto mskFecha
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecha_GotFocus"))
End Sub

Private Sub MskFecha_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtPersona.SetFocus
    Else
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFecha_KeyPress"))
End Sub

Private Sub mskFechaFinal_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFechaFinal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFinal_GotFocus"))
End Sub

Private Sub mskFechaFinal_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(mskFechaFinal.ClipText) = "" Then
            mskFechaFinal.Text = fdtmServerFecha
        End If
        cmdCargar.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFinal_KeyPress"))
End Sub

Private Sub mskFechaInicial_GotFocus()
On Error GoTo NotificaError
    
    pSelMkTexto mskFechaInicial

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicial_GotFocus"))
End Sub

Private Sub mskFechaInicial_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFechaFinal.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicial_KeyPress"))
End Sub

Private Sub mskFechaInicial_LostFocus()
On Error GoTo NotificaError
    
    If Trim(mskFechaInicial.ClipText) = "" Then
        mskFechaInicial.Text = fdtmServerFecha
    End If
    
    If Not IsDate(mskFechaInicial.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        pEnfocaMkTexto mskFechaInicial
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicial_LostFocus"))
End Sub

Private Sub optDeudor_Click()

    If llngCveConceptoDeudor = -1 And Not vlblnConsulta Then
        'No se ha configurado el parámetro de concepto de entrada de dinero para comprobación de gastos.
        MsgBox SIHOMsg(1483), vbOKOnly + vbInformation, "Mensaje"
        OptTipoPaciente(0).Value = True
    Else
        If Not lblnConsultaDeudor Then
            vlblnConsulta = False
            fraRecibo.Enabled = True
        End If
        vlblnPacienteSeleccionado = False
        txtMovimientoPaciente.Text = ""
        txtPaciente.Text = ""
        vlrfcPaciente = ""
        txtEmpresaPaciente.Text = ""
        txtTipoPaciente.Text = ""
        txtFechaFinal.Text = ""
        txtFechaInicial.Text = ""
        mskFecha.Mask = ""
        mskFecha.Text = ldtmFecha
        mskFecha.Mask = "##/##/####"
        txtPersona.Text = ""
        txtCantidad.Text = ""
        txtCantidadenLetras.Text = ""
        optPesos.Value = False
        optDolares.Value = False
        txtDocumento.Text = ""
        txtComentario.Text = ""
        lblPagoCancelado.Visible = False
        fraImprimir.Visible = False
    
        optSocio.Value = False
        OptTipoPaciente(0).Value = False
        OptTipoPaciente(1).Value = False
        optEntradaDinero.Value = True
        optSalidaDinero.Enabled = False
        lbCuenta.Caption = "Número de cuenta"
        pHabilita 0, 0, 0, 0, 0, 1, 0, 0
        If fblnCanFocus(txtPersona) Then txtPersona.SetFocus
    End If
End Sub

Private Sub optDolares_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
       If Val(Format(txtCantidad.Text, "################.00")) = 0 Then
            txtCantidadenLetras.Text = ""
        Else
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "dolares", "")
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optDolares_GotFocus"))
End Sub
Public Sub optDolares_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then

        txtComentario.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optDolares_KeyPress"))
End Sub

Private Sub optEntradaDinero_Click()
On Error GoTo NotificaError

    If Trim(txtPaciente.Text) <> "" Or optDeudor.Value Then
        pCargaConceptos "NO", lintDeducible, lintCoaseguro, lintCopago, lintCoaseguroAdicional

        If Val(grdConceptos.TextMatrix(1, lintColIdConcepto)) = 0 Then
            'No existen conceptos para pago.
            MsgBox SIHOMsg(298), vbOKOnly + vbInformation, "Mensaje"
            pEnfocaTextBox txtMovimientoPaciente
        Else
            If Not lblnManipulacion Then
                pCargaFolio "RE"
            End If
            If Not vlblnConsulta Then
                If Trim(txtFolioRecibo.Text) = "" Then
                    'No existen folios activos para este documento.
                    MsgBox SIHOMsg(291), vbCritical, "Mensaje"
                    'pEnfocaTextBox txtMovimientoPaciente
                Else
                    fraRecibo.Enabled = True
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optEntradaDinero_Click"))
End Sub

Public Sub pCargaFolio(strTipoDocumento As String)
On Error GoTo NotificaError

    Dim vllngFoliosRestantes As Long
    Dim vlstrFolioDocumento As String
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String

    vllngFoliosRestantes = 1
    vlstrFolioDocumento = ""
    
    If Not vlblnConsulta Then
        If Trim(strTipoDocumento) <> "" Then
            pCargaArreglo alstrParametrosSalida, vllngFoliosRestantes & "|" & adInteger & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
            frsEjecuta_SP "'" & Trim(strTipoDocumento) & "'|" & vgintNumeroDepartamento & "|0", "sp_gnFolios", , , alstrParametrosSalida
            pObtieneValores alstrParametrosSalida, vllngFoliosRestantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
            '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
            strSerie = Trim(strSerie)
            vlstrFolioDocumento = strSerie & strFolio
            If Trim(vlstrFolioDocumento) <> "0" Then
                If vllngFoliosRestantes > 0 Then
                    MsgBox "Faltan " & Trim(str(vllngFoliosRestantes)) + " recibos y será necesario aumentar folios!", vbOKOnly + vbInformation, "Mensaje"
                End If
                txtFolioRecibo.Text = vlstrFolioDocumento
            Else
                txtFolioRecibo.Text = ""
            End If
        Else
            txtFolioRecibo.Text = ""
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaFolio"))
End Sub

Private Sub pCargaConceptos(strTipo As String, intDeducible As Integer, intCoaseguro As Integer, intCopago As Integer, intCoaseguroAdicional As Integer)
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset
        
    pLimpiaConceptos
    If optDeudor.Value Then
        grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColNombreConcepto) = lstrDescConceptoDeudor
        grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColIdConcepto) = llngCveConceptoDeudor
        grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColTipoConcepto) = "NO"
    Else
        vgstrParametrosSP = "-1" & "|" & str(vgintClaveEmpresaContable) & "|" & strTipo & "|" & "1" & "|" & str(intDeducible) & "|" & str(intCoaseguro) & "|" & str(intCopago) & "|" & str(intCoaseguroAdicional)
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCONCEPTOPAGO")
        If rs.RecordCount <> 0 Then
            Do While Not rs.EOF
                grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColNombreConcepto) = rs!chrDescripcion
                grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColIdConcepto) = rs!intNumConcepto
                grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColTipoConcepto) = rs!chrTipo
                grdConceptos.TextMatrix(grdConceptos.Rows - 1, lintColNumeroCuenta) = rs!intNumeroCuenta
                grdConceptos.Rows = grdConceptos.Rows + 1
                rs.MoveNext
            Loop
            grdConceptos.Rows = grdConceptos.Rows - 1
        End If
    End If
    If vlnblnLocate = False Then
        grdConceptos_Click
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaConceptos"))
End Sub

Private Sub optEntradaDinero_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        optEntradaDinero_Click
        
        If fraRecibo.Enabled Then
            txtPersona.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optEntradaDinero_KeyDown"))
End Sub

Private Sub optMostrarSolo_Click(Index As Integer)
    cmdCargar_Click
End Sub

Private Sub optPesos_GotFocus()
On Error GoTo NotificaError
    
    
    If Not vlblnPacienteSeleccionado And optSocio.Value = False And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
         If Val(Format(txtCantidad.Text, "################.00")) = 0 Then
            txtCantidadenLetras.Text = ""
        Else
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optPesos_GotFocus"))
End Sub

Public Sub optPesos_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If KeyAscii = 13 Then
        If chkFacturaSustitutaDFP.Enabled Then
            chkFacturaSustitutaDFP.SetFocus
        Else
            txtComentario.SetFocus
        End If
    Else
     KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optPesos_KeyPress"))
End Sub
Public Sub optSalidaDinero_Click()
On Error GoTo NotificaError

    If Trim(txtPaciente.Text) <> "" Then
        pCargaConceptos "SD", 0, 0, 0, 0

        If Val(grdConceptos.TextMatrix(1, lintColIdConcepto)) = 0 Then
            'No existen conceptos para pago.
            MsgBox SIHOMsg(298), vbOKOnly + vbInformation, "Mensaje"
            pEnfocaTextBox txtMovimientoPaciente
        Else
            If Not lblnManipulacion Then
                pCargaFolio "SD"
            End If
            If Not vlblnConsulta Then
                If Trim(txtFolioRecibo.Text) = "" Then
                    'No existen folios activos para este documento.
                    MsgBox SIHOMsg(291), vbCritical, "Mensaje"
                    'pEnfocaTextBox txtMovimientoPaciente
                Else
                    fraRecibo.Enabled = True
                End If
            End If
        End If
    End If
    
    chkFacturaSustitutaDFP.Value = 0
    chkFacturaSustitutaDFP.Enabled = False
    
    lstFacturaASustituirDFP.Enabled = False
    lstFacturaASustituirDFP.Clear
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optSalidaDinero_Click"))
End Sub

Private Sub optSalidaDinero_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        optSalidaDinero_Click
    
        If fraRecibo.Enabled Then
            txtPersona.SetFocus
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optSalidaDinero_KeyDown"))
End Sub

Private Sub optSocio_Click()
On Error GoTo NotificaError
    
    If Not vlblnConsulta Then
        pEnfocaTextBox txtMovimientoPaciente
    End If
    
    OptTipoPaciente(0).Value = False
    OptTipoPaciente(1).Value = False
    
    lbCuenta.Caption = "Clave única"
    
    txtMovimientoPaciente.MaxLength = 20
    optEntradaDinero.Value = False
    optSalidaDinero.Value = False
    
    optSalidaDinero.Enabled = False
    optDolares.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optSocio_Click"))
End Sub

Private Sub optTipo_Click(Index As Integer)
On Error GoTo NotificaError

    If optTipo(2).Value Or optTipo(3).Value Then
        optTipoBusqueda(0).Value = True
        optTipoBusqueda(1).Enabled = False
    Else
        optTipoBusqueda(1).Enabled = True
    End If
    
    If optTipo(3).Value Then
        optTipoMovimiento(1).Value = True
        optTipoMovimiento(0).Enabled = False
        optTipoMovimiento(2).Enabled = False
    Else
        optTipoMovimiento(0).Value = True
        optTipoMovimiento(0).Enabled = True
        optTipoMovimiento(2).Enabled = True
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipo_Click"))
End Sub

Private Sub optTipo_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If optTipoBusqueda(0).Value Then
            optTipoBusqueda(0).SetFocus
        Else
            optTipoBusqueda(1).SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipo_KeyPress"))
End Sub

Private Sub optTipoBusqueda_Click(Index As Integer)
On Error GoTo NotificaError

    fraRangoFechas.Enabled = optTipoBusqueda(0).Value
    
    mskFechaInicial.Mask = ""
    mskFechaInicial.Text = ldtmFecha
    mskFechaInicial.Mask = "##/##/####"
    
    mskFechaFinal.Mask = ""
    mskFechaFinal.Text = ldtmFecha
    mskFechaFinal.Mask = "##/##/####"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoBusqueda_Click"))
End Sub

Private Sub optTipoBusqueda_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
        If optTipoBusqueda(0).Value Then
            If optTipoMovimiento(0).Value Then
                optTipoMovimiento(0).SetFocus
            ElseIf optTipoMovimiento(1).Value Then
                optTipoMovimiento(1).SetFocus
            Else
                optTipoMovimiento(2).SetFocus
            End If
        Else
            With FrmBusquedaPacientes
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Value = False
                .optSinFacturar.Enabled = True
                .optSoloActivos.Enabled = True
                .optTodos.Value = True
                .optTodos.Enabled = True
                
                If optTipo(1).Value Then 'Externos
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1750,4100"
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " Externos"
                Else
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,2200,1050,1050,4100"
                    .vgstrTipoPaciente = "I"
                    .Caption = .Caption & " Internos"
                End If
        
                llngNumCuenta = .flngRegresaPaciente()
            End With
            
            cmdCargar.SetFocus
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoBusqueda_KeyPress"))
End Sub

Private Sub optTipoMovimiento_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If fraRangoFechas.Enabled Then
            mskFechaInicial.SetFocus
        Else
            cmdCargar.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoMovimiento_KeyDown"))
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
On Error GoTo NotificaError
    
    If Not vlblnConsulta Then
        pEnfocaTextBox txtMovimientoPaciente
    End If
    
    optSocio.Value = False
    
    lbCuenta.Caption = "Número de cuenta"
    
    txtMovimientoPaciente.MaxLength = 10
    optSalidaDinero.Enabled = True
    optDolares.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":OptTipoPaciente_Click"))
End Sub

Private Sub txtCantidad_GotFocus()
On Error GoTo NotificaError
    
    pSelTextBox txtCantidad
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidad_GotFocus"))
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not fblnFormatoCantidad(txtCantidad, KeyAscii, 2) Then
       KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            optPesos.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidad_KeyPress"))
End Sub
Private Sub txtCantidad_LostFocus()
If Me.optPesos.Value Then
   If Val(Format(txtCantidad.Text, "################.00")) = 0 Then
            txtCantidadenLetras.Text = ""
        Else
            txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "pesos", "M.N.")
        End If
ElseIf Me.optDolares.Value Then

   If Val(Format(txtCantidad.Text, "################.00")) = 0 Then
       txtCantidadenLetras.Text = ""
   Else
       txtCantidadenLetras.Text = fstrNumeroenLetras(Val(Format(txtCantidad.Text, "################.00")), "dolares", "")
   End If
End If
End Sub

Private Sub txtCantidadenLetras_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optSocio.Value = False And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
        pSelTextBox txtCantidad
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtCantidadenLetras_GotFocus"))
End Sub

Private Sub txtComentario_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
        pSelTextBox txtComentario
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtComentario_GotFocus"))
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
     End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtComentario_KeyPress"))
End Sub

Private Sub txtDocumento_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDocumento_GotFocus"))
End Sub

Private Sub txtFolioRecibo_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
        pSelTextBox txtFolioRecibo
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioRecibo_GotFocus"))
End Sub

Private Sub txtFolioRecibo_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        mskFecha.SetFocus
    Else
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioRecibo_KeyPress"))
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
On Error GoTo NotificaError
    
    lblnConsultaDeudor = False
    If vlblnLimpiar Then
        If optDeudor.Value Then OptTipoPaciente(0).Value = True
        pLimpia
        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    Else
        vlblnLimpiar = True
    End If

    If vglngMovPaciente <> 0 Then
        txtMovimientoPaciente.Text = vglngMovPaciente
        OptTipoPaciente(0).Value = vgstrTipoPaciente = "I"
        OptTipoPaciente(1).Value = vgstrTipoPaciente = "E"
        OptTipoPaciente(0).Enabled = False
        OptTipoPaciente(1).Enabled = False
        optEntradaDinero.Value = vgstrTipoEntradaSalida = "E" Or vgstrTipoEntradaSalida = ""
        optSalidaDinero.Value = vgstrTipoEntradaSalida = "S"
        optEntradaDinero.Enabled = False
        optSalidaDinero.Enabled = False
        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_GotFocus"))
End Sub

Public Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    Dim rsSocio As New ADODB.Recordset
    Dim strParametrosSP As String
    Dim strTest As String
    strTest = ""
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
            If optSocio.Value = False Then
                With FrmBusquedaPacientes
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    
                    
                    If OptTipoPaciente(1).Value Then 'Externos
                        .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo, ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,1500,1750,4100"
                        .vgstrTipoPaciente = "E"
                        .Caption = .Caption & " Externos"
                    Else
                        .vgStrOtrosCampos = ", SiTipoIngreso.vchNombre as Tipo, ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,2200,1050,1050,4100"
                        .vgstrTipoPaciente = "I"
                        .Caption = .Caption & " Internos"
                    End If
                    .grdBusqueda.ColAlignment(1) = flexAlignCenterBottom
                    
                    txtMovimientoPaciente.Text = .flngRegresaPaciente()
                    
                    If txtMovimientoPaciente <> -1 Then
                        vlblnLimpiar = False
                        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                    Else
                        txtMovimientoPaciente.Text = ""
                    End If
                    
                End With
            Else
                ' Busca socio
                With frmSociosBusqueda
                    .Show vbModal
                    txtMovimientoPaciente.Text = .vgstrClaveUnica
                    
                    If .vgstrClaveUnica = "" Then Exit Sub
                    
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                    Unload frmSociosBusqueda
                End With
            End If
        Else
            If optSocio.Value = False Then
                If fblnDatosPaciente(Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")) Then
                    pRevisaControlAseguradora Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")
                    
                    txtPersona.Text = txtPaciente.Text
                    vlblnPacienteSeleccionado = True
                    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                    
                    If Not lblnManipulacion Then
                        optEntradaDinero.SetFocus
                        optEntradaDinero.Value = True
                    End If
                End If
            Else
                strParametrosSP = txtMovimientoPaciente & "|" & -1
                Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_PVSELSOCIOS")
                If rsSocio.RecordCount > 0 Then
                    txtPaciente.Text = Trim(rsSocio!vchApellidoPaterno) & " " & Trim(rsSocio!vchApellidoMaterno) & " " & Trim(rsSocio!vchNombre)
                    vlrfcPaciente = Trim(Replace(Replace(Replace(rsSocio!vchRFC, "-", ""), "_", ""), " ", ""))
                    txtTipoPaciente.Text = "SOCIO"
                    txtFechaInicial.Text = ""
                    txtFechaFinal.Text = ""
                    
                    txtPersona.Text = txtPaciente.Text
                    vlblnPacienteSeleccionado = True
                    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                    
                    vlblnLimpiar = False
                    
                    optEntradaDinero.Value = True
                    optEntradaDinero_Click
                    optPesos.Value = True
                    optDolares.Enabled = False
                    
                    If grdConceptos.Enabled = True Then grdConceptos.SetFocus
                End If
                
                rsSocio.Close
            End If
        End If
    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyDown"))
End Sub

Private Function fblnDatosPaciente(vllngxMovimiento As Long, vlstrxTipoPaciente As String) As Boolean
On Error GoTo NotificaError
    
    Dim vlrsPvSelDatosPaciente As New ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    If vlstrxTipoPaciente = "D" Then
        fblnDatosPaciente = True
        txtPaciente.Text = ""
        vlrfcPaciente = ""
        txtEmpresaPaciente.Text = ""
        txtTipoPaciente.Text = ""
        txtFechaInicial = ""
        txtFechaFinal = ""
        lintCveEmpresa = 0
    Else
        Set vlrsPvSelDatosPaciente = frsEjecuta_SP(CStr(vllngxMovimiento) & "|0|" & vlstrxTipoPaciente & "|" & vgintClaveEmpresaContable, "Sp_PvSelDatosPaciente")
        If vlrsPvSelDatosPaciente.RecordCount <> 0 Then
            If vlrsPvSelDatosPaciente!Facturada = 0 Or vlblnConsulta Then
                fblnDatosPaciente = True
                txtPaciente.Text = vlrsPvSelDatosPaciente!Nombre
                vlrfcPaciente = Trim(Replace(Replace(Replace(vlrsPvSelDatosPaciente!RFCPaciente, "-", ""), "_", ""), " ", ""))
                txtEmpresaPaciente.Text = IIf(IsNull(vlrsPvSelDatosPaciente!empresa), "", vlrsPvSelDatosPaciente!empresa)
                txtTipoPaciente.Text = vlrsPvSelDatosPaciente!tipo
                txtFechaInicial = IIf(IsNull(vlrsPvSelDatosPaciente!Ingreso), "", vlrsPvSelDatosPaciente!Ingreso)
                txtFechaFinal = IIf(IsNull(vlrsPvSelDatosPaciente!Egreso), "", vlrsPvSelDatosPaciente!Egreso)
                lintCveEmpresa = vlrsPvSelDatosPaciente!intcveempresa
            Else
                fblnDatosPaciente = False
                'La cuenta del paciente está completamente facturada.
                MsgBox SIHOMsg(597), vbExclamation, "Mensaje"
                vgblnCuentaFacturada = True '(CR) - Agregado para caso no. 6863
                pEnfocaTextBox txtMovimientoPaciente
            End If
        Else
            fblnDatosPaciente = False
            '¡La información no existe!
            MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            pEnfocaTextBox txtMovimientoPaciente
        End If
        vlrsPvSelDatosPaciente.Close
    End If
  
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosPaciente"))
End Function

Private Sub pRevisaControlAseguradora(lngnumCuenta As Long, strTipoPaciente As String)
On Error GoTo NotificaError

    Dim dblDeducible As Double
    Dim dblCoaseguro As Double
    Dim dblCoaseguroAdicional As Double
    Dim dblCopago As Double

    'Revisar si el paciente, tiene el control de aseguradora registrado:
    lintDeducible = 0
    lintCoaseguro = 0
    lintCoaseguroAdicional = 0
    lintCopago = 0
    
    vgstrParametrosSP = str(lngnumCuenta) & "|" & strTipoPaciente & "|" & str(lintCveEmpresa) & "|0"
    Set rsControlSeguro = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCONTROLSEGUROEMPRESA")
    If rsControlSeguro.RecordCount <> 0 Then
        dblDeducible = IIf(IsNull(rsControlSeguro!MNYCANTIDADDEDUCIBLE), 0, rsControlSeguro!MNYCANTIDADDEDUCIBLE)
        dblCoaseguro = IIf(IsNull(rsControlSeguro!MNYCANTIDADCOASEGURO), 0, rsControlSeguro!MNYCANTIDADCOASEGURO)
        dblCoaseguroAdicional = IIf(IsNull(rsControlSeguro!MNYCANTIDADCOASEGUROADICIONAL), 0, rsControlSeguro!MNYCANTIDADCOASEGUROADICIONAL)
        dblCopago = IIf(IsNull(rsControlSeguro!MNYCANTIDADCOPAGO), 0, rsControlSeguro!MNYCANTIDADCOPAGO)
        
        lintDeducible = IIf(rsControlSeguro!BITFACTURADEDUCIBLE = 0 And dblDeducible <> 0 And IsNull(rsControlSeguro!CHRFOLIORECIBODEDUCIBLE), 1, 0)
        lintCoaseguro = IIf(rsControlSeguro!BITFACTURACOASEGURO = 0 And dblCoaseguro <> 0 And IsNull(rsControlSeguro!CHRFOLIORECIBOCOASEGURO), 1, 0)
        lintCoaseguroAdicional = IIf(rsControlSeguro!bitFacturaCoaseguroAdicional = 0 And dblCoaseguroAdicional <> 0 And IsNull(rsControlSeguro!CHRFOLIORECIBOCOASEGUROADICION), 1, 0)
        lintCopago = IIf(rsControlSeguro!BITFACTURACOPAGO = 0 And dblCopago <> 0 And IsNull(rsControlSeguro!CHRFOLIORECIBOCOPAGO), 1, 0)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pRevisaControlAseguradora"))
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If optSocio.Value = False Then
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Or UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "D" Then
                OptTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
                OptTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
                optSocio.Value = UCase(Chr(KeyAscii)) = "S"
                optDeudor.Value = UCase(Chr(KeyAscii)) = "D"
            End If
            KeyAscii = 7
        End If
    Else
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Or UCase(Chr(KeyAscii)) = "S" Or UCase(Chr(KeyAscii)) = "D" Then
            OptTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            OptTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
            optSocio.Value = UCase(Chr(KeyAscii)) = "S"
            optDeudor.Value = UCase(Chr(KeyAscii)) = "D"
            KeyAscii = 7
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
            
            
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyPress"))
End Sub

Private Sub optTipoPaciente_GotFocus(Index As Integer)
On Error GoTo NotificaError
    
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoPaciente_GotFocus"))
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtMovimientoPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optTipoPaciente_KeyPress"))
End Sub

Private Sub txtPersona_GotFocus()
On Error GoTo NotificaError
    
    If Not vlblnPacienteSeleccionado And optSocio.Value = False And optDeudor.Value = False Then
        txtMovimientoPaciente.SetFocus
    Else
        pSelTextBox txtPersona
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPersona_GotFocus"))
End Sub

Private Sub txtPersona_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If optDeudor.Value Then
            If fblnCanFocus(txtCantidad) Then txtCantidad.SetFocus
        Else
            grdConceptos.Col = lintColNombreConcepto
            grdConceptos.SetFocus
        End If
    Else
            'A-Z                                    a-z                 # ''                    ""                              #            ñ              Ñ              '                    ,              .
        If KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii >= 34 And KeyAscii <= 57 Or KeyAscii = 127 Or (KeyAscii >= 97 And KeyAscii <= 241) Or KeyAscii = 39 Or KeyAscii = 44 Or KeyAscii = 46 Then  ' según pruebas se deben aceptar puntos, comas y apostrofos, ni pex
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
          
        Else
           KeyAscii = 0
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPersona_KeyPress"))
End Sub

Private Sub pSocios()
On Error GoTo NotificaError
    'Este procedimiento permite verificar si los socios estan activados
    Dim rs As ADODB.Recordset
    
    Set rs = frsRegresaRs("SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITUTILIZASOCIOS'")
    If Not rs.EOF Then
        'vgintSocios = rs!vchValor
        If rs!vchvalor = 1 Then
            optSocio.Enabled = True
            optSocio.Visible = True
            chkSocios.Value = 0
            chkSocios.Enabled = True
            chkSocios.Visible = True
            optDeudor.Left = 2630
        Else
            optSocio.Enabled = False
            optSocio.Visible = False
            chkSocios.Value = 0
            chkSocios.Enabled = False
            chkSocios.Visible = False
            optDeudor.Left = 1830
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSocios"))
End Sub

'- CASO 6894: Registra el movimiento de cancelación en el libro de bancos, si es que se pagó con transferencia -'
'// MODIFICADO PARA CASO 7442: Guardar cualquier forma de pago ligada a una cuenta de banco en tabla intermedia \\'
Private Sub pCancelaMovimiento(vlintNumPago As Long, vlstrFolio As String, vlStrReferencia As String, vlintCorteMovimiento As Long, vllngCorteActual As Long)
On Error GoTo NotificaError

    Dim rsMovimiento As ADODB.Recordset
    Dim lstrTipoDoc As String, lstrFecha As String
    Dim ldblCantidad As Double
    Dim rs As ADODB.Recordset
       
    If vlStrReferencia = "PA" Then  'Pago automático
            vlstrSentencia = "select distinct  pvcortepoliza.intnumcorte, pvcortepoliza.chrtipodocumento, pvfactura.intconsecutivo " & _
                     "from pvpagocortepoliza, pvcortepoliza, pvfactura " & _
                     "where trim(pvpagocortepoliza.chrfoliorecibo) = trim('" & vlstrFolio & "') " & _
                     "and pvpagocortepoliza.intconsecutivo = pvcortepoliza.intconsecutivo " & _
                     "and trim(pvfactura.chrfoliofactura) = trim(pvcortepoliza.chrfoliodocumento)"
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount > 0 Then
            vlintCorteMovimiento = rs!intNumCorte
            vlStrReferencia = rs!chrTipoDocumento
            vlintNumPago = rs!intConsecutivo
        End If
    End If
    
    vlstrSentencia = "SELECT MB.intFormaPago, MB.mnyCantidad, MB.mnyTipoCambio, FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco) AS IdBanco, mb.chrtipomovimiento " & _
                     " FROM PvMovimientoBancoForma MB " & _
                     " INNER JOIN PvFormaPago FP ON MB.intFormaPago = FP.intFormaPago " & _
                     " LEFT  JOIN CpBanco B ON B.intNumeroCuenta = FP.intCuentaContable " & _
                     " WHERE TRIM(MB.chrTipoDocumento) = '" & Trim(vlStrReferencia) & "' AND MB.intNumDocumento = " & vlintNumPago & _
                     " AND MB.intNumCorte = " & vlintCorteMovimiento
    Set rsMovimiento = frsRegresaRs(vlstrSentencia)
    If rsMovimiento.RecordCount > 0 Then
        lstrFecha = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) '- Fecha y hora del movimiento -'
    
        rsMovimiento.MoveFirst
        Do While Not rsMovimiento.EOF
            If rsMovimiento!chrTipo <> "C" Then
                '- Revisar tipo de forma de pago para determinar movimiento de cancelación -'
                If optEntradaDinero.Value Then ' Para Entradas de dinero
                    If rsMovimiento!chrTipoMovimiento = "CBA" Then
                        lstrTipoDoc = "CCB"               'Comisión bancaria
                    Else
                        Select Case rsMovimiento!chrTipo
                            Case "E": lstrTipoDoc = "CEP" 'Efectivo
                            Case "T": lstrTipoDoc = "CJP" 'Tarjeta de crédito
                            Case "B": lstrTipoDoc = "CTP" 'Transferencia bancaria
                            Case "H": lstrTipoDoc = "CQP" 'Cheque
                        End Select
                    End If
                Else ' Para Salidas de dinero
                    Select Case rsMovimiento!chrTipo
                        Case "E": lstrTipoDoc = "CSE" 'Efectivo
                        Case "T": lstrTipoDoc = "CST" 'Tarjeta de crédito
                        Case "B": lstrTipoDoc = "CSB" 'Transferencia bancaria
                        Case "H": lstrTipoDoc = "CSC" 'Cheque
                    End Select
                End If
    
                '- Cantidad negativa para que se tome como abono si se cancela una entrada de dinero, cantidad positiva si se cancela salida de dinero -'
                ldblCantidad = rsMovimiento!MNYCantidad * IIf(optEntradaDinero.Value, -1, 1)
    
                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rsMovimiento!intFormaPago & "|" & rsMovimiento!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rsMovimiento!mnytipocambio = 0, 1, 0) & "|" & rsMovimiento!mnytipocambio & "|" & lstrTipoDoc & "|" & vlStrReferencia & "|" & _
                                    vlintNumPago & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & lstrFecha & "|" & "1" & "|" & cgstrModulo
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
            End If
            rsMovimiento.MoveNext
        Loop
    End If
    rsMovimiento.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaMovimiento"))
End Sub

'- CASO 7442: Regresa tipo de movimiento según la forma de pago -'
Private Function fstrTipoMovimientoForma(lintCveForma As Integer) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    
    fstrTipoMovimientoForma = ""
    
    vlstrSentencia = "SELECT * FROM PvFormaPago WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        If optEntradaDinero.Value Then
        '- Entradas de Dinero -'
            Select Case rsForma!chrTipo
                Case "E": fstrTipoMovimientoForma = "EFP" 'Efectivo
                Case "T": fstrTipoMovimientoForma = "TAP" 'Tarjeta
                Case "B": fstrTipoMovimientoForma = "TPA" 'Transferencia
                Case "H": fstrTipoMovimientoForma = "CHP" 'Cheque
            End Select
        Else
        '- Salidas de dinero -'
            Select Case rsForma!chrTipo
                Case "E": fstrTipoMovimientoForma = "SEP" 'Efectivo
                Case "T": fstrTipoMovimientoForma = "STP" 'Tarjeta de crédito
                Case "B": fstrTipoMovimientoForma = "SBP" 'Transferencia
                Case "H": fstrTipoMovimientoForma = "SCP" 'Cheque
            End Select
        End If
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrTipoMovimientoForma"))
End Function

Private Function fblnGenerarComprobante() As Boolean
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select bitGenerarCFDI from PVConceptoPago where intNumConcepto = " & grdConceptos.TextMatrix(grdConceptos.Row, lintColIdConcepto))
    If Not rs.EOF Then
        fblnGenerarComprobante = rs!bitGenerarCFDI
    Else
        fblnGenerarComprobante = False
    End If
    rs.Close
End Function


Private Sub pDatosFiscales(strRFC As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    strSql = "select * from PVDatosFiscales where chrRFC = '" & strRFC & "'"
    Set rsTmp = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    Load frmDatosFiscales
    frmDatosFiscales.vgblnMostrarUsoCFDI = True
    If Not rsTmp.EOF Then
        frmDatosFiscales.vgstrNombre = IIf(IsNull(rsTmp!CHRNOMBRE), "", Trim(rsTmp!CHRNOMBRE))
        frmDatosFiscales.vgstrDireccion = IIf(IsNull(rsTmp!chrCalle), "", Trim(rsTmp!chrCalle))
        frmDatosFiscales.vgstrNumExterior = IIf(IsNull(rsTmp!VCHNUMEROEXTERIOR), "", rsTmp!VCHNUMEROEXTERIOR)
        frmDatosFiscales.vgstrNumInterior = IIf(IsNull(rsTmp!VCHNUMEROINTERIOR), "", rsTmp!VCHNUMEROINTERIOR)
        frmDatosFiscales.vgstrColonia = IIf(IsNull(rsTmp!VCHCOLONIA), "", rsTmp!VCHCOLONIA)
        frmDatosFiscales.vgstrCP = IIf(IsNull(rsTmp!VCHCODIGOPOSTAL), "", rsTmp!VCHCODIGOPOSTAL)
        If Not IsNull(rsTmp!INTCVECIUDAD) Then
            frmDatosFiscales.cboCiudad.ListIndex = flngLocalizaCbo(frmDatosFiscales.cboCiudad, str(rsTmp!INTCVECIUDAD))
        End If
        frmDatosFiscales.vgstrTelefono = IIf(IsNull(rsTmp!chrTelefono), "", Trim(rsTmp!chrTelefono))
        frmDatosFiscales.vgstrRFC = IIf(IsNull(rsTmp!CHRRFC), "", Trim(rsTmp!CHRRFC))
        frmDatosFiscales.vlstrNumRef = IIf(IsNull(rsTmp!intNumReferencia), "", rsTmp!intNumReferencia)
        frmDatosFiscales.vlstrTipo = IIf(IsNull(rsTmp!CHRTIPOPACIENTE), "", Trim(rsTmp!CHRTIPOPACIENTE))
        
        If Not IsNull(rsTmp!VCHREGIMENFISCAL) Then
            frmDatosFiscales.cboRegimenFiscal.ListIndex = flngLocalizaCbo(frmDatosFiscales.cboRegimenFiscal, rsTmp!VCHREGIMENFISCAL)
        End If
    Else
        frmDatosFiscales.vgstrRFC = strRFC
        frmDatosFiscales.vgstrNombre = txtPaciente.Text
        frmDatosFiscales.cboRegimenFiscal.ListIndex = 0
    End If
    rsTmp.Close
    frmDatosFiscales.sstDatos.Tab = 0
    frmDatosFiscales.Show vbModal
End Sub

Private Function pGeneraCFDI(blnGeneraCFD As Boolean, lngNumPagoSalida As Long, strAnoAprobacion As String, strNumeroAprobacion As String, strSerie As String, strFolio As String) As Integer

    Dim lstrsql As String
    Dim rsCnEmpresaContable As ADODB.Recordset
    Dim strNombreEmisor As String
    Dim strRFCEmisor As String
    Dim strCalleEmisor As String
    Dim strTelefonoEmisor As String
    Dim strColoniaEmisor As String
    Dim strCPEmisor As String
    Dim strCiudadEmisor As String
    Dim strEstadoEmisor As String
    Dim strPaisEmisor As String
    Dim strCiudadReceptor As String
    Dim strEstadoReceptor As String
    Dim strPaisReceptor As String
    Dim strParametros As String
    Dim rsSec_gncomproFD As ADODB.Recordset
    Dim lnSec_gncomproFD As Long
    Dim strSentencia As String
    Dim rsPago As ADODB.Recordset
    Dim rsRutas As ADODB.Recordset
    Dim strRuta As String
    Dim rs As New ADODB.Recordset
    pGeneraCFDI = 0
    If (vgstrVersionCFDI = "3.3" And fblnLicenciaCFDI33(False)) Or (vgstrVersionCFDI = "4.0" And fblnLicenciaCFDI33(False)) Then
            If blnGeneraCFD Then
                lstrsql = "SELECT * FROM CnEmpresaContable WHERE tnyClaveEmpresa = " & CStr(vgintClaveEmpresaContable)
                Set rsCnEmpresaContable = frsRegresaRs(lstrsql, adLockOptimistic, adOpenDynamic)
                If rsCnEmpresaContable.RecordCount > 0 Then
                    strNombreEmisor = IIf(IsNull(rsCnEmpresaContable!vchNombre), "", rsCnEmpresaContable!vchNombre)
                    strRFCEmisor = IIf(IsNull(rsCnEmpresaContable!vchRFC), "", rsCnEmpresaContable!vchRFC)
                    strCalleEmisor = IIf(IsNull(rsCnEmpresaContable!vchCalle), "", rsCnEmpresaContable!vchCalle) & " " & IIf(IsNull(rsCnEmpresaContable!VCHNUMEROEXTERIOR), "", rsCnEmpresaContable!VCHNUMEROEXTERIOR)
                    strTelefonoEmisor = IIf(IsNull(rsCnEmpresaContable!vchTelefono), "", rsCnEmpresaContable!vchTelefono)
                    strColoniaEmisor = IIf(IsNull(rsCnEmpresaContable!VCHCOLONIA), "", rsCnEmpresaContable!VCHCOLONIA)
                    strCPEmisor = IIf(IsNull(rsCnEmpresaContable!VCHCODIGOPOSTAL), "", rsCnEmpresaContable!VCHCODIGOPOSTAL)
                    
                    '- Ciudad -'
                    lstrsql = "SELECT * FROM Ciudad WHERE intCveCiudad = " & CStr(rsCnEmpresaContable!INTCVECIUDAD)
                    Set rs = frsRegresaRs(lstrsql)
                    If Not rs.EOF Then
                        strCiudadEmisor = rs!VCHDESCRIPCION
                        '- Estado -'
                        lstrsql = "SELECT * FROM Estado WHERE intCveEstado = " & CStr(rs!INTCVEESTADO)
                        rs.Close
                        Set rs = frsRegresaRs(lstrsql)
                        If Not rs.EOF Then
                            strEstadoEmisor = rs!VCHDESCRIPCION
                            '- País -'
                            lstrsql = "SELECT * FROM Pais WHERE intCvePais = " & CStr(rs!intCvePais)
                            rs.Close
                            Set rs = frsRegresaRs(lstrsql)
                            If Not rs.EOF Then
                                strPaisEmisor = rs!VCHDESCRIPCION
                            End If
                        End If
                    End If
                    rs.Close
                                    
                                    
                    '- Ciudad -'
                    lstrsql = "SELECT * FROM Ciudad WHERE intCveCiudad = " & lngCveCiudad
                    Set rs = frsRegresaRs(lstrsql)
                    If Not rs.EOF Then
                        strCiudadReceptor = rs!VCHDESCRIPCION
                        '- Estado -'
                        lstrsql = "SELECT * FROM Estado WHERE intCveEstado = " & CStr(rs!INTCVEESTADO)
                        rs.Close
                        Set rs = frsRegresaRs(lstrsql)
                        If Not rs.EOF Then
                            strEstadoReceptor = rs!VCHDESCRIPCION
                            '- País -'
                            lstrsql = "SELECT * FROM Pais WHERE intCvePais = " & CStr(rs!intCvePais)
                            rs.Close
                            Set rs = frsRegresaRs(lstrsql)
                            If Not rs.EOF Then
                                strPaisReceptor = rs!VCHDESCRIPCION
                            End If
                        End If
                    End If
                    rs.Close
                                    
                                    
                End If
                rsCnEmpresaContable.Close
                If vgstrVersionCFDI = "4.0" Then
                    vlstrRegimenFiscal = vlstrRegimenFiscal
                Else
                    vlstrRegimenFiscal = ""
                End If
                
                strRFCEmisor = Trim(Replace(Replace(Replace(strRFCEmisor, "-", ""), "_", ""), " ", ""))
                vlstrDFRFC = Trim(Replace(Replace(Replace(vlstrDFRFC, "-", ""), "_", ""), " ", ""))
                strParametros = CStr(lngNumPagoSalida) & "|" & "AN" & "|" & CStr(vgintClaveEmpresaContable) _
                                & "|" & IIf(Trim(strAnoAprobacion) = "0", "", Trim(strAnoAprobacion)) & "|" & Trim(strNumeroAprobacion) & "|" & strSerie & "|" & strFolio _
                                & "|" & vlstrDFNombre & "|" & vlstrDFRFC & "|" & vlstrDFDireccion _
                                & "|" & vlstrDFColonia & "|" & vlstrDFCodigoPostal & "|" & strCiudadReceptor & "|" & strEstadoReceptor _
                                & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora) & "|" & 0 & "|" & strNombreEmisor & "|" & strRFCEmisor _
                                & "|" & strCalleEmisor & "|" & strTelefonoEmisor & "|" & strColoniaEmisor _
                                & "|" & strCPEmisor & "|" & strCiudadEmisor & "|" & strCiudadReceptor _
                                & "|" & strCiudadEmisor & "|" & strEstadoEmisor & "|" & strPaisEmisor & "|" & strPaisReceptor _
                                & "|" & "" _
                                & "|" & vgstrVersionCFDI _
                                & "|" & vlstrDFNumExterior _
                                & "|" & vlstrDFNumInterior _
                                & "|" & vlstrRegimenFiscal
                Set rsSec_gncomproFD = frsEjecuta_SP(strParametros, "sp_PVInsDatosAntCFDI")
                
                If rsSec_gncomproFD.RecordCount > 0 Then lnSec_gncomproFD = rsSec_gncomproFD!numero
                
                pEjecutaSentencia "update PVPago set intIdComprobante = " & lnSec_gncomproFD & " where intNumPago = " & lngNumPagoSalida
                pEjecutaSentencia "UPDATE GNCOMPROBANTEFISCALDIGITAL SET VCHNOMBRERECEPTOR = '" & vlstrDFNombre & "' WHERE  INTIDCOMPROBANTE = " & lnSec_gncomproFD
            End If
            'Si se realizará una emisión digital
            'Barra de progreso CFD
            pgbBarraCFD.Value = 70
            'freBarraCFD.Top = 3200
            Screen.MousePointer = vbHourglass
            lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere..."
            freBarraCFD.Visible = True
            freBarraCFD.Refresh
            frmEntradaSalidaDinero.Enabled = False
            
            pLogTimbrado 2
            If Not fblnGeneraComprobanteDigital(lngNumPagoSalida, "AN", 0, CInt(strAnoAprobacion), strNumeroAprobacion, True, False) Then
                On Error Resume Next
                If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
                   'El donativo se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
                   MsgBox SIHOMsg(1306), vbInformation + vbOKOnly, "Mensaje"
                  '------------------------------------------------------------------
                  'Identificamos el donativo como pendiente de confirmar timbre fiscal
                  '------------------------------------------------------------------
                    If blnGeneraCFD Then
                        pMarcarPendienteTimbre lngNumPagoSalida, "AN", vgintNumeroDepartamento
                    End If
                    pGeneraCFDI = 1
                ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then  'No se realizó el timbrado
                  EntornoSIHO.ConeccionSIHO.RollbackTrans
                  pGeneraCFDI = 2
                  'Se restaura el estatus del cursor del puntero
                  Screen.MousePointer = vbDefault
                  freBarraCFD.Visible = False
                  frmEntradaSalidaDinero.Enabled = True
                  pLogTimbrado 1
                  Exit Function
                End If
            End If
           'Se consultan las rutas del CFD y CFDi para la impresión
            strSentencia = "SELECT TRIM(VCHSERIECOMPROBANTE) || TRIM(VCHFOLIOCOMPROBANTE) Folio FROM GNCOMPROBANTEFISCALDIGITAL WHERE intComprobante = " & lngNumPagoSalida & " AND CHRTIPOCOMPROBANTE = 'AN'"
            Set rsPago = frsRegresaRs(strSentencia)
            strSentencia = "SELECT VCHRUTAXML, VCHRUTAPDF FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
            Set rsRutas = frsRegresaRs(strSentencia)
            strRuta = rsRutas!vchRutaPDF & "\" & rsPago!Folio & ".pdf"
           'Barra de progreso CFD
            pgbBarraCFD.Value = 100
            'freBarraCFD.Top = 3200
            Screen.MousePointer = vbDefault
            lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere..."
            freBarraCFD.Visible = True
            freBarraCFD.Refresh
            freBarraCFD.Visible = False
            frmEntradaSalidaDinero.Enabled = True
            pLogTimbrado 1
    End If

End Function
Private Sub pFacturasDirectasAnteriores()
    Dim rsFacturas As New ADODB.Recordset
    'Dim vlstrsql As String
    Dim i As Integer
    
    vlstrsql = ""
      
    vlstrsql = "select PVPAGO.CHRFOLIORECIBO chrfoliofactura,PVPAGO.DTMFECHA DTMFECHAHORA,PVPAGO.MNYCANTIDAD mnyTotalFactura, CASE WHEN PVPAGO.BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, PVPAGO.MNYTIPOCAMBIO " & _
               "from PVPago " & _
               "inner join GNComprobanteFiscalDigital on GNComprobanteFiscalDigital.INTCOMPROBANTE = PVPago.intNumPago and GNComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = 'AN' " & _
               "where NOT VCHUUID IS NULL and PVPago.BITCANCELADO = 1 and pvpago.chrfoliofactura is null and pvpago.INTMOVPACIENTE = " & txtMovimientoPaciente
               
   Set rsFacturas = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
  
   If rsFacturas.RecordCount <> 0 Then
       chkFacturaSustitutaDFP.Enabled = True
       lstFacturaASustituirDFP.Enabled = True
   
       lstFacturaASustituirDFP.Clear
    Else
       chkFacturaSustitutaDFP.Enabled = False
       lstFacturaASustituirDFP.Enabled = False
    End If
    rsFacturas.Close
    vlnblnLocate = False
End Sub

Private Sub pRefacturacionAnteriores()
Dim i As Integer

    If chkFacturaSustitutaDFP.Value = 1 And lstFacturaASustituirDFP.ListCount > 0 Then
            For i = 0 To UBound(aFoliosPrevios())
                If aFoliosPrevios(i).chrfoliofactura <> "" Then
                    pEjecutaSentencia "INSERT INTO PVPAGOREFACTURACION (chrFoliopagoActivada, chrFoliopagoCancelada) " & " VALUES ('" & Trim(txtFolioRecibo.Text) & "', '" & aFoliosPrevios(i).chrfoliofactura & "')"
                End If
            Next i
        End If
End Sub
Private Sub pCargaRegimenFiscal()
    Dim rsTmp As ADODB.Recordset
        Set rsTmp = frsRegresaRs("SELECT VCHCLAVE, VCHDESCRIPCION FROM GNCATALOGOSATDETALLE WHERE INTIDCATALOGOSAT = 1", adLockReadOnly, adOpenForwardOnly)
        If Not rsTmp.EOF Then
            pLlenarCboRs frmDatosFiscales.cboRegimenFiscal, rsTmp, 0, 1
            frmDatosFiscales.cboRegimenFiscal.AddItem "<NINGUNO>", 0
            frmDatosFiscales.cboRegimenFiscal.ItemData(frmDatosFiscales.cboRegimenFiscal.newIndex) = 0
            frmDatosFiscales.cboRegimenFiscal.ListIndex = 0
        End If
        rsTmp.Close
End Sub

