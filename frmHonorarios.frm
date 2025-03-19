VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHonorarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Honorarios médicos"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10245
   ScaleMode       =   0  'User
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdComisiones 
      Height          =   855
      Left            =   4200
      TabIndex        =   83
      Top             =   13200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1508
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDetalleCorte 
      Height          =   1455
      Left            =   480
      TabIndex        =   82
      Top             =   11640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2566
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraComisiones 
      Height          =   2790
      Left            =   1635
      TabIndex        =   61
      Top             =   5400
      Visible         =   0   'False
      Width           =   6060
      Begin VB.CommandButton cmdAceptarComisiones 
         Caption         =   "&Aceptar"
         Height          =   435
         Left            =   2370
         TabIndex        =   71
         Top             =   2235
         Width           =   1365
      End
      Begin VB.ListBox lstComisionesAsignadas 
         Height          =   1230
         Left            =   3525
         TabIndex        =   69
         Top             =   750
         Width           =   2415
      End
      Begin VB.Frame freBotones 
         Height          =   1335
         Left            =   2625
         TabIndex        =   68
         Top             =   645
         Width           =   810
         Begin VB.CommandButton cmdSelecciona 
            Caption         =   "Inluir"
            Height          =   510
            Index           =   0
            Left            =   90
            MaskColor       =   &H80000014&
            Picture         =   "frmHonorarios.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Asignar un paquete al paciente"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   630
         End
         Begin VB.CommandButton cmdSelecciona 
            Caption         =   "Excluir"
            Enabled         =   0   'False
            Height          =   510
            Index           =   1
            Left            =   90
            MaskColor       =   &H80000014&
            Picture         =   "frmHonorarios.frx":017A
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Excluir un paquete al paciente"
            Top             =   750
            UseMaskColor    =   -1  'True
            Width           =   630
         End
      End
      Begin VB.ListBox lstComisiones 
         Height          =   1230
         Left            =   120
         TabIndex        =   63
         Top             =   750
         Width           =   2415
      End
      Begin VB.CommandButton cmdCerrarComisiones 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5730
         TabIndex        =   66
         Top             =   180
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   90
         X2              =   5925
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   5940
         Y1              =   2085
         Y2              =   2085
      End
      Begin VB.Label Label20 
         Caption         =   "Comisiones asignadas"
         Height          =   240
         Left            =   3510
         TabIndex        =   70
         Top             =   525
         Width           =   1560
      End
      Begin VB.Label Label16 
         Caption         =   "Comisiones"
         Height          =   210
         Left            =   135
         TabIndex        =   67
         Top             =   510
         Width           =   1350
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Selección de comisiones para honorarios médicos"
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
         Height          =   270
         Left            =   90
         TabIndex        =   62
         Top             =   165
         Width           =   4365
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   30
         Top             =   135
         Width           =   5985
      End
      Begin VB.Shape Shape2 
         Height          =   2700
         Left            =   0
         Top             =   90
         Width           =   6060
      End
   End
   Begin TabDlg.SSTab SSTabHonorario 
      Height          =   10900
      HelpContextID   =   1
      Left            =   -165
      TabIndex        =   44
      Top             =   -600
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   19235
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmHonorarios.frx":02F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRecibo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraBotonera"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frePaciente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraReciboHonorario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraHonorarios"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CommonDialog1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraCuentaPagar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmHonorarios.frx":0310
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "cmdActualizarRecibo"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraCuentaPagar 
         Caption         =   "Información de la cuenta por pagar"
         Enabled         =   0   'False
         Height          =   2115
         Left            =   255
         TabIndex        =   128
         Top             =   5970
         Visible         =   0   'False
         Width           =   12330
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHonoCxPProv 
            Height          =   1650
            Left            =   115
            TabIndex        =   129
            Top             =   300
            Width           =   12080
            _ExtentX        =   21299
            _ExtentY        =   2910
            _Version        =   393216
            FixedRows       =   0
            GridColor       =   -2147483631
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   9480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdActualizarRecibo 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -67620
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmHonorarios.frx":032C
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Actualizar la información del recibo"
         Top             =   9855
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -75000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmHonorarios.frx":066E
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Guardar"
         Top             =   0
         Width           =   495
      End
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   -69982
         TabIndex        =   111
         Top             =   9690
         Width           =   2920
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   1860
            Picture         =   "frmHonorarios.frx":09B0
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   "Cancelar honorario"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPagoAnticipado 
            Caption         =   "Pagar por anticipado"
            Height          =   495
            Left            =   60
            TabIndex        =   114
            Top             =   170
            Width           =   1800
         End
      End
      Begin VB.Frame fraHonorarios 
         Height          =   4080
         Left            =   255
         TabIndex        =   80
         Top             =   5970
         Visible         =   0   'False
         Width           =   12330
         Begin VB.CheckBox chkDolares 
            Caption         =   "Registrar los honorarios en dólares"
            Height          =   225
            Left            =   60
            TabIndex        =   95
            Top             =   2880
            Width           =   2760
         End
         Begin VB.CheckBox chkMuestraDetalle 
            Caption         =   "Mostrar los descuentos en el recibo"
            Enabled         =   0   'False
            Height          =   225
            Left            =   60
            TabIndex        =   94
            Top             =   3120
            Width           =   2955
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargaHonorarios 
            Height          =   1650
            Left            =   60
            TabIndex        =   81
            Top             =   165
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   2910
            _Version        =   393216
            FixedRows       =   0
            GridColor       =   -2147483631
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblTotalRTP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   131
            Top             =   2580
            Width           =   1695
         End
         Begin VB.Label Label29 
            Caption         =   "Retención RTP"
            Height          =   255
            Left            =   9075
            TabIndex        =   130
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblTotalMonto 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   93
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lblTotalTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   92
            Top             =   3600
            Width           =   1695
         End
         Begin VB.Label lblTotalComisionIVA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   91
            Top             =   3240
            Width           =   1695
         End
         Begin VB.Label lblTotalComisiones 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   90
            Top             =   2910
            Width           =   1695
         End
         Begin VB.Label lblTotalRetencion 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   10515
            TabIndex        =   89
            Top             =   2250
            Width           =   1695
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Total "
            Height          =   195
            Left            =   9075
            TabIndex        =   88
            Top             =   3600
            Width           =   510
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "IVA comisiones "
            Height          =   195
            Left            =   9075
            TabIndex        =   87
            Top             =   3300
            Width           =   1230
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Comisiones "
            Height          =   195
            Left            =   9075
            TabIndex        =   86
            Top             =   2970
            Width           =   945
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Retención de ISR"
            Height          =   195
            Left            =   9075
            TabIndex        =   85
            Top             =   2310
            Width           =   1380
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Monto "
            Height          =   195
            Left            =   9075
            TabIndex        =   84
            Top             =   1980
            Width           =   600
         End
      End
      Begin VB.Frame fraReciboHonorario 
         Height          =   360
         Left            =   300
         TabIndex        =   76
         Top             =   3600
         Width           =   2820
         Begin VB.TextBox txtReciboHonorario 
            Height          =   315
            Left            =   1515
            MaxLength       =   10
            TabIndex        =   13
            ToolTipText     =   "Número de recibo de honorario del médico"
            Top             =   45
            Width           =   1215
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Número de recibo "
            Height          =   195
            Left            =   60
            TabIndex        =   77
            Top             =   105
            Width           =   1305
         End
      End
      Begin VB.Frame frePaciente 
         Height          =   1375
         Left            =   255
         TabIndex        =   55
         Top             =   540
         Width           =   12330
         Begin VB.TextBox txtMovimientoPaciente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   "Número de cuenta del paciente"
            Top             =   240
            Width           =   1260
         End
         Begin VB.TextBox txtPaciente 
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   4
            ToolTipText     =   "Nombre del paciente"
            Top             =   585
            Width           =   5385
         End
         Begin VB.TextBox txtTipoPaciente 
            Height          =   315
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Tipo de paciente"
            Top             =   930
            Width           =   5385
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   390
            Left            =   2865
            TabIndex        =   56
            Top             =   180
            Width           =   3360
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "&Consulta externa"
               Height          =   195
               Index           =   2
               Left            =   1800
               TabIndex        =   113
               Top             =   135
               Width           =   1500
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "&Externo"
               Height          =   195
               Index           =   1
               Left            =   885
               TabIndex        =   2
               Top             =   135
               Width           =   855
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "&Interno"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   1
               Top             =   120
               Value           =   -1  'True
               Width           =   825
            End
         End
         Begin MSMask.MaskEdBox mskFechaRegistro 
            Height          =   315
            Left            =   8505
            TabIndex        =   97
            ToolTipText     =   "Fecha de atención del médico"
            Top             =   585
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label lblFechaRegistro 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de registro"
            Height          =   195
            Left            =   7170
            TabIndex        =   96
            Top             =   645
            Width           =   1230
         End
         Begin VB.Label lblEstadoHonorario 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   8505
            TabIndex        =   3
            Top             =   945
            Width           =   3690
         End
         Begin VB.Label lblTipoPago 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   7170
            TabIndex        =   74
            Top             =   1035
            Width           =   495
         End
         Begin VB.Label lblCuentaPaciente 
            AutoSize        =   -1  'True
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   75
            TabIndex        =   59
            Top             =   300
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   75
            TabIndex        =   58
            Top             =   645
            Width           =   555
         End
         Begin VB.Label lblTipoEmpresa 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   75
            TabIndex        =   57
            Top             =   990
            Width           =   1200
         End
      End
      Begin VB.Frame fraBotonera 
         Height          =   705
         Left            =   3850
         TabIndex        =   46
         Top             =   10020
         Width           =   5015
         Begin VB.CommandButton cmdPagoEfectivo 
            Caption         =   "Pagar con efectivo"
            Height          =   495
            Left            =   3525
            TabIndex        =   43
            ToolTipText     =   "Pagar con efectivo los honorarios "
            Top             =   150
            Width           =   1425
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmHonorarios.frx":0EA2
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   555
            Picture         =   "frmHonorarios.frx":12A4
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Anterior"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1050
            Picture         =   "frmHonorarios.frx":1416
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1545
            Picture         =   "frmHonorarios.frx":1588
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Siguiente"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2040
            Picture         =   "frmHonorarios.frx":16FA
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Ultimo"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Enabled         =   0   'False
            Height          =   495
            Left            =   2535
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmHonorarios.frx":186C
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Guardar"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   3030
            Picture         =   "frmHonorarios.frx":1BAE
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Imprimir"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Height          =   8835
         Left            =   -74790
         TabIndex        =   45
         Top             =   540
         Width           =   12375
         Begin VB.CommandButton cmdLimpiarCarta 
            Enabled         =   0   'False
            Height          =   375
            Left            =   8190
            Picture         =   "frmHonorarios.frx":1D50
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Borrar archivo"
            Top             =   8340
            Width           =   375
         End
         Begin VB.CommandButton cmdAbrirCarta 
            Enabled         =   0   'False
            Height          =   375
            Left            =   7815
            Picture         =   "frmHonorarios.frx":2182
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Guardar archivo y abrir"
            Top             =   8340
            Width           =   375
         End
         Begin VB.CommandButton cmdCargarCarta 
            Enabled         =   0   'False
            Height          =   375
            Left            =   7440
            Picture         =   "frmHonorarios.frx":25B4
            Style           =   1  'Graphical
            TabIndex        =   124
            ToolTipText     =   "Cargar archivo"
            Top             =   8340
            Width           =   375
         End
         Begin VB.CommandButton cmdMarcar 
            Caption         =   "Marcar / Desmarcar"
            Height          =   375
            Left            =   8655
            TabIndex        =   119
            Top             =   8340
            Width           =   1800
         End
         Begin VB.CommandButton cmdInvertirSel 
            Caption         =   "Invertir selección"
            Height          =   375
            Left            =   10470
            TabIndex        =   118
            Top             =   8340
            Width           =   1800
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar información"
            Height          =   345
            Left            =   8475
            TabIndex        =   116
            ToolTipText     =   "Cargar los honorarios "
            Top             =   1095
            Width           =   3795
         End
         Begin VB.Frame Frame7 
            Height          =   1320
            Left            =   105
            TabIndex        =   101
            Top             =   120
            Width           =   8265
            Begin VB.TextBox txtFilCuenta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1530
               TabIndex        =   105
               ToolTipText     =   "Número de cuenta del paciente"
               Top             =   480
               Width           =   1020
            End
            Begin VB.OptionButton optFilTipoPaciente 
               Caption         =   "Interno"
               Height          =   195
               Index           =   0
               Left            =   1530
               TabIndex        =   104
               ToolTipText     =   "Tipo de paciente interno"
               Top             =   225
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton optFilTipoPaciente 
               Caption         =   "Externo"
               Height          =   195
               Index           =   1
               Left            =   2385
               TabIndex        =   103
               ToolTipText     =   "Tipo de paciente externo"
               Top             =   225
               Width           =   855
            End
            Begin VB.TextBox txtFilPaciente 
               Height          =   315
               Left            =   2580
               Locked          =   -1  'True
               TabIndex        =   102
               ToolTipText     =   "Nombre del paciente"
               Top             =   480
               Width           =   5535
            End
            Begin VB.ComboBox cboFilMedico 
               Height          =   315
               Left            =   1530
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   107
               ToolTipText     =   "Médico"
               Top             =   855
               Width           =   6585
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Número de cuenta"
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   540
               Width           =   1320
            End
            Begin VB.Label lbl24 
               AutoSize        =   -1  'True
               Caption         =   "Médico"
               Height          =   195
               Left            =   120
               TabIndex        =   106
               Top             =   915
               Width           =   525
            End
         End
         Begin VB.Frame Frame4 
            Height          =   930
            Left            =   8475
            TabIndex        =   98
            Top             =   120
            Width           =   3795
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   315
               Left            =   675
               TabIndex        =   109
               ToolTipText     =   "Fecha de inicio de atención"
               Top             =   465
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox mskFin 
               Height          =   315
               Left            =   2460
               TabIndex        =   110
               ToolTipText     =   "Fecha de fin de atención"
               Top             =   465
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               Format          =   "dd/mmm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   " "
            End
            Begin VB.Label Label1 
               Caption         =   "Rango fechas de atención"
               Height          =   180
               Left            =   135
               TabIndex        =   115
               Top             =   195
               Width           =   2010
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Left            =   1965
               TabIndex        =   100
               Top             =   525
               Width           =   420
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Left            =   135
               TabIndex        =   99
               Top             =   525
               Width           =   465
            End
         End
         Begin VSFlex7LCtl.VSFlexGrid grdHonorarios 
            Height          =   6780
            Left            =   105
            TabIndex        =   112
            ToolTipText     =   "Lista de créditos según los filtros seleccionados"
            Top             =   1485
            Width           =   12165
            _cx             =   21458
            _cy             =   11959
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
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmHonorarios.frx":29E6
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
            Editable        =   1
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
      End
      Begin VB.Frame fraRecibo 
         Height          =   4080
         Left            =   255
         TabIndex        =   47
         Top             =   1900
         Width           =   12330
         Begin VB.TextBox txtRetencionRTP 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Retención de RTP que se hace al médico"
            Top             =   900
            Width           =   1500
         End
         Begin VB.CheckBox chkRetencionRTP 
            Caption         =   "Retención RTP"
            Height          =   195
            Left            =   7155
            TabIndex        =   24
            ToolTipText     =   "Retención de RTP (Impuesto sobre remuneraciones al trabajo personal no subordinado)"
            Top             =   980
            Width           =   1500
         End
         Begin VB.TextBox txtEmail 
            Height          =   315
            Left            =   4680
            MaxLength       =   100
            TabIndex        =   19
            ToolTipText     =   "Correo electrónico para envío de CFDI"
            Top             =   3690
            Width           =   2280
         End
         Begin VB.TextBox txtAdjuntar 
            Height          =   315
            Left            =   8475
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Carta de autorización"
            Top             =   3600
            Width           =   3000
         End
         Begin VB.CommandButton cmdAdjuntar 
            Caption         =   "Adjuntar carta"
            Height          =   315
            Left            =   7080
            TabIndex        =   33
            ToolTipText     =   "Adjuntar carta de autorización"
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   495
            Left            =   8475
            MaxLength       =   500
            TabIndex        =   32
            ToolTipText     =   "Observaciones"
            Top             =   3000
            Width           =   3720
         End
         Begin VB.CheckBox chkRequiereCFDI 
            Caption         =   "Requiere CFDI"
            Height          =   195
            Left            =   4680
            TabIndex        =   12
            ToolTipText     =   "Requiere CFDI"
            Top             =   1480
            Width           =   2175
         End
         Begin VB.ComboBox cboUsoCFDI 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Uso del CFDI"
            Top             =   2430
            Width           =   5400
         End
         Begin VB.CheckBox chkHonorarioFacturado 
            Caption         =   "Honorario médico ya facturado"
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            ToolTipText     =   "Honorario médico ya facturado"
            Top             =   1200
            Width           =   3135
         End
         Begin VB.CheckBox chkPagoDirecto 
            Caption         =   "Pago de honorarios directo al médico"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1560
            TabIndex        =   11
            ToolTipText     =   "Pago de honorarios directo al médico"
            Top             =   1480
            Width           =   3015
         End
         Begin VB.ComboBox cboTarifa 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   22
            ToolTipText     =   "Selección de la tarifa de ISR"
            Top             =   525
            Width           =   1620
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1500
         End
         Begin VB.TextBox txtCantidadComision 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Monto de comisiones"
            Top             =   1620
            Width           =   1500
         End
         Begin VB.TextBox txtRetencion 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Retención que se le hace al médico"
            Top             =   525
            Width           =   1500
         End
         Begin VB.TextBox txtNetoPagar 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Monto de los honorarios menos retención"
            Top             =   1260
            Width           =   1500
         End
         Begin VB.TextBox txtIVAComision 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "IVA de las comisiones"
            Top             =   1975
            Width           =   1500
         End
         Begin VB.TextBox txtPagoCredito 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   10695
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2310
            Width           =   1500
         End
         Begin VB.TextBox txtMontoHonorario 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10695
            TabIndex        =   20
            ToolTipText     =   "Monto de los honorarios"
            Top             =   180
            Width           =   1500
         End
         Begin VB.CommandButton cmdAgregar 
            Height          =   375
            Left            =   11670
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmHonorarios.frx":2ABA
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Agregar honorario"
            Top             =   3550
            UseMaskColor    =   -1  'True
            Width           =   525
         End
         Begin MSMask.MaskEdBox mskFechaAtencionFin 
            Height          =   315
            Left            =   3135
            TabIndex        =   9
            ToolTipText     =   "Fecha de atención del médico"
            Top             =   870
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtConcepto 
            Height          =   315
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Concepto del honorario médico"
            Top             =   2080
            Width           =   5400
         End
         Begin VB.TextBox txtReciboNombre 
            Height          =   315
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   16
            ToolTipText     =   "Persona a quien esta dirigido el recibo"
            Top             =   2780
            Width           =   5400
         End
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   1560
            MaxLength       =   13
            TabIndex        =   18
            ToolTipText     =   "Registro Federal de Contribuyentes"
            Top             =   3690
            Width           =   1515
         End
         Begin VB.TextBox txtDireccion 
            Height          =   495
            Left            =   1560
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            ToolTipText     =   "Dirección"
            Top             =   3130
            Width           =   5400
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Enabled         =   0   'False
            Height          =   2505
            Left            =   9690
            TabIndex        =   72
            Top             =   510
            Width           =   1770
         End
         Begin VB.CheckBox chkRetencion 
            Caption         =   "Retención ISR"
            Height          =   195
            Left            =   7155
            TabIndex        =   21
            ToolTipText     =   "Retención del ISR"
            Top             =   585
            Value           =   1  'Checked
            Width           =   1350
         End
         Begin VB.CheckBox chkComisiones 
            Caption         =   "Comisiones"
            Enabled         =   0   'False
            Height          =   225
            Left            =   7155
            TabIndex        =   27
            ToolTipText     =   "Comisiones"
            Top             =   1670
            Width           =   1110
         End
         Begin VB.ComboBox cboMedicos 
            Height          =   315
            ItemData        =   "frmHonorarios.frx":2FAC
            Left            =   1560
            List            =   "frmHonorarios.frx":2FAE
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Selección del médico"
            Top             =   180
            Width           =   5400
         End
         Begin VB.CommandButton cmdTodosMedicos 
            Caption         =   "Mostrar todos los médicos"
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   545
            Width           =   2055
         End
         Begin MSMask.MaskEdBox mskFechaAtencion 
            Height          =   315
            Left            =   1560
            TabIndex        =   8
            ToolTipText     =   "Fecha de atención del médico"
            Top             =   870
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
            Caption         =   "Correo electrónico"
            Height          =   195
            Left            =   3280
            TabIndex        =   127
            Top             =   3750
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   7155
            TabIndex        =   123
            Top             =   3100
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Uso del CFDI"
            Height          =   195
            Left            =   105
            TabIndex        =   122
            Top             =   2490
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de atención médica"
            Height          =   375
            Left            =   105
            TabIndex        =   78
            Top             =   800
            Width           =   1185
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "a"
            Height          =   195
            Left            =   2880
            TabIndex        =   79
            Top             =   930
            Width           =   90
         End
         Begin VB.Label lblPagosCredito 
            AutoSize        =   -1  'True
            Caption         =   "Pagos de crédito"
            Enabled         =   0   'False
            Height          =   195
            Left            =   7155
            TabIndex        =   75
            Top             =   2360
            Width           =   1200
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "IVA comisiones"
            Height          =   195
            Left            =   7155
            TabIndex        =   73
            Top             =   2005
            Width           =   1080
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   7155
            TabIndex        =   60
            Top             =   2690
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Médico"
            Height          =   195
            Left            =   105
            TabIndex        =   54
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            Height          =   195
            Left            =   105
            TabIndex        =   53
            Top             =   2140
            Width           =   690
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Recibo a nombre"
            Height          =   195
            Left            =   105
            TabIndex        =   52
            Top             =   2840
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
            Height          =   195
            Left            =   105
            TabIndex        =   51
            Top             =   3190
            Width           =   675
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Left            =   105
            TabIndex        =   50
            Top             =   3750
            Width           =   315
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Monto del honorario"
            Height          =   195
            Left            =   7155
            TabIndex        =   49
            Top             =   240
            Width           =   1410
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   7155
            TabIndex        =   48
            Top             =   1310
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "frmHonorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------
' Programa para registro, consulta y cancelación de honorarios
' Fecha de programación:  Martes 20 de Octubre de 2002
'------------------------------------------------------------------------------------------

Option Explicit

'Columnas de la búsqueda de honorarios
Const cintColFilConsecutivo = 1
Const cintColFilCuenta = 2
Const cintColFilTipo = 3
Const cintColFilPaciente = 4
Const cintColFilMedico = 5
Const cintColFilInicioAtencion = 6
Const cintColFilFinAtencion = 7
Const cintColFilRecibo = 8
Const cintColFilConcepto = 9
Const cintColFilMonto = 10
Const cintColFilRetencion = 11
Const cintColFilSubtotal = 12
Const cintColFilComision = 13
Const cintColFilIVA = 14
Const cintColFilTotal = 15
Const cintColFilEstado = 16
Const cintColFilFechaEnvio = 17
Const cintColFilCliente = 18
Const cintColFilEmpleado = 19
Const cintColFilEstatus = 20
Const cintColFilCveMedico = 21
Const cintColFilCxP = 22
Const cintColFilCarta = 23
Const cintColFilCartaMod = 24
Const cstrColumnas = "|Número|Cuenta|Tipo|Paciente|Médico|Inicio atención|Término atención|Recibo|Concepto|Monto|Retención|Subtotal|Comisión|IVA|Total|Estado|Envío cobro|Cliente|Persona registró|Estatus|Clave Médico|Cuenta por pagar|Carta de autorización|Carta_Mod"

'Constantes del gripd de guardado:
Const cintColCveMedico = 1
Const cintColCuentaContableMedico = 2
Const cintColNombreMedico = 3
Const cintColMonto = 4
Const cintColRetencion = 5
Const cintColComision = 6
Const cintColIVAComision = 7
Const cintColTotalPagar = 8
Const cintColReciboHonorario = 9
Const cintColFechaAtencionInicio = 10
Const cintColFechaAtencionFin = 11
Const cintColConcepto = 12
Const cintColReciboNombre = 13
Const cintColDireccion = 14
Const cIntColRFC = 15
Const cintColIdHonorario = 16
Const cintColIdTarifa = 17
Const cintColPagoDirecto = 18
Const cintColCarta = 19
Const cintColCartaCompleta = 20
Const cintColObservaciones = 21
Const cintColRetencionRTP = 22
Const cintColRetencionISR = 23

Private Type Comisiones
    vllngCveComision As Long
    vldblCantidad   As Double
    vldblIVA As Double
End Type

Private Type typTarifa
    lngIdTarifa As Long
    dblPorcentaje As Double
End Type

Dim vlstrSentencia As String
Dim vlintlimpiar As Integer
Dim vlblnLimpiar As Boolean
Dim vlblnConsulta As Boolean
Dim vlblnPacienteSeleccionado As Boolean
Dim vlblnEntrando As Boolean
Dim vgintTipoPaciente As Integer            'Para saber el tipo de paciente al momento de grabar
Dim vgintEmpresa As Long                 'La clave de la empresa pa traerla por toda la pantalla
Dim vgintCveExtra As Long
Dim llngTipoConvenio As Long 'Para el tipo de convenio del paciente

Dim vlComisiones() As Comisiones            'Arreglo para traer las comisiones que se le ponen a cada Honorario
Dim aFormasPago() As FormasPago
Dim arrTarifas() As typTarifa               'Tarifas de ISR
Dim vlintInterno As Integer

Public vllngNumeroOpcion As Long            'No de opción para el módulo que manda llamar este proceso

Dim rsPvSelComision As New ADODB.Recordset
Dim rsPvSelDatosPaciente As New ADODB.Recordset
Dim vllngCorteGrabando As Long
Dim vllngNumeroCorte As Long                'Numero de corte actual

Public vllngNumCuentaPagar  As Long 'Variable que almacen al consecutivo de la tabla cpcuentapagarmedico
Private vgrptReporte As CRAXDRT.Report
Public vglngCuentaPaciente As Long     'Variable que almacena el numero de cuenta del paciente
Public vldblHonorarioTotal As Double                 'Guarda el monto de los honorarios
Public vlintContador As Integer                'Variable auxiliar usada para realizar ciclos
Public vllngCuentaMedico As Long                   'Trae la cuenta acreedora del médico
Public vlblnContinuaRevision As Boolean        'Bandera para saber si se agrega o no un honorario repetido al grid

Dim vldblPagoCuenta As Double 'Pagos efectuados al honorario a crédito,?
Dim vllngNumHonorario As Long 'Consecutivo de PvHonorario
Dim vlblnEntraCorte As Boolean 'Indica si los honorarios que se pagan en efectivo o crédito afectan el corte
Dim rsDatosCliente As New ADODB.Recordset 'Se actualiza en la validacion de datos y se usa al guardar, datos del cliente, cuando los honorarios son a crédito
Dim vllngCtaHonorariosPagar As Long 'Cuenta contable para honorarios por pagar, esta debe ser una cuenta puente
Dim vllngCtaHonorariosCobrar As Long 'Cuenta contable para honorarios por cobrar, esta debe ser una cuenta puente
Dim vldblTipoCambio As Double 'Tipo de cambio del día
Dim vlintMoneda As Integer 'Moneda del honorario 1 = pesos, 0 = dólares
Dim vldblTipoCambioHonorario As Double 'Tipo de cambio al que fué registrado el honorario

Dim lngDeptoCliente As Long 'Departamento al que pertenece el cliente relacionado con la cuenta
Dim llngNumMedico As Long 'Num. de médico con que se abrirá la cuenta del paciente de consulta externa
Dim lstrAfiliacion As String 'Num. de afiliación del paciente
Dim lblnRecargarTarifas As Boolean 'Para saber cuando cargar el catálogo de tarifas
Dim vlblnesconvenio As Boolean
Dim vlblnEsCredito As Boolean
Dim vllngCuentaPagar As Long '0 = No existe cuentapagarmedico relacionada al honorario, 1 = Si existe
Dim ldtmFecha As Date 'Fecha actual
Dim llngTotalSel As Long 'No. de honorarios seleccionados

Dim vllngCtaCostoHonorarioFacturado As Long   'Cuenta contable de costo para los honorarios facturados

Dim lstrRFCConvPers As String
Dim lstrDireccionConvPers  As String
Dim lstrNombreConvPers As String
Dim lstrCorreoElectronico As String

Dim lstrRFCConvPersUsoCFDI As String
Dim lstrDireccionConvPersUsoCFDI  As String
Dim lstrNombreConvPersUsoCFDI As String
Dim lstrCorreoElectronicoUsoCFDI As String
Dim vlstrcveempresa As String

Dim lblnPacienteConvenio As Boolean

Dim vlTipoPacienteCatalogo As String
Dim vlstrAdTipoPaciente As String
Dim vlstrFolioDoc As String
Dim vllngNumCorte As Long
Dim dblPorcentajeRTP As Double
Dim dblTotalRetencionDetalle As Double

'Registra el movimiento de cancelación en el libro de bancos -'
Private Sub pCancelaMovimiento(vlintNumPago As Long, vlstrFolio As String, vlStrReferencia As String, vlintCorteMovimiento As Long, vllngCorteActual As Long, vllngPersonaGraba As Long)
On Error GoTo NotificaError

    Dim rsMovimiento As ADODB.Recordset
    Dim lstrTipoDoc As String, lstrFecha As String
    Dim ldblCantidad As Double
    Dim rs As ADODB.Recordset
    Dim vlstrSentencia As String

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
                Select Case rsMovimiento!chrTipo
                    Case "E": lstrTipoDoc = "CEH" 'Efectivo
                    Case "T": lstrTipoDoc = "CTH" 'Tarjeta de crédito
                    Case "B": lstrTipoDoc = "CRH" 'Transferencia bancaria
                    Case "H": lstrTipoDoc = "CQH" 'Cheque
                End Select

                '- Cantidad negativa para que se tome como abono si se cancela una entrada de dinero, cantidad positiva si se cancela salida de dinero -'
                ldblCantidad = rsMovimiento!MNYCantidad * -1

                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rsMovimiento!intFormaPago & "|" & rsMovimiento!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rsMovimiento!MNYTIPOCAMBIO = 0, 1, 0) & "|" & rsMovimiento!MNYTIPOCAMBIO & "|" & lstrTipoDoc & "|" & vlStrReferencia & "|" & _
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

Private Function fstrTipoMovimientoForma(lintCveForma As Integer) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    
    fstrTipoMovimientoForma = ""
    
    vlstrSentencia = "SELECT * FROM PvFormaPago WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        Select Case rsForma!chrTipo
            Case "E": fstrTipoMovimientoForma = "EFH" 'Efectivo
            Case "T": fstrTipoMovimientoForma = "TAH" 'Tarjeta
            Case "B": fstrTipoMovimientoForma = "TRH" 'Transferencia
            Case "H": fstrTipoMovimientoForma = "CHH" 'Cheque
        End Select
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrTipoMovimientoForma"))
End Function

Private Sub cboFilMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then mskInicio.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboFilMedico_KeyDown"))
End Sub

Private Sub cboMedicos_Click()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim rsComisionesMedico As New ADODB.Recordset
    Dim vllngContador As Long
    Dim vlblnTermina As Boolean
    Dim rsISRMedico As New ADODB.Recordset
    Dim rs As New ADODB.Recordset

    If cboMedicos.ListIndex > -1 Then

        pCargaComisiones

        'Verificar la configuración del médico
        Set rs = frsEjecuta_SP("-1|1", "SP_CNSELTARIFAISR")
        If rs.RecordCount > 0 Then
            'RETENCION ISR
            vlstrSentencia = "select bitISR, intcveRetencion, nvl(bitRTP,0) bitRTP from HoMedico where intCveMedico = " & Str(cboMedicos.ItemData(cboMedicos.ListIndex))
            Set rsISRMedico = frsRegresaRs(vlstrSentencia)
            If rsISRMedico.RecordCount <> 0 Then
                chkRetencion.Value = IIf(IsNull(rsISRMedico!bitISR), 0, rsISRMedico!bitISR)
                chkRetencionRTP.Value = IIf(IsNull(rsISRMedico!BITRTP), 0, rsISRMedico!BITRTP) ''bitRTP
                If Not IsNull(rsISRMedico!INTCVERETENCION) Then
                    'cboTarifa
                    cboTarifa.ListIndex = fintLocalizaCbo(cboTarifa, rsISRMedico!INTCVERETENCION)
                End If
            End If
            
        Else
            chkRetencion.Enabled = False
        End If
        
        'COMISIONES
        'Si al médico se le cargan comisiones, verificar las que tiene asignadas
        vlstrSentencia = "select bitComisiones from HoMedico where intCveMedico = " & Str(cboMedicos.ItemData(cboMedicos.ListIndex))
        If CBool(frsRegresaRs(vlstrSentencia).Fields(0)) Then
            'Si tiene comisiones asignadas, cargar lo asignado.... si no, se cargan las predeterminadas
            vlstrSentencia = "select count(*) from CcComisionMedicoEmpresa where intCveMedico = " & Str(cboMedicos.ItemData(cboMedicos.ListIndex))
            If frsRegresaRs(vlstrSentencia).Fields(0) > 0 Then
                If optTipoPaciente(0).Value Then
                    vlstrSentencia = "Select tnyCveTipoPaciente*-1 CveTipoPaciente, intCveEmpresa CveEmpresa from AdAdmision where numNumCuenta = " & txtMovimientoPaciente.Text
                Else
                    vlstrSentencia = "Select tnyCveTipoPaciente*-1 CveTipoPaciente, intClaveEmpresa CveEmpresa from RegistroExterno where intNumCuenta = " & txtMovimientoPaciente.Text
                End If
                Set rsTipoPaciente = frsRegresaRs(vlstrSentencia)
                If rsTipoPaciente.RecordCount <> 0 Then
                    vlstrSentencia = "" & _
                    "select " & _
                        "PvComision.smiCveComision " & _
                    "From " & _
                        "CcComisionMedicoEmpresa " & _
                        "inner join PvComision on CcComisionMedicoEmpresa.intCveComision = PvComision.smiCveComision " & _
                    "Where " & _
                        "CcComisionMedicoEmpresa.intCveMedico = " & Str(cboMedicos.ItemData(cboMedicos.ListIndex)) & _
                        " and (CcComisionMedicoEmpresa.intCveTipoPacienteEmpresa = 0 " & _
                        " or CcComisionMedicoEmpresa.intCveTipoPacienteEmpresa = " & IIf(IsNull(rsTipoPaciente!cveTipoPaciente), 0, rsTipoPaciente!cveTipoPaciente) & _
                        " or CcComisionMedicoEmpresa.intCveTipoPacienteEmpresa = " & IIf(IsNull(rsTipoPaciente!cveEmpresa), 0, rsTipoPaciente!cveEmpresa) & ")"

                    Set rsComisionesMedico = frsRegresaRs(vlstrSentencia)
                    If rsComisionesMedico.RecordCount <> 0 Then
                        Do While Not rsComisionesMedico.EOF
                            vllngContador = 0
                            vlblnTermina = False

                            Do While vllngContador <= lstComisiones.ListCount - 1 And Not vlblnTermina
                                If lstComisiones.ItemData(vllngContador) = rsComisionesMedico!smiCveComision Then
                                    lstComisiones.ListIndex = vllngContador
                                    cmdSelecciona_Click 0
                                    vlblnTermina = True
                                    chkComisiones.Value = 1
                                End If
                                vllngContador = vllngContador + 1
                            Loop
                            rsComisionesMedico.MoveNext
                        Loop
                    End If
                End If
            Else
                'Asigna comisiones
                pAsignaPredeterminadas
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboMedicos_Click"))
End Sub

Private Sub pAsignaPredeterminadas()
    On Error GoTo NotificaError

    Dim vllngContador As Long
    Dim vllngComisionPredeterminada As Long
    Dim vlstrSentencia As String
    Dim rsComisionesPredeterminadas As New ADODB.Recordset
    Dim vlblnEncontrado As Boolean

    vlstrSentencia = "select * from PvComision where bitAsignada = 1 and bitActivo = 1"
    Set rsComisionesPredeterminadas = frsRegresaRs(vlstrSentencia)

    If rsComisionesPredeterminadas.RecordCount <> 0 Then

        Do While Not rsComisionesPredeterminadas.EOF
            vllngContador = 0
            vlblnEncontrado = False
            Do While vllngContador <= lstComisiones.ListCount - 1 And Not vlblnEncontrado

                If rsComisionesPredeterminadas!smiCveComision = lstComisiones.ItemData(vllngContador) Then
                    vlblnEncontrado = True
                    lstComisiones.ListIndex = vllngContador
                    cmdSelecciona_Click 0
                    chkComisiones.Value = 1
                End If
                vllngContador = vllngContador + 1
            Loop
            rsComisionesPredeterminadas.MoveNext
        Loop

    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAsignaPredeterminadas"))
End Sub

Private Sub cboMedicos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then mskFechaAtencion.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboMedicos_KeyDown"))
End Sub

Private Sub cboTarifa_Click()
    pCalculoTotales
End Sub

Private Sub cboTarifa_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboTarifa_KeyDown"))
End Sub

Private Sub cboUsoCFDI_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtReciboNombre.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboUsoCFDI_KeyPress"))

End Sub

Private Sub chkComisiones_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkComisiones_KeyDown"))
End Sub

Private Sub chkHonorarioFacturado_Click()
   ' If lblnPacienteConvenio Then
   '     If chkHonorarioFacturado.Value = 0 Then
   '         chkPagoDirecto.Enabled = True
   '     Else
   '         chkPagoDirecto.Enabled = False
   '         chkPagoDirecto.Value = 0
   '     End If
   ' End If
End Sub

Private Sub chkHonorarioFacturado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkHonorarioFacturado_KeyDown"))
End Sub

Private Sub chkHonorarioFacturado_Validate(Cancel As Boolean)
 '   If lblnPacienteConvenio Then
 '       If chkHonorarioFacturado.Value = 0 Then
 '           chkPagoDirecto.Enabled = True
 '       Else
 '           chkPagoDirecto.Enabled = False
 '           chkPagoDirecto.Value = 0
 '       End If
 '   End If
End Sub

Private Sub chkPagoDirecto_Click()
    If chkPagoDirecto.Value = vbChecked Then
        txtReciboNombre.Text = lstrNombreConvPers
        txtDireccion.Text = lstrDireccionConvPers
        txtRFC.Text = lstrRFCConvPers
        txtEmail.Text = lstrCorreoElectronico
    Else
        txtReciboNombre.Text = vgstrNombreHospitalCH
        txtDireccion.Text = Trim(vgstrDireccionCH) & ", COLONIA " & Trim(vgstrColoniaCH) & ", CP " & Trim(vgstrCodPostalCH) & ", " & vgstrCiudadCH & ", " & vgstrEstadoCH
        txtRFC.Text = vgstrRfCCH
        txtEmail.Text = vgstrEmailCH
    End If
End Sub

Private Sub chkPagoDirecto_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkPagoDirecto_KeyDown"))
End Sub

Private Sub chkRequiereCFDI_Click()
    If chkRequiereCFDI.Value = vbChecked Then
        cboUsoCFDI.Enabled = True
    Else
        cboUsoCFDI.Enabled = False
        cboUsoCFDI.ListIndex = -1
    End If
End Sub

Private Sub chkRequiereCFDI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkRetencion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    pCalculoTotales
    
    If KeyCode = vbKeyReturn Then SendKeys vbTab
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencion_KeyDown"))
End Sub

Private Function fblnHonorarioValido() As Boolean
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rsTemp As New ADODB.Recordset
    
    fblnHonorarioValido = True
    
    If Val(Format(txtMontoHonorario.Text, "############.00")) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbCritical + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
        txtMontoHonorario.SetFocus
    End If
    
    If fblnHonorarioValido And cboMedicos.ListIndex < 0 Then
        'Seleccione el médico.
        MsgBox SIHOMsg(332), vbCritical, "Mensaje"
        fblnHonorarioValido = False
        cboMedicos.SetFocus
    End If
    
    If fblnHonorarioValido Then
        vlstrSentencia = "Select isnull(intNumeroCuenta,0) Cuenta From HoMedicoEmpresa " & _
                        " Where intCLAveMedico = " & Trim(Str(cboMedicos.ItemData(cboMedicos.ListIndex))) & _
                        " and TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
        Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsTemp.RecordCount <> 0 Then
            If rsTemp!cuenta <> 0 Then
                vllngCuentaMedico = rsTemp!cuenta
            Else
                'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                MsgBox SIHOMsg(519), vbCritical, "Mensaje"
                fblnHonorarioValido = False
            End If
        End If
    End If
    If fblnHonorarioValido And Not IsDate(mskFechaAtencion.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbCritical + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
        mskFechaAtencion.SetFocus
    End If
    If fblnHonorarioValido And Not IsDate(mskFechaAtencionFin.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbCritical + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
        mskFechaAtencionFin.SetFocus
    End If
    If fblnHonorarioValido Then
        If CDate(mskFechaAtencion.Text) > CDate(mskFechaAtencionFin.Text) Then
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbCritical + vbOKOnly, "Mensaje"
            fblnHonorarioValido = False
            mskFechaAtencion.SetFocus
        End If
    End If
    If fblnHonorarioValido Then
        If CDate(mskFechaAtencion.Text) > fdtmServerFecha Then
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbCritical + vbOKOnly, "Mensaje"
            fblnHonorarioValido = False
            mskFechaAtencion.SetFocus
        End If
    End If
    If fblnHonorarioValido Then
        If CDate(mskFechaAtencionFin.Text) > fdtmServerFecha Then
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbCritical + vbOKOnly, "Mensaje"
            fblnHonorarioValido = False
            mskFechaAtencionFin.SetFocus
        End If
    End If
    If fblnHonorarioValido And Trim(txtConcepto.Text) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbCritical + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
        txtConcepto.SetFocus
    End If
    If fblnHonorarioValido And chkRetencion.Value = 1 And cboTarifa.ListIndex = -1 Then
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbExclamation + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
    End If
    If fblnHonorarioValido And Val(Format(txtTotal.Text, "############.00")) < 0 Then
        'Dato incorrecto.
        MsgBox SIHOMsg(406), vbCritical + vbOKOnly, "Mensaje"
        fblnHonorarioValido = False
        txtMontoHonorario.SetFocus
    End If
     If fblnHonorarioValido And chkRetencionRTP.Value = 1 And dblPorcentajeRTP = 0 Then
        'Dato incorrecto.
        'MsgBox SIHOMsg(406), vbCritical + vbOKOnly, "Mensaje"
        '1649 No se encuentra configurado el porcentaje de retención de RTP en los parámetros de cuentas por pagar
        MsgBox SIHOMsg(1649), vbExclamation, "Mensaje"
        fblnHonorarioValido = False
        chkRetencionRTP.SetFocus
    End If
    
    

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnHonorarioValido"))
End Function

Private Sub chkRetencionRTP_Click()
On Error GoTo NotificaError
    Dim rs As New ADODB.Recordset

    If chkRetencionRTP.Value = 1 And dblPorcentajeRTP = 0 Then
        vlstrSentencia = "Select vchvalor from siparametro where vchnombre = 'NUMPORCENTAJERTP'"
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount <> 0 Then
            ''If IsNull
            dblPorcentajeRTP = IIf(IsNull(rs!vchvalor), 0, rs!vchvalor)
        Else
            '1649 No se encuentra configurado el porcentaje de retención de RTP en los parámetros de cuentas por pagar
            'MsgBox SIHOMsg(1649), vbExclamation, "Mensaje"
            dblPorcentajeRTP = 0
        End If
        
        If dblPorcentajeRTP = 0 Then
            '1649 No se encuentra configurado el porcentaje de retención de RTP en los parámetros de cuentas por pagar
            MsgBox SIHOMsg(1649), vbExclamation, "Mensaje"
        End If
    End If
    
    If txtMontoHonorario.Text <> "" Then
        pCalculoTotales
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencion_Click"))
End Sub

Private Sub chkRetencionRTP_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencionRTP_KeyDown"))
End Sub

Private Sub cmdAbrirCarta_Click()
    Dim rs As ADODB.Recordset
    Dim stmCarta As New ADODB.Stream
    Dim strDestino As String
    If grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta) <> "" Then
        If grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCartaMod) = "*" Then
            pAbrirArchivo grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta)
        Else
            'OPEN RECORDSET TO READ BLOB
            Set rs = frsRegresaRs("select * from PVHonorarioDocumentos where intCvehonorario = " & grdHonorarios.TextMatrix(grdHonorarios.Row, 1), adLockOptimistic, adOpenStatic)
            If Not rs.EOF Then
                stmCarta.Type = adTypeBinary
                stmCarta.Open
                stmCarta.Write rs!blbCartaAutorizacion
                strDestino = Environ$("temp") & "\" & rs!vchNombreCartaAutorizacion
                stmCarta.SaveToFile strDestino, adSaveCreateOverWrite
                stmCarta.Close
                pAbrirArchivo strDestino
            End If
            rs.Close
        End If
    End If
    
End Sub

Private Sub cmdActualizarRecibo_Click()
  On Error GoTo NotificaError
    
    Dim vllngPersona As Long
    Dim vllngContador As Long
    Dim vlstrSentencia As String
    Dim rsCarta As ADODB.Recordset
    Dim stmCarta As New ADODB.Stream
    If Not fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then Exit Sub
    
    vllngPersona = flngPersonaGraba(vgintNumeroDepartamento)
    
    If vllngPersona <> 0 Then
        
        For vllngContador = 1 To grdHonorarios.Rows - 1
            If Trim(grdHonorarios.TextMatrix(vllngContador, 0)) = "*" Then
                vlstrSentencia = _
                "update PvHonorario set " & _
                    "chrNumReciboHonorario =' " & grdHonorarios.TextMatrix(vllngContador, 8) & "'," & _
                    "intCveEmpleado = " & Str(vllngPersona) & _
                "Where " & _
                    "intConsecutivo = " & grdHonorarios.TextMatrix(vllngContador, 1)
                pEjecutaSentencia vlstrSentencia
                
                If grdHonorarios.TextMatrix(vllngContador, cintColFilCartaMod) = "*" Then
                    Set rsCarta = frsRegresaRs("select * from PVHonorarioDocumentos where intCveHonorario = " & grdHonorarios.TextMatrix(vllngContador, 1), adLockOptimistic, adOpenStatic)
                    If rsCarta.EOF Then
                        rsCarta.AddNew
                        rsCarta!intCveHonorario = grdHonorarios.TextMatrix(vllngContador, 1)
                    End If
                    If grdHonorarios.TextMatrix(vllngContador, cintColFilCarta) <> "" Then
                        rsCarta!vchNombreCartaAutorizacion = fstrSoloNombreArchivo(grdHonorarios.TextMatrix(vllngContador, cintColFilCarta))
                        stmCarta.Type = adTypeBinary
                        stmCarta.Open
                        stmCarta.LoadFromFile grdHonorarios.TextMatrix(vllngContador, cintColFilCarta)
                        rsCarta!blbCartaAutorizacion = stmCarta.Read
                    Else
                        rsCarta!vchNombreCartaAutorizacion = ""
                        rsCarta!blbCartaAutorizacion = ""
                    End If
                    rsCarta.Update
                End If
            End If
        Next vllngContador
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        cmdCargar_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdActualizarRecibo_Click"))

End Sub

Private Sub cmdAdjuntar_Click()
    CommonDialog1.FileName = txtAdjuntar.Text
    CommonDialog1.ShowOpen
    txtAdjuntar.Text = CommonDialog1.FileName
    SendKeys vbTab
End Sub

Private Sub cmdAgregar_Click()
    On Error GoTo NotificaError
    Dim lngIdTarifa As Long
        
    vlblnContinuaRevision = True

    '--------------------------------------------------------
    ' Validación de Datos
    '--------------------------------------------------------
    If fblnHonorarioValido() Then
    
        fraHonorarios.Enabled = True
        
        pVerificaHonorarioIgual
        
        If vlblnContinuaRevision Then
        
            lngIdTarifa = 0
            If cboTarifa.ListIndex <> -1 Then
                lngIdTarifa = cboTarifa.ItemData(cboTarifa.ListIndex)
            End If
        
            With grdCargaHonorarios
                If Trim(.TextMatrix(1, 1)) = "" Then
                    .Row = 1
                Else
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
                .TextMatrix(.Row, 0) = ""
                .TextMatrix(.Row, cintColNombreMedico) = cboMedicos.List(cboMedicos.ListIndex)
                .TextMatrix(.Row, cintColCveMedico) = cboMedicos.ItemData(cboMedicos.ListIndex)
                .TextMatrix(.Row, cintColFechaAtencionInicio) = CDate(mskFechaAtencion.Text)
                .TextMatrix(.Row, cintColConcepto) = Trim(txtConcepto.Text)
                .TextMatrix(.Row, cintColMonto) = FormatCurrency(txtMontoHonorario.Text)
                .TextMatrix(.Row, cintColRetencion) = FormatCurrency(dblTotalRetencionDetalle)
                .TextMatrix(.Row, cintColComision) = txtCantidadComision.Text
                .TextMatrix(.Row, cintColIVAComision) = txtIVAComision.Text
                .TextMatrix(.Row, cintColReciboNombre) = txtReciboNombre.Text
                .TextMatrix(.Row, cintColDireccion) = txtDireccion.Text
                .TextMatrix(.Row, cIntColRFC) = Trim(Replace(Replace(Replace(txtRFC.Text, "-", ""), "_", ""), " ", ""))
                .TextMatrix(.Row, cintColReciboHonorario) = txtReciboHonorario.Text
                .TextMatrix(.Row, cintColFechaAtencionFin) = CDate(mskFechaAtencionFin.Text)
                .TextMatrix(.Row, cintColTotalPagar) = txtTotal.Text
                .TextMatrix(.Row, cintColCuentaContableMedico) = vllngCuentaMedico
                .TextMatrix(.Row, cintColFechaAtencionInicio) = CDate(mskFechaAtencion.Text)
                .TextMatrix(.Row, cintColIdTarifa) = lngIdTarifa
                .TextMatrix(.Row, cintColPagoDirecto) = IIf(chkPagoDirecto.Value = vbChecked, "*", "")
                .TextMatrix(.Row, cintColCarta) = fstrSoloNombreArchivo(txtAdjuntar.Text)
                .TextMatrix(.Row, cintColCartaCompleta) = txtAdjuntar.Text
                .TextMatrix(.Row, cintColObservaciones) = txtObservaciones.Text
                .TextMatrix(.Row, cintColRetencionRTP) = txtRetencionRTP.Text
                .TextMatrix(.Row, cintColRetencionISR) = txtRetencion.Text
                
                llngNumMedico = cboMedicos.ItemData(cboMedicos.ListIndex)
                .Redraw = True
            End With
        End If

        For vlintContador = 0 To UBound(vlComisiones) - 1
            With grdComisiones
                .FixedCols = 0
                .Cols = 4
                .FormatString = "Honorario|cvecom|valor|iva"
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = grdCargaHonorarios.Row
                .TextMatrix(.Row, 1) = vlComisiones(vlintContador).vllngCveComision
                .TextMatrix(.Row, 2) = vlComisiones(vlintContador).vldblCantidad
                .TextMatrix(.Row, 3) = vlComisiones(vlintContador).vldblIVA
                .Visible = True
            End With
        Next
    
        'Limpia los datos del honorario
        txtReciboHonorario.Text = ""
        txtConcepto.Text = ""
        txtMontoHonorario.Text = ""
        txtRetencion.Text = ""
        txtRetencionRTP.Text = ""
        txtNetoPagar.Text = ""
        txtCantidadComision.Text = ""
        txtIVAComision.Text = ""
        txtPagoCredito.Text = ""
        txtTotal.Text = ""
        txtAdjuntar.Text = ""
        txtObservaciones.Text = ""
'        chkPagoDirecto.Value = vbUnchecked
        cboMedicos.SetFocus
        pCalculaTotales
        
        'chkHonorarioFacturado.Enabled = False
        
        If chkHonorarioFacturado.Value = 1 Then
            fraRecibo.Enabled = False
            txtReciboHonorario.Enabled = False
        Else
            fraRecibo.Enabled = True
            txtReciboHonorario.Enabled = True
        End If
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgregar_Click"))
End Sub

Private Function fblnConsulta() As Boolean

    fblnConsulta = True
    If Trim(mskInicio.ClipText) <> "" Then
        If Not IsDate(mskInicio.Text) Then
            fblnConsulta = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskInicio.SetFocus
        End If
    End If
    If fblnConsulta And Trim(mskFin.ClipText) <> "" Then
        If Not IsDate(mskFin.Text) Then
            fblnConsulta = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskFin.SetFocus
        End If
    End If
    If fblnConsulta And Trim(mskInicio.ClipText) <> "" And Trim(mskFin.ClipText) <> "" Then
        If CDate(mskFin.Text) < CDate(mskInicio.Text) Then
            fblnConsulta = False
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
            mskInicio.SetFocus
        End If
    End If

End Function

Private Sub pLimpiaHonorarios()

    llngTotalSel = 0
    
    cmdMarcar.Enabled = False
    cmdInvertirSel.Enabled = False

    cmdPagoAnticipado.Enabled = False
    cmdDelete.Enabled = False
    cmdActualizarRecibo.Enabled = False

    With grdHonorarios
        .Clear
        .Cols = 24
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .FormatString = cstrColumnas
    
        .ColWidth(0) = 200
        .ColWidth(cintColFilConsecutivo) = 0
        .ColWidth(cintColFilCuenta) = IIf(txtFilPaciente.Text <> "" And txtFilCuenta.Text <> "", 0, 1000)
        .ColWidth(cintColFilTipo) = IIf(txtFilPaciente.Text <> "" And txtFilCuenta.Text <> "", 0, 800)
        .ColWidth(cintColFilPaciente) = IIf(txtFilPaciente.Text <> "" And txtFilCuenta.Text <> "", 0, 3000)
        .ColWidth(cintColFilMedico) = IIf(cboFilMedico.ItemData(cboFilMedico.ListIndex) = 0, 3000, 0)
        .ColWidth(cintColFilInicioAtencion) = 1500
        .ColWidth(cintColFilFinAtencion) = 1500
        .ColWidth(cintColFilRecibo) = 1000
        .ColWidth(cintColFilConcepto) = 4000
        .ColWidth(cintColFilMonto) = 1000
        .ColWidth(cintColFilRetencion) = 1200
        .ColWidth(cintColFilSubtotal) = 1200
        .ColWidth(cintColFilComision) = 1200
        .ColWidth(cintColFilIVA) = 1200
        .ColWidth(cintColFilTotal) = 1200
        .ColWidth(cintColFilEstado) = 3500
        .ColWidth(cintColFilEmpleado) = 3000
        .ColWidth(cintColFilFechaEnvio) = 1500
        .ColWidth(cintColFilCliente) = 3000
        .ColWidth(cintColFilEstatus) = 0
        .ColWidth(cintColFilCveMedico) = 0
        .ColWidth(cintColFilCxP) = 0
        .ColWidth(cintColFilCarta) = 1000
        .ColWidth(cintColFilCartaMod) = 0
        
        .FixedAlignment(cintColFilConsecutivo) = flexAlignCenterCenter
        .FixedAlignment(cintColFilCuenta) = flexAlignCenterCenter
        .FixedAlignment(cintColFilTipo) = flexAlignCenterCenter
        .FixedAlignment(cintColFilPaciente) = flexAlignCenterCenter
        .FixedAlignment(cintColFilMedico) = flexAlignCenterCenter
        .FixedAlignment(cintColFilInicioAtencion) = flexAlignCenterCenter
        .FixedAlignment(cintColFilFinAtencion) = flexAlignCenterCenter
        .FixedAlignment(cintColFilRecibo) = flexAlignCenterCenter
        .FixedAlignment(cintColFilConcepto) = flexAlignCenterCenter
        .FixedAlignment(cintColFilMonto) = flexAlignCenterCenter
        .FixedAlignment(cintColFilRetencion) = flexAlignCenterCenter
        .FixedAlignment(cintColFilSubtotal) = flexAlignCenterCenter
        .FixedAlignment(cintColFilComision) = flexAlignCenterCenter
        .FixedAlignment(cintColFilIVA) = flexAlignCenterCenter
        .FixedAlignment(cintColFilTotal) = flexAlignCenterCenter
        .FixedAlignment(cintColFilEstado) = flexAlignCenterCenter
        .FixedAlignment(cintColFilEmpleado) = flexAlignCenterCenter
        .FixedAlignment(cintColFilFechaEnvio) = flexAlignCenterCenter
        .FixedAlignment(cintColFilCliente) = flexAlignCenterCenter
        .FixedAlignment(cintColFilCarta) = flexAlignCenterCenter
        
        .ColAlignment(cintColFilConsecutivo) = flexAlignCenterCenter
        .ColAlignment(cintColFilCuenta) = flexAlignRightCenter
        .ColAlignment(cintColFilTipo) = flexAlignLeftCenter
        .ColAlignment(cintColFilPaciente) = flexAlignLeftCenter
        .ColAlignment(cintColFilMedico) = flexAlignLeftCenter
        .ColAlignment(cintColFilInicioAtencion) = flexAlignCenterCenter
        .ColAlignment(cintColFilFinAtencion) = flexAlignCenterCenter
        .ColAlignment(cintColFilRecibo) = flexAlignLeftCenter
        .ColAlignment(cintColFilConcepto) = flexAlignLeftCenter
        .ColAlignment(cintColFilMonto) = flexAlignRightCenter
        .ColAlignment(cintColFilRetencion) = flexAlignRightCenter
        .ColAlignment(cintColFilSubtotal) = flexAlignRightCenter
        .ColAlignment(cintColFilComision) = flexAlignRightCenter
        .ColAlignment(cintColFilIVA) = flexAlignRightCenter
        .ColAlignment(cintColFilTotal) = flexAlignRightCenter
        .ColAlignment(cintColFilEstado) = flexAlignLeftCenter
        .ColAlignment(cintColFilEmpleado) = flexAlignLeftCenter
        .ColAlignment(cintColFilFechaEnvio) = flexAlignLeftCenter
        .ColAlignment(cintColFilCliente) = flexAlignLeftCenter
        .ColAlignment(cintColFilCarta) = flexAlignLeftCenter
    End With

End Sub

Private Sub cmdCargar_Click()
    Dim rs As New ADODB.Recordset
    Dim strFiltroFecha As String
    Dim strFechaInicio As String
    Dim strFechaFin As String
    Dim Y As Integer
    
    If Not fblnConsulta Then Exit Sub
    
    pLimpiaHonorarios
    
    strFiltroFecha = IIf(IsDate(mskInicio.Text), "1", "0")
    
    If IsDate(mskInicio.Text) And IsDate(mskFin.Text) Then
        strFechaInicio = fstrFechaSQL(mskInicio.Text)
        strFechaFin = fstrFechaSQL(mskFin.Text)
    Else
        strFechaInicio = fstrFechaSQL(fdtmServerFecha)
        strFechaFin = fstrFechaSQL(fdtmServerFecha)
    End If
    
    vgstrParametrosSP = _
    Trim(Str(vgintNumeroDepartamento)) & _
    "|" & Trim(Str(Val(txtFilCuenta.Text))) & _
    "|" & IIf(optFilTipoPaciente(0).Value, "I", IIf(optFilTipoPaciente(1).Value, "E", "*")) & _
    "|" & Trim(Str(cboFilMedico.ItemData(cboFilMedico.ListIndex))) & _
    "|" & strFiltroFecha & _
    "|" & strFechaInicio & _
    "|" & strFechaFin
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelHonorario")
    
    If rs.RecordCount <> 0 Then
        cmdMarcar.Enabled = True
        cmdInvertirSel.Enabled = True
       
        With grdHonorarios
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, cintColFilConsecutivo) = rs!intConsecutivo
                .TextMatrix(.Rows - 1, cintColFilCuenta) = rs!cuenta
                .TextMatrix(.Rows - 1, cintColFilTipo) = rs!tipo
                .TextMatrix(.Rows - 1, cintColFilPaciente) = IIf(IsNull(rs!Paciente), " ", rs!Paciente)
                .TextMatrix(.Rows - 1, cintColFilMedico) = IIf(IsNull(rs!Medico), " ", rs!Medico)
                .TextMatrix(.Rows - 1, cintColFilInicioAtencion) = Format(rs!InicioAtencion, "dd/mmm/yyyy")
                .TextMatrix(.Rows - 1, cintColFilFinAtencion) = Format(rs!FinAtencion, "dd/mmm/yyyy")
                .TextMatrix(.Rows - 1, cintColFilRecibo) = IIf(IsNull(rs!Recibo), " ", rs!Recibo)
                .TextMatrix(.Rows - 1, cintColFilConcepto) = IIf(IsNull(rs!Concepto), " ", rs!Concepto)
                .TextMatrix(.Rows - 1, cintColFilMonto) = FormatCurrency(rs!Monto, 2)
                .TextMatrix(.Rows - 1, cintColFilRetencion) = FormatCurrency(rs!Retencion, 2)
                .TextMatrix(.Rows - 1, cintColFilSubtotal) = FormatCurrency(rs!Subtotal, 2)
                .TextMatrix(.Rows - 1, cintColFilComision) = FormatCurrency(rs!Comision, 2)
                .TextMatrix(.Rows - 1, cintColFilIVA) = FormatCurrency(rs!IvaComision, 2)
                .TextMatrix(.Rows - 1, cintColFilTotal) = FormatCurrency(rs!Total, 2)
                .TextMatrix(.Rows - 1, cintColFilEstado) = IIf(IsNull(rs!Estado), " ", rs!Estado)
                .TextMatrix(.Rows - 1, cintColFilEmpleado) = IIf(IsNull(rs!Persona), " ", rs!Persona)
                If IsDate(rs!fechaEnvio) Then
                    .TextMatrix(.Rows - 1, cintColFilFechaEnvio) = Format(rs!fechaEnvio, "dd/mmm/yyyy")
                Else
                    .TextMatrix(.Rows - 1, cintColFilFechaEnvio) = " "
                End If
                .TextMatrix(.Rows - 1, cintColFilCliente) = IIf(IsNull(rs!NombreCliente), " ", rs!NombreCliente)
                .TextMatrix(.Rows - 1, cintColFilEstatus) = IIf(IsNull(rs!chrestatus), " ", rs!chrestatus)
                .TextMatrix(.Rows - 1, cintColFilCveMedico) = rs!CveMedico
                .TextMatrix(.Rows - 1, cintColFilCxP) = rs!CxP
                .TextMatrix(.Rows - 1, cintColFilCarta) = IIf(IsNull(rs!vchNombreCartaAutorizacion), "", rs!vchNombreCartaAutorizacion)
                If .TextMatrix(.Rows - 1, cintColFilEstatus) = "C" Then
                    pDarColor grdHonorarios.Rows - 1
                End If
                .Rows = .Rows + 1
                rs.MoveNext
            
            Loop
            .Rows = .Rows - 1
        End With
        
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbInformation + vbOKOnly, "Mensaje"
    End If

End Sub

Private Sub pDarColor(lngRenglon As Long)
    On Error GoTo NotificaError
    
    Dim lngColumna As Long
    
    grdHonorarios.Row = lngRenglon
    For lngColumna = 1 To grdHonorarios.Cols - 1
        grdHonorarios.Col = lngColumna
        grdHonorarios.CellForeColor = &HFF&
    Next lngColumna

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pDarColor"))
End Sub

Private Sub cmdCargarCarta_Click()
    CommonDialog1.Flags = &H1000
    CommonDialog1.FileName = grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta)
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta) Then
        grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta) = CommonDialog1.FileName
        grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCartaMod) = "*"
        grdHonorarios.TextMatrix(grdHonorarios.Row, 0) = "*"
                 cmdActualizarRecibo.Enabled = True
         cmdPagoAnticipado.Enabled = True
         cmdDelete.Enabled = True

    End If
End Sub

Private Sub cmdInvertirSel_Click()
    If grdHonorarios.Rows > 0 Then
        grdHonorarios.Col = 0
        '
        'If OptAccion(0).Value Then
        If Val(grdHonorarios.TextMatrix(1, cintColFilConsecutivo)) > 0 Then pMarca "*", -1
        'End If
        
    End If
End Sub

Private Sub cmdLimpiarCarta_Click()
    If grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta) <> "" Then
        grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCarta) = ""
        grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilCartaMod) = "*"
        grdHonorarios.TextMatrix(grdHonorarios.Row, 0) = "*"
                 cmdActualizarRecibo.Enabled = True
         cmdPagoAnticipado.Enabled = True
         cmdDelete.Enabled = True

    End If
End Sub

Private Sub cmdMarcar_Click()
    If grdHonorarios.Row > 0 Then
        grdHonorarios.Col = 0
        grdHonorarios_DblClick
    End If
End Sub

Private Function flngCuentaAcreedora(lngCveMedico As Long) As Long
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    flngCuentaAcreedora = 0
    
    vlstrSentencia = "select intNumeroCuenta intnumcuentacontable from HoMedicoempresa where intClaveMedico =" & lngCveMedico & _
                    " and tnyclaveempresa = " & vgintClaveEmpresaContable
    
    Set rs = frsRegresaRs(vlstrSentencia)
    If rs.RecordCount <> 0 Then
        flngCuentaAcreedora = IIf(IsNull(rs!INTNUMCUENTACONTABLE), 0, rs!INTNUMCUENTACONTABLE)
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngCuentaAcreedora"))
End Function

Private Function flngAbreCuenta(strNumPaciente As String, lngEmpleado As Long) As Long

    Dim rsregistroexterno As New ADODB.Recordset
    Dim strSentencia As String
    
    flngAbreCuenta = 1
    vgstrParametrosSP = strNumPaciente & "|10|" & _
                        lngEmpleado & "|" & _
                        vgintNumeroDepartamento & "|0|0|0|" & _
                        lngEmpleado & "|" & _
                        lngEmpleado & "|" & _
                        vgintTipoPaciente & "|" & _
                        vgintEmpresa & _
                        "|0|0|0|0|0|0|" & _
                        llngNumMedico & _
                        "|0|0|0|0|0|0|" & _
                        fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & _
                        fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & _
                        "|O||||||||||||||0|1|0|0|0|0|"
    
    frsEjecuta_SP vgstrParametrosSP, "sp_ExInSPacienteIngreso", True, flngAbreCuenta
    
End Function

Private Sub cmdPagoAnticipado_Click()
    On Error GoTo NotificaError
    
    Dim rsCpCuentaPagarMedico As New ADODB.Recordset
    Dim rsPagoAnt As New ADODB.Recordset
    Dim vldblcantidadregistrar As Double
    Dim vllngCtaAcreedoraMedico As Long
    Dim vllngNumCuentaPagar As Long
    Dim lngNumPolizaDetalle As Long
    Dim vllngPersonaGraba As Long
    Dim vllngResultado As Long
    Dim vlstrSentencia As String
    Dim vllngNumPoliza As Long
    Dim contador As Long
    Dim vlintMoneda As Integer
    Dim ldblTotal As Double
    Dim lblnPago As Boolean
    Dim lngTotal As Long
    
    If grdHonorarios.Rows > 0 Then
    
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)

        If vllngPersonaGraba <> 0 Then
            
            lblnPago = False
            lngTotal = 0
    
            For contador = 1 To grdHonorarios.Rows - 1
            
                If grdHonorarios.TextMatrix(contador, 0) = "*" And grdHonorarios.TextMatrix(contador, cintColFilEstatus) = "R" And Val(grdHonorarios.TextMatrix(contador, cintColFilCxP)) = 0 Then
    
                    vllngCtaAcreedoraMedico = flngCuentaAcreedora(CLng(grdHonorarios.TextMatrix(contador, cintColFilCveMedico)))
                    
                    If vllngCtaAcreedoraMedico <> 0 Then
                
                        EntornoSIHO.ConeccionSIHO.BeginTrans
                        
                        'Validar que no se esté realizando algún cierre
                        vllngResultado = 1
                        vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
                        frsEjecuta_SP vgstrParametrosSP, "sp_CnUpdEstatusCierre", True, vllngResultado
        
                        If vllngResultado = 1 Then
                        
                            If Not fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(fdtmServerFecha), Month(fdtmServerFecha)) Then
    
                                Set rsPagoAnt = frsEjecuta_SP(grdHonorarios.TextMatrix(contador, 1), "Sp_Pvselhonorario")
                                If rsPagoAnt.RecordCount > 0 Then
                            
                                    vlstrSentencia = "select * from CpCuentaPagarMedico where intNumCuentaPagar = -1"
                                    Set rsCpCuentaPagarMedico = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                                    
                                    'Cuando es un paciente tipo convenio a credito, la empresa absorbe la retención del ISR
                                    If (rsPagoAnt!convenio <> 0 And rsPagoAnt!convenio <> " " And Not IsNull(rsPagoAnt!convenio) And rsPagoAnt!EstadoCredito = 1) Then
                                        vldblcantidadregistrar = IIf(IsNull(rsPagoAnt!MontoHonorario), 0, rsPagoAnt!MontoHonorario) - IIf(IsNull(rsPagoAnt!MontoRetencion), 0, rsPagoAnt!MontoRetencion)
                                    Else
                                        vldblcantidadregistrar = IIf(IsNull(rsPagoAnt!MontoHonorario), 0, rsPagoAnt!MontoHonorario)
                                    End If
                                    vldblcantidadregistrar = vldblcantidadregistrar - IIf(IsNull(rsPagoAnt!MontoRetencionRTP), 0, rsPagoAnt!MontoRetencionRTP)
                    
                                    vlintMoneda = rsPagoAnt!EstadoMoneda
                                    vldblTipoCambio = IIf(IsNull(rsPagoAnt!TipoCambio), 0, rsPagoAnt!TipoCambio)
                                    ldblTotal = IIf(IsNull(rsPagoAnt!MontoHonorario), 0, rsPagoAnt!MontoHonorario) - IIf(IsNull(rsPagoAnt!MontoRetencion), 0, rsPagoAnt!MontoRetencion) - IIf(IsNull(rsPagoAnt!MontoRetencionRTP), 0, rsPagoAnt!MontoRetencionRTP)
                                    ldblTotal = ldblTotal - IIf(IsNull(rsPagoAnt!Comision), 0, rsPagoAnt!Comision) - IIf(IsNull(rsPagoAnt!IvaComision), 0, rsPagoAnt!IvaComision)
                                
                                    With rsCpCuentaPagarMedico
                                        .AddNew
                                        !intCveMedico = rsPagoAnt!IdMedico
                                        !intNumCuentaMedico = vllngCtaAcreedoraMedico
                                        !dtmfecha = fdtmServerFecha + fdtmServerHora
                                        !MNYCantidad = ldblTotal
                                        !bitestatusmoneda = vlintMoneda
                                        !MNYTIPOCAMBIO = vldblTipoCambio
                                        !intConsecutivo = CLng(grdHonorarios.TextMatrix(contador, 1))
                                        !mnymontohonorario = IIf(IsNull(rsPagoAnt!MontoHonorario), 0, rsPagoAnt!MontoHonorario)
                                        !mnyretencion = IIf(IsNull(rsPagoAnt!MontoRetencion), 0, rsPagoAnt!MontoRetencion)
                                        !mnyretencionRTP = IIf(IsNull(rsPagoAnt!MontoRetencionRTP), 0, rsPagoAnt!MontoRetencionRTP)
                                        !mnycomision = IIf(IsNull(rsPagoAnt!Comision), 0, rsPagoAnt!Comision)
                                        !mnyivacomision = IIf(IsNull(rsPagoAnt!IvaComision), 0, rsPagoAnt!IvaComision)
                                        .Update
                                    End With
        
                                    vllngNumCuentaPagar = flngObtieneIdentity("SEC_CPCUENTAPAGARMEDICO", rsCpCuentaPagarMedico!intnumcuentapagar)
                
                                    vllngNumPoliza = flngInsertarPoliza(fdtmServerFecha, "D", "PAGO ANTICIPADO A " & grdHonorarios.TextMatrix(contador, cintColFilMedico), vllngPersonaGraba)
                
                                    lngNumPolizaDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, vllngCtaHonorariosPagar, CDbl(Format(vldblcantidadregistrar, "############.00")) * IIf(vlintMoneda = 1, 1, vldblTipoCambio), 1)
                                    lngNumPolizaDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, vllngCtaAcreedoraMedico, CDbl(Format(vldblcantidadregistrar, "############.00")) * IIf(vlintMoneda = 1, 1, vldblTipoCambio), 0)
                                      
                                    vlstrSentencia = "insert into PvHonorarioPoliza (intCveHonorario, intNumeroPoliza) values(" & grdHonorarios.TextMatrix(contador, 1) & "," & vllngNumPoliza & ")"
                                    pEjecutaSentencia vlstrSentencia
                                      
                                    pEjecutaSentencia "update PvHonorario set PvHonorario.chrEstatus = 'M' where PvHonorario.intConsecutivo = " & grdHonorarios.TextMatrix(contador, 1)
        
                                    pEjecutaSentencia "update CnEstatusCierre set vchEstatus='Libre' where tnyClaveEmpresa=" + Str(vgintClaveEmpresaContable)
            
                                    EntornoSIHO.ConeccionSIHO.CommitTrans
                
                                    lblnPago = True
                
                                End If
                            Else
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                frmMensajePeriodoContableCerrado.Show vbModal
                            End If
                        Else
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            'En este momento se está realizando un cierre contable, espere un momento e intente de nuevo.
                            MsgBox SIHOMsg(714), vbOKOnly + vbExclamation, "Mensaje"
                        End If
                    Else
                        'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                        MsgBox SIHOMsg(519) & " " & grdHonorarios.TextMatrix(contador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        lngTotal = lngTotal + 1
                    End If
                Else
                    If grdHonorarios.TextMatrix(contador, cintColFilEstatus) <> "R" And grdHonorarios.TextMatrix(contador, 0) = "*" Then
                        'El honorario seleccionado no está registrado a crédito.
                        MsgBox SIHOMsg(881) & " " & grdHonorarios.TextMatrix(contador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        lngTotal = lngTotal + 1
                    ElseIf Val(grdHonorarios.TextMatrix(contador, cintColFilCxP)) <> 0 And grdHonorarios.TextMatrix(contador, 0) = "*" Then
                        'El honorario seleccionado ya se indicó para pago anticipado.
                        MsgBox SIHOMsg(882) & " " & grdHonorarios.TextMatrix(contador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        lngTotal = lngTotal + 1
                    End If
                End If
            Next contador
            If lblnPago Then
                cmdCargar_Click
                'La operación se realizó satisfactoriamente
                MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
            End If
            
            If lngTotal <> 0 Then
                'En algunos registros no pudo realizarse la operación.
                MsgBox SIHOMsg(880), vbOKOnly + vbInformation, "Mensaje"
            End If
            
            pLimpiaHonorarios
            cmdCargar.SetFocus

        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPagoAnticipado_Click"))
End Sub

Private Sub cmdPagoEfectivo_Click()
    On Error GoTo NotificaError
        Dim vlstrsql As String
        Dim rsPvDetalleCorte As ADODB.Recordset
        Dim vllngNumDetalleCorte As Long
        Dim vllngPersonaGraba As Long
    
        If fDatosValidosEfectivo Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba = 0 Then Exit Sub
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            Set rsPvDetalleCorte = frsRegresaRs("select * from PVDetalleCorte where intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
            
            vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
            vllngCorteGrabando = 1
            vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
            frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
            If vllngCorteGrabando <> 2 Then
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                Exit Sub
            End If
                
            With rsPvDetalleCorte
                .AddNew
                !intnumcorte = vllngNumeroCorte
                !dtmFechahora = fdtmServerFecha + fdtmServerHora
                !chrFolioDocumento = Trim(Str(vllngNumHonorario))
                !chrTipoDocumento = "HO"
                !intFormaPago = aFormasPago(0).vlintNumFormaPago
                !mnyCantidadPagada = aFormasPago(0).vldblCantidad * -1
                !MNYTIPOCAMBIO = aFormasPago(0).vldblTipoCambio
                !intfoliocheque = IIf(Trim(aFormasPago(0).vlstrFolio) = "", "0", Trim(aFormasPago(0).vlstrFolio))
                !intNumCorteDocumento = vllngNumCorte
                !INTCVEEMPLEADO = vllngPersonaGraba
                .Update
            End With
                                        
            vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
            
            If Trim(aFormasPago(0).vlstrRFC) <> "" And Trim(aFormasPago(0).vlstrBancoSAT) <> "" Then
                frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(0).vlstrRFC) & "'|'" & Trim(aFormasPago(0).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(0).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(0).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(0).vldtmFecha))) & "'|'" & Trim(aFormasPago(0).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
            End If
                                           
            pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", aFormasPago(0).vllngCuentaContable, aFormasPago(0).vldblCantidad, False
            pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", vllngCuentaMedico, aFormasPago(0).vldblCantidad, True
            
            vlstrsql = "UPDATE PVHONORARIO SET CHRESTATUS = 'G', VCHESTATUSPORTAL= 'PA' WHERE INTCONSECUTIVO = " & vllngNumHonorario
            pEjecutaSentencia vlstrsql
            
            pLiberaCorte vllngNumeroCorte
            
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "PAGO EFECTIVO DE HONORARIOS", Str(vllngNumHonorario))
            EntornoSIHO.ConeccionSIHO.CommitTrans
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            pMuestraHonorario
        End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPagoEfectivo_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError

    Dim rsPvHonorario As New ADODB.Recordset
    Dim rsPvComisionHonorario As New ADODB.Recordset
    Dim rsPvDetalleCorte As New ADODB.Recordset
    Dim rsPvHonorarioPersonas As New ADODB.Recordset
    Dim vllngPersona As Long
    Dim vlintContador As Integer
    Dim vlintContadorComisiones As Integer
    Dim vlintContadorFormasPago As Integer
    Dim vlstrSentencia As String
    Dim vldblProporcion As Double
    Dim vldblCantidadCorte As Double
    Dim lngNumMovimiento As Long
    Dim vldblCantidad As Double
    Dim strEstatus As String
    Dim dtmfecha As Date
    Dim dtmHora As Date
    Dim blnPagoDirecto As Boolean
    Dim lngNumCliente As Long
    Dim vlstrDocAfectable As String
    Dim vlstrDocSinCuenta As String
    Dim vllngNumDetalleCorte As Long
    Dim vldblAcumuladoCXP As Double
    Dim rsValor As New ADODB.Recordset
    Dim vlblnUsoCorte As Boolean
    Dim rsCnPoliza As New ADODB.Recordset
    Dim rsCnDetallePoliza As New ADODB.Recordset
    Dim vllngPolizaMaestro As Long
    Dim stmCarta As New ADODB.Stream 'Agregado Carta autorizacion
    Dim rsCarta As ADODB.Recordset
    Dim rsEmail As ADODB.Recordset
    Dim rsCorreo As ADODB.Recordset
    Dim clsEmail As clsCDOmail
    Dim vlPassDecoded As String
    Dim strMetodoPago As String
    Dim strMensaje As String
    Dim lngCveFormaPago As Long
    Dim strSQL As String
    Dim vlIntBitPortal  As Integer
    
    vlstrDocAfectable = ""
    vlstrDocSinCuenta = ""
    
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
        
        ' SP para obtener el número de cuenta de los honorarios por pagar
        Set rsPvHonorarioPersonas = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_CLAVE_HONORARIOS")
         
        If vlblnConsulta Then
            'CONSULTA: Solo se actualiza el folio del recibo del honorario del médico
            
            vllngPersona = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersona = 0 Then Exit Sub
            
            vlstrSentencia = "update PvHonorario set chrNumReciboHonorario = '" & txtReciboHonorario.Text & "' where intConsecutivo =" & vllngNumHonorario
            pEjecutaSentencia vlstrSentencia
            
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
            pHabilita 0, 0, 0, 0, 0, 0, 1, 0
            cmdPrint.SetFocus
        Else
            'NUEVO
            If fblnDatosValidos() Then
                vllngPersona = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersona = 0 Then Exit Sub
                
                'INICIA LA TRANSACCION
                EntornoSIHO.ConeccionSIHO.BeginTrans
                dtmfecha = fdtmServerFecha
                dtmHora = fdtmServerHora
                'Si es un paciente de consulta externa, se abre una cuenta cerrada para los honorarios
                If optTipoPaciente(2).Value Then
                    txtMovimientoPaciente.Text = flngAbreCuenta(txtMovimientoPaciente.Text, vllngPersona)
                End If
                
                vlblnUsoCorte = False
                
                'SI ESTÁ CONFIGURADO QUE ENTRA AL CORTE O SE SELECCIONÓ LA FORMA DE PAGO A CRÉDITO
                If chkHonorarioFacturado.Value = 0 Then
                    If vlblnEntraCorte Or aFormasPago(0).vlbolEsCredito Then
                        vlblnUsoCorte = True
                        
                        vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                        vllngCorteGrabando = 1
                        vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                        If vllngCorteGrabando <> 2 Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                    End If
                End If
                
                Set rsPvHonorario = frsRegresaRs("select * from PvHonorario where intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
                Set rsPvComisionHonorario = frsRegresaRs("select * from PvComisionHonorario where intCveHonorario = -1", adLockOptimistic, adOpenDynamic)
                Set rsPvDetalleCorte = frsRegresaRs("select * from PVDetalleCorte where intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
                
                If chkHonorarioFacturado.Value = 0 Then
                    If aFormasPago(0).vlbolEsCredito Then
                        strEstatus = "R" 'CREDITO
                    ElseIf Not vlblnEntraCorte Then
                        strEstatus = "G" 'PAGADO AL MEDICO
                    Else
                        strEstatus = "M" 'PAGO PENDIENTE AL MEDICO
                    End If
                Else
                    strEstatus = "M" 'PAGO PENDIENTE AL MEDICO
                End If
                
                For vlintContador = 1 To grdCargaHonorarios.Rows - 1
                    If chkHonorarioFacturado.Value = 0 Then
                        If Not aFormasPago(0).vlbolEsCredito Then
                            'Validar cuentas contables
                            If grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico) <> 0 Then
                                If Not fblnCuentaAfectable(fstrCuentaContable(grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vgintClaveEmpresaContable) Then
                                    'La cuenta contable asignada al médico no acepta movimientos.
                                    'MsgBox SIHOMsg(1236) & Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3), vbExclamation, "Mensaje"
                                    vlstrDocAfectable = Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3)
                                End If
                            Else
                                'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                                'MsgBox SIHOMsg(519) & Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3), vbOKOnly + vbInformation, "Mensaje"
                                vlstrDocSinCuenta = Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3)
                            End If
                        End If
                    Else
                        'Validar cuentas contables
                        If grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico) <> 0 Then
                            If Not fblnCuentaAfectable(fstrCuentaContable(grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vgintClaveEmpresaContable) Then
                                'La cuenta contable asignada al médico no acepta movimientos.
                                'MsgBox SIHOMsg(1236) & Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3), vbExclamation, "Mensaje"
                                vlstrDocAfectable = Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3)
                            End If
                        Else
                            'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                            'MsgBox SIHOMsg(519) & Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3), vbOKOnly + vbInformation, "Mensaje"
                            vlstrDocSinCuenta = Chr(13) & grdCargaHonorarios.TextMatrix(vlintContador, 3)
                        End If
                    End If
                Next vlintContador
                
                If vlstrDocSinCuenta <> "" Then
                    'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                    MsgBox SIHOMsg(519) & vlstrDocSinCuenta, vbExclamation, "Mensaje"
                End If
                If vlstrDocAfectable <> "" Then
                    'La cuenta contable asignada al médico no acepta movimientos.
                    MsgBox SIHOMsg(1236) & vlstrDocAfectable, vbExclamation, "Mensaje"
                End If
                If vlstrDocSinCuenta <> "" Or vlstrDocAfectable <> "" Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    Exit Sub
                End If
                
                vldblAcumuladoCXP = 0
                
                For vlintContador = 1 To grdCargaHonorarios.Rows - 1
                    vldblAcumuladoCXP = 0
                    With rsPvHonorario
                        
                        If chkHonorarioFacturado.Value = 0 Then
                            If aFormasPago(0).vlbolEsCredito And grdCargaHonorarios.TextMatrix(vlintContador, cintColPagoDirecto) = "*" Then
                                blnPagoDirecto = True
                                lngNumCliente = rsDatosCliente!intNumCliente
                            Else
                                blnPagoDirecto = False
                                lngNumCliente = -1
                            End If
                        Else
                            blnPagoDirecto = False
                            lngNumCliente = -1
                        End If
                        
                        .AddNew
                        !intNumCuenta = Val(txtMovimientoPaciente.Text)
                        !CHRTIPOPACIENTE = IIf(optTipoPaciente(0).Value, "I", "E")
                        !dtmFechaAtencion = CDate(grdCargaHonorarios.TextMatrix(vlintContador, cintColFechaAtencionInicio))
                        !intCveMedico = grdCargaHonorarios.TextMatrix(vlintContador, cintColCveMedico)
                        !chrConcepto = Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColConcepto))
                        !chrNombreRecibo = Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColReciboNombre))
                        !chrDireccion = IIf(Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColDireccion)) = "", " ", Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColDireccion)))
                        !chrRFC = Trim(Replace(Replace(Replace(Trim(grdCargaHonorarios.TextMatrix(vlintContador, cIntColRFC)), "-", ""), "_", ""), " ", ""))
                        !mnymontohonorario = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00"))
                        !mnyretencion = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionISR), "############.00"))
                        !mnycomision = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColComision), "############.00"))
                        !mnyivacomision = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColIVAComision), "############.00"))
                        !INTCVEEMPLEADO = vllngPersona
                        !BITPESOS = IIf(chkDolares.Value = 0, 1, 0)
                        !mnyPagoCuenta = 0
                        !chrestatus = strEstatus
                        !dtmfecha = dtmfecha + fdtmServerHora
                        !mnyretencionRTP = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionRTP), "############.00"))
                        If chkHonorarioFacturado.Value = 0 Then
                            !bitCredito = IIf(aFormasPago(0).vlbolEsCredito, 1, 0)
                        Else
                            !bitCredito = 0
                        End If
                        
                        !chrNumReciboHonorario = grdCargaHonorarios.TextMatrix(vlintContador, cintColReciboHonorario)
                        
                        If chkHonorarioFacturado.Value = 0 Then
                            !bitAfectoCorte = IIf(vlblnEntraCorte Or aFormasPago(0).vlbolEsCredito, 1, 0)
                        Else
                            !bitAfectoCorte = 0
                        End If
                        
                        !dtmFechaAtencionFin = CDate(grdCargaHonorarios.TextMatrix(vlintContador, cintColFechaAtencionFin))
                        !MNYTIPOCAMBIO = IIf(chkDolares.Value = 0, vldblTipoCambio, 0)
                        !smiDeptoCaptura = vgintNumeroDepartamento
                        !smiDeptoCredito = lngDeptoCliente
                        !bitConsultaExterna = IIf(optTipoPaciente(2).Value, 1, 0)
                        !bitPagoDirecto = IIf(blnPagoDirecto, "1", "0")
                        !BITHONORARIOFACTURADO = chkHonorarioFacturado.Value
                        !vchObservaciones = grdCargaHonorarios.TextMatrix(vlintContador, cintColObservaciones)
                        !bitRequiereRecibo = IIf(chkRequiereCFDI.Value = vbChecked, 1, 0)
                        If chkRequiereCFDI.Value = vbChecked Then
                            If chkHonorarioFacturado.Value = 0 Then
                                If aFormasPago(0).vlbolEsCredito Or vlblnUsoCorte Then
                                    strMetodoPago = "PPD"
                                    strSQL = "select CSD.INTIDREGISTRO from GNCatalogoSAT CS" & _
                                    " inner join GNCatalogoSATDetalle CSD on CSD.INTIDCATALOGOSAT = CS.INTIDCATALOGOSAT" & _
                                    " where CS.VCHNOMBRECATALOGO = 'c_FormaPago'" & _
                                    " and CSD.VCHCLAVE = '99'  and CSD.bitActivo = 1"
                                    lngCveFormaPago = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)!intIdRegistro
                                Else
                                    strMetodoPago = "PUE"
                                    lngCveFormaPago = flngCatalogoSATIdByNombreTipo("c_FormaPago", CLng(aFormasPago(0).vlintNumFormaPago), "FP", 0)
                                End If
                            Else
                                strMetodoPago = "PPD"
                                strSQL = "select CSD.INTIDREGISTRO from GNCatalogoSAT CS" & _
                                " inner join GNCatalogoSATDetalle CSD on CSD.INTIDCATALOGOSAT = CS.INTIDCATALOGOSAT" & _
                                " where CS.VCHNOMBRECATALOGO = 'c_FormaPago'" & _
                                " and CSD.VCHCLAVE = '99' and CSD.bitActivo = 1"
                                lngCveFormaPago = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)!intIdRegistro
                            End If
                            !vchMetodoPago = strMetodoPago
                            !intCveFormaPago = lngCveFormaPago
                        End If
                        
                        If aFormasPago(0).vlbolEsCredito Then
                            If blnPagoDirecto Then
                                !bitPagoCXP = 0
                                If lblnPacienteConvenio Then !vchEstatusPortal = "NU"
                            Else
                                !vchEstatusPortal = "CT"
                                !bitPagoCXP = 1
                            End If
                        Else
                            !vchEstatusPortal = "CO"
                            If vlblnUsoCorte Then
                                '!vchEstatusPortal = "CO"
                                !bitPagoCXP = 1
                            Else
                                '!vchEstatusPortal = "PE"
                                !bitPagoCXP = 0
                            End If
                        End If
                        
                        If cboUsoCFDI.ListIndex > -1 Then
                            !intCveUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
                        End If
                        
                        If lngNumCliente > 0 Then
                            !intNumCliente = lngNumCliente
                        End If
                        
                        !vchCorreo = txtEmail.Text
                        If chkRequiereCFDI.Value = vbChecked And Trim(!chrRFC) = Trim(Replace(Replace(Replace(Trim(vgstrRfCCH), "-", ""), "_", ""), " ", "")) Then
                            strSQL = "select vchvalor from siparametro where vchnombre = 'BITPORTALMEDICOS' and intcveempresacontable = " & vgintClaveEmpresaContable
                                    vlIntBitPortal = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)!vchvalor
                            !bitAdjuntarCFDI = vlIntBitPortal
                        Else
                            !bitAdjuntarCFDI = 0
                        End If
                        .Update
                    End With
                    
                    vllngNumHonorario = flngObtieneIdentity("SEC_PVHONORARIO", rsPvHonorario!intConsecutivo)
                    'Adjuntar la carta de autorizacion del seguro
                    If grdCargaHonorarios.TextMatrix(vlintContador, cintColCartaCompleta) <> "" Then
                        Set rsCarta = frsRegresaRs("select * from PVHonorarioDocumentos where intCveHonorario = " & vllngNumHonorario, adLockOptimistic, adOpenStatic)
                        If rsCarta.EOF Then
                            rsCarta.AddNew
                            rsCarta!intCveHonorario = vllngNumHonorario
                        End If
                        rsCarta!vchNombreCartaAutorizacion = grdCargaHonorarios.TextMatrix(vlintContador, cintColCarta)
                        stmCarta.Type = adTypeBinary
                        stmCarta.Open
                        stmCarta.LoadFromFile grdCargaHonorarios.TextMatrix(vlintContador, cintColCartaCompleta)
                        rsCarta!blbCartaAutorizacion = stmCarta.Read
                        rsCarta.Update
                    End If
                    'Enviar el correo del portal
                    Set rsCorreo = frsEjecuta_SP(CInt(vgintClaveEmpresaContable) & "|1", "Sp_CnSelCnCorreo")
                    If Not rsCorreo.EOF Then
                    
                        Set rsEmail = frsRegresaRs("select * from HOMedico where intCveMedico = " & grdCargaHonorarios.TextMatrix(vlintContador, cintColCveMedico), adLockReadOnly, adOpenForwardOnly)
                        If Not rsEmail.EOF Then
                            If IIf(IsNull(rsEmail!bitCorreoHonoraios), 0, rsEmail!bitCorreoHonoraios) <> 0 Then
                                
                                Set clsEmail = New clsCDOmail
            
                                'Se decodifica la contraseña
                                vlPassDecoded = Trim(rsCorreo!vchPassword)
                                
                                'Se reemplazan los caracteres especiales ("U"<-"?"   "l"<-"ñ"   "="<-"Ñ"   "=="<-"Ñ?")
                                vlPassDecoded = Replace(vlPassDecoded, "Ñ?", "==")
                                vlPassDecoded = Replace(vlPassDecoded, "Ñ", "=")
                                vlPassDecoded = Replace(vlPassDecoded, "ñ", "l")
                                vlPassDecoded = Replace(vlPassDecoded, "?", "U")
                                
                                vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (1a vez)
                                vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (2a vez)
                                vlPassDecoded = Decode(vlPassDecoded) 'Se decodifica la contraseña a partir de vlPassDecoded (3a vez)
                                                
                                With clsEmail
                                    'Datos para enviar
                                    .Servidor = Trim(rsCorreo!VCHSERVIDORSMTP)
                                    .Puerto = Val(rsCorreo!intPuerto)
                                    .UseAuntentificacion = True
                                    .SSL = IIf(rsCorreo!BITSSL = 1, True, False)
                                    .Usuario = Trim(rsCorreo!vchCorreo)
                                    .Password = vlPassDecoded
                                    .Asunto = "Honorarios " & txtPaciente.Text & " | " & Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColConcepto))
                                    
                                    '.AdjuntoPDF = IIf(chkPDF.Value = vbChecked, Trim(strArchivoPDF), "")
                                    '.AdjuntoXML = IIf(chkXML.Value = vbChecked, Trim(strArchivoXML), "")
                                    '.AdjuntoZIP = IIf(blnArchivoZIP, Trim(strRutaZIP), "")
                                    .De = Trim(rsCorreo!vchNombre) & " <" & Trim(rsCorreo!vchCorreo) & ">"
                                    .Para = Trim(rsEmail!vchNombre) & " " & Trim(rsEmail!vchApellidoPaterno) & " " & Trim(rsEmail!vchApellidoMaterno) & " <" & Trim(rsEmail!vchEmail) & ">"
                                    
                                    strMensaje = "RFC: " & Trim(Replace(Replace(Replace(Trim(grdCargaHonorarios.TextMatrix(vlintContador, cIntColRFC)), "-", ""), "_", ""), " ", "")) & vbCrLf
                                    strMensaje = strMensaje & "Nombre: " & Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColReciboNombre)) & vbCrLf
                                    strMensaje = strMensaje & "Domicilio: " & IIf(Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColDireccion)) = "", " ", Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColDireccion))) & vbCrLf
                                    strMensaje = strMensaje & "Concepto: " & Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColConcepto)) & vbCrLf
                                    strMensaje = strMensaje & "Método de pago: " & strMetodoPago & vbCrLf
                                    strMensaje = strMensaje & "Forma de pago: " & fstrCatalogoSATCveDescById(lngCveFormaPago, 2) & vbCrLf
                                    If cboUsoCFDI.ListIndex > -1 Then
                                        strMensaje = strMensaje & "Uso del CFDI: " & fstrCatalogoSATCveDescById(cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex), 2) & vbCrLf
                                    End If
                                    strMensaje = strMensaje & "Subtotal: " & Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) & vbCrLf
                                    strMensaje = strMensaje & "Impuestos retenidos: " & Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencion), "############.00")) & vbCrLf
                                    strMensaje = strMensaje & "Total: " & Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) - Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencion), "############.00")) & vbCrLf
                                    If Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColObservaciones)) <> "" Then
                                        strMensaje = strMensaje & vbCrLf & "Observaciones: " & Trim(grdCargaHonorarios.TextMatrix(vlintContador, cintColObservaciones)) & vbCrLf
                                    End If
                                    .mensaje = strMensaje
                                    'Enviar el correo
                                    If .fblnEnviarCorreo = True Then
                                        'Si se envía correctamente, grabar en el log de correos (SILOGCORREOS)
                                        'Call pGuardarLogCorreos(.Usuario, .Para, .CC, .Asunto, .AdjuntoPDF, .AdjuntoXML, .mensaje, lngEmpleado)
                                        
                                        'Guardar el log de transacciones
                                        'Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngEmpleado, Me.Caption, Trim(strFolioDocumento))
                                        'vgblnEnvioExitosoCorreo = True
                                    End If
                                End With
                                Set clsEmail = Nothing 'Se libera el objeto para el envío del correo
    
                            End If
                        End If
                        rsEmail.Close
                    End If
                    rsCorreo.Close
                    
                    If Val(grdCargaHonorarios.TextMatrix(vlintContador, cintColIdTarifa)) <> 0 Then
                        vgstrParametrosSP = CStr(vllngNumHonorario) & "|" & grdCargaHonorarios.TextMatrix(vlintContador, cintColIdTarifa)
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVINSHONORARIOTARIFA"
                    End If
                    
                    If Not blnPagoDirecto Then
                        vllngNumCuentaPagar = 0
                        grdCargaHonorarios.TextMatrix(vlintContador, cintColIdHonorario) = vllngNumHonorario
                        grdCargaHonorarios.TextMatrix(vlintContador, 0) = "*"
                        With rsPvComisionHonorario
                            For vlintContadorComisiones = 1 To grdComisiones.Rows - 1
                                If grdComisiones.TextMatrix(vlintContadorComisiones, 0) = vlintContador Then
                                    .AddNew
                                    !intCveHonorario = vllngNumHonorario
                                    !intComision = grdComisiones.TextMatrix(vlintContadorComisiones, 1)
                                    !MNYCantidad = Val(Format(grdComisiones.TextMatrix(vlintContadorComisiones, 2), "############.00"))
                                    !MNYIVA = Val(Format(grdComisiones.TextMatrix(vlintContadorComisiones, 3), "############.00"))
                                    .Update
                                End If
                            Next vlintContadorComisiones
                        End With
                                            
                        If chkHonorarioFacturado.Value = 0 Then
                            If aFormasPago(0).vlbolEsCredito Then
                                'ES A CRÉDITO
                                'If vgintEmpresa > 0 Then
                                '    vldblCantidad = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.##")) - Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencion), "############.##"))
                                'Else
                                    vldblCantidad = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.##"))
                                'End If
                                vgstrParametrosSP = _
                                fstrFechaSQL(fdtmServerFecha) _
                                & "|" & CStr(rsDatosCliente!intNumCliente) _
                                & "|" & CStr(rsDatosCliente!INTNUMCUENTACONTABLE) _
                                & "|" & CStr(vllngNumHonorario) _
                                & "|" & "HO" _
                                & "|" & vldblCantidad * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)) _
                                & "|" & Str(vgintNumeroDepartamento) _
                                & "|" & Str(vllngPersona) _
                                & "|" & " " & "|" & "0" & "|" & "0" & "|" & "0" _
                                & "|" & IIf(chkRetencion, Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionISR), "############.00")), 0) 'cboTarifa.ItemData
                                lngNumMovimiento = 1
                                frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, lngNumMovimiento
                                
                               'Qué debe actualizar aqui?
                                pEjecutaSentencia "UPDATE CCMOVIMIENTOCREDITO SET MNYRETENCIONRTP = " & IIf(chkRetencionRTP, Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionRTP), "############.00")), 0) & " WHERE intnummovimiento = " & lngNumMovimiento
    
                                'ENTRA AL CORTE
                                For vlintContadorFormasPago = 0 To UBound(aFormasPago(), 1)
                                    vldblProporcion = aFormasPago(vlintContadorFormasPago).vldblCantidad / (CDbl(Format(lblTotalMonto, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)))
                                    vldblCantidadCorte = (CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1))) * vldblProporcion
                                    With rsPvDetalleCorte
                                        .AddNew
                                        !intnumcorte = vllngNumeroCorte
                                        !dtmFechahora = fdtmServerFecha + fdtmServerHora
                                        !chrFolioDocumento = Trim(Str(vllngNumHonorario))
                                        !chrTipoDocumento = "HO"
                                        !intFormaPago = aFormasPago(vlintContadorFormasPago).vlintNumFormaPago
                                        If aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0 Or aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then  'Quiere decir que es una forma de pago en moneda nacional
                                            !mnyCantidadPagada = vldblCantidadCorte
                                        Else
                                            !mnyCantidadPagada = vldblCantidadCorte / vldblTipoCambio
                                        End If
                                        !MNYTIPOCAMBIO = aFormasPago(vlintContadorFormasPago).vldblTipoCambio
                                        !intfoliocheque = IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio) = "", "0", Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio))
                                        !intNumCorteDocumento = vllngNumeroCorte
                                        .Update
                                    End With
                                    
                                    vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                    
                                    If Not aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then
                                        If Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) <> "" And Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) <> "" Then
                                            frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(vlintContadorFormasPago).vldtmFecha))) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                        End If
                                    End If
                                    
                                    pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(vlintContadorFormasPago).vlbolEsCredito = False, aFormasPago(vlintContadorFormasPago).vllngCuentaContable, vllngCtaHonorariosCobrar), vldblCantidadCorte, True
                                    
                                    pEjecutaSentencia "UPDATE PVHONORARIO SET bitafectocorte = 1 WHERE intconsecutivo = " & vllngNumHonorario
                                    
                                    'vldblAcumuladoCXP = vldblCantidadCorte  'vldblAcumuladoCXP + vldblCantidadCorte
                                    vldblAcumuladoCXP = vldblAcumuladoCXP + vldblCantidadCorte
                                Next vlintContadorFormasPago
                                
                                If vldblAcumuladoCXP <> 0 Then
                                    ' Valida sí hubo una comisión
                                    'Sí la hubo, se guarda en la cuenta de comisiones, sí no, en la cuenta de honorarios de CNPARAMETRO
                                    If chkComisiones.Value = 1 Then
        '                                pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "###########.00")) * IIf(chkDolares.Value = 1, vldblTipoCambio, 1), False
                                        pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                    Else
                                        If Not rsPvHonorarioPersonas.EOF Then
        '                                    pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, rsPvHonorarioPersonas(0).Value), CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "###########.00")) * IIf(chkDolares.Value = 1, vldblTipoCambio, 1), False
                                            pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                        End If
                                    End If
                                End If
                            Else
                                'ES EN EFECTIVO, DÓLARES, TARJETA O CHEQUE
                                If vlblnEntraCorte Then
                                    'ENTRA AL CORTE
                                    pGeneraCuentaPorPagar Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColTotalPagar), "############.00")), grdCargaHonorarios.TextMatrix(vlintContador, cintColCveMedico), grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico), vllngNumCuentaPagar, IIf((chkDolares.Value = 1), 0, 1), vllngNumHonorario, Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionISR), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColComision), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColIVAComision), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionRTP), "############.00"))
                                    
                                    For vlintContadorFormasPago = 0 To UBound(aFormasPago(), 1)
                                        vldblProporcion = aFormasPago(vlintContadorFormasPago).vldblCantidad / (CDbl(Format(lblTotalMonto, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)))
                                        vldblCantidadCorte = (CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1))) * vldblProporcion
                                        With rsPvDetalleCorte
                                            .AddNew
                                            !intnumcorte = vllngNumeroCorte
                                            !dtmFechahora = fdtmServerFecha + fdtmServerHora
                                            !chrFolioDocumento = Trim(Str(vllngNumHonorario))
                                            !chrTipoDocumento = "HO"
                                            !intFormaPago = aFormasPago(vlintContadorFormasPago).vlintNumFormaPago
                                            If aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0 Or aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then  'Quiere decir que es una forma de pago en moneda nacional
                                                !mnyCantidadPagada = vldblCantidadCorte
                                            Else
                                                !mnyCantidadPagada = vldblCantidadCorte / vldblTipoCambio
                                            End If
                                            !MNYTIPOCAMBIO = aFormasPago(vlintContadorFormasPago).vldblTipoCambio
                                            !intfoliocheque = IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio) = "", "0", Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio))
                                            !intNumCorteDocumento = vllngNumeroCorte
                                            .Update
                                        End With
                                        
                                        vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                        
                                        If Not aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then
                                            If Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) <> "" And Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) <> "" Then
                                                frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(vlintContadorFormasPago).vldtmFecha))) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                            End If
                                        End If
                                        
                                        pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(vlintContadorFormasPago).vlbolEsCredito = False, aFormasPago(vlintContadorFormasPago).vllngCuentaContable, vllngCtaHonorariosCobrar), vldblCantidadCorte, True
                                        
                                        pEjecutaSentencia "UPDATE PVHONORARIO SET bitafectocorte = 1 WHERE intconsecutivo = " & vllngNumHonorario
                                        
                                        'vldblAcumuladoCXP = vldblCantidadCorte 'vldblAcumuladoCXP + vldblCantidadCorte
                                        vldblAcumuladoCXP = vldblAcumuladoCXP + vldblCantidadCorte
                                        
                                        '-----------------------------------------------------------------------------------'
                                        ' Guardar en el kárdex del banco si hubo pago por medio de transferencias bancarias '
                                        '-----------------------------------------------------------------------------------'
                                        vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(Format(dtmfecha, "dd/mm/yyyy"), Format(dtmHora, "hh:mm:ss")) & "|" & aFormasPago(vlintContadorFormasPago).vlintNumFormaPago & "|" & aFormasPago(vlintContadorFormasPago).lngIdBanco & "|" & _
                                                            IIf(aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0, aFormasPago(vlintContadorFormasPago).vldblCantidad, aFormasPago(vlintContadorFormasPago).vldblDolares) & "|" & IIf(aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContadorFormasPago).vldblTipoCambio & "|" & _
                                                            fstrTipoMovimientoForma(aFormasPago(vlintContadorFormasPago).vlintNumFormaPago) & "|" & "HO" & "|" & vllngNumHonorario & "|" & vllngPersona & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(dtmfecha, "dd/mm/yyyy"), Format(dtmHora, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo
                                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                                                                                
                                    Next vlintContadorFormasPago
                                    
                                    If vldblAcumuladoCXP <> 0 Then
                                        ' Valida sí hubo una comisión
                                        'Sí la hubo, se guarda en la cuenta de comisiones, sí no, en la cuenta de honorarios de CNPARAMETRO
                                        If chkComisiones.Value = 1 Then
            '                                pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "###########.00")) * IIf(chkDolares.Value = 1, vldblTipoCambio, 1), False
                                            pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                        Else
                                            If Not rsPvHonorarioPersonas.EOF Then
            '                                    pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, rsPvHonorarioPersonas(0).Value), CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "###########.00")) * IIf(chkDolares.Value = 1, vldblTipoCambio, 1), False
                                                pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                            End If
                                        End If
                                    End If
                                Else
                                    'NO ENTRA AL CORTE
                                    
                                    'vldblAcumuladoCXP = 0
                                    
                                    For vlintContadorFormasPago = 0 To UBound(aFormasPago(), 1)
                                        Set rsValor = frsRegresaRs("SELECT chrtipo FROM PVFORMAPAGO WHERE intformapago = " & aFormasPago(vlintContadorFormasPago).vlintNumFormaPago)
                                        If rsValor.RecordCount <> 0 Then
                                            If rsValor!chrTipo = "T" Or rsValor!chrTipo = "H" Or rsValor!chrTipo = "B" Then
                                                
                                                If Not vlblnUsoCorte Then
                                                    vlblnUsoCorte = True
                                        
                                                    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                                                    vllngCorteGrabando = 1
                                                    vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                                                    frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                                                    If vllngCorteGrabando <> 2 Then
                                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                                        Exit Sub
                                                    End If
                                                End If
                                            
                                                vldblProporcion = aFormasPago(vlintContadorFormasPago).vldblCantidad / (CDbl(Format(lblTotalMonto, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)))
                                                vldblCantidadCorte = (CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1))) * vldblProporcion
                                                With rsPvDetalleCorte
                                                    .AddNew
                                                    !intnumcorte = vllngNumeroCorte
                                                    !dtmFechahora = fdtmServerFecha + fdtmServerHora
                                                    !chrFolioDocumento = Trim(Str(vllngNumHonorario))
                                                    !chrTipoDocumento = "HO"
                                                    !intFormaPago = aFormasPago(vlintContadorFormasPago).vlintNumFormaPago
                                                    If aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0 Or aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then  'Quiere decir que es una forma de pago en moneda nacional
                                                        !mnyCantidadPagada = vldblCantidadCorte
                                                    Else
                                                        !mnyCantidadPagada = vldblCantidadCorte / vldblTipoCambio
                                                    End If
                                                    !MNYTIPOCAMBIO = aFormasPago(vlintContadorFormasPago).vldblTipoCambio
                                                    !intfoliocheque = IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio) = "", "0", Trim(aFormasPago(vlintContadorFormasPago).vlstrFolio))
                                                    !intNumCorteDocumento = vllngNumeroCorte
                                                    .Update
                                                End With
                                                
                                                vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVDETALLECORTE", 0)
                                                                
                                                If Not aFormasPago(vlintContadorFormasPago).vlbolEsCredito Then
                                                    If Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) <> "" And Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) <> "" Then
                                                        frsEjecuta_SP vllngNumeroCorte & "|" & vllngNumDetalleCorte & "|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrRFC) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(vlintContadorFormasPago).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(vlintContadorFormasPago).vldtmFecha))) & "'|'" & Trim(aFormasPago(vlintContadorFormasPago).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                                                    End If
                                                End If
                                                
                                                pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(vlintContadorFormasPago).vlbolEsCredito = False, aFormasPago(vlintContadorFormasPago).vllngCuentaContable, vllngCtaHonorariosCobrar), vldblCantidadCorte, True
                                                
                                                pEjecutaSentencia "UPDATE PVHONORARIO SET bitafectocorte = 1 WHERE intconsecutivo = " & vllngNumHonorario
                                                
                                                pEjecutaSentencia "UPDATE PVHONORARIO SET chrestatus = 'M' WHERE intconsecutivo = " & vllngNumHonorario
                                                
                                                'vldblAcumuladoCXP = vldblCantidadCorte 'vldblAcumuladoCXP + vldblCantidadCorte
                                                vldblAcumuladoCXP = vldblAcumuladoCXP + vldblCantidadCorte
                                                
                                                '-----------------------------------------------------------------------------------'
                                                ' Guardar en el kárdex del banco si hubo pago por medio de transferencias bancarias '
                                                '-----------------------------------------------------------------------------------'
                                                vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(Format(dtmfecha, "dd/mm/yyyy"), Format(dtmHora, "hh:mm:ss")) & "|" & aFormasPago(vlintContadorFormasPago).vlintNumFormaPago & "|" & aFormasPago(vlintContadorFormasPago).lngIdBanco & "|" & _
                                                                    IIf(aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0, aFormasPago(vlintContadorFormasPago).vldblCantidad, aFormasPago(vlintContadorFormasPago).vldblDolares) & "|" & IIf(aFormasPago(vlintContadorFormasPago).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContadorFormasPago).vldblTipoCambio & "|" & _
                                                                    fstrTipoMovimientoForma(aFormasPago(vlintContadorFormasPago).vlintNumFormaPago) & "|" & "HO" & "|" & vllngNumHonorario & "|" & vllngPersona & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(dtmfecha, "dd/mm/yyyy"), Format(dtmHora, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo
                                                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                                                
                                            End If
                                        End If
                                    Next vlintContadorFormasPago
                                    
                                    If vldblAcumuladoCXP <> 0 Then
                                    
                                        ' Valida sí hubo una comisión
                                        'Sí la hubo, se guarda en la cuenta de comisiones, sí no, en la cuenta de honorarios de CNPARAMETRO
                                        If chkComisiones.Value = 1 Then
                                            pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                        Else
                                            If Not rsPvHonorarioPersonas.EOF Then
                                                pInsCortePoliza vllngNumeroCorte, vllngNumHonorario, "HO", IIf(aFormasPago(0).vlbolEsCredito, vllngCtaHonorariosPagar, grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)), vldblAcumuladoCXP, False
                                            End If
                                        End If
                                    
                                        vldblProporcion = vldblAcumuladoCXP / ((CDbl(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1))))
                                        
                                        pGeneraCuentaPorPagar (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColTotalPagar), "############.00")) * vldblProporcion), grdCargaHonorarios.TextMatrix(vlintContador, cintColCveMedico), grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico), vllngNumCuentaPagar, IIf((chkDolares.Value = 1), 0, 1), vllngNumHonorario, (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")) * vldblProporcion), (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionISR), "############.00")) * vldblProporcion), (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColComision), "############.00")) * vldblProporcion), (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColIVAComision), "############.00")) * vldblProporcion), (Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionRTP), "############.00")) * vldblProporcion)
                                    End If
                                End If
                            End If
                        Else
                            pGeneraCuentaPorPagar Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColTotalPagar), "############.00")), grdCargaHonorarios.TextMatrix(vlintContador, cintColCveMedico), grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico), vllngNumCuentaPagar, IIf((chkDolares.Value = 1), 0, 1), vllngNumHonorario, Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionISR), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColComision), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColIVAComision), "############.00")), Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColRetencionRTP), "############.00"))
                        
                            Set rsCnPoliza = frsRegresaRs("SELECT * FROM CnPoliza WHERE intNumeroPoliza = -1", adLockOptimistic, adOpenDynamic)
                            With rsCnPoliza
                                .AddNew
                                !tnyclaveempresa = vgintClaveEmpresaContable
                                !smiEjercicio = Year(fdtmServerFecha)
                                !tnyMes = Month(fdtmServerFecha)
                                !dtmFechaPoliza = fdtmServerFecha
                                !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, "D", Year(fdtmServerFecha), Month(fdtmServerFecha), False)
                                !chrTipoPoliza = "D"
                                !vchConceptoPoliza = "HONORARIO POR PAGAR A " & Trim(cboMedicos.Text)
                                !smicvedepartamento = vgintNumeroDepartamento
                                !INTCVEEMPLEADO = vllngPersona
                                !vchNumero = " "
                                !bitAsentada = 0
                                .Update
                            End With
                            vllngPolizaMaestro = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!intNumeroPoliza)
                            rsCnPoliza.Close
                            
                            Set rsCnDetallePoliza = frsRegresaRs("select * from CnDetallePoliza where intNumeroPoliza = -1 ", adLockOptimistic, adOpenDynamic)
                            With rsCnDetallePoliza
                                .AddNew
                                !intNumeroPoliza = vllngPolizaMaestro
                                !intNumeroCuenta = vllngCtaCostoHonorarioFacturado
                                !bitNaturalezaMovimiento = 1
                                !mnyCantidadMovimiento = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00"))
                                !vchReferencia = " "
                                !vchConcepto = " "
                                .Update
                            End With
                            rsCnDetallePoliza.Close
                            
                            Set rsCnDetallePoliza = frsRegresaRs("select * from CnDetallePoliza where intNumeroPoliza = -1 ", adLockOptimistic, adOpenDynamic)
                            With rsCnDetallePoliza
                                .AddNew
                                !intNumeroPoliza = vllngPolizaMaestro
                                !intNumeroCuenta = grdCargaHonorarios.TextMatrix(vlintContador, cintColCuentaContableMedico)
                                !bitNaturalezaMovimiento = 0
                                !mnyCantidadMovimiento = Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColMonto), "############.00"))
                                !vchReferencia = " "
                                !vchConcepto = " "
                                .Update
                            End With
                            rsCnDetallePoliza.Close
                        End If
                    End If
                    
                    ' Update del portal
                    pEjecutaSentencia "UPDATE PVHONORARIO SET bitPagoCXP = 1, vchEstatusPortal = 'CO' where bitAfectoCorte = 1 and bitPagoCXP = 0 and vchEstatusPortal = 'PE' and bitPagoDirecto = 0 and intconsecutivo = " & vllngNumHonorario

                Next vlintContador
                                            
                vlstrSentencia = "delete from PvTipoPacienteProceso where PvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
                    "and PvTipoPacienteProceso.intproceso = " & enmTipoProceso.Honorarios
                pEjecutaSentencia vlstrSentencia
                vlstrSentencia = "insert into PvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.Honorarios & "," & IIf(optTipoPaciente(0).Value, "'I'", IIf(optTipoPaciente(1).Value, "'E'", "'C'")) & ")"
                pEjecutaSentencia vlstrSentencia
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersona, "CAPTURA DE HONORARIOS", txtMovimientoPaciente.Text)
                
                If chkHonorarioFacturado.Value = 0 Then
                    If vlblnEntraCorte Or aFormasPago(0).vlbolEsCredito Or vlblnUsoCorte Then
                        pLiberaCorte vllngNumeroCorte
                    End If
                End If
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                If optTipoPaciente(2).Value Then
                    'La información fue guardada con la cuenta:
                    MsgBox SIHOMsg(725) + " " + txtMovimientoPaciente.Text, vbInformation + vbOKOnly, "Mensaje"
                    lblCuentaPaciente.Caption = "Número de cuenta"
                    optTipoPaciente(1).Value = True
                    optTipoPaciente(2).Visible = False
                Else
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
                    chkComisiones.Enabled = False
                End If
                
                If chkHonorarioFacturado.Value = 0 Then
                    If aFormasPago(0).vlbolEsCredito Then
                        pLimpia
                        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                        txtMovimientoPaciente.SetFocus
                    Else
                        fraHonorarios.Enabled = False
                        fraRecibo.Enabled = False
                        fraReciboHonorario.Enabled = False
                        pHabilita 0, 0, 0, 0, 0, 0, 1, 0
                        cmdPrint.SetFocus
                    End If
                Else
                    fraHonorarios.Enabled = False
                    fraRecibo.Enabled = False
                    fraReciboHonorario.Enabled = False
                    pHabilita 0, 0, 0, 0, 0, 0, 1, 0
                    cmdPrint.SetFocus
                End If
            End If
        End If
        rsPvHonorarioPersonas.Close
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSave"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError

    Dim rsCuentasPuente As New ADODB.Recordset
    Dim vllngMensaje As Long
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    If vlblnEntrando Then
         'Tarifas de ISR
        Set rs = frsEjecuta_SP("-1|1", "SP_CNSELTARIFAISR")
        If rs.RecordCount = 0 Then MsgBox SIHOMsg(1411), vbExclamation, "Mensaje"
        
        vlblnEntrando = False
        
        vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    
        If vllngMensaje <> 0 Then
            'Cierre el corte actual antes de registrar este documento.
            'No existe un corte abierto.
            MsgBox SIHOMsg(Str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje"
            Unload Me
        Else
            vllngCtaHonorariosCobrar = 0
            vllngCtaHonorariosPagar = 0
            Set rsCuentasPuente = frsSelParametros("CN", vgintClaveEmpresaContable, "INTCTAHONORARIOSCOBRAR")
            If rsCuentasPuente.RecordCount <> 0 Then
                If Not IsNull(rsCuentasPuente!valor) Then vllngCtaHonorariosCobrar = rsCuentasPuente!valor
            End If
            Set rsCuentasPuente = frsSelParametros("CN", vgintClaveEmpresaContable, "INTCTAHONORARIOSPAGAR")
            If rsCuentasPuente.RecordCount <> 0 Then
                If Not IsNull(rsCuentasPuente!valor) Then vllngCtaHonorariosPagar = rsCuentasPuente!valor
            End If
            
'            vlstrSentencia = "select intNumCuentaHonorarioCobrar,intNumCuentaHonorarioPagar from ccparametro"
'            Set rsCuentasPuente = frsRegresaRs(vlstrSentencia)
'            If rsCuentasPuente.RecordCount <> 0 Then
'                If Not IsNull(rsCuentasPuente!intnumcuentahonorariocobrar) Then
'                    vllngCtaHonorariosCobrar = rsCuentasPuente!intnumcuentahonorariocobrar
'                End If
'                If Not IsNull(rsCuentasPuente!intnumcuentahonorarioPagar) Then
'                    vllngCtaHonorariosPagar = rsCuentasPuente!intnumcuentahonorarioPagar
'                End If
'            End If
            rsCuentasPuente.Close
            
            txtMovimientoPaciente.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Activate"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError

    If SSTabHonorario.Tab = 0 Then
        If fraComisiones.Visible Then
            Cancel = True
            If cmdAceptarComisiones.Enabled Then
                cmdAceptarComisiones_Click
            End If
        Else
            If cmdSave.Enabled Or vlblnConsulta Or Val(txtMovimientoPaciente.Text) <> 0 Then
                Cancel = True
                '¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pLimpia
                    If lblnRecargarTarifas Then
                        pCargaTarifas
                    End If
                    If txtMovimientoPaciente.Enabled Then
                        txtMovimientoPaciente.SetFocus
                    End If
                    pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                End If
            End If
        End If
    Else
        Cancel = True
        SSTabHonorario.Tab = 0
        pLimpia
        If lblnRecargarTarifas Then
            pCargaTarifas
        End If
        pHabilita 0, 0, 1, 0, 0, 0, 0, 0
        If txtMovimientoPaciente.Enabled Then
            txtMovimientoPaciente.SetFocus
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_QueryUnload"))
End Sub

Private Sub grdCargaHonorarios_Click()
    'If grdCargaHonorarios.Row > 0 And grdCargaHonorarios.Col = cintColPagoDirecto Then
    '    If Trim(grdCargaHonorarios.TextMatrix(grdCargaHonorarios.Row, cintColCveMedico)) <> "" Then
    '        grdCargaHonorarios.TextMatrix(grdCargaHonorarios.Row, cintColPagoDirecto) = IIf(grdCargaHonorarios.TextMatrix(grdCargaHonorarios.Row, cintColPagoDirecto) = "", "*", "")
    '    End If
    'End If
End Sub

Private Sub grdCargaHonorarios_DblClick()
    On Error GoTo NotificaError
    If grdCargaHonorarios.Col = cintColPagoDirecto Then Exit Sub
    With grdCargaHonorarios
        If .TextMatrix(.Row, 2) <> "" And .Row > 0 Then
            If .Rows > 2 Then
                pBorrarRegMshFGrdData grdCargaHonorarios.Row, grdCargaHonorarios
                .Redraw = True
                
                If Not fraRecibo.Enabled Then
                    fraRecibo.Enabled = True
                    txtReciboHonorario.Enabled = True
                    
                    cboMedicos.SetFocus
                End If
            Else
                grdCargaHonorarios.RowData(1) = -1
                pConfigura
                
                'chkHonorarioFacturado.Enabled = True
                
                If Not fraRecibo.Enabled Then
                    fraRecibo.Enabled = True
                    txtReciboHonorario.Enabled = True
                    
                    cboMedicos.SetFocus
                End If
            End If
        Else
            grdCargaHonorarios.RowData(1) = -1
            pConfigura
        End If
    End With
    
    pCalculaTotales
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAtencionFin_LostFocus"))
End Sub

Private Sub grdHonorarios_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo NotificaError
    
     grdHonorarios.TextMatrix(grdHonorarios.Row, 0) = "*"
    If grdHonorarios.Row = grdHonorarios.Rows - 1 Then
       
        cmdActualizarRecibo.SetFocus
    Else
        grdHonorarios.Row = grdHonorarios.Row + 1
         cmdActualizarRecibo.Enabled = True
         cmdPagoAnticipado.Enabled = True
         cmdDelete.Enabled = True
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHonorarios_AfterEdit"))
End Sub

Private Sub grdHonorarios_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo NotificaError

    If ((grdHonorarios.Col <> 8) Or (Val(grdHonorarios.TextMatrix(grdHonorarios.Row, 1)) = 0)) Then
        Cancel = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHonorarios_BeforeEdit"))
End Sub

Private Sub grdHonorarios_EnterCell()
    If grdHonorarios.Row = 1 Then
        If grdHonorarios.TextMatrix(1, 1) = "" Then Exit Sub
    End If
    If grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilEstatus) = "C" Then Exit Sub
        
    
    If grdHonorarios.Col = cintColFilCarta Then
        cmdCargarCarta.Enabled = True
        cmdLimpiarCarta.Enabled = True
        cmdAbrirCarta.Enabled = True
    Else
        cmdCargarCarta.Enabled = False
        cmdLimpiarCarta.Enabled = False
        cmdAbrirCarta.Enabled = False
    End If
End Sub

Private Sub grdHonorarios_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then grdHonorarios_DblClick

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHonorarios_KeyDown"))
End Sub

Private Sub grdHonorarios_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub mskFechaAtencionFin_GotFocus()
    On Error GoTo NotificaError

    pSelMkTexto mskFechaAtencionFin
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAtencionFin_GotFocus"))
End Sub

Private Sub mskFechaAtencionFin_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then chkRequiereCFDI.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAtencionFin_KeyPress"))
End Sub

Private Sub mskFin_GotFocus()
    pSelMkTexto mskFin
End Sub

Private Sub mskFin_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then cmdCargar.SetFocus

End Sub

Private Sub mskFin_LostFocus()

    If Not IsDate(mskFin.Text) And Trim(mskFin.ClipText) <> "" Then
        mskFin.Mask = ""
        mskFin.Text = ldtmFecha
        mskFin.Mask = "##/##/####"
    End If

End Sub

Private Sub mskInicio_GotFocus()
    pSelMkTexto mskInicio
End Sub

Private Sub mskInicio_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then mskFin.SetFocus

End Sub

Private Sub mskInicio_LostFocus()

    If Not IsDate(mskInicio.Text) And Trim(mskInicio.ClipText) <> "" Then
        mskInicio.Mask = ""
        mskInicio.Text = ldtmFecha
        mskInicio.Mask = "##/##/####"
    End If

End Sub

Private Sub optFilTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    If SSTabHonorario.Tab <> 0 Then
        txtFilCuenta.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":optFilTipoPaciente_Click"))
End Sub

Private Sub txtAdjuntar_GotFocus()
    pSelTextBox txtAdjuntar
End Sub

Private Sub txtAdjuntar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 8 Then
        txtAdjuntar.Text = ""
    End If
End Sub

Private Sub txtEmail_GotFocus()
    pSelTextBox txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoHonorario.SetFocus
    End If
End Sub

Private Sub txtFilCuenta_Change()
    txtFilPaciente.Text = ""
End Sub

Private Sub txtFilCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    Dim rsNombrePaciente As New ADODB.Recordset
    Dim vlstrSentencia As String

    If KeyCode = vbKeyReturn Then
        If RTrim(txtFilCuenta.Text) = "" Then
            With FrmBusquedaPacientes
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSinFacturar.Enabled = False
                .optSoloActivos.Enabled = False
                .optTodos.Value = True
                .optTodos.Enabled = False

                If optFilTipoPaciente(1).Value Then 'Externos
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as Fecha, isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " Externos"
                Else
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as INGRESO , ExPacienteIngreso.dtmFechaHoraEgreso as EGRESO, isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                    .vgstrTipoPaciente = "I"
                    .Caption = .Caption & " Internos"
                End If

                txtFilCuenta.Text = .flngRegresaPaciente()

                If txtFilCuenta.Text <> -1 Then
                    If optFilTipoPaciente(1).Value Then
                        vlstrSentencia = "Select ltrim(rtrim(Externo.chrApePaterno))||' '||ltrim(rtrim(Externo.chrApeMaterno))||' '||ltrim(rtrim(Externo.chrNombre)) as Paciente "
                        vlstrSentencia = vlstrSentencia & " From Externo inner join RegistroExterno on RegistroExterno.intNumPaciente = Externo.intNumPaciente "
                        vlstrSentencia = vlstrSentencia & " Where RegistroExterno.intNumcuenta = " & txtFilCuenta.Text
                    Else
                        vlstrSentencia = "select ltrim(rtrim(AdPaciente.vchApellidoPaterno))||' '||ltrim(rtrim(AdPaciente.vchApellidoMaterno))||' '||ltrim(rtrim(AdPaciente.vchNombre)) as Paciente "
                        vlstrSentencia = vlstrSentencia & " From AdPaciente inner join AdAdmision on AdAdmision.numCvePaciente = AdPaciente.numCvePaciente "
                        vlstrSentencia = vlstrSentencia & " where AdAdmision.numNumCuenta= " & txtFilCuenta.Text
                    End If
                    
                    Set rsNombrePaciente = frsRegresaRs(vlstrSentencia)
                    If rsNombrePaciente.RecordCount > 0 Then
                        txtFilPaciente.Text = rsNombrePaciente!Paciente
                        cboFilMedico.SetFocus
                    End If
                    rsNombrePaciente.Close
                Else
                    txtFilCuenta.Text = ""
                    txtFilPaciente.Text = ""
                End If
            End With
        Else
            If optFilTipoPaciente(1).Value Then
                vlstrSentencia = "Select ltrim(rtrim(Externo.chrApePaterno))||' '||ltrim(rtrim(Externo.chrApeMaterno))||' '||ltrim(rtrim(Externo.chrNombre)) as Paciente "
                vlstrSentencia = vlstrSentencia & " From Externo inner join RegistroExterno on RegistroExterno.intNumPaciente = Externo.intNumPaciente "
                vlstrSentencia = vlstrSentencia & " Where RegistroExterno.intNumcuenta = " & txtFilCuenta.Text
            Else
                vlstrSentencia = "select ltrim(rtrim(AdPaciente.vchApellidoPaterno))||' '||ltrim(rtrim(AdPaciente.vchApellidoMaterno))||' '||ltrim(rtrim(AdPaciente.vchNombre)) as Paciente "
                vlstrSentencia = vlstrSentencia & " From AdPaciente inner join AdAdmision on AdAdmision.numCvePaciente = AdPaciente.numCvePaciente "
                vlstrSentencia = vlstrSentencia & " where AdAdmision.numNumCuenta= " & txtFilCuenta.Text
            End If
            Set rsNombrePaciente = frsRegresaRs(vlstrSentencia)
            If rsNombrePaciente.RecordCount > 0 Then
                txtFilPaciente.Text = rsNombrePaciente!Paciente
                cboFilMedico_KeyDown vbKeyReturn, 0
                cboFilMedico.SetFocus
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                txtFilPaciente.Text = ""
                pEnfocaTextBox txtFilCuenta
            End If
            rsNombrePaciente.Close
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFilCuenta_KeyDown"))
End Sub

Private Sub txtFilCuenta_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optFilTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            optFilTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFilCuenta_KeyPress"))
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim lngNumeroPaciente As Long 'Número de paciente que se dió de alta o se actualizó para guardar el honorario de consulta externa
    Dim lngnumCuenta As Long 'para llamar el procedimiento de datos de paciente
    Dim strTipoPaciente As String 'para llamar el procedimiento de datos de paciente

    If KeyCode = vbKeyReturn Then
        If Trim(txtMovimientoPaciente.Text) = "" Then
        
            If optTipoPaciente(2).Value Then
                'Paciente de consulta externa
                'Registrar o buscar un paciente de consulta externa
                
                If Not cgstrModulo = "CC" Then ' si es del modulo de caja
                    
                    txtMovimientoPaciente.Text = ""
                    
                    frmAdmisionPaciente.vlblnMostrarTabGenerales = True
                    frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
                    frmAdmisionPaciente.vlblnMostrarTabInternos = False
                    frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
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
                    
                    frmAdmisionPaciente.vllngNumeroOpcionExterno = 1806
                    frmAdmisionPaciente.Show vbModal, Me
                    
                    lngNumeroPaciente = frmAdmisionPaciente.vglngExpediente
                    
                    If lngNumeroPaciente > 0 Then
                        
                        vlblnLimpiar = False
                        txtMovimientoPaciente.Text = lngNumeroPaciente
                        If fblnDatosPaciente(0, lngNumeroPaciente, "E") Then
                        
                            vlblnPacienteSeleccionado = True
                            pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                            
                            'chkHonorarioFacturado.Enabled = True
                            
                            fraRecibo.Enabled = True
                            txtReciboHonorario.Enabled = True
                            
                            cmdTodosMedicos_Click
                            
                            cboMedicos.SetFocus
                            
                            'aparecer la lista para agregar honorarios
                            grdCargaHonorarios.RowData(1) = -1
                            pConfigura
                            fraHonorarios.Visible = True
                            cmdAgregar.Visible = True
                            cmdAgregar.Enabled = True
                            grdComisiones.Rows = 1
                        End If
                    End If
                
                Else ' Si es del modulo de cxc
                    
                    With FrmBusquedaPacientes
                        .vgblnPideClave = False
                        .vgIntMaxRecords = 100
                        .vgstrMovCve = "C"
                        .optSinFacturar.Enabled = False
                        .optSoloActivos.Enabled = False
                        .optTodos.Value = True
                        .optTodos.Enabled = False
                        .vgStrOtrosCampos = ""
                        .vgstrTamanoCampo = "800,3400,1700,4100"
                        .vgstrTipoPaciente = "E"
                        .Caption = .Caption & " Externos"
                        .vgblndecredito = 1
                        lngNumeroPaciente = .flngRegresaPaciente()
                    End With
                    
                    If lngNumeroPaciente = -1 Then ' si el paciente no existe lo da de alta con la pantalla de pac.externos con algunos campos inhabilitados

                        frmAdmisionPaciente.vlblnMostrarTabGenerales = True
                        frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
                        frmAdmisionPaciente.vlblnMostrarTabInternos = False
                        frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
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
                        frmAdmisionPaciente.blnHonorariosCC = True
                        
                        frmAdmisionPaciente.vllngNumeroOpcionExterno = 1806
                        frmAdmisionPaciente.Show vbModal, Me
                        
                        lngNumeroPaciente = frmAdmisionPaciente.vglngExpediente
                        
                        If lngNumeroPaciente > 0 Then
                            vlblnLimpiar = False
                            txtMovimientoPaciente.Text = lngNumeroPaciente
                            If fblnDatosPaciente(0, lngNumeroPaciente, "E") Then
                                vlblnPacienteSeleccionado = True
                                pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                                
                                'chkHonorarioFacturado.Enabled = True
                                
                                fraRecibo.Enabled = True
                                txtReciboHonorario.Enabled = True
                                cmdTodosMedicos_Click
                                cboMedicos.SetFocus
                                'aparecer la lista para agregar honorarios
                                grdCargaHonorarios.RowData(1) = -1
                                pConfigura
                                fraHonorarios.Visible = True
                                cmdAgregar.Visible = True
                                cmdAgregar.Enabled = True
                                grdComisiones.Rows = 1
                            End If
                        End If
                    
                    Else
                    
                        If lngNumeroPaciente > 0 Then
                            vlblnLimpiar = False
                            txtMovimientoPaciente.Text = lngNumeroPaciente
                            If fblnDatosPaciente(0, lngNumeroPaciente, "E") Then
                                vlblnPacienteSeleccionado = True
                                pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                                
                                'chkHonorarioFacturado.Enabled = True
                                
                                fraRecibo.Enabled = True
                                txtReciboHonorario.Enabled = True
                                cmdTodosMedicos_Click
                                cboMedicos.SetFocus
                                'aparecer la lista para agregar honorarios
                                grdCargaHonorarios.RowData(1) = -1
                                pConfigura
                                fraHonorarios.Visible = True
                                cmdAgregar.Visible = True
                                cmdAgregar.Enabled = True
                                grdComisiones.Rows = 1
                            End If
                        End If
                        
                    End If
                    
                End If
            Else
                With FrmBusquedaPacientes
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Enabled = False
                    .optSoloActivos.Enabled = False
                    .optTodos.Value = True
                    .optTodos.Enabled = False
    
                    If optTipoPaciente(1).Value Then 'Externos
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as Fecha, isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,1700,4100"
                        .vgstrTipoPaciente = "E"
                        .Caption = .Caption & " Externos"
                    Else
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as INGRESO, ExPacienteIngreso.dtmFechaHoraEgreso as EGRESO, isnull(CCempresa.vchDescripcion, adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,990,990,4100"
                        .vgstrTipoPaciente = "I"
                        .Caption = .Caption & " Internos"
                    End If
    
                    lngnumCuenta = .flngRegresaPaciente()
    
                    If lngnumCuenta <> -1 Then
                        txtMovimientoPaciente.Text = lngnumCuenta
                        vlblnLimpiar = False
                        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                    End If
                End With
            End If
         Else
            lngnumCuenta = IIf(Not optTipoPaciente(2).Value, CLng(txtMovimientoPaciente.Text), 0)
            lngNumeroPaciente = IIf(optTipoPaciente(2).Value, CLng(txtMovimientoPaciente.Text), 0)
            strTipoPaciente = IIf(Not optTipoPaciente(0).Value, "E", "I")
         
            If fblnDatosPaciente(lngnumCuenta, lngNumeroPaciente, strTipoPaciente) Then
                
                vlblnPacienteSeleccionado = True
                pHabilita 0, 0, 0, 0, 0, 1, 0, 0
                
                'chkHonorarioFacturado.Enabled = True
                
                fraRecibo.Enabled = True
                txtReciboHonorario.Enabled = True
                pCargaMedicos

                cboMedicos.SetFocus
                'aparecer la lista para agregar honorarios
                grdCargaHonorarios.RowData(1) = -1
                cmdAgregar.Visible = True
                cmdAgregar.Enabled = True
                grdComisiones.Rows = 1
            
            End If
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyDown"))
End Sub

Private Sub pCalculoTotales()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vldblComisionTotal As Double
    Dim vldblIVA As Double
    Dim dblPorRetencion As Double
   
    

    vldblComisionTotal = 0
    vldblIVA = 0
    ReDim vlComisiones(0)

    dblPorRetencion = 0
    If cboTarifa.ListIndex <> -1 Then
        dblPorRetencion = arrTarifas(cboTarifa.ListIndex).dblPorcentaje
    End If

    If chkRetencion.Value = 1 Then
        txtRetencion.Text = FormatCurrency(Str(Val(Format(txtMontoHonorario.Text, "############.##")) * dblPorRetencion / 100), 2)
    Else
        txtRetencion.Text = FormatCurrency(0, 2)
    End If
    If chkRetencionRTP.Value = 1 Then
        txtRetencionRTP.Text = FormatCurrency(Str(Val(Format(txtMontoHonorario.Text, "############.##")) * dblPorcentajeRTP / 100), 2)
    Else
        txtRetencionRTP.Text = FormatCurrency(0, 2)
    End If
    dblTotalRetencionDetalle = Val(Format(txtRetencion.Text, "############.##")) + Val(Format(txtRetencionRTP.Text, "############.##"))
    
    txtNetoPagar.Text = FormatCurrency(Val(Format(txtMontoHonorario.Text, "############.##")) - dblTotalRetencionDetalle, 2)
    For vlintContador = 0 To lstComisionesAsignadas.ListCount - 1
        ReDim Preserve vlComisiones(vlintContador + 1)
        vlstrSentencia = "select mnyComision Comision, smyIva IVA from pvComision where smiCveComision = " & Trim(Str(lstComisionesAsignadas.ItemData(vlintContador)))
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        vldblComisionTotal = vldblComisionTotal + (rs!Comision * Val(Format(txtMontoHonorario.Text, "##############.##")) / 100)
        vldblIVA = vldblIVA + (rs!Comision * Val(Format(txtMontoHonorario.Text, "##############.##")) / 100) * rs!IVA / 100
        vlComisiones(vlintContador).vllngCveComision = lstComisionesAsignadas.ItemData(vlintContador)
        vlComisiones(vlintContador).vldblCantidad = (rs!Comision * Val(Format(txtMontoHonorario.Text, "##############.##")) / 100)
        vlComisiones(vlintContador).vldblIVA = (rs!Comision * Val(Format(txtMontoHonorario.Text, "##############.##")) / 100) * rs!IVA / 100
        rs.Close
    Next
    txtCantidadComision.Text = FormatCurrency(vldblComisionTotal, 2)
    txtIVAComision.Text = FormatCurrency(vldblIVA, 2)
    txtPagoCredito.Text = FormatCurrency(vldblPagoCuenta, 2)
    txtTotal.Text = FormatCurrency(Val(Format(txtNetoPagar.Text, "##############.##")) - Val(Format(txtCantidadComision.Text, "##############.##")) - Val(Format(txtIVAComision.Text, "##############.##")) - vldblPagoCuenta, 2)

    chkComisiones.Value = IIf(vldblComisionTotal = 0, 0, 1)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculoTotales"))
End Sub

Private Sub pCargaMedicos()
    On Error GoTo NotificaError
    Dim rsMedicos As New ADODB.Recordset
    
    'Médicos
    vlstrSentencia = "SELECT  HoMedico.intCveMedico, " & _
        " rtrim(ltrim(HoMedico.vchApellidoPaterno)) || ' ' || " & _
        " rtrim(ltrim(HoMedico.vchApellidoMaterno)) || ' ' || " & _
        " rtrim(LTrim(HoMedico.vchNombre)) As Nombre " & _
        " From exMedicoACargo " & _
        " INNER JOIN HoMedico ON Homedico.intCveMedico = exMedicoACargo.intCveMedico " & _
        " Where exMedicoACargo.numNumCuenta = " & Trim(txtMovimientoPaciente.Text) & _
        " and ExMedicoACargo.chrTipoPaciente = " & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & _
        " union "

    If optTipoPaciente(0).Value Then 'Internos
        vlstrSentencia = vlstrSentencia & _
            " SELECT  HoMedico.intCveMedico, " & _
            " rtrim(ltrim(HoMedico.vchApellidoPaterno)) || ' ' || " & _
            " rtrim(ltrim(HoMedico.vchApellidoMaterno)) || ' ' || " & _
            " rtrim(LTrim(HoMedico.vchNombre)) As Nombre " & _
            " From AdAdmision " & _
            " INNER JOIN HoMedico ON Homedico.intCveMedico = adAdmision.intCveMedicoCargo " & _
            " Where AdAdmision.numNumCuenta = " & Trim(txtMovimientoPaciente.Text)
    Else
        vlstrSentencia = vlstrSentencia & _
            " SELECT  HoMedico.intCveMedico, " & _
            " rtrim(ltrim(HoMedico.vchApellidoPaterno)) || ' ' || " & _
            " rtrim(ltrim(HoMedico.vchApellidoMaterno)) || ' ' || " & _
            " rtrim(LTrim(HoMedico.vchNombre)) As Nombre " & _
            " From RegistroExterno " & _
            " INNER JOIN HoMedico ON Homedico.intCveMedico = RegistroExterno.intMedico " & _
            " Where RegistroExterno.intNumCuenta = " & Trim(txtMovimientoPaciente.Text)
    End If

    Set rsMedicos = frsRegresaRs(vlstrSentencia)
    If rsMedicos.RecordCount <> 0 Then
        pLlenarCboRs cboMedicos, rsMedicos, 0, 1
        cboMedicos.ListIndex = 0
        cboMedicos_Click
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaMedicos"))
End Sub

Private Sub chkComisiones_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo NotificaError

    fraComisiones.Top = 1000
    fraComisiones.Visible = True
    frePaciente.Enabled = False
    fraRecibo.Enabled = False
    pCalculoTotales

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkComisiones_MouseUp"))
End Sub

Private Sub cmdAceptarComisiones_Click()
    On Error GoTo NotificaError
    
    Dim vldblcomisiones As Double
    pCalculoTotales
    If txtTotal <= 0 Then
        'Las comisiones generan honorarios negativos, corregir
        MsgBox SIHOMsg(928), vbOKOnly + vbInformation, "Mensaje"
    Else
        frePaciente.Enabled = True
        fraRecibo.Enabled = True
        txtReciboHonorario.Enabled = True
        chkComisiones.SetFocus
        fraComisiones.Visible = False
        If lstComisionesAsignadas.ListCount > 0 Then
            chkComisiones.Value = 1
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAceptarComisiones_Click"))
End Sub
Private Sub cmdBack_Click()
    On Error GoTo NotificaError

    If grdHonorarios.Row > 1 Then
        grdHonorarios.Row = grdHonorarios.Row - 1
    End If
    vllngNumHonorario = CLng(grdHonorarios.TextMatrix(grdHonorarios.Row, 1))
    pMuestraHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdBack_Click"))
End Sub
Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    
    Dim rsPvDetalleCorte As New ADODB.Recordset
    Dim rsHonorarioBorrar As New ADODB.Recordset
    Dim rsHonorarioEnPaqueteCobranza As New ADODB.Recordset 'Para saber si un honorario está en algun paquete de cobranza
    Dim rsFormasPago As New ADODB.Recordset
    Dim rsCortePoliza As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlblnError As Boolean
    Dim vllngPersonaGraba As Long
    Dim vldblcantidadpoliza As Long
    Dim vllngNumPoliza As Long
    Dim lngNumPolizaDetalle As Long
    Dim vllngCtaAcreedoraMedico As Long
    Dim lngContador As Long
    Dim lngTotal As Long
    
    vlblnError = False
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
    
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
    
        lngTotal = 0
    
        For lngContador = 1 To grdHonorarios.Rows - 1
        
            If grdHonorarios.TextMatrix(lngContador, 0) = "*" Then
        
                vlblnError = False
                
                vgstrParametrosSP = grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                Set rsHonorarioEnPaqueteCobranza = frsEjecuta_SP(vgstrParametrosSP, "SP_CCHONORARIOSENPAQCOBRANZA")
                
                'Se verifica que el honorario no se encuentre en un paquete de cobranza
                If rsHonorarioEnPaqueteCobranza.RecordCount <> 0 Then
                    MsgBox Left$(SIHOMsg(706), 34) & vbCrLf & "está dentro del paquete de cobranza " & rsHonorarioEnPaqueteCobranza!numero, vbOKOnly + vbInformation, "Mensaje"
                    vlblnError = True
                End If
            
                vgstrParametrosSP = grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo) & "|0"
                Set rsHonorarioBorrar = frsEjecuta_SP(vgstrParametrosSP, "SP_CCSELCUENTAPAGARHONORARIO")
                
                rsHonorarioBorrar.MoveFirst
                Do While Not rsHonorarioBorrar.EOF
                    If rsHonorarioBorrar!cheques <> 0 Then
                         'No se puede cancelar el honorario, ya se emitió cheque para el médico.
                        MsgBox SIHOMsg(706) & " " & grdHonorarios.TextMatrix(lngContador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        vlblnError = True
                    End If
                    If rsHonorarioBorrar!transferencias <> 0 Then
                         'No se puede cancelar el honorario, ya se realizó transferencia para el médico.
                        MsgBox "No se puede cancelar el honorario, ya se realizó transferencia para el médico. " & grdHonorarios.TextMatrix(lngContador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        vlblnError = True
                    End If
                    If Not vlblnError And rsHonorarioBorrar!autorizaciones <> 0 Then
                        'No se puede cancelar el honorario, ya se autorizó para pago.
                        MsgBox SIHOMsg(707) & " " & grdHonorarios.TextMatrix(lngContador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        vlblnError = True
                    End If
                    If Not vlblnError And rsHonorarioBorrar!pagos <> 0 Then
                        'No se puede cancelar el documento  el crédito tiene pagos registrados.
                        MsgBox SIHOMsg(368) & " " & grdHonorarios.TextMatrix(lngContador, cintColFilMedico), vbOKOnly + vbInformation, "Mensaje"
                        vlblnError = True
                    End If
                    rsHonorarioBorrar.MoveNext
                Loop
'---------------------------------------------------------***********************************************
                '| Valida que no se pueda cancelar un honorario que se generó automáticamente al realizar el pago de un crédito
                If Not vlblnError Then
                    vlstrSentencia = "Select Count(*) Co " & _
                                     "  From CpCuentaPagarMedico " & _
                                     "       Inner join CCPAGOHONORARIOMEDICOAUTOMATIC On CpCuentaPagarMedico.INTNUMCUENTAPAGAR = CCPAGOHONORARIOMEDICOAUTOMATIC.INTNUMCUENTAPAGAR " & _
                                     " Where CpCuentaPagarMedico.INTCONSECUTIVO = " & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    'Si el honorario fue generado automáticamente no se podrá cancelar
                    If rs.RecordCount <> 0 Then
                        If rs!CO > 0 Then
                            MsgBox "No se puede canacelar el honorario, fue generado automáticamente desde el proceso de pagos en cuentas por cobrar.", vbCritical, "Mensaje"
                            vlblnError = True
                        End If
                    End If
                End If
'---------------------------------------------------------***********************************************
                '| Valida que no se pueda cancelar un honorario que se generó automáticamente al realizar un pago de contado al generar una factura
                If Not vlblnError Then
                    vlstrSentencia = "Select Count(*) Co " & _
                                     "  From CpCuentaPagarMedico " & _
                                     "       Inner join PVFACTURAHONORARIOMEDAUTOMATIC On CpCuentaPagarMedico.INTNUMCUENTAPAGAR = PVFACTURAHONORARIOMEDAUTOMATIC.INTNUMCUENTAPAGAR " & _
                                     " Where CpCuentaPagarMedico.INTCONSECUTIVO = " & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    'Si el honorario fue generado automáticamente no se podrá cancelar
                    If rs.RecordCount <> 0 Then
                        If rs!CO > 0 Then
                            MsgBox "No se puede canacelar el honorario, fue generado automáticamente desde el proceso de facturación.", vbCritical, "Mensaje"
                            vlblnError = True
                        End If
                    End If
                End If
'---------------------------------------------------------***********************************************
                rsHonorarioBorrar.MoveFirst
                
                If Not vlblnError Then
                    
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    
                    If rsHonorarioBorrar!AfectoCorte = 1 Or rsHonorarioBorrar!credito = 1 Then
                        vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                        vllngCorteGrabando = 1
                        vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                        frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                        If vllngCorteGrabando <> 2 Then
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            'En este momento se está afectando el corte, espere un momento e intente de nuevo.
                            MsgBox SIHOMsg(779), vbOKOnly + vbInformation, "Mensaje"
                            Exit Sub
                        End If
                    End If
                    
                    'Borrar la cuenta por pagar, si es que existió
                    vlstrSentencia = "delete from CpCuentaPagarMedico where intconsecutivo = " & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                    pEjecutaSentencia vlstrSentencia
                    
                    'Cancelar el crédito, si existió:
                    vgstrParametrosSP = grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo) & "|" & "HO"
                    frsEjecuta_SP vgstrParametrosSP, "Sp_Ccupdcancelacredito"
                    
                    '--------------------------------------------------------------------------
                    'Si se trata de un honorario a credito y existe una cuenta por pagar
                    'se hace póliza para cancelar póliza que se genera al pagar por anticipado
                    '--------------------------------------------------------------------------
                    If vllngCuentaPagar > 0 And vlblnEsCredito Then
                        vllngCtaAcreedoraMedico = flngCuentaAcreedora(grdHonorarios.TextMatrix(lngContador, cintColFilCveMedico))
    
                        If vlblnesconvenio Then
                            vldblcantidadpoliza = CDbl(grdHonorarios.TextMatrix(lngContador, cintColFilSubtotal)) 'Honorario sin retención de ISR
                        Else
                            vldblcantidadpoliza = CDbl(grdHonorarios.TextMatrix(lngContador, cintColFilMonto))
                        End If
                        
                        vllngNumPoliza = flngInsertarPoliza(fdtmServerFecha, "D", "CANCELACION PAGO ANTICIPADO DE HONORARIO ", vllngPersonaGraba)
                        lngNumPolizaDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, vllngCtaHonorariosPagar, CDbl(Format(vldblcantidadpoliza, "############.00")), 0)
                        lngNumPolizaDetalle = flngInsertarPolizaDetalle(vllngNumPoliza, vllngCtaAcreedoraMedico, CDbl(Format(vldblcantidadpoliza, "############.00")), 1)
                    End If
                    
                    If rsHonorarioBorrar!AfectoCorte = 1 Or rsHonorarioBorrar!credito = 1 Then
                        vlstrSentencia = "select * from pvDetalleCorte where intConsecutivo = -1"
                        Set rsPvDetalleCorte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
                        vlstrSentencia = "Select * from PVDetalleCorte where chrFolioDocumento = " & "'" & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo) & "' and chrTipoDocumento='HO' "
                        Set rsFormasPago = frsRegresaRs(vlstrSentencia)
                        
                        If rsFormasPago.RecordCount <> 0 Then
                            rsFormasPago.MoveFirst
                            '--------------------------------------------
                            ' Cancelar el movimiento de la forma de pago
                            '--------------------------------------------
                            pCancelaMovimiento grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo), Trim(grdHonorarios.TextMatrix(lngContador, cintColFilRecibo)), "HO", rsFormasPago!intNumCorteDocumento, vllngNumeroCorte, vllngPersonaGraba
                            
                            Do While Not rsFormasPago.EOF
                                With rsPvDetalleCorte
                                    .AddNew
                                    !intnumcorte = vllngNumeroCorte
                                    !dtmFechahora = fdtmServerFecha + fdtmServerHora
                                    !chrFolioDocumento = grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                                    !chrTipoDocumento = "HO"
                                    !intFormaPago = rsFormasPago!intFormaPago
                                    !mnyCantidadPagada = rsFormasPago!mnyCantidadPagada * -1
                                    !MNYTIPOCAMBIO = rsFormasPago!MNYTIPOCAMBIO
                                    !intfoliocheque = rsFormasPago!intfoliocheque
                                    !intNumCorteDocumento = rsFormasPago!intNumCorteDocumento
                                    !INTCVEEMPLEADO = vllngPersonaGraba
                                    .Update
                                End With
                                rsFormasPago.MoveNext
                            Loop
                        End If
                    
                        vlstrSentencia = "select * from PvCortePoliza where chrFolioDocumento = '" & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo) & "' and chrTipoDocumento = 'HO'"
                        Set rsCortePoliza = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                        If rsCortePoliza.RecordCount > 0 Then
                            
                            rsCortePoliza.MoveFirst
                            Do While Not rsCortePoliza.EOF
                                pInsCortePoliza _
                                    vllngNumeroCorte, _
                                    grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo), _
                                    "HO", _
                                    rsCortePoliza!intNumCuenta, _
                                    rsCortePoliza!MNYCantidad * IIf(vllngNumeroCorte = rsCortePoliza!intnumcorte, -1, 1), _
                                    IIf(vllngNumeroCorte = rsCortePoliza!intnumcorte, rsCortePoliza!bitcargo, IIf(rsCortePoliza!bitcargo = 1, 0, 1))
                                rsCortePoliza.MoveNext
                            Loop
                        End If
                        
                        pLiberaCorte vllngNumeroCorte
                    End If
                    
                    'Poner estado de cancelado:
                    vgstrParametrosSP = grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo) & "|" & "C"
                    
                    frsEjecuta_SP vgstrParametrosSP, "Sp_Pvupdestadohonorario"
                    pEjecutaSentencia "update PVHonorario set vchEstatusPortal = 'CA' where intConsecutivo = " & grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                    
                    lngTotal = lngTotal + 1
                    
                    pGuardarLogTransaccion Me.Name, EnmGrabar, vllngPersonaGraba, "CANCELACION DE HONORARIOS", grdHonorarios.TextMatrix(lngContador, cintColFilConsecutivo)
                
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                End If
            End If
        Next lngContador
    
        If lngTotal <> 0 Then
            'La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        pLimpiaHonorarios
        cmdCargar.SetFocus
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdDelete"))
End Sub

Private Sub cmdEnd_Click()
    On Error GoTo NotificaError

    vllngNumHonorario = CLng(grdHonorarios.TextMatrix(grdHonorarios.Rows - 1, 1))
    pMuestraHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError

    SSTabHonorario.Tab = 1
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError

    If grdHonorarios.Row < grdHonorarios.Rows - 1 Then
        grdHonorarios.Row = grdHonorarios.Row + 1
    End If
    vllngNumHonorario = CLng(grdHonorarios.TextMatrix(grdHonorarios.Row, 1))
    pMuestraHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdNext_Click"))
End Sub

Private Sub cmdPrint_Click()
    Dim rsFacturas As New ADODB.Recordset
    Dim rsPvSelHonorario As New ADODB.Recordset
    Dim vlstrFacturas As String
    Dim vlIntCont As Integer

    Dim alstrParametros(2) As String

    On Error GoTo NotificaError

    If fraHonorarios.Visible Then     'Imprime los honorarios del grid
        For vlintContador = 1 To grdCargaHonorarios.Rows - 1
            vlstrFacturas = ""
            If grdCargaHonorarios.TextMatrix(vlintContador, 0) = "*" Then
                Set rsFacturas = frsRegresaRs("select DISTINCT * from pvfactura where intmovpaciente = " & txtMovimientoPaciente & " and chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and chrestatus <>'C'")
                If rsFacturas.RecordCount > 0 Then
                    For vlIntCont = 1 To rsFacturas.RecordCount
                        vlstrFacturas = Trim(vlstrFacturas) & IIf(vlstrFacturas = "", "", ", ") & rsFacturas!chrfoliofactura
                        rsFacturas.MoveNext
                    Next vlIntCont
                End If

                vgstrParametrosSP = grdCargaHonorarios.TextMatrix(vlintContador, cintColIdHonorario)

                Set rsPvSelHonorario = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelHonorario")
                If rsPvSelHonorario.RecordCount <> 0 Then
                    pInstanciaReporte vgrptReporte, "rptReciboHonorario.rpt"
                    vgrptReporte.DiscardSavedData
                    alstrParametros(0) = "Nombre Empresa" & ";" & Trim(vgstrNombreHospitalCH)
                    alstrParametros(1) = "vFacturas" & ";" & Trim(vlstrFacturas)
                    alstrParametros(2) = "vMuestraDetalle" & ";" & IIf(chkMuestraDetalle.Value, "1", "0")
                    pCargaParameterFields alstrParametros, vgrptReporte
                    pImprimeReporte vgrptReporte, rsPvSelHonorario, "P", "Honorarios"
                Else
                    MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
                End If

                rsPvSelHonorario.Close

                'Validación para que cuando sea cero no imprima el recibo
                If Val(Format(grdCargaHonorarios.TextMatrix(vlintContador, cintColComision), "############.##")) <> 0 Then
                    If MsgBox(SIHOMsg(520), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        vgstrParametrosSP = grdCargaHonorarios.TextMatrix(vlintContador, cintColIdHonorario)
                        Set rsPvSelComision = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelComision")
                        If rsPvSelComision.RecordCount > 0 Then
                            pInstanciaReporte vgrptReporte, "rptReciboComisiones.rpt"
                            vgrptReporte.DiscardSavedData
                            alstrParametros(0) = "Nombre Empresa" & ";" & Trim(vgstrNombreHospitalCH)
                            pCargaParameterFields alstrParametros, vgrptReporte
                            pImprimeReporte vgrptReporte, rsPvSelComision, "P", "Comisión honorarios"
                        Else
                            MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
                        End If
                        rsPvSelComision.Close
                    End If
                End If
            End If
        Next
    Else        'Aqui no hay grid, imprime el honorario de la consulta
        Set rsFacturas = frsRegresaRs("select * from pvfactura where intmovpaciente = " & txtMovimientoPaciente & " and chrtipopaciente = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and chrestatus <>'C'")
        If rsFacturas.RecordCount > 0 Then
            For vlIntCont = 1 To rsFacturas.RecordCount
                vlstrFacturas = Trim(vlstrFacturas) & IIf(vlstrFacturas = "", "", ", ") & rsFacturas!chrfoliofactura
                rsFacturas.MoveNext
            Next vlIntCont
        End If
        vgstrParametrosSP = vllngNumHonorario

        Set rsPvSelHonorario = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelHonorario")
        If rsPvSelHonorario.RecordCount <> 0 Then
            pInstanciaReporte vgrptReporte, "rptReciboHonorario.rpt"
            vgrptReporte.DiscardSavedData
            alstrParametros(0) = "Nombre Empresa" & ";" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = "vFacturas" & ";" & Trim(vlstrFacturas)
            alstrParametros(2) = "vMuestraDetalle" & ";" & IIf(chkMuestraDetalle.Value, "1", "0")
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsPvSelHonorario, "P", "Honorarios"
        Else
            MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
        End If

        rsPvSelHonorario.Close
        'Validación para que cuando sea cero no imprima el recibo
        If Val(Format(txtCantidadComision.Text, "############.##")) = 0 Then Exit Sub

        If MsgBox(SIHOMsg(520), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            vgstrParametrosSP = vllngNumHonorario
            Set rsPvSelComision = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelComision")
            If rsPvSelComision.RecordCount > 0 Then
                pInstanciaReporte vgrptReporte, "rptReciboComisiones.rpt"
                vgrptReporte.DiscardSavedData
                alstrParametros(0) = "Nombre Empresa" & ";" & Trim(vgstrNombreHospitalCH)
                pCargaParameterFields alstrParametros, vgrptReporte
                pImprimeReporte vgrptReporte, rsPvSelComision, "P", "Comisión honorarios"
            Else
                MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
            End If
            rsPvSelComision.Close
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim rsCreditoVigenteAsignado As New ADODB.Recordset
    Dim rsCuentasPuente As New ADODB.Recordset
    Dim vlstrTipoPacienteCredito As String
    Dim vllngCveClienteCredito As Long
    Dim vldbltotalpagar As Double
    Dim vlblnCredito As Boolean
    Dim vlintContador As Integer
    Dim vlblnAgregaCredito As Boolean
    
    vllngCtaCostoHonorarioFacturado = 0
    
    fblnDatosValidos = True
    
'For vlintContador = 1 To grdCargaHonorarios.Rows - 1
    If Trim(grdCargaHonorarios.TextMatrix(1, 1)) = "" Then
        MsgBox Mid(SIHOMsg(33), 1, 31) & ", no hay honorarios que guardar!", vbExclamation, "Mensaje"
        fblnDatosValidos = False
    End If
    
    If fblnDatosValidos Then
        If chkRequiereCFDI.Value = vbChecked Then
            If cboUsoCFDI.ListIndex = -1 Then
                MsgBox "Seleccione el uso del CFDI", vbExclamation, "Mensaje"
                cboUsoCFDI.SetFocus
                fblnDatosValidos = False
            End If
        End If
    End If
    
    vldbltotalpagar = lblTotalMonto 'grdCargaHonorarios.TextMatrix(vlintContador, cintColTotalPagar)
    If fblnDatosValidos Then
        vlstrSentencia = "select chrTipo from AdTipoPaciente where tnyCveTipoPaciente = " & Trim(Str(vgintTipoPaciente))
        Set rsTipoPaciente = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        vlstrTipoPacienteCredito = rsTipoPaciente!chrTipo
        If vlstrTipoPacienteCredito = "PA" Then
            If optTipoPaciente(0).Value Then
                vlstrTipoPacienteCredito = "PI"
            Else
                vlstrTipoPacienteCredito = "PE"
            End If
            vllngCveClienteCredito = CLng(txtMovimientoPaciente.Text) 'Igual que la clave del paciente
        ElseIf vlstrTipoPacienteCredito = "CO" Then
            vllngCveClienteCredito = vgintEmpresa
            'Cuando se trata de un paciente tipo convenio, la empresa absorve la retención del ISR y de RTP
            vldbltotalpagar = lblTotalMonto - IIf(chkPagoDirecto.Value, lblTotalRetencion, 0) - IIf(chkPagoDirecto.Value, lblTotalRTP, 0) '(lblTotalMonto - lblTotalRetencion)
        ElseIf vlstrTipoPacienteCredito <> "" Then
            'En la clave EXTRA va el número de Empleado o el de Médico
            vllngCveClienteCredito = vgintCveExtra
        Else
            vllngCveClienteCredito = 0
        End If

        rsTipoPaciente.Close
        
        Set rsCreditoVigenteAsignado = frsRegresaRs("select count(*) from CcCliente Inner Join Nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento " & _
                                     " where CcCliente.intNumReferencia=" & Str(vllngCveClienteCredito) & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
                                     " and CcCliente.chrTipoCliente = " & "'" & vlstrTipoPacienteCredito & "'" & " and CcCliente.bitActivo=1")
        If rsCreditoVigenteAsignado.Fields(0) = 0 And vlstrTipoPacienteCredito = "EM" Then
            vllngCveClienteCredito = Trim(txtMovimientoPaciente.Text)
            vlstrTipoPacienteCredito = IIf(optTipoPaciente(0).Value, "PI", IIf(optTipoPaciente(1).Value, "PE", "PP"))
            Set rsCreditoVigenteAsignado = frsRegresaRs("select count(*) from CcCliente Inner Join Nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento " & _
                                     " where CcCliente.intNumReferencia=" & Str(vllngCveClienteCredito) & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
                                     " and CcCliente.chrTipoCliente = '" & vlstrTipoPacienteCredito & "' and CcCliente.bitActivo=1")
        End If
        
        If lblnPacienteConvenio Then
            vlstrSentencia = "select * from cccliente inner join ccempresa on cccliente.INTNUMREFERENCIA = ccempresa.INTCVEEMPRESA where ccempresa.INTCVEEMPRESA = " & Trim(Str(vllngCveClienteCredito))
        Else
           vlstrSentencia = "select * from REGISTROEXTERNO INNER JOIN cccliente on REGISTROEXTERNO.intNumCuenta = cccliente.INTNUMREFERENCIA WHERE REGISTROEXTERNO.intNumCuenta = " & txtMovimientoPaciente.Text
        End If

        Set rsDatosCliente = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        
        If rsDatosCliente.RecordCount > 0 Then vlblnAgregaCredito = True
        
        If chkHonorarioFacturado.Value = 0 Then
            If Not lblnPacienteConvenio Then
                fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), Val(Format(vldbltotalpagar, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)), True, vldblTipoCambio, IIf(chkDolares.Value = 1, False, True), vllngCveClienteCredito, vlstrTipoPacienteCredito, Trim(txtRFC.Text), False, False, True, "Honorarios")
            Else
                If chkPagoDirecto.Value = 1 Then
                    fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), Val(Format(vldbltotalpagar, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)), True, vldblTipoCambio, IIf(chkDolares.Value = 1, False, True), vllngCveClienteCredito, vlstrTipoPacienteCredito, Trim(txtRFC.Text), False, False, True, "Honorarios")
                Else
                    If vlblnAgregaCredito Then
                        fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), Val(Format(vldbltotalpagar, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)), True, vldblTipoCambio, IIf(chkDolares.Value = 1, False, True), vllngCveClienteCredito, vlstrTipoPacienteCredito, Trim(txtRFC.Text), False, False, True, "Honorarios")
                    Else
                        fblnDatosValidos = fblnFormasPagoPos(aFormasPago(), Val(Format(vldbltotalpagar, "############.00")) * (IIf(chkDolares.Value = 1, vldblTipoCambio, 1)), True, vldblTipoCambio, False, vllngCveClienteCredito, vlstrTipoPacienteCredito, Trim(txtRFC.Text), False, False, True, "Honorarios")
                    End If
                End If
                
            End If
        Else
            vlstrSentencia = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'INTCTACOSTOHONORARIOFACTURADO' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
            Set rs = frsRegresaRs(vlstrSentencia)
            If rs.RecordCount <> 0 Then
                vllngCtaCostoHonorarioFacturado = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
            Else
                vllngCtaCostoHonorarioFacturado = 0
            End If
            
            If vllngCtaCostoHonorarioFacturado = 0 Then
                'No se encuentra configurada la cuenta de ingresos del concepto de facturación para los honorarios médicos, favor de verificar.
                MsgBox SIHOMsg(1450), vbOKOnly + vbInformation, "Mensaje"
                fblnDatosValidos = False
            Else
                If Not fblnCuentaAfectable(fstrCuentaContable(vllngCtaCostoHonorarioFacturado), vgintClaveEmpresaContable) Then
                    'La cuenta contable de ingresos del concepto de facturación para los honorarios médicos no acepta movimientos, favor de verificar.
                    MsgBox SIHOMsg(1451), vbOKOnly + vbInformation, "Mensaje"
                    fblnDatosValidos = False
                Else
                    fblnDatosValidos = True
                End If
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        If chkHonorarioFacturado.Value = 0 Then
            If aFormasPago(0).vlbolEsCredito = True Then
                vlblnCredito = True
            Else
                vlblnCredito = False
            End If
        Else
            vlblnCredito = False
        End If
    End If
    
    If fblnDatosValidos And vlblnCredito Then
        vlstrSentencia = "select * from ccCliente " & _
                        " Inner Join nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento " & _
                        " where ccCliente.intNumReferencia = " & Trim(Str(vllngCveClienteCredito)) & _
                        " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
                        " and ccCliente.chrTipoCliente = '" & vlstrTipoPacienteCredito & "'"

        Set rsDatosCliente = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsDatosCliente.RecordCount = 0 Then
            ' Se detecto un error en la información del cliente
            MsgBox SIHOMsg(367), vbCritical, "Mensaje"
            fblnDatosValidos = False
        End If
    End If
    
    If fblnDatosValidos And vlblnCredito And vllngCtaHonorariosCobrar = 0 Then
        'No se encuentra registrada la cuenta contable para honorarios por cobrar.'
        MsgBox SIHOMsg(703), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
    End If
    
    If fblnDatosValidos And vlblnCredito And vllngCtaHonorariosPagar = 0 Then
        'No se encuentra registrada la cuenta contable para honorarios por pagar.'
        MsgBox SIHOMsg(704), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
    End If
    'Validacion de cuentas afectables en caso de ser credito
    
    If fblnDatosValidos And vlblnCredito And Not fblnCuentaAfectable(fstrCuentaContable(vllngCtaHonorariosCobrar), vgintClaveEmpresaContable) Then
        'La cuenta contable asignada para honorarios por pagar no acepta movimientos.'
         MsgBox SIHOMsg(1237), vbExclamation, "Mensaje"
         fblnDatosValidos = False
    End If
    If fblnDatosValidos And vlblnCredito And Not fblnCuentaAfectable(fstrCuentaContable(vllngCtaHonorariosPagar), vgintClaveEmpresaContable) Then
        'La cuenta contable asignada para honorarios por cobrar no acepta movimientos.'
        MsgBox SIHOMsg(1238), vbExclamation, "Mensaje"
        fblnDatosValidos = False
    End If
    
    If fblnDatosValidos And cgstrModulo = "CC" And vlblnCredito Then
        If lngDeptoCliente <> vgintNumeroDepartamento Then
            'El crédito de la cuenta pertenece a otro departamento.'
            MsgBox SIHOMsg(723), vbOKOnly + vbExclamation, "Mensaje"
            fblnDatosValidos = False
        End If
    End If
'Next vlintContador
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function

Private Sub cmdSelecciona_Click(Index As Integer)
    On Error GoTo NotificaError

    If Index = 0 Then
        pSeleccionaLista lstComisiones.ListIndex, lstComisiones, lstComisionesAsignadas, cmdSelecciona(0), cmdSelecciona(1)
    Else
        pSeleccionaLista lstComisionesAsignadas.ListIndex, lstComisionesAsignadas, lstComisiones, cmdSelecciona(1), cmdSelecciona(0)
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSelecciona_Click"))
End Sub

Private Sub cmdTodosMedicos_Click()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    'Todos los Médicos
    vlstrSentencia = "SELECT  HoMedico.intCveMedico, " & _
        " rtrim(ltrim(HoMedico.vchApellidoPaterno)) || ' ' || " & _
        " rtrim(ltrim(HoMedico.vchApellidoMaterno)) || ' ' || " & _
        " rtrim(LTrim(HoMedico.vchNombre)) As Nombre " & _
        " From Homedico " & _
        " where bitEstaActivo = 1 "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboMedicos, rs, 0, 1
    If rs.RecordCount > 0 Then
        cboMedicos.ListIndex = 0
    End If
    rs.Close
    cboMedicos.SetFocus

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTodosMedicos_Click"))
End Sub

Private Sub cmdTop_Click()
    On Error GoTo NotificaError

    vllngNumHonorario = CLng(grdHonorarios.TextMatrix(1, 1))
    pMuestraHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdTop_Click"))
End Sub
Private Sub pMuestraHonorario()
    On Error GoTo NotificaError

    Dim vllngContador As Long
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsHonorario As New ADODB.Recordset
    Dim rscuentapagar As New ADODB.Recordset
    Dim vldblMontoHonorario As Double
    Dim vldblRetencion As Double
    Dim vldblRetencionRTP As Double
    Dim vldblPagosCredito As Double
    Dim vldblComision As Double
    Dim vldblIvaComision As Double
    Dim vlintCredito As Integer
    Dim rsDatosPaciente As New ADODB.Recordset
    Dim vlblnHabilitaPago As Boolean
    Dim rsFacturaCxPProveedor As New ADODB.Recordset
  
    vllngCuentaPagar = 0
    Set rsHonorario = frsEjecuta_SP(CStr(vllngNumHonorario), "Sp_Pvselhonorario")
    If rsHonorario.RecordCount <> 0 Then
        vgstrParametrosSP = CStr(vllngNumHonorario)
        Set rscuentapagar = frsEjecuta_SP(vgstrParametrosSP, "sp_ccselexistecuentapagarmedic")
        If rscuentapagar.RecordCount > 0 Then
            vllngCuentaPagar = 1
        Else
            vllngCuentaPagar = 0
        End If
        rscuentapagar.Close
        
        vlstrSentencia = "SELECT COPROVEEDOR.VCHNOMBRE, CPCUENTAPAGAR.DTMFECHARECEPCION, CPCUENTAPAGAR.DTMFECHADOCUMENTO, CPCUENTAPAGAR.VCHNUMEROFACTURA " & _
                                    "FROM PVHONORARIO INNER JOIN CPCUENTAPAGARMEDICO ON  PVHONORARIO.INTCONSECUTIVO = CPCUENTAPAGARMEDICO.INTCONSECUTIVO " & _
                                    "INNER JOIN CPHONORARIOCONTRARECIBO ON CPCUENTAPAGARMEDICO.INTNUMCUENTAPAGAR = CPHONORARIOCONTRARECIBO.INTCONSECUTIVOHONORARIO " & _
                                    "INNER JOIN CPCUENTAPAGAR ON CPHONORARIOCONTRARECIBO.INTNUMEROCXP = CPCUENTAPAGAR.INTNUMEROCXP " & _
                                    "INNER JOIN COPROVEEDOR ON CPCUENTAPAGAR.INTCVEPROVEEDOR = COPROVEEDOR.INTCVEPROVEEDOR  WHERE  PVHONORARIO.INTCONSECUTIVO = " & vllngNumHonorario
        Set rsFacturaCxPProveedor = frsRegresaRs(vlstrSentencia)
        
        vlblnConsulta = True
        vlblnPacienteSeleccionado = True
        
        vlstrAdTipoPaciente = rsHonorario!ChrTipoPac
        vlintMoneda = rsHonorario!EstadoMoneda
        vldblTipoCambioHonorario = IIf(IsNull(rsHonorario!TipoCambio), 0, rsHonorario!TipoCambio)
            
        vlintCredito = IIf(IsNull(rsHonorario!EstadoCredito), 0, rsHonorario!EstadoCredito)
        vlblnEsCredito = IIf(rsHonorario!EstadoCredito = 1, True, False)
        vldblMontoHonorario = IIf(IsNull(rsHonorario!MontoHonorario), 0, rsHonorario!MontoHonorario)
        vldblRetencion = IIf(IsNull(rsHonorario!MontoRetencion), 0, rsHonorario!MontoRetencion)
        vldblRetencionRTP = IIf(IsNull(rsHonorario!MontoRetencionRTP), 0, rsHonorario!MontoRetencionRTP)
        
        vldblPagosCredito = IIf(IsNull(rsHonorario!MontoPagos), 0, rsHonorario!MontoPagos)
        vldblComision = IIf(IsNull(rsHonorario!Comision), 0, rsHonorario!Comision)
        vldblIvaComision = IIf(IsNull(rsHonorario!IvaComision), 0, rsHonorario!IvaComision)
        
        txtMovimientoPaciente.Text = Str(rsHonorario!NumeroCuenta)
        optTipoPaciente(0).Value = rsHonorario!InternoExterno = "I"
        optTipoPaciente(1).Value = rsHonorario!InternoExterno = "E"
            
        txtPaciente.Text = rsHonorario!Paciente
        
        If IIf(IsNull(rsHonorario!convenio), "", Trim(rsHonorario!convenio)) = "" Then
            txtTipoPaciente.Text = IIf(IsNull(rsHonorario!TipoPaciente), "", rsHonorario!TipoPaciente)
        Else
            txtTipoPaciente.Text = IIf(IsNull(rsHonorario!convenio), "", rsHonorario!convenio)
        End If
        
        lblTipoEmpresa.Caption = IIf(Trim(rsHonorario!convenio) = "", "Tipo de paciente", "Empresa")
        If IsNull(rsHonorario!convenio) Or rsHonorario!convenio = 0 Or rsHonorario!convenio = " " Then
            vlblnesconvenio = False
        Else
            vlblnesconvenio = True
        End If
        
        mskFechaRegistro.Mask = ""
        mskFechaRegistro.Text = FormatDateTime(rsHonorario!FECHAREGISTRO, vbShortDate)
        mskFechaRegistro.Mask = "##/##/####"
        
        lblEstadoHonorario.Visible = True
        lblEstadoHonorario.Caption = rsHonorario!Estado
        
        cboMedicos.Clear
        cboMedicos.AddItem rsHonorario!Medico, 0
        cboMedicos.ItemData(cboMedicos.newIndex) = rsHonorario!IdMedico
        cboMedicos.ListIndex = 0
        
        mskFechaAtencion.Text = rsHonorario!FechaInicioAtencion
        mskFechaAtencionFin.Text = Format(rsHonorario!FechaFinAtencion, "dd/mm/yyyy")
        
        txtReciboHonorario.Text = IIf(IsNull(rsHonorario!NumeroRecibo), " ", rsHonorario!NumeroRecibo)
        
        txtConcepto.Text = IIf(IsNull(rsHonorario!Concepto), " ", rsHonorario!Concepto)
        chkPagoDirecto.Value = IIf(rsHonorario!bitPagoDirecto <> 0, vbChecked, vbUnchecked)
        
        If rsHonorario!BITHONORARIOFACTURADO Then
            txtReciboNombre.Text = vgstrNombreHospitalCH
            txtDireccion.Text = Trim(vgstrDireccionCH) & IIf(Trim(vgstrColoniaCH) = "", "", "   Col. " & Trim(vgstrColoniaCH)) & IIf(Trim(vgstrCodPostalCH) = "", "", "   CP " & Trim(vgstrCodPostalCH)) & IIf(Trim(vgstrCiudadCH) = "", "", " " & vgstrCiudadCH) & IIf(Trim(vgstrEstadoCH) = "", "", ", " & Trim(vgstrEstadoCH))
            txtRFC.Text = vgstrRfCCH
            txtEmail.Text = vgstrEmailCH
        Else
            txtReciboNombre.Text = IIf(IsNull(rsHonorario!NombreRecibo), " ", rsHonorario!NombreRecibo)
            txtDireccion.Text = IIf(IsNull(rsHonorario!Direccion), " ", rsHonorario!Direccion)
            txtRFC.Text = IIf(IsNull(rsHonorario!RFC), " ", Trim(Replace(Replace(Replace(rsHonorario!RFC, "-", ""), "_", ""), " ", "")))
            txtEmail.Text = IIf(IsNull(rsHonorario!vchCorreo), "", rsHonorario!vchCorreo)
        End If
        
        txtMontoHonorario.Text = FormatCurrency(Str(vldblMontoHonorario), 2)
        chkRetencion.Value = IIf(vldblRetencion = 0, 0, 1)
        chkRetencionRTP.Value = IIf(vldblRetencionRTP = 0, 0, 1)
        
        cboTarifa.ListIndex = -1
        If chkRetencion.Value = 1 Then
            
            vllngContador = 0
            Do While vllngContador <= cboTarifa.ListCount - 1
                If cboTarifa.ItemData(vllngContador) = rsHonorario!IdTarifa Then
                    cboTarifa.ListIndex = vllngContador
                End If
                vllngContador = vllngContador + 1
            Loop
            
            If cboTarifa.ListIndex = -1 Then
            
                lblnRecargarTarifas = True
            
                cboTarifa.AddItem rsHonorario!Tarifa, 0
                cboTarifa.ItemData(cboTarifa.newIndex) = rsHonorario!IdTarifa
                cboTarifa.ListIndex = 0
            End If
        End If
        
        txtRetencion.Text = FormatCurrency(Str(vldblRetencion), 2)
        txtRetencionRTP.Text = FormatCurrency(Str(vldblRetencionRTP), 2)
        dblTotalRetencionDetalle = Val(Format(txtRetencion.Text, "############.##")) + Val(Format(txtRetencionRTP.Text, "############.##"))
        
        txtNetoPagar.Text = FormatCurrency(vldblMontoHonorario - dblTotalRetencionDetalle, 2)
        chkComisiones.Value = IIf(vldblComision = 0, 0, 1)
        txtCantidadComision.Text = FormatCurrency(vldblComision, 2)
        txtIVAComision.Text = FormatCurrency(vldblIvaComision, 2)
        txtPagoCredito.Text = FormatCurrency(vldblPagosCredito, 2)
        txtTotal.Text = FormatCurrency(vldblMontoHonorario - dblTotalRetencionDetalle - vldblComision - vldblIvaComision, 2)
        
        chkHonorarioFacturado.Value = IIf(rsHonorario!BITHONORARIOFACTURADO <> 0, vbChecked, vbUnchecked)
        
        If IsNull(rsHonorario!intCveUsoCFDI) Then
            cboUsoCFDI.ListIndex = -1
        Else
            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, rsHonorario!intCveUsoCFDI)
        End If
        chkRequiereCFDI.Value = IIf(IsNull(rsHonorario!bitRequiereRecibo), vbUnchecked, IIf(rsHonorario!bitRequiereRecibo = 0, vbUnchecked, vbChecked))
        txtObservaciones.Text = IIf(IsNull(rsHonorario!vchObservaciones), "", rsHonorario!vchObservaciones)
        
        
        lblFechaRegistro.Visible = True
        mskFechaRegistro.Visible = True
        lblTipoPago.Visible = True
        optTipoPaciente(2).Visible = False

        If rsFacturaCxPProveedor.RecordCount > 0 Then
            With grdHonoCxPProv
                .Clear
                .Cols = 5
                .Rows = 2
                .FixedCols = 1
                .FixedRows = 1
                .FormatString = "|Proveedor/Acreedor|Folio factura|Fecha documento|Fecha recepción"
                .ColWidth(0) = 80
                .ColWidth(1) = 5500     'Nombre proveedor
                .ColWidth(2) = 3300      'Folio factura
                .ColWidth(3) = 1450    'Fecha documento
                .ColWidth(4) = 1450    'Fecha recepcion
                .ColAlignmentFixed(1) = flexAlignCenterCenter
                .ColAlignmentFixed(2) = flexAlignCenterCenter
                .ColAlignmentFixed(3) = flexAlignCenterCenter
                .ColAlignmentFixed(4) = flexAlignCenterCenter
                .ColAlignment(1) = flexAlignLeftCenter
                .ColAlignment(2) = flexAlignLeftCenter
                .ColAlignment(3) = flexAlignLeftCenter
                .ColAlignment(4) = flexAlignLeftCenter
            
                Do While Not rsFacturaCxPProveedor.EOF
                    .TextMatrix(.Rows - 1, 1) = rsFacturaCxPProveedor!vchNombre
                    .TextMatrix(.Rows - 1, 2) = rsFacturaCxPProveedor!VCHNUMEROFACTURA
                    .TextMatrix(.Rows - 1, 3) = Format(rsFacturaCxPProveedor!DTMFECHADOCUMENTO, "dd/mmm/yyyy")
                    .TextMatrix(.Rows - 1, 4) = Format(rsFacturaCxPProveedor!DTMFECHARECEPCION, "dd/mmm/yyyy")
                    rsFacturaCxPProveedor.MoveNext
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                Loop
                .Rows = .Rows - 1
            End With
            fraCuentaPagar.Visible = True
        Else
            fraCuentaPagar.Visible = False
        End If
        
        fraHonorarios.Visible = False
        
        If rsHonorario!Estado = "PAGO PENDIENTE AL MEDICO" And rsHonorario!BITHONORARIOFACTURADO = 0 Then vlblnHabilitaPago = True
        SSTabHonorario.Tab = 0
        If lblEstadoHonorario = "CANCELADO" Then
            pHabilita 1, 1, 1, 1, 1, 0, 0, 0
        Else
            pHabilita 1, 1, 1, 1, 1, 0, 1, IIf(vlblnHabilitaPago, 1, 0)
        End If
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pMuestraHonorario"))
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Function fblnEntraCorte() As Boolean
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rsIncluyeHonorarioCorte As New ADODB.Recordset

    fblnEntraCorte = False
    
    vlstrSentencia = "select bitIncluyeHonorarioCorte from PvParametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsIncluyeHonorarioCorte = frsRegresaRs(vlstrSentencia)
    If rsIncluyeHonorarioCorte.RecordCount <> 0 Then
        fblnEntraCorte = IIf(IsNull(rsIncluyeHonorarioCorte.Fields(0)), False, rsIncluyeHonorarioCorte.Fields(0) = 1)
    End If
        
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnEntraCorte"))
End Function

Private Sub Form_Load()
    On Error GoTo NotificaError
        
    vgstrNombreForm = Me.Name
        
    Me.Icon = frmMenuPrincipal.Icon

    ldtmFecha = fdtmServerFecha

    vlblnEntrando = True
    
    vldblTipoCambio = fdblTipoCambio(fdtmServerFecha, "V")

    vlblnEntraCorte = fblnEntraCorte()

    fraReciboHonorario.BorderStyle = 0
    chkMuestraDetalle = IIf(cgstrModulo = "PV", 0, 1)
    
    lblnRecargarTarifas = False
    chkHonorarioFacturado.Enabled = False
    pCargaTarifas
    pCargaUsosCFDI
    
    pLimpia
    pHabilita 0, 0, 1, 0, 0, 0, 0, 0
    vlintInterno = fintEsInterno(vglngNumeroLogin, enmTipoProceso.Honorarios)
    optTipoPaciente(0).Value = vlintInterno = 1
    optTipoPaciente(1).Value = vlintInterno = 2
    optTipoPaciente(2).Value = vlintInterno = 3
    
    vlTipoPacienteCatalogo = ""
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pCargaTarifas()
    On Error GoTo NotificaError
    Dim intcontador As Integer
    Dim rs As New ADODB.Recordset
    
    cboTarifa.Clear
    
    Set rs = frsEjecuta_SP("-1|1", "SP_CNSELTARIFAISR")
    
    intcontador = 0
    Do While Not rs.EOF
        ReDim Preserve arrTarifas(intcontador)

        cboTarifa.AddItem rs!Descripcion
        cboTarifa.ItemData(cboTarifa.newIndex) = rs!IdTarifa
        
        arrTarifas(intcontador).lngIdTarifa = rs!IdTarifa
        arrTarifas(intcontador).dblPorcentaje = rs!Porcentaje
    
        intcontador = intcontador + 1
    
        rs.MoveNext
    Loop
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaImpuestos"))
End Sub

Private Sub mskFechaAtencion_GotFocus()
    On Error GoTo NotificaError

    pSelMkTexto mskFechaAtencion
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAtencion_GotFocus"))
End Sub

Private Sub mskFechaAtencion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        mskFechaAtencionFin.SetFocus
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaAtencion_KeyPress"))
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
        
    If Not vlblnConsulta Then
        pEnfocaTextBox txtMovimientoPaciente
    End If
    
    If optTipoPaciente(0).Value Or optTipoPaciente(1).Value Then
        lblCuentaPaciente.Caption = "Número de cuenta"
    Else
        lblCuentaPaciente.Caption = "Número de paciente"
    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":OptTipoPaciente_Click"))
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

Private Sub SSTabHonorario_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTabHonorario.Tab = 1 Then
        pLimpiaHonorarios
        txtFilCuenta.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":SSTabHonorario_Click"))
End Sub

Private Sub txtConcepto_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtConcepto
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtConcepto_GotFocus"))
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        KeyAscii = 0
        'cboUsoCFDI.SetFocus
        SendKeys vbTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtConcepto_KeyPress"))
End Sub

Private Sub txtDireccion_GotFocus()
    On Error GoTo NotificaError

    'pSelTextBox txtDireccion
    txtDireccion.SelStart = Len(txtDireccion)

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDireccion_GotFocus"))
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        KeyAscii = 0
        txtRFC.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtDireccion_KeyPress"))
End Sub

Private Sub txtMontoHonorario_GotFocus()
    On Error GoTo NotificaError

    txtMontoHonorario.Text = Format(txtMontoHonorario.Text, "############.##")
    pSelTextBox txtMontoHonorario

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_GotFocus"))
End Sub

Private Sub txtMontoHonorario_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
 
    If KeyCode = vbKeyReturn Then
        pCalculoTotales
        SendKeys vbTab
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_KeyDown"))
End Sub

Private Sub txtMontoHonorario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

If Not fblnFormatoCantidad(txtMontoHonorario, KeyAscii, 2) Then
       KeyAscii = 7
    Else
            If txtMontoHonorario.Text <> "." And txtMontoHonorario.Text <> "" Then
            
                If CDbl(txtMontoHonorario.Text) > 0 Then
    
                    chkComisiones.Enabled = True
    
                Else: chkComisiones.Enabled = False

                End If
            End If
            
        pCalculoTotales
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMontoHonorario_KeyPress"))
End Sub

Private Sub txtMovimientoPaciente_GotFocus()

    On Error GoTo NotificaError

    pSelTextBox txtMovimientoPaciente

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_GotFocus"))
End Sub

Private Sub pLimpia(Optional vlguardar As String)
    On Error GoTo NotificaError
    
    vlTipoPacienteCatalogo = ""
    
    lblnPacienteConvenio = False
    
    lstrRFCConvPers = ""
    lstrDireccionConvPers = ""
    lstrNombreConvPers = ""
    lstrCorreoElectronico = ""
    chkPagoDirecto.Value = 0
    
    vlintlimpiar = 0
    vlblnConsulta = False
    vlblnPacienteSeleccionado = False
    
    chkHonorarioFacturado.Value = 0
    'chkHonorarioFacturado.Enabled = True
    
    '*-*-*-*-Consulta de honorarios-*-*-*-*-*
    txtFilCuenta.Text = ""
    txtFilPaciente.Text = ""
    pCargaFilMedico
    grdHonorarios.Clear
    grdHonorarios.Rows = 0
    grdHonorarios.Cols = 0
    '*-*-*-*-Consulta de honorarios-*-*-*-*-*
    
    vlintMoneda = 0
    vldblTipoCambioHonorario = 0
    
    lblFechaRegistro.Visible = False
    mskFechaRegistro.Visible = False
    lblTipoPago.Visible = False
    lblEstadoHonorario.Visible = False
    
    vldblPagoCuenta = 0

    SSTabHonorario.Tab = 0

    fraHonorarios.Visible = True

    txtMovimientoPaciente.Text = ""
    
    optTipoPaciente(2).Visible = True
    
    optFilTipoPaciente(0).Value = vlintInterno = 1
    optFilTipoPaciente(1).Value = vlintInterno = 2 Or vlintInterno = 3

    txtPaciente.Text = ""
    txtTipoPaciente.Text = ""
    lblTipoPago.Visible = False
    lblEstadoHonorario.Visible = False

    fraRecibo.Enabled = False
    cboMedicos.ListIndex = -1

    mskFechaAtencion.Mask = ""
    mskFechaAtencion.Text = ""
    mskFechaAtencion.Mask = "##/##/####"

    mskFechaAtencionFin.Mask = ""
    mskFechaAtencionFin.Text = ""
    mskFechaAtencionFin.Mask = "##/##/####"

    fraReciboHonorario.Enabled = True
    txtReciboHonorario.Text = ""

    txtConcepto.Text = ""
    txtReciboNombre.Text = ""
    txtDireccion.Text = ""
    txtRFC.Text = ""


    chkRetencion.Value = 0
    cboTarifa.ListIndex = -1
    chkComisiones.Value = 0
    chkRetencionRTP.Value = 0
    
    txtMontoHonorario.Text = ""
    txtRetencion.Text = ""
    txtRetencionRTP.Text = ""
    txtNetoPagar.Text = ""

    txtCantidadComision.Text = ""
    txtIVAComision.Text = ""
    txtPagoCredito.Text = ""
    txtTotal.Text = ""

    chkDolares.Value = 0
    cboUsoCFDI.ListIndex = -1
    txtEmail.Text = ""
    txtObservaciones.Text = ""
    txtAdjuntar.Text = ""
        
    chkRequiereCFDI.Value = vbUnchecked
    fraCuentaPagar.Visible = False
    pConfigura

    grdComisiones.Clear
    pCalculaTotales
    
    chkComisiones.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub pHabilita(vlintTop As Integer, vlintBack As Integer, vlintLocate As Integer, vlintNext As Integer, vlintEnd As Integer, vlintSave As Integer, vlintPrint As Integer, vlintEfectivo As Integer)
    On Error GoTo NotificaError

    cmdTop.Enabled = vlintTop = 1
    cmdBack.Enabled = vlintBack = 1
    cmdLocate.Enabled = vlintLocate = 1
    cmdNext.Enabled = vlintNext = 1
    cmdEnd.Enabled = vlintEnd = 1
    cmdSave.Enabled = vlintSave = 1
    cmdPrint.Enabled = vlintPrint = 1
    cmdPagoEfectivo.Enabled = vlintEfectivo = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilita"))
End Sub

Private Sub pCargaComisiones()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rsComisiones As New ADODB.Recordset
    Dim vlintContador As Integer

    pLimpiaComisiones '(los dos list)

    chkComisiones.Value = 0

    vlstrSentencia = "Select smiCveComision Clave, chrDescripcion Nombre, bitAsignada Asignada " & _
                        " From pvComision " & _
                        " Where bitActivo = 1 "
    Set rsComisiones = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)

    If rsComisiones.RecordCount > 0 Then
        Do While Not rsComisiones.EOF

            lstComisiones.AddItem rsComisiones!Nombre
            lstComisiones.ItemData(lstComisiones.newIndex) = rsComisiones!clave
            lstComisiones.Enabled = True
            cmdSelecciona(0).Enabled = True

            rsComisiones.MoveNext
        Loop
    End If

    If lstComisiones.ListCount <> 0 Then
        lstComisiones.ListIndex = 0
    End If

    rsComisiones.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaComisiones"))
End Sub

Private Sub pLimpiaComisiones()
    On Error GoTo NotificaError

    lstComisiones.Clear
    lstComisionesAsignadas.Clear
    lstComisiones.Enabled = False
    lstComisionesAsignadas.Enabled = False

    cmdSelecciona(0).Enabled = False
    cmdSelecciona(1).Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaComisiones"))
End Sub

Private Function fblnDatosPaciente(vllngxMovimiento As Long, lngNumPaciente As Long, vlstrxTipoPaciente As String) As Boolean
    Dim rsHono As ADODB.Recordset
    Dim rsUsoCFDIFact As ADODB.Recordset
    On Error GoTo NotificaError

    fblnDatosPaciente = False

    If Not vlblnConsulta Then
        vgstrParametrosSP = CStr(vllngxMovimiento) & "|" & CStr(lngNumPaciente) & "|" & vlstrxTipoPaciente & "|" & vgintClaveEmpresaContable
        
        Set rsPvSelDatosPaciente = frsEjecuta_SP(vgstrParametrosSP, "sp_pvSelDatosPaciente")
        
        If rsPvSelDatosPaciente.RecordCount <> 0 Then
            fblnDatosPaciente = True
            txtPaciente.Text = rsPvSelDatosPaciente!Nombre
            
            txtTipoPaciente.Text = IIf(IsNull(rsPvSelDatosPaciente!empresa), rsPvSelDatosPaciente!tipo, rsPvSelDatosPaciente!empresa)
            
            vlTipoPacienteCatalogo = rsPvSelDatosPaciente!TipoPacienteCatalogo
            
            lblTipoEmpresa.Caption = IIf(IsNull(rsPvSelDatosPaciente!empresa), "Tipo de paciente", "Empresa")
            If Not IsNull(rsPvSelDatosPaciente!Ingreso) Then
                mskFechaAtencion.Text = Format(rsPvSelDatosPaciente!Ingreso, "dd/mm/yyyy")
                mskFechaAtencionFin.Text = Format(rsPvSelDatosPaciente!Ingreso, "dd/mm/yyyy")
            End If
            vgintTipoPaciente = rsPvSelDatosPaciente!tnyCveTipoPaciente
            vgintEmpresa = rsPvSelDatosPaciente!intcveempresa
            llngTipoConvenio = rsPvSelDatosPaciente!TipoConvenio
            vgintCveExtra = rsPvSelDatosPaciente!intCveExtra
            lngDeptoCliente = IIf(IsNull(rsPvSelDatosPaciente!NumDeptoCliente), 0, rsPvSelDatosPaciente!NumDeptoCliente)
            lstrAfiliacion = IIf(IsNull(rsPvSelDatosPaciente!Afiliacion), " ", rsPvSelDatosPaciente!Afiliacion)
        
            If rsPvSelDatosPaciente!TipoPacienteCatalogo = "CONVENIO" Then
                lblnPacienteConvenio = True
            Else
                lblnPacienteConvenio = False
            End If
            
            chkRequiereCFDI.Value = 1
            If lblnPacienteConvenio Then
                vlstrcveempresa = rsPvSelDatosPaciente!intcveempresa
                lstrNombreConvPers = rsPvSelDatosPaciente!empresa
                lstrDireccionConvPers = rsPvSelDatosPaciente!DireccionEmpresa
                lstrRFCConvPers = Trim(Replace(Replace(Replace(rsPvSelDatosPaciente!RFCEmpresa, "-", ""), "_", ""), " ", ""))
                lstrCorreoElectronico = IIf(IsNull(rsPvSelDatosPaciente!correoempresa), " ", rsPvSelDatosPaciente!correoempresa)
                
                Set rsHono = frsRegresaRs("select * from CCEmpresaHonorarios where intCveEmpresa = " & vlstrcveempresa, adLockReadOnly, adOpenForwardOnly)
                If rsHono.RecordCount > 0 Then
                    chkPagoDirecto.Value = IIf(IsNull(rsHono!bitPagoDirecto), vbChecked, IIf(rsHono!bitPagoDirecto = 0, vbUnchecked, vbChecked))
                    If chkPagoDirecto.Value Then
                        txtReciboNombre.Text = lstrNombreConvPers
                        txtDireccion.Text = lstrDireccionConvPers
                        txtRFC.Text = lstrRFCConvPers
                        txtEmail.Text = IIf(IsNull(rsHono!vchEmail), "", rsHono!vchEmail)
                        txtObservaciones.Text = IIf(IsNull(rsHono!vchObservaciones), "", rsHono!vchObservaciones)
                        cboUsoCFDI.ListIndex = fintLocalizaCbo(cboUsoCFDI, IIf(IsNull(rsHono!intCveUsoCFDI), "0", rsHono!intCveUsoCFDI))
                    Else
                        Set rsUsoCFDIFact = frsRegresaRs("SELECT INTCVEUSOCFDIHONOFACTURADO FROM PVPARAMETRO WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable, adLockOptimistic, adOpenDynamic)
                        If rsUsoCFDIFact.RecordCount > 0 Then
                            If IIf(IsNull(rsUsoCFDIFact!INTCVEUSOCFDIHONOFACTURADO), 0, rsUsoCFDIFact!INTCVEUSOCFDIHONOFACTURADO) = 0 Then
                                'Falta configurar el uso de CFDI para los honorarios en parámetros del módulo.
                                MsgBox SIHOMsg(1563), vbCritical, "Mensaje"
                                pLimpia
                                If lblnRecargarTarifas Then
                                    pCargaTarifas
                                End If
                                txtMovimientoPaciente.SetFocus
                                pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                                fblnDatosPaciente = False
                                Exit Function
                            Else
                                cboUsoCFDI.ListIndex = fintLocalizaCbo(cboUsoCFDI, IIf(IsNull(rsUsoCFDIFact!INTCVEUSOCFDIHONOFACTURADO), "0", rsUsoCFDIFact!INTCVEUSOCFDIHONOFACTURADO))
                            End If
                        Else
                            'Falta configurar el uso de CFDI para los honorarios en parámetros del módulo.
                            MsgBox SIHOMsg(1563), vbCritical, "Mensaje"
                            pLimpia
                            If lblnRecargarTarifas Then
                                pCargaTarifas
                            End If
                            txtMovimientoPaciente.SetFocus
                            pHabilita 0, 0, 1, 0, 0, 0, 0, 0
                            fblnDatosPaciente = False
                            Exit Function
                        End If
                        txtReciboNombre.Text = vgstrNombreHospitalCH
                        txtDireccion.Text = Trim(vgstrDireccionCH) & IIf(Trim(vgstrColoniaCH) = "", "", "   Col. " & Trim(vgstrColoniaCH)) & IIf(Trim(vgstrCodPostalCH) = "", "", "   CP " & Trim(vgstrCodPostalCH)) & IIf(Trim(vgstrCiudadCH) = "", "", " " & vgstrCiudadCH) & IIf(Trim(vgstrEstadoCH) = "", "", ", " & Trim(vgstrEstadoCH))
                        txtRFC.Text = vgstrRfCCH
                        txtEmail.Text = vgstrEmailCH
                    End If
                Else
                    chkPagoDirecto.Value = vbChecked
                    txtReciboNombre.Text = lstrNombreConvPers
                    txtDireccion.Text = Trim(lstrDireccionConvPers) 'rsPvSelDatosPaciente!DireccionEmpresa
                    txtRFC.Text = lstrRFCConvPers
                End If
            Else
                lstrCorreoElectronico = IIf(IsNull(rsPvSelDatosPaciente!CORREO), " ", rsPvSelDatosPaciente!CORREO)
                lstrDireccionConvPers = IIf(IsNull(rsPvSelDatosPaciente!DireccionPaciente), "", rsPvSelDatosPaciente!DireccionPaciente)
                lstrNombreConvPers = rsPvSelDatosPaciente!Nombre
                lstrRFCConvPers = Trim(Replace(Replace(Replace(rsPvSelDatosPaciente!RFCPaciente, "-", ""), "_", ""), " ", ""))
                
                
                chkPagoDirecto.Value = vbUnchecked
                txtReciboNombre.Text = lstrNombreConvPers
                txtDireccion.Text = lstrDireccionConvPers
                txtRFC.Text = lstrRFCConvPers
                txtEmail.Text = lstrCorreoElectronico
            End If
        Else
            vlTipoPacienteCatalogo = ""
        
            '¡La información no existe!
            MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            pEnfocaTextBox txtMovimientoPaciente
        End If
        rsPvSelDatosPaciente.Close
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosPaciente"))
End Function

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Or UCase(Chr(KeyAscii)) = "C" Then
            optTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            optTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
            optTipoPaciente(2).Value = UCase(Chr(KeyAscii)) = "C"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtMovimientoPaciente_KeyPress"))
End Sub

Private Sub txtObservaciones_GotFocus()
    pSelTextBox txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If vlTipoPacienteCatalogo = "CONVENIO" Then
            cmdAdjuntar.SetFocus
        Else
            cmdAgregar.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtReciboNombre_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtReciboNombre

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtReciboNombre_GotFocus"))
End Sub

Private Sub txtReciboNombre_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 13 Then
        txtDireccion.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtReciboNombre_KeyPress"))
End Sub

Private Sub txtRFC_GotFocus()
    On Error GoTo NotificaError

    pSelTextBox txtRFC

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtRfc_GotFocus"))
End Sub

Private Sub txtRFC_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    Dim vlstrcaracter As String

    If KeyAscii = 13 Then
        txtEmail.SetFocus
    Else
        If KeyAscii <> 8 Then
            vlstrcaracter = fStrRFCValido(Chr(KeyAscii))
            If vlstrcaracter <> "" Then
                KeyAscii = Asc(UCase(vlstrcaracter))
            Else
                KeyAscii = 7
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtRFC_KeyPress"))
End Sub

Private Sub cmdCerrarComisiones_Click()
    On Error GoTo NotificaError

    cmdAceptarComisiones_Click
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCerrarComisiones_Click"))
End Sub

Private Sub lstComisiones_DblClick()
    On Error GoTo NotificaError

    cmdSelecciona_Click 0
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstComisiones_DblClick"))
End Sub

Private Sub lstComisionesAsignadas_DblClick()
    On Error GoTo NotificaError

    cmdSelecciona_Click 1
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":lstComisionesAsignadas_DblClick"))
End Sub

Private Sub chkRetencion_Click()
    On Error GoTo NotificaError
    Dim intcontador As Integer

    If txtMontoHonorario.Text <> "" Then
        pCalculoTotales
    End If

    cboTarifa.Enabled = chkRetencion.Value
    If Not cboTarifa.Enabled Then
        cboTarifa.ListIndex = -1
    Else
        intcontador = 0
        Do While intcontador <= UBound(arrTarifas(), 1)
            If arrTarifas(intcontador).dblPorcentaje = 10 Then
                cboTarifa.ListIndex = intcontador
            End If
            intcontador = intcontador + 1
        Loop
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkRetencion_Click"))
End Sub

Private Sub txtReciboHonorario_GotFocus()
    On Error GoTo NotificaError

    pHabilita 0, 0, 0, 0, 0, 1, 0, 0
    
    pSelTextBox txtReciboHonorario
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtReciboHonorario_GotFocus"))
End Sub

Private Sub txtReciboHonorario_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If Len(Trim(txtReciboHonorario.Text)) > 0 Then
            If vlblnConsulta Then
                If cmdSave.Enabled And cmdSave.Visible Then cmdSave.SetFocus
            End If
        End If
        pEnfocaTextBox txtConcepto
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtReciboHonorario_KeyDown"))
End Sub

Private Sub txtReciboHonorario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtReciboHonorario_KeyPress"))
End Sub

Private Sub pConfigura()
    On Error GoTo NotificaError

    With grdCargaHonorarios
        .Clear
        .Rows = 2
        .Cols = 21
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|||Médico|Monto|Retención|Comisión|IVA|Total|Recibo|Inicio|Fin|Concepto|Recibo|Dirección|R. F. C.|||Pago directo|Carta de autorización||Observaciones|RTP|ISR"

        .ColWidth(0) = 150
        .ColWidth(cintColCveMedico) = 0
        .ColWidth(cintColCuentaContableMedico) = 0
        .ColWidth(cintColNombreMedico) = 3800
        .ColWidth(cintColFechaAtencionInicio) = 1000
        .ColWidth(cintColFechaAtencionFin) = 1000
        .ColWidth(cintColReciboHonorario) = 1200
        .ColWidth(cintColConcepto) = 2500
        .ColWidth(cintColReciboNombre) = 2000
        .ColWidth(cintColDireccion) = 2000
        .ColWidth(cIntColRFC) = 1200
        .ColWidth(cintColMonto) = 1100
        .ColWidth(cintColRetencion) = 1100
        .ColWidth(cintColComision) = 1100
        .ColWidth(cintColIVAComision) = 1100
        .ColWidth(cintColTotalPagar) = 1100
        .ColWidth(cintColIdHonorario) = 0
        .ColWidth(cintColIdTarifa) = 0
        .ColWidth(cintColPagoDirecto) = 1000
        .ColWidth(cintColCarta) = 2000
        .ColWidth(cintColCartaCompleta) = 0
        .ColWidth(cintColObservaciones) = 3500
        .ColWidth(cintColRetencionRTP) = 0 'Ocultas para hacer la sumatoria de los totales separado por cada retención
        .ColWidth(cintColRetencionISR) = 0 '
        
        .ColAlignmentFixed(cintColNombreMedico) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColFechaAtencionInicio) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColFechaAtencionFin) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColReciboHonorario) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColConcepto) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColReciboNombre) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDireccion) = flexAlignCenterCenter
        .ColAlignmentFixed(cIntColRFC) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColMonto) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColRetencion) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColComision) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColIVAComision) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColTotalPagar) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColPagoDirecto) = flexAlignCenterCenter

        .ColAlignment(cintColNombreMedico) = flexAlignLeftCenter
        .ColAlignment(cintColFechaAtencionInicio) = flexAlignLeftCenter
        .ColAlignment(cintColFechaAtencionFin) = flexAlignLeftCenter
        .ColAlignment(cintColReciboHonorario) = flexAlignLeftCenter
        .ColAlignment(cintColConcepto) = flexAlignLeftCenter
        .ColAlignment(cintColReciboNombre) = flexAlignLeftCenter
        .ColAlignment(cintColDireccion) = flexAlignLeftCenter
        .ColAlignment(cIntColRFC) = flexAlignLeftCenter
        .ColAlignment(cintColMonto) = flexAlignRightCenter
        .ColAlignment(cintColRetencion) = flexAlignRightCenter
        .ColAlignment(cintColComision) = flexAlignRightCenter
        .ColAlignment(cintColIVAComision) = flexAlignRightCenter
        .ColAlignment(cintColTotalPagar) = flexAlignRightCenter
        .ColAlignment(cintColPagoDirecto) = flexAlignCenterCenter

        .ScrollBars = flexScrollBarBoth
    End With

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfigura"))

End Sub

Private Sub pVerificaHonorarioIgual()
    On Error GoTo NotificaError
        Dim vlblnDiferente As Boolean
        Dim vlaux As Integer
        Dim vlstrCampo As String

        If grdCargaHonorarios.Rows > 1 And grdCargaHonorarios.TextMatrix(1, 2) <> "" Then
            For vlintContador = 1 To grdCargaHonorarios.Rows - 1
                For vlaux = 2 To 14
                    Select Case vlaux
                        Case 2
                            vlstrCampo = cboMedicos.ItemData(cboMedicos.ListIndex)
                        Case 3
                            vlstrCampo = mskFechaAtencion.Text
                        Case 4
                            vlstrCampo = Trim(txtConcepto.Text)
                        Case 5
                            vlstrCampo = Val(Format(txtMontoHonorario.Text, "############.##"))
                        Case 6
                            vlstrCampo = Val(Format(txtRetencion.Text, "############.##"))
                        Case 7
                            vlstrCampo = Val(Format(txtCantidadComision.Text, "############.##"))
                        Case 8
                            vlstrCampo = Val(Format(txtIVAComision.Text, "############.##"))
                        Case 9
                            vlstrCampo = txtReciboNombre.Text
                        Case 10
                            vlstrCampo = txtDireccion.Text
                        Case 11
                            vlstrCampo = Trim(Replace(Replace(Replace(txtRFC.Text, "-", ""), "_", ""), " ", ""))
                        Case 12
                            vlstrCampo = txtReciboHonorario.Text
                        Case 13
                            vlstrCampo = mskFechaAtencionFin.Text
                        Case 14
                            vlstrCampo = txtTotal.Text
                    End Select
                    If grdCargaHonorarios.TextMatrix(vlintContador, vlaux) = vlstrCampo Then
                        If vlaux = 14 Then
                            If MsgBox(Mid(SIHOMsg(106), 1, 35) & "!" & Chr(13) & "¿desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Mensaje") = vbNo Then
                                vlblnContinuaRevision = False
                                Exit Sub
                            Else
                                vlblnContinuaRevision = True
                                Exit Sub
                            End If
                        Else
                            vlblnDiferente = False
                        End If
                    Else
                        vlblnDiferente = True
                        Exit For
                    End If
                Next
            Next
        End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pVerificaHonorarioIgual"))
End Sub

Private Sub pCargaFilMedico()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    vlstrSentencia = "SELECT  HoMedico.intCveMedico, " & _
        " rtrim(ltrim(HoMedico.vchApellidoPaterno)) || ' ' || " & _
        " rtrim(ltrim(HoMedico.vchApellidoMaterno)) || ' ' || " & _
        " rtrim(LTrim(HoMedico.vchNombre)) As Nombre " & _
        " From Homedico " & _
        " where bitEstaActivo = 1 "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboFilMedico, rs, 0, 1
    End If
    rs.Close

    cboFilMedico.AddItem "<TODOS>", 0
    cboFilMedico.ItemData(0) = 0
    cboFilMedico.ListIndex = 0

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pVerificaHonorarioIgual"))
End Sub

Private Sub grdHonorarios_DblClick()
    On Error GoTo NotificaError
    
    If grdHonorarios.Row > 0 Then
    'Si no es la columna para marcar el * o si no es la columna de recibo
        If Val(grdHonorarios.TextMatrix(grdHonorarios.Row, 1)) <> 0 And grdHonorarios.Col <> 8 Then
            If grdHonorarios.Col > 1 Then
                vllngNumHonorario = CLng(grdHonorarios.TextMatrix(grdHonorarios.Row, 1))
                pMuestraHonorario
                cmdLocate.SetFocus
            Else
                If grdHonorarios.TextMatrix(grdHonorarios.Row, cintColFilEstatus) <> "C" Then
                    grdHonorarios.Col = 0

                    pMarca "*", grdHonorarios.Row

                End If
            End If
        Else
            txtFilCuenta.SetFocus
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdHonorarios_DblClick"))
End Sub

Private Sub pCalculaTotales()
    On Error GoTo NotificaError

    Dim intIndex As Integer
    Dim dblMonto As Double
    Dim dblretencion As Double
    Dim dblComisiones As Double
    Dim dblComisioIVA As Double
    Dim dblTotal As Double
    Dim dblRTP As Double
        
    If Trim(grdCargaHonorarios.TextMatrix(1, 1)) <> "" Then
        For intIndex = 1 To grdCargaHonorarios.Rows - 1
            If grdCargaHonorarios.TextMatrix(intIndex, 2) <> "" Then
                dblMonto = dblMonto + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColMonto))
                dblretencion = dblretencion + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColRetencionISR))
                dblComisiones = dblComisiones + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColComision))
                dblComisioIVA = dblComisioIVA + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColIVAComision))
                dblTotal = dblTotal + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColTotalPagar))
                dblRTP = dblRTP + CDbl(grdCargaHonorarios.TextMatrix(intIndex, cintColRetencionRTP))
            End If
        Next
    End If
    lblTotalMonto.Caption = FormatCurrency(dblMonto)
    lblTotalRetencion.Caption = FormatCurrency(dblretencion)
    lblTotalRTP.Caption = FormatCurrency(dblRTP)
    lblTotalComisiones.Caption = FormatCurrency(dblComisiones)
    lblTotalComisionIVA.Caption = FormatCurrency(dblComisioIVA)
    lblTotalTotal.Caption = FormatCurrency(dblTotal)
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCalculaTotales"))
End Sub

Private Sub pMarca(lstrCaracter As String, llngRenglon As Long)
    Dim llngContador As Long
    
    With grdHonorarios
        
        If llngRenglon > 0 Then
            .TextMatrix(llngRenglon, 0) = IIf(.TextMatrix(llngRenglon, 0) = lstrCaracter, "", lstrCaracter)
            
            If Trim(.TextMatrix(llngRenglon, 0)) = "" Then
                llngTotalSel = llngTotalSel - 1
            Else
                llngTotalSel = llngTotalSel + 1
            End If
            
            .Col = 0
            .Row = llngRenglon
            .CellFontBold = vbBlackness
        Else
            'Todos o Invertir selección
            If llngRenglon = -1 Then
                For llngContador = 1 To .Rows - 1
                    If grdHonorarios.TextMatrix(llngContador, cintColFilEstatus) <> "C" Then
                        .TextMatrix(llngContador, 0) = IIf(.TextMatrix(llngContador, 0) = lstrCaracter, "", lstrCaracter)
                        
                        If Trim(.TextMatrix(llngContador, 0)) = "" Then
                            llngTotalSel = llngTotalSel - 1
                        Else
                            llngTotalSel = llngTotalSel + 1
                        End If
                        
                        .Col = 0
                        .Row = llngContador
                        .CellFontBold = vbBlackness
                    End If
                Next llngContador
            End If
        End If
        .Row = llngRenglon
    End With
    
    cmdPagoAnticipado.Enabled = llngTotalSel <> 0
    cmdDelete.Enabled = llngTotalSel <> 0

End Sub

Private Sub pCargaUsosCFDI()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDI, rsTmp, 0, 1
        cboUsoCFDI.ListIndex = -1
    End If
End Sub

Private Function fstrSoloNombreArchivo(strNombreArchivo) As String
    fstrSoloNombreArchivo = Mid(strNombreArchivo, InStrRev(strNombreArchivo, "\") + 1, Len(strNombreArchivo))
End Function

Public Function fDatosValidosEfectivo() As Boolean
    On Error GoTo NotificaError
    Dim vlIntCont As Integer
    Dim vlstrsql As String
    Dim rsPagosHonorario As ADODB.Recordset
    Dim rsFormaPago As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim vldblPagoTotal As Double
    
    fDatosValidosEfectivo = True
    
    Set rsPagosHonorario = frsEjecuta_SP(CStr(vllngNumHonorario), "sp_pvselpagohonorario")
    If rsPagosHonorario.RecordCount > 0 Then
        If IIf(IsNull(rsPagosHonorario!CuentaBancaria), 1, 0) = 1 Then
            vlstrFolioDoc = rsPagosHonorario!folio
            vllngNumCorte = rsPagosHonorario!NumCorte
            Do While Not rsPagosHonorario.EOF
                If Trim(rsPagosHonorario!tipoPago) = "B" Or Trim(rsPagosHonorario!tipoPago) = "T" Or Trim(rsPagosHonorario!tipoPago) = "H" Or Trim(rsPagosHonorario!tipoPago) = "E" Then
                    vldblPagoTotal = vldblPagoTotal + Round(rsPagosHonorario!cantidadPago * IIf(rsPagosHonorario!PESOS = 1, 1, rsPagosHonorario!TipoCambio), 2)
                Else
                    fDatosValidosEfectivo = False
                    'El honorario no fue pagado con tarjeta, transferencia o cheque.
                    MsgBox SIHOMsg(1576), vbCritical, "Mensaje"
                    Exit Function
                End If
                rsPagosHonorario.MoveNext
            Loop
        Else
            fDatosValidosEfectivo = False
            'El honorario ya tiene una cuenta bancaria asignada en cuentas por pagar.
            MsgBox SIHOMsg(1578), vbCritical, "Mensaje"
            Exit Function
        End If
    End If
    
    If fDatosValidosEfectivo Then
        vlstrSentencia = "Select isnull(intNumeroCuenta,0) Cuenta From HoMedicoEmpresa " & _
                        " Where intCLAveMedico = " & Trim(Str(cboMedicos.ItemData(cboMedicos.ListIndex))) & _
                        " and TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
        Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsTemp.RecordCount <> 0 Then
            If rsTemp!cuenta <> 0 Then
                If Not fblnCuentaAfectable(fstrCuentaContable(CStr(rsTemp!cuenta)), vgintClaveEmpresaContable) Then
                    fDatosValidosEfectivo = False
                    'La cuenta contable asignada al médico no acepta movimientos.
                    MsgBox SIHOMsg(1236), vbExclamation, "Mensaje"
                Else
                    vllngCuentaMedico = rsTemp!cuenta
                End If
            Else
                'El médico no tiene una cuenta contable asignada, favor de verificarlo.
                MsgBox SIHOMsg(519), vbCritical, "Mensaje"
                fDatosValidosEfectivo = False
            End If
        End If
    End If
    
    If fDatosValidosEfectivo Then
        fDatosValidosEfectivo = fblnFormasPagoPos(aFormasPago(), vldblPagoTotal, True, vldblTipoCambio, False, 0, vlstrAdTipoPaciente, Trim(txtRFC.Text), False, False, True, "Honorarios")
    End If
    
    If fDatosValidosEfectivo Then
        
        For vlIntCont = 0 To UBound(aFormasPago(), 1)
            If vlIntCont = 1 Then
                fDatosValidosEfectivo = False
                'Este pago solo puede ser de tipo "EFECTIVO".
                MsgBox SIHOMsg(1577), vbCritical, "Mensaje"
                Exit Function
            Else
                vlstrsql = "SELECT  BITPESOS, CHRTIPO FROM PVFORMAPAGO WHERE INTFORMAPAGO = " & aFormasPago(vlIntCont).vlintNumFormaPago
                Set rsFormaPago = frsRegresaRs(vlstrsql)
                If rsFormaPago.RecordCount > 0 Then
                    If Trim(rsFormaPago!chrTipo) <> "E" Then
                        fDatosValidosEfectivo = False
                        'Este pago solo puede ser de tipo "EFECTIVO".
                        MsgBox SIHOMsg(1577), vbCritical, "Mensaje"
                        Exit Function
                    ElseIf rsFormaPago!BITPESOS = 0 Then
                        fDatosValidosEfectivo = False
                        'Este pago solo puede ser de tipo "EFECTIVO".
                        MsgBox SIHOMsg(1577), vbCritical, "Mensaje"
                        Exit Function
                    End If
                End If
            End If
        Next vlIntCont
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fDatosValidosEfectivo"))
End Function





