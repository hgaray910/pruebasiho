VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGrupoFacturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupo de cuentas"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTFacturacionConsolidada 
      Height          =   17610
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   31062
      _Version        =   393216
      TabHeight       =   529
      TabCaption(0)   =   "Datos generales"
      TabPicture(0)   =   "frmGrupoFacturas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label21"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label24"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "msfgCargosAsignados"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraBotonera"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FreTotales"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkCosto"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Contenido del gru&po"
      TabPicture(1)   =   "frmGrupoFacturas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(1)=   "lblDetalleCargos"
      Tab(1).Control(2)=   "Label2"
      Tab(1).Control(3)=   "lblCtaSeleccionada"
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(5)=   "Shape2"
      Tab(1).Control(6)=   "Shape4"
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(8)=   "fraFiltrosContenidoGrupoE"
      Tab(1).Control(9)=   "fraFiltrosContenidoGrupoP"
      Tab(1).Control(10)=   "MSFGCuentasAsignadasGrupo"
      Tab(1).Control(11)=   "MSFGCuentasDisponiblesGrupo"
      Tab(1).Control(12)=   "MSFGCargosDisponibles"
      Tab(1).Control(13)=   "cmdQuitaTodos"
      Tab(1).Control(14)=   "cmdQuitaUno"
      Tab(1).Control(15)=   "cmdAgregaUno"
      Tab(1).Control(16)=   "cmdAgregaTodo"
      Tab(1).Control(17)=   "chkMuestraDetalleAutomaticamente"
      Tab(1).Control(18)=   "cmdVerDetalleCargos"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Consulta de grupos"
      TabPicture(2)   =   "frmGrupoFacturas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame5"
      Tab(2).Control(3)=   "fraOpcional"
      Tab(2).Control(4)=   "cmdBuscaGrupos"
      Tab(2).Control(5)=   "MSFGResultado"
      Tab(2).ControlCount=   6
      Begin VB.CheckBox chkCosto 
         Caption         =   "Incluir el costo del cargo"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Incluir costos"
         Top             =   8100
         Width           =   2280
      End
      Begin VB.Frame Frame7 
         Height          =   930
         Left            =   -68400
         TabIndex        =   5
         Top             =   400
         Width           =   5175
         Begin VB.CheckBox chkRangoFechasB 
            Caption         =   "Rango de fechas"
            Height          =   255
            Left            =   240
            TabIndex        =   99
            Top             =   0
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFechaInicialB 
            Height          =   315
            Left            =   1200
            TabIndex        =   100
            Top             =   400
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   60030977
            CurrentDate     =   38147
         End
         Begin MSComCtl2.DTPicker dtpFechaFinalB 
            Height          =   315
            Left            =   3600
            TabIndex        =   101
            Top             =   400
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   60030977
            CurrentDate     =   38147
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Fecha final"
            Height          =   195
            Left            =   2640
            TabIndex        =   7
            Top             =   460
            Width           =   780
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Fecha inicial"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   460
            Width           =   885
         End
      End
      Begin VB.Frame Frame10 
         Height          =   930
         Left            =   -74880
         TabIndex        =   4
         Top             =   400
         Width           =   6510
         Begin VB.OptionButton optTipoGrupoB 
            Caption         =   "Todos"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   97
            Top             =   210
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTipoGrupoB 
            Caption         =   "Particular"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   96
            Top             =   210
            Width           =   975
         End
         Begin VB.OptionButton optTipoGrupoB 
            Caption         =   "Empresa"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   95
            Top             =   210
            Width           =   1095
         End
         Begin VB.ComboBox cboEmpresa 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   480
            Width           =   6015
         End
      End
      Begin VB.Frame Frame5 
         Height          =   680
         Left            =   -74880
         TabIndex        =   76
         Top             =   1200
         Width           =   1335
         Begin VB.CheckBox chkFacturados 
            Caption         =   "Facturados"
            Height          =   375
            Left            =   150
            TabIndex        =   102
            Top             =   190
            Width           =   1160
         End
      End
      Begin VB.CommandButton cmdVerDetalleCargos 
         Caption         =   "&Ver"
         Height          =   375
         Left            =   -69280
         TabIndex        =   60
         Top             =   4555
         Width           =   495
      End
      Begin VB.CheckBox chkMuestraDetalleAutomaticamente 
         Caption         =   "Muestra detalle automáticamente"
         Height          =   255
         Left            =   -65880
         TabIndex        =   68
         ToolTipText     =   "Muestra detalle de la cuenta"
         Top             =   5335
         Width           =   2655
      End
      Begin VB.CommandButton cmdAgregaTodo 
         Caption         =   ">>"
         Height          =   375
         Left            =   -69280
         TabIndex        =   56
         Top             =   3100
         Width           =   495
      End
      Begin VB.CommandButton cmdAgregaUno 
         Caption         =   ">"
         Height          =   375
         Left            =   -69280
         TabIndex        =   57
         Top             =   3460
         Width           =   495
      End
      Begin VB.CommandButton cmdQuitaUno 
         Caption         =   "<"
         Height          =   375
         Left            =   -69280
         TabIndex        =   58
         Top             =   3820
         Width           =   495
      End
      Begin VB.CommandButton cmdQuitaTodos 
         Caption         =   "<<"
         Height          =   375
         Left            =   -69280
         TabIndex        =   59
         Top             =   4180
         Width           =   495
      End
      Begin VB.Frame Frame9 
         Height          =   975
         Left            =   120
         TabIndex        =   41
         Top             =   340
         Width           =   1455
         Begin VB.OptionButton optTipoGrupo 
            Caption         =   "Particulares"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optTipoGrupo 
            Caption         =   "Empresa"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Height          =   1335
         Left            =   9240
         TabIndex        =   36
         Top             =   340
         Width           =   2535
         Begin VB.TextBox txtDGFolioFactura 
            Enabled         =   0   'False
            Height          =   315
            Left            =   195
            TabIndex        =   38
            Top             =   880
            Width           =   2175
         End
         Begin VB.TextBox txtFechaCreacion 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   1275
            TabIndex        =   37
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Creación"
            Height          =   195
            Left            =   195
            TabIndex        =   40
            Top             =   300
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Folio factura"
            Height          =   195
            Left            =   195
            TabIndex        =   39
            Top             =   615
            Width           =   870
         End
      End
      Begin VB.Frame FreTotales 
         Enabled         =   0   'False
         Height          =   2060
         Left            =   7540
         TabIndex        =   24
         Top             =   7060
         Width           =   4235
         Begin VB.TextBox txtRetServicios 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   115
            ToolTipText     =   "Total del presupuesto"
            Top             =   1375
            Width           =   1890
         End
         Begin VB.TextBox txtTotalAPagar 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   114
            ToolTipText     =   "Total del presupuesto"
            Top             =   1670
            Width           =   1890
         End
         Begin VB.TextBox txtTotal 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   31
            ToolTipText     =   "Total del presupuesto"
            Top             =   1080
            Width           =   1890
         End
         Begin VB.TextBox txtIva 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "Iva del presupuesto"
            Top             =   785
            Width           =   1890
         End
         Begin VB.TextBox txtSubtotal 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Subtotal del presupuesto"
            Top             =   490
            Width           =   1890
         End
         Begin VB.TextBox txtDescuentos 
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
            Height          =   285
            Left            =   2215
            Locked          =   -1  'True
            TabIndex        =   29
            ToolTipText     =   "Total de descuentos"
            Top             =   195
            Width           =   1890
         End
         Begin VB.Label lblRetServicios 
            AutoSize        =   -1  'True
            Caption         =   "Retención servicios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   100
            TabIndex        =   117
            Top             =   1395
            Width           =   2055
         End
         Begin VB.Label lblTotalAPagar 
            AutoSize        =   -1  'True
            Caption         =   "Total a pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   100
            TabIndex        =   116
            Top             =   1695
            Width           =   1425
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Total cargos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   100
            TabIndex        =   35
            Top             =   1095
            Width           =   1335
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
            Height          =   240
            Left            =   100
            TabIndex        =   34
            Top             =   810
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   100
            TabIndex        =   32
            Top             =   512
            Width           =   870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Descuentos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   100
            TabIndex        =   33
            Top             =   217
            Width           =   1245
         End
      End
      Begin VB.Frame fraBotonera 
         Height          =   720
         Left            =   120
         TabIndex        =   16
         Top             =   8400
         Width           =   4620
         Begin VB.CommandButton cmdVistaPreliminar 
            Height          =   495
            Left            =   3560
            Picture         =   "frmGrupoFacturas.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Vista previa"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   4050
            Picture         =   "frmGrupoFacturas.frx":01F6
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Imprimir"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEliminar 
            Height          =   495
            Left            =   3060
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGrupoFacturas.frx":0398
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Eliminar grupo"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdFin 
            Height          =   495
            Left            =   2070
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGrupoFacturas.frx":053A
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguiente 
            Height          =   495
            Left            =   1575
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGrupoFacturas.frx":06AC
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   495
            Left            =   1080
            Picture         =   "frmGrupoFacturas.frx":081E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Búsqueda de pacientes"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdInicio 
            Height          =   495
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGrupoFacturas.frx":0990
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnterior 
            Height          =   495
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGrupoFacturas.frx":0B02
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Registro anterior"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabar 
            Height          =   495
            Left            =   2565
            Picture         =   "frmGrupoFacturas.frx":0FF4
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Grabar registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   1695
         TabIndex        =   13
         Top             =   345
         Width           =   7455
         Begin VB.TextBox txtCveGrupo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            MaxLength       =   5
            TabIndex        =   0
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox txtEmpresa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   555
            Width           =   6255
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Empresa"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   585
            Width           =   735
         End
      End
      Begin VB.Frame fraOpcional 
         Height          =   680
         Left            =   -73570
         TabIndex        =   8
         Top             =   1200
         Width           =   9755
         Begin VB.OptionButton optTipoPacB 
            Caption         =   "Ambos"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   106
            Top             =   280
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optTipoPacB 
            Caption         =   "Externos"
            Height          =   255
            Index           =   1
            Left            =   4800
            TabIndex        =   105
            Top             =   280
            Width           =   975
         End
         Begin VB.OptionButton optTipoPacB 
            Caption         =   "Internos"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   104
            Top             =   280
            Width           =   975
         End
         Begin VB.TextBox txtFolioFacturaB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   103
            Top             =   220
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Height          =   1215
            Left            =   2400
            TabIndex        =   9
            Top             =   0
            Width           =   25
         End
         Begin VB.TextBox txtCtaPacB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8400
            TabIndex        =   107
            Top             =   220
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Folio factura"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   280
            Width           =   870
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta del paciente"
            Height          =   195
            Left            =   6840
            TabIndex        =   11
            Top             =   280
            Width           =   1425
         End
         Begin VB.Label Label14 
            Caption         =   "Tipo de paciente"
            Height          =   255
            Left            =   2520
            TabIndex        =   10
            Top             =   280
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdBuscaGrupos 
         Height          =   495
         Left            =   -63765
         Picture         =   "frmGrupoFacturas.frx":16F6
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Búsqueda de pacientes"
         Top             =   1365
         UseMaskColor    =   -1  'True
         Width           =   540
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFGCargosDisponibles 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   69
         Top             =   5625
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   4683
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VSFlex7LCtl.VSFlexGrid MSFGCuentasDisponiblesGrupo 
         Height          =   2532
         Left            =   -74880
         TabIndex        =   70
         Top             =   2760
         Width           =   5412
         _cx             =   9551
         _cy             =   4471
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         WordWrap        =   -1  'True
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
      Begin VSFlex7LCtl.VSFlexGrid MSFGResultado 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   109
         Top             =   1920
         Width           =   11655
         _cx             =   20558
         _cy             =   11456
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         WordWrap        =   -1  'True
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfgCargosAsignados 
         Height          =   5175
         Left            =   120
         TabIndex        =   71
         Top             =   1800
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9128
         _Version        =   393216
         GridColor       =   12632256
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VSFlex7LCtl.VSFlexGrid MSFGCuentasAsignadasGrupo 
         Height          =   2535
         Left            =   -68640
         TabIndex        =   77
         Top             =   2760
         Width           =   5415
         _cx             =   9551
         _cy             =   4471
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         WordWrap        =   -1  'True
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
      Begin VB.Frame fraFiltrosContenidoGrupoP 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   80
         Top             =   460
         Visible         =   0   'False
         Width           =   11655
         Begin VB.Frame Frame14 
            Height          =   735
            Left            =   4560
            TabIndex        =   93
            Top             =   1005
            Width           =   5415
            Begin VB.ComboBox cboConceptoFacturacionP 
               Height          =   315
               Left            =   2160
               Style           =   2  'Dropdown List
               TabIndex        =   89
               Top             =   240
               Width           =   3015
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Concepto de facturación"
               Height          =   195
               Left            =   240
               TabIndex        =   94
               Top             =   300
               Width           =   1755
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "  Rango de fechas  "
            Height          =   1530
            Left            =   240
            TabIndex        =   83
            Top             =   210
            Width           =   4215
            Begin MSComCtl2.DTPicker dtpFechaInicialP 
               Height          =   315
               Left            =   2160
               TabIndex        =   84
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   60030977
               CurrentDate     =   38098
            End
            Begin MSComCtl2.DTPicker dtpFechaFinalP 
               Height          =   315
               Left            =   2160
               TabIndex        =   85
               Top             =   960
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   60030977
               CurrentDate     =   38098
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Fecha final"
               Height          =   195
               Left            =   600
               TabIndex        =   92
               Top             =   1020
               Width           =   780
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Fecha inicial"
               Height          =   195
               Left            =   600
               TabIndex        =   90
               Top             =   540
               Width           =   885
            End
         End
         Begin VB.CommandButton cmdEjecutarConsultaP 
            Caption         =   "&Ejecutar consulta"
            Height          =   735
            Left            =   10200
            TabIndex        =   91
            Top             =   600
            Width           =   1215
         End
         Begin VB.Frame Frame12 
            Height          =   735
            Left            =   4560
            TabIndex        =   81
            Top             =   210
            Width           =   5415
            Begin VB.OptionButton optTipoPacP 
               Caption         =   "Ambos"
               Height          =   255
               Index           =   2
               Left            =   4320
               TabIndex        =   88
               Top             =   320
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton optTipoPacP 
               Caption         =   "Externos"
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   87
               Top             =   320
               Width           =   975
            End
            Begin VB.OptionButton optTipoPacP 
               Caption         =   "Internos"
               Height          =   255
               Index           =   0
               Left            =   2160
               TabIndex        =   86
               Top             =   320
               Width           =   975
            End
            Begin VB.Label Label17 
               Caption         =   "Tipo de paciente"
               Height          =   255
               Left            =   240
               TabIndex        =   82
               Top             =   315
               Width           =   1335
            End
         End
      End
      Begin VB.Frame fraFiltrosContenidoGrupoE 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   42
         Top             =   460
         Width           =   11655
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   240
            TabIndex        =   62
            Top             =   840
            Width           =   4575
            Begin VB.OptionButton optTipoPac 
               Caption         =   "Internos"
               Height          =   255
               Index           =   0
               Left            =   1470
               TabIndex        =   48
               Top             =   240
               Width           =   880
            End
            Begin VB.OptionButton optTipoPac 
               Caption         =   "Externos"
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   49
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton optTipoPac 
               Caption         =   "Ambos"
               Height          =   255
               Index           =   2
               Left            =   3480
               TabIndex        =   50
               Top             =   240
               Value           =   -1  'True
               Width           =   790
            End
            Begin VB.TextBox txtNumCuenta 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3060
               TabIndex        =   51
               Top             =   560
               Width           =   1215
            End
            Begin VB.Label lblTipoPaciente 
               Caption         =   "Tipo de paciente"
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label lblNumCuenta 
               Caption         =   "Número de cuenta del paciente"
               Height          =   255
               Left            =   120
               TabIndex        =   63
               Top             =   590
               Width           =   2415
            End
         End
         Begin VB.CommandButton cmdEjecutaConsulta 
            Caption         =   "&Ejecutar consulta"
            Height          =   665
            Left            =   10200
            TabIndex        =   55
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cboTipoCargo 
            Height          =   315
            ItemData        =   "frmGrupoFacturas.frx":1868
            Left            =   6840
            List            =   "frmGrupoFacturas.frx":187B
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Frame Frame6 
            Height          =   735
            Left            =   4920
            TabIndex        =   52
            Top             =   160
            Width           =   6495
            Begin VB.Frame Frame1 
               Height          =   855
               Left            =   4080
               TabIndex        =   61
               Top             =   0
               Width           =   25
            End
            Begin VB.CheckBox chkIncluirCargosAtrasados 
               Caption         =   "Incluir cargos atrasados"
               Enabled         =   0   'False
               Height          =   375
               Left            =   4320
               TabIndex        =   47
               Top             =   240
               Width           =   2055
            End
            Begin VB.CheckBox chkRangoFechas 
               Caption         =   "Rango de fechas"
               Height          =   255
               Left            =   240
               TabIndex        =   44
               Top             =   0
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker dtpFechaIni 
               Height          =   315
               Left            =   480
               TabIndex        =   45
               Top             =   320
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   60030977
               CurrentDate     =   38098
            End
            Begin MSComCtl2.DTPicker dtpFechaFin 
               Height          =   315
               Left            =   2280
               TabIndex        =   46
               Top             =   320
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   60030977
               CurrentDate     =   38098
            End
         End
         Begin VB.ComboBox cboConvenio 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   480
            Width           =   4575
         End
         Begin VB.ComboBox cboConceptoFacturacion 
            Height          =   315
            Left            =   6840
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label lblTipoCargo 
            Caption         =   "Tipo de cargo"
            Height          =   255
            Left            =   4920
            TabIndex        =   67
            Top             =   1470
            Width           =   2415
         End
         Begin VB.Label lblConceptoFacturacion 
            Caption         =   "Concepto de facturación"
            Height          =   255
            Left            =   4920
            TabIndex        =   66
            Top             =   1110
            Width           =   2175
         End
         Begin VB.Label lblConvenio 
            Caption         =   "Empresa"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Precio modificado"
         Height          =   195
         Left            =   420
         TabIndex        =   113
         Top             =   7080
         Width           =   1260
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0080C0FF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   120
         Top             =   7080
         Width           =   255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   1770
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Fecha modificada"
         Height          =   195
         Left            =   2070
         TabIndex        =   112
         Top             =   7095
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Precio modificado"
         Height          =   195
         Left            =   -74580
         TabIndex        =   111
         Top             =   8400
         Width           =   1260
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H0080C0FF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   -74880
         Top             =   8400
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0FFFF&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Left            =   -73230
         Top             =   8400
         Width           =   255
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Fecha modificada"
         Height          =   195
         Left            =   -72930
         TabIndex        =   110
         Top             =   8415
         Width           =   1260
      End
      Begin VB.Label lblCtaSeleccionada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   78
         Top             =   5340
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Cuentas disponibles para agregar al grupo"
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   2500
         Width           =   3855
      End
      Begin VB.Label lblDetalleCargos 
         Caption         =   "Detalle de los cargos "
         Height          =   255
         Left            =   -74880
         TabIndex        =   74
         Top             =   5340
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Cuentas asignadas al grupo"
         Height          =   255
         Left            =   -68640
         TabIndex        =   73
         Top             =   2500
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Cargos en el grupo"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1440
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmGrupoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim cuenta As Integer
Dim tipo As String
Dim vlintEmpresa As Integer
Dim vlintCveConcepto As Long
Dim vldblPorcIVA As Double
Dim vldblMontoIVA As Double
Dim vlstrCveCargo As String
Dim vlstrTipoCargoEmpresa As String
Dim vlstrDescripcionCargo As String
Dim vldblPorcRetServicios As Double
Dim vldblCantRetServicios As Double
Dim vldblRetFactura As Double
Dim vldblDescFactura As Double
Dim vldblIVAFactura As Double
Dim vldblTotalFactura As Double

Private vgrptReporte As CRAXDRT.Report

Private Enum enmStatus
    stNuevo = 1
    stedicion = 2
    stEspera = 3
    stFacturado = 4
    stConsulta = 5
End Enum

Private vgintCveEmpresa As Integer 'Sirve para identificar la empresa con la cual se ejecutó la búsqueda de cargos o la consulta de un grupo
Private vgenmStatus As enmStatus 'Estado general del módulo
Private vgintGrupoAlEntrar As Integer 'No permite identificar si se cargo un grupo y despues se intenta grabar estos cargos con otro número de grupo
Private vgrsGrupoFacturas As New ADODB.Recordset 'Para hacer el scroll
Private vgintListIndexAlEntrar As Integer 'Para saber si cambio el combo de Convenio

Private Type varPaciente
    NoCuenta As Long
    tipo As String
End Type

Private Sub cboConvenio_Click()
    Dim rsRetencion As ADODB.Recordset
    Dim vlstrsql As String
    
    If cboConvenio.Enabled Then
        pInicializaContenido
        vlstrsql = "SELECT RELPORCENTAJESERVICIOSEMP FROM CCEMPRESA WHERE INTCVEEMPRESA = " & cboConvenio.ItemData(cboConvenio.ListIndex)
        Set rsRetencion = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        vldblPorcRetServicios = IIf(IsNull(rsRetencion!RELPORCENTAJESERVICIOSEMP), 0, rsRetencion!RELPORCENTAJESERVICIOSEMP)
    End If
End Sub

Private Sub cboConvenio_DropDown()
    vgintListIndexAlEntrar = cboConvenio.ListIndex
End Sub



Private Sub chkMuestraDetalleAutomaticamente_Click()
    If chkMuestraDetalleAutomaticamente.Value Then
        cmdVerDetalleCargos_Click
        cmdVerDetalleCargos.Enabled = False
    Else
        cmdVerDetalleCargos.Enabled = True
    End If
End Sub

Private Sub chkRangoFechas_Click()
    'Solo pone con negritas el filtro seleccionado
    If chkRangoFechas.Value Then
        chkIncluirCargosAtrasados.Enabled = True
        dtpFechaIni.Enabled = True
        dtpFechaFin.Enabled = True
        dtpFechaIni.SetFocus
    Else
        chkIncluirCargosAtrasados.Enabled = False
        dtpFechaIni.Enabled = False
        dtpFechaFin.Enabled = False
    End If
End Sub

Private Sub chkRangoFechasB_Click()
    If chkRangoFechasB.Value Then
        dtpFechaInicialB.Enabled = True
        dtpFechaFinalB.Enabled = True
        dtpFechaInicialB.SetFocus
    Else
        dtpFechaInicialB.Enabled = False
        dtpFechaFinalB.Enabled = False
    End If
End Sub

Private Sub cmdAgregaTodo_Click()
    Dim intCont As Integer
    If MSFGCuentasDisponiblesGrupo.TextMatrix(1, 1) <> "" Then
        MSFGCuentasDisponiblesGrupo.Row = 1
        For intCont = 0 To MSFGCuentasDisponiblesGrupo.Rows - 2
            cmdAgregaUno_Click
        Next
    End If
End Sub

Private Sub cmdAgregaUno_Click()
  
  pMueveCuenta MSFGCuentasDisponiblesGrupo, MSFGCuentasAsignadasGrupo
   Dim intRows As Integer
            
'   pConfiguraGridCargos msfgCargosAsignados
'   For intRows = 1 To MSFGCuentasAsignadasGrupo.Rows - 1
   pLlenaCargos Val(MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Rows - 1, 0)), _
   MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Rows - 1, 1), _
   "C", _
   msfgCargosAsignados, _
           MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Rows - 1, 3), _
                             1
       ' Next
pmarcados
End Sub

Private Sub pHabilitaBotonesPaso()
    '|  Si no hay cuentas disponibles deshabilita los botones para agregar y viceversa
    If MSFGCuentasDisponiblesGrupo.TextMatrix(1, 0) = "" Then
        cmdAgregaUno.Enabled = False
        cmdAgregaTodo.Enabled = False
    Else
        cmdAgregaUno.Enabled = True
        cmdAgregaTodo.Enabled = True
    End If
    '|  Si no hay cuentas asignadas deshabilita los botones para quitar y viceversa
    If MSFGCuentasAsignadasGrupo.TextMatrix(1, 0) = "" Then
        cmdQuitaUno.Enabled = False
        cmdQuitaTodos.Enabled = False
    Else
        cmdQuitaUno.Enabled = True
        cmdQuitaTodos.Enabled = True
    End If
End Sub


Private Sub pMueveCuenta(msfgOrigen As VSFlexGrid, msfgDestino As VSFlexGrid)
    With msfgOrigen
        '|  Validaciones de datos
        If .Row < 1 Then
            '|  ¡Dato no válido, seleccione un valor de la lista!
            MsgBox SIHOMsg(3), vbCritical, "Mensaje"
            msfgOrigen.Row = 1
            msfgOrigen.SetFocus
            Exit Sub
        End If
        If Not fblnExisteCuenta(msfgDestino, .TextMatrix(.Row, 0), .TextMatrix(.Row, 1)) Then
            '|  Si el primer renglón del grid de las cuentas asiganadas esta vacío
            '|  lo usa para pasar la cuenta, sino agrega un nuevo renglón
            If msfgDestino.TextMatrix(1, 1) = "" Then
                msfgDestino.TextMatrix(1, 0) = .TextMatrix(.Row, 0)
                msfgDestino.TextMatrix(1, 1) = .TextMatrix(.Row, 1)
                msfgDestino.TextMatrix(1, 2) = .TextMatrix(.Row, 2)
                msfgDestino.TextMatrix(1, 3) = .TextMatrix(.Row, 3)
            Else
                msfgDestino.AddItem .TextMatrix(.Row, 0) & Chr(9) & .TextMatrix(.Row, 1) & Chr(9) & .TextMatrix(.Row, 2) & Chr(9) & .TextMatrix(.Row, 3)
            End If
            '|  Si solo queda un renglón en el grid de las cuentas disponibles
            '|  lo limpia, sino lo elimina
            If .Rows = 2 Then
                pConfiguraGridCuentas msfgOrigen
            Else
                .RemoveItem .Row
            End If
            If vgenmStatus = stConsulta Then pPonEstado stedicion
        Else
            '|  Esta cuenta ya está registrada
            MsgBox SIHOMsg(266), vbCritical, "Mensaje"
        End If
        pHabilitaBotonesPaso
    End With
End Sub

Private Sub cmdAnterior_Click()
    If Not vgrsGrupoFacturas.BOF Then
        If vgenmStatus <> stEspera Then vgrsGrupoFacturas.MovePrevious
        If vgrsGrupoFacturas.BOF Then
            vgrsGrupoFacturas.MoveNext
            txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
            txtCveGrupo_KeyDown vbKeyReturn, 0
            cmdInicio.Enabled = False
            cmdAnterior.Enabled = False
        Else
            cmdSiguiente.Enabled = True
            cmdFin.Enabled = True
            txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
            txtCveGrupo_KeyDown vbKeyReturn, 0
        End If
    End If
End Sub

Private Sub cmdBuscaGrupos_Click()
    
    Dim strParametros As String
    Dim rsGrupos As New ADODB.Recordset
    
    strParametros = IIf(optTipoGrupoB(0).Value, IIf(cboEmpresa.ListIndex = 0, "-2", cboEmpresa.ItemData(cboEmpresa.ListIndex)), IIf(optTipoGrupoB(1).Value, "-1", "-3")) & "|" & _
                    IIf(chkRangoFechasB.Value, fstrFechaSQL(dtpFechaInicialB.Value, "00:00:00"), "") & "|" & _
                    fstrFechaSQL(dtpFechaFinalB.Value, "23:59:59") & "|" & _
                    IIf(txtFolioFacturaB.Text <> "", "1", IIf(chkFacturados.Value, "1", "0")) & "|" & _
                    IIf(optTipoPacB(2).Value, "-1", txtCtaPacB.Text) & "|" & _
                    IIf(optTipoPacB(0).Value, "I", "E") & "|" & _
                    IIf(txtFolioFacturaB.Text = "", "-1", txtFolioFacturaB.Text) & "|" & _
                    vgintClaveEmpresaContable
                    
    Set rsGrupos = frsEjecuta_SP(strParametros, "Sp_Pvselconsultagrupos")
    MSFGResultado.Clear
    If rsGrupos.RecordCount > 0 Then
        pLlenaVsfGrid MSFGResultado, rsGrupos
        pConfiguraGridResultadoBusqueda False
    Else
        pConfiguraGridResultadoBusqueda True
    End If

End Sub

Private Sub cmdBuscar_Click()
    SSTFacturacionConsolidada.Tab = 2
    cmdBuscaGrupos.SetFocus
End Sub

Private Sub cmdEjecutaConsulta_Click()
    Dim vlrsCuentasPorEmpresa As New ADODB.Recordset
    Dim vlstrParametros As String
    Dim vlstrTipoCargo As String
    Dim vldtmFechaIni As String
    Dim vldtmFechaFin As String
    Dim vlblnEncontrado As Boolean
    Dim intCont As Integer
        
    Screen.MousePointer = vbHourglass
    Select Case cboTipoCargo.Text
        Case "<TODOS>"
            vlstrTipoCargo = "*"
        Case "ARTÍCULO"
            vlstrTipoCargo = "AR"
        Case "ESTUDIO"
            vlstrTipoCargo = "ES"
        Case "EXAMEN"
            vlstrTipoCargo = "EX"
        Case "GRUPO DE EXÁMENES"
            vlstrTipoCargo = "GE"
        Case "OTROS CONCEPTOS"
            vlstrTipoCargo = "OC"
    End Select
    
    If chkRangoFechas.Value Then
        vldtmFechaIni = dtpFechaIni
        vldtmFechaFin = dtpFechaFin
    Else
        vldtmFechaIni = "01/01/1900"
        vldtmFechaFin = "01/01/1900"
    End If
    
    vlstrParametros = cboConvenio.ItemData(cboConvenio.ListIndex) & "|" & _
                      IIf(optTipoPac(0).Value, "I", IIf(optTipoPac(1).Value, "E", "*")) & "|" & _
                      IIf(Not optTipoPac(0).Value And Not optTipoPac(1).Value, "-1", IIf(txtNumCuenta.Text = "", "-1", txtNumCuenta.Text)) & "|" & _
                      vldtmFechaIni & "|" & _
                      vldtmFechaFin & "|" & _
                      chkIncluirCargosAtrasados.Value & "|" & _
                      IIf(cboConceptoFacturacion.ListIndex = 0, -1, cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex)) & "|" & _
                      vlstrTipoCargo & "|" & _
                      vgintClaveEmpresaContable
    Set vlrsCuentasPorEmpresa = frsEjecuta_SP(vlstrParametros, "Sp_PvSelCuentasDeEmpresa")
    If vlrsCuentasPorEmpresa.RecordCount > 0 Then
        pLimpiaMshFGd MSFGCargosDisponibles
        pConfiguraGridCargos MSFGCargosDisponibles
        MSFGCargosDisponibles.Row = 1
        pConfiguraGridCuentas MSFGCuentasDisponiblesGrupo
        intCont = 1
        While Not vlrsCuentasPorEmpresa.EOF
            MSFGCuentasDisponiblesGrupo.Rows = MSFGCuentasDisponiblesGrupo.Rows + 1
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 0) = vlrsCuentasPorEmpresa!NumCta
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 1) = vlrsCuentasPorEmpresa!TipoPac
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 2) = vlrsCuentasPorEmpresa!Nombre
            intCont = intCont + 1
            vlrsCuentasPorEmpresa.MoveNext
        Wend
        'Significa que la consulta no está vacía
        If (MSFGCuentasDisponiblesGrupo.TextMatrix(1, 1) <> "") Then
            MSFGCuentasDisponiblesGrupo.Rows = MSFGCuentasDisponiblesGrupo.Rows - 1
            MSFGCuentasDisponiblesGrupo.Row = 1
            MSFGCuentasDisponiblesGrupo.Col = 2
            MSFGCuentasDisponiblesGrupo.SetFocus
            txtEmpresa.Text = cboConvenio.Text
            vgintCveEmpresa = cboConvenio.ItemData(cboConvenio.ListIndex)
        Else
            '|  ¡No existe información!
            MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
        End If
    Else
        MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
    End If
    pHabilitaBotonesPaso
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEjecutarConsultaP_Click()
    Dim strParametros As String
    Dim rsCuentas As New ADODB.Recordset
    Dim intCont As Integer
    
    Screen.MousePointer = vbHourglass
    strParametros = fstrFechaSQL(dtpFechaInicialP.Value, "00:00:00") & "|" & _
                    fstrFechaSQL(dtpFechaFinalP.Value, "23:59:59") & "|" & _
                    IIf(optTipoPacP(0).Value, "I", IIf(optTipoPacP(1).Value, "E", "A")) & "|" & _
                    IIf(cboConceptoFacturacionP.ListIndex = 0, -1, cboConceptoFacturacionP.ItemData(cboConceptoFacturacionP.ListIndex)) & "|" & _
                    vgintClaveEmpresaContable
    Set rsCuentas = frsEjecuta_SP(strParametros, "Sp_Pvselcuentasdepaciente")
    If rsCuentas.RecordCount > 0 Then
        pLimpiaMshFGd MSFGCargosDisponibles
        pConfiguraGridCargos MSFGCargosDisponibles
        MSFGCargosDisponibles.Row = 1
        pConfiguraGridCuentas MSFGCuentasDisponiblesGrupo
        intCont = 1
        While Not rsCuentas.EOF
            MSFGCuentasDisponiblesGrupo.Rows = MSFGCuentasDisponiblesGrupo.Rows + 1
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 0) = rsCuentas!NumCta
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 1) = rsCuentas!TipoPac
            MSFGCuentasDisponiblesGrupo.TextMatrix(intCont, 2) = rsCuentas!Nombre
            intCont = intCont + 1
            rsCuentas.MoveNext
        Wend
        'Significa que la consulta no está vacía
        If (MSFGCuentasDisponiblesGrupo.TextMatrix(1, 1) <> "") Then
            MSFGCuentasDisponiblesGrupo.Rows = MSFGCuentasDisponiblesGrupo.Rows - 1
            MSFGCuentasDisponiblesGrupo.Row = 1
            MSFGCuentasDisponiblesGrupo.Col = 2
            MSFGCuentasDisponiblesGrupo.SetFocus
        Else
            '|  ¡No existe información!
            MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
        End If
    Else
        MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
    End If
    pHabilitaBotonesPaso
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdEliminar_Click()

On Error GoTo NotificaError
    '|  ¿Está seguro de eliminar los datos?
    If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        '"Elimina" el detalle del grupo
        pEjecutaSentencia "UPDATE PvCargo SET PvCargo.INTCVEGRUPO = NULL WHERE PvCargo.INTCVEGRUPO = " & vgrsGrupoFacturas!intCveGrupo
        'Elimina Cuentas
        pEjecutaSentencia "DELETE FROM PVDETALLEFACTURACONSOLID WHERE INTCVEGRUPO = " & vgrsGrupoFacturas!intCveGrupo
        'Elimina el maestro
        vgrsGrupoFacturas.Delete
        vgrsGrupoFacturas.Update
        vgrsGrupoFacturas.Requery
        EntornoSIHO.ConeccionSIHO.CommitTrans
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "FACTURACION CONSOLIDADA", txtCveGrupo.Text)
        txtCveGrupo.SetFocus
        pPonEstado stEspera
    End If
    Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
End Sub

Private Sub cmdFin_Click()
    vgrsGrupoFacturas.MoveLast
    txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
    txtCveGrupo_KeyDown vbKeyReturn, 0
    cmdSiguiente.Enabled = False
    cmdFin.Enabled = False
    cmdInicio.Enabled = True
    cmdAnterior.Enabled = True
End Sub

Private Sub cmdGrabar_Click()

Dim vllngCveGrupo As Long
Dim vlintRenglon As Integer
Dim vlrsGrupo As New ADODB.Recordset
Dim vllngPersonaGraba As Long
Dim vlrsPvFacturacionConsolidada As New ADODB.Recordset
Dim intCont As Long
Dim vlblnEsta As Boolean
Dim llngVariable As Long

On Error GoTo NotificaError
    If msfgCargosAsignados.TextMatrix(1, 1) = "" Then
        '|  ¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbCritical, "Mensaje"
        msfgCargosAsignados.SetFocus
        Exit Sub
    End If
    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        If vgenmStatus = stNuevo Then
            'Inserta el grupo
            Set vlrsPvFacturacionConsolidada = frsRegresaRs("SELECT * FROM PvFacturacionConsolidada WHERE INTCVEGRUPO = -1", adLockOptimistic, adOpenDynamic)
            With vlrsPvFacturacionConsolidada
                .AddNew
                !intcveempresa = vgintCveEmpresa
                !DTMFECHACREACION = fdtmServerFecha
                !tnyClaveEmpresa = vgintClaveEmpresaContable
                .Update
                txtCveGrupo.Text = flngObtieneIdentity("SEC_PVFACTURACIONCONSOLIDADA", !intCveGrupo)
                vllngCveGrupo = txtCveGrupo.Text
            End With
        End If
        '---------------------------------------------------------------------------
        '|  Libera los cargos asignados al grupo
        '---------------------------------------------------------------------------
        pEjecutaSentencia "UPDATE PvCargo SET PvCargo.INTCVEGRUPO = null WHERE PvCargo.INTCVEGRUPO = " & txtCveGrupo.Text
        '---------------------------------------------------------------------------
        '|  Asocia los cargos seleccionados al grupo
        '---------------------------------------------------------------------------
        For vlintRenglon = 1 To msfgCargosAsignados.Rows - 1
            pEjecutaSentencia "UPDATE PvCargo SET PvCargo.INTCVEGRUPO = " & txtCveGrupo.Text & " WHERE PvCargo.INTNUMCARGO = " & msfgCargosAsignados.RowData(vlintRenglon)
        Next vlintRenglon
        '---------------------------------------------------------------------------
        '|  Elimina el detalle del grupo
        '---------------------------------------------------------------------------
        pEjecutaSentencia "DELETE FROM PVDETALLEFACTURACONSOLID WHERE INTCVEGRUPO = " & txtCveGrupo.Text
        '---------------------------------------------------------------------------
        '|  Guarda las cuentas de pacientes que pertenecen al grupo
        '---------------------------------------------------------------------------
        For intCont = 1 To MSFGCuentasAsignadasGrupo.Rows - 1
            pEjecutaSentencia "INSERT INTO PVDETALLEFACTURACONSOLID ( INTCVEGRUPO, INTMOVPACIENTE,CHRTIPOPACIENTE ) VALUES ( " & txtCveGrupo.Text & ", " & MSFGCuentasAsignadasGrupo.TextMatrix(intCont, 0) & ", '" & MSFGCuentasAsignadasGrupo.TextMatrix(intCont, 1) & "')"
            vgstrParametrosSP = Trim(MSFGCuentasAsignadasGrupo.TextMatrix(intCont, 0)) & "|" & MSFGCuentasAsignadasGrupo.TextMatrix(intCont, 1) & "|0|"
            frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDCERRARABRIRCUENTA"
        Next intCont

        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        llngVariable = 1
        frsEjecuta_SP CLng(txtCveGrupo.Text) & "|1", "FN_UPDCARGOSPAQUETESFUERAGRUPO", True, llngVariable
        If llngVariable <> 0 Then
            '¡Los datos han sido guardados satisfactoriamente!
            MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
        
            'Se agregaron al grupo automáticamente los cargos asignados a los paquetes que se facturarán en el grupo, consulte de nuevo.
            MsgBox SIHOMsg(1583), vbOKOnly + vbInformation, "Mensaje"
        Else
            '¡Los datos han sido guardados satisfactoriamente!
            MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
        End If
   
        vgrsGrupoFacturas.Requery
        vgrsGrupoFacturas.MoveLast
        txtCveGrupo.Text = ""
        pPonEstado stEspera
    End If
    Exit Sub
NotificaError:
    EntornoSIHO.ConeccionSIHO.RollbackTrans
End Sub

Private Sub cmdInicio_Click()
    vgrsGrupoFacturas.MoveFirst
    txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
    txtCveGrupo_KeyDown vbKeyReturn, 0
    cmdSiguiente.Enabled = True
    cmdFin.Enabled = True
    cmdInicio.Enabled = False
    cmdAnterior.Enabled = False
End Sub

Private Sub cmdPrint_Click()

 On Error GoTo NotificaError
 
    pImprime "I"

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdImprimir_Click"))
End Sub
Private Sub pImprime(strDestino As String)
    Dim rs As New ADODB.Recordset
    Dim rsgrupo As New ADODB.Recordset
    Dim alstrParametros(4) As String
    
    pInstanciaReporte vgrptReporte, "rptGrupoFacturas.rpt"
    
    Set rsgrupo = frsEjecuta_SP(CStr(txtCveGrupo.Text), "Sp_PvSelGrupoFacturas")
    vgstrParametrosSP = CStr(txtCveGrupo.Text) & "|" & IIf(IsNull(rsgrupo!Folio), "0|-1", "1|" & Trim(rsgrupo!Folio)) & "|" & chkCosto.Value
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptGrupoCuentas")
    If rs.EOF Then
        'No existe información con esos parametros
        MsgBox SIHOMsg(236), vbInformation, "Mensaje"
    Else
        vgrptReporte.DiscardSavedData
        alstrParametros(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
        alstrParametros(1) = "Grupo;" & CStr(txtCveGrupo.Text)
        alstrParametros(2) = "Empresa;" & IIf(IsNull(rsgrupo!empresa), "PARTICULARES", rsgrupo!empresa)
        alstrParametros(3) = "Creacion;" & IIf(IsNull(rsgrupo!Fecha), "", Format(rsgrupo!Fecha, "DD/MMM/YYYY"))
        alstrParametros(4) = "Foliofactura;" & rsgrupo!Folio
        pCargaParameterFields alstrParametros, vgrptReporte
        pImprimeReporte vgrptReporte, rs, strDestino, "Facturación consolidada"
    End If
    rs.Close
    
End Sub


Private Sub cmdQuitaTodos_Click()
    Dim intCont As Integer
    
    If MSFGCuentasAsignadasGrupo.TextMatrix(1, 0) <> "" Then
        MSFGCuentasAsignadasGrupo.Row = 1
        For intCont = 0 To MSFGCuentasAsignadasGrupo.Rows - 2
            cmdQuitaUno_Click
        Next
    End If

End Sub

Private Sub cmdQuitaUno_Click()
    If MSFGCuentasAsignadasGrupo.Row > 0 Then
        pseleccioncuenta MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 0), MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 1)
        pMueveCuenta MSFGCuentasAsignadasGrupo, MSFGCuentasDisponiblesGrupo
        If chkMuestraDetalleAutomaticamente.Value = vbChecked And MSFGCuentasAsignadasGrupo.Row > 0 Then
            MSFGCuentasAsignadasGrupo_Click
        End If
    End If
   
End Sub

Private Sub cmdSiguiente_Click()
    If Not vgrsGrupoFacturas.EOF Then
        If vgrsGrupoFacturas.BOF Then vgrsGrupoFacturas.MoveNext
        vgrsGrupoFacturas.MoveNext
        If vgrsGrupoFacturas.EOF Then
            vgrsGrupoFacturas.MovePrevious
            txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
            txtCveGrupo_KeyDown vbKeyReturn, 0
            cmdSiguiente.Enabled = False
            cmdFin.Enabled = False
        Else
            cmdInicio.Enabled = True
            cmdAnterior.Enabled = True
            txtCveGrupo.Text = vgrsGrupoFacturas!intCveGrupo
            txtCveGrupo_KeyDown vbKeyReturn, 0
        End If
    End If
End Sub



Private Sub cmdVerDetalleCargos_Click()
    If MSFGCuentasAsignadasGrupo.Row > 0 Then
        If MSFGCuentasAsignadasGrupo.TextMatrix(1, 0) = "" Then Exit Sub
        '|  Si esta seleccionado un renglón del grid se los asignados muestra su detalle
       
        pLlenaCargos MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 0), _
                     MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 1), _
                     IIf(txtDGFolioFactura.Text = "", "C", "G"), _
                     MSFGCargosDisponibles, _
                     MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 3)
        pPonEtiqueta MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 0) & " " & _
                     MSFGCuentasAsignadasGrupo.TextMatrix(MSFGCuentasAsignadasGrupo.Row, 1)
        pmarcados
     
    Else
        If MSFGCuentasDisponiblesGrupo.TextMatrix(1, 0) = "" Then Exit Sub
        '|  Si esta seleccionado un renglón del grid se los disponibles muestra su detalle
        If MSFGCuentasDisponiblesGrupo.Row > 0 Then
            pLlenaCargos Val(MSFGCuentasDisponiblesGrupo.TextMatrix(MSFGCuentasDisponiblesGrupo.Row, 0)), _
                         MSFGCuentasDisponiblesGrupo.TextMatrix(MSFGCuentasDisponiblesGrupo.Row, 1), _
                         IIf(txtDGFolioFactura.Text = "", "C", "G"), _
                         MSFGCargosDisponibles, _
                         MSFGCuentasDisponiblesGrupo.TextMatrix(MSFGCuentasDisponiblesGrupo.Row, 3)
            pPonEtiqueta MSFGCuentasDisponiblesGrupo.TextMatrix(MSFGCuentasDisponiblesGrupo.Row, 0) & " " & _
                         MSFGCuentasDisponiblesGrupo.TextMatrix(MSFGCuentasDisponiblesGrupo.Row, 1)
                         
        End If
    End If
    
End Sub

Private Sub cmdVistaPreliminar_Click()
   pImprime "P"
End Sub

Private Sub dtpFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkIncluirCargosAtrasados.SetFocus
End Sub

Private Sub dtpFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpFechaFin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    vlintEmpresa = 0
    vlintCveConcepto = 0
    vldblPorcIVA = 0
    vldblMontoIVA = 0
    vlstrCveCargo = " "
    vlstrTipoCargoEmpresa = " "
    vlstrDescripcionCargo = " "
    Select Case KeyCode
        Case vbKeyReturn
            If Me.ActiveControl.Name <> "txtCveGrupo" Then SendKeys vbTab
        Case vbKeyEscape
            Select Case vgenmStatus
                Case stedicion, stNuevo
                    '|  ¿Desea abandonar la operación?
                    If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        pInicializaForma
                        pPonEstado stEspera
                        KeyCode = 0
                    End If
                Case stConsulta, stFacturado
                    pPonEstado stEspera
                Case stEspera
                    If SSTFacturacionConsolidada.Tab = 2 Then
                        pPonEstado stEspera
                        KeyCode = 0
                    Else
                        Unload Me
                    End If
            End Select
    End Select

End Sub


 
Private Sub Form_Load()
    Dim vlrsCarga As New ADODB.Recordset
    Dim strSQL As String
    Dim rsEmpresas As New ADODB.Recordset
    
    
    vlintEmpresa = 0
    vlintCveConcepto = 0
    vldblPorcIVA = 0
    vldblMontoIVA = 0
    vlstrCveCargo = " "
    vlstrTipoCargoEmpresa = " "
    vlstrDescripcionCargo = " "
    
    Me.Icon = frmMenuPrincipal.Icon
    '------------------------------------------------------
    '|           D A T O S    G E N E R A L E S
    '------------------------------------------------------
    SSTFacturacionConsolidada.Tab = 0
    'Carga RecordSet principal
    Set vgrsGrupoFacturas = frsRegresaRs("SELECT INTCVEGRUPO FROM PvFacturacionConsolidada ORDER BY INTCVEGRUPO", adLockOptimistic, adOpenDynamic)
    If vgrsGrupoFacturas.RecordCount = 0 Then
        cmdSiguiente.Enabled = False
        cmdAnterior.Enabled = False
        cmdBuscar.Enabled = False
        cmdFin.Enabled = False
        cmdInicio.Enabled = False
    End If
    '------------------------------------------------------
    '|         C O N T E N I D O   D E   G R U P O S
    '------------------------------------------------------
    '|  Carga combo de los convenios
    strSQL = "SELECT CcEmpresa.INTCVEEMPRESA, CcEmpresa.VCHDESCRIPCION FROM CcEmpresa"
    Set vlrsCarga = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboConvenio, vlrsCarga, 0, 1
    vgintCveEmpresa = IIf(vlrsCarga.RecordCount > 0, cboConvenio.ItemData(0), -1)
    '|  Carga combo de los conceptos de facturación
    strSQL = "SELECT PvConceptoFacturacion.SMICVECONCEPTO, RTRIM(PvConceptoFacturacion.CHRDESCRIPCION) as DESCRIPCION FROM PvConceptoFacturacion"
    Set vlrsCarga = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboConceptoFacturacion, vlrsCarga, 0, 1, 3
    pLlenarCboRs cboConceptoFacturacionP, vlrsCarga, 0, 1, 3
    '------------------------------------------------------
    '|         C O N S U L T A    D E    G R U P O S
    '------------------------------------------------------
    '|  Carga combo con las empresas que tengan grupos de facturas
    strSQL = "SELECT distinct CcEmpresa.INTCVEEMPRESA as Clave " & _
             "     , CcEmpresa.VCHDESCRIPCION as Nombre " & _
             "  FROM PvFacturacionConsolidada " & _
             "       INNER JOIN CcEmpresa ON (PvFacturacionConsolidada.INTCVEEMPRESA = CcEmpresa.INTCVEEMPRESA) "
    Set rsEmpresas = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresa, rsEmpresas, 0, 1, 3
    cboEmpresa.ListIndex = 0
    rsEmpresas.Close
    '| Inicializa fechas
    dtpFechaFinalB.Value = fdtmServerFecha
    dtpFechaInicialB.Value = fdtmServerFecha
    pPonEstado stEspera
    cmdPrint.Enabled = False
    cmdVistaPreliminar.Enabled = False
    
    
End Sub

Private Sub MSFGCargosDisponibles_DblClick()
pEnmarcaDesmarca
End Sub

Private Sub MSFGCuentasAsignadasGrupo_Click()
 
    '|  Quita la selección del grid MSFGCuentasDisponiblesGrupo
    MSFGCuentasDisponiblesGrupo.Row = -1
    With MSFGCuentasAsignadasGrupo
        '|  Si el renglón seleccionado no está vacío Y Se va a mostrar el detalle automáticamente
        If MSFGCuentasAsignadasGrupo.TextMatrix(1, 1) <> "" _
           And chkMuestraDetalleAutomaticamente.Value Then
            pLlenaCargos .TextMatrix(.Row, 0), _
                         .TextMatrix(.Row, 1), _
                         "C", _
                         MSFGCargosDisponibles, _
                         .TextMatrix(.Row, 3)
            pPonEtiqueta .TextMatrix(.Row, 0) & " " & .TextMatrix(.Row, 1)
        End If
    End With
    pmarcados
End Sub

Private Sub MSFGCuentasAsignadasGrupo_RowColChange()
    MSFGCuentasAsignadasGrupo_Click
End Sub

Private Sub MSFGCuentasDisponiblesGrupo_Click()
    With MSFGCuentasDisponiblesGrupo
        If .Row > 0 Then
            MSFGCuentasAsignadasGrupo.Row = -1
            If chkMuestraDetalleAutomaticamente.Value Then
                pLlenaCargos Val(.TextMatrix(.Row, 0)), _
                             .TextMatrix(.Row, 1), _
                             "C", _
                             MSFGCargosDisponibles, _
                             .TextMatrix(.Row, 3)
                pPonEtiqueta .TextMatrix(.Row, 0) & " " & .TextMatrix(.Row, 1)
            End If
        End If
    End With
End Sub

Private Sub MSFGCuentasDisponiblesGrupo_RowColChange()
    MSFGCuentasDisponiblesGrupo_Click
End Sub

Private Sub MSFGResultado_DblClick()
    'Evalúa si escogió un grupo
    If MSFGResultado.Row > 0 Then
        txtCveGrupo.Text = MSFGResultado.TextMatrix(MSFGResultado.Row, 0)
        txtCveGrupo_KeyDown vbKeyReturn, 0
        SSTFacturacionConsolidada.Tab = 0
    End If
End Sub

Private Sub optTipoGrupo_Click(Index As Integer)
    If Index = 0 Then
        fraFiltrosContenidoGrupoE.Visible = True
        fraFiltrosContenidoGrupoP.Visible = False
    Else
        fraFiltrosContenidoGrupoP.Visible = True
        fraFiltrosContenidoGrupoE.Visible = False
    End If
    txtCveGrupo.SetFocus
End Sub

Private Sub optTipoGrupoB_Click(Index As Integer)
    Select Case Index
        Case 0
            cboEmpresa.Enabled = True
            pEnfocaCbo cboEmpresa
        Case 1, 2
            cboEmpresa.Enabled = False
    End Select
End Sub

Private Sub optTipoPac_Click(Index As Integer)
    Select Case Index
        Case 0, 1
            txtNumCuenta.Enabled = True
            txtNumCuenta.SetFocus
        Case 2
            txtNumCuenta.Text = ""
            txtNumCuenta.Enabled = False
    End Select
End Sub

Private Sub optTipoPac_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtNumCuenta.Text = ""
        chkRangoFechas.SetFocus
        If fblnCanFocus(chkRangoFechas) Then chkRangoFechas.SetFocus
    End If
End Sub

Private Sub optTipoPacB_Click(Index As Integer)
    Select Case Index
        Case 0, 1
            txtCtaPacB.Enabled = True
            txtCtaPacB.SetFocus
        Case 2
            txtCtaPacB.Text = ""
            txtCtaPacB.Enabled = False
    End Select
End Sub

Private Sub SSTFacturacionConsolidada_Click(PreviousTab As Integer)
    Select Case SSTFacturacionConsolidada.Tab
        Case 0
            If vgenmStatus = stNuevo Or vgenmStatus = stedicion Then
                pTotales 'pLlenaCargosAsignados
            End If
        Case 1
            If fblnCanFocus(cboConvenio) Then cboConvenio.SetFocus
            If vgenmStatus = stedicion Or vgenmStatus = stConsulta Then
                cboConvenio.Enabled = False
                cboConvenio.ListIndex = fintLocalizaCbo(cboConvenio, CStr(vgintCveEmpresa))
            Else
                If vgintCveEmpresa = 0 Then
                    cboConvenio.ListIndex = 0
                Else
                    cboConvenio.ListIndex = fintLocalizaCbo(cboConvenio, CStr(vgintCveEmpresa))
                End If
                cboConvenio.Enabled = True
            End If
    End Select
End Sub

Private Sub txtCtaPacB_GotFocus()
    pEnfocaTextBox txtCtaPacB
End Sub

Private Sub txtCtaPacB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If RTrim(txtCtaPacB.Text) = "" Then
            With FrmBusquedaPacientes
                If optTipoPacB(1).Value Then 'Externos
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = True
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                ElseIf optTipoPacB(0).Value Then
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtCtaPacB.Text = .flngRegresaPaciente()
                
                If txtCtaPacB.Text <> -1 Then
                    txtCtaPacB_KeyDown vbKeyReturn, 0
                Else
                    txtCtaPacB.Text = ""
                End If
                .txtBusqueda.CausesValidation = True
            End With
        End If
    End If

End Sub

Private Sub txtCtaPacB_KeyPress(KeyAscii As Integer)
    '|  Si la tecla presionada no es un número
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        '|  Valida si la tecla presionada es una "E" o "P" para cambiar el tipo de grupo
        Select Case UCase(Chr(KeyAscii))
            Case "E"
                optTipoPacB(1).Value = True
            Case "I"
                optTipoPacB(0).Value = True
            Case "A"
                optTipoPacB(2).Value = True
        End Select
        KeyAscii = 7
    End If
End Sub

Private Sub txtCtaPacB_LostFocus()
    If txtCtaPacB.Text = "" Then optTipoPacB(2).Value = True
End Sub


Private Sub txtCveGrupo_GotFocus()
    pEnfocaTextBox txtCveGrupo
End Sub

Private Sub txtCveGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        '|  Si la clave del grupo esta vacía se asume que desea buscar un grupo
        If txtCveGrupo.Text = "" Then
            cmdBuscar_Click
        Else
            '|  Si se introdujo una clave de grupo de valida que el grupo exista
            If fblnExisteGrupo(CDbl(txtCveGrupo.Text)) Then
                pDespliegaGrupo txtCveGrupo.Text
                cmdPrint.Enabled = True
                cmdVistaPreliminar.Enabled = True
            Else '|  Si la clave no existe se inicia la captura de un nuevo grupo
                cmdPrint.Enabled = False
                cmdVistaPreliminar.Enabled = False
                pPonEstado stNuevo
                SSTFacturacionConsolidada.Tab = 1
                '|  Si el grupo es de particulares
                If optTipoGrupo(1).Value Then
                    txtEmpresa.Text = "PARTICULAR"
                    vgintCveEmpresa = -1
                    cmdEjecutarConsultaP_Click
                    
                End If
            End If
        End If
    End If

End Sub

Private Sub txtCveGrupo_KeyPress(KeyAscii As Integer)
    '|  Si la tecla presionada no es un número
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        '|  Valida si la tecla presionada es una "E" o "P" para cambiar el tipo de grupo
        Select Case UCase(Chr(KeyAscii))
            Case "E"
                optTipoGrupo(0).Value = True
            Case "P"
                optTipoGrupo(1).Value = True
        End Select
        KeyAscii = 7
    End If
End Sub

Private Sub txtFolioFacturaB_GotFocus()
    pEnfocaTextBox txtFolioFacturaB
End Sub

Private Sub txtFolioFacturaB_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNumCuenta_GotFocus()
    pEnfocaTextBox txtNumCuenta
End Sub

Private Sub txtNumCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If RTrim(txtNumCuenta.Text) = "" Then
            With FrmBusquedaPacientes
                If optTipoPac(1).Value Then 'Externos
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = True
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                ElseIf optTipoPac(0).Value Then
                    .txtBusqueda.CausesValidation = False
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtNumCuenta.Text = .flngRegresaPaciente()
                
                If txtNumCuenta <> -1 Then
                    txtNumCuenta_KeyDown vbKeyReturn, 0
                Else
                    txtNumCuenta.Text = ""
                End If
                .txtBusqueda.CausesValidation = True
            End With
        End If
    End If
End Sub

Private Sub txtNumCuenta_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optTipoPac(0).Value = UCase(Chr(KeyAscii)) = "I"
            optTipoPac(1).Value = UCase(Chr(KeyAscii)) = "E"
            optTipoPac(2).Value = UCase(Chr(KeyAscii)) = "A"
        End If
        KeyAscii = 7
    End If
End Sub

Private Sub pLlenaCargos(pintNumCta As Long, _
                         pstrTipoPac As String, _
                         pstrTipoCuenta As String, _
                         msfgGrid As Control, _
                         strPerteneceGrupo As String, _
                         Optional intAgregaAlFinal As Integer = 0)
    Dim vlstrSQLExterno As String
    Dim vlintContador As Integer
    Dim rsDescuentos As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim vlstrTipoDescuento As String
    Dim vldblSubtotal As Double
    Dim vldblDescuento As Double
    Dim vldblTotDescuento As Double
    Dim vldblIVA As Double
    Dim vldbltotal As Double
    Dim vlContadorColumnas As Integer
    Dim vlTemp As Variant
    Dim vlrsSeleccionaCargos As New ADODB.Recordset
    Dim vlstrTipoCargo As String
    Dim vldtmFechainicial As Date
    Dim vlstrSentencia As String
    Dim vllngCveConceptoFacturacion As Long
    Dim rs As New ADODB.Recordset
    
    Select Case cboTipoCargo.Text
        Case "<TODOS>"
            vlstrTipoCargo = "*"
        Case "ARTÍCULO"
            vlstrTipoCargo = "AR"
        Case "ESTUDIO"
            vlstrTipoCargo = "ES"
        Case "EXAMEN"
            vlstrTipoCargo = "EX"
        Case "GRUPO DE EXÁMENES"
            vlstrTipoCargo = "GE"
        Case "OTROS CONCEPTOS"
            vlstrTipoCargo = "OC"
    End Select
    
    Me.MousePointer = 11
    
    If intAgregaAlFinal = 0 Then
        pConfiguraGridCargos msfgGrid, True
        msfgGrid.Row = 1
    Else
        msfgGrid.Row = msfgGrid.Rows - 1
        If msfgGrid.TextMatrix(msfgGrid.Row, 2) <> "" Then msfgGrid.Rows = msfgGrid.Rows + 1
        msfgGrid.Row = msfgGrid.Rows - 1
    End If
    msfgGrid.Redraw = False
 
    vlintEmpresa = 0
    vlintCveConcepto = 0
    vldblPorcIVA = 0
    vldblMontoIVA = 0
    vlstrCveCargo = " "
    vlstrTipoCargoEmpresa = " "
    vlstrDescripcionCargo = " "
    vldtmFechainicial = IIf(chkRangoFechas.Value And chkIncluirCargosAtrasados.Value, CDate("01/01/1900"), dtpFechaIni.Value)
    
    Set vlrsSeleccionaCargos = frsEjecuta_SP(pintNumCta & "|" & pstrTipoPac & "|0|-1|" & pstrTipoCuenta & IIf(strPerteneceGrupo = "*", "|S", "|N") & "|0", "SP_PVSELCARGOSPACIENTEGRUPOS")
    With vlrsSeleccionaCargos
    
        'If .EOF And intAgregaAlFinal = 0 Then pConfiguraGridCargos msfgCargosAsignados
        Do While Not .EOF
            If (vlrsSeleccionaCargos!smicveconcepto = cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex) _
                 Or cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex) = 0) _
               And (IIf(vlrsSeleccionaCargos!chrTipoCargo = "GE", "EX", vlrsSeleccionaCargos!chrTipoCargo) = vlstrTipoCargo Or vlstrTipoCargo = "*") _
               And ((vlrsSeleccionaCargos!DTMFECHAHORA > vldtmFechainicial And vlrsSeleccionaCargos!DTMFECHAHORA < CDate(CStr(dtpFechaFin.Value) & " 23:59:59")) Or (chkRangoFechas.Value = False)) Then
            
                If ((IsNull(!Excluido) Or !Excluido = "") _
                     And (IsNull(!FolioFactura) Or RTrim(!FolioFactura) = "")) Then
                     
                     
                    '-----------------------------------------------------------------------------------------------------------------
                    ' Trae los datos de los cargos de la empresa
                    '-----------------------------------------------------------------------------------------------------------------
                    
                    vllngCveConceptoFacturacion = !smicveconcepto
                    
                    vlstrCveCargo = !chrCveCargo
                    vlstrTipoCargoEmpresa = !chrTipoCargo
                    
                    vlintEmpresa = IIf(IsNull(!ClaveEmpresaPaciente), 0, !ClaveEmpresaPaciente)
                    
'                    vlstrSentencia = "SELECT smicveconcepto, vchdescripcion FROM CCCATALOGOPCE " & _
'                                     " WHERE chrcvecargo = '" & Trim(vlstrCveCargo) & "'" & _
'                                     "   AND chrtipocargo = '" & Trim(vlstrTipoCargoEmpresa) & "'" & _
'                                     "   AND intCveEmpresa = " & vlintEmpresa
'                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                    If rs.RecordCount > 0 Then

                        vlintCveConcepto = 0
                        vlintCveConcepto = IIf(IsNull(!ClaveConceptoPCE) Or !ClaveConceptoPCE = 0, vllngCveConceptoFacturacion, !ClaveConceptoPCE)
                        
                        vlstrDescripcionCargo = " "
                        vlstrDescripcionCargo = IIf(IsNull(!DescripcionConceptoPCE) Or !DescripcionConceptoPCE = "", !DescripcionCargosSoloPVCARGO, !DescripcionConceptoPCE)
                
'                    Else
'                        Select Case vlstrTipoCargoEmpresa
'                            Case "AR"
'                                      vlstrSentencia = "SELECT SMICVECONCEPTFACT smicveconcepto, VCHNOMBRECOMERCIAL vchdescripcion FROM ivarticulo " & _
'                                                       " WHERE intidarticulo = " & (CLng(Trim(vlstrCveCargo)))
'                            Case "OC"
'                                      vlstrSentencia = "SELECT SMICONCEPTOFACT smicveconcepto, CHRDESCRIPCION vchdescripcion FROM pvotroconcepto " & _
'                                                       " WHERE intcveconcepto = " & (CLng(Trim(vlstrCveCargo)))
'                            Case "ES"
'                                      vlstrSentencia = "SELECT SMICONFACT smicveconcepto, VCHNOMBRE vchdescripcion FROM imestudio " & _
'                                                       " WHERE intcveestudio = " & (CLng(Trim(vlstrCveCargo)))
'                            Case "EX"
'                                      vlstrSentencia = "SELECT SMICONFACT smicveconcepto, CHRNOMBRE vchdescripcion FROM laexamen " & _
'                                                       " WHERE intcveexamen = " & (CLng(Trim(vlstrCveCargo)))
'                            Case "GE"
'                                      vlstrSentencia = "SELECT SMICONFACT smicveconcepto, CHRNOMBRE vchdescripcion FROM lagrupoexamen " & _
'                                                       " WHERE intcvegrupo = " & (CLng(Trim(vlstrCveCargo)))
'                        End Select
'                        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                        If rs.RecordCount > 0 Then
''                            vlintCveConcepto = rs!smiCveConcepto
'                            vlintCveConcepto = vllngCveConceptoFacturacion
'                            vlstrDescripcionCargo = rs!VCHDESCRIPCION
'                        Else
'                            vlstrDescripcionCargo = " "
'                        End If
'                    End If
'                    rs.Close
                    
'                    vlstrSentencia = "SELECT smyIva, chrdescripcion FROM pvConceptoFacturacion " & _
'                                     " WHERE smiCveConcepto = " & vlintCveConcepto
'                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                    If rs.RecordCount > 0 Then
                    
                        vldblPorcIVA = 0
                        vldblPorcIVA = !IVAConceptoFacturacion
                        msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 15) = LTrim(!DescConceptoFacturacion)
                    
'                    End If
'                    rs.Close
                    
                    
                                     
                    msfgGrid.Col = 0
                    msfgGrid.RowData(msfgGrid.Row) = !IntNumCargo
                    'msfgGrid.TextMatrix(msfgGrid.Row, 0) = "*"
                    msfgGrid.TextMatrix(msfgGrid.Row, 1) = !CHRTIPOPACIENTE 'pstrTipoPac
                    msfgGrid.Col = 2
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 2) = !INTMOVPACIENTE ' pintNumCta
                    msfgGrid.TextMatrix(msfgGrid.Row, 3) = !chrTipoCargo
                    msfgGrid.Col = 4
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 4) = IIf(IsNull(vlstrDescripcionCargo), "", vlstrDescripcionCargo)
                    'msfgGrid.TextMatrix(msfgGrid.Row, 4) = IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo)
                    msfgGrid.Col = 5
                    msfgGrid.CellAlignment = flexAlignRightTop
                    'msfgGrid.TextMatrix(msfgGrid.Row, 5) = Format(Round(!mnyPrecio, 2), "$ ###,###,###,###.00")
                    msfgGrid.TextMatrix(msfgGrid.Row, 5) = Format(Round(!MNYPRECIO, 6), "$ ###,###,###,###.00####")
                    msfgGrid.Col = 6
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 6) = !MNYCANTIDAD
                    msfgGrid.Col = 7
                    msfgGrid.CellAlignment = flexAlignRightTop
                    'msfgGrid.TextMatrix(msfgGrid.Row, 7) = Format(Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2), "$ ###,###,###,###.00") 'Subtotal
                    msfgGrid.TextMatrix(msfgGrid.Row, 7) = Format(Round(!MNYPRECIO, 6) * Round(!MNYCANTIDAD, 2), "$ ###,###,###,###.00####") 'Subtotal
                    msfgGrid.Col = 8
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 8) = 0 'Descuentos
                    msfgGrid.Col = 9
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 9) = 0 'Total
                    msfgGrid.Col = 10
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 10) = 0 'Cantidad de IVA
                    msfgGrid.Col = 11
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 11) = 0 'Total total
                    msfgGrid.Col = 12
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 12) = Format(CStr(!DTMFECHAHORA), "dd/mmm/yyyy HH:mm")
                    msfgGrid.Col = 13
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 13) = !TipoDocumento
                    msfgGrid.Col = 14
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 14) = !intFolioDocumento
                    msfgGrid.Col = 15
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    'msfgGrid.TextMatrix(msfgGrid.Row, 15) = LTrim(RTrim(!Concepto))
                    msfgGrid.Col = 16
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 16) = !NombreDepartamento
                    msfgGrid.Col = 17
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 17) = IIf(IsNull(!FolioFactura), "", !FolioFactura)
                    msfgGrid.TextMatrix(msfgGrid.Row, 18) = !intDescuentaInventario
                    msfgGrid.TextMatrix(msfgGrid.Row, 19) = !chrCveCargo
                    msfgGrid.TextMatrix(msfgGrid.Row, 20) = IIf(IsNull(!Excluido), "", !Excluido) 'Trae una "X" si está excluido
                    msfgGrid.TextMatrix(msfgGrid.Row, 23) = IIf(IsNull(!Excluido), "", !Excluido) 'El mismo que el 18 pero nomas pa que se vea y darle performance
                    msfgGrid.TextMatrix(msfgGrid.Row, 21) = vldblPorcIVA 'Porcentaje de IVA
                    msfgGrid.TextMatrix(msfgGrid.Row, 22) = vlintCveConcepto 'CveConceptoFacturacion
                    'msfgGrid.TextMatrix(msfgGrid.Row, 21) = !smyIva 'Porcentaje de IVA
                    'msfgGrid.TextMatrix(msfgGrid.Row, 22) = !smicveconcepto 'CveConceptoFacturacion
                    msfgGrid.TextMatrix(msfgGrid.Row, 24) = IIf(!Urgente = 0, "", Trim(Str(!Urgente * 100))) 'Cargo Urgente (informacion en tabla)
                    msfgGrid.TextMatrix(msfgGrid.Row, 25) = IIf(!Urgente = 0, "", Trim(Str(!Urgente * 100))) 'Cargo Urgente (informacion para cambio)
                    msfgGrid.TextMatrix(msfgGrid.Row, 26) = !CveDepartamento          'La clave del departamento que cargo ¿neta?
                    msfgGrid.TextMatrix(msfgGrid.Row, 27) = 0 'esta columna no se usa
                    msfgGrid.TextMatrix(msfgGrid.Row, 28) = 0 'esta columna no se usa
                    msfgGrid.TextMatrix(msfgGrid.Row, 29) = 0
                    msfgGrid.TextMatrix(msfgGrid.Row, 30) = 0
                    msfgGrid.Col = 31
                    msfgGrid.CellAlignment = flexAlignLeftTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 31) = IIf(IsNull(!NombrePaquete), "", LTrim(RTrim(!NombrePaquete))) 'La descripción del paquete que tiene guardado
                    msfgGrid.TextMatrix(msfgGrid.Row, 32) = IIf(IsNull(!ClavePaquete), 0, !ClavePaquete)  'La clave del paquete al que esta asignado un cargo
                    msfgGrid.Col = 33
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 33) = IIf(IsNull(!CantidadPaquete), "", IIf(!CantidadPaquete = 0, "", !CantidadPaquete)) 'Cantidad dentro del paquete e incluidos inicialmente
                    msfgGrid.Col = 34
                    msfgGrid.CellAlignment = flexAlignRightTop
                    msfgGrid.TextMatrix(msfgGrid.Row, 34) = IIf(IsNull(!CantidadExtraPaquete), "", IIf(!CantidadExtraPaquete = 0, "", !CantidadExtraPaquete)) 'Cantidad dentro del paquete pero Fuera de la configuracion inicial
                    msfgGrid.TextMatrix(msfgGrid.Row, 35) = "" 'Estatus de cambio en paquetes, Permite ver si un registro ha sido modificado o no, para optimizar la grabada
                    msfgGrid.TextMatrix(msfgGrid.Row, 36) = IIf(IsNull(!ConceptoFacturaPaquete), 0, !ConceptoFacturaPaquete) 'Concepto de Facturación del paquete
                    msfgGrid.TextMatrix(msfgGrid.Row, 37) = IIf(IsNull(!PrecioPaquete), 0, !PrecioPaquete) 'PRECIO del paquete
                    msfgGrid.TextMatrix(msfgGrid.Row, 38) = IIf(IsNull(!PaqueteCuentaIngreso), 0, !PaqueteCuentaIngreso) 'Cuenta Ingreso del paquete
                    msfgGrid.TextMatrix(msfgGrid.Row, 39) = IIf(IsNull(!PaqueteCuentaDescuento), 0, !PaqueteCuentaDescuento) 'Cuenta de descuento del paquete
                    msfgGrid.TextMatrix(msfgGrid.Row, 40) = 0 'esta columna no se usa
                    msfgGrid.TextMatrix(msfgGrid.Row, 41) = 0 'esta columna no se usa
                    msfgGrid.TextMatrix(msfgGrid.Row, 42) = IIf(IsNull(!IVAPaquete), 0, !IVAPaquete) 'IVA del paquete (ya multiplicado)
                    msfgGrid.TextMatrix(msfgGrid.Row, 43) = IIf(IsNull(!DescripConceptoPaquete), 0, !DescripConceptoPaquete) 'Nombre del concepto de facturacion
                    msfgGrid.ColWidth(17) = 0
                    msfgGrid.ColWidth(20) = 0
                    msfgGrid.ColWidth(23) = 0
                    'Descuentos
                    'msfgGrid.TextMatrix(msfgGrid.Row, 8) = Format(!MNYDESCUENTO, "$ ###,###,###,##0.00")
                    msfgGrid.TextMatrix(msfgGrid.Row, 8) = Format(!MNYDESCUENTO, "$ ###,###,###,##0.00####")
                    'Total
                    'vldbltotal = Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2) - Round(!MNYDESCUENTO, 2)
                    vldbltotal = Round(!MNYPRECIO, 6) * Round(!MNYCANTIDAD, 2) - Round(!MNYDESCUENTO, 6)
                    'msfgGrid.TextMatrix(msfgGrid.Row, 9) = Format(vldbltotal, "$ ###,###,###,##0.00")
                    msfgGrid.TextMatrix(msfgGrid.Row, 9) = Format(vldbltotal, "$ ###,###,###,##0.00####")
                    'IVA
                    'vldblMontoIVA = (Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2) - Round(!MNYDESCUENTO, 2)) * (vldblPorcIVA / 100)
                    vldblMontoIVA = Round((Round(!MNYPRECIO, 6) * Round(!MNYCANTIDAD, 2) - Round(!MNYDESCUENTO, 6)) * (vldblPorcIVA / 100), 6)
                    msfgGrid.TextMatrix(msfgGrid.Row, 10) = vldblMontoIVA 'Format(vldblIVA, "$ ###,###,###,##0.00")
                    'Total total
                    'msfgGrid.TextMatrix(msfgGrid.Row, 11) = Format(vldbltotal + vldblMontoIVA, "$ ###,###,###,##0.00")
                    msfgGrid.TextMatrix(msfgGrid.Row, 11) = Format(vldbltotal + vldblMontoIVA, "$ ###,###,###,##0.00####")
                   
                    'vldblIVA = !MontoIVA
                    'msfgGrid.TextMatrix(msfgGrid.Row, 10) = vldblIVA 'Format(vldblIVA, "$ ###,###,###,##0.00")
                    'Total total
                    'msfgGrid.TextMatrix(msfgGrid.Row, 11) = Format(vldbltotal + vldblIVA, "$ ###,###,###,##0.00")
                    
'                    If !PrecioManual Then
'                        For vlintcontador = 3 To msfgGrid.Cols - 1
'                            msfgGrid.Col = vlintcontador
'                            msfgGrid.CellBackColor = &H80000018
'                        Next
'                    Else
'                        For vlintcontador = 1 To msfgGrid.Cols - 1
'                            msfgGrid.Col = vlintcontador
'                            msfgGrid.CellForeColor = &H80000008
'                        Next
'                    End If
                     If !PrecioManual Then
                        msfgGrid.Col = 5
                        msfgGrid.CellBackColor = &HC0FFFF   'fondo amarillo
                     End If
                     If Not IsNull(!FechaManual) Then
                        If !FechaManual = 1 Then
                           msfgGrid.Col = 12
                           msfgGrid.CellBackColor = &HC0E0FF      'fondo naranja
                        End If
                     End If

                    
                    msfgGrid.Rows = msfgGrid.Rows + 1
                    msfgGrid.Row = msfgGrid.Rows - 1
                End If
            End If
            .MoveNext
        Loop
        If msfgGrid.TextMatrix(1, 1) <> "" Then msfgGrid.Rows = msfgGrid.Rows - 1
        .Close
    End With
    
    msfgGrid.Redraw = True
    
    Me.MousePointer = 0
    
End Sub

Sub pLimpiaGrid(ObjGrd As VSFlexGrid)
    Dim vlbytColumnas As Byte
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
End Sub

Private Sub pConfiguraGridCargos(pMSHGGrid As Control, Optional pblnPoneDosRenglones As Boolean = True, Optional blnInicializa As Boolean = True)
    Dim vlintcont As Integer
    
    With pMSHGGrid
        On Error Resume Next
        .Cols = 44
        If blnInicializa Then .Rows = 1
        .FixedCols = 5
        If pblnPoneDosRenglones Then .Rows = 2
        .FixedRows = 1
        .FormatString = "|^I/E|^Cuenta|^Tipo|^Descripción del cargo|^Precio|^Cant.|^Importe|^Descuento|^Total|^IVA|^Monto total|^Fecha/Hora|^Tipo doc.|^Referencia|^Concepto de facturación|^Departamento|^Factura||||||^Exclusión||^Urgente||||||^Paquete asignado||^Incluidos|^Extras"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 320  'Tipo de paciente
        .ColWidth(2) = 650 'No. Cta.
        .ColWidth(3) = 400  'Tipo de cargo (FIX)
        .ColWidth(4) = 4000 'Descripción
        .ColWidth(5) = 1000 'Precio
        .ColWidth(6) = 450  'Cantidad
        .ColWidth(7) = 1200 'Importe
        .ColWidth(8) = 1000 'Descuentos
        .ColWidth(9) = 1200 'Total
        .ColWidth(10) = 900 'Cantidad IVA
        .ColWidth(11) = 1200 'Total total
        .ColWidth(12) = 1590 'Fecha
        .ColWidth(13) = 1200 'Tipo de documento
        .ColWidth(14) = 1000 'Numero de documento (Referencia)
        .ColWidth(15) = 3500 'Concepto de facturación
        .ColWidth(16) = 2800 'Departamento
        .ColWidth(17) = 800 'Factura
        .ColWidth(18) = 0   'bitDescuentaInventario
        .ColWidth(19) = 0   'Clave del cargo
        .ColWidth(20) = 0   'Bit Excluido
        .ColWidth(21) = 0   'IVA
        .ColWidth(22) = 0   'Clave Concepto Facturacion
        .ColWidth(23) = 800 'Estatus Visual de ExClusion "X"
        .ColWidth(24) = 0   'Información en la tabla de BitUrgente
        .ColWidth(25) = 800 'Información Visual del BitUrgente
        .ColWidth(26) = 0   'Clave del departamento que carga
        .ColWidth(27) = 0   'esta columna no se usa
        .ColWidth(28) = 0   'Cuenta de descuentos del Concepto de facturacion
        .ColWidth(29) = 0   'esta columna no se usa
        .ColWidth(30) = 0   'esta columna no se usa
        .ColWidth(31) = 2800 'Descripción del paquete
        .ColWidth(32) = 0    'Clave del paquete
        .ColWidth(33) = 800  'Cantidad de cargos incluidos en el paquete
        .ColWidth(34) = 800  'Cantidad de cargos incluidos en el paquete pero EXTRA de la configuracion inicial
        .ColWidth(35) = 0    'Estatus de cambio en paquetes, Permite ver si un registro ha sido modificado o no, para optimizar la grabada
        .ColWidth(36) = 0    'Concepto de facturacion del paquete
        .ColWidth(37) = 0    'Precio del paquete
        .ColWidth(38) = 0    'Cuenta de ingresos del Concepto del Paquete
        .ColWidth(39) = 0    'Cuenta de descuentos del Concepto del Paquete
        .ColWidth(40) = 0    'esta columna no se usa
        .ColWidth(41) = 0    'esta columna no se usa
        .ColWidth(42) = 0    'IVA del paquete (ya multiplicado)
        .ColWidth(43) = 0    'Descripcion del concepto de facturacion del paquete
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Private Sub pInicializaForma()
    SSTFacturacionConsolidada.Tab = 0
    pInicialiazaDatosGenerales
    pInicializaContenido
    pInicializaConsulta
End Sub

Private Sub pInicialiazaDatosGenerales()
    txtEmpresa.Text = ""
    txtFechaCreacion.Text = ""
    txtDGFolioFactura.Text = ""
    pConfiguraGridCargos msfgCargosAsignados
    txtDescuentos.Text = ""
    txtSubtotal.Text = ""
    txtIva.Text = ""
    txtTotal.Text = ""
    txtRetServicios.Text = ""
    txtTotalAPagar.Text = ""
End Sub

Private Sub pInicializaContenido()
    chkRangoFechas.Value = 0
    dtpFechaIni.Value = fdtmServerFecha
    dtpFechaIni.Enabled = False
    dtpFechaFin.Value = fdtmServerFecha
    dtpFechaFin.Enabled = False
    dtpFechaInicialP.Value = fdtmServerFecha
    dtpFechaFinalP.Value = fdtmServerFecha
    chkIncluirCargosAtrasados.Value = 0
    chkIncluirCargosAtrasados.Enabled = False
    optTipoPac(2).Value = True
    optTipoPacP(2).Value = True
    txtNumCuenta.Text = ""
    cboConceptoFacturacion.ListIndex = 0
    cboConceptoFacturacionP.ListIndex = 0
    cboTipoCargo.ListIndex = 0
    pLimpiaAreaDeSeleccion
End Sub
'-----------------------------------------------------------------------------------
'|  Limpia el área en que se seleccionan las cuentas y el detalle de las mismas
'-----------------------------------------------------------------------------------
Private Sub pLimpiaAreaDeSeleccion()
    pConfiguraGridCargos MSFGCargosDisponibles
    pConfiguraGridCuentas MSFGCuentasAsignadasGrupo
    pConfiguraGridCuentas MSFGCuentasDisponiblesGrupo
    pHabilitaBotonesPaso
End Sub

Private Sub pInicializaConsulta()
    If cboEmpresa.ListCount > 1 Then cboEmpresa.ListIndex = 0
    chkRangoFechasB.Value = False
    dtpFechaInicialB.Value = CDate("01/" & Month(fdtmServerFecha) & "/" & Year(fdtmServerFecha))
    dtpFechaFinalB.Value = fdtmServerFecha
    chkFacturados.Value = False
    txtFolioFacturaB.Text = ""
    optTipoPacB(2).Value = True
    txtCtaPacB.Enabled = False
    optTipoGrupoB(2).Value = True
    txtCtaPacB.Text = ""
    cboEmpresa.Enabled = False
    pConfiguraGridResultadoBusqueda True
    cmdPrint.Enabled = False
    cmdVistaPreliminar.Enabled = False
End Sub

Private Sub pQuitaCuenta(pintNumCta As Long)
    Dim vlintcont As Integer
    Dim vlintRows As Integer
    Dim vlintCols As Integer
    
    vlintcont = 1
    vlintRows = MSFGCargosDisponibles.Rows
    While vlintcont < vlintRows
        If MSFGCargosDisponibles.TextMatrix(vlintcont, 2) = pintNumCta Then
            'pBorrarRenglon MSFGCargosDisponibles, vlintCont
            If vlintRows - 1 = 1 Then
                For vlintCols = 0 To MSFGCargosDisponibles.Cols - 1
                    MSFGCargosDisponibles.TextMatrix(1, vlintCols) = ""
                Next
                MSFGCargosDisponibles.RowData(1) = -1
            Else
                MSFGCargosDisponibles.RemoveItem vlintcont
                vlintRows = MSFGCargosDisponibles.Rows
                If MSFGCargosDisponibles.TextMatrix(1, 2) <> "" Then vlintcont = vlintcont - 1
            End If
        End If
        vlintcont = vlintcont + 1
    Wend
End Sub

' -------------------------------------------------------------------
'   Procedimiento para calcular TOTALES de los cargos
' -------------------------------------------------------------------
Private Sub pTotales()
    Dim vldblSubtotal As Double
    Dim vldblTotDescuento As Double
    Dim vldblIVA As Double
    Dim vldbltotal As Double
    Dim vlintContador As Integer
    Dim vldblT As Double
    Dim vldblImportes As Double
    Dim vldblImportePReten As Double
    Dim vldblDescuentoPReten As Double
    Dim vldblTotalAPagar As Double
    
    'SubTotales = Campo 5 del grid
    vldblSubtotal = 0
    vldblTotDescuento = 0
    vldblIVA = 0
    vldblT = 0
    vldblImportes = 0
    vldblImportePReten = 0
    vldblDescuentoPReten = 0
    vldblCantRetServicios = 0
    vldblTotalAPagar = 0
    
    For vlintContador = 1 To msfgCargosAsignados.Rows - 1
        'vldblSubtotal = vldblSubtotal + Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 9), "############.##"))
        'vldblSubtotal = vldblSubtotal + Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 9), "############.######"))
        vldblImportes = vldblImportes + Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 7), ""))
        'vldblTotDescuento = vldblTotDescuento + Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 8), "############.##"))
        vldblTotDescuento = vldblTotDescuento + Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 8), "############.######"))
        vldblIVA = vldblIVA + IIf(msfgCargosAsignados.TextMatrix(vlintContador, 10) = "", 0, msfgCargosAsignados.TextMatrix(vlintContador, 10))
        vldblImportePReten = vldblImportePReten + IIf(msfgCargosAsignados.TextMatrix(vlintContador, 10) = "0", 0, Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 7), "")))
        vldblDescuentoPReten = vldblDescuentoPReten + IIf(msfgCargosAsignados.TextMatrix(vlintContador, 10) = "0", 0, Val(Format(msfgCargosAsignados.TextMatrix(vlintContador, 8), "")))
        'msfgCargosAsignados.TextMatrix(vlintContador, 10) = Format(IIf(msfgCargosAsignados.TextMatrix(vlintContador, 10) = "", 0, msfgCargosAsignados.TextMatrix(vlintContador, 10)), "$ ###,###,###,##0.00")
        msfgCargosAsignados.TextMatrix(vlintContador, 10) = Format(IIf(msfgCargosAsignados.TextMatrix(vlintContador, 10) = "", 0, msfgCargosAsignados.TextMatrix(vlintContador, 10)), "$ ###,###,###,##0.00####")
        vldblT = vldblT + CDbl(IIf(msfgCargosAsignados.TextMatrix(vlintContador, 9) = "", 0, msfgCargosAsignados.TextMatrix(vlintContador, 9)))
    Next

    If txtDGFolioFactura.Text <> "" Then
        vldblSubtotal = vldblTotalFactura - vldblIVAFactura + vldblRetFactura
        txtSubtotal.Text = Format(vldblSubtotal, "$ ###,###,###,##0.00")
        txtDescuentos.Text = Format(vldblDescFactura, "$ ###,###,###,##0.00")
        txtIva.Text = Format(vldblIVAFactura, "$ ###,###,###,##0.00")
        vldbltotal = vldblSubtotal + vldblIVAFactura
        txtTotal.Text = Format(vldbltotal, "$ ###,###,###,##0.00")
    Else
        vldblSubtotal = Round(vldblImportes, 2) - Round(vldblTotDescuento, 2)
        txtSubtotal.Text = Format(vldblSubtotal, "$ ###,###,###,##0.00")
        txtDescuentos.Text = Format(vldblTotDescuento, "$ ###,###,###,##0.00")
        txtIva.Text = Format(vldblIVA, "$ ###,###,###,##0.00")
        vldbltotal = vldblSubtotal + vldblIVA
        txtTotal.Text = Format(vldbltotal, "$ ###,###,###,##0.00")
    End If
    If (txtDGFolioFactura.Text <> "" And vldblRetFactura = 0) Or txtDGFolioFactura.Text = "" Then
        vldblTotalAPagar = vldbltotal
    ElseIf txtDGFolioFactura.Text <> "" And vldblRetFactura > 0 Then
        vldblTotalAPagar = vldbltotal - vldblRetFactura
        vldblCantRetServicios = vldblRetFactura
    'ElseIf vldblPorcRetServicios > 0 Then
    '    vldblTotalAPagar = vldbltotal - vldblCantRetServicios
    '    vldblCantRetServicios = (vldblImportePReten - vldblDescuentoPReten) * vldblPorcRetServicios
    End If
    txtRetServicios.Text = Format(vldblCantRetServicios, "$ ###,###,###,##0.00")
    txtTotalAPagar.Text = Format(vldblTotalAPagar, "$ ###,###,###,##0.00")
End Sub

Private Function fblnExisteGrupo(intCveGrupo As Double) As Boolean
    fblnExisteGrupo = True
    vgrsGrupoFacturas.Find "INTCVEGRUPO = " & intCveGrupo, 0, adSearchForward, 1
    If vgrsGrupoFacturas.EOF Then fblnExisteGrupo = False
End Function

Private Sub txtNumCuenta_Validate(Cancel As Boolean)
    If Trim(txtNumCuenta.Text) = "" Then
        cmdEjecutaConsulta.Caption = "&Ejecutar consulta"
    Else
        If optTipoPac(0).Value = 0 And optTipoPac(1).Value = 0 Then
            MsgBox SIHOMsg(530) & Chr(13) & "Especifique si la cuenta es de un paciente interno o externo.", vbExclamation, "Mensaje"
            pEnfocaTextBox txtNumCuenta
        End If
    End If
End Sub

Private Sub pPonEstado(pstrEstado As enmStatus)
    Dim strSentencia As String
    
    Select Case pstrEstado
        Case 1
            '-----------------------------------------
            '|           N U E V O
            '-----------------------------------------
            cmdInicio.Enabled = False
            cmdAnterior.Enabled = False
            cmdBuscar.Enabled = False
            cmdGrabar.Enabled = True
            cmdEliminar.Enabled = False
            cmdSiguiente.Enabled = False
            cmdFin.Enabled = False
            pInicializaForma
            '|  Tab Contenido del grupo
            SSTFacturacionConsolidada.TabEnabled(1) = True
            '|  Tab Consulta de grupos
            SSTFacturacionConsolidada.TabEnabled(2) = False
            fraFiltrosContenidoGrupoE.Enabled = True
            fraFiltrosContenidoGrupoP.Enabled = True
            pHabilitaBotonesPaso
            vgenmStatus = stNuevo
            chkCosto.Enabled = False
            chkCosto.Value = 0
        Case 2
            '-----------------------------------------
            '|           E D I C I Ó N
            '-----------------------------------------
            cmdInicio.Enabled = False
            cmdAnterior.Enabled = False
            cmdBuscar.Enabled = False
            cmdGrabar.Enabled = True
            cmdEliminar.Enabled = False
            cmdSiguiente.Enabled = False
            cmdFin.Enabled = False
            '|  Tab Contenido del grupo
            SSTFacturacionConsolidada.TabEnabled(1) = True
            '|  Tab Consulta de grupos
            SSTFacturacionConsolidada.TabEnabled(2) = False
            fraFiltrosContenidoGrupoE.Enabled = True
            fraFiltrosContenidoGrupoP.Enabled = True
            pHabilitaBotonesPaso
            vgenmStatus = stedicion
        Case 3
            '-----------------------------------------
            '|           E S P E R A
            '-----------------------------------------
            pInicializaForma
            '|  Botonera de acciones
            cmdInicio.Enabled = True
            cmdAnterior.Enabled = True
            cmdBuscar.Enabled = True
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = False
            cmdSiguiente.Enabled = True
            cmdFin.Enabled = True
            If (vgrsGrupoFacturas.State <> 0) Then
                If (vgrsGrupoFacturas.RecordCount > 0) Then vgrsGrupoFacturas.MoveLast
            End If
            txtCveGrupo.Text = fSigConsecutivo("INTCVEGRUPO", "PVFACTURACIONCONSOLIDADA")
            pEnfocaTextBox txtCveGrupo
            '|  Tab Contenido del grupo
            SSTFacturacionConsolidada.TabEnabled(1) = False
            '|  Tab Consulta de grupos
            SSTFacturacionConsolidada.TabEnabled(2) = True
            fraFiltrosContenidoGrupoE.Enabled = False
            fraFiltrosContenidoGrupoP.Enabled = False
            pHabilitaBotonesPaso
            optTipoGrupo(0).Value = True
            vgenmStatus = stEspera
            chkCosto.Enabled = False
            chkCosto.Value = 0
        Case 4
            '-----------------------------------------
            '|           F A C T U R A D O
            '-----------------------------------------
            cmdInicio.Enabled = True
            cmdAnterior.Enabled = True
            cmdBuscar.Enabled = True
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = False
            cmdSiguiente.Enabled = True
            cmdFin.Enabled = True
            '|  Tab Contenido del grupo
            SSTFacturacionConsolidada.TabEnabled(1) = False
            '|  Tab Consulta de grupos
            SSTFacturacionConsolidada.TabEnabled(2) = True
            fraFiltrosContenidoGrupoE.Enabled = False
            fraFiltrosContenidoGrupoP.Enabled = False
            '|  Deshabilita botones de paso de cuentas
            cmdAgregaTodo.Enabled = False
            cmdAgregaUno.Enabled = False
            cmdQuitaUno.Enabled = False
            cmdQuitaTodos.Enabled = False
            pHabilitaBotonesPaso
            vgenmStatus = stFacturado
            chkCosto.Enabled = True
            chkCosto.Value = 0
        Case 5
            '-----------------------------------------
            '|           C O N S U L T A
            '-----------------------------------------
            cmdInicio.Enabled = True
            cmdAnterior.Enabled = True
            cmdBuscar.Enabled = True
            cmdGrabar.Enabled = False
            cmdEliminar.Enabled = True
            cmdSiguiente.Enabled = True
            cmdFin.Enabled = True
            '|  Tab Contenido del grupo
            SSTFacturacionConsolidada.TabEnabled(1) = True
            '|  Tab Consulta de grupos
            SSTFacturacionConsolidada.TabEnabled(2) = True
            fraFiltrosContenidoGrupoE.Enabled = True
            fraFiltrosContenidoGrupoP.Enabled = True
            '|  Deshabilita botones de paso de cuentas
            pHabilitaBotonesPaso
            vgenmStatus = stConsulta
            chkCosto.Enabled = True
            chkCosto.Value = 0
    End Select
    '----------------------------------------------------------------
    '|  Valida que se haya establecido un concepto de liquidación
    '|  para que se puedan crear grupos de particulares.
    '----------------------------------------------------------------
    strSentencia = "Select COUNT(*) Co From PVCONCEPTOPAGOempresa Where BITCONCEPTOLIQUIDACION = 1 and pvconceptopagoempresa.intcveempresa = " & vgintClaveEmpresaContable
    If frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)!CO = 0 Then
        optTipoGrupo(1).Enabled = False
        optTipoGrupoB(0).Value = True
        optTipoGrupoB(1).Enabled = False
        optTipoGrupoB(2).Enabled = False
    End If
End Sub

Private Sub pConfiguraGridCuentas(msfgGrid As VSFlexGrid)
    With msfgGrid
        .Clear
        .FixedRows = 1
        .Cols = 4
        .Rows = 2
        .ColWidth(0) = "800"
        .ColWidth(1) = "500"
        .ColWidth(2) = "4080"
        '|  Si la columna 3 tiene un "*" significa que la cuenta ya pertenece al grupo,
        '|  esto sirve para cuando se realiza la búsqueda de los cargos.
        .ColWidth(3) = "0"
        .TextMatrix(0, 0) = "Cuenta"
        .TextMatrix(0, 1) = "Tipo"
        .TextMatrix(0, 2) = "Nombre del paciente"
        .Row = 1
        .Col = 1
    End With
    
End Sub
'--------------------------------------------------------------------------------
'|  Verifica si existe la cuenta que se desea insertar o remover de un grid
'|  Regresa True si existe y False si no.
'--------------------------------------------------------------------------------
Private Function fblnExisteCuenta(msfgGrid As VSFlexGrid, _
                                  intCuenta As Long, _
                                  strTipoCta As String) As Boolean
    Dim intRow As Integer
    
    fblnExisteCuenta = False
    For intRow = 1 To msfgGrid.Rows - 1
        If Val(msfgGrid.TextMatrix(intRow, 0)) = intCuenta _
           And msfgGrid.TextMatrix(intRow, 1) = strTipoCta Then
           fblnExisteCuenta = True
        End If
    Next
                             
End Function
                             
'--------------------------------------------------------------------------------
'|  Despliega la cuenta que se esta mostrado en el grid detalle
'|  para que el usuario ubique que es lo que está consultando.
'--------------------------------------------------------------------------------
Private Sub pPonEtiqueta(strCuenta As String)
    lblDetalleCargos.Caption = "Detalle de los cargos de la cuenta "
    lblCtaSeleccionada.Caption = strCuenta
End Sub

Private Sub pConfiguraGridResultadoBusqueda(pblnInicializa As Boolean)
    With MSFGResultado
        .Redraw = False
        .Cols = 4
        If pblnInicializa Then .Rows = 2
        .FixedCols = 0
        .FixedRows = 1
        .ColWidth(0) = 1000
        .ColWidth(1) = 2000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .TextMatrix(0, 0) = "Clave"
        .TextMatrix(0, 1) = "Creación"
        .TextMatrix(0, 2) = "Factura"
        .TextMatrix(0, 3) = "Empresa"
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .Redraw = True
    End With
    
End Sub

'-----------------------------------------------------------------------------------
'|  Despliega todos los datos almacenados en la BD en los componentes de la forma
'-----------------------------------------------------------------------------------
Private Sub pDespliegaGrupo(intCveGrupo As Integer)
    Dim rsgrupo As New ADODB.Recordset
    Dim rsDetalleGrupo As New ADODB.Recordset
    Dim strSQL As String
    Dim vldbltotal As Double
    Dim vldblIVA As Double
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vllngNumCta As Long
    Dim vllngCveConceptoFacturacion As Long
    
    pInicialiazaDatosGenerales
    pInicializaContenido
    
    vlintEmpresa = 0
    vlintCveConcepto = 0
    vldblPorcIVA = 0
    vldblMontoIVA = 0
    vllngNumCta = 0
    vlstrCveCargo = " "
    vlstrTipoCargoEmpresa = " "
    vlstrDescripcionCargo = " "
    
    Set rsgrupo = frsEjecuta_SP(CStr(intCveGrupo), "Sp_PvSelGrupoFacturas")
    If rsgrupo.RecordCount = 0 Then Exit Sub
    With rsgrupo
        txtEmpresa.Text = !empresa
        txtFechaCreacion.Text = !Fecha
        txtDGFolioFactura.Text = IIf(IsNull(!Folio), "", !Folio)
        vgintCveEmpresa = !cveEmpresa
        vldblPorcRetServicios = IIf(IsNull(!RetServiciosEmpresa), 0, !RetServiciosEmpresa)
        vldblRetFactura = IIf(IsNull(!mnyretencion), 0, !mnyretencion)
        vldblDescFactura = IIf(IsNull(!MNYDESCUENTO), 0, !MNYDESCUENTO)
        vldblIVAFactura = IIf(IsNull(!MNYIVA), 0, !MNYIVA)
        vldblTotalFactura = IIf(IsNull(!mnyTotalFactura), 0, !mnyTotalFactura)
        If vgintCveEmpresa = -1 Then
            optTipoGrupo(1).Value = True
        Else
            optTipoGrupo(0).Value = True
        End If
        While Not .EOF
            '|  Si el primer renglón del grid de las cuentas asiganadas esta vacío
            '|  lo usa para pasar la cuenta, sino agrega un nuevo renglón
            If MSFGCuentasAsignadasGrupo.TextMatrix(1, 1) = "" Then
                MSFGCuentasAsignadasGrupo.TextMatrix(1, 0) = !cuenta
                MSFGCuentasAsignadasGrupo.TextMatrix(1, 1) = !tipo
                MSFGCuentasAsignadasGrupo.TextMatrix(1, 2) = !NombrePaciente
                MSFGCuentasAsignadasGrupo.TextMatrix(1, 3) = "*"
            Else
                MSFGCuentasAsignadasGrupo.AddItem !cuenta & Chr(9) & !tipo & Chr(9) & !NombrePaciente & Chr(9) & "*"
            End If
            .MoveNext
        Wend
        .Close
    End With
    If txtDGFolioFactura.Text <> "" Then
        pPonEstado stFacturado
    Else
        pPonEstado stConsulta
    End If
    Set rsDetalleGrupo = frsEjecuta_SP(intCveGrupo & "|G|" & IIf(txtDGFolioFactura.Text = "", "0|-1", "1|" & Trim(txtDGFolioFactura.Text)) & "|G|S|0", "SP_PVSELCARGOSPACIENTEGRUPOS")
    If rsDetalleGrupo.EOF Then Exit Sub
    With rsDetalleGrupo
    
    msfgCargosAsignados.Redraw = False
    
        '|  Llena el grid con los cargos facturados
        While Not rsDetalleGrupo.EOF

            '-----------------------------------------------------------------------------------------------------------------
            ' Trae los datos de los cargos de la empresa
            '-----------------------------------------------------------------------------------------------------------------
            
            vllngCveConceptoFacturacion = !smicveconcepto
            
            vlstrCveCargo = !chrCveCargo
            vlstrTipoCargoEmpresa = !chrTipoCargo
            vllngNumCta = !INTMOVPACIENTE ' pintNumCta
            
            vlintEmpresa = IIf(IsNull(!ClaveEmpresaPaciente), 0, !ClaveEmpresaPaciente)
            
'            vlstrSentencia = "SELECT smicveconcepto, vchdescripcion FROM CCCATALOGOPCE " & _
'                             " WHERE chrcvecargo = '" & Trim(vlstrCveCargo) & "'" & _
'                             "   AND chrtipocargo = '" & Trim(vlstrTipoCargoEmpresa) & "'" & _
'                                     "   AND intCveEmpresa = " & vlintEmpresa
'            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'            If rs.RecordCount > 0 Then

                vlintCveConcepto = 0
                vlintCveConcepto = IIf(IsNull(!ClaveConceptoPCE) Or !ClaveConceptoPCE = 0, vllngCveConceptoFacturacion, !ClaveConceptoPCE)
                
                vlstrDescripcionCargo = " "
                vlstrDescripcionCargo = IIf(IsNull(!DescripcionConceptoPCE) Or !DescripcionConceptoPCE = "", !DescripcionCargosSoloPVCARGO, !DescripcionConceptoPCE)
                
'            Else
'                Select Case vlstrTipoCargoEmpresa
'                    Case "AR"
'                              vlstrSentencia = "SELECT SMICVECONCEPTFACT smicveconcepto, VCHNOMBRECOMERCIAL vchdescripcion FROM ivarticulo " & _
'                                               " WHERE intidarticulo = " & (CLng(Trim(vlstrCveCargo)))
'                    Case "OC"
'                              vlstrSentencia = "SELECT SMICONCEPTOFACT smicveconcepto, CHRDESCRIPCION vchdescripcion FROM pvotroconcepto " & _
'                                               " WHERE intcveconcepto = " & (CLng(Trim(vlstrCveCargo)))
'                    Case "ES"
'                              vlstrSentencia = "SELECT SMICONFACT smicveconcepto, VCHNOMBRE vchdescripcion FROM imestudio " & _
'                                               " WHERE intcveestudio = " & (CLng(Trim(vlstrCveCargo)))
'                    Case "EX"
'                              vlstrSentencia = "SELECT SMICONFACT smicveconcepto, CHRNOMBRE vchdescripcion FROM laexamen " & _
'                                               " WHERE intcveexamen = " & (CLng(Trim(vlstrCveCargo)))
'                    Case "GE"
'                              vlstrSentencia = "SELECT SMICONFACT smicveconcepto, CHRNOMBRE vchdescripcion FROM lagrupoexamen " & _
'                                               " WHERE intcvegrupo = " & (CLng(Trim(vlstrCveCargo)))
'                End Select
'                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                If rs.RecordCount > 0 Then
'                    vlintCveConcepto = vllngCveConceptoFacturacion
'                    vlstrDescripcionCargo = rs!VCHDESCRIPCION
'                Else
'                    vlstrDescripcionCargo = " "
'                End If
'            End If
'            rs.Close
                    
'            vlstrSentencia = "SELECT smyIva, chrdescripcion FROM pvConceptoFacturacion " & _
'                             " WHERE smiCveConcepto = " & vlintCveConcepto
'            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                If rs.RecordCount > 0 Then

                    vldblPorcIVA = 0
                    vldblPorcIVA = !IVAConceptoFacturacion
                    msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 15) = LTrim(!DescConceptoFacturacion)
                    
'                End If
'            rs.Close
        
            msfgCargosAsignados.RowData(msfgCargosAsignados.Row) = !IntNumCargo
'            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 0) = "*"
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 1) = !CHRTIPOPACIENTE 'pstrTipoPac
            msfgCargosAsignados.Col = 2
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 2) = !INTMOVPACIENTE ' pintNumCta
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 3) = !chrTipoCargo
            msfgCargosAsignados.Col = 4
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 4) = IIf(IsNull(vlstrDescripcionCargo), "", vlstrDescripcionCargo)
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 4) = IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo)
            msfgCargosAsignados.Col = 5
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 5) = Format(Round(!mnyPrecio, 2), "$ ###,###,###,###.00")
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 5) = Format(Round(!MNYPRECIO, 6), "$ ###,###,###,###.00####")
            msfgCargosAsignados.Col = 6
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 6) = !MNYCANTIDAD
            msfgCargosAsignados.Col = 7
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 7) = Format(Round((!mnyPrecio * !mnyCantidad), 2), "$ ###,###,###,###.00") 'Subtotal  Format(Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2), "$ ###,###,###,###.00") 'Subtotal
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 7) = Format(Round((!MNYPRECIO * !MNYCANTIDAD), 6), "$ ###,###,###,###.00####") 'Subtotal  Format(Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2), "$ ###,###,###,###.00") 'Subtotal
            msfgCargosAsignados.Col = 8
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 8) = 0 'Descuentos
            msfgCargosAsignados.Col = 9
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 9) = 0 'Total
            msfgCargosAsignados.Col = 10
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 10) = 0 'Cantidad de IVA
            msfgCargosAsignados.Col = 11
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 11) = 0 'Total total
            msfgCargosAsignados.Col = 12
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 12) = Format(CStr(!DTMFECHAHORA), "dd/mmm/yyyy HH:mm")
            msfgCargosAsignados.Col = 13
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 13) = !TipoDocumento
            msfgCargosAsignados.Col = 14
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 14) = !intFolioDocumento
            msfgCargosAsignados.Col = 15
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 15) = LTrim(RTrim(!Concepto))
            msfgCargosAsignados.Col = 16
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 16) = !NombreDepartamento
            msfgCargosAsignados.Col = 17
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 17) = IIf(IsNull(!FolioFactura), "", !FolioFactura)
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 18) = !intDescuentaInventario
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 19) = !chrCveCargo
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 20) = IIf(IsNull(!Excluido), "", !Excluido) 'Trae una "X" si está excluido
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 23) = IIf(IsNull(!Excluido), "", !Excluido) 'El mismo que el 18 para que se vea y darle performance
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 21) = vldblPorcIVA 'Porcentaje de IVA
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 22) = vlintCveConcepto 'CveConceptoFacturacion
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 21) = !smyIVA 'Porcentaje de IVA
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 22) = !smiCveConcepto 'CveConceptoFacturacion
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 24) = IIf(!Urgente = 0, "", Trim(Str(!Urgente * 100))) 'Cargo Urgente (informacion en tabla)
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 25) = IIf(!Urgente = 0, "", Trim(Str(!Urgente * 100))) 'Cargo Urgente (informacion para cambio)
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 26) = !CveDepartamento
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 27) = 0 'esta columna no se usa
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 28) = 0 'esta columna no se usa
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 29) = 0
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 30) = 0
            msfgCargosAsignados.Col = 31
            msfgCargosAsignados.CellAlignment = flexAlignLeftTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 31) = IIf(IsNull(!NombrePaquete), "", LTrim(RTrim(!NombrePaquete))) 'La descripción del paquete que tiene guardado
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 32) = IIf(IsNull(!ClavePaquete), 0, !ClavePaquete)  'La clave del paquete al que esta asignado un cargo
            msfgCargosAsignados.Col = 33
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 33) = IIf(IsNull(!CantidadPaquete), "", IIf(!CantidadPaquete = 0, "", !CantidadPaquete)) 'Cantidad dentro del paquete e incluidos inicialmente
            msfgCargosAsignados.Col = 34
            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 34) = IIf(IsNull(!CantidadExtraPaquete), "", IIf(!CantidadExtraPaquete = 0, "", !CantidadExtraPaquete)) 'Cantidad dentro del paquete pero Fuera de la configuracion inicial
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 35) = "" 'Estatus de cambio en paquetes, Permite ver si un registro ha sido modificado o no, para optimizar la grabada
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 36) = IIf(IsNull(!ConceptoFacturaPaquete), 0, !ConceptoFacturaPaquete) 'Concepto de Facturación del paquete
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 37) = IIf(IsNull(!PrecioPaquete), 0, !PrecioPaquete) 'PRECIO del paquete
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 38) = IIf(IsNull(!PaqueteCuentaIngreso), 0, !PaqueteCuentaIngreso) 'Cuenta Ingreso del paquete
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 39) = IIf(IsNull(!PaqueteCuentaDescuento), 0, !PaqueteCuentaDescuento) 'Cuenta de descuento del paquete
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 40) = 0 'esta columna no se usa
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 41) = 0 'esta columna no se usa
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 42) = IIf(IsNull(!IVAPaquete), 0, !IVAPaquete) 'IVA del paquete (ya multiplicado)
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 43) = IIf(IsNull(!DescripConceptoPaquete), 0, !DescripConceptoPaquete) 'Nombre del concepto de facturacion
            'Descuentos
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 8) = Format(!MNYDESCUENTO, "$ ###,###,###,##0.00")
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 8) = Format(!MNYDESCUENTO, "$ ###,###,###,##0.00####")
            'Total
            'vldbltotal = Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2) - Round(!MNYDESCUENTO, 2)
            vldbltotal = Round(!MNYPRECIO, 6) * Round(!MNYCANTIDAD, 2) - Round(!MNYDESCUENTO, 6)
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 9) = Format(vldbltotal, "$ ###,###,###,##0.00")
            msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 9) = Format(vldbltotal, "$ ###,###,###,##0.00####")
            'IVA
            
            'vldblIVA = !MontoIVA
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 10) = vldblIVA 'Format(vldblIVA, "$ ###,###,###,##0.00")
            'Total total
            'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 11) = Format(vldbltotal + vldblIVA, "$ ###,###,###,##0.00")

             'vldblMontoIVA = (Round(!mnyPrecio, 2) * Round(!mnyCantidad, 2) - Round(!MNYDESCUENTO, 2)) * (vldblPorcIVA / 100)
             vldblMontoIVA = Round((Round(!MNYPRECIO, 6) * Round(!MNYCANTIDAD, 2) - Round(!MNYDESCUENTO, 6)) * (vldblPorcIVA / 100), 6)
             msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 10) = vldblMontoIVA 'Format(vldblIVA, "$ ###,###,###,##0.00")
             'Total total
             'msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 11) = Format(vldbltotal + vldblMontoIVA, "$ ###,###,###,##0.00")
             msfgCargosAsignados.TextMatrix(msfgCargosAsignados.Row, 11) = Format(vldbltotal + vldblMontoIVA, "$ ###,###,###,##0.00####")



             If !PrecioManual Then
                msfgCargosAsignados.Col = 5
                msfgCargosAsignados.CellBackColor = &HC0FFFF   'fondo amarillo
             End If
             If Not IsNull(!FechaManual) Then
                If !FechaManual = 1 Then
                   msfgCargosAsignados.Col = 12
                   msfgCargosAsignados.CellBackColor = &HC0E0FF      'fondo naranja
                End If
             End If
            msfgCargosAsignados.Rows = msfgCargosAsignados.Rows + 1
            msfgCargosAsignados.Row = msfgCargosAsignados.Rows - 1
            .MoveNext
        Wend
        msfgCargosAsignados.Rows = msfgCargosAsignados.Rows - 1
        msfgCargosAsignados.ColWidth(17) = 0
        msfgCargosAsignados.ColWidth(20) = 0
        msfgCargosAsignados.ColWidth(23) = 0
        .Close
    End With
    
    msfgCargosAsignados.Redraw = True
    
    pTotales
End Sub

Private Sub pLlenaCargosAsignados()
    Dim intRows As Integer
    
    pConfiguraGridCargos msfgCargosAsignados
    
    For intRows = 1 To MSFGCuentasAsignadasGrupo.Rows - 1
        pLlenaCargos Val(MSFGCuentasAsignadasGrupo.TextMatrix(intRows, 0)), _
                     MSFGCuentasAsignadasGrupo.TextMatrix(intRows, 1), _
                     "C", _
                     msfgCargosAsignados, _
                     MSFGCuentasAsignadasGrupo.TextMatrix(intRows, 3), _
                     1
    Next
    pTotales
End Sub

Private Sub pEnmarcaDesmarca()
Dim vlintContador As Integer
   With MSFGCargosDisponibles
        If MSFGCuentasAsignadasGrupo.Row < 0 Then Exit Sub
        
        If .Row > 0 Then
            .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "*", "", "*")
        End If
        If .TextMatrix(.Row, 0) = "" Then
            For vlintContador = 1 To msfgCargosAsignados.Rows - 1
                If ((MSFGCargosDisponibles.TextMatrix(.Row, 2) = msfgCargosAsignados.TextMatrix(vlintContador, 2)) And (MSFGCargosDisponibles.TextMatrix(.Row, 3) = msfgCargosAsignados.TextMatrix(vlintContador, 3)) And (MSFGCargosDisponibles.RowData(MSFGCargosDisponibles.Row) = msfgCargosAsignados.RowData(vlintContador))) Then
                   If msfgCargosAsignados.Rows > 2 Then
                         msfgCargosAsignados.RemoveItem vlintContador
                         If vgenmStatus = stConsulta Then pPonEstado stedicion
                   Else
                        pConfiguraGridCargos msfgCargosAsignados
                   End If
                   Exit For
                End If
            Next
        Else
            If msfgCargosAsignados.Rows > 2 Or (msfgCargosAsignados.Rows = 2 And Trim(msfgCargosAsignados.TextMatrix(1, 2)) <> "") Then
                msfgCargosAsignados.AddItem "" & Chr(9) & .TextMatrix(.Row, 1) & Chr(9) & .TextMatrix(.Row, 2) & Chr(9) & .TextMatrix(.Row, 3) & Chr(9) & .TextMatrix(.Row, 4) & Chr(9) & .TextMatrix(.Row, 5) & Chr(9) & .TextMatrix(.Row, 6) & Chr(9) & .TextMatrix(.Row, 7) & Chr(9) & .TextMatrix(.Row, 8) & Chr(9) & .TextMatrix(.Row, 9) & Chr(9) & .TextMatrix(.Row, 10) & Chr(9) & .TextMatrix(.Row, 11) & Chr(9) & .TextMatrix(.Row, 12) & Chr(9) & .TextMatrix(.Row, 13) & Chr(9) & .TextMatrix(.Row, 14) & Chr(9) & .TextMatrix(.Row, 15) & Chr(9) & .TextMatrix(.Row, 16) & Chr(9) & .TextMatrix(.Row, 17) & Chr(9) & .TextMatrix(.Row, 18) & Chr(9) & .TextMatrix(.Row, 19) & Chr(9) & .TextMatrix(.Row, 20) & Chr(9) & .TextMatrix(.Row, 21) & Chr(9) & .TextMatrix(.Row, 22)
                msfgCargosAsignados.RowData(msfgCargosAsignados.Rows - 1) = MSFGCargosDisponibles.RowData(.Row)
                msfgCargosAsignados.Row = msfgCargosAsignados.Rows - 1
                msfgCargosAsignados.RowHeight(msfgCargosAsignados.Row) = 240
                msfgCargosAsignados.Refresh
                If vgenmStatus = stConsulta Then pPonEstado stedicion
            Else
                msfgCargosAsignados.RowData(msfgCargosAsignados.Rows - 1) = MSFGCargosDisponibles.RowData(.Row)
                msfgCargosAsignados.Row = msfgCargosAsignados.Rows - 1
'                msfgCargosAsignados.TextMatrix(1, 0) = "*"
                msfgCargosAsignados.TextMatrix(1, 1) = .TextMatrix(.Row, 1)
                msfgCargosAsignados.TextMatrix(1, 2) = .TextMatrix(.Row, 2)
                msfgCargosAsignados.TextMatrix(1, 3) = .TextMatrix(.Row, 3)
                msfgCargosAsignados.TextMatrix(1, 4) = .TextMatrix(.Row, 4)
                msfgCargosAsignados.TextMatrix(1, 5) = .TextMatrix(.Row, 5)
                msfgCargosAsignados.TextMatrix(1, 6) = .TextMatrix(.Row, 6)
                msfgCargosAsignados.TextMatrix(1, 7) = .TextMatrix(.Row, 7)
                msfgCargosAsignados.TextMatrix(1, 8) = .TextMatrix(.Row, 8)
                msfgCargosAsignados.TextMatrix(1, 9) = .TextMatrix(.Row, 9)
                msfgCargosAsignados.TextMatrix(1, 10) = .TextMatrix(.Row, 10)
                msfgCargosAsignados.TextMatrix(1, 11) = .TextMatrix(.Row, 11)
                msfgCargosAsignados.TextMatrix(1, 12) = .TextMatrix(.Row, 12)
                msfgCargosAsignados.TextMatrix(1, 13) = .TextMatrix(.Row, 13)
                msfgCargosAsignados.TextMatrix(1, 14) = .TextMatrix(.Row, 14)
                msfgCargosAsignados.TextMatrix(1, 15) = .TextMatrix(.Row, 15)
                msfgCargosAsignados.TextMatrix(1, 16) = .TextMatrix(.Row, 16)
                msfgCargosAsignados.TextMatrix(1, 17) = .TextMatrix(.Row, 17)
                msfgCargosAsignados.TextMatrix(1, 18) = .TextMatrix(.Row, 18)
                msfgCargosAsignados.TextMatrix(1, 19) = .TextMatrix(.Row, 19)
                msfgCargosAsignados.TextMatrix(1, 20) = .TextMatrix(.Row, 20)
                msfgCargosAsignados.TextMatrix(1, 21) = .TextMatrix(.Row, 21)
                msfgCargosAsignados.TextMatrix(1, 22) = .TextMatrix(.Row, 22)
            End If
                
                
                
                
    '            msfgCargosAsignados.Col = 0
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 1
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 2
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 3
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 4
                msfgCargosAsignados.CellAlignment = flexAlignLeftTop
                msfgCargosAsignados.Col = 5
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 6
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 7
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 8
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 9
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 10
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 11
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 12
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 13
                msfgCargosAsignados.CellAlignment = flexAlignLeftTop
                msfgCargosAsignados.Col = 14
                msfgCargosAsignados.CellAlignment = flexAlignRightTop
                msfgCargosAsignados.Col = 15
                msfgCargosAsignados.CellAlignment = flexAlignLeftTop
                msfgCargosAsignados.Col = 16
                msfgCargosAsignados.CellAlignment = flexAlignLeftTop
                msfgCargosAsignados.Col = 17
                msfgCargosAsignados.CellAlignment = flexAlignLeftTop
    '            msfgCargosAsignados.Col = 18
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 19
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 20
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 21
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
    '            msfgCargosAsignados.Col = 22
    '            msfgCargosAsignados.CellAlignment = flexAlignRightTop
            

           
        End If
   End With
   Exit Sub
End Sub

Private Sub pmarcados()
    Dim vlintasignados As Integer
    Dim vlintdisponibles As Integer
    
    For vlintdisponibles = 1 To MSFGCargosDisponibles.Rows - 1

        For vlintasignados = 1 To msfgCargosAsignados.Rows - 1
            If MSFGCargosDisponibles.TextMatrix(vlintdisponibles, 2) <> "" Then
                If ((MSFGCargosDisponibles.TextMatrix(vlintdisponibles, 2) = msfgCargosAsignados.TextMatrix(vlintasignados, 2)) And (MSFGCargosDisponibles.TextMatrix(vlintdisponibles, 3) = msfgCargosAsignados.TextMatrix(vlintasignados, 3)) And (MSFGCargosDisponibles.RowData(vlintdisponibles) = msfgCargosAsignados.RowData(vlintasignados))) Then
                    MSFGCargosDisponibles.TextMatrix(vlintdisponibles, 0) = "*"
                End If
            End If
        Next
    Next

End Sub

Private Sub pseleccioncuenta(vllngCuenta As String, vlstrTipo As String)
    Dim vlintcont As Integer
    Dim blnSalir As Boolean
  
    For vlintcont = 1 To msfgCargosAsignados.Rows - 1
        If (vllngCuenta = msfgCargosAsignados.TextMatrix(vlintcont, 2)) And (vlstrTipo = msfgCargosAsignados.TextMatrix(vlintcont, 1)) Then
            msfgCargosAsignados.TextMatrix(vlintcont, 2) = ""
        End If
    Next
    vlintcont = 1
    Do While Not blnSalir
        If msfgCargosAsignados.TextMatrix(vlintcont, 2) = "" Then
            If msfgCargosAsignados.Rows = 2 Then
                pConfiguraGridCargos msfgCargosAsignados
                blnSalir = True
            Else
                msfgCargosAsignados.RemoveItem vlintcont
            End If
        Else
            vlintcont = vlintcont + 1
        End If
        If msfgCargosAsignados.Rows - 1 < vlintcont Then
            blnSalir = True
        End If
    Loop
     pConfiguraGridCargos MSFGCargosDisponibles

End Sub

