VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmConsultaPOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consultas"
   ClientHeight    =   7695
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freCargando 
      Height          =   1335
      Left            =   1440
      TabIndex        =   18
      Top             =   8160
      Visible         =   0   'False
      Width           =   7920
      Begin VB.Label Label17 
         Caption         =   "Consultando información, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   345
         Width           =   6855
      End
   End
   Begin VB.Frame freBarraCFD 
      Height          =   1005
      Left            =   820
      TabIndex        =   23
      Top             =   8160
      Visible         =   0   'False
      Width           =   9720
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   24
         Top             =   600
         Width           =   9640
         _ExtentX        =   17013
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
         Left            =   90
         TabIndex        =   25
         Top             =   180
         Width           =   9450
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Left            =   30
         Top             =   120
         Width           =   9660
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1290
      Left            =   1635
      TabIndex        =   20
      Top             =   2670
      Visible         =   0   'False
      Width           =   7680
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   360
         Left            =   1035
         TabIndex        =   21
         Top             =   675
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H80000002&
         Caption         =   "Cargando información, por favor espere..."
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
         TabIndex        =   22
         Top             =   180
         Width           =   7410
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   30
         Top             =   120
         Width           =   7620
      End
   End
   Begin VB.Frame freMuestraTicket 
      Height          =   5640
      Left            =   240
      TabIndex        =   5
      Top             =   8040
      Visible         =   0   'False
      Width           =   10965
      Begin VB.Frame Frame4 
         Height          =   2160
         Left            =   7185
         TabIndex        =   9
         Top             =   3360
         Width           =   3670
         Begin VB.TextBox TxtIEPS 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   645
            Width           =   1725
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   990
            Width           =   1725
         End
         Begin VB.TextBox txtDescuentos 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   300
            Width           =   1725
         End
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   1335
            Width           =   1725
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1815
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   1685
            Width           =   1725
         End
         Begin VB.Label lblIEPS 
            Caption         =   "IEPS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   435
            TabIndex        =   26
            Top             =   690
            Width           =   885
         End
         Begin VB.Label Label3 
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
            Height          =   255
            Left            =   435
            TabIndex        =   17
            Top             =   1035
            Width           =   1365
         End
         Begin VB.Label Label4 
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
            Height          =   255
            Left            =   435
            TabIndex        =   16
            Top             =   1380
            Width           =   1365
         End
         Begin VB.Label Label5 
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
            Height          =   255
            Left            =   435
            TabIndex        =   15
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Left            =   435
            TabIndex        =   14
            Top             =   1725
            Width           =   1230
         End
      End
      Begin VB.CommandButton cmdCerrarMuestraTicket 
         Caption         =   "&Cerrar"
         Height          =   420
         Left            =   105
         TabIndex        =   8
         Top             =   5100
         Width           =   1350
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdMuestraTicket 
         Height          =   2775
         Left            =   105
         TabIndex        =   6
         Top             =   555
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   4895
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta del ticket"
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
         Left            =   105
         TabIndex        =   7
         Top             =   150
         Width           =   9390
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   45
         Top             =   120
         Width           =   10875
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdArticulos 
      Height          =   1200
      Left            =   3165
      TabIndex        =   4
      Top             =   9345
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   2117
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame13 
      Caption         =   "Rango de fechas"
      Height          =   750
      Left            =   180
      TabIndex        =   2
      Top             =   450
      Width           =   2865
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
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   "Fecha final de la búsqueda"
         Top             =   300
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
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Fecha inicial de la búsqueda"
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy "
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFactura 
      Height          =   1185
      Left            =   225
      TabIndex        =   3
      Top             =   9345
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   2090
      _Version        =   393216
      Enabled         =   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstPOS 
      Height          =   7725
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   13626
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tickets"
      TabPicture(0)   =   "frmConsultaPOS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbTipoPacienteEmpresa"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdBuscaTickets"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "comPrinter"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboProcedencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Facturas"
      TabPicture(1)   =   "frmConsultaPOS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optMostrarSolo(1)"
      Tab(1).Control(1)=   "optMostrarSolo(2)"
      Tab(1).Control(2)=   "optMostrarSolo(3)"
      Tab(1).Control(3)=   "optMostrarSolo(4)"
      Tab(1).Control(4)=   "optMostrarSolo(0)"
      Tab(1).Control(5)=   "PB"
      Tab(1).Control(6)=   "Frame5"
      Tab(1).Control(7)=   "grdBuscaFacturas"
      Tab(1).Control(8)=   "Label57(0)"
      Tab(1).Control(9)=   "Label57(6)"
      Tab(1).Control(10)=   "Label57(7)"
      Tab(1).Control(11)=   "Label57(9)"
      Tab(1).Control(12)=   "Label57(13)"
      Tab(1).Control(13)=   "Label57(8)"
      Tab(1).Control(14)=   "Label57(14)"
      Tab(1).Control(15)=   "Label57(12)"
      Tab(1).Control(16)=   "Label57(15)"
      Tab(1).Control(17)=   "Label57(10)"
      Tab(1).ControlCount=   18
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de timbre fiscal"
         Height          =   255
         Index           =   1
         Left            =   -71880
         TabIndex        =   65
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de cancelar ante el SAT"
         Height          =   255
         Index           =   2
         Left            =   -71880
         TabIndex        =   64
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de autorización de cancelación"
         Height          =   255
         Index           =   3
         Left            =   -68040
         TabIndex        =   63
         Top             =   480
         Width           =   4335
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo cancelación rechazada"
         Height          =   255
         Index           =   4
         Left            =   -68040
         TabIndex        =   62
         Top             =   720
         Width           =   4335
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar todo"
         Height          =   255
         Index           =   0
         Left            =   -71880
         TabIndex        =   61
         Top             =   480
         Value           =   -1  'True
         Width           =   4335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Formas de pago"
         Height          =   765
         Left            =   4920
         TabIndex        =   51
         ToolTipText     =   "Forma de pago en que fueron emitidos los tickets"
         Top             =   450
         Width           =   2085
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Efectivo"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Forma de pago en que fueron emitidos los tickets"
            Top             =   465
            Width           =   885
         End
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Crédito"
            Height          =   255
            Index           =   1
            Left            =   1150
            TabIndex        =   33
            ToolTipText     =   "Forma de pago en que fueron emitidos los tickets"
            Top             =   465
            Width           =   855
         End
         Begin VB.OptionButton optFormaPago 
            Caption         =   "Todas"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Forma de pago en que fueron emitidos los tickets"
            Top             =   220
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.PictureBox PB 
         Height          =   135
         Left            =   -65400
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   46
         Top             =   7320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -71280
         TabIndex        =   43
         Top             =   6960
         Width           =   4605
         Begin VB.CommandButton cmdCancelaFacturasSAT 
            Caption         =   "Validar comprobantes pendientes de cancelación"
            Height          =   495
            Left            =   2175
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Cancelar factura(s) ante el SAT"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2295
         End
         Begin VB.CommandButton Cmdconfirmartimbre 
            Caption         =   "Confirmar timbre fiscal"
            Height          =   495
            Left            =   120
            Picture         =   "frmConsultaPOS.frx":0038
            TabIndex        =   44
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.ComboBox cboProcedencia 
         Height          =   315
         Left            =   7095
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         ToolTipText     =   "Procedencia"
         Top             =   720
         Width           =   4185
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   200
         TabIndex        =   36
         Top             =   6840
         Width           =   11070
         Begin VB.CommandButton cmdReimpresion 
            Caption         =   "Reimprimir"
            Height          =   390
            Left            =   240
            TabIndex        =   42
            Top             =   225
            Width           =   1545
         End
         Begin VB.Frame Frame3 
            Height          =   525
            Left            =   5895
            TabIndex        =   41
            Top             =   120
            Width           =   90
         End
         Begin VB.CommandButton cmdFacturaSeleccion 
            Caption         =   "Facturar"
            Height          =   390
            Left            =   1793
            TabIndex        =   40
            Top             =   225
            Width           =   1545
         End
         Begin VB.CommandButton cmdCancelaSelecion 
            Caption         =   "Cancelar"
            Height          =   390
            Left            =   3360
            TabIndex        =   39
            Top             =   225
            Width           =   1545
         End
         Begin VB.CommandButton cmdSeleTodo 
            Caption         =   "Seleccionar todo"
            Height          =   390
            Left            =   7170
            TabIndex        =   38
            Top             =   225
            Width           =   1875
         End
         Begin VB.CommandButton cmdDesSeleToto 
            Caption         =   "Quitar selección"
            Height          =   390
            Left            =   9060
            TabIndex        =   37
            Top             =   225
            Width           =   1755
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Incluir"
         Height          =   765
         Left            =   3255
         TabIndex        =   29
         Top             =   450
         Width           =   1590
         Begin VB.CheckBox chkCancelados 
            Caption         =   "Cancelados"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   180
            TabIndex        =   35
            Top             =   220
            Width           =   1170
         End
         Begin VB.CheckBox chkFacturados 
            Caption         =   "Facturados"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   180
            TabIndex        =   30
            Top             =   465
            Width           =   1170
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBuscaFacturas 
         Height          =   5325
         Left            =   -74820
         TabIndex        =   47
         Top             =   1290
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   9393
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSCommLib.MSComm comPrinter 
         Left            =   4560
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBuscaTickets 
         Height          =   5325
         Left            =   180
         TabIndex        =   48
         Top             =   1290
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   9393
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A "
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   -66720
         TabIndex        =   60
         Top             =   6720
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
         TabIndex        =   59
         Top             =   6720
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74400
         TabIndex        =   58
         Top             =   6735
         Width           =   840
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
         Left            =   -74760
         TabIndex        =   57
         Top             =   7500
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
         TabIndex        =   56
         Top             =   7245
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de cancelar ante el SAT"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   -74400
         TabIndex        =   55
         Top             =   6990
         Width           =   2565
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " A "
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   14
         Left            =   -74760
         TabIndex        =   54
         Top             =   6975
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de autorización de cancelación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -74400
         TabIndex        =   53
         Top             =   7260
         Width           =   3060
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación rechazada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74400
         TabIndex        =   52
         Top             =   7515
         Width           =   1680
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Facturas pendientes de timbre fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   -66435
         TabIndex        =   50
         Top             =   6735
         Width           =   2610
      End
      Begin VB.Label lbTipoPacienteEmpresa 
         Caption         =   "Tipo de paciente"
         Height          =   255
         Left            =   7080
         TabIndex        =   49
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmConsultaPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmConsultaPOS                                         -
'-------------------------------------------------------------------------------------
'| Objetivo: Es la conculta del
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 25/Oct/2001
'| Modificó                 : Nombre(s)
'| Fecha Terminación        :
'| Fecha última modificación: 14/May/2002
'-------------------------------------------------------------------------------------

Option Explicit

Private vgrptReporte As CRAXDRT.Report

Dim vlaryParametros() As String
Dim llngEmpresaTipoPaciente As Long 'Para saber si se está facturando a algun convenio
Dim llngFormatoFactura As Long
Dim intTipoEmisionComprobante As Integer 'Variable que compara el tipo de formato y folio a utilizar (0 = Error de formato y folios incompatibles, 1 = Físicos, 2 = Digitales)
Dim intTipoCFDFactura As Integer
Dim lstrLeyendaCliente As String    'Para la impresión del ticket

Dim aFormasPago() As FormasPago     'Para guardar los movimientos de los pagos

Dim vlstrTPaciente As String
Dim vllngExp As Long
Dim vllngSeleccionadas As Long
Dim vllngSeleccPendienteTimbre As Long
Dim vlblnLicenciaIEPS As Boolean
Dim aMovimientoBancoForma() As String
Dim vlintContadorMovs As Integer
Dim blnNoactivate As Boolean
Dim vlintUsoCFDI As Long

Dim lngCveFormato As Long           'Identifica el tipo de formato a utilizar
Dim vlngCveFormato As Long          'Tipo de formato
Dim vgintTipoPaciente As Integer
Dim vllngFormatoaUsar As Long       'Para saber que formato se va a utilizar
Dim blnLicenciaLealtadCliente As Boolean    'Para saber si se tiene licencia para generar la lealtad del cliente y el médico
Dim vllngNumeroPaciente As Long
Dim vlblnEsGlobal As Boolean

Private Sub pDatosPaciente(vllngNumCuenta As Long)
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim vlstrSentencia As String
    
    vllngNumeroPaciente = 0
    vlstrSentencia = "SELECT * FROM Expacienteingreso where expacienteingreso.INTNUMCUENTA = " & vllngNumCuenta
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        vllngNumeroPaciente = rs!intNumPaciente
    End If
    If rs.State <> adStateClosed Then rs.Close
                    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDatosPaciente"))
    Unload Me
End Sub

Public Function fblnCuentaRelacionadaComBancaria(lngCuenta As Long, strFolio As String) As Boolean
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    '----------------------------------------------------------------------'
    ' Verifica si la cuenta está relaconada con el concepto de facturación '
    ' o con la cuenta del iva cobrado
    '----------------------------------------------------------------------'
    
    fblnCuentaRelacionadaComBancaria = False
             
    If intBitCuentaPuenteBanco = 1 Then
        strSQL = "select 1 from dual where " & lngCuenta & _
                        " in (/*(select case when PvConceptoFacturacionDepartame.intnumcuentaingreso is null then " & _
                                        "pvconceptofacturacionempresa.intnumctaingreso " & _
                                    "else pvconceptofacturaciondepartame.intnumcuentaingreso end " & _
                               "from pvventapublico " & _
                                    "inner join pvdetalleventapublico on pvventapublico.intcveventa = pvdetalleventapublico.intcveventa " & _
                                    "left join pvconceptofacturacionempresa on pvdetalleventapublico.smicveconceptofacturacion = pvconceptofacturacionempresa.intcveconceptofactura " & _
                                    "left join pvconceptofacturaciondepartame on pvdetalleventapublico.smicveconceptofacturacion = pvconceptofacturaciondepartame.smicveconcepto " & _
                                                                            "and pvconceptofacturaciondepartame.smicvedepartamento = pvventapublico.intcvedepartamento " & _
                              "where trim(pvventapublico.chrfolioticket) = '" & strFolio & "'),*/ " & _
                             "(select vchvalor from siparametro " & _
                               "where vchnombre = 'INTCTAIVACOBRADO' " & _
                                 "and intcveempresacontable = " & vgintClaveEmpresaContable & "))"
                            
        Set rs = frsRegresaRs(strSQL)
        If rs.RecordCount <> 0 Then
            fblnCuentaRelacionadaComBancaria = True
        End If
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaRelacionadaComBancaria"))
End Function
Private Sub pCargaBusqueda()
    
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim lngAux As Long
    Dim lngAncho As Long
    Dim vlBackColor As Variant
    Dim vlForeColor As Variant


    freCargando.Top = 2500
    freCargando.Visible = True
    freCargando.Refresh
    
    grdBuscaFacturas.Redraw = False
    pLimpiaGrid grdBuscaFacturas
    pConfiguraGridBusqueda
    vllngSeleccionadas = 0
    vllngSeleccPendienteTimbre = 0
        lngAncho = 1000
        lngAux = 0
    vgstrParametrosSP = _
                        "-1" & _
                        "|" & IIf(Not optMostrarSolo(0).Value, fstrFechaSQL("01/01/2010"), fstrFechaSQL(mskFechaInicial.Text)) & _
                        "|" & IIf(Not optMostrarSolo(0).Value, fstrFechaSQL(fdtmServerFecha), fstrFechaSQL(mskFechaFinal.Text)) & _
                        "|" & "1" & _
                        "|" & "-1" & _
                        "|" & "A" & _
                        "|" & "-1" & _
                        "|" & CStr(vgintNumeroDepartamento) & _
                        "|" & vgintClaveEmpresaContable & _
                        "|" & IIf(optMostrarSolo(2).Value, 1, 0) & _
                        "|" & IIf(optMostrarSolo(1).Value, 1, 0) & _
                        "|" & IIf(optMostrarSolo(3).Value, 1, 0) & _
                        "|" & IIf(optMostrarSolo(4).Value, 1, 0)
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFacturaFiltro_NE")

    
    
    If rs.RecordCount > 10000 Then
        'El numero de registros es demasiado grande, sólo se mostrarán los primeros 10000. Pruebe con un rango de fechas menor...
        MsgBox SIHOMsg(403), vbInformation, "Mensaje"
    End If

    Do While Not rs.EOF
        With grdBuscaFacturas
            If .RowData(1) <> -1 Then
                 .Rows = .Rows + 1
                 .Row = .Rows - 1
            End If
            .RowData(.Row) = rs!IdFactura
            .TextMatrix(.Row, 0) = IIf(rs!PendienteTimbreFiscal = 1 Or (rs!PendienteCancelarSAT_NE <> "NP" And rs!PendienteCancelarSAT_NE <> "CR"), "*", "")
            '|.TextMatrix(.Row, 0) = IIf(rs!PendienteTimbreFiscal = 1, "*", "")
            
            .TextMatrix(.Row, 1) = rs!folio
            'ajustar la columna de los folios-------------
      
             PB.Font = .CellFontName
   
             PB.FontSize = .CellFontSize
    
             lngAux = PB.TextWidth(.TextMatrix(.Row, 1))
             If lngAux > lngAncho Then
                lngAncho = lngAux
             End If
           
            .TextMatrix(.Row, 2) = IIf(IsNull(rs!Paciente), "", rs!Paciente)
            .TextMatrix(.Row, 3) = rs!RazonSocial
            .TextMatrix(.Row, 4) = rs!fecha
            .TextMatrix(.Row, 5) = FormatCurrency(rs!Descuento, 2)
            .TextMatrix(.Row, 6) = FormatCurrency(rs!IEPS, 2)
            .TextMatrix(.Row, 7) = FormatCurrency(rs!TotalFactura - rs!IVA, 2)
            .TextMatrix(.Row, 8) = FormatCurrency(rs!IVA, 2)
            .TextMatrix(.Row, 9) = FormatCurrency(rs!TotalFactura, 2)
            .TextMatrix(.Row, 10) = rs!Moneda
            .TextMatrix(.Row, 11) = IIf(IsNull(rs!PersonaFacturo), "", rs!PersonaFacturo)
            .TextMatrix(.Row, 12) = IIf(IsNull(rs!PersonaCancelo), "", rs!PersonaCancelo)
            .TextMatrix(.Row, 13) = rs!PendienteCancelarSAT_NE
            .TextMatrix(.Row, 14) = rs!PendienteTimbreFiscal
            
            If (.TextMatrix(.Row, 13) <> "NP" And .TextMatrix(.Row, 13) <> "CR") Then vllngSeleccionadas = vllngSeleccionadas + 1
            
            '|If rs!PendienteTimbreFiscal > 0 Then vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
            
            .Col = 0
            If rs!chrEstatus = "C" Then
            
                If rs!PendienteCancelarSat = 1 Then
                    vlForeColor = &HFF&    '| Rojo
                    vlBackColor = &HC0E0FF '| Naranja suave
                Else
                    vlForeColor = &HFF&    '| Rojo
                    vlBackColor = &HFFFFFF '| Blanco
                End If
                
                '| La primer columna se pinta de rojo
                .Col = 1
                .CellForeColor = &HFF&
                '| Las demás columnas se pintan dependiendo del estado que guarde la factura
                For vlintContador = 3 To .Cols
                    .Col = .Col + 1
                    .CellForeColor = vlForeColor
                    .CellBackColor = vlBackColor
                Next
            Else
                Select Case rs!PendienteCancelarSAT_NE
                    Case "PA"
                        vlForeColor = &HFFFFFF '| Blanco
                        vlBackColor = &H80FF&  '| Naranja fuerte
                    Case "CR"
                        vlForeColor = &HFFFFFF '| Blanco
                        vlBackColor = &HFF&    '| Rojo
                    Case "NP"
                        vlForeColor = &H0&     '| Negro
                        vlBackColor = &HFFFFFF '| Blanco
                End Select
                
                '| La primer columna se pinta de negro
                .Col = 1
                .CellForeColor = &H0&
                '| Las demás columnas se pintan dependiendo del estado que guarde la factura
                For vlintContador = 3 To .Cols
                    .Col = .Col + 1
                    .CellForeColor = vlForeColor
                    .CellBackColor = vlBackColor
                Next
            
            
                If rs!PendienteTimbreFiscal = 1 Then
                   .Col = 1
                   For vlintContador = 3 To .Cols
                       .Col = .Col + 1
                       .CellBackColor = &H80FFFF 'amarillo
                   Next
                End If
            End If

                  
        End With
        rs.MoveNext
    Loop

    If lngAncho > 1000 Then grdBuscaFacturas.ColWidth(1) = lngAncho + 100
    grdBuscaFacturas.Redraw = True
    rs.Close
    freCargando.Visible = False
    'If Me.ActiveControl.Name = "cmdCancelaFacturasSAT" And vllngSeleccionadas = 0 Then grdBuscaFacturas.SetFocus
    cmdCancelaFacturasSAT.Enabled = vllngSeleccionadas > 0
    'If Me.ActiveControl.Name = "Cmdconfirmartimbre" And vllngSeleccPendienteTimbre = 0 Then grdBuscaFacturas.SetFocus
    Cmdconfirmartimbre.Enabled = vllngSeleccPendienteTimbre > 0
End Sub
Private Sub pConfiguraGridBusqueda()
    With grdBuscaFacturas
 .Cols = 15
        .FixedCols = 2
        .FixedRows = 1
        .FormatString = "|Factura|Paciente|Facturado a|Fecha|Descuento|IEPS|Subtotal|IVA|Total|Moneda|Empleado que facturó|Empleado que canceló"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 1000 'Numero de Factura
        .ColWidth(2) = 4000 'Paciente
        .ColWidth(3) = 3400 'Facturado a:
        .ColWidth(4) = 1430 'Fecha
        '------------------------------------------------------
        .ColWidth(5) = 1300 'Descuento
        .ColWidth(6) = IIf(vlblnLicenciaIEPS, 1300, 0) 'IEPS
        .ColWidth(7) = 1300 'Subtotal
        .ColWidth(8) = 1460 'Iva
        .ColWidth(9) = 1430 'Total
        '------------------------------------------------------
        .ColWidth(10) = 1000 'Moneda
        .ColWidth(11) = 4000 'Empleado Facturo
        .ColWidth(12) = 4000 'Empleado Cancelo
        .ColWidth(13) = 0 ' si la factura esta pendiente de capturar en la SAT
        .ColWidth(14) = 0 'si esta pendiente de timbre fiscal
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignLeftCenter
        .ColAlignment(11) = flexAlignLeftCenter
        .ColAlignment(12) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignLeftCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .ColAlignmentFixed(12) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .Clear
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

Private Sub cboProcedencia_Click()
On Error GoTo NotificaError

    If mskFechaInicial = "  /  /    " Then
        mskFechaInicial = fdtmServerFecha
    End If
    
    If mskFechaFinal = "  /  /    " Then
        mskFechaFinal = fdtmServerFecha
    End If
        
    sstPOS_Click sstPOS.Tab
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProcedencia_Click"))
End Sub

Private Sub cboProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If sstPOS.Tab = 0 And KeyCode = 13 Then
        grdBuscaTickets.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboProcedencia_KeyDown"))
End Sub

Private Sub chkFacturados_Click()
    If mskFechaFinal = "  /  /    " Then
        mskFechaFinal = fdtmServerFecha
    End If
    sstPOS_Click sstPOS.Tab
End Sub




Private Sub cmdCancelaFacturasSAT_Click()
'Cancelacion maciva de facturas ante el SAT, cancelacion del XML
Dim vlLngCantidadFacturas As Long
Dim vlLngCont As Long
Dim vllngPersonaGraba As Long
Dim vlintFacturasCanceladas As Long

On Error GoTo NotificaError

    '|  Los comprobantes seleccionados serán validados nuevamente ante el SAT.
    '|  ¿Desea continuar?
    If MsgBox(SIHOMsg(1249), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
        'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
        With grdBuscaFacturas
            vlLngCantidadFacturas = 0
            vlintFacturasCanceladas = 0
            For vlLngCont = 1 To .Rows - 1
                If .TextMatrix(vlLngCont, 0) = "*" And (.TextMatrix(vlLngCont, 13) <> "NP" And .TextMatrix(vlLngCont, 13) <> "CR") Then
                    vlLngCantidadFacturas = vlLngCantidadFacturas + 1
                    If fblnFacturaCancelable(grdBuscaFacturas.TextMatrix(vlLngCont, 1)) Then
                        pCancelaCFDiFacturaSiHO grdBuscaFacturas.TextMatrix(vlLngCont, 1), grdBuscaFacturas.TextMatrix(vlLngCont, 13), vllngPersonaGraba, 0, Me.Name, False
                        vlintFacturasCanceladas = vlintFacturasCanceladas + 1
                    End If
                End If
            Next vlLngCont
        End With
        If vlLngCantidadFacturas = vlintFacturasCanceladas Then
            '|  La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
        End If
        mskFechaFinal_KeyDown vbKeyReturn, 0
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelaFacturasSAT_Click"))
    Unload Me
End Sub
Private Sub cmdCerrarMuestraTicket_Click()
    sstPOS.Enabled = True
    Frame13.Enabled = True
    grdBuscaTickets.SetFocus
    freMuestraTicket.Visible = False
End Sub

Private Sub cmdConfirmartimbre_Click()
Dim vlLngCont As Long
Dim vllngPersonaGraba As Long
Dim vlngReg As Long

On Error GoTo NotificaError

blnNOMensajeErrorPAC = False 'de inicio siempre a False

'Los comprobantes seleccionados se encuentran pendientes de timbre fiscal ¿Desea confirmar el timbre fiscal?
If MsgBox(SIHOMsg(1310), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
   If vllngPersonaGraba = 0 Then Exit Sub
   'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
   With grdBuscaFacturas
        For vlLngCont = 1 To .Rows - 1
            If .TextMatrix(vlLngCont, 0) = "*" And .TextMatrix(vlLngCont, 14) = 1 Then
               pLogTimbrado 2
               pgbBarraCFD.Value = 70
               freBarraCFD.Top = 3200
               Screen.MousePointer = vbHourglass
               lblTextoBarraCFD.Caption = "Confirmando timbre fiscal, por favor espere..."
               freBarraCFD.Visible = True
               freBarraCFD.Refresh
               blnNOMensajeErrorPAC = True
               EntornoSIHO.ConeccionSIHO.BeginTrans
               If .TextMatrix(vlLngCont, 13) = "2" Or .TextMatrix(vlLngCont, 13) = "XX" Then
                  vgIntBanderaTImbradoPendiente = 0
                  vlngReg = flngRegistroFolio("FA", .RowData(vlLngCont))
                  If Not fblnGeneraCFDpCancelacion(.RowData(vlLngCont), "FA", fblnTCFDi(vlngReg), 0) Then
                     On Error Resume Next
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar/o que no alcanzó a llegar el timbrado
                           'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                           MsgBox Replace(SIHOMsg(1314), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbInformation + vbOKOnly, "Mensaje"
                           'la factura se queda igual, no se hace nada
                        ElseIf vgIntBanderaTImbradoPendiente = 2 Then
                             'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
                              MsgBox Replace(SIHOMsg(1313), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbExclamation + vbOKOnly, "Mensaje"
                              'Aqui se debe de cancelar la factura que vas a eliminar?
                              pEjecutaSentencia " UPDATE PVCANCELARCOMPROBANTES SET BITPENDIENTECANCELAR = 0 " & _
                              "Where INTCOMPROBANTE =" & .RowData(vlLngCont) & " And VCHTIPOCOMPROBANTE = 'FA'"
                              'se quita de los pendientes de timbre
                              pEliminaPendientesTimbre .RowData(vlLngCont), "FA"
                        End If
                  Else 'se confirmo el timbre correctamente
                          'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                           pEliminaPendientesTimbre .RowData(vlLngCont), "FA"
                           EntornoSIHO.ConeccionSIHO.CommitTrans
                           'Timbre fiscal de factura <FOLIO>: Confirmado.
                           MsgBox Replace(SIHOMsg(1315), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbInformation + vbOKOnly, "Mensaje"
                  End If
               Else
                  vlngReg = flngRegistroFolio("FA", .RowData(vlLngCont))
                  If Not fblnGeneraComprobanteDigital(.RowData(vlLngCont), "FA", 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
                      On Error Resume Next
                      EntornoSIHO.ConeccionSIHO.RollbackTrans
                      If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar/o que no alcanzó a llegar el timbrado
                          'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                          MsgBox Replace(SIHOMsg(1314), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbInformation + vbOKOnly, "Mensaje"
                          'la factura se queda igual, no se hace nada
                      ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
                          'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
                          MsgBox Replace(SIHOMsg(1313), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbExclamation + vbOKOnly, "Mensaje"
                          'Aqui se debe de cancelar la factura
                          pCancelarFactura Trim(.TextMatrix(vlLngCont, 1)), vllngPersonaGraba, Me.Name
                      End If
                  Else
                      'Se guarda el LOG
                      Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & .TextMatrix(vlLngCont, 1))
                      'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                      pEliminaPendientesTimbre .RowData(vlLngCont), "FA"
                      'Commit
                      EntornoSIHO.ConeccionSIHO.CommitTrans
                      'Timbre fiscal de factura <FOLIO>: Confirmado.
                      MsgBox Replace(SIHOMsg(1315), "<FOLIO>", Trim(.TextMatrix(vlLngCont, 1))), vbInformation + vbOKOnly, "Mensaje"
                  End If
               End If
               'Barra de progreso CFD
               pgbBarraCFD.Value = 100
               freBarraCFD.Top = 3200
               Screen.MousePointer = vbDefault
               freBarraCFD.Visible = False
               'guardamos el log del timbrado
               pLogTimbrado 1
            End If
        Next vlLngCont
   End With
      
   blnNOMensajeErrorPAC = False
   pCargaBusqueda
   grdBuscaFacturas.SetFocus
      
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub

Private Sub cmdFacturaSeleccion_Click()

    If fblnValidaSeleccion Then
        PGrabarFactura
        pCargaBusquedaT 0, grdBuscaTickets
    End If
End Sub

Private Function fblnValidaSeleccion() As Boolean
    Dim intCont As Integer
    Dim blnPrimeraVez As Boolean
    
    grdBuscaTickets.Col = 1
    fblnValidaSeleccion = True
    blnPrimeraVez = True
    For intCont = 1 To grdBuscaTickets.Rows - 1
        grdBuscaTickets.Row = intCont
        If (grdBuscaTickets.CellForeColor = vbRed Or grdBuscaTickets.CellForeColor = vbBlue) And grdBuscaTickets.TextMatrix(intCont, 0) = "*" Then
            If blnPrimeraVez Then
                If MsgBox("Se encuentran seleccionados tickets facturados y/o cancelados los cuales serán ignorados." & vbCrLf & "¿Desea continuar?", vbExclamation + vbYesNo, "Mensaje") = vbNo Then
                    fblnValidaSeleccion = False
                    Exit For
                Else
                    grdBuscaTickets.TextMatrix(intCont, 0) = ""
                    blnPrimeraVez = False
                End If
            Else
                grdBuscaTickets.TextMatrix(intCont, 0) = ""
            End If
        End If
    Next
    
    If fblnValidaSeleccion = True Then
        If optFormaPago(0).Value = True Then
            fblnValidaSeleccion = False
            MsgBox "Seleccione la forma de pago en que fueron emitidos los tickets.", vbExclamation, "Mensaje"
            optFormaPago(0).SetFocus
        End If
    End If
End Function

Private Function flngSeleccionValida() As Long
    'Regresa el número de mensaje cuando la selección de los tickets a facturar no es correcta, razones:
    'cuando los tickets pertenecen a pacientes de diferentes convenios o tipos de paciente
    'cuando fueron facturados a crédito y tienen pagos registrados
    'cuando pertenecen al mismo tipo de paciente pero con distinto nombre
    Dim vllngContador As Long
    Dim vlblnError As Boolean
    Dim vlblnPrimero As Boolean
    Dim rsDatos As New ADODB.Recordset
    Dim lngCveTipoPaciente As Long
    Dim lstrNombrePaciente As String
    flngSeleccionValida = 0
    
    llngEmpresaTipoPaciente = 0
    vlblnError = False
    vlblnPrimero = True
    vllngContador = 1
    
    Do While vllngContador <= grdBuscaTickets.Rows - 1 And Not vlblnError
        If grdBuscaTickets.TextMatrix(vllngContador, 0) = "*" Then
            If vlblnPrimero Then
                llngEmpresaTipoPaciente = Val(grdBuscaTickets.TextMatrix(vllngContador, 14))
                lstrNombrePaciente = grdBuscaTickets.TextMatrix(vllngContador, 9)
                If Val(grdBuscaTickets.TextMatrix(vllngContador, 18)) > 0 Then   'cuenta paciente
                    vgstrParametrosSP = Trim(grdBuscaTickets.TextMatrix(grdBuscaTickets.Row, 18)) & "|0|E|" & vgintClaveEmpresaContable
                    Set rsDatos = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
                    If rsDatos.RecordCount > 0 Then
                        lngCveTipoPaciente = rsDatos!tnyCveTipoPaciente
                    End If
                    If llngEmpresaTipoPaciente > 0 Then  'el paciente pertenece a una empresa
                        vgstrParametrosSP = vgintNumeroDepartamento & "|" & llngEmpresaTipoPaciente & "|" & lngCveTipoPaciente & "|E"
                    Else
                        vgstrParametrosSP = vgintNumeroDepartamento & "|0|" & lngCveTipoPaciente & "|E"
                    End If
                Else
                    vgstrParametrosSP = vgintNumeroDepartamento & "|0|0|E"
                End If
                llngFormatoFactura = 1
                frsEjecuta_SP vgstrParametrosSP, "Fn_Pvselformatofactura", False, llngFormatoFactura
                If llngFormatoFactura = 0 Then
                    'No se encontró un formato válido de factura, por favor de uno de alta.
                    flngSeleccionValida = 373
                    vlblnError = True
                End If
                vlblnPrimero = False
            End If
            If Not vlblnError Then
                If Val(grdBuscaTickets.TextMatrix(vllngContador, 14)) <> llngEmpresaTipoPaciente Then
                    'No se pueden facturar tickets que pertenecen a diferentes convenios o tipos de paciente.
                    flngSeleccionValida = 660
                    vlblnError = True
                Else
                    If Val(grdBuscaTickets.TextMatrix(vllngContador, 15)) <> 0 Then
                        'No se pueden facturar tickets que tienen pagos registrados en crédito.
                        flngSeleccionValida = 661
                        vlblnError = True
                    End If
                    If grdBuscaTickets.TextMatrix(vllngContador, 9) <> lstrNombrePaciente Then
                        'No se puede generar una factura con tickets de diferentes clientes.
                        flngSeleccionValida = 1367
                        vlblnError = True
                    End If
                End If
            End If
        End If
        
        vllngContador = vllngContador + 1
    Loop
End Function

Private Sub pDatosFiscalesPvParametros()
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "SELECT CHRRFCPOS RFC, null Clave, CHRNOMBREFACTURAPOS Nombre, CHRDIRECCIONPOS chrCalle," & _
                     "VCHNUMEROEXTERIORPOS vchNumeroExterior, VCHNUMEROINTERIORPOS vchNumeroInterior, " & _
                     "null Telefono,'OT' Tipo,  INTCVECIUDAD IdCiudad, VCHCOLONIAPOS Colonia, VCHCODIGOPOSTALPOS CP FROM PVParametro " & _
                     "WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic)
    
    If rs.RecordCount > 0 Then
            With frmDatosFiscales
                    .vgstrNombre = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
                    .vgstrDireccion = IIf(IsNull(rs!CHRCALLE), "", Trim(rs!CHRCALLE))
                    .vgstrNumExterior = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                    .vgstrNumInterior = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                    .vgstrColonia = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
                    .vgstrCP = IIf(IsNull(rs!CP), "", Trim(rs!CP))
                    .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, Str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
                    .vgstrTelefono = IIf(IsNull(rs!Telefono), "", Trim(rs!Telefono))
                    .vgstrRFC = IIf(IsNull(rs!RFC), "", Trim(rs!RFC))
                    .vlstrNumRef = IIf(rs!tipo = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
                    .vlstrTipo = IIf(IsNull(rs!tipo), "OT", rs!tipo)
                    .vglngDatosParametro = True
                    .vgActivaSujetoaIEPS = vlblnLicenciaIEPS
                    .vgBitSujetoaIEPS = 0
            End With
    Else
        frmDatosFiscales.sstDatos.Tab = 1
    End If
End Sub
Private Function pRegresaDatosFiscales() As Boolean
        Dim rs As New ADODB.Recordset
        Dim vlstrSentencia As String
        Dim vllngContador As Long
        Dim vlblnPrimerDato As Boolean
        Dim vllngTipoPaciente As Long
        Dim vllngCuentaPrimerDato As Long
        Dim vlblnDatosFiscalesEnBlanco As Boolean ' si los tickets pertenecen a diferentes personas o empresas ¿ como saber cuales datos fiscales son los correctos?
        Dim vllngExPaciente As Long
        Dim vlstrChrtipoPaciente As String
                
        vlblnPrimerDato = True ' para tomar los datos del primer renglon seleccionado
        vlblnDatosFiscalesEnBlanco = False
        pRegresaDatosFiscales = True
        
        vlstrTPaciente = "OT" ' variable que se utiliza par ver si se actualizan o no los datos fiscales en expaciente
        vllngExp = 0
        
        Do While vllngContador <= grdBuscaTickets.Rows - 1
           If grdBuscaTickets.TextMatrix(vllngContador, 0) = "*" Then
                If vlblnPrimerDato Then
                      vlblnPrimerDato = False 'colocamos en false la variable para que no vuelva a entrar aqui
                      vllngTipoPaciente = Val(grdBuscaTickets.TextMatrix(vllngContador, 14))
                      vllngCuentaPrimerDato = Val(grdBuscaTickets.TextMatrix(vllngContador, 18))
                      vlstrSentencia = "select INTNUMPACIENTE from EXPACIENTEINGRESO where INTNUMCUENTA =" & grdBuscaTickets.TextMatrix(vllngContador, 18)
                      Set rs = frsRegresaRs(vlstrSentencia)
                      If rs.RecordCount > 0 Then
                         vllngExPaciente = rs!intNumPaciente
                         vllngExp = vllngExPaciente  ' torcemos el numero de expediente
                      Else
                         vllngExPaciente = 0
                         vlblnDatosFiscalesEnBlanco = True 'Datos fiscales en blanco
                         vllngContador = grdBuscaTickets.Rows ' para que salga del ciclo, ya no se requiere buscar mas
                          'pRegresaDatosFiscales = False
                          Load frmDatosFiscales
                          pDatosFiscalesPvParametros
                          Exit Function
                      End If
                   Else ' si no es el primer renglon
                      If Val(grdBuscaTickets.TextMatrix(vllngContador, 14)) <> vllngTipoPaciente Then
                         vlblnDatosFiscalesEnBlanco = True 'Datos fiscales en blanco
                         vllngContador = grdBuscaTickets.Rows ' para que salga del ciclo, ya no se requiere buscar mas
                         pRegresaDatosFiscales = False
                      Else ' mismo tipo de paciente, se debe revisar si es el mismo paciente, esto para el caso de que no sea empresa
                         If vllngTipoPaciente < 0 Then ' ES UN TIPO  CONVENIO
                            vlstrSentencia = "select INTNUMPACIENTE from EXPACIENTEINGRESO where INTNUMCUENTA =" & grdBuscaTickets.TextMatrix(vllngContador, 18)
                            Set rs = frsRegresaRs(vlstrSentencia)
                            If rs.RecordCount > 0 Then
                               If vllngExPaciente <> rs!intNumPaciente Then
                                  vlblnDatosFiscalesEnBlanco = True 'Datos fiscales en blanco
                                  vllngContador = grdBuscaTickets.Rows ' para que salga del ciclo, ya no se requiere buscar mas
                                  pRegresaDatosFiscales = False
                               End If
                            Else
                               vlblnDatosFiscalesEnBlanco = True 'Datos fiscales en blanco
                               vllngContador = grdBuscaTickets.Rows ' para que salga del ciclo, ya no se requiere buscar mas
                               pRegresaDatosFiscales = False
                            End If
                         End If
                      End If
                   End If
           End If
           vllngContador = vllngContador + 1
        Loop
        Load frmDatosFiscales ' cargamos la pantalla de datos fiscales
        '' ahora para ver si cargamos datos fiscales o no
        
        If Not vlblnDatosFiscalesEnBlanco Then
            ' primero vemos si es de un convenio
            If vllngTipoPaciente > 0 Then ' es un convenio se trae la informacion de ccempresa
                vlstrChrtipoPaciente = "CO"
              
                vlstrSentencia = "SELECT chrRfcEmpresa RFC, intCveEmpresa Clave, vchRazonSocial Nombre, chrCalle, vchNumeroExterior,vchNumeroInterior, chrTelefonoEmpresa Telefono, 'CO' Tipo, CcEmpresa.intCveCiudad IdCiudad, CCEmpresa.vchColonia Colonia, TO_CHAR(CCEmpresa.vchCodigoPostal) CP, VCHREGIMENFISCAL FROM CCEmpresa  WHERE intCveEmpresa = " & CStr(vllngTipoPaciente)
                Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic)
                If rs.RecordCount > 0 Then
                    With frmDatosFiscales
                            .vgstrNombre = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
                            .vgstrDireccion = IIf(IsNull(rs!CHRCALLE), "", Trim(rs!CHRCALLE))
                            .vgstrNumExterior = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                            .vgstrNumInterior = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                            .vgstrColonia = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
                            .vgstrCP = IIf(IsNull(rs!CP), "", Trim(rs!CP))
                            .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, Str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
                            .vgstrTelefono = IIf(IsNull(rs!Telefono), "", Trim(rs!Telefono))
                            .vgstrRFC = IIf(IsNull(rs!RFC), "", Trim(rs!RFC))
                            .vlstrNumRef = IIf(vlstrChrtipoPaciente = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
                            .vlstrTipo = vlstrChrtipoPaciente
                            .vgBitSujetoaIEPS = 0
                            
                            .vlstrRegimenFiscal = rs!VCHREGIMENFISCAL
                            
                            .vgActivaSujetoaIEPS = vlblnLicenciaIEPS
                    End With
                End If
            Else ' se trae la informacion de externo
                vlstrSentencia = "select chrtipo from ADTIPOPACIENTE where TNYCVETIPOPACIENTE = " & CStr((vllngTipoPaciente * -1))
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount > 0 Then
                   vlstrChrtipoPaciente = rs!chrTipo
                Else
                   vlstrChrtipoPaciente = "OT"
                End If
                   
                vlstrSentencia = "SELECT chrRFC RFC, RegistroExterno.INTNUMCUENTA Clave, RTRIM(chrApePaterno) || ' ' || RTRIM(chrApeMaterno) || ' ' || RTRIM(chrNombre) Nombre, CHRCALLE, VCHNUMEROEXTERIOR, VCHNUMEROINTERIOR, chrTelefono Telefono, 'PE' Tipo, Externo.intCiudad IdCiudad, Externo.vchColonia Colonia, Externo.vchCodPostal CP " & _
                                " FROM Externo LEFT OUTER JOIN RegistroExterno ON (Externo.INTNUMPACIENTE = RegistroExterno.INTNUMPACIENTE AND RegistroExterno.DTMFECHAEGRESO IS NULL) WHERE Externo.intNumPaciente = " & vllngExPaciente
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount > 0 Then
                    ' cargamos los datos fiscales
                    With frmDatosFiscales
                        .vgstrNombre = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
                        .vgstrDireccion = IIf(IsNull(rs!CHRCALLE), "", Trim(rs!CHRCALLE))
                        .vgstrNumExterior = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                        .vgstrNumInterior = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                        .vgstrColonia = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
                        .vgstrCP = IIf(IsNull(rs!CP), "", Trim(rs!CP))
                        .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, Str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
                        .vgstrTelefono = IIf(IsNull(rs!Telefono), "", Trim(rs!Telefono))
                        .vgstrRFC = IIf(IsNull(rs!RFC), "", Trim(rs!RFC))
                        .vlstrNumRef = IIf(vlstrChrtipoPaciente = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
                        .vlstrTipo = vlstrChrtipoPaciente
                        .vgBitSujetoaIEPS = 0
                        .vgActivaSujetoaIEPS = vlblnLicenciaIEPS
                        
                        
                        'para saber si se actualizan los datos fiscales o no en expaciente (sólo EM y ME)
                        'sólo de aqui se puede obtener este dato para la facturación
                        vlstrTPaciente = vlstrChrtipoPaciente
                    End With
                End If
            End If
       End If
End Function
Private Sub pRellenaInfoGlobal()
    
    Dim i As Integer
    Dim j As Integer
    Dim dtmFechaMinima As Date
    Dim dtmFechaMaxima As Date
    Dim intDiferencia As Date
    
    i = 1
    j = 1
    
    'dtmFechaMinima = aTickets(i).dtmfecha
    For i = 1 To grdBuscaTickets.Rows - 1
        If grdBuscaTickets.TextMatrix(i, 0) = "*" Then
            dtmFechaMinima = CDate(grdBuscaTickets.TextMatrix(i, 3))
            Exit For
        End If
    Next i

    'dtmFechaMaxima = aTickets(j).dtmfecha
    For i = 1 To grdBuscaTickets.Rows - 1
        If grdBuscaTickets.TextMatrix(i, 0) = "*" Then
            dtmFechaMaxima = CDate(grdBuscaTickets.TextMatrix(i, 3))
        End If
    Next i

    intDiferencia = DateDiff("D", dtmFechaMinima, dtmFechaMaxima)
    
    vgStrAñoGlobal = Year(dtmFechaMaxima)
    vgStrMesesGlobal = Month(dtmFechaMaxima)
    If intDiferencia <= 1 Then
        vgStrPeriodicidad = "01"
    ElseIf intDiferencia <= 7 Then
        vgStrPeriodicidad = "02"
    ElseIf intDiferencia <= 15 Then
        vgStrPeriodicidad = "03"
    Else
        vgStrPeriodicidad = "04"
    End If
End Sub
Private Sub PGrabarFactura()
    Dim vlintContador As Integer                 'Para los ciclos
    Dim vlintConta2 As Integer                   'Igual que el de arriba, pero para otros ciclos, jeje
    Dim vlintPosicion As Integer                 'Para control de la posicion en el grid de Articulos
    Dim rsFormasPagos As New ADODB.Recordset     'Sólo para traer las formas de pago de cada pago
    Dim rstipoformapago As New ADODB.Recordset   'Para identificar las formas de pago tipo credito
    Dim rsFactura As New ADODB.Recordset         'RS tipo tabla para guardar la fractura
    Dim rsDetalleFactura As New ADODB.Recordset  'RS tipo tabla para el Detalle de la FACTURA
    Dim rsPvDocumentoCancelado As New ADODB.Recordset
    Dim rsDatosCliente As New ADODB.Recordset    'RS para los datos del cliente
    Dim rsTemp As New ADODB.Recordset            'RS para cualquier operación
    Dim rsDetalleVentaPublico As New ADODB.Recordset  'RS para Obtener los cargos de un ticket
    Dim vlstrSentencia As String                 'Sirve pa TODOS los RS's
    Dim vllngNumeroCorte As Long                 'Trae el numero de corte actual
    Dim a As Integer                             'Contador
    Dim vldblTipoCambio As Double                'Trae el tipo de cambio utilizado en la pantalla de Formas de pago
    Dim vlblnbandera As Boolean                  'Banderilla para control de flujo
    Dim vlstrFolioDocumento As String            'Este es el numero de factura a utilizar
    Dim vllngConsecutivoFactura As Long
    Dim rsPolizaTicket As New ADODB.Recordset    'Cargar los movimientos contables que hizo el ticket
    Dim rsFormasTicket As New ADODB.Recordset    'Cargar las formas de pago del ticket
    
    Dim vllngPersonaGraba As Long                'Persona que esta generando la factura
    Dim vllngFoliosFaltantes As Long             'Para el control de los folios faltantes
    ' Subtotales de la Factura
    Dim vldblSubtotal As Double                  'Subtotal de la cuenta, para la factura
    Dim vldblIVA As Double                       'IVA total de la cuenta, para la factura
    Dim vldblDescuentos As Double                'Descuento de la cuenta, para la factura
    Dim vldblCantidad As Double                  'Cantidad de articulos para el detalle de la factura
    Dim vldblIEPS As Double                      'Cantidad IEPS
    'Datos fiscales
    Dim vlstrNombreFactura As String
    Dim vlstrDireccion As String
    Dim vlstrNumeroExterior As String
    Dim vlstrNumeroInterior As String
    Dim vlstrTelefono As String
    Dim vlstrRFC As String
    Dim vlstrColonia As String
    Dim vlstrCP As String
    Dim vlblnSujetoIEPS As Boolean
    Dim vlstrregimenfiscalreceptor As String

    Dim vPrinter As Printer
    Dim vllngCorteGrabando As Long
    Dim vllngClaveVentaTicket As Long
    Dim vllngSeleccionValida As Long 'Indica si la seleccion que se hizo para facturar es válida
    Dim vldblCantidadCredito As Double 'Acumulado de lo que se vendió a crédito
    Dim vldblIVACredito As Double 'Acumulado del IVA de lo que se vendió a crédito
    Dim vllngNumeroCliente As Long  'Número de cliente cuando la venta fue registrada a credito
    Dim lngNumMovimiento As Long 'Consecutivo del movimiento de crédito creado
    Dim intTotalTickets As Integer
    Dim lngCuentaTicket As Long
    Dim lngCveCiudad As Long 'Clave de la cd, de facturacion
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    Dim blnExtranjero As Boolean
    Dim intTipoAgrupacion As Integer
    Dim rsTipoAgrupacion As New ADODB.Recordset
    Dim vldblIvaDescuento As Double
    Dim vldblIVACiclo As Double
    Dim vldblIVACicloConDecimales As Double
    Dim vldblIEPSCiclo As Double
    Dim vldblcveVenta As Double
    Dim vldblsumivaNoRed As Double
    Dim vldblsumivaRed As Double

    Dim vlblnAcutalizaDatosFiscales As Boolean
    Dim vlstrsql As String
    
    'calculo de las cantidades que gravan y las que no graban, se agrego para el manejo de IEPS
    Dim rsPvFacturaImporte As New ADODB.Recordset
    Dim rsCalculosPvFacturaImporte As New ADODB.Recordset
    Dim vldblImporteGravado As Double
    Dim vldblSumatoria As Double
    Dim vllngCorteUsado As Long
    Dim vlintFolio As Long
    Dim rsChequeTransCta As New ADODB.Recordset
    Dim vldblSubtotalCiclo As Double
    Dim vldblDescuentosCiclo As Double
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim vlintIVAConceptofacturacion As Integer
    Dim vlstrFoliosTicketsenFactura As String
    Dim rsPuntosTickets As New ADODB.Recordset
    Dim rsPuntosObtenidos As New ADODB.Recordset
    Dim rsPuntosUtilizados As New ADODB.Recordset

On Error GoTo NotificaError
    '-----------------------------------------------------------------
    '*****************************************************************************************************************************************************************
    '------------------------------------------------------------------
    '- ¿Al menos uno seleccionado? -
    '------------------------------------------------------------------
    vlblnbandera = False
    For vlintContador = 1 To grdBuscaTickets.Rows - 1
        If grdBuscaTickets.TextMatrix(vlintContador, 0) = "*" Then
            vlblnbandera = True
            Exit For
        End If
    Next
    
    If Not vlblnbandera Then Exit Sub
    
    vllngSeleccionValida = flngSeleccionValida()
    If vllngSeleccionValida <> 0 Then
        MsgBox SIHOMsg(Str(vllngSeleccionValida)), vbInformation + vbOKOnly, "Mensaje"
        Exit Sub
    End If
    
    '--------------------------------------------------------
    ' Validación de que tenga una impresora seleccionada
    '--------------------------------------------------------
    vlstrSentencia = "select chrNombreImpresora Impresora from ImpresoraDepartamento where chrTipo = 'FA' and smiCveDepartamento = " & Trim(Str(vgintNumeroDepartamento))
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsTemp.RecordCount > 0 Then
        For Each vPrinter In Printers
            If UCase(Trim(vPrinter.DeviceName)) = UCase(Trim(rsTemp!Impresora)) Then
                 Set Printer = vPrinter
            End If
        Next
    Else
        MsgBox SIHOMsg(492), vbCritical, "Mensaje"
        Exit Sub
    End If
    rsTemp.Close
    
    '--------------------------------------------------------
    ' Tipo de cambio de los DOLARES
    '--------------------------------------------------------
    vldblTipoCambio = fdblTipoCambio(fdtmServerFecha, "V") 'Tipo de cambio a la Venta
    If vldblTipoCambio = 0 Then
        MsgBox SIHOMsg(231), vbCritical, "Mensaje"
        Exit Sub
    End If
    '--------------------------------------------------------
    
    If Not fblnValidaCuentaPuenteBanco(vgintClaveEmpresaContable) Then Exit Sub
    
    
    
    '--------------------------------------------------------
    ' Identifica el tipo de formato a utilizar
    '--------------------------------------------------------
        
      lngCveFormato = 1
      vlngCveFormato = 0
      frsEjecuta_SP vgintNumeroDepartamento & "|" & vgintClaveEmpresaContable & "|" & vgintTipoPaciente & "|E", "fn_PVSelFormatoFactura2", True, lngCveFormato
      vllngFormatoaUsar = lngCveFormato

      'Se valida en caso de no haber formato activo mostrar mensaje y cancelar transacción
      If vllngFormatoaUsar = 0 Then
          MsgBox SIHOMsg(373), vbCritical, "Mensaje"  'No se encontró un formato válido de factura.
          Exit Sub
      Else
          pValidaFormato
      End If
    
    If vlngCveFormato = 2 Or vlngCveFormato = 3 Then
        If Not fblnValidaSAT Then Exit Sub
    Else
        If vlngCveFormato = 1 Then
            If Not fblnValidaSATotrosConceptosCONCEPTOFACTURACION Then Exit Sub
        End If
    End If
    


    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
       
    If vllngPersonaGraba = 0 Then Exit Sub

    '--------------------------------------------------------
    ' Pedir Datos Fiscales
    '--------------------------------------------------------
    If Not pRegresaDatosFiscales Then 'si no regresa datos fiscales
       'se activa la pantalla de datos fiscales pero en modobusqueda
       frmDatosFiscales.sstDatos.Tab = 1
       vlblnAcutalizaDatosFiscales = False 'no se actualizan datos fiscales
    Else
       vlblnAcutalizaDatosFiscales = True ' si se actualizan datos fiscales (falta ver el tipo de paciente)
    End If
    
    'se muestra la pantalla de datos fiscales
    frmDatosFiscales.vgblnMostrarUsoCFDI = True
    'frmDatosFiscales.vgstrTipoUsoCFDI = IIf(vgintEmpresa > 0, "EM", "TP")
    'frmDatosFiscales.vgintTipoPacEmp = IIf(vgintEmpresa > 0, vgintEmpresa, vgintTipoPaciente)
    frmDatosFiscales.Show vbModal
       
    'Inicializa la pantalla de datos fiscales en la busqueda / Otros
    With frmDatosFiscales
        vlstrNombreFactura = .vgstrNombre
        vlstrDireccion = .vgstrDireccion
        vlstrNumeroExterior = .vgstrNumExterior
        vlstrNumeroInterior = .vgstrNumInterior
        vlstrTelefono = .vgstrTelefono
        vlstrRFC = .vgstrRFC
        lngCveCiudad = .llngCveCiudad
        vlstrColonia = .vgstrColonia
        vlstrCP = .vgstrCP
        blnExtranjero = IIf(.vgBitExtranjero = 0, False, True)
        vlblnSujetoIEPS = IIf(.vgBitSujetoaIEPS = 0, False, True)
        vlintUsoCFDI = frmDatosFiscales.vgintUsoCFDI
        vlstrregimenfiscalreceptor = .vlstrRegimenFiscal
        
    End With
    Unload frmDatosFiscales
    Set frmDatosFiscales = Nothing
    If Trim(vlstrRFC) = "" Or Trim(vlstrNombreFactura) = "" Then Exit Sub
    
    If Trim(vlstrRFC) = "XAXX010101000" And Trim(vlstrNombreFactura) = "PUBLICO EN GENERAL" Then
        vlblnEsGlobal = True
        pRellenaInfoGlobal
    End If
   'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
    '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
    intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", llngFormatoFactura)

    If intTipoEmisionComprobante = 0 Then   'ERROR
        'Si es error, se cancela la transacción
        Exit Sub
    End If

    'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
    intTipoCFDFactura = fintTipoCFD("FA", llngFormatoFactura)
    
    'Si aparece un error terminar la transacción
    If intTipoCFDFactura = 3 Then   'ERROR
        'Si es error, se cancela la transacción
        Exit Sub
    End If

    '------------------------------------------------------------------
    ' Inicio de la TRANSACCION
    '------------------------------------------------------------------
    EntornoSIHO.ConeccionSIHO.BeginTrans

    '------------------------------------------------------------------
    '-Número de la factura-
    '------------------------------------------------------------------
    vllngFoliosFaltantes = 0
    
    pCargaArreglo vlaryParametros, vllngFoliosFaltantes & "|" & ADODB.adBSTR & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
    frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "sp_gnFolios", , , vlaryParametros
    pObtieneValores vlaryParametros, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
    '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
    strSerie = Trim(strSerie)
    vlstrFolioDocumento = strSerie & strFolio
    
    If Trim(vlstrFolioDocumento) = "0" Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        Exit Sub
    End If
    
    If Trim(vlstrFolioDocumento) <> "" Then
        '-------------------------------------
        ' Numero de corte
        '-------------------------------------
        vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
        pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
        
        'vllngCorteGrabando = 1
        'frsEjecuta_SP CStr(vllngNumeroCorte) & "|Grabando", "Sp_PvUpdEstatusCorte", True, vllngCorteGrabando
        'If vllngCorteGrabando = 2 Then               '(4)
            If grdBuscaTickets.RowData(1) > 0 Then 'Si no esta vacio
                '------------------------------------------------
                ' Var para el cálculo de los totales de la factura
                '------------------------------------------------
                vldblDescuentos = 0
                vldblIVA = 0
                vldblSubtotal = 0
                vldblIEPS = 0
                '------------------------------------------------
                '-Inicializo y configuro Grid Temporal(escondido) de cargos
                pInicioGridCargos
                pConfiguraGridCargos
                '-----------------------------------------------------------
                intTotalTickets = 0
                vldblCantidadCredito = 0
                vllngNumeroCliente = 0
                vldblIVACredito = 0
                
                ReDim aFormasPago(0) '- Preparar arreglo de las formas de pago -'
                vlintConta2 = 0
                ReDim aMovimientoBancoForma(0) '- Preparar arreglo de movimientos a bancos -'
                vlintContadorMovs = 0
                vlstrFoliosTicketsenFactura = ""
                
                For vlintContador = 1 To grdBuscaTickets.Rows - 1  '--Ciclo principal--                    '(2)
                    If grdBuscaTickets.TextMatrix(vlintContador, 0) = "*" Then
                        intTotalTickets = intTotalTickets + 1
                        
                        lngCuentaTicket = CLng(Val(grdBuscaTickets.TextMatrix(vlintContador, 18)))
                    
                        vllngNumeroCliente = IIf(vllngNumeroCliente > 0, vllngNumeroCliente, Val(grdBuscaTickets.TextMatrix(vlintContador, 16)))
                        '-------------------------------------------------------------------------------------------------'
                        ' Tomo los cargos del detalle de VentaPublico y los pongo en grdArticulos para generar la factura '
                        '-------------------------------------------------------------------------------------------------'
                        vlstrSentencia = "SELECT " & _
                                        " pvdetalleventapublico.INTCVEVENTA CveVenta, pvDetalleVentaPublico.intNumCargo NumCargo, " & _
                                        " pvDetalleVentaPublico.MNYPRECIO, pvDetalleVentaPublico.INTCANTIDAD, pvDetalleVentaPublico.MNYIVA, " & _
                                        " pvDetalleVentaPublico.mnyDescuento, pvDetalleVentaPublico.smiCveConceptoFacturacion ConceptoFacturacion, " & _
                                        " PvVentaPublico.INTCVEDEPARTAMENTO, PvVentaPublico.INTNUMCORTE, pvDetalleventapublico.mnyIEPS IEPS, pvdetalleventapublico.numporcentajeIEPS TASAIEPS" & _
                                        " FROM pvDetalleVentaPublico " & _
                                        " LEFT OUTER JOIN ivArticulo ON IvArticulo.intIdArticulo = pvDetalleVentaPublico.intCveCargo " & _
                                        " LEFT OUTER JOIN PvOtroConcepto ON PvOtroConcepto.intCveConcepto = pvDetalleVentaPublico.intCveCargo " & _
                                        " INNER JOIN PvVentaPublico ON PvDetalleVentaPublico.intCveVenta = PvVentaPublico.intCveVenta " & _
                                        " WHERE PvDetalleVentaPublico.intCveVenta = " & Trim(grdBuscaTickets.RowData(vlintContador))
                        Set rsDetalleVentaPublico = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
                        vllngClaveVentaTicket = CLng(grdBuscaTickets.RowData(vlintContador))
                        vldblIVACiclo = 0
                        vldblIVACicloConDecimales = 0
                        vldblIEPSCiclo = 0
                        vldblSubtotalCiclo = 0
                        vldblDescuentosCiclo = 0
                        Do While Not rsDetalleVentaPublico.EOF
                            With grdArticulos
                                If .RowData(1) <> -1 Then
                                    .Rows = .Rows + 1
                                End If
                                .Row = .Rows - 1
                                .RowData(.Row) = rsDetalleVentaPublico!NumCargo
                                .TextMatrix(.Row, 9) = rsDetalleVentaPublico!MNYPRECIO * rsDetalleVentaPublico!intCantidad 'importe
                                .TextMatrix(.Row, 2) = rsDetalleVentaPublico!MNYPRECIO
                                .TextMatrix(.Row, 3) = rsDetalleVentaPublico!intCantidad
                                .TextMatrix(.Row, 6) = rsDetalleVentaPublico!IEPS 'IEPS@
                                .TextMatrix(.Row, 11) = rsDetalleVentaPublico!MNYIVA 'IVA
                                .TextMatrix(.Row, 5) = rsDetalleVentaPublico!MNYDESCUENTO 'Descuento
                                .TextMatrix(.Row, 13) = rsDetalleVentaPublico!ConceptoFacturacion 'ConceptoFacturacion
                                .TextMatrix(.Row, 16) = rsDetalleVentaPublico!CveVenta 'Cve venta
                                .TextMatrix(.Row, 17) = rsDetalleVentaPublico!TASAIEPS / 100 'tasa de IEPS que aplica
                                vldblDescuentos = vldblDescuentos + rsDetalleVentaPublico!MNYDESCUENTO 'suma descuentos
                                vldblIVACiclo = vldblIVACiclo + Round(rsDetalleVentaPublico!MNYIVA, 2) 'Suma IVa
                                vldblIVACicloConDecimales = vldblIVACicloConDecimales + rsDetalleVentaPublico!MNYIVA 'Suma IVa
                                vldblSubtotal = vldblSubtotal + (rsDetalleVentaPublico!MNYPRECIO * rsDetalleVentaPublico!intCantidad) 'suma importes
                                vldblIEPS = vldblIEPS + Val(rsDetalleVentaPublico!IEPS)
                                vldblIEPSCiclo = vldblIEPSCiclo + Val(rsDetalleVentaPublico!IEPS)
                                vldblSubtotalCiclo = vldblSubtotalCiclo + Round((rsDetalleVentaPublico!MNYPRECIO * rsDetalleVentaPublico!intCantidad), 2) 'suma importes
                                vldblDescuentosCiclo = vldblDescuentosCiclo + rsDetalleVentaPublico!MNYDESCUENTO 'suma descuentos
                            End With
                            rsDetalleVentaPublico.MoveNext
                        Loop
                        'vldblIVA = vldblIVA + Format(Val(vldblIVACiclo), "Fixed")
                        vldblIVA = vldblIVA + vldblIVACicloConDecimales
                        rsDetalleVentaPublico.Close
                                        
                        '--------------------------------------------------------------'
                        ' Registrar los movimientos contables del ticket en la factura '
                        '--------------------------------------------------------------'
                        vgstrParametrosSP = Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)) & "|" & "TI" & "|" & grdBuscaTickets.TextMatrix(vlintContador, 17)
                        Set rsPolizaTicket = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelPolizaDocto")
                        Do While Not rsPolizaTicket.EOF
'                            vgstrParametrosSP = CStr(vllngNumeroCorte) _
'                                                & "|" & vlstrFolioDocumento _
'                                                & "|" & "FA" _
'                                                & "|" & CStr(rsPolizaTicket!INTNUMCUENTA) _
'                                                & "|" & CStr(rsPolizaTicket!mnyCantidad) _
'                                                & "|" & CStr(rsPolizaTicket!bitcargo)
'                            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                            
                            'Agregado por caso 10442, verifica si la cuenta se encuentra relacionada con un banco
                            'para registrar el movimiento contable a la cuenta puente para no afectar la cuenta de banco
                            vlintFolio = 0
                            If fblnCuentaRelacionadaConBancos(rsPolizaTicket!intNumCuenta) Then
                                vlintFolio = 1
                                frsEjecuta_SP Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)) & "|" & rsPolizaTicket!intNumCuenta & "|" & rsPolizaTicket!MNYCantidad & "|" & "TI", "fn_pvCuentaRelacionadaConBanco", True, vlintFolio
                            End If
                                                                    
                            pAgregarMovArregloCorte CStr(vllngNumeroCorte), vllngPersonaGraba, vlstrFolioDocumento, "FA", IIf(vlintFolio = 1, lngCuentaPuenteBanco, rsPolizaTicket!intNumCuenta), _
                            rsPolizaTicket!MNYCantidad, IIf(rsPolizaTicket!bitcargo = 1, True, False), "", 0, 0, "", 0, 2, "", ""
                            rsPolizaTicket.MoveNext
                        Loop
                        
                        '----------------------'
                        ' Se cancela el ticket '
                        '----------------------'
                        vlstrSentencia = "update PvVentaPublico set bitCancelado = 1, chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "' where intCveVenta = " & Trim(grdBuscaTickets.RowData(vlintContador))
                        pEjecutaSentencia (vlstrSentencia)
                        
                        vlstrSentencia = "Select * from pvDocumentoCancelado where chrFoliodocumento = '-1'"
                        Set rsPvDocumentoCancelado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                        'Se registra en documentos cancelados
                            With rsPvDocumentoCancelado
                                .AddNew
                                !chrFolioDocumento = Trim(grdBuscaTickets.RowData(vlintContador)) ' Trim(grdBuscaTickets.TextMatrix(vlintContador, 1))
                                !chrTipoDocumento = "TI"
                                !SMIDEPARTAMENTO = vgintNumeroDepartamento
                                !intEmpleado = vllngPersonaGraba
                                !dtmfecha = fdtmServerFecha
                                .Update
                            End With
                        '-------------------------------------------------------'
                        ' Registrar las formas de pago del ticket en la factura '
                        '-------------------------------------------------------'
                        vgstrParametrosSP = Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)) & "|" & "TI" & "|" & grdBuscaTickets.TextMatrix(vlintContador, 17)
                        Set rsFormasPagos = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaDoctoCorte")
                        Do While Not rsFormasPagos.EOF
                            'Sumar cantidades que son a credito
                            vgstrParametrosSP = CStr(rsFormasPagos!intFormaPago) & "|" & -1 & "|" & -1 & "|" & -1 & "|" & -1 & "|'*'"
                            Set rstipoformapago = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaPago")
                            If rstipoformapago!chrTipo = "C" Then
                                'vldblCantidadCredito = vldblCantidadCredito + rsFormasPagos!mnyCantidadPagada
                                vldblCantidadCredito = vldblCantidadCredito + ((vldblSubtotalCiclo - vldblDescuentosCiclo) + vldblIVACiclo + vldblIEPSCiclo)
                                vldblIVACredito = vldblIVACredito + vldblIVACiclo
                            End If
                          
'                            vgstrParametrosSP = CStr(vllngNumeroCorte) _
'                                                & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora) _
'                                                & "|" & vlstrFolioDocumento _
'                                                & "|" & "FA" _
'                                                & "|" & CStr(rsFormasPagos!intFormaPago) _
'                                                & "|" & CStr(rsFormasPagos!MNYCANTIDADPAGADA) _
'                                                & "|" & CStr(rsFormasPagos!mnytipocambio) _
'                                                & "|" & CStr(rsFormasPagos!intfoliocheque) _
'                                                & "|" & CStr(vllngNumeroCorte)
'                            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"

                            Set rsChequeTransCta = frsRegresaRs("SELECT PVCQTR.* FROM PVCORTECHEQUETRANSCTA PVCQTR INNER JOIN PVDETALLECORTE PVDC ON PVDC.INTCONSECUTIVO = PVCQTR.INTCONSECUTIVODETCORTE WHERE TRIM(PVDC.CHRFOLIODOCUMENTO) = '" & Trim(rsFormasPagos!chrFolioDocumento) & "' AND TRIM(PVDC.CHRTIPODOCUMENTO) = '" & Trim(rsFormasPagos!chrTipoDocumento) & "' AND PVDC.INTFORMAPAGO = " & rsFormasPagos!intFormaPago & " AND PVDC.MNYCANTIDADPAGADA = " & rsFormasPagos!mnyCantidadPagada & " AND PVDC.MNYTIPOCAMBIO = " & rsFormasPagos!mnytipocambio & " AND PVDC.INTFOLIOCHEQUE = " & rsFormasPagos!intfoliocheque & " ORDER BY PVCQTR.INTCONSECUTIVODETCORTE", adLockReadOnly, adOpenForwardOnly)
                            If rsChequeTransCta.RecordCount > 0 Then
                                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, vlstrFolioDocumento, "FA", 0, rsFormasPagos!mnyCantidadPagada, False, _
                                fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora, True), rsFormasPagos!intFormaPago, rsFormasPagos!mnytipocambio, _
                                CStr(rsFormasPagos!intfoliocheque), vllngNumeroCorte, 1, "", "", False, Trim(Replace(Replace(Replace(vlstrRFC, "-", ""), "_", ""), " ", "")), IIf(IsNull(rsChequeTransCta!CHRCLAVEBANCOSAT), "", rsChequeTransCta!CHRCLAVEBANCOSAT), IIf(IsNull(rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), "", rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), IIf(IsNull(rsChequeTransCta!VCHCUENTABANCARIA), "", rsChequeTransCta!VCHCUENTABANCARIA), IIf(IsNull(rsChequeTransCta!dtmfecha), fdtmServerFecha, rsChequeTransCta!dtmfecha)
                            Else
                                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, vlstrFolioDocumento, "FA", 0, rsFormasPagos!mnyCantidadPagada, False, _
                                fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora, True), rsFormasPagos!intFormaPago, rsFormasPagos!mnytipocambio, _
                                CStr(rsFormasPagos!intfoliocheque), vllngNumeroCorte, 1, "", ""
                            End If
                           
                            '- Llenar arreglo de las formas de pago para guardarlas en la tabla intermedia -'
                            ReDim Preserve aFormasPago(vlintConta2)
                            With aFormasPago(vlintConta2)
                                .vlbolEsCredito = rstipoformapago!chrTipo = "C"
                                .vldblCantidad = rsFormasPagos!mnyCantidadPagada
                                .vldblTipoCambio = rsFormasPagos!mnytipocambio
                                .vlintNumFormaPago = rsFormasPagos!intFormaPago
                                .vllngCuentaContable = rstipoformapago!INTCUENTACONTABLE
                                .vlstrFolio = vlstrFolioDocumento
                                .lngIdBanco = flngCuentaBanco(rstipoformapago!INTCUENTACONTABLE)
                            End With
                            vlintConta2 = vlintConta2 + 1
                                                   
                            rsFormasPagos.MoveNext
                        Loop
                        If rstipoformapago.State <> adStateClosed Then rstipoformapago.Close
                        
                        '---------------------------------------------------------'
                        ' Cancelar el ticket en el corte y su movimiento contable '
                        '---------------------------------------------------------'
'                        vgstrParametrosSP = Trim(grdBuscaTickets.TextMatrix(vlintcontador, 1)) & "|" & "TI" & "|" & grdBuscaTickets.TextMatrix(vlintcontador, 17) & "|" & CStr(vllngNumeroCorte)
'                        frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdCancelaDoctoCorte"
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)), "TI", 0, 0, _
                        False, "", 0, 0, "", CLng(grdBuscaTickets.TextMatrix(vlintContador, 17)), 3, "", ""
                    
                        'Debido a que los tickets son reemplazados por una factura, se cancelan los registros que se hayan hecho para crédito
                        vlstrSentencia = "update CcMovimientoCredito set bitCancelado = 1, dtmFechaCancelacion = " & fstrFechaSQL(fdtmServerFecha) & " where chrTipoReferencia = 'TI' and chrFolioReferencia = '" & Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)) & "'"
                        pEjecutaSentencia vlstrSentencia
                        
                        '---------------------------------------------------------'
                        ' Cancelar el movimiento de las formas de pago del ticket '
                        '---------------------------------------------------------'
                        If intBitCuentaPuenteBanco = 0 Then
                            pCancelaMovimiento Trim(grdBuscaTickets.RowData(vlintContador)), Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)), grdBuscaTickets.TextMatrix(vlintContador, 17), vllngNumeroCorte, vllngPersonaGraba, True
                        End If
                        
                        vlstrFoliosTicketsenFactura = vlstrFoliosTicketsenFactura & ", '" & Trim(grdBuscaTickets.TextMatrix(vlintContador, 1)) & "'"
                        
                    End If
                Next vlintContador
            End If
            
            '------------------------------------'
            ' Guardar en la factura y su detalle '
            '------------------------------------'
            vlstrSentencia = "select * from PVFactura where chrFolioFactura = '-1'"
            Set rsFactura = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            With rsFactura
                .AddNew
                !chrfoliofactura = vlstrFolioDocumento
                !dtmFechahora = fdtmServerFecha + fdtmServerHora
                !chrRFC = IIf(blnExtranjero, "XEXX010101000", IIf(Len(Trim(Replace(Replace(Replace(vlstrRFC, "-", ""), "_", ""), " ", ""))) < 12, "XAXX010101000", Trim(Replace(Replace(Replace(vlstrRFC, "-", ""), "_", ""), " ", ""))))
                !CHRNOMBRE = vlstrNombreFactura
                !CHRCALLE = vlstrDireccion
                !VCHNUMEROEXTERIOR = vlstrNumeroExterior
                !VCHNUMEROINTERIOR = vlstrNumeroInterior
                !CHRTELEFONO = vlstrTelefono
                !VCHCOLONIA = Trim(vlstrColonia)
                !VCHCODIGOPOSTAL = Trim(vlstrCP)
                
                'anexo de el regimen fiscal al guardar la factura
                !VCHREGIMENFISCALRECEPTOR = vlstrregimenfiscalreceptor
                
                '!smyIVA = Format(Val(vldblIVA), "Fixed")
                !smyIVA = Round(vldblIVA, 2)
                !MNYDESCUENTO = Round(vldblDescuentos, 2)
                !chrEstatus = " "
                !INTMOVPACIENTE = IIf(intTotalTickets = 1 And lngCuentaTicket <> 0, lngCuentaTicket, 0)
                !CHRTIPOPACIENTE = IIf(intTotalTickets = 1 And lngCuentaTicket <> 0, "E", "V")
                !SMIDEPARTAMENTO = vgintNumeroDepartamento
                !intCveEmpleado = vllngPersonaGraba
                !intnumcorte = vllngNumeroCorte
                !mnyAnticipo = 0
                !mnyTotalFactura = Round((Round(vldblSubtotal, 2) - vldblDescuentos + vldblIEPS + Round(vldblIVA, 2)), 2) '<----se agrega el ieps al total de la factura
                !BITPESOS = 1
                !mnytipocambio = 0
                !chrTipoFactura = IIf(llngEmpresaTipoPaciente = 0, "V", "E")
                !intCveVentaPublico = -1
                !INTCVECIUDAD = lngCveCiudad
                !intcveempresa = llngEmpresaTipoPaciente
                !mnyTotalPagar = Round((Round(vldblSubtotal, 2) - vldblDescuentos + vldblIEPS + Round(vldblIVA, 2)), 2)
                !vchSerie = strSerie
                !INTFOLIO = strFolio
'                !bitdesgloseIEPS = IIf(vldblIEPS > 0, 1, IIf(vlblnSujetoIEPS, 1, 0))
                !bitdesgloseIEPS = IIf(vlblnSujetoIEPS, 1, 0)
                !intTipoDetalleFactura = intTipoAgrupacion
                !intCveUsoCFDI = vlintUsoCFDI
                !bitFacturaGlobal = IIf(vlblnEsGlobal, 1, 0)
                .Update
            End With
            '--------------------------------------
            ' Detalle de la Factura
            '--------------------------------------
            ' Aquí trabajo con dos grid escondidos(grdFactura y grdArticulos),
            ' para que este mas fácil organizar los cargos,
            ' como en la facturación normal. (si ya se, copie el código)
            '--------------------------------------
            ' Preparo el Grid escondido para guardar los datos de la factura
            With grdFactura
                .Clear
                .ClearStructure
                .Cols = 8
                .Rows = 2
                .RowData(1) = -1
                .ColWidth(1) = 0 'Concepto de facturación (texto)
                .ColWidth(2) = 0 'Cargo
                .ColWidth(3) = 0 'Abono
                .ColWidth(4) = 0 'IVA
                .ColWidth(5) = 0 'Descuentos
                .ColWidth(6) = 0 'IEPS
                .ColWidth(7) = 0 'Tasa IEPS

            End With

            '---------------------------------------------------------------
            For vlintContador = 1 To grdArticulos.Rows - 1
                vlintPosicion = 0
                For vlintConta2 = 1 To grdFactura.Rows - 1  'For para ver si ya existe en el Grid temporal
                    If CLng(grdArticulos.TextMatrix(vlintContador, 13)) = grdFactura.RowData(vlintConta2) Then
                        vlintPosicion = vlintConta2
                        Exit For
                    End If
                Next vlintConta2
                If vlintPosicion <> 0 Then
                    ' La 2 es el CARGO
                    grdFactura.TextMatrix(vlintPosicion, 2) = Val(grdFactura.TextMatrix(vlintPosicion, 2)) + _
                    Round((Val(grdArticulos.TextMatrix(vlintContador, 9)) + IIf(vlblnSujetoIEPS = False, Val(grdArticulos.TextMatrix(vlintContador, 6)), 0)), 2)
                   
                    ' El 5 son los DESCUENTOS
                    grdFactura.TextMatrix(vlintPosicion, 5) = Val(grdFactura.TextMatrix(vlintPosicion, 5)) + _
                        (Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "############.##")))
                    
                    'El 6 es el IEPS
                    If vlblnSujetoIEPS Then
                        grdFactura.TextMatrix(vlintPosicion, 6) = Val(grdFactura.TextMatrix(vlintPosicion, 6)) + Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
                    End If
                    
                    'El 4 es el IVA
                    grdFactura.TextMatrix(vlintPosicion, 4) = Val(grdFactura.TextMatrix(vlintPosicion, 4)) + Val(Format(grdArticulos.TextMatrix(vlintContador, 11), ""))
                Else
                    If grdFactura.RowData(1) <> -1 Then
                        grdFactura.Rows = grdFactura.Rows + 1
                    End If
                    grdFactura.RowData(grdFactura.Rows - 1) = CLng(grdArticulos.TextMatrix(vlintContador, 13)) 'Clave del Concepto
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 1) = ""  'Descripción del Concepto
    
                    vldblCantidad = Round((Val(Format(Val(grdArticulos.TextMatrix(vlintContador, 9)) + IIf(vlblnSujetoIEPS = False, Val(grdArticulos.TextMatrix(vlintContador, 6)), 0), "Fixed"))), 3)
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 2) = vldblCantidad 'Cantidad
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 5) = Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "###########.##")) 'Descuentos individuales
                    
                    'Monto IEPS y Tasa, como está agrupado por concepto de facturación, no se deben mezclar articulos con diferente IEPS bajo el mismo concepto
                    If vlblnSujetoIEPS Then
                        grdFactura.TextMatrix(grdFactura.Rows - 1, 6) = Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
                        grdFactura.TextMatrix(grdFactura.Rows - 1, 7) = grdArticulos.TextMatrix(vlintContador, 17)
                    Else
                        grdFactura.TextMatrix(grdFactura.Rows - 1, 6) = "0"
                    End If
                        
                    'IVA
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = Val(Format(grdArticulos.TextMatrix(vlintContador, 11), ""))
                    
                End If
            Next vlintContador
            'Para que se calcule de la misma forma que la poliza y evitar descuadre de centavos
            vldblcveVenta = 0
            vldblsumivaNoRed = 0
            vldblsumivaRed = 0
            
            'IVA----------------------------------------------------------------------------------------------------------------------------------------------------
'            For vlintContador = 1 To grdArticulos.Rows - 1
'                If vldblcveVenta = 0 Or vldblcveVenta = grdArticulos.TextMatrix(vlintContador, 16) Then
'                    vldblsumivaNoRed = vldblsumivaNoRed + grdArticulos.TextMatrix(vlintContador, 11)
'                Else
'                    vldblsumivaRed = vldblsumivaRed + Format(Val(vldblsumivaNoRed), "Fixed")
'                    vldblsumivaNoRed = 0
'                    vldblsumivaNoRed = vldblsumivaNoRed + grdArticulos.TextMatrix(vlintContador, 11)
'                End If
'                vldblcveVenta = grdArticulos.TextMatrix(vlintContador, 16)
'                If vlintContador = grdArticulos.Rows - 1 Then
'                    vldblsumivaRed = vldblsumivaRed + Format(Val(vldblsumivaNoRed), "Fixed")
'                End If
'            Next vlintContador
'            grdFactura.TextMatrix(IIf(vlintPosicion <> 0, vlintPosicion, grdFactura.Rows - 1), 4) = vldblsumivaRed
            '-------------------------------------------------------------------------------------------------------------------------------------------------------
            vllngConsecutivoFactura = flngObtieneIdentity("SEC_PVFACTURA", rsFactura!intConsecutivo)
        
            vlstrSentencia = "select * from PVDetalleFactura where chrFolioFactura = '-1'"
            Set rsDetalleFactura = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            vldblIvaDescuento = 0
            
            With rsDetalleFactura
                For a = 1 To grdFactura.Rows - 1
                    If grdFactura.RowData(a) > 0 Then 'Porque pueden ser negativos con el descuento y Pagos
                        .AddNew
                        !chrfoliofactura = vlstrFolioDocumento
                        !smicveconcepto = grdFactura.RowData(a)
                        !MNYCantidad = Val(Format(grdFactura.TextMatrix(a, 2), "############.##"))
                        !MNYIVA = Format(Val(grdFactura.TextMatrix(a, 4)), "")
                        !MNYDESCUENTO = Val(grdFactura.TextMatrix(a, 5))
                        
'                        If !mnyCantidad - !MNYDESCUENTO = 0 Then
'                            strSQL = "select smyIVA from PVConceptoFacturacion where smiCveConcepto = " & IIf(grdFactura.RowData(a) < 0, grdFactura.RowData(a) * -1, grdFactura.RowData(a))
'                            Set rsTmp = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
'                            If Not rsTmp.EOF Then
'                                vlintIVAConceptofacturacion = IIf(rsTmp.Fields(0).Value = 0, 0, rsTmp.Fields(0).Value)
'                            Else
'                                vlintIVAConceptofacturacion = 0
'                            End If
'                            rsTmp.Close
'
'                            !mnyIVAConcepto = Format(!mnyCantidad * (vlintIVAConceptofacturacion / 100), "")
'                        Else
'                            !mnyIVAConcepto = Format(!mnyCantidad * (!MNYIVA / (!mnyCantidad - !MNYDESCUENTO)), "")
'                        End If
                        
                        !mnyIVAConcepto = Format(!MNYCantidad * (!MNYIVA / IIf((!MNYCantidad - !MNYDESCUENTO) > 0, (!MNYCantidad - !MNYDESCUENTO), 1)), "Fixed")
                        
                        If !MNYIVA <> 0 Then
                            vldblIvaDescuento = vldblIvaDescuento + (!MNYDESCUENTO * (vgdblCantidadIvaGeneral / 100))
                        End If
                        
                        !mnyIeps = Val(grdFactura.TextMatrix(a, 6))
                        If !mnyIeps <> 0 Then
                            !numTasaIEPS = Val(grdFactura.TextMatrix(a, 7))
                        End If

                        
                        !chrTipo = "NO"
                        .Update
                        If Val(grdFactura.TextMatrix(a, 5)) > 0 Then
                            .AddNew
                            !chrfoliofactura = vlstrFolioDocumento
                            !smicveconcepto = -2
                            !MNYCantidad = Val(grdFactura.TextMatrix(a, 5))
                            !MNYIVA = 0
                            !MNYDESCUENTO = 0
                            !chrTipo = "DE"
                            !mnyIVAConcepto = vldblIvaDescuento
                            .Update
                            
                            vldblIvaDescuento = 0
                        End If
                    End If
                Next a
            End With
            rsDetalleFactura.Close
            
            '--------------------------------------------------------------------------------------------------------------------------------------------------------
            ' CANTIDADES Y TASAS IEPS
            '--------------------------------------------------------------------------------------------------------------------------------------------------------
            If vlblnLicenciaIEPS Then pGrabaTasasIEPS vllngConsecutivoFactura
            
            vlstrSentencia = "SELECT * FROM PVFACTURAIMPORTE WHERE INTCONSECUTIVO = " & vllngConsecutivoFactura
            Set rsPvFacturaImporte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rsPvFacturaImporte.RecordCount = 0 Then frsEjecuta_SP Str(vllngConsecutivoFactura) & "|0|0|0|0", "SP_PVINSFACTURAIMPORTE"
               
            vldblImporteGravado = 0
            vldblSumatoria = 0
            vlstrSentencia = ""
            If rsFactura!smyIVA > 0 Then
               vlstrSentencia = "SELECT nvl(ROUND(NVL(SUM((PVDETALLEVENTAPUBLICO.MnyPrecio * PVDETALLEVENTAPUBLICO.IntCantidad) + PVDETALLEVENTAPUBLICO.MNYIEPS - PVDETALLEVENTAPUBLICO.MnyDescuento),0),2),0) AS ImporteGravado FROM PVVENTAPUBLICO INNER JOIN PVDETALLEVENTAPUBLICO ON PVDETALLEVENTAPUBLICO.IntCveVenta = PVVENTAPUBLICO.IntCveVenta WHERE PVDETALLEVENTAPUBLICO.MnyIva > 0 AND PVVENTAPUBLICO.ChrFolioFactura = '" & rsFactura!chrfoliofactura & "'"
            End If
             
            If Trim(vlstrSentencia) <> "" Then
               Set rsCalculosPvFacturaImporte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
               If rsCalculosPvFacturaImporte.RecordCount > 0 Then
                  vldblImporteGravado = rsCalculosPvFacturaImporte!ImporteGravado
               End If
            End If
            
            pEjecutaSentencia "UPDATE PVFACTURAIMPORTE SET MNYSUBTOTALGRAVADO = " & vldblImporteGravado & _
                                                       ", MNYSUBTOTALNOGRAVADO = " & rsFactura!mnyTotalFactura - rsFactura!smyIVA - vldblImporteGravado & _
                                                       " WHERE IntConsecutivo = " & vllngConsecutivoFactura
            rsFactura.Close
            
            '----------------- CASO 7442 : Guardar información de la forma de pago en tabla intermedia -----------------'
            ' (Se requiere el consecutivo de la factura, por eso se guarda después de hacer el movimiento en PvFactura) '
'            For vlintContador = 0 To UBound(aFormasPago(), 1)
'                If Not aFormasPago(vlintContador).vlbolEsCredito Then 'Formas de pago distintas a Crédito
'                    vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora) & "|" & aFormasPago(vlintContador).vlintNumFormaPago & "|" & aFormasPago(vlintContador).lngIdBanco & "|" & _
'                                        aFormasPago(vlintContador).vldblCantidad & "|" & IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContador).vldblTipoCambio & "|" & fstrTipoMovimientoForma(aFormasPago(vlintContador).vlintNumFormaPago) & "|" & _
'                                        "FA" & "|" & vllngConsecutivoFactura & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerFechaHora) & "|" & "1" & "|" & cgstrModulo
'                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
'                End If
'            Next vlintContador

            For vlintContador = 1 To UBound(aMovimientoBancoForma)
                vlstrSentencia = "update PvMovimientoBancoForma set intNumDocumento = " & vllngConsecutivoFactura & _
                                 " where intNumDocumento = -999 and intIdMovimiento = " & aMovimientoBancoForma(vlintContador)
                pEjecutaSentencia (vlstrSentencia)
            Next vlintContador
            '-----------------------------------------------------------------------------------------------------------'
            
            '------------------------------------------'
            ' Poner el número de factura en los cargos '
            '------------------------------------------'
            For vlintConta2 = 1 To grdArticulos.Rows - 1
                vlstrSentencia = "update PVCargo set chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "'" & _
                                 " where intNumCargo = " & Trim(Str(grdArticulos.RowData(vlintConta2)))
                pEjecutaSentencia (vlstrSentencia)
            Next vlintConta2
        
            '-----------------------------------------------------------------------------------------------'
            ' Finalmente si alguna venta fue a crédito se registra el movimiento de crédito para la factura '
            '-----------------------------------------------------------------------------------------------'
            If vldblCantidadCredito > 0 Then
                vlstrSentencia = "select * from CcCliente where intNumCliente = " & Str(vllngNumeroCliente)
                Set rsDatosCliente = frsRegresaRs(vlstrSentencia)
                If rsDatosCliente.RecordCount <> 0 Then
                   'pCrearMovtoCredito vldblCantidadCredito, vllngNumeroCliente, rsDatosCliente!INTNUMCUENTACONTABLE, vlstrFolioDocumento, "FA", ((vldblSubtotal - vldblDescuentos) + vldblIVA + vldblIEPS), vldblIVA, Str(vllngPersonaGraba)
                   pCrearMovtoCredito Val(Format(vldblCantidadCredito, "############.##")), vllngNumeroCliente, rsDatosCliente!INTNUMCUENTACONTABLE, vlstrFolioDocumento, "FA", vldblCantidadCredito, vldblIVACredito, Str(vllngPersonaGraba)
                End If
            End If
        
            '-----------------------------------------------------------------------------'
            ' Guarda los puntos de lealtad de los tickets cancelados en la nueva factura -'
            '-----------------------------------------------------------------------------'
            ' -- 17115 --
            If blnLicenciaLealtadCliente Then
                '-- Suma de los puntos acumulados
                vlstrSentencia = "select sum(mnysubtotalfactura) subTotalFactura, sum(intpuntospaciente) puntosPaciente, max(intNumCuenta) numCuentaPaciente, " & _
                                        "sum(intpuntosmedicolab) puntosMedicoLab, max(intcvemedicolab) medicoLab, " & _
                                        "sum(intpuntosmedicoimagen) puntosMedicoImagen, max(intcvemedicoImagen) medicoImagen, " & _
                                        "sum(intpuntosmedicoCargo) puntosMedicoCargo, max(intcvemedicoCargo) medicoCargo " & _
                                 "from pvPuntosObtenidosPaciente where trim(chrfoliofactura) in (" & Mid(vlstrFoliosTicketsenFactura, 2) & ")"
                Set rsPuntosTickets = frsRegresaRs(vlstrSentencia)
                If rsPuntosTickets.RecordCount <> 0 Then
                    If Not IsNull(rsPuntosTickets!numCuentaPaciente) Then
                        pDatosPaciente (rsPuntosTickets!numCuentaPaciente)
                        vlstrSentencia = "SELECT * FROM PvPuntosObtenidosPaciente where intIdFactura = -1"
                        Set rsPuntosObtenidos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                           
                        rsPuntosObtenidos.AddNew
                        rsPuntosObtenidos!intIdFactura = vllngConsecutivoFactura
                        rsPuntosObtenidos!chrfoliofactura = Trim(vlstrFolioDocumento)
                        rsPuntosObtenidos!intNumCuenta = rsPuntosTickets!numCuentaPaciente
                        rsPuntosObtenidos!intNumPaciente = vllngNumeroPaciente
                        rsPuntosObtenidos!mnySubtotalFactura = rsPuntosTickets!subTotalFactura
                        rsPuntosObtenidos!intpuntospaciente = rsPuntosTickets!puntosPaciente
                        rsPuntosObtenidos!intcvemedicolab = rsPuntosTickets!medicoLab
                        rsPuntosObtenidos!intpuntosMedicoLab = rsPuntosTickets!puntosMedicoLab
                        rsPuntosObtenidos!intcvemedicoImagen = rsPuntosTickets!medicoImagen
                        rsPuntosObtenidos!intpuntosMedicoImagen = rsPuntosTickets!puntosMedicoImagen
                        rsPuntosObtenidos!intcvemedicoCargo = rsPuntosTickets!medicoCargo
                        rsPuntosObtenidos!intpuntosMedicoCargo = rsPuntosTickets!puntosMedicoCargo
                        rsPuntosObtenidos!dtmFechahora = fdtmServerFecha + fdtmServerHora
                        rsPuntosObtenidos.Update
                        rsPuntosObtenidos.Close
                    End If
                End If
                rsPuntosTickets.Close
                
                '-- Suma de los puntos utilizados
                vlstrSentencia = "select sum(mnyDescuentoAplicado) descuentoAplicado, sum(intpuntospaciente) puntosPaciente, max(intNumCuenta) numCuentaPaciente, " & _
                                        "sum(intpuntosmedicolab) puntosMedicoLab, max(intcvemedicolab) medicoLab, " & _
                                        "sum(intpuntosmedicoimagen) puntosMedicoImagen, max(intcvemedicoImagen) medicoImagen, " & _
                                        "sum(intpuntosmedicoCargo) puntosMedicoCargo, max(intcvemedicoCargo) medicoCargo " & _
                                 "from pvPuntosUtilizadosPaciente where trim(chrfoliofactura) in (" & Mid(vlstrFoliosTicketsenFactura, 2) & ")"
                Set rsPuntosTickets = frsRegresaRs(vlstrSentencia)
                If rsPuntosTickets.RecordCount <> 0 Then
                    If Not IsNull(rsPuntosTickets!numCuentaPaciente) Then
                        pDatosPaciente (rsPuntosTickets!numCuentaPaciente)
                        vlstrSentencia = "SELECT * FROM PvPuntosUtilizadosPaciente where intIdFactura = -1"
                        Set rsPuntosUtilizados = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                           
                        rsPuntosUtilizados.AddNew
                        rsPuntosUtilizados!intIdFactura = vllngConsecutivoFactura
                        rsPuntosUtilizados!chrfoliofactura = Trim(vlstrFolioDocumento)
                        rsPuntosUtilizados!intNumCuenta = rsPuntosTickets!numCuentaPaciente
                        rsPuntosUtilizados!intNumPaciente = vllngNumeroPaciente
                        rsPuntosUtilizados!mnyDescuentoAplicado = rsPuntosTickets!descuentoAplicado
                        rsPuntosUtilizados!intpuntospaciente = rsPuntosTickets!puntosPaciente
                        rsPuntosUtilizados!intcvemedicolab = rsPuntosTickets!medicoLab
                        rsPuntosUtilizados!intpuntosMedicoLab = rsPuntosTickets!puntosMedicoLab
                        rsPuntosUtilizados!intcvemedicoImagen = rsPuntosTickets!medicoImagen
                        rsPuntosUtilizados!intpuntosMedicoImagen = rsPuntosTickets!puntosMedicoImagen
                        rsPuntosUtilizados!intcvemedicoCargo = rsPuntosTickets!medicoCargo
                        rsPuntosUtilizados!intpuntosMedicoCargo = rsPuntosTickets!puntosMedicoCargo
                        rsPuntosUtilizados!dtmFechahora = fdtmServerFecha + fdtmServerHora
                        rsPuntosUtilizados.Update
                        rsPuntosUtilizados.Close
                    End If
                End If
                rsPuntosTickets.Close
            End If
                        
                        
        vlstrSentencia = "Select intTipoAgrupaDigital From Formato Where Formato.INTNUMEROFORMATO = " & llngFormatoFactura
        Set rsTipoAgrupacion = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        intTipoAgrupacion = IIf(IsNull(rsTipoAgrupacion!intTipoAgrupaDigital), "1", rsTipoAgrupacion!intTipoAgrupaDigital)
        
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "GENERACION DE FACTURA", vlstrFolioDocumento)
            
        'actualizacion de datos fiscales sólo para medico y empleados, que uno sean RFC Genericos''''''''''''''''''''''''''''''''''''''''''''''
        If vlblnAcutalizaDatosFiscales And vllngExp <> 0 And (vlstrTPaciente = "EM" Or vlstrTPaciente = "ME") And vlstrRFC <> "XEXX010101000" And vlstrRFC <> "XAXX010101000" Then
           vlstrsql = vllngExp & "|1|" & fStrRFCValido(vlstrRFC) & _
                                   "|" & vlstrDireccion & _
                                   "|" & vlstrNumeroExterior & _
                                   "|" & vlstrNumeroInterior & _
                                   "|" & vlstrColonia & _
                                   "|" & lngCveCiudad & _
                                   "|" & vlstrCP & _
                                   "|" & vlstrTelefono
            frsEjecuta_SP vlstrsql, "SP_UpdDatosFiscalesExPaciente"
        End If
        
        '----------------------------------
        'Agregamos los movimientos al corte
        '----------------------------------
        vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte, True)
        
        If vllngCorteUsado = 0 Then
             EntornoSIHO.ConeccionSIHO.RollbackTrans
             'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
             MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
             Exit Sub
        End If
                   
        If vllngCorteUsado <> vllngNumeroCorte Then
           'actualizamos el corte en el que se registró la factura, esto es por si hay un cambio de corte al momento de hacer el registro d ela información de la factura
            pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & vllngConsecutivoFactura
        End If
        
        If intTipoEmisionComprobante = 2 Then
           If Not fblnValidaDatosCFDCFDi(vllngConsecutivoFactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacion), strNumeroAprobacion) Then
              EntornoSIHO.ConeccionSIHO.RollbackTrans
              Exit Sub
           End If
        End If
              
        EntornoSIHO.ConeccionSIHO.CommitTrans '*
                
        '-----------------------
        'Facturación electrónica
        '-----------------------
        If intTipoEmisionComprobante = 2 Then
           
          'Barra de progreso CFD
           pgbBarraCFD.Value = 70
           freBarraCFD.Top = 3200
           Screen.MousePointer = vbHourglass
           lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
           freBarraCFD.Visible = True
           freBarraCFD.Refresh
           frmConsultaPOS.Enabled = False
           If intTipoCFDFactura = 1 Then
              pLogTimbrado 2
              pMarcarPendienteTimbre vllngConsecutivoFactura, "FA", vgintNumeroDepartamento 'factura pendiente de timbre fiscal
           End If
           EntornoSIHO.ConeccionSIHO.BeginTrans 'inicia el proceso de timbrado de la factura
           
           If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", intTipoAgrupacion, CInt(strAnoAprobacion), strNumeroAprobacion, IIf(intTipoCFDFactura = 1, True, False)) Then
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 If intTipoCFDFactura = 1 Then pLogTimbrado 1
               On Error Resume Next
               If vgIntBanderaTImbradoPendiente = 1 Then
                  'El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
                   MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura"), vbInformation + vbOKOnly, "Mensaje"
               ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then
                      '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
                      MsgBox SIHOMsg(1338), vbCritical + vbOKOnly, "Mensaje"
                      pCancelarFactura Trim(vlstrFolioDocumento), vllngPersonaGraba, "frmFacturacion", True
                      
                      'Actualiza PDF al cancelar facturas
                      If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", 1, 0, "", False, True, -1) Then
                             On Error Resume Next
                      End If
                      
                      'imprimimos la factura cancelada
                      fblnImprimeComprobanteDigital vllngConsecutivoFactura, "FA", "I", llngFormatoFactura, intTipoAgrupacion
                      Screen.MousePointer = vbDefault
                      freBarraCFD.Visible = False
                      frmConsultaPOS.Enabled = True
                      Exit Sub
               End If
           Else
               EntornoSIHO.ConeccionSIHO.CommitTrans
               If intTipoCFDFactura = 1 Then
                  pLogTimbrado 1
                  pEliminaPendientesTimbre vllngConsecutivoFactura, "FA" 'quitamos la factur de pendientes de timbre fiscal
               End If
           End If
           'Barra de progreso CFD
           pgbBarraCFD.Value = 100
           freBarraCFD.Top = 3200
           Screen.MousePointer = vbDefault
           freBarraCFD.Visible = False
           frmConsultaPOS.Enabled = True
        End If
        
        '--------------------------------------
        ' Impresión de la Factura
        '--------------------------------------
        MsgBox SIHOMsg(343), vbInformation, "Mensaje"
        
        '|  facturación electrónica
        If intTipoEmisionComprobante = 2 Then
            If Not fblnImprimeComprobanteDigital(vllngConsecutivoFactura, "FA", "I", llngFormatoFactura, intTipoAgrupacion) Then
                Exit Sub
            End If
            
            'Verifica el parámetro de envío de CFDs por correo
            If fblnPermitirEnvio(lngCuentaTicket, vllngNumeroCliente) And vgIntBanderaTImbradoPendiente = 0 Then
                '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pEnviarCFD "FA", vllngConsecutivoFactura, CLng(vgintClaveEmpresaContable), Trim(vlstrRFC), vllngPersonaGraba, Me
                End If
            End If
        Else
            pImprimeFormato llngFormatoFactura, vllngConsecutivoFactura
        End If
   Else
      EntornoSIHO.ConeccionSIHO.RollbackTrans
   End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarFacturaPaciente_Click"))
End Sub
Private Sub pCrearMovtoCredito( _
                                dblCantidadCredito As Double, _
                                lngNumCliente As Long, _
                                lngCtaContable As Long, _
                                strFolio As String, _
                                strTipo As String, _
                                dbltotalfact As Double, _
                                dblivafact As Double, _
                                lngPersonaGraba As Long)

    Dim dblTotalFactura As Double
    Dim dblIVAFactura As Double
    Dim dblPorcentaje As Double
    Dim dblSubtotalCredito As Double
    Dim dblIVACredito As Double
    Dim lngMovimientoCredito As Long
    Dim vlaryParametrosSalida() As String
    
    dblTotalFactura = Val(Format(dbltotalfact, "############.##"))
    dblIVAFactura = Val(Format(dblivafact, "############.##"))
    dblPorcentaje = dblCantidadCredito / (dblTotalFactura)
    dblSubtotalCredito = Format((dblTotalFactura - dblIVAFactura) * dblPorcentaje, "###########.##")
    dblIVACredito = Val(Format((dblIVAFactura * dblPorcentaje), "###########.##"))
    
    vgstrParametrosSP = _
                        fstrFechaSQL(fdtmServerFecha) _
                        & "|" & lngNumCliente _
                        & "|" & lngCtaContable _
                        & "|" & strFolio _
                        & "|" & strTipo _
                        & "|" & dblCantidadCredito _
                        & "|" & Str(vgintNumeroDepartamento) _
                        & "|" & Str(lngPersonaGraba) _
                        & "|" & " " & "|" & "0" & "|" & dblSubtotalCredito & "|" & dblIVACredito
    lngMovimientoCredito = 1
    frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, lngMovimientoCredito
End Sub

Private Sub pInicioGridCargos()
    With grdArticulos
        .Clear
        .ClearStructure
        .Cols = 6
        .Rows = 2
        .RowData(1) = -1
    End With
End Sub

Private Sub pConfiguraGridCargos()
    With grdArticulos
        .Cols = 18
        .FixedCols = 1
        .FixedRows = 1
        .GridLines = flexGridRaised
        .SelectionMode = flexSelectionByRow
        .FormatString = "|Descripción|Precio|Cant.|Importe|Descuento|IEPS|Subtotal|IVA|Total"
        '.FormatString = "|Descripción|Precio|Cant.|Subtotal|Descuento|Total|IEPS|IVA|Monto total"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 3000 'Descripción
        .ColWidth(2) = 1000 'Precio
        .ColWidth(3) = 450  'Cantidad
        .ColWidth(4) = 1200 'Importe
        .ColWidth(5) = 1200 'Descuentos
        .ColWidth(6) = 1200 'IEPS @
        .ColWidth(7) = 1200 'Subtotal @
        .ColWidth(8) = 1200 'Cantidad IVA
        .ColWidth(9) = 1500 'Total total
        .ColWidth(10) = 0   'Clave del cargo
        .ColWidth(11) = 0   'IVA
        .ColWidth(12) = 0   'Tipo de Cargo
        .ColWidth(13) = 0   'Concepto de Facturación
        .ColWidth(14) = 0   'Tipo de descuento del inventario
        .ColWidth(15) = 0   'Clave con la que se realizó el cargo en la tabla de cargos
        .ColWidth(16) = 0   'Clave del ticket
        .ColWidth(17) = 0   'Tasa IEPS
        .ColAlignment(0) = flexAlignCenterBottom
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
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
End Sub
Private Sub cmdCancelaSelecion_Click()
    Dim vllngPersonaGraba As Long
    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    Dim vlblnbandera As Boolean
        
    '------------------------------------------------------------------
    '- ¿Al menos uno seleccionado? -
    '------------------------------------------------------------------
    vlblnbandera = False
    For vlintContador = 1 To grdBuscaTickets.Rows - 1
        If grdBuscaTickets.TextMatrix(vlintContador, 0) = "*" Then
            vlblnbandera = True
            Exit For
        End If
    Next
    If Not vlblnbandera Then Exit Sub
    
    If Not fblnValidaSeleccion Then Exit Sub
    
    '-----------------------------------------
    '   Persona que graba
    '-----------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    With grdBuscaTickets
        If .RowData(1) = -1 Then Exit Sub
        '------------------------------------------
        ' Inicio de Transaccion
        '------------------------------------------
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        For vlintContador = 1 To .Rows - 1
            If .TextMatrix(vlintContador, 0) = "*" Then 'Para los que esten seleccionados nadamas
                .Row = vlintContador
                pCancelarTicket .RowData(.Row), vllngPersonaGraba, .TextMatrix(.Row, 1)
            End If
        Next
        EntornoSIHO.ConeccionSIHO.CommitTrans
        pCargaBusquedaT 0, grdBuscaTickets
    End With
End Sub
Private Function pCancelarTicket(vllngNumTicket As Long, vllngPersonaGraba As Long, vlstrFolioTicket As String)
    Dim vllngNumeroCorte As Long
    Dim vlstrSentencia As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPvVentaPublico As New ADODB.Recordset
    Dim rsPvDocumentoCancelado As New ADODB.Recordset
    Dim vllngCorteGrabando As Long
    Dim vllngValorDevuelto As Long
    Dim lnPagos As Long
    Dim strParametros As String
    
On Error GoTo NotificaError
    
    lnPagos = 1
    
    ' Función para obtener pagos del crédito
    strParametros = "TI" & "|" & vlstrFolioTicket
    frsEjecuta_SP strParametros, "FN_CCSELPAGOSCREDITO", True, lnPagos
        
    If lnPagos = 0 Then
        'Rs de DocumentoCancelado
        vlstrSentencia = "Select * from pvDocumentoCancelado where chrFoliodocumento = '-1'"
        Set rsPvDocumentoCancelado = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        
        'Rs de pvDetalleVentaPublico
        vlstrSentencia = "select * from pvVentaPublico " & _
                        " where pvVentaPublico.chrTipoRecivo = 'T' " & _
                        " and pvVentaPublico.intCveDepartamento = " & Trim(Str(vgintNumeroDepartamento)) & _
                        " and pvVentaPublico.intCveVenta = " & Trim(Str(vllngNumTicket))
        Set rsPvVentaPublico = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If IsNull(rsPvVentaPublico!chrfoliofactura) Or Trim(rsPvVentaPublico!chrfoliofactura) = "" Then
            vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
        
            vllngCorteGrabando = 1
            frsEjecuta_SP CStr(vllngNumeroCorte) & "|Grabando", "SP_PVUPDESTATUSCORTE", True, vllngCorteGrabando
            If vllngCorteGrabando = 2 Then
                '------------------------------------------------------------------
                '1.- Borrar los cargos que se hicieron en la venta
                '------------------------------------------------------------------
                
                ' 30/03/2009 E.O.
                ' Se agrega opción para que regrese el numero de almacén en lugar del numero de departamento que regresa el artículo.
                
                Dim rsVentaAlmacen As New ADODB.Recordset
                Dim vlintAlmacenVenta As Integer
                vlstrSentencia = "select smicvedepartamento from nodepartamento where nodepartamento.chrClasificacion = 'A' and smicvedepartamento=" & vgintNumeroDepartamento
                Set rsVentaAlmacen = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                If Not rsVentaAlmacen.RecordCount = 0 Then
                    With rsVentaAlmacen
                        vlintAlmacenVenta = !SMICVEDEPARTAMENTO
                    End With
                Else
                    vlstrSentencia = "select intnumalmacen from pvalmacenes where intnumdepartamento =" & vgintNumeroDepartamento
                    Set rsVentaAlmacen = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    With rsVentaAlmacen
                        vlintAlmacenVenta = !intnumalmacen
                    End With
                End If
                
                vlstrSentencia = " select * from pvCargo " & _
                                " where chrTipoDocumento = 'T' " & _
                                " and intFolioDocumento = " & Trim(Str(vllngNumTicket))
                Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsTemp.RecordCount > 0 Then
                    Do While Not rsTemp.EOF
                        vgstrParametrosSP = rsTemp!IntNumCargo & "|" & "EVP" & "|" & vllngPersonaGraba & "|" & vlintAlmacenVenta & "|" & "T" & "|" & Trim(Str(vllngNumTicket)) & "|" & 0 & "|" & 0 & "|" & 2
                        vllngValorDevuelto = 1
                        frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDBORRACARGO", False, vllngValorDevuelto
                        rsTemp.MoveNext
                    Loop
                End If
                rsTemp.Close
                
                '2.- Se cancela el Ticket
                vlstrSentencia = "update PvVentaPublico set bitCancelado = 1 where intCveVenta = " & Trim(Str(vllngNumTicket))
                pEjecutaSentencia (vlstrSentencia)
                
                '3.- Se registra en documentos cancelados
                With rsPvDocumentoCancelado
                    .AddNew
                    !chrFolioDocumento = Trim(Str(vllngNumTicket))
                    !chrTipoDocumento = "TI"
                    !SMIDEPARTAMENTO = vgintNumeroDepartamento
                    !intEmpleado = vllngPersonaGraba
                    !dtmfecha = fdtmServerFecha
                    .Update
                End With
                
                '4.- Se afecta el corte con cantidad negativa
                vgstrParametrosSP = Trim(rsPvVentaPublico!chrFolioTicket) & "|" & vllngPersonaGraba & "|" & "TIF" & "|" & CStr(rsPvVentaPublico!intnumcorte) & "|" & CStr(vllngNumeroCorte)
                frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdCancelaDoctoCorte"
                
                '5.- Se cancela el movimiento de crédito, si es que hubo:
                vlstrSentencia = "update CcMovimientoCredito set bitCancelado = 1, dtmFechaCancelacion = " & fstrFechaSQL(fdtmServerFecha) & " where chrTipoReferencia = 'TI' and chrFolioReferencia = '" & vlstrFolioTicket & "'"
                pEjecutaSentencia vlstrSentencia
                
                '6.-  Se cancela el movimiento de la forma de pago
                pCancelaMovimiento vllngNumTicket, rsPvVentaPublico!chrFolioTicket, rsPvVentaPublico!intnumcorte, vllngNumeroCorte, vllngPersonaGraba, False
                
                pLiberaCorte vllngNumeroCorte
            End If
        End If
        
        rsPvVentaPublico.Close
        rsPvDocumentoCancelado.Close
    Else
        MsgBox SIHOMsg(1042), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCancelarTicket"))
End Function

Private Function fblnBloqueoCuenta(vllngMovimientoPaciente As Long, vlStrTipoPaciente As String) As Boolean
On Error GoTo NotificaError
    Dim X As Integer
    Dim vlblnTermina As Boolean
    Dim vlstrBloqueo As String
                
    fblnBloqueoCuenta = False
    
    vlblnTermina = False
    X = 1
    Do While X <= cgintIntentoBloqueoCuenta And Not vlblnTermina
        vlstrBloqueo = fstrBloqueaCuenta(vllngMovimientoPaciente, vlStrTipoPaciente)
        If vlstrBloqueo = "F" Then
            vlblnTermina = True
            'La cuenta ya ha sido facturada, no se pudo realizar ningún movimiento.
            MsgBox SIHOMsg(299), vbOKOnly + vbInformation, "Mensaje"
        Else
            If vlstrBloqueo = "O" Then
                If X = cgintIntentoBloqueoCuenta Then
                    vlblnTermina = True
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnBloqueoCuenta"))
End Function

Private Sub cmdDesSeleToto_Click()
    pPonQuitaLetra " "
End Sub

Private Sub cmdReimpresion_Click()
    Dim vllngPersonaGraba As Long
    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    Dim rsDatosEmpresa As New ADODB.Recordset    'RS para Reimpresion del ticket
    Dim vlstrNombreHospital As String            'Para la impresión del ticket
    Dim vlstrRegistro As String                  'Para la impresión del ticket
    Dim vlstrDireccionHospital As String         'Para la impresión del ticket
    Dim vlstrTelefonoHospital As String          'Para la impresión del ticket
    Dim vldblTotSubtotal As Double               'Subtotal de la cuenta, para el ticket
    Dim vldblTotIva As Double                    'IVA total de la cuenta, para el ticket
    Dim vldblTotDescuentos As Double             'Descuento de la cuenta, para el ticket
    Dim vlstrFolioTicket As String               'Folio del Ticket
    Dim vlrscmdSelTicket As New ADODB.Recordset
    Dim alstrParametros(10) As String
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    Dim vlblnImpresoraSerial As Boolean          'Parametro para que imprima un ticket en impresora serial o normal
    Dim rsTemp As New ADODB.Recordset
    Dim vlrsSelTicket As New ADODB.Recordset
    Dim a As Integer
    Dim vldtmFechaHoy As Date                    'Varible con la Fecha actual
    Dim vldtmHoraHoy As Date                     'Varible con la Hora actual
    Dim vldblSubtotalTicket As Double
    Dim vldblIVATicket As Double
    Dim vldblDescuentosTicket As Double
    Dim vlblnbandera As Boolean
    Dim strCurrPrinter As String
    
    '------------------------------------------------------------------
    '- ¿Al menos uno seleccionado? -
    '------------------------------------------------------------------
    vlblnbandera = False
    For vlintContador = 1 To grdBuscaTickets.Rows - 1
        If grdBuscaTickets.TextMatrix(vlintContador, 0) = "*" Then
            vlblnbandera = True
            Exit For
        End If
    Next
    If Not vlblnbandera Then Exit Sub
    
    '--------------------------------------------------------
    ' Fecha y hora del Sistema
    '--------------------------------------------------------
    vldtmFechaHoy = fdtmServerFecha
    vldtmHoraHoy = fdtmServerHora

    '-----------------------------------------
    ' ¿Esta seguro que desea reimprimir los tickets seleccionados?
    '-----------------------------------------
    If MsgBox(SIHOMsg(476), vbYesNo + vbQuestion, "Mensaje") = vbNo Then Exit Sub
    
    '-----------------------------------------
    '   Persona que graba
    '-----------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    '--------------------------------------------------------
    ' Parametro de si tiene instalada una impresora serial
    '--------------------------------------------------------
    vlblnImpresoraSerial = False
    
    vlstrSentencia = "select count(*) from PvLoginImpresoraTicket where intNumeroLogin = " & Str(vglngNumeroLogin)
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsTemp.Fields(0) <> 0 Then
        vlblnImpresoraSerial = True
    End If
    rsTemp.Close
    
    '--------------------------------------------------------
    ' Traer Leyenda de informacion al cliente
    '--------------------------------------------------------
    vlstrSentencia = "select vchleyendacliente from pvparametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsTemp = frsRegresaRs(vlstrSentencia)
    lstrLeyendaCliente = IIf(IsNull(rsTemp!vchleyendacliente), "", rsTemp!vchleyendacliente)
    rsTemp.Close
    
    With grdBuscaTickets
        If .RowData(1) = -1 Then Exit Sub
        '--------------------------------------------------------
        ' Traer datos generales del Hospital
        '--------------------------------------------------------
        
        vlstrNombreHospital = Trim(vgstrNombreHospitalCH)
        vlstrRegistro = " R.SSA " & Trim(vgstrSSACH) & " RFC " & Trim(vgstrRfCCH)
        vlstrDireccionHospital = Trim(vgstrDireccionCH) & " CP. " & Trim(vgstrCodPostalCH)
        vlstrTelefonoHospital = Trim(vgstrTelefonoCH)
        
        For vlintContador = 1 To .Rows - 1
            If .TextMatrix(vlintContador, 0) = "*" Then 'Para los que esten seleccionados nadamas
                .Row = vlintContador
                '--------------------------------------------------------
                ' Impresion del Ticket
                '--------------------------------------------------------
                Set vlrsSelTicket = frsEjecuta_SP(Val(.RowData(vlintContador)) & "|" & fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), "SP_PVSELTICKET")
                vlstrFolioTicket = vlrsSelTicket!FolioTicket
                a = vlrsSelTicket.RecordCount
                If vlblnImpresoraSerial Then
                    Dim liCveVenta As Double
                    liCveVenta = (Val(.RowData(vlintContador)))
                
                    AbrirCOM1
                    AbrirCaja
                    ActivarSonidoCaja
                    pImprimeTicket CStr(liCveVenta)   'CStr(RTrim(vlstrFolioTicket))
                    CerrarCOM1
                    vlrsSelTicket.Close
                Else
                    Set vlrscmdSelTicket = frsEjecuta_SP(.RowData(vlintContador) & "|0", "Sp_PvSelTicket")
                    If vlrscmdSelTicket.RecordCount > 0 Then
                        vlstrFolioTicket = vlrscmdSelTicket!FolioTicket
                        vlstrx = grdBuscaTickets.RowData(vlintContador)
                        Set rsReporte = frsEjecuta_SP(vlstrx & "|" & fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), "SP_PVSELTICKET")
                        If rsReporte.RecordCount > 0 Then
                            vgrptReporte.DiscardSavedData
                            
                            alstrParametros(0) = "Direccion;" & vlstrDireccionHospital
                            alstrParametros(1) = "Empleado;" & Trim(vlrscmdSelTicket!Empleado)
                            alstrParametros(2) = "FechaActual;" & UCase(Format(vlrscmdSelTicket!fecha, "dd/mmm/yyyy"))
                            alstrParametros(3) = "FolioVenta;" & "FOLIO: " & Trim(vlstrFolioTicket)
                            alstrParametros(4) = "HoraActual;" & UCase(Format(vlrscmdSelTicket!fecha, "HH:mm"))
                            alstrParametros(5) = "NombreCliente;" & vlrscmdSelTicket!CLIENTE
                            alstrParametros(6) = "NombreEmpresa;" & vlstrNombreHospital
                            alstrParametros(7) = "Registro;" & vlstrRegistro
                            alstrParametros(8) = "Telefono;" & "TEL. " & Format(RTrim(vlstrTelefonoHospital), "###-##-##")
                            alstrParametros(9) = "LeyendaInformacionCliente;" & Trim(lstrLeyendaCliente)
                            pCargaParameterFields alstrParametros, vgrptReporte
                            strCurrPrinter = fstrCurrPrinter
                            fblnAsignaImpresora vgintNumeroDepartamento, "TI"
                            pSetReportPrinterSettings vgrptReporte
                            pImprimeReporte vgrptReporte, rsReporte, "P", "Ticket de venta"
                            pSetPrinter strCurrPrinter
                        End If
                    Else
                        'No existe información con esos parámetros.
                        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
                        vlrscmdSelTicket.Close
                    End If
                End If
                
                vlstrSentencia = "update pvVentaPublico set intPersonaReimprime = " & Trim(Str(vllngPersonaGraba)) & _
                                 " where intCveVenta = " & Trim(Str(grdBuscaTickets.RowData(vlintContador)))
                pEjecutaSentencia vlstrSentencia
                Call pGuardarLogTransaccion(Me.Name, EnmConsulta, vllngPersonaGraba, "REIMPRESION DE TICKET", CStr(grdBuscaTickets.RowData(vlintContador)))
                
            End If
            If rsReporte.State <> adStateClosed Then rsReporte.Close
        Next vlintContador
        
        pCargaBusquedaT 0, grdBuscaTickets
    End With
End Sub

Private Function fintObtieneFormato() As Integer
    Dim vlstrSentencia As String
    Dim vlrsFormato As New ADODB.Recordset

    fintObtieneFormato = -1
    '----------------------------------------------------------------------
    '««  Busca un formato predeterminado para el departamento del login  »»
    '----------------------------------------------------------------------
    vlstrSentencia = "Select * From PVTICKETDEPARTAMENTO Where smiDepartamento = " & CStr(vgintNumeroDepartamento)
    Set vlrsFormato = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If vlrsFormato.RecordCount > 0 Then
        fintObtieneFormato = vlrsFormato!intCveFormatoTicket
    End If
End Function

Private Sub pImprimeTicket(pstrCveTicket As String)
    Dim vlintCont As Integer       '[  Contador general  ]
    Dim vlintSecciones As Integer  '[  Contador de secciones  ]
    Dim vlintCveFormato As Integer '[  Clave del formato  ]
    Dim vlintRenglon As Integer    '[  Sirve para identificar el cambio de renglón  ]
    Dim vlintPosicion As Integer   '[  Posición del campo(field) en el RecordSet  ]
    Dim vlstrLinea As String       '[  Cadena que se va a imprimir  ]
    Dim vlstrSeccion As String     '[  Indica la sección, E = Encabezado, C = Cuerpo, P = Pie]
    Dim vlstrSentencia As String
    Dim vlrsValoresTicket As New ADODB.Recordset
    Dim vlrsFormatoTicket As New ADODB.Recordset
    Dim vlintDesgloseIEPS As Integer
        
    vlintCveFormato = fintObtieneFormato
    '[  Si existe un formato  ]
    If vlintCveFormato <> -1 Then
       vlintDesgloseIEPS = fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0)
         Set vlrsValoresTicket = frsEjecuta_SP(pstrCveTicket & "|" & vlintDesgloseIEPS, "Sp_Pvselticket")
        '[  Si existe un ticket registrado con esa clave  ]
        If vlrsValoresTicket.RecordCount > 0 Then
            '[  Es un ciclo de tres vueltas porque son tres secciones (Encabezado, Cuerpo y Pie)  ]
            For vlintSecciones = 0 To 2
                Select Case vlintSecciones
                    Case 0 '[  Encabezado  ]
                        vlstrSeccion = "E"
                    Case 1 '[  Cuerpo  ]
                        vlstrSeccion = "C"
                    Case 2 '[  Pie  ]
                        vlstrSeccion = "P"
                End Select
                '[  Obtiene los detalles del formato del ticket  ]
                vlstrSentencia = "                  Select * "
                vlstrSentencia = vlstrSentencia & " From pvFormatoTicket "
                vlstrSentencia = vlstrSentencia & "      Inner Join pvDetalleFormatoTicket On (pvFormatoTicket.intCveFormatoTicket = pvDetalleFormatoTicket.intCveFormatoTicket)"
                vlstrSentencia = vlstrSentencia & " Where pvFormatoTicket.intCveFormatoTicket = " & vlintCveFormato & " And"
                vlstrSentencia = vlstrSentencia & "       pvDetalleFormatoTicket.vchSeccion = '" & vlstrSeccion & "'"
                vlstrSentencia = vlstrSentencia & " Order by intRenglon, intColumna"
                Set vlrsFormatoTicket = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                With vlrsFormatoTicket
                    '[  Dado que solo el cuerpo provocará un ciclo, forza a que las demás secciones solo den una vuelta  ]
                    If vlstrSeccion <> "C" Then vlrsValoresTicket.MoveLast
                    Do While Not vlrsValoresTicket.EOF
                        vlintRenglon = 0
                        '------------------------------------------------------
                        '««  Recorre el formato del ticket según su sección  »»
                        '------------------------------------------------------
                        Do While Not .EOF
                            vlintRenglon = vlrsFormatoTicket!intRenglon
                            vlstrLinea = ""
                            Do While vlintRenglon = vlrsFormatoTicket!intRenglon
                                '[  Valores fijos  ]
                                If vlrsFormatoTicket!vchTipoValor = "F" Then
                                    vlstrLinea = vlstrLinea & IIf(vlrsFormatoTicket!vchvalor = "×", " ", vlrsFormatoTicket!vchvalor)
                                Else '[  Campos insertables  ]
                                    '------------------------------------------------------
                                    '««  Localiza la posición del campo en el RecordSet  »»
                                    '------------------------------------------------------
                                    For vlintCont = 0 To vlrsValoresTicket.Fields.Count - 1
                                        If UCase(vlrsValoresTicket.Fields(vlintCont).Name) = UCase(vlrsFormatoTicket!vchvalor) Then
                                            vlintPosicion = vlintCont
                                            Exit For
                                        End If
                                    Next
                                    If (vlintDesgloseIEPS = 0 And vlintPosicion <> 36 And vlintPosicion <> 37) Or vlintDesgloseIEPS = 1 Then
                                        Select Case vlrsFormatoTicket!intTipoDato
                                            Case 0 '[  Cadena  ]
                                                If vlrsFormatoTicket!intLongMax > 0 Then
                                                    vlstrLinea = vlstrLinea & RTrim(Mid(vlrsValoresTicket.Fields(vlintPosicion), 1, vlrsFormatoTicket!intLongMax))
                                                Else
                                                    vlstrLinea = vlstrLinea & Mid(vlrsValoresTicket.Fields(vlintPosicion), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud))
                                                End If
                                            Case 1 '[  Fecha   ]
                                                vlstrLinea = vlstrLinea & RTrim(Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "dd/mmm/yyyy"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud)))
                                            Case 2 '[  Moneda  ]
                                                vlstrLinea = vlstrLinea & Space((Len(vlstrLinea) + vlrsFormatoTicket!intLonguitud - Len(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,##0.00")) - Len(vlstrLinea))) & Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,##0.00"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud))
                                            Case 3 '[  Número  ]
                                                vlstrLinea = vlstrLinea & RTrim(Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "###,###,###"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud)))
                                            Case 4 '[   Hora   ]
                                                vlstrLinea = vlstrLinea & RTrim(Mid(Format(vlrsValoresTicket.Fields(vlintPosicion), "hh:mm"), 1, IIf(vlrsFormatoTicket!intLongMax > 0, vlrsFormatoTicket!intLongMax, vlrsFormatoTicket!intLonguitud)))
                                        End Select
                                   End If
                                End If
                                .MoveNext
                                If .EOF Then Exit Do
                            Loop
                            ImprimirCOM1 vlstrLinea & vbCrLf
                        Loop
                        vlrsValoresTicket.MoveNext
                        If vlstrSeccion = "C" Then .MoveFirst
                    Loop
                    vlrsValoresTicket.MoveFirst
                End With
            Next
        Else
            '[  ¡No existe información!  ]
            MsgBox SIHOMsg(13), vbCritical, "Mensaje"
        End If
    Else
        '[  No existen registrados formatos de impresión.  ]
        MsgBox SIHOMsg(277), vbCritical, "Mensaje"
    End If
End Sub

Private Sub cmdSeleTodo_Click()
    pPonQuitaLetra "*"
End Sub

Private Sub Form_Activate()
  fblnCargaPermisos
  fblnHabilitaObjetos frmConsultaPOS
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rsProcedencia As ADODB.Recordset
    
    pInstanciaReporte vgrptReporte, "rptTicket.rpt"
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlblnLicenciaIEPS = fblLicenciaIEPS '<-------
    pPreparaIEPS
    
    Set rsProcedencia = frsEjecuta_SP("", "SP_PVSELPROCEDENCIA")
    
    If rsProcedencia.RecordCount <> 0 Then
        pLlenarCboRs cboProcedencia, rsProcedencia, 0, 1
    End If
   
    rsProcedencia.Close
    
    cboProcedencia.AddItem "VENTA AL PÚBLICO"
    cboProcedencia.ItemData(cboProcedencia.newIndex) = 10000
    
    cboProcedencia.AddItem "<TODOS>", 0
    cboProcedencia.ItemData(0) = 0
    cboProcedencia.ListIndex = 0
    
    'Validar licencia para generar la lealtad del cliente y el médico con el hospital por medio del otorgamiento de puntos
    blnLicenciaLealtadCliente = fblnLicenciaLealtadCliente
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub
Private Sub pPreparaIEPS()
If Not vlblnLicenciaIEPS Then ' no hay licencia IEPS ajustamos la pantalla
   Me.cmdCerrarMuestraTicket.Top = 4845
   Me.lblIEPS.Visible = False
   Me.TxtIEPS.Visible = False
   Me.Label5.Top = 338
   Me.txtDescuentos.Top = 300
   Me.Label3.Top = 690
   Me.txtSubtotal.Top = 645
   Me.Label4.Top = 1028
   Me.txtIVA.Top = 990
   Me.Label6.Top = 1373
   Me.txtTotal.Top = 1335
   Frame4.Height = 1920
   Me.freMuestraTicket.Height = 5400
End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If freMuestraTicket.Visible Then
        Cancel = 1
        cmdCerrarMuestraTicket_Click
    End If
End Sub
Private Sub grdBuscaFacturas_Click()
   With grdBuscaFacturas
        If .MouseCol = 0 And .MouseRow > 0 Then
           If .TextMatrix(.Row, 14) = 1 Then
                If .TextMatrix(.Row, 0) = "*" Then
                   If vllngSeleccPendienteTimbre > 0 Then
                      vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre - 1
                   End If
                      .TextMatrix(.Row, 0) = ""
                Else
                      vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                     .TextMatrix(.Row, 0) = "*"
                End If
                Me.Cmdconfirmartimbre.Enabled = vllngSeleccPendienteTimbre > 0
            ElseIf (.TextMatrix(.Row, 13) <> "NP" And .TextMatrix(.Row, 13) <> "CR") Then
                If .TextMatrix(.Row, 0) = "*" Then
                    If vllngSeleccionadas > 0 Then
                       vllngSeleccionadas = vllngSeleccionadas - 1
                    End If
                    .TextMatrix(.Row, 0) = ""
                Else
                    vllngSeleccionadas = vllngSeleccionadas + 1
                    .TextMatrix(.Row, 0) = "*"
                End If
                Me.cmdCancelaFacturasSAT.Enabled = vllngSeleccionadas > 0
           End If
        End If
    End With
End Sub

Private Sub grdBuscaTickets_Click()
    If grdBuscaTickets.Col = 1 Then Exit Sub
    'If grdBuscaTickets.TextMatrix(grdBuscaTickets.Row, 10) = "C" Then Exit Sub 'Que esta cancelado'
    If grdBuscaTickets.RowData(grdBuscaTickets.Row) = -1 Then Exit Sub 'Que No hay nada en el Grid
    
    If grdBuscaTickets.TextMatrix(grdBuscaTickets.Row, 0) = "*" Then
        grdBuscaTickets.TextMatrix(grdBuscaTickets.Row, 0) = ""
    Else
        grdBuscaTickets.Col = 0
        grdBuscaTickets.CellFontBold = vbBlackness
        grdBuscaTickets.TextMatrix(grdBuscaTickets.Row, 0) = "*"
    End If
End Sub

Private Sub grdBuscaTickets_DblClick()
    If grdBuscaTickets.RowData(grdBuscaTickets.Row) = -1 Then Exit Sub 'Que No hay nada en el Grid
    
    pConfiguraGrdTickets grdMuestraTicket
    pCargaBusquedaT grdBuscaTickets.RowData(grdBuscaTickets.Row), grdMuestraTicket
    pTotalesMuestraTicket
    sstPOS.Enabled = False
    Frame13.Enabled = False
    freMuestraTicket.Top = 720
    freMuestraTicket.Visible = True
    cmdCerrarMuestraTicket.Enabled = True
End Sub
Private Sub pTotalesMuestraTicket()
    Dim vllngContador As Long
    Dim vldblSubtotal As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim vldbltotal As Double
    Dim vldblIEPS As Double

    vldblSubtotal = 0
    vldblDescuento = 0
    vldblIVA = 0
    vldbltotal = 0
    With grdMuestraTicket
        For vllngContador = 1 To .Rows - 1
            vldblDescuento = vldblDescuento + Val(Format(.TextMatrix(vllngContador, 3), "############.##"))
            vldblIEPS = vldblIEPS + Val(Format(.TextMatrix(vllngContador, 4), "############.##"))
            vldblSubtotal = vldblSubtotal + Val(Format(.TextMatrix(vllngContador, 5), "############.##"))
            vldblIVA = vldblIVA + Val(Format(.TextMatrix(vllngContador, 6), "############.##"))
            vldbltotal = vldbltotal + Val(Format(.TextMatrix(vllngContador, 7), "############.##"))
        Next
    End With
    txtSubtotal.Text = FormatCurrency(vldblSubtotal)
    txtDescuentos.Text = FormatCurrency(vldblDescuento)
    Me.TxtIEPS.Text = IIf(vlblnLicenciaIEPS, FormatCurrency(vldblIEPS), "")
    Me.TxtIEPS.Enabled = vlblnLicenciaIEPS
    txtIVA.Text = FormatCurrency(vldblIVA)
    txtTotal.Text = FormatCurrency(vldbltotal)
End Sub

Private Sub grdBuscaTickets_GotFocus()
    If Me.sstPOS.Tab = 1 Then
        If optMostrarSolo(0).Enabled Then optMostrarSolo(0).SetFocus
    End If
End Sub

Private Sub mskFechaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If mskFechaFinal = "  /  /    " Then
            mskFechaFinal = fdtmServerFecha
        End If
        sstPOS_Click sstPOS.Tab
        If sstPOS.Tab = 0 Then
            grdBuscaTickets.SetFocus
        Else
            If optMostrarSolo(0).Enabled Then optMostrarSolo(0).SetFocus
            'grdBuscaFacturas.SetFocus
        End If
    End If
End Sub

Private Sub mskFechaInicial_GotFocus()
    mskFechaInicial.Text = fdtmServerFecha - 1
    mskFechaFinal.Text = fdtmServerFecha
    pEnfocaMkTexto mskFechaInicial
    sstPOS_Click 0
End Sub

Private Sub mskFechaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And mskFechaInicial = "  /  /    " Then
        mskFechaInicial = fdtmServerFecha
    End If
End Sub

Private Sub mskFechaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        pEnfocaMkTexto mskFechaFinal
    End If
End Sub

Private Sub optFormaPago_Click(Index As Integer)
    pCargaBusquedaT 0, grdBuscaTickets
End Sub

Private Sub optFormaPago_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If fblnCanFocus(cboProcedencia) Then cboProcedencia.SetFocus
    End If
End Sub


Private Sub optMostrarSolo_Click(Index As Integer)
    pCargaBusqueda
    If grdBuscaFacturas.Enabled = True Then grdBuscaFacturas.SetFocus
End Sub

Private Sub sstPOS_Click(PreviousTab As Integer)
    If IsDate(mskFechaInicial) And IsDate(mskFechaFinal) Then
        frmConsultaPOS.Refresh
        If sstPOS.Tab = 1 Then
            If optMostrarSolo(0) = False Then
               mskFechaFinal.Enabled = False
               mskFechaInicial.Enabled = False
            End If
            pCargaBusqueda      'Factura
            grdBuscaFacturas.Refresh
        Else
            mskFechaFinal.Enabled = True
            mskFechaInicial.Enabled = True
            pCargaBusquedaT 0, grdBuscaTickets     'Tickets
            grdBuscaTickets.Visible = True
            grdBuscaTickets.Refresh
        End If
    Else
        pEnfocaMkTexto mskFechaInicial
    End If
End Sub
Private Sub grdBuscaFacturas_DblClick()
    If grdBuscaFacturas.TextMatrix(grdBuscaFacturas.Row, 13) <> "2" And grdBuscaFacturas.TextMatrix(grdBuscaFacturas.Row, 13) <> "XX" Then
        Set frmConsultaFactura.vgfrmFacturas = prjCaja.frmConsultaPOS
        frmConsultaFactura.pConsultaFacturas grdBuscaFacturas.TextMatrix(grdBuscaFacturas.Row, 1), True
        pCargaBusqueda
    Else
        MsgBox SIHOMsg(1277), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub
Private Sub pCargaBusquedaT(vllngNumeroTicket As Long, grdGridCarga As MSHFlexGrid)  'Para los Tickets
On Error GoTo NotificaError
    Dim vlintContador As Integer
    Dim strParametros As String
    Dim rs As New ADODB.Recordset
    Dim vlstrCualTabla As String
    
    '------------------------------------------------
    ' Progres bar
    '------------------------------------------------
    pgbBarra.Value = 10
    freBarra.Visible = True
    frmConsultaPOS.Refresh
    sstPOS.MousePointer = ssHourglass
    
    If vllngNumeroTicket > 0 Then
        vlstrCualTabla = "pvDetalleVentaPublico."
    Else
        vlstrCualTabla = "pvVentaPublico."
    End If
               
    grdGridCarga.Redraw = False
    pLimpiaGrid grdGridCarga
    pConfiguraGrdTickets grdGridCarga
    
    strParametros = fstrFechaSQL(mskFechaInicial.Text, "00:00:00") & "|" & fstrFechaSQL(mskFechaFinal.Text, "23:59:59") _
                    & "|" & CStr(vgintNumeroDepartamento) & "|" & CStr(chkCancelados.Value) & "|" & CStr(chkFacturados.Value) _
                    & "|" & CStr(cboProcedencia.ItemData(cboProcedencia.ListIndex)) & "|" & CStr(vllngNumeroTicket) & "|" & IIf(optFormaPago(1).Value = True, "C", IIf(optFormaPago(2).Value = True, "E", "T"))
    Set rs = frsEjecuta_SP(strParametros, "SP_PVSELTICKETFACTURA")
    If rs.RecordCount > 10000 Then
        'El número de registros es demasiado grande, sólo se mostrarán los primeros 10,000. Pruebe con un rango de fechas menor...
        MsgBox SIHOMsg(403), vbInformation, "Mensaje"
    End If

    Do While Not rs.EOF
        With grdGridCarga
            If .RowData(1) <> -1 Then
                 .Rows = .Rows + 1
                 .Row = .Rows - 1
            End If
            If .Row Mod 50 = 0 Then
                pgbBarra.Value = (.Row / rs.RecordCount) * 100
            End If
            
            .RowData(.Row) = rs!clave
            .Col = 0
            .CellFontBold = True
            '.TextMatrix(.Row, 0) = IIf(Not IsNull(rs!chrFolioFactura), "F", IIf(rs!bitcancelado = 1, "C", ""))
            If vllngNumeroTicket = 0 Then
                .TextMatrix(.Row, 1) = IIf(IsNull(rs!FolioTicket), "", rs!FolioTicket)
            Else
                .TextMatrix(.Row, 1) = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
            End If
            
            If vllngNumeroTicket = 0 And grdGridCarga.Name = "grdBuscaTickets" Then
                .TextMatrix(.Row, 2) = IIf(IsNull(rs!chrfoliofactura), "", rs!chrfoliofactura)
            End If
            
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 3, 2)) = rs!FechaVenta
            '------------------------------------------------------------------------------------------------
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 4, 3)) = FormatCurrency(rs!Descuento, 2) 'Descuento
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 5, 4)) = FormatCurrency(rs!mnyIeps, 2) 'IEPS
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 6, 5)) = FormatCurrency(rs!Subtotal - rs!Descuento + rs!mnyIeps, 2) 'Subtotal + IEPS
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 7, 6)) = FormatCurrency(rs!IVA, 2) 'IVA
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 8, 7)) = FormatCurrency(rs!Total, 2) 'Total
            '------------------------------------------------------------------------------------------------
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 9, 8)) = IIf(IsNull(rs!Persona), "", rs!Persona)
            '.TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 10, 9)) = IIf(rs!bitcancelado = 1, "C", "") 'Estatus del ticket
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 11, 10)) = IIf(IsNull(rs!NombreEmpresaTipoPaciente), "", rs!NombreEmpresaTipoPaciente)
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 12, 11)) = IIf(IsNull(rs!Empleado), "", rs!Empleado)
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 13, 12)) = IIf(IsNull(rs!Reimpresion), "", rs!Reimpresion)
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 14, 13)) = rs!EmpresaTipoPaciente
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 15, 14)) = rs!CantidadPagada
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 16, 15)) = rs!NumeroCliente
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 17, 16)) = rs!NumeroCorte
            .TextMatrix(.Row, IIf(grdGridCarga.Name = "grdBuscaTickets", 18, 17)) = rs!NumeroCuenta
            
            'Ticket Cancelado
            If rs!bitcancelado = 1 Then
                .Col = 0
                For vlintContador = 2 To .Cols - 1
                    .Col = .Col + 1
                    If Not IsNull(rs!chrfoliofactura) Then
                        .CellForeColor = vbBlue '&HFF0000
                    Else
                        .CellForeColor = vbRed '&HFF&
                    End If
                Next
            End If

            If rs!Reimpresion <> "" Then
                .Col = 1
                For vlintContador = 2 To .Cols - 1
                    .Col = .Col + 1
                    .CellBackColor = &H80000018
                Next
            End If
        End With
        rs.MoveNext
    Loop
    grdGridCarga.Redraw = True
    rs.Close
    freBarra.Visible = False
    sstPOS.MousePointer = ssArrow

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaBusquedaT"))
End Sub

Private Sub pConfiguraGrdTickets(grdGrid As MSHFlexGrid)
    With grdGrid
        .Cols = IIf(grdGrid.Name = "grdBuscaTickets", 19, 18)
        .FixedRows = 1
        
        If grdGrid.Name = "grdBuscaTickets" Then
            .FixedCols = 2
            .FormatString = "|Ticket|Factura|Fecha|Descuento|IEPS|Subtotal|IVA|Total|Nombre del cliente||Tipo de paciente|Persona que vendió|Persona que reimprimió|Num.Corte|Num.Cuenta"
            
            .ColWidth(0) = 300  'Fix
            .ColWidth(1) = 1200 'Numero
            .ColWidth(2) = IIf(chkFacturados.Value = 0, 0, 1200) 'Folio de la factura
            .ColWidth(3) = 1430 'Fecha
            '---------------------------------------------------------
            .ColWidth(4) = 1275 'Descuento
            .ColWidth(5) = IIf(vlblnLicenciaIEPS, 1300, 0) 'IEPS
            .ColWidth(6) = 1430 'Subtotal
            .ColWidth(7) = 1300 'Iva
            .ColWidth(8) = 1430 'Total
            '---------------------------------------------------------
            .ColWidth(9) = 4000 'Empleado
            .ColWidth(10) = 0    'Cancelado "" o "C"
            .ColWidth(11) = 4000 'Nombre empresa tipo paciente de la persona a quien se le vendio
            .ColWidth(12) = 4000 'Empleado
            .ColWidth(13) = 4000 'Empleado reimprimio
            .ColWidth(14) = 0 'Clave de la empresa o convenio del paciente a quien se le vendió
            .ColWidth(15) = 0 'Cantidad que se ha pagado en un ticket a crédito
            .ColWidth(16) = 0 'Numero del cliente cuando la venta fue a credito
            .ColWidth(17) = 0 'Num. en que se guardó el ticket
            .ColWidth(18) = 0 'Num. de cuenta del ticket
            
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(3) = flexAlignLeftCenter
            .ColAlignment(4) = flexAlignRightCenter
            .ColAlignment(5) = flexAlignRightCenter
            .ColAlignment(6) = flexAlignRightCenter
            .ColAlignment(7) = flexAlignRightCenter
            .ColAlignment(8) = flexAlignRightCenter
            .ColAlignment(11) = flexAlignLeftCenter
        Else
            .FixedCols = 1
            .FormatString = "|Descripción|Fecha|Descuento|IEPS|Subtotal|IVA|Total|Nombre del cliente||Tipo de paciente|Persona que vendió|Persona que reimprimió|Num.Corte|Num.Cuenta"
            .ColWidth(0) = 100  'Fix
            .ColWidth(1) = 3990 'Descripcion
            .ColWidth(2) = 0    'Fecha
            
            '---------------------------------------------------------------
            .ColWidth(3) = 1275 'Descuento
            .ColWidth(4) = IIf(vlblnLicenciaIEPS, 1300, 0) 'IEPS
            .ColWidth(5) = 1380 'Subtotal
            .ColWidth(6) = 1250 'Iva
            .ColWidth(7) = 1380 'Total
            '---------------------------------------------------------------
            .ColWidth(8) = 0    'Empleado
            .ColWidth(9) = 0    'Cancelado "" o "C"
            .ColWidth(10) = 0 'Nombre empresa tipo paciente de la persona a quien se le vendio
            .ColWidth(11) = 0 'Empleado
            .ColWidth(12) = 0 'Empleado reimprimio
            .ColWidth(13) = 0 'Clave de la empresa o convenio del paciente a quien se le vendió
            .ColWidth(14) = 0 'Cantidad que se ha pagado en un ticket a crédito
            .ColWidth(15) = 0 'Numero del cliente cuando la venta fue a credito
            .ColWidth(16) = 0 'Num. en que se guardó el ticket
            .ColWidth(17) = 0 'Num. de cuenta del ticket
        
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(2) = flexAlignLeftCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignRightCenter
            .ColAlignment(5) = flexAlignRightCenter
            .ColAlignment(6) = flexAlignRightCenter
            .ColAlignment(7) = flexAlignRightCenter
            .ColAlignment(10) = flexAlignLeftCenter
        End If
        
        .ColAlignmentFixed = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Private Sub pPonQuitaLetra(Caracter As String)
    Dim vlCont As Integer
    If grdBuscaTickets.RowData(1) <> -1 Then ' Que esta bacio
        For vlCont = 1 To grdBuscaTickets.Rows - 1
            'If grdBuscaTickets.TextMatrix(vlCont, 10) <> "C" Then  'Osea que esta cancelada
                grdBuscaTickets.TextMatrix(vlCont, 0) = Caracter
                grdBuscaTickets.Col = 0
                grdBuscaTickets.Row = vlCont
                grdBuscaTickets.CellFontBold = vbBlackness
            'End If
        Next vlCont
    End If
End Sub

Private Sub chkCancelados_Click()
    If mskFechaFinal = "  /  /    " Then
        mskFechaFinal = fdtmServerFecha
    End If
    sstPOS_Click sstPOS.Tab
End Sub

Private Sub ImprimirLineasEnBlanco(vlintNumeroLineas As Integer)
    Dim X As Integer ' Variable que sirve para imprimir lineas en blanco

    For X = 1 To vlintNumeroLineas
        comPrinter.Output = vbLf
    Next X
End Sub

Private Sub AbrirCOM1()
    comPrinter.CommPort = 1
    comPrinter.Settings = "9600,n,8,1"

    If (comPrinter.PortOpen = True) Then
        comPrinter.PortOpen = False
        DoEvents: DoEvents
    End If

    comPrinter.PortOpen = True
    DoEvents
End Sub

Private Sub CerrarCOM1()
    comPrinter.PortOpen = False
    DoEvents
End Sub

Private Sub ImprimirCOM1(vlstrTexto As String)
    comPrinter.Output = vlstrTexto
    DoEvents
End Sub

Private Sub AbrirCaja()
    comPrinter.Output = Chr$(7) ' Commando para abrir la Caja
    DoEvents
End Sub

Private Sub ActivarSonidoCaja()
    comPrinter.Output = Chr$(30) ' Sound Buzzer
End Sub

'- CASO 7442: Cancelar el movimiento de las formas de pago del ticket -'
Private Sub pCancelaMovimiento(vllngNumTicket As Long, vlstrFolio As String, vllngCorteTicket As Long, vllngCorteActual As Long, vllngPersonaGraba As Long, vlblnAgregaEnFactura As Boolean)
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim lstrSentencia As String, lstrTipoDoc As String, lstrFecha As String
    Dim ldblCantidad As Double
    Dim vlintContador As Integer
    Dim vllngidmovimiento As Long
    Dim vlintNumArreglo As Integer
                 
    lstrSentencia = "SELECT MB.intFormaPago, MB.mnyCantidad, MB.mnyTipoCambio, FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco) AS IdBanco, mb.chrtipomovimiento " & _
                    " FROM PvMovimientoBancoForma MB " & _
                    " INNER JOIN PvFormaPago FP ON MB.intFormaPago = FP.intFormaPago " & _
                    " LEFT  JOIN CpBanco B ON B.intNumeroCuenta = FP.intCuentaContable " & _
                    " WHERE TRIM(MB.chrTipoDocumento) = 'TI' AND MB.intNumDocumento = " & vllngNumTicket & _
                    " AND MB.intNumCorte = " & vllngCorteTicket & _
                    " AND ((mb.mnycantidad > 0 AND mb.chrtipomovimiento not in ('CBA', 'CCB')) " & _
                           " OR (mb.mnycantidad < 0 AND mb.chrtipomovimiento = 'CBA')) " & _
                    " group by MB.intFormaPago, MB.mnyCantidad, MB.mnyTipoCambio, FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco), mb.chrtipomovimiento"
    Set rs = frsRegresaRs(lstrSentencia)
    If Not rs.EOF Then
        rs.MoveFirst
        If vlblnAgregaEnFactura Then
            vlintNumArreglo = UBound(aMovimientoBancoForma)
            ReDim Preserve aMovimientoBancoForma(vlintNumArreglo + rs.RecordCount)
        End If
        
        Do While Not rs.EOF
            If rs!chrTipo <> "C" Then
                lstrFecha = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) '- Fecha y hora del movimiento -'
                
                '- Revisar tipo de forma de pago para determinar movimiento de cancelación -'
                If rs!chrTipoMovimiento = "CBA" Then
                    lstrTipoDoc = "CCB"                 'Comisión bancaria
                Else
                    Select Case rs!chrTipo
                        Case "E": lstrTipoDoc = "CFT"   'Efectivo
                        Case "T": lstrTipoDoc = "CJT"   'Tarjeta de crédito
                        Case "B": lstrTipoDoc = "CTT"   'Transferencia bancaria
                        Case "H": lstrTipoDoc = "CET"   'Cheque
                    End Select
                End If
                
                ldblCantidad = rs!MNYCantidad * (-1) 'Cantidad negativa para que se tome como abono
    
                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rs!intFormaPago & "|" & rs!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rs!mnytipocambio = 0, 1, 0) & "|" & rs!mnytipocambio & "|" & lstrTipoDoc & "|" & "TI" & "|" & vllngNumTicket & "|" & _
                                    vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & lstrFecha & "|" & "1" & "|" & cgstrModulo
                frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                
                If vlblnAgregaEnFactura Then
                    vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rs!intFormaPago & "|" & rs!IdBanco & "|" & rs!MNYCantidad & "|" & _
                                        IIf(rs!mnytipocambio = 0, 1, 0) & "|" & rs!mnytipocambio & "|" & IIf(rs!chrTipoMovimiento = "CBA", rs!chrTipoMovimiento, fstrTipoMovimientoForma(rs!intFormaPago)) & "|" & "FA" & "|" & -999 & "|" & _
                                        vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & lstrFecha & "|" & "1" & "|" & cgstrModulo
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                    vlintContadorMovs = vlintContadorMovs + 1
                    aMovimientoBancoForma(vlintContadorMovs) = flngObtieneIdentity("SEC_PVMOVIMIENTOBANCOFORMA", vllngidmovimiento)
                End If
                        
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCancelaMovimiento"))
End Sub

'- CASO 7442: Regresa el número de banco de la cuenta de la forma de pago si pertenece a un banco activo -'
Private Function flngCuentaBanco(lintNumeroCuenta As Long) As Long
On Error GoTo NotificaError

    Dim rsCuentaBanco As New ADODB.Recordset
    Dim lstrSentencia As String
    
    flngCuentaBanco = 0
    lstrSentencia = "SELECT tnyNumeroBanco FROM CpBanco WHERE bitEstatus = 1 AND intNumeroCuenta = " & lintNumeroCuenta
    Set rsCuentaBanco = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsCuentaBanco.RecordCount <> 0 Then
        flngCuentaBanco = rsCuentaBanco!tnynumerobanco
    End If
    rsCuentaBanco.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":flngCuentaBanco"))
End Function

'- CASO 7442: Regresa tipo de movimiento según la forma de pago -'
Private Function fstrTipoMovimientoForma(lintCveForma As Integer) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    Dim lstrSentencia As String
    
    fstrTipoMovimientoForma = ""
    
    lstrSentencia = "SELECT * FROM PvFormaPago WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        Select Case rsForma!chrTipo
            Case "E": fstrTipoMovimientoForma = "EFV"
            Case "T": fstrTipoMovimientoForma = "TAV"
            Case "B": fstrTipoMovimientoForma = "TPV"
            Case "H": fstrTipoMovimientoForma = "CQV"
        End Select
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrTipoMovimientoForma"))
End Function

'- CASO 6217: Verifica si se va a mostrar la pantalla de envío de CFD por correo electrónico -'
Private Function fblnPermitirEnvio(lngNumPaciente As Long, lngNumCliente As Long) As Boolean
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    
    fblnPermitirEnvio = False
    
    '- Revisar que el parámetro de envío de CFD esté activado -'
    If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
        fblnPermitirEnvio = True
    Else
        Exit Function
    End If
    
    fblnPermitirEnvio = False
    
'    '- Revisar si se factura a un cliente, que no pertenezca a una empresa -'
'    If lngNumCliente <> 0 Then
'        vlstrSentencia = "SELECT chrTipoCliente FROM CcCliente WHERE intNumCliente = " & lngNumCliente
'        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
'        If rs.RecordCount <> 0 Then
'            If Trim(rs!chrTipoCliente) <> "CO" Then
'                fblnPermitirEnvio = True
'            End If
'        End If
'    Else
'    '- Revisar que el paciente no pertenezca a un convenio -'
'        vgstrParametrosSP = lngNumPaciente & "|" & Str(vgintClaveEmpresaContable)
'        If OptTipoPaciente(0).Value = True Then
'            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
'        Else
'            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
'        End If
'        If rs.RecordCount > 0 Then
'            If IsNull(rs!cveEmpresa) Then
'                fblnPermitirEnvio = True
'            End If
'        End If
'        rs.Close
'    End If

    If llngEmpresaTipoPaciente = 0 Then
        fblnPermitirEnvio = True
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPermitirEnvio"))
End Function
Private Sub pGrabaTasasIEPS(vllngConsecutivoFactura As Long)
Dim ArrId() As Double
Dim vintCantidadTasas As Integer
Dim vintPosicion As Integer
Dim vlngcont As Long
Dim vintcont As Integer
Dim vstrSentencia As String


With grdArticulos
     vintCantidadTasas = 0
     Erase ArrId
     ReDim ArrId(3, 0)
             
     'RECORREMOS EL GRID PARA REVISAR LAS TASAS, Y VER LAS CANTIDADES QUE GRAVAN
     For vlngcont = 1 To .Rows - 1
         If Val(Format(.TextMatrix(vlngcont, 6))) > 0 Then ' se cuenta con una cantidad de IEPS
            vintPosicion = 0
            For vintcont = 1 To vintCantidadTasas
                If Val(.TextMatrix(vlngcont, 17)) = ArrId(1, vintcont) Then
                   vintPosicion = vintcont
                   Exit For
                End If
            Next vintcont
            
            If vintPosicion = 0 Then 'AGREGAMOS
               vintCantidadTasas = vintCantidadTasas + 1
               ReDim Preserve ArrId(3, vintCantidadTasas)
               ArrId(1, vintCantidadTasas) = Val(.TextMatrix(vlngcont, 17)) * 100 'TASA DE IEPS aqui la tasa del IEPS se guarda tal y como esta en el grid ya que no viene dividida por 100
               ArrId(2, vintCantidadTasas) = Val(Format(.TextMatrix(vlngcont, 6), "")) 'CANTIDAD DE IEPS @
               ArrId(3, vintCantidadTasas) = (Val(Format(.TextMatrix(vlngcont, 2), "")) * Val(.TextMatrix(vlngcont, 3))) - Val(Format(.TextMatrix(vlngcont, 5), "")) 'CANTIDAD QUE GRAVA IEPS
            Else 'ACTUALIZAMOS
               ArrId(2, vintPosicion) = ArrId(2, vintPosicion) + Val(Format(.TextMatrix(vlngcont, 6), "")) 'CANTIDAD DE IEPS @
               ArrId(3, vintPosicion) = ArrId(3, vintPosicion) + (Val(Format(.TextMatrix(vlngcont, 2), "")) * Val(.TextMatrix(vlngcont, 3))) - Val(Format(.TextMatrix(vlngcont, 5), "")) 'CANTIDAD QUE GRAVA IEPS
            End If
         End If
         
     Next vlngcont
End With
     
     For vintcont = 1 To vintCantidadTasas
         vstrSentencia = "Insert into pviepscomprobante(INTCOMPROBANTE,CHRTIPOCOMPROBANTE,NUMTASAIEPS,MNYCANTIDADGRAVADA,MNYCANTIDADIEPS) " & _
                     "values(" & vllngConsecutivoFactura & ",'FA'," & ArrId(1, vintcont) & "," & ArrId(3, vintcont) & "," & ArrId(2, vintcont) & ")"
     
         pEjecutaSentencia vstrSentencia
     Next vintcont

End Sub


Private Function fblnValidaSAT() As Boolean
    Dim rsDetalleVentaPublico As ADODB.Recordset
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim rsTemp As New ADODB.Recordset
    
    If vgstrVersionCFDI <> "3.2" Then
        For vlintContador = 1 To grdBuscaTickets.Rows - 1
            If grdBuscaTickets.TextMatrix(vlintContador, 0) = "*" Then
                
                vlstrSentencia = "select" & _
                " case PVDetalleVentaPublico.chrTipoCargo when 'AR' then IVArticulo.vchNombreComercial when 'OC' then PVOtroConcepto.chrDescripcion else null end Descripcion" & _
                ", PVDetalleVentaPublico.smiCveConceptoFacturacion ConceptoFacturacion, PVDetalleVentaPublico.intCveCargo, PVDetalleVentaPublico.chrTipoCargo" & _
                " from PVDetalleVentaPublico " & _
                " left join IVArticulo on IVArticulo.intIdArticulo = PVDetalleVentaPublico.intCveCargo and PVDetalleVentaPublico.chrTipoCargo = 'AR'" & _
                " left join PVOtroConcepto on PVOtroConcepto.intCveConcepto = PVDetalleVentaPublico.intCveCargo and PVDetalleVentaPublico.chrTipoCargo = 'OC'" & _
                " where PVDetalleVentaPublico.intCveVenta = " & Trim(grdBuscaTickets.RowData(vlintContador))
                
                Set rsDetalleVentaPublico = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
                Do While Not rsDetalleVentaPublico.EOF
                    If flngCatalogoSATIdByNombreTipo("c_ClaveUnidad", rsDetalleVentaPublico!ConceptoFacturacion, "CF", 0) = 0 Then
                        If rsDetalleVentaPublico!chrTipoCargo = "AR" Then
                            'Revisa si el artículo tiene definida la clave del SAT
                            vlstrSentencia = "Select IvArticulo.intIdArticulo clave From IvArticulo " & _
                                             "inner join gnCatalogoSatRelacion on IvArticulo.intIdArticulo  = gnCatalogoSatRelacion.intCveConcepto and gnCatalogoSatRelacion.chrTipoConcepto='AR' " & _
                                             "Where  gnCatalogoSatRelacion.chrTipoConcepto='AR' and gnCatalogoSatRelacion.intCveConcepto  = " & rsDetalleVentaPublico!intCveCargo
                            Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                            If rsTemp.RecordCount <= 0 Then
                                MsgBox SIHOMsg(1549) & Chr(13) & rsDetalleVentaPublico!Descripcion & ".", vbExclamation, "Mensaje"
                                fblnValidaSAT = False
                                rsTemp.Close
                                Exit Function
                            End If
                            rsTemp.Close
                        Else
                            If rsDetalleVentaPublico!chrTipoCargo = "OC" Then
                                'Revisa si "Otro Concepto" tiene definida la clave del SAT
                                vlstrSentencia = "Select pvOtroConcepto.intCveConcepto clave From pvOtroConcepto " & _
                                                 "inner join gnCatalogoSatRelacion on pvOtroConcepto.intCveConcepto = gnCatalogoSatRelacion.intCveConcepto and gnCatalogoSatRelacion.chrTipoConcepto='OC' " & _
                                                 "Where gnCatalogoSatRelacion.chrTipoConcepto = 'OC' and gnCatalogoSatRelacion.intDiferenciador = 1 and gnCatalogoSatRelacion.intCveConcepto = " & rsDetalleVentaPublico!intCveCargo
                                Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                If rsTemp.RecordCount <= 0 Then
                                    MsgBox SIHOMsg(1549) & Chr(13) & rsDetalleVentaPublico!Descripcion & ".", vbExclamation, "Mensaje"
                                    fblnValidaSAT = False
                                    rsTemp.Close
                                    Exit Function
                                End If
                                rsTemp.Close
                            End If
                        End If
                    End If
                    
                   rsDetalleVentaPublico.MoveNext
                Loop
                rsDetalleVentaPublico.Close
            End If
        Next
    End If
    fblnValidaSAT = True
End Function




Public Sub pValidaFormato()
     
     Dim rsAgrupaDigital As New ADODB.Recordset
     Dim vlstrAgrupaDigital As String
    
     vlstrAgrupaDigital = "SELECT intTipoAgrupaDigital FROM Formato WHERE Formato.INTNUMEROFORMATO = " & vllngFormatoaUsar
     Set rsAgrupaDigital = frsRegresaRs(vlstrAgrupaDigital, adLockReadOnly, adOpenForwardOnly)
     If rsAgrupaDigital.RecordCount > 0 Then
         vlngCveFormato = rsAgrupaDigital!intTipoAgrupaDigital
      End If
      rsAgrupaDigital.Close

End Sub

Public Function fblnValidaSATotrosConceptosCONCEPTOFACTURACION()

    Dim intRow As Integer
    Dim rsDetalleVentaPublico As ADODB.Recordset
    Dim vlstrSentencia As String
    
    'Por concepto de facturación
    If vgstrVersionCFDI <> "3.2" Then
        For intRow = 1 To grdBuscaTickets.Rows - 1
            If grdBuscaTickets.TextMatrix(intRow, 0) = "*" Then
                
                vlstrSentencia = "select" & _
                " case PVDetalleVentaPublico.chrTipoCargo when 'AR' then IVArticulo.vchNombreComercial when 'OC' then PVOtroConcepto.chrDescripcion else null end Descripcion" & _
                ", PVDetalleVentaPublico.smiCveConceptoFacturacion ConceptoFacturacion, PVDetalleVentaPublico.intCveCargo, PVDetalleVentaPublico.chrTipoCargo" & _
                ", PvConceptoFacturacion.chrdescripcion" & _
                " from PVDetalleVentaPublico" & _
                " left join IVArticulo on IVArticulo.intIdArticulo = PVDetalleVentaPublico.intCveCargo and PVDetalleVentaPublico.chrTipoCargo = 'AR'" & _
                " left join PVOtroConcepto on PVOtroConcepto.intCveConcepto = PVDetalleVentaPublico.intCveCargo and PVDetalleVentaPublico.chrTipoCargo = 'OC'" & _
                " left join PvConceptoFacturacion on PVDetalleVentaPublico.smicveconceptofacturacion = PvConceptoFacturacion.smicveconcepto" & _
                " where PVDetalleVentaPublico.intCveVenta = " & Trim(grdBuscaTickets.RowData(intRow))
                
                Set rsDetalleVentaPublico = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
                Do While Not rsDetalleVentaPublico.EOF
                    If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", rsDetalleVentaPublico!ConceptoFacturacion, "CF", 1) = 0 Then
                        MsgBox "No está definida la clave del SAT para el producto/servicio " & rsDetalleVentaPublico!chrDescripcion, vbExclamation, "Mensaje"
                        fblnValidaSATotrosConceptosCONCEPTOFACTURACION = False
                        Exit Function
                    End If

                   rsDetalleVentaPublico.MoveNext
                Loop
                rsDetalleVentaPublico.Close
            End If
        Next
    End If
    fblnValidaSATotrosConceptosCONCEPTOFACTURACION = True

End Function
