VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas de crédito y cargo"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmNotas.frx":0000
   ScaleHeight     =   10800
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   1440
      TabIndex        =   151
      Top             =   11040
      Visible         =   0   'False
      Width           =   8760
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   152
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
         Left            =   90
         TabIndex        =   153
         Top             =   180
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
   End
   Begin VB.Frame freBarra 
      Height          =   1335
      Left            =   660
      TabIndex        =   52
      Top             =   11160
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbCargando 
         Height          =   360
         Left            =   165
         TabIndex        =   53
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
         ForeColor       =   &H80000009&
         Height          =   240
         Left            =   105
         TabIndex        =   54
         Top             =   180
         Width           =   7965
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Left            =   30
         Top             =   120
         Width           =   8145
      End
   End
   Begin TabDlg.SSTab sstCargos 
      Height          =   4455
      Left            =   75
      TabIndex        =   125
      Top             =   2535
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmNotas.frx":0C42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbBuscaCargos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbFacturasPaciente"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraIncluir"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkFacturasPaciente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtBuscaCargo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboFacturasPaciente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraCargos"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.Frame fraCargos 
         Caption         =   "Cargos"
         Height          =   3015
         Left            =   120
         TabIndex        =   130
         Top             =   450
         Width           =   11340
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargos 
            Height          =   2565
            Left            =   80
            TabIndex        =   131
            ToolTipText     =   "Detalle de los cargos"
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   4524
            _Version        =   393216
            GridColor       =   -2147483633
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.ComboBox cboFacturasPaciente 
         Height          =   315
         ItemData        =   "frmNotas.frx":0C5E
         Left            =   10000
         List            =   "frmNotas.frx":0C60
         Style           =   2  'Dropdown List
         TabIndex        =   134
         ToolTipText     =   "Lista de facturas del paciente"
         Top             =   120
         Width           =   1440
      End
      Begin VB.TextBox txtBuscaCargo 
         Height          =   315
         Left            =   8040
         TabIndex        =   132
         Top             =   120
         Width           =   3400
      End
      Begin VB.CheckBox chkFacturasPaciente 
         Caption         =   "Mostrar facturas sin crédito"
         Height          =   195
         Left            =   3900
         TabIndex        =   133
         ToolTipText     =   "Facturas que no son a crédito"
         Top             =   170
         Width           =   2295
      End
      Begin VB.Frame fraIncluir 
         Height          =   810
         Left            =   120
         TabIndex        =   126
         Top             =   3460
         Width           =   11340
         Begin VB.CheckBox chkPorcentajePaciente 
            Caption         =   "Usar porcentaje"
            Height          =   255
            Left            =   3900
            TabIndex        =   150
            ToolTipText     =   "Porcentaje aplicado a la nota"
            Top             =   330
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtCantidadCargo 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   8760
            TabIndex        =   128
            ToolTipText     =   "Cantidad del concepto"
            Top             =   300
            Width           =   1530
         End
         Begin VB.CommandButton cmdIncluirCargo 
            Height          =   540
            Left            =   10680
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":0C62
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "Incluir los movimientos capturados para la nota"
            Top             =   185
            UseMaskColor    =   -1  'True
            Width           =   570
         End
         Begin VB.Label lbCantidadCargo 
            Alignment       =   1  'Right Justify
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   5520
            TabIndex        =   129
            Top             =   345
            Width           =   3030
         End
      End
      Begin VB.Label lbFacturasPaciente 
         Caption         =   "Facturas"
         Height          =   255
         Left            =   9240
         TabIndex        =   148
         Top             =   170
         Width           =   735
      End
      Begin VB.Label lbBuscaCargos 
         Caption         =   "Buscar cargos"
         Height          =   255
         Left            =   6720
         TabIndex        =   135
         Top             =   170
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab sstObj 
      Height          =   11535
      Left            =   -45
      TabIndex        =   22
      Top             =   -585
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   20346
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmNotas.frx":1154
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "freBotones"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tmrCargos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraMetodoForma"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDatosCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraMotivoNota"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraEncabezadoNota"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraCliente"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraTipoPaciente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPendienteTimbre"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "sstFacturasCreditos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "freDetalleNota"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtComentario"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmNotas.frx":1170
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape3(6)"
      Tab(1).Control(1)=   "Label57(11)"
      Tab(1).Control(2)=   "Label57(10)"
      Tab(1).Control(3)=   "Label57(15)"
      Tab(1).Control(4)=   "Label57(12)"
      Tab(1).Control(5)=   "Label57(14)"
      Tab(1).Control(6)=   "Label57(8)"
      Tab(1).Control(7)=   "Label57(13)"
      Tab(1).Control(8)=   "Label57(9)"
      Tab(1).Control(9)=   "Label57(7)"
      Tab(1).Control(10)=   "Label57(6)"
      Tab(1).Control(11)=   "Frame1"
      Tab(1).Control(12)=   "optMostrarSolo(0)"
      Tab(1).Control(13)=   "optMostrarSolo(4)"
      Tab(1).Control(14)=   "optMostrarSolo(3)"
      Tab(1).Control(15)=   "optMostrarSolo(2)"
      Tab(1).Control(16)=   "optMostrarSolo(1)"
      Tab(1).Control(17)=   "Frame3"
      Tab(1).Control(18)=   "Frame5"
      Tab(1).Control(19)=   "Frame2"
      Tab(1).ControlCount=   20
      Begin VB.Frame Frame2 
         Height          =   8440
         Left            =   -74880
         TabIndex        =   23
         Top             =   1800
         Width           =   11600
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusqueda 
            Height          =   8175
            Left            =   75
            TabIndex        =   48
            ToolTipText     =   "Lista de notas en rango de fechas"
            Top             =   165
            Width           =   11445
            _ExtentX        =   20188
            _ExtentY        =   14420
            _Version        =   393216
            GridColor       =   -2147483637
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   -71235
         TabIndex        =   171
         Top             =   10350
         Width           =   4755
         Begin VB.CommandButton cmdConfirmartimbrefiscal 
            Caption         =   "Confirmar timbre fiscal"
            Enabled         =   0   'False
            Height          =   615
            Left            =   60
            Picture         =   "frmNotas.frx":118C
            TabIndex        =   49
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2175
         End
         Begin VB.CommandButton cmdCancelaNotasSAT 
            Caption         =   "Validar comprobantes pendientes de cancelación"
            Enabled         =   0   'False
            Height          =   615
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Cancelar nota(s) de crédito/cargo ante el SAT"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   -72960
         TabIndex        =   51
         Top             =   580
         Width           =   5355
         Begin VB.CommandButton cmdCargarDatos 
            Caption         =   "Cargar datos"
            Height          =   315
            Left            =   4150
            TabIndex        =   47
            Top             =   800
            Width           =   1095
         End
         Begin VB.Frame fraConsultaPaciente 
            Height          =   430
            Left            =   1080
            TabIndex        =   136
            Top             =   330
            Visible         =   0   'False
            Width           =   4180
            Begin VB.OptionButton OptConsultaTipoPaciente 
               Caption         =   "Todos"
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   46
               Top             =   155
               Width           =   855
            End
            Begin VB.OptionButton OptConsultaTipoPaciente 
               Caption         =   "Externo"
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   45
               Top             =   155
               Width           =   855
            End
            Begin VB.OptionButton OptConsultaTipoPaciente 
               Caption         =   "Interno"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   44
               Top             =   155
               Width           =   855
            End
            Begin VB.Label lbConsultaTipoPaciente 
               Caption         =   "Tipo de paciente"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   155
               Width           =   1335
            End
         End
         Begin VB.TextBox txtNombreCliente 
            Height          =   315
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   43
            ToolTipText     =   "Nombre del cliente"
            Top             =   800
            Width           =   3975
         End
         Begin VB.TextBox txtNumCliente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            MaxLength       =   20
            TabIndex        =   42
            ToolTipText     =   "Número de cliente"
            Top             =   415
            Width           =   930
         End
         Begin VB.Label lbClientePaciente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   90
            TabIndex        =   55
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de timbre fiscal"
         Height          =   255
         Index           =   1
         Left            =   -67530
         TabIndex        =   38
         Top             =   865
         Width           =   3135
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de cancelar ante el SAT"
         Height          =   255
         Index           =   2
         Left            =   -67530
         TabIndex        =   39
         Top             =   1100
         Width           =   3855
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de autorización de cancelación"
         Height          =   255
         Index           =   3
         Left            =   -67530
         TabIndex        =   40
         Top             =   1330
         Width           =   4335
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo cancelación rechazada"
         Height          =   255
         Index           =   4
         Left            =   -67530
         TabIndex        =   41
         Top             =   1560
         Width           =   2895
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar todo"
         Height          =   255
         Index           =   0
         Left            =   -67530
         TabIndex        =   37
         Top             =   650
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtComentario 
         Height          =   975
         Left            =   200
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   117
         ToolTipText     =   "Comentario adicional"
         Top             =   9480
         Width           =   7425
      End
      Begin VB.Frame freDetalleNota 
         Caption         =   "Detalle de la nota"
         Height          =   2940
         Left            =   150
         TabIndex        =   28
         Top             =   7650
         Width           =   11570
         Begin VB.TextBox txtDescuentoTot 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
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
            Left            =   9555
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1860
            Width           =   1935
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
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
            Left            =   9555
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2550
            Width           =   1935
         End
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
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
            Left            =   9555
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   2205
            Width           =   1935
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
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
            Left            =   9555
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1515
            Width           =   1935
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdNotas 
            Height          =   1225
            Left            =   75
            TabIndex        =   5
            ToolTipText     =   "Conceptos que integran la nota"
            Top             =   240
            Width           =   11400
            _ExtentX        =   20108
            _ExtentY        =   2170
            _Version        =   393216
            Cols            =   6
            GridColor       =   -2147483633
            ScrollBars      =   2
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label lblObservaciones 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Left            =   75
            TabIndex        =   116
            Top             =   1560
            Width           =   1065
         End
         Begin VB.Label lblDescuento 
            AutoSize        =   -1  'True
            Caption         =   "Descuento"
            Height          =   195
            Left            =   8310
            TabIndex        =   32
            Top             =   1920
            Width           =   780
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   8310
            TabIndex        =   31
            Top             =   2610
            Width           =   360
         End
         Begin VB.Label lblIVA 
            AutoSize        =   -1  'True
            Caption         =   "IVA"
            Height          =   195
            Left            =   8310
            TabIndex        =   30
            Top             =   2265
            Width           =   255
         End
         Begin VB.Label lblSubtotal 
            AutoSize        =   -1  'True
            Caption         =   "Subtotal"
            Height          =   195
            Left            =   8310
            TabIndex        =   29
            Top             =   1560
            Width           =   585
         End
      End
      Begin TabDlg.SSTab sstFacturasCreditos 
         Height          =   4745
         Left            =   120
         TabIndex        =   77
         Top             =   2850
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   8361
         _Version        =   393216
         TabHeight       =   494
         Enabled         =   0   'False
         TabCaption(0)   =   "Facturas"
         TabPicture(0)   =   "frmNotas.frx":167E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraConcepto"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraDetalleFactura"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "chkFacturasPagadas"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Créditos directos"
         TabPicture(1)   =   "frmNotas.frx":169A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraConceptos"
         Tab(1).Control(1)=   "Frame4"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Información del crédito"
         TabPicture(2)   =   "frmNotas.frx":16B6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lbInfCredito"
         Tab(2).Control(1)=   "lbInfSaldo"
         Tab(2).Control(2)=   "lbInfTotalpagos"
         Tab(2).Control(3)=   "lbInfTotalNotas"
         Tab(2).Control(4)=   "grdInformacionNota"
         Tab(2).Control(5)=   "txtInfCredito"
         Tab(2).Control(6)=   "txtInfSaldo"
         Tab(2).Control(7)=   "txtInfTotalpagos"
         Tab(2).Control(8)=   "txtInfTotalNotas"
         Tab(2).ControlCount=   9
         Begin VB.TextBox txtInfTotalNotas 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   -65760
            Locked          =   -1  'True
            TabIndex        =   142
            ToolTipText     =   "Total notas"
            Top             =   2520
            Width           =   2055
         End
         Begin VB.TextBox txtInfTotalpagos 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   -65760
            Locked          =   -1  'True
            TabIndex        =   141
            ToolTipText     =   "Total pagos"
            Top             =   3000
            Width           =   2055
         End
         Begin VB.TextBox txtInfSaldo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   -65760
            Locked          =   -1  'True
            TabIndex        =   140
            ToolTipText     =   "Saldo"
            Top             =   3480
            Width           =   2055
         End
         Begin VB.TextBox txtInfCredito 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   -65760
            Locked          =   -1  'True
            TabIndex        =   139
            ToolTipText     =   "Cantidad a crédito de la factura"
            Top             =   3960
            Width           =   2055
         End
         Begin VB.CheckBox chkFacturasPagadas 
            Caption         =   "Mostrar facturas sin crédito"
            Height          =   255
            Left            =   8880
            TabIndex        =   138
            ToolTipText     =   "Mostar facturas que ya fueron pagadas"
            Top             =   320
            Width           =   2535
         End
         Begin VB.Frame fraConceptos 
            Height          =   2175
            Left            =   -74880
            TabIndex        =   105
            Top             =   2640
            Width           =   11340
            Begin VB.ListBox lstConceptosFact 
               Height          =   1620
               Left            =   2040
               TabIndex        =   111
               Top             =   360
               Width           =   5895
            End
            Begin VB.TextBox txtDescuentoCR 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9195
               TabIndex        =   108
               ToolTipText     =   "Descuento del concepto"
               Top             =   1080
               Width           =   1770
            End
            Begin VB.TextBox txtCantidadCR 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9195
               TabIndex        =   107
               ToolTipText     =   "Cantidad del concepto"
               Top             =   720
               Width           =   1770
            End
            Begin VB.CommandButton cmdIncluirCR 
               Enabled         =   0   'False
               Height          =   540
               Left            =   10395
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmNotas.frx":16D2
               Style           =   1  'Graphical
               TabIndex        =   106
               ToolTipText     =   "Incluir los movimientos capturados para la nota"
               Top             =   1485
               UseMaskColor    =   -1  'True
               Width           =   570
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Conceptos de facturación"
               Height          =   195
               Left            =   120
               TabIndex        =   112
               Top             =   360
               Width           =   1830
            End
            Begin VB.Label lbDescuentoCR 
               AutoSize        =   -1  'True
               Caption         =   "Descuento"
               Height          =   195
               Left            =   8160
               TabIndex        =   110
               Top             =   1140
               Width           =   780
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
               Height          =   195
               Left            =   8160
               TabIndex        =   109
               Top             =   780
               Width           =   630
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Detalle del crédito directo"
            Height          =   2175
            Left            =   -74880
            TabIndex        =   104
            Top             =   480
            Width           =   11340
            Begin VB.ComboBox cboCreditosDirectos 
               Height          =   315
               ItemData        =   "frmNotas.frx":1BC4
               Left            =   5707
               List            =   "frmNotas.frx":1BC6
               Style           =   2  'Dropdown List
               TabIndex        =   113
               ToolTipText     =   "Lista de facturas del cliente"
               Top             =   240
               Width           =   1440
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCreditoDirecto 
               Height          =   1365
               Left            =   75
               TabIndex        =   114
               ToolTipText     =   "Detalle de los conceptos que integran la factura"
               Top             =   705
               Width           =   11175
               _ExtentX        =   19711
               _ExtentY        =   2408
               _Version        =   393216
               GridColor       =   -2147483633
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Crédito directo"
               Height          =   195
               Left            =   4597
               TabIndex        =   115
               Top             =   300
               Width           =   1020
            End
         End
         Begin VB.Frame fraDetalleFactura 
            Caption         =   "Detalle de la factura"
            Height          =   2055
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   11340
            Begin VB.ComboBox cboFactura 
               Height          =   315
               ItemData        =   "frmNotas.frx":1BC8
               Left            =   720
               List            =   "frmNotas.frx":1BCA
               Style           =   2  'Dropdown List
               TabIndex        =   98
               ToolTipText     =   "Lista de facturas del cliente"
               Top             =   240
               Width           =   1710
            End
            Begin VB.TextBox txtCuenta 
               Height          =   315
               Left            =   4995
               Locked          =   -1  'True
               TabIndex        =   97
               ToolTipText     =   "Cuenta a la que pertenece la factura"
               Top             =   240
               Width           =   1050
            End
            Begin VB.TextBox txtNombrePaciente 
               Height          =   315
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   96
               ToolTipText     =   "Paciente al que pertenece la factura"
               Top             =   240
               Width           =   5190
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFactura 
               Height          =   1365
               Left            =   75
               TabIndex        =   99
               ToolTipText     =   "Detalle de los conceptos que integran la factura"
               Top             =   585
               Width           =   11175
               _ExtentX        =   19711
               _ExtentY        =   2408
               _Version        =   393216
               GridColor       =   -2147483633
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
            End
            Begin VB.Label lbFactura 
               AutoSize        =   -1  'True
               Caption         =   "Factura"
               Height          =   195
               Left            =   90
               TabIndex        =   103
               Top             =   300
               Width           =   540
            End
            Begin VB.Label lbCuenta 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta"
               Height          =   195
               Left            =   4425
               TabIndex        =   102
               Top             =   300
               Width           =   510
            End
            Begin VB.Label lbFecha 
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               Height          =   195
               Left            =   2556
               TabIndex        =   101
               Top             =   300
               Width           =   450
            End
            Begin VB.Label lblFechaFactura 
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   3090
               TabIndex        =   100
               Top             =   240
               Width           =   1245
            End
         End
         Begin VB.Frame fraConcepto 
            Height          =   2130
            Left            =   120
            TabIndex        =   78
            Top             =   2520
            Width           =   11340
            Begin VB.CheckBox chkPorcentaje 
               Caption         =   "Usar porcentaje"
               Height          =   255
               Left            =   9360
               TabIndex        =   149
               ToolTipText     =   "Porcentaje aplicado a la nota"
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtDescuento 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9360
               TabIndex        =   82
               ToolTipText     =   "Descuento del concepto"
               Top             =   1080
               Width           =   1530
            End
            Begin VB.TextBox txtCantidad 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9360
               TabIndex        =   81
               ToolTipText     =   "Cantidad del concepto"
               Top             =   720
               Width           =   1530
            End
            Begin VB.CommandButton cmdIncluir 
               Height          =   540
               Left            =   10320
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmNotas.frx":1BCC
               Style           =   1  'Graphical
               TabIndex        =   80
               ToolTipText     =   "Incluir los movimientos capturados para la nota"
               Top             =   1485
               UseMaskColor    =   -1  'True
               Width           =   570
            End
            Begin VB.CommandButton cmdCargar 
               Caption         =   "Cargar información"
               Height          =   315
               Left            =   6060
               TabIndex        =   79
               ToolTipText     =   "Cargar los elementos a la nota."
               Top             =   560
               Visible         =   0   'False
               Width           =   1560
            End
            Begin TabDlg.SSTab sstElementos 
               Height          =   1860
               Left            =   120
               TabIndex        =   83
               ToolTipText     =   "Elementos que puede contener la nota."
               Top             =   180
               Width           =   7635
               _ExtentX        =   13467
               _ExtentY        =   3281
               _Version        =   393216
               Style           =   1
               Tabs            =   5
               Tab             =   4
               TabsPerRow      =   5
               TabHeight       =   529
               WordWrap        =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Artículos"
               TabPicture(0)   =   "frmNotas.frx":20BE
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "chkMedicamentos"
               Tab(0).Control(1)=   "optDescripcion"
               Tab(0).Control(2)=   "optClave"
               Tab(0).Control(3)=   "txtSeleArticulo"
               Tab(0).Control(4)=   "lstArticulos"
               Tab(0).ControlCount=   5
               TabCaption(1)   =   "Servicios auxiliares"
               TabPicture(1)   =   "frmNotas.frx":20DA
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "lstEstudios"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Laboratorio"
               TabPicture(2)   =   "frmNotas.frx":20F6
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "lstExamenes"
               Tab(2).ControlCount=   1
               TabCaption(3)   =   "Otros conceptos"
               TabPicture(3)   =   "frmNotas.frx":2112
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "lstOtrosConceptos"
               Tab(3).ControlCount=   1
               TabCaption(4)   =   "Conceptos de facturación"
               TabPicture(4)   =   "frmNotas.frx":212E
               Tab(4).ControlEnabled=   -1  'True
               Tab(4).Control(0)=   "lstConceptosFacturacion"
               Tab(4).Control(0).Enabled=   0   'False
               Tab(4).ControlCount=   1
               Begin VB.ListBox lstConceptosFacturacion 
                  Height          =   840
                  Left            =   120
                  TabIndex        =   92
                  ToolTipText     =   "Lista de conceptos de facturación disponibles"
                  Top             =   810
                  Width           =   7395
               End
               Begin VB.ListBox lstOtrosConceptos 
                  Height          =   1035
                  ItemData        =   "frmNotas.frx":214A
                  Left            =   -74880
                  List            =   "frmNotas.frx":214C
                  TabIndex        =   91
                  ToolTipText     =   "Lista de otros conceptos disponibles."
                  Top             =   740
                  Width           =   7395
               End
               Begin VB.ListBox lstExamenes 
                  Height          =   840
                  ItemData        =   "frmNotas.frx":214E
                  Left            =   -74880
                  List            =   "frmNotas.frx":2150
                  TabIndex        =   90
                  ToolTipText     =   "Lista de exámenes de laboratorio disponibles."
                  Top             =   780
                  Width           =   7400
               End
               Begin VB.ListBox lstEstudios 
                  Height          =   840
                  ItemData        =   "frmNotas.frx":2152
                  Left            =   -74880
                  List            =   "frmNotas.frx":2154
                  TabIndex        =   89
                  ToolTipText     =   "Lista de estudios de servicios auxiliares disponibles."
                  Top             =   810
                  Width           =   7400
               End
               Begin VB.ListBox lstArticulos 
                  DragIcon        =   "frmNotas.frx":2156
                  Height          =   645
                  ItemData        =   "frmNotas.frx":25A0
                  Left            =   -74880
                  List            =   "frmNotas.frx":25A2
                  TabIndex        =   88
                  ToolTipText     =   "Lista de artículos disponibles."
                  Top             =   1050
                  Width           =   7400
               End
               Begin VB.TextBox txtSeleArticulo 
                  Height          =   315
                  Left            =   -74880
                  TabIndex        =   87
                  ToolTipText     =   "Teclee la clave o la descripcion del artículo"
                  Top             =   615
                  Width           =   5260
               End
               Begin VB.OptionButton optClave 
                  Caption         =   "&Clave"
                  Height          =   225
                  Left            =   -69520
                  TabIndex        =   86
                  Top             =   670
                  Width           =   705
               End
               Begin VB.OptionButton optDescripcion 
                  Caption         =   "&Descripción"
                  Height          =   225
                  Left            =   -68700
                  TabIndex        =   85
                  Top             =   670
                  Value           =   -1  'True
                  Width           =   1140
               End
               Begin VB.CheckBox chkMedicamentos 
                  Caption         =   "Sólo medicamentos"
                  Height          =   225
                  Left            =   -69520
                  TabIndex        =   84
                  Top             =   400
                  Value           =   1  'Checked
                  Width           =   1695
               End
            End
            Begin VB.Label lbDescuento 
               AutoSize        =   -1  'True
               Caption         =   "Descuento"
               Height          =   195
               Left            =   8085
               TabIndex        =   94
               Top             =   1140
               Width           =   780
            End
            Begin VB.Label lbCantidad 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
               Height          =   195
               Left            =   8085
               TabIndex        =   93
               Top             =   780
               Width           =   630
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdInformacionNota 
            Height          =   1725
            Left            =   -74880
            TabIndex        =   143
            ToolTipText     =   "Notas dentro del crédito de la factura"
            Top             =   480
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   3043
            _Version        =   393216
            GridColor       =   -2147483633
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lbInfTotalNotas 
            Caption         =   "Total notas"
            Height          =   375
            Left            =   -67800
            TabIndex        =   147
            Top             =   2595
            Width           =   1695
         End
         Begin VB.Label lbInfTotalpagos 
            Caption         =   "Total pagos"
            Height          =   375
            Left            =   -67800
            TabIndex        =   146
            Top             =   3075
            Width           =   1695
         End
         Begin VB.Label lbInfSaldo 
            Caption         =   "Saldo"
            Height          =   375
            Left            =   -67800
            TabIndex        =   145
            Top             =   3555
            Width           =   1695
         End
         Begin VB.Label lbInfCredito 
            Caption         =   "Cantidad del crédito"
            Height          =   375
            Left            =   -67800
            TabIndex        =   144
            Top             =   4035
            Width           =   1695
         End
      End
      Begin VB.TextBox txtPendienteTimbre 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Pendiente de cancelarse ante el SAT"
         Top             =   645
         Width           =   4095
      End
      Begin VB.Frame fraTipoPaciente 
         Height          =   400
         Left            =   3120
         TabIndex        =   154
         Top             =   550
         Width           =   2895
         Begin VB.OptionButton OptTipoPaciente 
            Caption         =   "Interno"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   156
            Top             =   125
            Width           =   1095
         End
         Begin VB.OptionButton OptTipoPaciente 
            Caption         =   "Externo"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   155
            Top             =   125
            Width           =   1095
         End
      End
      Begin VB.Frame fraCliente 
         Height          =   400
         Left            =   150
         TabIndex        =   118
         Top             =   550
         Width           =   2895
         Begin VB.OptionButton optPaciente 
            Caption         =   "Paciente"
            Height          =   255
            Left            =   1560
            TabIndex        =   120
            Top             =   125
            Width           =   1215
         End
         Begin VB.OptionButton optCliente 
            Caption         =   "Cliente"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   125
            Width           =   1215
         End
      End
      Begin VB.Frame fraEncabezadoNota 
         Height          =   510
         Left            =   150
         TabIndex        =   24
         Top             =   925
         Width           =   8940
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   5490
            TabIndex        =   3
            ToolTipText     =   "Fecha de la nota"
            Top             =   130
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.OptionButton optNotaCargo 
            Caption         =   "Nota de c&argo"
            Height          =   240
            Left            =   1605
            TabIndex        =   1
            ToolTipText     =   "Nota de cargo"
            Top             =   167
            Width           =   1320
         End
         Begin VB.OptionButton optNotaCredito 
            Caption         =   "Nota de c&rédito"
            Height          =   240
            Left            =   105
            TabIndex        =   0
            ToolTipText     =   "Nota de crédito"
            Top             =   167
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.Label lblEstatus 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   7305
            TabIndex        =   4
            ToolTipText     =   "Estado de la nota"
            Top             =   130
            Width           =   1560
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   6750
            TabIndex        =   27
            Top             =   190
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha"
            Height          =   195
            Left            =   4905
            TabIndex        =   26
            Top             =   190
            Width           =   480
         End
         Begin VB.Label Label11 
            Caption         =   "Folio"
            Height          =   195
            Left            =   3105
            TabIndex        =   25
            Top             =   190
            Width           =   420
         End
         Begin VB.Label lblFolio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3570
            TabIndex        =   2
            ToolTipText     =   "Folio de la nota"
            Top             =   130
            Width           =   1245
         End
      End
      Begin VB.Frame fraMotivoNota 
         Height          =   890
         Left            =   9120
         TabIndex        =   121
         Top             =   560
         Width           =   2595
         Begin VB.OptionButton OptMotivoNota 
            Caption         =   "Otorgamiento de descuentos"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   124
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton OptMotivoNota 
            Caption         =   "Error de facturación"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   123
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lbMotivoNota 
            AutoSize        =   -1  'True
            Caption         =   "Motivo de la nota de crédito"
            Height          =   180
            Left            =   120
            TabIndex        =   122
            Top             =   140
            Width           =   2340
         End
      End
      Begin VB.Frame fraDatosCliente 
         Height          =   880
         Left            =   150
         TabIndex        =   58
         Top             =   1390
         Width           =   11570
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   8920
            Locked          =   -1  'True
            TabIndex        =   73
            ToolTipText     =   "Número de cliente"
            Top             =   480
            Width           =   2510
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   315
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   72
            ToolTipText     =   "Número de cliente"
            Top             =   480
            Width           =   7395
         End
         Begin VB.TextBox txtCliente 
            Height          =   315
            Left            =   2235
            Locked          =   -1  'True
            TabIndex        =   71
            ToolTipText     =   "Número de cliente"
            Top             =   145
            Width           =   9190
         End
         Begin VB.Frame freSeleccionaCliente 
            Caption         =   "Seleccionar cliente para capturar"
            Height          =   2385
            Left            =   11700
            TabIndex        =   63
            Top             =   255
            Visible         =   0   'False
            Width           =   2970
            Begin VB.OptionButton optTipoCliente 
               Caption         =   "&Empleado"
               Height          =   195
               Index           =   2
               Left            =   240
               TabIndex        =   70
               Top             =   240
               Width           =   1110
            End
            Begin VB.ListBox lstFPBuscaCliente 
               Height          =   1035
               Left            =   60
               TabIndex        =   69
               Top             =   1200
               Width           =   2850
            End
            Begin VB.TextBox txtFPBuscaCliente 
               Height          =   285
               Left            =   60
               TabIndex        =   68
               Top             =   915
               Width           =   2850
            End
            Begin VB.OptionButton optTipoCliente 
               Caption         =   "&Convenio"
               Height          =   195
               Index           =   1
               Left            =   1605
               TabIndex        =   67
               Top             =   240
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton optTipoCliente 
               Caption         =   "E&xterno"
               Height          =   195
               Index           =   5
               Left            =   240
               TabIndex        =   66
               Top             =   690
               Width           =   960
            End
            Begin VB.OptionButton optTipoCliente 
               Caption         =   "&Interno"
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   65
               Top             =   465
               Width           =   885
            End
            Begin VB.OptionButton optTipoCliente 
               Caption         =   "&Médico"
               Height          =   195
               Index           =   3
               Left            =   1605
               TabIndex        =   64
               Top             =   465
               Width           =   960
            End
         End
         Begin VB.TextBox txtCveCliente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   960
            MaxLength       =   15
            TabIndex        =   59
            ToolTipText     =   "Número de cliente"
            Top             =   145
            Width           =   1245
         End
         Begin VB.Label lbCliente 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   205
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Domicilio"
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   540
            Width           =   630
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Left            =   8480
            TabIndex        =   74
            Top             =   540
            Width           =   315
         End
      End
      Begin VB.Frame FraMetodoForma 
         Height          =   535
         Left            =   150
         TabIndex        =   159
         Top             =   2230
         Width           =   11570
         Begin VB.ComboBox cboFormaPago 
            Height          =   315
            Left            =   1310
            Style           =   2  'Dropdown List
            TabIndex        =   60
            ToolTipText     =   "Forma de pago"
            Top             =   155
            Width           =   2535
         End
         Begin VB.ComboBox cboMetodoPago 
            Height          =   315
            Left            =   5210
            Style           =   2  'Dropdown List
            TabIndex        =   61
            ToolTipText     =   "Método de pago"
            Top             =   155
            Width           =   2535
         End
         Begin VB.ComboBox cboUsoCFDI 
            Height          =   315
            Left            =   8920
            Style           =   2  'Dropdown List
            TabIndex        =   62
            ToolTipText     =   "Uso del CFDI"
            Top             =   155
            Width           =   2535
         End
         Begin VB.Label lblFormaPago 
            Caption         =   "Forma de pago"
            Height          =   195
            Left            =   120
            TabIndex        =   162
            Top             =   215
            Width           =   1095
         End
         Begin VB.Label lblMetodoPago 
            Caption         =   "Método de pago"
            Height          =   195
            Left            =   3920
            TabIndex        =   161
            Top             =   215
            Width           =   1215
         End
         Begin VB.Label lblUsoCFDI 
            Caption         =   "Uso del CFDI"
            Height          =   195
            Left            =   7850
            TabIndex        =   160
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.Timer tmrCargos 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   240
         Top             =   10800
      End
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   -74880
         TabIndex        =   34
         Top             =   580
         Width           =   1875
         Begin MSMask.MaskEdBox mskFechaFinal 
            Height          =   315
            Left            =   600
            TabIndex        =   36
            ToolTipText     =   "Fecha final"
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaInicial 
            Height          =   315
            Left            =   600
            TabIndex        =   35
            ToolTipText     =   "Fecha de inicio"
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   90
            TabIndex        =   57
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   90
            TabIndex        =   56
            Top             =   780
            Width           =   420
         End
      End
      Begin VB.Frame freBotones 
         Height          =   705
         Left            =   2160
         TabIndex        =   33
         Top             =   10590
         Width           =   7470
         Begin VB.CommandButton cmdAddenda 
            DisabledPicture =   "frmNotas.frx":25A4
            Enabled         =   0   'False
            Height          =   495
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":2DB6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Modifica los datos de la addenda de la aseguradora"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdConfirmartimbre 
            Caption         =   "Confirmar timbre fiscal"
            Height          =   495
            Left            =   4000
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1230
         End
         Begin VB.CommandButton cmdComprobante 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5205
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":35C8
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDesglose 
            Caption         =   "Desglose de facturas"
            Height          =   495
            Left            =   5700
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Reporte de facturas incluídas en la nota"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1710
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3510
            Picture         =   "frmNotas.frx":3EE6
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Cancelar nota"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   495
            Left            =   45
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":43D8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   495
            Left            =   540
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":454A
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   495
            Left            =   1035
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":46BC
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   495
            Left            =   1530
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":4BAE
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   495
            Left            =   2025
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":4D20
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   495
            Left            =   3000
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmNotas.frx":4E92
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Guardar el registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
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
         Left            =   -74880
         TabIndex        =   170
         Top             =   10320
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74520
         TabIndex        =   169
         Top             =   10335
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
         Left            =   -74880
         TabIndex        =   168
         Top             =   11100
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
         Left            =   -74880
         TabIndex        =   167
         Top             =   10845
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de cancelar ante el SAT"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   166
         Top             =   10590
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
         Left            =   -74880
         TabIndex        =   165
         Top             =   10575
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de autorización de cancelación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -74520
         TabIndex        =   164
         Top             =   10860
         Width           =   3060
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación rechazada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74520
         TabIndex        =   163
         Top             =   11115
         Width           =   1680
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "A"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   -65990
         TabIndex        =   158
         Top             =   10330
         Width           =   135
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Notas pendientes de timbre fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   -65700
         TabIndex        =   157
         Top             =   10335
         Width           =   2355
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   6
         Left            =   -66050
         Top             =   10320
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------
' Programa para relacion de cargos por departamento
' Fecha de programación: Marzo, 2003
'-----------------------------------------------------------------------------------

Option Explicit

Const cintTipoDocumento = 8 'Nota de crédito
Const cintColCantidadNota = 3
Const cintColDescuentoNota = 4
Const cintColIVANota = 5
Const cintColTotalNota = 6

Private Type ResumenFactura
    vlstrFolioFactura As String
    vldblSubtotal As Double
    vldblDescuento As Double
    vldblIVA As Double
    vldtmFecha As Date
    vldblTipoCambio As Double
End Type

Dim rs As New ADODB.Recordset
Dim rsDatosCliente As New ADODB.Recordset

Dim vlstrSentencia As String

Dim vllngNumeroCuenta As Long
Dim vllngNumeroTipoFormato As Long

Dim vlngCveConcepto As Long
Dim vlStrConcepto As String
Dim vldblImporteConcepto As Double
Dim vldblDescuentoConcepto As Double
Dim vldblIvaConcepto As Double
Dim vldblPorcentajeIVA As Double
Dim vllngCveCtaIngresos As Long
Dim vllngCveCtaDescuento As Long
Dim vllngCveCtaIva As Long
Dim vldblIVAcero As Double
Dim vldblIVApago As Double
Dim vldblImporteIvaCero As Double
Dim vldblImporteIvaNoCero As Double
Dim vldblImportePagado As Double


Dim vlblnConsulta As Boolean            'Estatus cuando se está consultando una nota
Dim vlblnSalir As Boolean
Dim vlstrCualLista As String
Dim lstListaSeleccionada As ListBox
Dim vlblnElementoSeleccionado As Boolean
Dim vlblnDePaso As Boolean
Dim alstrParametrosSalida() As String
Dim vlblnSerieUnica As Boolean
Dim vlstrFormato As String 'Formato de moneda
Dim strPacienteImpresion As String
Dim strFacturaImpresion As String

Dim vlblnErrorManejoFolio As Boolean 'Para saber si se ha registrado en parámetros el manejo del folio de notas, único o separado

Dim vldblTipoCambio As Double 'Tipo de cambio del día
Dim lngIDnota As Long           'Id. de la nota que se consulta o que se guardó

Dim aFacturas() As ResumenFactura
Dim aFacturasTemporal() As ResumenFactura

Dim blnPestaniaFac As Boolean
Dim blnPestaniaCR As Boolean

Dim vldblSubTotalTemporal As Double
Dim vldblDescuentoTemporal As Double
Dim vldblIvaTemporal As Double
Dim vldTotal As Double
Dim ArrgrdFactura() As Double

Dim vlstrFolio As String
Dim vlstrSerie As String
Dim vlstrAnoAprobacion As String
Dim vlstrNumeroAprobacion As String
Dim vlstrFolioDocumento As String

'Para el detalle de la nota:
Const vlintColFactura = 1
Const vlintColDescripcionConcepto = 2
Const vlintColCantidad = 3
Const vlintColDescuento = 4
Const vlintColIVA = 5
Const vlintColTotal = 6
Const vlintColCveConcepto = 7
Const vlintColCtaIngresos = 8
Const vlintColCtaDescuentos = 9
Const vlintColCtaIVA = 10
Const vlintColTipoCargo = 11
Const vlintColTipoNotaFARCDetalle = 12
Const vlintColIVANotaSinRedondear = 15
Const vlintColCantidadBase = 16
Const vlintColIVABase = 17
Const vlintColIVANotaSinRedondearBase = 18
Const vlintColFechaDocumento = 19
Const vlintColDescuentoBase = 21
Const vlintColPorcentajeFactura = 22
Const vlintColPorcentajeFacturaConDescuento = 23
Const vlintColTipoCambio = 24
Const vlintColSeleccionoOtroConcepto = 25
Const vlintColDescuentoEspecial = 26

'Para la búsqueda de notas:
Const vlintColFechaNota = 1
Const vlintColFolioNota = 2
Const vlintColTipoNota = 3
Const vlintColNumCliente = 4
Const vlintColNombreCliente = 5
Const vlintColEstadoNota = 6
Const vlintColMotivoNota = 7
Const vlintColDomicilioCliente = 8
Const vlintColRFCCliente = 9
Const vlintColSubtotalNota = 10
Const vlintColDescuentoNota = 11
Const vlintColIVANota = 12
Const vlintColTotalNota = 13
Const vlintColchrTipo = 14
Const vlintColchrEstatus = 15
Const vlintColintNumPoliza = 16
Const vlintColdtmFechaRegistro = 17
Const vlintColCuentaContable = 18
Const vlintColTipoNotaFACR = 19
Const vlintColIdNota = 20
Const vlintColPFacSAT = 21
Const vlintcolPTimbre = 22

'Para el color en la búsqueda de notas
Const vllngColorCanceladas = &HC0&
Const vllngColorActivas = &H80000012
Const llnColorPenCancelaSAT = &HC0E0FF
Const llncolorCanceladasSAT = &H80000005

'Para el tamaño de la consulta de las notas
Const clngfreDetalleNotaHeightAlta = 2955
Const clngfreDetalleNotaHeightConsulta = 7245

Const clngfreDetalleNotaTopAlta = 6540
Const clngfreDetalleNotaTopConsulta = 2400

Const clnggrdNotasHeightAlta = 1200
Const clnggrdNotasHeightConsulta = 5000

Const clngtxtSubtotalTopAlta = 1500
Const clngtxtSubtotalTopConsulta = 5760

Const clnglblSubtotalTopAlta = 1540
Const clnglblSubtotalTopConsulta = 5800

Dim BlnAjusteIVA As Boolean
Dim IntAjusteIVAContador As Integer
Dim lngNumeroCliente As Long
Dim intTipoEmisionComprobante As Integer
Dim intTipoCFDNota As Integer
Dim vlstrTipoCFD As String
Dim vldbltotal As Double
Private intLimiteAjusteIVA As Integer   ' Límite de recursiones en pAjustaIVA
Dim vlcentavo As Boolean ' si se activa entonces se realizo el ajuste del centavo al total de la factura
Public vlblnlimpiaNotas As Boolean
Public vglngCveAddenda As Long 'Clave de la addenda (en caso de aplicar para la empresa)
Dim vgstrTipoPacienteAddenda As String
Dim vglngCuentaPacienteAddenda As Long
Dim vglngCveEmpresaCliente As Long
Dim vgMostrarMsjAddenda As Boolean 'Variable que define si se mostrará el mensaje para addendas "Se incluyó más de una factura en la nota de crédito por lo que no se generará CFD con addenda.  ¿Desea continuar?"
Dim vgdtmFechaIngreso As Date  'Guarda la fecha de ingreso del paciente
Dim vgblnFacturaDirecta As Boolean 'Indica que se está aplicando una nota de crédito/cargo a una factura directa (Caso 7249)
Dim vgblnCancelada As Boolean  'Guarda el estado de la nota, si está cancelada habilitará el botón de descarga del acuse de cancelación (Caso 7994)
Dim vgblnCFDI As Boolean 'Indica si la nota es CFDi'
Dim vllngSeleccionadas As Long
Dim vllngSeleccPendienteTimbre As Long
Public blnCancelaNota As Boolean
Dim blnNoMensaje As Boolean
Dim vlintBitSaldarCuentas As Long               'Variable que indica el valor del bit pvConceptoFacturacion.BitSaldarCuentas, que nos dice si la cuenta del ingreso se salda con la del descuento
Dim vlblnCuentaIngresoSaldada As Boolean        'Variable que indica si la cuenta del ingreso fue saldada con la cuenta del descuento

Dim dblPorcentajeNota As Double

Dim vldblTotalIVAConceptosCargos As Double
Dim vldblTotalIVAConceptosSeguros As Double
Dim vldblPorcentajeFactura As Double
Dim vldblPorcentajeIVAFactura As Double

Dim vldblTotalConceptosCargos As Double
Dim vldblTotalConceptosSeguros As Double
Dim vldblTotalConceptosCargosConDescuento As Double
Dim vldblPorcentajeFacturaConDescuento As Double
Dim vlstrtipofactura As String

Dim vldblTotalConceptosCargosSIgrava As Double
Dim vldblTotalConceptosCargosConDescuentoSIgrava As Double
Dim vldblTotalConceptosCargosNOgrava As Double
Dim vldblTotalConceptosCargosConDescuentoNOgrava As Double
Dim vldblTotalConceptosSegurosSIgrava As Double
Dim vldblTotalConceptosSegurosNOgrava As Double
Dim vldblPorcentajeFacturaSIgrava As Double
Dim vldblPorcentajeFacturaNOgrava As Double
Dim vldblPorcentajeFacturaConDescuentoSIgrava As Double
Dim vldblPorcentajeFacturaConDescuentoNOgrava As Double

Dim vlblnSeleccionoOtroConcepto As Boolean      'Variable que indica si se eligió un concepto de facturación diferente al concepto seleccionado de la factura
Dim vldblTmpDescuentoEspecial As Double         'variable para guardar eldescuento del concepto, sera el puente del concepo al grd notas

Dim vlstrCodigo As String               'Código postal del cliente, que debe ser obligatorio Facturacion 4.0
Dim vlstrRegimen As String              'Regimen del receptor, que debe ser obligatorio Facturacion 4.0
Dim vlRazonSocialComprobante As String  'Razón social para el comprobante
Dim vlRFCComprobante As String          'RFC para el comprobante

Public Function fblnConceptoAseguradora(intConceptoFactura As Long) As Boolean
On Error GoTo NotificaError
    'Revisa que el concepto de facturación sea del tipo aseguradora'

    Dim rs As New ADODB.Recordset
    Dim lstrSentencia As String
    
    lstrSentencia = "SELECT count(*) aseguradora FROM pvConceptoFacturacion " & _
                    "WHERE inttipo = 1 and smicveconcepto = " & intConceptoFactura
                                       
    Set rs = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
    fblnConceptoAseguradora = IIf(rs!Aseguradora = 0, False, True)
    rs.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnConceptoAseguradora"))
    Unload Me
End Function


Private Function fstrConceptoFacturacion(vlintCveConcepto As Long, vlstrTipoCargo As String) As String
    Dim rsCF As ADODB.Recordset

On Error GoTo NotificaError

    fstrConceptoFacturacion = ""
    If vlstrTipoCargo = "OC" Then
        Set rsCF = frsRegresaRs("SELECT pvConceptoFacturacion.chrDescripcion, pvotroconcepto.chrDescripcion chrDescripcionOtroconcepto FROM pvConceptoFacturacion " & _
                                            " INNER JOIN pvotroconcepto ON pvotroconcepto.smiconceptofact = pvconceptofacturacion.smicveconcepto " & _
                                "WHERE pvotroconcepto.intcveconcepto = " & vlintCveConcepto)
    ElseIf vlstrTipoCargo = "AR" Then
        Set rsCF = frsRegresaRs("SELECT pvConceptoFacturacion.chrDescripcion, ivarticulo.vchnombrecomercial chrDescripcionArticulo FROM pvConceptoFacturacion " & _
                                            " INNER JOIN ivarticulo ON ivarticulo.smicveconceptfact = pvconceptofacturacion.smicveconcepto " & _
                                "WHERE ivarticulo.intidarticulo = " & vlintCveConcepto)
    ElseIf vlstrTipoCargo = "ES" Then
        Set rsCF = frsRegresaRs("SELECT pvConceptoFacturacion.chrDescripcion, imestudio.vchnombre chrDescripcionEstudio FROM pvConceptoFacturacion " & _
                                            " INNER JOIN imestudio ON imestudio.smiconfact = pvconceptofacturacion.smicveconcepto " & _
                                "WHERE imestudio.intcveestudio = " & vlintCveConcepto)
    ElseIf vlstrTipoCargo = "EX" Then
        Set rsCF = frsRegresaRs("SELECT pvConceptoFacturacion.chrDescripcion, laexamen.chrnombre chrDescripcionExamen FROM pvConceptoFacturacion " & _
                                            " INNER JOIN laexamen ON laexamen.smiconfact = pvconceptofacturacion.smicveconcepto " & _
                                "WHERE laexamen.intcveexamen = " & vlintCveConcepto)
    ElseIf vlstrTipoCargo = "GE" Then
        Set rsCF = frsRegresaRs("SELECT pvConceptoFacturacion.chrDescripcion, lagrupoexamen.chrnombre chrDescripcionGrupoExamen FROM pvConceptoFacturacion " & _
                                            " INNER JOIN lagrupoexamen ON lagrupoexamen.smiconfact = pvconceptofacturacion.smicveconcepto " & _
                                "WHERE lagrupoexamen.intcvegrupo = " & vlintCveConcepto)
    Else
        Set rsCF = frsRegresaRs("SELECT chrDescripcion FROM pvConceptoFacturacion WHERE smiCveConcepto = " & vlintCveConcepto)
    End If
    While Not rsCF.EOF
        If vlstrTipoCargo = "OC" Then
            fstrConceptoFacturacion = rsCF!chrdescripcion & " configurado para el otro concepto " & rsCF!chrDescripcionOtroconcepto
        ElseIf vlstrTipoCargo = "AR" Then
            fstrConceptoFacturacion = rsCF!chrdescripcion & " configurado para el artículo " & rsCF!chrDescripcionArticulo
        ElseIf vlstrTipoCargo = "ES" Then
            fstrConceptoFacturacion = rsCF!chrdescripcion & " configurado para el estudio de servicios auxiliares " & rsCF!chrDescripcionEstudio
        ElseIf vlstrTipoCargo = "EX" Then
            fstrConceptoFacturacion = rsCF!chrdescripcion & " configurado para el exámen de laboratorio " & rsCF!chrDescripcionExamen
        ElseIf vlstrTipoCargo = "GE" Then
            fstrConceptoFacturacion = rsCF!chrdescripcion & " configurado para el grupo de exámenes de laboratorio " & rsCF!chrDescripcionGrupoExamen
        Else
            fstrConceptoFacturacion = rsCF!chrdescripcion
        End If
        rsCF.MoveNext
    Wend
    rsCF.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrConceptoFacturacion"))
End Function


Public Function fblnConceptoPaquete(strFolioFactura As String, intConceptoFactura As Long) As Boolean
On Error GoTo NotificaError
    'Revisa que el concepto de facturación pertenezca a un paquete'

    Dim rs As New ADODB.Recordset
    Dim lstrSentencia As String
    
    lstrSentencia = "SELECT count(*) paquete FROM pvcargo " & _
                    "INNER JOIN pvpaquete ON pvcargo.intnumpaquete = pvpaquete.intnumpaquete " & _
                    "INNER JOIN pvconceptofacturacion ON pvpaquete.smiconceptofactura = pvconceptofacturacion.smicveconcepto " & _
                    "WHERE TRIM(pvcargo.chrfoliofactura) = TRIM('" & strFolioFactura & "') AND pvpaquete.smiconceptofactura = " & intConceptoFactura & " " & _
                    "AND pvcargo.intnumpaquete <> 0"
                                       
    Set rs = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
    fblnConceptoPaquete = IIf(rs!Paquete = 0, False, True) 'Si no corresponde a un paquete
    rs.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnConceptoPaquete"))
    Unload Me
End Function

Private Function fstrDepartamento(vlintCveDepartamento As Integer) As String
    Dim rsDep As ADODB.Recordset

On Error GoTo NotificaError

    fstrDepartamento = ""
    Set rsDep = frsRegresaRs("SELECT vchDescripcion FROM noDepartamento WHERE smiCveDepartamento = " & vlintCveDepartamento)
    While Not rsDep.EOF
        fstrDepartamento = rsDep!VCHDESCRIPCION
        rsDep.MoveNext
    Wend
    rsDep.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrDepartamento"))
End Function

Private Sub pCargaTipoPaciente()
    If fintEsInterno(vglngNumeroLogin, 11) > 0 Then
        If fintEsInterno(vglngNumeroLogin, 11) = 1 Then
            OptTipoPaciente(0).Value = True
        Else
            OptTipoPaciente(1).Value = True
        End If
    End If
End Sub

Private Sub pColorearRenglon(lngColor As Long)
    Dim intCol As Integer
    Dim i As Integer
    Dim j As Integer
    Dim blBorrar As Boolean

' Compara contenido del grid de cargos contra grid de notas
    For i = 1 To grdCargos.Rows - 1
        blBorrar = True
        grdCargos.Row = i
        For j = 1 To grdNotas.Rows - 1
            grdNotas.Row = j
            If grdCargos.TextMatrix(i, 2) = grdNotas.TextMatrix(j, 2) Then
                blBorrar = False
            End If
        Next j
        
        ' Colorea el renglón actual
        If blBorrar = True And grdCargos.CellForeColor = &HC0& Or grdNotas.TextMatrix(1, 2) = "" Then
            For intCol = 1 To grdCargos.Cols - 1
                grdCargos.Col = intCol
                grdCargos.CellForeColor = lngColor
            Next intCol
        End If
    Next i
End Sub

Private Sub pHabilitaControles()
    If optNotaCredito.Value Then
        cmdCargar.Enabled = False
        txtSeleArticulo.Enabled = False
    End If
                        
    If optNotaCargo.Value Then
        cmdCargar.Enabled = True
        txtSeleArticulo.Enabled = True
    End If
End Sub

Private Sub pHabilitaCuadros(blnValor As Boolean)
    txtCantidadCargo.Enabled = blnValor
    cmdIncluirCargo.Enabled = blnValor
End Sub

Private Sub pHabilitaFrames(blnValor As Boolean, blnMotivoNota)
    fraDetalleFactura.Enabled = blnValor
    fraConcepto.Enabled = blnValor
    freDetalleNota.Enabled = blnValor
    fraDatosCliente.Enabled = blnValor
    fraMotivoNota.Enabled = blnMotivoNota
    fraTipoPaciente.Enabled = blnValor
    sstCargos.Enabled = blnValor
    fraEncabezadoNota.Enabled = blnValor
    lblUsoCFDI.Enabled = blnValor
    lblMetodoPago.Enabled = blnValor
    lblFormaPago.Enabled = blnValor
    
    cboUsoCFDI.Enabled = blnValor
    cboMetodoPago.Enabled = blnValor
    cboFormaPago.Enabled = blnValor
End Sub

Private Sub pLimpiarTodosList()
    lstArticulos.Clear
    lstEstudios.Clear
    lstExamenes.Clear
    lstOtrosConceptos.Clear
    lstConceptosFacturacion.Clear
End Sub

Private Sub pProrratearNota(TotalNota)
    Dim i As Integer
    Dim j As Integer
    Dim dblTotalSeleccionNota As Double
    Dim dblSubTotalTemporal As Double
    Dim dblIvaTemporal As Double
    Dim dblTotal As Double
    Dim strReferencia As String
    Dim dblDescuentoTemporal As Double

On Error GoTo NotificaError
        
    '***** (CR) - Modificado para Caso 6864: >>Round<< fue cambiado por >>Format<< *****'
    
    For i = 1 To grdNotas.Rows - 1
        If optCliente.Value = True Then
            For j = 1 To grdFactura.Rows - 1
                If grdFactura.TextMatrix(j, 6) = grdNotas.TextMatrix(i, 7) And Trim(grdNotas.TextMatrix(i, 1)) = Trim(cboFactura.Text) Then
                    grdNotas.TextMatrix(i, vlintColCantidad) = FormatCurrency(Format(CDbl(grdFactura.TextMatrix(j, 2)) - CDbl(grdFactura.TextMatrix(j, 3)), vlstrFormato), 2)
                    grdNotas.TextMatrix(i, vlintColIVA) = FormatCurrency(Format(grdFactura.TextMatrix(j, 4), vlstrFormato), 2)
                    grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear) = (grdFactura.TextMatrix(j, 12))
                    grdNotas.TextMatrix(i, vlintColCantidadBase) = CDbl(grdNotas.TextMatrix(i, vlintColCantidad))
                    grdNotas.TextMatrix(i, vlintColIVABase) = CDbl(grdNotas.TextMatrix(i, vlintColIVA))
                    grdNotas.TextMatrix(i, vlintColIVANotaSinRedondearBase) = CDbl(grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear))
                    grdNotas.TextMatrix(i, vlintColDescuentoBase) = CDbl(grdNotas.TextMatrix(i, vlintColDescuento))
                    grdNotas.TextMatrix(i, vlintColFechaDocumento) = CDate(lblFechaFactura.Caption)
                End If
            Next j
        Else
            For j = 1 To grdCargos.Rows - 1
                If grdCargos.TextMatrix(j, 6) = grdNotas.TextMatrix(i, 7) And Trim(grdNotas.TextMatrix(i, 1)) = Trim(cboFacturasPaciente.Text) Then
                    If OptMotivoNota(0).Value = True Then
                        grdNotas.TextMatrix(i, vlintColCantidad) = FormatCurrency(Format(CDbl(grdCargos.TextMatrix(j, 2)), vlstrFormato), 2)
                        grdNotas.TextMatrix(i, vlintColDescuento) = FormatCurrency(Format(CDbl(grdCargos.TextMatrix(j, 3)), vlstrFormato), 2)
                    Else
                        grdNotas.TextMatrix(i, vlintColCantidad) = FormatCurrency(Format(CDbl(grdCargos.TextMatrix(j, 2)) - CDbl(grdCargos.TextMatrix(j, 3)), vlstrFormato), 2)
                    End If

                    grdNotas.TextMatrix(i, vlintColIVA) = FormatCurrency(Format(grdCargos.TextMatrix(j, 4), vlstrFormato), 2)
                    grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear) = (grdCargos.TextMatrix(j, 4))
                    grdNotas.TextMatrix(i, vlintColCantidadBase) = CDbl(grdNotas.TextMatrix(i, vlintColCantidad))
                    grdNotas.TextMatrix(i, vlintColIVABase) = CDbl(grdNotas.TextMatrix(i, vlintColIVA))
                    grdNotas.TextMatrix(i, vlintColIVANotaSinRedondearBase) = CDbl(grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear))
                    grdNotas.TextMatrix(i, vlintColDescuentoBase) = CDbl(grdNotas.TextMatrix(i, vlintColDescuento))
                    If optPaciente.Value = True And chkFacturasPaciente.Value = 1 Then
                        grdNotas.TextMatrix(i, vlintColFechaDocumento) = lblFechaFactura.Caption
                    Else
                        grdNotas.TextMatrix(i, vlintColFechaDocumento) = CDate(grdCargos.TextMatrix(j, 1))
                    End If
                End If
            Next j
        End If
    Next i

    'Obtiene el total base de la nota
    For i = 1 To grdNotas.Rows - 1
        If OptMotivoNota(0).Value = True Then
            dblTotalSeleccionNota = dblTotalSeleccionNota + CDbl(grdNotas.TextMatrix(i, vlintColCantidadBase)) - CDbl(grdNotas.TextMatrix(i, vlintColDescuento)) + CDbl(grdNotas.TextMatrix(i, vlintColIVABase))
        Else
            dblTotalSeleccionNota = dblTotalSeleccionNota + CDbl(grdNotas.TextMatrix(i, vlintColCantidadBase)) + CDbl(grdNotas.TextMatrix(i, vlintColIVABase))
            'dblTotalSeleccionNota = dblTotalSeleccionNota + CDbl(grdNotas.TextMatrix(i, vlintColCantidadBase)) + CDbl(grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear))  ' Revisar GM
        End If
    Next i

    If dblTotalSeleccionNota >= TotalNota Then
        dblPorcentajeNota = TotalNota / dblTotalSeleccionNota
    Else
        dblPorcentajeNota = 1
    End If
    'dblPorcentajeNota = Round(dblPorcentajeNota, 4)
    
    ReDim aFacturas(0)
    For i = 1 To grdNotas.Rows - 1
        'dblSubTotalTemporal = Format(dblSubTotalTemporal + CDbl((grdNotas.TextMatrix(i, vlintColCantidadBase)) * dblPorcentajeNota), vlstrFormato)
        'dblIvaTemporal = Format(dblIvaTemporal + (grdNotas.TextMatrix(i, vlintColIVANotaSinRedondearBase) * dblPorcentajeNota), vlstrFormato)
        'dblTotal = Format(dblSubTotalTemporal + dblIvaTemporal, vlstrFormato)

        grdNotas.TextMatrix(i, vlintColCantidad) = FormatCurrency(Format(grdNotas.TextMatrix(i, vlintColCantidadBase) * dblPorcentajeNota, vlstrFormato), 2)
'        grdNotas.TextMatrix(i, vlintColCantidad) = FormatCurrency(Format((grdNotas.TextMatrix(i, vlintColCantidadBase) - grdNotas.TextMatrix(i, vlintColDescuento)) * dblPorcentajeNota, vlstrFormato), 2)

        grdNotas.TextMatrix(i, vlintColDescuento) = FormatCurrency(Format(grdNotas.TextMatrix(i, vlintColDescuentoBase) * dblPorcentajeNota, vlstrFormato), 2)

        grdNotas.TextMatrix(i, vlintColIVA) = FormatCurrency(Format(grdNotas.TextMatrix(i, vlintColIVABase) * dblPorcentajeNota, vlstrFormato), 2)
        grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear) = CDbl(grdNotas.TextMatrix(i, vlintColIVANotaSinRedondearBase)) * dblPorcentajeNota

        '- (CR) - Modificado para Casos 6446 y 6864 -'
        dblSubTotalTemporal = dblSubTotalTemporal + grdNotas.TextMatrix(i, vlintColCantidad)
        'dblSubTotalTemporal = dblSubTotalTemporal + grdNotas.TextMatrix(i, vlintColCantidadBase) * dblPorcentajeNota       'Revisar GM

        dblDescuentoTemporal = dblDescuentoTemporal + grdNotas.TextMatrix(i, vlintColDescuento)
        dblIvaTemporal = dblIvaTemporal + grdNotas.TextMatrix(i, vlintColIVA)
        
        dblTotal = dblSubTotalTemporal - dblDescuentoTemporal + dblIvaTemporal
        '--------------------------------------------'

        'pAcumulaFactura grdNotas.TextMatrix(i, vlintColFactura), grdNotas.TextMatrix(i, vlintColCantidad), 0, grdNotas.TextMatrix(i, vlintColIVA), 1, CDate(grdNotas.TextMatrix(i, vlintColFechaDocumento))
        'pAcumulaFactura grdNotas.TextMatrix(i, vlintColFactura), grdNotas.TextMatrix(i, vlintColCantidadBase) * dblPorcentajeNota, grdNotas.TextMatrix(i, vlintColDescuento), grdNotas.TextMatrix(i, vlintColIVANotaSinRedondear), 1, grdNotas.TextMatrix(i, vlintColTipoCambio), CDate(grdNotas.TextMatrix(i, vlintColFechaDocumento))
        pAcumulaFactura grdNotas.TextMatrix(i, vlintColFactura), grdNotas.TextMatrix(i, vlintColCantidad), grdNotas.TextMatrix(i, vlintColDescuento), grdNotas.TextMatrix(i, vlintColIVA), 1, grdNotas.TextMatrix(i, vlintColTipoCambio), CDate(grdNotas.TextMatrix(i, vlintColFechaDocumento))
    Next i

    txtSubtotal.Text = FormatCurrency(dblSubTotalTemporal, 2)
    txtDescuentoTot.Text = FormatCurrency(dblDescuentoTemporal, 2)
    txtIVA.Text = FormatCurrency(dblIvaTemporal, 2)
    txtTotal.Text = FormatCurrency(dblTotal, 2)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pProrratearNota"))
    Unload Me
End Sub

Private Function fTotalFactura(grid, Factura As String) As Double
    Dim i As Integer
    Dim dblTotalFactura As Double
    Dim rsTotalFactura As ADODB.Recordset
    Dim vldblCentavo As Double ' identifica la diferencia entre el total de la factura y la suma de los conceptos
    
    Dim rsDatosCompDig As ADODB.Recordset
    Dim blnCentavo As Boolean
    Dim blnSeCambio As Boolean
On Error GoTo NotificaError
    
    blnSeCambio = False
    Set rsTotalFactura = frsEjecuta_SP(Factura, "Sp_PvSelFactura")
    If rsTotalFactura.RecordCount > 0 Then fTotalFactura = rsTotalFactura!TotalFactura
    rsTotalFactura.Close
    
    For i = 1 To grid.Rows - 1
        'dblTotalFactura = dblTotalFactura + Format(CDbl(grid.TextMatrix(i, 11)), "fixed")
        'dblTotalFactura = dblTotalFactura + Round(CDbl(grid.TextMatrix(i, 11)), 2)  'Revisar GM
        dblTotalFactura = dblTotalFactura + CDbl(grid.TextMatrix(i, 5))     ' Revisar GM 13052
        
    Next i
    
    If vlcentavo = False Then
        vldblCentavo = Format(Abs(fTotalFactura - dblTotalFactura), "Fixed")
        If vldblCentavo > 0.01 Then
            dblTotalFactura = 0
            For i = 1 To grid.Rows - 1
                dblTotalFactura = dblTotalFactura + CDbl(grid.TextMatrix(i, 11))
            Next i
            vldblCentavo = Abs(fTotalFactura - dblTotalFactura)
        End If
        
        dblTotalFactura = Format(dblTotalFactura, "Fixed")
        If fTotalFactura <> dblTotalFactura And vldblCentavo = 0.01 Then
            For i = 1 To grid.Rows - 1
                blnCentavo = False
                Set rsDatosCompDig = frsEjecuta_SP(Factura & "|" & CLng(grid.TextMatrix(i, 6)), "SP_CCSelComprobanteDigital")
                If rsDatosCompDig.RecordCount > 0 Then
                    If fTotalFactura > dblTotalFactura Then
                        If FormatCurrency(CDbl(rsDatosCompDig!NUMIMPORTECONCEPTO)) <> FormatCurrency(CDbl(grid.TextMatrix(i, 2))) Then
                            grid.TextMatrix(i, 2) = FormatCurrency(CDbl(grid.TextMatrix(i, 2)) + vldblCentavo)
                            grid.TextMatrix(i, 5) = FormatCurrency(CDbl(grid.TextMatrix(i, 5)) + vldblCentavo)
                            blnCentavo = True
                            blnSeCambio = True
                        Else
                            If FormatCurrency(CDbl(rsDatosCompDig!NUMIVACONCEPTO)) <> FormatCurrency(CDbl(grid.TextMatrix(i, 4))) Then
                                grid.TextMatrix(i, 4) = FormatCurrency(CDbl(grid.TextMatrix(i, 4)) + vldblCentavo)
                                grid.TextMatrix(i, 5) = FormatCurrency(CDbl(grid.TextMatrix(i, 5)) + vldblCentavo)
                                blnCentavo = True
                                blnSeCambio = True
                            End If
                        End If
                    Else
                        If FormatCurrency(CDbl(rsDatosCompDig!NUMIMPORTECONCEPTO)) <> FormatCurrency(CDbl(grid.TextMatrix(i, 2))) Then
                            grid.TextMatrix(i, 2) = FormatCurrency(CDbl(grid.TextMatrix(i, 2)) - vldblCentavo)
                            grid.TextMatrix(i, 5) = FormatCurrency(CDbl(grid.TextMatrix(i, 5)) - vldblCentavo)
                            blnCentavo = True
                            blnSeCambio = True
                        Else
                            If FormatCurrency(CDbl(rsDatosCompDig!NUMIVACONCEPTO)) <> FormatCurrency(CDbl(grid.TextMatrix(i, 4))) Then
                                grid.TextMatrix(i, 4) = FormatCurrency(CDbl(grid.TextMatrix(i, 4)) - vldblCentavo)
                                grid.TextMatrix(i, 5) = FormatCurrency(CDbl(grid.TextMatrix(i, 5)) - vldblCentavo)
                                blnCentavo = True
                                blnSeCambio = True
                            End If
                        End If
                    End If
                    rsDatosCompDig.Close
                    If blnCentavo = True Then
                        vlcentavo = True
                        Exit For
                    End If
                End If
                
            Next i
        End If
       
        If blnSeCambio = False Then
            If fTotalFactura > dblTotalFactura Then
                grid.TextMatrix(1, 2) = FormatCurrency(CDbl(grid.TextMatrix(1, 2)) + vldblCentavo)
                grid.TextMatrix(1, 5) = FormatCurrency(CDbl(grid.TextMatrix(1, 5)) + vldblCentavo)
                blnCentavo = True
            Else
                grid.TextMatrix(1, 2) = FormatCurrency(CDbl(grid.TextMatrix(1, 2)) - vldblCentavo)
                grid.TextMatrix(1, 5) = FormatCurrency(CDbl(grid.TextMatrix(1, 5)) - vldblCentavo)
                blnCentavo = True
            End If
        End If
        
        dblTotalFactura = 0
        For i = 1 To grid.Rows - 1
            dblTotalFactura = dblTotalFactura + CDbl(grid.TextMatrix(i, 5))     ' Revisar GM 13052
        Next i
        fTotalFactura = dblTotalFactura
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fTotalFactura"))
    Unload Me
End Function

Private Sub cboCreditosDirectos_Click()
On Error GoTo NotificaError
    
    Dim rsDatosCreditosDirectos As New ADODB.Recordset

    If cboCreditosDirectos.ListIndex <> -1 Then
        'Cuenta y paciente al que pertenece la factura
        Set rsDatosCreditosDirectos = frsEjecuta_SP(cboCreditosDirectos.ItemData(cboCreditosDirectos.ListIndex), "SP_CCSELCREDITODETALLE")
        If rsDatosCreditosDirectos.RecordCount <> 0 Then
            If rsDatosCreditosDirectos.RecordCount <> 0 Then
                Do While Not rsDatosCreditosDirectos.EOF
                    grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 1) = Format(rsDatosCreditosDirectos!fecha, "dd/mmm/yyyy")
                    grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 2) = FormatCurrency(rsDatosCreditosDirectos!CantidadCredito, 2)
                    grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 3) = FormatCurrency(rsDatosCreditosDirectos!CantidadPagada, 2)
                    grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 4) = FormatCurrency(rsDatosCreditosDirectos!Saldo, 2)
                    rsDatosCreditosDirectos.MoveNext
                Loop
            End If
        End If
        rsDatosCreditosDirectos.Close
        vldblImporteConcepto = CDbl(Format(grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 4), vlstrFormato))
        vldblImportePagado = CDbl(Format(grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 3), vlstrFormato))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboCreditosDirectos_Click"))
    Unload Me
End Sub

Private Sub cboCreditosDirectos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstConceptosFact.SetFocus
        lstConceptosFact.ListIndex = 0
    End If
End Sub

Private Sub cboFacturasPaciente_Click()
On Error GoTo NotificaError
    
    Dim rsDatosFISC As New ADODB.Recordset
    Dim rsDatosFactura As New ADODB.Recordset
    Dim vlblnMonedaValida As Boolean
    Dim vlintContador As Integer
    
    grdCargos.Redraw = False

    If cboFacturasPaciente.ListIndex <> -1 Then
        vgstrParametrosSP = Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))
        Set rsDatosFactura = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFactura")
        If rsDatosFactura.RecordCount <> 0 Then
            txtCuenta.Text = rsDatosFactura!cuenta
            txtNombrePaciente.Text = IIf(IsNull(rsDatosFactura!NOMBREPACIENTE), "", rsDatosFactura!NOMBREPACIENTE)
            lblFechaFactura.Caption = Format(rsDatosFactura!fecha, "dd/mmm/yyyy")
            vlstrtipofactura = rsDatosFactura!chrTipoFactura
                        
            If IIf(IsNull(rsDatosFactura!CP), "", rsDatosFactura!CP) <> "" Then
                vlstrCodigo = IIf(IsNull(rsDatosFactura!CP), "", rsDatosFactura!CP)
            End If
            
            If IIf(IsNull(rsDatosFactura!RazonSocial), "", rsDatosFactura!RazonSocial) <> "" Then
                vlRazonSocialComprobante = IIf(IsNull(rsDatosFactura!RazonSocial), "", rsDatosFactura!RazonSocial)
            End If
            
            If IIf(IsNull(rsDatosFactura!RFC), "", rsDatosFactura!RFC) <> "" Then
                vlRFCComprobante = IIf(IsNull(rsDatosFactura!RFC), "", rsDatosFactura!RFC)
            End If
            
            Set rsDatosFISC = frsRegresaRs("SELECT VCHREGIMENFISCALRECEPTOR FROM PVFACTURA WHERE trim(chrfoliofactura) = '" & Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)) & "'")
            If rsDatosFISC.RecordCount > 0 Then
                If IIf(IsNull(rsDatosFISC!VCHREGIMENFISCALRECEPTOR), "", rsDatosFISC!VCHREGIMENFISCALRECEPTOR) <> "" Then
                    vlstrRegimen = IIf(IsNull(rsDatosFISC!VCHREGIMENFISCALRECEPTOR), "", rsDatosFISC!VCHREGIMENFISCALRECEPTOR)
                End If
            End If
            
            ' --------
            ' Verifica que la moneda de la factura seleccionada sea igual a la moneda de la(s) factura(s) ya incluida(s) en la nota
            vlblnMonedaValida = True
            For vlintContador = 1 To grdNotas.Rows - 1
                If rsDatosFactura!TipoCambio = 0 Then
                    If Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 0 And Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 1 Then
                        vlblnMonedaValida = False
                        Exit For
                    End If
                ElseIf rsDatosFactura!TipoCambio <> 0 Then
                    If Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 0 And Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) = 1 Then
                        vlblnMonedaValida = False
                        Exit For
                    End If
                End If
            Next vlintContador
            If vlblnMonedaValida = False Then
                pLimpiaGridFacturaPaciente
                pConfiguraGridCargos
                'txtCantidad.Text = ""
                rsDatosFactura.Close
                MsgBox SIHOMsg(1135), vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
            End If
            ' --------
            
            pLimpiaGridFacturaPaciente
            
            If rsDatosFactura.RecordCount <> 0 Then
                Do While Not rsDatosFactura.EOF
                    If rsDatosFactura!chrTipo = "NO" Or rsDatosFactura!chrTipo = "OC" Then
                        If Trim(grdCargos.TextMatrix(1, 1)) = "" Then
                            grdCargos.Row = 1
                        Else
                            grdCargos.Rows = grdCargos.Rows + 1
                            grdCargos.Row = grdCargos.Rows - 1
                        End If
                        grdCargos.TextMatrix(grdCargos.Row, 1) = rsDatosFactura!Concepto
                        grdCargos.TextMatrix(grdCargos.Row, 2) = FormatCurrency(rsDatosFactura!Importe, 2)
                        grdCargos.TextMatrix(grdCargos.Row, 3) = FormatCurrency(rsDatosFactura!Descuento, 2)
                        grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(rsDatosFactura!IVA, 2)
                        grdCargos.TextMatrix(grdCargos.Row, 5) = FormatCurrency(rsDatosFactura!Importe - rsDatosFactura!Descuento + rsDatosFactura!IVA, 2)
                        grdCargos.TextMatrix(grdCargos.Row, 6) = rsDatosFactura!smicveconcepto
                        grdCargos.TextMatrix(grdCargos.Row, 7) = IIf(IsNull(rsDatosFactura!CuentaIngresos), 0, rsDatosFactura!CuentaIngresos)
                        grdCargos.TextMatrix(grdCargos.Row, 8) = IIf(IsNull(rsDatosFactura!CuentaDescuentos), 0, rsDatosFactura!CuentaDescuentos)
                        grdCargos.TextMatrix(grdCargos.Row, 9) = 0
                        grdCargos.TextMatrix(grdCargos.Row, 10) = Format(rsDatosFactura!IVACONCEPTO, vlstrFormato)
                        grdCargos.TextMatrix(grdCargos.Row, 11) = FormatCurrency(rsDatosFactura!Importe - rsDatosFactura!Descuento + rsDatosFactura!IVA, 15)
                        grdCargos.TextMatrix(grdCargos.Row, 13) = IIf(rsDatosFactura!BITPESOS = 1, 1, rsDatosFactura!TipoCambio)
                    End If
                    rsDatosFactura.MoveNext
                Loop
            End If
            
            pConfiguraGridCargos

            freDetalleNota.Enabled = True
            lblUsoCFDI.Enabled = True
            lblMetodoPago.Enabled = True
            lblFormaPago.Enabled = True
            
            cboUsoCFDI.Enabled = True
            cboMetodoPago.Enabled = True
            cboFormaPago.Enabled = True

            txtComentario.Enabled = True
                        
            vlcentavo = False
            txtCantidadCargo.Text = FormatCurrency(fTotalFactura(grdCargos, vgstrParametrosSP), 2)
            
            DoEvents
            'If fblnCanFocus(cboFacturasPaciente) Then cboFacturasPaciente.SetFocus
        Else
            txtCuenta.Text = ""
            txtNombrePaciente.Text = ""
            lblFechaFactura.Caption = ""
        End If

        rsDatosFactura.Close
    Else
        If txtCveCliente.Text <> "" And txtCliente.Text <> "" Then 'And txtRFC.Text <> "" And txtDomicilio.Text <> "" Then
            MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
             
            chkFacturasPaciente.Value = False
            pEnfocaTextBox txtCveCliente
            lbBuscaCargos.Visible = True
            txtBuscaCargo.Visible = True
            cboFacturasPaciente.Visible = False
            lbFacturasPaciente.Visible = False
            optCliente.Enabled = True
            chkPorcentajePaciente.Visible = False
            pHabilitaCuadros True
            
            '- Agregado para caso 7374 -'
            If optPaciente.Value And OptMotivoNota(0).Value Then
                OptMotivoNota(1).Value = True
                chkFacturasPaciente.Enabled = True
            End If
            '---------------------------'
        Else
            txtCveCliente.Text = ""
            txtCliente.Text = ""
            txtRFC.Text = ""
            txtDomicilio.Text = ""
            pEnfocaTextBox txtCveCliente
        
            lbBuscaCargos.Visible = True
            txtBuscaCargo.Visible = True
    
            cboFacturasPaciente.Visible = False
            lbFacturasPaciente.Visible = False
                        
            chkFacturasPaciente.Value = 0
        
            optCliente.Enabled = True
        
            chkPorcentajePaciente.Visible = False
        End If
    End If
    grdCargos.Redraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFacturasPaciente_Click"))
    Unload Me
End Sub

Private Sub cboFacturasPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = 13 Then
        grdCargos.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFacturasPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub cboFormaPago_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        cboMetodoPago.SetFocus
    End If
End Sub

Private Sub cboMetodoPago_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        cboUsoCFDI.SetFocus
    End If
End Sub

Private Sub cboUsoCFDI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        If optCliente.Value Then
            If cboCreditosDirectos.ListCount <> 0 Then
                If sstFacturasCreditos.TabEnabled(0) Then
                    sstFacturasCreditos.Tab = 0
                    cboFactura.SetFocus
                Else
                    sstFacturasCreditos.Tab = 1
                    cboCreditosDirectos.SetFocus
                End If
            Else
                sstFacturasCreditos.TabEnabled(1) = False
                If cboFactura.Enabled = True Then cboFactura.SetFocus
            End If
        Else
            If txtBuscaCargo.Visible = True Then txtBuscaCargo.SetFocus
            If fblnCanFocus(cboFacturasPaciente) Then cboFacturasPaciente.SetFocus
        End If
    End If
End Sub

Private Sub chkFacturasPaciente_Click()
On Error GoTo NotificaError

    If chkFacturasPaciente.Value = 1 Then
        lbBuscaCargos.Visible = False
        txtBuscaCargo.Visible = False
        chkPorcentajePaciente.Visible = True
        cboFacturasPaciente.Visible = True
        lbFacturasPaciente.Visible = True
        
        lbCantidadCargo.Caption = "Total nota"
        fraCargos.Caption = "Detalle de la factura"
        grdCargos.ToolTipText = "Detalle de la factura"
        txtCantidadCargo.ToolTipText = "Cantidad total de la nota"
        
        pLimpiaGridNota
        pConfiguraGridNota "FA"

        pFacturasPaciente
        grdCargos.Redraw = False
        
        cmdIncluirCargo.Enabled = False
        cboFacturasPaciente_Click
        grdCargos.Redraw = True
     Else
        lbBuscaCargos.Visible = True
        txtBuscaCargo.Visible = True
        cboFacturasPaciente.Visible = False
        lbFacturasPaciente.Visible = False
        chkPorcentajePaciente.Visible = False
        
        lbCantidadCargo.Caption = "Cantidad"
        fraCargos.Caption = "Cargos"
        grdCargos.ToolTipText = "Detalle de los cargos"
        txtCantidadCargo.ToolTipText = "Cantidad del concepto"
        txtCantidadCargo.Text = ""
        
        pLimpiaGridNota
        pConfiguraGridNota "FA"
        
        grdCargos.Redraw = False
'        pLimpiaGridCargos
'        pConfiguraGridCargos
        tmrCargos.Enabled = True
        cmdIncluirCargo.Enabled = True
        grdCargos.Redraw = True
        
        'If fblnCanFocus(txtBuscaCargo) Then
        'txtBuscaCargo.SetFocus
        'End If
       'If txtCveCliente.Text = "" Then
      ' txtCveCliente.SetFocus
       'End If
    End If
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFacturasPaciente_Click"))
    Unload Me
End Sub

Private Sub chkFacturasPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = 13 Then
        If txtCantidadCargo.Enabled = True Then txtCantidadCargo.SetFocus
        pSelTextBox txtCantidadCargo
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFacturasPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub chkFacturasPagadas_Click()
On Error GoTo NotificaError

    pLimpiaConcepto
    pLimpiaNota
    
    fraConcepto.Enabled = True
    freDetalleNota.Enabled = True
    lblUsoCFDI.Enabled = True
    lblMetodoPago.Enabled = True
    lblFormaPago.Enabled = True
    
    cboUsoCFDI.Enabled = True
    cboMetodoPago.Enabled = True
    cboFormaPago.Enabled = True
    
    If chkFacturasPagadas.Value = 1 Then
        sstFacturasCreditos.TabEnabled(2) = False
    Else
        sstFacturasCreditos.TabEnabled(2) = True
    End If
    
    pFacturas
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFacturasPagadas_Click"))
    Unload Me
End Sub

Private Sub chkFacturasPagadas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = 13 Then txtCantidad.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkFacturasPagadas_KeyDown"))
    Unload Me
End Sub

Private Sub chkPorcentaje_Click()
On Error GoTo NotificaError

    If chkPorcentaje.Value = 1 Then
        lbCantidad.Caption = "Total porcentaje"
        txtCantidad = Val(Format(txtCantidad, "##########.00")) / fTotalFactura(grdFactura, Trim(cboFactura.List(cboFactura.ListIndex))) * 100
        txtCantidad.ToolTipText = "Total porcentaje aplicado a la nota"
    Else
        If lbCantidad.Caption = "Total porcentaje" Then
            txtCantidad = (Val(Format(txtCantidad, "##########.00")) / 100) * fTotalFactura(grdFactura, Trim(cboFactura.List(cboFactura.ListIndex)))
        End If
        
        lbCantidad.Caption = "Total nota"
        txtCantidad.ToolTipText = "Cantidad total de la nota"
    End If
        
    fraConcepto.Enabled = True
    freDetalleNota.Enabled = True
    lblUsoCFDI.Enabled = True
    lblMetodoPago.Enabled = True
    lblFormaPago.Enabled = True
    
    cboUsoCFDI.Enabled = True
    cboMetodoPago.Enabled = True
    cboFormaPago.Enabled = True
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkPorcentaje_Click"))
    Unload Me
End Sub

Private Sub chkPorcentajePaciente_Click()
On Error GoTo NotificaError

    If chkPorcentajePaciente.Value = 1 Then
        lbCantidadCargo.Caption = "Porcentaje sobre importe de la factura"
        txtCantidadCargo = Val(Format(txtCantidadCargo, "##########.00")) / fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))) * 100
        txtCantidadCargo.Text = Format(txtCantidadCargo.Text, "Fixed")
        txtCantidadCargo.ToolTipText = "Total porcentaje aplicado a la factura"
    Else
        If lbCantidadCargo.Caption = "Porcentaje sobre importe de la factura" Then
          ' txtCantidadCargo = Format((Val(Format(txtCantidadCargo, "##########.00")) / 100) * fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))), 2)
          txtCantidadCargo = fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)))
        End If
        
        lbCantidadCargo.Caption = "Total nota"
        txtCantidadCargo.ToolTipText = "Cantidad total de la nota"
        txtCantidadCargo = FormatCurrency(txtCantidadCargo, 2)
    End If
    
    If vlblnlimpiaNotas Then
        pLimpiaNota
    End If
    
    fraConcepto.Enabled = True
    freDetalleNota.Enabled = True
    lblUsoCFDI.Enabled = True
    lblMetodoPago.Enabled = True
    lblFormaPago.Enabled = True
    
    cboUsoCFDI.Enabled = True
    cboMetodoPago.Enabled = True
    cboFormaPago.Enabled = True
    
    If chkPorcentajePaciente.Value = 1 Then
        If CDbl(txtCantidadCargo.Text) > 100 Then
            txtCantidadCargo.Text = 100
        Else
            If txtCantidadCargo.Enabled And Trim(txtCantidadCargo.Text) <> "" Then txtCantidadCargo.Text = Replace(txtCantidadCargo.Text, "$", "")
            txtCantidadCargo.Text = FormatCurrency(txtCantidadCargo.Text, 2)
            If chkPorcentaje.Visible = True Then
                vldbltotal = fTotalFactura(grdCargos, vgstrParametrosSP)
                If CDbl(txtCantidadCargo.Text) > vldbltotal Then
                    txtCantidadCargo.Text = FormatCurrency(vldbltotal, 2)
                End If
            End If
        End If
    End If
          
     
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkPorcentajePaciente_Click"))
    Unload Me
End Sub

Private Sub chkPorcentajePaciente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = 13 Then
        txtCantidadCargo.SetFocus
        pSelTextBox txtCantidadCargo
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkPorcentajePaciente_KeyDown"))
    Unload Me
End Sub

Private Sub cmdAddenda_Click()
    With frmAddendaDatos
        .lngAddenda = vglngCveAddenda
        .lngCuenta = vglngCuentaPacienteAddenda
        .strTipoIngreso = vgstrTipoPacienteAddenda
        .lngCveEmpresaContable = vgintClaveEmpresaContable
        .lngCveEmpresaPaciente = vglngCveEmpresaCliente
        .Show vbModal
    End With
End Sub
Private Sub cmdCancelaNotasSAT_Click()
    'Cancelacion masiva de facturas ante el SAT, cancelacion del XML
    Dim vlLngCantidadFacturas As Long
    Dim vlintFacturasCanceladas As Long
    Dim vlLngCont As Long
    Dim vllngPersonaGraba As Long
    Dim vlblnCancelarSiHO As Boolean

    On Error GoTo NotificaError
    '|  Los comprobantes seleccionados serán validados nuevamente ante el SAT.
    '|  ¿Desea continuar?
    If MsgBox(SIHOMsg(1249), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
            frmMotivosCancelacion.blnActivaUUID = False
            frmMotivosCancelacion.Show vbModal, Me
                If vgMotivoCancelacion = "" Then Exit Sub
        'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
        With grdBusqueda
            vlLngCantidadFacturas = 0
            vlintFacturasCanceladas = 0
            For vlLngCont = 1 To .Rows - 1
                '|  Parámetros:  Columna fixed         Estado nuevo esquema cancelación                     Estado nuevo esquema cancelación
                If .TextMatrix(vlLngCont, 0) = "*" And (.TextMatrix(vlLngCont, vlintColPFacSAT) <> "NP" And .TextMatrix(vlLngCont, vlintColPFacSAT) <> "CR") Then
                    vlLngCantidadFacturas = vlLngCantidadFacturas + 1
                    If .TextMatrix(vlLngCont, vlintColPFacSAT) = "PC" Then
                        vlblnCancelarSiHO = False
                    Else
                        vlblnCancelarSiHO = True
                    End If
                    frmFechaCancelacionNotas.pCancelaNota grdBusqueda.TextMatrix(vlLngCont, vlintColFolioNota), Trim(grdBusqueda.TextMatrix(vlLngCont, vlintColNombreCliente)), "ACTUAL", vlblnCancelarSiHO, False, True, vllngPersonaGraba
                End If
            Next vlLngCont
        End With
        If vlLngCantidadFacturas = vlintFacturasCanceladas Then
            '|  La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
        End If
        '|  Refresca la información
        pCargaNotas
        grdBusqueda.SetFocus
    End If
    Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelaFacturasSAT_Click"))
    Unload Me
    
    
End Sub

Private Sub cmdCargarDatos_Click()
On Error GoTo NotificaError

    pCargaNotas
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargarDatos_Click"))
    Unload Me
End Sub

Private Sub cmdComprobante_Click()
On Error GoTo NotificaError
    
    If vlstrTipoCFD = "CFD" Then
        frmComprobanteFiscalDigital.lngComprobante = lngIDnota
        frmComprobanteFiscalDigital.strTipoComprobante = IIf(optNotaCargo.Value, "CA", "CR")
        frmComprobanteFiscalDigital.blnCancelado = vgblnCancelada
        frmComprobanteFiscalDigital.Show vbModal, Me
    ElseIf vlstrTipoCFD = "CFDi" Then
        frmComprobanteFiscalDigitalInternet.lngComprobante = lngIDnota
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = IIf(optNotaCargo.Value, "CA", "CR")
        frmComprobanteFiscalDigitalInternet.blnCancelado = vgblnCancelada
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdComprobante_Click"))
    Unload Me
End Sub

Private Sub cmdConfirmartimbre_Click()
Dim vlLngCont As Long
Dim vllngPersonaGraba As Long
Dim vllngconsecutivoNota As Long
Dim chrTipoCancel As String
Dim ObjRs As New ADODB.Recordset
Dim vlngReg As Long

On Error GoTo NotificaError

blnNOMensajeErrorPAC = False 'de inicio siempre a False
vllngconsecutivoNota = CLng(frmNotas.grdBusqueda.TextMatrix(frmNotas.grdBusqueda.Row, vlintColIdNota))
'Los comprobantes seleccionados se encuentran pendientes de timbre fiscal. ¿Desea confirmar el timbre fiscal?
If MsgBox(Replace(SIHOMsg(1310), "Los comprobantes seleccionados se encuentran pendientes de timbre fiscal. ", ""), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   
   vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
   If vllngPersonaGraba = 0 Then Exit Sub
        
       pgbBarraCFD.Value = 35
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbHourglass
       lblTextoBarraCFD.Caption = "Confirmando timbre fiscal para la nota, por favor espere..."
       freBarraCFD.Visible = True
       freBarraCFD.Refresh
       frmNotas.Enabled = False
       freBarraCFD.Refresh
       pLogTimbrado 2
       
       blnNOMensajeErrorPAC = True
       EntornoSIHO.ConeccionSIHO.BeginTrans
       
       vlstrSentencia = "select gcf.VCHCODIGOPOSTALDFRECEPTOR, gcf.VCHREGIMENFISCALRECEPTOR from ccnota ccn " & _
                        " left join gncomprobantefiscaldigital gcf on (gcf.VCHSERIECOMPROBANTE || gcf.VCHFOLIOCOMPROBANTE) = Trim(ccn.VCHFACTURAIMPRESION)" & _
                        " where ccn.INTCONSECUTIVO = " & vllngconsecutivoNota & " and rownum = 1 and gcf.CHRTIPOCOMPROBANTE = 'FA' "
       Set ObjRs = frsRegresaRs(vlstrSentencia)
       
       If ObjRs.RecordCount > 0 Then
               
        With ObjRs
            If Trim(vlstrCodigo) = "" Then
                vlstrCodigo = !VCHCODIGOPOSTALDFRECEPTOR
            End If
            If vlstrRegimen = "" Then
                vlstrRegimen = !VCHREGIMENFISCALRECEPTOR
            End If
        End With
              
       End If
       
       ObjRs.Close
                                                                       
       vlngReg = flngRegistroFolio(IIf(optNotaCargo.Value, "CA", "CR"), vllngconsecutivoNota)
       
        If vgstrVersionCFDI = "4.0" Then
            vgstrTipoNotaSig = IIf(optNotaCredito.Value = True, "CR", "CA")
            vgstrCodigoPostalSig = Trim(vlstrCodigo)
            vglngIdComprobanteSig = vllngconsecutivoNota
            vgstrRegimenFiscalSig = vlstrRegimen
        End If
       
       If Not fblnGeneraComprobanteDigital(vllngconsecutivoNota, IIf(optNotaCargo.Value, "CA", "CR"), 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
           On Error Resume Next
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          
            vgstrTipoNotaSig = ""
            vgstrCodigoPostalSig = ""
            vglngIdComprobanteSig = 0
            vgstrRegimenFiscalSig = ""
                    
          If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar
             'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
              MsgBox Replace(SIHOMsg(1314), "factura <FOLIO>", "nota de " & IIf(optNotaCargo.Value, "cargo", "crédito")), vbInformation + vbOKOnly, "Mensaje"
             'la factura se queda igual, no se hace nada
          ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
             'No es posible realizar el timbrado de la nota <FOLIO>, la nota será cancelada.
              MsgBox Replace(Replace(SIHOMsg(1313), "factura <FOLIO>", "nota " & Trim(lblFolio.Caption)), "factura", " nota"), vbExclamation + vbOKOnly, "Mensaje"
  
              blnCancelaNota = True
              'Aqui se debe de cancelar la nota__________________________________
              'cargamos la nota(0)_______________________________________________
              'Se revisa el parametro del tipo de cancelación de documento
              If chrTipoCancel = "" Then
                 vlstrSentencia = "select chrTipoCancel from ccTipoCancelacion where smiCveDepartamento = " & vgintNumeroDepartamento & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
                 Set ObjRs = frsRegresaRs(vlstrSentencia)
                 If ObjRs.RecordCount > 0 Then 'Si tiene configurado el parametro seleccionar la opción indicada
                    chrTipoCancel = Trim(ObjRs!chrTipoCancel)
                 Else
                    chrTipoCancel = "DOCUMENTO"
                 End If
              End If
                          
              If chrTipoCancel = "ELEGIR" Then
                 frmFechaCancelacionNotas.Show vbModal
              Else
                 frmFechaCancelacionNotas.pCancelaNota grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColFolioNota), txtCliente.Text, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColFolioNota), True, True, True, vllngPersonaGraba
              End If '___________________________________________________________
              
              If blnCancelaNota = True Then
                 'Eliminamos la informacion de la nota de la tabla de pendientes de timbre fiscal
                 pEliminaPendientesTimbre vllngconsecutivoNota, IIf(optNotaCargo.Value, "CA", "CR")
              End If
          End If
       Else
       
       
            vgstrTipoNotaSig = ""
            vgstrCodigoPostalSig = ""
            vglngIdComprobanteSig = 0
            vgstrRegimenFiscalSig = ""
            
            vlstrCodigo = ""
            vlstrRegimen = ""
            vlRazonSocialComprobante = ""
            vlRFCComprobante = ""
       
          'Se guarda el LOG
           Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & lblFolio.Caption)
          'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
           pEliminaPendientesTimbre vllngconsecutivoNota, IIf(optNotaCargo.Value, "CA", "CR")
          'Commit
                    
           EntornoSIHO.ConeccionSIHO.CommitTrans
          'Timbre fiscal de factura <FOLIO>: Confirmado.
           MsgBox Replace(SIHOMsg(1315), "factura <FOLIO> ", "nota de " & IIf(optNotaCargo.Value, "cargo ", "crédito ")), vbInformation + vbOKOnly, "Mensaje"
       End If

       'Barra de progreso CFD
       pgbBarraCFD.Value = 100
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbDefault
       freBarraCFD.Visible = False
       frmNotas.Enabled = True
       pLogTimbrado 1
       
       If vgIntBanderaTImbradoPendiente <> 1 Then
          InicializaComponentes
          If optNotaCargo.Value Then
             optNotaCargo.SetFocus
          Else
             optNotaCredito.SetFocus
          End If
       End If
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub
Private Sub cmdconfirmartimbrefiscal_Click()
Dim vlLngCont As Integer
Dim vllngPersonaGraba As Long
Dim Idnota As Long
Dim chrTipoCancel As String
Dim ObjRs As New ADODB.Recordset
Dim vlngReg As Long

On Error GoTo NotificaError

blnNOMensajeErrorPAC = False 'de inicio siempre a False
chrTipoCancel = ""
'Los comprobantes seleccionados se encuentran pendientes de timbre fiscal ¿Desea confirmar el timbre fiscal?
If MsgBox(SIHOMsg(1310), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   
   vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
   If vllngPersonaGraba = 0 Then Exit Sub
     
     'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
      With grdBusqueda
           For vlLngCont = 1 To .Rows - 1
               If .TextMatrix(vlLngCont, 0) = "*" And .TextMatrix(vlLngCont, vlintcolPTimbre) = 1 Then
                                                           
                  Idnota = CLng(.TextMatrix(vlLngCont, vlintColIdNota))
                                                           
                  blnNOMensajeErrorPAC = True
                  EntornoSIHO.ConeccionSIHO.BeginTrans
                  
                  If .TextMatrix(vlLngCont, vlintColPFacSAT) = "XX" Then
                     pLogTimbrado 2
                     'es una factura que esta pendiente de timbre fiscal pero se debe de cancelar después de la confirmación del timbre
                     vgIntBanderaTImbradoPendiente = 0
                     vlngReg = flngRegistroFolio("CR", Idnota)
                     If Not fblnGeneraCFDpCancelacion(Idnota, "CR", fblnTCFDi(vlngReg), 0) Then
                        On Error Resume Next
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        
                        pLogTimbrado 1
                        
                        If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar/o que no alcanzó a llegar el timbrado
                           'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                           MsgBox Replace(SIHOMsg(1314), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), vbInformation + vbOKOnly, "Mensaje"
                           
                        ElseIf vgIntBanderaTImbradoPendiente = 2 Then
                          'No es posible realizar el timbrado de la nota <FOLIO>, la nota será cancelada.
                          MsgBox Replace(Replace(SIHOMsg(1313), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), "factura", "nota"), vbExclamation + vbOKOnly, "Mensaje"
                          'Aqui se debe de cancelar la nota
                          pEjecutaSentencia " UPDATE PVCANCELARCOMPROBANTES SET BITPENDIENTECANCELAR = 0 " & _
                         "Where INTCOMPROBANTE =" & Idnota & " And VCHTIPOCOMPROBANTE = 'CR'"
                                                    
                          pEliminaPendientesTimbre Idnota, "CR"
                          
                        End If
                     Else  'se confirmo el timbre correctamente
                           pLogTimbrado 1
                           'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                           pEliminaPendientesTimbre Idnota, "CR"
                           EntornoSIHO.ConeccionSIHO.CommitTrans
                           'Timbre fiscal de la factura <FOLIO>: Confirmado.
                           MsgBox Replace(SIHOMsg(1315), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), vbInformation + vbOKOnly, "Mensaje"
                     End If
                  
                  Else
                     pLogTimbrado 2
                     vlngReg = flngRegistroFolio(Trim(.TextMatrix(vlLngCont, vlintColchrTipo)), Idnota)
                     
                     If Not fblnGeneraComprobanteDigital(Idnota, Trim(.TextMatrix(vlLngCont, vlintColchrTipo)), 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
                       On Error Resume Next
                       
                       EntornoSIHO.ConeccionSIHO.RollbackTrans
                                              
                       'guardamos log del timbrado
                       pLogTimbrado 1
                       
                       If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar
                          'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                          MsgBox Replace(SIHOMsg(1314), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), vbInformation + vbOKOnly, "Mensaje"
                          'la factura se queda igual, no se hace nada
                       ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
                          'No es posible realizar el timbrado de la nota <FOLIO>, la nota será cancelada.
                          MsgBox Replace(Replace(SIHOMsg(1313), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), "factura", "nota"), vbExclamation + vbOKOnly, "Mensaje"
                          blnCancelaNota = True
                          'Aqui se debe de cancelar la nota__________________________________
                          'cargamos la nota(0)_______________________________________________
                          .Row = vlLngCont
                          pMostrarNota vlLngCont, .TextMatrix(vlLngCont, vlintColTipoNotaFACR)
                          'Se revisa el parametro del tipo de cancelación de documento
                          If chrTipoCancel = "" Then
                             vlstrSentencia = "select chrTipoCancel from ccTipoCancelacion where smiCveDepartamento = " & vgintNumeroDepartamento & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
                             Set ObjRs = frsRegresaRs(vlstrSentencia)
                             If ObjRs.RecordCount > 0 Then 'Si tiene configurado el parametro seleccionar la opción indicada
                                chrTipoCancel = Trim(ObjRs!chrTipoCancel)
                             Else
                                chrTipoCancel = "DOCUMENTO"
                             End If
                          End If
                          
                          If chrTipoCancel = "ELEGIR" Then
                             frmFechaCancelacionNotas.Show vbModal
                          Else
                             frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio.Caption), txtCliente.Text, chrTipoCancel, True, True, True, vllngPersonaGraba
                          End If '___________________________________________________________
                          
                          If blnCancelaNota = True Then
                            'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                            pEliminaPendientesTimbre Idnota, Trim(.TextMatrix(vlLngCont, vlintColchrTipo))
                          End If
                       End If
                     Else
                      'Se guarda el LOG
                       Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & .TextMatrix(vlLngCont, 1))
                       'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                       pEliminaPendientesTimbre Idnota, Trim(.TextMatrix(vlLngCont, vlintColchrTipo))
                       'Commit
                       
                       'guardamos el log del timbrado
                       pLogTimbrado 1
                       
                       EntornoSIHO.ConeccionSIHO.CommitTrans
                      'Timbre fiscal de factura <FOLIO>: Confirmado.
                       MsgBox Replace(SIHOMsg(1315), "factura <FOLIO>", "nota " & Trim(.TextMatrix(vlLngCont, vlintColFolioNota))), vbInformation + vbOKOnly, "Mensaje"
                     End If
                  End If
               End If
           
           Next vlLngCont
      End With
      
      blnNOMensajeErrorPAC = False
      pCargaNotas
      grdBusqueda.SetFocus
      
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub

Private Sub cmdDesglose_Click()
    Dim rptReporte As CRAXDRT.report
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(8) As String
    Dim strParametros As String
    
    strParametros = Str(lngIDnota) & "|" & IIf(grdNotas.TextMatrix(grdNotas.Row, vlintColTipoNotaFARCDetalle) = "FA", 1, 2)
    Set rsReporte = frsEjecuta_SP(strParametros, "Sp_CcRptDesgloseNotaFactura")
    If rsReporte.RecordCount <> 0 Then
        pInstanciaReporte rptReporte, "rptDesgloseNotaFactura.rpt"
        rptReporte.DiscardSavedData
        
        alstrParametros(0) = "NombreHospital" & ";" & Trim(vgstrNombreHospitalCH) & ";TRUE"
        alstrParametros(1) = "FolioNota" & ";" & Trim(lblFolio.Caption) & ";TRUE"
        alstrParametros(2) = "FechaNota" & ";" & Trim(mskFecha.Text) & ";TRUE"
        alstrParametros(3) = "TipoNota" & ";" & Trim(IIf(optNotaCargo.Value, "Nota de cargo", "Nota de crédito")) & ";TRUE"
        alstrParametros(4) = "NombreCliente" & ";" & Trim(txtCliente.Text) & ";TRUE"
        alstrParametros(5) = "DireccionCliente" & ";" & Trim(txtDomicilio.Text) & ";TRUE"
        alstrParametros(6) = "RFCCliente" & ";" & Trim(txtRFC.Text) & ";TRUE"
        alstrParametros(7) = "FacturaCR" & ";" & IIf(grdNotas.TextMatrix(grdNotas.Row, vlintColTipoNotaFARCDetalle) = "FA", 1, 2) & ";TRUE"
        alstrParametros(8) = "Cliente" & ";" & IIf(lbCliente.Caption = "Cliente", "Cliente", "Paciente") & ";TRUE"
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, "P", "Desglose de facturas de la nota"
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    rsReporte.Close
End Sub

Private Sub cmdIncluirCargo_Click()
    Dim vldblCantidad As Double
    Dim vldblIVA As Double
    Dim rsCveConcepto As ADODB.Recordset
    Dim strParametros As String
    Dim dblCantidad As Double
    Dim dblDescuento As Double
    
    Dim vldblDescuento As Double
    Dim vldblSumaImportes As Double
    Dim vldblSumaDescuentos As Double
    Dim vldblIVACorrespondiente As Double
    Dim dblTotalNota As Double

On Error GoTo NotificaError

    ' Validar facturas pagadas
    If chkFacturasPaciente.Value = 0 Then
        If fblnExisteCargo() Then Exit Sub '(CR) - Agregado para caso 6864
    
        dblCantidad = CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 3) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 3))) * CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 4) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 4)))
        dblDescuento = CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 5) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 5)))
                   
        If (dblCantidad - dblDescuento) < CDbl(Val(Format(txtCantidadCargo.Text, vlstrFormato))) Then
        'El importe de la nota de crédito no puede ser mayor al importe del cargo!
            MsgBox SIHOMsg(1000), vbOKOnly + vbInformation, "Mensaje"
            txtCantidadCargo.Text = ""
            txtCantidadCargo.SetFocus
            Exit Sub
        End If
            
        If Val(Format(txtCantidadCargo.Text, vlstrFormato)) = 0 Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            If txtCantidadCargo.Enabled = True Then txtCantidadCargo.SetFocus
        Else
            vldblCantidad = CDbl(Val(Format(txtCantidadCargo.Text, vlstrFormato)))
            strParametros = CStr(vgintClaveEmpresaContable) & "|" & grdCargos.TextMatrix(grdCargos.Row, 10)
            Set rsCveConcepto = frsEjecuta_SP(strParametros, "SP_PVSELCVECONCEPTO")
            If rsCveConcepto.RecordCount > 0 Then
                If optNotaCredito.Value = True And OptMotivoNota(1).Value = True Then
                    vllngCveCtaDescuento = rsCveConcepto!NumCtaDescNota
                Else
                    vllngCveCtaDescuento = rsCveConcepto!intCuentaDescuento
                End If
                vllngCveCtaIngresos = rsCveConcepto!INTCUENTACONTABLE
                vldblPorcentajeIVA = rsCveConcepto!smyIVA / 100
            Else
                vllngCveCtaDescuento = 0
                vllngCveCtaIngresos = 0
                vldblPorcentajeIVA = 0
            End If
            rsCveConcepto.Close
            
            If (dblCantidad - dblDescuento) = CDbl(Val(Format(txtCantidadCargo.Text, vlstrFormato))) Then
                vldblIVA = CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 6) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 6)))
            Else
                If CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 6) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 6))) > 0 Then
                    vldblIVA = (vldblCantidad) * (CDbl(IIf(grdCargos.TextMatrix(grdCargos.Row, 6) = "", 0, grdCargos.TextMatrix(grdCargos.Row, 6))) / (dblCantidad - dblDescuento))
                Else
                    vldblIVA = 0
                End If
            End If
            
            pIncluyeConcepto grdCargos.TextMatrix(grdCargos.Row, 2), vldblCantidad, 0, 0, vldblIVA, _
                             grdCargos.TextMatrix(grdCargos.Row, 10), vllngCveCtaIngresos, vllngCveCtaDescuento, _
                             glngCtaIVACobrado, "CF", 1, vlblnSeleccionoOtroConcepto, grdCargos.TextMatrix(grdCargos.Row, 11), grdCargos.TextMatrix(grdCargos.Row, 12)
            
            pAcumulaFactura cboCreditosDirectos.List(cboCreditosDirectos.ListIndex), vldblCantidad, 0, vldblIVA, 1, 1, CDate(grdCargos.TextMatrix(grdCargos.Row, 1))
            
            freDetalleNota.Enabled = True
            txtComentario.Enabled = True
            lblUsoCFDI.Enabled = True
            lblMetodoPago.Enabled = True
            lblFormaPago.Enabled = True
            
            cboUsoCFDI.Enabled = True
            cboMetodoPago.Enabled = True
            cboFormaPago.Enabled = True
            
            txtCantidadCargo.Text = ""
            strPacienteImpresion = Trim(txtCliente.Text)
        End If
    Else
        ' Facturas pagadas
        If Val(Format(txtCantidadCargo.Text, vlstrFormato)) = 0 Or (Val(Format(txtCantidadCargo.Text, vlstrFormato)) = 0 And Val(Format(txtDescuento.Text, vlstrFormato)) = 0) Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            txtCantidadCargo.SetFocus
        Else
            pSeleccionaElementoFactura
            
            If Not fblnExisteConcepto() Then
                strPacienteImpresion = Trim(txtNombrePaciente.Text)
                strFacturaImpresion = Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))
                
                If chkPorcentajePaciente.Value = 0 Then
                    dblTotalNota = Val(Format(txtCantidadCargo, "##########.00"))
                Else
                    dblTotalNota = (Val(Format(txtCantidadCargo, "##########.00")) / 100) * fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)))
                End If
            
                vldblCantidad = Format(dblTotalNota, vlstrFormato)
                
                '------------------------------------'
                '   N O T A S   D E   C R É D I T O  '
                '------------------------------------'
                'If Round((vldblImporteConcepto - vldblDescuentoConcepto), 2) <= Round(vldblCantidad, 2) Then
                If Format((vldblImporteConcepto - vldblDescuentoConcepto), vlstrFormato) <= vldblCantidad Then
                    If OptMotivoNota(0).Value = True Then
                        vldblDescuento = grdCargos.TextMatrix(grdCargos.Row, 3)
                        vldblCantidad = grdCargos.TextMatrix(grdCargos.Row, 2)
                    Else
                        vldblDescuento = 0
                        vldblCantidad = grdCargos.TextMatrix(grdCargos.Row, 2) - grdCargos.TextMatrix(grdCargos.Row, 3)
                    End If
                    
                    vldblIVA = vldblIvaConcepto
                    vldblIVA = grdCargos.TextMatrix(grdCargos.Row, 4)
    
                    pIncluyeConcepto vlStrConcepto, vldblCantidad, vldblDescuento, 0, vldblIVA, vlngCveConcepto, vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, "CF", grdCargos.TextMatrix(grdCargos.Row, 13), vlblnSeleccionoOtroConcepto
                    If optPaciente.Value = True And chkFacturasPaciente.Value = 1 Then
                        pAcumulaFactura cboFacturasPaciente.List(cboFacturasPaciente.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, 1, lblFechaFactura.Caption
                    Else
                        pAcumulaFactura cboFacturasPaciente.List(cboFacturasPaciente.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, 1, CDate(grdCargos.TextMatrix(grdCargos.Row, 1))
                    End If
                    grdCargos.Enabled = True
                    sstFacturasCreditos.TabEnabled(1) = False
                    grdCargos.SetFocus
                                   
                    pProrratearNota dblTotalNota
                Else
                    If (vldblImporteConcepto - vldblDescuentoConcepto) < vldblCantidad Then
                        'La cantidad no puede se mayor al importe del concepto.
                        MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                        txtCantidadCargo.SetFocus
                    Else
                        If (vldblDescuentoConcepto + vldblDescuento) > (vldblImporteConcepto - vldblCantidad) Then
                            'Los descuentos exceden el importe del concepto.
                            MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                            
                            If txtDescuento.Visible = True And txtDescuento.Enabled = True Then
                                txtDescuento.SetFocus
                            Else
                                txtCantidadCargo.SetFocus
                            End If
                        Else
                            
                            If OptMotivoNota(0).Value = True Then
                                vldblDescuento = grdCargos.TextMatrix(grdCargos.Row, 3)
                                vldblCantidad = grdCargos.TextMatrix(grdCargos.Row, 2)
                            Else
                                vldblDescuento = 0
                                vldblCantidad = grdCargos.TextMatrix(grdCargos.Row, 2) - grdCargos.TextMatrix(grdCargos.Row, 3)
                            End If
                            If vldblIvaConcepto > 0 Then
                                vldblIVA = (vldblCantidad - vldblDescuento) * (vldblIvaConcepto / (vldblImporteConcepto - vldblDescuentoConcepto))
                            Else
                                vldblIVA = 0
                            End If
                                                                
                            pIncluyeConcepto vlStrConcepto, vldblCantidad, vldblDescuento, 0, vldblIVA, vlngCveConcepto, vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, "CF", grdCargos.TextMatrix(grdCargos.Row, 13), vlblnSeleccionoOtroConcepto
                            pAcumulaFactura cboFacturasPaciente.List(cboFacturasPaciente.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, grdCargos.TextMatrix(grdCargos.Row, 13)
    
                            pProrratearNota dblTotalNota
    
                            grdCargos.SetFocus
                        End If
                    End If
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIncluirCargo_Click"))
    Unload Me
End Sub

Private Sub cmdIncluirCR_Click()
On Error GoTo NotificaError
    
    Dim vldblCantidad As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim vldblSumaImportes As Double
    Dim vldblSumaDescuentos As Double
    Dim vldblIVACorrespondiente As Double
    Dim dblAjustaIVA As Double
    Dim dblCantidadCredito As Double
    Dim vlDblTotalGrid As Double
    Dim vlDblTotalGrid0 As Double
    Dim vlIntContfor  As Integer
    
    vlIntContfor = 0
    vlDblTotalGrid = 0
    
    vldblIVAcero = 1
    vldblIVApago = 1

    If Val(Format(txtCantidadCR.Text, vlstrFormato)) = 0 And Val(Format(txtDescuentoCR.Text, vlstrFormato)) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtCantidadCR.SetFocus
    Else
        pSeleccionaElementoCR
'        If vllngCveCtaDescuento = 0 Then
'            MsgBox Replace(SIHOMsg(907), "contable", "contable de descuento") & Chr(13) & lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex) & "  " & lstListaSeleccionada.Text, vbOKOnly + vbInformation, "Mensaje"
'            Exit Sub
'        End If
'        If vllngCveCtaIngresos = 0 Then
'            MsgBox Replace(SIHOMsg(907), "contable", "contable de ingreso") & Chr(13) & lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex) & "  " & lstListaSeleccionada.Text, vbOKOnly + vbInformation, "Mensaje"
'            Exit Sub
'        End If
        If vllngCveCtaIva = 0 Then
            'No se encuentran registradas las cuentas de IVA cobrado y no cobrado en los parámetros generales del sistema.
            MsgBox SIHOMsg(729), vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
        End If
        
'        If vldblIVApago = 0 Then
            'La cantidad no puede se mayor al importe del concepto.
'            MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
'            Exit Sub
'        End If
        
       ' If vldblPorcentajeIVA = 0 And vldblIVAcero = 0 Then
        If vldblIVAcero = 0 Then
            'Porcentaje de IVA incorrecto.
            MsgBox Replace(SIHOMsg(400), "Porcentaje", "Porcentaje de IVA"), vbOKOnly + vbExclamation, "Mensaje"
            Exit Sub
         Else
            If vldblIVApago = 0 Then
                'La cantidad no puede se mayor al importe del concepto.
                MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
            End If
         End If
        
        If Not fblnExisteConceptoCR() Then
'            strPacienteImpresion = Trim(txtNombrePaciente.Text)
            'strFacturaImpresion = Trim(cboCreditosDirectos.List.List(cboCreditosDirectos.ListIndex))
'
            vldblCantidad = CDbl(Val(Format(txtCantidadCR.Text, vlstrFormato)))
            vldblDescuento = CDbl(Val(Format(txtDescuentoCR.Text, vlstrFormato)))
'
            If optNotaCredito.Value Then
            
                '------------------------------------'
                '   N O T A S   D E   C R É D I T O  '
                '------------------------------------'
                  For vlIntContfor = 1 To grdNotas.Rows - 1
                   If grdNotas.TextMatrix(vlIntContfor, 1) = IIf(sstFacturasCreditos.Tab = 0, cboFactura.List(cboFactura.ListIndex), cboCreditosDirectos.List(cboCreditosDirectos.ListIndex)) Then 'And grdNotas.TextMatrix(vlIntContfor, 5) <> 0 Then
                     If grdNotas.TextMatrix(vlIntContfor, 5) <> 0 Then
                       vlDblTotalGrid = vlDblTotalGrid + CDbl(grdNotas.TextMatrix(vlIntContfor, 3)) - CDbl(grdNotas.TextMatrix(vlIntContfor, 4)) + CDbl(grdNotas.TextMatrix(vlIntContfor, 5))
                     Else
                      vlDblTotalGrid0 = vlDblTotalGrid0 + CDbl(grdNotas.TextMatrix(vlIntContfor, 3)) - CDbl(grdNotas.TextMatrix(vlIntContfor, 4)) + CDbl(grdNotas.TextMatrix(vlIntContfor, 5))
                     End If
                   End If
                Next vlIntContfor
                
                 If (txtCantidadCR + vlDblTotalGrid - vldblDescuento) > vldblImporteIvaNoCero * (1 + vgdblCantidadIvaGeneral) And vldblPorcentajeIVA <> 0 Then
                  'La cantidad no puede se mayor al importe del concepto.
                    MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                    txtCantidadCR.SetFocus
                    Exit Sub
                Else
                 If (txtCantidadCR + vlDblTotalGrid0 - vldblDescuento) > vldblImporteIvaCero And vldblPorcentajeIVA = 0 Then
                 'La cantidad no puede se mayor al importe del concepto.
                    MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                    txtCantidadCR.SetFocus
                    Exit Sub
                 End If
                
                End If

                        
                   'If vldblImporteConcepto < (vldblCantidad + ((vldblCantidad - vldblDescuento) * vldblPorcentajeIVA)) + vlDblTotalGrid cantidad existente abajo en el grid ! Then
                If vldblImporteConcepto < Round(((vldblCantidad - vldblDescuento) + ((vldblCantidad - vldblDescuento) * vldblPorcentajeIVA) + vlDblTotalGrid + vlDblTotalGrid0), 2) Then
                    'La cantidad no puede se mayor al importe del concepto.
                    MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                    txtCantidadCR.SetFocus
                Else
                    If (vldblDescuentoConcepto + vldblDescuento) > (vldblImporteConcepto - vldblCantidad) Then
                        'Los descuentos exceden el importe del concepto.
                        MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                        If txtDescuentoCR.Visible = True Then
                            txtDescuentoCR.SetFocus
                        Else
                            txtCantidadCR.SetFocus
                        End If
                    Else
                        pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, 0, (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, 1, vlblnSeleccionoOtroConcepto
                        pAcumulaFactura cboCreditosDirectos.List(cboCreditosDirectos.ListIndex), vldblCantidad, vldblDescuento, (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA, 1, 1
                        
                        If blnPestaniaFac Then
                            sstFacturasCreditos.TabEnabled(0) = False
                        End If
                        plimpiaListaCR
                        grdCreditoDirecto.SetFocus
                    End If
                End If
            Else
                '--------------------------------'
                '   N O T A S   D E   C A R G O  '
                '--------------------------------'
                If (vldblDescuentoConcepto + vldblDescuento) > (vldblImporteConcepto + vldblCantidad) Then
                    'Los descuentos exceden el importe del concepto.
                    MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                    txtDescuentoCR.SetFocus
                Else
                    pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, 0, (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, 1, vlblnSeleccionoOtroConcepto
                    pAcumulaFactura cboCreditosDirectos.List(cboCreditosDirectos.ListIndex), vldblCantidad, vldblDescuento, (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA, 1, 1
                    grdCreditoDirecto.SetFocus
                End If
            End If
        Else
            plimpiaListaCR
            grdCreditoDirecto.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIncluirCR_Click"))
    Unload Me
End Sub


Private Sub plimpiaListaCR()
    txtCantidadCR.Text = ""
    txtDescuentoCR.Text = ""
    lstConceptosFact.ListIndex = 0
    cmdIncluirCR.Enabled = False
End Sub
Private Sub Command1_Click()
Dim vlintRenglon As Integer
    For vlintRenglon = 0 To UBound(aFacturas)
        MsgBox ("vlstrFolioFactura = " & aFacturas(vlintRenglon).vlstrFolioFactura & "  vldblSubtotal = " & aFacturas(vlintRenglon).vldblSubtotal & "  vldblIVA = " & aFacturas(vlintRenglon).vldblIVA)
    Next
End Sub
Private Sub cmdIncluirCR_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdIncluirCR.Enabled Then
            cmdIncluirCR.SetFocus
        Else
            txtCantidadCR.SetFocus
        End If
    End If
End Sub
Private Sub grdBusqueda_Click()
With grdBusqueda
    If .MouseCol = 0 And .MouseRow > 0 Then
       If .TextMatrix(.Row, vlintcolPTimbre) = "1" Then
           If .TextMatrix(.Row, 0) = "*" Then
                If vllngSeleccPendienteTimbre > 0 Then
                   vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre - 1
                End If
                .TextMatrix(.Row, 0) = ""
           Else
                vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                .TextMatrix(.Row, 0) = "*"
           End If
           Me.cmdConfirmarTimbreFiscal.Enabled = vllngSeleccPendienteTimbre > 0
       
       ElseIf .TextMatrix(.Row, vlintColPFacSAT) = "PC" Or .TextMatrix(.Row, vlintColPFacSAT) = "PA" Or .TextMatrix(.Row, vlintColPFacSAT) = "XX" Then
            If .TextMatrix(.Row, 0) = "*" Then
                If vllngSeleccionadas > 0 Then
                   vllngSeleccionadas = vllngSeleccionadas - 1
                End If
                .TextMatrix(.Row, 0) = ""
            Else
                vllngSeleccionadas = vllngSeleccionadas + 1
                .TextMatrix(.Row, 0) = "*"
            End If
       Me.cmdCancelaNotasSAT.Enabled = vllngSeleccionadas > 0
       
       End If
    End If
End With
End Sub

Private Sub GrdCargos_DblClick()

    Dim intCol As Integer
    Dim intRow As Integer
    
    Dim vlintContador As Integer
    Dim vblnBanderaEncontrado As Boolean

On Error GoTo NotificaError

    If chkFacturasPaciente.Value = 0 Then
        intRow = grdCargos.Row
        
        If optPaciente.Value = True Then pColorearRenglon (&O0)
        
        grdCargos.Row = intRow
        
        For intCol = 1 To grdCargos.Cols - 3
            grdCargos.Col = intCol
            grdCargos.CellForeColor = &HC0&
        Next intCol
           
        If txtCantidadCargo.Enabled = True Then
            If Val(txtCantidadCargo) > 100 Then
                MsgBox SIHOMsg(400), vbOKOnly + vbInformation, "Mensaje"
                txtCantidadCargo.SetFocus
                Exit Sub
            End If
            
            txtCantidadCargo.SetFocus
        End If
    Else
        vblnBanderaEncontrado = False

        If vlblnConsulta Then Exit Sub
    
        If Trim(grdCargos.TextMatrix(1, 1)) <> "" Then
            vlStrConcepto = grdCargos.TextMatrix(grdCargos.Row, 1)
            vldblImporteConcepto = CDbl(Format(grdCargos.TextMatrix(grdCargos.Row, 2), vlstrFormato))
            vldblDescuentoConcepto = CDbl(Format(grdCargos.TextMatrix(grdCargos.Row, 3), vlstrFormato))
            vldblIvaConcepto = CDbl(grdCargos.TextMatrix(grdCargos.Row, 4))
            vldblPorcentajeIVA = CDbl(Format(grdCargos.TextMatrix(grdCargos.Row, 10), vlstrFormato))
            vlngCveConcepto = CLng(Format(grdCargos.TextMatrix(grdCargos.Row, 6), vlstrFormato))
        End If
                           
        cmdIncluir.Enabled = False
        
        If txtCantidadCargo.Enabled = True Then
            If Val(txtCantidadCargo) > 100 Then
                MsgBox SIHOMsg(400), vbOKOnly + vbInformation, "Mensaje"
                txtCantidadCargo.SetFocus
                Exit Sub
            End If
            
            txtCantidadCargo.SetFocus
        End If
        
        cmdIncluirCargo_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_DblClick"))
    Unload Me
End Sub

Private Sub grdCargos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then GrdCargos_DblClick
End Sub

Private Sub grdCreditoDirecto_DblClick()
    lstConceptosFact.SetFocus
    lstConceptosFact.ListIndex = 0
End Sub

Private Sub grdCreditoDirecto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstConceptosFact.SetFocus
        lstConceptosFact.ListIndex = 0
    End If
End Sub

Private Sub lstConceptosFact_DblClick()
    cmdIncluirCR.Enabled = True
    vldblImporteConcepto = CDbl(Format(grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 4), vlstrFormato))
    txtCantidadCR.SetFocus
End Sub

Private Sub lstConceptosFact_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdIncluirCR.Enabled = True
        vldblImporteConcepto = CDbl(Format(grdCreditoDirecto.TextMatrix(grdCreditoDirecto.Row, 4), vlstrFormato))
        txtCantidadCR.SetFocus
    End If
End Sub

Private Sub lstConceptosFacturacion_Click()
    If Val(lstConceptosFacturacion.ItemData(lstConceptosFacturacion.ListIndex)) <> Val(grdFactura.TextMatrix(grdFactura.Row, 6)) Then
        vlblnSeleccionoOtroConcepto = True
    Else
        vlblnSeleccionoOtroConcepto = False
    End If
End Sub

Private Sub MskFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        fraDatosCliente.Enabled = True
        
        If fraMotivoNota.Enabled = True Then
            If optCliente.Value Then
                OptMotivoNota(0).SetFocus
            Else
                OptMotivoNota(1).SetFocus
            End If
        Else
            txtCveCliente.SetFocus
        End If
    End If
End Sub

Private Sub optCliente_Click()
    lbCliente.Caption = "Cliente"
    optNotaCargo.Enabled = True
    optNotaCredito.Value = True
    If optNotaCredito.Visible = True And optNotaCredito.Enabled Then optNotaCredito.SetFocus
    
    OptTipoPaciente(0).Visible = False
    OptTipoPaciente(1).Visible = False
    fraTipoPaciente.Visible = False
    sstFacturasCreditos.TabEnabled(1) = True
    
    sstFacturasCreditos.Visible = True
    sstFacturasCreditos.Height = 4745 '4935
    sstCargos.Visible = False
    
    grdNotas.ColWidth(vlintColFactura) = 1100
    
    fraMotivoNota.Enabled = True
    OptMotivoNota(0).Value = True
    
    fraDatosCliente.Top = 1450
    
    FraMetodoForma.Visible = True
    
    InicializaComponentes
End Sub

Private Sub OptCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        optNotaCredito.SetFocus
    End If
End Sub

Private Sub OptConsultaTipoPaciente_Click(Index As Integer)
    txtNumCliente.Text = ""
    txtNombreCliente.Text = ""
    
    If OptConsultaTipoPaciente(2).Value = True Then
        txtNumCliente.Enabled = False
    Else
        txtNumCliente.Enabled = True
    End If
End Sub

Private Sub OptConsultaTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtNumCliente.Enabled = True Then
            txtNumCliente.SetFocus
        Else
            cmdCargarDatos.SetFocus
        End If
    End If
End Sub

Private Sub optMostrarSolo_Click(Index As Integer)
    If optMostrarSolo(0).Value Then
        Me.mskFechaInicial.Enabled = True
        Me.mskFechaFinal.Enabled = True
        Me.txtNumCliente.Enabled = True
        Me.txtNombreCliente.Enabled = True
        Me.cmdCargarDatos.Enabled = True
        
        Me.OptConsultaTipoPaciente(0).Enabled = True
        Me.OptConsultaTipoPaciente(1).Enabled = True
        Me.OptConsultaTipoPaciente(2).Enabled = True
    Else
        Me.mskFechaInicial.Enabled = False
        Me.mskFechaFinal.Enabled = False
        Me.txtNumCliente.Enabled = False
        Me.txtNombreCliente.Enabled = False
        Me.cmdCargarDatos.Enabled = False
        
        Me.OptConsultaTipoPaciente(0).Enabled = False
        Me.OptConsultaTipoPaciente(1).Enabled = False
        Me.OptConsultaTipoPaciente(2).Enabled = False
    End If
    
    pCargaNotas
End Sub

Private Sub OptMotivoNota_Click(Index As Integer)
    If OptMotivoNota(0).Value Then
        If Not vlblnConsulta Then pLimpiaNota
        lbDescuento.Visible = True
        txtDescuento.Visible = True
        txtDescuento.Enabled = True
        
        lbDescuentoCR.Visible = True
        txtDescuentoCR.Visible = True
        
        lbCantidad.Caption = "Cantidad"
        
        chkFacturasPagadas.Visible = False
        chkPorcentaje.Visible = False
        
        txtCantidad.ToolTipText = "Cantidad del concepto"
        
        '----- Agregado para caso 7374 -----'
        If optPaciente.Value And Trim(txtCveCliente.Text) <> "" And Not vlblnConsulta Then
            chkFacturasPaciente.Value = vbChecked
            chkFacturasPaciente.Enabled = Not cboFacturasPaciente.Visible
        End If
        '-----------------------------------'
        
        If optCliente.Value And optNotaCredito.Value Then
            txtDescuento.Enabled = False
        End If
    Else
        If Not vlblnConsulta Then pLimpiaNota
        lbDescuento.Visible = False
        txtDescuento.Visible = False
        
        lbDescuentoCR.Visible = False
        txtDescuentoCR.Visible = False
        
        lbCantidad.Caption = "Total nota"
        
        chkFacturasPagadas.Visible = True
        chkPorcentaje.Visible = True
        chkPorcentaje.Value = 0
        
        txtCantidad.ToolTipText = "Cantidad total de la nota"
        
        '----- Agregado para caso 7374 -----'
        If optPaciente.Value And chkFacturasPaciente.Value = vbChecked And Not vlblnConsulta Then
            chkFacturasPaciente.Value = vbUnchecked
            chkFacturasPaciente.Enabled = True
        End If
        '-----------------------------------'
        cmdIncluir.Enabled = False
    End If
End Sub

Private Sub OptMotivoNota_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        fraDatosCliente.Enabled = True
        txtCveCliente.Enabled = True
        txtCveCliente.SetFocus
    End If
End Sub

Private Sub optNotaCargo_GotFocus()
    optNotaCargo.Value = True
    
    fstrFolioDocumento 0
    lblFolio.Caption = vlstrFolioDocumento
    If Trim(lblFolio.Caption) = "0" Then
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical + vbOKOnly, "Mensaje"
        vlblnSalir = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub optNotaCargo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFecha
   End If
End Sub

Private Sub optNotaCredito_GotFocus()
    optNotaCredito.Value = True
    
    fstrFolioDocumento 0
    lblFolio.Caption = vlstrFolioDocumento
    If Trim(lblFolio.Caption) = "0" Then
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical + vbOKOnly, "Mensaje"
        vlblnSalir = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub OptPaciente_Click()
    lbCliente.Caption = "Cuenta"
    optNotaCargo.Enabled = False
    OptTipoPaciente(0).Visible = True
    OptTipoPaciente(1).Visible = True
    fraTipoPaciente.Visible = True
    OptTipoPaciente(0).Value = True
    OptMotivoNota(1).Value = True
    
    If optNotaCredito.Visible = True And optNotaCredito.Enabled = True Then optNotaCredito.SetFocus
    
    sstFacturasCreditos.TabEnabled(1) = False
    
    sstFacturasCreditos.Height = 4455
    sstFacturasCreditos.Visible = False
    sstCargos.Visible = True
    
    If chkFacturasPaciente.Value = 0 Then pConfiguraGridCargos
    
    grdNotas.ColWidth(vlintColFactura) = 0
    
    fraDatosCliente.Top = 1800
    
    fraTipoPaciente.Top = 1470
    fraTipoPaciente.Left = 150
    
    InicializaComponentes
    
    fraMotivoNota.Enabled = True    'Modificado para caso 7374
        
    pCargaTipoPaciente
End Sub

Private Sub OptPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OptTipoPaciente(0).SetFocus
    End If
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    If Not vlblnConsulta Then
        InicializaComponentes
        fraMotivoNota.Enabled = True
        fraDatosCliente.Enabled = True
    End If
End Sub

Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        optNotaCredito.SetFocus
        If txtCveCliente.Enabled = True Then txtCveCliente.SetFocus
    End If
End Sub

Private Sub sstElementos_Click(PreviousTab As Integer)
    If vlblnElementoSeleccionado And Not vlblnDePaso Then
        vlblnDePaso = True
        sstElementos.Tab = PreviousTab
        Exit Sub
    End If
    
    cmdCargar.Visible = (sstElementos.Tab <> 0)
    
    If (optNotaCredito.Value = True) And (sstElementos.Tab = 4) Then
        cmdCargar.Enabled = False
    Else
        cmdCargar.Enabled = True
    End If
    
    vlblnDePaso = False
    If vlblnElementoSeleccionado Then
        If txtCantidad.Enabled = True Then
            txtCantidad.SetFocus
        End If
    End If
    
    ' pHabilitaControles
End Sub

Private Sub pPosicionarCantidad()
    vlblnElementoSeleccionado = True
    If OptMotivoNota(1).Value = True Then
        grdFactura.SetFocus
    Else
        cmdIncluir.Enabled = True
    End If
    pEnfocaTextBox txtCantidad
End Sub

Private Sub lstArticulos_DblClick()
    pPosicionarCantidad
End Sub

Private Sub lstArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pPosicionarCantidad
    End If
End Sub

Private Sub lstConceptosFacturacion_DblClick()
    pPosicionarCantidad
End Sub

Private Sub lstConceptosFacturacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pPosicionarCantidad
    End If
End Sub

Private Sub lstEstudios_DblClick()
    pPosicionarCantidad
End Sub

Private Sub lstEstudios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pPosicionarCantidad
    End If
End Sub

Private Sub lstExamenes_DblClick()
    pPosicionarCantidad
End Sub

Private Sub lstExamenes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pPosicionarCantidad
    End If
End Sub

Private Sub lstOtrosConceptos_DblClick()
    pPosicionarCantidad
End Sub

Private Sub lstOtrosConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        pPosicionarCantidad
    End If
End Sub

Private Sub pSeleccionaElemento()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    Select Case sstElementos.Tab
        Case 0
            Set lstListaSeleccionada = lstArticulos
            vlstrCualLista = "AR"
            If lstListaSeleccionada.ListIndex = -1 Then Exit Sub
            vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intNumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                             " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                             " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                             " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                             " AND PvConceptoFacturacion.smiCveConcepto = (SELECT smiCveConceptFact FROM IvArticulo WHERE intIDArticulo = " & Trim(Str(lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex))) & ")"
        Case 1
            Set lstListaSeleccionada = lstEstudios
            vlstrCualLista = "ES"
            If lstListaSeleccionada.ListIndex = -1 Then Exit Sub
            vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intNumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                             " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                             " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa on PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                             " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                             " AND PvConceptoFacturacion.smiCveConcepto = (SELECT smiConFact FROM ImEstudio WHERE intCveEstudio = " & Trim(Str(lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex))) & ")"
        Case 2
            Set lstListaSeleccionada = lstExamenes
            If lstListaSeleccionada.ListIndex = -1 Then Exit Sub
            If lstExamenes.ItemData(lstExamenes.ListIndex) < 0 Then
                vlstrCualLista = "GE"
                vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intNumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                                 " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                                 " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                                 " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                                 " AND PvConceptoFacturacion.smiCveConcepto = (SELECT smiConFact FROM LaGrupoExamen WHERE intCveGrupo = " & Trim(Str(lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex) * -1)) & ")"
            Else
                vlstrCualLista = "EX"
                vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intNumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                                 " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                                 " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                                 " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                                 " AND PvConceptoFacturacion.smiCveConcepto = (SELECT smiConFact FROM LaExamen WHERE intCveExamen = " & Trim(Str(lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex))) & ")"
            End If
        Case 3
            Set lstListaSeleccionada = lstOtrosConceptos
            vlstrCualLista = "OC"
            If lstListaSeleccionada.ListIndex = -1 Then Exit Sub
            vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intNumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                             " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                             " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smicveconcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                             " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                             " AND  PvConceptoFacturacion.smiCveConcepto = (SELECT smiConceptoFact FROM PvOtroConcepto WHERE intCveConcepto = " & Trim(Str(lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex))) & ")"
        Case 4
            Set lstListaSeleccionada = lstConceptosFacturacion
            vlstrCualLista = "CF"
            vlstrSentencia = "SELECT PvConceptoFacturacionEmpresa.intnumCtaingreso intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                             " PvConceptoFacturacionEmpresa.intNumCtaDescuento intCuentaDescuento, PvConceptoFacturacionEmpresa.intNumCtaDescNota NumCtaDescNota " & _
                             " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                             " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                             " AND PvConceptoFacturacion.smiCveConcepto = " & lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex)
    End Select
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount = 0 Then
        vllngCveCtaDescuento = 0
        vllngCveCtaIngresos = 0
        vllngCveCtaIva = 0
        vldblPorcentajeIVA = 0
    Else
        If optNotaCredito.Value = True And OptMotivoNota(1).Value = True Then
            vllngCveCtaDescuento = IIf(IsNull(rs!NumCtaDescNota), 0, rs!NumCtaDescNota)
        Else
            vllngCveCtaDescuento = IIf(IsNull(rs!intCuentaDescuento), 0, rs!intCuentaDescuento)
        End If
        vllngCveCtaIngresos = IIf(IsNull(rs!INTCUENTACONTABLE), 0, rs!INTCUENTACONTABLE)
        vllngCveCtaIva = glngCtaIVANoCobrado
        vldblPorcentajeIVA = rs!smyIVA / 100
    End If
     
    rs.Close
    vlblnElementoSeleccionado = False
End Sub

Private Sub pSeleccionaElementoFactura()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "SELECT NVL(PvConceptoFacturacionEmpresa.intNumCtaIngreso, 0) intCuentaContable, PvConceptoFacturacion.smyIVA, " & _
                     " NVL(PvConceptoFacturacionEmpresa.intnumCtaDescuento, 0) intCuentaDescuento, NVL(PvConceptoFacturacionEmpresa.intNumCtaDescNota, 0) NumCtaDescNota " & _
                     " FROM PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                     " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                     " AND PvConceptoFacturacion.smiCveConcepto = " & Val(grdCargos.TextMatrix(grdCargos.Row, 6))
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount = 0 Then
        vllngCveCtaDescuento = 0
        vllngCveCtaIngresos = 0
        vllngCveCtaIva = 0
        vldblPorcentajeIVA = 0
    Else
        If optNotaCredito.Value = True And OptMotivoNota(1).Value = True Then
            vllngCveCtaDescuento = rs!NumCtaDescNota
        Else
            vllngCveCtaDescuento = rs!intCuentaDescuento
        End If
        vllngCveCtaIngresos = rs!INTCUENTACONTABLE
        vllngCveCtaIva = glngCtaIVANoCobrado
        vldblPorcentajeIVA = rs!smyIVA / 100
    End If
    rs.Close

    vlblnElementoSeleccionado = False
End Sub

Private Sub pSeleccionaElementoCR()
    Dim vlstrSentencia As String
    Dim vlstrSentencia1 As String
    Dim vlstrsentenciai As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim ri As New ADODB.Recordset
    Dim vdblIvatotal As Double
    
    
    
    Set lstListaSeleccionada = lstConceptosFact
    vlstrCualLista = "CF"
    vlstrSentencia = "SELECT NVL(PvConceptoFacturacionEmpresa.intNumCtaingreso, 0) intCuentaContable, PvConceptoFacturacion.smyIVA," & _
                     " NVL(PvConceptoFacturacionEmpresa.intNumCtaDescuento, 0) intCuentaDescuento FROM PvConceptoFacturacion " & _
                     " INNER JOIN PvConceptoFacturacionEmpresa ON PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura " & _
                     " WHERE PvConceptoFacturacionEmpresa.intCveEmpresaContable = " & vgintClaveEmpresaContable & _
                     " AND PvConceptoFacturacion.smiCveConcepto = " & lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex)
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    vlstrsentenciai = "Select relporcentaje from cnimpuesto where smicveimpuesto = (Select vchvalor from siparametro where vchnombre = 'INTTASAIMPUESTOHOSPITAL' and intcveempresacontable = " & vgintClaveEmpresaContable & ") and bitactivo = 1"
    Set ri = frsRegresaRs(vlstrsentenciai, adLockReadOnly, adOpenForwardOnly)
    
    vdblIvatotal = ri!relPorcentaje / 100
   
  vlstrSentencia1 = "Select " & _
                       "CASE WHEN (ROUND (MNYSUBTOTAL,2) - ROUND(MNYIVA/" & vdblIvatotal & ",2)) < 1 THEN ROUND (MNYIVA/" & vdblIvatotal & ",2) + (ROUND (MNYSUBTOTAL,2) - ROUND(MNYIVA/" & vdblIvatotal & ",2)) ELSE ROUND (MNYIVA/" & vdblIvatotal & ",2) END AS MNYCONIVA, " & _
                       "MNYSUBTOTAL - MNYIVA /" & vdblIvatotal & " as MNYSINIVA, " & _
                       "chrfolioreferencia ," & _
                       "ROUND (MNYSUBTOTAL,2) as MNYSUBTOTAL , " & _
                       "CASE WHEN (ROUND (MNYSUBTOTAL,2) - ROUND(MNYIVA/" & vdblIvatotal & ",2)) < 1 THEN ROUND (MNYIVA/" & vdblIvatotal & ",2) + (ROUND (MNYSUBTOTAL,2) - ROUND(MNYIVA/" & vdblIvatotal & ",2)) ELSE ROUND (MNYIVA/" & vdblIvatotal & ",2) END AS MNYIVA, " & _
                       "case when (ROUND (MNYSUBTOTAL,2) - ROUND(MNYIVA/" & vdblIvatotal & ",2)) <= 0 then 1 else 0 end as DIF " & _
                       " From ccmovimientocredito " & _
                       " Where INTNUMCLIENTE = " & Trim(txtCveCliente) & _
                       " And chrfolioreferencia = '" & Trim(cboCreditosDirectos.List(cboCreditosDirectos.ListIndex)) & "'"
      

   Set rs1 = frsRegresaRs(vlstrSentencia1, adLockReadOnly, adOpenForwardOnly)
   vldblImporteIvaCero = rs1!MNYSINIVA
   vldblImporteIvaNoCero = rs1!MNYCONIVA
   
   If rs1.RecordCount <> 0 Then
     'Verificar IVACero
     If rs1!dif = 0 Then
        If (rs1!MNYCONIVA <> 0 And rs1!MNYSINIVA <> 0) Or (rs1!MNYCONIVA = 0 And rs1!MNYSINIVA = 0) Then
           vldblIVAcero = 1
        End If
     Else
        vldblIVAcero = 0
     End If
     
   End If
   'Verificar Pago
    If rs.RecordCount <> 0 Then
      If rs!smyIVA > 1 Then
            vldblIVAcero = 1
        If Val(Format(txtCantidadCR.Text, vlstrFormato)) <= Val(Format(rs1!MNYCONIVA, vlstrFormato)) Or Val(Format(rs1!MNYCONIVA, vlstrFormato)) <= 0 And rs1!MNYCONIVA <> 0 Then
            vldblIVApago = 1
        Else
            vldblIVApago = 0
            If rs1!MNYCONIVA = 0 Then
               vldblIVAcero = 0
            End If
        End If
      Else
        If Val(Format(txtCantidadCR.Text, vlstrFormato)) <= Val(Format(rs1!MNYSINIVA, vlstrFormato)) Or Val(Format(rs1!MNYSINIVA, vlstrFormato)) >= 0 And Val(rs1!MNYSINIVA) <> 0 Then
            vldblIVApago = 1
        Else
            vldblIVApago = 0
        End If
      End If
    End If
   
    
    If rs.RecordCount = 0 Then
        vllngCveCtaDescuento = 0
        vllngCveCtaIngresos = 0
        vllngCveCtaIva = 0
        vldblPorcentajeIVA = 0
    Else
        vllngCveCtaDescuento = rs!intCuentaDescuento
        vllngCveCtaIngresos = rs!INTCUENTACONTABLE
        vllngCveCtaIva = glngCtaIVANoCobrado
        vldblPorcentajeIVA = rs!smyIVA / 100
    End If
    rs.Close
End Sub

Private Sub chkMedicamentos_Click()
    If fblnCanFocus(txtSeleArticulo) Then txtSeleArticulo.SetFocus
    txtSeleArticulo_KeyUp 7, 0
End Sub

Private Sub optClave_Click()
    lstArticulos.Clear
    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 11
    If fblnCanFocus(txtSeleArticulo) Then txtSeleArticulo.SetFocus
End Sub

Private Sub optDescripcion_Click()
    lstArticulos.Clear
    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 30
    If fblnCanFocus(txtSeleArticulo) Then txtSeleArticulo.SetFocus
End Sub

Private Sub sstFacturasCreditos_Click(PreviousTab As Integer)
    Dim rsInfCredito As ADODB.Recordset
    Dim strParametros As String
    Dim i As Integer
    Dim strFactura As String
    Dim strTipoReferencia As String
    
On Error GoTo NotificaError
    
    If sstFacturasCreditos.Tab = 0 Then
        pConfiguraGridNota "FA"
    End If
    
    If sstFacturasCreditos.Tab = 1 Then
        pConfiguraGridNota "MA"
    End If
    
    freDetalleNota.Enabled = True
    lblUsoCFDI.Enabled = True
    lblMetodoPago.Enabled = True
    lblFormaPago.Enabled = True
    
    cboUsoCFDI.Enabled = True
    cboMetodoPago.Enabled = True
    cboFormaPago.Enabled = True

    cmdGrabarRegistro.Enabled = True
    
    If sstFacturasCreditos.Tab = 2 Then
        pConfiguraGridNotaInfo
        freDetalleNota.Enabled = False
        lblUsoCFDI.Enabled = False
        lblMetodoPago.Enabled = False
        lblFormaPago.Enabled = False
        
        cboUsoCFDI.Enabled = False
        cboMetodoPago.Enabled = False
        cboFormaPago.Enabled = False
                
        If PreviousTab = 0 Then
            strFactura = cboFactura.Text
            strTipoReferencia = "FA"
        End If
        
        If PreviousTab = 1 Then
            strFactura = cboCreditosDirectos.Text
            strTipoReferencia = "MA"
        End If
                    
        strParametros = strFactura & "|" & txtCveCliente & "|" & strTipoReferencia
        Set rsInfCredito = frsEjecuta_SP(strParametros, "SP_CCSELNOTAS")
    
        pLlenarMshFGrdRs grdInformacionNota, rsInfCredito, 0
                    
        For i = 1 To grdInformacionNota.Rows - 1
            With grdInformacionNota
                .TextMatrix(i, cintColCantidadNota) = IIf(.TextMatrix(i, cintColCantidadNota) = "", "", FormatCurrency(Val(.TextMatrix(i, cintColCantidadNota)), 2))
                .TextMatrix(i, cintColDescuentoNota) = IIf(.TextMatrix(i, cintColDescuentoNota) = "", "", FormatCurrency(Val(.TextMatrix(i, cintColDescuentoNota)), 2))
                .TextMatrix(i, cintColIVANota) = IIf(.TextMatrix(i, cintColIVANota) = "", "", FormatCurrency(Val(.TextMatrix(i, cintColIVANota)), 2))
                .TextMatrix(i, cintColTotalNota) = IIf(.TextMatrix(i, cintColTotalNota) = "", "", FormatCurrency(Val(.TextMatrix(i, cintColTotalNota)), 2))
            End With
        Next i
                                
        pFormatoGridNotaInfo
        
        Set rsInfCredito = frsEjecuta_SP(strParametros, "SP_CCSELINFCREDITO")
        If rsInfCredito.RecordCount > 0 Then
            txtInfTotalNotas.Text = FormatCurrency(rsInfCredito!Notas, 2)
            txtInfTotalpagos.Text = FormatCurrency(rsInfCredito!pagos, 2)
            txtInfSaldo.Text = FormatCurrency(rsInfCredito!Saldo, 2)
            txtInfCredito.Text = FormatCurrency(rsInfCredito!credito, 2)
        Else
            '¡La información no existe!
            MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
        End If
                    
        rsInfCredito.Close
        cmdGrabarRegistro.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstFacturasCreditos_Click"))
    Unload Me
End Sub

Private Sub tmrCargos_Timer()
On Error GoTo NotificaError

    If chkFacturasPaciente.Value Then
        chkFacturasPaciente_Click
        tmrCargos.Enabled = False
    Else
        pLlenaCargos txtCveCliente.Text
        tmrCargos.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrCargos_Timer"))
    Unload Me
End Sub

Private Sub txtBuscaCargo_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intLongitud As Integer

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or KeyAscii = 32 Then
        txtBuscaCargo_KeyUp KeyAscii, 0
    Else
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then KeyAscii = 7
        End If
    End If
    
    If KeyAscii = 13 Then
        grdCargos.Row = 1
        grdCargos.SetFocus
    End If
End Sub

Private Sub txtBuscaCargo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    pLlenaCargos txtCveCliente.Text

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBuscaCargo_KeyUp"))
    Unload Me
End Sub

Private Sub txtCantidad_LostFocus()
On Error GoTo NotificaError

    If chkPorcentaje.Value = 1 Then
        If Val(txtCantidad.Text) > 100 Then txtCantidad.Text = 100
    Else
        If txtCantidad.Enabled And Trim(txtCantidad.Text) <> "" And Val(txtCantidad.Text) > 0 Then
            txtCantidad.Text = FormatCurrency(txtCantidad.Text, 2)
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_LostFocus"))
    Unload Me
End Sub

Private Sub txtCantidadCargo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdIncluirCargo.Enabled = True Then cmdIncluirCargo.SetFocus
        
        If chkFacturasPaciente.Value = 1 Then
            If CDbl(Val(txtCantidadCargo)) > fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))) Then
                txtCantidadCargo.Text = CStr(fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex))))
            End If
            
            grdCargos.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadCargo_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidadCargo_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidadCargo, KeyAscii, 2) Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtCantidadCargo_LostFocus()
On Error GoTo NotificaError
        
    If chkPorcentajePaciente.Value = 1 Then
        If txtCantidadCargo.Text = "" Then
            txtCantidadCargo.Text = 100#
        ElseIf CDbl(txtCantidadCargo.Text) > 100 Then
            txtCantidadCargo.Text = 100#
        End If
    Else
        If txtCantidadCargo.Enabled And Trim(txtCantidadCargo.Text) <> "" Then
            txtCantidadCargo.Text = Replace(txtCantidadCargo.Text, "$", "")
            txtCantidadCargo.Text = FormatCurrency(txtCantidadCargo.Text, 2)
            If chkPorcentajePaciente.Visible Then
                vldbltotal = fTotalFactura(grdCargos, vgstrParametrosSP)
                If CDbl(txtCantidadCargo.Text) > vldbltotal Then
                    txtCantidadCargo.Text = FormatCurrency(vldbltotal, 2)
                End If
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadCargo_LostFocus"))
    Unload Me
End Sub

Private Sub txtCantidadCR_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtCantidadCR
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadCR_GotFocus"))
    Unload Me
End Sub

Private Sub txtCantidadCR_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        txtCantidadCR.Text = Format(txtCantidadCR.Text, vlstrFormato)
        If txtDescuentoCR.Visible Then
            If fblnCanFocus(txtDescuentoCR) Then txtDescuentoCR.SetFocus
        Else
            If fblnCanFocus(cmdIncluirCR) Then cmdIncluirCR.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadCR_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidadCR_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not fblnFormatoCantidad(txtCantidadCR, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidadCR_KeyPress"))
    Unload Me
End Sub

Private Sub txtCantidadCR_LostFocus()
    If txtCantidadCR.Enabled And Trim(txtCantidadCR.Text) <> "" Then
        txtCantidadCR.Text = FormatCurrency(txtCantidadCR.Text, 2)
    End If
End Sub

Private Sub txtComentario_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtComentario
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtComentario_GotFocus"))
    Unload Me
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 And cmdGrabarRegistro.Enabled = True Then
        cmdGrabarRegistro.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtComentario_KeyPress"))
    Unload Me
End Sub

Private Sub txtCveCliente_Change()
On Error GoTo NotificaError
    
    txtCliente.Text = ""
    txtDomicilio.Text = ""
    txtRFC.Text = ""
    
    If optCliente.Value = True Then
        pLimpiaDetalleFactura
        pLimpiaConcepto
        pLimpiaNota
    Else
        grdCargos.Redraw = False
        pLimpiaGridCargos
        pConfiguraGridCargos
        cboFacturasPaciente.Clear
        chkFacturasPaciente.Value = 0
        grdCargos.Redraw = True
        If txtCveCliente = "" Then
            pLimpiaGridNota
            pConfiguraGridNota "FA"
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveCliente_Change"))
    Unload Me
End Sub

Private Sub txtDescuentoCR_GotFocus()
On Error GoTo NotificaError

    pSelTextBox txtDescuentoCR
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuentoCR_GotFocus"))
    Unload Me
End Sub

Private Sub txtDescuentoCR_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        txtDescuentoCR.Text = Format(txtDescuentoCR.Text, vlstrFormato)
        If cmdIncluirCR.Enabled Then
            cmdIncluirCR.SetFocus
        Else
            txtCantidadCR.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuentoCR_KeyDown"))
    Unload Me
End Sub

Private Sub txtDescuentoCR_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not fblnFormatoCantidad(txtDescuentoCR, KeyAscii, 2) Then
        KeyAscii = 7
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuentoCR_KeyPress"))
    Unload Me
End Sub

Private Sub txtNombreCliente_GotFocus()
Me.cmdCargarDatos.SetFocus
End Sub

Private Sub txtNumCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vllngNumCliente As Long
    Dim vgstrParametrosSP As String
    Dim rsPaciente As ADODB.Recordset

On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        If optCliente.Value = True Then
            If Trim(txtNumCliente.Text) = "" Then
                vllngNumCliente = flngNumCliente(False, 1)
                If vllngNumCliente <> 0 Then
                   txtNumCliente.Text = vllngNumCliente
                End If
            End If
            
            If Trim(txtNumCliente.Text) <> "" Then
                pDatosCliente txtNumCliente.Text
                If rsDatosCliente.RecordCount <> 0 Then
                    If rsDatosCliente!smicvedepartamento = vgintNumeroDepartamento Then
                        pAsignarTxtNombre rsDatosCliente!NombreCliente, txtNombreCliente
                        cmdCargarDatos.SetFocus
                    Else
                        'El cliente seleccionado no pertenece a este departamento.
                        MsgBox SIHOMsg(646), vbOKOnly + vbInformation, "Mensaje"
                        txtNumCliente.Text = ""
                    End If
                End If
            Else
                cmdCargarDatos.SetFocus
            End If
        Else
            If Trim(txtNumCliente.Text) = "" Then
                With FrmBusquedaPacientes
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    
                    If OptConsultaTipoPaciente(1).Value Then 'Externos
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,1700,4100"
                        .vgstrTipoPaciente = "E"
                        .Caption = .Caption & " Externos"
                    Else
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,990,990,4100"
                        .vgstrTipoPaciente = "I"
                        .Caption = .Caption & " Internos"
                    End If
            
                    txtNumCliente.Text = .flngRegresaPaciente()
                End With
            Else
                cmdCargarDatos.SetFocus
            End If  'End If Trim(txtNumCliente.Text) = ""
                
            ' Carga datos del paciente
            If Trim(txtNumCliente.Text) <> "" Then
                vgstrParametrosSP = Trim(txtNumCliente.Text) & "|" & Str(vgintClaveEmpresaContable)
                
                If OptConsultaTipoPaciente(0).Value = True Then
                    Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
                Else
                    Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
                End If
                
                If rsPaciente.RecordCount > 0 Then
                    pAsignarTxtNombre IIf(IsNull(rsPaciente!Nombre), " ", rsPaciente!Nombre), txtNombreCliente
                Else
                    If txtNumCliente.Text <> -1 Then
                        '¡La información no existe!
                        MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                    End If
                    
                    txtNumCliente.Text = ""
                    txtNombreCliente.Text = ""
                    txtNumCliente.SetFocus
                End If ' End rsPaciente.RecordCount > 0
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                txtNumCliente.Text = ""
                txtNombreCliente.Text = ""
                txtNumCliente.SetFocus
            End If ' End if Trim(txtNumCliente.Text) <> ""
        End If ' End If OptCliente.Value
    End If ' End if KeyCode
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumCliente_KeyDown"))
    Unload Me
End Sub

Private Sub txtNumCliente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        Select Case UCase(Chr(KeyAscii))
            Case "I"
                If OptConsultaTipoPaciente(0).Enabled Then
                    OptConsultaTipoPaciente(0).Value = True
                    txtNumCliente.SetFocus
                End If
            Case "E"
                If OptConsultaTipoPaciente(1).Enabled Then
                    OptConsultaTipoPaciente(1).Value = True
                    txtNumCliente.SetFocus
                End If
        End Select
        
        KeyAscii = 7
    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumCliente_KeyPress"))
    Unload Me
End Sub

Private Sub txtPendienteTimbre_GotFocus()
If fblnCanFocus(Me.cmdPrimerRegistro) Then Me.cmdPrimerRegistro.SetFocus
End Sub

Private Sub txtSeleArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If lstArticulos.Enabled Then
            lstArticulos.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtSeleArticulo_KeyDown"))
    Unload Me
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtSeleArticulo_KeyPress"))
    Unload Me
End Sub

Private Sub txtSeleArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    Dim vlstrOtroFiltro As String
    
On Error GoTo NotificaError

    If chkMedicamentos.Value = 1 Then
        vlstrOtroFiltro = " and chrCveArtMedicamen = '1'"
    Else
        vlstrOtroFiltro = ""
    End If
    
    If optDescripcion.Value Then
        vlstrSentencia = "Select intIDArticulo, vchNombreComercial from ivarticulo"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstArticulos, "vchNombreComercial", 100, vlstrOtroFiltro, "vchNombreComercial"
    Else
        vlstrSentencia = "Select intIDArticulo, vchNombreComercial from ivarticulo"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstArticulos, "chrCveArticulo", 100, vlstrOtroFiltro, "vchNombreComercial"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtSeleArticulo_KeyUp"))
    Unload Me
End Sub

Private Sub cmdCargar_Click()
On Error GoTo NotificaError

    pCargaArticulos True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargar_Click"))
    Unload Me
End Sub

Private Sub pCargaArticulos(vlblnManipularBotonCargar As Boolean)
    Dim vlstrSentencia As String
    Dim rsDatos As New ADODB.Recordset
    Dim vlintContador As Integer

On Error GoTo NotificaError

    pgbCargando.Value = 0
    lblTextoBarra.Caption = SIHOMsg(280)
    
    If vlblnManipularBotonCargar Then cmdCargar.Enabled = False
    
    lstArticulos.Visible = False
    
    Select Case sstElementos.Tab
        Case 0
            lstArticulos.Clear
            vlstrSentencia = "SELECT chrCveArticulo, vchNombreComercial FROM IvArticulo WHERE vchEstatus='ACTIVO' order by vchNombreComercial"
        Case 1
            lstEstudios.Visible = False
            lstEstudios.Clear
            vlstrSentencia = "SELECT intCveEstudio, vchNombre FROM ImEstudio WHERE bitStatusActivo=1 order by vchNombre"
        Case 2
            lstExamenes.Visible = False
            lstExamenes.Clear
            vlstrSentencia = "SELECT intCveExamen, chrNombre FROM LaExamen WHERE bitEstatusActivo = 1 Union SELECT LaGrupoExamen.INTCVEGRUPO * -1, LaGrupoExamen.chrNombre From LaGrupoExamen Where bitEstatusActivo = 1 ORDER BY 2,1"
        Case 3
            lstOtrosConceptos.Visible = False
            lstOtrosConceptos.Clear
            vlstrSentencia = "SELECT intCveConcepto, chrDescripcion FROM PvOtroConcepto WHERE bitEstatus=1 order by chrDescripcion "
        Case 4
            lstConceptosFacturacion.Visible = False
            lstConceptosFacturacion.Clear
            vlstrSentencia = "SELECT smiCveConcepto, chrDescripcion FROM PvConceptoFacturacion WHERE bitActivo = 1 order by chrDescripcion"
    End Select
    
    Set rsDatos = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If rsDatos.RecordCount > 500 Then
        lblTextoBarra.Caption = SIHOMsg(280)
        freBarra.Top = 4560
        freBarra.Visible = True
        freBarra.Refresh
    End If
    
    Do While Not rsDatos.EOF
        Select Case sstElementos.Tab
            Case 0
                lstArticulos.AddItem (rsDatos!VCHNOMBRECOMERCIAL)
                lstArticulos.ItemData(lstArticulos.newIndex) = CLng(rsDatos!chrcvearticulo)
            Case 1
                lstEstudios.AddItem (rsDatos!vchNombre)
                lstEstudios.ItemData(lstEstudios.newIndex) = rsDatos!intCveEstudio
            Case 2
                lstExamenes.AddItem (rsDatos!CHRNOMBRE)
                lstExamenes.ItemData(lstExamenes.newIndex) = rsDatos!IntCveExamen
            Case 3
                lstOtrosConceptos.AddItem (rsDatos!chrdescripcion)
                lstOtrosConceptos.ItemData(lstOtrosConceptos.newIndex) = rsDatos!intCveConcepto
            Case 4
                lstConceptosFacturacion.AddItem (rsDatos!chrdescripcion)
                lstConceptosFacturacion.ItemData(lstConceptosFacturacion.newIndex) = rsDatos!smicveconcepto
        End Select
    
        rsDatos.MoveNext
        If Not rsDatos.EOF And rsDatos.RecordCount > 500 Then
            If rsDatos.Bookmark Mod 100 = 0 Then
                pgbCargando.Value = (rsDatos.Bookmark / rsDatos.RecordCount) * 100
            End If
        End If
    Loop
    
    Select Case sstElementos.Tab
        Case 0
            If lstArticulos.ListCount > 0 Then
                lstArticulos.ListIndex = 0
            End If
            lstArticulos.Visible = True
        Case 1
            If lstEstudios.ListCount > 0 Then
                lstEstudios.ListIndex = 0
            End If
            lstEstudios.Visible = True
        Case 2
            If lstExamenes.ListCount > 0 Then
                lstExamenes.ListIndex = 0
            End If
            lstExamenes.Visible = True
        Case 3
            If lstOtrosConceptos.ListCount > 0 Then
                lstOtrosConceptos.ListIndex = 0
            End If
            lstOtrosConceptos.Visible = True
        Case 4
            If lstConceptosFacturacion.ListCount > 0 Then
                lstConceptosFacturacion.ListIndex = 0
            End If
            lstConceptosFacturacion.Visible = True
    End Select
    
    rsDatos.Close
    lstArticulos.Visible = True
    freBarra.Visible = False
    If vlblnManipularBotonCargar Then cmdCargar.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaArticulos"))
    Unload Me
End Sub
Private Sub pCargaNotas()
    Dim vllngContador As Long
    Dim vllngColor As Long
    Dim lngcolorSub As Long
    Dim rsNotas As New ADODB.Recordset
    Dim vlBackColor As Variant
    Dim vlForeColor As Variant
    Dim vlblnColorNE As Boolean

On Error GoTo NotificaError
    vllngSeleccionadas = 0
    vllngSeleccPendienteTimbre = 0
    If (Not IsDate(mskFechaInicial) Or Not IsDate(mskFechaFinal)) And Me.optMostrarSolo(2).Value = False Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
        pEnfocaMkTexto mskFechaInicial
    Else
       
        If txtNumCliente.Text = "" Then txtNombreCliente.Text = ""
        
        pConfiguraBusqueda
                
        vgstrParametrosSP = Str(vgintNumeroDepartamento) & "|" _
                            & IIf(Me.optMostrarSolo(0).Value = False, fstrFechaSQL("01/01/2010"), fstrFechaSQL(mskFechaInicial.Text)) _
                            & "|" & IIf(Me.optMostrarSolo(0).Value = False, fstrFechaSQL(fdtmServerFecha), fstrFechaSQL(mskFechaFinal.Text)) _
                            & "|" & Trim(Str(Val(txtNumCliente.Text))) _
                            & "|" & IIf(optCliente.Value = True, "C", "P") _
                            & "|" & IIf(Me.optMostrarSolo(0).Value = False, "TO", IIf(OptConsultaTipoPaciente(0) = True, "I", IIf(OptConsultaTipoPaciente(1).Value = True, "E", "TO"))) _
                            & "|" & Str(vgintClaveEmpresaContable) _
                            & "|" & IIf(Me.optMostrarSolo(2).Value = True, 1, 0) _
                            & "|" & IIf(Me.optMostrarSolo(1).Value = True, 1, 0) _
                            & "|" & IIf(Me.optMostrarSolo(3).Value = True, 1, 0) _
                            & "|" & IIf(Me.optMostrarSolo(4).Value = True, 1, 0)
                                                        
        Set rsNotas = frsEjecuta_SP(vgstrParametrosSP, "Sp_CcSelNotaCreditoCargo")
        If rsNotas.RecordCount = 0 Then
            'No existe información con esos parámetros.
            If Not blnNoMensaje Then MsgBox SIHOMsg(236), vbInformation + vbOKOnly, "Mensaje"
            If Me.optMostrarSolo(2).Value = True Then
               optMostrarSolo(2).SetFocus
            ElseIf Me.optMostrarSolo(1).Value = True Then
               optMostrarSolo(1).SetFocus
            Else
               pEnfocaMkTexto mskFechaInicial
            End If
        Else
            With grdBusqueda
                Do While Not rsNotas.EOF
                    .Row = .Rows - 1
                         
'                    If vgblnNuevoEsquemaCancelacion Then
'                        .TextMatrix(.Row, 0) = IIf(rsNotas!PendienteCancelarSat_NE <> "CA", "*", "")
'                    Else
'                        .TextMatrix(.Row, 0) = IIf(rsNotas!PendienteTimbreFiscal = 1 Or rsNotas!PendienteCancelarSat > 0, "*", "")
'                    End If
                                             
                    If rsNotas!PendienteTimbreFiscal = 1 Then
                        If rsNotas!PendienteCancelarSAT_NE = "CR" Or rsNotas!PendienteCancelarSAT_NE = "PA" Then
                            .TextMatrix(.Row, 0) = ""
                        Else
                            .TextMatrix(.Row, 0) = IIf(rsNotas!PendienteCancelarSAT_NE <> "CA", "*", "")
                        End If
                       
                       lngcolorSub = &H80FFFF
                       If .TextMatrix(.Row, 0) = "*" Then
                            vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                       End If
                       vllngColor = vllngColorActivas
                    Else
                       If rsNotas!chrestatus = "C" Then
                          vllngColor = vllngColorCanceladas
                       Else
                          vllngColor = vllngColorActivas
                       End If

                       If rsNotas!PendienteCancelarSat > 0 Then
                            If rsNotas!PendienteCancelarSAT_NE = "CR" Then
                                .TextMatrix(.Row, 0) = ""
                            Else
                                .TextMatrix(.Row, 0) = IIf(rsNotas!PendienteCancelarSAT_NE <> "CA", "*", "")
                            End If

                            If .TextMatrix(.Row, 0) = "*" Then
                                 vllngSeleccionadas = vllngSeleccionadas + 1
                            End If
                            lngcolorSub = llnColorPenCancelaSAT
                       Else
                          lngcolorSub = llncolorCanceladasSAT
                       End If
                    End If
                    
                    .TextMatrix(.Row, vlintcolPTimbre) = rsNotas!PendienteTimbreFiscal
                    
                    .TextMatrix(.Row, vlintColPFacSAT) = rsNotas!PendienteCancelarSAT_NE
                    
                    vlblnColorNE = False
                    If vgblnNuevoEsquemaCancelacion Then
                        If rsNotas!PendienteTimbreFiscal <> 1 Then
                            If rsNotas!chrestatus = "C" Then
                                Select Case rsNotas!PendienteCancelarSAT_NE
                                    Case "PC"
                                        vlForeColor = &HFF&    '| Rojo
                                        vlBackColor = &HC0E0FF '| Naranja
                                        vlblnColorNE = True
                                    Case "PA"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &H80FF&  '| Naranja fuerte
                                        vlblnColorNE = True
                                    Case "CR"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &HFF&    '| Rojo
                                        vlblnColorNE = True
                                    Case "NP"
                                        vlForeColor = &HFF&    '| Blanco
                                        vlBackColor = &HFFFFFF '| Rojo
                                        vlblnColorNE = True
                                End Select
                            Else
                                Select Case rsNotas!PendienteCancelarSAT_NE
                                    Case "PC"
                                        vlForeColor = &HFF&    '| Rojo
                                        vlBackColor = &HC0E0FF '| Naranja
                                        vlblnColorNE = True
                                    Case "PA"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &H80FF&  '| Naranja fuerte
                                        vlblnColorNE = True
                                    Case "CR"
                                        vlForeColor = &HFFFFFF '| Blanco
                                        vlBackColor = &HFF&    '| Rojo
                                        vlblnColorNE = True
                                    Case "NP"
                                        vlForeColor = &H0&     '| Negro
                                        vlBackColor = &HFFFFFF '| Rojo
                                        vlblnColorNE = True
                                End Select
                            End If
                        End If
                    End If
                                    
                    .Col = vlintColIdNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!intConsecutivo
                    
                    .Col = vlintColFechaNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = Format(rsNotas!dtmfecha, "dd/mmm/yyyy")
                    
                    .Col = vlintColFolioNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!chrFolioNota
                    
                    .Col = vlintColTipoNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!TipoNota
                    
                    .Col = vlintColNumCliente
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!intCliente
                    
                    .Col = vlintColNombreCliente
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!NombreCliente), "", rsNotas!NombreCliente)
                    
                    .Col = vlintColEstadoNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!EstadoNota), "", rsNotas!EstadoNota)
                    
                    .Col = vlintColDomicilioCliente
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!DomicilioCliente), "", rsNotas!DomicilioCliente)
                    
                    .Col = vlintColRFCCliente
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!RFCCliente), "", rsNotas!RFCCliente)
                    
                    .Col = vlintColSubtotalNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!MNYSUBTOTAL), "", rsNotas!MNYSUBTOTAL)
                    
                    .Col = vlintColDescuentoNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!MNYDESCUENTO), 0, rsNotas!MNYDESCUENTO)
                    
                    .Col = vlintColIVANota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!MNYIVA), 0, rsNotas!MNYIVA)
                    
                    .Col = vlintColTotalNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!TotalNota), 0, rsNotas!TotalNota)
                    
                    .Col = vlintColchrTipo
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!chrTipo
                    
                    .Col = vlintColchrEstatus
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!chrestatus
                    
                    .Col = vlintColintNumPoliza
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(IsNull(rsNotas!intNumPoliza), "", rsNotas!intNumPoliza)
                    
                    .Col = vlintColdtmFechaRegistro
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!dtmFechaRegistro
                    
                    .Col = vlintColCuentaContable
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!cuentacontable
                    
                    .Col = vlintColTipoNotaFACR
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = rsNotas!TipoFolio
                    
                    .Col = vlintColMotivoNota
                    If vlblnColorNE Then
                        .CellForeColor = vlForeColor
                        .CellBackColor = vlBackColor
                    Else
                        .CellForeColor = vllngColor
                        .CellBackColor = lngcolorSub
                    End If
                    .TextMatrix(.Row, .Col) = IIf(rsNotas!chrTipo = "CR", IIf(rsNotas!MotivoNota = "E", "Error de facturación", "Otorgamiento de descuentos"), "")
                    
                    rsNotas.MoveNext
                    .Rows = .Rows + 1
                Loop
                .Rows = .Rows - 1
                .Row = 1
                .Col = 2
                If .Visible And .Enabled Then .SetFocus
            End With
        End If
        rsNotas.Close
    End If
       Me.cmdCancelaNotasSAT.Enabled = vllngSeleccionadas > 0
       Me.cmdConfirmarTimbreFiscal.Enabled = vllngSeleccPendienteTimbre > 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaNotas"))
    Unload Me
End Sub

Private Sub pHabilita(vlblnTop As Boolean, vlblnPrevious As Boolean, vlblnBusqueda As Boolean, vlblnnext As Boolean, vlblnEnd As Boolean, vlblnsave As Boolean, vlblnDelete As Boolean, vlblnDesglose As Boolean)
On Error GoTo NotificaError
    
    cmdPrimerRegistro.Enabled = vlblnTop
    cmdAnteriorRegistro.Enabled = vlblnPrevious
    cmdBuscar.Enabled = vlblnBusqueda
    cmdSiguienteRegistro.Enabled = vlblnnext
    cmdUltimoRegistro.Enabled = vlblnEnd
    cmdGrabarRegistro.Enabled = vlblnsave
    cmdDelete.Enabled = vlblnDelete
    cmdDesglose.Enabled = vlblnDesglose

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
    Unload Me
End Sub

Private Sub cboFactura_Click()
On Error GoTo NotificaError
    
    Dim rsDatosFactura As New ADODB.Recordset
    Dim llngConceptos As Long
    Dim vldblFacturaIVA As Double
    Dim vldblTotalFactura As Double
    Dim vlintContador As Integer
    Dim vlblnMonedaValida As Boolean
    Dim RsDescuentoEspecial As ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vldblDescEspecial As Double
    Dim vldblProporcionalSeguros As Double
    
    If cboFactura.ListIndex <> -1 Then
        'Cuenta y paciente al que pertenece la factura
        vgstrParametrosSP = Trim(cboFactura.List(cboFactura.ListIndex))
        Set rsDatosFactura = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFactura")
        If rsDatosFactura.RecordCount <> 0 Then
            vlstrSentencia = "SELECT NVL(PVFACTURA.NUMPORCENTDESCUENTOESPECIAL,0) PORCDESCESPECIAL FROM PVFACTURA WHERE CHRFOLIOFACTURA = '" & Trim(cboFactura.List(cboFactura.ListIndex) & "'")
            Set RsDescuentoEspecial = frsRegresaRs(vlstrSentencia)
        
            txtCuenta.Text = rsDatosFactura!cuenta
            txtNombrePaciente.Text = IIf(IsNull(rsDatosFactura!NOMBREPACIENTE), "", rsDatosFactura!NOMBREPACIENTE)
            lblFechaFactura.Caption = Format(rsDatosFactura!fecha, "dd/mmm/yyyy")
            vlstrtipofactura = rsDatosFactura!chrTipoFactura
            
            ' --------
            ' Verifica que la moneda de la factura seleccionada sea igual a la moneda de la(s) factura(s) ya incluida(s) en la nota
            vlblnMonedaValida = True
            For vlintContador = 1 To grdNotas.Rows - 1
                If rsDatosFactura!TipoCambio = 0 Then
                    If Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 0 And Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 1 Then
                        vlblnMonedaValida = False
                        Exit For
                    End If
                ElseIf rsDatosFactura!TipoCambio <> 0 Then
                    If Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) <> 0 And Val(grdNotas.TextMatrix(vlintContador, vlintColTipoCambio)) = 1 Then
                        vlblnMonedaValida = False
                        Exit For
                    End If
                End If
            Next vlintContador
            If vlblnMonedaValida = False Then
                pLimpiaGridFactura
                pConfiguraGridFactura
                txtCantidad.Text = ""
                rsDatosFactura.Close
                MsgBox IIf(optNotaCargo.Value = True, Replace(SIHOMsg(1135), "crédito", "cargo"), SIHOMsg(1135)), vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
            End If
            ' --------
                        
            grdFactura.Visible = False
            pLimpiaGridFactura
            If rsDatosFactura.RecordCount <> 0 Then
                '------------- (CR) - Agregado para caso 7249 -------------'
                vgblnFacturaDirecta = (rsDatosFactura!chrTipoFactura = "C") 'Revisar si es factura directa
                sstElementos.TabEnabled(0) = Not vgblnFacturaDirecta
                sstElementos.TabEnabled(1) = Not vgblnFacturaDirecta
                sstElementos.TabEnabled(2) = Not vgblnFacturaDirecta
                sstElementos.TabEnabled(3) = Not vgblnFacturaDirecta
                If vgblnFacturaDirecta Then
                    sstElementos.Tab = 4 'Seleccionar la pestaña de Conceptos de Facturación
                Else
                    sstElementos.Tab = 0 'Seleccionar la pestaña de Artículos
                End If
                '----------------------------------------------------------'
            
                '---------------- Agregado para caso 12808 ----------------'
                vldblTotalConceptosCargos = 0
                vldblTotalIVAConceptosCargos = 0
                vldblTotalConceptosSeguros = 0
                vldblTotalIVAConceptosSeguros = 0
                vldblTotalConceptosCargosConDescuento = 0
            
                vldblTotalConceptosCargosSIgrava = 0
                vldblTotalConceptosCargosNOgrava = 0
                vldblTotalConceptosCargosConDescuentoSIgrava = 0
                vldblTotalConceptosCargosConDescuentoNOgrava = 0
                vldblTotalConceptosSegurosSIgrava = 0
                vldblTotalConceptosSegurosNOgrava = 0
                vldblPorcentajeFacturaSIgrava = 0
                vldblPorcentajeFacturaNOgrava = 0
                vldblPorcentajeFacturaConDescuentoSIgrava = 0
                vldblPorcentajeFacturaConDescuentoNOgrava = 0
                
                rsDatosFactura.MoveFirst
                Do While Not rsDatosFactura.EOF
                    If rsDatosFactura!chrTipo = "NO" Or rsDatosFactura!chrTipo = "OC" Then
                        vldblTotalConceptosCargos = vldblTotalConceptosCargos + rsDatosFactura!Importe
                        vldblTotalConceptosCargosConDescuento = vldblTotalConceptosCargosConDescuento + rsDatosFactura!Importe - rsDatosFactura!Descuento
                        vldblTotalIVAConceptosCargos = vldblTotalIVAConceptosCargos + rsDatosFactura!IVA
                        If rsDatosFactura!IVA <> 0 Then
                            vldblTotalConceptosCargosSIgrava = vldblTotalConceptosCargosSIgrava + rsDatosFactura!Importe
                            vldblTotalConceptosCargosConDescuentoSIgrava = vldblTotalConceptosCargosConDescuentoSIgrava + rsDatosFactura!Importe - rsDatosFactura!Descuento
                        Else
                            vldblTotalConceptosCargosNOgrava = vldblTotalConceptosCargosNOgrava + rsDatosFactura!Importe
                            vldblTotalConceptosCargosConDescuentoNOgrava = vldblTotalConceptosCargosConDescuentoNOgrava + rsDatosFactura!Importe - rsDatosFactura!Descuento
                        End If
                    ElseIf rsDatosFactura!chrTipo = "OD" Then
                        vldblTotalConceptosSeguros = vldblTotalConceptosSeguros + rsDatosFactura!Importe
                        vldblTotalIVAConceptosSeguros = vldblTotalIVAConceptosSeguros + rsDatosFactura!IVA
                        If rsDatosFactura!IVA <> 0 Then
                            If rsDatosFactura!mnyCantidadGravada <> 0 Then
                                vldblTotalConceptosSegurosSIgrava = vldblTotalConceptosSegurosSIgrava + rsDatosFactura!mnyCantidadGravada
                                vldblTotalConceptosSegurosNOgrava = vldblTotalConceptosSegurosNOgrava + (rsDatosFactura!Importe - rsDatosFactura!mnyCantidadGravada)
                            Else
                                vldblTotalConceptosSegurosSIgrava = vldblTotalConceptosSegurosSIgrava + rsDatosFactura!Importe
                            End If
                        Else
                            vldblTotalConceptosSegurosNOgrava = vldblTotalConceptosSegurosNOgrava + rsDatosFactura!Importe
                        End If
                    End If
                    rsDatosFactura.MoveNext
                Loop
            
                If vldblTotalConceptosCargos = 0 Then
                    vldblPorcentajeFactura = 1
                Else
                    vldblPorcentajeFactura = 1 - (vldblTotalConceptosSeguros / vldblTotalConceptosCargos)
                End If
                If vldblTotalConceptosCargosConDescuento = 0 Then
                    vldblPorcentajeFacturaConDescuento = 1
                Else
                    vldblPorcentajeFacturaConDescuento = 1 - (vldblTotalConceptosSeguros / vldblTotalConceptosCargosConDescuento)
                End If
                If vldblTotalIVAConceptosCargos = 0 Then
                    vldblPorcentajeIVAFactura = 1
                Else
                    vldblPorcentajeIVAFactura = 1 - (vldblTotalIVAConceptosSeguros / vldblTotalIVAConceptosCargos)
                End If
                
                If vldblTotalConceptosCargosSIgrava = 0 Then
                    vldblPorcentajeFacturaSIgrava = 1
                Else
                    vldblPorcentajeFacturaSIgrava = 1 - (vldblTotalConceptosSegurosSIgrava / vldblTotalConceptosCargosSIgrava)
                End If
                
                If vldblTotalConceptosCargosNOgrava = 0 Then
                    vldblPorcentajeFacturaNOgrava = 1
                Else
                    vldblPorcentajeFacturaNOgrava = 1 - (vldblTotalConceptosSegurosNOgrava / vldblTotalConceptosCargosNOgrava)
                End If
                
                If vldblTotalConceptosCargosConDescuentoSIgrava = 0 Then
                    vldblPorcentajeFacturaConDescuentoSIgrava = 1
                Else
                    vldblPorcentajeFacturaConDescuentoSIgrava = 1 - (vldblTotalConceptosSegurosSIgrava / vldblTotalConceptosCargosConDescuentoSIgrava)
                End If
                
                If vldblTotalConceptosCargosConDescuentoNOgrava = 0 Then
                    vldblPorcentajeFacturaConDescuentoNOgrava = 1
                Else
                    vldblPorcentajeFacturaConDescuentoNOgrava = 1 - (vldblTotalConceptosSegurosNOgrava / vldblTotalConceptosCargosConDescuentoNOgrava)
                End If
                
                '----------------------------------------------------------'
            
                rsDatosFactura.MoveFirst
                llngConceptos = 0
                vldblFacturaIVA = 0
                Do While Not rsDatosFactura.EOF
                    If rsDatosFactura!chrTipo = "NO" Or rsDatosFactura!chrTipo = "OC" Then
                        If Trim(grdFactura.TextMatrix(1, 1)) = "" Then
                            grdFactura.Row = 1
                        Else
                            grdFactura.Rows = grdFactura.Rows + 1
                            grdFactura.Row = grdFactura.Rows - 1
                        End If
                        'vldblFacturaIVA = ((rsDatosFactura!Importe * vldblPorcentajeFactura) - rsDatosFactura!Descuento) * (rsDatosFactura!IVAConcepto / 100)
                        vldblFacturaIVA = 0
                        
                        vldblProporcionalSeguros = (rsDatosFactura!Importe / IIf(rsDatosFactura!IVA <> 0, vldblTotalConceptosCargosSIgrava, vldblTotalConceptosCargosNOgrava)) * IIf(rsDatosFactura!IVA <> 0, vldblTotalConceptosSegurosSIgrava, vldblTotalConceptosSegurosNOgrava)
                        vldblDescEspecial = (rsDatosFactura!Importe - vldblProporcionalSeguros - rsDatosFactura!Descuento) * (RsDescuentoEspecial!PORCDESCESPECIAL / 100)
                        
                        If rsDatosFactura!IVA <> 0 Then
                            'vldblFacturaIVA = ((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaSIgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaSIgrava)) - rsDatosFactura!Descuento) * (rsDatosFactura!IVAConcepto / 100)
                            'vldblFacturaIVA = ((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaSIgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaSIgrava)) - rsDatosFactura!Descuento)
                            ''''vldblFacturaIVA = rsDatosFactura!IVA * vldblPorcentajeIVAFactura
                            If rsDatosFactura!chrTipo = "OC" And rsDatosFactura!mnyCantidadGravada <> 0 Then
                                vldblFacturaIVA = rsDatosFactura!IVA
                            Else
                                vldblFacturaIVA = (rsDatosFactura!Importe - rsDatosFactura!Descuento - vldblProporcionalSeguros - vldblDescEspecial) * (rsDatosFactura!IVACONCEPTO / 100)
                            End If
                        End If
                        grdFactura.TextMatrix(grdFactura.Row, 1) = rsDatosFactura!Concepto
                        If rsDatosFactura!IVA = 0 Then
                            'grdFactura.TextMatrix(grdFactura.Row, 2) = FormatCurrency(rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaNOgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaNOgrava), 2)
                            grdFactura.TextMatrix(grdFactura.Row, 2) = FormatCurrency(rsDatosFactura!Importe * vldblPorcentajeFacturaNOgrava, 2)
                        Else
                            'grdFactura.TextMatrix(grdFactura.Row, 2) = FormatCurrency(rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaSIgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaSIgrava), 2)
                            grdFactura.TextMatrix(grdFactura.Row, 2) = FormatCurrency(rsDatosFactura!Importe * vldblPorcentajeFacturaSIgrava, 2)
                        End If
                        'grdFactura.TextMatrix(grdFactura.Row, 2) = FormatCurrency(rsDatosFactura!Importe * vldblPorcentajeFactura, 2)
                        grdFactura.TextMatrix(grdFactura.Row, 3) = FormatCurrency(rsDatosFactura!Descuento + FormatCurrency(vldblDescEspecial, 2), 2) 'FormatCurrency(((rsDatosFactura!Importe * IIf(rsDatosFactura!IVA = 0, vldblPorcentajeFacturaNOgrava, vldblPorcentajeFacturaSIgrava)) - rsDatosFactura!Descuento) * (rsDescuentoEspecial!PORCDESCESPECIAL / 100), 2), 2) 'FormatCurrency(rsDatosFactura!Descuento, 2)
                    'grdFactura.TextMatrix(grdFactura.Row, 4) = FormatCurrency(rsDatosFactura!IVA * vldblPorcentajeIVAFactura, 2)
                        grdFactura.TextMatrix(grdFactura.Row, 4) = FormatCurrency(vldblFacturaIVA, 2)
                    'grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFactura) - rsDatosFactura!Descuento + (rsDatosFactura!IVA * vldblPorcentajeIVAFactura), 2)
                    'grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((grdFactura.TextMatrix(grdFactura.Row, 2) - grdFactura.TextMatrix(grdFactura.Row, 3) + grdFactura.TextMatrix(grdFactura.Row, 4)), 2)
                        'grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFactura) - rsDatosFactura!Descuento + vldblFacturaIVA, 2)
                        If rsDatosFactura!IVA = 0 Then
                            'grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaNOgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaNOgrava)) - rsDatosFactura!Descuento, 2)
                            grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((grdFactura.TextMatrix(grdFactura.Row, 2) - grdFactura.TextMatrix(grdFactura.Row, 3)), 2)
                        Else
                            'grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaSIgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaSIgrava)) - rsDatosFactura!Descuento + vldblFacturaIVA, 2)
                            grdFactura.TextMatrix(grdFactura.Row, 5) = FormatCurrency((grdFactura.TextMatrix(grdFactura.Row, 2) - grdFactura.TextMatrix(grdFactura.Row, 3) + grdFactura.TextMatrix(grdFactura.Row, 4)), 2)
                        End If
                        grdFactura.TextMatrix(grdFactura.Row, 6) = rsDatosFactura!smicveconcepto
                        grdFactura.TextMatrix(grdFactura.Row, 7) = IIf(IsNull(rsDatosFactura!CuentaIngresos), 0, rsDatosFactura!CuentaIngresos)
                        grdFactura.TextMatrix(grdFactura.Row, 8) = IIf(IsNull(rsDatosFactura!CuentaDescuentos), 0, rsDatosFactura!CuentaDescuentos)
                        grdFactura.TextMatrix(grdFactura.Row, 9) = 0
                        grdFactura.TextMatrix(grdFactura.Row, 10) = Format(rsDatosFactura!IVACONCEPTO, vlstrFormato)
                    'grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFactura) - rsDatosFactura!Descuento + (rsDatosFactura!IVA * vldblPorcentajeIVAFactura), 15)
                        'grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFactura) - rsDatosFactura!Descuento + vldblFacturaIVA, 15)
                        If rsDatosFactura!IVA = 0 Then
                            'grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaNOgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaNOgrava)) - rsDatosFactura!Descuento, 15)
                            grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFacturaNOgrava) - rsDatosFactura!Descuento, 15)
                        Else
                            'grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * IIf(vldblPorcentajeFacturaSIgrava = 1, vldblPorcentajeFactura, vldblPorcentajeFacturaSIgrava)) - rsDatosFactura!Descuento + vldblFacturaIVA, 15)
                            grdFactura.TextMatrix(grdFactura.Row, 11) = FormatCurrency((rsDatosFactura!Importe * vldblPorcentajeFacturaSIgrava) - rsDatosFactura!Descuento + vldblFacturaIVA, 15)
                        End If
                    'grdFactura.TextMatrix(grdFactura.Row, 12) = FormatCurrency(rsDatosFactura!IVA * vldblPorcentajeIVAFactura, 15)
                        grdFactura.TextMatrix(grdFactura.Row, 12) = FormatCurrency(vldblFacturaIVA, 15)
                        grdFactura.TextMatrix(grdFactura.Row, 13) = IIf(rsDatosFactura!BITPESOS = 1, 1, rsDatosFactura!TipoCambio)
                        grdFactura.TextMatrix(grdFactura.Row, 14) = vldblDescEspecial
                        llngConceptos = llngConceptos + 1
                    End If
                    rsDatosFactura.MoveNext
                Loop
            End If
            pConfiguraGridFactura
            grdFactura.Visible = True
            vlcentavo = False
            If OptMotivoNota(1).Value = True And llngConceptos > 0 Then
                txtCantidad.Text = CStr(fTotalFactura(grdFactura, vgstrParametrosSP))
                chkPorcentaje_Click
            ElseIf OptMotivoNota(0).Value = True And llngConceptos > 0 Then
                vldblTotalFactura = CStr(fTotalFactura(grdFactura, vgstrParametrosSP))
            ElseIf optNotaCargo.Value = True And llngConceptos > 0 Then
                vldblTotalFactura = CStr(fTotalFactura(grdFactura, vgstrParametrosSP))
            End If
        Else
            txtCuenta.Text = ""
            txtNombrePaciente.Text = ""
            lblFechaFactura.Caption = ""
        End If

        rsDatosFactura.Close
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFactura_Click"))
    Unload Me
End Sub

Private Sub cboFactura_GotFocus()
On Error GoTo NotificaError
    
    If vlblnConsulta Then Exit Sub
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFactura_GotFocus"))
    Unload Me
End Sub

Private Sub cboFactura_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        grdFactura.Col = 1
        grdFactura.Row = 1
        grdFactura.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboFactura_KeyDown"))
    Unload Me
End Sub

Private Sub cmdAnteriorRegistro_Click()
    If grdBusqueda.Row > 1 Then
        grdBusqueda.Row = grdBusqueda.Row - 1
    End If
    
    pMostrarNota grdBusqueda.Row, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTipoNotaFACR)
    If optCliente.Value = True Then
        pHabilita True, True, True, True, True, False, IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) <> "A", False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "A", True, False)
    Else
        pHabilita True, True, True, True, True, False, IIf((grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "C" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC") Or (grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P" And Not fblnNotaAutomatica(lblFolio.Caption)), False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P", True, False)
    End If

End Sub

Private Sub cmdBuscar_Click()
On Error GoTo NotificaError
      
    sstObj.Tab = 1
    pEnfocaMkTexto mskFechaInicial
    
    sstCargos.Visible = False
   
    If fraTipoPaciente.Visible = True Then
        fraTipoPaciente.Visible = False
    End If
   
    If optPaciente.Value = True Then
        'Frame2.Height = 9410
        'Frame2.Top = 1600
        'grdBusqueda.Height = 9080
'        Frame3.Height = 1000
        lbClientePaciente.Caption = "Cuenta"
'        lbClientePaciente.Top = 685
'        txtNumCliente.Top = 610
'        txtNombreCliente.Top = 610
'        cmdCargarDatos.Top = 600
        fraConsultaPaciente.Visible = True
        OptConsultaTipoPaciente(2).Value = True
        txtNumCliente.Enabled = False
        optMostrarSolo(3).Enabled = False
        optMostrarSolo(4).Enabled = False
        'fraTipoPaciente.Visible = False
    Else
        'Frame2.Height = 9810
        'Frame2.Top = 1200
        'grdBusqueda.Height = 9480
'        Frame3.Height = 630
        lbClientePaciente.Caption = "Cliente"
'        lbClientePaciente.Top = 285
'        txtNumCliente.Top = 210
'        txtNombreCliente.Top = 210
'        cmdCargarDatos.Top = 200
        fraConsultaPaciente.Visible = False
        txtNumCliente.Enabled = True
        optMostrarSolo(3).Enabled = IIf(optNotaCargo.Value, True, False)
        optMostrarSolo(4).Enabled = IIf(optNotaCargo.Value, True, False)
    End If

    If vlblnConsulta = True Then
        freDetalleNota.Enabled = True
        lblUsoCFDI.Enabled = True
        lblMetodoPago.Enabled = True
        lblFormaPago.Enabled = True
        
        cboUsoCFDI.Enabled = True
        cboMetodoPago.Enabled = True
        cboFormaPago.Enabled = True

    Else
        freDetalleNota.Enabled = False
        lblUsoCFDI.Enabled = False
        lblMetodoPago.Enabled = False
        lblFormaPago.Enabled = False
        
        cboUsoCFDI.Enabled = False
        cboMetodoPago.Enabled = False
        cboFormaPago.Enabled = False

    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
    Unload Me
End Sub
Private Sub pConfiguraBusqueda()
On Error GoTo NotificaError
    With grdBusqueda
        .Clear
        .Cols = 23
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Fecha|Folio|Tipo|Número|Cliente|Estado|Motivo de la nota"
        .ColWidth(0) = 200
        .ColWidth(vlintColIdNota) = 0
        .ColWidth(vlintColPFacSAT) = 0 ' indica si la nota esta pendiente de cancelarse ante la sat
        .ColWidth(vlintcolPTimbre) = 0 ' indica si la nota esta pendiente de timbre fiscal
        .ColWidth(vlintColFechaNota) = 1100
        .ColWidth(vlintColFolioNota) = 1200
        .ColWidth(vlintColTipoNota) = 1500
        .ColWidth(vlintColNumCliente) = 1000
        .ColWidth(vlintColNombreCliente) = 3500
        .ColWidth(vlintColEstadoNota) = 1500
        .ColWidth(vlintColDomicilioCliente) = 0
        .ColWidth(vlintColRFCCliente) = 0
        .ColWidth(vlintColSubtotalNota) = 0
        .ColWidth(vlintColDescuentoNota) = 0
        .ColWidth(vlintColIVANota) = 0
        .ColWidth(vlintColTotalNota) = 0
        .ColWidth(vlintColchrTipo) = 0
        .ColWidth(vlintColchrEstatus) = 0
        .ColWidth(vlintColintNumPoliza) = 0
        .ColWidth(vlintColdtmFechaRegistro) = 0
        .ColWidth(vlintColCuentaContable) = 0
        .ColWidth(vlintColTipoNotaFACR) = 0
        .ColWidth(vlintColMotivoNota) = 2200
        
        .ColAlignmentFixed(vlintColFechaNota) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColFolioNota) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColTipoNota) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColNumCliente) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColNombreCliente) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColEstadoNota) = flexAlignCenterCenter
        .ColAlignmentFixed(vlintColMotivoNota) = flexAlignCenterCenter
   
        .ColAlignment(vlintColFechaNota) = flexAlignLeftCenter
        .ColAlignment(vlintColFolioNota) = flexAlignLeftCenter
        .ColAlignment(vlintColTipoNota) = flexAlignLeftCenter
        .ColAlignment(vlintColNumCliente) = flexAlignRightCenter
        .ColAlignment(vlintColNombreCliente) = flexAlignLeftCenter
        .ColAlignment(vlintColEstadoNota) = flexAlignLeftCenter
        .ColAlignmentFixed(vlintColMotivoNota) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraBusqueda"))
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo NotificaError
    
    Dim rsNotasFacturas As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    Dim vllngNumPoliza  As Long
    Dim vllngResultado As Long
    Dim vllngContador As Long
    Dim vllngDetallePoliza As Long
    Dim vlstrEstadoNota As String
    Dim vlstrSentencia As String
    Dim vldblCantidad As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim dblTotalIVA As Double 'para hacer el movimiento a IVA no cobrado
    Dim intMensaje As Integer 'Mensaje que regresa la función fintErrorCancelar()
    Dim intNumeroCuenta As Long
    Dim rsDetalleNotaElectronica As ADODB.Recordset

    If Not vgblnCFDI Then 'Agregado para caso 7994
        'Se revisa el parametro del tipo de cancelación de documento
        vlstrSentencia = "select chrTipoCancel from ccTipoCancelacion where smiCveDepartamento = " & vgintNumeroDepartamento & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount > 0 Then 'Si tiene configurado el parametro seleccionar la opción indicada
            If Trim(rs!chrTipoCancel) = "DOCUMENTO" Then
                frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio), txtCliente.Text, "DOCUMENTO", True, True, False
            ElseIf Trim(rs!chrTipoCancel) = "ACTUAL" Then
                frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio), txtCliente.Text, "ACTUAL", True, True, False
            ElseIf Trim(rs!chrTipoCancel) = "ELEGIR" Then
                frmFechaCancelacionNotas.Show vbModal
            End If
        Else 'Si no tiene configurado el parametro, realizar la cancelación por default (Cancelación a la fecha del documento)
            frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio), txtCliente.Text, "DOCUMENTO", True, True, False
        End If
    Else 'Si la nota es CFDi y el PAC es PAX, cancelar con la fecha actual
        'Modificado para el caso 8886, para incluir la cancelación por Buzón
        frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio), txtCliente.Text, "ACTUAL", True, True, False
    End If

    InicializaComponentes

    If optNotaCargo.Value Then
        optNotaCargo.SetFocus
    Else
        optNotaCredito.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
    Unload Me
End Sub

Private Sub InicializaComponentes()
On Error GoTo NotificaError

    pLimpiaTodo
    pHabilita False, False, True, False, False, False, False, False
    
    optMostrarSolo(0).Value = True
    
    freDetalleNota.Height = 2940
    
    If optCliente.Value Then
        freDetalleNota.Top = 7670
    Else
        freDetalleNota.Top = 7720
    End If
    
    grdNotas.Height = 1200
    
    lblObservaciones.Top = 1560
    
    If optCliente.Value Then
        FraMetodoForma.Top = 2250
    Else
        FraMetodoForma.Top = 2600
    End If
    
    If optCliente.Value Then
        txtComentario.Top = 9500
    Else
        txtComentario.Top = 9500 '9350
    End If
    
    txtSubtotal.Top = 1515
    txtDescuentoTot.Top = 1860
    txtIVA.Top = 2205
    txtTotal.Top = 2550

    lblSubtotal.Top = 1560
    lblDescuento.Top = 1920
    lblIVA.Top = 2265
    lblTotal.Top = 2610
    
    txtComentario.Locked = False
    
    sstObj.Tab = 0
    
    If optCliente.Value = True Then
        sstFacturasCreditos.Visible = True
        fraMotivoNota.Enabled = True
    Else
        sstCargos.Visible = True
        fraTipoPaciente.Enabled = True
        fraMotivoNota.Enabled = True 'Caso 7374: Modificado para permitir aplicar notas a pacientes por "Error de facturación"
        grdNotas.ColWidth(vlintColFactura) = 0
        fraTipoPaciente.Visible = True
    End If
    
    If optNotaCredito.Value = True Then
        If optPaciente.Value = True Then
            OptMotivoNota(0).Value = False
            OptMotivoNota(1).Value = True
        Else
            OptMotivoNota(0).Value = True
            OptMotivoNota(1).Value = False
        End If
    Else
        fraMotivoNota.Enabled = False
        OptMotivoNota(0).Value = False
        OptMotivoNota(1).Value = False
    End If
        
    optCliente.Enabled = True
    optPaciente.Enabled = True
    
    optNotaCredito.Enabled = True
    optNotaCargo.Enabled = True
    
    If optPaciente.Value Then
        optNotaCargo.Enabled = False
    Else
        optNotaCargo.Enabled = True
    End If
    
    cmdComprobante.Enabled = False
    Me.cmdConfirmarTimbre.Enabled = False
    Me.txtPendienteTimbre.Visible = False
    blnNoMensaje = True
    blnNoMensaje = False
    cmdAddenda.Enabled = False
    vglngCveAddenda = 0
    vgstrTipoPacienteAddenda = ""
    vglngCuentaPacienteAddenda = 0
    vglngCveEmpresaCliente = 0
    
    '- (CR) Agregado para caso 7249 -'
    vgblnFacturaDirecta = False
    sstElementos.TabEnabled(0) = True
    sstElementos.TabEnabled(1) = True
    sstElementos.TabEnabled(2) = True
    sstElementos.TabEnabled(3) = True
    sstElementos.TabEnabled(4) = True
    '-------------------------------'
    
    Me.mskFechaInicial.Enabled = True
    Me.mskFechaFinal.Enabled = True
    Me.txtNumCliente.Enabled = True
    Me.txtNombreCliente.Enabled = True
    Me.cmdCargarDatos.Enabled = True
    Me.OptConsultaTipoPaciente(0).Enabled = True
    Me.OptConsultaTipoPaciente(1).Enabled = True
    Me.OptConsultaTipoPaciente(2).Enabled = True
    
    FraMetodoForma.Visible = True
    cboUsoCFDI.Visible = True
    cboMetodoPago.Visible = True
    cboFormaPago.Visible = True
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":InicializaComponentes"))
    Unload Me
End Sub

Private Function fblnDatosValidos() As Boolean
    Dim rsValidaPago As New ADODB.Recordset
    Dim rsValidaCuenta As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vlngNumDepto As Long
    Dim vldblDiferenciaIVA As Double
    
    Dim vldblSaldoActual As Double
    Dim rsReporte As New ADODB.Recordset
    Dim vllngNumCliente As Long
    Dim vllngDeptoCliente As Long
    Dim vllngCuentaContableCredito As Long
    Dim vldblLimiteCredito As Double
    Dim vlstrsSQL As String
    Dim rsCuentaContableCredito As New ADODB.Recordset

On Error GoTo NotificaError

    fblnDatosValidos = True

    '----------------------------------'
    ' Si tiene permisos de "Escritura" '
    '----------------------------------'
    Select Case cgstrModulo
        Case "CC"
            If Not fblnRevisaPermiso(vglngNumeroLogin, 634, "E") Then
                fblnDatosValidos = False
            End If
                
        Case "PV"
            If Not fblnRevisaPermiso(vglngNumeroLogin, 2296, "E") Then
                fblnDatosValidos = False
            End If
    End Select
    
    '------------------'
    ' Fecha de la nota '
    '------------------'
    If fblnDatosValidos Then
        If Not IsDate(mskFecha.Text) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            pEnfocaMkTexto mskFecha
            fblnDatosValidos = False
        Else
            If CDate(mskFecha.Text) > fdtmServerFecha Then
                '¡La fecha debe ser menor o igual a la del sistema!
                MsgBox SIHOMsg(40), vbOKOnly + vbInformation, "Mensaje"
                pEnfocaMkTexto mskFecha
                fblnDatosValidos = False
            End If
        End If
    End If
    
    '---------------------------------------------------------------------'
    ' Validación para que la nota de crédito no exceda la cantidad pagada '
    '---------------------------------------------------------------------'
    If fblnDatosValidos And optNotaCredito.Value Then
        vlintContador = 0
        Do While vlintContador <= UBound(aFacturas(), 1) And fblnDatosValidos
            If aFacturas(vlintContador).vldblSubtotal <> 0 Or aFacturas(vlintContador).vldblDescuento <> 0 Or aFacturas(vlintContador).vldblIVA <> 0 Then
                vlstrSentencia = "SELECT intnummovimiento, mnycantidadCredito, mnycantidadPagada FROM ccmovimientocredito "
                vlstrSentencia = vlstrSentencia & " WHERE chrtiporeferencia = 'FA' AND chrfolioreferencia = '" & Trim(aFacturas(vlintContador).vlstrFolioFactura) & "'"
                If optCliente.Value = True Then '(CR) - Para que se distingan los créditos de los clientes
                    vlstrSentencia = vlstrSentencia & " AND intNumCliente = " & Trim(txtCveCliente.Text)
                End If
                Set rsValidaPago = frsRegresaRs(vlstrSentencia)
                If rsValidaPago.RecordCount > 0 And chkFacturasPagadas.Value = 0 Then
                    'If rsValidaPago!mnyCantidadPagada + (aFacturas(vlintContador).vldblSubtotal - aFacturas(vlintContador).vldblDescuento + Val(Format(aFacturas(vlintContador).vldblIVA, vlstrFormato))) > rsValidaPago!mnyCantidadCredito Then
                    'If rsValidaPago!mnyCantidadPagada + (Val(Format(aFacturas(vlintContador).vldblSubtotal, vlstrFormato)) - aFacturas(vlintContador).vldblDescuento + Val(Format(Round(aFacturas(vlintContador).vldblIVA, 2), vlstrFormato))) > rsValidaPago!mnyCantidadCredito Then
                     If rsValidaPago!mnyCantidadPagada + Val(Format(((aFacturas(vlintContador).vldblSubtotal * aFacturas(vlintContador).vldblTipoCambio) - (aFacturas(vlintContador).vldblDescuento * aFacturas(vlintContador).vldblTipoCambio) + (aFacturas(vlintContador).vldblIVA * aFacturas(vlintContador).vldblTipoCambio)), vlstrFormato)) > rsValidaPago!mnyCantidadCredito Then
                        MsgBox SIHOMsg(689) + ": " + Trim(aFacturas(vlintContador).vlstrFolioFactura), vbExclamation, "Mensaje"
                        fblnDatosValidos = False
                    End If
                End If
                rsValidaPago.Close
            End If
            vlintContador = vlintContador + 1
        Loop
    End If
    
    '-----------------------------------'
    ' Que haya movimientos para guardar '
    '-----------------------------------'
    If fblnDatosValidos And Trim(grdNotas.TextMatrix(1, 2)) = "" Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        fblnDatosValidos = False
    End If
    
    '-------------------------------'
    ' Tipo de cambio de los DOLARES '
    '-------------------------------'
    If fblnDatosValidos Then
        vldblTipoCambio = fdblTipoCambio(fdtmServerFecha, "V") 'Tipo de cambio a la Venta
        If vldblTipoCambio = 0 Then
            'No está registrado el tipo de cambio del día.
            MsgBox SIHOMsg(231), vbCritical, "Mensaje"
            fblnDatosValidos = False
        End If
    End If
    
    '----------------------------------------------------------------------'
    ' Que la fecha de la nota sea siempre mayor a la fecha de las facturas '
    '----------------------------------------------------------------------'
    If fblnDatosValidos Then
        vlintContador = 0
        Do While vlintContador <= UBound(aFacturas(), 1) And fblnDatosValidos
            If aFacturas(vlintContador).vldblSubtotal <> 0 Or aFacturas(vlintContador).vldblDescuento <> 0 Or aFacturas(vlintContador).vldblIVA <> 0 Then
                If aFacturas(vlintContador).vldtmFecha > CDate(mskFecha.Text) Then
                    If optCliente.Value = True Then
                        'No se puede incluir en la nota una factura cuya fecha es mayor a la fecha de la nota.
                        MsgBox SIHOMsg(730) & " " & aFacturas(vlintContador).vlstrFolioFactura, vbOKOnly + vbExclamation, "Mensaje"
                    Else
                        'No se puede incluir en la nota un cargo cuya fecha es mayor a la fecha de la nota.
                        MsgBox SIHOMsg(1176) & " " & aFacturas(vlintContador).vlstrFolioFactura, vbOKOnly + vbExclamation, "Mensaje"
                    End If
                    
                    fblnDatosValidos = False
                End If
            End If
            vlintContador = vlintContador + 1
        Loop
    End If
    
    '--------------------------------------------------------------------------------------------'
    ' Que la fecha de la nota sea siempre mayor a la fecha de apertura de la cuenta del paciente '
    '--------------------------------------------------------------------------------------------'
    If fblnDatosValidos Then
        If optPaciente.Value Then
            If vgdtmFechaIngreso > CDate(mskFecha.Text) + fdtmServerHora Then
                'No se puede aplicar una nota de crédito con una fecha menor a la fecha de apertura de la cuenta
                MsgBox SIHOMsg(1003), vbOKOnly + vbExclamation, "Mensaje"
                fblnDatosValidos = False
            End If
        End If
    End If
    
    '---------------------------------------------------------------------------'
    ' Validar que existan las cuentas contables de cada concepto de facturación '
    '---------------------------------------------------------------------------'
'    If fblnDatosValidos Then
'        If optNotaCredito.Value = True And OptMotivoNota(1).Value = True And sstFacturasCreditos.Tab = 0 Then
'            For vlintContador = 1 To grdNotas.Rows - 1
'                '- Valida la cuenta de ingresos -'
'                If Val(grdNotas.TextMatrix(vlintContador, vlintColCtaIngresos)) = 0 Then
'                    'No se encontró la cuenta contable
'                    MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto '" & Trim(grdNotas.TextMatrix(vlintContador, vlintColDescripcionConcepto)) & "'.", vbOKOnly + vbInformation, "Mensaje"
'                    fblnDatosValidos = False
'                    Exit For
'                Else
'                    '- Valida si existe cuenta de descuento por nota de crédito en conceptos de facturación -'
'                    vlstrSentencia = "SELECT NVL(intNumCtaDescNota, 0) NumCtaDescNota " & _
'                                     " FROM PvConceptoFacturacionEmpresa " & _
'                                     " WHERE intCveConceptoFactura = " & Str(Val(grdNotas.TextMatrix(vlintContador, vlintColCveConcepto))) & _
'                                     " AND intCveEmpresaContable = " & vgintClaveEmpresaContable
'                    Set rsValidaCuenta = frsRegresaRs(vlstrSentencia)
'                    If rsValidaCuenta.RecordCount <> 0 Then
'                        If rsValidaCuenta!NumCtaDescNota = 0 Then
'                            fblnDatosValidos = False
'                        End If
'                    Else
'                        fblnDatosValidos = False
'                    End If
'
'                    '- Valida si existe cuenta de descuento por nota de crédito en excepciones contables -'
'                    If Not fblnDatosValidos And optPaciente.Value = True Then
'                        If Val(grdNotas.TextMatrix(vlintContador, 13)) = 0 Then
'                            vlngNumDepto = Val(grdNotas.TextMatrix(vlintContador, 14))
'                        Else
'                            vlngNumDepto = 1
'                            'Departamento que solicita requisición
'                            frsEjecuta_SP Val(grdNotas.TextMatrix(vlintContador, 13)), "FN_CCSELDEPTOSOLICITAREQ", True, vlngNumDepto
'                            If vlngNumDepto = 0 Then
'                                vlngNumDepto = Val(grdNotas.TextMatrix(vlintContador, 14))
'                            End If
'                        End If
'
'                        vlstrSentencia = "SELECT NVL(intNumCuentaDescNota, 0) NumCtaDescNota " & _
'                                         " FROM PvConceptoFacturacionDepartame " & _
'                                         " WHERE smiCveConcepto = " & Str(Val(grdNotas.TextMatrix(vlintContador, vlintColCveConcepto))) & _
'                                         " AND smiCveDepartamento = " & vlngNumDepto
'                        Set rsValidaCuenta = frsRegresaRs(vlstrSentencia)
'                        If rsValidaCuenta.RecordCount <> 0 Then
'                            If rsValidaCuenta!NumCtaDescNota = 0 Then
'                                fblnDatosValidos = False
'                            Else
'                                fblnDatosValidos = True
'                            End If
'                        Else
'                            fblnDatosValidos = False
'                        End If
'                    End If
'
'                    If Not fblnDatosValidos Then
'                        'No existe cuenta contable para el concepto de facturación:
'                        MsgBox SIHOMsg(907) & " '" & Trim(grdNotas.TextMatrix(vlintContador, vlintColDescripcionConcepto)) & "' asignada para descuento por nota de crédito.", vbOKOnly + vbInformation, "Mensaje"
'                        Exit For
'                    End If
'                End If
'            Next vlintContador
'        End If
'    End If
    
    '-------------------------------------------------------------------------------------------'
    ' Validar que el total de IVA de la nota, no sea mayor a la proporción de IVA de la factura '
    '-------------------------------------------------------------------------------------------'
'    If fblnDatosValidos And UBound(aFacturas(), 1) = 0 And cboFactura.List(cboFactura.ListIndex) <> "" And optCliente.Value = True And optNotaCredito.Value = True And OptMotivoNota(1).Value = True Then
'        vldblDiferenciaIVA = Val(Format(txtIVA.Text, vlstrFormato)) - Format((vldblTotalIVAConceptosCargos - vldblTotalIVAConceptosSeguros) * IIf(chkPorcentaje.Value = 1, Val(Format(txtCantidad.Text, vlstrFormato)) / 100, (Val(Format(txtCantidad.Text, vlstrFormato)) / fTotalFactura(grdFactura, Trim(cboFactura.List(cboFactura.ListIndex))))), vlstrFormato)
'        If Format(vldblDiferenciaIVA, vlstrFormato) >= 0.04 Then
'            fblnDatosValidos = False
'            MsgBox Replace(Replace(SIHOMsg(1134), "las notas de crédito", "la nota"), "verifique", "correspondiente al total de la nota, verifique"), vbExclamation, "Mensaje"
'        End If
'    End If
    
    'Revisa si con el importe de la nota de cargo no se excede el limite de crédito del cliente
    If fblnDatosValidos Then
        If optCliente.Value = True And optNotaCargo.Value = True Then
            vlstrsSQL = "select intnumcliente, SMICVEDEPARTAMENTO, intNumCuentaContable, mnyLimiteCredito " & _
                        " from CcCliente " & _
                        " Where INTNUMCLIENTE = " & txtCveCliente.Text
            Set rsCuentaContableCredito = frsRegresaRs(vlstrsSQL)
            If rsCuentaContableCredito.RecordCount <> 0 Then
                vllngNumCliente = rsCuentaContableCredito!intNumCliente
                vllngDeptoCliente = rsCuentaContableCredito!smicvedepartamento
                vllngCuentaContableCredito = rsCuentaContableCredito!intnumcuentacontable
                vldblLimiteCredito = rsCuentaContableCredito!mnyLimiteCredito
            End If
            
            ' Determina el saldo actual
            vldblSaldoActual = 0
            vgstrParametrosSP = "*" & "|" & "-1" & "|" & txtCveCliente.Text & "|" & txtCveCliente.Text & "|" & fstrFechaSQL(Trim(fdtmServerFecha)) & "|" & "0" & "|" & "0" & "|" & "0" & "|" & vllngDeptoCliente & "|" & vgintClaveEmpresaContable
            Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_Ccrptantiguedadsaldos")
            If rsReporte.RecordCount <> 0 Then
                Do While Not rsReporte.EOF
                    vldblSaldoActual = vldblSaldoActual + rsReporte!Saldo
                    rsReporte.MoveNext
                Loop
            Else
                vldblSaldoActual = 0
            End If
            
            If (Val(Format(Trim(txtTotal.Text), "############.##")) + vldblSaldoActual) > vldblLimiteCredito And vldblLimiteCredito > 0 Then
                'La cantidad capturada más el saldo actual del cliente excede el límite de crédito otorgado.
                MsgBox SIHOMsg(734), vbOKOnly + vbExclamation, "Mensaje"
        
                fblnDatosValidos = False
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        If Val(Format(txtIVA.Text, vlstrFormato)) <> 0 Then
            If glngCtaIVACobrado = 0 Or glngCtaIVANoCobrado = 0 Then
                fblnDatosValidos = False
                'No se encuentran registradas las cuentas de IVA cobrado y no cobrado en los parámetros generales del sistema.
                MsgBox SIHOMsg(729), vbOKOnly + vbExclamation, "Mensaje"
            ElseIf Not fblnCuentaAfectable(fstrCuentaContable(glngCtaIVACobrado), vgintClaveEmpresaContable) Then
                fblnDatosValidos = False
                MsgBox SIHOMsg(825) & vbCrLf & vbCrLf & fstrCuentaContable(glngCtaIVACobrado) & "  " & fstrNombreCuentaContable(glngCtaIVACobrado), vbOKOnly + vbExclamation, "Mensaje"
            ElseIf Not fblnCuentaAfectable(fstrCuentaContable(glngCtaIVANoCobrado), vgintClaveEmpresaContable) Then
                fblnDatosValidos = False
                MsgBox SIHOMsg(825) & vbCrLf & vbCrLf & fstrCuentaContable(glngCtaIVANoCobrado) & "  " & fstrNombreCuentaContable(glngCtaIVANoCobrado), vbOKOnly + vbExclamation, "Mensaje"
            End If
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
    Unload Me
End Function

Private Sub ChecarCentavosPoliza(vllngNoPoliza As Long)
On Error GoTo NotificaError
    Dim rsCcDetallePoliza As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlIntCent As Integer
    Dim vlStrCent As String
    Dim X As Integer
    
    
    vlstrSentencia = "select INTNUMEROREGISTRO, mnycantidadmovimiento, " & _
                        " (select sum(mnycantidadmovimiento)from CnDetallePoliza where intNumeroPoliza = " & vllngNoPoliza & " and bitnaturalezamovimiento = 1) suma  " & _
                        " from CnDetallePoliza where intNumeroPoliza = " & vllngNoPoliza & " and bitnaturalezamovimiento = 1 "
    Set rsCcDetallePoliza = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsCcDetallePoliza.RecordCount > 0 Then
        If Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) >= -0.03 And Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) <= 0.03 Then
            'MsgBox "- " & CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)), vbOKOnly + vbInformation, "Mensaje"
            vlIntCent = 0
            If CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = ".01" Or CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = "-.01" Then
                vlIntCent = 1
            ElseIf CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = ".02" Or CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = "-.02" Then
                    vlIntCent = 2
            ElseIf CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = ".03" Or CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)) = "-.03" Then
                    vlIntCent = 3
            End If
            If vlIntCent > 0 Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                If rsCcDetallePoliza.RecordCount = 1 Then
                    If Format((rsCcDetallePoliza!suma) < Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) Then
                        vlStrCent = "0.0" + CStr(vlIntCent)
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                         " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento + " & vlStrCent & _
                                         " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                    Else
                        'se agregó porque no estaba inicializada la variable y marcaba error - caso 20504
                        vlStrCent = "0.0" + CStr(vlIntCent)
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                         " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento - " & vlStrCent & _
                                         " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                    End If
                ElseIf rsCcDetallePoliza.RecordCount = 2 Then
                    If Format((rsCcDetallePoliza!suma) < Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) Then
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                         " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento + 0.01 " & _
                                         " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                        If vlIntCent > 1 Then
                            vlStrCent = "0.0" + CStr(vlIntCent - 1)
                            rsCcDetallePoliza.MoveNext
                            vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                             " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento + " & vlStrCent & _
                                             " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                            pEjecutaSentencia vlstrSentencia
                        End If
                    Else
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                         " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento - 0.01 " & _
                                         " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                        If vlIntCent > 1 Then
                            vlStrCent = "0.0" + CStr(vlIntCent - 1)
                            rsCcDetallePoliza.MoveNext
                            vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                             " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento - " & vlStrCent & _
                                             " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                            pEjecutaSentencia vlstrSentencia
                        End If
                    End If
                ElseIf rsCcDetallePoliza.RecordCount >= 3 Then
                    If Format((rsCcDetallePoliza!suma) < Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) Then
                        For X = 1 To vlIntCent
                            'MsgBox "- " & CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)), vbOKOnly + vbInformation, "Mensaje"
                            vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                             " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento + 0.01 " & _
                                             " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                            pEjecutaSentencia vlstrSentencia
                            If vlIntCent < 4 Then
                                rsCcDetallePoliza.MoveNext
                            End If
                        Next X
                    Else
                        For X = 1 To vlIntCent
                            'MsgBox "- " & CStr(Format((rsCcDetallePoliza!suma) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)), vbOKOnly + vbInformation, "Mensaje"
                            vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                             " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento - 0.01 " & _
                                             " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                            pEjecutaSentencia vlstrSentencia
                            If vlIntCent < 4 Then
                                rsCcDetallePoliza.MoveNext
                            End If
                        Next X
                    End If
                End If
                
                rsCcDetallePoliza.Close
                vlStrCent = "0.0" + CStr(vlIntCent)
                vlstrSentencia = "select INTNUMEROREGISTRO, mnycantidadmovimiento, " & _
                            " (select sum(mnycantidadmovimiento)from CnDetallePoliza where intNumeroPoliza = " & vllngNoPoliza & " and bitnaturalezamovimiento = 0) suma  " & _
                            " from CnDetallePoliza where intNumeroPoliza = " & vllngNoPoliza & " and bitnaturalezamovimiento = 0 "
                Set rsCcDetallePoliza = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                
                If rsCcDetallePoliza.RecordCount > 0 Then
                    If Format((rsCcDetallePoliza!suma) < Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) Then
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                                     " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento + " & vlStrCent & _
                                                     " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                    Else
                        vlstrSentencia = "UPDATE CnDetallePoliza " & _
                                                     " SET CnDetallePoliza.mnycantidadmovimiento = CnDetallePoliza.mnycantidadmovimiento - " & vlStrCent & _
                                                     " WHERE INTNUMEROREGISTRO = " & CStr(rsCcDetallePoliza!intNumeroRegistro)
                        pEjecutaSentencia vlstrSentencia
                    End If
                End If
                EntornoSIHO.ConeccionSIHO.CommitTrans
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":ChecarCentavosPoliza"))
    Unload Me

End Sub


Private Sub cmdGrabarRegistro_Click()
On Error GoTo NotificaError

    Dim rsCcNotaDetalle As New ADODB.Recordset   'Recordset dinámico para el detalle de la nota
    Dim rsCcNota As New ADODB.Recordset          'Para guardar la información maestro
    Dim rsCcNotaFactura As New ADODB.Recordset   'Para guardar la información de resumen de totales de la nota para la factura
    Dim rsNota As New ADODB.Recordset
    Dim vllngPersonaGraba As Long                'Persona que esta generando la factura
    Dim vllngSecuencia As Long                   'Consecutivo de CcNota
    Dim vllngContador As Long                    'Para ciclos
    Dim vllngFoliosFaltantes  As Long            'Para validar si existen folios para el documento
    Dim vllngResultado As Long                   'Resultado de la revisión si se está realizando algún cierre contable
    Dim vllngNumPoliza As Long                   'Consecutivo de la póliza para la nota
    Dim vllngDetallePoliza As Long               'Consecutivo del detalle de la póliza que fué insertado
    Dim vllngCuenta As Long 'Contador
    Dim strTotalLetras As String
    Dim vlstrFecha As String
    
    Dim vlstrSentencia As String                 'Para formar las instrucciones SQL
    
    Dim vldblCantidad As Double                  'Subtotal del concepto en la factura
    Dim vldblDescuento As Double                 'Descuento del concepto en la factura
    Dim vldblIVA As Double                       'IVA del concepto en la factura
    Dim vllngnumcredito As Long                  'Consecutivo del movimiento de crédito
    Dim intTotalFacturas As Integer              'Total de facturas incluidas en la nota
    Dim dblTotalIVA As Double                    'Total del IVA de la nota para hacer el cargo o abono a IVA cobrado y no cobrado
    'Dim vlstrFormatoLargo As String
    
    Dim strTipoReferencia As String
    Dim BlnBit As Boolean
    Dim intNumeroCuenta As Long
    Dim lngNumDepto As Long
    Dim intDepartamento As Integer
    Dim strParametros As String
    Dim lngSecuencia As Long                     'Consecutivo de CcNota para CFDs
    Dim vlAddendaValida As Long
    
    Dim vlintContador As Integer
    Dim chrTipoCancel As String
    Dim ObjRs As New ADODB.Recordset
    Dim rsDepartamentoConcepto As New ADODB.Recordset
    Dim vldblCantidadTotal As Double
    Dim vldblFactorDescuento As Double
    Dim rsSeleccionaCargos As New ADODB.Recordset
    Dim vllngCtaIngresos As Long
    Dim vllngCtaDescuentos As Long
    Dim vldblCantidadTotalDesc As Double
    'Dim vlstrCF As String
    'Dim rsCF As ADODB.Recordset
    Dim vldblCantidadTotalConcepto As Double
    Dim vldblPorcentajeDepartamento As Double
    Dim vldblCantidadConceptoNota As Double
    Dim vldblDescuentoConceptoNota As Double
    Dim vldblTipoCambio  As Double
    Dim intIndex As Integer
    Dim vl_strConceptoPoliza As String
    Dim vl_strConceptoDtallePoliza As String
    Dim vl_strReferenciaDetallePoliza As String
    Dim vl_strNotaTipo As String
    Dim vl_TipoCH As String
    Dim vllngConRefDetallePoliza As Integer
    
    intTipoEmisionComprobante = fintTipoEmisionComprobante(IIf(optNotaCargo.Value, "NA", "NC"), vllngNumeroTipoFormato, vgintFolioUnico)
    If intTipoEmisionComprobante = 0 Then
        Exit Sub
    End If
    
    If intTipoEmisionComprobante = 2 Then
        'Se revisa el tipo de CFD de la nota (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
        intTipoCFDNota = fintTipoCFD(IIf(optNotaCargo.Value, "NA", "NC"), vllngNumeroTipoFormato, vgintFolioUnico)
        
        'Si aparece un error terminar la transacción
        If intTipoCFDNota = 3 Then   'ERROR
            'Si es error, se cancela la transacción
            Exit Sub
        End If
    End If
          
    If Not fblnDatosValidos() Then Exit Sub
    
    ' Validación de la cuenta puente para notas de crédito a pacientes
    If optPaciente.Value = True And chkFacturasPaciente.Value = 0 Then
        intNumeroCuenta = 1
        strParametros = "INTNUMCUENTAPUENTEDESCTOSPACIENTE" & "|" & vgintClaveEmpresaContable
        ' Función que regresa el número de cuenta contable de la cuenta puente para pacientes
        frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
        
        If intNumeroCuenta = 0 Then
            ' No se ha registrado la "Cuenta puente para notas de crédito a pacientes"
            MsgBox SIHOMsg(1001), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    
    ' Validación de la cuenta puente para notas de crédito en facturas pagadas
    If chkFacturasPagadas.Value = 1 Or chkFacturasPaciente.Value = 1 Then
        intNumeroCuenta = 1
        strParametros = "INTNUMCPUENTEDESCTOSNOTASFACPAGADAS" & "|" & vgintClaveEmpresaContable
        ' Función que regresa el número de cuenta contable de la cuenta puente para pacientes
        frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
        
        If intNumeroCuenta = 0 Then
            ' No se ha registrado la "Cuenta puente para notas de crédito en facturas pagadas"
            MsgBox SIHOMsg(1079), vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
       
    'Validación del mensaje de addenda...
    If vgMostrarMsjAddenda = True Then
        'Se incluyó más de una factura en la nota de crédito por lo que no se generará CFD con addenda.  ¿Desea continuar?
        If MsgBox(SIHOMsg(1147), vbQuestion + vbYesNo, "Mensaje") = vbNo Then
            Exit Sub
        End If
    End If

    If optPaciente.Value = True And optNotaCredito.Value And OptMotivoNota(1).Value Then
        If chkFacturasPaciente.Value = 0 Then
            If OptTipoPaciente(0).Value Then 'Internos
                vgstrParametrosSP = txtCveCliente.Text & "|" & Str(vgintClaveEmpresaContable)
                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
            Else  'Externos
                vgstrParametrosSP = txtCveCliente.Text & "|" & Str(vgintClaveEmpresaContable)
                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
            End If
    
            If rs.RecordCount <> 0 Then
                If rs!Facturado = 1 Or rs!Facturado = True Then
                    'La cuenta ya ha sido facturada, no se pudo realizar ningún movimiento.
                    MsgBox SIHOMsg(299), vbOKOnly + vbInformation, "Mensaje"
                    grdCargos.Redraw = True
                    
                    InicializaComponentes
                    If optNotaCargo.Value Then
                       optNotaCargo.SetFocus
                    Else
                       optNotaCredito.SetFocus
                    End If
                    vlblnConsulta = False
                    vlblnSalir = False
                    cmdComprobante.Enabled = False
                    Exit Sub
                End If
            End If
        End If
    End If

    If optPaciente.Value = True And optNotaCredito.Value And OptMotivoNota(1).Value Then
        If chkFacturasPaciente.Value = 0 Then
            Set rsSeleccionaCargos = frsEjecuta_SP(txtCveCliente.Text & "|" & IIf(OptTipoPaciente(0), "I", "E") & "|" & 0 & "|" & "-1|N|0", "Sp_PvselcargospacienteNotas")
            If rsSeleccionaCargos.RecordCount = 0 Then
                MsgBox SIHOMsg(288), vbOKOnly + vbInformation, "Mensaje"
                grdCargos.Redraw = True
                
                InicializaComponentes
                If optNotaCargo.Value Then
                   optNotaCargo.SetFocus
                Else
                   optNotaCredito.SetFocus
                End If
                vlblnConsulta = False
                vlblnSalir = False
                cmdComprobante.Enabled = False
                Exit Sub
            End If
        End If
    End If
    
    If Not fblnValidaSAT() Then
        Exit Sub
    End If
    
    '-------------------'
    ' Persona que graba '
    '-------------------'
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    '--------------------------'
    ' Inicio de la transacción '
    '--------------------------'
    EntornoSIHO.ConeccionSIHO.BeginTrans

    If grdNotas.TextMatrix(1, 3) <> "" Then
        intLimiteAjusteIVA = 0
        BlnAjusteIVA = False
        IntAjusteIVAContador = 0
        Call pAjustaIVA(grdNotas)
    End If
    
    '--------------------------------------------------------------------'
    ' Revisar que no se esté haciendo un cierre contable en este momento '
    '--------------------------------------------------------------------'
    vllngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "Sp_CnUpdEstatusCierre", True, vllngResultado
    If vllngResultado <> 1 Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'En este momento se está realizando un cierre contable, espere un momento e intente de nuevo.
        MsgBox SIHOMsg(714), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    '----------------------------------------------'
    ' Revisar que el periodo contable esté abierto '
    '----------------------------------------------'
    If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(mskFecha.Text)), Month(CDate(mskFecha.Text))) Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'El periodo contable esta cerrado.
        MsgBox SIHOMsg(209), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    '----------------------------'
    ' Cargar el folio de la nota '
    '----------------------------'
    fstrFolioDocumento 1
    If Trim(vlstrFolioDocumento) = "0" Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        vlblnSalir = True
        Exit Sub
    End If
    
    vlstrSentencia = "SELECT * FROM CcNota WHERE intConsecutivo = -1"
    Set rsCcNota = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vlstrSentencia = "SELECT * FROM CcNotaFactura WHERE intConsecutivo = -1"
    Set rsCcNotaFactura = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vlstrSentencia = "SELECT * FROM CcNotaDetalle WHERE intConsecutivo = -1"
    Set rsCcNotaDetalle = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
     '-Aqui se asigna el concepto general de las notas de CREDITO O CARGO de la póliza que se genera
    vl_strConceptoPoliza = ""
    If optNotaCredito.Value Then                            'CREDITOS
        If optCliente Then vl_strNotaTipo = "X"             'Clientes
        If optPaciente Then vl_strNotaTipo = "2"            'Pacientes
    End If
    If optNotaCargo.Value Then vl_strNotaTipo = "W"         'CARGOS
    If fblnObtenerConfigPolizanva(0, vl_strNotaTipo, frmNotas.Name, optCliente) = True Then
        For intIndex = 0 To UBound(ConceptosPoliza, 1)
            Select Case ConceptosPoliza(intIndex).Tipo
                Case "1"
                    vl_strConceptoPoliza = vl_strConceptoPoliza & ConceptosPoliza(intIndex).Texto
                Case "14"
                    vl_strConceptoPoliza = vl_strConceptoPoliza & Trim(lblFolio.Caption)
                Case "20"
                    vl_strConceptoPoliza = vl_strConceptoPoliza & Trim(txtCliente.Text)
            End Select
        Next intIndex
    Else
        Me.MousePointer = 0
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        Exit Sub
    End If
    vl_strNotaTipo = ""
    '------------------- aqui termina el agregado de concepto de poliza
    
    'vllngNumPoliza = flngInsertarPoliza(CDate(mskFecha.Text), "D", "NOTA " & IIf(optNotaCredito.Value, "CREDITO ", "CARGO ") & Trim(vlstrFolioDocumento) & " PARA " & Trim(txtCliente.Text), vllngPersonaGraba)
    vllngNumPoliza = flngInsertarPoliza(CDate(mskFecha.Text), "D", vl_strConceptoPoliza, vllngPersonaGraba)
    
    '---------------------------'
    ' Guardar la nota (maestro) '
    '---------------------------'
    With rsCcNota
        .AddNew
        !chrFolioNota = Trim(vlstrFolioDocumento)
        !chrFolio = Trim(vlstrFolio)
        !chrSerie = Trim(vlstrSerie)
        !dtmfecha = CDate(mskFecha.Text)
        !dtmFechahora = CDate(mskFecha.Text) + fdtmServerHora
        !intCliente = IIf(optCliente.Value = True, Val(txtCveCliente.Text), Null)
        !chrTipo = IIf(optNotaCargo.Value, "CA", "CR")
        !MNYSUBTOTAL = CCur(txtSubtotal)
        !MNYDESCUENTO = CCur(txtDescuentoTot)
        !MNYIVA = CCur(txtIVA)
        !chrestatus = "A"
        !intPersonaGraba = vllngPersonaGraba
        !intPersonaBorra = 0
        !intNumPoliza = vllngNumPoliza
        !smicvedepartamento = vgintNumeroDepartamento
        !dtmFechaRegistro = fdtmServerFecha
        !intNumPolizaCancelacion = 0
        !vchPacienteImpresion = Mid(strPacienteImpresion, 1, 100)
        !vchFacturaImpresion = Mid(strFacturaImpresion, 1, 20)
        !vchComentario = Trim(txtComentario.Text)
        !CHRTIPOPACIENTE = IIf(optCliente.Value = True, Null, IIf(OptTipoPaciente(0).Value = True, "I", "E"))
        !INTMOVPACIENTE = Val(txtCveCliente.Text)
        !chrnotadirigida = IIf(optCliente.Value = True, "C", "P")
        !chrmotivonota = IIf(OptMotivoNota(0).Value = True, "E", "O")
        !chrFacturaPagada = IIf(chkFacturasPagadas.Value = 1 Or chkFacturasPaciente.Value = 1, "P", Null)
        If cboUsoCFDI.ListIndex > -1 Then
            !intCveUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
        End If
        
        If cboFormaPago.ListIndex > -1 Then
            !VCHFORMAPAGO = Left(Trim(cboFormaPago.Text), 2)
        End If
        
        If cboMetodoPago.ListIndex > -1 Then
            !vchMetodoPago = Left(Trim(cboMetodoPago.Text), 3)
        End If
        
        If Not blnValidarDiferentesFacturas Then
            'Si fue mas de una factura se cambia el paciente y factura para la impresión de la nota.
            rsCcNota!vchPacienteImpresion = "VARIOS"
            rsCcNota!vchFacturaImpresion = "VARIAS"
        Else
            rsCcNota!vchFacturaImpresion = strFacturaImpresion
        End If
        .Update
        
        vllngSecuencia = flngObtieneIdentity("SEC_CCNOTA", rsCcNota!intConsecutivo)
        lngIDnota = vllngSecuencia
    End With
    
    '-------------------------------------------'
    ' Guardar las facturas que integran la nota '
    '-------------------------------------------'
    If sstFacturasCreditos.Tab = 0 Then
        strTipoReferencia = "FA"
    Else
        strTipoReferencia = "MA"
    End If
    
    With rsCcNotaFactura
        vllngContador = 0
        intTotalFacturas = 0
        Do While vllngContador <= UBound(aFacturas(), 1)
            If aFacturas(vllngContador).vldblSubtotal <> 0 Or aFacturas(vllngContador).vldblDescuento <> 0 Or aFacturas(vllngContador).vldblIVA <> 0 Then
                If optNotaCredito.Value Then
                    '-----------------------------------------------------------------------------------------------------------'
                    ' Afectar el crédito, cuando es una nota de crédito: Se aumenta la cantidad pagada al movimiento de crédito '
                    '-----------------------------------------------------------------------------------------------------------'
                    If optCliente.Value = True And chkFacturasPagadas.Value = 0 Then
                        vlstrSentencia = "SELECT intNumMovimiento FROM CcMovimientoCredito WHERE chrTipoReferencia = '" & strTipoReferencia & "' AND trim(chrFolioReferencia) = trim('" & aFacturas(vllngContador).vlstrFolioFactura & "') AND intnumcliente = " & Val(txtCveCliente.Text) & " AND bitcancelado = 0"
                        'vlstrSentencia = "SELECT intNumMovimiento FROM CcMovimientoCredito WHERE chrTipoReferencia = '" & strTipoReferencia & "' AND chrFolioReferencia = '" & aFacturas(vllngContador).vlstrFolioFactura & "' AND intnumcliente = '" & txtCveCliente.Text & "' AND bitcancelado = 0"
                        vllngnumcredito = frsRegresaRs(vlstrSentencia).Fields(0)
                        vlstrSentencia = "UPDATE CCMOVIMIENTOCREDITO " & _
                                         " SET CCMOVIMIENTOCREDITO.MNYCANTIDADPAGADA = CCMOVIMIENTOCREDITO.MNYCANTIDADPAGADA + " & (((aFacturas(vllngContador).vldblSubtotal * aFacturas(vllngContador).vldblTipoCambio) - (aFacturas(vllngContador).vldblDescuento * aFacturas(vllngContador).vldblTipoCambio) + (aFacturas(vllngContador).vldblIVA * aFacturas(vllngContador).vldblTipoCambio))) & _
                                         " WHERE intNumMovimiento = " & (vllngnumcredito)
                        pEjecutaSentencia vlstrSentencia
                    End If
                    Else
                        vgstrParametrosSP = fstrFechaSQL(mskFecha.Text) _
                                        & "|" & txtCveCliente.Text _
                                        & "|" & rsDatosCliente!intnumcuentacontable _
                                        & "|" & Trim(vlstrFolioDocumento) _
                                        & "|" & "CA" _
                                        & "|" & (aFacturas(vllngContador).vldblSubtotal * aFacturas(vllngContador).vldblTipoCambio) - (aFacturas(vllngContador).vldblDescuento * aFacturas(vllngContador).vldblTipoCambio) + Format((aFacturas(vllngContador).vldblIVA * aFacturas(vllngContador).vldblTipoCambio), vlstrFormato) _
                                        & "|" & Str(vgintNumeroDepartamento) _
                                        & "|" & Str(vllngPersonaGraba) _
                                        & "|" & " " & "|" & "0" & "|" & (aFacturas(vllngContador).vldblSubtotal * aFacturas(vllngContador).vldblTipoCambio) - (aFacturas(vllngContador).vldblDescuento * aFacturas(vllngContador).vldblTipoCambio) & "|" & Format((aFacturas(vllngContador).vldblIVA * aFacturas(vllngContador).vldblTipoCambio), vlstrFormato)
                        vllngnumcredito = 1
                        frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, vllngnumcredito
                    End If
                .AddNew
                !intConsecutivo = vllngSecuencia
                !chrfoliofactura = Trim(aFacturas(vllngContador).vlstrFolioFactura)
                !MNYSUBTOTAL = aFacturas(vllngContador).vldblSubtotal
                !MNYDESCUENTO = aFacturas(vllngContador).vldblDescuento
                !MNYIVA = Format(aFacturas(vllngContador).vldblIVA, vlstrFormato)
                !intNumMovimientoCredito = vllngnumcredito
                !chrTipoFolio = IIf(sstFacturasCreditos.Tab = 0, "FA", "MA")
                .Update
                
                intTotalFacturas = intTotalFacturas + 1
            End If
            vllngContador = vllngContador + 1
        Loop
    End With

    '---------------------------------'
    ' Detalle de las notas de crédito '
    '---------------------------------'

    If optNotaCredito Then
        'CREDITO
        If optCliente Then vl_TipoCH = "X"  'cliente
        If optPaciente Then vl_TipoCH = "2" 'paciente
    Else
        'CARGO
        vl_TipoCH = "W"
    End If
    
    With grdNotas
        dblTotalIVA = 0
        vldblCantidadTotal = 0
        vldblCantidadTotalDesc = 0
        vldblTipoCambio = 1
        
        For vllngContador = 1 To .Rows - 1
        
'+++++++++aqui construir la referencia del detalle y el concepto de detalle para asignarlo a la funcion flngInsertarPolizaDetalle
            vl_strConceptoDtallePoliza = ""
            vl_strReferenciaDetallePoliza = ""
            
            For vllngConRefDetallePoliza = 1 To 2
                If fblnObtenerConfigPolizanva(vllngConRefDetallePoliza, vl_TipoCH, frmNotas.Name) = True Then     'Concepto detalle
                    For intIndex = 0 To UBound(ConceptosPoliza, 1)
                        Select Case ConceptosPoliza(intIndex).Tipo
                            Case "1"
                                If vllngConRefDetallePoliza = 1 Then vl_strConceptoDtallePoliza = vl_strConceptoDtallePoliza & ConceptosPoliza(intIndex).Texto
                                If vllngConRefDetallePoliza = 2 Then vl_strReferenciaDetallePoliza = vl_strReferenciaDetallePoliza & ConceptosPoliza(intIndex).Texto
                            Case "10"
                                If vllngConRefDetallePoliza = 1 Then vl_strConceptoDtallePoliza = vl_strConceptoDtallePoliza & Trim(.TextMatrix(vllngContador, 1)) 'No. Factura
                                If vllngConRefDetallePoliza = 2 Then vl_strReferenciaDetallePoliza = vl_strReferenciaDetallePoliza & Trim(.TextMatrix(vllngContador, 1)) 'No. Factura
                        End Select
                    Next intIndex
                End If
            Next vllngConRefDetallePoliza
'++++++++++aqui termina la asignacion de referencia y concepto de detalle de poliza
        
            vldblTipoCambio = Val(.TextMatrix(vllngContador, vlintColTipoCambio))
            vldblCantidad = Val(Format(.TextMatrix(vllngContador, vlintColCantidad), vlstrFormato))
            'vldblCantidad = Val(Format(grdNotas.TextMatrix(vllngContador, vlintColCantidadBase) * IIf(dblPorcentajeNota = 0, 1, dblPorcentajeNota)))
            vldblDescuento = Val(Format(.TextMatrix(vllngContador, vlintColDescuento), vlstrFormato))
            vldblIVA = Val(Format(.TextMatrix(vllngContador, vlintColIVA), vlstrFormato))
            'vldblIVA = Val((.TextMatrix(vllngContador, vlintColIVANotaSinRedondear)))
            
            rsCcNotaDetalle.AddNew
            rsCcNotaDetalle!intConsecutivo = vllngSecuencia
            rsCcNotaDetalle!intConcepto = Val(.TextMatrix(vllngContador, vlintColCveConcepto))
            rsCcNotaDetalle!MNYCantidad = vldblCantidad
            rsCcNotaDetalle!MNYDESCUENTO = vldblDescuento
            rsCcNotaDetalle!MNYIVA = vldblIVA
            
            ' Valida que sea cliente o paciente
            If optCliente.Value = True Then
                If chkFacturasPagadas.Value = 0 Then
                    rsCcNotaDetalle!intCuentaIngreso = Val(.TextMatrix(vllngContador, vlintColCtaIngresos))
                Else
                    intNumeroCuenta = 1
                    strParametros = "INTNUMCPUENTEDESCTOSNOTASFACPAGADAS" & "|" & vgintClaveEmpresaContable
                    ' Función que regresa el número de cuenta contable de la cuenta puente en facturas pagadas
                    frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                    rsCcNotaDetalle!intCuentaIngreso = intNumeroCuenta
                End If
                rsCcNotaDetalle!intCuentaDescuento = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos))
            Else
                '----- Agregado para caso 7374 -----'
                If OptMotivoNota(0).Value = True Then
                    rsCcNotaDetalle!intCuentaIngreso = Val(.TextMatrix(vllngContador, vlintColCtaIngresos))
                Else
                    intNumeroCuenta = 1
                    If chkFacturasPaciente.Value = 0 Then
                        strParametros = "INTNUMCUENTAPUENTEDESCTOSPACIENTE" & "|" & vgintClaveEmpresaContable
                        ' Función que regresa el número de cuenta contable de la cuenta puente para pacientes
                        frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                    Else
                        strParametros = "INTNUMCPUENTEDESCTOSNOTASFACPAGADAS" & "|" & vgintClaveEmpresaContable
                        ' Función que regresa el número de cuenta contable de la cuenta puente en facturas pagadas
                        frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                    End If
                    
                    rsCcNotaDetalle!intCuentaIngreso = intNumeroCuenta
                End If
                    
                If Val(.TextMatrix(vllngContador, 13)) = 0 Then
                    If optNotaCredito.Value = True And OptMotivoNota(1).Value = True And sstFacturasCreditos.Tab = 0 Then
                        rsCcNotaDetalle!intCuentaDescuento = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), Val(.TextMatrix(vllngContador, 14)), "NOTA")
                    Else
                        rsCcNotaDetalle!intCuentaDescuento = flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), Val(.TextMatrix(vllngContador, 14)), "DESCUENTO")
                    End If
                Else
                    lngNumDepto = 1
                    ' Función que regresa número de departamento que solicita requisición
                    frsEjecuta_SP Val(.TextMatrix(vllngContador, 13)), "FN_CCSELDEPTOSOLICITAREQ", True, lngNumDepto
                    If lngNumDepto = 0 Then
                        intDepartamento = Val(.TextMatrix(vllngContador, 14))
                    Else
                        intDepartamento = lngNumDepto
                    End If
                        
                    If optNotaCredito.Value = True And OptMotivoNota(1).Value = True And sstFacturasCreditos.Tab = 0 Then
                        rsCcNotaDetalle!intCuentaDescuento = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA")
                    Else
                        rsCcNotaDetalle!intCuentaDescuento = flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO")
                    End If
                End If
            End If

            rsCcNotaDetalle!intCuentaIVA = Val(.TextMatrix(vllngContador, vlintColCtaIVA))
            rsCcNotaDetalle!chrTipoCargo = IIf(sstFacturasCreditos.Tab = 0, .TextMatrix(vllngContador, vlintColTipoCargo), "CF")
            rsCcNotaDetalle!chrfoliofactura = Trim(.TextMatrix(vllngContador, vlintColFactura))
            rsCcNotaDetalle!chrTipoNota = IIf(sstFacturasCreditos.Tab = 0, "FA", "MA")
            
            If IIf(sstFacturasCreditos.Tab = 0, .TextMatrix(vllngContador, vlintColTipoCargo), "CF") <> "CF" Then
                rsCcNotaDetalle!numImporteGravado = vldblIVA / (fdblIVAConcepto(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), IIf(sstFacturasCreditos.Tab = 0, .TextMatrix(vllngContador, vlintColTipoCargo), "CF")) / 100)
            Else
                If optCliente.Value = True Then
                    If fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto))) And flngCveEmpresaCliente(CLng(txtCveCliente.Text)) <> 0 Then
                        rsCcNotaDetalle!numImporteGravado = vldblIVA / (fdblTasaIVAEmpresa(flngCveEmpresaCliente(CLng(txtCveCliente.Text))) / 100)
                    Else
                        rsCcNotaDetalle!numImporteGravado = vldblIVA / (fdblIVAConcepto(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), IIf(sstFacturasCreditos.Tab = 0, .TextMatrix(vllngContador, vlintColTipoCargo), "CF")) / 100)
                    End If
                Else
                    If fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto))) And flngCveEmpresaPaciente(CLng(txtCveCliente.Text)) <> 0 Then
                        rsCcNotaDetalle!numImporteGravado = vldblIVA / (fdblTasaIVAEmpresa(flngCveEmpresaPaciente(CLng(txtCveCliente.Text))) / 100)
                    Else
                        rsCcNotaDetalle!numImporteGravado = vldblIVA / (fdblIVAConcepto(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), IIf(sstFacturasCreditos.Tab = 0, .TextMatrix(vllngContador, vlintColTipoCargo), "CF")) / 100)
                    End If
                End If
            End If
            
            rsCcNotaDetalle.Update
            
            vldblCantidad = Val(Format(.TextMatrix(vllngContador, vlintColCantidad), vlstrFormato)) * vldblTipoCambio
            vldblDescuento = Val(Format(.TextMatrix(vllngContador, vlintColDescuento), vlstrFormato)) * vldblTipoCambio
            'vldblIVA = Val(Format(.TextMatrix(vllngContador, vlintColIVA), vlstrFormato)) * vldblTipoCambio
            
            '- Valida notas dirigidas a CLIENTES -'
            If optCliente.Value = True Then
                ' Valida que sea nota de crédito
                If optNotaCredito.Value = True Then
                    ' Valida el motivo de la nota
                    If OptMotivoNota(0).Value = True Then
                        If sstFacturasCreditos.Tab = 1 Or _
                                (sstFacturasCreditos.Tab = 0 And fblnConceptoPaquete(.TextMatrix(vllngContador, vlintColFactura), Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                                (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "P" And fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                                (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "C") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "OC") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "AR") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "ES") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "GE") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "EX") Or _
                                (sstFacturasCreditos.Tab = 0 And (.TextMatrix(vllngContador, vlintColTipoCargo) = "CF" And .TextMatrix(vllngContador, vlintColSeleccionoOtroConcepto) = True)) Then
                                
                            vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                            '--------------------------------------------------------------------------
                            'Cambio para caso 8736
                            'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                            'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                            vlblnCuentaIngresoSaldada = False
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                vlintBitSaldarCuentas = 1
                                frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                If vlintBitSaldarCuentas = 1 Then
                                    '-----------------------------------'
                                    ' Abono para el Ingreso - Descuento '
                                    '-----------------------------------'
                                    If (vldblCantidad - vldblDescuento) > 0 Then
                                        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                        vlblnCuentaIngresoSaldada = True
                                    ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                        vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                    ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                        vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                    End If
                                End If
                            End If
                            
                            If vlblnCuentaIngresoSaldada = False Then
                                'Cargo o abono a la cuenta de ingreso del concepto
                                If vldblCantidad <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                                
                                'Abono a la cuenta de descuento del concepto
                                If vldblDescuento <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                            End If
                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1)
                        Else
                            vgstrParametrosSP = txtCuenta.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & txtCveCliente.Text
                            Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")

                            If rsDepartamentoConcepto.RecordCount > 0 Then
                                Do While Not rsDepartamentoConcepto.EOF
                                    vldblCantidad = Format((rsDepartamentoConcepto!mnyCantidadSinDescuento * .TextMatrix(vllngContador, 20) * CDbl(.TextMatrix(vllngContador, vlintColPorcentajeFactura))), vlstrFormato)
                                    'vldblDescuento = Format((rsDepartamentoConcepto!MNYDESCUENTO * .TextMatrix(vllngContador, 20)), vlstrFormato)
                                    vldblDescuento = (((rsDepartamentoConcepto!MNYDESCUENTO + .TextMatrix(vllngContador, vlintColDescuentoEspecial)) * .TextMatrix(vllngContador, 20)))
                                    vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                    vldblCantidadTotalDesc = vldblCantidadTotalDesc + vldblDescuento
                                    '--------------------------------------------------------------------------
                                    'Cambio para caso 8736
                                    'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                                    'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                                    vlblnCuentaIngresoSaldada = False
                                    vllngCtaIngresos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO")
                                    If vllngCtaIngresos = 0 Then
                                        MsgBox "No se encontró la cuenta de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                    If vllngCtaIngresos = vllngCtaDescuentos Then
                                        'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                        vlintBitSaldarCuentas = 1
                                        frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                        If vlintBitSaldarCuentas = 1 Then
                                            '-----------------------------------'
                                            ' Abono para el Ingreso - Descuento '
                                            '-----------------------------------'
                                            If (vldblCantidad - vldblDescuento) > 0 Then
                                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0), .TextMatrix(vllngContador, vlintColFactura))
                                                vlblnCuentaIngresoSaldada = True
                                            ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                                vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                            ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                                vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                            End If
                                        End If
                                    End If
                                    
                                    If vlblnCuentaIngresoSaldada = False Then
                                        'Cargo o abono a la cuenta de ingreso del concepto
                                        If vldblCantidad <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                        
                                        'Abono a la cuenta de descuento del concepto
                                        If vldblDescuento <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0)
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 0, .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 0, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                    End If
                                    rsDepartamentoConcepto.MoveNext
                                Loop
                            Else
                                vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                '--------------------------------------------------------------------------
                                'Cambio para caso 8736
                                'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                                'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                                vlblnCuentaIngresoSaldada = False
                                If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = 0 Then
                                    'No se encontró la cuenta contable
                                    MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                                If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                    'No se encontró la cuenta contable
                                    MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                                If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                    'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                    vlintBitSaldarCuentas = 1
                                    frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                    If vlintBitSaldarCuentas = 1 Then
                                        '-----------------------------------'
                                        ' Abono para el Ingreso - Descuento '
                                        '-----------------------------------'
                                        If (vldblCantidad - vldblDescuento) > 0 Then
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                            vlblnCuentaIngresoSaldada = True
                                        ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                            vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                        ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                            vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                        End If
                                    End If
                                End If
                                
                                If vlblnCuentaIngresoSaldada = False Then
                                    'Cargo o abono a la cuenta de ingreso del concepto
                                    If vldblCantidad <> 0 Then
                                        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                    End If
                                    
                                    'Abono a la cuenta de descuento del concepto
                                    If vldblDescuento <> 0 Then
                                        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0)
                                    End If
                                End If
                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1)
                            End If
                        End If
                    Else
                        'Cargo a la cuenta de descuento del concepto
                        '- Otorgamiento de descuentos -'   '- Clientes -'
                        If sstFacturasCreditos.Tab = 1 Or _
                                (sstFacturasCreditos.Tab = 0 And fblnConceptoPaquete(.TextMatrix(vllngContador, vlintColFactura), Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                                (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "P" And fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                                (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "C") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "OC") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "AR") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "ES") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "GE") Or _
                                (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "EX") Or _
                                (sstFacturasCreditos.Tab = 0 And (.TextMatrix(vllngContador, vlintColTipoCargo) = "CF" And .TextMatrix(vllngContador, vlintColSeleccionoOtroConcepto) = True)) Then
                            
                            If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                        
                            vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1)
                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                        Else
                            vgstrParametrosSP = txtCuenta.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & txtCveCliente.Text
                            Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                
                            If rsDepartamentoConcepto.RecordCount > 0 Then
                                Do While Not rsDepartamentoConcepto.EOF
                                    'vldblCantidad = Format(((rsDepartamentoConcepto!MNYCANTIDAD - .TextMatrix(vllngContador, vlintColDescuentoEspecial)) * dblPorcentajeNota * CDbl(.TextMatrix(vllngContador, vlintColPorcentajeFacturaConDescuento))), vlstrFormato)
                                                                        
                                    vldblCantidad = rsDepartamentoConcepto!mnyCantidadSinDescuento * CDbl(.TextMatrix(vllngContador, vlintColPorcentajeFacturaConDescuento))
                                    vldblCantidad = vldblCantidad - rsDepartamentoConcepto!MNYDESCUENTO - CDbl(.TextMatrix(vllngContador, vlintColDescuentoEspecial))
                                    vldblCantidad = vldblCantidad * dblPorcentajeNota
                                    vldblCantidad = Format(vldblCantidad, vlstrFormato)
                                    
                                    'vldblDescuento = Format((rsDepartamentoConcepto!MNYDESCUENTO * .TextMatrix(vllngContador, 20)), vlstrFormato)
                                    vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    
                                    'Cargo a la cuenta de descuento del concepto
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1)
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA"), vldblCantidad, 1, .TextMatrix(vllngContador, vlintColFactura))
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA"), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                    rsDepartamentoConcepto.MoveNext
                                Loop
                            Else
                                If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                    'No se encontró la cuenta contable
                                    MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                            
                                vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblCantidad, 1)
                            End If
                        End If
                    End If
                Else
                    If sstFacturasCreditos.Tab = 1 Or _
                            (sstFacturasCreditos.Tab = 0 And fblnConceptoPaquete(.TextMatrix(vllngContador, vlintColFactura), Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                            (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "P" And fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                            (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "C") Or _
                            (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "OC") Or _
                            (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "AR") Or _
                            (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "ES") Or _
                            (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "GE") Or _
                            (sstFacturasCreditos.Tab = 0 And .TextMatrix(vllngContador, vlintColTipoCargo) = "EX") Or _
                            (sstFacturasCreditos.Tab = 0 And (.TextMatrix(vllngContador, vlintColTipoCargo) = "CF" And .TextMatrix(vllngContador, vlintColSeleccionoOtroConcepto) = True)) Then
                            
                        vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                        '--------------------------------------------------------------------------
                        'Cambio para caso 8736
                        'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                        'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                        vlblnCuentaIngresoSaldada = False
                        If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = 0 Then
                            'No se encontró la cuenta contable
                            MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                        If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                            'No se encontró la cuenta contable
                            MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                            Exit Sub
                        End If
                        
                        If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                            'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                            vlintBitSaldarCuentas = 1
                            frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                            If vlintBitSaldarCuentas = 1 Then
                                '-----------------------------------'
                                ' Abono para el Ingreso - Descuento '
                                '-----------------------------------'
                                If (vldblCantidad - vldblDescuento) > 0 Then
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                    vlblnCuentaIngresoSaldada = True
                                ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                    vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                    vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                End If
                            End If
                        End If
                        
                        If vlblnCuentaIngresoSaldada = False Then
                            'Cargo o abono a la cuenta de ingreso del concepto
                            If vldblCantidad <> 0 Then
                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                            End If
                            
                            'Cargo a la cuenta de descuento del concepto
                            If vldblDescuento <> 0 Then
                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 1)
                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                            End If
                        End If
                    Else
                        vgstrParametrosSP = txtCuenta.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & txtCveCliente.Text
                        Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                        vldblCantidadConceptoNota = vldblCantidad
                        vldblDescuentoConceptoNota = vldblDescuento
                        vldblCantidadTotalConcepto = 0
                        
                        If rsDepartamentoConcepto.RecordCount > 0 Then
'                            If rsDepartamentoConcepto.RecordCount > 0 Then
                                rsDepartamentoConcepto.MoveFirst
                                Do While Not rsDepartamentoConcepto.EOF
                                    vldblCantidad = (rsDepartamentoConcepto!mnyCantidadSinDescuento * .TextMatrix(vllngContador, 20) * CDbl(.TextMatrix(vllngContador, vlintColPorcentajeFactura)))
                                    vldblCantidadTotalConcepto = vldblCantidadTotalConcepto + vldblCantidad
                                    rsDepartamentoConcepto.MoveNext
                                Loop
'                            End If
                            
'                            If rsDepartamentoConcepto.RecordCount > 0 Then
                                rsDepartamentoConcepto.MoveFirst
                                Do While Not rsDepartamentoConcepto.EOF
                                    vldblCantidad = Format((rsDepartamentoConcepto!mnyCantidadSinDescuento * .TextMatrix(vllngContador, 20) * CDbl(.TextMatrix(vllngContador, vlintColPorcentajeFactura))), vlstrFormato)
                                    'vldblFactorDescuento = .TextMatrix(vllngContador, vlintColDescuento) / (.TextMatrix(vllngContador, vlintColCantidad) / vldblPorcentajeFactura)
                                    'vldblDescuento = Format((rsDepartamentoConcepto!mnycantidadSinDescuento * .TextMatrix(vllngContador, 20) * vldblFactorDescuento), vlstrFormato)
                                    'vldblDescuento = ((rsDepartamentoConcepto!mnycantidadSinDescuento * .TextMatrix(vllngContador, 20) * vldblFactorDescuento))
                                    
                                    vldblPorcentajeDepartamento = vldblCantidad / vldblCantidadTotalConcepto
                                    vldblCantidad = vldblCantidadConceptoNota * vldblPorcentajeDepartamento
                                    vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                    
                                    vldblDescuento = vldblDescuentoConceptoNota * vldblPorcentajeDepartamento
                                    vldblCantidadTotalDesc = vldblCantidadTotalDesc + vldblDescuento
                                    '--------------------------------------------------------------------------
                                    'Cambio para caso 8736
                                    'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                                    'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                                    vlblnCuentaIngresoSaldada = False
                                    vllngCtaIngresos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO")
                                    If vllngCtaIngresos = 0 Then
                                        MsgBox "No se encontró la cuenta de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                    If vllngCtaIngresos = vllngCtaDescuentos Then
                                        'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                        vlintBitSaldarCuentas = 1
                                        frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                        If vlintBitSaldarCuentas = 1 Then
                                            '-----------------------------------'
                                            ' Abono para el Ingreso - Descuento '
                                            '-----------------------------------'
                                            If (vldblCantidad - vldblDescuento) > 0 Then
                                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0), .TextMatrix(vllngContador, vlintColFactura))
                                                vlblnCuentaIngresoSaldada = True
                                            ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                                vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                            ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                                vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                            End If
                                        End If
                                    End If
                                    
                                    If vlblnCuentaIngresoSaldada = False Then
                                        'Cargo o abono a la cuenta de ingreso del concepto
                                        If vldblCantidad <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                        
                                        'Cargo a la cuenta de descuento del concepto
                                        If vldblDescuento <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 1)
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 1, .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                    End If
                                    rsDepartamentoConcepto.MoveNext
                                Loop
'                            End If
                        Else
                            vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                            '--------------------------------------------------------------------------
                            'Cambio para caso 8736
                            'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                            'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                            vlblnCuentaIngresoSaldada = False
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                vlintBitSaldarCuentas = 1
                                frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                If vlintBitSaldarCuentas = 1 Then
                                    '-----------------------------------'
                                    ' Abono para el Ingreso - Descuento '
                                    '-----------------------------------'
                                    If (vldblCantidad - vldblDescuento) > 0 Then
                                        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), IIf(optNotaCredito.Value, 1, 0))
                                        vlblnCuentaIngresoSaldada = True
                                    ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                        vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                    ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                        vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                    End If
                                End If
                            End If
                            
                            If vlblnCuentaIngresoSaldada = False Then
                                'Cargo o abono a la cuenta de ingreso del concepto
                                If vldblCantidad <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0))
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                                
                                'Cargo a la cuenta de descuento del concepto
                                If vldblDescuento <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 1)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                            End If
                        End If
                    End If
                End If
            '- Valida notas dirigidas a PACIENTES -'
            Else
                '- Modificado para Caso 7374: Permitir aplicar Notas de Crédito por "Error de facturación" a Pacientes -'
                If optNotaCredito.Value = True Then
                    If OptMotivoNota(0).Value = True Then
                    '- Error de Facturación -'
                        If (sstFacturasCreditos.Tab = 0 And fblnConceptoPaquete(.TextMatrix(vllngContador, vlintColFactura), Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                           (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "P" And fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Then
                            vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                            '--------------------------------------------------------------------------
                            'Cambio para caso 8736
                            'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                            'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                            vlblnCuentaIngresoSaldada = False
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            If Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) = 0 Then
                                'No se encontró la cuenta contable
                                MsgBox Replace(SIHOMsg(222), ".", "") & " de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                EntornoSIHO.ConeccionSIHO.RollbackTrans
                                Exit Sub
                            End If
                            If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                vlintBitSaldarCuentas = 1
                                frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                If vlintBitSaldarCuentas = 1 Then
                                    '-----------------------------------'
                                    ' Abono para el Ingreso - Descuento '
                                    '-----------------------------------'
                                    If (vldblCantidad - vldblDescuento) > 0 Then
                                        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), 1)
                                        vlblnCuentaIngresoSaldada = True
                                    ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                        vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                    ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                        vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                    End If
                                End If
                            End If
                            
                            If vlblnCuentaIngresoSaldada = False Then
                                'Cargo a la cuenta de ingreso del concepto
                                If vldblCantidad <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, 1)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                            
                                'Abono a la cuenta de descuento del concepto
                                If vldblDescuento <> 0 Then
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                            End If
                        Else
                            vgstrParametrosSP = txtCuenta.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & 0
                            Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                
                            If rsDepartamentoConcepto.RecordCount > 0 Then
                                Do While Not rsDepartamentoConcepto.EOF
                                    'vldblCantidad = Format((rsDepartamentoConcepto!mnycantidad * dblPorcentajeNota), vlstrFormato)
                                    vldblCantidad = Format((rsDepartamentoConcepto!mnyCantidadSinDescuento * dblPorcentajeNota), vlstrFormato)
                                    vldblDescuento = Format((rsDepartamentoConcepto!MNYDESCUENTO * dblPorcentajeNota), vlstrFormato)
                                    vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                    vldblCantidadTotalDesc = vldblCantidadTotalDesc + vldblDescuento
                                    
                                    '--------------------------------------------------------------------------
                                    'Cambio para caso 8736
                                    'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
                                    'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
                                    vlblnCuentaIngresoSaldada = False
                                    vllngCtaIngresos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO")
                                    If vllngCtaIngresos = 0 Then
                                        MsgBox "No se encontró la cuenta de ingresos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'If Val(.TextMatrix(vllngContador, vlintColCtaIngresos)) = Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)) Then
                                    If vllngCtaIngresos = vllngCtaDescuentos Then
                                        'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                                        vlintBitSaldarCuentas = 1
                                        frsEjecuta_SP CStr(.TextMatrix(vllngContador, vlintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                                        If vlintBitSaldarCuentas = 1 Then
                                            '-----------------------------------'
                                            ' Abono para el Ingreso - Descuento '
                                            '-----------------------------------'
                                            If (vldblCantidad - vldblDescuento) > 0 Then
                                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), (vldblCantidad - vldblDescuento), 1)
                                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), (vldblCantidad - vldblDescuento), 1, .TextMatrix(vllngContador, vlintColFactura))
                                                vlblnCuentaIngresoSaldada = True
                                            ElseIf (vldblCantidad - vldblDescuento) < 0 Then
                                                vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                                            ElseIf (vldblCantidad - vldblDescuento) = 0 Then
                                                vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                                            End If
                                        End If
                                    End If
                                    
                                    If vlblnCuentaIngresoSaldada = False Then
                                        'Cargo a la cuenta de ingreso del concepto
                                        If vldblCantidad <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaIngresos)), vldblCantidad, 1)
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, 1, .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "INGRESO"), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                    
                                        'Abono a la cuenta de descuento del concepto
                                        If vldblDescuento <> 0 Then
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, Val(.TextMatrix(vllngContador, vlintColCtaDescuentos)), vldblDescuento, 0)
                                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 0, .TextMatrix(vllngContador, vlintColFactura))
                                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "DESCUENTO"), vldblDescuento, 0, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                        End If
                                    End If
                                    rsDepartamentoConcepto.MoveNext
                                Loop
                            End If
                        End If
                    Else
                    '- Otorgamiento de descuentos -'
                        If (Val(.TextMatrix(vllngContador, 13)) = 0 Or (Val(.TextMatrix(vllngContador, 13)) <> 0 And .TextMatrix(vllngContador, vlintColFactura) = "")) Then
                            If .TextMatrix(vllngContador, vlintColFactura) = "" Then
                                intDepartamento = Val(.TextMatrix(vllngContador, 14))
                                vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                If sstFacturasCreditos.Tab = 0 Then
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & fstrDepartamento(intDepartamento) & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                Else
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO")
                                    If vllngCtaDescuentos = 0 Then
                                        MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & fstrDepartamento(intDepartamento) & ".", vbOKOnly + vbInformation, "Mensaje"
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO"), vldblCantidad, 1, .TextMatrix(vllngContador, vlintColFactura))
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO"), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                End If
                            Else
                                If (sstFacturasCreditos.Tab = 0 And fblnConceptoPaquete(.TextMatrix(vllngContador, vlintColFactura), Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Or _
                                   (sstFacturasCreditos.Tab = 0 And vlstrtipofactura = "P" And fblnConceptoAseguradora(Val(.TextMatrix(vllngContador, vlintColCveConcepto)))) Then
                                    
                                    If rsDepartamentoConcepto.State = 0 Then
                                        If vlstrtipofactura = "P" Then
                                            vgstrParametrosSP = txtCuenta.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & 0
                                            Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                                        Else
                                            vgstrParametrosSP = txtCveCliente.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & 0
                                            Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                                        End If
                                    End If
                                    
                                    vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                    vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA")
                                    If vllngCtaDescuentos = 0 Then
                                        If rsDepartamentoConcepto.RecordCount = 0 Then
                                            MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                        Else
                                            MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                        End If
                                        
                                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                                        Exit Sub
                                    End If
                                    'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1)
                                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                Else
                                    vgstrParametrosSP = txtCveCliente.Text & "|" & .TextMatrix(vllngContador, vlintColFactura) & "|" & .TextMatrix(vllngContador, vlintColCveConcepto) & "|" & 0
                                    Set rsDepartamentoConcepto = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDEPARTAMENTOSCONCEPTO")
                                    If rsDepartamentoConcepto.RecordCount > 0 Then
                                        Do While Not rsDepartamentoConcepto.EOF
                                            If sstFacturasCreditos.Tab = 0 Then
                                                vldblCantidad = Format((rsDepartamentoConcepto!MNYCantidad * dblPorcentajeNota), vlstrFormato)
                                                vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                                                vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA")
                                                If vllngCtaDescuentos = 0 Then
                                                    MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & " del departamento " & rsDepartamentoConcepto!VCHDESCRIPCION & ".", vbOKOnly + vbInformation, "Mensaje"
                                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                                    Exit Sub
                                                End If
                                                'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA"), Format(rsDepartamentoConcepto!MNYCantidad * dblPorcentajeNota, vlstrFormato), 1, .TextMatrix(vllngContador, vlintColFactura))
                                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), rsDepartamentoConcepto!SMIDEPARTAMENTO, "NOTA"), Format(rsDepartamentoConcepto!MNYCantidad * dblPorcentajeNota, vlstrFormato), 1, vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
                                            End If
                                            rsDepartamentoConcepto.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        Else
                            lngNumDepto = 1
                            ' Función que regresa número de departamento que solicita requisición
                            frsEjecuta_SP Val(.TextMatrix(vllngContador, 13)), "FN_CCSELDEPTOSOLICITAREQ", True, lngNumDepto
                            If lngNumDepto = 0 Then
                                intDepartamento = Val(.TextMatrix(vllngContador, 14))
                            Else
                                intDepartamento = lngNumDepto
                            End If
                            
                            If sstFacturasCreditos.Tab = 0 Then
                                vllngCtaDescuentos = flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA")
                                If vllngCtaDescuentos = 0 Then
                                    MsgBox "No se encontró la cuenta de descuentos por nota de crédito para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1)
                            Else
                                vllngCtaDescuentos = flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO")
                                If vllngCtaDescuentos = 0 Then
                                    MsgBox "No se encontró la cuenta de descuentos para el concepto " & fstrConceptoFacturacion(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), .TextMatrix(vllngContador, vlintColTipoCargo)) & ".", vbOKOnly + vbInformation, "Mensaje"
                                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                                    Exit Sub
                                End If
                                vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO"), vldblCantidad, 1, .TextMatrix(vllngContador, vlintColFactura))
                            End If
                            vldblCantidadTotal = vldblCantidadTotal + vldblCantidad
                        End If
                    
'                        If sstFacturasCreditos.Tab = 0 Then
'                            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDeptoNota(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "NOTA"), vldblCantidad, 1)
'                        Else
'                            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, flngCuentaConceptoDepartamento(Val(.TextMatrix(vllngContador, vlintColCveConcepto)), intDepartamento, "DESCUENTO"), vldblCantidad, 1, .TextMatrix(vllngContador, vlintColFactura))
'                        End If
                    End If
                End If
            End If
            dblTotalIVA = dblTotalIVA + (vldblIVA * .TextMatrix(vllngContador, vlintColTipoCambio))
        Next vllngContador
    End With
    
    '- Modificado para el caso 6723. Se cambió la función a Format porque se presentaron diferencias de centavos -'
    dblTotalIVA = Format(dblTotalIVA, vlstrFormato) 'Round(dblTotalIVA, 2)
    
    If vldblTipoCambio = 1 Then
        If Format((dblTotalIVA) - Val(Format(txtIVA.Text, vlstrFormato)), vlstrFormato) = 0.01 Then
            dblTotalIVA = dblTotalIVA - 0.01
        ElseIf Format((dblTotalIVA) - Val(Format(txtIVA.Text, vlstrFormato)), vlstrFormato) = -0.01 Then
            dblTotalIVA = dblTotalIVA + 0.01
        Else
            dblTotalIVA = Val(Format(txtIVA.Text, vlstrFormato))
        End If
    End If
    vldblCantidadTotal = vldblCantidadTotal + dblTotalIVA
    
    If vldblTipoCambio = 1 Then
        If Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) >= 0.01 And Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) <= 0.03 Then
            If dblTotalIVA <> 0 Then
                dblTotalIVA = dblTotalIVA - Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)
                vldblCantidadTotal = vldblCantidadTotal - Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato)
            End If
        ElseIf Format(Val(Format(txtTotal.Text, vlstrFormato)) - (vldblCantidadTotal - vldblCantidadTotalDesc), vlstrFormato) >= 0.01 And Format(Val(Format(txtTotal.Text, vlstrFormato)) - (vldblCantidadTotal - vldblCantidadTotalDesc), vlstrFormato) <= 0.03 Then
            If dblTotalIVA <> 0 Then
                dblTotalIVA = dblTotalIVA + Format(Val(Format(txtTotal.Text, vlstrFormato)) - (vldblCantidadTotal - vldblCantidadTotalDesc), vlstrFormato)
                vldblCantidadTotal = vldblCantidadTotal + Format(Val(Format(txtTotal.Text, vlstrFormato)) - (vldblCantidadTotal - vldblCantidadTotalDesc), vlstrFormato)
            End If
        End If
    End If

    'Cargo o abono a la cuenta de IVA cobrado y no cobrado
    If optCliente.Value = True Then
        If dblTotalIVA <> 0 Then
            'Cargo/abono a IVA NO cobrado
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, glngCtaIVANoCobrado, dblTotalIVA, IIf(optNotaCredito.Value, 1, 0))
            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, glngCtaIVANoCobrado, dblTotalIVA, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
        End If
    Else
        If dblTotalIVA <> 0 Then
            'Cargo/abono a IVA cobrado
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, glngCtaIVACobrado, dblTotalIVA, IIf(optNotaCredito.Value, 1, 0))
            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, glngCtaIVACobrado, dblTotalIVA, IIf(optNotaCredito.Value, 1, 0), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
        End If
    End If
       
    If vldblTipoCambio = 1 Then
        If Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) = 0.01 Then
            vldblCantidadTotal = (vldblCantidadTotal - vldblCantidadTotalDesc)
        ElseIf Format((vldblCantidadTotal - vldblCantidadTotalDesc) - Val(Format(txtTotal.Text, vlstrFormato)), vlstrFormato) = -0.01 Then
            vldblCantidadTotal = (vldblCantidadTotal - vldblCantidadTotalDesc)
        Else
            vldblCantidadTotal = Val(Format(txtTotal.Text, vlstrFormato))
        End If
    Else
        vldblCantidadTotal = (vldblCantidadTotal - vldblCantidadTotalDesc)
    End If
    
    'Cargo o abono a la cuenta del cliente (según si es nota de crédito o cargo)
    If optCliente.Value = True Then
        If chkFacturasPagadas.Value = 0 Then
            intNumeroCuenta = 1
            ' Función que regresa el número de cuenta contable del cliente
            frsEjecuta_SP txtCveCliente, "FN_CCSELCUENTACLIENTE", True, intNumeroCuenta
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(txtTotal.Text, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
        Else
            intNumeroCuenta = 1
            strParametros = "INTNUMCPUENTEDESCTOSNOTASFACPAGADAS" & "|" & vgintClaveEmpresaContable
            ' Función que regresa el número de cuenta contable de la cuenta puente en facturas pagadas
            frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(txtTotal.Text, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
            'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
            vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
        End If
    Else
        intNumeroCuenta = 1
        If chkFacturasPaciente.Value = 0 Then
            strParametros = "INTNUMCUENTAPUENTEDESCTOSPACIENTE" & "|" & vgintClaveEmpresaContable
            ' Función que regresa el número de cuenta contable de la cuenta puente para pacientes
            frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
        Else
            strParametros = "INTNUMCPUENTEDESCTOSNOTASFACPAGADAS" & "|" & vgintClaveEmpresaContable
            ' Función que regresa el número de cuenta contable de la cuenta puente en facturas pagadas
            frsEjecuta_SP strParametros, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
        End If
        
        'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(txtTotal.Text, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
        'vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1))
        vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(vldblCantidadTotal, vlstrFormato)), IIf(optNotaCredito.Value, 0, 1), vl_strReferenciaDetallePoliza, vl_strConceptoDtallePoliza)
    End If

    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, IIf(optNotaCargo.Value, "NOTA DE CARGO", "NOTA DE CREDITO"), vlstrFolioDocumento)
    
    '------------------------------------------------------------'
    '------------------------- ADDENDA --------------------------'
    '------------------------------------------------------------'
    If cmdAddenda.Enabled = True Then
        'Valida la información de la addenda
        vlAddendaValida = 1
        strParametros = CStr(vglngCveAddenda) & "|" & CStr(vglngCuentaPacienteAddenda) & "|" & vgstrTipoPacienteAddenda
        frsEjecuta_SP strParametros, "FN_PVSELADDENDADATOSVALIDOS", True, vlAddendaValida
        If vlAddendaValida = 2 Then
            'No se ha configurado la addenda
            MsgBox SIHOMsg(1143), vbCritical, "Mensaje"
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            cmdAddenda.SetFocus
            Exit Sub
        ElseIf vlAddendaValida = 0 Then
            'Hay información incorrecta en la addenda
            MsgBox SIHOMsg(1140), vbCritical, "Mensaje"
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            cmdAddenda.SetFocus
            Exit Sub
        End If
    End If
    '------------------------------------------------------------'
    '--------------------------- FIN ----------------------------'
    '------------------------------------------------------------'
    
    '***********************************************************************************************'
    '***** (CR) Guardar información de la póliza fuera del corte para reporte en Corte de caja *****'
    vgstrParametrosSP = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P") & "|" & _
                        fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & IIf(optNotaCargo.Value, "NA", "NC") & _
                        "|" & CStr(vllngNumPoliza) & "|" & CStr(vllngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento)
    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSMOVIMIENTOFUERACORTE"
    '***********************************************************************************************'
    '***********************************************************************************************'
    
    If optNotaCredito.Value = True Then optNotaCargo.Enabled = False
    If optPaciente.Value = True Then
        'SP que guarda la configuración de Internos y Externos según el proceso:
         strParametros = vglngNumeroLogin & "|" & 11 & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
         frsEjecuta_SP strParametros, "SP_CCINSCONFIGURACIONPACIENTE"
    End If
    
    pHabilitaCuadros False
    rsCcNota.Close
    rsCcNotaFactura.Close
    rsCcNotaDetalle.Close
    pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
    
    vlblnlimpiaNotas = False
    chkPorcentajePaciente.Value = 0
    vlblnlimpiaNotas = True
                         
    '-------------------------------------------------------------------------------------------------
    'VALIDACIÓN DE LOS DATOS ANTES DE INSERTAR EN GNCOMPROBANTEFISCALDIGITAL EN EL PROCESO DE TIMBRADO
    '-------------------------------------------------------------------------------------------------
    If intTipoEmisionComprobante = 2 Then
       If Not fblnValidaDatosCFDCFDi(vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR"), IIf(intTipoCFDNota = 1, True, False), CInt(vlstrAnoAprobacion), vlstrNumeroAprobacion) Then
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          Exit Sub
       End If
    End If
                      
    EntornoSIHO.ConeccionSIHO.CommitTrans '*
    
    'Se agregó para el caso 19547
    ChecarCentavosPoliza vllngNumPoliza
                             
      If intTipoEmisionComprobante = 2 Then ' Si se realizará una emisión digital
      
        vgblRazonSocial = Trim(txtCliente.Text)
        If vlRazonSocialComprobante <> "" Then
            vgblRazonSocial = Trim(vlRazonSocialComprobante)
        End If
        
        'Barra de progreso CFD
        pgbBarraCFD.Value = 35
        freBarraCFD.Top = 3200
        Screen.MousePointer = vbHourglass
        lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital para la nota, por favor espere..."
        freBarraCFD.Visible = True
        freBarraCFD.Refresh
        frmNotas.Enabled = False
        If intTipoCFDNota = 1 Then
           pLogTimbrado 2
           pMarcarPendienteTimbre vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR"), vgintNumeroDepartamento
        End If
        EntornoSIHO.ConeccionSIHO.BeginTrans 'se abre otra transacción
        
        'Actualizar el comprobante para guardar el código postal y el regimen
        If vgstrVersionCFDI = "4.0" Then
            vgstrTipoNotaSig = IIf(optNotaCredito.Value = True, "CR", "CA")
            vgstrCodigoPostalSig = Trim(vlstrCodigo)
            vglngIdComprobanteSig = vllngSecuencia
            vgstrRegimenFiscalSig = vlstrRegimen
        End If
        
        If Not fblnGeneraComprobanteDigital(vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR"), 0, Val(vlstrAnoAprobacion), vlstrNumeroAprobacion, IIf(intTipoCFDNota = 1, True, False), , vglngCveAddenda) Then
           On Error Resume Next

            vgstrTipoNotaSig = ""
            vgstrCodigoPostalSig = ""
            vglngIdComprobanteSig = 0
            vgstrRegimenFiscalSig = ""

           EntornoSIHO.ConeccionSIHO.CommitTrans
           If intTipoCFDNota = 1 Then pLogTimbrado 1
           If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
              'El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
              MsgBox Replace(SIHOMsg(1306), "El comprobante", "La nota de " & IIf(optNotaCargo.Value, "cargo", "crédito")), vbInformation + vbOKOnly, "Mensaje"
           ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then 'No se realizó el timbrado
             '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
              MsgBox Replace(SIHOMsg(1338), "factura", "nota"), vbCritical + vbOKOnly, "Mensaje"
                  '___________________________________________________________________________________________________________________________________cancelar la nota
                  mskFechaInicial.Text = mskFecha.Text
                  mskFechaFinal.Text = mskFecha.Text
                  pCargaNotas ' debemos de llenar el grid con la información de la nota para poder cancelar la nota
                  If Me.grdBusqueda.Rows > 1 Then
                     For vlintContador = 1 To Me.grdBusqueda.Rows - 1 'localizar el row que corresponda a la nota que se acaba de agregar
                            If Val(Me.grdBusqueda.TextMatrix(vlintContador, vlintColIdNota)) = vllngSecuencia Then
                               grdBusqueda.Row = vlintContador
                               Exit For
                             End If
                     Next vlintContador
                  End If
                  blnCancelaNota = True
                  If chrTipoCancel = "" Then
                     vlstrSentencia = "select chrTipoCancel from ccTipoCancelacion where smiCveDepartamento = " & vgintNumeroDepartamento & " and tnyClaveEmpresa = " & vgintClaveEmpresaContable
                     Set ObjRs = frsRegresaRs(vlstrSentencia)
                     If ObjRs.RecordCount > 0 Then 'Si tiene configurado el parametro seleccionar la opción indicada
                        chrTipoCancel = Trim(ObjRs!chrTipoCancel)
                     Else
                        chrTipoCancel = "DOCUMENTO"
                     End If
                  End If
                  If chrTipoCancel = "ELEGIR" Then
                     frmFechaCancelacionNotas.Show vbModal
                  Else
                        frmFechaCancelacionNotas.pCancelaNota Trim(lblFolio.Caption), txtCliente.Text, chrTipoCancel, True, True, True, vllngPersonaGraba
                  End If
                  If blnCancelaNota = True Then
                     'Eliminamos la informacion de la nota de la tabla de pendientes de timbre fiscal
                     pEliminaPendientesTimbre vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR") 'quitamos la factura de pendientes de timbre fiscal
                  End If
                  '___________________________________________________________________________________________________________________________________________________
                  pLogTimbrado 1
                  frsEjecuta_SP CFDilngNumError & "|" & Left(CFDistrDescripError, 200) & "|" & cgstrModulo & "|" & Left(CFDistrProcesoError, 50) & " Linea:" & CFDiintLineaError & "|" & "", "SP_GNINSREGISTROERRORES", True
                  fblnImprimeComprobanteDigital vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR"), "I", vllngNumeroTipoFormato, 0
                  Screen.MousePointer = vbDefault
                  frmNotas.Enabled = True
                  freBarraCFD.Visible = False
                  InicializaComponentes
                  If optNotaCargo.Value Then
                     optNotaCargo.SetFocus
                  Else
                     optNotaCredito.SetFocus
                  End If
                  vlblnConsulta = False
                  vlblnSalir = False
                  cmdComprobante.Enabled = False
                  Exit Sub
           End If
        Else
        
        
            vgstrTipoNotaSig = ""
            vgstrCodigoPostalSig = ""
            vglngIdComprobanteSig = 0
            vgstrRegimenFiscalSig = ""
            
            vlstrCodigo = ""
            vlstrRegimen = ""
            vlRazonSocialComprobante = ""
            vlRFCComprobante = ""
            
           EntornoSIHO.ConeccionSIHO.CommitTrans
           If intTipoCFDNota = 1 Then
              pLogTimbrado 1
              pEliminaPendientesTimbre vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR")
           End If
        End If
                
        'Barra de progreso CFD
        pgbBarraCFD.Value = 100
        freBarraCFD.Top = 3200
        Screen.MousePointer = vbDefault
        freBarraCFD.Visible = False
        frmNotas.Enabled = True
    End If
                         

    'Asegúrese de que la impresora esté lista y presione aceptar.
    MsgBox SIHOMsg(343), vbOKOnly + vbInformation, "Mensaje"
    If intTipoEmisionComprobante = 2 Then ' Si se realizará una emisión digital
        If Not fblnImprimeComprobanteDigital(vllngSecuencia, IIf(optNotaCargo.Value, "CA", "CR"), "I", vllngNumeroTipoFormato, 0) Then
            Exit Sub
        End If

        'Verifica si debe mostrarse la pantalla de envío de CFDs por correo electrónico
        If fblnPermitirEnvio And vgIntBanderaTImbradoPendiente = 0 Then
            '¿Desea enviar por e-mail la información del comprobante fiscal digital?
            If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pEnviarCFD IIf(optNotaCargo.Value, "CA", "CR"), vllngSecuencia, CLng(vgintClaveEmpresaContable), Trim(txtRFC.Text), vllngPersonaGraba, Me
            End If
        End If
    Else
        
        txtComentario.Locked = True
        strTotalLetras = fstrNumeroenLetras(CDbl(Format(txtTotal.Text, "#############.00")), "pesos", "M.N.")
        vlstrFecha = Format(mskFecha.Text, "Long Date")
        vgstrParametrosSP = vllngSecuencia & "|" & strTotalLetras & "|" & vlstrFecha & "|" & IIf(optCliente.Value = True, "C", "P")
        Set rsNota = frsEjecuta_SP(vgstrParametrosSP, "Sp_CcRptNota")
        pImpFormato rsNota, 8, vllngNumeroTipoFormato
    End If

'    vlblnConsulta = True
'    pHabilita False, False, True, False, False, False, False, IIf(optPaciente.Value = True, False, True)
'    Me.cmdConfirmartimbre.Enabled = False
'    If vgIntBanderaTImbradoPendiente = 1 Then
'        Me.txtPendienteTimbre.Text = "Pendiente de timbre fiscal"
'        Me.txtPendienteTimbre.ForeColor = &H0&
'        Me.txtPendienteTimbre.BackColor = &HFFFF&
'        Me.txtPendienteTimbre.Visible = True
'        Me.cmdDelete.Enabled = False
'        Me.cmdComprobante.Enabled = False
'    End If
'    cmdAddenda.Enabled = False
'    If cmdDesglose.Enabled = True Then cmdDesglose.SetFocus
'    sstFacturasCreditos.Enabled = False
'    txtComentario.Enabled = False
'    pHabilitaFrames False, False 'Inhabilitar frames para evitar cambios en la información una vez guardada la nota
    
     InicializaComponentes
     If optNotaCargo.Value Then
        optNotaCargo.SetFocus
     Else
        optNotaCredito.SetFocus
     End If
     vlblnConsulta = False
     vlblnSalir = False
     cmdComprobante.Enabled = False
       
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
End Sub

Private Function blnValidarDiferentesFacturas() As Boolean
On Error GoTo NotificaError
    
    Dim i As Long
    Dim strCadenaUno As String
    
    If grdNotas.Rows > 2 Then
        For i = 1 To grdNotas.Rows - 1
            If i = 1 Then
                strCadenaUno = Trim(CStr(grdNotas.TextMatrix(i, vlintColFactura)))
            Else
                If Trim(grdNotas.TextMatrix(i, vlintColFactura)) <> strCadenaUno Then
                    blnValidarDiferentesFacturas = False
                    Exit Function
                End If
            End If
        Next i
    Else
        strCadenaUno = Trim(CStr(grdNotas.TextMatrix(1, vlintColFactura)))
    End If
    
    strFacturaImpresion = strCadenaUno
    blnValidarDiferentesFacturas = True
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":blnValidarDiferentesFacturas"))
    Unload Me
End Function

Private Sub cmdIncluir_Click()
On Error GoTo NotificaError
    
    Dim vldblCantidad As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim vldblSumaImportes As Double
    Dim vldblSumaDescuentos As Double
    Dim vldblIVACorrespondiente As Double
    Dim dblTotalNota As Double
    Dim vlstrFolioFactura As String
    Dim vlintContFacturas As Integer
    Dim dblIVACantidad As Double

    vlstrFolioFactura = ""
    vlintContFacturas = 0
    
    If optCliente.Value And optNotaCredito.Value And OptMotivoNota(0).Value Then
        If lstListaSeleccionada.ListIndex = -1 Then
            txtCantidad.Text = ""
            grdFactura.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(Format(txtCantidad.Text, vlstrFormato)) = 0 Or (Val(Format(txtCantidad.Text, vlstrFormato)) = 0 And Val(Format(txtDescuento.Text, vlstrFormato)) = 0) Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtCantidad.SetFocus
    Else
        pSeleccionaElemento
        
        If Not fblnExisteConcepto() Then
            strPacienteImpresion = Trim(txtNombrePaciente.Text)
            strFacturaImpresion = Trim(cboFactura.List(cboFactura.ListIndex))
        
            If chkPorcentaje.Value = 0 Or OptMotivoNota(0).Value = True Then
                dblTotalNota = Val(Format(txtCantidad, "##########.00"))
            Else
                dblTotalNota = (Val(Format(txtCantidad, "##########.00")) / 100) * fTotalFactura(grdFactura, Trim(cboFactura.List(cboFactura.ListIndex)))
            End If
            
            vldblCantidad = CDbl(Val(Format(dblTotalNota, vlstrFormato)))
            vldblDescuento = CDbl(Val(Format(txtDescuento.Text, vlstrFormato)))
            
            If optNotaCredito.Value Then
                '-----------------------------------'
                '  N O T A S   D E   C R É D I T O  '
                '-----------------------------------'
                If IIf(OptMotivoNota(1).Value = True, Round((vldblImporteConcepto - vldblDescuentoConcepto), 2) <= Round(vldblCantidad, 2), vldblImporteConcepto = vldblCantidad) Then
                    If OptMotivoNota(1).Value = True Then
                        vldblDescuento = IIf(OptMotivoNota(1).Value = True, 0, vldblDescuentoConcepto)
                        vldblCantidad = grdFactura.TextMatrix(grdFactura.Row, 2) - grdFactura.TextMatrix(grdFactura.Row, 3)
                        vldblIVA = grdFactura.TextMatrix(grdFactura.Row, 12)

                        pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, vldblTmpDescuentoEspecial, vldblIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, grdFactura.TextMatrix(grdFactura.Row, 13), vlblnSeleccionoOtroConcepto
                        pAcumulaFactura cboFactura.List(cboFactura.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, grdFactura.TextMatrix(grdFactura.Row, 13)

                        Call pProcesoAddenda(True)

                        grdFactura.Enabled = True
                        sstFacturasCreditos.TabEnabled(1) = False
                        grdFactura.SetFocus

                        If OptMotivoNota(1).Value = True Then
                            pProrratearNota dblTotalNota
                        End If
                    Else
                        vldblDescuento = IIf(OptMotivoNota(1).Value = True, 0, vldblDescuentoConcepto)
                        vldblIVA = vldblIvaConcepto
                        
                        If OptMotivoNota(1).Value = False Then
                            If vldblImporteConcepto >= dblTotalNota Then
                                dblPorcentajeNota = (dblTotalNota / vldblImporteConcepto)  '(vldblImporteConcepto - vldblDescuentoConcepto)
                            Else
                                dblPorcentajeNota = 1
                            End If
                        End If
                        'dblPorcentajeNota = Round(dblPorcentajeNota, 4)
                        
                        pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, vldblTmpDescuentoEspecial, vldblIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, grdFactura.TextMatrix(grdFactura.Row, 13), vlblnSeleccionoOtroConcepto, , , dblPorcentajeNota
                        pAcumulaFactura cboFactura.List(cboFactura.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, CDbl(grdFactura.TextMatrix(grdFactura.Row, 13))

                        Call pProcesoAddenda(True)
                                                
                        grdFactura.Enabled = True
                        sstFacturasCreditos.TabEnabled(1) = False
                        vldblTmpDescuentoEspecial = 0
                        grdFactura.SetFocus
                    End If
                Else
                    If IIf(OptMotivoNota(1).Value = True, vldblImporteConcepto - vldblDescuentoConcepto < vldblCantidad, vldblImporteConcepto < vldblCantidad) And vlstrCualLista = "CF" Then
                        'La cantidad no puede ser mayor al importe del concepto.
                        MsgBox SIHOMsg(643), vbOKOnly + vbInformation, "Mensaje"
                        txtCantidad.SetFocus
                    Else
                        If vldblDescuento > vldblDescuentoConcepto And vlstrCualLista = "CF" Then
                            'La cantidad no puede ser mayor al importe del concepto.
                            MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                            
                            If txtDescuento.Visible = True And txtDescuento.Enabled = True Then
                                txtDescuento.SetFocus
                            Else
                                txtCantidad.SetFocus
                            End If
                        Else
                            'If (vldblDescuentoConcepto + vldblDescuento) > (vldblImporteConcepto - vldblCantidad) And vlstrCualLista = "CF" Then
                            If vldblDescuento > vldblCantidad And vlstrCualLista = "CF" Then
                                'Los descuentos exceden el importe del concepto.
                                MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                                
                                If txtDescuento.Visible = True And txtDescuento.Enabled = True Then
                                    txtDescuento.SetFocus
                                Else
                                    txtCantidad.SetFocus
                                End If
                            Else
                                If vlstrCualLista = "CF" Then '| Si se está realizando la nota por concepto de facturación
                                    If vldblIvaConcepto > 0 Then
                                        vldblIVA = (vldblCantidad - vldblDescuento) * (vldblIvaConcepto / (vldblImporteConcepto - vldblDescuentoConcepto))
                                    Else
                                        vldblIVA = 0
                                    End If
                                Else '|  Si no se está realizando por concepto de facturación calcula el IVA en base al porcentaje del IVA del elemento seleccionado
                                    vldblIVA = (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA
                                End If
                                                                    
                                If OptMotivoNota(1).Value = False Then
                                    If vldblImporteConcepto >= dblTotalNota Then
                                        dblPorcentajeNota = (dblTotalNota / vldblImporteConcepto)  '(vldblImporteConcepto - vldblDescuentoConcepto)
                                    Else
                                        dblPorcentajeNota = 1
                                    End If
                                End If
                                'dblPorcentajeNota = Round(dblPorcentajeNota, 4)
                                
                                pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, vldblTmpDescuentoEspecial, vldblIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, grdFactura.TextMatrix(grdFactura.Row, 13), vlblnSeleccionoOtroConcepto, , , dblPorcentajeNota
                                pAcumulaFactura cboFactura.List(cboFactura.ListIndex), vldblCantidad, vldblDescuento, vldblIVA, 1, grdFactura.TextMatrix(grdFactura.Row, 13)

                                Call pProcesoAddenda(True)
                                                                
                                If OptMotivoNota(1).Value = True Then
                                    pProrratearNota dblTotalNota
                                End If
                                
                                If blnPestaniaCR Then
                                    sstFacturasCreditos.TabEnabled(1) = False
                                End If
                                
                                vldblTmpDescuentoEspecial = 0
                                pLimpiarTodosList
                                grdFactura.SetFocus
                            End If
                        End If
                    End If
                End If
            Else
                '-------------------------------'
                '  N O T A S   D E   C A R G O  '
                '-------------------------------'
                If (vldblDescuentoConcepto + vldblDescuento) > (vldblImporteConcepto + vldblCantidad) And vlstrCualLista = "CF" Then
                    'Los descuentos exceden el importe del concepto.
                    MsgBox SIHOMsg(644), vbOKOnly + vbInformation, "Mensaje"
                    txtDescuento.SetFocus
                Else
                    If vldblImporteConcepto >= dblTotalNota Then
                        dblPorcentajeNota = (dblTotalNota / vldblImporteConcepto)  '(vldblImporteConcepto - vldblDescuentoConcepto)
                    Else
                        dblPorcentajeNota = 1
                    End If
                    'dblPorcentajeNota = Round(dblPorcentajeNota, 4)
                
                    'dblIVACantidad = (((vldblCantidad / vldblPorcentajeFactura) - (vldblDescuento / vldblPorcentajeFactura)) * vldblPorcentajeIVA) * vldblPorcentajeIVAFactura
                    dblIVACantidad = (((vldblCantidad) - (vldblDescuento)) * vldblPorcentajeIVA)
'                    pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, (vldblCantidad - vldblDescuento) * vldblPorcentajeIVA, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, , , dblPorcentajeNota
                    pIncluyeConcepto lstListaSeleccionada.Text, vldblCantidad, vldblDescuento, 0, dblIVACantidad, lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex), vllngCveCtaIngresos, vllngCveCtaDescuento, vllngCveCtaIva, vlstrCualLista, grdFactura.TextMatrix(grdFactura.Row, 13), vlblnSeleccionoOtroConcepto, , , dblPorcentajeNota
                    pAcumulaFactura cboFactura.List(cboFactura.ListIndex), vldblCantidad, vldblDescuento, dblIVACantidad, 1, grdFactura.TextMatrix(grdFactura.Row, 13)

                    Call pProcesoAddenda(True)
                     
                    blnPestaniaCR = False
                    sstFacturasCreditos.TabEnabled(1) = blnPestaniaCR
                    grdFactura.SetFocus
                End If
            End If
        Else
            pLimpiarTodosList
            grdFactura.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIncluir_Click"))
    Unload Me
End Sub

Public Function fblnExisteConcepto() As Boolean
On Error GoTo NotificaError
    
    Dim vllngContador As Long
    If optCliente.Value And optNotaCredito.Value And OptMotivoNota(0).Value Then
        If lstListaSeleccionada.ListIndex = -1 Then Exit Function
    End If
    fblnExisteConcepto = False
    
    If chkFacturasPaciente.Value = 0 Then
        vllngContador = 1
        Do While Not fblnExisteConcepto And vllngContador <= grdNotas.Rows - 1
            If IIf(grdNotas.TextMatrix(vllngContador, vlintColTipoCargo) = "GE", Val(grdNotas.TextMatrix(vllngContador, vlintColCveConcepto)) * -1, Val(grdNotas.TextMatrix(vllngContador, vlintColCveConcepto))) = lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex) _
                And Trim(grdNotas.TextMatrix(vllngContador, vlintColFactura)) = Trim(cboFactura.List(cboFactura.ListIndex)) Then
                fblnExisteConcepto = True
            End If
            vllngContador = vllngContador + 1
        Loop
    Else
        vllngContador = 1
        Do While Not fblnExisteConcepto And vllngContador <= grdNotas.Rows - 1
            If Val(grdNotas.TextMatrix(vllngContador, vlintColCveConcepto)) = Format(grdCargos.TextMatrix(grdCargos.Row, 6), vlstrFormato) _
                And Trim(grdNotas.TextMatrix(vllngContador, vlintColFactura)) = Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)) Then
                fblnExisteConcepto = True
            End If
            vllngContador = vllngContador + 1
        Loop
    End If

    If fblnExisteConcepto Then
        'Este concepto ya está registrado.
        MsgBox SIHOMsg(319), vbInformation + vbOKOnly, "Mensaje"
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteConcepto"))
    Unload Me
End Function

Public Function fblnExisteConceptoCR() As Boolean
On Error GoTo NotificaError
    
    Dim vllngContador As Long

    fblnExisteConceptoCR = False
    
    vllngContador = 1
    Do While Not fblnExisteConceptoCR And vllngContador <= grdNotas.Rows - 1
        If Val(grdNotas.TextMatrix(vllngContador, vlintColCveConcepto)) = lstListaSeleccionada.ItemData(lstListaSeleccionada.ListIndex) _
            And Trim(grdNotas.TextMatrix(vllngContador, vlintColFactura)) = Trim(cboCreditosDirectos.List(cboCreditosDirectos.ListIndex)) Then
            fblnExisteConceptoCR = True
        End If
        vllngContador = vllngContador + 1
    Loop

    If fblnExisteConceptoCR Then
        'Este concepto ya está registrado.
        MsgBox SIHOMsg(319), vbInformation + vbOKOnly, "Mensaje"
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteConceptoCR"))
    Unload Me
End Function

Public Function fblnExisteCargo() As Boolean
On Error GoTo NotificaError
'(CR) - Revisa que el cargo que se va agregar no se encuentre en el grid de las notas'

    Dim vllngContador As Long

    fblnExisteCargo = False
    
    vllngContador = 1
    Do While Not fblnExisteCargo And vllngContador <= grdNotas.Rows - 1
        If Val(grdNotas.TextMatrix(vllngContador, vlintColCveConcepto)) = Format(grdCargos.TextMatrix(grdCargos.Row, 10), vlstrFormato) _
            And Trim(grdNotas.TextMatrix(vllngContador, vlintColDescripcionConcepto)) = Trim(grdCargos.TextMatrix(grdCargos.Row, 2)) _
            And Val(grdNotas.TextMatrix(vllngContador, 13)) = Val(grdCargos.TextMatrix(grdCargos.Row, 11)) Then
            fblnExisteCargo = True
        End If
        vllngContador = vllngContador + 1
    Loop

    If fblnExisteCargo Then
        'Este concepto ya está registrado.
        MsgBox SIHOMsg(319), vbInformation + vbOKOnly, "Mensaje"
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteConceptoCR"))
    Unload Me
End Function

Private Sub pIncluyeConcepto(vlStrConcepto As String, _
                             vldblCantidad As Double, _
                             vldblDescuento As Double, _
                             vldblDescuentoEspecial As Double, _
                             vldblIVA As Double, _
                             vllngCveConcepto As Long, _
                             vllngCtaIngresos As Long, _
                             vllngCtaDescuentos As Long, _
                             vllngCtaIVA As Long, _
                             vlstrTipoCargo As String, _
                             vlstrTipoCambio As Double, _
                             vlblnSeleccionoOtroConcepto As Boolean, _
                             Optional vllngintFolioDocumento As Long, _
                             Optional vllngCveDepartamento As Long, _
                             Optional vldblPorcentajeNota As Double)
On Error GoTo NotificaError
    
    If Trim(grdNotas.TextMatrix(1, 2)) = "" Then
        grdNotas.Row = 1
    Else
        grdNotas.Rows = grdNotas.Rows + 1
        grdNotas.Row = grdNotas.Rows - 1
    End If
    
    With grdNotas
        If chkFacturasPaciente.Value = 0 Then
            .TextMatrix(.Row, vlintColFactura) = IIf(sstFacturasCreditos.Tab = 0, cboFactura.List(cboFactura.ListIndex), cboCreditosDirectos.List(cboCreditosDirectos.ListIndex))
        Else
            .TextMatrix(.Row, vlintColFactura) = cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)
        End If
        .TextMatrix(.Row, vlintColDescripcionConcepto) = vlStrConcepto
        .TextMatrix(.Row, vlintColCantidad) = FormatCurrency(vldblCantidad, 2)
        .TextMatrix(.Row, vlintColDescuento) = FormatCurrency(vldblDescuento, 2)
        .TextMatrix(.Row, vlintColIVA) = FormatCurrency(vldblIVA, 2)
        .TextMatrix(.Row, vlintColCveConcepto) = IIf(vlstrTipoCargo = "GE", vllngCveConcepto * -1, vllngCveConcepto)
        .TextMatrix(.Row, vlintColCtaIngresos) = vllngCtaIngresos
        .TextMatrix(.Row, vlintColCtaDescuentos) = vllngCtaDescuentos
        .TextMatrix(.Row, vlintColCtaIVA) = vllngCtaIVA
        .TextMatrix(.Row, vlintColTipoCargo) = vlstrTipoCargo
        .TextMatrix(.Row, vlintColTipoNotaFARCDetalle) = IIf(sstFacturasCreditos.Tab = 0, "FA", "MA")
        .TextMatrix(.Row, 13) = vllngintFolioDocumento
        .TextMatrix(.Row, 14) = vllngCveDepartamento
        .TextMatrix(.Row, vlintColIVANotaSinRedondear) = FormatCurrency(vldblIVA, 15)
        .TextMatrix(.Row, 20) = vldblPorcentajeNota
        .TextMatrix(.Row, vlintColPorcentajeFactura) = IIf(vldblIVA > 0, vldblPorcentajeFacturaSIgrava, vldblPorcentajeFacturaNOgrava)
'        .TextMatrix(.Row, vlintColPorcentajeFacturaConDescuento) = IIf(vldblIVA > 0, vldblPorcentajeFacturaConDescuentoSIgrava, vldblPorcentajeFacturaConDescuentoNOgrava)
        .TextMatrix(.Row, vlintColPorcentajeFacturaConDescuento) = IIf(vldblIVA > 0, vldblPorcentajeFacturaSIgrava, vldblPorcentajeFacturaNOgrava)
        .TextMatrix(.Row, vlintColTipoCambio) = vlstrTipoCambio
        .TextMatrix(.Row, vlintColSeleccionoOtroConcepto) = vlblnSeleccionoOtroConcepto
        .TextMatrix(.Row, vlintColDescuentoEspecial) = vldblDescuentoEspecial
    End With
    
    vldblSubTotalTemporal = vldblSubTotalTemporal + vldblCantidad
    txtSubtotal.Text = FormatCurrency(vldblSubTotalTemporal, 2)
    
    vldblDescuentoTemporal = vldblDescuentoTemporal + vldblDescuento
    txtDescuentoTot.Text = FormatCurrency(vldblDescuentoTemporal, 2)
    
    'vldblIvaTemporal = vldblIvaTemporal + FormatCurrency(vldblIVA, 2)
    'vldblIvaTemporal = vldblIvaTemporal + vldblIVA  'Se cambio la linea por la de arriba ya que no cuadraba la nota (caso 9686)
    vldblIvaTemporal = vldblIvaTemporal + Round(vldblIVA, 2)
    
    txtIVA.Text = FormatCurrency(vldblIvaTemporal, 2)

    vldTotal = vldblSubTotalTemporal - vldblDescuentoTemporal + vldblIvaTemporal
    txtTotal.Text = FormatCurrency(vldTotal, 2)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIncluyeConcepto"))
    Unload Me
End Sub

Private Sub pAcumulaFactura(vlstrFactura As String, _
                            vldblCantidad As Double, _
                            vldblDescuento As Double, _
                            vldblIVA As Double, _
                            vldblFactor As Double, _
                            vlstrTipoCambio As Double, _
                            Optional vldtmFecha As Date)

    Dim vlintContador As Integer
    Dim vlintRenglon As Integer
    Dim vlblnEncontrada As Boolean
    Dim vlintTamaño As Integer
    Dim vlintNumeroRenglon  As Integer
    
    If Trim(aFacturas(0).vlstrFolioFactura) <> "" Then
        vlblnEncontrada = False
        vlintContador = 0
        Do While vlintContador <= UBound(aFacturas(), 1) And Not vlblnEncontrada
            If Trim(aFacturas(vlintContador).vlstrFolioFactura) = Trim(vlstrFactura) Then
                vlintRenglon = vlintContador
                vlblnEncontrada = True
            End If
            vlintContador = vlintContador + 1
        Loop
        
        If Not vlblnEncontrada Then
            ReDim Preserve aFacturas(UBound(aFacturas(), 1) + 1)
            vlintRenglon = UBound(aFacturas(), 1)
        End If
    Else
        vlintRenglon = 0
    End If
    
    If vldblFactor <> -1 Then
        'Significa que se está agregando la factura
        'If chkFacturasPaciente.Value = 0 Then
        '    aFacturas(vlintRenglon).vlstrFolioFactura = IIf(sstFacturasCreditos.Tab = 0, cboFactura.List(cboFactura.ListIndex), cboCreditosDirectos.List(cboCreditosDirectos.ListIndex))
        'Else
        '    aFacturas(vlintRenglon).vlstrFolioFactura = cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)
        'End If
        '- (CR) Modificado para caso 7578: No permitía guardar nota aplicada a más de una factura -'
        aFacturas(vlintRenglon).vlstrFolioFactura = vlstrFactura 'Agregar siempre el folio enviado como parámetro
    End If
    
'    aFacturas(vlintRenglon).vldblSubtotal = aFacturas(vlintRenglon).vldblSubtotal + (vldblCantidad * vldblFactor)
'    aFacturas(vlintRenglon).vldblDescuento = aFacturas(vlintRenglon).vldblDescuento + (vldblDescuento * vldblFactor)
'    aFacturas(vlintRenglon).vldblIVA = aFacturas(vlintRenglon).vldblIVA + (vldblIVA * vldblFactor)
    
    'aFacturas(vlintRenglon).vldblSubtotal = Round(aFacturas(vlintRenglon).vldblSubtotal + (vldblCantidad * vldblFactor), 2)
    aFacturas(vlintRenglon).vldblSubtotal = (aFacturas(vlintRenglon).vldblSubtotal + (vldblCantidad * vldblFactor))
    aFacturas(vlintRenglon).vldblDescuento = Round(aFacturas(vlintRenglon).vldblDescuento + (vldblDescuento * vldblFactor), 2)
    'aFacturas(vlintRenglon).vldblIVA = Round(aFacturas(vlintRenglon).vldblIVA, 2) + (Round(vldblIVA, 2) * vldblFactor)
    aFacturas(vlintRenglon).vldblIVA = (aFacturas(vlintRenglon).vldblIVA) + ((Round(vldblIVA, 2)) * vldblFactor)
    aFacturas(vlintRenglon).vldblTipoCambio = vlstrTipoCambio
    
    If optCliente.Value = True Then
        aFacturas(vlintRenglon).vldtmFecha = CDate(Format(IIf(sstFacturasCreditos.Tab = 0, lblFechaFactura.Caption, grdCreditoDirecto.TextMatrix(1, 1)), "dd/mm/yyyy"))
    Else
        '- (CR) Modificado para que se tome la fecha del cargo en notas para pacientes -'
        aFacturas(vlintRenglon).vldtmFecha = CDate(Format(IIf(sstFacturasCreditos.Tab = 0, vldtmFecha, grdCreditoDirecto.TextMatrix(1, 1)), "dd/mm/yyyy"))
    End If
    
    'Si la factura ya se eliminó del grid, se elimina también del arreglo aFacturas
    If vldblFactor = -1 And aFacturas(vlintRenglon).vldblSubtotal <= 0 And aFacturas(vlintRenglon).vldblDescuento <= 0 And aFacturas(vlintRenglon).vldblIVA <= 0 Then
        If UBound(aFacturas()) > 0 Then
            vlintNumeroRenglon = 0
            For vlintTamaño = 0 To UBound(aFacturas())
                If vlintTamaño <> vlintRenglon Then
                    ReDim Preserve aFacturasTemporal(vlintNumeroRenglon + 1)
                    aFacturasTemporal(vlintNumeroRenglon).vlstrFolioFactura = aFacturas(vlintTamaño).vlstrFolioFactura
                    aFacturasTemporal(vlintNumeroRenglon).vldblSubtotal = aFacturas(vlintTamaño).vldblSubtotal
                    aFacturasTemporal(vlintNumeroRenglon).vldblDescuento = aFacturas(vlintTamaño).vldblDescuento
                    aFacturasTemporal(vlintNumeroRenglon).vldblIVA = aFacturas(vlintTamaño).vldblIVA
                    aFacturasTemporal(vlintNumeroRenglon).vldtmFecha = aFacturas(vlintTamaño).vldtmFecha
                    vlintNumeroRenglon = vlintNumeroRenglon + 1
                End If
            Next
            aFacturas = aFacturasTemporal
            ReDim Preserve aFacturas(vlintNumeroRenglon - 1)
        End If
    End If
    
End Sub

Private Sub cmdPrimerRegistro_Click()
    grdBusqueda.Row = 1
    pMostrarNota grdBusqueda.Row, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTipoNotaFACR)
    If optCliente.Value = True Then
        pHabilita True, True, True, True, True, False, IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) <> "A", False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "A", True, False)
    Else
        pHabilita True, True, True, True, True, False, IIf((grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "C" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC") Or (grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P" And Not fblnNotaAutomatica(lblFolio.Caption)), False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P", True, False)
    End If
End Sub

Private Sub cmdSiguienteRegistro_Click()
    If grdBusqueda.Row < grdBusqueda.Rows - 1 Then
        grdBusqueda.Row = grdBusqueda.Row + 1
    End If
    pMostrarNota grdBusqueda.Row, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTipoNotaFACR)
    If optCliente.Value = True Then
        pHabilita True, True, True, True, True, False, IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) <> "A", False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "A", True, False)
    Else
        pHabilita True, True, True, True, True, False, IIf((grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "C" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC") Or (grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P" And Not fblnNotaAutomatica(lblFolio.Caption)), False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P", True, False)
    End If
End Sub

Private Sub cmdUltimoRegistro_Click()
    grdBusqueda.Row = grdBusqueda.Rows - 1
    pMostrarNota grdBusqueda.Row, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTipoNotaFACR)
    If optCliente.Value = True Then
        pHabilita True, True, True, True, True, False, IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) <> "A", False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "A", True, False)
    Else
        pHabilita True, True, True, True, True, False, IIf((grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "C" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC") Or (grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P" And Not fblnNotaAutomatica(lblFolio.Caption)), False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P", True, False)
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
   
    Dim vPrinter As Printer
    Dim rsImpresoras As ADODB.Recordset

    vgstrNombreForm = Me.Name

    vllngNumeroTipoFormato = flngFormatoDepto(vgintNumeroDepartamento, cintTipoDocumento, "*")

    sstCargos.Caption = ""

    'Validación de Impresoras
    vlstrSentencia = "SELECT chrNombreImpresora Impresora FROM ImpresoraDepartamento WHERE chrTipo = 'NO' AND smiCveDepartamento = " & Trim(Str(vgintNumeroDepartamento))
    Set rsImpresoras = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsImpresoras.RecordCount > 0 Then
        For Each vPrinter In Printers
            If UCase(Trim(vPrinter.DeviceName)) = UCase(Trim(rsImpresoras!Impresora)) Then
                Set Printer = vPrinter
            End If
        Next
    Else
        'No se tiene asignada una impresora en la cual imprimir las notas
        MsgBox SIHOMsg(993), vbCritical, "Mensaje"
        vlblnSalir = True
        Unload Me
        Exit Sub
    End If
    rsImpresoras.Close
    
    If vlblnErrorManejoFolio Then
        MsgBox SIHOMsg(383) & " Manejo de la serie de folios de notas.", vbOKOnly + vbExclamation, "Mensaje"
        vlblnSalir = True
        Unload Me
        Exit Sub
    End If
    
    If vllngNumeroTipoFormato = 0 Then
        MsgBox SIHOMsg(383) & " Formato para impresión de nota.", vbOKOnly + vbExclamation, "Mensaje"
        vlblnSalir = True
        Unload Me
        Exit Sub
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub pLimpiaGridFactura()
On Error GoTo NotificaError

    With grdFactura
       .Clear
       .Cols = 15
       .Rows = 2
       .FixedCols = 1
       .FixedRows = 1
       .FormatString = "|Descripción|Cantidad|Descuento|IVA|Total"
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridFactura"))
    Unload Me
End Sub

Private Sub pLimpiaGridFacturaPaciente()
On Error GoTo NotificaError

    With grdCargos
        .Clear
        .Cols = 14
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción|Cantidad|Descuento|IVA|Total"
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridFacturaPaciente"))
    Unload Me
End Sub

Private Sub pConfiguraGridFactura()
On Error GoTo NotificaError

   With grdFactura
      .FormatString = "|Descripción|Cantidad|Descuento|IVA|Total"
      .ColWidth(0) = 100     ' Fixed              (0)
      .ColWidth(1) = 5100    ' Descripcion        (1)
      .ColWidth(2) = 1200    ' Cantidad           (2)
      .ColWidth(3) = 1100    ' Descuento          (3)
      .ColWidth(4) = 1100    ' IVA                (3)
      .ColWidth(5) = 1200    ' Total              (4)
      .ColWidth(6) = 0       ' CveConceptoNota
      .ColWidth(7) = 0       ' CuentaIngreso
      .ColWidth(8) = 0       ' CuentaDescuento
      .ColWidth(9) = 0       ' No se utiliza
      .ColWidth(10) = 0      ' % de IVA del concepto
      .ColWidth(11) = 0      ' Total sin redondear
      .ColWidth(12) = 0      ' IVA sin redondear
      .ColWidth(13) = 0      ' Tipo de cambio
      .ColWidth(14) = 0     'DESCUENTO ESPECIAL
      
      .ColAlignment(1) = flexAlignLeftCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
      .ColAlignmentFixed(1) = flexAlignCenterCenter
      .ColAlignmentFixed(2) = flexAlignCenterCenter
      .ColAlignmentFixed(3) = flexAlignCenterCenter
      .ColAlignmentFixed(4) = flexAlignCenterCenter
      .ColAlignmentFixed(5) = flexAlignCenterCenter
      .ScrollBars = flexScrollBarBoth
   End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridFactura"))
    Unload Me
End Sub

Private Sub pConfiguraGridCargos()
On Error GoTo NotificaError
            
    If chkFacturasPaciente.Value = 0 Then
        With grdCargos
           .FormatString = "|Fecha|Descripción|Cantidad|Precio|Descuento|IVA|Total|Departamento que cargó"
           .ColWidth(0) = 100     'Fixed
           .ColWidth(1) = 1200    'Fecha del cargo
           .ColWidth(2) = 3000    'Descripción del cargo
           .ColWidth(3) = 1000    'Cantidad
           .ColWidth(4) = 1000    'Precio
           .ColWidth(5) = 1000    'Descuento
           .ColWidth(6) = 800     'IVA
           .ColWidth(7) = 800     'Total
           .ColWidth(8) = 2500    'Departamento que cargó
           .ColWidth(9) = 0
           .ColWidth(10) = 0
           .ColWidth(11) = 0
           .ColWidth(12) = 0
           .ColWidth(13) = 0      'Tipo de cambio
                              
           .ColAlignment(1) = flexAlignLeftCenter
           .ColAlignment(2) = flexAlignLeftCenter
           .ColAlignment(3) = flexAlignRightCenter
           .ColAlignment(4) = flexAlignRightCenter
           .ColAlignment(5) = flexAlignRightCenter
           .ColAlignment(6) = flexAlignRightCenter
           .ColAlignment(7) = flexAlignRightCenter
           .ColAlignmentFixed(1) = flexAlignCenterCenter
           .ColAlignmentFixed(2) = flexAlignCenterCenter
           .ColAlignmentFixed(3) = flexAlignCenterCenter
           .ColAlignmentFixed(4) = flexAlignCenterCenter
           .ColAlignmentFixed(5) = flexAlignCenterCenter
           .ColAlignmentFixed(6) = flexAlignCenterCenter
           .ColAlignmentFixed(7) = flexAlignCenterCenter
           .ColAlignmentFixed(8) = flexAlignCenterCenter
           
           .ScrollBars = flexScrollBarBoth
        End With
    Else
        With grdCargos
            .FormatString = "|Descripción|Cantidad|Descuento|IVA|Total"
            .ColWidth(0) = 100     'Fixed              (0)
            .ColWidth(1) = 6400    'Descripcion        (1)
            .ColWidth(2) = 1200    'Cantidad           (2)
            .ColWidth(3) = 1100    'Descuento          (3)
            .ColWidth(4) = 1100    'IVA                (3)
            .ColWidth(5) = 1200    'Total              (4)
            .ColWidth(6) = 0       ' CveConceptoNota
            .ColWidth(7) = 0       ' CuentaIngreso
            .ColWidth(8) = 0       ' CuentaDescuento
            .ColWidth(9) = 0       'No se utiliza
            .ColWidth(10) = 0       ' % de IVA del concepto
            .ColWidth(13) = 0      'Tipo de cambio
      
            .ColAlignment(1) = flexAlignLeftCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignRightCenter
            .ColAlignment(4) = flexAlignRightCenter
            .ColAlignment(5) = flexAlignRightCenter
            .ColAlignmentFixed(1) = flexAlignCenterCenter
            .ColAlignmentFixed(2) = flexAlignCenterCenter
            .ColAlignmentFixed(3) = flexAlignCenterCenter
            .ColAlignmentFixed(4) = flexAlignCenterCenter
            .ColAlignmentFixed(5) = flexAlignCenterCenter
            
            .ScrollBars = flexScrollBarBoth
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    Unload Me
End Sub

Private Sub pLimpiaGridCargos()
On Error GoTo NotificaError

   If chkFacturasPaciente.Value = 0 Then
        With grdCargos
           .Clear
           .Cols = 14
           .Rows = 2
           .FixedCols = 1
           .FixedRows = 1
           .FormatString = "|Fecha|Descripción|Cantidad|Precio|Descuento|IVA|Total|Departamento que cargó"
        End With
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridCargos"))
    Unload Me
End Sub

Private Sub pConfiguraGridCreditos()
On Error GoTo NotificaError

   With grdCreditoDirecto
      .FormatString = "|Fecha|Monto crédito|Cantidad pagada|Saldo"
      .ColWidth(0) = 100     'Fixed                (0)
      .ColWidth(1) = 1200    'Fecha                (1)
      .ColWidth(2) = 1500    'Cantidad crédito     (2)
      .ColWidth(3) = 1500    'Cantidad pagada      (3)
      .ColWidth(4) = 1500    'Saldo                (4)
      
      .ColAlignment(1) = flexAlignRightCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(3) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignmentFixed(1) = flexAlignCenterCenter
      .ColAlignmentFixed(2) = flexAlignCenterCenter
      .ColAlignmentFixed(3) = flexAlignCenterCenter
      .ColAlignmentFixed(4) = flexAlignCenterCenter
      .ScrollBars = flexScrollBarBoth
   End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCreditos"))
    Unload Me
End Sub

Private Sub pConfiguraGridNota(strFacturaCR As String)
On Error GoTo NotificaError
      
   With grdNotas
      .FormatString = "|" & IIf(Trim(strFacturaCR) = "FA", "Factura", "Crédito directo") & "|Descripción|Cantidad|Descuento|IVA|Total|Fecha"
      .ColWidth(0) = 100
      .ColWidth(vlintColFactura) = IIf(optCliente.Value = True Or chkFacturasPaciente.Value = 1, 1500, 0)
      .ColWidth(vlintColDescripcionConcepto) = 6000
      .ColWidth(vlintColCantidad) = 1200
      .ColWidth(vlintColDescuento) = IIf(OptMotivoNota(0).Value Or chkFacturasPaciente = True, 1100, 0)
      .ColWidth(vlintColIVA) = 1100
      .ColWidth(vlintColTotal) = 0
      .ColWidth(vlintColCveConcepto) = 0
      .ColWidth(vlintColCtaIngresos) = 0
      .ColWidth(vlintColCtaDescuentos) = 0
      .ColWidth(vlintColCtaIVA) = 0
      .ColWidth(vlintColTipoCargo) = 0
      .ColWidth(vlintColTipoNotaFARCDetalle) = 0
      .ColWidth(13) = 0
      .ColWidth(14) = 0
      .ColWidth(15) = 0
      .ColWidth(16) = 0
      .ColWidth(17) = 0
      .ColWidth(18) = 0
      .ColWidth(19) = 0
      .ColWidth(20) = 0
      .ColWidth(21) = 0
      .ColWidth(vlintColPorcentajeFactura) = 0
      .ColWidth(vlintColPorcentajeFacturaConDescuento) = 0
      .ColWidth(vlintColTipoCambio) = 0
      .ColWidth(vlintColSeleccionoOtroConcepto) = 0
      .ColWidth(vlintColDescuentoEspecial) = 0
      
      .ColAlignment(vlintColFactura) = flexAlignLeftCenter
      .ColAlignment(vlintColDescripcionConcepto) = flexAlignLeftCenter
      .ColAlignment(vlintColCantidad) = flexAlignRightCenter
      .ColAlignment(vlintColDescuento) = flexAlignRightCenter
      .ColAlignment(vlintColIVA) = flexAlignRightCenter
      .ColAlignment(vlintColTotal) = flexAlignRightCenter
      
      .ColAlignmentFixed(vlintColFactura) = flexAlignCenterCenter
      .ColAlignmentFixed(vlintColDescripcionConcepto) = flexAlignCenterCenter
      .ColAlignmentFixed(vlintColCantidad) = flexAlignCenterCenter
      .ColAlignmentFixed(vlintColDescuento) = flexAlignCenterCenter
      .ColAlignmentFixed(vlintColIVA) = flexAlignCenterCenter
      .ColAlignmentFixed(vlintColTotal) = flexAlignCenterCenter
      
      .ScrollBars = flexScrollBarBoth
   End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridNota"))
    Unload Me
End Sub

Private Sub pLimpiaGridNota()
On Error GoTo NotificaError

     With grdNotas
        .Clear
        .Cols = 27
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Folio|Descripción|Cantidad|Descuento|IVA|Total"
    End With
    
    ReDim aFacturas(0)
    aFacturas(0).vlstrFolioFactura = ""

    txtSubtotal.Text = ""
    txtDescuentoTot.Text = ""
    txtIVA.Text = ""
    txtTotal.Text = ""
    
    InicializaVariables

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGridNota"))
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
   
   If KeyCode = 82 And Shift = 4 Then
      optNotaCredito.SetFocus
      optNotaCredito_Click
   ElseIf KeyCode = 65 And Shift = 4 Then
      optNotaCargo.SetFocus
      optNotaCargo_Click
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
        fraCliente.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
        
    optMostrarSolo(0).Value = True
    
    vlstrFormato = "###############.00"
    
    vlblnSerieUnica = fblnSerieUnica
    
    blnPestaniaFac = False
    blnPestaniaCR = False
    
    'P A R A M E T R O S
    vlblnErrorManejoFolio = False
    
    If frsRegresaRs("Select ISNULL(intFolioUnicoNotas,0) from CcParametro").RecordCount <> 0 Then
        vgintFolioUnico = frsRegresaRs(" Select intFolioUnicoNotas from CcParametro ").Fields(0)
    Else
        vlblnErrorManejoFolio = True
    End If
        
    pLimpiaTodo
    pHabilita False, False, True, False, False, False, False, False
    
    pCargarOpciones
    
    If optPaciente.Value = True Then pCargaTipoPaciente
    
    ReDim aFoliosFacturas(0)
    
    chkFacturasPagadas.Visible = False
    
    cboFacturasPaciente.Visible = False
    lbFacturasPaciente.Visible = False
    vlblnlimpiaNotas = True
    
    pCargaUsosCFDI
    pCargaFormasPago
    pCargaMetodoPago
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    If vlblnSalir = True Then
        Unload Me
        vlblnSalir = False
    Else
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            If sstObj.Tab = 0 And vlblnConsulta = False And cmdGrabarRegistro.Enabled = False Then
                 Unload Me
            Else
                InicializaComponentes
                If optNotaCargo.Value Then
                    optNotaCargo.SetFocus
                Else
                    optNotaCredito.SetFocus
                End If
                
                Cancel = True
                vlblnConsulta = False
                vlblnSalir = False
                cmdComprobante.Enabled = False
            End If
        Else
            Cancel = True
        End If
   End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub pMostrarNota(vlintRenglon As Integer, strFacturaCR As String)
    Dim rsDetalleNota As New ADODB.Recordset
    Dim rsDetalleNotaElectronica As New ADODB.Recordset
    Dim strParametrosSP As String

    vlblnConsulta = True
    
    fraEncabezadoNota.Enabled = False
    
    lngIDnota = Val(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColIdNota))
    
    optNotaCredito.Value = grdBusqueda.TextMatrix(vlintRenglon, vlintColchrTipo) = "CR"
    optNotaCargo.Value = grdBusqueda.TextMatrix(vlintRenglon, vlintColchrTipo) = "CA"
    
    If optNotaCredito.Value = True Then
        If grdBusqueda.TextMatrix(grdBusqueda.Row, 7) = "Error de facturación" Then
            OptMotivoNota(0).Value = True
            OptMotivoNota(1).Value = False
        Else
            OptMotivoNota(1).Value = True
            OptMotivoNota(0).Value = False
        End If
    Else
        OptMotivoNota(1).Value = False
        OptMotivoNota(0).Value = False
    End If
    
    lblFolio.Caption = grdBusqueda.TextMatrix(vlintRenglon, vlintColFolioNota)
    mskFecha.Mask = ""
    mskFecha.Text = Format(grdBusqueda.TextMatrix(vlintRenglon, vlintColFechaNota), "dd/mm/yyyy")
    mskFecha.Mask = "##/##/####"
      
    lblEstatus.ForeColor = IIf(grdBusqueda.TextMatrix(vlintRenglon, vlintColchrEstatus) = "C", vllngColorCanceladas, vllngColorActivas)
    
    lblEstatus.Caption = grdBusqueda.TextMatrix(vlintRenglon, vlintColEstadoNota)
    txtCveCliente.Text = grdBusqueda.TextMatrix(vlintRenglon, vlintColNumCliente)
    txtCliente.Text = grdBusqueda.TextMatrix(vlintRenglon, vlintColNombreCliente)
    txtDomicilio.Text = grdBusqueda.TextMatrix(vlintRenglon, vlintColDomicilioCliente)
    txtRFC.Text = grdBusqueda.TextMatrix(vlintRenglon, vlintColRFCCliente)
        
    fraDetalleFactura.Visible = False
    fraConcepto.Visible = False
    
    freDetalleNota.Height = clngfreDetalleNotaHeightConsulta
    
    If optCliente.Value = True Then
        freDetalleNota.Top = clngfreDetalleNotaTopConsulta
    Else
        freDetalleNota.Top = 2700
    End If
    
    grdNotas.Height = clnggrdNotasHeightConsulta
    
    txtSubtotal.Top = clngtxtSubtotalTopConsulta
    txtDescuentoTot.Top = clngtxtSubtotalTopConsulta + 340
    txtIVA.Top = clngtxtSubtotalTopConsulta + (340 * 2)
    txtTotal.Top = clngtxtSubtotalTopConsulta + (340 * 3)

    lblSubtotal.Top = clnglblSubtotalTopConsulta
    lblDescuento.Top = clnglblSubtotalTopConsulta + 340
    lblIVA.Top = clnglblSubtotalTopConsulta + (340 * 2)
    lblTotal.Top = clnglblSubtotalTopConsulta + (340 * 3)
    txtComentario.Top = clngtxtSubtotalTopConsulta + 2750
    
    lblObservaciones.Top = clngtxtSubtotalTopConsulta + 50

    txtComentario.Locked = True
    
    pLimpiaGridNota
    pConfiguraGridNota strFacturaCR
    vgstrParametrosSP = grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColIdNota)
    Set rsDetalleNota = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelNotaCreditoCargoDet")
    If rsDetalleNota.RecordCount <> 0 Then
        With grdNotas
            .Row = 1
            Do While Not rsDetalleNota.EOF
                .TextMatrix(.Row, vlintColFactura) = rsDetalleNota!chrfoliofactura
                .TextMatrix(.Row, vlintColDescripcionConcepto) = IIf(IsNull(rsDetalleNota!DescripcionDetalle), "", rsDetalleNota!DescripcionDetalle)
                .TextMatrix(.Row, vlintColCantidad) = FormatCurrency(rsDetalleNota!MNYCantidad, 2)
                .TextMatrix(.Row, vlintColDescuento) = IIf(IsNull(rsDetalleNota!MNYDESCUENTO), 0, FormatCurrency(rsDetalleNota!MNYDESCUENTO, 2))
                .TextMatrix(.Row, vlintColIVA) = FormatCurrency(rsDetalleNota!MNYIVA, 2)
                .TextMatrix(.Row, vlintColTotal) = FormatCurrency(rsDetalleNota!MNYCantidad - IIf(IsNull(rsDetalleNota!MNYDESCUENTO), 0, rsDetalleNota!MNYDESCUENTO) + rsDetalleNota!MNYIVA, 2)
                .TextMatrix(.Row, vlintColCveConcepto) = rsDetalleNota!intConcepto
                .TextMatrix(.Row, vlintColCtaIngresos) = IIf(IsNull(rsDetalleNota!intCuentaIngreso), "", rsDetalleNota!intCuentaIngreso)
                .TextMatrix(.Row, vlintColCtaDescuentos) = IIf(IsNull(rsDetalleNota!intCuentaDescuento), "", rsDetalleNota!intCuentaDescuento)
                .TextMatrix(.Row, vlintColCtaIVA) = rsDetalleNota!intCuentaIVA
                .TextMatrix(.Row, vlintColTipoNotaFARCDetalle) = rsDetalleNota!TipoNota
                txtComentario.Text = rsDetalleNota!Comentario
                If optPaciente.Value = True Then
                    If rsDetalleNota!TipoPaciente = "I" Then OptTipoPaciente(0).Value = True
                    If rsDetalleNota!TipoPaciente = "E" Then OptTipoPaciente(1).Value = True
                    
                    vlblnConsulta = True
                End If
                .Rows = .Rows + 1
                .Row = .Rows - 1
                rsDetalleNota.MoveNext
            Loop
            .Rows = .Rows - 1
        End With
    End If

    txtSubtotal.Text = FormatCurrency(Val(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColSubtotalNota)), 2)
    txtDescuentoTot.Text = FormatCurrency(Val(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColDescuentoNota)), 2)
    txtIVA.Text = FormatCurrency(Val(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColIVANota)), 2)
    txtTotal.Text = FormatCurrency(Val(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTotalNota)), 2)
    
    Me.txtPendienteTimbre.Visible = False
    
    If grdBusqueda.TextMatrix(vlintRenglon, vlintcolPTimbre) = "1" Then
       cmdComprobante.Enabled = False
       Me.txtPendienteTimbre.Visible = True
       Me.txtPendienteTimbre.Text = "Pendiente de timbre fiscal"
       Me.txtPendienteTimbre.ForeColor = &H0&
       Me.txtPendienteTimbre.BackColor = &HFFFF&
       Me.cmdDelete.Enabled = False
       Me.cmdConfirmarTimbre.Enabled = True
    Else
        Set rsDetalleNotaElectronica = frsRegresaRs("SELECT * FROM GnComprobanteFiscalDigital INNER JOIN CCnota ON GnComprobanteFiscalDigital.INTCOMPROBANTE = CCNota.INTCONSECUTIVO AND GnComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = CCNota.CHRTIPO WHERE CCNota.ChrFolioNota = '" & Trim(lblFolio.Caption) & "'")
        If rsDetalleNotaElectronica.RecordCount <> 0 Then
            vlstrTipoCFD = IIf(IsNull(rsDetalleNotaElectronica!INTNUMEROAPROBACION), "CFDi", "CFD")
            cmdComprobante.Enabled = True
            frmFechaCancelacionNotas.vlblnActivaMotivo = True
        Else
            cmdComprobante.Enabled = False
            frmFechaCancelacionNotas.vlblnActivaMotivo = False
        End If

        If grdBusqueda.TextMatrix(vlintRenglon, vlintColPFacSAT) = "PC" Then
           Me.txtPendienteTimbre.Visible = True
           Me.txtPendienteTimbre.Text = "Pendiente de cancelarse ante el SAT"
           Me.txtPendienteTimbre.ForeColor = &HFF&
           Me.txtPendienteTimbre.BackColor = &HC0E0FF
        Else
            If grdBusqueda.TextMatrix(vlintRenglon, vlintColPFacSAT) = "PA" Then
                Me.txtPendienteTimbre.Visible = True
                Me.txtPendienteTimbre.Text = "Pendiente de autorización"
                Me.txtPendienteTimbre.ForeColor = &H80000005
                Me.txtPendienteTimbre.BackColor = &H80FF&
            Else
                If grdBusqueda.TextMatrix(vlintRenglon, vlintColPFacSAT) = "CR" Then
                    Me.txtPendienteTimbre.Visible = True
                    Me.txtPendienteTimbre.Text = "Cancelación rechazada"
                    Me.txtPendienteTimbre.ForeColor = &H80000005
                    Me.txtPendienteTimbre.BackColor = &HFF&
                End If
            End If
        End If

        Me.cmdConfirmarTimbre.Enabled = False
    End If
    
    FraMetodoForma.Visible = False
    cboUsoCFDI.Visible = False
    cboMetodoPago.Visible = False
    cboFormaPago.Visible = False
     
    '- Agregados para caso 7994 -'
    vgblnCancelada = grdBusqueda.TextMatrix(vlintRenglon, vlintColchrEstatus) = "C"
    vgblnCFDI = fblnNotaCFDI(lngIDnota, grdBusqueda.TextMatrix(vlintRenglon, vlintColchrTipo))
End Sub


Private Sub grdBusqueda_DblClick()
If grdBusqueda.MouseCol > 0 And grdBusqueda.MouseRow > 0 Then
    If Trim(grdBusqueda.TextMatrix(1, 1)) <> "" Then
       If grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "XX" Then
         'Este comprobante procede de un proceso incompleto en el sistema por lo que no puede ser consultado.
          MsgBox SIHOMsg(1277), vbOKOnly + vbExclamation, "Mensaje"
       Exit Sub
       End If
    
        sstObj.Tab = 0
        pMostrarNota grdBusqueda.Row, grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColTipoNotaFACR)
        
        If optCliente.Value = True Then
            pHabilita True, True, True, True, True, False, IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) <> "A", False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "A", True, False)
        Else
            pHabilita True, True, True, True, True, False, IIf((grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "C" Or grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColPFacSAT) = "PC") Or (grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P" And Not fblnNotaAutomatica(lblFolio.Caption)), False, True), IIf(grdBusqueda.TextMatrix(grdBusqueda.Row, vlintColchrEstatus) = "P", True, False)
        End If
        
        cmdBuscar.SetFocus
        
        vlblnConsulta = True
        
        If optCliente.Value = True Then
            sstCargos.Visible = False
            sstFacturasCreditos.Visible = False
        End If
        
        If grdBusqueda.TextMatrix(grdBusqueda.Row, 7) = "Error de facturación" Then
            OptMotivoNota(0).Value = True
        Else
            OptMotivoNota(1).Value = True
        End If
        
        If grdBusqueda.TextMatrix(grdBusqueda.Row, 7) = "" Then
            OptMotivoNota(0).Value = False
            OptMotivoNota(1).Value = False
        End If
        
        If Trim(grdBusqueda.TextMatrix(grdBusqueda.Row, 6)) = "APLICADA" Then
            cmdDesglose.Enabled = True
        End If
            
        fraCliente.Enabled = False
        fraMotivoNota.Enabled = False
        
        If optPaciente.Value = True Then fraTipoPaciente.Enabled = False
    End If
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdBusqueda_DblClick"))
    Unload Me
End Sub

Private Sub grdBusqueda_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
   If KeyAscii = vbKeyReturn Then
      grdBusqueda_DblClick
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdBusqueda_KeyPress"))
    Unload Me
End Sub

Private Sub grdFactura_DblClick()
On Error GoTo NotificaError

    Dim vlintContador As Integer
    Dim vblnBanderaEncontrado As Boolean

    vblnBanderaEncontrado = False
    
    If vlblnConsulta Then Exit Sub

    If Trim(grdFactura.TextMatrix(1, 1)) <> "" Then
        vlblnElementoSeleccionado = False
        
        sstElementos.Tab = 4
        
        If optNotaCredito.Value Then
            For vlintContador = 0 To lstConceptosFacturacion.ListCount - 1
                If grdFactura.TextMatrix(grdFactura.Row, 1) = lstConceptosFacturacion.List(vlintContador) Then
                    lstConceptosFacturacion.ListIndex = vlintContador
                    vblnBanderaEncontrado = True
                    Exit For
                End If
            Next vlintContador
            
            If Not vblnBanderaEncontrado Then
                lstConceptosFacturacion.Clear
                lstConceptosFacturacion.AddItem grdFactura.TextMatrix(grdFactura.Row, 1)
                lstConceptosFacturacion.ItemData(lstConceptosFacturacion.newIndex) = Val(grdFactura.TextMatrix(grdFactura.Row, 6))
                lstConceptosFacturacion.ListIndex = lstConceptosFacturacion.ListCount - 1
                        
                pSeleccionaElemento
                If fblnExisteConcepto() Then
                    pLimpiarTodosList
                    grdFactura.SetFocus
                    Exit Sub
                End If
            End If
        Else
            pCargaArticulos False
            lstConceptosFacturacion.ListIndex = fintLocalizaLst(lstConceptosFacturacion, grdFactura.TextMatrix(grdFactura.Row, 6))
        End If

        vldblImporteConcepto = CDbl(Format(grdFactura.TextMatrix(grdFactura.Row, 2), vlstrFormato))
        vldblDescuentoConcepto = CDbl(Format(grdFactura.TextMatrix(grdFactura.Row, 3), vlstrFormato))
        vldblIvaConcepto = CDbl(grdFactura.TextMatrix(grdFactura.Row, 12))
        vldblPorcentajeIVA = CDbl(Format(grdFactura.TextMatrix(grdFactura.Row, 10), vlstrFormato))
        vldblTmpDescuentoEspecial = CDbl(Format(grdFactura.TextMatrix(grdFactura.Row, 14), vlstrFormato))
        
        If optCliente.Value And optNotaCredito.Value And OptMotivoNota(0).Value Then
            cmdIncluir.Enabled = False
        Else
            cmdIncluir.Enabled = True
        End If
        
        txtCantidad.SetFocus
        
        If OptMotivoNota(1).Value = True Then cmdIncluir_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdFactura_DblClick"))
    Unload Me
End Sub

Private Sub grdFactura_GotFocus()
On Error GoTo NotificaError
    
    If vlblnConsulta Then Exit Sub
    
    cmdIncluir.Enabled = False
    
    vldblImporteConcepto = 0
    vldblDescuentoConcepto = 0
    vldblIvaConcepto = 0
    vldblPorcentajeIVA = 0
    
    If OptMotivoNota(0).Value = True Then txtCantidad.Text = ""
    txtDescuento.Text = ""
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdFactura_GotFocus"))
    Unload Me
End Sub

Private Sub grdFactura_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        grdFactura_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdFactura_KeyDown"))
    Unload Me
End Sub

Private Sub grdNotas_DblClick()
On Error GoTo NotificaError
       
    If vlblnConsulta Then Exit Sub
    
    Dim vldblCantidad As Double
    Dim vldblDescuento As Double
    Dim vldblIVA As Double
    Dim dblTotalNota As Double
    
    If Trim(grdNotas.TextMatrix(1, 1)) <> "" Or Trim(grdNotas.TextMatrix(1, 2)) <> "" Then
        vldblCantidad = Val(Format(grdNotas.TextMatrix(grdNotas.Row, vlintColCantidad), vlstrFormato))
        vldblDescuento = Val(Format(grdNotas.TextMatrix(grdNotas.Row, vlintColDescuento), "##########.00000"))
        vldblIVA = Val(Format(grdNotas.TextMatrix(grdNotas.Row, vlintColIVA), "##########.00000"))
        
        pAcumulaFactura grdNotas.TextMatrix(grdNotas.Row, vlintColFactura), vldblCantidad, vldblDescuento, vldblIVA, -1, grdNotas.TextMatrix(grdNotas.Row, vlintColTipoCambio)
                
        If grdNotas.Rows - 1 = 1 Then
            txtSubtotal.Text = ""
            txtDescuentoTot.Text = ""
            txtIVA.Text = ""
            txtTotal.Text = ""
            InicializaVariables
            
            pLimpiaGridNota
            pConfiguraGridNota IIf(sstFacturasCreditos.Tab = 0, "FA", "MA")
            
            If optCliente.Value = True Then
                grdNotas.ColWidth(vlintColFactura) = 1100
            Else
                grdNotas.ColWidth(vlintColFactura) = 0
            End If
            
            pLimpiarTodosList
            If blnPestaniaCR Then
                sstFacturasCreditos.TabEnabled(1) = blnPestaniaCR
            End If
            If blnPestaniaFac Then
                sstFacturasCreditos.TabEnabled(0) = blnPestaniaFac
            End If
            
            cmdAddenda.Enabled = False
        Else
            txtSubtotal.Text = FormatCurrency(Val(Format(txtSubtotal.Text, vlstrFormato)) - vldblCantidad)
            txtDescuentoTot.Text = FormatCurrency(Val(Format(txtDescuentoTot.Text, "#########.000")) - vldblDescuento)
            txtIVA.Text = FormatCurrency(Val(Format(txtIVA.Text, "#########.000")) - vldblIVA)
            txtTotal.Text = FormatCurrency(Val(Format(txtSubtotal.Text, vlstrFormato)) - Val(Format(txtDescuentoTot.Text, vlstrFormato)) + Val(Format(txtIVA.Text, vlstrFormato)), 2)
            vldblSubTotalTemporal = CDbl(txtSubtotal.Text)
            vldblDescuentoTemporal = CDbl(txtDescuentoTot.Text)
            vldblIvaTemporal = CDbl(txtIVA.Text)
            vldTotal = CDbl(txtTotal.Text)
                            
            pBorrarRegMshFGrd grdNotas, grdNotas.Row
            
            If OptMotivoNota(1).Value = True Then
                If optCliente.Value = True Then
                    If chkPorcentaje.Value = 0 Or OptMotivoNota(0).Value = True Then
                        dblTotalNota = Val(Format(txtCantidad, "##########.00"))
                    Else
                        dblTotalNota = (Val(Format(txtCantidad, "##########.00")) / 100) * fTotalFactura(grdFactura, Trim(cboFactura.List(cboFactura.ListIndex)))
                    End If
                    If sstFacturasCreditos.Tab <> 1 Then
                        pProrratearNota Val(Format(dblTotalNota, "##########.00"))
                    End If
                Else
                    If chkFacturasPaciente.Value = 1 Then
                        If chkPorcentajePaciente.Value = 0 Then
                            dblTotalNota = Val(Format(txtCantidadCargo, "##########.00"))
                        Else
                            dblTotalNota = (Val(Format(txtCantidadCargo, "##########.00")) / 100) * fTotalFactura(grdCargos, Trim(cboFacturasPaciente.List(cboFacturasPaciente.ListIndex)))
                        End If
                    
                        pProrratearNota Val(Format(dblTotalNota, "##########.00"))
                    End If
                End If
            End If
            
            cmdIncluir.Enabled = False

            Call pProcesoAddenda(False)
        End If
        
        If optPaciente.Value = True Then pColorearRenglon (&O0)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdNotas_DblClick"))
    Unload Me
End Sub

Private Sub mskFechaFinal_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If KeyAscii = vbKeyReturn Then
        cmdCargarDatos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinal_KeyPress"))
    Unload Me
End Sub

Private Sub mskFechaInicial_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If KeyAscii = vbKeyReturn Then
        pEnfocaMkTexto mskFechaFinal
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicial_KeyPress"))
    Unload Me
End Sub

Private Sub optNotaCargo_Click()
On Error GoTo NotificaError

    If vlblnConsulta Then Exit Sub
    
    pLimpiaTodo
    pHabilita False, False, True, False, False, False, False, False
    
    fraMotivoNota.Enabled = False
    OptMotivoNota(0).Value = False
    OptMotivoNota(1).Value = False
    
    lbDescuento.Visible = True
    txtDescuento.Visible = True
    txtDescuento.Enabled = True
    
    lbDescuentoCR.Visible = True
    txtDescuentoCR.Visible = True
    
    cmdAddenda.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optNotaCargo_Click"))
    Unload Me
End Sub

Private Sub optNotaCredito_Click()
On Error GoTo NotificaError

    If vlblnConsulta Then Exit Sub
    
    pLimpiaTodo
    pHabilita False, False, True, False, False, False, False, False
    fraMotivoNota.Enabled = True
    OptMotivoNota(1).Value = True

    cmdAddenda.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optNotaCredito_Click"))
    Unload Me
End Sub

Private Sub fstrFolioDocumento(vlintAumentaFolio As Integer)
    'Regresa el folio para nota de cargo no crédito

    Dim vllngFoliosRestantes As Long
    Dim vlstrFolioDoc As String
    Dim vlstrSerieDoc As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    
    vllngFoliosRestantes = 1
    vlstrFolioDoc = ""
    vlstrSerieDoc = ""
    strNumeroAprobacion = ""
    strAnoAprobacion = ""
       
    pCargaArreglo alstrParametrosSalida, vllngFoliosRestantes & "|" & ADODB.adBSTR & "|" & vlstrFolioDoc & "|" & ADODB.adBSTR & "|" & vlstrSerieDoc & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
    
    If vgintFolioUnico = 0 Then
        frsEjecuta_SP IIf(optNotaCargo.Value, "NA", "NC") & "|" & IIf(vlblnSerieUnica, -1, vgintNumeroDepartamento) & "|" & Str(vlintAumentaFolio), "sp_gnFolios", , , alstrParametrosSalida
    Else
        frsEjecuta_SP "CC" & "|" & IIf(vlblnSerieUnica, -1, vgintNumeroDepartamento) & "|" & Str(vlintAumentaFolio), "sp_gnFolios", , , alstrParametrosSalida
    End If
    
    pObtieneValores alstrParametrosSalida, vllngFoliosRestantes, vlstrFolioDoc, vlstrSerieDoc, strNumeroAprobacion, strAnoAprobacion
    
    vlstrSerieDoc = Trim(vlstrSerieDoc)
    If vllngFoliosRestantes > 0 Then
        MsgBox "Faltan " & Trim(Str(vllngFoliosRestantes)) + " notas y será necesario aumentar folios!", vbOKOnly + vbExclamation, "Mensaje"
    End If

    vlstrFolio = vlstrFolioDoc
    vlstrSerie = vlstrSerieDoc
    vlstrAnoAprobacion = strAnoAprobacion
    vlstrNumeroAprobacion = strNumeroAprobacion
    vlstrFolioDocumento = Trim(vlstrSerieDoc) & Trim(vlstrFolioDoc)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": fstrFolioDocumento"))
    Unload Me
End Sub

Private Sub pLimpiaDetalleFactura()
    fraDetalleFactura.Visible = True
    fraDetalleFactura.Enabled = False
    
    cboFactura.Clear
    lblFechaFactura.Caption = ""
    txtCuenta.Text = ""
    txtNombrePaciente.Text = ""
    
    pLimpiaGridFactura
    pConfiguraGridFactura
End Sub

Private Sub pLimpiaConcepto()
    vlstrCualLista = ""
    
    fraConcepto.Visible = True
    fraConcepto.Enabled = False
    
    txtCantidad.Text = ""
    txtDescuento.Text = ""
    
    cmdIncluir.Enabled = False
    
    txtSeleArticulo = ""
    chkMedicamentos.Value = 1
    optDescripcion.Value = True
    
    pLimpiarTodosList
    sstElementos.Tab = 0
    cmdCargar.Visible = False
    
    vlblnElementoSeleccionado = False
End Sub

Private Sub pLimpiaNota()
    txtComentario.Text = ""
    
    If vlblnConsulta = True Then
        freDetalleNota.Enabled = True
        lblUsoCFDI.Enabled = True
        lblMetodoPago.Enabled = True
        lblFormaPago.Enabled = True
        
        cboUsoCFDI.Enabled = True
        cboMetodoPago.Enabled = True
        cboFormaPago.Enabled = True

    Else
        freDetalleNota.Enabled = False
        lblUsoCFDI.Enabled = False
        lblMetodoPago.Enabled = False
        lblFormaPago.Enabled = False
        
        cboUsoCFDI.Enabled = False
        cboMetodoPago.Enabled = False
        cboFormaPago.Enabled = False

    End If

    pLimpiaGridNota
    pConfiguraGridNota "FA"
End Sub

Private Sub pLimpiaTodo()
On Error GoTo NotificaError
    
    vlblnConsulta = False
    
    lngIDnota = 0
    
    fraEncabezadoNota.Enabled = True
    
    If fraTipoPaciente.Visible = False Then
        mskFecha.Mask = ""
        mskFecha.Text = fdtmServerFecha
        mskFecha.Mask = "##/##/####"
    End If
    
    lblEstatus.ForeColor = vllngColorActivas
    lblEstatus.Caption = "NUEVA"
    
    fraDatosCliente.Enabled = False
    
    txtCveCliente.Text = ""
    txtCliente.Text = ""
    txtDomicilio.Text = ""
    txtRFC.Text = ""
    chkFacturasPagadas.Value = 0
    
    pLimpiaDetalleFactura
    
    pLimpiaCreditosDirectos
    
    pLimpiaConcepto
    
    pLimpiaNota
    
    'Tab de búsqueda
    mskFechaInicial.Mask = ""
    mskFechaInicial.Text = fdtmServerFecha
    mskFechaInicial.Mask = "##/##/####"
    
    mskFechaFinal.Mask = ""
    mskFechaFinal.Text = fdtmServerFecha
    mskFechaFinal.Mask = "##/##/####"
    
    txtNumCliente.Text = ""
    txtNombreCliente.Text = ""
    vlstrtipofactura = ""

    pConfiguraBusqueda
    
    cboUsoCFDI.ListIndex = -1
    cboFormaPago.ListIndex = -1
    cboMetodoPago.ListIndex = -1
    vlblnSeleccionoOtroConcepto = False
    vlstrCodigo = ""
    vlstrRegimen = ""
    vlRazonSocialComprobante = ""
    vlRFCComprobante = ""
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaTodo"))
    Unload Me
End Sub

Private Sub pLimpiaCreditosDirectos()
    cboCreditosDirectos.Clear
    grdCreditoDirecto.Clear
    pConfiguraGridCreditos
    lstConceptosFact.Clear
    sstFacturasCreditos.Enabled = False
    cmdIncluirCR.Enabled = False
    sstFacturasCreditos.Tab = 0
End Sub

Private Sub optNotaCredito_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
   
   If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFecha
   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optNotaCredito_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidad_GotFocus()
    pSelTextBox txtCantidad
    'cmdIncluir.Enabled = False
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlstrCualLista As String

    If KeyCode = vbKeyReturn Then
        If optCliente.Value And optNotaCredito.Value And OptMotivoNota(0).Value Then
            Select Case sstElementos.Tab
                Case 0
                    Set lstListaSeleccionada = lstArticulos
                    vlstrCualLista = "AR"
                Case 1
                    Set lstListaSeleccionada = lstEstudios
                    vlstrCualLista = "ES"
                Case 2
                    Set lstListaSeleccionada = lstExamenes
                    If lstExamenes.ItemData(lstExamenes.ListIndex) < 0 Then
                        vlstrCualLista = "GE"
                    Else
                        vlstrCualLista = "EX"
                    End If
                Case 3
                    Set lstListaSeleccionada = lstOtrosConceptos
                    vlstrCualLista = "OC"
                Case 4
                    Set lstListaSeleccionada = lstConceptosFacturacion
                    vlstrCualLista = "CF"
            End Select
                            
            If vlstrCualLista = "CF" Then
                If vldblImporteConcepto <> 0 And cboFactura.ListIndex <> -1 Then
                    If Val(Format(txtCantidad.Text, vlstrFormato)) <= vldblImporteConcepto Then
                        vlstrSentencia = "SELECT NVL(SUM(mnyCantidad), 0) SumaNotas " & _
                                         "FROM ccNotaDetalle, ccNota " & _
                                         "WHERE ccNota.intconsecutivo = ccNotaDetalle.intconsecutivo " & _
                                         " AND ccNota.CHRESTATUS <> 'C' " & _
                                         " AND ccNotaDetalle.intConcepto = " & Val(grdFactura.TextMatrix(grdFactura.Row, 6)) & _
                                         " AND trim(chrFolioFactura) = '" & Trim(cboFactura.Text) & "'" & _
                                         " AND ccNota.CHRTIPO = 'CR'"
                        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                        If rs.RecordCount <> 0 Then
                            If rs!SumaNotas + Val(Format(txtCantidad.Text, vlstrFormato)) > vldblImporteConcepto Then
                                'La cantidad excede el importe del concepto, ya que tiene notas de crédito por
                                MsgBox SIHOMsg(1287) & FormatCurrency(rs!SumaNotas, 2) & ".", vbOKOnly + vbInformation, "Mensaje"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If Val(Format(txtCantidad.Text, vlstrFormato)) = vldblImporteConcepto Then
                    txtDescuento.Text = FormatCurrency(vldblDescuentoConcepto, 2)
                Else
                    If optCliente.Value And optNotaCredito.Value And OptMotivoNota(0).Value Then
                        If vldblDescuentoConcepto <> 0 Then
                            txtDescuento.Text = FormatCurrency((vldblDescuentoConcepto * Val(Format(txtCantidad.Text, vlstrFormato))) / vldblImporteConcepto, 2)
                        Else
                            txtDescuento.Text = FormatCurrency(vldblDescuentoConcepto, 2)
                        End If
                        txtDescuento.Enabled = False
                    End If
                End If
                cmdIncluir.Enabled = True
            Else
                txtDescuento.Text = ""
                txtDescuento.Enabled = False
            End If
        Else
            If Val(Format(txtCantidad.Text, vlstrFormato)) = vldblImporteConcepto Then
                txtDescuento.Text = FormatCurrency(vldblDescuentoConcepto, 2)
            End If
        End If
        
        If txtDescuento.Visible = True And txtDescuento.Enabled = True Then
            txtDescuento.SetFocus
        Else
            If cmdIncluir.Enabled = True Then cmdIncluir.SetFocus
        End If
        If OptMotivoNota(1).Value = True Then
            grdFactura.SetFocus
        End If
    End If
      
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCantidad, KeyAscii, 2) Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtCveCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    Dim rsTemp As New ADODB.Recordset
    Dim vllngNumCliente As Long
    Dim intClaveCliente As Long
    Dim rsTipoCliente As ADODB.Recordset
    Dim rsCargos As New ADODB.Recordset
    Dim vgstrParametrosSP As String
    Dim rsPaciente As ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlblnClienteValido As Boolean
    Dim rsDatosFISC As New ADODB.Recordset
        
    If KeyCode = vbKeyReturn Then
        '***** Se está aplicando una nota a un Cliente *****'
        If optCliente.Value = True Then
            If Trim(txtCveCliente.Text) = "" Then
               vllngNumCliente = flngNumCliente(False, 1)
               If vllngNumCliente <> 0 Then
                  txtCveCliente.Text = vllngNumCliente
               End If
            End If
            
            If Trim(txtCveCliente.Text) <> "" Then
                vlblnClienteValido = pDatosCliente(txtCveCliente.Text)
                If vlblnClienteValido Then
                
                
                    If rsDatosCliente.RecordCount > 0 Then
                        vlstrSentencia = "select trim(ccnota.chrfolionota) folio " & _
                                                ", gncomprobantefiscaldigital.vchformapago " & _
                                                ", gncomprobantefiscaldigital.vchmetodopago " & _
                                                ", gncomprobantefiscaldigital.vchusocfdi " & _
                                            "From ccnota " & _
                                                "inner join gncomprobantefiscaldigital on trim(ccnota.chrfolionota) = trim(gncomprobantefiscaldigital.vchseriecomprobante) || trim(gncomprobantefiscaldigital.vchfoliocomprobante) " & _
                                                    "and gncomprobantefiscaldigital.chrtipocomprobante in ('CR','CA') " & _
                                            "Where ccnota.intcliente = " & txtCveCliente.Text & " " & _
                                            "order by gncomprobantefiscaldigital.intidcomprobante desc"
                        Set rs = frsRegresaRs(vlstrSentencia)
                        If rs.RecordCount <> 0 Then
                            If rs!VCHFORMAPAGO <> "" Then
                                cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, Trim(rs!VCHFORMAPAGO) & " - ")
                            Else
                                cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, "99 - ")
                            End If
                            
                            If rs!vchMetodoPago <> "" Then
                                cboMetodoPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboMetodoPago, Trim(rs!vchMetodoPago) & " - ")
                            Else
                                cboMetodoPago.ListIndex = 0
                            End If
                            
    '                        If rs!vchusocfdi <> "" Then
    '                            cboUsoCFDI.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboUsoCFDI, Trim(rs!vchusocfdi) & " - ")
    '                        Else
    '                            cboUsoCFDI.ListIndex = -1
    '                        End If
                        Else
                            cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, "99 - ")
                            cboMetodoPago.ListIndex = 0
                        End If
                    
                        'MsgBox rsDatosCliente!chrTipoCliente
                        If rsDatosCliente!chrTipoCliente = "CO" Then
                            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", rsDatosCliente!intNumReferencia, "EM", IIf(optNotaCredito.Value, 2, 1)))
                        End If
                        If rsDatosCliente!chrTipoCliente = "ME" Then
                            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", 0, "ME", IIf(optNotaCredito.Value, 2, 1)))
                        End If
                        If rsDatosCliente!chrTipoCliente = "EM" Then
                            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", 0, "EP", IIf(optNotaCredito.Value, 2, 1)))
                        End If
                        If rsDatosCliente!chrTipoCliente = "PE" Then
                            Set rsPaciente = frsRegresaRs("select intCveTipoPaciente from EXPacienteIngreso where chrTipoIngreso = 'E' and intNumCuenta = " & rsDatosCliente!intNumReferencia)
                            If Not rsPaciente.EOF Then
                                cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", rsPaciente!intCveTipoPaciente, "TP", IIf(optNotaCredito.Value, 2, 1)))
                            End If
                            rsPaciente.Close
                        End If
                        If rsDatosCliente!chrTipoCliente = "PI" Then
                            Set rsPaciente = frsRegresaRs("select intCveTipoPaciente from EXPacienteIngreso where chrTipoIngreso = 'I' and intNumCuenta = " & rsDatosCliente!intNumReferencia)
                            If Not rsPaciente.EOF Then
                                cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", rsPaciente!intCveTipoPaciente, "TP", IIf(optNotaCredito.Value, 2, 1)))
                            End If
                            rsPaciente.Close
                        End If
                        
                        sstFacturasCreditos.Enabled = True
                        sstFacturasCreditos.TabEnabled(0) = True
                        sstFacturasCreditos.TabEnabled(1) = True
                        If rsDatosCliente!smicvedepartamento = vgintNumeroDepartamento Then
                            pAsignarTxtNombre IIf(IsNull(rsDatosCliente!RazonSocial), " ", rsDatosCliente!RazonSocial), txtCliente
                            pAsignarTxtRFC IIf(IsNull(rsDatosCliente!RFCCliente), "", rsDatosCliente!RFCCliente), txtRFC
                            pAsignarTxtDomicilio IIf(IsNull(rsDatosCliente!DomicilioCliente), "", rsDatosCliente!DomicilioCliente), txtDomicilio
                            
                            pFacturas
                            
                            pCreditosDirectos
                                                                                                                           
                            optPaciente.Enabled = False
                            
                            
                            If cboFactura.ListCount <> 0 Then
                                pHabilita False, False, False, False, False, True, False, False
                                fraDetalleFactura.Enabled = True
                                freDetalleNota.Enabled = True
                                lblUsoCFDI.Enabled = True
                                lblMetodoPago.Enabled = True
                                lblFormaPago.Enabled = True
                                
                                cboUsoCFDI.Enabled = True
                                cboMetodoPago.Enabled = True
                                cboFormaPago.Enabled = True
    
                                txtComentario.Enabled = True
                                fraDatosCliente.Enabled = False
                                fraConcepto.Enabled = True
                                blnPestaniaFac = True
                            Else
                                sstFacturasCreditos.TabEnabled(0) = False
                            End If
                                                    
                            If cboCreditosDirectos.ListCount <> 0 Then
                                pHabilita False, False, False, False, False, True, False, False
                                blnPestaniaCR = True
                                freDetalleNota.Enabled = True
                                lblUsoCFDI.Enabled = True
                                lblMetodoPago.Enabled = True
                                lblFormaPago.Enabled = True
                                
                                cboUsoCFDI.Enabled = True
                                cboMetodoPago.Enabled = True
                                cboFormaPago.Enabled = True
    
                                txtComentario.Enabled = True
                                If sstFacturasCreditos.TabEnabled(0) Then
                                    sstFacturasCreditos.Tab = 0
                                    If optCliente.Value Then
                                        If fblnCanFocus(cboFormaPago) Then cboFormaPago.SetFocus
                                    Else
                                        If fblnCanFocus(cboFactura) Then cboFactura.SetFocus
                                    End If
                                Else
                                    sstFacturasCreditos.Tab = 1
                                    If optCliente.Value Then
                                        If fblnCanFocus(cboFormaPago) Then cboFormaPago.SetFocus
                                    Else
                                        If fblnCanFocus(cboCreditosDirectos) Then cboCreditosDirectos.SetFocus
                                    End If
                                End If
                            Else
                                sstFacturasCreditos.TabEnabled(1) = False
                                If optCliente.Value Then
                                    If fblnCanFocus(cboFormaPago) Then cboFormaPago.SetFocus
                                Else
                                    If cboFactura.Enabled = True Then cboFactura.SetFocus
                                End If
                            End If
                                                    
                            If sstFacturasCreditos.TabEnabled(1) = False And sstFacturasCreditos.TabEnabled(0) = False Then
                                'No se entraron facturas pendientes de cobro.
                                MsgBox SIHOMsg(713), vbOKOnly + vbInformation, "Mensaje"
                                txtCveCliente.Text = ""
                                txtCliente.Text = ""
                                txtRFC.Text = ""
                                txtDomicilio.Text = ""
                                pLimpiaTodo
                                pEnfocaTextBox txtCveCliente
                                optNotaCredito.SetFocus
                                optPaciente.Enabled = True
                            End If
                            
                            
                            
                        Else
                            pAsignarTxtNombre IIf(IsNull(rsDatosCliente!NombreCliente), " ", rsDatosCliente!NombreCliente), txtCliente
                            'El cliente seleccionado no pertenece a este departamento.
                            MsgBox SIHOMsg(646), vbOKOnly + vbInformation, "Mensaje"
                            txtCveCliente.Text = ""
                            txtCliente.Text = ""
                            txtRFC.Text = ""
                            txtDomicilio.Text = ""
                            pLimpiaTodo
                            pEnfocaTextBox txtCveCliente
                            optNotaCredito.SetFocus
                            
                            optPaciente.Enabled = True
                        End If
                    Else
                        '¡La información no existe!
                        MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                        txtCveCliente.Text = ""
                        txtCliente.Text = ""
                        txtRFC.Text = ""
                        txtDomicilio.Text = ""
                        pEnfocaTextBox txtCveCliente
                    End If
                
                End If
                
            End If
        Else
            '***** Se está aplicando una nota a un Paciente *****'
            If RTrim(txtCveCliente.Text) = "" Then
                With FrmBusquedaPacientes
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = True
                    
                    If OptTipoPaciente(1).Value Then 'Externos
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,1700,4100"
                        .vgstrTipoPaciente = "E"
                        .Caption = .Caption & " Externos"
                    Else 'Internos
                        .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                        .vgstrTamanoCampo = "800,3400,990,990,4100"
                        .vgstrTipoPaciente = "I"
                        .Caption = .Caption & " Internos"
                    End If
            
                    txtCveCliente.Text = .flngRegresaPaciente()
                End With
            End If  'End If Trim(txtCveCliente.Text) <> ""
                
            If Trim(txtCveCliente.Text) <> "" Then
                vlstrSentencia = "select trim(ccnota.chrfolionota) folio " & _
                                        ", gncomprobantefiscaldigital.vchformapago " & _
                                        ", gncomprobantefiscaldigital.vchmetodopago " & _
                                        ", gncomprobantefiscaldigital.vchusocfdi " & _
                                    "From ccnota " & _
                                        "inner join gncomprobantefiscaldigital on trim(ccnota.chrfolionota) = trim(gncomprobantefiscaldigital.vchseriecomprobante) || trim(gncomprobantefiscaldigital.vchfoliocomprobante) " & _
                                            "and gncomprobantefiscaldigital.chrtipocomprobante in ('CR','CA') " & _
                                    "Where ccnota.chrnotadirigida = 'P' and ccnota.chrtipopaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "' and intmovpaciente = " & txtCveCliente.Text & " " & _
                                    "order by gncomprobantefiscaldigital.intidcomprobante desc"
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount <> 0 Then
                    If rs!VCHFORMAPAGO <> "" Then
                        cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, Trim(rs!VCHFORMAPAGO) & " - ")
                    Else
                        cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, "99 - ")
                    End If
                    
                    If rs!vchMetodoPago <> "" Then
                        cboMetodoPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboMetodoPago, Trim(rs!vchMetodoPago) & " - ")
                    Else
                        cboMetodoPago.ListIndex = 0
                    End If
                    
'                        If rs!vchusocfdi <> "" Then
'                            cboUsoCFDI.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboUsoCFDI, Trim(rs!vchusocfdi) & " - ")
'                        Else
'                            cboUsoCFDI.ListIndex = -1
'                        End If
                Else
                    cboFormaPago.ListIndex = fintLocalizaCritCboFormaMetodoUso(cboFormaPago, "99 - ")
                    cboMetodoPago.ListIndex = 0
                End If
            
                vgstrParametrosSP = Trim(txtCveCliente.Text) & "|" & Str(vgintClaveEmpresaContable)
                
                If OptTipoPaciente(0).Value = True Then
                    Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
                Else
                    Set rsPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
                End If
                
                If rsPaciente.RecordCount > 0 Then
                    sstFacturasCreditos.Enabled = True
                    sstFacturasCreditos.TabEnabled(0) = True
                    sstFacturasCreditos.TabEnabled(1) = True
                    
                    cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", rsPaciente!cveTipoPaciente, "TP", IIf(optNotaCredito.Value, 2, 1)))
                    
                    pAsignarTxtNombre IIf(IsNull(rsPaciente!NombreFiscal), " ", rsPaciente!NombreFiscal), txtCliente
                    pAsignarTxtRFC IIf(IsNull(rsPaciente!RFC), "", rsPaciente!RFC), txtRFC
                    pAsignarTxtDomicilio IIf(IsNull(Trim(rsPaciente!Direccion)), "", Trim(rsPaciente!Direccion)), txtDomicilio
                    
                    If vgstrVersionCFDI = "4.0" Then
                        Set rsDatosFISC = frsRegresaRs("SELECT chrnombre, CHRRFC, VCHCODIGOPOSTAL, VCHREGIMENFISCALRECEPTOR FROM PVFACTURA WHERE CHRTIPOFACTURA = 'P' AND INTMOVPACIENTE = " & Trim(txtCveCliente.Text) & " order by intconsecutivo desc")
                        If rsDatosFISC.RecordCount > 0 Then
                            vlstrRegimen = IIf(IsNull(rsDatosFISC!VCHREGIMENFISCALRECEPTOR), "", rsDatosFISC!VCHREGIMENFISCALRECEPTOR)
                            vlstrCodigo = IIf(IsNull(rsDatosFISC!VCHCODIGOPOSTAL), "", rsDatosFISC!VCHCODIGOPOSTAL)
                            vlRazonSocialComprobante = IIf(IsNull(rsDatosFISC!CHRNOMBRE), "", rsDatosFISC!CHRNOMBRE)
                            vlRFCComprobante = IIf(IsNull(rsDatosFISC!chrRFC), "", rsDatosFISC!chrRFC)
                            
                                Set rsDatosFISC = frsRegresaRs("SELECT VCHREGIMENFISCAL, VCHCODIGOPOSTAL, chrrfc, chrnombre FROM PVDATOSFISCALES WHERE intnumcuenta = '" & Trim(txtCveCliente.Text) & "' order BY intid desc")
                                If rsDatosFISC.RecordCount > 0 Then
                                    If vlstrRegimen = "" Then
                                        vlstrRegimen = IIf(IsNull(rsDatosFISC!vchregimenfiscal), "", rsDatosFISC!vchregimenfiscal)
                                    End If
                                    
                                    If vlstrCodigo = "" Then
                                        vlstrCodigo = IIf(IsNull(rsDatosFISC!VCHCODIGOPOSTAL), "", rsDatosFISC!VCHCODIGOPOSTAL)
                                    End If
                                    
                                    If vlRazonSocialComprobante = "" Then
                                        vlRazonSocialComprobante = IIf(IsNull(rsDatosFISC!CHRNOMBRE), "", rsDatosFISC!CHRNOMBRE)
                                    End If
                                    
                                    If vlRFCComprobante = "" Then
                                        vlRFCComprobante = IIf(IsNull(rsDatosFISC!chrRFC), "", rsDatosFISC!chrRFC)
                                    End If
                                End If
                        Else
                            vlRazonSocialComprobante = ""
                            vlstrRegimen = ""
                            vlstrCodigo = ""
                            Set rsDatosFISC = frsRegresaRs("SELECT VCHREGIMENFISCAL, VCHCODIGOPOSTAL, chrrfc, chrnombre FROM PVDATOSFISCALES WHERE intnumcuenta = '" & Trim(txtCveCliente.Text) & "' order BY intid desc")
                            If rsDatosFISC.RecordCount > 0 Then
                                vlstrRegimen = IIf(IsNull(rsDatosFISC!vchregimenfiscal), "", rsDatosFISC!vchregimenfiscal)
                                vlstrCodigo = IIf(IsNull(rsDatosFISC!VCHCODIGOPOSTAL), "", rsDatosFISC!VCHCODIGOPOSTAL)
                                vlRazonSocialComprobante = IIf(IsNull(rsDatosFISC!CHRNOMBRE), "", rsDatosFISC!CHRNOMBRE)
                                vlRFCComprobante = IIf(IsNull(rsDatosFISC!chrRFC), "", rsDatosFISC!chrRFC)
                            End If
                            
                            Set rsDatosFISC = frsRegresaRs("SELECT trim(vchnombre) || ' ' || trim(vchapellidopaterno) || ' ' || trim(vchapellidomaterno) chrnombre, EXPACIENTE.VCHRFC, GNDOMICILIO.VCHCODIGOPOSTAL, PVDATOSFISCALES.VCHREGIMENFISCAL FROM EXPACIENTEINGRESO LEFT JOIN EXPACIENTE ON EXPACIENTEINGRESO.INTNUMPACIENTE = EXPACIENTE.INTNUMPACIENTE LEFT JOIN EXPACIENTEDOMICILIO ON EXPACIENTEDOMICILIO.INTNUMPACIENTE = EXPACIENTE.INTNUMPACIENTE LEFT JOIN GNDOMICILIO ON EXPACIENTEDOMICILIO.INTCVEDOMICILIO = GNDOMICILIO.INTCVEDOMICILIO LEFT JOIN PVDATOSFISCALES ON TRIM(EXPACIENTE.VCHRFC) = TRIM(PVDATOSFISCALES.CHRRFC) AND EXPACIENTEINGRESO.INTNUMCUENTA = PVDATOSFISCALES.INTNUMCUENTA OR 1 = 1 WHERE EXPACIENTEINGRESO.INTNUMCUENTA = " & Trim(txtCveCliente.Text) & " AND ROWNUM = 1 ORDER BY VCHREGIMENFISCAL DESC")
                            If rsDatosFISC.RecordCount > 0 Then
                                If vlstrRegimen = "" Then
                                    vlstrRegimen = IIf(IsNull(rsDatosFISC!vchregimenfiscal), "", rsDatosFISC!vchregimenfiscal)
                                End If
                                
                                If vlstrCodigo = "" Then
                                    vlstrCodigo = IIf(IsNull(rsDatosFISC!VCHCODIGOPOSTAL), "", rsDatosFISC!VCHCODIGOPOSTAL)
                                End If
                                
                                If vlRazonSocialComprobante = "" Then
                                    vlRazonSocialComprobante = IIf(IsNull(rsDatosFISC!CHRNOMBRE), "", rsDatosFISC!CHRNOMBRE)
                                End If
                                
                                If vlRFCComprobante = "" Then
                                    vlRFCComprobante = IIf(IsNull(rsDatosFISC!vchRFC), "", rsDatosFISC!vchRFC)
                                End If
                            End If
                        End If
                        
                        If vlstrCodigo = "" Then
                            rsDatosFISC.Close
                            MsgBox "No se encuentra configurado el código postal correspondiente al paciente " & Trim(IIf(IsNull(rsPaciente!NombreFiscal), " ", rsPaciente!NombreFiscal)) & ", se puede configurar en la información del paciente.", vbExclamation + vbOKOnly, "Mensaje"
                            txtCveCliente.Text = ""
                            Exit Sub
                        End If
                        If vlstrRegimen = "" Then
                           rsDatosFISC.Close
                            MsgBox "No se encuentra configurado el régimen fiscal correspondiente al paciente " & Trim(IIf(IsNull(rsPaciente!NombreFiscal), " ", rsPaciente!NombreFiscal)) & ", se puede configurar en la información del paciente.", vbExclamation + vbOKOnly, "Mensaje"
                            txtCveCliente.Text = ""
                            pLimpiaTodo
                            Exit Sub
                        End If
                        rsDatosFISC.Close
                    End If
                    
                    If Not IsNull(rsPaciente!FechaIngreso) Then vgdtmFechaIngreso = CDate(rsPaciente!FechaIngreso)
                    
                    optCliente.Enabled = False
                    
                    pHabilita False, False, False, False, False, True, False, False
                    
                    '-- Agregado para caso 7374 --'
                    If OptMotivoNota(0).Value Then
                        chkFacturasPaciente.Value = vbChecked
                        chkFacturasPaciente.Enabled = False
                    Else
                        tmrCargos.Enabled = True
                    End If
                    
                    pHabilitaFrames True, True 'Modificado para caso 7374, para habilitar frame de Motivo de la nota

                    pHabilitaCuadros True
                    
                    If cboFormaPago.Enabled Then
                        cboFormaPago.SetFocus
                    End If
                Else
                    If txtCveCliente.Text <> -1 Then
                        '¡La información no existe!
                        MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                    End If
                    
                    txtCveCliente.Text = ""
                    txtCliente.Text = ""
                    txtRFC.Text = ""
                    txtDomicilio.Text = ""
                    pEnfocaTextBox txtCveCliente
                    
                    optCliente.Enabled = True
                End If ' End rsPaciente.RecordCount > 0
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                txtCveCliente.Text = ""
                txtCliente.Text = ""
                txtRFC.Text = ""
                txtDomicilio.Text = ""
                pEnfocaTextBox txtCveCliente
            End If ' End if Trim(txtCveCliente.Text) <> ""
        End If  'optCliente.Value
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveCliente_KeyDown"))
    Unload Me
End Sub

Public Sub pLlenaCargos(strCuentaPaciente As String)
On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    Dim rsSeleccionaCargos As New ADODB.Recordset
    Dim intLongitud As Integer
    Dim strCompara As String
        
    '-------------------------------------------------------------------
    ' SP para cargar los cargos del paciente
    '-------------------------------------------------------------------
    'grdCargos.Redraw = False
    pLimpiaGridCargos
    pConfiguraGridCargos
    If txtCliente.Text = "" And txtDomicilio.Text = "" And txtRFC.Text = "" Then
        pHabilitaCuadros False
        pLimpiaNota
    Else
        vgstrParametrosSP = strCuentaPaciente & "|" & IIf(OptTipoPaciente(0), "I", "E") & "|" & 0 & "|" & "-1|N|0"
        Set rsSeleccionaCargos = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvselcargospacienteNotas")
        If rsSeleccionaCargos.RecordCount = 0 Then
            FrmBusquedaPacientes.Visible = False
        
            '¡La información no existe!
            If txtCveCliente.Text <> "" Then MsgBox SIHOMsg(288), vbOKOnly + vbInformation, "Mensaje"
                grdCargos.Redraw = True
                chkFacturasPaciente.Value = 1
            Else
                If optPaciente.Value = True And optNotaCredito.Value And OptMotivoNota(1).Value Then
                    If OptTipoPaciente(0).Value Then 'Internos
                        vgstrParametrosSP = txtCveCliente.Text & "|" & Str(vgintClaveEmpresaContable)
                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
                    Else  'Externos
                        vgstrParametrosSP = txtCveCliente.Text & "|" & Str(vgintClaveEmpresaContable)
                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
                    End If
            
                    If rs.RecordCount <> 0 Then
                        If rs!Facturado = 1 Or rs!Facturado = True Then
                            'La cuenta del paciente está completamente facturada.
                            MsgBox SIHOMsg(597), vbOKOnly + vbInformation, "Mensaje"
                            grdCargos.Redraw = True
                            
                            InicializaComponentes
                            If optNotaCargo.Value Then
                               optNotaCargo.SetFocus
                            Else
                               optNotaCredito.SetFocus
                            End If
            
                            Exit Sub
                        End If
                    End If
                End If
            
                With rsSeleccionaCargos
                    Do While Not .EOF
                        If txtBuscaCargo.Text = "" Then
                            If grdCargos.RowData(1) <> -1 Then
                                 grdCargos.Row = grdCargos.Rows - 1
                            End If
                        
                            grdCargos.TextMatrix(grdCargos.Row, 1) = Format(!dtmFechahora, "dd/mmm/yyyy")
                            grdCargos.TextMatrix(grdCargos.Row, 2) = IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo)
                            grdCargos.TextMatrix(grdCargos.Row, 3) = !MNYCantidad
                            grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(!MNYPRECIO, 2)
                            grdCargos.TextMatrix(grdCargos.Row, 5) = FormatCurrency(!MNYDESCUENTO, 2)
'                            grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((!smyIVA / 100 * (!mnyPrecio - !MNYDESCUENTO)) * !mnycantidad, 2)
'                            grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(!mnyPrecio * !mnycantidad - !MNYDESCUENTO + ((!mnyPrecio * !mnycantidad) * !smyIVA / 100), 2)
                            grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((!smyIVA / 100 * (!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO)), 2)
                            grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO + ((!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO) * !smyIVA / 100), 2)
                            grdCargos.TextMatrix(grdCargos.Row, 8) = !NombreDepartamento
                            grdCargos.TextMatrix(grdCargos.Row, 9) = !chrTipoCargo
                            grdCargos.TextMatrix(grdCargos.Row, 10) = !CveConceptoFacturacion
                            grdCargos.TextMatrix(grdCargos.Row, 11) = !intFolioDocumento
                            grdCargos.TextMatrix(grdCargos.Row, 12) = !CveDepartamento
                            
                            If grdCargos.RowData(1) <> -1 Then
                                grdCargos.Rows = grdCargos.Rows + 1
                            End If
                        Else
                            intLongitud = Len(txtBuscaCargo.Text)
                            strCompara = Mid(IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo), 1, intLongitud)
                            If strCompara = txtBuscaCargo.Text Then
                                If grdCargos.RowData(1) <> -1 Then
                                    grdCargos.Row = grdCargos.Rows - 1
                                End If
                            
                                grdCargos.TextMatrix(grdCargos.Row, 1) = Format(!dtmFechahora, "dd/mm/yyyy")
                                grdCargos.TextMatrix(grdCargos.Row, 2) = IIf(IsNull(!DescripcionCargo), "", !DescripcionCargo)
                                grdCargos.TextMatrix(grdCargos.Row, 3) = !MNYCantidad
                                grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(!MNYPRECIO, 2)
                                grdCargos.TextMatrix(grdCargos.Row, 5) = FormatCurrency(!MNYDESCUENTO, 2)
'                                grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((!smyIVA / 100 * (!mnyPrecio - !MNYDESCUENTO)) * !mnycantidad, 2)
'                                grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(!mnyPrecio * !mnycantidad - !MNYDESCUENTO + ((!mnyPrecio * !mnycantidad) * !smyIVA / 100), 2)
                                grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((!smyIVA / 100 * (!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO)), 2)
                                grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO + ((!MNYPRECIO * !MNYCantidad - !MNYDESCUENTO) * !smyIVA / 100), 2)
                                grdCargos.TextMatrix(grdCargos.Row, 8) = !NombreDepartamento
                                grdCargos.TextMatrix(grdCargos.Row, 9) = !chrTipoCargo
                                grdCargos.TextMatrix(grdCargos.Row, 10) = !CveConceptoFacturacion
                                grdCargos.TextMatrix(grdCargos.Row, 11) = !intFolioDocumento
                                grdCargos.TextMatrix(grdCargos.Row, 12) = !CveDepartamento
                                
                                If grdCargos.RowData(1) <> -1 Then
                                    grdCargos.Rows = grdCargos.Rows + 1
                                End If
                            End If
                        End If
                        .MoveNext
                    Loop
                    .Close
                
                    grdCargos.Rows = grdCargos.Row + 1
                End With
            '  grdCargos.Redraw = True
        End If
    End If
 
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCargos"))
    Unload Me
End Sub

Private Sub pCargarOpciones()
On Error GoTo NotificaError
    
    Select Case cgstrModulo
        Case "CC"
            optCliente.Value = True
            
            sstCargos.Visible = False
            sstFacturasCreditos.Visible = True
                
        Case "PV"
            optPaciente.Value = True
            OptTipoPaciente(0).Value = True
            sstFacturasCreditos.TabEnabled(1) = False
                        
            sstFacturasCreditos.Visible = False
            sstCargos.Visible = True
            sstCargos.Caption = ""
            pConfiguraGridCargos
    End Select

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargarOpciones"))
End Sub

Private Sub pProcesoAddenda(blnAgregar As Boolean)
'Para el parámetro blnAgregar (True = Se van a agregan datos en la nota, False = Se van a eliminar datos en la nota)
On Error GoTo NotificaError

    Dim strParametros As String
    Dim rsDatosPacienteAddenda As New ADODB.Recordset
    Dim vlintContFacturas As Integer
    Dim vlstrFolioFactura As String
    Dim rsAddendasCliente As New ADODB.Recordset

    vglngCveAddenda = 0
    vglngCveEmpresaCliente = 0
    vgstrTipoPacienteAddenda = ""
    vglngCuentaPacienteAddenda = 0
    vlstrFolioFactura = ""
    vlintContFacturas = 0
    vgMostrarMsjAddenda = False
    
    If optCliente.Value = True Then 'Aplica únicamente en clientes

        Dim i As Integer
        For i = 1 To grdNotas.Rows - 1
            If vlstrFolioFactura <> Trim(grdNotas.TextMatrix(i, 1)) Then
                vlintContFacturas = vlintContFacturas + 1
            End If
            
            vlstrFolioFactura = Trim(grdNotas.TextMatrix(i, 1))
        Next i
    
        'Si solo hay cargo de una sola factura sigue el proceso
        If vlintContFacturas = 1 Then
            'Revisa la información de la factura para identificar si el paciente de la factura utiliza convenio con addenda
            vglngCveAddenda = 1
            strParametros = "|" & IIf(optNotaCargo.Value, "CA", "CR") & "|" & IIf(sstFacturasCreditos.Tab = 0, IIf(blnAgregar = True, Trim(cboFactura.List(cboFactura.ListIndex)), Trim(grdNotas.TextMatrix(grdNotas.Row, 1))), " ")
            frsEjecuta_SP strParametros & "|" & vgintClaveEmpresaContable, "FN_PVSELADDENDAEMPRESACLIENTE", True, vglngCveAddenda
            
            'Se obtiene la empresa cliente a la cual se le otorga la addenda
            vglngCveEmpresaCliente = 1
            strParametros = "|" & IIf(optNotaCargo.Value, "CA", "CR") & "|" & IIf(sstFacturasCreditos.Tab = 0, IIf(blnAgregar = True, Trim(cboFactura.List(cboFactura.ListIndex)), Trim(grdNotas.TextMatrix(grdNotas.Row, 1))), " ")
            frsEjecuta_SP strParametros & "|" & vgintClaveEmpresaContable, "FN_PVSELEMPRESACLIENTEADDENDA", True, vglngCveEmpresaCliente
            
            'Valida si se tiene licencia para emitir la addenda seleccionada
            vglngCveAddenda = IIf(fblLicenciaAddenda(vglngCveAddenda) = True, vglngCveAddenda, 0)

            If vglngCveAddenda = 0 Then
                cmdAddenda.Enabled = False
            Else
                cmdAddenda.Enabled = True
                
                'Después de habilitar el botón de addenda, se consulta la información del paciente relacionado con la factura
                strParametros = IIf(sstFacturasCreditos.Tab = 0, IIf(blnAgregar = True, Trim(cboFactura.List(cboFactura.ListIndex)), Trim(grdNotas.TextMatrix(grdNotas.Row, 1))), " ")
                Set rsDatosPacienteAddenda = frsEjecuta_SP(strParametros, "SP_CCDATOSPACIENTENOTA")
                
                If rsDatosPacienteAddenda.RecordCount > 0 Then
                    vgstrTipoPacienteAddenda = IIf(IsNull(rsDatosPacienteAddenda!TipoPaciente), "", rsDatosPacienteAddenda!TipoPaciente)
                    vglngCuentaPacienteAddenda = IIf(IsNull(rsDatosPacienteAddenda!cuenta), 0, rsDatosPacienteAddenda!cuenta)
                End If
            End If
            
            vgMostrarMsjAddenda = False
        Else
            'Verifica si la empresa cliente tiene configuradas addendas, para saber si se muestra el mensaje
            Set rsAddendasCliente = frsEjecuta_SP(Trim(txtCveCliente.Text), "SP_PVSELADDENDACLIENTE")
            
            If rsAddendasCliente.RecordCount > 0 Then
                If Val(rsAddendasCliente!cont) > 0 Then
                    vgMostrarMsjAddenda = True
                End If
            End If
            cmdAddenda.Enabled = False
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pProcesoAddenda"))
End Sub

Private Sub pFacturas()
On Error GoTo NotificaError
    
    Dim rsFacturas As New ADODB.Recordset
    Dim strParametros As String
    
    strParametros = txtCveCliente.Text & "|" & IIf(chkFacturasPagadas.Value = 1, "P", "N")
    Set rsFacturas = frsEjecuta_SP(strParametros, "Sp_CcSelFacturaNota")
    If rsFacturas.RecordCount <> 0 Then
        cboFactura.Clear
        Do While Not rsFacturas.EOF
            cboFactura.AddItem rsFacturas!FolioFactura
            rsFacturas.MoveNext
        Loop
        cboFactura.ListIndex = 0
    Else
        chkFacturasPagadas.Value = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFacturas"))
    Unload Me
End Sub

Private Sub pFacturasPaciente()
On Error GoTo NotificaError
    
    Dim rsFacturasPaciente As New ADODB.Recordset
    Dim strParametros As String
    
    strParametros = txtCveCliente.Text & "|" & IIf(OptTipoPaciente(0).Value = True, "I", "E")
    Set rsFacturasPaciente = frsEjecuta_SP(strParametros, "SP_CCSELFACTURASPACIENTE")
    
    cboFacturasPaciente.Clear
    If rsFacturasPaciente.RecordCount <> 0 Then
        Do While Not rsFacturasPaciente.EOF
            cboFacturasPaciente.AddItem rsFacturasPaciente!FolioFactura
            rsFacturasPaciente.MoveNext
        Loop
        cboFacturasPaciente.ListIndex = 0
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFacturasPaciente"))
    Unload Me
End Sub

Private Function pDatosCliente(vlstrNumCliente As String) As Boolean
On Error GoTo NotificaError
    Dim vlstrsql As String
    
    Dim vlstrCatalgo, vlstrTipo As String
    Dim rsRegimen As New ADODB.Recordset
    Dim rsDatosFISC As New ADODB.Recordset
    Set rsDatosCliente = frsEjecuta_SP(Str(vlstrNumCliente) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|1", "sp_CcSelDatosCliente")
    
    'Si la version de CFDI es 4.0, valida que el cliente tenga el codigo postal y el regimen fiscal
    pDatosCliente = True
    
    
    If vgstrVersionCFDI = "4.0" Then
        If rsDatosCliente.RecordCount <> 0 Then
            If rsDatosCliente!chrTipoCliente = "CO" Then
                vlstrsql = "select vchRegimenFiscal, vchCodigoPostal  from CCEmpresa where intcveEmpresa = " & rsDatosCliente!intNumReferencia
                vlstrTipo = "empresa"
                vlstrCatalgo = "empresas"
            ElseIf rsDatosCliente!chrTipoCliente = "ME" Then
                   vlstrsql = "select vchRegimenFiscal, VCHCONSULCODPOSTAL as vchCodigoPostal from HoMedico where intcveMedico = " & rsDatosCliente!intNumReferencia
                   vlstrCatalgo = "médicos"
                   vlstrTipo = "médico"
            ElseIf rsDatosCliente!chrTipoCliente = "EM" Then
                 vlstrsql = "select vchRegimenFiscal, chrCodigoPostal as vchCodigoPostal from NoEmpleado where intcveEmpleado = " & rsDatosCliente!intNumReferencia
                 vlstrTipo = "empleado"
                 vlstrCatalgo = "empleados"
            Else
                Set rsDatosFISC = frsRegresaRs("SELECT VCHREGIMENFISCAL, VCHCODIGOPOSTAL, CHRNOMBRE, CHRRFC FROM PVDATOSFISCALES WHERE VCHCODIGOPOSTAL IS NOT NULL AND VCHREGIMENFISCAL IS NOT NULL AND INTNUMCUENTA = " & rsDatosCliente!intNumReferencia & " ORDER BY INTNUMCUENTA DESC")
                If rsDatosFISC.RecordCount = 0 Then
                    Set rsDatosFISC = frsRegresaRs("SELECT VCHREGIMENFISCAL, VCHCODIGOPOSTAL, CHRNOMBRE, CHRRFC FROM PVDATOSFISCALES WHERE VCHCODIGOPOSTAL IS NOT NULL AND VCHREGIMENFISCAL IS NOT NULL AND TRIM(CHRRFC) = '" & Trim(rsDatosCliente!RFCCliente) & "' ORDER BY INTNUMCUENTA DESC")
                End If
                vlstrTipo = "paciente"
                vlstrsql = ""
            End If
        
        'End If
            If vlstrsql <> "" Then
                Set rsRegimen = frsRegresaRs(vlstrsql)
                If rsRegimen.RecordCount <> 0 Then
                    vlstrCodigo = IIf(IsNull(rsRegimen!VCHCODIGOPOSTAL), "", rsRegimen!VCHCODIGOPOSTAL)
                    vlstrRegimen = IIf(IsNull(rsRegimen!vchregimenfiscal), "", rsRegimen!vchregimenfiscal)
                    
                    vlRazonSocialComprobante = ""
                    vlRFCComprobante = ""
                    
                    If vlstrCodigo = "" Then
                        txtCveCliente.Text = ""
                        MsgBox "No se encuentra configurado el código postal correspondiente al cliente " & Trim(rsDatosCliente!NombreCliente) & " de tipo " & vlstrTipo & ", se puede configurar en el catálogo de " & vlstrCatalgo, vbExclamation + vbOKOnly, "Mensaje"
                       pDatosCliente = False
                       Exit Function
                    End If
                    If vlstrRegimen = "" Then
                       txtCveCliente.Text = ""
                        MsgBox "No se encuentra configurado el régimen fiscal correspondiente al cliente " & Trim(rsDatosCliente!NombreCliente) & " de tipo " & vlstrTipo & ", se puede configurar en el catálogo de " & vlstrCatalgo, vbExclamation + vbOKOnly, "Mensaje"
                       pDatosCliente = False
                       Exit Function
                    End If
                End If
            Else
                If rsDatosFISC.RecordCount > 0 Then
                    vlstrCodigo = IIf(IsNull(rsDatosFISC!VCHCODIGOPOSTAL), "", Trim(rsDatosFISC!VCHCODIGOPOSTAL))
                    vlstrRegimen = IIf(IsNull(rsDatosFISC!vchregimenfiscal), "", Trim(rsDatosFISC!vchregimenfiscal))
                    
                    vlRazonSocialComprobante = IIf(IsNull(rsDatosFISC!CHRNOMBRE), "", Trim(rsDatosFISC!CHRNOMBRE))
                    vlRFCComprobante = IIf(IsNull(rsDatosFISC!chrRFC), "", Trim(rsDatosFISC!chrRFC))
                Else
                    vlstrCodigo = ""
                    vlstrRegimen = ""
                    
                    vlRazonSocialComprobante = ""
                    vlRFCComprobante = ""
                End If
                If vlstrCodigo = "" Then
                    txtCveCliente.Text = ""
                    MsgBox "No se encuentra configurado el código postal correspondiente al cliente " & Trim(rsDatosCliente!NombreCliente) & " de tipo " & vlstrTipo & ", se puede configurar en la información del paciente.", vbExclamation + vbOKOnly, "Mensaje"
                   pDatosCliente = False
                   Exit Function
                End If
                If vlstrRegimen = "" Then
                   txtCveCliente.Text = ""
                    MsgBox "No se encuentra configurado el régimen fiscal correspondiente al cliente " & Trim(rsDatosCliente!NombreCliente) & " de tipo " & vlstrTipo & ", se puede configurar en la información del paciente.", vbExclamation + vbOKOnly, "Mensaje"
                   pDatosCliente = False
                   Exit Function
                End If
            End If
        End If
    End If
         
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDatosCliente"))
    Unload Me
End Function

Private Sub pAsignarTxtNombre(vlstrNombre As String, txtNombre As TextBox)
    txtNombre.Text = vlstrNombre
End Sub

Private Sub pAsignarTxtDireccion(vlstrDireccion As String, txtDireccion As TextBox)
    txtDireccion.Text = vlstrDireccion
End Sub

Private Sub pAsignarTxtRFC(vlstrRFC As String, txtRFC As TextBox)
    txtRFC.Text = vlstrRFC
End Sub

Private Sub pAsignarTxtDomicilio(vlstrDomicilio As String, txtDomicilio As TextBox)
    txtDomicilio.Text = vlstrDomicilio
End Sub

Private Sub txtCveCliente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        Select Case UCase(Chr(KeyAscii))
            Case "I"
                If OptTipoPaciente(0).Enabled Then
                    OptTipoPaciente(0).Value = True
                    txtCveCliente.SetFocus
                End If
            Case "E"
                If OptTipoPaciente(1).Enabled Then
                    OptTipoPaciente(1).Value = True
                    txtCveCliente.SetFocus
                End If
        End Select
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveCliente_KeyPress"))
    Unload Me
End Sub

Private Sub txtDescuento_GotFocus()
    pSelTextBox txtDescuento
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If cmdIncluir.Enabled Then
            cmdIncluir.SetFocus
        Else
            txtCantidad.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_KeyDown"))
    Unload Me
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
   
    If Not fblnFormatoCantidad(txtDescuento, KeyAscii, 2) Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_KeyPress"))
    Unload Me
End Sub

Private Function fintErrorCancelar() As Integer
    '----------------------------------------------------------------------------------------------------------------------'
    ' Función que revisa que la nota de cargo no estés incluida en un paquete de cobranza o que no tenga pagos registrados '
    '----------------------------------------------------------------------------------------------------------------------'
    Dim rs As New ADODB.Recordset
    Dim rsPagos As New ADODB.Recordset
    
    fintErrorCancelar = 0
    'que el o los créditos de la factura no tengan pagos registrados
    vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha) & _
                        "|" & fstrFechaSQL(fdtmServerFecha) & _
                        "|" & txtCveCliente.Text & _
                        "|" & "0" & _
                        "|" & "CA" & _
                        "|" & "0" & _
                        "|" & Trim(lblFolio.Caption) & _
                        "|" & "0" & _
                        "|" & "0" & _
                        "|" & "*" & _
                        "|" & "0"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_CcSelCredito")
    If rs.RecordCount <> 0 Then
        If IsDate(rs!fechaEnvio) Then
            'No se puede cancelar el documento, los créditos fueron incluídos en un paquete de cobranza.
            fintErrorCancelar = 718
        Else
            vgstrParametrosSP = Str(rs!Movimiento) & "|" & "0" & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "I"
            Set rsPagos = frsEjecuta_SP(vgstrParametrosSP, "Sp_CcSelPagosCredito")
            If rsPagos.RecordCount <> 0 Then
                'No se puede cancelar el documento  el crédito tiene pagos registrados.
                fintErrorCancelar = 368
            End If
            rsPagos.Close
        End If
    Else
        '¡La información no existe!
        fintErrorCancelar = 12
        Exit Function
    End If
    rs.Close
End Function

Private Sub pCreditosDirectos()
On Error GoTo NotificaError
    
    Dim rsCreditosDirectos As New ADODB.Recordset
    
    Set rsCreditosDirectos = frsEjecuta_SP(CInt(txtCveCliente.Text), "SP_CCSELCREDIRECTOS")
    If rsCreditosDirectos.RecordCount <> 0 Then
        Do While Not rsCreditosDirectos.EOF
            cboCreditosDirectos.AddItem rsCreditosDirectos!FolioCreditoDirecto
            cboCreditosDirectos.ItemData(cboCreditosDirectos.newIndex) = rsCreditosDirectos!NumMovimiento
            rsCreditosDirectos.MoveNext
        Loop
        cboCreditosDirectos.ListIndex = 0
    End If
    rsCreditosDirectos.Close
    
    Set rsCreditosDirectos = frsRegresaRs("SELECT smiCveConcepto, chrDescripcion FROM PvConceptoFacturacion WHERE bitActivo = 1 ORDER BY chrDescripcion")
    While Not rsCreditosDirectos.EOF
        lstConceptosFact.AddItem rsCreditosDirectos!chrdescripcion
        lstConceptosFact.ItemData(lstConceptosFact.newIndex) = rsCreditosDirectos!smicveconcepto
        rsCreditosDirectos.MoveNext
    Wend
    rsCreditosDirectos.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCreditosDirectos"))
    Unload Me
End Sub

Private Sub InicializaVariables()
    vldblSubTotalTemporal = 0
    vldblDescuentoTemporal = 0
    vldblIvaTemporal = 0
    vldTotal = 0
End Sub

Private Sub pAjustaIVA(grid As MSHFlexGrid)
On Error GoTo NotificaError

    Dim vllngContador As Integer
    Dim vldblIVA As Double
    Dim dbldiferencia As Double
    Dim rsCantidadCredito As ADODB.Recordset
    Dim dblCantidad As Double
    Dim dblDescuento As Double
    Dim dblIVA As Double
    Dim dblCantidadCredito As Double
    Dim strParametrosCredito As String
    Dim intNumeroFactura As Integer
    Dim vllngContadorFactura As Integer
    Dim vlstrFormatoLargo As String
    
    vlstrFormatoLargo = "###############.0000000000000000"
    
    intLimiteAjusteIVA = intLimiteAjusteIVA + 1
    
    If BlnAjusteIVA = False And intLimiteAjusteIVA < 1000 Then
        With grid
            intNumeroFactura = 0
            
            ' Limpia IVA de la estructura
            'For vllngContador = 0 To UBound(aFacturas(), 1)
            '    aFacturas(vllngContador).vldblIVA = 0
            'Next vllngContador
            
            ' Guarda IVA en la estructura aFacturas
            For vllngContador = 1 To .Rows - 1
                vldblIVA = vldblIVA + Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo))
            
                'If aFacturas(intNumeroFactura).vlstrFolioFactura <> .TextMatrix(vllngContador, 1) And intNumeroFactura < UBound(aFacturas(), 1) Then
                '    intNumeroFactura = intNumeroFactura + 1
                'End If
                
                'aFacturas(intNumeroFactura).vldblIVA = aFacturas(intNumeroFactura).vldblIVA + Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo))
            Next vllngContador
        End With
        
        ' If para validar grid con IVA mayor
        If Round(vldblIVA, 2) - Round(CDbl(txtIVA.Text), 2) > 0 And Round(vldblIVA, 2) - Round(CDbl(txtIVA.Text), 2) < 0.05 Then
            With grid
                If (.Rows - 1) > 1 Then
                    If .Rows - 1 = IntAjusteIVAContador Then IntAjusteIVAContador = 0
                    
                    For vllngContador = 1 + IntAjusteIVAContador To .Rows - 1
                        ' Hace ajuste al IVA
                        If BlnAjusteIVA = False And .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) <> 0 Then
                            If sstFacturasCreditos.Tab = 1 Then
                                strParametrosCredito = Trim(.TextMatrix(vllngContador, vlintColFactura)) & "|" & "MA"
                                Set rsCantidadCredito = frsEjecuta_SP(strParametrosCredito, "SP_CANTIDADCREDITO")
                                If rsCantidadCredito.RecordCount > 0 Then
                                    dblCantidad = Val(Format(.TextMatrix(vllngContador, vlintColCantidad), vlstrFormato))
                                    dblDescuento = Val(Format(.TextMatrix(vllngContador, vlintColDescuento), vlstrFormato))
                                    dblIVA = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo))
                                    dblCantidadCredito = dblCantidad - dblDescuento + dblIVA
                                    
                                    If dblCantidadCredito > rsCantidadCredito(0).Value Then
                                        .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) - 0.01
                                        IntAjusteIVAContador = vllngContador
                                        Call pAjustaIVA(grid)
                                    End If
                                End If
                                rsCantidadCredito.Close
                            Else
                                .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) - 0.01
                                IntAjusteIVAContador = vllngContador
                                Call pAjustaIVA(grid)
                            End If  ' End if sstFacturasCreditos.Tab
                        End If  ' End if hace ajuste al IVA
                    Next vllngContador
                Else
                    If BlnAjusteIVA = False And .TextMatrix(1, vlintColIVANotaSinRedondear) <> 0 Then
                        .TextMatrix(1, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(1, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) - 0.01
                        Call pAjustaIVA(grid)
                    End If
                End If  ' End if .Rows > 1
            End With ' End Grid
        Else
            ' If para validar grid con IVA menor
            If Round(CDbl(txtIVA.Text), 2) - Round(vldblIVA, 2) > 0 And Round(CDbl(txtIVA.Text), 2) - Round(vldblIVA, 2) < 0.05 Then
                With grid
                    If (.Rows - 1) > 1 Then
                        If .Rows - 1 = IntAjusteIVAContador Then IntAjusteIVAContador = 0
                        
                        For vllngContador = 1 + IntAjusteIVAContador To .Rows - 1
                            ' Hace ajuste al IVA
                            If BlnAjusteIVA = False And .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) <> 0 Then
                                If sstFacturasCreditos.Tab = 1 Then
                                    strParametrosCredito = Trim(.TextMatrix(vllngContador, vlintColFactura)) & "|" & "MA"
                                    Set rsCantidadCredito = frsEjecuta_SP(strParametrosCredito, "SP_CANTIDADCREDITO")
                                    If rsCantidadCredito.RecordCount > 0 Then
                                        dblCantidad = Val(Format(.TextMatrix(vllngContador, vlintColCantidad), vlstrFormato))
                                        dblDescuento = Val(Format(.TextMatrix(vllngContador, vlintColDescuento), vlstrFormato))
                                        dblIVA = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo))
                                        dblCantidadCredito = dblCantidad - dblDescuento + dblIVA
                                        
                                        If rsCantidadCredito(0).Value > dblCantidadCredito Then
                                            .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) + 0.01
                                            IntAjusteIVAContador = vllngContador
                                            Call pAjustaIVA(grid)
                                        End If
                                    End If
                                    rsCantidadCredito.Close
                                Else
                                    .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) + 0.01
                                    IntAjusteIVAContador = vllngContador
                                    Call pAjustaIVA(grid)
                                End If  ' End If sstFacturasCreditos.Tab
                            End If  ' End If BlnAjusteIVA
                        Next vllngContador
                    Else
                        If BlnAjusteIVA = False And .TextMatrix(1, vlintColIVANotaSinRedondear) <> 0 Then
                            .TextMatrix(1, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(1, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) + 0.01
                            Call pAjustaIVA(grid)
                        End If
                    End If  ' End if .Rows > 1
                End With    ' End Grid
            Else
                BlnAjusteIVA = True
                
                For vllngContadorFactura = 0 To UBound(aFacturas(), 1)
                    'Validar que el monto no exeda el crédito
                    If sstFacturasCreditos.Tab = 1 Then
                        strParametrosCredito = aFacturas(vllngContadorFactura).vlstrFolioFactura & "|" & "MA"
                        Set rsCantidadCredito = frsEjecuta_SP(strParametrosCredito, "SP_CANTIDADCREDITO")
                        If rsCantidadCredito.RecordCount > 0 Then
                            If (aFacturas(vllngContadorFactura).vldblSubtotal - aFacturas(vllngContadorFactura).vldblDescuento _
                                + aFacturas(vllngContadorFactura).vldblIVA) > rsCantidadCredito(0).Value Then
                                With grid
                                    If .Rows - 1 = IntAjusteIVAContador Then IntAjusteIVAContador = 0
                                    
                                    For vllngContador = 1 + IntAjusteIVAContador To .Rows - 1
                                        If .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) <> 0 Then
                                            .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) - 0.01
                                            .TextMatrix(vllngContador, vlintColIVA) = Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormato)
                                            
                                            vllngContador = .Rows - 1
                                        End If
                                    Next vllngContador
                                End With
                                
                                aFacturas(vllngContadorFactura).vldblIVA = aFacturas(vllngContadorFactura).vldblIVA - 0.01
                                                        
                                txtIVA.Text = FormatCurrency(CDbl(txtIVA.Text) - 0.01, 2)
                                txtTotal.Text = FormatCurrency(CDbl(txtTotal.Text) - 0.01, 2)
                            End If ' End if compara crédito
                        End If ' End if rsCantidadCredito.RecordCount
                    Else
                        strParametrosCredito = aFacturas(vllngContadorFactura).vlstrFolioFactura & "|" & "FA"
                        Set rsCantidadCredito = frsEjecuta_SP(strParametrosCredito, "SP_CANTIDADCREDITO")
                        If rsCantidadCredito.RecordCount > 0 Then
                            If Round(aFacturas(vllngContadorFactura).vldblSubtotal - aFacturas(vllngContadorFactura).vldblDescuento _
                                + aFacturas(vllngContadorFactura).vldblIVA, 2) > rsCantidadCredito(0).Value Then
                                With grid
                                    If .Rows - 1 = IntAjusteIVAContador Then IntAjusteIVAContador = 0
                                    
                                    For vllngContador = 1 + IntAjusteIVAContador To .Rows - 1
                                        If .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) <> 0 Then
                                            .TextMatrix(vllngContador, vlintColIVANotaSinRedondear) = Val(Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormatoLargo)) - 0.01
                                            .TextMatrix(vllngContador, vlintColIVA) = Format(.TextMatrix(vllngContador, vlintColIVANotaSinRedondear), vlstrFormato)
                                            
                                            vllngContador = .Rows - 1
                                        End If
                                    Next vllngContador
                                End With
                                
                                aFacturas(vllngContadorFactura).vldblIVA = aFacturas(vllngContadorFactura).vldblIVA - 0.01
                                txtIVA.Text = FormatCurrency(CDbl(txtIVA.Text) - 0.01, 2)
                                txtTotal.Text = FormatCurrency(CDbl(txtTotal.Text) - 0.01, 2)
                            End If  ' End if compara crédito
                        End If ' End if rsCantidadCredito.RecordCount
                    End If  ' End if sstFacturasCreditos.Tab
                    
                    rsCantidadCredito.Close
                Next vllngContadorFactura  ' End For Facturas
            End If  '  End If para validar grid con IVA menor
        End If  ' End If para validar grid con IVA mayor
    End If ' End if BlnAjusteIVA

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAjustaIVA"))
    Unload Me
End Sub

Private Sub pConfiguraGridNotaInfo()
On Error GoTo NotificaError

    With grdInformacionNota
       .Clear
       .Cols = 7
       .Rows = 2
       .FixedCols = 1
       .FixedRows = 1
       .FormatString = "||Folio nota|Cantidad|Descuento|IVA|Total"
       .ColWidth(0) = 100     'Fixed              (0)
       .ColWidth(1) = 0       'Consecutivo        (1)
       .ColWidth(2) = 2000    'Folio              (2)
       .ColWidth(3) = 1200    'Cantidad           (3)
       .ColWidth(4) = 1100    'Descuento          (4)
       .ColWidth(5) = 1100    'IVA                (5)
       .ColWidth(6) = 1200    'Total              (6)
       
       .ColAlignment(1) = flexAlignLeftCenter
       .ColAlignment(2) = flexAlignRightCenter
       .ColAlignment(3) = flexAlignRightCenter
       .ColAlignment(4) = flexAlignRightCenter
       .ColAlignment(5) = flexAlignRightCenter
       .ColAlignment(6) = flexAlignRightCenter
       
       .ColAlignmentFixed(1) = flexAlignCenterCenter
       .ColAlignmentFixed(2) = flexAlignCenterCenter
       .ColAlignmentFixed(3) = flexAlignCenterCenter
       .ColAlignmentFixed(4) = flexAlignCenterCenter
       .ColAlignmentFixed(5) = flexAlignCenterCenter
       .ColAlignmentFixed(6) = flexAlignCenterCenter
       
       .ScrollBars = flexScrollBarBoth
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridNotaInfo"))
    Unload Me
End Sub

Private Sub pFormatoGridNotaInfo()
On Error GoTo NotificaError

    With grdInformacionNota
       .FormatString = "||Folio nota|Cantidad|Descuento|IVA|Total"
       .ColWidth(0) = 100     'Fixed              (0)
       .ColWidth(1) = 0       'Consecutivo        (1)
       .ColWidth(2) = 2000    'Folio              (2)
       .ColWidth(3) = 1200    'Cantidad           (3)
       .ColWidth(4) = 1100    'Descuento          (4)
       .ColWidth(5) = 1100    'IVA                (5)
       .ColWidth(6) = 1200    'Total              (6)
       
       .ColAlignment(1) = flexAlignLeftCenter
       .ColAlignment(2) = flexAlignLeftCenter
       .ColAlignment(3) = flexAlignRightCenter
       .ColAlignment(4) = flexAlignRightCenter
       .ColAlignment(5) = flexAlignRightCenter
       .ColAlignment(6) = flexAlignRightCenter
       
       .ColAlignmentFixed(1) = flexAlignCenterCenter
       .ColAlignmentFixed(2) = flexAlignCenterCenter
       .ColAlignmentFixed(3) = flexAlignCenterCenter
       .ColAlignmentFixed(4) = flexAlignCenterCenter
       .ColAlignmentFixed(5) = flexAlignCenterCenter
       .ColAlignmentFixed(6) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoGridNotaInfo"))
    Unload Me
End Sub

Public Function flngCuentaConceptoDeptoNota(vllngConceptoFactura As Long, vlintNumeroDepartamento As Integer, vlstrTipoCuenta As String, Optional vllngCveTipoPaciente As Long) As Long
'-----------------------------------------------------------------------------------------------------------------------'
' Regresa la cuenta contable del ingreso, descuento o nota de crédito del concepto de facturación según el departamento '
'-----------------------------------------------------------------------------------------------------------------------'
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlblnEncontro As Boolean
    Dim vlstrCuenta As String
    
    flngCuentaConceptoDeptoNota = 0
    
    vlblnEncontro = False
    
    'Identifica si es socio (la clave de tipo paciente es <> 0) busca por tipo de paciente socio
    If vllngCveTipoPaciente > 0 Then 'Por tipo de paciente
        vlstrSentencia = "SELECT " & IIf(Trim(vlstrTipoCuenta) = "INGRESO", "intNumCuentaIngreso", "intNumCuentaDescuento") & _
                         " FROM PvConceptoFactPaciente " & _
                         " WHERE smiCveConcepto = " & Trim(Str(vllngConceptoFactura)) & _
                         " AND smiCveTipoPaciente = " & Trim(Str(vllngCveTipoPaciente))
    Else
        'Si no es socio, busca por departamento
        Select Case Trim(vlstrTipoCuenta)
            Case "INGRESO"
                vlstrCuenta = "intNumCuentaIngreso"
            Case "DESCUENTO"
                vlstrCuenta = "intNumCuentaDescuento"
            Case "NOTA"
                vlstrCuenta = "intNumCuentaDescNota"
        End Select
        'vlstrSentencia = "SELECT " & IIf(Trim(vlstrTipoCuenta) = "INGRESO", "intNumCuentaIngreso", "intNumCuentaDescuento")
        vlstrSentencia = "SELECT " & vlstrCuenta & _
                         " FROM PvConceptoFacturacionDepartame " & _
                         " WHERE smiCveConcepto = " & Trim(Str(vllngConceptoFactura)) & _
                         " AND smiCveDepartamento = " & Trim(Str(vlintNumeroDepartamento))
    End If
    Set rs = frsRegresaRs(vlstrSentencia)
    If rs.RecordCount <> 0 Then
        If rs.Fields(0) <> 0 Then
            flngCuentaConceptoDeptoNota = rs.Fields(0)
            vlblnEncontro = True
        End If
    End If
    
    If Not vlblnEncontro And vllngCveTipoPaciente = 0 Then 'Busca por empresa como última opción cuando NO es socio
        Select Case Trim(vlstrTipoCuenta)
            Case "INGRESO"
                vlstrCuenta = "intNumCtaIngreso"
            Case "DESCUENTO"
                vlstrCuenta = "intNumCtaDescuento"
            Case "NOTA"
                vlstrCuenta = "intNumCtaDescNota"
        End Select
    
        vlstrSentencia = "SELECT " & vlstrCuenta & _
                         " FROM PvConceptoFacturacionEmpresa " & _
                         " WHERE intCveConceptoFactura = " & Str(vllngConceptoFactura) & _
                         " AND intCveEmpresaContable = " & vgintClaveEmpresaContable
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount <> 0 Then
            flngCuentaConceptoDeptoNota = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngCuentaConceptoDeptoNota"))
End Function

'- Indica si la Nota de crédito/cargo es CFDi y el PAC activo es PAX -'
'- Modificado para la cancelación con Buzón Fiscal -'
Private Function fblnNotaCFDI(llngConsecutivoNota As Long, lstrTipo As String) As Boolean
On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim lstrSentencia As String
    
    '- Revisar que la nota sea CFDi -'
    lstrSentencia = "SELECT vchUUID FROM GnComprobanteFiscalDigital " & _
                    " WHERE intComprobante = " & llngConsecutivoNota & " AND chrTipoComprobante = '" & lstrTipo & "'"
    Set rs = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
    fblnNotaCFDI = IIf(IsNull(rs!VCHUUID), False, True) 'Si no existe UUID entonces es CFD
    rs.Close
    
    If fblnNotaCFDI Then
        '- Revisar si el PAC activo es PAX -'
        '- Modificado para caso 8886 para Buzón Fiscal -'
        Set rs = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGPAC")
        If rs.RecordCount > 0 Then
            fblnNotaCFDI = True 'Val(rs!PAC) = 2
        Else
            fblnNotaCFDI = False 'Si no hay PAC activo
        End If
        rs.Close
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
End Function

'- CASO 6217: Verifica si se va a mostrar la pantalla de envío de CFD por correo electrónico -'
Private Function fblnPermitirEnvio() As Boolean
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    'Dim lstrTipo As String
    
    fblnPermitirEnvio = False
    
    '- Revisar que el parámetro de envío de CFD esté activado -'
    If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
        fblnPermitirEnvio = True
    Else
        Exit Function
    End If
    
    fblnPermitirEnvio = False
    
    If optCliente.Value = True Then
        '- Revisar que el cliente no pertenezca a una empresa -'
        vlstrSentencia = "SELECT chrTipoCliente FROM CcCliente WHERE intNumCliente = " & Trim(txtCveCliente.Text)
        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        If rs.RecordCount <> 0 Then
            If Trim(rs!chrTipoCliente) <> "CO" Then
                fblnPermitirEnvio = True
            End If
        End If
    Else
        '- Revisar que la factura no pertenezca a una empresa -'
'        Set rs = frsEjecuta_SP(Trim(cboFacturasPaciente.Text), "SP_PvSelFactura")
'        If rs.RecordCount > 0 Then
'            lstrTipo = rs!chrTipoFactura
'        End If
'        rs.Close
    
        '- Revisar que el paciente no pertenezca a un convenio -'
        vgstrParametrosSP = Trim(txtCveCliente.Text) & "|" & Str(vgintClaveEmpresaContable)
        If OptTipoPaciente(0).Value = True Then
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
        Else
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
        End If
        If rs.RecordCount > 0 Then
            If IsNull(rs!cveEmpresa) Or chkFacturasPaciente.Value = 0 Then 'lstrTipo = "P" Then
                fblnPermitirEnvio = True
            End If
        End If
        rs.Close
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPermitirEnvio"))
End Function

Private Sub pCargaUsosCFDI()
On Error GoTo NotificaError

Dim rsTmp As ADODB.Recordset

    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDI, rsTmp, 0, 1
        cboUsoCFDI.ListIndex = -1
    End If
    rsTmp.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaUsosCFDI"))
End Sub

Private Sub pCargaFormasPago()
On Error GoTo NotificaError

Dim rsTmp As ADODB.Recordset
Dim strSql As String

    strSql = "SELECT INTIDREGISTRO, TRIM(VCHCLAVE) ||' - ' || TRIM(VCHDESCRIPCION) NOMBRE FROM GNCATALOGOSATDETALLE WHERE INTIDCATALOGOSAT = 2 AND BITACTIVO = 1"
    Set rsTmp = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rsTmp.EOF Then
        pLlenarCboRs cboFormaPago, rsTmp, 0, 1
        cboFormaPago.ListIndex = -1
    End If
    rsTmp.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaFormasPago"))
End Sub

Private Sub pCargaMetodoPago()
On Error GoTo NotificaError

    cboMetodoPago.AddItem "PUE - Pago en una sola exhibición", 0
    cboMetodoPago.AddItem "PPD - Pago en parcialidades o diferido", 1
    cboMetodoPago.ListIndex = -1
         
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaMetodoPago"))
End Sub

Private Function fblnValidaSAT() As Boolean
    Dim intRow As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    If vgstrVersionCFDI <> "3.2" Then
    
        strSql = "select IVArticulo.intCveUniMinimaVta ""cveUnidad""" & _
        " from IVArticulo" & _
        " where IVArticulo.intIdArticulo = "

       
        'Uso del CFDI
        If cboUsoCFDI.ListIndex = -1 Then
            MsgBox "Seleccione el uso del comprobante", vbExclamation, "Mensaje"
            cboUsoCFDI.SetFocus
            fblnValidaSAT = False
            Exit Function
        End If
        
        For intRow = 1 To grdNotas.Rows - 1
            If grdNotas.TextMatrix(intRow, 11) = "AR" Then
                If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", grdNotas.TextMatrix(intRow, 7), "AR", 0) = 0 Then
                    MsgBox "No está definida la clave del SAT para el producto/servicio " & grdNotas.TextMatrix(intRow, 2), vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    Exit Function
                End If
                Set rsTmp = frsRegresaRs(strSql & grdNotas.TextMatrix(intRow, 7), adLockReadOnly, adOpenForwardOnly)
                If Not rsTmp.EOF Then
                    If flngCatalogoSATIdByNombreTipo("c_ClaveUnidad", rsTmp!cveUnidad, "UV", 0) = 0 Then
                        MsgBox "No está definida la clave del SAT para la unidad del producto/servicio " & grdNotas.TextMatrix(intRow, 2), vbExclamation, "Mensaje"
                        fblnValidaSAT = False
                        Exit Function
                    End If
                Else
                    MsgBox "No está definida la clave del SAT para la unidad del producto/servicio " & grdNotas.TextMatrix(intRow, 2), vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    Exit Function
                End If
                rsTmp.Close
            Else
                If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", grdNotas.TextMatrix(intRow, 7), grdNotas.TextMatrix(intRow, 11), 1) = 0 Then
                    MsgBox "No está definida la clave del SAT para el producto/servicio " & grdNotas.TextMatrix(intRow, 2), vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    Exit Function
                End If
                If flngCatalogoSATIdByNombreTipo("c_ClaveUnidad", grdNotas.TextMatrix(intRow, 7), grdNotas.TextMatrix(intRow, 11), 2) = 0 Then
                    MsgBox "No está definida la clave del SAT para la unidad del producto/servicio " & grdNotas.TextMatrix(intRow, 2), vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    Exit Function
                End If
            End If
        Next
        
        fblnValidaSAT = True
    
    Else
        fblnValidaSAT = True
    End If
        
End Function

Private Function fdblIVAConcepto(lngCveConcepto As Long, strTipoConcepto As String) As Double
    Dim rsTmp As ADODB.Recordset
    Dim lngCveConFact As Long
    Dim strSql As String
    lngCveConFact = 0
    Select Case strTipoConcepto
        Case "AR": strSql = "select smiCveConceptFact from IVArticulo where intIdArticulo = " & lngCveConcepto
        Case "OC": strSql = "select smiConceptoFact from PVOtroConcepto where intCveConcepto = " & lngCveConcepto
        Case "ES": strSql = "select smiConFact from IMEstudio where intCveEstudio = " & lngCveConcepto
        Case "EX": strSql = "select smiConFact from LAExamen where intCveExamen = " & lngCveConcepto
        Case "GE": strSql = "select smiConFact from LAGrupoExamen where intCveGrupo = " & lngCveConcepto
    End Select
    If strTipoConcepto <> "CF" Then
        Set rsTmp = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
        If Not rsTmp.EOF Then
            lngCveConFact = rsTmp.Fields(0).Value
        End If
        rsTmp.Close
    Else
        lngCveConFact = lngCveConcepto
    End If
    
    strSql = "select smyIVA from PVConceptoFacturacion where smiCveConcepto = " & lngCveConFact
    Set rsTmp = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rsTmp.EOF Then
        fdblIVAConcepto = IIf(rsTmp.Fields(0).Value = 0, vgdblCantidadIvaGeneral, rsTmp.Fields(0).Value)
    Else
        fdblIVAConcepto = vgdblCantidadIvaGeneral
    End If
    rsTmp.Close
End Function

Private Function fintLocalizaCritCboFormaMetodoUso(ObjCbo As ComboBox, vlstrCriterio As String) As Integer
'-------------------------------------------------------------------------------------------
' Busca un criterio dentro del combobox para las formas de pago, metodo de pago y uso de CFDI
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumReg As Integer
    Dim vlintseq As Integer
    Dim vlintInicio As Integer
    
    vlintNumReg = ObjCbo.ListCount
    If Len(vlstrCriterio) > 0 Then
        For vlintseq = 0 To vlintNumReg
            vlintInicio = 0
            vlintInicio = InStr(1, ObjCbo.List(vlintseq), vlstrCriterio)
            If vlintInicio > 0 Then
                fintLocalizaCritCboFormaMetodoUso = vlintseq
                Exit For
            Else
                fintLocalizaCritCboFormaMetodoUso = -1
            End If
        Next vlintseq
    Else
        fintLocalizaCritCboFormaMetodoUso = -1
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaCritCboFormaMetodoUso"))
End Function

Private Function fdblTasaIVAEmpresa(vllngEmpresaCat As Integer) As Double
    Dim ObjRs As New ADODB.Recordset
    Dim objSTR As String
    
    fdblTasaIVAEmpresa = 0
    
    If vllngEmpresaCat = 0 Then
        fdblTasaIVAEmpresa = vgdblCantidadIvaGeneral
    Else
        objSTR = "select relporcentaje from ccempresa inner join cnimpuesto on cnimpuesto.smicveimpuesto = SMICVEIMPUESTOCONCEPSSEG where intcveempresa = " & vllngEmpresaCat
        Set ObjRs = frsRegresaRs(objSTR, adLockOptimistic)
        If ObjRs.RecordCount > 0 Then
            fdblTasaIVAEmpresa = ObjRs!relPorcentaje
        Else
            fdblTasaIVAEmpresa = vgdblCantidadIvaGeneral
        End If
    End If

End Function

Private Function flngCveEmpresaCliente(vlClave As Long) As Long
    Dim ObjRs As New ADODB.Recordset
    Dim objSTR As String
    
    flngCveEmpresaCliente = 0
    
    If vlClave = 0 Then
        flngCveEmpresaCliente = 0
    Else
        objSTR = "Select intNumReferencia From CcCliente Where chrtipocliente = 'CO' and intNumCliente = " & vlClave
        Set ObjRs = frsRegresaRs(objSTR, adLockOptimistic)
        If ObjRs.RecordCount > 0 Then
            flngCveEmpresaCliente = ObjRs!intNumReferencia
        Else
            flngCveEmpresaCliente = 0
        End If
    End If

End Function

Private Function flngCveEmpresaPaciente(vlClave As Long) As Long
    Dim ObjRs As New ADODB.Recordset
    Dim objSTR As String
    
    flngCveEmpresaPaciente = 0
    
    If vlClave = 0 Then
        flngCveEmpresaPaciente = 0
    Else
        objSTR = "select intcveempresa from EXPacienteIngreso where intNumCuenta = " & vlClave
        Set ObjRs = frsRegresaRs(objSTR, adLockOptimistic)
        If ObjRs.RecordCount > 0 Then
            flngCveEmpresaPaciente = IIf(IsNull(ObjRs!intcveempresa), 0, ObjRs!intcveempresa)
        Else
            flngCveEmpresaPaciente = 0
        End If
    End If

End Function
