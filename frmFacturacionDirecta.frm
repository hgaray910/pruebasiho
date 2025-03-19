VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturacionDirecta 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación directa a clientes"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   1320
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   8760
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   36
         Top             =   600
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarraCFD 
         BackColor       =   &H80000002&
         Caption         =   "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
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
         Left            =   65
         TabIndex        =   37
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
         Top             =   110
         Width           =   8700
      End
   End
   Begin TabDlg.SSTab SSTFactura 
      Height          =   10245
      Left            =   0
      TabIndex        =   38
      Top             =   -645
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   18071
      _Version        =   393216
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmFacturacionDirecta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboUsoCFDI"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkMovimiento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraTotales"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraMoneda"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraFolioFecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraCliente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraDetalle"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkFacturaSustitutaDFP"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lstFacturaASustituirDFP"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmFacturacionDirecta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optMostrarSolo(0)"
      Tab(1).Control(1)=   "optMostrarSolo(2)"
      Tab(1).Control(2)=   "optMostrarSolo(1)"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame6"
      Tab(1).Control(5)=   "Frame1"
      Tab(1).Control(6)=   "PB"
      Tab(1).Control(7)=   "optMostrarSolo(4)"
      Tab(1).Control(8)=   "optMostrarSolo(3)"
      Tab(1).Control(9)=   "grdFactura"
      Tab(1).Control(10)=   "Label57(10)"
      Tab(1).Control(11)=   "Label57(0)"
      Tab(1).Control(12)=   "Label57(6)"
      Tab(1).Control(13)=   "Label57(7)"
      Tab(1).Control(14)=   "Label57(9)"
      Tab(1).Control(15)=   "Label57(13)"
      Tab(1).Control(16)=   "Label57(8)"
      Tab(1).Control(17)=   "Label57(14)"
      Tab(1).Control(18)=   "Label57(12)"
      Tab(1).Control(19)=   "Label57(15)"
      Tab(1).Control(20)=   "Label57(11)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmFacturacionDirecta.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdCreditosaFacturar"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame3 
         Height          =   815
         Left            =   -74880
         TabIndex        =   109
         Top             =   720
         Width           =   7540
         Begin VB.CheckBox ChkTodosCreditosaFacturar 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            ToolTipText     =   "Mostrar todos"
            Top             =   320
            Width           =   800
         End
         Begin VB.CommandButton cmdAgregarDatos 
            Caption         =   "Agregar créditos"
            Height          =   330
            Left            =   6120
            TabIndex        =   114
            ToolTipText     =   "Agregar créditos"
            Top             =   295
            Width           =   1300
         End
         Begin VB.CommandButton cmdCargarCreditosaFacturar 
            Caption         =   "Cargar créditos"
            Height          =   330
            Left            =   4800
            TabIndex        =   113
            ToolTipText     =   "Cargar créditos"
            Top             =   295
            Width           =   1250
         End
         Begin MSMask.MaskEdBox mskCFFechaFin 
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
            Left            =   3525
            TabIndex        =   112
            ToolTipText     =   "Fecha final de la búsqueda"
            Top             =   295
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCFFechaIni 
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
            Left            =   1680
            TabIndex        =   111
            ToolTipText     =   "Fecha inicial de la búsqueda"
            Top             =   295
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy "
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label26 
            Caption         =   "Selección de créditos para facturar"
            Height          =   205
            Left            =   120
            TabIndex        =   119
            Top             =   -10
            Width           =   2505
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   3000
            TabIndex        =   117
            Top             =   355
            Width           =   420
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   1095
            TabIndex        =   116
            Top             =   355
            Width           =   465
         End
      End
      Begin VB.ListBox lstFacturaASustituirDFP 
         Height          =   450
         ItemData        =   "frmFacturacionDirecta.frx":0054
         Left            =   1320
         List            =   "frmFacturacionDirecta.frx":005B
         TabIndex        =   23
         ToolTipText     =   "Facturas directas a las cuales sustituye"
         Top             =   7860
         Width           =   1695
      End
      Begin VB.CheckBox chkFacturaSustitutaDFP 
         Caption         =   "Factura sustituta"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Indica que la factura directa que se generará es sustituta de otra previamente cancelada"
         Top             =   7860
         Width           =   1035
      End
      Begin VB.Frame fraDetalle 
         Height          =   4565
         Left            =   210
         TabIndex        =   41
         Top             =   2740
         Width           =   11100
         Begin VB.CommandButton cmdCreditosaFacturar 
            Caption         =   "Seleccionar créditos para facturar"
            Height          =   375
            Left            =   8210
            TabIndex        =   118
            ToolTipText     =   "Seleccionar créditos para facturar"
            Top             =   150
            Width           =   2820
         End
         Begin VB.TextBox txtNumeroPredial 
            Height          =   315
            Left            =   8015
            MaxLength       =   150
            TabIndex        =   20
            Top             =   3390
            Width           =   3015
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   975
            Left            =   1320
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Observaciones de la factura"
            Top             =   3030
            Width           =   4935
         End
         Begin VB.ComboBox CboMotivosFactura 
            Height          =   315
            Left            =   8015
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "Motivo de la factura"
            Top             =   3030
            Width           =   3015
         End
         Begin VB.CheckBox chkMostrar 
            Caption         =   "Mostrar todos los conceptos de facturación"
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   3375
         End
         Begin VB.CommandButton cmdBorrar 
            Height          =   495
            Left            =   10515
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmFacturacionDirecta.frx":0078
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Eliminar un concepto"
            Top             =   3950
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAgregarCargSub 
            Height          =   495
            Left            =   10020
            MaskColor       =   &H00FF00FF&
            Picture         =   "frmFacturacionDirecta.frx":021A
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Seleccionar cargos de servicios subrogados"
            Top             =   3950
            UseMaskColor    =   -1  'True
            Width           =   495
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
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   42
            Text            =   "frmFacturacionDirecta.frx":055C
            Top             =   4110
            Width           =   4815
         End
         Begin VSFlex7LCtl.VSFlexGrid vsfConcepto 
            Height          =   2385
            Left            =   105
            TabIndex        =   17
            ToolTipText     =   "Lista de créditos según los filtros seleccionados"
            Top             =   570
            Width           =   10920
            _cx             =   19262
            _cy             =   4207
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmFacturacionDirecta.frx":0582
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
         Begin VB.Label lblnumpredial 
            Caption         =   "Número de predial"
            Height          =   195
            Left            =   6480
            TabIndex        =   108
            Top             =   3450
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "Observaciones"
            Height          =   255
            Left            =   105
            TabIndex        =   107
            Top             =   3100
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Motivo de la factura"
            Height          =   195
            Left            =   6480
            TabIndex        =   104
            Top             =   3105
            Width           =   1455
         End
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   58
         Top             =   1400
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo facturas sin cancelar ante el SAT"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   60
         Top             =   1940
         Width           =   3735
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo facturas pendientes de timbre fiscal"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   59
         Top             =   1670
         Width           =   3735
      End
      Begin VB.Frame fraCliente 
         Height          =   2040
         Left            =   210
         TabIndex        =   80
         Top             =   715
         Width           =   8160
         Begin VB.TextBox txtNumCliente 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1260
            TabIndex        =   0
            ToolTipText     =   "Número de cliente"
            Top             =   195
            Width           =   915
         End
         Begin VB.CheckBox chkBitExtranjero 
            Caption         =   "Extranjero"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3780
            TabIndex        =   3
            Top             =   630
            Width           =   1000
         End
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   540
            Width           =   2160
         End
         Begin VB.Label lblNumeroExterior 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   5
            ToolTipText     =   "Número exterior del cliente"
            Top             =   1245
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   100
            TabIndex        =   89
            Top             =   255
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Left            =   100
            TabIndex        =   88
            Top             =   945
            Width           =   345
         End
         Begin VB.Label lblCliente 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   2205
            TabIndex        =   1
            ToolTipText     =   "Nombre del cliente"
            Top             =   195
            Width           =   5870
         End
         Begin VB.Label lblDomicilio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            ToolTipText     =   "Calle del cliente"
            Top             =   900
            Width           =   6815
         End
         Begin VB.Label lblCiudad 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5805
            TabIndex        =   10
            ToolTipText     =   "Ciudad del cliente"
            Top             =   1605
            Width           =   2265
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   5145
            TabIndex        =   87
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   86
            Top             =   600
            Width           =   315
         End
         Begin VB.Label lblTelefono 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3360
            TabIndex        =   9
            ToolTipText     =   "Teléfono del cliente"
            Top             =   1605
            Width           =   1600
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Left            =   2685
            TabIndex        =   85
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label16 
            Caption         =   "Número exterior"
            Height          =   255
            Left            =   105
            TabIndex        =   84
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Número interior"
            Height          =   255
            Left            =   2685
            TabIndex        =   83
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label lblNumeroInterior 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3780
            TabIndex        =   6
            ToolTipText     =   "Número interior del cliente"
            Top             =   1245
            Width           =   1185
         End
         Begin VB.Label lblCP 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   8
            ToolTipText     =   "Código postal del cliente"
            Top             =   1605
            Width           =   1185
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Código postal"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   100
            TabIndex        =   82
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label lblColonia 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5805
            TabIndex        =   7
            ToolTipText     =   "Colonia del cliente"
            Top             =   1245
            Width           =   2265
         End
         Begin VB.Label Label20 
            Caption         =   "Colonia"
            Height          =   255
            Left            =   5145
            TabIndex        =   81
            Top             =   1320
            Width           =   555
         End
      End
      Begin VB.Frame fraFolioFecha 
         Height          =   1380
         Left            =   8415
         TabIndex        =   77
         Top             =   715
         Width           =   2895
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   960
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "dd/mmm/yyyy"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Folio de factura"
            Height          =   195
            Left            =   105
            TabIndex        =   79
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   105
            TabIndex        =   78
            Top             =   1020
            Width           =   450
         End
         Begin VB.Label lblFolio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   105
            TabIndex        =   11
            Top             =   480
            Width           =   2700
         End
      End
      Begin VB.Frame fraMoneda 
         Caption         =   "Factura en "
         Height          =   555
         Left            =   8415
         TabIndex        =   76
         Top             =   2200
         Width           =   2895
         Begin VB.OptionButton optPesos 
            Caption         =   "Pesos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Cantidad en pesos"
            Top             =   270
            Width           =   735
         End
         Begin VB.OptionButton optPesos 
            Caption         =   "Dólares"
            Height          =   195
            Index           =   1
            Left            =   870
            TabIndex        =   14
            ToolTipText     =   "Cantidad en pesos"
            Top             =   270
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.Label lblTipoCambio 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1755
            TabIndex        =   15
            Top             =   165
            Width           =   1050
         End
      End
      Begin VB.Frame fraTotales 
         Height          =   2730
         Left            =   5520
         TabIndex        =   68
         Top             =   7260
         Width           =   5790
         Begin VB.CheckBox chkRetencionIVA 
            Caption         =   "Retención IVA"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Retención de IVA del honorario"
            Top             =   1992
            Width           =   1890
         End
         Begin VB.CheckBox chkRetencionISR 
            Caption         =   "Retención ISR"
            Enabled         =   0   'False
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
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Retención de ISR del honorario"
            Top             =   1632
            Width           =   1890
         End
         Begin VB.ComboBox cboTarifa 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Selección de la tasa del ISR"
            Top             =   1595
            Width           =   1725
         End
         Begin VB.Label lblRetencionIVA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   106
            ToolTipText     =   "Monto del IVA retenido"
            Top             =   1955
            Width           =   1770
         End
         Begin VB.Label lblRetencionISR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   105
            ToolTipText     =   "Monto de la retención de ISR"
            Top             =   1595
            Width           =   1770
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
            Left            =   120
            TabIndex        =   27
            Top             =   912
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
            Left            =   120
            TabIndex        =   26
            Top             =   552
            Width           =   1245
         End
         Begin VB.Label Label2 
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
            Left            =   120
            TabIndex        =   28
            Top             =   1272
            Width           =   375
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   2352
            Width           =   1425
         End
         Begin VB.Label lblDescuentos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   75
            Top             =   515
            Width           =   1770
         End
         Begin VB.Label lblSubtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   73
            Top             =   875
            Width           =   1770
         End
         Begin VB.Label lblIVA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   71
            Top             =   1235
            Width           =   1770
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   70
            Top             =   2315
            Width           =   1770
         End
         Begin VB.Label lblImporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3840
            TabIndex        =   69
            Top             =   150
            Width           =   1770
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   120
            TabIndex        =   25
            Top             =   192
            Width           =   795
         End
      End
      Begin VB.CheckBox chkMovimiento 
         Caption         =   "Registrar un movimiento de crédito por cada concepto de la factura"
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   7380
         Width           =   5070
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -74865
         TabIndex        =   63
         Top             =   715
         Width           =   3900
         Begin MSMask.MaskEdBox mskFechaFin 
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
            Left            =   2565
            TabIndex        =   54
            ToolTipText     =   "Fecha final de la búsqueda"
            Top             =   195
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaIni 
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
            Left            =   705
            TabIndex        =   53
            ToolTipText     =   "Fecha inicial de la búsqueda"
            Top             =   195
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy "
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2055
            TabIndex        =   67
            Top             =   255
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   135
            TabIndex        =   66
            Top             =   255
            Width           =   465
         End
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   -70935
         TabIndex        =   52
         Top             =   715
         Width           =   7295
         Begin VB.CommandButton cmdCargar 
            Caption         =   "&Cargar datos"
            Height          =   330
            Left            =   6030
            TabIndex        =   57
            ToolTipText     =   "Cargar los datos con los parámetros "
            Top             =   195
            Width           =   1110
         End
         Begin VB.TextBox txtNumClienteBusqueda 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   645
            TabIndex        =   55
            ToolTipText     =   "Número de cliente <enter> para cargar la búsqueda"
            Top             =   195
            Width           =   735
         End
         Begin VB.Label lblNombreCliente 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1395
            TabIndex        =   56
            Top             =   188
            Width           =   4605
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   75
            TabIndex        =   61
            Top             =   255
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   "Siguiente pago"
         Top             =   9300
         Width           =   5080
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3030
            Picture         =   "frmFacturacionDirecta.frx":0656
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Cancelar factura"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2520
            Picture         =   "frmFacturacionDirecta.frx":0B48
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Guardar factura"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2040
            Picture         =   "frmFacturacionDirecta.frx":0CBA
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Ultimo pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1545
            Picture         =   "frmFacturacionDirecta.frx":0E2C
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Siguiente pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1050
            Picture         =   "frmFacturacionDirecta.frx":0F9E
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   555
            Picture         =   "frmFacturacionDirecta.frx":1110
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Anterior pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   60
            Picture         =   "frmFacturacionDirecta.frx":1282
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Primer pago"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdCFD 
            Height          =   495
            Left            =   4510
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmFacturacionDirecta.frx":13A4
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdConfirmarTimbre 
            Caption         =   "Confirmar timbre fiscal"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmFacturacionDirecta.frx":1CC2
            TabIndex        =   44
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   -71385
         TabIndex        =   40
         Top             =   9075
         Width           =   4725
         Begin VB.CommandButton cmdCancelaFacturasSAT 
            Caption         =   "Validar comprobantes pendientes de cancelación"
            Enabled         =   0   'False
            Height          =   615
            Left            =   2235
            Style           =   1  'Graphical
            TabIndex        =   103
            ToolTipText     =   "Cancelar factura(s) ante el SAT"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2415
         End
         Begin VB.CommandButton cmdConfirmarTimbreFiscal 
            Caption         =   "Confirmar timbre fiscal"
            Enabled         =   0   'False
            Height          =   615
            Left            =   60
            Picture         =   "frmFacturacionDirecta.frx":209A
            TabIndex        =   102
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.PictureBox PB 
         Height          =   135
         Left            =   -68280
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   39
         Top             =   9660
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cboUsoCFDI 
         Height          =   315
         Left            =   1315
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Uso del CFDI"
         Top             =   8580
         Width           =   3975
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo cancelación rechazada"
         Height          =   255
         Index           =   4
         Left            =   -70440
         TabIndex        =   64
         Top             =   1670
         Width           =   4335
      End
      Begin VB.OptionButton optMostrarSolo 
         Caption         =   "Mostrar sólo pendientes de autorización de cancelación"
         Height          =   255
         Index           =   3
         Left            =   -70440
         TabIndex        =   62
         Top             =   1400
         Width           =   4335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFactura 
         Height          =   6735
         Left            =   -74880
         TabIndex        =   65
         Top             =   2240
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   11880
         _Version        =   393216
         Cols            =   6
         GridColor       =   12632256
         FormatString    =   "|Fecha|Folio|Número|Cliente|Estado"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCreditosaFacturar 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   115
         ToolTipText     =   "Lista de créditos para facturar según los filtros seleccionados"
         Top             =   1630
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   14631
         _Version        =   393216
         Cols            =   9
         GridColor       =   12632256
         FormatString    =   "|Folio|Concepto de facturación|Importe|Subtotal|I.V.A.|Descuento|Clave Concepto|Número movimiento"
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
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
         Index           =   10
         Left            =   -66000
         TabIndex        =   101
         Top             =   9030
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de timbre fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   -65625
         TabIndex        =   100
         Top             =   9045
         Width           =   1890
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
         TabIndex        =   99
         Top             =   9030
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74520
         TabIndex        =   98
         Top             =   9045
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
         TabIndex        =   97
         Top             =   9810
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
         TabIndex        =   96
         Top             =   9555
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de cancelar ante el SAT"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   95
         Top             =   9300
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
         TabIndex        =   94
         Top             =   9285
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de autorización de cancelación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -74520
         TabIndex        =   93
         Top             =   9570
         Width           =   3060
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación rechazada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74520
         TabIndex        =   92
         Top             =   9825
         Width           =   1680
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Facturas timbradas ante el SAT y que no se encuentran en el SIHO"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   -74690
         TabIndex        =   91
         Top             =   320
         Width           =   4785
      End
      Begin VB.Label Label18 
         Caption         =   "Uso del CFDI"
         Height          =   255
         Left            =   240
         TabIndex        =   90
         Top             =   8635
         Width           =   1095
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "R. F. C."
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "frmFacturacionDirecta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------
' Programa para facturación directa a clientes
' Fecha de desarrollo: Agosto 23, 2006
'--------------------------------------------------------------------------------------------------------

Option Explicit

'<vsfConcepto>
Public cintColCveConcepto As Integer    'Columna donde se guardará la clave del concepto seleccionado
Public cintColIVAConcepto As Integer    'Columna donde se guardará el porcentaje que tiene el concepto de facturación
Public cintColDescripcion As Integer    'Descripción del concepto seleccionado
Public cintColCantidad As Integer       'Cantidad a facturar del concepto
Public cintColPrecioUnitario As Integer           'Precio unitario del concepto
Public cintColImporte As Integer                'Importe del concepto (Cantidad por Precio unitario)
Public cintColDescuento As Integer               'Descuento a facturar del concepto
Public cintColIVA As Integer                     'IVA a facturar del concepto
Public cintColCtaIngreso As Integer              'Cuenta contable para el ingreso
Public cintColCtaDescuento As Integer           'Cuenta contable para el descuento
Public cintColDeptoConcepto As Integer            'Cve. del departamento del concepto
Public cintColCveCargoMultiEmp As Integer
Public cintColBitExento As Integer             'Indica si el concepto es exento de IVA (1=Exento de IVA, 0=Sujeto a IVA)
Public cintColumnas As Integer                'Núm. de columnas de <vsfConcepto>

Const cstrFormato = "|||Concepto|Cantidad|Precio unitario|Importe|Descuento|IVA"

'<grdFactura>
Const cintColchrEstatus = 1
Const cintColNumPoliza = 2
Const cIntColNumCorte = 3
Const cintColFecha = 4
Const cIntColFolio = 5
Const cintColNumCliente = 6
Const cIntColRazonSocial = 7
Const cIntColRFC = 8
Const cintColTotalFactura = 9
Const cintColIVAConsulta = 10
Const cintColDescuentos = 11
Const cintColSubtotal = 12
Const cintColMoneda = 13
Const cIntColEstado = 14
Const cintColFacturo = 15
Const cintColCancelo = 16
Const cintColPFacSAT = 17 ' agregada para saber si la factura esta pendiente de facturarse ante el SAT(CGR)
Const cintColPTimbre = 18
Const cintColEstadoNuevoEsquemaCancelacion = 19
Const cintColgrdFactura = 20            'Núm. de columnas de <grdFactura>
Const cstrFormatgrdFactura = "||||Fecha|Folio|Cliente|Razón social|RFC|Total|IVA|Descuento|Subtotal|Moneda|Estado|Facturó|Canceló|NuevoEsquemaCancelacion"
Const llngColorCanceladas = &HFF&
Const llngColorActivas = &H80000012
Const llnColorPenCancelaSAT = &HC0E0FF
Const llncolorCanceladasSAT = &H80000005
Const cstrCFDI = "4.0"

Public cstrCantidad As String 'Para formatear a número
Public cstrCantidad4Decimales As String 'Para formatear a número

Const cintTipoFormato = 9               'Formato para factura directa en <TipoFormato> CC

Dim lstrConceptos As String             'Cadena con los conceptos de factura

Dim ldblImporteFactura As Double        'Total del importe de la factura
Dim ldblDescuentosFactura As Double     'Total de descuentos de la factura
Dim ldblCantidadFactura As Double       'Total de candidad menos descuento
Dim ldblIVAFactura As Double            'Total de IVA de la factura

Dim ldblDescuentoConcepto As Double     'Descuento del concepto
Dim ldblImporteConcepto As Double       'Importe del concepto
Dim ldblCantidadConcepto As Double      'Cantidad del concepto
Dim ldblIVAConcepto As Double           'IVA del concepto

Dim lblnConsulta As Boolean             'Para saber si se está consultando una factura
Dim lblnEntraCorte As Boolean           'Para saber si la factura entra o no en el corte
Dim llngNumCorte As Long                'Num. de corte en el que se está guardando
Dim llngNumFormaCredito As Long         'Num. de forma de pago CREDITO para el departamento
Dim lblnCreditoVigente As Boolean       'Indica si el cliente tiene crédito vigente o no
Dim llngNumPoliza As Long               'Num. de póliza
Dim llngNumCtaCliente As Long           'Num. de cuenta contable del cliente
Dim llngPersonaGraba As Long            'Num. de empleado que graba la factura
Dim llngNumReferencia As Long           'Nummero de referencia del cliente
Dim lstrTipoCliente As String           'Tipo de cliente
Dim strSentencia As String              'Usos varios

Dim lstrCalleNumero As String           'Para guardar en la factura
Dim lstrColonia As String               'Para guardar en la factura
Dim lstrCiudad As String                'Para guardar en la factura
Dim llngCveCiudad As Long               'Para guardar en la factura
Dim lstrEstado As String                'Para guardar en la factura
Dim lstrCodigo As String                'Para guardar en la factura

Dim llngFormato As Long                 'Num. del formato de factura para el departamento

Dim apoliza() As TipoPoliza             'Para formar la poliza de la factura
Dim vlblnMultiempresa As Boolean
Dim vllngCveProveedorM  As Long         'Para la cve del proveedor subrogado asignado en facturacion multiempresa
Dim vlstrproveedorM As String
Dim vllngFormatoaUsar As Long               'Para saber que formato se va a utilizar
Dim lngCveFormato As Long                   'Para saber el formato que se va a utilizar (relacionado con pvDocumentoDepartamento.intNumFormato)
Dim intTipoEmisionComprobante As Integer    'Variable que compara el tipo de formato y folio a utilizar (0 = Error de formato y folios incompatibles, 1 = Físicos, 2 = Digitales)
Dim strFolio As String                      'Folio de la factura
Dim strSerie As String                      'Serie de la factura
Dim strNumeroAprobacion As String           'Número de aprobación del folio
Dim strAnoAprobacion As String              'Año de aprobación del folio
Dim vgConsecutivoMuestraPvFactura As Long   'Consecutivo de la tabla PvFactura al momento de seleccionar un registro del grid en pMostrar
Dim vlstrTipoCFD As String
Dim vgintnumemprelacionada As Integer
Dim intTipoCFDFactura As Integer        'Variable que regresa el tipo de CFD de la factura(0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
Dim vlstrRazonSocial As String

Dim vlblnCancelada As Boolean               'Guarda el estado de la factura
Dim vllngSeleccionadas As Long
Dim vllngSeleccPendienteTimbre As Long
Dim blnNoMensaje As Boolean
Dim aFormasPago() As FormasPago
Dim vlblnPagoForma As Boolean           'Variable que indica si se utilizó la pantalla de formas de pago
Dim aPoliza2() As RegistroPoliza
Dim rsPVFacturaFueraCorteForma As New ADODB.Recordset 'Para guardar las formas de pago cuando la factura no entra al corte
Dim vldblTipoCambio As Double
Dim dblProporcionIVA As Double
Dim vldblTotalIVACredito As Double
Dim vlintBitSaldarCuentas As Long               'Variable que indica el valor del bit pvConceptoFacturacion.BitSaldarCuentas, que nos dice si la cuenta del ingreso se salda con la del descuento
Dim vlblnCuentaIngresoSaldada As Boolean        'Variable que indica si la cuenta del ingreso fue saldada con la cuenta del descuento
Dim vldblComisionIvaBancaria As Double          'Cantidad que corresponde al iva de la comisión bancaria aplicada a cada forma de pago

Dim vllngConsecutivoFactura As Long             'Consecutivo PvFactura

Dim vlblnEsCredito As Boolean

Dim arrTarifas() As typTarifaImpuesto
 
Dim vlnblnEmpresaPersonaFisica As Boolean
Dim vlnblnLocate As Boolean
Dim vlblnEmpresa As Boolean
Dim vldblRetServicios As Double
Public blnNoFolios As Boolean

Dim vldtmFechaFactura As Date
Dim vlStrVersionCFDI As String
Dim vlstrCodigoP As String

Const cstrgrdCreditosaFacturar = "|Fecha|Folio|Concepto de facturación|Importe|Descuento|Subtotal|IVA|Clave Concepto|Número movimiento|Total"


Private Sub pConfiguragrdCreditosaFacturar()
    Dim intcontador As Integer

    grdCreditosaFacturar.Clear
        
    grdCreditosaFacturar.Rows = 2
    grdCreditosaFacturar.Cols = 11
    grdCreditosaFacturar.FormatString = cstrgrdCreditosaFacturar

    grdCreditosaFacturar.Row = 1
    
    With grdCreditosaFacturar
        .FixedCols = 1
        .ColWidth(0) = 200
        .ColWidth(1) = 1200 'Fecha
        .ColWidth(2) = 1500 'Folio
        .ColWidth(3) = 3200 'Concepto de facturación
        .ColWidth(4) = 1000 'Importe
        .ColWidth(5) = 1000 'Descuento
        .ColWidth(6) = 1000 'Subtotal
        .ColWidth(7) = 1000 'IVA
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 1000
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        
        For intcontador = 1 To .Cols - 1
            .ColAlignmentFixed(intcontador) = flexAlignCenterCenter
        Next intcontador
    End With
End Sub


Private Sub pObtenerConsecutivo()
    Dim rs As New ADODB.Recordset

    
    vgstrParametrosSP = Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFactura")
        
       
    If rs.RecordCount <> 0 Then
        vllngConsecutivoFactura = rs!IdFactura
    Else
        '¡La información no existe!
        MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub

Private Sub pCancelaMovimiento(vllngNumFactura As Long, vlstrFolio As String, vllngCorteFactura As Long, vllngCorteActual As Long, vllngPersonaGraba As Long, Optional vlblnRefacturar As Boolean = False, Optional vllngNuevaFactura As Long = 0)
On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim lstrSentencia As String, lstrTipoDoc As String, lstrFecha As String
    Dim ldblCantidad As Double
    
    lstrSentencia = "SELECT MB.intFormaPago, MB.chrTipoMovimiento, MB.mnyCantidad, MB.mnyTipoCambio," & _
                    " FP.chrTipo, ISNULL(B.tnyNumeroBanco, MB.intCveBanco) AS IdBanco " & _
                    " FROM PvMovimientoBancoForma MB " & _
                    " INNER JOIN PvFormaPago FP ON MB.intFormaPago = FP.intFormaPago " & _
                    " LEFT  JOIN CpBanco B ON B.intNumeroCuenta = FP.intCuentaContable " & _
                    " WHERE TRIM(MB.chrTipoDocumento) = 'FA' AND MB.intNumDocumento = " & vllngNumFactura & _
                    " AND MB.intNumCorte = " & vllngCorteFactura & _
                    " AND ((mb.mnycantidad > 0 AND mb.chrtipomovimiento <> 'CBA') " & _
                           " OR (mb.mnycantidad < 0 AND mb.chrtipomovimiento = 'CBA')) "
    Set rs = frsRegresaRs(lstrSentencia)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            If rs!chrTipo <> "C" Then
                lstrFecha = fstrFechaSQL(fdtmServerFecha, fdtmServerHora) '- Fecha y hora del movimiento -'
                
                '- Revisar tipo de movimiento para determinar la cancelación -'
                If rs!chrTipoMovimiento = "CBA" Then
                    lstrTipoDoc = "CCB"                 'Comisión bancaria
                Else
                    Select Case rs!chrTipoMovimiento
                        '- Movimientos de facturación -'
                        Case "EFC": lstrTipoDoc = "CEC"   'Efectivo en factura de cliente
                        Case "TAC": lstrTipoDoc = "CJC"   'Tarjeta de crédito en factura de cliente
                        Case "TCL": lstrTipoDoc = "CTC"   'Transferencia bancaria en factura de cliente
                        Case "CHC": lstrTipoDoc = "CQC"   'Cheque en factura de cliente
                    End Select
                End If
                
                ldblCantidad = rs!MNYCantidad * (-1) 'Cantidad negativa para que se tome como abono
    
                '- Guardar información en tabla intermedia -'
                vgstrParametrosSP = vllngCorteActual & "|" & lstrFecha & "|" & rs!intFormaPago & "|" & rs!IdBanco & "|" & ldblCantidad & "|" & _
                                    IIf(rs!MNYTIPOCAMBIO = 0, 1, 0) & "|" & rs!MNYTIPOCAMBIO & "|" & lstrTipoDoc & "|" & "FA" & "|" & vllngNumFactura & "|" & _
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

Private Function fstrConceptos() As String
    Dim rs As New ADODB.Recordset

    vgstrParametrosSP = "0|1|" & IIf(Chkmostrar.Value = 1, "-1", "2")
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelConceptoFactura")
    
    fstrConceptos = ""
    Do While Not rs.EOF
        fstrConceptos = fstrConceptos & "|#" & Trim(Str(rs!smicveconcepto)) & ";" & Trim(rs!chrdescripcion)
        rs.MoveNext
    Loop
End Function

Private Sub CboMotivosFactura_Click()
    'If vlnblnEmpresaPersonaFisica Then
        If CboMotivosFactura.ListIndex = 0 Then
            chkRetencionISR.Enabled = False
            chkRetencionISR.Value = 0
            
            chkRetencionIVA.Enabled = False
            chkRetencionIVA.Value = 0
            txtNumeroPredial.Enabled = False
            txtNumeroPredial.Text = ""
            lblnumpredial.Enabled = False
            cboTarifa.Enabled = False
            
            chkMovimiento.Enabled = True
        ElseIf CboMotivosFactura.ListIndex = 3 Then
            chkRetencionISR.Enabled = False
            chkRetencionISR.Value = 0
            cboTarifa.Enabled = False
            
            txtNumeroPredial.Enabled = False
            txtNumeroPredial.Text = ""
            lblnumpredial.Enabled = False
            
                        
            chkRetencionIVA.Enabled = True
            chkRetencionIVA.Value = 0
            chkMovimiento.Value = 0
            chkMovimiento.Enabled = False
        ElseIf CboMotivosFactura.ListIndex = 2 Then
            chkRetencionISR.Enabled = True
            chkRetencionIVA.Enabled = True
            cboTarifa.Enabled = True
            txtNumeroPredial.Enabled = True
            lblnumpredial.Enabled = True
            chkMovimiento.Value = 0
            chkMovimiento.Enabled = False
        Else
            chkRetencionISR.Enabled = True
            chkRetencionIVA.Enabled = True
            cboTarifa.Enabled = True
            txtNumeroPredial.Enabled = False
            txtNumeroPredial.Text = ""
            lblnumpredial.Enabled = False
            chkMovimiento.Value = 0
            chkMovimiento.Enabled = False
        End If
    'Else
    '    chkRetencionISR.Enabled = False
    '    chkRetencionISR.Value = 0
    '
    '    If CboMotivosFactura.ListIndex = 0 Then
    '        chkRetencionIVA.Enabled = False
    '        chkRetencionIVA.Value = 0
    '        cboTarifa.Enabled = False
    '    ElseIf CboMotivosFactura.ListIndex = 1 Then
    '        chkRetencionIVA.Enabled = False
    '        chkRetencionIVA.Value = 0
    '      cboTarifa.Enabled = False
    '    End If
        
    '   chkMovimiento.Enabled = True
    'End If
    
    pAsignaTotales
End Sub

Private Sub cboTarifa_Click()
    pAsignaTotales
End Sub





Private Sub chkFacturaSustitutaDFP_Click()
Dim i As Integer

 If Not vlnblnLocate Then
    If chkFacturaSustitutaDFP.Value = 0 Then
        lstFacturaASustituirDFP.Clear
        ReDim aFoliosPrevios(0)
    Else
        If chkFacturaSustitutaDFP.Value = 1 Then
        
            frmBusquedaFacturasDirectasPrevias.vlchrtipofactura = "C"
            frmBusquedaFacturasDirectasPrevias.vlchrtipopaciente = "C"
            frmBusquedaFacturasDirectasPrevias.vlintmovpaciente = CLng(IIf(txtNumCliente = "", "0", txtNumCliente))
            frmBusquedaFacturasDirectasPrevias.Show vbModal, Me
            
            lstFacturaASustituirDFP.Clear
            For i = 0 To UBound(aFoliosPrevios())
                If aFoliosPrevios(i).chrfoliofactura <> "" Then
                    lstFacturaASustituirDFP.AddItem aFoliosPrevios(i).chrfoliofactura
                End If
            Next i
            
            If frmBusquedaFacturasDirectasPrevias.vlchrfoliofactura = "" Then
                chkFacturaSustitutaDFP.Value = 0
            End If
        End If
    End If
  End If
End Sub

Private Sub chkFacturaSustitutaDFP_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkMostrar_Click()
    lstrConceptos = fstrConceptos()
End Sub

Private Sub chkRetencionISR_Click()
    Dim rs As ADODB.Recordset

    If Not lblnConsulta Then
        Set rs = frsEjecuta_SP("-1|1", "SP_CNSELTARIFAISR")
        If rs.RecordCount = 0 Then
            cboTarifa.Clear
        End If
    
        If Trim(cboTarifa.Text) = "" Then
            pCargaTasasRetencionISR
        End If
    
        pAsignaTotales
    End If
End Sub

Private Sub chkRetencionISR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkRetencionISR.Value = 0 Then
            If chkRetencionIVA.Enabled Then
                chkRetencionIVA.SetFocus
            Else
                cboUsoCFDI.SetFocus
            End If
        Else
            cboTarifa.SetFocus
        End If
    End If
End Sub

Private Sub chkRetencionIVA_Click()
    If Not lblnConsulta Then
        pAsignaTotales
    End If
End Sub

Private Sub ChkTodosCreditosaFacturar_Click()
    If ChkTodosCreditosaFacturar.Value = 1 Then
        Label24.Enabled = False
        mskCFFechaIni.Enabled = False
        Label25.Enabled = False
        mskCFFechaFin.Enabled = False
    Else
        Label24.Enabled = True
        mskCFFechaIni.Enabled = True
        Label25.Enabled = True
        mskCFFechaFin.Enabled = True
        mskCFFechaIni.Mask = ""
        mskCFFechaIni.Text = fdtmServerFecha
        mskCFFechaIni.Mask = "##/##/####"
    
        mskCFFechaFin.Mask = ""
        mskCFFechaFin.Text = fdtmServerFecha
        mskCFFechaFin.Mask = "##/##/####"
        pEnfocaMkTexto mskCFFechaIni
    End If
End Sub

Private Sub cmdAgregarCargSub_Click()
On Error GoTo NotificaError
Dim vlintContador As Integer
Dim vlintCont As Integer

'    pLimpiavsfConcepto
'    pConfiguravsfConcepto
    frmSeleccionarCargosSub.vlintPestañaInicial = 0
    frmSeleccionarCargosSub.txtProveedorSubrCarg.Text = vlstrproveedorM
    frmSeleccionarCargosSub.txtidProveedor.Text = vllngCveProveedorM
    frmSeleccionarCargosSub.Show vbModal, Me
'    vsfConcepto.FormatString = "|||Descripcíon|Cantidad|Descuento|IVA"
'    With vsfConcepto
'        .FixedCols = 1
'        .ColWidth(0) = 100
'        .ColWidth(cintColCveConcepto) = 0
'        .ColWidth(cintColIVAConcepto) = 0
'        .ColWidth(cintColDescripcion) = 5500
'        .ColWidth(cintColPrecioUnitario) = 1500
'        .ColWidth(cintColDescuento) = 1500
'        .ColWidth(cintColIVA) = 1500
'        .ColWidth(cintColCtaIngreso) = 0
'        .ColWidth(cintColCtaDescuento) = 0
'        .ColWidth(cintColDeptoConcepto) = 0
'        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
'        .ColAlignment(cintColPrecioUnitario) = flexAlignRightCenter
'        .ColAlignment(cintColDescuento) = flexAlignRightCenter
'        .ColAlignment(cintColIVA) = flexAlignRightCenter
'        .FixedAlignment(cintColDescripcion) = flexAlignCenterCenter
'        .FixedAlignment(cintColPrecioUnitario) = flexAlignCenterCenter
'        .FixedAlignment(cintColDescuento) = flexAlignCenterCenter
'        .FixedAlignment(cintColIVA) = flexAlignCenterCenter
'    End With    'limpíar grid
    For vlintCont = 1 To frmSeleccionarCargosSub.grdCargosProveedores.Rows - 1
        If frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 0) = "*" Then
            vlintContador = 1
        End If
    Next
    If vlintContador = 1 Then
        PcargarInformacionSubrogado
    End If
    pAsignaTotales
    Unload frmSeleccionarCargosSub
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregarCargSub_Click"))
End Sub

Private Sub cmdAgregarDatos_Click()
    Dim vlintCont As Integer
    Dim vlintTotal As Integer
    Dim vldblImporte As Double
    Dim vldblDescuento As Double
    Dim vldblSubtotal As Double
    Dim vldblIVA As Double
    Dim vldbltotal As Double
    Dim vlStrMsg As String
    Dim rsDatosConcepto As New ADODB.Recordset
    Dim vlintContRep As Integer
    Dim vlblnExisteConFact As Boolean
       
    vldblImporte = 0
    vldblDescuento = 0
    vldblSubtotal = 0
    vldblIVA = 0
    vldbltotal = 0
    
    vlintTotal = 0
    vlintAgregarCreditos = 1
    ReDim allngAgregarCreditos(vlintAgregarCreditos)
    pLimpiagrdFactura
    pConfiguragrdFactura
    'vsfConcepto.Rows = vsfConcepto.Rows + 1
    For vlintCont = 1 To grdCreditosaFacturar.Rows - 1
        If Trim(grdCreditosaFacturar.TextMatrix(vlintCont, 0)) = "*" Then
             'Verifica que la fecha del crédito sea menor o igual a la fecha de la factura, sino cancela los movimientos
             If CDate(grdCreditosaFacturar.TextMatrix(vlintCont, 1)) > CDate(MskFecha.Text) Then
                pLimpiagrdFactura
                pConfiguragrdFactura
                lblnCreditosaFacturar = False
                'Debe seleccionar créditos con fecha menor o igual a la fecha de la factura.
                MsgBox "Debe seleccionar créditos con fecha menor o igual a la fecha de la factura.", vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
             End If
             'Verifica que el concepto tenga cuenta, sino se cancela los movimientos
             vgstrParametrosSP = grdCreditosaFacturar.TextMatrix(vlintCont, 8) & "|-1|-1|" & vgintClaveEmpresaContable
             Set rsDatosConcepto = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelConceptoFacturacion")
             If rsDatosConcepto!INTCUENTACONTABLE <> 0 And rsDatosConcepto!intCuentaDescuento <> 0 Then
                'Se checa que el concepto exista y si existe se le suman las cantidades
                vlblnExisteConFact = False
                For vlintContRep = 1 To vsfConcepto.Rows - 1
                    If vsfConcepto.TextMatrix(vlintContRep, cintColCveConcepto) = rsDatosConcepto!smicveconcepto Then
                        'Importe
                        vsfConcepto.TextMatrix(vlintContRep, cintColImporte) = FormatCurrency(CStr(CDbl(vsfConcepto.TextMatrix(vlintContRep, cintColImporte)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 4))))
                        'Precio unitario
                        vsfConcepto.TextMatrix(vlintContRep, cintColPrecioUnitario) = FormatCurrency(CStr(CDbl(vsfConcepto.TextMatrix(vlintContRep, cintColPrecioUnitario)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 4))))
                        'Descuento
                        vsfConcepto.TextMatrix(vlintContRep, cintColDescuento) = FormatCurrency(CStr(CDbl(vsfConcepto.TextMatrix(vlintContRep, cintColDescuento)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 5))))
                        'IVA
                        vsfConcepto.TextMatrix(vlintContRep, cintColIVA) = FormatCurrency(CStr(CDbl(vsfConcepto.TextMatrix(vlintContRep, cintColIVA)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 7))))
                        
                        'Se agrega el número de credito al arreglo
                        allngAgregarCreditos(vlintAgregarCreditos) = grdCreditosaFacturar.TextMatrix(vlintCont, 9)
                        
                        vlintTotal = vlintTotal + 1
                        
                        vldblImporte = vldblImporte + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 4))
                        vldblDescuento = vldblDescuento + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 5))
                        vldblSubtotal = vldblSubtotal + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 6))
                        vldblIVA = vldblIVA + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 7))
                        vldbltotal = vldbltotal + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 6)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 7))
                                                
                        vlintAgregarCreditos = vlintAgregarCreditos + 1
                        ReDim Preserve allngAgregarCreditos(vlintAgregarCreditos)
                        vlblnExisteConFact = True
                        
                        Exit For
                    End If
                Next
                'Si no existe de agrega en un nuevo renglon
                If vlblnExisteConFact = False Then
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCveConcepto) = rsDatosConcepto!smicveconcepto
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColIVAConcepto) = 0 'rsConcepto!smyIVA
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCtaIngreso) = rsDatosConcepto!INTCUENTACONTABLE
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCtaDescuento) = rsDatosConcepto!intCuentaDescuento
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDeptoConcepto) = rsDatosConcepto!SMIDEPARTAMENTO
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColBitExento) = 0 'rsConcepto!bitExentoIva
                    
                    'Concepto
                    'vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCveConcepto) = grdCreditosaFacturar.TextMatrix(vlintCont, 7)
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescripcion) = grdCreditosaFacturar.TextMatrix(vlintCont, 3)
                    'Importe
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColImporte) = grdCreditosaFacturar.TextMatrix(vlintCont, 4)
                    'Precio unitario
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColPrecioUnitario) = grdCreditosaFacturar.TextMatrix(vlintCont, 4)
                    'Descuento
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescuento) = grdCreditosaFacturar.TextMatrix(vlintCont, 5)
                    'IVA
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColIVA) = grdCreditosaFacturar.TextMatrix(vlintCont, 7)
                    'Cantidad
                    vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCantidad) = "1.00"
                                
                    'Se agrega el número de credito al arreglo
                    allngAgregarCreditos(vlintAgregarCreditos) = grdCreditosaFacturar.TextMatrix(vlintCont, 9)
                    
                    vlintTotal = vlintTotal + 1
                    
                    vldblImporte = vldblImporte + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 4))
                    vldblDescuento = vldblDescuento + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 5))
                    vldblSubtotal = vldblSubtotal + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 6))
                    vldblIVA = vldblIVA + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 7))
                    vldbltotal = vldbltotal + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 6)) + CDbl(grdCreditosaFacturar.TextMatrix(vlintCont, 7))
                    
                    'If vlintCont < grdCreditosaFacturar.Rows - 1 Then
                        vsfConcepto.Rows = vsfConcepto.Rows + 1
                        vlintAgregarCreditos = vlintAgregarCreditos + 1
                        ReDim Preserve allngAgregarCreditos(vlintAgregarCreditos)
                    'End If
                End If
             Else
                'blnerror = True
                pLimpiagrdFactura
                pConfiguragrdFactura
                lblnCreditosaFacturar = False
                'No existe cuenta contable para el concepto de facturación:
                MsgBox SIHOMsg(907) & vsfConcepto.ComboItem(vsfConcepto.ComboIndex), vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
             End If
        End If
    Next
    
    If vlintTotal = 0 Then
        If MsgBox(SIHOMsg(819) & "¿Desea continuar?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            vlintAgregarCreditos = 0
            SSTFactura.Tab = 0
            cmdAgregarDatos.Enabled = False
            lblnCreditosaFacturar = False
        End If
    Else
        'Se desactivan los controles para que no se puedan agregar más datos
        Chkmostrar.Value = 0
        Chkmostrar.Enabled = False
        chkFacturaSustitutaDFP.Value = 0
        chkFacturaSustitutaDFP.Enabled = False
        lstFacturaASustituirDFP.Text = ""
        lstFacturaASustituirDFP.Enabled = False
        cmdCreditosaFacturar.Enabled = False
        vsfConcepto.Enabled = False
        MskFecha.Enabled = False
        SSTFactura.Tab = 0
        'Se agregan las sumas de los datos
        'lblImporte.Caption = FormatCurrency(vldblImporte, 4)
        'lblDescuentos.Caption = FormatCurrency(vldblDescuento, 4)
        'lblSubtotal.Caption = FormatCurrency(vldblSubtotal, 4)
        'lblIVA.Caption = FormatCurrency(vldblIVA, 4)
        lblTotal.Caption = FormatCurrency(vldbltotal, 4)
        
        ldblImporteFactura = vldblImporte
        ldblDescuentosFactura = vldblDescuento
        'lblSubtotal.Caption = FormatCurrency(vldblSubtotal, 4)
        ldblIVAFactura = vldblIVA
        
        lblImporte.Caption = FormatCurrency(ldblImporteFactura, 4)
        lblDescuentos.Caption = FormatCurrency(ldblDescuentosFactura, 4)
        lblSubtotal.Caption = FormatCurrency(ldblImporteFactura - ldblDescuentosFactura, 4)
        lblIVA.Caption = FormatCurrency(ldblIVAFactura, 4)
        
        lblnCreditosaFacturar = True
        
        cmdSave.SetFocus
    End If
    
    'For vlintCont = 1 To vlintAgregarCreditos
    '    vlStrMsg = vlStrMsg & allngAgregarCreditos(vlintCont) & " - "
    'Next
    'MsgBox vlStrMsg, vbInformation + vbOKOnly, "Mensaje"
    'If vsfConcepto.Rows = 3 Then
    '    vsfConcepto.Rows = vsfConcepto.Rows - 1
    'End If
End Sub

Private Sub cmdBack_Click()
    If grdFactura.Row > 1 Then
        grdFactura.Row = grdFactura.Row - 1
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, IIf(grdFactura.TextMatrix(grdFactura.Row, cintColPTimbre) = 1, 0, IIf(Trim(grdFactura.TextMatrix(grdFactura.Row, cintColchrEstatus)) = "C", 0, 1))
End Sub

Private Sub cmdBorrar_Click()
    ldblDescuentoConcepto = Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColDescuento), cstrCantidad))
    ldblCantidadConcepto = Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColPrecioUnitario), cstrCantidad))
    ldblIVAConcepto = Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColIVA), cstrCantidad))
    ldblImporteConcepto = Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColImporte), cstrCantidad))

    vsfConcepto.RemoveItem vsfConcepto.Row
    
    ldblDescuentosFactura = ldblDescuentosFactura - ldblDescuentoConcepto
    ldblCantidadFactura = ldblCantidadFactura - ldblCantidadConcepto
    ldblIVAFactura = ldblIVAFactura - ldblIVAConcepto
    ldblImporteFactura = ldblImporteFactura - ldblImporteConcepto
        
    If Not vlblnMultiempresa Then
        vsfConcepto_RowColChange
    Else
        CmdBorrar.Enabled = False
    End If

    pAsignaTotales
End Sub

Private Sub cmdCancelaFacturasSAT_Click()
    'Cancelacion masiva de facturas ante el SAT, cancelacion del XML
    Dim vlLngCantidadFacturas As Long
    Dim vlintFacturasCanceladas As Long
    Dim vlLngCont As Long
    Dim vllngPersonaGraba As Long

    On Error GoTo NotificaError
    '|  Los comprobantes seleccionados serán validados nuevamente ante el SAT.
    '|  ¿Desea continuar?
    If MsgBox(SIHOMsg(1249), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
        'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
        With grdFactura
            vlLngCantidadFacturas = 0
            vlintFacturasCanceladas = 0
            For vlLngCont = 1 To .Rows - 1
                '|  Parámetros:  Columna fixed         Estado nuevo esquema cancelación                                          Estado nuevo esquema cancelación
                If .TextMatrix(vlLngCont, 0) = "*" And (.TextMatrix(vlLngCont, cintColEstadoNuevoEsquemaCancelacion) <> "NP" And .TextMatrix(vlLngCont, cintColEstadoNuevoEsquemaCancelacion) <> "CR") Then
                    vlLngCantidadFacturas = vlLngCantidadFacturas + 1
                    '|  Parámetros:         (     Folio factura      )
                    If fblnFacturaCancelable(.TextMatrix(vlLngCont, cIntColFolio)) Then
                        '|  Parámetros:                     Folio factura,                                Estado nuevo esquema,                    Persona graba, Honorarios, Nombre forma, Muestra mensaje de cancelación satisfactoria
                        pCancelaCFDiFacturaSiHO .TextMatrix(vlLngCont, cIntColFolio), .TextMatrix(vlLngCont, cintColEstadoNuevoEsquemaCancelacion), vllngPersonaGraba, 0, Me.Name, False
                        vlintFacturasCanceladas = vlintFacturasCanceladas + 1
                    End If
                End If
            Next vlLngCont
        End With
        If vlLngCantidadFacturas = vlintFacturasCanceladas Then
            '|  La operación se realizó satisfactoriamente.
            MsgBox SIHOMsg(420), vbInformation + vbOKOnly, "Mensaje"
        End If
        '|  Refresca la información
        cmdCargar_Click
        grdFactura.SetFocus
    End If
    Exit Sub
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelaFacturasSAT_Click"))
    Unload Me

End Sub



Public Sub pCancelaSiHO(pintRow As Integer, plngPersonaGraba As Long)
    Dim vllngNumCorteFactura As Long
    Dim vllngNumCorte As Long
    
    Dim vlstrsql As String
    Dim rsCCDatosPolizaCF As New ADODB.Recordset
    
    Dim vlstrsqlCredFact As String
    Dim rsCredFact As New ADODB.Recordset
        
    EntornoSIHO.ConeccionSIHO.BeginTrans
    '1.- Cancela la factura y registrar en documentos cancelados
    vgstrParametrosSP = grdFactura.TextMatrix(pintRow, cIntColFolio) & "|" & Str(vgintNumeroDepartamento) & "|" & Str(plngPersonaGraba)
    frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaFactura", True

    '2.- Cancelar los créditos
    vgstrParametrosSP = grdFactura.TextMatrix(pintRow, cIntColFolio) & "|" & "FA"
    frsEjecuta_SP vgstrParametrosSP, "Sp_CcUpdCancelaCredito", True
    
    vgstrParametrosSP = grdFactura.TextMatrix(pintRow, cIntColFolio) & "|" & vgintClaveEmpresaContable & "|" & grdFactura.TextMatrix(pintRow, cintColNumCliente)
    frsEjecuta_SP vgstrParametrosSP, "SP_CCDELMULTEMPCARGOS", True
    
    vllngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")

    '3.- Si la factura no entró al corte, cancelar la póliza
    If Val(grdFactura.TextMatrix(pintRow, cintColNumPoliza)) <> 0 Then
        ' Cancelar la póliza
        pCancelarPoliza CLng(grdFactura.TextMatrix(pintRow, cintColNumPoliza)), "CANCELACION DE FACTURA " & grdFactura.TextMatrix(pintRow, cIntColFolio) & " (REUTILIZAR POLIZA) "
        
        ' Busca el corte de la factura y corte actual para cancelar el movimiento al libro de banco
        vllngNumCorteFactura = frsRegresaRs("SELECT intNumCorte FROM pvMovimientoFueraCorte WHERE intnumpoliza = " & CLng(grdFactura.TextMatrix(pintRow, cintColNumPoliza))).Fields(0)
        If vllngNumCorte = 0 Then
            vllngNumCorte = vllngNumCorteFactura
        End If
        pCancelaMovimiento vgConsecutivoMuestraPvFactura, grdFactura.TextMatrix(pintRow, cIntColFolio), vllngNumCorteFactura, vllngNumCorte, plngPersonaGraba
        
        ' Liberar para que se realicen cierres
        pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
    Else
        ' Cancelar el documento en el corte
        vgstrParametrosSP = grdFactura.TextMatrix(pintRow, cIntColFolio) & "|" & plngPersonaGraba & "|" & "FA" & "|" & grdFactura.TextMatrix(pintRow, cIntColNumCorte) & "|" & Str(vllngNumCorte)
        frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaDoctoCorte", True
    
        pCancelaMovimiento vgConsecutivoMuestraPvFactura, grdFactura.TextMatrix(pintRow, cIntColFolio), grdFactura.TextMatrix(pintRow, cIntColNumCorte), vllngNumCorte, plngPersonaGraba
    
        ' Liberar el corte
        pLiberaCorte vllngNumCorte
    End If
    
    '4.- Si se agregaron creditos a facturar, se cancela la poliza
    vlstrsql = "SELECT DISTINCT(INTNUMEROPOLIZACANCELACION) NUMEROPOLIZA FROM CCMOVIMIENTOCREDITOPOLIZA WHERE CHRFOLIOFACTURA = '" + Trim(grdFactura.TextMatrix(pintRow, cIntColFolio)) + "'"
    Set rsCCDatosPolizaCF = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rsCCDatosPolizaCF.RecordCount > 0 Then
         vgstrParametrosSP = grdFactura.TextMatrix(pintRow, cIntColFolio) & "|" & rsCCDatosPolizaCF!NumeroPoliza
         frsEjecuta_SP vgstrParametrosSP, "sp_CCUpdCancelaPolCredaFact", True
    End If
    
    pGuardarLogTransaccion Me.Name, EnmGrabar, plngPersonaGraba, "CANCELACION DE FACTURA DIRECTA", grdFactura.TextMatrix(pintRow, cIntColFolio)
    
    '<Si es comprobante fiscal digital, cancelar el CFD>
    ''vgstrParametrosSP = vgConsecutivoMuestraPvFactura & "|" & "FA" & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora)
    vgstrParametrosSP = vgConsecutivoMuestraPvFactura & "|" & "FA" & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|'" & vgMotivoCancelacion & "'"
    frsEjecuta_SP vgstrParametrosSP, "SP_GNUPDCANCELACOMPROBANTEFIS", True
    
    pMensajeCanelacionCFDi vgConsecutivoMuestraPvFactura, "FA"
    
    'Actualiza PDF al cancelar facturas
    pObtenerConsecutivo
    If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, False, True, -1) Then
            On Error Resume Next
    End If
    
    EntornoSIHO.ConeccionSIHO.CommitTrans

End Sub


Private Sub cmdCargar_Click()
    Dim rs As New ADODB.Recordset
    Dim lngColor As Long
    Dim lngcolorSub As Long
    Dim lngAux As Long
    Dim lngAncho As Long
       
    If fblnDatosBus() Then
        pLimpiagrdFactura
        pConfiguragrdFactura
        vllngSeleccionadas = 0
        vllngSeleccPendienteTimbre = 0
        
        lngAncho = 1100
        lngAux = 0
        
        vgstrParametrosSP = _
        "-1" & _
        "|" & IIf(Not optMostrarSolo(0), fstrFechaSQL("01/01/2010"), fstrFechaSQL(mskFechaIni.Text)) & _
        "|" & IIf(Not optMostrarSolo(0), fstrFechaSQL(fdtmServerFecha), fstrFechaSQL(mskFechaFin.Text)) & _
        "|1|" & _
        IIf(Not optMostrarSolo(0), "-1", IIf(Val(txtNumClienteBusqueda.Text) = 0, "-1", txtNumClienteBusqueda.Text)) & _
        "|" & "C" & _
        "|" & "-1" & _
        "|" & CStr(vgintNumeroDepartamento) & _
        "|" & vgintClaveEmpresaContable & _
        "|" & IIf(optMostrarSolo(2).Value, 1, 0) & _
        "|" & IIf(optMostrarSolo(1).Value, 1, 0) & _
        "|" & IIf(optMostrarSolo(3).Value, 1, 0) & "|" & IIf(optMostrarSolo(4).Value, 1, 0)
                      
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFacturaFiltro_NE")
        If rs.RecordCount <> 0 Then
            With grdFactura
                .Visible = False
                Do While Not rs.EOF
                    .Row = .Rows - 1
                    If rs!chrestatus = "C" Then
                        lngColor = llngColorCanceladas
                        lngcolorSub = &HFFFFFF '| Blanco
                        .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                    Else
                        '
                        Select Case rs!PendienteCancelarSAT_NE
                            Case "PA" '| Pendiente de autorización
                                lngColor = &HFFFFFF '| Blanco
                                lngcolorSub = &H80FF&  '| Naranja fuerte
                                .TextMatrix(.Row, 0) = "*"
                                vllngSeleccionadas = vllngSeleccionadas + 1
                                .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                            Case "CR" '| Cancelación rechazada
                                lngColor = &HFFFFFF '| Blanco
                                lngcolorSub = &HFF&    '| Rojo
                                .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                            Case "NP" '| No se encuentra pendiente de cancelación
                                lngColor = &H0&     '| Negro
                                lngcolorSub = &HFFFFFF '| Blanco
                                .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                        End Select
                    End If
                    
                    If rs!PendienteCancelarSAT_NE = "PC" Then '| Pendiente de cancelación
                       .TextMatrix(.Row, 0) = "*"
                       vllngSeleccionadas = vllngSeleccionadas + 1
                       lngcolorSub = llnColorPenCancelaSAT
                       lngColor = &HFF&  '| Rojo
                       .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                    Else
                       If rs!PendienteTimbreFiscal = 1 Then
                          lngColor = &H0&     '| Negro
                          lngcolorSub = &H80FFFF    '| Amarillo
                          .TextMatrix(.Row, 0) = "*"
                          vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                       Else
                          'lngcolorSub = llncolorCanceladasSAT
                       End If
                    End If
                    
                    .TextMatrix(.Row, cintColPFacSAT) = rs!PendienteCancelarSat
                    .TextMatrix(.Row, cintColPTimbre) = rs!PendienteTimbreFiscal
                    .Col = cIntColNumCorte
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColNumCorte) = rs!NumCorte
                    .Col = cintColNumPoliza
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColNumPoliza) = rs!numpoliza
                    .Col = cintColchrEstatus
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColchrEstatus) = rs!chrestatus
                    .Col = cintColFecha
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColFecha) = Format(rs!fecha, "dd/mmm/yyyy")
                    .Col = cIntColFolio
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColFolio) = rs!Folio
                    'ajustar la columna de los folios-------------
                    PB.Font = .CellFontName
                    PB.FontSize = .CellFontSize
                    lngAux = PB.TextWidth(.TextMatrix(.Row, cIntColFolio))
                    If lngAux > lngAncho Then
                       lngAncho = lngAux
                    End If
                    '-----------------------------------------------------------------------------------------------
                    .Col = cintColNumCliente
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColNumCliente) = rs!NumCliente
                    .Col = cIntColRazonSocial
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColRazonSocial) = rs!RazonSocial
                    .Col = cIntColRFC
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColRFC) = rs!RFC
                    .Col = cintColTotalFactura
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColTotalFactura) = FormatCurrency(rs!TotalFactura, 4)
                    .Col = cintColIVAConsulta
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColIVAConsulta) = FormatCurrency(rs!IVA, 4)
                    .Col = cintColDescuentos
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColDescuentos) = FormatCurrency(rs!Descuento, 4)
                    .Col = cintColSubtotal
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColSubtotal) = FormatCurrency(rs!Subtotal, 4)
                    .Col = cintColMoneda
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColMoneda) = rs!Moneda
                    .Col = cIntColEstado
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColEstado) = IIf(IsNull(rs!Estado), "", rs!Estado)
                    .Col = cintColFacturo
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColFacturo) = rs!PersonaFacturo
                    .Col = cintColCancelo
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColCancelo) = IIf(IsNull(rs!PersonaCancelo), "", rs!PersonaCancelo) 'Caso 20313
                    .RowData(.Row) = rs!IdFactura
                    .Rows = .Rows + 1
                    rs.MoveNext
                Loop
                If lngAncho > 1100 Then .ColWidth(cIntColFolio) = lngAncho + 100
                .Rows = .Rows - 1
                .Visible = True
            End With
            grdFactura.Col = cintColFecha
            grdFactura.Row = 1
            grdFactura.SetFocus
        Else
            'No existe información con esos parámetros.
          If Not blnNoMensaje Then MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
          If optMostrarSolo(0).Value Then
            mskFechaIni.SetFocus
          End If
        End If
        Me.cmdCancelaFacturasSAT.Enabled = vllngSeleccionadas > 0
        Me.cmdConfirmartimbrefiscal.Enabled = vllngSeleccPendienteTimbre > 0
      
    End If
End Sub

Private Function fblnDatosBus() As Boolean
    fblnDatosBus = True
    If optMostrarSolo(2) Then Exit Function
        
    If Not IsDate(mskFechaIni.Text) Then
        fblnDatosBus = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaIni.SetFocus
    End If
    
    If fblnDatosBus And Not IsDate(mskFechaFin.Text) Then
        fblnDatosBus = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaFin.SetFocus
    End If
    
    If fblnDatosBus Then
        If CDate(mskFechaFin.Text) < CDate(mskFechaIni.Text) Then
            fblnDatosBus = False
            '¡Rango incorrecto!
            MsgBox SIHOMsg(26), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaIni.SetFocus
        End If
    End If
End Function

Private Sub cmdCargarCreditosaFacturar_Click()
'Se agregó para mostrar los créditos a facturar que no estan en una factura
On Error GoTo NotificaError
    Dim vlstrParametros As String
    Dim rsCreditoafacturar As New ADODB.Recordset

    'grdCreditosaFacturar.Rows = 2
    'grdCreditosaFacturar.Cols = 9
    
    If mskCFFechaIni.Text = "  /  /    " Or mskCFFechaFin.Text = "  /  /    " Then
        'Fecha no valida.
        MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
        pEnfocaMkTexto mskCFFechaIni
        Exit Sub
    End If
    
    pConfiguragrdCreditosaFacturar

    'Obtiene los artículos de la lista
    vgstrParametrosSP = fstrFechaSQL(mskCFFechaIni) & "|" & fstrFechaSQL(mskCFFechaFin) & "|" & ChkTodosCreditosaFacturar.Value & "|" & CInt(Trim(txtNumCliente.Text))
    Set rsCreditoafacturar = frsEjecuta_SP(vgstrParametrosSP, "SP_CCSELCREDITOSAFACTURAR")
    
    If rsCreditoafacturar.RecordCount = 0 Then
        cmdAgregarDatos.Enabled = False
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    Do While Not rsCreditoafacturar.EOF
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 1) = Format(rsCreditoafacturar!DTMFECHAMOVIMIENTO, "dd/mmm/yyyy")
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 2) = rsCreditoafacturar!chrfolioreferencia
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 3) = rsCreditoafacturar!DESCCONCEPTOFACT
            'Importe
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 4) = FormatCurrency(rsCreditoafacturar!MNYSUBTOTAL + rsCreditoafacturar!MNYDESCUENTO)
            'Descuento
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 5) = FormatCurrency(rsCreditoafacturar!MNYDESCUENTO)
            'Subtotal
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 6) = FormatCurrency(rsCreditoafacturar!MNYSUBTOTAL)
            'IVA
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 7) = FormatCurrency(rsCreditoafacturar!MNYIVA)
            
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 8) = rsCreditoafacturar!INTCVECONCEPTOFACT
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 9) = rsCreditoafacturar!intNumMovimiento
            
            'Total
            grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Rows - 1, 10) = FormatCurrency(rsCreditoafacturar!MNYSUBTOTAL + rsCreditoafacturar!MNYIVA)
            
            grdCreditosaFacturar.Rows = grdCreditosaFacturar.Rows + 1
            rsCreditoafacturar.MoveNext
    Loop
    
    grdCreditosaFacturar.Rows = grdCreditosaFacturar.Rows - 1
    cmdAgregarDatos.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCargarCreditosaFacturar_Click"))
End Sub

Private Sub cmdCFD_Click()
On Error GoTo NotificaError

    If vlstrTipoCFD = "CFD" Then
        frmComprobanteFiscalDigital.lngComprobante = vgConsecutivoMuestraPvFactura
        frmComprobanteFiscalDigital.strTipoComprobante = "FA"
        frmComprobanteFiscalDigital.blnCancelado = vlblnCancelada
        frmComprobanteFiscalDigital.Show vbModal, Me
    ElseIf vlstrTipoCFD = "CFDi" Then
        frmComprobanteFiscalDigitalInternet.lngComprobante = vgConsecutivoMuestraPvFactura
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = "FA"
        frmComprobanteFiscalDigitalInternet.blnCancelado = vlblnCancelada
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCFD_Click"))
    Unload Me
End Sub
Private Sub cmdConfirmartimbre_Click()
Dim vlLngCont As Long
Dim vllngPersonaGraba As Long
Dim vlngReg As Long

On Error GoTo NotificaError

blnNOMensajeErrorPAC = False 'de inicio siempre a False

'Los comprobantes seleccionados se encuentran pendientes de timbre fiscal. ¿Desea confirmar el timbre fiscal?
If MsgBox(Replace(SIHOMsg(1310), "Los comprobantes seleccionados se encuentran pendientes de timbre fiscal. ", ""), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   
   vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
   If vllngPersonaGraba = 0 Then Exit Sub
        
       pgbBarraCFD.Value = 70
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbHourglass
       lblTextoBarraCFD.Caption = "Confirmando timbre fiscal, por favor espere..."
       freBarraCFD.Visible = True
       freBarraCFD.Refresh
       pLogTimbrado 2
       blnNOMensajeErrorPAC = True
       EntornoSIHO.ConeccionSIHO.BeginTrans
       vlngReg = flngRegistroFolio("FA", vgConsecutivoMuestraPvFactura)
       If Not fblnGeneraComprobanteDigital(vgConsecutivoMuestraPvFactura, "FA", 1, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
          On Error Resume Next
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar
             'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
              MsgBox Replace(SIHOMsg(1314), " <FOLIO>", ""), vbInformation + vbOKOnly, "Mensaje"
             'la factura se queda igual, no se hace nada
          ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
              'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
              MsgBox Replace(SIHOMsg(1313), " <FOLIO>", ""), vbExclamation + vbOKOnly, "Mensaje"
              'Aqui se debe de cancelar la factura
              pCancelarFacturaDirecta Me, CLng(grdFactura.TextMatrix(grdFactura.Row, cintColNumCliente)), Trim(lblFolio.Caption), Not (Val(grdFactura.TextMatrix(grdFactura.Row, cintColNumPoliza)) <> 0), vllngPersonaGraba, MskFecha.Text, CLng(grdFactura.TextMatrix(grdFactura.Row, cIntColNumCorte)), CLng(grdFactura.TextMatrix(grdFactura.Row, cintColNumPoliza))
              pEliminaPendientesTimbre vgConsecutivoMuestraPvFactura, "FA"
          End If
       Else
          'Se guarda el LOG
           Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura" & lblFolio.Caption)
          'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
           pEliminaPendientesTimbre vgConsecutivoMuestraPvFactura, "FA"
          'Commit
           EntornoSIHO.ConeccionSIHO.CommitTrans
          'Timbre fiscal de factura <FOLIO>: Confirmado.
           MsgBox Replace(SIHOMsg(1315), "<FOLIO> ", ""), vbInformation + vbOKOnly, "Mensaje"
           'secarga de nuevo la factura
           'pConsultaFacturas Trim(lblFolio.Caption), 0
       End If
                  
       'Barra de progreso CFD
       pgbBarraCFD.Value = 100
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbDefault
       freBarraCFD.Visible = False
       blnNOMensajeErrorPAC = False
       pLogTimbrado 1
           
       If vgIntBanderaTImbradoPendiente = 0 Or vgIntBanderaTImbradoPendiente = 2 Then
           pReinicia
           txtNumCliente.SetFocus
       End If
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub
Private Sub cmdconfirmartimbrefiscal_Click()
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
      With grdFactura
           
           For vlLngCont = 1 To .Rows - 1
               If .TextMatrix(vlLngCont, 0) = "*" And .TextMatrix(vlLngCont, cintColPTimbre) = 1 Then
                  pgbBarraCFD.Value = 70
                  freBarraCFD.Top = 3200
                  Screen.MousePointer = vbHourglass
                  lblTextoBarraCFD.Caption = "Confirmando timbre fiscal, por favor espere..."
                  freBarraCFD.Visible = True
                  freBarraCFD.Refresh
                  pLogTimbrado 2
                  blnNOMensajeErrorPAC = True
                  EntornoSIHO.ConeccionSIHO.BeginTrans
                  vlngReg = flngRegistroFolio("FA", .RowData(vlLngCont))
                  If Not fblnGeneraComprobanteDigital(.RowData(vlLngCont), "FA", 1, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
                      On Error Resume Next
                       
                       EntornoSIHO.ConeccionSIHO.RollbackTrans
                       
                                           
                       If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar/o no se alcanzo a llegar al timbre
                          'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                          MsgBox Replace(SIHOMsg(1314), "<FOLIO>", Trim(.TextMatrix(vlLngCont, cIntColFolio))), vbInformation + vbOKOnly, "Mensaje"
                          'la factura se queda igual, no se hace nada
                       ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
                          'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
                          MsgBox Replace(SIHOMsg(1313), "<FOLIO>", Trim(.TextMatrix(vlLngCont, cIntColFolio))), vbExclamation + vbOKOnly, "Mensaje"
                          'Aqui se debe de cancelar la factura
                          pCancelarFactura Trim(.TextMatrix(vlLngCont, 5)), vllngPersonaGraba, Me.Name
                       End If
                  Else
                      'Se guarda el LOG
                       Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & .TextMatrix(vlLngCont, 1))
                       'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
                       pEliminaPendientesTimbre .RowData(vlLngCont), "FA"
                       'Commit
                       EntornoSIHO.ConeccionSIHO.CommitTrans
                      'Timbre fiscal de factura <FOLIO>: Confirmado.
                       MsgBox Replace(SIHOMsg(1315), "<FOLIO>", Trim(.TextMatrix(vlLngCont, cIntColFolio))), vbInformation + vbOKOnly, "Mensaje"
                       
                  End If
                  
                  'Barra de progreso CFD
                   pgbBarraCFD.Value = 100
                   freBarraCFD.Top = 3200
                   Screen.MousePointer = vbDefault
                   freBarraCFD.Visible = False
                    pLogTimbrado 1
               End If
           Next vlLngCont
      End With
      
      blnNOMensajeErrorPAC = False
      cmdCargar_Click
      grdFactura.SetFocus
      
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub

Private Sub cmdCreditosaFacturar_Click()
    SSTFactura.Tab = 2
    
    mskCFFechaIni.Mask = ""
    mskCFFechaIni.Text = fdtmServerFecha
    mskCFFechaIni.Mask = "##/##/####"
    
    mskCFFechaFin.Mask = ""
    mskCFFechaFin.Text = fdtmServerFecha
    mskCFFechaFin.Mask = "##/##/####"
    
    ChkTodosCreditosaFacturar.Value = 1
    ChkTodosCreditosaFacturar.SetFocus
    
    pConfiguragrdCreditosaFacturar
    
    cmdAgregarDatos.Enabled = False
    Call cmdCargarCreditosaFacturar_Click
End Sub

Private Sub cmdDelete_Click()
    Dim intError As Integer         'Error en transacción
    
    If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 2260, 609), "E") Then Exit Sub
    
    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If llngPersonaGraba = 0 Then Exit Sub
           frmMotivosCancelacion.blnActivaUUID = True
           frmMotivosCancelacion.intNumCliente = CInt(Trim(txtNumCliente))
           frmMotivosCancelacion.strdtmFechaHora = vldtmFechaFactura
           frmMotivosCancelacion.Show vbModal, Me
           If vgMotivoCancelacion = "" Then Exit Sub
    
    If Not fblnCancelaCFDi(vgConsecutivoMuestraPvFactura, "FA") Then
  '     EntornoSIHO.ConeccionSIHO.RollbackTrans
        If vlstrMensajeErrorCancelacionCFDi <> "" Then MsgBox vlstrMensajeErrorCancelacionCFDi, vbOKOnly + vlintTipoMensajeErrorCancelacionCFDi, "Mensaje"
        pMuestra
        pReinicia
        txtNumCliente.SetFocus
       Exit Sub
    End If
        
    
    intError = fintErrorCancelar()
    
    If intError = 0 Then
        
        pCancelaSiHO grdFactura.Row, llngPersonaGraba
        pReinicia
        txtNumCliente.SetFocus
        
    Else
        
        If intError = 964 Then
           MsgBox Replace(SIHOMsg(intError), "!", "") & " o " & Replace(SIHOMsg(1059), "E", "e"), vbOKOnly + vbExclamation, "Mensaje"
        ElseIf intError = 783 Then
           'No se encontró la forma de pago equivalente para el departamento seleccionado. --> No se encontró la forma de pago de la factura.
           MsgBox Replace(SIHOMsg(intError), "equivalente para el departamento seleccionado.", "de la factura."), vbOKOnly + vbExclamation, "Mensaje"
        Else
            MsgBox SIHOMsg(intError), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
End Sub


Private Function fintErrorCancelar() As Integer
    Dim rs As New ADODB.Recordset
    Dim rsPagos As New ADODB.Recordset
    Dim rsFormaPago As New ADODB.Recordset
    
    
    fintErrorCancelar = 0
    
    'Que la factura aún esté activa
    Set rs = frsEjecuta_SP(Trim(lblFolio.Caption), "sp_PvSelFactura")
    If rs.RecordCount <> 0 Then
        If rs!chrestatus = "C" Then
            'La información ha cambiado, consulte de nuevo.
            fintErrorCancelar = 381
            Exit Function
        End If
    End If
    rs.Close
    
    'Que el o los créditos de la factura no tengan pagos registrados
    vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha) & _
                        "|" & fstrFechaSQL(fdtmServerFecha) & _
                        "|" & txtNumCliente.Text & _
                        "|" & "0" & _
                        "|" & "FA" & _
                        "|" & "0" & _
                        "|" & Trim(lblFolio.Caption) & _
                        "|" & "0" & _
                        "|" & "0" & _
                        "|" & "*" & _
                        "|" & "0"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelCredito")
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF And fintErrorCancelar = 0
            If IsDate(rs!fechaEnvio) Then
                'No se puede cancelar el documento, los créditos fueron incluídos en un paquete de cobranza.
                fintErrorCancelar = 718
            Else
                vgstrParametrosSP = Str(rs!Movimiento) & "|" & "0" & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & "P"
                Set rsPagos = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelPagosCredito")
                If rsPagos.RecordCount <> 0 Then
                    'No se puede cancelar el documento  el crédito tiene pagos registrados.
                    fintErrorCancelar = 368
                Else
                    '-------------------------------------------------------------------------'
                    '- CASO 7249: Revisar que la factura no tenga pagos por Notas de Crédito -'
                    '-------------------------------------------------------------------------'
                    strSentencia = " SELECT mnyCantidadPagada FROM CcMovimientoCredito " & _
                                   " WHERE intNumMovimiento = " & Str(rs!Movimiento)
                    Set rsPagos = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rsPagos.Fields(0) > 0 Then
                        'No se puede cancelar el documento el crédito tiene pagos registrados.
                        fintErrorCancelar = 368
                    End If
                    rsPagos.Close
                    '-------------------------------------------------------------------------'
                End If
            End If
            rs.MoveNext
        Loop
        If fintErrorCancelar <> 0 Then
            Exit Function
        End If
    Else
        'revisar si la factura fué pagada al contado
        '-------------------------------------------
        Set rsFormaPago = frsEjecuta_SP(Trim(Me.lblFolio.Caption), "sp_ccSelFormasPago")
               
        If rsFormaPago.RecordCount > 0 Then
           If rsFormaPago!credito > 0 Then
              If rsFormaPago!Contado > 0 Then 'se pagó parte crédito y parte contado
                 'No se puede cancelar el documento el crédito tiene pagos registrados.
                 fintErrorCancelar = 368
                 Exit Function
              Else ' se pago todo a crédito
                   '¡La factura ya fue pagada!
                   fintErrorCancelar = 964
                   Exit Function
              End If
           Else
              If rsFormaPago!Contado = 0 Then 'se pagó todo a contado
                 'No se encontró la forma de pago equivalente para el departamento seleccionado.
                  fintErrorCancelar = 783
                  Exit Function
              End If
           End If
        Else
           'No se encontró la forma de pago equivalente para el departamento seleccionado.
           fintErrorCancelar = 783
           Exit Function
        End If
               
    End If
    
    'Si la factura tiene poliza directa, que el periodo contable esté abierto para cancelarla
    If Val(grdFactura.TextMatrix(grdFactura.Row, cintColNumPoliza)) <> 0 Then
        fintErrorCancelar = fintErrorContable(MskFecha.Text)
    End If
    
    'Si la factura entró en un corte, como se registrará la cancelación en el corte actual, validar que exista uno abiero
    If fintErrorCancelar = 0 And Val(grdFactura.TextMatrix(grdFactura.Row, cintColNumPoliza)) = 0 Then
        llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
        If llngNumCorte = 0 Then
            fintErrorCancelar = 659 'No se encontró un corte abierto.
            Exit Function
        End If
    End If
    
    'Si la factura entró en un corte, bloquearlo para registrar la cancelación
    If Val(grdFactura.TextMatrix(grdFactura.Row, cintColNumPoliza)) = 0 Then
        fintErrorCancelar = fintErrorBloqueoCorte(llngNumCorte)
    End If
End Function

Private Sub cmdEnd_Click()
    grdFactura.Row = grdFactura.Rows - 1
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, IIf(grdFactura.TextMatrix(grdFactura.Row, cintColPTimbre) = 1, 0, IIf(Trim(grdFactura.TextMatrix(grdFactura.Row, cintColchrEstatus)) = "C", 0, 1))
End Sub

Private Sub cmdLocate_Click()
    pReinicia
    SSTFactura.Tab = 1
    vlnblnLocate = True '' Caso [15290]
    If optMostrarSolo(2).Value = vbChecked Then
       optMostrarSolo(2).SetFocus
    ElseIf optMostrarSolo(1).Value = vbChecked Then
       optMostrarSolo(1).SetFocus
    Else
        mskFechaIni.SetFocus
    End If
        
End Sub

Private Sub cmdNext_Click()
    If grdFactura.Row < grdFactura.Rows - 1 Then
        grdFactura.Row = grdFactura.Row + 1
    End If
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, IIf(grdFactura.TextMatrix(grdFactura.Row, cintColPTimbre) = 1, 0, IIf(Trim(grdFactura.TextMatrix(grdFactura.Row, cintColchrEstatus)) = "C", 0, 1))
End Sub

Private Sub cmdSave_Click()
    
    If Not fblnDatosValidos() Then Exit Sub
    
    Dim arrDatosFisc() As DatosFiscales
    ReDim arrDatosFisc(0)
    arrDatosFisc(0).strDomicilio = Left(Trim(lblDomicilio.Caption), 380)
    arrDatosFisc(0).strNumExterior = Left(Trim(lblNumeroExterior.Caption), 50)
    arrDatosFisc(0).strNumInterior = Left(Trim(lblNumeroInterior.Caption), 50)
    arrDatosFisc(0).strTelefono = Left(Trim(lblTelefono.Caption), 50)
    arrDatosFisc(0).lstrCalleNumero = Left(Trim(lstrCalleNumero), 250)
    arrDatosFisc(0).lstrColonia = Left(Trim(lstrColonia), 100)
    arrDatosFisc(0).lstrCiudad = Left(Trim(lstrCiudad), 100)
    arrDatosFisc(0).lstrEstado = Left(Trim(lstrEstado), 100)
    arrDatosFisc(0).lstrCodigo = Left(Trim(lstrCodigo), 20)
    arrDatosFisc(0).llngCveCiudad = llngCveCiudad
    
    Dim tipoPago As DirectaMasiva
    tipoPago.intDirectaMasiva = 0
    vllngFormatoaUsar = llngFormato
    
    If CboMotivosFactura.ListIndex = 2 And txtNumeroPredial.Text <> "" Then
        vgblnhaynumpredial = True
        VGLNUMPREDIAL = Trim(txtNumeroPredial.Text)
    ElseIf CboMotivosFactura.ListIndex = 2 And txtNumeroPredial.Text = "" Then
        vgblnhaynumpredial = True
        VGLNUMPREDIAL = ""
    Else
        vgblnhaynumpredial = False
        VGLNUMPREDIAL = ""
    End If
    
    
    pGeneraFacturaDirecta Me, vllngFormatoaUsar, intTipoEmisionComprobante, intTipoCFDFactura, chkBitExtranjero.Value, txtRFC.Text, vlstrRazonSocial, llngPersonaGraba, vlblnEsCredito, vlblnPagoForma, chkMovimiento.Value, _
                            lblTotal.Caption, aFormasPago(), IIf(optPesos(0).Value, 1, 0), vldblTipoCambio, llngNumReferencia, lstrTipoCliente, lblnEntraCorte, MskFecha.Text, lblFolio.Caption, llngNumPoliza, cboUsoCFDI.ListIndex, _
                        cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex), CboMotivosFactura.ListIndex, chkRetencionIVA.Value, _
                        lblIVA.Caption, lblDescuentos.Caption, txtNumCliente.Text, lblTipoCambio.Caption, _
                        cstrCantidad4Decimales, strSerie, lblRetencionISR.Caption, lblRetencionIVA.Caption, Left(Trim(txtObservaciones.Text), 200), cboTarifa.ItemData(cboTarifa.ListIndex), llngNumCorte, _
                        chkFacturaSustitutaDFP.Value, lstFacturaASustituirDFP.ListCount, strAnoAprobacion, strNumeroAprobacion, cstrCantidad, cintTipoFormato, vlblnMultiempresa, chkRetencionISR.Value, cboTarifa.ListIndex, arrTarifas(), vgintnumemprelacionada, vlblnCuentaIngresoSaldada, vlintBitSaldarCuentas, apoliza(), arrDatosFisc, lblnConsulta, vldblTotalIVACredito, llngNumCtaCliente, llngNumFormaCredito, lblSubtotal.Caption, dblProporcionIVA, vldblComisionIvaBancaria, tipoPago, aPoliza2()
                        
                        
End Sub

'Private Sub cmdSave_Click()
'    Dim intError As Integer         'Error en transacción
'    Dim clsFacturaDirecta As clsFactura
'    Dim rsFactura As New ADODB.Recordset
'    Dim strTotalLetras As String
'    Dim lngidfactura As Long
'    Dim vlRFC As String
'    Dim vlNombre As String
'    Dim strSentencia As String
'    Dim rs As New ADODB.Recordset
'    Dim vllngPvFacturaConsecutivo As Long
'    Dim vllngCorteUsado As Long
'    Dim vlblnBandera As Boolean
'    Dim vlstrTipoPacienteCredito As String          'Sería 'PI' 'PE' 'EM' 'CO' 'ME'
'    Dim vllngCveClienteCredito As Long              'Clave del empledo o del médico
'    Dim intUsoCFDI As Integer
'    Dim intcontador As Integer
'    Dim i As Integer
'
' On Error GoTo NotificaError
'
'    If Not fblnDatosValidos() Then Exit Sub
'
'    '*********************************** OPCIONES AGREGADAS PARA CFD'S ************************************
'    'Identifica el tipo de formato a utilizar
'    vllngFormatoaUsar = llngFormato
'
'    'Se valida en caso de no haber formato activo mostrar mensaje y cancelar transacción
'    If vllngFormatoaUsar = 0 Then
'        'No se encontró un formato válido de factura.
'        MsgBox SIHOMsg(373), vbCritical, "Mensaje"
'        pReinicia
'        Exit Sub
'    End If
'
'    'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
'    '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
'    intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", vllngFormatoaUsar)
'
'    'Si los folios y los formatos no son compatibles...
'    If intTipoEmisionComprobante = 0 Then   'ERROR
'        'Si es error, se cancela la transacción
'        Exit Sub
'    End If
'
'    If intTipoEmisionComprobante = 2 Then
'        'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
'        intTipoCFDFactura = fintTipoCFD("FA", vllngFormatoaUsar)
'
'        'Si aparece un error terminar la transacción
'        If intTipoCFDFactura = 3 Then   'ERROR
'            'Si es error, se cancela la transacción
'            Exit Sub
'        End If
'    End If
'
'    If chkBitExtranjero.Value = 1 Then
'        vlRFC = "XEXX010101000"
'    Else
'        vlRFC = IIf(Len(fStrRFCValido(txtRFC.Text)) < 12 Or Len(fStrRFCValido(txtRFC.Text)) > 13, "XAXX010101000", fStrRFCValido(txtRFC.Text))
'    End If
'
'    'vlNombre = Trim(lblCliente.Caption)
'    vlNombre = Trim(vlstrRazonSocial)
'
''****************************************************************************************
'    'Validar uso del comprobante y las claves de productos/servicios y unidades
'    If Not fblnValidaSAT Then
'        Exit Sub
'    End If
'
'    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
'    If llngPersonaGraba = 0 Then Exit Sub
'
'    '------------------------------------------------------------------------------------
'    ' Agregado para caso 8644
'    vlblnEsCredito = False
'    vlblnPagoForma = False
'    If chkMovimiento.Value = 0 Then
'        If Val(Format(lblTotal.Caption, "")) > 0 Then
'            vlblnPagoForma = fblnFormasPagoPos(aFormasPago(), IIf(optPesos(0).Value, Val(Format(lblTotal.Caption, "")), Val(Format(lblTotal.Caption, "")) * vldblTipoCambio), True, vldblTipoCambio, True, llngNumReferencia, lstrTipoCliente, Trim(Replace(Replace(Replace(txtRFC.Text, "-", ""), "_", ""), " ", "")), False, False, True, "frmFacturacionDirecta")
'
'            If vlblnPagoForma Then
'                intcontador = 0
'                Do While intcontador <= UBound(aFormasPago(), 1)
'                    If aFormasPago(intcontador).vlbolEsCredito Then
'                        vlblnEsCredito = True
'                    End If
'                    intcontador = intcontador + 1
'                Loop
'            End If
'        End If
'
'        If Not (vlblnPagoForma) Then Exit Sub             ' Si <ESC> a las formas de pago
'    End If
'    '------------------------------------------------------------------------------------
'
'    EntornoSIHO.ConeccionSIHO.BeginTrans
'
'    intError = fintErrorGrabar()
'
'    If intError = 0 Then
'        Set clsFacturaDirecta = New clsFactura
'        If Not lblnEntraCorte Then
'            '------------------------------------------------------'
'            ' 1.- Insertar la póliza de la factura y tomar el id.
'            llngNumPoliza = flngInsertarPoliza(CDate(mskFecha.Text), "D", "FACTURA " & Trim(lblFolio.Caption), llngPersonaGraba)
'            '------------------------------------------------------'
'            ' 2.- Guardar la factura
'            If cboUsoCFDI.ListIndex > -1 Then
'                intUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
'            Else
'                intUsoCFDI = 0
'            End If
'
'            If CboMotivosFactura.ListIndex = 3 And chkRetencionIVA.Value = 1 Then
'                lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(lblFolio.Caption), (CDate(mskFecha.Text) + fdtmServerHora), vlRFC, vlNombre, lblDomicilio.Caption, lblNumeroExterior.Caption, lblNumeroInterior.Caption, Val(Format(lblIVA.Caption, cstrCantidad4Decimales)), Val(Format(lblDescuentos.Caption, cstrCantidad4Decimales)), " ", CLng(txtNumCliente.Text), "C", vgintNumeroDepartamento, llngPersonaGraba, 0, 0, Val(Format(lblTotal.Caption, cstrCantidad4Decimales)), IIf(optPesos(0).Value, 1, 0), Val(Format(lblTipoCambio.Caption)), lblTelefono.Caption, "C", CLng(txtNumCliente.Text), 0, llngNumPoliza, lstrCalleNumero, lstrColonia, lstrCiudad, lstrEstado, lstrCodigo, glngCveImpuesto, llngCveCiudad, strFolio, strSerie, intUsoCFDI, CDbl(lblRetencionISR.Caption), 0, CboMotivosFactura.ListIndex, 0, CDbl(lblRetencionIVA.Caption), txtObservaciones.Text)
'
'            Else
'                lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(lblFolio.Caption), (CDate(mskFecha.Text) + fdtmServerHora), vlRFC, vlNombre, lblDomicilio.Caption, lblNumeroExterior.Caption, lblNumeroInterior.Caption, Val(Format(lblIVA.Caption, cstrCantidad4Decimales)), Val(Format(lblDescuentos.Caption, cstrCantidad4Decimales)), " ", CLng(txtNumCliente.Text), "C", vgintNumeroDepartamento, llngPersonaGraba, 0, 0, Val(Format(lblTotal.Caption, cstrCantidad4Decimales)), IIf(optPesos(0).Value, 1, 0), Val(Format(lblTipoCambio.Caption)), lblTelefono.Caption, "C", CLng(txtNumCliente.Text), 0, llngNumPoliza, lstrCalleNumero, lstrColonia, lstrCiudad, lstrEstado, lstrCodigo, glngCveImpuesto, llngCveCiudad, strFolio, strSerie, intUsoCFDI, CDbl(lblRetencionISR.Caption), CDbl(lblRetencionIVA.Caption), CboMotivosFactura.ListIndex, cboTarifa.ItemData(cboTarifa.ListIndex), , txtObservaciones.Text)
'            End If
'
'            '------------------------------------------------------'
'            ' 3.- Guardar del detalle de la factura
'            pGuardaDetalleFactura (lngidfactura)
'            '------------------------------------------------------'
'            ' 4.- Guardar el detalle de la póliza
'            pGuardaDetallePoliza (lngidfactura)
'            '------------------------------------------------------'
'            ' 5.- Guardar el movimiento de crédito
'            pGuardaCredito lngidfactura
'            '------------------------------------------------------'
'            ' 6.- Liberar para que se pueda hacer un cierre
'            pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
'            '------------------------------------------------------'
'            ' 7.- Insertar en tabla de movimientos fuera del corte
'            vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & "FA" & "|" & CStr(llngNumPoliza) & "|" & CStr(llngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento)
'            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSMOVIMIENTOFUERACORTE"
'        Else
'            '------------------------------------------------------'
'            ' 1.- Guardar la factura
'            If cboUsoCFDI.ListIndex > -1 Then
'                intUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
'            Else
'                intUsoCFDI = 0
'            End If
'            If CboMotivosFactura.ListIndex = 3 And chkRetencionIVA.Value = 1 Then
'                lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(lblFolio.Caption), (CDate(mskFecha.Text) + fdtmServerHora), vlRFC, vlNombre, lblDomicilio.Caption, lblNumeroExterior.Caption, lblNumeroInterior.Caption, Val(Format(lblIVA.Caption, cstrCantidad4Decimales)), Val(Format(lblDescuentos.Caption, cstrCantidad4Decimales)), " ", CLng(txtNumCliente.Text), "C", vgintNumeroDepartamento, llngPersonaGraba, llngNumCorte, 0, Val(Format(lblTotal.Caption, cstrCantidad4Decimales)), IIf(optPesos(0).Value, 1, 0), Val(Format(lblTipoCambio.Caption)), lblTelefono.Caption, "C", CLng(txtNumCliente.Text), 0, 0, lstrCalleNumero, lstrColonia, lstrCiudad, lstrEstado, lstrCodigo, glngCveImpuesto, llngCveCiudad, strFolio, strSerie, intUsoCFDI, CDbl(lblRetencionISR.Caption), 0, CboMotivosFactura.ListIndex, 0, CDbl(lblRetencionIVA.Caption), txtObservaciones.Text)
'            Else
'                lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(lblFolio.Caption), (CDate(mskFecha.Text) + fdtmServerHora), vlRFC, vlNombre, lblDomicilio.Caption, lblNumeroExterior.Caption, lblNumeroInterior.Caption, Val(Format(lblIVA.Caption, cstrCantidad4Decimales)), Val(Format(lblDescuentos.Caption, cstrCantidad4Decimales)), " ", CLng(txtNumCliente.Text), "C", vgintNumeroDepartamento, llngPersonaGraba, llngNumCorte, 0, Val(Format(lblTotal.Caption, cstrCantidad4Decimales)), IIf(optPesos(0).Value, 1, 0), Val(Format(lblTipoCambio.Caption)), lblTelefono.Caption, "C", CLng(txtNumCliente.Text), 0, 0, lstrCalleNumero, lstrColonia, lstrCiudad, lstrEstado, lstrCodigo, glngCveImpuesto, llngCveCiudad, strFolio, strSerie, intUsoCFDI, CDbl(lblRetencionISR.Caption), CDbl(lblRetencionIVA.Caption), CboMotivosFactura.ListIndex, cboTarifa.ItemData(cboTarifa.ListIndex), , txtObservaciones.Text)
'            End If
'            '------------------------------------------------------'
'            ' 2.- Guardar del detalle de la factura
'            pGuardaDetalleFactura (lngidfactura)
'            '------------------------------------------------------'
'
'            'inicializamos el arreglo del corte
'            pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
'
'            '------------------------------------------------------'
'            ' 3.- Guardar la factura en el corte
'            pGuardaFacturaCorte (lngidfactura)
'            '------------------------------------------------------'
'            ' 4.- Guardar la póliza en el corte
'            pGuardaPolizaCorte
'            '------------------------------------------------------'
'            ' 5.- Guardar el movimiento de crédito
'            pGuardaCredito lngidfactura
'            '------------------------------------------------------'
'            ' 6.- Registra en el corte
'            vllngCorteUsado = fRegistrarMovArregloCorte(llngNumCorte, True)
'
'            If vllngCorteUsado = 0 Then
'               EntornoSIHO.ConeccionSIHO.RollbackTrans
'               'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
'               MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
'               Exit Sub
'            Else
'              If vllngCorteUsado <> llngNumCorte Then
'             'actualizamos el corte en el que se registró la factura, esto es por si hay un cambio de corte al momento de hacer el registro de la información de la factura
'              pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & lngidfactura
'              End If
'            End If
'        End If
'
'
'        If chkFacturaSustitutaDFP.Value = 1 And lstFacturaASustituirDFP.ListCount > 0 Then
'            For i = 0 To UBound(aFoliosPrevios())
'                If aFoliosPrevios(i).chrfoliofactura <> "" Then
'                    pEjecutaSentencia "INSERT INTO PVREFACTURACION (chrFolioFacturaActivada, chrFolioFacturaCancelada) " & " VALUES ('" & Trim(lblFolio.Caption) & "', '" & aFoliosPrevios(i).chrfoliofactura & "')"
'                End If
'            Next i
'        End If
'
'
'
'        '-------------------------------------------------------------------------------------------------
'        'VALIDACIÓN DE LOS DATOS ANTES DE INSERTAR EN GNCOMPROBANTEFISCLADIGITAL EN EL PROCESO DE TIMBRADO
'        '-------------------------------------------------------------------------------------------------
'        If intTipoEmisionComprobante = 2 Then
'           If Not fblnValidaDatosCFDCFDi(lngidfactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacion), strNumeroAprobacion) Then
'              EntornoSIHO.ConeccionSIHO.RollbackTrans
'              Exit Sub
'           End If
'        End If
'
'        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, llngPersonaGraba, "FACTURACION DIRECTA A CLIENTES", lblFolio.Caption)
'        EntornoSIHO.ConeccionSIHO.CommitTrans 'cerramos transacción, ya esta lista la factura
'
'        '*** GENERACIÓN DEL CFD ***
'
'            '<Si se realizará una emisión digital>
'        If intTipoEmisionComprobante = 2 Then
'            '|Genera el comprobante fiscal digital para la factura
'            'Barra de progreso CFD
'            pgbBarraCFD.Value = 70
'            freBarraCFD.Top = 3200
'            Screen.MousePointer = vbHourglass
'            lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere..."
'            freBarraCFD.Visible = True
'            freBarraCFD.Refresh
'            frmFacturacionDirecta.Enabled = False
'            If intTipoCFDFactura = 1 Then
'               pLogTimbrado 2
'               pMarcarPendienteTimbre lngidfactura, "FA", vgintNumeroDepartamento
'            End If
'            EntornoSIHO.ConeccionSIHO.BeginTrans 'iniciamos transaccion de timbrado
'            If Not fblnGeneraComprobanteDigital(lngidfactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, IIf(intTipoCFDFactura = 1, True, False)) Then
'                On Error Resume Next
'
'                EntornoSIHO.ConeccionSIHO.CommitTrans
'                If intTipoCFDFactura = 1 Then pLogTimbrado 1
'                If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
'                   'El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
'                   MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura directa"), vbInformation + vbOKOnly, "Mensaje"
'                ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then  'No se realizó el timbrado
'                   '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
'                   MsgBox SIHOMsg(1338), vbCritical + vbOKOnly, "Mensaje"
'                   pCancelarFacturaDirecta CLng(Trim(Me.txtNumCliente.Text)), Trim(lblFolio.Caption), lblnEntraCorte, llngPersonaGraba, vllngCorteUsado, llngNumPoliza
'
'                   'Actualiza PDF al cancelar facturas
'                   If Not fblnGeneraComprobanteDigital(lngidfactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, False, True, -1) Then
'                          On Error Resume Next
'                   End If
'
'                   If intTipoCFDFactura = 1 Then pEliminaPendientesTimbre lngidfactura, "FA" 'quitamos la factura de pendientes de timbre fiscal
'                   'Imprimimos la factura cancelada
'                   fblnImprimeComprobanteDigital lngidfactura, "FA", "I", llngFormato, 1
'                   Screen.MousePointer = vbDefault
'                   frmFacturacionDirecta.Enabled = True
'                   freBarraCFD.Visible = False
'                   pReinicia
'                   txtNumCliente.SetFocus
'                   Exit Sub
'              End If
'            Else
'
'
'               EntornoSIHO.ConeccionSIHO.CommitTrans
'               If intTipoCFDFactura = 1 Then
'                  pLogTimbrado 1
'                  pEliminaPendientesTimbre lngidfactura, "FA" 'quitamos la factura de pendientes de timbre fiscal
'               End If
'            End If
'            pgbBarraCFD.Value = 100
'            freBarraCFD.Top = 3200
'            Screen.MousePointer = vbDefault
'            freBarraCFD.Visible = False
'            frmFacturacionDirecta.Enabled = True
'        End If
'
'        '*** IMPRESIÓN DEL CFD ***
'
'        '<Si se realizará una emisión digital>
'        If intTipoEmisionComprobante = 2 Then
'            If Not fblnImprimeComprobanteDigital(lngidfactura, "FA", "I", llngFormato, 1) Then
'                Exit Sub
'            End If
'
'            'Verifica si debe mostrarse la pantalla de envío de CFDs por correo electrónico
'            If fblnPermitirEnvio And vgIntBanderaTImbradoPendiente = 0 Then
'                If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
'                    pEnviarCFD "FA", lngidfactura, CLng(vgintClaveEmpresaContable), vlRFC, llngPersonaGraba, Me
'                End If
'            End If
'
'        Else
'        '<Emisión física>
'            'Asegúrese de que la impresora esté   lista y  presione aceptar.
'            MsgBox SIHOMsg(343), vbOKOnly + vbInformation, "Mensaje"
'            If vgintNumeroModulo <> 2 Then
'                strTotalLetras = fstrNumeroenLetras(CDbl(Format(lblTotal.Caption, cstrCantidad)), IIf(optPesos(0).Value, "pesos", "dólares"), IIf(optPesos(0).Value, "M.N.", " "))
'                vgstrParametrosSP = Trim(lblFolio.Caption) & "|" & strTotalLetras
'                Set rsFactura = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptFactura")
'                pImpFormato rsFactura, cintTipoFormato, llngFormato
'            Else
'                pImprimeFormato llngFormato, lngidfactura
'            End If
'        End If
'
'        pReinicia
'        txtNumCliente.SetFocus
'
'    Else
'        EntornoSIHO.ConeccionSIHO.RollbackTrans
'        MsgBox SIHOMsg(intError), vbOKOnly + vbExclamation, "Mensaje"
'    End If
'
'Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
'    cmdSave.Enabled = False
'    lblnConsulta = False
'    Unload Me
'
'End Sub

Private Sub PcargarInformacionSubrogado()
On Error GoTo NotificaError
    Dim vlintCont As Integer
    Dim vlintCicl As Integer
    Dim rsCveConcepto As ADODB.Recordset
    Dim strParametros As String
       
'Const cintColCveConcepto = 1            'Columna donde se guardará la clave del concepto seleccionado
'Const cintColIVAConcepto = 2            'Columna donde se guardará el porcentaje que tiene el concepto de facturación
'Const cintColDescripcion = 3            'Descripción del concepto seleccionado
'Const cintColPrecioUnitario = 4               'Cantidad a facturar del concepto
'Const cintColDescuento = 5              'Descuento a facturar del concepto
'Const cintColIVA = 6                    'IVA a facturar del concepto
'Const cintColCtaIngreso = 7             'Cuenta contable para el ingreso
'Const cintColCtaDescuento = 8           'Cuenta contable para el descuento
'Const cintColDeptoConcepto = 9          'Cve. del departamento del concepto

'.TextMatrix(.Row, 1) = rs!Nombre 'Nombre del Paciente
'.TextMatrix(.Row, 2) = Format(rs!Fecha, "dd/mm/yyyy") 'Descripción del servicio
'.TextMatrix(.Row, 3) = rs!cantidadcargo 'Nombre de la empresa
'.TextMatrix(.Row, 4) = rs!descripcion 'Tipo de convenio
'.TextMatrix(.Row, 5) = IIf(rs!inttipoacuerdo = 1, rs!mnyCantidad & "%", Format(rs!mnyCantidad, "$ ###,###,###,###,###.00"))
'.TextMatrix(.Row, 6) = rs!MNYIVA
'.TextMatrix(.Row, 7) = rs!intNumeroCuenta
'.TextMatrix(.Row, 8) = rs!inttipoacuerdo
'.TextMatrix(.Row, 9) = rs!IVA



    For vlintCont = 1 To frmSeleccionarCargosSub.grdCargosProveedores.Rows - 1
        If frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 0) = "*" Then
            If frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont) <> 0 Then
               'If vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescripcion) <> "" Then vsfConcepto.Rows = vsfConcepto.Rows + 1
               For vlintCicl = 1 To vsfConcepto.Rows - 1
                If Val(vsfConcepto.TextMatrix(vlintCicl, cintColCveCargoMultiEmp)) = frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont) Then
                    frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 0) = ""
                End If
               Next
            End If
        End If
    Next

    For vlintCont = 1 To frmSeleccionarCargosSub.grdCargosProveedores.Rows - 1
        If frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 0) = "*" Then
            If frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont) <> 0 Then
               If vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescripcion) <> "" Then vsfConcepto.Rows = vsfConcepto.Rows + 1
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCveConcepto) = frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 11)
'               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColIVAConcepto) = "test"
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescripcion) = frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 5)
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColPrecioUnitario) = Format(IIf(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 9) = 0, frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 4) / 100)), "$ ###,###,###,###,###.00")
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDescuento) = "0.00"
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColIVA) = Format(IIf(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 9) = 0, CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 10)), CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 7) / 100)), "$ ###,###,###,###,###.00")
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCveCargoMultiEmp) = frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont)
                strParametros = CStr(vgintClaveEmpresaContable) & "|" & frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 11)
                Set rsCveConcepto = frsEjecuta_SP(strParametros, "SP_PVSELCVECONCEPTO")
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCtaIngreso) = rsCveConcepto!INTCUENTACONTABLE
               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCtaDescuento) = rsCveConcepto!intCuentaDescuento
               
'               vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColDeptoConcepto) = "test"
               ldblCantidadFactura = IIf(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 9) = 0, CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")), CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 4) / 100)) + ldblCantidadFactura
               ldblIVAFactura = IIf(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 9) = 0, CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 10)), CDbl(Replace(frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 6), "%", "")) * (frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 7) / 100)) + ldblIVAFactura
                                
            End If
        End If
    Next
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":PcargarInformacionSubrogado"))
End Sub

Private Sub pGrabarFactMultiempSub()
On Error GoTo NotificaError
    Dim vlintCont As Integer
    Dim strSentencia As String
       
    For vlintCont = 1 To frmSeleccionarCargosSub.grdCargosProveedores.Rows - 1
        If frmSeleccionarCargosSub.grdCargosProveedores.TextMatrix(vlintCont, 0) = "*" Then
            If frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont) <> 0 Then
               vgstrParametrosSP = frmSeleccionarCargosSub.grdCargosProveedores.RowData(vlintCont) _
                            & "|" & Trim(lblFolio.Caption) _
                            & "|" & vgintnumemprelacionada _
                            & "|" & vgintClaveEmpresaContable _
                            & "|" & txtNumCliente.Text
                frsEjecuta_SP vgstrParametrosSP, "SP_CCINSMULTEMPCARGOS"
            End If
        End If
    Next
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabarFactMultiempSub"))
End Sub

Private Function fblnDatosValidos() As Boolean
    Dim intcontador As Integer
    Dim rs As ADODB.Recordset
    Dim vintNumConceptos As Integer
    Dim vlblniva As Boolean
    Dim vlstrTipoCliente As String
    Dim vlstrSentencia As String
    Dim rsDatosConcepto As ADODB.Recordset
    
    fblnDatosValidos = True
    
    fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 2260, 609), "E")
    
    If fblnDatosValidos And Trim(lblCliente.Caption) = "" Then
        fblnDatosValidos = False
        'Seleccione el cliente.
        MsgBox SIHOMsg(322), vbExclamation + vbOKOnly, "Mensaje"
        txtNumCliente.SetFocus
    End If
    If fblnDatosValidos And Not lblnCreditoVigente Then
        fblnDatosValidos = False
        'El cliente seleccionado no tiene crédito activo.
        MsgBox SIHOMsg(727), vbExclamation + vbOKOnly, "Mensaje"
        txtNumCliente.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(MskFecha.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
        MskFecha.SetFocus
    End If
    If fblnDatosValidos Then
        If CDate(MskFecha.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
            MskFecha.SetFocus
       ElseIf CDate(MskFecha.Text) < frmMenuPrincipal.vgdateFechaInicioOperaciones Then
            fblnDatosValidos = False
            'La fecha debe ser mayor o igual a la de inicio de operaciones de la empresa.
            MsgBox SIHOMsg(681), vbExclamation + vbOKOnly, "Mensaje"
            MskFecha.SetFocus
       End If
    End If
    If fblnDatosValidos And optPesos(1).Value Then
        'lblTipoCambio.Caption = FormatCurrency(fdblTipoCambio(CDate(mskFecha.Text), "C"), 2)
        If Val(Format(lblTipoCambio.Caption, cstrCantidad)) = 0 Then
            fblnDatosValidos = False
            'No está registrado el tipo de cambio del día.
            MsgBox SIHOMsg(231), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
    If fblnDatosValidos And Val(vsfConcepto.TextMatrix(1, cintColCveConcepto)) = 0 Then
        fblnDatosValidos = False
        'Seleccione al menos un concepto de facturación.
        MsgBox SIHOMsg(482), vbExclamation + vbOKOnly, "Mensaje"
        vsfConcepto.Row = 1
        vsfConcepto.Col = cintColDescripcion
        vsfConcepto.SetFocus
    End If
    If fblnDatosValidos Then
        If Val(Format(lblTotal.Caption)) = 0 Then
            fblnDatosValidos = False
            'No se puede realizar la operación con cantidad cero o menor que cero
            MsgBox SIHOMsg(651), vbExclamation + vbOKOnly, "Mensaje"
            vsfConcepto.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        intcontador = 1
        Do While intcontador <= IIf(vlblnMultiempresa, vsfConcepto.Rows - 1, vsfConcepto.Rows - 2) And fblnDatosValidos
            If Val(Format(vsfConcepto.TextMatrix(intcontador, cintColImporte), cstrCantidad)) = 0 Then
                fblnDatosValidos = False
                vsfConcepto.Row = intcontador
                If Val(vsfConcepto.TextMatrix(intcontador, cintColCantidad)) = "0" Or Val(vsfConcepto.TextMatrix(intcontador, cintColCantidad)) = "0.00" Then
                    vsfConcepto.Col = cintColCantidad
                Else
                    If Val(Format(vsfConcepto.TextMatrix(intcontador, cintColPrecioUnitario), cstrCantidad)) = 0 Then
                        vsfConcepto.Col = cintColPrecioUnitario
                    End If
                End If
                'No se puede realizar la operación con cantidad cero o menor que cero
                MsgBox SIHOMsg(651), vbExclamation + vbOKOnly, "Mensaje"
                vsfConcepto.SetFocus
            End If
            intcontador = intcontador + 1
        Loop
    End If
    
    If fblnDatosValidos Then
        If chkRetencionISR.Value = 1 Then
            vgstrParametrosSP = "-1|1"
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
            If rs.RecordCount = 0 Then
                cboTarifa.Clear
            End If
        
            If Trim(cboTarifa.Text) = "" Then
                pCargaTasasRetencionISR
                pAsignaTotales
                
                If Trim(cboTarifa.Text) = "" Then
                    'No se han configurado las tasas para la retención de ISR, favor de verificar.
                    MsgBox SIHOMsg(1546), vbExclamation, "Mensaje"
                    
                    fblnDatosValidos = False
                                        
                    chkRetencionISR.SetFocus
                End If
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        If chkRetencionIVA.Value = 1 And CboMotivosFactura.ListIndex <> 3 Then
            Dim ObjRs As New ADODB.Recordset
            Set ObjRs = frsRegresaRs("Select vchvalor from siparametro where vchnombre = 'MNYPORCENTAJERETIVA'", adLockOptimistic)
            If ObjRs.RecordCount > 0 Then
               If CDbl(IIf(Trim(ObjRs!vchvalor) = "", "0", Trim(ObjRs!vchvalor))) = 0 Then
                  gdblPorcentajeRetIVA = 0
                    'No se encuentra registrado el % de retención de IVA en los parámetros generales.'
                    MsgBox SIHOMsg(748), vbExclamation + vbOKOnly, "Mensaje"
                  
                  fblnDatosValidos = False
                  pAsignaTotales
                  
                  chkRetencionIVA.SetFocus
               Else
                    gdblPorcentajeRetIVA = CDbl(IIf(Trim(ObjRs!vchvalor) = "", "0", Trim(ObjRs!vchvalor)))
                    pAsignaTotales
               End If
            Else
                gdblPorcentajeRetIVA = 0
                'No se encuentra registrado el % de retención de IVA en los parámetros generales.'
                MsgBox SIHOMsg(748), vbExclamation + vbOKOnly, "Mensaje"
                
                fblnDatosValidos = False
                pAsignaTotales
                
                chkRetencionIVA.SetFocus
            End If
        End If
    End If
    
    If fblnDatosValidos Then
        If CboMotivosFactura.ListIndex <> 0 Then
            vlblniva = False
            For intcontador = 1 To IIf(vlblnMultiempresa, vsfConcepto.Rows - 1, vsfConcepto.Rows - 2)
                If CDbl(Format(vsfConcepto.TextMatrix(intcontador, cintColIVA), cstrCantidad4Decimales)) > 0 Then vlblniva = True
            Next intcontador
            'Se valida si se obliga a retener los impuestos ISR y/o IVA
            If chkRetencionISR.Value = 0 And chkRetencionIVA.Value = 0 And vlblniva And fblnValidaRetencion Then
                'Seleccione el dato.
                MsgBox SIHOMsg(431), vbExclamation + vbOKOnly, "Mensaje"
                fblnDatosValidos = False
                
                If chkRetencionISR.Enabled Then
                    chkRetencionISR.SetFocus
                Else
                    If chkRetencionIVA.Enabled Then
                        chkRetencionIVA.SetFocus
                    End If
                End If
            End If
        
            If CboMotivosFactura.ListIndex = 1 Then
                'Honorarios profesionales
                
                If fblnDatosValidos And chkRetencionISR.Value = 1 Then
                    'Provision
                    If fblnCuentaProvision(1, 1, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionISR.SetFocus
                    End If
                End If
                If fblnDatosValidos And chkRetencionISR.Value = 1 Then
                    'Retención
                    If fblnCuentaRetencionImpuestos(1, 1, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionISR.SetFocus
                    End If
                End If
                
                If fblnDatosValidos And chkRetencionIVA.Value = 1 Then
                    'Provision
                    If fblnCuentaProvision(1, 2, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionIVA.SetFocus
                    End If
                End If
                If fblnDatosValidos And chkRetencionIVA.Value = 1 Then
                    'Retención
                    If fblnCuentaRetencionImpuestos(1, 2, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionIVA.SetFocus
                    End If
                End If
            ElseIf CboMotivosFactura.ListIndex = 2 Then
                'Arrendamiento
                
                If fblnDatosValidos And chkRetencionISR.Value = 1 Then
                    'Provision
                    If fblnCuentaProvision(2, 1, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionISR.SetFocus
                    End If
                End If
                If fblnDatosValidos And chkRetencionISR.Value = 1 Then
                    'Retención
                    If fblnCuentaRetencionImpuestos(2, 1, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionISR.SetFocus
                    End If
                End If
                
                If fblnDatosValidos And chkRetencionIVA.Value = 1 Then
                    'Provision
                    If fblnCuentaProvision(2, 2, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionIVA.SetFocus
                    End If
                End If
                If fblnDatosValidos And chkRetencionIVA.Value = 1 Then
                    'Retención
                    If fblnCuentaRetencionImpuestos(2, 2, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionIVA.SetFocus
                    End If
                End If
            Else
                If fblnDatosValidos And chkRetencionIVA.Value = 1 Then
                    'Retención
                    If fblnCuentaProvision(3, 2, True) = 0 Then
                        fblnDatosValidos = False
                        chkRetencionIVA.SetFocus
                    End If
                End If
            End If
        End If
    End If
    
    If fblnDatosValidos And Trim(cboUsoCFDI.Text) = "" Then
        fblnDatosValidos = False
        MsgBox "Selecccione el uso del CFDI", vbExclamation + vbOKOnly, "Mensaje"
        cboUsoCFDI.SetFocus
    End If
    
    If fblnDatosValidos Then
        fblnDatosValidos = fblnAsignaImpresora(vgintNumeroDepartamento, "FA")
        If Not fblnDatosValidos Then
            'No se tiene asignada una impresora en la cual imprimir las facturas
            MsgBox SIHOMsg(492), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If

    If fblnDatosValidos Then
        intcontador = 1
        vintNumConceptos = IIf(vlblnMultiempresa, vsfConcepto.Rows - 1, vsfConcepto.Rows - 2)
        If vintNumConceptos = 1 And Val(vsfConcepto.TextMatrix(intcontador, cintColBitExento)) = 1 And CboMotivosFactura.ListIndex <> 0 Then
            fblnDatosValidos = False
            'No se puede realizar la operación, el concepto de facturación es exento de IVA
            MsgBox "No se puede realizar la operación, el concepto de facturación es exento de IVA", vbExclamation + vbOKOnly, "Mensaje"
            vsfConcepto.SetFocus
        End If
    End If
    
    If fblnDatosValidos Then
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
    
      
     vlstrTipoCliente = ""
     Select Case lstrTipoCliente
        Case "PI"
            vlstrTipoCliente = "Paciente interno"
        Case "PE"
            vlstrTipoCliente = "Paciente externo"
        Case "EM"
            vlstrTipoCliente = "Empledo"
        Case "CO"
            vlstrTipoCliente = "Empresa"
        Case "ME"
            vlstrTipoCliente = "Médico"
    End Select
     
    
     If fblnDatosValidos Then
            If vlStrVersionCFDI = cstrCFDI Then
                If vlStrREgimenFiscal = "" Then
                    fblnDatosValidos = False
                    MsgBox "No se cuenta con el régimen fiscal del cliente: " & vlstrTipoCliente, vbExclamation + vbOKOnly, "Mensaje"
                    txtRFC.SetFocus
                Else
                    If lblCP.Caption = "" Then
                    fblnDatosValidos = False
                    MsgBox "No se cuenta con el código postal del cliente: " & vlstrTipoCliente, vbExclamation + vbOKOnly, "Mensaje"
                    txtRFC.SetFocus
                End If
            End If
     End If
 End If
 
 'Se valida que las cuentas puente de los créditos para facturar acepten movimientos
   If fblnDatosValidos And lblnCreditosaFacturar = True Then
        vlstrSentencia = "select CnCuenta.bitEstatusMovimientos bitEstatusMovimientos From CnCuenta " & _
         " Where CnCuenta.bitEstatusActiva = 1 " & _
         " and CnCuenta.INTNUMEROCUENTA = (select vchvalor from siparametro " & _
         " where vchnombre = 'INTCTAINGRESOCREDITOSAFACTURAR' AND CHRMODULO = 'CN' " & _
         "  AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable & ") "
        Set rsDatosConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsDatosConcepto!bitEstatusMovimientos = 0 Then
             fblnDatosValidos = False
             'La cuenta contable de ingresos para créditos para facturar no acepta movimientos
             MsgBox "La cuenta contable de ingresos para créditos para facturar no acepta movimientos.", vbExclamation + vbOKOnly, "Mensaje"
         Else
            rsDatosConcepto.Close
            vlstrSentencia = "select CnCuenta.bitEstatusMovimientos bitEstatusMovimientos From CnCuenta " & _
             " Where CnCuenta.bitEstatusActiva = 1 " & _
             " and CnCuenta.INTNUMEROCUENTA = (select vchvalor from siparametro " & _
             " where vchnombre = 'INTCTADESCUENTOCREDITOSAFACTURAR' AND CHRMODULO = 'CN' " & _
             "  AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable & ") "
            Set rsDatosConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rsDatosConcepto!bitEstatusMovimientos = 0 Then
                 fblnDatosValidos = False
                 'La cuenta contable de descuento para créditos para facturar no acepta movimientos
                 MsgBox "La cuenta contable de descuento para créditos para facturar no acepta movimientos.", vbExclamation + vbOKOnly, "Mensaje"
             Else
                rsDatosConcepto.Close
                vlstrSentencia = "select CnCuenta.bitEstatusMovimientos bitEstatusMovimientos From CnCuenta " & _
                 " Where CnCuenta.bitEstatusActiva = 1 " & _
                 " and CnCuenta.INTNUMEROCUENTA = (select vchvalor from siparametro " & _
                 " where vchnombre = 'INTCTACLIENTESCREDITOSAFACTURAR' AND CHRMODULO = 'CN' " & _
                 "  AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable & ") "
                Set rsDatosConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsDatosConcepto!bitEstatusMovimientos = 0 Then
                     fblnDatosValidos = False
                     'La cuenta contable de clientes para créditos para facturar no acepta movimientos
                     MsgBox "La cuenta contable de clientes para créditos para facturar no acepta movimientos.", vbExclamation + vbOKOnly, "Mensaje"
                 End If
             End If
         End If
   End If
'----
End Function

Private Sub cmdTop_Click()
    grdFactura.Row = 1
    pMuestra
    pHabilita 1, 1, 1, 1, 1, 0, IIf(grdFactura.TextMatrix(grdFactura.Row, cintColPTimbre) = 1, 0, IIf(Trim(grdFactura.TextMatrix(grdFactura.Row, cintColchrEstatus)) = "C", 0, 1))
End Sub

Private Sub Form_Activate()
    Dim intMensaje As Integer

    intMensaje = CInt(flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P"))

    If intMensaje <> 0 Then
        'Cierre el corte actual antes de registrar este documento.
        'No existe un corte abierto
        MsgBox SIHOMsg(Str(intMensaje)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

    If glngCveImpuesto = 0 Then
        'No se encuentra registrada la tasa de IVA en los parámetros generales del sistema.
        MsgBox SIHOMsg(731), vbOKOnly + vbExclamation, "Mensaje"
        Unload Me
        Exit Sub
    End If
    If llngFormato = 0 Then
        'Configure el formato de impresión en los parámetros del módulo.
        MsgBox SIHOMsg(732), vbOKOnly + vbExclamation, "Mensaje"
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
       ' vlnblnLocate = False '' Caso [15290]
        Unload Me
   End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (ActiveControl.Name <> "vsfConcepto" And ActiveControl.Name <> "chkRetencionISR") Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim vgstrParametrosSP As String
    Dim rs As ADODB.Recordset
    Dim intcontador As Long

    cintColCveConcepto = 1
    cintColIVAConcepto = 2
    cintColDescripcion = 3
    cintColCantidad = 4
    cintColPrecioUnitario = 5
    cintColImporte = 6
    cintColDescuento = 7
    cintColIVA = 8
    cintColCtaIngreso = 9
    cintColCtaDescuento = 10
    cintColDeptoConcepto = 11
    cintColCveCargoMultiEmp = 12
    cintColBitExento = 13
    cintColumnas = 14
    
    cstrCantidad = "#############.00"
    cstrCantidad4Decimales = "#############.0000"
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlnblnEmpresaPersonaFisica = IIf(Len(Trim(Replace(Replace(Replace(vgstrRfCCH, "-", ""), "_", ""), " ", ""))) = 13, True, False)
    If Trim(Replace(Replace(Replace(vgstrRfCCH, "-", ""), "_", ""), " ", "")) = "MAG041126GT8" Then
        vlnblnEmpresaPersonaFisica = True
    End If
    
    pCargaCboMotFactura False
    pCargaTasasRetencionISR
        
    'If Not vlnblnEmpresaPersonaFisica Then
    '    CboMotivosFactura.Enabled = True
    '    chkRetencionISR.Enabled = False
    '    chkRetencionISR.Value = 0
        
    '    chkRetencionIVA.Enabled = False
    '    chkRetencionIVA.Value = 0
    '    cboTarifa.Enabled = False
    'Else
        CboMotivosFactura.Enabled = True
        
        chkRetencionISR.Enabled = False
        chkRetencionISR.Value = 0
        
        chkRetencionIVA.Enabled = False
        chkRetencionIVA.Value = 0
        
        cboTarifa.Enabled = False
    'End If
    
        
        
    Chkmostrar.Enabled = fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 2263, 621), "C")
    
    lstrConceptos = fstrConceptos()
    
    'Busca el formato de factura directa configurado para el departamento
    llngFormato = flngFormatoDepto(vgintNumeroDepartamento, cintTipoFormato, "*")

    blnNoMensaje = False
    
    pReinicia
    pCargaUsosCFDI
    
    strSentencia = " Select VCHVERSIONCFDI from CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(strSentencia)
    
    If rs.RecordCount > 0 Then
        vlStrVersionCFDI = Trim(rs!vchVersionCFDI)
    End If
     
    
    'chkFacturaSustitutaDFP.Visible = False  '' Caso [15290]
    'lstFacturaASustituirDFP.Visible = False '' Caso [15290]
    
    cboUsoCFDI.Enabled = False
    
    SSTFactura.Tab = 0
End Sub

Private Sub pCargaTasasRetencionISR()
    cboTarifa.Clear
    ReDim arrTarifas(0)
    Dim rs As ADODB.Recordset
    Dim intcontador As Integer
    
    vgstrParametrosSP = "-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
    If rs.RecordCount <> 0 Then
    
        intcontador = 0
        Do While Not rs.EOF
            ReDim Preserve arrTarifas(intcontador)
    
            cboTarifa.AddItem rs!Descripcion
            cboTarifa.ItemData(cboTarifa.newIndex) = rs!IdTarifa
            
            arrTarifas(intcontador).lngId = rs!IdTarifa
            arrTarifas(intcontador).dblPorcentaje = rs!Porcentaje
        
            intcontador = intcontador + 1
        
            rs.MoveNext
        Loop
        
        cboTarifa.ListIndex = 0
    End If
End Sub

Public Sub pReinicia()
    lblnConsulta = False
    
    fraTotales.Enabled = True
    
    fraCliente.Enabled = True
    fraFolioFecha.Enabled = True
    fraMoneda.Enabled = True
    fraDetalle.Enabled = False
    MskFecha.Enabled = True
    
    txtNumCliente.Text = ""
    txtObservaciones.Text = ""
    
    MskFecha.Mask = ""
    MskFecha.Text = fdtmServerFecha
    MskFecha.Mask = "##/##/####"

    chkBitExtranjero.Enabled = True
    lblFolio.ForeColor = llngColorActivas

    pCargaFolio 0
    
    Chkmostrar.Visible = True
    chkMovimiento.Visible = True

    pLimpiavsfConcepto
    pConfiguravsfConcepto
    
    optPesos(0).Value = True
    optPesos_Click 0
   
    mskFechaIni.Mask = ""
    mskFechaIni.Text = fdtmServerFecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    
    txtNumClienteBusqueda.Text = ""
    lblNombreCliente.Caption = ""
    
    pLimpiagrdFactura
    pConfiguragrdFactura
    
    chkBitExtranjero.Value = vbUnchecked
    cmdCFD.Enabled = False
    cmdConfirmarTimbre.Enabled = False
    cmdConfirmartimbrefiscal.Enabled = False
    cmdCancelaFacturasSAT.Enabled = False
    txtPendienteTimbre.Visible = False
    
    vgConsecutivoMuestraPvFactura = 0
    vlblnMultiempresa = False
    
    pHabilita 0, 0, 1, 0, 0, 0, 0
    cmdAgregarCargSub.Enabled = False
    optMostrarSolo(0).Value = True
    blnNoMensaje = True
    blnNoMensaje = False
    mskFechaIni.Enabled = True
    mskFechaFin.Enabled = True
    txtNumClienteBusqueda.Enabled = True
    lblNombreCliente.Enabled = True
    cmdCargar.Enabled = True
    
    CboMotivosFactura.ListIndex = 0
    
    If Not vlnblnEmpresaPersonaFisica Then
        CboMotivosFactura.Enabled = True
        chkRetencionISR.Enabled = False
        chkRetencionISR.Value = 0
        
        chkRetencionIVA.Enabled = False
        chkRetencionIVA.Value = 0
        cboTarifa.Enabled = False
    Else
        CboMotivosFactura.Enabled = True
        
        chkRetencionISR.Enabled = False
        chkRetencionISR.Value = 0
        
        chkRetencionIVA.Enabled = False
        chkRetencionIVA.Value = 0
        
        cboTarifa.Enabled = False
    End If
    
    cboUsoCFDI.Enabled = False
    
    CboMotivosFactura_Click
    
    chkMovimiento.Enabled = False
    
   
    chkFacturaSustitutaDFP.Value = 0
    lstFacturaASustituirDFP.Clear
    chkFacturaSustitutaDFP.Enabled = False
    lstFacturaASustituirDFP.Enabled = False
    
    chkFacturaSustitutaDFP.Visible = True
    lstFacturaASustituirDFP.Visible = True
    
    lblnCreditosaFacturar = False
    cmdCreditosaFacturar.Visible = True
    cmdCreditosaFacturar.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SSTFactura.Tab = 0 Then
        If cmdSave.Enabled Or lblnConsulta Then
            Cancel = True
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pReinicia
                txtNumCliente.SetFocus
            End If
        End If
    End If
    
    If SSTFactura.Tab = 1 Or SSTFactura.Tab = 2 Then
        Cancel = True
        SSTFactura.Tab = 0
        pReinicia
        txtNumCliente.SetFocus
    End If
End Sub

Private Sub grdCreditosaFacturar_DblClick()
    If Trim(grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Row, 0)) = "" And Trim(grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Row, 8)) <> "" Then
        grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Row, 0) = "*"
    Else
        grdCreditosaFacturar.TextMatrix(grdCreditosaFacturar.Row, 0) = ""
    End If
    
    cmdAgregarDatos.Enabled = True
End Sub

Private Sub grdFactura_Click()
    With grdFactura
        If .MouseCol = 0 And .MouseRow > 0 Then
            If .TextMatrix(.Row, cintColPTimbre) = "1" Then
                If .TextMatrix(.Row, 0) = "*" Then
                    If vllngSeleccPendienteTimbre > 0 Then
                       vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre - 1
                    End If
                    .TextMatrix(.Row, 0) = ""
                Else
                    vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                    .TextMatrix(.Row, 0) = "*"
                End If
                Me.cmdConfirmartimbrefiscal.Enabled = vllngSeleccPendienteTimbre > 0
            ElseIf (.TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) <> "NP" And .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) <> "CR") Then
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
        
'''           If .TextMatrix(.Row, cintColPFacSAT) = "1" Or .TextMatrix(.Row, cintColPFacSAT) = "2" Then
'''                If .TextMatrix(.Row, 0) = "*" Then
'''                    If vllngSeleccionadas > 0 Then
'''                       vllngSeleccionadas = vllngSeleccionadas - 1
'''                    End If
'''                    .TextMatrix(.Row, 0) = ""
'''                Else
'''                    vllngSeleccionadas = vllngSeleccionadas + 1
'''                    .TextMatrix(.Row, 0) = "*"
'''                End If
'''              Me.cmdCancelaFacturasSAT.Enabled = vllngSeleccionadas > 0
              
        End If
    End With
End Sub

Private Sub grdFactura_DblClick()
If grdFactura.MouseCol > 0 And grdFactura.MouseRow > 0 Then
    If Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio)) <> "" Then
        If grdFactura.TextMatrix(grdFactura.Row, cintColPFacSAT) <> "2" Then
           SSTFactura.Tab = 0
           pMuestra
           pHabilita 1, 1, 1, 1, 1, 0, IIf(grdFactura.TextMatrix(grdFactura.Row, cintColPTimbre) = "1", 0, IIf(Trim(grdFactura.TextMatrix(grdFactura.Row, cintColchrEstatus)) = "C", 0, 1))
           cmdLocate.SetFocus
        Else
           'no existen información
           MsgBox SIHOMsg(13), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
End If
End Sub

Private Sub grdFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdFactura_DblClick
    End If
End Sub

Private Sub mskCFFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdCargarCreditosaFacturar.SetFocus
    End If
End Sub

Private Sub mskCFFechaFin_LostFocus()
    If mskCFFechaFin.Text <> "  /  /    " Then
        If Not IsDate(mskCFFechaFin.Text) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
             MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
             pEnfocaMkTexto mskCFFechaIni
             Exit Sub
        End If
        If CDate(mskCFFechaFin.Text) < CDate("01/01/1900") Then
             'Fecha no valida.
             MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
             pEnfocaMkTexto mskCFFechaFin
             Exit Sub
        End If
        If mskCFFechaFin.Text <> "  /  /    " And mskCFFechaIni.Text <> "  /  /    " Then
            If CDate(mskCFFechaFin.Text) < CDate(mskCFFechaIni.Text) Then
                 'Fecha no valida.
                 MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
                 pEnfocaMkTexto mskCFFechaFin
                 Exit Sub
            End If
        End If
    End If
End Sub

Private Sub mskCFFechaIni_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        pEnfocaMkTexto mskCFFechaFin
'    End If
End Sub

Private Sub mskCFFechaIni_LostFocus()
    If mskCFFechaIni.Text <> "  /  /    " Then
        If Not IsDate(mskCFFechaIni.Text) Then
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
             MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
             pEnfocaMkTexto mskCFFechaIni
             Exit Sub
        End If
        If CDate(mskCFFechaIni.Text) < CDate("01/01/1900") Then
             'Fecha no valida.
             MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
             pEnfocaMkTexto mskCFFechaIni
             Exit Sub
        Else
            pEnfocaMkTexto mskCFFechaFin
        End If
    Else
        'Fecha no valida.
        MsgBox SIHOMsg(254), vbExclamation, "Mensaje"
        pEnfocaMkTexto mskCFFechaIni
        Exit Sub
    End If
    
End Sub

Private Sub mskFecha_Change()
    optPesos(1).Enabled = IsDate(MskFecha.Text)
End Sub

Private Sub mskFecha_GotFocus()
    pSelMkTexto MskFecha
End Sub

Private Sub MskFecha_LostFocus()
    Dim X As Integer
    
    vldblTipoCambio = fdblTipoCambio(CDate(MskFecha.Text), "O")
    If Val(Format(vldblTipoCambio, cstrCantidad)) = 0 Then
        'Registre el tipo de cambio del día.
        MsgBox SIHOMsg(335), vbOKOnly + vbInformation, "Mensaje"
    End If
    If optPesos(0).Value = True Then
        lblTipoCambio.Caption = ""
    Else
        lblTipoCambio.Caption = Trim(Str(fdblTipoCambio(CDate(MskFecha.Text), "O")))
    End If
End Sub

Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaIni_GotFocus()
    pSelMkTexto mskFechaIni
End Sub

Private Sub optMostrarSolo_Click(Index As Integer)
    If Index = 0 Then
        mskFechaIni.Enabled = True
        mskFechaFin.Enabled = True
        txtNumClienteBusqueda.Enabled = True
        lblNombreCliente.Enabled = True
        cmdCargar.Enabled = True
    Else
        mskFechaIni.Enabled = False
        mskFechaFin.Enabled = False
        txtNumClienteBusqueda.Enabled = False
        lblNombreCliente.Enabled = False
        cmdCargar.Enabled = False
    End If
    cmdCargar_Click
End Sub

Private Sub optPesos_Click(Index As Integer)
    Dim X As Integer
    
    If Index = 0 Then
        lblTipoCambio.Caption = ""
        vldblTipoCambio = fdblTipoCambio(CDate(MskFecha.Text), "O")
        If Val(Format(vldblTipoCambio, cstrCantidad)) = 0 Then
            'Registre el tipo de cambio del día.
            MsgBox SIHOMsg(335), vbOKOnly + vbInformation, "Mensaje"
        Else
            With vsfConcepto
                For X = 1 To .Rows - 1
                    If .TextMatrix(X, cintColDescripcion) <> "" Then
                        .TextMatrix(X, cintColPrecioUnitario) = FormatCurrency(.TextMatrix(X, cintColPrecioUnitario) * vldblTipoCambio, 4)
                        .TextMatrix(X, cintColDescuento) = FormatCurrency(.TextMatrix(X, cintColDescuento) * vldblTipoCambio, 4)
                        .TextMatrix(X, cintColIVA) = FormatCurrency(.TextMatrix(X, cintColIVA) * vldblTipoCambio, 4)
                        .TextMatrix(X, cintColImporte) = FormatCurrency(.TextMatrix(X, cintColImporte) * vldblTipoCambio, 4)
                    End If
                Next X
            End With
            ldblImporteFactura = ldblImporteFactura * vldblTipoCambio
            ldblDescuentosFactura = ldblDescuentosFactura * vldblTipoCambio
            ldblCantidadFactura = ldblCantidadFactura * vldblTipoCambio
            ldblIVAFactura = ldblIVAFactura * vldblTipoCambio
            pAsignaTotales
        End If
    Else
        'lblTipoCambio.Caption = FormatCurrency(fdblTipoCambio(CDate(MskFecha.Text), "C"), 2)
        lblTipoCambio.Caption = Trim(Str(fdblTipoCambio(CDate(MskFecha.Text), "O")))
        vldblTipoCambio = fdblTipoCambio(CDate(MskFecha.Text), "O")
        If Val(Format(lblTipoCambio.Caption, cstrCantidad)) = 0 Then
            'Registre el tipo de cambio del día.
            MsgBox SIHOMsg(335), vbOKOnly + vbInformation, "Mensaje"
        Else
            With vsfConcepto
                For X = 1 To .Rows - 1
                    If .TextMatrix(X, cintColDescripcion) <> "" Then
                        .TextMatrix(X, cintColPrecioUnitario) = FormatCurrency(.TextMatrix(X, cintColPrecioUnitario) / vldblTipoCambio, 4)
                        .TextMatrix(X, cintColDescuento) = FormatCurrency(.TextMatrix(X, cintColDescuento) / vldblTipoCambio, 4)
                        .TextMatrix(X, cintColIVA) = FormatCurrency(.TextMatrix(X, cintColIVA) / vldblTipoCambio, 4)
                        .TextMatrix(X, cintColImporte) = FormatCurrency(.TextMatrix(X, cintColImporte) / vldblTipoCambio, 4)
                    End If
                Next X
            End With
            ldblImporteFactura = ldblImporteFactura / vldblTipoCambio
            ldblDescuentosFactura = ldblDescuentosFactura / vldblTipoCambio
            ldblCantidadFactura = ldblCantidadFactura / vldblTipoCambio
            ldblIVAFactura = ldblIVAFactura / vldblTipoCambio
            pAsignaTotales
        End If
    End If
End Sub

Private Sub pLimpiavsfConcepto()
    vsfConcepto.Clear
    vsfConcepto.Rows = 2
    vsfConcepto.Cols = cintColumnas
    vsfConcepto.FormatString = cstrFormato

    vsfConcepto.Col = cintColDescripcion
    vsfConcepto.Row = 1

    ldblDescuentosFactura = 0
    ldblCantidadFactura = 0
    ldblIVAFactura = 0
    ldblImporteFactura = 0

    pAsignaTotales

    CmdBorrar.Enabled = False
End Sub

Private Sub pLimpiagrdFactura()
    grdFactura.Clear
    grdFactura.Rows = 2
    grdFactura.Cols = cintColgrdFactura
    grdFactura.FormatString = cstrFormatgrdFactura
End Sub



Private Sub txtNumCliente_Change()
    lblCliente.Caption = ""
    lblDomicilio.Caption = ""
    lblCiudad.Caption = ""
    txtRFC.Text = ""
    lblTelefono.Caption = ""
    chkBitExtranjero.Value = vbUnchecked
    lblNumeroExterior.Caption = ""
    lblNumeroInterior.Caption = ""
    lblColonia.Caption = ""
    lblCP.Caption = ""

    cboUsoCFDI.ListIndex = -1
    
    lstFacturaASustituirDFP.Clear
    chkFacturaSustitutaDFP.Value = 0
End Sub

Private Sub txtNumCliente_GotFocus()
    pSelTextBox txtNumCliente
End Sub

Private Sub pAsignaDatosCliente(rs As ADODB.Recordset)
    Dim strTipoUsoCFDI As String
    Dim lngTipoUsoCFDI As Long
    Dim rsTmp As ADODB.Recordset
    Dim rsDatosFISC As ADODB.Recordset
    txtNumCliente.Text = rs!intNumCliente
    lblCliente.Caption = IIf(IsNull(rs!NombreCliente), " ", rs!NombreCliente)
    vlstrRazonSocial = IIf(IsNull(rs!RazonSocial), " ", rs!RazonSocial)
    lblCiudad.Caption = IIf(IsNull(rs!ciudadcliente), " ", rs!ciudadcliente)
    If IsNull(rs!RFCCliente) Then
        txtRFC.Text = ""
    Else
        txtRFC.Text = fStrRFCValido(rs!RFCCliente)
    End If
    lblTelefono.Caption = IIf(IsNull(rs!Telefono), " ", rs!Telefono)
    lblDomicilio.Caption = IIf(IsNull(rs!chrCalle), " ", rs!chrCalle)
    lblNumeroExterior.Caption = IIf(IsNull(rs!VCHNUMEROEXTERIOR), " ", rs!VCHNUMEROEXTERIOR)
    lblNumeroInterior.Caption = IIf(IsNull(rs!VCHNUMEROINTERIOR), " ", rs!VCHNUMEROINTERIOR)
    lblColonia.Caption = IIf(IsNull(rs!Colonia), " ", rs!Colonia)
    lblCP.Caption = Trim(IIf(IsNull(rs!Codigo), " ", rs!Codigo))
    vlStrREgimenFiscal = ""
    If rs!chrTipoCliente <> "PI" And rs!chrTipoCliente <> "PE" Then
        vlStrREgimenFiscal = IIf(IsNull(rs!REGIMENFISCAL), " ", Trim(rs!REGIMENFISCAL))
    Else
        Set rsDatosFISC = frsRegresaRs("SELECT PVDATOSFISCALES.VCHREGIMENFISCAL, MAX(INTNUMCUENTA) FROM PVDATOSFISCALES WHERE TRIM(CHRRFC) = '" & Trim(rs!RFCCliente) & "' GROUP BY PVDATOSFISCALES.VCHREGIMENFISCAL")
        If rsDatosFISC.RecordCount > 0 Then
            vlStrREgimenFiscal = IIf(IsNull(rsDatosFISC!vchregimenfiscal), "", Trim(rsDatosFISC!vchregimenfiscal))
        Else
            vlStrREgimenFiscal = ""
        End If
    End If
    
    Select Case rs!chrTipoCliente
        Case "EM"
            strTipoUsoCFDI = "EP"
            lngTipoUsoCFDI = 0
        Case "CO"
            strTipoUsoCFDI = "EM"
            lngTipoUsoCFDI = rs!intNumReferencia
        Case "PI", "PE"
            strTipoUsoCFDI = "TP"
            lngTipoUsoCFDI = 0
            Set rsTmp = frsRegresaRs("select intCveTipoPaciente from EXPacienteIngreso where intNumCuenta = " & rs!intNumReferencia & " and chrTipoIngreso = '" & IIf(rs!chrTipoCliente = "PI", "I", "E") & "'")
            If Not rsTmp.EOF Then
                lngTipoUsoCFDI = rsTmp!intCveTipoPaciente
            End If
            rsTmp.Close
        Case "ME"
            strTipoUsoCFDI = "ME"
            lngTipoUsoCFDI = 0
        Case Else
            strTipoUsoCFDI = "XX"
            lngTipoUsoCFDI = 0
    End Select
    cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, flngCatalogoSATIdByNombreTipo("c_UsoCFDI", lngTipoUsoCFDI, strTipoUsoCFDI, 1))
   
End Sub

Private Sub txtNumCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vllngNumCliente As Long
    Dim rs As New ADODB.Recordset
    Dim rsMultiemp As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        vllngCveProveedorM = 0
        vlstrproveedorM = ""
        vgintnumemprelacionada = 0
        vlblnMultiempresa = False
        cmdAgregarCargSub.Enabled = False
        If Trim(txtNumCliente.Text) = "" Then
            vllngNumCliente = flngNumCliente(True, 1)
        Else
            vllngNumCliente = Val(txtNumCliente.Text)
        End If
        If vllngNumCliente <> 0 Then
            Set rs = frsEjecuta_SP(Str(vllngNumCliente) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|1", "sp_CcSelDatosCliente")
            If rs.RecordCount <> 0 Then
                'verificar si multiempresa y todo eso
                'IF multiempresa y sub y todo eso then vlblnMultiempresa = True
                'Set rsMultiemp = frsRegresaRs("select * from siempresacliente simp left join cccliente ccc on  simp.INTIDEMPRESACLIENTE = ccc.INTNUMREFERENCIA and ccc.CHRTIPOCLIENTE = 'CO' where simp.TNYIDEMPRESA = " & vgintClaveEmpresaContable)
                vgstrParametrosSP = vgintClaveEmpresaContable & "|" & rs!intNumReferencia & "|" & vgintNumeroDepartamento
                Set rsMultiemp = frsEjecuta_SP(vgstrParametrosSP, "sp_CCSelEmpresaProveedor")
                If rsMultiemp.RecordCount <> 0 Then
                    Do While Not rsMultiemp.EOF
                        If rsMultiemp!idempresacliente = rs!intNumReferencia Then
                            If rsMultiemp!idproveedor <> 0 Then
                                vlblnMultiempresa = True
                                cmdAgregarCargSub.Enabled = True
                                vllngCveProveedorM = rsMultiemp!idproveedor
                                vlstrproveedorM = rsMultiemp!proveedor
                                vgintnumemprelacionada = rsMultiemp!empresa
                            Else
                               ' Exit Sub
                            End If
                        End If
                        rsMultiemp.MoveNext
                    Loop
                End If
                rsMultiemp.Close
                lblnCreditoVigente = rs!bitactivo = 1
                llngNumCtaCliente = rs!intnumcuentacontable
                llngNumReferencia = rs!intNumReferencia
                lstrTipoCliente = rs!chrTipoCliente
                                 
                lstrCalleNumero = IIf(IsNull(rs!callenumero), "", rs!callenumero)
                lstrColonia = IIf(IsNull(rs!Colonia), "", rs!Colonia)
                lstrCiudad = IIf(IsNull(rs!ciudadcliente), "", rs!ciudadcliente)
                llngCveCiudad = IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)
                lstrEstado = IIf(IsNull(rs!Estado), "", rs!Estado)
                lstrCodigo = Trim(IIf(IsNull(rs!Codigo), "", rs!Codigo))
                vlblnEmpresa = IIf(IsNull(rs!chrTipoCliente), False, IIf(rs!chrTipoCliente = "CO", True, False))
                vldblRetServicios = IIf(IsNull(rs!RetServicios), 0, rs!RetServicios) * 100
                If vlblnEmpresa And vldblRetServicios <> 0 Then pCargaCboMotFactura True
                
                cboUsoCFDI.Enabled = True
                fraDetalle.Enabled = True
                chkMovimiento.Enabled = True
                
                CboMotivosFactura_Click
                    
                pAsignaDatosCliente rs
                pHabilita 0, 0, 0, 0, 0, 1, 0
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                pEnfocaTextBox txtNumCliente
            End If
        End If
    End If
    vlnblnLocate = True
    pFacturasDirectasAnteriores
End Sub

Private Sub txtNumCliente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub pConfiguravsfConcepto()
    With vsfConcepto
        .FixedCols = 1
        .ColWidth(0) = 100
        .ColWidth(cintColCveConcepto) = 0
        .ColWidth(cintColIVAConcepto) = 0
        .ColWidth(cintColDescripcion) = 5200
        .ColWidth(cintColCantidad) = 750
        .ColWidth(cintColPrecioUnitario) = 1200
        .ColWidth(cintColImporte) = 1300
        .ColWidth(cintColDescuento) = 1100
        .ColWidth(cintColIVA) = 1200
        .ColWidth(cintColCtaIngreso) = 0
        .ColWidth(cintColCtaDescuento) = 0
        .ColWidth(cintColDeptoConcepto) = 0
        .ColWidth(12) = 0
        .ColWidth(cintColBitExento) = 0
        .ColWidth(10) = 0
        
        .ColAlignment(cintColDescripcion) = flexAlignLeftCenter
        .ColAlignment(cintColCantidad) = flexAlignRightCenter
        .ColAlignment(cintColPrecioUnitario) = flexAlignRightCenter
        .ColAlignment(cintColImporte) = flexAlignRightCenter
        .ColAlignment(cintColDescuento) = flexAlignRightCenter
        .ColAlignment(cintColIVA) = flexAlignRightCenter
        .FixedAlignment(cintColDescripcion) = flexAlignCenterCenter
        .FixedAlignment(cintColCantidad) = flexAlignCenterCenter
        .FixedAlignment(cintColPrecioUnitario) = flexAlignCenterCenter
        .FixedAlignment(cintColImporte) = flexAlignCenterCenter
        .FixedAlignment(cintColDescuento) = flexAlignCenterCenter
        .FixedAlignment(cintColIVA) = flexAlignCenterCenter
        
        .ColDataType(cintColCantidad) = flexDTString
        
        
    End With
End Sub

Private Sub pMuestra()
    Dim rs As New ADODB.Recordset
    Dim rsnumpredial As New ADODB.Recordset
    
    Dim vlaux As String
    Dim vlSQL As String
    Dim vlrsAux As New ADODB.Recordset
    Dim rsFacturas As New ADODB.Recordset
    Dim vlstrsql As String
    Dim vllngContador As Integer
    Dim i As Integer
      
    CboMotivosFactura.Enabled = False
    fraTotales.Enabled = False
    
    vgstrParametrosSP = Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFactura_NE")
        
    fraCliente.Enabled = False
    fraFolioFecha.Enabled = False
    fraMoneda.Enabled = False
    fraDetalle.Enabled = False
        
    If rs.RecordCount <> 0 Then
        lblnConsulta = True
        
        If rs!chrrTipoCliente = "CO" And IIf(IsNull(rs!PorcentajeRetenServ), 0, rs!PorcentajeRetenServ) > 0 Then pCargaCboMotFactura True
        
        Chkmostrar.Visible = False
        cmdCreditosaFacturar.Visible = False
        chkMovimiento.Visible = False
        
        txtNumCliente.Text = rs!NumCliente
        lblCliente.Caption = rs!NombreCliente
        lblCiudad.Caption = IIf(IsNull(rs!Ciudad), " ", rs!Ciudad)
        lblTelefono.Caption = IIf(IsNull(rs!Telefono), " ", rs!Telefono)
        If IsNull(rs!RFC) Then
            txtRFC.Text = " "
        Else
            txtRFC.Text = fStrRFCValido(rs!RFC)
        End If
        vgConsecutivoMuestraPvFactura = rs!IdFactura
        lblDomicilio.Caption = IIf(IsNull(rs!chrCalle), " ", rs!chrCalle)
        lblNumeroExterior.Caption = IIf(IsNull(rs!VCHNUMEROEXTERIOR), " ", rs!VCHNUMEROEXTERIOR)
        lblNumeroInterior.Caption = IIf(IsNull(rs!VCHNUMEROINTERIOR), " ", rs!VCHNUMEROINTERIOR)
        lblColonia.Caption = IIf(IsNull(rs!Colonia), " ", rs!Colonia)
        lblCP.Caption = IIf(IsNull(rs!CP), " ", rs!CP)
        vlaux = IIf(IsNull(rs!Serie), "", rs!Serie)
        txtObservaciones.Text = IIf(IsNull(rs!Observaciones), " ", rs!Observaciones)
        
        Set rsnumpredial = frsRegresaRs("SELECT * FROM GnComprobanteFiscalDigital INNER JOIN PVFactura ON GnComprobanteFiscalDigital.INTCOMPROBANTE = PVFactura.INTCONSECUTIVO AND GnComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = 'FA' WHERE PVFactura.ChrFolioFactura = '" & Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio)) & "'")
        
        If rsnumpredial.RecordCount > 0 Then
            If IsNull(rsnumpredial!VCHNUMEROPREDIAL) Or rsnumpredial!VCHNUMEROPREDIAL = "" Then
                txtNumeroPredial.Text = ""
            Else
                txtNumeroPredial.Text = IIf(IsNull(rsnumpredial!VCHNUMEROPREDIAL), "", rsnumpredial!VCHNUMEROPREDIAL)
            End If
        End If
        
        If rs!PendienteTimbre = 0 Then
            Set vlrsAux = frsRegresaRs("SELECT * FROM GnComprobanteFiscalDigital INNER JOIN PVFactura ON GnComprobanteFiscalDigital.INTCOMPROBANTE = PVFactura.INTCONSECUTIVO AND GnComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = 'FA' WHERE PVFactura.ChrFolioFactura = '" & Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio)) & "'")
            If vlrsAux.RecordCount <> 0 Then
                cmdCFD.Enabled = True
                vlstrTipoCFD = IIf(IsNull(vlrsAux!INTNUMEROAPROBACION), "CFDi", "CFD")
            Else
                cmdCFD.Enabled = False
            End If
            cmdConfirmarTimbre.Enabled = False
            If rs!PendienteCancelarSat = 1 Then
               Me.txtPendienteTimbre.Text = "Pendiente de cancelarse ante el SAT"
               Me.txtPendienteTimbre.ForeColor = &HFF&
               Me.txtPendienteTimbre.BackColor = &HC0E0FF
               Me.txtPendienteTimbre.Visible = True
            Else
               txtPendienteTimbre.Visible = False
                Select Case rs!PendienteCancelarSAT_NE
                    Case "PA"
                        Me.txtPendienteTimbre.Text = "Pendiente de autorización de cancelación"
                        txtPendienteTimbre.ForeColor = &HFFFFFF '| Blanco
                        txtPendienteTimbre.BackColor = &H80FF&  '| Naranja fuerte
                        txtPendienteTimbre.Visible = True
                    Case "CR"
                        Me.txtPendienteTimbre.Text = "Cancelación rechazada"
                        txtPendienteTimbre.ForeColor = &HFFFFFF '| Blanco
                        txtPendienteTimbre.BackColor = &HFF&    '| Rojo
                        txtPendienteTimbre.Visible = True
                    Case "NP"
                        txtPendienteTimbre.Visible = False
                End Select

            End If
            
            If cgstrModulo = "PV" Then
                If fblnRevisaPermiso(vglngNumeroLogin, 3090, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3090, "C", True) Then
                     Me.cmdDelete.Enabled = True
                Else
                    Me.cmdDelete.Enabled = False
                End If
            ElseIf cgstrModulo = "CC" Then
                If fblnRevisaPermiso(vglngNumeroLogin, 3089, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3089, "C", True) Then
                    Me.cmdDelete.Enabled = True
                Else
                    Me.cmdDelete.Enabled = False
                End If
            End If
        Else
            cmdCFD.Enabled = False
            cmdConfirmarTimbre.Enabled = True
            Me.txtPendienteTimbre.Text = "Pendiente de timbre fiscal"
            Me.txtPendienteTimbre.ForeColor = &H0&
            Me.txtPendienteTimbre.BackColor = &HFFFF&
            Me.txtPendienteTimbre.Visible = True
            Me.cmdDelete.Enabled = False
        End If
        
        'Se habilita el chkExtranjero en caso de ser cliente Extranjero
        If Trim(rs!RFC) = "XEXX010101000" Then
            chkBitExtranjero.Value = vbChecked
        Else
            chkBitExtranjero.Value = vbUnchecked
        End If
        
        lblFolio.ForeColor = IIf(rs!chrestatus = "C", llngColorCanceladas, llngColorActivas)
        lblFolio.Caption = Trim(grdFactura.TextMatrix(grdFactura.Row, cIntColFolio))
        
        MskFecha.Mask = ""
        MskFecha.Text = Format(rs!fecha, "dd/mm/yyyy")
        vldtmFechaFactura = rs!fecha
        MskFecha.Mask = "##/##/####"
        
        optPesos(0).Value = rs!BITPESOS = 1
        optPesos(1).Value = rs!BITPESOS = 0
        
        'lblTipoCambio.Caption = FormatCurrency(rs!TipoCambio, 2)
        If optPesos(1).Value = True Then
            lblTipoCambio.Caption = rs!TipoCambio
        Else
            lblTipoCambio.Caption = ""
        End If
        vlblnCancelada = rs!chrestatus = "C"
        
        pLimpiavsfConcepto
        pConfiguravsfConcepto
        
        CboMotivosFactura.ListIndex = rs!intTipoFactura
        CboMotivosFactura_Click
        
        chkMovimiento.Enabled = False
        
        lblRetencionISR.Caption = FormatCurrency(0, 4)
        lblRetencionIVA.Caption = FormatCurrency(0, 4)
        chkRetencionISR.Value = 0
        chkRetencionIVA.Value = 0
        
        cboTarifa.ListIndex = flngLocalizaCbo(cboTarifa, rs!intIdTarifaRetencionISR)
        If cboTarifa.ListIndex = -1 Then
            cboTarifa.ListIndex = 0
        End If
        
        If rs!MNYRETENCIONISR <> 0 Then
            chkRetencionISR.Enabled = True
            chkRetencionISR.Value = 1
        End If
        If rs!MNYRETENCIONIVA <> 0 Then
            chkRetencionIVA.Enabled = True
            chkRetencionIVA.Value = 1
        ElseIf rs!MNYRETENSERVICIOS <> 0 Then
            chkRetencionIVA.Enabled = True
            chkRetencionIVA.Value = 1
        End If
        
        If rs!MNYRETENCIONISR <> 0 Then
            lblRetencionISR.Caption = FormatCurrency(rs!MNYRETENCIONISR, 4)
        End If
        If rs!MNYRETENCIONIVA <> 0 Then
            lblRetencionIVA.Caption = FormatCurrency(rs!MNYRETENCIONIVA, 4)
        ElseIf rs!MNYRETENSERVICIOS <> 0 Then
            lblRetencionIVA.Caption = FormatCurrency(rs!MNYRETENSERVICIOS, 4)
        End If
        
        lblDescuentos.Caption = FormatCurrency(rs!DescuentoFactura, 4)
        lblSubtotal.Caption = FormatCurrency(rs!Subtotal, 4)
        lblIVA.Caption = FormatCurrency(rs!IVAFactura, 4)
                
        lblTotal.Caption = FormatCurrency(rs!TotalFactura, 4)
        lblImporte.Caption = FormatCurrency(rs!Subtotal + rs!DescuentoFactura, 4)
        
        cboUsoCFDI.Enabled = False
        
        If Not IsNull(rs!intCveUsoCFDI) Then
            cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, rs!intCveUsoCFDI)
        End If
        With vsfConcepto
            Do While Not rs.EOF And Not IsNull(rs!Concepto)
                .TextMatrix(.Rows - 1, cintColDescripcion) = rs!Concepto
                .TextMatrix(.Rows - 1, cintColPrecioUnitario) = FormatCurrency(rs!Importe / rs!Unidades, 4)
                .TextMatrix(.Rows - 1, cintColDescuento) = FormatCurrency(rs!Descuento, 4)
                .TextMatrix(.Rows - 1, cintColIVA) = FormatCurrency(rs!IVA, 4)
                .TextMatrix(.Rows - 1, cintColCantidad) = FormatNumber(Format(rs!Unidades, cstrCantidad), 2)
                .TextMatrix(.Rows - 1, cintColImporte) = FormatCurrency(rs!Importe, 4)
                .Rows = .Rows + 1
                rs.MoveNext
            Loop
            .Rows = .Rows - 1
        End With
        
        chkFacturaSustitutaDFP.Visible = False
        lstFacturaASustituirDFP.Visible = False
        
    Else
        '¡La información no existe!
        MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub

Private Sub pConfiguragrdFactura()
    Dim intcontador As Integer

    With grdFactura
        .FixedCols = 1
        .ColWidth(0) = 200
        .ColWidth(cintColNumPoliza) = 0
        .ColWidth(cIntColNumCorte) = 0
        .ColWidth(cintColchrEstatus) = 0
        .ColWidth(cintColFecha) = 1100
        .ColWidth(cIntColFolio) = 1100
        .ColWidth(cintColNumCliente) = 950
        .ColWidth(cIntColRazonSocial) = 3200
        .ColWidth(cIntColRFC) = 1500
        .ColWidth(cintColTotalFactura) = 1500
        .ColWidth(cintColIVAConsulta) = 1200
        .ColWidth(cintColDescuentos) = 1100
        .ColWidth(cintColSubtotal) = 1500
        .ColWidth(cintColMoneda) = 800
        .ColWidth(cIntColEstado) = 1000
        .ColWidth(cintColFacturo) = 3000
        .ColWidth(cintColCancelo) = 3000
        .ColWidth(cintColPFacSAT) = 0 'CGR
        .ColWidth(cintColPTimbre) = 0 'CGR
        .ColWidth(cintColEstadoNuevoEsquemaCancelacion) = 0
        .ColAlignment(cintColFecha) = flexAlignLeftCenter
        .ColAlignment(cIntColFolio) = flexAlignLeftCenter
        .ColAlignment(cintColNumCliente) = flexAlignRightCenter
        .ColAlignment(cIntColRazonSocial) = flexAlignLeftCenter
        .ColAlignment(cIntColRFC) = flexAlignLeftCenter
        .ColAlignment(cintColTotalFactura) = flexAlignRightCenter
        .ColAlignment(cintColIVAConsulta) = flexAlignRightCenter
        .ColAlignment(cintColDescuentos) = flexAlignRightCenter
        .ColAlignment(cintColSubtotal) = flexAlignRightCenter
        .ColAlignment(cintColMoneda) = flexAlignLeftCenter
        .ColAlignment(cIntColEstado) = flexAlignLeftCenter
        .ColAlignment(cintColFacturo) = flexAlignLeftCenter
        .ColAlignment(cintColCancelo) = flexAlignLeftCenter
        
        For intcontador = 1 To .Cols - 1
            .ColAlignmentFixed(intcontador) = flexAlignCenterCenter
        Next intcontador
    End With
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer)
On Error GoTo NotificaError
    
    cmdTop.Enabled = intTop
    cmdBack.Enabled = intBack
    cmdLocate.Enabled = intlocate
    cmdNext.Enabled = intNext
    cmdEnd.Enabled = intEnd
    cmdSave.Enabled = intSave
    
     If cgstrModulo = "PV" Then
        If fblnRevisaPermiso(vglngNumeroLogin, 3090, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3090, "C", True) Then
             cmdDelete.Enabled = intDelete = 1
        Else
            Me.cmdDelete.Enabled = False
        End If
     ElseIf cgstrModulo = "CC" Then
        If fblnRevisaPermiso(vglngNumeroLogin, 3089, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3089, "C", True) Then
            cmdDelete.Enabled = intDelete = 1
        Else
            Me.cmdDelete.Enabled = False
        End If
     End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
    Unload Me
End Sub

Private Sub txtNumClienteBusqueda_Change()
    lblNombreCliente.Caption = ""
End Sub

Private Sub txtNumClienteBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vllngNumCliente As Long
    Dim rs As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtNumClienteBusqueda.Text) = "" Then
            vllngNumCliente = flngNumCliente(True, 1)
        Else
            vllngNumCliente = Val(txtNumClienteBusqueda.Text)
        End If
        If vllngNumCliente <> 0 Then
            Set rs = frsEjecuta_SP(Str(vllngNumCliente) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|1", "sp_CcSelDatosCliente")
            If rs.RecordCount <> 0 Then
                txtNumClienteBusqueda.Text = rs!intNumCliente
                lblNombreCliente.Caption = IIf(IsNull(rs!NombreCliente), " ", rs!NombreCliente)
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbOKOnly + vbInformation, "Mensaje"
                pEnfocaTextBox txtNumClienteBusqueda
            End If
        End If
    End If
End Sub

Private Sub txtNumeroPredial_GotFocus()
    pSelTextBox txtNumeroPredial

End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPendienteTimbre_GotFocus()
    If fblnCanFocus(cmdTop) Then cmdTop.SetFocus
End Sub
Private Sub vsfConcepto_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblPrecioUnitario As Double
    Dim dblCantidad As Double
    Dim dblDescuento As Double
    Dim dblImporte As Double
    Dim dblIVA As Double
    Dim blnerror As Boolean
    Dim lngRenglonConcepto As Long
    Dim rsConcepto As New ADODB.Recordset
    
    If Col = cintColDescripcion And vsfConcepto.ComboIndex <> -1 Then
        lngRenglonConcepto = flngExisteDatoColVsf(vsfConcepto, cintColCveConcepto, vsfConcepto.ComboData(vsfConcepto.ComboIndex))
        
        If lngRenglonConcepto <> Row Then
            If lngRenglonConcepto = -1 Then
                vgstrParametrosSP = vsfConcepto.ComboData(vsfConcepto.ComboIndex) & "|-1|-1|" & vgintClaveEmpresaContable
                Set rsConcepto = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelConceptoFacturacion")
                If rsConcepto!INTCUENTACONTABLE <> 0 And rsConcepto!intCuentaDescuento <> 0 Then
                    vsfConcepto.TextMatrix(Row, cintColCveConcepto) = rsConcepto!smicveconcepto
                    vsfConcepto.TextMatrix(Row, cintColIVAConcepto) = rsConcepto!smyIVA
                    vsfConcepto.TextMatrix(Row, cintColCtaIngreso) = rsConcepto!INTCUENTACONTABLE
                    vsfConcepto.TextMatrix(Row, cintColCtaDescuento) = rsConcepto!intCuentaDescuento
                    vsfConcepto.TextMatrix(Row, cintColDeptoConcepto) = rsConcepto!SMIDEPARTAMENTO
                    vsfConcepto.TextMatrix(Row, cintColBitExento) = rsConcepto!bitExentoIva
                    
                    If Val(vsfConcepto.TextMatrix(vsfConcepto.Rows - 1, cintColCveConcepto)) <> 0 Then
                        vsfConcepto.TextMatrix(Row, cintColCantidad) = "1.00"
                        vsfConcepto.TextMatrix(Row, cintColPrecioUnitario) = FormatCurrency(0, 4)
                        vsfConcepto.TextMatrix(Row, cintColDescuento) = FormatCurrency(0, 4)
                        vsfConcepto.TextMatrix(Row, cintColImporte) = FormatCurrency(0, 4)
                        vsfConcepto.TextMatrix(Row, cintColIVA) = FormatCurrency(0, 4)
                        vsfConcepto.Rows = vsfConcepto.Rows + 1
                    End If
                Else
                    blnerror = True
                    'No existe cuenta contable para el concepto de facturación:
                    MsgBox SIHOMsg(907) & vsfConcepto.ComboItem(vsfConcepto.ComboIndex), vbOKOnly + vbInformation, "Mensaje"
                End If
            Else
                blnerror = True
                'Este dato ya está registrado.
                MsgBox SIHOMsg(404), vbOKOnly + vbInformation, "Mensaje"
            End If
        End If
    End If
    
    If Col = cintColCantidad Or Col = cintColPrecioUnitario Or Col = cintColDescuento Or Col = cintColDescripcion Then
        'Da formato a las columnas
        vsfConcepto.TextMatrix(Row, cintColDescuento) = FormatCurrency(Val(Format(vsfConcepto.TextMatrix(Row, cintColDescuento), cstrCantidad4Decimales)), 4)
        vsfConcepto.TextMatrix(Row, cintColPrecioUnitario) = FormatCurrency(Val(Format(vsfConcepto.TextMatrix(Row, cintColPrecioUnitario), cstrCantidad4Decimales)), 4)
        vsfConcepto.TextMatrix(Row, cintColCantidad) = FormatNumber(Val(Format(vsfConcepto.TextMatrix(Row, cintColCantidad), cstrCantidad4Decimales)), 2)
        
        ' Multiplica la cantidad por el precio unitario para actualizar el importe
        If (Col = cintColCantidad Or Col = cintColPrecioUnitario) Then
            dblImporte = Val(Format(vsfConcepto.TextMatrix(Row, cintColCantidad), cstrCantidad4Decimales)) * Val(Format(vsfConcepto.TextMatrix(Row, cintColPrecioUnitario), cstrCantidad4Decimales))
            vsfConcepto.TextMatrix(Row, cintColImporte) = FormatCurrency(dblImporte, 4)
        End If
        'Valida que el Importe sea menor que el descuento
        If (Col = cintColDescuento Or Col = cintColCantidad Or Col = cintColPrecioUnitario) And Val(Format(vsfConcepto.TextMatrix(Row, cintColImporte), cstrCantidad4Decimales)) < Val(Format(vsfConcepto.TextMatrix(Row, cintColDescuento), cstrCantidad4Decimales)) Then
            blnerror = True
            'Los descuentos exceden el importe del concepto.
            MsgBox SIHOMsg(644), vbOKOnly + vbExclamation, "Mensaje"
            vsfConcepto.TextMatrix(Row, cintColDescuento) = FormatCurrency(0, 4)
        Else
            dblImporte = Val(Format(vsfConcepto.TextMatrix(Row, cintColImporte), cstrCantidad4Decimales))
            If dblImporte <> 0 Then
                dblDescuento = Val(Format(vsfConcepto.TextMatrix(Row, cintColDescuento), cstrCantidad4Decimales))
                If Not vlblnMultiempresa Then
                    dblIVA = Val(Format(vsfConcepto.TextMatrix(Row, cintColIVAConcepto), cstrCantidad4Decimales))
                    vsfConcepto.TextMatrix(Row, cintColIVA) = FormatCurrency((dblImporte - dblDescuento) * dblIVA / 100, 4)
                End If
            Else
                vsfConcepto.TextMatrix(Row, cintColDescuento) = FormatCurrency(0, 4)
                vsfConcepto.TextMatrix(Row, cintColIVA) = FormatCurrency(0, 4)
            End If
        End If
    End If
    
    ldblImporteFactura = ldblImporteFactura - ldblImporteConcepto + Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColImporte), cstrCantidad4Decimales))
    ldblDescuentosFactura = ldblDescuentosFactura - ldblDescuentoConcepto + Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColDescuento), cstrCantidad4Decimales))
    ldblCantidadFactura = ldblCantidadFactura - ldblCantidadConcepto + Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColPrecioUnitario), cstrCantidad4Decimales))
    ldblIVAFactura = ldblIVAFactura - ldblIVAConcepto + Val(Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColIVA), cstrCantidad4Decimales))
    pAsignaTotales
    
    If Not blnerror Then
        pMueveRenglonColumna
    End If
End Sub

Private Sub pAsignaTotales()
    Dim vldblRetISR As Double
    Dim vldblRetIVA As Double
    Dim vlintCont As Integer
    
    lblImporte.Caption = FormatCurrency(ldblImporteFactura, 4)
    lblDescuentos.Caption = FormatCurrency(ldblDescuentosFactura, 4)
    lblSubtotal.Caption = FormatCurrency(ldblImporteFactura - ldblDescuentosFactura, 4)
    lblIVA.Caption = FormatCurrency(ldblIVAFactura, 4)
        
    vldblRetISR = 0
    vldblRetIVA = 0
    
    'If vlnblnEmpresaPersonaFisica Then
        If CboMotivosFactura.ListIndex <> 0 Then
            If chkRetencionISR.Value = 1 And Trim(cboTarifa.Text) <> "" Then
                vldblRetISR = (ldblImporteFactura - ldblDescuentosFactura) * (arrTarifas(cboTarifa.ListIndex).dblPorcentaje / 100)
            End If
                
            If ldblIVAFactura > 0 Then
                If CboMotivosFactura.ListIndex = 3 Then
                    gdblPorcentajeRetIVA = vldblRetServicios / 100
                    If chkRetencionIVA.Value = 1 Then
                        For vlintCont = 1 To vsfConcepto.Rows - 2
                            If CDbl(Format(vsfConcepto.TextMatrix(vlintCont, cintColIVA), cstrCantidad4Decimales)) > 0 Then
                                vldblRetIVA = vldblRetIVA + (CDbl(Format(vsfConcepto.TextMatrix(vlintCont, cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(vsfConcepto.TextMatrix(vlintCont, cintColDescuento), cstrCantidad4Decimales))) * gdblPorcentajeRetIVA
                            End If
                        Next vlintCont
                    End If
                Else
                    Dim ObjRs As New ADODB.Recordset
                    Set ObjRs = frsRegresaRs("Select vchvalor from siparametro where vchnombre = 'MNYPORCENTAJERETIVA'", adLockOptimistic)
                    If ObjRs.RecordCount > 0 Then
                       If CDbl(IIf(Trim(ObjRs!vchvalor) = "", "0", Trim(ObjRs!vchvalor))) = 0 Then
                          gdblPorcentajeRetIVA = 0
                       Else
                            gdblPorcentajeRetIVA = CDbl(IIf(Trim(ObjRs!vchvalor) = "", "0", Trim(ObjRs!vchvalor)))
                       End If
                    Else
                        gdblPorcentajeRetIVA = 0
                    End If
                
                    chkRetencionIVA.Enabled = True
                    If chkRetencionIVA.Value = 1 Then
                        vldblRetIVA = ldblIVAFactura * (gdblPorcentajeRetIVA / 100)
                    End If
                End If
            Else
                chkRetencionIVA.Enabled = False
                chkRetencionIVA.Value = 0
            End If
        End If
    'End If
    
    lblRetencionISR.Caption = FormatCurrency(vldblRetISR, 4)
    lblRetencionIVA.Caption = FormatCurrency(vldblRetIVA, 4)
    
    lblTotal.Caption = FormatCurrency(ldblImporteFactura - ldblDescuentosFactura + ldblIVAFactura - vldblRetISR - vldblRetIVA, 4)
End Sub

Private Sub pMueveRenglonColumna()
    If vsfConcepto.Col = cintColDescripcion Then
        vsfConcepto.Col = cintColCantidad
    Else
        If vsfConcepto.Col = cintColCantidad Then
            vsfConcepto.Col = cintColPrecioUnitario
        Else
            If vsfConcepto.Col = cintColPrecioUnitario Then
                vsfConcepto.Col = cintColDescuento
            Else
                If vsfConcepto.Col = cintColDescuento Then
                    vsfConcepto.Col = cintColDescripcion
                    If Not vlblnMultiempresa Then
                        If vsfConcepto.Row + 1 < vsfConcepto.Rows Then
                            vsfConcepto.Row = vsfConcepto.Row + 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub vsfConcepto_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    ldblDescuentoConcepto = Val(Format(vsfConcepto.TextMatrix(Row, cintColDescuento), cstrCantidad4Decimales))
    ldblCantidadConcepto = Val(Format(vsfConcepto.TextMatrix(Row, cintColPrecioUnitario), cstrCantidad4Decimales))
    ldblIVAConcepto = Val(Format(vsfConcepto.TextMatrix(Row, cintColIVA), cstrCantidad4Decimales))
    ldblImporteConcepto = Val(Format(vsfConcepto.TextMatrix(Row, cintColImporte), cstrCantidad4Decimales))

    'Si es la columna de descripción, cargar el combo
    If Col = cintColDescripcion And Not vlblnMultiempresa Then
        vsfConcepto.ComboList = lstrConceptos
        vsfConcepto.ComboIndex = 0
    Else
        vsfConcepto.ComboList = ""
    End If
    
    'Si es la columna de IVA, no se edita
    If Col = cintColIVA Or Col = cintColImporte Then
        Cancel = True
    End If
    
    'Si es la columna de cantidad, Precio unitario o descuento y no hay concepto, no se edita
    If (Col = cintColCantidad Or Col = cintColPrecioUnitario Or Col = cintColDescuento) And Val(vsfConcepto.TextMatrix(Row, cintColCveConcepto)) = 0 Then
        Cancel = True
    End If
    
    'Si es la columna de descuento y no hay cantidad, no se edita
    If Col = cintColDescuento And Val(Format(vsfConcepto.TextMatrix(Row, cintColPrecioUnitario), cstrCantidad4Decimales)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsfConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsfConcepto.Col = cintColPrecioUnitario Then
            vsfConcepto.TextMatrix(vsfConcepto.Row, cintColPrecioUnitario) = Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColPrecioUnitario), cstrCantidad4Decimales)
        End If
        
        If vsfConcepto.Col = cintColDescuento Then
            vsfConcepto.TextMatrix(vsfConcepto.Row, cintColDescuento) = Format(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColDescuento), cstrCantidad4Decimales)
        End If
    End If
End Sub

Private Sub vsfConcepto_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col <> cintColDescripcion Then
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = Asc(".") Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub vsfConcepto_RowColChange()
    If vsfConcepto.Row <> -1 Then
        CmdBorrar.Enabled = Val(IIf(Trim(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColCveConcepto)) = "", "0", Trim(vsfConcepto.TextMatrix(vsfConcepto.Row, cintColCveConcepto)))) <> 0
    End If
End Sub

Private Sub pCargaUsosCFDI()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDI, rsTmp, 0, 1
        cboUsoCFDI.ListIndex = -1
    End If
End Sub

Public Function fblnValidaSAT() As Boolean
    Dim intRow As Integer
    If vgstrVersionCFDI <> "3.2" Then
        If cboUsoCFDI.ListIndex = -1 Then
            MsgBox "Seleccione el uso del comprobante", vbExclamation, "Mensaje"
            cboUsoCFDI.SetFocus
            fblnValidaSAT = False
            Exit Function
        End If
        For intRow = 1 To vsfConcepto.Rows - 1
            If Val(vsfConcepto.TextMatrix(intRow, 1)) > 0 Then
                If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", vsfConcepto.TextMatrix(intRow, 1), "CF", 1) = 0 Then
                    MsgBox "No está definida la clave del SAT para el producto/servicio " & vsfConcepto.TextMatrix(intRow, 3), vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    Exit Function
                End If
                If flngCatalogoSATIdByNombreTipo("c_ClaveUnidad", vsfConcepto.TextMatrix(intRow, 1), "CF", 2) = 0 Then
                    MsgBox "No está definida la clave del SAT para la unidad del producto/servicio " & vsfConcepto.TextMatrix(intRow, 3), vbExclamation, "Mensaje"
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

Private Sub pFacturasDirectasAnteriores()
    Dim rsFacturas As New ADODB.Recordset
    Dim vlstrsql As String
    Dim i As Integer
    
'        vlstrsql = "SELECT CHRFOLIOFACTURA, DTMFECHAHORA, MNYTOTALFACTURA, CASE WHEN BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, MNYTIPOCAMBIO FROM PVFACTURA " & _
'                           "INNER JOIN GNCOMPROBANTEFISCALDIGITAL ON GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'FA' " & _
'                           "AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVFACTURA.INTCONSECUTIVO " & _
'                   "WHERE NOT VCHUUID IS NULL AND INTMOVPACIENTE = " & CLng(IIf(txtNumCliente.Text = "", "0", txtNumCliente.Text)) & " AND CHRTIPOPACIENTE = 'C' AND CHRTIPOFACTURA = 'C' AND CHRESTATUS = 'C' " & _
'                   "ORDER BY INTCONSECUTIVO DESC"
        vlstrsql = "SELECT CHRFOLIOFACTURA, DTMFECHAHORA, MNYTOTALFACTURA, CASE WHEN BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, MNYTIPOCAMBIO FROM PVFACTURA " & _
                   "INNER JOIN GNCOMPROBANTEFISCALDIGITAL ON GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'FA' " & _
                   "AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVFACTURA.INTCONSECUTIVO " & _
                   "WHERE NOT VCHUUID IS NULL AND INTMOVPACIENTE = " & CLng(IIf(txtNumCliente.Text = "", "0", txtNumCliente.Text)) & " AND CHRTIPOPACIENTE = 'C' AND CHRTIPOFACTURA = 'C'" & _
                   "ORDER BY INTCONSECUTIVO DESC"
               
        Set rsFacturas = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        
        If rsFacturas.RecordCount <> 0 Then
            chkFacturaSustitutaDFP.Enabled = True
            lstFacturaASustituirDFP.Enabled = True
        Else
            chkFacturaSustitutaDFP.Enabled = False
            lstFacturaASustituirDFP.Enabled = False
        End If
        rsFacturas.Close
         vlnblnLocate = False
End Sub

Private Sub pCargaCboMotFactura(blnServicios As Boolean)
    
    CboMotivosFactura.Clear
    CboMotivosFactura.AddItem "OTROS", 0
    CboMotivosFactura.ItemData(CboMotivosFactura.newIndex) = 0
    CboMotivosFactura.AddItem "HONORARIOS PROFESIONALES", 1
    CboMotivosFactura.ItemData(CboMotivosFactura.newIndex) = 1
    CboMotivosFactura.AddItem "ARRENDAMIENTO", 2
    CboMotivosFactura.ItemData(CboMotivosFactura.newIndex) = 2
    If blnServicios Then
        CboMotivosFactura.AddItem "SERVICIOS", 3
        CboMotivosFactura.ItemData(CboMotivosFactura.newIndex) = 3
    End If
    
    CboMotivosFactura.ListIndex = 0
End Sub

Public Sub pCargaFolio(intAumenta As Integer)
    Dim vllngFoliosRestantes As Long
    Dim vlstrFolioDocumento As String
    Dim alstrParametrosSalida() As String
    Dim vllngFoliosFaltantes As Long

    vllngFoliosFaltantes = 1
    vlstrFolioDocumento = ""
    pCargaArreglo alstrParametrosSalida, vllngFoliosFaltantes & "|" & ADODB.adBSTR & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
    frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|" & Str(intAumenta), "sp_gnFolios", , , alstrParametrosSalida
    pObtieneValores alstrParametrosSalida, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
    '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
    strSerie = Trim(strSerie)
    If vllngFoliosFaltantes > 0 Then
        MsgBox "Faltan " & Trim(Str(vllngFoliosFaltantes)) + " facturas y será necesario aumentar folios!", vbOKOnly + vbInformation, "Mensaje"
    End If
    lblFolio.Caption = Trim(strSerie) + Trim(strFolio)
    strFolio = Trim(lblFolio.Caption)

    'Habilitar el chkBitExtranjero si el folio es de tipo digital
    If Trim(strNumeroAprobacion) <> "" And Trim(strAnoAprobacion) <> "" Then
        chkBitExtranjero.Enabled = True
    End If
End Sub
'Solo cuando la empresa contable emisora de la factura sea persona fisica y el receptor sea persona moral, será cuando obligue la selecciona de esas 2 casillas.
'Cuando tenga retención que marque predeterminadamente y qué no obligue a nada, si lo quitan será decisión del usuario
Private Function fblnValidaRetencion() As Boolean
On Error GoTo NotificaError
    Dim ObjRs As New ADODB.Recordset
    Set ObjRs = frsRegresaRs("select length(vchrfc) longitud from CNEMPRESACONTABLE where tnyclaveempresa = " & vgintClaveEmpresaContable)
    If ObjRs.RecordCount > 0 Then
       If ObjRs!Longitud = 13 And Len(Trim(txtRFC.Text)) = 12 Then
          fblnValidaRetencion = True
       Else
           fblnValidaRetencion = False
       End If
    End If
    ObjRs.Close
    Set ObjRs = Nothing
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnValidaRetencion"))
End Function
