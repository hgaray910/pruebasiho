VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFacturacionMembresiaSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación a socios"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   1260
      TabIndex        =   24
      Top             =   9840
      Visible         =   0   'False
      Width           =   8760
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   25
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
         TabIndex        =   26
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
   Begin TabDlg.SSTab SSTFactura 
      Height          =   10245
      Left            =   0
      TabIndex        =   1
      Top             =   -480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   18071
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmFacturacionMembresiaSocios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFolioFecha"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmFacturacionMembresiaSocios.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape3(6)"
      Tab(1).Control(1)=   "Shape3(4)"
      Tab(1).Control(2)=   "Shape3(5)"
      Tab(1).Control(3)=   "Label1(4)"
      Tab(1).Control(4)=   "Label1(2)"
      Tab(1).Control(5)=   "Label1(3)"
      Tab(1).Control(6)=   "Label1(1)"
      Tab(1).Control(7)=   "Label57(10)"
      Tab(1).Control(8)=   "Label57(11)"
      Tab(1).Control(9)=   "Label57(6)"
      Tab(1).Control(10)=   "Label57(7)"
      Tab(1).Control(11)=   "Label57(15)"
      Tab(1).Control(12)=   "Label57(12)"
      Tab(1).Control(13)=   "Label57(2)"
      Tab(1).Control(14)=   "Label57(3)"
      Tab(1).Control(15)=   "Label57(4)"
      Tab(1).Control(16)=   "Label57(5)"
      Tab(1).Control(17)=   "Label57(8)"
      Tab(1).Control(18)=   "Label57(9)"
      Tab(1).Control(19)=   "Shape1"
      Tab(1).Control(20)=   "Label21"
      Tab(1).Control(21)=   "Shape2"
      Tab(1).Control(22)=   "Label22"
      Tab(1).Control(23)=   "Label23"
      Tab(1).Control(24)=   "Label13"
      Tab(1).Control(25)=   "grdBusquedaFactura"
      Tab(1).Control(26)=   "Frame4"
      Tab(1).Control(27)=   "Frame1"
      Tab(1).Control(28)=   "cmdCargar"
      Tab(1).Control(29)=   "ChkFacturasCancelaNoSAT"
      Tab(1).Control(30)=   "ChkPendientesTimbre"
      Tab(1).Control(31)=   "Frame3"
      Tab(1).ControlCount=   32
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -70215
         TabIndex        =   75
         Top             =   7920
         Width           =   1860
         Begin VB.CommandButton cmdconfirmartimbrefiscal 
            Caption         =   "Confirmar timbre fiscal"
            Enabled         =   0   'False
            Height          =   495
            Left            =   30
            Picture         =   "frmFacturacionMembresiaSocios.frx":0038
            TabIndex        =   11
            ToolTipText     =   "Confirmar timbre"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelaFacturasSAT 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1275
            Picture         =   "frmFacturacionMembresiaSocios.frx":052A
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Cancelar factura(s) ante el SAT"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.CheckBox ChkPendientesTimbre 
         Caption         =   "Mostrar sólo facturas pendientes de timbre fiscal"
         Height          =   255
         Left            =   -71160
         TabIndex        =   9
         ToolTipText     =   "Mostrar sólo facturas pendientes de timbre fiscal"
         Top             =   1250
         Width           =   4695
      End
      Begin VB.CheckBox ChkFacturasCancelaNoSAT 
         Caption         =   "Mostrar sólo facturas sin cancelar ante el SAT"
         Height          =   255
         Left            =   -74865
         TabIndex        =   8
         ToolTipText     =   "Mostrar sólo facturas sin cancelar ante el SAT"
         Top             =   1250
         Width           =   3615
      End
      Begin VB.CommandButton cmdCargar 
         Caption         =   "&Cargar datos"
         Height          =   525
         Left            =   -64635
         TabIndex        =   7
         ToolTipText     =   "Cargar los datos con los filtros establecidos"
         Top             =   670
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Socio"
         Height          =   615
         Left            =   -71520
         TabIndex        =   52
         Top             =   600
         Width           =   6855
         Begin VB.TextBox txtBusquedaSocio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Número de socio"
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label lblBusquedaNombreSocio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   1785
            TabIndex        =   6
            ToolTipText     =   "Nombre del socio"
            Top             =   240
            Width           =   5010
         End
      End
      Begin VB.Frame fraCliente 
         Height          =   2040
         Left            =   210
         TabIndex        =   32
         Top             =   615
         Width           =   8715
         Begin VB.TextBox txtClaveSocio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            TabIndex        =   51
            ToolTipText     =   "Número de socio"
            Top             =   540
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.TextBox txtRFC 
            Height          =   315
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "RFC del socio"
            Top             =   540
            Width           =   1800
         End
         Begin VB.CheckBox chkBitExtranjero 
            Caption         =   "Extranjero"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3420
            TabIndex        =   15
            ToolTipText     =   "Extranjero"
            Top             =   570
            Width           =   1000
         End
         Begin VB.TextBox txtClaveUnica 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1500
            TabIndex        =   0
            ToolTipText     =   "Número de socio"
            Top             =   195
            Width           =   1800
         End
         Begin VB.Label Label20 
            Caption         =   "Colonia"
            Height          =   255
            Left            =   5265
            TabIndex        =   49
            Top             =   1320
            Width           =   555
         End
         Begin VB.Label lblColonia 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5925
            TabIndex        =   48
            ToolTipText     =   "Número exterior del socio"
            Top             =   1245
            Width           =   2730
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Código postal"
            Height          =   195
            Left            =   225
            TabIndex        =   47
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label lblCP 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1500
            TabIndex        =   46
            ToolTipText     =   "Código postal del socio"
            Top             =   1605
            Width           =   1185
         End
         Begin VB.Label lblNumeroInterior 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   45
            ToolTipText     =   "Número interior del socio"
            Top             =   1245
            Width           =   1125
         End
         Begin VB.Label lblNumeroExterior 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1500
            TabIndex        =   44
            ToolTipText     =   "Número exterior del socio"
            Top             =   1245
            Width           =   1185
         End
         Begin VB.Label Label17 
            Caption         =   "Número exterior"
            Height          =   255
            Left            =   225
            TabIndex        =   43
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Left            =   2805
            TabIndex        =   42
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label lblTelefono 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4020
            TabIndex        =   41
            ToolTipText     =   "Teléfono del socio"
            Top             =   1605
            Width           =   1125
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   40
            Top             =   600
            Width           =   315
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
            Height          =   195
            Left            =   5265
            TabIndex        =   39
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblCiudad 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   5925
            TabIndex        =   38
            ToolTipText     =   "Ciudad del socio"
            Top             =   1605
            Width           =   2730
         End
         Begin VB.Label lblDomicilio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1500
            TabIndex        =   37
            ToolTipText     =   "Calle del socio"
            Top             =   900
            Width           =   7155
         End
         Begin VB.Label lblSocio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Enabled         =   0   'False
            Height          =   315
            Left            =   3405
            TabIndex        =   36
            ToolTipText     =   "Nombre del socio"
            Top             =   195
            Width           =   5250
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Calle"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   945
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Clave única"
            Height          =   195
            Left            =   225
            TabIndex        =   34
            Top             =   255
            Width           =   840
         End
         Begin VB.Label Label16 
            Caption         =   "Número interior"
            Height          =   255
            Left            =   2805
            TabIndex        =   50
            Top             =   1320
            Width           =   1215
         End
      End
      Begin VB.Frame fraFolioFecha 
         Height          =   2040
         Left            =   8940
         TabIndex        =   27
         Top             =   615
         Width           =   2400
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   360
            TabIndex        =   28
            ToolTipText     =   "Fecha de la factura"
            Top             =   1080
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   "dd/mmm/yyyy"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Folio de factura"
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   360
            TabIndex        =   30
            Top             =   840
            Width           =   450
         End
         Begin VB.Label lblFolio 
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   360
            TabIndex        =   29
            ToolTipText     =   "Folio de la factura"
            Top             =   480
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fecha"
         Height          =   615
         Left            =   -74865
         TabIndex        =   2
         Top             =   600
         Width           =   3300
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
            Left            =   1965
            TabIndex        =   4
            ToolTipText     =   "Fecha final de la búsqueda"
            Top             =   240
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
            Left            =   465
            TabIndex        =   3
            ToolTipText     =   "Fecha inicial de la búsqueda"
            Top             =   240
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
            Caption         =   "Al"
            Height          =   195
            Left            =   1730
            TabIndex        =   53
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            Height          =   195
            Left            =   135
            TabIndex        =   13
            Top             =   300
            Width           =   240
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdBusquedaFactura 
         Height          =   6015
         Left            =   -74865
         TabIndex        =   10
         ToolTipText     =   "Facturas de socios"
         Top             =   1515
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   10610
         _Version        =   393216
         Cols            =   6
         GridColor       =   12632256
         FormatString    =   "|Fecha|Folio|Número|Cliente|Estado"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   240
         TabIndex        =   54
         Top             =   2760
         Width           =   11105
         _ExtentX        =   19579
         _ExtentY        =   8705
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Factura"
         TabPicture(0)   =   "frmFacturacionMembresiaSocios.frx":0A1C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label9"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraTotales"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "grdConceptos"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkOtrosDatosFiscales"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtpendientetimbre"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cboUsoCFDI"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Cargos"
         TabPicture(1)   =   "frmFacturacionMembresiaSocios.frx":0A38
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdCargos"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Pagos"
         TabPicture(2)   =   "frmFacturacionMembresiaSocios.frx":0A54
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chkIncluyePagosFacturados"
         Tab(2).Control(1)=   "grdPagos"
         Tab(2).ControlCount=   2
         Begin VB.ComboBox cboUsoCFDI 
            Height          =   315
            Left            =   1200
            TabIndex        =   16
            Text            =   "cboUsoCFDI"
            ToolTipText     =   "Uso del CFDI"
            Top             =   2950
            Width           =   4975
         End
         Begin VB.TextBox txtpendientetimbre 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   120
            TabIndex        =   76
            Text            =   "Pendiente de cancelarse ante el SAT"
            Top             =   3360
            Visible         =   0   'False
            Width           =   4495
         End
         Begin VB.CheckBox chkIncluyePagosFacturados 
            Caption         =   "Incluir facturados"
            Height          =   255
            Left            =   -74880
            TabIndex        =   68
            ToolTipText     =   "Incluir pagos que ya fueron facturados"
            Top             =   4560
            Width           =   1815
         End
         Begin VSFlex7LCtl.VSFlexGrid grdPagos 
            Height          =   4095
            Left            =   -74880
            TabIndex        =   67
            ToolTipText     =   "Pagos realizados por el socio"
            Top             =   405
            Width           =   10815
            _cx             =   19076
            _cy             =   7223
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
         Begin VB.CheckBox chkOtrosDatosFiscales 
            Caption         =   "Solicitar otros datos fiscales"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   "Solicitar otros datos fiscales antes de facturar"
            Top             =   4440
            Width           =   2535
         End
         Begin VSFlex7LCtl.VSFlexGrid grdConceptos 
            Height          =   2460
            Left            =   120
            TabIndex        =   55
            ToolTipText     =   "Factura del socio"
            Top             =   400
            Width           =   10815
            _cx             =   19076
            _cy             =   4339
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
         Begin VSFlex7LCtl.VSFlexGrid grdCargos 
            Height          =   4380
            Left            =   -74880
            TabIndex        =   66
            ToolTipText     =   "Cargos del socio"
            Top             =   405
            Width           =   10815
            _cx             =   19076
            _cy             =   7726
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
         Begin VB.Frame fraTotales 
            Enabled         =   0   'False
            Height          =   1965
            Left            =   7035
            TabIndex        =   56
            Top             =   2850
            Width           =   3885
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
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
               Height          =   240
               Left            =   230
               TabIndex        =   94
               Top             =   885
               Width           =   555
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
               Left            =   1785
               TabIndex        =   65
               ToolTipText     =   "Total a pagar"
               Top             =   1500
               Width           =   1875
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
               Left            =   1785
               TabIndex        =   64
               ToolTipText     =   "IVA"
               Top             =   555
               Width           =   1875
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
               Index           =   0
               Left            =   210
               TabIndex        =   63
               Top             =   1537
               Width           =   1425
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
               Left            =   240
               TabIndex        =   62
               Top             =   585
               Width           =   375
            End
            Begin VB.Label Label12 
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
               Left            =   210
               TabIndex        =   61
               Top             =   270
               Width           =   870
            End
            Begin VB.Label lblTotalPagos 
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
               Left            =   1785
               TabIndex        =   60
               ToolTipText     =   "Total de pagos realizados"
               Top             =   1185
               Width           =   1875
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Pagos"
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
               Left            =   210
               TabIndex        =   59
               Top             =   1218
               Width           =   690
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
               Left            =   1785
               TabIndex        =   58
               ToolTipText     =   "Subtotal"
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label lblTotalFactura 
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
               Left            =   1785
               TabIndex        =   57
               ToolTipText     =   "Total de la factura"
               Top             =   870
               Width           =   1875
            End
         End
         Begin VB.Label Label9 
            Caption         =   "Uso del CFDI"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   3000
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Height          =   720
         Left            =   4100
         TabIndex        =   14
         ToolTipText     =   "Siguiente pago"
         Top             =   7750
         Width           =   3230
         Begin VB.CommandButton cmdConfirmartimbre 
            Caption         =   "Confirmar timbre fiscal"
            Height          =   495
            Left            =   1580
            Picture         =   "frmFacturacionMembresiaSocios.frx":0A70
            TabIndex        =   21
            ToolTipText     =   "Confirmar timbre fiscal"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1080
         End
         Begin VB.CommandButton cmdCFD 
            Height          =   495
            Left            =   2660
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmFacturacionMembresiaSocios.frx":0F62
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Comprobante fiscal digital"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   90
            Picture         =   "frmFacturacionMembresiaSocios.frx":1880
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Consulta de facturas"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   585
            Picture         =   "frmFacturacionMembresiaSocios.frx":19F2
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Grabar factura"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   1080
            Picture         =   "frmFacturacionMembresiaSocios.frx":1B64
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Cancelar factura"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   -74865
         TabIndex        =   90
         Top             =   8100
         Width           =   225
      End
      Begin VB.Label Label23 
         Caption         =   "Cancelación rechazada"
         Height          =   255
         Left            =   -74520
         TabIndex        =   93
         Top             =   8390
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         ForeColor       =   &H00FFFFFF&
         Height          =   235
         Left            =   -74875
         TabIndex        =   92
         Top             =   8360
         Width           =   255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74880
         Top             =   8332
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "Pendientes de autorización de cancelación"
         Height          =   255
         Left            =   -74520
         TabIndex        =   91
         Top             =   8095
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   235
         Left            =   -74880
         Top             =   8070
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   -74640
         TabIndex        =   89
         Top             =   0
         Width           =   840
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
         Index           =   8
         Left            =   -75000
         TabIndex        =   88
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   -74640
         TabIndex        =   87
         Top             =   0
         Width           =   840
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
         Index           =   4
         Left            =   -75000
         TabIndex        =   86
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   85
         Top             =   0
         Width           =   840
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
         Index           =   2
         Left            =   -75000
         TabIndex        =   84
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de autorización de cancelación"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -75000
         TabIndex        =   81
         Top             =   0
         Width           =   3060
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Cancelación rechazada"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74640
         TabIndex        =   80
         Top             =   795
         Width           =   1680
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74640
         TabIndex        =   79
         Top             =   15
         Width           =   840
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
         Left            =   -75000
         TabIndex        =   78
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "A"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   -66540
         TabIndex        =   74
         Top             =   7575
         Width           =   135
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Pendientes de timbre fiscal"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   -66300
         TabIndex        =   73
         Top             =   7575
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Canceladas"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   72
         Top             =   7575
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "A"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   3
         Left            =   -74805
         TabIndex        =   71
         Top             =   7825
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "A"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   -74805
         TabIndex        =   70
         Top             =   7575
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pendientes de cancelar ante el SAT"
         Height          =   195
         Index           =   4
         Left            =   -74520
         TabIndex        =   69
         Top             =   7825
         Width           =   2565
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   5
         Left            =   -74880
         Top             =   7815
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   -74880
         Top             =   7560
         Width           =   255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   6
         Left            =   -66600
         Top             =   7560
         Width           =   255
      End
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   225
      Index           =   1
      Left            =   0
      Top             =   18600
      Width           =   255
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      Caption         =   "Canceladas"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   83
      Top             =   0
      Width           =   840
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
      Index           =   0
      Left            =   0
      TabIndex        =   82
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "R. F. C."
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "frmFacturacionMembresiaSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------
' Programa para facturación directa a clientes
' Fecha de desarrollo: Agosto 23, 2006
'--------------------------------------------------------------------------------------------------------

Option Explicit

Private Type TipoPoliza
    lngnumCuenta As Long
    dblCantidad As Double
    intNaturaleza As Integer
End Type

'<grdConceptos>
Const cintColCveConcepto = 1               'Clave del concepto
Const cintColDescripcionConcepto = 2       'Descripción del concepto
Const cintColCantidadConcepto = 3          'Cantidad del concepto
Const cintColIVAConcepto = 4               'Iva del concepto
Const cintColCtaIngresoConcepto = 5

Const cintColsgrdConceptos = 6             '|  Núm. de columnas de <grdConcepto>

Const cstrFormatoConcepto = "||Concepto|Cantidad|IVA|"

'<grdCargos>
Const cintColCveCargo = 1               'Clave del cargo
Const cintColFechaCargo = 2             'Fecha del cargo
Const cintColDescripcionCargo = 3       'Descripción del cargo
Const cintColCantidadCargo = 4          'Cantidad del cargo
Const cintIVACargo = 5                  'Iva del cargo
Const cintColCveConceptoCargo = 6

Const cintColsgrdCargos = 7            '|  Núm. de columnas de <grdConcepto>

Const cstrFormatoCargo = "||Fecha|Otro concepto|Monto|IVA|"

'<grdBusquedaFactura>
Const cintColchrEstatus = 1
Const cintColNumPoliza = 2
Const cIntColNumCorte = 3
Const cIntColFolio = 4
Const cintColNumSocio = 5
Const cIntColRFC = 6
Const cIntColRazonSocial = 7
Const cintColFecha = 8
Const cintColTotalFactura = 9
Const cintColIVAConsulta = 10
Const cintColSubtotal = 11
Const cIntColEstado = 12
Const cintColFacturo = 13
Const cintColCancelo = 14
Const cintColgrdBusquedaFactura = 18 '19 '17            'Núm. de columnas de <grdBusquedaFactura>
Const cintColPCancelarNoSAt = 15
Const cintColPTimbre = 16
Const cintColEstadoNuevoEsquemaCancelacion = 17


Const cstrFormatoBusquedaFactura = "||||Folio|Socio|RFC|Razón social|Fecha|Total|IVA|Subtotal|Estado|Facturó|Canceló"


Const llngColorCanceladas = &HC0&
Const llngColorActivas = &H80000012
Const llnColorPenCancelaSAT = &HC0E0FF
Const llncolorCanceladasSAT = &H80000005

Const cstrCantidad = "#############.00" 'Para formatear a número

Const cintTipoFormato = 9               'Formato para factura directa en <TipoFormato> CC

Dim lstrConceptos As String             'Cadena con los conceptos de factura

Dim ldblDescuentosFactura As Double     'Total de descuentos de la factura
Dim ldblCantidadFactura As Double       'Total de candidad menos descuento
Dim ldblIVAFactura As Double            'Total de IVA de la factura

Dim ldblDescuentoConcepto As Double     'Descuento del concepto
Dim ldblCantidadConcepto As Double      'Cantidad del concepto
Dim ldblIVAConcepto As Double           'IVA del concepto

Dim lblnConsulta As Boolean             'Para saber si se está consultando una factura
Dim lblnEntraCorte As Boolean           'Para saber si la factura entra o no en el corte
Dim llngNumCorte As Long                'Num. de corte en el que se está guardando
Dim llngNumFormaPago As Long            'Num. de forma de pago
Dim lblnCreditoVigente As Boolean       'Indica si el cliente tiene crédito vigente o no
Dim llngNumPoliza As Long               'Num. de póliza
Dim llngNumCtaCliente As Long           'Num. de cuenta contable del cliente
Dim llngPersonaGraba As Long            'Num. de empleado que graba la factura
Dim strSentencia As String              'Usos varios

Dim lstrCalleNumero As String           'Para guardar en la factura
Dim lstrColonia As String               'Para guardar en la factura
Dim lstrCiudad As String                'Para guardar en la factura
Dim llngCveCiudad As Long               'Para guardar en la factura
Dim lstrEstado As String                'Para guardar en la factura
Dim lstrCodigo As String                'Para guardar en la factura

Dim llngFormato As Long                 'Num. del formato de factura para el departamento

Dim apoliza() As TipoPoliza             'Para formar la poliza de la factura

Dim lngCveFormato As Long                   'Para saber el formato que se va a utilizar (relacionado con pvDocumentoDepartamento.intNumFormato)
Dim intTipoEmisionComprobante As Integer    'Variable que compara el tipo de formato y folio a utilizar (0 = Error de formato y folios incompatibles, 1 = Físicos, 2 = Digitales)
Dim strFolio As String                      'Folio de la factura
Dim strSerie As String                      'Serie de la factura
Dim strNumeroAprobacion As String           'Número de aprobación del folio
Dim strAnoAprobacion As String              'Año de aprobación del folio
Dim vgConsecutivoMuestraPvFactura As Long   'Consecutivo de la tabla PvFactura al momento de seleccionar un registro del grid en pMostrar
Dim vlstrTipoCFD As String
Dim intTipoCFDFactura As Integer        'Variable que regresa el tipo de CFD de la factura(0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)


Dim lngCuentaCuotasPorCobrar As Long        '|  Cuenta puente de las cuotas por cobrar de socios
Dim lngCuentaCuotasPorDevengar As Long      '|  Cuenta puente de las cuotas por devengar de socios

Dim aFormasPago() As FormasPago
Dim vllngSeleccionadas As Long
Dim vllngSeleccPendienteTimbre As Long
Dim lngConsecutivoFactura As Long
Dim blnActivarconsulta As Boolean
Dim strVersionCFDISocios As String

Private Sub pCargaFolio(intAumenta As Integer)
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
    lblFolio.Caption = Trim(strSerie) + Trim(strFolio)
    
    If vllngFoliosFaltantes > 0 Then
        MsgBox "Faltan " & Trim(Str(vllngFoliosFaltantes)) + " facturas y será necesario aumentar folios!", vbOKOnly + vbInformation, "Mensaje"
    End If

    'Habilitar el chkBitExtranjero si el folio es de tipo digital
    If Trim(strNumeroAprobacion) <> "" And Trim(strAnoAprobacion) <> "" Then
        chkBitExtranjero.Enabled = True
    End If
End Sub



Private Sub ChkFacturasCancelaNoSAT_Click()
    Me.mskFechaIni.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    Me.mskFechaFin.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    Me.txtBusquedaSocio.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    Me.lblBusquedaNombreSocio.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    Me.ChkPendientesTimbre.Value = vbUnchecked
    Me.ChkPendientesTimbre.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    Me.cmdCargar.Enabled = Not ChkFacturasCancelaNoSAT.Value = vbChecked
    If blnActivarconsulta Then cmdCargar_Click
End Sub
Private Sub chkIncluyePagosFacturados_Click()
    pLlenaPagos
End Sub
Private Sub ChkPendientesTimbre_Click()
    Me.mskFechaIni.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    Me.mskFechaFin.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    Me.txtBusquedaSocio.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    Me.lblBusquedaNombreSocio.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    Me.ChkFacturasCancelaNoSAT.Value = vbUnchecked
    Me.ChkFacturasCancelaNoSAT.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    Me.cmdCargar.Enabled = Not ChkPendientesTimbre.Value = vbChecked
    If blnActivarconsulta Then cmdCargar_Click
End Sub

Private Sub cmdCancelaFacturasSAT_Click()
'Cancelacion maciva de facturas ante el SAT, cancelacion del XML
Dim ArrIdFacturas() As String
Dim vlLngCantidadFacturas As Long
Dim vlLngCont As Long
Dim vllngPersonaGraba As Long

On Error GoTo NotificaError

If MsgBox(SIHOMsg(1249), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
   If vllngPersonaGraba = 0 Then Exit Sub
     'Recorremos el grid para poder cargar el arreglo con los Id de las facturas que vamos a cancelar
      With grdBusquedaFactura
           vlLngCantidadFacturas = 0
           Erase ArrIdFacturas
           ReDim ArrIdFacturas(3, 0)
           For vlLngCont = 1 To .Rows - 1
               If .TextMatrix(vlLngCont, 0) = "*" And .TextMatrix(vlLngCont, cintColPCancelarNoSAt) > 0 Then
                   vlLngCantidadFacturas = vlLngCantidadFacturas + 1
                   ReDim Preserve ArrIdFacturas(3, vlLngCantidadFacturas)
                   ArrIdFacturas(1, vlLngCantidadFacturas) = .RowData(vlLngCont)
                   ArrIdFacturas(2, vlLngCantidadFacturas) = "FA"
                   ArrIdFacturas(3, vlLngCantidadFacturas) = 1
               End If
           Next vlLngCont
      End With
      'enviamos el arreglo a cancelacion'''''''''''''
'| Comentado Temporalmente para compilar      pCancelaCFDiMasivo ArrIdFacturas, vlLngCantidadFacturas, "frmFacturacionMembresiaSocios", vllngPersonaGraba
      cmdCargar_Click
      grdBusquedaFactura.SetFocus
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCancelaFacturasSAT_Click"))
    Unload Me
    
    
End Sub

Private Sub cmdCargar_Click()
    Dim rs As New ADODB.Recordset
    Dim lngColor As Long
    Dim lngcolorSub As Long
           
    If fblnDatosValidosBusqueda() Then
        pLimpiagrdBusquedaFactura
        pConfiguragrdBusquedaFactura
        vllngSeleccionadas = 0
        vllngSeleccPendienteTimbre = 0
        vgstrParametrosSP = ""
        vgstrParametrosSP = _
                            "-1" & _
                            "|" & IIf((ChkFacturasCancelaNoSAT.Value = vbChecked Or Me.ChkPendientesTimbre.Value = vbChecked), fstrFechaSQL("01/01/2010"), fstrFechaSQL(mskFechaIni.Text)) & _
                            "|" & IIf((ChkFacturasCancelaNoSAT.Value = vbChecked Or Me.ChkPendientesTimbre.Value = vbChecked), fstrFechaSQL(fdtmServerFecha), fstrFechaSQL(mskFechaFin.Text)) & _
                            "|1" & _
                            "|" & "-1" & _
                            "|" & "S|" & _
                             IIf((Me.ChkFacturasCancelaNoSAT.Value = vbChecked Or Me.ChkPendientesTimbre.Value = vbChecked), "-1", IIf(Trim(txtBusquedaSocio.Text) = "", "-1", flngObtieneClaveSocio(txtBusquedaSocio.Text))) & _
                            "|" & CStr(vgintNumeroDepartamento) & _
                            "|" & CLng(vgintClaveEmpresaContable) & _
                            "|" & IIf(Me.ChkFacturasCancelaNoSAT.Value = vbChecked, 1, 0) & _
                            "|" & IIf(ChkPendientesTimbre.Value = vbChecked, 1, 0) & _
                            "|" & "0" & _
                            "|" & "0"
                             
        
        Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFacturaFiltro_NE")
        If rs.RecordCount <> 0 Then
            With grdBusquedaFactura
                .Visible = False
                Do While Not rs.EOF
                    .Row = .Rows - 1
                    If rs!chrEstatus = "C" Then
                        lngColor = llngColorCanceladas
                        lngcolorSub = &HFFFFFF '| Blanco
                        .TextMatrix(.Row, cintColEstadoNuevoEsquemaCancelacion) = rs!PendienteCancelarSAT_NE
                    Else
                        'lngColor = llngColorActivas
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
                        lngcolorSub = &H80FFFF
                        .TextMatrix(.Row, 0) = "*"
                        vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                       Else
                        'lngcolorSub = llncolorCanceladasSAT
                       End If
                    End If
                    
                     .TextMatrix(.Row, cintColPTimbre) = rs!PendienteTimbreFiscal
                      .TextMatrix(.Row, cintColPCancelarNoSAt) = rs!PendienteCancelarSat
                    .Col = cIntColNumCorte
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cIntColNumCorte) = rs!NumCorte
                    .Col = cintColNumPoliza
                    .CellForeColor = lngColor
                    .CellBackColor = lngcolorSub
                    .TextMatrix(.Row, cintColNumPoliza) = rs!NumPoliza
                    .Col = cintColchrEstatus
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColchrEstatus) = rs!chrEstatus
                    .Col = cintColFecha
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColFecha) = Format(rs!fecha, "dd/mmm/yyyy")
                    .Col = cIntColFolio
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cIntColFolio) = rs!Folio
                    .Col = cintColNumSocio
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColNumSocio) = rs!Paciente
                    .Col = cIntColRazonSocial
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cIntColRazonSocial) = rs!RazonSocial
                    .Col = cIntColRFC
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cIntColRFC) = rs!RFC
                    .Col = cintColTotalFactura
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColTotalFactura) = FormatCurrency(rs!TotalFactura, 2)
                    .Col = cintColIVAConsulta
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColIVAConsulta) = FormatCurrency(rs!IVA, 2)
                    .Col = cintColSubtotal
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColSubtotal) = FormatCurrency(rs!Subtotal, 2)
                    .Col = cIntColEstado
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cIntColEstado) = rs!Estado
                    .Col = cintColFacturo
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColFacturo) = rs!PersonaFacturo
                    .Col = cintColCancelo
                     .CellBackColor = lngcolorSub
                    .CellForeColor = lngColor
                    .TextMatrix(.Row, cintColCancelo) = rs!PersonaCancelo
                    .RowData(.Row) = rs!IdFactura
                    .Rows = .Rows + 1
                    rs.MoveNext
                Loop
                .Rows = .Rows - 1
                .Visible = True
            End With
            grdBusquedaFactura.Col = cintColFecha
            grdBusquedaFactura.Row = 1
        Else
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
    End If
    Me.cmdCancelaFacturasSAT.Enabled = vllngSeleccionadas > 0
    Me.cmdconfirmartimbrefiscal.Enabled = vllngSeleccPendienteTimbre > 0
End Sub

Private Function fblnDatosValidosBusqueda() As Boolean
    fblnDatosValidosBusqueda = True
    If Not IsDate(mskFechaIni.Text) Then
        fblnDatosValidosBusqueda = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaIni.SetFocus
    End If
    If fblnDatosValidosBusqueda And Not IsDate(mskFechaFin.Text) Then
        fblnDatosValidosBusqueda = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        mskFechaFin.SetFocus
    End If
    If fblnDatosValidosBusqueda Then
        If CDate(mskFechaFin.Text) < CDate(mskFechaIni.Text) Then
            fblnDatosValidosBusqueda = False
            '¡Rango incorrecto!
            MsgBox SIHOMsg(26), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaIni.SetFocus
        End If
    End If
End Function

Private Sub cmdCFD_Click()
On Error GoTo NotificaError

    If vlstrTipoCFD = "CFD" Then
        frmComprobanteFiscalDigital.lngComprobante = vgConsecutivoMuestraPvFactura
        frmComprobanteFiscalDigital.strTipoComprobante = "FA"
        frmComprobanteFiscalDigital.Show vbModal, Me
    ElseIf vlstrTipoCFD = "CFDi" Then
        frmComprobanteFiscalDigitalInternet.lngComprobante = vgConsecutivoMuestraPvFactura
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = "FA"
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCFD_Click"))
    Unload Me
End Sub

Private Function fintErrorCancelar() As Integer
    Dim rs As New ADODB.Recordset
    Dim rsPagos As New ADODB.Recordset
    
    fintErrorCancelar = 0
    
    'que la factura aún esté activa
    Set rs = frsEjecuta_SP(Trim(lblFolio.Caption), "sp_PvSelFactura")
    If rs.RecordCount <> 0 Then
        If rs!chrEstatus = "C" Then
            'La información ha cambiado, consulte de nuevo.
            fintErrorCancelar = 381
            Exit Function
        End If
    End If
    rs.Close
    'que el o los créditos de la factura no tengan pagos registrados
    vgstrParametrosSP = fstrFechaSQL(fdtmServerFecha) & _
                        "|" & fstrFechaSQL(fdtmServerFecha) & _
                        "|" & txtClaveUnica.Text & _
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
                End If
            End If
            rs.MoveNext
        Loop
        If fintErrorCancelar <> 0 Then
            Exit Function
        End If
    Else
        '¡La factura ya fue pagada!
        fintErrorCancelar = 964
        Exit Function
    End If
    'si la factura tiene poliza directa, que el periodo contable esté abierto para cancelarla
    If Val(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColNumPoliza)) <> 0 Then
        fintErrorCancelar = fintErrorContable(mskFecha.Text)
    End If
    'si la factura entró en un corte, como se registrará la cancelación en el corte actual, validar que exista uno abiero
    If fintErrorCancelar = 0 And Val(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColNumPoliza)) = 0 Then
        llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
        If llngNumCorte = 0 Then
            fintErrorCancelar = 659 'No se encontró un corte abierto.
            Exit Function
        End If
    End If
    'si la factura entró en un corte, bloquearlo para registrar la cancelación
    If Val(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColNumPoliza)) = 0 Then
        fintErrorCancelar = fintErrorBloqueoCorte()
    End If
End Function

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
      lblTextoBarraCFD.Caption = "Confirmando timbre fiscal para la factura, por favor espere..."
      freBarraCFD.Visible = True
      freBarraCFD.Refresh
      frmFacturacionMembresiaSocios.Enabled = False
      pLogTimbrado 2
      blnNOMensajeErrorPAC = True
      EntornoSIHO.ConeccionSIHO.BeginTrans
      vlngReg = flngRegistroFolio("FA", lngConsecutivoFactura)
      If Not fblnGeneraComprobanteDigital(lngConsecutivoFactura, "FA", 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
         On Error Resume Next
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
             'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
               MsgBox Replace(SIHOMsg(1314), "<FOLIO>", Trim(lblFolio.Caption)), vbInformation + vbOKOnly, "Mensaje"
              
             'la factura se queda igual, no se hace nada
          ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
              'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
              MsgBox Replace(SIHOMsg(1313), "<FOLIO> ", Trim(lblFolio.Caption)), vbExclamation + vbOKOnly, "Mensaje"
              
              'Aqui se debe de cancelar la factura
              pCancelarFactura Trim(lblFolio.Caption), vllngPersonaGraba, Me.Name
              'se carga de nuevo la factura
                  
          End If
      Else
          'Se guarda el LOG
           Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & lblFolio.Caption)
          'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
           pEliminaPendientesTimbre lngConsecutivoFactura, "FA"
          'Commit
           EntornoSIHO.ConeccionSIHO.CommitTrans
          'Timbre fiscal de factura <FOLIO>: Confirmado.
           MsgBox Replace(SIHOMsg(1315), " <FOLIO>", Trim(lblFolio.Caption)), vbInformation + vbOKOnly, "Mensaje"

      End If
                  
'     'Barra de progreso CFDi
       pgbBarraCFD.Value = 100
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbDefault
       freBarraCFD.Visible = False
       frmFacturacionMembresiaSocios.Enabled = True
       blnNOMensajeErrorPAC = False
       pLogTimbrado 1
       If vgIntBanderaTImbradoPendiente <> 1 Then
          pReinicia
          txtClaveUnica.SetFocus
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
      With grdBusquedaFactura
           
           For vlLngCont = 1 To .Rows - 1
               If .TextMatrix(vlLngCont, 0) = "*" And .TextMatrix(vlLngCont, cintColPTimbre) = 1 Then
                  pgbBarraCFD.Value = 70
                  freBarraCFD.Top = 3200
                  Screen.MousePointer = vbHourglass
                  lblTextoBarraCFD.Caption = "Confirmando timbre fiscal para la factura, por favor espere..."
                  freBarraCFD.Visible = True
                  freBarraCFD.Refresh
                  frmFacturacionMembresiaSocios.Enabled = False
                  pLogTimbrado 2
                  blnNOMensajeErrorPAC = True
                  EntornoSIHO.ConeccionSIHO.BeginTrans
                                                           
                  vlngReg = flngRegistroFolio("FA", .RowData(vlLngCont))
                  If Not fblnGeneraComprobanteDigital(.RowData(vlLngCont), "FA", 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
                     
                       On Error Resume Next
                       EntornoSIHO.ConeccionSIHO.RollbackTrans
                       If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
                          'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
                          MsgBox Replace(SIHOMsg(1314), "<FOLIO>", Trim(.TextMatrix(vlLngCont, cIntColFolio))), vbInformation + vbOKOnly, "Mensaje"
                          'la factura se queda igual, no se hace nada
                       ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
                          'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
                          MsgBox Replace(SIHOMsg(1313), "<FOLIO>", Trim(.TextMatrix(vlLngCont, cIntColFolio))), vbExclamation + vbOKOnly, "Mensaje"
                          'Aqui se debe de cancelar la factura
                          pCancelarFactura Trim(.TextMatrix(vlLngCont, 1)), vllngPersonaGraba, Me.Name
                       End If
                  Else
                      'Se guarda el LOG
                       Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura" & .TextMatrix(vlLngCont, 1))
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
                    frmFacturacionMembresiaSocios.Enabled = True
                    pLogTimbrado 1
               End If
           Next vlLngCont
      End With
      
      blnNOMensajeErrorPAC = False
      cmdCargar_Click
      grdBusquedaFactura.SetFocus
      
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    Dim vllngPersonaGraba As Long
    Dim vllngMensaje As Long
    Dim vllngNumeroCorte As Long
    Dim vllngCorteGrabando As Long
    Dim vllngNumCorteFactura As Long
    Dim vlstrSentencia As String
    Dim rsDC As New ADODB.Recordset
    Dim rsPvDetalleCorte As New ADODB.Recordset
    Dim vldblTotalIVACredito As Double
    Dim vlrsTemp As New ADODB.Recordset
    Dim vlrsCorte As New ADODB.Recordset
    Dim vlintPolizaFactura As Long
    
    '------------------------------------------------------------------
    ' Persona que graba
    '------------------------------------------------------------------
    If Not fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 2417, 609), "E") Then Exit Sub
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub

    If Not fblnCancelaCFDi(lngConsecutivoFactura, "FA") Then
       'EntornoSIHO.ConeccionSIHO.RollbackTrans
       If vlstrMensajeErrorCancelacionCFDi <> "" Then MsgBox vlstrMensajeErrorCancelacionCFDi, vbOKOnly + vbCritical, "Mensaje"
       Exit Sub
    End If
        
    '------------------------------------------------------------------
    '-Inicio de Transacción
    '------------------------------------------------------------------
    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    '|  Pone la fecha actual como fecha de cancelación en GnComprobanteFiscalDigital
    frsEjecuta_SP lngConsecutivoFactura & "|FA" & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|'" & vgMotivoCancelacion & "'", "SP_GNUPDCANCELACOMPROBANTEFIS"

    '------------------------------------------------------------------
    '-Obtener el numero de corte actual
    '------------------------------------------------------------------
    vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If vllngMensaje <> 0 Then
        MsgBox SIHOMsg(Str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje" 'Que el corte debe ser cerrado por cambio de día, Que no existe corte abierto
        Exit Sub
    End If

    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    vllngCorteGrabando = 1
    frsEjecuta_SP vllngNumeroCorte & "|Grabando", "Sp_PvUpdEstatusCorte", True, vllngCorteGrabando
        
    If vllngCorteGrabando = 2 Then               '(2)
        vllngNumCorteFactura = frsRegresaRs("select intNumCorte from PvFactura where ltrim(rtrim(chrFolioFactura))='" & Trim(lblFolio.Caption) & "'").Fields(0)
        '------------------------------------------------------------------
        ' Generar registros al reves en PVDetalleCorte para cancelar
        '------------------------------------------------------------------
        vlstrSentencia = "select * from pvdetallecorte " & _
                         " where chrFolioDocumento = '" & Trim(lblFolio.Caption) & "'" & _
                         " and chrTipoDocumento = 'FA'"
        Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)           'RS de consulta
        Set rsPvDetalleCorte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic) 'RS tipo Tabla
        If rsDC.RecordCount > 0 Then
            With rsPvDetalleCorte
                Do While Not rsDC.EOF
                    .AddNew
                    !intnumcorte = vllngNumeroCorte
                    !dtmFechahora = fdtmServerFecha + fdtmServerHora
                    !chrFolioDocumento = rsDC!chrFolioDocumento
                    !chrTipoDocumento = rsDC!chrTipoDocumento
                    !intFormaPago = rsDC!intFormaPago
                    !mnyCantidadPagada = rsDC!mnyCantidadPagada * -1  'Cantidad Negativa
                    !MNYTIPOCAMBIO = rsDC!MNYTIPOCAMBIO
                    !intfoliocheque = rsDC!intfoliocheque
                    !intNumCorteDocumento = rsDC!intNumCorteDocumento
                    .Update
                    rsDC.MoveNext
                Loop
            End With
        End If
        vlstrSentencia = "Select distinct  chrFolioDocumento, chrTipoDocumento, intFormaPago, " & _
                        " mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                        " intNumCorteDocumento   " & _
                        " From pvDetalleCorte " & _
                        " Where chrFolioDocumento in (Select chrFolioRecibo " & _
                                                        " from pvPago " & _
                                                        " Where chrFolioFactura = '" & Trim(lblFolio.Caption) & "'" & _
                                                        " and smidepartamento = (" & _
                                                        " select smicvedepartamento from nodepartamento where " & _
                                                        " tnyclaveempresa = " & vgintClaveEmpresaContable & " and smicvedepartamento = " & vgintNumeroDepartamento & ")) " & _
                        " And mnyCantidadPagada > 0 " & _
                        " And chrTipoDocumento = 'RE' "

        vlstrSentencia = vlstrSentencia & _
                        " UNION Select distinct  chrFolioDocumento, chrTipoDocumento, intFormaPago, " & _
                        " mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                        " intNumCorteDocumento " & _
                        " From pvDetalleCorte " & _
                        " Where chrFolioDocumento in (Select chrFolioRecibo " & _
                                                        " From pvSalidaDinero " & _
                                                        " Where chrFolioFactura = '" & Trim(lblFolio.Caption) & "' " & _
                                                        " and smidepartamento = (" & _
                                                    " select smicvedepartamento from nodepartamento where " & _
                                                    " tnyclaveempresa = " & vgintClaveEmpresaContable & " and smicvedepartamento = " & vgintNumeroDepartamento & ")) " & _
                        " And mnyCantidadPagada > 0 " & _
                        " And chrTipoDocumento = 'SD' "
            
    End If
    Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly) 'RS de consulta
    vldblTotalIVACredito = 0
    If rsDC.RecordCount > 0 Then
        With rsPvDetalleCorte
            Do While Not rsDC.EOF
                'Aqui se obtiene el total del Iva a crédito para poder hacer
                Set vlrsTemp = frsRegresaRs("Select pvformapago.chrTipo From pvformapago Where pvformapago.INTFORMAPAGO = " & !intFormaPago, adLockOptimistic, adOpenForwardOnly)
                If vlrsTemp!chrTipo = "C" Then vldblTotalIVACredito = vldblTotalIVACredito + rsDC!mnyCantidadPagada
                .AddNew
                !intnumcorte = vllngNumeroCorte
                !dtmFechahora = fdtmServerFecha + fdtmServerHora
                !chrFolioDocumento = rsDC!chrFolioDocumento
                !chrTipoDocumento = rsDC!chrTipoDocumento
                !intFormaPago = rsDC!intFormaPago
                !mnyCantidadPagada = rsDC!mnyCantidadPagada
                !MNYTIPOCAMBIO = rsDC!MNYTIPOCAMBIO
                !intfoliocheque = rsDC!intfoliocheque
                !intNumCorteDocumento = rsDC!intNumCorteDocumento
                .Update
                rsDC.MoveNext
            Loop
        End With
        vlrsTemp.Close
    End If

    ' Cancelacion de la factura
    vlstrSentencia = "UPDATE pvFactura set chrEstatus = 'C' WHERE chrFolioFactura = '" & Trim(lblFolio.Caption) & "'"
    pEjecutaSentencia (vlstrSentencia)

    '------------------------------------------------------------------
    ' Guardo en documentos cancelados
    '------------------------------------------------------------------
    vlstrSentencia = "insert into PVDocumentoCancelado values('" & Trim(lblFolio.Caption) & "','FA'," & _
                    Trim(Str(vgintNumeroDepartamento)) & "," & Trim(Str(vllngPersonaGraba)) & ",getdate())"
    pEjecutaSentencia (vlstrSentencia)
    
    '--------------------------------------------------------------------
    '|  Hace los movimientos inversos de los cargos del socio
    '--------------------------------------------------------------------
    vlstrSentencia = "Select INTNUMCORTE From PVCORTEPOLIZA Where Trim(PVCORTEPOLIZA.CHRFOLIODOCUMENTO) = Trim('" & lblFolio.Caption & "') AND TRIM(PVCORTEPOLIZA.CHRTIPODOCUMENTO) = 'FA'"
    Set vlrsCorte = frsRegresaRs(vlstrSentencia)
    vlintPolizaFactura = 0
    If Not vlrsCorte.EOF Then
        vlintPolizaFactura = IIf(IsNull(vlrsCorte!intnumcorte), 0, vlrsCorte!intnumcorte)
    End If
    
    If vlintPolizaFactura = vllngNumeroCorte Then
        '|  Abono a la cuenta por cobrar de cuotas de socios
        pInsCortePoliza vllngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorCobrar, CDbl(Format(lblTotalFactura.Caption, cstrCantidad)) * -1, 0
        '|  Cargo a la cuenta por devengar de cuotas de socios
        pInsCortePoliza vllngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorDevengar, CDbl(Format(lblTotalFactura.Caption, cstrCantidad)) * -1, 1
    Else
        '|  Abono a la cuenta por cobrar de cuotas de socios
        pInsCortePoliza vllngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorCobrar, CDbl(Format(lblTotalFactura.Caption, cstrCantidad)), 1
        '|  Cargo a la cuenta por devengar de cuotas de socios
        pInsCortePoliza vllngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorDevengar, CDbl(Format(lblTotalFactura.Caption, cstrCantidad)), 0
    End If
            
'''    '------------------------------------------------------------------
'''    ' Quitar el estatus de "Facturada" a la cuenta del paciente
'''    '------------------------------------------------------------------
'''    If Not optTipoPaciente(2).Value And Not optTipoPaciente(3).Value Then
'''        'Factura de una cuenta:
'''        vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & "0"
'''        frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCUENTAFACTURADA"
'''    Else
'''        If optTipoPaciente(2).Value Then
'''            'Factura de un grupo
'''            Set rs = frsEjecuta_SP(txtMovimientoPaciente.Text, "SP_PVSELCUENTAGRUPO")
'''            Do While Not rs.EOF
'''                vgstrParametrosSP = rs!intMovPaciente & "|" & rs!chrtipopaciente & "|" & "0"
'''                frsEjecuta_SP vgstrParametrosSP, "SP_PVUPDCUENTAFACTURADA"
'''                rs.MoveNext
'''            Loop
'''        End If
'''    End If
             
    rsPvDetalleCorte.Close
    rsDC.Close
    pLiberaCorte vllngNumeroCorte

'***************************************************************************************
'***************************************************************************************
'***************************************************************************************

    '------------------------------------------------------------------
    ' Quito el numero de factura del cargo, para que los pueda borrar
    '------------------------------------------------------------------
    vlstrSentencia = "update PvCargo set chrFolioFactura = null where chrFolioFactura = '" & RTrim(lblFolio.Caption) & "'"
    pEjecutaSentencia vlstrSentencia
        
    '------------------------------------------------------------------------------------'
    ' Genera todos los movimientos de la factura en la poliza con el siguiente criterio: '
    '   » Si el Corte en el que se realizó la factura y el corte actual son IGUALES      '
    '       » Si es un Cargo. Cantidad = Cantidad * -1 y Tipo Movimiento = Cargo         '
    '       » Si es un Abono. Cantidad = Cantidad * -1 y Tipo Movimiento = Abono         '
    '   » Si el Corte en el que se realizó la factura y el corte actual son DIFERENTES   '
    '       » Si es un Cargo. Cantidad = Cantidad y Tipo Movimiento = Abono              '
    '       » Si es un Abono. Cantidad = Cantidad y Tipo Movimiento = Cargo              '
    '------------------------------------------------------------------------------------'
    vgstrParametrosSP = CStr(vllngNumeroCorte) & "|" & Trim(lblFolio.Caption)
    frsEjecuta_SP vgstrParametrosSP, "sp_PvInsPolizaCancelaFactura"
        
    '------------------------------------------------------------------
    ' Quitar el cancelado de PAGOS (BitCancelado) y Quitar el numero de factura
    '------------------------------------------------------------------
    vlstrSentencia = "Update PvPago set chrFolioFactura = NULL, bitCancelado = 0 " & _
                " where chrFolioFactura = '" & Trim(lblFolio.Caption) & "'"
    pEjecutaSentencia (vlstrSentencia)
    
    '------------------------------------------------------------------
    ' Quitar el cancelado de DEVOLUCIONES (BitCancelado) y Quitar el numero de factura
    '------------------------------------------------------------------
    vlstrSentencia = "Update PvSalidaDinero set chrFolioFactura = NULL, bitCancelado = 0 " & _
                " where chrFolioFactura = '" & Trim(lblFolio.Caption) & "'"
    pEjecutaSentencia (vlstrSentencia)

    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "CANCELACION DE FACTURA DE MEMBRESIA SOCIO", lblFolio.Caption)
    
    '------------------------------------------------------------------
    ' Darle COMMIT a la TRANSACTION
    '------------------------------------------------------------------
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    'La factura se canceló satisfactoriamente.
    MsgBox SIHOMsg(365), vbInformation, "Mensaje"
    
    pReinicia
    txtClaveUnica.SetFocus
End Sub

Private Sub cmdLocate_Click()
    '| Inicializa búsqueda
    pLimpiagrdBusquedaFactura
    pConfiguragrdBusquedaFactura
    blnActivarconsulta = False
    Me.ChkFacturasCancelaNoSAT.Value = vbUnchecked
    Me.ChkPendientesTimbre.Value = vbUnchecked
    blnActivarconsulta = True
    mskFechaIni.Mask = ""
    mskFechaIni.Text = fdtmServerFecha
    mskFechaIni.Mask = "##/##/####"
    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    txtBusquedaSocio.Text = ""
    lblBusquedaNombreSocio.Caption = ""
    SSTFactura.Tab = 1
    mskFechaIni.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim intError As Integer         'Error en transacción
    Dim clsFacturaDirecta As clsFactura
    Dim rsFactura As New ADODB.Recordset
    Dim strTotalLetras As String
    Dim lngidfactura As Long
    Dim strSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vllngPvFacturaConsecutivo As Long
    Dim vlblnbandera As Boolean
    
    '|  Datos fiscales
    Dim strNombreFactura As String
    Dim strDireccion As String
    Dim strCodigoPostal As String
    Dim strColonia As String
    Dim strNumeroExterior As String
    Dim strNumeroInterior As String
    Dim bitExtranjero As Integer
    Dim strTelefono As String
    Dim strRFC As String
    Dim strCalleNumero As String
    Dim lngCveCiudad As Long
    
    Dim vlstrFolioDocumento As String
    Dim vllngFoliosFaltantes As Long
    Dim alstrParametrosSalida() As String
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    Dim vllngNumeroCorte As Long
    Dim vllngCorteUsado As Long
    Dim intUsoCFDI As Long
    
On Error GoTo NotificaError

    If Not fblnDatosValidos() Then Exit Sub
    
'*********************************** OPCIONES AGREGADAS PARA CFD'S ************************************
    
    'Se valida en caso de no haber formato activo mostrar mensaje y cancelar transacción
    If llngFormato = 0 Then
        'No se encontró un formato válido de factura.
        MsgBox SIHOMsg(373), vbCritical, "Mensaje"
        pReinicia
        Exit Sub
    End If
    
    'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
    '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
    intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", llngFormato)
    
    'Si los folios y los formatos no son compatibles...
    If intTipoEmisionComprobante = 0 Then   'ERROR
        'Si es error, se cancela la transacción
        Exit Sub
    End If
    
    If intTipoEmisionComprobante = 2 Then
        'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
        intTipoCFDFactura = fintTipoCFD("FA", llngFormato)
        
        'Si aparece un error terminar la transacción
        If intTipoCFDFactura = 3 Then   'ERROR
            'Si es error, se cancela la transacción
            Exit Sub
        End If
    End If
    
      If cboUsoCFDI.ListIndex = -1 Then
            MsgBox "Seleccione el uso del comprobante", vbExclamation, "Mensaje"
            cboUsoCFDI.SetFocus
            Exit Sub
        End If
    
    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If llngPersonaGraba = 0 Then Exit Sub
    
    '--------------------------------------------------------
    '                      Formas de pago
    '--------------------------------------------------------
    vlblnbandera = False
    If Val(Format(lblTotal.Caption, "")) > 0 Then
       vlblnbandera = fblnFormasPagoPos(aFormasPago(), Val(Format(lblTotal.Caption, "")), True, 1, False, -1, "CO", Trim(Replace(Replace(Replace(txtRFC.Text, "-", ""), "_", ""), " ", "")))
    Else
        If Val(Format(lblTotal.Caption, "")) < 0 Then
            MsgBox SIHOMsg(369), vbCritical, "Mensaje"
            Exit Sub
        End If
    End If
    '|  Le puse esta condición para que deje facturar en cero. - DM -
    If Val(Format(lblTotal.Caption, "")) <> 0 Then
        If Not vlblnbandera Then Exit Sub  'Para que ya no haga nada
    End If
    
    '------------------------------------------------------------------
    '               Carga los datos fiscales del socio
    '------------------------------------------------------------------
    strNombreFactura = Trim(lblSocio.Caption)
    strDireccion = Trim(lblDomicilio.Caption)
    strNumeroExterior = Trim(lblNumeroExterior.Caption)
    strNumeroInterior = Trim(lblNumeroInterior.Caption)
    bitExtranjero = IIf(chkBitExtranjero.Value, 1, 0)
    strTelefono = Trim(lblTelefono.Caption)
    If chkBitExtranjero.Value = 1 Then
        strRFC = "XEXX010101000"
    Else
        strRFC = IIf(Len(fStrRFCValido(txtRFC.Text)) < 12 Or Len(fStrRFCValido(txtRFC.Text)) > 13, "XAXX010101000", fStrRFCValido(txtRFC.Text))
    End If
    lngCveCiudad = llngCveCiudad
    strCodigoPostal = Trim(lblCP.Caption)
    strColonia = Trim(lblColonia.Caption)
    strCalleNumero = Trim(strDireccion) & " " & Trim(strNumeroExterior) & " " & IIf(Trim(strNumeroInterior) = "", "", " Int. " & strNumeroInterior)
    
    '------------------------------------------------------------------
    '                     Otros datos fiscales
    '------------------------------------------------------------------
    If chkOtrosDatosFiscales.Value = 1 Then
        Load frmDatosFiscales
        frmDatosFiscales.sstDatos.Tab = 1
        frmDatosFiscales.Show vbModal
        With frmDatosFiscales
            strNombreFactura = .vgstrNombre
            strDireccion = .vgstrDireccion
            strNumeroExterior = .vgstrNumExterior
            strNumeroInterior = .vgstrNumInterior
            bitExtranjero = .vgBitExtranjero
            strTelefono = .vgstrTelefono
            strRFC = fStrRFCValido(.vgstrRFC)
            lngCveCiudad = .llngCveCiudad
            strCodigoPostal = .vgstrCP
            strColonia = .vgstrColonia
            strCalleNumero = Trim(strDireccion) & " " & Trim(strNumeroExterior) & " " & IIf(Trim(strNumeroInterior) = "", "", " Int. " & strNumeroInterior)
        End With
        Unload frmDatosFiscales
        Set frmDatosFiscales = Nothing
        If Trim(strRFC) = "" Or Trim(strNombreFactura) = "" Then Exit Sub
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '------------------------------------------------------------------
    '-Número de la factura-
    '------------------------------------------------------------------
    vllngFoliosFaltantes = 0
    
    pCargaArreglo alstrParametrosSalida, vllngFoliosFaltantes & "|" & ADODB.adBSTR & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
    frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "sp_gnFolios", , , alstrParametrosSalida
    pObtieneValores alstrParametrosSalida, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
    '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
    vlstrFolioDocumento = Trim(strSerie) & strFolio
    If Trim(vlstrFolioDocumento) = "0" Then
        EntornoSIHO.ConeccionSIHO.RollbackTrans
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        Exit Sub
    End If
    lblFolio.Caption = vlstrFolioDocumento
    
    '------------------------------------------------------------------
    '- Número del corte actual
    '------------------------------------------------------------------
    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    
    '------------------------------------------------------------------
    '  Inicializa la póliza
    '------------------------------------------------------------------
    ReDim apoliza(0)
    apoliza(0).lngnumCuenta = 0

    Set clsFacturaDirecta = New clsFactura
    '---------------------------------------------------
    '1.- Guardar la factura
      If cboUsoCFDI.ListIndex > -1 Then
                intUsoCFDI = cboUsoCFDI.ItemData(cboUsoCFDI.ListIndex)
      Else
                intUsoCFDI = 0
      End If
      
    lngidfactura = clsFacturaDirecta.flngInsFactura(vlstrFolioDocumento, (CDate(mskFecha.Text) + fdtmServerHora), strRFC, strNombreFactura, strDireccion, strNumeroExterior, strNumeroInterior, Val(Format(lblIVA.Caption, cstrCantidad)), 0, " ", CLng(txtClaveSocio.Text), "S", vgintNumeroDepartamento, llngPersonaGraba, vllngNumeroCorte, Val(Format(lblTotalPagos.Caption, cstrCantidad)), Val(Format(lblTotalFactura.Caption, cstrCantidad)), 1, 1, strTelefono, "S", 0, 0, 0, strCalleNumero, strColonia, " ", " ", strCodigoPostal, glngCveImpuesto, lngCveCiudad, strFolio, strSerie, intUsoCFDI)
    '---------------------------------------------------
    '2.- Guardar del detalle de la factura
    pGuardaDetalleFactura (lngidfactura)
    '---------------------------------------------------
    '2.1.- Inicializamos el arreglo del corte
    pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
    '--------------------------------------------------
    '3.- Cancelación de los pagos realizados
    pCancelaPagos vlstrFolioDocumento, vllngNumeroCorte
    '---------------------------------------------------
    '4.- Guardar la factura en el corte
    pGuardaFacturaCorte vllngNumeroCorte
    '---------------------------------------------------
    '5.- Guardar la póliza en el corte
    pGuardaPolizaCorte vllngNumeroCorte
    '---------------------------------------------------
    '6.- Pone el folio de la factura a los cargos
    pPoneFolioFacturaACargos
    '---------------------------------------------------
    '7.- Liberar el corte
    vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte, True)
    If vllngCorteUsado = 0 Then
       EntornoSIHO.ConeccionSIHO.RollbackTrans
       'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
       MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
       Exit Sub
    Else
       If vllngCorteUsado <> vllngNumeroCorte And vllngCorteUsado > 0 Then
          'actualizamos el corte en el que se registró la factura, esto es por si hay un cambio de corte al momento de hacer el registro de la información de la factura
          pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & lngidfactura
       End If
    End If
    
    '-------------------------------------------------------------------------------------------------
    'VALIDACIÓN DE LOS DATOS ANTES DE INSERTAR EN GNCOMPROBANTEFISCLADIGITAL EN EL PROCESO DE TIMBRADO
    '-------------------------------------------------------------------------------------------------
    If intTipoEmisionComprobante = 2 Then
       If Not fblnValidaDatosCFDCFDi(lngidfactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacion), strNumeroAprobacion) Then
              EntornoSIHO.ConeccionSIHO.RollbackTrans
            Exit Sub
           End If
    End If
       
    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, llngPersonaGraba, "FACTURACION MEMBRESIA SOCIO", lblFolio.Caption)
    EntornoSIHO.ConeccionSIHO.CommitTrans

    '*** GENERACIÓN DEL CFDi ***
    '<Si se realizará una emisión digital>
    If intTipoEmisionComprobante = 2 Then
        pgbBarraCFD.Value = 70
        freBarraCFD.Top = 3200
        Screen.MousePointer = vbHourglass
        lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
        freBarraCFD.Visible = True
        freBarraCFD.Refresh
        frmFacturacionMembresiaSocios.Enabled = False
        If intTipoCFDFactura = 1 Then
           pLogTimbrado 2
           pMarcarPendienteTimbre lngidfactura, "FA", vgintNumeroDepartamento
        End If
        EntornoSIHO.ConeccionSIHO.BeginTrans 'iniciamos transaccion de timbrado
        If Not fblnGeneraComprobanteDigital(lngidfactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, IIf(intTipoCFDFactura = 1, True, False)) Then
           On Error Resume Next

           EntornoSIHO.ConeccionSIHO.CommitTrans
           If intTipoCFDFactura = 1 Then pLogTimbrado 1
           If vgIntBanderaTImbradoPendiente = 1 Then
              'El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
              MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura"), vbInformation + vbOKOnly, "Mensaje"
           ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then 'No se realizó el timbrado
                 '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
                  MsgBox SIHOMsg(1338), vbCritical + vbOKOnly, "Mensaje"
                  pCancelarFactura Trim(lblFolio.Caption), llngPersonaGraba, "frmFacturacionMembresiaSocios", True
                  fblnImprimeComprobanteDigital lngidfactura, "FA", "I", llngFormato, 1
                  Screen.MousePointer = vbDefault
                  freBarraCFD.Visible = False
                  frmFacturacionMembresiaSocios.Enabled = True
                  freBarraCFD.Visible = False
                  pReinicia
                  txtClaveUnica.SetFocus
                  Exit Sub
           End If
        Else


            EntornoSIHO.ConeccionSIHO.CommitTrans
            If intTipoCFDFactura = 1 Then
               pEliminaPendientesTimbre lngidfactura, "FA" 'quitamos la factura de pendientes de timbre fiscal
               pLogTimbrado 1
            End If
        End If
        'Barra de progreso CFDi
        pgbBarraCFD.Value = 100
        freBarraCFD.Top = 3200
        Screen.MousePointer = vbDefault
        freBarraCFD.Visible = False
        frmFacturacionMembresiaSocios.Enabled = True
    End If
   
    '*** IMPRESIÓN DEL CFD ***
                                
        '<Si se realizará una emisión digital>
        If intTipoEmisionComprobante = 2 Then
           If Not fblnImprimeComprobanteDigital(lngidfactura, "FA", "I", llngFormato, 1) Then
               Exit Sub
           End If
            
           If vgIntBanderaTImbradoPendiente = 0 Then
              'Preguntar si se enviará el CFD por correo electrónico
              If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                 pEnviarCFD "FA", lngidfactura, CLng(vgintClaveEmpresaContable), txtRFC.Text, llngPersonaGraba, Me
              End If
           End If
        Else
        '<Emisión física>
            'Asegúrese de que la impresora esté   lista y  presione aceptar.
            MsgBox SIHOMsg(343), vbOKOnly + vbInformation, "Mensaje"
            If vgintNumeroModulo <> 2 Then
                strTotalLetras = fstrNumeroenLetras(CDbl(Format(lblTotal.Caption, cstrCantidad)), "pesos", "M.N.")
                vgstrParametrosSP = lblFolio.Caption & "|" & strTotalLetras
                Set rsFactura = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptFactura")
                pImpFormato rsFactura, cintTipoFormato, llngFormato
            Else
                pImprimeFormato llngFormato, lngidfactura
            End If
        End If
        
    pReinicia
    txtClaveUnica.SetFocus
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
    cmdSave.Enabled = False
    lblnConsulta = False
    Unload Me
End Sub
Private Sub pGuardaPolizaCorte(lngNumeroCorte As Long)
    Dim intcontador As Integer
    Dim dblTotalCliente As Double
    Dim dblIVA As Double
       
    '--------------------------------------------------------------------
    '|  Hace los movimientos inversos de los cargos del socio
    '--------------------------------------------------------------------
    
    '|  Abono a la cuenta por cobrar de cuotas de socios
    'pInsCortePoliza lngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorCobrar, lblTotalFactura.Caption, 0
    pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, txtClaveSocio.Text, "SO", lngCuentaCuotasPorCobrar, lblTotalFactura.Caption, False, "", 0, 0, "", 0, 2, txtClaveSocio.Text, "SO"
    '|  Cargo a la cuenta por devengar de cuotas de socios
    'pInsCortePoliza lngNumeroCorte, txtClaveSocio.Text, "SO", lngCuentaCuotasPorDevengar, lblTotalFactura.Caption, 1
    pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, txtClaveSocio.Text, "SO", lngCuentaCuotasPorDevengar, lblTotalFactura.Caption, True, "", 0, 0, "", 0, 2, txtClaveSocio.Text, "SO"
    '--------------------------------------------------------------------
    '|  Hace los movimientos de las cuentas de ingresos
    '--------------------------------------------------------------------
    intcontador = 0
    Do While intcontador <= UBound(apoliza(), 1)
        'pInsCortePoliza lngNumeroCorte, lblFolio.Caption, "FA", aPoliza(intContador).lngNumCuenta, aPoliza(intContador).dblCantidad, aPoliza(intContador).intNaturaleza = 1
        pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, lblFolio.Caption, "FA", apoliza(intcontador).lngnumCuenta, apoliza(intcontador).dblCantidad, apoliza(intcontador).intNaturaleza = 1, _
        "", 0, 0, "", 0, 2, lblFolio.Caption, "FA"
        intcontador = intcontador + 1
    Loop
    '------------------------------------------------
    '   Hace los cargos a las formas de pago
    '------------------------------------------------
    If Val(Format(lblTotal.Caption, "")) > 0 Then
        For intcontador = 0 To UBound(aFormasPago)
            'pInsCortePoliza lngNumeroCorte, lblFolio.Caption, "FA", aFormasPago(intContador).vllngCuentaContable, aFormasPago(intContador).vldblCantidad, True
            pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, lblFolio.Caption, "FA", aFormasPago(intcontador).vllngCuentaContable, aFormasPago(intcontador).vldblCantidad, True, _
            "", 0, 0, "", 0, 2, lblFolio.Caption, "FA", aFormasPago(intcontador).vlbolEsCredito, aFormasPago(intcontador).vlstrRFC, aFormasPago(intcontador).vlstrBancoSAT, aFormasPago(intcontador).vlstrBancoExtranjero, aFormasPago(intcontador).vlstrCuentaBancaria, aFormasPago(intcontador).vldtmFecha
        Next
    End If
    
    '------------------------------------------------
    '   Abona la cuenta del IVA cobrado
    '------------------------------------------------
    dblIVA = CDbl(Format(lblIVA.Caption, cstrCantidad))
    If dblIVA <> 0 Then
        'pInsCortePoliza lngNumeroCorte, lblFolio.Caption, "FA", glngCtaIVACobrado, dblIVA, 0
        pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, lblFolio.Caption, "FA", glngCtaIVACobrado, dblIVA, False, "", 0, 0, "", 0, 2, lblFolio.Caption, "FA"
    End If
End Sub

Private Sub pPoneFolioFacturaACargos()
    Dim vlstrSentencia As String
    Dim vlstrAux As String
    Dim vlintContador As Integer
    
    '--------------------------------------
    ' Poner el número de factura en los CARGOS
    '--------------------------------------
    vlstrSentencia = "  Update PVCargo " & _
                     "     set chrFolioFactura = '" & Trim(lblFolio.Caption) & "'" & _
                     "   where intNumCargo IN ("
    For vlintContador = 1 To grdCargos.Rows - 1
        vlstrSentencia = vlstrSentencia & Trim(Str(grdCargos.TextMatrix(vlintContador, cintColCveCargo)))
        vlstrAux = vlstrAux & "," & Trim(Str(grdCargos.TextMatrix(vlintContador, cintColCveCargo)))
         If vlintContador < grdCargos.Rows - 1 Then
            vlstrSentencia = vlstrSentencia & ", "
        End If
        If vlintContador Mod 50 = 0 And vlstrAux <> "" Then
            vlstrSentencia = Trim(vlstrSentencia)
            If Mid(vlstrSentencia, Len(vlstrSentencia), 1) = "," Then
                vlstrSentencia = Mid(vlstrSentencia, 1, Len(vlstrSentencia) - 1) & ") "
            Else
                vlstrSentencia = vlstrSentencia & ")"
            End If
            pEjecutaSentencia (vlstrSentencia)
            vlstrSentencia = ""
            vlstrSentencia = vlstrSentencia & " Update PVCargo " & _
                                              "    Set chrFolioFactura = '" & Trim(lblFolio.Caption) & "'" & _
                                              "  Where intNumCargo IN ("
            vlstrAux = ""
        End If
    Next vlintContador
    
    vlintContador = vlintContador - 1
    If vlintContador Mod 50 <> 0 And vlstrAux <> "" Then
        vlstrSentencia = Trim(vlstrSentencia)
        If Mid(vlstrSentencia, Len(vlstrSentencia), 1) = "," Then
            vlstrSentencia = Mid(vlstrSentencia, 1, Len(vlstrSentencia) - 1) & ") "
        Else
            vlstrSentencia = vlstrSentencia & ")"
        End If
        pEjecutaSentencia (vlstrSentencia)
    End If
End Sub
Private Sub pGuardaFacturaCorte(lngNumeroCorte As Long)
    Dim rs As New ADODB.Recordset
    Dim intFormaPago As Integer

    If CDbl(Format(lblTotal.Caption, cstrCantidad)) <> 0 Then
        For intFormaPago = 0 To UBound(aFormasPago)
'            vgstrParametrosSP = CStr(lngNumeroCorte) & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & Trim(lblFolio.Caption) _
'                                & "|" & "FA" & "|" & aFormasPago(intFormaPago).vlintNumFormaPago & "|" & aFormasPago(intFormaPago).vldblCantidad _
'                                & "|" & "0" & "|" & aFormasPago(intFormaPago).vlstrFolio & "|" & CStr(lngNumeroCorte)
'            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
            pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, Trim(lblFolio.Caption), "FA", 0, aFormasPago(intFormaPago).vldblCantidad, False, fstrFechaSQL(fdtmServerFecha, fdtmServerHora, True), CLng(aFormasPago(intFormaPago).vlintNumFormaPago), _
            0, aFormasPago(intFormaPago).vlstrFolio, lngNumeroCorte, 1, Trim(lblFolio.Caption), "FA", aFormasPago(intFormaPago).vlbolEsCredito, aFormasPago(intFormaPago).vlstrRFC, aFormasPago(intFormaPago).vlstrBancoSAT, aFormasPago(intFormaPago).vlstrBancoExtranjero, aFormasPago(intFormaPago).vlstrCuentaBancaria, aFormasPago(intFormaPago).vldtmFecha
        Next
    End If
End Sub
Private Sub pLlenaPagos()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vldblTotalPagos As Double
    Dim vldblCantidad As Double
    Dim vlintContador As Integer
    
    vlstrSentencia = "Select intNumPago " & _
                     "     , pvpago.intNumConcepto " & _
                     "     , chrDescripcion Concepto " & _
                     "     , dtmFecha Fecha" & _
                     "     , chrFolioRecibo Recibo" & _
                     "     , mnyCantidad Cantidad " & _
                     "     , Case bitPesos when 1 then 'Pesos' Else 'Dolares' end as Moneda " & _
                     "     , mnyTipoCambio TipoCambio" & _
                     "     , chrTipo TipoPago " & _
                     "     , isnull(chrFolioFactura,'') Factura" & _
                     "     , 'E' EntradaSalida " & _
                     "     , intNumCorte Corte " & _
                     "     , pvconceptopagoempresa.intnumerocuenta CuentaConcepto " & _
                     "  From pvpago " & _
                     "       inner join pvConceptoPago on pvPago.intNumConcepto = pvConceptoPago.intNumConcepto " & _
                     "       inner join pvconceptopagoempresa on pvconceptopago.intnumconcepto = pvconceptopagoempresa.intnumconcepto " & _
                     " Where pvconceptopagoempresa.intcveempresa = " & vgintClaveEmpresaContable & _
                     "   and chrTipoPaciente = 'S'" & _
                     "   and intMovPaciente = " & Trim(txtClaveSocio.Text)
    vlstrSentencia = vlstrSentencia & _
                     IIf(chkIncluyePagosFacturados.Value, " and not (bitCancelado = 1 and chrFolioFactura is null)", " and bitCancelado = 0")
    vlstrSentencia = vlstrSentencia & _
                    " UNION " & _
                    "Select intNumSalida " & _
                    "     , pvSalidaDinero.intNumConcepto " & _
                    "     , chrDescripcion Concepto " & _
                    "     , dtmFecha Fecha " & _
                    "     , chrFolioRecibo Recibo " & _
                    "     , mnyCantidad*-1 Cantidad " & _
                    "     , Case bitPesos " & _
                    "            when 1 then 'Pesos'  " & _
                    "            Else 'Dolares'  " & _
                    "       end as Moneda " & _
                    "     , mnyTipoCambio TipoCambio " & _
                    "     , 'SD' TipoPago " & _
                    "     , IsNull(chrFolioFactura,'') Factura " & _
                    "     , 'S' EntradaSalida " & _
                    "     , intNumCorte Corte " & _
                    "     , pvconceptopagoempresa.intnumerocuenta CuentaConcepto " & _
                    "  From pvSalidaDinero " & _
                    "       Inner Join pvConceptoPago on pvSalidaDinero.intNumConcepto = pvConceptoPago.intNumConcepto " & _
                    "       Inner Join pvconceptopagoempresa on pvconceptopago.intnumconcepto = pvconceptopagoempresa.intnumconcepto " & _
                    " Where chrTipoPaciente = 'S' " & _
                    "   And intMovPaciente = " & Trim(txtClaveSocio.Text) & _
                    "   And pvconceptopagoempresa.intcveempresa = " & vgintClaveEmpresaContable & _
                    IIf(chkIncluyePagosFacturados.Value, " and not (bitCancelado = 1 and chrFolioFactura is null)", " and bitCancelado = 0")
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLimpiaGrid grdPagos
    pConfiguraGridPagos
    vldblTotalPagos = 0
    With grdPagos
        .Redraw = False 'Optimización
        rs.Sort = "Fecha"
        
        Do While Not rs.EOF
            If grdPagos.RowData(1) <> -1 Then
                 grdPagos.Rows = grdPagos.Rows + 1
                 grdPagos.Row = grdPagos.Rows - 1
            End If

            .Col = 0
            .CellFontBold = True
            If (Trim(rs!Factura) = "") Or IsNull(rs!Factura) Then
                .TextMatrix(.Row, 0) = ""
            Else
                .TextMatrix(.Row, 0) = "F"
            End If
            .RowData(.Row) = rs!intNumPago 'No me sirve de nada
            .TextMatrix(.Row, 1) = rs!Concepto
            .TextMatrix(.Row, 2) = Format(rs!fecha, "dd/mmm/yyyy")
            .TextMatrix(.Row, 3) = Format(rs!Cantidad, "$ ###,###,###,###.00") 'La cantidad se pone igual que como se grabó (Sin hacer nungún tipo de conversión dado que se despliega también la moneda)
            vldblCantidad = IIf(rs!Moneda = "Pesos", 1, rs!TipoCambio) * rs!Cantidad 'Convierte a pesos (Si es necesario)
            .TextMatrix(.Row, 4) = rs!Moneda
            .TextMatrix(.Row, 5) = Trim(rs!Recibo)
            .TextMatrix(.Row, 6) = IIf(IsNull(rs!Factura), "", rs!Factura)
            .TextMatrix(.Row, 7) = rs!tipoPago
            .TextMatrix(.Row, 8) = "" 'Disponible
            .TextMatrix(.Row, 9) = rs!EntradaSalida
            .TextMatrix(.Row, 10) = rs!corte
            .TextMatrix(.Row, 11) = rs!CuentaConcepto

            If (rs!tipoPago = "NO" Or rs!tipoPago = "SD") And (.TextMatrix(.Row, 0) <> "F") Then
                vldblTotalPagos = vldblTotalPagos + vldblCantidad
            End If
            
            For vlintContador = 2 To .Cols - 1
                If rs!EntradaSalida = "S" Then
                    .Col = vlintContador
                    .CellBackColor = &H80000018
                Else
                    .Col = vlintContador
                    .CellForeColor = &H80000008
                End If
            Next
            rs.MoveNext
        Loop
    .Redraw = True 'Optimización
    rs.Close
    End With
'    txtPagos.Text = Format(vldblTotalPagos, "$ ###,###,###,##0.00")
End Sub

Private Sub pCancelaPagos(strFolioDocumento As String, lngNumeroCorte As Long)
    Dim vlintContador As Integer
    Dim vldtmFechaHoy As Date    '|  Varible con la Fecha actual
    Dim vldtmHoraHoy As Date     '|  Varible con la Hora actual
    Dim vlstrSentencia As String
    Dim rsFormasPagos As New ADODB.Recordset
   
    vldtmFechaHoy = fdtmServerFecha
    vldtmHoraHoy = fdtmServerHora

    '-------------------------------------
    '|  Cancelo los pagos
    '-------------------------------------
    If Trim(grdPagos.TextMatrix(1, 1)) <> "" Then '|  Si no esta vacia la cuadrícula de Pagos
        For vlintContador = 1 To grdPagos.Rows - 1
            '-------------------------------------------------------
            '|  Se tomarán en cuenta solo los pagos NORMALES que no hayan sido facturados previamente
            '-------------------------------------------------------
            If (grdPagos.TextMatrix(vlintContador, 7) = "NO" Or grdPagos.TextMatrix(vlintContador, 7) = "SD") And Trim(grdPagos.TextMatrix(vlintContador, 6)) = "" Then
                '-------------------------------------
                ' Cancelo el Pago o Salida de Efectivo
                '-------------------------------------
                vlstrSentencia = "Update " & IIf(grdPagos.TextMatrix(vlintContador, 9) = "E", "pvPago", "pvSalidaDinero") & _
                                "    Set bitCancelado = 1" & _
                                "      , chrFolioFactura = '" & Trim(strFolioDocumento) & "'" & _
                                "  Where rtrim(chrFolioRecibo) = '" & Trim(grdPagos.TextMatrix(vlintContador, 5)) & "'" & _
                                "    And intMovPaciente = '" & Trim(txtClaveSocio.Text) & "'" & _
                                "    And chrTipoPaciente = 'S'"
                pEjecutaSentencia (vlstrSentencia)
                '-------------------------------------
                ' Registrar las formas del Pago o Salida en la factura
                '-------------------------------------
                vgstrParametrosSP = Trim(grdPagos.TextMatrix(vlintContador, 5)) & "|" & IIf(grdPagos.TextMatrix(vlintContador, 9) = "E", "RE", "SD") & "|" & grdPagos.TextMatrix(vlintContador, 10)
                Set rsFormasPagos = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaDoctoCorte")
                Do While Not rsFormasPagos.EOF
'                    vgstrParametrosSP = CStr(lngNumeroCorte) _
'                                        & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) _
'                                        & "|" & strFolioDocumento _
'                                        & "|" & "FA" _
'                                        & "|" & CStr(rsFormasPagos!intFormaPago) _
'                                        & "|" & CStr(rsFormasPagos!MNYCANTIDADPAGADA * IIf(Trim(rsFormasPagos!chrTipoDocumento) = "SD", -1, 1)) _
'                                        & "|" & CStr(rsFormasPagos!mnytipocambio) _
'                                        & "|" & CStr(rsFormasPagos!intfoliocheque) _
'                                        & "|" & CStr(lngNumeroCorte)
'                    frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
                    
                    pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, strFolioDocumento, "FA", 0, rsFormasPagos!mnyCantidadPagada * IIf(Trim(rsFormasPagos!chrTipoDocumento) = "SD", -1, 1), False, _
                    fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")), rsFormasPagos!intFormaPago, rsFormasPagos!MNYTIPOCAMBIO, _
                    rsFormasPagos!intfoliocheque, lngNumeroCorte, 1, strFolioDocumento, "FA"
                    
                    

                    'Movimiento de cargo a la cuenta de la forma de pago del recibo, pero como factura:
'                    vgstrParametrosSP = CStr(lngNumeroCorte) _
'                                        & "|" & Trim(strFolioDocumento) _
'                                        & "|" & "FA" _
'                                        & "|" & CStr(rsFormasPagos!INTCUENTACONTABLE) _
'                                        & "|" & IIf(rsFormasPagos!mnytipocambio = 0, rsFormasPagos!MNYCANTIDADPAGADA, rsFormasPagos!MNYCANTIDADPAGADA * rsFormasPagos!mnytipocambio) _
'                                        & "|" & IIf(Trim(rsFormasPagos!chrtipodocumento) = "SD", 0, 1)
'                    frsEjecuta_SP vgstrParametrosSP, "SP_PVINSPVCORTEPOLIZA"
                    pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, Trim(strFolioDocumento), "FA", rsFormasPagos!INTCUENTACONTABLE, IIf(rsFormasPagos!MNYTIPOCAMBIO = 0, rsFormasPagos!mnyCantidadPagada, rsFormasPagos!mnyCantidadPagada * rsFormasPagos!MNYTIPOCAMBIO), _
                    IIf(Trim(rsFormasPagos!chrTipoDocumento) = "SD", False, True), "", 0, 0, "", 0, 2, Trim(strFolioDocumento), "FA"

                    rsFormasPagos.MoveNext
                Loop
                rsFormasPagos.Close

                '-------------------------------------
                ' Cancelar el Pago o Salida en el corte y su movimiento contable
                '-------------------------------------
'                vgstrParametrosSP = Trim(grdPagos.TextMatrix(vlintcontador, 5)) _
'                                    & "|" & IIf(grdPagos.TextMatrix(vlintcontador, 9) = "E", "RE", "SD") _
'                                    & "|" & grdPagos.TextMatrix(vlintcontador, 10) _
'                                    & "|" & CStr(lngNumeroCorte)
'                frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdCancelaDoctoCorte"
                pAgregarMovArregloCorte lngNumeroCorte, llngPersonaGraba, Trim(grdPagos.TextMatrix(vlintContador, 5)), IIf(grdPagos.TextMatrix(vlintContador, 9) = "E", "RE", "SD"), 0, 0, False, "", 0, 0, "", CLng(grdPagos.TextMatrix(vlintContador, 10)), _
                3, Trim(strFolioDocumento), "FA"
                

            End If
        Next vlintContador
    End If
End Sub

Private Sub pGuardaDetalleFactura(lngidfactura As Long)
    Dim intcontador As Integer
    Dim rs As New ADODB.Recordset
    Dim dblCantidad As Double
    Dim dblimportegravado As Double
    Dim dblSubTotal As Double
    
    strSentencia = "select * from PvDetalleFactura where chrFolioFactura = '*'"
    Set rs = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    dblimportegravado = 0
    dblSubTotal = 0
    
    For intcontador = 1 To grdConceptos.Rows - 1
        With grdConceptos
            rs.AddNew
            rs!chrfoliofactura = lblFolio.Caption
            rs!smicveconcepto = Val(.TextMatrix(intcontador, cintColCveConcepto))
            rs!MNYCantidad = CDbl(Format(.TextMatrix(intcontador, cintColCantidadConcepto), cstrCantidad))
            rs!MNYDESCUENTO = 0
            rs!MNYIVA = CDbl(Format(.TextMatrix(intcontador, cintColIVAConcepto), cstrCantidad))
            rs!chrTipo = "NO"
            
            If rs!MNYIVA <> 0 Then
                rs!mnyIVAConcepto = rs!MNYCantidad * (vgdblCantidadIvaGeneral / 100)
            End If
            
            rs.Update
            
            '--Calcular el importe gravado y el descuento sobre el importe gravado
            If CDbl(Format(.TextMatrix(intcontador, cintColIVAConcepto), cstrCantidad)) > 0 Then
                dblimportegravado = dblimportegravado + CDbl(Format(.TextMatrix(intcontador, cintColCantidadConcepto), cstrCantidad))
            End If
            
            dblCantidad = CDbl(Format(.TextMatrix(intcontador, cintColCantidadConcepto), cstrCantidad))
            pLlenaPoliza CLng(.TextMatrix(intcontador, cintColCtaIngresoConcepto)), dblCantidad, 0
            
        End With
    Next intcontador
    rs.Close
    '--Insertar datos en pvfacturaimporte
    dblSubTotal = Val(Format(lblTotalFactura.Caption, cstrCantidad)) - Val(Format(lblIVA.Caption, cstrCantidad))
    vgstrParametrosSP = lngidfactura & "|" & dblimportegravado & "|" & dblSubTotal - dblimportegravado & "|" & 0 & "|" & 0
    frsEjecuta_SP vgstrParametrosSP, "sp_PvInsFacturaImporte", True
    If cboUsoCFDI.Enabled Then
        cboUsoCFDI.SetFocus
    End If
End Sub

Private Function fintErrorBloqueoCorte() As Integer
    '----------------------------------------------------------------------------------------------
    'Función para bloquear el corte
    '----------------------------------------------------------------------------------------------
    Dim lngCorteGrabando  As Long   'Resultado de la validación del estado del corte
    
    lngCorteGrabando = 1
    vgstrParametrosSP = llngNumCorte & "|" & "Grabando"
    frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, lngCorteGrabando
    If lngCorteGrabando <> 2 Then
        fintErrorBloqueoCorte = 720 'No se puede realizar la operación, inténtelo en unos minutos.
    End If
End Function

Private Function fblnDatosValidos() As Boolean
    Dim intcontador As Integer

    fblnDatosValidos = True
    
    fblnDatosValidos = fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 2417, 609), "E")
    
    If fblnDatosValidos And Trim(lblSocio.Caption) = "" Then
        fblnDatosValidos = False
        'Seleccione el cliente.
        MsgBox SIHOMsg(322), vbExclamation + vbOKOnly, "Mensaje"
        txtClaveUnica.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskFecha.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
        mskFecha.SetFocus
    End If
    If fblnDatosValidos Then
        If CDate(mskFecha.Text) > fdtmServerFecha Then
            fblnDatosValidos = False
            '¡La fecha debe ser menor o igual a la del sistema!
            MsgBox SIHOMsg(40), vbExclamation + vbOKOnly, "Mensaje"
            mskFecha.SetFocus
        End If
    End If
    If fblnDatosValidos Then
        intcontador = 1
        Do While intcontador <= grdCargos.Rows - 2 And fblnDatosValidos
            If Val(Format(grdCargos.TextMatrix(intcontador, cintColCantidadCargo), cstrCantidad)) = 0 Then
                fblnDatosValidos = False
                grdCargos.Row = intcontador
                grdCargos.Col = cintColCantidadCargo
                'No se puede realizar la operación con cantidad cero o menor que cero
                MsgBox SIHOMsg(651), vbExclamation + vbOKOnly, "Mensaje"
                grdCargos.SetFocus
            End If
            intcontador = intcontador + 1
        Loop
    End If
    If fblnDatosValidos Then
        fblnDatosValidos = fblnAsignaImpresora(vgintNumeroDepartamento, "FA")
        If Not fblnDatosValidos Then
            'No se tiene asignada una impresora en la cual imprimir las facturas
            MsgBox SIHOMsg(492), vbExclamation + vbOKOnly, "Mensaje"
        End If
    End If
End Function

Private Sub Form_Activate()
    Dim intMensaje As Integer

    intMensaje = CInt(flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P"))

    If intMensaje <> 0 Then
        'Cierre el corte actual antes de registrar este documento.
        'No existe un corte abierto
        MsgBox SIHOMsg(Str(intMensaje)), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    If glngCtaIVACobrado = 0 Or glngCtaIVANoCobrado = 0 Then
        'No se encuentran registradas las cuentas de IVA cobrado y no cobrado en los parámetros generales del sistema.
        MsgBox SIHOMsg(729), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    If glngCveImpuesto = 0 Then
        'No se encuentra registrada la tasa de IVA en los parámetros generales del sistema.
        MsgBox SIHOMsg(731), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    If llngFormato = 0 Then
        'Configure el formato de impresión en los parámetros del módulo.
        MsgBox Mid(SIHOMsg(732), 1, 51) & ".", vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
    If Trim(lblFolio.Caption) = "0" Then
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291), vbCritical, "Mensaje"
        '|  Inicializo estas variables para que se salga sin preguntar
        cmdSave.Enabled = False
        lblnConsulta = False
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()
    Dim strParametrosSP As String

    Dim lngFoliosRestantes As Long
    Dim strFolioDocumento As String
    Dim lngMensaje As Long
    Dim rsTipoFacturacion As New ADODB.Recordset
    Dim strSentencia As String
    Dim lngCveFormato As Long
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    Dim alstrParametrosSalida() As String
    
    strVersionCFDISocios = vgstrVersionCFDI
    'vgstrVersionCFDI = "3.2"
    
    Me.Icon = frmMenuPrincipal.Icon
    llngFormato = flngFormatoDepto(vgintNumeroDepartamento, cintTipoFormato, "S")
    
    '|  Obtiene el número de cuenta contable de la cuenta puente de cuotas por cobrar de socios
    strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
    lngCuentaCuotasPorCobrar = 1
    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, lngCuentaCuotasPorCobrar
    '|  Obtiene el número de cuenta contable de la cuenta puente de cuotas por devengar de socios
    strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
    lngCuentaCuotasPorDevengar = 1
    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, lngCuentaCuotasPorDevengar
    
    pReinicia
    
    pCargaUsosCFDI
    
    blnActivarconsulta = True
    
    SSTFactura.Tab = 0
End Sub

Private Sub pReinicia()
    lblnConsulta = False
    
    pLimpiaEncabezado
    
    pLimpiagrdBusquedaFactura
    pConfiguragrdBusquedaFactura
    
    pLimpiaGrid grdCargos
    pConfiguraGridCargos
    
    pLimpiaGrid grdConceptos
    pLimpiaConceptos
    
    pLimpiaGrid grdPagos
    pConfiguraGridPagos
    
    cmdCFD.Enabled = False
    vgConsecutivoMuestraPvFactura = 0
    
    lblSubtotal.Caption = FormatCurrency(0, 2)
    lblIVA.Caption = FormatCurrency(0, 2)
    lblTotalFactura.Caption = FormatCurrency(0, 2)
    lblTotalPagos.Caption = FormatCurrency(0, 2)
    lblTotal.Caption = FormatCurrency(0, 2)
    
    lblFolio.ForeColor = llngColorActivas
    chkOtrosDatosFiscales.Value = False
    
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    pHabilita 0, 0, 1, 0, 0, 0, 0
 
    Me.cmdCancelaFacturasSAT.Enabled = False
    Me.cmdConfirmartimbre.Enabled = False
    Me.cmdconfirmartimbrefiscal.Enabled = False
    blnActivarconsulta = False
    Me.ChkFacturasCancelaNoSAT.Value = vbUnchecked
    Me.ChkPendientesTimbre.Value = vbUnchecked
    blnActivarconsulta = True
    Me.txtpendientetimbre.Visible = False
    
    lngConsecutivoFactura = 0
    SSTab1.Tab = 0
    
    cboUsoCFDI.Enabled = False
    chkOtrosDatosFiscales.Enabled = False
End Sub

Private Sub pLimpiaEncabezado()
    fraCliente.Enabled = True
    fraFolioFecha.Enabled = True
        
    txtClaveUnica.Text = ""
    txtClaveSocio.Text = ""
    
    mskFecha.Mask = ""
    mskFecha.Text = fdtmServerFecha
    mskFecha.Mask = "##/##/####"

    chkBitExtranjero.Enabled = True
    'lblFolio.ForeColor = llngColorActivas

    pCargaFolio 0

    mskFechaIni.Mask = ""
    mskFechaIni.Text = fdtmServerFecha
    mskFechaIni.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    
    txtBusquedaSocio.Text = ""
    lblBusquedaNombreSocio.Caption = ""

    chkBitExtranjero.Value = vbUnchecked
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If SSTFactura.Tab = 0 Then
        If cmdSave.Enabled Or lblnConsulta Then
            Cancel = True
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pReinicia
                txtClaveUnica.SetFocus
            End If
        End If
    End If
    If SSTFactura.Tab = 1 Then
        Cancel = True
        SSTFactura.Tab = 0
        cmdLocate.SetFocus
    End If
    If Not Cancel Then
        vgstrVersionCFDI = strVersionCFDISocios
    End If
End Sub
Private Sub grdBusquedaFactura_Click()
    With grdBusquedaFactura
        If .MouseCol = 0 And .MouseRow > 0 Then
           If IIf(Trim(.TextMatrix(.Row, cintColPCancelarNoSAt)) = "", 0, Val(.TextMatrix(.Row, cintColPCancelarNoSAt))) > 0 Then
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
           ElseIf .TextMatrix(.Row, cintColPTimbre) = "1" Then
                If .TextMatrix(.Row, 0) = "*" Then
                    If vllngSeleccPendienteTimbre > 0 Then
                       vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre - 1
                    End If
                    .TextMatrix(.Row, 0) = ""
                Else
                    vllngSeleccPendienteTimbre = vllngSeleccPendienteTimbre + 1
                    .TextMatrix(.Row, 0) = "*"
                End If
           Me.cmdconfirmartimbrefiscal.Enabled = vllngSeleccPendienteTimbre > 0
           End If
        End If
    End With
End Sub
Private Sub grdBusquedaFactura_DblClick()
    If Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cIntColFolio)) <> "" Then
        SSTFactura.Tab = 0
        pMuestra
        If grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColEstadoNuevoEsquemaCancelacion) = "CR" Or grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColEstadoNuevoEsquemaCancelacion) = "PA" Then
            pHabilita 1, 1, 1, 1, 1, 0, IIf(Me.txtpendientetimbre.Visible = False, 0, IIf(Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColEstadoNuevoEsquemaCancelacion)) = "CR" Or Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColEstadoNuevoEsquemaCancelacion)) = "PA", 1, 0))
        Else
             pHabilita 1, 1, 1, 1, 1, 0, IIf(Me.txtpendientetimbre.Visible = True, 0, IIf(Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cintColchrEstatus)) = "C", 0, 1))
        End If
        cmdLocate.SetFocus
    End If
End Sub
Private Sub grdBusquedaFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdBusquedaFactura_DblClick
    End If
End Sub

Private Sub mskFecha_GotFocus()
    pSelMkTexto mskFecha
End Sub
Private Sub mskFechaFin_GotFocus()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaIni_GotFocus()
    pSelMkTexto mskFechaIni
End Sub
Private Sub pLimpiagrdBusquedaFactura()
    grdBusquedaFactura.Clear
    grdBusquedaFactura.Rows = 2
    grdBusquedaFactura.Cols = cintColgrdBusquedaFactura
    grdBusquedaFactura.FormatString = cstrFormatoBusquedaFactura
End Sub

Private Sub txtBusquedaSocio_Change()
    lblBusquedaNombreSocio.Caption = ""
End Sub
Private Sub txtBusquedaSocio_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCveSocio As Long
    Dim strParametrosSP As String
    Dim rsSocio As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If Trim(txtBusquedaSocio.Text) = "" Then
            With frmSociosBusqueda
                .Show vbModal, Me
                If .vglngClaveSocio <> 0 Then
                     txtBusquedaSocio.Text = .vgstrClaveUnica
                     lblBusquedaNombreSocio.Caption = .vgstrNombreSocio
                End If
                Unload frmSociosBusqueda
            End With
        Else
            lngCveSocio = flngObtieneClaveSocio(txtClaveUnica.Text)
            If lngCveSocio = -1 Then
                '|  ¡No existe información!
                MsgBox SIHOMsg(13), vbInformation, "Mensaje"
            Else
                strParametrosSP = CStr(lngCveSocio) & "|T"
                Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_SORPTSELSOCIO")
                If rsSocio.RecordCount > 0 Then
                    lblBusquedaNombreSocio.Caption = IIf(IsNull(rsSocio!Nombre), "", rsSocio!Nombre)
                End If
            End If
        End If
    End If
End Sub
Private Sub txtBusquedaSocio_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtClaveUnica_Change()
    lblSocio.Caption = ""
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
End Sub
Private Sub txtClaveUnica_GotFocus()
    pSelTextBox txtClaveUnica
End Sub
Private Sub pAsignaDatosSocio(rs As ADODB.Recordset)
    txtClaveUnica.Text = rs!intNumCliente
    lblSocio.Caption = IIf(IsNull(rs!NombreCliente), " ", rs!NombreCliente)
    lblCiudad.Caption = IIf(IsNull(rs!ciudadcliente), " ", rs!ciudadcliente)
    If IsNull(rs!RFCCliente) Then
        txtRFC.Text = ""
    Else
        txtRFC.Text = fStrRFCValido(rs!RFCCliente)
    End If
    lblTelefono.Caption = IIf(IsNull(rs!Telefono), " ", rs!Telefono)
    lblDomicilio.Caption = IIf(IsNull(rs!CHRCALLE), " ", rs!CHRCALLE)
    lblNumeroExterior.Caption = IIf(IsNull(rs!VCHNUMEROEXTERIOR), " ", rs!VCHNUMEROEXTERIOR)
    lblNumeroInterior.Caption = IIf(IsNull(rs!VCHNUMEROINTERIOR), " ", rs!VCHNUMEROINTERIOR)
    lblColonia.Caption = IIf(IsNull(rs!Colonia), " ", rs!Colonia)
    lblCP.Caption = IIf(IsNull(rs!codigo), " ", rs!codigo)
End Sub
Private Sub pMuestra()
    Dim rs As New ADODB.Recordset
    Dim vlaux As String
    Dim vlSQL As String
    Dim vlrsAux As New ADODB.Recordset
    Dim rsCiudad As New ADODB.Recordset
    
    vgstrParametrosSP = Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cIntColFolio))
    
     Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFactura_NE")
    
    fraCliente.Enabled = False
    fraFolioFecha.Enabled = False
        
    If rs.RecordCount <> 0 Then
        pLimpiaEncabezado
        
        lngConsecutivoFactura = rs!IdFactura 'para poder hacer la cancelacion, confirmacion del timbre etc etc
        
        txtClaveSocio.Text = rs!cuenta
        lblFolio.ForeColor = IIf(rs!chrEstatus = "C", llngColorCanceladas, llngColorActivas)
        lblFolio.Caption = rs!Serie & rs!Folio
        Set vlrsAux = frsRegresaRs("SELECT * FROM GnComprobanteFiscalDigital INNER JOIN PVFactura ON GnComprobanteFiscalDigital.INTCOMPROBANTE = PVFactura.INTCONSECUTIVO AND GnComprobanteFiscalDigital.CHRTIPOCOMPROBANTE = 'FA' WHERE PVFactura.ChrFolioFactura = '" & Trim(grdBusquedaFactura.TextMatrix(grdBusquedaFactura.Row, cIntColFolio)) & "'")
        If vlrsAux.RecordCount <> 0 Then
            cmdCFD.Enabled = True
            vlstrTipoCFD = IIf(IsNull(vlrsAux!INTNUMEROAPROBACION), "CFDi", "CFD")
        Else
            cmdCFD.Enabled = False
        End If
        
        lblnConsulta = True
        
        txtClaveUnica.Text = rs!ClaveSocio
        lblSocio.Caption = rs!NombreSocio
        cboUsoCFDI.ListIndex = flngLocalizaCbo(cboUsoCFDI, rs!intCveUsoCFDI)
        'Carga la ciudad del domicilio fiscal
        vgstrParametrosSP = IIf(IsNull(rs!CveCiudad), 0, rs!CveCiudad) & "|-1|-1"
        Set rsCiudad = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELCIUDAD")
        If rsCiudad.RecordCount <> 0 Then
            lblCiudad.Caption = IIf(IsNull(rsCiudad!vchDescripcion), "", rsCiudad!vchDescripcion)
        End If
        
        lblTelefono.Caption = IIf(IsNull(rs!Telefono), " ", rs!Telefono)
        If IsNull(rs!RFC) Then
            txtRFC.Text = " "
        Else
            txtRFC.Text = fStrRFCValido(rs!RFC)
        End If
        vgConsecutivoMuestraPvFactura = rs!IdFactura
        lblDomicilio.Caption = IIf(IsNull(rs!CHRCALLE), " ", rs!CHRCALLE)
        lblNumeroExterior.Caption = IIf(IsNull(rs!VCHNUMEROEXTERIOR), " ", rs!VCHNUMEROEXTERIOR)
        lblNumeroInterior.Caption = IIf(IsNull(rs!VCHNUMEROINTERIOR), " ", rs!VCHNUMEROINTERIOR)
        lblColonia.Caption = IIf(IsNull(rs!Colonia), " ", rs!Colonia)
        lblCP.Caption = IIf(IsNull(rs!CP), " ", rs!CP)
        vlaux = IIf(IsNull(rs!Serie), "", rs!Serie)
        
        'Se habilita el chkExtranjero en caso de ser cliente Extranjero
        If Trim(rs!RFC) = "XEXX010101000" Then
            chkBitExtranjero.Value = vbChecked
        Else
            chkBitExtranjero.Value = vbUnchecked
        End If
        
        mskFecha.Mask = ""
        mskFecha.Text = Format(rs!fecha, "dd/mm/yyyy")
        mskFecha.Mask = "##/##/####"
        
        lblSubtotal.Caption = FormatCurrency(rs!Subtotal, 2)
        lblIVA.Caption = FormatCurrency(rs!IVAFactura, 2)
        lblTotalFactura.Caption = FormatCurrency(rs!TotalFactura, 2)
        lblTotalPagos.Caption = FormatCurrency(fdblPagosFacturados(rs!IdFactura), 2)
        lblTotal.Caption = FormatCurrency(rs!TotalFactura - fdblPagosFacturados(rs!IdFactura), 2)
        
    If rs!PendienteTimbre = 0 Then
        If rs!PendienteCancelarSat = 1 Then
            Me.cmdConfirmartimbre.Enabled = True
            Me.cmdCFD.Enabled = False
            Me.cmdDelete.Enabled = False
            Me.txtpendientetimbre.Text = "Pendiente de cancelarse ante el SAT"
            Me.txtpendientetimbre.ForeColor = &HFF&
            Me.txtpendientetimbre.BackColor = &HC0E0FF
            Me.txtpendientetimbre.Visible = True
           
        Else
               If rs!PendienteCancelarSat = 1 Then
                  Me.txtpendientetimbre.Text = "Pendiente de cancelarse ante el SAT"
                  Me.txtpendientetimbre.ForeColor = &HFF&
                  Me.txtpendientetimbre.BackColor = &HC0E0FF
                  Me.txtpendientetimbre.Visible = True
               Else
                  Me.txtpendientetimbre.Visible = False
                  Select Case rs!PendienteCancelarSAT_NE
                    Case "PA"
                        Me.txtpendientetimbre.Text = "Pendiente de autorización de cancelación"
                        txtpendientetimbre.ForeColor = &HFFFFFF '| Blanco
                        txtpendientetimbre.BackColor = &H80FF&  '| Naranja fuerte
                        txtpendientetimbre.Visible = True
                    Case "CR"
                        Me.txtpendientetimbre.Text = "Cancelación rechazada"
                        txtpendientetimbre.ForeColor = &HFFFFFF '| Blanco
                        txtpendientetimbre.BackColor = &HFF&    '| Rojo
                        txtpendientetimbre.Visible = True
                        
                        Me.cmdDelete.Enabled = True
                    Case "NP"
                        txtpendientetimbre.Visible = False
                  End Select
               End If
            
               Me.cmdConfirmartimbre.Enabled = False
               If Trim(vlstrTipoCFD) <> "" Then Me.cmdCFD.Enabled = True
        End If
    Else
          Me.txtpendientetimbre.Text = "Pendiente de timbre fiscal"
          Me.txtpendientetimbre.ForeColor = &H0&
          Me.txtpendientetimbre.BackColor = &HFFFF&
          Me.txtpendientetimbre.Visible = True
          
          Me.cmdConfirmartimbre.Enabled = True
        
        
    End If
    
    
        pLimpiaGrid grdConceptos
        pLimpiaConceptos
        With grdConceptos
            Do While Not rs.EOF And Not IsNull(rs!Concepto)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, cintColCveConcepto) = rs!smicveconcepto
                .TextMatrix(.Rows - 1, cintColDescripcionConcepto) = rs!Concepto
                .TextMatrix(.Rows - 1, cintColCantidadConcepto) = FormatCurrency(rs!Importe, 2)
                .TextMatrix(.Rows - 1, cintColIVAConcepto) = FormatCurrency(rs!IVA, 2)
                rs.MoveNext
            Loop
        End With
        '|  Oculta las pestañas Cargos y Pagos
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False

        
    Else
        '¡La información no existe!
        MsgBox SIHOMsg(12), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub
Private Sub pConfiguragrdBusquedaFactura()
    Dim intcontador As Integer

    With grdBusquedaFactura
        .FixedCols = 1
        .ColWidth(0) = 200
        .ColWidth(cintColNumPoliza) = 0
        .ColWidth(cIntColNumCorte) = 0
        .ColWidth(cintColchrEstatus) = 0
        .ColWidth(cintColFecha) = 1100
        .ColWidth(cIntColFolio) = 1100
        .ColWidth(cintColNumSocio) = 3150
        .ColWidth(cIntColRazonSocial) = 3150
        .ColWidth(cIntColRFC) = 1450
        .ColWidth(cintColTotalFactura) = 1100
        .ColWidth(cintColIVAConsulta) = 1000
        .ColWidth(cintColSubtotal) = 1100
        .ColWidth(cIntColEstado) = 1200
        .ColWidth(cintColFacturo) = 3150
        .ColWidth(cintColCancelo) = 3150
        .ColWidth(cintColPCancelarNoSAt) = 0
        .ColWidth(cintColPTimbre) = 0
        .ColWidth(cintColEstadoNuevoEsquemaCancelacion) = 0
        
        .ColAlignment(cintColFecha) = flexAlignLeftCenter
        .ColAlignment(cIntColFolio) = flexAlignLeftCenter
        .ColAlignment(cintColNumSocio) = flexAlignLeftCenter
        .ColAlignment(cIntColRazonSocial) = flexAlignLeftCenter
        .ColAlignment(cIntColRFC) = flexAlignLeftCenter
        .ColAlignment(cintColTotalFactura) = flexAlignRightCenter
        .ColAlignment(cintColIVAConsulta) = flexAlignRightCenter
        .ColAlignment(cintColSubtotal) = flexAlignRightCenter
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
    
    cmdLocate.Enabled = intlocate = 1
    cmdSave.Enabled = intSave = 1
    cmdDelete.Enabled = intDelete = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
    Unload Me
End Sub
Private Sub txtClaveUnica_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(txtClaveUnica.Text) = "" Then
            With frmSociosBusqueda
                .vglngClaveSocio = 0
                .Show vbModal, Me
                If .vglngClaveSocio <> 0 Then
                    pLlenaInformacionSocio .vglngClaveSocio
                End If
                Unload frmSociosBusqueda
                If fblnCanFocus(chkBitExtranjero) Then chkBitExtranjero.SetFocus
            End With
        Else
            pLlenaInformacionSocio flngObtieneClaveSocio(txtClaveUnica.Text)
        End If
        cboUsoCFDI.Enabled = True
        chkOtrosDatosFiscales.Enabled = True
    End If
End Sub
Private Function flngObtieneClaveSocio(strClaveSocio As String) As Long
    Dim rsSocio As New ADODB.Recordset
    Dim strParametrosSP As String
    
    strParametrosSP = CStr(strClaveSocio) & "|-1"
    Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_PVSELSOCIOS")
    If rsSocio.RecordCount > 0 Then
        flngObtieneClaveSocio = rsSocio!intcvesocio
    Else
        flngObtieneClaveSocio = -1
    End If
    rsSocio.Close
End Function
Private Sub pLlenaInformacionSocio(lngCveSocio As Long)
    Dim strParametrosSP As String
    Dim rsSocio As New ADODB.Recordset
    Dim rsCargos As New ADODB.Recordset
    Dim intcontador As Integer
    
    '|  Consulta la información del socio
    strParametrosSP = CStr(lngCveSocio) & "|T"
    Set rsSocio = frsEjecuta_SP(strParametrosSP, "SP_SORPTSELSOCIO")
    '|  Si existe la información del socio
    If rsSocio.RecordCount > 0 Then
        pReinicia
        With rsSocio
            txtClaveSocio.Text = lngCveSocio
            txtClaveUnica.Text = IIf(IsNull(!VCHCLAVESOCIO), "", !VCHCLAVESOCIO)
            lblSocio.Caption = IIf(IsNull(!Nombre), "", !Nombre)
            txtRFC.Text = IIf(IsNull(!RFC), "", !RFC)
            lblDomicilio.Caption = IIf(IsNull(!Domicilio), "", !Domicilio)
            lblNumeroExterior.Caption = IIf(IsNull(!NumeroExterior), "", !NumeroExterior)
            lblNumeroInterior.Caption = IIf(IsNull(!NumeroInterior), "", !NumeroInterior)
            lblCP.Caption = IIf(IsNull(!CP), "", !CP)
            lblTelefono.Caption = IIf(IsNull(!Telefono), "", !Telefono)
            lblColonia.Caption = IIf(IsNull(!Colonia), "", !Colonia)
            lblCiudad.Caption = IIf(IsNull(!DescripcionCiudad), "", !DescripcionCiudad)
            llngCveCiudad = IIf(IsNull(!ClaveCiudad), 0, !ClaveCiudad)
        End With
        
        '| Consulta los cargos del socio
        strParametrosSP = CStr(lngCveSocio)
        Set rsCargos = frsEjecuta_SP(strParametrosSP, "SP_PVSELCARGOSDESOCIOS")
        '| Si el socio tiene cargos
        If rsCargos.RecordCount > 0 Then
            With grdCargos
                Do While Not rsCargos.EOF
                    If .RowData(1) <> -1 Then
                         .Rows = .Rows + 1
                         .Row = .Rows - 1
                    End If
                    .RowData(.Row) = rsCargos!ClaveCargo
                    .TextMatrix(.Row, cintColCveCargo) = rsCargos!ClaveCargo
                    .TextMatrix(.Row, cintColFechaCargo) = Format(rsCargos!fecha, "dd/mmm/yyyy") & " " & FormatDateTime(rsCargos!fecha, vbShortTime)
                    .TextMatrix(.Row, cintColDescripcionCargo) = rsCargos!Concepto
                    .TextMatrix(.Row, cintColCveConceptoCargo) = rsCargos!ClaveConceptoFactura
                    .TextMatrix(.Row, cintColCantidadCargo) = FormatCurrency(rsCargos!Monto, 2)
                    .TextMatrix(.Row, cintIVACargo) = FormatCurrency(rsCargos!IVA, 2)
                    
                    '|  Si el precio fue modificado se pone de color amarillo
                    If rsCargos!PrecioManual = 1 Then
                        For intcontador = 1 To grdCargos.Cols - 1
                            .Col = intcontador
                            .CellBackColor = &H80000018
                        Next
                    End If
                    rsCargos.MoveNext
                Loop
            End With
            lblnConsulta = True
            pHabilita 0, 0, 0, 0, 0, 1, 0
            rsCargos.Close
            pLlenaFactura
            pLlenaPagos
            pCalculaTotales
            If fblnCanFocus(cmdSave) Then cmdSave.SetFocus
        Else
            '|  No existen cargos pendientes de facturar.
            MsgBox SIHOMsg(288), vbInformation, "Mensaje"
            lblnConsulta = True
        End If
    Else
        '|  ¡No existe información!
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
        pReinicia
    End If
End Sub
Private Sub pLlenaFactura()
    Dim intCargo As Integer
    Dim intConcepto As Integer
    Dim blnConceptoNuevo As Boolean
    Dim dblCantidadAcumuladaConcepto As Double
    Dim dblCantidadCargo As Double
    Dim dblCantidadAcumuladaIVAConcepto As Double
    Dim dblCantidadIVACargo As Double
    Dim strParametrosSP As String
    Dim rsConcepto As New ADODB.Recordset
    
    pLimpiaConceptos
    '|  Recorre los cargos
    For intCargo = 1 To grdCargos.Rows - 1
        blnConceptoNuevo = True
        '|  Recorre los conceptos de facturación
        For intConcepto = 1 To grdConceptos.Rows - 1
            '|  Si el concepto de facturación ya existe
            If grdCargos.TextMatrix(intCargo, cintColCveConceptoCargo) = grdConceptos.TextMatrix(intConcepto, cintColCveConcepto) Then
                '|  Acumula la cantidad del cargo
                dblCantidadAcumuladaConcepto = grdConceptos.TextMatrix(intConcepto, cintColCantidadConcepto)
                dblCantidadCargo = grdCargos.TextMatrix(intCargo, cintColCantidadCargo)
                dblCantidadAcumuladaConcepto = dblCantidadAcumuladaConcepto + dblCantidadCargo
                grdConceptos.TextMatrix(intConcepto, cintColCantidadConcepto) = FormatCurrency(dblCantidadAcumuladaConcepto, 2)
                '|  Acumula el IVA
                dblCantidadAcumuladaIVAConcepto = grdConceptos.TextMatrix(intConcepto, cintColIVAConcepto)
                dblCantidadIVACargo = grdCargos.TextMatrix(intCargo, cintIVACargo)
                dblCantidadAcumuladaIVAConcepto = dblCantidadAcumuladaIVAConcepto + dblCantidadIVACargo
                grdConceptos.TextMatrix(intConcepto, cintColIVAConcepto) = FormatCurrency(dblCantidadAcumuladaIVAConcepto, 2)
                
                blnConceptoNuevo = False
                Exit For
            End If
        Next
        '|  Si no se ha agregado el concepto de facturación lo agrega
        If blnConceptoNuevo Then
            strParametrosSP = grdCargos.TextMatrix(intCargo, cintColCveConceptoCargo) & "|-1|-1|" & vgintClaveEmpresaContable
            Set rsConcepto = frsEjecuta_SP(strParametrosSP, "sp_PvSelConceptoFacturacion")
            If rsConcepto!INTCUENTACONTABLE <> 0 Then
                grdConceptos.AddItem vbTab & _
                                     rsConcepto!smicveconcepto & vbTab & _
                                     rsConcepto!chrDescripcion & vbTab & _
                                     grdCargos.TextMatrix(intCargo, cintColCantidadCargo) & vbTab & _
                                     grdCargos.TextMatrix(intCargo, cintIVACargo) & vbTab & _
                                     rsConcepto!INTCUENTACONTABLE
            Else
                'No existe cuenta contable para el concepto de facturación:
                MsgBox SIHOMsg(907) & grdCargos.TextMatrix(intCargo, cintColCveConceptoCargo), vbOKOnly + vbInformation, "Mensaje"
            End If
        End If
    Next
End Sub
Private Sub pLimpiaConceptos()
    With grdConceptos
        .Clear
        .Cols = cintColsgrdCargos
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = cstrFormatoConcepto
        .ColWidth(0) = 100  ' Fix
        .ColWidth(cintColCveConcepto) = 0    ' Clave
        .ColWidth(cintColDescripcionConcepto) = 7150 ' Concepto de facturación
        .ColWidth(cintColCantidadConcepto) = 2000  ' Cantidad
        .ColWidth(cintColIVAConcepto) = 1500  ' IVA
        .ColWidth(cintColCtaIngresoConcepto) = 0  ' Cuenta de ingreso
        
        .ColAlignment(cintColDescripcionConcepto) = flexAlignLeftCenter
        .ColAlignment(cintColCantidadConcepto) = flexAlignRightCenter
        .ColAlignment(cintColIVAConcepto) = flexAlignRightCenter
        
        .FixedAlignment(cintColDescripcionConcepto) = flexAlignCenterCenter
        .FixedAlignment(cintColCantidadConcepto) = flexAlignCenterCenter
        .FixedAlignment(cintColIVAConcepto) = flexAlignCenterCenter
                
        .ScrollBars = flexScrollBarVertical
    End With
End Sub
Private Sub pConfiguraGridCargos()
On Error GoTo NotificaError

    With grdCargos
        .Clear
        .Cols = cintColsgrdCargos
        .Rows = 2
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = cstrFormatoCargo
        .ColWidth(0) = 100                        '| Fix
        .ColWidth(cintColCveCargo) = 0            '| Clave
        .ColWidth(cintColFechaCargo) = 2000       '| Fecha del cargo
        .ColWidth(cintColDescripcionCargo) = 5150 '| Otro concepto
        .ColWidth(cintColCantidadCargo) = 2000    '| Cantidad
        .ColWidth(cintIVACargo) = 1000            '| IVA
        .ColWidth(cintColCveConceptoCargo) = 0    '| Clave del concepto
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .FixedAlignment(4) = flexAlignCenterCenter
        .FixedAlignment(5) = flexAlignCenterCenter
                
        .ScrollBars = flexScrollBarBoth
        
        .RowData(1) = -1
        .Row = 1
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
End Sub
Private Sub pAsignaTotales()
    lblSubtotal.Caption = FormatCurrency(ldblCantidadFactura - ldblDescuentosFactura, 2)
    lblIVA.Caption = FormatCurrency(ldblIVAFactura, 2)
    lblTotal.Caption = FormatCurrency(ldblCantidadFactura - ldblDescuentosFactura + ldblIVAFactura, 2)
End Sub
Private Sub pCalculaTotales()
    Dim intCont As Integer
    Dim dblSubTotal As Double
    Dim dblTotalIVA As Double
    Dim dblTotalPagos As Double
    Dim dblTotalFactura As Double
        
    dblSubTotal = 0
    dblTotalIVA = 0
    
    For intCont = 1 To grdCargos.Rows - 1
        dblSubTotal = dblSubTotal + CDbl(grdCargos.TextMatrix(intCont, cintColCantidadCargo))
        dblTotalIVA = dblTotalIVA + CDbl(grdCargos.TextMatrix(intCont, cintIVACargo))
    Next
    '|  Subtotal
    lblSubtotal.Caption = FormatCurrency(dblSubTotal, 2)
    '|  IVA
    lblIVA.Caption = FormatCurrency(dblTotalIVA, 2)
    '|  Total de la factura
    dblTotalFactura = dblSubTotal + dblTotalIVA
    lblTotalFactura.Caption = FormatCurrency(dblTotalFactura, 2)
    '|  Pagos
    dblTotalPagos = fdblTotalPagos(txtClaveSocio.Text)
    lblTotalPagos = FormatCurrency(dblTotalPagos, 2)
    '|  Total a pagar
    lblTotal.Caption = FormatCurrency(dblTotalFactura - dblTotalPagos, 2)
End Sub
Private Function fdblTotalPagos(lngCveSocio As Long) As Double
    Dim strParametrosSP As String
    Dim rsPagos As New ADODB.Recordset
    Dim dblPagos As Double
    
    dblPagos = 0
    strParametrosSP = "*|P|S|" & fstrFechaSQL("01-01-1900") & "|" & fstrFechaSQL("01-01-1900") & "|" & lngCveSocio & "|" & vgintClaveEmpresaContable & "|1"
    Set rsPagos = frsEjecuta_SP(strParametrosSP, "SP_PVSELENTRADASALIDADINEROFEC")
    While Not rsPagos.EOF
        If rsPagos!Cancelado = 0 Then
            '|  Sumariza la cantidad del pago y resta la salida de dinero, si fue en dólares hace la conversión
            dblPagos = dblPagos + (rsPagos!Cantidad * IIf(rsPagos!TipoMovimiento = "E", 1, -1) * IIf(rsPagos!Moneda = "Pesos", 1, rsPagos!TipoCambio))
        End If
        rsPagos.MoveNext
    Wend
    fdblTotalPagos = dblPagos
    rsPagos.Close
End Function
Private Sub pLlenaPoliza(lngnumCuenta As Long, dblCantidad As Double, intTipoMovto As Integer)
    Dim intTamaño As Integer
    Dim intPosicion As Integer
    Dim blnEstaCuenta As Boolean
    Dim intcontador As Integer
    
    If apoliza(0).lngnumCuenta = 0 Then
        apoliza(0).lngnumCuenta = lngnumCuenta
        apoliza(0).dblCantidad = dblCantidad
        apoliza(0).intNaturaleza = intTipoMovto
    Else
        blnEstaCuenta = False
        intTamaño = UBound(apoliza(), 1)
        For intcontador = 0 To intTamaño
            If apoliza(intcontador).lngnumCuenta = lngnumCuenta Then
                blnEstaCuenta = True
                intPosicion = intcontador
            End If
        Next intcontador
        If blnEstaCuenta Then
            If apoliza(intPosicion).intNaturaleza = intTipoMovto Then
                apoliza(intPosicion).dblCantidad = apoliza(intPosicion).dblCantidad + dblCantidad
            Else
                If apoliza(intPosicion).dblCantidad > dblCantidad Then
                    apoliza(intPosicion).dblCantidad = apoliza(intPosicion).dblCantidad - dblCantidad
                Else
                    If apoliza(intPosicion).dblCantidad < dblCantidad Then
                        apoliza(intPosicion).dblCantidad = dblCantidad - apoliza(intPosicion).dblCantidad
                        If apoliza(intPosicion).intNaturaleza = 1 Then
                            apoliza(intPosicion).intNaturaleza = 0
                        Else
                            apoliza(intPosicion).intNaturaleza = 1
                        End If
                    Else
                        apoliza(intPosicion).lngnumCuenta = 0
                        apoliza(intPosicion).dblCantidad = 0
                        apoliza(intPosicion).intNaturaleza = 0
                    End If
                End If
            End If
        Else
            ReDim Preserve apoliza(intTamaño + 1)
            apoliza(intTamaño + 1).lngnumCuenta = lngnumCuenta
            apoliza(intTamaño + 1).dblCantidad = dblCantidad
            apoliza(intTamaño + 1).intNaturaleza = intTipoMovto
        End If
    End If
End Sub
Private Function fintErrorContable(strFecha As String) As Integer
    '----------------------------------------------------------------------------------------------
    'Función para revisar que se pueda introducir una póliza o que el periodo contable esté abierto
    '----------------------------------------------------------------------------------------------
    Dim lngResultado As Long        'Resultado de la validación del periodo contable
    
    fintErrorContable = 0
    
    lngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "SP_CNUPDESTATUSCIERRE", True, lngResultado
    If lngResultado = 1 Then
        If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(strFecha)), Month(CDate(strFecha))) Then
            fintErrorContable = 209 'El periodo contable esta cerrado.
            Exit Function
        End If
    Else
        fintErrorContable = 720 'No se puede realizar la operación, inténtelo en unos minutos.
        Exit Function
    End If
End Function
Sub pLimpiaGrid(ObjGrd As VSFlexGrid)
    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .FormatString = ""
        .Rows = 2
        .Row = 1
        .Col = 1
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
End Sub
Private Sub pConfiguraGridPagos()
    With grdPagos
        .Cols = 12
        .FixedCols = 2
        .FixedRows = 1
        .FormatString = "|Concepto del pago|Fecha|Cantidad|Moneda|Recibo|Factura"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 4000 'Concepto del pago
        .ColWidth(2) = 1430 'Fecha
        .ColWidth(3) = 1430 'Cantidad
        .ColWidth(4) = 700  'Moneda
        .ColWidth(5) = 1430 'Recibo
        .ColWidth(6) = 1500 'Factura
        .ColWidth(7) = 0    'Tipo de Pago que se realizó
        .ColWidth(8) = 0    '*Disponible*
        .ColWidth(9) = 0    'Si es un pago "E" o una devolución "S" (Entrada o Salida)
        .ColWidth(10) = 0   'El numero de corte
        .ColWidth(11) = 0   'La cuenta contable segun el Concepto de pago/salida de dinero
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignLeftCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        .FixedAlignment(3) = flexAlignCenterCenter
        .FixedAlignment(4) = flexAlignCenterCenter
        .FixedAlignment(5) = flexAlignCenterCenter
        .FixedAlignment(6) = flexAlignCenterCenter
        .FixedAlignment(7) = flexAlignCenterCenter
        .FixedAlignment(8) = flexAlignCenterCenter
        .FixedAlignment(9) = flexAlignCenterCenter
        .FixedAlignment(10) = flexAlignCenterCenter
        .FixedAlignment(11) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub
Private Function fdblPagosFacturados(lngFacturaId As Long)
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    vlstrSentencia = "Select NVL(MNYANTICIPO, 0) Pagos " & _
                     "  From PvFactura " & _
                     " Where INTCONSECUTIVO = " & CStr(lngFacturaId)
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    fdblPagosFacturados = rs!pagos
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaPagos"))
End Function
Private Sub txtClaveUnica_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtpendientetimbre_Change()
If fblnCanFocus(Me.cmdLocate) Then cmdLocate.SetFocus
End Sub

Private Sub pCargaUsosCFDI()
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frsCatalogoSAT("c_UsoCFDI")
    If Not rsTmp.EOF Then
        pLlenarCboRs cboUsoCFDI, rsTmp, 0, 1
        cboUsoCFDI.ListIndex = -1
    End If
End Sub
