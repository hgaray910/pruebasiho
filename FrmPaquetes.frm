VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMtoPaquetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de paquetes, planes y cirugías"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12960
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDgArchivo 
      Left            =   1800
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   9960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FreTotales 
      Enabled         =   0   'False
      Height          =   1255
      Left            =   120
      TabIndex        =   54
      Top             =   3360
      Width           =   3495
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
         Left            =   1495
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Iva del paquete"
         Top             =   520
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
         Left            =   1495
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "Subtotal del paquete"
         Top             =   210
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
         Left            =   1495
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Total del paquete"
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label Label11 
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
         Left            =   90
         TabIndex        =   66
         Top             =   542
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
         Left            =   90
         TabIndex        =   65
         Top             =   232
         Width           =   870
      End
      Begin VB.Label Label59 
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
         Left            =   90
         TabIndex        =   64
         Top             =   862
         Width           =   555
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1335
      Left            =   2400
      TabIndex        =   0
      Top             =   7920
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbCargando 
         Height          =   360
         Left            =   165
         TabIndex        =   44
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
         Left            =   120
         TabIndex        =   45
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
   Begin TabDlg.SSTab SSTObj 
      Height          =   9795
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   13080
      _ExtentX        =   23072
      _ExtentY        =   17277
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   732
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "FrmPaquetes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "SysInfo1"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Contenido del paquete"
      TabPicture(1)   =   "FrmPaquetes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrmMtoPaquetes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdTotales"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Departamentos para registro de paciente"
      TabPicture(2)   =   "FrmPaquetes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(2)=   "lstDepartamentos"
      Tab(2).Control(3)=   "lstDepartamentosSel"
      Tab(2).Control(4)=   "cmdEliminaTodo"
      Tab(2).Control(5)=   "cmdEliminaUno"
      Tab(2).Control(6)=   "cmdAsignaUno"
      Tab(2).Control(7)=   "cmdAsignaTodo"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Honorarios médicos del paquete de cirugía"
      TabPicture(3)   =   "FrmPaquetes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "fraHonorarios"
      Tab(3).Control(2)=   "Frame9"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Búsqueda de paquetes"
      TabPicture(4)   =   "FrmPaquetes.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label17"
      Tab(4).Control(1)=   "grdHBusqueda"
      Tab(4).Control(2)=   "txtBusqueda"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Paquetes de presupuestos"
      TabPicture(5)   =   "FrmPaquetes.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdPresupuestos"
      Tab(5).Control(1)=   "Frame10"
      Tab(5).Control(2)=   "fraRangoFolios"
      Tab(5).Control(3)=   "frmFechas"
      Tab(5).ControlCount=   4
      Begin VB.Frame frmFechas 
         Caption         =   "Rango de fechas de presupuestos"
         Height          =   760
         Left            =   -74790
         TabIndex        =   113
         Top             =   720
         Width           =   4455
         Begin MSMask.MaskEdBox mskFechaFin 
            Height          =   315
            Left            =   2925
            TabIndex        =   117
            ToolTipText     =   "Fecha final"
            Top             =   315
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaInicio 
            Height          =   315
            Left            =   840
            TabIndex        =   118
            ToolTipText     =   "Fecha inicial"
            Top             =   315
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label22 
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2290
            TabIndex        =   115
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label21 
            Caption         =   "Desde"
            Height          =   195
            Left            =   255
            TabIndex        =   114
            Top             =   375
            Width           =   600
         End
      End
      Begin VB.Frame fraRangoFolios 
         Caption         =   "Rango de números de presupuestos"
         Height          =   765
         Left            =   -70110
         TabIndex        =   109
         Top             =   720
         Width           =   4335
         Begin VB.TextBox txtFolioInicial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   900
            TabIndex        =   107
            ToolTipText     =   "Número de presupuesto inicial"
            Top             =   315
            Width           =   1215
         End
         Begin VB.TextBox txtFolioFinal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2895
            TabIndex        =   108
            ToolTipText     =   "Número de presupuesto final"
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label lblDesde 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   255
            TabIndex        =   112
            Top             =   375
            Width           =   465
         End
         Begin VB.Label lblHasta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2280
            TabIndex        =   111
            Top             =   375
            Width           =   420
         End
      End
      Begin VB.Frame Frame10 
         Height          =   780
         Left            =   -65580
         TabIndex        =   106
         Top             =   720
         Width           =   735
         Begin VB.CommandButton cmdBuscarPresupuestos 
            Height          =   480
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":00A8
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Búsqueda"
            Top             =   190
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   -74400
         TabIndex        =   104
         Top             =   480
         Width           =   11775
         Begin VB.OptionButton optFormaAsignacionSeleccionar 
            Caption         =   "Permitir seleccionar cuales aplican"
            Height          =   255
            Left            =   6480
            TabIndex        =   89
            ToolTipText     =   "Incluir en la admisión (prepagos y convenio)"
            Top             =   220
            Width           =   3375
         End
         Begin VB.OptionButton optFormaAsignacionTodos 
            Caption         =   "Todos al asignar el paquete"
            Height          =   255
            Left            =   3840
            TabIndex        =   88
            Top             =   220
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.Label Label18 
            Caption         =   "Forma de asignación en el registro del paciente"
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   -73080
         TabIndex        =   103
         ToolTipText     =   "Buscar"
         Top             =   480
         Width           =   4215
      End
      Begin VB.Frame fraHonorarios 
         Height          =   3375
         Left            =   -74400
         TabIndex        =   101
         Top             =   3000
         Width           =   11775
         Begin VB.CommandButton cmdModificar 
            Caption         =   "Modificar"
            Height          =   495
            Left            =   9720
            TabIndex        =   96
            ToolTipText     =   "Modificar honorario médico"
            Top             =   2760
            Width           =   1215
         End
         Begin VB.CommandButton cmdBorrarHonorario 
            Height          =   495
            Left            =   11040
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":021A
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Borrar"
            Top             =   2760
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VSFlex7LCtl.VSFlexGrid grdDetalleHonorarios 
            Height          =   2325
            Left            =   120
            TabIndex        =   94
            ToolTipText     =   "Honorarios médicos que forman parte del paquete"
            Top             =   240
            Width           =   11415
            _cx             =   20135
            _cy             =   4101
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
            BackColorBkg    =   -2147483636
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
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPaquetes.frx":03BC
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
      End
      Begin VB.Frame Frame5 
         Height          =   1935
         Left            =   -74400
         TabIndex        =   87
         Top             =   1080
         Width           =   11775
         Begin VB.CommandButton cmdAgregar 
            Height          =   480
            Left            =   7200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":0451
            Style           =   1  'Graphical
            TabIndex        =   93
            ToolTipText     =   "Agregar honorario médico"
            Top             =   1320
            UseMaskColor    =   -1  'True
            Width           =   525
         End
         Begin VB.ComboBox cboFuncion 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   90
            ToolTipText     =   "Funciones de los participantes de cirugía"
            Top             =   240
            Width           =   5415
         End
         Begin VB.ComboBox cboOtroConcepto 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   91
            ToolTipText     =   "Concepto de cargo"
            Top             =   600
            Width           =   5415
         End
         Begin VB.TextBox txtImporteHonorario 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2280
            TabIndex        =   92
            ToolTipText     =   "Cantidad del importe del honorario médico"
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Función de cirugía"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   300
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "Otro concepto de cargo"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   650
            Width           =   2055
         End
         Begin VB.Label Label16 
            Caption         =   "Importe del honorario médico"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   1000
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdAsignaTodo 
         Height          =   495
         Left            =   -68745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmPaquetes.frx":0943
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Agregar todos los departamentos"
         Top             =   1400
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAsignaUno 
         Height          =   495
         Left            =   -68745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmPaquetes.frx":0AB5
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Agregar el departamento seleccionado"
         Top             =   1950
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdEliminaUno 
         Height          =   495
         Left            =   -68745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmPaquetes.frx":0C27
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Quitar el departamento seleccionado"
         Top             =   2550
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdEliminaTodo 
         Height          =   495
         Left            =   -68745
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FrmPaquetes.frx":0D99
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Quitar todos los departamentos"
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.ListBox lstDepartamentosSel 
         Height          =   3570
         Left            =   -68115
         Sorted          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Departamentos que podrán asignar el paquete al paciente al momento de su registro"
         Top             =   840
         Width           =   3810
      End
      Begin VB.ListBox lstDepartamentos 
         Height          =   3570
         Left            =   -72720
         Sorted          =   -1  'True
         TabIndex        =   58
         ToolTipText     =   "Departamentos activos clasificados como administrativos, de enfermería o bien almacenes que no son de tipo consignación"
         Top             =   840
         Width           =   3810
      End
      Begin VSFlex7LCtl.VSFlexGrid grdTotales 
         Height          =   2175
         Left            =   2040
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   7560
         Width           =   10920
         _cx             =   19262
         _cy             =   3836
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   14
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPaquetes.frx":0F0B
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ComboSearch     =   0
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   1
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Frame Frame6 
         Caption         =   "Agrupar reporte por "
         Height          =   1255
         Left            =   -64800
         TabIndex        =   70
         Top             =   3360
         Width           =   2535
         Begin VB.OptionButton optAgruparRpt 
            Caption         =   "Departamento"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   20
            ToolTipText     =   "Agrupar por departamento"
            Top             =   900
            Width           =   2175
         End
         Begin VB.OptionButton optAgruparRpt 
            Caption         =   "Concepto de facturación"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            ToolTipText     =   "Agrupar por concepto de facturación"
            Top             =   550
            Width           =   2175
         End
         Begin VB.OptionButton optAgruparRpt 
            Caption         =   "Tipo de cargo"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            ToolTipText     =   "Agrupar por tipo de cargo"
            Top             =   250
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin SysInfoLib.SysInfo SysInfo1 
         Left            =   -64320
         Top             =   7380
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame Frame4 
         Height          =   690
         Left            =   -70080
         TabIndex        =   52
         Top             =   3900
         Width           =   3675
         Begin VB.CommandButton cmdPrint 
            Height          =   480
            Left            =   3105
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":0F93
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Imprimir"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   60
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":1135
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":12A7
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1080
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":1419
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1590
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":158B
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2100
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":16FD
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Enabled         =   0   'False
            Height          =   480
            Left            =   2610
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmPaquetes.frx":186F
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Guardar el registro"
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2865
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   12630
         Begin VB.TextBox txtFechaAltaPaquete 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   11230
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "01/Ene/1999"
            Top             =   2485
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox ChkValidaPaquete 
            Caption         =   "Validar cargos permitidos en paquete"
            Height          =   315
            Left            =   8925
            TabIndex        =   9
            ToolTipText     =   "Validar cargos permitidos en paquete"
            Top             =   2180
            Width           =   3495
         End
         Begin VB.ComboBox cboMovConceptoFactura 
            ForeColor       =   &H80000012&
            Height          =   315
            IntegralHeight  =   0   'False
            ItemData        =   "FrmPaquetes.frx":1BB1
            Left            =   3960
            List            =   "FrmPaquetes.frx":1BB3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            ToolTipText     =   "Concepto de facturación"
            Top             =   2160
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.CheckBox chkprecio 
            Caption         =   "Actualizar automáticamente costos y precios"
            Height          =   315
            Left            =   8925
            TabIndex        =   8
            ToolTipText     =   "Activar actualización automática de costos y precios"
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txtAnticipo 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   10600
            MaxLength       =   15
            TabIndex        =   6
            ToolTipText     =   "Anticipo sugerido"
            Top             =   945
            Width           =   1815
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            ItemData        =   "FrmPaquetes.frx":1BB5
            Left            =   10600
            List            =   "FrmPaquetes.frx":1BC2
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Tipo de paquete"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cboTratamiento 
            Height          =   315
            ItemData        =   "FrmPaquetes.frx":1BDE
            Left            =   10600
            List            =   "FrmPaquetes.frx":1BE8
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Tipo de tratamiento"
            Top             =   255
            Width           =   1815
         End
         Begin VB.ComboBox cboConceptoFactura 
            Height          =   315
            Left            =   1800
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Concepto de facturación utilizado para este paquete"
            Top             =   945
            Width           =   6975
         End
         Begin VB.TextBox txtCvePaquete 
            Height          =   315
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   1
            ToolTipText     =   "Clave única del paquete"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1800
            MaxLength       =   100
            TabIndex        =   2
            ToolTipText     =   "Descripción del paquete"
            Top             =   600
            Width           =   6975
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdGrupos 
            Height          =   1120
            Left            =   180
            TabIndex        =   21
            ToolTipText     =   "Conceptos de facturación de los grupos de cargo"
            Top             =   1560
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   1984
            _Version        =   393216
            Cols            =   3
            AllowBigSelection=   0   'False
            FocusRect       =   0
            HighLight       =   2
            BandDisplay     =   1
            RowSizingMode   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).BandIndent=   5
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSMask.MaskEdBox mskFecha 
            Height          =   315
            Left            =   11230
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Fecha de última actualización de costos y precios"
            Top             =   1545
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            Height          =   315
            Left            =   8925
            TabIndex        =   10
            ToolTipText     =   "Estatus del paquete Activo o Inactivo"
            Top             =   2430
            Width           =   745
         End
         Begin VB.Label lblFechaAltaPaquete 
            Caption         =   "Fecha de alta "
            Height          =   195
            Left            =   10080
            TabIndex        =   119
            Top             =   2485
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "costos y precios"
            Height          =   255
            Left            =   8920
            TabIndex        =   82
            Top             =   1605
            Width           =   2880
         End
         Begin VB.Label Label4 
            Caption         =   "Última actualización de"
            Height          =   315
            Left            =   8920
            TabIndex        =   77
            Top             =   1380
            Width           =   2880
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Conceptos de facturación para excedente por grupos de cargos"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   1335
            Width           =   4530
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Anticipo sugerido"
            Height          =   195
            Left            =   8920
            TabIndex        =   69
            Top             =   1005
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   8920
            TabIndex        =   68
            Top             =   660
            Width           =   315
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Tratamiento"
            Height          =   195
            Left            =   8920
            TabIndex        =   67
            Top             =   315
            Width           =   840
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Concepto de factura"
            Height          =   195
            Left            =   180
            TabIndex        =   53
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   315
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   660
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cargos que incluye el paquete "
         Height          =   4075
         Left            =   105
         TabIndex        =   46
         Top             =   3480
         Width           =   12820
         Begin VB.CommandButton cmdImportar 
            Caption         =   "Importar"
            Height          =   405
            Left            =   11590
            TabIndex        =   122
            ToolTipText     =   "Importar de formato Excel"
            Top             =   3590
            Width           =   1095
         End
         Begin VB.CommandButton cmdExportar 
            Caption         =   "Exportar"
            Height          =   405
            Left            =   10490
            TabIndex        =   121
            ToolTipText     =   "Exportar en formato Excel"
            Top             =   3590
            Width           =   1095
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   240
            MaxLength       =   3
            TabIndex        =   72
            Top             =   2160
            Visible         =   0   'False
            Width           =   615
         End
         Begin VSFlex7LCtl.VSFlexGrid grdPaquete 
            Height          =   3300
            Left            =   85
            TabIndex        =   85
            ToolTipText     =   "Elementos incluidos en el paquete"
            Top             =   240
            Width           =   12645
            _cx             =   22304
            _cy             =   5821
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
            BackColorBkg    =   -2147483636
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   32
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPaquetes.frx":1C00
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   2
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   0
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
            AutoSizeMouse   =   0   'False
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            Begin VB.TextBox txtMontoLimite 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   120
               TabIndex        =   86
               Top             =   1440
               Visible         =   0   'False
               Width           =   615
            End
         End
      End
      Begin VB.Frame FrmMtoPaquetes 
         Height          =   2985
         Left            =   105
         TabIndex        =   47
         Top             =   480
         Width           =   12820
         Begin VB.Frame Frame8 
            Caption         =   "Tipo de ingreso para descuento en predeterminadas "
            Height          =   525
            Left            =   7560
            TabIndex        =   81
            Top             =   1775
            Width           =   5150
            Begin VB.OptionButton optTipoPacienteDesc 
               Caption         =   "Urgencias"
               Height          =   255
               Index           =   3
               Left            =   3840
               TabIndex        =   40
               ToolTipText     =   "Tipo de ingreso urgencias para descuentos"
               Top             =   220
               Width           =   1095
            End
            Begin VB.OptionButton optTipoPacienteDesc 
               Caption         =   "Externo"
               Height          =   255
               Index           =   2
               Left            =   2640
               TabIndex        =   39
               ToolTipText     =   "Tipo de ingreso externo para descuentos"
               Top             =   220
               Width           =   855
            End
            Begin VB.OptionButton optTipoPacienteDesc 
               Caption         =   "Interno"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   38
               ToolTipText     =   "Tipo de ingreso interno para descuentos"
               Top             =   220
               Width           =   855
            End
            Begin VB.OptionButton optTipoPacienteDesc 
               Caption         =   "Todos"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   37
               ToolTipText     =   "Todos los tipos de ingreso para descuentos"
               Top             =   220
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame7 
            Height          =   760
            Left            =   5840
            TabIndex        =   76
            Top             =   1955
            Width           =   1310
            Begin VB.CommandButton cmdIncluir 
               Caption         =   "Incluir"
               Height          =   540
               Index           =   3
               Left            =   75
               MaskColor       =   &H80000014&
               Picture         =   "FrmPaquetes.frx":1E78
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Incluir un cargo al paquete"
               Top             =   145
               UseMaskColor    =   -1  'True
               Width           =   570
            End
            Begin VB.CommandButton cmdExcluir 
               Caption         =   "Excluir"
               Height          =   540
               Index           =   2
               Left            =   660
               MaskColor       =   &H80000014&
               Picture         =   "FrmPaquetes.frx":1FD2
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Excluir un cargo al paquete"
               Top             =   145
               UseMaskColor    =   -1  'True
               Width           =   570
            End
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "Actualizar costos y precios"
            Height          =   425
            Left            =   10590
            TabIndex        =   43
            ToolTipText     =   "Actualizar costos y precios en predeterminadas y comparativa"
            Top             =   2425
            Width           =   2115
         End
         Begin VB.Frame frmCosto 
            Caption         =   "Costo base "
            Height          =   525
            Left            =   7560
            TabIndex        =   75
            Top             =   2340
            Width           =   2895
            Begin VB.OptionButton OptPolitica 
               Caption         =   "Última compra"
               Height          =   255
               Index           =   1
               Left            =   1520
               TabIndex        =   42
               ToolTipText     =   "Costo basado en la última compra"
               Top             =   220
               Width           =   1298
            End
            Begin VB.OptionButton OptPolitica 
               Caption         =   "Costo más alto"
               Height          =   255
               Index           =   0
               Left            =   100
               TabIndex        =   41
               ToolTipText     =   "Costo basado en el más alto"
               Top             =   220
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.Frame fraTipoPaciente 
            Caption         =   "Tipo de ingreso "
            Height          =   525
            Left            =   7680
            TabIndex        =   74
            Top             =   1100
            Width           =   4905
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Todos"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   33
               ToolTipText     =   "Todos los tipos de ingreso para comparación"
               Top             =   240
               Value           =   -1  'True
               Width           =   743
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Interno"
               Height          =   195
               Index           =   5
               Left            =   1320
               TabIndex        =   34
               ToolTipText     =   "Tipo de ingreso interno para comparación"
               Top             =   240
               Width           =   863
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Externo"
               Height          =   195
               Index           =   4
               Left            =   2520
               TabIndex        =   35
               ToolTipText     =   "Tipo de ingreso externo para comparación"
               Top             =   240
               Width           =   908
            End
            Begin VB.OptionButton optTipoPaciente 
               Caption         =   "Urgencias"
               Height          =   195
               Index           =   3
               Left            =   3720
               TabIndex        =   36
               ToolTipText     =   "Tipo de ingreso urgencias para comparación"
               Top             =   240
               Width           =   1013
            End
         End
         Begin VB.OptionButton optDescripcion 
            Caption         =   "&Descripción"
            Height          =   225
            Left            =   4320
            TabIndex        =   26
            ToolTipText     =   "Buscar por descripción"
            Top             =   600
            Value           =   -1  'True
            Width           =   1133
         End
         Begin VB.OptionButton optClave 
            Caption         =   "&Clave"
            Height          =   225
            Left            =   3480
            TabIndex        =   25
            ToolTipText     =   "Buscar por clave"
            Top             =   600
            Width           =   698
         End
         Begin VB.ListBox lstElementos 
            DragIcon        =   "FrmPaquetes.frx":212C
            Height          =   1815
            ItemData        =   "FrmPaquetes.frx":2576
            Left            =   195
            List            =   "FrmPaquetes.frx":2578
            TabIndex        =   28
            ToolTipText     =   "Elementos a agregar al paquete"
            Top             =   900
            Width           =   5370
         End
         Begin VB.TextBox txtSeleArticulo 
            Height          =   315
            Left            =   200
            TabIndex        =   24
            ToolTipText     =   "Teclee la clave o la descripción del elemento"
            Top             =   555
            Width           =   3210
         End
         Begin TabDlg.SSTab sstElementos 
            Height          =   2700
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Elementos que puede contener un paquete"
            Top             =   150
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   4763
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
            TabCaption(0)   =   "      Artículos      "
            TabPicture(0)   =   "FrmPaquetes.frx":257A
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "chkMedicamentos"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Servicios auxiliares"
            TabPicture(1)   =   "FrmPaquetes.frx":2596
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "     Laboratorio     "
            TabPicture(2)   =   "FrmPaquetes.frx":25B2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Otros conceptos"
            TabPicture(3)   =   "FrmPaquetes.frx":25CE
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            TabCaption(4)   =   "Grupos de cargos"
            TabPicture(4)   =   "FrmPaquetes.frx":25EA
            Tab(4).ControlEnabled=   -1  'True
            Tab(4).ControlCount=   0
            Begin VB.CheckBox chkMedicamentos 
               Caption         =   "Sólo medicamentos"
               Height          =   225
               Left            =   -69480
               TabIndex        =   27
               ToolTipText     =   "Filtrar sólo medicamentos"
               Top             =   450
               Value           =   1  'Checked
               Width           =   1680
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Selección de información para comparación de precios"
            Height          =   1580
            Left            =   7560
            TabIndex        =   78
            Top             =   150
            Width           =   5145
            Begin VB.ComboBox cboEmpresas 
               Height          =   315
               Left            =   1440
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   32
               ToolTipText     =   "Lista de empresas para comparación"
               Top             =   620
               Width           =   3600
            End
            Begin VB.ComboBox cboTipoPaciente 
               Height          =   315
               Left            =   1440
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   31
               ToolTipText     =   "Lista de procedencias para comparación"
               Top             =   240
               Width           =   3600
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Empresa"
               Height          =   195
               Left            =   100
               TabIndex        =   80
               Top             =   660
               Width           =   855
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de paciente"
               Height          =   195
               Left            =   100
               TabIndex        =   79
               Top             =   300
               Width           =   1200
            End
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         Height          =   3885
         Left            =   -74880
         TabIndex        =   100
         Top             =   840
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   6853
         _Version        =   393216
         GridColor       =   12632256
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPresupuestos 
         Height          =   2925
         Left            =   -74790
         TabIndex        =   116
         Top             =   1560
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   5159
         _Version        =   393216
         GridColor       =   12632256
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label17 
         Caption         =   "Buscar por descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   530
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Departamentos asignados"
         Height          =   195
         Left            =   -68160
         TabIndex        =   84
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Departamentos disponibles"
         Height          =   195
         Left            =   -72720
         TabIndex        =   83
         Top             =   600
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmMtoPaquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmMtoPaquetes
'-------------------------------------------------------------------------------------
'| Objetivo: Mantenimiento del catálogo de Paquetes
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 3/Noviembre/2000
'| Fecha última modificación: 10/Oct/2002
'-------------------------------------------------------------------------------------
'| Modificó                 : Rosenda Hernández Anaya
'| Fecha última modificación: 04/Agosto/2003
'|      Se modifica para que de mantenimiento a los nuevos campos de la tabla PvPaquete
'|------------------------------------------------------------------------------------

Option Explicit
Public rsPaquetes As New ADODB.Recordset

Dim vgstrEstadoManto As String
Dim vglngTipoParticular As Long

Dim vllngDesktop  As Long
Dim vllngSizeNormal As Long
Dim vllngSizeGrande As Long
Dim vllngSizeHonorarios As Long
Dim vlstrsql As String
Dim rspvselObtenerPrecio As New ADODB.Recordset
Private vgrptReporte As CRAXDRT.Report
Dim vlaryResultados() As String

Dim vgTipoIngresoDescuento As String

Dim vgdblPrecioPredGrupo As Double
Dim vgstrTipoPredGrupo As String
Dim vglngClavePredGrupo As Long
Dim vgdtmFechaActualizacion As Date

Dim vgdblPrecioConvGrupo As Double
Dim vgstrTipoConvGrupo As String
Dim vglngClaveConvGrupo As Long
Dim vgblnValida As Boolean

Dim vlintValidarCargosEnPaquete As Integer

Dim vlintCambioPrecio As Integer

Private Const CB_SETITEMHEIGHT = &H153
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim vlblnAgregaNuevo As Boolean
Dim vlintCantCargo As Integer

'Tipo de Permiso en String
Dim vlstrTipoPermiso As String

''Honorarios
Const cintColClaveFuncion = 1
Const cintColDescripcionFuncion = 2
Const cintColClaveConcepto = 3
Const cintColDescripcionConcepto = 4
Const cintColImporteHonorario = 5
Const cintColEstadoRegistro = 6        'Numero de columna para indicar si es 1=nuevo 2 = editado 3 = borrar

Dim vllngCveConceptoFactHonorarioMedico As Long 'Valor del parametro con la clave concepto de facturacion de Honorarios médicos
Dim vlintEstadoRegistro As Integer
Dim vlngRowActualizar As Long 'Número de row en el grid del elemento que se selecciona para modificar

Const intTotalColsgrdPresupuesto = 7
Const strFormatgrdPaquete = "|Presupuesto|Fecha|Paquete|Descripción del paquete|Cuenta del paciente|Nombre del paciente"
Const ColRowDatagrdPaquete = 0         'RowData
Const ColNumeroPresupuesto = 1         'Número del presupuesto
Const ColFechaPresupuesto = 2          'Fecha del presupuesto
Const ColNumeroPaquete = 3             'Numero del paquete
Const ColDescripcionPaquete = 4        'Descripción del paquete
Const ColNumCuentaPaciente = 5         'Número de cuenta del paciente
Const ColNombrePaciente = 6            'Nombre del paciente

Private Sub pConfiguraGridPresupuestos()
    With grdPresupuestos
    
        .Cols = intTotalColsgrdPresupuesto
        .Rows = 2
        .Clear

        .FormatString = strFormatgrdPaquete
        .ColWidth(0) = 0        'Fix
        .ColWidth(1) = 1200      'Número del presupuesto
        .ColWidth(2) = 1100     'Fecha
        .ColWidth(3) = 1000     'Numero del paquete
        .ColWidth(4) = 3500     'Descripción del paquete
        .ColWidth(5) = 1700     'Número de cuenta del paciente
        .ColWidth(6) = 3700     'Nombre del paciente
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        
'        .ColAlignmentFixed(1) = flexAlignCenterCenter
'        .ColAlignmentFixed(2) = flexAlignCenterCenter
'        .ColAlignmentFixed(3) = flexAlignCenterCenter
'        .ColAlignmentFixed(4) = flexAlignCenterCenter
'        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With

End Sub


Private Sub pLlenaGridPresupuestos()
    On Error GoTo NotificaError

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlstrFolioInicio As String
    Dim vlstrFolioFin As String

    'Validar el rango de fechas
    If Not IsDate(mskFechaInicio) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        pEnfocaMkTexto mskFechaInicio
        Exit Sub
    End If
    'Validar el rango de fechas
    If Not IsDate(mskFechaFin) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        pEnfocaMkTexto mskFechaFin
        Exit Sub
    End If
    If IsDate(mskFechaFin) And IsDate(mskFechaInicio) Then
        If CDate(mskFechaFin) < CDate(mskFechaInicio) Then
            '¡La fecha final debe ser mayor a la fecha inicial!
            MsgBox SIHOMsg(379), vbExclamation, "Mensaje"
            pEnfocaMkTexto mskFechaFin
            Exit Sub
        End If
    End If
    If Val(txtFolioInicial.Text) <= 0 Then
        '¡El folio final debe ser mayor a 0!
        MsgBox Replace(SIHOMsg(943), "folio", "número de presupuesto"), vbExclamation, "Mensaje"
        pEnfocaTextBox txtFolioInicial
        Exit Sub
    End If
    If Val(txtFolioFinal.Text) <= 0 Then
        '¡El folio inicial debe ser mayor a 0!
        MsgBox Replace(SIHOMsg(944), "folio", "número de presupuesto"), vbExclamation, "Mensaje"
        pEnfocaTextBox txtFolioInicial
        Exit Sub
    End If
    If Val(txtFolioInicial.Text) > Val(txtFolioFinal.Text) Then
        MsgBox Replace(SIHOMsg(201), "folio", "número de presupuesto"), vbExclamation, "Mensaje"
        pEnfocaTextBox txtFolioInicial
        Exit Sub
    End If
    
    vlstrFolioInicio = Trim(str(Val(txtFolioInicial.Text)))
    vlstrFolioFin = Trim(str(Val(txtFolioFinal.Text)))

    pConfiguraGridPresupuestos
    
    vlstrSentencia = "SELECT PvPresupuesto.intcvepresupuesto, " & _
                        "TO_CHAR (dtmfechapresupuesto, 'DD/MM/YYYY') AS fechapresupuesto, " & _
                        "PvPaquete.chrdescripcion descripcionPaquete, PvPresupuesto.intnumpaquete, " & _
                        "PvPaquetePaciente.intmovpaciente, " & _
                        "ExPaciente.vchapellidopaterno || ' ' || ExPaciente.vchapellidomaterno || ' ' || ExPaciente.vchnombre as nombrePaciente " & _
                     "FROM PvPresupuesto " & _
                        "inner join PvPaquetePaciente on PvPresupuesto.intnumpaquete = PvPaquetePaciente.intnumpaquete " & _
                        "inner join PvPaquete on PvPaquetePaciente.intnumpaquete = PvPaquete.intnumpaquete " & _
                        "inner join ExPacienteIngreso on PvPaquetePaciente.intmovpaciente = ExPacienteIngreso.intnumcuenta " & _
                        "inner join ExPaciente on ExPacienteIngreso.intnumpaciente = ExPaciente.intnumpaciente " & _
                     "WHERE smicvedepartamento = " & vgintNumeroDepartamento & _
                      " AND (dtmfechapresupuesto between to_date(" & fstrFechaSQL(mskFechaInicio, "00:00:00") & ", 'yyyy-mm-dd hh24:mi:ss') And to_date(" & fstrFechaSQL(mskFechaFin, "23:59:59") & ", 'yyyy-mm-dd hh24:mi:ss') " & _
                            "OR intcvepresupuesto between " & vlstrFolioInicio & " and " & vlstrFolioFin & ")" & _
                      " AND chrEstado = 'A' order by intcvepresupuesto"

    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            With grdPresupuestos
                If IIf(.TextMatrix(1, ColRowDatagrdPaquete) = "", -1, .TextMatrix(1, ColRowDatagrdPaquete)) <> -1 Then
                    .Rows = .Rows + 1
                End If
                .Row = .Rows - 1
                .TextMatrix(.Row, ColRowDatagrdPaquete) = IIf(IsNull(rs!INTCVEPRESUPUESTO), 0, rs!INTCVEPRESUPUESTO)
                .TextMatrix(.Row, ColNumeroPresupuesto) = IIf(IsNull(rs!INTCVEPRESUPUESTO), 0, rs!INTCVEPRESUPUESTO)
                .TextMatrix(.Row, ColFechaPresupuesto) = IIf(IsNull(rs!fechapresupuesto), "", Format(rs!fechapresupuesto, "dd/mmm/yyyy"))
                .TextMatrix(.Row, ColNumeroPaquete) = rs!intnumpaquete
                .TextMatrix(.Row, ColDescripcionPaquete) = rs!descripcionPaquete
                .TextMatrix(.Row, ColNumCuentaPaciente) = rs!INTMOVPACIENTE
                .TextMatrix(.Row, ColNombrePaciente) = rs!NOMBREPACIENTE
            End With
            rs.MoveNext
        Loop
    Else
        MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
        'pNuevo
        Exit Sub
    End If
    rs.Close
        
    grdPresupuestos.SetFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaGridPresupuestos"))
End Sub


Private Sub pCalculaTotales()
' LLena el Grid de los totales a partir del grid de los elementos del paquete

    Dim vllngContador As Integer
    Dim vlstrTipoGrupoCargo As String
    
    Dim vldblCosto As Double
    Dim vldblImporte As Double
    Dim vldblDescuento As Double
    Dim vldblSubtotal As Double
    Dim vldblIVA As Double
    Dim vldbltotal As Double
    Dim vldblCostoConv As Double
    Dim vldblImporteConv As Double
    Dim vldblDescConv As Double
    Dim vldblSubtotalConv As Double
    Dim vldblIvaConv As Double
    Dim vldblTotalConv As Double
    
    Dim vldblCostoAR As Double
    Dim vldblImporteAR As Double
    Dim vldblDescuentoAR As Double
    Dim vldblSubtotalAR As Double
    Dim vldblIvaAR As Double
    Dim vldblTotalAR As Double
    Dim vldblCostoARConv As Double
    Dim vldblImporteARConv As Double
    Dim vldblDescARConv As Double
    Dim vldblSubtotalARConv As Double
    Dim vldblIvaARConv As Double
    Dim vldblTotalARConv As Double
    
    Dim vldblCostoME As Double
    Dim vldblImporteME As Double
    Dim vldblDescuentoME As Double
    Dim vldblSubtotalME As Double
    Dim vldblIvaME As Double
    Dim vldblTotalME As Double
    Dim vldblCostoMEConv As Double
    Dim vldblImporteMEConv As Double
    Dim vldblDescMEConv As Double
    Dim vldblSubtotalMEConv As Double
    Dim vldblIvaMEConv As Double
    Dim vldblTotalMEConv As Double
    
    Dim vldblCostoES As Double
    Dim vldblImporteES As Double
    Dim vldblDescuentoES As Double
    Dim vldblSubtotalES As Double
    Dim vldblIvaES As Double
    Dim vldblTotalES As Double
    Dim vldblCostoESConv As Double
    Dim vldblImporteESConv As Double
    Dim vldblDescESConv As Double
    Dim vldblSubtotalESConv As Double
    Dim vldblIvaESConv As Double
    Dim vldblTotalESConv As Double
    
    Dim vldblCostoEX As Double
    Dim vldblImporteEX As Double
    Dim vldblDescuentoEX As Double
    Dim vldblSubtotalEX As Double
    Dim vldblIvaEX As Double
    Dim vldblTotalEX As Double
    Dim vldblCostoEXConv As Double
    Dim vldblImporteEXConv As Double
    Dim vldblDescEXConv As Double
    Dim vldblSubtotalEXConv As Double
    Dim vldblIvaEXConv As Double
    Dim vldblTotalEXConv As Double
    
    Dim vldblCostoOC As Double
    Dim vldblImporteOC As Double
    Dim vldblDescuentoOC As Double
    Dim vldblSubtotalOC As Double
    Dim vldblIvaOC As Double
    Dim vldblTotalOC As Double
    Dim vldblCostoOCConv As Double
    Dim vldblImporteOCConv As Double
    Dim vldblDescOCConv As Double
    Dim vldblSubtotalOCConv As Double
    Dim vldblIvaOCConv As Double
    Dim vldblTotalOCConv As Double
    
    Dim rsArticulos As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    With grdPaquete
        For vllngContador = 2 To .Rows - 1
            
            vlstrTipoGrupoCargo = ""
            If .TextMatrix(vllngContador, 0) = "GC" Then
                Set rs = frsRegresaRs("SELECT chrtipo FROM PVGRUPOCARGO WHERE intcvegrupo = " & .TextMatrix(vllngContador, 15), adLockReadOnly, adOpenForwardOnly)
                vlstrTipoGrupoCargo = IIf(rs.RecordCount = 0, "", rs!chrTipo)
                rs.Close
            End If
        
            vldblCosto = vldblCosto + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
            vldblImporte = vldblImporte + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
            vldblDescuento = vldblDescuento + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
            vldblSubtotal = vldblSubtotal + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
            vldblIVA = vldblIVA + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
            vldbltotal = vldbltotal + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
            
            vldblCostoConv = vldblCostoConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
            vldblImporteConv = vldblImporteConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
            vldblDescConv = vldblDescConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
            vldblSubtotalConv = vldblSubtotalConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
            vldblIvaConv = vldblIvaConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
            vldblTotalConv = vldblTotalConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            
            If .TextMatrix(vllngContador, 0) = "AR" Then
                Set rsArticulos = frsRegresaRs("SELECT CASE WHEN chrCveArtMedicamen = 1 THEN 'ME' ELSE 'AR' END Tipo FROM IVARTICULO Where intIDArticulo = " & .TextMatrix(vllngContador, 15), adLockReadOnly, adOpenForwardOnly)
                If rsArticulos!tipo = "AR" Then
                    vldblCostoAR = vldblCostoAR + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                    vldblImporteAR = vldblImporteAR + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                    vldblDescuentoAR = vldblDescuentoAR + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                    vldblSubtotalAR = vldblSubtotalAR + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                    vldblIvaAR = vldblIvaAR + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                    vldblTotalAR = vldblTotalAR + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                    
                    vldblCostoARConv = vldblCostoARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                    vldblImporteARConv = vldblImporteARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                    vldblDescARConv = vldblDescARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                    vldblSubtotalARConv = vldblSubtotalARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                    vldblIvaARConv = vldblIvaARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                    vldblTotalARConv = vldblTotalARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
                Else
                    vldblCostoME = vldblCostoME + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                    vldblImporteME = vldblImporteME + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                    vldblDescuentoME = vldblDescuentoME + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                    vldblSubtotalME = vldblSubtotalME + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                    vldblIvaME = vldblIvaME + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                    vldblTotalME = vldblTotalME + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                    
                    vldblCostoMEConv = vldblCostoMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                    vldblImporteMEConv = vldblImporteMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                    vldblDescMEConv = vldblDescMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                    vldblSubtotalMEConv = vldblSubtotalMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                    vldblIvaMEConv = vldblIvaMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                    vldblTotalMEConv = vldblTotalMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
                End If
                rsArticulos.Close
            End If
            
            If .TextMatrix(vllngContador, 0) = "GC" And vlstrTipoGrupoCargo = "AR" Then
                vldblCostoAR = vldblCostoAR + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                vldblImporteAR = vldblImporteAR + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                vldblDescuentoAR = vldblDescuentoAR + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                vldblSubtotalAR = vldblSubtotalAR + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                vldblIvaAR = vldblIvaAR + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                vldblTotalAR = vldblTotalAR + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                    
                vldblCostoARConv = vldblCostoARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                vldblImporteARConv = vldblImporteARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                vldblDescARConv = vldblDescARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                vldblSubtotalARConv = vldblSubtotalARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                vldblIvaARConv = vldblIvaARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                vldblTotalARConv = vldblTotalARConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            End If
            
            If .TextMatrix(vllngContador, 0) = "GC" And vlstrTipoGrupoCargo = "ME" Then
                vldblCostoME = vldblCostoME + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                vldblImporteME = vldblImporteME + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                vldblDescuentoME = vldblDescuentoME + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                vldblSubtotalME = vldblSubtotalME + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                vldblIvaME = vldblIvaME + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                vldblTotalME = vldblTotalME + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                
                vldblCostoMEConv = vldblCostoMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                vldblImporteMEConv = vldblImporteMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                vldblDescMEConv = vldblDescMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                vldblSubtotalMEConv = vldblSubtotalMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                vldblIvaMEConv = vldblIvaMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                vldblTotalMEConv = vldblTotalMEConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            End If
            
            If .TextMatrix(vllngContador, 0) = "ES" Or vlstrTipoGrupoCargo = "ES" Then
                vldblCostoES = vldblCostoES + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                vldblImporteES = vldblImporteES + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                vldblDescuentoES = vldblDescuentoES + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                vldblSubtotalES = vldblSubtotalES + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                vldblIvaES = vldblIvaES + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                vldblTotalES = vldblTotalES + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                
                vldblCostoESConv = vldblCostoESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                vldblImporteESConv = vldblImporteESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                vldblDescESConv = vldblDescESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                vldblSubtotalESConv = vldblSubtotalESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                vldblIvaESConv = vldblIvaESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                vldblTotalESConv = vldblTotalESConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            End If
            
            If (.TextMatrix(vllngContador, 0) = "EX" Or .TextMatrix(vllngContador, 0) = "GE") Or vlstrTipoGrupoCargo = "EX" Then
                vldblCostoEX = vldblCostoEX + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                vldblImporteEX = vldblImporteEX + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                vldblDescuentoEX = vldblDescuentoEX + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                vldblSubtotalEX = vldblSubtotalEX + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                vldblIvaEX = vldblIvaEX + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                vldblTotalEX = vldblTotalEX + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                
                vldblCostoEXConv = vldblCostoEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                vldblImporteEXConv = vldblImporteEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                vldblDescEXConv = vldblDescEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                vldblSubtotalEXConv = vldblSubtotalEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                vldblIvaEXConv = vldblIvaEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                vldblTotalEXConv = vldblTotalEXConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            End If
            
            If .TextMatrix(vllngContador, 0) = "OC" Or vlstrTipoGrupoCargo = "OC" Then
                vldblCostoOC = vldblCostoOC + CDbl(Val(Format(.TextMatrix(vllngContador, 6), "############.##")))
                vldblImporteOC = vldblImporteOC + CDbl(Val(Format(.TextMatrix(vllngContador, 8), "############.##")))
                vldblDescuentoOC = vldblDescuentoOC + CDbl(Val(Format(.TextMatrix(vllngContador, 9), "############.##")))
                vldblSubtotalOC = vldblSubtotalOC + CDbl(Val(Format(.TextMatrix(vllngContador, 10), "############.##")))
                vldblIvaOC = vldblIvaOC + CDbl(Val(Format(.TextMatrix(vllngContador, 11), "############.##")))
                vldblTotalOC = vldblTotalOC + CDbl(Val(Format(.TextMatrix(vllngContador, 12), "############.##")))
                
                vldblCostoOCConv = vldblCostoOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 17), "############.##")))
                vldblImporteOCConv = vldblImporteOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 19), "############.##")))
                vldblDescOCConv = vldblDescOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 20), "############.##")))
                vldblSubtotalOCConv = vldblSubtotalOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 21), "############.##")))
                vldblIvaOCConv = vldblIvaOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 22), "############.##")))
                vldblTotalOCConv = vldblTotalOCConv + CDbl(Val(Format(.TextMatrix(vllngContador, 23), "############.##")))
            End If
            
        Next vllngContador
        txtSubtotal.Text = FormatCurrency(vldblSubtotal, 2)
        txtIva.Text = FormatCurrency(vldblIVA, 2)
        txtTotal.Text = FormatCurrency(vldbltotal, 2)
        
        With grdTotales
            .TextMatrix(2, 2) = FormatCurrency(vldblCostoME, 2)
            .TextMatrix(2, 3) = FormatCurrency(vldblImporteME, 2)
            .TextMatrix(2, 4) = FormatCurrency(vldblDescuentoME, 2)
            .TextMatrix(2, 5) = FormatCurrency(vldblSubtotalME, 2)
            .TextMatrix(2, 6) = FormatCurrency(vldblIvaME, 2)
            .TextMatrix(2, 7) = FormatCurrency(vldblTotalME, 2)
            .TextMatrix(2, 8) = FormatCurrency(vldblCostoMEConv, 2)
            .TextMatrix(2, 9) = FormatCurrency(vldblImporteMEConv, 2)
            .TextMatrix(2, 10) = FormatCurrency(vldblDescMEConv, 2)
            .TextMatrix(2, 11) = FormatCurrency(vldblSubtotalMEConv, 2)
            .TextMatrix(2, 12) = FormatCurrency(vldblIvaMEConv, 2)
            .TextMatrix(2, 13) = FormatCurrency(vldblTotalMEConv, 2)
            
            .TextMatrix(3, 2) = FormatCurrency(vldblCostoAR, 2)
            .TextMatrix(3, 3) = FormatCurrency(vldblImporteAR, 2)
            .TextMatrix(3, 4) = FormatCurrency(vldblDescuentoAR, 2)
            .TextMatrix(3, 5) = FormatCurrency(vldblSubtotalAR, 2)
            .TextMatrix(3, 6) = FormatCurrency(vldblIvaAR, 2)
            .TextMatrix(3, 7) = FormatCurrency(vldblTotalAR, 2)
            .TextMatrix(3, 8) = FormatCurrency(vldblCostoARConv, 2)
            .TextMatrix(3, 9) = FormatCurrency(vldblImporteARConv, 2)
            .TextMatrix(3, 10) = FormatCurrency(vldblDescARConv, 2)
            .TextMatrix(3, 11) = FormatCurrency(vldblSubtotalARConv, 2)
            .TextMatrix(3, 12) = FormatCurrency(vldblIvaARConv, 2)
            .TextMatrix(3, 13) = FormatCurrency(vldblTotalARConv, 2)
            
            .TextMatrix(4, 2) = FormatCurrency(vldblCostoES + vldblCostoEX + vldblCostoOC, 2)
            .TextMatrix(4, 3) = FormatCurrency(vldblImporteES + vldblImporteEX + vldblImporteOC, 2)
            .TextMatrix(4, 4) = FormatCurrency(vldblDescuentoES + vldblDescuentoEX + vldblDescuentoOC, 2)
            .TextMatrix(4, 5) = FormatCurrency(vldblSubtotalES + vldblSubtotalEX + vldblSubtotalOC, 2)
            .TextMatrix(4, 6) = FormatCurrency(vldblIvaES + vldblIvaEX + vldblIvaOC, 2)
            .TextMatrix(4, 7) = FormatCurrency(vldblTotalES + vldblTotalEX + vldblTotalOC, 2)
            .TextMatrix(4, 8) = FormatCurrency(vldblCostoESConv + vldblCostoEXConv + vldblCostoOCConv, 2)
            .TextMatrix(4, 9) = FormatCurrency(vldblImporteESConv + vldblImporteEXConv + vldblImporteOCConv, 2)
            .TextMatrix(4, 10) = FormatCurrency(vldblDescESConv + vldblDescEXConv + vldblDescOCConv, 2)
            .TextMatrix(4, 11) = FormatCurrency(vldblSubtotalESConv + vldblSubtotalEXConv + vldblSubtotalOCConv, 2)
            .TextMatrix(4, 12) = FormatCurrency(vldblIvaESConv + vldblIvaEXConv + vldblIvaOCConv, 2)
            .TextMatrix(4, 13) = FormatCurrency(vldblTotalESConv + vldblTotalEXConv + vldblTotalOCConv, 2)
            
            .TextMatrix(5, 2) = FormatCurrency(vldblCostoES, 2)
            .TextMatrix(5, 3) = FormatCurrency(vldblImporteES, 2)
            .TextMatrix(5, 4) = FormatCurrency(vldblDescuentoES, 2)
            .TextMatrix(5, 5) = FormatCurrency(vldblSubtotalES, 2)
            .TextMatrix(5, 6) = FormatCurrency(vldblIvaES, 2)
            .TextMatrix(5, 7) = FormatCurrency(vldblTotalES, 2)
            .TextMatrix(5, 8) = FormatCurrency(vldblCostoESConv, 2)
            .TextMatrix(5, 9) = FormatCurrency(vldblImporteESConv, 2)
            .TextMatrix(5, 10) = FormatCurrency(vldblDescESConv, 2)
            .TextMatrix(5, 11) = FormatCurrency(vldblSubtotalESConv, 2)
            .TextMatrix(5, 12) = FormatCurrency(vldblIvaESConv, 2)
            .TextMatrix(5, 13) = FormatCurrency(vldblTotalESConv, 2)
            
            .TextMatrix(6, 2) = FormatCurrency(vldblCostoEX, 2)
            .TextMatrix(6, 3) = FormatCurrency(vldblImporteEX, 2)
            .TextMatrix(6, 4) = FormatCurrency(vldblDescuentoEX, 2)
            .TextMatrix(6, 5) = FormatCurrency(vldblSubtotalEX, 2)
            .TextMatrix(6, 6) = FormatCurrency(vldblIvaEX, 2)
            .TextMatrix(6, 7) = FormatCurrency(vldblTotalEX, 2)
            .TextMatrix(6, 8) = FormatCurrency(vldblCostoEXConv, 2)
            .TextMatrix(6, 9) = FormatCurrency(vldblImporteEXConv, 2)
            .TextMatrix(6, 10) = FormatCurrency(vldblDescEXConv, 2)
            .TextMatrix(6, 11) = FormatCurrency(vldblSubtotalEXConv, 2)
            .TextMatrix(6, 12) = FormatCurrency(vldblIvaEXConv, 2)
            .TextMatrix(6, 13) = FormatCurrency(vldblTotalEXConv, 2)
            
            .TextMatrix(7, 2) = FormatCurrency(vldblCostoOC, 2)
            .TextMatrix(7, 3) = FormatCurrency(vldblImporteOC, 2)
            .TextMatrix(7, 4) = FormatCurrency(vldblDescuentoOC, 2)
            .TextMatrix(7, 5) = FormatCurrency(vldblSubtotalOC, 2)
            .TextMatrix(7, 6) = FormatCurrency(vldblIvaOC, 2)
            .TextMatrix(7, 7) = FormatCurrency(vldblTotalOC, 2)
            .TextMatrix(7, 8) = FormatCurrency(vldblCostoOCConv, 2)
            .TextMatrix(7, 9) = FormatCurrency(vldblImporteOCConv, 2)
            .TextMatrix(7, 10) = FormatCurrency(vldblDescOCConv, 2)
            .TextMatrix(7, 11) = FormatCurrency(vldblSubtotalOCConv, 2)
            .TextMatrix(7, 12) = FormatCurrency(vldblIvaOCConv, 2)
            .TextMatrix(7, 13) = FormatCurrency(vldblTotalOCConv, 2)
            
            .TextMatrix(8, 2) = FormatCurrency(vldblCosto, 2)
            .TextMatrix(8, 3) = FormatCurrency(vldblImporte, 2)
            .TextMatrix(8, 4) = FormatCurrency(vldblDescuento, 2)
            .TextMatrix(8, 5) = FormatCurrency(vldblSubtotal, 2)
            .TextMatrix(8, 6) = FormatCurrency(vldblIVA, 2)
            .TextMatrix(8, 7) = FormatCurrency(vldbltotal, 2)
            .TextMatrix(8, 8) = FormatCurrency(vldblCostoConv, 2)
            .TextMatrix(8, 9) = FormatCurrency(vldblImporteConv, 2)
            .TextMatrix(8, 10) = FormatCurrency(vldblDescConv, 2)
            .TextMatrix(8, 11) = FormatCurrency(vldblSubtotalConv, 2)
            .TextMatrix(8, 12) = FormatCurrency(vldblIvaConv, 2)
            .TextMatrix(8, 13) = FormatCurrency(vldblTotalConv, 2)
            
        End With
    End With
End Sub

Private Sub pModificaRegistro()
    Dim rsCargosSeleccionados As New ADODB.Recordset
    Dim rsDepartamentos As New ADODB.Recordset
    Dim rsHonorarios As New ADODB.Recordset
    Dim rsArticuloSKU As New ADODB.Recordset
    
    Dim vlstrSentencia As String
    Dim vlintcontador As Integer
    Dim vllngContenido As Long
    Dim vlintModoDescuentoInventario As Integer
    Dim vgstrSentencia As String
    Dim vgintCont As Integer
    Dim rsUnidad As New ADODB.Recordset
    
    
    ' Permite realizar la modificación de la descripción de un registro
    vgstrEstadoManto = "M"
    pHabilitaCampos (True)
    vlblnAgregaNuevo = False
    pCargarDepartamentos
    
    lstDepartamentosSel.Clear
    'Se cargan los departamentos asignados al paquete
    vgstrSentencia = "SELECT PVPAQUETEDEPARTAMENTO.SMICVEDEPARTAMENTO Cve, TRIM(NODEPARTAMENTO.vchDescripcion) Nombre " & _
                     "FROM PVPAQUETEDEPARTAMENTO " & _
                        "INNER JOIN NODEPARTAMENTO ON NODEPARTAMENTO.SMICVEDEPARTAMENTO = PVPAQUETEDEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                     "WHERE PVPAQUETEDEPARTAMENTO.INTNUMPAQUETE = " & rsPaquetes!intnumpaquete & _
                     " ORDER BY NOMBRE"
    Set rsDepartamentos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)

    If rsDepartamentos.RecordCount > 0 Then
        With lstDepartamentosSel
            Do While Not rsDepartamentos.EOF
                .AddItem rsDepartamentos!Nombre, .ListCount
                .ItemData(.newIndex) = rsDepartamentos!Cve
                For vgintCont = 0 To lstDepartamentos.ListCount - 1
                    If lstDepartamentos.ItemData(vgintCont) = rsDepartamentos!Cve Then
                        lstDepartamentos.RemoveItem (vgintCont)
                        Exit For
                    End If
                Next
                rsDepartamentos.MoveNext
            Loop
        End With
    End If
    rsDepartamentos.Close
    
    pHabilitaBotonesDeptos
    
    With rsPaquetes
        txtCvePaquete.Text = !intnumpaquete
        txtDescripcion.Text = Trim(!chrDescripcion)
        cmdActualizar.Enabled = IIf(chkprecio.Value, False, True)
        cboTratamiento.ListIndex = fintLocalizaCritCbo(cboTratamiento, Trim(!chrTratamiento))
        cboTipo.ListIndex = fintLocalizaCritCbo(cboTipo, Trim(!chrTipo))
        txtAnticipo.Text = FormatCurrency(str(!mnyAnticipoSugerido))
        mskFecha = IIf(IsNull(!dtmFechaActualizacion), "  /  /    ", !dtmFechaActualizacion)
        txtSeleArticulo.Text = ""
        lstElementos.Clear
        optTipoPacienteDesc(0).Value = IIf(!chrTipoIngresoDescuento = "T", True, False)
        optTipoPacienteDesc(1).Value = IIf(!chrTipoIngresoDescuento = "I", True, False)
        optTipoPacienteDesc(2).Value = IIf(!chrTipoIngresoDescuento = "E", True, False)
        optTipoPacienteDesc(3).Value = IIf(!chrTipoIngresoDescuento = "U", True, False)
        vgTipoIngresoDescuento = !chrTipoIngresoDescuento
        OptPolitica(0).Value = IIf(!bitcostobase, False, True)
        OptPolitica(1).Value = IIf(!bitcostobase, True, False)
        optFormaAsignacionTodos.Value = IIf(!bitSeleccionarHonorarios = 1, False, True)
        optFormaAsignacionSeleccionar.Value = IIf(!bitSeleccionarHonorarios = 1, True, False)
     ''  chkBasico.Value = IIf(!bitbasico Or !bitbasico = 1, 1, 0)
        '*****Fecha de alta de paquete *****************************************
        If Not !dtmFechaAltaPaquete = vbNullString Then
            txtFechaAltaPaquete.Text = Format(!dtmFechaAltaPaquete, "dd/mmm/yyyy")
            txtFechaAltaPaquete.Visible = True
            lblFechaAltaPaquete.Visible = True
        Else
            lblFechaAltaPaquete.Visible = False
            txtFechaAltaPaquete.Visible = False
        End If
        '*******************************************************************
        cboConceptoFactura.ListIndex = fintLocalizaCbo(cboConceptoFactura, !SMICONCEPTOFACTURA)
        
        chkActivo.Value = IIf(!bitactivo Or !bitactivo = 1, 1, 0)
        ChkValidaPaquete.Value = IIf(!bitValidaCargosPaquete Or !bitValidaCargosPaquete = 1, 1, 0)
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(4) = True
        pHabilitaHonorario ' SSTObj.TabEnabled(3) = tratamiento= "QUIRURGICO"
        
        grdPaquete.Visible = False
        lblTextoBarra.Caption = "Cargando datos, por favor espere..."
        freBarra.Top = 400
        freBarra.Visible = True
        pgbCargando.Value = 0
        freBarra.Refresh
        grdPaquete.Clear
        pConfiguraGridCargos

        grdTotales.Clear
        grdTotales.Rows = 3
        pConfiguraGridTotales
        
        grdGrupos.Clear 'Limpiar los datos seleccionados en el grid
        grdGrupos.Rows = 2
        pConfiguraGridGrupos
        
        grdPaquete.Rows = 3
        grdPaquete.TextMatrix(2, 15) = -1

        vlstrSentencia = "SELECT cast(PvDetallePaquete.intCveCargo as int) as CveCargo, PvDetallePaquete.smiCantidad, " & _
                "PvDetallePaquete.chrTipoCargo, " & _
                "Case " & _
                "when PvDetallePaquete.chrTipoCargo = 'AR' then IvArticulo.vchNombreComercial " & _
                "when PvDetallePaquete.chrTipoCargo = 'ES' then imEstudio.vchNombre " & _
                "when PvDetallePaquete.chrTipoCargo = 'OC' then PvOtroConcepto.chrDescripcion " & _
                "when PvDetallePaquete.chrTipoCargo = 'EX' then LaExamen.chrNombre " & _
                "when PvDetallePaquete.chrTipoCargo = 'GE' then LaGrupoExamen.chrNombre " & _
                "when PvDetallePaquete.chrTipoCargo = 'GC' then PvGrupoCargo.vchNombre " & _
                "else 'Invalido' " & _
                "end As Cargo, "
        vlstrSentencia = vlstrSentencia & _
                "Case  " & _
                "when PvDetallePaquete.chrTipoCargo = 'AR' then IvArticulo.smiCveConceptFact " & _
                "when PvDetallePaquete.chrTipoCargo = 'ES' then imEstudio.smiConFact " & _
                "when PvDetallePaquete.chrTipoCargo = 'OC' then PvOtroConcepto.smiConceptoFact " & _
                "when PvDetallePaquete.chrTipoCargo = 'EX' then LaExamen.smiConFact " & _
                "when PvDetallePaquete.chrTipoCargo = 'GE' then LaGrupoExamen.smiConFact " & _
                "when PvDetallePaquete.chrTipoCargo = 'GC' then 0" & _
                "else 0 " & _
                "end As ConceptoFacturacion, " & _
                "PvDetallePaquete.intDescuentoInventario, " & _
                "pvDetallePaquete.mnyPrecio," & _
                "pvDetallePaquete.mnyIVA," & _
                "ivArticulo.intContenido Contenido, ivArticulo.intIDArticulo, " & _
                "IvArticulo.chrCveArticulo, " & _
                "pvDetallePaquete.smiCveConcepto, " & _
                "pvDetallePaquete.mnyMontoLimite, " & _
                "pvDetallePaquete.BitControl, " & _
                "pvDetallePaquete.MnyCosto, " & _
                "pvDetallePaquete.MnyDescuento, " & _
                "pvDetallePaquete.MnyPrecioEspecifico, " & _
                "NVL(IvFamilia.vchdescripcion,0) familia, " & _
                "NVL(IvSubFamilia.vchdescripcion,0) subfamilia " & _
            "FROM PvDetallePaquete "
        vlstrSentencia = vlstrSentencia & _
            " LEFT OUTER JOIN PvOtroConcepto ON PvDetallePaquete.intCveCargo = PvOtroConcepto.intCveConcepto " & _
            " LEFT OUTER JOIN LaExamen ON PvDetallePaquete.intCveCargo = LaExamen.IntCveExamen " & _
            " LEFT OUTER JOIN LaGrupoExamen ON PvDetallePaquete.intCveCargo = LaGrupoExamen.IntCveGrupo " & _
            " LEFT OUTER JOIN ImEstudio ON PvDetallePaquete.intCveCargo = ImEstudio.intCveEstudio " & _
            " LEFT OUTER JOIN PvGrupoCargo ON PvDetallePaquete.intCveCargo = PvGrupoCargo.intCveGrupo " & _
            " LEFT OUTER JOIN IvArticulo ON PvDetallePaquete.intCveCargo = IvArticulo.intIdArticulo " & _
                " LEFT OUTER JOIN ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                " LEFT OUTER JOIN ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
            " LEFT OUTER JOIN IvFamilia ON IvArticulo.CHRCVEFAMILIA = IvFamilia.CHRCVEFAMILIA and IvArticulo.CHRCVEARTMEDICAMEN = IvFamilia.CHRCVEARTMEDICAMEN" & _
            " LEFT OUTER JOIN IvSubFamilia  ON IvArticulo.CHRCVEFAMILIA = IvSubFamilia.CHRCVEFAMILIA and  IvArticulo.CHRCVESUBFAMILIA = IvSubFamilia.CHRCVESUBFAMILIA and IvSubFamilia.CHRCVEARTMEDICAMEN = ivarticulo.CHRCVEARTMEDICAMEN" & _
            " WHERE PvDetallePaquete.intNumPaquete = " & CStr(!intnumpaquete) & " order by cargo asc"

        rsCargosSeleccionados.CursorType = adOpenForwardOnly
        rsCargosSeleccionados.LockType = adLockReadOnly
        rsCargosSeleccionados.ActiveConnection = EntornoSIHO.ConeccionSIHO
        rsCargosSeleccionados.Source = vlstrSentencia
        rsCargosSeleccionados.Open

        Do While Not rsCargosSeleccionados.EOF
            pgbCargando.Value = rsCargosSeleccionados.Bookmark / rsCargosSeleccionados.RecordCount * 100
            With grdPaquete
                vllngContenido = 0
                vlintModoDescuentoInventario = 0
                If rsCargosSeleccionados!chrTipoCargo = "AR" Then 'Nomas para los articulos
                    vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
                    vllngContenido = rsCargosSeleccionados!Contenido 'Este es el contenido de IVarticulo
                    If vllngContenido > 1 Then
                        vlintModoDescuentoInventario = rsCargosSeleccionados!INTDESCUENTOINVENTARIO 'Descuento por unidad Minima
                    End If
                End If
            If rsCargosSeleccionados!chrTipoCargo = "AR" Then
                vlstrSentencia = "Select intContenido Contenido, " & _
                                    " substring(vchNombreComercial,1,50) Articulo, " & _
                                    " ivUA.vchDescripcion UnidadAlterna, " & _
                                    " ivUM.vchDescripcion UnidadMinima, " & _
                                    " ivArticulo.chrCveArticulo " & _
                                        " From ivArticulo " & _
                                    " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                    " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                    " WHERE intIDArticulo = " & rsCargosSeleccionados!intIdArticulo
                Set rsUnidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                'SKU delarticulo
                vlstrSentencia = "select VCHCVEEXTERNA from SIEQUIVALENCIADETALLE where INTCVEEQUIVALENCIA = 15 and VCHCVELOCAL = " & rsCargosSeleccionados!intIdArticulo
                Set rsArticuloSKU = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
             End If
                If grdPaquete.TextMatrix(2, 15) <> -1 Then
                    .Rows = .Rows + 1
                End If
                
                .Row = .Rows - 1
                .TextMatrix(.Row, 15) = IIf(rsCargosSeleccionados.Fields("CveCargo").Value < 0, rsCargosSeleccionados!CveCargo * -1, rsCargosSeleccionados!CveCargo)
                .TextMatrix(.Row, 0) = rsCargosSeleccionados!chrTipoCargo
                .TextMatrix(.Row, 1) = IIf(IsNull(rsCargosSeleccionados!Cargo), "", rsCargosSeleccionados!Cargo)
                .TextMatrix(.Row, 2) = rsCargosSeleccionados!SMICANTIDAD
                .TextMatrix(.Row, 30) = rsCargosSeleccionados!SMICANTIDAD
                
                If vlintModoDescuentoInventario = 1 And rsCargosSeleccionados!chrTipoCargo = "AR" Then
                    .TextMatrix(.Row, 3) = rsUnidad!UnidadMinima
                ElseIf vlintModoDescuentoInventario = 2 And rsCargosSeleccionados!chrTipoCargo = "AR" Then
                    .TextMatrix(.Row, 3) = rsUnidad!UnidadAlterna
                Else
                    .TextMatrix(.Row, 3) = ""
                End If
                
                If rsCargosSeleccionados!chrTipoCargo = "AR" Then
                    .TextMatrix(.Row, 31) = IIf(IsNull(rsCargosSeleccionados!chrcvearticulo), " ", rsCargosSeleccionados!chrcvearticulo)
                    If rsArticuloSKU.RecordCount > 0 Then
                        .TextMatrix(.Row, 25) = IIf(IsNull(rsArticuloSKU!VCHCVEEXTERNA), "", rsArticuloSKU!VCHCVEEXTERNA)
                    Else
                        .TextMatrix(.Row, 25) = ""
                    End If
                Else
                    .TextMatrix(.Row, 31) = ""
                End If
                
                .TextMatrix(.Row, 4) = IIf(rsCargosSeleccionados!chrTipoCargo = "GC", FormatCurrency(rsCargosSeleccionados!mnyMontoLimite, 2), "")
                .TextMatrix(.Row, 5) = FormatCurrency(rsCargosSeleccionados!mnycosto, 2)
                .TextMatrix(.Row, 6) = FormatCurrency(Val(Format(.TextMatrix(.Row, 5), "")) * Val(Format(.TextMatrix(.Row, 2), "")), 2)
                .TextMatrix(.Row, 7) = FormatCurrency(rsCargosSeleccionados!mnyPrecio, 2)
                .TextMatrix(.Row, 8) = FormatCurrency(rsCargosSeleccionados!mnyPrecio * rsCargosSeleccionados!SMICANTIDAD, 2)
                .TextMatrix(.Row, 9) = FormatCurrency(rsCargosSeleccionados!MNYDESCUENTO, 2)
                .TextMatrix(.Row, 10) = FormatCurrency((rsCargosSeleccionados!SMICANTIDAD * rsCargosSeleccionados!mnyPrecio) - rsCargosSeleccionados!MNYDESCUENTO, 2)
                .TextMatrix(.Row, 11) = FormatCurrency(rsCargosSeleccionados!MNYIVA, 2)
                .TextMatrix(.Row, 12) = FormatCurrency(((rsCargosSeleccionados!SMICANTIDAD * rsCargosSeleccionados!mnyPrecio) - rsCargosSeleccionados!MNYDESCUENTO) + rsCargosSeleccionados!MNYIVA, 2)
                .TextMatrix(.Row, 13) = vlintModoDescuentoInventario
                .TextMatrix(.Row, 14) = vllngContenido
                .TextMatrix(.Row, 16) = FormatCurrency(.TextMatrix(.Row, 5), 2)
                .TextMatrix(.Row, 17) = FormatCurrency(.TextMatrix(.Row, 6), 2)
                .TextMatrix(.Row, 24) = rsCargosSeleccionados!MNYPRECIOESPECIFICO
                .TextMatrix(.Row, 26) = IIf(IsNull(rsCargosSeleccionados!smicveconcepto), "", IIf(rsCargosSeleccionados!chrTipoCargo = "GC", rsCargosSeleccionados!smicveconcepto, ""))
                .TextMatrix(.Row, 27) = IIf(rsCargosSeleccionados!familia = 0, "", rsCargosSeleccionados!familia)
                .TextMatrix(.Row, 28) = IIf(rsCargosSeleccionados!subfamilia = 0, "", rsCargosSeleccionados!subfamilia)
                .TextMatrix(.Row, 29) = IIf(rsCargosSeleccionados!chrTipoCargo = "GC", FormatCurrency(rsCargosSeleccionados!mnyMontoLimite, 2), "")
                .Col = 0
                
                If rsCargosSeleccionados!chrTipoCargo = "GC" Then
                    pAgregaGrupoConcepto rsCargosSeleccionados!CveCargo, IIf(IsNull(rsCargosSeleccionados!Cargo), "", rsCargosSeleccionados!Cargo), rsCargosSeleccionados!smicveconcepto
                End If
                
                rsCargosSeleccionados.MoveNext
            End With
        Loop
        If grdPaquete.TextMatrix(2, 1) <> "" Then
            cmdExportar.Enabled = True
            cmdImportar.Enabled = True
        Else
            cmdExportar.Enabled = False
            cmdImportar.Enabled = True
        End If
        chkprecio.Value = IIf(IsNull(!bitincrementoautomatico), 0, !bitincrementoautomatico)
        rsCargosSeleccionados.Close
        pCalculaTotales
        freBarra.Visible = False
        grdPaquete.Visible = True
        
        
        
    End With
    
    'Cargar los honorarios asignados al paquete
    vlstrSentencia = ""
    'Se cargan los departamentos asignados al paquete
     vgstrSentencia = "SELECT PVPAQUETEHONORARIOS.INTCVEPAQUETE " & _
                        ",PVPAQUETEHONORARIOS.INTCVEFUNCION, TRIM(EXFUNCIONPARTICIPANTECIRUGIA.VCHDESCRIPCION) NOMBREFUNCION " & _
                        " ,PVPAQUETEHONORARIOS.INTCVECONCEPTO, TRIM(PVOTROCONCEPTO.CHRDESCRIPCION) DESCRIPCIONCONCEPTO" & _
                        " , PVPAQUETEHONORARIOS.MNYIMPORTEHONORARIO" & _
                        " From PVPAQUETEHONORARIOS " & _
                        " INNER JOIN EXFUNCIONPARTICIPANTECIRUGIA ON EXFUNCIONPARTICIPANTECIRUGIA.INTCVEFUNCION = PVPAQUETEHONORARIOS.INTCVEFUNCION " & _
                        " INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = PVPAQUETEHONORARIOS.INTCVECONCEPTO " & _
                     " WHERE PVPAQUETEHONORARIOS.INTCVEPAQUETE = " & rsPaquetes!intnumpaquete & _
                     " ORDER BY NOMBREFUNCION"
    Set rsHonorarios = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)

    With grdDetalleHonorarios
        .Rows = 1
        If rsHonorarios.RecordCount <> 0 Then
            Do While Not rsHonorarios.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, cintColDescripcionFuncion) = Trim(rsHonorarios!NOMBREFUNCION)
                .TextMatrix(.Rows - 1, cintColClaveFuncion) = Trim(rsHonorarios!INTCVEFUNCION)
                .TextMatrix(.Rows - 1, cintColDescripcionConcepto) = Trim(rsHonorarios!DescripcionConcepto)
                .TextMatrix(.Rows - 1, cintColClaveConcepto) = Trim(rsHonorarios!intCveConcepto)
                .TextMatrix(.Rows - 1, cintColImporteHonorario) = FormatCurrency(rsHonorarios!MNYIMPORTEHONORARIO, 2)
                rsHonorarios.MoveNext
            Loop
        End If
    End With
    rsHonorarios.Close
    
End Sub

Private Sub pSeleccionaElemento()
    Dim vlstrCualLista As String
    Dim vllngPosicion As Long
    Dim vlstrSentencia As String
    Dim vldblPrecio As Double
    Dim vldblDescuento As Double
    Dim vldblDescuentoConvenio As Double
    Dim vldblPorceDescuento As Double
    Dim vldblPorceDescuentoConv As Double
    Dim vldblPrecioConvenio As Double
    Dim vldblIVA As Integer
    Dim vldblIvaConv As Integer
    Dim vllngContenido As Long
    Dim vlstrCveArticulo As String
    Dim vlintModoDescuentoInventario As Integer
    Dim vlstrx As String, vlstrY As String
    
    Dim vldblSubtotal As Double
    Dim vldblSubtotalConvenio As Double
    Dim vldblCantidad As Double
    Dim vldblCostoPred As Double
    Dim vldblCostoConv As Double
    Dim rs As New ADODB.Recordset
    
    Dim vlintcontador As Integer
    Dim rsEnComun As New ADODB.Recordset
    Dim vlStrGrupos As String
    Dim vlstrCadenaMsj As String
    Dim vlstrPaqueteMsj As String
    Dim vlstrSentenciaConcepto As String
    Dim vlstrConceptoFacturacion As String
    
    Dim lstListas As ListBox
    Dim vlblnNuevoElemento As Boolean
    
    Dim vlaryParametrosSalida() As String
    Dim rsUnidad As New ADODB.Recordset
    
    Set lstListas = lstElementos
    With grdPaquete
        Select Case sstElementos.Tab
            Case 0
                vlstrCualLista = "AR"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & Trim(str(lstListas.ItemData(lstListas.ListIndex))) & ")"
            Case 1
                vlstrCualLista = "ES"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & Trim(str(lstListas.ItemData(lstListas.ListIndex))) & ")"
            Case 2
                vlstrCualLista = IIf(lstElementos.ItemData(lstElementos.ListIndex) < 0, "GE", "EX")
                If lstElementos.ItemData(lstElementos.ListIndex) < 0 Then
                    vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & Trim(str(lstListas.ItemData(lstListas.ListIndex) * -1)) & ")"
                Else
                    vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & Trim(str(lstListas.ItemData(lstListas.ListIndex))) & ")"
                End If
            Case 3
                vlstrCualLista = "OC"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & Trim(str(lstListas.ItemData(lstListas.ListIndex))) & ")"
            Case 4
                vlstrCualLista = "GC"
                pElementoGrupoPred (lstListas.ItemData(lstListas.ListIndex))
                
                Select Case vgstrTipoPredGrupo
                    Case "ME"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiCveConceptFact smiConFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo
                    Case "AR"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiCveConceptFact smiConFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo
                    Case "ES"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClavePredGrupo
                    Case "EX"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClavePredGrupo
                    Case "GE"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClavePredGrupo
                    Case "OC"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConceptoFact smiConFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClavePredGrupo
                End Select
        End Select

        If vlstrSentencia <> "" Then
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            vldblIVA = IIf(rs.RecordCount = 0, 0, rs!IVA)
            rs.Close
                        
        Else
            vldblIVA = 0
        End If
        If vlstrSentenciaConcepto <> "" Then
            Set rs = frsRegresaRs(vlstrSentenciaConcepto, adLockReadOnly, adOpenForwardOnly)
            vlstrConceptoFacturacion = IIf(rs.RecordCount = 0, 0, rs!smiConFact)
            rs.Close
                        
        Else
            vlstrConceptoFacturacion = 0
        End If

        
        vlstrSentencia = ""
        Select Case sstElementos.Tab
            Case 4
                pElementoGrupoConv (lstListas.ItemData(lstListas.ListIndex))
                
                Select Case vgstrTipoConvGrupo
                    Case "ME"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                    Case "AR"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                    Case "ES"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClaveConvGrupo & ")"
                    Case "EX"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClaveConvGrupo & ")"
                    Case "GE"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClaveConvGrupo & ")"
                    Case "OC"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClaveConvGrupo & ")"
                End Select
                
                If vlstrSentencia <> "" Then
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    vldblIvaConv = IIf(rs.RecordCount = 0, 0, rs!IVA)
                    rs.Close
                Else
                    vldblIvaConv = 0
                End If
            Case Else
                vldblIvaConv = vldblIVA
        End Select

        If vlstrCualLista = "GC" Then
            vlStrGrupos = ""
            With grdGrupos
                For vlintcontador = 1 To .Rows - 1
                    If Trim(str(lstListas.ItemData(lstListas.ListIndex))) <> .RowData(vlintcontador) Then
                        vlStrGrupos = IIf(vlintcontador = 1, .RowData(vlintcontador), vlStrGrupos & IIf(Trim(vlStrGrupos) = "", "", ",") & .RowData(vlintcontador))
                    End If
                Next
                If vlStrGrupos <> "" Then
                    Set rsEnComun = frsRegresaRs("SELECT DGC.intCveGrupo CveGrupo, TRIM(GC.vchNombre) DescGrupo, DGC.chrTipoCargo TipoCargo, DGC.intCveCargo CveCargo " & _
                                                   ",DECODE(DGC.chrTipoCargo " & _
                                                       ",'AR',TRIM(AR.vchnombrecomercial), 'ME',TRIM(ME.vchnombrecomercial) " & _
                                                       ",'OC',TRIM(OC.chrdescripcion), 'ES',TRIM(ES.vchnombre) " & _
                                                       ",'EX',TRIM(EX.chrnombre), 'GE',TRIM(GE.chrnombre),'') DescCargo " & _
                                                   ",DECODE(DGC.chrTipoCargo " & _
                                                       ",'AR',TRIM(AR.chrCveArticulo), 'ME',TRIM(ME.chrCveArticulo) " & _
                                                       ",'OC',TRIM(OC.intCveConcepto), 'ES',TRIM(ES.intCveEstudio) " & _
                                                       ",'EX',TRIM(EX.intCveExamen), 'GE',TRIM(GE.intcvegrupo),'') CveCargoReal " & _
                                                "FROM PVDETALLEGRUPOCARGO DGC " & _
                                                   "LEFT JOIN PVGRUPOCARGO GC ON GC.intCveGrupo = DGC.intCveGrupo " & _
                                                   "LEFT OUTER JOIN IVARTICULO AR ON DGC.intCveCargo = AR.intIDArticulo AND AR.chrCveArtMedicamen <> 1 " & _
                                                   "LEFT OUTER JOIN IVARTICULO ME ON DGC.intCveCargo = ME.intIDArticulo AND ME.chrCveArtMedicamen = 1 " & _
                                                   "LEFT OUTER JOIN PVOTROCONCEPTO OC ON DGC.intCveCargo = OC.intCveConcepto " & _
                                                   "LEFT OUTER JOIN IMESTUDIO ES ON DGC.intCveCargo = ES.intCveEstudio " & _
                                                   "LEFT OUTER JOIN LAEXAMEN EX ON DGC.intCveCargo = EX.intCveExamen " & _
                                                   "LEFT OUTER JOIN LAGRUPOEXAMEN GE ON DGC.intCveCargo = GE.intcvegrupo " & _
                                                "WHERE DGC.intCveGrupo IN (" & vlStrGrupos & ") " & _
                                                   "AND (DGC.chrTipoCargo, DGC.intCveCargo) IN (SELECT chrTipoCargo, intCveCargo " & _
                                                                                               "FROM PVDETALLEGRUPOCARGO " & _
                                                                                               "WHERE intCveGrupo = " & Trim(str(lstListas.ItemData(lstListas.ListIndex))) & ") " & _
                                                "ORDER BY CveGrupo, DescGrupo, CveCargoReal, DescCargo", adLockReadOnly, adOpenForwardOnly)
                    With rsEnComun
                        If .RecordCount > 0 Then
                            vlstrCadenaMsj = ""
                            vlstrPaqueteMsj = ""
                            .MoveFirst
                            For vlintcontador = 1 To .RecordCount
                                If vlstrPaqueteMsj = !DescGrupo Then
                                    vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & "     " & !cveCargoReal & " " & !DescCargo
                                Else
                                    vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & Format(!cveGrupo, "########") & " " & !DescGrupo & Chr(13) & "     " & !cveCargoReal & " " & !DescCargo
                                End If
                                vlstrPaqueteMsj = !DescGrupo
                                .MoveNext
                            Next vlintcontador
                            MsgBox SIHOMsg(1101) & vlstrCadenaMsj, vbOKOnly + vbInformation, "Mensaje"
                            Exit Sub
                        End If
                    End With
                    rsEnComun.Close
                End If
            End With
        End If
        
        If lstListas.ListIndex > -1 Then
            .Redraw = False
            vllngPosicion = FintBuscaEnRowData(grdPaquete, lstListas.ItemData(lstListas.ListIndex), vlstrCualLista)
            If vllngPosicion = -1 Then        'Cuando no esta en la lista
                If IIf(.TextMatrix(2, 15) = "", -1, .TextMatrix(2, 15)) <> -1 Then
                    '.AddItem ("Nada")
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                Else
                    .Row = 2
                End If
                vlblnNuevoElemento = True
            Else
                .Row = vllngPosicion
                vlblnNuevoElemento = False
                MsgBox SIHOMsg(1632) & vlstrCadenaMsj, vbOKOnly + vbInformation, "Mensaje"
                .Redraw = True
                .Refresh
                Exit Sub
            End If
            
            vlstrCveArticulo = ""
            vlintModoDescuentoInventario = 0
            vllngContenido = 1
            
            If vlstrCualLista = "AR" Or (vlstrCualLista = "GC" And (vgstrTipoPredGrupo = "AR" Or vgstrTipoPredGrupo = "ME")) Then 'Nomas para los articulos
                ' Tipo de descuento de Inventario
                vlstrSentencia = "Select intContenido Contenido, " & _
                                " substring(vchNombreComercial,1,50) Articulo, " & _
                                " ivUA.vchDescripcion UnidadAlterna, " & _
                                " ivUM.vchDescripcion UnidadMinima, " & _
                                " ivArticulo.chrCveArticulo " & _
                                    " From ivArticulo " & _
                                " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", Trim(str(lstListas.ItemData(lstListas.ListIndex))), vglngClavePredGrupo)
                
                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
                vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                If vllngContenido > 1 Then
                    If MsgBox("¿Desea realizar la venta de " & Trim(rs!Articulo) & " por " & Trim(rs!UnidadAlterna) & "?" & Chr(13) & "Si selecciona NO, se venderá por " & Trim(rs!UnidadMinima) & ".", vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                        vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
                    End If
                End If
                rs.Close
            End If
            
            '-----------------------
            'Precio unitario PREDETERMINADAS
            '-----------------------
            
            vldblPrecio = FormatCurrency(IIf(.TextMatrix(.Row, 7) = "", 0, .TextMatrix(.Row, 7)), 2)
            If vlblnNuevoElemento Or chkprecio.Value Then
                If vlstrCualLista <> "GC" Then
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = str(IIf(lstListas.ItemData(lstListas.ListIndex) < 0, lstListas.ItemData(lstListas.ListIndex) * -1, lstListas.ItemData(lstListas.ListIndex))) & _
                    "|" & vlstrCualLista & _
                    "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & _
                    "|" & "0" & "|" & IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    '"|" & CStr(vglngTipoParticular) & _
                    '"|" & "0" & "|E|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio
                    
                    If vldblPrecio = -1 Then
                       vldblPrecio = 0
                    Else
                        'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                        If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                            vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                        End If
                    End If
                Else
                    vldblPrecio = vgdblPrecioPredGrupo
                    If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                        vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                    End If
                End If
            End If
                
            '-----------------------
            'Precio unitario CONVENIO
            '-----------------------
            If vlstrCualLista <> "GC" Then
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                vgstrParametrosSP = str(IIf(lstListas.ItemData(lstListas.ListIndex) < 0, lstListas.ItemData(lstListas.ListIndex) * -1, lstListas.ItemData(lstListas.ListIndex))) & _
                "|" & vlstrCualLista & _
                "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & _
                "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|" & IIf(cboEmpresas.ItemData(cboEmpresas.ListIndex) <> 0, IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")), "E") & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                pObtieneValores vlaryResultados, vldblPrecioConvenio
                
                If vldblPrecioConvenio = -1 Then
                   vldblPrecioConvenio = 0
                Else
                    'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                    If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                        vldblPrecioConvenio = vldblPrecioConvenio / CDbl(vllngContenido)
                    End If
                End If
            Else
                vldblPrecioConvenio = vgdblPrecioConvGrupo
                If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                    vldblPrecioConvenio = vldblPrecioConvenio / CDbl(vllngContenido)
                End If
            End If
            
            If vlblnNuevoElemento Or .TextMatrix(.Row, 9) = "$0.00" Or chkprecio.Value Then
                vldblPorceDescuento = 0
            Else
                vldblPorceDescuento = FormatCurrency(.TextMatrix(.Row, 9), 2) / (Val(.TextMatrix(.Row, 2)) * FormatCurrency(.TextMatrix(.Row, 7), 2))
            End If
            
            'If vlblnNuevoElemento Or .TextMatrix(.Row, 20) = "$0.00" Or chkprecio.Value Then
            vldblPorceDescuentoConv = 0
            If Val(Format(.TextMatrix(.Row, 2))) * Val(Format(.TextMatrix(.Row, 18))) <> 0 Then
            '    vldblPorceDescuentoConv = FormatCurrency(.TextMatrix(.Row, 20), 2) / (Val(.TextMatrix(.Row, 2)) * FormatCurrency(.TextMatrix(.Row, 18), 2))
                vldblPorceDescuentoConv = Val(Format(.TextMatrix(.Row, 20))) / Val(Format(.TextMatrix(.Row, 2))) * Val(Format(.TextMatrix(.Row, 18)))
            End If
            
            vldblCantidad = Val(.TextMatrix(.Row, 2)) + 1
            
            '-----------------------
            'DESCUENTOS PREDETERMINADAS
            '-----------------------
                vldblDescuento = 0
                If vlblnNuevoElemento Or chkprecio.Value Then
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    frsEjecuta_SP IIf(optTipoPacienteDesc(0).Value, "A", IIf(optTipoPacienteDesc(1).Value, "I", IIf(optTipoPacienteDesc(2).Value, "E", "U"))) & "|" & _
                                    IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|0|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & _
                                    IIf(vlstrCualLista <> "GC", str(IIf(lstListas.ItemData(lstListas.ListIndex) < 0, lstListas.ItemData(lstListas.ListIndex) * -1, lstListas.ItemData(lstListas.ListIndex))), vglngClavePredGrupo) & "|" & _
                                    IIf(vlintModoDescuentoInventario <> 1, vldblPrecio, vldblPrecio * CDbl(vllngContenido)) & "|" & _
                                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                    0 & "|" & CDbl(vllngContenido) & "|" & vldblCantidad & "|" & _
                                    vlintModoDescuentoInventario, _
                                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                                    
                                    'CStr(vglngTipoParticular) & "|0|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & _

                    pObtieneValores vlaryParametrosSalida, vldblDescuento
                End If
                
            '-----------------------
            'DESCUENTOS CONVENIO
            '-----------------------
                vldblDescuentoConvenio = 0
                If vlblnNuevoElemento Or chkprecio.Value Then
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    frsEjecuta_SP IIf(optTipoPaciente(6).Value, "A", IIf(optTipoPaciente(5).Value, "I", IIf(optTipoPaciente(4).Value, "E", "U"))) & "|" & _
                                    IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoConvGrupo) & "|" & _
                                    IIf(vlstrCualLista <> "GC", str(IIf(lstListas.ItemData(lstListas.ListIndex) < 0, lstListas.ItemData(lstListas.ListIndex) * -1, lstListas.ItemData(lstListas.ListIndex))), vglngClaveConvGrupo) & "|" & _
                                    IIf(vlintModoDescuentoInventario <> 1, vldblPrecioConvenio, vldblPrecioConvenio * CDbl(vllngContenido)) & "|" & _
                                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                    0 & "|" & CDbl(vllngContenido) & "|" & vldblCantidad & "|" & _
                                    vlintModoDescuentoInventario, _
                                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblDescuentoConvenio
                End If

            vldblSubtotal = (vldblCantidad * vldblPrecio) - IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuento, (vldblPrecio * vldblCantidad) * vldblPorceDescuento)
            vldblSubtotalConvenio = (vldblCantidad * vldblPrecioConvenio) - IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuentoConvenio, (vldblPrecioConvenio * vldblCantidad) * vldblPorceDescuentoConv)
                            
            vlstrSentencia = "Select intContenido Contenido, " & _
                                " substring(vchNombreComercial,1,50) Articulo, " & _
                                " ivUA.vchDescripcion UnidadAlterna, " & _
                                " ivUM.vchDescripcion UnidadMinima, " & _
                                " ivArticulo.chrCveArticulo,IvFamilia.VCHDESCRIPCION familia, IvSubFamilia.VCHDESCRIPCION subfamilia " & _
                                    " From ivArticulo " & _
                                " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                " inner Join IvFamilia ON IvArticulo.CHRCVEFAMILIA = IvFamilia.CHRCVEFAMILIA and IvArticulo.CHRCVEARTMEDICAMEN = IvFamilia.CHRCVEARTMEDICAMEN" & _
                                " inner Join IvSubFamilia  ON IvArticulo.CHRCVEFAMILIA = IvSubFamilia.CHRCVEFAMILIA and  IvArticulo.CHRCVESUBFAMILIA = IvSubFamilia.CHRCVESUBFAMILIA and IvSubFamilia.CHRCVEARTMEDICAMEN = ivarticulo.CHRCVEARTMEDICAMEN " & _
                                " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", Trim(str(lstListas.ItemData(lstListas.ListIndex))), vglngClavePredGrupo)
                
            Set rsUnidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                            
            If vlstrCualLista = "GE" Then
                .TextMatrix(.Row, 1) = Mid(lstListas.List(lstListas.ListIndex), 1, Len(lstListas.List(lstListas.ListIndex)) - 8)
            Else
                .TextMatrix(.Row, 1) = lstListas.List(lstListas.ListIndex)
            End If
            .TextMatrix(.Row, 0) = vlstrCualLista
            .TextMatrix(.Row, 2) = vldblCantidad
            .TextMatrix(.Row, 30) = vldblCantidad
            
            If vlintModoDescuentoInventario = 1 And vlstrCualLista = "AR" Then
                .TextMatrix(.Row, 3) = rsUnidad!UnidadMinima
            ElseIf vlintModoDescuentoInventario = 2 And vlstrCualLista = "AR" Then
                .TextMatrix(.Row, 3) = rsUnidad!UnidadAlterna
            Else
                .TextMatrix(.Row, 3) = ""
            End If
            
            .TextMatrix(.Row, 26) = IIf(vlstrConceptoFacturacion = 0, "", vlstrConceptoFacturacion)
            If vlstrCualLista = "AR" Then
                .TextMatrix(.Row, 27) = rsUnidad!familia
                .TextMatrix(.Row, 28) = rsUnidad!subfamilia
                .TextMatrix(.Row, 31) = IIf(IsNull(rsUnidad!chrcvearticulo), " ", rsUnidad!chrcvearticulo)
            Else
                .TextMatrix(.Row, 27) = ""
                .TextMatrix(.Row, 28) = ""
                .TextMatrix(.Row, 31) = ""
            End If

            rsUnidad.Close
            
            'SKU delarticulo
            If vlstrCualLista = "AR" Then
                vlstrSentencia = "select VCHCVEEXTERNA from SIEQUIVALENCIADETALLE where INTCVEEQUIVALENCIA = 15 and VCHCVELOCAL = " & Trim(str(lstListas.ItemData(lstListas.ListIndex)))
                Set rsUnidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsUnidad.RecordCount > 0 Then
                    .TextMatrix(.Row, 25) = rsUnidad!VCHCVEEXTERNA
                Else
                    .TextMatrix(.Row, 25) = ""
                End If
            End If
            
            .TextMatrix(.Row, 4) = IIf(vlstrCualLista <> "GC", "", IIf(vlblnNuevoElemento, FormatCurrency(0, 2), FormatCurrency(IIf(.TextMatrix(.Row, 4) = "", 0, .TextMatrix(.Row, 4)), 2)))
            .TextMatrix(.Row, 7) = FormatCurrency(vldblPrecio, 2)
            .TextMatrix(.Row, 8) = FormatCurrency(vldblPrecio * vldblCantidad, 2)
            .TextMatrix(.Row, 9) = FormatCurrency(IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuento, (vldblPrecio * vldblCantidad) * vldblPorceDescuento), 2)
            .TextMatrix(.Row, 10) = FormatCurrency(vldblSubtotal, 2)
            .TextMatrix(.Row, 11) = FormatCurrency(str(vldblSubtotal * vldblIVA / 100), 2)
            .TextMatrix(.Row, 12) = FormatCurrency(str(vldblSubtotal + (vldblSubtotal * vldblIVA / 100)), 2)
            .TextMatrix(.Row, 13) = vlintModoDescuentoInventario
            .TextMatrix(.Row, 14) = vllngContenido
            .TextMatrix(.Row, 15) = IIf(lstListas.ItemData(lstListas.ListIndex) < 0, lstListas.ItemData(lstListas.ListIndex) * -1, lstListas.ItemData(lstListas.ListIndex))
            .TextMatrix(.Row, 29) = IIf(vlstrCualLista <> "GC", "", IIf(vlblnNuevoElemento, FormatCurrency(0, 2), FormatCurrency(IIf(.TextMatrix(.Row, 4) = "", 0, .TextMatrix(.Row, 4)), 2)))
            If vlblnNuevoElemento Or chkprecio.Value Then
                vldblCostoPred = flngCosto(IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 0), vgstrTipoPredGrupo), IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 15), vglngClavePredGrupo))
                .TextMatrix(.Row, 5) = FormatCurrency(IIf(vlintModoDescuentoInventario <> 1, vldblCostoPred, vldblCostoPred / CDbl(vllngContenido)), 2)
                If vlstrCualLista <> "GC" Then
                    .TextMatrix(.Row, 16) = .TextMatrix(.Row, 5)
                Else
                    vldblCostoConv = flngCosto(IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 0), vgstrTipoConvGrupo), IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 15), vglngClaveConvGrupo))
                    .TextMatrix(.Row, 16) = FormatCurrency(IIf(vlintModoDescuentoInventario <> 1, vldblCostoConv, vldblCostoConv / CDbl(vllngContenido)), 2)
                End If
            End If
            .TextMatrix(.Row, 6) = FormatCurrency(.TextMatrix(.Row, 5) * CInt(.TextMatrix(.Row, 2)), 2)
            .TextMatrix(.Row, 17) = FormatCurrency(.TextMatrix(.Row, 16) * CInt(.TextMatrix(.Row, 2)), 2)
            
            .TextMatrix(.Row, 18) = FormatCurrency(vldblPrecioConvenio, 2)
            .TextMatrix(.Row, 19) = FormatCurrency(vldblPrecioConvenio * vldblCantidad, 2)
            .TextMatrix(.Row, 20) = FormatCurrency(IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuentoConvenio, (vldblPrecioConvenio * vldblCantidad) * vldblPorceDescuentoConv), 2)
            .TextMatrix(.Row, 21) = FormatCurrency(vldblSubtotalConvenio, 2)
            .TextMatrix(.Row, 22) = FormatCurrency(str(vldblSubtotalConvenio * vldblIvaConv / 100), 2)
            .TextMatrix(.Row, 23) = FormatCurrency(str(vldblSubtotalConvenio + (vldblSubtotalConvenio * vldblIvaConv / 100)), 2)
            .TextMatrix(.Row, 24) = 0
            vlblnAgregaNuevo = False
            If fPaqueteEnPrecioPorCargos(Val(txtCvePaquete.Text), False) Then vlblnAgregaNuevo = vlblnNuevoElemento
            
            .Col = 0
            .Redraw = True
            .Refresh
            pCalculaTotales
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
            If vlstrCualLista = "GC" Then
                pAgregaGrupoConcepto lstListas.ItemData(lstListas.ListIndex), lstListas.List(lstListas.ListIndex), -1
            End If
            If vlblnNuevoElemento Then pMuestraColumnasConv
        End If
    End With
    
End Sub

Private Sub cboLista_Click()
    If Trim(txtCvePaquete) <> "" Then
        pRecalculaPrecios
    End If
End Sub

Private Sub cboConceptoFactura_GotFocus()
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False 'Búsqueda
End Sub

Private Sub cboConceptoFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
    
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboEmpresas_Click()
    If chkprecio.Value Then pRecalculaPrecios
End Sub

Private Sub cboEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If optTipoPaciente(6).Value Then
            optTipoPaciente(6).SetFocus
        Else
            If optTipoPaciente(5).Value Then
                optTipoPaciente(5).SetFocus
            Else
                If optTipoPaciente(4).Value Then
                    optTipoPaciente(4).SetFocus
                Else
                    optTipoPaciente(3).SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub cboMovConceptoFactura_Click()
    Call SendMessage(cboMovConceptoFactura.hwnd, CB_SETITEMHEIGHT, 0, ByVal 13)
    cboMovConceptoFactura_KeyDown 1, 0
End Sub

Private Sub cboMovConceptoFactura_GotFocus()
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False 'Búsqueda
End Sub

Private Sub cboMovConceptoFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    'Para verificar que tecla fue presionada en el textbox
    With grdGrupos
        Select Case KeyCode
            Case 27   'ESC
                 cboMovConceptoFactura.Visible = False
            Case 13
                If Trim(cboMovConceptoFactura.Text) <> "" Then
                    .Text = Trim(cboMovConceptoFactura.Text)
                    .TextMatrix(.Row, 2) = cboMovConceptoFactura.ItemData(cboMovConceptoFactura.ListIndex)
                    .SetFocus
                End If
            Case 1
                If Trim(cboMovConceptoFactura.Text) <> "" And .TextMatrix(.Row, 1) <> "" And cmdPrint.Enabled = False Then ' And vgblnValida Then
                'If Trim(cboMovConceptoFactura.Text) <> "" Then
                    .Text = Trim(cboMovConceptoFactura.Text)
                    .TextMatrix(.Row, 2) = cboMovConceptoFactura.ItemData(cboMovConceptoFactura.ListIndex)
                    '.SetFocus
                End If
        End Select
    End With
End Sub

Private Sub cboMovConceptoFactura_LostFocus()
    cboMovConceptoFactura.Visible = False
End Sub

Private Sub cboTipo_GotFocus()
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(4) = False
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboTipoPaciente_Click()
    Dim vlintIndex As Integer
    Dim rs As New ADODB.Recordset
    
    If cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) = 0 Then
        cboEmpresas.ListIndex = fintLocalizaCbo(cboEmpresas, 0)
        cboEmpresas.Enabled = False
        pLimpiarCantidadesConv
    Else
        Set rs = frsRegresaRs("SELECT bitutilizaconvenio " & _
                              "FROM ADTIPOPACIENTE " & _
                              "WHERE tnycvetipopaciente = " & cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), adLockReadOnly, adOpenForwardOnly)
        If rs!bitUtilizaConvenio = 0 And cboEmpresas.Enabled Then cboEmpresas.ListIndex = fintLocalizaCbo(cboEmpresas, 0)
        cboEmpresas.Enabled = IIf(rs!bitUtilizaConvenio = 1, True, False)
        rs.Close
    End If
    
    pMuestraColumnasConv
    
    If chkprecio.Value Then pRecalculaPrecios
    
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboEmpresas.Enabled Then
            cboEmpresas.SetFocus
        Else
            If optTipoPaciente(6).Value Then
                optTipoPaciente(6).SetFocus
            Else
                If optTipoPaciente(5).Value Then
                    optTipoPaciente(5).SetFocus
                Else
                    If optTipoPaciente(4).Value Then
                        optTipoPaciente(4).SetFocus
                    Else
                        optTipoPaciente(3).SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboTratamiento_Click()
    pHabilitaHonorario
End Sub

Private Sub pHabilitaHonorario()
    If Trim(cboTratamiento.List(cboTratamiento.ListIndex)) = "QUIRURGICO" Then
        SSTObj.TabEnabled(3) = True
    Else
        SSTObj.TabEnabled(3) = False
    End If
End Sub

Private Sub cboTratamiento_GotFocus()
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(4) = False
End Sub

Private Sub cboTratamiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkActivo_GotFocus()
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
    
    pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
    'cmdGrabarRegistro.Enabled = True
    
    SSTObj.TabEnabled(1) = True
    SSTObj.TabEnabled(2) = True
    SSTObj.TabEnabled(3) = True
    SSTObj.TabEnabled(4) = False
End Sub

Private Sub chkActivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub


Private Sub chkBasico_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
     If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAgregar.SetFocus
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":chkBasico_KeyPress"))
End Sub

Private Sub ChkValidaPaquete_GotFocus()
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
    
    pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
    'cmdGrabarRegistro.Enabled = True
    
    SSTObj.TabEnabled(1) = True
    SSTObj.TabEnabled(2) = True
    SSTObj.TabEnabled(3) = True
    SSTObj.TabEnabled(4) = False
End Sub

Private Sub ChkValidaPaquete_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkMedicamentos_GotFocus()
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub chkprecio_Click()
    cmdActualizar.Enabled = IIf(chkprecio.Value, False, True)
    If chkprecio.Value Then pRecalculaPrecios
End Sub

Private Sub chkprecio_GotFocus()
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
    
    pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
    'cmdGrabarRegistro.Enabled = True
    
    SSTObj.TabEnabled(1) = True
    SSTObj.TabEnabled(2) = True
    SSTObj.TabEnabled(3) = True
    SSTObj.TabEnabled(4) = False
End Sub

Private Sub chkprecio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdActualizar_Click()
    pRecalculaPrecios
End Sub

Private Sub cmdAsignaTodo_Click()
    pAsigna True, True
End Sub

Private Sub cmdAsignaUno_Click()
    pAsigna True
End Sub

Private Sub cmdBuscarPresupuestos_Click()
    pLlenaGridPresupuestos
End Sub

Private Sub cmdEliminaTodo_Click()
    pAsigna False, True
End Sub

Private Sub cmdEliminaUno_Click()
    pAsigna False
End Sub

Private Sub cmdExcluir_Click(Index As Integer)
    grdPaquete_dblClick
End Sub

Private Sub cmdExcluir_GotFocus(Index As Integer)
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        'cmdGrabarRegistro.Enabled = True
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro

        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
Dim o_Excel As Object
Dim o_Libro As Object
Dim o_Sheet As Object
Dim intRowExcel As Integer
Dim intRow As Integer
Dim vlintcontador As Integer

CDgArchivo.FileName = ""
CDgArchivo.CancelError = False
CDgArchivo.InitDir = App.Path
CDgArchivo.Filter = "Documentos excel(*.xls;*.xlsx)|*.xls;*.xlsx"
CDgArchivo.DialogTitle = "Exportación de cargos"
CDgArchivo.FilterIndex = 1
CDgArchivo.Flags = cdlOFNOverwritePrompt
CDgArchivo.ShowSave

If CDgArchivo.FileName = "" Then
    freBarra.Visible = False
    Exit Sub
End If

Set o_Excel = CreateObject("Excel.Application")
Set o_Libro = o_Excel.Workbooks.Add
Set o_Sheet = o_Libro.Worksheets(1)

If Not IsObject(o_Excel) Then
    MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
       vbExclamation, "Mensaje"
    Exit Sub
End If
intRowExcel = 4

lblTextoBarra.Caption = "Exportando información, por favor espere..."
freBarra.Top = 4000
freBarra.Visible = True
freBarra.Refresh
pgbCargando.Value = 0
lblTextoBarra.Refresh
'columnas de titulo principal del paquete
o_Excel.Cells(1, 1).Value = "Número"
o_Excel.Cells(1, 2).Value = Trim(txtCvePaquete.Text)
o_Excel.Cells(2, 1).Value = "Descripción"
o_Excel.Cells(2, 2).Value = Trim(txtDescripcion.Text)
o_Excel.range(o_Excel.Cells(2, 2), o_Excel.Cells(2, 5)).Merge
'columnas titulos
o_Excel.Cells(3, 1).Value = "Clave"
o_Excel.Cells(3, 2).Value = "Clave externa"
o_Excel.Cells(3, 3).Value = "Tipo cargo"
o_Excel.Cells(3, 4).Value = "Descripción"
o_Excel.Cells(3, 5).Value = "Cantidad"
o_Excel.Cells(3, 6).Value = "Unidad"
o_Excel.Cells(3, 7).Value = "Monto límite"
o_Excel.Cells(3, 8).Value = "Costo base"
o_Excel.Cells(3, 9).Value = "Total costo"
o_Excel.Cells(3, 10).Value = "Precio"
o_Excel.Cells(3, 11).Value = "Importe"
o_Excel.Cells(3, 12).Value = "Descuento"
o_Excel.Cells(3, 13).Value = "IVA"
o_Excel.Cells(3, 14).Value = "Total"
o_Excel.Cells(3, 15).Value = "Familia"
o_Excel.Cells(3, 16).Value = "Subfamilia"
'Diseño excel
o_Sheet.range("A3:R3").HorizontalAlignment = -4108
o_Sheet.range("A3:R3").VerticalAlignment = -4108
o_Sheet.range("A3:R3").WrapText = True
o_Sheet.range("A4").Select
o_Excel.ActiveWindow.FreezePanes = True
o_Sheet.range("A3:R3").Interior.ColorIndex = 15 '15 48
o_Sheet.range("B:B").HorizontalAlignment = -4152
o_Sheet.range("B3:B3").HorizontalAlignment = -4108

'Tamaños celdas
o_Sheet.range("A:A").Columnwidth = 12
o_Sheet.range("B:B").Columnwidth = 13
o_Sheet.range("C:C").Columnwidth = 10
o_Sheet.range("D:D").Columnwidth = 50
o_Sheet.range("F:F").Columnwidth = 13
o_Sheet.range("G:G").Columnwidth = 12
o_Sheet.range("H:N").Columnwidth = 12
o_Sheet.range("O:P").Columnwidth = 25
'Formato
o_Sheet.range("B:B").NumberFormat = "@"
o_Sheet.range("H:N").NumberFormat = "$#,##0.00"

pgbCargando.Max = grdPaquete.Rows - 2

'Recorre el grid y llena el Excel
For intRow = 2 To grdPaquete.Rows - 1
    If grdPaquete.RowHeight(intRow - 1) > 0 Then
        With grdPaquete
            o_Sheet.Cells(intRowExcel, 1).NumberFormat = "@"
            If (.TextMatrix(intRow, 0) = "AR") Then
                o_Sheet.Cells(intRowExcel, 1).Value = .TextMatrix(intRow, 31) & " " 'Clave del articulo
            Else
                o_Sheet.Cells(intRowExcel, 1).Value = .TextMatrix(intRow, 15) & " " 'Clave del cargo
            End If
            o_Sheet.Cells(intRowExcel, 2).Value = IIf(.TextMatrix(intRow, 0) = "AR", .TextMatrix(intRow, 25) & "  ", "  ") 'Clave externa articulo
            o_Sheet.Cells(intRowExcel, 3).Value = .TextMatrix(intRow, 0) & " " 'Chr tipo cargo
            o_Sheet.Cells(intRowExcel, 4).Value = .TextMatrix(intRow, 1) & " " 'Descripcion
            o_Sheet.Cells(intRowExcel, 5).Value = .TextMatrix(intRow, 2) & " " 'Cantidad
            o_Sheet.Cells(intRowExcel, 6).Value = .TextMatrix(intRow, 3) & " " 'Unidad
            o_Sheet.Cells(intRowExcel, 7).Value = .TextMatrix(intRow, 4) & " " 'Monto limite
            o_Sheet.Cells(intRowExcel, 8).Value = Val(Format(.TextMatrix(intRow, 5), "############.##")) & " " 'Costo base
            o_Sheet.Cells(intRowExcel, 9).Value = Val(Format(.TextMatrix(intRow, 6), "############.##")) & " " 'Total costo
            o_Sheet.Cells(intRowExcel, 10).Value = Val(Format(.TextMatrix(intRow, 7), "############.##")) & " " 'Precio
            o_Sheet.Cells(intRowExcel, 11).Value = Val(Format(.TextMatrix(intRow, 8), "############.##")) & " " 'Importe
            o_Sheet.Cells(intRowExcel, 12).Value = Val(Format(.TextMatrix(intRow, 9), "############.##")) & " " 'Descuento
            o_Sheet.Cells(intRowExcel, 13).Value = Val(Format(.TextMatrix(intRow, 11), "############.##")) & " " 'IVA
            o_Sheet.Cells(intRowExcel, 14).Value = Val(Format(.TextMatrix(intRow, 12), "############.##")) & " " 'Total
            o_Sheet.Cells(intRowExcel, 15).Value = IIf(.TextMatrix(intRow, 0) = "AR", .TextMatrix(intRow, 27), "") 'Familia
            o_Sheet.Cells(intRowExcel, 16).Value = IIf(.TextMatrix(intRow, 0) = "AR", .TextMatrix(intRow, 28), "") 'SubFamilia
        End With
        intRowExcel = intRowExcel + 1
    End If
    pgbCargando.Value = pgbCargando.Value + 1
Next



'La información ha sido exportada exitosamente
If CDgArchivo.FileName <> "" Then
    o_Excel.DisplayAlerts = False 'Deshabilitamos la alerta de reemplazar porque ya esta al seleccionar el archivo
    o_Libro.SaveAs CDgArchivo.FileName, -4143
    o_Excel.DisplayAlerts = True 'Habilitamos
End If
MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
o_Excel.Visible = True
freBarra.Visible = False
pgbCargando.Max = 100
pgbCargando.Value = 0

Set o_Excel = Nothing


Exit Sub

NotificaError:
    ' -- Cierra la hoja y la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Sheet Is Nothing Then Set o_Sheet = Nothing
    freBarra.Visible = False
    pgbCargando.Max = 100
    pgbCargando.Value = 0
    lblTextoBarra.Caption = "Cargando datos, por favor espere..."
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExportar_Click"))

End Sub

Private Sub cmdGrabarRegistro_Click()
'**********************************************************
' Procedimiento para grabar una alta o modificación       *
'**********************************************************
    Dim rsDetallePaquete As New ADODB.Recordset
    Dim rsDescuentoInventario As New ADODB.Recordset
    Dim vlintcontador As Integer
    Dim vllngPersonaGraba As Long
    Dim vlstrSentenciaDescuentoInventario As String
    Dim vlstrSentencia As String
    Dim vlintRow As Integer
    
    Dim rsCostoPaquete As New ADODB.Recordset
    Dim rsHonorarios As New ADODB.Recordset
    Dim vlintsubtotal As Double
    Dim vlIntCont As Integer
    
    If fblnDatosValidos() Then
        With rsPaquetes
            '--------------------------------------------------------
            ' Persona que graba
            '--------------------------------------------------------
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba = 0 Then Exit Sub
            
            '--------------------------------------
            EntornoSIHO.ConeccionSIHO.BeginTrans
            '--------------------------------------
            ' Grabar el Concepto de Facturación
            '--------------------------------------
            If vgstrEstadoManto = "A" Then ' Solo cuando es una alta
                .AddNew
            End If
            !chrDescripcion = Trim(txtDescripcion.Text)
            !bitactivo = chkActivo.Value
            !bitValidaCargosPaquete = ChkValidaPaquete.Value
            !SMICONCEPTOFACTURA = cboConceptoFactura.ItemData(cboConceptoFactura.ListIndex)
            !chrTratamiento = Trim(cboTratamiento.List(cboTratamiento.ListIndex))
            !chrTipo = Trim(cboTipo.List(cboTipo.ListIndex))
            !mnyAnticipoSugerido = Val(Format(txtAnticipo.Text, "###############.##"))
            !dtmFechaActualizacion = IIf(vgdtmFechaActualizacion = CDate("01/01/1900"), IIf(mskFecha <> "  /  /    ", mskFecha, Null), vgdtmFechaActualizacion)
            !bitcostobase = IIf(OptPolitica(0).Value, 0, 1)
            !chrTipoIngresoDescuento = IIf(optTipoPacienteDesc(0).Value, "T", IIf(optTipoPacienteDesc(1).Value, "I", IIf(optTipoPacienteDesc(2).Value, "E", "U")))
            !bitincrementoautomatico = chkprecio.Value
            !bitSeleccionarHonorarios = IIf(optFormaAsignacionSeleccionar.Value, 1, 0)
            If vgstrEstadoManto = "A" Then ' Solo cuando es una alta
                !dtmFechaAltaPaquete = Now
            End If
            
           '!bitbasico = chkBasico.Value BITSELECCIONARHONORARIOS
            .Update
            If vgstrEstadoManto = "A" Then ' Solo cuando es una alta
                txtCvePaquete.Text = flngObtieneIdentity("SEC_PVPAQUETE", rsPaquetes!intnumpaquete)
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "PAQUETE", txtCvePaquete.Text)
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "PAQUETE", txtCvePaquete.Text)
            End If
            
            If !bitincrementoautomatico = 1 Then
           ' vlintsubtotal = Val(Mid(txtSubtotal, 2, Len(txtSubtotal)))
                vlintsubtotal = Val(Format(Mid(txtSubtotal, 2, Len(txtSubtotal)), "#############.00"))
                vgstrParametrosSP = Val(vlintsubtotal) & "|" & Val(txtCvePaquete)
                frsEjecuta_SP vgstrParametrosSP, "SP_PVACTUALIZALISTAPRECIOS"
            
            End If
            '---------------------------------------------------------
            ' Grabar los Items incluidos en ese paquete, sólo cuando ya se dio de alta
            '---------------------------------------------------------
            vlstrSentenciaDescuentoInventario = "SELECT * FROM PvDetallePaquete WHERE intNumPaquete = " & CLng(txtCvePaquete.Text)
            Set rsDescuentoInventario = frsRegresaRs(vlstrSentenciaDescuentoInventario, adLockReadOnly, adOpenForwardOnly)
            If rsDescuentoInventario.RecordCount > 0 Then
                Do While Not rsDescuentoInventario.EOF
                    vlintcontador = 2
                    Do While vlintcontador <= grdPaquete.Rows - 1
                        If Val(grdPaquete.TextMatrix(vlintcontador, 15)) = rsDescuentoInventario!intCveCargo And Val(Format(grdPaquete.TextMatrix(vlintcontador, 7), "###############.##")) = rsDescuentoInventario!mnyPrecio Then
                            If grdPaquete.TextMatrix(vlintcontador, 13) = rsDescuentoInventario!INTDESCUENTOINVENTARIO Then
                                grdPaquete.TextMatrix(vlintcontador, 13) = rsDescuentoInventario!INTDESCUENTOINVENTARIO 'Unidad
                            End If
                            Exit Do
                        End If
                        vlintcontador = vlintcontador + 1
                    Loop
                    rsDescuentoInventario.MoveNext
                Loop
            End If
            rsDescuentoInventario.Close
            
            If vlblnAgregaNuevo Then
                'El paquete maneja precios específicos por cargo, por lo que el precio de venta se actualizará a cero en las listas de precios. Será necesario guardar de nuevo la configuración de precios específicos por cargo para que se calcule un nuevo precio de venta y se actualicen las listas de precios.
                MsgBox SIHOMsg(1623), vbOKOnly + vbInformation, "Mensaje"
                vlstrSentencia = "UPDATE PVDETALLELISTA SET MNYPRECIO = 0" & _
                " WHERE CHRCVECARGO = " & txtCvePaquete.Text & _
                " AND CHRTIPOCARGO = '" & "PA" & "'"
                pEjecutaSentencia vlstrSentencia
            End If
                
                Call PBorraCargosAsignados(txtCvePaquete.Text)
                vlstrsql = "select * from PvDetallePaquete where intNumPaquete = -5" ' Paque no regrese nada
                Set rsDetallePaquete = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
                If IIf(grdPaquete.TextMatrix(2, 15) = "", -1, grdPaquete.TextMatrix(2, 15)) <> -1 Then
                    For vlintcontador = 2 To grdPaquete.Rows - 1
                        rsDetallePaquete.AddNew
                        rsDetallePaquete!intnumpaquete = IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text))
                        rsDetallePaquete!intCveCargo = grdPaquete.TextMatrix(vlintcontador, 15)
                        rsDetallePaquete!chrTipoCargo = grdPaquete.TextMatrix(vlintcontador, 0)
                        rsDetallePaquete!SMICANTIDAD = grdPaquete.TextMatrix(vlintcontador, 2)
                        rsDetallePaquete!INTDESCUENTOINVENTARIO = CInt(Val(grdPaquete.TextMatrix(vlintcontador, 13)))
                        rsDetallePaquete!mnyMontoLimite = Val(Format(grdPaquete.TextMatrix(vlintcontador, 4), "############.##"))
                        rsDetallePaquete!mnycosto = Val(Format(grdPaquete.TextMatrix(vlintcontador, 5), "############.##"))
                        rsDetallePaquete!mnyPrecio = Val(Format(grdPaquete.TextMatrix(vlintcontador, 7), "############.##"))
                        rsDetallePaquete!MNYDESCUENTO = Val(Format(grdPaquete.TextMatrix(vlintcontador, 9), "############.##"))
                        rsDetallePaquete!MNYIVA = Val(Format(grdPaquete.TextMatrix(vlintcontador, 11), "############.##"))
                        rsDetallePaquete!smicveconcepto = IIf(grdPaquete.TextMatrix(vlintcontador, 0) = "GC", flngBuscaConcepto(grdPaquete.TextMatrix(vlintcontador, 15)), 0)
                        rsDetallePaquete!MNYPRECIOESPECIFICO = grdPaquete.TextMatrix(vlintcontador, 24)
                        rsDetallePaquete.Update
                    Next
                    
                End If
                rsDetallePaquete.Close
            'End If
            
            Set rsCostoPaquete = frsRegresaRs("SELECT * FROM PVCOSTOCARGOS WHERE intCveEmpresaContable = " & vgintClaveEmpresaContable & " AND chrTipo = 'PA' AND intCveCargo = " & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text)), adLockOptimistic, adOpenDynamic)
            If rsCostoPaquete.RecordCount > 0 Then
                pEjecutaSentencia ("UPDATE PVCOSTOCARGOS SET NumCosto = " & CDbl(grdTotales.TextMatrix(8, 2)) & " WHERE intCveEmpresaContable = " & vgintClaveEmpresaContable & " AND chrTipo = 'PA' AND intCveCargo = " & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text)))
            Else
                pEjecutaSentencia ("INSERT INTO PVCOSTOCARGOS VALUES(" & vgintClaveEmpresaContable & "," & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text)) & ",'PA'," & CDbl(IIf(grdTotales.TextMatrix(8, 2) <> "", grdTotales.TextMatrix(8, 2), 0)) & ")")
            End If
            rsCostoPaquete.Close
            
            'Se borra la relacion del paquete con sus departamentos
            vlstrSentencia = "DELETE FROM PVPAQUETEDEPARTAMENTO WHERE INTNUMPAQUETE = " & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text))
            pEjecutaSentencia vlstrSentencia

            'Se alimentan los departamentos para el paquete con los nuevos seleccionados
            For vlintRow = 0 To lstDepartamentosSel.ListCount - 1
                vlstrSentencia = "INSERT INTO PVPAQUETEDEPARTAMENTO VALUES(" & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text)) & "," & lstDepartamentosSel.ItemData(vlintRow) & ")"
                pEjecutaSentencia vlstrSentencia
            Next vlintRow
            
            'Guardado de honorarios médicos
            If grdDetalleHonorarios.Rows > 1 Then
                If Trim(grdDetalleHonorarios.TextMatrix(1, 1)) <> "" Then
                    ''If (grdDetalleHonorarios.Rows > 1 Or (grdDetalleHonorarios.Rows = 2 And Trim(grdDetalleHonorarios.TextMatrix(1, 1))) <> "") Then
                         'Se borra la relacion del paquete
                        vlstrSentencia = "delete from pvPaqueteHonorarios where intCvePaquete = " & IIf(CLng(txtCvePaquete.Text) < 0, CLng(txtCvePaquete.Text) * -1, CLng(txtCvePaquete.Text))
                        pEjecutaSentencia vlstrSentencia
                        
                        vlstrSentencia = "select * from pvPaqueteHonorarios where intcvePaquete=-1"
                        Set rsHonorarios = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                         
                        For vlintcontador = 1 To grdDetalleHonorarios.Rows - 1
                                rsHonorarios.AddNew
                                rsHonorarios!intCvePaquete = Trim(txtCvePaquete.Text)
                                rsHonorarios!INTCVEFUNCION = Val(grdDetalleHonorarios.TextMatrix(vlintcontador, cintColClaveFuncion))
                                rsHonorarios!intCveConcepto = Val(grdDetalleHonorarios.TextMatrix(vlintcontador, cintColClaveConcepto))
                                rsHonorarios!MNYIMPORTEHONORARIO = Val(Format(grdDetalleHonorarios.TextMatrix(vlintcontador, cintColImporteHonorario), "############.##"))
                                rsHonorarios.Update
                        Next vlintcontador
                        rsHonorarios.Close
                End If
            End If
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End With
        rsPaquetes.Requery
        Call pNuevoRegistro
        
    End If
End Sub

Private Function fblnDatosValidos() As Boolean
    Dim vlintcontador As Integer
    
    fblnDatosValidos = True
    
    If RTrim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(2) + Chr(13) + txtDescripcion.ToolTipText, vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    If fblnDatosValidos And cboTratamiento.ListIndex = -1 Then
        fblnDatosValidos = False
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
        cboTratamiento.SetFocus
    End If
    If fblnDatosValidos And cboTipo.ListIndex = -1 Then
        fblnDatosValidos = False
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
        cboTipo.SetFocus
    End If
    If fblnDatosValidos And cboConceptoFactura.ListIndex = -1 Then
        fblnDatosValidos = False
        'Seleccione el dato.
        MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
        cboConceptoFactura.SetFocus
    End If

    With grdGrupos
        For vlintcontador = 1 To .Rows - 1
            If Val(.TextMatrix(vlintcontador, 2)) = -1 Then
                fblnDatosValidos = False
                'Seleccione el dato.
                MsgBox SIHOMsg(431), vbOKOnly + vbInformation, "Mensaje"
                'MsgBox SIHOMsg(431) & " concepto de facturación para excedente por grupo de cargo", vbOKOnly + vbInformation, "Mensaje"
                .Col = 3
                .Row = vlintcontador
                .SetFocus
'                Call pEditarConcepto(13, cboMovConceptoFactura, grdGrupos)
                Exit Function
            End If
        Next
    End With

End Function

Private Sub pConceptosFactura()
    Dim rsConcepto As New ADODB.Recordset
    Dim rsConceptoHosp As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "SELECT smiCveConcepto, chrDescripcion " & _
                     "FROM PvConceptoFacturacion " & _
                     "WHERE bitActivo = 1 " & _
                     "ORDER BY chrDescripcion"
    Set rsConcepto = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboConceptoFactura, rsConcepto, 0, 1
    rsConcepto.Close
    
    vlstrSentencia = "SELECT smiCveConcepto, TRIM(chrDescripcion) " & _
                     "FROM PvConceptoFacturacion " & _
                     "WHERE bitActivo = 1 " & _
                     "AND intTipo = 0"
    Set rsConceptoHosp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    pLlenarCboRs cboMovConceptoFactura, rsConceptoHosp, 0, 1
    cboMovConceptoFactura.ListIndex = 0
    rsConceptoHosp.Close
End Sub

Private Sub pConfiguraGrid()
    With grdHBusqueda
        .FormatString = "|Número|Descripción|Tratamiento|Tipo|Estado"
        .ColWidth(0) = 100 'Fix
        .ColWidth(1) = 1400  'Clave
        .ColWidth(2) = 6500 'Descripcion
        .ColWidth(3) = 1433 'Tratamiento
        .ColWidth(4) = 1433 'Tipo
        .ColWidth(5) = 1433 'Estado
        
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarVertical
    End With
End Sub

Private Sub pConfiguraGridCargos()
    With grdPaquete
        .FixedCols = 2
        .Rows = 2
        '.FormatString = "|Descripción|Cant.||Monto limite|Costo base|Total costo|Precio|Descuento|Subtotal|IVA|Total||"
        
        .MergeCells = flexMergeRestrictRows
        .TextMatrix(0, 0) = " "
        .TextMatrix(0, 1) = " "
        .TextMatrix(0, 2) = " "
        .TextMatrix(0, 3) = " "
        .TextMatrix(0, 4) = " "
        .TextMatrix(0, 5) = " "
        .TextMatrix(0, 6) = " "
        .TextMatrix(0, 7) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 8) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 9) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 10) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 11) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 12) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 16) = " "
        .TextMatrix(0, 17) = " "
        .TextMatrix(0, 18) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 19) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 20) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 21) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 22) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 23) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", "  ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 24) = " "
        .TextMatrix(0, 25) = " "
        .TextMatrix(0, 26) = " "
        .TextMatrix(0, 27) = " " 'Familia segun sea el caso
        .TextMatrix(0, 28) = " " 'Subfamilia segun sea el caso
        .TextMatrix(0, 29) = " " 'Monto limite, es el original para comparar cuando editen
        .TextMatrix(0, 30) = " " 'Cantidad, es la original para comparar cuando editen
        .TextMatrix(0, 31) = " " 'Chr del articulo
        .MergeRow(0) = True
        
        .TextMatrix(1, 0) = "Descripción"
        .TextMatrix(1, 1) = "Descripción"
        .TextMatrix(1, 2) = "Cantidad"
        .TextMatrix(1, 3) = "Unidad"
        .TextMatrix(1, 4) = "Monto límite"
        .TextMatrix(1, 5) = "Costo base"
        .TextMatrix(1, 6) = "Total costo"
        
        .TextMatrix(1, 7) = "Precio"
        .TextMatrix(1, 8) = "Importe"
        .TextMatrix(1, 9) = "Descuento"
        .TextMatrix(1, 10) = "Subtotal"
        .TextMatrix(1, 11) = "IVA"
        .TextMatrix(1, 12) = "Total"
        
        .TextMatrix(1, 16) = "Costo base"
        .TextMatrix(1, 17) = "Total costo"
        
        .TextMatrix(1, 18) = "Precio"
        .TextMatrix(1, 19) = "Importe"
        .TextMatrix(1, 20) = "Descuento"
        .TextMatrix(1, 21) = "Subtotal"
        .TextMatrix(1, 22) = "IVA"
        .TextMatrix(1, 23) = "Total"
        
        .MergeRow(1) = True
        
        .ColWidth(0) = 310      ' Tipo
        .ColWidth(1) = 3800     ' Descripción
        .ColWidth(2) = 700      ' Cantidad
        .ColWidth(3) = 2000     ' Unidad
        .ColWidth(4) = 940      ' Monto limite
        .ColWidth(5) = 863      ' Costo base
        .ColWidth(6) = 968      ' Total costo
        .ColWidth(7) = 880      ' Precio
        .ColWidth(8) = 968      ' Importe
        .ColWidth(9) = 880      ' Descuento
        .ColWidth(10) = 968      ' Subtotal
        .ColWidth(11) = 880     ' IVA
        .ColWidth(12) = 968     ' Total
        
        .ColWidth(13) = 0       ' Tipo de descuento de inventario
        .ColWidth(14) = 0       ' Contenido
        .ColWidth(15) = 0       ' Clave del elemento
        
        .ColWidth(16) = 0     ' Costo base
        .ColWidth(17) = 0       ' Total costo
        
        .ColWidth(18) = 0       ' Precio
        .ColWidth(19) = 0       ' Importe
        .ColWidth(20) = 0       ' Descuento
        .ColWidth(21) = 0       ' Subtotal
        .ColWidth(22) = 0       ' IVA
        .ColWidth(23) = 0       ' Total
        .ColWidth(24) = 0       ' PrecioEspecifico
        .ColWidth(25) = 0       ' Clave externa
        .ColWidth(26) = 0       ' Conc facturacion
        .ColWidth(27) = 0       ' Familia
        .ColWidth(28) = 0       ' Subfamilia
        .ColWidth(29) = 0       'Monto limite original
        .ColWidth(30) = 0       'Cantidad original
        .ColWidth(31) = 0
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(13) = flexAlignRightCenter
        .ColAlignment(14) = flexAlignRightCenter
        .ColAlignment(15) = flexAlignRightCenter
        .ColAlignment(16) = flexAlignRightCenter
        .ColAlignment(17) = flexAlignRightCenter
        .ColAlignment(18) = flexAlignRightCenter
        .ColAlignment(19) = flexAlignRightCenter
        .ColAlignment(20) = flexAlignRightCenter
        .ColAlignment(21) = flexAlignRightCenter
        .ColAlignment(22) = flexAlignRightCenter
        .ColAlignment(23) = flexAlignRightCenter

        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
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
        .FixedAlignment(12) = flexAlignCenterCenter
        .FixedAlignment(13) = flexAlignCenterCenter
        .FixedAlignment(14) = flexAlignCenterCenter
        .FixedAlignment(15) = flexAlignCenterCenter
        .FixedAlignment(16) = flexAlignCenterCenter
        .FixedAlignment(17) = flexAlignCenterCenter
        .FixedAlignment(18) = flexAlignCenterCenter
        .FixedAlignment(19) = flexAlignCenterCenter
        .FixedAlignment(20) = flexAlignCenterCenter
        .FixedAlignment(21) = flexAlignCenterCenter
        .FixedAlignment(22) = flexAlignCenterCenter
        .FixedAlignment(23) = flexAlignCenterCenter
        
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 15) = -1
    End With
End Sub

Private Sub cmdImportar_Click()
On Error GoTo NotificaError
Dim objXLApp As Object
Dim txtRuta As String
Dim intLoopCounter As Integer
Dim txtClave As String
Dim intRows As Integer
Dim vblnBandera As Boolean
Dim intResul As Integer
Dim intRowsAct As Integer
Dim lngCantidad As Long
Dim strMontoLimite As String
Dim strTipoCargo As String
Dim strTipoCargoGrid As String
Dim strDescripcion As String
Dim strDescripcionGrid As String
Dim strUnidad As String
Dim strUnidadGrid As String
Dim vlstrValidador As String

Set objXLApp = CreateObject("Excel.Application")
    CommonDialog1.DialogTitle = "Abrir archivo"
    CommonDialog1.Filter = "Documentos excel|*.xls;*.xlsx;"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.CancelError = True
    On Error Resume Next
    CommonDialog1.ShowOpen
    If Err Then
        'Si se cancela el cuadro de diálogo
        Exit Sub
    End If
    txtRuta = CommonDialog1.FileName
    intRowsAct = grdPaquete.Rows - 2
    lblTextoBarra.Caption = "Importando información, por favor espere..."
    freBarra.Top = 4000
    freBarra.Visible = True
    freBarra.Refresh
    pgbCargando.Value = 0
    lblTextoBarra.Refresh
    
   
    With objXLApp
        .Workbooks.Open txtRuta
        .Workbooks(1).Worksheets(1).Select
         pgbCargando.Max = (CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 3)
        'Verificamos que sea la misma cantidad de articulos del grid contra el excel
'        If grdPaquete.Rows - 2 <> (CInt(.cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 3) Then
'            'Información no valida
'            MsgBox "La cantidad de cargos que se intentan importar son diferentes a los cargos que ya se tienen asignados.", vbOKOnly + vbInformation, "Mensaje"
'            .Workbooks(1).Close False
'            .Quit
'            objXLApp = Nothing
'            pRestablecerBarra
'            Exit Sub
'        End If
        
        If Trim(.range("B" & 1)) <> Trim(txtCvePaquete) Then
            vlstrValidador = "El número del paquete no coincide con el documento que se intenta importar."
        End If
        
        If Trim(.range("B" & 2)) <> Trim(txtDescripcion) Then
            vlstrValidador = "La descripción del paquete no coincide con el documento que se intenta importar."
        End If
        
        If vlstrValidador <> "" Then
            'Información no valida
            MsgBox vlstrValidador, vbOKOnly + vbInformation, "Mensaje"
            .Workbooks(1).Close False
            .Quit
            objXLApp = Nothing
            pRestablecerBarra
            Exit Sub
        End If
        
        'Validamos que no repitan cargos
        For intLoopCounter = 4 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
            'Entra el primer articulo y lo comparamos con todos, solo se puede estar 1 vez.
            Dim intLoopTotal As Integer
            Dim vlstrTipoCargoE As String
            Dim vlngCveCargoE As Long
            Dim vlngTotalEncontrados As Long
            
            vlstrTipoCargoE = .range("C" & intLoopCounter)
            vlngCveCargoE = .range("A" & intLoopCounter)
            vlngTotalEncontrados = 0
            
            For intLoopTotal = 4 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
                If vlstrTipoCargoE = .range("C" & intLoopTotal) Then
                    If vlngCveCargoE = .range("A" & intLoopTotal) Then
                        vlngTotalEncontrados = vlngTotalEncontrados + 1
                    End If
                End If
            Next intLoopTotal
            
            If vlngTotalEncontrados > 1 Then
                MsgBox "El cargo ya está incluido en la lista." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                pRestablecerBarra
                Exit Sub
            End If
            
        Next intLoopCounter
        'Validamos los datos de excel, cantidad y monto limite, si es que es GC
        For intLoopCounter = 4 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
            lngCantidad = .range("E" & intLoopCounter)
            If Trim(.range("E" & intLoopCounter)) <> "" Then
                If Not IsNumeric(Trim(.range("E" & intLoopCounter))) Then
                    'La cantidad del Cargo es incorrecta
                    MsgBox "La cantidad del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
                End If
                If Trim(.range("E" & intLoopCounter)) < 0 Then
                    'La cantidad del Cargo es incorrecta
                    MsgBox "La cantidad del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
                End If
                If (InStr(1, CDbl(Trim(.range("E" & intLoopCounter))), ".") >= 1) Or (InStr(1, CDbl(Trim(.range("E" & intLoopCounter))), ",") >= 1) Then
                    'La cantidad del Cargo es incorrecta
                    MsgBox "La cantidad del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
                End If
            Else
                'La cantidad del Cargo es incorrecta
                MsgBox "La cantidad del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                pRestablecerBarra
                Exit Sub
            End If
            
            If Trim(.range("C" & intLoopCounter)) <> "GC" Then
                If Trim(.range("G" & intLoopCounter)) <> "" Then
                    MsgBox "El monto límite del cargo solo aplica para grupos de cargos." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
                End If
            End If
                        
            If Trim(.range("C" & intLoopCounter)) = "GC" Then
                If Trim(.range("G" & intLoopCounter)) <> "" Then
                    If Not IsNumeric(Trim(.range("H" & intLoopCounter))) Then
                        MsgBox "El monto límite del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        pRestablecerBarra
                        Exit Sub
                    End If
                    If Trim(.range("G" & intLoopCounter)) < 0 Then
                        MsgBox "El monto límite del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        pRestablecerBarra
                        Exit Sub
                    End If
                Else
                    MsgBox "El monto límite del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
                End If
                strMontoLimite = .range("G" & intLoopCounter)
            End If
            If Not IsNumeric(Trim(.range("A" & intLoopCounter))) Then
                MsgBox "La clave del cargo es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                pRestablecerBarra
                Exit Sub
            End If
            'Despues validamos la información con los que estan en la base de datos
            Dim vlstrSentencia As String
            Dim rsCargo As New ADODB.Recordset
            Dim vlstrTipoCargo As String
            Dim vsltrCveCargo As String
            Dim strArticulo As String
            
            vlstrTipoCargo = Trim(.range("C" & intLoopCounter))
            vsltrCveCargo = Trim(.range("A" & intLoopCounter))
            Select Case Trim(vlstrTipoCargo)
                Case "AR"
                    vlstrSentencia = "Select substring(vchNombreComercial,1,50) Articulo, alterna.VCHDESCRIPCION alterna, minima.VCHDESCRIPCION minima, intcontenido  From ivArticulo " & _
                    " inner join IvUnidadVenta alterna on alterna.INTCVEUNIDADVENTA = ivArticulo.INTCVEUNIALTERNAVTA " & _
                    " inner join IvUnidadVenta minima on minima.INTCVEUNIDADVENTA = ivArticulo.INTCVEUNIMINIMAVTA WHERE CHRCVEARTICULO = " & vsltrCveCargo
                Case "ES"
                    vlstrSentencia = "SELECT substring(VCHNOMBRE,1,50) Articulo FROM IMESTUDIO WHERE intCveEstudio = " & vsltrCveCargo
                Case "EX"
                    vlstrSentencia = "SELECT substring(CHRNOMBRE,1,50) Articulo FROM LAEXAMEN WHERE intCveExamen = " & vsltrCveCargo
                Case "GE"
                    vlstrSentencia = "SELECT substring(CHRNOMBRE,1,50) Articulo FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vsltrCveCargo
                Case "OC"
                    vlstrSentencia = "SELECT substring(CHRDESCRIPCION,1,50) Articulo FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vsltrCveCargo
                Case Else 'GC
                    vlstrSentencia = "SELECT substring(VCHNOMBRE,1,50) Articulo FROM PVGRUPOCARGO WHERE INTCVEGRUPO = " & vsltrCveCargo
            End Select

            Set rsCargo = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
            If rsCargo.RecordCount = 0 Then
                Dim vlstrTipo As String
                Select Case Trim(vlstrTipoCargo)
                     Case "AR"
                        vlstrTipo = "artículo"
                    Case "ES"
                        vlstrTipo = "estudio"
                    Case "EX"
                        vlstrTipo = "examen"
                    Case "GE"
                        vlstrTipo = "grupo de examen"
                    Case "OC"
                        vlstrTipo = "otros conceptos"
                    Case Else
                        vlstrTipo = "grupos de cargos"
                End Select
                    MsgBox "La clave del " & vlstrTipo & " es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    pRestablecerBarra
                    Exit Sub
            Else
                If Trim(vlstrTipoCargo) = "AR" Then
                    Dim vlstrTipoUnidad As String
                    Dim vlintContenido As Integer
                    
                    vlstrTipoUnidad = Trim(objXLApp.range("F" & intLoopCounter))
                    vlintContenido = rsCargo!intContenido
                    
                    If vlintContenido = 1 Then
                        If Trim(vlstrTipoUnidad) <> Trim(rsCargo!alterna) Then
                            MsgBox "El tipo de unidad es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            pRestablecerBarra
                            Exit Sub
                        End If
                    Else
                        If Trim(vlstrTipoUnidad) <> Trim(rsCargo!alterna) Then
                            If Trim(vlstrTipoUnidad) <> Trim(rsCargo!MINIMA) Then
                                MsgBox "El tipo de unidad es incorrecta." & " Renglón " & intLoopCounter & ".", vbOKOnly + vbInformation, "Mensaje"
                                .Workbooks(1).Close False
                                .Quit
                                objXLApp = Nothing
                                pRestablecerBarra
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next intLoopCounter
        'Limpiamos el grid
        grdPaquete.Clear
        grdPaquete.Rows = 3
        pConfiguraGridCargos
        
        Dim intcontador As Integer
        Dim intCanitdad As Integer
        Dim lngCveCargo As String
        Dim intDescuenta As Integer
        Dim lngContenido As Long
        Dim lngTotalCargos As Long
        Dim clveCargo As Long
        Dim lngRow As Long
        Dim vlstrUnidad As String
        Dim intLista As Integer
        intcontador = 2
        vlstrUnidad = ""
        lngTotalCargos = CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) 'Restandole los 4 de los titulos
        For intRows = 4 To lngTotalCargos
            'grdPaquete.TextMatrix(intRows, 2) = Trim(.range("E" & intcontador))
            grdPaquete.Rows = grdPaquete.Rows + 1
            Select Case Trim(.range("C" & intRows))
                Case "AR"
                    intLista = 0
                Case "ES"
                    intLista = 1
                Case "OC"
                    intLista = 3
                Case "GC"
                    intLista = 4
                Case Else
                    intLista = 2
            End Select
            intCanitdad = CInt(Trim(.range("E" & intRows)))
            lngCveCargo = Trim(.range("A" & intRows))
            vlstrUnidad = Trim(.range("F" & intRows))
            grdPaquete.TextMatrix(intcontador, 2) = intCanitdad
            pRecalculaImportesExcel intCanitdad, intcontador, lngCveCargo, intLista, vlstrUnidad, Trim(.range("C" & intRows))
            If Trim(.range("C" & intRows)) = "GC" Then
                grdPaquete.TextMatrix(intcontador, 4) = FormatCurrency(Trim(.range("G" & intRows)), 2)
            End If
            intcontador = intcontador + 1
            lngRow = lngRow + 1
            pgbCargando.Value = pgbCargando.Value + 1
        Next intRows
        .Workbooks(1).Close False
        .Quit
    End With
Set objXLApp = Nothing
MsgBox "La información ha sido importada exitosamente.", vbOKOnly + vbInformation, "Mensaje"
cmdExportar.Enabled = False
'cmdImportar.Enabled = False
pRestablecerBarra
pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Command1_Click"))
    Unload Me
End Sub

Private Sub cmdIncluir_Click(Index As Integer)
    If fblnSeleccionoElemento() Then pSeleccionaElemento
End Sub

Private Sub cmdIncluir_GotFocus(Index As Integer)
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        'cmdGrabarRegistro.Enabled = True
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    
    vlstrx = CStr(CLng(Val(txtCvePaquete.Text))) 'Clave del Paquete
    vlstrx = vlstrx & "|" & IIf(optAgruparRpt(0).Value, 0, IIf(optAgruparRpt(1).Value, 1, 2))
    
    vgstrParametrosSP = CStr(CLng(Val(txtCvePaquete.Text))) & "|" & IIf(optAgruparRpt(0).Value, 0, IIf(optAgruparRpt(1).Value, 1, 2)) & "|" & Trim(vgstrNombreHospitalCH) & "|" & vgintClaveEmpresaContable
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_pvRPTPaquete")
    
    If rsReporte.RecordCount > 0 Then
        vgrptReporte.DiscardSavedData
        pImprimeReporte vgrptReporte, rsReporte, "P", "Paquetes, planes y cirugías"
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    If rsReporte.State <> adStateClosed Then rsReporte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
End Sub

Private Sub cmdSelecciona_Click(Index As Integer)
    If Index = 0 Then
        If fblnSeleccionoElemento() Then
            pSeleccionaElemento
        End If
    Else
        grdPaquete_dblClick
    End If
End Sub

Private Function fblnSeleccionoElemento() As Boolean
    fblnSeleccionoElemento = IIf(lstElementos.ListCount = 0, False, True)
End Function

Private Sub Command1_Click()
    pRecalculaPrecios
End Sub

Private Sub Form_Activate()
    If Not vgblnValida Then
        fblnHabilitaObjetos Me
        SSTObj.Tab = 0
        pConceptosFactura
        If txtDescripcion <> "" Then
        'cboConceptoFactura.ListIndex = 0
            With rsPaquetes
                cboConceptoFactura.ListIndex = fintLocalizaCbo(cboConceptoFactura, !SMICONCEPTOFACTURA)
            End With
        End If
    End If
    vgblnValida = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 7
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Me.Icon = frmMenuPrincipal.Icon
    
    Set rs = frsSelParametros("SI", vgintClaveEmpresaContable, "BITVALIDARCARGOSENPAQUETE")
    If Not rs.EOF Then
        vlintValidarCargosEnPaquete = IIf(IsNull(rs!Valor), 0, Val(rs!Valor))
    End If
    rs.Close
    
    vllngSizeNormal = 5175
    vllngSizeGrande = 10245
    vllngSizeHonorarios = 7000

    Me.Height = vllngSizeNormal
    
    pInstanciaReporte vgrptReporte, "rptReportePaquete.rpt"
            
    vgstrEstadoManto = ""
    vlstrsql = "select * from PvPaquete"
    Set rsPaquetes = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        
    ' Para ver cual es el Tipo de Paciente Particular
    vlstrSentencia = 0
    Set rs = frsSelParametros("SI", -1, "INTTIPOPARTICULAR")
    If rs.RecordCount > 0 Then vglngTipoParticular = IIf(IsNull(rs!Valor), 0, rs!Valor)
    rs.Close
        
    Set rs = frsRegresaRs("SELECT tnyCveTipoPaciente Clave, vchDescripcion " & _
                          "FROM ADTIPOPACIENTE " & _
                          "ORDER BY vchDescripcion", adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboTipoPaciente, rs, 0, 1, 0, False
    cboTipoPaciente.AddItem "<NINGUNO>", 0
    cboTipoPaciente.ListIndex = 0
    rs.Close
        
    Set rs = frsRegresaRs("SELECT intCveEmpresa Clave, vchDescripcion " & _
                          "FROM CCEMPRESA " & _
                          "ORDER BY vchDescripcion", adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboEmpresas, rs, 0, 1, 0, False
    cboEmpresas.AddItem "<NINGUNA>", 0
    cboEmpresas.ListIndex = 0
    rs.Close
    vgblnValida = False
    vgTipoIngresoDescuento = "T"

    pCargarDepartamentos

    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = fdtmServerFecha
    mskFechaInicio.Mask = "##/##/####"

    mskFechaFin.Mask = ""
    mskFechaFin.Text = fdtmServerFecha
    mskFechaFin.Mask = "##/##/####"
    
    vlstrTipoPermiso = fblnRevisaPermisoStr(vglngNumeroLogin)
    cmdExportar.Enabled = False
    cmdImportar.Enabled = False
    
End Sub
Public Function fblnRevisaPermisoStr(vllngxNumeroLogin As Long) As String
'-------------------------------------------------------------------------------
' Validar si un login tiene permiso y por que opcion
'-------------------------------------------------------------------------------
    Dim rspermisodelusuario As New ADODB.Recordset
    Dim vlstrsql As String
    
    If UCase(vgstrNombreUsuario) <> "ADMINISTRADOR" Then
        vlstrsql = "select chrPermiso from Permiso where intNumeroLogin=" + str(vllngxNumeroLogin) + " and intNumeroOpcion=322"
        Set rspermisodelusuario = frsRegresaRs(vlstrsql)
        If rspermisodelusuario.RecordCount <> 0 Then
            If rspermisodelusuario!chrpermiso = "C" Then
                fblnRevisaPermisoStr = "C"
            ElseIf rspermisodelusuario!chrpermiso = "L" Then
                fblnRevisaPermisoStr = "L"
            ElseIf rspermisodelusuario!chrpermiso = "S" Then
                fblnRevisaPermisoStr = "S"
            ElseIf rspermisodelusuario!chrpermiso = "E" Then
                fblnRevisaPermisoStr = "E"
            End If
        End If
    Else
        fblnRevisaPermisoStr = "E"
    End If
    
    
End Function

Public Sub pBloqueaboton(strtipopermiso As String, cmdabloquear As CommandButton)
    If strtipopermiso = "C" Or strtipopermiso = "E" Then
        cmdabloquear.Enabled = True
    ElseIf strtipopermiso = "L" Or strtipopermiso = "S" Then
        cmdabloquear.Enabled = False
    End If
End Sub



Private Function fmskBorrarRegVSFlex(vllngRenglon As Long, grdNombre As VSFlexGrid) As VSFlexGrid
    Dim vllngContador As Long, vllngContador1 As Long
                    
    With grdNombre
        If .Rows > 3 Then
            For vllngContador = .Row + 1 To .Rows - 1
                .TextMatrix(vllngContador - 1, 15) = .TextMatrix(vllngContador, 15)
                For vllngContador1 = 0 To .Cols - 1
                    .TextMatrix(vllngContador - 1, vllngContador1) = .TextMatrix(vllngContador, vllngContador1)
                Next vllngContador1
            Next vllngContador
        Else
            If .Rows = 3 Then
                .TextMatrix(2, 15) = -1
                For vllngContador = 0 To .Cols - 1
                    .TextMatrix(2, vllngContador) = ""
                Next vllngContador
            End If
        End If
        If .Rows > 3 Then
            .Rows = .Rows - 1
            .Row = .Rows - 1
        End If
    End With
    Set fmskBorrarRegVSFlex = grdNombre

End Function

Private Function fmskBorrarRegMSFlex(vllngRenglon As Long, grdNombre As MSHFlexGrid) As MSHFlexGrid
    Dim vllngContador As Long, vllngContador1 As Long
                    
    With grdNombre
        If .Rows > 2 Then
            For vllngContador = .Row + 1 To .Rows - 1
                .RowData(vllngContador - 1) = .RowData(vllngContador)
                For vllngContador1 = 0 To .Cols - 1
                    .TextMatrix(vllngContador - 1, vllngContador1) = .TextMatrix(vllngContador, vllngContador1)
                Next vllngContador1
            Next vllngContador
        Else
            If .Rows = 2 Then
                .RowData(1) = -1
                For vllngContador = 0 To .Cols - 1
                    .TextMatrix(1, vllngContador) = ""
                Next vllngContador
            End If
        End If
        If .Rows > 2 Then
            .Rows = .Rows - 1
            .Row = .Rows - 1
        End If
    End With
    Set fmskBorrarRegMSFlex = grdNombre

End Function

Private Sub grdDetalleHonorarios_DblClick()
     cmdBorrarHonorario_Click
End Sub

Private Sub grdGrupos_Click()
    If grdGrupos.MouseCol = 3 And grdGrupos.RowData(grdGrupos.Row) <> -1 Then  'Columna que puede ser editada
        Call pEditarConcepto(13, cboMovConceptoFactura, grdGrupos)
    End If
End Sub

Private Sub grdGrupos_Scroll()
    If cboMovConceptoFactura.Visible Then
        If grdGrupos.Col = 3 Then
            If cboMovConceptoFactura.Left <> grdGrupos.Left + grdGrupos.CellLeft Or _
                cboMovConceptoFactura.Top <> grdGrupos.Top + grdGrupos.CellTop Or _
                cboMovConceptoFactura.Width <> grdGrupos.CellWidth - 8 Or _
                cboMovConceptoFactura.Height <> grdGrupos.CellHeight - 8 Then
                cboMovConceptoFactura.Move grdGrupos.Left + grdGrupos.CellLeft - 18, grdGrupos.Top + grdGrupos.CellTop - 32, grdGrupos.CellWidth + 6
                Call SendMessage(cboMovConceptoFactura.hwnd, CB_SETITEMHEIGHT, -1&, ByVal 13)
            End If
        End If
    End If
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdHBusqueda_DblClick
End Sub

Private Sub grdPaquete_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    txtCantidad.Text = ""
    txtMontoLimite.Text = ""
    txtCantidad.Visible = False
    txtMontoLimite.Visible = False
End Sub

Private Sub grdPaquete_Click()
    If grdPaquete.Col = 4 And grdPaquete.TextMatrix(grdPaquete.Row, 0) = "GC" And grdPaquete.Row > 1 Then
        Call pEditarColumnaMonto(13, txtMontoLimite, grdPaquete, True)
    End If
    If grdPaquete.Col = 2 And Trim(grdPaquete.TextMatrix(grdPaquete.Row, 0)) <> "" And grdPaquete.Row > 1 Then
        txtCantidad.Text = ""
        Call pEditarColumnaCantidad(13, txtCantidad, grdPaquete, True)
    End If
End Sub

Private Sub grdPaquete_dblClick()
    Dim vldblIvaCorresponde As Double
    Dim vldblIVA As Double
    Dim vldblIvaConvCorresponde As Double
    Dim vldblIvaConv As Double
    Dim vldblPrecio As Double
    Dim vldblPrecioConvenio As Double
    Dim vldblCantidad As Double
    Dim vllngContador As Long
    Dim vllngContGC As Long
    Dim vllngPorcDescuento As Double
    Dim vllngPorcDescuentoConv As Double
    Dim vllngGridRows As Long
    Dim vllngCantidadGrid As Long
    
    vllngGridRows = grdPaquete.Rows
    With grdPaquete
        If .Row <> -1 Then
            .Redraw = False
            
            If Val(.TextMatrix(.Row, 2)) > 1 Then
                vllngCantidadGrid = Val(.TextMatrix(.Row, 2))
                vldblIVA = Val(Format(.TextMatrix(.Row, 11), "############.00"))
                vldblIvaCorresponde = vldblIVA / Val(.TextMatrix(.Row, 2))
                
                vldblIvaConv = Val(Format(.TextMatrix(.Row, 22), "############.00"))
                vldblIvaConvCorresponde = vldblIvaConv / Val(.TextMatrix(.Row, 2))
                
                vldblPrecio = Val(Format(.TextMatrix(.Row, 7), "############.00"))
                vldblPrecioConvenio = Val(Format(.TextMatrix(.Row, 18), "############.00"))
                vldblCantidad = Val(.TextMatrix(.Row, 2)) - 1
                If Val(Format(.TextMatrix(.Row, 9), "")) = 0 Then
                    vllngPorcDescuento = 0
                Else
                    vllngPorcDescuento = Val(Format(.TextMatrix(.Row, 9), "")) / (vldblPrecio * Val(Format(.TextMatrix(.Row, 2), "")))
                End If
                
                If Val(Format(.TextMatrix(.Row, 20), "")) = 0 Then
                    vllngPorcDescuentoConv = 0
                Else
                    vllngPorcDescuentoConv = Val(Format(.TextMatrix(.Row, 20), "")) / (vldblPrecioConvenio * Val(Format(.TextMatrix(.Row, 2), "")))
                End If
                
                .TextMatrix(.Row, 2) = vldblCantidad
                .TextMatrix(.Row, 8) = FormatCurrency(vldblCantidad * vldblPrecio, 2)
                .TextMatrix(.Row, 9) = FormatCurrency((vldblCantidad * vldblPrecio) * vllngPorcDescuento, 2)
                .TextMatrix(.Row, 10) = FormatCurrency((vldblCantidad * vldblPrecio) - Val(Format(.TextMatrix(.Row, 9), "")), 2)
                .TextMatrix(.Row, 11) = FormatCurrency(vldblIvaCorresponde * vldblCantidad, 2)
                .TextMatrix(.Row, 12) = FormatCurrency(((vldblCantidad * vldblPrecio) - Val(Format(.TextMatrix(.Row, 9), ""))) + (vldblIvaCorresponde * vldblCantidad), 2)
                .TextMatrix(.Row, 6) = FormatCurrency(.TextMatrix(.Row, 5) * Val(Format(.TextMatrix(.Row, 2), "")), 2)
                .TextMatrix(.Row, 17) = FormatCurrency(Val(Format(.TextMatrix(.Row, 16), "")) * Val(Format(.TextMatrix(.Row, 2), "")), 2)
                .TextMatrix(.Row, 19) = FormatCurrency(vldblCantidad * vldblPrecioConvenio, 2)
                .TextMatrix(.Row, 20) = FormatCurrency((vldblCantidad * vldblPrecioConvenio) * vllngPorcDescuentoConv, 2)
                .TextMatrix(.Row, 21) = FormatCurrency((vldblCantidad * vldblPrecioConvenio) - Val(Format(.TextMatrix(.Row, 20), "")), 2)
                .TextMatrix(.Row, 22) = FormatCurrency(vldblIvaConvCorresponde * vldblCantidad, 2)
                .TextMatrix(.Row, 23) = FormatCurrency(((vldblCantidad * vldblPrecioConvenio) - Val(Format(.TextMatrix(.Row, 20), ""))) + (vldblIvaConvCorresponde * vldblCantidad), 2)
                txtCantidad.Visible = False
            Else
                If .TextMatrix(.Row, 0) = "GC" Then
                    For vllngContGC = 1 To grdGrupos.Rows - 1
                        grdGrupos.Row = vllngContGC
                        If grdGrupos.RowData(grdGrupos.Row) = .TextMatrix(.Row, 15) Then
                            grdGrupos = fmskBorrarRegMSFlex(grdGrupos.Row, grdGrupos)
                            txtCantidad.Visible = False
                            Exit For
                        End If
                    Next
                End If
                grdPaquete = fmskBorrarRegVSFlex(.Row, grdPaquete)
                txtCantidad.Visible = False
                pMuestraColumnasConv
            End If
            pCalculaTotales
            .Redraw = True
            .Refresh
            If fPaqueteEnPrecioPorCargos(Val(txtCvePaquete.Text), False) Then vlblnAgregaNuevo = True
            If vllngGridRows = grdPaquete.Rows And grdPaquete.TextMatrix(2, 1) <> "" And vllngCantidadGrid = vldblCantidad Then
                cmdExportar.Enabled = True
                cmdImportar.Enabled = True
            Else
                cmdExportar.Enabled = False
                cmdImportar.Enabled = False
            End If
        End If
    End With
    
End Sub

Private Function FintBuscaEnRowData(grdHBusca As VSFlexGrid, vlintCriterio As Long, vlstrTipoElemento As String)
    Dim vlintcontador As Long
    FintBuscaEnRowData = -1
    If vlstrTipoElemento = "GE" And vlintCriterio < 0 Then
        vlintCriterio = vlintCriterio * -1
    End If
    With grdHBusca
    For vlintcontador = 2 To .Rows - 1
        If IIf(.TextMatrix(vlintcontador, 15) = "", 0, .TextMatrix(vlintcontador, 15)) = vlintCriterio And vlstrTipoElemento = .TextMatrix(vlintcontador, 0) Then
            FintBuscaEnRowData = vlintcontador
            Exit For
        End If
    Next
    End With
End Function

Private Sub grdPaquete_GotFocus()
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        
        'cmdGrabarRegistro.Enabled = True
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub grdPaquete_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If grdPaquete.Row > 1 Then
        If grdPaquete.Col = 2 Then
            If KeyCode = vbKeyF2 Then 'para que se edite el contenido de la celda como en excel
                Call pEditarColumnaCantidad(13, txtCantidad, grdPaquete, False)
            End If
        Else
            If grdPaquete.Col = 4 Then
                If KeyCode = vbKeyF2 Then 'para que se edite el contenido de la celda como en excel
                    Call pEditarColumnaMonto(13, txtMontoLimite, grdPaquete, False)
                End If
            Else
                If KeyCode = vbKeyReturn Then
                    With grdPaquete
                        .Col = 0
                        .Col = 1
                        If .Row - 1 < .Rows Then
                            If .Row = .Rows - 1 Then
                                .Row = 1
                            Else
                                .Row = .Row + 1
                                .Row = IIf(.Row = .Rows - 1, 1, .Row + 1)
                            End If
                        End If
                    End With
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaquete_KeyDown"))
    Unload Me
End Sub

Private Sub lstDepartamentos_DblClick()
    pAsigna True
End Sub

Private Sub lstDepartamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna True
End Sub

Private Sub lstDepartamentosSel_DblClick()
    pAsigna False
End Sub

Private Sub lstDepartamentosSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna False
End Sub

Private Sub lstElementos_DblClick()
    If fblnSeleccionoElemento() Then pSeleccionaElemento
End Sub

Private Sub lstElementos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If fblnSeleccionoElemento() Then
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
            
            pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
            'cmdGrabarRegistro.Enabled = True
            SSTObj.TabEnabled(1) = True
            SSTObj.TabEnabled(2) = True
                
            SSTObj.TabEnabled(3) = True
            SSTObj.TabEnabled(4) = False
            pSeleccionaElemento
        End If
    End If
End Sub

Private Sub MskFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaFin_Click()
    pSelMkTexto mskFechaFin
End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtFolioInicial.SetFocus
End Sub


Private Sub mskFechaInicio_GotFocus()
    pSelMkTexto mskFechaInicio
End Sub


Private Sub mskFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mskFechaFin.SetFocus
End Sub


Private Sub optClave_GotFocus()
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub optDescripcion_GotFocus()
    If SSTObj.TabEnabled(1) Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False
    End If
End Sub

Private Sub OptPolitica_Click(Index As Integer)
    If Not cmdActualizar.Enabled Then pRecalculaPrecios
End Sub

Private Sub OptPolitica_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    If chkprecio.Value Then pRecalculaPrecios
End Sub

Private Sub optTipoPaciente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If optTipoPacienteDesc(0).Value Then
            optTipoPacienteDesc(0).SetFocus
        Else
            If optTipoPacienteDesc(1).Value Then
                optTipoPacienteDesc(1).SetFocus
            Else
                If optTipoPacienteDesc(2).Value Then
                    optTipoPacienteDesc(2).SetFocus
                Else
                    optTipoPacienteDesc(3).SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub optTipoPacienteDesc_Click(Index As Integer)
    If chkprecio.Value Then pRecalculaPrecios
End Sub

Private Sub optTipoPacienteDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub sstElementos_Click(PreviousTab As Integer)
    pEnfocaTextBox txtSeleArticulo
    pLimpiaBusqueda
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
    vllngDesktop = (Me.Top * 2) + Me.Height
    
    If SSTObj.Tab = 4 Then
        Me.Height = vllngSizeNormal
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
    
        grdHBusqueda.Enabled = True
        pLlenaGrid
        grdHBusqueda.SetFocus
        FreTotales.Visible = False
        ChkValidaPaquete.Visible = False
    ElseIf SSTObj.Tab = 0 Then
        Me.Height = vllngSizeNormal
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
        
        If vgstrEstadoManto = "" Then
            txtCvePaquete.SetFocus
        Else
            txtDescripcion.SetFocus
        End If
        FreTotales.Visible = True
        ChkValidaPaquete.Visible = True
    ElseIf SSTObj.Tab = 2 Then
        Me.Height = vllngSizeNormal
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
        FreTotales.Visible = False
        
        If lstDepartamentos.ListCount <> 0 Then
            If lstDepartamentos.ListIndex = -1 Then
                lstDepartamentos.ListIndex = 0
            End If
        End If
        
        If lstDepartamentosSel.ListCount <> 0 Then
            If lstDepartamentosSel.ListIndex = -1 Then
                lstDepartamentosSel.ListIndex = 0
            End If
        End If
        
        If lstDepartamentosSel.ListCount > 0 Then
            lstDepartamentosSel.SetFocus
        Else
            lstDepartamentos.SetFocus
        End If
        
        pHabilitaBotonesDeptos
    ElseIf SSTObj.Tab = 3 Then
    ''Honorarios
        Me.Height = vllngSizeHonorarios
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
        FreTotales.Visible = False
        cboFuncion.SetFocus
    ElseIf SSTObj.Tab = 5 Then
        Me.Height = vllngSizeNormal
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
    
        grdPresupuestos.Enabled = True
        pConfiguraGridPresupuestos
        'pLlenaGridPresupuestos
        mskFechaInicio.SetFocus
        FreTotales.Visible = False
        ChkValidaPaquete.Visible = False
    Else
    
        Me.Height = vllngSizeGrande
        Me.Top = Int((vllngDesktop - Me.Height) / 2)
        FreTotales.Visible = False
        
        If FrmMtoPaquetes.Enabled Then
            txtSeleArticulo.SetFocus
        End If
    End If
End Sub

Private Sub txtAnticipo_GotFocus()
    txtAnticipo.Text = Format(txtAnticipo.Text, "#########.##")
    pSelTextBox txtAnticipo
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
    
    pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
    'cmdGrabarRegistro.Enabled = True
    
    SSTObj.TabEnabled(1) = True
    SSTObj.TabEnabled(2) = True
    
    SSTObj.TabEnabled(3) = True
    SSTObj.TabEnabled(4) = False
End Sub

Private Sub txtAnticipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtAnticipo_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtAnticipo, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtAnticipo_LostFocus()
    txtAnticipo.Text = FormatCurrency(str(Val(Format(txtAnticipo.Text, "#########.##"))))
End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtBusqueda_Change()
On Error GoTo NotificaError
    pLlenaGrid
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_GotFocus"))
End Sub

Private Sub txtCantidad_Click()
    If grdPaquete.Row > 1 Then
        txtCantidad_KeyDown 1, 0
    End If
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlintCantidad As Integer
    On Error GoTo NotificaError

    'Para verificar que tecla fue presionada en el textbox
    If grdPaquete.Row > 1 Then
        With grdPaquete
            Select Case KeyCode
                Case 27   'ESC
                     txtCantidad.Visible = False
                    .SetFocus
                Case 13
                    txtCantidad.Text = Format(txtCantidad.Text, "####")
                    If Val(txtCantidad.Text) > 0 Then
                        vlintCantidad = Val(.Text)
                        .Text = txtCantidad.Text
                        txtCantidad.Visible = False
                        pRecalculaImportes vlintCantidad
                        .SetFocus
                    End If
                Case 38   'Flecha para arriba
                    txtCantidad.Text = Format(txtCantidad.Text, "####")
                    If Val(txtCantidad.Text) > 0 Then
                        vlintCantidad = Val(.Text)
                        .Text = txtCantidad.Text
                        txtCantidad.Visible = False
                        pRecalculaImportes vlintCantidad
                    End If
                    .SetFocus
                    DoEvents
                    .Row = IIf(.Row > .FixedRows, .Row - 1, .Row)
                Case 40  'Flecha para abajo
                    txtCantidad.Text = Format(txtCantidad.Text, "####")
                    If Val(txtCantidad.Text) > 0 Then
                        vlintCantidad = Val(.Text)
                        .Text = txtCantidad.Text
                        txtCantidad.Visible = False
                        pRecalculaImportes vlintCantidad
                    End If
                    .SetFocus
                    DoEvents
                    .Row = IIf(.Row < .Rows - 1, .Row + 1, 1)
                Case 1
                    txtCantidad.Text = Format(txtCantidad.Text, "####")
                    If Val(txtCantidad.Text) > 0 Then
                        vlintCantidad = Val(txtCantidad.Text)
                        .Text = vlintCantidad
                        'txtCantidad.Visible = False
                        pRecalculaImportes vlintCantidad
                    End If
                Case 2
                    txtCantidad.Text = Format(txtCantidad.Text, "####")
                    If Val(txtCantidad.Text) > 0 Then
                        vlintCantidad = Val(txtCantidad.Text)
                        .Text = txtCantidad.Text
                        pRecalculaImportes vlintCantidad
                    End If
            End Select
        End With
    End If
    
    If vlintCantCargo <> vlintCantidad Then If fPaqueteEnPrecioPorCargos(Val(txtCvePaquete.Text), False) Then vlblnAgregaNuevo = True
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_KeyDown"))
    Unload Me
End Sub

Private Sub txtCantidad_KeyUp(KeyCode As Integer, Shift As Integer)
    If Val(txtCantidad) <> 0 Then
        If txtCantidad <> grdPaquete.TextMatrix(grdPaquete.Row, 30) Then
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
        End If
        txtCantidad_KeyDown 2, 0
       
    End If
End Sub

Private Sub txtCvePaquete_GotFocus()
    Call pHabilitaBotonModifica(False)
    SSTObj.TabEnabled(4) = True
    cmdBuscar.Enabled = True
    
    cmdGrabarRegistro.Enabled = False
    pNuevoRegistro
    SSTObj.Tab = 0
End Sub

Private Sub pHabilitaCampos(vlblnEstatus As Boolean)
    txtDescripcion.Enabled = vlblnEstatus
    cboConceptoFactura.Enabled = vlblnEstatus
    chkActivo.Enabled = vlblnEstatus
    ChkValidaPaquete.Enabled = vlblnEstatus
    cboTratamiento.Enabled = vlblnEstatus
    cboTipo.Enabled = vlblnEstatus
    txtAnticipo.Enabled = vlblnEstatus
    chkprecio.Enabled = vlblnEstatus
    grdGrupos.Enabled = vlblnEstatus
    lstDepartamentos.Enabled = vlblnEstatus
    cmdAsignaTodo.Enabled = vlblnEstatus
    cmdAsignaUno.Enabled = vlblnEstatus
    cmdEliminaUno.Enabled = vlblnEstatus
    cmdEliminaTodo.Enabled = vlblnEstatus
    lstDepartamentosSel.Enabled = vlblnEstatus
End Sub

Private Sub pNuevoRegistro()
    Dim vgintCont As Integer
    
    
    FreTotales.Visible = True
    
    txtCvePaquete.Text = fintSigNumRs(rsPaquetes, 0)
    txtDescripcion.Text = ""
    cboTipoPaciente.ListIndex = 0
    cboEmpresas.ListIndex = 0
    optTipoPaciente(6).Value = True
    txtSeleArticulo.Text = ""
    sstElementos.Tab = 0
    cboConceptoFactura.ListIndex = 0
    cboTratamiento.ListIndex = -1
    cboTipo.ListIndex = -1
    txtAnticipo.Text = ""
    mskFecha = "  /  /    "
    vgdtmFechaActualizacion = CDate("01/01/1900")
    OptPolitica(0).Value = True
    OptPolitica(1).Value = False
    txtSeleArticulo.Text = ""
    lstElementos.Clear
    optTipoPacienteDesc(0).Value = True
    optTipoPacienteDesc(1).Value = False
    optTipoPacienteDesc(2).Value = False
    optTipoPacienteDesc(3).Value = False
    vgTipoIngresoDescuento = "T"
    chkprecio.Value = 0
    cmdActualizar.Enabled = True
    txtSubtotal.Text = ""
    txtIva.Text = ""
    txtTotal.Text = ""
    chkActivo.Value = 1
    ChkValidaPaquete.Value = vlintValidarCargosEnPaquete
    ChkValidaPaquete.Enabled = True ''*
    vgstrEstadoManto = ""
    pHabilitaCampos (False)
    If rsPaquetes.RecordCount = 0 Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        SSTObj.TabEnabled(3) = False
    End If
    grdPaquete.Clear 'Limpiar los datos seleccionados en el grid
    grdPaquete.Rows = 3
    pConfiguraGridCargos
    
    grdPaquete.Rows = 3
    grdPaquete.TextMatrix(2, 15) = -1
    
    grdGrupos.Clear 'Limpiar los datos seleccionados en el grid
    grdGrupos.Rows = 2
    pConfiguraGridGrupos
    
    grdTotales.Clear
    grdTotales.Rows = 3
    pConfiguraGridTotales
    
    pCargarDepartamentos
    
    lstDepartamentosSel.Clear
    For vgintCont = 0 To lstDepartamentos.ListCount - 1
        lstDepartamentosSel.AddItem lstDepartamentos.List(vgintCont), lstDepartamentosSel.ListCount
        lstDepartamentosSel.ItemData(lstDepartamentosSel.newIndex) = lstDepartamentos.ItemData(vgintCont)
    Next
    lstDepartamentos.Clear
    
    pHabilitaBotonesDeptos
    
    'Limpiar honorarios medicos
    pInicializaHonorarios
    pLimpiaFechaAltaPaquete
    Call pEnfocaTextBox(txtCvePaquete)
End Sub
Private Sub pLimpiaFechaAltaPaquete()
    lblFechaAltaPaquete.Visible = False
    txtFechaAltaPaquete.Visible = False
    txtFechaAltaPaquete.Text = vbNullString
End Sub
Private Sub txtCvePaquete_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    Dim vlintNumero As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            'Buscar criterio
            If (Len(txtCvePaquete.Text) <= 0) Then
                txtCvePaquete.Text = "0"
            End If
            pHabilitaCampos (True)
            If fintSigNumRs(rsPaquetes, 0) = CLng(txtCvePaquete.Text) Then
                FrmMtoPaquetes.Enabled = True
                Frame1.Enabled = True
            
                vgstrEstadoManto = "A" 'Alta
                Call pEnfocaTextBox(txtDescripcion)
                
                pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
                'cmdGrabarRegistro.Enabled = True
                
                cmdBuscar.Enabled = False
                SSTObj.TabEnabled(1) = True     'Asignación
                SSTObj.TabEnabled(2) = True     'Departamentos
                SSTObj.TabEnabled(3) = True     'Honorarios
                SSTObj.TabEnabled(4) = False    'Búsqueda
                chkActivo.Value = 1
                chkActivo.Enabled = False
                cmdImportar.Enabled = True
                cmdExportar.Enabled = True
                
                ChkValidaPaquete.Value = vlintValidarCargosEnPaquete
            Else
                If fintLocalizaPkRs(rsPaquetes, 0, txtCvePaquete.Text) > 0 Then
                    pModificaRegistro
                    pLimpiarCantidadesConv
                    vgstrEstadoManto = "M" 'Modificacion
                    Call pEnfocaTextBox(txtDescripcion)
                    
                    If vlintCambioPrecio <> 0 Then
                        pHabilitaBotonModifica (False)
                        cmdBuscar.Enabled = False
                        
                        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
                        'cmdGrabarRegistro.Enabled = True
                        
                        SSTObj.TabEnabled(1) = True
                        SSTObj.TabEnabled(2) = True
                        SSTObj.TabEnabled(3) = True
                        SSTObj.TabEnabled(4) = False 'Búsqueda
                    Else
                        pHabilitaBotonModifica (True)
                    End If
                    
                    chkActivo.Enabled = True
                    
                    ChkValidaPaquete.Enabled = True ''*
                Else
                    Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
                    Call pEnfocaTextBox(txtCvePaquete)
                End If
            End If
    End Select
    
    vlintCambioPrecio = 0
End Sub

Private Sub grdHBusqueda_DblClick()
'' If fintLocalizaPkRs(rsCnRenglon, 0, str(grdConsulta.RowData(grdConsulta.row))) <> 0 Then
    If grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1) <> "" Then
         If fintLocalizaPkRs(rsPaquetes, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)) > 0 Then
            grdGrupos.Clear 'Limpiar los datos seleccionados en el grid
            grdGrupos.Rows = 2
            pConfiguraGridGrupos
        
            pModificaRegistro
            SSTObj.Tab = 0
            Call pEnfocaTextBox(txtDescripcion)
            pHabilitaBotonModifica (True)
            pHabilitaHonorario
        Else
            Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
            Call pEnfocaMkTexto(txtCvePaquete)
        End If
    Else
        txtCvePaquete.SetFocus
       ' Call pEnfocaMkTexto(txtCvePaquete)
        '
    End If
    
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SSTObj.Tab <> 0 Then
        SSTObj.Tab = 0
        If vgstrEstadoManto = "" Then
            txtCvePaquete.SetFocus
        Else
            txtDescripcion.SetFocus
        End If
        Cancel = True
    Else
        If vgstrEstadoManto <> "" Then
            If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                txtCvePaquete.SetFocus
                lstDepartamentosSel.Clear
                
                pHabilitaBotonesDeptos
                
                
                cmdGrabarRegistro.Enabled = False
            End If
            Cancel = True
        Else
            rsPaquetes.Close
        End If
        vgblnValida = True
    End If
End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
    ''aqui?
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    SSTObj.TabEnabled(1) = vlblnHabilita
    SSTObj.TabEnabled(2) = vlblnHabilita
    SSTObj.TabEnabled(3) = vlblnHabilita
    SSTObj.TabEnabled(4) = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
    'cmdGrabarRegistro.Enabled = Not vlblnHabilita
    cmdPrint.Enabled = vlblnHabilita
End Sub

Private Sub txtCvePaquete_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub PBorraCargosAsignados(vlintCargo As Long)
    Dim vlstrSentencia As String
    vlstrSentencia = "Delete  from PvDetallePaquete where intNumPaquete = " & CStr(vlintCargo)
    Call pEjecutaSentencia(vlstrSentencia)
End Sub

Private Sub cmdBuscar_Click()
    grdHBusqueda.Enabled = True
    SSTObj.Tab = 4
    'pLlenaGrid
    grdHBusqueda.SetFocus
End Sub

Private Sub pLlenaGrid()
    Dim vlstrSentencia As String
    Dim rsPaquetes As New ADODB.Recordset
    Dim vlintcontador As Integer
    grdHBusqueda.Clear
    
    vlstrSentencia = "SELECT intNumPaquete,chrDescripcion,chrTratamiento,chrTipo,bitActivo FROM PvPaquete  where upper(chrDescripcion) like '" & txtBusqueda.Text & "%' Order by chrDescripcion"
    Set rsPaquetes = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If rsPaquetes.RecordCount <> 0 Then
        Call pLlenarMshFGrdRs(grdHBusqueda, rsPaquetes)
        pConfiguraGrid
        If grdHBusqueda.Rows > 1 Then
        With grdHBusqueda
            For vlintcontador = 1 To .Rows - 1
                .TextMatrix(vlintcontador, 5) = IIf(.TextMatrix(vlintcontador, 5), "Activo", "Inactivo")
            Next
        End With
        End If
    Else
         grdHBusqueda.Clear
         grdHBusqueda.Rows = 2
    End If

    rsPaquetes.Close
End Sub

Private Sub cmdAnteriorRegistro_Click()
    Call pPosicionaRegRs(rsPaquetes, "A")
    pModificaRegistro
End Sub

Private Sub cmdPrimerRegistro_Click()
    Call pPosicionaRegRs(rsPaquetes, "I")
    pModificaRegistro
End Sub

Private Sub cmdSiguienteRegistro_Click()
    Call pPosicionaRegRs(rsPaquetes, "S")
    pModificaRegistro
End Sub

Private Sub cmdUltimoRegistro_Click()
    Call pPosicionaRegRs(rsPaquetes, "U")
    pModificaRegistro
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        cboConceptoFactura.SetFocus
    ElseIf txtDescripcion <> "" And fblnVerificaAlfanumerico(KeyCode) Then
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False 'Búsqueda
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub chkMedicamentos_Click()
    txtSeleArticulo.SetFocus
    txtSeleArticulo_KeyUp 7, 0
End Sub

Private Sub optClave_Click()
    pLimpiaBusqueda
End Sub

Private Sub optDescripcion_Click()
    pLimpiaBusqueda
End Sub

Private Sub txtFolioFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdBuscarPresupuestos.SetFocus
End Sub


Private Sub txtFolioFinal_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioFinal_KeyPress"))
End Sub


Private Sub txtFolioInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtFolioFinal.SetFocus
End Sub


Private Sub txtFolioInicial_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtFolioInicial_KeyPress"))
End Sub


Private Sub txtImporteHonorario_GotFocus()
    txtImporteHonorario.Text = Format(txtImporteHonorario.Text, "#########.##")
    pSelTextBox txtImporteHonorario
End Sub

Private Sub txtMontoLimite_Click()
    If grdPaquete.Row > 1 Then
        txtMontoLimite_KeyDown 1, 0
    End If
End Sub

Private Sub txtMontoLimite_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlintCantidad As Integer

    'Para verificar que tecla fue presionada en el textbox
    With grdPaquete
        Select Case KeyCode
            Case 27   'ESC
                 txtMontoLimite.Visible = False
                .SetFocus
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                .Row = IIf(.Row > .FixedRows, .Row - 1, .Row)
            Case 13
                txtMontoLimite.Text = FormatCurrency(IIf(txtMontoLimite.Text = "" Or txtMontoLimite.Text = ".", "0", txtMontoLimite.Text), 2)
                .Text = FormatCurrency(txtMontoLimite.Text, 2)
                txtMontoLimite.Visible = False
                .SetFocus
            Case 40 'Flecha para abajo
                .SetFocus
                DoEvents
                .Row = IIf(.Row < .Rows - 1, .Row + 1, 1)
            Case 1
                txtMontoLimite.Text = FormatCurrency(IIf(txtMontoLimite.Text = "", "0", txtMontoLimite.Text), 2)
                .Text = FormatCurrency(txtMontoLimite.Text, 2)
                txtMontoLimite.Visible = False
                .SetFocus
            Case 2
                .Text = FormatCurrency(txtMontoLimite.Text, 2)
        End Select
    End With
End Sub

Private Sub txtMontoLimite_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    ' Solo permite números
    If Not fblnFormatoCantidad(txtMontoLimite, KeyAscii, 2) Or KeyAscii = 7 Then KeyAscii = 7
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMontoLimite_KeyPress"))
    Unload Me
End Sub

Private Sub txtMontoLimite_KeyUp(KeyCode As Integer, Shift As Integer)
    If CDbl(IIf(txtMontoLimite.Text = "." Or txtMontoLimite.Text = "", 0, txtMontoLimite.Text)) <> 0 Then
        If txtMontoLimite.Text <> grdPaquete.TextMatrix(grdPaquete.Row, 29) Then
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
        End If
        txtMontoLimite_KeyDown 2, 0
    Else
        grdPaquete.Text = FormatCurrency(0, 2)
    End If
End Sub

Private Sub txtMontoLimite_LostFocus()
    txtMontoLimite.Visible = False
End Sub

Private Sub txtSeleArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If lstElementos.Enabled Then lstElementos.SetFocus
    End If
End Sub

Private Sub txtSeleArticulo_KeyPress(KeyAscii As Integer)
    If optClave.Value Then
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtSeleArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    Dim vlstrOtroFiltro As String
    
    vlstrOtroFiltro = " "
    vlstrSentencia = ""
    
    Select Case sstElementos.Tab
        Case 0
            vlstrSentencia = "SELECT intIDArticulo, vchNombreComercial FROM IVARTICULO"
            vlstrOtroFiltro = " AND chrCostoGasto <> 'G' AND vchEstatus = 'ACTIVO'"
            vlstrOtroFiltro = IIf(chkMedicamentos.Value = 1, vlstrOtroFiltro & " and chrCveArtMedicamen = '1'", vlstrOtroFiltro)
            PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstElementos, IIf(optDescripcion.Value, "vchNombreComercial", "chrCveArticulo"), 1000, vlstrOtroFiltro, "vchNombreComercial"
        Case 1
            vlstrSentencia = "SELECT intCveEstudio, vchNombre FROM IMESTUDIO"
            PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstElementos, IIf(optDescripcion.Value, "vchNombre", "intCveEstudio"), 1000, "AND bitStatusActivo = 1", "vchNombre, intCveEstudio"
        Case 2
            vlstrSentencia = "SELECT * FROM (SELECT intCveExamen Clave, chrNombre FROM LAEXAMEN WHERE bitEstatusActivo = 1 UNION SELECT (intCveGrupo * -1) Clave, TRIM(chrNombre) || ' (GRUPO)' chrNombre FROM LAGRUPOEXAMEN WHERE bitEstatusActivo = 1)"
            PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstElementos, IIf(optDescripcion.Value, "chrNombre", "Clave"), 1000, IIf(optDescripcion.Value, vlstrOtroFiltro, " OR Clave LIKE '-" & txtSeleArticulo & "%'"), "chrNombre"
        Case 3
            vlstrSentencia = "SELECT intCveConcepto, chrDescripcion FROM PVOTROCONCEPTO"
            PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstElementos, IIf(optDescripcion.Value, "chrDescripcion", "intCveConcepto"), 1000, " and bitEstatus = 1", "chrDescripcion"
        Case 4
            vlstrSentencia = "SELECT intCveGrupo, vchNombre FROM PVGRUPOCARGO"
            PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstElementos, IIf(optDescripcion.Value, "vchNombre", "intCveGrupo"), 1000, " AND BitActivo = 1", "vchNombre"
    End Select
End Sub

Private Sub pConfiguraGridTotales()
    With grdTotales
        .ColWidth(0) = 1510     '
        .ColWidth(1) = 0        ' Columna vacia para que funcione el merge
        .ColWidth(2) = 968      ' Total Costo
        .ColWidth(3) = 968      ' Importe
        .ColWidth(4) = 880      ' Descuento
        .ColWidth(5) = 968      ' Subtotal Precio
        .ColWidth(6) = 860      ' IVA
        .ColWidth(7) = 968      ' Total
        .ColWidth(8) = 968      ' Total Costo
        .ColWidth(9) = 968      ' Importe
        .ColWidth(10) = 880     ' Descuento
        .ColWidth(11) = 968     ' Subtotal Precio
        .ColWidth(12) = 860     ' IVA
        .ColWidth(13) = 968     ' Total
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        .ColAlignment(13) = flexAlignRightCenter
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
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
        .FixedAlignment(12) = flexAlignCenterCenter
        .FixedAlignment(13) = flexAlignCenterCenter

        .MergeCells = flexMergeRestrictRows
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(0, 1) = " "
        .TextMatrix(0, 2) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 3) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 4) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 5) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 6) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 7) = "LISTAS PREDETERMINADAS"
        .TextMatrix(0, 8) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 9) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 10) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 11) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 12) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        .TextMatrix(0, 13) = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
        
        .MergeRow(0) = True
        .MergeCol(0) = True
        
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = "Total costo"
        .TextMatrix(1, 3) = "Importe"
        .TextMatrix(1, 4) = "Descuento"
        .TextMatrix(1, 5) = "Subtotal"
        .TextMatrix(1, 6) = "IVA"
        .TextMatrix(1, 7) = "Total"
        .TextMatrix(1, 8) = "Total costo"
        .TextMatrix(1, 9) = "Importe"
        .TextMatrix(1, 10) = "Descuento"
        .TextMatrix(1, 11) = "Subtotal"
        .TextMatrix(1, 12) = "IVA"
        .TextMatrix(1, 13) = "Total"

        .TextMatrix(2, 0) = "Medicamentos"
        .Row = 2
        .Col = 0
        .CellFontBold = True
        
        .AddItem ("Artículos")
        .Row = 3
        .Col = 0
        .CellFontBold = True
        
        .AddItem ("Servicios")
        .Row = 4
        .Col = 0
        .CellFontBold = True
        
        .AddItem ("  Servicios auxiliares")
        .AddItem ("  Laboratorio")
        .AddItem ("  Otros conceptos")
        
        .AddItem ("Totales")
        .Row = 8
        .Col = 0
        .CellFontBold = True
                
    End With
End Sub

Private Sub pLimpiaBusqueda()
    lstElementos.Clear
    txtSeleArticulo.Text = ""
    
    If FrmMtoPaquetes.Enabled Then
        txtSeleArticulo.SetFocus
    End If
End Sub

Private Sub pConfiguraGridGrupos()
    With grdGrupos
        .FixedCols = 2
        .FormatString = "|Grupos de cargos||Conceptos de facturación"
        .ColWidth(0) = 100  ' Fix
        .ColWidth(1) = 4000 ' Descripción del grupo
        .ColWidth(2) = 0 ' Clave Concepto de facturacion
        .ColWidth(3) = 4190 ' Descripcion Concepto de facturacion
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignLeftCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 2) = ""
        .RowData(1) = -1
    End With
End Sub

Private Sub pAgregaGrupoConcepto(vlintClaveGrupo As Integer, vlstrDescGrupo As String, vllngConcepGrupo As Long)
' Agrega elemento al grid de cargos
    Dim vlintcontador As Integer
    Dim rs As New ADODB.Recordset
    On Error GoTo NotificaError
    
    cboMovConceptoFactura.ListIndex = 0
    With grdGrupos
        For vlintcontador = 1 To .Rows - 1
            If .RowData(vlintcontador) = vlintClaveGrupo Then Exit Sub
        Next
    
        .Row = .Rows - 1
        If .TextMatrix(.Row, 1) <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
        .RowData(.Row) = vlintClaveGrupo
        .TextMatrix(.Row, 1) = vlstrDescGrupo
        .TextMatrix(.Row, 2) = vllngConcepGrupo
        
        Set rs = frsRegresaRs("SELECT Trim(chrDescripcion) Nombre FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = " & vllngConcepGrupo, adLockReadOnly, adOpenForwardOnly)
        If Not rs.EOF Then
            .TextMatrix(.Row, 3) = IIf(rs.RecordCount = 0, "", rs!Nombre)
        Else
            .TextMatrix(.Row, 3) = ""
        End If
        rs.Close
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregaGrupoConcepto"))
End Sub

Public Sub pEditarColumnaCantidad(KeyAscii As Integer, txtEdit As TextBox, grid As VSFlexGrid, vlSelText As Boolean)
    On Error GoTo NotificaError
    Dim vlintTexto As Integer
    
    With txtEdit
        '.Text = grid 'Inicialización del Textbox
       
        Select Case KeyAscii
'            Case 0 To 32 'Edita el texto de la celda en la que está posicionado
'                .SelStart = 0
'                .SelLength = 1000
            Case 8, 48 To 57 ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
'            Case 46 ' Reemplaza el texto actual solo si se teclean números
'                .Text = "."
'                .SelStart = 1
        End Select
    End With
    
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft + 32, .Top + .CellTop + 37, .CellWidth - 8, .CellHeight - 8
    End With
    txtEdit.Visible = True
    txtEdit.SetFocus
    
    If txtEdit.Text = "" Then txtEdit.Text = Val(grid.TextMatrix(grid.Row, 2))
    If vlSelText Then pSelTextBox txtEdit

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumnaCantidad"))
    Unload Me
End Sub

Private Sub grdPaquete_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If grdPaquete.Col = 2 And grdPaquete.Text <> "" And grdPaquete.Row > 1 Then 'Columna que puede ser editada
        txtCantidad.Text = ""
        Call pEditarColumnaCantidad(KeyAscii, txtCantidad, grdPaquete, False)
    End If
    
    If grdPaquete.Col = 4 And grdPaquete.TextMatrix(grdPaquete.Row, 0) = "GC" And grdPaquete.Row > 1 Then 'Columna que puede ser editada
        txtMontoLimite.Text = ""
        Call pEditarColumnaMonto(KeyAscii, txtMontoLimite, grdPaquete, False)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdPaquete_KeyPress_KeyPress"))
    Unload Me
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    ' Solo permite números
    If KeyAscii = 46 Then
        KeyAscii = 7
    Else
        If Not fblnFormatoCantidad(txtCantidad, KeyAscii, 0) Or KeyAscii = 7 Then
            KeyAscii = 7
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCantidad_KeyPress"))
    Unload Me
End Sub

Private Sub txtCantidad_LostFocus()
    txtCantidad.Visible = False
End Sub

Private Sub pRecalculaImportes(vlintCantidad As Integer)
    Dim vlstrCualLista As String
    Dim vlstrSentencia As String
    Dim vldblPrecio As Double
    Dim vllngPorcDescuento As Double
    Dim vllngPorcDescuentoConv As Double
    Dim vldblIVA As Integer
    Dim rs As New ADODB.Recordset
    
    With grdPaquete
        vlstrCualLista = .TextMatrix(.Row, 0)
        Select Case .TextMatrix(.Row, 0)
            Case "AR"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & Trim(str(.TextMatrix(.Row, 15))) & ")"
            Case "ES"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & Trim(str(.TextMatrix(.Row, 15))) & ")"
            Case "EX"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & Trim(str(.TextMatrix(.Row, 15))) & ")"
            Case "GE"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & Trim(str(.TextMatrix(.Row, 15))) & ")"
            Case "OC"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & Trim(str(.TextMatrix(.Row, 15))) & ")"
            Case "GC"
                vlstrSentencia = ""
        End Select

        If vlstrSentencia <> "" Then
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            vldblIVA = IIf(rs.RecordCount = 0, 0, rs!IVA)
            rs.Close
        Else
            If Val(Format(.TextMatrix(.Row, 10), "")) <> 0 Then
                vldblIVA = CDbl(Val(Format(.TextMatrix(.Row, 11), ""))) / CDbl(Val(Format(.TextMatrix(.Row, 10), ""))) * 100
            Else
                vldblIVA = 0
            End If
        End If
        
        If Val(Format(.TextMatrix(.Row, 9), "")) > 0 Then
            vllngPorcDescuento = CDbl(Val(Format(.TextMatrix(.Row, 9), ""))) / CDbl(Val(Format(.TextMatrix(.Row, 8), "")))
        Else
            vllngPorcDescuento = 0
        End If
        
        If Val(Format(.TextMatrix(.Row, 20), "")) = 0 Then
            vllngPorcDescuentoConv = 0
        Else
            If .TextMatrix(.Row, 20) = "" Then
                vllngPorcDescuentoConv = CDbl(Format(0, "0.00"))
            Else
                vllngPorcDescuentoConv = Val(Format(.TextMatrix(.Row, 20), "")) / Val(Format(.TextMatrix(.Row, 19), ""))
            End If
        End If

        .TextMatrix(.Row, 8) = FormatCurrency(Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 7), "############.00")), 2)
        .TextMatrix(.Row, 9) = FormatCurrency((Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 7), "############.00"))) * vllngPorcDescuento, 2)
        .TextMatrix(.Row, 6) = FormatCurrency(Val(Format(.TextMatrix(.Row, 5), "############.00")) * Val(Format(.TextMatrix(.Row, 2), "############.00")), 2)
        .TextMatrix(.Row, 17) = FormatCurrency(Val(Format(.TextMatrix(.Row, 16), "############.00")) * Val(Format(.TextMatrix(.Row, 2), "")), 2)
        .TextMatrix(.Row, 10) = FormatCurrency(Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 7), "############.00")) - Val(Format(.TextMatrix(.Row, 9), "############.00")), 2)
        .TextMatrix(.Row, 11) = FormatCurrency(Val(Format(.TextMatrix(.Row, 10), "############.00")) * vldblIVA / 100, 2)
        .TextMatrix(.Row, 12) = FormatCurrency(Val(Format(.TextMatrix(.Row, 10), "############.00")) + (Val(Format(.TextMatrix(.Row, 10), "############.00")) * vldblIVA / 100), 2)
        .TextMatrix(.Row, 19) = FormatCurrency(Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 18), "############.00")), 2)
        .TextMatrix(.Row, 20) = FormatCurrency((Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 18), "############.00"))) * vllngPorcDescuentoConv, 2)
        .TextMatrix(.Row, 21) = FormatCurrency(Val(Format(.TextMatrix(.Row, 2), "")) * Val(Format(.TextMatrix(.Row, 18), "############.00")) - Val(Format(.TextMatrix(.Row, 20), "############.00")), 2)
        .TextMatrix(.Row, 22) = FormatCurrency((Val(Format(.TextMatrix(.Row, 21), "############.00")) * vldblIVA / 100), 2)
        .TextMatrix(.Row, 23) = FormatCurrency(Val(Format(.TextMatrix(.Row, 21), "############.00")) + (Val(Format(.TextMatrix(.Row, 21), "############.00")) * vldblIVA / 100), 2)
        
        pCalculaTotales
    End With
    
End Sub

Public Sub pEditarConcepto(KeyAscii As Integer, CboBox As ComboBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    Dim vlintTexto As Integer
    
    ' Muestra el Combo en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        CboBox.Move .Left + .CellLeft - 18, .Top + .CellTop - 32, .CellWidth + 6
        CboBox.ListIndex = fintLocalizaCritCbo(CboBox, .TextMatrix(.Row, 3))
        If CboBox.ListIndex = -1 Then CboBox.ListIndex = 0
    End With
    CboBox.Visible = True
    CboBox.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarConcepto"))
    Unload Me
End Sub

Private Function flngBuscaConcepto(vlintClaveGrupo As Long) As Long
' Funcion que busca el grupo en el grid y regresa su numero de concepto de facturacion
    Dim vlintcontador As Integer
    
    flngBuscaConcepto = 0
    
    With grdGrupos
        For vlintcontador = 1 To .Rows - 1
            flngBuscaConcepto = IIf(.RowData(vlintcontador) = vlintClaveGrupo, .TextMatrix(vlintcontador, 2), flngBuscaConcepto)
        Next
    End With

End Function

Private Sub pRecalculaPrecios()
    Dim vlintIndex As Integer
    Dim vldblPrecio As Double
    Dim vldblPrecioConvenio As Double
    Dim vldblDescuento As Double
    Dim vldblDescuentoConvenio As Double
    Dim vldblIVA As Double
    Dim vldblIvaConv As Double
    Dim vldblCostoPred As Double
    Dim vldblCostoConv As Double
    Dim vlstrSentencia As String
    Dim vlstrTituloConvenio As String
    Dim rsTemp As ADODB.Recordset
    Dim vlaryParametrosSalida() As String
    Dim rsContenido As ADODB.Recordset
    Dim vlstrSentenciaContenido As String
    Dim vllngContenidoGC As Long
    Dim rsDescuentoInventario As ADODB.Recordset
    Dim vlstrSentenciaDescuentoInventario As String
    Dim PrecioUnidad As Double
    
    
    
    For vlintIndex = 2 To grdPaquete.Rows - 1

        If IIf(grdPaquete.TextMatrix(vlintIndex, 15) = "", -1, grdPaquete.TextMatrix(vlintIndex, 15)) > -1 Then
            With grdPaquete
                'Precio unitario PREDETERMINADAS
                If .TextMatrix(vlintIndex, 0) <> "GC" Then
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    'vgstrParametrosSP = .TextMatrix(vlintIndex, 15) & "|" & .TextMatrix(vlintIndex, 0) & "|" & vglngTipoParticular & "|" & 0 & "|E|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    vgstrParametrosSP = .TextMatrix(vlintIndex, 15) & "|" & .TextMatrix(vlintIndex, 0) & "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|" & 0 & "|" & IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecio
                    vldblPrecio = IIf(vldblPrecio = -1, 0, vldblPrecio)
                Else
                    pElementoGrupoPred (.TextMatrix(vlintIndex, 15))
                    vldblPrecio = vgdblPrecioPredGrupo
                End If
                
                'Precio unitario CONVENIO
                If .TextMatrix(vlintIndex, 0) <> "GC" Then
                    pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                    vgstrParametrosSP = .TextMatrix(vlintIndex, 15) & _
                    "|" & .TextMatrix(vlintIndex, 0) & _
                    "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), vglngTipoParticular) & _
                    "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & _
                    "|" & IIf(cboEmpresas.ItemData(cboEmpresas.ListIndex) <> 0, IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")), "E") & _
                    "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                    pObtieneValores vlaryResultados, vldblPrecioConvenio
                    vldblPrecioConvenio = IIf(vldblPrecioConvenio = -1, 0, vldblPrecioConvenio)
                Else
                    pElementoGrupoConv (.TextMatrix(vlintIndex, 15))
                    vldblPrecioConvenio = vgdblPrecioConvGrupo
                End If
                
                'DESCUENTOS PREDETERMINADAS
                    vldblDescuento = 0
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    frsEjecuta_SP IIf(optTipoPacienteDesc(0).Value, "A", IIf(optTipoPacienteDesc(1).Value, "I", IIf(optTipoPacienteDesc(2).Value, "E", "U"))) & "|" & _
                                    IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|0|0|" & IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 0), vgstrTipoPredGrupo) & "|" & _
                                    IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 15), vglngClavePredGrupo) & "|" & _
                                    vldblPrecio & "|" & _
                                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                    0 & "|" & IIf(Val(.TextMatrix(vlintIndex, 14)) = 0, 1, Val(.TextMatrix(vlintIndex, 14))) & "|" & Val(.TextMatrix(vlintIndex, 2)) & "|" & _
                                    Val(.TextMatrix(vlintIndex, 13)), _
                                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblDescuento
                    'vglngTipoParticular & "|0|0|" & IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 0), vgstrTipoPredGrupo) & "|" & _

                'DESCUENTOS CONVENIO
                    vldblDescuentoConvenio = 0
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    frsEjecuta_SP IIf(optTipoPaciente(6).Value, "A", IIf(optTipoPaciente(5).Value, "I", IIf(optTipoPaciente(4).Value, "E", "U"))) & "|" & _
                                    IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|0|" & IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 0), vgstrTipoConvGrupo) & "|" & _
                                    IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 15), vglngClaveConvGrupo) & "|" & _
                                    vldblPrecioConvenio & "|" & _
                                    vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                    0 & "|" & IIf(Val(.TextMatrix(vlintIndex, 14)) = 0, 1, Val(.TextMatrix(vlintIndex, 14))) & "|" & Val(.TextMatrix(vlintIndex, 2)) & "|" & _
                                    Val(.TextMatrix(vlintIndex, 13)), _
                                    "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblDescuentoConvenio
                
                vlstrSentencia = ""
                Select Case .TextMatrix(vlintIndex, 0)
                    Case "AR"
                        vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join IvArticulo on IvArticulo.smiCveConceptFact = PvConceptoFacturacion.smiCveConcepto where IvArticulo.intIdArticulo = " & .TextMatrix(vlintIndex, 15)
                    Case "ES"
                        vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join imEstudio on imEstudio.smiConFact = PvConceptoFacturacion.smiCveConcepto where imEstudio.intCveEstudio = " & .TextMatrix(vlintIndex, 15)
                    Case "OC"
                        vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join PvOtroConcepto on PvOtroConcepto.smiConceptoFact = PvConceptoFacturacion.smiCveConcepto where PvOtroConcepto.intCveConcepto = " & .TextMatrix(vlintIndex, 15)
                    Case "EX"
                        vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join LaExamen on LaExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto where LaExamen.IntCveExamen = " & .TextMatrix(vlintIndex, 15)
                    Case "GE"
                        vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join LaGrupoExamen on LaGrupoExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto where LaGrupoExamen.IntCveGrupo = " & .TextMatrix(vlintIndex, 15)
                    Case "GC"
                        Select Case vgstrTipoPredGrupo
                            Case "ME"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join IvArticulo on IvArticulo.smiCveConceptFact = PvConceptoFacturacion.smiCveConcepto where IvArticulo.intIdArticulo = " & vglngClavePredGrupo
                            Case "AR"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join IvArticulo on IvArticulo.smiCveConceptFact = PvConceptoFacturacion.smiCveConcepto where IvArticulo.intIdArticulo = " & vglngClavePredGrupo
                            Case "ES"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join imEstudio on imEstudio.smiConFact = PvConceptoFacturacion.smiCveConcepto where imEstudio.intCveEstudio = " & vglngClavePredGrupo
                            Case "EX"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join LaExamen on LaExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto where LaExamen.IntCveExamen = " & vglngClavePredGrupo
                            Case "GE"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join LaGrupoExamen on LaGrupoExamen.smiConFact = PvConceptoFacturacion.smiCveConcepto where LaGrupoExamen.IntCveGrupo = " & vglngClavePredGrupo
                            Case "OC"
                                vlstrSentencia = "select PvConceptoFacturacion.smyIVA from PvConceptoFacturacion inner join PvOtroConcepto on PvOtroConcepto.smiConceptoFact = PvConceptoFacturacion.smiCveConcepto where PvOtroConcepto.intCveConcepto = " & vglngClavePredGrupo
                        End Select
                End Select
                                    
                If vlstrSentencia <> "" Then
                    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If Not rsTemp.EOF Then
                        vldblIVA = rsTemp!smyIVA
                    Else
                        vldblIVA = 0
                    End If
                    rsTemp.Close
                End If
                
                vlstrSentencia = ""
                Select Case .TextMatrix(vlintIndex, 0)
                    Case "GC"
                        pElementoGrupoConv (.TextMatrix(vlintIndex, 15))
                        Select Case vgstrTipoConvGrupo
                            Case "ME"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                            Case "AR"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                            Case "ES"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClaveConvGrupo & ")"
                            Case "EX"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClaveConvGrupo & ")"
                            Case "GE"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClaveConvGrupo & ")"
                            Case "OC"
                                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClaveConvGrupo & ")"
                        End Select
                        
                        If vlstrSentencia <> "" Then
                            Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                            vldblIvaConv = IIf(rsTemp.RecordCount = 0, 0, rsTemp!IVA)
                            rsTemp.Close
                        Else
                            vldblIvaConv = 0
                        End If
                    Case Else
                        vldblIvaConv = vldblIVA
                End Select
                
                vllngContenidoGC = 1
                If .TextMatrix(vlintIndex, 13) = 1 Then
                    vldblPrecio = vldblPrecio / .TextMatrix(vlintIndex, 14)
                    vldblPrecioConvenio = vldblPrecioConvenio / .TextMatrix(vlintIndex, 14)
                Else
                    PrecioUnidad = 0
                    If .TextMatrix(vlintIndex, 13) = 0 And .TextMatrix(vlintIndex, 0) = "GC" And vgstrTipoConvGrupo = "AR" Then 'Nada más para los artículos
                        vlstrSentenciaContenido = "Select intContenido Contenido " & _
                                                  "From ivArticulo " & _
                                                  "WHERE intIDArticulo = " & vglngClavePredGrupo
                        Set rsContenido = frsRegresaRs(vlstrSentenciaContenido, adLockReadOnly, adOpenForwardOnly)
                        If rsContenido.RecordCount > 0 Then
                            vllngContenidoGC = rsContenido!Contenido 'Este es el contenido de IVarticulo
                            
                            PrecioUnidad = 2 'Descuento por Unidad Alterna
                            If vllngContenidoGC > 1 Then
                                vlstrSentenciaDescuentoInventario = "SELECT intDescuentoInventario FROM PvDetallePaquete WHERE intNumPaquete = " & Trim(txtCvePaquete) & _
                                                                    " and intcvecargo = " & grdPaquete.TextMatrix(vlintIndex, 15)
                                Set rsDescuentoInventario = frsRegresaRs(vlstrSentenciaDescuentoInventario, adLockReadOnly, adOpenForwardOnly)
                                If rsDescuentoInventario.RecordCount > 0 Then
                                    PrecioUnidad = rsDescuentoInventario!INTDESCUENTOINVENTARIO 'Descuento por unidad Minima
                                End If
                                rsDescuentoInventario.Close
                            End If
                            If PrecioUnidad = 1 Then
                                vldblPrecio = vldblPrecio / vllngContenidoGC
                                vldblPrecioConvenio = vldblPrecioConvenio / vllngContenidoGC
                            Else
                                If PrecioUnidad = 2 Then
                                    vllngContenidoGC = 1
                                End If
                            End If
                        End If
                        rsContenido.Close
                    End If
                End If
                
               
                vldblPrecio = Round(vldblPrecio, 2)
                
                vldblCostoPred = flngCosto(IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 0), vgstrTipoPredGrupo), IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 15), vglngClavePredGrupo))
                If .TextMatrix(vlintIndex, 13) = 0 And .TextMatrix(vlintIndex, 0) = "GC" And vgstrTipoConvGrupo = "AR" Then 'Nada más para los artículos
                    .TextMatrix(vlintIndex, 5) = FormatCurrency(vldblCostoPred / IIf(Val(.TextMatrix(vlintIndex, 14)) = 0, vllngContenidoGC, Val(.TextMatrix(vlintIndex, 14))), 2)
                Else
                    .TextMatrix(vlintIndex, 5) = FormatCurrency(IIf(Val(.TextMatrix(vlintIndex, 13)) <> 1, vldblCostoPred, vldblCostoPred / IIf(Val(.TextMatrix(vlintIndex, 14)) = 0, 1, Val(.TextMatrix(vlintIndex, 14)))), 2)
                End If
                
                
                .TextMatrix(vlintIndex, 6) = FormatCurrency(.TextMatrix(vlintIndex, 5) * CInt(.TextMatrix(vlintIndex, 2)), 2)
                If .TextMatrix(vlintIndex, 7) <> vldblPrecio Then
                    vlintCambioPrecio = vlintCambioPrecio + 1
                End If
                .TextMatrix(vlintIndex, 7) = FormatCurrency(vldblPrecio, 2)
                .TextMatrix(vlintIndex, 8) = FormatCurrency(vldblPrecio * CInt(.TextMatrix(vlintIndex, 2)), 2)
                .TextMatrix(vlintIndex, 9) = FormatCurrency(vldblDescuento, 2)
                .TextMatrix(vlintIndex, 10) = FormatCurrency((vldblPrecio * CInt(.TextMatrix(vlintIndex, 2))) - vldblDescuento, 2)
                .TextMatrix(vlintIndex, 11) = FormatCurrency((.TextMatrix(vlintIndex, 10)) * vldblIVA / 100, 2)
                .TextMatrix(vlintIndex, 12) = FormatCurrency((.TextMatrix(vlintIndex, 10)) * (vldblIVA / 100 + 1), 2)
                
                If cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0 Then
                    If .TextMatrix(vlintIndex, 0) <> "GC" Then
                        .TextMatrix(vlintIndex, 16) = .TextMatrix(vlintIndex, 5)
                    Else
                        vldblCostoConv = flngCosto(IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 0), vgstrTipoConvGrupo), IIf(.TextMatrix(vlintIndex, 0) <> "GC", .TextMatrix(vlintIndex, 15), vglngClaveConvGrupo))
                        .TextMatrix(vlintIndex, 16) = FormatCurrency(IIf(Val(.TextMatrix(vlintIndex, 13)) <> 1, vldblCostoConv, vldblCostoConv / IIf(Val(.TextMatrix(vlintIndex, 14)) = 0, 1, Val(.TextMatrix(vlintIndex, 14)))), 2)
                    End If
                    .TextMatrix(vlintIndex, 17) = FormatCurrency(.TextMatrix(vlintIndex, 16) * CInt(.TextMatrix(vlintIndex, 2)), 2)
                    .TextMatrix(vlintIndex, 18) = FormatCurrency(vldblPrecioConvenio, 2)
                    .TextMatrix(vlintIndex, 19) = FormatCurrency(vldblPrecioConvenio * CInt(.TextMatrix(vlintIndex, 2)), 2)
                    .TextMatrix(vlintIndex, 20) = FormatCurrency(vldblDescuentoConvenio, 2)
                    .TextMatrix(vlintIndex, 21) = FormatCurrency((vldblPrecioConvenio * CInt(.TextMatrix(vlintIndex, 2))) - vldblDescuentoConvenio, 2)
                    .TextMatrix(vlintIndex, 22) = FormatCurrency((.TextMatrix(vlintIndex, 21)) * vldblIvaConv / 100, 2)
                    .TextMatrix(vlintIndex, 23) = FormatCurrency((.TextMatrix(vlintIndex, 21)) * (vldblIvaConv / 100 + 1), 2)
                Else
                    .TextMatrix(vlintIndex, 16) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 17) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 18) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 19) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 20) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 21) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 22) = FormatCurrency(0, 2)
                    .TextMatrix(vlintIndex, 23) = FormatCurrency(0, 2)
                End If

            End With
        End If
    Next
    vlstrTituloConvenio = IIf(Trim(cboTipoPaciente.Text) = "<NINGUNO>", " ", IIf(Trim(cboEmpresas.Text) = "<NINGUNA>", Trim(cboTipoPaciente.Text), Trim(cboEmpresas.Text)))
    grdPaquete.TextMatrix(0, 18) = vlstrTituloConvenio
    grdPaquete.TextMatrix(0, 19) = vlstrTituloConvenio
    grdPaquete.TextMatrix(0, 20) = vlstrTituloConvenio
    grdPaquete.TextMatrix(0, 21) = vlstrTituloConvenio
    grdPaquete.TextMatrix(0, 22) = vlstrTituloConvenio
    grdPaquete.TextMatrix(0, 23) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 8) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 9) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 10) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 11) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 12) = vlstrTituloConvenio
    grdTotales.TextMatrix(0, 13) = vlstrTituloConvenio
    

    pCalculaTotales
    
    vgdtmFechaActualizacion = fdtmServerFecha
    
    If vlintCambioPrecio <> 0 Then
        MsgBox "Cambiaron los precios del contenido de paquete, es necesario guardar el paquete para actualizar la información.", vbOKOnly + vbInformation, "Mensaje"
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        
        pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
        'cmdGrabarRegistro.Enabled = True
        SSTObj.TabEnabled(1) = True
        SSTObj.TabEnabled(2) = True
        SSTObj.TabEnabled(3) = True
        SSTObj.TabEnabled(4) = False 'Búsqueda
    End If
    
    
End Sub

Private Function flngCosto(vlstrTipo As String, vlIntClave As Long) As Double
' Funcion que regresa el costo base del elemento
    Dim vlstrSentencia As String
    Dim rsTemp As ADODB.Recordset

    flngCosto = 0
    If vlstrTipo = "AR" Then
        vlstrSentencia = " SELECT " & IIf(OptPolitica(0).Value, "IVARTICULOEMPRESAS.MnyCostoMasAlto", "IVARTICULOEMPRESAS.MnyCostoUltEntrada") & " Costo" & _
                         " FROM IVARTICULO" & _
                         " INNER JOIN IVARTICULOEMPRESAS ON IVARTICULO.chrcvearticulo = IVARTICULOEMPRESAS.chrcvearticulo AND IVARTICULOEMPRESAS.tnyclaveempresa = " & vgintClaveEmpresaContable & _
                         " WHERE intIDArticulo = " & vlIntClave
    Else
        vlstrSentencia = "SELECT NumCosto Costo FROM PVCOSTOCARGOS WHERE intCveCargo = " & vlIntClave & " AND chrTipo = '" & vlstrTipo & "'" & " AND intcveempresacontable = " & vgintClaveEmpresaContable
    End If
    
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If Not rsTemp.EOF Then
        flngCosto = rsTemp!costo
    End If
    
End Function

Private Sub pLimpiarCantidadesConv()
' Limpiar las columnas de importes de acuerdo al convenio
    Dim vlintIndex As Integer
    
    With grdPaquete
        For vlintIndex = 2 To .Rows - 1
            .TextMatrix(vlintIndex, 18) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 19) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 20) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 21) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 22) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 23) = FormatCurrency(0, 2)
        Next
        
        .TextMatrix(0, 18) = " "
        .TextMatrix(0, 19) = " "
        .TextMatrix(0, 20) = " "
        .TextMatrix(0, 21) = " "
        .TextMatrix(0, 22) = " "
        .TextMatrix(0, 23) = " "
    End With
    
    With grdTotales
        For vlintIndex = 2 To .Rows - 1
            .TextMatrix(vlintIndex, 8) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 9) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 10) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 11) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 12) = FormatCurrency(0, 2)
            .TextMatrix(vlintIndex, 13) = FormatCurrency(0, 2)
        Next
        
        .TextMatrix(0, 8) = " "
        .TextMatrix(0, 9) = " "
        .TextMatrix(0, 10) = " "
        .TextMatrix(0, 11) = " "
        .TextMatrix(0, 12) = " "
        .TextMatrix(0, 13) = " "
    End With
    
End Sub

Private Sub pMuestraColumnasConv()
' Limpiar las columnas de importes de acuerdo al convenio

    grdTotales.Width = IIf(cboTipoPaciente.ListIndex = 0, 7180, IIf(grdGrupos.Rows = 2 And grdGrupos.TextMatrix(1, 1) = "", 11840, 12820))
    grdTotales.Left = IIf(cboTipoPaciente.ListIndex = 0, 5740, IIf(grdGrupos.Rows = 2 And grdGrupos.TextMatrix(1, 1) = "", 1075, 100))
    
    With grdPaquete
        .ColWidth(16) = IIf(cboTipoPaciente.ListIndex = 0, 0, IIf(grdGrupos.Rows = 2 And grdGrupos.TextMatrix(1, 1) = "", 0, 863))
        .ColWidth(17) = IIf(cboTipoPaciente.ListIndex = 0, 0, IIf(grdGrupos.Rows = 2 And grdGrupos.TextMatrix(1, 1) = "", 0, 968))
        .ColWidth(18) = IIf(cboTipoPaciente.ListIndex = 0, 0, 880)
        .ColWidth(19) = IIf(cboTipoPaciente.ListIndex = 0, 0, 968)
        .ColWidth(20) = IIf(cboTipoPaciente.ListIndex = 0, 0, 880)
        .ColWidth(21) = IIf(cboTipoPaciente.ListIndex = 0, 0, 968)
        .ColWidth(22) = IIf(cboTipoPaciente.ListIndex = 0, 0, 880)
        .ColWidth(23) = IIf(cboTipoPaciente.ListIndex = 0, 0, 968)
    End With
    grdTotales.ColWidth(8) = IIf(cboTipoPaciente.ListIndex = 0, 0, IIf(grdGrupos.Rows = 2 And grdGrupos.TextMatrix(1, 1) = "", 0, 968))
End Sub

Private Sub pElementoGrupoPred(vlIntClave As Long)
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
            vllngEstaEnPaquete = FintBuscaEnRowData(grdPaquete, rsTemp!clave, IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo))
            If vllngEstaEnPaquete = -1 Then
                vldblPrecio = 0
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                'vgstrParametrosSP = rsTemp!clave & "|" & IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo) & "|" & CStr(vglngTipoParticular) & "|0|E|0|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
                vgstrParametrosSP = rsTemp!clave & "|" & IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo) & "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|0|" & IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")) & "|0|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
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
            vllngEstaEnPaquete = FintBuscaEnRowData(grdPaquete, rsTemp!clave, IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo))
            If vllngEstaEnPaquete = -1 Then
                vldblPrecio = 0
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
                vgstrParametrosSP = rsTemp!clave & "|" & IIf(rsTemp!tipo = "ME", "AR", rsTemp!tipo) & "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|" & IIf(cboEmpresas.ItemData(cboEmpresas.ListIndex) <> 0, IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")), "E") & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
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

Public Sub pEditarColumnaMonto(KeyAscii As Integer, txtEdit As TextBox, grid As VSFlexGrid, vlSelText As Boolean)
    On Error GoTo NotificaError
    Dim vlintTexto As Integer
    
    With txtEdit
        .Text = grid 'Inicialización del Textbox
       
        Select Case KeyAscii
            Case 0 To 32 'Edita el texto de la celda en la que está posicionado
                .SelStart = 0
                .SelLength = 1000
            Case 8, 48 To 57 ' Reemplaza el texto actual solo si se teclean números
                If grid <> Chr(KeyAscii) Then
                    cmdExportar.Enabled = False
                    cmdImportar.Enabled = False
                End If
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
            Case 46 ' Reemplaza el texto actual solo si se teclean números
                .Text = "."
                .SelStart = 1
        End Select
    End With
    
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft - 80, .Top + .CellTop - 235, .CellWidth - 8, .CellHeight - 8
    End With
    txtEdit.Visible = True
    txtEdit.SetFocus
    
    If txtEdit.Text = "" Or txtEdit.Text = "$0.00" Then txtEdit.Text = IIf(CDbl(grid.TextMatrix(grid.Row, 4)) > 0, CDbl(grid.TextMatrix(grid.Row, 4)), Val(grid.TextMatrix(grid.Row, 2)) * CDbl(grid.TextMatrix(grid.Row, 7)))
    If vlSelText Then pSelTextBox txtEdit
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEditarColumnaMonto"))
    Unload Me
End Sub

Private Sub pAsigna(Asigna As Boolean, Optional Todos As Boolean)
'Procedimiento que asigna o elimina areas
    Dim vgintCont As Integer
    Dim vlintSeleccion As Integer

    If Asigna Then
        If lstDepartamentos.ListCount > 0 Then
            If Todos Then
                For vgintCont = 0 To lstDepartamentos.ListCount - 1
                    lstDepartamentosSel.AddItem lstDepartamentos.List(vgintCont), lstDepartamentosSel.ListCount
                    lstDepartamentosSel.ItemData(lstDepartamentosSel.newIndex) = lstDepartamentos.ItemData(vgintCont)
                Next
                
                lstDepartamentos.Clear
                
                If lstDepartamentosSel.ListCount <> 0 Then
                    If lstDepartamentosSel.ListIndex = -1 Then
                        lstDepartamentosSel.ListIndex = 0
                    End If
                End If
            Else
                If lstDepartamentos.ListIndex = -1 Then
                    lstDepartamentos.SetFocus
                    Exit Sub
                End If
                If fValida(lstDepartamentos.ItemData(lstDepartamentos.ListIndex)) = False Then
                    lstDepartamentos.SetFocus
                    Exit Sub
                End If
                vlintSeleccion = lstDepartamentos.ListIndex
                lstDepartamentosSel.AddItem lstDepartamentos.List(lstDepartamentos.ListIndex), lstDepartamentosSel.ListCount
                lstDepartamentosSel.ItemData(lstDepartamentosSel.newIndex) = lstDepartamentos.ItemData(lstDepartamentos.ListIndex)
                lstDepartamentos.RemoveItem (lstDepartamentos.ListIndex)
                
                If lstDepartamentos.ListCount <> 0 Then
                    lstDepartamentos.ListIndex = IIf(vlintSeleccion = 0, 0, vlintSeleccion - 1)
                Else
                    If lstDepartamentosSel.ListCount <> 0 Then
                        If lstDepartamentosSel.ListIndex = -1 Then
                            lstDepartamentosSel.ListIndex = 0
                        End If
                    End If
                End If
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
            
            pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
            'cmdGrabarRegistro.Enabled = True
            
            SSTObj.TabEnabled(1) = True
            SSTObj.TabEnabled(2) = True
            SSTObj.TabEnabled(3) = True
            SSTObj.TabEnabled(4) = False 'Búsqueda
        End If
        
        pHabilitaBotonesDeptos
        
        If Todos Then
            If lstDepartamentosSel.ListCount > 0 Then
                lstDepartamentosSel.SetFocus
            Else
                lstDepartamentos.SetFocus
            End If
        Else
            If lstDepartamentos.ListCount > 0 Then
                lstDepartamentos.SetFocus
            Else
                lstDepartamentosSel.SetFocus
            End If
        End If
    Else
        If lstDepartamentosSel.ListCount > 0 Then
            If Todos Then
                For vgintCont = 0 To lstDepartamentosSel.ListCount - 1
                    lstDepartamentos.AddItem lstDepartamentosSel.List(vgintCont), lstDepartamentos.ListCount
                    lstDepartamentos.ItemData(lstDepartamentos.newIndex) = lstDepartamentosSel.ItemData(vgintCont)
                Next
                lstDepartamentosSel.Clear
                
                If lstDepartamentos.ListCount <> 0 Then
                    If lstDepartamentos.ListIndex = -1 Then
                        lstDepartamentos.ListIndex = 0
                    End If
                End If
            Else
                If lstDepartamentosSel.ListIndex = -1 Then
                    lstDepartamentosSel.SetFocus
                    Exit Sub
                End If
                vlintSeleccion = lstDepartamentosSel.ListIndex
                lstDepartamentos.AddItem lstDepartamentosSel.List(lstDepartamentosSel.ListIndex), lstDepartamentos.ListCount
                lstDepartamentos.ItemData(lstDepartamentos.newIndex) = lstDepartamentosSel.ItemData(lstDepartamentosSel.ListIndex)
                lstDepartamentosSel.RemoveItem (lstDepartamentosSel.ListIndex)
                
                If lstDepartamentosSel.ListCount <> 0 Then
                    lstDepartamentosSel.ListIndex = IIf(vlintSeleccion = 0, 0, vlintSeleccion - 1)
                Else
                    If lstDepartamentos.ListCount <> 0 Then
                        If lstDepartamentos.ListIndex = -1 Then
                            lstDepartamentos.ListIndex = 0
                        End If
                    End If
                End If
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
            
            pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
            'cmdGrabarRegistro.Enabled = True
            
            SSTObj.TabEnabled(1) = True
            SSTObj.TabEnabled(2) = True
            SSTObj.TabEnabled(3) = True
            SSTObj.TabEnabled(4) = False 'Búsqueda
        End If
        
        pHabilitaBotonesDeptos
        
        If Todos Then
            If lstDepartamentos.ListCount > 0 Then
                lstDepartamentos.SetFocus
            Else
                lstDepartamentosSel.SetFocus
            End If
        Else
            If lstDepartamentosSel.ListCount > 0 Then
                lstDepartamentosSel.SetFocus
            Else
                lstDepartamentos.SetFocus
            End If
        End If
    End If
End Sub

Private Sub pHabilitaBotonesDeptos()
    cmdAsignaUno.Enabled = IIf(lstDepartamentos.ListCount > 0, True, False)
    cmdAsignaTodo.Enabled = IIf(lstDepartamentos.ListCount > 0, True, False)
    cmdEliminaUno.Enabled = IIf(lstDepartamentosSel.ListCount > 0, True, False)
    cmdEliminaTodo.Enabled = IIf(lstDepartamentosSel.ListCount > 0, True, False)
End Sub

Private Sub pCargarDepartamentos()
    On Error GoTo NotificaError
    Dim rsDatos As New ADODB.Recordset
    
    With lstDepartamentos
        .Clear
        Set rsDatos = frsRegresaRs("SELECT smicvedepartamento CVE, RTRIM(vchdescripcion) NOMBRE FROM NODEPARTAMENTO WHERE bitestatus = 1 AND (TRIM(chrclasificacion) IN ('G','E') or (TRIM(chrclasificacion) = 'A' and bitconsignacion = 0)) ORDER BY NOMBRE", adLockReadOnly, adOpenForwardOnly)
        If rsDatos.RecordCount > 0 Then
            Do While Not rsDatos.EOF
                .AddItem rsDatos!Nombre, .ListCount
                .ItemData(.newIndex) = rsDatos!Cve
                rsDatos.MoveNext
            Loop
        End If
        rsDatos.Close
    End With
    
    pHabilitaBotonesDeptos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargarDepartamentos"))
    Unload Me
End Sub

Private Function fValida(Cve As Long) As Boolean
'Valida que el elemento no este asignado anteriormente
    Dim vgintCont As Integer
    
    fValida = True
    
    With lstDepartamentosSel
        If .ListCount > 0 Then
            For vgintCont = 0 To .ListCount - 1
                If Cve = .ItemData(vgintCont) Then
                    fValida = False
                    Call Beep
                    Exit For
                End If
            Next
        End If
    End With

End Function

Private Function fPaqueteEnPrecioPorCargos(vlCvePaquete As Long, vlblnMensaje As Boolean) As Boolean
    Dim dblPrecioPaquetePorCargo As Double
    Dim rsTotalCargosPaquete As New ADODB.Recordset
    
    fPaqueteEnPrecioPorCargos = False
    
    dblPrecioPaquetePorCargo = 0
    Set rsTotalCargosPaquete = frsRegresaRs("Select FN_PVSELTOTALPAQUETEPORCARGO(" & vlCvePaquete & ") Info From Dual", adLockReadOnly, adOpenStatic)
    If rsTotalCargosPaquete.RecordCount <> 0 Then
        dblPrecioPaquetePorCargo = Trim(IIf(IsNull(rsTotalCargosPaquete!Info), 0, rsTotalCargosPaquete!Info))
    End If
    rsTotalCargosPaquete.Close
    
    If dblPrecioPaquetePorCargo <> 0 Then
        If vlblnMensaje Then
            'No se podrán hacer cambios al paquete porque ya se encuentra configurado en la pantalla de precios de cargos en paquetes, si requiere hacer cambios al paquete primero elimine la configuración de precios de cargos en paquete.
            MsgBox SIHOMsg(1595), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        fPaqueteEnPrecioPorCargos = True
    End If

End Function

''Comienza parte de configuración de honorarios
Private Sub pInicializaHonorarios()
    Dim vlstrSentencia As String
    Dim rsParametro  As New ADODB.Recordset
    Me.Icon = frmMenuPrincipal.Icon
     
    vlstrSentencia = "select INTCVECONCEPTOHONORARIOMEDICO from PvParametro where  TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rsParametro = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsParametro.RecordCount > 0 Then
        vllngCveConceptoFactHonorarioMedico = rsParametro!intCveConceptoHonorarioMedico
    End If
    rsParametro.Close
    
    pCargarCombos
    vlintEstadoRegistro = 1
    pConfiguraGridHonorarios
    vlngRowActualizar = 0
End Sub

Private Sub pCargarCombos()
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select intcveFuncion, vchDescripcion from ExFuncionParticipanteCirugia where bitActiva <> 0 and chrEstatus='M' order by vchDescripcion")
    If Not rs.EOF Then
        pLlenarCboRs cboFuncion, rs, 0, 1
    End If
    rs.Close
    'Los otros conceptos son los relacionados al concepto de facturacion de Honorarios médicos
    Set rs = frsRegresaRs("select intCveConcepto, chrDescripcion from PvOtroConcepto where bitEstatus <> 0 and smiConceptoFact=" & vllngCveConceptoFactHonorarioMedico & " order by chrDescripcion")
    If Not rs.EOF Then
        pLlenarCboRs cboOtroConcepto, rs, 0, 1
    End If
    rs.Close
Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargarCombos"))
End Sub

Private Sub cmdAgregar_Click()
On Error GoTo NotificaError

    Dim vlblnIncrementarRow As Boolean
    vlblnIncrementarRow = True
    Dim vllngRow As Long
    
    If Not fblnDatosValidosAgregar Then
       Exit Sub
    End If
    
    If Not fnInformacionRepetida Then
            With grdDetalleHonorarios
                
                If .Rows = 2 Then
                    If Trim(.TextMatrix(1, 1)) = "" Or vlngRowActualizar > 0 Then
                        vlblnIncrementarRow = False
                    'Cuando está el primero en blanco ya no incrementa
                    End If
                End If
                If vlngRowActualizar > 0 Then
                    vlblnIncrementarRow = False
                End If
                
                If vlblnIncrementarRow Then
                    .Rows = .Rows + 1
                End If
                
                If vlngRowActualizar > 0 Then
                    vllngRow = vlngRowActualizar
                   ' vlintEstadoRegistro = 2
                Else
                    vllngRow = .Rows - 1
                   ' vlintEstadoRegistro = 1
                End If
                
                .TextMatrix(vllngRow, cintColClaveFuncion) = cboFuncion.ItemData(cboFuncion.ListIndex)
                .TextMatrix(vllngRow, cintColDescripcionFuncion) = cboFuncion.Text
                .TextMatrix(vllngRow, cintColClaveConcepto) = cboOtroConcepto.ItemData(cboOtroConcepto.ListIndex)
                .TextMatrix(vllngRow, cintColDescripcionConcepto) = cboOtroConcepto.Text
                .TextMatrix(vllngRow, cintColImporteHonorario) = txtImporteHonorario.Text
                .TextMatrix(vllngRow, cintColEstadoRegistro) = vlintEstadoRegistro
                grdDetalleHonorarios.Col = 1
            End With
            'cmdGrabarRegistro.Enabled = True
            pBloqueaboton vlstrTipoPermiso, cmdGrabarRegistro
    End If
    cboFuncion.SetFocus
    pLimpiaConfiguracion
    
Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregar_Click"))
End Sub

Private Function fblnDatosValidosAgregar() As Boolean
   On Error GoTo NotificaError
    fblnDatosValidosAgregar = True
    
    If fblnDatosValidosAgregar And Trim(cboFuncion.Text) = "" Then
        fblnDatosValidosAgregar = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboFuncion.SetFocus
    End If
    If fblnDatosValidosAgregar And cboOtroConcepto.ListIndex = -1 Then
        fblnDatosValidosAgregar = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        cboOtroConcepto.SetFocus
    End If
    If fblnDatosValidosAgregar And Trim(txtImporteHonorario.Text) = "" Then
        fblnDatosValidosAgregar = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtImporteHonorario.SetFocus
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
    
End Function

Public Function fnInformacionRepetida() As Boolean
 On Error GoTo NotificaError
 Dim contador As Integer
    If vlngRowActualizar > 0 Then
        fnInformacionRepetida = False
        Exit Function
    End If
    For contador = 1 To grdDetalleHonorarios.Rows - 1
        If Not grdDetalleHonorarios.TextMatrix(contador, cintColClaveFuncion) = "" Then
            If Val(grdDetalleHonorarios.TextMatrix(contador, cintColClaveFuncion)) = cboFuncion.ItemData(cboFuncion.ListIndex) Then
                MsgBox "La información ya existe", vbOKOnly + vbExclamation, "Mensaje"
                fnInformacionRepetida = True
                Exit Function
            End If
        End If
    Next contador
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fnInformacionRepetida"))
End Function

Private Sub pConfiguraGridHonorarios()
On Error GoTo NotificaError
    
    With grdDetalleHonorarios
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 7
        .FormatString = "||Función de cirugía||Otro concepto de cargo|Importe honorario|"
        .ColWidth(0) = 100
        .ColWidth(cintColClaveFuncion) = 0
        .ColWidth(cintColDescripcionFuncion) = 4000
        .ColWidth(cintColClaveConcepto) = 0
        .ColWidth(cintColDescripcionConcepto) = 4000
        .ColWidth(cintColImporteHonorario) = 2000
        .ColWidth(cintColEstadoRegistro) = 0
        .TextMatrix(1, cintColClaveFuncion) = ""
        .TextMatrix(1, cintColDescripcionFuncion) = ""
        .TextMatrix(1, cintColClaveConcepto) = ""
        .TextMatrix(1, cintColDescripcionConcepto) = ""
        .TextMatrix(1, cintColImporteHonorario) = ""
        .TextMatrix(1, cintColEstadoRegistro) = ""
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridHonorarios"))
End Sub

Private Sub pLimpiaConfiguracion()
    cboFuncion.ListIndex = -1
    cboOtroConcepto.ListIndex = -1
    txtImporteHonorario = ""
    vlngRowActualizar = 0
End Sub

Private Sub cmdBorrarHonorario_Click()
    On Error GoTo NotificaError
    With grdDetalleHonorarios
        
        If .Rows > 1 Then
            If .Row > 0 Then
                .RemoveItem .Row
            End If
'        Else
'            If .Row = 1 Then
'                .TextMatrix(1, cintColClaveFuncion) = ""
'                .TextMatrix(1, cintColDescripcionFuncion) = ""
'                .TextMatrix(1, cintColClaveConcepto) = ""
'                .TextMatrix(1, cintColDescripcionConcepto) = ""
'                .TextMatrix(1, cintColImporteHonorario) = ""
'                .TextMatrix(1, cintColEstadoRegistro) = ""
'            End If
        End If
    End With
    
    pLimpiaConfiguracion
    
Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrarHonorario_Click"))
End Sub

Private Sub cmdModificar_Click()
On Error GoTo NotificaError
    
     With grdDetalleHonorarios
        If .Row > 0 Then
            If .Rows > 2 Then
                If .Row > 0 Then
                    vlngRowActualizar = .Row
                End If
            Else
                If .Row = 1 And .TextMatrix(1, cintColClaveFuncion) <> "" Then
                    vlngRowActualizar = 1
                End If
            End If
       End If
    End With
    If vlngRowActualizar > 0 Then
        pCargarDatosModificar
    End If
    Exit Sub
NotificaError:
End Sub

Private Sub pCargarDatosModificar()
    cboFuncion.ListIndex = flngLocalizaCbo(cboFuncion, grdDetalleHonorarios.TextMatrix(vlngRowActualizar, cintColClaveFuncion))
    cboOtroConcepto.ListIndex = flngLocalizaCbo(cboOtroConcepto, grdDetalleHonorarios.TextMatrix(vlngRowActualizar, cintColClaveConcepto))
    txtImporteHonorario.Text = grdDetalleHonorarios.TextMatrix(vlngRowActualizar, cintColImporteHonorario)
    cboFuncion.SetFocus
End Sub

Private Sub CboFuncion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
     If KeyAscii = 13 Then
        KeyAscii = 0
        cboOtroConcepto.SetFocus
    End If
       
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":CboFuncion_KeyPress"))
End Sub

Private Sub CboOtroConcepto_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
     If KeyAscii = 13 Then
        KeyAscii = 0
        txtImporteHonorario.SetFocus
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":CboOtroConcepto_KeyPress"))
End Sub

Private Sub txtImporteHonorario_LostFocus()
On Error GoTo NotificaError
    txtImporteHonorario.Text = FormatCurrency(Val(txtImporteHonorario.Text), 2)
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtImporteHonorario_LostFocus"))
End Sub

Private Sub TxtImporteHonorario_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
     If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAgregar.SetFocus
    End If
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteHonorario)) Then
        KeyAscii = 7
    End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":TxtImporteHonorario_KeyPress"))
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
     If KeyAscii = 13 Then
        KeyAscii = 0
        cmdAgregar_Click
    End If
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdAgregar_KeyPress"))
End Sub
Private Function pRecalculaImportesExcel(vlintCantidad As Integer, vlintRow As Integer, vlstrCargo As String, vlintTipoCargo As Integer, vlstrUnidad As String, vlstrTipoCargo As String)
    Dim vlstrCualLista As String
    Dim vllngPosicion As Long
    Dim vlstrSentencia As String
    Dim vldblPrecio As Double
    Dim vldblDescuento As Double
    Dim vldblDescuentoConvenio As Double
    Dim vldblPorceDescuento As Double
    Dim vldblPorceDescuentoConv As Double
    Dim vldblPrecioConvenio As Double
    Dim vldblIVA As Integer
    Dim vldblIvaConv As Integer
    Dim vllngContenido As Long
    Dim vlstrCveArticulo As String
    Dim vlintModoDescuentoInventario As Integer
    Dim vlstrx As String, vlstrY As String
    
    Dim vldblSubtotal As Double
    Dim vldblSubtotalConvenio As Double
    Dim vldblCantidad As Double
    Dim vldblCostoPred As Double
    Dim vldblCostoConv As Double
    Dim rs As New ADODB.Recordset
    
    Dim vlintcontador As Integer
    Dim rsEnComun As New ADODB.Recordset
    Dim vlStrGrupos As String
    Dim vlstrCadenaMsj As String
    Dim vlstrPaqueteMsj As String
    Dim vlstrSentenciaConcepto As String
    Dim vlstrConceptoFacturacion As String
    
    Dim lstListas As ListBox
    Dim vlblnNuevoElemento As Boolean
    
    Dim vlaryParametrosSalida() As String
    Dim rsUnidad As New ADODB.Recordset
    Dim vlstrSentenciaNombre As String
    Dim rsNombreCargo As New ADODB.Recordset
    Dim strNombre As String
    vlblnNuevoElemento = True
    With grdPaquete
        .Row = vlintRow
        Select Case vlintTipoCargo
            Case 0
                vlstrCualLista = "AR"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE CHRCVEARTICULO = " & Trim(vlstrCargo) & ")"
                vlstrSentenciaNombre = "Select substring(vchNombreComercial,1,50) Articulo,INTIDARTICULO From ivArticulo  WHERE CHRCVEARTICULO = " & Trim(vlstrCargo)
            Case 1
                vlstrCualLista = "ES"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & Trim(vlstrCargo) & ")"
                vlstrSentenciaNombre = "SELECT substring(VCHNOMBRE,1,50) Articulo FROM IMESTUDIO WHERE intCveEstudio = " & Trim(vlstrCargo)
            Case 2
                vlstrCualLista = IIf(vlstrTipoCargo = "GE", "GE", "EX")
                If vlstrTipoCargo = "GE" Then
                    vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & Trim(vlstrCargo) & ")"
                    vlstrSentenciaNombre = "SELECT substring(CHRNOMBRE,1,50) Articulo FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & Trim(vlstrCargo)
                Else
                    vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & Trim(vlstrCargo) & ")"
                    vlstrSentenciaNombre = "SELECT substring(CHRNOMBRE,1,50) Articulo FROM LAEXAMEN WHERE intCveExamen = " & Trim(vlstrCargo)
                End If
            Case 3
                vlstrCualLista = "OC"
                vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & Trim(vlstrCargo) & ")"
                vlstrSentenciaNombre = "SELECT substring(CHRDESCRIPCION,1,50) Articulo FROM PVOTROCONCEPTO WHERE intCveConcepto = " & Trim(vlstrCargo)
            Case 4
                vlstrCualLista = "GC"
                vlstrSentenciaNombre = "SELECT substring(VCHNOMBRE,1,50) Articulo FROM PVGRUPOCARGO WHERE INTCVEGRUPO = " & Trim(vlstrCargo)
                pElementoGrupoPred (vlstrCargo)
                
                Select Case vgstrTipoPredGrupo
                    Case "ME"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiCveConceptFact smiConFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo
                    Case "AR"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiCveConceptFact smiConFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClavePredGrupo
                    Case "ES"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClavePredGrupo
                    Case "EX"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClavePredGrupo
                    Case "GE"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClavePredGrupo
                    Case "OC"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClavePredGrupo & ")"
                        vlstrSentenciaConcepto = "SELECT smiConceptoFact smiConFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClavePredGrupo
                End Select
        End Select

        If vlstrSentencia <> "" Then
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            vldblIVA = IIf(rs.RecordCount = 0, 0, rs!IVA)
            rs.Close
                        
        Else
            vldblIVA = 0
        End If
        If vlstrSentenciaConcepto <> "" Then
            Set rs = frsRegresaRs(vlstrSentenciaConcepto, adLockReadOnly, adOpenForwardOnly)
            vlstrConceptoFacturacion = IIf(rs.RecordCount = 0, 0, rs!smiConFact)
            rs.Close
                        
        Else
            vlstrConceptoFacturacion = 0
        End If
        
        If vlstrSentenciaNombre <> "" Then
            Set rsNombreCargo = frsRegresaRs(vlstrSentenciaNombre, adLockReadOnly, adOpenForwardOnly)
            If rsNombreCargo.RecordCount > 0 Then
                strNombre = rsNombreCargo!Articulo
                If vlstrCualLista = "AR" Then
                    vlstrCargo = rsNombreCargo!intIdArticulo
                End If
            End If
            rsNombreCargo.Close
                    
        Else
            strNombre = ""
        End If

        
        vlstrSentencia = ""
        Select Case vlintTipoCargo
            Case 4
                pElementoGrupoConv (vlstrCargo)
                
                Select Case vgstrTipoConvGrupo
                    Case "ME"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                    Case "AR"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiCveConceptFact FROM IVARTICULO WHERE intIDArticulo = " & vglngClaveConvGrupo & ")"
                    Case "ES"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM IMESTUDIO WHERE intCveEstudio = " & vglngClaveConvGrupo & ")"
                    Case "EX"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAEXAMEN WHERE intCveExamen = " & vglngClaveConvGrupo & ")"
                    Case "GE"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConFact FROM LAGRUPOEXAMEN WHERE intCveGrupo = " & vglngClaveConvGrupo & ")"
                    Case "OC"
                        vlstrSentencia = "SELECT ISNULL(smyIVA,0) IVA FROM PVCONCEPTOFACTURACION WHERE smiCveConcepto = (SELECT smiConceptoFact FROM PVOTROCONCEPTO WHERE intCveConcepto = " & vglngClaveConvGrupo & ")"
                End Select
                
                If vlstrSentencia <> "" Then
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    vldblIvaConv = IIf(rs.RecordCount = 0, 0, rs!IVA)
                    rs.Close
                Else
                    vldblIvaConv = 0
                End If
            Case Else
                vldblIvaConv = vldblIVA
        End Select
        If vlstrCualLista = "GC" Then
            vlStrGrupos = ""
            With grdGrupos
                For vlintcontador = 1 To .Rows - 1
                    If Trim(vlstrCargo) <> .RowData(vlintcontador) Then
                        vlStrGrupos = IIf(vlintcontador = 1, .RowData(vlintcontador), vlStrGrupos & IIf(Trim(vlStrGrupos) = "", "", ",") & .RowData(vlintcontador))
                    End If
                Next
                If vlStrGrupos <> "" Then
                    Set rsEnComun = frsRegresaRs("SELECT DGC.intCveGrupo CveGrupo, TRIM(GC.vchNombre) DescGrupo, DGC.chrTipoCargo TipoCargo, DGC.intCveCargo CveCargo " & _
                                                   ",DECODE(DGC.chrTipoCargo " & _
                                                       ",'AR',TRIM(AR.vchnombrecomercial), 'ME',TRIM(ME.vchnombrecomercial) " & _
                                                       ",'OC',TRIM(OC.chrdescripcion), 'ES',TRIM(ES.vchnombre) " & _
                                                       ",'EX',TRIM(EX.chrnombre), 'GE',TRIM(GE.chrnombre),'') DescCargo " & _
                                                   ",DECODE(DGC.chrTipoCargo " & _
                                                       ",'AR',TRIM(AR.chrCveArticulo), 'ME',TRIM(ME.chrCveArticulo) " & _
                                                       ",'OC',TRIM(OC.intCveConcepto), 'ES',TRIM(ES.intCveEstudio) " & _
                                                       ",'EX',TRIM(EX.intCveExamen), 'GE',TRIM(GE.intcvegrupo),'') CveCargoReal " & _
                                                "FROM PVDETALLEGRUPOCARGO DGC " & _
                                                   "LEFT JOIN PVGRUPOCARGO GC ON GC.intCveGrupo = DGC.intCveGrupo " & _
                                                   "LEFT OUTER JOIN IVARTICULO AR ON DGC.intCveCargo = AR.intIDArticulo AND AR.chrCveArtMedicamen <> 1 " & _
                                                   "LEFT OUTER JOIN IVARTICULO ME ON DGC.intCveCargo = ME.intIDArticulo AND ME.chrCveArtMedicamen = 1 " & _
                                                   "LEFT OUTER JOIN PVOTROCONCEPTO OC ON DGC.intCveCargo = OC.intCveConcepto " & _
                                                   "LEFT OUTER JOIN IMESTUDIO ES ON DGC.intCveCargo = ES.intCveEstudio " & _
                                                   "LEFT OUTER JOIN LAEXAMEN EX ON DGC.intCveCargo = EX.intCveExamen " & _
                                                   "LEFT OUTER JOIN LAGRUPOEXAMEN GE ON DGC.intCveCargo = GE.intcvegrupo " & _
                                                "WHERE DGC.intCveGrupo IN (" & vlStrGrupos & ") " & _
                                                   "AND (DGC.chrTipoCargo, DGC.intCveCargo) IN (SELECT chrTipoCargo, intCveCargo " & _
                                                                                               "FROM PVDETALLEGRUPOCARGO " & _
                                                                                               "WHERE intCveGrupo = " & Trim(vlstrCargo) & ") " & _
                                                "ORDER BY CveGrupo, DescGrupo, CveCargoReal, DescCargo", adLockReadOnly, adOpenForwardOnly)
                    With rsEnComun
                        If .RecordCount > 0 Then
                            vlstrCadenaMsj = ""
                            vlstrPaqueteMsj = ""
                            .MoveFirst
                            For vlintcontador = 1 To .RecordCount
                                If vlstrPaqueteMsj = !DescGrupo Then
                                    vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & "     " & !cveCargoReal & " " & !DescCargo
                                Else
                                    vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & Format(!cveGrupo, "########") & " " & !DescGrupo & Chr(13) & "     " & !cveCargoReal & " " & !DescCargo
                                End If
                                vlstrPaqueteMsj = !DescGrupo
                                .MoveNext
                            Next vlintcontador
                            MsgBox SIHOMsg(1101) & vlstrCadenaMsj, vbOKOnly + vbInformation, "Mensaje"
                            Exit Function
                        End If
                    End With
                    rsEnComun.Close
                End If
            End With
        End If
                              
        vlstrCveArticulo = ""
        vlintModoDescuentoInventario = 0
        vllngContenido = 1
        If vlstrCualLista = "AR" Or (vlstrCualLista = "GC" And (vgstrTipoPredGrupo = "AR" Or vgstrTipoPredGrupo = "ME")) Then 'Nomas para los articulos
                ' Tipo de descuento de Inventario
                vlstrSentencia = "Select intContenido Contenido, " & _
                                " substring(vchNombreComercial,1,50) Articulo, " & _
                                " ivUA.vchDescripcion UnidadAlterna, " & _
                                " ivUM.vchDescripcion UnidadMinima, " & _
                                " ivArticulo.chrCveArticulo " & _
                                    " From ivArticulo " & _
                                " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", vlstrCargo, vglngClavePredGrupo)
                
                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rs!Contenido = 1 Then
                    vlintModoDescuentoInventario = 2
                Else
                    If Trim(vlstrUnidad) = Trim(rs!UnidadAlterna) Then
                        vlintModoDescuentoInventario = 2
                    Else
                        vlintModoDescuentoInventario = 1
                    End If
                End If
                
                vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                rs.Close
        End If
        
        '-----------------------
        'Precio unitario PREDETERMINADAS
        '-----------------------
        
        vldblPrecio = FormatCurrency(IIf(.TextMatrix(.Row, 7) = "", 0, .TextMatrix(.Row, 7)), 2)
        If vlblnNuevoElemento Or chkprecio.Value Then
            If vlstrCualLista <> "GC" Then
                pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
'                vgstrParametrosSP = str(IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)) & _
'                "|" & vlstrCualLista & _
'                "|" & CStr(vglngTipoParticular) & _
'                "|" & "0" & "|E|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable

                vgstrParametrosSP = str(IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)) & _
                "|" & vlstrCualLista & _
                "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & _
                "|" & "0" & "|" & IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")) & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable

                frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
                pObtieneValores vlaryResultados, vldblPrecio
                
                If vldblPrecio = -1 Then
                    vldblPrecio = 0
                Else
                    'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                    If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                        vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                    End If
                End If
            Else
                vldblPrecio = vgdblPrecioPredGrupo
                If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                    vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                End If
            End If
        End If
            
        '-----------------------
        'Precio unitario CONVENIO
        '-----------------------
        If vlstrCualLista <> "GC" Then
            pCargaArreglo vlaryResultados, "|" & vbDouble & "||" & vbDouble
            vgstrParametrosSP = str(IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)) & _
            "|" & vlstrCualLista & _
            "|" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & _
            "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|" & IIf(cboEmpresas.ItemData(cboEmpresas.ListIndex) <> 0, IIf(optTipoPaciente(3).Value, "U", IIf(optTipoPaciente(4).Value, "E", "I")), "E") & "|" & 0 & "|" & CDate("01/01/1900") & "|" & vgintClaveEmpresaContable
            frsEjecuta_SP vgstrParametrosSP, "sp_pvselObtenerPrecio", False, , vlaryResultados
            pObtieneValores vlaryResultados, vldblPrecioConvenio
            
            If vldblPrecioConvenio = -1 Then
                vldblPrecioConvenio = 0
            Else
                'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                    vldblPrecioConvenio = vldblPrecioConvenio / CDbl(vllngContenido)
                End If
            End If
        Else
            vldblPrecioConvenio = vgdblPrecioConvGrupo
            If vlintModoDescuentoInventario = 1 Then  'Descuento Unidad MINIMA
                vldblPrecioConvenio = vldblPrecioConvenio / CDbl(vllngContenido)
            End If
        End If
        
        If vlblnNuevoElemento Or .TextMatrix(.Row, 9) = "$0.00" Or chkprecio.Value Then
            vldblPorceDescuento = 0
        Else
            vldblPorceDescuento = FormatCurrency(.TextMatrix(.Row, 9), 2) / (Val(.TextMatrix(.Row, 2)) * FormatCurrency(.TextMatrix(.Row, 7), 2))
        End If
        
        'If vlblnNuevoElemento Or .TextMatrix(.Row, 20) = "$0.00" Or chkprecio.Value Then
        vldblPorceDescuentoConv = 0
        If Val(Format(.TextMatrix(.Row, 2))) * Val(Format(.TextMatrix(.Row, 18))) <> 0 Then
        '    vldblPorceDescuentoConv = FormatCurrency(.TextMatrix(.Row, 20), 2) / (Val(.TextMatrix(.Row, 2)) * FormatCurrency(.TextMatrix(.Row, 18), 2))
            vldblPorceDescuentoConv = Val(Format(.TextMatrix(.Row, 20))) / Val(Format(.TextMatrix(.Row, 2))) * Val(Format(.TextMatrix(.Row, 18)))
        End If
        
        vldblCantidad = Val(.TextMatrix(.Row, 2))
        
        '-----------------------
        'DESCUENTOS PREDETERMINADAS
        '-----------------------
            vldblDescuento = 0
            If vlblnNuevoElemento Or chkprecio.Value Then
                pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                frsEjecuta_SP IIf(optTipoPacienteDesc(0).Value, "A", IIf(optTipoPacienteDesc(1).Value, "I", IIf(optTipoPacienteDesc(2).Value, "E", "U"))) & "|" & _
                                IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|0|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & _
                                IIf(vlstrCualLista <> "GC", str(IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)), vglngClavePredGrupo) & "|" & _
                                IIf(vlintModoDescuentoInventario <> 1, vldblPrecio, vldblPrecio * CDbl(vllngContenido)) & "|" & _
                                vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                0 & "|" & CDbl(vllngContenido) & "|" & vldblCantidad & "|" & _
                                vlintModoDescuentoInventario, _
                                "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vldblDescuento
                
                'CStr(vglngTipoParticular) & "|0|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoPredGrupo) & "|" & _

            End If
            
        '-----------------------
        'DESCUENTOS CONVENIO
        '-----------------------
            vldblDescuentoConvenio = 0
            If vlblnNuevoElemento Or chkprecio.Value Then
                pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                frsEjecuta_SP IIf(optTipoPaciente(6).Value, "A", IIf(optTipoPaciente(5).Value, "I", IIf(optTipoPaciente(4).Value, "E", "U"))) & "|" & _
                                IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) <> 0, cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex), CStr(vglngTipoParticular)) & "|" & cboEmpresas.ItemData(cboEmpresas.ListIndex) & "|0|" & IIf(vlstrCualLista <> "GC", vlstrCualLista, vgstrTipoConvGrupo) & "|" & _
                                IIf(vlstrCualLista <> "GC", str(IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)), vglngClaveConvGrupo) & "|" & _
                                IIf(vlintModoDescuentoInventario <> 1, vldblPrecioConvenio, vldblPrecioConvenio * CDbl(vllngContenido)) & "|" & _
                                vgintNumeroDepartamento & "|" & fdtmServerFecha & "|" & _
                                0 & "|" & CDbl(vllngContenido) & "|" & vldblCantidad & "|" & _
                                vlintModoDescuentoInventario, _
                                "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vldblDescuentoConvenio
            End If
            
        vldblSubtotal = (vldblCantidad * vldblPrecio) - IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuento, (vldblPrecio * vldblCantidad) * vldblPorceDescuento)
        vldblSubtotalConvenio = (vldblCantidad * vldblPrecioConvenio) - IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuentoConvenio, (vldblPrecioConvenio * vldblCantidad) * vldblPorceDescuentoConv)
                        
        vlstrSentencia = "Select intContenido Contenido, " & _
                            " substring(vchNombreComercial,1,50) Articulo, " & _
                            " ivUA.vchDescripcion UnidadAlterna, " & _
                            " ivUM.vchDescripcion UnidadMinima, " & _
                            " ivArticulo.chrCveArticulo,IvFamilia.VCHDESCRIPCION familia, IvSubFamilia.VCHDESCRIPCION subfamilia,INTIDARTICULO " & _
                                " From ivArticulo " & _
                            " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                            " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                            " inner Join IvFamilia ON IvArticulo.CHRCVEFAMILIA = IvFamilia.CHRCVEFAMILIA and IvArticulo.CHRCVEARTMEDICAMEN = IvFamilia.CHRCVEARTMEDICAMEN" & _
                            " inner Join IvSubFamilia  ON IvArticulo.CHRCVEFAMILIA = IvSubFamilia.CHRCVEFAMILIA and  IvArticulo.CHRCVESUBFAMILIA = IvSubFamilia.CHRCVESUBFAMILIA and IvSubFamilia.CHRCVEARTMEDICAMEN = ivarticulo.CHRCVEARTMEDICAMEN " & _
                            " WHERE intIDArticulo = " & IIf(vlstrCualLista = "AR", Trim(str(vlstrCargo)), vglngClavePredGrupo)
            
        Set rsUnidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
        .TextMatrix(.Row, 1) = strNombre

        .TextMatrix(.Row, 0) = vlstrCualLista
        .TextMatrix(.Row, 2) = vldblCantidad
        .TextMatrix(.Row, 30) = vldblCantidad
        
        If vlintModoDescuentoInventario = 1 And vlstrCualLista = "AR" Then
            .TextMatrix(.Row, 3) = rsUnidad!UnidadMinima
        ElseIf vlintModoDescuentoInventario = 2 And vlstrCualLista = "AR" Then
            .TextMatrix(.Row, 3) = rsUnidad!UnidadAlterna
        Else
            .TextMatrix(.Row, 3) = ""
        End If
        
        .TextMatrix(.Row, 26) = IIf(vlstrConceptoFacturacion = 0, "", vlstrConceptoFacturacion)
        If vlstrCualLista = "AR" Then
            .TextMatrix(.Row, 27) = rsUnidad!familia
            .TextMatrix(.Row, 28) = rsUnidad!subfamilia
            .TextMatrix(.Row, 31) = IIf(IsNull(rsUnidad!chrcvearticulo), " ", rsUnidad!chrcvearticulo)
            .TextMatrix(.Row, 15) = rsUnidad!intIdArticulo
        Else
            .TextMatrix(.Row, 27) = ""
            .TextMatrix(.Row, 28) = ""
            .TextMatrix(.Row, 31) = ""
            .TextMatrix(.Row, 15) = IIf(vlstrCargo < 0, vlstrCargo * -1, vlstrCargo)
        End If

        rsUnidad.Close
        
        'SKU delarticulo
        If vlstrCualLista = "AR" Then
            vlstrSentencia = "select VCHCVEEXTERNA from SIEQUIVALENCIADETALLE where INTCVEEQUIVALENCIA = 15 and VCHCVELOCAL = " & Trim(vlstrCargo)
            Set rsUnidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rsUnidad.RecordCount > 0 Then
                .TextMatrix(.Row, 25) = rsUnidad!VCHCVEEXTERNA
            Else
                .TextMatrix(.Row, 25) = ""
            End If
        End If
        
        .TextMatrix(.Row, 4) = IIf(vlstrCualLista <> "GC", "", IIf(vlblnNuevoElemento, FormatCurrency(0, 2), FormatCurrency(IIf(.TextMatrix(.Row, 4) = "", 0, .TextMatrix(.Row, 4)), 2)))
        .TextMatrix(.Row, 7) = FormatCurrency(vldblPrecio, 2)
        .TextMatrix(.Row, 8) = FormatCurrency(vldblPrecio * vldblCantidad, 2)
        .TextMatrix(.Row, 9) = FormatCurrency(IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuento, (vldblPrecio * vldblCantidad) * vldblPorceDescuento), 2)
        .TextMatrix(.Row, 10) = FormatCurrency(vldblSubtotal, 2)
        .TextMatrix(.Row, 11) = FormatCurrency(str(vldblSubtotal * vldblIVA / 100), 2)
        .TextMatrix(.Row, 12) = FormatCurrency(str(vldblSubtotal + (vldblSubtotal * vldblIVA / 100)), 2)
        .TextMatrix(.Row, 13) = vlintModoDescuentoInventario
        .TextMatrix(.Row, 14) = vllngContenido
        .TextMatrix(.Row, 29) = IIf(vlstrCualLista <> "GC", "", IIf(vlblnNuevoElemento, FormatCurrency(0, 2), FormatCurrency(IIf(.TextMatrix(.Row, 4) = "", 0, .TextMatrix(.Row, 4)), 2)))
        If vlblnNuevoElemento Or chkprecio.Value Then
            vldblCostoPred = flngCosto(IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 0), vgstrTipoPredGrupo), IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 15), vglngClavePredGrupo))
            .TextMatrix(.Row, 5) = FormatCurrency(IIf(vlintModoDescuentoInventario <> 1, vldblCostoPred, vldblCostoPred / CDbl(vllngContenido)), 2)
            If vlstrCualLista <> "GC" Then
                .TextMatrix(.Row, 16) = .TextMatrix(.Row, 5)
            Else
                vldblCostoConv = flngCosto(IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 0), vgstrTipoConvGrupo), IIf(vlstrCualLista <> "GC", .TextMatrix(.Row, 15), vglngClaveConvGrupo))
                .TextMatrix(.Row, 16) = FormatCurrency(IIf(vlintModoDescuentoInventario <> 1, vldblCostoConv, vldblCostoConv / CDbl(vllngContenido)), 2)
            End If
        End If
        .TextMatrix(.Row, 6) = FormatCurrency(.TextMatrix(.Row, 5) * CInt(.TextMatrix(.Row, 2)), 2)
        .TextMatrix(.Row, 17) = FormatCurrency(.TextMatrix(.Row, 16) * CInt(.TextMatrix(.Row, 2)), 2)
        
        .TextMatrix(.Row, 18) = FormatCurrency(vldblPrecioConvenio, 2)
        .TextMatrix(.Row, 19) = FormatCurrency(vldblPrecioConvenio * vldblCantidad, 2)
        .TextMatrix(.Row, 20) = FormatCurrency(IIf(vlblnNuevoElemento Or chkprecio.Value, vldblDescuentoConvenio, (vldblPrecioConvenio * vldblCantidad) * vldblPorceDescuentoConv), 2)
        .TextMatrix(.Row, 21) = FormatCurrency(vldblSubtotalConvenio, 2)
        .TextMatrix(.Row, 22) = FormatCurrency(str(vldblSubtotalConvenio * vldblIvaConv / 100), 2)
        .TextMatrix(.Row, 23) = FormatCurrency(str(vldblSubtotalConvenio + (vldblSubtotalConvenio * vldblIvaConv / 100)), 2)
        .TextMatrix(.Row, 24) = 0
        
        vlblnAgregaNuevo = False
        If fPaqueteEnPrecioPorCargos(Val(txtCvePaquete.Text), False) Then vlblnAgregaNuevo = vlblnNuevoElemento
        
        .Col = 0
        .Redraw = True
        .Refresh
        pCalculaTotales
        If vlstrCualLista = "GC" Then
            pAgregaGrupoConcepto CInt(vlstrCargo), strNombre, -1
        End If
        If vlblnNuevoElemento Then pMuestraColumnasConv
    End With
End Function
Private Function pRestablecerBarra()
    freBarra.Visible = False
    pgbCargando.Max = 100
    pgbCargando.Value = 0
End Function
