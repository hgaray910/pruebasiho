VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMantoDescuentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de descuentos"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freActualizando 
      Height          =   1335
      Left            =   3345
      TabIndex        =   56
      Top             =   8670
      Visible         =   0   'False
      Width           =   4560
      Begin VB.Label Label10 
         Caption         =   "Actualizando descuentos, por favor espere..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   390
         TabIndex        =   57
         Top             =   315
         Width           =   4020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   4538
      TabIndex        =   62
      Top             =   2600
      Width           =   2835
      Begin VB.CommandButton cmdVerDescuentos 
         Caption         =   "Asignar descuentos"
         Enabled         =   0   'False
         Height          =   405
         Left            =   45
         TabIndex        =   64
         ToolTipText     =   "Ver/cancelar administración de descuentos"
         Top             =   150
         Width           =   1815
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1875
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoDescuento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Guardar la información"
         Top             =   150
         Width           =   900
      End
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   2340
      Top             =   8170
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame freBarra 
      Height          =   1275
      Left            =   1620
      TabIndex        =   35
      Top             =   82220
      Visible         =   0   'False
      Width           =   8205
      Begin MSComctlLib.ProgressBar pgbCargando 
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
         Left            =   105
         TabIndex        =   37
         Top             =   150
         Width           =   7890
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
   Begin VB.Frame freDescuento 
      Height          =   1230
      Left            =   6510
      TabIndex        =   42
      Top             =   6120
      Visible         =   0   'False
      Width           =   5130
      Begin VB.OptionButton optTipoDescuento 
         Caption         =   "C&osto"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   67
         Top             =   510
         Width           =   885
      End
      Begin VB.OptionButton optTipoDescuento 
         Caption         =   "&Cantidad"
         Height          =   195
         Index           =   1
         Left            =   3180
         TabIndex        =   66
         Top             =   510
         Width           =   945
      End
      Begin VB.OptionButton optTipoDescuento 
         Caption         =   "&Porcentaje"
         Height          =   195
         Index           =   0
         Left            =   2055
         TabIndex        =   65
         Top             =   510
         Value           =   -1  'True
         Width           =   1185
      End
      Begin MSMask.MaskEdBox mskFechaInicioVigencia 
         Height          =   315
         Left            =   90
         TabIndex        =   59
         ToolTipText     =   "Fecha de inicio de la vigencia"
         Top             =   855
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CheckBox chkVigencia 
         Caption         =   "Con &vigencia"
         Height          =   250
         Left            =   90
         TabIndex        =   58
         Top             =   465
         Width           =   1245
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   2
         EndProperty
         Height          =   315
         Left            =   3420
         TabIndex        =   43
         Top             =   855
         Width           =   1365
      End
      Begin MSMask.MaskEdBox mskFechaFinVigencia 
         Height          =   315
         Left            =   1845
         TabIndex        =   60
         ToolTipText     =   "Fecha de fin de la vigencia"
         Top             =   855
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "al"
         Height          =   195
         Left            =   1605
         TabIndex        =   61
         Top             =   915
         Width           =   120
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento"
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   75
         TabIndex        =   45
         Top             =   180
         Width           =   1065
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   45
         Top             =   135
         Width           =   5040
      End
      Begin VB.Label lblPorcentaje 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   4830
         TabIndex        =   44
         Top             =   900
         Width           =   210
      End
   End
   Begin VB.Frame FreConDescuento 
      Caption         =   "Pacientes con descuentos (I)=Interno (E)=Externo"
      Height          =   2280
      Left            =   6120
      TabIndex        =   38
      Top             =   315
      Width           =   5780
      Begin VB.ListBox lstConDescuento 
         Height          =   1815
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   330
         Width           =   5655
      End
   End
   Begin VB.Frame FreElementos 
      Caption         =   "Elementos a Incluir"
      Height          =   3840
      Left            =   50
      TabIndex        =   10
      Top             =   4000
      Width           =   5400
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar información"
         Height          =   315
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Cargar la información"
         Top             =   1035
         Width           =   1680
      End
      Begin TabDlg.SSTab sstElementos 
         Height          =   3450
         Left            =   45
         TabIndex        =   12
         ToolTipText     =   "Elementos a asignar descuentos"
         Top             =   300
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   6085
         _Version        =   393216
         Tabs            =   7
         Tab             =   6
         TabsPerRow      =   4
         TabHeight       =   529
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
         TabPicture(0)   =   "frmMantoDescuento.frx":0342
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(1)=   "chkMedicamentos"
         Tab(0).Control(2)=   "optDescripcion"
         Tab(0).Control(3)=   "optClave"
         Tab(0).Control(4)=   "txtSeleArticulo"
         Tab(0).Control(5)=   "lstArticulos"
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Imagenología"
         TabPicture(1)   =   "frmMantoDescuento.frx":035E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label6"
         Tab(1).Control(1)=   "lstEstudios"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Laboratorio"
         TabPicture(2)   =   "frmMantoDescuento.frx":037A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label7"
         Tab(2).Control(1)=   "lstExamenes"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Otros"
         TabPicture(3)   =   "frmMantoDescuento.frx":0396
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label8"
         Tab(3).Control(1)=   "lstOtrosConceptos"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "Cirugías"
         TabPicture(4)   =   "frmMantoDescuento.frx":03B2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Label9"
         Tab(4).Control(1)=   "lstCirugias"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "CONCEPTOS"
         TabPicture(5)   =   "frmMantoDescuento.frx":03CE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label12"
         Tab(5).Control(1)=   "lstConceptos"
         Tab(5).ControlCount=   2
         TabCaption(6)   =   "Paquetes"
         TabPicture(6)   =   "frmMantoDescuento.frx":03EA
         Tab(6).ControlEnabled=   -1  'True
         Tab(6).Control(0)=   "Label11"
         Tab(6).Control(0).Enabled=   0   'False
         Tab(6).Control(1)=   "lstPaquetes"
         Tab(6).Control(1).Enabled=   0   'False
         Tab(6).ControlCount=   2
         Begin VB.ListBox lstPaquetes 
            Height          =   2205
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   72
            ToolTipText     =   "Listado de paquetes disponibles"
            Top             =   1080
            Width           =   4845
         End
         Begin VB.ListBox lstArticulos 
            DragIcon        =   "frmMantoDescuento.frx":0406
            Height          =   1815
            Left            =   -74865
            Sorted          =   -1  'True
            TabIndex        =   50
            ToolTipText     =   "Artículos disponibles"
            Top             =   1550
            Width           =   4845
         End
         Begin VB.TextBox txtSeleArticulo 
            Height          =   315
            Left            =   -74865
            TabIndex        =   49
            ToolTipText     =   "Teclee la clave o la descripcion del artículo"
            Top             =   1200
            Width           =   2880
         End
         Begin VB.OptionButton optClave 
            Caption         =   "Clave"
            Height          =   225
            Left            =   -71880
            TabIndex        =   48
            Top             =   1125
            Width           =   795
         End
         Begin VB.OptionButton optDescripcion 
            Caption         =   "Descripción"
            Height          =   225
            Left            =   -71880
            TabIndex        =   47
            Top             =   1320
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.CheckBox chkMedicamentos 
            Caption         =   "Sólo medicamentos"
            Height          =   225
            Left            =   -71880
            TabIndex        =   46
            Top             =   833
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.ListBox lstConceptos 
            Height          =   2205
            Left            =   -74865
            Sorted          =   -1  'True
            TabIndex        =   40
            ToolTipText     =   "Conceptos de factura"
            Top             =   1080
            Width           =   4845
         End
         Begin VB.ListBox lstEstudios 
            Height          =   2205
            Left            =   -74860
            Sorted          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Estudios de Imagenología"
            Top             =   1150
            Width           =   4845
         End
         Begin VB.ListBox lstExamenes 
            Height          =   2205
            Left            =   -74865
            Sorted          =   -1  'True
            TabIndex        =   15
            ToolTipText     =   "Exámenes de laboratorio de esa clasificación"
            Top             =   1150
            Width           =   4845
         End
         Begin VB.ListBox lstOtrosConceptos 
            Height          =   2205
            Left            =   -74860
            Sorted          =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Otros conceptos"
            Top             =   1150
            Width           =   4845
         End
         Begin VB.ListBox lstCirugias 
            Height          =   2205
            Left            =   -74860
            Sorted          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Cirugías disponibles"
            Top             =   1080
            Width           =   4845
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Paquetes disponibles"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "Catálogo de Artículos"
            Height          =   240
            Left            =   -74880
            TabIndex        =   51
            Top             =   825
            Width           =   1980
         End
         Begin VB.Label Label12 
            Caption         =   "Conceptos de facturación"
            Height          =   240
            Left            =   -74880
            TabIndex        =   41
            Top             =   840
            Width           =   1980
         End
         Begin VB.Label Label7 
            Caption         =   "Exámenes de Laboratorio"
            Height          =   240
            Left            =   -74880
            TabIndex        =   20
            Top             =   945
            Width           =   1980
         End
         Begin VB.Label Label8 
            Caption         =   "Otros Conceptos"
            Height          =   285
            Left            =   -74880
            TabIndex        =   19
            Top             =   945
            Width           =   1290
         End
         Begin VB.Label Label9 
            Caption         =   "Cirugías disponibles"
            Height          =   285
            Left            =   -74880
            TabIndex        =   18
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Estudios de Imagenología"
            Height          =   240
            Left            =   -74880
            TabIndex        =   17
            Top             =   930
            Width           =   1980
         End
      End
   End
   Begin VB.Frame freElementosIncuidos 
      Caption         =   "Descuentos a incluir"
      Height          =   3855
      Left            =   6300
      TabIndex        =   8
      Top             =   4000
      Width           =   5505
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDescuentos 
         Height          =   3450
         Left            =   45
         TabIndex        =   9
         ToolTipText     =   "Descuentos asignados"
         Top             =   300
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   6085
         _Version        =   393216
         Cols            =   5
         FixedRows       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         FormatString    =   "|Clave|Descricpion|Descuento|Tipo"
         BandDisplay     =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   5
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2100
      Left            =   5520
      TabIndex        =   5
      Top             =   4170
      Width           =   705
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   615
         Left            =   75
         MaskColor       =   &H80000005&
         Picture         =   "frmMantoDescuento.frx":0850
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Seleccionar todos"
         Top             =   1410
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Excluir"
         Height          =   615
         Index           =   1
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmMantoDescuento.frx":0B02
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Excluir de la lista"
         Top             =   780
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSelecciona 
         Caption         =   "Incluir"
         Height          =   615
         Index           =   0
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmMantoDescuento.frx":0C7C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Seleccionar"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame Frame6 
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   3210
      Width           =   11715
   End
   Begin TabDlg.SSTab SSTDescuentos 
      Height          =   8185
      Left            =   0
      TabIndex        =   21
      Top             =   -15
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   14446
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Paciente"
      TabPicture(0)   =   "frmMantoDescuento.frx":0DF6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrePaciente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tipo de paciente"
      TabPicture(1)   =   "frmMantoDescuento.frx":0E12
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "freTipoPaciente"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Convenio"
      TabPicture(2)   =   "frmMantoDescuento.frx":0E2E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FreEmpresa"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   -65500
         TabIndex        =   80
         Top             =   3400
         Width           =   2310
         Begin VB.CommandButton Command5 
            Caption         =   "Exportar"
            Height          =   405
            Left            =   45
            TabIndex        =   82
            ToolTipText     =   "Exportar"
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Importar"
            Height          =   405
            Left            =   1150
            TabIndex        =   81
            ToolTipText     =   "Importar"
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   -65500
         TabIndex        =   77
         Top             =   3400
         Width           =   2310
         Begin VB.CommandButton Command3 
            Caption         =   "Exportar"
            Height          =   405
            Left            =   45
            TabIndex        =   79
            ToolTipText     =   "Exportar"
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Importar"
            Height          =   405
            Left            =   1150
            TabIndex        =   78
            ToolTipText     =   "Importar"
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   9500
         TabIndex        =   74
         Top             =   3400
         Width           =   2310
         Begin VB.CommandButton cmdImportar 
            Caption         =   "Importar"
            Height          =   405
            Left            =   1150
            TabIndex        =   76
            ToolTipText     =   "Importar"
            Top             =   150
            Width           =   1095
         End
         Begin VB.CommandButton cmdExportar 
            Caption         =   "Exportar"
            Height          =   405
            Left            =   45
            TabIndex        =   75
            ToolTipText     =   "Exportar"
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.Frame FrePaciente 
         Height          =   2280
         Left            =   50
         TabIndex        =   28
         Top             =   330
         Width           =   6060
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "&Externo"
            Height          =   255
            Index           =   1
            Left            =   3750
            TabIndex        =   34
            Top             =   330
            Width           =   975
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "&Interno"
            Height          =   255
            Index           =   0
            Left            =   2670
            TabIndex        =   33
            Top             =   330
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtPaciente 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   645
            Width           =   4530
         End
         Begin VB.TextBox txtMovimientoPaciente 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   0
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox txtEmpresaPaciente 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   1005
            Width           =   4530
         End
         Begin VB.TextBox txtTipoPaciente 
            Height          =   315
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1365
            Width           =   2880
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   75
            TabIndex        =   32
            Top             =   330
            Width           =   1320
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   75
            TabIndex        =   31
            Top             =   705
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Left            =   75
            TabIndex        =   30
            Top             =   1065
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de paciente"
            Height          =   195
            Left            =   75
            TabIndex        =   29
            Top             =   1425
            Width           =   1200
         End
      End
      Begin VB.Frame freTipoPaciente 
         Caption         =   "Tipos de paciente"
         Height          =   2280
         Left            =   -74950
         TabIndex        =   26
         Top             =   330
         Width           =   6060
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "Urgencias"
            Height          =   195
            Index           =   3
            Left            =   4920
            TabIndex        =   70
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "Todos"
            Height          =   195
            Index           =   2
            Left            =   4920
            TabIndex        =   68
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "Externo"
            Height          =   195
            Index           =   1
            Left            =   4920
            TabIndex        =   55
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optTipoPac2 
            Caption         =   "Interno"
            Height          =   195
            Index           =   0
            Left            =   4920
            TabIndex        =   54
            Top             =   600
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.ListBox lstTiposPaciente 
            Height          =   1815
            Left            =   45
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   345
            Width           =   4845
         End
      End
      Begin VB.Frame FreEmpresa 
         Caption         =   "Empresas activas"
         Height          =   2280
         Left            =   -74950
         TabIndex        =   22
         Top             =   330
         Width           =   6060
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "Urgencias"
            Height          =   195
            Index           =   3
            Left            =   4920
            TabIndex        =   71
            ToolTipText     =   "Tipo paciente"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "Todos"
            Height          =   195
            Index           =   2
            Left            =   4920
            TabIndex        =   69
            ToolTipText     =   "Tipo paciente"
            Top             =   1680
            Width           =   975
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "Externo"
            Height          =   195
            Index           =   1
            Left            =   4920
            TabIndex        =   25
            ToolTipText     =   "Tipo paciente"
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton OptTipoPac 
            Caption         =   "Interno"
            Height          =   195
            Index           =   0
            Left            =   4920
            TabIndex        =   24
            ToolTipText     =   "Tipo paciente"
            Top             =   600
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.ComboBox cboTipoConvenio 
            Height          =   315
            Left            =   50
            Style           =   2  'Dropdown List
            TabIndex        =   52
            ToolTipText     =   "Tipos de convenio disponibles"
            Top             =   225
            Width           =   4845
         End
         Begin VB.ListBox lstEmpresas 
            Height          =   1620
            Left            =   50
            Sorted          =   -1  'True
            TabIndex        =   23
            Top             =   555
            Width           =   4830
         End
      End
   End
End
Attribute VB_Name = "frmMantoDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmMantoDescuentos
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de Descuentos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 13/Noviembre/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: 14-Nov-05
'-------------------------------------------------------------------------------------
' Fecha:            09/Enero/2004
' Autor:            Rosenda Hernández Anaya
' Descripción:      Capturar vigencia del descuento
'-------------------------------------------------------------------------------------
' Fecha:            04/Noviembre/2005
' Autor:            Rodolfo Ramos García
' Descripción:      Que elimine del query los registros que tengan que ver con conceptos de facturación
'                   que no existan, esto pasa cuando se borran conceptos de facturación
'                   sin borrar el de los descuentos.
'-------------------------------------------------------------------------------------

Option Explicit
Dim vgstrEstadoManto As String
Dim vglngDesktop As Long     'Para saber el tamaño del desktop
Const cgintFactorVentana = 200
Dim rsDescuentos As New ADODB.Recordset
Private Type Cargos
    lngClave As Long
    strDescripcion As String
End Type
Private cargosFaltantes() As Cargos 'Array para mostrar cuales cargos no estan incluidos en el excel
    
'------------------------------------------------
' Estados de vgstrEstadoManto
' ""  = Inicio
' "A" = Alta de un descuento
' "AS" = Asignando descuentos en una alta
' "MS" = Asignando descuentos en un cambio
' "M" = Consulta/Modificación de descuentos
'------------------------------------------------
Private Sub pCargaGuardados(vlstrCual As String)
    On Error GoTo NotificaError

    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If vlstrCual = "P" Then 'Pacientes
        vlstrSentencia = _
        "Select " & _
               "Distinct " & _
               "PvDescuento.chrTipoDescuento," & _
               "PvDescuento.intCveAfectada," & _
               "PvDescuento.chrTipoPaciente," & _
               "case when PvDescuento.chrTipoPaciente = 'I' then " & _
                   "rtrim(AdPaciente.vchApellidoPaterno) || ' ' || rtrim(AdPaciente.vchApellidoMaterno) || ' ' || rtrim(AdPaciente.vchNombre) " & _
               "Else " & _
                   "' ' " & _
               "end as Dato," & _
               "case when PvDescuento.chrTipoPaciente = 'E' then " & _
                    "rtrim(Externo.chrApePaterno) || ' ' || rtrim(Externo.chrApeMaterno) || ' ' || rtrim(Externo.chrNombre) " & _
               "Else " & _
                   "' ' " & _
               "end As NomExterno " & _
        "From " & _
               "PvDescuento " & _
               "left outer join RegistroExterno on PvDescuento.INTCVEAFECTADA = RegistroExterno.INTNUMCUENTA " & _
               "left outer join Externo on RegistroExterno.INTNUMPACIENTE = Externo.INTNUMPACIENTE " & _
               "left outer join AdAdmision on PvDescuento.INTCVEAFECTADA = AdAdmision.NUMNUMCUENTA " & _
               "left outer join AdPaciente on AdAdmision.NUMCVEPACIENTE = AdPaciente.NUMCVEPACIENTE " & _
        "Where " & _
               "chrTipoDescuento = 'P'" & _
               " AND pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
    ElseIf vlstrCual = "T" Then 'Tipos de Paciente
        vlstrSentencia = "SELECT DISTINCT PvDescuento.chrTipoDescuento, " & _
                "PvDescuento.intCveAfectada, PvDescuento.chrTipoPaciente, " & _
                "AdTipoPaciente.vchDescripcion as Dato " & _
                "FROM PvDescuento LEFT OUTER JOIN " & _
                "AdTipoPaciente ON " & _
                "PvDescuento.intCveAfectada = AdTipoPaciente.tnyCveTipoPaciente " & _
                "Where chrTipoDescuento = 'T' and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
    Else 'Empresas
        vlstrSentencia = "SELECT DISTINCT PvDescuento.chrTipoDescuento, " & _
                "PvDescuento.intCveAfectada, " & _
                "PvDescuento.chrTipoPaciente, " & _
                "CcEmpresa.vchDescripcion as Dato " & _
                "FROM PvDescuento LEFT OUTER JOIN " & _
                "CcEmpresa ON " & _
                "PvDescuento.intCveAfectada = CcEmpresa.intCveEmpresa " & _
                "Where chrTipoDescuento = 'E' and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
    End If
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    lstConDescuento.Clear
    Do While Not rs.EOF
        If rs!CHRTIPOPACIENTE = "E" And SSTDescuentos.Tab = 0 Then 'Solo para los externos
            lstConDescuento.AddItem RTrim(rs!NomExterno) & "  (E)"
        Else
            lstConDescuento.AddItem RTrim(rs!Dato) & IIf(rs!CHRTIPOPACIENTE = " ", "", "  (" & IIf(rs!CHRTIPOPACIENTE = "A", "T", rs!CHRTIPOPACIENTE) & ")")
        End If
        lstConDescuento.ItemData(lstConDescuento.newIndex) = rs!intCveAfectada
        rs.MoveNext
    Loop
    
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaGuardados"))
    Unload Me
End Sub

Private Sub pCargaArticulos()
    On Error GoTo NotificaError
    
    pgbCargando.Value = 0
    lblTextoBarra.Caption = SIHOMsg(280)
    cmdCargar.Enabled = False
    Dim vlstrSentencia As String
    Dim rsDatos As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim lstListas As ListBox
    
    Select Case sstElementos.Tab
        Case 0
            Set lstListas = lstArticulos
            vlstrSentencia = "SELECT intIDArticulo, vchNombreComercial FROM IvArticulo where vchEstatus='ACTIVO' order by vchNombreComercial"
            lstListas.Visible = False
        Case 1
            Set lstListas = lstEstudios
            vlstrSentencia = "SELECT intCveEstudio, vchNombre FROM ImEstudio where bitStatusActivo=1 order by vchNombre"
        Case 2
            Set lstListas = lstExamenes
            vlstrSentencia = "SELECT LaGrupoExamen.intCveGrupo as Clave, " & _
                "LaGrupoExamen.chrNombre as Nombre, " & _
                "'G' as Tipo " & _
                "From LaGrupoExamen " & _
                "where bitEstatusActivo=1" & _
                "Union " & _
                "SELECT LaExamen.intCveExamen as Clave, " & _
                "LaExamen.chrNombre AS Nombre, " & _
                "'E' as Tipo " & _
                "From LaExamen " & _
                "where bitEstatusActivo=1 order by Tipo,Nombre"

        Case 3
            Set lstListas = lstOtrosConceptos
            vlstrSentencia = "SELECT intCveConcepto, chrDescripcion FROM PvOtroConcepto where bitEstatus=1 order by chrDescripcion "
        Case 4
            Set lstListas = lstCirugias
            vlstrSentencia = "SELECT intCveCirugia, vchDescripcion FROM ExCirugia where bitActiva=1 order by vchDescripcion "
        Case 5 'Conceptos de facturación
            Set lstListas = lstConceptos
            vlstrSentencia = "SELECT smiCveCOncepto, chrDescripcion FROM PvConceptoFacturacion where bitActivo=1 order by chrDescripcion "
        Case 6
            Set lstListas = lstPaquetes
            vlstrSentencia = "SELECT intNumPaquete, trim(chrDescripcion) descripcion FROM PvPaquete where nvl(bitActivo,0) = 1 order by chrDescripcion "
    End Select
    
    lstListas.Visible = False
    lstListas.Clear
    Set rsDatos = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    
    If rsDatos.RecordCount > 500 Then
        freBarra.Top = 2220
        freBarra.Visible = True
        freBarra.Refresh
    End If
    
    Do While Not rsDatos.EOF
        Select Case sstElementos.Tab
            Case 0 'Articulos
                lstListas.AddItem (rsDatos!vchNombreComercial)
                lstListas.ItemData(lstArticulos.newIndex) = CLng(rsDatos!intIdArticulo)
            Case 1 'Estudios de Imagenologia
                lstListas.AddItem (rsDatos!vchNombre)
                lstListas.ItemData(lstEstudios.newIndex) = rsDatos!intCveEstudio
            Case 2 'Examenes de Laboratorio
                lstListas.AddItem (RTrim(rsDatos!Nombre) + IIf(rsDatos!tipo = "G", "  (G)", ""))
                lstListas.ItemData(lstExamenes.newIndex) = rsDatos!clave
            Case 3 ' Otros conceptos
                lstListas.AddItem (rsDatos!chrDescripcion)
                lstListas.ItemData(lstOtrosConceptos.newIndex) = rsDatos!intCveConcepto
            Case 4 'Cirugías
                lstListas.AddItem (rsDatos!VCHDESCRIPCION)
                lstListas.ItemData(lstCirugias.newIndex) = rsDatos!intCveCirugia
            Case 5 'Conceptos de facturación
                lstListas.AddItem (rsDatos!chrDescripcion)
                lstListas.ItemData(lstConceptos.newIndex) = rsDatos!smicveconcepto
            Case 6 'Paquetes
                lstListas.AddItem (rsDatos!Descripcion)
                lstListas.ItemData(lstPaquetes.newIndex) = rsDatos!intnumpaquete
        End Select
        
        rsDatos.MoveNext
        If Not rsDatos.EOF And rsDatos.RecordCount > 500 Then
            If rsDatos.Bookmark Mod 100 = 0 Then
                pgbCargando.Value = (rsDatos.Bookmark / rsDatos.RecordCount) * 100
            End If
        End If
    Loop
   
    If lstListas.ListCount > 0 Then
        lstListas.ListIndex = 0
    End If
    
    rsDatos.Close
    freBarra.Visible = False
    cmdCargar.Enabled = True
    lstListas.Visible = True
    If lstListas.Enabled And lstListas.Visible Then lstListas.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaArticulos"))
    Unload Me
End Sub

Private Sub pCargaTiposPaciente()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select tnyCveTipoPaciente, vchDescripcion from AdTipoPaciente order by vchDescripcion"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Do While Not rs.EOF
        lstTiposPaciente.AddItem rs!VCHDESCRIPCION
        lstTiposPaciente.ItemData(lstTiposPaciente.newIndex) = rs!tnyCveTipoPaciente
        rs.MoveNext
    Loop
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTiposPaciente"))
    Unload Me
End Sub

Private Sub pCargaTipoConvenio()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    vlstrSentencia = "Select tnyCveTipoConvenio, vchDescripcion from ccTipoConvenio order by vchDescripcion"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    cboTipoConvenio.AddItem "<Todas>"
    cboTipoConvenio.ItemData(cboTipoConvenio.newIndex) = 0
    Do While Not rs.EOF
        cboTipoConvenio.AddItem rs!VCHDESCRIPCION
        cboTipoConvenio.ItemData(cboTipoConvenio.newIndex) = rs!tnyCveTipoConvenio
        rs.MoveNext
    Loop
    rs.Close
    If cboTipoConvenio.ListCount > 0 Then
        cboTipoConvenio.ListIndex = 0
    Else
        cboTipoConvenio.Enabled = False
        lstEmpresas.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTipoConvenio"))
    Unload Me
End Sub

Private Sub pCargaEmpresas()
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    
    If cboTipoConvenio.ListIndex > 0 Then
        vlstrSentencia = "Select intCveEmpresa, vchDescripcion " & _
                            " from ccEmpresa " & _
                            "Where tnyCveTipoConvenio = " & RTrim(str(cboTipoConvenio.ItemData(cboTipoConvenio.ListIndex))) & _
                            " order by vchDescripcion "
    Else
        vlstrSentencia = "Select intCveEmpresa, vchDescripcion " & _
                            " from ccEmpresa " & _
                            " order by vchDescripcion "
    End If
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    lstEmpresas.Clear
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            lstEmpresas.AddItem rs!VCHDESCRIPCION
            lstEmpresas.ItemData(lstEmpresas.newIndex) = rs!intcveempresa
            rs.MoveNext
        Loop
        lstEmpresas.Enabled = True
    Else
        lstEmpresas.Enabled = False
    End If
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaEmpresas"))
    Unload Me
End Sub

Private Sub cboTipoConvenio_Click()
    On Error GoTo NotificaError
    
    pCargaEmpresas

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoConvenio_Click"))
    Unload Me
End Sub





Private Sub chkVigencia_Click()
    On Error GoTo NotificaError

    mskFechaInicioVigencia.Enabled = chkVigencia.Value = 1
    mskFechaFinVigencia.Enabled = chkVigencia.Value = 1
    
    If chkVigencia.Value = 0 Then
        mskFechaInicioVigencia.Mask = ""
        mskFechaInicioVigencia.Text = ""
        mskFechaInicioVigencia.Mask = "##/##/####"
        
        mskFechaFinVigencia.Mask = ""
        mskFechaFinVigencia.Text = ""
        mskFechaFinVigencia.Mask = "##/##/####"
    Else
        pEnfocaMkTexto mskFechaInicioVigencia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkVigencia_Click"))
    Unload Me
End Sub

Private Sub chkVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        chkVigencia_Click
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkVigencia_KeyDown"))
    Unload Me
End Sub

Private Sub cmdCargar_Click()
    On Error GoTo NotificaError
    
    pCargaArticulos

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargar_Click"))
    Unload Me
End Sub

Private Sub pConfiguraGridCargos()
    On Error GoTo NotificaError
    
    With grdDescuentos
        .Cols = 8
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Descripción|Descuento|Tipo||Inicio vigencia|Fin vigencia|clave"
        .ColWidth(0) = 100 'Fix
        .ColWidth(1) = 3550 'Descripción
        .ColWidth(2) = 900 'Descuento
        .ColWidth(3) = 500  'Tipo
        .ColWidth(4) = 200  'Porciento o cantidad
        .ColWidth(5) = 1200  'Fecha inicio vigencia
        .ColWidth(6) = 1200  'Fecha fin vigencia
        .ColWidth(7) = 0  'Clave del cargo
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
        .TextMatrix(1, 2) = ""
        .RowData(1) = -1
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    Unload Me
End Sub

Private Sub cmdExportar_Click()
On Error GoTo NotificaError
Dim o_Excel As Object
Dim o_Libro As Object
Dim o_Sheet As Object
Dim intRowExcel As Integer
Dim intRow As Integer
Dim vlintContador As Integer
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Sheet = o_Libro.Worksheets(1)
    
    If Not IsObject(o_Excel) Then
        MsgBox "Necesitas Microsoft Excel para utilizar esta funcionalidad", _
           vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    'datos del repote
    If SSTDescuentos.Tab = 0 Then
        o_Excel.Cells(1, 1).Value = "Número de cuenta"
        o_Excel.Cells(1, 3).Value = Trim(txtMovimientoPaciente)
    ElseIf SSTDescuentos.Tab = 1 Then
        For vlintContador = 0 To optTipoPac2.Count
            If optTipoPac2(vlintContador).Value Then
                o_Excel.Cells(1, 1).Value = "Clave"
                o_Excel.Cells(1, 3).Value = lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex)
                Exit For
            End If
        Next
    ElseIf SSTDescuentos.Tab = 2 Then
        o_Excel.Cells(1, 1).Value = "Número convenio"
        o_Excel.Cells(1, 3).Value = lstEmpresas.ItemData(lstEmpresas.ListIndex)
    End If
    
    If SSTDescuentos.Tab = 0 Then
        o_Excel.Cells(2, 1).Value = "Nombre del paciente"
        o_Excel.Cells(2, 3).Value = txtPaciente
    ElseIf SSTDescuentos.Tab = 1 Then
        o_Excel.Cells(2, 1).Value = "Tipo de paciente"
        o_Excel.Cells(2, 3).Value = lstTiposPaciente.Text
    ElseIf SSTDescuentos.Tab = 2 Then
        o_Excel.Cells(2, 1).Value = "Empresa"
        o_Excel.Cells(2, 3).Value = lstEmpresas.Text
    End If
    'columnas titulos
    o_Excel.Cells(4, 1).Value = "Tipo"
    o_Excel.Cells(4, 2).Value = "Clave"
    o_Excel.Cells(4, 3).Value = "Descripción"
    o_Excel.Cells(4, 4).Value = "Tipo descuento"
    o_Excel.Cells(4, 5).Value = "Descuento"
    o_Excel.Cells(4, 6).Value = "Inicio vigencia"
    o_Excel.Cells(4, 7).Value = "Fin vigencia"
    'Diseño
    o_Sheet.range("A4:G4").HorizontalAlignment = -4108
    o_Sheet.range("A4:G4").VerticalAlignment = -4108
    o_Sheet.range("A4:G4").WrapText = True
    o_Sheet.range("A5").Select
    o_Excel.ActiveWindow.FreezePanes = True
    o_Sheet.range("A4:G4").Interior.ColorIndex = 15 '15 48
    'fin diseño
    o_Sheet.range("B:B").NumberFormat = "0"
    o_Sheet.range("F:F").NumberFormat = "dd/mmm/yyyy"
    o_Sheet.range("G:G").NumberFormat = "dd/mmm/yyyy"
    o_Excel.range(o_Excel.Cells(1, 1), o_Excel.Cells(1, 2)).Merge 'Merge numero convenio,cuenta del paciente, tipo paciente
    o_Excel.range(o_Excel.Cells(2, 1), o_Excel.Cells(2, 2)).Merge 'Merge numero convenio,cuenta del paciente, tipo paciente
    'Tamaño
    o_Sheet.range("C:C").Columnwidth = 50
    o_Sheet.range("D:D").Columnwidth = 15
    o_Sheet.range("F:F").Columnwidth = 15
    o_Sheet.range("G:G").Columnwidth = 15
    intRowExcel = 5
    grdDescuentos.Sort = flexSortStringNoCaseAscending
    'Recorre el grid y llena el Excel
    For intRow = 2 To grdDescuentos.Rows
        If grdDescuentos.RowHeight(intRow - 1) > 0 Then
            If grdDescuentos.TextMatrix(intRow - 1, 4) = "%" Then
                o_Sheet.Cells(intRowExcel, 5).NumberFormat = "#.0000%"
            ElseIf grdDescuentos.TextMatrix(intRow - 1, 4) = "$" Then
                o_Sheet.Cells(intRowExcel, 5).NumberFormat = "$0.0000"
            Else
                o_Sheet.Cells(intRowExcel, 5).NumberFormat = "General"
            End If
            With grdDescuentos
                o_Sheet.Cells(intRowExcel, 1).Value = .TextMatrix(intRow - 1, 3) & " "
                o_Sheet.Cells(intRowExcel, 2).Value = .TextMatrix(intRow - 1, 7) & " "
                o_Sheet.Cells(intRowExcel, 3).Value = .TextMatrix(intRow - 1, 1) & " "
                o_Sheet.Cells(intRowExcel, 4).Value = IIf(.TextMatrix(intRow - 1, 4) = "%", "PORCENTAJE", IIf(.TextMatrix(intRow - 1, 4) = "C", "COSTO", IIf(.TextMatrix(intRow - 1, 4) = "", "", "CANTIDAD"))) & " "
                If .TextMatrix(intRow - 1, 4) = "%" Then
                    o_Sheet.Cells(intRowExcel, 5).Value = (Val(.TextMatrix(intRow - 1, 2)) / 100)
                ElseIf .TextMatrix(intRow - 1, 4) = "$" Then
                    o_Sheet.Cells(intRowExcel, 5).Value = .TextMatrix(intRow - 1, 2)
                Else
                    o_Sheet.Cells(intRowExcel, 5).Value = ""
                End If
                'o_Sheet.Cells(intRowExcel, 5).Value = IIf(IsNumeric(Format(.TextMatrix(intRow - 1, 2), "##.0000")), (Val(.TextMatrix(intRow - 1, 2)) / 100), "") & " "
                o_Sheet.Cells(intRowExcel, 6).Value = .TextMatrix(intRow - 1, 5) & " "
                o_Sheet.Cells(intRowExcel, 7).Value = .TextMatrix(intRow - 1, 6) & " "
            End With
            intRowExcel = intRowExcel + 1
        End If
    Next
    'La información ha sido exportada exitosamente
    MsgBox SIHOMsg(1185), vbOKOnly + vbInformation, "Mensaje"
    o_Excel.Visible = True
    
    Set o_Excel = Nothing
        
Exit Sub
NotificaError:
    ' -- Cierra la hoja y la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Sheet Is Nothing Then Set o_Sheet = Nothing
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExportar_Click"))
End Sub

Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    Dim vlstrSentencia As String
    Dim vlstrTipoDescuento As String
    Dim vllngCveAfectada As Long
    Dim vlStrTipoPaciente As String
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    
        
    If Not fblnRevisaPermiso(vglngNumeroLogin, 303, "E", True) Then
        '¡El usuario no tiene permiso para grabar datos!
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    Else
        '-------------------------------------------------------------------
        '   Valida si la cuenta se encuentra bloqueada por trabajo social
        '-------------------------------------------------------------------
        If SSTDescuentos.Tab = 0 Then
            If fblnCuentaBloqueada(Trim(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")) Then
                'No se puede realizar ésta operación. La cuenta se encuentra bloqueada por trabajo social.
                MsgBox SIHOMsg(662), vbCritical, "Mensaje"
                Exit Sub
            End If
        End If
        
        '--------------------------------------------------------
        ' Persona que graba
        '--------------------------------------------------------
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
        '--------------------------------------------------------
        
        Set rsDescuentos = frsRegresaRs("SELECT * FROM PVDESCUENTO where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockOptimistic, adOpenDynamic)
    
        With rsDescuentos
            If .State = 0 Then
                .Open
            End If
            EntornoSIHO.ConeccionSIHO.BeginTrans
            Select Case SSTDescuentos.Tab
                Case 0 'Pacientes
                    vlstrTipoDescuento = "P"
                    vllngCveAfectada = txtMovimientoPaciente.Text
                    vlStrTipoPaciente = IIf(OptTipoPaciente(0).Value, "I", IIf(OptTipoPaciente(1).Value, "E", "A"))
                Case 1 'Tipos de pacientes
                    vlstrTipoDescuento = "T"
                    vllngCveAfectada = lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex)
                    vlStrTipoPaciente = IIf(optTipoPac2(0).Value, "I", IIf(optTipoPac2(1).Value, "E", IIf(optTipoPac2(3).Value, "U", "A")))
                Case 2 'Empresa
                    vlstrTipoDescuento = "E"
                    vllngCveAfectada = lstEmpresas.ItemData(lstEmpresas.ListIndex)
                    vlStrTipoPaciente = IIf(optTipoPac(0).Value, "I", IIf(optTipoPac(1).Value, "E", IIf(optTipoPac(3).Value, "U", "A")))
            End Select
    
            '--------------------------
            ' Borrado de los elementos previamente guardados
            vlstrSentencia = "Delete from PvDescuento " & _
                             "where chrTipoDescuento = " & "'" & vlstrTipoDescuento & "'" & _
                             " and intCveAfectada = " & vllngCveAfectada & _
                             " and chrTipoPaciente = " & "'" & vlStrTipoPaciente & "'" & _
                             " and tnyclaveempresa = " & vgintClaveEmpresaContable
            pEjecutaSentencia vlstrSentencia
            '--------------------------
            If grdDescuentos.RowData(1) <> -1 Then
                 For vlintContador = 1 To grdDescuentos.Rows - 1
                    .AddNew
                    !chrTipoDescuento = vlstrTipoDescuento
                    !intCveAfectada = vllngCveAfectada
                    !CHRTIPOPACIENTE = vlStrTipoPaciente
                    !chrTipoCargo = grdDescuentos.TextMatrix(vlintContador, 3)
                    If grdDescuentos.TextMatrix(vlintContador, 4) = "C" Then
                        !MNYDESCUENTO = 0
                    Else
                        !MNYDESCUENTO = CDec(grdDescuentos.TextMatrix(vlintContador, 2))
                    End If
                    !intTipoDescuento = IIf(grdDescuentos.TextMatrix(vlintContador, 4) = "%", 1, IIf(grdDescuentos.TextMatrix(vlintContador, 4) = "$", 0, 2))
                    If IsDate(grdDescuentos.TextMatrix(vlintContador, 5)) Then
                        !dtmFechaInicioVigencia = CDate(grdDescuentos.TextMatrix(vlintContador, 5) + " 00:00:00")
                        !dtmFechaFinVigencia = CDate(grdDescuentos.TextMatrix(vlintContador, 6) + " 23:59:59")
                    End If
                    If grdDescuentos.TextMatrix(vlintContador, 3) = "CF" Then 'Conceptos de facturación
                       !smicveconcepto = grdDescuentos.RowData(vlintContador)
                       !intCveCargo = 0
                    Else
                       !intCveCargo = grdDescuentos.RowData(vlintContador)
                       !smicveconcepto = 0
                    End If
                     !tnyClaveEmpresa = vgintClaveEmpresaContable
                    .Update
                 Next
            End If
            Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "ASIGNACION DE DESCUENTOS", vlstrTipoDescuento & " " & CStr(vllngCveAfectada) & " " & vlStrTipoPaciente)
            EntornoSIHO.ConeccionSIHO.CommitTrans
         End With
        
        If SSTDescuentos.Tab = 0 Then 'Sólo para tipo MovPaciente
            freActualizando.Top = 2370
            frmMantoDescuentos.Refresh
            freActualizando.Visible = True
            freActualizando.Refresh
            freActualizando.Visible = False
        End If
        SQL = "delete from PvTipoPacienteProceso where PvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
          "and PvTipoPacienteProceso.intproceso = " & enmTipoProceso.Descuento
        pEjecutaSentencia SQL
        
        SQL = "insert into PvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.Descuento & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")"
        pEjecutaSentencia SQL
    
        pCancelar 'Solamente para limpiar la pantalla
        ' Cargar la lista  de los guardados
        If SSTDescuentos.Tab = 0 Then
            pCargaGuardados "P"
        ElseIf SSTDescuentos.Tab = 1 Then
            pCargaGuardados "T"
        Else
            pCargaGuardados "E"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Unload Me
End Sub

Private Sub cmdVerDescuentos_Click()
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlstrCondicion As String
    Dim vlintContador As Integer
    
    If frmMantoDescuentos.Height = 8445 Then ' Cancelar
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            pCancelar
        End If
    Else ' Edicion
        vgstrEstadoManto = "A"
        For vlintContador = frmMantoDescuentos.Height To 7620 Step cgintFactorVentana
            frmMantoDescuentos.Height = vlintContador
            frmMantoDescuentos.Top = Int((vglngDesktop - frmMantoDescuentos.Height) / 2)
        Next
        frmMantoDescuentos.Height = 8445
        frmMantoDescuentos.Top = Int((vglngDesktop - frmMantoDescuentos.Height) / 2)
        If SSTDescuentos.Tab = 0 Then
            SSTDescuentos.TabEnabled(1) = False
            SSTDescuentos.TabEnabled(2) = False
        ElseIf SSTDescuentos.Tab = 1 Then
            SSTDescuentos.TabEnabled(0) = False
            SSTDescuentos.TabEnabled(2) = False
        Else
            SSTDescuentos.TabEnabled(0) = False
            SSTDescuentos.TabEnabled(1) = False
        End If
        cmdVerDescuentos.Caption = "Cancelar asignación"
        cmdGrabarRegistro.Enabled = True
        '---------------------------------------------------------
        ' 1º Creación de la sentencia
        vlstrSentencia = "SELECT PvDescuento.chrTipoDescuento, PvDescuento.intTipoDescuento, " & _
            "PvDescuento.chrTipoCargo, PvDescuento.intCveAfectada, PvDescuento.mnyDescuento," & _
            "PvDescuento.chrTipoPaciente, PvDescuento.intCveCargo, PvDescuento.smiCveConcepto,PvDescuento.dtmFechaInicioVigencia,PvDescuento.dtmFechaFinVigencia, " & _
            "PvConceptoFacturacion.chrDescripcion AS ConceptoFacturacion, " & _
            "LaExamen.chrNombre AS Examen, " & _
            "LaGrupoExamen.chrNombre AS GrupoExamen, " & _
            "PvOtroConcepto.chrDescripcion AS OtroConcepto, " & _
            "ExCirugia.vchDescripcion AS Cirugia, " & _
            "IvArticulo.vchNombreComercial AS Articulo, " & _
            "ImEstudio.vchNombre AS Estudio, " & _
            "trim(pvPaquete.chrDescripcion) as Paquete " & _
            "FROM PvDescuento " & _
            "LEFT OUTER JOIN PvConceptoFacturacion ON PvDescuento.smiCveConcepto = PvConceptoFacturacion.smiCveConcepto " & _
            "LEFT OUTER JOIN LaExamen ON " & _
            "PvDescuento.intCveCargo = LaExamen.IntCveExamen LEFT OUTER " & _
            "Join " & _
            "LaGrupoExamen ON " & _
            "PvDescuento.intCveCargo = LaGrupoExamen.intCveGrupo LEFT OUTER "
            vlstrSentencia = vlstrSentencia & _
            "Join " & _
            "ImEstudio ON " & _
            "PvDescuento.intCveCargo = ImEstudio.intCveEstudio LEFT OUTER " & _
            "Join " & _
            "PvOtroConcepto ON " & _
            "PvDescuento.intCveCargo = PvOtroConcepto.intCveConcepto  LEFT " & _
            "Outer Join " & _
            "ExCirugia ON " & _
            "PvDescuento.intCveCargo = ExCirugia.intCveCirugia LEFT OUTER  " & _
            "Join " & _
            "IvArticulo ON " & _
            "PvDescuento.intCveCargo = IvArticulo.intIDArticulo " & _
            "Left join " & _
            "PvPaquete on PvDescuento.intCveCargo = PvPaquete.intNumPaquete "
            
            
            
        'Crear la condicion del select
        If SSTDescuentos.Tab = 0 Then
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'P' " & _
                             "and PvDescuento.intCveAfectada = " & RTrim(txtMovimientoPaciente.Text) & _
                             " and chrTipoPaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
                             
        ElseIf SSTDescuentos.Tab = 1 Then
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'T' " & _
                            "and PvDescuento.intCveAfectada = " & str(lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex)) & _
                             " and chrTipoPaciente = '" & IIf(optTipoPac2(0).Value, "I", IIf(optTipoPac2(1).Value, "E", IIf(optTipoPac2(3).Value, "U", "A"))) & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
        Else
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'E' " & _
                            "and PvDescuento.intCveAfectada = " & str(lstEmpresas.ItemData(lstEmpresas.ListIndex)) & _
                             " and chrTipoPaciente = '" & IIf(optTipoPac(0).Value, "I", IIf(optTipoPac(1).Value, "E", IIf(optTipoPac(3).Value, "U", "A"))) & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
        End If
        
        vlstrSentencia = vlstrSentencia & vlstrCondicion
        vlstrCondicion = " and (" & _
                                "(PvConceptoFacturacion.chrDescripcion is not null and pvDescuento.chrTipoCargo = 'CF') or " & _
                                "(LaExamen.chrNombre is not null and pvDescuento.chrTipoCargo = 'EX') or " & _
                                "(LaGrupoExamen.chrNombre is not null  and pvDescuento.chrTipoCargo = 'GE') or " & _
                                "(PvOtroConcepto.chrDescripcion is not null  and pvDescuento.chrTipoCargo = 'OC') or " & _
                                "(ExCirugia.vchDescripcion is not null  and pvDescuento.chrTipoCargo = 'CI') or " & _
                                "(IvArticulo.vchNombreComercial is not null  and pvDescuento.chrTipoCargo = 'AR') or " & _
                                "(PvPaquete.chrDescripcion is not null  and pvDescuento.chrTipoCargo = 'PA') or " & _
                                "(ImEstudio.vchNombre is not null and pvDescuento.chrTipoCargo = 'ES'))"
        vlstrSentencia = vlstrSentencia & vlstrCondicion
        '---------------------------------------------------------
        ' 2º Ejecucion de la consulta
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        
        '---------------------------------------------------------
        ' 3º Cargar el Grid
        pConfiguraGridCargos
        With grdDescuentos
            .Redraw = False
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                   If .RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    
                    Select Case rs!chrTipoCargo
                    Case "AR"
                        .TextMatrix(.Row, 1) = rs!Articulo
                    Case "CF"
                        .TextMatrix(.Row, 1) = rs!ConceptoFacturacion
                    Case "CI"
                        .TextMatrix(.Row, 1) = rs!Cirugia
                    Case "OC"
                        .TextMatrix(.Row, 1) = rs!OtroConcepto
                    Case "EX"
                        .TextMatrix(.Row, 1) = rs!Examen
                    Case "ES"
                        .TextMatrix(.Row, 1) = rs!Estudio
                    Case "GE"
                        .TextMatrix(.Row, 1) = rs!GrupoExamen
                    Case "PA"
                        .TextMatrix(.Row, 1) = rs!Paquete
                    End Select
                    
                    If rs!intTipoDescuento = 1 Then 'Porcentaje
                        .TextMatrix(.Row, 2) = rs!MNYDESCUENTO
                        .TextMatrix(.Row, 4) = "%"
                    ElseIf rs!intTipoDescuento = 0 Then 'Cantidad
                        .TextMatrix(.Row, 2) = Format(str(rs!MNYDESCUENTO), "$###,###,###.##")
                        .TextMatrix(.Row, 4) = "$"
                    Else
                        .TextMatrix(.Row, 2) = "Costo" 'Costo
                        .TextMatrix(.Row, 4) = "C"
                    End If
                    .TextMatrix(.Row, 3) = rs!chrTipoCargo
                    If IsDate(rs!dtmFechaInicioVigencia) Then
                        .TextMatrix(.Row, 5) = Format(rs!dtmFechaInicioVigencia, "dd/mmm/yyyy")
                    End If
                    If IsDate(rs!dtmFechaFinVigencia) Then
                        .TextMatrix(.Row, 6) = Format(rs!dtmFechaFinVigencia, "dd/mmm/yyyy")
                    End If
                    
                    If rs!chrTipoCargo = "CF" Then
                        .TextMatrix(.Row, 7) = rs!smicveconcepto
                    Else
                        .TextMatrix(.Row, 7) = rs!intCveCargo
                    End If
                    
                    If rs!chrTipoCargo = "CF" Then
                        .RowData(.Row) = rs!smicveconcepto
                    Else
                        .RowData(.Row) = rs!intCveCargo
                    End If
                    rs.MoveNext
                Loop
            End If
            sstElementos.Tab = 5
            cmdCargar_Click
            If lstConceptos.Enabled And lstConceptos.Visible Then lstConceptos.SetFocus
            .Redraw = True
            .Refresh
        
        End With
        FrePaciente.Enabled = False
        freTipoPaciente.Enabled = False
        FreEmpresa.Enabled = False
        FreConDescuento.Enabled = False
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVerDescuentos_Click"))
    Unload Me
End Sub

Private Sub pSeleccionaElemento()
    On Error GoTo NotificaError
    
    Dim vlstrCualLista As String
    Dim vlintPosicion As Integer
    Dim lstListas As ListBox
    
    If Val(Format(txtDescuento.Text, "#########.##")) > 0 Or optTipoDescuento(2).Value Then
        With grdDescuentos
            Select Case sstElementos.Tab
            Case 6
                Set lstListas = lstPaquetes
                vlstrCualLista = "PA"
            Case 5
                Set lstListas = lstConceptos
                vlstrCualLista = "CF"
            Case 4
                Set lstListas = lstCirugias
                vlstrCualLista = "CI"
            Case 3
                Set lstListas = lstOtrosConceptos
                vlstrCualLista = "OC"
            Case 2
                Set lstListas = lstExamenes
                If Mid(lstListas.List(lstListas.ListIndex), Len(lstListas.List(lstListas.ListIndex)) - 2, 3) = "(G)" Then
                    vlstrCualLista = "GE"
                Else
                    vlstrCualLista = "EX"
                End If
            Case 1
               Set lstListas = lstEstudios
                vlstrCualLista = "ES"
            Case 0
               Set lstListas = lstArticulos
                vlstrCualLista = "AR"
            End Select
            If lstListas.ListIndex <> -1 Then
                vlintPosicion = FintBuscaEnRowData(grdDescuentos, lstListas.ItemData(lstListas.ListIndex), vlstrCualLista)
                If vlintPosicion = -1 Then        'Cuando no esta en la lista
                    If .RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                Else
                    .Row = vlintPosicion 'Funciona como modificación
                End If
                .TextMatrix(.Row, 1) = lstListas.List(lstListas.ListIndex)
                If optTipoDescuento(2).Value Then
                    .TextMatrix(.Row, 2) = "Costo"
                Else
                    .TextMatrix(.Row, 2) = IIf(optTipoDescuento(1).Value, Format(str(txtDescuento.Text), "$###,###,###.##"), txtDescuento.Text)
                End If
                .TextMatrix(.Row, 3) = vlstrCualLista
                .TextMatrix(.Row, 4) = IIf(optTipoDescuento(1).Value, "$", IIf(optTipoDescuento(0).Value, "%", "C"))
                .TextMatrix(.Row, 5) = IIf(chkVigencia.Value = 0, "", Format(mskFechaInicioVigencia.Text, "dd/mmm/yyyy"))
                .TextMatrix(.Row, 6) = IIf(chkVigencia.Value = 0, "", Format(mskFechaFinVigencia.Text, "dd/mmm/yyyy"))
                .TextMatrix(.Row, 7) = lstListas.ItemData(lstListas.ListIndex)
                .RowData(.Row) = lstListas.ItemData(lstListas.ListIndex)
                .Redraw = True
                .Refresh
            Else
                MsgBox SIHOMsg(3), vbCritical, "Mensaje"
            End If
            If lstListas.Enabled And lstListas.Visible Then lstListas.SetFocus
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaElemento"))
    Unload Me
End Sub
Private Sub pepeLuis(hola As String)
    
    Dim vlstrCualLista As String
    Dim vlintPosicion As Integer
    Dim lstListas As ListBox
    
        With grdDescuentos
            Select Case hola
            Case "PA"
                Set lstListas = lstPaquetes
                vlstrCualLista = "PA"
            Case "CF"
                Set lstListas = lstConceptos
                vlstrCualLista = "CF"
            Case "CI"
                Set lstListas = lstCirugias
                vlstrCualLista = "CI"
            Case "OC"
                Set lstListas = lstOtrosConceptos
                vlstrCualLista = "OC"
            Case "GE"
                Set lstListas = lstExamenes
                vlstrCualLista = "GE"
            Case "EX"
                Set lstListas = lstExamenes
                vlstrCualLista = "EX"
            Case "ES"
               Set lstListas = lstEstudios
                vlstrCualLista = "ES"
            Case "AR"
               Set lstListas = lstArticulos
                vlstrCualLista = "AR"
            End Select
            If lstListas.ListIndex <> -1 Then
                vlintPosicion = FintBuscaEnRowData(grdDescuentos, lstListas.ItemData(lstListas.ListIndex), vlstrCualLista)
                If Not vlintPosicion = -1 Then        'Cuando no esta en la lista
                    .Row = vlintPosicion 'Funciona como modificación
                End If
                .RowData(.Row) = lstListas.ItemData(lstListas.ListIndex)
                .Redraw = True
                .Refresh
            End If
            If lstListas.Enabled And lstListas.Visible Then lstListas.SetFocus
        End With
End Sub
Private Sub cmdSelecciona_Click(Index As Integer)
    On Error GoTo NotificaError
    
    If Index = 0 Then
        pPideDescuento
    ElseIf Index = 2 Then
        pPideDescuento
    Else
        grdDescuentos_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSelecciona_Click"))
    Unload Me
End Sub

Private Function FintBuscaEnRowData(grdHBusca As MSHFlexGrid, vlintCriterio As Long, vlstrTipoElemento)
    On Error GoTo NotificaError
    
    Dim vlintContador As Long
    
    FintBuscaEnRowData = -1
    With grdHBusca
    For vlintContador = 1 To .Rows - 1
        If .RowData(vlintContador) = vlintCriterio And vlstrTipoElemento = .TextMatrix(vlintContador, 3) Then
            FintBuscaEnRowData = vlintContador
            Exit For
        End If
    Next
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":FintBuscaEnRowData"))
    Unload Me
End Function

Private Sub cmdImportar_Click()
On Error GoTo NotificaError
Dim objXLApp As Object
Dim intLoopCounter As Integer
Dim txtRuta As String
Dim intRows As Integer
Dim vlblnbandera As Boolean
Dim intRowsAct As Integer
Dim intResul As Integer
Dim pbytDecimales As Byte
Dim vlintContador As Integer
Dim vlintContadorConceptos As Integer
Dim vlstrCualLista As String
Dim vlintPosicion As Integer
Dim lstListas As ListBox
Dim vlintDimensionBorrar As Integer
Dim vlintCargos As Integer
Dim vlinRepetidos As Integer
Dim vlstrValidador As String
Dim vlblnCargos As Boolean

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
    vlblnbandera = False
    intRowsAct = grdDescuentos.Rows - 1
    intResul = 0
    vlstrValidador = ""
    With objXLApp
        .Workbooks.Open txtRuta
        .Workbooks(1).Worksheets(1).Select
        
        If grdDescuentos.Rows - 1 <> (CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 4) Then
            'Información no valida
            MsgBox "La cantidad de cargos que se intentan importar son diferentes a los cargos que ya se tienen asignados", vbOKOnly + vbInformation, "Mensaje"
            .Workbooks(1).Close False
            .Quit
            objXLApp = Nothing
            Exit Sub
        End If

        If SSTDescuentos.Tab = 0 Then
            If Not txtMovimientoPaciente.Text = Trim(.range("C" & 1)) Then
                vlstrValidador = "El número de cuenta no coincide con el documento que se intenta importar"
            Else
                If Not txtPaciente.Text = Trim(.range("C" & 2)) Then
                    vlstrValidador = "El nombre del paciente no coincide con el documento que se intenta importar"
                End If
            End If
        
        ElseIf SSTDescuentos.Tab = 1 Then
            If Not lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex) = Trim(.range("C" & 1)) Then
                vlstrValidador = "La clave del tipo paciente  no coincide con el documento que se intenta importar"
            End If
    
            If Not lstTiposPaciente.Text = Trim(.range("C" & 2)) Then
                vlstrValidador = "El tipo paciente  no coincide con el documento que se intenta importar"
            End If
        ElseIf SSTDescuentos.Tab = 2 Then
            If Not lstEmpresas.ItemData(lstEmpresas.ListIndex) = Trim(.range("C" & 1)) Then
                vlstrValidador = "El número de convenio  no coincide con el documento que se intenta importar"
            End If
    
            If Not lstEmpresas.Text = Trim(.range("C" & 2)) Then
                vlstrValidador = "El nombre de la empresa  no coincide con el documento que se intenta importar"
            End If
        End If
        
        If vlstrValidador <> "" Then
            'Información no valida
            MsgBox vlstrValidador, vbOKOnly + vbInformation, "Mensaje"
            .Workbooks(1).Close False
            .Quit
            objXLApp = Nothing
            Exit Sub
        End If
        
        vlintDimensionBorrar = 0
        Erase cargosFaltantes
        vlinRepetidos = 4
        ReDim Preserve cargosFaltantes(CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row) - 4)
        vlblnCargos = False
        
        For intLoopCounter = 5 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
            Dim txtClaveCargo As String
            Dim txtDescripcion As String
            txtClaveCargo = .range("B" & intLoopCounter)
            txtDescripcion = Trim(UCase(.range("C" & intLoopCounter)))
            If txtClaveCargo <> "" And txtClaveCargo <> 0 Then
                For intRows = 1 To grdDescuentos.Rows - 1
                    Dim txtClaveCargoGrid As String
                    Dim txtDescripcionGrid As String
                    txtClaveCargoGrid = grdDescuentos.TextMatrix(intRows, 7)
                    txtDescripcionGrid = Trim(grdDescuentos.TextMatrix(intRows, 1))
                    If txtClaveCargo = txtClaveCargoGrid Then
                        vlblnCargos = True
                        intResul = intResul + 1
                        If UCase(txtDescripcion) <> UCase(txtDescripcionGrid) Then
                            'Descripción del cargo
                            MsgBox SIHOMsg(1658) & " " & txtDescripcion & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                            grdDescuentos.Clear
                            grdDescuentos.Rows = 2
                            pConfiguraGridCargos
                            pLlenarDescuentos
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        End If
                    Else
                        If Not vlblnCargos Then
                            vlblnCargos = False
                        End If
                    End If
                
         
                Next intRows
                    If Not vlblnCargos Then
                        cargosFaltantes(intLoopCounter - 5).lngClave = txtClaveCargo
                        cargosFaltantes(intLoopCounter - 5).strDescripcion = .range("C" & intLoopCounter)
                    Else
                        cargosFaltantes(intLoopCounter - 5).lngClave = 0
                        cargosFaltantes(intLoopCounter - 5).strDescripcion = ""
                    End If
                    
                    vlblnCargos = False
            Else
                'Clave Cargo invalido
                MsgBox SIHOMsg(1654) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                grdDescuentos.Clear
                grdDescuentos.Rows = 2
                pConfiguraGridCargos
                pLlenarDescuentos
                .Workbooks(1).Close False
                .Quit
                objXLApp = Nothing
                Exit Sub
            End If
            
        Next intLoopCounter
        
        If intResul = intRowsAct Then
            For intRows = grdDescuentos.Rows - 1 To 0 Step -1
                grdDescuentos.RemoveItem (intRows)
            Next intRows
            For intLoopCounter = 5 To CInt(.Cells.Find("*", SearchOrder:=1, SearchDirection:=2).Row)
            '---------------------------VALIDACIONES DEL EXCEL ANTES DE IMPORTAR---------------------------------
            '---------------------------Validar el tipo----------------------------------------------------------
            If Not (Trim(UCase(.range("A" & intLoopCounter))) = "CF" Or Trim(UCase(.range("A" & intLoopCounter))) = "AR" Or Trim(UCase(.range("A" & intLoopCounter))) = "CI" Or Trim(UCase(.range("A" & intLoopCounter))) = "ES" Or Trim(UCase(.range("A" & intLoopCounter))) = "EX" Or Trim(UCase(.range("A" & intLoopCounter))) = "OC" Or Trim(UCase(.range("A" & intLoopCounter))) = "PA" Or Trim(UCase(.range("A" & intLoopCounter))) = "GE") Then
                    'Tipo invalido
                    MsgBox SIHOMsg(1657) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                    grdDescuentos.Clear
                    grdDescuentos.Rows = 2
                    pConfiguraGridCargos
                    pLlenarDescuentos
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
            '---------------------------FIN Validar el tipo----------------------------------------------------------
            '---------------------------Validar tipo descuento---------------------------------------------------
                If Not (Trim(UCase(.range("D" & intLoopCounter))) = "CANTIDAD" Or Trim(UCase(.range("D" & intLoopCounter))) = "PORCENTAJE" Or Trim(UCase(.range("D" & intLoopCounter))) = "COSTO" Or (Trim(UCase(UCase(.range("D" & intLoopCounter)))) = "COSTO" And (Trim(.range("E" & intLoopCounter)) = ""))) Then
                    'Tipo de descuento invalido
                    MsgBox SIHOMsg(1655) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                    grdDescuentos.Clear
                    grdDescuentos.Rows = 2
                    pConfiguraGridCargos
                    pLlenarDescuentos
                    .Workbooks(1).Close False
                    .Quit
                    objXLApp = Nothing
                    Exit Sub
                End If
            '---------------------------FIN Validar tipo descuento---------------------------------------------------
            
            '---------------------------Validar las fechas si trae---------------------------------------------------
                If Trim(.range("F" & intLoopCounter)) <> "" Or Trim(.range("G" & intLoopCounter)) <> "" Then
                    If Not IsDate(Trim(.range("F" & intLoopCounter))) Then
                        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                        MsgBox SIHOMsg(1661) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                        grdDescuentos.Clear
                        grdDescuentos.Rows = 2
                        pConfiguraGridCargos
                        pLlenarDescuentos
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    Else
                        'Primero validamos que le rango de fechas aun este valido
                        If Not (((CDate(Trim(.range("G" & intLoopCounter))) >= fdtmServerFecha)) And ((CDate(Trim(.range("F" & intLoopCounter))) <= CDate(Trim(.range("G" & intLoopCounter)))))) Then
                            'Si el rango de fechas esta vencido avisar, es decir: que la fecha final sea mayor que la fecha inicial.
                            If Not (((CDate(Trim(.range("G" & intLoopCounter))) >= CDate(Trim(.range("F" & intLoopCounter)))))) Then
                                'Si el rango ya esta vencido validamos desde la fecha inicio
                                If CDate(Trim(.range("F" & intLoopCounter))) < fdtmServerFecha Then
                                    '¡Fecha fin vigencia no válida! Formato de fecha dd/mm/aaaa.
                                        MsgBox SIHOMsg(1661) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                                        grdDescuentos.Clear
                                        grdDescuentos.Rows = 2
                                        pConfiguraGridCargos
                                        pLlenarDescuentos
                                        .Workbooks(1).Close False
                                        .Quit
                                        objXLApp = Nothing
                                        Exit Sub
                                Else
                                    If Not IsDate(Trim(.range("G" & intLoopCounter))) Then
                                        '¡Fecha fin vigencia no válida! Formato de fecha dd/mm/aaaa.
                                        MsgBox SIHOMsg(1664) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                                        grdDescuentos.Clear
                                        grdDescuentos.Rows = 2
                                        pConfiguraGridCargos
                                        pLlenarDescuentos
                                        .Workbooks(1).Close False
                                        .Quit
                                        objXLApp = Nothing
                                        Exit Sub
                                    Else
                                        If CDate(Trim(.range("F" & intLoopCounter))) > CDate(Trim(.range("G" & intLoopCounter))) Then
                                            '¡Rango de fechas no válido!
                                            MsgBox SIHOMsg(1664) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                                            grdDescuentos.Clear
                                            grdDescuentos.Rows = 2
                                            pConfiguraGridCargos
                                            pLlenarDescuentos
                                            .Workbooks(1).Close False
                                            .Quit
                                            objXLApp = Nothing
                                            Exit Sub
                                        End If
                                    End If
                                End If 'Fin del validacion de fecha de inicio
                            Else
                                '¡Rango de fechas no válido!
                                MsgBox SIHOMsg(1665) & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                                grdDescuentos.Clear
                                grdDescuentos.Rows = 2
                                pConfiguraGridCargos
                                pLlenarDescuentos
                                .Workbooks(1).Close False
                                .Quit
                                objXLApp = Nothing
                                Exit Sub
                            End If 'Fin rango fechas vencidas avisar
                        End If 'Fin rango de fechas valido entre inicial, final y servidor.
                    End If
                End If
                '---------------------------FIN Validar las fechas si trae-----------------------------------------------
                '---------------------------FIN VALIDACIONES DEL EXCEL ANTES DE IMPORTAR---------------------------------
                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 1) = Trim(.range("C" & intLoopCounter)) 'Descripcion
                If Trim(UCase(.range("D" & intLoopCounter))) = "PORCENTAJE" Then
                    If IsNumeric(Trim(.range("E" & intLoopCounter))) Or Trim(.range("E" & intLoopCounter)) = "Costo" Then
                        If Trim(.range("E" & intLoopCounter)) > 0 Then
                            Dim txtDescuento As Double
                            Dim x1 As Integer
                            Dim x2 As Double
                            Dim strCadena As String
                            Dim vlCiclo As Integer
                            Dim strCaracter As String
                            Dim vlContador As Integer
                            txtDescuento = Format(Round(Trim(.range("E" & intLoopCounter)) * 100, 4), "##.0000")
                            x1 = Int(txtDescuento)
                            x2 = txtDescuento - x1
                            If x2 > 0 Then
                                vlContador = 0
                                strCadena = CStr(txtDescuento)
                                For vlCiclo = Len(strCadena) To 0 Step -1
                                    strCaracter = Mid(strCadena, vlCiclo, 1)
                                    If strCaracter = "0" Then
                                        vlContador = vlContador + 1
                                    End If
                                Next vlCiclo
                                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 2) = Round(txtDescuento, 4 - vlContador)
                            Else
                                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 2) = x1
                            End If
                            grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 4) = IIf(Trim(.range("D" & intLoopCounter)) = "PORCENTAJE", "%", IIf(Trim(.range("D" & intLoopCounter)) = "COSTO", "C", "$")) 'Tipo descuento
                        Else
                            '¡El formato del descuento  es incorrecto!
                            MsgBox SIHOMsg(1659) & " " & LCase((.range("D" & intLoopCounter))) & "es incorrecto!" & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                            grdDescuentos.Clear
                            grdDescuentos.Rows = 2
                            pConfiguraGridCargos
                            pLlenarDescuentos
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        End If
                    Else
                        '¡El formato del descuento del cargo es incorrecto!
                        MsgBox SIHOMsg(1659) & " " & LCase((.range("D" & intLoopCounter))) & "es incorrecto!" & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                        grdDescuentos.Clear
                        grdDescuentos.Rows = 2
                        pConfiguraGridCargos
                        pLlenarDescuentos
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                ElseIf Trim(UCase(.range("D" & intLoopCounter))) = "CANTIDAD" Then
                    If IsNumeric(Trim(.range("E" & intLoopCounter))) Or Trim(.range("E" & intLoopCounter)) = "Costo" Or Val(Trim(.range("E" & intLoopCounter))) > 0 Then
                        If Trim(.range("E" & intLoopCounter)) > 0 Then
                            txtDescuento = Format(Round(Trim(.range("E" & intLoopCounter)), 2), "##.00")
                            grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 2) = Round(txtDescuento, 2)
                            grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 4) = IIf(Trim(.range("D" & intLoopCounter)) = "PORCENTAJE", "%", IIf(Trim(.range("D" & intLoopCounter)) = "COSTO", "C", "$")) 'Tipo descuento
                        Else
                            '¡El formato del descuento del cargo es incorrecto!
                            MsgBox SIHOMsg(1659) & " " & LCase((.range("D" & intLoopCounter))) & "es incorrecto!" & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                            grdDescuentos.Clear
                            grdDescuentos.Rows = 2
                            pConfiguraGridCargos
                            pLlenarDescuentos
                            .Workbooks(1).Close False
                            .Quit
                            objXLApp = Nothing
                            Exit Sub
                        End If
                    Else
                        '¡El formato del descuento del cargo es incorrecto!
                        MsgBox SIHOMsg(1659) & " " & LCase((.range("D" & intLoopCounter))) & "es incorrecto!" & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                        grdDescuentos.Clear
                        grdDescuentos.Rows = 2
                        pConfiguraGridCargos
                        pLlenarDescuentos
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    End If
                ElseIf Trim(UCase(.range("D" & intLoopCounter))) = "COSTO" Then
                    If Trim(.range("E" & intLoopCounter)) <> "" Then
                        '¡El formato del descuento del cargo es incorrecto!
                        MsgBox SIHOMsg(1659) & " " & LCase((.range("D" & intLoopCounter))) & "es incorrecto!" & " Renglón " & intLoopCounter, vbOKOnly + vbInformation, "Mensaje"
                        grdDescuentos.Clear
                        grdDescuentos.Rows = 2
                        pConfiguraGridCargos
                        pLlenarDescuentos
                        .Workbooks(1).Close False
                        .Quit
                        objXLApp = Nothing
                        Exit Sub
                    Else
                        grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 2) = "Costo" 'Descuento
                        grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 4) = "C"
                    End If
                End If
                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 3) = Trim(.range("A" & intLoopCounter)) 'Tipo
                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 5) = Trim(.range("F" & intLoopCounter)) 'Inicio vigencia
                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 6) = Trim(.range("G" & intLoopCounter)) 'Fin vigencia
                grdDescuentos.TextMatrix(grdDescuentos.Rows - 1, 7) = Trim(.range("B" & intLoopCounter)) 'clave cargo
                grdDescuentos.AddItem ""
            Next intLoopCounter
            grdDescuentos.Rows = grdDescuentos.Rows - 1
            grdDescuentos.Sort = flexSortStringNoCaseAscending
            
            For vlintContador = 1 To grdDescuentos.Rows - 1
                grdDescuentos.RowData(vlintContador) = grdDescuentos.TextMatrix(vlintContador, 7)
            Next
            MsgBox "La información ha sido importada exitosamente.", vbOKOnly + vbInformation, "Mensaje"
        Else
            Dim vlstrMensaje As String
            For vlintCargos = 0 To UBound(cargosFaltantes)
                If cargosFaltantes(vlintCargos).strDescripcion <> "" Then
                    vlstrMensaje = vlstrMensaje & cargosFaltantes(vlintCargos).strDescripcion & ","
                End If
            Next vlintCargos
            vlstrMensaje = Left(vlstrMensaje, Len(vlstrMensaje) - 1)
            
            MsgBox SIHOMsg(1656) & " " & vlstrMensaje, vbOKOnly + vbInformation, "Mensaje"
            vlblnbandera = True
            .Workbooks(1).Close False
            .Quit
            Exit Sub
        End If
        
'        If vlblnbandera And (intRowsAct <> grdDescuentos.Rows - 1) Then
'            For intRows = grdDescuentos.Rows - 1 To intRowsAct + 1 Step -1
'                grdDescuentos.RemoveItem intRows
'            Next intRows
'        End If
        
        .Workbooks(1).Close False
        .Quit
    End With

    Set objXLApp = Nothing

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Command1_Click"))
    Unload Me
End Sub

Private Sub Command2_Click()
    Call cmdImportar_Click
End Sub

Private Sub Command3_Click()
    Call cmdExportar_Click
End Sub

Private Sub Command4_Click()
    Call cmdImportar_Click
End Sub

Private Sub Command5_Click()
    Call cmdExportar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    Select Case vgstrEstadoManto
        Case "A"
            Cancel = True
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pCancelar
            End If
        Case "AS", "MS"
            Cancel = True
            freDescuento.Visible = False
            FreElementos.Enabled = True
            freElementosIncuidos.Enabled = True
            chkTodos.Value = 0
            optTipoDescuento(0).Enabled = True
            optTipoDescuento(1).Enabled = True
            optTipoDescuento(2).Enabled = True
            Select Case sstElementos.Tab
            Case 6
                If lstPaquetes.Visible And lstPaquetes.Enabled Then lstPaquetes.SetFocus
            Case 5
                If lstConceptos.Visible And lstConceptos.Enabled Then lstConceptos.SetFocus
            Case 4
                If lstCirugias.Visible And lstCirugias.Enabled Then lstCirugias.SetFocus
            Case 3
                If lstOtrosConceptos.Visible And lstOtrosConceptos.Enabled Then lstOtrosConceptos.SetFocus
            Case 2
                If lstExamenes.Visible And lstExamenes.Enabled Then lstExamenes.SetFocus
            Case 1
                If lstEstudios.Visible And lstEstudios.Enabled Then lstEstudios.SetFocus
            Case 0
                If lstArticulos.Visible And lstArticulos.Enabled Then lstArticulos.SetFocus
            End Select
    End Select

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo NotificaError

    If rsDescuentos.State = 1 Then
        rsDescuentos.Close
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
    Unload Me
End Sub

Private Sub grdDescuentos_DblClick()
    On Error GoTo NotificaError
    
    With grdDescuentos
        If .Rows > 2 Then
            pBorrarRegMshFGrdData grdDescuentos.Row, grdDescuentos, True
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
            Command3.Enabled = False
            Command2.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
        Else
            pLimpiaMshFGrid grdDescuentos
            .Rows = 2
            pConfiguraGridCargos
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
            Command3.Enabled = False
            Command2.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDescuentos_DblClick"))
    Unload Me
End Sub

Private Sub pPideDescuento()
    On Error GoTo NotificaError
    
    'La opción del descuento por costo aplica sólo para Artículos
    optTipoDescuento(2).Enabled = sstElementos.Tab = 0 Or sstElementos.Tab = 5
    '-------------------------------------------------------------
    
    freDescuento.Visible = True
    FreElementos.Enabled = False
    freElementosIncuidos.Enabled = False
    cmdGrabarRegistro.Enabled = False
    cmdVerDescuentos.Enabled = False
    pEnfocaTextBox txtDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pPideDescuento"))
    Unload Me
End Sub

Private Sub chkTodos_Click()
    On Error GoTo NotificaError
    
    If chkTodos.Value = 1 And freDescuento.Visible = False Then
        optTipoDescuento(0).Value = True
        optTipoDescuento(1).Enabled = False
        pPideDescuento
    ElseIf chkTodos.Value = 0 And freDescuento.Visible = True Then
            freDescuento.Visible = False
            FreElementos.Enabled = True
            freElementosIncuidos.Enabled = True
            chkTodos.Value = 0
            optTipoDescuento(0).Enabled = True
            optTipoDescuento(1).Enabled = True
            optTipoDescuento(2).Enabled = True
            Select Case sstElementos.Tab
            Case 6
                If lstPaquetes.Visible And lstPaquetes.Enabled Then lstPaquetes.SetFocus
            Case 5
                If lstConceptos.Visible And lstConceptos.Enabled Then lstConceptos.SetFocus
            Case 4
                If lstCirugias.Visible And lstCirugias.Enabled Then lstCirugias.SetFocus
            Case 3
                If lstOtrosConceptos.Visible And lstOtrosConceptos.Enabled Then lstOtrosConceptos.SetFocus
            Case 2
                If lstExamenes.Visible And lstExamenes.Enabled Then lstExamenes.SetFocus
            Case 1
                If lstEstudios.Visible And lstEstudios.Enabled Then lstEstudios.SetFocus
            Case 0
                If lstArticulos.Visible And lstArticulos.Enabled Then lstArticulos.SetFocus
            End Select
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkTodos_Click"))
    Unload Me
End Sub

Private Sub lstArticulos_DblClick()
    On Error GoTo NotificaError
    
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstArticulos_DblClick"))
    Unload Me
End Sub

Private Sub lstArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstArticulos_KeyDown"))
    Unload Me
End Sub

Private Sub lstCirugias_DblClick()
    On Error GoTo NotificaError
    
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstCirugias_DblClick"))
    Unload Me
End Sub

Private Sub lstCirugias_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstCirugias_KeyDown"))
    Unload Me
End Sub

Private Sub lstConceptos_DblClick()
    On Error GoTo NotificaError
    
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstConceptos_DblClick"))
    Unload Me
End Sub

Private Sub lstPaquetes_DblClick()
    On Error GoTo NotificaError
    
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstPaquetes_DblClick()"))
    Unload Me
End Sub
Private Sub lstPaquetes_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstPaquetes_KeyDown"))
    Unload Me
End Sub
Private Sub lstConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstConceptos_KeyDown"))
    Unload Me
End Sub

Private Sub lstEstudios_DblClick()
    On Error GoTo NotificaError
        
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEstudios_DblClick"))
    Unload Me
End Sub

Private Sub lstEstudios_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEstudios_KeyDown"))
    Unload Me
End Sub

Private Sub lstExamenes_DblClick()
    On Error GoTo NotificaError
        
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstExamenes_DblClick"))
    Unload Me
End Sub

Private Sub lstExamenes_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstExamenes_KeyDown"))
    Unload Me
End Sub

Private Sub lstOtrosConceptos_DblClick()
    On Error GoTo NotificaError
    
    pPideDescuento

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstOtrosConceptos_DblClick"))
    Unload Me
End Sub

Private Sub lstOtrosConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pPideDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstOtrosConceptos_KeyDown"))
    Unload Me
End Sub


Private Sub lstConDescuento_DblClick()
    On Error GoTo NotificaError
    
    If SSTDescuentos.Tab = 0 Then
        OptTipoPaciente(0).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "I"
        OptTipoPaciente(1).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "E"
        txtMovimientoPaciente.Text = lstConDescuento.ItemData(lstConDescuento.ListIndex)
        If txtMovimientoPaciente.Visible And txtMovimientoPaciente.Enabled Then txtMovimientoPaciente.SetFocus
        txtMovimientoPaciente_KeyDown 13, 0
        If Trim(txtMovimientoPaciente.Text) <> "" And txtPaciente.Text <> "" Then
            cmdVerDescuentos_Click
        End If
        
    ElseIf SSTDescuentos.Tab = 1 Then
        lstTiposPaciente.ListIndex = fintLocalizaEnLista(lstTiposPaciente, lstConDescuento.ItemData(lstConDescuento.ListIndex))
        optTipoPac2(0).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "I"
        optTipoPac2(1).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "E"
        optTipoPac2(2).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "T"
        optTipoPac2(3).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "U"
        lstTiposPaciente_DblClick
        cmdVerDescuentos_Click
    Else
        cboTipoConvenio.ListIndex = 0
        lstEmpresas.ListIndex = fintLocalizaEnLista(lstEmpresas, lstConDescuento.ItemData(lstConDescuento.ListIndex))
        optTipoPac(0).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "I"
        optTipoPac(1).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "E"
        optTipoPac(2).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "T"
        optTipoPac(3).Value = Mid(Trim(lstConDescuento.List(lstConDescuento.ListIndex)), Len(Trim(lstConDescuento.List(lstConDescuento.ListIndex))) - 1, 1) = "U"
        lstEmpresas_DblClick
        cmdVerDescuentos_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstConDescuento_DblClick"))
    Unload Me
End Sub

Private Function fintLocalizaEnLista(lstLista As ListBox, intClave As Integer) As Integer
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    fintLocalizaEnLista = -1   'Regresa un -1 si no lo encuentra
    For vlintContador = 0 To lstLista.ListCount - 1
        If lstLista.ItemData(vlintContador) = intClave Then
            fintLocalizaEnLista = vlintContador
            vlintContador = lstLista.ListCount + 1
        End If
    Next vlintContador

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaEnLista"))
    Unload Me
End Function

Private Sub lstEmpresas_DblClick()
    On Error GoTo NotificaError
    
    cmdVerDescuentos.Enabled = True
    If cmdVerDescuentos.Visible And cmdVerDescuentos.Enabled Then cmdVerDescuentos.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_DblClick"))
    Unload Me
End Sub

Private Sub lstEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cmdVerDescuentos.Enabled = True
        If cmdVerDescuentos.Visible And cmdVerDescuentos.Enabled Then cmdVerDescuentos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstEmpresas_KeyDown"))
    Unload Me
End Sub


Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    SSTDescuentos.Tab = 0
    vglngDesktop = SysInfo1.WorkAreaHeight
    vgstrEstadoManto = ""
    txtDescuento.Text = 0
    pCargaTiposPaciente
    pCargaTipoConvenio
    pCargaGuardados ("P")
    
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Descuento) > 0 Then
        If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Descuento) = 1 Then
          OptTipoPaciente(0).Value = True
        Else
          OptTipoPaciente(1).Value = True
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    chkVigencia_Click
    
    vgstrNombreForm = Me.Name
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub


Private Sub lstTiposPaciente_DblClick()
    On Error GoTo NotificaError
    
    cmdVerDescuentos.Enabled = True
    If cmdVerDescuentos.Visible And cmdVerDescuentos.Enabled Then cmdVerDescuentos.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTiposPaciente_DblClick"))
    Unload Me
End Sub

Private Sub lstTiposPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        cmdVerDescuentos.Enabled = True
        If cmdVerDescuentos.Visible And cmdVerDescuentos.Enabled Then cmdVerDescuentos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstTiposPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaFinVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pEnfocaTextBox txtDescuento
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFinVigencia_KeyDown"))
    Unload Me
End Sub

Private Sub mskFechaInicioVigencia_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFechaFinVigencia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicioVigencia_KeyDown"))
    Unload Me
End Sub

Private Sub optTipoDescuento_Click(Index As Integer)
    On Error GoTo NotificaError
    
    If optTipoDescuento(1).Value Then 'Cantidad
        txtDescuento.Enabled = True
        txtDescuento.MaxLength = 15
        txtDescuento.Text = "$0.00"
        lblPorcentaje.Visible = False
        txtDescuento.Locked = False
    ElseIf optTipoDescuento(0).Value Then 'Porcentaje
        txtDescuento.Enabled = True
        txtDescuento.MaxLength = 7
        txtDescuento.Text = "0"
        lblPorcentaje.Visible = True
        txtDescuento.Locked = False
    Else
        txtDescuento.Text = ""
        txtDescuento.Locked = True
    End If
    If freDescuento.Visible Then
        If optTipoDescuento(0).Value Or optTipoDescuento(1).Value Then
            pEnfocaTextBox txtDescuento
        Else
            txtDescuento_KeyDown vbKeyReturn, 0
            optTipoDescuento(0).Value = True
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoDescuento_click"))
    Unload Me

End Sub

Private Sub optTipoDescuento_GotFocus(Index As Integer)
    On Error GoTo NotificaError
    
    vgstrEstadoManto = vgstrEstadoManto & "S"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoDescuento_GotFocus"))
    Unload Me

End Sub

Private Sub optTipoDescuento_LostFocus(Index As Integer)
    On Error GoTo NotificaError
    
    vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 1)

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":opTipoDescuento_LostFocus"))
    Unload Me

End Sub

Private Sub optTipoPac_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        lstEmpresas_DblClick
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPac_KeyPress"))
    Unload Me
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    pCancelar

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
    Unload Me
End Sub

Private Sub pCancelar()
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    '-----Botones de importar o exportar-----
    cmdExportar.Enabled = True
    cmdImportar.Enabled = True
    Command3.Enabled = True
    Command2.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    '------- Los frames --------------------
    FrePaciente.Enabled = True
    freTipoPaciente.Enabled = True
    FreEmpresa.Enabled = True
    FreConDescuento.Enabled = True
    '----------------------------------------
    If txtMovimientoPaciente.Visible And txtMovimientoPaciente.Enabled Then txtMovimientoPaciente.SetFocus
    txtEmpresaPaciente.Text = ""
    txtTipoPaciente.Text = ""
    txtPaciente.Text = ""
    For vlintContador = frmMantoDescuentos.Height To 3765 Step -cgintFactorVentana
        frmMantoDescuentos.Height = vlintContador
        frmMantoDescuentos.Top = Int((vglngDesktop - frmMantoDescuentos.Height) / 2)
    Next
    frmMantoDescuentos.Height = 3705
    frmMantoDescuentos.Top = Int((vglngDesktop - frmMantoDescuentos.Height) / 2)
    SSTDescuentos.TabEnabled(0) = True
    SSTDescuentos.TabEnabled(1) = True
    SSTDescuentos.TabEnabled(2) = True
    cmdVerDescuentos.Caption = "Asignar descuentos"
    cmdGrabarRegistro.Enabled = False
    cmdVerDescuentos.Enabled = False
    grdDescuentos.Clear
    grdDescuentos.Rows = 2
    pConfiguraGridCargos
    vgstrEstadoManto = ""
    Select Case SSTDescuentos.Tab
        Case 0
            If txtMovimientoPaciente.Visible And txtMovimientoPaciente.Enabled Then txtMovimientoPaciente.SetFocus
        Case 1
            If lstTiposPaciente.Visible And lstTiposPaciente.Enabled Then lstTiposPaciente.SetFocus
        Case 2
            If lstEmpresas.Visible And lstEmpresas.Enabled Then lstEmpresas.SetFocus
    End Select
    optTipoDescuento(0).Value = True
    txtDescuento.Text = 0
    txtDescuento.Enabled = True
    freDescuento.Visible = False
    chkVigencia.Value = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pcancelar"))
    Unload Me
End Sub

Private Sub SSTDescuentos_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If SSTDescuentos.Tab = 0 Then
        If txtMovimientoPaciente.Visible And txtMovimientoPaciente.Enabled Then txtMovimientoPaciente.SetFocus
        FreConDescuento.Caption = "Pacientes con descuentos (I)=Interno (E)=Externo"
        pCargaGuardados "P"
    ElseIf SSTDescuentos.Tab = 1 Then
        If lstTiposPaciente.Visible And lstTiposPaciente.Enabled Then lstTiposPaciente.SetFocus
        FreConDescuento.Caption = "Tipos de paciente con descuentos"
        pCargaGuardados "T"
    Else
        If cboTipoConvenio.Visible And cboTipoConvenio.Enabled Then cboTipoConvenio.SetFocus
        FreConDescuento.Caption = "Empresas con descuentos"
        pCargaGuardados "E"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":SSTDescuentos_Click"))
    Unload Me
End Sub

Private Sub txtDescuento_GotFocus()
    On Error GoTo NotificaError
    
    vgstrEstadoManto = vgstrEstadoManto & "S"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_GotFocus"))
    Unload Me
End Sub

Private Sub txtDescuento_LostFocus()
    On Error GoTo NotificaError
    
    vgstrEstadoManto = Mid(vgstrEstadoManto, 1, 1)
    cmdGrabarRegistro.Enabled = True
    cmdVerDescuentos.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_LostFocus"))
    Unload Me
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintContador As Integer
    
    If KeyCode = vbKeyReturn Then
        If chkVigencia.Value = 0 Or fblnFechasValidas() Then
            cmdExportar.Enabled = False
            cmdImportar.Enabled = False
            Command3.Enabled = False
            Command2.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            If chkTodos.Value = 1 Then
                If optTipoDescuento(2).Value Then 'Si es por Costo
                    freDescuento.Visible = False
                    FreElementos.Enabled = True
                    freElementosIncuidos.Enabled = True
                    cmdGrabarRegistro.Enabled = True
                    cmdVerDescuentos.Enabled = True
                    For vlintContador = 0 To lstConceptos.ListCount - 1
                        lstConceptos.ListIndex = vlintContador
                        pSeleccionaElemento
                    Next
                    chkTodos.Value = 0
                    optTipoDescuento(0).Enabled = True '%
                    optTipoDescuento(1).Enabled = True '$
                    optTipoDescuento(2).Enabled = True 'Costo
                Else
                  If Val(txtDescuento.Text) > 100 Then
                    MsgBox SIHOMsg(36) & "menor o igual a 100%", vbCritical, "Mensaje"
                    txtDescuento.Text = 0
                    freDescuento.Visible = True
                    pEnfocaTextBox txtDescuento
                  Else
                    freDescuento.Visible = False
                    FreElementos.Enabled = True
                    freElementosIncuidos.Enabled = True
                    For vlintContador = 0 To lstConceptos.ListCount - 1
                        lstConceptos.ListIndex = vlintContador
                        pSeleccionaElemento
                    Next
                    chkTodos.Value = 0
                    optTipoDescuento(0).Enabled = True '%
                    optTipoDescuento(1).Enabled = True '$
                    optTipoDescuento(2).Enabled = True 'Costo
                  End If
                End If
            Else
                If optTipoDescuento(1).Value Then 'Si es por Cantidad
                    freDescuento.Visible = False
                    FreElementos.Enabled = True
                    freElementosIncuidos.Enabled = True
                    pSeleccionaElemento
                Else
                  If Val(txtDescuento.Text) > 100 Then
                    MsgBox SIHOMsg(36) & "menor o igual a 100%", vbCritical, "Mensaje"
                    txtDescuento.Text = 0
                    pEnfocaTextBox txtDescuento
                  Else
                    freDescuento.Visible = False
                    FreElementos.Enabled = True
                    freElementosIncuidos.Enabled = True
                    pSeleccionaElemento
                  End If
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_KeyDown"))
    Unload Me
End Sub
Private Function fblnFechasValidas() As Boolean
    On Error GoTo NotificaError

    fblnFechasValidas = True
    
    If chkVigencia.Value = 1 Then
        If Not IsDate(mskFechaInicioVigencia.Text) Then
            fblnFechasValidas = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            pEnfocaMkTexto mskFechaInicioVigencia
        Else
            If mskFechaInicioVigencia.Text < fdtmServerFecha Then
                '¡La fecha de inicio de vigencia debe ser mayor o igual a la del sistema.!
                MsgBox SIHOMsg(1661), vbOKOnly + vbInformation, "Mensaje"
                pEnfocaMkTexto mskFechaInicioVigencia
           Else
                If Not IsDate(mskFechaFinVigencia.Text) Then
                    fblnFechasValidas = False
                    '¡Fecha no válida!, formato de fecha dd/mm/aaaa
                    MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
                    pEnfocaMkTexto mskFechaFinVigencia
                Else
                    If CDate(mskFechaInicioVigencia.Text) > CDate(mskFechaFinVigencia.Text) Then
                        fblnFechasValidas = False
                        '¡Rango de fechas no válido!
                        MsgBox SIHOMsg(64), vbOKOnly + vbInformation, "Mensaje"
                        pEnfocaMkTexto mskFechaInicioVigencia
                    End If
                End If
            End If
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFechasValidas"))
    Unload Me
End Function

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
    Dim pbytDecimales As Byte
    On Error GoTo NotificaError
    
    If optTipoDescuento(0).Value Then 'Porcentaje
        pbytDecimales = 4
    Else
        pbytDecimales = 2
    End If
    
    If Not fblnFormatoCantidad(txtDescuento, KeyAscii, pbytDecimales) Then
       KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_KeyPress"))
    Unload Me
End Sub


Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            OptTipoPaciente(0).Value = UCase(Chr(KeyAscii)) = "I"
            OptTipoPaciente(1).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
    Unload Me
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsPostergado As ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                If OptTipoPaciente(1).Value Then 'Externos
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " Externos"
                    .optSinFacturar.Value = True
                    .optSoloActivos.Enabled = False
                    .optTodos.Enabled = False
                    .vgIntMaxRecords = 50
                    .vgstrMovCve = "M"
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                Else
                    .vgstrTipoPaciente = "I"
                    .Caption = .Caption & " Internos"
                    .optSinFacturar.Value = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = False
                    .vgIntMaxRecords = 50
                    .vgstrMovCve = "M"
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                '.vgblnPideClave = True
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                If txtMovimientoPaciente <> -1 Then
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
        Else
            If OptTipoPaciente(0).Value Then 'Internos
                vlstrSentencia = "SELECT rtrim(AdPaciente.vchApellidoPaterno)||' '||rtrim(AdPaciente.vchApellidoMaterno)||' '||rtrim(AdPaciente.vchNombre) as Nombre, " & _
                        "ccEmpresa.vchDescripcion as Empresa, AdTipoPaciente.vchDescripcion as Tipo " & _
                        "FROM AdAdmision " & _
                        "INNER JOIN AdPaciente ON " & _
                           "AdAdmision.numCvePaciente = AdPaciente.numCvePaciente " & _
                        "INNER JOIN AdTipoPaciente ON " & _
                           "AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON " & _
                           "AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa " & _
                        " INNER JOIN NODEPARTAMENTO ON ADADMISION.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where AdAdmision.bitFacturado = 0 and AdAdmision.numNumCuenta = " & txtMovimientoPaciente.Text & " And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            Else   'Externos
                vlstrSentencia = "SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as Nombre, " & _
                        "RegistroExterno.intClaveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "RegistroExterno.tnyCveTipoPaciente as cveTipoPaciente, AdTipoPaciente.vchDescripcion  as Tipo, '' as Cuarto " & _
                        "FROM RegistroExterno " & _
                        "INNER JOIN Externo ON " & _
                            "RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
                        "INNER JOIN AdTipoPaciente ON " & _
                           "RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                         "LEFT OUTER Join CcEmpresa ON " & _
                            "RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
                        " INNER JOIN NODEPARTAMENTO ON REGISTROEXTERNO.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where RegistroExterno.bitFacturado = 0 and intNumCuenta = " & txtMovimientoPaciente.Text & " And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            End If
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            Set rsPostergado = frsRegresaRs("SELECT BITPOSTERGADA FROM EXPACIENTEINGRESO WHERE INTNUMCUENTA = " & txtMovimientoPaciente.Text, adLockOptimistic, adOpenDynamic)
            If rsPostergado.RecordCount <> 0 Then
                If rsPostergado!BITPOSTERGADA = 1 Then
                    MsgBox "La cuenta se encuentra postergada, no es posible realizar cambios.", vbOKOnly + vbInformation, "Mensaje"
                    Exit Sub
                End If
            End If
            If rs.RecordCount <> 0 Then
                txtPaciente.Text = rs!Nombre
                txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                txtTipoPaciente.Text = rs!tipo
                cmdVerDescuentos.Enabled = True
                vgstrEstadoManto = "A"
                If cmdVerDescuentos.Visible And cmdVerDescuentos.Enabled Then cmdVerDescuentos.SetFocus
            Else
                'La información no existe
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                cmdVerDescuentos.Enabled = False
                pCancelar
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub

Private Sub sstElementos_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstElementos.Tab = 0 Then
        If txtSeleArticulo.Visible And txtSeleArticulo.Enabled Then txtSeleArticulo.SetFocus
        cmdCargar.Visible = False
    Else
        cmdCargar.Visible = True
    End If
    chkTodos.Enabled = sstElementos.Tab = 5

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstElementos_Click"))
    Unload Me
End Sub

Private Sub chkMedicamentos_Click()
    On Error GoTo NotificaError
    
    If txtSeleArticulo.Visible And txtSeleArticulo.Enabled Then txtSeleArticulo.SetFocus
    txtSeleArticulo_KeyUp 7, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkMedicamentos_Click"))
    Unload Me
End Sub

Private Sub optClave_Click()
    On Error GoTo NotificaError
    
    lstArticulos.Clear
    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 11
    If txtSeleArticulo.Visible And txtSeleArticulo.Enabled Then txtSeleArticulo.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optClave_Click"))
    Unload Me
End Sub

Private Sub optDescripcion_Click()
    On Error GoTo NotificaError
    
    lstArticulos.Clear
    txtSeleArticulo.Text = ""
    txtSeleArticulo.MaxLength = 30
    If txtSeleArticulo.Visible And txtSeleArticulo.Enabled Then txtSeleArticulo.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optDescripcion_Click"))
    Unload Me
End Sub

Private Sub txtSeleArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If lstArticulos.Enabled And lstArticulos.Visible Then
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
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vlstrOtroFiltro As String
    
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

Private Sub pLlenarDescuentos()
On Error GoTo NotificaError
Dim rs As New ADODB.Recordset
Dim vlstrSentencia As String
Dim vlstrCondicion As String
Dim vlintContador As Integer
    
        vgstrEstadoManto = "A"
        
        If SSTDescuentos.Tab = 0 Then
            SSTDescuentos.TabEnabled(1) = False
            SSTDescuentos.TabEnabled(2) = False
        ElseIf SSTDescuentos.Tab = 1 Then
            SSTDescuentos.TabEnabled(0) = False
            SSTDescuentos.TabEnabled(2) = False
        Else
            SSTDescuentos.TabEnabled(0) = False
            SSTDescuentos.TabEnabled(1) = False
        End If
        cmdVerDescuentos.Caption = "Cancelar asignación"
        cmdGrabarRegistro.Enabled = True
        '---------------------------------------------------------
        ' 1º Creación de la sentencia
        vlstrSentencia = "SELECT PvDescuento.chrTipoDescuento, PvDescuento.intTipoDescuento, " & _
            "PvDescuento.chrTipoCargo, PvDescuento.intCveAfectada, PvDescuento.mnyDescuento," & _
            "PvDescuento.chrTipoPaciente, PvDescuento.intCveCargo, PvDescuento.smiCveConcepto,PvDescuento.dtmFechaInicioVigencia,PvDescuento.dtmFechaFinVigencia, " & _
            "PvConceptoFacturacion.chrDescripcion AS ConceptoFacturacion, " & _
            "LaExamen.chrNombre AS Examen, " & _
            "LaGrupoExamen.chrNombre AS GrupoExamen, " & _
            "PvOtroConcepto.chrDescripcion AS OtroConcepto, " & _
            "ExCirugia.vchDescripcion AS Cirugia, " & _
            "IvArticulo.vchNombreComercial AS Articulo, " & _
            "ImEstudio.vchNombre AS Estudio, " & _
            "trim(pvPaquete.chrDescripcion) as Paquete " & _
            "FROM PvDescuento " & _
            "LEFT OUTER JOIN PvConceptoFacturacion ON PvDescuento.smiCveConcepto = PvConceptoFacturacion.smiCveConcepto " & _
            "LEFT OUTER JOIN LaExamen ON " & _
            "PvDescuento.intCveCargo = LaExamen.IntCveExamen LEFT OUTER " & _
            "Join " & _
            "LaGrupoExamen ON " & _
            "PvDescuento.intCveCargo = LaGrupoExamen.intCveGrupo LEFT OUTER "
            vlstrSentencia = vlstrSentencia & _
            "Join " & _
            "ImEstudio ON " & _
            "PvDescuento.intCveCargo = ImEstudio.intCveEstudio LEFT OUTER " & _
            "Join " & _
            "PvOtroConcepto ON " & _
            "PvDescuento.intCveCargo = PvOtroConcepto.intCveConcepto  LEFT " & _
            "Outer Join " & _
            "ExCirugia ON " & _
            "PvDescuento.intCveCargo = ExCirugia.intCveCirugia LEFT OUTER  " & _
            "Join " & _
            "IvArticulo ON " & _
            "PvDescuento.intCveCargo = IvArticulo.intIDArticulo " & _
            "Left join " & _
            "PvPaquete on PvDescuento.intCveCargo = PvPaquete.intNumPaquete "
            
            
            
        'Crear la condicion del select
        If SSTDescuentos.Tab = 0 Then
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'P' " & _
                             "and PvDescuento.intCveAfectada = " & RTrim(txtMovimientoPaciente.Text) & _
                             " and chrTipoPaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
                             
        ElseIf SSTDescuentos.Tab = 1 Then
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'T' " & _
                            "and PvDescuento.intCveAfectada = " & str(lstTiposPaciente.ItemData(lstTiposPaciente.ListIndex)) & _
                             " and chrTipoPaciente = '" & IIf(optTipoPac2(0).Value, "I", IIf(optTipoPac2(1).Value, "E", IIf(optTipoPac2(3).Value, "U", "A"))) & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
        Else
            vlstrCondicion = "where PvDescuento.chrTipoDescuento = 'E' " & _
                            "and PvDescuento.intCveAfectada = " & str(lstEmpresas.ItemData(lstEmpresas.ListIndex)) & _
                             " and chrTipoPaciente = '" & IIf(optTipoPac(0).Value, "I", IIf(optTipoPac(1).Value, "E", IIf(optTipoPac(3).Value, "U", "A"))) & "'" & _
                             " and pvdescuento.tnyclaveempresa = " & vgintClaveEmpresaContable
        End If
        
        vlstrSentencia = vlstrSentencia & vlstrCondicion
        vlstrCondicion = " and (" & _
                                "(PvConceptoFacturacion.chrDescripcion is not null and pvDescuento.chrTipoCargo = 'CF') or " & _
                                "(LaExamen.chrNombre is not null and pvDescuento.chrTipoCargo = 'EX') or " & _
                                "(LaGrupoExamen.chrNombre is not null  and pvDescuento.chrTipoCargo = 'GE') or " & _
                                "(PvOtroConcepto.chrDescripcion is not null  and pvDescuento.chrTipoCargo = 'OC') or " & _
                                "(ExCirugia.vchDescripcion is not null  and pvDescuento.chrTipoCargo = 'CI') or " & _
                                "(IvArticulo.vchNombreComercial is not null  and pvDescuento.chrTipoCargo = 'AR') or " & _
                                "(PvPaquete.chrDescripcion is not null  and pvDescuento.chrTipoCargo = 'PA') or " & _
                                "(ImEstudio.vchNombre is not null and pvDescuento.chrTipoCargo = 'ES'))"
        vlstrSentencia = vlstrSentencia & vlstrCondicion
        '---------------------------------------------------------
        ' 2º Ejecucion de la consulta
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        
        '---------------------------------------------------------
        ' 3º Cargar el Grid
        pConfiguraGridCargos
        With grdDescuentos
            .Redraw = False
            If rs.RecordCount > 0 Then
                Do While Not rs.EOF
                   If .RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    
                    Select Case rs!chrTipoCargo
                    Case "AR"
                        .TextMatrix(.Row, 1) = rs!Articulo
                    Case "CF"
                        .TextMatrix(.Row, 1) = rs!ConceptoFacturacion
                    Case "CI"
                        .TextMatrix(.Row, 1) = rs!Cirugia
                    Case "OC"
                        .TextMatrix(.Row, 1) = rs!OtroConcepto
                    Case "EX"
                        .TextMatrix(.Row, 1) = rs!Examen
                    Case "ES"
                        .TextMatrix(.Row, 1) = rs!Estudio
                    Case "GE"
                        .TextMatrix(.Row, 1) = rs!GrupoExamen
                    Case "PA"
                        .TextMatrix(.Row, 1) = rs!Paquete
                    End Select
                    
                    If rs!intTipoDescuento = 1 Then 'Porcentaje
                        .TextMatrix(.Row, 2) = rs!MNYDESCUENTO
                        .TextMatrix(.Row, 4) = "%"
                    ElseIf rs!intTipoDescuento = 0 Then 'Cantidad
                        .TextMatrix(.Row, 2) = Format(str(rs!MNYDESCUENTO), "$###,###,###.##")
                        .TextMatrix(.Row, 4) = "$"
                    Else
                        .TextMatrix(.Row, 2) = "Costo" 'Costo
                        .TextMatrix(.Row, 4) = "C"
                    End If
                    .TextMatrix(.Row, 3) = rs!chrTipoCargo
                    If IsDate(rs!dtmFechaInicioVigencia) Then
                        .TextMatrix(.Row, 5) = Format(rs!dtmFechaInicioVigencia, "dd/mmm/yyyy")
                    End If
                    If IsDate(rs!dtmFechaFinVigencia) Then
                        .TextMatrix(.Row, 6) = Format(rs!dtmFechaFinVigencia, "dd/mmm/yyyy")
                    End If
                    
                    If rs!chrTipoCargo = "CF" Then
                        .TextMatrix(.Row, 7) = rs!smicveconcepto
                    Else
                        .TextMatrix(.Row, 7) = rs!intCveCargo
                    End If
                    
                    If rs!chrTipoCargo = "CF" Then
                        .RowData(.Row) = rs!smicveconcepto
                    Else
                        .RowData(.Row) = rs!intCveCargo
                    End If
                    rs.MoveNext
                Loop
            End If
            sstElementos.Tab = 5
            cmdCargar_Click
            If lstConceptos.Enabled And lstConceptos.Visible Then lstConceptos.SetFocus
            .Redraw = True
            .Refresh
        
        End With
        FrePaciente.Enabled = False
        freTipoPaciente.Enabled = False
        FreEmpresa.Enabled = False
        FreConDescuento.Enabled = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarDescuentos"))
    Unload Me
End Sub
