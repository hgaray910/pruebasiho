VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuPrincipal 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Integral Hospitalario"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmMenuPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAlertaSumaAsegurada 
      Interval        =   30000
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer tmrPendientestimbre 
      Interval        =   800
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer tmrVerificarPendientes 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin MyFramePanel.MyFrame Frame2 
      Height          =   2925
      Left            =   120
      Top             =   1080
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   5159
      BackColor       =   16777215
      ForeColor       =   8677689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackgroundAlignment=   4
      BorderColor     =   8677689
      Caption         =   "Caja"
      CaptionAlignment=   4
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   13284230
      HeaderColorTopRight=   13284230
      HeaderColorBottomLeft=   16777215
      HeaderColorBottomRight=   16777215
      Begin MyCommandButton.MyButton cmdCajaChica 
         Height          =   975
         Left            =   630
         TabIndex        =   6
         ToolTipText     =   "Caja chica"
         Top             =   1755
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":1CFA
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":2DFC
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDescuentos 
         Height          =   975
         Left            =   2670
         TabIndex        =   8
         ToolTipText     =   "Registro de descuentos a pacientes, tipos de paciente, etc."
         Top             =   1755
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":5376
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":84BA
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdManejoCuenta 
         Height          =   975
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "Manejo de cuentas"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":AA34
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":BB36
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdHonorarios 
         Height          =   975
         Left            =   3210
         TabIndex        =   3
         ToolTipText     =   "Registro de honorarios a crédito"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":E0B0
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":F1B2
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPresupuestos 
         Height          =   975
         Left            =   1650
         TabIndex        =   7
         ToolTipText     =   "Presupuestos"
         Top             =   1755
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":1172C
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":1282E
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturacion 
         Height          =   975
         Left            =   2190
         TabIndex        =   2
         ToolTipText     =   "Facturación de cuentas"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":14DA8
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":15EAA
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCorte 
         Height          =   975
         Left            =   3690
         TabIndex        =   9
         ToolTipText     =   "Corte de caja"
         Top             =   1755
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":18424
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":1B568
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReportes 
         Height          =   975
         Left            =   4710
         TabIndex        =   10
         ToolTipText     =   "Reportes del módulo"
         Top             =   1755
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":1DAE2
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":1EBE4
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPagos 
         Height          =   975
         Left            =   1170
         TabIndex        =   1
         ToolTipText     =   "Registro de pagos y devoluciones de efectivo"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":2115E
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":22260
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCargos 
         Height          =   975
         Left            =   150
         TabIndex        =   0
         ToolTipText     =   "Cargos directos a pacientes"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":247DA
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":258DC
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdSocios 
         Height          =   975
         Left            =   5250
         TabIndex        =   5
         ToolTipText     =   "Socios"
         Top             =   675
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":27E56
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":28F58
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00846939&
         X1              =   240
         X2              =   6000
         Y1              =   480
         Y2              =   480
      End
   End
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   825
      Left            =   120
      Top             =   120
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   1455
      BackColor       =   16777215
      ForeColor       =   8677689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackgroundAlignment=   4
      BorderColor     =   8677689
      Caption         =   ""
      CaptionAlignment=   4
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   13284230
      HeaderColorTopRight=   13284230
      HeaderColorBottomLeft=   16777215
      HeaderColorBottomRight=   16777215
      Begin VB.Timer tmrDocCanceladosNoSAT 
         Interval        =   800
         Left            =   5760
         Top             =   480
      End
      Begin VB.Timer tmrCert 
         Interval        =   800
         Left            =   5760
         Top             =   0
      End
      Begin VB.Label lblNomEmpresa 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00846939&
         Height          =   555
         Left            =   240
         TabIndex        =   16
         Top             =   150
         Width           =   5775
      End
   End
   Begin MyFramePanel.MyFrame Frame3 
      Height          =   1380
      Left            =   120
      Top             =   4080
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   2434
      BackColor       =   16777215
      ForeColor       =   8677689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      BackgroundAlignment=   4
      BorderColor     =   8677689
      Caption         =   ""
      CaptionAlignment=   4
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   13284230
      HeaderColorTopRight=   13284230
      HeaderColorBottomLeft=   16777215
      HeaderColorBottomRight=   16777215
      Begin MyCommandButton.MyButton cmdRequisiciones 
         Height          =   975
         Left            =   4710
         TabIndex        =   15
         ToolTipText     =   "Requisiciones al almacén"
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":2B4D2
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":2E616
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdParametros 
         Height          =   975
         Left            =   3690
         TabIndex        =   14
         ToolTipText     =   "Parámetros del módulo"
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":30B90
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":31C92
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdListasPrecios 
         Height          =   975
         Left            =   1650
         TabIndex        =   12
         ToolTipText     =   "Registro y asignación de listas de precios"
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":3420C
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":37350
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCatalogos 
         Height          =   975
         Left            =   2670
         TabIndex        =   13
         ToolTipText     =   "Catálogos del sistema"
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":398CA
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":3A9CC
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPOS 
         Height          =   975
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Venta al público"
         Top             =   195
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   16777215
         Picture         =   "frmMenuPrincipal.frx":3CF46
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   ""
         CaptionAlignment=   7
         CaptionPosition =   4
         DepthEvent      =   1
         PictureDisabled =   "frmMenuPrincipal.frx":4008A
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
   Begin MyFramePanel.MyFrame fraFecha 
      Height          =   490
      Left            =   -180
      Top             =   7600
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   873
      BackColor       =   16777215
      ForeColor       =   8677689
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      AppearanceThemes=   2
      BackgroundAlignment=   4
      BorderColor     =   5800032
      BorderStyle     =   0
      Caption         =   ""
      CaptionAlignment=   4
      CaptionOffsetX  =   10
      CornerRadius    =   25
      CornerTopLeft   =   -1  'True
      CornerTopRight  =   -1  'True
      CornerBottomLeft=   -1  'True
      CornerBottomRight=   -1  'True
      HeaderHeight    =   32
      HeaderColorTopLeft=   16249836
      HeaderColorTopRight=   16249836
      HeaderColorBottomLeft=   13284230
      HeaderColorBottomRight=   13284230
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00846939&
         Height          =   375
         Left            =   6240
         TabIndex        =   18
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label LblAlertaSumaAsegurada 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Existen pacientes que exceden la suma asegurada, para ver la lista dar doble clic aquí"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   6240
      Width           =   6225
   End
   Begin VB.Label lblPendientesTimbre 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Existen facturas, notas de crédito, notas de cargo y donativos pendientes de timbre fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   21
      Top             =   7160
      Width           =   6225
   End
   Begin VB.Label LblDocNoCanceladosSAT 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Existen facturas, notas de crédito, notas de cargo y donativos pendientes de cancelar ante el SAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   6225
   End
   Begin VB.Label lblFechaCertificado 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "La licencia del SiHO expirará en ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   6225
   End
   Begin VB.Label lblPacientePendienteAnticipo 
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Existen pacientes pendientes de anticipo, para ver la lista dar doble clic aqui"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5880
      Width           =   6225
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAyuda 
         Caption         =   "Ayuda"
      End
      Begin VB.Menu mnu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de"
      End
   End
End
Attribute VB_Name = "frmMenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
' Pantalla principal del módulo de Caja
' Fecha de desarrollo: Abril 2001
'--------------------------------------------------------------------------------------

Dim vlstrsql As String
Dim vlblnValidoCorte As Boolean 'Bandera que indica si ya se paso por el proceso pCorte que se manda llamar en el activate
Private Const cintMinutosMensage As Integer = 10
Dim ACC As AlertaCargoCuarto.CargarCuarto
Attribute ACC.VB_VarHelpID = -1
Dim lintIntervalo As Integer
Dim ldtmHora As Date
Dim lintDepa As Integer
Public vgintSocios As Integer
Public vgdateFechaInicioOperaciones As Date ' indica la fecha de inicio de opraciones contables de la empresa, valida fechas de pagos y facturas

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

Private Sub pRevisaVencimientoPresupuestos()
On Error GoTo NotificaError
    Dim strparametro As String
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    strparametro = fRegresaParametro("IntDiasSinRespPresupuesto", "PvParametro", 0)
    If Val(strparametro) > 0 Then
        vlstrSentencia = "select * from pvpresupuesto " & _
                         "where dtmfechapresupuesto Is Not Null and chrestado = 'C' " & _
                         "and (dtmfechapresupuesto + intdiasvencimiento + " & Val(strparametro) & ") < to_date(" & fstrFechaSQL(fdtmServerFecha, "23:59:59") & ", 'YYYY-MM-DD HH24:MI:SS')"
        
        Set rs = frsRegresaRs(vlstrSentencia)
        If rs.RecordCount > 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            Do While Not rs.EOF
                '-- Modifica el estado de los presupuestos a SIN RESPUESTA de aquellos que ha vencido su fecha de espera
                vlstrSentencia = "update pvpresupuesto Set chrEstado = 'S' " & _
                                 "where intCvePresupuesto = " & rs!intcvepresupuesto
                pEjecutaSentencia (vlstrSentencia)
                Call pGuardarLogPresupuesto("S", 0, rs!intcvepresupuesto, "Cambio de estado automático a SIN RESPUESTA")
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "PRESUPUESTO (Cambio de estado a SIN RESPUESTA)", rs!intcvepresupuesto)
                rs.MoveNext
            Loop
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pRevisaVencimientoPresupuestos"))
End Sub


Private Sub cmdCajaChica_Click()
    On Error GoTo NotificaError

    frmMenuCajaChica.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCajaChica_Click"))
End Sub

Private Sub cmdCargos_Click()
    On Error GoTo NotificaError
    
    frmCargosDirectosCaja.vllngNumeroOpcion = 300
    frmCargosDirectosCaja.llngNumOpcionHabilitaCambioFecha = 3012
    frmCargosDirectosCaja.llngNumOpcionHabilitaMedicamentoAplicado = 361
    frmCargosDirectosCaja.HelpContextID = 3
    frmCargosDirectosCaja.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargos_Click"))
End Sub

Private Sub cmdCatalogos_Click()
    On Error GoTo NotificaError
    
    frmMenuCatalogosCaja.HelpContextID = 26
    frmMenuCatalogosCaja.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCatalogos_Click"))
End Sub

Private Sub cmdCorte_Click()
    On Error GoTo NotificaError
    
    frmMenuCorte.HelpContextID = 19
    frmMenuCorte.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCorte_Click"))
End Sub

Private Sub cmdDescuentos_Click()
    On Error GoTo NotificaError
    
    frmMenuDescuentos.HelpContextID = 16
    frmMenuDescuentos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDescuentos_Click"))
End Sub

Private Sub cmdEstadoCuenta_Click()
    On Error GoTo NotificaError
    
    frmReporteEstadoCuenta.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEstadoCuenta_Click"))
End Sub

Private Sub cmdFacturacion_Click()
    On Error GoTo NotificaError
    
    frmFacturacion.vllngNumeroOpcion = 304
    frmFacturacion.HelpContextID = 5
    frmFacturacion.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdFacturacion_Click"))
End Sub

Private Sub cmdHonorarios_Click()
    On Error GoTo NotificaError
    frmHonorarios.vllngNumeroOpcion = 337 'El número de opción es para saber de donde se está llamando (Caja o Crédito)
    frmHonorarios.HelpContextID = 6
    frmHonorarios.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdHonorarios_Click"))
End Sub

Private Sub cmdListasPrecios_Click()
    On Error GoTo NotificaError
    
    frmMenuListasPrecios.HelpContextID = 25
    frmMenuListasPrecios.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdListasPrecios_Click"))
End Sub

Private Sub cmdPos_Click()
    cmdPOS.Enabled = False
    frmPOS.HelpContextID = 24
    frmPOS.Show vbModal
    cmdPOS.Enabled = True
End Sub

Private Sub cmdManejoCuenta_Click()
    frmMenuManejoCuenta.HelpContextID = 7
    frmMenuManejoCuenta.Show vbModal
End Sub

Private Sub cmdPagos_Click()
    On Error GoTo NotificaError
    
    Me.Enabled = False
    frmEntradaSalidaDinero.HelpContextID = 4
    frmEntradaSalidaDinero.lblnManipulacion = False
    frmEntradaSalidaDinero.Show vbModal
    Me.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPagos_Click"))
End Sub

Private Sub cmdParametros_Click()
    On Error GoTo NotificaError
    
    frmMenuParametrosCaja.HelpContextID = 27
    frmMenuParametrosCaja.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdParametros_Click"))
End Sub

Private Sub cmdPresupuestos_Click()
    On Error GoTo NotificaError
    
    frmPresupuestos.HelpContextID = 15
    frmPresupuestos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPresupuestos_Click"))
End Sub

Private Sub cmdReportes_Click()
    On Error GoTo NotificaError
    
    frmMenuReportesPV.HelpContextID = 23
    frmMenuReportesPV.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReportes_Click"))
End Sub

Private Sub cmdRequisiciones_Click()
    On Error GoTo NotificaError

    frmMenuRequisicion.HelpContextID = 28
    frmMenuRequisicion.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRequisiciones_Click"))
End Sub

Private Sub cmdTraslado_Click()
    frmTrasladoCargos.Show vbModal
End Sub

Private Sub cmdSocios_Click()
    frmMenuSocios.Show vbModal, Me
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    '*******************************************************
    'Creación del corte del departamento con el que se entro
    '*******************************************************
    
    vgstrNombreForm = Me.Name
    
    If Not vlblnValidoCorte Then
        pCorte
    End If
    
    fblnCargaPermisos
    fblnHabilitaObjetos Me
    pSocios

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub pCorte()
    On Error GoTo NotificaError
    
    vlblnValidoCorte = True
    
    frsEjecuta_SP CStr(vgintNumeroDepartamento) & "|" & CStr(vglngNumeroEmpleado), "SP_GNINSCORTE"
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCorte"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim strparametro As String
    Dim rsAnoInicio As ADODB.Recordset
    Dim rsMesInicio As ADODB.Recordset
    Dim vlTopAlerta As Long
    Dim ObjRS As New ADODB.Recordset
    Dim ObjStr As String
    
1    vgstrNombreForm = Me.Name
    
2    fraFecha.Top = 5520
3    Me.Height = 6480
    
4    vgintNumeroModulo = 2
5    vlblnValidoCorte = False
    
6    strparametro = fRegresaParametro("dtmHoraIniMsgCargo", "PvParametro", 0)
7    If IsDate(strparametro) Then ldtmHora = fRegresaParametro("dtmHoraIniMsgCargo", "PvParametro", 0)
8    strparametro = fRegresaParametro("smiCveDepartamentoMsg", "PvParametro", 0)
9    If strparametro <> "" Then lintDepa = fRegresaParametro("smiCveDepartamentoMsg", "PvParametro", 0)
10    strparametro = fRegresaParametro("intIntervaloMsgCargo", "PvParametro", 0)
11    If strparametro <> "" Then lintIntervalo = fRegresaParametro("intIntervaloMsgCargo", "PvParametro", 0)
    
12    If (vgintNumeroDepartamento = lintDepa And lintIntervalo > 0 And Not IsNull(ldtmHora)) Then
13        Set ACC = New AlertaCargoCuarto.CargarCuarto
14        ACC.pInicia "PV", fstrRegresaConeccion, vgstrBaseDatosUtilizada, App.Path
15    End If
        
16    lblNomEmpresa = Trim(vgstrNombreHospitalCH)
17    fraFecha.Caption = Format(fdtmServerFecha, "Long Date")
    
18    pVerificarPendientes
19    pCerrarCuentasAut
20    pParametrosCuartos
21    pParametrosAdmision
22    pSocios
23    pDocNoCanceladosSAT
24    pDocPendientesTimbre
        
25    tmrAlertaSumaAsegurada.Enabled = False
26    LblAlertaSumaAsegurada.Caption = ""
27    ObjStr = "select vchvalor from siparametro where vchnombre ='BITALERTASUMAASEGURADA' and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
28    Set ObjRS = frsRegresaRs(ObjStr, adLockOptimistic)
29    If ObjRS.RecordCount <> 0 Then
30       If ObjRS!vchvalor = "1" Then
31            pVerificaAlertaSumaAsegurada
32            tmrAlertaSumaAsegurada.Enabled = True
33       End If
34    End If
    
   'cagar fecha de inicio de operaciones contables de la empresa
35    Set rsAnoInicio = frsSelParametros("CN", vgintClaveEmpresaContable, "SMIEJERCICIOINICIOOPERACIONES")
36    Set rsMesInicio = frsSelParametros("CN", vgintClaveEmpresaContable, "TNYMESINICIOOPERACIONES")
37    vgdateFechaInicioOperaciones = CDate("01/" & rsMesInicio!valor & "/" & rsAnoInicio!valor)
    
    'Descarga el certificado para validar las fechas
38    lblFechaCertificado.Caption = fstrTiempoRestanteCertificado

    'Se ocultan las etiquetas que están vacías
39    lblFechaCertificado.Visible = IIf(Trim(lblFechaCertificado.Caption) = "", False, True)
40    lblPacientePendienteAnticipo.Visible = IIf(Trim(lblPacientePendienteAnticipo.Caption) = "", False, True)
41    Me.LblDocNoCanceladosSAT.Visible = IIf(Trim(Me.LblDocNoCanceladosSAT.Caption) = "", False, True)
42    Me.lblPendientesTimbre.Visible = IIf(Trim(Me.lblPendientesTimbre.Caption) = "", False, True)
43    Me.LblAlertaSumaAsegurada.Visible = IIf(Trim(Me.LblAlertaSumaAsegurada.Caption) = "", False, True)
    
44    pAlertasYRedimiencion
45    pRevisaVencimientoPresupuestos
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load" & " Linea:" & Erl()))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (vgintNumeroDepartamento = lintDepa And lintIntervalo <> 0) Then ACC.pTermina
        Unload frmfondo
    End
End Sub

Private Sub Label1_Click()
    frmAcerca.Show vbModal, Me
End Sub

Private Sub LblAlertaSumaAsegurada_Click()
    If LblAlertaSumaAsegurada.Caption <> "" Then
        frmAlertaSumaAsegurada.vlblnPermiteAbrir = fblnRevisaPermiso(vglngNumeroLogin, 4130, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4130, "E", True)
        frmAlertaSumaAsegurada.vlblnPermiteCerrar = fblnRevisaPermiso(vglngNumeroLogin, 4129, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4129, "E", True)
                
        frmAlertaSumaAsegurada.Show vbModal, Me
        pVerificaAlertaSumaAsegurada
    End If
End Sub

Private Sub LblDocNoCanceladosSAT_DblClick()
    Me.tmrDocCanceladosNoSAT.Enabled = False
    If Me.LblDocNoCanceladosSAT.Visible = True Then Me.LblDocNoCanceladosSAT.Visible = False
    LblDocNoCanceladosSAT.Caption = ""
    pAlertasYRedimiencion
End Sub

Private Sub lblPendientesTimbre_DblClick()
    Me.tmrPendientestimbre.Enabled = False
    If Me.lblPendientesTimbre.Visible = True Then Me.lblPendientesTimbre.Visible = False
    lblPendientesTimbre.Caption = ""
    pAlertasYRedimiencion
End Sub

Private Sub mnuAcerca_Click()
    frmAcerca.Show vbModal, Me
End Sub

Private Sub mnuAyuda_Click()
    pMostrarAyuda App.HelpFile
End Sub

Private Sub lblPacientePendienteAnticipo_DblClick()
    pVerificarPendientes
    If lblPacientePendienteAnticipo.Caption <> "" Then
        frmPendientesAnticipo.Show vbModal, Me
    End If
End Sub

Private Sub tmrCert_Timer()
    If Trim(lblFechaCertificado.Caption) <> "" Then
        lblFechaCertificado.Visible = Not lblFechaCertificado.Visible
    End If
End Sub

Private Sub tmrDocCanceladosNoSAT_Timer()
    If Trim(Me.LblDocNoCanceladosSAT.Caption) <> "" Then
       LblDocNoCanceladosSAT.Visible = Not LblDocNoCanceladosSAT.Visible
    End If
End Sub

Private Sub tmrPendientestimbre_Timer()
    If Trim(Me.lblPendientesTimbre.Caption) <> "" Then
       lblPendientesTimbre.Visible = Not lblPendientesTimbre.Visible
    End If
End Sub

Private Sub tmrVerificarPendientes_Timer()
    On Error GoTo NotificaError
    Static intMinutos As Integer
    intMinutos = intMinutos + 1
    If intMinutos = cintMinutosMensage Then
        intMinutos = 0
        pVerificarPendientes
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrVerificarPendientes_Timer"))
End Sub

Private Sub tmrAlertaSumaAsegurada_Timer()
    On Error GoTo NotificaError
    
    pVerificaAlertaSumaAsegurada
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrAlertaSumaAsegurada_Timer"))
End Sub

Private Sub pVerificaAlertaSumaAsegurada()
    On Error GoTo NotificaError
    
    Dim rsCuentas As ADODB.Recordset
    
    Set rsCuentas = frsEjecuta_SP(Str(vgintClaveEmpresaContable), "sp_PVRPTCUENTAPENDIENTEALERTA")
    If rsCuentas.RecordCount <> 0 Then
        If LblAlertaSumaAsegurada.Caption = "" Then
            LblAlertaSumaAsegurada.Caption = "Existen pacientes que exceden la suma asegurada, para ver la lista dar doble clic aquí"
            LblAlertaSumaAsegurada.Visible = True
            pAlertasYRedimiencion
        Else
            LblAlertaSumaAsegurada.Caption = "Existen pacientes que exceden la suma asegurada, para ver la lista dar doble clic aquí"
            LblAlertaSumaAsegurada.Visible = True
        End If
    Else
        If LblAlertaSumaAsegurada.Caption <> "" Then
            LblAlertaSumaAsegurada.Caption = ""
            pAlertasYRedimiencion
        Else
            LblAlertaSumaAsegurada.Caption = ""
        End If
    End If
    rsCuentas.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pVerificaAlertaSumaAsegurada"))
End Sub

Private Sub pVerificarPendientes()
    On Error GoTo NotificaError

    If frmPendientesAnticipo.fblnVerificarPendientes Then
        If lblPacientePendienteAnticipo.Caption = "" Then
            lblPacientePendienteAnticipo.Caption = "Existen pacientes pendientes de anticipo, para ver la lista dar doble clic aquí"
            pAlertasYRedimiencion
        Else
            lblPacientePendienteAnticipo.Caption = "Existen pacientes pendientes de anticipo, para ver la lista dar doble clic aquí"
        End If
    Else
        If lblPacientePendienteAnticipo.Caption <> "" Then
            lblPacientePendienteAnticipo.Caption = ""
            pAlertasYRedimiencion
        Else
            lblPacientePendienteAnticipo.Caption = ""
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pVerificarPendientes"))
End Sub

Private Sub pCerrarCuentasAut()
    On Error GoTo NotificaError
    Dim intDias As Integer
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select bitCerrarCuentasExtAut, intDiasAbrirCuentasExternos from PVParametro where tnyclaveempresa = " & vgintClaveEmpresaContable)
    If Not rs.EOF Then
        If Not IsNull(rs!bitCerrarCuentasExtAut) And Not IsNull(rs!intDiasAbrirCuentasExternos) Then
            If rs!bitCerrarCuentasExtAut <> 0 Then
                intDias = rs!intDiasAbrirCuentasExternos
                frsEjecuta_SP intDias & "|" & vglngNumeroLogin, "sp_PVCierreAutomaticoCuentas", True
            End If
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCerrarCuentasAut"))
End Sub

Private Sub pSocios()
    On Error GoTo NotificaError
    'este procedimiento permite
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITUTILIZASOCIOS'")
    If Not rs.EOF Then
        vgintSocios = rs!vchvalor
        If rs!vchvalor = 1 Then
            frmMenuPrincipal.Width = 6705
            Frame1.Width = 6330
            lblNomEmpresa.Left = 300
            lblNomEmpresa.Top = 120
            Frame2.Width = 6330
            cmdCargos.Left = 150
            cmdPagos.Left = 1170
            cmdFacturacion.Left = 2190
            cmdHonorarios.Left = 3210
            cmdManejoCuenta.Left = 4230
            cmdCajaChica.Left = 630
            cmdPresupuestos.Left = 1650
            cmdDescuentos.Left = 2670
            cmdCorte.Left = 3690
            cmdReportes.Left = 4710
            Frame3.Width = 6330
            cmdPOS.Left = 630
            cmdListasPrecios.Left = 1650
            cmdCatalogos.Left = 2670
            cmdParametros.Left = 3690
            cmdRequisiciones.Left = 4710
            cmdSocios.Visible = True
        Else
            frmMenuPrincipal.Width = 6525
            Frame1.Height = 825
            Frame1.Left = 120
            Frame1.Top = 120
            Frame1.Width = 6210
            lblNomEmpresa.Left = 240
            lblNomEmpresa.Top = 150
            Frame2.Width = 6210
            cmdCargos.Left = 630
            cmdPagos.Left = 1650
            cmdFacturacion.Left = 2670
            cmdHonorarios.Left = 3690
            cmdManejoCuenta.Left = 4710
            cmdCajaChica.Left = 630
            cmdPresupuestos.Left = 1650
            cmdDescuentos.Left = 2670
            cmdCorte.Left = 3690
            cmdReportes.Left = 4710
            Frame3.Width = 6210
            cmdPOS.Left = 630
            cmdListasPrecios.Left = 1650
            cmdCatalogos.Left = 2670
            cmdParametros.Left = 3690
            cmdRequisiciones.Left = 4710
            cmdSocios.Visible = False
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSocios"))
End Sub

Private Sub pDocNoCanceladosSAT()
    On Error GoTo NotificaError
    
    Dim ObjRS As New ADODB.Recordset
    Dim ObjStr As String
    Dim ObjInt As Integer
    Dim ObjDato As String

    Set ObjRS = frsEjecuta_SP(CStr(vgintNumeroDepartamento), "sp_pvSelNoCanceladosSAT")
    ObjStr = ""
    
    If ObjRS.RecordCount > 0 Then
    ObjRS.MoveFirst
     For ObjInt = 0 To ObjRS.RecordCount - 1
     ObjDato = ObjRS(0)
        Select Case ObjDato
        Case "FA"
        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "facturas", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y facturas", ObjStr & ", facturas"))
        Case "CR"
        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "notas de crédito", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y notas de crédito", ObjStr & ", notas de crédito"))
        Case "CA"
        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "notas de cargo", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y notas de cargo", ObjStr & ", notas de cargo"))
        Case "DO"
        If cgstrModulo = "CN" Then ObjStr = IIf(ObjStr = "", "donativos", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y donativos", ObjStr & ", donativos"))
        End Select
     ObjRS.MoveNext
     Next ObjInt
     
    End If
    If ObjStr = "" Then
       Me.LblDocNoCanceladosSAT.Caption = ""
    Else
       Me.LblDocNoCanceladosSAT.Caption = "Existen " & ObjStr & " pendientes de cancelar ante el SAT"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pDocNoCanceladosSAT"))
End Sub

Private Sub pDocPendientesTimbre()
    On Error GoTo NotificaError

    Dim ObjRS As New ADODB.Recordset
    Dim ObjStr As String
    Dim ObjInt As Integer
    Dim ObjDato As String

    Set ObjRS = frsEjecuta_SP(CStr(vgintNumeroDepartamento), "sp_pvSelPendientesTimbre")
    ObjStr = ""
    
    If ObjRS.RecordCount > 0 Then
       ObjRS.MoveFirst
       For ObjInt = 0 To ObjRS.RecordCount - 1
           ObjDato = ObjRS(0)
           Select Case ObjDato
                  Case "FA"
                        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "facturas", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y facturas", ObjStr & ", facturas"))
                  Case "CR"
                        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "notas de crédito", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y notas de crédito", ObjStr & ", notas de crédito"))
                  Case "CA"
                        If cgstrModulo = "CC" Or cgstrModulo = "PV" Then ObjStr = IIf(ObjStr = "", "notas de cargo", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y notas de cargo", ObjStr & ", notas de cargo"))
                  Case "DO"
                        If cgstrModulo = "CN" Then ObjStr = IIf(ObjStr = "", "donativos", IIf((ObjInt + 1 = ObjRS.RecordCount), ObjStr & " y donativos", ObjStr & ", donativos"))
           End Select
           ObjRS.MoveNext
       Next ObjInt
    End If
    
    If ObjStr = "" Then
       Me.lblPendientesTimbre.Caption = ""
    Else
       Me.lblPendientesTimbre.Caption = "Existen " & ObjStr & " pendientes de timbre fiscal"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pDocPendientesTimbre"))
End Sub

Private Sub pAlertasYRedimiencion()
    On Error GoTo NotificaError
        
        '--- LOGICA PARA MOSTRAR LA ALERTA E IR AUMENTANDO LA PANTALLA DE TAMAÑO ---
        'Inicia la pantalla como si no hubiera ninguna alerta
    
        Me.AutoRedraw = False
    
        fraFecha.Top = 5520
        Me.Height = 6480

        vlTopAlerta = 5520 ' Top de inicio para las alertas
        vlAumentoSencillo = 360 ' Tamaño de aumento cuando se agregue una alerta que ocupe un solo renglon
        vlAumentoDoble = 500 ' Tamaño de aumento cuando se agregue una alerta que ocupe dos renglones

        'Se obtiene el tiempo restante de arrendamiento en caso de aplicar (para mostrar la notificación en el menú principal (reemplazando la etiqueta de la vigencia de certificados)
        If Trim(frmLogin.strTiempoRestanteTotal) <> "" Then 'Si es de arrendamiento
            If lblFechaCertificado.Visible = False Then
                lblFechaCertificado.Visible = True
            End If
            lblFechaCertificado.Caption = "La licencia del SiHO expirará en " & frmLogin.strTiempoRestanteTotal & IIf(Val(frmLogin.strTiempoRestanteTotal) = 1, " día.", " días.")
        End If
        
        'La licencia del SiHO expirará en ...
        If Trim(lblFechaCertificado.Caption) <> "" Then
            lblFechaCertificado.Top = vlTopAlerta
            fraFecha.Top = fraFecha.Top + vlAumentoSencillo
            Me.Height = Me.Height + vlAumentoSencillo
            
            vlTopAlerta = vlTopAlerta + vlAumentoSencillo
        End If
        
        'Existen pacientes pendientes de anticipo, para ver la lista dar doble clic aqui
        If Trim(lblPacientePendienteAnticipo.Caption) <> "" Then
            lblPacientePendienteAnticipo.Top = vlTopAlerta
            fraFecha.Top = fraFecha.Top + vlAumentoSencillo
            Me.Height = Me.Height + vlAumentoSencillo
            
            vlTopAlerta = vlTopAlerta + vlAumentoSencillo
        End If
        
        'Existen pacientes que exceden la suma asegurada, para ver la lista dar doble clic aqui
        If Trim(LblAlertaSumaAsegurada.Caption) <> "" Then
            LblAlertaSumaAsegurada.Top = vlTopAlerta
            fraFecha.Top = fraFecha.Top + vlAumentoSencillo
            Me.Height = Me.Height + vlAumentoSencillo
            
            vlTopAlerta = vlTopAlerta + vlAumentoSencillo
        End If
        
        'Existen facturas, notas de crédito, notas de cargo y donativos pendientes de cancelar ante el SAT
        If Trim(LblDocNoCanceladosSAT.Caption) <> "" Then
            
            LblDocNoCanceladosSAT.Top = vlTopAlerta
            If Len(Trim(LblDocNoCanceladosSAT.Caption)) > 70 Then
                fraFecha.Top = fraFecha.Top + vlAumentoDoble
                Me.Height = Me.Height + vlAumentoDoble
                
                vlTopAlerta = vlTopAlerta + vlAumentoDoble
            Else
                fraFecha.Top = fraFecha.Top + vlAumentoSencillo
                Me.Height = Me.Height + vlAumentoSencillo
                
                vlTopAlerta = vlTopAlerta + vlAumentoSencillo
            End If
        End If
    
        'Existen facturas, notas de crédito, notas de cargo y donativos pendientes de timbre fiscal
        If Trim(lblPendientesTimbre.Caption) <> "" Then
            lblPendientesTimbre.Top = vlTopAlerta
            If Len(Trim(lblPendientesTimbre.Caption)) > 70 Then
                fraFecha.Top = fraFecha.Top + vlAumentoDoble
                Me.Height = Me.Height + vlAumentoDoble
                
                vlTopAlerta = vlTopAlerta + vlAumentoDoble
            Else
                fraFecha.Top = fraFecha.Top + vlAumentoSencillo
                Me.Height = Me.Height + vlAumentoSencillo
                
                vlTopAlerta = vlTopAlerta + vlAumentoSencillo
            End If
        End If
    
    If Trim(Me.LblDocNoCanceladosSAT.Caption) = "" Then Me.tmrDocCanceladosNoSAT.Enabled = False
    If Trim(Me.lblPendientesTimbre.Caption) = "" Then Me.tmrPendientestimbre.Enabled = False
    
    Me.AutoRedraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAlertasYRedimiencion"))
End Sub
