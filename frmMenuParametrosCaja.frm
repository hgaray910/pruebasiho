VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuParametrosCaja 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de parámetros "
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "frmMenuParametrosCaja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   6390
      Left            =   195
      Top             =   90
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11271
      BackColor       =   16777215
      ForeColor       =   11682635
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
      Begin MyCommandButton.MyButton cmdValidaEdoCta 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         ToolTipText     =   "Conexión para validación de estado de cuenta"
         Top             =   2805
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Conexión para validación de estado de cuenta"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFormatosTicket 
         Height          =   315
         Left            =   5115
         TabIndex        =   17
         ToolTipText     =   "Mantenimiento de formatos de ticket"
         Top             =   1500
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Formatos de ticket"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAsignacionCF 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         ToolTipText     =   "Conceptos para facturas parciales"
         Top             =   1935
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Conceptos de facturas parciales"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdEquivalenciaFormaPago 
         Height          =   315
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "Equivalencias de forma de pago"
         Top             =   4980
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Equivalencias de forma de pago"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdParametroTipoCorte 
         Height          =   315
         Left            =   5115
         TabIndex        =   24
         ToolTipText     =   "Parámetros del tipo de corte"
         Top             =   4545
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Tipo de corte por departamento"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdExclusionDescuento 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Exclusión de descuentos por departamento"
         Top             =   5415
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Exclusión de descuentos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDeptosPrecioEspecial 
         Height          =   315
         Left            =   150
         TabIndex        =   10
         ToolTipText     =   "Departamentos con precio especial"
         Top             =   4545
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Departamentos con precio especial"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPreciosEspeciales 
         Height          =   315
         Left            =   5115
         TabIndex        =   22
         ToolTipText     =   "Parámetros para precios especiales por horario"
         Top             =   3675
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Precios especiales por horario"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdParametrosModulo 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         ToolTipText     =   "Mantenimiento de parámetros del módulo"
         Top             =   630
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Parámetros del módulo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAsignaPermiso 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         ToolTipText     =   "Asignación de permisos de acceso a las opciones del módulo"
         Top             =   225
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         MaskColor       =   16777215
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Asignación de permisos de acceso"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoFormatoFactura 
         Height          =   315
         Left            =   5115
         TabIndex        =   14
         ToolTipText     =   "Mantenimiento de formatos de factura"
         Top             =   225
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Formatos de factura"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoTipoCambio 
         Height          =   315
         Left            =   5115
         TabIndex        =   25
         ToolTipText     =   "Mantenimiento del tipo de cambio por día"
         Top             =   4980
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Tipos de cambio"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoFolios 
         Height          =   315
         Left            =   150
         TabIndex        =   13
         ToolTipText     =   "Mantenimiento de folios de documentos"
         Top             =   5850
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Folios de documentos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdDeptoCajaChica 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         ToolTipText     =   "Departamentos con caja chica"
         Top             =   4110
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Departamentos con caja chica"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCostos 
         Height          =   315
         Left            =   150
         TabIndex        =   8
         ToolTipText     =   "Mantenimiento de los costos de conceptos de cargo y paquetes"
         Top             =   3675
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Costos por conceptos de cargo y paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoNombreMed 
         Height          =   315
         Left            =   5115
         TabIndex        =   20
         ToolTipText     =   "Parametrizar el nombre de los medicamentos"
         Top             =   2805
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Nombre genérico por tipo de paciente/convenio"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAlmacenesPOS 
         Height          =   315
         Left            =   5115
         TabIndex        =   18
         ToolTipText     =   "Selecciona el almacén que surtirá venta al público"
         Top             =   1935
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Manejo almacenes venta público"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFormatoFacturaEmpresaPac 
         Height          =   315
         Left            =   5115
         TabIndex        =   15
         ToolTipText     =   "Mantenimiento de formatos de factura por tipo de paciente y empresa"
         Top             =   630
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Formatos de factura por tipo de paciente y empresa"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMedicoPromotor 
         Height          =   315
         Left            =   5115
         TabIndex        =   19
         ToolTipText     =   "Selecciona el almacén que surtirá venta al público"
         Top             =   2370
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Médicos por promotor"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdConceptosExcluidosAreaProductividad 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         ToolTipText     =   "Mantenimiento de formatos de ticket"
         Top             =   2370
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Conceptos excluidos por área de productividad"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoUnidadesMedida 
         Height          =   315
         Left            =   5115
         TabIndex        =   27
         ToolTipText     =   "Mantenimiento de las unidades de medida por concepto de facturación"
         Top             =   5850
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Unidades de medida para conceptos de factura"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPrecioscargospaquetes 
         Height          =   315
         Left            =   5115
         TabIndex        =   21
         ToolTipText     =   "Parámetros para precios de cargo en paquetes"
         Top             =   3240
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Precios de cargos en paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoTiposPacienteConvenio 
         Height          =   315
         Left            =   5115
         TabIndex        =   26
         ToolTipText     =   "Mantenimiento de los tipos de paciente y convenio para factura de asistencia social"
         Top             =   5415
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Tipos de paciente y convenio para factura de asistencia social"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdTabuladorPuntos 
         Height          =   315
         Left            =   5115
         TabIndex        =   23
         ToolTipText     =   "Configuración del tabulador de puntos para cliente leal"
         Top             =   4110
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Tabulador de puntos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdConceptosDesglosa 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Conceptos agrupa cargos factura mixta"
         Top             =   1500
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Conceptos agrupa cargos factura mixta"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdConfiguracionUnicoConcepto 
         Height          =   315
         Left            =   150
         TabIndex        =   7
         ToolTipText     =   "Configuración para facturación con un concepto"
         Top             =   3240
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Configuración para facturación con un concepto"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdMantoFormatoRecibo 
         Height          =   315
         Left            =   5115
         TabIndex        =   16
         ToolTipText     =   "Mantenimiento de formatos de recibos"
         Top             =   1065
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Formatos de recibo"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdComisionesPromotores 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Mantenimiento de formatos de recibos"
         Top             =   1065
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   556
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
         AppearanceThemes=   1
         BackColorDown   =   15133676
         TransparentColor=   16249836
         Caption         =   "Comisiones para promotores"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuParametrosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlmacenesPOS_Click()
    frmAlmacenPOS.Show vbModal
End Sub

Private Sub cmdAsignacionCF_Click()
    frmMantoAsignacionCF.HelpContextID = 27
    frmMantoAsignacionCF.Show vbModal, Me
End Sub

Private Sub cmdAsignaPermiso_Click()
    pFormaPermisos 335
End Sub

Private Sub cmdComisionesPromotores_Click()
    frmComisionPromotor.llngNumOpcion = 2298
    frmComisionPromotor.HelpContextID = 27
    frmComisionPromotor.Show vbModal
End Sub

Private Sub cmdConceptosDesglosa_Click()
    frmConfiguracionDesglosaC.Show vbModal
    
End Sub

Private Sub cmdConceptosExcluidosAreaProductividad_Click()
    frmConceptosExcluidosAreaP.llngNumOpcion = 2299
    frmConceptosExcluidosAreaP.HelpContextID = 27
    frmConceptosExcluidosAreaP.Show vbModal
End Sub

Private Sub cmdConfiguracionUnicoConcepto_Click()
    frmConfiguracionFUC.Show vbModal
End Sub

Private Sub cmdCostos_Click()
    frmCostoCargos.lstrModulo = "PV"
    frmCostoCargos.Show vbModal, Me
End Sub

Private Sub cmdDeptoCajaChica_Click()
    frmDepartamentoCajachica.Show vbModal
End Sub

Private Sub cmdDeptosPrecioEspecial_Click()
    frmDepartamentoPrecioEspecial.HelpContextID = 27
    frmDepartamentoPrecioEspecial.Show vbModal, Me
End Sub

Private Sub cmdEquivalenciaFormaPago_Click()
    frmEquivalenciaFormaPago.HelpContextID = 27
    frmEquivalenciaFormaPago.Show vbModal
End Sub

Private Sub cmdExclusionDescuento_Click()
    frmParametroExclusionDescuento.HelpContextID = 27
    frmParametroExclusionDescuento.Show vbModal
End Sub

Private Sub cmdFormatoFacturaEmpresaPac_Click()
    frmFormatoFacturaPacEmpresa.llngNumeroOpcionModulo = IIf(cgstrModulo = "SI", 2284, 2283)
    frmFormatoFacturaPacEmpresa.Show vbModal
End Sub

Private Sub cmdFormatosTicket_Click()
    frmParametrizacionTicket.Show vbModal
End Sub

Private Sub cmdMantoFolios_Click()
    frmMantoFolios.HelpContextID = 26
    frmMantoFolios.Show vbModal
End Sub

Private Sub cmdMantoFormatoFactura_Click()
    vglngNumeroTipoFormato = 2
    frmMantenimientoFormatos.vllngNumeroOpcionModulo = IIf(cgstrModulo = "SI", 1185, 332)
    frmMantenimientoFormatos.HelpContextID = 27
    frmMantenimientoFormatos.Show vbModal
End Sub

Private Sub cmdMantoFormatoRecibo_Click()
    frmFormatosRecibos.Show vbModal, Me
End Sub

Private Sub cmdMantoNombreMed_Click()
    frmMantoNombreMed.Show vbModal
End Sub

Private Sub cmdMantoTipoCambio_Click()
    frmMantoTipoCambio.vllngNumeroOpcionModulo = IIf(cgstrModulo = "SI", 1189, 331)
    frmMantoTipoCambio.HelpContextID = 27
    frmMantoTipoCambio.Show vbModal
End Sub

Private Sub cmdMantoTiposPacienteConvenio_Click()
    
    frmPacientesConvenios.llngNumOpcion = IIf(cgstrModulo = "SI", 4121, 4120)
    frmPacientesConvenios.HelpContextID = 27
    frmPacientesConvenios.Show vbModal
    
End Sub

'- (CR) Agregado para caso 6911 -'
Private Sub cmdMantoUnidadesMedida_Click()
    frmUnidadesMedida.vllngNumeroOpcion = flngObtenOpcion(cmdMantoUnidadesMedida.Name)
    frmUnidadesMedida.Show vbModal
End Sub

Private Sub cmdMedicoPromotor_Click()
    frmMedicosPorPromotor.llngNumOpcion = 2300
    frmMedicosPorPromotor.HelpContextID = 27
    frmMedicosPorPromotor.Show vbModal
End Sub

Private Sub cmdParametrosModulo_Click()
    frmParametros.HelpContextID = 27
    frmParametros.Show vbModal
End Sub

Private Sub cmdParametroTipoCorte_Click()
    frmParametroTipoCorte.HelpContextID = 27
    frmParametroTipoCorte.Show vbModal
End Sub

Private Sub cmdPrecioscargospaquetes_Click() ''*
    frmPrecioCargosPaquete.HelpContextID = 27
    frmPrecioCargosPaquete.Show vbModal
End Sub

Private Sub cmdPreciosEspeciales_Click()
    frmPrecioHorario.HelpContextID = 27
    frmPrecioHorario.Show vbModal
End Sub

Private Sub cmdTabuladorPuntos_Click()
    frmTabuladorPuntos.llngNumOpcion = IIf(cgstrModulo = "PV", 4136, 4136)
    frmTabuladorPuntos.HelpContextID = 27
    frmTabuladorPuntos.Show vbModal
End Sub

Private Sub cmdValidaEdoCta_Click()
    frmConValidaEdoCta.HelpContextID = 27
    frmConValidaEdoCta.Show vbModal
    
End Sub

Private Sub Form_Activate()
    fblnCargaPermisos  '***
    fblnHabilitaObjetos Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    
    '- Validación para que no aparezca el botón de unidades de medida en el Hospital San José de Hermosillo -'
    'If Replace(Replace(UCase(vgstrRfCCH), "-", ""), " ", "") = "HSJ040622G70" Then
    '    cmdMantoUnidadesMedida.Visible = False
    'End If
End Sub

Private Sub MyButton1_Click()

End Sub

