VERSION 5.00
Object = "{6B45E0EA-D03D-4CBB-94F4-B6AD155551A1}#1.1#0"; "MyFramePanel.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmMenuReportesPV 
   BackColor       =   &H00F7F3EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menú de reportes"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8595
   Icon            =   "frmMenuReportesPV.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin MyFramePanel.MyFrame Frame1 
      Height          =   8475
      Left            =   120
      Top             =   120
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   14949
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
      Begin MyCommandButton.MyButton cmdCuentasReabiertas 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Reporte de cuentas que excedieron el límite máximo de días para abrir"
         Top             =   3090
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cuentas reabiertas"
         DepthEvent      =   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdVentasCredito 
         Height          =   315
         Left            =   4230
         TabIndex        =   37
         ToolTipText     =   "Reporte de ventas a crédito"
         Top             =   7140
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ventas a crédito"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteador 
         Height          =   315
         Left            =   4230
         TabIndex        =   39
         ToolTipText     =   "Reporteador"
         Top             =   7950
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Reporteador"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdIngresos 
         Height          =   315
         Left            =   4230
         TabIndex        =   26
         ToolTipText     =   "Reporte de ingresos por tipo de paciente"
         Top             =   2685
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por tipo de paciente"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturasCanceladas 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Reporte de facturas canceladas en el departamento"
         Top             =   5925
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Facturas canceladas en el departamento"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPacientesAtendidos 
         Height          =   315
         Left            =   4230
         TabIndex        =   28
         ToolTipText     =   "Reporte de pacientes atendidos"
         Top             =   3495
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Pacientes atendidos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteDescuentos 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Reporte de descuentos asignados"
         Top             =   3495
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Descuentos asignados"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCargosempresas 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Reporte de cargos facturados por procedencia"
         Top             =   1065
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cargos facturados por procedencia"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteProductividadMedicos 
         Height          =   315
         Left            =   4230
         TabIndex        =   31
         ToolTipText     =   "Reporte de productividad de médicos"
         Top             =   4710
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Productividad de médicos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdEstadoCuenta 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Reporte del estado de cuenta del paciente"
         Top             =   4710
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Estado de cuenta del paciente"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReportePresupuestos 
         Height          =   315
         Left            =   4230
         TabIndex        =   30
         ToolTipText     =   "Reporte de presupuestos realizados"
         Top             =   4305
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Presupuestos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCuentaPendiente 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Reporte de cuentas pendientes de facturar"
         Top             =   2685
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cuentas pendientes de facturar"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptDescuentosAsignados 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Reporte de descuentos asignados por tipo de paciente"
         Top             =   3900
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Descuentos por tipo de paciente"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptRelacionFacturas 
         Height          =   315
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Reporte relación de facturas"
         Top             =   5520
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Facturas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteGanancia 
         Height          =   315
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "Reporte de ganancias"
         Top             =   7140
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ganancias"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFacturasPorIngreso 
         Height          =   315
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "Reporte de facturas por lugar de ingreso al hospital"
         Top             =   6330
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Facturas por lugar de ingreso al hospital"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdReporteListaPrecios 
         Height          =   315
         Left            =   4230
         TabIndex        =   27
         ToolTipText     =   "Reporte de listas de precios"
         Top             =   3090
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Listas de precios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptHonorariosPagados 
         Height          =   315
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Honorarios pagados en efectivo que no entraron al corte de caja"
         Top             =   7545
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Honorarios en efectivo fuera del corte"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptTraslado 
         Height          =   315
         Left            =   4230
         TabIndex        =   36
         ToolTipText     =   "Reporte de traslado de cargos"
         Top             =   6735
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Traslado de cargos"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCjaIngresosTurno 
         Height          =   315
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "Reporte de ingresos"
         Top             =   7950
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos de caja por turno"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptIngresoConcepto 
         Height          =   315
         Left            =   4230
         TabIndex        =   21
         ToolTipText     =   "Reporte de ingresos por concepto de factura"
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por concepto de factura"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdIngresosTickets 
         Height          =   315
         Left            =   4230
         TabIndex        =   25
         ToolTipText     =   "Reporte de ingresos por tickets"
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por tickets"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdingresosempresa 
         Height          =   315
         Left            =   4230
         TabIndex        =   23
         ToolTipText     =   "Reporte de ingresos por referida del paciente"
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por empresa referida"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdingresosconmonto 
         Height          =   315
         Left            =   4230
         TabIndex        =   32
         ToolTipText     =   "Reporte de relación de cuentas de pacientes"
         Top             =   5115
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Relación de cuentas de pacientes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdAnticipos 
         Height          =   315
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Reporte de entradas y salidas de dinero"
         Top             =   4305
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Entradas y salidas de dinero"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdConciliacionEmpresas 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Reporte de servicios entre empresas contables"
         Top             =   1875
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Conciliación de servicios entre empresas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRelacionNotas 
         Height          =   315
         Left            =   4230
         TabIndex        =   33
         ToolTipText     =   "Reporte de relación de notas de cargo y crédito"
         Top             =   5520
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Relación de notas de cargo y crédito"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCargosEliminados 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Reporte de cargos eliminados"
         Top             =   255
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cargos eliminados"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdComisionesPromotores 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Reporte de comisiones para promotores"
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
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
      Begin MyCommandButton.MyButton cmdCargosExcedentesEnSA 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Reporte de cargos excedentes de la suma asegurada"
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cargos excedentes de la suma asegurada"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdIngresosPaquetes 
         Height          =   315
         Left            =   4230
         TabIndex        =   24
         ToolTipText     =   "Reporte de ingresos por paquetes"
         Top             =   1875
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por paquetes"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdCronologicoFacturas 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Reporte cronológico de facturas y notas"
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Cronológico de facturas y notas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdIngresosDepto 
         Height          =   315
         Left            =   4230
         TabIndex        =   22
         ToolTipText     =   "Reporte de ingresos por departamento"
         Top             =   1065
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos por departamento"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptFacturacionIntegrada 
         Height          =   315
         Left            =   240
         TabIndex        =   12
         ToolTipText     =   "Reporte de facturación integrada"
         Top             =   5115
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Facturación integrada"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdResumenCargos 
         Height          =   315
         Left            =   4230
         TabIndex        =   34
         ToolTipText     =   "Reporte de resumen diario de cargos y formas de pago"
         Top             =   5925
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Resumen diario de cargos y formas de pago"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdVentasCruzadas 
         Height          =   315
         Left            =   4230
         TabIndex        =   38
         ToolTipText     =   "Reporte de ventas cruzadas"
         Top             =   7545
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ventas cruzadas"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdPacientesAtendidosAseguradoras 
         Height          =   315
         Left            =   4230
         TabIndex        =   29
         ToolTipText     =   "Reporte de pacientes atendidos de aseguradoras"
         Top             =   3900
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Pacientes atendidos de aseguradoras"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdFoliosAntesVenta 
         Height          =   315
         Left            =   240
         TabIndex        =   16
         ToolTipText     =   "Reporte de folios antes de venta"
         Top             =   6735
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Folios antes de venta"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdRptSalidasCajaChica 
         Height          =   315
         Left            =   4230
         TabIndex        =   35
         ToolTipText     =   "Reporte de salidas de caja chica"
         Top             =   6330
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Salidas de caja chica"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdIngresosDiarios 
         Height          =   315
         Left            =   4230
         TabIndex        =   20
         ToolTipText     =   "Reporte de ingresos diarios"
         Top             =   255
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Ingresos diarios"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
      Begin MyCommandButton.MyButton cmdTrazabilidadVentaPublico 
         Height          =   315
         Left            =   4230
         TabIndex        =   40
         ToolTipText     =   "Trazabilidad de medicamentos en ventas al público"
         Top             =   -10000
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Trazabilidad de medicamentos en ventas al público"
         CaptionPosition =   4
         DepthEvent      =   1
         PictureAlignment=   1
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMenuReportesPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lblnSalir As Boolean '|  Indica si la forma se descará en el Activate

Private Sub cmdAnticipos_Click()
On Error GoTo NotificaError

    frmRptAnticiposPaciente.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnticipos_Click"))
    Unload Me
End Sub

Private Sub cmdCargosEliminados_Click()
On Error GoTo NotificaError

    frmCargosEliminados.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargosEliminados_Click"))
End Sub

Private Sub cmdCargosempresas_Click()
On Error GoTo NotificaError
    
    frmFilMedicamentoFarmacia.HelpContextID = 23
    frmFilMedicamentoFarmacia.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargosempresas_Click"))
    Unload Me
End Sub

Private Sub cmdCargosExcedentesEnSA_Click()
On Error GoTo NotificaError

    frmRptCargosExcedentesEnSA.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCargosExcedentesEnSA_Click"))
    Unload Me
End Sub

Private Sub cmdCjaIngresosTurno_Click()
On Error GoTo NotificaError
    
    frmReporteIngresosTurno.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCjaIngresosTurno_Click"))
    Unload Me
End Sub

Private Sub cmdComisionesPromotores_Click()
On Error GoTo NotificaError

    frmFilComisionesPromotores.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdComisionesPromotores_Click"))
    Unload Me
End Sub

Private Sub cmdConciliacionEmpresas_Click()
On Error GoTo NotificaError

    frmRptConciliacionEmpresas.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConciliacionEmpresas_Click"))
    Unload Me
End Sub

Private Sub cmdCronologicoFacturas_Click()
On Error GoTo NotificaError

    frmRptFacturacionCronologica.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCronologicoFacturas_Click"))
End Sub

Private Sub cmdCuentaPendiente_Click()
On Error GoTo NotificaError
    
    frmRptCuentaPendiente.vllngNumOpcion = 1839
    frmRptCuentaPendiente.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCuentaPendiente_Click"))
    Unload Me
End Sub

Private Sub cmdCuentasReabiertas_Click()
On Error GoTo NotificaError

    frmReporteCuentasReabiertas.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCuentasReabiertas_Click"))
    Unload Me
End Sub

Private Sub cmdEstadoCuenta_Click()
On Error GoTo NotificaError
    
    Me.Enabled = False
    frmReporteEstadoCuenta.HelpContextID = 23
    frmReporteEstadoCuenta.Show vbModal
    Me.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEstadoCuenta_Click"))
    Unload Me
End Sub

Private Sub cmdFacturasCanceladas_Click()
On Error GoTo NotificaError

    frmReporteFacturaCancelada.vglngNumeroOpcion = 343
    frmReporteFacturaCancelada.HelpContextID = 23
    frmReporteFacturaCancelada.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdFacturasCanceladas_Click"))
    Unload Me
End Sub

Private Sub cmdFacturasPorIngreso_Click()
On Error GoTo NotificaError
    
    frmFacturasPorIngreso.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdFacturasPorIngreso_Click"))
    Unload Me
End Sub

Private Sub cmdFoliosAntesVenta_Click()
On Error GoTo NotificaError

   frmRptFoliosAntesVenta.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdFoliosAntesVenta_Click"))
    Unload Me
End Sub

Private Sub cmdIngresos_Click()
On Error GoTo NotificaError
    
    frmReporteIngresos.HelpContextID = 23
    frmReporteIngresos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIngresos_Click"))
    Unload Me
End Sub

Private Sub cmdingresosconmonto_Click()
On Error GoTo NotificaError

    frmReportesIngreso.vgstrmodulo = "PV"
    frmReportesIngreso.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdingresosconmonto_Click"))
    Unload Me
End Sub

Private Sub cmdIngresosDepto_Click()
On Error GoTo NotificaError

    frmrptIngresosporDepartamento.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIngresosDepto_Click"))
    Unload Me
End Sub

Private Sub cmdIngresosDiarios_Click()
    frmRptIngresosX.Show vbModal
    
    
End Sub

Private Sub cmdingresosempresa_Click()
On Error GoTo NotificaError

   Frmrptingresosempresapaciente.Show vbModal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdingresosempresa_Click"))
    Unload Me
End Sub

Private Sub cmdIngresosPaquetes_Click()
On Error GoTo NotificaError
    
    frmRptIngresoPaquetes.HelpContextID = 23
    frmRptIngresoPaquetes.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIngresosPaquetes_Click"))
    Unload Me
End Sub

Private Sub cmdIngresosTickets_Click()
On Error GoTo NotificaError
    
    frmRptIngresosPorTickets.HelpContextID = 23
    frmRptIngresosPorTickets.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdIngresosTickets_Click"))
    Unload Me
End Sub

Private Sub cmdPacientesAtendidos_Click()
On Error GoTo NotificaError
    
    frmPacientesAtendidos.HelpContextID = 23
    frmPacientesAtendidos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPacientesAtendidos_Click"))
    Unload Me
End Sub

Private Sub cmdPacientesAtendidosAseguradoras_Click()
On Error GoTo NotificaError
    
'    frmPacientesAtendidos.HelpContextID = 23
    frmPacientesAtendidosAseguradoras.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPacientesAtendidosAseguradoras_Click"))
    Unload Me
End Sub

Private Sub cmdRelacionNotas_Click()
On Error GoTo NotificaError

    frmFilRelacionNotas.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRelacionNotas_Click"))
    Unload Me
End Sub

Private Sub cmdReporteador_Click()
On Error GoTo NotificaError

    frmShowReport.HelpContextID = 23
    frmShowReport.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteador_Click"))
    Unload Me
End Sub

Private Sub cmdReporteDescuentos_Click()
On Error GoTo NotificaError
    
    frmReporteDescuentos.HelpContextID = 23
    frmReporteDescuentos.Show vbModal, Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteDescuentos_Click"))
    Unload Me
End Sub

Private Sub cmdReporteGanancia_Click()
On Error GoTo NotificaError
    
    frmReporteGanancias.HelpContextID = 23
    frmReporteGanancias.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteGanancia_Click"))
    Unload Me
End Sub

Private Sub cmdReporteListaPrecios_Click()
On Error GoTo NotificaError
    
    FrmRptListaPrecio.HelpContextID = 23
    FrmRptListaPrecio.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteListaPrecios_Click"))
    Unload Me
End Sub

Private Sub cmdReportePresupuestos_Click()
On Error GoTo NotificaError
    
    frmReportePresupuestos.HelpContextID = 23
    frmReportePresupuestos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReportePresupuestos_Click"))
    Unload Me
End Sub

Private Sub cmdReporteProductividadMedicos_Click()
On Error GoTo NotificaError
    
    frmReporteProductividadMedicos.HelpContextID = 23
    frmReporteProductividadMedicos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdReporteProductividadMedicos_Click"))
    Unload Me
End Sub

Private Sub cmdResumenCargos_Click()
    frmRptResumenDiarioCargos.Show vbModal, Me
End Sub

Private Sub cmdRptDescuentosAsignados_Click()
On Error GoTo NotificaError
    
    FrmFilRptDescuentos.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptDescuentosAsignados_Click"))
    Unload Me
End Sub

Private Sub cmdRptFacturacionIntegrada_Click()
On Error GoTo NotificaError
    
    frmReporteFacturacionIntegrada.HelpContextID = 23
    frmReporteFacturacionIntegrada.vglngNumeroOpcion = 4012
    frmReporteFacturacionIntegrada.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptFacturacionIntegrada_Click"))
    Unload Me
End Sub

Private Sub cmdRptHonorariosPagados_Click()
On Error GoTo NotificaError
    
    frmRptHonorarioFueraCorte.Show vbModal, Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptHonorariosPagados_Click"))
    Unload Me
End Sub

Private Sub cmdRptIngresoConcepto_Click()
On Error GoTo NotificaError

    frmRptIngresoConcepto.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptIngresoConcepto_Click"))
    Unload Me
End Sub

Private Sub cmdRptRelacionFacturas_Click()
On Error GoTo NotificaError
    
    frmRptRelacionFacturas.HelpContextID = 23
    frmRptRelacionFacturas.vglngNumeroOpcion = 1881
    frmRptRelacionFacturas.Show vbModal

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptRelacionFacturas_Click"))
    Unload Me
End Sub

Private Sub cmdRptSalidasCajaChica_Click()
On Error GoTo NotificaError

    frmRptSalidasCajaChica.Show vbModal
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptTraslado_Click"))
    Unload Me
End Sub

Private Sub cmdRptTraslado_Click()
On Error GoTo NotificaError

    frmRptTraslado.Show vbModal
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRptTraslado_Click"))
    Unload Me
End Sub

Private Sub cmdTrazabilidadVentaPublico_Click()
    frmReportesTrazabilidadVentas.Show vbModal, Me
End Sub

Private Sub cmdVentasCredito_Click()
On Error GoTo NotificaError

    frmRptVentaClientes.Show vbModal, Me
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVentasCredito_Click"))
    Unload Me
End Sub

Private Sub cmdVentasCruzadas_Click()
On Error GoTo NotificaError

    frmRptVentasCruzadas.Show vbModal, Me
  
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVentasCredito_Click"))
    Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError

    UsoTrazavilidad
    
    fblnHabilitaObjetos Me
    If lblnSalir Then Unload Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then Unload Me

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
Dim strTitulo As String

    Me.Icon = frmMenuPrincipal.Icon
    lblnSalir = False
    strTitulo = fRegresaParametro("VCHTITULOCTASPENDFACT", "PvParametro", 0)
    If strTitulo <> "" Then
        cmdCuentaPendiente.Caption = strTitulo
    Else
        'Debe configurar el parámetro Título para el reporte de "Cuentas pendientes de facturar"
        MsgBox SIHOMsg(755), vbCritical, "Mensaje"
        lblnSalir = True
    End If
    
End Sub

Private Sub UsoTrazavilidad()
    Dim rsTrazabilidas As New ADODB.Recordset
    
    Set rsTrazabilidas = frsRegresaRs("SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR " & _
                            "FROM SIPARAMETRO WHERE SIPARAMETRO.VCHNOMBRE = 'BITTRAZABILIDAD' AND SIPARAMETRO.CHRMODULO='IV'")
    If rsTrazabilidas.RecordCount > 0 Then
        If rsTrazabilidas!valor = 1 Then
            CfgFrmTrazabilidad
            cfgTabIndexTrazabilidad
        End If
    End If
End Sub

Private Sub CfgFrmTrazabilidad()
    Me.Height = 9370
    Frame1.Height = 8835
    
    cmdIngresosDiarios.Top = 8355
    cmdIngresosDiarios.Left = 240
    
    cmdRptIngresoConcepto.Top = 255
    cmdIngresosDepto.Top = 660
    cmdingresosempresa.Top = 1065
    cmdIngresosPaquetes.Top = 1470
    cmdIngresosTickets.Top = 1875
    cmdIngresos.Top = 2280
    cmdReporteListaPrecios.Top = 2685
    cmdPacientesAtendidos.Top = 3090
    cmdPacientesAtendidosAseguradoras.Top = 3495
    cmdReportePresupuestos.Top = 3900
    cmdReporteProductividadMedicos.Top = 4305
    cmdingresosconmonto.Top = 4710
    cmdRelacionNotas.Top = 5115
    cmdResumenCargos.Top = 5520
    cmdRptSalidasCajaChica.Top = 5925
    cmdRptTraslado.Top = 6330
    
    cmdVentasCredito.Top = 7140
    cmdVentasCruzadas.Top = 7545

    cmdTrazabilidadVentaPublico.Top = 6735
End Sub

Private Sub cfgTabIndexTrazabilidad()
    cmdRptTraslado.TabIndex = 36
    cmdTrazabilidadVentaPublico.TabIndex = 37
    cmdVentasCredito.TabIndex = 38
    cmdVentasCruzadas.TabIndex = 39
    cmdReporteador.TabIndex = 40
End Sub
