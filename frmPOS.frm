VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas al público"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFoliosVentaImportada 
      Height          =   3855
      Left            =   4320
      TabIndex        =   102
      Top             =   2880
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton cmdAceptarSeleccionTickets 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   1080
         TabIndex        =   106
         ToolTipText     =   "Aceptar"
         Top             =   3360
         Width           =   915
      End
      Begin VB.ListBox lstFoliosTickets 
         Height          =   2535
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   105
         Top             =   720
         Width           =   2880
      End
      Begin VB.CommandButton cmdInvertir 
         Caption         =   "Invertir selección"
         Height          =   375
         Left            =   1350
         TabIndex        =   104
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label24 
         Caption         =   "Folios a incluir"
         Height          =   225
         Left            =   120
         TabIndex        =   103
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.Frame freControlAseguradora 
      Height          =   2385
      Left            =   2625
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   6525
      Begin VB.Frame fraTipoCopago 
         Height          =   390
         Left            =   1035
         TabIndex        =   67
         Top             =   1095
         Width           =   2205
         Begin VB.OptionButton optControlCopagoPorciento 
            Caption         =   "Porcentaje"
            Height          =   195
            Left            =   1050
            TabIndex        =   69
            Top             =   150
            Width           =   1065
         End
         Begin VB.OptionButton optCopagoCantidad 
            Caption         =   "Cantidad"
            Height          =   195
            Left            =   30
            TabIndex        =   68
            Top             =   135
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Frame fraTipoCoaseguro 
         Height          =   360
         Left            =   1035
         TabIndex        =   64
         Top             =   780
         Width           =   2205
         Begin VB.OptionButton optTipoCoaseguro 
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   66
            Top             =   135
            Width           =   1065
         End
         Begin VB.OptionButton optTipoCoaseguro 
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   65
            Top             =   135
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Frame fraTipoDeducible 
         Height          =   375
         Left            =   1035
         TabIndex        =   60
         Top             =   465
         Width           =   2205
         Begin VB.OptionButton optTipoDeducible 
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   1
            Left            =   1050
            TabIndex        =   62
            Top             =   135
            Width           =   1065
         End
         Begin VB.OptionButton optTipoDeducible 
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   61
            Top             =   135
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.TextBox txtCoaseguro 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3270
         MaxLength       =   20
         TabIndex        =   47
         Top             =   870
         Width           =   1185
      End
      Begin VB.CommandButton cmdGrabaControlAseguradora 
         Height          =   495
         Left            =   3090
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPOS.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Grabar control de aseguradora"
         Top             =   1770
         Width           =   495
      End
      Begin VB.TextBox txtDeducible 
         Alignment       =   1  'Right Justify
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3270
         MaxLength       =   20
         TabIndex        =   46
         Top             =   555
         Width           =   1185
      End
      Begin VB.CheckBox chkFacturarCoaseguro 
         Caption         =   "Facturar coaseguro"
         Height          =   210
         Left            =   4725
         TabIndex        =   51
         Top             =   900
         Width           =   1755
      End
      Begin VB.CheckBox chkFacturarCopago 
         Caption         =   "Facturar copago"
         Height          =   210
         Left            =   4725
         TabIndex        =   52
         Top             =   1245
         Width           =   1755
      End
      Begin VB.CheckBox chkFacturarDeducible 
         Caption         =   "Facturar deducible"
         Height          =   210
         Left            =   4725
         TabIndex        =   50
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox txtCopago 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3270
         TabIndex        =   48
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CommandButton cmdCierraControlAseguradora 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6255
         TabIndex        =   44
         Top             =   150
         Width           =   255
      End
      Begin VB.Label lblSignoDeducible 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4515
         TabIndex        =   63
         Top             =   600
         Width           =   120
      End
      Begin VB.Label lblSignoCoaseguro 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4515
         TabIndex        =   57
         Top             =   915
         Width           =   120
      End
      Begin VB.Label Label44 
         Caption         =   "Deducible"
         Height          =   210
         Left            =   135
         TabIndex        =   56
         Top             =   585
         Width           =   870
      End
      Begin VB.Label Label45 
         Caption         =   "Coaseguro"
         Height          =   225
         Left            =   135
         TabIndex        =   55
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label20 
         Caption         =   "Copago"
         Height          =   225
         Left            =   135
         TabIndex        =   54
         Top             =   1230
         Width           =   585
      End
      Begin VB.Label lblSignoPorcientoCopago 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4515
         TabIndex        =   53
         Top             =   1245
         Width           =   120
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Control de aseguradora"
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
         Height          =   225
         Left            =   60
         TabIndex        =   45
         Top             =   165
         Width           =   2235
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   30
         Top             =   105
         Width           =   6480
      End
   End
   Begin VB.Frame freBusqueda 
      Height          =   3420
      Left            =   2745
      TabIndex        =   9
      Top             =   9240
      Width           =   6300
      Begin VB.ListBox lstBuscaArticulos 
         Height          =   2205
         Left            =   75
         TabIndex        =   14
         Top             =   1065
         Width           =   6150
      End
      Begin VB.TextBox txtBuscaArticulo 
         Height          =   285
         Left            =   75
         TabIndex        =   13
         Top             =   735
         Width           =   6150
      End
      Begin VB.CommandButton cmdTachitaBusqueda 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5955
         TabIndex        =   11
         Top             =   165
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   75
         TabIndex        =   12
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda de artículos"
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
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   180
         Width           =   2235
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   30
         Top             =   120
         Width           =   6225
      End
   End
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   967
      TabIndex        =   81
      Top             =   9240
      Visible         =   0   'False
      Width           =   9720
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   82
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
         TabIndex        =   83
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFactura 
      Height          =   1185
      Left            =   180
      TabIndex        =   17
      Top             =   10890
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2090
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmrHora 
      Interval        =   3000
      Left            =   11040
      Top             =   360
   End
   Begin MSComDlg.CommonDialog CDgArchivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "Exportación de pólizas"
      FileName        =   "poliza.txt"
      Filter          =   "Texto (*.txt)|*.txt| Todos los archivos (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog cdgExcel 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freDatosPaciente 
      Height          =   1635
      Left            =   60
      TabIndex        =   19
      Top             =   0
      Width           =   11640
      Begin VB.TextBox txtTicketPrevio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9860
         MaxLength       =   9
         TabIndex        =   88
         ToolTipText     =   "Folio generado antes de la venta"
         Top             =   1155
         Width           =   975
      End
      Begin VB.CommandButton cmdConsultaTicketsNoFacturados 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   10870
         TabIndex        =   84
         ToolTipText     =   "Buscar folio"
         Top             =   1155
         Width           =   660
      End
      Begin VB.TextBox txtGafete 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   77
         Top             =   525
         Width           =   1335
      End
      Begin VB.OptionButton optEmpleado 
         Caption         =   "Empleado"
         Height          =   255
         Left            =   3900
         TabIndex        =   75
         Top             =   200
         Width           =   1000
      End
      Begin VB.OptionButton optMedico 
         Caption         =   "Médico"
         Height          =   255
         Left            =   2600
         TabIndex        =   74
         Top             =   200
         Width           =   1000
      End
      Begin VB.OptionButton optPaciente 
         Caption         =   "Paciente"
         Height          =   255
         Left            =   1305
         TabIndex        =   73
         Top             =   200
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1305
         MaxLength       =   10
         TabIndex        =   1
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label lblHora 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9860
         TabIndex        =   87
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dólar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   9860
         TabIndex        =   86
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Folio antes de venta"
         Height          =   195
         Left            =   8350
         TabIndex        =   85
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   11600
         X2              =   6000
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   6000
         X2              =   6000
         Y1              =   120
         Y2              =   1560
      End
      Begin VB.Label Label22 
         Caption         =   "Gafete"
         Height          =   255
         Left            =   3950
         TabIndex        =   76
         Top             =   575
         Width           =   600
      End
      Begin VB.Label Label21 
         Caption         =   "Médico"
         Height          =   195
         Left            =   150
         TabIndex        =   70
         Top             =   2000
         Width           =   855
      End
      Begin VB.Label lblFactura 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   6720
         TabIndex        =   59
         Top             =   1155
         Width           =   1305
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6120
         TabIndex        =   29
         Top             =   240
         Width           =   3720
      End
      Begin VB.Label lblPaciente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   24
         Top             =   870
         Width           =   4590
      End
      Begin VB.Label lblEmpresa 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1305
         TabIndex        =   20
         Top             =   1215
         Width           =   4605
      End
      Begin VB.Label lblEmpresaTipoPaciente 
         AutoSize        =   -1  'True
         Caption         =   "Tipo paciente"
         Height          =   195
         Left            =   150
         TabIndex        =   28
         Top             =   1275
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Número cuenta"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   585
         Width           =   1110
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Left            =   6120
         TabIndex        =   25
         Top             =   1200
         Width           =   540
      End
   End
   Begin VB.Frame frmMedico 
      Height          =   600
      Left            =   60
      TabIndex        =   71
      Top             =   1530
      Width           =   11640
      Begin VB.CheckBox chkFacturaSustitutaDFP 
         Caption         =   "Factura sustituta"
         Height          =   375
         Left            =   8350
         TabIndex        =   22
         ToolTipText     =   "Indica que la factura que se generará es sustituta de otra previamente cancelada"
         Top             =   130
         Width           =   1480
      End
      Begin VB.ListBox lstFacturaASustituirDFP 
         Height          =   380
         IntegralHeight  =   0   'False
         ItemData        =   "frmPOS.frx":0342
         Left            =   9850
         List            =   "frmPOS.frx":0349
         TabIndex        =   23
         ToolTipText     =   "Facturas a las cuales sustituye"
         Top             =   150
         Width           =   1695
      End
      Begin VB.ComboBox cboMedico 
         Height          =   315
         ItemData        =   "frmPOS.frx":0366
         Left            =   1305
         List            =   "frmPOS.frx":0368
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Seleccione el médico que emite la receta"
         Top             =   185
         Width           =   6120
      End
      Begin VB.Label lblMedico 
         Caption         =   "Médico"
         Height          =   195
         Left            =   150
         TabIndex        =   72
         Top             =   245
         Width           =   975
      End
   End
   Begin VB.Frame FreDetalle 
      Height          =   6945
      Left            =   60
      TabIndex        =   2
      Top             =   2000
      Width           =   11640
      Begin VB.CommandButton cmdDescuentoPuntos 
         Caption         =   "Aplicar puntos de cliente leal"
         Enabled         =   0   'False
         Height          =   350
         Left            =   4635
         TabIndex        =   94
         ToolTipText     =   "Aplicación/Desaplicación de puntos de cliente leal"
         Top             =   4680
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Frame FreTotales 
         Enabled         =   0   'False
         Height          =   3220
         Left            =   7410
         TabIndex        =   107
         Top             =   3435
         Width           =   4130
         Begin VB.TextBox txtImporte 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   119
            ToolTipText     =   "Importe"
            Top             =   315
            Width           =   2175
         End
         Begin VB.TextBox txtIEPS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   112
            ToolTipText     =   "IEPS"
            Top             =   1215
            Width           =   2175
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   111
            ToolTipText     =   "Subtotal "
            Top             =   1665
            Width           =   2175
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   110
            ToolTipText     =   "Total"
            Top             =   2565
            Width           =   2175
         End
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   109
            ToolTipText     =   "IVA"
            Top             =   2115
            Width           =   2175
         End
         Begin VB.TextBox txtDescuentos 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1750
            TabIndex        =   108
            ToolTipText     =   "Descuento"
            Top             =   765
            Width           =   2175
         End
         Begin VB.Label Label13 
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
            Left            =   255
            TabIndex        =   120
            Top             =   405
            Width           =   795
         End
         Begin VB.Label lblIEPS 
            AutoSize        =   -1  'True
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
            Height          =   240
            Left            =   240
            TabIndex        =   117
            Top             =   1305
            Width           =   525
         End
         Begin VB.Label Label17 
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
            Left            =   240
            TabIndex        =   116
            Top             =   1755
            Width           =   870
         End
         Begin VB.Label Label11 
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
            Left            =   255
            TabIndex        =   115
            Top             =   2655
            Width           =   555
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
            Left            =   255
            TabIndex        =   114
            Top             =   2205
            Width           =   375
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Descuento"
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
            Left            =   255
            TabIndex        =   113
            Top             =   855
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdImportarVenta 
         Caption         =   "Importar venta"
         Height          =   495
         Left            =   10125
         TabIndex        =   101
         ToolTipText     =   "Importar archivo Excel con información de ventas"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Frame frePrecios 
         Height          =   4740
         Left            =   2460
         TabIndex        =   30
         Top             =   9150
         Visible         =   0   'False
         Width           =   6105
         Begin VB.Frame Frame6 
            Height          =   180
            Left            =   105
            TabIndex        =   40
            Top             =   4005
            Width           =   5955
         End
         Begin VB.TextBox txtConsultaDescripcion 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   210
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   840
            Width           =   5670
         End
         Begin VB.TextBox txtConsultaDescuento 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   3285
            Width           =   3870
         End
         Begin VB.TextBox txtConsultaPrecio 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   1965
            Width           =   3870
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta de precios"
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
            Height          =   225
            Left            =   75
            TabIndex        =   39
            Top             =   180
            Width           =   2235
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   195
            TabIndex        =   38
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "Presione <ESC> para salir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   1380
            TabIndex        =   35
            Top             =   4275
            Width           =   3450
         End
         Begin VB.Label Label2 
            Caption         =   "Descuento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   255
            TabIndex        =   33
            Top             =   2940
            Width           =   1380
         End
         Begin VB.Label Label1 
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   31
            Top             =   1605
            Width           =   885
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H80000002&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   330
            Left            =   45
            Top             =   135
            Width           =   6060
         End
      End
      Begin VB.Frame freGraba 
         Height          =   780
         Left            =   5685
         TabIndex        =   97
         Top             =   5875
         Width           =   1680
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   590
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPOS.frx":036A
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "<END>  Grabar venta"
            Top             =   180
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   495
            Left            =   1090
            Picture         =   "frmPOS.frx":06AC
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Consultas"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdImprimeTicketSinFacturar 
            Height          =   495
            Left            =   80
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPOS.frx":0B9E
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "Imprimir consumo antes de guardar la venta"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame freDescuentos 
         Caption         =   "Descuentos"
         Height          =   1100
         Left            =   4635
         TabIndex        =   90
         Top             =   3435
         Width           =   2730
         Begin VB.CommandButton cmdControlAseguradora 
            Caption         =   "Control de aseguradoras"
            Enabled         =   0   'False
            Height          =   350
            Left            =   240
            TabIndex        =   93
            ToolTipText     =   "Permite capturar el control de la aseguradora."
            Top             =   630
            Width           =   2230
         End
         Begin VB.CommandButton cmdDescuenta 
            Caption         =   "Aplicar"
            Height          =   315
            Left            =   1500
            TabIndex        =   92
            ToolTipText     =   "Aplicar descuento capturado"
            Top             =   250
            Width           =   975
         End
         Begin VB.TextBox txtDescuento 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            MaxLength       =   6
            TabIndex        =   91
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   1230
            TabIndex        =   96
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label18 
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
            ForeColor       =   &H00E0E0E0&
            Height          =   255
            Left            =   1200
            TabIndex        =   95
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00;(#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   58
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSCommLib.MSComm comPrinter 
         Left            =   8280
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.Frame Frame8 
         Caption         =   "Mensajes"
         Height          =   1600
         Left            =   120
         TabIndex        =   41
         Top             =   5055
         Width           =   5500
         Begin VB.Label lblBusquedaPacientes 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<F4> - Buscar pacientes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   80
            Top             =   1150
            Width           =   5250
         End
         Begin VB.Label lblBusquedaMedicos 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<F6> - Buscar médicos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   79
            Top             =   850
            Width           =   5250
         End
         Begin VB.Label lblBusquedaEmpleados 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "<F7> - Buscar empleados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   78
            Top             =   550
            Width           =   5250
         End
         Begin VB.Label lblMensajes 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   120
            TabIndex        =   42
            Top             =   250
            Width           =   5250
         End
      End
      Begin VB.Frame freConsultaPrecios 
         Height          =   1100
         Left            =   1800
         TabIndex        =   36
         Top             =   3435
         Width           =   2450
         Begin VB.CheckBox chkConsultaPrecios 
            Caption         =   "<F5> Consulta de precios"
            Height          =   435
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Permite consultar precios"
            Top             =   350
            Width           =   2200
         End
      End
      Begin VB.Frame freFacturaPesosDolares 
         Caption         =   "Facturar en "
         Height          =   1100
         Left            =   120
         TabIndex        =   18
         Top             =   3435
         Width           =   1290
         Begin VB.OptionButton optPesos 
            Caption         =   "Dólares"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   700
            Width           =   900
         End
         Begin VB.OptionButton optPesos 
            Caption         =   "Pesos"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   4
            Top             =   350
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.OptionButton optOtrosConceptos 
         Caption         =   "Otros conceptos"
         Height          =   195
         Left            =   6705
         TabIndex        =   16
         Top             =   495
         Width           =   1500
      End
      Begin VB.OptionButton optArticulo 
         Caption         =   "Artículos"
         Height          =   195
         Left            =   5625
         TabIndex        =   15
         Top             =   495
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4500
         MaxLength       =   4
         TabIndex        =   3
         Top             =   420
         Width           =   930
      End
      Begin VB.TextBox txtClaveArticulo 
         Height          =   315
         Left            =   150
         TabIndex        =   0
         Top             =   420
         Width           =   3450
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdArticulos 
         Height          =   2475
         Left            =   150
         TabIndex        =   89
         Top             =   885
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   4366
         _Version        =   393216
         GridColor       =   12632256
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblPuntosLealtad 
         Alignment       =   1  'Right Justify
         Caption         =   "Puntos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   215
         Left            =   4520
         TabIndex        =   118
         Top             =   6670
         Visible         =   0   'False
         Width           =   7000
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad"
         Height          =   225
         Left            =   3720
         TabIndex        =   8
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código de barras"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   180
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmPOS                                                 -
'-------------------------------------------------------------------------------------
'| Objetivo: Es el punto de venta al público de farmacia
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 25/Oct/2001
'| Modificó                 : Nombre(s)
'------------------------------------------------------------------------------------

Option Explicit
Dim vlstrClaveDepartamento As String  'variable en la cual guardamos la clave del departamento con el que se ingreso al modulo y que usaremos para comparar al cargar el grid de productos vendidos sin facturar
Dim vlbflag As Boolean
Dim vgParametroConvenio As Integer
Dim vgintEmpresa As Integer
Dim vgstrNombreFactura As String
Dim vgstrDireccionFactura As String
Dim vgstrNumeroExteriorFactura As String
Dim vgstrNumeroInteriorFactura As String
Dim vgBitExtranjeroFactura As Integer
Dim vgstrColoniaFactura As String
Dim vgstrCPFactura As String
Dim vgstrTelefonoFactura As String
Dim vgstrRFCFactura As String
Dim vgstrNombreFacturaPaciente As String
Dim vgstrDireccionFacturaPaciente As String
Dim vgstrNumeroExteriorFacturaPaciente As String
Dim vgstrNumeroInteriorFacturaPaciente As String
Dim vgBitExtranjeroFacturaPaciente As Integer
Dim lngCveFormato As Long                           'Para saber el formato que se va a utilizar (relacionado con pvDocumentoDepartamento.intNumFormato)
Dim vllngFormatoaUsar As Long                       'Para saber que formato se va a utilizar
Dim intTipoEmisionComprobante As Integer            'Variable que compara el tipo de formato y folio a utilizar (0 = Error de formato y folios incompatibles, 1 = Físicos, 2 = Digitales)
Dim vgstrColoniaFacturaPaciente As String
Dim vgstrCPFacturaPaciente As String
Dim llngCveCiudadPaciente As Long
Dim llngCveCiudad As Long
Dim vgstrTelefonoFacturaPaciente As String
Dim vgstrRFCFacturaPaciente As String

Dim strCurrPrinter As String

Dim vgintTipoPaciente As Integer
Dim vgintTipoParticular As Integer
Dim vglngCveExtra As Long                        'Clave del empleado, médico, empresa etc. que se guardará en PvDatosFiscales.intNumReferencia
Dim vgstrTipoPaciente As String                     'Tipo del cliente: Empleado, médico, empresa etc. que se guardará en PvDatosFiscales.chrTipoCliente
Dim vgstrEstadoManto As String
Dim aFormasPago() As FormasPago                     'Estas formas de pago son para cuando paga la EMPRESA o en su defecto, el paciente que no tiene capturado el control de aseguradora.
Dim aFormasPagoPaciente() As FormasPago             'Estas son para cuando paga el PACIENTE, cuando tiene configurado el control de aseguradoras y que se le da FACTURA.
Dim aFormasPagoReciboPaciente() As FormasPago       'Estas son para cuando paga el PACIENTE, cuando tiene configurado el control de aseguradoras y que se le da RECIBO.
Dim rsConsultaFacturas As New ADODB.Recordset
Dim vlblnPrimeraVez As Boolean
Dim vldblTipoCambio As Double                       'Trae el tipo de cambio utilizado en la pantalla de Formas de pago
Dim lblnImpresoraSerial As Boolean                  'Parametro para que imprima un ticket en impresora serial o normal
Dim lstrLeyendaCliente As String                    'Para la impresión del ticket
Dim vllngMedicoDefaultPOS As Long
Dim lngTipoPacMedico As Long
Dim lngTipoPacEmpleado As Long
Dim vlintValorPMP As Integer                     'Bandera para validar el Precio máximo al publico de los medicamentos de venta al público

Private vgrptReporte As CRAXDRT.Report              ' Tipo de dato para Arreglo para grabar información de la factura del paciente
Private Type typFacturaPaciente
    dblCantidad As Double
    dblCantidadIVA As Double
    lngConceptoFacturacion As Long
    strTipo As String
End Type

Private Type typTicketsSeleccionados
    lngTicket As Long
    dtmfecha As Date
End Type
Dim aTickets() As typTicketsSeleccionados

Dim rsGnParametrosTipoCorte As New ADODB.Recordset
Dim vlblnParametroTipoCorte As Boolean
Dim vllngPersonaGraba As Long
Dim lCveLista As Long
Dim vlbsinalmacen As Boolean
Dim vlstrNumRef  As String
Dim vlstrTipo As String
Const intAplicado = 0                           'Indica que tomará el concepto de facturación base para el cálculo de los descuentos

Dim llngCveConceptoDeducible As Long            'Clave para factura deducible
Dim llngCveConceptoCoaseguro As Long            'Clave para factura coaseguro
Dim llngCveConceptoCopago As Long               'Clave para factura copago

Dim ldblDeducible As Double                     'Cantidad de deducible
Dim ldblCoaseguro As Double                     'Cantidad de coaseguro
Dim ldblCopago As Double                        'Cantidad de copago
Dim lblnGuardoControl As Boolean                'Para identificar si se agregaron mas cargos después de guardar el control del seguro
Dim lblnExisteControl As Boolean                'Para identificar cuando el paciente tiene control de aseguradora

                                                ' Variables agregadas para la funcionalidad de CFD's
Dim vllngFoliosRestantes As Long                'Solo para mostrar el mensaje de advertencia
Dim vlstrFolioDocumento As String
Dim vlaryParametrosSalida() As String

Dim strFolio As String
Dim strSerie As String
Dim strFolioPaciente As String
Dim strSeriePaciente As String
Dim strNumeroAprobacionPaciente As String
Dim strAnoAprobacionPaciente As String
Dim strFolioDocumento As String
Dim strSerieDocumento As String
Dim strNumeroAprobacionDocumento As String
Dim strAnoAprobacionDocumento As String

Dim lngCargos As Long
Dim intFacturarVentaPublico As Integer          'Indica si se facturan automáticamente los tikets de venta al público
Dim llngFormatoFacturaAUsar As Long

Dim blnFacturaDeducible As Boolean              'Si se factura o no el Deducible
Dim blnFacturaCoaseguro As Boolean              'Si se factura o no el Coaseguro
Dim blnFacturaCopago As Boolean                 'Si se factura o no el CoPago

Dim blnFactuaAutomatica As Boolean
Dim vlstrSentencia As String

Dim vgstrRegimenFiscal As String
Dim strNombrePOS  As String
Dim strCallePOS  As String
Dim strNumeroExteriorPOS  As String
Dim strNumeroInteriorPOS  As String
Dim strColoniaPOS  As String
Dim strCPPOS  As String
Dim lngCveCiudadPOS As Long
Dim vgintSocioRelacionado As Long   'variable en la que se almacena la clave del socio relacionado (en caso de haber ingresado a un paciente de tipo socio)
Dim vlblnLicenciaIEPS As Boolean
Dim vlintBitSaldarCuentas As Long               'Variable que indica el valor del bit pvConceptoFacturacion.BitSaldarCuentas, que nos dice si la cuenta del ingreso se salda con la del descuento
Dim vlblnCuentaIngresoSaldada As Boolean        'Variable que indica si la cuenta del ingreso fue saldada con la cuenta del descuento
Dim intCuentaNueva As Integer ' 1 = Es cuenta nueva 0 = No es cuenta nueva

Dim vldblFacturaPacienteSubTotal As Double   'Aqui traigo la suma de los otros conceptos para la factura del paciente
Dim vldblFacturaPacienteIVA As Double        'Aqui traigo la cantidad de IVA de la factura del paciente
Dim vldblFacturaPacienteTotal As Double      'Aqui traigo la cantidad TOTAL de la factura del paciente
Dim blnFacturaAutomatica As Boolean     'Para saber si se va a genera una facturaautomática del ticket
Dim vlbytNumeroConceptosPaciente As Byte     'Esta me sirve para saber cuantos conceptos tiene la factura del paciente, para poder prorratear el IVA
Dim aFPFacturaPaciente() As typFacturaPaciente 'Aqui traigo el detalle de la factura del paciente
Dim blnFacturarConceptosSeguro As Boolean
Dim vldblTotalControlAseguradorasSinIVA As Double
Dim vldblTotalControlAseguradoras As Double
Dim vldblTotIEPS As Double                   'IEPS total de la cuenta
Dim vlstrFolioDocumentoPaciente As String

Dim vgstrRFCPersonaSeleccionada As String
Dim vldblIngresosPuente As Double           'Cantidad del asiento contable que corresponde a los ingresos
Dim vl_dblClickCmdDescuenta As Boolean      'Variable usada para controlar el el keyup del gridArticulos
Dim vlblnConsultaTicketPrevio As Boolean

Dim vlstrNombreFacturaPaciente As String     'El nombre de la factura del paciente, en el caso de aseguradora
Dim vlstrRFCFacturaPaciente As String        'RFC para la factura  del paciente, en el caso de aseguradora
Dim vlstrDireccionFacturaPaciente As String  'Dirección para la Factura  del paciente, en el caso de aseguradora
Dim vlstrNumeroExteriorFacturaPaciente As String  'Número Exterior para la Factura  del paciente, en el caso de aseguradora
Dim vlstrNumeroInteriorFacturaPaciente As String  'Número Interior para la Factura  del paciente, en el caso de aseguradora
Dim vlBitExtranjeroFacturaPaciente
Dim vlstrColoniaFacturaPaciente As String    'Colonia para la Factura  del paciente, en el caso de aseguradora
Dim vlstrCPFacturaPaciente As String         'Código postal para la Factura  del paciente, en el caso de aseguradora
Dim vlstrTelefonoFacturaPaciente As String   'Telefono para la factura del paciente, en el caso de aseguradora
Dim vllngCveCiudadFacturaPaciente As Long

Dim vlintCopiasTicket As Integer


'    ' Datos Fiscales para la facturada
Dim vlstrNombreFactura As String             'El nombre de aquien se le va a facturar
Dim vlstrRFC As String                       'RFC para la factura
Dim vlstrDireccion As String                 'Dirección para la Factura
Dim vlstrNumeroExterior As String            'Número Exterior para la Factura
Dim vlstrNumeroInterior As String            'Número Interior para la Factura
Dim vlBitExtranjero As Integer
Dim vlstrColonia As String                   'Colonia para la factura
Dim vlstrCP As String                        'Código postal para la factura
Dim vlstrTelefono As String                  'Telefono para la factura
Dim vlintUsoCFDI As Long

Dim fs As FileSystemObject
Dim f As Variant
Dim xlsApp As Object 'Excel.Application
Dim hoja As Object 'Excel.Worksheet
Dim rngfnd As Object 'Excel.Range

Dim vlblnImportoVenta As Boolean

Dim vldblTotalIVAGuardado As Double
Dim vldblSubtotalGuardado As Double

Dim vllngConsecutivoFactura As Long          'Consecutivo de la factura BASE para la impresión
Dim lstrFolioDocumentoF As String
Dim vlrsSelTicket As New ADODB.Recordset     'Para eliminar el uso del comando cmdSelTicketr

Dim alstrParamTickets(9) As String

'    ' Variables para agregar un registro en PVFACTURAIMPORTE cuando se facturan conceptos de seguro al paciente
Dim vldblsubtotalgravado As Double
Dim vldblsubtotalNogravado As Double
Dim vldbldescuentogravado As Double
Dim vldblDescuentoNoGravado As Double

Dim alstrParametros(6) As String
Dim vlngCveFormato As Long

Dim vlnblnLocate As Boolean
Dim vllngNumeroPaciente As Long
Dim vlngEmpresa As Long
Dim llngNumCliente As Long
Dim vlstrsql As String

Dim blnLicenciaLealtadCliente As Boolean    'Para saber si se tiene licencia para generar la lealtad del cliente y el médico
Dim vlstrMensajePuntos As String
Dim vldblMontoDisponiblePuntos As Double
Dim vllngPersonaGrabaPuntos As Long
Dim dblDescuentoAplicadoPuntos As Double
Dim vldblPuntosDisponibles As Double
Dim vldblPuntosAplicados As Double

Dim vlchrIncluirConceptosSeguro As String

Public lngCantidad As Long
Public strLote As String
Public strCveArticulo As String
Public strVariosLotes As String

Dim vlLngIvLoteProcedencia As Long

Dim vlblnAplicaTrazabilidad As Boolean  'Indica si aplica Trazabilidad - CASO 19582
Dim vlstrCveArticuloTraza As String 'Guardamos la clave del articulo para mostrarla en la captura del lote Trazabilidad
Dim vlblnBitTrazabilidad As Boolean

'lote de salida LMM
Dim vllngdeptolote As Long
Public vlblnSeleccionoLote As Boolean

Dim vlBlnManejaLotes As Boolean 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]

Private Sub pDesAplicarDescuentoPuntos()
On Error GoTo NotificaError
    Dim llngRow As Long
    Dim llngPersonaGraba As Long
    Dim lstrSentencia As String
    Dim rsTabulador As New ADODB.Recordset
    Dim dblDescuentoPuntos As Double
    Dim vldblSubtotal As Double
    Dim vllngContador As Long
    Dim vldblDescuentoPuntos As Double
    Dim vldblDescuentoCargo As Double
    Dim vldblImporteFactura As Double
    
    ' -- 17115 --
    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If llngPersonaGraba = 0 Then Exit Sub
    
    vllngPersonaGrabaPuntos = llngPersonaGraba
    
    '---------------------------------------------
    ' Eliminar descuento por aplicación de puntos
    '---------------------------------------------
    For vllngContador = 1 To grdArticulos.Rows - 1
        With grdArticulos
            .TextMatrix(vllngContador, 5) = .TextMatrix(vllngContador, 27)  ' Descuento antes de aplicar descuento por puntos
            .TextMatrix(vllngContador, 27) = "0"
            .TextMatrix(vllngContador, 28) = "0"   'Descuento aplicado por puntos
        End With
    Next
    
    ' Aplica porcentaje de descuentos originales
    For vllngContador = 1 To grdArticulos.Rows - 1
        vldblDescuentoCargo = CDbl(grdArticulos.TextMatrix(vllngContador, 5)) / (CDbl(grdArticulos.TextMatrix(vllngContador, 24)) * CLng(grdArticulos.TextMatrix(vllngContador, 3))) * 100
        pActualizaDescuentos (vldblDescuentoCargo), vllngContador
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGrabaPuntos, "DESCUENTOS EN PANTALLA DE VENTA AL PÜBLICO A CLIENTE LEAL (Desaplicar descuento puntos)", "Cta. " & txtMovimientoPaciente.Text & " Cargo " & grdArticulos.RowData(vllngContador) & " Descto. " & Format(Val(vldblDescuentoPuntos), "#,###.##00")) ''F10
    Next
        
    grdArticulos.Redraw = True
    grdArticulos.Refresh
    If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
    grdArticulos.Col = 1
               
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDesAplicarDescuentoPuntos"))
End Sub


Private Sub pActualizaDescuentoPuntos(vlintContador As Integer)

    ' Si se aplicó descuento por puntos de cliente leal,
    ' guarda el descuento del cargo antes de aplicarlo el descuento por puntos
    ' y guarda el descuento aplicado por puntos
    
    If Val(Format(grdArticulos.TextMatrix(vlintContador, 28), "")) > 0 Then
        pEjecutaSentencia "Update PvCargo Set mnyDescuentoOriginal = " & Val(Format(grdArticulos.TextMatrix(vlintContador, 27), "")) & _
                                           ", mnydescuentoPuntos = " & Val(Format(grdArticulos.TextMatrix(vlintContador, 28), "")) & _
                          " where intnumcargo = " & grdArticulos.TextMatrix(vlintContador, 15)
    End If
    
End Sub


Private Sub pActualizaDescuentos(vldblDescuento As Double, vllngRenglon As Long, Optional vlSocio As Integer)
    Dim vlintContador As Integer
'    Dim vldblDescuento As Double
    Dim vldblDesctoAplicado As Double
    Dim vldblPrecio As Double
    Dim vlintCantidad As Double
    Dim vldblSubtotal As Double
    Dim vldblIVA As Double
    Dim vlaryParametrosSalida() As String
    Dim vldblTotDescuento As Double
    Dim vldblPorcentajeD As Double
    Dim vlstrMensaje As String
    Dim vllngSeleccionados As Long
'    Dim vllngDescontados As Long
    
    vldblDescuento = vldblDescuento / 100
    vlintContador = vllngRenglon
       
    If grdArticulos.RowData(1) <> -1 Then
        grdArticulos.Redraw = False
        vlstrMensaje = ""
'        vllngSeleccionados = 0
        'vllngDescontados = 0
               
'        For vlintContador = 1 To grdArticulos.Rows - 1
            'If grdArticulos.TextMatrix(vlintContador, 0) = "*" Then
               'vllngSeleccionados = vllngSeleccionados + 1
               With grdArticulos
                    '------------------------------------
                    'Se revisa la exclusión del descuento
                    '------------------------------------
'                    If .TextMatrix(vlintContador, 20) = "2" Then
'                       If vlstrMensaje = "" Then
'                          vlstrMensaje = .TextMatrix(vlintContador, 1)
'                       Else
'                          vlstrMensaje = vlstrMensaje & vbNewLine & .TextMatrix(vlintContador, 1)
'                       End If
'
'
'                    Else
                        'vllngDescontados = vllngDescontados + 1
                        vlintCantidad = Val(.TextMatrix(vlintContador, 3))
                        
                        If CDbl(.TextMatrix(vlintContador, 24)) <> 0 Then
                            vldblDesctoAplicado = (Val(Format(.TextMatrix(vlintContador, 24), "")) * vldblDescuento * vlintCantidad)
                        Else
                            vldblDesctoAplicado = (Val(Format(.TextMatrix(vlintContador, 2), "")) * vldblDescuento * vlintCantidad)
                        End If
                        
                        
                        .Row = vlintContador
                        '-----------------------
                        'Descuentos
                        '-----------------------
                        '-----------------------
                        ' Procedimiento para obtener el IVA
                        '-----------------------
                        vldblIVA = fdblObtenerIva(.RowData(.Row), .TextMatrix(.Row, 12)) / 100
                        If Val(.TextMatrix(vlintContador, 18)) > 0 Then
                            vldblPorcentajeD = Val(.TextMatrix(vlintContador, 18)) + Val(txtDescuento.Text)
                        Else
                            vldblPorcentajeD = IIf(Val(Format(.TextMatrix(vlintContador, 5), "")) > 0, Format((vldblDesctoAplicado * 100) / .TextMatrix(vlintContador, 4), "###.00"), Val(txtDescuento.Text))
                        End If
                        .TextMatrix(.Row, 27) = IIf(.TextMatrix(.Row, 5) = "", "0", .TextMatrix(.Row, 5)) ' Descuento antes de aplicar descuento por puntos
                        .TextMatrix(.Row, 5) = Format(vldblDesctoAplicado, "$###,###,###,###.00")
                        .TextMatrix(.Row, 28) = vldblDesctoAplicado - CDbl(Format(.TextMatrix(.Row, 27), ""))   'Descuento aplicado por puntos
                        
                        If CDbl(.TextMatrix(vlintContador, 24)) <> 0 Then
                            vldblSubtotal = Round((Val(Format(.TextMatrix(vlintContador, 24), "")) * vlintCantidad), 2) - Val(Format(.TextMatrix(.Row, 5), "############.00")) 'vldblDescuento
                        Else
                            vldblSubtotal = (Val(Format(.TextMatrix(vlintContador, 2), "")) * vlintCantidad) - Val(Format(.TextMatrix(.Row, 5), "############.00")) 'vldblDescuento
                        End If

                        .TextMatrix(.Row, 6) = Format(vldblSubtotal * Val(Format(.TextMatrix(.Row, 19), "")), "$###,###,###,###.00") 'IEPS@
                        vldblSubtotal = vldblSubtotal + (vldblSubtotal * Val(Format(.TextMatrix(.Row, 19), ""))) 'IEPS + SUBTOTAL
                        .TextMatrix(.Row, 7) = Format(vldblSubtotal, "$###,###,###,###.00") '@
                        .TextMatrix(.Row, 8) = Format(vldblSubtotal * vldblIVA, "$###,###,###,###.00") '
                        .TextMatrix(.Row, 9) = Format(vldblSubtotal * (vldblIVA + 1), "$###,###,###,###.00") '
                        .Col = 9
                        .CellFontBold = True
                        .TextMatrix(.Row, 11) = vldblSubtotal * vldblIVA
                        .TextMatrix(vlintContador, 0) = ""
                        .TextMatrix(.Row, 18) = vldblPorcentajeD '
'                    End If
                End With
                pCalculaTotales
            'End If
'        Next
        
'        If vllngSeleccionados > 0 Then 'hay seleccionados
'           If vllngDescontados = 0 Then ' no hay descuentos por que hay exclusiones
'              'Mensaje: No se puede aplicar el descuento debido a que los articulos cuentan con exclusion de descuento.
'               MsgBox SIHOMsg(1273), vbOKOnly + vbExclamation, "Mensaje"
'           Else
'               If vllngDescontados < vllngSeleccionados Then
'                  'Mensaje: No se puede aplicar el descuento en los siguientes artículos debido a que se encontró con una exclusión de descuento:
'                   MsgBox SIHOMsg(1274) & vbNewLine & vlstrMensaje, vbOKOnly + vbExclamation, "Mensaje"
'               End If
'           End If
'        End If
               
        grdArticulos.Redraw = True
        grdArticulos.Refresh
        vl_dblClickCmdDescuenta = True
        If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
        grdArticulos.Col = 1
    End If

End Sub


Private Sub pAplicarDescuentoPuntos()
On Error GoTo NotificaError
    Dim llngRow As Long
    Dim llngPersonaGraba As Long
    Dim lstrSentencia As String
    Dim rsTabulador As New ADODB.Recordset
    Dim dblDescuentoPuntos As Double
    Dim vldblSubtotal As Double
    Dim vllngContador As Long
    Dim vldblDescuentoPuntos As Double
    Dim vldblDescuentoCargo As Double
    Dim vldblImporteFactura As Double
    Dim vldblFacturaExcedente As Double
    Dim vldblFacturaDeducible As Double
    Dim vldblFacturaCoaseguro As Double
    Dim vldblFacturaCoaseguroMedico As Double
    Dim vldblFacturaCoaseguroAdicional As Double
    Dim vldblFacturaCopago As Double
    Dim vldblImporteConceptosSeguro As Double
    Dim vldblSubtotalConcepto As Double
    Dim vldblSaldoDisponiblePuntos As Double
    Dim vldblMontoAplicablePuntos As Double
    Dim vldblImporteCargosExcluidos As Double
    Dim vldblSubtotalCargo As Double
    Dim vldblImporteCargo As Double
    Dim dblDescuentoAplicadoPuntosLocal As Double

    ' -- 17115 --
    llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If llngPersonaGraba = 0 Then Exit Sub

    vllngPersonaGrabaPuntos = llngPersonaGraba

    '------------------------------------
    ' Descuento por aplicación de puntos
    '------------------------------------

    ' Aplica porcentaje de descuentos por puntos
    dblDescuentoAplicadoPuntos = 0
    If Val(Format(txtSubtotal.Text, "")) >= vldblMontoDisponiblePuntos Then
        dblDescuentoAplicadoPuntos = vldblMontoDisponiblePuntos
    ElseIf Val(Format(txtSubtotal.Text, "")) < vldblMontoDisponiblePuntos Then
        dblDescuentoAplicadoPuntos = Val(Format(txtSubtotal.Text, ""))
    End If
    vldblImporteFactura = CDbl(txtSubtotal.Text) + CDbl(txtDescuentos.Text)
    vldblDescuentoPuntos = dblDescuentoAplicadoPuntos / vldblImporteFactura * 100
    dblDescuentoAplicadoPuntosLocal = dblDescuentoAplicadoPuntos

    'Aplica descuento en cargos
    vldblPuntosAplicados = 0
    For vllngContador = 1 To grdArticulos.Rows - 1
        If vldblPuntosAplicados < dblDescuentoAplicadoPuntos Then
            'If grdArticulos.TextMatrix(vllngContador, 0) = "*" Then
                vldblDescuentoCargo = CDbl(grdArticulos.TextMatrix(vllngContador, 5)) / (CDbl(grdArticulos.TextMatrix(vllngContador, 24)) * CLng(grdArticulos.TextMatrix(vllngContador, 3))) * 100
                vldblImporteCargo = CDbl(grdArticulos.TextMatrix(vllngContador, 24)) * CLng(grdArticulos.TextMatrix(vllngContador, 3))
                vldblSubtotalCargo = (CDbl(grdArticulos.TextMatrix(vllngContador, 24)) * CLng(grdArticulos.TextMatrix(vllngContador, 3))) - CDbl(grdArticulos.TextMatrix(vllngContador, 5))
                If vldblSubtotalCargo >= dblDescuentoAplicadoPuntosLocal Then
                    vldblMontoAplicablePuntos = dblDescuentoAplicadoPuntosLocal
                ElseIf vldblSubtotalCargo < dblDescuentoAplicadoPuntosLocal Then
                    vldblMontoAplicablePuntos = vldblSubtotalCargo
                End If
                vldblDescuentoPuntos = vldblMontoAplicablePuntos / vldblImporteCargo * 100
                pActualizaDescuentos (vldblDescuentoCargo + vldblDescuentoPuntos), vllngContador
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGrabaPuntos, "DESCUENTOS EN PANTALLA DE VENTA AL PÚBLICO A CLIENTE LEAL", "Cta. " & txtMovimientoPaciente.Text & " Cargo " & grdArticulos.RowData(vllngContador) & " Descto. " & Format(Val(vldblDescuentoPuntos), "#,###.##00")) ''F10
                vldblPuntosAplicados = vldblPuntosAplicados + vldblMontoAplicablePuntos
                dblDescuentoAplicadoPuntosLocal = dblDescuentoAplicadoPuntosLocal - vldblMontoAplicablePuntos
            'End If
        End If
    Next
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAplicarDescuentoPuntos"))
End Sub


Private Sub pAcumularPuntos(lngConsecutivoFactura As Long, strFolioDocumento As String, vldtmFechaHoy As Date, vldtmHoraHoy As Date)
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsRango As New ADODB.Recordset
    Dim rsPuntosObtenidos As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim vllngCveMedicoLab As Long
    Dim vllngCveMedicoImagen As Long
    Dim vllngCveMedicoCargo As Long
    Dim vllngCveTipoIngreso As Long
    
    '-- 17115 --
    vllngCveTipoIngreso = 0
    vllngCveMedicoLab = 0
    vllngCveMedicoImagen = 0
    vllngCveMedicoCargo = 0
    
    'Busca si el Subtotal entra en algún rango en el tabulador de puntos
    vlstrSentencia = " Select min(ta.mnylimitesuperior), ta.intPuntosPaciente, ta.intPuntosMedico From PvTabuladorPuntosLealtad ta Where " & Val(Format(txtSubtotal.Text, "")) & " >= ta.mnylimiteinferior" & _
                     " And " & Val(Format(txtSubtotal.Text, "")) & " <= ta.mnylimitesuperior " & _
                     " Group by ta.intpuntospaciente, ta.intpuntosmedico"
    
    Set rsRango = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
    If rsRango.RecordCount > 0 Then
        vlstrSentencia = "select intCveTipoIngreso, intCveTipoPaciente from ExPacienteIngreso where intNumCuenta = " & Val(txtMovimientoPaciente.Text)
        Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
        If rs.RecordCount > 0 Then vllngCveTipoIngreso = rs!intCveTipoIngreso
        
        '- 8 EXTERNO
        '-10 CONSULTA EXTERNA
        If vllngCveTipoIngreso = 8 Or vllngCveTipoIngreso = 10 Then
            'Busca si existe una solicitud de laboratorio, para generar puntos al médico
            vlstrSentencia = "select nvl(min(intnummedico), 0) intnummedico from LASolicitudExamen where intMovPaciente = " & Val(txtMovimientoPaciente.Text)
            Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
            If rs.RecordCount > 0 Then vllngCveMedicoLab = rs!intnummedico
        
            'Busca si existe una solicitud de imagenologia, para generar puntos al médico
            vlstrSentencia = "select nvl(min(intnummedico), 0) intnummedico from IMSolicitudEstudio where intMovPaciente = " & Val(txtMovimientoPaciente.Text)
            Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
            If rs.RecordCount > 0 Then vllngCveMedicoImagen = rs!intnummedico
        End If
                
        '- 1   INTERNAMIENTO Normal
        '- 2   AMBULATORIO
        '- 4   INTERNO FUE URGENCIAS
        '- 5   INTERNO FUE AMBULATORIO
        '- 6   RECIÉN NACIDO
        '-11  CORTA ESTANCIA
        '-12  INTERNO FUE CORTA ESTANCIA
        '-13  CORTA ESTANCIA FUE AMBULATORIO
        If vllngCveTipoIngreso = 1 Or vllngCveTipoIngreso = 2 Or vllngCveTipoIngreso = 4 Or vllngCveTipoIngreso = 5 Or vllngCveTipoIngreso = 6 Or vllngCveTipoIngreso = 11 Or vllngCveTipoIngreso = 12 Or vllngCveTipoIngreso = 13 Then
            'Busca si hay médico a cargo, para generar puntos al médico
            vlstrSentencia = "select nvl(min(intcvemedicotratante), 0) intnummedico from ExPacienteIngreso where intnumcuenta = " & Val(txtMovimientoPaciente.Text)
            Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
            If rs.RecordCount > 0 Then vllngCveMedicoCargo = rs!intnummedico
        Else
            If cboMedico.ListIndex <> -1 Then
                vllngCveMedicoCargo = cboMedico.ItemData(cboMedico.ListIndex)
            End If
        End If
                
        vlstrSentencia = "SELECT * FROM PvPuntosObtenidosPaciente"
        Set rsPuntosObtenidos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
           
        rsPuntosObtenidos.AddNew
        rsPuntosObtenidos!intIdFactura = lngConsecutivoFactura
        rsPuntosObtenidos!chrfoliofactura = Trim(strFolioDocumento)
        rsPuntosObtenidos!intNumCuenta = CLng(txtMovimientoPaciente.Text)
        rsPuntosObtenidos!intnumpaciente = vllngNumeroPaciente
        rsPuntosObtenidos!mnySubtotalFactura = Val(Format(txtSubtotal.Text, ""))
        rsPuntosObtenidos!intpuntospaciente = rsRango!intpuntospaciente
        rsPuntosObtenidos!intcvemedicolab = vllngCveMedicoLab
        rsPuntosObtenidos!intpuntosMedicoLab = IIf(vllngCveMedicoLab = 0, 0, rsRango!intPuntosMedico)
        rsPuntosObtenidos!intcvemedicoImagen = vllngCveMedicoImagen
        rsPuntosObtenidos!intpuntosMedicoImagen = IIf(vllngCveMedicoImagen = 0, 0, rsRango!intPuntosMedico)
        rsPuntosObtenidos!intcvemedicoCargo = vllngCveMedicoCargo
        rsPuntosObtenidos!intpuntosMedicoCargo = IIf(vllngCveMedicoCargo = 0, 0, rsRango!intPuntosMedico)
        rsPuntosObtenidos!dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
        rsPuntosObtenidos.Update
        rsPuntosObtenidos.Close
        
    End If
    rsRango.Close
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAcumularPuntos"))
End Sub



Private Sub pAplicarPuntos(lngConsecutivoFactura As Long, strFolioDocumento As String, vldtmFechaHoy As Date, vldtmHoraHoy As Date)
On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsRango As New ADODB.Recordset
    Dim rsPuntosUtilizados As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsIngreso As New ADODB.Recordset
    Dim vldblPuntosPaciente As Double
    Dim vldblPuntosMedLab As Double
    Dim vldblPuntosMedImagen As Double
    Dim vldblPuntosMedCargo As Double
    Dim vldblSaldoPuntos As Double
    Dim vldblValorPuntoLealtad As Double
    
    vldblPuntosPaciente = 0
    vldblPuntosMedLab = 0
    vldblPuntosMedImagen = 0
    vldblPuntosMedCargo = 0
    vldblSaldoPuntos = 0
    vldblValorPuntoLealtad = 0
    
    '-- 17115 --
    vlstrSentencia = "SELECT * FROM PvPuntosUtilizadosPaciente"
    Set rsPuntosUtilizados = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
       
    rsPuntosUtilizados.AddNew
    rsPuntosUtilizados!intIdFactura = lngConsecutivoFactura
    rsPuntosUtilizados!chrfoliofactura = Trim(strFolioDocumento)
    rsPuntosUtilizados!intNumCuenta = CLng(txtMovimientoPaciente.Text)
    rsPuntosUtilizados!intnumpaciente = vllngNumeroPaciente
    rsPuntosUtilizados!mnyDescuentoAplicado = vldblPuntosAplicados  'dblDescuentoAplicadoPuntos
    
    vldblSaldoPuntos = Round(vldblPuntosAplicados / fdblValorPuntoLealtad)
    
    'Consulta datos del ingreso del paciente
    vlstrSentencia = "select pi.intNumPaciente, pi.intcvetipoingreso, pi.intcvetipopaciente, nvl(pi.intcvemedicorelacionado, 0) intcvemedicorelacionado, tp.bitfamiliar, tp.chrtipo from ExPacienteIngreso pi inner join AdTipoPaciente tp on pi.intcvetipopaciente = tp.tnycvetipopaciente where pi.intNumCuenta = " & CLng(txtMovimientoPaciente.Text)
    Set rsIngreso = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
    If rsIngreso.RecordCount > 0 Then
        'Si el paciente es médico o familiar de médico
        If (rsIngreso!bitFamiliar = 1 And rsIngreso!chrTipo = "ME") Or rsIngreso!chrTipo = "ME" Then
            
            'Consulta puntos paciente
            vlstrSentencia = "select nvl(sum(puntosPaciente), 0) puntosPaciente from ( " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intPuntosPaciente end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intPuntosPaciente end), 0) puntosPaciente from pvPuntosObtenidosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "where pp.intNumPaciente = " & rsIngreso!intnumpaciente & " " & _
                                " union all " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intPuntosPaciente end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intPuntosPaciente end), 0) puntosPaciente from pvPuntosUtilizadosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intNumPaciente = " & rsIngreso!intnumpaciente & ")"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rs.RecordCount > 0 Then vldblPuntosPaciente = rs!puntosPaciente
            rs.Close
            
            'Consulta puntos médico relacionado laboratorio
            vlstrSentencia = "select nvl(sum(puntosMedLab), 0) puntosMedLab from ( " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicolab end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicolab end), 0) puntosMedLab from pvPuntosObtenidosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicolab = " & rsIngreso!intCveMedicoRelacionado & _
                                " union all " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicolab end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicolab end), 0) puntosMedLab from pvPuntosUtilizadosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicolab = " & rsIngreso!intCveMedicoRelacionado & ")"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rs.RecordCount > 0 Then vldblPuntosMedLab = rs!puntosMedLab
            rs.Close
        
            'Consulta puntos médico relacionado imagenología
            vlstrSentencia = "select nvl(sum(puntosmedimagen), 0) puntosmedimagen from ( " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicoImagen end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicoImagen end), 0) puntosMedImagen from pvPuntosObtenidosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicoImagen = " & rsIngreso!intCveMedicoRelacionado & _
                                " union all " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicoImagen end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicoImagen end), 0) puntosMedImagen from pvPuntosUtilizadosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicoImagen = " & rsIngreso!intCveMedicoRelacionado & ")"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rs.RecordCount > 0 Then vldblPuntosMedImagen = rs!puntosMedImagen
            rs.Close
            
            'Consulta puntos médico a cargo relacionado
            vlstrSentencia = "select nvl(sum(puntosMedCargo), 0) puntosMedCargo from ( " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicoCargo end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicoCargo end), 0) puntosMedCargo from pvPuntosObtenidosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicoCargo = " & rsIngreso!intCveMedicoRelacionado & _
                                " union all " & _
                                "select nvl(sum(case when vp.intcveventa is null then 0 else pp.intpuntosmedicoCargo end + " & _
                                               "case when fa.intconsecutivo is null then 0 else pp.intpuntosmedicoCargo end), 0) puntosMedCargo from pvPuntosUtilizadosPaciente pp " & _
                                            "left join PvFactura fa on pp.intIdFactura = fa.intconsecutivo and fa.chrestatus <> 'C' " & _
                                            "left join PvVentaPublico vp ON pp.intidfactura = vp.intCveVenta AND vp.chrTipoRecivo = 'T' AND vp.bitCancelado = 0 " & _
                                "Where pp.intcvemedicoCargo = " & rsIngreso!intCveMedicoRelacionado & ")"
            Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rs.RecordCount > 0 Then vldblPuntosMedCargo = rs!puntosMedCargo
            rs.Close
            
            'vldblSaldoPuntos = vldblPuntosDisponibles
            If vldblPuntosPaciente > 0 Then
                vldblSaldoPuntos = vldblSaldoPuntos - vldblPuntosPaciente
                rsPuntosUtilizados!intpuntospaciente = vldblPuntosPaciente
            Else
                rsPuntosUtilizados!intpuntospaciente = 0
            End If
            If vldblSaldoPuntos > 0 Then
                If vldblPuntosMedLab > 0 Then
                    If vldblSaldoPuntos >= vldblPuntosMedLab Then
                        vldblSaldoPuntos = vldblSaldoPuntos - vldblPuntosMedLab
                        rsPuntosUtilizados!intpuntosMedicoLab = vldblPuntosMedLab
                    ElseIf vldblSaldoPuntos < vldblPuntosMedLab Then
                        rsPuntosUtilizados!intpuntosMedicoLab = vldblSaldoPuntos
                        vldblSaldoPuntos = 0
                    End If
                    rsPuntosUtilizados!intcvemedicolab = rsIngreso!intCveMedicoRelacionado
                End If
            End If
            If vldblSaldoPuntos > 0 Then
                If vldblPuntosMedImagen > 0 Then
                    If vldblSaldoPuntos >= vldblPuntosMedImagen Then
                        vldblSaldoPuntos = vldblSaldoPuntos - vldblPuntosMedImagen
                        rsPuntosUtilizados!intpuntosMedicoImagen = vldblPuntosMedImagen
                    ElseIf vldblSaldoPuntos < vldblPuntosMedImagen Then
                        rsPuntosUtilizados!intpuntosMedicoImagen = vldblSaldoPuntos
                        vldblSaldoPuntos = 0
                    End If
                    rsPuntosUtilizados!intcvemedicoImagen = rsIngreso!intCveMedicoRelacionado
                End If
            End If
            If vldblSaldoPuntos > 0 Then
                If vldblPuntosMedCargo > 0 Then
                    If vldblSaldoPuntos >= vldblPuntosMedCargo Then
                        vldblSaldoPuntos = vldblSaldoPuntos - vldblPuntosMedCargo
                        rsPuntosUtilizados!intpuntosMedicoCargo = vldblPuntosMedCargo
                    ElseIf vldblSaldoPuntos < vldblPuntosMedCargo Then
                        rsPuntosUtilizados!intpuntosMedicoCargo = vldblSaldoPuntos
                        vldblSaldoPuntos = 0
                    End If
                    rsPuntosUtilizados!intcvemedicoCargo = rsIngreso!intCveMedicoRelacionado
                End If
            End If
        Else
            rsPuntosUtilizados!intpuntospaciente = vldblSaldoPuntos
        End If
        
        rsPuntosUtilizados!dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
        rsPuntosUtilizados.Update
        rsPuntosUtilizados.Close
        rsIngreso.Close
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGrabaPuntos, "APLICACIÓN DE PUNTOS DE LEALTAD", strFolioDocumento)
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pAplicarPuntos"))
    
End Sub


Private Sub pControlPuntos(lngConsecutivoFactura As Long, strFolioDocumento As String, vldtmFechaHoy As Date, vldtmHoraHoy As Date)
    If blnLicenciaLealtadCliente Then
        If Val(txtMovimientoPaciente.Text) > 0 Then
            If dblDescuentoAplicadoPuntos > 0 Then
                pAplicarPuntos lngConsecutivoFactura, strFolioDocumento, vldtmFechaHoy, vldtmHoraHoy
            End If
            pAcumularPuntos lngConsecutivoFactura, strFolioDocumento, vldtmFechaHoy, vldtmHoraHoy
        End If
    End If
End Sub


Private Sub pEtiquetaVar(strUso As String, strMensaje As String)

    lblPuntosLealtad.Caption = ""
    lblPuntosLealtad.Visible = False

    If Val(txtMovimientoPaciente.Text) <> 0 And strMensaje <> "" Then
        lblPuntosLealtad.Caption = strMensaje
        lblPuntosLealtad.Left = 4520
        lblPuntosLealtad.Top = 6670
        lblPuntosLealtad.Width = 7000
        lblPuntosLealtad.ForeColor = &HFF0000
        lblPuntosLealtad.FontBold = True
        lblPuntosLealtad.Alignment = 1
        lblPuntosLealtad.Visible = True
    End If
End Sub


Private Function fdblValorPuntoLealtad() As Double

    Dim ObjRS As New ADODB.Recordset
    Dim objSTR As String
    
    fdblValorPuntoLealtad = 0
    
    objSTR = "select nvl(vchvalor, 0) valor from siparametro where vchnombre = 'MNYVALORPUNTOLEALTAD' "
    'and INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
    Set ObjRS = frsRegresaRs(objSTR, adLockOptimistic)
    
    If ObjRS.RecordCount > 0 Then
        fdblValorPuntoLealtad = ObjRS!Valor
    End If

End Function
Private Sub pDatosPaciente(vllngNumCuenta As Long)
On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    
    vllngNumeroPaciente = 0
    vlstrSentencia = "SELECT * FROM Expacienteingreso where expacienteingreso.INTNUMCUENTA = " & vllngNumCuenta
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        vllngNumeroPaciente = rs!intnumpaciente
    End If
    If rs.State <> adStateClosed Then rs.Close
                    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDatosPaciente"))
    Unload Me
End Sub


Private Sub pIniciaVarFactAutomatica()
    vlstrNombreFactura = strNombrePOS
    vlstrDireccion = strCallePOS
    vlstrNumeroExterior = strNumeroExteriorPOS
    vlstrNumeroInterior = strNumeroInteriorPOS
    vlBitExtranjero = 0
    vlstrColonia = strColoniaPOS
    vlstrCP = strCPPOS
    llngCveCiudad = lngCveCiudadPOS
    vlstrTelefono = ""
    vlstrRFC = "XAXX010101000"
    vlstrNumRef = "0"
    vlstrTipo = "0"
End Sub


Private Function fRegistrarMovArregloCorte2(vllngCorte As Long, Optional vlblnRegistraInfoExtra As Boolean) As Long
    Dim vlintContador As Integer

On Error GoTo NotificaError
    
    fRegistrarMovArregloCorte2 = 0
    '--------------------------------------------------
    'Recorremos el arreglo de los movmimentos del corte
    '--------------------------------------------------
    For vlintContador = 0 To vlintElementosMovCorteIng - 1
        If vlarrMovCorteIngresos(vlintContador).vlblnSeExcluye = False Then 'sólo los que no fueron excluidos
           Select Case vlarrMovCorteIngresos(vlintContador).vlintTipoOperacion
                         '------------------------------
                  Case 2 'inserta en la poliza del corte
                         '------------------------------
                          pInsCortePoliza vllngCorte, _
                                          vlarrMovCorteIngresos(vlintContador).vlstrFolioDocumento, _
                                          vlarrMovCorteIngresos(vlintContador).vlstrTipoDocumento, _
                                          vlarrMovCorteIngresos(vlintContador).vllngCuentaContable, _
                                          vlarrMovCorteIngresos(vlintContador).vldblCantidad, _
                                          vlarrMovCorteIngresos(vlintContador).vlblnCargo, _
                                          vlarrMovCorteIngresos(vlintContador).vlstrTipoMovimiento, _
                                          "PVCORTEPOLIZAINGRESOS"
           End Select
        End If
    Next vlintContador
    
    fRegistrarMovArregloCorte2 = vllngCorte
         
Exit Function
NotificaError:
    fRegistrarMovArregloCorte2 = 0
End Function


Private Sub pAgregarMovArregloCorte2(vllngCorte As Long, vlstrFolioDocumentoCorte As String, vlstrTipoDocumento As String, vllngCuentaContable As Long, vldblCantidad As Double, vlblnCargo As Boolean, vlStrFechaHora As String, vlintFormaPago As Long, vldblTipoCambio As Double, vlstrfolioCheque As String, vllngCorteDocto As Long, vlintTipoOperacion As Integer, vlstrFolioDocOriginaMov As String, vlstrTipoDocOriginaMov As String, Optional vlbolEsCredito As Boolean, Optional vlstrRFC As String, Optional vlstrBancoSAT As String, Optional vlstrBancoExtranjero As String, Optional vlstrCuentaBancaria As String, Optional vldtmFecha As Date, Optional vlstrTipoMovimientoPoliza As String, Optional vlblnDescuento As Boolean)
    If intBitCuentaPuenteIngresos = 1 Then
        '----------------------------
        '- Agregado para caso 11017 -
        '----------------------------
        If vlintTipoOperacion = 0 Then
            pAgregarMovArregloCorteIngresos 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
        End If
        If vlstrTipoDocumento = "TI" Or vlstrTipoDocumento = "FA" Then
            pAgregarMovArregloCorteIngresos vllngCorte, vllngPersonaGraba, vlstrFolioDocumentoCorte, vlstrTipoDocumento, vllngCuentaContable, vldblCantidad, vlblnCargo, vlStrFechaHora, vlintFormaPago, vldblTipoCambio, vlstrfolioCheque, vllngCorteDocto, vlintTipoOperacion, vlstrFolioDocOriginaMov, vlstrTipoDocOriginaMov, vlbolEsCredito, vlstrRFC, vlstrBancoSAT, vlstrBancoExtranjero, vlstrCuentaBancaria, vldtmFecha, vlstrTipoMovimientoPoliza
            'Si el bitUtilizaCuentaPuenteIngresos = 1 y tipo de documento es "TI", no se registran asientos contables de los ingresos y se registran en cuenta puente donde se acumula (ingreso - descuento + iva + ieps)
            vldblIngresosPuente = vldblIngresosPuente + IIf(vlblnDescuento, vldblCantidad * -1, vldblCantidad)
        End If
        If vlstrTipoDocumento = "FA" Then
            pAgregarMovArregloCorte vllngCorte, vllngPersonaGraba, vlstrFolioDocumentoCorte, vlstrTipoDocumento, vllngCuentaContable, vldblCantidad, vlblnCargo, vlStrFechaHora, vlintFormaPago, vldblTipoCambio, vlstrfolioCheque, vllngCorteDocto, vlintTipoOperacion, vlstrFolioDocOriginaMov, vlstrTipoDocOriginaMov, vlbolEsCredito, vlstrRFC, vlstrBancoSAT, vlstrBancoExtranjero, vlstrCuentaBancaria, vldtmFecha, vlstrTipoMovimientoPoliza
        End If
    Else
        pAgregarMovArregloCorte vllngCorte, vllngPersonaGraba, vlstrFolioDocumentoCorte, vlstrTipoDocumento, vllngCuentaContable, vldblCantidad, vlblnCargo, vlStrFechaHora, vlintFormaPago, vldblTipoCambio, vlstrfolioCheque, vllngCorteDocto, vlintTipoOperacion, vlstrFolioDocOriginaMov, vlstrTipoDocOriginaMov, vlbolEsCredito, vlstrRFC, vlstrBancoSAT, vlstrBancoExtranjero, vlstrCuentaBancaria, vldtmFecha, vlstrTipoMovimientoPoliza
    End If
End Sub

Private Sub pAbreDatosFiscalesPacienteEmpresa()
    ' DATOS FISCALES DE LA EMPRESA O DEL CLIENTE O DEL PACIENTE CUANDO NO ES CONVENIO
       Load frmDatosFiscales
       frmDatosFiscales.vgblnMostrarUsoCFDI = True
       If txtMovimientoPaciente.Text <> "" Then
            pDatosFiscales3
       Else
          frmDatosFiscales.sstDatos.Tab = 0
          pDatosFiscalesPvParametros
       End If
       pDatosFiscales2
End Sub

Private Sub pAbreDatosFiscales()
    MsgBox SIHOMsg(581), vbInformation, "Mensaje" 'Capture los datos de la factura del paciente.
    Load frmDatosFiscales
    frmDatosFiscales.vgblnMostrarUsoCFDI = True
    
    If txtMovimientoPaciente.Text <> "" Then
        pAsignaVariablesDatosFiscales
    Else
        frmDatosFiscales.sstDatos.Tab = 0
        pDatosFiscalesPvParametros
    End If
    
    pDatosFiscales
    
    vlstrNumRef = vglngCveExtra
    vlstrTipo = vgstrTipoPaciente
    'NO HAY DESGLOCE DE IEPS PARA PACIENTES YA QUE SOLO SE FACTURAN LOS CONCEPTOS DE ASEGURADORA
    
    Unload frmDatosFiscales
    Set frmDatosFiscales = Nothing
End Sub

Private Sub pAsignaVariablesDatosFiscales()
    frmDatosFiscales.vgstrNombre = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNombreFactura))
    frmDatosFiscales.vgstrDireccion = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrDireccionFactura))
    frmDatosFiscales.vgstrNumExterior = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNumeroExteriorFactura))
    frmDatosFiscales.vgstrNumInterior = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNumeroInteriorFactura))
    frmDatosFiscales.vgstrColonia = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrColoniaFactura))
    frmDatosFiscales.vgstrCP = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrCPFactura))
    frmDatosFiscales.cboCiudad.ListIndex = IIf(Trim(txtMovimientoPaciente.Text) = "", -1, flngLocalizaCbo(frmDatosFiscales.cboCiudad, str(llngCveCiudad)))
    frmDatosFiscales.vgstrTelefono = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrTelefonoFactura))
    frmDatosFiscales.vgstrRFC = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrRFCFactura))
    frmDatosFiscales.vlstrNumRef = CStr(vglngCveExtra)
    frmDatosFiscales.vlstrTipo = vgstrTipoPaciente
    
    frmDatosFiscales.vlstrRegimenFiscal = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrRegimenFiscal))
    
    frmDatosFiscales.vgActivaSujetoaIEPS = IIf(vldblFacturaPacienteTotal > 0, False, True)
    frmDatosFiscales.vgBitSujetoaIEPS = 0
    frmDatosFiscales.sstDatos.Tab = 0
End Sub

''Dim vgintUtilizaConvenio As Integer         'Indica si la cuenta es de convenio

Private Sub pVerificaAlmacen()
    Dim vlstrSentencia As String
    Dim rsVerificaAlmacen As New ADODB.Recordset
    
    vlstrSentencia = "Select smicveDepartamento,RTrim (vchDescripcion) From noDepartamento Where chrClasificacion = 'A' and bitEstatus = 1  And smicvedepartamento = " & vgintNumeroDepartamento
    Set rsVerificaAlmacen = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsVerificaAlmacen.EOF Then
        vlstrSentencia = "select intnumalmacen from pvalmacenes where intnumdepartamento =" & vgintNumeroDepartamento
        Set rsVerificaAlmacen = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        If rsVerificaAlmacen.EOF Then
            MsgBox SIHOMsg(961), vbExclamation, "Mensaje"
            vlbsinalmacen = True
            rsVerificaAlmacen.Close
            Unload Me
        End If
    End If
End Sub

Private Sub cboMedico_GotFocus()
    On Error GoTo NotificaError
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaPersona"))
End Sub

Private Sub cboMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtClaveArticulo.Enabled = True Then
        If KeyCode = vbKeyReturn Then
            If txtClaveArticulo.Enabled Then
                txtClaveArticulo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chkFacturaSustitutaDFP_Click()
    Dim i As Integer

 If Not vlnblnLocate Then
    If chkFacturaSustitutaDFP.Value = 0 Then
        lstFacturaASustituirDFP.Clear
        ReDim aFoliosPrevios(0)
    Else
        If chkFacturaSustitutaDFP.Value = 1 Then
        
            frmBusquedaFacturasPreviasVP.vlstrsql = vlstrsql
            frmBusquedaFacturasPreviasVP.Show vbModal, Me
            
            lstFacturaASustituirDFP.Clear
            For i = 0 To UBound(aFoliosPrevios())
                If aFoliosPrevios(i).chrfoliofactura <> "" Then
                    lstFacturaASustituirDFP.AddItem aFoliosPrevios(i).chrfoliofactura
                End If
            Next i
            
            If frmBusquedaFacturasPreviasVP.vlchrfoliofactura = "" Then
                chkFacturaSustitutaDFP.Value = 0
            End If
        End If
    End If
  End If

End Sub

Private Sub cmdAceptarSeleccionTickets_Click()
    Dim intcontador As Long
    Dim intContador2 As Long
    Dim vlIntCont As Integer
    Dim vlblnTicketSeleccionado As Boolean
    Dim vlblnTerminoInformacion As Boolean
    Dim vlintContador As Integer
    Dim vlintContador2 As Integer
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim lstrClaveSinEquivalencia As String
    
    Dim vldblImporte As Double
    Dim vldblDescuentos As Double
    Dim vldblIEPS As Double
    Dim vldblTotalIVA As Double
    
    ReDim aTickets(0)
    
    intContador2 = 1
    For intcontador = 0 To lstFoliosTickets.ListCount - 1
        If lstFoliosTickets.Selected(intcontador) Then
            ReDim Preserve aTickets(intContador2)
            
            lstFoliosTickets.ListIndex = intcontador
            
            aTickets(intContador2).lngTicket = lstFoliosTickets.Text
            aTickets(intContador2).dtmfecha = CDate(Trim(hoja.Cells(intContador2 + 1, 4)))
            intContador2 = intContador2 + 1
        End If
    Next intcontador
    intContador2 = intContador2 - 1
    
    If intContador2 = 0 Then
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    lstrClaveSinEquivalencia = ""
    
    vlblnTerminoInformacion = False
    vlintContador = 1
    vlintContador2 = 1
    
    vldblImporte = 0
    vldblDescuentos = 0
    vldblIEPS = 0
    vldblTotalIVA = 0
    
    Do While vlblnTerminoInformacion = False
        If Trim(hoja.Cells(vlintContador + 1, 1)) = "" Then
            vlblnTerminoInformacion = True
        Else
            grdArticulos.Redraw = False

            vlblnTicketSeleccionado = False
            For vlIntCont = 1 To intContador2
                If Trim(aTickets(vlIntCont).lngTicket) = CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 1)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 1)))) Then
                    vlblnTicketSeleccionado = True
                End If
            Next vlIntCont

            If vlblnTicketSeleccionado = True Then

                SQL = "SELECT SIEQUIVALENCIADETALLE.*, PVOTROCONCEPTO.CHRDESCRIPCION, PVOTROCONCEPTO.SMICONCEPTOFACT, PVCONCEPTOFACTURACION.CHRDESCRIPCION DESCRIPCIONCONCEPTOFACT " & _
                        "From SIEQUIVALENCIADETALLE " & _
                            "INNER JOIN SIEQUIVALENCIA ON SIEQUIVALENCIA.INTCVEEQUIVALENCIA = SIEQUIVALENCIADETALLE.INTCVEEQUIVALENCIA " & _
                            "INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = SIEQUIVALENCIADETALLE.VCHCVELOCAL " & _
                            "INNER JOIN PVCONCEPTOFACTURACION ON PVOTROCONCEPTO.SMICONCEPTOFACT = PVCONCEPTOFACTURACION.SMICVECONCEPTO " & _
                        "WHERE SIEQUIVALENCIA.VCHDESCRIPCION = 'OTROS CONCEPTOS DE CARGO' " & _
                            "AND trim(VCHCVEEXTERNA) = '" & Trim(hoja.Cells(vlintContador + 1, 2)) & "'"
                Set rs = frsRegresaRs(SQL)
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    
                    vlblnImportoVenta = True
            
                    If grdArticulos.RowData(1) <> -1 Then
                        grdArticulos.Rows = grdArticulos.Rows + 1
                        grdArticulos.Row = grdArticulos.Rows - 1
                    End If
                    
                    'grdArticulos.TextMatrix(vlintContador2, 1) = Trim(hoja.Cells(vlintContador + 1, 3)) 'Descripcion
                    grdArticulos.TextMatrix(vlintContador2, 1) = Trim(rs!chrDescripcion) 'Descripcion
                    
                    grdArticulos.TextMatrix(vlintContador2, 2) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 5)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 5)))), "$###,###,###,###.00") 'Precio
                    grdArticulos.TextMatrix(vlintContador2, 24) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 5)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 5)))), "$###,###,###,###.00") 'Precio
                    
                    grdArticulos.TextMatrix(vlintContador2, 3) = 1 'Cantidad
                    grdArticulos.TextMatrix(vlintContador2, 4) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 5)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 5)))), "$###,###,###,###.00") 'Importe
                    vldblImporte = vldblImporte + Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 5)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 5)))), 2)
                    
                    grdArticulos.TextMatrix(vlintContador2, 5) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 6)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 6)))), "$###,###,###,###.00") 'Descuento
                    grdArticulos.TextMatrix(vlintContador2, 25) = Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 6)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 6)))), 2) 'Descuento
                    vldblDescuentos = vldblDescuentos + Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 6)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 6)))), 2)
                    
                    If vlblnLicenciaIEPS Then
                        grdArticulos.TextMatrix(vlintContador2, 6) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7)))), "$###,###,###,###.00") 'IEPS
                        vldblIEPS = vldblIEPS + Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7)))), 2)
                    Else
                        grdArticulos.TextMatrix(vlintContador2, 6) = Format(0, "$###,###,###,###.00") 'IEPS
                    End If
                    
                    If vlblnLicenciaIEPS Then
                        grdArticulos.TextMatrix(vlintContador2, 7) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 10)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 10)))), "$###,###,###,###.00") 'Subtotal
                    Else
                        grdArticulos.TextMatrix(vlintContador2, 7) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 10)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 10)))) - CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7)))), "$###,###,###,###.00") 'Subtotal
                    End If
                    
                    If CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))) = 0 Then
                        grdArticulos.TextMatrix(vlintContador2, 8) = Format(0, "$###,###,###,###.00") 'IVA
                        grdArticulos.TextMatrix(vlintContador2, 26) = 0 'IVA
                    Else
                        If vlblnLicenciaIEPS Then
                            grdArticulos.TextMatrix(vlintContador2, 8) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), "$###,###,###,###.00") 'IVA
                            grdArticulos.TextMatrix(vlintContador2, 26) = Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), 2) 'IVA
                            vldblTotalIVA = vldblTotalIVA + Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), 2)
                        Else
                            If CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7)))) = 0 Then
                                grdArticulos.TextMatrix(vlintContador2, 8) = Format(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), "$###,###,###,###.00") 'IVA
                                grdArticulos.TextMatrix(vlintContador2, 26) = Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), 2) 'IVA
                                vldblTotalIVA = vldblTotalIVA + Round(CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 11)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 11)))), 2)
                            Else
                                grdArticulos.TextMatrix(vlintContador2, 8) = Format((CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 10)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 10)))) - CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7))))) * (CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 13)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 13)))) / 100), "$###,###,###,###.00") 'IVA
                                grdArticulos.TextMatrix(vlintContador2, 26) = Round((CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 10)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 10)))) - CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7))))) * (CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 13)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 13)))) / 100), 2) 'IVA
                                vldblTotalIVA = vldblTotalIVA + Round((CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 10)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 10)))) - CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 7)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 7))))) * (CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 13)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 13)))) / 100), 2)
                            End If
                        End If
                    End If
                    
                    grdArticulos.TextMatrix(vlintContador2, 9) = Format(CDbl(grdArticulos.TextMatrix(vlintContador2, 7)) + CDbl(grdArticulos.TextMatrix(vlintContador2, 8)), "$###,###,###,###.00") 'Total
                    grdArticulos.Col = 9
                    grdArticulos.CellFontBold = True
                    
                    'grdArticulos.TextMatrix(vlintContador2, 10) = Trim(hoja.Cells(vlintContador + 1, 2))
                    grdArticulos.TextMatrix(vlintContador2, 10) = Trim(rs!VCHCVELOCAL)
                    
                    grdArticulos.TextMatrix(vlintContador2, 11) = CDbl(grdArticulos.TextMatrix(vlintContador2, 8)) 'IVA
                    
                    'grdArticulos.TextMatrix(vlintContador2, 12) = UCase(Mid(Trim(hoja.Cells(vlintContador + 1, 2)), 1, 2)) 'Tipo de cargo
                    grdArticulos.TextMatrix(vlintContador2, 12) = "OC" 'Tipo de cargo
                    
                    grdArticulos.TextMatrix(vlintContador2, 13) = Trim(rs!SMICONCEPTOFACT) 'Concepto de facturacion
                    grdArticulos.TextMatrix(vlintContador2, 14) = 0 'Descuento por unidad alterna(2) o unidad mínima(1)
                    grdArticulos.TextMatrix(vlintContador2, 16) = 0 'Contenido de IVArticulo
                    grdArticulos.TextMatrix(vlintContador2, 17) = 0 'Clave de la lista de precios
                    grdArticulos.TextMatrix(vlintContador2, 18) = CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 6)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 6)))) / CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 5)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 5)))) 'Porcentaje de descuento
                    
                    If vlblnLicenciaIEPS Then
                        grdArticulos.TextMatrix(vlintContador2, 19) = CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 9)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 9)))) / 100 'Porcentaje de IEPS
                    Else
                        grdArticulos.TextMatrix(vlintContador2, 19) = 0 'Porcentaje de IEPS
                    End If
                    
                    grdArticulos.TextMatrix(vlintContador2, 20) = 0 '2 = tiene exclusión de descuento, 3 = no tiene exclusion de descuento
                    grdArticulos.TextMatrix(vlintContador2, 21) = Trim(rs!DESCRIPCIONCONCEPTOFACT) 'Descripción del concepto de Facturación
                    grdArticulos.TextMatrix(vlintContador2, 22) = Trim(hoja.Cells(vlintContador + 1, 1)) 'Folio del ticket por cada concepto de la factura global
                    grdArticulos.RowData(vlintContador2) = Trim(rs!VCHCVELOCAL)
                    grdArticulos.TextMatrix(vlintContador2, 23) = "0"
                    
                    grdArticulos.Redraw = True
                    grdArticulos.Refresh
'                    pCalculaTotales
                    
                    vlintContador2 = vlintContador2 + 1
                Else
                    lstrClaveSinEquivalencia = lstrClaveSinEquivalencia & IIf(lstrClaveSinEquivalencia = "", Trim(hoja.Cells(vlintContador + 1, 2)), Chr(13) & Trim(hoja.Cells(vlintContador + 1, 2)))
                End If
                rs.Close
            End If

            vlintContador = vlintContador + 1
        End If
    Loop
    
    txtImporte.Text = Format(vldblImporte, "$###,###,###,###.00")
    txtDescuentos.Text = Format(vldblDescuentos, "$###,###,###.00")
    txtSubtotal.Text = Format(Round(vldblImporte, 2) - Round(vldblDescuentos, 2) + Round(vldblIEPS, 2), "$###,###,###,###.00")
    txtIEPS.Text = Format(vldblIEPS, "$###,###,###,###.00")
    txtIva.Text = Format(vldblTotalIVA, "$###,###,###,###.00")
    txtTotal.Text = Format(Round(vldblImporte, 2) - Round(vldblDescuentos, 2) + Round(vldblIEPS, 2) + Round(vldblTotalIVA, 2), "$###,###,###,###.00")
                                  
    xlsApp.Quit
    Set xlsApp = Nothing
    
    f.Close
    
    FreDetalle.Enabled = True
    frmMedico.Enabled = True
    freControlAseguradora.Enabled = True
    freBusqueda.Enabled = True
    
    grdArticulos.SetFocus
    frmFoliosVentaImportada.Visible = False
    
    If vlblnImportoVenta = False Then
        txtClaveArticulo.Enabled = True
        txtCantidad.Enabled = True
        optArticulo.Enabled = True
        optOtrosConceptos.Enabled = True
                    
        freDatosPaciente.Enabled = True
        freConsultaPrecios.Enabled = True
        freDescuentos.Enabled = True
        
        cmdImportarVenta.Enabled = True
        
        freGraba.Enabled = True
        freFacturaPesosDolares.Enabled = True
    End If
    
    grdArticulos.Redraw = True

    
    If lstrClaveSinEquivalencia <> "" Then
        'Los siguientes conceptos encontrados en el archivo no cuentan con su respectiva equivalencia, favor de verificar.
        MsgBox SIHOMsg(1510) & Chr(13) & Trim(lstrClaveSinEquivalencia), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub

Private Sub cmdCierraControlAseguradora_Click()
    FreDetalle.Enabled = True
    freGraba.Enabled = True
    freDescuentos.Enabled = True
    freConsultaPrecios.Enabled = True
    freFacturaPesosDolares.Enabled = True
    freControlAseguradora.Visible = False
    If txtClaveArticulo.Enabled Then
        txtClaveArticulo.SetFocus
    End If
End Sub


Private Sub pConsultaDepto()
   On Error GoTo NotificaError
    
    Dim rsDepto As New ADODB.Recordset 'Declaramos un nuevo recordset para asignar el departamento
    Dim vlstrDepto As String ' string que contiene la consulta SQL, que despues se asigna al procedimiento que regresa la consulta
    
        vlstrDepto = "SELECT intcvedepartamento FROM PvVentaPublicoNoFacturado WHERE intCveVenta =" + Me.txtTicketPrevio
        Set rsDepto = frsRegresaRs(vlstrDepto)
        vlstrClaveDepartamento = rsDepto!intCveDepartamento
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConsultaDepto"))

End Sub

Private Sub cmdConsultaTicketsNoFacturados_Click()
    Dim rsSelTicketBusqueda As New ADODB.Recordset
    Dim rsSelTicket As New ADODB.Recordset
    Dim vlintContador As Integer
    Dim vlStrTipoPaciente As String
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim blnContinuar As Boolean
    Dim i As Long
    
   
    pNuevo
    If Val(txtTicketPrevio.Text) = 0 Or Val(txtTicketPrevio.Text) = fSiguienteClaveVentaPublicoNoFacturado Then
    
        Set rsSelTicketBusqueda = frsEjecuta_SP(str(vgintClaveEmpresaContable), "SP_PVSELVENTASPUBLICONOFACT")
        If rsSelTicketBusqueda.RecordCount > 0 Then
    
            frmVentaPublicoNoFacturado.Show vbModal, Me
            
            If Trim(txtTicketPrevio.Text) = "" Then
                txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
                txtTicketPrevio.SetFocus
                Exit Sub
            End If
        Else
            'Escriba el folio o numero de referencia.
            MsgBox Replace(SIHOMsg(295), " del documento", ""), vbOKOnly + vbInformation, "Mensaje"
            txtTicketPrevio.SetFocus
            Exit Sub
        End If
    End If
    
    Set rsSelTicket = frsEjecuta_SP(Val(txtTicketPrevio.Text) & "|" & IIf(vlblnLicenciaIEPS, 1, 0), "SP_PVSELTICKETNOFACTURADO")

    If rsSelTicket.RecordCount > 0 Then

        If rsSelTicket!bitcancelado = 1 Then
            'La venta ya ha sido guardada
            MsgBox SIHOMsg(1426), vbInformation, "Mensaje"
            vlblnConsultaTicketPrevio = False
            
            txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
            
            txtTicketPrevio.SetFocus
            pEnfocaTextBox txtTicketPrevio
        Else
            vlblnConsultaTicketPrevio = True
            freDatosPaciente.Enabled = False
            frmMedico.Enabled = False
            vlintContador = 1
            pConsultaDepto
            
            blnContinuar = vlstrClaveDepartamento = vgintNumeroDepartamento
            
            If Not blnContinuar Then
                blnContinuar = MsgBox("El folio es de otro departamento. ¿Desea continuar?", vbQuestion + vbYesNo, "Mensaje") = vbYes
            End If
            
            If blnContinuar Then
           
                grdArticulos.Redraw = False
           
                Do While Not rsSelTicket.EOF
                    With rsSelTicket
                    
                    lblPaciente.Caption = !CLIENTE
                    optPaciente.Value = IIf(!INTMOVPACIENTE = 0, 0, 1)
                    optMedico.Value = IIf(!intCveMedico = 0, 0, 1)
                    optEmpleado.Value = IIf(!intCveEmpleado = 0, 0, 1)
                    If !INTMOVPACIENTE <> 0 Then
                        txtMovimientoPaciente.Text = !INTMOVPACIENTE
                    ElseIf !intCveMedico <> 0 Then
                        txtMovimientoPaciente.Text = !intCveMedico
                    ElseIf !intCveEmpleado <> 0 Then
                        txtMovimientoPaciente.Text = !intCveEmpleado
                    End If
                    If Val(txtMovimientoPaciente.Text) = 0 Then
                        txtMovimientoPaciente.Text = ""
                        optPaciente.Value = True
                        lblEmpresa.Caption = "PARTICULAR"
                    Else
                        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                    End If
                    
                    optArticulo.Value = IIf(!chrTipoCargo = "AR", True, False)
                    optOtrosConceptos.Value = IIf(!chrTipoCargo = "OC", True, False)
        
                    If grdArticulos.RowData(1) <> -1 Then
                        grdArticulos.Rows = grdArticulos.Rows + 1
                        grdArticulos.Row = grdArticulos.Rows - 1
                    End If
        
                    grdArticulos.TextMatrix(vlintContador, 1) = !Descripcion
                    
                    grdArticulos.TextMatrix(vlintContador, 2) = Format(!mnyPrecio, "$###,###,###,###.00")
                    grdArticulos.TextMatrix(vlintContador, 24) = !mnyPrecio
                    
                    grdArticulos.TextMatrix(vlintContador, 3) = !intCantidad
                    grdArticulos.TextMatrix(vlintContador, 4) = Format(!mnyPrecio * CLng(!intCantidad), "$###,###,###,###.00")
                    grdArticulos.TextMatrix(vlintContador, 5) = Format(!MNYDESCUENTO, "$###,###,###,###.00")
                    grdArticulos.TextMatrix(vlintContador, 6) = Format(!IEPS_ELEMENTO, "$###,###,###,###.00")  '@
                    grdArticulos.TextMatrix(vlintContador, 7) = Format((!mnyPrecio * !intCantidad) - !MNYDESCUENTO + !IEPS_ELEMENTO, "$###,###,###,###.00") '@
                    grdArticulos.TextMatrix(vlintContador, 8) = Format(!MNYIVA, "$###,###,###,###.00")
                    grdArticulos.TextMatrix(vlintContador, 9) = Format((!mnyPrecio * !intCantidad) - !MNYDESCUENTO + !IEPS_ELEMENTO + !MNYIVA, "$###,###,###,###.00")
                    grdArticulos.Col = 9
                    grdArticulos.CellFontBold = True
                    grdArticulos.TextMatrix(vlintContador, 10) = !intCveCargo
                    grdArticulos.TextMatrix(vlintContador, 11) = !MNYIVA
                    grdArticulos.TextMatrix(vlintContador, 12) = !chrTipoCargo
                    grdArticulos.TextMatrix(vlintContador, 13) = !smiCveConceptoFacturacion  'Clave del concepto de Facturación
                    grdArticulos.TextMatrix(vlintContador, 14) = !IntModoDescuentoInventario 'Descuento por unidad alterna(2) o unidad mínima(1)
                    grdArticulos.TextMatrix(vlintContador, 19) = !numporcentajeieps / 100 'Porcentaje de IEPS
                    grdArticulos.TextMatrix(vlintContador, 20) = !intAuxExclusionDescuento '2 = tiene exclusión de descuento, 3 = no tiene exclusion de descuento
                    grdArticulos.TextMatrix(vlintContador, 23) = !bitPrecioManual   'Indica si se modificó el precio
                    grdArticulos.RowData(grdArticulos.Row) = !intCveCargo
                    pCalculaTotales
                    vlintContador = vlintContador + 1
                    .MoveNext
                    End With
                 Loop
                 
                 grdArticulos.Redraw = True
        
             Else
                pNuevo
                txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
                pEnfocaTextBox txtTicketPrevio
            End If
        End If
    Else
        If Trim(txtTicketPrevio.Text) = fSiguienteClaveVentaPublicoNoFacturado Then
            If grdArticulos.Enabled Then
                grdArticulos.SetFocus
            End If
        Else
            MsgBox Replace(SIHOMsg(550), "documento", "folio"), vbInformation, "Mensaje" '"Número de folio no encontrado"
            vlblnConsultaTicketPrevio = False
            txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
            txtTicketPrevio.SetFocus
            pEnfocaTextBox txtTicketPrevio
        End If
    End If
End Sub

Private Sub cmdControlAseguradora_Click()
    Dim rsControl As New ADODB.Recordset
    Dim rsControlEmpresa As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    ' Que no salga la pantalla cuando la cuenta este en cero
    If Val(Format(txtTotal.Text)) = 0 Then Exit Sub
    
    vlstrSentencia = "SELECT * FROM pvControlAseguradora " & _
                     " WHERE intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & _
                     " AND chrTipoPaciente = 'E'" & _
                     " AND (chrFolioFacturaDeducible is null or chrFolioFacturaCoaseguro is null or chrFolioFacturaCopago is null or chrFolioFacturaExcedente is null or chrFolioFacturaEmpresa is null) " & _
                     " AND intCveEmpresa = " & str(vgintEmpresa)
                    
    Set rsControl = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsControl.RecordCount > 0 Then 'Que SI se encuentra
        With rsControl
            'Deducible:
            optTipoDeducible(0).Value = Trim(!CHRTIPODEDUCIBLE) = "C"
            optTipoDeducible(1).Value = Trim(!CHRTIPODEDUCIBLE) = "P"
            optTipoDeducible_Click IIf(optTipoDeducible(0).Value, 0, 1)
            txtDeducible.Text = FormatNumber(IIf(optTipoDeducible(0).Value, !MNYCANTIDADDEDUCIBLE, !NUMPORCENTAJEDEDUCIBLE), 2)
            chkFacturarDeducible.Value = !BITFACTURADEDUCIBLE
            'Coaseguro:
            optTipoCoaseguro(0).Value = Trim(!CHRTIPOCOASEGURO) = "C"
            optTipoCoaseguro(1).Value = Trim(!CHRTIPOCOASEGURO) = "P"
            optTipoCoaseguro_Click IIf(optTipoCoaseguro(0).Value, 0, 1)
            txtCoaseguro.Text = FormatNumber(IIf(optTipoCoaseguro(0).Value, !MNYCANTIDADCOASEGURO, !NUMPORCENTAJECOASEGURO), 2)
            chkFacturarCoaseguro.Value = !BITFACTURACOASEGURO
            'Copago:
            optCopagoCantidad.Value = Trim(!CHRTIPOCOPAGO) = "C"
            optControlCopagoPorciento.Value = Trim(!CHRTIPOCOPAGO) = "P"
            If optCopagoCantidad.Value Then
                optCopagoCantidad_Click
            Else
                optControlCopagoPorciento_Click
            End If
            txtCopago.Text = FormatNumber(IIf(optCopagoCantidad.Value, !MNYCANTIDADCOPAGO, !NUMPORCENTAJECOPAGO), 2)
            chkFacturarCopago.Value = !BITFACTURACOPAGO
            vlchrIncluirConceptosSeguro = IIf(IsNull(!chrIncluirConceptosSeguro), "I", !chrIncluirConceptosSeguro)
        End With
    Else
        'Poner el último tipo de deducible, coaseguro y copago usado esta aseguradora:
        Set rsControlEmpresa = frsEjecuta_SP(str(vgintEmpresa), "SP_PVSELCONTROLASEGURADORAEMPR")
        If rsControlEmpresa.RecordCount <> 0 Then
            optTipoDeducible(0).Value = Trim(rsControlEmpresa!CHRTIPODEDUCIBLE) = "C"
            optTipoDeducible(1).Value = Trim(rsControlEmpresa!CHRTIPODEDUCIBLE) = "P"
        
            optTipoCoaseguro(0).Value = Trim(rsControlEmpresa!CHRTIPOCOASEGURO) = "C"
            optTipoCoaseguro(1).Value = Trim(rsControlEmpresa!CHRTIPOCOASEGURO) = "P"
        
            optCopagoCantidad.Value = Trim(rsControlEmpresa!CHRTIPOCOPAGO) = "C"
            optControlCopagoPorciento.Value = Trim(rsControlEmpresa!CHRTIPOCOPAGO) = "P"
            
            chkFacturarDeducible.Value = rsControlEmpresa!BITFACTURADEDUCIBLE
            chkFacturarCoaseguro.Value = rsControlEmpresa!BITFACTURACOASEGURO
            chkFacturarCopago.Value = rsControlEmpresa!BITFACTURACOPAGO
            vlchrIncluirConceptosSeguro = IIf(IsNull(rsControlEmpresa!chrIncluirConceptosSeguro), "I", rsControlEmpresa!chrIncluirConceptosSeguro)
        Else
            optTipoDeducible(0).Value = True
            optTipoCoaseguro(0).Value = True
            optCopagoCantidad.Value = True
            chkFacturarDeducible.Value = 0
            chkFacturarCoaseguro.Value = 0
            chkFacturarCopago.Value = 0
        End If
        rsControlEmpresa.Close
        
        If optTipoDeducible(0).Value Then
            optTipoDeducible_Click 0
        Else
            optTipoDeducible_Click 1
        End If
        If optTipoCoaseguro(0).Value Then
            optTipoCoaseguro_Click 0
        Else
            optTipoCoaseguro_Click 1
        End If
        If optCopagoCantidad.Value Then
            optCopagoCantidad_Click
        Else
            optControlCopagoPorciento_Click
        End If
        
        txtDeducible.Text = ""
        txtCoaseguro.Text = ""
        txtCopago.Text = ""
    End If
    rsControl.Close

    freControlAseguradora.Visible = True
    pEnfocaTextBox txtDeducible
    
    FreDetalle.Enabled = False
    freGraba.Enabled = False
    freDescuentos.Enabled = False
    freConsultaPrecios.Enabled = False
    freFacturaPesosDolares.Enabled = False
        
End Sub

Private Sub cmdControlAseguradora_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= cmdControlAseguradora.Left And _
        X <= cmdControlAseguradora.Left + cmdControlAseguradora.Width And _
        Y >= cmdControlAseguradora.Top And _
        Y >= cmdControlAseguradora.Top + cmdControlAseguradora.Height Then
        
        If Val(Format(txtTotal.Text, "")) = 0 Then
            cmdControlAseguradora.Enabled = False
        End If
    End If
End Sub

Private Sub cmdDescuenta_Click()
    Dim vlintContador As Integer
    Dim vldblDescuento As Double
    Dim vldblDesctoAplicado As Double
    Dim vldblPrecio As Double
    Dim vlintCantidad As Double
    Dim vldblSubtotal As Double
    Dim vldblIVA As Double
    Dim vlaryParametrosSalida() As String
    Dim vldblTotDescuento As Double
    Dim vldblPorcentajeD As Double
    Dim vlstrMensaje As String
    Dim vllngSeleccionados As Long
    Dim vllngDescontados As Long
    '---------------------------------------------------
    ' Tiene que tener Permisos para dar descuentos
    '---------------------------------------------------
    If Not fblnRevisaPermiso(vglngNumeroLogin, 382, "C") Then
        '---------------------------------------------------
        ' Para que sólo los logins autorizados puedan buscar
        ' por descripción (sólo artículos)
        ' y que pregute cuando el Login inicial no tenga permiso
        '---------------------------------------------------
        Load frmLogin
        frmLogin.vgblnCargaVariablesGlobales = False
        frmLogin.Show vbModal
        If frmLogin.vgintLogin = -1 Or Not fblnRevisaPermiso(frmLogin.vgintLogin, 382, "C") Then
            MsgBox SIHOMsg(635), vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
    
    vldblDescuento = Val(txtDescuento.Text) / 100
    
    If vldblDescuento > 1 Then
        MsgBox SIHOMsg("400"), vbExclamation, "Mensaje"
        txtDescuento.Text = ""
        txtDescuento.SetFocus
        Exit Sub
    End If
    
    If grdArticulos.RowData(1) <> -1 Then
        grdArticulos.Redraw = False
        vlstrMensaje = ""
        vllngSeleccionados = 0
        vllngDescontados = 0
               
        For vlintContador = 1 To grdArticulos.Rows - 1
            If grdArticulos.TextMatrix(vlintContador, 0) = "*" Then
               vllngSeleccionados = vllngSeleccionados + 1
               With grdArticulos
                    '------------------------------------
                    'Se revisa la exclusión del descuento
                    '------------------------------------
                    If .TextMatrix(vlintContador, 20) = "2" Then
                       If vlstrMensaje = "" Then
                          vlstrMensaje = .TextMatrix(vlintContador, 1)
                       Else
                          vlstrMensaje = vlstrMensaje & vbNewLine & .TextMatrix(vlintContador, 1)
                       End If


                    Else
                        vllngDescontados = vllngDescontados + 1
                        vlintCantidad = Val(.TextMatrix(vlintContador, 3))
                        
                        If CDbl(.TextMatrix(vlintContador, 24)) <> 0 Then
                            vldblDesctoAplicado = (Val(Format(.TextMatrix(vlintContador, 24), "")) * vldblDescuento * vlintCantidad)
                        Else
                            vldblDesctoAplicado = (Val(Format(.TextMatrix(vlintContador, 2), "")) * vldblDescuento * vlintCantidad)
                        End If
                        
                        
                        .Row = vlintContador
                        '-----------------------
                        'Descuentos
                        '-----------------------
                        '-----------------------
                        ' Procedimiento para obtener el IVA
                        '-----------------------
                        vldblIVA = fdblObtenerIva(.RowData(.Row), .TextMatrix(.Row, 12)) / 100
                        If Val(.TextMatrix(vlintContador, 18)) > 0 Then
                            vldblPorcentajeD = Val(.TextMatrix(vlintContador, 18)) + Val(txtDescuento.Text)
                        Else
                            vldblPorcentajeD = IIf(Val(Format(.TextMatrix(vlintContador, 5), "")) > 0, Format((vldblDesctoAplicado * 100) / .TextMatrix(vlintContador, 4), "###.00"), Val(txtDescuento.Text))
                        End If
                        .TextMatrix(.Row, 5) = Format(vldblDesctoAplicado, "$###,###,###,###.00")
                        
                        If CDbl(.TextMatrix(vlintContador, 24)) <> 0 Then
                            vldblSubtotal = Round((Val(Format(.TextMatrix(vlintContador, 24), "")) * vlintCantidad), 2) - Val(Format(.TextMatrix(.Row, 5), "############.00")) 'vldblDescuento
                        Else
                            vldblSubtotal = (Val(Format(.TextMatrix(vlintContador, 2), "")) * vlintCantidad) - Val(Format(.TextMatrix(.Row, 5), "############.00")) 'vldblDescuento
                        End If

                        .TextMatrix(.Row, 6) = Format(vldblSubtotal * Val(Format(.TextMatrix(.Row, 19), "")), "$###,###,###,###.00") 'IEPS@
                        vldblSubtotal = vldblSubtotal + (vldblSubtotal * Val(Format(.TextMatrix(.Row, 19), ""))) 'IEPS + SUBTOTAL
                        .TextMatrix(.Row, 7) = Format(vldblSubtotal, "$###,###,###,###.00") '@
                        .TextMatrix(.Row, 8) = Format(vldblSubtotal * vldblIVA, "$###,###,###,###.00") '
                        .TextMatrix(.Row, 9) = Format(vldblSubtotal * (vldblIVA + 1), "$###,###,###,###.00") '
                        .Col = 9
                        .CellFontBold = True
                        .TextMatrix(.Row, 11) = vldblSubtotal * vldblIVA
                        .TextMatrix(vlintContador, 0) = ""
                        .TextMatrix(.Row, 18) = vldblPorcentajeD '
                    End If
                End With
                pCalculaTotales
            End If
        Next
        
        If vllngSeleccionados > 0 Then 'hay seleccionados
           If vllngDescontados = 0 Then ' no hay descuentos por que hay exclusiones
              'Mensaje: No se puede aplicar el descuento debido a que los articulos cuentan con exclusion de descuento.
               MsgBox SIHOMsg(1273), vbOKOnly + vbExclamation, "Mensaje"
           Else
               If vllngDescontados < vllngSeleccionados Then
                  'Mensaje: No se puede aplicar el descuento en los siguientes artículos debido a que se encontró con una exclusión de descuento:
                   MsgBox SIHOMsg(1274) & vbNewLine & vlstrMensaje, vbOKOnly + vbExclamation, "Mensaje"
               End If
           End If
        End If
               
        grdArticulos.Redraw = True
        grdArticulos.Refresh
        vl_dblClickCmdDescuenta = True
        If fblnCanFocus(grdArticulos) Then grdArticulos.SetFocus
        grdArticulos.Col = 1
    End If
End Sub

Private Function fblnControlValido() As Boolean

    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    fblnControlValido = True
    
    '--------------------------------------------
    'Validación de que exista los conceptos para DEDUCIBLE, COASEGURO Y COPAGO
    '--------------------------------------------
    vlstrSentencia = "SELECT ISNULL(intConceptoDeducible,0) ConceptoDeducible, " & _
                     " ISNULL(intConceptoCoaseguro,0) ConceptoCoaseguro, " & _
                     " ISNULL(intConceptoCoPago,0) ConceptoCoPago " & _
                     " FROM pvParametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount = 0 Then
        fblnControlValido = False
        'No se tiene asignado el concepto de facturación para el deducible y coaseguro.
        MsgBox SIHOMsg(371), vbCritical, "Mensaje"
    End If
    If fblnControlValido Then
        If rs!ConceptoDeducible = 0 Or rs!ConceptoCoaseguro = 0 Or rs!ConceptoCoPago = 0 Then
            fblnControlValido = False
            'No se tiene asignado el concepto de facturación para el deducible y coaseguro.
            MsgBox SIHOMsg(371), vbCritical, "Mensaje"
        Else
            llngCveConceptoDeducible = rs!ConceptoDeducible
            llngCveConceptoCoaseguro = rs!ConceptoCoaseguro
            llngCveConceptoCopago = rs!ConceptoCoPago
        End If
    End If

End Function

Private Sub cmdDescuentoPuntos_Click()
    
    txtDescuento.Text = ""
    If Val(txtMovimientoPaciente.Text) > 0 Then
        If cmdDescuentoPuntos.Caption = "Aplicar puntos de cliente leal" Then
            If Val(Format(txtSubtotal.Text, "")) > 0 Then
                pAplicarDescuentoPuntos
                cmdDescuentoPuntos.Caption = "Desaplicar puntos de cliente leal"
            End If
        ElseIf cmdDescuentoPuntos.Caption = "Desaplicar puntos de cliente leal" Then
            pDesAplicarDescuentoPuntos
            cmdDescuentoPuntos.Caption = "Aplicar puntos de cliente leal"
        End If
    End If
    
End Sub
Private Sub pRellenaInfoGlobal()
    
    Dim i As Integer
    Dim j As Integer
    Dim dtmFechaMinima As Date
    Dim dtmFechaMaxima As Date
    Dim intDiferencia As Date
    
    i = 1
    j = 1
    
    dtmFechaMinima = aTickets(i).dtmfecha
    For i = 1 To UBound(aTickets())
        If dtmFechaMinima > aTickets(i).dtmfecha Then
            dtmFechaMinima = aTickets(i).dtmfecha
        End If
    Next i

    dtmFechaMaxima = aTickets(j).dtmfecha
    For j = 1 To UBound(aTickets())
        If dtmFechaMaxima < aTickets(j).dtmfecha Then
            dtmFechaMaxima = aTickets(j).dtmfecha
        End If
    Next j

    intDiferencia = DateDiff("D", dtmFechaMinima, dtmFechaMaxima)
    
    vgStrAñoGlobal = Year(dtmFechaMinima)
    vgStrMesesGlobal = Month(dtmFechaMinima)
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

Private Function obtieneimporte(vlintContador As Integer)
        If CDbl(IIf(Trim(grdArticulos.TextMatrix(vlintContador, 24)) = "", "0", Trim(grdArticulos.TextMatrix(vlintContador, 24)))) <> 0 Then
            obtieneimporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 24), ""))
        Else
            obtieneimporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 2), ""))
        End If
End Function

Private Sub cmdSave_Click()
     Dim vllngConsecFacPac As Long  'Consecutivo de la factura del paciente para la impresión
     Dim vldblFacturaBaseTotal As Double          'Monto de la factura Base del POS (osea la de la empresa en el caso de que sea por aseguradora)
     Dim vldblCantidadPagoConceptos As Double     'Trae la cantidad que pagaría el paciente y que no se le facturaria
     Dim rsPvDetalleCorte As New ADODB.Recordset  'Aqui añado los registros del detalle del corte
     Dim rsFactura As New ADODB.Recordset         'RS tipo tabla para guardar la factura
     Dim rsDetalleFactura As New ADODB.Recordset  'RS tipo tabla para el Detalle de la FACTURA
     Dim rsDatosCliente As New ADODB.Recordset    'RS para los datos del cliente
     Dim rsTipoPaciente As New ADODB.Recordset    'RS para el Tipo de paciente en crédito
     Dim rsCredito  As New ADODB.Recordset
     Dim rsVentaPublico As New ADODB.Recordset    'RS para guardar la Venta, Ticket
     Dim rsDetalleVentaPublico As New ADODB.Recordset    'RS para guardar del detalle de la Venta, Ticket
     Dim vllngNumeroCorte As Long                 'Trae el numero de corte actual
     Dim a As Long                                'Contador
     Dim vlintContador As Integer                 'Para los ciclos
     Dim vlblnOkFormasPago As Boolean             'Bandera para saber si se capturaron las formas de pago de la EMPRESA o Paciente SIN Control de aseguradora
     Dim vlblnOkFormasPagoPac As Boolean     'Bandera para saber si se capturaron las formas de pago de la EMPRESA o Paciente SIN Control de aseguradora
     Dim vlblnOkFormasPagoRecPac As Boolean 'Bandera para saber si se capturaron las formas de pago del PACIENTE cuando se le da RECIBO
     Dim vllngConsecutivoVenta As Long            'Todas las ventas se van a guardar y los cargos y pagos llevarán este consecutivo
     Dim vlstrTipoPacienteCredito As String       'Sería 'PI' 'PE' 'EM' 'CO' 'ME'
     Dim vllngCveClienteCredito As Long           'Clave del empledo o del médico
     Dim vlstrFacturaTicket As String             'Trae una "F" si es Factura o una "T" si es Ticket
     Dim vldblCantidad As Double                  'Para la facturación
     Dim vllngNumCliente As Long                  'Numero de cliente para Venta a credito
     Dim vllngFoliosFaltantes As Long             'Para almacenar los folios restantes
     Dim vlstrtipopago As String                  'El tipo de pago sería 'DE', 'CO', 'CP'
     Dim vlblnSujetoIEPS As Boolean ' indica si se va a utilizar desglose de ieps en facturas que no son de tipo convenio y que sea factura de la empresa
     Dim vldblTotSubtotal As Double               'Subtotal de la cuenta, para el ticket
     Dim vldblTotIva As Double                    'IVA total de la cuenta, para el ticket
     Dim vldblTotDescuentos As Double             'Descuento de la cuenta, para el ticket
     Dim vldtmFechaHoy As Date                    'Varible con la Fecha actual
     Dim vldtmHoraHoy As Date                     'Varible con la Hora actual
     Dim vllngAux As Long
     Dim vlaryParametrosSalida() As String
     Dim vlStrConcepto As String
     Dim vlblnDatosFiscales As Boolean
     Dim intTipoEmisionComprobante As Integer
     Dim intTipoCFDFactura As Integer
     Dim rsTipoAgrupacion As New ADODB.Recordset
     Dim intTipoAgrupacion As Integer
     Dim vlaryParametrosSalidaF() As String
     Dim llngFoliosFaltantesF As Long
     Dim strFolioDocumentoF As String
     Dim strSerieDocumentoF As String
     Dim strNumeroAprobacionDocumentoF As String
     Dim strAnoAprobacionDocumentoF As String
     Dim rsPvFacturaImporte As New ADODB.Recordset
     Dim rsCalcPvFactImp As New ADODB.Recordset
     Dim vldblimportegravado As Double   ' Importe gravado
     Dim vldblSumatoria As Double  ' Sumatoria de conceptos de seguro en la factura de la empresa
     Dim vllngCorteUsado As Long
     Dim vldblComisionIvaBancaria As Double
     Dim vlstrSentencia As String
     Dim vllngnoerror As Integer
     Dim rsTemp As New ADODB.Recordset
     Dim i As Integer
     Dim vldblImporte As Double
     Dim rsReporte As New ADODB.Recordset
     
     

     On Error GoTo NotificaError
    
    If Not fblnDatosValidos() Then Exit Sub
    If Not fblnValidaCuentaPuenteIngresos(vgintClaveEmpresaContable) Then Exit Sub
    
    ' Inicializa las variables de cliente y numero de referencia '
    pInicializaVariablesGuardado
    
    If lblnExisteControl Then
        
        vldblCantidadPagoConceptos = 0
        For vlintContador = 0 To 2
            vlstrtipopago = IIf(vlintContador = 0, "DE", IIf(vlintContador = 1, "CO", "CP"))
            vldblCantidad = IIf(vlintContador = 0, ldblDeducible, IIf(vlintContador = 1, ldblCoaseguro, ldblCopago))

            If vldblCantidad > 0 Then
                vldblTotalControlAseguradorasSinIVA = vldblTotalControlAseguradorasSinIVA + Round((IIf(vlstrtipopago = "DE", ldblDeducible, IIf(vlstrtipopago = "CO", ldblCoaseguro, ldblCopago)) - (vldblFacturaPacienteIVA * (((IIf(vlstrtipopago = "DE", ldblDeducible, IIf(vlstrtipopago = "CO", ldblCoaseguro, ldblCopago)) * 100) / vldblTotalControlAseguradoras) / 100))), 2)
            End If

            If IIf(vlintContador = 0, blnFacturaDeducible, IIf(vlintContador = 1, blnFacturaCoaseguro, blnFacturaCopago)) Then
                ReDim Preserve aFPFacturaPaciente(UBound(aFPFacturaPaciente) + 1)
                aFPFacturaPaciente(UBound(aFPFacturaPaciente) - 1).dblCantidad = vldblCantidad
                aFPFacturaPaciente(UBound(aFPFacturaPaciente) - 1).strTipo = vlstrtipopago
                aFPFacturaPaciente(UBound(aFPFacturaPaciente) - 1).dblCantidadIVA = (vldblFacturaPacienteIVA * (((IIf(vlstrtipopago = "DE", ldblDeducible, IIf(vlstrtipopago = "CO", ldblCoaseguro, ldblCopago)) * 100) / vldblTotalControlAseguradoras) / 100))
                aFPFacturaPaciente(UBound(aFPFacturaPaciente) - 1).lngConceptoFacturacion = IIf(vlstrtipopago = "DE", llngCveConceptoDeducible, IIf(vlstrtipopago = "CO", llngCveConceptoCoaseguro, IIf(vlstrtipopago = "CP", llngCveConceptoCopago, 0)))
                vldblFacturaPacienteSubTotal = vldblFacturaPacienteSubTotal + vldblCantidad
            Else 'Osea que no se factura ese concepto, sino que se daría un recibo
                vldblCantidadPagoConceptos = vldblCantidadPagoConceptos + vldblCantidad
            End If
        Next vlintContador
        
        vlstrFacturaTicket = "F" 'Siempre se factura a la empresa cuando hay control de aseguradora
    Else
        If vlblnImportoVenta = True Then
            vlstrFacturaTicket = "F"
        Else
'             Se factura o se imprime Ticket únicamente cuando NO se ha capturado el Control de aseguradoras
'            ¿Desea factura?
            If MsgBox(SIHOMsg(932), vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje") = vbYes Then
                vlstrFacturaTicket = "F"
            Else
                vlblnDatosFiscales = False
                vlblnSujetoIEPS = vlblnLicenciaIEPS
                vlstrFacturaTicket = "T"
                If intFacturarVentaPublico = 1 Then blnFacturaAutomatica = True
            End If
        End If
    End If
    
    'Se validan los folios de los tickets...
    If vlstrFacturaTicket = "T" Then
        pCargaArreglo vlaryParametrosSalida, vllngFoliosRestantes & "|" & adInteger & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionPaciente & "|" & ADODB.adBSTR & "|" & strAnoAprobacionPaciente & "|" & ADODB.adBSTR
        frsEjecuta_SP "TI|" & vgintNumeroDepartamento & "|0", "Sp_GnFolios", , , vlaryParametrosSalida
        pObtieneValores vlaryParametrosSalida, vllngFoliosRestantes, strFolio, strSerie, strNumeroAprobacionPaciente, strAnoAprobacionPaciente
         'Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
        vlstrFolioDocumento = Trim(strSerie) & strFolio
        
        If Trim(vlstrFolioDocumento) = "0" Then
            MsgBox SIHOMsg(291) + " Tickets.", vbCritical + vbOKOnly, "Mensaje" 'No existen folios activos para este documento.
            Exit Sub
        End If
    End If
    
    'VALIDACION DE FORMATO/FOLIO (FISICO, DIGITAL)
    'Si es factura iniciar la validación
    If vlstrFacturaTicket = "F" Or blnFacturaAutomatica Then
        'Identifica el tipo de formato a utilizar
        lngCveFormato = 1
        frsEjecuta_SP vgintNumeroDepartamento & "|" & vgintEmpresa & "|" & vgintTipoPaciente & "|E", "fn_PVSelFormatoFactura2", True, lngCveFormato
        vllngFormatoaUsar = lngCveFormato

        'Se valida en caso de no haber formato activo mostrar mensaje y cancelar transacción
        If vllngFormatoaUsar = 0 Then
            MsgBox SIHOMsg(373), vbCritical, "Mensaje"  'No se encontró un formato válido de factura.
            Exit Sub
        Else
            pValidaFormato
        End If
        
        If vlblnImportoVenta = True Then
            'Revisa si se tiene configurado el rpt para desglosar por cargo la factura al ser factura global
            Set rsTemp = frsRegresaRs("SELECT * FROM FORMATO WHERE NOT VCHDESCRIPCIONAGRUPA3 IS NULL AND intnumeroformato = " & vllngFormatoaUsar)
            If rsTemp.EOF Then
                'No se ha configurado el archivo para el formato de factura desglosado por cargo, favor de verificar.
                MsgBox SIHOMsg(1512), vbExclamation, "Mensaje"
                Exit Sub
            End If
        End If

        'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
        '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
        intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", vllngFormatoaUsar)

        'ERROR, 'Si es error, se cancela la transacción
        If intTipoEmisionComprobante = 0 Then Exit Sub
        
        If intTipoEmisionComprobante = 2 Then
            'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
            intTipoCFDFactura = fintTipoCFD("FA", vllngFormatoaUsar)
            
             'ERROR 'Si aparece un error terminar la transacción
            If intTipoCFDFactura = 3 Then Exit Sub
        End If
    End If
        
    If Not fblnValidaCuentasIngresoDescuento Then Exit Sub
    
    'DATOS FISCALES FACTURA
    If vlngCveFormato = 2 Or vlngCveFormato = 3 Then
        If Not fblnValidaSAT Then Exit Sub
    Else
        'If vlngCveFormato = 1 And optOtrosConceptos.Value = True Then
        If vlngCveFormato = 1 Then
            If Not fblnValidaSATotrosConceptosCONCEPTOFACTURACION Then Exit Sub
        End If
    End If

    vlblnSujetoIEPS = False
    If vlstrFacturaTicket = "F" Then

       'DATOS FISCALES PACIENTE
       If vldblFacturaPacienteTotal > 0 Then
          pAbreDatosFiscales
          
          If Trim(vlstrRFCFacturaPaciente) = "" Or Trim(vlstrNombreFacturaPaciente) = "" Then Exit Sub
    
          'Pedir Datos Fiscales (de la factura BASE) Capture los datos de la factura de la empresa.
          MsgBox SIHOMsg(582), vbInformation, "Mensaje"
       End If
        
       pAbreDatosFiscalesPacienteEmpresa
        
        vlblnSujetoIEPS = IIf(frmDatosFiscales.vgBitSujetoaIEPS = 1, True, False)
        
        Unload frmDatosFiscales
        Set frmDatosFiscales = Nothing
        
       If Trim(vlstrRFC) = "" Or Trim(vlstrNombreFactura) = "" Then Exit Sub
            
    'DATOS FISCALES TICKET
    ElseIf (vlstrFacturaTicket = "T" And vlblnDatosFiscales) Or blnFacturaAutomatica Then
        'NO ESTA CONFIGURADA LA FACTURA AUTOMATICA
        If Not blnFacturaAutomatica Then
           'DATOS FISCALES PACIENTE
           If vldblFacturaPacienteTotal > 0 Then
              MsgBox SIHOMsg(581), vbInformation, "Mensaje" 'Capture los datos de la factura del paciente.
              Load frmDatosFiscales
              frmDatosFiscales.vgblnMostrarUsoCFDI = True

              If txtMovimientoPaciente.Text <> "" Then pAsignaVariablesDatosFiscales
              
              frmDatosFiscales.sstDatos.Tab = IIf(txtMovimientoPaciente.Text <> "", 0, 1)
              
              pDatosFiscales
              vlstrNumRef = frmDatosFiscales.vlstrNumRef
              vlstrTipo = frmDatosFiscales.vlstrTipo
                
              Unload frmDatosFiscales
              Set frmDatosFiscales = Nothing
                
              If Trim(vlstrRFCFacturaPaciente) = "" Or Trim(vlstrNombreFacturaPaciente) = "" Then Exit Sub
           End If
        
           ' DATOS FISCALES PARA EMPRESA, CLIENTE O PACIENTE QUE NO ES DE TIPO CONVENIO
           If vldblFacturaPacienteTotal > 0 Then MsgBox SIHOMsg(582), vbInformation, "Mensaje" 'Capture los datos de la factura de la empresa.
            
           Load frmDatosFiscales
           frmDatosFiscales.vgblnMostrarUsoCFDI = True

           If txtMovimientoPaciente.Text <> "" Then
                pAsignaVariablesDatosFiscales
           Else
                frmDatosFiscales.sstDatos.Tab = 0
                pDatosFiscalesPvParametros
           End If
           
            pDatosFiscales2
            
            Unload frmDatosFiscales
            Set frmDatosFiscales = Nothing

        ' SE TIENE CONFIGURADA LA FACTURACIÓN AUTOMATICA DE TICKETS
        Else 'Se va a facturar automáticamente
            pIniciaVarFactAutomatica
            vlblnSujetoIEPS = False
        End If

        If vlblnDatosFiscales Then If Trim(vlstrRFC) = "" Or Trim(vlstrNombreFactura) = "" Then Exit Sub
    End If

    ' Las formas de pago para la FACTURA del PACIENTE CON CONTROL de aseguradora.
    vlblnOkFormasPagoPac = True
    If vldblFacturaPacienteTotal > 0.1 Then 'Le puse 0.1 por aquello de los cálculos no exactos
        MsgBox SIHOMsg(579), vbInformation, "Mensaje" 'Por favor seleccione la forma de pago para la factura del paciente.
        vlblnOkFormasPagoPac = fblnFormasPagoPos(aFormasPagoPaciente(), vldblFacturaPacienteTotal, True, vldblTipoCambio, True, CLng(Val(txtMovimientoPaciente.Text)), "PE", Trim(Replace(Replace(Replace(IIf(Trim(vlstrRFC) <> "", Trim(vlstrRFC), Trim(vgstrRFCPersonaSeleccionada)), "-", ""), "_", ""), " ", "")), , , , "frmPOS")
        If Not vlblnOkFormasPagoPac Then Exit Sub
    End If

    ' Las formas de pago para el RECIBO DE PAGO del PACIENTE CON CONTROL de aseguradora.
    vlblnOkFormasPagoRecPac = True
    If vldblCantidadPagoConceptos > 0 Then
        MsgBox SIHOMsg(591), vbInformation, "Mensaje" 'Por favor seleccione la forma de pago para el RECIBO del paciente.
        vlblnOkFormasPagoRecPac = fblnFormasPagoPos(aFormasPagoReciboPaciente(), vldblCantidadPagoConceptos, True, vldblTipoCambio, False, 0, "PE", Trim(Replace(Replace(Replace(IIf(Trim(vlstrRFC) <> "", Trim(vlstrRFC), IIf(Trim(vlstrRFCFacturaPaciente) <> "", Trim(vlstrRFCFacturaPaciente), Trim(vgstrRFCPersonaSeleccionada))), "-", ""), "_", ""), " ", "")), , , , "frmPOS")
        If Not vlblnOkFormasPagoPac Then Exit Sub
    End If
    
    vldblFacturaBaseTotal = Val(Format(txtTotal.Text, "")) - vldblFacturaPacienteTotal
    
    ' Las formas de pago para la empresa o paciente sin control de aseguradora. (Factura BASE)
    If vldblFacturaBaseTotal >= 0 Then
        vlblnOkFormasPago = False
        vlstrTipoPacienteCredito = ""
        vllngCveClienteCredito = 0
            
        If optPaciente.Value Then
           'Pacientes
            Set rsTipoPaciente = frsRegresaRs("select chrTipo from AdTipoPaciente where tnyCveTipoPaciente = " & Trim(str(vgintTipoPaciente)), adLockReadOnly, adOpenForwardOnly)
            vlstrTipoPacienteCredito = rsTipoPaciente!chrTipo
            
            If vlstrTipoPacienteCredito = "PA" Then
                vlstrTipoPacienteCredito = "PE"
                vllngCveClienteCredito = Val(txtMovimientoPaciente.Text)
            ElseIf vlstrTipoPacienteCredito = "CO" Then
            'Si es paciente de convenio
                vllngCveClienteCredito = vgintEmpresa
            ElseIf vlstrTipoPacienteCredito <> "" Then
            'Si está relacionado con un empleado o un médico
                vllngCveClienteCredito = vglngCveExtra
            Else
                vllngCveClienteCredito = 0
            End If
            rsTipoPaciente.Close
            
        ElseIf optMedico.Value Then
        'Médicos
            vlstrTipoPacienteCredito = "ME"
            vllngCveClienteCredito = vglngCveExtra
            
        ElseIf optEmpleado.Value Then
        'Empleados
            vlstrTipoPacienteCredito = "EM"
            vllngCveClienteCredito = vglngCveExtra
        End If
        
        If vldblFacturaPacienteTotal > 0 Or vldblCantidadPagoConceptos > 0 Then MsgBox SIHOMsg(580), vbInformation, "Mensaje" 'Por favor seleccione la forma de pago para la empresa.
        
        Set rsCredito = frsRegresaRs("select count(*) from CcCliente Inner join nodepartamento on cccliente.smicvedepartamento = nodepartamento.smicvedepartamento Where cccliente.intNumReferencia = " + str(vllngCveClienteCredito) + " And cccliente.chrTipoCliente = " + " '" + vlstrTipoPacienteCredito + "'" + " and cccliente.bitActivo=1 and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable)
        If rsCredito.Fields(0) = 0 And vlstrTipoPacienteCredito = "EM" Then
            vllngCveClienteCredito = txtMovimientoPaciente.Text
            vlstrTipoPacienteCredito = "PE"
        End If
        
        If vldblFacturaBaseTotal > 0 Then vlblnOkFormasPago = fblnFormasPagoPos(aFormasPago(), vldblFacturaBaseTotal - vldblCantidadPagoConceptos, True, vldblTipoCambio, rsCredito.RecordCount <> 0, vllngCveClienteCredito, vlstrTipoPacienteCredito, Trim(Replace(Replace(Replace(IIf(Trim(vlstrRFC) <> "", Trim(vlstrRFC), IIf(Trim(vlstrRFCFacturaPaciente) <> "", Trim(vlstrRFCFacturaPaciente), Trim(vgstrRFCPersonaSeleccionada))), "-", ""), "_", ""), " ", "")), , , True, "frmPOS")
    Else
        '¡No se puede facturar si la cantidad es cero o menor que cero!
        MsgBox SIHOMsg(429), vbExclamation, "Mensaje"
    End If
    
    
    ' Formato de factura a utilizar
    If vlstrFacturaTicket = "F" Then
        llngFormatoFacturaAUsar = 1
        frsEjecuta_SP IIf(optPaciente.Value And Trim(txtMovimientoPaciente.Text) <> "", vgintNumeroDepartamento & "|" & vgintEmpresa & "|" & vgintTipoPaciente & "|E", vgintNumeroDepartamento & "|0|0|E"), "Fn_Pvselformatofactura", False, llngFormatoFacturaAUsar
        If llngFormatoFacturaAUsar = 0 Then
            MsgBox SIHOMsg(373), vbCritical, "Mensaje" 'No se encontró un formato válido de factura, por favor de uno de alta.
            Exit Sub
        End If
    End If
    
    If vldblFacturaBaseTotal > 0 Then
       If Not (vlblnOkFormasPago And vlblnOkFormasPagoRecPac And vlblnOkFormasPagoPac And Val(Format(txtTotal.Text, "")) > 0) Then Exit Sub
    End If

    
    ' Persona que graba
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    ' En caso de que el corte se maneje por empleado valida que sea el mismo de persona graba si no pone 0 en vllngPersonaGraba
    If vlblnParametroTipoCorte Then
        If rsGnParametrosTipoCorte!intTipoCorte = 2 Then
            If vllngPersonaGraba <> vglngNumeroEmpleado Then vllngPersonaGraba = 0
        End If
    End If
       
    If vllngPersonaGraba = 0 Then
       'El corte no pertenece a la persona que está registrando.
       MsgBox SIHOMsg(933), vbOKOnly + vbExclamation, "Mensaje"
       Exit Sub
    End If
    
    ' Fecha y hora del Sistema
    vldtmFechaHoy = fdtmServerFecha
    vldtmHoraHoy = fdtmServerHora

    
    'Corte
    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    pAgregarMovArreglo
    
    'Inicio de Transacción
    EntornoSIHO.ConeccionSIHO.BeginTrans
        
    
    'FACTURA DEL PACIENTE
    If vldblFacturaPacienteTotal > 0 And vlstrFacturaTicket = "F" Then
       'folio factura del paciente
       pCargaArreglo vlaryParametrosSalida, vllngFoliosFaltantes & "|" & adInteger & "|" & strFolioDocumento & "|" & ADODB.adBSTR & "|" & strSerieDocumento & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionDocumento & "|" & ADODB.adBSTR & "|" & strAnoAprobacionDocumento & "|" & ADODB.adBSTR
       frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "Sp_GnFolios", , , vlaryParametrosSalida
       pObtieneValores vlaryParametrosSalida, vllngFoliosFaltantes, strFolioDocumento, strSerieDocumento, strNumeroAprobacionDocumento, strAnoAprobacionDocumento
       strSerieDocumento = Trim(strSerieDocumento)
       vlstrFolioDocumento = strSerieDocumento & strFolioDocumento
       If Trim(vlstrFolioDocumento) = "0" Then
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          'No existen folios activos para este documento.
          MsgBox SIHOMsg(291), vbCritical, "Mensaje"
          Exit Sub
       End If
                  
       ' Formas de pago para el corte de la factura del paciente
       ' Afecta corte (PvDetalleCorte) (De la factura del paciente)
       Set rsPvDetalleCorte = frsRegresaRs("SELECT * FROM PvDetalleCorte WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
       For a = 0 To UBound(aFormasPagoPaciente(), 1)
           pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), _
                                   IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), 0, _
                                   IIf((aFormasPago(a).vldblTipoCambio = 0), aFormasPagoPaciente(a).vldblCantidad, aFormasPagoPaciente(a).vldblDolares), _
                                   False, CStr(vldtmFechaHoy + vldtmHoraHoy), CLng(aFormasPagoPaciente(a).vlintNumFormaPago), aFormasPagoPaciente(a).vldblTipoCambio, _
                                   IIf(Trim(aFormasPagoPaciente(a).vlstrFolio) = "", "0", Trim(aFormasPagoPaciente(a).vlstrFolio)), _
                                   vllngNumeroCorte, 1, "FACTURAPACIENTE", IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                                   aFormasPagoPaciente(a).vlbolEsCredito, aFormasPagoPaciente(a).vlstrRFC, aFormasPagoPaciente(a).vlstrBancoSAT, aFormasPagoPaciente(a).vlstrBancoExtranjero, aFormasPagoPaciente(a).vlstrCuentaBancaria, aFormasPagoPaciente(a).vldtmFecha
                                
           'Cargo a la cuenta de la forma de pago
           pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), IIf(vlstrFacturaTicket = "F", "FA", "TI"), aFormasPagoPaciente(a).vllngCuentaContable, IIf(aFormasPagoPaciente(a).vldblTipoCambio = 0, aFormasPagoPaciente(a).vldblCantidad, aFormasPagoPaciente(a).vldblDolares * aFormasPagoPaciente(a).vldblTipoCambio), True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA"
                         
            ' Agregado para caso 8741
            ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
            If aFormasPagoPaciente(a).vllngCuentaComisionBancaria <> 0 And aFormasPagoPaciente(a).vldblCantidadComisionBancaria <> 0 Then
                ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), IIf(vlstrFacturaTicket = "F", "FA", "TI"), aFormasPagoPaciente(a).vllngCuentaComisionBancaria, aFormasPagoPaciente(a).vldblCantidadComisionBancaria, True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
                If aFormasPagoPaciente(a).vldblIvaComisionBancaria <> 0 Then
                    ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                    pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), IIf(vlstrFacturaTicket = "F", "FA", "TI"), glngCtaIVAPagado, aFormasPagoPaciente(a).vldblIvaComisionBancaria, True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
                End If
                ' Se genera un abono por la cantidad de la comisión bancaria y su iva que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), IIf(vlstrFacturaTicket = "F", "FA", "TI"), aFormasPagoPaciente(a).vllngCuentaContable, (aFormasPagoPaciente(a).vldblCantidadComisionBancaria + aFormasPagoPaciente(a).vldblIvaComisionBancaria), False, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
             End If
                                   
           ' Generar movimiento de CREDITO. '
           If aFormasPagoPaciente(a).vlbolEsCredito Then
              
              ' Para ver que numero de cliente es este paciente '
              Set rsDatosCliente = frsRegresaRs("SELECT * FROM CcCliente INNER JOIN NoDepartamento ON CcCliente.smiCveDepartamento = NoDepartamento.smiCveDepartamento WHERE CcCliente.intNumReferencia = " & Trim(txtMovimientoPaciente.Text) & " AND CcCliente.chrTipoCliente = 'PE' AND NoDepartamento.tnyClaveEmpresa = " & vgintClaveEmpresaContable, adLockReadOnly, adOpenForwardOnly)
              If rsDatosCliente.RecordCount = 0 Then
                 EntornoSIHO.ConeccionSIHO.RollbackTrans
                 rsDatosCliente.Close
                 'Se detectó un error en la información del cliente.
                 MsgBox SIHOMsg(367), vbCritical, "Mensaje"
                 Exit Sub
              Else
                 vllngNumCliente = rsDatosCliente!intNumCliente
              End If
              
              pCrearMovtoCredito aFormasPago(a).vldblCantidad, vllngNumCliente, rsDatosCliente!INTNUMCUENTACONTABLE, vlstrFolioDocumento, IIf(vlstrFacturaTicket = "F", "FA", "TI"), vldblTipoCambio, vllngPersonaGraba
              rsDatosCliente.Close
           End If
       Next a
    
       ' GRABAR FACTURA DEL PACIENTE
       If optPesos(0).Value Then vldblTipoCambio = 1
                   
       Set rsFactura = frsRegresaRs("SELECT * FROM PVFactura WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
       With rsFactura
            .AddNew
            !chrfoliofactura = vlstrFolioDocumento
            vlstrFolioDocumentoPaciente = vlstrFolioDocumento
            !dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
            !CHRRFC = IIf(vlBitExtranjero, "XEXX010101000", IIf(Len(Trim(Replace(Replace(Replace(vlstrRFCFacturaPaciente, "-", ""), "_", ""), " ", ""))) < 12 Or Len(Trim(Replace(Replace(Replace(vlstrRFCFacturaPaciente, "-", ""), "_", ""), " ", ""))) > 13, "XAXX010101000", Trim(Replace(Replace(Replace(vlstrRFCFacturaPaciente, "-", ""), "_", ""), " ", ""))))
            !vchSerie = strSerieDocumento
            !INTFOLIO = IIf(Trim(strFolioDocumento) = "", Null, strFolioDocumento)
            !CHRNOMBRE = vlstrNombreFacturaPaciente
            !chrCalle = vlstrDireccionFacturaPaciente
            !VCHNUMEROEXTERIOR = vlstrNumeroExteriorFacturaPaciente
            !VCHNUMEROINTERIOR = vlstrNumeroInteriorFacturaPaciente
            !VCHCOLONIA = vlstrColoniaFacturaPaciente
            !VCHCODIGOPOSTAL = vlstrCPFacturaPaciente
            !chrTelefono = vlstrTelefonoFacturaPaciente
            !smyIVA = Round((vldblFacturaPacienteIVA / vldblTipoCambio), 2)
            !MNYDESCUENTO = 0
            !chrEstatus = " "
            !INTMOVPACIENTE = Val(txtMovimientoPaciente.Text)
            !CHRTIPOPACIENTE = "E"
            !SMIDEPARTAMENTO = vgintNumeroDepartamento
            !intCveEmpleado = vllngPersonaGraba
            !intNumCorte = vllngNumeroCorte
            !mnyAnticipo = 0
            !mnyTotalFactura = Round((vldblFacturaPacienteTotal / vldblTipoCambio), 2)
            !BITPESOS = IIf(optPesos(0).Value, 1, 0)
            !mnytipocambio = IIf(optPesos(0).Value, 0, vldblTipoCambio)
            !chrTipoFactura = "P" 'Es el Tipo de Factura si es "P" es Paciente, si es "E" es empresa
            !intCveVentaPublico = vllngConsecutivoVenta
            !VCHREGIMENFISCALRECEPTOR = vgstrRegimenFiscal
            !intNumCliente = vllngNumCliente
            !intCveCiudad = vllngCveCiudadFacturaPaciente 'llngCveCiudadPaciente
            !intcveempresa = 0
            !bitdesgloseIEPS = IIf(vlblnImportoVenta, 1, IIf(vlblnSujetoIEPS, 1, 0))
            !intCveUsoCFDI = IIf(vlintUsoCFDI = 0, 64, vlintUsoCFDI)
            !bitFacturaGlobal = IIf(vlblnImportoVenta = True, 1, 0)
            .Update
            
            vldblsubtotalNogravado = vldblsubtotalNogravado + (!mnyTotalFactura - !smyIVA)
            vldblDescuentoNoGravado = vldblDescuentoNoGravado + Round(!MNYDESCUENTO, 2)
       End With
           
       vllngConsecFacPac = flngObtieneIdentity("SEC_PvFactura", rsFactura!intConsecutivo)
                        
       'Guardar información de la forma de pago en tabla intermedia para factura de paciente
       If vlblnOkFormasPagoPac Then
          For a = 0 To UBound(aFormasPagoPaciente(), 1)
              If Not aFormasPagoPaciente(a).vlbolEsCredito Then 'Formas de pago distintas a Crédito
                    frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPagoPaciente(a).vlintNumFormaPago & "|" & aFormasPagoPaciente(a).lngIdBanco & "|" & IIf(aFormasPagoPaciente(a).vldblTipoCambio = 0, aFormasPagoPaciente(a).vldblCantidad, aFormasPagoPaciente(a).vldblDolares) & "|" & IIf(aFormasPagoPaciente(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPagoPaciente(a).vldblTipoCambio & "|" & fstrTipoMovimientoForma(aFormasPagoPaciente(a).vlintNumFormaPago, "F") & "|" & "FA" & "|" & vllngConsecFacPac & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                    If aFormasPagoPaciente(a).vllngCuentaComisionBancaria <> 0 And aFormasPagoPaciente(a).vldblCantidadComisionBancaria <> 0 Then
                        frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPagoPaciente(a).vlintNumFormaPago & "|" & aFormasPagoPaciente(a).lngIdBanco & "|" & (aFormasPagoPaciente(a).vldblCantidadComisionBancaria + aFormasPagoPaciente(a).vldblIvaComisionBancaria) * -1 & "|" & IIf(aFormasPagoPaciente(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPagoPaciente(a).vldblTipoCambio & "|" & "CBA" & "|" & "FA" & "|" & vllngConsecFacPac & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    End If
              End If
          Next a
       End If
       
       ' Detalle de la Factura (del paciente) '
       ' Aqui trabajo con un grid escondido(grdFactura) organizar los cargos, como en la facturación normal.
       ' Por si no tiene cargos
       If grdArticulos.RowData(1) = -1 Then Exit Sub
            
       pPreparaGridParaDatosFactura
            
       If grdFactura.RowData(1) <> -1 Then grdFactura.Rows = grdFactura.Rows + 1
       
       grdFactura.TextMatrix(grdFactura.Rows - 1, 1) = ""  'Descripción del Concepto
       grdFactura.TextMatrix(grdFactura.Rows - 1, 5) = 0 'Descuentos individuales
       
       ' Para el DEDUCIBLE (Un registro de detalle para el deducible) '
       If ldblDeducible > 0 And blnFacturaDeducible Then
            grdFactura.RowData(grdFactura.Rows - 1) = llngCveConceptoDeducible 'Clave del concepto
            grdFactura.TextMatrix(grdFactura.Rows - 1, 2) = IIf(blnFacturaDeducible, Round((ldblDeducible - (vldblFacturaPacienteIVA * (((ldblDeducible * 100) / vldblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), ldblDeducible)  'Cantidad
            grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = IIf(blnFacturaDeducible, Round(vldblFacturaPacienteIVA * (((ldblDeducible * 100) / vldblTotalControlAseguradoras) / 100) / vldblTipoCambio, 4), 0)  'IVA
       End If
       
       ' Para el COASEGURO (Un registro de detalle para el COASEGURO) '
       If ldblCoaseguro > 0 And blnFacturaCoaseguro Then
            grdFactura.Rows = grdFactura.Rows + 1
            grdFactura.RowData(grdFactura.Rows - 1) = llngCveConceptoCoaseguro 'Clave del concepto
            grdFactura.TextMatrix(grdFactura.Rows - 1, 2) = IIf(blnFacturaCoaseguro, Round((ldblCoaseguro - (vldblFacturaPacienteIVA * (((ldblCoaseguro * 100) / vldblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), ldblCoaseguro)     'Cantidad
            grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = IIf(blnFacturaCoaseguro, Round(vldblFacturaPacienteIVA * (((ldblCoaseguro * 100) / vldblTotalControlAseguradoras) / 100) / vldblTipoCambio, 4), 0)  'IVA
       End If
       
       ' Para el COPAGO (Un registro de detalle para el COPAGO) '
       If ldblCopago > 0 And blnFacturaCopago Then
            grdFactura.Rows = grdFactura.Rows + 1
            grdFactura.RowData(grdFactura.Rows - 1) = llngCveConceptoCopago 'Clave del concepto
            grdFactura.TextMatrix(grdFactura.Rows - 1, 2) = IIf(blnFacturaCopago, Round((ldblCopago - (vldblFacturaPacienteIVA * (((ldblCopago * 100) / vldblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), ldblCopago)     'Cantidad
            grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = IIf(blnFacturaCopago, Round(vldblFacturaPacienteIVA * (((ldblCopago * 100) / vldblTotalControlAseguradoras) / 100) / vldblTipoCambio, 4), 0) 'IVA
       End If
                            
       Set rsDetalleFactura = frsRegresaRs("SELECT * FROM PVDetalleFactura WHERE chrFolioFactura = ''", adLockOptimistic, adOpenDynamic)
       With rsDetalleFactura
           For a = 1 To grdFactura.Rows - 1
               If grdFactura.RowData(a) > 0 Then 'Porque pueden ser negativos con el descuento y Pagos
                   .AddNew
                   !chrfoliofactura = vlstrFolioDocumento
                   !smicveconcepto = grdFactura.RowData(a)
                   !MNYCantidad = Val(grdFactura.TextMatrix(a, 2))
                   !MNYIVA = Val(grdFactura.TextMatrix(a, 4))
                   !MNYDESCUENTO = Val(grdFactura.TextMatrix(a, 5))
                   !mnyIVAConcepto = Val(grdFactura.TextMatrix(a, 4))
                   !chrTipo = "OC" 'Concepto "Normal"
                   '!VCHFOLIOTICKETFACTGLOBAL = IIf(vlblnImportoVenta = True, grdFactura.TextMatrix(a, 8), "")
                   
                   .Update
        
                   vldblsubtotalgravado = vldblsubtotalgravado + Round(!MNYCantidad - !MNYDESCUENTO, 2)
                   vldblsubtotalNogravado = vldblsubtotalNogravado - Round(!MNYCantidad - !MNYDESCUENTO, 2)
                   vldbldescuentogravado = vldbldescuentogravado + !MNYDESCUENTO
                   vldblDescuentoNoGravado = vldblDescuentoNoGravado - Round(!MNYDESCUENTO, 2)
           
                   
                   ' Registros de la póliza del Deducible, Coaseguro y Copago
                   'Abono a la cuenta de ingreso
                   pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", flngCuentaConceptoDepartamento(grdFactura.RowData(a), vgintNumeroDepartamento, "INGRESO"), Val(grdFactura.TextMatrix(a, 2)), False, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA"
                   'Abono a la cuenta de IVA por pagar
                   If Val(grdFactura.TextMatrix(a, 4)) <> 0 Then
                       pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", flngCuentaConceptoDepartamento(grdFactura.RowData(a), vgintNumeroDepartamento, "IVA"), Val(grdFactura.TextMatrix(a, 4)), False, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA"
                   End If
               End If
           Next a
       End With
            
       frsEjecuta_SP str(vllngConsecFacPac) & "|" & str(Round(vldblsubtotalgravado, 2)) & "|" & str(Round(vldblsubtotalNogravado, 2)) & "|" & str(Round(vldbldescuentogravado, 2)) & "|" & str(Round(vldblDescuentoNoGravado, 2)), "SP_PVINSFACTURAIMPORTE"
       rsDetalleFactura.Close
       rsFactura.Close
    End If
        
    
    ' Número de la factura BASE/EMPRESA/PACIENTE(cuando solamente de factura al paciente) o el numero del recibo de pago '
    vllngFoliosFaltantes = 0 'Para que validar los folios faltantes (aqui no aplica)
    vlstrFolioDocumento = "0"
    
    ' Folio de factura BASE, osea la de la empresa en caso de la aseguradora o tambien del paciente si es que no tiene aseguradora
    If vlstrFacturaTicket = "F" Then
       pCargaArreglo vlaryParametrosSalida, vllngFoliosFaltantes & "|" & adInteger & "|" & strFolioDocumento & "|" & ADODB.adBSTR & "|" & strSerieDocumento & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionDocumento & "|" & ADODB.adBSTR & "|" & strAnoAprobacionDocumento & "|" & ADODB.adBSTR
       frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "Sp_GnFolios", , , vlaryParametrosSalida
       pObtieneValores vlaryParametrosSalida, vllngFoliosFaltantes, strFolioDocumento, strSerieDocumento, strNumeroAprobacionDocumento, strAnoAprobacionDocumento
        
      '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
       strSerieDocumento = Trim(strSerieDocumento)
       vlstrFolioDocumento = strSerieDocumento & strFolioDocumento
       
       If Trim(vlstrFolioDocumento) = "0" Then
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          MsgBox SIHOMsg(291), vbCritical, "Mensaje" 'No existen folios activos para este documento.
          Exit Sub
       End If
    Else
       pCargaArreglo vlaryParametrosSalida, vllngFoliosFaltantes & "|" & adInteger & "|" & strFolioDocumento & "|" & ADODB.adBSTR & "|" & strSerieDocumento & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionDocumento & "|" & ADODB.adBSTR & "|" & strAnoAprobacionDocumento & "|" & ADODB.adBSTR
       frsEjecuta_SP "TI|" & vgintNumeroDepartamento & "|1", "Sp_GnFolios", , , vlaryParametrosSalida
       pObtieneValores vlaryParametrosSalida, vllngFoliosFaltantes, strFolioDocumento, strSerieDocumento, strNumeroAprobacionDocumento, strAnoAprobacionDocumento
        
       '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
       strSerieDocumento = Trim(strSerieDocumento)
       vlstrFolioDocumento = strSerieDocumento & strFolioDocumento
       
       If Trim(vlstrFolioDocumento) = "0" Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            MsgBox SIHOMsg(291), vbCritical, "Mensaje" 'No existen folios activos para este documento.
            Exit Sub
       End If
       
       If blnFacturaAutomatica Then
           llngFoliosFaltantesF = 0
           pCargaArreglo vlaryParametrosSalidaF, llngFoliosFaltantesF & "|" & adInteger & "|" & strFolioDocumentoF & "|" & ADODB.adBSTR & "|" & strSerieDocumentoF & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionDocumentoF & "|" & ADODB.adBSTR & "|" & strAnoAprobacionDocumentoF & "|" & ADODB.adBSTR
           frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "sp_gnFolios", , , vlaryParametrosSalidaF
           pObtieneValores vlaryParametrosSalidaF, llngFoliosFaltantesF, strFolioDocumentoF, strSerieDocumentoF, strNumeroAprobacionDocumentoF, strAnoAprobacionDocumentoF
           '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
           strSerieDocumentoF = Trim(strSerieDocumentoF)
           lstrFolioDocumentoF = strSerieDocumentoF & strFolioDocumentoF
            
           If Trim(strFolioDocumentoF) = "0" Then
              EntornoSIHO.ConeccionSIHO.RollbackTrans
              'No existen folios activos para este documento.
              MsgBox SIHOMsg(291), vbCritical, "Mensaje"
              Exit Sub
           End If
       End If
    End If

    
    ' Generar Registro de la Venta
    Set rsVentaPublico = frsRegresaRs("SELECT * FROM PvVentaPublico WHERE intCveVenta = -1", adLockOptimistic, adOpenDynamic)
    With rsVentaPublico
        .AddNew
        !dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
        !intCveEmpleado = vllngPersonaGraba
        !chrFolioTicket = Trim(vlstrFolioDocumento)
        !intCveDepartamento = vgintNumeroDepartamento
        !chrCliente = Trim(lblPaciente.Caption)
        !chrTipoRecivo = vlstrFacturaTicket
        !intNumCorte = vllngNumeroCorte
        
        If vlstrFacturaTicket = "F" Then
            !chrfoliofactura = vlstrFolioDocumento
        Else
            If blnFacturaAutomatica Then !chrfoliofactura = lstrFolioDocumentoF
        End If
        
        For vlintContador = 1 To grdArticulos.Rows - 1
            vldblTotSubtotal = vldblTotSubtotal + IIf(vlblnImportoVenta = False, Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), ""))), 2), Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), ""))))
            vldblTotIva = vldblTotIva + IIf(CDbl(IIf(grdArticulos.TextMatrix(vlintContador, 26) = "", "0", grdArticulos.TextMatrix(vlintContador, 26))) <> 0, Val(Format(grdArticulos.TextMatrix(vlintContador, 26), "")), Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")))
            vldblTotDescuentos = vldblTotDescuentos + IIf(CDbl(IIf(grdArticulos.TextMatrix(vlintContador, 25) = "", "0", grdArticulos.TextMatrix(vlintContador, 25))) <> 0, Val(Format(grdArticulos.TextMatrix(vlintContador, 25), "")), Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")))
            vldblTotIEPS = vldblTotIEPS + Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
        Next
        
        !MNYSUBTOTAL = Round(vldblTotSubtotal, 2)
        !MNYIVA = vldblTotIva
        !MNYDESCUENTO = vldblTotDescuentos
        !bitcancelado = IIf(blnFacturaAutomatica, 1, 0)
        !INTMOVPACIENTE = Val(txtMovimientoPaciente)
        !intCveMedico = IIf(cboMedico.ListIndex = -1, 0, cboMedico.ItemData(cboMedico.ListIndex))
        !bitFacturaAutomatica = IIf(blnFacturaAutomatica, 1, 0)
        !mnyIeps = vldblTotIEPS
        .Update
        vllngConsecutivoVenta = flngObtieneIdentity("SEC_PVVENTAPUBLICO", !INTCVEVENTA)
    End With
    
    
    ' Generar detalle de la venta
    vldblIngresosPuente = 0
    Set rsDetalleVentaPublico = frsRegresaRs("SELECT * FROM PvDetalleVentaPublico WHERE intCveVenta = -1", adLockOptimistic, adOpenDynamic)
    For vlintContador = 1 To grdArticulos.Rows - 1
        
        ' Generar un cargo por cada registro de DetalleVentaPublico
        Dim rsVentaAlmacen As New ADODB.Recordset
        Dim vlintAlmacenVenta As Integer
        Set rsVentaAlmacen = frsRegresaRs("SELECT smiCveDepartamento FROM NoDepartamento WHERE chrClasificacion = 'A' AND smiCveDepartamento = " & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
        
        If Not rsVentaAlmacen.RecordCount = 0 Then
            vlintAlmacenVenta = rsVentaAlmacen!smicvedepartamento
        Else
            Set rsVentaAlmacen = frsRegresaRs("SELECT intNumAlmacen FROM PvAlmacenes WHERE intnumdepartamento =" & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
            vlintAlmacenVenta = rsVentaAlmacen!intnumalmacen
        End If
        
        With grdArticulos
            If vlblnImportoVenta = True Then
'                'Se importaron los conceptos de la venta así que realizan los cargos directamente

                    'Graba el cargo
                    Set rsTemp = frsRegresaRs("SELECT * FROM PVCARGO WHERE INTNUMCARGO = -1", adLockOptimistic, adOpenDynamic)
                    With rsTemp
                        .AddNew
                        !chrTipoDocumento = "T"
                        !intFolioDocumento = vllngConsecutivoVenta
                        !INTMOVPACIENTE = IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, CLng(Val(txtMovimientoPaciente.Text)))
                        !CHRTIPOPACIENTE = IIf(txtMovimientoPaciente.Text = "", "T", "E")
                        !chrfoliofactura = Trim(vlstrFolioDocumento)
                        !CHRCVECARGO = grdArticulos.RowData(vlintContador)
                        !smicveconcepto = Trim(grdArticulos.TextMatrix(vlintContador, 13))
                        !chrTipoCargo = Trim(grdArticulos.TextMatrix(vlintContador, 12))
                        !MNYCantidad = 1
                        !mnyPrecio = IIf(CDbl(grdArticulos.TextMatrix(vlintContador, 24)) <> 0, CDbl(grdArticulos.TextMatrix(vlintContador, 24)), CDbl(grdArticulos.TextMatrix(vlintContador, 2)))
                        !MNYDESCUENTO = IIf(CDbl(grdArticulos.TextMatrix(vlintContador, 25)) <> 0, CDbl(grdArticulos.TextMatrix(vlintContador, 25)), CDbl(grdArticulos.TextMatrix(vlintContador, 5)))
                        !MNYIVA = IIf(CDbl(grdArticulos.TextMatrix(vlintContador, 26)) <> 0, CDbl(grdArticulos.TextMatrix(vlintContador, 26)), CDbl(grdArticulos.TextMatrix(vlintContador, 8)))
                        !dtmFechahora = fdtmServerFecha
                        !intEmpleado = vllngPersonaGraba
                        !SMIDEPARTAMENTO = vlintAlmacenVenta
                        !bitExcluido = 0
                        !bitHojaConsumo = 0
                        !intNumCirugia = 0
                        !intDescuentaInventario = 0
                        !INTNUMKARDEX = 0
                        !mnyIncrementoHorario = 0
                        !bitPrecioManual = Val((grdArticulos.TextMatrix(vlintContador, 23)))
                        !mnyIeps = CDbl(grdArticulos.TextMatrix(vlintContador, 6))
                        !VCHFOLIOTICKETFACTGLOBAL = IIf(vlblnImportoVenta = True, Trim(grdArticulos.TextMatrix(vlintContador, 22)), "")
                        .Update
                    End With
                    
                    .TextMatrix(vlintContador, 15) = flngObtieneIdentity("SEC_PVCARGO", rsTemp!IntNumCargo)
                    
                    pEjecutaSentencia "DELETE FROM PvDescuento WHERE chrTipoDescuento = 'V' AND intCveAfectada = " & IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, txtMovimientoPaciente.Text) & " and tnyclaveempresa = " & vgintClaveEmpresaContable
                    If Val(.TextMatrix(vlintContador, 15)) < 0 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(Val(.TextMatrix(vlintContador, 15)) * -1) & "(" & Trim(.TextMatrix(vlintContador, 1)) & ")", vbExclamation, "Mensaje"
                        Exit Sub
                    End If
            Else
                ' Cuando la cantidad en IVArticulo es 1 then Se carga como UNIDAD ALTERNA (es por eso de la validacion del parametro de "DescuentaInventario")
                vllngAux = 1
                frsEjecuta_SP "V|" & IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, txtMovimientoPaciente.Text) & "|" & IIf(txtMovimientoPaciente.Text = "", "T", "E") & "|0|" & .RowData(vlintContador) & "|AR|" & Val(Format(.TextMatrix(vlintContador, 5), "")) & "|0|" & vgintClaveEmpresaContable, "fn_PvInsDescuento", True, vllngAux
                
                If vllngAux = -99 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    'No se puede realizar la operación, inténtelo en unos minutos.
                    MsgBox SIHOMsg(720), vbExclamation + vbOKOnly, "Mensaje"
                    Exit Sub
                End If
                vllngAux = 1
                
                If Val(.TextMatrix(vlintContador, 14)) <> 0 Then
                    vllngnoerror = flngBloqueaArticulo2(fstrObtenCveArticulo(CLng(.TextMatrix(vlintContador, 10))), vgintNumeroDepartamento, cgstrModulo, Me.Name)
                    frsEjecuta_SP .RowData(vlintContador) & "|" & vlintAlmacenVenta & "|" & "T" & "|" & vllngConsecutivoVenta & "|" & IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, CLng(Val(txtMovimientoPaciente.Text))) & "|" & IIf(txtMovimientoPaciente.Text = "", "T", "E") & "|" & .TextMatrix(vlintContador, 12) & "|" & 0 & "|" & CLng(Val(.TextMatrix(vlintContador, 3))) & "|" & vllngPersonaGraba & "|" & Val(.TextMatrix(vlintContador, 14)) & "|" & "SVP" & "|" & vllngConsecutivoVenta & "|" & 2, "Sp_PvUpdCargos", True, vllngAux
                    vllngnoerror = flngLiberaArticulo(fstrObtenCveArticulo(CLng(.TextMatrix(vlintContador, 10))), vgintNumeroDepartamento)
                Else
                    frsEjecuta_SP .RowData(vlintContador) & "|" & vlintAlmacenVenta & "|" & "T" & "|" & vllngConsecutivoVenta & "|" & IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, CLng(Val(txtMovimientoPaciente.Text))) & "|" & IIf(txtMovimientoPaciente.Text = "", "T", "E") & "|" & .TextMatrix(vlintContador, 12) & "|" & 0 & "|" & CLng(Val(.TextMatrix(vlintContador, 3))) & "|" & vllngPersonaGraba & "|" & Val(.TextMatrix(vlintContador, 14)) & "|" & "SVP" & "|" & vllngConsecutivoVenta & "|" & 2, "Sp_PvUpdCargos", True, vllngAux
                End If
                
                .TextMatrix(vlintContador, 15) = CStr(vllngAux)
                pActualizaPrecio vlintContador
                
                'Existencias de trazabilidad
               tValidarExistencias vgblnTrazabilidad, CLng(vlintContador)
                
                                        
                pEjecutaSentencia "DELETE FROM PvDescuento WHERE chrTipoDescuento = 'V' AND intCveAfectada = " & IIf(txtMovimientoPaciente.Text = "", vllngConsecutivoVenta, txtMovimientoPaciente.Text) & " and tnyclaveempresa = " & vgintClaveEmpresaContable
                If Val(.TextMatrix(vlintContador, 15)) < 0 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(Val(.TextMatrix(vlintContador, 15)) * -1) & "(" & Trim(.TextMatrix(vlintContador, 1)) & ")", vbExclamation, "Mensaje"
                    Exit Sub
                End If
                
                'si tiene IEPS se agrega al cargo
                If Val(Format(.TextMatrix(vlintContador, 6), "")) > 0 Then
                   pEjecutaSentencia "Update Pvcargo set mnyIEPS= " & Val(Format(.TextMatrix(vlintContador, 6), "")) & " where intnumcargo = " & .TextMatrix(vlintContador, 15) '@
                End If
                
                pActualizaDescuentoPuntos (vlintContador)
            End If
        End With
        
         vldblImporte = obtieneimporte(vlintContador)
'        If CDbl(IIf(Trim(grdArticulos.TextMatrix(vlintContador, 24)) = "", "0", Trim(grdArticulos.TextMatrix(vlintContador, 24)))) <> 0 Then


'            vldblImporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 24), ""))
'        Else


'            vldblImporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 2), ""))
'        End If

        
        
        
        ' Generar un registro de DetalleVentaPublico '
        With rsDetalleVentaPublico
            pNuevoDetalleVenta rsDetalleVentaPublico, vllngConsecutivoVenta, vlintContador
            'Cambio para caso 8736
            'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
            'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
            vlblnCuentaIngresoSaldada = False
            If flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "INGRESO") = flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "DESCUENTO") Then
                'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
                vlintBitSaldarCuentas = 1
                frsEjecuta_SP CStr(grdArticulos.TextMatrix(vlintContador, 13)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
                If vlintBitSaldarCuentas = 1 Then
                    
                    ' Abono para el Ingreso - Descuento '
                    
                    If (Val(grdArticulos.TextMatrix(vlintContador, 3)) * Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "")) - Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))) > 0 Then
                        pAgregarMovArregloCorte2 vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                                                IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                                                flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "INGRESO"), _
                                                ((Val(grdArticulos.TextMatrix(vlintContador, 3)) * vldblImporte) - Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))), False, "", 0, 0, "", 0, 2, _
                                                IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK", False
                        vlblnCuentaIngresoSaldada = True
                    ElseIf (Val(grdArticulos.TextMatrix(vlintContador, 3)) * Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "")) - Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))) < 0 Then
                        vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                    ElseIf (Val(grdArticulos.TextMatrix(vlintContador, 3)) * Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "")) - Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))) = 0 Then
                        vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                    End If
                End If
            End If
            
            If vlblnCuentaIngresoSaldada = False Then
                'Abono a la cuenta del ingreso
                pAgregarMovArregloCorte2 vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                                        IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                                        flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "INGRESO"), _
                                        Val(grdArticulos.TextMatrix(vlintContador, 3)) * vldblImporte, False, "", 0, 0, "", 0, 2, _
                                        IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK", False
                
                'Cargo a la cuenta de descuentos
                If Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")) <> 0 Then
                    pAgregarMovArregloCorte2 vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                                        IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                                        flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "DESCUENTO"), _
                                        Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")), True, "", 0, 0, "", 0, 2, _
                                        IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK", True
                End If
            End If
            .Update
        End With
        
        'Movimientos de trazabilidad
        If vgblnTrazabilidad = True Then
           pNuevaTrazabilidad vllngConsecutivoVenta, vlLngIvLoteProcedencia, vlintContador
        End If
        
        If vlBlnManejaLotes Then 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
        ' Graba las salida de los lotes ?LMM caso 20274
          pGrabaLotes Val(vllngConsecutivoVenta), "SVP", CInt(vllngdeptolote), "S", True
          vlBlnManejaLotes = False 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
        End If 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
    Next
    
    If vlblnConsultaTicketPrevio Then
        pEjecutaSentencia "update pvventapubliconofacturado set bitcancelado = 1, INTCVEVENTAPUBLICO = " & vllngConsecutivoVenta & " where intcveventa = " & Val(txtTicketPrevio.Text)
    End If
    
    ' Actualizar en la póliza el descuento por concepto de Deducible, Coaseguro, Copago en el caso en que se facturen estos conceptos
    If vldblFacturaPacienteTotal > 0 Then
        For a = 0 To UBound(aFPFacturaPaciente) - 1
            ' Descuento por la cantidad del DEDUCIBLE, CONSECUTIVO, COPAGO
            'Cargo a la cuenta de descuentos
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", flngCuentaConceptoDepartamento(aFPFacturaPaciente(a).lngConceptoFacturacion, vgintNumeroDepartamento, "DESCUENTO"), aFPFacturaPaciente(a).dblCantidad, True, "", 0, 0, "", 0, 2, Trim(vlstrFolioDocumento), "FA"
                       
            'Abono a la cuenta de IVA por pagar (Del DEDUCIBLE, COASEGURO y COPAGO)
            If aFPFacturaPaciente(a).dblCantidadIVA <> 0 Then
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", glngCtaIVACobrado, aFPFacturaPaciente(a).dblCantidadIVA, True, "", 0, 0, "", 0, 2, Trim(vlstrFolioDocumento), "FA"
            End If
        Next
    End If
    
    ' Poner los totales en el maestro de VentaPublico
    rsDetalleVentaPublico.Close
    
    ' Iniciar numero de cliente con CERO
    vllngNumCliente = 0
    
    ' Formas de pago para el corte de la factura BASE
    If vldblFacturaBaseTotal > 0 Then
        
        ' Afecta corte (PvDetalleCorte) (De la factura Base)
        
        For a = 0 To UBound(aFormasPago(), 1)
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
            IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
            0, IIf((aFormasPago(a).vldblTipoCambio = 0), aFormasPago(a).vldblCantidad, _
            aFormasPago(a).vldblDolares), False, CStr(vldtmFechaHoy + vldtmHoraHoy), CLng(aFormasPago(a).vlintNumFormaPago), aFormasPago(a).vldblTipoCambio, _
            IIf(Trim(aFormasPago(a).vlstrFolio) = "", "0", Trim(aFormasPago(a).vlstrFolio)), _
            vllngNumeroCorte, 1, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
            aFormasPago(a).vlbolEsCredito, aFormasPago(a).vlstrRFC, aFormasPago(a).vlstrBancoSAT, aFormasPago(a).vlstrBancoExtranjero, aFormasPago(a).vlstrCuentaBancaria, aFormasPago(a).vldtmFecha, "TIK"
        
            'Cargo a la cuenta de la forma de pago
            'pInsCortePoliza vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), aFormasPago(a).vllngCuentaContable, IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, aFormasPago(a).vldblDolares * aFormasPago(a).vldblTipoCambio), True
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
            IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
            aFormasPago(a).vllngCuentaContable, IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, _
            aFormasPago(a).vldblDolares * aFormasPago(a).vldblTipoCambio), True, "", 0, 0, "", 0, 2, _
            IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
            aFormasPago(a).vlbolEsCredito, aFormasPago(a).vlstrRFC, aFormasPago(a).vlstrBancoSAT, aFormasPago(a).vlstrBancoExtranjero, aFormasPago(a).vlstrCuentaBancaria, aFormasPago(a).vldtmFecha, "TIK"
            
            ' Agregado para caso 8741
            ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
            If aFormasPago(a).vllngCuentaComisionBancaria <> 0 And aFormasPago(a).vldblCantidadComisionBancaria <> 0 Then
                 ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                    IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                    aFormasPago(a).vllngCuentaComisionBancaria, aFormasPago(a).vldblCantidadComisionBancaria, True, "", 0, 0, "", 0, 2, _
                    IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "CBA"

                If aFormasPago(a).vldblIvaComisionBancaria <> 0 Then
                    ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                    pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                        IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                        glngCtaIVAPagado, aFormasPago(a).vldblIvaComisionBancaria, True, "", 0, 0, "", 0, 2, _
                        IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "CBA"
                End If
                ' Se genera un abono por la cantidad de la comisión bancaria y su iva que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                    IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                    aFormasPago(a).vllngCuentaContable, (aFormasPago(a).vldblCantidadComisionBancaria + aFormasPago(a).vldblIvaComisionBancaria), False, "", 0, 0, "", 0, 2, _
                    IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "CBA"
             End If
            
            
            ' Generar movimiento de CREDITO.
            If aFormasPago(a).vlbolEsCredito Then
                
                ' Para ver que numero de cliente es este paciente
                Set rsDatosCliente = frsRegresaRs("SELECT * FROM CcCliente INNER JOIN NoDepartamento on CcCliente.smicvedepartamento = NoDepartamento.smicvedepartamento WHERE CcCliente.intNumReferencia = " & Trim(str(vllngCveClienteCredito)) & " AND CcCliente.chrTipoCliente = '" & vlstrTipoPacienteCredito & "' AND NoDepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable, adLockReadOnly, adOpenForwardOnly)
                If rsDatosCliente.RecordCount = 0 Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    rsDatosCliente.Close
                    'Se detectó un error en la información del cliente.
                    MsgBox SIHOMsg(367), vbCritical, "Mensaje"
                    Exit Sub
                Else
                    vllngNumCliente = rsDatosCliente!intNumCliente
                End If
                
                pCrearMovtoCredito aFormasPago(a).vldblCantidad, vllngNumCliente, rsDatosCliente!INTNUMCUENTACONTABLE, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), vlstrFolioDocumento), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), vldblTipoCambio, vllngPersonaGraba
                rsDatosCliente.Close
            End If
        
            'Movimientos contables del IVA
            'Abono a la cuenta contable del IVA, según sea efectivo o crédito:
            If Val(Format(txtIva.Text, "###########.##")) <> 0 Then pAgregarMovArregloCorte2 vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), IIf(aFormasPago(a).vlbolEsCredito, glngCtaIVANoCobrado, glngCtaIVACobrado), Val(Format(txtIva.Text, "##########.##")) * IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, aFormasPago(a).vldblDolares * aFormasPago(a).vldblTipoCambio) / Val(Format(txtTotal.Text, "###########.##")), False, "", 0, 0, "", 0, 2, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK", False
                        
            'Movimientos contables del IEPS
            'Abono a la cuenta contable del IEPS, según sea efectivo o crédito:
            If Val(Format(txtIEPS.Text, "###########.##")) <> 0 Then pAgregarMovArregloCorte2 vllngNumeroCorte, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), IIf(aFormasPago(a).vlbolEsCredito, glngctaIEPSNoCobrado, glngctaIEPSCobrado), Val(Format(txtIEPS.Text, "##########.##")) * IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, aFormasPago(a).vldblDolares * aFormasPago(a).vldblTipoCambio) / Val(Format(txtTotal.Text, "###########.##")), False, "", 0, 0, "", 0, 2, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK", False
        Next a
    End If
    
    If intBitCuentaPuenteIngresos = 1 And vlstrFacturaTicket = "T" And blnFacturaAutomatica = False Then
        'Abono a la cuenta puente del ingreso
        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), _
                                IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), _
                                lngCuentaPuenteIngresos, _
                                vldblIngresosPuente, False, "", 0, 0, "", 0, 2, _
                                IIf(blnFacturaAutomatica, Trim(lstrFolioDocumentoF), Trim(vlstrFolioDocumento)), IIf((vlstrFacturaTicket = "F" Or blnFacturaAutomatica), "FA", "TI"), , , , , , , "TIK"
    End If
    
    
    ' Formas de pago para el corte del RECIBO del paciente
    If vldblCantidadPagoConceptos > 0 Then
        
        ' Afecta corte (PvDetalleCorte) (Del RECIBO del paciente)
        Set rsPvDetalleCorte = frsRegresaRs("SELECT * FROM PVDetalleCorte WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
        For a = 0 To UBound(aFormasPagoReciboPaciente(), 1)
                
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), _
            "FA", 0, IIf((aFormasPagoReciboPaciente(a).vldblTipoCambio = 0), aFormasPagoReciboPaciente(a).vldblCantidad, aFormasPagoReciboPaciente(a).vldblDolares), _
            False, CStr(vldtmFechaHoy + vldtmHoraHoy), CLng(aFormasPagoReciboPaciente(a).vlintNumFormaPago), aFormasPagoReciboPaciente(a).vldblTipoCambio, _
            IIf(Trim(aFormasPagoReciboPaciente(a).vlstrFolio) = "", "0", Trim(aFormasPagoReciboPaciente(a).vlstrFolio)), _
            vllngNumeroCorte, 1, "RECIBOPACIENTE", "FA", _
            aFormasPagoReciboPaciente(a).vlbolEsCredito, aFormasPagoReciboPaciente(a).vlstrRFC, aFormasPagoReciboPaciente(a).vlstrBancoSAT, aFormasPagoReciboPaciente(a).vlstrBancoExtranjero, aFormasPagoReciboPaciente(a).vlstrCuentaBancaria, aFormasPagoReciboPaciente(a).vldtmFecha
                            
            'Cargo a la cuenta de la forma de pago
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", aFormasPagoReciboPaciente(a).vllngCuentaContable, IIf(aFormasPagoReciboPaciente(a).vldblTipoCambio = 0, aFormasPagoReciboPaciente(a).vldblCantidad, aFormasPagoReciboPaciente(a).vldblDolares * aFormasPagoReciboPaciente(a).vldblTipoCambio), True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA"
            
            ' Agregado para caso 8741
            ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
            If aFormasPagoReciboPaciente(a).vllngCuentaComisionBancaria <> 0 And aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria <> 0 Then
                ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", aFormasPagoReciboPaciente(a).vllngCuentaComisionBancaria, aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria, True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
                If aFormasPagoReciboPaciente(a).vldblIvaComisionBancaria <> 0 Then
                    ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                    pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", glngCtaIVAPagado, aFormasPagoReciboPaciente(a).vldblIvaComisionBancaria, True, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
                End If
                ' Se genera un abono por la cantidad de la comisión bancaria y su iva que corresponde a la forma de pago
                pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, Trim(vlstrFolioDocumento), "FA", aFormasPagoReciboPaciente(a).vllngCuentaContable, (aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria + aFormasPagoReciboPaciente(a).vldblIvaComisionBancaria), False, "", 0, 0, "", 0, 2, "FACTURAPACIENTE", "FA", , , , , , , "CBA"
             End If
         Next a
    End If

    If vlstrFacturaTicket = "F" Then
        
        ' Guardar en la factura y su detalle (Factura BASE)
        Set rsTipoAgrupacion = frsRegresaRs("SELECT intTipoAgrupaDigital FROM Formato WHERE Formato.INTNUMEROFORMATO = " & llngFormatoFacturaAUsar, adLockReadOnly, adOpenForwardOnly)
        intTipoAgrupacion = IIf(IsNull(rsTipoAgrupacion!intTipoAgrupaDigital), "1", rsTipoAgrupacion!intTipoAgrupaDigital)
        
        ' Si se importa la venta para generar una factura global tiene que sacarse en formato desglosado por cargo para poder incluir el folio del ticket a nivel de concepto en el CFDI
        If vlblnImportoVenta = True Then intTipoAgrupacion = 3
        
        If optPesos(0).Value Then vldblTipoCambio = 1
        
        Set rsFactura = frsRegresaRs("SELECT * FROM PvFactura WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
        With rsFactura
            .AddNew
            !chrfoliofactura = vlstrFolioDocumento
            !dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
            !CHRRFC = IIf(vlBitExtranjero, "XEXX010101000", IIf(Len(fStrRFCValido(vlstrRFC)) < 12 Or Len(fStrRFCValido(vlstrRFC)) > 13, "XAXX010101000", fStrRFCValido(vlstrRFC)))
            !vchSerie = strSerieDocumento
            !INTFOLIO = strFolioDocumento
            !CHRNOMBRE = vlstrNombreFactura
            !chrCalle = vlstrDireccion
            !VCHREGIMENFISCALRECEPTOR = vgstrRegimenFiscal
            !VCHNUMEROEXTERIOR = vlstrNumeroExterior
            !VCHNUMEROINTERIOR = vlstrNumeroInterior
            !VCHCOLONIA = vlstrColonia
            !VCHCODIGOPOSTAL = vlstrCP
            !chrTelefono = vlstrTelefono
            !smyIVA = Round(((Val(Format(txtIva.Text, "")) - IIf(blnFacturarConceptosSeguro, vldblFacturaPacienteIVA, 0)) / vldblTipoCambio), 2)
            !MNYDESCUENTO = Round(((Val(Format(txtDescuentos.Text, "")) + IIf(blnFacturarConceptosSeguro, vldblTotalControlAseguradorasSinIVA, 0)) / vldblTipoCambio), 2)
            !chrEstatus = " "
            !INTMOVPACIENTE = Val(txtMovimientoPaciente.Text)
            !CHRTIPOPACIENTE = IIf(Val(txtMovimientoPaciente.Text) = 0, "V", "E")
            !SMIDEPARTAMENTO = vgintNumeroDepartamento
            !intCveEmpleado = vllngPersonaGraba
            !intNumCorte = vllngNumeroCorte
            !mnyAnticipo = 0
            !mnyTotalFactura = Round(((Val(Format(txtTotal.Text, "")) - vldblFacturaPacienteTotal) / vldblTipoCambio), 2)
            !BITPESOS = IIf(optPesos(0).Value, 1, 0)
            !mnytipocambio = IIf(optPesos(0).Value, 0, vldblTipoCambio)
            !chrTipoFactura = IIf(vlstrTipoPacienteCredito = "CO", "E", "P")
            !chrIncluirConceptosSeguro = IIf(lblnExisteControl, vlchrIncluirConceptosSeguro, Null)
            !intCveVentaPublico = vllngConsecutivoVenta
            !intNumCliente = vllngNumCliente
            !intCveCiudad = llngCveCiudad
            !intcveempresa = IIf(vlstrTipoPacienteCredito = "CO", vllngCveClienteCredito, 0)
            !mnyTotalPagar = Round(((Val(Format(txtTotal.Text, "")) - vldblFacturaPacienteTotal) / vldblTipoCambio), 2)
            !bitdesgloseIEPS = IIf(vlblnImportoVenta, 1, IIf(vlblnSujetoIEPS, 1, 0))
            !intTipoDetalleFactura = intTipoAgrupacion
            !intCveUsoCFDI = IIf(vlintUsoCFDI = 0, 64, vlintUsoCFDI)
            !bitFacturaGlobal = IIf(vlblnImportoVenta = True, 1, 0)
            .Update
        End With
        
        vllngConsecutivoFactura = flngObtieneIdentity("SEC_PvFactura", rsFactura!intConsecutivo)
        
        ' Detalle de la Factura
        pGrabaDetalleFactura vlstrFolioDocumento, vldblFacturaPacienteIVA, vldblTotalControlAseguradoras, ldblDeducible, ldblCoaseguro, ldblCopago, IIf(vlblnImportoVenta, True, vlblnSujetoIEPS)
                
        ' CANTIDADES Y TASAS IEPS
        If vlblnLicenciaIEPS Then pGrabaTasasIEPS vllngConsecutivoFactura
        
        ' Ajustar PVFACTURAIMPORTE
        Set rsPvFacturaImporte = frsRegresaRs("SELECT * FROM PVFACTURAIMPORTE WHERE INTCONSECUTIVO = " & vllngConsecutivoFactura, adLockOptimistic, adOpenDynamic)
        If rsPvFacturaImporte.RecordCount = 0 Then frsEjecuta_SP str(vllngConsecutivoFactura) & "|0|0|0|0", "SP_PVINSFACTURAIMPORTE"
        

        vldblimportegravado = 0
        vldblSumatoria = 0
                   
        Set rsCalcPvFactImp = frsRegresaRs(IIf(rsFactura!smyIVA = 0, "SELECT nvl(ROUND(NVL(SUM((PVDETALLEVENTAPUBLICO.MnyPrecio * PVDETALLEVENTAPUBLICO.IntCantidad) + PVDETALLEVENTAPUBLICO.MNYIEPS - PVDETALLEVENTAPUBLICO.MnyDescuento),0),2),0) AS ImporteGravado FROM PVVENTAPUBLICO INNER JOIN PVDETALLEVENTAPUBLICO ON PVDETALLEVENTAPUBLICO.IntCveVenta = PVVENTAPUBLICO.IntCveVenta WHERE (PVDETALLEVENTAPUBLICO.MnyIva = -1) AND PVVENTAPUBLICO.ChrFolioFactura = '" & rsFactura!chrfoliofactura & "'", "SELECT nvl(ROUND(NVL(SUM((PVDETALLEVENTAPUBLICO.MnyPrecio * PVDETALLEVENTAPUBLICO.IntCantidad) + PVDETALLEVENTAPUBLICO.MNYIEPS - PVDETALLEVENTAPUBLICO.MnyDescuento),0),2),0) AS ImporteGravado FROM PVVENTAPUBLICO INNER JOIN PVDETALLEVENTAPUBLICO ON PVDETALLEVENTAPUBLICO.IntCveVenta = PVVENTAPUBLICO.IntCveVenta WHERE (PVDETALLEVENTAPUBLICO.MnyIva > 0) AND PVVENTAPUBLICO.ChrFolioFactura = '" & rsFactura!chrfoliofactura & "'"), adLockOptimistic, adOpenDynamic)
        If rsCalcPvFactImp.RecordCount > 0 Then
            vldblimportegravado = rsCalcPvFactImp!ImporteGravado
        End If
                    
        Set rsCalcPvFactImp = frsRegresaRs("SELECT nvl((ROUND(NVL(SUM(MnyCantidad),0),2)),0) AS ImporteConceptosSegudo FROM PVDETALLEFACTURA WHERE ChrFolioFactura = '" & rsFactura!chrfoliofactura & "' AND ChrTipo = 'OD'", adLockOptimistic, adOpenDynamic)
        If rsCalcPvFactImp.RecordCount > 0 Then
            vldblSumatoria = rsCalcPvFactImp!ImporteConceptosSegudo
        End If
    
        If rsFactura!BITPESOS = 1 Then
            pEjecutaSentencia "UPDATE PVFACTURAIMPORTE SET MNYSUBTOTALGRAVADO = " & vldblimportegravado - vldblSumatoria & ", MNYSUBTOTALNOGRAVADO = " & rsFactura!mnyTotalFactura - rsFactura!smyIVA - vldblimportegravado - vldblSumatoria & " WHERE IntConsecutivo = " & vllngConsecutivoFactura
        Else
            pEjecutaSentencia "UPDATE PVFACTURAIMPORTE SET MNYSUBTOTALGRAVADO = " & Round((vldblimportegravado - vldblSumatoria) / rsFactura!mnytipocambio, 2) & ", MNYSUBTOTALNOGRAVADO = " & rsFactura!mnyTotalFactura - rsFactura!smyIVA - Round((vldblimportegravado - vldblSumatoria) / rsFactura!mnytipocambio, 2) & " WHERE IntConsecutivo = " & vllngConsecutivoFactura
        End If
        rsFactura.Close
        
        ' Poner el número de factura en los cargos
        vlstrSentencia = "UPDATE PvCargo SET chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "' WHERE intNumCargo IN ("
        For a = 1 To grdArticulos.Rows - 1
            vlstrSentencia = vlstrSentencia & Trim(grdArticulos.TextMatrix(a, 15))
            vlstrSentencia = IIf(a = grdArticulos.Rows - 1, vlstrSentencia & ")", vlstrSentencia & ",")
        Next a
        pEjecutaSentencia (vlstrSentencia)
        
        ' Que se actualice el CONTROL DE ASEGURADORAS '
        pEjecutaSentencia "UPDATE PvControlAseguradora SET chrFolioFacturaEmpresa = '" & Trim(vlstrFolioDocumento) & "' WHERE intMovPaciente = " & Trim(str(Val(txtMovimientoPaciente.Text))) & " AND chrTipoPaciente = 'E' AND (chrFolioFacturaDeducible IS NULL OR chrFolioFacturaCoaseguro IS NULL OR chrFolioFacturaCopago IS NULL OR chrFolioFacturaExcedente IS NULL OR chrFolioFacturaEmpresa IS NULL) AND intCveEmpresa = " & str(vgintEmpresa)
        
        'Guardar información de la forma de pago en tabla intermedia
        If vldblFacturaBaseTotal > 0 And vlblnOkFormasPago Then
            For a = 0 To UBound(aFormasPago(), 1)
                If Not aFormasPago(a).vlbolEsCredito Then 'Formas de pago distintas a Crédito
                    frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(a).vlintNumFormaPago & "|" & aFormasPago(a).lngIdBanco & "|" & _
                                        IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, aFormasPago(a).vldblDolares) & "|" & IIf(aFormasPago(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(a).vldblTipoCambio & "|" & _
                                        fstrTipoMovimientoForma(aFormasPago(a).vlintNumFormaPago, "F") & "|" & "FA" & "|" & vllngConsecutivoFactura & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & _
                                        fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                    vldblComisionIvaBancaria = 0
                    If aFormasPago(a).vllngCuentaComisionBancaria <> 0 And aFormasPago(a).vldblCantidadComisionBancaria <> 0 Then
                        
                        vldblComisionIvaBancaria = (aFormasPago(a).vldblCantidadComisionBancaria + aFormasPago(a).vldblIvaComisionBancaria) * -1
                        If aFormasPago(a).vldblTipoCambio <> 0 Then
                            vldblComisionIvaBancaria = (aFormasPago(a).vldblCantidadComisionBancaria + aFormasPago(a).vldblIvaComisionBancaria) / aFormasPago(a).vldblTipoCambio * -1
                        End If
                        
                        frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(a).vlintNumFormaPago & "|" & aFormasPago(a).lngIdBanco & "|" & _
                                            vldblComisionIvaBancaria & "|" & IIf(aFormasPago(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(a).vldblTipoCambio & "|" & _
                                            "CBA" & "|" & "FA" & "|" & vllngConsecutivoFactura & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & _
                                            fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    End If
                End If
            Next a
        End If
                
        'Guardar información de la forma de pago en tabla intermedia para recibo del paciente
        If vldblCantidadPagoConceptos > 0 And vlblnOkFormasPagoRecPac Then
            For a = 0 To UBound(aFormasPagoReciboPaciente(), 1)
                If Not aFormasPagoReciboPaciente(a).vlbolEsCredito Then 'Formas de pago distintas a Crédito
                    frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPagoReciboPaciente(a).vlintNumFormaPago & "|" & aFormasPagoReciboPaciente(a).lngIdBanco & "|" & _
                                        IIf(aFormasPagoReciboPaciente(a).vldblTipoCambio = 0, aFormasPagoReciboPaciente(a).vldblCantidad, aFormasPagoReciboPaciente(a).vldblDolares) & "|" & IIf(aFormasPagoReciboPaciente(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPagoReciboPaciente(a).vldblTipoCambio & "|" & _
                                        fstrTipoMovimientoForma(aFormasPagoReciboPaciente(a).vlintNumFormaPago, "F") & "|" & "FA" & "|" & vllngConsecutivoFactura & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & _
                                        fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                    vldblComisionIvaBancaria = 0
                    If aFormasPagoReciboPaciente(a).vllngCuentaComisionBancaria <> 0 And aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria <> 0 Then
                        
                        vldblComisionIvaBancaria = (aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria + aFormasPagoReciboPaciente(a).vldblIvaComisionBancaria) * -1
                        If aFormasPagoReciboPaciente(a).vldblTipoCambio <> 0 Then
                            vldblComisionIvaBancaria = (aFormasPagoReciboPaciente(a).vldblCantidadComisionBancaria + aFormasPagoReciboPaciente(a).vldblIvaComisionBancaria) / aFormasPagoReciboPaciente(a).vldblTipoCambio * -1
                        End If
                        
                        frsEjecuta_SP vgstrParametrosSP = vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPagoReciboPaciente(a).vlintNumFormaPago & "|" & aFormasPagoReciboPaciente(a).lngIdBanco & "|" & _
                                            vldblComisionIvaBancaria & "|" & IIf(aFormasPagoReciboPaciente(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPagoReciboPaciente(a).vldblTipoCambio & "|" & _
                                            "CBA" & "|" & "FA" & "|" & vllngConsecutivoFactura & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & _
                                            fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    End If
                End If
            Next a
        End If
       
        ' actualizacion de datos fiscales (sólo aplica para empleados y para medicos)
        If (vgstrTipoPaciente = "EM" Or vgstrTipoPaciente = "ME") And Trim(txtMovimientoPaciente.Text) <> "" Then
           If vlstrRFCFacturaPaciente <> "" Then ' lleva factura de empresa pero sólo se actualizan los datos fiscales del paciente
              If vlstrRFCFacturaPaciente <> "XEXX010101000" And vlstrRFCFacturaPaciente <> "XAXX010101000" Then
                 frsEjecuta_SP Me.txtMovimientoPaciente.Text & "|0|" & fStrRFCValido(vlstrRFCFacturaPaciente) & "|" & vlstrDireccionFacturaPaciente & "|" & vlstrNumeroExteriorFacturaPaciente & "|" & vlstrNumeroInteriorFacturaPaciente & "|" & vlstrColoniaFacturaPaciente & "|" & llngCveCiudad & "|" & vlstrCPFacturaPaciente & "|" & vlstrTelefonoFacturaPaciente, "SP_UpdDatosFiscalesExPaciente"
              End If
           Else ' no lleva factura de empresa entonces se actualizan los datos fiscales del paciente en expaciente
              If vlstrRFC <> "XEXX010101000" And vlstrRFC <> "XAXX010101000" Then
                 frsEjecuta_SP Me.txtMovimientoPaciente.Text & "|0|" & fStrRFCValido(vlstrRFC) & "|" & vlstrDireccion & "|" & vlstrNumeroExterior & "|" & vlstrNumeroInterior & "|" & vlstrColonia & "|" & llngCveCiudad & "|" & vlstrCP & "|" & vlstrTelefono, "SP_UpdDatosFiscalesExPaciente"
              End If
           End If
        End If
        
        pActualizaEmpleadoIngresa

        'Se registran los puntos acumulados y los puntos utilizados en la factura
        pControlPuntos vllngConsecutivoFactura, vlstrFolioDocumento, vldtmFechaHoy, vldtmHoraHoy

        ' Cerrar cuenta del paciente Externo
        ' Actualizar fecha de egreso del paciente Externo y poner la cuenta como facturada '
        'If Trim(txtMovimientoPaciente.Text) <> "" And optPaciente.Value Then
        If Trim(txtMovimientoPaciente.Text) <> "" Then frsEjecuta_SP Trim(txtMovimientoPaciente.Text), " SP_PVUPDCUENTAPACIENTEPOS"
               
        'Guardamos el log de la transacción
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "VENTAS AL PUBLICO", CStr(vllngConsecutivoVenta))
        'Insertamos movimientos en el corte y actualizamos los cortes si es que el corte de cerro mientras se hacia la venta(por si la moscas)
        vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte, True)
        If vllngCorteUsado <> 0 Then vllngCorteUsado = fRegistrarMovArregloCorte2(vllngCorteUsado, True)
        
        If vllngCorteUsado = 0 Then 'error al insertar los movimientos en el corte
           EntornoSIHO.ConeccionSIHO.RollbackTrans
           'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
           MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
           Exit Sub
        End If
                          
        If vllngCorteUsado <> vllngNumeroCorte Then
          If vldblFacturaPacienteTotal > 0 Then pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & vllngConsecFacPac
          pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & vllngConsecutivoFactura
          pEjecutaSentencia "UPdate pvventapublico set INTNUMCORTE = " & vllngCorteUsado & " where intcveventa = " & vllngConsecutivoVenta
        End If
          
         pRefacturacionAnteriores
          
        'VALIDACIÓN DE LOS DATOS ANTES DE INSERTAR EN GNCOMPROBANTEFISCLADIGITAL EN EL PROCESO DE TIMBRADO
        If intTipoEmisionComprobante = 2 Then
           If vldblFacturaPacienteTotal > 0 Then
              If Not fblnValidaDatosCFDCFDi(vllngConsecFacPac, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacionPaciente), strNumeroAprobacionPaciente) Then
                 EntornoSIHO.ConeccionSIHO.RollbackTrans
                 Exit Sub
              End If
           End If
           
           If Not fblnValidaDatosCFDCFDi(vllngConsecutivoFactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacionPaciente), strNumeroAprobacionPaciente) Then
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                Exit Sub
           End If
        End If
        
        'FIN de Transacción
        EntornoSIHO.ConeccionSIHO.CommitTrans ' hasta aqui ya tenemos la(s) factura(s)
                
        'Factura digital
        If intTipoEmisionComprobante = 2 Then
        
           'TIMBRE DE LA FACTURA DEL PACIENTE
           If vldblFacturaPacienteTotal > 0 Then
              pBarraCFD3
              If intTipoCFDFactura = 1 Then
                 pMarcarPendienteTimbre vllngConsecFacPac, "FA", vgintNumeroDepartamento 'factura de paciente pendiente de timbre
                 pLogTimbrado 2
              End If
              EntornoSIHO.ConeccionSIHO.BeginTrans
              pLogTimbrado 2
              If Not fblnGeneraComprobanteDigital(vllngConsecFacPac, "FA", intTipoAgrupacion, CInt(strAnoAprobacionPaciente), strNumeroAprobacionPaciente, IIf(intTipoCFDFactura = 1, True, False)) Then
                 On Error Resume Next
                 If intTipoCFDFactura = 1 Then pLogTimbrado 1 'guarda el log del timbre
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 If vgIntBanderaTImbradoPendiente = 1 Then
                    If intTipoCFDFactura = 1 Then
                       intcontadorCFDiPendienteCancelar = 0
                       ReDim vlArrCFDiPendienteCancelar(0)
                       pCFDiPendienteCancelar vllngConsecFacPac, "FA"
                       pCFDiPendienteCancelar 0, "", 1
                    End If
                    '(1306) El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal.
                    '(1319) El proceso ha quedado incompleto.
                    '(1307) La factura será cancelada en el sistema, será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT.
                    MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura del paciente") & vbNewLine & SIHOMsg(1319) & vbNewLine & Replace(Replace(SIHOMsg(1307), "La factura será cancelada", "La factura del paciente y la factura de la empresa serán canceladas"), ".", " de la factura del paciente."), vbCritical + vbOKOnly, "Mensaje"
                 ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then 'No se realizó el timbrado
                        '(33)   ¡No se pueden guardar los datos!
                        '(1319) El proceso ha quedado incompleto.
                        '(1307) La factura será cancelada en el sistema, será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT.
                        MsgBox SIHOMsg(33) & vbNewLine & SIHOMsg(1319) & vbNewLine & Replace(Replace(SIHOMsg(1307), "La factura será cancelada", "La factura del paciente y la factura de la empresa serán canceladas"), ", será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT", ""), vbCritical + vbOKOnly, "Mensaje"
                        If intTipoCFDFactura = 1 Then pEliminaPendientesTimbre vllngConsecFacPac, "FA"
                 End If
                 pCancelarFactura Trim(vlstrFolioDocumentoPaciente), vllngPersonaGraba, "frmPOS", True, False  'cancelación de la factura del paciente
                 'agregamos la factura de la empresa a  gncomprobantefiscaldigital
                 fblnGeneraComprobanteDigital vllngConsecutivoFactura, "FA", intTipoAgrupacion, CInt(strAnoAprobacionDocumento), strNumeroAprobacionDocumento, IIf(intTipoCFDFactura = 1, True, False), , , False
                                  
                 pCancelarFactura Trim(vlstrFolioDocumento), vllngPersonaGraba, "frmPOS", False, False 'cancelación de la factura de la empresa
                 ' se imprimen ambas facturas
                 fblnImprimeComprobanteDigital vllngConsecFacPac, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                 fblnImprimeComprobanteDigital vllngConsecutivoFactura, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                 pQuitaBarraCFD
                 pNuevo
                 pEnfocaTextBox txtClaveArticulo
                 cboMedico.ListIndex = 0
                 Exit Sub 'termina todo
              Else 'timbrado de la factura del paciente EXITOSO
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 If intTipoCFDFactura = 1 Then
                    pEliminaPendientesTimbre vllngConsecFacPac, "FA"
                    pLogTimbrado 1
                 End If
                 pBarraCFD4
              End If
           End If
                      
           'TIMBRE DE LA FACTURA BASE/EMPRESA/PACIENTE(CUANDO EL PACIENTE ES TIPO ASEGURADORA PERO SE REALIZA LA VENTA A LA EMPRESA) SIEMPRE QUE HAY FACTURA ENTRA
           pBarraCFD3
           If intTipoCFDFactura = 1 Then
              pMarcarPendienteTimbre vllngConsecutivoFactura, "FA", vgintNumeroDepartamento 'factura base pendiente de timbre
              pLogTimbrado 2
           End If
           EntornoSIHO.ConeccionSIHO.BeginTrans
                      
           If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", intTipoAgrupacion, CInt(strAnoAprobacionDocumento), strNumeroAprobacionDocumento, IIf(intTipoCFDFactura = 1, True, False)) Then
              On Error Resume Next
              
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 If intTipoCFDFactura = 1 Then pLogTimbrado 1
                 If vgIntBanderaTImbradoPendiente = 1 Then
                    If vldblFacturaPacienteTotal > 0 Then
                       If intTipoCFDFactura = 1 Then
                          intcontadorCFDiPendienteCancelar = 0
                          ReDim vlArrCFDiPendienteCancelar(0)
                          pCFDiPendienteCancelar vllngConsecFacPac, "FA" 'pendiente cancelar ante el SAT Paciente
                          pCFDiPendienteCancelar vllngConsecutivoFactura, "FA"         'pendiente cancelar ante el SAT Empresa
                          pCFDiPendienteCancelar 0, "", 1                              'guardamos pendiente de cancelación
                       End If
                                              
                       '(1306) El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal.
                       '(1319) El proceso ha quedado incompleto.
                       '(1307) La factura será cancelada en el sistema, será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT.
                       MsgBox Replace(SIHOMsg(1306), "El comprobante", "La Factura de la empresa") & vbNewLine & SIHOMsg(1319) & vbNewLine & Replace(Replace(Replace(SIHOMsg(1307), "La factura será cancelada", "Ambas facturas serán canceladas"), "fiscal", "fiscal de la factura de la empresa"), ".", " de ambos documentos"), vbCritical + vbOKOnly, "Mensaje"
                       
                       pCancelarFactura Trim(vlstrFolioDocumentoPaciente), vllngPersonaGraba, "frmPOS", True, False  'cancelación de la factura del paciente
                       pCancelarFactura Trim(vlstrFolioDocumento), vllngPersonaGraba, "frmPOS", False, False 'cancelación de la factura de la empresa
                       pQuitaBarraCFD
                       fblnImprimeComprobanteDigital vllngConsecFacPac, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                       fblnImprimeComprobanteDigital vllngConsecutivoFactura, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                       pNuevo
                       pEnfocaTextBox txtClaveArticulo
                       cboMedico.ListIndex = 0
                       Exit Sub 'termina todo
                    Else 'no hay factura del paciente, solo Base
                       '(1306) El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal.
                       MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura"), vbInformation + vbOKOnly, "Mensaje"
                       pQuitaBarraCFD
                    End If
                 
                 ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then
                        EntornoSIHO.ConeccionSIHO.CommitTrans
                                             
                     If vldblFacturaPacienteTotal > 0 Then
                        If intTipoCFDFactura = 1 Then
                            intcontadorCFDiPendienteCancelar = 0
                            ReDim vlArrCFDiPendienteCancelar(0)
                            pCFDiPendienteCancelar vllngConsecFacPac, "FA" 'pendiente cancelar ante el SAT Paciente
                            pCFDiPendienteCancelar 0, "", 1                              'guardamos pendiente de cancelación
                        End If
                                            
                        '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
                        '(1307) La factura será cancelada en el sistema, será necesario realizar la cancelación ante el SAT.
                        MsgBox Replace(Replace(SIHOMsg(1338), "La factura", "La factura de la empresa"), ", será cancelada en el sistema", "") & vbNewLine & SIHOMsg(1319) & vbNewLine & Replace(Replace(Replace(SIHOMsg(1307), "La factura será cancelada", "Ambas facturas serán canceladas"), "confirmar el timbre fiscal y ", ""), ".", " de la factura del paciente."), vbCritical + vbOKOnly, "Mensaje"
                       
                        pCancelarFactura Trim(vlstrFolioDocumentoPaciente), vllngPersonaGraba, "frmPOS", True, False  'cancelación de la factura del paciente
                        pCancelarFactura Trim(vlstrFolioDocumento), vllngPersonaGraba, "frmPOS", False, False 'cancelación de la factura de la empresa
                        pEliminaPendientesTimbre vllngConsecutivoFactura, "FA"
                        pQuitaBarraCFD
                        fblnImprimeComprobanteDigital vllngConsecFacPac, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                        fblnImprimeComprobanteDigital vllngConsecutivoFactura, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                        pNuevo
                        pEnfocaTextBox txtClaveArticulo
                        cboMedico.ListIndex = 0
                        Exit Sub 'termina todo
                    Else
                        '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
                        MsgBox SIHOMsg(1338), vbCritical + vbOKOnly, "Mensaje"
                        pCancelarFactura Trim(vlstrFolioDocumento), vllngPersonaGraba, "frmPOS", False  'cancelación de la factura de la empresa
                        
                        'Actualiza PDF al cancelar facturas
                        If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", 1, 0, "", False, True, -1) Then On Error Resume Next
                        
                        pEliminaPendientesTimbre vllngConsecutivoFactura, "FA"
                        fblnImprimeComprobanteDigital vllngConsecutivoFactura, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion
                        pQuitaBarraCFD
                        pNuevo
                        pEnfocaTextBox txtClaveArticulo
                        cboMedico.ListIndex = 0
                        Exit Sub 'termina todo
                    End If
                 End If
           Else 'timbrado correcto
              EntornoSIHO.ConeccionSIHO.CommitTrans
              If intTipoCFDFactura = 1 Then
                 pEliminaPendientesTimbre vllngConsecutivoFactura, "FA"
                 pLogTimbrado 1 'guarda el log del timbre
              End If
              pQuitaBarraCFD
           End If
        End If
             
        ' Impresión del RECIBO del paciente
        If vldblCantidadPagoConceptos > 0 Then
            Set rsReporte = frsRegresaRs("SELECT * FROM PARAMETROS")
            If rsReporte.RecordCount > 0 Then
              pInstanciaReporte vgrptReporte, "rptReciboPagoPos.rpt"
              vgrptReporte.DiscardSavedData
            
              alstrParametros(0) = "Cantidad;" & FormatCurrency(vldblCantidadPagoConceptos, 2)
              alstrParametros(1) = "CantidadLetras;" & fstrNumeroenLetras(vldblCantidadPagoConceptos, "PESOS", "M.N.")
              vlStrConcepto = "PAGO DE "
              If Not blnFacturaCoaseguro And ldblCoaseguro > 0 Then vlStrConcepto = vlStrConcepto & "COASEGURO"
              
              If Not blnFacturaDeducible And ldblDeducible > 0 Then
                  If vlStrConcepto <> "PAGO DE " Then vlStrConcepto = vlStrConcepto & ", "
                  vlStrConcepto = vlStrConcepto & "DEDUCIBLE"
              End If
              If Not blnFacturaCopago And ldblCopago > 0 Then
                  If vlStrConcepto <> "PAGO DE " Then vlStrConcepto = vlStrConcepto & ", "
                  vlStrConcepto = vlStrConcepto & "COPAGO"
              End If

              pTerminaParametrosReporte rsReporte!INTCIUDAD, vlStrConcepto

              pCargaParameterFields alstrParametros, vgrptReporte
              
              pImprimeReporte vgrptReporte, rsReporte, "P", "Recibo de pago"
            Else
                MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetros.
            End If
            If rsReporte.State <> adStateClosed Then rsReporte.Close
        End If
        
        'Impresión de la Factura (del PACIENTE)
        MsgBox SIHOMsg(420) & vbNewLine & SIHOMsg(343), vbInformation, "Mensaje"
        
        If vldblFacturaPacienteTotal > 0 Then
          'Facturación digital
            If intTipoEmisionComprobante = 2 Then
               If Not fblnImprimeComprobanteDigital(vllngConsecFacPac, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion) Then
                   Exit Sub
               End If
                
               'Verifica el parámetro de envío de CFDs por correo
               If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) And vlstrTipoPacienteCredito <> "CO" Then
                    '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                    If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        pEnviarCFD "FA", vllngConsecFacPac, CLng(vgintClaveEmpresaContable), Trim(vlstrRFCFacturaPaciente), vllngPersonaGraba, Me
                    End If
               End If
            Else
                pImprimeFormato llngFormatoFacturaAUsar, vllngConsecFacPac
            End If
        End If
        
        'Impresión de la Factura (BASE)
        If intTipoEmisionComprobante = 2 Then
            If Not fblnImprimeComprobanteDigital(vllngConsecutivoFactura, "FA", "I", llngFormatoFacturaAUsar, intTipoAgrupacion) Then
                Exit Sub
            End If
                'Verifica el parámetro de envío de CFDs por correo
                If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) And vlstrTipoPacienteCredito <> "CO" And vgIntBanderaTImbradoPendiente = 0 Then
                    '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                    If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                        pEnviarCFD "FA", vllngConsecutivoFactura, CLng(vgintClaveEmpresaContable), Trim(vlstrRFC), vllngPersonaGraba, Me
                    End If
                End If
         Else
            pImprimeFormato llngFormatoFacturaAUsar, vllngConsecutivoFactura
        End If
    Else 'Si es una "T" Para los TICKETS, osea que selecciono NO a Factura
        'Si se factura automáticamente
        If blnFacturaAutomatica Then
           
           EntornoSIHO.ConeccionSIHO.Execute "SAVEPOINT FACTURA" 'SAVE POINT
           'Las facturas automáticas siempre son en pesos
           vldblTipoCambio = 1
           
           'Guardar la factura y su detalle
           Set rsTipoAgrupacion = frsRegresaRs("SELECT intTipoAgrupaDigital FROM Formato WHERE Formato.INTNUMEROFORMATO = " & vllngFormatoaUsar, adLockReadOnly, adOpenForwardOnly)
           intTipoAgrupacion = IIf(IsNull(rsTipoAgrupacion!intTipoAgrupaDigital), "1", rsTipoAgrupacion!intTipoAgrupaDigital)

           Set rsFactura = frsRegresaRs("SELECT * FROM PvFactura WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
           With rsFactura
                .AddNew
                !chrfoliofactura = lstrFolioDocumentoF
                !dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
                !CHRRFC = vlstrRFC
                !VCHREGIMENFISCALRECEPTOR = vgstrRegimenFiscal
                !vchSerie = strSerieDocumentoF
                !INTFOLIO = strFolioDocumentoF
                !CHRNOMBRE = vlstrNombreFactura
                !chrCalle = vlstrDireccion
                !VCHNUMEROEXTERIOR = vlstrNumeroExterior
                !VCHNUMEROINTERIOR = vlstrNumeroInterior
                !VCHCOLONIA = vlstrColonia
                !VCHCODIGOPOSTAL = vlstrCP
                !chrTelefono = vlstrTelefono
                !intCveCiudad = llngCveCiudad
                !smyIVA = Val(Format(txtIva.Text, ""))
                !MNYDESCUENTO = Val(Format(txtDescuentos.Text, ""))
                !mnyTotalFactura = Val(Format(txtTotal.Text, ""))
                !mnyTotalPagar = Val(Format(txtTotal.Text, ""))
                !mnyAnticipo = 0
                !BITPESOS = 1
                !mnytipocambio = 0
                !chrEstatus = " "
                !INTMOVPACIENTE = Val(txtMovimientoPaciente.Text)
                !CHRTIPOPACIENTE = IIf(Val(txtMovimientoPaciente.Text) = 0, "V", "E")
                !chrTipoFactura = IIf(vlstrTipoPacienteCredito = "CO", "E", "P")
                !chrIncluirConceptosSeguro = IIf(lblnExisteControl, vlchrIncluirConceptosSeguro, Null)
                !SMIDEPARTAMENTO = vgintNumeroDepartamento
                !intCveEmpleado = vllngPersonaGraba
                !intNumCorte = vllngNumeroCorte
                !intCveVentaPublico = vllngConsecutivoVenta
                !intNumCliente = vllngNumCliente
                !intcveempresa = IIf(vlstrTipoPacienteCredito = "CO", vllngCveClienteCredito, 0)
'                !bitdesgloseIEPS = IIf(vlblnImportoVenta, 1, IIf(vlblnSujetoIEPS, 1, 0))
                !bitdesgloseIEPS = 0
                !intTipoDetalleFactura = intTipoAgrupacion
                !intCveUsoCFDI = IIf(vlintUsoCFDI = 0, 64, vlintUsoCFDI)
                !bitFacturaGlobal = IIf(vlblnImportoVenta = True, 1, 0)
                .Update
           End With
           vllngConsecutivoFactura = flngObtieneIdentity("SEC_PvFactura", rsFactura!intConsecutivo)
           pGrabaDetalleFactura lstrFolioDocumentoF, 0, 0, 0, 0, 0, IIf(vlblnImportoVenta, 1, IIf(vlblnSujetoIEPS, 1, 0))
                      
           ' CANTIDADES Y TASAS IEPS
           If vlblnLicenciaIEPS Then pGrabaTasasIEPS vllngConsecutivoFactura

           ' Ajustar PVFACTURAIMPORTE '
           Set rsPvFacturaImporte = frsRegresaRs("SELECT * FROM PVFACTURAIMPORTE WHERE INTCONSECUTIVO = " & vllngConsecutivoFactura, adLockOptimistic, adOpenDynamic)
            If rsPvFacturaImporte.RecordCount = 0 Then frsEjecuta_SP str(vllngConsecutivoFactura) & "|0|0|0|0", "SP_PVINSFACTURAIMPORTE"

            vldblimportegravado = 0
            vldblSumatoria = 0
            Set rsCalcPvFactImp = frsRegresaRs(IIf(rsFactura!smyIVA = 0, "SELECT nvl(ROUND(NVL(SUM((PVDETALLEVENTAPUBLICO.MnyPrecio * PVDETALLEVENTAPUBLICO.IntCantidad) + PVDETALLEVENTAPUBLICO.MNYIEPS - PVDETALLEVENTAPUBLICO.MnyDescuento),0),2),0) AS ImporteGravado FROM PVVENTAPUBLICO INNER JOIN PVDETALLEVENTAPUBLICO ON PVDETALLEVENTAPUBLICO.IntCveVenta = PVVENTAPUBLICO.IntCveVenta WHERE (PVDETALLEVENTAPUBLICO.MnyIva = -1) AND PVVENTAPUBLICO.ChrFolioFactura = '" & rsFactura!chrfoliofactura & "'", "SELECT nvl(ROUND(NVL(SUM((PVDETALLEVENTAPUBLICO.MnyPrecio * PVDETALLEVENTAPUBLICO.IntCantidad) + PVDETALLEVENTAPUBLICO.MNYIEPS - PVDETALLEVENTAPUBLICO.MnyDescuento),0),2),0) AS ImporteGravado FROM PVVENTAPUBLICO INNER JOIN PVDETALLEVENTAPUBLICO ON PVDETALLEVENTAPUBLICO.IntCveVenta = PVVENTAPUBLICO.IntCveVenta WHERE (PVDETALLEVENTAPUBLICO.MnyIva > 0) AND PVVENTAPUBLICO.ChrFolioFactura = '" & rsFactura!chrfoliofactura & "'"), adLockOptimistic, adOpenDynamic)
            If rsCalcPvFactImp.RecordCount > 0 Then
                vldblimportegravado = rsCalcPvFactImp!ImporteGravado
            End If
            Set rsCalcPvFactImp = frsRegresaRs("SELECT nvl((ROUND(NVL(SUM(MnyCantidad),0),2)),0) AS ImporteConceptosSegudo FROM PVDETALLEFACTURA WHERE ChrFolioFactura = '" & rsFactura!chrfoliofactura & "' AND ChrTipo = 'OD'", adLockOptimistic, adOpenDynamic)
            If rsCalcPvFactImp.RecordCount > 0 Then
                vldblSumatoria = rsCalcPvFactImp!ImporteConceptosSegudo
            End If
            If rsFactura!BITPESOS = 1 Then
                pEjecutaSentencia "UPDATE PVFACTURAIMPORTE SET MNYSUBTOTALGRAVADO = " & vldblimportegravado - vldblSumatoria & ", MNYSUBTOTALNOGRAVADO = " & rsFactura!mnyTotalFactura - rsFactura!smyIVA - vldblimportegravado - vldblSumatoria & " WHERE IntConsecutivo = " & vllngConsecutivoFactura
            Else
                pEjecutaSentencia "UPDATE PVFACTURAIMPORTE SET MNYSUBTOTALGRAVADO = " & Round((vldblimportegravado - vldblSumatoria) / rsFactura!mnytipocambio, 2) & ", MNYSUBTOTALNOGRAVADO = " & rsFactura!mnyTotalFactura - rsFactura!smyIVA - Round((vldblimportegravado - vldblSumatoria) / rsFactura!mnytipocambio, 2) & " WHERE IntConsecutivo = " & vllngConsecutivoFactura
            End If
            rsFactura.Close
           
           ' Poner el número de factura en los cargo
           vlstrSentencia = "UPDATE PVCargo SET chrFolioFactura = '" & Trim(lstrFolioDocumentoF) & "' WHERE intNumCargo IN ("
           For a = 1 To grdArticulos.Rows - 1
               vlstrSentencia = vlstrSentencia & Trim(grdArticulos.TextMatrix(a, 15))
               vlstrSentencia = IIf(a = grdArticulos.Rows - 1, vlstrSentencia & ")", vlstrSentencia & ",")
           Next a
           pEjecutaSentencia (vlstrSentencia)
           Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "GENERACION DE FACTURA AUTOMATICA", lstrFolioDocumentoF)
           
           If intTipoEmisionComprobante = 2 Then
              If Not fblnValidaDatosCFDCFDi(vllngConsecutivoFactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacionDocumentoF), strNumeroAprobacionDocumentoF, False) Then
                 EntornoSIHO.ConeccionSIHO.Execute "ROLLBACK TO SAVEPOINT FACTURA" 'para atras la factura
                 blnFacturaAutomatica = False
                 ''quitamos el folio de la factura de la vental al publico
                 pEjecutaSentencia "Update pvventapublico set CHRFOLIOFACTURA = null , BITCANCELADO = 0, BITFACTURAAUTOMATICA=0 where INTCVEVENTA =" & vllngConsecutivoVenta
                 pCambioFacturaTicket Trim(vlstrFolioDocumento)
              End If
           End If
        End If
        
        '---------------
        ' Guardar información de la forma de pago en tabla intermedia para el ticket
        '---------------
        If vldblFacturaBaseTotal > 0 And vlblnOkFormasPago Then
           For a = 0 To UBound(aFormasPago(), 1)
               If Not aFormasPago(a).vlbolEsCredito Then 'Formas de pago distintas a Crédito
                  frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(a).vlintNumFormaPago & "|" & aFormasPago(a).lngIdBanco & "|" & _
                                      IIf(aFormasPago(a).vldblTipoCambio = 0, aFormasPago(a).vldblCantidad, aFormasPago(a).vldblDolares) & "|" & IIf(aFormasPago(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(a).vldblTipoCambio & "|" & _
                                      fstrTipoMovimientoForma(aFormasPago(a).vlintNumFormaPago, IIf(blnFacturaAutomatica, "F", "T")) & "|" & IIf(blnFacturaAutomatica, "FA", "TI") & "|" & IIf(blnFacturaAutomatica, vllngConsecutivoFactura, vllngConsecutivoVenta) & "|" & _
                                      vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                    
                  ' Agregado para caso 8741
                  ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                  vldblComisionIvaBancaria = 0
                  If aFormasPago(a).vllngCuentaComisionBancaria <> 0 And aFormasPago(a).vldblCantidadComisionBancaria <> 0 Then
                     
                     If aFormasPago(a).vldblTipoCambio = 0 Then
                        vldblComisionIvaBancaria = (aFormasPago(a).vldblCantidadComisionBancaria + aFormasPago(a).vldblIvaComisionBancaria) * -1
                     Else
                        vldblComisionIvaBancaria = (aFormasPago(a).vldblCantidadComisionBancaria + aFormasPago(a).vldblIvaComisionBancaria) / aFormasPago(a).vldblTipoCambio * -1
                     End If
                                          
                     frsEjecuta_SP vllngNumeroCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(a).vlintNumFormaPago & "|" & aFormasPago(a).lngIdBanco & "|" & vldblComisionIvaBancaria & "|" & IIf(aFormasPago(a).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(a).vldblTipoCambio & "|" & "CBA" & "|" & IIf(blnFacturaAutomatica, "FA", "TI") & "|" & IIf(blnFacturaAutomatica, vllngConsecutivoFactura, vllngConsecutivoVenta) & "|" & vllngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo, "Sp_PvInsMovimientoBancoForma"
                  End If
               End If
           Next a
        End If
        'guardamos el log de la transacción
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "VENTAS AL PUBLICO", CStr(vllngConsecutivoVenta))

        'insertamos movimientos en el corte y actualizamos los cortes si es que el corte de cerró mientras se hacia la venta(por si la moscas)
        vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte, True)
        If vllngCorteUsado <> 0 Then
            vllngCorteUsado = fRegistrarMovArregloCorte2(vllngCorteUsado, True)
        End If
        If vllngCorteUsado = 0 Then
           EntornoSIHO.ConeccionSIHO.RollbackTrans
           'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
           MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
           Exit Sub
        End If
        If vllngCorteUsado <> vllngNumeroCorte Then
           pEjecutaSentencia "UPdate pvventapublico set INTNUMCORTE = " & vllngCorteUsado & " where intcveventa = " & vllngConsecutivoVenta
           If blnFacturaAutomatica Then
              pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & vllngConsecutivoFactura
           End If
        End If
        
        '------------------------------------------------------------------------
        'Se registran los puntos acumulados y los puntos utilizados en la factura
        '------------------------------------------------------------------------
        pControlPuntos vllngConsecutivoVenta, vlstrFolioDocumento, vldtmFechaHoy, vldtmHoraHoy
        
        ' FIN de Transacción
        EntornoSIHO.ConeccionSIHO.CommitTrans
                      
        If blnFacturaAutomatica Then
           'TIMBRE FACTURA AUTOMATICA
           If intTipoEmisionComprobante = 2 Then
              pMovimientosTimbrado
              
              EntornoSIHO.ConeccionSIHO.BeginTrans
              If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", intTipoAgrupacion, CInt(strAnoAprobacionDocumentoF), strNumeroAprobacionDocumentoF, IIf(intTipoCFDFactura = 1, True, False)) Then
                 On Error Resume Next
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 If vgIntBanderaTImbradoPendiente = 1 Then

                 ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then
                    pEliminaPendientes
                    Exit Sub
                 End If
              Else
                 EntornoSIHO.ConeccionSIHO.CommitTrans
                 pEliminaPendientesTimbre vllngConsecutivoFactura, "FA"
              End If
              pBarraCFD
           End If
        End If
        ' Impresión del Ticket
         MsgBox SIHOMsg(420), vbInformation, "Mensaje"
        
        If vlstrFacturaTicket = "T" Then
            pAlistaImpresionTicket vllngConsecutivoVenta
        End If
        If Trim(txtMovimientoPaciente.Text) <> "" Then pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
    End If
    
    pLimpiaVariablesFacGlobal
    Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, ("frmPOS" & ":cmdSave_Click"))
    Unload Me
End Sub

Private Sub pLimpiaVariablesFacGlobal()
    vgStrPeriodicidad = ""
    vgStrMesesGlobal = ""
    vgStrAñoGlobal = ""
    
    pLimpia
End Sub

Private Sub pActualizaEmpleadoIngresa()
    'Actualizar el empleado que ingresa
    pEjecutaSentencia "UPDATE ExPacienteIngreso SET intCveEmpleadoIngreso = " & CStr(vllngPersonaGraba) & " WHERE intNumCuenta = " & Val(txtMovimientoPaciente.Text) & " AND intCveTipoIngreso IN (7, 8)"
End Sub

Private Sub pAlistaImpresionTicket(vllngConsecutivoVenta As Long)
    Dim rsReporte As New ADODB.Recordset
    Dim a As Long

    Set vlrsSelTicket = frsEjecuta_SP(Val(vllngConsecutivoVenta) & "|" & IIf(vlblnLicenciaIEPS, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0), "SP_PVSELTICKET")
    a = vlrsSelTicket.RecordCount
    If lblnImpresoraSerial Then
        pImpresionSerial vllngConsecutivoVenta
        vlrsSelTicket.Close
    Else 'Impresora paralela normal
        Set rsReporte = frsEjecuta_SP(Val(vllngConsecutivoVenta) & "|" & IIf(vlblnLicenciaIEPS, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0), "SP_PVSELTICKET")
        If rsReporte.RecordCount > 0 Then
            pInstanciaReporte vgrptReporte, "rptTicket.rpt"
            vgrptReporte.DiscardSavedData
            pCopiasTicket
            pCargaParametrosTicket
            pCargaParameterFields alstrParamTickets, vgrptReporte
            strCurrPrinter = fstrCurrPrinter
            fblnAsignaImpresora vgintNumeroDepartamento, "TI"
            pSetReportPrinterSettings vgrptReporte
    
            pImprimeReporte vgrptReporte, rsReporte, "I", "Ticket de pago", , vlintCopiasTicket
            pSetPrinter strCurrPrinter
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetros.
        End If
        If rsReporte.State <> adStateClosed Then rsReporte.Close
        vlrsSelTicket.Close
    End If
End Sub

Private Sub cmdGrabaControlAseguradora_Click()
    Dim dblTotal As Double
    Dim rsCarta As ADODB.Recordset
    Dim vllngCveCarta As Long
    
    If fblnControlValido() Then
    
        dblTotal = Val(Format(txtTotal.Text, ""))
    
        'Calcular deducible, coaseguro y copago:
        ldblDeducible = 0
        If Val(Format(txtDeducible.Text, "")) <> 0 Then
            ldblDeducible = IIf(optTipoDeducible(0).Value, Val(Format(txtDeducible.Text, "")), dblTotal * Val(Format(txtDeducible.Text, "")) / 100)
        End If
        
        ldblCoaseguro = 0
        If Val(Format(txtCoaseguro.Text, "")) <> 0 And (dblTotal - ldblDeducible) > 0 Then
            ldblCoaseguro = IIf(optTipoCoaseguro(0).Value, Val(Format(txtCoaseguro.Text, "")), (dblTotal - ldblDeducible) * Val(Format(txtCoaseguro.Text, "")) / 100)
        End If
        
        ldblCopago = 0
        If Val(Format(txtCopago.Text, "")) <> 0 And (dblTotal - ldblDeducible - ldblCoaseguro) > 0 Then
            ldblCopago = IIf(optCopagoCantidad.Value, Val(Format(txtCopago.Text, "")), (dblTotal - ldblDeducible - ldblCoaseguro) * Val(Format(txtCopago.Text, "")) / 100)
        End If
        
        If (ldblDeducible + ldblCoaseguro + ldblCopago) > dblTotal Then
            'La cantidad excede el importe a pagar.
            MsgBox SIHOMsg(930), vbOKOnly + vbExclamation, "Mensaje"
            pEnfocaTextBox txtDeducible
        Else
            ' CUENTA | TIPOPACIENTE | CVEEMPRESA | DEDUCIBLE | PORCENTAJECOASEGURO
            ' SUMAASEGURADA | NOMBREASEGURADO | PARENTESCO | FORMAFACTURACION | HONORARIOS
            ' COMENTARIOS | EXCEDENTESUMAASEGURADA | BITFACTURADEDUCIBLE | BITFACTURACOASEGURO
            ' BITFACTURACOPAGO | PORCENTAJECOPAGO | CANTIDADCOPAGO | CANTIDADCOASEGURO
            ' NUMPORCENTAJEDEDUCIBLE | CHRTIPODEDUCIBLE | CHRTIPOCOASEGURO | CHRTIPOCOPAGO
            
            ' MNYCANTIDADCOASEGUROADICION | NUMPORCENTAJECOASEGUROADICI | BITFACTURACOASEGUROADICIONA
            ' CHRTIPOCOASEGUROADICIONAL | MNYDESCUENTOEXCEDENTE | MNYDESCUENTODEDUCIBLE
            ' MNYDESCUENTOCOASEGURO | MNYDESCUENTOCOASEGUROADICIO | MNYDESCUENTOCOPAGO
            ' INTAUTORIZA | INTTIPOPOLIZA | INTNUMCONTROL | INTNUMPOLIZA

            ' BITFACTURACOASEGUROMEDICO | NUMPORCENTAJECOASEGUROMEDICO | MNYCANTIDADCOASEGUROMEDICO
            ' MNYDESCUENTOCOASEGUROMEDICO | MNYCANTIDADMAXIMACOASEGURO | CHRTIPOCOASEGUROMEDICO
            ' MNYHONORARIOSAFACTURAR | MNYCANTIDADCMAFACTURAR
                
            'PVCARTACONTROLSEGURO
    
            vlstrSentencia = "select * from PVCARTACONTROLSEGURO where intcveCarta=-1"
            Set rsCarta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                                              
            With rsCarta
                .AddNew
                !VCHDESCRIPCION = "CARTA DE AUTORIZACIÓN GENERAL"
                !intNumCuenta = CLng(txtMovimientoPaciente.Text)
                !intcveempresa = vgintEmpresa
                !BITDEFAULT = 1
                .Update
                
                vllngCveCarta = flngObtieneIdentity(UCase("SEQ_PVCARTACONTROLSEGURO"), !intCveCarta)
            End With
            
            vgstrParametrosSP = _
            txtMovimientoPaciente.Text _
            & "|" & "E" _
            & "|" & Trim(str(vgintEmpresa)) _
            & "|" & str(ldblDeducible) _
            & "|" & str(IIf(optTipoCoaseguro(1).Value, Val(Format(txtCoaseguro.Text, "")), 0)) _
            & "|" & "0" _
            & "|" & Trim(lblPaciente.Caption) _
            & "|" & "Mismo" _
            & "|" & "FFS" _
            & "|" & "0" _
            & "|" & " " _
            & "|" & "0" _
            & "|" & IIf(chkFacturarDeducible.Value = 1, "1", "0") _
            & "|" & IIf(chkFacturarCoaseguro.Value = 1, "1", "0") _
            & "|" & IIf(chkFacturarCopago.Value = 1, "1", "0") _
            & "|" & str(IIf(optControlCopagoPorciento.Value, Val(Format(txtCopago.Text, "")), 0)) _
            & "|" & str(ldblCopago) _
            & "|" & str(ldblCoaseguro) _
            & "|" & str(IIf(optTipoDeducible(1).Value, Val(Format(txtDeducible.Text, "")), 0)) _
            & "|" & IIf(optTipoDeducible(0).Value, "C", "P") _
            & "|" & IIf(optTipoCoaseguro(0).Value, "C", "P") _
            & "|" & IIf(optCopagoCantidad.Value, "C", "P")
        
            vgstrParametrosSP = vgstrParametrosSP _
            & "|0|0|0|C|0|0|0|0|0|0|0|0|0|0|0|0|0|0|C|0|0|" & vllngCveCarta & "|" & Null & "|" & Null & "|" & Null & "|" & Null

            vgstrParametrosSP = vgstrParametrosSP & ""
            
            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSCONTROLASEGURADORA"
        
            'La información se actualizó satisfactoriamente.
            MsgBox SIHOMsg(284), vbInformation, "Mensaje"
            
            lblnGuardoControl = True
            
            If ldblDeducible = 0 And ldblCopago = 0 And ldblCoaseguro = 0 Then
                lblnExisteControl = False
            Else
                lblnExisteControl = True
            End If
            
            
            freControlAseguradora.Visible = False
            
            FreDetalle.Enabled = True
            freGraba.Enabled = True
            freDescuentos.Enabled = True
            freConsultaPrecios.Enabled = True
            freFacturaPesosDolares.Enabled = True
        End If
    End If
    
End Sub

Private Function fblnDatosValidos() As Boolean
    Dim vlstrSentencia  As String
    Dim rsTemp As New ADODB.Recordset
    Dim vPrinter As Printer

    fblnDatosValidos = True

    If grdArticulos.RowData(1) = -1 Then
        fblnDatosValidos = False
        MsgBox SIHOMsg(372), vbCritical, "Mensaje" 'No existen conceptos para facturar.
    End If
    
    If vlblnLicenciaIEPS Then
        If fblnDatosValidos And (glngctaIEPSCobrado = 0 Or glngctaIEPSNoCobrado = 0) Then
            fblnDatosValidos = False
            MsgBox SIHOMsg(1258), vbCritical, "Mensaje" 'No se encuentran registradas las cuentas de IEPS cobrado y no cobrado en los parámetros generales del sistema.
        End If
        
        
        If fblnDatosValidos Then
                vlstrSentencia = "Select BITESTATUSMOVIMIENTOS from cncuenta where intnumerocuenta = " & glngctaIEPSCobrado
                Set rsTemp = frsRegresaRs(vlstrSentencia, adLockOptimistic)
                
                If rsTemp.RecordCount = 0 Then
                   fblnDatosValidos = False
                   MsgBox SIHOMsg(1258), vbCritical, "Mensaje" 'No se encuentran registradas las cuentas de IEPS cobrado y no cobrado en los parámetros generales del sistema.
                Else
                    If rsTemp!Bitestatusmovimientos = 0 Then
                       fblnDatosValidos = False
                       '¡La cuenta contable para el manejo IEPS cobrado no acepta movimientos!
                       MsgBox SIHOMsg(1259), vbCritical, "Mensaje"
                    End If
                End If
        End If
          
        If fblnDatosValidos Then
                vlstrSentencia = "Select BITESTATUSMOVIMIENTOS from cncuenta where intnumerocuenta = " & glngctaIEPSNoCobrado
                Set rsTemp = frsRegresaRs(vlstrSentencia, adLockOptimistic)
                
                If rsTemp.RecordCount = 0 Then
                   fblnDatosValidos = False
                   MsgBox SIHOMsg(1258), vbCritical, "Mensaje" 'No se encuentran registradas las cuentas de IEPS cobrado y no cobrado en los parámetros generales del sistema.
                Else
                    If rsTemp!Bitestatusmovimientos = 0 Then
                       fblnDatosValidos = False
                       '¡La cuenta contable para el manejo IEPS cobrado no acepta movimientos!
                       MsgBox SIHOMsg(1260), vbCritical, "Mensaje"
                    End If
                End If
        End If
    End If
    
    If fblnDatosValidos And (glngCtaIVACobrado = 0 Or glngCtaIVANoCobrado = 0) Then
       fblnDatosValidos = False
       MsgBox SIHOMsg(729), vbCritical, "Mensaje" 'No se encuentran registradas las cuentas de IVA cobrado y no cobrado en los parámetros generales del sistema.
    End If
    
    If fblnDatosValidos Then 'Que tengan configurada una impresora:
        vlstrSentencia = "select chrNombreImpresora Impresora from ImpresoraDepartamento where chrTipo = 'FA' and smiCveDepartamento = " & Trim(str(vgintNumeroDepartamento))
        Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsTemp.RecordCount > 0 Then
            For Each vPrinter In Printers
                If UCase(Trim(vPrinter.DeviceName)) = UCase(Trim(rsTemp!Impresora)) Then
                     Set Printer = vPrinter
                End If
            Next
        Else
            fblnDatosValidos = False
            MsgBox SIHOMsg(492), vbCritical, "Mensaje"  'No se tiene asignada una impresora en la cual imprimir las facturas
        End If
        rsTemp.Close
    End If
    
    If fblnDatosValidos Then
        '--------------------------------------------------------
        ' Tipo de cambio de los DOLARES
        '--------------------------------------------------------
        vldblTipoCambio = fdblTipoCambio(fdtmServerFecha, "V")
        If vldblTipoCambio = 0 Then
            fblnDatosValidos = False
            MsgBox SIHOMsg(231), vbCritical, "Mensaje" 'No está registrado el tipo de cambio del día.
        End If
    End If
    'Que no se hayan agregado más cargos, cuando ya se calculó el deducible, coaseguro y/o copago:
    If fblnDatosValidos And lblnExisteControl And Not lblnGuardoControl Then
        fblnDatosValidos = False
        'Actualice el registro del control de aseguradora, la cuenta ha cambiado.
        MsgBox SIHOMsg(931), vbOKOnly + vbInformation, "Mensaje"
    End If

End Function

Private Sub cmdImportarVenta_Click()
    pImportarVenta
End Sub

Private Sub cmdImprimeTicketSinFacturar_Click()
    Dim rsVentaPublico As New ADODB.Recordset    'RS para guardar la Venta, Ticket
    Dim vldtmFechaHoy As Date                    'Varible con la Fecha actual
    Dim vldtmHoraHoy As Date                     'Varible con la Hora actual
    Dim vlintContador As Integer                 'Para los ciclos
    Dim vldblTotSubtotal As Double               'Subtotal de la cuenta, para el ticket
    Dim vldblTotIva As Double                    'IVA total de la cuenta, para el ticket
    Dim vldblTotDescuentos As Double             'Descuento de la cuenta, para el ticket
    Dim vllngConsecutivoVenta As Long            'Todas las ventas se van a guardar y los cargos y pagos llevarán este consecutivo
    Dim rsDetalleVentaPublico As New ADODB.Recordset    'RS para guardar del detalle de la Venta, Ticket
    Dim vllngAux As Long
    Dim vlrsSelTicketSinFact As New ADODB.Recordset     'Para eliminar el uso del comando cmdSelTicket
    Dim alstrParam(9) As String
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrNombreHospital As String            'Para la impresión del ticket
    Dim vlstrRegistro As String                  'Para la impresión del ticket
    Dim vlstrDireccionHospital As String         'Para la impresión del ticket
    Dim vlstrTelefonoHospital As String          'Para la impresión del ticket
    Dim strCurrPrinter As String
    
    On Error GoTo NotificaError
    
    If lblnExisteControl Then Exit Sub

    If Val(Format(txtTotal.Text, "")) <= 0 Then
        MsgBox SIHOMsg(1425), vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    If Not fblnDatosValidos() Then Exit Sub
       
    '------------------'
    ' Inicializa las variables de cliente y numero de referencia '
    pInicializaVariablesGuardado
    
    '----------------'
    ' Persona que graba
    '----------------'
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
          
    '----------------'
    ' Fecha y hora del Sistema
    '----------------'
    vldtmFechaHoy = fdtmServerFecha
    vldtmHoraHoy = fdtmServerHora
    vldblTotSubtotal = 0
    vldblTotIva = 0
    vldblTotDescuentos = 0
    vldblTotIEPS = 0
              
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    '----------------'
    ' Generar Registro de la Venta
    '----------------'
    If vlblnConsultaTicketPrevio Then
        Set rsVentaPublico = frsRegresaRs("SELECT * FROM PvVentaPublicoNoFacturado WHERE intCveVenta = " & Val(txtTicketPrevio.Text), adLockOptimistic, adOpenDynamic)
        With rsVentaPublico
        For vlintContador = 1 To grdArticulos.Rows - 1
            With grdArticulos
                If CDbl(grdArticulos.TextMatrix(vlintContador, 24)) <> 0 Then
                    vldblTotSubtotal = vldblTotSubtotal + Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), ""))), 2)
                Else
                    vldblTotSubtotal = vldblTotSubtotal + Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), "")))
                End If
                vldblTotIva = vldblTotIva + Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2)
                vldblTotDescuentos = vldblTotDescuentos + Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))
                vldblTotIEPS = vldblTotIEPS + Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
            End With
        Next
        
        !MNYSUBTOTAL = vldblTotSubtotal
        !MNYIVA = vldblTotIva
        !MNYDESCUENTO = vldblTotDescuentos
        !INTMOVPACIENTE = Val(txtMovimientoPaciente)
        !intCveMedico = IIf(cboMedico.ListIndex = -1, 0, cboMedico.ItemData(cboMedico.ListIndex))
        !mnyIeps = vldblTotIEPS
        !intPersonaReImprime = vllngPersonaGraba
        .Update
        End With
        vllngConsecutivoVenta = Val(txtTicketPrevio.Text)
    Else
        Set rsVentaPublico = frsRegresaRs("SELECT * FROM PvVentaPublicoNoFacturado WHERE intCveVenta = -1", adLockOptimistic, adOpenDynamic)
        With rsVentaPublico
            .AddNew
            !dtmFechahora = vldtmFechaHoy + vldtmHoraHoy
            !intCveEmpleado = IIf(optEmpleado.Value, Val(txtMovimientoPaciente), 0)
            !intCveDepartamento = vgintNumeroDepartamento
            !chrCliente = Trim(lblPaciente.Caption)
                   
            For vlintContador = 1 To grdArticulos.Rows - 1
                With grdArticulos
                    If CDbl(grdArticulos.TextMatrix(vlintContador, 24)) <> 0 Then
                        vldblTotSubtotal = vldblTotSubtotal + Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), ""))), 2)
                    Else
                        vldblTotSubtotal = vldblTotSubtotal + Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "") * Val(Format(grdArticulos.TextMatrix(vlintContador, 3), "")))
                    End If
                                    
                    vldblTotIva = vldblTotIva + Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2)
                    vldblTotDescuentos = vldblTotDescuentos + Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))
                    vldblTotIEPS = vldblTotIEPS + Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
                End With
            Next
            
            !MNYSUBTOTAL = vldblTotSubtotal
            !MNYIVA = vldblTotIva
            !MNYDESCUENTO = vldblTotDescuentos
            !bitcancelado = 0
            !INTMOVPACIENTE = IIf(optPaciente.Value, Val(txtMovimientoPaciente), 0)
            !intCveMedico = IIf(optMedico.Value, Val(txtMovimientoPaciente), 0)
            !mnyIeps = vldblTotIEPS
            !intPersonaReImprime = vllngPersonaGraba
            .Update
            vllngConsecutivoVenta = flngObtieneIdentity("SEC_PVVENTAPUBLICONOFACTURADO", !INTCVEVENTA)
        End With
    End If
    
    '----------------'
    ' Generar detalle de la venta
    '----------------'
    If vlblnConsultaTicketPrevio Then
        pEjecutaSentencia "DELETE FROM PvDetalleVentaPublicoNoFact WHERE intCveVenta = " & Val(txtTicketPrevio.Text)
    End If
    Set rsDetalleVentaPublico = frsRegresaRs("SELECT * FROM PvDetalleVentaPublicoNoFact WHERE intCveVenta = -1", adLockOptimistic, adOpenDynamic)
    For vlintContador = 1 To grdArticulos.Rows - 1
        '----------------'
        ' Generar un registro de DetalleVentaPublico '
        '----------------'
        With rsDetalleVentaPublico
            .AddNew
            !INTCVEVENTA = vllngConsecutivoVenta
            !intCveCargo = grdArticulos.RowData(vlintContador)
            !chrTipoCargo = grdArticulos.TextMatrix(vlintContador, 12)
            !intCantidad = grdArticulos.TextMatrix(vlintContador, 3)
            If CDbl(grdArticulos.TextMatrix(vlintContador, 24)) <> 0 Then
                !mnyPrecio = Val(Format(grdArticulos.TextMatrix(vlintContador, 24), ""))
            Else
                !mnyPrecio = Val(Format(grdArticulos.TextMatrix(vlintContador, 2), ""))
            End If
            !MNYDESCUENTO = Val(Format(grdArticulos.TextMatrix(vlintContador, 5), ""))
            !MNYIVA = Val(Format(grdArticulos.TextMatrix(vlintContador, 11), ""))
            !smiCveConceptoFacturacion = grdArticulos.TextMatrix(vlintContador, 13)
            !mnyIeps = Val(Format(grdArticulos.TextMatrix(vlintContador, 6), ""))
            !numporcentajeieps = Val(grdArticulos.TextMatrix(vlintContador, 19)) * 100
            !IntModoDescuentoInventario = grdArticulos.TextMatrix(vlintContador, 14)
            !intAuxExclusionDescuento = grdArticulos.TextMatrix(vlintContador, 20)
            !bitPrecioManual = Val((grdArticulos.TextMatrix(vlintContador, 23)))
            .Update
        End With
    Next

    rsDetalleVentaPublico.Close
    
    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "VENTAS AL PÚBLICO ANTES DE LA VENTA", CStr(vllngConsecutivoVenta))
    EntornoSIHO.ConeccionSIHO.CommitTrans
    
    ' Impresión del Ticket
    ' MsgBox SIHOMsg(420), vbInformation, "Mensaje"

    ' Traer datos generales del Hospital
    vlstrNombreHospital = Trim(vgstrNombreHospitalCH)
    vlstrRegistro = "R.SSA " & Trim(vgstrSSACH) & " RFC " & Trim(vgstrRfCCH)
    vlstrDireccionHospital = Trim(vgstrDireccionCH) & " CP. " & Trim(vgstrCodPostalCH)
    vlstrTelefonoHospital = Trim(vgstrTelefonoCH)
    
    Set vlrsSelTicketSinFact = frsEjecuta_SP(Val(vllngConsecutivoVenta) & "|" & IIf(vlblnLicenciaIEPS, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0), "SP_PVSELTICKETNOFACTURADO")
    'a = vlrsSelTicketSinFact.RecordCount
    If lblnImpresoraSerial Then
        'pImpresionSerial vllngConsecutivoVenta
        vlrsSelTicketSinFact.Close
    Else 'Impresora paralela normal
        Set rsReporte = frsEjecuta_SP(Val(vllngConsecutivoVenta) & "|" & IIf(vlblnLicenciaIEPS, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0), "SP_PVSELTICKETNOFACTURADO")
        If rsReporte.RecordCount > 0 Then
            pInstanciaReporte vgrptReporte, "rptTicketNoFacturado.rpt"
            
            vgrptReporte.DiscardSavedData

            alstrParam(0) = "Direccion;" & vlstrDireccionHospital
            alstrParam(1) = "Empleado;"
            If vlrsSelTicketSinFact.RecordCount > 0 Then
                alstrParam(1) = "Empleado;" & Trim(vlrsSelTicketSinFact!Empleado)
            End If
            alstrParam(2) = "FechaActual;" & UCase(Format(fdtmServerFecha, "dd/mmm/yyyy"))
            alstrParam(3) = "FolioVenta;" & "FOLIO  " & Trim(vllngConsecutivoVenta)
            alstrParam(4) = "HoraActual;" & UCase(Format(fdtmServerHora, "HH:mm"))
            If Trim(txtMovimientoPaciente.Text) = "" Or Val(txtMovimientoPaciente.Text) = 0 Then
                alstrParam(5) = "NombreCliente;" & "Público"
            Else
                alstrParam(5) = "NombreCliente;" & Trim(lblPaciente.Caption) & " / " & Trim(lblEmpresa.Caption)
            End If
            alstrParam(6) = "NombreEmpresa;" & Trim(vlstrNombreHospital)
            alstrParam(7) = "Registro;" & vlstrRegistro
            alstrParam(8) = "Telefono;" & "TEL. " & Format(RTrim(vlstrTelefonoHospital), "###-##-##")
            alstrParam(9) = "LeyendaInformacionCliente;" & Trim(lstrLeyendaCliente)

            pCargaParameterFields alstrParam, vgrptReporte
            strCurrPrinter = fstrCurrPrinter
            fblnAsignaImpresora vgintNumeroDepartamento, "TI"
            pSetReportPrinterSettings vgrptReporte
            pImprimeReporte vgrptReporte, rsReporte, "I", "Ticket de pago"
            pSetPrinter strCurrPrinter
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje" 'No existe información con esos parámetros.
        End If
        If rsReporte.State <> adStateClosed Then rsReporte.Close
        vlrsSelTicketSinFact.Close
    End If
    pNuevo
    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
    If txtClaveArticulo.Enabled Then
        txtClaveArticulo.SetFocus
    End If
Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, ("frmPOS" & ":cmdImprimeTicketSinFacturar_Click"))
    Unload Me
End Sub

Private Sub cmdInvertir_Click()
    Dim intcontador As Long

    For intcontador = 0 To lstFoliosTickets.ListCount - 1
        If lstFoliosTickets.Selected(intcontador) Then
            lstFoliosTickets.Selected(intcontador) = False
        Else
            lstFoliosTickets.Selected(intcontador) = True
        End If
    Next intcontador
End Sub

Private Sub pNuevoDetalleVenta(rs As Recordset, lngConsecutivoVenta As Long, intcontador As Integer)
        rs.AddNew
        rs!INTCVEVENTA = lngConsecutivoVenta
        rs!intCveCargo = grdArticulos.RowData(intcontador)
        rs!chrTipoCargo = grdArticulos.TextMatrix(intcontador, 12)
        rs!intCantidad = grdArticulos.TextMatrix(intcontador, 3)
        rs!mnyPrecio = IIf(Val(Format(grdArticulos.TextMatrix(intcontador, 24), "")) <> 0, Val(Format(grdArticulos.TextMatrix(intcontador, 24), "")), Val(Format(grdArticulos.TextMatrix(intcontador, 2), "")))
        rs!MNYDESCUENTO = IIf(Val(Format(grdArticulos.TextMatrix(intcontador, 25), "")) <> 0, Val(Format(grdArticulos.TextMatrix(intcontador, 25), "")), Val(Format(grdArticulos.TextMatrix(intcontador, 5), "")))
        rs!MNYIVA = IIf(Val(Format(grdArticulos.TextMatrix(intcontador, 26), "")) <> 0, Val(Format(grdArticulos.TextMatrix(intcontador, 26), "")), Val(Format(grdArticulos.TextMatrix(intcontador, 11), "")))
        rs!IntNumCargo = grdArticulos.TextMatrix(intcontador, 15)
        rs!smiCveConceptoFacturacion = grdArticulos.TextMatrix(intcontador, 13)
        rs!mnyIeps = Val(Format(grdArticulos.TextMatrix(intcontador, 6), ""))
        rs!numporcentajeieps = Val(grdArticulos.TextMatrix(intcontador, 19)) * 100
        rs!bitPrecioManual = Val((grdArticulos.TextMatrix(intcontador, 23)))
End Sub

Private Sub pCargaParametrosTicket()
    alstrParamTickets(0) = "Direccion;" & Trim(vgstrDireccionCH) & " CP. " & Trim(vgstrCodPostalCH)
    alstrParamTickets(1) = "Empleado;"
    If vlrsSelTicket.RecordCount > 0 Then
        alstrParamTickets(1) = "Empleado;" & Trim(vlrsSelTicket!Empleado)
    End If
    alstrParamTickets(2) = "FechaActual;" & UCase(Format(fdtmServerFecha, "dd/mmm/yyyy"))
    alstrParamTickets(3) = "FolioVenta;" & "FOLIO: " & Trim(vlstrFolioDocumento)
    alstrParamTickets(4) = "HoraActual;" & UCase(Format(fdtmServerHora, "HH:mm"))
    alstrParamTickets(5) = "NombreCliente;" & IIf(Trim(txtMovimientoPaciente.Text) = "", "Público", Trim(lblPaciente.Caption) & " / " & Trim(lblEmpresa.Caption))
    alstrParamTickets(6) = "NombreEmpresa;" & Trim(vgstrNombreHospitalCH)
    alstrParamTickets(7) = "Registro;" & "R.SSA " & Trim(vgstrSSACH) & " RFC " & Trim(vgstrRfCCH)
    alstrParamTickets(8) = "Telefono;" & "TEL. " & Format(RTrim(Trim(vgstrTelefonoCH)), "###-##-##")
    alstrParamTickets(9) = "LeyendaInformacionCliente;" & Trim(lstrLeyendaCliente)
End Sub

Private Sub pEliminaPendientes()
    pEliminaPendientesTimbre vllngConsecutivoFactura, "FA"
    pCancelarFactura Trim(lstrFolioDocumentoF), vllngPersonaGraba, "frmFacturacion", False, False 'cancelamos factura
    pQuitaBarraCFD
End Sub

Private Sub pMovimientosTimbrado()
    pBarraCFD2
    pLogTimbrado 2
    pMarcarPendienteTimbre vllngConsecutivoFactura, "FA", vgintNumeroDepartamento
End Sub

Private Sub pLimpia()
    pNuevo
    pEnfocaTextBox txtClaveArticulo
    cboMedico.ListIndex = 0
    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
End Sub

Private Sub pCopiasTicket()
Dim rsCopiasTicket As New ADODB.Recordset

    '|  Obtiene el número de copias que se imprimirán del ticket
    Set rsCopiasTicket = frsRegresaRs("Select SiParametro.VCHVALOR From SiParametro Where SiParametro.intCveEmpresaContable = " & vgintClaveEmpresaContable & " And SiParametro.VCHNOMBRE = 'NUMCOPIASTICKET'", adLockReadOnly, adOpenForwardOnly)
    If rsCopiasTicket.RecordCount > 0 Then
        vlintCopiasTicket = rsCopiasTicket!VCHVALOR
    Else
        vlintCopiasTicket = 1
    End If
End Sub

Private Sub pDatosFiscales()
    frmDatosFiscales.vgstrTipoUsoCFDI = IIf(vgintEmpresa > 0, "EM", "TP")
    frmDatosFiscales.vgintTipoPacEmp = IIf(vgintEmpresa > 0, vgintEmpresa, vgintTipoPaciente)
    frmDatosFiscales.Show vbModal
    vlstrNombreFacturaPaciente = frmDatosFiscales.vgstrNombre
    vlstrDireccionFacturaPaciente = frmDatosFiscales.vgstrDireccion
    vlstrNumeroExteriorFacturaPaciente = frmDatosFiscales.vgstrNumExterior
    vlstrNumeroInteriorFacturaPaciente = frmDatosFiscales.vgstrNumInterior
    vlBitExtranjeroFacturaPaciente = frmDatosFiscales.vgBitExtranjero
    vlstrColoniaFacturaPaciente = frmDatosFiscales.vgstrColonia
    vlstrCPFacturaPaciente = frmDatosFiscales.vgstrCP
    
    vgstrRegimenFiscal = frmDatosFiscales.vlstrRegimenFiscal
    
    vllngCveCiudadFacturaPaciente = frmDatosFiscales.llngCveCiudad
    vlstrTelefonoFacturaPaciente = frmDatosFiscales.vgstrTelefono
    vlstrRFCFacturaPaciente = frmDatosFiscales.vgstrRFC
    vlintUsoCFDI = frmDatosFiscales.vgintUsoCFDI
    
End Sub

Private Sub pDatosFiscales2()
    If vlblnImportoVenta = True Then
        frmDatosFiscales.vgstrTipoUsoCFDI = IIf(vgintEmpresa > 0, "EM", "TP")
        frmDatosFiscales.vgintTipoPacEmp = IIf(vgintEmpresa > 0, vgintEmpresa, vgintTipoPaciente)
        vlstrNombreFactura = IIf(vgstrVersionCFDI = "3.3", frmDatosFiscales.vgstrNombre, "PUBLICO EN GENERAL")
        vlstrDireccion = frmDatosFiscales.vgstrDireccion
        vlstrNumeroExterior = frmDatosFiscales.vgstrNumExterior
        vlstrNumeroInterior = frmDatosFiscales.vgstrNumInterior
        vlBitExtranjero = frmDatosFiscales.vgBitExtranjero
        vlstrColonia = frmDatosFiscales.vgstrColonia
        vlstrCP = frmDatosFiscales.vgstrCP
        llngCveCiudad = frmDatosFiscales.cboCiudad.ItemData(frmDatosFiscales.cboCiudad.ListIndex)
        vlstrTelefono = frmDatosFiscales.vgstrTelefono
        vlstrRFC = "XAXX010101000"
        vlstrNumRef = frmDatosFiscales.vlstrNumRef
        vgstrRegimenFiscal = frmDatosFiscales.vlstrRegimenFiscal
        vlstrTipo = frmDatosFiscales.vlstrTipo
        vlintUsoCFDI = 64
    Else
        frmDatosFiscales.vgstrTipoUsoCFDI = IIf(vgintEmpresa > 0, "EM", "TP")
        frmDatosFiscales.vgintTipoPacEmp = IIf(vgintEmpresa > 0, vgintEmpresa, vgintTipoPaciente)
        frmDatosFiscales.Show vbModal
        vlstrNombreFactura = frmDatosFiscales.vgstrNombre
        vlstrDireccion = frmDatosFiscales.vgstrDireccion
        vlstrNumeroExterior = frmDatosFiscales.vgstrNumExterior
        vgstrRegimenFiscal = frmDatosFiscales.vlstrRegimenFiscal
        vlstrNumeroInterior = frmDatosFiscales.vgstrNumInterior
        vlBitExtranjero = frmDatosFiscales.vgBitExtranjero
        vlstrColonia = frmDatosFiscales.vgstrColonia
        vlstrCP = frmDatosFiscales.vgstrCP
        llngCveCiudad = frmDatosFiscales.llngCveCiudad
        vlstrTelefono = frmDatosFiscales.vgstrTelefono
        vlstrRFC = frmDatosFiscales.vgstrRFC
        vlstrNumRef = frmDatosFiscales.vlstrNumRef
        vlstrTipo = frmDatosFiscales.vlstrTipo
        vlintUsoCFDI = frmDatosFiscales.vgintUsoCFDI
    End If
    
End Sub

Private Sub pDatosFiscales3()
'datos fiscales cuando se tiene convenio
    frmDatosFiscales.vgstrNombre = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNombreFactura))
    frmDatosFiscales.vgstrDireccion = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrDireccionFactura))
    frmDatosFiscales.vgstrNumExterior = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNumeroExteriorFactura))
    frmDatosFiscales.vgstrNumInterior = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrNumeroInteriorFactura))
    frmDatosFiscales.vgstrColonia = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrColoniaFactura))
    frmDatosFiscales.vgstrCP = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrCPFactura))
    frmDatosFiscales.cboCiudad.ListIndex = IIf(Trim(txtMovimientoPaciente.Text) = "", -1, flngLocalizaCbo(frmDatosFiscales.cboCiudad, str(llngCveCiudadPaciente)))
    frmDatosFiscales.vgstrTelefono = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrTelefonoFactura))
    frmDatosFiscales.vgstrRFC = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrRFCFactura))
    
    frmDatosFiscales.vlstrRegimenFiscal = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrRegimenFiscal))
    
    
    
    frmDatosFiscales.vlstrNumRef = CStr(vglngCveExtra)
    frmDatosFiscales.vlstrTipo = vgstrTipoPaciente
    frmDatosFiscales.vgActivaSujetoaIEPS = IIf(vldblFacturaPacienteTotal > 0, False, True)
    frmDatosFiscales.vgBitSujetoaIEPS = 0
    frmDatosFiscales.vgstrCorreo = IIf(Trim(txtMovimientoPaciente.Text) = "", "", Trim(vgstrEmailCH))       'JASM 20211227
    frmDatosFiscales.sstDatos.Tab = 0
End Sub

Private Sub pImpresionSerial(vllngConsecVenta As Long)
    'Impresión en la impresora serial con cajón de dinero
    pAbrirCOM1
    pAbrirCaja
    pActivarSonidoCaja
    pImprimeTicket CStr(vllngConsecVenta)
    pCerrarCOM1
End Sub

Private Sub pPreparaGridParaDatosFactura()
    ' Prepara el Grid escondido para guardar los datos de la factura
    With grdFactura
         .Clear
         .ClearStructure
         .Cols = 9
         .Rows = 2
         .RowData(1) = -1
         .ColWidth(1) = 0 'Concepto de facturación (texto)
         .ColWidth(2) = 0 'Cargo
         .ColWidth(3) = 0 'Abono
         .ColWidth(4) = 0 'IVA
         .ColWidth(5) = 0 'Descuentos
         .ColWidth(6) = 0 'Descuentos
         .ColWidth(7) = 0 'Descuentos
         .ColWidth(8) = 0 'Folio del ticket para las facturas globales
         .Row = 1
    End With
End Sub


Private Sub pQuitaBarraCFD()
    Screen.MousePointer = vbDefault
    freBarraCFD.Visible = False
    frmPOS.Enabled = True
End Sub

Private Sub pBarraCFD()
    pLogTimbrado 1
    'Barra de progreso CFD
    pgbBarraCFD.Value = 100
    freBarraCFD.Top = 3200
    pQuitaBarraCFD
End Sub

Private Sub pBarraCFD2()
    pgbBarraCFD.Value = 70
    freBarraCFD.Top = 3200
    Screen.MousePointer = vbHourglass
    lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere..."
    freBarraCFD.Visible = True
    freBarraCFD.Refresh
    frmPOS.Enabled = False
End Sub

Private Sub pBarraCFD3()
    pgbBarraCFD.Value = 70
    freBarraCFD.Top = 3200
    Screen.MousePointer = vbHourglass
    lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
    freBarraCFD.Visible = True
    freBarraCFD.Refresh
    frmPOS.Enabled = False
End Sub

Private Sub pBarraCFD4()
    pgbBarraCFD.Value = 100
    freBarraCFD.Top = 3200
    pQuitaBarraCFD
End Sub

Private Sub pInicializaVariablesGuardado()
    Dim vlintContador As Integer

    If vlblnImportoVenta = True And vgstrVersionCFDI = "4.0" Then
        'Rellenar variables para el nodo InformacionGlobal del CFDi
        pRellenaInfoGlobal
    End If

    vldblsubtotalgravado = 0
    vldblsubtotalNogravado = 0
    vldbldescuentogravado = 0
    vldblDescuentoNoGravado = 0
    vlstrNumRef = 0
    vlstrTipo = 0
    vldblFacturaPacienteSubTotal = 0
    vldblFacturaPacienteIVA = 0
    vldblFacturaPacienteTotal = 0
    blnFacturaAutomatica = False
    vlbytNumeroConceptosPaciente = 0
    ReDim aFPFacturaPaciente(0)
    ReDim aTickets(0)
    blnFacturarConceptosSeguro = False
    vldblTotalControlAseguradorasSinIVA = 0
    vldblTotIEPS = 0
    vlstrFolioDocumentoPaciente = ""
    vlngCveFormato = 0
    lngCveFormato = 0
    
    vldblTotalIVAGuardado = 0
    vldblSubtotalGuardado = 0
    blnFacturaDeducible = chkFacturarDeducible.Value = 1
    blnFacturaCoaseguro = chkFacturarCoaseguro.Value = 1
    blnFacturaCopago = chkFacturarCopago.Value = 1
    
    If lblnExisteControl Then
        For vlintContador = 1 To grdArticulos.Rows - 1
            vldblSubtotalGuardado = vldblSubtotalGuardado + Val(Format(grdArticulos.TextMatrix(vlintContador, 4), "#########.##"))
            vldblTotalIVAGuardado = vldblTotalIVAGuardado + Val(Format(grdArticulos.TextMatrix(vlintContador, 8), "#########.##"))
        Next vlintContador
        
        vldblFacturaPacienteSubTotal = IIf(blnFacturaDeducible, ldblDeducible, 0) + IIf(blnFacturaCoaseguro, ldblCoaseguro, 0) + IIf(blnFacturaCopago, ldblCopago, 0)
        blnFacturarConceptosSeguro = IIf(vldblFacturaPacienteSubTotal > 0, True, False)
        vldblTotalControlAseguradoras = ldblDeducible + ldblCoaseguro + ldblCopago
        vldblFacturaPacienteIVA = (vldblTotalIVAGuardado * (vldblTotalControlAseguradoras / (vldblSubtotalGuardado + vldblTotalIVAGuardado)))
        vldblFacturaPacienteTotal = vldblFacturaPacienteSubTotal
        vlbytNumeroConceptosPaciente = IIf(ldblDeducible > 0, 1, 0) + IIf(ldblCoaseguro > 0, 1, 0) + IIf(ldblCopago > 0, 1, 0)
    End If
End Sub

Private Sub pDatosFiscalesPvParametros()

    'asignacion de variables para mostrar empresas y otros pacientes que no tengan convenio
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If vlblnImportoVenta = True Then
        vlstrSentencia = "SELECT CHRRFCPOS RFC, null Clave, CHRNOMBREFACTURAPOS Nombre, CHRDIRECCIONPOS chrCalle," & _
                         "VCHNUMEROEXTERIORPOS vchNumeroExterior, VCHNUMEROINTERIORPOS vchNumeroInterior, " & _
                         "null Telefono,'OT' Tipo,  INTCVECIUDAD IdCiudad, VCHCOLONIAPOS Colonia, VCHCODIGOPOSTALPOS CP FROM PVParametro " & _
                         "WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic)
        
        If rs.RecordCount > 0 Then
            With frmDatosFiscales
                .vgstrNombre = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
                .vgstrDireccion = IIf(IsNull(rs!chrCalle), "", Trim(rs!chrCalle))
                .vgstrNumExterior = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                .vgstrNumInterior = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                .vgstrColonia = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
                .vgstrCP = IIf(IsNull(rs!CP), "", Trim(rs!CP))
                .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
                .vgstrTelefono = ""
                .vgstrRFC = "XAXX010101000"
                .vlstrNumRef = IIf(vlstrTipo = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
                .vlstrTipo = IIf(IsNull(rs!tipo), "OT", rs!tipo)
                .vgActivaSujetoaIEPS = False
                .vgBitSujetoaIEPS = 0
                .vglngDatosParametro = True
            End With
        Else
            frmDatosFiscales.sstDatos.Tab = 1
        End If
    Else
        vlstrSentencia = "SELECT CHRRFCPOS RFC, null Clave, CHRNOMBREFACTURAPOS Nombre, CHRDIRECCIONPOS chrCalle," & _
                         "VCHNUMEROEXTERIORPOS vchNumeroExterior, VCHNUMEROINTERIORPOS vchNumeroInterior, " & _
                         "null Telefono,'OT' Tipo,  INTCVECIUDAD IdCiudad, VCHCOLONIAPOS Colonia, VCHCODIGOPOSTALPOS CP FROM PVParametro " & _
                         "WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
        Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic)
        
        If rs.RecordCount > 0 Then
            With frmDatosFiscales
                    .vgstrNombre = IIf(IsNull(rs!Nombre), "", Trim(rs!Nombre))
                    .vgstrDireccion = IIf(IsNull(rs!chrCalle), "", Trim(rs!chrCalle))
                    .vgstrNumExterior = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", Trim(rs!VCHNUMEROEXTERIOR))
                    .vgstrNumInterior = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", Trim(rs!VCHNUMEROINTERIOR))
                    .vgstrColonia = IIf(IsNull(rs!Colonia), "", Trim(rs!Colonia))
                    .vgstrCP = IIf(IsNull(rs!CP), "", Trim(rs!CP))
                    .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, str(IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)))
                    .vgstrTelefono = IIf(IsNull(rs!Telefono), "", Trim(rs!Telefono))
                    .vgstrRFC = IIf(IsNull(rs!RFC), "", Trim(rs!RFC))
                    .vlstrNumRef = IIf(vlstrTipo = "OT", "NULL", IIf(IsNull(rs!clave), 0, rs!clave))
                    .vlstrTipo = IIf(IsNull(rs!tipo), "OT", rs!tipo)
                    .vgActivaSujetoaIEPS = IIf(txtIEPS.Text > 0, True, False)
                    .vgBitSujetoaIEPS = IIf(txtIEPS.Text > 0, 1, 0)
                    .vglngDatosParametro = True
            End With
        Else
            frmDatosFiscales.sstDatos.Tab = 1
        End If
    End If
End Sub

Private Sub pGrabaDetalleFactura(strFolioFactura As String, dblFacturaPacienteIVA As Double, dblTotalControlAseguradoras As Double, dblDeducible As Double, dblCoaseguro As Double, dblCopago As Double, Optional vlblnSujetoaIEPS As Boolean)
On Error GoTo NotificaError
    Dim dblDescuentos As Double
    Dim vlstrFolioDocumentoDetalleFac As Integer
    Dim intPosicion As Integer
    Dim intcontador As Integer
    Dim intContador2 As Integer
    Dim dblCantidad As Double
    Dim lnga As Long
    Dim rsDetalleFactura As ADODB.Recordset
    Dim vldblIvaDescuento As Double

    '--------------------------------------
    ' Detalle de la Factura
    '--------------------------------------
    ' Aqui trabajo con un grid escondido(grdFactura), para que este mas fácil
    ' organizar los cargos, como en la facturación normal.
    '--------------------------------------
    blnFacturaDeducible = chkFacturarDeducible.Value = 1
    blnFacturaCoaseguro = chkFacturarCoaseguro.Value = 1
    blnFacturaCopago = chkFacturarCopago.Value = 1
    
    'Por si no tiene cargos
    If grdArticulos.RowData(1) = -1 Then Exit Sub
            
    'Se prepara el grid escondido para guardar los datos de la factura
    With grdFactura
        .Clear
        .ClearStructure
        .Cols = 9
        .Rows = 2
        .RowData(1) = -1
        .ColWidth(1) = 0 'Concepto de facturación (texto)
        .ColWidth(2) = 0 'Cargo
        .ColWidth(3) = 0 'Abono
        .ColWidth(4) = 0 'IVA
        .ColWidth(5) = 0 'Descuentos
        .ColWidth(6) = 0 'IEPS
        .ColWidth(7) = 0 'Tasa IEPS
        .ColWidth(8) = 0 'Folio del ticket para las facturas globales
    End With

    dblDescuentos = 0 'Para el cálculo de los descuentos
    For intcontador = 1 To grdArticulos.Rows - 1
        
        intPosicion = 0
        For intContador2 = 1 To grdFactura.Rows - 1  'For para ver si ya existe en el arreglo
            If CLng(grdArticulos.TextMatrix(intcontador, 13)) = grdFactura.RowData(intContador2) Then
                intPosicion = intContador2
                Exit For
            End If
        Next
        
        If intPosicion <> 0 Then
           ' La 2 es el CARGO
                If vlblnImportoVenta = False Then
                    grdFactura.TextMatrix(intPosicion, 2) = Val(grdFactura.TextMatrix(intPosicion, 2)) + _
                                                           ((Val(Format(grdArticulos.TextMatrix(intcontador, 4), "")) + _
                                                           IIf(vlblnLicenciaIEPS, IIf(vlblnSujetoaIEPS = False, Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")), 0), 0)) / vldblTipoCambio) '@
                Else
                    grdFactura.TextMatrix(intPosicion, 2) = Val(grdFactura.TextMatrix(intPosicion, 2)) + _
                                                           ((Val(Format(grdArticulos.TextMatrix(intcontador, 24), "")) + _
                                                           IIf(vlblnLicenciaIEPS, IIf(vlblnSujetoaIEPS = False, Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")), 0), 0)) / vldblTipoCambio) '@
                End If
                                                           
                ' El 4 es el IVA
                If vlblnImportoVenta = False Then
                    grdFactura.TextMatrix(intPosicion, 4) = Val(grdFactura.TextMatrix(intPosicion, 4)) + (Val(Format(grdArticulos.TextMatrix(intcontador, 11), "")) / vldblTipoCambio)
                Else
                    grdFactura.TextMatrix(intPosicion, 4) = Val(grdFactura.TextMatrix(intPosicion, 4)) + (Val(Format(grdArticulos.TextMatrix(intcontador, 26), "")) / vldblTipoCambio)
                End If
                    
                ' El 5 son los DESCUENTOS
                If vlblnImportoVenta = False Then
                    grdFactura.TextMatrix(intPosicion, 5) = Val(grdFactura.TextMatrix(intPosicion, 5)) + _
                                                            (Val(Format(grdArticulos.TextMatrix(intcontador, 5), "")) / vldblTipoCambio)
                Else
                    grdFactura.TextMatrix(intPosicion, 5) = Val(grdFactura.TextMatrix(intPosicion, 5)) + _
                                                            (Val(Format(grdArticulos.TextMatrix(intcontador, 25), "")) / vldblTipoCambio)
                End If
                                                    
                'El 6 es el IEPS
                If vlblnSujetoaIEPS Then
                    'grdFactura.TextMatrix(intPosicion, 6) = Val(grdFactura.TextMatrix(intPosicion, 6)) + Val(Format(grdArticulos.TextMatrix(intcontador, 6), ""))
                    grdFactura.TextMatrix(intPosicion, 6) = Val(grdFactura.TextMatrix(intPosicion, 6)) + (Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")) / vldblTipoCambio)
                End If
                
                If vlblnImportoVenta = True Then
                    grdFactura.TextMatrix(intPosicion, 8) = grdArticulos.TextMatrix(intPosicion, 22)
                End If
       Else
            If grdFactura.RowData(1) <> -1 Then grdFactura.Rows = grdFactura.Rows + 1
                
                grdFactura.RowData(grdFactura.Rows - 1) = CLng(grdArticulos.TextMatrix(intcontador, 13)) 'Clave del Concepto
                grdFactura.TextMatrix(grdFactura.Rows - 1, 1) = ""  'Descripción del Concepto
    
                If vlblnImportoVenta = False Then
                    dblCantidad = Round( _
                        ( _
                            Val( _
                                  Val(Format(grdArticulos.TextMatrix(intcontador, 4), "")) + _
                                  IIf(vlblnLicenciaIEPS, IIf(vlblnSujetoaIEPS = False, Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")), 0), 0) _
                                ) _
                        ) / vldblTipoCambio _
                                        , 3) '@
                Else
                    dblCantidad = Round( _
                        ( _
                            Val( _
                                  Val(Format(grdArticulos.TextMatrix(intcontador, 24), "")) + _
                                  IIf(vlblnLicenciaIEPS, IIf(vlblnSujetoaIEPS = False, Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")), 0), 0) _
                                ) _
                        ) / vldblTipoCambio _
                                        , 3) '@
                End If
                
                grdFactura.TextMatrix(grdFactura.Rows - 1, 2) = dblCantidad 'Cantidad
    
                If vlblnImportoVenta = False Then
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = Val(Format(grdArticulos.TextMatrix(intcontador, 11), "")) / vldblTipoCambio 'IVA
                Else
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 4) = Val(Format(grdArticulos.TextMatrix(intcontador, 26), "")) / vldblTipoCambio 'IVA
                End If
                    
                If vlblnImportoVenta = False Then
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 5) = Val(Format(grdArticulos.TextMatrix(intcontador, 5), "")) / vldblTipoCambio 'Descuentos individuales
                Else
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 5) = Val(Format(grdArticulos.TextMatrix(intcontador, 25), "")) / vldblTipoCambio 'Descuentos individuales
                End If
                
                'Monto IEPS y Tasa, como está agrupado por concepto de facturación, no se deben mezclar articulos con diferente IEPS bajo el mismo concepto
                If vlblnSujetoaIEPS Then
                    'grdFactura.TextMatrix(grdFactura.Rows - 1, 6) = Val(Format(grdArticulos.TextMatrix(intcontador, 6), ""))
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 6) = Val(Format(grdArticulos.TextMatrix(intcontador, 6), "")) / vldblTipoCambio
                    
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 7) = grdArticulos.TextMatrix(intcontador, 19)
                Else
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 6) = "0"
                End If
                
                If vlblnImportoVenta = True Then
                    grdFactura.TextMatrix(grdFactura.Rows - 1, 8) = grdArticulos.TextMatrix(intcontador, 22)
                End If
        End If
        
        dblDescuentos = dblDescuentos + Val(Format(grdArticulos.TextMatrix(intcontador, 5), "")) / vldblTipoCambio 'Sumar Descuentos
    
    Next

    vlstrSentencia = "Select * From PvDetalleFactura Where chrFolioFactura = ''"
    Set rsDetalleFactura = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    With rsDetalleFactura
        
        vldblIvaDescuento = 0
        
        For lnga = 1 To grdFactura.Rows - 1
            If grdFactura.RowData(lnga) > 0 Then 'Porque pueden ser negativos con el descuento y Pagos
                .AddNew
                !chrfoliofactura = strFolioFactura
                !smicveconcepto = grdFactura.RowData(lnga)
                !MNYCantidad = Val(grdFactura.TextMatrix(lnga, 2))
                !MNYIVA = Val(grdFactura.TextMatrix(lnga, 4))
                !MNYDESCUENTO = Val(grdFactura.TextMatrix(lnga, 5))
                !chrTipo = "NO"  'Osea que es un concepto "Normal"
                
                !mnyIVAConcepto = !MNYCantidad * (!MNYIVA / IIf((!MNYCantidad - !MNYDESCUENTO) > 0, (!MNYCantidad - !MNYDESCUENTO), 1))
                
                If !MNYIVA <> 0 Then
                    vldblIvaDescuento = vldblIvaDescuento + (!MNYDESCUENTO * (vgdblCantidadIvaGeneral / 100))
                End If
                
                !mnyIeps = Val(grdFactura.TextMatrix(lnga, 6))
                If !mnyIeps <> 0 Then
                    !numTasaIEPS = Val(grdFactura.TextMatrix(lnga, 7))
                End If
                
                '!VCHFOLIOTICKETFACTGLOBAL = IIf(vlblnImportoVenta = True, grdFactura.TextMatrix(lnga, 8), "")
                
                .Update
            End If
        Next lnga
        
        If dblDescuentos > 0 Then
            .AddNew
            !chrfoliofactura = strFolioFactura
            !smicveconcepto = -2
            !MNYCantidad = dblDescuentos
            !MNYIVA = 0
            !MNYDESCUENTO = dblDescuentos
            !chrTipo = "DE"  'Osea que es un concepto de tipo "Descuento"
            
            !mnyIVAConcepto = vldblIvaDescuento
            
            .Update
        End If

        ' Registro para el Deducible
        If ldblDeducible > 0 Then
            .AddNew
            !chrfoliofactura = strFolioFactura
            !smicveconcepto = llngCveConceptoDeducible
            !MNYCantidad = IIf(blnFacturaDeducible, Round((dblDeducible - (dblFacturaPacienteIVA * (((dblDeducible * 100) / dblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), dblDeducible)
            !MNYIVA = IIf(blnFacturaDeducible, dblFacturaPacienteIVA * (((dblDeducible * 100) / dblTotalControlAseguradoras) / 100) / vldblTipoCambio, 0) 'IVA
            !MNYDESCUENTO = 0
            !chrTipo = IIf(blnFacturaDeducible, "OD", "OP") 'Osea que es un concepto "Deducible, CoAseguro o CoPago"
            !mnyIVAConcepto = IIf(blnFacturaDeducible, !MNYIVA, 0)
            .Update
        End If
        
        'Registro para el Coaseguro
        If ldblCoaseguro > 0 Then
            .AddNew
            !chrfoliofactura = strFolioFactura
            !smicveconcepto = llngCveConceptoCoaseguro
            !MNYCantidad = IIf(blnFacturaCoaseguro, Round((dblCoaseguro - (dblFacturaPacienteIVA * (((dblCoaseguro * 100) / dblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), dblCoaseguro)
            !MNYIVA = IIf(blnFacturaCoaseguro, dblFacturaPacienteIVA * (((dblCoaseguro * 100) / dblTotalControlAseguradoras) / 100) / vldblTipoCambio, 0)  'IVA
            !MNYDESCUENTO = 0
            !chrTipo = IIf(blnFacturaCoaseguro, "OD", "OP")  'Osea que es un concepto "Deducible, CoAseguro o CoPago"
            !mnyIVAConcepto = IIf(blnFacturaDeducible, !MNYIVA, 0)
            .Update
        End If

        'Registro para el Copago
        If dblCopago > 0 Then
            .AddNew
            !chrfoliofactura = strFolioFactura
            !smicveconcepto = llngCveConceptoCopago
            !MNYCantidad = IIf(blnFacturaCopago, Round((dblCopago - (dblFacturaPacienteIVA * (((dblCopago * 100) / dblTotalControlAseguradoras) / 100))) / vldblTipoCambio, 2), dblCopago)
            !MNYIVA = IIf(blnFacturaCopago, dblFacturaPacienteIVA * (((ldblCopago * 100) / dblTotalControlAseguradoras) / 100) / vldblTipoCambio, 0)  'IVA
            !MNYDESCUENTO = 0
            !chrTipo = IIf(blnFacturaCopago, "OD", "OP")  'Osea que es un concepto "Deducible, CoAseguro o CoPago"
            !mnyIVAConcepto = IIf(blnFacturaDeducible, !MNYIVA, 0)
            .Update
        End If
    End With
    
    rsDetalleFactura.Close
            
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGrabaDetalleFactura"))
    Unload Me
End Sub

Function fblnFormatoCantidadPorciento(pstrTexto As TextBox, pintKeyascii As Integer, pbytDecimales As Byte) As Boolean
'--------------------------------------
'Funcion para validar que en una cantidad no se puedan poner mas de un punto
'y que le da formato a las cantidades ###,###,###,###.##
'y que sólo le pone el numero de digitos que le mandan por parametros
'--------------------------------------
    Dim vlintContador As Integer 'Para la contada de los decimales
    Dim vlblnTienePunto As Boolean 'Bandera para saber si la cantidad tiene punto
    Dim vlstrFormato As String 'Texto para ver como quedaría el formato
    Dim vlstrRollback As String
    Dim vlintSelStart As Integer 'Temporal para no perder la posición del cursor
    
    Dim intcontador As Integer
    Dim blnPunto As Boolean
    
    blnPunto = False
    
    vlblnTienePunto = False 'Inicializada
    vlstrFormato = "###.00"
    pstrTexto.Text = RTrim(pstrTexto.Text) 'Quitar espacios
    vlstrRollback = pstrTexto.Text
    fblnFormatoCantidadPorciento = True
    
    If Not pintKeyascii = vbKeyBack And Not pintKeyascii = vbKeyReturn Then
        If IsNumeric(Chr(pintKeyascii)) Or pintKeyascii = 46 Then
            If pbytDecimales = 0 And pintKeyascii = 46 Then
               fblnFormatoCantidadPorciento = False
            Else
                If pstrTexto.SelLength = 0 Or pstrTexto.SelLength <> Len(pstrTexto.Text) Then
                    For vlintContador = 1 To Len(pstrTexto.Text)
                        If Mid(pstrTexto.Text, vlintContador, 1) = "." Then
                            If Len(Mid(pstrTexto.Text, vlintContador)) = pbytDecimales + 1 Or pintKeyascii = 46 Then
                                If Not pintKeyascii = vbKeyBack And Not pintKeyascii = vbKeyReturn Then
                                    If (vlblnTienePunto And pintKeyascii = 46) Or (pstrTexto.SelStart > Len(pstrTexto.Text) - pbytDecimales + 1) Then
                                        fblnFormatoCantidadPorciento = False
                                        vlintContador = Len(pstrTexto.Text) + 1
                                    End If
                                End If
                            End If
                            vlblnTienePunto = True
                        Else
                            fblnFormatoCantidadPorciento = True
                        End If
                    Next
                Else
                    fblnFormatoCantidadPorciento = True
                End If
            End If
        Else
            fblnFormatoCantidadPorciento = False
        End If
        
        If fblnFormatoCantidadPorciento And Not pintKeyascii = 46 Then
            If Val(pstrTexto.Text) > 0 Then
                If pstrTexto.SelLength = Len(RTrim(pstrTexto.Text)) Then
                    pstrTexto.Text = Chr(pintKeyascii)
                Else
                    pstrTexto.Text = Mid(pstrTexto.Text, 1, pstrTexto.SelStart) & Chr(pintKeyascii) & Mid(pstrTexto.Text, pstrTexto.SelStart + 1, Len(pstrTexto.Text))
                End If
                vlintSelStart = Len(pstrTexto.Text)
                
                If vlblnTienePunto Then
                    If pbytDecimales > 0 Then
                        vlstrFormato = vlstrFormato & "."
                        For vlintContador = 1 To pbytDecimales
                            vlstrFormato = vlstrFormato & "0"
                        Next
                    End If
                Else
                    vlintSelStart = Len(pstrTexto.Text)
                End If
                pintKeyascii = 0
                pstrTexto.SelStart = vlintSelStart
                
            End If
        Else
            If pintKeyascii = 46 Then
                If pstrTexto.SelLength = Len(RTrim(pstrTexto.Text)) Then
                    pstrTexto.Text = Chr(pintKeyascii)
                Else
                    pstrTexto.Text = Mid(pstrTexto.Text, 1, pstrTexto.SelStart) & Chr(pintKeyascii) & Mid(pstrTexto.Text, pstrTexto.SelStart + 1, Len(pstrTexto.Text))
                End If
                If Not IsNumeric(Format(pstrTexto.Text, "#############")) And pstrTexto.Text <> "." Then
                    pstrTexto.Text = vlstrRollback
                End If
                pintKeyascii = 0
                pstrTexto.SelStart = Len(pstrTexto.Text)
            End If
        End If
        
        For intcontador = 1 To Len(txtDescuento.Text)
            If Mid(txtDescuento.Text, intcontador, 1) = "." Then
                blnPunto = True
            End If
        Next intcontador
           
        For intcontador = 1 To Len(txtDescuento.Text)
            If Not Mid(txtDescuento.Text, intcontador, 1) = "." And intcontador > 3 _
            And blnPunto = False Then
                txtDescuento.Text = Left$(txtDescuento.Text, 3)
            End If
        Next intcontador
    End If
End Function

Private Sub pCrearMovtoCredito( _
dblCantidadCredito As Double, _
lngNumCliente As Long, _
lngCtaContable As Long, _
strFolio As String, _
strTipo As String, _
dblTipoCambio As Double, _
lngPersonaGraba As Long)

    Dim dblTotalFactura As Double
    Dim dblIVAFactura As Double
    Dim dblPorcentaje As Double
    Dim dblSubtotalCredito As Double
    Dim dblIVACredito As Double
    Dim lngMovimientoCredito As Long
    Dim vlaryParametrosSalida() As String
    
    dblTotalFactura = Val(Format(txtTotal.Text, "############.##"))
    dblIVAFactura = Val(Format(txtIva.Text, "############.##"))
    dblPorcentaje = CDbl(dblCantidadCredito / (dblTotalFactura * IIf(optPesos(0).Value, 1, dblTipoCambio)))
    dblSubtotalCredito = CDbl(Format(((dblTotalFactura - dblIVAFactura) * IIf(optPesos(0).Value, 1, dblTipoCambio)) * dblPorcentaje, "###########.##"))
    dblIVACredito = Val(Format((dblIVAFactura * IIf(optPesos(0).Value, 1, dblTipoCambio)) * dblPorcentaje, "###########.##"))
    
    vgstrParametrosSP = _
    fstrFechaSQL(fdtmServerFecha) _
    & "|" & lngNumCliente _
    & "|" & lngCtaContable _
    & "|" & strFolio _
    & "|" & strTipo _
    & "|" & dblCantidadCredito _
    & "|" & str(vgintNumeroDepartamento) _
    & "|" & str(lngPersonaGraba) _
    & "|" & " " & "|" & "0" & "|" & dblSubtotalCredito & "|" & dblIVACredito
    
    lngMovimientoCredito = 1
    frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, lngMovimientoCredito
End Sub

Private Sub Form_Deactivate()
    vlblnPrimeraVez = False
End Sub

Private Sub grdArticulos_KeyPress(KeyAscii As Integer)
    If grdArticulos.Col = 2 Then
        If fblnRevisaPermiso(vglngNumeroLogin, 307, "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, 307, "E", True) Then
            If Trim(grdArticulos.TextMatrix(grdArticulos.Row, 1)) <> "" Then
                txtPrecio.Move grdArticulos.Left + grdArticulos.CellLeft, grdArticulos.Top + grdArticulos.CellTop, grdArticulos.CellWidth, grdArticulos.CellHeight
                txtPrecio.Text = grdArticulos.TextMatrix(grdArticulos.Row, grdArticulos.Col)
                txtPrecio.Visible = True
                pEnfocaTextBox txtPrecio
            End If
        End If
    End If
End Sub

Private Sub optCopagoCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtCopago
End Sub

Private Sub optControlCopagoPorciento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtCopago
End Sub

Private Sub optEmpleado_Click()
    If lngTipoPacEmpleado < 1 Then
        MsgBox SIHOMsg(967) & "empleado", vbCritical, "Mensaje" '¡No se ha configurado el tipo de paciente para la cuenta del empleado
        optPaciente.Value = True
        If txtClaveArticulo.Enabled Then
            txtClaveArticulo.SetFocus
        End If
    Else
        txtGafete.Enabled = True
    End If
End Sub

Private Sub optEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtGafete.SetFocus
End Sub

Private Sub optMedico_Click()
    txtGafete.Enabled = False
    If lngTipoPacMedico < 1 Then
        MsgBox SIHOMsg(967) & "médico!", vbCritical, "Mensaje" '¡No se ha configurado el tipo de paciente para la cuenta del médico
        optPaciente.Value = True
        If txtClaveArticulo.Enabled Then
            txtClaveArticulo.SetFocus
        End If
    End If
End Sub

Private Sub optMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtMovimientoPaciente.SetFocus
End Sub
Private Sub OptPaciente_Click()
    txtGafete.Enabled = False
End Sub
Private Sub OptPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtMovimientoPaciente.SetFocus
End Sub

Private Sub optTipoCoaseguro_Click(Index As Integer)
    lblSignoCoaseguro.Visible = Index = 1
End Sub

Private Sub optTipoCoaseguro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtCoaseguro
End Sub

Private Sub optTipoDeducible_Click(Index As Integer)
    lblSignoDeducible.Visible = Index = 1
End Sub

Private Sub optTipoDeducible_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtDeducible
End Sub

Private Sub txtCoaseguro_Change()
    chkFacturarCoaseguro.Enabled = Val(Format(txtCoaseguro.Text, "")) <> 0
    If Not chkFacturarCoaseguro.Enabled Then
        chkFacturarCoaseguro.Value = 0
    End If
End Sub

Private Sub txtCoaseguro_LostFocus()
    txtCoaseguro.Text = FormatNumber(Val(Format(txtCoaseguro.Text, "")), 2)
End Sub

Private Sub txtCopago_Change()
    chkFacturarCopago.Enabled = Val(Format(txtCopago.Text, "")) <> 0
    If Not chkFacturarCopago.Enabled Then
        chkFacturarCopago.Value = 0
    End If
End Sub

Private Sub txtDeducible_Change()
    chkFacturarDeducible.Enabled = Val(Format(txtDeducible.Text, "")) <> 0
    If Not chkFacturarDeducible.Enabled Then
        chkFacturarDeducible.Value = 0
    End If
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If fblnCanFocus(cmdDescuenta) Then cmdDescuenta.SetFocus
    End If
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
       
If Not fblnFormatoCantidadPorciento(txtDescuento, KeyAscii, 2) Then ' Solo permite números
    KeyAscii = 7
End If
          
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescuento_KeyPress"))
    Unload Me
End Sub

Private Sub txtGafete_GotFocus()
    lblMensajes.Caption = "Teclee el número de gafete del empleado"
End Sub

Private Sub txtGafete_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim llngNumCliente As Long
    Dim rsGafete As New ADODB.Recordset

    If KeyCode = vbKeyReturn Then
        If Val(Trim(txtGafete.Text)) > 0 Then
            Set rsGafete = frsRegresaRs("select count(*) cuenta from noempleado where intNumeroGafete = " & txtGafete.Text)
            If rsGafete!cuenta = 0 Then
                txtGafete.Text = ""
                MsgBox SIHOMsg(12), vbCritical, "Mensaje" '¡La información no existe!
                pNuevo
                txtClaveArticulo.SetFocus
                Exit Sub
            End If
        End If
        If Val(Trim(txtGafete.Text)) > 0 Then
            'Busca el número de empleado
            llngNumCliente = 1
            frsEjecuta_SP Trim(txtGafete.Text), "sp_PvSelNumClienteEmpleado", True, llngNumCliente
            If llngNumCliente > 0 Then
                pObtieneCuentaPaciente llngNumCliente
            Else
                txtGafete.Text = ""
                MsgBox SIHOMsg(966), vbCritical, "Mensaje" '¡El empleado no tiene credito!
                pNuevo
            End If
        End If
    End If

End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsPaquete As New ADODB.Recordset
    Dim rsEmpresa As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vldblValidoDolares As String
    'Dim vllngNumeroPaciente As Long
    Dim vllngCuentaExterno As Long
    'Dim llngNumCliente As Long
    Dim lintCreditoActivo As Integer    'Se utiliza para saber si se selecciono un paciente tipo medico o empleado con credito  1=Credito activo 0=sin credito
    Dim vldblMontoDisponiblePuntosMed As Double
    Dim intCveMedicoRelacionado As Double
    Dim rsPuntos As New ADODB.Recordset
    Dim vllngCveTipoPaciente As Long
    Dim rsPuntosMedico As New ADODB.Recordset
    Dim vldblMontoDisponiblePuntosPac As Double
    Dim rsMedico As New ADODB.Recordset
    Dim rsEmpleado As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        If fdblTipoCambio(fdtmServerFecha, "V") = 0 Then
            MsgBox SIHOMsg(231), vbCritical, "Mensaje"
            Exit Sub
        End If
        
        If chkConsultaPrecios.Value = 1 Then 'Caso 20487
            Exit Sub
        End If
        
        If Trim(txtMovimientoPaciente.Text) = "" Then
            
            If optPaciente.Value Then
                With FrmBusquedaPacientes
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "C"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = True
                    .optTodos.Value = True
                    .vgstrForma = "frmPOS"
                    .vgStrOtrosCampos = ", case when isnull(CCempresa.vchDescripcion,'*')='*' then AdTipoPaciente.vchDescripcion else CcEmpresa.vchDescripcion end as Empresa, case when isnull(Clientes.intNumCliente,0) = 0 then 'Sin crédito' else case when Clientes.bitActivo = 1 then 'Activo' else 'Inactivo' end end as ""Estado crédito"", " & _
                    " (Select Trim(GnDomicilio.vchCalle)||' '||Trim(GnDomicilio.vchNumeroExterior)||Case When GnDomicilio.vchNumeroInterior Is Null Then '' Else '  Int. '||Trim(GnDomicilio.vchNumeroInterior) End " & _
                    " From ExPacienteDomicilio " & _
                    " Inner Join GnDomicilio ON ExPacienteDomicilio.intCveDomicilio = GnDomicilio.intCveDomicilio " & _
                    " And GnDomicilio.intCveTipoDomicilio = 1 " & _
                    " Where ExPacienteDomicilio.intNumPaciente = ExPaciente.intNumPaciente) as Dirección, " & _
                    " ExPaciente.dtmFechaNacimiento as ""Fecha Nac."", " & _
                    " (Select GnTelefono.vchTelefono " & _
                    " From ExPacienteTelefono " & _
                    " Inner Join GnTelefono On ExPacienteTelefono.intCveTelefono = GnTelefono.intCveTelefono " & _
                    " And GnTelefono.intCveTipoTelefono = 1 " & _
                    " Where ExPacienteTelefono.intNumPaciente = ExPaciente.intNumpaciente) as Telefono "
                    .vgstrTamanoCampo = "800,3400,3000,1500,4100,990,980"
                    
                    pEnfocaTextBox txtClaveArticulo
                    
                    vllngNumeroPaciente = .flngRegresaPaciente()
                    lintCreditoActivo = 0
                    If vllngNumeroPaciente <> -1 Then 'Que SI está
                        If fblnCuentaAbierta(vllngNumeroPaciente) Then 'Que si no hay una cuenta abierta
                            MsgBox "Existe una cuenta abierta para este paciente, por favor primero cierre la cuenta.", vbCritical, "Mensaje"
                            pEnfocaTextBox txtClaveArticulo
                            Exit Sub
                        Else
                            lintCreditoActivo = IIf(.vgblnCreditoActivo = True, 1, 0)
                        End If
                    Else
                        frmAdmisionPaciente.vlblnMostrarTabGenerales = True
                        frmAdmisionPaciente.vlblnMostrarTabInternamiento = True
                        frmAdmisionPaciente.vlblnMostrarTabInternos = False
                        frmAdmisionPaciente.vlblnMostrarTabPrepagos = False
                        frmAdmisionPaciente.vlblnMostrarTabIngresosPrevios = False
                        frmAdmisionPaciente.vlblnMostrarTabEgresados = False
                        frmAdmisionPaciente.vlblnMostrarTabExternos = False
                        frmAdmisionPaciente.vlintPestañaInicial = 0
                        
                        frmAdmisionPaciente.blnAbrirCuenta = False
                        frmAdmisionPaciente.blnActivar = False
                        frmAdmisionPaciente.blnHabilitarAbrirCuenta = False
                        frmAdmisionPaciente.blnHabilitarActivar = False
                        frmAdmisionPaciente.blnHabilitarReporte = False
                        frmAdmisionPaciente.blnConsulta = False
                        frmAdmisionPaciente.vglngExpedienteConsulta = 0
                        frmAdmisionPaciente.blnHonorariosCC = False
                        
                        frmAdmisionPaciente.vllngNumeroOpcionExterno = 352
                        frmAdmisionPaciente.Show vbModal, Me
                        
                        vllngNumeroPaciente = frmAdmisionPaciente.vglngExpediente
                        
                        lintCreditoActivo = False
                    End If
                    
                    If vllngNumeroPaciente > 0 Then
                        vlstrSentencia = "SELECT ISNULL(EXTERNO.TNYTIPOCONVENIO,0) TIPOCONVENIO, " & _
                                         " ISNULL(EXTERNO.INTCLAVEEMPRESA,0) EMPRESA, " & _
                                         " TNYTIPOPACIENTE TIPOPACIENTE, " & _
                                         " EXTERNO.VCHNUMAFILIACION,EXTERNO.INTCIUDAD, Externo.tnyTipoPaciente " & _
                                         " FROM EXTERNO " & _
                                         " WHERE EXTERNO.INTNUMPACIENTE = " & Trim(str(vllngNumeroPaciente))
                        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                        'Se obtiene la clave de socio relacionado (en caso de que exista dicha relación), si es que no se regresó nada desde la pantalla de admisión
                        If vgintSocioRelacionado = 0 And (CLng(rs!tnyTipoPaciente) = flngTipoPacienteSocio) And (flngTipoPacienteSocio <> 0) Then
                            vgintSocioRelacionado = frmAdmisionPaciente.vglngPacienteSocioRel
                            
                            If vgintSocioRelacionado = 0 Then
                                Set rs = frsEjecuta_SP(CStr(vllngNumeroPaciente), "SP_SOSELSOCIORELACIONADO")
                        
                                If rs.RecordCount > 0 Then
                                    If Val(rs!intcvesocio) Then
                                        vgintSocioRelacionado = Val(rs!intcvesocio) 'clave del socio relacionado
                                    End If
                                End If
                            End If
                        End If
                        
                        vlngEmpresa = Trim(rs!empresa)
                        vllngCuentaExterno = 1
                        vgstrParametrosSP = vllngNumeroPaciente & "|" & rs!TipoConvenio & "|" & rs!empresa & "|" & rs!TipoPaciente & "|" & vllngMedicoDefaultPOS & "|" & vgintNumeroDepartamento & "|" & rs!VCHNUMAFILIACION & "||0|||||0|0|8|" & lintCreditoActivo & "|" & IIf(vgintSocioRelacionado = 0, "", vgintSocioRelacionado)
                        frsEjecuta_SP vgstrParametrosSP, "sp_GnNumCuentaExterno", True, vllngCuentaExterno
                        vgintSocioRelacionado = 0
                        txtMovimientoPaciente.Text = vllngCuentaExterno
                        rs.Close
                        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                    End If
                End With
                
                 If txtMovimientoPaciente <> "" Then
                        pFacturasDirectasAnteriores
                     End If
             
            Else
            'Médicos o Pacientes
                pEnfocaTextBox txtClaveArticulo
                llngNumCliente = flngCliente(IIf(optMedico.Value, "ME", "EM"))
                If llngNumCliente > 0 Then
                    pObtieneCuentaPaciente llngNumCliente
                    pFacturasDirectasAnteriores
                End If
            End If
        Else
        If Trim(txtMovimientoPaciente.Text) <> "" And optPaciente.Value Then
            intCuentaNueva = 1
        End If
            vlstrSentencia = "SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as Nombre, " & _
                    "RegistroExterno.intClaveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                    "RegistroExterno.tnyCveTipoPaciente cveTipoPaciente, AdTipoPaciente.vchDescripcion as Tipo, " & _
                    "Case when NVL(RegistroExterno.intClaveEmpresa, 0) = 0 " & _
                        "then 0 " & _
                        "else 1 " & _
                    " end as bitUtilizaConvenio, " & _
                    "EXTERNO.CHRCALLE, EXTERNO.VCHNUMEROEXTERIOR, EXTERNO.VchNumeroInterior, " & _
                    "Externo.vchcolonia as Colonia, Externo.vchcodpostal as CP, " & _
                    "Externo.chrTelefono as Telefono, ' ' as FechaIngreso, " & _
                    "' ' as Medico, " & _
                    "RegistroExterno.intCveExtra, chrRFC as RFC, (SELECT VCHCORREOELECTRONICO FROM EXPACIENTE E WHERE E.INTNUMPACIENTE = RegistroExterno.INTNUMPACIENTE) as VCHCORREOELECTRONICO, " & _
                    "' ' as Diagnostico, '' as Cuarto, ccTipoConvenio.bitAseguradora, " & _
                    "RegistroExterno.tnyTipoConvenio TipoConvenio, AdTipoPaciente.CHRTIPO TipoPac, EXTERNO.INTCIUDAD, ccempresa.INTCVECIUDAD " & _
                    "FROM RegistroExterno " & _
                    "INNER JOIN Externo ON " & _
                        "RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
                    "INNER JOIN AdTipoPaciente ON " & _
                       "RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                    " INNER JOIN NODEPARTAMENTO ON REGISTROEXTERNO.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                     "LEFT OUTER Join CcEmpresa ON " & _
                        "RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
                    "LEFT OUTER Join CcTipoConvenio ON  " & _
                        "ccEmpresa.tnyCveTipoConvenio = ccTipoConvenio.tnyCveTipoConvenio " & _
                    "Where RegistroExterno.intNumCuenta = " & txtMovimientoPaciente.Text & " AND NODEPARTAMENTO.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
                                
            'Si no es un paciente, se filtra también por el tipo de paciente sea ME o EM
            If Not optPaciente.Value Then
                vlstrSentencia = vlstrSentencia & " and RegistroExterno.TNYCVETIPOPACIENTE = " & IIf(optMedico.Value, lngTipoPacMedico, lngTipoPacEmpleado)
            End If
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            If rs.RecordCount <> 0 Then
                freDatosPaciente.Enabled = False
                '-------------------------------
                'Datos generales del Paciente
                '-------------------------------
                lblPaciente.Caption = " " & rs!Nombre
                vgstrRFCPersonaSeleccionada = IIf(IsNull(rs!RFC), "", rs!RFC)
                vgstrNombreFacturaPaciente = rs!Nombre
                vgstrDireccionFacturaPaciente = IIf(IsNull(rs!chrCalle), "", rs!chrCalle)
                vgstrNumeroExteriorFacturaPaciente = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", rs!VCHNUMEROEXTERIOR)
                vgstrNumeroInteriorFacturaPaciente = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", rs!VCHNUMEROINTERIOR)
                If rs!TipoPac = "CO" Then
                     llngCveCiudadPaciente = IIf(IsNull(rs!intCveCiudad), 0, rs!intCveCiudad)
                Else
                     llngCveCiudadPaciente = IIf(IsNull(rs!INTCIUDAD), 0, rs!INTCIUDAD)
                End If
                vgstrTelefonoFacturaPaciente = IIf(IsNull(rs!Telefono), "", rs!Telefono)
                vgstrRFCFacturaPaciente = IIf(IsNull(rs!RFC), "", rs!RFC)
                vgstrColoniaFacturaPaciente = IIf(IsNull(rs!Colonia), "", rs!Colonia)
                vgstrCPFacturaPaciente = IIf(IsNull(rs!CP), "", rs!CP)
                vgstrTipoPaciente = IIf(IsNull(rs!TipoPac), "", rs!TipoPac)
                vgintEmpresa = IIf(IsNull(rs!cveEmpresa), 0, rs!cveEmpresa)
                lblEmpresa.Caption = " " + IIf(rs!bitUtilizaConvenio = 1, IIf(IsNull(rs!empresa), "", rs!empresa), rs!tipo)
                vgintTipoPaciente = rs!cveTipoPaciente
                cmdControlAseguradora.Enabled = vgintEmpresa <> 0 And IIf(rs!bitAseguradora = 1, True, False)
                vglngCveExtra = IIf(IsNull(rs!intCveExtra), 0, rs!intCveExtra) 'Esta puede ser el numero de empleado o del médico
                'en esta parte podria ser posible la busqueda de los datos de el medico LAGP
                
                ' Datos para la factura
                If vgintEmpresa <> 0 Then 'Poner datos de la Empresa
                    vlstrSentencia = "select * from ccEmpresa where intCveEmpresa = " & Trim(str(vgintEmpresa))
                    Set rsEmpresa = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rsEmpresa.RecordCount > 0 Then
                        vgstrNombreFactura = IIf(IsNull(rsEmpresa!VCHRAZONSOCIAL), "", Trim(rsEmpresa!VCHRAZONSOCIAL))
                        vgstrDireccionFactura = IIf(IsNull(rsEmpresa!chrCalle), "", Trim(rsEmpresa!chrCalle)) 'Direccion para la factura
                        vgstrNumeroExteriorFactura = IIf(IsNull(rsEmpresa!VCHNUMEROEXTERIOR), "", Trim(rsEmpresa!VCHNUMEROEXTERIOR)) 'Número Exterior para la factura
                        vgstrNumeroInteriorFactura = IIf(IsNull(rsEmpresa!VCHNUMEROINTERIOR), "", Trim(rsEmpresa!VCHNUMEROINTERIOR)) 'Número Interior para la factura
                        vgstrColoniaFactura = IIf(IsNull(rsEmpresa!VCHCOLONIA), "", Trim(rsEmpresa!VCHCOLONIA)) 'Colonia para la factura
                        vgstrCPFactura = IIf(IsNull(rsEmpresa!VCHCODIGOPOSTAL), "", rsEmpresa!VCHCODIGOPOSTAL) 'Código postal para la factura
                        llngCveCiudad = IIf(IsNull(rsEmpresa!intCveCiudad), 0, rsEmpresa!intCveCiudad)
                                            
                                         
                        vgstrRegimenFiscal = IIf(IsNull(rsEmpresa!VCHREGIMENFISCAL), "", rsEmpresa!VCHREGIMENFISCAL)
                        
                        
                        vgstrTelefonoFactura = IIf(IsNull(rsEmpresa!chrTelefonoEmpresa), "", Trim(rsEmpresa!chrTelefonoEmpresa)) 'Telefono para la factura
                        vgstrRFCFactura = IIf(IsNull(rsEmpresa!chrRFCempresa), "", Trim(rsEmpresa!chrRFCempresa)) 'Telefono para la factura
                        vglngCveExtra = rsEmpresa!intcveempresa
                    Else
                        vgstrNombreFactura = ""
                        vgstrDireccionFactura = ""
                        vgstrNumeroExteriorFactura = ""
                        vgstrNumeroInteriorFactura = ""
                        vgBitExtranjeroFactura = 0
                        vgstrRegimenFiscal = ""
                        vgstrColoniaFactura = ""
                        vgstrCPFactura = ""
                        llngCveCiudad = 0
                        vgstrTelefonoFactura = ""
                        vgstrRFCFactura = ""
                    End If
                    rsEmpresa.Close
                Else   'Poner datos del paciente
                    If optMedico.Value Then
                        vlstrSentencia = "select * from homedico where intcvemedico = " & Trim(str(vglngCveExtra))
                        Set rsMedico = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                              
                    
                        vgstrNombreFactura = IIf(IsNull(rsMedico!vchNombre), "", Trim(rsMedico!vchNombre) & " " & Trim(rsMedico!vchApellidoPaterno) & " " & Trim(rsMedico!vchApellidoMaterno)) 'Nombre de Paciente
                        vgstrDireccionFactura = IIf(IsNull(rsMedico!VCHCONSULCALLE), "", rsMedico!VCHCONSULCALLE) 'Direccion para la factura
                        vgstrNumeroExteriorFactura = IIf(IsNull(rsMedico!VCHCONSULNUMEROEXTERIOR), "", rsMedico!VCHCONSULNUMEROEXTERIOR) 'Número Exterior para la factura
                        vgstrNumeroInteriorFactura = IIf(IsNull(rsMedico!VCHCONSULNUMEROINTERIOR), "", rsMedico!VCHCONSULNUMEROINTERIOR) 'Número Interior para la factura
                        vgstrColoniaFactura = IIf(IsNull(rsMedico!VCHCONSULCOLONIA), "", rsMedico!VCHCONSULCOLONIA) 'Colonia para la factura
                        vgstrCPFactura = IIf(IsNull(rsMedico!VCHCONSULCODPOSTAL), "", rsMedico!VCHCONSULCODPOSTAL) 'Código postal para la factura
                        
                        vgstrRegimenFiscal = IIf(IsNull(rsMedico!VCHREGIMENFISCAL), "", rsMedico!VCHREGIMENFISCAL)
                        
                        vgstrEmailCH = IIf(IsNull(rsMedico!vchEmail), "", rsMedico!vchEmail) 'EMAIL para la Factura JASM 20211227
                        llngCveCiudad = IIf(IsNull(rsMedico!intCveCiudad), 0, rsMedico!intCveCiudad)
                        vgstrTelefonoFactura = "" 'Telefono para la factura
                        vgstrRFCFactura = IIf(IsNull(rsMedico!vchRfcMedico), "", rsMedico!vchRfcMedico) 'rfc
                    
                    ElseIf optEmpleado.Value Then ' en este caso es empleado
                        'Se modificó porque marcaba error cuando la variable venia vacía y se igualaba a 0 ya que el query no traía datos, cuando es 0 trae los datos del empleado con la referencia del cliente
                        If vglngCveExtra = 0 Then
                            vlstrSentencia = "select * from noempleado where intcveempleado = (SELECT NVL(intNumReferencia, 0) FROM CCCLIENTE WHERE intNumCliente = " & llngNumCliente & ")"
                        Else
                            vlstrSentencia = "select * from noempleado where intcveempleado = " & Trim(str(vglngCveExtra))
                        End If
                        Set rsEmpleado = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    
                        vgstrNombreFactura = IIf(IsNull(rsEmpleado!vchNombre), "", rsEmpleado!vchNombre & " " & rsEmpleado!vchApellidoPaterno & " " & rsEmpleado!vchApellidoMaterno) 'Nombre de Paciente
                        vgstrDireccionFactura = IIf(IsNull(rsEmpleado!chrCalle), "", rsEmpleado!chrCalle) 'Direccion para la factura
                        vgstrNumeroExteriorFactura = IIf(IsNull(rsEmpleado!VCHNUMEROEXTERIOR), "", rsEmpleado!VCHNUMEROEXTERIOR) 'Número Exterior para la factura
                        vgstrNumeroInteriorFactura = IIf(IsNull(rsEmpleado!VCHNUMEROINTERIOR), "", rsEmpleado!VCHNUMEROINTERIOR) 'Número Interior para la factura
                        vgstrColoniaFactura = IIf(IsNull(rsEmpleado!chrColonia), "", rsEmpleado!chrColonia) 'Colonia para la factura
                        vgstrCPFactura = IIf(IsNull(rsEmpleado!chrCodigoPostal), "", rsEmpleado!chrCodigoPostal) 'Código postal para la factura
                        vgstrRegimenFiscal = IIf(IsNull(rsEmpleado!VCHREGIMENFISCAL), "", rsEmpleado!VCHREGIMENFISCAL)
                        vgstrEmailCH = IIf(IsNull(rsEmpleado!vchCorreo), "", rsEmpleado!vchCorreo) 'EMAIL para la Factura JASM 20211227
                        llngCveCiudad = IIf(IsNull(rsEmpleado!intCveCiudad), 0, rsEmpleado!intCveCiudad)
                        vgstrTelefonoFactura = IIf(IsNull(rsEmpleado!chrTelefono), "", rsEmpleado!chrTelefono) 'Telefono para la factura
                        vgstrRFCFactura = IIf(IsNull(rsEmpleado!CHRRFC), "", rsEmpleado!CHRRFC) 'Telefono para la factura
                    Else ' en este caso es cuando es paciente sin convenio
                    
                        vgstrNombreFactura = IIf(IsNull(rs!Nombre), "", rs!Nombre) 'Nombre de Paciente
                        vgstrDireccionFactura = IIf(IsNull(rs!chrCalle), "", rs!chrCalle) 'Direccion para la factura
                        vgstrNumeroExteriorFactura = IIf(IsNull(rs!VCHNUMEROEXTERIOR), "", rs!VCHNUMEROEXTERIOR) 'Número Exterior para la factura
                        vgstrNumeroInteriorFactura = IIf(IsNull(rs!VCHNUMEROINTERIOR), "", rs!VCHNUMEROINTERIOR) 'Número Interior para la factura
                        vgstrColoniaFactura = IIf(IsNull(rs!Colonia), "", rs!Colonia) 'Colonia para la factura
                        vgstrCPFactura = IIf(IsNull(rs!CP), "", rs!CP) 'Código postal para la factura
                        vgstrEmailCH = IIf(IsNull(rs!vchCorreoElectronico), "", rs!vchCorreoElectronico) 'EMAIL para la Factura JASM 20211227
                        llngCveCiudad = IIf(IsNull(rs!INTCIUDAD), 0, rs!INTCIUDAD)
                        vgstrTelefonoFactura = IIf(IsNull(rs!Telefono), "", rs!Telefono) 'Telefono para la factura
                        vgstrRFCFactura = IIf(IsNull(rs!RFC), "", rs!RFC) 'Telefono para la factura
                    End If
                    
                End If
                '--------------------------------------
                'Estados de los frames, Tabs y estatus
                '--------------------------------------
                vgstrEstadoManto = "V" 'Venta
                
                pEnfocaTextBox txtClaveArticulo
            Else
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            End If
            rs.Close
            
            ' -- Caso 17115 --
            ' Se muestra en una etiqueta los puntos disponibles con los que cuenta el paciente, en caso de tener
            ' la licencia activa para generar la lealtad del cliente con el hospital por medio del otorgamiento de puntos
            vlstrMensajePuntos = ""
            vldblMontoDisponiblePuntos = 0
            vldblMontoDisponiblePuntosPac = 0
            vldblMontoDisponiblePuntosMed = 0
            vldblPuntosDisponibles = 0
            If blnLicenciaLealtadCliente Then
                pDatosPaciente txtMovimientoPaciente.Text
                '-- puntos paciente
                vgstrParametrosSP = Val(txtMovimientoPaciente.Text) & "|1|-1"
                Set rsPuntos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELPUNTOSDISPONIBLES")
        
                If rsPuntos.RecordCount > 0 Then vldblMontoDisponiblePuntosPac = rsPuntos!puntosDisponibles

                vlstrSentencia = "select pi.intcvetipoingreso, pi.intcvetipopaciente, nvl(pi.intCveMedicoRelacionado, 0) intCveMedicoRelacionado, pi.intcvemedicotratante, tp.bitfamiliar, tp.chrtipo from ExPacienteIngreso pi inner join AdTipoPaciente tp on pi.intcvetipopaciente = tp.tnycvetipopaciente where pi.intNumCuenta = " & Val(txtMovimientoPaciente.Text)
                Set rs = frsRegresaRs(vlstrSentencia, adOpenDynamic, adLockOptimistic)
                If rs.RecordCount > 0 Then
                    vllngCveTipoPaciente = rs!intCveTipoPaciente
            
                    If (rs!bitFamiliar = 1 And rs!chrTipo = "ME") Or rs!chrTipo = "ME" Then
                        '-- puntos médico o familiar de médico
                        intCveMedicoRelacionado = rs!intCveMedicoRelacionado
                        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) & "|0|" & intCveMedicoRelacionado
                        Set rsPuntosMedico = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELPUNTOSDISPONIBLES")
                
                        If rsPuntosMedico.RecordCount > 0 Then vldblMontoDisponiblePuntosMed = rsPuntosMedico!puntosDisponiblesMed
                    End If
                End If
                
                If (vldblMontoDisponiblePuntosPac + vldblMontoDisponiblePuntosMed) > 0 Then
                    vldblPuntosDisponibles = (vldblMontoDisponiblePuntosPac + vldblMontoDisponiblePuntosMed)
                    vldblMontoDisponiblePuntos = ((vldblMontoDisponiblePuntosPac + vldblMontoDisponiblePuntosMed) * fdblValorPuntoLealtad)
                    If vldblMontoDisponiblePuntos > 0 Then
                        vlstrMensajePuntos = "Puntos acumulados " & (vldblMontoDisponiblePuntosPac + vldblMontoDisponiblePuntosMed) & " que equivalen a " & Format(vldblMontoDisponiblePuntos, "$ ###,###,###,###.00")
                    End If
                Else
                    vlstrMensajePuntos = "Puntos acumulados " & (vldblMontoDisponiblePuntosPac + vldblMontoDisponiblePuntosMed) & " que equivalen a " & Format(vldblMontoDisponiblePuntos, "$ ###,###,###,##0.00")
                End If
                pEtiquetaVar "Generales", vlstrMensajePuntos
                If vldblMontoDisponiblePuntos > 0 Then
                    cmdDescuentoPuntos.Enabled = True
                Else
                    'pEtiquetaVar "", ""
                    cmdDescuentoPuntos.Enabled = False
                End If
            End If
        End If
    End If
    
End Sub
Private Sub pSeleccionaElemento(vllngClaveElemento As Long, vlintCantidad As Long, vlstrTipoCargo As String)
    
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim rsPMP As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlStrTipoPaciente As String
    Dim vlintPosicion As Integer
    Dim vlintContador As Integer
    Dim lstListas As ListBox
    Dim vldblPrecio As Double
    Dim vldblIncrementoTarifa As Double 'Incremento en la tarifa de los precios
    Dim vldblSubtotal As Double
    Dim vldblDescuento As Double
    Dim vldblIEPS As Double '<----------
    Dim vldblIEPSPercent As Double '<---
    Dim vldblIVA As Double
    Dim vldblDescUnitario As Double
    Dim vldblTotDescuento As Double
    Dim vlstrTipoDescuento As String
    Dim vlintModoDescuentoInventario As Integer
    Dim vllngContenido As Long
    Dim vlstrAux As String
    Dim a As Integer
    Dim vlstrPrecio As String
    Dim vlstrIncrementoTarifa As String
    Dim vlaryParametrosSalida() As String
    Dim vlstrCveArticulo As String
    Dim vlchrTipoDescuento As String
    Dim vldblPorcentajeDescuento As Double
    Dim vlintAuxExclusionDescuento As Long
    Dim DescuentoInventario As Integer
    Dim MNYCantidad As Integer
    Dim rsVentaAlmacen As New ADODB.Recordset
    Dim vlintAlmacenVenta As Integer
    Dim vlbExistencias As Boolean
    Dim vlchrDescripcion As String
    'Dim vlintUV As Integer
    'Dim vlintUM As Integer
    Dim lngNumCargo As Long
    Dim vlintUV As Long ' caso 20342
    Dim vlintUM As Long ' caso 20342
    Dim rsDepto As New ADODB.Recordset
    Dim rslotes As New ADODB.Recordset
    Dim blnConvertir As Boolean ' Caso 20453
    Dim lngRow As Long ' Caso 20453
    Dim lngAuxExistencia As Long ' Caso 20453
    'Caso 20417
    Dim intRespuesta As Integer
    Dim vlLngCantidadAlternaDesc As Long
    Dim rsUnidadesAltMin As New ADODB.Recordset
    'Caso 20417
    
    blnConvertir = False ' caso 20453
        
    If Val(Format(vlintCantidad, "")) > 0 Then
        With grdArticulos
            vllngContenido = 1  'Este es el Contenido en IvArticulo
            If vllngClaveElemento <> -1 Then
                vlintModoDescuentoInventario = 0
                If vlstrTipoCargo = "AR" Then 'Nomas para los articulos
                    '-------------------
                    ' Tipo de descuento de Inventario
                    '-------------------
                    vlstrSentencia = "SELECT intContenido Contenido,vchNombreComercial Descripcion, substring(vchNombreComercial,1,50) Articulo,  ivUA.vchDescripcion UnidadAlterna,  ivUM.vchDescripcion UnidadMinima," & _
                                    " chrCveArticulo CveArticulo, ivArticulo.bitVentaPublico" & _
                                    " From ivArticulo " & _
                                    " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                    " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                    " WHERE intIdArticulo = " & Trim(str(vllngClaveElemento))

                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rs.RecordCount > 0 Then
                        If rs!bitVentaPublico = 0 Then
                            MsgBox SIHOMsg(1548), vbInformation, "Mensaje"
                            rs.Close
                            Exit Sub
                        End If
                        vlstrCveArticulo = rs!cveArticulo
                        vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
                        vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                        vlchrDescripcion = rs!Descripcion
                        
                        If chkConsultaPrecios.Value = 0 Then ' Caso 20487
                            'caso 20417
                            If vlstrTipoCargo = "AR" Then
                                Set rsVentaAlmacen = frsRegresaRs("SELECT smiCveDepartamento FROM NoDepartamento WHERE NoDepartamento.chrClasificacion = 'A' AND smiCveDepartamento = " & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                        
                                If Not rsVentaAlmacen.RecordCount = 0 Then
                                    vlintAlmacenVenta = rsVentaAlmacen!smicvedepartamento
                                Else
                                    Set rsVentaAlmacen = frsRegresaRs("SELECT intNumAlmacen FROM PvAlmacenes WHERE intnumdepartamento =" & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                                    vlintAlmacenVenta = rsVentaAlmacen!intnumalmacen
                                End If
                            End If
                                                   
                             Set rsUnidadesAltMin = frsRegresaRs("SELECT INTEXISTENCIADEPTOUV VL_UV, INTEXISTENCIADEPTOUM VL_UM FROM IVUBICACION WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta, adLockOptimistic, adLockOptimistic, adOpenDynamic)
                             If Not rsUnidadesAltMin.RecordCount = 0 Then
                                 If rsUnidadesAltMin!VL_UV = 0 Then
                                     MsgBox "El artículo " & Trim(rs!Articulo) & " sólo cuenta con unidades mínimas, la venta se realizara con esta unidad", vbInformation + vbOKOnly, "Mensaje"
                                         vlintModoDescuentoInventario = 1
                                 End If
                            
                                 'caso 20417 / 20453
                                 If rsUnidadesAltMin!VL_UV > 0 Then
                                     If vllngContenido > 1 Then
                                         If MsgBox("¿Desea realizar la venta de " & Trim(rs!Articulo) & " por " & Trim(rs!UnidadAlterna) & "?" & Chr(13) & "Si selecciona NO, se venderá por " & Trim(rs!UnidadMinima) & ".", vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                                             vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
                                         End If
                                     End If
                                 End If
                             End If
                             rs.Close
                        End If
                    End If
                End If
                '-----------------------
                'Precio unitario (En TODO el POS es SIN AUMENTO de TARIFAS, así lo dijo SERGIO)
                '-----------------------
                pCargaArreglo vlaryParametrosSalida, "|" & adDecimal & "||" & adDecimal
                frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|1|01/01/1900|" & vgintClaveEmpresaContable, "SP_PVSELOBTENERPRECIO", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vldblPrecio, vldblIncrementoTarifa
                
                If vldblPrecio = -1 Or vldblPrecio = 0 Then
                    MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                '*********************************************
                If vlintValorPMP = 1 Then
                    'Se compara el precio con el Precio Máximo al Publico
                     vlstrSentencia = "SELECT NVL(MNYPRECIOMAXIMOPUBLICO,0) precio"
                    vlstrSentencia = vlstrSentencia & " FROM IVARTICULO INNER JOIN IVARTICULOEMPRESAS ON IVARTICULO.chrcvearticulo = IVARTICULOEMPRESAS.chrcvearticulo"
                    vlstrSentencia = vlstrSentencia & " and IVARTICULO.INTIDARTICULO =  " & vllngClaveElemento & ""
                    vlstrSentencia = vlstrSentencia & " AND IVARTICULOEMPRESAS.tnyclaveempresa = " & vgintClaveEmpresaContable & ""
                    vlstrSentencia = vlstrSentencia & " WHERE  IVARTICULO.BITVENTAPUBLICO = 1 AND IVARTICULO.CHRCVEARTMEDICAMEN = 1"
    
                    Set rsPMP = frsRegresaRs(vlstrSentencia)
                    If rsPMP.RecordCount <> 0 Then
                        If (IIf(IsNull(rsPMP!precio), 0, rsPMP!precio) > 0 And vldblPrecio >= rsPMP!precio) Then vldblPrecio = rsPMP!precio
                    End If
                    rsPMP.Close: Set rsPMP = Nothing
                End If
                '*********************************************
'                vldblPrecio = Format(vldblPrecio, "############.00")

                'Existencias
                vlbExistencias = False
                DescuentoInventario = vlintModoDescuentoInventario
                MNYCantidad = vlintCantidad
                
                If vlstrTipoCargo = "AR" Then
                    Set rsVentaAlmacen = frsRegresaRs("SELECT smiCveDepartamento FROM NoDepartamento WHERE NoDepartamento.chrClasificacion = 'A' AND smiCveDepartamento = " & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                    
                    If Not rsVentaAlmacen.RecordCount = 0 Then
                        vlintAlmacenVenta = rsVentaAlmacen!smicvedepartamento
                    Else
                        Set rsVentaAlmacen = frsRegresaRs("SELECT intNumAlmacen FROM PvAlmacenes WHERE intnumdepartamento =" & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                        vlintAlmacenVenta = rsVentaAlmacen!intnumalmacen
                    End If
                    vlstrSentencia = "SELECT INTEXISTENCIADEPTOUV VL_UV, INTEXISTENCIADEPTOUM VL_UM FROM IVUBICACION WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rs.RecordCount > 0 Then
                        vlintUV = rs!VL_UV
                        vlintUM = rs!VL_UM
                    Else
                        vlintUV = 0
                        vlintUM = 0
                    End If
                    
                    'Por unidad minima
                    If DescuentoInventario = 1 And vllngContenido > 1 Then
                        If (((vlintUV * vllngContenido) + vlintUM) < MNYCantidad) Then
                            vlbExistencias = True
                        End If
                    'Por unidad alterna
                    Else
                        If vllngContenido = 1 And DescuentoInventario = 1 Then
                            DescuentoInventario = 2
                        End If
                        If (((vlintUV * vllngContenido) + vlintUM)) < (MNYCantidad * vllngContenido) And DescuentoInventario = 2 Then
                            vlbExistencias = True
                        End If
                    End If
                    
                     If vlbExistencias Then
                        If chkConsultaPrecios.Value = 0 Then
                           MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(317) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
                           Exit Sub
                        End If
                     End If
                     
                     'Caso 20417
                     'Se realiza la conversión de unidad alterna a minima
                      If vlintModoDescuentoInventario = 1 Then
                          If vlintUM <> MNYCantidad Then
                              'Validar si el producto se puede convertir
                              Set rs = frsRegresaRs("select INTCVEUNIALTERNAVTA ALTERNAVTA, INTCVEUNIMINIMAVTA MINIMAVTA, INTCONTENIDO CONTENIDO from IVARTICULO WHERE CHRCVEARTICULO ='" & vlstrCveArticulo & "'", adLockReadOnly, adOpenForwardOnly)
                              If rs.RecordCount > 0 Then
                                  If rs!ALTERNAVTA <> rs!MINIMAVTA Then
                                        If vlintUM <= MNYCantidad Then
                                            blnConvertir = True
                                        Else
                                            With grdArticulos
                                                For lngRow = 1 To .Rows - 1
                                                    If Val(.TextMatrix(lngRow, 10)) = LngCveArticulo(vlstrCveArticulo) And Val(.TextMatrix(lngRow, 14)) = vlintModoDescuentoInventario Then
                                                       lngAuxExistencia = lngAuxExistencia + .TextMatrix(lngRow, 3)
                                                    End If
                                                Next
                                            End With
                                            
                                            If lngAuxExistencia = vlintUM Then
                                                If MsgBox("Todas las unidades mínima ya estan capturadas en la cuadrícula, ¿desea que el sistema realice la conversión de unidades?" & Chr(13) & Chr(13) & "Sí, para continuar" & Chr(13) & "No, para cancelar", vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje") = vbYes Then
                                                    blnConvertir = True
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                   
                                        End If
                                        
                                        If blnConvertir Then
                                            intRespuesta = MsgBox("¿Desea realizar la conversión de unidad alterna a mínima de " & Trim(vlchrDescripcion) & "?" & Chr(13) & Chr(13) & "Sí, realizar la conversión y esta será definitiva." & Chr(13) & "No, no realizar la conversión." & Chr(13) & "Cancel, cancela la conversión.", vbYesNoCancel + vbQuestion + vbDefaultButton2, "Mensaje")
                                            
                                            'en caso de cancelar sale del proceso
                                            If intRespuesta = vbCancel Then
                                                Exit Sub
                                            End If
                                            
                                            'se realiza la conversión a unidad minima
                                            If intRespuesta = vbYes Then
                                                
                                                If MsgBox("Una vez que se realice la conversión, las ventas de este artículo / medicamento deben de realizarse por cantidad mínima, ¿desea continuar?" & Chr(13) & Chr(13) & "Sí, para continuar" & Chr(13) & "No, para cancelar", vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje") = vbYes Then
                                            
                                                    ' saber las cantidades a descontar en unidad alterna
                                                    If (vlintUV * vllngContenido) + vlintUM > MNYCantidad Then
                                                        vlLngCantidadAlternaDesc = RoundUP(MNYCantidad / vllngContenido)
                                                        If MsgBox("Hay en existencia " & vlintUV & " unidades alternas, se tomarán " & vlLngCantidadAlternaDesc & " para convertirlas en unidades minímas." & Chr(13) & Chr(13) & "¿Desea continuar con la conversión?", vbYesNo + vbQuestion + vbDefaultButton2, "Mensaje") = vbNo Then
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        vlLngCantidadAlternaDesc = vlintUV
                                                    End If
                                                    
                                                    ' Se valida cantidad a descontar contra el stock en alternas
                                                    If vlLngCantidadAlternaDesc > vlintUV Then
                                                        MsgBox "No se pudo completar la operación de convertir unidad alterna a mínima; " & SIHOMsg(317) & Chr(13) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
                                                        pCalculaTotales
                                                        Exit Sub
                                                    End If
                                                    
                                                    'se realiza la conversión a unidad minima
                                                    vlintUM = (vlLngCantidadAlternaDesc * vllngContenido) + vlintUM
                                                    
                                                    'se resta la unidad alterna al almacén
                                                    vlstrSentencia = "UPDATE IVUBICACION SET INTEXISTENCIADEPTOUV = INTEXISTENCIADEPTOUV - " & vlLngCantidadAlternaDesc & ", INTEXISTENCIADEPTOUM = " & vlintUM & " WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta
                                                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                                    
                                                    'convierte en  la tabla de IVLOTEUBICACION
                                                    vlstrSentencia = "UPDATE IVLOTEUBICACION SET INTEXISTENCIADEPTOUM =" & vlintUM & ", INTEXISTENCIADEPTOUV = INTEXISTENCIADEPTOUV -" & vlLngCantidadAlternaDesc & " WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta & " AND INTEXISTENCIADEPTOUV=" & vlintUV
                                                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                                Else
                                                    Exit Sub
                                                End If
                                            End If
                                      End If
                                  End If
                              End If
                          End If
                      End If
                     'Caso 20417
                End If
                
                '----------------------
                'Clave de la lista de precios
                '----------------------
                pCargaArreglo vlaryParametrosSalida, "|" & adInteger
                frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|" & vgintClaveEmpresaContable, "SP_PVSELLISTAPRECIO", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, lCveLista
                
                '-----------------------
                'Exclusión de descuentos
                '-----------------------
                vlintAuxExclusionDescuento = 1
                frsEjecuta_SP "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & vlstrTipoCargo & "|" & vllngClaveElemento & "|" & vgintNumeroDepartamento, "FN_PVSELEXCLUSIONDESCUENTO", True, vlintAuxExclusionDescuento
                '-------------------------------------------------------------
                'vlintAuxExclusionDescuento = 2 SI HAY EXCLUSION DE DESCUENTOS
                'vlintAuxExclusionDescuento = 3 NO HAY EXCLUSION DE DESCUENTOS
                '-------------------------------------------------------------
                If vlintAuxExclusionDescuento = 2 Then
                   vldblTotDescuento = 0
                Else
                    '-----------------------
                    'Descuentos
                    '-----------------------
                    vldblTotDescuento = 0
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    vgstrParametrosSP = "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & CStr(Val(txtMovimientoPaciente.Text)) & "|" & vlstrTipoCargo & "|" & CStr(vllngClaveElemento) & "|" & vldblPrecio & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & Trim(str(intAplicado)) & "|" & CStr(vllngContenido) & "|" & CStr(vlintCantidad) & "|" & CStr(vlintModoDescuentoInventario)
                    frsEjecuta_SP vgstrParametrosSP, "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblTotDescuento
                End If
                
                '-----------------------
                'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                '-----------------------
                If vlintModoDescuentoInventario = 1 Then
                    vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                End If
                
'                vldblPrecio = Format(vldblPrecio, "############.00")
                
                vldblSubtotal = (vldblPrecio * CLng(vlintCantidad)) - vldblTotDescuento
                
                '----------------------------------
                ' Porcentaje de IEPS
                '----------------------------------
                If vlblnLicenciaIEPS And vlstrTipoCargo = "AR" Then
                   vldblIEPSPercent = fdblObtenerPorcentajeIEPS(CLng(vllngClaveElemento)) / 100 '<----------------
                Else
                   vldblIEPSPercent = 0
                End If
                '----------------------------------
                vldblIEPS = vldblSubtotal * vldblIEPSPercent '<---sacamos cantidada de IEPS
                vldblSubtotal = vldblSubtotal + vldblIEPS '<-----sumamos el IEPS al Subtotal
                               
                '-----------------------
                ' Procedimiento para obtener % del IVA
                '-----------------------
                vldblIVA = fdblObtenerIva(CLng(vllngClaveElemento), vlstrTipoCargo) / 100
                
                '-----------------------
                ' Datos del Artículo ó del Otro concepto
                '-----------------------
                If optArticulo.Value Then
                    vlstrSentencia = "select AR.vchNombreComercial Descripcion, AR.smiCveConceptFact Concepto, CF.chrDescripcion ConceptoDesc from IVArticulo AR inner join PVConceptoFacturacion CF on CF.smiCveConcepto = AR.smiCveConceptFact where AR.intIdArticulo = " & Trim(str(vllngClaveElemento))
                Else
                    vlstrSentencia = "select OC.chrDescripcion Descripcion, OC.smiConceptoFact Concepto, CF.chrDescripcion ConceptoDesc from PVOtroConcepto OC inner join PVConceptoFacturacion CF on CF.smiCveConcepto = OC.smiConceptoFact where OC.intCveConcepto = " & Trim(str(vllngClaveElemento))
                End If
                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
                '-----------------------
                ' Llenado del grid o de la consulta de precios
                '-----------------------
                If chkConsultaPrecios.Value = 0 Then
                    
                    lblnGuardoControl = False
                
                    If .RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    

                    .TextMatrix(.Row, 1) = rs!Descripcion
                    .TextMatrix(.Row, 2) = Format(vldblPrecio, "$###,###,###,###.0000")
                    .TextMatrix(.Row, 24) = vldblPrecio
                    
                    .TextMatrix(.Row, 3) = vlintCantidad
                    .TextMatrix(.Row, 4) = Format(vldblPrecio * CLng(vlintCantidad), "$###,###,###,###.00")
                    .TextMatrix(.Row, 5) = Format(vldblTotDescuento, "$###,###,###,###.00")
                    .TextMatrix(.Row, 6) = Format(vldblIEPS, "$###,###,###,###.00") '@
                    .TextMatrix(.Row, 7) = Format(vldblSubtotal, "$###,###,###,###.00") '@
                    .TextMatrix(.Row, 8) = Format(vldblSubtotal * vldblIVA, "$###,###,###,###.00")
                    .TextMatrix(.Row, 9) = Format(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 8)), "$###,###,###,###.00")
                    .Col = 9
                    .CellFontBold = True
                    .TextMatrix(.Row, 10) = vllngClaveElemento
                    .TextMatrix(.Row, 11) = vldblSubtotal * vldblIVA
                    .TextMatrix(.Row, 12) = vlstrTipoCargo
                    .TextMatrix(.Row, 13) = rs!Concepto  'Clave del concepto de Facturación
                    .TextMatrix(.Row, 14) = vlintModoDescuentoInventario 'Descuento por unidad alterna(2) o unidad mínima(1)
                    .TextMatrix(.Row, 16) = vllngContenido   'Contenido de IVArticulo
                    .TextMatrix(.Row, 17) = lCveLista   'Clave de la lista de precios
                    .TextMatrix(.Row, 18) = vldblPorcentajeDescuento 'Porcentaje de descuento
                    .TextMatrix(.Row, 19) = vldblIEPSPercent 'Porcentaje de IEPS
                    .TextMatrix(.Row, 20) = vlintAuxExclusionDescuento '2 = tiene exclusión de descuento, 3 = no tiene exclusion de descuento
                    .TextMatrix(.Row, 21) = rs!ConceptoDesc 'Descripción del concepto de Facturación
                    .TextMatrix(.Row, 23) = "0"
                    .RowData(.Row) = vllngClaveElemento
                    .Redraw = True
                    .Refresh

                     lngNumCargo = Val(.RowData(.Row))
                      
                    'Caso 20370
                    'Validar el stock de inventario Vs Solicitud de salidas(venta al publico)
                    'GIRM
                    
                    If BlnValidarCantidad(vlintUM, vlintUV, vllngClaveElemento, vlintModoDescuentoInventario) Then
                        MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(317) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
                        
                        'caso 20384
                        If grdArticulos.Rows > 2 Then
                            grdArticulos.RemoveItem (grdArticulos.Row)
                        Else
                            pLimpiaGrid
                            pConfiguraGridCargos
                        End If
                        'caso 20384
                        
                        txtClaveArticulo.Text = ""
                        pCalculaTotales
                        Exit Sub
                    End If
                    'Caso 20370
                      
                    'MANEJO DE CADUCIDADES POR MEDIO DEL LOTE
                    'LMM 20274
                     Set rsDepto = frsRegresaRs(" select INTNUMALMACEN from  pvAlmacenes WHERE INTNUMDEPARTAMENTO = " & vgintNumeroDepartamento)
                     If rsDepto.RecordCount > 0 Then
                         vllngdeptolote = rsDepto!intnumalmacen
                     End If
                    If Len(Trim(vlstrCveArticulo)) > 0 Then
                        If fblnManejaLotes(vlstrCveArticulo) And vllngdeptolote > 0 Then
                            Set rslotes = frsEjecuta_SP(vlstrCveArticulo & "|'*'|" & "-1" & "|" & vllngdeptolote, "SP_IVSELARTICULOSLOTES", , , , True)
                            If rslotes.RecordCount <> 0 Then
                                'Manejos de caducidades
                                pManejaCaducidadSal (vllngClaveElemento), "S", vlstrCveArticulo, vlintModoDescuentoInventario
                                vlBlnManejaLotes = True 'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
                                'Si no selecciono lote se remueve la fila agregada
                                If vgblnCapturoLoteYCaduc = False Then
                                    If grdArticulos.Rows > 2 Then
                                        grdArticulos.RemoveItem grdArticulos.Rows - 1
                                    Else
                                        pLimpiaGrid
                                    End If
                                 Exit Sub
                                End If
                            End If
                        End If
                    End If
                    pCalculaTotales
                Else
                    ' Consultar precios
                    txtConsultaPrecio.Text = ""
                    txtConsultaDescuento.Text = ""
                    txtConsultaDescripcion.Text = ""
                    txtConsultaDescripcion.Text = Trim(rs!Descripcion)
                    txtConsultaPrecio.Text = Format(vldblPrecio + (vldblPrecio * vldblIEPSPercent), "$###,###,###,###.00")
                    txtConsultaDescuento.Text = Format(vldblTotDescuento, "$###,###,###,###.00")
                    frePrecios.Top = 0
                    frePrecios.Left = 2800
                    frePrecios.Visible = True
                    frePrecios.ZOrder 0
                    vgstrEstadoManto = "P"
                    pBuscarPrecio True
                End If
                rs.Close
            Else
                MsgBox SIHOMsg(3), vbCritical, "Mensaje"
            End If
            vlintCantidad = 1
            txtClaveArticulo.Text = ""
            
            ' Deshabilitar el frame de pedida de datos
            If grdArticulos.RowData(1) <> -1 Then
                freDatosPaciente.Enabled = False
                If txtMovimientoPaciente.Text = "" Then
                    optPaciente.Value = True
                    lblPaciente.Caption = "VENTA AL PÚBLICO"
                    vgstrRFCPersonaSeleccionada = ""
                    vlStrTipoPaciente = "Select vchDescripcion from adtipopaciente where tnycvetipopaciente = " & vgintTipoPaciente
                    Set rsTipoPaciente = frsRegresaRs(vlStrTipoPaciente, adLockReadOnly, adOpenForwardOnly)
                    
                    If rsTipoPaciente.RecordCount > 0 Then
                        lblEmpresa.Caption = rsTipoPaciente!VCHDESCRIPCION
                    Else
                        lblEmpresa.Caption = "PARTICULAR"
                    End If
                End If
                lblBusquedaMedicos.FontBold = False
            lblBusquedaEmpleados.FontBold = False
            lblBusquedaPacientes.FontBold = False
            End If
        End With
        
    End If
    
    'se deja el totalizar en esta parte por si hay algun brinco no controlado
    pCalculaTotales
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaElemento"))
    
End Sub

'Private Sub pManejaCaducidadsalida(vllngClaveElemento As String)
Public Sub pManejaCaducidadSal(lngNumCargo As Long, vlstrTipoMov As String, vlstrCveArticulo As String, vlintUnidad As Integer)
On Error GoTo NotificaError

    Dim vllngContenido As Long
    Dim vlstrDescUnidadAlterna As String
    Dim vlstrDescUnidadMinima As String
    Dim rs As ADODB.Recordset
    Dim rslotes As New ADODB.Recordset
    Dim nombrearticulo As String
    Dim rsDepto As ADODB.Recordset
    
    'MANEJO DE CADUCIDADES POR MEDIO DEL LOTE
    'LMM 20274
     Set rsDepto = frsRegresaRs(" select INTNUMALMACEN from  pvAlmacenes WHERE INTNUMDEPARTAMENTO = " & vgintNumeroDepartamento)
     If rsDepto.RecordCount > 0 Then
         vllngdeptolote = rsDepto!intnumalmacen
     End If
    
    If fblnManejaLotes(vlstrCveArticulo) And vllngdeptolote > 0 Then
        Set rslotes = frsEjecuta_SP(vlstrCveArticulo & "|'*'|" & "-1" & "|" & vllngdeptolote, "SP_IVSELARTICULOSLOTES", , , , True)
        If rslotes.RecordCount <> 0 Then
            vgblnCapturoLoteYCaduc = False
    
            Set rs = frsRegresaRs(" SELECT IvArticulo.vchnombrecomercial, intContenido, intIdArticulo, IA.VCHDESCRIPCION ALTERNA ,UM.VCHDESCRIPCION MINIMA" & _
                                " FROM IvArticulo " & _
                                " LEFT OUTER JOIN IvUnidadVentA IA ON IA.INTCVEUNIDADVENTA = IvArticulo.INTCVEUNIALTERNAVTA " & _
                                " LEFT OUTER JOIN IvUnidadVenta UM ON UM.INTCVEUNIDADVENTA = IvArticulo.INTCVEUNIMINIMAVTA " & _
                                " WHERE chrCveArticulo = '" & vlstrCveArticulo & "' ")
            If rs.RecordCount > 0 Then
                vllngContenido = rs!intContenido
                vlstrDescUnidadAlterna = rs!alterna
                vlstrDescUnidadMinima = rs!MINIMA
                nombrearticulo = rs!VCHNOMBRECOMERCIAL
            End If
                        
            frmCaptSalidaLotePV.vlstrTablaReferencias = "SVP"
            frmCaptSalidaLotePV.vlstrTablaRefInicial = ""
            frmCaptSalidaLotePV.vlstrChrCveArticulo = vlstrCveArticulo
            frmCaptSalidaLotePV.vlstrNoMovimiento = Val(lngNumCargo)
            frmCaptSalidaLotePV.vllngNumHojaConsumo = 0
            frmCaptSalidaLotePV.vlstrTipoMovimiento = IIf(vlintUnidad = 1, "UM", "UV")
            frmCaptSalidaLotePV.vlintContenidoArt = flngContenido(vlstrCveArticulo)
            frmCaptSalidaLotePV.vlStrTitUM = StrConv(vlstrDescUnidadMinima, vbProperCase)
            frmCaptSalidaLotePV.vlStrTitUV = StrConv(vlstrDescUnidadAlterna, vbProperCase)
            frmCaptSalidaLotePV.txtDescripcionLgaArt.Text = nombrearticulo
            frmCaptSalidaLotePV.txtCantDevol.Text = txtCantidad.Text
            frmCaptSalidaLotePV.txtTotalADevolver.Text = txtCantidad.Text
            frmCaptSalidaLotePV.vlblnSoloMoviento = False
            frmCaptSalidaLotePV.vlintNoDepartamento = vllngdeptolote 'vglngCveAlmacenGeneral
           ' frmCaptSalidaLote.vlblnPermiteCambioUnidad = False 'IIf(grdManejo.TextMatrix(grdManejo.Row, intColUnidad) = "A", True, False)
            
            If vlblnBitTrazabilidad And vlstrCveArticuloTraza <> "" Then
                'Limpiamos el codigo de barras, si viene llena la variable le asignara el valor y asi no se queda el dato guardado
                frmCaptSalidaLotePV.txtCodigoBarras = ""
                frmCaptSalidaLotePV.vlblnCargada = False
                frmCaptSalidaLotePV.chkCodigoBarras.Value = 1
                frmCaptSalidaLotePV.txtCodigoBarras.Text = "" ' txtCodigoBarras.Text
                frmCaptSalidaLotePV.vlblnGuardarTraza = True
            Else
                frmCaptSalidaLotePV.txtCodigoBarras = ""
                frmCaptSalidaLotePV.vlblnCargada = False
                 frmCaptSalidaLotePV.chkCodigoBarras.Value = 0
            End If
            
            Load frmCaptSalidaLotePV
            frmCaptSalidaLotePV.Show vbModal, Me
            frmCaptSalidaLotePV.vlblnGuardarTraza = False
            rs.Close
            
            If frmCaptSalidaLotePV.vlblnEncontrado = False And vlstrCveArticuloTraza <> "" Then
               ' grdManejo.TextMatrix(grdManejo.Row, intColCantidadASurtir) = 0
                MsgBox "No se encontró el lote de la etiqueta capturada.", vbExclamation + vbOKOnly, "Mensaje"
                Exit Sub
            End If
            
            'If vgblnForzarLoteYCaduc And Not vgblnCapturoLoteYCaduc Then
               ' grdManejo.TextMatrix(grdManejo.Row, intColCantidadASurtir) = Trim(str(txtCantidad.Text))
               ' Exit Sub
            'End If
                        
'            If lblnCodigoBarras Then chkCodigoBarras.SetFocus
'            pSelTextBox txtCodigoBarras
'            txtCodigoBarras.Text = ""
       End If
        rslotes.Close
        
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pManejaCaducidad"))
    Unload Me
End Sub

Public Function fblnManejaLotes(vlstrCveArticulo As String) As Long
On Error GoTo NotificaError

    fblnManejaLotes = frsRegresaRs("SELECT bitmanejacaducidad FROM IVARTICULO WHERE chrCveArticulo = '" & Trim(vlstrCveArticulo) & "'").Fields(0)

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " :modProcedimientos " & ":fblnManejaLotes"))
End Function
Private Function fblnCuentaAbierta(vllngNumeroPaciente As Long) As Boolean
    '--------------------------------------------------------
    ' Función para determinar si un numero de paciente
    ' tiene una cuenta abierta
    ' Buscar una cuenta abierta
    '--------------------------------------------------------
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset

    vlstrSentencia = "select count(intnumPaciente) Abierta from RegistroExterno" & _
                    " Where dtmFechaEgreso Is Null " & _
                    " and intNumPaciente  = " & Trim(str(vllngNumeroPaciente))
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    fblnCuentaAbierta = rs!Abierta > 0
    rs.Close

End Function

Private Sub pNuevo()
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    '-----------------------------------
    ' Limpieza de campos
    '-----------------------------------
    txtMovimientoPaciente.Text = ""
    lblPaciente.Caption = ""
    vgstrRFCPersonaSeleccionada = ""
    lblEmpresa.Caption = ""
    vgintTipoPaciente = 0
    vgintEmpresa = 0
    txtCantidad.Text = 1
    pLimpiaGrid
    txtImporte.Text = ""
    txtDescuentos.Text = ""
    txtDescuento.Text = ""
    txtSubtotal.Text = ""
    txtIEPS.Text = ""
    txtIva.Text = ""
    txtTotal.Text = ""
    pConfiguraGridCargos
    freDatosPaciente.Enabled = True
    vgstrEstadoManto = ""
    
    vlblnImportoVenta = False
    cmdImportarVenta.Enabled = True
    txtClaveArticulo.Enabled = True
    txtCantidad.Enabled = True
    optArticulo.Enabled = True
    optOtrosConceptos.Enabled = True
    
    freGraba.Enabled = True
    
    freConsultaPrecios.Enabled = True
    freDescuentos.Enabled = True
    
    'Frame del control de seguros:
    lblnGuardoControl = False
    lblnExisteControl = False
    optTipoDeducible(0).Value = True
    optTipoCoaseguro(0).Value = True
    optCopagoCantidad.Value = True
    txtDeducible.Text = ""
    txtCoaseguro.Text = ""
    txtCopago.Text = ""
    chkFacturarDeducible.Value = 0
    chkFacturarCoaseguro.Value = 0
    chkFacturarCopago.Value = 0
    ldblDeducible = 0
    ldblCoaseguro = 0
    ldblCopago = 0
    
    txtDeducible_Change
    txtCoaseguro_Change
    txtCopago_Change
    
    optPaciente.Value = True
    intCuentaNueva = 0
    lblBusquedaMedicos.FontBold = True
    lblBusquedaEmpleados.FontBold = True
    lblBusquedaPacientes.FontBold = True
    
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   '|                   VERIFICACIÓN DE FOLIOS                  |
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
   
    'Verificación de folios de factura
    vllngFoliosRestantes = 1
    vlstrFolioDocumento = ""
    pCargaArreglo vlaryParametrosSalida, vllngFoliosRestantes & "|" & adInteger & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacionPaciente & "|" & ADODB.adBSTR & "|" & strAnoAprobacionPaciente & "|" & ADODB.adBSTR
    frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|0", "Sp_GnFolios", , , vlaryParametrosSalida
    pObtieneValores vlaryParametrosSalida, vllngFoliosRestantes, strFolio, strSerie, strNumeroAprobacionPaciente, strAnoAprobacionPaciente
     '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
    vlstrFolioDocumento = Trim(strSerie) & strFolio
    
    If vllngFoliosRestantes > 0 Then
        MsgBox "Faltan " & Trim(str(vllngFoliosRestantes)) + " facturas y será necesario aumentar folios!", vbOKOnly + vbInformation, "Mensaje"
    End If
    lblFactura.Caption = vlstrFolioDocumento
    If Trim(lblFactura.Caption) = "0" Then
        'No existen folios activos para este documento.
        MsgBox SIHOMsg(291) + " Facturas.", vbCritical + vbOKOnly, "Mensaje"
    End If
    
'    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
      
    '-----------------------
    'El Tipo de paciente por default, para venta al público
    '-----------------------
'    vlstrSentencia = "select IntTipoParticular from parametros"
    vlstrSentencia = "select distinct vchValor from SiParametro where trim(VCHNOMBRE) = 'INTTIPOPARTICULAR'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    vgintTipoPaciente = rs!VCHVALOR

    vlblnConsultaTicketPrevio = False
    freDatosPaciente.Enabled = True
    frmMedico.Enabled = True
    
    chkFacturaSustitutaDFP.Enabled = False
    lstFacturaASustituirDFP.Enabled = False
    
    lstFacturaASustituirDFP.Clear
    chkFacturaSustitutaDFP.Value = 0
    
    vlnblnLocate = False
    
    pEtiquetaVar "", ""
    cmdDescuentoPuntos.Caption = "Aplicar puntos de cliente leal"
    dblDescuentoAplicadoPuntos = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not vlbsinalmacen = True Then
    If vgstrEstadoManto = "B" Then        'Busqueda de articulos
        Cancel = 1
        pBuscaArticulos False
    Else
        If vgstrEstadoManto = "P" Then   'Busqueda de precios
            Cancel = 1
            pBuscarPrecio False
        ElseIf freControlAseguradora.Visible Then
            Cancel = 1
            cmdCierraControlAseguradora_Click
        Else
            If grdArticulos.RowData(1) <> -1 Then
                Cancel = 1
                If MsgBox(SIHOMsg(9), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                    If Trim(txtMovimientoPaciente.Text) <> "" Then
        
                        If cmdDescuentoPuntos.Caption = "Desaplicar puntos de cliente leal" Then
                            MsgBox "Es necesario desaplicar los puntos de cliente leal antes de salir de la pantalla.", vbExclamation, "Mensaje"
                            cmdDescuentoPuntos.SetFocus
                            Exit Sub
                        End If
                        
                        'Revisa si la cuenta tiene cargos
                        lngCargos = 1
                        frsEjecuta_SP Trim(txtMovimientoPaciente.Text), "sp_PvSelNumCargosCuenta", False, lngCargos
                        
                        If lngCargos = 0 Then
                            If intCuentaNueva = 0 Then
                                pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                            ElseIf intCuentaNueva = 1 Then
                                pBorrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                            End If
                        End If
                        
                    End If
                    
                    pNuevo
                    txtClaveArticulo.Text = ""
                    If txtClaveArticulo.Enabled Then
                        txtClaveArticulo.SetFocus
                    End If
                    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
                    frmCaptSalidaLotePV.BorrarLote ' caso 20417
                End If
            Else
                If vgstrEstadoManto = "V" Then  'Estatus de Venta, osea que ya creo el número de cuenta
                    
                    If cmdDescuentoPuntos.Caption = "Desaplicar puntos de cliente leal" Then
                        MsgBox "Es necesario desaplicar los puntos de cliente leal antes de salir de la pantalla.", vbExclamation, "Mensaje"
                        cmdDescuentoPuntos.SetFocus
                        Exit Sub
                    End If
                    
                    Cancel = 1
        
                    'Revisa si la cuenta tiene cargos
                    If Trim(txtMovimientoPaciente.Text) <> "" Then
                        lngCargos = 1
                        frsEjecuta_SP Trim(txtMovimientoPaciente.Text), "sp_PvSelNumCargosCuenta", False, lngCargos
                        
                        If lngCargos = 0 Then
                            If intCuentaNueva = 0 Then
                                pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                            ElseIf intCuentaNueva = 1 Then
                                pBorrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                            End If
                        End If
                    End If
                    
                    pNuevo
                    txtClaveArticulo.Text = ""
                    If txtClaveArticulo.Enabled Then
                        txtClaveArticulo.SetFocus
                    End If
                    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
                ElseIf vgstrEstadoManto = "" Then 'Ya creó numero de cuenta pero no se realiza la venta
                        If cmdDescuentoPuntos.Caption = "Desaplicar puntos de cliente leal" Then
                            MsgBox "Es necesario desaplicar los puntos de cliente leal antes de salir de la pantalla.", vbExclamation, "Mensaje"
                            cmdDescuentoPuntos.SetFocus
                            Exit Sub
                        End If
                        
                        If intCuentaNueva = 0 Then
                        pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                     ElseIf intCuentaNueva = 1 Then
                        pBorrarCuenta CLng(Val(txtMovimientoPaciente.Text))
                     End If
                End If
            End If
        End If
    End If
Else
    Unload Me
End If
End Sub
Private Sub pBorrarCuenta(vllngNumeroCuenta As Long)
'Si no había cuenta existente y abrió una pero se cancela la venta, entonces borra la cuenta
    EntornoSIHO.ConeccionSIHO.BeginTrans
        frsEjecuta_SP Trim(txtMovimientoPaciente.Text) & "|E", "SP_EXDELCUENTAPACIENTE"
    EntornoSIHO.ConeccionSIHO.CommitTrans
End Sub

Private Sub pCerrarCuenta(lngCveCuenta As Long)
'Si ya hay una cuenta existente, solo cierra la cuenta, no la borra
Dim rsCuenta As New ADODB.Recordset
Dim blnCierraCuenta As Boolean
Dim strTipoCuenta As String

'-----------------------------------'
' Actualizar el estado de la cuenta '
'-----------------------------------'
vlstrSentencia = "SELECT * FROM Expacienteingreso where expacienteingreso.INTNUMCUENTA = " & lngCveCuenta
Set rsCuenta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
If rsCuenta.RecordCount <> 0 Then
    strTipoCuenta = rsCuenta!chrTipoIngreso
    blnCierraCuenta = True
    
        If blnCierraCuenta = True Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
                vgstrParametrosSP = lngCveCuenta & "|" & _
                                    Trim(strTipoCuenta) & "|1|" & _
                                    IIf(Trim(strTipoCuenta) = "E", fstrFechaSQL(fdtmServerFecha, fdtmServerHora), Null)
                frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDCERRARABRIRCUENTA"
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
Else
    Exit Sub 'No existe esa cuenta
End If
End Sub

Private Sub pBuscarPrecio(vlblnbandera As Boolean)
    If vlblnbandera Then
        freDatosPaciente.Enabled = False
        FreDetalle.Enabled = False
        frmMedico.Enabled = False
        lblBusquedaMedicos.FontBold = False
        lblBusquedaEmpleados.FontBold = False
        lblBusquedaPacientes.FontBold = False
    Else
        FreDetalle.Enabled = True
        frmMedico.Enabled = True
        freDatosPaciente.Enabled = True
        lblBusquedaMedicos.FontBold = True
        lblBusquedaEmpleados.FontBold = True
        lblBusquedaPacientes.FontBold = True
    
        pEnfocaTextBox txtClaveArticulo
        frePrecios.Visible = False
        chkConsultaPrecios.Value = 0
        pEnfocaTextBox txtClaveArticulo
        vgstrEstadoManto = ""
    End If
End Sub

Private Sub grdArticulos_DblClick()
    With grdArticulos
        If .RowData(1) <> -1 Then
            .TextMatrix(.Row, 0) = IIf(.TextMatrix(.Row, 0) = "*", "", "*")
            .Col = 0
            .CellFontBold = True
            .CellFontSize = 12
            .Col = 1
        End If
    End With
End Sub

Private Sub pCalculaTotales()
    Dim vlintContador As Integer
    Dim vldblImporte As Double
    Dim vldblDescuentos As Double
    Dim vldblTotalIVA As Double
    Dim vlintConta2 As Integer
    Dim vldblIEPS As Double
    Dim vldblImporteGrd As Double
    
    
    vldblImporte = 0
    vldblDescuentos = 0
    vldblTotalIVA = 0
    vldblIEPS = 0

    
    With grdArticulos
        
        If vgblnTrazabilidad Then
            If vlintContador > 0 Then
                If .TextMatrix(vlintContador, 4) <> "" Then
                    For vlintContador = 1 To .Rows - 1
                         vldblImporteGrd = CDbl(.TextMatrix(vlintContador, 4))
                        .TextMatrix(vlintContador, 7) = Format(Val(.TextMatrix(vlintContador, 3)) * vldblImporteGrd, "$###,###,###,###.00")
                        .TextMatrix(vlintContador, 9) = Format(Val(.TextMatrix(vlintContador, 3)) * vldblImporteGrd, "$###,###,###,###.00")
                    Next vlintContador
                End If
            End If
        End If
    
        If .RowData(1) <> -1 Then
        
            cmdImportarVenta.Enabled = False
            
            If vlblnImportoVenta = True Then
                txtClaveArticulo.Enabled = False
                txtCantidad.Enabled = False
                optArticulo.Enabled = False
                optOtrosConceptos.Enabled = False
                            
                freConsultaPrecios.Enabled = False
                freDescuentos.Enabled = False
                
                freDatosPaciente.Enabled = False
                
                freGraba.Enabled = True
                
                optPaciente.Value = True
                optMedico.Value = False
                optEmpleado.Value = False
                txtMovimientoPaciente.Text = ""
                txtGafete.Text = ""
                lblPaciente.Caption = ""
                lblEmpresa.Caption = ""
            End If
        
            For vlintContador = 1 To .Rows - 1
                vldblImporte = vldblImporte + IIf(vgblnTrazabilidad = True, Format(Val(.TextMatrix(vlintContador, 3)) * CDbl(.TextMatrix(vlintContador, 2)), ""), Val(Format(.TextMatrix(vlintContador, 4), "")))
                vldblDescuentos = vldblDescuentos + Val(Format(.TextMatrix(vlintContador, 5), ""))
                vldblTotalIVA = vldblTotalIVA + Round(Val(.TextMatrix(vlintContador, 11)), 2)
                .Row = vlintContador
                
                If vlblnLicenciaIEPS Then vldblIEPS = vldblIEPS + Val(Format(.TextMatrix(vlintContador, 6), "")) '@
                              
                For vlintConta2 = 1 To .Cols - 1
                    .Col = vlintConta2
                    .CellBackColor = IIf(.TextMatrix(vlintContador, 14) = 1, &H80FFFF, &H80000005)
                    .CellForeColor = IIf(.TextMatrix(vlintContador, 14) = 1, &HFF0000, &H80000012)
                Next
            Next
        Else
            cmdImportarVenta.Enabled = True
            vlblnImportoVenta = False
            
            txtClaveArticulo.Enabled = True
            txtCantidad.Enabled = True
            optArticulo.Enabled = True
            optOtrosConceptos.Enabled = True
                        
            freConsultaPrecios.Enabled = True
            freDescuentos.Enabled = True
            
            freDatosPaciente.Enabled = True
                
            freGraba.Enabled = True
            
            freFacturaPesosDolares.Enabled = True
        End If
        
    txtImporte.Text = Format(vldblImporte, "$###,###,###,###.00")
    txtDescuentos.Text = Format(vldblDescuentos, "$###,###,###.00")
    txtSubtotal.Text = Format(Round(vldblImporte, 2) - Round(vldblDescuentos, 2) + Round(vldblIEPS, 2), "$###,###,###,###.00")
    txtIEPS.Text = Format(vldblIEPS, "$###,###,###,###.00")
    txtIva.Text = Format(vldblTotalIVA, "$###,###,###,###.00")
    txtTotal.Text = Format(Round(vldblImporte, 2) - Round(vldblDescuentos, 2) + Round(vldblIEPS, 2) + Round(vldblTotalIVA, 2), "$###,###,###,###.00")
    End With
End Sub

Private Function fdblObtenerIva(vllngCveCargo As Long, vlstrTipoCargo As String) As Variant
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
        
        Select Case vlstrTipoCargo
        Case "OC"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM PvOtroConcepto INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "PvOtroConcepto.smiConceptoFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where PvOtroConcepto.intCveConcepto = " & RTrim(str(vllngCveCargo))
        Case "AR"
            vlstrSentencia = "SELECT PvConceptoFacturacion.smyIva Iva " & _
                        "FROM IvArticulo INNER JOIN " & _
                        "PvConceptoFacturacion ON " & _
                        "IvArticulo.smiCveConceptFact = PvConceptoFacturacion.smiCveConcepto " & _
                        "where IvArticulo.intIdArticulo = " & RTrim(str(vllngCveCargo))
        End Select
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    fdblObtenerIva = rs!IVA
    rs.Close
End Function
Private Function fdblObtenerPorcentajeIEPS(vllngCveCargo As Long) As Variant
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
        
    vlstrSentencia = "select nvl(cnieps.RELPORCENTAJE,0)" & _
                       " From ivarticulo" & _
                       " left join cnieps on cnieps.SMICVEIEPS = ivarticulo.INTCVEIEPS" & _
                       " where ivarticulo.INTIDARTICULO=" & CStr(vllngCveCargo) & ""
    
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    fdblObtenerPorcentajeIEPS = IIf(rs.RecordCount = 0, 0, rs.Fields(0))
    rs.Close
End Function


Private Sub Form_Activate()
    
    vlbsinalmacen = False
    Call pVerificaAlmacen
    
    Dim vlrsAux As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vllngMensaje As Long

    vlstrSentencia = "SELECT PVPARAMETROVENTADEFAULTXDEPTO.SMIVENTADEFAULT FROM PVPARAMETROVENTADEFAULTXDEPTO WHERE PVPARAMETROVENTADEFAULTXDEPTO.SMICVEDEPARTAMENTO = " & vgintNumeroDepartamento
    Set vlrsAux = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If vlblnPrimeraVez Then
        If Not vlrsAux.EOF Then
            optArticulo.Value = IIf(vlrsAux.Fields(0).Value, False, True)
            optOtrosConceptos.Value = IIf(vlrsAux.Fields(0).Value, True, False)
        Else
            optArticulo.Value = True
            optOtrosConceptos.Value = False
        End If
    End If
    
    lblFecha.Caption = Format(fdtmServerFecha, "Long Date")
    lblTipoCambio.Caption = "Dólar " & Trim(str(fdblTipoCambio(fdtmServerFecha, "V")))
    lblHora.Caption = Mid(Time(), 1, Len(Time()) - 3)
    
    
    vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    
    
    
    If vllngMensaje <> 0 Then
        'Cierre el corte actual antes de registrar este documento.
        MsgBox SIHOMsg(str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje" 'No existe un corte abierto
        Unload Me
    End If
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If frmFoliosVentaImportada.Visible Then
            FreDetalle.Enabled = True
            grdArticulos.SetFocus
            freDatosPaciente.Enabled = True
            frmMedico.Enabled = True
            freControlAseguradora.Enabled = True
            freBusqueda.Enabled = True
            
            frmFoliosVentaImportada.Visible = False
            
            If vlblnImportoVenta = False Then
                txtClaveArticulo.Enabled = True
                txtCantidad.Enabled = True
                optArticulo.Enabled = True
                optOtrosConceptos.Enabled = True
                            
                freConsultaPrecios.Enabled = True
                freDescuentos.Enabled = True
                
                cmdImportarVenta.Enabled = True
                                
                freGraba.Enabled = True
                
                freFacturaPesosDolares.Enabled = True
            End If
            
            xlsApp.Quit
            Set xlsApp = Nothing
            
            f.Close
            
            Exit Sub
        End If
    
        KeyAscii = 0
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Dim vlrsAux As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsDatos As New ADODB.Recordset
    Dim vgstrSentencia As String
    
    'Caso 20507 [Modifico/Agrego: GIRM | Fecha:  04/09/2024 ]
    vlBlnManejaLotes = False

    vlbflag = False
    Me.Icon = frmMenuPrincipal.Icon
    
    'Validar licencia para generar la lealtad del cliente y el médico con el hospital por medio del otorgamiento de puntos
    blnLicenciaLealtadCliente = fblnLicenciaLealtadCliente
    If blnLicenciaLealtadCliente Then
        cmdDescuentoPuntos.Visible = True
    End If
    
    vlblnLicenciaIEPS = fblLicenciaIEPS '<-------
    pPreparaIEPS
    
    pInstanciaReporte vgrptReporte, "rptReciboPagoPos.rpt"
    vlblnPrimeraVez = True
'    vlstrSentencia = "select intTipoParticular from Parametros"
    vlstrSentencia = "select distinct vchValor from SiParametro where trim(VCHNOMBRE) = 'INTTIPOPARTICULAR'"
    Set vlrsAux = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    vgintTipoParticular = vlrsAux.Fields(0).Value 'Para saber cual es el TipoPacienteParticular
    vlrsAux.Close
    
    '--------------------------------------------------------------
    '-Identificar el parametro de tipo de corte en el departamento
    '--------------------------------------------------------------
    vlstrSentencia = "Select * " & _
            " From GnParametroTipoCorte " & _
            " Where intCveDepartamento = " & vgintNumeroDepartamento
    vlblnParametroTipoCorte = False
    Set rsGnParametrosTipoCorte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsGnParametrosTipoCorte.RecordCount > 0 Then
        vlblnParametroTipoCorte = True
    End If
        
    pLimpiaGrid
    pConfiguraGridCargos
    
    txtTicketPrevio.Text = fSiguienteClaveVentaPublicoNoFacturado
    
    '--------------------------------------------------------
    ' Parametro de si tiene instalada una impresora serial
    '--------------------------------------------------------
    lblnImpresoraSerial = False
    vlstrSentencia = "select count(*) from PvLoginImpresoraTicket where intNumeroLogin = " & str(vglngNumeroLogin)
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsTemp.Fields(0) <> 0 Then
        lblnImpresoraSerial = True
    End If
    rsTemp.Close
    
    '--------------------------------------------------------
    ' Traer Leyenda de informacion al cliente y el valor del bit para validar la venta de medicamentos sin revasar el Precio Máximo al Público
    '--------------------------------------------------------
    vlstrSentencia = "select vchleyendacliente, BitValidacionPMPVentaPublic from pvparametro where tnyclaveempresa = " & vgintClaveEmpresaContable
    Set rsTemp = frsRegresaRs(vlstrSentencia)
    lstrLeyendaCliente = IIf(IsNull(rsTemp!vchleyendacliente), "", rsTemp!vchleyendacliente)
    vlintValorPMP = IIf(IsNull(rsTemp!BitValidacionPMPVentaPublic), 0, rsTemp!BitValidacionPMPVentaPublic)
    rsTemp.Close
    
    'Parámetros de caja necesarios
    Set rsTemp = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "Sp_PvSelParametro")
    If rsTemp.RecordCount > 0 Then
        With rsTemp
            vllngMedicoDefaultPOS = 0
            lngTipoPacMedico = IIf(IsNull(!intTipoPacMedico), 0, !intTipoPacMedico)
            lngTipoPacEmpleado = IIf(IsNull(!intTipoPacEmpleado), 0, !intTipoPacEmpleado)
            intFacturarVentaPublico = IIf(IsNull(!bitFacturarVentaPublico), 0, !bitFacturarVentaPublico)
            strNombrePOS = IIf(IsNull(!chrNombreFacturaPOS), "", Trim(!chrNombreFacturaPOS))
            strCallePOS = IIf(IsNull(!chrDireccionPOS), "", Trim(!chrDireccionPOS))
            strNumeroExteriorPOS = IIf(IsNull(!vchNumeroExteriorPOS), "", Trim(!vchNumeroExteriorPOS))
            strNumeroInteriorPOS = IIf(IsNull(!vchNumeroInteriorPOS), "", Trim(!vchNumeroInteriorPOS))
            strColoniaPOS = IIf(IsNull(!vchColoniaPOS), "", Trim(!vchColoniaPOS))
            strCPPOS = IIf(IsNull(!vchCodigoPostalPOS), "", Trim(!vchCodigoPostalPOS))
            lngCveCiudadPOS = IIf(IsNull(!intCveCiudad), 0, !intCveCiudad)
        End With
    Else
        vllngMedicoDefaultPOS = 0
        lngTipoPacMedico = 0
        lngTipoPacEmpleado = 0
        intFacturarVentaPublico = 0
    End If
    rsTemp.Close
    
    pNuevo
    
    fraTipoDeducible.BorderStyle = 0
    fraTipoCoaseguro.BorderStyle = 0
    fraTipoCopago.BorderStyle = 0
    
    vgstrSentencia = "Select intCveMedico Clave, ltrim(rtrim(vchApellidoPaterno))||' '||ltrim(rtrim(vchApellidoMaterno))||' '||ltrim(rtrim(vchNombre)) Nombre " & _
    "From HoMedico Where bitEstaActivo = 1 Order by Nombre"
    Set rsDatos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsDatos.RecordCount > 0 Then
        pLlenarCboRs cboMedico, rsDatos, 0, 1
        cboMedico.AddItem "NINGUNO", 0
        cboMedico.ListIndex = 0
    End If
    rsDatos.Close
    
    Me.txtIEPS.Enabled = vlblnLicenciaIEPS '<----
    Me.lblIEPS.Enabled = vlblnLicenciaIEPS '<----
    
    vlblnConsultaTicketPrevio = False
    
    vgblnTrazabilidad = fblnrevisaUsoTrazabilidad()
    
End Sub
Private Sub pPreparaIEPS()
If Not vlblnLicenciaIEPS Then ' no hay licencia IEPS ajustamos la pantalla
    If blnLicenciaLealtadCliente Then
        Me.lblIEPS.Visible = False
        Me.txtIEPS.Visible = False
'        Me.Label17.Top = 1305
'        Me.txtSubtotal.Top = 1215
'        Me.Label10.Top = 1770
'        Me.txtIva.Top = 1680
'        Me.Label11.Top = 2235
'        Me.txtTotal.Top = 2145

        Me.Label13.Top = 615
        Me.txtImporte.Top = 525
        Me.Label12.Top = 1065
        Me.txtDescuentos.Top = 975
        
        Me.Label17.Top = 1510
        Me.txtSubtotal.Top = 1425
        Me.Label10.Top = 1980
        Me.txtIva.Top = 1890
        Me.Label11.Top = 2445
        Me.txtTotal.Top = 2355

'        Me.cmdDescuentoPuntos.Left = 8810
'        Me.cmdDescuentoPuntos.Top = 6300

        'Me.FreTotales.Height = 2780
        'Me.freGraba.Top = 5460
        'Me.FreDetalle.Height = 6480
        'Me.Frame8.Top = 4630
        'Me.Height = 8940
        'Me.Refresh
    Else
        Me.lblIEPS.Visible = False
        Me.txtIEPS.Visible = False
        Me.Label17.Top = 1305
        Me.txtSubtotal.Top = 1215
        Me.Label10.Top = 1770
        Me.txtIva.Top = 1680
        Me.Label11.Top = 2235
        Me.txtTotal.Top = 2145
        Me.FreTotales.Height = 2800
        Me.freGraba.Top = 5460
        Me.FreDetalle.Height = 6480
        Me.Frame8.Top = 4630
        Me.Height = 8940
        Me.Refresh
    End If
End If
End Sub
Private Sub grdArticulos_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        With grdArticulos
            If .Rows > 2 Then
                pBorrarRegMshFGrdData grdArticulos.Row, grdArticulos
                grdArticulos.Row = 1
                grdArticulos.Col = 1
                grdArticulos.SetFocus
            Else
                pLimpiaGrid
                .Rows = 2
                pConfiguraGridCargos
                
                If txtClaveArticulo.Enabled Then
                    txtClaveArticulo.SetFocus
                End If
            End If
            pCalculaTotales
        End With
    ElseIf KeyCode = vbKeyEnd Then
        cmdSave_Click
    ElseIf KeyCode = vbKeyReturn And Not vl_dblClickCmdDescuenta Then
        grdArticulos_DblClick
    End If
    vl_dblClickCmdDescuenta = False
End Sub

Private Sub lstBuscaArticulos_DblClick()
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    pBuscaArticulos False
    frmPOS.Refresh
    pSeleccionaElemento lstBuscaArticulos.ItemData(lstBuscaArticulos.ListIndex), CLng(Val(txtCantidad.Text)), IIf(optArticulo.Value, "AR", "OC")
    
    txtCantidad.Text = 1
    Exit Sub
End Sub

Private Sub lstBuscaArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then lstBuscaArticulos_DblClick
End Sub

Private Sub txtBuscaArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown) And lstBuscaArticulos.Enabled Then
        lstBuscaArticulos.SetFocus
    End If
End Sub

Private Sub txtBuscaArticulo_KeyPress(KeyAscii As Integer)
    '------------------------------------
    ' Validar el apostrofe porque truena cuando lo pones una comilla sencilla ( ' )
    '------------------------------------
    If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub txtBuscaArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
        
    If optArticulo.Value Then
        vlstrSentencia = "Select intidArticulo Clave, vchNombreComercial Descripcion from ivArticulo"
        PSuperBusqueda txtBuscaArticulo, vlstrSentencia, lstBuscaArticulos, "vchNombreComercial", 100, "And IvArticulo.chrCostoGasto <> 'G' and upper(vchestatus) = 'ACTIVO'", "vchNombreComercial"
    Else
        vlstrSentencia = " SELECT * "
        vlstrSentencia = vlstrSentencia & " FROM PVOTROCONCEPTODEPTO"
        vlstrSentencia = vlstrSentencia & " WHERE PVOTROCONCEPTODEPTO.SMICVEDEPARTAMENTO = " & vgintNumeroDepartamento
        If frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic).EOF Then
            vlstrSentencia = "Select intCveConcepto Clave, chrDescripcion Descripcion from PvOtroConcepto"
        Else
            vlstrSentencia = " SELECT intCveConcepto Clave, chrDescripcion Descripcion "
            vlstrSentencia = vlstrSentencia & " FROM PVOTROCONCEPTO"
            vlstrSentencia = vlstrSentencia & "   inner join (SELECT PVOTROCONCEPTODEPTO.INTCVEOTROCONCEPTO"
            vlstrSentencia = vlstrSentencia & "                  FROM PVOTROCONCEPTODEPTO"
            vlstrSentencia = vlstrSentencia & "                  WHERE PVOTROCONCEPTODEPTO.SMICVEDEPARTAMENTO = " & vgintNumeroDepartamento & ") Temp"
            vlstrSentencia = vlstrSentencia & "                  ON PVOTROCONCEPTO.INTCVECONCEPTO = Temp.intcveOtroconcepto  "
        End If
        PSuperBusqueda txtBuscaArticulo, vlstrSentencia, lstBuscaArticulos, "chrDescripcion", 100, " AND BITESTATUS = 1 ", "chrDescripcion"
    End If
End Sub

Private Sub pLimpiaGrid()
    
    With grdArticulos
        .Clear
        .Rows = 2
        .RowData(1) = -1
    End With
    If grdArticulos.TextMatrix(0, 1) = "" Then
        pConfiguraGridCargos
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtClaveArticulo_GotFocus()
    lblMensajes.Caption = IIf(optArticulo.Value, "<F8> - Buscar artículos", "<F8> - Buscar conceptos")
    
    If Trim(vlstrFolioDocumento) = "0" Or Trim(lblFactura.Caption) = "0" Then
        Unload Me
    End If
End Sub

Private Sub txtClaveArticulo_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "*" Then
        txtCantidad.Text = IIf(Val(Mid(txtClaveArticulo.Text, 1, txtClaveArticulo.SelStart)) = 0, 1, Val(Mid(txtClaveArticulo.Text, 1, txtClaveArticulo.SelStart)))
        KeyAscii = 0
        txtClaveArticulo.Text = ""
    Else
        If optArticulo.Value Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
                KeyAscii = 7
            End If
        End If
    End If
End Sub

Private Sub txtClaveArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    cboMedico.Enabled = True
    
    If KeyCode = vbKeyReturn Then
        Select Case vgblnTrazabilidad
            Case True
                vgBuscarArticuloTrazabilidad
            Case False
                If optArticulo.Value Then
                    vlstrSentencia = "select intIdArticulo Elemento from ivArticulo " & _
                                    " Where chrCveArticulo = " & _
                                    " (Select chrCveArticulo from ivCodigoBarrasArticulo where rtrim(ltrim(vchCodigoBarras)) = '" & Replace(Trim(txtClaveArticulo.Text), "'", "''") & "'and rownum < 2)"
                Else
                    vlstrSentencia = "select intCveConcepto Elemento from PvOtroConcepto " & _
                                    " Where intCveConcepto = " & Val(Mid(Trim(txtClaveArticulo.Text), 1, 9))
                End If
                
                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
                If rs.RecordCount > 0 Then
                    pSeleccionaElemento rs!Elemento, CLng(txtCantidad.Text), IIf(optArticulo.Value, "AR", "OC")
                    txtCantidad.Text = 1
                Else
                    MsgBox SIHOMsg(13) & Chr(vbKeyReturn) & "Artículo inválido.", vbExclamation, "Mensaje"
                    pEnfocaTextBox txtClaveArticulo
                End If
                rs.Close
         End Select
     Else
         If KeyCode = vbKeyF8 Then
             If optArticulo.Value Then
                 If Not fblnRevisaPermiso(vglngNumeroLogin, 383, "C") Then
                     '---------------------------------------------------
                     'Para que sólo los logins autorizados puedan buscar
                     ' por descripción (sólo artículos)
                     ' y que pregute cuando el Login inicial no tenga permiso
                     '---------------------------------------------------
                     Load frmLogin
                     frmLogin.vgblnCargaVariablesGlobales = False
                     frmLogin.Show vbModal
                     If frmLogin.vgintLogin = -1 Or Not fblnRevisaPermiso(frmLogin.vgintLogin, 383, "C") Then
                         MsgBox SIHOMsg(635), vbInformation, "Mensaje"
                         Exit Sub
                     End If
                 End If
             End If
             pBuscaArticulos True
         End If
     End If
    ' If txtMovimientoPaciente = "" Then
         'pFacturasDirectasAnteriores
    ' End If
End Sub
Private Sub pBuscaArticulos(vlblnActiva As Boolean)
    If vlblnActiva Then
        'Mostrar la búsqueda de artículos
        freBusqueda.Top = 1200
        freBusqueda.Visible = True
        txtBuscaArticulo.SetFocus
        cboMedico.Enabled = False
        FreDetalle.Enabled = False
        freDatosPaciente.Enabled = False
        lblBusquedaMedicos.FontBold = False
        lblBusquedaEmpleados.FontBold = False
        lblBusquedaPacientes.FontBold = False
        
        txtBuscaArticulo.Text = ""
        lstBuscaArticulos.Clear
        lstBuscaArticulos.Enabled = False
        vgstrEstadoManto = "B"
    Else
        'Quitar la búsqueda de artículos
        lblMensajes.Caption = "<END> Terminar venta"
        FreDetalle.Enabled = True
        cboMedico.Enabled = True
        pEnfocaTextBox txtClaveArticulo
        freBusqueda.Visible = False
        If Trim(lblEmpresa.Caption) = "" Then
            freDatosPaciente.Enabled = True
            lblBusquedaMedicos.FontBold = True
            lblBusquedaEmpleados.FontBold = True
            lblBusquedaPacientes.FontBold = True
        End If
        vgstrEstadoManto = ""
    End If
     'If txtMovimientoPaciente = "" Then
        'pFacturasDirectasAnteriores
     'End If
  End Sub

Private Sub pConfiguraGridCargos()
    With grdArticulos
        '.Cols = 29
        .Cols = 33
        .FixedCols = 1
        .FixedRows = 1
        .SelectionMode = flexSelectionFree
        .FormatString = "|Descripción|Precio|Cantidad|Importe|Descuento|IEPS|Subtotal|IVA|Total"
        .ColWidth(0) = IIf(vlblnLicenciaIEPS, 150, 190)  'Fix
        .ColWidth(1) = IIf(vlblnLicenciaIEPS, 3270, 3330) 'Descripción
        .ColWidth(2) = IIf(vlblnLicenciaIEPS, 980, 1050) 'Precio
        .ColWidth(3) = IIf(vlblnLicenciaIEPS, 700, 960)  'Cantidad
        .ColWidth(4) = IIf(vlblnLicenciaIEPS, 980, 1050) 'Importe
        .ColWidth(5) = IIf(vlblnLicenciaIEPS, 970, 1050) 'Descuento
        .ColWidth(6) = IIf(vlblnLicenciaIEPS, 700, 0) 'IEPS<---------------------@
        .ColWidth(7) = IIf(vlblnLicenciaIEPS, 1000, 1050) 'Subtotal@
        .ColWidth(8) = IIf(vlblnLicenciaIEPS, 980, 1050) 'IVA
        .ColWidth(9) = IIf(vlblnLicenciaIEPS, 1300, 1300) 'Total
        .ColWidth(10) = 0   'Clave del cargo
        .ColWidth(11) = 0   'IVA
        .ColWidth(12) = 0   'Tipo de Cargo
        .ColWidth(13) = 0   'Concepto de Facturación
        .ColWidth(14) = 0   'Tipo de descuento del inventario
        .ColWidth(15) = 0   'Clave con la que se realizó el cargo en la tabla de cargos
        .ColWidth(16) = 0   'Contenido de IVArticulo,
        .ColWidth(17) = 0   'Clave de la lista de precios
        .ColWidth(18) = 0   'Porcentaje de descuento
        .ColWidth(19) = 0   'Porcentaje de IEPS <---------------------------------
        .ColWidth(20) = 0   'Indica si tiene exclusión de descuento
        .ColWidth(21) = 0   'Descripción del concepto de facturación
        .ColWidth(22) = 0   'Descripción del concepto de facturación
        .ColWidth(23) = 0   'Bit cambio de precio, 1 cuando hubo cambio de precio, vacio o cero si no se movio
        .ColWidth(24) = 0   'Precio unitario con todos los decimales
        .ColWidth(25) = 0
        .ColWidth(26) = 0
        .ColWidth(27) = 0   'Descuento antes de aplicar descuento por puntos de cliente leal
        .ColWidth(28) = 0   'Descuento aplicado en puntos de cliente leal
        .ColWidth(29) = 0   'Código de barras de trazabilidad
        .ColWidth(30) = 0   'Clave del artículo
        .ColWidth(31) = 0   'Lote de trazabilidad
        .ColWidth(32) = 0   'Varios Lotes
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
        .ColAlignment(10) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignLeftCenter
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
        .ColAlignmentFixed(11) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Private Sub cmdTachitaBusqueda_Click()
    pBuscaArticulos False
    cboMedico.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEnd Then
        If vgstrEstadoManto <> "B" And vgstrEstadoManto <> "P" And Not freControlAseguradora.Visible Then
            cmdSave_Click
        End If
    ElseIf KeyCode = vbKeyF5 Then
    'Consulta de precios
        If vgstrEstadoManto <> "B" And vgstrEstadoManto <> "P" Then
            chkConsultaPrecios.Value = IIf(chkConsultaPrecios.Value = 1, 0, 1)
            chkConsultaPrecios_Click
        End If
    ElseIf KeyCode = vbKeyF4 Then
    'Búsqueda de pacientes
        If vgstrEstadoManto <> "B" And vgstrEstadoManto <> "P" And vgstrEstadoManto <> "V" And Not freControlAseguradora.Visible Then
            optPaciente.Value = True
            txtMovimientoPaciente_KeyDown vbKeyReturn, 0
        End If
    ElseIf KeyCode = vbKeyF6 Then
    'Búsqueda de médicos
        If vgstrEstadoManto <> "B" And vgstrEstadoManto <> "P" And vgstrEstadoManto <> "V" And Not freControlAseguradora.Visible Then
            If lngTipoPacMedico > 0 Then
                optMedico.Value = True
                txtMovimientoPaciente_KeyDown vbKeyReturn, 0
            Else
                MsgBox SIHOMsg(967) & "médico!", vbCritical, "Mensaje" '¡No se ha configurado el tipo de paciente para la cuenta del médico
                optPaciente.Value = True
                If txtClaveArticulo.Enabled Then
                    txtClaveArticulo.SetFocus
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyF7 Then
    'Búsqueda de empleados
        If vgstrEstadoManto <> "B" And vgstrEstadoManto <> "P" And vgstrEstadoManto <> "V" And Not freControlAseguradora.Visible Then
            If lngTipoPacEmpleado > 0 Then
                optEmpleado.Value = True
                txtMovimientoPaciente_KeyDown vbKeyReturn, 0
            Else
                MsgBox SIHOMsg(967) & "empleado!", vbCritical, "Mensaje" '¡No se ha configurado el tipo de paciente para la cuenta del empleado
                optPaciente.Value = True
                If txtClaveArticulo.Enabled Then
                    txtClaveArticulo.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub optArticulo_Click()
    pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub optArticulo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    optArticulo_Click
End Sub

Private Sub optOtrosConceptos_Click()
    pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub optOtrosConceptos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    optOtrosConceptos_Click
End Sub

Private Sub tmrHora_Timer()
    lblHora.Caption = Mid(Time(), 1, Len(Time()) - 3)
End Sub

Private Sub pAumentaFolioFactura()
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    vlstrSentencia = "update RegistroFolio set intNumeroActual=intNumeroActual+1 "
    vlstrSentencia = vlstrSentencia + "where smiDepartamento = " + str(vgintNumeroDepartamento) + " "
    vlstrSentencia = vlstrSentencia + "and chrTipoDocumento='FA' and intNumeroActual<=intNumeroFinal"
    pEjecutaSentencia vlstrSentencia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAumentaFolio"))
End Sub

Private Sub pAumentaFolio()
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    vlstrSentencia = "update RegistroFolio set intNumeroActual=intNumeroActual+1 "
    vlstrSentencia = vlstrSentencia + "where smiDepartamento = " + str(vgintNumeroDepartamento) + " "
    vlstrSentencia = vlstrSentencia + "and chrTipoDocumento='RE' and intNumeroActual<=intNumeroFinal"
    pEjecutaSentencia vlstrSentencia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAumentaFolio"))
End Sub

Private Sub optUM_Click()
    pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub optUM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    optUM_Click
End Sub

Private Sub optUV_Click()
    pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub optUV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pEnfocaTextBox txtClaveArticulo
End Sub

Private Sub chkConsultaPrecios_Click()
    If txtClaveArticulo.Enabled Then
        If txtClaveArticulo.Enabled Then
            txtClaveArticulo.SetFocus
        End If
    End If
End Sub

Private Sub chkConsultaPrecios_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtClaveArticulo.Enabled Then
        txtClaveArticulo.SetFocus
    End If
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub cmdBuscar_Click()
    frmConsultaPOS.Show vbModal
    If txtClaveArticulo.Enabled Then
        txtClaveArticulo.SetFocus
        vlnblnLocate = False
    End If
End Sub

Private Sub txtClaveArticulo_LostFocus()
    lblMensajes = "<END> Terminar venta"
End Sub

Private Sub txtMovimientoPaciente_LostFocus()
    lblMensajes = "<END> Terminar venta"
End Sub
Private Sub grdArticulos_GotFocus()
    lblMensajes = "<DEL> = Borrar artículo  <ENTER> = Seleccionar"
End Sub

Private Sub grdArticulos_LostFocus()
    lblMensajes = "<END> Terminar venta"
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
    ' Verifica folios
    If Trim(vlstrFolioDocumento) = "0" Or Trim(lblFactura.Caption) = "0" Then
        Unload Me
    End If

    lblMensajes = "<ENTER> Buscar " & IIf(optPaciente.Value, "paciente", IIf(optMedico.Value, "médico", "empleado"))
    
    If Trim(txtMovimientoPaciente.Text) <> "" And optPaciente.Value Then
        
        'Revisa si la cuenta tiene cargos
        lngCargos = 1
        frsEjecuta_SP Trim(txtMovimientoPaciente.Text), "sp_PvSelNumCargosCuenta", False, lngCargos
        
        If lngCargos = 0 Then
            If intCuentaNueva = 0 Then
                pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
            ElseIf intCuentaNueva = 1 Then
                pBorrarCuenta CLng(Val(txtMovimientoPaciente.Text))
            End If
        End If
        
    ElseIf Trim(txtMovimientoPaciente.Text) <> "" And (optMedico.Value Or optEmpleado.Value) Then
    
        'Revisa si la cuenta tiene cargos
        lngCargos = 1
        frsEjecuta_SP Trim(txtMovimientoPaciente.Text), "sp_PvSelNumCargosCuenta", False, lngCargos
    
        If lngCargos = 0 Then
            If intCuentaNueva = 0 Then
                pCerrarCuenta CLng(Val(txtMovimientoPaciente.Text))
            ElseIf intCuentaNueva = 1 Then
                pBorrarCuenta CLng(Val(txtMovimientoPaciente.Text))
            End If
            lblPaciente.Caption = ""
            vgstrRFCPersonaSeleccionada = ""
            lblEmpresa.Caption = ""
        End If
        
    End If
    
    txtMovimientoPaciente.Text = ""
    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
End Sub

Private Sub txtBuscaArticulo_GotFocus()
    lblMensajes.Caption = "Teclee la descripción del artículo"
End Sub

Private Sub txtCopago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdGrabaControlAseguradora.SetFocus
End Sub

Private Sub txtCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtCopago
End Sub

Private Sub txtDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pEnfocaTextBox txtCoaseguro
End Sub

Private Sub txtDeducible_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtDeducible, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCoaseguro_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCoaseguro, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCopago_KeyPress(KeyAscii As Integer)
    If Not fblnFormatoCantidad(txtCopago, KeyAscii, 2) Then KeyAscii = 7
End Sub

Private Sub txtCopago_GotFocus()
    pEnfocaTextBox txtCopago
End Sub

Private Sub txtCoaseguro_GotFocus()
    pEnfocaTextBox txtCoaseguro
End Sub

Private Sub txtDeducible_GotFocus()
    pEnfocaTextBox txtDeducible
End Sub

Private Sub txtDeducible_LostFocus()
    txtDeducible.Text = FormatNumber(Val(Format(txtDeducible.Text, "")), 2)
End Sub

Private Sub txtCopago_LostFocus()
    txtCopago.Text = FormatNumber(Val(Format(txtCopago.Text, "")), 2)
End Sub

Private Sub optCopagoCantidad_Click()
    lblSignoPorcientoCopago.Visible = False
End Sub

Private Sub optControlCopagoPorciento_Click()
    lblSignoPorcientoCopago.Visible = True
End Sub

Private Sub pAbrirCOM1()
    comPrinter.CommPort = 1
    comPrinter.Settings = "9600,n,8,1"

    If (comPrinter.PortOpen = True) Then
        comPrinter.PortOpen = False
        DoEvents: DoEvents
    End If

    comPrinter.PortOpen = True
    DoEvents
End Sub

Private Sub pCerrarCOM1()
    comPrinter.PortOpen = False
    DoEvents
End Sub

Private Sub ImprimirCOM1(vlstrTexto As String)
    comPrinter.Output = vlstrTexto
    DoEvents
End Sub

Private Sub pAbrirCaja()
    comPrinter.Output = Chr$(7) ' Commando para abrir la Caja
    DoEvents
End Sub

Private Sub pActivarSonidoCaja()
    comPrinter.Output = Chr$(30) ' Sound Buzzer
End Sub
    
'Función que obtiene la clave del empleado, médico, empresa etc. que se guardará en PvDatosFiscales.intNumReferencia
'y carga la variable pstrTipo para ser guardada en PvDatosFiscales.chrTipoCliente
Private Function fstrObtieneNumRefTipo(pintClave As Long, pstrTipo As String) As String
    Dim vlstrSentencia As String
    Dim vlrsClaveTipo As New ADODB.Recordset

    vlstrSentencia = " Select RegistroExterno.INTCVEEXTRA, adTipoPaciente.CHRTIPO  "
    vlstrSentencia = vlstrSentencia & " From RegistroExterno "
    vlstrSentencia = vlstrSentencia & "   Inner Join Externo On (RegistroExterno.INTNUMPACIENTE = Externo.INTNUMPACIENTE)"
    vlstrSentencia = vlstrSentencia & "   Inner Join AdTipoPaciente On (RegistroExterno.TNYCVETIPOPACIENTE = AdTipoPaciente.TNYCVETIPOPACIENTE) "
    vlstrSentencia = vlstrSentencia & " Where RegistroExterno.INTNUMCUENTA = " & pintClave
    Set vlrsClaveTipo = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)
    fstrObtieneNumRefTipo = vlrsClaveTipo!intCveExtra
    pstrTipo = vlrsClaveTipo!chrTipo
End Function

Private Sub pImprimeTicket(pstrCveTicket As String)
    Dim vlIntCont As Integer       '[  Contador general  ]
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
    vlintDesgloseIEPS = IIf(vlblnLicenciaIEPS, fRegresaParametro("BITDESGLOSEIEPSTICKET", "PvParametro", 0), 0)
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
                                    vlstrLinea = vlstrLinea & IIf(vlrsFormatoTicket!VCHVALOR = "×", " ", vlrsFormatoTicket!VCHVALOR)
                                Else '[  Campos insertables  ]
                                    '------------------------------------------------------
                                    '««  Localiza la posición del campo en el RecordSet  »»
                                    '------------------------------------------------------
                                    For vlIntCont = 0 To vlrsValoresTicket.Fields.Count - 1
                                        If UCase(vlrsValoresTicket.Fields(vlIntCont).Name) = UCase(vlrsFormatoTicket!VCHVALOR) Then
                                            vlintPosicion = vlIntCont
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
            MsgBox SIHOMsg(13), vbCritical, "Mensaje" '¡No existe información!
        End If
    Else
        MsgBox SIHOMsg(277), vbCritical, "Mensaje" 'No existen registrados formatos de impresión.
    End If
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

Private Function flngCliente(strTipoCliente As String) As Long
    frmBusquedaClienteMedEmp.lstrTipoCliente = strTipoCliente
    frmBusquedaClienteMedEmp.Show vbModal
    flngCliente = frmBusquedaClienteMedEmp.llngNumCliente
End Function

Private Sub pObtieneCuentaPaciente(lngCliente As Long)
    Dim lngCuentaExterno As Long
    Dim alstrParametroSalida() As String

    'Obtiene la cuenta del paciente relacionado con el cliente
    pCargaArreglo alstrParametroSalida, intCuentaNueva & "|" & ADODB.adInteger
    lngCuentaExterno = 1
    vgstrParametrosSP = lngCliente & "|" & IIf(optMedico.Value, "ME", "EM") & "|" & IIf(optMedico.Value, lngTipoPacMedico, lngTipoPacEmpleado) & "|" & vllngMedicoDefaultPOS & "|" & vgintClaveEmpresaContable & "|" & vgintNumeroDepartamento
    frsEjecuta_SP vgstrParametrosSP, "sp_PvCtaPacienteMedicoEmpleado", True, lngCuentaExterno, alstrParametroSalida
    pObtieneValores alstrParametroSalida, intCuentaNueva ' 1 = Es cuenta nueva 0 = No es cuenta nueva

    If Trim(txtGafete.Text) <> "" Then txtGafete.Text = ""
    
    If lngCuentaExterno > 0 Then
        txtMovimientoPaciente.Text = lngCuentaExterno
        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
    Else
        If optMedico.Value Then
            optMedico.SetFocus
        ElseIf optEmpleado.Value Then
            optEmpleado.SetFocus
        Else
            optPaciente.SetFocus
        End If
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    lblMensajes = "<ENTER> Grabar precio"
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    Dim llngElement As Long
    Dim llngCantidad As Long
    Dim lstrTCargo As String
'    Dim laryParamArticulo() As String
    Dim lstrCveCargo As String
    Dim lstrCveLista As String
    Dim ldblPrecio As Double
    Dim ldblTotDescuento As Double
    Dim ldblSubtotal As Double
    Dim ldblIEPSPercent As Double
    Dim ldblIEPS As Double
    Dim ldblIVA As Double
    Dim rsPMP As New ADODB.Recordset
    Dim vlstrSentencia As String
        
    If Not fblnFormatoCantidad(txtPrecio, KeyAscii, 2) Then
        KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            If txtPrecio = "" Then
                txtPrecio.Visible = False
                ElseIf CDbl(txtPrecio) = 0 Then
                MsgBox SIHOMsg(788), vbInformation, "Mensaje"
                txtPrecio.SetFocus
            ElseIf CDbl(grdArticulos.TextMatrix(grdArticulos.Row, 2)) <> CDbl(txtPrecio.Text) Then
                llngElement = Val(grdArticulos.TextMatrix(grdArticulos.Row, 10))
                llngCantidad = Val(grdArticulos.TextMatrix(grdArticulos.Row, 3))
                lstrTCargo = grdArticulos.TextMatrix(grdArticulos.Row, 12)
                lstrCveLista = grdArticulos.TextMatrix(grdArticulos.Row, 17)
                
                ldblPrecio = CDbl(txtPrecio.Text)
                
                '*********************************************
                If vlintValorPMP = 1 Then
                    'Se compara el precio con el Precio Máximo al Publico
                     vlstrSentencia = "SELECT NVL(MNYPRECIOMAXIMOPUBLICO,0) precio"
                    vlstrSentencia = vlstrSentencia & " FROM IVARTICULO INNER JOIN IVARTICULOEMPRESAS ON IVARTICULO.chrcvearticulo = IVARTICULOEMPRESAS.chrcvearticulo"
                    vlstrSentencia = vlstrSentencia & " and IVARTICULO.INTIDARTICULO =  " & llngElement & ""
                    vlstrSentencia = vlstrSentencia & " AND IVARTICULOEMPRESAS.tnyclaveempresa = " & vgintClaveEmpresaContable & ""
                    vlstrSentencia = vlstrSentencia & " WHERE  IVARTICULO.BITVENTAPUBLICO = 1 AND IVARTICULO.CHRCVEARTMEDICAMEN = 1"
    
                    Set rsPMP = frsRegresaRs(vlstrSentencia)
                    If rsPMP.RecordCount <> 0 Then
                        If (IIf(IsNull(rsPMP!precio), 0, rsPMP!precio) > 0 And ldblPrecio >= rsPMP!precio) Then ldblPrecio = rsPMP!precio
                    End If
                    rsPMP.Close: Set rsPMP = Nothing
                End If
                '*********************************************
                
                If Val(grdArticulos.TextMatrix(grdArticulos.Row, 20)) = 2 Then
                   ldblTotDescuento = 0    ' SI hay exclusion de descuento
                Else 'Descuentos
                    ldblTotDescuento = 0
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    vgstrParametrosSP = "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & CStr(Val(txtMovimientoPaciente.Text)) & "|" & lstrTCargo & "|" & CStr(llngElement) & "|" & ldblPrecio & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & Trim(str(intAplicado)) & "|" & Trim(grdArticulos.TextMatrix(grdArticulos.Row, 16)) & "|" & CStr(llngCantidad) & "|" & Trim(grdArticulos.TextMatrix(grdArticulos.Row, 20))
                    frsEjecuta_SP vgstrParametrosSP, "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, ldblTotDescuento
                End If
                
                ldblSubtotal = (ldblPrecio * CLng(llngCantidad)) - ldblTotDescuento
                
                '----------------------------------
                ' Porcentaje de IEPS
                '----------------------------------
                If vlblnLicenciaIEPS And lstrTCargo = "AR" Then
                   ldblIEPSPercent = fdblObtenerPorcentajeIEPS(CLng(llngElement)) / 100 '<----------------
                Else
                   ldblIEPSPercent = 0
                End If
                '----------------------------------
                ldblIEPS = ldblSubtotal * ldblIEPSPercent '<---sacamos cantidada de IEPS
                ldblSubtotal = ldblSubtotal + ldblIEPS '<-----sumamos el IEPS al Subtotal
                               
                '-----------------------
                ' Procedimiento para obtener % del IVA
                '-----------------------
                ldblIVA = fdblObtenerIva(CLng(llngElement), lstrTCargo) / 100

                grdArticulos.TextMatrix(grdArticulos.Row, 2) = Format(ldblPrecio, "$###,###,###,###.00")
                grdArticulos.TextMatrix(grdArticulos.Row, 24) = ldblPrecio
                
                grdArticulos.TextMatrix(grdArticulos.Row, 4) = Format(ldblPrecio * CLng(llngCantidad), "$###,###,###,###.00")
                grdArticulos.TextMatrix(grdArticulos.Row, 5) = Format(ldblTotDescuento, "$###,###,###,###.00")
                grdArticulos.TextMatrix(grdArticulos.Row, 6) = Format(ldblIEPS, "$###,###,###,###.00") '@
                grdArticulos.TextMatrix(grdArticulos.Row, 7) = Format(ldblSubtotal, "$###,###,###,###.00") '@
                grdArticulos.TextMatrix(grdArticulos.Row, 8) = Format(ldblSubtotal * ldblIVA, "$###,###,###,###.00")
                grdArticulos.TextMatrix(grdArticulos.Row, 9) = Format(ldblSubtotal * (ldblIVA + 1), "$###,###,###,###.00")
                grdArticulos.TextMatrix(grdArticulos.Row, 11) = ldblSubtotal * ldblIVA
                grdArticulos.TextMatrix(grdArticulos.Row, 18) = "0" 'Porcentaje de descuento
                grdArticulos.TextMatrix(grdArticulos.Row, 19) = ldblIEPSPercent 'Porcentaje de IEPS
                grdArticulos.TextMatrix(grdArticulos.Row, 23) = "1"
                pCalculaTotales
                txtPrecio.Visible = False
                If fblnCanFocus(txtClaveArticulo) Then txtClaveArticulo.SetFocus
            ElseIf CDbl(grdArticulos.TextMatrix(grdArticulos.Row, 2)) = CDbl(txtPrecio.Text) Then
                txtPrecio.Visible = False
            End If
        End If
    End If
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.Visible = False
    lblMensajes = "<END> Terminar venta"
End Sub
'- CASO 7442: Regresa tipo de movimiento según la forma de pago -'
Private Function fstrTipoMovimientoForma(lintCveForma As Integer, lstrTipoDoc As String) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    Dim lstrSentencia As String
    
    fstrTipoMovimientoForma = ""
    
    lstrSentencia = "SELECT * FROM PVFORMAPAGO WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        If lstrTipoDoc = "F" Then 'Movimientos para Factura en Venta al público
            Select Case rsForma!chrTipo
                Case "E": fstrTipoMovimientoForma = "EFV"
                Case "T": fstrTipoMovimientoForma = "TAV"
                Case "B": fstrTipoMovimientoForma = "TPV"
                Case "H": fstrTipoMovimientoForma = "CQV"
            End Select
        Else 'Movimientos para Tickets
            Select Case rsForma!chrTipo
                Case "E": fstrTipoMovimientoForma = "EFT"
                Case "T": fstrTipoMovimientoForma = "TAT"
                Case "B": fstrTipoMovimientoForma = "TPT"
                Case "H": fstrTipoMovimientoForma = "CQT"
            End Select
        End If
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fstrTipoMovimientoForma"))
End Function

Private Sub pGrabaTasasIEPS(vllngConsecutivoFacturaTasaIEPS As Long)
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
         If Val(Format(.TextMatrix(vlngcont, 6))) > 0 Then ' se cuenta con una cantidad de IEPS @
            vintPosicion = 0
            For vintcont = 1 To vintCantidadTasas
                If (Val(.TextMatrix(vlngcont, 19)) * 100) = ArrId(1, vintcont) Then
                   vintPosicion = vintcont
                   Exit For
                End If
            Next vintcont
            
            If vintPosicion = 0 Then 'AGREGAMOS
               vintCantidadTasas = vintCantidadTasas + 1
               ReDim Preserve ArrId(3, vintCantidadTasas)
               ArrId(1, vintCantidadTasas) = Val(.TextMatrix(vlngcont, 19)) * 100 'TASA DE IEPS
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
                     "values(" & vllngConsecutivoFacturaTasaIEPS & ",'FA'," & ArrId(1, vintcont) & "," & ArrId(3, vintcont) & "," & ArrId(2, vintcont) & ")"
     
         pEjecutaSentencia vstrSentencia
     Next vintcont

End Sub

Private Sub txtTicketPrevio_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cmdConsultaTicketsNoFacturados.SetFocus
    End If
End Sub


Private Sub txtTicketPrevio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 0
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Function fblnValidaSAT() As Boolean
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim vgstrSentencia As String
    
    
    'Por cargo (agrupado/desglosado)
    If vgstrVersionCFDI <> "3.2" Then
        For intRow = 1 To grdArticulos.Rows - 1
            If grdArticulos.TextMatrix(intRow, 12) = "OC" Then
                'Revisa si "Otro Concepto" tiene definida la clave del SAT
                vlstrSentencia = "Select pvOtroConcepto.intCveConcepto clave From pvOtroConcepto " & _
                                 "inner join gnCatalogoSatRelacion on pvOtroConcepto.intCveConcepto = gnCatalogoSatRelacion.intCveConcepto and gnCatalogoSatRelacion.chrTipoConcepto='OC' " & _
                                 "Where gnCatalogoSatRelacion.chrTipoConcepto = 'OC' and gnCatalogoSatRelacion.intDiferenciador = 1 and gnCatalogoSatRelacion.intCveConcepto = " & grdArticulos.TextMatrix(intRow, 10)
                Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                If rsTemp.RecordCount <= 0 Then
                    MsgBox SIHOMsg(1549) & Chr(13) & grdArticulos.TextMatrix(intRow, 1) & ".", vbExclamation, "Mensaje"
                    fblnValidaSAT = False
                    rsTemp.Close
                    Exit Function
                End If
                rsTemp.Close
            Else
                'Revisa si el artículo tiene definida la clave del SAT
                If grdArticulos.TextMatrix(intRow, 12) = "AR" Then
                    'Revisa si el artículo tiene definida la clave del SAT
                    vlstrSentencia = "Select IvArticulo.intIdArticulo clave From IvArticulo " & _
                                     "inner join gnCatalogoSatRelacion on IvArticulo.intIdArticulo  = gnCatalogoSatRelacion.intCveConcepto and gnCatalogoSatRelacion.chrTipoConcepto='AR' " & _
                                     "Where  gnCatalogoSatRelacion.chrTipoConcepto='AR' and gnCatalogoSatRelacion.intCveConcepto  = " & grdArticulos.TextMatrix(intRow, 10)
                    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rsTemp.RecordCount <= 0 Then
                        MsgBox SIHOMsg(1549) & Chr(13) & grdArticulos.TextMatrix(intRow, 1) & ".", vbExclamation, "Mensaje"
                        fblnValidaSAT = False
                        rsTemp.Close
                        Exit Function
                    End If
                    rsTemp.Close
                End If
            End If
        Next
    End If
    fblnValidaSAT = True
End Function

'|  Importa información de ventas
Public Sub pImportarVenta()
    Dim strLinea As String
    Dim strArchivoImportacion As String         '|  Archivo que se intenta importar
    Dim strSentencia As String                  '|  Variable genérica para armar instrucciones SQL
    Dim lngPersonaGraba As Long
    Dim lngResultado As Long
    Dim lngContador As Long
    Dim vlblnTerminoInformacion As Boolean
    Dim vlintContador As Integer
    Dim intcontador As Long
    Dim intContador2 As Long
    Dim vllngClaveTicket As Long
    Dim vllngClaveTicket2 As Long
    Dim vlblnContieneInformacion As Boolean
        
On Error GoTo NotificaError
            
    Set xlsApp = Nothing
                
    strArchivoImportacion = fstrAbreArchivo

    If strArchivoImportacion = "" Then Exit Sub
            
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    Set f = fs.OpenTextFile(strArchivoImportacion, ForReading, , TristateUseDefault)
                   
    Set xlsApp = CreateObject("Excel.Application")
    DoEvents
    xlsApp.Workbooks.Open strArchivoImportacion
    DoEvents
    Set hoja = xlsApp.Worksheets(1)
    DoEvents
                   
    If UCase(hoja.Cells(1, 1)) <> "FOLIO" Or _
            UCase(hoja.Cells(1, 4)) <> "FECHA" Or _
            UCase(hoja.Cells(1, 5)) <> "IMPORTE" Or _
            UCase(hoja.Cells(1, 6)) <> "DESCUENTO" Or _
            UCase(hoja.Cells(1, 7)) <> "IEPS" Or _
            UCase(hoja.Cells(1, 10)) <> "SUBTOTAL" Or _
            UCase(hoja.Cells(1, 11)) <> "IVA" Or _
            UCase(hoja.Cells(1, 14)) <> "TOTAL" Then
        MsgBox "La información contenida en el archivo no cumple con el formato requerido.", vbCritical, "Mensaje"
        
        
        xlsApp.Quit
        Set xlsApp = Nothing
        
        f.Close
        
        Exit Sub
    End If
            
    Me.MousePointer = 11
    
    vlblnContieneInformacion = False
    
    vlblnTerminoInformacion = False
    vlintContador = 1
    lstFoliosTickets.Clear
    Do While vlblnTerminoInformacion = False
        If Trim(hoja.Cells(vlintContador + 1, 1)) = "" Then
            vlblnTerminoInformacion = True
        Else
            vlblnContieneInformacion = True
        
            'Agregar folios
            lstFoliosTickets.AddItem CDbl(IIf(Trim(hoja.Cells(vlintContador + 1, 1)) = "", "0", Trim(hoja.Cells(vlintContador + 1, 1))))
            lstFoliosTickets.ItemData(lstFoliosTickets.newIndex) = vlintContador
            lstFoliosTickets.Selected(lstFoliosTickets.newIndex) = True
                        
            vlintContador = vlintContador + 1
        End If
    Loop
                       
    For intcontador = 0 To lstFoliosTickets.ListCount
        If intcontador < lstFoliosTickets.ListCount Then
            lstFoliosTickets.ListIndex = intcontador
            vllngClaveTicket = lstFoliosTickets.Text
    
            For intContador2 = intcontador + 1 To lstFoliosTickets.ListCount
                If intContador2 < lstFoliosTickets.ListCount Then
                    lstFoliosTickets.ListIndex = intContador2
                    vllngClaveTicket2 = lstFoliosTickets.Text
                
                    If vllngClaveTicket = vllngClaveTicket2 Then
                        lstFoliosTickets.RemoveItem (intContador2)
                        
                        intContador2 = intContador2 - 1
                    End If
                End If
            Next intContador2
        End If
    Next intcontador
    
    If vlblnContieneInformacion = True Then
        frmFoliosVentaImportada.Visible = True
        
        cmdInvertir.SetFocus
        
        FreDetalle.Enabled = False
        freDatosPaciente.Enabled = False
        frmMedico.Enabled = False
        freControlAseguradora.Enabled = False
        freBusqueda.Enabled = False
        
        txtClaveArticulo.Enabled = False
        txtCantidad.Enabled = False
        optArticulo.Enabled = False
        optOtrosConceptos.Enabled = False
                    
        freConsultaPrecios.Enabled = False
        freDescuentos.Enabled = False
        
        cmdImportarVenta.Enabled = False
    
        freGraba.Enabled = True
        
        freFacturaPesosDolares.Enabled = True
        
'        optPaciente.Value = True
'        optMedico.Value = False
'        optEmpleado.Value = False
'        txtMovimientoPaciente.Text = ""
'        txtGafete.Text = ""
'        lblPaciente.Caption = ""
'        lblEmpresa.Caption = ""
    Else
        MsgBox "La información contenida en el archivo no cumple con el formato requerido.", vbCritical, "Mensaje"
        
        xlsApp.Quit
        Set xlsApp = Nothing
        
        f.Close
    End If
    
    Me.MousePointer = 0
    
Exit Sub
NotificaError:
    Me.MousePointer = 0
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImportarVenta"))
    xlsApp.Quit
    Set xlsApp = Nothing
End Sub

'|  Verifica que el formato del archivo de texto sea el especificado por el layout
Private Function fblnValidacionesGenerales(strArchivo As String) As Boolean
On Error GoTo NotificaError
    
    Dim fs As FileSystemObject
    Dim f As Variant
    Dim strLinea As String
    Dim blnFormatoErroneo As String
    Dim strPolizasImportadasAnteriormente As String     '|  Lista de las pólizas que ya han sido importadas con anterioridad
    Dim lstrArchivo As String
    
    fblnValidacionesGenerales = False
    blnFormatoErroneo = False
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(strArchivo, ForReading, TristateFalse)
    
    '--------------------------------------------------------------------'
    '|  Recorre el archivo para validar los encabezados de las pólizas  |'
    '--------------------------------------------------------------------'
    Do Until f.AtEndOfStream
        DoEvents
        strLinea = f.ReadLine
        
        If Mid(strLinea, 1, 1) = "P" Then
            '---------------------------------------------------------------------------------'
            '|  Verifica que el encabezado del archivo coincida con el formato especificado  |'
            '---------------------------------------------------------------------------------'
            If Mid(strLinea, 2, 1) <> " " _
               Or Mid(strLinea, 11, 1) <> " " _
               Or Mid(strLinea, 13, 1) <> " " _
               Or Mid(strLinea, 22, 1) <> " " _
               Or Mid(strLinea, 24, 1) <> " " _
               Or Mid(strLinea, 28, 1) <> " " _
               Or Mid(strLinea, 129, 1) <> " " _
               Or Mid(strLinea, 132, 1) <> " " _
               Or Mid(strLinea, 134, 1) <> " " Then
               blnFormatoErroneo = True
               Exit Do
            End If
        End If
    Loop
    f.Close
        
    If blnFormatoErroneo Then
        '|  Formato interno de archivo erróneo.
        MsgBox SIHOMsg(746), vbCritical, "Mensaje"
        Exit Function
    End If
    
    lstrArchivo = ""
    If lstrArchivo Then
        lstrArchivo = fstrAbreArchivo()
        If lstrArchivo = "" Then
            Exit Function
        End If
    End If
    
    fblnValidacionesGenerales = True
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidacionesGenerales(" & strArchivo & ")"))
End Function

'|  Muestra el diálogo de abrir archivo y regresa la ruta del archivo seleccionado
Private Function fstrAbreArchivo() As String
On Error GoTo NotificaError
    
    fstrAbreArchivo = ""
    CDgArchivo.CancelError = True

    With cdgExcel
        .FileName = ""
        .DialogTitle = "Abrir archivo para importación de datos del comprobante"
        .InitDir = App.Path
        .Flags = cdlOFNHideReadOnly
        .Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx"
        .ShowOpen
        fstrAbreArchivo = .FileName
    End With
    
Exit Function
NotificaError:
End Function

Private Sub pTerminaParametrosReporte(vllngCiudad As Long, vlStrConcepto As String)
    alstrParametros(2) = "Concepto;" & vlStrConcepto
    alstrParametros(3) = "LugarFecha;" & frsRegresaRs("select ciudad.vchdescripcion from ciudad where intcveciudad = " & vllngCiudad).Fields(0) & "., a " & Format(fdtmServerFecha, "Long Date")
    alstrParametros(4) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
    alstrParametros(5) = "Recibi;" & Trim(lblPaciente.Caption)
    alstrParametros(6) = "Usuario;" & frsRegresaRs("SELECT NoEmpleado.vchApellidoPaterno || ' ' || NoEmpleado.vchApellidoMaterno || ' ' || NoEmpleado.vchNombre AS Empleado FROM noEmpleado WHERE intCveEmpleado = " & Trim(str(vllngPersonaGraba)), adLockReadOnly, adOpenForwardOnly)!Empleado
End Sub

Private Sub pAgregarMovArreglo()
    pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
    pAgregarMovArregloCorte2 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
End Sub

Public Function fblnValidaSATotrosConceptosCONCEPTOFACTURACION()
    Dim intRow As Integer
    
    'Por concepto de facturación
    If vgstrVersionCFDI <> "3.2" Then
        For intRow = 1 To grdArticulos.Rows - 1
            If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", grdArticulos.TextMatrix(intRow, 13), "CF", 1) = 0 Then
                MsgBox "No está definida la clave del SAT para el producto/servicio " & grdArticulos.TextMatrix(intRow, 21), vbExclamation, "Mensaje"
                fblnValidaSATotrosConceptosCONCEPTOFACTURACION = False
                Exit Function
            End If
        Next
    End If
    fblnValidaSATotrosConceptosCONCEPTOFACTURACION = True
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

Private Sub pFacturasDirectasAnteriores()
    Dim rsFacturas As New ADODB.Recordset
    'Dim vlstrsql As String
    Dim i As Integer
    
    vlstrsql = ""
      
    If lblPaciente.Caption <> "VENTA AL PÚBLICO" And txtMovimientoPaciente.Text <> "" Then
       If vgstrTipoPaciente = "PA" Then
         vlstrsql = "SELECT PVFACTURA.CHRFOLIOFACTURA, PVFACTURA.DTMFECHAHORA, PVFACTURA.MNYTOTALFACTURA, CASE WHEN BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, PVFACTURA.MNYTIPOCAMBIO FROM PVFACTURA " & _
                            "INNER JOIN GNCOMPROBANTEFISCALDIGITAL ON GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'FA' " & _
                            "AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVFACTURA.INTCONSECUTIVO " & _
                            "INNER JOIN EXPACIENTEINGRESO ON PVFACTURA.INTMOVPACIENTE = EXPACIENTEINGRESO.INTNUMCUENTA " & _
                    "WHERE NOT VCHUUID IS NULL AND PVFACTURA.CHRTIPOPACIENTE = 'E' AND PVFACTURA.CHRTIPOFACTURA = 'P' AND PVFACTURA.CHRESTATUS = 'C' AND NVL(PVFACTURA.INTCVEVENTAPUBLICO,0) <> 0 " & _
                            "AND EXPACIENTEINGRESO.INTNUMPACIENTE = " & vllngNumeroPaciente & _
                    "ORDER BY INTCONSECUTIVO DESC "
       Else
            If vgstrTipoPaciente = "ME" Or vgstrTipoPaciente = "EM" Then
                vlstrsql = "SELECT PVFACTURA.CHRFOLIOFACTURA, PVFACTURA.DTMFECHAHORA, PVFACTURA.MNYTOTALFACTURA, CASE WHEN BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, PVFACTURA.MNYTIPOCAMBIO FROM PVFACTURA " & _
                                "INNER JOIN GNCOMPROBANTEFISCALDIGITAL ON GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'FA' " & _
                                "AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVFACTURA.INTCONSECUTIVO " & _
                                "INNER JOIN PVVENTAPUBLICO ON PVVENTAPUBLICO.INTCVEVENTA = PVFACTURA.INTCVEVENTAPUBLICO " & _
                           "WHERE NOT VCHUUID IS NULL AND PVFACTURA.INTMOVPACIENTE IN " & _
                                "(SELECT EXPACIENTEINGRESO.INTNUMCUENTA FROM EXPACIENTEINGRESO WHERE EXPACIENTEINGRESO.INTNUMPACIENTE IN (SELECT EXPACIENTEINGRESO.INTNUMPACIENTE FROM EXPACIENTEINGRESO WHERE EXPACIENTEINGRESO.INTNUMCUENTA = " & Trim(txtMovimientoPaciente.Text) & ")" & ")" & _
                                " AND PVFACTURA.CHRTIPOPACIENTE = 'E' AND PVFACTURA.CHRTIPOFACTURA = 'P' AND PVFACTURA.CHRESTATUS = 'C' AND NVL(PVFACTURA.INTCVEVENTAPUBLICO,0) <> 0 " & _
                           " ORDER BY INTCONSECUTIVO DESC"
            Else
                vlstrsql = "SELECT PVFACTURA.CHRFOLIOFACTURA, PVFACTURA.DTMFECHAHORA, PVFACTURA.MNYTOTALFACTURA, CASE WHEN BITPESOS = 1 THEN 'Pesos' ELSE 'Dólares' END PESOS, PVFACTURA.MNYTIPOCAMBIO FROM PVFACTURA " & _
                                "INNER JOIN GNCOMPROBANTEFISCALDIGITAL ON GNCOMPROBANTEFISCALDIGITAL.CHRTIPOCOMPROBANTE = 'FA' " & _
                                "AND GNCOMPROBANTEFISCALDIGITAL.INTCOMPROBANTE = PVFACTURA.INTCONSECUTIVO " & _
                                "INNER JOIN EXPACIENTEINGRESO ON PVFACTURA.INTMOVPACIENTE = EXPACIENTEINGRESO.INTNUMCUENTA " & _
                           "WHERE NOT VCHUUID IS NULL AND PVFACTURA.CHRTIPOPACIENTE = 'E' AND PVFACTURA.CHRTIPOFACTURA = 'E' AND PVFACTURA.CHRESTATUS = 'C' AND NVL(PVFACTURA.INTCVEVENTAPUBLICO,0) <> 0 " & _
                                "AND PVFACTURA.INTCVEEMPRESA = " & vlngEmpresa & _
                           "ORDER BY INTCONSECUTIVO DESC"
            End If
        End If
            Set rsFacturas = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        
            If rsFacturas.RecordCount <> 0 Then
                chkFacturaSustitutaDFP.Enabled = True
                lstFacturaASustituirDFP.Enabled = True
            
                lstFacturaASustituirDFP.Clear
            Else
                chkFacturaSustitutaDFP.Enabled = False
                lstFacturaASustituirDFP.Enabled = False
            End If
                rsFacturas.Close
       End If
                vlnblnLocate = False
End Sub

Private Sub pRefacturacionAnteriores()
Dim i As Integer

    If chkFacturaSustitutaDFP.Value = 1 And lstFacturaASustituirDFP.ListCount > 0 Then
            For i = 0 To UBound(aFoliosPrevios())
                If aFoliosPrevios(i).chrfoliofactura <> "" Then
                    pEjecutaSentencia "INSERT INTO PVREFACTURACION (chrFolioFacturaActivada, chrFolioFacturaCancelada) " & " VALUES ('" & Trim(lblFactura.Caption) & "', '" & aFoliosPrevios(i).chrfoliofactura & "')"
                End If
            Next i
        End If
End Sub

Private Sub pActualizaPrecio(vlintContador As Integer)

    If Val((grdArticulos.TextMatrix(vlintContador, 23))) = 1 Then ' si se modificó el precio se actualiza en pvcargo
        If CDbl(IIf(Trim(grdArticulos.TextMatrix(vlintContador, 24)) = "", "0", Trim(grdArticulos.TextMatrix(vlintContador, 24)))) <> 0 Then
            pEjecutaSentencia "Update pvCargo set mnyPrecio = " & Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "")) & ", mnyIVA = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2))) & _
                            ", mnyDescuento = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")), 2))) & ", bitPrecioManual = 1 where intNumCargo = " & grdArticulos.TextMatrix(vlintContador, 15)
        Else
            pEjecutaSentencia "Update pvCargo set mnyPrecio = " & Format(Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "")), "###########0.00") & ", mnyIVA = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2))) & _
                            ", mnyDescuento = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")), 2))) & ", bitPrecioManual = 1 where intNumCargo = " & grdArticulos.TextMatrix(vlintContador, 15)
        End If
    Else
        'Se agregó esta parte para que siempre actualice el precio del cargo y lo deje tal y como está en la venta, para evitar problemas si se cambia el precio del artículo durante la facturación del POS
        If CDbl(IIf(Trim(grdArticulos.TextMatrix(vlintContador, 24)) = "", "0", Trim(grdArticulos.TextMatrix(vlintContador, 24)))) <> 0 Then
            pEjecutaSentencia "Update pvCargo set mnyPrecio = " & Val(Format(grdArticulos.TextMatrix(vlintContador, 24), "")) & ", mnyIVA = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2))) & _
                            ", mnyDescuento = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")), 2))) & " where intNumCargo = " & grdArticulos.TextMatrix(vlintContador, 15)
        Else
            pEjecutaSentencia "Update pvCargo set mnyPrecio = " & Format(Val(Format(grdArticulos.TextMatrix(vlintContador, 2), "")), "###########0.00") & ", mnyIVA = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 11), "")), 2))) & _
                            ", mnyDescuento = " & Trim(str(Round(Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")), 2))) & " where intNumCargo = " & grdArticulos.TextMatrix(vlintContador, 15)
        End If
                        
    End If

End Sub

Public Function fblnValidaCuentasIngresoDescuento()
    Dim vldblImporte As Double
    Dim vlintContador As Integer
    
    fblnValidaCuentasIngresoDescuento = True
    
    For vlintContador = 1 To grdArticulos.Rows - 1
        If CDbl(IIf(Trim(grdArticulos.TextMatrix(vlintContador, 24)) = "", "0", Trim(grdArticulos.TextMatrix(vlintContador, 24)))) <> 0 Then
            vldblImporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 24), ""))
        Else
            vldblImporte = Val(Format(grdArticulos.TextMatrix(vlintContador, 2), ""))
        End If
        
        If vldblImporte <> 0 Then
            If flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "INGRESO") = 0 Then
                MsgBox Replace(SIHOMsg(907), "contable", "contable de ingreso") & Chr(13) & Trim(grdArticulos.TextMatrix(vlintContador, 13)) & "  " & fstrConceptoFacturacion(Val(grdArticulos.TextMatrix(vlintContador, 13))), vbOKOnly + vbInformation, "Mensaje"
                fblnValidaCuentasIngresoDescuento = False
                Exit Function
            End If
        End If
    
        If Val(Format(grdArticulos.TextMatrix(vlintContador, 5), "")) <> 0 Then
            If flngCuentaConceptoDepartamento(grdArticulos.TextMatrix(vlintContador, 13), vgintNumeroDepartamento, "DESCUENTO") = 0 Then
                MsgBox Replace(SIHOMsg(907), "contable", "contable de descuento") & Chr(13) & Trim(grdArticulos.TextMatrix(vlintContador, 13)) & "  " & fstrConceptoFacturacion(Val(grdArticulos.TextMatrix(vlintContador, 13))), vbOKOnly + vbInformation, "Mensaje"
                fblnValidaCuentasIngresoDescuento = False
                Exit Function
            End If
        End If
    Next vlintContador
        
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaCuentasIngresoDescuento"))
End Function

Private Function fstrConceptoFacturacion(vlintCveConcepto As Long) As String
    Dim rsCF As ADODB.Recordset

On Error GoTo NotificaError

    fstrConceptoFacturacion = ""
    Set rsCF = frsRegresaRs("SELECT chrDescripcion FROM pvConceptoFacturacion WHERE smiCveConcepto = " & vlintCveConcepto)
    While Not rsCF.EOF
        fstrConceptoFacturacion = rsCF!chrDescripcion
        rsCF.MoveNext
    Wend
    rsCF.Close
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrConceptoFacturacion"))
End Function

Private Function fSiguienteClaveVentaPublicoNoFacturado() As Long
    On Error GoTo NotificaError
    
    fSiguienteClaveVentaPublicoNoFacturado = 1
    
    Dim rsUltimaClave As New ADODB.Recordset
    
    Set rsUltimaClave = frsRegresaRs("select max(nvl(INTCVEVENTA,0)) + 1 as Siguiente from PVVENTAPUBLICONOFACTURADO")
    If rsUltimaClave.RecordCount <> 0 Then
        fSiguienteClaveVentaPublicoNoFacturado = IIf(IsNull(rsUltimaClave!Siguiente), 1, rsUltimaClave!Siguiente)
    Else
        fSiguienteClaveVentaPublicoNoFacturado = 1
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fSiguienteClaveVentaPublicoNoFacturado"))
End Function

Private Sub vgBuscarArticuloTrazabilidad()
    Dim blnEncontrado As Boolean
    Dim blnTrazaEncontrado As Boolean
    Dim rs As New ADODB.Recordset
       
    blnEncontrado = False
    blnTrazaEncontrado = False
    
     'Se busca en la opcion estandar de código de barra
        If optArticulo.Value Then
            'Para artículos
            vlstrSentencia = "select intIdArticulo Elemento from ivArticulo " & _
                             " Where chrCveArticulo = " & _
                             " (Select chrCveArticulo from ivCodigoBarrasArticulo where rtrim(ltrim(vchCodigoBarras)) = '" & Replace(Trim(txtClaveArticulo.Text), "'", "''") & "'and rownum < 2)"
        Else
            'Otros conceptos
            vlstrSentencia = "select intCveConcepto Elemento from PvOtroConcepto " & _
                             " Where intCveConcepto = " & Val(Mid(Trim(txtClaveArticulo.Text), 1, 9))
        End If
    
        'Se realiza la consulta de la forma estandar
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        
        'Se valida si no viene en cero los registros
        If rs.RecordCount > 0 Then
            pSeleccionaElemento rs!Elemento, CLng(txtCantidad.Text), IIf(optArticulo.Value, "AR", "OC")
            txtCantidad.Text = 1
            blnEncontrado = True
        End If
        
        'Se busca en la opcion de trazabilidad
        If optArticulo.Value Then
            
            vlstrSentencia = "Select count(*) as reg from ivetiqueta"
            
            'Se realiza la consulta de la forma estandar
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            
            If rs.RecordCount > 0 Then
                If rs!REG > 0 Then
        
                    vlstrSentencia = "select intIdArticulo Elemento from ivArticulo " & _
                                     " Where chrCveArticulo = " & _
                                     "(Select CAST(CHRCVEARTICULO as CHAR(10)) from ivetiqueta where rtrim(ltrim(INTIDETIQUETA)) = '" & Replace(Trim(txtClaveArticulo.Text), "'", "''") & "'and rownum < 2)"
                                     
                    'Se realiza la consulta de la forma estandar
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
                    If rs.RecordCount > 0 Then
                        pSeleccionaElementoTrazabilidad rs!Elemento, CLng(Trim(txtClaveArticulo.Text)), CLng(txtCantidad.Text), "AR"
                        txtCantidad.Text = 1
                        blnEncontrado = True
                    
                    End If
                Else
                    'Mensaje cuando no hay registro nen la tabla IvEtiqueta
                    'El artículo que está buscando no existe o no sé a generado la etiqueta para este, verifique por favor.
                    MsgBox SIHOMsg(1675) & Chr(vbKeyReturn), vbExclamation, "Mensaje"
                    Exit Sub
                End If
            Else
                'Mensaje cuando no hay registro nen la tabla IvEtiqueta
                'El artículo que está buscando no existe o no sé a generado la etiqueta para este, verifique por favor.
                MsgBox SIHOMsg(1675) & Chr(vbKeyReturn), vbExclamation, "Mensaje"
                Exit Sub
            End If
        Else
           'Mensaje al seleccionar optOtrosConceptos
           MsgBox SIHOMsg(1674) & Chr(vbKeyReturn), vbExclamation, "Mensaje"
        End If
                             
        'Mensaje que se muestra si no encuentra en ninguna de las dos opciones de busqueda
        If Not blnEncontrado Then
            MsgBox SIHOMsg(13) & Chr(vbKeyReturn) & "Artículo inválido.", vbExclamation, "Mensaje"
            pEnfocaTextBox txtClaveArticulo
        End If
        rs.Close
End Sub

Private Sub pSeleccionaElementoTrazabilidad(vllngClaveElemento As Long, lngIdEtiqueta As Long, vlintCantidad As Long, vlstrTipoCargo As String)
    
On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim rsTipoPaciente As New ADODB.Recordset
    Dim rsPMP As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlStrTipoPaciente As String
    Dim vlintPosicion As Integer
    Dim vlintContador As Integer
    Dim lstListas As ListBox
    Dim vldblPrecio As Double
    Dim vldblIncrementoTarifa As Double 'Incremento en la tarifa de los precios
    Dim vldblSubtotal As Double
    Dim vldblDescuento As Double
    Dim vldblIEPS As Double '<----------
    Dim vldblIEPSPercent As Double '<---
    Dim vldblIVA As Double
    Dim vldblDescUnitario As Double
    Dim vldblTotDescuento As Double
    Dim vlstrTipoDescuento As String
    Dim vlintModoDescuentoInventario As Integer
    Dim vllngContenido As Long
    Dim vlstrAux As String
    Dim a As Integer
    Dim vlstrPrecio As String
    Dim vlstrIncrementoTarifa As String
    Dim vlaryParametrosSalida() As String
    Dim vlstrCveArticulo As String
    Dim vlchrTipoDescuento As String
    Dim vldblPorcentajeDescuento As Double
    Dim vlintAuxExclusionDescuento As Long
    Dim DescuentoInventario As Integer
    Dim MNYCantidad As Integer
    Dim rsVentaAlmacen As New ADODB.Recordset
    Dim vlintAlmacenVenta As Integer
    Dim vlbExistencias As Boolean
    Dim vlchrDescripcion As String
    Dim vlintUV As Long
    Dim vlintUM As Long
    Dim lngNumCargo As Long
        
    If Val(Format(vlintCantidad, "")) > 0 Then
        With grdArticulos
            vllngContenido = 1  'Este es el Contenido en IvArticulo
            If vllngClaveElemento <> -1 Then
                vlintModoDescuentoInventario = 0
                If vlstrTipoCargo = "AR" Then 'Nomas para los articulos
                    '-------------------
                    ' Tipo de descuento de Inventario
                    '-------------------
                    vlstrSentencia = "SELECT intContenido Contenido,vchNombreComercial Descripcion, substring(vchNombreComercial,1,50) Articulo,  ivUA.vchDescripcion UnidadAlterna,  ivUM.vchDescripcion UnidadMinima," & _
                                    " chrCveArticulo CveArticulo, ivArticulo.bitVentaPublico" & _
                                    " From ivArticulo " & _
                                    " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
                                    " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
                                    " WHERE intIdArticulo = " & Trim(str(vllngClaveElemento))

                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rs.RecordCount > 0 Then
                        If rs!bitVentaPublico = 0 Then
                            MsgBox SIHOMsg(1548), vbInformation, "Mensaje"
                            rs.Close
                            Exit Sub
                        End If
                        vlstrCveArticulo = rs!cveArticulo
                        vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
                        vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
                        vlchrDescripcion = rs!Descripcion
                        If vllngContenido > 1 Then
                            If MsgBox("¿Desea realizar la venta de " & Trim(rs!Articulo) & " por " & Trim(rs!UnidadAlterna) & "?" & Chr(13) & "Si selecciona NO, se venderá por " & Trim(rs!UnidadMinima) & ".", vbYesNo + vbQuestion, "Mensaje") = vbNo Then
                                vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
                            End If
                        End If
                        rs.Close
                    End If
                End If
                '-----------------------
                'Precio unitario (En TODO el POS es SIN AUMENTO de TARIFAS, así lo dijo SERGIO)
                '-----------------------
                pCargaArreglo vlaryParametrosSalida, "|" & adDecimal & "||" & adDecimal
                frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|1|01/01/1900|" & vgintClaveEmpresaContable, "SP_PVSELOBTENERPRECIO", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, vldblPrecio, vldblIncrementoTarifa
                
                If vldblPrecio = -1 Or vldblPrecio = 0 Then
                    MsgBox SIHOMsg(301), vbInformation, "Mensaje"
                    Exit Sub
                End If
                
                '*********************************************
                If vlintValorPMP = 1 Then
                    'Se compara el precio con el Precio Máximo al Publico
                     vlstrSentencia = "SELECT NVL(MNYPRECIOMAXIMOPUBLICO,0) precio"
                    vlstrSentencia = vlstrSentencia & " FROM IVARTICULO INNER JOIN IVARTICULOEMPRESAS ON IVARTICULO.chrcvearticulo = IVARTICULOEMPRESAS.chrcvearticulo"
                    vlstrSentencia = vlstrSentencia & " and IVARTICULO.INTIDARTICULO =  " & vllngClaveElemento & ""
                    vlstrSentencia = vlstrSentencia & " AND IVARTICULOEMPRESAS.tnyclaveempresa = " & vgintClaveEmpresaContable & ""
                    vlstrSentencia = vlstrSentencia & " WHERE  IVARTICULO.BITVENTAPUBLICO = 1 AND IVARTICULO.CHRCVEARTMEDICAMEN = 1"
    
                    Set rsPMP = frsRegresaRs(vlstrSentencia)
                    If rsPMP.RecordCount <> 0 Then
                        If (IIf(IsNull(rsPMP!precio), 0, rsPMP!precio) > 0 And vldblPrecio >= rsPMP!precio) Then vldblPrecio = rsPMP!precio
                    End If
                    rsPMP.Close: Set rsPMP = Nothing
                End If
                '*********************************************
'                vldblPrecio = Format(vldblPrecio, "############.00")

                'Existencias
                vlbExistencias = False
                DescuentoInventario = vlintModoDescuentoInventario
                MNYCantidad = vlintCantidad
                
                If vlstrTipoCargo = "AR" Then
                    Set rsVentaAlmacen = frsRegresaRs("SELECT smiCveDepartamento FROM NoDepartamento WHERE NoDepartamento.chrClasificacion = 'A' AND smiCveDepartamento = " & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                    
                    If Not rsVentaAlmacen.RecordCount = 0 Then
                        vlintAlmacenVenta = rsVentaAlmacen!smicvedepartamento
                    Else
                        Set rsVentaAlmacen = frsRegresaRs("SELECT intNumAlmacen FROM PvAlmacenes WHERE intnumdepartamento =" & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
                        vlintAlmacenVenta = rsVentaAlmacen!intnumalmacen
                    End If
                    vlstrSentencia = "SELECT INTEXISTENCIADEPTOUV VL_UV, INTEXISTENCIADEPTOUM VL_UM FROM IVUBICACION WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    If rs.RecordCount > 0 Then
                        vlintUV = rs!VL_UV
                        vlintUM = rs!VL_UM
                    Else
                        vlintUV = 0
                        vlintUM = 0
                    End If
                    
                    'Por unidad minima
                    If DescuentoInventario = 1 And vllngContenido > 1 Then
                        If (((vlintUV * vllngContenido) + vlintUM) < MNYCantidad) Then
                            vlbExistencias = True
                        End If
                    'Por unidad alterna
                    Else
                        If vllngContenido = 1 And DescuentoInventario = 1 Then
                            DescuentoInventario = 2
                        End If
                        If (((vlintUV * vllngContenido) + vlintUM)) < (MNYCantidad * vllngContenido) And DescuentoInventario = 2 Then
                            vlbExistencias = True
                        End If
                    End If
                    
                     If vlbExistencias Then
                        If chkConsultaPrecios.Value = 0 Then
                           MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(317) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
                           Exit Sub
                        End If
                     End If
                End If
                
                '----------------------
                'Clave de la lista de precios
                '----------------------
                pCargaArreglo vlaryParametrosSalida, "|" & adInteger
                frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|" & vgintClaveEmpresaContable, "SP_PVSELLISTAPRECIO", , , vlaryParametrosSalida
                pObtieneValores vlaryParametrosSalida, lCveLista
                
                '-----------------------
                'Exclusión de descuentos
                '-----------------------
                vlintAuxExclusionDescuento = 1
                frsEjecuta_SP "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & vlstrTipoCargo & "|" & vllngClaveElemento & "|" & vgintNumeroDepartamento, "FN_PVSELEXCLUSIONDESCUENTO", True, vlintAuxExclusionDescuento
                '-------------------------------------------------------------
                'vlintAuxExclusionDescuento = 2 SI HAY EXCLUSION DE DESCUENTOS
                'vlintAuxExclusionDescuento = 3 NO HAY EXCLUSION DE DESCUENTOS
                '-------------------------------------------------------------
                If vlintAuxExclusionDescuento = 2 Then
                   vldblTotDescuento = 0
                Else
                    '-----------------------
                    'Descuentos
                    '-----------------------
                    vldblTotDescuento = 0
                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
                    vgstrParametrosSP = "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & CStr(Val(txtMovimientoPaciente.Text)) & "|" & vlstrTipoCargo & "|" & CStr(vllngClaveElemento) & "|" & vldblPrecio & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & Trim(str(intAplicado)) & "|" & CStr(vllngContenido) & "|" & CStr(vlintCantidad) & "|" & CStr(vlintModoDescuentoInventario)
                    frsEjecuta_SP vgstrParametrosSP, "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
                    pObtieneValores vlaryParametrosSalida, vldblTotDescuento
                End If
                
                '-----------------------
                'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
                '-----------------------
                If vlintModoDescuentoInventario = 1 Then
                    vldblPrecio = vldblPrecio / CDbl(vllngContenido)
                End If
                
'                vldblPrecio = Format(vldblPrecio, "############.00")
                
                vldblSubtotal = (vldblPrecio * CLng(vlintCantidad)) - vldblTotDescuento
                
                '----------------------------------
                ' Porcentaje de IEPS
                '----------------------------------
                If vlblnLicenciaIEPS And vlstrTipoCargo = "AR" Then
                   vldblIEPSPercent = fdblObtenerPorcentajeIEPS(CLng(vllngClaveElemento)) / 100 '<----------------
                Else
                   vldblIEPSPercent = 0
                End If
                '----------------------------------
                vldblIEPS = vldblSubtotal * vldblIEPSPercent '<---sacamos cantidada de IEPS
                vldblSubtotal = vldblSubtotal + vldblIEPS '<-----sumamos el IEPS al Subtotal
                               
                '-----------------------
                ' Procedimiento para obtener % del IVA
                '-----------------------
                vldblIVA = fdblObtenerIva(CLng(vllngClaveElemento), vlstrTipoCargo) / 100
                
                '-----------------------
                ' Datos del Artículo ó del Otro concepto
                '-----------------------
                If optArticulo.Value Then
                    vlstrSentencia = "select AR.vchNombreComercial Descripcion, AR.smiCveConceptFact Concepto, CF.chrDescripcion ConceptoDesc from IVArticulo AR inner join PVConceptoFacturacion CF on CF.smiCveConcepto = AR.smiCveConceptFact where AR.intIdArticulo = " & Trim(str(vllngClaveElemento))
                Else
                    vlstrSentencia = "select OC.chrDescripcion Descripcion, OC.smiConceptoFact Concepto, CF.chrDescripcion ConceptoDesc from PVOtroConcepto OC inner join PVConceptoFacturacion CF on CF.smiCveConcepto = OC.smiConceptoFact where OC.intCveConcepto = " & Trim(str(vllngClaveElemento))
                End If
                
                Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
                '-----------------------
                ' Llenado del grid o de la consulta de precios
                '-----------------------
                If chkConsultaPrecios.Value = 0 Then
                    
                    lblnGuardoControl = False
                
                    If .RowData(1) <> -1 Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    End If
                    
                    .TextMatrix(.Row, 1) = rs!Descripcion
                    .TextMatrix(.Row, 2) = Format(vldblPrecio, "$###,###,###,###.0000")
                    .TextMatrix(.Row, 24) = vldblPrecio
                    
                    .TextMatrix(.Row, 3) = vlintCantidad
                    .TextMatrix(.Row, 4) = Format(vldblPrecio * CLng(vlintCantidad), "$###,###,###,###.00")
                    .TextMatrix(.Row, 5) = Format(vldblTotDescuento, "$###,###,###,###.00")
                    .TextMatrix(.Row, 6) = Format(vldblIEPS, "$###,###,###,###.00") '@
                    .TextMatrix(.Row, 7) = Format(vldblSubtotal, "$###,###,###,###.00") '@
                    .TextMatrix(.Row, 8) = Format(vldblSubtotal * vldblIVA, "$###,###,###,###.00")
                    .TextMatrix(.Row, 9) = Format(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 8)), "$###,###,###,###.00")
                    .Col = 9
                    .CellFontBold = True
                    .TextMatrix(.Row, 10) = vllngClaveElemento
                    .TextMatrix(.Row, 11) = vldblSubtotal * vldblIVA
                    .TextMatrix(.Row, 12) = vlstrTipoCargo
                    .TextMatrix(.Row, 13) = rs!Concepto  'Clave del concepto de Facturación
                    .TextMatrix(.Row, 14) = vlintModoDescuentoInventario 'Descuento por unidad alterna(2) o unidad mínima(1)
                    .TextMatrix(.Row, 16) = vllngContenido   'Contenido de IVArticulo
                    .TextMatrix(.Row, 17) = lCveLista   'Clave de la lista de precios
                    .TextMatrix(.Row, 18) = vldblPorcentajeDescuento 'Porcentaje de descuento
                    .TextMatrix(.Row, 19) = vldblIEPSPercent 'Porcentaje de IEPS
                    .TextMatrix(.Row, 20) = vlintAuxExclusionDescuento '2 = tiene exclusión de descuento, 3 = no tiene exclusion de descuento
                    .TextMatrix(.Row, 21) = rs!ConceptoDesc 'Descripción del concepto de Facturación
                    .TextMatrix(.Row, 23) = "0"
                    .RowData(.Row) = vllngClaveElemento
                    .Redraw = True
                    .Refresh
                    
                    
                   lngNumCargo = Val(.RowData(.Row))
                   
                   'Caso 20370
                    'Validar el stock de inventario Vs Solicitud de salidas(venta al publico)
                    'GIRM
                    
                    If BlnValidarCantidad(vlintUM, vlintUV, vllngClaveElemento, vlintModoDescuentoInventario) Then
                        MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(317) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
                        grdArticulos.RemoveItem (grdArticulos.Row)
                        txtClaveArticulo.Text = ""
                        Exit Sub
                    End If
                    'Caso 20370
                   
                    'Manejos de caducidades
                    pManejaCaducidad (vllngClaveElemento), "S", vlstrCveArticulo, vlintModoDescuentoInventario
                        
                    'Si no selecciono el lote no lo agregara al Grid caso 20274
                    If vgblnCapturoLoteYCaduc = False Then
                        If grdArticulos.Rows > 2 Then
                            grdArticulos.RemoveItem grdArticulos.Rows - 1
                        Else
                            pLimpiaGrid
                        End If
                        Exit Sub
                    End If
'
'                    'Manejos de caducidades
'                    pManejaCaducidad (lngNumCargo), "S", vlstrCveArticulo, vlintModoDescuentoInventario
'
                    
                    
                    If strVariosLotes = "" Then
                        'Cambio la cantidad
                        .TextMatrix(.Row, 3) = lngCantidad
                        'Clave del articulo
                        .TextMatrix(.Row, 30) = strCveArticulo

                        'Lote de trazabilidad
                        .TextMatrix(.Row, 31) = strLote
                    Else
                        'Cambio la cantidad
                        .TextMatrix(.Row, 3) = lngCantidad

                        'Varios lotes en un mismos producto
                        .TextMatrix(.Row, 32) = strVariosLotes
                        strVariosLotes = ""
                    End If

                    
                    pCalculaTotales
                Else
                    txtConsultaPrecio.Text = ""
                    txtConsultaDescuento.Text = ""
                    txtConsultaDescripcion.Text = ""
                    txtConsultaDescripcion.Text = Trim(rs!Descripcion)
                    txtConsultaPrecio.Text = Format(vldblPrecio + (vldblPrecio * vldblIEPSPercent), "$###,###,###,###.00")
                    txtConsultaDescuento.Text = Format(vldblTotDescuento, "$###,###,###,###.00")
                    frePrecios.Top = 0
                    frePrecios.Left = 2800
                    frePrecios.Visible = True
                    frePrecios.ZOrder 0
                    vgstrEstadoManto = "P"
                    pBuscarPrecio True
                End If
                rs.Close
            Else
                MsgBox SIHOMsg(3), vbCritical, "Mensaje"
            End If
            vlintCantidad = 1
            txtClaveArticulo.Text = ""
            
            ' Deshabilitar el frame de pedida de datos
            If grdArticulos.RowData(1) <> -1 Then
                freDatosPaciente.Enabled = False
                If txtMovimientoPaciente.Text = "" Then
                    optPaciente.Value = True
                    lblPaciente.Caption = "VENTA AL PÚBLICO"
                    vgstrRFCPersonaSeleccionada = ""
                    vlStrTipoPaciente = "Select vchDescripcion from adtipopaciente where tnycvetipopaciente = " & vgintTipoPaciente
                    Set rsTipoPaciente = frsRegresaRs(vlStrTipoPaciente, adLockReadOnly, adOpenForwardOnly)
                    
                    If rsTipoPaciente.RecordCount > 0 Then
                        lblEmpresa.Caption = rsTipoPaciente!VCHDESCRIPCION
                    Else
                        lblEmpresa.Caption = "PARTICULAR"
                    End If
                End If
                lblBusquedaMedicos.FontBold = False
            lblBusquedaEmpleados.FontBold = False
            lblBusquedaPacientes.FontBold = False
            End If
        End With
    End If
    'limpiar
    vgblnCapturoLoteYCaduc = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaElementoTrazabilidad"))
    
End Sub
'Private Sub pSeleccionaElementoTrazabilidad2(vllngClaveElemento As Long, lngIdEtiqueta As Long, vlintCantidad As Long, vlstrTipoCargo As String)
'
'On Error GoTo NotificaError
'
'    Dim rs As New ADODB.Recordset
'    Dim rsVentaAlmacen As New ADODB.Recordset
'    Dim rsTipoPaciente As New ADODB.Recordset
'
'    Dim vldblPrecio As Double
'    Dim vldblIncrementoTarifa As Double 'Incremento en la tarifa de los precios
'    Dim vldblPorcentajeDescuento As Double
'    Dim vldblTotDescuento As Double
'    Dim vldblSubtotal As Double
'    Dim vldblIEPS As Double
'    Dim vldblIEPSPercent As Double
'    Dim vldblIVA As Double
'
'    Dim vlintModoDescuentoInventario As Integer
'    Dim DescuentoInventario As Integer
'    Dim MNYCantidad As Integer
'    Dim vlintAlmacenVenta As Integer
'
'    Dim vllngContenido As Long
'    Dim vlngExistencia As Long
'    Dim vlintAuxExclusionDescuento As Long
'    Dim lngNumCargo As Long
'
'    Dim vlstrCveArticulo As String
'    Dim vlchrDescripcion As String
'    Dim vlchrTipoDescuento As String
'    Dim vlStrTipoPaciente As String
'
'
'    Dim vlbExistencias As Boolean
'
'    If Val(Format(vlintCantidad, "")) > 0 Then
'        With grdArticulos
'            vllngContenido = 1  'Este es el Contenido en IvArticulo
'            If vllngClaveElemento <> -1 Then
'
'                'Tipo de descuento de Inventario
'                vlstrSentencia = "SELECT intContenido Contenido,vchNombreComercial Descripcion, substring(vchNombreComercial,1,50) Articulo,  ivUA.vchDescripcion UnidadAlterna,  ivUM.vchDescripcion UnidadMinima," & _
'                                    " chrCveArticulo CveArticulo, ivArticulo.bitVentaPublico" & _
'                                    " From ivArticulo " & _
'                                    " inner Join ivUnidadVenta ivUA on ivUA.intCveUnidadVenta = ivArticulo.intCveUniAlternaVta " & _
'                                    " inner Join ivUnidadVenta ivUM on ivUM.intCveUnidadVenta = ivArticulo.intCveUniMinimaVta " & _
'                                    " WHERE intIdArticulo = " & Trim(str(vllngClaveElemento))
'
'                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'                    If rs.RecordCount > 0 Then
'                        If rs!bitVentaPublico = 0 Then
'                            MsgBox SIHOMsg(1548), vbInformation, "Mensaje"
'                            rs.Close
'                            Exit Sub
'                        End If
'
'                        vlstrCveArticulo = rs!cveArticulo
'                        vlintModoDescuentoInventario = 2 'Descuento por Unidad Alterna
'                        vllngContenido = rs!Contenido 'Este es el contenido de IVarticulo
'                        vlchrDescripcion = rs!Descripcion
'
'
'                        If vllngContenido > 1 Then
'                            If MsgBox("¿Desea realizar la venta de " & Trim(rs!Articulo) & " por " & Trim(rs!UnidadAlterna) & "?" & Chr(13) & "Si selecciona NO, se venderá por " & Trim(rs!UnidadMinima) & ".", vbYesNo + vbQuestion, "Mensaje") = vbNo Then
'                                vlintModoDescuentoInventario = 1 'Descuento por unidad Minima
'                            End If
'                        End If
'                        rs.Close
'                    End If
'
'
'                    'Precio unitario
'                    pCargaArreglo vlaryParametrosSalida, "|" & adDecimal & "||" & adDecimal
'                    frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|1|01/01/1900|" & vgintClaveEmpresaContable, "SP_PVSELOBTENERPRECIO", , , vlaryParametrosSalida
'                    pObtieneValores vlaryParametrosSalida, vldblPrecio, vldblIncrementoTarifa
'
'                    If vldblPrecio = -1 Or vldblPrecio = 0 Then
'                        MsgBox SIHOMsg(301), vbInformation, "Mensaje"
'                        Exit Sub
'                    End If
'
'                    'Existencias
'                    vlbExistencias = False
'                    DescuentoInventario = vlintModoDescuentoInventario
'                    MNYCantidad = vlintCantidad
'
'                    Set rsVentaAlmacen = frsRegresaRs("SELECT smiCveDepartamento FROM NoDepartamento WHERE NoDepartamento.chrClasificacion = 'A' AND smiCveDepartamento = " & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
'
'                    If Not rsVentaAlmacen.RecordCount = 0 Then
'                        vlintAlmacenVenta = rsVentaAlmacen!smicvedepartamento
'                    Else
'                        Set rsVentaAlmacen = frsRegresaRs("SELECT intNumAlmacen FROM PvAlmacenes WHERE intnumdepartamento =" & vgintNumeroDepartamento, adLockOptimistic, adOpenDynamic)
'                        vlintAlmacenVenta = rsVentaAlmacen!intnumalmacen
'                    End If
'
'                    'vlstrSentencia = "SELECT INTEXISTENCIADEPTOUV VL_UV, INTEXISTENCIADEPTOUM VL_UM FROM IVUBICACION WHERE chrcvearticulo = " & vlstrCveArticulo & " AND smicvedepartamento = " & vlintAlmacenVenta
'                    vlstrSentencia = "Select Existencia from IVLOTEPROCEDENCIA WHERE INTIDLOTEPROCEDENCIA =(select INTIDLOTEPROCEDENCIA from ivetiqueta where INTIDETIQUETA =" & lngIdEtiqueta & ")"
'
'                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'
'                    If rs!EXISTENCIA = Null Then
'
'                        vlngExistencia = 0
'                    End If
'
'                    If rs.RecordCount > 0 Then
'                        If rs!EXISTENCIA > 0 Then
'                            vlngExistencia = rs!EXISTENCIA
'
'
'                        Else
'                            vlngExistencia = 0
'                        End If
'                    Else
'                        vlngExistencia = 0
'                    End If
'
'                    If DescuentoInventario = 1 And vllngContenido > 1 Then
'                        If vlngExistencia < MNYCantidad Then
'                            vlbExistencias = True
'                        End If
'                     End If
'
'                    If (vlngExistencia < MNYCantidad) And DescuentoInventario = 2 Then
'                        vlbExistencias = True
'                    End If
'
'                    If vlbExistencias Then
'                        If chkConsultaPrecios.Value = 0 Then
'                           MsgBox "No se pudo completar la operación." & Chr(13) & SIHOMsg(317) & "(" & Trim(vlchrDescripcion) & ")", vbExclamation, "Mensaje"
'                           Exit Sub
'                        End If
'                    End If
'
'                    'Clave de la lista de precios
'                    pCargaArreglo vlaryParametrosSalida, "|" & adInteger
'                    frsEjecuta_SP vllngClaveElemento & "|" & vlstrTipoCargo & "|" & vgintTipoPaciente & "|" & vgintEmpresa & "|E|" & vgintClaveEmpresaContable, "SP_PVSELLISTAPRECIO", , , vlaryParametrosSalida
'                    pObtieneValores vlaryParametrosSalida, lCveLista
'
'
'                    'Exclusión de descuentos
'                    vlintAuxExclusionDescuento = 1
'                    frsEjecuta_SP "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & vlstrTipoCargo & "|" & vllngClaveElemento & "|" & vgintNumeroDepartamento, "FN_PVSELEXCLUSIONDESCUENTO", True, vlintAuxExclusionDescuento
'
'                    If vlintAuxExclusionDescuento = 2 Then
'                        vldblTotDescuento = 0
'                    Else
'                        'Descuentos
'                        vldblTotDescuento = 0
'                        pCargaArreglo vlaryParametrosSalida, "|" & adDecimal
'                        vgstrParametrosSP = "E|" & vgintTipoPaciente & "|" & vgintEmpresa & "|" & CStr(Val(txtMovimientoPaciente.Text)) & "|" & vlstrTipoCargo & "|" & CStr(vllngClaveElemento) & "|" & vldblPrecio & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(fdtmServerFecha) & "|" & Trim(str(intAplicado)) & "|" & CStr(vllngContenido) & "|" & CStr(vlintCantidad) & "|" & CStr(vlintModoDescuentoInventario)
'                        frsEjecuta_SP vgstrParametrosSP, "sp_PvSelDescuentoCantidad", , , vlaryParametrosSalida
'                        pObtieneValores vlaryParametrosSalida, vldblTotDescuento
'                    End If
'
'                    'El Precio del artículo, según el tipo de descuento y CONTENIDO de Ivarticulo
'                    If vlintModoDescuentoInventario = 1 Then
'                        vldblPrecio = vldblPrecio / CDbl(vllngContenido)
'                    End If
'
'                    vldblSubtotal = (vldblPrecio * CLng(vlintCantidad)) - vldblTotDescuento
'
'                    ' Porcentaje de IEPS
'                    If vlblnLicenciaIEPS And vlstrTipoCargo = "AR" Then
'                       vldblIEPSPercent = fdblObtenerPorcentajeIEPS(CLng(vllngClaveElemento)) / 100 '<----------------
'                    Else
'                       vldblIEPSPercent = 0
'                    End If
'
'                    vldblIEPS = vldblSubtotal * vldblIEPSPercent '<---sacamos cantidada de IEPS
'                    vldblSubtotal = vldblSubtotal + vldblIEPS '<-----sumamos el IEPS al Subtotal
'
'                    ' Procedimiento para obtener % del IVA
'                    vldblIVA = fdblObtenerIva(CLng(vllngClaveElemento), vlstrTipoCargo) / 100
'
'                    ' Datos del Artículo ó del Otro concepto
'                    If optArticulo.Value Then
'                        vlstrSentencia = "select AR.vchNombreComercial Descripcion, AR.smiCveConceptFact Concepto, CF.chrDescripcion ConceptoDesc from IVArticulo AR inner join PVConceptoFacturacion CF on CF.smiCveConcepto = AR.smiCveConceptFact where AR.intIdArticulo = " & Trim(str(vllngClaveElemento))
'                    End If
'
'                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
'
'                    ' Llenado del grid o de la consulta de precios
'                    If chkConsultaPrecios.Value = 0 Then
'
'                        lblnGuardoControl = False
'
'                        If .RowData(1) <> -1 Then
'                            .Rows = .Rows + 1
'                            .Row = .Rows - 1
'                        End If
'
'                        .TextMatrix(.Row, 1) = rs!Descripcion
'                        .TextMatrix(.Row, 2) = Format(vldblPrecio, "$###,###,###,###.00")
'                        .TextMatrix(.Row, 24) = vldblPrecio
'
'                        .TextMatrix(.Row, 3) = vlintCantidad
'                        .TextMatrix(.Row, 4) = Format(vldblPrecio * CLng(vlintCantidad), "$###,###,###,###.00")
'                        .TextMatrix(.Row, 5) = Format(vldblTotDescuento, "$###,###,###,###.00")
'                        .TextMatrix(.Row, 6) = Format(vldblIEPS, "$###,###,###,###.00") '@
'                        .TextMatrix(.Row, 7) = Format(vldblSubtotal, "$###,###,###,###.00") '@
'                        .TextMatrix(.Row, 8) = Format(vldblSubtotal * vldblIVA, "$###,###,###,###.00")
'                        .TextMatrix(.Row, 9) = Format(CDbl(.TextMatrix(.Row, 7)) + CDbl(.TextMatrix(.Row, 8)), "$###,###,###,###.00")
'                        .Col = 9
'                        .CellFontBold = True
'                        .TextMatrix(.Row, 10) = vllngClaveElemento
'                        .TextMatrix(.Row, 11) = vldblSubtotal * vldblIVA
'                        .TextMatrix(.Row, 12) = vlstrTipoCargo
'                        .TextMatrix(.Row, 13) = rs!Concepto  'Clave del concepto de Facturación
'                        .TextMatrix(.Row, 14) = vlintModoDescuentoInventario 'Descuento por unidad alterna(2) o unidad mínima(1)
'                        .TextMatrix(.Row, 16) = vllngContenido   'Contenido de IVArticulo
'                        .TextMatrix(.Row, 17) = lCveLista   'Clave de la lista de precios
'                        .TextMatrix(.Row, 18) = vldblPorcentajeDescuento 'Porcentaje de descuento
'                        .TextMatrix(.Row, 19) = vldblIEPSPercent 'Porcentaje de IEPS
'                        .TextMatrix(.Row, 20) = vlintAuxExclusionDescuento '2 = tiene exclusión de descuento, 3 = no tiene exclusion de descuento
'                        .TextMatrix(.Row, 21) = rs!ConceptoDesc 'Descripción del concepto de Facturación
'                        .TextMatrix(.Row, 23) = "0"
'                        .TextMatrix(.Row, 29) = txtClaveArticulo.Text
'                        .RowData(.Row) = vllngClaveElemento
'                        .Redraw = True
'                        .Refresh
'
'                        lngNumCargo = Val(.RowData(.Row))
'
'                        'Manejos de caducidades
'                        pManejaCaducidad (lngNumCargo), "S", vlstrCveArticulo, vlintModoDescuentoInventario
'
'
'                        If strVariosLotes = "" Then
'                            'Cambio la cantidad
'                            .TextMatrix(.Row, 3) = lngCantidad
'                            'Clave del articulo
'                            .TextMatrix(.Row, 30) = StrCveArticulo
'
'                            'Lote de trazabilidad
'                            .TextMatrix(.Row, 31) = strLote
'                        Else
'                            'Cambio la cantidad
'                            .TextMatrix(.Row, 3) = lngCantidad
'
'                            'Varios lotes en un mismos producto
'                            .TextMatrix(.Row, 32) = strVariosLotes
'                            strVariosLotes = ""
'                        End If
'
'                        'Totales
'                        pCalculaTotales
'
'                    Else
'                        txtConsultaPrecio.Text = ""
'                        txtConsultaDescuento.Text = ""
'                        txtConsultaDescripcion.Text = ""
'                        txtConsultaDescripcion.Text = Trim(rs!Descripcion)
'                        txtConsultaPrecio.Text = Format(vldblPrecio + (vldblPrecio * vldblIEPSPercent), "$###,###,###,###.00")
'                        txtConsultaDescuento.Text = Format(vldblTotDescuento, "$###,###,###,###.00")
'                        frePrecios.Top = 0
'                        frePrecios.Left = 2800
'                        frePrecios.Visible = True
'                        frePrecios.ZOrder 0
'                        vgstrEstadoManto = "P"
'                        pBuscarPrecio True
'                    End If
'                    rs.Close
'                Else
'                    MsgBox SIHOMsg(3), vbCritical, "Mensaje"
'                End If
'
'                vlintCantidad = 1
'                txtClaveArticulo.Text = ""
'
'                ' Deshabilitar el frame de pedida de datos
'                If grdArticulos.RowData(1) <> -1 Then
'                    freDatosPaciente.Enabled = False
'                    If txtMovimientoPaciente.Text = "" Then
'                        optPaciente.Value = True
'                        lblPaciente.Caption = "VENTA AL PÚBLICO"
'                        vgstrRFCPersonaSeleccionada = ""
'                        vlStrTipoPaciente = "Select vchDescripcion from adtipopaciente where tnycvetipopaciente = " & vgintTipoPaciente
'                        Set rsTipoPaciente = frsRegresaRs(vlStrTipoPaciente, adLockReadOnly, adOpenForwardOnly)
'
'                        If rsTipoPaciente.RecordCount > 0 Then
'                            lblEmpresa.Caption = rsTipoPaciente!VCHDESCRIPCION
'                        Else
'                            lblEmpresa.Caption = "PARTICULAR"
'                        End If
'                    End If
'
'                    lblBusquedaMedicos.FontBold = False
'                    lblBusquedaEmpleados.FontBold = False
'                    lblBusquedaPacientes.FontBold = False
'            End If
'        End With
'    End If
'
'Exit Sub
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaElementoTrazabilidad"))
'
'
'End Sub

Public Sub pManejaCaducidad(lngNumCargo As Long, vlstrTipoMov As String, vlstrCveArticulo As String, vlintUnidad As Integer)
    On Error GoTo NotificaError
    
    Dim vllngContenido As Long

    Dim vlstrDescUnidadAlterna As String
    Dim vlstrDescUnidadMinima As String
    Dim vlstrLote As String
    Dim vlstrFechaCaduca As String
    
    Dim rs As ADODB.Recordset
    Dim rsDepto As ADODB.Recordset
    
    Const intColdescripcion = 1
    
    vllngContenido = 0
    
    'MANEJO DE CADUCIDADES POR MEDIO DEL LOTE
    'LMM 20274
     Set rsDepto = frsRegresaRs(" select INTNUMALMACEN from  pvAlmacenes WHERE INTNUMDEPARTAMENTO = " & vgintNumeroDepartamento)
     If rsDepto.RecordCount > 0 Then
         vllngdeptolote = rsDepto!intnumalmacen
     End If

    Set rs = frsRegresaRs(" select intContenido, intIdArticulo, IA.VCHDESCRIPCION ALTERNA ,UM.VCHDESCRIPCION MINIMA" & _
                          " from IvArticulo " & _
                          " LEFT OUTER join IvUnidadVentA IA on IA.INTCVEUNIDADVENTA = ivarticulo.INTCVEUNIALTERNAVTA " & _
                          " LEFT OUTER join IvUnidadVenta UM on UM.INTCVEUNIDADVENTA = ivarticulo.INTCVEUNIMINIMAVTA " & _
                          " where chrCveArticulo = '" & Trim(vlstrCveArticulo) & "' ")
                          
    If rs.RecordCount > 0 Then
        vllngContenido = rs!intContenido
        vlstrDescUnidadAlterna = rs!alterna
        vlstrDescUnidadMinima = rs!MINIMA
    End If
    
    Set rs = frsRegresaRs("select INTIDETIQUETA, CHRCVEARTICULO, CHRLOTE, DTMFECHACADUCIDAD, INTIDLOTEPROCEDENCIA " & _
                          "from ivetiqueta where INTIDETIQUETA = " & Trim(txtClaveArticulo.Text))
                          
    If rs.RecordCount > 0 Then
        vlstrLote = Trim(rs!chrlote)
        vlstrFechaCaduca = rs!dtmFechaCaducidad
    End If
                    
    If vlstrTipoMov = "S" Then
        frmCapturaLotePV.vlstrTablaReferencias = "SVP"
        frmCapturaLotePV.vlstrChrCveArticulo = vlstrCveArticulo
        frmCapturaLotePV.vlstrNoMovimiento = Val(lngNumCargo)
        frmCapturaLotePV.vlstrTipoMovimiento = IIf(vlintUnidad = 1, "UM", "UV")
        frmCapturaLotePV.vlintContenidoArt = vllngContenido
        frmCapturaLotePV.vlStrTitUM = StrConv(vlstrDescUnidadMinima, vbProperCase)
        frmCapturaLotePV.vlStrTitUV = StrConv(vlstrDescUnidadAlterna, vbProperCase)
        frmCapturaLotePV.txtDescripcionLgaArt.Text = grdArticulos.TextMatrix(grdArticulos.Row, intColdescripcion)
        frmCapturaLotePV.txtCantRecep.Text = Val(txtCantidad)
        frmCapturaLotePV.txtTotalARecibir.Text = Val(txtCantidad)
        frmCapturaLotePV.txtLote.Text = vlstrLote
        frmCapturaLotePV.txtFechaCaduce.Text = vlstrFechaCaduca
        frmCapturaLotePV.vlintNoDepartamento = vllngdeptolote 'ajuste caso 20274'vgintNumeroDepartamento
        frmCapturaLotePV.TxtCBarra.Text = txtClaveArticulo.Text
        frmCapturaLotePV.Label2.Caption = "Total salida"
        
        Load frmCapturaLotePV
    
        frmCapturaLotePV.Show vbModal, Me
        'pGrabaLotes Val(lngNumCargo), "SVP", CInt(vgintNumeroDepartamento), "S", True
    Else
        frmCapturaLotePV.vlstrTablaReferencias = "ECCD"
        frmCapturaLotePV.vlstrChrCveArticulo = vlstrCveArticulo
        frmCapturaLotePV.vlstrNoMovimiento = Val(lngNumCargo)
        frmCapturaLotePV.vlstrTipoMovimiento = IIf(vllngContenido > 1, "UM", "UV")
        frmCapturaLotePV.vlintContenidoArt = vllngContenido
        frmCapturaLotePV.vlStrTitUM = StrConv(vlstrDescUnidadMinima, vbProperCase)
        frmCapturaLotePV.vlStrTitUV = StrConv(vlstrDescUnidadAlterna, vbProperCase)
        frmCapturaLotePV.vlintNoDepartamento = vgintNumeroDepartamento
        frmCapturaLotePV.vlblnValidaMovimiento = False
        Load frmCapturaLotePV
        
        frmCapturaLotePV.Show vbModal, Me
        'pGrabaLotes Val(lngNumCargo), "ECCD", CInt(vgintNumeroDepartamento), "E", True
    End If
    rs.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pManejaCaducidad"))
    Unload Me
End Sub

Private Sub pNuevaTrazabilidad(lngConsecutivoVenta As Long, lngIdProcedencia As Long, intRow As Integer)
    
    
    Dim lngRegistros As Long
    Dim lngCantidad As Long
    Dim strQry As String
    Dim strLotes() As String
    Dim strRenglon() As String
    Dim strRenglones As String
    Dim vllngRenglon As Long
    
On Error GoTo NotificaError

    If grdArticulos.TextMatrix(intRow, 32) = "" Then
    
        If grdArticulos.TextMatrix(intRow, 30) <> "" And grdArticulos.TextMatrix(intRow, 31) <> "" And grdArticulos.TextMatrix(intRow, 29) <> "" Then
    
            'Agrega el registro de la trazabilidad
            GuardarVentaTrazabilidas grdArticulos.TextMatrix(intRow, 30), grdArticulos.TextMatrix(intRow, 31), CLng(grdArticulos.TextMatrix(intRow, 29)), lngConsecutivoVenta
            lngCantidad = lngCantidad + grdArticulos.TextMatrix(intRow, 3)
            
            'Descuenta la cantidad tomada de IVLOTEPROCEDENCIA
            
            pEjecutaSentencia "UPDATE ivloteprocedencia SET existencia = existencia - " & lngCantidad & " WHERE intidloteprocedencia =" & lngIdProcedencia
            
        End If
    
    Else
        strLotes = Split(grdArticulos.TextMatrix(intRow, 32), "|")
        For vllngRenglon = 0 To UBound(strLotes)
            strRenglon = Split(strLotes(vllngRenglon), ";")
            GuardarVentaTrazabilidas strRenglon(1), strRenglon(2), CLng(strRenglon(3)), lngConsecutivoVenta
            
            lngCantidad = lngCantidad + grdArticulos.TextMatrix(intRow, 3)
            
            'Descuenta la cantidad tomada de IVLOTEPROCEDENCIA
            pEjecutaSentencia "UPDATE ivloteprocedencia SET existencia = existencia - " & lngCantidad & " WHERE intidloteprocedencia =" & lngIdProcedencia
        Next vllngRenglon
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNuevaTrazabilidad"))

End Sub

'Funcion que verifica si se hace uso del bit de trazabilidad
Private Function fblnrevisaUsoTrazabilidad() As Boolean
    Dim strQry As String
    Dim rs As ADODB.Recordset
    
    fblnrevisaUsoTrazabilidad = False
    
    strQry = "SELECT VCHVALOR FROM SIPARAMETRO WHERE VCHNOMBRE ='BITTRAZABILIDAD' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(strQry, adLockOptimistic, adOpenDynamic)
    
    If rs.RecordCount > 0 Then
        If rs!VCHVALOR = "1" Then
            fblnrevisaUsoTrazabilidad = True
        End If
    End If
    rs.Close
End Function


Private Sub tValidarExistencias(blnTrazabilidad As Boolean, intcontador As Long)
    Dim rsTrazabilidas As ADODB.Recordset
    
    If blnTrazabilidad = True Then
        Set rsTrazabilidas = frsRegresaRs("SELECT INTIDLOTEPROCEDENCIA, EXISTENCIA FROM ivloteprocedencia WHERE intidloteprocedencia = (SELECT intidloteprocedencia FROM ivetiqueta WHERE intidetiqueta = '" & grdArticulos.TextMatrix(intcontador, 29) & "')", adLockOptimistic, adOpenDynamic)
        If rsTrazabilidas.RecordCount > 0 Then
            'If Val(.TextMatrix(vlintContador, 15)) < 0 Then
            vlLngIvLoteProcedencia = rsTrazabilidas!INTIDLOTEPROCEDENCIA
            grdArticulos.TextMatrix(intcontador, 15) = rsTrazabilidas!EXISTENCIA
            'End If
        End If
        rsTrazabilidas.Close
    End If
End Sub


Private Sub GuardarVentaTrazabilidas(vlstrCveArticulo As String, vlstrLote As String, vllngLote As Long, lngConsecutivoVenta As Long)
    Dim rsTrazabilidad As New ADODB.Recordset
    Dim rs As New ADODB.Recordset

    'Se obtiene el identificador unicos
    Set rs = frsRegresaRs("select SEC_PVVENTAPUBLICOLOTE.nextval  as Id from dual", adLockOptimistic, adOpenDynamic)

    'se graba la informacion si el bit de trazabilidad esta encendido
    Set rsTrazabilidad = frsRegresaRs("SELECT * FROM PVVENTAPUBLICOLOTE WHERE intConsecutivo = -1", adLockOptimistic, adOpenDynamic)
    With rsTrazabilidad
        .AddNew
        !chrcvearticulo = vlstrCveArticulo
        !VCHLOTE = vlstrLote
        !DTMFECHAVENTA = fdtmServerFecha
        !INTCVEVENTA = lngConsecutivoVenta
        !intEmpleado = vllngPersonaGraba
        !IntLote = vllngLote
        !intConsecutivo = rs!id
        .Update
    End With
       
    rs.Close
    rsTrazabilidad.Close
       
End Sub

'Caso 20370
Function grdCantidad(LngCveArticulo As Long, IntModoDescuentoInventario As Integer) As Long
    Dim contador As Long, i As Long
    contador = 0
    For i = 1 To grdArticulos.Rows - 1
        'If grdArticulos.TextMatrix(i, 10) <> "" Then
            If LngCveArticulo = CLng(grdArticulos.TextMatrix(i, 10)) And IntModoDescuentoInventario = CLng(grdArticulos.TextMatrix(i, 14)) Then
                contador = contador + CLng(grdArticulos.TextMatrix(i, 3))
            End If
        'End If
    Next i
    grdCantidad = contador
End Function

Function BlnValidarCantidad(LngUM As Long, LngUV As Long, LngCveArticulo As Long, vlintModoDescuentoInventario As Integer) As Boolean
    Dim Validar As Boolean
    Validar = False
    
    'caso 20384
    If vlintModoDescuentoInventario = 1 Then
        'unidades minima
        If LngUM <> 0 Then
            If LngUM < grdCantidad(LngCveArticulo, vlintModoDescuentoInventario) Then
                Validar = True
            End If
        End If
    Else
        'unidad alternitiva
        If LngUV <> 0 Then
            If LngUV < grdCantidad(LngCveArticulo, vlintModoDescuentoInventario) Then
              Validar = True
            End If
        End If
    End If
    'caso 20384
    
    BlnValidarCantidad = Validar
    
End Function

'caso 20370

'Caso 20417
Function RoundUP(vldblNumero As Double) As Long
    If vldblNumero > Int(vldblNumero) Then
        RoundUP = Int(vldblNumero) + 1
    Else
        RoundUP = Int(vldblNumero)
    End If
End Function
'Caso 20417

'caso 20453
Function LngCveArticulo(strCveArticulo As String) As Long
    Dim rs As New ADODB.Recordset
    Dim strSentencia As String
    
    LngCveArticulo = 0
    
    strSentencia = "SELECT intidarticulo FROM ivarticulo WHERE chrcvearticulo='" & strCveArticulo & "'"
    
    Set rs = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then
        LngCveArticulo = rs!intIdArticulo
    End If
End Function

'caso 20453
