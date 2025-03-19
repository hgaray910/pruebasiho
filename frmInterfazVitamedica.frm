VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInterfazVitamedica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interfaz con Vitamédica"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTVitamedica 
      Height          =   9585
      Left            =   -8
      TabIndex        =   32
      Top             =   -330
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   16907
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmInterfazVitamedica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FreBotonera"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdQuitarSeleccion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdInvertirSeleccion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSeleccionarTodos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FrePaciente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CDgArchivo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FreDetalle"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmInterfazVitamedica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FreDetalleTXT"
      Tab(1).Control(1)=   "FreFiltros"
      Tab(1).ControlCount=   2
      Begin VB.Frame FreDetalle 
         Enabled         =   0   'False
         Height          =   3105
         Left            =   100
         TabIndex        =   43
         Top             =   4560
         Width           =   13575
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargos 
            Height          =   2750
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Cargos en la cuenta del paciente"
            Top             =   225
            Width           =   13320
            _ExtentX        =   23495
            _ExtentY        =   4842
            _Version        =   393216
            Cols            =   11
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FormatString    =   "|Fecha|Descripción|Cantidad|Precio|Importe|Descuento|Subtotal|IVA|Total|Factura"
            BandDisplay     =   1
            RowSizingMode   =   1
            _NumberOfBands  =   1
            _Band(0).BandIndent=   5
            _Band(0).Cols   =   11
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame FreDetalleTXT 
         Enabled         =   0   'False
         Height          =   7000
         Left            =   -74900
         TabIndex        =   49
         Top             =   2100
         Visible         =   0   'False
         Width           =   13575
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTXT 
            Height          =   6630
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   "Archivos de texto para Vitamédica"
            Top             =   225
            Width           =   13320
            _ExtentX        =   23495
            _ExtentY        =   11695
            _Version        =   393216
            Cols            =   7
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FormatString    =   "|Fecha|Número cuenta |Nombre paciente|Empresa |Tipo de comanda|Persona envió"
            RowSizingMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   7
         End
      End
      Begin VB.Frame FreFiltros 
         Height          =   1600
         Left            =   -74895
         TabIndex        =   44
         ToolTipText     =   "Filtros de búsqueda de archivos TXT"
         Top             =   400
         Width           =   13575
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar"
            Height          =   390
            Left            =   11040
            TabIndex        =   56
            ToolTipText     =   "Cargar la información con los filtros seleccionados"
            Top             =   990
            Width           =   2340
         End
         Begin VB.OptionButton OptPaciente 
            Caption         =   "Externo"
            Height          =   255
            Index           =   3
            Left            =   5340
            TabIndex        =   54
            ToolTipText     =   "Paciente de tipo externo"
            Top             =   645
            Width           =   975
         End
         Begin VB.OptionButton OptPaciente 
            Caption         =   "Interno"
            Height          =   255
            Index           =   2
            Left            =   4305
            TabIndex        =   53
            ToolTipText     =   "Paciente de tipo interno"
            Top             =   645
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox txtNombrePaciente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "Nombre del paciente"
            Top             =   990
            Width           =   5370
         End
         Begin VB.TextBox txtCuentaPaciente 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1575
            MaxLength       =   10
            TabIndex        =   52
            ToolTipText     =   "Número de cuenta del paciente"
            Top             =   630
            Width           =   1605
         End
         Begin MSMask.MaskEdBox mskFechaInicial 
            Height          =   315
            Left            =   1575
            TabIndex        =   50
            ToolTipText     =   "Fecha de inicio de generación del archivo de texto"
            Top             =   255
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFechaFinal 
            Height          =   315
            Left            =   5340
            TabIndex        =   51
            ToolTipText     =   "Fecha de fin de generación del archivo de texto"
            Top             =   255
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha final"
            Height          =   195
            Left            =   4305
            TabIndex        =   48
            Top             =   315
            Width           =   1035
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha inicial"
            Height          =   195
            Left            =   150
            TabIndex        =   47
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   150
            TabIndex        =   46
            ToolTipText     =   "Número de cuenta"
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Nombre"
            Height          =   195
            Left            =   150
            TabIndex        =   45
            ToolTipText     =   "Nombre del paciente"
            Top             =   1050
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog CDgArchivo 
         Left            =   480
         Top             =   8430
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame FrePaciente 
         Height          =   4050
         Left            =   100
         TabIndex        =   34
         Top             =   400
         Width           =   13575
         Begin VB.TextBox txtDescripcionProc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7500
            MaxLength       =   150
            TabIndex        =   21
            ToolTipText     =   "Descripción del procedimiento"
            Top             =   3240
            Width           =   3150
         End
         Begin VB.TextBox txtDescripcionDiagnostico 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7500
            MaxLength       =   150
            TabIndex        =   20
            ToolTipText     =   "Descripción del diagnóstico"
            Top             =   2880
            Width           =   3150
         End
         Begin VB.TextBox txtCodigoCPT 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   11
            ToolTipText     =   "Código CPT"
            Top             =   3240
            Width           =   3600
         End
         Begin VB.TextBox txtCodigoICD 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   10
            ToolTipText     =   "Código ICD"
            Top             =   2880
            Width           =   3600
         End
         Begin VB.TextBox txtClaveMedico 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   12
            ToolTipText     =   "Clave del médico"
            Top             =   3600
            Width           =   3600
         End
         Begin VB.TextBox txtBeneficiario 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   16
            TabIndex        =   8
            ToolTipText     =   "Número de beneficiario"
            Top             =   2160
            Width           =   3600
         End
         Begin VB.TextBox txtNomina 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   7
            ToolTipText     =   "Número de nómina"
            Top             =   1800
            Width           =   3600
         End
         Begin VB.ComboBox cboTipoFactura 
            Height          =   315
            Left            =   7500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "Tipo de factura"
            Top             =   2520
            Width           =   3150
         End
         Begin VB.ComboBox cboTipoEgreso 
            Height          =   315
            Left            =   7500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Tipo de egreso"
            Top             =   1800
            Width           =   3150
         End
         Begin VB.ComboBox cboTipoIngreso 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            ToolTipText     =   "Tipo de ingreso"
            Top             =   2520
            Width           =   3600
         End
         Begin VB.ComboBox cboFrecuencia 
            Height          =   315
            Left            =   7500
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Tipo de frecuencia"
            Top             =   2160
            Width           =   3150
         End
         Begin VB.ListBox lstFacturas 
            Height          =   735
            ItemData        =   "frmInterfazVitamedica.frx":0038
            Left            =   10920
            List            =   "frmInterfazVitamedica.frx":003A
            Style           =   1  'Checkbox
            TabIndex        =   22
            ToolTipText     =   "Selección de facturas"
            Top             =   525
            Width           =   2355
         End
         Begin VB.TextBox txtFechaEgreso 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7500
            TabIndex        =   14
            ToolTipText     =   "Fecha de egreso del paciente"
            Top             =   630
            Width           =   1605
         End
         Begin VB.TextBox txtFechaIngreso 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7500
            TabIndex        =   13
            ToolTipText     =   "Fecha de ingreso del paciente"
            Top             =   255
            Width           =   1605
         End
         Begin VB.TextBox txtEmpresaPaciente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   5
            ToolTipText     =   "Empresa relacionada con la cuenta del paciente"
            Top             =   990
            Width           =   3600
         End
         Begin VB.TextBox txtMovimientoPaciente 
            Alignment       =   1  'Right Justify
            CausesValidation=   0   'False
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   1
            ToolTipText     =   "Número de cuenta del paciente"
            Top             =   240
            Width           =   1300
         End
         Begin VB.TextBox txtPaciente 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Nombre del paciente"
            Top             =   630
            Width           =   3600
         End
         Begin VB.OptionButton OptTipoPaciente 
            Caption         =   "Interno"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   2
            ToolTipText     =   "Paciente de tipo interno"
            Top             =   270
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptTipoPaciente 
            Caption         =   "Externo"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   3
            ToolTipText     =   "Paciente de tipo externo"
            Top             =   270
            Width           =   975
         End
         Begin VB.TextBox txtNumeroControl 
            Enabled         =   0   'False
            Height          =   315
            Left            =   7500
            Locked          =   -1  'True
            TabIndex        =   15
            ToolTipText     =   "Número de control capturado en el registro del paciente"
            Top             =   990
            Width           =   3150
         End
         Begin VB.TextBox txtPreAutorizacion 
            Height          =   315
            Left            =   7500
            MaxLength       =   9
            TabIndex        =   16
            ToolTipText     =   "Número de pre-autorización otorgado por Vitamédica para el paciente"
            Top             =   1395
            Width           =   3150
         End
         Begin VB.ComboBox cboComanda 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Tipo de comanda"
            Top             =   1395
            Width           =   3600
         End
         Begin VB.Label lblDescripcionProc 
            Caption         =   "Procedimiento"
            Height          =   195
            Left            =   6000
            TabIndex        =   70
            ToolTipText     =   "Procedimiento"
            Top             =   3300
            Width           =   1935
         End
         Begin VB.Label lblDescripcionDiagnostico 
            Caption         =   "Diagnóstico"
            Height          =   195
            Left            =   6000
            TabIndex        =   69
            ToolTipText     =   "Diagnóstico"
            Top             =   2940
            Width           =   1935
         End
         Begin VB.Label lblCodigoCPT 
            Caption         =   "Código CPT"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            ToolTipText     =   "Código CPT"
            Top             =   3300
            Width           =   1935
         End
         Begin VB.Label lblCodigoICD 
            Caption         =   "Código ICD"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "Código ICD"
            Top             =   2940
            Width           =   1935
         End
         Begin VB.Label lblClaveMedico 
            Caption         =   "Clave del médico"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            ToolTipText     =   "Clave del médico"
            Top             =   3660
            Width           =   1935
         End
         Begin VB.Label lblBeneficiario 
            Caption         =   "Número de beneficiario"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            ToolTipText     =   "Número de beneficiario"
            Top             =   2220
            Width           =   1935
         End
         Begin VB.Label lblNomina 
            Caption         =   "Número de nómina"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            ToolTipText     =   "Número de nómina"
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label lblTipoFactura 
            Caption         =   "Tipo de factura"
            Height          =   195
            Left            =   6000
            TabIndex        =   63
            ToolTipText     =   "Tipo de factura"
            Top             =   2580
            Width           =   1335
         End
         Begin VB.Label lblTipoEgreso 
            Caption         =   "Tipo de egreso"
            Height          =   195
            Left            =   6000
            TabIndex        =   62
            ToolTipText     =   "Tipo de egreso"
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label lblTipoIngreso 
            Caption         =   "Tipo de ingreso"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            ToolTipText     =   "Tipo de ingreso"
            Top             =   2580
            Width           =   1335
         End
         Begin VB.Label lblFrecuencia 
            Caption         =   "Frecuencia"
            Height          =   195
            Left            =   6000
            TabIndex        =   59
            ToolTipText     =   "Frecuencia"
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Facturas"
            Height          =   195
            Left            =   10920
            TabIndex        =   58
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Nombre"
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Número de cuenta"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "Número de cuenta"
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo de comanda"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Tipo de comanda"
            Top             =   1455
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha ingreso"
            Height          =   195
            Left            =   6000
            TabIndex        =   38
            Top             =   315
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha egreso"
            Height          =   195
            Left            =   6000
            TabIndex        =   37
            Top             =   690
            Width           =   1035
         End
         Begin VB.Label Label6 
            Caption         =   "Número de control"
            Height          =   195
            Left            =   6000
            TabIndex        =   36
            Top             =   1050
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Pre-autorización"
            Height          =   195
            Left            =   6000
            TabIndex        =   35
            Top             =   1455
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdSeleccionarTodos 
         Caption         =   "Seleccionar todo"
         Height          =   390
         Left            =   8955
         TabIndex        =   25
         ToolTipText     =   "Seleccionar todo"
         Top             =   7755
         Width           =   2340
      End
      Begin VB.CommandButton cmdInvertirSeleccion 
         Caption         =   "Seleccionar/Quitar selección"
         Height          =   390
         Left            =   6585
         TabIndex        =   24
         ToolTipText     =   "Seleccionar/Quitar selección"
         Top             =   7755
         Width           =   2340
      End
      Begin VB.CommandButton cmdQuitarSeleccion 
         Caption         =   "Quitar selección"
         Height          =   390
         Left            =   11325
         TabIndex        =   26
         ToolTipText     =   "Quitar selección"
         Top             =   7755
         Width           =   2340
      End
      Begin VB.Frame FreBotonera 
         Height          =   850
         Left            =   6165
         TabIndex        =   33
         Top             =   8445
         Width           =   2010
         Begin VB.CommandButton cmdImprimir 
            Height          =   550
            Left            =   1320
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmInterfazVitamedica.frx":003C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Imprimir"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton cmdEnviar 
            Height          =   550
            Left            =   720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmInterfazVitamedica.frx":01DE
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Enviar correo"
            Top             =   180
            Width           =   555
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   550
            Left            =   120
            Picture         =   "frmInterfazVitamedica.frx":0F04
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Consulta de los archivos de texto generados"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   555
         End
      End
   End
   Begin VB.Frame freBarra 
      Height          =   1335
      Left            =   1680
      TabIndex        =   0
      Top             =   9480
      Visible         =   0   'False
      Width           =   8205
      Begin VB.PictureBox pgbBarra 
         Height          =   360
         Left            =   165
         ScaleHeight     =   300
         ScaleWidth      =   7875
         TabIndex        =   30
         Top             =   675
         Width           =   7935
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
         Left            =   75
         TabIndex        =   31
         Top             =   135
         Width           =   7875
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
   Begin VB.Label Label15 
      Caption         =   "Frecuencia"
      Height          =   195
      Left            =   5400
      TabIndex        =   60
      Top             =   1980
      Width           =   1335
   End
End
Attribute VB_Name = "frmInterfazVitamedica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmInterfazVitamedica                                  -
'-------------------------------------------------------------------------------------
'| Objetivo: Generar un archivo de texto de la cuenta del paciente para Vitamédica
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Teresita de J. Zubía Ramos
'| Autor                    : Teresita de J. Zubía Ramos
'| Fecha de Creación        : 01/Mayo/2020
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : 15/Mayo/2020
'| Fecha última modificación:
'-------------------------------------------------------------------------------------
'| Fecha última modificación:
'-------------------------------------------------------------------------------------
'| Fecha última modificación:
'|
'-------------------------------------------------------------------------------------
'| Fecha última modificación:
'-------------------------------------------------------------------------------------


Option Explicit
Private vgrptReporte As CRAXDRT.Report

Public strCorreoDestinatario As String      'Correo del destinatario cargado automáticamente de la tabla correspondiente
Public strAsunto As String                  'Asunto del correo electrónico
Public strMensaje As String                 'Mensaje del correo electrónico
Public strRutaTXT As String                 'Ruta donde se almacenan los archivos TXT para Vitamédica
Public strNombreArchivoTXT As String        'Nombre del archivo TXT
Public vllngNumeroOpcion As Long
Public blnArchivoTXT As Boolean             'Indica si se anexará un archivo TXT

Dim Paterno As String
Dim Materno As String
Dim Nombre As String

Dim vgbitParametros As Boolean
Dim vgintEmpresa As Integer
Dim vgintTipoPaciente As Integer
Dim vgstrEstadoManto As String
Dim vllngPersonaGraba As Long
Dim llngTotalSel As Long                    'No. de cargos seleccionados
Dim vlblnFechaInicial As Boolean
Dim vlblnfechaFinal As Boolean
Dim vlblnCargos As Boolean                  'Identificar si el paciente tiene cargos

Dim vlblnFiltroFacturasTrabajando As Boolean
Dim vlblnFormatoVitamedica As Boolean

Private Sub pConfiguraGridCargos()
    On Error GoTo NotificaError

    llngTotalSel = 0
    
    grdCargos.Redraw = False
    
    With grdCargos
        .Cols = 52
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Fecha|Descripción|Cantidad|Precio|Importe|Descuento|Subtotal|IVA|Total|||||||||||||||Factura|conceptoFacturacion|paquete"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 1550 'Fecha ingreso
        .ColWidth(2) = 3360 'Descripción
        .ColWidth(3) = 800  'Cantidad
        .ColWidth(4) = 1000 'Precio
        .ColWidth(5) = 1200 'Importe
        .ColWidth(6) = 990 'Descuento
        .ColWidth(7) = 1150 'Subtotal
        .ColWidth(8) = 800 'IVA cargo
        .ColWidth(9) = 1200 'Total
        
        .ColWidth(10) = 0   'Monto facturado
        .ColWidth(11) = 0   'Importe gravado
        .ColWidth(12) = 0   'IVA
        .ColWidth(13) = 0   'Fecha sistema
        .ColWidth(14) = 0   'Hora sistema
        .ColWidth(15) = 0   'Nombre del hospital
        .ColWidth(16) = 0   'RFC
        .ColWidth(17) = 0   'Número del paciente
        .ColWidth(18) = 0   'Cuenta
        .ColWidth(19) = 0   'Tipo de paciente: interno/externo
        
        .ColWidth(20) = 0       'Nombre del paciente
        .ColWidth(21) = 0       'Fecha egreso
        .ColWidth(22) = 0       'Tipo paciente
        .ColWidth(23) = 0       'Nombre de la empresa
        .ColWidth(24) = 1000    'Facturas
        .ColWidth(25) = 0       'Paquetes
        .ColWidth(26) = 0       'Tasa IVA
        .ColWidth(27) = 0       'Importe facturado
        .ColWidth(28) = 0       'Desglose conceptos seguros
        .ColWidth(29) = 0       'Descuento especial facturado
        
        .ColWidth(30) = 0   'Descuento especial porcentaje
        .ColWidth(31) = 0   'Descuento especial límite
        .ColWidth(32) = 0   'Monto
        .ColWidth(33) = 0   'Total a pagar
        .ColWidth(34) = 0
        .ColWidth(35) = 0
        .ColWidth(36) = 0
        .ColWidth(37) = 0
        .ColWidth(38) = 0
        .ColWidth(39) = 0
        .ColWidth(40) = 0
        .ColWidth(41) = 0
        .ColWidth(42) = 0
        .ColWidth(43) = 0
        .ColWidth(44) = 0   'Para calcular el subtotal2 archivo TXT
        .ColWidth(45) = 0
        .ColWidth(46) = 0   'Para calcular el subtotal archivo TXT
        .ColWidth(47) = 0
        .ColWidth(48) = 0
        .ColWidth(49) = 0
        .ColWidth(50) = 0
        .ColWidth(51) = 0
        
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignRightBottom
        .ColAlignment(4) = flexAlignRightBottom
        .ColAlignment(5) = flexAlignRightBottom
        .ColAlignment(6) = flexAlignRightBottom
        .ColAlignment(7) = flexAlignRightBottom
        .ColAlignment(8) = flexAlignRightBottom
        .ColAlignment(9) = flexAlignRightBottom
        .ColAlignment(10) = flexAlignRightBottom
        .ColAlignment(24) = flexAlignLeftBottom
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
        .ColAlignmentFixed(24) = flexAlignCenterCenter

        .ScrollBars = flexScrollBarBoth
    End With
    
    grdCargos.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    Unload Me
End Sub


Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    On Error GoTo NotificaError
    
    Dim vlbytColumnas As Byte
    
    grdCargos.Redraw = False
    
       With ObjGrd
        ' .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        '.Clear
        For vlbytColumnas = 0 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
            .Col = vlbytColumnas
            .BackColor = &H80000005
            .ForeColor = &H80000008
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
        
        End With
        
    grdCargos.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
    Unload Me
End Sub

Private Sub pCancelar()
    On Error GoTo NotificaError
    
    FreDetalle.Enabled = False
    FrePaciente.Enabled = True
    
    pinhabilita
    txtPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    vgstrEstadoManto = ""
    cboComanda.Clear
    
    txtMovimientoPaciente.Locked = False
    OptTipoPaciente(0).Enabled = True
    OptTipoPaciente(1).Enabled = True
    cmdInvertirSeleccion.Enabled = False
    cmdSeleccionarTodos.Enabled = False
    cmdQuitarSeleccion.Enabled = False
    
    'strCorreoDestinatario = ""
    strAsunto = ""
    strMensaje = ""
    strRutaTXT = ""
    Paterno = ""
    Materno = ""
    Nombre = ""
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pcancelar"))
    Unload Me
End Sub

Private Sub cboComanda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Function fblnConsulta() As Boolean

    fblnConsulta = True
    If Trim(mskFechaInicial.ClipText) <> "" Then
        If Not IsDate(mskFechaInicial.Text) Then
            fblnConsulta = False
            vlblnFechaInicial = True
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaInicial.SetFocus
        End If
    End If
    If Not IsDate(mskFechaFinal.Text) Then vlblnfechaFinal = True
    If fblnConsulta And Trim(mskFechaFinal.ClipText) <> "" Then
        If Not IsDate(mskFechaFinal.Text) Then
            fblnConsulta = False
            vlblnfechaFinal = True
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaFinal.SetFocus
        End If
    End If
    If fblnConsulta And Trim(mskFechaInicial.ClipText) <> "" And Trim(mskFechaFinal.ClipText) <> "" Then
        If CDate(mskFechaFinal.Text) < CDate(mskFechaInicial.Text) Then
            fblnConsulta = False
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
            mskFechaInicial.SetFocus
        End If
    End If

End Function

Private Sub pCargaCboFactura()
On Error GoTo NotificaError
    Dim vlstrFactura As String
    Dim rsFactura As New ADODB.Recordset
    
    vlstrFactura = "select distinct pvcargo.chrfoliofactura, nvl(pvfactura.intConsecutivo,0) intConsecutivo from pvcargo left join pvfactura on trim(pvfactura.chrfoliofactura) = trim(pvcargo.chrfoliofactura) where pvcargo.bitexcluido = 0 and pvcargo.INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " order by nvl(pvfactura.intConsecutivo,0), trim(pvcargo.chrfoliofactura)"
    
    Set rsFactura = frsRegresaRs(vlstrFactura, adLockReadOnly, adOpenForwardOnly)
        
    lstFacturas.Clear
    
    If rsFactura.RecordCount > 0 Then
        Do While Not rsFactura.EOF
            If rsFactura!intConsecutivo = 0 Then
                lstFacturas.AddItem "<SIN FACTURA>"
                lstFacturas.ItemData(lstFacturas.newIndex) = 0
                lstFacturas.Selected(lstFacturas.newIndex) = True
            Else
                lstFacturas.AddItem Trim(rsFactura!chrfoliofactura)
                lstFacturas.ItemData(lstFacturas.newIndex) = rsFactura!intConsecutivo
                lstFacturas.Selected(lstFacturas.newIndex) = True
            End If
            rsFactura.MoveNext
        Loop
    End If
    rsFactura.Close
    
    vlblnFiltroFacturasTrabajando = True
    
    lstFacturas.ListIndex = 0
    
Exit Sub
NotificaError:
     Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaCboFactura"))
End Sub

Private Sub cboFrecuencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboTipoEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboTipoFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cboTipoIngreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError

    SSTVitamedica.Tab = 1
    pCancelarConsultaTXT
    mskFechaInicial.Text = fdtmServerFecha
    mskFechaFinal.Text = fdtmServerFecha
    mskFechaInicial.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdLocate_Click"))
End Sub

Private Sub cmdCargar_Click()
    Dim rs As New ADODB.Recordset
    Dim strFiltroFecha As String
    Dim strFechaInicio As String
    Dim strFechafinal As String
    Dim Y As Integer
    
    If Trim(txtCuentaPaciente.Text) = "" Then
        'Debe indicar un número de cuenta.
        MsgBox SIHOMsg(1579), vbExclamation, "Mensaje"
        txtCuentaPaciente.SetFocus
        Exit Sub
    End If
    
    If Not fblnConsulta Then Exit Sub
    
    FreDetalleTXT.Enabled = True
    FreDetalleTXT.Visible = True
    pConfiguraGridTXT
    grdTXT.Cols = 10
    pLimpiaGrid grdTXT
    
    strFiltroFecha = IIf(IsDate(mskFechaInicial.Text), "1", "0")
                                    
    If IsDate(mskFechaInicial.Text) And IsDate(mskFechaFinal.Text) Then
        strFechaInicio = fstrFechaSQL(mskFechaInicial.Text)
        strFechafinal = fstrFechaSQL(mskFechaFinal.Text)
    Else
        strFechaInicio = fstrFechaSQL(fdtmServerFecha)
        strFechafinal = fstrFechaSQL(fdtmServerFecha)
        mskFechaInicial.Text = fdtmServerFecha
        mskFechaFinal.Text = fdtmServerFecha
    End If
    
 
    vgstrParametrosSP = _
    txtCuentaPaciente.Text & _
    "|" & IIf(optPaciente(2).Value, "I", "E") & _
    "|" & strFechaInicio & _
    "|" & strFechafinal
    
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelTxtVitamedica")
    
    If rs.RecordCount <> 0 Then
       
        With grdTXT
            Do While Not rs.EOF
                .TextMatrix(.Rows - 1, 1) = Format(rs!dtmfechatxt, "dd/mmm/yyyy hh:mm")
                .TextMatrix(.Rows - 1, 2) = rs!cuenta
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(txtNombrePaciente), " ", txtNombrePaciente)
                .TextMatrix(.Rows - 1, 4) = IIf(IsNull(rs!empresa), " ", rs!empresa)
                .TextMatrix(.Rows - 1, 5) = IIf(IsNull(rs!chrcomanda), " ", rs!chrcomanda)
                .TextMatrix(.Rows - 1, 6) = IIf(IsNull(rs!Empleado), " ", rs!Empleado)

                .Rows = .Rows + 1
                rs.MoveNext
            Loop
            .Rows = .Rows - 1
        End With
        
    Else
        'No existe información con esos parámetros.
        MsgBox SIHOMsg(236), vbInformation + vbOKOnly, "Mensaje"
        vgstrEstadoManto = ""
        pCancelarConsultaTXT
        mskFechaInicial.Text = fdtmServerFecha
        mskFechaFinal.Text = fdtmServerFecha
        mskFechaInicial.SetFocus
    
    End If

End Sub

Private Sub cmdEnviar_Click()
    pExportaTXT
End Sub

Private Sub cmdImprimir_Click()
    Dim alstrParametros(110) As String
    Dim alstrParametros2(100) As String
    Dim rsDatosPaciente As New ADODB.Recordset
    Dim rsEstadoCuenta As New ADODB.Recordset
    Dim lintCveEmpresaPaciente As Integer
    Dim vlstrDestino As String
    Dim vlstrSentencia As String
    Dim rsEspecialidad As New ADODB.Recordset
    Dim rsDiagnostico As New ADODB.Recordset
    Dim rsProcedimiento As New ADODB.Recordset
    Dim dbleTotalConIVA As Double
    Dim dbleTotalSinIVA As Double
    Dim vlsumatotal As Double
    Dim vlLngCont As Long
    Dim dia As String
    Dim mes As String
    Dim año As String
    Dim Hora As String
    Dim minutos As String
    Dim horaMinutos As String
    Dim vlstrTipoIngreso As String
    Dim vlstrDiagnostico As String
    Dim vlblnbandera As Boolean 'Nos ayudara a saber si entro o no en un concepto  y saber si desglosa
    'Variables de conceptos sumatorias
    'Paquetes
    Dim dblTotalPaquete As Double
    Dim dblTotalPaqueteCorta As Double
    Dim dblTotalPaqueteOtros As Double
    'Estancia
    Dim dblTotalCuartoPrivado As Double
    Dim dblTotalCuna As Double
    Dim dblTotalCrecimiento As Double
    Dim dblTotalTerapiaNeonatal As Double
    Dim dblTotalTerapiaPediatrica As Double
    Dim dblTotalTerapiaAdulto As Double
    Dim dblTotalTerapiaIntermediaNeonatal As Double
    Dim dblTotalTerapiaIntermediaPediatrica As Double
    Dim dblTotalTerapiaIntermediaAdulto As Double
    Dim dblTotalCoronaria As Double
    'Salas
    Dim dblTotalSalaCirugiaMenorPrimera As Double
    Dim dblTotalSalaCirugiaMenorFraccion As Double
    Dim dblTotalSalaCortaEstancia As Double
    Dim dblTotalSalaQuirurgica As Double
    Dim dblTotalSalaQuirurgicaFraccion As Double
    Dim dblTotalSalaRecuperacion As Double
    Dim dblTotalSalaExpulsionHora As Double
    Dim dblTotalSalaExpulsionFraccion As Double
    Dim dblTotalSalaLabor As Double
    Dim dblTotalOtrosConceptosSalas As Double
    'Medicamentos
    Dim dblTotalMedicamentoGravable As Double
    Dim dblTotalMedicamentoNoGravable As Double
    'OtrosConceptos
    Dim dblTotalOtrosEstancia As Double
    Dim dblTotalOtrosTerapiaIntensiva As Double
    Dim dblTotalOtrosTerapiaIntermedia As Double
    Dim dblTotalOtrosCuna As Double
    Dim dblTotalMaterialCuracion As Double
    Dim dblTotalMaterialQuirofanos As Double
    Dim dblTotalEquiposUtilizados As Double
    Dim dblTotalEquiposDesechables As Double
    Dim dblTotalLaboratorioClinico As Double
    Dim dblTotalLaboratorioPAtologia As Double
    Dim dblTotalRadiologiaDiagnostica As Double
    Dim dblTotalRadiologiaTerapeutica As Double
    Dim dblTotalResonanciaMAgnetica As Double
    Dim dblTotalMedicinaNuclear As Double
    Dim dblTotalTomografia As Double
    Dim dblTotalUltrasonido As Double
    Dim dblTotalInhaloterapia As Double
    Dim dblTotalEquipoAnestesia As Double
    Dim dblTotalInstrumentalMicrocirugia As Double
    Dim dblTotalOtros1 As Double
    Dim dblTotalOtros2 As Double
    Dim dblTotalOtros3 As Double
    Dim dblTotalOtros4 As Double
    Dim dblTotalOtros5 As Double
    Dim strOtros1 As String
    Dim strOtros2 As String
    Dim strOtros3 As String
    Dim strOtros4 As String
    Dim strOtros5 As String
    
    'Paquetes
    Dim intTotalPaqueteUNIDAD As Double
    Dim intTotalPaqueteCortaUNIDAD As Double
    Dim intTotalPaqueteOtrosUNIDAD As Double
    'Estancia
    Dim intTotalCuartoPrivadoUNIDAD As Double
    Dim intTotalCunaUNIDAD As Double
    Dim intTotalCrecimientoUNIDAD As Double
    Dim intTotalTerapiaNeonatalUNIDAD As Double
    Dim intTotalTerapiaPediatricaUNIDAD As Double
    Dim intTotalTerapiaAdultoUNIDAD As Double
    Dim intTotalTerapiaIntermediaNeonatalUNIDAD As Double
    Dim intTotalTerapiaIntermediaPediatricaUNIDAD As Double
    Dim intTotalTerapiaIntermediaAdultoUNIDAD As Double
    Dim intTotalCoronariaUNIDAD As Double
    'Salas
    Dim intTotalSalaCirugiaMenorPrimeraUNIDAD As Double
    Dim intTotalSalaCirugiaMenorFraccionUNIDAD As Double
    Dim intTotalSalaCortaEstanciaUNIDAD As Double
    Dim intTotalSalaQuirurgicaUNIDAD As Double
    Dim intTotalSalaQuirurgicaFraccionUNIDAD As Double
    Dim intTotalSalaRecuperacionUNIDAD As Double
    Dim intTotalSalaExpulsionHoraUNIDAD As Double
    Dim intTotalSalaExpulsionFraccionUNIDAD As Double
    Dim intTotalSalaLaborUNIDAD As Double
    Dim intTotalOtrosConceptosSalasUNIDAD As Double
    'Medicamentos
    Dim intTotalMedicamentoGravableUNIDAD As Double
    Dim intTotalMedicamentoNoGravableUNIDAD As Double
    'OtrosConceptos
    Dim intTotalOtrosEstanciaUNIDAD As Double
    Dim intTotalOtrosTerapiaIntensivaUNIDAD As Double
    Dim intTotalOtrosTerapiaIntermediaUNIDAD As Double
    Dim intTotalOtrosCunaUNIDAD As Double
    Dim intTotalMaterialCuracionUNIDAD As Double
    Dim intTotalMaterialQuirofanosUNIDAD As Double
    Dim intTotalEquiposUtilizadosUNIDAD As Double
    Dim intTotalEquiposDesechablesUNIDAD As Double
    Dim intTotalLaboratorioClinicoUNIDAD As Double
    Dim intTotalLaboratorioPAtologiaUNIDAD As Double
    Dim intTotalRadiologiaDiagnosticaUNIDAD As Double
    Dim intTotalRadiologiaTerapeuticaUNIDAD As Double
    Dim intTotalResonanciaMAgneticaUNIDAD As Double
    Dim intTotalMedicinaNuclearUNIDAD As Double
    Dim intTotalTomografiaUNIDAD As Double
    Dim intTotalUltrasonidoUNIDAD As Double
    Dim intTotalInhaloterapiaUNIDAD As Double
    Dim intTotalEquipoAnestesiaUNIDAD As Double
    Dim intTotalInstrumentalMicrocirugia As Double
    Dim intTotalOtros1 As Double
    Dim intTotalOtros2 As Double
    Dim intTotalOtros3 As Double
    Dim intTotalOtros4 As Double
    Dim intTotalOtros5 As Double
    
    Dim vlstrCodigo As String
    Dim vlstrProcedimiento As String
    Dim vlstrCodigoProcedimiento As String
    Dim vlstrEspecialidad As String
    Dim vlblnSiGeneraraInfo As Boolean
    vlblnSiGeneraraInfo = False
    
    'Inicialización de variables a 0
    intTotalPaqueteUNIDAD = 0
    intTotalPaqueteCortaUNIDAD = 0
    intTotalPaqueteOtrosUNIDAD = 0
    intTotalCuartoPrivadoUNIDAD = 0
    intTotalCunaUNIDAD = 0
    intTotalCrecimientoUNIDAD = 0
    intTotalTerapiaNeonatalUNIDAD = 0
    intTotalTerapiaPediatricaUNIDAD = 0
    intTotalTerapiaAdultoUNIDAD = 0
    intTotalTerapiaIntermediaNeonatalUNIDAD = 0
    intTotalTerapiaIntermediaPediatricaUNIDAD = 0
    intTotalTerapiaIntermediaAdultoUNIDAD = 0
    intTotalCoronariaUNIDAD = 0
    intTotalSalaCirugiaMenorPrimeraUNIDAD = 0
    intTotalSalaCirugiaMenorFraccionUNIDAD = 0
    intTotalSalaCortaEstanciaUNIDAD = 0
    intTotalSalaQuirurgicaUNIDAD = 0
    intTotalSalaQuirurgicaFraccionUNIDAD = 0
    intTotalSalaRecuperacionUNIDAD = 0
    intTotalSalaExpulsionHoraUNIDAD = 0
    intTotalSalaExpulsionFraccionUNIDAD = 0
    intTotalSalaLaborUNIDAD = 0
    intTotalOtrosConceptosSalasUNIDAD = 0
    intTotalMedicamentoGravableUNIDAD = 0
    intTotalMedicamentoNoGravableUNIDAD = 0
    intTotalOtrosEstanciaUNIDAD = 0
    intTotalOtrosTerapiaIntensivaUNIDAD = 0
    intTotalOtrosTerapiaIntermediaUNIDAD = 0
    intTotalOtrosCunaUNIDAD = 0
    intTotalMaterialCuracionUNIDAD = 0
    intTotalMaterialQuirofanosUNIDAD = 0
    intTotalEquiposUtilizadosUNIDAD = 0
    intTotalEquiposDesechablesUNIDAD = 0
    intTotalLaboratorioClinicoUNIDAD = 0
    intTotalLaboratorioPAtologiaUNIDAD = 0
    intTotalRadiologiaDiagnosticaUNIDAD = 0
    intTotalRadiologiaTerapeuticaUNIDAD = 0
    intTotalResonanciaMAgneticaUNIDAD = 0
    intTotalMedicinaNuclearUNIDAD = 0
    intTotalTomografiaUNIDAD = 0
    intTotalUltrasonidoUNIDAD = 0
    intTotalInhaloterapiaUNIDAD = 0
    intTotalEquipoAnestesiaUNIDAD = 0
    intTotalOtros1 = 0
    intTotalOtros2 = 0
    intTotalOtros3 = 0
    intTotalOtros4 = 0
    intTotalOtros5 = 0
    dblTotalPaquete = 0
    dblTotalPaqueteCorta = 0
    dblTotalPaqueteOtros = 0
    dblTotalCuartoPrivado = 0
    dblTotalCuna = 0
    dblTotalCrecimiento = 0
    dblTotalTerapiaNeonatal = 0
    dblTotalTerapiaPediatrica = 0
    dblTotalTerapiaAdulto = 0
    dblTotalTerapiaIntermediaNeonatal = 0
    dblTotalTerapiaIntermediaPediatrica = 0
    dblTotalTerapiaIntermediaAdulto = 0
    dblTotalCoronaria = 0
    dblTotalSalaCirugiaMenorPrimera = 0
    dblTotalSalaCirugiaMenorFraccion = 0
    dblTotalSalaCortaEstancia = 0
    dblTotalSalaQuirurgica = 0
    dblTotalSalaQuirurgicaFraccion = 0
    dblTotalSalaRecuperacion = 0
    dblTotalSalaExpulsionHora = 0
    dblTotalSalaExpulsionFraccion = 0
    dblTotalSalaLabor = 0
    dblTotalOtrosConceptosSalas = 0
    dblTotalMedicamentoGravable = 0
    dblTotalMedicamentoNoGravable = 0
    dblTotalOtrosEstancia = 0
    dblTotalOtrosTerapiaIntensiva = 0
    dblTotalOtrosTerapiaIntermedia = 0
    dblTotalOtrosCuna = 0
    dblTotalMaterialCuracion = 0
    dblTotalMaterialQuirofanos = 0
    dblTotalEquiposUtilizados = 0
    dblTotalEquiposDesechables = 0
    dblTotalLaboratorioClinico = 0
    dblTotalLaboratorioPAtologia = 0
    dblTotalRadiologiaDiagnostica = 0
    dblTotalRadiologiaTerapeutica = 0
    dblTotalResonanciaMAgnetica = 0
    dblTotalMedicinaNuclear = 0
    dblTotalTomografia = 0
    dblTotalUltrasonido = 0
    dblTotalInhaloterapia = 0
    dblTotalEquipoAnestesia = 0
    dblTotalInstrumentalMicrocirugia = 0
    intTotalInstrumentalMicrocirugia = 0
    dblTotalOtros1 = 0
    dblTotalOtros2 = 0
    dblTotalOtros3 = 0
    dblTotalOtros4 = 0
    dblTotalOtros5 = 0
    vlsumatotal = 0
    dbleTotalConIVA = 0
    dbleTotalSinIVA = 0
    vlstrDiagnostico = ""
    vlstrCodigo = 0
    vlstrProcedimiento = ""
    vlstrCodigoProcedimiento = ""
    vlstrEspecialidad = ""
    strOtros1 = ""
    strOtros2 = ""
    strOtros3 = ""
    strOtros4 = ""
    strOtros5 = ""
    
    
    If grdCargos.Rows = 1 Then
        'No se han seleccionado cargos para enviar.
        MsgBox "No se han encontrado cargos para imprimir", vbOKOnly + vbInformation, "Mensaje"
        Exit Sub
    End If
    
    
    
    For vlLngCont = 1 To grdCargos.Rows - 1

                
        dbleTotalSinIVA = dbleTotalSinIVA + IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 44))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 44)))
        If grdCargos.TextMatrix(vlLngCont, 42) <> "" Then
            dbleTotalConIVA = dbleTotalConIVA + CDbl(grdCargos.TextMatrix(vlLngCont, 42))
        Else
            dbleTotalConIVA = dbleTotalConIVA + 0
        End If
        
        vlsumatotal = dbleTotalSinIVA + dbleTotalConIVA
           
        Next vlLngCont
    
    'vlsumatotal = FormatNumber(vlsumatotal, 4)
    'dbleTotalConIVA = FormatNumber(dbleTotalConIVA, 4)
    'dbleTotalSinIVA = FormatNumber(dbleTotalSinIVA, 4)
    vlstrDestino = "P"
    
    
    pInstanciaReporte vgrptReporte, "rptVidamedica1.rpt"
    
    Set rsDatosPaciente = frsEjecuta_SP(Trim(txtMovimientoPaciente.Text) & "|0|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable, "SP_PVSELDATOSPACIENTEVITA")
    If rsDatosPaciente.RecordCount <> 0 Then
               
        vlstrDiagnostico = Trim(txtDescripcionDiagnostico.Text)
        vlstrProcedimiento = Trim(txtDescripcionProc.Text)
        
        vlstrTipoIngreso = rsDatosPaciente!TipoIngreso
        vlstrSentencia = "select vchdescripcion from HOESPECIALIDADMEDICO " & _
                          " inner join HOESPECIALIDAD ON HOESPECIALIDADMEDICO.TNYCVEESPECIALIDAD = HOESPECIALIDAD.TNYCVEESPECIALIDAD " & _
                          " where HOESPECIALIDADMEDICO.INTCVEMEDICO = " & rsDatosPaciente!ClaveMedico & " and rownum=1 order by HOESPECIALIDAD.TNYCVEESPECIALIDAD asc "
        Set rsEspecialidad = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        'Especialidad medico
        If rsEspecialidad.RecordCount > 0 Then
            vlstrEspecialidad = rsEspecialidad!VCHDESCRIPCION
            alstrParametros(0) = "Especialidad;" & vlstrEspecialidad
        Else
            alstrParametros(0) = "Especialidad;" & ""
        End If
        'Encabezado
        alstrParametros(1) = "numElegibilidad;" & Trim(txtNumeroControl.Text)
        alstrParametros(2) = "numPreAutorizacion;" & Trim(txtPreAutorizacion.Text)
        
        'Datos del paciente
        alstrParametros(3) = "Cliente;" & Trim(rsDatosPaciente!empresa)
        alstrParametros(4) = "NumNomina;" & Trim(txtNomina.Text)
        alstrParametros(5) = "NumBeneficiario;" & Trim(txtBeneficiario.Text)
        alstrParametros(6) = "ApellidoPaterno;" & Trim(rsDatosPaciente!Paterno)
        alstrParametros(7) = "ApellidoMaterno;" & Trim(rsDatosPaciente!Materno)
        alstrParametros(8) = "Nombres;" & Trim(rsDatosPaciente!Nombre)
        alstrParametros(9) = "FechaIngresoDia;"
        alstrParametros(10) = "FechaIngresoMes;"
        alstrParametros(11) = "FechaIngresoAño;"
        alstrParametros(12) = "HoraIngreso;"
        alstrParametros(13) = "MinutosIngreso;"
                      
        If Not IsNull(rsDatosPaciente!Ingreso) Then
            alstrParametros(9) = "FechaIngresoDia;" & Format(rsDatosPaciente!Ingreso, "dd")
            alstrParametros(10) = "FechaIngresoMes;" & Format(rsDatosPaciente!Ingreso, "mm")
            alstrParametros(11) = "FechaIngresoAño;" & Format(rsDatosPaciente!Ingreso, "yyyy")
            alstrParametros(12) = "HoraIngreso;" & Format(rsDatosPaciente!Ingreso, "hh")
            alstrParametros(13) = "MinutosIngreso;" & Format(rsDatosPaciente!Ingreso, "mm")
        End If
        alstrParametros(14) = "FechaEgresoDia;"
        alstrParametros(15) = "FechaEgresoMes;"
        alstrParametros(16) = "FechaEgresoAño;"
        alstrParametros(17) = "HoraEgreso;"
        alstrParametros(18) = "MinutosEgreso;"
        
        If Not IsNull(rsDatosPaciente!Egreso) Then
            alstrParametros(14) = "FechaEgresoDia;" & Format(rsDatosPaciente!Egreso, "dd")
            alstrParametros(15) = "FechaEgresoMes;" & Format(rsDatosPaciente!Egreso, "mm")
            alstrParametros(16) = "FechaEgresoAño;" & Format(rsDatosPaciente!Egreso, "yyyy")
            alstrParametros(17) = "HoraEgreso;" & Format(rsDatosPaciente!Egreso, "hh")
            alstrParametros(18) = "MinutosEgreso;" & Format(rsDatosPaciente!Egreso, "mm")
        End If
        'Medico tratante
        alstrParametros(19) = "ClaveMedico;" & Trim(txtClaveMedico.Text)
        alstrParametros(20) = "NombreMedico;" & Trim(rsDatosPaciente!Medico)
        'Datos de ingreso
        alstrParametros(21) = "TipoIngreso;" & cboTipoIngreso.ListIndex
        alstrParametros(22) = "CodigoICD;" & Trim(txtCodigoICD.Text)
        alstrParametros(23) = "DescripcionDiagnostico;" & Trim(vlstrDiagnostico)
        alstrParametros(24) = "CodigoCPT;" & Trim(txtCodigoCPT.Text)
        alstrParametros(25) = "DescripcionProcedimiento;" & Trim(vlstrProcedimiento)
        'Datos de egreso
        alstrParametros(26) = "MotivoEgreso;" & cboTipoEgreso.ListIndex
        'Datos de facturación
        alstrParametros(27) = "ClaveHospital;" & "7595"
        alstrParametros(28) = "NombreHospital;" & Trim("SANATORIO OFTALMOLÓGICO MÉRIDA")
        alstrParametros(29) = "TipoFactura;" & cboTipoFactura.ListIndex
        alstrParametros(30) = "Frecuencia;" & cboFrecuencia.ListIndex
        alstrParametros(31) = "TotalSinIVa;" & dbleTotalSinIVA
        alstrParametros(32) = "TotalConIVA;" & vlsumatotal
    End If

    Dim dblSinIVA As Double
    
    For vlLngCont = 1 To grdCargos.Rows - 1
        vlblnbandera = False
        
        dblSinIVA = IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 44))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 44)))
                
        If grdCargos.TextMatrix(vlLngCont, 50) = "PAQUETES" Then
            dblTotalPaquete = dblTotalPaquete + dblSinIVA
            intTotalPaqueteUNIDAD = 1 'intTotalPaqueteUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "PAQUETE CORTA ESTANCIA" Then
            dblTotalPaqueteCorta = dblTotalPaqueteCorta + dblSinIVA
            intTotalPaqueteCortaUNIDAD = 1 'intTotalPaqueteCortaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS PAQUETES" Then
            dblTotalPaqueteOtros = dblTotalPaqueteOtros + dblSinIVA
            intTotalPaqueteOtrosUNIDAD = 1 'intTotalPaqueteOtrosUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "CUARTO PRIVADO" Then
            dblTotalCuartoPrivado = dblTotalCuartoPrivado + dblSinIVA
            intTotalCuartoPrivadoUNIDAD = 1 'intTotalCuartoPrivadoUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "CUNA" Then
            dblTotalCuna = dblTotalCuna + dblSinIVA
            intTotalCunaUNIDAD = 1 'intTotalCunaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "CRECIMIENTO Y DESARROLLO" Then
            dblTotalCrecimiento = dblTotalCrecimiento + dblSinIVA
            intTotalCrecimientoUNIDAD = 1 'intTotalCrecimientoUNIDAD + 1
            vlblnbandera = True
            
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTENSIVA NEONATAL" Then
            dblTotalTerapiaNeonatal = dblTotalTerapiaNeonatal + dblSinIVA
            intTotalTerapiaNeonatalUNIDAD = 1 'intTotalTerapiaNeonatalUNIDAD + 1
            vlblnbandera = True
        End If
        'aaaaaa
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTENSIVA PEDIÁTRICA" Then
            dblTotalTerapiaPediatrica = dblTotalTerapiaPediatrica + dblSinIVA
            intTotalTerapiaPediatricaUNIDAD = 1 'intTotalTerapiaPediatricaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTENSIVA ADULTO" Then
            dblTotalTerapiaAdulto = dblTotalTerapiaAdulto + dblSinIVA
            intTotalTerapiaAdultoUNIDAD = 1 'intTotalTerapiaAdultoUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTERMEDIA NEONATAL" Then
            dblTotalTerapiaIntermediaNeonatal = dblTotalTerapiaIntermediaNeonatal + dblSinIVA
            intTotalTerapiaIntermediaNeonatalUNIDAD = 1 'intTotalTerapiaIntermediaNeonatalUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTERMEDIA PEDIÁTRICA" Then
            dblTotalTerapiaIntermediaPediatrica = dblTotalTerapiaIntermediaPediatrica + dblSinIVA
            intTotalTerapiaIntermediaPediatricaUNIDAD = 1 'intTotalTerapiaIntermediaPediatricaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TERAPIA INTERMEDIA ADULTO" Then
            dblTotalTerapiaIntermediaAdulto = dblTotalTerapiaIntermediaAdulto + dblSinIVA
            intTotalTerapiaIntermediaAdultoUNIDAD = 1 'intTotalTerapiaIntermediaAdultoUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "UNIDAD CORONARIA" Then
            dblTotalCoronaria = dblTotalCoronaria + dblSinIVA
            intTotalCoronariaUNIDAD = 1 'intTotalCoronariaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE CIRUGÍA MENOR PRIMERA HORA" Then
            dblTotalSalaCirugiaMenorPrimera = dblTotalSalaCirugiaMenorPrimera + dblSinIVA
            intTotalSalaCirugiaMenorPrimeraUNIDAD = 1 'intTotalSalaCirugiaMenorPrimeraUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE CIRUGÍA MENOR FRACCIÓN SUBSECUENTE" Then
            dblTotalSalaCirugiaMenorFraccion = dblTotalSalaCirugiaMenorFraccion + dblSinIVA
            intTotalSalaCirugiaMenorFraccionUNIDAD = 1 'intTotalSalaCirugiaMenorFraccionUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE CORTA ESTANCIA" Then
            dblTotalSalaCortaEstancia = dblTotalSalaCortaEstancia + dblSinIVA
            intTotalSalaCortaEstanciaUNIDAD = 1 'intTotalSalaCortaEstanciaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA QUIRÚRGICA 1 HORA" Then
            dblTotalSalaQuirurgica = dblTotalSalaQuirurgica + dblSinIVA
            intTotalSalaQuirurgicaUNIDAD = 1 'intTotalSalaQuirurgicaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA QUIRÚRGICA (FRACCIÓN ADICIONAL)" Then
            dblTotalSalaQuirurgicaFraccion = dblTotalSalaQuirurgicaFraccion + dblSinIVA
            intTotalSalaQuirurgicaFraccionUNIDAD = 1 'intTotalSalaQuirurgicaFraccionUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE RECUPERACIÓN" Then
            dblTotalSalaRecuperacion = dblTotalSalaRecuperacion + dblSinIVA
            intTotalSalaRecuperacionUNIDAD = 1 'intTotalSalaRecuperacionUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE EXPULSIÓN 1 HORA" Then
            dblTotalSalaExpulsionHora = dblTotalSalaExpulsionHora + dblSinIVA
            intTotalSalaExpulsionHoraUNIDAD = 1 'intTotalSalaExpulsionHoraUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE EXPULSIÓN (FRACCIÓN ADICIONAL)" Then
            dblTotalSalaExpulsionFraccion = dblTotalSalaExpulsionFraccion + dblSinIVA
            intTotalSalaExpulsionFraccionUNIDAD = 1 'intTotalSalaExpulsionFraccionUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "SALA DE LABOR" Then
            dblTotalSalaLabor = dblTotalSalaLabor + dblSinIVA
            intTotalSalaLaborUNIDAD = 1 'intTotalSalaLaborUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS CONCEPTOS SALAS" Then
            dblTotalOtrosConceptosSalas = dblTotalOtrosConceptosSalas + dblSinIVA
            intTotalOtrosConceptosSalasUNIDAD = 1 'intTotalOtrosConceptosSalasUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "MEDICAMENTOS GRAVABLES DE IMPUESTO" Then
            dblTotalMedicamentoGravable = dblTotalMedicamentoGravable + dblSinIVA
            intTotalMedicamentoGravableUNIDAD = 1 'intTotalMedicamentoGravableUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "MEDICAMENTOS NO GRAVABLES DE IMPUESTO" Then
            dblTotalMedicamentoNoGravable = dblTotalMedicamentoNoGravable + dblSinIVA
            intTotalMedicamentoNoGravableUNIDAD = 1 'intTotalMedicamentoNoGravableUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS ESTANCIA" Then
            dblTotalOtrosEstancia = dblTotalOtrosEstancia + dblSinIVA
            intTotalOtrosEstanciaUNIDAD = 1 'intTotalOtrosEstanciaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS TERAPIA INTENSIVA" Then
            dblTotalOtrosTerapiaIntensiva = dblTotalOtrosTerapiaIntensiva + dblSinIVA
            intTotalOtrosTerapiaIntensivaUNIDAD = 1 'intTotalOtrosTerapiaIntensivaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS TERAPIA INTERMEDIA" Then
            dblTotalOtrosTerapiaIntermedia = dblTotalOtrosTerapiaIntermedia + dblSinIVA
            intTotalOtrosTerapiaIntermediaUNIDAD = 1 'intTotalOtrosTerapiaIntermediaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "OTROS CUNA" Then
            dblTotalOtrosCuna = dblTotalOtrosCuna + dblSinIVA
            intTotalOtrosCunaUNIDAD = 1 'intTotalOtrosCunaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "MATERIAL DE CURACIÓN" Then
            dblTotalMaterialCuracion = dblTotalMaterialCuracion + dblSinIVA
            intTotalMaterialCuracionUNIDAD = 1 'intTotalMaterialCuracionUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "MATERIAL QUIRÓFANOS" Then
            dblTotalMaterialQuirofanos = dblTotalMaterialQuirofanos + dblSinIVA
            intTotalMaterialQuirofanosUNIDAD = 1 'intTotalMaterialQuirofanosUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "EQUIPOS UTILIZADOS" Then
            dblTotalEquiposUtilizados = dblTotalEquiposUtilizados + dblSinIVA
            intTotalEquiposUtilizadosUNIDAD = 1 'intTotalEquiposUtilizadosUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "EQUIPOS DESECHABLES" Then
            dblTotalEquiposDesechables = dblTotalEquiposDesechables + dblSinIVA
            intTotalEquiposDesechablesUNIDAD = 1 'intTotalEquiposDesechablesUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "LABORATORIO CLÍNICO" Then
            dblTotalLaboratorioClinico = dblTotalLaboratorioClinico + dblSinIVA
            intTotalLaboratorioClinicoUNIDAD = 1 'intTotalLaboratorioClinicoUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "LABORATORIO DE PATOLOGÍA" Then
            dblTotalLaboratorioPAtologia = dblTotalLaboratorioPAtologia + dblSinIVA
            intTotalLaboratorioPAtologiaUNIDAD = 1 'intTotalLaboratorioPAtologiaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "RADIOLOGÍA DIAGNÓSTICA" Then
            dblTotalRadiologiaDiagnostica = dblTotalRadiologiaDiagnostica + dblSinIVA
            intTotalRadiologiaDiagnosticaUNIDAD = 1 'intTotalRadiologiaDiagnosticaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "RADIOLOGÍA TERAPÉUTICA" Then
            dblTotalRadiologiaTerapeutica = dblTotalRadiologiaTerapeutica + dblSinIVA
            intTotalRadiologiaTerapeuticaUNIDAD = 1 'intTotalRadiologiaTerapeuticaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "RESONANCIA MAGNÉTICA" Then
            dblTotalResonanciaMAgnetica = dblTotalResonanciaMAgnetica + dblSinIVA
            intTotalResonanciaMAgneticaUNIDAD = 1 'intTotalResonanciaMAgneticaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "MEDICINA NUCLEAR" Then
            dblTotalMedicinaNuclear = dblTotalMedicinaNuclear + dblSinIVA
            intTotalMedicinaNuclearUNIDAD = 1 'intTotalMedicinaNuclearUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "TOMOGRAFÍA" Then
            dblTotalTomografia = dblTotalTomografia + dblSinIVA
            intTotalTomografiaUNIDAD = 1 'intTotalTomografiaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "ULTRASONIDO" Then
            dblTotalUltrasonido = dblTotalUltrasonido + dblSinIVA
            intTotalUltrasonidoUNIDAD = 1 'intTotalUltrasonidoUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "INHALOTERAPIA" Then
            dblTotalInhaloterapia = dblTotalInhaloterapia + dblSinIVA
            intTotalInhaloterapiaUNIDAD = 1 'intTotalInhaloterapiaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "EQUIPO DE ANESTESIA" Then
            dblTotalEquipoAnestesia = dblTotalEquipoAnestesia + dblSinIVA
            intTotalEquipoAnestesiaUNIDAD = 1 'intTotalEquipoAnestesiaUNIDAD + 1
            vlblnbandera = True
        End If
        
        If grdCargos.TextMatrix(vlLngCont, 50) = "INSTRUMENTAL DE MICROCIRUGÍA" Then
            dblTotalInstrumentalMicrocirugia = dblTotalInstrumentalMicrocirugia + dblSinIVA
            intTotalInstrumentalMicrocirugia = 1 'intTotalInstrumentalMicrocirugia + 1
            vlblnbandera = True
        End If
        
        If vlblnbandera = False Then
            If strOtros1 = "" Or strOtros1 = grdCargos.TextMatrix(vlLngCont, 50) Then
                If grdCargos.TextMatrix(vlLngCont, 51) Then
                    dblTotalOtros1 = dblTotalOtros1 + dblSinIVA
                    intTotalOtros1 = 1 'intTotalOtros1 + 1
                    strOtros1 = grdCargos.TextMatrix(vlLngCont, 2)
                    vlblnbandera = True
                Else
                    dblTotalOtros1 = dblTotalOtros1 + dblSinIVA
                    intTotalOtros1 = 1 'intTotalOtros1 + 1
                    strOtros1 = grdCargos.TextMatrix(vlLngCont, 50)
                    vlblnbandera = True
                End If
            End If
            If vlblnbandera = False Then
                If strOtros2 = "" Or strOtros2 = grdCargos.TextMatrix(vlLngCont, 50) Then
                    If grdCargos.TextMatrix(vlLngCont, 51) Then
                        dblTotalOtros2 = dblTotalOtros2 + dblSinIVA
                        intTotalOtros2 = 1 'intTotalOtros2 + 1
                        strOtros2 = grdCargos.TextMatrix(vlLngCont, 2)
                        vlblnbandera = True
                    Else
                        dblTotalOtros2 = dblTotalOtros2 + dblSinIVA
                        intTotalOtros2 = 1 'intTotalOtros2 + 1
                        strOtros2 = grdCargos.TextMatrix(vlLngCont, 50)
                        vlblnbandera = True
                    End If
                End If
            End If
            If vlblnbandera = False Then
                If strOtros3 = "" Or strOtros3 = grdCargos.TextMatrix(vlLngCont, 50) Then
                    If grdCargos.TextMatrix(vlLngCont, 51) Then
                        dblTotalOtros3 = dblTotalOtros3 + dblSinIVA
                        intTotalOtros3 = 1 'intTotalOtros3 + 1
                        strOtros3 = grdCargos.TextMatrix(vlLngCont, 2)
                        vlblnbandera = True
                    Else
                        dblTotalOtros3 = dblTotalOtros3 + dblSinIVA
                        intTotalOtros3 = 1 'intTotalOtros3 + 1
                        strOtros3 = grdCargos.TextMatrix(vlLngCont, 50)
                        vlblnbandera = True
                    End If
                End If
            End If
            If vlblnbandera = False Then
                If strOtros4 = "" Or strOtros4 = grdCargos.TextMatrix(vlLngCont, 50) Then
                    If grdCargos.TextMatrix(vlLngCont, 51) Then
                        dblTotalOtros4 = dblTotalOtros4 + dblSinIVA
                        intTotalOtros4 = 1 'intTotalOtros4 + 1
                        strOtros4 = grdCargos.TextMatrix(vlLngCont, 2)
                        vlblnbandera = True
                    Else
                        dblTotalOtros4 = dblTotalOtros4 + dblSinIVA
                        intTotalOtros4 = 1 'intTotalOtros4 + 1
                        strOtros4 = grdCargos.TextMatrix(vlLngCont, 50)
                        vlblnbandera = True
                    End If
                End If
            End If
            If vlblnbandera = False Then
                If strOtros5 = "" Or strOtros5 = grdCargos.TextMatrix(vlLngCont, 50) Then
                    If grdCargos.TextMatrix(vlLngCont, 51) Then
                        dblTotalOtros5 = dblTotalOtros5 + dblSinIVA
                        intTotalOtros5 = 1 'intTotalOtros5 + 1
                        strOtros5 = grdCargos.TextMatrix(vlLngCont, 2)
                        vlblnbandera = True
                    Else
                        dblTotalOtros5 = dblTotalOtros5 + dblSinIVA
                        intTotalOtros5 = 1 'intTotalOtros5 + 1
                        strOtros5 = grdCargos.TextMatrix(vlLngCont, 50)
                        vlblnbandera = True
                    End If
                End If
            End If
            
        End If
        
        
        
           
    Next vlLngCont
    'Parametros de totales
    alstrParametros(33) = "dblTotalPaqueteCorta;" & dblTotalPaqueteCorta
    alstrParametros(34) = "intTotalPaqueteCortaUNIDAD;" & intTotalPaqueteCortaUNIDAD
    alstrParametros(35) = "dblTotalPaquete;" & dblTotalPaquete
    alstrParametros(36) = "intTotalPaqueteUNIDAD;" & intTotalPaqueteUNIDAD
    alstrParametros(37) = "dblTotalCuartoPrivado;" & dblTotalCuartoPrivado
    alstrParametros(38) = "intTotalCuartoPrivadoUNIDAD;" & intTotalCuartoPrivadoUNIDAD
    alstrParametros(39) = "dblTotalCuna;" & dblTotalCuna
    alstrParametros(40) = "intTotalCunaUNIDAD;" & intTotalCunaUNIDAD
    alstrParametros(41) = "dblTotalCrecimiento;" & dblTotalCrecimiento
    alstrParametros(42) = "intTotalCrecimientoUNIDAD;" & intTotalCrecimientoUNIDAD
    alstrParametros(43) = "dblTotalTerapiaNeonatal;" & dblTotalTerapiaNeonatal
    alstrParametros(44) = "intTotalTerapiaNeonatalUNIDAD;" & intTotalTerapiaNeonatalUNIDAD
    alstrParametros(45) = "dblTotalTerapiaPediatrica;" & dblTotalTerapiaPediatrica
    alstrParametros(46) = "intTotalTerapiaPediatricaUNIDAD;" & intTotalTerapiaPediatricaUNIDAD
    alstrParametros(47) = "dblTotalTerapiaAdulto;" & dblTotalTerapiaAdulto
    alstrParametros(48) = "intTotalTerapiaAdultoUNIDAD;" & intTotalTerapiaAdultoUNIDAD
    alstrParametros(49) = "dblTotalTerapiaIntermediaNeonatal;" & dblTotalTerapiaIntermediaNeonatal
    alstrParametros(50) = "intTotalTerapiaIntermediaNeonatalUNIDAD;" & intTotalTerapiaIntermediaNeonatalUNIDAD
    alstrParametros(51) = "dblTotalTerapiaIntermediaPediatrica;" & dblTotalTerapiaIntermediaPediatrica
    alstrParametros(52) = "intTotalTerapiaIntermediaPediatricaUNIDAD;" & intTotalTerapiaIntermediaPediatricaUNIDAD
    alstrParametros(53) = "dblTotalTerapiaIntermediaAdulto;" & dblTotalTerapiaIntermediaAdulto
    alstrParametros(54) = "intTotalTerapiaIntermediaAdultoUNIDAD;" & intTotalTerapiaIntermediaAdultoUNIDAD
    alstrParametros(55) = "dblTotalCoronaria;" & dblTotalCoronaria
    alstrParametros(56) = "intTotalCoronariaUNIDAD;" & intTotalCoronariaUNIDAD
    alstrParametros(57) = "dblTotalSalaCirugiaMenorPrimera;" & dblTotalSalaCirugiaMenorPrimera
    alstrParametros(58) = "intTotalSalaCirugiaMenorPrimeraUNIDAD;" & intTotalSalaCirugiaMenorPrimeraUNIDAD
    alstrParametros(59) = "dblTotalSalaCirugiaMenorFraccion;" & dblTotalSalaCirugiaMenorFraccion
    alstrParametros(60) = "intTotalSalaCirugiaMenorFraccionUNIDAD;" & intTotalSalaCirugiaMenorFraccionUNIDAD
    alstrParametros(61) = "dblTotalSalaCortaEstancia;" & dblTotalSalaCortaEstancia
    alstrParametros(62) = "intTotalSalaCortaEstanciaUNIDAD;" & intTotalSalaCortaEstanciaUNIDAD
    alstrParametros(63) = "dblTotalPaqueteOtros;" & dblTotalPaqueteOtros
    alstrParametros(64) = "intTotalPaqueteOtrosUNIDAD;" & intTotalPaqueteOtrosUNIDAD
            
    pCargaParameterFields alstrParametros, vgrptReporte
    pImprimeReporte vgrptReporte, rsDatosPaciente, vlstrDestino, "Reclamación de servicios médicos de hospitalización"
    
    pInstanciaReporte vgrptReporte, "rptVidamedica2.rpt"
    
    alstrParametros2(0) = "dblTotalSalaQuirurgica;" & dblTotalSalaQuirurgica
    alstrParametros2(1) = "intTotalSalaQuirurgicaUNIDAD;" & intTotalSalaQuirurgicaUNIDAD
    alstrParametros2(2) = "dblTotalSalaQuirurgicaFraccion;" & dblTotalSalaQuirurgicaFraccion
    alstrParametros2(3) = "intTotalSalaQuirurgicaFraccionUNIDAD;" & intTotalSalaQuirurgicaFraccionUNIDAD
    alstrParametros2(4) = "dblTotalSalaRecuperacion;" & dblTotalSalaRecuperacion
    alstrParametros2(5) = "intTotalSalaRecuperacionUNIDAD;" & intTotalSalaRecuperacionUNIDAD
    alstrParametros2(6) = "dblTotalSalaExpulsionHora;" & dblTotalSalaExpulsionHora
    alstrParametros2(7) = "intTotalSalaExpulsionHoraUNIDAD;" & intTotalSalaExpulsionHoraUNIDAD
    alstrParametros2(8) = "dblTotalSalaExpulsionFraccion;" & dblTotalSalaExpulsionFraccion
    alstrParametros2(9) = "intTotalSalaExpulsionFraccionUNIDAD;" & intTotalSalaExpulsionFraccionUNIDAD
    alstrParametros2(10) = "dblTotalSalaLabor;" & dblTotalSalaLabor
    alstrParametros2(11) = "intTotalSalaLaborUNIDAD;" & intTotalSalaLaborUNIDAD
    alstrParametros2(12) = "dblTotalOtrosConceptosSalas;" & dblTotalOtrosConceptosSalas
    alstrParametros2(13) = "intTotalOtrosConceptosSalasUNIDAD;" & intTotalOtrosConceptosSalasUNIDAD
    alstrParametros2(14) = "dblTotalMedicamentoGravable;" & dblTotalMedicamentoGravable
    alstrParametros2(15) = "intTotalMedicamentoGravableUNIDAD;" & intTotalMedicamentoGravableUNIDAD
    alstrParametros2(16) = "dblTotalMedicamentoNoGravable;" & dblTotalMedicamentoNoGravable
    alstrParametros2(17) = "intTotalMedicamentoNoGravableUNIDAD;" & intTotalMedicamentoNoGravableUNIDAD
    alstrParametros2(18) = "dblTotalOtrosEstancia;" & dblTotalOtrosEstancia
    alstrParametros2(19) = "intTotalOtrosEstanciaUNIDAD;" & intTotalOtrosEstanciaUNIDAD
    alstrParametros2(20) = "dblTotalOtrosTerapiaIntensiva;" & dblTotalOtrosTerapiaIntensiva
    alstrParametros2(21) = "intTotalOtrosTerapiaIntensivaUNIDAD;" & intTotalOtrosTerapiaIntensivaUNIDAD
    alstrParametros2(22) = "dblTotalOtrosTerapiaIntermedia;" & dblTotalOtrosTerapiaIntermedia
    alstrParametros2(23) = "intTotalOtrosTerapiaIntermediaUNIDAD;" & intTotalOtrosTerapiaIntermediaUNIDAD
    alstrParametros2(24) = "dblTotalOtrosCuna;" & dblTotalOtrosCuna
    alstrParametros2(25) = "intTotalOtrosCunaUNIDAD;" & intTotalOtrosCunaUNIDAD
    alstrParametros2(26) = "dblTotalMaterialCuracion;" & dblTotalMaterialCuracion
    alstrParametros2(27) = "intTotalMaterialCuracionUNIDAD;" & intTotalMaterialCuracionUNIDAD
    alstrParametros2(28) = "dblTotalMaterialQuirofanos;" & dblTotalMaterialQuirofanos
    alstrParametros2(29) = "intTotalMaterialQuirofanosUNIDAD;" & intTotalMaterialQuirofanosUNIDAD
    alstrParametros2(30) = "dblTotalEquiposUtilizados;" & dblTotalEquiposUtilizados
    alstrParametros2(31) = "intTotalEquiposUtilizadosUNIDAD;" & intTotalEquiposUtilizadosUNIDAD
    alstrParametros2(32) = "dblTotalEquiposDesechables;" & dblTotalEquiposDesechables
    alstrParametros2(33) = "intTotalEquiposDesechablesUNIDAD;" & intTotalEquiposDesechablesUNIDAD
    alstrParametros2(34) = "dblTotalLaboratorioClinico;" & dblTotalLaboratorioClinico
    alstrParametros2(35) = "intTotalLaboratorioClinicoUNIDAD;" & intTotalLaboratorioClinicoUNIDAD
    alstrParametros2(36) = "dblTotalLaboratorioPAtologia;" & dblTotalLaboratorioPAtologia
    alstrParametros2(37) = "intTotalLaboratorioPAtologiaUNIDAD;" & intTotalLaboratorioPAtologiaUNIDAD
    alstrParametros2(38) = "dblTotalRadiologiaDiagnostica;" & dblTotalRadiologiaDiagnostica
    alstrParametros2(39) = "intTotalRadiologiaDiagnosticaUNIDAD;" & intTotalRadiologiaDiagnosticaUNIDAD
    alstrParametros2(40) = "dblTotalRadiologiaTerapeutica;" & dblTotalRadiologiaTerapeutica
    alstrParametros2(41) = "intTotalRadiologiaTerapeuticaUNIDAD;" & intTotalRadiologiaTerapeuticaUNIDAD
    alstrParametros2(42) = "dblTotalResonanciaMAgnetica;" & dblTotalResonanciaMAgnetica
    alstrParametros2(43) = "intTotalResonanciaMAgneticaUNIDAD;" & intTotalResonanciaMAgneticaUNIDAD
    alstrParametros2(44) = "dblTotalMedicinaNuclear;" & dblTotalMedicinaNuclear
    alstrParametros2(45) = "intTotalMedicinaNuclearUNIDAD;" & intTotalMedicinaNuclearUNIDAD
    alstrParametros2(46) = "dblTotalTomografia;" & dblTotalTomografia
    alstrParametros2(47) = "intTotalTomografiaUNIDAD;" & intTotalTomografiaUNIDAD
    alstrParametros2(48) = "dblTotalUltrasonido;" & dblTotalUltrasonido
    alstrParametros2(49) = "intTotalUltrasonidoUNIDAD;" & intTotalUltrasonidoUNIDAD
    alstrParametros2(50) = "dblTotalInhaloterapia;" & dblTotalInhaloterapia
    alstrParametros2(51) = "intTotalInhaloterapiaUNIDAD;" & intTotalInhaloterapiaUNIDAD
    alstrParametros2(52) = "dblTotalEquipoAnestesia;" & dblTotalEquipoAnestesia
    alstrParametros2(53) = "intTotalEquipoAnestesiaUNIDAD;" & intTotalEquipoAnestesiaUNIDAD
    
    alstrParametros2(54) = "FechaIngresoDia;"
    alstrParametros2(55) = "FechaIngresoMes;"
    alstrParametros2(56) = "FechaIngresoAño;"
                      
    If Not IsNull(rsDatosPaciente!Ingreso) Then
        alstrParametros2(54) = "FechaIngresoDia;" & Format(rsDatosPaciente!Ingreso, "dd")
        alstrParametros2(55) = "FechaIngresoMes;" & Format(rsDatosPaciente!Ingreso, "mm")
        alstrParametros2(56) = "FechaIngresoAño;" & Format(rsDatosPaciente!Ingreso, "yyyy")

    End If
        alstrParametros2(57) = "FechaEgresoDia;"
        alstrParametros2(58) = "FechaEgresoMes;"
        alstrParametros2(59) = "FechaEgresoAño;"
    
    If Not IsNull(rsDatosPaciente!Egreso) Then
        alstrParametros2(57) = "FechaEgresoDia;" & Format(rsDatosPaciente!Egreso, "dd")
        alstrParametros2(58) = "FechaEgresoMes;" & Format(rsDatosPaciente!Egreso, "mm")
        alstrParametros2(59) = "FechaEgresoAño;" & Format(rsDatosPaciente!Egreso, "yyyy")
    End If
    alstrParametros2(60) = "dblTotalInstrumentalMicrocirugia;" & dblTotalInstrumentalMicrocirugia
    alstrParametros2(61) = "intTotalInstrumentalMicrocirugia;" & intTotalInstrumentalMicrocirugia
    'Otros
    alstrParametros2(62) = "dblTotalOtros1;" & dblTotalOtros1
    alstrParametros2(63) = "intTotalOtros1;" & intTotalOtros1
    alstrParametros2(64) = "dblTotalOtros2;" & dblTotalOtros2
    alstrParametros2(65) = "intTotalOtros2;" & intTotalOtros2
    alstrParametros2(66) = "dblTotalOtros3;" & dblTotalOtros3
    alstrParametros2(67) = "intTotalOtros3;" & intTotalOtros3
    alstrParametros2(68) = "dblTotalOtros4;" & dblTotalOtros4
    alstrParametros2(69) = "intTotalOtros4;" & intTotalOtros4
    alstrParametros2(70) = "dblTotalOtros5;" & dblTotalOtros5
    alstrParametros2(71) = "intTotalOtros5;" & intTotalOtros5
    'OTros nombre
    alstrParametros2(72) = "strOtros1;" & strOtros1
    alstrParametros2(73) = "strOtros2;" & strOtros2
    alstrParametros2(74) = "strOtros3;" & strOtros3
    alstrParametros2(75) = "strOtros4;" & strOtros4
    alstrParametros2(76) = "strOtros5;" & strOtros5
    
    
    
    Set rsDatosPaciente = frsEjecuta_SP(Trim(txtMovimientoPaciente.Text) & "|0|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable, "SP_PVSELDATOSPACIENTEVITA")

    pCargaParameterFields alstrParametros2, vgrptReporte
    pImprimeReporte vgrptReporte, rsDatosPaciente, vlstrDestino, "Reclamación de servicios médicos de hospitalización"
    
End Sub

Private Sub cmdInvertirSeleccion_Click()
Dim X As Long
    If grdCargos.Row > 0 Then
        grdCargos.Col = 0
        GrdCargos_DblClick
    End If
Exit Sub
End Sub

Private Sub cmdQuitarSeleccion_Click()
On Error GoTo NotificaError
Dim X As Long
    
    lstFacturas.Enabled = True
    
    grdCargos.Redraw = False
    
    For X = 1 To grdCargos.Rows - 1
        If Trim(grdCargos.TextMatrix(X, 0)) = "*" Then
            grdCargos.TextMatrix(X, 0) = ""
        End If
    Next X
    
    X = 0
    For X = 0 To lstFacturas.ListCount - 1
        lstFacturas.Selected(X) = False
    Next
    
    grdCargos.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdQuitarSeleccion_Click"))
    Unload Me
End Sub

Private Sub cmdSeleccionarTodos_Click()
On Error GoTo NotificaError
    Dim X As Long
    
    lstFacturas.Enabled = True
    
    grdCargos.Redraw = False
    
    For X = 1 To grdCargos.Rows - 1
        grdCargos.TextMatrix(X, 0) = "*"
    Next X

    X = 0
    For X = 0 To lstFacturas.ListCount - 1
        lstFacturas.Selected(X) = True
    Next
    
    grdCargos.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSeleccionarTodos_Click"))
    Unload Me
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
   
    '-------------------------------------------------------
    'Revisamos si tiene licencia Vitamédica
    '-------------------------------------------------------
    If Not fblnLicenciaVitamedica Then
        MsgBox SIHOMsg(1575), vbExclamation, "Mensaje"
        Unload Me
    End If

      
    frmInterfazVitamedica.Refresh
    
    pConfiguraGridCargos
    
    vgstrNombreForm = Me.Name

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = vbKeyEscape Then
        If vgstrEstadoManto <> "CE" Then
           txtMovimientoPaciente.Text = ""
        End If
        KeyAscii = 0
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Me.Icon = frmMenuPrincipal.Icon

    vllngPersonaGraba = 0
    vgstrEstadoManto = ""
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Cargos) > 0 Then
      If fintEsInterno(vglngNumeroLogin, enmTipoProceso.Cargos) = 1 Then
        OptTipoPaciente(0).Value = True
      Else
        OptTipoPaciente(1).Value = True
      End If
    End If
    
    vgintEmpresa = 0
    vgintTipoPaciente = 0
    pinhabilita
    pLimpiaGrid grdCargos
    pConfiguraGridCargos
    pLimpiaGrid grdTXT
    pConfiguraGridTXT
    pCorreoElectronicoDestinatario
    
    vlstrSentencia = "select vchvalor from siparametro where vchnombre = 'BITVALIDARFORMATOVITAMEDICA'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then vlblnFormatoVitamedica = IIf(IsNull(rs!vchvalor), "", rs!vchvalor)
    rs.Close
    cmdImprimir.Enabled = False
    If vlblnFormatoVitamedica = False Then
        cboTipoIngreso.Visible = False
        cboTipoEgreso.Visible = False
        cboTipoFactura.Visible = False
        cboFrecuencia.Visible = False
        lblTipoIngreso.Visible = False
        lblTipoEgreso.Visible = False
        lblTipoFactura.Visible = False
        lblFrecuencia.Visible = False
        lblNomina.Visible = False
        lblBeneficiario.Visible = False
        lblClaveMedico.Visible = False
        lblCodigoCPT.Visible = False
        lblCodigoICD.Visible = False
        txtNomina.Visible = False
        txtBeneficiario.Visible = False
        txtClaveMedico.Visible = False
        txtCodigoCPT.Visible = False
        txtCodigoICD.Visible = False
        txtDescripcionDiagnostico.Visible = False
        txtDescripcionProc.Visible = False
        lblDescripcionDiagnostico.Visible = False
        lblDescripcionProc.Visible = False
        FrePaciente.Height = 2370
        FreDetalle.Height = 4905
        FreDetalle.Top = 2755
        FreBotonera.Width = 1410
        cmdImprimir.Visible = False
        grdCargos.Height = 4500
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo NotificaError
    
    If vgstrEstadoManto = "C" Then
        Cancel = 1
        vgstrEstadoManto = ""
        pCancelar
        pCancelarConsultaTXT
        If SSTVitamedica.Tab = 0 Then
            txtMovimientoPaciente.Text = ""
            pLimpiaGrid grdCargos
            pEnfocaTextBox txtMovimientoPaciente
        Else
            mskFechaInicial.Text = fdtmServerFecha
            mskFechaFinal.Text = fdtmServerFecha
            pLimpiaGrid grdTXT
            mskFechaInicial.SetFocus
        End If
    Else
        If vgstrEstadoManto = "CE" Then
            If SSTVitamedica.Tab = 0 Then
                Cancel = 1
                vgstrEstadoManto = "C"
            End If
        Else
            If SSTVitamedica.Tab = 1 Then
                Cancel = 1
                SSTVitamedica.Tab = 0
                pCancelar
                pConfiguraGridCargos
                txtMovimientoPaciente.Text = ""
                pLimpiaGrid grdCargos
                pEnfocaTextBox txtMovimientoPaciente
            End If
        End If
    End If
    cmdImprimir.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
    Unload Me
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub GrdCargos_DblClick()
    Dim X As Long
    
    If grdCargos.Row > 0 Then
        grdCargos.Col = 0
        pMarca "*", grdCargos.Row
            
        X = 0
        For X = 0 To lstFacturas.ListCount - 1
            vlblnFiltroFacturasTrabajando = False
            lstFacturas.Selected(X) = False
            vlblnFiltroFacturasTrabajando = True
        Next
        
        lstFacturas.Enabled = False
    End If
    
    Exit Sub
End Sub

Private Sub lstFacturas_Click()
    If vlblnFiltroFacturasTrabajando Then
        pCargaSeleccionDeFacturas
    End If
End Sub

Private Sub lstFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaFinal_GotFocus()
    pSelMkTexto mskFechaFinal
End Sub

Private Sub mskFechaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub mskFechaInicial_GotFocus()
    pSelMkTexto mskFechaInicial
End Sub

Private Sub mskFechaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtBeneficiario_GotFocus()
     pSelTextBox txtBeneficiario
End Sub

Private Sub txtBeneficiario_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtBeneficiario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtClaveMedico_GotFocus()
    pSelTextBox txtClaveMedico
End Sub

Private Sub txtClaveMedico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtClaveMedico_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodigoCPT_GotFocus()
    pSelTextBox txtCodigoCPT
End Sub

Private Sub txtCodigoCPT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtCodigoCPT_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCodigoICD_GotFocus()
    pSelTextBox txtCodigoICD
End Sub

Private Sub txtCodigoICD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtCodigoICD_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCuentaPaciente_GotFocus()
    pSelTextBox txtCuentaPaciente
End Sub


Private Sub txtCuentaPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtCuentaPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                If optPaciente(3).Value Then 'Externos
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = False
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                Else
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = False
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha ing."", TO_CHAR(ExPacienteIngreso.dtmFechaHoraEgreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtCuentaPaciente.Text = .flngRegresaPaciente()
                
                If txtCuentaPaciente <> -1 Then
                    txtCuentaPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtCuentaPaciente.Text = ""
                End If
            End With
        Else
            If optPaciente(2).Value Then 'Internos
                vlstrSentencia = "SELECT rtrim(AdPaciente.vchApellidoPaterno)||' '||rtrim(AdPaciente.vchApellidoMaterno)||' '||rtrim(AdPaciente.vchNombre) as Nombre, " & _
                        "AdAdmision.intCveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "AdAdmision.tnyCveTipoPaciente cveTipoPaciente, AdTipoPaciente.vchDescripcion as Tipo,  " & _
                        "AdAdmision.vchNumCuarto Cuarto," & _
                        "AdAdmision.bitCuentaCerrada CuentaCerrada, " & _
                        "AdAdmision.VCHNUMAFILIACION NUMAFILIACION, " & _
                        "AdAdmision.DTMFECHAEGRESO FECHAEGRESO, " & _
                        "AdAdmision.DTMFECHAINGRESO FECHAINGRESO " & _
                        "FROM AdAdmision " & _
                        "INNER JOIN AdPaciente ON AdAdmision.numCvePaciente = AdPaciente.numCvePaciente " & _
                        "INNER JOIN AdTipoPaciente ON AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON ADADMISION.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where AdAdmision.numNumCuenta = " & txtCuentaPaciente.Text & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            Else 'Externos
                vlstrSentencia = "SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as Nombre, " & _
                        "RegistroExterno.intClaveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "RegistroExterno.vchNumAfiliacion NUMAFILIACION,  " & _
                        "RegistroExterno.dtmFecha FECHAEGRESO,  " & _
                        "RegistroExterno.dtmFechaEgreso FECHAINGRESO,  " & _
                        "RegistroExterno.tnyCveTipoPaciente as cveTipoPaciente, AdTipoPaciente.vchDescripcion  as Tipo, '' as Cuarto, RegistroExterno.bitCuentaCerrada CuentaCerrada " & _
                        "FROM RegistroExterno " & _
                        "INNER JOIN Externo ON RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
                        "INNER JOIN AdTipoPaciente ON RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON REGISTROEXTERNO.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where RegistroExterno.intNumCuenta = " & txtCuentaPaciente.Text & " And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            End If
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            
            If rs.RecordCount <> 0 Then
                'If rs!CuentaCerrada = 0 Then
                
                    mskFechaInicial.Enabled = True
                    mskFechaFinal.Enabled = True
                                        
                    vgstrEstadoManto = "C" 'Cargando
                    txtCuentaPaciente.Locked = True
                    optPaciente(2).Enabled = False
                    optPaciente(3).Enabled = False
                    txtNombrePaciente.Text = rs!Nombre
                    
                    If rs!cveTipoPaciente <> 2 Then
                        'El paciente seleccionado no pertenece a un convenio.
                        MsgBox SIHOMsg(351), vbExclamation, "Mensaje"
                        vgstrEstadoManto = ""
                        pCancelarConsultaTXT
                        txtCuentaPaciente.Text = ""
                        pEnfocaTextBox txtCuentaPaciente
                        Exit Sub
                    End If
                    
                    cmdCargar.Enabled = True
                    cmdCargar.SetFocus
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                vgstrEstadoManto = ""
                pCancelarConsultaTXT
                txtCuentaPaciente.Text = ""
                pEnfocaTextBox txtCuentaPaciente
                
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaPaciente_KeyDown"))
    Unload Me
End Sub


Private Sub txtCuentaPaciente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optPaciente(2).Value = UCase(Chr(KeyAscii)) = "I"
            optPaciente(3).Value = UCase(Chr(KeyAscii)) = "E"
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuentaPaciente_KeyPress"))
    Unload Me
End Sub

Private Sub txtDescripcionDiagnostico_GotFocus()
    pSelTextBox txtDescripcionDiagnostico
End Sub

Private Sub txtDescripcionDiagnostico_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtDescripcionDiagnostico_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDescripcionProc_GotFocus()
    pSelTextBox txtDescripcionProc
End Sub

Private Sub txtDescripcionProc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtDescripcionProc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
    pSelTextBox txtMovimientoPaciente
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsFactura As New ADODB.Recordset
    Dim vlstrFactura As String
    Dim rsPreautorizacion As New ADODB.Recordset
    Dim vlstrPreautorizacion As String
    Dim vlintConvenio As Integer
    
    vlblnCargos = True
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtMovimientoPaciente.Text) = "" Then
            With FrmBusquedaPacientes
                If OptTipoPaciente(1).Value Then 'Externos
                    .vgstrTipoPaciente = "E"
                    .Caption = .Caption & " externos"
                    .vgblnPideClave = False
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSoloActivos.Enabled = True
                    .optSinFacturar.Enabled = True
                    .optTodos.Enabled = False
                    .optSinFacturar.Value = True
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,1700,4100"
                Else
                    .vgstrTipoPaciente = "I"  'Internos
                    .vgblnPideClave = False
                    .Caption = .Caption & " internos"
                    .vgIntMaxRecords = 100
                    .vgstrMovCve = "M"
                    .optSinFacturar.Value = True
                    .optSinFacturar.Enabled = True
                    .optSoloActivos.Enabled = True
                    .optTodos.Enabled = False
                    .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha ing."", TO_CHAR(ExPacienteIngreso.dtmFechaHoraEgreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                    .vgstrTamanoCampo = "800,3400,990,990,4100"
                End If
                
                txtMovimientoPaciente.Text = .flngRegresaPaciente()
                
                If txtMovimientoPaciente <> -1 Then
                    txtMovimientoPaciente_KeyDown vbKeyReturn, 0
                Else
                    txtMovimientoPaciente.Text = ""
                End If
            End With
        Else
            If OptTipoPaciente(0).Value Then 'Internos
                vlstrSentencia = "SELECT rtrim(AdPaciente.vchApellidoPaterno)||' '||rtrim(AdPaciente.vchApellidoMaterno)||' '||rtrim(AdPaciente.vchNombre) as NombreCompleto, " & _
                        "AdAdmision.intCveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "AdAdmision.tnyCveTipoPaciente cveTipoPaciente, AdTipoPaciente.vchDescripcion as Tipo,  " & _
                        "AdAdmision.vchNumCuarto Cuarto," & _
                        "AdAdmision.bitCuentaCerrada CuentaCerrada, " & _
                        "AdAdmision.VCHNUMAFILIACION NUMAFILIACION, " & _
                        "AdAdmision.DTMFECHAEGRESO FECHAEGRESO, " & _
                        "AdAdmision.DTMFECHAINGRESO FECHAINGRESO, " & _
                        "rtrim(AdPaciente.vchApellidoPaterno) as Paterno, " & _
                        "rtrim(AdPaciente.vchApellidoMaterno) as Materno, " & _
                        "rtrim(AdPaciente.vchNombre) as Nombre " & _
                        "FROM AdAdmision " & _
                        "INNER JOIN AdPaciente ON AdAdmision.numCvePaciente = AdPaciente.numCvePaciente " & _
                        "INNER JOIN AdTipoPaciente ON AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON ADADMISION.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where AdAdmision.numNumCuenta = " & txtMovimientoPaciente.Text & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            Else 'Externos
                vlstrSentencia = "SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as NombreCompleto, " & _
                        "RegistroExterno.intClaveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, " & _
                        "RegistroExterno.vchNumAfiliacion NUMAFILIACION,  " & _
                        "RegistroExterno.dtmFecha FECHAEGRESO,  " & _
                        "RegistroExterno.dtmFechaEgreso FECHAINGRESO,  " & _
                        "RegistroExterno.tnyCveTipoPaciente as cveTipoPaciente, AdTipoPaciente.vchDescripcion  as Tipo, '' as Cuarto, RegistroExterno.bitCuentaCerrada CuentaCerrada, " & _
                        "rtrim(chrApePaterno) as Paterno, " & _
                        "rtrim(chrApeMaterno) as Materno, " & _
                        "rtrim(chrNombre) as Nombre " & _
                        "FROM RegistroExterno " & _
                        "INNER JOIN Externo ON RegistroExterno.intNumPaciente = Externo.intNumPaciente " & _
                        "INNER JOIN AdTipoPaciente ON RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente " & _
                        "LEFT OUTER Join CcEmpresa ON RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa " & _
                        "INNER JOIN NODEPARTAMENTO ON REGISTROEXTERNO.INTCVEDEPARTAMENTO = NODEPARTAMENTO.SMICVEDEPARTAMENTO " & _
                        "Where RegistroExterno.intNumCuenta = " & txtMovimientoPaciente.Text & " And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            End If
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            
            If rs.RecordCount <> 0 Then
                'If rs!CuentaCerrada = 0 Then
                
                    FrePaciente.Enabled = True
                    pHabilita
                    FreDetalle.Enabled = False
                    
                    vgstrEstadoManto = "C" 'Cargando
                    txtMovimientoPaciente.Locked = True
                    OptTipoPaciente(0).Enabled = False
                    OptTipoPaciente(1).Enabled = False
                    txtEmpresaPaciente.Locked = True
                    
                    txtPaciente.Text = rs!nombreCompleto
                    txtNumeroControl = IIf(IsNull(rs!NumAfiliacion), "", rs!NumAfiliacion)
                    pComboComanda
                    
                    pComboTipoIngreso
                    pComboMotivoEgreso
                    pComboTipoFactura
                    pComboFrecuencia
                    
                    cmdImprimir.Enabled = True
                    
                    Paterno = rs!Paterno
                    Materno = rs!Materno
                    Nombre = rs!Nombre
                    
                    If vlblnFormatoVitamedica Then
                        vlintConvenio = 3
                    Else
                        vlintConvenio = 2
                    End If
                    
                    If (IsNull(rs!empresa)) Or rs!cveTipoPaciente <> vlintConvenio Then
                        'El paciente seleccionado no pertenece a un convenio.
                        MsgBox SIHOMsg(351), vbExclamation, "Mensaje"
                        vgstrEstadoManto = ""
                        pCancelar
                        txtMovimientoPaciente.Text = ""
                        pLimpiaGrid grdCargos
                        pEnfocaTextBox txtMovimientoPaciente
                        cmdImprimir.Enabled = False
                        Exit Sub
                    Else
                         txtEmpresaPaciente.Text = IIf(IsNull(rs!empresa), "", rs!empresa)
                    End If
                    vgintTipoPaciente = rs!cveTipoPaciente
                    vgintEmpresa = rs!cveEmpresa

                    vlstrPreautorizacion = "select * from PVAUTORIZAVITAMEDICA where pvautorizavitamedica.numnumcuenta = " & Trim(txtMovimientoPaciente.Text)
                    Set rsPreautorizacion = frsRegresaRs(vlstrPreautorizacion, adLockOptimistic, adOpenDynamic)
                    If rsPreautorizacion.RecordCount <> 0 Then txtPreAutorizacion = IIf(IsNull(rsPreautorizacion!vchnumpreautorizacion), "", rsPreautorizacion!vchnumpreautorizacion)
                    rsPreautorizacion.Close
                                        
                    pLlenaCargos
                    
                    If vlblnCargos = False Then Exit Sub
                    
                    FreDetalle.Enabled = True
                    pHabilita
                    pComboComanda
                    pCargaCboFactura
                    cboComanda.SetFocus
            Else
                '¡La información no existe!
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                vgstrEstadoManto = ""
                pCancelar
                txtMovimientoPaciente.Text = ""
                pLimpiaGrid grdCargos
                pEnfocaTextBox txtMovimientoPaciente
                
            End If
            freBarra.Visible = False
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown"))
    Unload Me
End Sub
Private Sub OptTipoPaciente_Click(Index As Integer)
    On Error GoTo NotificaError
    
    pEnfocaTextBox txtMovimientoPaciente

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
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

Public Sub pHabilita()
    cboComanda.Enabled = True
    txtPreAutorizacion.Enabled = True
    lstFacturas.Enabled = True
    cmdInvertirSeleccion.Enabled = True
    cmdSeleccionarTodos.Enabled = True
    cmdQuitarSeleccion.Enabled = True
    cboFrecuencia.Enabled = True
    cboTipoIngreso.Enabled = True
    cboTipoEgreso.Enabled = True
    cboTipoFactura.Enabled = True
    txtNomina.Enabled = True
    txtBeneficiario.Enabled = True
    txtClaveMedico.Enabled = True
    txtCodigoCPT.Enabled = True
    txtCodigoICD.Enabled = True
    txtDescripcionDiagnostico.Enabled = True
    txtDescripcionProc.Enabled = True
End Sub

Private Sub txtNomina_GotFocus()
    pSelTextBox txtNomina
End Sub

Private Sub txtNomina_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtNomina_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtPreAutorizacion_GotFocus()
    pSelTextBox txtPreAutorizacion
End Sub

Private Sub txtPreAutorizacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtPreAutorizacion_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPreAutorizacion_KeyPress"))
    Unload Me
End Sub

Public Sub pinhabilita()

    txtPaciente.Enabled = False
    txtEmpresaPaciente.Enabled = False
    txtFechaIngreso.Enabled = False
    txtFechaEgreso.Enabled = False
    txtNumeroControl.Enabled = False
    cboComanda.Enabled = False
    txtPreAutorizacion.Enabled = False
    lstFacturas.Enabled = False
    
    txtPaciente.Text = ""
    txtEmpresaPaciente.Text = ""
    txtFechaIngreso.Text = "  /  /    "
    txtFechaEgreso.Text = "  /  /    "
    txtNumeroControl.Text = ""
    txtPreAutorizacion.Text = ""
    lstFacturas.Clear
    
    'strCorreoDestinatario = ""
    strAsunto = ""
    strMensaje = ""
    strRutaTXT = ""
    Paterno = ""
    Materno = ""
    Nombre = ""
    vlblnFechaInicial = True
    vlblnfechaFinal = True

    cmdInvertirSeleccion.Enabled = False
    cmdSeleccionarTodos.Enabled = False
    cmdQuitarSeleccion.Enabled = False
    cboFrecuencia.Enabled = False
    cboTipoIngreso.Enabled = False
    cboTipoEgreso.Enabled = False
    cboTipoFactura.Enabled = False
    txtNomina.Enabled = False
    txtBeneficiario.Enabled = False
    txtCodigoICD.Enabled = False
    txtCodigoCPT.Enabled = False
    txtClaveMedico.Enabled = False
    txtDescripcionDiagnostico.Enabled = False
    txtDescripcionProc.Enabled = False
    txtNomina.Text = ""
    txtBeneficiario.Text = ""
    txtCodigoICD.Text = ""
    txtCodigoCPT.Text = ""
    txtClaveMedico.Text = ""
    txtDescripcionDiagnostico.Text = ""
    txtDescripcionProc.Text = ""

End Sub

Public Sub pComboComanda()

    cboComanda.Clear
    cboComanda.AddItem "PRELIMINAR", 0
    cboComanda.AddItem "CORTE", 1
    cboComanda.AddItem "DIARIA", 2
    cboComanda.AddItem "FINAL", 3
    cboComanda.ListIndex = 0

End Sub
Public Sub pComboFrecuencia()

    cboFrecuencia.Clear
    cboFrecuencia.AddItem "FACTURA ÚNICA", 0
    cboFrecuencia.AddItem "PRIMERA FACTURA", 1
    cboFrecuencia.AddItem "CONTINUACIÓN FACTURA", 2
    cboFrecuencia.AddItem "ÚLTIMA FACTURA", 3
    cboFrecuencia.ListIndex = 0

End Sub
Public Sub pComboTipoIngreso()

    cboTipoIngreso.Clear
    cboTipoIngreso.AddItem "URGENCIA", 0
    cboTipoIngreso.AddItem "PROGRAMADO", 1
    cboTipoIngreso.AddItem "TRASLADO DE OTRA UNIDAD", 2
    cboTipoIngreso.ListIndex = 0

End Sub
Public Sub pComboMotivoEgreso()

    cboTipoEgreso.Clear
    cboTipoEgreso.AddItem "ALTA POR MEJORÍA", 0
    cboTipoEgreso.AddItem "ALTA VOLUNTARIA", 1
    cboTipoEgreso.AddItem "DEFUNCIÓN", 2
    cboTipoEgreso.AddItem "TRASLADO DE OTRO HOSPITAL", 3
    cboTipoEgreso.ListIndex = 0

End Sub
Public Sub pComboTipoFactura()

    cboTipoFactura.Clear
    cboTipoFactura.AddItem "PACIENTE INTERNADO", 0
    cboTipoFactura.AddItem "PACIENTE AMBULATORIO", 1
    cboTipoFactura.ListIndex = 0

End Sub
Public Sub pLlenaCargos()
On Error GoTo NotificaError

    Dim alstrParametros(41) As String
    Dim alstrParametros2(1) As String
    Dim vlstrSentencia As String
    Dim vlstrFacturasPaciente As String
    Dim vlstrGruposCuenta As String
    Dim rs As New ADODB.Recordset
    Dim rsInformacionFaltantePCE As ADODB.Recordset
    Dim rsEstadoCuenta As New ADODB.Recordset
    Dim rsDatosPaciente As New ADODB.Recordset
    Dim lblnFueraCatalogo As Boolean
    Dim vldblTotalPaquetes As Double
    Dim vldblIVA As Double
    Dim vlstrGrupoCuentas As String 'Agregado para caso 6776
    Dim lngConceptosAseguradora As Long
    Dim dblDescuentoEspecial As Double
    Dim X As Integer
    
    
    Me.MousePointer = 11
                
    Set rsDatosPaciente = frsEjecuta_SP(Trim(txtMovimientoPaciente.Text) & "|0|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable, "Sp_PvSelDatosPaciente")
    If rsDatosPaciente.RecordCount <> 0 Then
                                
        txtFechaIngreso = Format(rsDatosPaciente!Ingreso, "dd/mmm/yyyy hh:mm") 'Fecha Ingreso formato pantalla
        txtFechaEgreso = Format(rsDatosPaciente!Egreso, "dd/mmm/yyyy hh:mm") 'Fecha egreso formato pantalla
        
            
        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                            & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                            & "|" & 2 _
                            & "|" & 0 _
                            & "|" & 0 _
                            & "|" & 0 _
                            & "|" & "*" _
                            & "|" & 1 _
                            & "|" & fstrFechaSQL("01/01/1900 00:00:00") _
                            & "|" & fstrFechaSQL("31/12/3999 23:59:59") _
                            & "|" & vgintClaveEmpresaContable _
                            & "|" & 0 _
                            & "|" & 1
                            
        Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvInterfazVitamedica")
        If rsEstadoCuenta.EOF Then
            'No existe información con esos parámetros.
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
            vlblnCargos = False
            vgstrEstadoManto = ""
            pCancelar
            txtMovimientoPaciente.Text = ""
            pLimpiaGrid grdCargos
            pEnfocaTextBox txtMovimientoPaciente
            cmdImprimir.Enabled = False
        Else
            vlblnCargos = True
            pLimpiaGrid grdCargos
            pConfiguraGridCargos
                                  
            grdCargos.Redraw = False
                                  
            '--------------------------------------------------'
            '                     Paquetes                     '
            '--------------------------------------------------'
            vlstrFacturasPaciente = ""
            vlstrSentencia = "SELECT DISTINCT PvPaquete.CHRDESCRIPCION, PvPaquetePaciente.MNYPRECIOPAQUETE, PvPaquetePaciente.MNYPRECIOPAQUETE*(PvConceptoFacturacion.SMYIVA/100) IVA, Trim(PVPAQUETEPACIENTEFACTURADO.CHRFOLIOFACTURA) Factura, PvConceptoFacturacion.CHRDESCRIPCION conceptoFacturacion" & _
                         " FROM PvCargo " & _
                         " INNER JOIN PvPaquete ON PvCargo.INTNUMPAQUETE = PvPaquete.INTNUMPAQUETE" & _
                         " INNER JOIN PvPaquetePaciente ON PvPaquete.INTNUMPAQUETE = PvPaquetePaciente.INTNUMPAQUETE AND PvPaquetePaciente.INTMOVPACIENTE = PvCargo.INTMOVPACIENTE AND PvPaquetePaciente.CHRTIPOPACIENTE = PvCargo.CHRTIPOPACIENTE" & _
                         " LEFT JOIN PvPaquetePacienteFacturado ON PVPAQUETEPACIENTEFACTURADO.INTNUMPAQUETE = pvcargo.intNumPaquete " & _
                                    " AND PVPAQUETEPACIENTEFACTURADO.CHRTIPOPACIENTE = pvcargo.CHRTIPOPACIENTE " & _
                                    " AND PVPAQUETEPACIENTEFACTURADO.INTMOVPACIENTE = pvcargo.INTMOVPACIENTE " & _
                                    " AND Trim(PVPAQUETEPACIENTEFACTURADO.CHRFOLIOFACTURA) = Trim(pvcargo.CHRFOLIOFACTURA) " & _
                                    " AND PVPAQUETEPACIENTEFACTURADO.CHRestatus = 'F'" & _
                         " INNER JOIN PvConceptoFacturacion ON PvPaquete.SMICONCEPTOFACTURA = PvConceptoFacturacion.SMICVECONCEPTO" & _
                         " WHERE PvCargo.intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & _
                         " AND PvCargo.chrTipoPaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'" & _
                         " AND (" & 0 & " = 1 OR PvCargo.bitExcluido = " & 0 & ")" & _
                         " AND (" & "'*'" & " = '*' OR PvCargo.chrFolioFactura = '" & "<TODAS>" & "')" & _
                         " AND (PVCARGO.INTCANTIDADPAQUETE <> 0 OR PVCARGO.INTCANTIDADEXTRAPAQUETE <> 0)" & _
                         " AND (PVCARGO.INTCANTIDADPAQUETE IS NOT NULL OR PVCARGO.INTCANTIDADEXTRAPAQUETE IS NOT NULL)" & _
                         " order by PvPaquete.CHRDESCRIPCION"
            Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                        
            X = 0
            vldblTotalPaquetes = 0
            Do While Not rs.EOF
                If X > 0 Then
                    grdCargos.Rows = grdCargos.Rows + 1
                    grdCargos.Row = grdCargos.Rows - 1
                End If
                X = X + 1
                grdCargos.Row = X
            
                grdCargos.TextMatrix(grdCargos.Row, 0) = "*"
                         
                If Len(Trim(rs!chrDescripcion)) >= 100 Then
                    grdCargos.TextMatrix(grdCargos.Row, 2) = Trim(rs!chrDescripcion) 'Descripción del paquete
                Else
                    grdCargos.TextMatrix(grdCargos.Row, 2) = Trim(rs!chrDescripcion) & String(100 - Len(Trim(rs!chrDescripcion)), " ") 'Pone la DESCRIPCIÓN del insumo en 100 caracteres
                End If
                         
                grdCargos.TextMatrix(grdCargos.Row, 3) = IIf(IsNull(rsEstadoCuenta!CantidadPaquete), "", rsEstadoCuenta!CantidadPaquete) 'Cantidad paquetes
                grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(rs!MNYPRECIOPAQUETE, 2) 'Precio del paquete grid
                grdCargos.TextMatrix(grdCargos.Row, 43) = rs!MNYPRECIOPAQUETE 'Precio paquete TXT
                grdCargos.TextMatrix(grdCargos.Row, 5) = FormatCurrency((rs!MNYPRECIOPAQUETE * rsEstadoCuenta!CantidadPaquete), 2) 'Importe pantalla
                grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((rsEstadoCuenta!Descuento * rsEstadoCuenta!fuerapaquete), 2) 'Descuento grid
                grdCargos.TextMatrix(grdCargos.Row, 41) = rsEstadoCuenta!Descuento * rsEstadoCuenta!fuerapaquete 'Descuento TXT
                grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(((rs!MNYPRECIOPAQUETE * rsEstadoCuenta!CantidadPaquete) - (rsEstadoCuenta!Descuento * rsEstadoCuenta!fuerapaquete)), 2) 'Subtotal grid
                grdCargos.TextMatrix(grdCargos.Row, 44) = ((rs!MNYPRECIOPAQUETE * rsEstadoCuenta!CantidadPaquete) - (rsEstadoCuenta!Descuento * rsEstadoCuenta!fuerapaquete)) 'Subtotal2 TXT
                grdCargos.TextMatrix(grdCargos.Row, 46) = rs!MNYPRECIOPAQUETE * rsEstadoCuenta!CantidadPaquete 'Subtotal TXT
                grdCargos.TextMatrix(grdCargos.Row, 40) = grdCargos.TextMatrix(grdCargos.Row, 7) 'Importe archivoTXT
                         
                grdCargos.TextMatrix(grdCargos.Row, 24) = IIf(IsNull(rs!Factura), "", Trim(rs!Factura)) 'Folio de la factura del paquete
                grdCargos.TextMatrix(grdCargos.Row, 50) = Trim(rs!ConceptoFacturacion)
                grdCargos.TextMatrix(grdCargos.Row, 51) = 1
                
                vldblIVA = vldblIVA + rs!IVA
                vldblTotalPaquetes = vldblTotalPaquetes + rs!MNYPRECIOPAQUETE
                rs.MoveNext
            Loop
            rs.Close
            
            grdCargos.TextMatrix(grdCargos.Row, 25) = vldblIVA 'IVA por paquete
            grdCargos.TextMatrix(grdCargos.Row, 26) = vldblTotalPaquetes 'Total paquetes
            grdCargos.TextMatrix(grdCargos.Row, 45) = vgdblCantidadIvaGeneral 'Tasa IVA
            
            If IsNull(rsEstadoCuenta!CantidadPaquete) Then
                grdCargos.TextMatrix(grdCargos.Row, 8) = FormatCurrency(vldblIVA, 2) 'IVA paquetes grid
                grdCargos.TextMatrix(grdCargos.Row, 42) = vldblIVA 'IVA paquetes TXT
                If grdCargos.TextMatrix(grdCargos.Row, 44) <> "" Then
                    grdCargos.TextMatrix(grdCargos.Row, 9) = FormatCurrency((grdCargos.TextMatrix(grdCargos.Row, 44) + vldblIVA), 2) 'Total
                End If
            Else
                grdCargos.TextMatrix(grdCargos.Row, 8) = FormatCurrency((rsEstadoCuenta!CantidadPaquete * vldblIVA), 2) 'IVA paquetes grid
                grdCargos.TextMatrix(grdCargos.Row, 42) = rsEstadoCuenta!CantidadPaquete * vldblIVA 'IVA paquetes TXT
                grdCargos.TextMatrix(grdCargos.Row, 9) = FormatCurrency(((grdCargos.TextMatrix(grdCargos.Row, 44)) + (rsEstadoCuenta!CantidadPaquete * vldblIVA)), 2) 'Total
            End If
            
            
            Do While Not rsEstadoCuenta.EOF
                If rsEstadoCuenta!bitMuestraCargo = 1 Then
                    If X > 0 Then
                        grdCargos.Rows = grdCargos.Rows + 1
                        grdCargos.Row = grdCargos.Rows - 1
                    End If
                    X = X + 1
                    grdCargos.Row = X
                     
                     grdCargos.TextMatrix(grdCargos.Row, 0) = "*"
                     If (Trim(rsEstadoCuenta!Campo1)) <> "" Then
                         grdCargos.TextMatrix(grdCargos.Row, 1) = Format(rsEstadoCuenta!FechaCargo, "dd/mmm/yyyy hh:mm") 'Fecha cargo formato pantalla
                         grdCargos.TextMatrix(grdCargos.Row, 39) = Format(rsEstadoCuenta!FechaCargo, "dd/mm/yyyy") 'Fecha cargo formato archivo TXT
                         grdCargos.TextMatrix(grdCargos.Row, 35) = Format(rsEstadoCuenta!FechaCargo, "hh:mm") ' Hora de cargo
                     End If
                     grdCargos.TextMatrix(grdCargos.Row, 36) = Trim(rsEstadoCuenta!TipoCargo) 'Código
                     grdCargos.TextMatrix(grdCargos.Row, 2) = Trim(rsEstadoCuenta!Campo4) 'Descripción
                     grdCargos.TextMatrix(grdCargos.Row, 3) = rsEstadoCuenta!Campo5 'Cantidad
                     grdCargos.TextMatrix(grdCargos.Row, 4) = FormatCurrency(rsEstadoCuenta!campo6, 2) 'Precio grid
                     grdCargos.TextMatrix(grdCargos.Row, 43) = rsEstadoCuenta!campo6 'Precio TXT
                     grdCargos.TextMatrix(grdCargos.Row, 5) = FormatCurrency(rsEstadoCuenta!Campo7, 2) 'Importe pantalla
                     grdCargos.TextMatrix(grdCargos.Row, 6) = FormatCurrency((rsEstadoCuenta!Campo5 * rsEstadoCuenta!Descuento), 2) 'Descuento grid
                     grdCargos.TextMatrix(grdCargos.Row, 41) = rsEstadoCuenta!Campo5 * rsEstadoCuenta!Descuento 'Descuento TXT
                     grdCargos.TextMatrix(grdCargos.Row, 7) = FormatCurrency(((rsEstadoCuenta!campo6 * rsEstadoCuenta!Campo5) - (rsEstadoCuenta!Campo5 * rsEstadoCuenta!Descuento)), 2) 'Subtotal grid
                     grdCargos.TextMatrix(grdCargos.Row, 44) = (rsEstadoCuenta!campo6 * rsEstadoCuenta!Campo5) - (rsEstadoCuenta!Campo5 * rsEstadoCuenta!Descuento) 'Subtotal2 TXT
                     grdCargos.TextMatrix(grdCargos.Row, 46) = (rsEstadoCuenta!campo6 * rsEstadoCuenta!Campo5) 'Subtotal TXT
                     grdCargos.TextMatrix(grdCargos.Row, 8) = FormatCurrency(rsEstadoCuenta!IVA, 2) 'IVA grid
                     grdCargos.TextMatrix(grdCargos.Row, 42) = rsEstadoCuenta!IVA 'IVA TXT
                     grdCargos.TextMatrix(grdCargos.Row, 9) = FormatCurrency((((rsEstadoCuenta!campo6 * rsEstadoCuenta!Campo5) - (rsEstadoCuenta!Campo5 * rsEstadoCuenta!Descuento)) + rsEstadoCuenta!IVA), 2) 'Total
                     grdCargos.TextMatrix(grdCargos.Row, 40) = grdCargos.TextMatrix(grdCargos.Row, 7) 'Importe archivoTXT
                     grdCargos.TextMatrix(grdCargos.Row, 10) = rsEstadoCuenta!montofacturado 'Monto facturado
                     grdCargos.TextMatrix(grdCargos.Row, 11) = rsEstadoCuenta!ImporteGravado 'Importe gravado
                     grdCargos.TextMatrix(grdCargos.Row, 12) = rsEstadoCuenta!ivacargo 'IVA cargo
                     grdCargos.TextMatrix(grdCargos.Row, 32) = rsEstadoCuenta!montopago 'Monto
                     grdCargos.TextMatrix(grdCargos.Row, 33) = rsEstadoCuenta!totalpagar 'Total a pagar
                     
                     grdCargos.TextMatrix(grdCargos.Row, 24) = IIf(IsNull(rsEstadoCuenta!Factura), "", Trim(rsEstadoCuenta!Factura)) 'Folio de la factura del cargo
                     
                     grdCargos.TextMatrix(grdCargos.Row, 13) = fdtmServerFecha
                     grdCargos.TextMatrix(grdCargos.Row, 14) = fdtmServerHora
                     grdCargos.TextMatrix(grdCargos.Row, 15) = Trim(vgstrNombreHospitalCH)
                     grdCargos.TextMatrix(grdCargos.Row, 16) = vgstrRfCCH
                     grdCargos.TextMatrix(grdCargos.Row, 17) = rsDatosPaciente!NumPaciente
                     grdCargos.TextMatrix(grdCargos.Row, 18) = txtMovimientoPaciente.Text
                     grdCargos.TextMatrix(grdCargos.Row, 19) = IIf(OptTipoPaciente(0).Value, "INTERNO", "EXTERNO")
                     grdCargos.TextMatrix(grdCargos.Row, 20) = IIf(IsNull(rsDatosPaciente!Nombre), "", rsDatosPaciente!Nombre)
                     If Not IsNull(rsDatosPaciente!Egreso) Then
                         grdCargos.TextMatrix(grdCargos.Row, 37) = Format(rsDatosPaciente!Egreso, "dd/mm/yyyy") 'Fecha egreso formato arcvivo TXT
                     End If
                     grdCargos.TextMatrix(grdCargos.Row, 22) = IIf(IsNull(rsDatosPaciente!tipo), "", rsDatosPaciente!tipo) ' Tipo paciente
                     grdCargos.TextMatrix(grdCargos.Row, 23) = IIf(IsNull(rsDatosPaciente!empresa), "", rsDatosPaciente!empresa) 'Nombre de la empresa
                     If Not IsNull(rsDatosPaciente!Ingreso) Then
                         grdCargos.TextMatrix(grdCargos.Row, 38) = Format(rsDatosPaciente!Ingreso, "dd/mm/yyyy") 'Fecha Ingreso formato arcvivo TXT
                     End If
                     grdCargos.TextMatrix(grdCargos.Row, 50) = Trim(rsEstadoCuenta!campo3)
                     grdCargos.TextMatrix(grdCargos.Row, 51) = 0

                
                End If
                     
                rsEstadoCuenta.MoveNext
            Loop
     
            grdCargos.Redraw = True
     End If
     rsEstadoCuenta.Close
 Else
     'No se encontró la información del paciente.
     MsgBox SIHOMsg(355), vbOKOnly + vbExclamation, "Mensaje"
 End If
 
 rsDatosPaciente.Close
 Me.MousePointer = 0


Exit Sub
NotificaError:
    Me.MousePointer = 0
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenarCargos"))
End Sub

'Exporta archivo TXT PARA Vitamédica
Public Sub pExportaTXT()

On Error GoTo NotificaError
    Dim vlLngCont As Long
    Dim vlblnSiGeneraraInfo As Boolean
    Dim vlstrDescripcion As String
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlstrvchvalor As String
    Dim vlsdescuento As Double
    Dim vlIVA As Double
    Dim vlPU As Double
    Dim vlImporte As Double
    Dim vlsumasubtotal As Double
    Dim vlsumadescuento As Double
    Dim vlsumasubtotal2 As Double
    Dim vlsumasubtotalIVA As Double
    Dim vlsumatotal As Double
    Dim vlstrNombreTXT As String
    Dim i As Integer
    Dim ldblCantidad As Double
    Dim lstrDescripcion As String
    Dim X As Integer
    Dim vlfacturas As String

    vlblnSiGeneraraInfo = False
    vlstrvchvalor = " "
    
    vlstrSentencia = "select vchvalor from siparametro where vchnombre = 'VCHNUMPROVEEDORVITAMEDICA'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then vlstrvchvalor = IIf(IsNull(rs!vchvalor), "", rs!vchvalor)
    rs.Close
    
    
    For vlLngCont = 1 To grdCargos.Rows - 1
        If grdCargos.TextMatrix(vlLngCont, 0) = "*" Then
            vlblnSiGeneraraInfo = True
            Exit For
        End If
    Next vlLngCont
    
    If vlblnSiGeneraraInfo Then
    
        CDgArchivo.CancelError = False
        CDgArchivo.InitDir = App.Path
        CDgArchivo.Flags = cdlOFNOverwritePrompt
        CDgArchivo.FileName = Trim(cboComanda.Text) & "_" & Trim(Paterno) & "_" & Trim(Materno) & "_" & Trim(Nombre) & ".txt"
        vlstrNombreTXT = Trim(cboComanda.Text) & "_" & Trim(Paterno) & "_" & Trim(Materno) & "_" & Trim(Nombre) & ".txt"
        CDgArchivo.DialogTitle = "Mensaje"

 
        CDgArchivo.ShowSave
        If CDgArchivo.FileName <> vlstrNombreTXT Then
            strRutaTXT = CDgArchivo.FileName
        Else
            Exit Sub
        End If
       
        
        Open CDgArchivo.FileName For Output As #1  ' Open file for output.
        
        strNombreArchivoTXT = CDgArchivo.FileTitle
        
        vlfacturas = ""
        X = 0
        For X = 0 To lstFacturas.ListCount - 1
            If lstFacturas.Selected(X) = True And lstFacturas.List(X) <> "<SIN FACTURA>" Then
                If vlfacturas = "" Then
                    vlfacturas = lstFacturas.List(X)
                Else
                    vlfacturas = vlfacturas & ", " & lstFacturas.List(X)
                End If
            End If
        Next
                    
        Print #1, "HOSPITAL" & Trim(vgstrNombreHospitalCH)
        Print #1, "FECHA DE EMISION:" & Trim(grdCargos.TextMatrix(grdCargos.Row, 13))
        Print #1, "CODIGO FISCAL:" & vlfacturas
        Print #1, "CLIENTE:" & Trim(txtEmpresaPaciente.Text)
        Print #1, "PACIENTE:" & Trim(txtPaciente.Text)
        Print #1, "FECHA DE INGRESO:" & Trim(grdCargos.TextMatrix(grdCargos.Row, 38))
        Print #1, "FECHA DE EGRESO:" & Trim(grdCargos.TextMatrix(grdCargos.Row, 37))
        Print #1, "PREAUTORIZACION:" & Trim(txtPreAutorizacion)
        Print #1, "ELEGIBILIDAD:" & Trim(txtNumeroControl)
        Print #1, "PROVEEDOR:" & Trim(vlstrvchvalor)
        
        Print #1, ""
        Print #1, "CONSEC" & Chr(9) & "FECHA APLICA" & Chr(9) & "HORA APLICA" & Chr(9) & "CODIGO" & Chr(9) & "DESCRIPCION" & Chr(9) & "CANTIDAD" & Chr(9) & "P.U." & Chr(9) & "DESCTO" & Chr(9) & "IVA" & Chr(9) & "IMPORTE"
    
        vlsumasubtotal = 0
        vlsumadescuento = 0
        vlsumasubtotal2 = 0
        vlsumasubtotalIVA = 0
        vlsumatotal = 0
        i = 0
        ldblCantidad = 0
        
        For vlLngCont = 1 To grdCargos.Rows - 1
            If grdCargos.TextMatrix(vlLngCont, 0) = "*" Then
                
                i = i + 1
                
                vlsdescuento = grdCargos.TextMatrix(vlLngCont, 41)
                If IsNull(grdCargos.TextMatrix(vlLngCont, 42)) Or grdCargos.TextMatrix(vlLngCont, 42) = "" Then
                    vlIVA = 0
                    grdCargos.TextMatrix(vlLngCont, 42) = 0
                Else
                    vlIVA = grdCargos.TextMatrix(vlLngCont, 42)
                End If
                vlPU = grdCargos.TextMatrix(vlLngCont, 43)
                vlImporte = grdCargos.TextMatrix(vlLngCont, 40)
                ldblCantidad = Val(grdCargos.TextMatrix(vlLngCont, 3))
                
                If Len(Trim(grdCargos.TextMatrix(vlLngCont, 2))) >= 99 Then
                    lstrDescripcion = Mid(Trim(grdCargos.TextMatrix(vlLngCont, 2)), 1, 99) & "/"
                Else
                    lstrDescripcion = Trim(grdCargos.TextMatrix(vlLngCont, 2)) & "/" & String(99 - Len(Trim(grdCargos.TextMatrix(vlLngCont, 2))), " ")
                End If
                
                Print #1, i & Chr(9) & _
                grdCargos.TextMatrix(vlLngCont, 39) & Chr(9) & _
                grdCargos.TextMatrix(vlLngCont, 35) & Chr(9) & _
                grdCargos.TextMatrix(vlLngCont, 36) & Chr(9) & _
                lstrDescripcion & Chr(9) & _
                Format(ldblCantidad, "##########0.00") & Chr(9) & _
                Format(vlPU, "############0.00") & Chr(9) & _
                Format(vlsdescuento, "###########0.00") & Chr(9) & _
                Format(vlIVA, "###########0.00") & Chr(9) & _
                Format(vlImporte, "############0.00")
                

                vlsumasubtotal = vlsumasubtotal + IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 46))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 46)))
                vlsumadescuento = vlsumadescuento + IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 41))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 41)))
                vlsumasubtotal2 = vlsumasubtotal2 + IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 44))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 44)))
                vlsumasubtotalIVA = vlsumasubtotalIVA + IIf(IsNull(CDbl(grdCargos.TextMatrix(vlLngCont, 42))), 0, CDbl(grdCargos.TextMatrix(vlLngCont, 42)))
                vlsumatotal = vlsumasubtotal2 + vlsumasubtotalIVA
                
            End If
        Next vlLngCont
    
        Print #1, ""
        Print #1, "SUBTOTAL:" & Trim((Format(vlsumasubtotal, "###########0.00")))
        Print #1, "DESCUENTO:" & Trim((Format(vlsumadescuento, "###########0.00")))
        Print #1, "SUBTOTAL:" & Trim((Format(vlsumasubtotal2, "###########0.00")))
        Print #1, "IVA:" & Trim((Format(vlsumasubtotalIVA, "###########0.00")))
        Print #1, "TOTAL:" & Trim((Format(vlsumatotal, "###########0.00")))
    
    
        Close #1
        
        frmDatosCorreoVitamedica.strCorreoDestinatario = strCorreoDestinatario
        frmDatosCorreoVitamedica.strAsunto = Trim(cboComanda.Text) & "_" & Trim(Paterno) & "_" & Trim(Materno) & "_" & Trim(Nombre) & ".txt"
        frmDatosCorreoVitamedica.strMensaje = "Se envía archivo de texto del paciente."
        frmDatosCorreoVitamedica.strRutaTXT = strRutaTXT
        frmDatosCorreoVitamedica.strNombreArchivoTXT = strNombreArchivoTXT
        frmDatosCorreoVitamedica.blnArchivoTXT = True
        frmDatosCorreoVitamedica.Show vbModal, Me
        
        If frmDatosCorreoVitamedica.blnArchivoTXT = True Then 'Generación y envío del archivo TXT exitosa; entonces se guarda el archivo
            pAltaTxtVitamedica
            pCancelar
            txtMovimientoPaciente.Text = ""
            pLimpiaGrid grdCargos
            txtMovimientoPaciente.SetFocus
        End If
    Else
        'No se han seleccionado cargos para enviar.
        MsgBox SIHOMsg(1581), vbOKOnly + vbInformation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Close #1
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    Me.MousePointer = 0
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pExportaTXT"))
End Sub

Public Sub pMarca(lstrCaracter As String, llngRenglon As Long)
    Dim llngContador As Long
    
    With grdCargos
        
        If llngRenglon > 0 Then
            .TextMatrix(llngRenglon, 0) = IIf(.TextMatrix(llngRenglon, 0) = lstrCaracter, "", lstrCaracter)
            
            If Trim(.TextMatrix(llngRenglon, 0)) = "" Then
                llngTotalSel = llngTotalSel - 1
            Else
                llngTotalSel = llngTotalSel + 1
            End If
            
            .Col = 0
            .Row = llngRenglon
            .CellFontBold = vbBlackness
        Else
            'Todos o Invertir selección
            If llngRenglon = -1 Then
                For llngContador = 1 To .Rows - 1
                    .TextMatrix(llngContador, 0) = IIf(.TextMatrix(llngContador, 0) = lstrCaracter, "", lstrCaracter)
                    
                    If Trim(.TextMatrix(llngContador, 0)) = "" Then
                        llngTotalSel = llngTotalSel - 1
                    Else
                        llngTotalSel = llngTotalSel + 1
                    End If
                    
                    .Col = 0
                    .Row = llngContador
                    .CellFontBold = vbBlackness
                Next llngContador
            End If
        End If
        .Row = llngRenglon
    End With
End Sub

Public Sub pAltaTxtVitamedica()

    Dim rsTXT As New ADODB.Recordset
    Dim vlstrTXT As String
    Dim rsPreautorizacion As New ADODB.Recordset
    Dim vlstrPreautorizacion As String
    On Error GoTo NotificaError
    
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    
    
    '-------------------------------------------------------------------------
    '  Guarda en PVTXTVITAMEDICA
    '-------------------------------------------------------------------------
    
    vlstrTXT = "select * from PVTXTVITAMEDICA where pvtxtvitamedica.numnumcuenta = " & Trim(txtMovimientoPaciente.Text)
    Set rsTXT = frsRegresaRs(vlstrTXT, adLockOptimistic, adOpenDynamic)
    
    With rsTXT
    
        .AddNew
        !numNumCuenta = Trim(txtMovimientoPaciente.Text)
        !CHRTIPOPACIENTE = IIf(OptTipoPaciente(0).Value, "I", "E")
        !dtmfechatxt = Format(Now, "dd/mmm/yyyy HH:mm")
        !INTCVEEMPRESAPACIENTE = vgintEmpresa
        !chrcomanda = cboComanda.Text
        !intCveEmpleado = vllngPersonaGraba
        .Update
        
    End With
    rsTXT.Close
    
    
    '-------------------------------------------------------------------------
    '  Guarda en PVAUTORIZAVITAMEDICA
    '-------------------------------------------------------------------------
    
    vlstrPreautorizacion = "select * from PVAUTORIZAVITAMEDICA where pvautorizavitamedica.numnumcuenta = " & Trim(txtMovimientoPaciente.Text)
    Set rsPreautorizacion = frsRegresaRs(vlstrPreautorizacion, adLockOptimistic, adOpenDynamic)
    
    With rsPreautorizacion
        
        If .RecordCount = 0 Then .AddNew
        !numNumCuenta = Trim(txtMovimientoPaciente.Text)
        !CHRTIPOPACIENTE = IIf(OptTipoPaciente(0).Value, "I", "E")
        !vchnumpreautorizacion = txtPreAutorizacion
        .Update
        
    End With
    rsPreautorizacion.Close
    
    '¡Los datos han sido guardados satisfactoriamente!
    MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAltaTxtVitamedica"))
End Sub

Public Sub pCancelarConsultaTXT()
    On Error GoTo NotificaError
    
    FreDetalleTXT.Enabled = True
    FreDetalleTXT.Visible = True
    FreFiltros.Enabled = True
    
    vgstrEstadoManto = ""
    txtNombrePaciente.Text = ""
    txtCuentaPaciente = ""
    txtCuentaPaciente.Locked = False
    
    
    'mskFechaInicial.Text = fdtmServerFecha
    'mskFechaFinal.Text = fdtmServerFecha
    mskFechaInicial.Enabled = True
    mskFechaFinal.Enabled = True
    optPaciente(2).Enabled = True
    optPaciente(3).Enabled = True
    cmdCargar.Enabled = True
        
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCancelarConsultaTXT"))
    Unload Me
End Sub

Public Sub pConfiguraGridTXT()
    On Error GoTo NotificaError
    
    With grdTXT
        .Cols = 10
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Fecha|Número cuenta |Nombre paciente|Empresa |Tipo de comanda|Persona envió"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 1550 'Fecha
        .ColWidth(2) = 1200 'Número cuenta
        .ColWidth(3) = 2800 'Nombre paciente
        .ColWidth(4) = 4800 'Empresa
        .ColWidth(5) = 1350 'Tipo de comanda
        .ColWidth(6) = 2800 'Persona envió
        
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
         
        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignLeftBottom
        .ColAlignment(4) = flexAlignLeftBottom
        .ColAlignment(5) = flexAlignLeftBottom
        .ColAlignment(6) = flexAlignLeftBottom
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
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridTXT"))
    Unload Me
End Sub

Public Sub pCorreoElectronicoDestinatario()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vlstrvchvalor As String


    vlstrvchvalor = " "
    
    vlstrSentencia = "select vchvalor from siparametro where vchnombre = 'VCHCORREOVITAMEDICA'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount <> 0 Then vlstrvchvalor = IIf(IsNull(rs!vchvalor), "", rs!vchvalor)
    rs.Close
    
    strCorreoDestinatario = vlstrvchvalor
    
    
End Sub

Public Sub pCargaSeleccionDeFacturas()
    Dim i As Integer
    Dim X As Long
    
    grdCargos.Redraw = False
    
    X = 0
    For X = 1 To grdCargos.Rows - 1
        If Trim(grdCargos.TextMatrix(X, 0)) = "*" Then
            grdCargos.TextMatrix(X, 0) = ""
        End If
    Next X
    
    i = 0
    For i = 0 To lstFacturas.ListCount - 1
        For X = 1 To grdCargos.Rows - 1
            If lstFacturas.List(i) = "<SIN FACTURA>" And lstFacturas.Selected(i) = True And grdCargos.TextMatrix(X, 24) = "" Then
                grdCargos.TextMatrix(X, 0) = "*"
            End If
            
            If lstFacturas.List(i) = grdCargos.TextMatrix(X, 24) And lstFacturas.Selected(i) = True Then
                grdCargos.TextMatrix(X, 0) = "*"
            End If
        Next X
    Next
    
    grdCargos.Redraw = True
End Sub
Function FormatearDecimalesConMiles(Valor As Double) As String
  ' Validar si el valor es entero
  If Int(Valor) = Valor Then
    ' Si es entero, mostrar 12.00
    FormatearDecimalesConMiles = Format(Valor, "00,00")
  Else
    ' Si no es entero, mostrar solo 2 decimales sin redondeo y con formato de miles
    FormatearDecimalesConMiles = Format(Valor, "#,0.00##")
  End If
End Function
