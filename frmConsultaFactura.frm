VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConsultaFactura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de facturas"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtObservacionesC 
      Height          =   2275
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      ToolTipText     =   "Observaciones de la factura"
      Top             =   6600
      Width           =   5175
   End
   Begin VB.Frame Frame11 
      Enabled         =   0   'False
      Height          =   3915
      Left            =   5280
      TabIndex        =   46
      Top             =   6240
      Width           =   3600
      Begin VB.TextBox txtRetencionServ 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   69
         Top             =   3120
         Width           =   1350
      End
      Begin VB.TextBox txtIVA 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   55
         ToolTipText     =   "Iva del presupuesto"
         Top             =   1680
         Width           =   1350
      End
      Begin VB.TextBox txtDescuentos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   54
         ToolTipText     =   "Total de descuentos"
         Top             =   240
         Width           =   1350
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   53
         ToolTipText     =   "Subtotal del presupuesto"
         Top             =   1320
         Width           =   1350
      End
      Begin VB.TextBox txtTotalFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   52
         ToolTipText     =   "Total del presupuesto"
         Top             =   2040
         Width           =   1350
      End
      Begin VB.TextBox txtPagos 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   51
         Top             =   2400
         Width           =   1350
      End
      Begin VB.TextBox txtNotas 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   50
         Top             =   2760
         Width           =   1350
      End
      Begin VB.TextBox txtTotalPagado 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   49
         Top             =   3480
         Width           =   1350
      End
      Begin VB.TextBox txtIEPS 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   48
         ToolTipText     =   "IEPS del presupuesto"
         Top             =   960
         Width           =   1350
      End
      Begin VB.TextBox txtDescuentoEspecial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1605
         TabIndex        =   47
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label lblRetencionServ 
         Caption         =   "Retención servicios"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
         Height          =   195
         Left            =   90
         TabIndex        =   65
         Top             =   1740
         Width           =   255
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
         Height          =   195
         Left            =   90
         TabIndex        =   64
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal"
         Height          =   195
         Left            =   90
         TabIndex        =   63
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total factura"
         Height          =   195
         Left            =   90
         TabIndex        =   62
         Top             =   2100
         Width           =   900
      End
      Begin VB.Label lblMoneda 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2960
         TabIndex        =   61
         Top             =   2058
         Width           =   615
      End
      Begin VB.Label lbPagos 
         Caption         =   "Pagos"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label lbNotas 
         Caption         =   "Nota de crédito"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Label lbTotalPagado 
         Caption         =   "Total pagado"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   3540
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "IEPS"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Descuento especial"
         Height          =   195
         Left            =   90
         TabIndex        =   56
         Top             =   660
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   780
      TabIndex        =   38
      Top             =   10160
      Width           =   7335
      Begin VB.CommandButton cmdAplAnt 
         Caption         =   "Anticipos CFDi"
         Height          =   480
         Left            =   5715
         TabIndex        =   45
         ToolTipText     =   "Comprobante por aplicación de anticipos"
         Top             =   165
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancelarFactura 
         Height          =   480
         Left            =   1050
         Picture         =   "frmConsultaFactura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancelar factura"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdCFD 
         Height          =   480
         Left            =   3260
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaFactura.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Comprobante fiscal digital"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdRefacturacion 
         Enabled         =   0   'False
         Height          =   480
         Left            =   1545
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaFactura.frx":0E10
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Refacturación"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdControlAseguradora 
         Caption         =   "Control de aseguradora"
         Height          =   480
         Left            =   3740
         TabIndex        =   44
         Top             =   165
         Width           =   1980
      End
      Begin VB.CommandButton cmdSiguienteFactura 
         Height          =   480
         Index           =   0
         Left            =   555
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaFactura.frx":139A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Siguiente registro"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAnterior 
         Height          =   480
         Index           =   0
         Left            =   65
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaFactura.frx":150C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Anterior registro"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdconfirmatimbre 
         Caption         =   "Confirmar timbre fiscal"
         Height          =   480
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConsultaFactura.frx":167E
         TabIndex        =   42
         ToolTipText     =   "Confirmar timbre fiscal"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   90
      TabIndex        =   30
      Top             =   10920
      Visible         =   0   'False
      Width           =   8760
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   180
         Width           =   8610
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Left            =   30
         Top             =   120
         Width           =   8700
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1650
      Left            =   60
      TabIndex        =   27
      Top             =   6600
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label lblFoliosRelacionados 
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   285
         Visible         =   0   'False
         Width           =   4920
      End
   End
   Begin VB.Frame Frame3 
      Height          =   930
      Left            =   60
      TabIndex        =   16
      Top             =   9225
      Width           =   5175
      Begin VB.Label lblEmpleadoCancela 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   720
         TabIndex        =   20
         Top             =   540
         Width           =   4365
      End
      Begin VB.Label lblEmpleadoFactura 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   210
         Width           =   4365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Facturó"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblEmpleadoCancelo 
         AutoSize        =   -1  'True
         Caption         =   "Canceló"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   600
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3240
      Left            =   60
      TabIndex        =   1
      Top             =   3000
      Width           =   8820
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConsultaFactura 
         Height          =   3060
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   5398
         _Version        =   393216
         GridColor       =   12632256
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame freDatos 
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      Height          =   2970
      Left            =   60
      TabIndex        =   0
      Top             =   20
      Width           =   8820
      Begin VB.TextBox txtEstadoFactura 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   68
         Text            =   "<EstadoFactura>"
         Top             =   270
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtRFCFactura 
         Height          =   315
         Left            =   1560
         TabIndex        =   67
         Top             =   975
         Width           =   1620
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   0
         Left            =   4410
         TabIndex        =   37
         Top             =   270
         Width           =   795
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   1
         Left            =   5340
         TabIndex        =   36
         Top             =   270
         Width           =   840
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   2
         Left            =   6315
         TabIndex        =   35
         Top             =   270
         Width           =   735
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Venta al público"
         Height          =   255
         Index           =   3
         Left            =   7185
         TabIndex        =   34
         Top             =   270
         Width           =   1485
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Left            =   5570
         TabIndex        =   33
         Top             =   600
         Width           =   3105
      End
      Begin VB.TextBox txtCiudad 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   26
         Top             =   2445
         Width           =   4200
      End
      Begin VB.TextBox txtColonia 
         Height          =   315
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   23
         Top             =   2070
         Width           =   4200
      End
      Begin VB.TextBox txtCP 
         Height          =   315
         Left            =   7110
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2070
         Width           =   1560
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   270
         Width           =   885
      End
      Begin VB.TextBox txtFolio 
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   615
         Width           =   1620
      End
      Begin VB.TextBox txtCanceladada 
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
         Left            =   4365
         TabIndex        =   11
         Text            =   "Pendiente de timbre fiscal"
         Top             =   960
         Width           =   4305
      End
      Begin VB.TextBox txtTelefonoFactura 
         Height          =   315
         Left            =   7110
         TabIndex        =   5
         Top             =   2430
         Width           =   1560
      End
      Begin VB.TextBox txtDireccionFactura 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1695
         Width           =   7115
      End
      Begin VB.TextBox txtNombreFactura 
         Height          =   315
         Left            =   1560
         MaxLength       =   300
         TabIndex        =   3
         Top             =   1335
         Width           =   7115
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Código postal"
         Height          =   195
         Left            =   5895
         TabIndex        =   25
         Top             =   2130
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Colonia"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   2130
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ciudad"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2490
         Width           =   495
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Folio"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   675
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   5025
         TabIndex        =   10
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "RFC"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1035
         Width           =   435
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   5895
         TabIndex        =   8
         Top             =   2490
         Width           =   630
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         Caption         =   "Dirección"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1755
         Width           =   675
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "Razón social"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1395
         Width           =   915
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   6330
      Width           =   1575
   End
End
Attribute VB_Name = "frmConsultaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmConsultaFactura                                     -
'-------------------------------------------------------------------------------------
'| Objetivo: Consultar y cancelar las facturas                                       -
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.                                       -
'| Autor                    : Rodolfo Ramos G.                                       -
'| Fecha de Creación        : 26/Jun/2001                                            -
'| Modificó                 : Nombre(s)                                              -
'| Fecha Terminación        : HOY                                                    -
'| Fecha última modificación: 03/Oct/2002                                            -
'-------------------------------------------------------------------------------------

Option Explicit

Dim vgstrFolioFactura As String
Public vgfrmFacturas As Form                'Cual es la forma que la está llamando
Private vglngCveFactura As Long             'Consecutivo de la factura que se va a Cancelar
Dim vgstrTipoFactura As String              'El tipo de factura podría ser "N"=Normal "T"=Tikets "D"=DirectaPOS
Dim vglngMovPaciente As Long                'Cual paciente es...
Dim vgstrTipoPaciente As String             'El Tipo de Paciente
Dim vgstrFacturaPacienteEmpresa As String   'Si la factura es de la empresa o del paciente
Dim dtmFechaMinima As Date                  'Fecha minima para fecha del ingreso
Dim dtmfechaingreso As Date                 'Fecha del ingreso original de la factura
Dim blnValidarFacturaEmpresa As Integer
Dim vglngCveEmpresaConvenio As Long
Dim strFechaFactura As String
Dim lngConsecutivoFactura As Long
Dim lblnCalcularEnBaseACargos As Boolean
Dim lblnSinComprobante As Boolean
Dim ldblHonorariosFacturados As Double
Dim vlstrTipoCFD As String
Public vgstrEquipo As String
Public vgstrIP As String
Public strRequestXML As String
Public strResponseXML As String
Dim vlblnLicenciaIEPS As Boolean
Dim vldblTipoCambio As Double
Dim rsTemp As New ADODB.Recordset
Dim rsEsPagoporCancelacion As New ADODB.Recordset
Dim vllngAplAnt As Long
Dim vlblnAplAntPend As Boolean
Dim vllngConsecutivoFactura As Long         'Consecutivo PvFactura
Dim vldblRetServicios As Double
'Dim vlblnMotivos As Boolean ''*
Dim vllngCveCarta As Long                   'Clave de la carta de la factura



Public Sub pConsultaFacturas(vlstrFolioFactura As String, vlbolMuestraForma As Boolean)
    '-----------------------------------------------------------------------
    ' Este procedure nos muestra una factura.
    '-----------------------------------------------------------------------
    Dim vlstrSentencia As String
    Dim vlstrConceptoExcedenteSumaAsegurada As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsFactura As New ADODB.Recordset
    Dim rsCiudad As New ADODB.Recordset
    Dim rsFoiosRelacionados As ADODB.Recordset
    Dim strFolios As String
    Dim rsAplAnt As ADODB.Recordset
    Dim rsAplAntPend As ADODB.Recordset

    'Para obtener el Concepto del EXCEDENTE de los PARAMETROS
    vlstrSentencia = "SELECT chrDescripcion ConceptoSumaAsegurada " & _
                     " FROM PvConceptoFacturacion " & _
                     " WHERE smiCveConcepto = " & _
                     " (SELECT intConceptoSumaAsegurada FROM PvParametro WHERE tnyclaveempresa = " & vgintClaveEmpresaContable & " ) "
    Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If Not rsTemp.EOF Then
        vlstrConceptoExcedenteSumaAsegurada = rsTemp!ConceptoSumaAsegurada
    End If
    rsTemp.Close
    
    'Inicialización de las variables
    vglngMovPaciente = 0
    vglngCveEmpresaConvenio = 0
    vgstrTipoPaciente = ""
    vgstrFacturaPacienteEmpresa = ""
    txtNombreFactura = ""
    txtDireccionFactura = ""
    txtTelefonoFactura = ""
    txtRFCFactura = ""
    txtCiudad.Text = ""
    txtCP.Text = ""
    txtColonia.Text = ""
    lblEmpleadoCancela.Caption = ""
    lblEmpleadoFactura.Caption = ""
    blnValidarFacturaEmpresa = 0
    pLimpiaGrid grdConsultaFactura
    pConfiguraGridFactura
    lblnSinComprobante = False
    
    Set rsFactura = frsEjecuta_SP(vlstrFolioFactura, "Sp_PvSelFactura_NE")
    If rsFactura.RecordCount <> 0 Then
        With rsFactura
            '--------------------------'
            ' Encabezado de la factura '
            '--------------------------'
            txtNombreFactura = IIf(IsNull(!RazonSocial), "", !RazonSocial)
            txtDireccionFactura = IIf(IsNull(!DireccionCompleta), "", !DireccionCompleta)
            txtTelefonoFactura = IIf(IsNull(!Telefono), "", !Telefono)
            txtRFCFactura = IIf(IsNull(!RFC), "", !RFC)
            txtColonia.Text = IIf(IsNull(!Colonia), "", !Colonia)
            txtCP.Text = IIf(IsNull(!CP), "", !CP)
            'Carga la ciudad del domicilio fiscal
            vgstrParametrosSP = !cveCiudad & "|-1|-1"
            Set rsCiudad = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELCIUDAD")
            If rsCiudad.RecordCount <> 0 Then
                txtCiudad.Text = IIf(IsNull(rsCiudad!VCHDESCRIPCION), "", rsCiudad!VCHDESCRIPCION)
            End If
    
            lblEmpleadoCancela.Caption = Trim(IIf(IsNull(!PersonaCancelo), "", !PersonaCancelo))
            lblEmpleadoCancela.Visible = !chrEstatus = "C"
            lblEmpleadoCancelo.Visible = !chrEstatus = "C"
            lblEmpleadoFactura.Caption = !PersonaFacturo
            
            lblMoneda.Caption = IIf(!BITPESOS = 1, "Pesos", "Dólares")
            vldblTipoCambio = !TipoCambio
            vglngCveFactura = !IdFactura
            
            lngConsecutivoFactura = !IdFactura
            vllngConsecutivoFactura = !IdFactura
            
            ldblHonorariosFacturados = IIf(IsNull(!mnyHonorariosFacturados), 0, !mnyHonorariosFacturados)
            
            ' Esto es que una factura puede ser :
            ' "Normal" = Que no se hizo en el POS
            ' "Ticket" = Que se facturo uno o varios tickets
            ' "Directo" = Que fue Venta al publico pero con Factura, sin ticket
            vgstrTipoFactura = IIf(!intCveVentaPublico = 0, "N", IIf(!intCveVentaPublico = -1, "T", "D"))
            
            grdConsultaFactura.Row = 1
            Do While Not .EOF
                If grdConsultaFactura.RowData(1) <> -1 Then
                   grdConsultaFactura.Rows = grdConsultaFactura.Rows + 1
                End If
                grdConsultaFactura.Row = grdConsultaFactura.Rows - 1
                grdConsultaFactura.RowData(grdConsultaFactura.Row) = 1
                grdConsultaFactura.TextMatrix(grdConsultaFactura.Row, 1) = IIf(IsNull(!Concepto), "", !Concepto)
                If !chrTipo = "NO" Or !chrTipo = "OC" Then
                    grdConsultaFactura.TextMatrix(grdConsultaFactura.Row, 2) = Format(!Importe, "$ ###,###,###,###.00")
                Else
                    grdConsultaFactura.TextMatrix(grdConsultaFactura.Row, 3) = Format(!Importe, "$ ###,###,###,###.00")
                End If
                .MoveNext
            Loop
            .MoveFirst
            
            txtSubtotal.Text = Format(!Subtotal, "$ ###,###,###,###.00")
            txtDescuentos.Text = Format((!DescuentoFactura - !MNYDESCUENTOESPECIAL), "$ ###,###,###,###.00")
            txtDescuentoEspecial.Text = Format(!MNYDESCUENTOESPECIAL, "$ ###,###,###,###.00")
            txtIVA.Text = Format(!IVAFactura, "$ ###,###,###,###.00")
            txtIEPS.Text = Format(IIf(vlblnLicenciaIEPS, !IEPS, 0), "$ ###,###,###,###.00")
            txtRetencionServ.Text = Format(!MNYRETENSERVICIOS, "$ ###,###,###,###.00")
            txtTotalFactura.Text = Format(!TotalFactura + !MNYRETENSERVICIOS, "$ ###,###,###,###.00")
            txtFolio.Text = vlstrFolioFactura
            txtFecha.Text = Format(!fecha, "Long Date")
            strFechaFactura = Format(!fecha, "DD/MM/YYYY")
            txtEstadoFactura.Text = IIf(IsNull(!PendienteCancelarSAT_NE), "NP", !PendienteCancelarSAT_NE)
            TxtObservacionesC.Text = Trim(IIf(IsNull(!Observaciones), "", !Observaciones))
            
            If !PendienteTimbre = 1 Then
               txtCanceladada.Text = "Pendiente de timbre fiscal"
               txtCanceladada.Width = 3105
               txtCanceladada.Left = 5565
               'txtCanceladada.Font.Size = 9.75
               txtCanceladada.ForeColor = &H0&       'negro
               txtCanceladada.BackColor = &HFFFF&     'amarillo
               txtCanceladada.Visible = True
            Else
               If !chrEstatus = "C" Then
                  If !PendienteCancelarSat = 1 Then
                     txtCanceladada.Text = "Pendiente de cancelarse ante el SAT"
                     'txtCanceladada.Font.Size = 9
                     txtCanceladada.Width = 4305
                     txtCanceladada.Left = 4365
                     txtCanceladada.ForeColor = &HFF& 'rojo
                     txtCanceladada.BackColor = &HC0E0FF ' naranja suave
                     txtCanceladada.Visible = True
                  Else
                     txtCanceladada.Width = 3105
                     txtCanceladada.Left = 5565
                     'txtCanceladada.Font.Size = 9.75
                     txtCanceladada.Text = "Factura cancelada"
                     txtCanceladada.ForeColor = &HFF& 'rojo
                     txtCanceladada.BackColor = &H80000005 'blanco
                     txtCanceladada.Visible = True
                  End If
                  lblnSinComprobante = True
               Else
                    Select Case !PendienteCancelarSAT_NE
                        Case "PA"
                            txtCanceladada.Text = "Pendiente de autorización"
                            'txtCanceladada.Font.Size = 9
                            txtCanceladada.Width = 4305
                            txtCanceladada.Left = 4365
                            txtCanceladada.ForeColor = &H80000005 '| Blanco
                            txtCanceladada.BackColor = &H80FF&    '| Naranja fuerte
                            txtCanceladada.Visible = True
                        Case "CR"
                            txtCanceladada.Text = "Cancelación rechazada"
                            'txtCanceladada.Font.Size = 9
                            txtCanceladada.Width = 4305
                            txtCanceladada.Left = 4365
                            txtCanceladada.ForeColor = &H80000005 '| Blanco
                            txtCanceladada.BackColor = &HFF&      '| Rojo
                            txtCanceladada.Visible = True
                        Case Else '| NP o cualquier otro
                            txtCanceladada.Text = ""
                            txtCanceladada.Visible = False
                            txtCanceladada.ForeColor = &H0&       'negro
                            txtCanceladada.BackColor = &H80000005 'blanco
                    End Select
               End If
            End If
            
            If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3088, 4114), "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3088, 4114), "C", True) Then
                cmdCancelarFactura.Enabled = !chrEstatus <> "C" And !PendienteTimbre = 0 '**-**-**' Valida permiso para cancelar facturas
            Else
                cmdCancelarFactura.Enabled = False
            End If
            
            Dim rsrefact As New ADODB.Recordset
            Dim paramrefact As String
            Set rsrefact = frsRegresaRs("select vchvalor from siparametro where trim(vchnombre) = 'BITREFACTURACIONACTIVA'")
            paramrefact = 0
            If rsrefact.RecordCount > 0 Then
               paramrefact = rsrefact!vchvalor
            Else
                paramrefact = 0
            End If
            
            If (fblnRevisaPermiso(vglngNumeroLogin, 3091, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3091, "C", True)) And paramrefact = "1" Then
                cmdRefacturacion.Enabled = True '| Se deshabilitó hasta que se defina como va a funcionar el proceso con el nuevo esquema de facturación !chrEstatus <> "C" And !PendienteTimbre = 0
            Else
                cmdRefacturacion.Enabled = False
            End If
            cmdconfirmatimbre.Enabled = !PendienteTimbre = 1
            cmdCFD.Enabled = !PendienteTimbre = 0 'solo cuando este pendiente de timbre quedará deshabilidata
            'vlblnMotivos = !PendienteTimbre = 0 ''*
                
            lblCuenta.Caption = IIf(Trim(!CHRTIPOPACIENTE) = "G", "Clave", "Cuenta")
            optTipoPaciente(0).Value = Trim(!CHRTIPOPACIENTE) = "I"
            optTipoPaciente(1).Value = Trim(!CHRTIPOPACIENTE) = "E"
            optTipoPaciente(2).Value = Trim(!CHRTIPOPACIENTE) = "G"
            optTipoPaciente(3).Value = Trim(!CHRTIPOPACIENTE) = "V"
            
            txtMovimientoPaciente.Text = !cuenta
            vglngMovPaciente = !cuenta
            vgstrTipoPaciente = !CHRTIPOPACIENTE
            vgstrFacturaPacienteEmpresa = !chrTipoFactura
            vllngCveCarta = IIf(IsNull(!intCveCarta), 0, !intCveCarta)
            
            strFolios = ""
            Set rsFoiosRelacionados = frsRegresaRs("select * from PVFacturaFolios where smiCveDepartamento = " & rsFactura!CveDepartamento & " and TRIM(chrFolioFactura) = '" & vlstrFolioFactura & "'")
            If Not rsFoiosRelacionados.EOF Then
                strFolios = "Folios relacionados: "
                Do Until rsFoiosRelacionados.EOF
                    strFolios = strFolios & Trim(rsFoiosRelacionados!chrFolioRelacionado) & " "
                    rsFoiosRelacionados.MoveNext
                Loop
            End If
            rsFoiosRelacionados.Close
            cmdAplAnt.Enabled = False
            vlblnAplAntPend = False
            If !PendienteTimbre = 0 Then
                Set rsAplAnt = frsRegresaRs("select * from PVAplicacionAnticipo where chrFolioFactura = '" & Trim(txtFolio.Text) & "'", adLockReadOnly, adOpenForwardOnly)
                If Not rsAplAnt.EOF Then
                    vllngAplAnt = rsAplAnt!INTCOMPROBANTE
                    Set rsAplAntPend = frsRegresaRs("select count(*) from GNPendientesTimbreFiscal where chrTipoComprobante = 'AA' and intComprobante = " & vllngAplAnt, adLockReadOnly, adOpenForwardOnly)
                    If rsAplAntPend.Fields(0).Value > 0 Then
                        cmdAplAnt.Caption = "Anticipos confirmar timbre"
                        vlblnAplAntPend = True
                    Else
                        cmdAplAnt.Caption = "Anticipos CFDi"
                    End If
                    rsAplAntPend.Close
                    cmdAplAnt.Enabled = True
                Else
                    vllngAplAnt = 0
                End If
                rsAplAnt.Close
            End If

            
            
        End With
        lblFoliosRelacionados.Caption = strFolios
                
    End If
    
    '---------------------------------------------------------------------'
    ' Para saber si habilito el botón para ver el control de aseguradoras '
    '---------------------------------------------------------------------'
    cmdControlAseguradora.Enabled = False
    
    vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & Str(vgintClaveEmpresaContable)
    
    'Internos ó Externos
    Set rsTemp = IIf(optTipoPaciente(0).Value, frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA"), frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA"))
    If rsTemp.RecordCount > 0 Then
        If rsTemp!bitUtilizaConvenio = 1 And rsTemp!bitAseguradora = 1 Then
            cmdControlAseguradora.Enabled = True
            blnValidarFacturaEmpresa = IIf(vgstrFacturaPacienteEmpresa = "P", 1, 0)
            vglngCveEmpresaConvenio = rsTemp!cveEmpresa
        End If
    End If
    rsTemp.Close
    
    pLlenaPagos
    
    vlstrTipoCFD = ""
    'vlblnMotivos = False ''*
    cmdCFD.Enabled = fblnCFD
    If cmdCFD.Enabled = False And lblnSinComprobante = True Then
        cmdCFD.Enabled = True
        'vlblnMotivos = True ''*
    End If
    
    pPreparaIEPS
    If vlbolMuestraForma Then frmConsultaFactura.Show vbModal
End Sub

Private Sub pConfiguraGridFactura()
    With grdConsultaFactura
        .Cols = 10
        .FixedCols = 2
        .FixedRows = 1
        .FormatString = "|Concepto|Cargo|Abono"
        .ColWidth(0) = 200  'Fix
        .ColWidth(1) = 4500 'Concepto de facturación
        .ColWidth(2) = 1430 'Cargo
        .ColWidth(3) = 1430 'Abono
        .ColWidth(4) = 0    'IVA
        .ColWidth(5) = 0    'Descuentos si la factura es consolidada
        .ColWidth(6) = 0    'Disponible
        .ColWidth(7) = 0    'Disponible
        .ColWidth(8) = 0    'Disponible
        .ColWidth(9) = 0    'Disponible
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignLeftCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        .ScrollBars = flexScrollBarBoth
    End With
End Sub

Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
End Sub

Private Sub cmdAnterior_Click(Index As Integer)
    If vgfrmFacturas.grdBuscaFacturas.Row > 1 Then
        vgfrmFacturas.grdBuscaFacturas.Row = vgfrmFacturas.grdBuscaFacturas.Row - 1
        pConsultaFacturas vgfrmFacturas.grdBuscaFacturas.TextMatrix(vgfrmFacturas.grdBuscaFacturas.Row, 1), False
    End If
End Sub


Private Sub cmdAplAnt_Click()
    Dim lngCveFormato As Long
    Dim blnResutado As Boolean
    If vlblnAplAntPend Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        pLogTimbrado 2
        
        If Not fblnGeneraComprobanteDigital(vllngAplAnt, "AA", 0, 0, "", True, False) Then
              On Error Resume Next
              If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
                 'El donativo se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
                 MsgBox SIHOMsg(1306), vbInformation + vbOKOnly, "Mensaje"
                 
                 pLogTimbrado 1
                 
                 EntornoSIHO.ConeccionSIHO.CommitTrans
              ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then  'No se realizó el timbrado
                EntornoSIHO.ConeccionSIHO.RollbackTrans
                
                pLogTimbrado 1
              End If
        Else
          blnResutado = True
          pEliminaPendientesTimbre vllngAplAnt, "AA"
          vlblnAplAntPend = False
          cmdAplAnt.Caption = "Anticipos CFDi"
          
          pLogTimbrado 1
          
          EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
        lngCveFormato = 1
        frsEjecuta_SP vgintNumeroDepartamento & "|0|0|T", "fn_PVSelFormatoFactura2", True, lngCveFormato
        fblnImprimeComprobanteDigital vllngAplAnt, "AA", "I", lngCveFormato, 0
        If blnResutado Then
            If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
                '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    pEnviarCFD "AA", vllngAplAnt, CLng(vgintClaveEmpresaContable), Trim(txtRFCFactura.Text), vglngNumeroEmpleado, Me
                End If
            End If
        End If

    Else
        frmComprobanteFiscalDigitalInternet.lngComprobante = vllngAplAnt
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = "AA"
        frmComprobanteFiscalDigitalInternet.blnCancelado = txtCanceladada.Visible
        frmComprobanteFiscalDigitalInternet.blnFacturaSinComprobante = False
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    End If
End Sub

Private Sub cmdCancelarFactura_Click()
    
    '******************************************************************************************************
    'CUALQUIER CAMBIO EN ESTE PROCEDIMIENTO DEBE TAMBIEN AFECTAR A PCANCELARFACTURA EN EL MODPROCEDIMIENTOS
    '******************************************************************************************************
    '-------------------------------'
    ' Sólo cancelación de Pacientes '
    '-------------------------------'

    Dim vllngPersonaGraba As Long
    
    If fblnFacturaCancelable(Trim(txtFolio.Text)) Then
        '-------------------'
        ' Persona que graba '
        '-------------------'
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
                frmMotivosCancelacion.blnActivaUUID = False ''*
                frmMotivosCancelacion.Show vbModal, Me ''*
                If vgMotivoCancelacion = "" Then Exit Sub ''*
        pCancelaCFDiFacturaSiHO Trim(txtFolio.Text), txtEstadoFactura.Text, vllngPersonaGraba, ldblHonorariosFacturados, Me.Name

        pConsultaFacturas Trim(Trim(txtFolio.Text)), 0
    End If
End Sub

Private Sub pLlenaPagos()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim vldblTotalPagos As Double
    Dim vldblCantidad As Double
    Dim strParametros As String
    Dim rsCantidadNota As ADODB.Recordset
    Dim dblCantidad As Double
    Dim rsTipoCambio As ADODB.Recordset
    Dim intTipoCambio As Integer
    
On Error GoTo NotificaError
    
    vlstrSentencia = "SELECT intNumPago " & _
                     ", pvpago.intNumConcepto " & _
                     ", chrDescripcion Concepto " & _
                     ", dtmFecha Fecha" & _
                     ", chrFolioRecibo Recibo" & _
                     ", mnyCantidad Cantidad " & _
                     ", CASE bitPesos WHEN 1 THEN 'Pesos' ELSE 'Dolares' END AS Moneda " & _
                     ", mnyTipoCambio TipoCambio" & _
                     ", chrTipo TipoPago " & _
                     ", isnull(chrFolioFactura,'') Factura" & _
                     ", 'E' EntradaSalida " & _
                     ", intNumCorte Corte " & _
                     "  FROM PvPago " & _
                     "  INNER JOIN PvConceptoPago ON PvPago.intNumConcepto = PvConceptoPago.intNumConcepto " & _
                     "  WHERE chrFolioFactura = '" & Trim(txtFolio) & "'"
    If frmFacturacion.optGrupoCuenta(0).Value Then
        vlstrSentencia = vlstrSentencia & _
                         " AND (PVPAGO.intmovpaciente, PVPAGO.chrtipopaciente) " & _
                         "       IN (SELECT Distinct INTMOVPACIENTE, CHRTIPOPACIENTE FROM PVDETALLEFACTURACONSOLID WHERE INTCVEGRUPO = " & Trim(txtMovimientoPaciente.Text) & ")"
    Else
        vlstrSentencia = vlstrSentencia & _
                         " AND chrTipoPaciente = " & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & _
                         " AND intMovPaciente = " & Trim(txtMovimientoPaciente.Text)
    End If
    
    vlstrSentencia = vlstrSentencia & " UNION SELECT intNumSalida, PvSalidaDinero.intNumConcepto, " & _
                    " chrDescripcion Concepto, dtmFecha Fecha, chrFolioRecibo Recibo, " & _
                    " mnyCantidad*-1 Cantidad, " & _
                    " CASE bitPesos WHEN 1 THEN 'Pesos' ELSE 'Dolares' END AS Moneda, " & _
                    " mnyTipoCambio TipoCambio, 'SD' TipoPago, " & _
                    " isnull(chrFolioFactura,'') Factura, " & _
                    " 'S' EntradaSalida, " & _
                    " intNumCorte Corte " & _
                    " FROM PvSalidaDinero " & _
                    " INNER JOIN PvConceptoPago ON PvSalidaDinero.intNumConcepto = PvConceptoPago.intNumConcepto " & _
                    " WHERE chrTipoPaciente = " & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & _
                    " AND intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & _
                    " AND chrFolioFactura = '" & Trim(txtFolio) & "'"
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    vldblTotalPagos = 0
    
    Do While Not rs.EOF
        vldblCantidad = IIf(rs!Moneda = "Pesos", 1, rs!TipoCambio) * rs!cantidad 'Convierte a pesos (Si es necesario)
        
        If lblMoneda.Caption = "Dólares" Then
            strParametros = rs!fecha
            Set rsTipoCambio = frsEjecuta_SP(strParametros, "Sp_GnSelTipoCambio")
            If rsTipoCambio.RecordCount > 0 Then intTipoCambio = rsTipoCambio!Venta
            rsTipoCambio.Close
        
            vldblCantidad = Round(vldblCantidad / intTipoCambio, 2)
        End If
        
        If (rs!tipoPago = "NO" Or rs!tipoPago = "SD") Then
            vldblTotalPagos = vldblTotalPagos + vldblCantidad
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
    txtPagos.Text = Format(vldblTotalPagos, "$ ###,###,###,###.00")
    
    'Regresa cantidad de la nota(s)
    strParametros = txtFolio & "|" & CStr(vgintClaveEmpresaContable) & "|" & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & "|" & txtMovimientoPaciente.Text
    Set rsCantidadNota = frsEjecuta_SP(strParametros, "SP_PVSELNOTADECREDITO")
    If rsCantidadNota.RecordCount > 0 Then
        dblCantidad = rsCantidadNota!cantidad
        
        If lblMoneda.Caption = "Dólares" Then
            strParametros = strFechaFactura
            Set rsTipoCambio = frsEjecuta_SP(strParametros, "Sp_GnSelTipoCambio")
            If rsTipoCambio.RecordCount > 0 Then intTipoCambio = rsTipoCambio!Venta
            rsTipoCambio.Close
            
            dblCantidad = Round(dblCantidad / intTipoCambio, 2)
        End If
    Else
        dblCantidad = 0
    End If
    rsCantidadNota.Close
    
    txtNotas.Text = Format(dblCantidad, "$ ###,###,###,###.00")
    txtTotalPagado = Format(Val(CDbl(txtTotalFactura) - (CDbl(txtPagos) + CDbl(txtNotas)) - CDbl(txtRetencionServ)), "$ ###,###,###,###.00")
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaPagos"))
End Sub

Private Sub pDeshabilitarFechaIngreso()
   txtFolio.Enabled = True
   txtFecha.Enabled = True
   txtNombreFactura.Enabled = True
   txtRFCFactura.Enabled = True
   txtDireccionFactura.Enabled = True
   txtTelefonoFactura.Enabled = True
   txtCP.Enabled = True
   txtColonia.Enabled = True
   txtCiudad.Enabled = True
   txtMovimientoPaciente.Enabled = True
   optTipoPaciente(0).Enabled = True
   optTipoPaciente(1).Enabled = True
   optTipoPaciente(2).Enabled = True
   optTipoPaciente(3).Enabled = True
   freDatos.Enabled = False
End Sub

Private Sub cmdCierra_Click()
    Unload Me
End Sub

Private Sub cmdconfirmatimbre_Click()
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
                                                 
       blnNOMensajeErrorPAC = True
       
       pLogTimbrado 2
       EntornoSIHO.ConeccionSIHO.BeginTrans
       vlngReg = flngRegistroFolio("FA", lngConsecutivoFactura)
       If Not fblnGeneraComprobanteDigital(lngConsecutivoFactura, "FA", 0, fintAnoAprobacion(vlngReg), fStrNumeroAprobacion(vlngReg), fblnTCFDi(vlngReg)) Then
          On Error Resume Next
          EntornoSIHO.ConeccionSIHO.RollbackTrans
          frsEjecuta_SP CFDilngNumError & "|" & Left(CFDistrDescripError, 200) & "|" & cgstrModulo & "|" & Left(CFDistrProcesoError, 50) & " Linea:" & CFDiintLineaError & "|" & "", "SP_GNINSREGISTROERRORES", True
          If vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 3 Then 'timbre pendiente de confirmar
             'Por el momento no es posible confirmar el timbre de la factura <FOLIO>, intente de nuevo en unos minutos.
              MsgBox Replace(SIHOMsg(1314), " <FOLIO>", ""), vbInformation + vbOKOnly, "Mensaje"
             'la factura se queda igual, no se hace nada
             
             pLogTimbrado 1
             
          ElseIf vgIntBanderaTImbradoPendiente = 2 Then 'No se realizó el timbrado
              pLogTimbrado 1
              'No es posible realizar el timbrado de la factura <FOLIO>, la factura será cancelada.
              MsgBox Replace(SIHOMsg(1313), " <FOLIO>", ""), vbExclamation + vbOKOnly, "Mensaje"
              'Aqui se debe de cancelar la factura
              pCancelarFactura Trim(Me.txtFolio.Text), vllngPersonaGraba, Me.Name
              'se carga de nuevo la factura
              pConsultaFacturas Trim(txtFolio.Text), 0
          End If
       Else
          'Se guarda el LOG
           Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, Me.Caption, "Confirmación de timbre factura " & Me.txtFolio.Text)
          'Eliminamos la informacion de la factura de la tabla de pendientes de timbre fiscal
           pEliminaPendientesTimbre lngConsecutivoFactura, "FA"
          'Commit
           EntornoSIHO.ConeccionSIHO.CommitTrans
           pLogTimbrado 1
          'Timbre fiscal de factura <FOLIO>: Confirmado.
           MsgBox Replace(SIHOMsg(1315), "<FOLIO> ", ""), vbInformation + vbOKOnly, "Mensaje"
           'se carga de nuevo la factura
           pConsultaFacturas Trim(txtFolio.Text), 0
       End If
       
       'Barra de progreso CFD
       pgbBarraCFD.Value = 100
       freBarraCFD.Top = 3200
       Screen.MousePointer = vbDefault
       lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
       freBarraCFD.Visible = True
       freBarraCFD.Refresh
       freBarraCFD.Visible = False
       blnNOMensajeErrorPAC = False
   
End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdConfirmarTimbre_Click"))
    'Unload Me
End Sub

Private Sub cmdControlAseguradora_Click()
    With frmConsultaControlAseguradora
        .vglngMovPaciente = txtMovimientoPaciente.Text
        .vgstrInternoExterno = IIf(optTipoPaciente(0).Value, "I", "E")
        .vgstrTipoFactura = vgstrFacturaPacienteEmpresa
        .vgstrFolioFactura = Trim(txtFolio.Text)
        .vglngCveCarta = vllngCveCarta
        .Show vbModal
    End With
End Sub

Private Sub cmdRefacturacion_Click()
'*************************************************************************************************************************************
'******** Comentado Temporalmente para compilar con el nuevo esquema de cancelación, hasta definir si se retoma esta funcionalidad
'*************************************************************************************************************************************
    Dim rsDC As New ADODB.Recordset             'Es el detalle corte pero como consulta
    Dim rsChequeTransCta As New ADODB.Recordset
    Dim rsChecaCredito As New ADODB.Recordset   'RS para saber si la factura fue a crédito y ya tiene pagos
    Dim rsPvDetalleCorte As New ADODB.Recordset 'Aqui añado los registros del detalle del corte
    Dim rsTemp As New ADODB.Recordset           'RS Temporal para lo que sea
    Dim rsCorteTiKets As New ADODB.Recordset    'RS para guardar los tikets que se estan reactivando despues de cancelar la factura
    Dim rsFPEOld As New ADODB.Recordset
    Dim rsFPENew As New ADODB.Recordset

    Dim vlstrSentencia As String                'Sirve pa TODOS los RS's
    Dim vllngNumeroCorte As Long                'Trae el numero de corte actual

    Dim vlblnCorteValido As Boolean
    Dim vllngNumCorteFactura As Long            'Es el número de corte en el que se registró la factura.
    Dim vllngCorteGrabando As Long
    Dim vllngFoliosFaltantes As Long
    Dim vlstrFolioDocumento As String
    Dim vllngCveFacturaNueva As Long
    Dim vPrinter As Printer                      'Para saber a cual impresora se manda la factura
    Dim vllngFormatoaUsar As Long                'Para saber que formato se va a utilizar
    Dim vlrsNuevaFactura As New ADODB.Recordset
    Dim vlrsAnteriorFactura As New ADODB.Recordset
    Dim vllngMensaje As Long
    Dim lngPersonaGraba As Long 'Clave del empleado que realiza la refacturación
    Dim vllngPersonaGraba As Long 'Clave del empleado que se inserta en los registros del detalle del corte

    Dim strParametros As String

    Dim vlaryParametrosSalida() As String
    Dim strTipoPacienteFactura As String

    '(PEMEX)
    Dim lngCveFormato As Long
    Dim rsDatosFactura As ADODB.Recordset
    Dim blnFacturaMultiple As Boolean
    Dim lngRenglonesDetalle As Long
    Dim lngTotalDocumentos As Long
    Dim arrFolios() As String
    Dim rsFolios As ADODB.Recordset
    Dim lngIndexFolios As Long
    Dim lng
    Dim blnCancelarFacturacion As Boolean
    Dim blnFoliosOK As Boolean
    Dim strIdentificador As String
    Dim lngInicial As Long
    Dim lngFinal As Long
    Dim rsDatosPac As ADODB.Recordset
    Dim lngCveTipoPaciente As Long
    Dim lngCveEmpresaPac As Long
    Dim lngContador As Long
    Dim blnValidarCatalogoCargos As Boolean
    '(FIN PEMEX)
    Dim strFolio As String
    Dim strSerie As String
    Dim strNumeroAprobacion As String
    Dim strAnoAprobacion As String
    Dim intTipoEmisionComprobante As Integer '|  0 = Error. Existen errores de configuración o faltan datos de comprobantes.
                                             '|  1 = Emisión física. Existen configurados folios y formato físico.
                                             '|  2 = Emisión digital. Existen configurados folios y formato digital.
    Dim intTipoDetalleFactura As Integer
    Dim lngAddendaComprobante As Long
    Dim blnFacturaAutomatica As Boolean
    Dim intTipoCFDFactura As Integer         'Variable que regresa el tipo de CFD de la factura(0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)

    Dim rsLogInterfazFactura As New ADODB.Recordset
    Dim vglngCveInterfazWS As Long
    Dim vllngCorteUsado As Long
    Dim vlRFCTemp As String
    Dim vlintEstadoFacturado As Integer
    Dim rsdf As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim vgintTipoPaciente As Integer

    Dim vlstrSentenciaDF As String

    If vllngAplAnt > 0 Then
        MsgBox SIHOMsg(1243), vbExclamation, "Mensaje" '¡No es posible realizar esta acción con esta factura! Por favor cancele la factura y genérela de nuevo.
        Exit Sub
    End If


    vlstrSentencia = "SELECT SMYIVA, MNYDESCUENTO, CHRESTATUS, INTMOVPACIENTE, CHRTIPOPACIENTE,"
    vlstrSentencia = vlstrSentencia & " SMIDEPARTAMENTO, INTCVEEMPLEADO,INTNUMCORTE, MNYANTICIPO, MNYTOTALFACTURA,"
    vlstrSentencia = vlstrSentencia & " BITPESOS, MNYTIPOCAMBIO, '" & frmDatosFiscales.txtTelefonoFactura & "', CHRTIPOFACTURA,"
    vlstrSentencia = vlstrSentencia & " INTNUMCLIENTE, INTCVEVENTAPUBLICO, INTCVEEMPRESA, MNYTOTALPAGAR,"
    vlstrSentencia = vlstrSentencia & " MNYTOTALNOTASCREDITO, MNYHONORARIOSFACTURADOS, MNYDESCUENTOESPECIAL, INTCVEUSOCFDI, NUMPORCENTDESCUENTOESPECIAL, CHRINCLUIRCONCEPTOSSEGURO, INTDESGLOSECONCEPTOSVICFDI, BITFACTURAGLOBAL"
    vlstrSentencia = vlstrSentencia & " FROM PvFactura "
    vlstrSentencia = vlstrSentencia & " WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
    Set vlrsAnteriorFactura = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If Not vlrsAnteriorFactura.EOF Then
        If vlrsAnteriorFactura!chrEstatus = "C" Then
            MsgBox SIHOMsg(1229), vbExclamation, "Mensaje"   'No se puede refacturar, el estado de la factura cambió. Consulte de nuevo.
            Exit Sub
        End If

        If Not fblnPermiteRefacturacionCambioIVA Then
            MsgBox SIHOMsg(1243), vbExclamation, "Mensaje" '¡No es posible realizar esta acción con esta factura! Por favor cancele la factura y genérela de nuevo.
            Exit Sub
        End If

        'Licencia Seguros Descuentos
        If Not IsNull(vlrsAnteriorFactura!chrIncluirConceptosSeguro) Then
            If Not fblnLicenciaCFDIDesc(vlrsAnteriorFactura!chrIncluirConceptosSeguro) Then Exit Sub
        End If

        strTipoPacienteFactura = vlrsAnteriorFactura!CHRTIPOPACIENTE
        '----------------------------------------------------------'
        ' (PEMEX) Se busca el tipo de formato que se va a utilizar '
        '----------------------------------------------------------'
        blnValidarCatalogoCargos = False
        lngCveFormato = 1

        Select Case strTipoPacienteFactura
            Case "I", "E"
                Set rsDatosPac = frsEjecuta_SP(vglngMovPaciente & "|0|" & vgstrTipoPaciente & "|" & vgintClaveEmpresaContable, "sp_PvSelDatosPaciente")
                If Not rsDatosPac.EOF Then
                    lngCveTipoPaciente = IIf(IsNull(rsDatosPac!tnyCveTipoPaciente), 0, rsDatosPac!tnyCveTipoPaciente)
                    lngCveEmpresaPac = IIf(IsNull(rsDatosPac!intcveempresa), 0, rsDatosPac!intcveempresa)
                    If vlrsAnteriorFactura!chrTipoFactura = "P" Then
                        If lngCveEmpresaPac = 0 Then
                            frsEjecuta_SP vgintNumeroDepartamento & "|0|" & lngCveTipoPaciente & "|" & vgstrTipoPaciente, "fn_PVSelFormatoFactura", True, lngCveFormato
                        Else
                            frsEjecuta_SP vgintNumeroDepartamento & "|" & lngCveEmpresaPac & "|" & lngCveTipoPaciente & "|" & vgstrTipoPaciente, "fn_PVSelFormatoFactura2", True, lngCveFormato
                        End If
                    ElseIf vlrsAnteriorFactura!chrTipoFactura = "E" Then
                        frsEjecuta_SP vgintNumeroDepartamento & "|" & vlrsAnteriorFactura!intcveempresa & "|" & lngCveTipoPaciente & "|" & vgstrTipoPaciente, "fn_PVSelFormatoFactura", True, lngCveFormato
                        blnValidarCatalogoCargos = True
                    Else
                       lngCveFormato = 0
                    End If
                Else
                    lngCveFormato = 0
                End If
                rsDatosPac.Close
            Case "G"
                frsEjecuta_SP vgintNumeroDepartamento & "|" & vlrsAnteriorFactura!intcveempresa & "|0|G", "fn_PVSelFormatoFactura", True, lngCveFormato
                blnValidarCatalogoCargos = True
            Case "V"
                frsEjecuta_SP vgintNumeroDepartamento & "|0|0|E", "fn_PVSelFormatoFactura", True, lngCveFormato
        End Select

        'Verifica si se usa un catálogo especial y si todos los cargos estan dentro de él (PEMEX, PCE)
        If blnValidarCatalogoCargos Then
            If fblnManejaCatalogoCargos(vlrsAnteriorFactura!intcveempresa) Then
                If fblnCargosFueraCatalogo(vglngMovPaciente, strTipoPacienteFactura, vlrsAnteriorFactura!intcveempresa) Then
                   Exit Sub
                End If
            End If
        End If
        vllngFormatoaUsar = lngCveFormato
        '-------------------------------------------------------
        '(FIN PEMEX)
        '-------------------------------------------------------

        '-------------------------------------------'
        ' Que exista formato de factura configurado '
        '-------------------------------------------'
        If vllngFormatoaUsar = 0 Then
            'No se encontró un formato válido de factura, por favor de uno de alta.
            MsgBox SIHOMsg(373), vbCritical, "Mensaje"
            Exit Sub
        End If
        '----------------------------------------------------'
        ' Validación de que tenga una impresora seleccionada '
        '----------------------------------------------------'
        vlstrSentencia = "select chrNombreImpresora Impresora from ImpresoraDepartamento where chrTipo = 'FA' and smiCveDepartamento = " & Trim(Str(vgintNumeroDepartamento))
        Set rsTemp = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsTemp.RecordCount > 0 Then
            For Each vPrinter In Printers
                If UCase(Trim(vPrinter.DeviceName)) = UCase(Trim(rsTemp!Impresora)) Then
                     Set Printer = vPrinter
                End If
            Next
        Else
            'No se tiene asignada una impresora en la cual imprimir las facturas
            MsgBox SIHOMsg(492), vbCritical, "Mensaje"
            Exit Sub
        End If
        rsTemp.Close

        ''****************************************************************************************
        '       VALIDACION DE FORMATO/FOLIO (FISICO, DIGITAL)
        '****************************************************************************************
        'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
        '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
        intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", vllngFormatoaUsar)

        If intTipoEmisionComprobante = 0 Then   'ERROR
            'Si es error, se cancela la transacción
            Exit Sub
        End If

        'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
        intTipoCFDFactura = fintTipoCFD("FA", vllngFormatoaUsar)

        'Si aparece un error terminar la transacción
        If intTipoCFDFactura = 3 Then   'ERROR
            'Si es error, se cancela la transacción
            Exit Sub
        End If
        '****************************************************************************************
        If Not fblnValidaCuentaPuenteBanco(vgintClaveEmpresaContable) Then Exit Sub

        lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If lngPersonaGraba = 0 Then Exit Sub

        vllngPersonaGraba = lngPersonaGraba
        lngAddendaComprobante = 0
        vlstrSentencia = "SELECT CFD.INTTIPODETALLEFACTURA, CFD.INTADDENDACOMPROBANTE " & _
                             "  FROM GNCOMPROBANTEFISCALDIGITAL CFD " & _
                             " WHERE CFD.INTCOMPROBANTE =  " & lngConsecutivoFactura & _
                             "   AND CFD.CHRTIPOCOMPROBANTE = 'FA'"
        Set rsTemp = frsRegresaRs(vlstrSentencia)
        If rsTemp.RecordCount > 0 Then
           intTipoDetalleFactura = IIf(IsNull(rsTemp!intTipoDetalleFactura), "0", rsTemp!intTipoDetalleFactura)
           'Se obtiene la addenda con la que se generó el comprobante (en caso de que así haya sido)
           lngAddendaComprobante = IIf(IsNull(rsTemp!intAddendaComprobante), "0", rsTemp!intAddendaComprobante)
        Else
           intTipoDetalleFactura = 1
        End If

        '------------------------------------------------------------------------'
        ' Checo que esa factura no tenga pagos registrados en Crédito y cobranza '
        '------------------------------------------------------------------------'
        vlstrSentencia = "SELECT COUNT(intNumMovimiento) FROM ccMovimientoCredito " & _
                         " WHERE chrFolioReferencia = '" & Trim(txtFolio.Text) & "'" & _
                         " AND chrTipoReferencia = 'FA' " & _
                         " AND mnyCantidadPagada > 0 "
        Set rsChecaCredito = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rsChecaCredito.Fields(0) > 0 Then
            MsgBox SIHOMsg(368), vbCritical, "Mensaje"
            Exit Sub
        End If
        rsChecaCredito.Close

        Set rsTemp = frsEjecuta_SP(Trim(txtFolio.Text), "Sp_PvSelFacturaAutomatica")
        If rsTemp.RecordCount > 0 Then
            blnFacturaAutomatica = True
        End If
        rsTemp.Close

        Load frmDatosFiscales

        frmDatosFiscales.vgActivaSujetoaIEPS = fblnAvtivasujetoIEPS

        If vlrsAnteriorFactura!BITFACTURAGLOBAL = 1 Then
            frmDatosFiscales.vgblnModalResult = True

            vlstrSentenciaDF = "SELECT CHRRFCPOS RFC, null Clave, CHRNOMBREFACTURAPOS Nombre, CHRDIRECCIONPOS chrCalle," & _
                             "VCHNUMEROEXTERIORPOS vchNumeroExterior, VCHNUMEROINTERIORPOS vchNumeroInterior, " & _
                             "null Telefono,'OT' Tipo,  INTCVECIUDAD IdCiudad, VCHCOLONIAPOS Colonia, VCHCODIGOPOSTALPOS CP FROM PVParametro " & _
                             "WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
            Set rsdf = frsRegresaRs(vlstrSentenciaDF, adLockOptimistic)

            If rsdf.RecordCount > 0 Then
                With frmDatosFiscales
                    .vgstrNombre = IIf(IsNull(rsdf!Nombre), "", Trim(rsdf!Nombre))
                    .vgstrDireccion = IIf(IsNull(rsdf!CHRCALLE), "", Trim(rsdf!CHRCALLE))
                    .vgstrNumExterior = IIf(IsNull(rsdf!VCHNUMEROEXTERIOR), "", Trim(rsdf!VCHNUMEROEXTERIOR))
                    .vgstrNumInterior = IIf(IsNull(rsdf!VCHNUMEROINTERIOR), "", Trim(rsdf!VCHNUMEROINTERIOR))
                    .vgstrColonia = IIf(IsNull(rsdf!Colonia), "", Trim(rsdf!Colonia))
                    .vgstrCP = IIf(IsNull(rsdf!CP), "", Trim(rsdf!CP))
                    .cboCiudad.ListIndex = flngLocalizaCbo(.cboCiudad, Str(IIf(IsNull(rsdf!IdCiudad), 0, rsdf!IdCiudad)))
                    .llngCveCiudad = .cboCiudad.ItemData(frmDatosFiscales.cboCiudad.ListIndex)
                    .vgstrTelefono = ""
                    .vgstrRFC = "XAXX010101000"
                    .vlstrNumRef = "NULL"
                    .vlstrTipo = IIf(IsNull(rsdf!Tipo), "OT", rsdf!Tipo)
                    .vglngDatosParametro = True

                    .chkIEPS.Value = IIf(frmDatosFiscales.vgActivaSujetoaIEPS = True, 1, 0)

                    .vgstrTipoUsoCFDI = "TP"

                    Set rs = frsRegresaRs("select distinct vchValor from SiParametro where trim(VCHNOMBRE) = 'INTTIPOPARTICULAR'", adLockReadOnly, adOpenForwardOnly)
                    vgintTipoPaciente = rs!vchvalor
                    .vgintTipoPacEmp = vgintTipoPaciente

                    .chkExtranjero.Value = False
                    .txtRFC.Text = "XAXX010101000"
                    .txtNombreFactura.Text = IIf(IsNull(rsdf!Nombre), "", Trim(rsdf!Nombre))
                    .txtDireccionFactura.Text = IIf(IsNull(rsdf!CHRCALLE), "", Trim(rsdf!CHRCALLE))
                    .txtNumExterior.Text = IIf(IsNull(rsdf!VCHNUMEROEXTERIOR), "", Trim(rsdf!VCHNUMEROEXTERIOR))
                    .txtNumInterior.Text = IIf(IsNull(rsdf!VCHNUMEROINTERIOR), "", Trim(rsdf!VCHNUMEROINTERIOR))
                    .txtColonia.Text = IIf(IsNull(rsdf!Colonia), "", Trim(rsdf!Colonia))
                    .txtCP.Text = IIf(IsNull(rsdf!CP), "", Trim(rsdf!CP))
                    .txtTelefonoFactura.Text = ""
                End With
            Else
                frmDatosFiscales.sstDatos.Tab = 1
                frmDatosFiscales.Show vbModal
            End If
        Else
            'Inicializa la pantalla de datos fiscales en la busqueda / Otros
            frmDatosFiscales.sstDatos.Tab = 1
            frmDatosFiscales.Show vbModal
        End If

        If frmDatosFiscales.vgblnModalResult Then
            '----------------------
            ' Inicio de Transacción
            '----------------------
            EntornoSIHO.ConeccionSIHO.BeginTrans

            '----------------------------------
            ' Obtener el numero de corte actual
            '----------------------------------
            vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")

            vlblnCorteValido = True
            If vllngMensaje <> 0 Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'Que el corte debe ser cerrado por cambio de día ó Que no existe corte abierto
               vlblnCorteValido = False
               MsgBox SIHOMsg(Str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje"
               Exit Sub
            End If

            vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
            pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            '||    C R E A C I Ó N   D E   L A   N U E V A   F A C T U R A     ||'
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            '------------------------------------'
            '        Número de la factura        '
            '------------------------------------'
            vllngFoliosFaltantes = 0
            pCargaArreglo vlaryParametrosSalida, vllngFoliosFaltantes & "|" & ADODB.adBSTR & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
            frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "sp_gnFolios", , , vlaryParametrosSalida
            pObtieneValores vlaryParametrosSalida, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
            '|  Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
            strSerie = Trim(strSerie)
            vlstrFolioDocumento = strSerie & strFolio
            If Trim(vlstrFolioDocumento) = "0" Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'No existen folios activos para este documento.
               MsgBox SIHOMsg(291), vbCritical, "Mensaje"
               Exit Sub
            End If
            '---------------------------------------------------------------'
            ' Actualiza la fecha de facturación en caso de que sea un grupo '
            '---------------------------------------------------------------'
            If Not optTipoPaciente(0).Value And Not optTipoPaciente(1).Value Then
               vlstrSentencia = "UPDATE PvFacturacionConsolidada SET chrFolioFactura = '" & vlstrFolioDocumento & "' WHERE INTCVEGRUPO = " & txtMovimientoPaciente.Text
               pEjecutaSentencia vlstrSentencia
            End If

            Set vlrsNuevaFactura = frsRegresaRs("SELECT * FROM PvFactura WHERE PvFactura.INTCONSECUTIVO = -1", adLockOptimistic, adOpenDynamic)
            With vlrsNuevaFactura
                 .AddNew
                 !chrfoliofactura = vlstrFolioDocumento
                 !dtmFechahora = fdtmServerFecha + fdtmServerHora
                 !chrRFC = IIf(frmDatosFiscales.chkExtranjero, "XEXX010101000", IIf(Len(Trim(frmDatosFiscales.txtRFC)) < 12, "XAXX010101000", Trim(frmDatosFiscales.txtRFC)))
                 !CHRNOMBRE = IIf(Trim(frmDatosFiscales.txtNombreFactura) = "", " ", frmDatosFiscales.txtNombreFactura)
                 !CHRCALLE = IIf(Trim(frmDatosFiscales.txtDireccionFactura) = "", " ", frmDatosFiscales.txtDireccionFactura)
                 !VCHNUMEROEXTERIOR = IIf(Trim(frmDatosFiscales.txtNumExterior) = "", "", frmDatosFiscales.txtNumExterior)
                 !VCHNUMEROINTERIOR = IIf(Trim(frmDatosFiscales.txtNumInterior) = "", "", frmDatosFiscales.txtNumInterior)
                 !VCHCOLONIA = IIf(Trim(frmDatosFiscales.txtColonia) = "", " ", frmDatosFiscales.txtColonia)
                 !VCHCODIGOPOSTAL = IIf(Trim(frmDatosFiscales.txtCP) = "", " ", frmDatosFiscales.txtCP)
                 !smyIVA = vlrsAnteriorFactura!smyIVA
                 !MNYDESCUENTO = vlrsAnteriorFactura!MNYDESCUENTO
                 !chrEstatus = vlrsAnteriorFactura!chrEstatus
                 !INTMOVPACIENTE = vlrsAnteriorFactura!INTMOVPACIENTE
                 !CHRTIPOPACIENTE = vlrsAnteriorFactura!CHRTIPOPACIENTE
                 !SMIDEPARTAMENTO = vlrsAnteriorFactura!SMIDEPARTAMENTO
                 !intCveEmpleado = lngPersonaGraba
                 !intnumcorte = vllngNumeroCorte
                 !mnyAnticipo = vlrsAnteriorFactura!mnyAnticipo
                 !mnyTotalFactura = vlrsAnteriorFactura!mnyTotalFactura
                 !BITPESOS = vlrsAnteriorFactura!BITPESOS
                 !MNYTIPOCAMBIO = vlrsAnteriorFactura!MNYTIPOCAMBIO
                 vldblTipoCambio = IIf(vlrsAnteriorFactura!MNYTIPOCAMBIO = 0, 1, vlrsAnteriorFactura!MNYTIPOCAMBIO)
                 !CHRTELEFONO = frmDatosFiscales.txtTelefonoFactura
                 !chrTipoFactura = vlrsAnteriorFactura!chrTipoFactura
                 !intNumCliente = vlrsAnteriorFactura!intNumCliente
                 !intCveVentaPublico = IIf(blnFacturaAutomatica, -1, vlrsAnteriorFactura!intCveVentaPublico)
                 !INTCVECIUDAD = frmDatosFiscales.llngCveCiudad
                 !intcveempresa = vlrsAnteriorFactura!intcveempresa
                 !mnyTotalPagar = vlrsAnteriorFactura!mnyTotalPagar
                 !mnyTotalNotasCredito = vlrsAnteriorFactura!mnyTotalNotasCredito
                 !vchSerie = strSerie
                 !INTFOLIO = strFolio
                 !mnyHonorariosFacturados = vlrsAnteriorFactura!mnyHonorariosFacturados
                 !bitdesgloseIEPS = IIf(frmDatosFiscales.chkIEPS.Value = vbChecked, 1, 0)
                 !MNYDESCUENTOESPECIAL = vlrsAnteriorFactura!MNYDESCUENTOESPECIAL
                 !intCveUsoCFDI = vlrsAnteriorFactura!intCveUsoCFDI
                 !NUMPORCENTDESCUENTOESPECIAL = vlrsAnteriorFactura!NUMPORCENTDESCUENTOESPECIAL
                 !intTipoDetalleFactura = intTipoDetalleFactura '<--el detalle que se usa para el timbrado, pudiera cambiar....
                 !chrIncluirConceptosSeguro = vlrsAnteriorFactura!chrIncluirConceptosSeguro
                 'Se agrega para refacturacion
                 !VCHREGIMENFISCALRECEPTOR = frmDatosFiscales.cboRegimenFiscal.ItemData(frmDatosFiscales.cboRegimenFiscal.ListIndex)

                 !INTDESGLOSECONCEPTOSVICFDI = vlrsAnteriorFactura!INTDESGLOSECONCEPTOSVICFDI
                 !BITFACTURAGLOBAL = vlrsAnteriorFactura!BITFACTURAGLOBAL
                 .Update
                 vllngCveFacturaNueva = flngObtieneIdentity("SEC_PVFACTURA", !intConsecutivo)
                 vlRFCTemp = Trim(!chrRFC) 'Se graba el valor del RFC para posible envío de CFD/CFDi
            End With
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            'IEPS
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            If (vgstrTipoFactura = "T" Or vgstrTipoFactura = "D") And vlblnLicenciaIEPS Then
               If fblnActualizaIEPS(vllngCveFacturaNueva, IIf(frmDatosFiscales.chkIEPS.Value = vbChecked, True, False)) = False Then
                  vlstrSentencia = " INSERT INTO PvDetalleFactura (CHRFOLIOFACTURA, SMICVECONCEPTO, MNYCANTIDAD, MNYDESCUENTO, MNYIVA,CHRTIPO,MNYCANTIDADGRAVADA, MNYIVACONCEPTO, MNYIEPS, NUMTASAIEPS)"
                  vlstrSentencia = vlstrSentencia & " SELECT  '" & vlstrFolioDocumento & "', SMICVECONCEPTO,MNYCANTIDAD,MNYDESCUENTO,MNYIVA,CHRTIPO,MNYCANTIDADGRAVADA, MNYIVACONCEPTO, MNYIEPS, NUMTASAIEPS "
                  vlstrSentencia = vlstrSentencia & " FROM PvDetalleFactura "
                  vlstrSentencia = vlstrSentencia & " WHERE PvDetalleFactura.CHRFOLIOFACTURA = '" & Trim(txtFolio.Text) & "'"
                  pEjecutaSentencia (vlstrSentencia)

                  vlstrSentencia = " INSERT INTO PvFacturaimporte (INTCONSECUTIVO, MNYSUBTOTALGRAVADO, MNYSUBTOTALNOGRAVADO, MNYDESCUENTOGRAVADO, MNYDESCUENTONOGRAVADO)"
                  vlstrSentencia = vlstrSentencia & " SELECT  " & vllngCveFacturaNueva & "  , PvFacturaIMPORTE.MNYSUBTOTALGRAVADO,PvFacturaIMPORTE.MNYSUBTOTALNOGRAVADO, "
                  vlstrSentencia = vlstrSentencia & " PVFACTURAIMPORTE.MNYDESCUENTOGRAVADO, PVFACTURAIMPORTE.MNYDESCUENTONOGRAVADO FROM PvFacturaIMPORTE INNER JOIN PvFACTURA ON PVFACTURA.INTCONSECUTIVO = PVFACTURAIMPORTE.INTCONSECUTIVO"
                  vlstrSentencia = vlstrSentencia & " WHERE PvFactura.CHRFOLIOFACTURA = '" & Trim(txtFolio.Text) & "'"
                  pEjecutaSentencia (vlstrSentencia)
               End If
            Else
               vlstrSentencia = " INSERT INTO PvDetalleFactura (CHRFOLIOFACTURA, SMICVECONCEPTO, MNYCANTIDAD, MNYDESCUENTO, MNYIVA,CHRTIPO,MNYCANTIDADGRAVADA, MNYIVACONCEPTO)"
               vlstrSentencia = vlstrSentencia & " SELECT  '" & vlstrFolioDocumento & "', SMICVECONCEPTO,MNYCANTIDAD,MNYDESCUENTO,MNYIVA,CHRTIPO,MNYCANTIDADGRAVADA, MNYIVACONCEPTO "
               vlstrSentencia = vlstrSentencia & " FROM PvDetalleFactura "
               vlstrSentencia = vlstrSentencia & " WHERE PvDetalleFactura.CHRFOLIOFACTURA = '" & Trim(txtFolio.Text) & "'"
               pEjecutaSentencia (vlstrSentencia)

               vlstrSentencia = " INSERT INTO PvFacturaimporte (INTCONSECUTIVO, MNYSUBTOTALGRAVADO, MNYSUBTOTALNOGRAVADO, MNYDESCUENTOGRAVADO, MNYDESCUENTONOGRAVADO)"
               vlstrSentencia = vlstrSentencia & " SELECT  " & vllngCveFacturaNueva & "  , PvFacturaIMPORTE.MNYSUBTOTALGRAVADO,PvFacturaIMPORTE.MNYSUBTOTALNOGRAVADO, "
               vlstrSentencia = vlstrSentencia & " PVFACTURAIMPORTE.MNYDESCUENTOGRAVADO, PVFACTURAIMPORTE.MNYDESCUENTONOGRAVADO FROM PvFacturaIMPORTE INNER JOIN PvFACTURA ON PVFACTURA.INTCONSECUTIVO = PVFACTURAIMPORTE.INTCONSECUTIVO"
               vlstrSentencia = vlstrSentencia & " WHERE PvFactura.CHRFOLIOFACTURA = '" & Trim(txtFolio.Text) & "'"
               pEjecutaSentencia (vlstrSentencia)
            End If
            '----------------------------'
            ' Actualiza PvCargoExcedente '
            '----------------------------'
            vlstrSentencia = "UPDATE PvCargoExcedente SET chrfoliofactura = '" & vlstrFolioDocumento & "'" & " WHERE chrfoliofactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia vlstrSentencia
            '-----------------------------------'
            ' Actualiza PvFacturaParcialEmpresa '
            '-----------------------------------'
            vlstrSentencia = "UPDATE PvFacturaParcialEmpresa SET intFacturaEmpresa = '" & CStr(vllngCveFacturaNueva) & "' "
            vlstrSentencia = vlstrSentencia & "WHERE PvFacturaParcialEmpresa.intFacturaEmpresa = " & CStr(vglngCveFactura)
            pEjecutaSentencia (vlstrSentencia)
            '--------------------------------------'
            ' Actualizar PvFacturaPacienteConcepto '
            '--------------------------------------'
            vlstrSentencia = "UPDATE PvFacturaPacienteConcepto SET chrFolioFactura = '" & vlstrFolioDocumento
            vlstrSentencia = vlstrSentencia & "' WHERE PvFacturaPacienteConcepto.chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)

            vllngNumCorteFactura = frsRegresaRs("SELECT intNumCorte FROM PvFactura WHERE ltrim(rtrim(chrFolioFactura)) = '" & Trim(txtFolio.Text) & "'").Fields(0)
            vlstrSentencia = "SELECT * FROM PvDetalleCorte WHERE chrFolioDocumento = '" & Trim(txtFolio.Text) & "' AND chrTipoDocumento = 'FA'"
            Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)           'RS de consulta

            If rsDC.RecordCount > 0 Then
               Do While Not rsDC.EOF
                  '~~~~~~~~~~~~~~~~~~~  F A C T U R A C I Ó N  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                  '||  Generar registros en PVDetalleCorte para generar la NUEVA FACTURA  ||'
                  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

                  Set rsChequeTransCta = frsRegresaRs("SELECT * FROM PVCORTECHEQUETRANSCTA WHERE INTCONSECUTIVODETCORTE = " & rsDC!intConsecutivo, adLockReadOnly, adOpenForwardOnly)
                  If rsChequeTransCta.RecordCount > 0 Then
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, vlstrFolioDocumento, rsDC!chrTipoDocumento, 0, rsDC!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                        1, "", "", False, Trim(Replace(Replace(Replace(vlRFCTemp, "-", ""), "_", ""), " ", "")), IIf(IsNull(rsChequeTransCta!CHRCLAVEBANCOSAT), "", rsChequeTransCta!CHRCLAVEBANCOSAT), IIf(IsNull(rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), "", rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), IIf(IsNull(rsChequeTransCta!VCHCUENTABANCARIA), "", rsChequeTransCta!VCHCUENTABANCARIA), IIf(IsNull(rsChequeTransCta!dtmfecha), fdtmServerFecha, rsChequeTransCta!dtmfecha)
                  Else
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, vlstrFolioDocumento, rsDC!chrTipoDocumento, 0, rsDC!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                        1, "", ""
                  End If

                  rsDC.MoveNext
               Loop
            End If

            vlstrSentencia = "SELECT DISTINCT  chrFolioDocumento, chrTipoDocumento, intFormaPago, mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                                " intNumCorteDocumento FROM PvDetalleCorte " & _
                                " WHERE chrFolioDocumento IN (SELECT chrFolioRecibo " & _
                                                            " FROM PvPago " & _
                                                            " WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "') " & _
                                " AND mnyCantidadPagada > 0 " & _
                                " AND chrTipoDocumento = 'RE' "
            If vgstrFacturaPacienteEmpresa = "P" Then  ' Sólo las facturas de pacientes tienen salidas de Efectivo
               vlstrSentencia = vlstrSentencia & " UNION SELECT DISTINCT  chrFolioDocumento, chrTipoDocumento, intFormaPago, " & _
                                    " mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                                    " intNumCorteDocumento  " & _
                                    " FROM PvDetalleCorte " & _
                                    " WHERE chrFolioDocumento IN (SELECT chrFolioRecibo " & _
                                                                " FROM PvSalidaDinero " & _
                                                                " WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "') " & _
                                    " AND mnyCantidadPagada > 0 " & _
                                    " AND chrTipoDocumento = 'SD' "
            End If
            Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly) 'RS de consulta
            If rsDC.RecordCount > 0 Then
               Do While Not rsDC.EOF
                  '~~~~~~~~~~~~~~~~ F A C T U R A C I Ó N ~~~~~~~~~~~~~~~~'
                  '||        Cancelación de los pagos en el corte       ||'
                  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                  pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsDC!chrFolioDocumento, rsDC!chrTipoDocumento, 0, (rsDC!mnyCantidadPagada * -1), _
                  False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                  1, "", ""

                  rsDC.MoveNext
               Loop
            End If

            '-----------------------------------------------------------------------------------------------------'
            ' Registrar en el corte movimientos, solo para Facturas de Tikets, para que se quede la venta intacta '
            '-----------------------------------------------------------------------------------------------------'
            If vgstrTipoFactura = "T" Then 'Sólo se reactivan las ventas de las facturas de tickets
               vlstrSentencia = "SELECT DISTINCT chrFolioDocumento, " & _
                                " chrTipoDocumento, intFormaPago, mnyCantidadPagada, " & _
                                " mnyTipoCambio, intFolioCheque, intNumCorteDocumento " & _
                                " FROM PvDetalleCorte " & _
                                " WHERE mnyCantidadPagada > 0 " & _
                                " AND chrTipoDocumento = 'TI' " & _
                                " AND chrFolioDocumento IN (SELECT chrFolioTicket FROM PvVentaPublico WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "')"
               Set rsCorteTiKets = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
               If rsCorteTiKets.RecordCount > 0 Then
                  'Guardar en el Corte
                  Do While Not rsCorteTiKets.EOF
                     '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                     '||    F A C T U R A C I Ó N     ||'
                     '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                     pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsCorteTiKets!chrFolioDocumento, "TI", 0, (rsCorteTiKets!mnyCantidadPagada * -1), _
                     False, (fdtmServerFecha + fdtmServerHora), rsCorteTiKets!intFormaPago, rsCorteTiKets!MNYTIPOCAMBIO, rsCorteTiKets!intfoliocheque, _
                     rsCorteTiKets!intNumCorteDocumento, 1, "", ""
                     rsCorteTiKets.MoveNext
                  Loop
               End If
               rsCorteTiKets.Close
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            '||  A C T U A L I Z A   L A   F A C T U R A   D E   F A C T U R A S   P A R C I A L E S  ||'
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            vlstrSentencia = " SELECT PvFacturasParciales.INTFACTURABASE FROM PvFacturasParciales "
            vlstrSentencia = vlstrSentencia & " WHERE PvFacturasParciales.INTFACTURABASE = " & vglngCveFactura
            Set rsTemp = frsRegresaRs(vlstrSentencia)
            If rsTemp.RecordCount > 0 Then
               'Si es factura base
               vlstrSentencia = " UPDATE PvFacturasParciales SET PvFacturasParciales.INTFACTURABASE = " & vllngCveFacturaNueva & _
                                " WHERE PvFacturasParciales.INTFACTURABASE = " & vglngCveFactura
               pEjecutaSentencia vlstrSentencia
            Else
               vlstrSentencia = " SELECT PvFacturasParciales.INTFACTURAPARCIAL FROM PvFacturasParciales "
               vlstrSentencia = vlstrSentencia & " WHERE PvFacturasParciales.INTFACTURAPARCIAL = " & vglngCveFactura
               Set rsTemp = frsRegresaRs(vlstrSentencia)
               If rsTemp.RecordCount > 0 Then
                  'Si es factura parcial
                  vlstrSentencia = " UPDATE PvFacturaParcialEmpresa SET PvFacturaParcialEmpresa.INTFACTURAPARCIAL = " & vllngCveFacturaNueva & _
                                   " WHERE PvFacturaParcialEmpresa.INTFACTURAPARCIAL =" & vglngCveFactura
                  pEjecutaSentencia vlstrSentencia

                  vlstrSentencia = " UPDATE PvFacturasParciales SET PvFacturasParciales.INTFACTURAPARCIAL = " & vllngCveFacturaNueva & _
                                   " WHERE PvFacturasParciales.INTFACTURAPARCIAL = " & vglngCveFactura
                  pEjecutaSentencia vlstrSentencia
               End If
            End If

            '-------------------------------------------------------------------------------------------'
            'Durante este proceso se actualiza el folio de la factura de los pagos, salidas de dinero y '
            'venta al público en lugar de cancelar y reactivar las cosas                                '
            '-------------------------------------------------------------------------------------------'
            vlstrSentencia = "UPDATE PvPago SET chrFolioFactura = '" & vlstrFolioDocumento & "' WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            vlstrSentencia = "UPDATE PvSalidaDinero SET chrFolioFactura = '" & vlstrFolioDocumento & "' WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            If blnFacturaAutomatica Then
               'Si es factura automática, que se quite el bitFacturaAutomatica
               'porque se supone que si refacturan no va a ser factura por venta al público
               vlstrSentencia = "UPDATE PvVentaPublico SET bitFacturaAutomatica = 0 WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "' And chrTipoRecivo = 'T'"
               pEjecutaSentencia vlstrSentencia
            End If
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            vlstrSentencia = "UPDATE PvVentaPublico SET CHRFOLIOTICKET = '" & vlstrFolioDocumento & "', chrFolioFactura = '" & vlstrFolioDocumento & "'" & _
                             " WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "' AND CHRTIPORECIVO = 'F'"
            pEjecutaSentencia vlstrSentencia
            '---------------------------------------------------------------------------------------------------------------------------------------------------------
            vlstrSentencia = "UPDATE PvVentaPublico SET chrFolioFactura = '" & vlstrFolioDocumento & "' WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & _
                             "' AND CHRTIPORECIVO = 'T'"
            pEjecutaSentencia vlstrSentencia
            '--------------------------------------------------------------------------------'
            ' Actualizar los movimientos de CCMovimientoCredito al folio de la nueva factura '
            '--------------------------------------------------------------------------------'
            vgstrParametrosSP = Trim(txtFolio.Text) & "|" & vlstrFolioDocumento
            frsEjecuta_SP vgstrParametrosSP, "SP_PVINSCREDITOREFACTURACION"
            rsDC.Close
            '-------------------------------------------------------------------'
            ' Se actualiza el folio de la factura los cargos con el nuevo Folio '
            '-------------------------------------------------------------------'
            vlstrSentencia = "UPDATE PvCargo SET chrFolioFactura = '" & vlstrFolioDocumento & "' WHERE chrFolioFactura = '" & RTrim(txtFolio.Text) & "'"
            pEjecutaSentencia vlstrSentencia
            '-----------------------------------------------------------'
            'Se actualiza el Folio de la factura del paciente en PvCargo'
            '-----------------------------------------------------------'
            If vgstrFacturaPacienteEmpresa = "P" And lblnCalcularEnBaseACargos Then
                vlstrSentencia = "UPDATE PvCargo SET chrFolioFacturaPaciente = '" & vlstrFolioDocumento & "' WHERE chrFolioFacturaPaciente = '" & RTrim(txtFolio.Text) & "'"
                pEjecutaSentencia vlstrSentencia
            End If
            '-------------------------------------------'
            ' Creación de la nueva factura, SOLO POLIZA '
            '-------------------------------------------'
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, vlstrFolioDocumento, "", 0, 0, False, "", 0, 0, "", 0, 5, txtFolio.Text, ""
            '----------------------------------------------'
            ' Actualiza el folio del control de aseguradora'
            '----------------------------------------------'
            vlstrSentencia = _
                            "UPDATE PVCONTROLASEGURADORA SET " & _
                                "CHRFOLIOFACTURADEDUCIBLE = CASE WHEN TRIM(CHRFOLIOFACTURADEDUCIBLE) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURADEDUCIBLE END " & _
                                ",CHRFOLIOFACTURACOASEGURO = CASE WHEN TRIM(CHRFOLIOFACTURACOASEGURO) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURACOASEGURO END " & _
                                ",CHRFOLIOFACTURACOPAGO = CASE WHEN TRIM(CHRFOLIOFACTURACOPAGO) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURACOPAGO END " & _
                                ",CHRFOLIOFACTURAEXCEDENTE = CASE WHEN TRIM(CHRFOLIOFACTURAEXCEDENTE) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURAEXCEDENTE END " & _
                                ",CHRFOLIOFACTURAEMPRESA = CASE WHEN TRIM(CHRFOLIOFACTURAEMPRESA) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURAEMPRESA END " & _
                                ",CHRFOLIOFACTURACOASEGUROADICI = CASE WHEN TRIM(CHRFOLIOFACTURACOASEGUROADICI) = '" & Trim(txtFolio.Text) & "' THEN '" & vlstrFolioDocumento & "' ELSE CHRFOLIOFACTURACOASEGUROADICI END " & _
                            "WHERE " & _
                                "TRIM(CHRFOLIOFACTURADEDUCIBLE) = '" & Trim(txtFolio.Text) & "' OR TRIM(CHRFOLIOFACTURACOASEGURO) = '" & Trim(txtFolio.Text) & "' OR TRIM(CHRFOLIOFACTURACOPAGO) = '" & Trim(txtFolio.Text) & "' OR TRIM(CHRFOLIOFACTURAEXCEDENTE) = '" & Trim(txtFolio.Text) & "' OR TRIM(CHRFOLIOFACTURAEMPRESA) = '" & Trim(txtFolio.Text) & "'" & " OR TRIM(CHRFOLIOFACTURACOASEGUROADICI) = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia vlstrSentencia
            '------------------------------------------------------------------------------'
            ' (PEMEX)Se consulta los datos de la factura para saber el número de registros '
            '------------------------------------------------------------------------------'
            blnFacturaMultiple = False
            If fblnObtieneDatosFormatoProc(lngCveFormato, vllngCveFacturaNueva, lngRenglonesDetalle, lngTotalDocumentos, rsDatosFactura) Then
                blnFacturaMultiple = True
            End If
            If blnFacturaMultiple Then
                blnCancelarFacturacion = False
                blnFoliosOK = False
                lngInicial = 0
                lngFinal = 0
                strIdentificador = ""
                Do
                    ReDim arrFolios(1)
                    arrFolios(1) = vlstrFolioDocumento
                    lngIndexFolios = 1
                    blnFoliosOK = True
                    If lngTotalDocumentos > 1 Then
                        Set rsFolios = frsEjecuta_SP("FA|" & vgintNumeroDepartamento & "|" & lngTotalDocumentos - 1 & "|1|" & strIdentificador & "|" & lngInicial & "|" & lngFinal & "|0", "sp_GNSelFoliosFactMult")
                        Do Until rsFolios.EOF
                            lngIndexFolios = lngIndexFolios + 1
                            ReDim Preserve arrFolios(lngIndexFolios)
                            If Not IsNull(rsFolios!vchFolio) Then
                               arrFolios(lngIndexFolios) = rsFolios!vchFolio
                            Else
                                blnFoliosOK = False
                            End If
                            rsFolios.MoveNext
                        Loop
                        rsFolios.Close
                    End If
                    If Not blnFoliosOK Then
                        If fmsgNuevaSerieFolios("FA", vgintNumeroDepartamento, strIdentificador, lngInicial, lngFinal) = vbCancel Then
                            blnCancelarFacturacion = True
                        End If
                    End If
                Loop Until blnCancelarFacturacion Or blnFoliosOK
                If blnCancelarFacturacion Then
                    EntornoSIHO.ConeccionSIHO.RollbackTrans
                    Exit Sub
                End If
                'Registro de folios
                For lngContador = 2 To lngIndexFolios
                    pEjecutaSentencia "insert into PVFacturaFolios(chrFolioFactura, chrFolioRelacionado, smiCveDepartamento, intnumDocumento) values('" & vlstrFolioDocumento & "', '" & arrFolios(lngContador) & "', " & vgintNumeroDepartamento & ", " & lngContador & ")"
                Next
            End If

            '------------------------------------------------------------------------------------------'
            ' Actualiza la factura en CcNota, CcNotaFactura, CcNotaDetalle y PVPAQUETEPACIENTEFACTURADO'
            '------------------------------------------------------------------------------------------'
            'CcNota
            vlstrSentencia = "Update CcNota Set vchFacturaImpresion = '" & vlstrFolioDocumento & "' Where vchFacturaImpresion = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)

            'CcNotaFactura
            vlstrSentencia = "Update CcNotaFactura Set chrFolioFactura = '" & vlstrFolioDocumento & "' Where chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)

            'CcNotaDetalle
            vlstrSentencia = "Update CcNotaDetalle Set chrFolioFactura = '" & vlstrFolioDocumento & "' Where chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)

            '-------------------------------------------------------------------------------------------------'
            'Pone como CANCELADA la factura que se re facturará en PVPAQUETEPACIENTEFACTURADO
            '-------------------------------------------------------------------------------------------------'
            strParametros = txtMovimientoPaciente.Text & "|" & IIf(optTipoPaciente(2).Value, "'G'", IIf(optTipoPaciente(0).Value, "'I'", "'E'")) & "|" & Trim(txtFolio.Text) & "|C"
            frsEjecuta_SP strParametros, "SP_PVUPDPAQUETESFACTURADOS"

            'Inserta en PVPAQUETEPACIENTEFACTURADO
            strParametros = txtMovimientoPaciente.Text & "|" & IIf(optTipoPaciente(2).Value, "'G'", IIf(optTipoPaciente(0).Value, "'I'", "'E'")) & "|" & Trim(vlstrFolioDocumento)
            frsEjecuta_SP strParametros, "SP_PVINSPAQUETESFACTURADOS"

            '---------------------------------------------------------------------------------------------------------------------------------------------'
            'Pone nuevamente como FACTURADA la factura que se re facturará en PVPAQUETEPACIENTEFACTURADO para no afectar las validaciones subsecuentes
            '---------------------------------------------------------------------------------------------------------------------------------------------'
            strParametros = txtMovimientoPaciente.Text & "|" & IIf(optTipoPaciente(2).Value, "'G'", IIf(optTipoPaciente(0).Value, "'I'", "'E'")) & "|" & Trim(txtFolio.Text) & "|F"
            frsEjecuta_SP strParametros, "SP_PVUPDPAQUETESFACTURADOS"

'            lngAddendaComprobante = 0
'            vlstrSentencia = "SELECT CFD.INTTIPODETALLEFACTURA, CFD.INTADDENDACOMPROBANTE " & _
'                             "  FROM GNCOMPROBANTEFISCALDIGITAL CFD " & _
'                             " WHERE CFD.INTCOMPROBANTE =  " & lngConsecutivoFactura & _
'                             "   AND CFD.CHRTIPOCOMPROBANTE = 'FA'"
'            Set rsTemp = frsRegresaRs(vlstrSentencia)
'            If rsTemp.RecordCount > 0 Then
'                intTipoDetalleFactura = IIf(IsNull(rsTemp!intTipoDetalleFactura), "0", rsTemp!intTipoDetalleFactura)
'                'Se obtiene la addenda con la que se generó el comprobante (en caso de que así haya sido)
'                lngAddendaComprobante = IIf(IsNull(rsTemp!intAddendaComprobante), "0", rsTemp!intAddendaComprobante)
'            Else
'                intTipoDetalleFactura = 1
'            End If

            '====================================================================================================='
            '================================= INICIO REFACTURACIÓN INTERFAZ AXA ================================='
            '====================================================================================================='
            'Se valida si la empresa seleccionada está configurada para usarse con alguna interfaz de WS
            vglngCveInterfazWS = 1
            frsEjecuta_SP lngCveEmpresaPac & "|" & vgintClaveEmpresaContable, "FN_GNSELINTERFAZWS", True, vglngCveInterfazWS

            'Se verifica si se cuenta con una licencia para la interfaz obtenida
            vglngCveInterfazWS = IIf(fblnLicenciaWS(vglngCveInterfazWS) = True, vglngCveInterfazWS, 0)
            If vglngCveInterfazWS <> 0 Then
                'Se selecciona el registro del log de transacciones de la interfaz para obtener la información necesaria
                vgstrParametrosSP = Trim(txtFolio.Text) & "|" & vglngMovPaciente & "|" & Trim(vgstrTipoPaciente)
                Set rsLogInterfazFactura = frsEjecuta_SP(vgstrParametrosSP, "SP_GNSELLOGINTERFAZAXAFACTURA")

                'Se seleccionan los datos de la factura original para que sean enviados en la nueva factura...
                If rsLogInterfazFactura.RecordCount > 0 Then
                    'Se obtiene la IP y el nombre de la máquina huesped
                    Call ObtenerPCIP
                    vgstrIP = vgstrNumeroIP
                    vgstrEquipo = vgstrNombreMaquina

                    '---------------------------------------------------------------------------------------------------------
                    If Not fblnValidacionWSAXA(Trim(rsLogInterfazFactura!VCHFOLIORECETA), Trim(vlstrFolioDocumento), CLng(rsLogInterfazFactura!INTCVEFOLIO), CLng(rsLogInterfazFactura!IntCveTipoIngreso), vglngMovPaciente, CLng(rsLogInterfazFactura!intNumPaciente), Trim(vgstrTipoPaciente), lngPersonaGraba) Then
                        On Error Resume Next
                        EntornoSIHO.ConeccionSIHO.RollbackTrans

                        'Se almacena en el log de transacciones de la interfaz de AXA después de una conexión fallida
                        pCargaArreglo vlaryParametrosSalida, "|" & adDouble 'adBSTR 'adVarChar1
                        vgstrParametrosSP = Trim(txtFolio.Text) & "|" & vgintNumeroModulo & "|" & CLng(rsLogInterfazFactura!IntCveTipoIngreso) & "|" & vglngMovPaciente & "|" & CLng(rsLogInterfazFactura!intNumPaciente) & "|" & Trim(vgstrTipoPaciente) & "|" & Trim(vgstrIP) & "|" & "FA" & "|NO|" & Trim(rsLogInterfazFactura!CLBXMLREQUEST) & "|" & strResponseXML & "|" & Trim(vgstrEquipo) & "|" & lngPersonaGraba & "|0|" & Trim(rsLogInterfazFactura!VCHFOLIORECETA)
                        frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa", , , vlaryParametrosSalida
                        pObtieneValores vlaryParametrosSalida, 0


                        '1307 La factura será cancelada en el sistema, será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT.
                        MsgBox "No se pudo validar la factura con el web service de AXA, se cancelará la transacción.", vbCritical, "Mensaje"

                        Exit Sub
                    End If
                    '---------------------------------------------------------------------------------------------------------
                End If
            End If
            '====================================================================================================='
            '================================= FINAL REFACTURACIÓN INTERFAZ AXA  ================================='
            '====================================================================================================='

            '----------------------------------------------------
            'insertar movimientos de la nueva factura en el corte
            '----------------------------------------------------
            vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte, True)

            If vllngCorteUsado = 0 Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
               MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
               Exit Sub
            End If

            If vllngCorteUsado <> vllngNumeroCorte Then
              'actualizamos el corte en el que se registró la factura, esto es por si hay un cambio de corte al momento de hacer el registro d ela información de la factura
               pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & vllngCveFacturaNueva
               vllngNumeroCorte = vllngCorteUsado
            End If

            If intTipoEmisionComprobante = 2 Then
               If Not fblnValidaDatosCFDCFDi(vllngCveFacturaNueva, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacion), strNumeroAprobacion) Then
                  EntornoSIHO.ConeccionSIHO.RollbackTrans
                 Exit Sub
               End If
            End If

            EntornoSIHO.ConeccionSIHO.CommitTrans 'aqui ya tenemos la nueva factura creada

            '--------------------------
            'Timbre de la nueva factura
            '--------------------------
            If intTipoEmisionComprobante = 2 Then
               pgbBarraCFD.Value = 70
               freBarraCFD.Top = 3200
               Screen.MousePointer = vbHourglass
               freBarraCFD.Visible = True
               freBarraCFD.Refresh
               frmConsultaFactura.Enabled = False

               pMarcarPendienteTimbre vllngCveFacturaNueva, "FA", vgintNumeroDepartamento
               blnNOMensajeErrorPAC = False
               pLogTimbrado 2

               EntornoSIHO.ConeccionSIHO.BeginTrans 'inicia el proceso de timbrado de la factura

               '-----------------------------------------'
               ' Asocia la factura cancelada y la creada '
               '-----------------------------------------'
               vlstrSentencia = "INSERT INTO PVREFACTURACION (chrFolioFacturaActivada, chrFolioFacturaCancelada) " & _
                                                     " VALUES ('" & vlstrFolioDocumento & "', '" & txtFolio.Text & "')"
               pEjecutaSentencia vlstrSentencia

               If Not fblnGeneraComprobanteDigital(vllngCveFacturaNueva, "FA", intTipoDetalleFactura, CInt(strAnoAprobacion), strNumeroAprobacion, IIf(intTipoCFDFactura = 1, True, False), , lngAddendaComprobante) Then
                  On Error Resume Next
                  Unload frmDatosFiscales
                  Set frmDatosFiscales = Nothing
                  pLogTimbrado 1
                  If vgIntBanderaTImbradoPendiente = 1 Then
                     EntornoSIHO.ConeccionSIHO.CommitTrans
                     intcontadorCFDiPendienteCancelar = 0
                     ReDim vlArrCFDiPendienteCancelar(0)
                     pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
                     pCFDiPendienteCancelar 0, 0, 1
                     '33 !No se pueden guardar los datos!
                     '1306 El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal.
                     '1307 La factura será cancelada en el sistema, será necesario confirmar el timbre fiscal y realizar la cancelación ante el SAT.
                     MsgBox SIHOMsg(33) & vbNewLine & SIHOMsg(1319) & vbNewLine & _
                            Replace(SIHOMsg(1306), "El comprobante", "La nueva factura") & vbNewLine & _
                            Replace(SIHOMsg(1307), "La factura", "La nueva factura"), vbCritical, "Mensaje"
                     '________________________________________________________________________________________________________________________________________________
                     'Regresamos la información a la factura anterior oJo

                     pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, True _
                                , SIHOMsg(1319) & vbNewLine & _
                                  Replace(SIHOMsg(1306), "El comprobante", "La nueva factura")

                     Screen.MousePointer = vbDefault
                     freBarraCFD.Visible = False
                     frmConsultaFactura.Enabled = True
                     pConsultaFacturas Trim(txtFolio.Text), 0
                     Exit Sub

                  ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then

                     EntornoSIHO.ConeccionSIHO.CommitTrans

                    '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
                     MsgBox Replace(SIHOMsg(1338), "La factura", "La nueva factura"), vbCritical + vbOKOnly, "Mensaje"


                     '________________________________________________________________________________________________________________________________________________
                     'Regresamos la información a la factura anterior oJo

                     pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, True, False, True, Replace(SIHOMsg(1338), "La factura", "La nueva factura")

                     'Actualiza PDF al cancelar facturas
                     If Not fblnGeneraComprobanteDigital(vllngCveFacturaNueva, "FA", 1, 0, "", False, True, -1) Then
                            On Error Resume Next
                     End If



                     Screen.MousePointer = vbDefault
                     freBarraCFD.Visible = False
                     frmConsultaFactura.Enabled = True
                     pConsultaFacturas Trim(txtFolio.Text), 0
                     Exit Sub
                 End If
              Else
                 pLogTimbrado 1
                 pEliminaPendientesTimbre vllngCveFacturaNueva, "FA" 'quitamos la factura de pendientes de timbre fiscal
                 EntornoSIHO.ConeccionSIHO.CommitTrans
              End If
              'Barra de progreso CFD
              pgbBarraCFD.Value = 100
              freBarraCFD.Top = 3200
              Screen.MousePointer = vbDefault
              freBarraCFD.Visible = False
              frmConsultaFactura.Enabled = True
            End If

            '----------------------------------
            'Cancelacion de la factura anterior*
            '----------------------------------
            '----------------------
            ' Inicio de Transacción
            '----------------------
             EntornoSIHO.ConeccionSIHO.BeginTrans
            '----------------------------------
            ' Obtener el numero de corte actual
            '----------------------------------
            vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
            vlblnCorteValido = True
            If vllngMensaje <> 0 Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'Que el corte debe ser cerrado por cambio de día ó Que no existe corte abierto
               vlblnCorteValido = False
               MsgBox SIHOMsg(Str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje"
             '|Cancelamos la factura nueva---------------------------------------------------------------------|
               intcontadorCFDiPendienteCancelar = 0
               ReDim vlArrCFDiPendienteCancelar(0)
               pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
               pCFDiPendienteCancelar 0, 0, 1
               pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, False, SIHOMsg(Str(vllngMensaje))
             '|________________________________________________________________________________________________|
               Exit Sub
            End If

            vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
            pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""


           '|Pone la fecha actual como fecha de cancelación en GnComprobanteFiscalDigital
            frsEjecuta_SP lngConsecutivoFactura & "|FA" & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora), "SP_GNUPDCANCELACOMPROBANTEFIS"

            vlstrSentencia = "SELECT * FROM PvDetalleCorte WHERE chrFolioDocumento = '" & Trim(txtFolio.Text) & "'" & _
                             " AND chrTipoDocumento = 'FA'"
            Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly) 'RS de consulta
            If rsDC.RecordCount > 0 Then
               Do While Not rsDC.EOF
                  '~~~~~~~~~~~~~~~~~~~  C A N C E L A C I Ó N  ~~~~~~~~~~~~~~~~~~~~~~'
                  '||  Generar registros al reves en PVDetalleCorte para cancelar  ||'
                  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                  pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsDC!chrFolioDocumento, rsDC!chrTipoDocumento, 0, (rsDC!mnyCantidadPagada * -1), _
                  False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                  1, "", ""
                  rsDC.MoveNext
               Loop
            End If

            vlstrSentencia = "SELECT DISTINCT  chrFolioDocumento, chrTipoDocumento, intFormaPago, " & _
                             " mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                             " intNumCorteDocumento " & _
                             " FROM PvDetalleCorte " & _
                             " WHERE chrFolioDocumento IN (SELECT chrFolioRecibo " & _
                                                            " FROM PvPago " & _
                                                            " WHERE chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "') " & _
                                " AND mnyCantidadPagada > 0 " & _
                                " AND chrTipoDocumento = 'RE' "

            If vgstrFacturaPacienteEmpresa = "P" Then  ' Sólo las facturas de pacientes tienen salidas de Efectivo
               vlstrSentencia = vlstrSentencia & " UNION SELECT DISTINCT  chrFolioDocumento, chrTipoDocumento, intFormaPago, " & _
                                " mnyCantidadPagada, mnyTipoCambio, intFolioCheque, " & _
                                " intNumCorteDocumento  " & _
                                " FROM PvDetalleCorte " & _
                                " WHERE chrFolioDocumento IN (SELECT chrFolioRecibo " & _
                                                                " FROM PvSalidaDinero " & _
                                                                " WHERE chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "') " & _
                                " AND mnyCantidadPagada > 0 " & _
                                " AND chrTipoDocumento = 'SD' "
            End If
            Set rsDC = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly) 'RS de consulta
            If rsDC.RecordCount > 0 Then
               Do While Not rsDC.EOF
                  '~~~~~~~~~~~~~~~~ C A N C E L A C I Ó N ~~~~~~~~~~~~~~~~'
                  '||       Reactivación de los pagos en el corte       ||'
                  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                    Set rsChequeTransCta = frsRegresaRs("SELECT PVCQTR.* FROM PVCORTECHEQUETRANSCTA PVCQTR INNER JOIN PVDETALLECORTE PVDC ON PVDC.INTCONSECUTIVO = PVCQTR.INTCONSECUTIVODETCORTE WHERE TRIM(PVDC.CHRFOLIODOCUMENTO) = '" & Trim(rsDC!chrFolioDocumento) & "' AND TRIM(PVDC.CHRTIPODOCUMENTO) = '" & Trim(rsDC!chrTipoDocumento) & "' AND PVDC.INTFORMAPAGO = " & rsDC!intFormaPago & " AND PVDC.MNYCANTIDADPAGADA = " & rsDC!mnyCantidadPagada & " AND PVDC.MNYTIPOCAMBIO = " & rsDC!MNYTIPOCAMBIO & " AND PVDC.INTFOLIOCHEQUE = " & rsDC!intfoliocheque & " AND PVDC.intNumCorteDocumento = " & rsDC!intNumCorteDocumento & " ORDER BY PVCQTR.INTCONSECUTIVODETCORTE", adLockReadOnly, adOpenForwardOnly)
                    If rsChequeTransCta.RecordCount > 0 Then
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsDC!chrFolioDocumento, rsDC!chrTipoDocumento, 0, rsDC!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                        1, "", "", False, Trim(Replace(Replace(Replace(vlRFCTemp, "-", ""), "_", ""), " ", "")), IIf(IsNull(rsChequeTransCta!CHRCLAVEBANCOSAT), "", rsChequeTransCta!CHRCLAVEBANCOSAT), IIf(IsNull(rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), "", rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), IIf(IsNull(rsChequeTransCta!VCHCUENTABANCARIA), "", rsChequeTransCta!VCHCUENTABANCARIA), IIf(IsNull(rsChequeTransCta!dtmfecha), fdtmServerFecha, rsChequeTransCta!dtmfecha)
                    Else
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsDC!chrFolioDocumento, rsDC!chrTipoDocumento, 0, rsDC!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsDC!intFormaPago, rsDC!MNYTIPOCAMBIO, rsDC!intfoliocheque, rsDC!intNumCorteDocumento, _
                        1, "", ""
                    End If
                  rsDC.MoveNext
               Loop
            End If

            '-----------------------------------------------------------------------------------------------------'
            ' Registrar en el corte movimientos, solo para Facturas de Tikets, para que se quede la venta intacta '
            '-----------------------------------------------------------------------------------------------------'
            If vgstrTipoFactura = "T" Then 'Sólo se reactivan las ventas de las facturas de tickets
               vlstrSentencia = "SELECT DISTINCT chrFolioDocumento, " & _
                                " chrTipoDocumento, intFormaPago, mnyCantidadPagada, " & _
                                " mnyTipoCambio, intFolioCheque, intNumCorteDocumento " & _
                                " FROM PvDetalleCorte " & _
                                " WHERE mnyCantidadPagada > 0 " & _
                                " AND chrTipoDocumento = 'TI' " & _
                                " AND chrFolioDocumento IN (SELECT chrFolioTicket FROM PvVentaPublico WHERE chrFolioFactura = '" & Trim(vlstrFolioDocumento) & "')"
               Set rsCorteTiKets = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
               If rsCorteTiKets.RecordCount > 0 Then
                  Do While Not rsCorteTiKets.EOF
                     '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                     '||    C A N C E L A C I Ó N     ||'
                     '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
                     Set rsChequeTransCta = frsRegresaRs("SELECT PVCQTR.* FROM PVCORTECHEQUETRANSCTA PVCQTR INNER JOIN PVDETALLECORTE PVDC ON PVDC.INTCONSECUTIVO = PVCQTR.INTCONSECUTIVODETCORTE WHERE TRIM(PVDC.CHRFOLIODOCUMENTO) = '" & Trim(rsCorteTiKets!chrFolioDocumento) & "' AND TRIM(PVDC.CHRTIPODOCUMENTO) = '" & Trim(rsCorteTiKets!chrTipoDocumento) & "' AND PVDC.INTFORMAPAGO = " & rsCorteTiKets!intFormaPago & " AND PVDC.MNYCANTIDADPAGADA = " & rsCorteTiKets!mnyCantidadPagada & " AND PVDC.MNYTIPOCAMBIO = " & rsCorteTiKets!MNYTIPOCAMBIO & " AND PVDC.INTFOLIOCHEQUE = " & rsCorteTiKets!intfoliocheque & " AND PVDC.intNumCorteDocumento = " & rsCorteTiKets!intNumCorteDocumento & " ORDER BY PVCQTR.INTCONSECUTIVODETCORTE", adLockReadOnly, adOpenForwardOnly)
                     If rsChequeTransCta.RecordCount > 0 Then
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsCorteTiKets!chrFolioDocumento, "TI", 0, rsCorteTiKets!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsCorteTiKets!intFormaPago, rsCorteTiKets!MNYTIPOCAMBIO, rsCorteTiKets!intfoliocheque, _
                        rsCorteTiKets!intNumCorteDocumento, 1, "", "", False, Trim(Replace(Replace(Replace(vlRFCTemp, "-", ""), "_", ""), " ", "")), IIf(IsNull(rsChequeTransCta!CHRCLAVEBANCOSAT), "", rsChequeTransCta!CHRCLAVEBANCOSAT), IIf(IsNull(rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), "", rsChequeTransCta!VCHBANCOORIGENEXTRANJERO), IIf(IsNull(rsChequeTransCta!VCHCUENTABANCARIA), "", rsChequeTransCta!VCHCUENTABANCARIA), IIf(IsNull(rsChequeTransCta!dtmfecha), fdtmServerFecha, rsChequeTransCta!dtmfecha)
                     Else
                        pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, rsCorteTiKets!chrFolioDocumento, "TI", 0, rsCorteTiKets!mnyCantidadPagada, _
                        False, (fdtmServerFecha + fdtmServerHora), rsCorteTiKets!intFormaPago, rsCorteTiKets!MNYTIPOCAMBIO, rsCorteTiKets!intfoliocheque, _
                        rsCorteTiKets!intNumCorteDocumento, 1, "", ""
                     End If
                     rsCorteTiKets.MoveNext
                  Loop
               End If
               rsCorteTiKets.Close
            End If

            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            '||    C A N C E L A C I Ó N    D E    F A C T U R A     ||'
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            vlstrSentencia = "SELECT chrestatus FROM PvFactura WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            Set rsTemp = frsRegresaRs(vlstrSentencia)
            If rsTemp.RecordCount > 0 Then
               If rsTemp!chrEstatus = "C" Then
                  EntornoSIHO.ConeccionSIHO.RollbackTrans
                  MsgBox SIHOMsg(1229), vbExclamation, "Mensaje"   'No se puede refacturar, el estado de la factura cambió. Consulte de nuevo.
                 '|Cancelamos la factura nueva---------------------------------------------------------------------|
                  intcontadorCFDiPendienteCancelar = 0
                  ReDim vlArrCFDiPendienteCancelar(0)
                  pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
                  pCFDiPendienteCancelar 0, 0, 1
                  pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, False, SIHOMsg(1229)
                 '|________________________________________________________________________________________________|
                  Me.MousePointer = 0
                  Exit Sub
               End If
            End If
            vlstrSentencia = "UPDATE PvFactura SET chrEstatus = 'C' WHERE chrFolioFactura = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia (vlstrSentencia)
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            '||          Cancelar el CFDi por medio del PAC          ||'
            '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
            If intTipoEmisionComprobante = 2 Then
               If Not fblnCancelaCFDi(lngConsecutivoFactura, "FA") Then
                  EntornoSIHO.ConeccionSIHO.RollbackTrans
                  If vlstrMensajeErrorCancelacionCFDi <> "" Then MsgBox vlstrMensajeErrorCancelacionCFDi, vbOKOnly + vbCritical, "Mensaje"
                 '|Cancelamos la factura nueva---------------------------------------------------------------------|
                  intcontadorCFDiPendienteCancelar = 0
                  ReDim vlArrCFDiPendienteCancelar(0)
                  pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
                  pCFDiPendienteCancelar 0, 0, 1
                  pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, False, vlstrMensajeErrorCancelacionCFDi
                 '|________________________________________________________________________________________________|
                  Screen.MousePointer = vbDefault
                  freBarraCFD.Visible = False
                  frmConsultaFactura.Enabled = True
                  Exit Sub
               End If
            End If
            '---------------------------------'
            ' Guardo en documentos cancelados '
            '---------------------------------'
            vlstrSentencia = "SELECT chrFolioDocumento FROM PvDocumentoCancelado WHERE chrFolioDocumento = '" & Trim(txtFolio.Text) & "' and smiDepartamento = " & Trim(Str(vgintNumeroDepartamento))
            Set rsTemp = frsRegresaRs(vlstrSentencia)
            If rsTemp.RecordCount > 0 Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               MsgBox SIHOMsg(1229), vbExclamation, "Mensaje"   'No se puede refacturar, el estado de la factura cambió. Consulte de nuevo.
             '|Cancelamos la factura nueva---------------------------------------------------------------------|
               intcontadorCFDiPendienteCancelar = 0
               ReDim vlArrCFDiPendienteCancelar(0)
               pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
               pCFDiPendienteCancelar 0, 0, 1
               pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, False, SIHOMsg(1229)
             '|________________________________________________________________________________________________|
               Me.MousePointer = 0
               Exit Sub
            End If
            vlstrSentencia = "INSERT INTO PVDocumentoCancelado VALUES('" & Trim(txtFolio.Text) & "','FA'," & _
                             Trim(Str(vgintNumeroDepartamento)) & "," & Trim(Str(lngPersonaGraba)) & ",getdate())"
            pEjecutaSentencia (vlstrSentencia)

            '----------------------------------------'
            ' Cancelación de la factura, SOLO POLIZA '
            '----------------------------------------'
            pAgregarMovArregloCorte vllngNumeroCorte, vllngPersonaGraba, txtFolio.Text, "", 0, 0, False, "", 0, 0, "", 0, 6, vlstrFolioDocumento, ""
            rsTemp.Close

            '---------------------------------------'
            'Actualiza en PVPAQUETEPACIENTEFACTURADO'
            '---------------------------------------'
            strParametros = txtMovimientoPaciente.Text & "|" & IIf(optTipoPaciente(2).Value, "'G'", IIf(optTipoPaciente(0).Value, "'I'", "'E'")) & "|" & Trim(txtFolio.Text)
            frsEjecuta_SP strParametros, "SP_PVUPDPAQUETESFACTURADOS"

            'insertar movimientos al corte
             vllngCorteUsado = fRegistrarMovArregloCorte(vllngNumeroCorte)
             'insertamos mov en el corte
            If vllngCorteUsado = 0 Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
               MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
             '|Cancelamos la factura nueva---------------------------------------------------------------------|
               intcontadorCFDiPendienteCancelar = 0
               ReDim vlArrCFDiPendienteCancelar(0)
               pCFDiPendienteCancelar vllngCveFacturaNueva, "FA", 0
               pCFDiPendienteCancelar 0, 0, 1
               pRegresaDatosFacturaAnterior txtFolio.Text, vlstrFolioDocumento, vglngCveFactura, vllngCveFacturaNueva, blnFacturaAutomatica, lngPersonaGraba, strTipoPacienteFactura, vlrsAnteriorFactura!INTMOVPACIENTE, vllngFormatoaUsar, intTipoDetalleFactura, vgstrFacturaPacienteEmpresa, lblnCalcularEnBaseACargos, optTipoPaciente(0).Value, optTipoPaciente(1).Value, txtMovimientoPaciente.Text, False, False, False, SIHOMsg(1320)
             '|________________________________________________________________________________________________|
              Exit Sub
            End If

            '--------------------------------------------------------------------------------------------'
            ' Cancelar los movimientos de la forma de pago de la vieja Factura y agregar los de la nueva '
            '--------------------------------------------------------------------------------------------'
            ' Agregado para caso 8758
            'Si el BitCuentaPuenteBanco está activo y la factura no es a paciente si se cancela el movimiento
            If intBitCuentaPuenteBanco = 0 Then
               pCancelaMovimientoRef lngConsecutivoFactura, Trim(txtFolio.Text), vllngNumCorteFactura, vllngNumeroCorte, lngPersonaGraba, True, vllngCveFacturaNueva
            End If

            '| Actualiza los honorarios médicos generados automáticamente por pagos de contado
            vlstrSentencia = "Update PVFACTURAHONORARIOMEDAUTOMATIC Set CHRFOLIOFACTURA = '" & vlstrFolioDocumento & "' Where CHRFOLIOFACTURA = '" & Trim(txtFolio.Text) & "'"
            pEjecutaSentencia vlstrSentencia


            EntornoSIHO.ConeccionSIHO.CommitTrans 'cerramos transacción de la cancelación de la factura

            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngPersonaGraba, "REFACTURACIÓN", txtFolio.Text)

           'La factura se canceló satisfactoriamente.
            MsgBox SIHOMsg(633) & "." & vbNewLine & SIHOMsg(343), vbInformation, "Mensaje"

            'Actualiza PDF al cancelar facturas
            If Not fblnGeneraComprobanteDigital(vllngConsecutivoFactura, "FA", 1, 0, "", False, True, -1) Then
                On Error Resume Next
            End If




           '| Facturación digital
           If intTipoEmisionComprobante = 2 Then
              If Not fblnImprimeComprobanteDigital(vllngCveFacturaNueva, "FA", "I", vllngFormatoaUsar, intTipoDetalleFactura) Then
                 Exit Sub
              End If

             '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ENVÍO DE CFD @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
             'Verifica el parámetro de envío de CFDs por correo
             If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
                '¿Desea enviar por e-mail la información del comprobante fiscal digital?
                If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                   pEnviarCFD "FA", vllngCveFacturaNueva, CLng(vgintClaveEmpresaContable), vlRFCTemp, lngPersonaGraba, Me
                End If
             End If
             vlRFCTemp = ""
             '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ FIN ENVÍO DE CFD @@@@@@@@@@@@@@@@@@@@@@@@@@@@
          Else
                '(PEMEX)
                If blnFacturaMultiple Then
                    pImprimeFormatoProc vllngFormatoaUsar, lngRenglonesDetalle, lngTotalDocumentos, rsDatosFactura, arrFolios
                Else
                    pImprimeFormato vllngFormatoaUsar, vllngCveFacturaNueva
                End If
                '(FIN PEMEX)
          End If
          pConsultaFacturas Trim(txtFolio.Text), 0
          Unload frmDatosFiscales
          Set frmDatosFiscales = Nothing
        Else
            Unload frmDatosFiscales
            Set frmDatosFiscales = Nothing
        End If
    End If
    vlrsAnteriorFactura.Close
End Sub

Private Sub cmdSiguienteFactura_Click(Index As Integer)
    If vgfrmFacturas.grdBuscaFacturas.Row < vgfrmFacturas.grdBuscaFacturas.Rows - 1 Then
        vgfrmFacturas.grdBuscaFacturas.Row = vgfrmFacturas.grdBuscaFacturas.Row + 1
        pConsultaFacturas vgfrmFacturas.grdBuscaFacturas.TextMatrix(vgfrmFacturas.grdBuscaFacturas.Row, 1), False
    End If
End Sub

''Private Sub Command1_Click()
''    frmMotivosCancelacion.Show vbModal, Me
''
''
''End Sub

Private Sub Form_Activate()
    Dim vllngMensaje As Long
    vllngMensaje = flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If vllngMensaje <> 0 Then
        'Que el corte debe ser cerrado por cambio de día
        'Que no existe corte abierto
        MsgBox SIHOMsg(Str(vllngMensaje)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If
End Sub

Private Function fblnCFD()
    Dim strSentencia As String
    Dim rsCFD As New ADODB.Recordset
    Dim rsPTimbre As New ADODB.Recordset
    
    fblnCFD = True
    
    strSentencia = "SELECT CFD.INTIDCOMPROBANTE, CFD.INTNUMEROAPROBACION ,VCHUUID" & _
                   " FROM GNCOMPROBANTEFISCALDIGITAL CFD " & _
                   " WHERE trim(CFD.VCHSERIECOMPROBANTE) || trim(CFD.VCHFOLIOCOMPROBANTE) = '" & Trim(txtFolio.Text) & "'" & _
                   " AND CHRTIPOCOMPROBANTE = 'FA'"
    Set rsCFD = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rsCFD.RecordCount > 0 Then
       vlstrTipoCFD = IIf(IsNull(rsCFD!INTNUMEROAPROBACION), "CFDi", "CFD")
    End If
    
    rsCFD.Close
    
    If vlstrTipoCFD = "" Then
       fblnCFD = False
    Else
        '--------- Activación del botón CFDi -----------
        strSentencia = "Select * from gnpendientestimbrefiscal where chrtipocomprobante = 'FA' and intcomprobante = " & lngConsecutivoFactura
        Set rsPTimbre = frsRegresaRs(strSentencia, adLockOptimistic)
        
        If rsPTimbre.RecordCount > 0 Then
           fblnCFD = False
        End If
        
        rsPTimbre.Close
    End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub
Private Function fblnBloqueoCuenta() As Boolean
On Error GoTo NotificaError
    
    Dim X As Integer
    Dim vlblnTermina As Boolean
    Dim vlstrBloqueo As String
                
    fblnBloqueoCuenta = False
    
    vlblnTermina = False
    X = 1
    Do While X <= cgintIntentoBloqueoCuenta And Not vlblnTermina
        vlstrBloqueo = fstrBloqueaCuenta(Val(txtMovimientoPaciente.Text), IIf(optTipoPaciente(0).Value, "I", "E"))
        If vlstrBloqueo = "F" Then
            vlblnTermina = True
            'La cuenta ya ha sido facturada, no se pudo realizar ningún movimiento.
            MsgBox SIHOMsg(299), vbOKOnly + vbInformation, "Mensaje"
        Else
            If vlstrBloqueo = "O" Then
                If X = cgintIntentoBloqueoCuenta Then
                    vlblnTermina = True
                    'La cuenta esta siendo usada por otra persona, intente de nuevo.
                    MsgBox SIHOMsg(300), vbOKOnly + vbInformation, "Mensaje"
                End If
            Else
                If vlstrBloqueo = "L" Then
                    vlblnTermina = True
                    fblnBloqueoCuenta = True
                End If
            End If
        End If
        X = X + 1
    Loop

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnBloqueoCuenta"))
End Function

Private Function fblnExisteTicket(prsTemp As ADODB.Recordset) As Boolean
    Dim vlintCont As Integer
    
    fblnExisteTicket = False
    For vlintCont = 0 To prsTemp.RecordCount
        If prsTemp!chrTipoDocumento = "T" Then
            fblnExisteTicket = True
            Exit Function
        End If
    Next
End Function

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    lblMoneda.BorderStyle = 0
    vlblnLicenciaIEPS = fblLicenciaIEPS '<-------
    lblnCalcularEnBaseACargos = False
    Set rs = frsEjecuta_SP(Str(vgintClaveEmpresaContable), "Sp_PvSelParametro")
    If Not rs.EOF Then
        lblnCalcularEnBaseACargos = IIf(IsNull(rs!BITCALCULARENBASEACARGOS), False, rs!BITCALCULARENBASEACARGOS)
    End If
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3088, 4114), "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 3088, 4114), "C", True) Then
        cmdCancelarFactura.Enabled = True
    Else
        cmdCancelarFactura.Enabled = False
    End If
End Sub
Private Sub pPreparaIEPS()
'Label48 Top = 300 'descuento
'Label8 = 660  'descuento especial
'Label7 = 1020 'IEPS
'Label16 = 1380 'subtotal
'Label60 = 1740 'IVa
'Label14 = 2100'Total factura
'lbPagos = 2460' pagos
'lbNotas = 2820 'notas
'lbTotalPagado = 3180 'totaltotal
'--------------------------
'txtDescuento = 240
'txtDescuentoEspecial = 600
'txtIEPS = 960
'txtSubtotal = 1320
'txtIVA = 1680
'txttotalfctura = 2040
'lblMoneda = 2100
'txtPagos = 2400
'txtNotas = 2760
'txtTotalPagado = 3120
'Frame11.Height = 6240
'frame3.top = 8925 * frame donde se encuentra el nombre de quien factura y quien cancela
'Frame2.Top = 9960 'botonera
'---------------------------
'Frame4.Height = 2700
'lblFoliosRelacinados.Height = 2415
'----------------------------------
'frmConsultaFactura.Height = 11190

If Not vlblnLicenciaIEPS Then ' NO HAY LICENCIA DE IEPS
   If Val(Format(Me.txtDescuentoEspecial.Text, "")) = 0 Then 'NO HAY DESCUENTO ESPECIAL
        Label48.Top = 300 'descuento
        Label8.Visible = False
        Label7.Visible = False
        Label16.Top = 1380 - 720 'subtotal
        Label60.Top = 1740 - 720 'IVa
        Label14.Top = 2100 - 720 'Total factura
        lbPagos.Top = 2460 - 720 ' pagos
        lbNotas.Top = 2820 - 720 'notas
        lblRetencionServ.Top = 3180 - 720 'retencion
        lbTotalPagado.Top = 3540 - 720 'totaltotal
        '--------------------------
        txtDescuentos.Top = 240
        txtDescuentoEspecial.Visible = False
        txtIEPS.Visible = False
        txtSubtotal.Top = 1320 - 720
        txtIVA.Top = 1680 - 720
        txtTotalFactura.Top = 2040 - 720
        lblMoneda.Top = 2100 - 720
        txtPagos.Top = 2400 - 720
        txtNotas.Top = 2760 - 720
        txtRetencionServ.Top = 3120 - 720
        txtTotalPagado.Top = 3480 - 720
        Frame11.Height = 3915 - 720
        Frame4.Height = 1720 - 720
        TxtObservacionesC.Height = 1920 ''*
        TxtObservacionesC.Top = 6580
        Label9.Top = 6280
        Frame3.Top = 9225 - 720 '* frame donde se encuentra el nombre de quien factura y quien cancela
        'Frame3.Top = 9225 - 360 '* frame donde se encuentra el nombre de quien factura y quien cancela
        Frame2.Top = 10160 - 720 'botonera
        '---------------------------
        'Frame4.Height = 2700 - 720
        lblFoliosRelacionados.Height = 1215 - 720
        '----------------------------------
        frmConsultaFactura.Height = 11390 - 720
   Else 'SI HAY DESCUENTO ESPECIAL*
        Label48.Top = 300 'descuento
        Label8.Top = 660  'descuento especial
        Label8.Visible = True
        Label7.Visible = False
        Label16.Top = 1380 - 360 'subtotal
        Label60.Top = 1740 - 360 'IVa
        Label14.Top = 2100 - 360 'Total factura
        lbPagos.Top = 2460 - 360 ' pagos
        lbNotas.Top = 2820 - 360 'notas
        lblRetencionServ.Top = 3180 - 360 'retencion
        lbTotalPagado.Top = 3540 - 360 'totaltotal
        '--------------------------
        txtDescuentos.Top = 240
        txtDescuentoEspecial.Top = 600
        txtDescuentoEspecial.Visible = True
        txtIEPS.Visible = False
        txtSubtotal.Top = 1320 - 360
        txtIVA.Top = 1680 - 360
        txtTotalFactura.Top = 2040 - 360
        lblMoneda.Top = 2100 - 360
        txtPagos.Top = 2400 - 360
        txtNotas.Top = 2760 - 360
        txtRetencionServ.Top = 3120 - 360
        txtTotalPagado.Top = 3480 - 360
        Frame11.Height = 3915 - 360
        Frame4.Height = 1720 - 360
        TxtObservacionesC.Top = 6580
        Label9.Top = 6280
        Frame3.Top = 9225 - 360 '* frame donde se encuentra el nombre de quien factura y quien cancela
        Frame2.Top = 10160 - 360 'botonera
        '---------------------------
        lblFoliosRelacionados.Height = 1215 - 360
        '----------------------------------
        frmConsultaFactura.Height = 11390 - 360
   End If
   Me.Refresh
Else 'SI HAY LICENCIA DE IEPS
   If Val(Format(Me.txtDescuentoEspecial.Text, "")) = 0 Then 'NO HAY DESCUENTO ESPECIAL
        Label48.Top = 300 'descuento
        Label8.Visible = False
        Label7.Top = 660 'IEPS
        Label16.Top = 1380 - 360 'subtotal
        Label60.Top = 1740 - 360 'IVa
        Label14.Top = 2100 - 360 'Total factura
        lbPagos.Top = 2460 - 360 ' pagos
        lbNotas.Top = 2820 - 360 'notas
        lblRetencionServ.Top = 3180 - 360 'retencion
        lbTotalPagado.Top = 3540 - 360 'totaltotal
        '--------------------------
        txtDescuentos.Top = 240
        txtIEPS.Top = 600
        txtDescuentoEspecial.Visible = False
        txtSubtotal.Top = 1320 - 360
        txtIVA.Top = 1680 - 360
        txtTotalFactura.Top = 2040 - 360
        lblMoneda.Top = 2100 - 360
        txtPagos.Top = 2400 - 360
        txtNotas.Top = 2760 - 360
        txtRetencionServ.Top = 3120 - 360
        txtTotalPagado.Top = 3480 - 360
        Frame11.Height = 3915 - 360
        Frame4.Height = 1720 - 360
        TxtObservacionesC.Top = 6580
        Label9.Top = 6280
        Frame3.Top = 9225 - 360 '* frame donde se encuentra el nombre de quien factura y quien cancela
        Frame2.Top = 10160 - 360 'botonera
        lblFoliosRelacionados.Height = 1215 - 360
        '----------------------------------
        frmConsultaFactura.Height = 11390 - 360
        Me.Refresh
   'Else 'SI HAY DESCUENTO ESPECIAL PERO NO SE REQUIERE AJUSTE DE PANTALLA
   End If
End If
End Sub
Private Sub cmdCFD_Click()
On Error GoTo NotificaError

    If vlstrTipoCFD = "CFD" Then
        frmComprobanteFiscalDigital.lngComprobante = lngConsecutivoFactura
        frmComprobanteFiscalDigital.strTipoComprobante = "FA"
        frmComprobanteFiscalDigital.blnCancelado = txtCanceladada.Visible
        frmComprobanteFiscalDigital.Show vbModal, Me
    ElseIf vlstrTipoCFD = "CFDi" Then
        frmComprobanteFiscalDigitalInternet.lngComprobante = lngConsecutivoFactura
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = "FA"
        frmComprobanteFiscalDigitalInternet.blnCancelado = txtCanceladada.Visible
        frmComprobanteFiscalDigitalInternet.blnFacturaSinComprobante = False
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    Else
        frmComprobanteFiscalDigitalInternet.lngComprobante = lngConsecutivoFactura
        frmComprobanteFiscalDigitalInternet.strTipoComprobante = "FA"
        frmComprobanteFiscalDigitalInternet.blnCancelado = txtCanceladada.Visible
        frmComprobanteFiscalDigitalInternet.blnFacturaSinComprobante = True
        frmComprobanteFiscalDigitalInternet.llngMovPaciente = txtMovimientoPaciente.Text
        frmComprobanteFiscalDigitalInternet.strFolioFactura = txtFolio.Text
        frmComprobanteFiscalDigitalInternet.dblTotal = txtTotalPagado.Text
        frmComprobanteFiscalDigitalInternet.Show vbModal, Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdCFD_Click"))
    Unload Me
End Sub


Private Function fblnValidacionWSAXA(vlstrFolioReceta As String, vlstrFolioDocumento As String, lngFolioLog As Long, vllngCveTipoIngreso As Long, lngnumCuenta As Long, lngNumPaciente As Long, strTipoPaciente As String, lngPersonaGraba As Long) As Boolean

    Dim DOMRequestXML As MSXML2.DOMDocument
    Dim SerializerWS As SoapSerializer30 'Para serializar el XML
    Dim ReaderRespuestaWS As SoapReader30      'Para leer la respuesta del WebService
    Dim ConectorWS As ISoapConnector 'Para conectarse al WebService
    Dim rsConexion As New ADODB.Recordset
    Dim vlaryParametrosSalida() As String
    Dim strDestino As String
    Dim objShell As Object
    Dim intMessage As Integer
    Dim rsLogInterfaz As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vllngFolioTrans As Long
    
    'Se inicializa la variable del folio
'    vllngFolioTrans = 0
    Err.Clear
    fblnValidacionWSAXA = False
    strRequestXML = ""
    strResponseXML = ""
    
    
    '_________________________________________________________________________________________________
        vlstrSentencia = "Select CLBXMLREQUEST FROM GNLOGINTERFAZAXA Where INTCVEFOLIO = " & lngFolioLog
        Set rsLogInterfaz = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        
        If rsLogInterfaz.RecordCount <> 0 Then
            strRequestXML = rsLogInterfaz!CLBXMLREQUEST
            strRequestXML = Replace(Replace(Trim(strRequestXML), Chr(10), ""), Chr(13), "") 'Se eliminan los saltos de linea
        Else
            GoTo NotificaError
        End If
    '_________________________________________________________________________________________________
    
    
        'Se especifican las rutas para la conexión con el servicio de timbrado
        Set rsConexion = frsEjecuta_SP(CStr(vgintClaveEmpresaContable), "SP_GNSELCONFIGINTERFAZWS")
    
        'Se elimina el contenido de la carpeta temporal
        strDestino = Environ$("temp") & "\fm-Axa9"
        pCreaDirectorio strDestino
        On Error Resume Next
        
        'Checa que el directorio se haya creado, para proceder a eliminar el archivo (evita el error 53)
        If Dir$(strDestino & "\") <> "" Then
            Kill strDestino & "\*.*"
            If Err.Number = 53 Or Err.Number = 0 Then '|  File not found
                Err.Clear
            Else
                Err.Raise Err.Number
            End If
        End If
        
        
        Set ConectorWS = New HttpConnector30
        ' La URL que atenderá nuestra solicitud
        ConectorWS.Property("EndPointURL") = Trim(rsConexion!URLWSConexion) '"http://www.axa-assistance-la.com:8082/wsre/eRecetario.asmx?WSDL"
                    
        ' Ruta del WebMethod según el tipo de ingreso
        ConectorWS.Property("SoapAction") = Trim(rsConexion!URLWMRecepcion)  '"http://www.axa-assistance-la.com/RecepcionServicio"
            
        '###########################################################################################################################
        '########################################## INICIA CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
        '###########################################################################################################################
        ConectorWS.Connect
    
        ConectorWS.BeginMessage
            Set SerializerWS = New SoapSerializer30
            SerializerWS.Init ConectorWS.InputStream
            
            SerializerWS.StartEnvelope
                SerializerWS.StartBody
                    SerializerWS.WriteXml strRequestXML
                SerializerWS.EndBody
            SerializerWS.EndEnvelope
            
        ConectorWS.EndMessage
        
        Set ReaderRespuestaWS = New SoapReader30
        ReaderRespuestaWS.Load ConectorWS.OutputStream
        
        strResponseXML = CStr(ReaderRespuestaWS.Body.xml)
        '#############################################################################################################################
        '########################################## FINAL DE CONEXIÓN CON EL SERVICIO DE TIMBRADO ##########################################
        '#############################################################################################################################
    
        Dim strMensajeError As String
        If Not ReaderRespuestaWS.FaultDetail Is Nothing Then
            Dim DOMNodoCodigoError As MSXML2.IXMLDOMNode
            
            'Se obtiene el codigo del error (en caso de haberlo)
            Set DOMNodoCodigoError = ReaderRespuestaWS.FaultDetail.childNodes(0).Attributes.getNamedItem("codigo")
            
            'Se captura el error
            If DOMNodoCodigoError Is Nothing Then
                Err.Raise 1000, "Comunicación AXA", "Error"
            Else
                strMensajeError = "Ocurrió un error al comunicarse con AXA" & vbNewLine & vbNewLine & _
                                            "Número de error: " & DOMNodoCodigoError.Text & vbNewLine & _
                                            "Descripción: " & ReaderRespuestaWS.FaultString.Text
                'Se muestra el mensaje de error en pantalla
                MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
            End If
        End If
    
        'Se valida si tiene un error en la conexión al WS
        If Err.Number = 5400 Or Err.Number = -2147024809 Then GoTo NotificaError
    
        'Se determina si se regresó un mensaje de error al solicitar información con AXA
        Dim DOMElementoRespuestaWS As MSXML2.IXMLDOMElement
        Dim DOMnodoAuxiliar As MSXML2.IXMLDOMNode
        Set DOMElementoRespuestaWS = ReaderRespuestaWS.Body
         
        If Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text) <> "" Then 'Si se regresó un mensaje de error...
        
            'Se almacena en el log de transacciones de la interfaz de AXA
            vgstrParametrosSP = Trim(vlstrFolioDocumento) & "|" & vgintNumeroModulo & "|" & vllngCveTipoIngreso & "|" & lngnumCuenta & "|" & lngNumPaciente & "|" & strTipoPaciente & "|" & Trim(vgstrIP) & "|" & "FA" & "|NO|" & strRequestXML & "|" & CStr(ReaderRespuestaWS.Body.xml) & "|" & Trim(vgstrEquipo) & "|" & lngPersonaGraba & "|" & lngFolioLog & "|" & Trim(vlstrFolioReceta) & "|"
            frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa"
        
            'Generación el MsgBox con hipervínculo al chat de AXA
            Set objShell = CreateObject("Wscript.Shell")
            intMessage = MsgBox("Información incorrecta: " & vbNewLine & vbNewLine & "- " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text) & vbNewLine & vbNewLine & "¿Desea abrir el chat en línea con AXA?", vbYesNo + vbExclamation, "Mensaje")
            
            'Se valida si no hay link de ayuda por medio de la asignación de un nodo auxiliar
            Set DOMnodoAuxiliar = DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda")
            
            If intMessage = vbYes And DOMnodoAuxiliar Is Nothing Then
                MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                Exit Function
            End If
            
            'Si se selecciona que sí, se abre la ventana del chat en línea con AXA
            If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 6) = "Object" Then
                MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                Exit Function
            End If
            
            If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 3) = "The" Then
                MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                Exit Function
            End If
            
            If intMessage = vbYes And Left(Trim(DOMElementoRespuestaWS.selectSingleNode("//@mensajeError").Text), 10) = "Conversion" Then
                MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                Exit Function
            End If
            
            If intMessage = vbYes And Trim(DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text) <> "" Then
                objShell.Run DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text
            ElseIf intMessage = vbYes And Trim(DOMElementoRespuestaWS.selectSingleNode("//@linkAyuda").Text) = "" Then
                MsgBox "No se proporcionó una liga de acceso al chat en línea de AXA", vbInformation, "Mensaje"
                Exit Function
            Else
                Exit Function
            End If
            
        Else
            'Se selecciona el mensaje de conexión exitosa correspondiente
            MsgBox "Conexión exitosa. " & vbNewLine & vbNewLine & "Folio de transacción: " & Trim(DOMElementoRespuestaWS.selectSingleNode("//@folioEnvio").Text), vbInformation, "Mensaje"
                            
            'Se almacena en el log de transacciones de la interfaz de AXA
'            vgstrParametrosSP = Trim(vlstrFolioDocumento) & "|" & vgintNumeroModulo & "|" & vllngCveTipoIngreso & "|" & lngNumCuenta & "|" & lngNumPaciente & "|" & strTipoPaciente & "|" & Trim(vgstrIP) & "|" & "FA" & "|SI|" & strRequestXML & "|" & strResponseXML & "|" & Trim(vgstrEquipo) & "|" & lngPersonaGraba & "|" & lngFolioLog & "|" & Trim(vlstrFolioReceta) & "|"
'            frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa"
            pCargaArreglo vlaryParametrosSalida, "|" & adDouble 'adBSTR 'adVarChar1
            vgstrParametrosSP = Trim(vlstrFolioDocumento) & "|" & vgintNumeroModulo & "|" & vllngCveTipoIngreso & "|" & lngnumCuenta & "|" & lngNumPaciente & "|" & strTipoPaciente & "|" & Trim(vgstrIP) & "|" & "FA" & "|SI|" & strRequestXML & "|" & strResponseXML & "|" & Trim(vgstrEquipo) & "|" & lngPersonaGraba & "|0|" & Trim(vlstrFolioReceta)
            frsEjecuta_SP vgstrParametrosSP, "sp_GNINSloginterfazaxa", , , vlaryParametrosSalida
            pObtieneValores vlaryParametrosSalida, 0
            
            'Se regresa variable de estatus de conexión correcta
            fblnValidacionWSAXA = True
            
        End If
        
        Set DOMElementoRespuestaWS = Nothing
        
Exit Function
NotificaError:
    strMensajeError = "Ocurrió un error al comunicarse con AXA" & vbNewLine & vbNewLine & _
                                "Verifique que el equipo cuente con acceso a Internet y/o que la información enviada a AXA esté correcta."
    MsgBox strMensajeError, vbCritical + vbOKOnly, "Mensaje"
    Err.Clear
    fblnValidacionWSAXA = False
    Exit Function
    
End Function
Private Function fblnPermiteRefacturacionCambioIVA() As Boolean
'funcion que valida que no se hagan refacturaciones si el iva cambio(colocada para validar el cambio de IVa del 2013 al 2014)
Dim ObjRs As New ADODB.Recordset
Dim strSentencia As String
Dim vlStrFechaFactura As String

fblnPermiteRefacturacionCambioIVA = True

strSentencia = "select dtmfechahora from pvfactura where chrfoliofactura='" & Trim(Me.txtFolio.Text) & "'"
Set ObjRs = frsRegresaRs(strSentencia, adLockOptimistic)

If ObjRs.RecordCount > 0 Then vlStrFechaFactura = CStr(ObjRs!dtmFechahora)

If vlStrFechaFactura <> "" Then
   If CDate(Format(vlStrFechaFactura, "dd/mm/yyyy HH:mm:ss")) <= CDate("31/12/2013 23:59:59") Then
       If Format(fdtmServerFechaHora(), "dd/mm/yyyy HH:mm:ss") > CDate("31/12/2013 23:59:59") Then
         strSentencia = "select vchvalor from siparametro where vchnombre = 'INTSECAMBIOIVA2014'"
         Set ObjRs = frsRegresaRs(strSentencia, adLockOptimistic)
         If ObjRs.RecordCount > 0 Then
            If ObjRs!vchvalor = "1" Then
               fblnPermiteRefacturacionCambioIVA = False
            End If
         End If
       End If
   End If
End If
End Function
Private Function fblnActualizaIEPS(vllngConsecutivoNF As Long, vlblnDesglosa As Boolean) As Boolean
' función para hacer la actualización del detalle de la factura cuando se requiera a causa del desglose de IEPS
Dim ObjRs As New ADODB.Recordset
Dim ObjRsVP As New ADODB.Recordset
Dim ObjRsDetalleFactura As New ADODB.Recordset
Dim ObjRsPvFactura As New ADODB.Recordset
Dim objSTR As String
Dim ArrFacturas() As Double
Dim vllngCantArrFacturas As Long
Dim vllngPosArrFacturas As Long
Dim vllngContArrFacturas As Long
Dim vldblDescuentoIndividual As Double
Dim vldblIVADescuentoIndividual As Double
Dim vldblImporteGravado As Double
Dim vldblImporteNogravado As Double
Dim vldbldescuentogravado As Double
Dim vldblDescuentoNoGravado As Double
Dim vldblSumatoria As Double
Dim vldblIEPS As Double
  
fblnActualizaIEPS = False

'********************************************************************************************************************************************************************
'guardamos las tasas del ieps tal y como estan en la factura anterior pero ahora con la nueva factura ya que sin importar si desglosa IEPS la nueva factura
'de todas maneras se deben de guardar las tasas
'********************************************************************************************************************************************************************
 objSTR = "Insert into PVIEPSCOMPROBANTE (INTCOMPROBANTE, CHRTIPOCOMPROBANTE, NUMTASAIEPS, MNYCANTIDADGRAVADA, MNYCANTIDADIEPS)" & _
          " select " & vllngConsecutivoNF & " " & Chr(34) & "INTCOMPROBANTE" & Chr(34) & ", CHRTIPOCOMPROBANTE, NUMTASAIEPS, MNYCANTIDADGRAVADA, MNYCANTIDADIEPS From PVIEPSCOMPROBANTE " & _
          " where INTCOMPROBANTE = " & lngConsecutivoFactura & " and CHRTIPOCOMPROBANTE='FA'"
pEjecutaSentencia (objSTR)
'********************************************************************************************************************************************************************

'Encontrar si la factura anterior desglosó IEPS
objSTR = "Select nvl(bitdesgloseIEPS,0) from PVfactura where intConsecutivo = " & lngConsecutivoFactura
Set ObjRs = frsRegresaRs(objSTR, adLockOptimistic)

If Val(ObjRs.Fields(0)) = 0 Then ' la factura anterior no desglosó IEPS
   If vlblnDesglosa Then ' la nueva factura si desglosa ieps
   '-----------------------------------------------------------------------------------------------------------------------------------------------------------------
      objSTR = "SELECT pvdetalleventapublico.INTCVEVENTA CveVenta, pvDetalleVentaPublico.intNumCargo NumCargo, " & _
               " pvDetalleVentaPublico.MNYPRECIO * pvDetalleVentaPublico.INTCANTIDAD IMPORTE, pvDetalleVentaPublico.MNYIVA, " & _
               " pvDetalleVentaPublico.mnyDescuento, pvDetalleVentaPublico.smiCveConceptoFacturacion Concepto, " & _
               " PvVentaPublico.INTCVEDEPARTAMENTO, PvVentaPublico.INTNUMCORTE, pvDetalleventapublico.mnyIEPS IEPS, pvdetalleventapublico.numporcentajeIEPS TASAIEPS" & _
               " FROM pvDetalleVentaPublico " & _
               " LEFT OUTER JOIN ivArticulo ON IvArticulo.intIdArticulo = pvDetalleVentaPublico.intCveCargo " & _
               " LEFT OUTER JOIN PvOtroConcepto ON PvOtroConcepto.intCveConcepto = pvDetalleVentaPublico.intCveCargo " & _
               " INNER JOIN PvVentaPublico ON PvDetalleVentaPublico.intCveVenta = PvVentaPublico.intCveVenta " & _
               " WHERE PvVentaPublico.CHRFOLIOFACTURA = '" & Trim(Me.txtFolio.Text) & "'"
      
      vldblDescuentoIndividual = 0
      Set ObjRsVP = frsRegresaRs(objSTR, adLockOptimistic)
      If ObjRsVP.RecordCount > 0 Then
      vldblIEPS = 0
          Set ObjRsPvFactura = frsRegresaRs("Select * from pvfactura where intconsecutivo =" & vllngConsecutivoNF, adLockOptimistic) 'Traemos la información de PVfactura
          
          If ObjRsPvFactura.EOF Then Exit Function
          
          vllngCantArrFacturas = 0
          
          Erase ArrFacturas
   
          ReDim ArrFacturas(6, 0)
      
          
          With ObjRsVP
               .MoveFirst
               Do While Not .EOF
                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    'DETALLE FACTURA
                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    vllngPosArrFacturas = -1
                    For vllngContArrFacturas = 1 To vllngCantArrFacturas
                        If !Concepto = ArrFacturas(1, vllngContArrFacturas) Then
                           vllngPosArrFacturas = vllngContArrFacturas
                           Exit For
                        End If
                    Next vllngContArrFacturas
                    
                    If vllngPosArrFacturas = -1 Then 'un nuevo
                       vllngCantArrFacturas = vllngCantArrFacturas + 1
                       ReDim Preserve ArrFacturas(6, vllngCantArrFacturas)
                       
                       ArrFacturas(1, vllngCantArrFacturas) = Val(!Concepto)
                       ArrFacturas(2, vllngCantArrFacturas) = Val(!Importe) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO) 'el importe entra sin la cantidad de IEPS
                       ArrFacturas(3, vllngCantArrFacturas) = Val(!MNYIVA) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(4, vllngCantArrFacturas) = Val(!MNYDESCUENTO) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(5, vllngCantArrFacturas) = Val(!IEPS)
                       ArrFacturas(6, vllngCantArrFacturas) = Val(!TASAIEPS) / 100
                    Else ' se suman
                       ArrFacturas(2, vllngPosArrFacturas) = ArrFacturas(2, vllngPosArrFacturas) + Val(!Importe) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO) 'el importe entra sin la cantidad de IEPS
                       ArrFacturas(3, vllngPosArrFacturas) = ArrFacturas(3, vllngPosArrFacturas) + Val(!MNYIVA) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(4, vllngPosArrFacturas) = ArrFacturas(4, vllngPosArrFacturas) + Val(!MNYDESCUENTO) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(5, vllngPosArrFacturas) = ArrFacturas(5, vllngPosArrFacturas) + Val(!IEPS)
                    End If
         
                    vldblDescuentoIndividual = vldblDescuentoIndividual + Val(!MNYDESCUENTO) '<---los descuentos se agregan como si fueran otro concepto de facturación
                    
                    If Val(!MNYIVA) > 0 Then '<----------------------importe gravado solo a lo que se le aplica IVA
                       vldblImporteGravado = vldblImporteGravado + (Val(!Importe) + Val(!IEPS) - Val(!MNYDESCUENTO))
                    End If
                    
                    vldblIEPS = vldblIEPS + Val(!IEPS)
               .MoveNext
               Loop
          End With
          
          Set ObjRsDetalleFactura = frsRegresaRs("Select * From PvDetalleFactura Where chrFolioFactura = ''", adLockOptimistic, adOpenDynamic)
          '-----------------------------------------------------------------------------------------------------------------------------------------------------------
          'INSERTAMOS LOS DETALLES DE LAS FACTURAS
          '-----------------------------------------------------------------------------------------------------------------------------------------------------------
           For vllngContArrFacturas = 1 To vllngCantArrFacturas
                ObjRsDetalleFactura.AddNew
                ObjRsDetalleFactura!chrfoliofactura = Trim(ObjRsPvFactura!chrfoliofactura)
                ObjRsDetalleFactura!smicveconcepto = ArrFacturas(1, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYCantidad = ArrFacturas(2, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYIVA = ArrFacturas(3, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYDESCUENTO = ArrFacturas(4, vllngContArrFacturas)
                ObjRsDetalleFactura!chrTipo = "NO"
                     
                ObjRsDetalleFactura!mnyIVAConcepto = ObjRsDetalleFactura!MNYCantidad * _
                (ObjRsDetalleFactura!MNYIVA / IIf((ObjRsDetalleFactura!MNYCantidad - ObjRsDetalleFactura!MNYDESCUENTO) > 0, (ObjRsDetalleFactura!MNYCantidad - ObjRsDetalleFactura!MNYDESCUENTO), 1))
                
                If ObjRsDetalleFactura!MNYIVA <> 0 Then
                   vldblIVADescuentoIndividual = vldblIVADescuentoIndividual + (ObjRsDetalleFactura!MNYDESCUENTO * (vgdblCantidadIvaGeneral / 100))
                End If
                
                ObjRsDetalleFactura!mnyIeps = ArrFacturas(5, vllngContArrFacturas)
                ObjRsDetalleFactura!numTasaIEPS = ArrFacturas(6, vllngContArrFacturas)
                
                ObjRsDetalleFactura.Update
           Next vllngContArrFacturas
          '---------------
          'Descuentos
          '---------------
           If vldblDescuentoIndividual > 0 Then
                ObjRsDetalleFactura.AddNew
                ObjRsDetalleFactura!chrfoliofactura = ObjRsPvFactura!chrfoliofactura
                ObjRsDetalleFactura!smicveconcepto = -2
                ObjRsDetalleFactura!MNYCantidad = vldblDescuentoIndividual
                ObjRsDetalleFactura!MNYIVA = 0
                ObjRsDetalleFactura!MNYDESCUENTO = vldblDescuentoIndividual
                ObjRsDetalleFactura!chrTipo = "DE"
                ObjRsDetalleFactura!mnyIVAConcepto = vldblIVADescuentoIndividual
                ObjRsDetalleFactura.Update
           End If
           ObjRsDetalleFactura.Close
          
          '----------------------------------------------------------------------------------------------------------------------------------------------------------
          'AJUSTAMOS LOS IMPORTES
          '----------------------------------------------------------------------------------------------------------------------------------------------------------
           vldblImporteGravado = vldblImporteGravado / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
           vldblImporteNogravado = ObjRsPvFactura!mnyTotalFactura - ObjRsPvFactura!smyIVA - vldblImporteGravado
           vldbldescuentogravado = vldblDescuentoIndividual / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
           vldblDescuentoNoGravado = ObjRsPvFactura!MNYDESCUENTO - vldbldescuentogravado
           
           pEjecutaSentencia "INSERT INTO PVFACTURAIMPORTE (INTCONSECUTIVO, MNYSUBTOTALGRAVADO, MNYSUBTOTALNOGRAVADO, MNYDESCUENTOGRAVADO, MNYDESCUENTONOGRAVADO) " & _
           "VALUES (" & ObjRsPvFactura!intConsecutivo & "," & Round(vldblImporteGravado, 2) & "," & Round(vldblImporteNogravado, 2) & "," & Round(vldbldescuentogravado, 2) & "," & Round(vldblDescuentoNoGravado, 2) & ")"
           
           fblnActualizaIEPS = True
      End If
      
   '-----------------------------------------------------------------------------------------------------------------------------------------------------------------
   Else ' la nueva factura tampoco desglosa IEPS
      Exit Function 'salimos no hay cambios
   End If
Else ' la factura anterior desglosó IEPS
   If vlblnDesglosa Then ' la nueva factura tambien desglosa IEPS
      Exit Function
   Else 'la nueva factura no desglosa IEPS
      '-----------------------------------------------------------------------------------------------------------------------------------------------------------------
      objSTR = "SELECT pvdetalleventapublico.INTCVEVENTA CveVenta, pvDetalleVentaPublico.intNumCargo NumCargo, " & _
               " pvDetalleVentaPublico.MNYPRECIO * pvDetalleVentaPublico.INTCANTIDAD IMPORTE, pvDetalleVentaPublico.MNYIVA, " & _
               " pvDetalleVentaPublico.mnyDescuento, pvDetalleVentaPublico.smiCveConceptoFacturacion Concepto, " & _
               " PvVentaPublico.INTCVEDEPARTAMENTO, PvVentaPublico.INTNUMCORTE, pvDetalleventapublico.mnyIEPS IEPS, pvdetalleventapublico.numporcentajeIEPS TASAIEPS" & _
               " FROM pvDetalleVentaPublico " & _
               " LEFT OUTER JOIN ivArticulo ON IvArticulo.intIdArticulo = pvDetalleVentaPublico.intCveCargo " & _
               " LEFT OUTER JOIN PvOtroConcepto ON PvOtroConcepto.intCveConcepto = pvDetalleVentaPublico.intCveCargo " & _
               " INNER JOIN PvVentaPublico ON PvDetalleVentaPublico.intCveVenta = PvVentaPublico.intCveVenta " & _
               " WHERE PvVentaPublico.CHRFOLIOFACTURA = '" & Trim(Me.txtFolio.Text) & "'"
      
      vldblDescuentoIndividual = 0
      Set ObjRsVP = frsRegresaRs(objSTR, adLockOptimistic)
      If ObjRsVP.RecordCount > 0 Then
         Set ObjRsPvFactura = frsRegresaRs("Select * from pvfactura where intconsecutivo =" & vllngConsecutivoNF, adLockOptimistic) 'Traemos la información de PVfactura
          If ObjRsPvFactura.EOF Then Exit Function
      
          vllngCantArrFacturas = 0
       
          Erase ArrFacturas
       
          ReDim ArrFacturas(4, 0)
     
          
          With ObjRsVP
               .MoveFirst
               Do While Not .EOF
                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    'DETALLE FACTURA
                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    vllngPosArrFacturas = -1
                    For vllngContArrFacturas = 1 To vllngCantArrFacturas
                        If !Concepto = ArrFacturas(1, vllngContArrFacturas) Then
                           vllngPosArrFacturas = vllngContArrFacturas
                           Exit For
                        End If
                    Next vllngContArrFacturas
                    
                    If vllngPosArrFacturas = -1 Then 'un nuevo
                       vllngCantArrFacturas = vllngCantArrFacturas + 1
                       ReDim Preserve ArrFacturas(4, vllngCantArrFacturas)
                       
                       ArrFacturas(1, vllngCantArrFacturas) = Val(!Concepto)
                       ArrFacturas(2, vllngCantArrFacturas) = (Val(!Importe) + Val(!IEPS)) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO) 'el importe entra CON la cantidad de IEPS
                       ArrFacturas(3, vllngCantArrFacturas) = Val(!MNYIVA) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(4, vllngCantArrFacturas) = Val(!MNYDESCUENTO) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                    Else ' se suman
                       ArrFacturas(2, vllngPosArrFacturas) = ArrFacturas(2, vllngPosArrFacturas) + ((Val(!Importe) + Val(!IEPS)) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)) 'el importe entra CON la cantidad de IEPS
                       ArrFacturas(3, vllngPosArrFacturas) = ArrFacturas(3, vllngPosArrFacturas) + Val(!MNYIVA) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                       ArrFacturas(4, vllngPosArrFacturas) = ArrFacturas(4, vllngPosArrFacturas) + Val(!MNYDESCUENTO) / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
                    End If
         
                    vldblDescuentoIndividual = vldblDescuentoIndividual + Val(!MNYDESCUENTO) '<---los descuentos se agregan como si fueran otro concepto de facturación/también se utliza para la tabla de pvfacturaimportes
                    If Val(!MNYIVA) > 0 Then
                       vldblImporteGravado = vldblImporteGravado + (Val(!Importe) + Val(!IEPS) - Val(!MNYDESCUENTO)) '<<
                    End If
                     vldblIEPS = vldblIEPS + Val(!IEPS)
                
               .MoveNext
               Loop
          End With
          
          Set ObjRsDetalleFactura = frsRegresaRs("Select * From PvDetalleFactura Where chrFolioFactura = ''", adLockOptimistic, adOpenDynamic)
          '-----------------------------------------------------------------------------------------------------------------------------------------------------------
          'INSERTAMOS LOS DETALLES DE LAS FACTURAS
          '-----------------------------------------------------------------------------------------------------------------------------------------------------------
           For vllngContArrFacturas = 1 To vllngCantArrFacturas
                ObjRsDetalleFactura.AddNew
                ObjRsDetalleFactura!chrfoliofactura = Trim(ObjRsPvFactura!chrfoliofactura)
                ObjRsDetalleFactura!smicveconcepto = ArrFacturas(1, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYCantidad = ArrFacturas(2, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYIVA = ArrFacturas(3, vllngContArrFacturas)
                ObjRsDetalleFactura!MNYDESCUENTO = ArrFacturas(4, vllngContArrFacturas)
                ObjRsDetalleFactura!chrTipo = "NO"
                     
                ObjRsDetalleFactura!mnyIVAConcepto = ObjRsDetalleFactura!MNYCantidad * _
                (ObjRsDetalleFactura!MNYIVA / IIf((ObjRsDetalleFactura!MNYCantidad - ObjRsDetalleFactura!MNYDESCUENTO) > 0, (ObjRsDetalleFactura!MNYCantidad - ObjRsDetalleFactura!MNYDESCUENTO), 1))
                
                If ObjRsDetalleFactura!MNYIVA <> 0 Then
                   vldblIVADescuentoIndividual = vldblIVADescuentoIndividual + (ObjRsDetalleFactura!MNYDESCUENTO * (vgdblCantidadIvaGeneral / 100))
                End If
                
                ObjRsDetalleFactura!mnyIeps = 0
                ObjRsDetalleFactura.Update
                
            Next vllngContArrFacturas
          '---------------
          'Descuentos
          '---------------
           If vldblDescuentoIndividual > 0 Then
                ObjRsDetalleFactura.AddNew
                ObjRsDetalleFactura!chrfoliofactura = Trim(ObjRsPvFactura!chrfoliofactura)
                ObjRsDetalleFactura!smicveconcepto = -2
                ObjRsDetalleFactura!MNYCantidad = vldblDescuentoIndividual
                ObjRsDetalleFactura!MNYIVA = 0
                ObjRsDetalleFactura!MNYDESCUENTO = vldblDescuentoIndividual
                ObjRsDetalleFactura!chrTipo = "DE"
                ObjRsDetalleFactura!mnyIVAConcepto = vldblIVADescuentoIndividual
                ObjRsDetalleFactura.Update
           End If
           ObjRsDetalleFactura.Close
          
          '----------------------------------------------------------------------------------------------------------------------------------------------------------
          'AJUSTAMOS LOS IMPORTES
          '----------------------------------------------------------------------------------------------------------------------------------------------------------
           vldblImporteGravado = vldblImporteGravado / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
           vldblImporteNogravado = ObjRsPvFactura!mnyTotalFactura - ObjRsPvFactura!smyIVA - vldblImporteGravado
           vldbldescuentogravado = vldblDescuentoIndividual / IIf(ObjRsPvFactura!BITPESOS = 1, 1, ObjRsPvFactura!MNYTIPOCAMBIO)
           vldblDescuentoNoGravado = ObjRsPvFactura!MNYDESCUENTO - vldbldescuentogravado
           
           pEjecutaSentencia "INSERT INTO PVFACTURAIMPORTE (INTCONSECUTIVO, MNYSUBTOTALGRAVADO, MNYSUBTOTALNOGRAVADO, MNYDESCUENTOGRAVADO, MNYDESCUENTONOGRAVADO) " & _
           "VALUES (" & ObjRsPvFactura!intConsecutivo & "," & Round(vldblImporteGravado, 2) & "," & Round(vldblImporteNogravado, 2) & "," & Round(vldbldescuentogravado, 2) & "," & Round(vldblDescuentoNoGravado, 2) & ")"
           
           fblnActualizaIEPS = True
      End If
   
   End If
End If
End Function
Private Function fblnAvtivasujetoIEPS() As Boolean
  Dim ObjRsPvFactura As New ADODB.Recordset
  Dim ObjRsPvFacturaVentaPublico As New ADODB.Recordset
  
  Dim objSTR As String
  
  fblnAvtivasujetoIEPS = False
  
  'If vgfrmFacturas.Name <> "frmConsultaPOS" Then Exit Function esta validación no es necesaria, que se comporte igual sin importar de donde se abre la pantalla
  
  objSTR = "Select * from Pvfactura where chrfoliofactura = '" & Trim(txtFolio.Text) & "'"
  Set ObjRsPvFactura = frsRegresaRs(objSTR, adLockOptimistic)
  
  If ObjRsPvFactura.RecordCount = 0 Then Exit Function
  
  If ObjRsPvFactura!intCveVentaPublico = 0 Then
     Exit Function
  ElseIf ObjRsPvFactura!intCveVentaPublico = -1 Then
     fblnAvtivasujetoIEPS = True
  Else
     objSTR = "Select count(*) " & Chr(34) & "FACTURAS" & Chr(34) & "from pvfactura where INTCVEVENTAPUBLICO = " & ObjRsPvFactura!intCveVentaPublico & "and CHRESTATUS <> 'C'"
     Set ObjRsPvFacturaVentaPublico = frsRegresaRs(objSTR, adLockOptimistic)
     
     If ObjRsPvFacturaVentaPublico.RecordCount > 0 Then
        If Val(ObjRsPvFacturaVentaPublico!FACTURAS) = 1 Then fblnAvtivasujetoIEPS = True
     End If
  End If
End Function

Private Function fblnLicenciaCFDIDesc(strTipo As String) As Boolean
    Dim strSql As String
    Dim strEncriptado As String
    Dim rsTemp As ADODB.Recordset
    
    fblnLicenciaCFDIDesc = False
    
    If strTipo = "D" Then
        strSql = "SELECT TRIM(REPLACE(REPLACE(REPLACE(CNEMPRESACONTABLE.VCHRFC,'-',''),'_',''),' ','')) AS RFC, TRIM(SIPARAMETRO.VCHVALOR) AS VALOR " & _
        "FROM SIPARAMETRO,CNEMPRESACONTABLE WHERE " & _
        "CNEMPRESACONTABLE.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable & " AND SIPARAMETRO.VCHNOMBRE = 'VCHLICENCIACFDISEGUROS' AND SIPARAMETRO.INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable
        Set rsTemp = frsRegresaRs(strSql)
        If Not rsTemp.EOF Then
            strEncriptado = fstrEncrypt(rsTemp!RFC, "SUMASEGCFDI33")
            fblnLicenciaCFDIDesc = IIf(rsTemp!valor = strEncriptado, True, False)
        End If
    Else
        fblnLicenciaCFDIDesc = True
    End If
    If Not fblnLicenciaCFDIDesc Then
        ' No se adquirió licencia para emitir CFDi para esta configuración de aseguradoras
        MsgBox SIHOMsg(1087) & " para esta configuración de aseguradoras", vbCritical, "Mensaje"
    End If
End Function

