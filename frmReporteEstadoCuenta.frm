VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReporteEstadoCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de cuenta del paciente"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   12870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNotasInternas 
      Height          =   2475
      Left            =   7750
      MaxLength       =   800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      ToolTipText     =   "Nota interna del paciente"
      Top             =   4040
      Width           =   5000
   End
   Begin VB.Frame fraFTP 
      Caption         =   "Validar con aseguradora"
      Height          =   2055
      Left            =   7750
      TabIndex        =   61
      Top             =   0
      Width           =   5000
      Begin VB.CheckBox chkCataCargosEmpresa 
         Caption         =   "Usar catálogo de cargos por empresa para códigos y descripciones"
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   43
         ToolTipText     =   "Usar catálogo de cargos por empresa para códigos y descripciones"
         Top             =   1100
         Width           =   2655
      End
      Begin VB.OptionButton optTipoValidacion 
         Caption         =   "Definitiva"
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   40
         Top             =   1450
         Width           =   975
      End
      Begin VB.OptionButton optTipoValidacion 
         Caption         =   "Consulta"
         Height          =   195
         Index           =   0
         Left            =   3000
         TabIndex        =   39
         Top             =   1150
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CommandButton cmdFTP 
         Caption         =   "Enviar"
         Height          =   375
         Left            =   4100
         TabIndex        =   44
         ToolTipText     =   "Enviar estado de cuenta"
         Top             =   1300
         Width           =   735
      End
      Begin VB.Frame frmFiltroEnvio 
         Caption         =   "Filtro por "
         Height          =   555
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   3015
         Begin VB.CheckBox chkFiltroEnvioFactura 
            Caption         =   "Factura"
            Height          =   255
            Left            =   1920
            TabIndex        =   42
            ToolTipText     =   "Filtrar los cargos por el folio de factura"
            Top             =   210
            Width           =   850
         End
         Begin VB.CheckBox chkFiltroEnvioRango 
            Caption         =   "Rango de fechas"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Filtrar los cargos por el rango de fechas"
            Top             =   210
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox txtMensaje 
      Height          =   1380
      Left            =   7750
      MaxLength       =   800
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   2325
      Width           =   5000
   End
   Begin VB.Frame Frame5 
      Caption         =   "Filtros"
      Height          =   4410
      Left            =   2160
      TabIndex        =   55
      Top             =   2085
      Width           =   5490
      Begin VB.ComboBox cboFactura 
         Height          =   315
         Left            =   2745
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   69
         ToolTipText     =   "Selección de la factura"
         Top             =   1635
         Width           =   1965
      End
      Begin VB.ComboBox cboCartas 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   34
         ToolTipText     =   "Selección de la carta"
         Top             =   3900
         Width           =   4155
      End
      Begin VB.CheckBox chkMostrarConSeguroSF 
         Caption         =   "Mostrar conceptos de seguro sin facturar"
         Enabled         =   0   'False
         Height          =   355
         Left            =   2745
         TabIndex        =   33
         ToolTipText     =   "Mostrar conceptos de seguro que no se han facturado"
         Top             =   3480
         Width           =   2600
      End
      Begin VB.CheckBox chkMostrarCirugias 
         Caption         =   "Mostrar cirugías del paciente"
         Height          =   195
         Left            =   2745
         TabIndex        =   32
         ToolTipText     =   "Mostrar cirugías del paciente"
         Top             =   3240
         Width           =   2610
      End
      Begin VB.CheckBox chkMostrarCuatroDecimales 
         Caption         =   "Mostrar importes con 4 decimales"
         Height          =   195
         Left            =   2745
         TabIndex        =   31
         ToolTipText     =   "Mostrar importes con 4 decimales"
         Top             =   3000
         Width           =   2650
      End
      Begin VB.CheckBox chkDesglosarPaquete 
         Caption         =   "Desglosar contenido del paquete"
         Height          =   195
         Left            =   2745
         TabIndex        =   30
         ToolTipText     =   "Desglosar el contenido del paquete"
         Top             =   2760
         Width           =   2640
      End
      Begin VB.CheckBox chkHora 
         Caption         =   "Incluir hora"
         Height          =   195
         Left            =   2745
         TabIndex        =   27
         ToolTipText     =   "Incluir hora "
         Top             =   2040
         Width           =   1800
      End
      Begin VB.CheckBox chkRangoFechas 
         Caption         =   "Rango de fechas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Por rango de fechas"
         Top             =   225
         Width           =   1530
      End
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   705
         TabIndex        =   16
         ToolTipText     =   "Fecha inicial"
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   840
         Left            =   3645
         TabIndex        =   58
         Top             =   75
         Visible         =   0   'False
         Width           =   1800
         Begin VB.OptionButton optCargo 
            Caption         =   "Cargos excluídos"
            Height          =   240
            Index           =   2
            Left            =   15
            TabIndex        =   18
            Top             =   105
            Width           =   1740
         End
         Begin VB.OptionButton optCargo 
            Caption         =   "Cargos no excluídos"
            Height          =   240
            Index           =   1
            Left            =   15
            TabIndex        =   20
            Top             =   570
            Width           =   1770
         End
         Begin VB.OptionButton optCargo 
            Caption         =   "Todos los cargos"
            Height          =   240
            Index           =   0
            Left            =   15
            TabIndex        =   19
            Top             =   315
            Width           =   1755
         End
      End
      Begin VB.CheckBox chkPagos 
         Caption         =   "Incluir detalle de pagos"
         Height          =   195
         Left            =   2745
         TabIndex        =   28
         ToolTipText     =   "Incluir detalle de pagos"
         Top             =   2280
         Width           =   2460
      End
      Begin VB.CheckBox chkCosto 
         Caption         =   "Incluir costos"
         Height          =   195
         Left            =   2745
         TabIndex        =   29
         ToolTipText     =   "Incluir costos"
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Frame FraTipos 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton optEstadoCuenta 
            Caption         =   "Estado de cuenta del paciente"
            Height          =   495
            Index           =   0
            Left            =   25
            TabIndex        =   21
            ToolTipText     =   "Estado de cuenta del paciente"
            Top             =   120
            Value           =   -1  'True
            Width           =   2445
         End
         Begin VB.OptionButton optEstadoCuenta 
            Caption         =   "Estado de cuenta de la empresa"
            Height          =   495
            Index           =   1
            Left            =   25
            TabIndex        =   22
            ToolTipText     =   "Estado de cuenta de la empresa"
            Top             =   675
            Width           =   2445
         End
         Begin VB.OptionButton optEstadoCuenta 
            Caption         =   "Estado de cuenta Paciente / Empresa"
            Height          =   495
            Index           =   2
            Left            =   25
            TabIndex        =   23
            ToolTipText     =   "Estado de cuenta paciente / empresa"
            Top             =   1200
            Width           =   2445
         End
      End
      Begin VB.Frame FraTiposHospitalFarmacia 
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   80
         TabIndex        =   67
         Top             =   1320
         Width           =   2535
         Begin VB.OptionButton optEstadoCuentaHospitalFarm 
            Caption         =   "Estado de cuenta consolidado"
            Height          =   495
            Index           =   2
            Left            =   25
            TabIndex        =   26
            ToolTipText     =   "Estado de cuenta consolidado"
            Top             =   1080
            Width           =   2445
         End
         Begin VB.OptionButton optEstadoCuentaHospitalFarm 
            Caption         =   "Estado de cuenta hospital"
            Height          =   495
            Index           =   1
            Left            =   25
            TabIndex        =   25
            ToolTipText     =   "Estado de cuenta hospital"
            Top             =   600
            Width           =   2445
         End
         Begin VB.OptionButton optEstadoCuentaHospitalFarm 
            Caption         =   "Estado de cuenta farmacia"
            Height          =   495
            Index           =   0
            Left            =   25
            TabIndex        =   24
            ToolTipText     =   "Estado de cuenta farmacia"
            Top             =   240
            Width           =   2445
         End
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   2505
         TabIndex        =   17
         ToolTipText     =   "Fecha final"
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblcarta 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         Height          =   195
         Left            =   120
         TabIndex        =   68
         Top             =   3960
         Width           =   375
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   1980
         TabIndex        =   60
         Top             =   585
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   135
         TabIndex        =   59
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         Height          =   195
         Left            =   2745
         TabIndex        =   56
         Top             =   1365
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Agrupación del reporte"
      Height          =   4410
      Left            =   105
      TabIndex        =   54
      Top             =   2085
      Width           =   1995
      Begin VB.OptionButton optOrden 
         Caption         =   "Cargos agrupados por fecha y concepto"
         Height          =   675
         Index           =   6
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Agrupar por cargos agrupados por fecha y concepto"
         Top             =   3600
         Width           =   1680
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Cargos agrupados"
         Height          =   315
         Index           =   5
         Left            =   100
         TabIndex        =   13
         ToolTipText     =   "Agrupar por cargos agrupados"
         Top             =   3120
         Width           =   1680
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Departamento que realizó el cargo"
         Height          =   435
         Index           =   4
         Left            =   100
         TabIndex        =   9
         ToolTipText     =   "Agrupar por departamento"
         Top             =   945
         Width           =   1680
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Cargo"
         Height          =   315
         Index           =   3
         Left            =   100
         TabIndex        =   12
         ToolTipText     =   "Agrupar por cargo"
         Top             =   2640
         Width           =   1680
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Concepto de factura"
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Agrupar por concepto de factura"
         Top             =   1560
         Width           =   1770
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Fecha"
         Height          =   315
         Index           =   1
         Left            =   100
         TabIndex        =   11
         ToolTipText     =   "Agrupar por fecha"
         Top             =   2110
         Width           =   1680
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Departamento"
         Height          =   315
         Index           =   2
         Left            =   100
         TabIndex        =   8
         ToolTipText     =   "Agrupar por departamento"
         Top             =   450
         Value           =   -1  'True
         Width           =   1680
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   5865
      TabIndex        =   53
      Top             =   6600
      Width           =   1140
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         Picture         =   "frmReporteEstadoCuenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         Picture         =   "frmReporteEstadoCuenta.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame fraDatosPaciente 
      Height          =   2055
      Left            =   105
      TabIndex        =   45
      Top             =   0
      Width           =   7530
      Begin VB.TextBox txtMovimientoPacienteOtro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   6210
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   64
         ToolTipText     =   "Número de cuenta"
         Top             =   240
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2760
         TabIndex        =   46
         Top             =   180
         Width           =   1815
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Interno"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   125
            Width           =   825
         End
         Begin VB.OptionButton optTipoPaciente 
            Caption         =   "Externo"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   2
            Top             =   125
            Width           =   855
         End
      End
      Begin VB.TextBox txtFechaFinal 
         Height          =   315
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Fecha final de atención"
         Top             =   1605
         Width           =   1710
      End
      Begin VB.TextBox txtFechaInicial 
         Height          =   315
         Left            =   1520
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Fecha de inicio de atención"
         Top             =   1605
         Width           =   1710
      End
      Begin VB.TextBox txtTipoPaciente 
         Height          =   315
         Left            =   1520
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Tipo de paciente"
         Top             =   1260
         Width           =   5860
      End
      Begin VB.TextBox txtEmpresaPaciente 
         Height          =   315
         Left            =   1520
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Nombre de la empresa del paciente"
         Top             =   915
         Width           =   5860
      End
      Begin VB.TextBox txtPaciente 
         Height          =   315
         Left            =   1520
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "Nombre del paciente"
         Top             =   585
         Width           =   5860
      End
      Begin VB.TextBox txtMovimientoPaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1520
         MaxLength       =   9
         TabIndex        =   0
         ToolTipText     =   "Número de cuenta del paciente"
         Top             =   240
         Width           =   1170
      End
      Begin VB.Label lblCuentaDe 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta de _______"
         Height          =   195
         Left            =   4785
         TabIndex        =   65
         Top             =   300
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "a"
         Height          =   195
         Left            =   3360
         TabIndex        =   52
         Top             =   1665
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de atención"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   1665
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   975
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   645
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Número de cuenta"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.Label lblNotasInternas 
      AutoSize        =   -1  'True
      Caption         =   "Notas internas"
      Height          =   195
      Left            =   7750
      TabIndex        =   62
      Top             =   3800
      Width           =   1020
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones"
      Height          =   195
      Left            =   7750
      TabIndex        =   57
      Top             =   2085
      Width           =   1065
   End
End
Attribute VB_Name = "frmReporteEstadoCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------
'Estado de cuenta del paciente
'Fecha de programación: Miércoles 18 de Abril de 2001
'--------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
' Fecha:        9/Octubre/2002
' Autor:        Rodolfo Ramos García
' Descripción:  Se le puso el filtro para que no contabilizara los pagos de Tipo Ded y Coa en el total a pagar
'--------------------------------------------------------------------------------------

Public llngNumeroCuenta As Long     'Para cuando se manda llamar pasándole los datos del paciente
Public lstrTipoPaciente As String   'Para cuando se manda llamar pasándole los datos del paciente, I = interno, E = externo
Public lstrNombreForma As String    'Para cuando se manda llamar pasándole el nombre de la forma que lo llama
Public llngNumeroCarta As Long      'Para cuando se manda llamar pasándole los datos de la carta de seguro

Private vgrptReporte1 As CRAXDRT.Report
Private vgrptReporte2 As CRAXDRT.Report
Private vgrptReporte3 As CRAXDRT.Report
Private vgrptReporte4 As CRAXDRT.Report
Private vgrptReporte5 As CRAXDRT.Report

Dim vlblnLimpiar As Boolean
Dim vlblnEntrando As Boolean
Dim vlblnUtilizaConvenio As Boolean
Dim vlblnPacienteSeleccionado As Boolean
Dim gintAseguradora As Integer                  'Indica si la empresa del paciente es aseguradora

Dim vlstrx As String

Dim vldblPagos As Double
Dim vldbltotal As Double
Dim vldblSubtotal As Double
Dim vldblDescuento As Double
Dim vldblTotalFacturado As Double
Dim vldblTotalSinFacturar As Double
Dim lintCveEmpresaPaciente As Integer
Dim rsFacturas As New ADODB.Recordset
Dim llngCveEmpresaPCE As Long
Dim vgempresapaciente1 As String

Dim vlblnDesglosarPaquete As Boolean
Dim vldblRetencion As Double
Dim strFechaNota As String
Dim strNotasInternas As String
Dim blnDatosCuenta As Boolean
Dim blnCartaEncontrada As Boolean


Dim blnEsHospitalMultiempresaFarm As Boolean 'Variable para indicar que se trata de la empresa definida como Hospital en la funcionalidad de Farmacia Multiempresa
Dim blnEsFarmaciaMultiempresaFarm As Boolean 'Variable para indicar que se trata de la empresa definida como Farmacia en la funcionalidad de Farmacia Multiempresa
Public vglngNumeroOpcion As Long
Dim blnCostos As Boolean

Dim vlintCveInterfazATC As Integer
Dim indiceCbo As Long
Dim blnEstadoAgrupado As Boolean

Private Sub pCalculaConceptos(vlblnFacturado As Boolean, vllngClaveCarta As Long)
On Error GoTo NotificaError

    If vlblnFacturado = 0 Or vlblnFacturado = False Then
        frmFacturacion.txtMovimientoPaciente = Trim(Me.txtMovimientoPaciente.Text)
        frmFacturacion.OptTipoPaciente(0) = Me.OptTipoPaciente(0)
        frmFacturacion.OptTipoPaciente(1) = Me.OptTipoPaciente(1)
        frmFacturacion.claveCartaEdoCta = vllngClaveCarta
        frmFacturacion.pMovimientoPaciente
        frmFacturacion.pCancelar

    Else
       frmFacturacion.pConfiguraGridFacturaPaciente
    End If
    
    ' -- Inserta en tabla intermedia los conceptos de aseguradora
    pInsertaConceptos

Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCalculaConceptos"))
End Sub



Private Sub pInsertaConceptos()
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    
    
 '     -- Inserta en tabla intermedia los conceptos de aseguradora para que sean tomados en cuenta en el procedimiento que muestra el estado de cuenta
    With frmFacturacion.grdFacturaPaciente
        If .Rows > 1 Then
            For vlintcontador = 1 To .Rows - 1
                If .TextMatrix(vlintcontador, 7) = "EX" Or .TextMatrix(vlintcontador, 7) = "DE" Or .TextMatrix(vlintcontador, 7) = "CO" Or .TextMatrix(vlintcontador, 7) = "CA" Or .TextMatrix(vlintcontador, 7) = "CP" Or .TextMatrix(vlintcontador, 7) = "CM" Then
                    If Val(Format(.TextMatrix(vlintcontador, 9), "############.00####")) > 0 Then
                        vlstrSentencia = "insert into PvTmpRptConceptosSeguros (campo4, descuento, ivacargo, montofacturado, ConceptoFactura) " & _
                                "values ('" & .TextMatrix(vlintcontador, 1) & "', " & Format(.TextMatrix(vlintcontador, 5), "############.00####") & ", " & .TextMatrix(vlintcontador, 4) & ", " & .TextMatrix(vlintcontador, 9) & ", " & .RowData(vlintcontador) & ")"
                                    
                        pEjecutaSentencia vlstrSentencia
                    End If
                End If
            Next vlintcontador
        End If
    End With



Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInsertaConceptos"))
End Sub

Private Sub cboCartas_Click()
    Dim querycartas As String
    Dim rsCartas As New ADODB.Recordset
    If blnCartaEncontrada = True And cboCartas.ListIndex > -1 Then
        querycartas = "select pvcartacontrolseguro.intcveempresa, ccempresa.VCHDESCRIPCION Descripcion, pvcartacontrolseguro.intcveempresa ClaveEmpresa from pvcartacontrolseguro inner join ccempresa on ccempresa.intcveempresa = pvcartacontrolseguro.intcveempresa where pvcartacontrolseguro.INTCVECARTA = '" & cboCartas.ItemData(cboCartas.ListIndex) & "'"
        Set rsCartas = frsRegresaRs(querycartas, adLockOptimistic, adOpenDynamic)
        If rsCartas.RecordCount > 0 Then
            txtEmpresaPaciente.Text = IIf(IsNull(rsCartas!Descripcion), "", rsCartas!Descripcion)
        End If
    End If
End Sub

Private Sub cboCartas_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim querycartas As String
    Dim rsCartas As New ADODB.Recordset
    Dim rsBitFTP As New ADODB.Recordset
    If blnCartaEncontrada = True Then
        If KeyCode = 13 Then
            querycartas = "select pvcartacontrolseguro.intcveempresa, ccempresa.VCHDESCRIPCION Descripcion, pvcartacontrolseguro.intcveempresa ClaveEmpresa from pvcartacontrolseguro inner join ccempresa on ccempresa.intcveempresa = pvcartacontrolseguro.intcveempresa where pvcartacontrolseguro.INTCVECARTA = '" & cboCartas.ItemData(cboCartas.ListIndex) & "'"
            Set rsCartas = frsRegresaRs(querycartas, adLockOptimistic, adOpenDynamic)
            If rsCartas.RecordCount > 0 Then
                pInterfazFTP rsCartas!claveEmpresa
                txtEmpresaPaciente.Text = IIf(IsNull(rsCartas!Descripcion), "", rsCartas!Descripcion)
            End If
            txtMensaje.SetFocus
            
        End If
    End If
End Sub

Private Sub cboCartas_LostFocus()
   Dim querycartas As String
    Dim rsCartas As New ADODB.Recordset
    If blnCartaEncontrada = True Then
            querycartas = "select pvcartacontrolseguro.intcveempresa, ccempresa.VCHDESCRIPCION Descripcion, pvcartacontrolseguro.intcveempresa ClaveEmpresa from pvcartacontrolseguro inner join ccempresa on ccempresa.intcveempresa = pvcartacontrolseguro.intcveempresa where pvcartacontrolseguro.INTCVECARTA = '" & cboCartas.ItemData(cboCartas.ListIndex) & "'"
            Set rsCartas = frsRegresaRs(querycartas, adLockOptimistic, adOpenDynamic)
            If rsCartas.RecordCount > 0 Then
                pInterfazFTP rsCartas!claveEmpresa
                txtEmpresaPaciente.Text = IIf(IsNull(rsCartas!Descripcion), "", rsCartas!Descripcion)
            End If
    End If
End Sub

Private Sub cboFactura_Click()
    If gintAseguradora = 1 And optEstadoCuenta(0).Value And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then
        chkMostrarConSeguroSF.Enabled = True
    Else
        chkMostrarConSeguroSF.Value = 0
        chkMostrarConSeguroSF.Enabled = False
    End If
End Sub

Private Sub cboFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.chkHora.SetFocus
    End If
End Sub

Private Sub chkCosto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkDesglosarPaquete.Enabled Then
            chkDesglosarPaquete.SetFocus
        Else
            If chkMostrarCuatroDecimales.Enabled And chkMostrarCuatroDecimales.Visible Then
                chkMostrarCuatroDecimales.SetFocus
            Else
                chkMostrarCirugias.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chkDesglosarPaquete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If chkMostrarCuatroDecimales.Enabled And chkMostrarCuatroDecimales.Visible Then
            chkMostrarCuatroDecimales.SetFocus
        Else
            chkMostrarCirugias.SetFocus
        End If
    End If
End Sub

Private Sub chkFiltroEnvioFactura_Click()
    If chkFiltroEnvioFactura.Value = 1 Then
        chkFiltroEnvioRango.Value = 0
    End If
End Sub

Private Sub chkFiltroEnvioFactura_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkFiltroEnvioRango_Click()
    If chkFiltroEnvioRango.Value = 1 Then
        chkFiltroEnvioFactura.Value = 0
    End If
End Sub

Private Sub chkFiltroEnvioRango_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkCataCargosEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkHora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub



Private Sub chkMostrarCirugias_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub






Private Sub chkMostrarConSeguroSF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fblnCanFocus(Me.txtMensaje) Then
            Me.txtMensaje.SetFocus
        End If
    End If
End Sub

Private Sub chkMostrarCuatroDecimales_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub chkPagos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If fblnCanFocus(chkCosto) Then
            chkCosto.SetFocus
        Else
            cmdPreview.SetFocus
        End If
    End If
End Sub

Private Sub chkRangoFechas_Click()
    If chkRangoFechas.Value = vbChecked Then
        Label10.Enabled = True
        Label11.Enabled = True
        mskFechaFin.Enabled = True
        mskFechaInicio.Enabled = True
    Else
        Label10.Enabled = False
        Label11.Enabled = False
        mskFechaFin.Enabled = False
        mskFechaInicio.Enabled = False
    End If
End Sub

Private Sub chkRangoFechas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub cmdFTP_Click()
    Dim lngRet As Long
    Dim rsChecaCartas As New ADODB.Recordset
    Dim vlintbitcartas As Integer
    
    pEjecutaSentencia "update PVFTPESTADOCUENTA set BITCARGOSPOREMPRESA = " & chkCataCargosEmpresa.Value & " where intcveinterfaz = " & vlintCveInterfazATC
    lngRet = 1
    Set rsChecaCartas = frsRegresaRs("SELECT * from all_source where (name = 'FN_PVFTPESTADOCUENTAFILTROS' OR name = 'FN_PVFTPESTADOCUENTA') and text LIKE '%IN_INTCVECARTASEGURO%'", adLockOptimistic, adOpenDynamic)
      
    If chkFiltroEnvioRango.Value = 1 Or chkFiltroEnvioFactura.Value = 1 Then
        If rsChecaCartas.RecordCount = 0 Then
            frsEjecuta_SP "-1|" & IIf(optTipoValidacion(0), "C", "V") & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & "|" & IIf(chkFiltroEnvioFactura.Value <> vbChecked, "*", IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex)))) & "|" & IIf(chkFiltroEnvioRango.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) & "|" & IIf(chkFiltroEnvioRango.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")), "fn_PVFTPEstadoCuentafiltros", True, lngRet
        Else
            frsEjecuta_SP "-1|" & IIf(optTipoValidacion(0), "C", "V") & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & "|" & IIf(chkFiltroEnvioFactura.Value <> vbChecked, "*", IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex)))) & "|" & IIf(chkFiltroEnvioRango.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) & "|" & IIf(chkFiltroEnvioRango.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")) & "|" & cboCartas.ItemData(cboCartas.ListIndex), "fn_PVFTPEstadoCuentafiltros", True, lngRet
        End If
    Else
        If rsChecaCartas.RecordCount = 0 Then
            frsEjecuta_SP "-1|" & IIf(optTipoValidacion(0), "C", "V") & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'"), "fn_PVFTPEstadoCuenta", True, lngRet
        Else
            frsEjecuta_SP "-1|" & IIf(optTipoValidacion(0), "C", "V") & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & "|" & cboCartas.ItemData(cboCartas.ListIndex), "fn_PVFTPEstadoCuenta", True, lngRet
        End If
    End If
    
    If lngRet = 0 Then
        MsgBox SIHOMsg(420), vbInformation, "Mensaje"
    Else
        MsgBox SIHOMsg(1), vbExclamation, "Mensaje"
    End If
End Sub

Private Sub cmdPreview_Click()
'1     On Error GoTo NotificaError
           
          'Validar el rango de fechas
2         If IsDate(mskFechaFin) And IsDate(mskFechaInicio) Then
3             If CDate(mskFechaFin) < CDate(mskFechaInicio) Then
4                     MsgBox SIHOMsg(379), vbExclamation, "Mensaje"
5                     pEnfocaMkTexto mskFechaFin
6                 Exit Sub
7             End If
8         End If
          
9         If FraTipos.Visible Then
10            If optEstadoCuenta(0).Enabled And optEstadoCuenta(0).Value = True And FraTipos.Enabled = True Then
11                optEstadoCuenta(0).SetFocus
12            Else
13                If optEstadoCuenta(1).Enabled And optEstadoCuenta(1).Value = True And FraTipos.Enabled = True Then
14                    optEstadoCuenta(1).SetFocus
15                Else
16                    If optEstadoCuenta(2).Enabled And optEstadoCuenta(2).Value = True And FraTipos.Enabled = True Then
17                        optEstadoCuenta(2).SetFocus
18                    End If
19                End If
20            End If
21        End If
          
22        If FraTiposHospitalFarmacia.Visible Then
23            If optEstadoCuentaHospitalFarm(0).Enabled And optEstadoCuentaHospitalFarm(0).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
24                optEstadoCuentaHospitalFarm(0).SetFocus
25            Else
26                If optEstadoCuentaHospitalFarm(1).Enabled And optEstadoCuentaHospitalFarm(1).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
27                    optEstadoCuentaHospitalFarm(1).SetFocus
28                Else
29                    If optEstadoCuentaHospitalFarm(2).Enabled And optEstadoCuentaHospitalFarm(2).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
30                        optEstadoCuentaHospitalFarm(2).SetFocus
31                    End If
32                End If
33            End If
34        End If
          
35        cmdPreview.Enabled = False
36        cmdPrint.Enabled = False
37        pReporte "P"
          
38        cmdPrint.Enabled = True
39        cmdPreview.Enabled = True
          
40        cmdPreview.SetFocus
          
41    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click" & " Linea:" & Erl()))
End Sub

Private Sub pReporte(vlstrDestino As String)
1     On Error GoTo NotificaError

          Dim alstrParametros(45) As String
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
          Dim rsTemp As New ADODB.Recordset
          Dim lblbActivaResponsableAD As Boolean
          'se agregaron estas variables para quitar los meses de la edad - caso 20431
          Dim vlstrEdadConMeses As String
          Dim vlstrEdadSinMeses As String
          Dim vlIntCont As Integer
          
    '*****************caso 20015
            Set rsTemp = frsRegresaRs("SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR " & _
                         "FROM SIPARAMETRO WHERE SIPARAMETRO.VCHNOMBRE = 'BITACTIVARESPONSABLECTAPACIENTE' AND SIPARAMETRO.CHRMODULO='PV'")
           lblbActivaResponsableAD = False
           If rsTemp.RecordCount <> 0 Then
            lblbActivaResponsableAD = rsTemp!Valor
           End If
     '*************************************
          
2         If vlblnPacienteSeleccionado Then
3             Me.MousePointer = 11
                      
4             Set rsDatosPaciente = frsEjecuta_SP(Trim(txtMovimientoPaciente.Text) & "|0|" & IIf(OptTipoPaciente(0).Value, "I", "E") & "|" & vgintClaveEmpresaContable, "Sp_PvSelDatosPaciente")
5             If rsDatosPaciente.RecordCount <> 0 Then
6                 pEjecutaSentencia "DELETE PvMensajeEstadoCuenta WHERE intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & " AND chrTipoPaciente = " & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
7                  pEjecutaSentencia "INSERT INTO pvMensajeEstadoCuenta VALUES(" & Trim(txtMovimientoPaciente.Text) & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ",'" & Trim(txtMensaje.Text) & "','" & IIf(Len(Trim(TxtNotasInternas.Text)) > 11, TxtNotasInternas.Text, "") & "')"
                  'Valida que las fechas de los filtros sean correctas
8                 If Not IsDate(mskFechaFin.Text) Then
9                     MsgBox SIHOMsg(29), vbInformation, "Mensaje"
10                    pEnfocaMkTexto mskFechaFin
11                    Exit Sub
12                End If
                  
13                If Not IsDate(mskFechaInicio.Text) Then
14                    MsgBox SIHOMsg(29), vbInformation, "Mensaje"
15                    pEnfocaMkTexto mskFechaInicio
16                    Exit Sub
17                End If
                  
18                vgrptReporte1.DiscardSavedData
                  
19                pEjecutaSentencia "DELETE PvTmpRptConceptosSeguros"

                  If Not (optOrden(5).Value) Or optOrden(6).Value Then
                      ' -- Si la cuenta del paciente es de convenio y la empresa de convenio es de tipo aseguradora,
                      ' -- se obtiene el cálculo de los conceptos de aseguradora que no han sido facturados
                      ' -- mediante los procesos que se utilizan para calcularlos al momento de facturarlos
                    If gintAseguradora = 1 And chkMostrarConSeguroSF.Value Then
                        If lstrNombreForma <> "frmFacturacion" Then
                            If OptTipoPaciente(0).Value Then 'Internos
                                vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & str(vgintClaveEmpresaContable)
                                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELINTERNOFACTURA")
                            Else  'Externos
                                 vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & str(vgintClaveEmpresaContable)
                                 Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXTERNOFACTURA")
                            End If
                          
                            If rs.RecordCount <> 0 Then
                                If cboCartas.List(cboCartas.ListIndex) = "<TODOS>" Then
                                    For G = 0 To cboCartas.ListCount - 1
                                        If cboCartas.List(G) <> "<TODOS>" Then
                                            pCalculaConceptos rs!Facturado, cboCartas.ItemData(G)
                                        End If
                                    Next
                                Else
                                    If cboCartas.ListIndex = -1 Then
                                        pCalculaConceptos rs!Facturado, 0
                                    Else
                                        pCalculaConceptos rs!Facturado, cboCartas.ItemData(cboCartas.ListIndex)
                                    End If
                                End If
                            End If
                        Else
                            ' -- Inserta en tabla intermedia los conceptos de aseguradora
                            pInsertaConceptos
                        End If
                    End If
                      
53                    If FraTiposHospitalFarmacia.Visible = True And FraTipos.Visible = False Then
''54                        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
''                                          & "|" & IIf(optTipoPaciente(0).Value, "I", "E") _
''                                          & "|" & IIf(optOrden(0).Value, 1, IIf(optOrden(1).Value, 2, IIf(optOrden(2).Value, 3, IIf(optOrden(3).Value, 4, 5)))) _
''                                          & "|" & 1 _
''                                          & "|" & IIf(optCargo(0).Value, 0, IIf(optCargo(2).Value, 1, 0)) _
''                                          & "|" & CStr(chkCosto.Value) _
''                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
''                                          & "|" & 0 _
''                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) _
''                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")) _
''                                          & "|" & vgintClaveEmpresaContable _
''                                          & "|" & str(chkDesglosarPaquete.Value) _
''                                          & "|" & IIf(chkHora.Value, 1, 0) _
''                                          & "|" & IIf(optEstadoCuentaHospitalFarm(0).Value, 1, IIf(optEstadoCuentaHospitalFarm(1).Value, 2, 3)) & "|" & -1
                                          
                          vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                          & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                          & "|" & IIf(optOrden(0).Value, 1, IIf(optOrden(1).Value, 2, IIf(optOrden(2).Value, 3, IIf(optOrden(3).Value, 4, 5)))) _
                                          & "|" & 1 _
                                          & "|" & IIf(optCargo(0).Value, 0, IIf(optCargo(2).Value, 1, 0)) _
                                          & "|" & CStr(chkCosto.Value) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & 0 _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")) _
                                          & "|" & vgintClaveEmpresaContable _
                                          & "|" & str(chkDesglosarPaquete.Value) _
                                          & "|" & IIf(chkHora.Value, 1, 0) _
                                          & "|" & IIf(optEstadoCuentaHospitalFarm(0).Value, 1, IIf(optEstadoCuentaHospitalFarm(1).Value, 2, 3)) & "|" & -1 _
                                          & "|" & IIf(blnEstadoAgrupado, 1, 0)

55                    Else
                            
                            indiceCbo = -1
                            'LMM
                            If cboCartas.ListCount > 0 Then
                                If (cboCartas.ListIndex = -1 Or cboCartas.ListIndex = 0) And optEstadoCuenta(0).Value Then
                                    indiceCbo = -1
                                Else
                                    indiceCbo = cboCartas.ItemData(cboCartas.ListIndex)
                                End If
                                
                            End If
                            
                            
'56
                                          
                          vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                          & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                          & "|" & IIf(optOrden(0).Value, 1, IIf(optOrden(1).Value, 2, IIf(optOrden(2).Value, 3, IIf(optOrden(3).Value, 4, 5)))) _
                                          & "|" & IIf(optEstadoCuenta(2).Value Or Not optEstadoCuenta(2).Enabled, 1, 0) _
                                          & "|" & IIf(optCargo(0).Value, 0, IIf(optCargo(2).Value, 1, 0)) _
                                          & "|" & CStr(chkCosto.Value) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & IIf(Me.optEstadoCuenta(0).Value, 0, IIf(optEstadoCuenta(1).Value, 1, 2)) _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")) _
                                          & "|" & vgintClaveEmpresaContable _
                                          & "|" & str(chkDesglosarPaquete.Value) _
                                          & "|" & IIf(chkHora.Value, 1, 0) _
                                          & "|" & 0 & "|" & indiceCbo _
                                          & "|" & IIf(blnEstadoAgrupado, 1, 0)
                                          
57                    End If
                      
58                    Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelEstadoCuenta")
59                ElseIf optOrden(5).Value Then
60                    If FraTiposHospitalFarmacia.Visible = True And FraTipos.Visible = False Then
61                        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                          & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                          & "|" & IIf(lntCveEmpresaPaciente = 0, 2, IIf(optEstadoCuenta(0).Value, 1, IIf(optEstadoCuenta(1).Value, 0, 2))) _
                                          & "|" & IIf(optEstadoCuentaHospitalFarm(0).Value, 1, IIf(optEstadoCuentaHospitalFarm(1).Value, 2, 3))
                                          
62                        Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELESTADOCTACARGOSHOSFARM")
63                    Else
                            
                            indiceCbo = -1
                            If cboCartas.ListCount > 0 Then
                                If (cboCartas.ListIndex = -1 Or cboCartas.ListIndex = 0) And optEstadoCuenta(0).Value Then
                                    indiceCbo = -1
                                Else
                                    indiceCbo = cboCartas.ItemData(cboCartas.ListIndex)
                                End If
                                
                            End If
64                        vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                          & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                          & "|" & IIf(lintCveEmpresaPaciente = 0, 2, IIf(optEstadoCuenta(0).Value, 1, IIf(optEstadoCuenta(1).Value, 0, 2))) & "|" & indiceCbo
                                          
65                        Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelEstadoCuentaCargos")
66                    End If
                    Else
                        If FraTiposHospitalFarmacia.Visible = True And FraTipos.Visible = False Then
                            vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                              & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                              & "|" & IIf(lntCveEmpresaPaciente = 0, 2, IIf(optEstadoCuenta(0).Value, 1, IIf(optEstadoCuenta(1).Value, 0, 2))) _
                                              & "|" & IIf(optEstadoCuentaHospitalFarm(0).Value, 1, IIf(optEstadoCuentaHospitalFarm(1).Value, 2, 3))
                                              
                            Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELESTADOCTACARGOSHOSFARM")
                        Else
                             
                             '(cboFactura.ItemData(cboFactura.ListIndex)
                            indiceCbo = -1
                            If cboCartas.ListCount > 0 Then
                                If (cboCartas.ListIndex = -1 Or cboCartas.ListIndex = 0) And optEstadoCuenta(0).Value Then
                                    indiceCbo = -1
                                Else
                                    indiceCbo = cboCartas.ItemData(cboCartas.ListIndex)
                                End If
                                
                            End If
                            
                            vgstrParametrosSP = Val(txtMovimientoPaciente.Text) _
                                          & "|" & IIf(OptTipoPaciente(0).Value, "I", "E") _
                                          & "|" & IIf(optOrden(0).Value, 1, IIf(optOrden(1).Value, 2, IIf(optOrden(2).Value, 3, IIf(optOrden(3).Value, 4, 5)))) _
                                          & "|" & IIf(optEstadoCuenta(2).Value Or Not optEstadoCuenta(2).Enabled, 1, 0) _
                                          & "|" & IIf(optCargo(0).Value, 0, IIf(optCargo(2).Value, 1, 0)) _
                                          & "|" & CStr(chkCosto.Value) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & IIf(Me.optEstadoCuenta(0).Value, 0, IIf(optEstadoCuenta(1).Value, 1, 2)) _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("01/01/1900 00:00:00"), fstrFechaSQL(mskFechaInicio.Text, "00:00:00")) _
                                          & "|" & IIf(chkRangoFechas.Value <> vbChecked, fstrFechaSQL("31/12/3999 23:59:59"), fstrFechaSQL(mskFechaFin.Text, "23:59:59")) _
                                          & "|" & vgintClaveEmpresaContable _
                                          & "|" & Trim(str(chkDesglosarPaquete.Value)) _
                                          & "|" & IIf(chkHora.Value, 1, 0) _
                                          & "|" & 0 & "|" & indiceCbo _
                                          & "|" & IIf(blnEstadoAgrupado, 1, 0)
'

                            Set rsEstadoCuenta = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelEstadoCuenta")
                        End If
67                End If
68                If rsEstadoCuenta.EOF Then
                      'No existe información con esos parámetros.
69                    MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
70                Else
                        
71                    alstrParametros(0) = "Fecha;" & fdtmServerFecha
72                    alstrParametros(1) = "Hora;" & fdtmServerHora
73                    If chkCosto.Value = 0 Then
74                        alstrParametros(2) = "Costo;"
75                    Else
76                        alstrParametros(2) = "Costo;" & "Costo"
77                    End If
78                    alstrParametros(3) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
79                    alstrParametros(4) = "DireccionHospital;" & Trim(vgstrDireccionCH) & " " & Trim(vgstrColoniaCH) & " " & Trim(vgstrCiudadCH)
80                    alstrParametros(5) = "RFC;" & "RFC " & vgstrRfCCH

                      'Se agrego para ocultar expediente si es externo
                      If OptTipoPaciente(1).Value Then
                        Set rs = frsRegresaRs("SELECT * FROM SIPARAMETRO WHERE VCHNOMBRE = 'BITOCULTAEXPEDIENTE' AND INTCVEEMPRESACONTABLE = " & vgintClaveEmpresaContable, adLockOptimistic)
                        If rs.RecordCount > 0 Then
                            If rs!VCHVALOR = 0 Then
                                alstrParametros(6) = "NúmeroPaciente;" & rsDatosPaciente!NumPaciente
                            Else
                                alstrParametros(6) = "NúmeroPaciente;" & "0"
                            End If
                        Else
81                          alstrParametros(6) = "NúmeroPaciente;" & rsDatosPaciente!NumPaciente
                        End If
                      Else
                        alstrParametros(6) = "NúmeroPaciente;" & rsDatosPaciente!NumPaciente
                      End If
                        alstrParametros(45) = "NumControl;" & IIf(IsNull(rsDatosPaciente!NumControl), "", rsDatosPaciente!NumControl)
82                    alstrParametros(7) = "Cuenta;" & txtMovimientoPaciente.Text
83                    alstrParametros(8) = "Tipo;" & IIf(OptTipoPaciente(0).Value, "INTERNO", "EXTERNO")
84                    alstrParametros(9) = "Nombre;" & IIf(IsNull(rsDatosPaciente!Nombre), "", rsDatosPaciente!Nombre)
85                    alstrParametros(10) = "Domicilio;" & IIf(IsNull(rsDatosPaciente!Domicilio), "", rsDatosPaciente!Domicilio)
86                    alstrParametros(11) = "Ciudad;" & IIf(IsNull(rsDatosPaciente!Ciudad), "", rsDatosPaciente!Ciudad)
87                    alstrParametros(12) = "Estado;" & IIf(IsNull(rsDatosPaciente!Estado), "", rsDatosPaciente!Estado)
88                    alstrParametros(13) = "FechaNacimiento;"
89                    alstrParametros(43) = "MostrarCirugias;" & IIf(chkMostrarCirugias.Value, 1, 0)
                      
90                    If Not IsNull(rsDatosPaciente!FechaNacimiento) Then
91                        alstrParametros(13) = "FechaNacimiento;" & Format(rsDatosPaciente!FechaNacimiento, "dd/mmm/yyyy")
92                    End If
                      
93                    alstrParametros(14) = "Edad;"
                      
94                    If Not IsNull(rsDatosPaciente!FechaNacimiento) Then
                          'Se modifico para que no muestre los meses en la edad - caso 20431
                          vlstrEdadConMeses = fstrObtieneEdad(rsDatosPaciente!FechaNacimiento, fdtmServerFecha)
                          
                          For vlIntCont = 1 To Len(vlstrEdadConMeses)
                            If Mid(vlstrEdadConMeses, vlIntCont, 1) <> "/" Then
                              vlstrEdadSinMeses = vlstrEdadSinMeses & Mid(vlstrEdadConMeses, vlIntCont, 1)
                            Else
                              vlstrEdadSinMeses = vlstrEdadSinMeses & " AÑO(S)"
                              Exit For
                            End If
                          Next vlIntCont
                          
95                        alstrParametros(14) = "Edad;" & vlstrEdadSinMeses
                          'Hasta aquí
                          'alstrParametros(14) = "Edad;" & fstrObtieneEdad(rsDatosPaciente!FechaNacimiento, fdtmServerFecha)
96                    End If
                      
97                    alstrParametros(15) = "FechaIngreso;"
                      
98                    If Not IsNull(rsDatosPaciente!Ingreso) Then
99                        alstrParametros(15) = "FechaIngreso;" & Format(rsDatosPaciente!Ingreso, "dd/mmm/yyyy hh:mm")
100                   End If
                      
101                   alstrParametros(16) = "FechaEgreso;"
                      
102                   If Not IsNull(rsDatosPaciente!Egreso) Then
103                      alstrParametros(16) = "FechaEgreso;" & Format(rsDatosPaciente!Egreso, "dd/mmm/yyyy hh:mm")
104                   End If
                      
105                   alstrParametros(17) = "UltimoCuarto;" & IIf(IsNull(rsDatosPaciente!Cuarto), "", rsDatosPaciente!Cuarto)

                     '*****************caso 20015
                    If lblbActivaResponsableAD And IsNull(rsDatosPaciente!Responsable) = False Then
                          
                          alstrParametros(18) = "Responsable;" & IIf(IsNull(rsDatosPaciente!Responsable), "", rsDatosPaciente!Responsable)
                          
                     Else
                      '*****************
                          'condicionamos este parametro para que funcione segun las carts de control de aseguradora
                          
                        If blnCartaEncontrada Then
                              alstrParametros(18) = "Responsable;" & IIf(IsNull(txtEmpresaPaciente.Text), "", txtEmpresaPaciente.Text)
                        Else
106                           alstrParametros(18) = "Responsable;" & IIf(IsNull(rsDatosPaciente!Responsable), "", rsDatosPaciente!Responsable)
                        End If
                      End If
107                   alstrParametros(19) = "MedicoTratante;" & IIf(IsNull(rsDatosPaciente!Medico), "", rsDatosPaciente!Medico)
108                   alstrParametros(20) = "TipoPaciente;" & IIf(IsNull(rsDatosPaciente!tipo), "", rsDatosPaciente!tipo)
                      If blnCartaEncontrada Then
109                     alstrParametros(21) = "Empresa;" & IIf(IsNull(txtEmpresaPaciente.Text), "", txtEmpresaPaciente.Text)
                      Else
                        alstrParametros(21) = "Empresa;" & IIf(IsNull(rsDatosPaciente!empresa), "", rsDatosPaciente!empresa)
                      End If
110                   If optEstadoCuenta(0).Value Then
111                       If FraTiposHospitalFarmacia.Visible = True And FraTipos.Visible = False Then
112                           If optEstadoCuentaHospitalFarm(0).Value = True Then
113                               alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA DE FARMACIA"
114                           Else
115                               If optEstadoCuentaHospitalFarm(1).Value = True Then
116                                   alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA DE HOSPITAL"
117                               Else
118                                   alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA CONSOLIDADO"
119                               End If
120                           End If
121                       Else
122                           alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA PACIENTE"
123                       End If
124                   Else
125                       If optEstadoCuenta(1).Value Then
126                           alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA EMPRESA"
127                       Else
128                           alstrParametros(22) = "Leyenda;" & "ESTADO DE CUENTA PACIENTE / EMPRESA"
129                       End If
130                   End If
                      '----------------------------------------------
                      'cambio para que siempre que existan grupos, estos se muestren sin importar si estan facturados o no (CGR)
                      
131                   vlstrGrupoCuentas = ""
132                   If cboFactura.ListIndex > -1 Then
133                       If Not optEstadoCuenta(0).Value Then ' solamente cuando NO es el estado de cuenta del paciente
134                           vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & _
                                               "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & _
                                               "|" & lintCveEmpresaPaciente & _
                                               "|" & vgintClaveEmpresaContable & _
                                               "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", cboFactura.List(cboFactura.ListIndex))
135                           Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelGruposEdoCuenta")
                              
136                           If rs.RecordCount > 0 Then
137                               Do While Not rs.EOF
138                                   If Not IsNull(rs!NumGrupo) Then
139                                       If vlstrGrupoCuentas = "" Then
140                                           vlstrGrupoCuentas = "GRUPOS DE CUENTAS " & rs!NumGrupo
141                                       Else
142                                           vlstrGrupoCuentas = vlstrGrupoCuentas & ", " & rs!NumGrupo
143                                       End If
144                                   End If
145                                   rs.MoveNext
146                               Loop
147                           End If
148                       End If
149                   End If
                      
150                   alstrParametros(34) = "GrupoFacturas;" & vlstrGrupoCuentas
                      
                      '----------------------------------------------
                      
                    Dim seleccion As Integer
                    seleccion = -1
                    If optEstadoCuenta(0) Then
                        seleccion = 0
                    ElseIf optEstadoCuenta(1) Then
                        seleccion = 1
                    ElseIf optEstadoCuenta(2) Then
                        seleccion = 2
                    End If
                        
                        
                      
                      Select Case seleccion
                        Case 0
                            If cboFactura.ListIndex > -1 Then
                                If cboFactura.ItemData(cboFactura.ListIndex) = -1 Then
                                    vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & _
                                                              "|" & IIf(optEstadoCuenta(0).Value, "P", IIf(optEstadoCuenta(1).Value, "E", "A")) & _
                                                              "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
                                      Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFolioFacturasCuenta")
                                          
                                      vlstrFacturasPaciente = ""
                                     Do While Not rs.EOF
                                         vlstrFacturasPaciente = vlstrFacturasPaciente & ", " & Trim(rs!chrfoliofactura)
                                           rs.MoveNext
                                      Loop
                                       vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
                                       rs.Close
                                Else
                                   vlstrFacturasPaciente = cboFactura.List(cboFactura.ListIndex)
                                End If

                            End If
                            alstrParametros(23) = "Facturas;" & vlstrFacturasPaciente
                        Case 1
                                If blnCartaEncontrada = False Then
                                      
                                      If cboFactura.ItemData(cboFactura.ListIndex) = -1 Then
                                             vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & _
                                                            "|" & "E" & _
                                                            "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
                                          Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFolioFacturasCuenta")
                                        
                                          vlstrFacturasPaciente = ""
                                          Do While Not rs.EOF
                                              vlstrFacturasPaciente = vlstrFacturasPaciente & ", " & Trim(rs!chrfoliofactura)
                                              rs.MoveNext
                                          Loop
                                          vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
                                          rs.Close
                                      Else
                                            vlstrFacturasPaciente = cboFactura.List(cboFactura.ListIndex)
                                      End If
                                      alstrParametros(23) = "Facturas;" & vlstrFacturasPaciente
                                Else
                                
                                    vlstrSentencia = "SELECT * FROM PVFACTURA " & _
                                    "INNER JOIN PVCARTACONTROLSEGURO ON PVFACTURA.INTCVECARTA = PVCARTACONTROLSEGURO.INTCVECARTA AND PVCARTACONTROLSEGURO.INTCVEEMPRESA =  PVFACTURA.INTCVEEMPRESA " & _
                                    "WHERE INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " AND PVFACTURA.INTCVECARTA = " & cboCartas.ItemData(cboCartas.ListIndex) & "and not pvfactura.chrestatus = 'C'"
                                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                    
                                    If cboFactura.ListIndex > 0 Then rs.Filter = "CHRFOLIOFACTURA = " & "'" & Trim(cboFactura.List(cboFactura.ListIndex) & "'")
                                    
                                    Do While Not rs.EOF
                                        vlstrFacturasPaciente = vlstrFacturasPaciente & ", " & Trim(rs!chrfoliofactura)
                                        rs.MoveNext
                                    Loop
                                    vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
                                    
                                    rs.Close
                                    alstrParametros(23) = "Facturas;" & vlstrFacturasPaciente
                                End If
                        Case 2
                            If blnCartaEncontrada = True Then
                             vlstrFacturasPaciente = ""
                             vlstrSentencia = "SELECT * FROM PVFACTURA " & _
                                    "WHERE INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " AND PVFACTURA.INTCVECARTA = " & cboCartas.ItemData(cboCartas.ListIndex) & "and not pvfactura.chrestatus = 'C'"
                                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                                    
                                    Do While Not rs.EOF
                                        vlstrFacturasPaciente = vlstrFacturasPaciente & ", " & Trim(rs!chrfoliofactura)
                                        rs.MoveNext
                                    Loop
                                    vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
                                    
                                    rs.Close
                                    alstrParametros(23) = "Facturas;" & vlstrFacturasPaciente
                            Else
                                If cboFactura.ListIndex > -1 Then
                                    If cboFactura.ItemData(cboFactura.ListIndex) = -1 Then
                                        vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & _
                                                                  "|" & "A" & _
                                                                  "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
                                          Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFolioFacturasCuenta")
                                              
                                          vlstrFacturasPaciente = ""
                                         Do While Not rs.EOF
                                             vlstrFacturasPaciente = vlstrFacturasPaciente & ", " & Trim(rs!chrfoliofactura)
                                               rs.MoveNext
                                          Loop
                                           vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
                                           rs.Close
                                    Else
                                       vlstrFacturasPaciente = cboFactura.List(cboFactura.ListIndex)
                                    End If
                                End If
                                alstrParametros(23) = "Facturas;" & vlstrFacturasPaciente
                            End If
                        End Select
                        
166
                                     
167                   vlstrFacturasPaciente = ""

'---------------------------------------------------------------------------
'                               PAQUETES
'---------------------------------------------------------------------------
        
        Dim vlstrSentenciaPaq As String
        If cboCartas.ListIndex = -1 Or cboCartas.ListIndex = 0 Then
            vlstrSentenciaPaq = ""
        Else
            vlstrSentenciaPaq = "and pvcargo.intcvecarta = " & cboCartas.ItemData(cboCartas.ListIndex)
        End If
        
                    vlstrSentencia = "SELECT DISTINCT PvPaquete.CHRDESCRIPCION, PvPaquetePaciente.MNYPRECIOPAQUETE, PvPaquetePaciente.MNYPRECIOPAQUETE*(PvConceptoFacturacion.SMYIVA/100) IVA" & _
                                        " FROM PvCargo " & _
                                        " INNER JOIN PvPaquete ON PvCargo.INTNUMPAQUETE = PvPaquete.INTNUMPAQUETE" & _
                                        " INNER JOIN PvPaquetePaciente ON PvPaquete.INTNUMPAQUETE = PvPaquetePaciente.INTNUMPAQUETE AND PvPaquetePaciente.INTMOVPACIENTE = PvCargo.INTMOVPACIENTE AND PvPaquetePaciente.CHRTIPOPACIENTE = PvCargo.CHRTIPOPACIENTE" & _
                                        " INNER JOIN PvConceptoFacturacion ON PvPaquete.SMICONCEPTOFACTURA = PvConceptoFacturacion.SMICVECONCEPTO" & _
                                        " WHERE PvCargo.intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & _
                                        " AND PvCargo.chrTipoPaciente = '" & IIf(OptTipoPaciente(0).Value, "I", "E") & "'" & _
                                        " AND (" & IIf(optCargo(0).Value, 1, 0) & " = 1 OR PvCargo.bitExcluido = " & IIf(optCargo(0).Value, 0, IIf(optCargo(2).Value, 1, 0)) & ")" & _
                                        " AND (" & IIf(Me.cboFactura.ListIndex > 0, "'" & Me.cboFactura.Text & "'", "'*'") & " = '*' OR PvCargo.chrFolioFactura = '" & Me.cboFactura.Text & "')" & _
                                        " AND (PVCARGO.INTCANTIDADPAQUETE <> 0 OR PVCARGO.INTCANTIDADEXTRAPAQUETE <> 0)" & _
                                        " AND (PVCARGO.INTCANTIDADPAQUETE IS NOT NULL OR PVCARGO.INTCANTIDADEXTRAPAQUETE IS NOT NULL) " & vlstrSentenciaPaq
                    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                    vldblTotalPaquetes = 0
                    Do While Not rs.EOF
                       vldblIVA = vldblIVA + rs!IVA
                        vldblTotalPaquetes = vldblTotalPaquetes + rs!MNYPRECIOPAQUETE
                        vlstrFacturasPaciente = vlstrFacturasPaciente & "," & Trim(rs!chrDescripcion)
                        rs.MoveNext
                    Loop
        
        




177                   vlstrFacturasPaciente = Mid(vlstrFacturasPaciente, 2, Len(vlstrFacturasPaciente))
178                   rs.Close
                      
179                   alstrParametros(24) = "Paquetes;" & vlstrFacturasPaciente
                                      
180                   If optOrden(0).Value Then 'Concepto de facturación
181                       alstrParametros(25) = "Campo1;" & "Fecha"
182                       alstrParametros(26) = "Campo2;" & "Número"
183                       alstrParametros(27) = "Campo3;" & "Clave"
184                       alstrParametros(28) = "Campo4;" & "Descripción"
185                   Else
186                       If optOrden(1).Value Then 'Fecha
187                           alstrParametros(25) = "Campo1;" & "Número"
188                           alstrParametros(26) = "Campo2;" & "Clave"
189                           alstrParametros(27) = "Campo3;" & "Concepto"
190                           alstrParametros(28) = "Campo4;" & "Descripción"
191                       ElseIf optOrden(2).Value Then  'Departamento
192                           alstrParametros(25) = "Campo1;" & "Fecha"
193                           alstrParametros(26) = "Campo2;" & "Número"
194                           alstrParametros(27) = "Campo3;" & "Clave"
195                           alstrParametros(28) = "Campo4;" & "Descripción"
196                       Else ' cargo
197                           alstrParametros(25) = "Campo1;" & "Fecha"
198                           alstrParametros(26) = "Campo2;" & "Clave"
199                           alstrParametros(27) = "Campo3;" & "Número"
                              alstrParametros(28) = "Campo4;" & "Descripción"
201                       End If
202                   End If
                      
203                   alstrParametros(29) = "telefono;" & rsDatosPaciente!TelefonoPaciente
204                   alstrParametros(30) = "Comentario;" & Trim(txtMensaje.Text)
205                   alstrParametros(31) = "diagnóstico;" & IIf(IsNull(rsDatosPaciente!Diagnostico), "", rsDatosPaciente!Diagnostico)
206                   alstrParametros(32) = "TasaIVA;" & vgdblCantidadIvaGeneral
                      
                      Dim vgStrConsultaImporteFactura As String
                      Dim vintSumatoria As Double
                      
                      Select Case seleccion
                        Case 0
                        
                                vgstrParametrosSP = str(chkRangoFechas.Value) & "|" & fstrFechaSQL(mskFechaInicio.Text) & "|" & fstrFechaSQL(mskFechaFin.Text) _
                                                    & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                                    & "|" & IIf(optEstadoCuenta(0).Value, "P", IIf(optEstadoCuenta(1).Value, "E", "*")) _
                                                    & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E")
                                  Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELIMPORTEFACTURADO")
                                
                                  alstrParametros(33) = "ImporteFacturado;" & rs!ImporteFacturado
                                
                                  rs.Close
                        Case 1
                            If blnCartaEncontrada Then
                                
                                vgStrConsultaImporteFactura = ""
                                vintSumatoria = 0
                                vgStrConsultaImporteFactura = "SELECT * FROM PVFACTURA WHERE INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " AND (PVFACTURA.INTCVECARTA = " & cboCartas.ItemData(cboCartas.ListIndex) & " AND TRIM(PVFACTURA.CHRTIPOFACTURA) = 'E' AND PVFACTURA.CHRESTATUS <> 'C')"
                                Set rs = frsRegresaRs(vgStrConsultaImporteFactura, adLockOptimistic, adOpenDynamic)
                                
                                If cboFactura.ListIndex > 0 Then rs.Filter = "CHRFOLIOFACTURA = " & "'" & Trim(cboFactura.List(cboFactura.ListIndex) & "'")

                                Do While Not rs.EOF
                                    vintSumatoria = vintSumatoria + rs!mnyTotalFactura
                                    rs.MoveNext
                                Loop
                                alstrParametros(33) = "ImporteFacturado;" & vintSumatoria
                            Else
                                 vgstrParametrosSP = str(chkRangoFechas.Value) & "|" & fstrFechaSQL(mskFechaInicio.Text) & "|" & fstrFechaSQL(mskFechaFin.Text) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & IIf(optEstadoCuenta(0).Value, "P", IIf(optEstadoCuenta(1).Value, "E", "*")) _
                                          & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E")
                                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELIMPORTEFACTURADO")
                      
                                alstrParametros(33) = "ImporteFacturado;" & rs!ImporteFacturado
                      
                                rs.Close
                            
                            End If
                            
                        Case 2
                             If blnCartaEncontrada Then
                                vgStrConsultaImporteFactura = ""
                                vintSumatoria = 0
                                vgStrConsultaImporteFactura = "SELECT * FROM PVFACTURA WHERE INTMOVPACIENTE = " & txtMovimientoPaciente.Text & " AND (PVFACTURA.INTCVECARTA = " & cboCartas.ItemData(cboCartas.ListIndex) & ") AND PVFACTURA.CHRESTATUS <> 'C'"
                                Set rs = frsRegresaRs(vgStrConsultaImporteFactura, adLockOptimistic, adOpenDynamic)
                                Do While Not rs.EOF
                                    vintSumatoria = vintSumatoria + rs!mnyTotalFactura
                                    rs.MoveNext
                                Loop
                                alstrParametros(33) = "ImporteFacturado;" & vintSumatoria
                            Else
                                 vgstrParametrosSP = str(chkRangoFechas.Value) & "|" & fstrFechaSQL(mskFechaInicio.Text) & "|" & fstrFechaSQL(mskFechaFin.Text) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & IIf(optEstadoCuenta(0).Value, "P", IIf(optEstadoCuenta(1).Value, "E", "*")) _
                                          & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E")
                                Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELIMPORTEFACTURADO")
                      
                                alstrParametros(33) = "ImporteFacturado;" & rs!ImporteFacturado
                      
                                rs.Close
                            
                            End If
                      End Select
                      
                      'Cálculo del importe facturado según el estado de cuenta que se pidió (paciente, empresa, paciente - empresa)

                      
                      '----------------------------------------------------------------------------------------------------
211                   lngConceptosAseguradora = 1
212                   If optEstadoCuenta(1).Value Or optEstadoCuenta(2).Value Then
213                       frsEjecuta_SP txtMovimientoPaciente.Text & "|1|" & IIf(OptTipoPaciente(0).Value, "I", "E"), "Fn_PvSelBitDesglose", True, lngConceptosAseguradora
214                   End If
                      
215                   alstrParametros(35) = "DesgloseConceptosSeguros;" & IIf(lngConceptosAseguradora = 3, 1, 0)
                                      
                      '-------------------------------
                      'Descuento especial ya facturado
                      '-------------------------------
216                   dblDescuentoEspecial = 0
217                   If optEstadoCuenta(1).Value Or optEstadoCuenta(2).Value Then
218                      vgstrParametrosSP = str(chkRangoFechas.Value) & "|" & fstrFechaSQL(mskFechaInicio.Text) & "|" & fstrFechaSQL(mskFechaFin.Text) _
                                          & "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", Trim(cboFactura.List(cboFactura.ListIndex))) _
                                          & "|" & txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E")
219                      Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDESCUENTOESPECIAL")
                         
220                      dblDescuentoEspecial = IIf(IsNull(rs!DescuentoEspecial), 0, rs!DescuentoEspecial)
221                   End If
222                   alstrParametros(36) = "DescuentoEspecialFacturado;" & dblDescuentoEspecial
                      
                      '---------------------------------------------------
                      'Datos para calcular descuento especial sin facturar
                      '---------------------------------------------------
223                   Set rs = frsRegresaRs("Select * from PvDescuentoEspecial where dtmfechainicial <= to_date(" & _
                                              fstrFechaSQL(fdtmServerFecha) & ",'yyyy-mm-dd') and dtmfechafinal >= to_date(" & _
                                              fstrFechaSQL(fdtmServerFecha) & ",'yyyy-mm-dd') and intcveempresa = " & lintCveEmpresaPaciente, adLockOptimistic)
                         
224                   If rs.RecordCount > 0 And (optEstadoCuenta(1).Value Or optEstadoCuenta(2).Value) Then
225                      alstrParametros(37) = "DescuentoEspecialPorcentaje;" & IIf(IsNull(rs!NUMPORCENTAJE), 0, rs!NUMPORCENTAJE)
226                      alstrParametros(38) = "DescuentoEspecialLimite;" & IIf(IsNull(rs!MNYMONTOAPLICARLIMITE), 0, rs!MNYMONTOAPLICARLIMITE)
227                      alstrParametros(44) = "DescuentoEspecialExclusion;" & IIf(IsNull(rs!bitConsideraExcluidos), 0, rs!bitConsideraExcluidos)
228                   Else
229                      alstrParametros(37) = "DescuentoEspecialPorcentaje;0"
230                      alstrParametros(38) = "DescuentoEspecialLimite;0"
231                      alstrParametros(44) = "DescuentoEspecialExclusion;0"
232                   End If
233                   alstrParametros(39) = "Treporte;" & IIf(Me.optEstadoCuenta(0).Value, 0, IIf(optEstadoCuenta(1).Value, 1, 2))
234                   If Not (optOrden(5).Value Or optOrden(6).Value) Then
235                       alstrParametros(40) = "tipoAgrupacion;" & IIf(optOrden(0).Value, 1, IIf(optOrden(1).Value, 2, IIf(optOrden(2).Value, 3, 4)))
236                   Else
237                       alstrParametros(40) = "tipoAgrupacion;" & 5
238                   End If
                      
239                   alstrParametros(41) = "CuatroDecimales;" & chkMostrarCuatroDecimales.Value
240                   alstrParametros(42) = "PorcentajeRetencion;" & vldblRetencion
                      '------------------------------------------------------------------------------
                      
'241                   If Not optOrden(5).Value Then
'242                       pCargaParameterFields alstrParametros, vgrptReporte1
'243                       pImprimeReporte vgrptReporte1, rsEstadoCuenta, vlstrDestino, "Estado de cuenta"
'244                   Else
'245                       pCargaParameterFields alstrParametros, vgrptReporte5
'246                       pImprimeReporte vgrptReporte5, rsEstadoCuenta, vlstrDestino, "Estado de cuenta cargos"
'247                   End If

241                   If optOrden(5).Value Then
242                       pCargaParameterFields alstrParametros, vgrptReporte4
243                       pImprimeReporte vgrptReporte4, rsEstadoCuenta, vlstrDestino, "Estado de cuenta cargos"
244                   ElseIf optOrden(6).Value Then
245                       pCargaParameterFields alstrParametros, vgrptReporte5
246                       pImprimeReporte vgrptReporte5, rsEstadoCuenta, vlstrDestino, "Estado de cuenta cargos agrupados"
                         Else
                               pCargaParameterFields alstrParametros, vgrptReporte1
                               pImprimeReporte vgrptReporte1, rsEstadoCuenta, vlstrDestino, "Estado de cuenta"
247                   End If
                      
248                   If chkPagos.Value = 1 Then
249                       pReportePagos vlstrDestino, IIf(IsNull(rsDatosPaciente!Nombre), "", rsDatosPaciente!Nombre) 'El destino sería "P" o "I"
250                   End If
251               End If
252               rsEstadoCuenta.Close
253           Else
                  'No se encontró la información del paciente.
254               MsgBox SIHOMsg(355), vbOKOnly + vbExclamation, "Mensaje"
255           End If
       
256           rsDatosPaciente.Close
257           Me.MousePointer = 0
258           If optEstadoCuenta(1).Value Or optEstadoCuenta(2).Value Then
259                If lintCveEmpresaPaciente = llngCveEmpresaPCE Then
260                   vgstrParametrosSP = IIf(OptTipoPaciente(0).Value, "I", "E") & txtMovimientoPaciente.Text & "," & "|" & llngCveEmpresaPCE
261                   Set rsInformacionFaltantePCE = frsEjecuta_SP(vgstrParametrosSP, "sp_CCVerificarInfoFaltantePCE")
262                   If Not rsInformacionFaltantePCE.EOF Then
263                       If MsgBox("Existen pacientes que tiene cargos sin código de PCE. ¿Desea obtener un listado de los cargos?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
264                           vgrptReporte3.DiscardSavedData
265                           alstrParametros2(0) = "NombreHospital;" & Trim(vgstrNombreHospitalCH)
266                           pCargaParameterFields alstrParametros2, vgrptReporte3
267                           pImprimeReporte vgrptReporte3, rsInformacionFaltantePCE, "P", "Pacientes sin código de PCE", False
268                       End If
269                   End If
270                   rsInformacionFaltantePCE.Close
271               ElseIf fblnManejaCatalogoCargos(lintCveEmpresaPaciente) Then
272                   lblnFueraCatalogo = fblnCargosFueraCatalogo(txtMovimientoPaciente.Text, IIf(OptTipoPaciente(0).Value, "I", "E"), lintCveEmpresaPaciente)
273               End If
274           End If
275       Else
              'Seleccione el paciente.
276           MsgBox SIHOMsg(353), vbOKOnly + vbInformation, "Mensaje"
277       End If

278   Exit Sub
NotificaError:
       Me.MousePointer = 0
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pReporte" & " Linea:" & Erl()))
End Sub

Private Sub pReportePagos(vlstrDestino As String, vlstrNombrePaciente As String)
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlstrNombreHospital As String
    Dim vlstrRegistro As String
    Dim vlstrDireccionHospital As String
    Dim vlstrTelefonoHospital As String
    Dim vlstrDepartamento As String
    Dim vlstrx As String
    Dim alstrParametros(4) As String
    
    '-------Traer datos generales del Hospital-----------
    vlstrNombreHospital = Trim(vgstrNombreHospitalCH)
    vlstrRegistro = "R.SSA " & Trim(vgstrSSACH) & " RFC " & Trim(vgstrRfCCH)
    vlstrDireccionHospital = Trim(vgstrDireccionCH) & " CP. " & Trim(vgstrCodPostalCH)
    vlstrTelefonoHospital = Trim(vgstrTelefonoCH)
    
    vlstrSentencia = "SELECT vchDescripcion FROM Nodepartamento WHERE smiCveDepartamento = " & Trim(str(vgintNumeroDepartamento))
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rs.RecordCount > 0 Then vlstrDepartamento = rs!VCHDESCRIPCION
    rs.Close
    
    Set rs = frsEjecuta_SP(Val(txtMovimientoPaciente.Text) & "|" & IIf(OptTipoPaciente(0).Value, "I", "E"), "SP_PVSELREPORTEPAGOSPACIENTE")
    If rs.EOF Then
        MsgBox "El paciente seleccionado no tiene pagos registrados", vbOKOnly + vbInformation, "Mensaje"
    Else
        vgrptReporte2.DiscardSavedData
        alstrParametros(0) = "FechaActual;" & Format(fdtmServerFecha, "dd/mmm/yyyy")
        alstrParametros(1) = "MovPaciente;" & txtMovimientoPaciente
        alstrParametros(2) = "NombreEmpresa;" & Trim(vlstrNombreHospital)
        alstrParametros(3) = "NombrePaciente;" & vlstrNombrePaciente
        alstrParametros(4) = "NombreReporte;" & "PAGOS REALIZADOS"
        pCargaParameterFields alstrParametros, vgrptReporte2
        pImprimeReporte vgrptReporte2, rs, vlstrDestino, "Pagos"
    End If
    rs.Close
End Sub

Private Sub pCalculaTotales()
1     On Error GoTo NotificaError

          Dim rsPagos As New ADODB.Recordset
          Dim rsCargos As New ADODB.Recordset
          
2         vldblSubtotal = 0
3         vldblDescuento = 0
4         vldblIVA = 0
5         vldblPagos = 0
6         vldbltotal = 0
7         vldblTotalFacturado = 0
8         vldblTotalSinFacturar = 0
          
9         vlstrx = "SELECT * FROM PvEstadoCuenta WHERE intNumCuenta = " & txtMovimientoPaciente.Text & " AND chrTipoPaciente = " & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
10        Set rsCargos = frsRegresaRs(vlstrx)
11        If rsCargos.RecordCount <> 0 Then
12            rsCargos.MoveFirst
13            Do While Not rsCargos.EOF
14                If rsCargos!mnyCampo7 > 0 Then
15                    vldblSubtotal = vldblSubtotal + rsCargos!mnyCampo7
16                    vldblDescuento = vldblDescuento + rsCargos!mnyCampo8
17                    vldblIVA = vldblIVA + rsCargos!mnyCampo9
18                End If
19                If IsNull(rsCargos!chrfoliofactura) Then
20                    vldblTotalSinFacturar = vldblTotalSinFacturar + rsCargos!mnyCampo7 - rsCargos!mnyCampo8 + rsCargos!mnyCampo9
21                Else
22                    vldblTotalFacturado = vldblTotalFacturado + rsCargos!mnyCampo7 - rsCargos!mnyCampo8 + rsCargos!mnyCampo9
23                End If
24                rsCargos.MoveNext
25            Loop
              
26            vlstrx = "SELECT (SELECT sum(mnyCantidad) " & _
                       "FROM PvPago " & _
                       "WHERE chrTipoPago = 'NO' " & _
                       "AND bitCancelado = 0 " & _
                       "AND bitPesos = 1 " & _
                       "AND intMovPaciente = " & txtMovimientoPaciente.Text & _
                       "AND chrTipoPaciente = " & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ")" & _
                       "+" & _
                       "ISNULL((SELECT SUM(mnyCantidad*mnyTipoCambio) " & _
                              " FROM PvPago " & _
                              " WHERE chrTipoPago = 'NO' " & _
                              " AND bitCancelado = 0 " & _
                              " AND bitPesos = 0 " & _
                              " AND intMovPaciente = " & txtMovimientoPaciente.Text & _
                              " AND chrTipoPaciente = " & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & "),0)"
27            Set rsPagos = frsRegresaRs(vlstrx)
28            If rsPagos.RecordCount <> 0 Then
29                If Not IsNull(rsPagos.Fields(0)) Then
30                    vldblPagos = rsPagos.Fields(0)
31                End If
32            End If
              
33            vldbltotal = vldblSubtotal - vldblDescuento + vldblIVA
34        End If

35    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCalculaTotales" & " Linea:" & Erl()))
End Sub

Private Sub cmdPrint_Click()
1     On Error GoTo NotificaError
          
2         If FraTipos.Visible Then
3             If optEstadoCuenta(0).Enabled And optEstadoCuenta(0).Value = True And FraTipos.Enabled = True Then
4                 optEstadoCuenta(0).SetFocus
5             Else
6                 If optEstadoCuenta(1).Enabled And optEstadoCuenta(1).Value = True And FraTipos.Enabled = True Then
7                     optEstadoCuenta(1).SetFocus
8                 Else
9                     If optEstadoCuenta(2).Enabled And optEstadoCuenta(2).Value = True And FraTipos.Enabled = True Then
10                        optEstadoCuenta(2).SetFocus
11                    End If
12                End If
13            End If
14        End If
          
15        If FraTiposHospitalFarmacia.Visible Then
16            If optEstadoCuentaHospitalFarm(0).Enabled And optEstadoCuentaHospitalFarm(0).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
17                optEstadoCuentaHospitalFarm(0).SetFocus
18            Else
19                If optEstadoCuentaHospitalFarm(1).Enabled And optEstadoCuentaHospitalFarm(1).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
20                    optEstadoCuentaHospitalFarm(1).SetFocus
21                Else
22                    If optEstadoCuentaHospitalFarm(2).Enabled And optEstadoCuentaHospitalFarm(2).Value = True And FraTiposHospitalFarmacia.Enabled = True Then
23                        optEstadoCuentaHospitalFarm(2).SetFocus
24                    End If
25                End If
26            End If
27        End If
          
28        cmdPreview.Enabled = False
29        cmdPrint.Enabled = False
30        pReporte "I"
          
31        cmdPrint.Enabled = True
32        cmdPreview.Enabled = True
          
33        cmdPrint.SetFocus

34    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click" & " Linea:" & Erl()))
End Sub

Private Sub Form_Activate()
On Error GoTo NotificaError
    
    If vlblnEntrando Then
        optOrden(2).Value = True
        optCargo(0).Value = True
        OptTipoPaciente(0).Value = True
        vlblnEntrando = False
        
        If llngNumeroCuenta <> 0 Then
            txtMovimientoPaciente.Text = llngNumeroCuenta
            OptTipoPaciente(0).Value = lstrTipoPaciente = "I"
            OptTipoPaciente(1).Value = lstrTipoPaciente = "E"
            
            txtMovimientoPaciente_KeyDown vbKeyReturn, 1
            fraDatosPaciente.Enabled = False
        End If
    End If
    
    If Not (vgintNumeroModulo = 2 Or vgintNumeroModulo = 15) Then
        If chkCosto.Enabled Then
            chkCosto.Visible = False
        End If
        If fblnRevisaPermiso(vglngNumeroLogin, 305, "C") Then
            chkCosto.Visible = True
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
1     On Error GoTo NotificaError
          Dim vlstrsql As String
          Dim rsMostrarCirugias As New ADODB.Recordset
          
2         vgstrNombreForm = Me.Name
          
          Dim rs As New ADODB.Recordset
          
3         Me.Icon = frmMenuPrincipal.Icon
          
4         vlstrsql = "select vchvalor from siparametro where vchnombre = 'BITMOSTRARCIRUGIASESTADOCUENTA'"
5         Set rsMostrarCirugias = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
6         If rsMostrarCirugias.RecordCount > 0 Then
7             chkMostrarCirugias.Value = rsMostrarCirugias!VCHVALOR
8         Else
9             chkMostrarCirugias.Value = 1
10        End If
          
11        pCargaEmpresaPCE
              
12        pInstanciaReporte vgrptReporte1, "rptPVEstadoCuenta.rpt"
13        pInstanciaReporte vgrptReporte2, "rptPagos.rpt"
14        pInstanciaReporte vgrptReporte3, "rptPacientesInfoFaltantePCE.rpt"
15        pInstanciaReporte vgrptReporte4, "rptPVEstadoCuentaCargos.rpt"
            pInstanciaReporte vgrptReporte5, "rptPVEstadoCuentaCargosAgrupados.rpt"

              
16        vlblnLimpiar = True
17        vlblnEntrando = True
18        vlblnDesglosarPaquete = True
19        blnDatosCuenta = False
          
20        FraTipos.Enabled = True
21        FraTipos.Visible = True
          
22        FraTiposHospitalFarmacia.Enabled = False
23        FraTiposHospitalFarmacia.Visible = False
                             
24        If fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 4167, 4168), "C") Then
25             chkCosto.Enabled = True
               blnCostos = True
26           Else
27             chkCosto.Enabled = False
               blnCostos = False
28        End If
          'Se agregó para que solo permita agregar observaciones y notas internas con permiso C o E
          If fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 305, 1532), "C") Or fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 305, 1532), "E") Then
             txtMensaje.Locked = False
             TxtNotasInternas.Locked = False
          Else
             txtMensaje.Locked = True
             TxtNotasInternas.Locked = True
          End If
          
          chkCataCargosEmpresa.Enabled = fblnATCConCargosEmpresa
          vlintCveInterfazATC = 0
          
         
          
           chkRangoFechas.Enabled = False
          mskFechaInicio.Enabled = False
          mskFechaFin.Enabled = False
          blnEstadoAgrupado = False
29    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load" & " Linea:" & Erl()))
End Sub

Private Sub pCargaEmpresaPCE()
On Error GoTo NotificaError
    Dim rsEmpresaPCE As New ADODB.Recordset

   'Empresa PCE:
    llngCveEmpresaPCE = 0
      
    vgstrParametrosSP = vgintClaveEmpresaContable
    Set rsEmpresaPCE = frsEjecuta_SP(vgstrParametrosSP, "sp_GNSelParametroPCE")
    If rsEmpresaPCE.RecordCount <> 0 Then
        llngCveEmpresaPCE = IIf(IsNull(rsEmpresaPCE!intCveEmpresaPCE), 0, rsEmpresaPCE!intCveEmpresaPCE)
    End If
    rsEmpresaPCE.Close
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaEmpresaPCE"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    llngNumeroCuenta = 0
    lstrTipoPaciente = ""
    lstrNombreForma = ""
    llngNumeroCarta = 0
End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then SendKeys vbTab
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_KeyDown"))
End Sub

Private Sub mskFechaFin_LostFocus()
On Error GoTo NotificaError

    If IsDate(mskFechaFin.Text) Then
        If mskFechaFin > fdtmServerFecha Then
            MsgBox SIHOMsg(40), vbExclamation, "Mensaje"
            mskFechaFin.Text = fdtmServerFecha
            pEnfocaMkTexto mskFechaFin
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        pEnfocaMkTexto mskFechaFin
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_KeyDown"))
End Sub

Private Sub mskFechaInicio_LostFocus()
On Error GoTo NotificaError

    If IsDate(mskFechaInicio.Text) Then
        If mskFechaInicio > fdtmServerFecha Then
            MsgBox SIHOMsg(40), vbExclamation, "Mensaje"
            mskFechaInicio.Text = fdtmServerFecha
            pEnfocaMkTexto mskFechaInicio
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFechaInicio_LostFocus"))
End Sub

Private Sub optCargo_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = vbKeyReturn Then
        pEnfocaTextBox txtMensaje
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optCargo_KeyPress"))
End Sub

Private Sub optEstadoCuenta_Click(Index As Integer)
    Dim vlstrCartaDefault As String
    Dim rs As New ADODB.Recordset
    'chkRangoFechas.Value = 0 'caso 20362
    chkMostrarConSeguroSF.Value = 0
    chkMostrarConSeguroSF.Enabled = False
  
    If Not (optOrden(5).Value Or optOrden(6).Value) Then
      cboFactura.Enabled = True
      chkRangoFechas.Enabled = True
      chkHora.Enabled = True
      chkPagos.Enabled = True
      'chkCosto.Enabled = True
      chkMostrarCirugias.Enabled = True
      'chkDesglosarPaquete.Enabled = True
      Label8.Enabled = True
      Label10.Enabled = False
      Label11.Enabled = False
      Select Case Index
         Case 0
            If vlblnUtilizaConvenio Then
                optCargo(2).Value = True
                cboFactura.Enabled = True
            Else
                optCargo(1).Value = True
                cboFactura.Enabled = True
            End If
            chkMostrarCuatroDecimales.Enabled = False
            chkMostrarCuatroDecimales.Value = 0
            If gintAseguradora = 1 And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then chkMostrarConSeguroSF.Enabled = True
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                                
                If blnCartaEncontrada = True Then
                    cboCartas.ListIndex = 0
                    
                Else
                    cboCartas.ListIndex = -1
                End If
            Else
                Dim contlabel As Integer
                Dim j As Integer
                contlabel = 0
                For j = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(j) = "<TODOS>" Then
                        contlabel = contlabel + 1
                    End If
                Next
                If contlabel = 0 Then
                    cboCartas.AddItem "<TODOS>", 0
                End If
'                cboCartas.Enabled = False
'                lblcarta.Enabled = False
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                Else
                    cboCartas.ListIndex = -1
                End If
            End If
         Case 1
            'If vlblnUtilizaConvenio Then
                 optCargo(1).Value = True
                 cboFactura.Enabled = True
            'Else
            '    OptCargo(1).Value = True
            '    cboFactura.Enabled = True
            'End If
            chkMostrarCuatroDecimales.Enabled = True
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                
                 If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                Else
                    cboCartas.ListIndex = -1
                End If
            Else
                Dim G As Integer
                For G = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(G) = "<TODOS>" Then
                        cboCartas.RemoveItem (G)
                        Exit For
                    End If
                Next
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        If cboCartas.ListCount > 0 Then
                            cboCartas.ListIndex = 0
                            vlstrCartaDefault = "SELECT intcvecarta, vchdescripcion FROM PVCARTACONTROLSEGURO WHERE INTNUMCUENTA = '" & Trim(txtMovimientoPaciente.Text) & "' and bitdefault = 1 and chrEstatus <> 'I'"
                            Set rs = frsRegresaRs(vlstrCartaDefault, adLockOptimistic, adOpenDynamic)
                            If rs.RecordCount > 0 Then
                                cboCartas.ListIndex = flngLocalizaCbo(cboCartas, rs!intCveCarta)
                            Else
                                cboCartas.ListIndex = 0
                            End If
                        End If
                    End If
                Else
                    cboCartas.ListIndex = -1
                End If
            End If
         Case 2
            optCargo(0).Value = True
            cboFactura.ListIndex = 0
            cboFactura.Enabled = False
            chkMostrarCuatroDecimales.Enabled = True
            
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                Else
                    cboCartas.ListIndex = -1
                End If
            Else
                Dim h As Integer
                For h = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(h) = "<TODOS>" Then
                        cboCartas.RemoveItem (h)
                        Exit For
                    End If
                Next
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                If lstrNombreForma = "frmFacturacion" Then
                    cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                Else
                    If cboCartas.ListCount > 0 Then
                        cboCartas.ListIndex = 0
                        vlstrCartaDefault = "SELECT intcvecarta, vchdescripcion FROM PVCARTACONTROLSEGURO WHERE INTNUMCUENTA = '" & Trim(txtMovimientoPaciente.Text) & "' and bitdefault = 1 and chrEstatus <> 'I'"
                        Set rs = frsRegresaRs(vlstrCartaDefault, adLockOptimistic, adOpenDynamic)
                        If rs.RecordCount > 0 Then
                            cboCartas.ListIndex = flngLocalizaCbo(cboCartas, rs!intCveCarta)
                        Else
                            If blnCartaEncontrada = True Then
                                cboCartas.ListIndex = 0
                            Else
                                cboCartas.ListIndex = -1
                            End If
                        End If
                    End If
                End If
                
            End If
      End Select
      
    Else
    Select Case Index
         Case 0
            
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = False
                lblcarta.Enabled = False
                                                
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                    blnCartaEncontrada = True
                Else
                '*aqui
                    cboCartas.Enabled = True
                lblcarta.Enabled = True
                End If
            Else
                Dim contlabel1 As Integer
                Dim j1 As Integer
                contlabel = 0
                For j1 = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(j1) = "<TODOS>" Then
                        contlabel1 = contlabel1 + 1
                    End If
                Next
                If contlabel1 = 0 Then
                    cboCartas.AddItem "<TODOS>", 0
                End If
                cboCartas.Enabled = False
                lblcarta.Enabled = False
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                    cboCartas.Enabled = True
                lblcarta.Enabled = True
                Else
                    cboCartas.ListIndex = -1
                End If
            End If
         Case 1
            optCargo(0).Value = False
            optCargo(2).Value = False
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = False
                lblcarta.Enabled = False
                                
                 If blnCartaEncontrada = True Then
                    cboCartas.ListIndex = 0
                    'aqui
                    cboCartas.Enabled = True
                lblcarta.Enabled = True
                Else
                    cboCartas.ListIndex = -1
                End If
                
            Else
                Dim g1 As Integer
                For g1 = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(g1) = "<TODOS>" Then
                        cboCartas.RemoveItem (g1)
                        Exit For
                    End If
                Next
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                         If cboCartas.ListCount > 0 Then
                            cboCartas.ListIndex = 0
                            vlstrCartaDefault = "SELECT intcvecarta, vchdescripcion FROM PVCARTACONTROLSEGURO WHERE INTNUMCUENTA = '" & Trim(txtMovimientoPaciente.Text) & "' and bitdefault = 1 and chrEstatus <> 'I'"
                            Set rs = frsRegresaRs(vlstrCartaDefault, adLockOptimistic, adOpenDynamic)
                            If rs.RecordCount > 0 Then
                                cboCartas.ListIndex = flngLocalizaCbo(cboCartas, rs!intCveCarta)
                            Else
                                cboCartas.ListIndex = 0
                            End If
                        End If
                    End If
                Else
                    cboCartas.ListIndex = -1
                End If
            End If
         Case 2
           
            
            If cboCartas.ListCount = 0 Then
                cboCartas.Enabled = False
                lblcarta.Enabled = False
                                
                If blnCartaEncontrada = True Then
                    If lstrNombreForma = "frmFacturacion" Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                    Else
                        cboCartas.ListIndex = 0
                    End If
                    
                    cboCartas.Enabled = True
                lblcarta.Enabled = True
                Else
                    cboCartas.ListIndex = -1
                End If
            Else
                Dim h1 As Integer
                For h1 = 0 To cboCartas.ListCount - 1
                    If cboCartas.List(h1) = "<TODOS>" Then
                        cboCartas.RemoveItem (h1)
                        Exit For
                    End If
                Next
                cboCartas.Enabled = True
                lblcarta.Enabled = True
                                
                If lstrNombreForma = "frmFacturacion" Then
                    cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                Else
                    'cboCartas.ListIndex = 0
                    If cboCartas.ListIndex <> -1 Then cboCartas.ListIndex = 0
                    vlstrCartaDefault = "SELECT intcvecarta, vchdescripcion FROM PVCARTACONTROLSEGURO WHERE INTNUMCUENTA = '" & Trim(txtMovimientoPaciente.Text) & "' and bitdefault = 1 and chrEstatus <> 'I'"
                    Set rs = frsRegresaRs(vlstrCartaDefault, adLockOptimistic, adOpenDynamic)
                    If rs.RecordCount > 0 Then
                        cboCartas.ListIndex = flngLocalizaCbo(cboCartas, rs!intCveCarta)
                    Else
                        If blnCartaEncontrada = True Then
                            cboCartas.ListIndex = 0
                        Else
                            cboCartas.ListIndex = -1
                        End If
                    End If
                End If
            End If
      End Select
      
       ''*
        If (optOrden(6).Value And optEstadoCuenta(0).Value) Or (optOrden(6).Value And optEstadoCuenta(1).Value) Or (optOrden(6).Value And optEstadoCuenta(2).Value) Then
            chkMostrarCirugias.Enabled = False
            chkRangoFechas.Enabled = True
            chkRangoFechas.Value = 0
            mskFechaInicio.Enabled = False
            mskFechaFin.Enabled = False
            Label10.Enabled = False
            Label11.Enabled = False
            If gintAseguradora = 1 And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then
               chkMostrarConSeguroSF.Value = 0
               chkMostrarConSeguroSF.Enabled = True
            Else
               chkMostrarConSeguroSF.Value = 0
               chkMostrarConSeguroSF.Enabled = False
            End If
            Label8.Enabled = True
            cboFactura.Enabled = True
            chkHora.Enabled = True
            chkPagos.Enabled = True
            If blnCostos Then chkCosto.Enabled = True Else chkCosto.Enabled = False
            chkDesglosarPaquete.Enabled = True
            chkDesglosarPaquete.Value = 0
            If (optOrden(6).Value And optEstadoCuenta(0).Value) Then
                chkMostrarCuatroDecimales.Enabled = False
                chkMostrarCuatroDecimales.Value = 0
            Else
                chkMostrarCuatroDecimales.Enabled = True
                chkMostrarCuatroDecimales.Value = 0
            End If
            
            chkMostrarCirugias.Enabled = True
        
        Else
    
            cboFactura.Enabled = False
            chkMostrarCuatroDecimales.Enabled = False
            chkMostrarCuatroDecimales.Value = 0
            chkMostrarCirugias.Enabled = False
            If (optOrden(5).Value) Then
                chkRangoFechas.Enabled = False
            Else
                chkRangoFechas.Enabled = True
            End If
            mskFechaInicio.Enabled = False
            mskFechaFin.Enabled = False
            chkHora.Enabled = False
            chkPagos.Enabled = False
            chkDesglosarPaquete.Enabled = False
            chkDesglosarPaquete.Value = 0
            Label8.Enabled = False
            Label10.Enabled = False
            Label11.Enabled = False
            chkMostrarConSeguroSF.Value = 0
            chkMostrarConSeguroSF.Enabled = False
        End If
      

    End If
    
    cboFactura.Clear
    
    If Val(txtMovimientoPaciente.Text) > 0 And Not optEstadoCuenta(2).Value Then
        vgstrParametrosSP = CStr(Val(txtMovimientoPaciente.Text)) & _
                            "|" & IIf(optEstadoCuenta(0).Value, "P", "E") & _
                            "|" & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
        Set rsFacturas = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelFolioFacturasCuenta")
        If rsFacturas.RecordCount <> 0 Then
            Do While Not rsFacturas.EOF
                cboFactura.AddItem rsFacturas!chrfoliofactura
                cboFactura.ItemData(cboFactura.newIndex) = 0
                rsFacturas.MoveNext
            Loop
        End If
        rsFacturas.Close
    End If
        
    cboFactura.AddItem "<TODAS>", 0
    cboFactura.ItemData(cboFactura.newIndex) = -1
    cboFactura.ListIndex = 0
   
    
End Sub

Private Sub optEstadoCuenta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not (optOrden(5).Value Or optOrden(6).Value) Then
            If Me.cboFactura.Enabled = True Then
                Me.cboFactura.SetFocus
            Else
                Me.chkPagos.SetFocus
            End If
        Else
            txtMensaje.SetFocus
        End If
    End If
End Sub

Private Sub optEstadoCuentaHospitalFarm_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not (optOrden(5).Value Or optOrden(6).Value) Then
            If Me.cboFactura.Enabled = True Then
                Me.cboFactura.SetFocus
            Else
                Me.chkPagos.SetFocus
            End If
        Else
            txtMensaje.SetFocus
        End If
    End If
End Sub

Private Sub optOrden_Click(Index As Integer)

'*********************+caso 20362*************
'chkRangoFechas.Value = 0


        If Not (optOrden(5).Value Or optOrden(6).Value) Then
            cboFactura.Enabled = True
            chkRangoFechas.Enabled = True
            chkHora.Enabled = True
            chkPagos.Enabled = True
            If blnCostos Then
               chkCosto.Enabled = True
            Else
                chkCosto.Enabled = False
            End If
            chkMostrarCirugias.Enabled = True
            Label8.Enabled = True
            Label10.Enabled = False
            Label11.Enabled = False
            If vlblnDesglosarPaquete = False Then
                chkDesglosarPaquete.Enabled = True
            End If
            If optEstadoCuenta(0) Then
                chkMostrarCuatroDecimales.Enabled = False
                chkMostrarCuatroDecimales.Value = 0
                If gintAseguradora = 1 And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then
                    chkMostrarConSeguroSF.Enabled = True
                Else
                    chkMostrarConSeguroSF.Enabled = False
                End If
            Else
                chkMostrarCuatroDecimales.Enabled = True
            End If
            blnEstadoAgrupado = False
        Else
             If optOrden(5).Value Then
                cboFactura.Enabled = False
                If (optOrden(5).Value) Then
                    chkRangoFechas.Enabled = False
                Else
                    chkRangoFechas.Enabled = True
                End If
                mskFechaInicio.Enabled = False
                mskFechaFin.Enabled = False
                chkHora.Enabled = False
                chkPagos.Enabled = False
                chkCosto.Enabled = False
                chkMostrarCirugias.Enabled = False
                Label8.Enabled = False
                Label10.Enabled = False
                Label11.Enabled = False
                If chkDesglosarPaquete.Enabled = True And vlblnDesglosarPaquete = True Then
                   vlblnDesglosarPaquete = False
                End If
                chkMostrarCuatroDecimales.Enabled = False
                chkMostrarCuatroDecimales.Value = 0
                chkDesglosarPaquete.Enabled = False
                chkDesglosarPaquete.Value = 0
                chkMostrarConSeguroSF.Value = 0
                chkMostrarConSeguroSF.Enabled = False
                
                blnEstadoAgrupado = False
            Else
                chkRangoFechas.Enabled = True
                chkRangoFechas.Value = 0
                mskFechaInicio.Enabled = False
                mskFechaFin.Enabled = False
                optCargo(2).Value = False
                Label10.Enabled = False
                Label11.Enabled = False
                If chkDesglosarPaquete.Enabled = True And vlblnDesglosarPaquete = True Then
                    vlblnDesglosarPaquete = False
                End If
                If (optOrden(6).Value And optEstadoCuenta(0).Value) Then
                    If gintAseguradora = 1 And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then
                       chkMostrarConSeguroSF.Enabled = True
                    Else
                       chkMostrarConSeguroSF.Enabled = False
                    End If
                
                End If
                
                Label8.Enabled = True
                cboFactura.Enabled = True
                chkHora.Enabled = True
                chkPagos.Enabled = True
               If blnCostos Then chkCosto.Enabled = True Else chkCosto.Enabled = False
                chkDesglosarPaquete.Enabled = True
                chkDesglosarPaquete.Value = 0
                If (optOrden(6).Value And optEstadoCuenta(0).Value) Then
                    chkMostrarCuatroDecimales.Enabled = False
                    chkMostrarCuatroDecimales.Value = 0
                Else
                    chkMostrarCuatroDecimales.Enabled = True
                    chkMostrarCuatroDecimales.Value = 0
                End If
                chkMostrarCirugias.Enabled = True
                blnEstadoAgrupado = True
            End If
        End If
       
        '*********************+caso 20362*************
'chkRangoFechas.Value = 0
    If chkRangoFechas.Value = vbChecked And chkRangoFechas.Enabled = True Then
        Label10.Enabled = True
        Label11.Enabled = True
        mskFechaFin.Enabled = True
        mskFechaInicio.Enabled = True
    Else
        Label10.Enabled = False
        Label11.Enabled = False
        mskFechaFin.Enabled = False
        mskFechaInicio.Enabled = False
    End If

        '*********************+caso 20362*************
End Sub

Private Sub optOrden_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = 13 Then
      '  SendKeys vbTab
'*7092023
  If chkRangoFechas.Enabled Then
            chkRangoFechas.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
'*
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optOrden_KeyPress"))
End Sub

Private Sub optTipoValidacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtMensaje_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And txtMovimientoPaciente.Text <> "" Then
        TxtNotasInternas.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMensaje_KeyDown"))

End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMensaje_KeyPress"))
End Sub

Private Sub txtMovimientoPaciente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        If Asc(UCase(Chr(KeyAscii))) = vbKeyI Then
            OptTipoPaciente(0).Value = True
        ElseIf Asc(UCase(Chr(KeyAscii))) = vbKeyE Then
            OptTipoPaciente(1).Value = True
        End If
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyPress"))
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
On Error GoTo NotificaError
    
    If Not vlblnConsulta Then
        txtMovimientoPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":OptTipoPaciente_Click"))
End Sub

Private Sub optTipoPaciente_GotFocus(Index As Integer)
On Error GoTo NotificaError
    
    pLimpia

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_GotFocus"))
End Sub

Private Sub optTipoPaciente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtMovimientoPaciente.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTipoPaciente_KeyPress"))
End Sub

Private Sub pLimpia()
1     On Error GoTo NotificaError
          
2         FraTipos.Enabled = True
3         FraTipos.Visible = True
          
4         FraTiposHospitalFarmacia.Enabled = False
5         FraTiposHospitalFarmacia.Visible = False
          
6         chkRangoFechas.Value = 0
7         Label10.Enabled = False
8         Label11.Enabled = False
9         mskFechaFin.Enabled = False
10        mskFechaInicio.Enabled = False
          
11        txtFechaFinal.Text = ""
12        txtFechaInicial.Text = ""
          
13        vlblnPacienteSeleccionado = False
          
14        cboFactura.Enabled = True
15        chkMostrarCuatroDecimales.Enabled = True
16        chkRangoFechas.Enabled = True
17        chkHora.Enabled = True
18        chkPagos.Enabled = True
          'chkCosto.Enabled = True
19        Label8.Enabled = True
          
20        txtPaciente.Text = ""
21        txtEmpresaPaciente.Text = ""
22        txtTipoPaciente.Text = ""
23        txtFechaFinal.Text = ""
24        txtFechaInicial.Text = ""
25        cboFactura.Clear
26        optEstadoCuenta(0).Value = True

27

28        optEstadoCuenta(1).Enabled = False
29        optEstadoCuenta(2).Enabled = False
            vgempresapaciente1 = ""
30        chkDesglosarPaquete.Value = 0
31        chkDesglosarPaquete.Enabled = False
          
32        fraFTP.Visible = True
            cmdFTP.Enabled = False
        optTipoValidacion(0).Enabled = False
        optTipoValidacion(1).Enabled = False
          
33        chkFiltroEnvioRango.Enabled = False
34        chkFiltroEnvioFactura.Enabled = False
          
35        frmFiltroEnvio.Enabled = False
          
36        chkFiltroEnvioRango.Value = 0
37        chkFiltroEnvioFactura.Value = 0
          
38        txtMensaje.Text = ""
39        TxtNotasInternas.Text = ""

40        chkMostrarConSeguroSF.Value = False
41        chkMostrarConSeguroSF.Enabled = False
          
42        lblCuentaDe.Visible = False
43        txtMovimientoPacienteOtro.Visible = False
          
44        lblCuentaDe.Caption = "Cuenta de _______"
45        txtMovimientoPacienteOtro.Text = ""
46        blnEsHospitalMultiempresaFarm = False
47        blnEsFarmaciaMultiempresaFarm = False
        cboCartas.Clear
            blnCartaEncontrada = False

48    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia" & " Linea:" & Erl()))
End Sub

Private Sub txtMovimientoPaciente_GotFocus()
On Error GoTo NotificaError
    
    pHabilitaFiltros
    vlblnDesglosarPaquete = True
    If vlblnLimpiar And llngNumeroCuenta = 0 Then
        pLimpia
    Else
        vlblnLimpiar = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_GotFocus"))
End Sub

Private Sub txtMovimientoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
1     On Error GoTo NotificaError
          Dim querycartas As String
          Dim rsCartas As New ADODB.Recordset
          Dim rs As New ADODB.Recordset
2         blnDatosCuenta = False
3         If KeyCode = vbKeyReturn Then
4             If RTrim(txtMovimientoPaciente.Text) = "" Then
5                 With FrmBusquedaPacientes
6                     .vgblnPideClave = False
7                     .vgIntMaxRecords = 100
8                     .vgstrMovCve = "M"
                      
9                     .optTodos.Value = True
10                    .optSinFacturar.Enabled = True
11                    .optSoloActivos.Enabled = True
12                    .optTodos.Enabled = True
                      
                      ' condicion para no saturar con tantos registros dependiendo si es externo o interno
13                    If OptTipoPaciente(1).Value Then 'Externos
14                        .vgIntMaxRecords = 1000
15                    Else
16                        .vgIntMaxRecords = 500
17                    End If
                      
18                    If OptTipoPaciente(1).Value Then 'Externos
19                        .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
20                        .vgstrTamanoCampo = "800,3400,1700,4100"
21                        .vgstrTipoPaciente = "E"
22                        .Caption = .Caption & " Externos"
23                    Else
24                        .vgStrOtrosCampos = ", TO_CHAR(ExPacienteIngreso.dtmFechaHoraIngreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha ing."", TO_CHAR(ExPacienteIngreso.dtmFechaHoraEgreso, 'dd/mm/yyyy hh:mi:ss am') as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
25                        .vgstrTamanoCampo = "800,3400,990,990,4100"
26                        .vgstrTipoPaciente = "I"
27                        .Caption = .Caption & " Internos"
28                    End If
              
29                    txtMovimientoPaciente.Text = .flngRegresaPaciente()
30                    If txtMovimientoPaciente <> -1 Then
31                        vlblnLimpiar = False
32                        txtMovimientoPaciente_KeyDown vbKeyReturn, 0
33                    Else
34                        txtMovimientoPaciente.Text = ""
35                    End If
36                End With
37            Else
                    cboCartas.ListIndex = -1
                    
                    querycartas = "SELECT intcvecarta, vchdescripcion FROM PVCARTACONTROLSEGURO WHERE INTNUMCUENTA = '" & Trim(txtMovimientoPaciente.Text) & "' AND CHRESTATUS <> 'I' order by  bitdefault desc"
                    Set rsCartas = frsRegresaRs(querycartas, adLockOptimistic, adOpenDynamic)
                    If rsCartas.RecordCount > 0 Then
                        'pLlenarCboRs cboCartas, rsCartas, 0, 1, 3

                        Do Until rsCartas.EOF
                            cboCartas.AddItem rsCartas!VCHDESCRIPCION
                            cboCartas.ItemData(cboCartas.newIndex) = rsCartas.Collect("intcvecarta")
                            rsCartas.MoveNext
                        Loop
                        
                        If lstrNombreForma = "frmFacturacion" Then
                            cboCartas.ListIndex = flngLocalizaCbo(cboCartas, str(llngNumeroCarta))
                        Else
                            cboCartas.ListIndex = 0
                        End If
                        cboCartas.Enabled = True
                        blnCartaEncontrada = True
                        
                    Else
                        cboCartas.Clear
                        cboCartas.AddItem "<TODOS>", 0
                        cboCartas.Enabled = True
                        lblcarta.Enabled = True
                        blnCartaEncontrada = False
                    End If


38                If fblnDatosPaciente(Val(txtMovimientoPaciente.Text), IIf(OptTipoPaciente(0).Value, "I", "E")) Then
39                    vlblnPacienteSeleccionado = True
                      'Habilitar la opción de desglosar paquete:
40                    vgstrParametrosSP = txtMovimientoPaciente.Text & "|" & IIf(OptTipoPaciente(0).Value, "I", "E")
41                    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELPAQUETEPACIENTE")
42                    chkDesglosarPaquete.Enabled = rs.RecordCount <> 0
                      
43                    If optOrden(0).Value Then
44                        optOrden(0).SetFocus
45                    Else
46                        If optOrden(1).Value Then
47                            optOrden(1).SetFocus
48                        Else
49                            optOrden(2).SetFocus
50                        End If
51                    End If
52                End If
53            End If
54            blnDatosCuenta = True
55        End If



56    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMovimientoPaciente_KeyDown" & " Linea:" & Erl()))
End Sub

Private Function fblnDatosPaciente(vllngxMovimiento As Long, vlstrxTipoPaciente As String) As Boolean
1     On Error GoTo NotificaError
          
          Dim rsComentario As New ADODB.Recordset
          Dim rsFacturas As New ADODB.Recordset
          Dim vlstrSentencia As String
          Dim vlrsPvSelDatosPaciente As New ADODB.Recordset
          Dim rsRetencion As ADODB.Recordset
          
2         chkMostrarCuatroDecimales.Value = 0
3         vldblRetencion = 0
4         vgstrParametrosSP = vllngxMovimiento & "|" & "0" & "|" & vlstrxTipoPaciente & "|" & vgintClaveEmpresaContable
5         Set vlrsPvSelDatosPaciente = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELDATOSPACIENTE")
6         With vlrsPvSelDatosPaciente
7             If .RecordCount <> 0 Then
8                 fblnDatosPaciente = True
                  
9                 txtPaciente.Text = !Nombre
10                txtEmpresaPaciente.Text = IIf(IsNull(!empresa), "", !empresa)
                  
                  vgempresapaciente1 = txtEmpresaPaciente.Text
11                lintCveEmpresaPaciente = !intcveempresa
                  
12                If !intcveempresa <> 0 Then
13                    vlstrSentencia = "SELECT NVL(RELPORCENTAJESERVICIOSEMP, 0) RETENCION FROM CCEMPRESA WHERE INTCVEEMPRESA = " & !intcveempresa
14                    Set rsRetencion = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
15                    If rsRetencion.RecordCount > 0 Then vldblRetencion = rsRetencion!Retencion
16                End If
                  
17                txtTipoPaciente.Text = !tipo
18                txtFechaInicial.Text = ""
19                If Not IsNull(!Ingreso) Then
20                    txtFechaInicial.Text = Format(!Ingreso, "dd/mmm/yyyy hh:mm")
21                End If
22                txtFechaFinal.Text = ""
23                If Not IsNull(!Egreso) Then
24                    txtFechaFinal.Text = Format(!Egreso, "dd/mmm/yyyy hh:mm")
25                End If
                  
                  'Agregar la Fecha de inicio para el filtro del reporte de estado de cuenta
26                If Not IsNull(!Ingreso) Then
27                    mskFechaInicio.Text = Format(!Ingreso, "dd/mm/yyyy")
28                End If
                  
                  'Agregar la Fecha de término para el filtro del reporte de estado de cuenta
29                mskFechaFin.Text = fdtmServerFecha
                  
                  'Comentario guardado para el estado de cuenta
30                vlstrSentencia = "Select vchMensaje Mensaje, VCHNOTASINTERNAS NOTAS from pvMensajeEstadoCuenta where intMovPaciente = " & Trim(str(vllngxMovimiento)) & _
                                  " and chrTipoPaciente = '" & vlstrxTipoPaciente & "'"
                                  
31                Set rsComentario = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
32                If rsComentario.RecordCount > 0 Then
33                    txtMensaje.Text = Trim(IIf(IsNull(rsComentario!Mensaje), "", rsComentario!Mensaje))
34                    TxtNotasInternas.Text = Trim(IIf(IsNull(rsComentario!Notas), "", rsComentario!Notas))
35                Else
36                    txtMensaje.Text = ""
37                    TxtNotasInternas.Text = ""
38                End If
39                rsComentario.Close
                      
40                vlblnUtilizaConvenio = IIf(!bitUtilizaConvenio = 0, False, True)
                  '-- Indica si la empresa de convenio del paciente es aseguradora
41                gintAseguradora = IIf((!bitUtilizaConvenio = 1 And !Aseguradora = 1), 1, 0)
42                chkMostrarConSeguroSF.Value = False
43                If gintAseguradora = 1 And cboFactura.ListIndex = 0 And (cgstrModulo = "PV" Or cgstrModulo = "CC") Then
44                    chkMostrarConSeguroSF.Enabled = True
45                Else
46                    chkMostrarConSeguroSF.Enabled = False
47                End If
48                Me.optEstadoCuenta(0).Value = True
49                optEstadoCuenta_Click 0
50                If vlblnUtilizaConvenio Then
51                    optEstadoCuenta(1).Enabled = True
52                    optEstadoCuenta(2).Enabled = True
                      
                      '-- Habilita opción para indicar si el detalle y totales del reporte son mostrados con 4 decimales
                      '-- si no se selecciona la opción se mostrarán como normalmente se hace a 2 decimales
      '                If cboFactura.ListIndex > -1 Then
      '                    'If Not optEstadoCuenta(0).Value Then ' solamente cuando NO es el estado de cuenta del paciente
      '                        vgstrParametrosSP = Trim(txtMovimientoPaciente.Text) & _
      '                                         "|" & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & _
      '                                         "|" & lintCveEmpresaPaciente & _
      '                                         "|" & vgintClaveEmpresaContable & _
      '                                         "|" & IIf(cboFactura.ItemData(cboFactura.ListIndex) = -1, "*", cboFactura.List(cboFactura.ListIndex))
      '                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelGruposEdoCuenta")
      '
      '                        If rs.RecordCount > 0 Then
      '                            Do While Not rs.EOF
      '                                If Not IsNull(rs!NumGrupo) Then
      '                                    chkMostrarCuatroDecimales.Visible = True
      '                                End If
      '                                rs.MoveNext
      '                            Loop
      '                        End If
      '                    'End If
      '                End If
53                Else
54                    optEstadoCuenta(1).Enabled = False
55                    optEstadoCuenta(2).Enabled = False
56                End If
                  
57                pInterfazFTP !intcveempresa
                  
58                If fblnFarmaciaMultiempresa Then
59                    pCargaOtraCuentaMultiempresaFarmacia
60                End If
                  If gintAseguradora = 1 Then
                      pAsignarCargosSinCarta Trim(txtMovimientoPaciente.Text), lintCveEmpresaPaciente
                      cboCartas_KeyDown vbKeyReturn, 0
                  End If
61            Else
62                fblnDatosPaciente = False
                  '¡La información no existe!
63                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
64                pEnfocaTextBox txtMovimientoPaciente
65            End If
66            .Close
67        End With

68    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosPaciente" & " Linea:" & Erl()))
End Function

Private Sub pInterfazFTP(lngCveEmpresa As Long)
    Dim rsInterfaz As ADODB.Recordset
    Dim strSql As String
    
    vlintCveInterfazATC = 0
    
    strSql = "select * from PVFTPEstadoCuenta " & _
    "inner join PVFTPEstadoCuentaDetalle on PVFTPEstadoCuentaDetalle.INTCVEINTERFAZ = PVFTPEstadoCuenta.INTCVEINTERFAZ " & _
    "where PVFTPEstadoCuentaDetalle.intCveEmpresa = " & lngCveEmpresa & " and PVFTPEstadoCuenta.bitactivo = 1"
    Set rsInterfaz = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rsInterfaz.EOF Then
        vlintCveInterfazATC = rsInterfaz!intcveinterfaz
    
        optTipoValidacion(0).Value = True
        '7092023
        fraFTP.Visible = True
        '*
        cmdFTP.Enabled = True
        optTipoValidacion(0).Enabled = True
        optTipoValidacion(1).Enabled = True
        frmFiltroEnvio.Enabled = rsInterfaz!BITUSAFILTROS = 1
        
        chkFiltroEnvioRango.Enabled = rsInterfaz!BITUSAFILTROS = 1
        chkFiltroEnvioFactura.Enabled = rsInterfaz!BITUSAFILTROS = 1
        
        chkCataCargosEmpresa.Value = rsInterfaz!BITCARGOSPOREMPRESA
    Else
        vlintCveInterfazATC = 0
    
        'fraFTP.Visible = False
        cmdFTP.Enabled = False
        optTipoValidacion(0).Enabled = False
        optTipoValidacion(1).Enabled = False
        frmFiltroEnvio.Enabled = False
        
        chkFiltroEnvioRango.Enabled = False
        chkFiltroEnvioFactura.Enabled = False
    End If
    rsInterfaz.Close
End Sub

Public Sub pHabilitaFiltros()
    cboFactura.Enabled = True
    chkRangoFechas.Enabled = True
    chkHora.Enabled = True
    chkPagos.Enabled = True
    'chkCosto.Enabled = True
    Label8.Enabled = True
    Label10.Enabled = False
    Label11.Enabled = False
End Sub

Private Sub TxtNotasInternas_GotFocus()
    If TxtNotasInternas.Locked = False Then
        TxtNotasInternas.FontBold = False
        strNotasInternas = ""
        If TxtNotasInternas = "" And blnDatosCuenta = True Then
            strFechaNota = Format$(Date, "dd/mmm/yyyy") & ":"
            TxtNotasInternas.Text = Trim(strFechaNota)
            TxtNotasInternas.SelStart = Len(TxtNotasInternas.Text)
            TxtNotasInternas.SetFocus
        End If
    End If
End Sub

Private Sub TxtNotasInternas_KeyPress(KeyAscii As Integer)
1     On Error GoTo NotificaError

2         KeyAscii = Asc(UCase(Chr(KeyAscii)))
3         If KeyAscii = 13 Then
4             strFechaNota = Replace(Format$(Date, "dd/mmm/yyyy") & ":", vbCrLf, "")
5             strNotasInternas = Trim(TxtNotasInternas.Text)
6             If TxtNotasInternas.Text = "" And blnDatosCuenta = True Then
7                 TxtNotasInternas.Text = ""
8                 TxtNotasInternas.Text = Trim(strNotasInternas) & strFechaNota
9                 TxtNotasInternas.SelStart = Len(TxtNotasInternas.Text)
10                TxtNotasInternas.SetFocus
11            Else
12                If txtMovimientoPaciente <> "" Then
13                    TxtNotasInternas.Text = Replace(Trim(strNotasInternas) & vbCrLf & RTrim(strFechaNota), vbCrLf & vbCrLf, vbCrLf)
14                    TxtNotasInternas.SelStart = Len(Trim$(TxtNotasInternas.Text))
15                End If
16            End If
17            KeyAscii = 7
18        End If
          
19        If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab And txtMovimientoPaciente.Text <> "" Then
20            cmdPreview.SetFocus
21        End If

22    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtNotasInternas_KeyPress" & " Linea:" & Erl()))
End Sub

Private Sub TxtNotasInternas_LostFocus()
    If Trim(txtMovimientoPaciente.Text) <> "" And TxtNotasInternas.Locked = False Then
        pEjecutaSentencia "DELETE PvMensajeEstadoCuenta WHERE intMovPaciente = " & Trim(txtMovimientoPaciente.Text) & " AND chrTipoPaciente = " & IIf(OptTipoPaciente(0).Value, "'I'", "'E'")
        pEjecutaSentencia "INSERT INTO pvMensajeEstadoCuenta VALUES(" & Trim(txtMovimientoPaciente.Text) & "," & IIf(OptTipoPaciente(0).Value, "'I'", "'E'") & ",'" & Trim(txtMensaje.Text) & "','" & IIf(Len(Trim(TxtNotasInternas.Text)) > 11, TxtNotasInternas.Text, "") & "')"
    End If
End Sub

Private Function fblnFarmaciaMultiempresa() As Boolean
1     On Error GoTo NotificaError
          
          Dim rs As New ADODB.Recordset
          Dim vlstrSentencia As String
          
2         lblCuentaDe.Visible = False
3         txtMovimientoPacienteOtro.Visible = False
4         lblCuentaDe.Caption = "Cuenta de _______"
5         txtMovimientoPacienteOtro.Text = ""
6         blnEsHospitalMultiempresaFarm = False
7         blnEsFarmaciaMultiempresaFarm = False
8         fblnFarmaciaMultiempresa = False
          
9         Set rs = frsRegresaRs("Select bitactivo, tnyclaveempresahospital, tnyclaveempresafarmacia from GNPARAMETROSFARMACIA where bitactivo = 1", adLockReadOnly, adOpenForwardOnly)
10        If rs.RecordCount = 0 Then
11            fblnFarmaciaMultiempresa = False
12        Else
13            If rs!bitactivo = 1 Then
14                If rs!tnyclaveempresahospital = vgintClaveEmpresaContable Then
15                    fblnFarmaciaMultiempresa = True
16                    blnEsHospitalMultiempresaFarm = True
17                    blnEsFarmaciaMultiempresaFarm = False
18                    lblCuentaDe.Caption = "Cuenta de farmacia"
19                End If
20                If rs!TNYCLAVEEMPRESAFARMACIA = vgintClaveEmpresaContable Then
21                    fblnFarmaciaMultiempresa = True
22                    blnEsHospitalMultiempresaFarm = False
23                    blnEsFarmaciaMultiempresaFarm = True
24                    lblCuentaDe.Caption = "Cuenta de hospital"
25                End If
26            End If
27        End If

28    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnFarmaciaMultiempresa" & " Linea:" & Erl()))
End Function

Private Sub pCargaOtraCuentaMultiempresaFarmacia()
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    If blnEsHospitalMultiempresaFarm = False And blnEsFarmaciaMultiempresaFarm = False Then
    
    Else
        If lintCveEmpresaPaciente <> 0 Then Exit Sub
        
        If blnEsHospitalMultiempresaFarm = True And blnEsFarmaciaMultiempresaFarm = False Then
            strSql = "SELECT intNumCuentaAdmision, intNumCuentaFarmacia FROM ADRELACIONCUENTASFARMACIA where intNumCuentaAdmision = " & Val(txtMovimientoPaciente.Text)
            Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
            If Not rs.EOF Then
                lblCuentaDe.Visible = True
                txtMovimientoPacienteOtro.Visible = True
                txtMovimientoPacienteOtro.Text = rs!intNumCuentaFarmacia
                
                optEstadoCuenta_Click 0
                FraTipos.Enabled = False
                FraTipos.Visible = False
                
                FraTiposHospitalFarmacia.Enabled = True
                FraTiposHospitalFarmacia.Visible = True
                
                optEstadoCuentaHospitalFarm(0).Value = True
            End If
            rs.Close
        Else
            If blnEsHospitalMultiempresaFarm = False And blnEsFarmaciaMultiempresaFarm = True Then
                strSql = "SELECT intNumCuentaAdmision, intNumCuentaFarmacia FROM ADRELACIONCUENTASFARMACIA where intNumCuentaFarmacia = " & Val(txtMovimientoPaciente.Text)
                Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
                If Not rs.EOF Then
                    lblCuentaDe.Visible = True
                    txtMovimientoPacienteOtro.Visible = True
                    txtMovimientoPacienteOtro.Text = rs!intNumCuentaAdmision
                    
                    optEstadoCuenta_Click 0
                    FraTipos.Enabled = False
                    FraTipos.Visible = False
                    
                    FraTiposHospitalFarmacia.Enabled = True
                    FraTiposHospitalFarmacia.Visible = True
                    
                    optEstadoCuentaHospitalFarm(0).Value = True
                End If
                rs.Close
            End If
        End If
    End If
End Sub

Private Function fblnATCConCargosEmpresa() As Boolean
    Dim strSql As String
    Dim strEncriptado As String
    Dim rsTemp As ADODB.Recordset
    
    fblnATCConCargosEmpresa = False
    
    strSql = "SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR FROM SIPARAMETRO WHERE VCHNOMBRE = 'VCHATCCATCARGOSEMPRESA'"
    Set rsTemp = frsRegresaRs(strSql)
    If Not rsTemp.EOF Then
        If fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 7028, 7029), "C", True) Or fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 2, 7028, 7029), "E", True) Then
            fblnATCConCargosEmpresa = IIf(IIf(IsNull(rsTemp!Valor), 0, rsTemp!Valor) = "1", True, False)
        End If
    End If
End Function


Public Sub pAsignarCargosSinCarta(lngnumCuenta As Long, intEmpresa As Integer)
    Dim vlstrSentencia As String
    Dim rsPrimeraCartaActiva As ADODB.Recordset
    Dim rsCargosSinCarta As ADODB.Recordset
    Dim rsFacturasSinCarta As ADODB.Recordset
    Dim rsControlSinCarta As ADODB.Recordset

    vgblnCambiaCarta = False
    
    'Elimina de los cargos cartas asignadas que se encuentren en estatus de 'Inactiva'
    vlstrSentencia = "UPDATE PvCargo set intcvecarta = null where intmovpaciente = " & lngnumCuenta & _
                     " and chrfoliofactura is null " & _
                     " and intcvecarta in ( select pvcartacontrolseguro.intcvecarta from PvCartaControlSeguro " & _
                                            "where chrestatus = 'I' and intnumcuenta = pvcargo.intmovpaciente) "
                                            '"and intCveEmpresa = " & intEmpresa & ")"
    pEjecutaSentencia (vlstrSentencia)
    
    vlstrSentencia = "select Nvl(Min(intCveCarta), 0) numCarta from PVCARTACONTROLSEGURO where intNumCuenta = " & lngnumCuenta & " AND chrEstatus = 'A' and intCveEmpresa = " & intEmpresa
    Set rsPrimeraCartaActiva = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If rsPrimeraCartaActiva.RecordCount > 0 Then
        If rsPrimeraCartaActiva!numCarta > 0 Then
            
            'Se asigna a los cargos que no tienen carta, la primera de las cartas activas de la cuenta del paciente
            vlstrSentencia = "select count(*) cargosSinCarta from PVCARGO where intMovPaciente = " & lngnumCuenta & " and chrfoliofactura is null AND intCveCarta is null"
            Set rsCargosSinCarta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rsCargosSinCarta.RecordCount > 0 Then
                If rsCargosSinCarta!cargosSinCarta > 0 Then
                    pEjecutaSentencia ("UPDATE PvCargo set intcvecarta = " & rsPrimeraCartaActiva!numCarta & " where intmovpaciente = " & lngnumCuenta & " and chrfoliofactura is null and intcvecarta is null")
                End If
            End If
                
            'Se asigna a las facturas que no tienen carta, la primera de las cartas activas de la cuenta del paciente
            vlstrSentencia = "select count(*) facturasSinCarta from PVFACTURA where intMovPaciente = " & lngnumCuenta & " AND intCveCarta is null and chrEstatus <> 'C'"
            Set rsFacturasSinCarta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rsFacturasSinCarta.RecordCount > 0 Then
                If rsFacturasSinCarta!facturasSinCarta > 0 Then
                    pEjecutaSentencia ("UPDATE PvFactura set intcvecarta = " & rsPrimeraCartaActiva!numCarta & " where intmovpaciente = " & lngnumCuenta & " and intcvecarta is null and chrEstatus <> 'C'")
                End If
            End If
            
            'Se asigna a el control de aseguradora que no tiene carta, la primera de las cartas activas de la cuenta del paciente
            vlstrSentencia = "select count(*) controlSinCarta from PVCONTROLASEGURADORA where intMovPaciente = " & lngnumCuenta & " AND intCveCarta is null and intCveEmpresa = " & intEmpresa
            Set rsControlSinCarta = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            If rsControlSinCarta.RecordCount = 1 Then
                If rsControlSinCarta!controlSinCarta > 0 Then
                    pEjecutaSentencia ("UPDATE PvControlAseguradora set intcvecarta = " & rsPrimeraCartaActiva!numCarta & " where intmovpaciente = " & lngnumCuenta & " and intcvecarta is null and intCveEmpresa = " & intEmpresa)
                End If
            End If
            
            'pCambiaCarta
            vgblnCambiaCarta = True
            
        End If
    End If

End Sub


