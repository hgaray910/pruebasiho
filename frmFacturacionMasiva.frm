VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFacturacionMasiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación masiva"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarraCFD 
      Height          =   1125
      Left            =   960
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   10080
      Begin MSComctlLib.ProgressBar pgbBarraCFD 
         Height          =   495
         Left            =   45
         TabIndex        =   12
         Top             =   600
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarraCFD 
         BackColor       =   &H80000002&
         Caption         =   "Generando el Comprobante Fiscal Digital para la factura, por favor espere..."
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
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   9795
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   465
         Index           =   0
         Left            =   30
         Top             =   105
         Width           =   10020
      End
   End
   Begin VB.Frame fraPrincipal 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11895
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   360
         Top             =   5040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboNombreHospital 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   180
         Width           =   6255
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   5400
         TabIndex        =   9
         Top             =   4820
         Width           =   1230
         Begin VB.CommandButton cmdEnviar 
            Enabled         =   0   'False
            Height          =   550
            Left            =   620
            MaskColor       =   &H00EFEFEF&
            Picture         =   "frmFacturacionMasiva.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Enviar correo"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   555
         End
         Begin VB.CommandButton cmdSave 
            Height          =   550
            Left            =   60
            MaskColor       =   &H80000000&
            Picture         =   "frmFacturacionMasiva.frx":0D26
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Iniciar timbrado masivo"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   550
         End
      End
      Begin VB.CommandButton cmdCargarArchivo 
         Caption         =   "Cargar archivo Excel"
         Height          =   315
         Left            =   10080
         TabIndex        =   0
         ToolTipText     =   "Seleccionar archivo con la información de las facturas"
         Top             =   180
         Width           =   1815
      End
      Begin VB.Frame fraDatosEspecificos 
         Caption         =   "Detalle de la factura"
         Height          =   2055
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   11775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfConcepto 
            Height          =   1695
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   2990
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraDatosGenerales 
         Caption         =   "Facturas"
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   11775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFacturas 
            Height          =   1695
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   2990
            _Version        =   393216
            HighLight       =   2
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   220
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmFacturacionMasiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------------
' Programa para facturación masiva a clientes
' Fecha de desarrollo: Diciembre 02, 2021
'--------------------------------------------------------------------------------------------------------
  
    Dim vllngFormatoaUsar As Long               'Para saber que formato se va a utilizar
    Dim llngFormato As Long                     'Num. del formato de factura para el departamento
    Const cintTipoFormato = 9                   'Formato para factura directa en <TipoFormato> CC
    Dim intTipoEmisionComprobante As Integer    'Variable que compara el tipo de formato y folio a utilizar (0 = Error de formato y folios incompatibles, 1 = Físicos, 2 = Digitales)
    Dim intTipoCFDFactura As Integer            'Variable que regresa el tipo de CFD de la factura(0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
    Dim llngPersonaGraba As Long                'Num. de empleado que graba la factura
    Dim vlblnEsCredito As Boolean
    Dim vlblnPagoForma As Boolean               'Variable que indica si se utilizó la pantalla de formas de pago
    Dim aFormasPago() As FormasPago
    Dim vldblTipoCambio As Double
    Public cstrCantidad As String               'Para formatear a número
    Public cstrCantidad4Decimales As String     'Para formatear a número
    Dim llngNumReferencia As Long               'Nummero de referencia del cliente
    Dim lstrTipoCliente As String           'Tipo de cliente
    Dim lblnEntraCorte As Boolean           'Para saber si la factura entra o no en el corte
    Dim strFolio As String
    Dim llngNumPoliza As Long               'Num. de póliza
    Dim vlstrError As String
    Dim strSerie As String
    Dim llngNumCorte As Long                'Num. de corte en el que se está guardando
    Dim strAnoAprobacion As String              'Año de aprobación del folio
    Dim strNumeroAprobacion As String           'Número de aprobación del folio
    Dim vlblnMultiempresa As Boolean
    Dim arrTarifas() As typTarifaImpuesto
    Dim vgintnumemprelacionada As Integer
    Dim vlblnCuentaIngresoSaldada As Boolean        'Variable que indica si la cuenta del ingreso fue saldada con la cuenta del descuento
    Dim vlintBitSaldarCuentas As Long               'Variable que indica el valor del bit pvConceptoFacturacion.BitSaldarCuentas, que nos dice si la cuenta del ingreso se salda con la del descuento
    Dim apoliza() As TipoPoliza             'Para formar la poliza de la factura
    Dim lblnConsulta As Boolean             'Para saber si se está consultando una factura
    Dim vldblTotalIVACredito As Double
    Dim llngNumCtaCliente As Long           'Num. de cuenta contable del cliente
    Dim llngNumFormaCredito As Long         'Num. de forma de pago CREDITO para el departamento
    Dim dblProporcionIVA As Double
    Dim vldblComisionIvaBancaria As Double          'Cantidad que corresponde al iva de la comisión bancaria aplicada a cada forma de pago
    Dim tipoPago As DirectaMasiva
    Dim vldtmfechaServer As Date
    Public cintColCveConcepto As Integer
    Public cintColImporte As Integer
    Public cintColCantidad As Integer
    Public cintColDescuento As Integer
    Public cintColIVA As Integer
    Public cintColBitExento As Integer
    Public cintColCtaIngreso As Integer
    Public cintColCtaDescuento As Integer
    Public cintColDeptoConcepto As Integer
    Public blnNoFolios As Boolean
    Dim aPoliza2() As RegistroPoliza
    
    
    Dim oExcel As Object
    Dim oLibro As Object
    Dim oHoja As Object
    
Private Sub cmdCargarArchivo_Click()
    
    CommonDialog1.Filter = "Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*"
    CommonDialog1.DefaultExt = "xlsx"
    CommonDialog1.DialogTitle = "Seleccionar archivo"
    CommonDialog1.ShowOpen
    'True Then
    If CommonDialog1.FileName <> "" Then
        Set oExcel = CreateObject("Excel.Application")
        Set oLibro = oExcel.Workbooks.Open(FileName:=CommonDialog1.FileName) '"C:\Users\W7-siho\Desktop\Facturas.xlsx") 'CommonDialog1.FileName)
        CommonDialog1.FileName = ""
        Set oHoja = oLibro.Worksheets(1)
        cmdCargarArchivo.Enabled = False
        pLlenaGridFacturas

        
    Else

    End If
    
End Sub

Private Sub pLlenaGridFacturas()

On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim rsConcepto As ADODB.Recordset
    Dim rsISR As ADODB.Recordset
    Dim rsiva As ADODB.Recordset
    Dim rsFormasPago As ADODB.Recordset
    Dim rsBancoSAT As ADODB.Recordset
    Dim X As Integer
    Dim z As Integer
    Dim vlstrParametrosSP As String
    Dim vlstrsSQL As String
    X = 1 'Para recorrer la columna de M o D
    z = 0 'Para saber cuantas facturas van
    
    grdFacturas.Visible = False
    vsfConcepto.Visible = False
    
    If oHoja.cells(X, 1) = "" Then
        vlstrError = "El documento está vacío o tiene un formato incorrecto."
        GoTo NotificaError
    End If
    
    Do While oHoja.cells(X, 1) <> ""

        If oHoja.cells(X, 1) = "M" Then
            z = z + 1
            If z <> 1 Then
                grdFacturas.Rows = grdFacturas.Rows + 1
            End If
            grdFacturas.TextMatrix(z, 1) = oHoja.cells(X, 2)                        'Id del cliente

            vlstrParametrosSP = Trim(Str(oHoja.cells(X, 2))) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|1"
            vlstrError = "No se encontró información del cliente " & Trim(Str(oHoja.cells(X, 2))) & " o se encuentra inactivo. Celda " & X & "B"
            Set rs = frsEjecuta_SP(vlstrParametrosSP, "sp_CcSelDatosCliente")
            grdFacturas.TextMatrix(z, 2) = Trim(rs!NombreCliente)                   'Nombre cliente
            grdFacturas.TextMatrix(z, 3) = Trim(rs!RFCCliente)                      'RFC
            grdFacturas.TextMatrix(z, 33) = UCase(CStr(oHoja.cells(X, 16)))         'Observaciones
            grdFacturas.TextMatrix(z, 15) = Trim(rs!RazonSocial)                    'Razón social
            grdFacturas.TextMatrix(z, 16) = CStr(oHoja.cells(X, 3))                 'Motivo de la factura
            If Not (grdFacturas.TextMatrix(z, 16) = "0" Or grdFacturas.TextMatrix(z, 16) = "1" Or grdFacturas.TextMatrix(z, 16) = "2" Or grdFacturas.TextMatrix(z, 16) = "3") Then
                vlstrError = "El motivo de la factura no es válido. Celda " & X & "C"
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 17) = CStr(rs!RetServicios)                   'Retención por servicios
            If grdFacturas.TextMatrix(z, 16) = "3" And grdFacturas.TextMatrix(z, 17) = "0" Then
                vlstrError = "No se registró la retención de IVA del cliente " & Trim(rs!NombreCliente) & "."
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 19) = frsCveUsoCFDI("c_UsoCFDI", CStr(oHoja.cells(X, 4)))
            If grdFacturas.TextMatrix(z, 19) = "-2" Then
                vlstrError = "El uso del CFDI es incorrecto o no se capturó. Celda " & X & "D"
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 20) = CStr(oHoja.cells(X, 5))
            If grdFacturas.TextMatrix(z, 16) = "1" Then 'Honorarios
                Set rsISR = frsEjecuta_SP(grdFacturas.TextMatrix(z, 20) & "|1", "SP_CNSELTARIFAISR") 'Porcentaje
                If rsISR.RecordCount = 0 Then
                    vlstrError = "No se configuró la tasa de ISR o se encuentra inactiva. Celda " & X & "E"
                    GoTo NotificaError
                End If
                rsISR.Close
            End If
            
            grdFacturas.TextMatrix(z, 25) = CStr(oHoja.cells(X, 6))
            vlstrsSQL = "select * from PVFORMAPAGO where intformapago = " & grdFacturas.TextMatrix(z, 25) & " and smiDepartamento = " & Trim(Str(vgintNumeroDepartamento))
            Set rsFormasPago = frsRegresaRs(vlstrsSQL)
            If rsFormasPago.RecordCount = 0 Then
                vlstrError = "No se encontró información de la forma de pago " & grdFacturas.TextMatrix(z, 25) & ", pertenece a otro departamento o se encuentra inactiva. Celda " & X & "F"
                GoTo NotificaError
            End If
            
            If (CStr(oHoja.cells(X, 7)) = "1" Or CStr(oHoja.cells(X, 7)) = "2") Then
                grdFacturas.TextMatrix(z, 12) = IIf(oHoja.cells(X, 7) = 1, "PESOS", "DOLARES") 'Moneda
                If grdFacturas.TextMatrix(z, 12) = "DOLARES" Then
                    grdFacturas.TextMatrix(z, 18) = fdblTipoCambio(CDate(fdtmServerFecha), "O")
                    If Val(Format(grdFacturas.TextMatrix(z, 18), cstrCantidad)) = 0 Then
                        vlstrError = "Registre el tipo de cambio del día."
                        GoTo NotificaError
                    End If
                Else
                    grdFacturas.TextMatrix(z, 18) = "0"
                End If
            Else
                vlstrError = "El código de moneda es incorrecto o no se capturó. Celda " & X & "G"
                GoTo NotificaError
            End If
            
            grdFacturas.TextMatrix(z, 22) = CStr(oHoja.cells(X, 14))                       'Referencia de pago
            
            
            If (Len(CStr(grdFacturas.TextMatrix(z, 22))) > 3 And IsNumeric(grdFacturas.TextMatrix(z, 22))) Or grdFacturas.TextMatrix(z, 22) = "" Then

            Else
                vlstrError = "La referencia del pago tiene un formato incorrecto. Celda " & X & "N"
                GoTo NotificaError
            End If
            
            grdFacturas.TextMatrix(z, 23) = IIf(CStr(oHoja.cells(X, 8)) = "", Trim(rs!RFCCliente), CStr(oHoja.cells(X, 8)))
            
            grdFacturas.TextMatrix(z, 24) = IIf(CStr(oHoja.cells(X, 11)) <> "", CStr(oHoja.cells(X, 11)), CStr(vldtmfechaServer))
            
            If CDate(grdFacturas.TextMatrix(z, 24)) > vldtmfechaServer Then
                vlstrError = "La fecha de pago de la factura " & z & " debe ser menor o igual a la del sistema"
                GoTo NotificaError
            End If
            
            If rsFormasPago!VCHDESCRIPCIONCFD = "" Then
                vlstrError = "No se configuró el Método de pago del SAT para CFDI, para la forma de pago " & grdFacturas.TextMatrix(z, 25)
                GoTo NotificaError
            End If
            
            If rsFormasPago!INTIDFORMAPAGOSAT = "" Then
                vlstrError = "No se configuró el Método de pago del SAT para contabilidad electrónica, para la forma de pago " & grdFacturas.TextMatrix(z, 25)
                GoTo NotificaError
            End If
            
            grdFacturas.TextMatrix(z, 32) = rsFormasPago!chrTipo
          
            grdFacturas.TextMatrix(z, 27) = IIf(CStr(oHoja.cells(X, 15)) <> "", CStr(oHoja.cells(X, 15)), IIf(IsNull(rs!CORREO), "", rs!CORREO))
            If grdFacturas.TextMatrix(z, 27) = "" Then
                vlstrError = "No se encontró un correo electrónico para el cliente " & Trim(rs!NombreCliente) & ". Celda " & X & "O"
                GoTo NotificaError
            End If
            
            grdFacturas.TextMatrix(z, 34) = IIf(IsNull(rs!REGIMENFISCAL), "", rs!REGIMENFISCAL)
            If Trim(grdFacturas.TextMatrix(z, 34)) = "" Then
                vlstrError = "No se cuenta con el régimen fiscal del cliente " & Trim(rs!NombreCliente) & "."
                GoTo NotificaError
            End If
            
            grdFacturas.TextMatrix(z, 35) = IIf(IsNull(rs!Codigo), "", rs!Codigo)
            If Trim(grdFacturas.TextMatrix(z, 35)) = "" Then
                vlstrError = "No se cuenta con el código postal del cliente " & Trim(rs!NombreCliente) & "."
                GoTo NotificaError
            End If
            
            If rsFormasPago!chrTipo = "T" Or rsFormasPago!chrTipo = "B" Or rsFormasPago!chrTipo = "H" Then
                grdFacturas.TextMatrix(z, 28) = IIf(CStr(oHoja.cells(X, 9)) <> "", CStr(oHoja.cells(X, 9)), "0")
                vlstrsSQL = "SELECT * FROM CPBANCOSAT WHERE bitactivo = 1 AND chrclave = '" & grdFacturas.TextMatrix(z, 28) & "'"
                Set rsBancoSAT = frsRegresaRs(vlstrsSQL)
                If rsBancoSAT.RecordCount = 0 Then
                    vlstrError = "No se encontró información del banco con la clave " & grdFacturas.TextMatrix(z, 28) & " o se encuentra inactivo. Celda " & X & "I"
                    GoTo NotificaError
                End If
                rsBancoSAT.Close
                
                grdFacturas.TextMatrix(z, 29) = IIf(CStr(oHoja.cells(X, 12)) <> "", CStr(oHoja.cells(X, 12)), "0")
                vlstrsSQL = "SELECT * FROM CPBANCO WHERE bitestatus = 1 AND TNYNUMEROBANCO = '" & grdFacturas.TextMatrix(z, 29) & "'"
                Set rsBancoSAT = frsRegresaRs(vlstrsSQL)
                If rsBancoSAT.RecordCount = 0 Then
                    vlstrError = "No se encontró información de la cuenta bancaria con clave " & grdFacturas.TextMatrix(z, 29) & " o se encuentra inactiva. Celda " & X & "L"
                    GoTo NotificaError
                End If
                rsBancoSAT.Close
                
                grdFacturas.TextMatrix(z, 31) = IIf(CStr(oHoja.cells(X, 10)) <> "", CStr(oHoja.cells(X, 10)), "0")
                If fblnCuentaOrdenanteValida(True, grdFacturas.TextMatrix(z, 31), rsFormasPago!VCHDESCRIPCIONCFD) = False Then
                    vlstrError = "El tamaño de la cuenta bancaria emisora del pago no tiene la longitud esperada. Celda " & X & "J"
                    GoTo NotificaError
                End If
                
                grdFacturas.TextMatrix(z, 30) = IIf(CStr(oHoja.cells(X, 13)) <> "", CStr(oHoja.cells(X, 13)), "")
                If grdFacturas.TextMatrix(z, 30) <> "" Then
                    vlstrParametrosSP = rsFormasPago!intFormaPago & "|" & grdFacturas.TextMatrix(z, 30) & "|" & vgintClaveEmpresaContable
                    Set rsBancoSAT = frsEjecuta_SP(vlstrParametrosSP, "sp_PvSelComisionCargoBancario") 'Se realiza la consulta
                    If rsBancoSAT.RecordCount = 0 Then
                        vlstrError = "No se encontró información del tipo de cargo bancario con clave " & grdFacturas.TextMatrix(z, 30) & ", no se ha configurado correctamente o se encuentra inactivo."
                        GoTo NotificaError
                    End If
                End If
                
                If Len(CStr(grdFacturas.TextMatrix(z, 22))) > 3 And IsNumeric(grdFacturas.TextMatrix(z, 22)) Then

                Else
                    vlstrError = "No se registró la referencia del pago o tiene un formato incorrecto. Celda " & X & "N"
                    GoTo NotificaError
                End If
                
            End If
        Else
            If CStr(oHoja.cells(X, 3)) = "" Or CStr(oHoja.cells(X, 3)) = "0" Then
                vlstrError = "No se capturó la cantidad o es igual a cero. Celda " & X & "C"
                GoTo NotificaError
            End If
            If CStr(oHoja.cells(X, 4)) = "" Or CStr(oHoja.cells(X, 4)) = "0" Then
                vlstrError = "No se capturó el precio unitario o es igual a cero. Celda " & X & "D"
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 4) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 4) = "", 0, grdFacturas.TextMatrix(z, 4))) + (CDbl(oHoja.cells(X, 3)) * CDbl(oHoja.cells(X, 4))), 4)   'Importe
            If grdFacturas.TextMatrix(z, 4) = 0 Then
                vlstrError = "El importe de un concepto es 0 en la factura " & z & "."
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 5) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 5) = "", 0, grdFacturas.TextMatrix(z, 5))) + CDbl(oHoja.cells(X, 5)), 4)   'Descuento
            If CDbl(grdFacturas.TextMatrix(z, 5)) > CDbl(grdFacturas.TextMatrix(z, 4)) Then
                vlstrError = "El descuento no puede ser mayor al importe. Celda " & X & "E"
                GoTo NotificaError
            End If
            grdFacturas.TextMatrix(z, 6) = FormatCurrency(CDbl(grdFacturas.TextMatrix(z, 4)) - CDbl(grdFacturas.TextMatrix(z, 5)), 4)   'Subtotal
            vlstrParametrosSP = CStr(oHoja.cells(X, 2)) & "|-1|-1|" & vgintClaveEmpresaContable
            Set rsConcepto = frsEjecuta_SP(vlstrParametrosSP, "sp_PvSelConceptoFacturacion")
            If rsConcepto.RecordCount > 0 Then
                grdFacturas.TextMatrix(z, 7) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 7) = "", 0, grdFacturas.TextMatrix(z, 7))) + ((CDbl(oHoja.cells(X, 3)) * CDbl(oHoja.cells(X, 4)) - CDbl(oHoja.cells(X, 5))) * ((rsConcepto!smyIVA) / 100)), 4) 'IVA
                grdFacturas.TextMatrix(z, 8) = FormatCurrency(CDbl(grdFacturas.TextMatrix(z, 6)) + CDbl(grdFacturas.TextMatrix(z, 7)), 4)  'Total
            Else
                vlstrError = "No se encontró información del concepto " & Trim(Str(oHoja.cells(X, 2))) & " o se encuentra inactivo. Celda " & X & "B"
                GoTo NotificaError
            End If
            If fblnValidaSAT(CLng(oHoja.cells(X, 2)), rsConcepto!chrdescripcion) = False Then
                GoTo NotificaError
            End If
            If grdFacturas.TextMatrix(z, 16) = "0" Then 'Otros
                grdFacturas.TextMatrix(z, 9) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 9) = "", 0, grdFacturas.TextMatrix(z, 9))) + CDbl(0), 4)   'CStr(oHoja.Cells(x, 5) * 100) & "%"      'Retención de ISR
                grdFacturas.TextMatrix(z, 10) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 10) = "", 0, grdFacturas.TextMatrix(z, 10))) + CDbl(0), 4) 'CStr(oHoja.Cells(x, 6) * 100) & "%"     'Retención de IVA
            ElseIf grdFacturas.TextMatrix(z, 16) = "1" Then 'Honorarios
                Set rsISR = frsEjecuta_SP(grdFacturas.TextMatrix(z, 20) & "|1", "SP_CNSELTARIFAISR") 'Porcentaje
                    If rsISR.RecordCount > 0 Then
                        grdFacturas.TextMatrix(z, 9) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 9) = "", 0, grdFacturas.TextMatrix(z, 9))) + (CDbl(grdFacturas.TextMatrix(z, 6)) * CDbl(rsISR!Porcentaje / 100)), 4)   'CStr(oHoja.Cells(x, 5) * 100) & "%"      'Retención de ISR
                        grdFacturas.TextMatrix(z, 10) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 10) = "", 0, grdFacturas.TextMatrix(z, 10))) + CDbl(0), 4) 'CStr(oHoja.Cells(x, 6) * 100) & "%"     'Retención de IVA
                        grdFacturas.TextMatrix(z, 21) = pBuscaIndiceTasaRetencion(CLng(grdFacturas.TextMatrix(z, 20)))
                    End If
            ElseIf grdFacturas.TextMatrix(z, 16) = "2" Then 'Arrendamiento
                grdFacturas.TextMatrix(z, 9) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 9) = "", 0, grdFacturas.TextMatrix(z, 9))) + CDbl(0), 4)   'CStr(oHoja.Cells(x, 5) * 100) & "%"      'Retención de ISR
                grdFacturas.TextMatrix(z, 10) = FormatCurrency(CDbl(CDbl(grdFacturas.TextMatrix(z, 7)) * CDbl(2 / 3)), 4) 'CStr(oHoja.Cells(x, 6) * 100) & "%"     'Retención de IVA
            ElseIf grdFacturas.TextMatrix(z, 16) = "3" Then 'Servicios
                grdFacturas.TextMatrix(z, 9) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 9) = "", 0, grdFacturas.TextMatrix(z, 9))) + CDbl(0), 4)   'CStr(oHoja.Cells(x, 5) * 100) & "%"      'Retención de ISR
                grdFacturas.TextMatrix(z, 10) = FormatCurrency(CDbl(IIf(grdFacturas.TextMatrix(z, 10) = "", 0, grdFacturas.TextMatrix(z, 10))) + (CDbl(grdFacturas.TextMatrix(z, 6)) * CDbl(grdFacturas.TextMatrix(z, 17))), 4) 'CStr(oHoja.Cells(x, 6) * 100) & "%"     'Retención de IVA
            End If
            grdFacturas.TextMatrix(z, 11) = FormatCurrency(CDbl(grdFacturas.TextMatrix(z, 8)) - (CDbl(grdFacturas.TextMatrix(z, 9) + CDbl(grdFacturas.TextMatrix(z, 10)))), 4) 'Total a cobrar
            
        End If
        X = X + 1
    Loop
    
    With grdFacturas
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    grdFacturas.Redraw = True
    grdFacturas.Visible = True
    Call grdFacturas_Click
    
Exit Sub
NotificaError:
    'Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
    MsgBox vlstrError, vbExclamation + vbOKOnly, "Mensaje"
    pLimpia
End Sub

Private Sub pLlenaGridFacturasDetalle()

    Dim vlblnEncontrado As Boolean
    Dim vlintFacturaAEncontrar As Integer
    Dim vlintXPosicion As Integer
    Dim vlintZ As Integer
    Dim rsConcepto As ADODB.Recordset
    Dim vlstrParametrosSP As String
    vlblnEncontrado = False
    vlintFacturaAEncontrar = grdFacturas.Row
    vlintFacturaEncontrada = 0
    vlintXPosicion = 1
    vlintZ = 0
    
    vsfConcepto.Visible = False
    pConfiguraGridFacturasDetalle
    Do While vlblnEncontrado = False
        If grdFacturas.TextMatrix(1, 1) = "" Then
            Exit Sub
        End If
        If oHoja.cells(vlintXPosicion, 1) = "M" Then 'Si la primer casilla es una M (Maestro) se revisa
            vlintFacturaEncontrada = vlintFacturaEncontrada + 1 'Se aumentan las facturas encontradas
            If vlintFacturaEncontrada = vlintFacturaAEncontrar Then 'Se busca si el numero de factura corresponde con el numero de factura encontrada en el excel
                vlblnEncontrado = True 'Se evita más entradas al loop, la factura se encontro!!
                vlintXPosicion = vlintXPosicion + 1 'Significa que el siguiente registro a la factura es un D (Detalle)
                
                Do While oHoja.cells(vlintXPosicion, 1) = "D" 'Mientras siga siendo D (Detalle) insertamos la info en la tabla
                    
                    vlintZ = vlintZ + 1                                         '
                    If vlintZ <> 1 Then                                         ' Mecanismo para agregar registros en blanco en la tabla
                        vsfConcepto.Rows = vsfConcepto.Rows + 1   ' para rellenarlos con la información del concepto
                    End If                                                      '
                    
                    vlstrParametrosSP = CStr(oHoja.cells(vlintXPosicion, 2)) & "|-1|-1|" & vgintClaveEmpresaContable 'Se generan los parámetros para consultar la info de los conceptos de factura
                    Set rsConcepto = frsEjecuta_SP(vlstrParametrosSP, "sp_PvSelConceptoFacturacion") 'Se realiza la consulta
                    
                    vsfConcepto.TextMatrix(vlintZ, 11) = CStr(oHoja.cells(vlintXPosicion, 2))
                    vsfConcepto.TextMatrix(vlintZ, 1) = rsConcepto!chrdescripcion    'Descripcion del concepto
                    vsfConcepto.TextMatrix(vlintZ, 2) = Format(CDbl(oHoja.cells(vlintXPosicion, 3)), "#.00") 'Cantidad
                    vsfConcepto.TextMatrix(vlintZ, 3) = FormatCurrency(CDbl(oHoja.cells(vlintXPosicion, 4)), 4) 'Precio unitario
                    vsfConcepto.TextMatrix(vlintZ, 4) = FormatCurrency(CDbl(oHoja.cells(vlintXPosicion, 4)) * CDbl(oHoja.cells(vlintXPosicion, 3)), 4) 'Importe
                    vsfConcepto.TextMatrix(vlintZ, 5) = FormatCurrency(CDbl(oHoja.cells(vlintXPosicion, 5)), 4) 'Descuento
                    vsfConcepto.TextMatrix(vlintZ, 6) = FormatCurrency(((CDbl(oHoja.cells(vlintXPosicion, 3)) * CDbl(oHoja.cells(vlintXPosicion, 4)) - CDbl(oHoja.cells(vlintXPosicion, 5))) * ((rsConcepto!smyIVA) / 100)), 4) 'IVA
                    vsfConcepto.TextMatrix(vlintZ, 7) = rsConcepto!bitExentoIva
                    vsfConcepto.TextMatrix(vlintZ, 8) = rsConcepto!INTCUENTACONTABLE
                    vsfConcepto.TextMatrix(vlintZ, 9) = rsConcepto!intCuentaDescuento
                    vsfConcepto.TextMatrix(vlintZ, 10) = rsConcepto!SMIDEPARTAMENTO
                    vlintXPosicion = vlintXPosicion + 1
                    
                Loop
                
            Else
                vlintXPosicion = vlintXPosicion + 1 'No es la factura buscada, seguimos girando el loop
            End If
        Else
            vlintXPosicion = vlintXPosicion + 1 'Si la primer casilla No es una M (Maestro) no se revisa
        End If
        
    Loop
    vsfConcepto.Visible = True
End Sub

Private Sub cmdEnviar_Click()
    Dim vlintFacturas As Integer 'Contador de facturas a timbrar
    If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
    '- Revisar que el parámetro de envío de CFD esté activado -'
        If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
            For vlintFacturas = 1 To grdFacturas.Rows - 1
                With grdFacturas
                    .Row = vlintFacturas
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                grdFacturas.Redraw = True
                If grdFacturas.TextMatrix(vlintFacturas, 26) <> "" Then
                    pEnviarMasivo "FA", CLng(grdFacturas.TextMatrix(vlintFacturas, 26)), CLng(vgintClaveEmpresaContable), grdFacturas.TextMatrix(vlintFacturas, 3), llngPersonaGraba, grdFacturas.TextMatrix(vlintFacturas, 27), Me
                End If
            Next vlintFacturas
            MsgBox "Proceso de envío efectuado, verificar el estado del envío.", vbOKOnly + vbInformation, "Mensaje"
            cmdEnviar.Enabled = False
        Else
            MsgBox "No está activo el envío de correos para esta empresa contable.", vbOKOnly + vbInformation, "Mensaje"
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdSave_Click()

    If grdFacturas.TextMatrix(1, 1) = "" Then
        MsgBox "No se ha cargado información", vbExclamation + vbOKOnly, "Mensaje"
        Exit Sub
    End If

    Dim vlintFacturas As Integer 'Contador de facturas a timbrar
    Dim vlstrParametrosSP As String
    Dim rs As ADODB.Recordset
    Dim vlintFacturasFinal As Integer 'Contador de facturas en blanco para ponerlas no timbradas
    For vlintFacturas = 1 To grdFacturas.Rows - 1
        If blnNoFolios = True Then
            
            Exit For
        End If
        With grdFacturas
            .Row = vlintFacturas
            .Col = 0
            .ColSel = .Cols - 1
        End With
        grdFacturas.Redraw = True
        Call grdFacturas_Click
        vlstrParametrosSP = grdFacturas.TextMatrix(vlintFacturas, 1) & "|0|*|*|" & CStr(vgintClaveEmpresaContable) & "|1"
        Set rs = frsEjecuta_SP(vlstrParametrosSP, "sp_CcSelDatosCliente")
        llngNumReferencia = rs!INTNUMREFERENCIA
        lstrTipoCliente = rs!chrTipoCliente
        llngNumCtaCliente = rs!INTNUMCUENTACONTABLE
        Dim arrDatosFisc() As DatosFiscales
        ReDim arrDatosFisc(0)
        arrDatosFisc(0).strDomicilio = IIf(IsNull(rs!CHRCALLE), " ", rs!CHRCALLE)
        arrDatosFisc(0).strNumExterior = IIf(IsNull(rs!vchNumeroExterior), " ", rs!vchNumeroExterior)
        arrDatosFisc(0).strNumInterior = IIf(IsNull(rs!vchNumeroInterior), " ", rs!vchNumeroInterior)
        arrDatosFisc(0).strTelefono = IIf(IsNull(rs!Telefono), " ", rs!Telefono)
        arrDatosFisc(0).lstrCalleNumero = IIf(IsNull(rs!callenumero), "", rs!callenumero)
        arrDatosFisc(0).lstrColonia = IIf(IsNull(rs!Colonia), "", rs!Colonia)
        arrDatosFisc(0).lstrCiudad = IIf(IsNull(rs!ciudadcliente), "", rs!ciudadcliente)
        arrDatosFisc(0).lstrEstado = IIf(IsNull(rs!Estado), "", rs!Estado)
        arrDatosFisc(0).lstrCodigo = IIf(IsNull(rs!Codigo), "", rs!Codigo)
        arrDatosFisc(0).llngCveCiudad = IIf(IsNull(rs!IdCiudad), 0, rs!IdCiudad)
        vlblnEsCredito = False
        vlblnPagoForma = False
        lblnEntraCorte = False
        vlblnMultiempresa = False
        vgstrParametrosSP = vgintClaveEmpresaContable & "|" & rs!INTNUMREFERENCIA & "|" & vgintNumeroDepartamento
        Set rsMultiemp = frsEjecuta_SP(vgstrParametrosSP, "sp_CCSelEmpresaProveedor")
        If rsMultiemp.RecordCount <> 0 Then
            Do While Not rsMultiemp.EOF
                If rsMultiemp!idempresacliente = rs!INTNUMREFERENCIA Then
                    If rsMultiemp!idproveedor <> 0 Then
                        vlblnMultiempresa = True
                        vgintnumemprelacionada = rsMultiemp!empresa
                    Else
                       vlblnMultiempresa = False
                    End If
                End If
                rsMultiemp.MoveNext
            Loop
        End If
        llngNumCorte = 0
        pCargaFolio 0
        lblnConsulta = False
        Dim fecha As String
        fecha = fdtmServerFecha
        tipoPago.intDirectaMasiva = 1
        tipoPago.strFormaPago = grdFacturas.TextMatrix(vlintFacturas, 25)
        tipoPago.strFechaPago = grdFacturas.TextMatrix(vlintFacturas, 24)
        tipoPago.strReferenciaPago = grdFacturas.TextMatrix(vlintFacturas, 22)
        tipoPago.vlStrRFCPago = grdFacturas.TextMatrix(vlintFacturas, 23)
        tipoPago.vlStrClaveBancoSAT = grdFacturas.TextMatrix(vlintFacturas, 28)
        tipoPago.vlstrTipoPago = grdFacturas.TextMatrix(vlintFacturas, 32)
        tipoPago.vlstrCuenta = grdFacturas.TextMatrix(vlintFacturas, 31)
        tipoPago.vlstrClaveCuentaBancaria = grdFacturas.TextMatrix(vlintFacturas, 29)
        tipoPago.vlStrTipoCargoBancario = grdFacturas.TextMatrix(vlintFacturas, 30)
        
        vlstrRegimenFiscal = Trim(grdFacturas.TextMatrix(vlintFacturas, 34))
        
        vllngFormatoaUsar = llngFormato
        pGeneraFacturaDirecta Me, vllngFormatoaUsar, intTipoEmisionComprobante, intTipoCFDFactura, CInt(0), Trim(rs!RFCCliente), Trim(rs!RazonSocial), llngPersonaGraba, _
        vlblnEsCredito, vlblnPagoForma, CInt(0), CStr(grdFacturas.TextMatrix(vlintFacturas, 11)), aFormasPago(), CInt(IIf(grdFacturas.TextMatrix(vlintFacturas, 12) = "PESOS", 1, 0)), CDbl(grdFacturas.TextMatrix(vlintFacturas, 18)), llngNumReferencia, lstrTipoCliente, lblnEntraCorte, _
        fecha, strFolio, llngNumPoliza, CInt(1), CInt(grdFacturas.TextMatrix(vlintFacturas, 19)), CInt(grdFacturas.TextMatrix(vlintFacturas, 16)), CInt(IIf(grdFacturas.TextMatrix(vlintFacturas, 16) = "2" Or grdFacturas.TextMatrix(vlintFacturas, 16) = "3", 1, 0)), grdFacturas.TextMatrix(vlintFacturas, 7), grdFacturas.TextMatrix(vlintFacturas, 5), grdFacturas.TextMatrix(vlintFacturas, 1), _
        grdFacturas.TextMatrix(vlintFacturas, 18), cstrCantidad4Decimales, strSerie, grdFacturas.TextMatrix(vlintFacturas, 9), grdFacturas.TextMatrix(vlintFacturas, 10), grdFacturas.TextMatrix(vlintFacturas, 33), CInt(IIf(grdFacturas.TextMatrix(vlintFacturas, 20) = "", "0", grdFacturas.TextMatrix(vlintFacturas, 20))), llngNumCorte, CInt(0), CInt(0), strAnoAprobacion, strNumeroAprobacion, cstrCantidad, cintTipoFormato, vlblnMultiempresa, CInt(IIf(grdFacturas.TextMatrix(vlintFacturas, 16) = "1", 1, 0)), CInt(IIf(grdFacturas.TextMatrix(vlintFacturas, 21) = "", "0", grdFacturas.TextMatrix(vlintFacturas, 21))), arrTarifas(), _
        vgintnumemprelacionada, vlblnCuentaIngresoSaldada, vlintBitSaldarCuentas, apoliza(), arrDatosFisc, lblnConsulta, vldblTotalIVACredito, llngNumCtaCliente, llngNumFormaCredito, grdFacturas.TextMatrix(vlintFacturas, 6), dblProporcionIVA, vldblComisionIvaBancaria, tipoPago, aPoliza2()
        
        If llngPersonaGraba = 0 Then Exit Sub
        
    Next vlintFacturas
    llngPersonaGraba = 0

    For vlintFacturasFinal = 1 To grdFacturas.Rows - 1
        With grdFacturas
            .Row = vlintFacturasFinal
            .Col = 0
            .ColSel = .Cols - 1
        End With
        grdFacturas.Redraw = True
        If grdFacturas.TextMatrix(grdFacturas.RowSel, 13) = "" Then
            grdFacturas.TextMatrix(grdFacturas.RowSel, 13) = "NO TIMBRADA"
        End If
    Next vlintFacturasFinal

    cmdEnviar.Enabled = True
    cmdSave.Enabled = False
End Sub

Private Sub pCargaTasasRetencionISR()
    ReDim arrTarifas(0)
    Dim rs As ADODB.Recordset
    Dim intcontador As Integer
    
    vgstrParametrosSP = "-1|1"
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_CNSELTARIFAISR")
    If rs.RecordCount <> 0 Then
    
        intcontador = 0
        Do While Not rs.EOF
            ReDim Preserve arrTarifas(intcontador)
    
'            cboTarifa.AddItem rs!Descripcion
'            cboTarifa.ItemData(cboTarifa.newIndex) = rs!IdTarifa
'
            arrTarifas(intcontador).lngId = rs!IdTarifa
            arrTarifas(intcontador).dblPorcentaje = rs!Porcentaje
        
            intcontador = intcontador + 1
        
            rs.MoveNext
        Loop
        
    End If
End Sub

Private Function pBuscaIndiceTasaRetencion(lngIdABuscar As Long) As Integer

    Dim intcontador As Integer
    intcontador = 0
    Dim intPosicion As Integer
    For intcontador = 0 To UBound(arrTarifas)
        
        If arrTarifas(intcontador).lngId = lngIdABuscar Then
            intPosicion = intcontador
        End If
        
    Next intcontador
    
    pBuscaIndiceTasaRetencion = intPosicion

End Function

Private Sub Form_Activate()

    Dim intMensaje As Integer
    intMensaje = CInt(flngCorteValido(vgintNumeroDepartamento, vglngNumeroEmpleado, "P"))

    If intMensaje <> 0 Then
        'Cierre el corte actual antes de registrar este documento.
        'No existe un corte abierto
        MsgBox SIHOMsg(Str(intMensaje)), vbOKOnly + vbInformation, "Mensaje"
        Unload Me
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Then
        If grdFacturas.TextMatrix(1, 1) <> "" Then
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pLimpia
            End If
        Else
            Unload Me
        End If
    End If

End Sub

Private Sub Form_Load()
    Dim vlstrsql As String
    Dim rs As ADODB.Recordset
    Me.Icon = frmMenuPrincipal.Icon
    blnNoFolios = False
    
    'Configuración del nombre del hospital
    vlstrsql = "SELECT VCHNOMBRE FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
    cboNombreHospital.Text = rs!vchNombre
    rs.Close
    vldtmfechaServer = fdtmServerFecha
    
    cstrCantidad = "#############.00"
    cstrCantidad4Decimales = "#############.0000"
    cintColCantidad = 2
    cintColImporte = 4
    cintColDescuento = 5
    cintColIVA = 6
    cintColBitExento = 7
    cintColCtaIngreso = 8
    cintColCtaDescuento = 9
    cintColDeptoConcepto = 10
    cintColCveConcepto = 11
    
    pLimpia
    
    llngFormato = flngFormatoDepto(vgintNumeroDepartamento, cintTipoFormato, "*")
    pCargaTasasRetencionISR
    
End Sub

Private Sub pLimpia()

    pConfiguraGridFacturas
    pConfiguraGridFacturasDetalle
    vsfConcepto.Visible = True
    cmdCargarArchivo.Enabled = True
    cmdSave.Enabled = True
    cmdEnviar.Enabled = False
    blnNoFolios = False

End Sub

Private Sub pConfiguraGridFacturas()
    
    With grdFacturas
        'Inicializada...
        .Rows = 2
        .Cols = 36
        .Clear
        
        'Configurada
        .FormatString = "|Cliente|Nombre del cliente|RFC|Importe|Descuento|Subtotal|IVA|Total|Retención de ISR|Retención de IVA|Total a cobrar|Moneda|Estado|Folio|||||||||||||||||||Observaciones"
        .ColWidth(0) = 100
        .ColWidth(1) = 900  'Cliente
        .ColWidth(2) = 3000 'Nombre
        .ColWidth(3) = 1400 'RFC
        .ColWidth(4) = 1200 'Importe
        .ColWidth(5) = 1200 'Descuento
        .ColWidth(6) = 1200 'Subtotal
        .ColWidth(7) = 1200 'IVA
        .ColWidth(8) = 1200 'Total
        .ColWidth(9) = 1400 'Retención de ISR
        .ColWidth(10) = 1400 'Retención de IVA
        .ColWidth(11) = 1200 'Total a cobrar
        .ColWidth(12) = 1100 'Moneda
        .ColWidth(13) = 2500 'Estado
        .ColWidth(14) = 1200 'Folio
        .ColWidth(15) = 0 'Razon social
        .ColWidth(16) = 0 'Motivo de factura
        .ColWidth(17) = 0 'Retencion servicios para cuando la factura es por motivo de servicios
        .ColWidth(18) = 0 'Tipo de cambio
        .ColWidth(19) = 0 'Clave interna Uso CFDI
        .ColWidth(20) = 0 'Clave de la retención de ISR
        .ColWidth(21) = 0 'List Index de la tasa de ISR
        .ColWidth(22) = 0 'Referencia de pago
        .ColWidth(23) = 0 'RFC relacionado al pago
        .ColWidth(24) = 0 'Fecha relacionada al pago
        .ColWidth(25) = 0 'Forma de pago relacionada al pago
        .ColWidth(26) = 0 'lngidfactura Luego del timbrado
        .ColWidth(27) = 0 'Correo electrónico
        .ColWidth(28) = 0 'Clave del Banco de pago
        .ColWidth(29) = 0 'Clave de la cuenta bancaria de pago
        .ColWidth(30) = 0 'Clave del tipo de cargo del pago
        .ColWidth(31) = 0 'Tarjeta/Cheque/Transferencia
        .ColWidth(32) = 0 'Tipo de pago
        .ColWidth(33) = 4500 'Observaciones
        .ColWidth(34) = 0 'Regimen fiscal
        .ColWidth(35) = 0 'Regimen fiscal
        
        .ColAlignment(1) = flexAlignRightBottom
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignLeftBottom
        .ColAlignment(4) = flexAlignRightBottom
        .ColAlignment(5) = flexAlignRightBottom
        .ColAlignment(6) = flexAlignRightBottom
        .ColAlignment(7) = flexAlignRightBottom
        .ColAlignment(8) = flexAlignRightBottom
        .ColAlignment(9) = flexAlignRightBottom
        .ColAlignment(10) = flexAlignRightBottom
        .ColAlignment(11) = flexAlignRightBottom
        .ColAlignment(12) = flexAlignLeftBottom
        .ColAlignment(13) = flexAlignLeftBottom
        .ColAlignment(14) = flexAlignLeftCenter
        .ColAlignment(15) = flexAlignLeftCenter
        .ColAlignment(16) = flexAlignLeftCenter
        .ColAlignment(17) = flexAlignLeftCenter
        .ColAlignment(18) = flexAlignLeftCenter
        .ColAlignment(19) = flexAlignLeftCenter
        .ColAlignment(20) = flexAlignLeftCenter
        .ColAlignment(21) = flexAlignLeftCenter
        .ColAlignment(22) = flexAlignLeftCenter
        .ColAlignment(23) = flexAlignLeftCenter
        .ColAlignment(24) = flexAlignLeftCenter
        .ColAlignment(25) = flexAlignLeftCenter
        .ColAlignment(33) = flexAlignLeftCenter
        
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
        .ColAlignmentFixed(12) = flexAlignCenterCenter
        .ColAlignmentFixed(13) = flexAlignCenterCenter
        .ColAlignmentFixed(14) = flexAlignCenterCenter
        .ColAlignmentFixed(15) = flexAlignCenterCenter
        .ColAlignmentFixed(16) = flexAlignCenterCenter
        .ColAlignmentFixed(17) = flexAlignCenterCenter
        .ColAlignmentFixed(18) = flexAlignCenterCenter
        .ColAlignmentFixed(19) = flexAlignCenterCenter
        .ColAlignmentFixed(20) = flexAlignCenterCenter
        .ColAlignmentFixed(21) = flexAlignCenterCenter
        .ColAlignmentFixed(22) = flexAlignCenterCenter
        .ColAlignmentFixed(23) = flexAlignCenterCenter
        .ColAlignmentFixed(24) = flexAlignCenterCenter
        .ColAlignmentFixed(25) = flexAlignCenterCenter
        .ColAlignmentFixed(33) = flexAlignCenterCenter
        
        .Redraw = True
        .Visible = True
    End With
    
End Sub

Private Sub pConfiguraGridFacturasDetalle()

    vsfConcepto.Clear
    With vsfConcepto
        'Inicializada...
        .Rows = 2
        .Cols = 12
        .Clear
        
        'Configurada
        .FormatString = "|Concepto|Cantidad|Precio unitario|Importe|Descuento|IVA||||"
        .ColWidth(0) = 100
        .ColWidth(1) = 3000 'Concepto
        .ColWidth(cintColCantidad) = 1000 'Cantidad
        .ColWidth(3) = 1400 'Precio unitario
        .ColWidth(cintColImporte) = 1200 'Importe
        .ColWidth(cintColDescuento) = 1200 'Descuento
        .ColWidth(cintColIVA) = 1200 'IVA
        .ColWidth(cintColBitExento) = 0 'Bit exento
        .ColWidth(cintColCtaIngreso) = 0 'Cuenta Ingreso
        .ColWidth(cintColCtaDescuento) = 0 'Cuenta Descuento
        .ColWidth(cintColDeptoConcepto) = 0 'Depto Concepto
        .ColWidth(cintColCveConcepto) = 0 'Cve Concepto
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(cintColCantidad) = flexAlignRightBottom
        .ColAlignment(3) = flexAlignRightBottom
        .ColAlignment(cintColImporte) = flexAlignRightBottom
        .ColAlignment(cintColDescuento) = flexAlignRightBottom
        .ColAlignment(cintColIVA) = flexAlignRightBottom
        .ColAlignment(cintColBitExento) = flexAlignRightBottom
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColCantidad) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColImporte) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColDescuento) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColIVA) = flexAlignCenterCenter
        .ColAlignmentFixed(cintColBitExento) = flexAlignCenterCenter
        
        .Redraw = True
        '.Visible = True
    End With
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not oLibro Is Nothing Then
        oLibro.Close
        
        oExcel.Quit
        Set oExcel = Nothing
        Set oLibro = Nothing
        Set oHoja = Nothing
        cmdCargarArchivo.Enabled = True
    End If
End Sub

Private Sub grdFacturas_Click()
    If grdFacturas.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    
    pLlenaGridFacturasDetalle

End Sub

Public Sub pCargaFolio(intAumenta As Integer)
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
    strSerie = Trim(strSerie)
    If vllngFoliosFaltantes > 0 And intAumenta = 1 Then
        MsgBox "¡Faltan " & Trim(Str(vllngFoliosFaltantes)) + " facturas y será necesario aumentar folios!", vbOKOnly + vbInformation, "Mensaje"
    End If
    strFolio = Trim(strSerie) + Trim(strFolio)

    'Habilitar el chkBitExtranjero si el folio es de tipo digital
    If Trim(strNumeroAprobacion) <> "" And Trim(strAnoAprobacion) <> "" Then
        'chkBitExtranjero.Enabled = True
    End If
End Sub

Public Function fblnValidaSAT(intConcepto As Long, strNombreConcepto) As Boolean
    Dim intRow As Integer
    If vgstrVersionCFDI <> "3.2" Then
        If intConcepto > 0 Then
            If flngCatalogoSATIdByNombreTipo("c_ClaveProdServ", intConcepto, "CF", 1) = 0 Then
                vlstrError = "No está definida la clave del SAT para el producto/servicio " & intConcepto & " - " & strNombreConcepto
                fblnValidaSAT = False
                Exit Function
            End If
            If flngCatalogoSATIdByNombreTipo("c_ClaveUnidad", intConcepto, "CF", 2) = 0 Then
                vlstrError = "No está definida la clave del SAT para la unidad del producto/servicio " & intConcepto & " - " & strNombreConcepto
                fblnValidaSAT = False
                Exit Function
            End If
        End If
        fblnValidaSAT = True
    Else
       fblnValidaSAT = True
    End If

End Function

Public Function frsCveUsoCFDI(strNombreCatalogo As String, strDescripcion As String) As Integer
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    strSQL = "select GNCatalogoSATDetalle.intIdRegistro, GNCatalogoSATDetalle.vchClave, GNCatalogoSATDetalle.vchDescripcion" & _
    " from GNCatalogoSAT inner join GNCatalogoSATDetalle on GNCatalogoSAT.intIdCatalogoSAT = GNCatalogoSATDetalle.intIdCatalogoSAT" & _
    " where GNCatalogoSAT.vchNombreCatalogo = '" & strNombreCatalogo & "' and GNCatalogoSATDetalle.bitActivo = 1 and VCHCLAVE LIKE '%" & strDescripcion & "%'" & _
    " order by vchClave"
    Set rs = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
    frsCveUsoCFDI = IIf(IsNull(rs!intIdRegistro) Or rs.RecordCount = 0, -2, rs!intIdRegistro)
End Function

Public Sub pEnviarMasivo(strTipoDocumento As String, lngIdDocumento As Long, lngCveEmpresa As Long, strRFC As String, lngIdEmpleado As Long, vlstrCorreo As String, frmEnvia As Form)
On Error GoTo NotificaError
    Dim rsCorreo As New ADODB.Recordset
    Dim rsDestinatario As New ADODB.Recordset
    Dim rsFolio As New ADODB.Recordset
    Dim rsRutaPDF As New ADODB.Recordset
    Dim rsRutaXML As New ADODB.Recordset
    Dim strSentencia As String
    
    '- Verifica configuración de la cuenta de correo de la empresa -'
    Set rsCorreo = frsEjecuta_SP(CInt(lngCveEmpresa) & "|0", "Sp_CnSelCnCorreo")
    If rsCorreo.RecordCount = 0 Then
        'No se ha configurado la cuenta de correo.
        MsgBox SIHOMsg(1202), vbCritical, "Mensaje"
        Exit Sub
    End If
    
    '- Verifica el folio del documento a enviar -'
    strSentencia = "SELECT TRIM(VCHSERIECOMPROBANTE) || TRIM(VCHFOLIOCOMPROBANTE) Folio FROM GNCOMPROBANTEFISCALDIGITAL WHERE intComprobante = " & lngIdDocumento & " AND trim(CHRTIPOCOMPROBANTE) = " & "'" & Trim(strTipoDocumento) & "'"
    Set rsFolio = frsRegresaRs(strSentencia)
    If rsFolio.RecordCount > 0 Then
        frmDatosCorreo.strFolioDocumento = Trim(rsFolio!folio)
    Else
        'Error al procesar el folio del documento a enviar.
        MsgBox SIHOMsg(1199), vbCritical, "Mensaje"
        Exit Sub
    End If
    
    '- Verifica las rutas de los archivos PDF -'
    strSentencia = "SELECT trim(VCHRUTAPDF) RutaPDF FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & lngCveEmpresa
    Set rsRutaPDF = frsRegresaRs(strSentencia)
    If rsRutaPDF.RecordCount > 0 Then
        frmDatosCorreo.strRutaPDF = Trim(rsRutaPDF!RutaPDF)
    Else
        'No se ha configurado la ruta de los archivos PDF.
        MsgBox SIHOMsg(1200), vbCritical, "Mensaje"
        Exit Sub
    End If

    '- Verifica las rutas de los archivos XML -'
    strSentencia = "SELECT trim(VCHRUTAXML) RutaXML FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & lngCveEmpresa
    Set rsRutaXML = frsRegresaRs(strSentencia)
    If rsRutaXML.RecordCount > 0 Then
        frmDatosCorreo.strRutaXML = Trim(rsRutaXML!RutaXML)
    Else
        'No se ha configurado la ruta de los archivos XML.
        MsgBox SIHOMsg(1201), vbCritical, "Mensaje"
        Exit Sub
    End If
    
    frmDatosCorreo.strTipoDocumento = Trim(strTipoDocumento) 'Establece el valor del tipo de documento en la pantalla de envío
    frmDatosCorreo.lngIdDocumento = lngIdDocumento  'Establece la clave del documento segun el valor del tipo en la pantalla de envío
    frmDatosCorreo.lngEmpleado = lngIdEmpleado 'Se establece el valor del empleado que realiza el envío
    frmDatosCorreo.strCorreoDestinatario = vlstrCorreo
    frmDatosCorreo.strEnvioMasivo = True
    Dim valorNormal As Boolean
    valorNormal = vgblnAutomatico
    
    vgblnAutomatico = True
    
    'frmDatosCorreo.Show vbModal, frmEnvia 'Muestra la forma para envío de correo
    
    Call frmDatosCorreo.Form_Load
    Call frmDatosCorreo.cmdEnviar_Click
    
    If frmDatosCorreo.vlblnEnvio = True Then
        grdFacturas.TextMatrix(grdFacturas.RowSel, 13) = "TIMBRADA - ENVIADA"
    Else
        grdFacturas.TextMatrix(grdFacturas.RowSel, 13) = "TIMBRADA - NO ENVIADA"
    End If
    
    vgblnAutomatico = valorNormal
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrBloqueaCuenta"))
End Sub

Private Function fblnCuentaOrdenanteValida(vlblnTexto As Boolean, txtCuentaBancaria As String, claveFormaPago As String) As Boolean
    fblnCuentaOrdenanteValida = False
    
    If vlblnTexto Then
        If claveFormaPago = "" Then
            fblnCuentaOrdenanteValida = True
        Else
            If claveFormaPago = "02" Then
                If Len(Trim(txtCuentaBancaria)) = 11 Or Len(Trim(txtCuentaBancaria)) = 18 Then
                    fblnCuentaOrdenanteValida = True
                Else
                    fblnCuentaOrdenanteValida = False
                End If
            Else
                If claveFormaPago = "03" Then
                    If Len(Trim(txtCuentaBancaria)) = 10 Or Len(Trim(txtCuentaBancaria)) = 16 Or Len(Trim(txtCuentaBancaria)) = 18 Then
                        fblnCuentaOrdenanteValida = True
                    Else
                        fblnCuentaOrdenanteValida = False
                    End If
                Else
                    If claveFormaPago = "04" Then
                        If Len(Trim(txtCuentaBancaria)) = 16 Then
                            fblnCuentaOrdenanteValida = True
                        Else
                            fblnCuentaOrdenanteValida = False
                        End If
                    Else
                        If claveFormaPago = "05" Then
                            If Len(Trim(txtCuentaBancaria)) = 10 Or Len(Trim(txtCuentaBancaria)) = 11 Or Len(Trim(txtCuentaBancaria)) = 15 Or Len(Trim(txtCuentaBancaria)) = 16 Or Len(Trim(txtCuentaBancaria)) = 18 Or (Len(Trim(txtCuentaBancaria)) >= 10 And Len(Trim(txtCuentaBancaria)) <= 50) Then
                                fblnCuentaOrdenanteValida = True
                            Else
                                fblnCuentaOrdenanteValida = False
                            End If
                        Else
                            If claveFormaPago = "06" Then
                                If Len(Trim(txtCuentaBancaria)) = 10 Then
                                    fblnCuentaOrdenanteValida = True
                                Else
                                    fblnCuentaOrdenanteValida = False
                                End If
                            Else
                                If claveFormaPago = "28" Then
                                    If Len(Trim(txtCuentaBancaria)) = 16 Then
                                        fblnCuentaOrdenanteValida = True
                                    Else
                                        fblnCuentaOrdenanteValida = False
                                    End If
                                Else
                                    If claveFormaPago = "29" Then
                                        If Len(Trim(txtCuentaBancaria)) = 15 Or Len(Trim(txtCuentaBancaria)) = 16 Then
                                            fblnCuentaOrdenanteValida = True
                                        Else
                                            fblnCuentaOrdenanteValida = False
                                        End If
                                    Else
                                        fblnCuentaOrdenanteValida = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If claveFormaPago = "" Then
            fblnCuentaOrdenanteValida = True
        Else
            If claveFormaPago = "02" Then
                If Len(Trim(cboCuentasPrevias.Text)) = 11 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Then
                    fblnCuentaOrdenanteValida = True
                Else
                    fblnCuentaOrdenanteValida = False
                End If
            Else
                If claveFormaPago = "03" Then
                    If Len(Trim(cboCuentasPrevias.Text)) = 10 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Then
                        fblnCuentaOrdenanteValida = True
                    Else
                        fblnCuentaOrdenanteValida = False
                    End If
                Else
                    If claveFormaPago = "04" Then
                        If Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                            fblnCuentaOrdenanteValida = True
                        Else
                            fblnCuentaOrdenanteValida = False
                        End If
                    Else
                        If claveFormaPago = "05" Then
                            If Len(Trim(cboCuentasPrevias.Text)) = 10 Or Len(Trim(cboCuentasPrevias.Text)) = 11 Or Len(Trim(cboCuentasPrevias.Text)) = 15 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Or Len(Trim(cboCuentasPrevias.Text)) = 18 Or (Len(Trim(cboCuentasPrevias.Text)) >= 10 And Len(Trim(cboCuentasPrevias.Text)) <= 50) Then
                                fblnCuentaOrdenanteValida = True
                            Else
                                fblnCuentaOrdenanteValida = False
                            End If
                        Else
                            If claveFormaPago = "06" Then
                                If Len(Trim(cboCuentasPrevias.Text)) = 10 Then
                                    fblnCuentaOrdenanteValida = True
                                Else
                                    fblnCuentaOrdenanteValida = False
                                End If
                            Else
                                If claveFormaPago = "28" Then
                                    If Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                                        fblnCuentaOrdenanteValida = True
                                    Else
                                        fblnCuentaOrdenanteValida = False
                                    End If
                                Else
                                    If claveFormaPago = "29" Then
                                        If Len(Trim(cboCuentasPrevias.Text)) = 15 Or Len(Trim(cboCuentasPrevias.Text)) = 16 Then
                                            fblnCuentaOrdenanteValida = True
                                        Else
                                            fblnCuentaOrdenanteValida = False
                                        End If
                                    Else
                                        fblnCuentaOrdenanteValida = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
End Function

Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdFacturas.TextMatrix(1, 1) = "" Then
        Exit Sub
    End If
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
        pLlenaGridFacturasDetalle
    End If
End Sub
