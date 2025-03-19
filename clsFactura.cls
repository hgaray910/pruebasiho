VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------------
' Clase factura,
' Fecha: Septiembre, 2006
' <flngInsFactura> = para insertar facturas, regresa el consecutivo
'------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim lstrFolio As String
Dim ldtmFecha As Date
Dim lstrRFC As String
Dim lstrNombre As String
Dim lstrDireccion As String
Dim llngCveCiudad As Long
Dim lstrTelefono As String
Dim ldblIVA As Double
Dim ldblDescuento As Double
Dim lstrEstado As String
Dim llngCuenta As Long
Dim lstrTipoPaciente As String
Dim ldblDeducible As Double
Dim ldblExcedenteSuma As Double
Dim ldblCoaseguro As Double
Dim ldblCopago As Double
Dim lintDepartamento As Integer
Dim llngEmpleado As Long
Dim lintTipoDeducible As String
Dim llngNumeroCorte As Long
Dim ldblAnticipo As Double
Dim ldblTotal As Double
Dim lintTipoMoneda As Integer
Dim ldblTipoCambio As Double
Dim lstrTipoFactura As String
Dim llngNumeroCliente As Long
Dim llngNumeroVenta As Long
Dim llngNumeroPoliza As Long



Public Property Let Folio(ByVal Dato As String)
    lstrFolio = Dato
End Property
Public Property Let fecha(ByVal Dato As Date)
    ldtmFecha = Dato
End Property
Public Property Let RFC(ByVal Dato As String)
    lstrRFC = Dato
End Property
Public Property Let Nombre(ByVal Dato As String)
    lstrNombre = Dato
End Property
Public Property Let Direccion(ByVal Dato As String)
    lstrDireccion = Dato
End Property

Public Property Let IdCiudad(ByVal Dato As String)
    llngCveCiudad = Dato
End Property

Public Property Let Telefono(ByVal Dato As String)
    lstrTelefono = Dato
End Property
Public Property Let IVA(ByVal Dato As Double)
    ldblIVA = Dato
End Property
Public Property Let Descuento(ByVal Dato As Double)
    ldblDescuento = Dato
End Property
Public Property Let Estado(ByVal Dato As String)
    lstrEstado = Dato
End Property
Public Property Let cuenta(ByVal Dato As Long)
    llngCuenta = Dato
End Property
Public Property Let TipoPaciente(ByVal Dato As String)
    lstrTipoPaciente = Dato
End Property
Public Property Let Deducible(ByVal Dato As Double)
    ldblDeducible = Dato
End Property
Public Property Let ExcedenteSuma(ByVal Dato As Double)
    ldblExcedenteSuma = Dato
End Property
Public Property Let Coaseguro(ByVal Dato As Double)
    ldblCoaseguro = Dato
End Property
Public Property Let Copago(ByVal Dato As Double)
    ldblCopago = Dato
End Property
Public Property Let Departamento(ByVal Dato As Long)
    lintDepartamento = Dato
End Property
Public Property Let Empleado(ByVal Dato As Long)
    llngEmpleado = Dato
End Property
Public Property Let TipoDeducible(ByVal Dato As Integer)
    lintTipoDeducible = Dato
End Property
Public Property Let NumeroCorte(ByVal Dato As Long)
    llngNumeroCorte = Dato
End Property
Public Property Let Anticipo(ByVal Dato As Double)
    ldblAnticipo = Dato
End Property
Public Property Let Total(ByVal Dato As Double)
    ldblTotal = Dato
End Property
Public Property Let TipoMoneda(ByVal Dato As Integer)
    lintTipoMoneda = Dato
End Property
Public Property Let TipoCambio(ByVal Dato As Double)
    ldblTipoCambio = Dato
End Property
Public Property Let TipoFactura(ByVal Dato As String)
    lstrTipoFactura = Dato
End Property
Public Property Let NumeroCliente(ByVal Dato As Long)
    llngNumeroCliente = Dato
End Property
Public Property Let NumeroVenta(ByVal Dato As Long)
    llngNumeroVenta = Dato
End Property
Public Property Let NumeroPoliza(ByVal Dato As String)
    llngNumeroPoliza = Dato
End Property

Public Function flngInsFactura( _
strFolio As String, dtmfecha As Date, _
strRFC As String, strNombre As String, _
strDireccion As String, strNumeroExterior As String, _
strNumeroInterior As String, dblIVA As Double, _
dblDescuento As Double, strEstadoFactura As String, _
lngCuenta As Long, strTipoPaciente As String, _
intDepartamento As Integer, lngEmpleado As Long, _
lngNumeroCorte As Long, dblAnticipo As Double, _
dblTotal As Double, intTipoMoneda As Integer, _
dblTipoCambio As Double, strTelefono As String, _
strTipoFactura As String, lngNumeroCliente As Long, _
lngNumeroVenta As Long, lngNumeroPoliza As Long, _
strCalleNumero As String, strColonia As String, _
strCiudad As String, strEstado As String, _
strCodigo As String, lngCveImpuesto As Long, _
lngCveCiudad As Long, strFolioSeparado As String, _
strSerie As String, Optional intCveUsoCFDI As Long = 45, _
Optional dblRetencionISR As Double, Optional dblRetencionIVA As Double, _
Optional intTipoFactura As Integer, Optional intIdTarifaRetencionISR As Integer, Optional dblRetServ As Double = 0, _
Optional strObservaciones As String, Optional StrREgimenFiscal As String) As Long


'-------------------------------------------------------------------------------------------------------------
' Inserta la factura y regresa el consecutivo <PvFactura.intConsecutivo>
'-------------------------------------------------------------------------------------------------------------
    Dim strSentencia As String
    Dim rs As New ADODB.Recordset
    
    strSentencia = "select PvFactura.* from PvFactura where PvFactura.intConsecutivo = -1"
    Set rs = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    
    rs.AddNew
    rs!chrfoliofactura = strFolio
    rs!dtmFechahora = dtmfecha
    rs!CHRRFC = Trim(strRFC)
    rs!CHRNOMBRE = Trim(strNombre)
    'rs!chrDireccion = Trim(strDireccion)
    rs!chrCalle = Trim(strDireccion)
    rs!VCHNUMEROEXTERIOR = Trim(strNumeroExterior)
    rs!VCHNUMEROINTERIOR = Trim(strNumeroInterior)
    rs!vchSerie = Trim(strSerie)
    rs!INTFOLIO = Val(strFolioSeparado)
    rs!smyIVA = dblIVA
    rs!MNYDESCUENTO = dblDescuento
    rs!chrEstatus = strEstadoFactura
    rs!INTMOVPACIENTE = lngCuenta
    rs!CHRTIPOPACIENTE = strTipoPaciente
    rs!SMIDEPARTAMENTO = intDepartamento
    rs!intCveEmpleado = lngEmpleado
    rs!intNumCorte = lngNumeroCorte
    rs!mnyAnticipo = dblAnticipo
    rs!mnyTotalFactura = dblTotal
    rs!BITPESOS = intTipoMoneda
    rs!mnytipocambio = dblTipoCambio
    rs!chrTelefono = Trim(strTelefono)
    rs!chrTipoFactura = strTipoFactura
    rs!intNumCliente = lngNumeroCliente
    rs!intCveVentaPublico = lngNumeroVenta
    rs!intNumPoliza = lngNumeroPoliza
    rs!vchCalleNumero = Trim(strCalleNumero)
    rs!VCHCOLONIA = Trim(strColonia)
    rs!VCHCIUDAD = Trim(strCiudad)
    rs!vchEstado = Trim(strEstado)
    rs!VCHCODIGOPOSTAL = Trim(strCodigo)
    rs!intCveImpuesto = lngCveImpuesto
    rs!INTCVECIUDAD = lngCveCiudad
    rs!intcveempresa = 0
    rs!mnyTotalPagar = dblTotal
    rs!intTipoDetalleFactura = 1 'para facturas directas y socios se coloca 1 (agrupada por concepto de facturación)
    If intCveUsoCFDI > 0 Then
        rs!intCveUsoCFDI = intCveUsoCFDI
    End If
    rs!MNYRETENCIONISR = dblRetencionISR
    rs!MNYRETENCIONIVA = dblRetencionIVA
    rs!intTipoFactura = intTipoFactura
    rs!intIdTarifaRetencionISR = intIdTarifaRetencionISR
    rs!MNYRETENSERVICIOS = dblRetServ
    rs!CHROBSERVACIONES = strObservaciones
    rs!VCHREGIMENFISCALRECEPTOR = StrREgimenFiscal
    
    rs.Update
    flngInsFactura = flngObtieneIdentity("SEC_PVFACTURA", rs!intConsecutivo)
    rs.Close
 
End Function



