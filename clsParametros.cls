VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-Variables de memoria
Dim vmstrIDHospital As String
Dim vmstrPswAdmin As String
Dim vmintCuartoPred As Integer
Dim vmintEstadoSalud As Integer
Dim vmintCveDisponible As Integer
Dim vmintCveOcupado As Integer
Dim vmintEmpresaActiva As Integer
Dim vmintForzarCobro As Integer
Dim vmintForzarAutorizacionSubRogados As Integer
Dim vmintAlertaSumaAsegurada As Integer
Dim vmintAdmEmpleado As Long
Dim vmintParticular As Integer
Dim vmblnNomina As Integer
Dim vmintTiempoMensajes As Integer
Dim vmstrArchivoWAV As String
Dim vmintCveDptoEmergencias As Integer
Dim vmintCveNacPredeterminada As Integer
Dim vmintCveProveedorAlmacenGral As Long
Dim vmintCostoMasIVA As Integer
Dim vmintIVAacreditable As Integer
Dim vmintCostoUrgAInterno As Integer
Dim vmintOcultarExpediente As Integer
Dim vmblnActualizacionAutomatica As Boolean
Dim vmblnAutoverificacion As Boolean
Dim vmblnEnvioPromociones As Boolean
Dim vmblnForzaLoteYCad As Boolean
Dim vmstrRutaReportes As String
Dim vmblnCentralizados As Boolean
Dim vmblnCtaRnaCtaMde As Boolean
Dim vmstrLinkBDCentralizada As String
Dim vmintDiasRequisicion As Integer
Dim vmintCveDeptoCE As Integer
Dim vmstrRFCFisicaInicial As String
Dim vmstrRFCFisicaFecha As String
Dim vmstrRFCFisicaHomoclave As String
Dim vmstrRFCMoralInicial As String
Dim vmstrRFCMoralFecha As String
Dim vmstrRFCMoralHomoclave As String
Dim vmintIVATasaHospiConcepto As Integer
Dim vmintDesglosaIVA As Integer
Dim vmintCveEmpresa As Integer
Dim vmdblPorcentajeRetIVA As Double
Dim vmdblporcentajeRetfletes As Double
Dim vmintfrecuenciacargosautomaticos As Double
Dim vmstrCargosExcluidos As String
Dim vmstrInstitucionEstablecimiento As String
Dim vmintAutorizarArtRequisicion As Integer
Dim vmintAutorizarCargosSolicitud As Integer
Dim vmintAutorizarCargosIndicacion As Integer
Dim vmblnActualizaExpediente As Boolean
Dim vmintCveIdiomaPredeterminado As Integer
Dim vmintBitManejaCuadroBasico As Integer
Dim vmintBitRequiereAutorizacion As Integer
Dim vmintCvePacienteSocio As Integer
Dim vmintInterfazPolizas As Integer
Dim vmintBitEnvioCFD As Integer
Dim vmintBitCancelaCFDiNoSAT As Integer
Dim vmintBitIVATipoCambioPago As Integer
Dim vmintCveAreaUci As Integer
Dim vmintCveAreaUcip As Integer
Dim vmblnFarmaciaSub As Boolean

Private Sub pRefresca(rs As ADODB.Recordset, intTipo As Integer)
    'Parametros generales, tabla SiParametro, sin clave de empresa contable
    If intTipo = 0 Then
        Do Until rs.EOF
            Select Case rs.Fields("vchNombre").Value
                Case "INTTIPOPARTICULAR": vmintParticular = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHIDHOSPITAL": vmstrIDHospital = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "INTCVENACIONALIDADPREDETERMINA": vmintCveNacPredeterminada = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHLINKBDCENTRALIZADA": vmstrLinkBDCentralizada = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "INTCVEDEPTOCONSULTAEXT": vmintCveDeptoCE = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITFORZARLOTEYCADUCIDAD": vmblnForzaLoteYCad = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITAUTOVERIFICACIONVERSION": vmblnAutoverificacion = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITENVIODEPROMOCIONES": vmblnEnvioPromociones = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHRUTAREPORTES": vmstrRutaReportes = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCFISICAINICIAL": vmstrRFCFisicaInicial = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCFISICAFECHA": vmstrRFCFisicaFecha = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCFISICAHOMOCLAVE": vmstrRFCFisicaHomoclave = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCMORALINICIAL": vmstrRFCMoralInicial = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCMORALFECHA": vmstrRFCMoralFecha = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "CHRRFCMORALHOMOCLAVE": vmstrRFCMoralHomoclave = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "VCHADMCONTRASENA": vmstrPswAdmin = IIf(IsNull(rs.Fields("vchValor").Value), "", rs.Fields("vchValor").Value)
                Case "MNYPORCENTAJERETIVA": vmdblPorcentajeRetIVA = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "TNYCVEDISPONIBLE": vmintCveDisponible = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "TNYCVEOCUPADO": vmintCveOcupado = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "RELPORCENTAJERETENCIONFLETES": vmdblporcentajeRetfletes = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTFRECUENCIACARGOSAUTOMATICOS": vmintfrecuenciacargosautomaticos = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEIDIOMAPREDETERMINADO": vmintCveIdiomaPredeterminado = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITFARMACIASUBROGADA": vmblnFarmaciaSub = IIf(IsNull(rs.Fields("vchvalor").Value), 0, rs.Fields("vchValor").Value)
                
            End Select
            rs.MoveNext
        Loop
    End If
    
    'Parametros por empresa, SiParametro
    If intTipo = 1 Then
        Do Until rs.EOF
            Select Case rs.Fields("vchNombre").Value
                Case "INTCVETIPOCUARTOPRED": vmintCuartoPred = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEESTADOSALUDPRED": vmintEstadoSalud = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEDEPTOEMERGENCIAS": vmintCveDptoEmergencias = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEPROVEEDORALMACENGENERAL": vmintCveProveedorAlmacenGral = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITTRASLADOAUTOMATICODECARGOS": vmblnActualizacionAutomatica = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITTRASLADOAUTOMATICOEXPEDIENTE": vmblnActualizaExpediente = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITAGRUPARCTARNACTAMADRE": vmblnCtaRnaCtaMde = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITCOSTOMASIVA": vmintCostoMasIVA = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITIVAACREDITABLE": vmintIVAacreditable = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITCOSTOURGAINTERNO": vmintCostoUrgAInterno = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITOCULTAEXPEDIENTE": vmintOcultarExpediente = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITDESGLOSAIVATASAHOSPICONCEPTO": vmintIVATasaHospiConcepto = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITFORZARAUTORIZACIONSUBROG": vmintForzarAutorizacionSubRogados = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITALERTASUMAASEGURADA": vmintAlertaSumaAsegurada = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITFORZARCOBRO": vmintForzarCobro = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHARCHIVOWAV": vmstrArchivoWAV = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTTIEMPOMENSAJES": vmintTiempoMensajes = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTDIASREQUISICION": vmintDiasRequisicion = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEEMPLEADOADM": vmintAdmEmpleado = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHCARGOSEXCLUIDOS": vmstrCargosExcluidos = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITAUTORIZARARTICULOSREQUISICION": vmintAutorizarArtRequisicion = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITAUTORIZARCARGOSOLICITUDSA": vmintAutorizarCargosSolicitud = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITAUTORIZARCARGOSINDICACIONSA": vmintAutorizarCargosIndicacion = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITMANEJACUADROBASICO": vmintBitManejaCuadroBasico = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITREQUIEREAUTORIZACION": vmintBitRequiereAutorizacion = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "VCHINSTITUCIONESTABLECIMIENTO": vmstrInstitucionEstablecimiento = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVETIPOPACIENTESOCIO": vmintCvePacienteSocio = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTINTERFAZPOLIZA": vmintInterfazPolizas = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITENVIOCFD": vmintBitEnvioCFD = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITCANCELACFDINOSAT": vmintBitCancelaCFDiNoSAT = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "BITIVATIPOCAMBIOPAGO": vmintBitIVATipoCambioPago = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEAREAUCI": vmintCveAreaUci = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
                Case "INTCVEAREAUCIP": vmintCveAreaUcip = IIf(IsNull(rs.Fields("vchValor").Value), 0, rs.Fields("vchValor").Value)
            End Select
            rs.MoveNext
        Loop
    End If
End Sub

Private Sub pLimpia(intTipo As Integer)
    'Parametros generales, tabla Parametros
    If intTipo = 0 Then
        vmintParticular = 0
        vmstrIDHospital = ""
        vmintCveNacPredeterminada = 0
        vmstrLinkBDCentralizada = ""
        vmintCveDeptoCE = 0
        vmblnForzaLoteYCad = 0
        vmblnAutoverificacion = 0
        vmblnEnvioPromociones = 0
        vmstrRutaReportes = ""
        vmstrRFCFisicaInicial = ""
        vmstrRFCFisicaFecha = ""
        vmstrRFCFisicaHomoclave = ""
        vmstrRFCMoralInicial = ""
        vmstrRFCMoralFecha = ""
        vmstrRFCMoralHomoclave = ""
        vmblnNomina = 0
        vmblnCentralizados = 0
        vmstrPswAdmin = ""
        vmdblPorcentajeRetIVA = 0
        vmintCveDisponible = 0
        vmintCveOcupado = 0
        vmdblporcentajeRetfletes = 0
        vmintfrecuenciacargosautomaticos = 0
        vmintCveIdiomaPredeterminado = 0
        vmblnFarmaciaSub = 0
    End If
    
    'Parametros por empresa, tabla SiParametro
    If intTipo = 1 Then
        vmintCuartoPred = 0
        vmintEstadoSalud = 0
        vmintCveDptoEmergencias = 0
        vmintCveProveedorAlmacenGral = 0
        vmblnActualizacionAutomatica = 0
        vmblnActualizaExpediente = 0
        vmblnCtaRnaCtaMde = 0
        vmintCostoMasIVA = 0
        vmintIVAacreditable = 0
        vmintCostoUrgAInterno = 0
        vmintOcultarExpediente = 0
        vmintIVATasaHospiConcepto = 0
        vmintDesglosaIVA = 0
        vmintForzarAutorizacionSubRogados = 0
        vmintAlertaSumaAsegurada = 0
        vmintForzarCobro = 0
        vmstrArchivoWAV = ""
        vmintTiempoMensajes = 0
        vmintDiasRequisicion = 0
        vmintAdmEmpleado = 0
        vmstrCargosExcluidos = ""
        vmintAutorizarArtRequisicion = 0
        vmintAutorizarCargosSolicitud = 0
        vmintAutorizarCargosIndicacion = 0
        vmintBitManejaCuadroBasico = 0
        vmintBitRequiereAutorizacion = 0
        vmstrInstitucionEstablecimiento = ""
        vmintCvePacienteSocio = 0
        vmintInterfazPolizas = 0
        vmintBitEnvioCFD = 0
        vmintBitCancelaCFDiNoSAT = 0
        vmintBitIVATipoCambioPago = 0
        vmintCveAreaUci = 0
        vmintCveAreaUcip = 0
    End If
End Sub

Public Property Get DptoCE() As Integer
   DptoCE = vmintCveDeptoCE
End Property

Public Property Let DptoCE(ByVal vData As Integer)
   vmintCveDeptoCE = vData
End Property

Public Property Get Nomina() As Boolean
   Nomina = vmblnNomina
End Property

Public Property Let Nomina(ByVal vData As Boolean)
   vmblnNomina = vData
End Property

Public Property Get DptoEmergencias() As Integer
   DptoEmergencias = vmintCveDptoEmergencias
End Property

Public Property Let DptoEmergencias(ByVal vData As Integer)
   vmintCveDptoEmergencias = vData
End Property

Public Property Get NacPredeterminada() As Integer
   NacPredeterminada = vmintCveNacPredeterminada
End Property

Public Property Let NacPredeterminada(ByVal vData As Integer)
   vmintCveNacPredeterminada = vData
End Property

Public Property Get PacienteSocio() As Integer
   PacienteSocio = vmintCvePacienteSocio
End Property

Public Property Let PacienteSocio(ByVal vData As Integer)
   vmintCvePacienteSocio = vData
End Property

Public Property Get IdiomaPredeterminado() As Integer
   IdiomaPredeterminado = vmintCveIdiomaPredeterminado
End Property

Public Property Let IdiomaPredeterminado(ByVal vData As Integer)
   vmintCveIdiomaPredeterminado = vData
End Property

Public Property Get ProveedorAlmacenGral() As Long
   ProveedorAlmacenGral = vmintCveProveedorAlmacenGral
End Property

Public Property Let ProveedorAlmacenGral(ByVal vData As Long)
   vmintCveProveedorAlmacenGral = vData
End Property

Public Property Get ForzarCobro() As Integer
   ForzarCobro = vmintForzarCobro
End Property

Public Property Let ForzarCobro(ByVal vData As Integer)
   vmintForzarCobro = vData
End Property

Public Property Get CostoIVA() As Integer
    CostoIVA = vmintCostoMasIVA
End Property

Public Property Let CostoIVA(ByVal vData As Integer)
    vmintCostoMasIVA = vData
End Property

Public Property Get IVAacreditable() As Integer
    IVAacreditable = vmintIVAacreditable
End Property

Public Property Let IVAacreditable(ByVal vData As Integer)
    vmintIVAacreditable = vData
End Property

Public Property Get CostoUrgAInterno() As Integer
    CostoUrgAInterno = vmintCostoUrgAInterno
End Property

Public Property Let CostoUrgAInterno(ByVal vData As Integer)
    vmintCostoUrgAInterno = vData
End Property

Public Property Get OcultarExpediente() As Integer
    OcultarExpediente = vmintOcultarExpediente
End Property

Public Property Let OcultarExpediente(ByVal vData As Integer)
    vmintOcultarExpediente = vData
End Property
Public Property Get Particular() As Integer
   Particular = vmintParticular
End Property
Public Property Let Particular(ByVal vData As Integer)
   vmintParticular = vData
End Property

Public Property Get AdmEmpleado() As Long
   AdmEmpleado = vmintAdmEmpleado
End Property

Public Property Let AdmEmpleado(ByVal vData As Long)
   vmintAdmEmpleado = vData
End Property

Public Property Get EmpresaActiva() As Integer
   EmpresaActiva = vmintEmpresaActiva
End Property

Public Property Let EmpresaActiva(ByVal vData As Integer)
   vmintEmpresaActiva = vData
End Property

Public Property Get CveOcupado() As Integer
   CveOcupado = vmintCveOcupado
End Property

Public Property Let CveOcupado(ByVal vData As Integer)
   vmintCveOcupado = vData
End Property

Public Property Get CveDisponible() As Integer
   CveDisponible = vmintCveDisponible
End Property

Public Property Let CveDisponible(ByVal vData As Integer)
   vmintCveDisponible = vData
End Property

Public Property Get EstadoSalud() As Integer
   EstadoSalud = vmintEstadoSalud
End Property

Public Property Let EstadoSalud(ByVal vData As Integer)
   vmintEstadoSalud = vData
End Property

Public Property Get CuartoPred() As Integer
   CuartoPred = vmintCuartoPred
End Property

Public Property Let CuartoPred(ByVal vData As Integer)
   vmintCuartoPred = vData
End Property

Public Property Get RFCFisicaInicial() As String
   RFCFisicaInicial = vmstrRFCFisicaInicial
End Property

Public Property Let RFCFisicaInicial(ByVal vData As String)
    vmstrRFCFisicaInicial = vData
End Property

Public Property Get RFCFisicaFecha() As String
   RFCFisicaFecha = vmstrRFCFisicaFecha
End Property

Public Property Let RFCFisicaFecha(ByVal vData As String)
    vmstrRFCFisicaFecha = vData
End Property

Public Property Get RFCFisicaHomoclave() As String
   RFCFisicaHomoclave = vmstrRFCFisicaHomoclave
End Property

Public Property Let RFCFisicaHomoclave(ByVal vData As String)
    vmstrRFCFisicaHomoclave = vData
End Property

Public Property Get RFCMoralInicial() As String
   RFCMoralInicial = vmstrRFCMoralInicial
End Property

Public Property Let RFCMoralInicial(ByVal vData As String)
    vmstrRFCMoralInicial = vData
End Property

Public Property Get RFCMoralFecha() As String
   RFCMoralFecha = vmstrRFCMoralFecha
End Property

Public Property Let RFCMoralFecha(ByVal vData As String)
    vmstrRFCMoralFecha = vData
End Property

Public Property Get RFCMoralHomoclave() As String
   RFCMoralHomoclave = vmstrRFCMoralHomoclave
End Property

Public Property Let RFCMoralHomoclave(ByVal vData As String)
    vmstrRFCMoralHomoclave = vData
End Property

Public Property Get IVATasaHospitalConceptoSeg() As Integer
   IVATasaHospitalConceptoSeg = vmintIVATasaHospiConcepto
End Property

Public Property Let IVATasaHospitalConceptoSeg(ByVal vData As Integer)
    vmintIVATasaHospiConcepto = vData
End Property

Public Property Get DesglosaIVA() As Integer
   DesglosaIVA = vmintDesglosaIVA
End Property

Public Property Let DesglosaIVA(ByVal vData As Integer)
    vmintDesglosaIVA = vData
End Property

Public Property Get PswAdmin() As String
   PswAdmin = vmstrPswAdmin
End Property

Public Property Let PswAdmin(ByVal vData As String)
   vmstrPswAdmin = vData
End Property

Public Property Get TiempoMensajes() As Integer
   TiempoMensajes = vmintTiempoMensajes
End Property

Public Property Let TiempoMensajes(ByVal vData As Integer)
   vmintTiempoMensajes = vData
End Property

Public Property Get ArchivoWAV() As String
   ArchivoWAV = vmstrArchivoWAV
End Property

Public Property Let ArchivoWAV(ByVal vData As String)
   vmstrArchivoWAV = vData
End Property

Public Property Get ForzarAutorizacionSubrog() As Integer
   ForzarAutorizacionSubrog = vmintForzarAutorizacionSubRogados
End Property

Public Property Get AlertaSumaAsegurada() As Integer
   AlertaSumaAsegurada = vmintAlertaSumaAsegurada
End Property

Public Property Let ForzarAutorizacionSubrog(ByVal vData As Integer)
   vmintForzarAutorizacionSubRogados = vData
End Property

Public Property Let AlertaSumaAsegurada(ByVal vData As Integer)
   vmintAlertaSumaAsegurada = vData
End Property

Public Property Get IdHospital() As String
   IdHospital = vmstrIDHospital
End Property

Public Property Let IdHospital(ByVal vData As String)
   vmstrIDHospital = vData
End Property

Public Property Get ActualizacionAutomatica() As Boolean
   ActualizacionAutomatica = vmblnActualizacionAutomatica
End Property

Public Property Let ActualizacionAutomatica(ByVal vData As Boolean)
   vmblnActualizacionAutomatica = vData
End Property

Public Property Get ActualizaExpediente() As Boolean
   ActualizaExpediente = vmblnActualizaExpediente
End Property

Public Property Let ActualizaExpediente(ByVal vData As Boolean)
   vmblnActualizaExpediente = vData
End Property

Public Property Get ForzaLoteYCad() As Boolean
   ForzaLoteYCad = vmblnForzaLoteYCad
End Property

Public Property Let ForzaLoteYCad(ByVal vData As Boolean)
   vmblnForzaLoteYCad = vData
End Property

Public Property Get Autoverificacion() As Boolean
   Autoverificacion = vmblnAutoverificacion
End Property

Public Property Get EnvioPromociones() As Boolean
   EnvioPromociones = vmblnEnvioPromociones
End Property

Public Property Let Autoverificacion(ByVal vData As Boolean)
   vmblnAutoverificacion = vData
End Property

Public Property Let EnvioPromociones(ByVal vData As Boolean)
   vmblnEnvioPromociones = vData
End Property

Public Property Get RutaReportes() As String
   RutaReportes = vmstrRutaReportes
End Property

Public Property Let RutaReportes(ByVal vData As String)
   vmstrRutaReportes = vData
End Property

Public Property Get Centralizados() As Boolean
   Centralizados = vmblnCentralizados
End Property

Public Property Let Centralizados(ByVal vData As Boolean)
   vmblnCentralizados = vData
End Property

Public Property Get AgruparCtaRnaACtaMde() As Boolean
   AgruparCtaRnaACtaMde = vmblnCtaRnaCtaMde
End Property

Public Property Let AgruparCtaRnaACtaMde(ByVal vData As Boolean)
   vmblnCtaRnaCtaMde = vData
End Property

Public Property Get EnlaceBDCentralizada() As String
   EnlaceBDCentralizada = vmstrLinkBDCentralizada
End Property

Public Property Let EnlaceBDCentralizada(ByVal vData As String)
   vmstrLinkBDCentralizada = vData
End Property

Public Property Get DiasRequisicion() As Integer
   DiasRequisicion = vmintDiasRequisicion
End Property

Public Property Let DiasRequisicion(ByVal vData As Integer)
   vmintDiasRequisicion = vData
End Property

Public Property Get PorcentajeRetIVA() As Double
    PorcentajeRetIVA = vmdblPorcentajeRetIVA
End Property

Public Property Let PorcentajeRetIVA(ByVal vData As Double)
    vmdblPorcentajeRetIVA = vData
End Property

Public Property Get PorcentajeRetfletes() As Double
    PorcentajeRetfletes = vmdblporcentajeRetfletes
End Property

Public Property Let PorcentajeRetfletes(ByVal vData As Double)
    vmdblporcentajeRetfletes = vData
End Property

Public Property Get FrecuenciaCargosAutomaticos() As Double
    FrecuenciaCargosAutomaticos = vmintfrecuenciacargosautomaticos
End Property

Public Property Let FrecuenciaCargosAutomaticos(ByVal vData As Double)
    vmintfrecuenciacargosautomaticos = vData
End Property

Public Property Get CargosExcluidos() As String
    CargosExcluidos = vmstrCargosExcluidos
End Property

Public Property Let CargosExcluidos(ByVal vData As String)
    vmstrCargosExcluidos = vData
End Property

'---------------------------------------------------
Public Property Get InstitucionEstablecimiento() As String
    InstitucionEstablecimiento = vmstrInstitucionEstablecimiento
End Property

Public Property Let InstitucionEstablecimiento(ByVal vData As String)
    vmstrInstitucionEstablecimiento = vData
End Property
'--------------------------------------------------

Public Property Get AutorizarArticulosRequisicion() As Integer
   AutorizarArticulosRequisicion = vmintAutorizarArtRequisicion
End Property

Public Property Let AutorizarArticulosRequisicion(ByVal vData As Integer)
    vmintAutorizarArtRequisicion = vData
End Property

Public Property Get AutorizarCargosSolicitud() As Integer
   AutorizarCargosSolicitud = vmintAutorizarCargosSolicitud
End Property

Public Property Let AutorizarCargosSolicitud(ByVal vData As Integer)
    vmintAutorizarCargosSolicitud = vData
End Property

Public Property Get AutorizarCargosIndicacionSA() As Integer
   AutorizarCargosIndicacionSA = vmintAutorizarCargosIndicacion
End Property

Public Property Let AutorizarCargosIndicacionSA(ByVal vData As Integer)
    vmintAutorizarCargosIndicacion = vData
End Property

Public Property Get BitManejaCuadroBasico() As Integer
   BitManejaCuadroBasico = vmintBitManejaCuadroBasico
End Property

Public Property Let BitManejaCuadroBasico(ByVal vData As Integer)
    vmintBitManejaCuadroBasico = vData
End Property

Public Property Get BitRequiereAutorizacion() As Integer
   BitRequiereAutorizacion = vmintBitRequiereAutorizacion
End Property

Public Property Let BitRequiereAutorizacion(ByVal vData As Integer)
    vmintBitRequiereAutorizacion = vData
End Property

Public Property Get InterfazPolizas() As Integer
   InterfazPolizas = vmintInterfazPolizas
End Property

Public Property Let InterfazPolizas(ByVal vData As Integer)
    vmintInterfazPolizas = vData
End Property

Public Property Get BitEnvioCFD() As Integer
   BitEnvioCFD = vmintBitEnvioCFD
End Property

Public Property Let BitEnvioCFD(ByVal vData As Integer)
    vmintBitEnvioCFD = vData
End Property

Public Property Get BitCancelaCFDiNoSAT() As Integer
   BitCancelaCFDiNoSAT = vmintBitCancelaCFDiNoSAT
End Property

Public Property Let BitCancelaCFDiNoSAT(ByVal vData As Integer)
   vmintBitCancelaCFDiNoSAT = vData
End Property
Public Property Get BitIVATipoCambioPago() As Integer
   BitIVATipoCambioPago = vmintBitIVATipoCambioPago
End Property

Public Property Let BitIVATipoCambioPago(ByVal vData As Integer)
   vmintBitIVATipoCambioPago = vData
End Property

Public Property Get AreaUci() As Integer
   AreaUci = vmintCveAreaUci
End Property

Public Property Let AreaUci(ByVal vData As Integer)
   vmintCveAreaUci = vData
End Property

Public Property Get AreaUcip() As Integer
   AreaUcip = vmintCveAreaUcip
End Property

Public Property Let AreaUcip(ByVal vData As Integer)
   vmintCveAreaUcip = vData
End Property
Public Property Get FarmaciaSub() As Boolean
   FarmaciaSub = vmblnFarmaciaSub
End Property

Public Property Let FarmaciaSub(ByVal vData As Boolean)
   vmblnFarmaciaSub = vData
End Property
Public Function fblnCargaRs(Optional intcveempresa As Integer = -1) As Boolean
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    vmintCveEmpresa = intcveempresa
    strSql = "select * from SIParametro where chrModulo = 'SI' and intCveEmpresaContable is null"
    Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        pRefresca rs, 0
        fblnCargaRs = True
    Else
        pLimpia 0
        fblnCargaRs = False
    End If
    rs.Close
    
    strSql = "select * from SIParametro where chrModulo = 'SI' and intCveEmpresaContable = " & intcveempresa
    Set rs = frsRegresaRs(strSql, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        pLimpia 1
        pRefresca rs, 1
    Else
        pLimpia 1
    End If
    rs.Close
End Function

Public Sub pGuarda()
    frsEjecuta_SP "-1|INTTIPOPARTICULAR|SI|" & vmintParticular & "|Clave del tipo de paciente que corresponde a particular (relacionado con AdTipoPaciente)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|VCHIDHOSPITAL|SI|" & vmstrIDHospital & "|Identificador del Hospital Parametro general que permite identificar a el Hospital", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|INTCVENACIONALIDADPREDETERMINA|SI|" & vmintCveNacPredeterminada & "|Clave de la nacionalidad que se utiliza en forma predeterminada (relacionado con Nacionalidad)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|VCHLINKBDCENTRALIZADA|SI|" & vmintCveNacPredeterminada & "|Especifíca el nombre del enlace por medio del cual se tendrá acceso al servidor central", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|INTCVEDEPTOCONSULTAEXT|SI|" & vmintCveDeptoCE & "|Clave del departamento de consulta externa", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|BITFORZARLOTEYCADUCIDAD|SI|" & IIf(vmblnForzaLoteYCad, "1", "0") & "|Forzar captura de lote y caducidades (0 = No forzar, 1 = Si forzar la captura)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|BITAUTOVERIFICACIONVERSION|SI|" & IIf(vmblnAutoverificacion, "1", "0") & "|Indica si se va a verificar que la version sea correcta 1 = se verifica, 0 = no se verifica", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|BITENVIODEPROMOCIONES|SI|" & IIf(vmblnEnvioPromociones, "1", "0") & "|Solicitar confirmación del paciente al momento de la admisión para recibir información de promociones", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|VCHRUTAREPORTES|SI|" & vmstrRutaReportes & "|Ruta de red de los reportes", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCFISICAINICIAL|SI|" & vmstrRFCFisicaInicial & "|Caracteres del inicio de un RFC de persona física, ejemplo: AAAA", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCFISICAFECHA|SI|" & vmstrRFCFisicaFecha & "|Caracteres de la fecha de un RFC de persona física, ejemplo: 999999", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCFISICAHOMOCLAVE|SI|" & vmstrRFCFisicaHomoclave & "|Caracteres de la homoclave de un RFC de persona física ejemplo: AAA", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCMORALINICIAL|SI|" & vmstrRFCMoralInicial & "|Caracteres del inicio de un RFC de persona moral, ejemplo: AAA", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCMORALFECHA|SI|" & vmstrRFCMoralFecha & "|Caracteres de la fecha de un RFC de persona moral, ejemplo: 999999", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|CHRRFCMORALHOMOCLAVE|SI|" & vmstrRFCMoralHomoclave & "|Caracteres de la homoclave de un RFC de persona moral, ejemplo: AAA", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|VCHADMCONTRASENA|SI|" & vmstrPswAdmin & "|Contraseña encriptada del administrador del sistema", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|MNYPORCENTAJERETIVA|SI|" & vmdblPorcentajeRetIVA & "|Porcentaje de retención de IVA sobre honorarios", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|TNYCVEDISPONIBLE|SI|" & vmintCveDisponible & "|Clave del estado de cuarto que corresponde a disponible (relacionado con AdEstadoCuarto)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|TNYCVEOCUPADO|SI|" & vmintCveOcupado & "|Clave del estado de cuarto que corresponde a ocupado (relacionado con AdEstadoCuarto)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|RELPORCENTAJERETENCIONFLETES|SI|" & vmdblporcentajeRetfletes & "|Clave del estado de cuarto que corresponde a ocupado (relacionado con AdEstadoCuarto)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|INTFRECUENCIACARGOSAUTOMATICOS|SI|" & vmintfrecuenciacargosautomaticos & "|Frecuencia en minutos para la ejecución del proceso de cargos automáticos por cuidados especiales", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|INTCVEIDIOMAPREDETERMINADO|SI|" & vmintCveIdiomaPredeterminado & "|Clave del idioma que se utiliza en forma predeterminado (relacionado con GnIdiomas)", "sp_GNUpdParametro", True
    frsEjecuta_SP "-1|BITFARMACIASUBROGADA|SI|" & IIf(vmblnFarmaciaSub, "1", "0") & "|Indica si el hospital cuenta con farmacia subrogada", "sp_GNUpdParametro", True

    If vmintCveEmpresa > 0 Then
        frsEjecuta_SP vmintCveEmpresa & "|INTCVETIPOCUARTOPRED|SI|" & vmintCuartoPred & "|Clave del tipo de cuarto que se utiliza en forma predeterminada en la admisión (relacionado con ADTipoCuarto)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEESTADOSALUDPRED|SI|" & vmintEstadoSalud & "|Clave del estado de salud que se utiliza en forma predeterminada en la admisión (relacionado con EXEstadoSalud)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEDEPTOEMERGENCIAS|SI|" & vmintCveDptoEmergencias & "|Clave del departamento que corresponde a emergencias (relacionado con NODepartamento)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEPROVEEDORALMACENGENERAL|SI|" & vmintCveProveedorAlmacenGral & "|Clave del proveedor que corresponde al almacén de abastecimiento (relacionado con COProveedor)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITTRASLADOAUTOMATICODECARGOS|SI|" & IIf(vmblnActualizacionAutomatica, "1", "0") & "|Indica si al momento del ingreso en admisión de un paciente externo, se realiza de manera automática un traslado de cargos", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITTRASLADOAUTOMATICOEXPEDIENTE|SI|" & IIf(vmblnActualizaExpediente, "1", "0") & "|Indica si al momento del ingreso en admisión de un paciente externo (INTERNO FUE EXTERNO), se realiza de manera automática un traslado del expediente", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITAGRUPARCTARNACTAMADRE|SI|" & IIf(vmblnCtaRnaCtaMde, "1", "0") & "|Forma en que se aplican los cargos de la cta del RN (1 a la cta de la Madre, 0 a la cta del RN)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITCOSTOMASIVA|SI|" & vmintCostoMasIVA & "|Indica si el calculo del costo promedio va a ser incluyendo el IVA de la recepción", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITIVAACREDITABLE|SI|" & vmintIVAacreditable & "|Indica si contabilizar o no el IVA acreditable no pagado al registrar la compra del servicio.", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITCOSTOURGAINTERNO|SI|" & vmintCostoUrgAInterno & "|Indica si el costo de pacientes de urgencias que se internaron tendrá costo de 0 o no en el reporte de uso de expediente.", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITOCULTAEXPEDIENTE|SI|" & vmintOcultarExpediente & "|Indica si se oculta el número de expediente en reporte (guarda 1 = si se oculta, 0 = no se oculta)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITDESGLOSAIVATASAHOSPICONCEPTO|SI|" & vmintIVATasaHospiConcepto & "|Indica si se desglosa el IVA a la tasa del hospital en los conceptos de seguro que desglosan IVA", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITFORZARAUTORIZACIONSUBROG|SI|" & vmintForzarAutorizacionSubRogados & "|Indica si se debe exigir la validación de los estudios subrogados", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITALERTASUMAASEGURADA|SI|" & vmintAlertaSumaAsegurada & "|Indica si se mostrará una alerta en caja y crédito por pacientes que excedan la suma asegurada", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITFORZARCOBRO|SI|" & vmintForzarCobro & "|Indica si el sistema debe de validar el cobro de un examen o de un estudio al momento de la entrega de los mismos", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|VCHARCHIVOWAV|SI|" & vmstrArchivoWAV & "|Sonido para los mensajes", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTTIEMPOMENSAJES|SI|" & vmintTiempoMensajes & "|Tiempo de espera para actualizar mensajes", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTDIASREQUISICION|SI|" & vmintDiasRequisicion & "|Número de días que se tomaran en cuenta para revisar las requisiciones pendientes", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEEMPLEADOADM|SI|" & vmintAdmEmpleado & "|Clave del empleado que es administrador del sistema (relacionado con NOEmpleado)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|VCHCARGOSEXCLUIDOS|SI|" & vmstrCargosExcluidos & "|Texto a mostrar en el documento de cargos excluidos de la admisión", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|VCHINSTITUCIONESTABLECIMIENTO|SI|" & vmstrInstitucionEstablecimiento & "|Nombre de la institución a la que pertenece el establecimiento", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITAUTORIZARARTICULOSREQUISICION|SI|" & vmintAutorizarArtRequisicion & "|Indica si se deben introducir los datos de autorización del artículo al generar la requisición cuando el paciente pertenece a una empresa y se esté solicitando un artículo no cubierto", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITAUTORIZARCARGOSOLICITUDSA|SI|" & vmintAutorizarCargosSolicitud & "|Indica si se deben introducir los datos de autorización al generar la solicitud de servicios auxiliares o laboratorio cuando el cargo es excluido por la empresa y con estatus de autorización", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITAUTORIZARCARGOSINDICACIONSA|SI|" & vmintAutorizarCargosIndicacion & "|Indica si se deben introducir los datos de autorización al generar la indicación de servicios auxiliares o laboratorio cuando el cargo es excluido por la empresa y con estatus de autorización", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITMANEJACUADROBASICO|SI|" & vmintBitManejaCuadroBasico & "|Indica si el hospital maneja un cuadro básico de medicamentos", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITREQUIEREAUTORIZACION|SI|" & vmintBitRequiereAutorizacion & "|Indica si se requiere de autorización cuando se realiza una requisición o indicación de un medicamento que no forma parte del cuadro básico", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVETIPOPACIENTESOCIO|SI|" & vmintCvePacienteSocio & "|Clave del tipo de paciente que será del tipo socio (relacionado con AdTipoPaciente)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTINTERFAZPOLIZA|SI|" & vmintInterfazPolizas & "|Interfaz utilizada para la importación y exportación de pólizas, 1 = CONTPAQ, 2 = Microsip.", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITENVIOCFD|SI|" & vmintBitEnvioCFD & "|Indica si se mostrará el mensaje para el envío de CFD/CFDi", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITCANCELACFDINOSAT|SI|" & vmintBitCancelaCFDiNoSAT & "|Indica si se pemitirá cancelar un CFDi cuando el sistema de cancelación del SAT no está disponible", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|BITIVATIPOCAMBIOPAGO|SI|" & vmintBitIVATipoCambioPago & "|Indica si se debe tomar el tipo de cambio de la fecha en que se realiza el pago para calcular el IVA acreditable y las retenciones de ISR,IVA y fletes", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEAREAUCI|SI|" & vmintCveAreaUci & "|Indica el área para UCI (relacionado con AdArea)", "sp_GNUpdParametro", True
        frsEjecuta_SP vmintCveEmpresa & "|INTCVEAREAUCIP|SI|" & vmintCveAreaUcip & "|Indica el área para UCIP (relacionado con AdArea)", "sp_GNUpdParametro", True
    End If
End Sub

Public Function fintAltaParametros(intcveempresa As Integer) As Integer
On Error GoTo NotificaError

    vmintCveEmpresa = intcveempresa
    pGuarda
    fintAltaParametros = 1
    
Exit Function
NotificaError:
     fintAltaParametros = 0
End Function

