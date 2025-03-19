Attribute VB_Name = "modCaja"
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Módulo        : modCaja.bas
'-------------------------------------------------------------------------------------
'| Objetivo: Permite crear constantes, banderas y variables públicas que serán utilizadas en el
'| proyecto de Caja.
'-------------------------------------------------------------------------------------

'IMPORTANTE
'IMPORTANTE
'IMPORTANTE
'IMPORTANTE
'IMPORTANTE
'IMPORTANTE
'SI ES NECESARIO MOVER ALGO RELACIONADO A LA FACTURACIÓN DIRECTA, ES NECESARIO HACER EL MISMO CAMBIO EN MODCREDITO EN EL MÓDULO DE CRÉDITO

Option Private Module
Option Explicit
'/****************************************************************************
'Variables declaradas durante la migración a ORACLE
Public Type CtrlPermiso
    vlstrObjeto As String
    vlstrForma As String
    vlIntOrdenTab As Long
    vlstrTipoPermiso As String
    vlintTotalRegistros As Integer
End Type
Public vgblnTerminate As Boolean
Global vgstrBaseDatosUtilizada As String  'Variable para distinguir que base de datos es: "MSSQL" y "ORACLE"
Public First_time As Boolean
Global vgblnNomina As Boolean                   'Si la empresa va a utilizar el modulo de nomina
Global vgstrParametrosSP As String
Public aControlesModulo() As CtrlPermiso 'Arreglo para cargar todos los controles del modulo.

Public Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long

Public Enum TipoOperacion
        EnmLogIn = 1
        EnmGrabar = 2
        EnmBorrar = 3
        EnmCambiar = 4
        EnmCancelacion = 5
        EnmLogout = 6
        EnmReImpresion = 7
        EnmConsulta = 8
        EnmPinPad = 9
End Enum

Public Type DirectaMasiva
    intDirectaMasiva As Integer
    strReferenciaPago As String
    strFormaPago As String
    strFechaPago As String
    vlStrRFCPago As String
    vlStrClaveBancoSAT As String
    vlstrClaveCuentaBancaria As String
    vlStrTipoCargoBancario As String
    vlstrCuenta As String
    vlstrtipopago As String
End Type

Public Type typTarifaImpuesto
    lngId As Long
    dblPorcentaje As Double
End Type
Public Type DatosFiscales
    strDomicilio As String
    strNumExterior As String
    strNumInterior As String
    strTelefono As String
    lstrCalleNumero As String
    lstrColonia As String
    lstrCiudad As String
    lstrEstado As String
    lstrCodigo As String
    llngCveCiudad As Long
End Type
Public Type TipoPoliza
    lngnumCuenta As Long
    dblCantidad As Double
    intNaturaleza As Integer
End Type

Public Type RegistroPoliza
    vllngNumeroCuenta As Long
    vldblCantidadMovimiento As Double
    vlintTipoMovimiento As Integer
End Type

'Constantes
Public Const cgstrModulo As String = "PV" 'Constante para registrar el nombre del módulo en manejo de errores

'Números de opción del proceso de transferencias
Public Const cintNumOpcionFondoFijo = 312
Public Const cintNumOpcionTransBanco = 356
Public Const cintNumOpcionTransDepto = 344
Public Const cintNumOpcionRecepcion = 364
Public Const cintNumOpcionFormasPago = 320
Public Const cintNumOpcionConceptoCaja = 368
Public Const cintNumOpcionCorte = 311
Public Const cintIdTipoFactura = 2
Public Const cintNumOpcionListas = 307
Public Const cintNumOpcionDeptosCajaChica = 355
Public Const cintNumOpcionCambioformas = 2176

'Este tipo de dato sirve para identificar en que proceso estas
Public Enum enmTipoProceso
        Cargos = 1
        pagos = 2
        Facturacion = 3
        Honorarios = 4
        TrasladoCargos = 5
        AbrirCerrarCuentas = 6
        Exclusion = 7
        AsignarPaquetes = 8
        Presupuesto = 9
        Descuento = 10
End Enum

'Banderas
Public vgblnExistioError As Boolean 'Bandera que permite salir del formulario
Public vgblnAuxNotas As Boolean 'Auxiliar para la inicialización de pantalla de notas


Public vgblnErrorIngreso As Boolean 'Bandera para verificar si existe un error al ingreso de datos como no se ha escrito nada en una variable

Public vgblnFlagGrd As Boolean

'Variables
Public vgintColOrd As Integer 'Contiene el número de la columna a ordenar
Public vgintTipoOrd As Integer 'Contiene el tipo de ordenación

Public vgstrNombreEmpresaContable As String 'Nombre empresa
Public vgstrNombreCortoEmpresaContable As String 'Nombre corto empresa
Public vgstrRepresentanteLegalEmpresaContable As String 'Representante legal
Public vgstrRFCEmpresaContable As String 'RFC
Public vgstrDireccionEmpresaContable As String 'Direccion
Public vgstrTelefonoEmpresaContable As String 'Telefono
Public vglngCuentaResultadoEjercicio As Long 'Cuenta para mostrar el resultado del ejercicio
Public vgintEjercicioInicioOperaciones As Integer 'Ejercicio en el cual se comienzan operaciones
Public vgintMesInicioOperaciones As Integer 'Mes en el cual se comienzan operaciones

' variables para cambiar mediante drag & drop el orden de las columnas del grid
Public vgstrNombreForm As String 'Indica el nombre del formulario en el que ocurre el error
Public vgstrNombreProcedimiento As String

Public vgintColLoc As Integer 'Variable que contiene la columna que va a ordenar al buscar dentro del MshFlexGrid

Public vgstrAcumTextoBusqueda As String 'Acumula los caracteres capturados para la búsqueda en el grdHBusqueda

'Variables globales para trabajar con cualquier formulario, procedimiento, o donde se lo aplique como parametros de entrada o salida
Public vgstrVarIntercam As String 'Variable de intercambio de información entre procedimientos
Public vgstrVarIntercam2 As String 'Variable de intercambio de información entre procedimientos

Public vllngCuentaHonorarioPagar As Long       'Cuenta puente honorario pagar
Public vllngCuentaHonorarioCobrar As Long      'Cuenta puente honorario cobrar

'Parametros Generales
Public vgstrNombreHospitalCH As String * 70 'Variable que contiene el nombre del hospital
Public vgstrNombCortoCH As String * 10 'Nombre corto del centro hospitalario
Public vgstrRfCCH As String * 15 'Rfc del centro hospitalario
Public vgstrIMSSCH As String * 20 'Registro IMSS del centro hospitalario
Public vgstrSSACH As String * 20 'Licencia SSA del centro hospitalario
Public vgstrRepLegalCH As String * 50 'Nombre del respresentante legal del centro hospitalario
Public vgstrDirGnralCH As String * 50 'Nombre del director general del centro hospitalario
Public vgstrDirMedCH As String * 50 'Nombre del director medico del centro hospitalario
Public vgstrAdmGnralCH As String * 50 'Nombre del administrador general del centro hospitalario
Public vgstrDireccionCH As String * 60 'Direccion del centro hospitalario
Public vgstrColoniaCH As String * 25 'Colonia del centro hospitalario
Public vgintCveCiudadCH As Integer 'Clave de la ciudad donde se encuentra el centro hospitalario
Public vgstrCiudadCH As String 'Ciudad donde se encuentra el centro hospitalario
Public vgintCveEstadoCH As Integer 'Clave del estado donde se encuentra el centro hospitalario
Public vgstrEstadoCH As String 'Estado donde se encuentra el centro hospitalario
Public vgintCvePaisCH As Integer 'Clave del pais donde se encuentra el centro hospitalario
Public vgstrPaisCH As String 'Pais donde se encuentra el centro hospitalario
Public vgstrTelefonoCH As String * 10 'Telefono del centro hospitalario
Public vgimgLogoCH As Variant 'Para el logo del hospital
Public vgstrFaxCH As String * 10 'Fax del centro hospitalario
Public vgstrEmailCH As String * 25 'Email del centro hospitalario
Public vgstrWebCH As String * 30 'Pagina Web del centro hospitalario
Public vgstrCodPostalCH As String * 8 'Codigo postal del centro hospitalario
Public vgstrApartPostalCH As String * 8 'Apartado postal del del centro hospitalario
Public vgintTipoCuartoCH As Integer 'Clave del Tipo de cuarto predeterminado para mostrarlo en la admision
Public vgintEstadoSaludCH As Integer 'Clave del Estado de Salud que por omision debe ser ingresado por admision
Public vgbytDisponibleCH As Byte 'Codigo del Estado de cuarto disponible
Public vgbytOcupadoCH As Byte 'Codigo del Estado de cuarto ocupado
Public vglngCveNacionalidad As Long             'Clave de la nacionalidad predeterminada para catálogos de pacientes
Public vgstrNacionalidad As String              'Descripción de la nacionalidad predeterminada para catálogos de pacientes
Public vgintDiasRequisicion As Integer          'Número de dias anteriores sobre los cuales se revisaran las requisiciones


'----------------------------------------------------------------
' caso 12619 envio automatico de comprobante de nomina al timbrar
Public vgblnAutomatico As Boolean
Public vgblnEnvioExitosoCorreo As Boolean
'----------------------------------------------------------------

'Parametros de almacén
Public Type departamento
    vlintDepartamento As Integer
End Type

Public Type MenuPermiso
    vlintNumMenu As Integer
    vlstrNombreMenu As String
    vlstrNombreBoton As String
    vlstrPermisos As String
End Type

Public vglngCveAlmacenGeneral As Long 'Indica el numero del proveedor asignado como Almacen General o Abastecimiento

'Expediente
Public vgstrTipoConvenioPac As String       'Descripcion del convenio del paciente
Public vgstrNombreEmpresa As String         'Numero de empresa del paciente
Public vgstrTipoPacienteAdm As String       'Descripcion del tipo de paciente Convenio, particular, accionista,..
Public vgstrNomArea As String               'Para cargar el nombre del area
Public vgstrNombrePaciente As String        'Nombre del paciente
Public vgstrNumeroAfiliacion As String      'Numero de afiliacion del paciente
Public vgstrNumeroCuarto As String          'Número Cuarto
Public vgstrEdad As String                  'Edad del paciente
Public vgstrTipoPaciente As String          'Tipo de paciente <I>Interno  <E>Externo
Public vglngNumeroCuenta As Long            'Número de cuenta (Número de movimiento)de paciente, sea I o E
Public vglngNumeroPaciente As Long          'Número de de paciente (Número de expediente)
Public vgdtmFechaIngreso As Date            'Fecha de Ingreso
Public vgstrNombreMedico As String          'Nombre del médico responsable del paciente
Public vgstrSexoPaciente As String          'Sexo del paciente "M"asculino o "F"emenino

'Para trabajar con el formulario frmLista
Public vgstrDataMember As String 'Para saber de que objeto comando traer y revisar la informacion

'Para trabajar con el formulario frmBusquedaPaciente con paso de parametros
Public vgstrTipoIngreso As String * 2 'Para saber que tipo de ingreso escogio el usuario
Public vgdblNumExpediente  As Double 'Para saber el numero del expediente el paciente escogido por el usuario
Public vgdblNumCuenta As Double 'Para saber el numero de cuenta del paciente escogido por el usuario
Public vgblnSeEscogio As Boolean 'Para saber si el usuario escogio un registro(al no escoger ninguno se crea un nuevo)

Public vgblnContrasenaCorr As Boolean 'Bandera para verificar si la contrasena tecleada fue la correcta

Public vgstrNombreCbo As String

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' S E G U R I D A D
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public vgintNumeroModulo As Integer  'Número de módulo de la tabla Modulo
Public vglngNumeroEmpleado As Long  'Número de empleado logueado
Public vgintNumeroDepartamento As Integer  'Numero de departamento con que se logueo el empleado
Public vglngNumeroLogin As Long  'Número de login que le corresponde en la tabla Login
Public vgstrNombreUsuario As String  'Login del usuario
Public vgstrNombreDepartamento As String  'Nombre del departamento personalizado con que se loguea el empleado
Public aPermisos() As Permiso  'Arreglo para seguridad

Public Type Permiso
    vllngNumeroOpcion As Long
    vlstrTipoPermiso As String
End Type

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' G E N E R A L E S   D E L   M O D U L O
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public Type FormasPago
    vlintNumFormaPago As Integer
    vldblCantidad As Double
    vllngFolio As Double
    vlstrFolio As String
    vldblComision As Double
    vllngCuentaContable As Long
    vldblTipoCambio As Double
    vlbolEsCredito As Boolean
    vldblDolares As Double
    lngIdBanco As Long
    intMoneda As Long
    vllngCuentaComisionBancaria As Long
    vldblCantidadComisionBancaria As Double
    vldblIvaComisionBancaria As Double
    
    vlstrRFC As String
    vlstrBancoSAT As String
    vlstrBancoExtranjero As String
    vlstrCuentaBancaria As String
    vldtmFecha As Date
End Type

Public vgintClaveEmpresaContable As Integer  'Clave de empresa registrada en Parametros
Public vgstrEstructuraCuentaContable As String  'Estructura de la cuenta contable registrada en Parametros
Public vgdblCantidadIvaGeneral As Double  'Porcentaje IVA general registrada en Parametros
Public vglngNumeroTipoFormato As Long 'Número de formato para impresión de cheque

'--------------------------------------------------------------
'06/Febrero/2003 Con la actualización del catálogo externos y tarjetas de cme
'--------------------------------------------------------------
Public vgOrden As Byte
Public vlY As Long              'Variable que contiene el numero de renglon actual en la impresion
Public vYCol As Long            'Almacena el renglon para inicializar la impresion en caso de que sean mas de una columna
Public vSigueCol As Boolean     'Valida si ya termino de imprimir la segunda columna

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' Variables del modulo de Nomina
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public vgintNumIMSS As Integer   '
Public vgintDiasVaca As Integer
Public vgintDiasAusen As Integer
Public vgintNumPosISPT As Integer   'Numero de ispt deduccion
Public vgintNumNegISPT As Integer   'Numero de ispt percepcion
Public vgdbldiasMes As Double
Public vgdblSubsidio As Double
Public vgdbldiasAno As Double
Public vgintTopeHoras As Integer
Public vgdblTasaEMe As Double
Public vgdblTasaEMd As Double
Public vgdblTasaPens As Double
Public vgdblTasaIV As Double
Public vgdblTasaCV As Double
Public vgintLimEMe As Integer
Public vgintLimEMd As Integer
Public vgintLimPens As Integer
Public vgintLimIV As Integer
Public vgintLimCV As Integer
Public vgintCvePeriodo As Integer
Public vgintNumPeriodoMes As Integer
Public vgintCveVacaciones As Integer
Public vgintcvePrimaVaca As Integer
Public vgintCveDiasLabora As Integer
Public vgintCveFondoAhorro As Integer
Public vgdtmFechaIniAhorro As Date
Public vgdtmFechaFinAhorro As Date
Public vgdblLimComedor As Double
Public vgintCveComedor As Integer
Public vgintCveAguinaldo As Integer
Public vgintCveExedenteFA As Integer
Public vgintCveSueldoNormal As Integer
Public vgintCveBonoAsistencia As Integer
Public vgintCveDesctoComedor As Integer
Public vgintCveFaltasMater As Integer
Public vgintCvePTU As Integer
Public vgblnStaConsPTU As Boolean
Public vgdblPorPTU As Double
Public vgstrPeriodoIni As String
Public vgintCvePrestamoAhorro As Integer
Public vgintCvePremioPuntualidad As Integer
'    --29/08/2001--
Public vglngCveIndemniza As Long
Public vglngCvePrimaAntig As Long
Public vgDblDiasPrimaAntig As Double
Public vgDblDiasIndemniza As Double
Public vgintCveEFAneg As Integer
Public vgintCveRetardo As Integer
Public vgintCveBanco As Integer
Public vgintCveCreditoEmpleado As Integer
Public vgintCveIMSSPagado As Integer

Public Const cgintIntentoBloqueoCuenta As Integer = 100

'--------------------------------------------------------------
' Tipo de Datos para manejar el arreglo de las urgencias
'--------------------------------------------------------------
Public Type typUrgencias
    intConsecutivoUrgencia As Integer
    strTipoPaciente As String
    dblCantidadUrgencia As Double
End Type

Public vgstrGrabaMedicoEmpleado As String       'Indica si grabó un empleado o médico
Public vglngPersonaGraba As Long                'Número de empleado o médico que grabó

Public vgintCveDepartamento As Integer  'para frmRequisicionCargoPac
Public vlstrRegimenFiscal As String

Public vgintFolioUnico As Long



'Boletines
Public Type DatosBoletines
    lote As String
    Url As String
    Descripcion As String
    FechaCaduca As Date
End Type
Public aBoletines() As DatosBoletines

'&&&&&&VARIABLES PARA MOVIMIENTOS DE CRÉDITOS PARA FACTURAR, CASO 18330&&&&&&&&&&&&&&&&
Public allngAgregarCreditos() As Long  'Arreglo donde se guardan los créditos a facturar agregados a la factura
Public vlintAgregarCreditos As Integer 'Dimensión de allngAgregarCreditos
Public lblnCreditosaFacturar As Boolean 'Variable para saber si se agregaron créditos para facturar
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

'Devoluciones de Almacenes
Public vgstrFechaDevolucion As Date
Public vgintNumDev As Long
Public vgintNumRef As Long
Public vgstrDeptoDevuelve As String
Public vgstrDeptoRecibe As String
Public vgstrEmpleadoRecibe As String

'Banderas
Public vlblnExistioError As Boolean 'Bandera que permite salir del formulario
Public vgblnHuboCambio As Boolean 'Bandera para controlar si existe algun cambio y confirmar el mismo
Public vgstrEmpleado As String 'Guarda el nombre del encargado del departamento de Administración

Public vgintColArrastra As Integer 'Variable para identificar que columna se esta desplazando en
Public vgintMousePosXdn As Integer, vgintMousePosYdn As Integer 'Variables que contienen la posicion de X y Y cuando se utiliza el ratón en el grid
Public vgblnArrastrarOk As Boolean 'Variable para identificar si se ha realizado un drag & drop
Public vgstrCveArtRecep As String  ' Es el arículo que se recibe

'Public Function flngNumeroCorte(intCveDepto As Integer, lngCveEmpleado As Long, strTipoCorte As String) As Long
'    On Error GoTo NotificaError
'    '----------------------------------------------------------------------------------------
'    ' Regresa el número de corte abierto en este momento
'    '----------------------------------------------------------------------------------------
'    'lngCveDepto = Clave del departamento
'    'lngCveEmpleado = Clave del empleado
'    'strTipoCorte = Tipo de corte, P = caja de ingresos, C = caja chica
'
'    vgstrParametrosSP = CStr(intCveDepto) & "|" & CStr(lngCveEmpleado) & "|" & Trim(strTipoCorte)
'
'    flngNumeroCorte = 1
'    frsEjecuta_SP vgstrParametrosSP, "SP_GNSELNUMEROCORTE", True, flngNumeroCorte
'
'Exit Function
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngNumeroCorte"))
'End Function

'Public Function flngCorteValido(intCveDepto As Integer, lngCveEmpleado As Long, chrTipoCorte As String) As Long
''----------------------------------------------------------------------------------------
'' Revisar si el corte actual es válido o debe cerrarse
''----------------------------------------------------------------------------------------
'    On Error GoTo NotificaError
'
'    flngCorteValido = 1
'    frsEjecuta_SP CStr(intCveDepto) & "|" & CStr(lngCveEmpleado) & "|" & Trim(chrTipoCorte), "SP_PVSELCORTEVALIDO", True, flngCorteValido
'
'Exit Function
'NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCorteValido"))
'End Function

Public Function fstrFolioDocumento(vlintxDepartamento As Integer, vlstrxTipoDocto As String, vlblnxAviso As Boolean) As String
      '----------------------------------------------------------------------------------------
      'Función para regresar el siguiente folio a usar de un documento x,
      ' "RE" Recibos
      ' "FA" Facturas
      ' "NC" Nota de crédito
      ' "NA" Nota de cargo
      '----------------------------------------------------------------------------------------
1         On Error GoTo NotificaError
          
          Dim vlstrSentencia As String
          Dim rsRegistroFolio As New ADODB.Recordset
          
2         fstrFolioDocumento = ""
          
3         vlstrSentencia = vlstrSentencia + " select * from RegistroFolio where intNumeroActual<=intNumeroFinal "
4         vlstrSentencia = vlstrSentencia + "and smiDepartamento=" + Str(vlintxDepartamento) + " "
5         vlstrSentencia = vlstrSentencia + "and chrTipoDocumento=" + "'" + vlstrxTipoDocto + "'"
          
6         Set rsRegistroFolio = frsRegresaRs(vlstrSentencia)
7         If rsRegistroFolio.RecordCount = 0 Then
              'No existen folios activos para este documento.
8             MsgBox SIHOMsg(291), vbOKOnly + vbInformation, "Mensaje"
9         Else
10            If vlblnxAviso And (rsRegistroFolio!intNumeroFinal - rsRegistroFolio!intNumeroActual + 1) <= rsRegistroFolio!smiFoliosAviso Then
11                MsgBox "Faltan " + Trim(Str(rsRegistroFolio!intNumeroFinal - rsRegistroFolio!intNumeroActual + 1)) + " " & _
                  IIf(vlstrxTipoDocto = "RE", "recibos ", IIf(vlstrxTipoDocto = "FA", "facturas ", IIf(vlstrxTipoDocto = "NC", "notas de crédito", "notas de cargo"))) & _
                  "y será necesario actualizar folios!", vbOKOnly + vbInformation, "Mensaje"
12            End If
13            fstrFolioDocumento = Trim(rsRegistroFolio!chrCveDocumento) + Trim(Str(rsRegistroFolio!intNumeroActual))
14        End If

15    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrFolioDocumento" & " Linea:" & Erl()))
End Function

Public Function fblnFormasPagoPos(axFormasPago() As FormasPago, vldblCantidad As Double, vlblnPesos As Boolean, vldblxTipoCambio As Double, vlblnxIncluirFormasdeCredito As Boolean, vlLngReferencia As Long, vlstrTipoCliente As String, Optional vlstrRFC As String, Optional pblnPermiteSalir As Boolean = False, Optional pblnPermitePagarMasQueLaDeuda As Boolean = False, Optional blnFormaTrans As Boolean, Optional strforma As String, Optional intDirectaMasiva As Integer = 0, Optional strReferenciaPago As String, Optional strFormaPago As String, Optional strFechaPago As String, Optional vlStrRFCPago As String, Optional vlstrCuenta As String, Optional tipoPago As String, Optional vlstrCveBancoSAT As String, Optional vlstrClaveCuentaBancaria As String, Optional vlStrTipoCargoBancario As String) As Boolean
'----------------------------------------------------------------------------------------
' Función que llama la forma frmPagoPOS para registrar las formas de pago en las cuales
' se paga una cantidad x, regresa un falso si no se registro nada o si no existen formas
' de pago
' axFormasPago()                = arreglo en donde se dejaran las formas de pago
' vldblCantidad                 = cantidad que se tiene que cubrir
' vlblnPesos                    = moneda
' vldblxTipoCambio              = tipo de cambio del día
' vlblnxIncluirFormasdeCredito  = bandera para incluir o no formas de pago a credito
' vllngReferencia               = si es que se incluyen formas a credito, numero de referencia del cliente
' vlstrTipoCliente              = si es que se incluyen formas a credito, tipo de cliente
'                                   PI = Paciente interno
'                                   PE = Paciente externo
'                                   EM = Empleado
'                                   ME = Médico
'                                   CO = Empresa
' intDirectaMasiva              = Si es facturación directa (Todo normal), o Masiva (los campos se rellenan automáticamente)
'----------------------------------------------------------------------------------------
    Dim vlstrsql As String
    Dim X As Integer
    
    fblnFormasPagoPos = False
    With frmPagoPos
        .lblnFormaTrans = blnFormaTrans
        .vldblCantidadPago = vldblCantidad
        .vlblnIncluirFormasCredito = vlblnxIncluirFormasdeCredito
        .vlLngReferencia = vlLngReferencia
        .vlstrTipoCliente = vlstrTipoCliente
        .vgstrForma = strforma
        .vlblnPesos = vlblnPesos
        .vldblTipoCambioDia = vldblxTipoCambio
        .txtCantidad.Text = Format(Str(vldblCantidad), "###,###,###,###.00")
        .txtImporte.Text = Format(Str(vldblCantidad), "###,###,###,###.00")
        .txtDiferencia.Text = Format(Str(vldblCantidad), "###,###,###,###.00")
        .vgblnPermiteSalir = pblnPermitePagarMasQueLaDeuda 'pblnPermiteSalir
        '.vgstrForma = strforma
        .vlstrRFCOriginal = vlstrRFC
        If intDirectaMasiva = 0 Then
            .Show vbModal
        Else
            .intModoMasivo = 1
            Call .Form_Activate
            Call .txtBancoExtranjero_KeyPress(13)
            .pSeleccionaBancoSAT vlstrCveBancoSAT, tipoPago
            .pSeleccionaPago strFormaPago
            .pConfiguraMasiva vlStrRFCPago, strReferenciaPago, strFechaPago, vlstrCuenta, tipoPago
            .pSeleccionaCuentaBancaria vlstrClaveCuentaBancaria, tipoPago
            .pSeleccionaTipoCargo vlStrTipoCargoBancario, tipoPago
            Call .txtCantidad_GotFocus
            Call .txtCantidad_KeyPress(13)
            Call .cmdAceptar_Click
        End If
    
        If .vgintSalidaOK = 1 Then
            ReDim axFormasPago(0)
            For X = 1 To .grdFormas.Rows - 1
                ReDim Preserve axFormasPago(X - 1)
                axFormasPago(X - 1).vldblCantidad = Val(.grdFormas.TextMatrix(X, 6))
                axFormasPago(X - 1).vlintNumFormaPago = .grdFormas.RowData(X)
                axFormasPago(X - 1).vllngFolio = Val(.grdFormas.TextMatrix(X, 2))
                axFormasPago(X - 1).vlstrFolio = Trim(.grdFormas.TextMatrix(X, 2))
                axFormasPago(X - 1).vllngCuentaContable = Val(.grdFormas.TextMatrix(X, 4))
                axFormasPago(X - 1).vldblTipoCambio = Val(.grdFormas.TextMatrix(X, 5))
                axFormasPago(X - 1).vlbolEsCredito = CBool(.grdFormas.TextMatrix(X, 7))
                axFormasPago(X - 1).vldblDolares = Val(.grdFormas.TextMatrix(X, 8))
                axFormasPago(X - 1).lngIdBanco = Val(.grdFormas.TextMatrix(X, 9))
                axFormasPago(X - 1).intMoneda = Val(.grdFormas.TextMatrix(X, 10))
                axFormasPago(X - 1).vllngCuentaComisionBancaria = Val(.grdFormas.TextMatrix(X, 11))
                axFormasPago(X - 1).vldblCantidadComisionBancaria = Val(.grdFormas.TextMatrix(X, 12))
                axFormasPago(X - 1).vldblIvaComisionBancaria = Val(.grdFormas.TextMatrix(X, 13))
                
                axFormasPago(X - 1).vlstrRFC = Trim(.grdFormas.TextMatrix(X, 14))
                axFormasPago(X - 1).vlstrBancoSAT = Trim(.grdFormas.TextMatrix(X, 15))
                axFormasPago(X - 1).vlstrBancoExtranjero = Trim(.grdFormas.TextMatrix(X, 16))
                axFormasPago(X - 1).vlstrCuentaBancaria = Trim(.grdFormas.TextMatrix(X, 17))
                
                If Trim(.grdFormas.TextMatrix(X, 18)) <> "" Then
                    axFormasPago(X - 1).vldtmFecha = CDate(Trim(.grdFormas.TextMatrix(X, 18)))
                End If
            Next X
            fblnFormasPagoPos = True
        End If
    End With
    
    Unload frmPagoPos
End Function

Private Function fblnEsValida(strData As String, strCensables As String, strNoCensables As String) As Boolean
    Dim intPosC As Integer
    Dim intPosN As Integer
    intPosC = InStr(1, strData, "C")
    intPosN = InStr(1, strData, "N")
    If intPosC = 0 Then
        fblnEsValida = False
        Exit Function
    Else
        If intPosN = 0 Or intPosN <> Len(strData) Then
            fblnEsValida = False
            Exit Function
        Else
            strCensables = Mid(strData, 1, InStr(1, strData, "C") - 1)
            strNoCensables = Mid(strData, InStr(1, strData, "C") + 1, InStr(1, strData, "N") - InStr(1, strData, "C") - 1)
            If Not IsNumeric(strCensables) Then
                If strCensables <> "UL" Then
                    fblnEsValida = False
                    Exit Function
                End If
            End If
            If Not IsNumeric(strNoCensables) Then
                If strNoCensables <> "UL" Then
                    fblnEsValida = False
                    Exit Function
                End If
            End If
        End If
    End If
    fblnEsValida = True
End Function

Public Function Space(vlintNumeroDeEspacios As Integer) As String
    Dim vlIntCont As Integer
    
    Space = ""
    For vlIntCont = 1 To vlintNumeroDeEspacios
        Space = Space & " "
    Next vlIntCont
End Function

Public Sub pLimpiaVSFlexGrid(ObjGrid As VSFlexGrid)
'-------------------------------------------------------------------------------------------
' Limpia o Inicia completamente un Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    With ObjGrid
        .Clear
        .Rows = 0
        .Cols = 0
        .FixedCols = 0
        .FixedRows = 0
    End With
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaMshFGrid"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Public Sub pBorrarRegVSFlexGrid(vlintNumeroRenglonBorrar As Integer, grdNombre As VSFlexGrid, Optional vlblnNolimpia As Boolean)
'--------------------------------------
'Funcion borrar un registro del grid sin que se pierda el rowdata
'--------------------------------------
    Dim acontenido() As Variant
    Dim X As Integer
    Dim z As Integer
    Dim Y As Integer
    Dim vlintColumnasArreglo As Integer
    Dim vlintRenglonesArreglo As Integer

    If grdNombre.Rows = 2 Then
        If vlblnNolimpia Then
            For vlintColumnasArreglo = 1 To grdNombre.Cols - 1
                grdNombre.TextMatrix(1, vlintColumnasArreglo) = ""
            Next
            grdNombre.RowData(1) = -1
        Else
            grdNombre.Rows = 0
        End If
    Else
        vlintRenglonesArreglo = grdNombre.Rows - 2
        vlintColumnasArreglo = grdNombre.Cols
        
        ReDim acontenido(vlintRenglonesArreglo, vlintColumnasArreglo)
        
        z = 0
        For X = 1 To grdNombre.Rows - 1
            If X <> vlintNumeroRenglonBorrar Then
                For Y = 1 To grdNombre.Cols
                    If Y = grdNombre.Cols - 1 Then
                        acontenido(z, Y - 1) = grdNombre.RowData(X)
                    Else
                        acontenido(z, Y - 1) = grdNombre.TextMatrix(X, Y)
                    End If
                Next Y
                z = z + 1
            End If
        Next X
    
        grdNombre.Rows = grdNombre.Rows - 1
    
        For X = 0 To vlintRenglonesArreglo - 1
            For Y = 0 To vlintColumnasArreglo - 1
                If Y + 1 = vlintColumnasArreglo Then
                    grdNombre.RowData(X + 1) = acontenido(X, Y)
                Else
                    grdNombre.TextMatrix(X + 1, Y + 1) = acontenido(X, Y)
                End If
            Next Y
        Next X
    End If
End Sub
Public Sub pCancelaPolizaCredito(lngNumPoliza As Long, lngEmpleado As Long, Optional strNumeroPago As String, Optional strDescripcion As String, Optional intAnno As Integer, Optional intMes As Integer, Optional dtmFechaPago As Date)
    '------------------------------------------------------------------------------------------------------------
    ' Procedimiento que cancela una póliza, deja el registro maestro en CnPoliza
    ' e inserta una nueva póliza de reversa con la fecha actual y descripción alusiva al número de pago cancelado,
    ' no se elimina de CnPoliza porque los consecutivos del tipo de póliza se verían afectados
    '------------------------------------------------------------------------------------------------------------

    Dim rsCnPoliza As New ADODB.Recordset
    Dim strDescripcionPoliza As String 'Para almacenar la descripción de la poliza de reversa
    Dim vlstrsql As String 'Para almacenar el query a la tabla cnpoliza
    Dim vllngNumeroPoliza As Long 'Para almacenar el número (id) de la nueva póliza de reversa

    If strDescripcion <> "" Then
        'Si es un pago
        If strNumeroPago <> "" Then
            strDescripcionPoliza = strDescripcion & " " & strNumeroPago
        Else
            strDescripcionPoliza = strDescripcion
        End If
    End If

    vlstrsql = "select * from CnPoliza where intNumeroPoliza = -1"
    Set rsCnPoliza = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)

    With rsCnPoliza
        .AddNew
        !tnyClaveEmpresa = vgintClaveEmpresaContable
        !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, "D", intAnno, intMes, False)
        !smiEjercicio = intAnno
        !tnyMes = intMes
        !dtmFechaPoliza = dtmFechaPago
        !vchConceptoPoliza = strDescripcionPoliza
        !chrTipoPoliza = "D"
        !SMICVEDEPARTAMENTO = vgintNumeroDepartamento
        !intCveEmpleado = lngEmpleado
        .Update
    End With

    vllngNumeroPoliza = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!INTNUMEROPOLIZA)

    vgstrParametrosSP = lngNumPoliza & "|" & vllngNumeroPoliza

    frsEjecuta_SP vgstrParametrosSP, "SP_CCINSERTADETALLEPOLIZA", True

    rsCnPoliza.Close

End Sub
Public Sub pCancelaPolizaCreditoFechaActual(lngNumPoliza As Long, lngEmpleado As Long, Optional strNumeroPago As String, Optional strDescripcion As String)
    '------------------------------------------------------------------------------------------------------------
    ' Procedimiento que cancela una póliza, deja el registro maestro en CnPoliza
    ' e inserta una nueva póliza de reversa con la fecha actual y descripción alusiva al número de pago cancelado,
    ' no se elimina de CnPoliza porque los consecutivos del tipo de póliza se verían afectados
    '------------------------------------------------------------------------------------------------------------
    
    Dim rsCnPoliza As New ADODB.Recordset
    Dim strDescripcionPoliza As String 'Para almacenar la descripción de la poliza de reversa
    Dim vlstrsql As String 'Para almacenar el query a la tabla cnpoliza
    Dim vllngNumeroPoliza As Long 'Para almacenar el número (id) de la nueva póliza de reversa
    
    If strDescripcion <> "" Then
        'Si es un pago
        If strNumeroPago <> "" Then
            strDescripcionPoliza = strDescripcion & " " & strNumeroPago
        Else
            strDescripcionPoliza = strDescripcion
        End If
    End If
    
    vlstrsql = "select * from CnPoliza where intNumeroPoliza = -1"
    Set rsCnPoliza = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    
    With rsCnPoliza
        .AddNew
        !tnyClaveEmpresa = vgintClaveEmpresaContable
        !intClavePoliza = flngFolioPoliza(vgintClaveEmpresaContable, "D", Year(fdtmServerFecha), Month(fdtmServerFecha), False)
        !smiEjercicio = Year(fdtmServerFecha)
        !tnyMes = Month(fdtmServerFecha)
        !dtmFechaPoliza = fdtmServerFecha
        !vchConceptoPoliza = strDescripcionPoliza
        !chrTipoPoliza = "D"
        !SMICVEDEPARTAMENTO = vgintNumeroDepartamento
        !intCveEmpleado = lngEmpleado
        .Update
    End With

    vllngNumeroPoliza = flngObtieneIdentity("SEC_CNPOLIZA", rsCnPoliza!INTNUMEROPOLIZA)
    
    vgstrParametrosSP = lngNumPoliza & "|" & vllngNumeroPoliza
    
    frsEjecuta_SP vgstrParametrosSP, "SP_CCINSERTADETALLEPOLIZA", True
    
    rsCnPoliza.Close
    
End Sub


Public Function fblnCuentaHonorarioCobrar() As Boolean
    On Error GoTo NotificaError
    Dim rsTemp As New ADODB.Recordset
    Dim SQL As String
    
    fblnCuentaHonorarioCobrar = False
    
    SQL = " SELECT  intnumcuentahonorariocobrar"
    SQL = SQL & " FROM CCPARAMETRO"
    Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
    If rsTemp.RecordCount > 0 Then
        If rsTemp!intnumcuentahonorariocobrar <> 0 Then
            vllngCuentaHonorarioCobrar = rsTemp!intnumcuentahonorariocobrar
            fblnCuentaHonorarioCobrar = True
        End If
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaHonorarioCobrar"))
End Function

Public Function fblnCuentaHonorarioPagar() As Boolean
    On Error GoTo NotificaError
    Dim rsTemp As New ADODB.Recordset
    Dim SQL As String
    
    fblnCuentaHonorarioPagar = False
    
    SQL = " SELECT  intnumcuentahonorarioPagar"
    SQL = SQL & " FROM CCPARAMETRO"
    Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
    If rsTemp.RecordCount > 0 Then
        If rsTemp!intnumcuentahonorarioPagar <> 0 Then
            vllngCuentaHonorarioPagar = rsTemp!intnumcuentahonorarioPagar
            fblnCuentaHonorarioPagar = True
        End If
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaHonorarioPagar"))
End Function

Public Function fintLocalizaItemDataCbo(pcboCombo As ComboBox, pintItemData As Integer) As Integer
    Dim vlIntCont As Integer
    
    fintLocalizaItemDataCbo = -1
    For vlIntCont = 0 To pcboCombo.ListCount - 1
        If pcboCombo.ItemData(vlIntCont) = pintItemData Then
            fintLocalizaItemDataCbo = vlIntCont
            Exit For
        End If
    Next vlIntCont

End Function

Public Sub pLimpiaMshFGd(ObjGrid As Control)
'-------------------------------------------------------------------------------------------
' Limpia o Inicia completamente un Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    With ObjGrid
        .Clear
        '.ClearStructure
        .FixedCols = 0
        .FixedRows = 0
        .Rows = 0
        .Cols = 0
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaMshFGrid"))
End Sub

Public Function fblnExisteCargoEnElGrid(pmsfgGrid As Control, pintCveCargo As Long) As Boolean
    Dim vlIntCont As Integer
    
    fblnExisteCargoEnElGrid = False
    For vlIntCont = 1 To pmsfgGrid.Rows - 1
        If pmsfgGrid.RowData(vlIntCont) = pintCveCargo Then
            fblnExisteCargoEnElGrid = True
            Exit For
        End If
    Next

End Function


Public Function flngBusquedaCuentasContables(Optional vlblnTodasCuentas As Boolean, Optional intclaveempresa As Integer, Optional vlintLvl As Integer) As Long
'----------------------------------------------------------------------------------------
' Realiza una consulta del catalogo de cuentas por descripcion y regresa el numero
' de cuenta seleccionada
' Es necesaria la forma frmBusquedaCuentasContables para funcionar
' Actualizaciones:
' 2001-05-16
'----------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    If IsNull(vlblnTodasCuentas) Then
        vlblnTodasCuentas = False
    End If
'// agregar un parametro de nivel a esta funcion para cambiar la variabel de nivel en form show
    frmBusquedaCuentasContables.vlblnTodasCuentas = vlblnTodasCuentas
    frmBusquedaCuentasContables.vlintcveempresa = intclaveempresa
    frmBusquedaCuentasContables.vlintLvl = vlintLvl
    frmBusquedaCuentasContables.Show vbModal
    flngBusquedaCuentasContables = frmBusquedaCuentasContables.vllngNumeroCuenta

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngBusquedaCuentasContables"))
End Function

Public Function fmsgNuevaSerieFolios(strTipoDoc As String, intNumDepto As Integer, ByRef strIdentificador As String, ByRef lngInicio As Long, ByRef lngFinal As Long) As VbMsgBoxResult
    fmsgNuevaSerieFolios = frmNuevaSerieFolios.fmsgPedirFolios(strTipoDoc, intNumDepto, strIdentificador, lngInicio, lngFinal)
End Function

Public Function fblnEstaEnListaNegra(lngCvePaciente As Long, lngCveTipoPaciente As Long, lngRelacion As Long, ByRef lngCLN As Long) As Boolean
    On Error GoTo NotificaError
    Dim lngCveListaNegra As Long
    
    lngCveListaNegra = 1
    frsEjecuta_SP lngCvePaciente & "|" & lngCveTipoPaciente & "|" & lngRelacion, "sp_ADVerificaListaNegra", True, lngCveListaNegra
    If lngCveListaNegra > 0 Then
        lngCLN = lngCveListaNegra
        If frmListaNegra.fblnAutorizacion(lngCveListaNegra) Then
            fblnEstaEnListaNegra = False
        Else
            fblnEstaEnListaNegra = True
        End If
        Unload frmListaNegra
    Else
        fblnEstaEnListaNegra = False
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnEstaEnListaNegra"))
End Function

Public Function fstrBloqueaCuenta(vllngxMovimiento As Long, vlstrxTipoPaciente As String) As String
      '----------------------------------------------------------------------------------------
      ' Bloquear la cuenta del paciente, el fin es evitar que sea facturada cuando se esta
      ' accesando
      ' Cuenta facturada  regresa 'F'
      ' Cuenta ocupada    regresa 'O'
      ' Cuenta libre      regresa 'L'
      '----------------------------------------------------------------------------------------
      'Modificación.
      'Fecha: 17-Jun-2004
      'Motivo: Se modificó para que si es un grupo se bloqueen todas cuentas que contenga dicho grupo
1     On Error GoTo NotificaError

          Dim rsEstatusCuenta As New ADODB.Recordset
          Dim vlstrSentencia As String
          'Variables necesarias para el grupo
          Dim vlintCuentas As Integer
          Dim vlrsCuentas As New ADODB.Recordset
          Dim vlblnEsGrupo As Boolean
          
2         vlintCuentas = 1
3         vlblnEsGrupo = False
4         If vlstrxTipoPaciente = "G" Then
5             Set vlrsCuentas = frsRegresaRs("SELECT DISTINCT intMovPaciente, chrTipoPaciente FROM PvCargo WHERE PvCargo.INTCVEGRUPO = " & vllngxMovimiento)
6             vlintCuentas = vlrsCuentas.RecordCount
7             vlblnEsGrupo = True
8         End If
          
9         For vlintCuentas = 1 To vlintCuentas
10            If vlblnEsGrupo Then 'Grupo de facturas
11                vllngxMovimiento = vlrsCuentas!INTMOVPACIENTE
12                vlstrxTipoPaciente = vlrsCuentas!CHRTIPOPACIENTE
13            End If
              
14            If vlstrxTipoPaciente = "I" Then 'Pacientes internos
15                vlstrSentencia = "select bitEstatusOcupado, bitFacturado from AdAdmision where numNumCuenta = " & Str(vllngxMovimiento)
16            Else 'Pacientes externos
17                vlstrSentencia = "select bitEstatusOcupado, bitFacturado from RegistroExterno where intNumCuenta = " & Str(vllngxMovimiento)
18            End If
19            Set rsEstatusCuenta = frsRegresaRs(vlstrSentencia)
20            If rsEstatusCuenta.RecordCount <> 0 Then
                  'If Not rsEstatusCuenta!BITFACTURADO Then
21                If rsEstatusCuenta!BITFACTURADO = 0 Then '(CR) - Modificado para caso 6863
                      'If Not rsEstatusCuenta!bitEstatusOcupado Then
22                    If rsEstatusCuenta!bitEstatusOcupado = 0 Then '(CR) - Modificado para caso 6863
23                        frsEjecuta_SP vlstrxTipoPaciente & "|" & Str(vllngxMovimiento) & "|1", "SP_EXUPDCUENTAOCUPADA"
24                        fstrBloqueaCuenta = "L"
25                    Else
26                        fstrBloqueaCuenta = "O"
27                    End If
28                Else
29                    fstrBloqueaCuenta = "F"
30                End If
31            Else
32                fstrBloqueaCuenta = "X"
33            End If
              'Si se trata de un grupo y alguna cuenta esta en un estado
              'diferente a "Libre" el grupo toma ese estado
34            If fstrBloqueaCuenta <> "L" Then Exit For
35            If vlblnEsGrupo Then vlrsCuentas.MoveNext
36        Next vlintCuentas

37    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fstrBloqueaCuenta" & " Linea:" & Erl()))
End Function

Public Function flngNumCliente(blnTodos As Boolean, IntActivos As Integer) As Long
    ' IntActivos = 0 Si se van a mostrar todos Activos e Inactivos, 1 Si seran solo Activos

    frmBusquedaCliente.lblnTodosClientes = blnTodos
    frmBusquedaCliente.lIntActivos = IntActivos
    frmBusquedaCliente.Show vbModal
    flngNumCliente = frmBusquedaCliente.vllngNumCliente

End Function

'- CASO 6217: Envío de CFDs por correo electrónico -'
Public Sub pEnviarCFD(strTipoDocumento As String, lngIdDocumento As Long, lngCveEmpresa As Long, strRFC As String, lngIdEmpleado As Long, frmEnvia As Form, Optional vgblnAutomatico As Boolean = False, Optional vlblnSinXMLConPDF As Boolean = False)
1     On Error GoTo NotificaError
          Dim rsCorreo As New ADODB.Recordset
          Dim rsDestinatario As New ADODB.Recordset
          Dim rsFolio As New ADODB.Recordset
          Dim rsRutaPDF As New ADODB.Recordset
          Dim rsRutaXML As New ADODB.Recordset
          Dim strSentencia As String
          
          '- Verifica configuración de la cuenta de correo de la empresa -'
2         Set rsCorreo = frsEjecuta_SP(CInt(lngCveEmpresa) & "|0", "Sp_CnSelCnCorreo")
3         If rsCorreo.RecordCount = 0 Then
              'No se ha configurado la cuenta de correo.
4             MsgBox SIHOMsg(1202), vbCritical, "Mensaje"
5             Exit Sub
6         End If
          
          '- Verifica el destinatario del correo con datos del consecutivo y el tipo del comprobante -'
7         Set rsDestinatario = frsEjecuta_SP(CStr(lngIdDocumento) & "|'" & strTipoDocumento & "'", "SP_GNSELEMAIL")
8         If rsDestinatario.RecordCount > 0 Then
9             frmDatosCorreo.strCorreoDestinatario = Trim(rsDestinatario!CORREO)
10        Else
11            frmDatosCorreo.strCorreoDestinatario = ""
12        End If
          
          '- Verifica el folio del documento a enviar -'
13        strSentencia = "SELECT TRIM(VCHSERIECOMPROBANTE) || TRIM(VCHFOLIOCOMPROBANTE) Folio FROM GNCOMPROBANTEFISCALDIGITAL WHERE intComprobante = " & lngIdDocumento & " AND trim(CHRTIPOCOMPROBANTE) = " & "'" & Trim(strTipoDocumento) & "'"
14        Set rsFolio = frsRegresaRs(strSentencia)
15        If rsFolio.RecordCount > 0 Then
16            frmDatosCorreo.strFolioDocumento = Trim(rsFolio!Folio)
17        Else
              'Error al procesar el folio del documento a enviar.
18            MsgBox SIHOMsg(1199), vbCritical, "Mensaje"
19            Exit Sub
20        End If
          
          '- Verifica las rutas de los archivos PDF -'
21        strSentencia = "SELECT trim(VCHRUTAPDF) RutaPDF FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & lngCveEmpresa
22        Set rsRutaPDF = frsRegresaRs(strSentencia)
23        If rsRutaPDF.RecordCount > 0 Then
24            frmDatosCorreo.strRutaPDF = Trim(rsRutaPDF!RutaPDF)
25        Else
              'No se ha configurado la ruta de los archivos PDF.
26            MsgBox SIHOMsg(1200), vbCritical, "Mensaje"
27            Exit Sub
28        End If

          '- Verifica las rutas de los archivos XML -'
29        strSentencia = "SELECT trim(VCHRUTAXML) RutaXML FROM CNEMPRESACONTABLE WHERE TNYCLAVEEMPRESA = " & lngCveEmpresa
30        Set rsRutaXML = frsRegresaRs(strSentencia)
31        If rsRutaXML.RecordCount > 0 Then
32            frmDatosCorreo.strRutaXML = Trim(rsRutaXML!RutaXML)
33        Else
              'No se ha configurado la ruta de los archivos XML.
34            MsgBox SIHOMsg(1201), vbCritical, "Mensaje"
35            Exit Sub
36        End If
          
37        frmDatosCorreo.strTipoDocumento = Trim(strTipoDocumento) 'Establece el valor del tipo de documento en la pantalla de envío
38        frmDatosCorreo.lngIdDocumento = lngIdDocumento  'Establece la clave del documento segun el valor del tipo en la pantalla de envío
39        frmDatosCorreo.lngEmpleado = lngIdEmpleado 'Se establece el valor del empleado que realiza el envío
          
40        If vlblnSinXMLConPDF Then
              'Para las prefacturas activa por default solo el PDF
41            frmDatosCorreo.chkXML.Value = 0
42            frmDatosCorreo.chkXML.Enabled = False
              
43            frmDatosCorreo.chkPDF.Value = 1
44            frmDatosCorreo.chkPDF.Enabled = True
45        End If
          
46        frmDatosCorreo.Show vbModal, frmEnvia 'Muestra la forma para envío de correo
          
47    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pEnviarCFD" & " Linea:" & Erl()))
End Sub

'Verifica el parámetro de envío de CFDs por correo
Public Function fblnRevisaEnvioCorreo(lintClaveEmpresa As Integer) As Boolean
On Error GoTo NotificaError
    
    Dim rsCorreo As New ADODB.Recordset
    Dim lstrSentencia As String
    
    lstrSentencia = "SELECT TRIM(SIPARAMETRO.VCHVALOR) AS VALOR FROM SIPARAMETRO " & _
                    "INNER JOIN CNEMPRESACONTABLE ON SIPARAMETRO.INTCVEEMPRESACONTABLE = CNEMPRESACONTABLE.TNYCLAVEEMPRESA " & _
                    "WHERE CNEMPRESACONTABLE.TNYCLAVEEMPRESA = " & lintClaveEmpresa & _
                    "AND SIPARAMETRO.VCHNOMBRE = 'BITENVIOCFD'"
    Set rsCorreo = frsRegresaRs(lstrSentencia)
    If rsCorreo.RecordCount > 0 Then
        fblnRevisaEnvioCorreo = Trim(rsCorreo!valor) = "1"
    Else
        fblnRevisaEnvioCorreo = False
    End If
    
Exit Function
NotificaError:
    fblnRevisaEnvioCorreo = False
End Function

Private Sub pLlenapoliza(lngnumCuenta As Long, dblCantidad As Double, intTipoMovto As Integer, ByRef apoliza() As TipoPoliza)
    Dim intTamaño As Integer
    Dim intPosicion As Integer
    Dim blnEstaCuenta As Boolean
    Dim intcontador As Integer
    
    If apoliza(0).lngnumCuenta = 0 Then
        apoliza(0).lngnumCuenta = lngnumCuenta
        apoliza(0).dblCantidad = dblCantidad
        apoliza(0).intNaturaleza = intTipoMovto
    Else
        blnEstaCuenta = False
        intTamaño = UBound(apoliza(), 1)
        For intcontador = 0 To intTamaño
            If apoliza(intcontador).lngnumCuenta = lngnumCuenta Then
                blnEstaCuenta = True
                intPosicion = intcontador
            End If
        Next intcontador
        
        If blnEstaCuenta Then
            If apoliza(intPosicion).intNaturaleza = intTipoMovto Then
                apoliza(intPosicion).dblCantidad = apoliza(intPosicion).dblCantidad + dblCantidad
            Else
                If apoliza(intPosicion).dblCantidad > dblCantidad Then
                    apoliza(intPosicion).dblCantidad = apoliza(intPosicion).dblCantidad - dblCantidad
                Else
                    If apoliza(intPosicion).dblCantidad < dblCantidad Then
                        apoliza(intPosicion).dblCantidad = dblCantidad - apoliza(intPosicion).dblCantidad
                        If apoliza(intPosicion).intNaturaleza = 1 Then
                            apoliza(intPosicion).intNaturaleza = 0
                        Else
                            apoliza(intPosicion).intNaturaleza = 1
                        End If
                    Else
                        apoliza(intPosicion).lngnumCuenta = 0
                        apoliza(intPosicion).dblCantidad = 0
                        apoliza(intPosicion).intNaturaleza = 0
                    End If
                End If
            End If
        Else
            ReDim Preserve apoliza(intTamaño + 1)
            apoliza(intTamaño + 1).lngnumCuenta = lngnumCuenta
            apoliza(intTamaño + 1).dblCantidad = dblCantidad
            apoliza(intTamaño + 1).intNaturaleza = intTipoMovto
        End If
    End If
End Sub

Private Function fstrTipoMovimientoForma(lintCveForma As Integer, ByRef frm As Form) As String
On Error GoTo NotificaError

    Dim rsForma As New ADODB.Recordset
    Dim lstrSentencia As String
    
    fstrTipoMovimientoForma = ""
    
    lstrSentencia = "SELECT * FROM PvFormaPago WHERE intFormaPago = " & lintCveForma
    Set rsForma = frsRegresaRs(lstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsForma.RecordCount > 0 Then
        Select Case rsForma!chrTipo
            Case "E": fstrTipoMovimientoForma = "EFC"
            Case "T": fstrTipoMovimientoForma = "TAC"
            Case "B": fstrTipoMovimientoForma = "TCL"
            Case "H": fstrTipoMovimientoForma = "CHC"
        End Select
    End If
    rsForma.Close
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (frm.Name & ":fstrTipoMovimientoForma"))
End Function

Public Function fintErrorBloqueoCorte(llngNumCorte As Long) As Integer
    '----------------------------------------------------------------------------------------------
    'Función para bloquear el corte
    '----------------------------------------------------------------------------------------------
    Dim lngCorteGrabando  As Long   'Resultado de la validación del estado del corte
    
    lngCorteGrabando = 1
    vgstrParametrosSP = llngNumCorte & "|" & "Grabando"
    frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, lngCorteGrabando
    If lngCorteGrabando <> 2 Then
        fintErrorBloqueoCorte = 720 'No se puede realizar la operación, inténtelo en unos minutos.
    End If
End Function

Public Function fintErrorContable(strFecha As String) As Integer
    '----------------------------------------------------------------------------------------------
    'Función para revisar que se pueda introducir una póliza o que el periodo contable esté abierto
    '----------------------------------------------------------------------------------------------
    Dim lngResultado As Long        'Resultado de la validación del periodo contable
    
    fintErrorContable = 0
    
    lngResultado = 1
    vgstrParametrosSP = vgintClaveEmpresaContable & "|" & "Grabando poliza"
    frsEjecuta_SP vgstrParametrosSP, "SP_CNUPDESTATUSCIERRE", True, lngResultado
    If lngResultado = 1 Then
        If fblnPeriodoCerrado(vgintClaveEmpresaContable, Year(CDate(strFecha)), Month(CDate(strFecha))) Then
            fintErrorContable = 209 'El periodo contable esta cerrado.
            Exit Function
        End If
    Else
        fintErrorContable = 720 'No se puede realizar la operación, inténtelo en unos minutos.
        Exit Function
    End If
End Function

Public Function fstrNoLetras(ByVal sString As String) As String
    Dim i As Integer
    For i = 1 To Len(sString)
        If IsNumeric(Mid(sString, i, 1)) Then
            fstrNoLetras = fstrNoLetras & Mid(sString, i, 1)
        End If
    Next i
End Function

Private Sub pGuardaPolizaCreditosaFacturar(llngNumPoliza As Long, vlStrFolioFactura As String)
On Error GoTo NotificaError
    
    'Dim vlintContador As Integer
    'Dim lngClavePoliza As Long
    'Dim intIndex As Integer
    
    Dim rsCnDetallePolizaCF As New ADODB.Recordset
    Dim rsCCMovCreditoPolizaCF As New ADODB.Recordset
    Dim rsCCDatosPolizaCF As New ADODB.Recordset
    Dim vlIntCont As Integer
    Dim vlstrsql As String
    
    For vlIntCont = 1 To vlintAgregarCreditos
        vlstrsql = "select * from CnDetallePoliza where intnumeropoliza in (select intnumeropoliza from CCMovimientoCreditoPoliza where intnummovimiento = " + CStr(allngAgregarCreditos(vlIntCont)) + ")"
        Set rsCCDatosPolizaCF = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        
        vlstrsql = "select * from CnDetallePoliza where intNumeroPoliza=-1"
        Set rsCnDetallePolizaCF = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
        
        With rsCnDetallePolizaCF
            Do While Not rsCCDatosPolizaCF.EOF
                .AddNew
                !INTNUMEROPOLIZA = llngNumPoliza
                !intNumeroCuenta = rsCCDatosPolizaCF!intNumeroCuenta
                !bitNaturalezaMovimiento = IIf(rsCCDatosPolizaCF!bitNaturalezaMovimiento = 0, 1, 0)
                !mnyCantidadMovimiento = rsCCDatosPolizaCF!mnyCantidadMovimiento
                !vchReferencia = " "
                !vchConcepto = " "
                .Update
                
                rsCCDatosPolizaCF.MoveNext
            Loop
        End With
        
        pEjecutaSentencia "UPDATE CCMovimientoCreditoPoliza SET INTNUMEROPOLIZACANCELACION = " + CStr(llngNumPoliza) + ", CHRFOLIOFACTURA = '" + CStr(vlStrFolioFactura) + "' where intnummovimiento = " + CStr(allngAgregarCreditos(vlIntCont))
    Next
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGuardaPoliza"))
End Sub


Public Sub pGeneraFacturaDirecta(ByRef frm As Form, ByRef vllngFormatoaUsar As Long, ByRef intTipoEmisionComprobante As Integer, ByRef intTipoCFDFactura As Integer, bitExtranjero As Integer, _
                                ByRef strRFC As String, ByRef vlstrRazonSocial As String, ByRef llngPersonaGraba As Long, ByRef vlblnEsCredito As Boolean, ByRef vlblnPagoForma As Boolean, intMovimiento As Integer, strTotal As String, _
                                ByRef aFormasPago() As FormasPago, intPesos As Integer, ByRef vldblTipoCambio As Double, ByRef llngNumReferencia As Long, ByRef lstrTipoCliente As String, ByRef lblnEntraCorte As Boolean, _
                                strFecha As String, strFolio As String, llngNumPoliza As Long, intCboUsoCFDiListIndex As Integer, intCboUsoCFDiItemData As Long, intMotivosFacturaListIndex As Integer, intRetencionIVA As Integer, _
                                strIVA As String, strDescuentos As String, strNumCliente As String, strTipoCambio As String, ByRef cstrCantidad4Decimales As String, ByRef strSerie As String, strRetencionISR As String, strRetencionIVA As String, _
                                strObservaciones As String, intCboTarifaItemData As Integer, ByRef llngNumCorte As Long, intFacturaSustitutaDFP As Integer, intFacturaASustituirDFPListCount As Integer, ByRef strAnoAprobacion As String, ByRef strNumeroAprobacion As String, _
                                ByRef cstrCantidad As String, ByRef cintTipoFormato As Integer, ByRef vlblnMultiempresa As Boolean, intRetencionISR As Integer, intCboTarifaListIndex As Integer, ByRef arrTarifas() As typTarifaImpuesto, ByRef vgintnumemprelacionada As Integer, _
                                ByRef vlblnCuentaIngresoSaldada As Boolean, ByRef vlintBitSaldarCuentas As Long, ByRef apoliza() As TipoPoliza, ByRef arrDatosFisc() As DatosFiscales, ByRef lblnConsulta As Boolean, ByRef vldblTotalIVACredito As Double, ByRef llngNumCtaCliente As Long, _
                                     ByRef llngNumFormaCredito As Long, strSubtotal As String, ByRef dblProporcionIVA As Double, ByRef vldblComisionIvaBancaria As Double, tipoPago As DirectaMasiva, ByRef aPoliza2() As RegistroPoliza) 'Parametro ultimo: 0 si es directa, 1 si es masiva
                                      
          Dim intError As Integer         'Error en transacción
          Dim clsFacturaDirecta As clsFactura
          Dim rsFactura As New ADODB.Recordset
          Dim strTotalLetras As String
          Dim lngidfactura As Long
          Dim vlRFC As String
          Dim vlNombre As String
          Dim strSentencia As String
          Dim rs As New ADODB.Recordset
          Dim vllngPvFacturaConsecutivo As Long
          Dim vllngCorteUsado As Long
          Dim vlblnbandera As Boolean
          Dim vlstrTipoPacienteCredito As String          'Sería 'PI' 'PE' 'EM' 'CO' 'ME'
          Dim vllngCveClienteCredito As Long              'Clave del empledo o del médico
          Dim intUsoCFDI As Long
          Dim intcontador As Integer
          Dim i As Integer
        Dim vlstrFolioDocumento As String
        Dim alstrParametrosSalida() As String
        Dim vllngFoliosFaltantes As Long
        
        'Variables para Créditos para facturar
        Dim vllngNumPolizaCF As Long
        Dim vlStrMsg As String
        Dim vlIntCont As Integer
        Dim vlstrsqlCredFact As String
        Dim rsCredFact As New ADODB.Recordset
              
2      On Error GoTo NotificaError
          
          '*********************************** OPCIONES AGREGADAS PARA CFD'S ************************************
          'Identifica el tipo de formato a utilizar
          'vllngFormatoaUsar = llngFormato
          
          'Se valida en caso de no haber formato activo mostrar mensaje y cancelar transacción
3         If vllngFormatoaUsar = 0 Then
              'No se encontró un formato válido de factura.
4             MsgBox SIHOMsg(373), vbCritical, "Mensaje"
5             frm.pReinicia
6             Exit Sub
7         End If
          
          'Se compara el tipo de folio con el tipo de formato a utilizar con la fn "fintTipoEmisionComprobante"
          '(intTipoEmisionComprobante: 0 = Error, 1 = Físico, 2 = Digital)
8         intTipoEmisionComprobante = fintTipoEmisionComprobante("FA", vllngFormatoaUsar)
          
          'Si los folios y los formatos no son compatibles...
9         If intTipoEmisionComprobante = 0 Then   'ERROR
10            frm.blnNoFolios = True
              'Si es error, se cancela la transacción
11            Exit Sub
12        End If
          
13        If intTipoEmisionComprobante = 2 Then
              'Se revisa el tipo de CFD de la Factura (0 = CFD, 1 = CFDi, 2 = Físico, 3 = Error)
14            intTipoCFDFactura = fintTipoCFD("FA", vllngFormatoaUsar)
              
              'Si aparece un error terminar la transacción
15            If intTipoCFDFactura = 3 Then   'ERROR
                  'Si es error, se cancela la transacción
16                Exit Sub
17            End If
18        End If
          
19        If bitExtranjero = 1 Then
20            vlRFC = "XEXX010101000"
21        Else
22            vlRFC = IIf(Len(fStrRFCValido(strRFC)) < 12 Or Len(fStrRFCValido(strRFC)) > 13, "XAXX010101000", fStrRFCValido(strRFC))
23        End If
          
          'vlNombre = Trim(lblCliente.Caption)
24        vlNombre = Trim(vlstrRazonSocial)
          
      '****************************************************************************************
          'Validar uso del comprobante y las claves de productos/servicios y unidades
25        If tipoPago.intDirectaMasiva = 0 Then
26            If Not frm.fblnValidaSAT Then
27                Exit Sub
28            End If
29        End If
          
30        If tipoPago.intDirectaMasiva = 1 And llngPersonaGraba = 0 Then
31            llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
32        ElseIf tipoPago.intDirectaMasiva = 0 Then
33            llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
34        End If
35        If llngPersonaGraba = 0 Then Exit Sub
          
          '------------------------------------------------------------------------------------
          ' Agregado para caso 8644
36        vlblnEsCredito = False
37        vlblnPagoForma = False
38        If intMovimiento = 0 Then
39            If Val(Format(strTotal, "")) > 0 Then
40                If tipoPago.intDirectaMasiva = 0 Then
41                    vlblnPagoForma = fblnFormasPagoPos(aFormasPago(), IIf(intPesos = 1, Val(Format(strTotal, "")), Val(Format(strTotal, "")) * vldblTipoCambio), True, vldblTipoCambio, True, llngNumReferencia, lstrTipoCliente, Trim(Replace(Replace(Replace(strRFC, "-", ""), "_", ""), " ", "")), False, False, True, "frmFacturacionDirecta")
42                Else
43                    vlblnPagoForma = fblnFormasPagoPos(aFormasPago(), IIf(intPesos = 1, Val(Format(strTotal, "")), Val(Format(strTotal, "")) * vldblTipoCambio), True, vldblTipoCambio, True, llngNumReferencia, lstrTipoCliente, Trim(Replace(Replace(Replace(strRFC, "-", ""), "_", ""), " ", "")), False, False, True, "frmFacturacionDirecta", 1, tipoPago.strReferenciaPago, tipoPago.strFormaPago, tipoPago.strFechaPago, tipoPago.vlStrRFCPago, tipoPago.vlstrCuenta, tipoPago.vlstrtipopago, tipoPago.vlStrClaveBancoSAT, tipoPago.vlstrClaveCuentaBancaria, tipoPago.vlStrTipoCargoBancario)
44                End If
                  
45                If vlblnPagoForma Then
46                    intcontador = 0
47                    Do While intcontador <= UBound(aFormasPago(), 1)
48                        If aFormasPago(intcontador).vlbolEsCredito Then
49                            vlblnEsCredito = True
50                        End If
51                        intcontador = intcontador + 1
52                    Loop
53                End If
54            End If

55            If Not (vlblnPagoForma) Then Exit Sub             ' Si <ESC> a las formas de pago
56        End If
          '------------------------------------------------------------------------------------
             
57        EntornoSIHO.ConeccionSIHO.BeginTrans
          
58        intError = fintErrorGrabar(frm, strFolio, llngNumCorte, strFecha, lblnEntraCorte, llngNumFormaCredito)
            '------------------------'
            '- Folio de la factura -'
            '------------------------'
            vllngFoliosFaltantes = 1
            vlstrFolioDocumento = ""
            pCargaArreglo alstrParametrosSalida, vllngFoliosFaltantes & "|" & ADODB.adBSTR & "|" & strFolio & "|" & ADODB.adBSTR & "|" & strSerie & "|" & ADODB.adBSTR & "|" & strNumeroAprobacion & "|" & ADODB.adBSTR & "|" & strAnoAprobacion & "|" & ADODB.adBSTR
            frsEjecuta_SP "FA|" & vgintNumeroDepartamento & "|1", "sp_gnFolios", , , alstrParametrosSalida
            pObtieneValores alstrParametrosSalida, vllngFoliosFaltantes, strFolio, strSerie, strNumeroAprobacion, strAnoAprobacion
    
            '|Si la serie está vacía el SP regresa un espacio en blanco por eso se debe de hacer el TRIM
            vlstrFolioDocumento = Trim(strSerie) & strFolio
            
            If Trim(vlstrFolioDocumento) = "0" Then
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               'No existen folios activos para este documento.
               MsgBox SIHOMsg(291), vbCritical, "Mensaje"
               Exit Sub
            End If
    
            If Trim(vlstrFolioDocumento) = "" Then
               'No se pudo obtener el folio para este documento, intente de nuevo.
               MsgBox SIHOMsg(1390), vbCritical, "Mensaje"
               EntornoSIHO.ConeccionSIHO.RollbackTrans
               Exit Sub
            End If
            strFolio = vlstrFolioDocumento
            frmFacturacionDirecta.lblFolio.Caption = strFolio
            '------------------------------------------'
            
59        If intError = 0 Then
60            Set clsFacturaDirecta = New clsFactura
61            If Not lblnEntraCorte Then
                  '------------------------------------------------------'
                  ' 1.- Insertar la póliza de la factura y tomar el id.
62                llngNumPoliza = flngInsertarPoliza(CDate(strFecha), "D", "FACTURA " & Trim(strFolio), llngPersonaGraba)
                  '------------------------------------------------------'
                  ' 2.- Guardar la factura
63                If intCboUsoCFDiListIndex > -1 Then
64                    intUsoCFDI = intCboUsoCFDiItemData
65                Else
66                    intUsoCFDI = 0
67                End If
                  
68                If intMotivosFacturaListIndex = 3 And intRetencionIVA = 1 Then
69                    lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(strFolio), (CDate(strFecha) + fdtmServerHora), vlRFC, vlNombre, arrDatosFisc(0).strDomicilio, arrDatosFisc(0).strNumExterior, arrDatosFisc(0).strNumInterior, Val(Format(strIVA, cstrCantidad4Decimales)), Val(Format(strDescuentos, cstrCantidad4Decimales)), " ", CLng(strNumCliente), "C", vgintNumeroDepartamento, llngPersonaGraba, 0, 0, Val(Format(strTotal, cstrCantidad4Decimales)), IIf(intPesos = 1, 1, 0), Val(Format(strTipoCambio)), arrDatosFisc(0).strTelefono, "C", CLng(strNumCliente), 0, llngNumPoliza, arrDatosFisc(0).lstrCalleNumero, arrDatosFisc(0).lstrColonia, arrDatosFisc(0).lstrCiudad, arrDatosFisc(0).lstrEstado, arrDatosFisc(0).lstrCodigo, glngCveImpuesto, arrDatosFisc(0).llngCveCiudad, fstrNoLetras(Replace(strFolio, strSerie, "")), strSerie, intUsoCFDI, CDbl(strRetencionISR), 0, intMotivosFacturaListIndex, 0, CDbl(strRetencionIVA), strObservaciones, vlstrRegimenFiscal)
                      
70                Else
71                    lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(strFolio), (CDate(strFecha) + fdtmServerHora), vlRFC, vlNombre, arrDatosFisc(0).strDomicilio, arrDatosFisc(0).strNumExterior, arrDatosFisc(0).strNumInterior, Val(Format(strIVA, cstrCantidad4Decimales)), Val(Format(strDescuentos, cstrCantidad4Decimales)), " ", CLng(strNumCliente), "C", vgintNumeroDepartamento, llngPersonaGraba, 0, 0, Val(Format(strTotal, cstrCantidad4Decimales)), IIf(intPesos = 1, 1, 0), Val(Format(strTipoCambio)), arrDatosFisc(0).strTelefono, "C", CLng(strNumCliente), 0, llngNumPoliza, arrDatosFisc(0).lstrCalleNumero, arrDatosFisc(0).lstrColonia, arrDatosFisc(0).lstrCiudad, arrDatosFisc(0).lstrEstado, arrDatosFisc(0).lstrCodigo, glngCveImpuesto, arrDatosFisc(0).llngCveCiudad, fstrNoLetras(Replace(strFolio, strSerie, "")), strSerie, intUsoCFDI, CDbl(strRetencionISR), CDbl(strRetencionIVA), intMotivosFacturaListIndex, intCboTarifaItemData, , strObservaciones, vlstrRegimenFiscal)
72                End If

                  '------------------------------------------------------'
                  ' 3.- Guardar del detalle de la factura
73                pGuardaDetalleFactura lngidfactura, frm, strFolio, vlblnMultiempresa, cstrCantidad4Decimales, intMotivosFacturaListIndex, intRetencionISR, arrTarifas(), intCboTarifaListIndex, intRetencionIVA, strNumCliente, vgintnumemprelacionada, vlblnCuentaIngresoSaldada, vlintBitSaldarCuentas, cstrCantidad, intPesos, vldblTipoCambio, apoliza(), strTotal, strIVA, strRetencionIVA, strRetencionISR, strDescuentos, tipoPago
                  '------------------------------------------------------'
                  ' 4.- Guardar el detalle de la póliza
74                pGuardaDetallePoliza lngidfactura, apoliza(), llngNumPoliza, vldblTotalIVACredito, vlblnPagoForma, strTotal, cstrCantidad, intPesos, llngNumCtaCliente, intMotivosFacturaListIndex, intRetencionISR, strRetencionISR, vlblnEsCredito, cstrCantidad4Decimales, vldblTipoCambio, intRetencionIVA, strRetencionIVA, strIVA, aFormasPago(), frm, strFolio, strNumCliente, aPoliza2(), strSubtotal, strFecha, llngPersonaGraba, dblProporcionIVA, llngNumCorte, vldblComisionIvaBancaria
                  '------------------------------------------------------'
                  ' 5.- Guardar el movimiento de crédito
75                pGuardaCredito lngidfactura, frm, vlblnMultiempresa, intMovimiento, cstrCantidad4Decimales, strRetencionIVA, strRetencionISR, strIVA, strSubtotal, strTotal, intPesos, vldblTipoCambio, vlblnPagoForma, strNumCliente, llngNumCtaCliente, strFolio, llngPersonaGraba, strFecha
                  '------------------------------------------------------'
                  'Se agrego para liberar los movimientos en caso de que se hayan agregado créditos para facturar
                  If lblnCreditosaFacturar = True Then
                      For vlIntCont = 1 To vlintAgregarCreditos
                          'Se agregó para que guarde la póliza con el numero de folio
                          vlstrsqlCredFact = "Select chrfolioreferencia from ccmovimientocredito where intnummovimiento = " + CStr(allngAgregarCreditos(vlIntCont))
                          Set rsCredFact = frsRegresaRs(vlstrsqlCredFact, adLockOptimistic, adOpenDynamic)
                          If rsCredFact.RecordCount > 0 Then
                              vlStrMsg = vlStrMsg & Trim(CStr(rsCredFact!chrfolioreferencia)) & ", "
                          End If
                      Next
                      rsCredFact.Close
                      vllngNumPolizaCF = flngInsertarPoliza(CDate(strFecha), "D", "FACTURACIÓN DE CRÉDITO PARA FACTURAR " & vlStrMsg, llngPersonaGraba)
                      pGuardaPolizaCreditosaFacturar vllngNumPolizaCF, strFolio
                  End If
                  '------------------------------------------------------'
                  ' 6.- Liberar para que se pueda hacer un cierre
76                pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
                  '------------------------------------------------------'
                  ' 7.- Insertar en tabla de movimientos fuera del corte
77                vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & "FA" & "|" & CStr(llngNumPoliza) & "|" & CStr(llngPersonaGraba) & "|" & CStr(vgintNumeroDepartamento)
78                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSMOVIMIENTOFUERACORTE"
79            Else
                  '------------------------------------------------------'
                  ' 1.- Guardar la factura
80                If intCboUsoCFDiListIndex > -1 Then
81                    intUsoCFDI = intCboUsoCFDiItemData
82                Else
83                    intUsoCFDI = 0
84                End If
85                If intMotivosFacturaListIndex = 3 And intRetencionIVA = 1 Then
86                    lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(strFolio), (CDate(strFecha) + fdtmServerHora), vlRFC, vlNombre, arrDatosFisc(0).strDomicilio, arrDatosFisc(0).strNumExterior, arrDatosFisc(0).strNumInterior, Val(Format(strIVA, cstrCantidad4Decimales)), Val(Format(strDescuentos, cstrCantidad4Decimales)), " ", CLng(strNumCliente), "C", vgintNumeroDepartamento, llngPersonaGraba, llngNumCorte, 0, Val(Format(strTotal, cstrCantidad4Decimales)), IIf(intPesos = 1, 1, 0), Val(Format(strTipoCambio)), arrDatosFisc(0).strTelefono, "C", CLng(strNumCliente), 0, 0, arrDatosFisc(0).lstrCalleNumero, arrDatosFisc(0).lstrColonia, arrDatosFisc(0).lstrCiudad, arrDatosFisc(0).lstrEstado, arrDatosFisc(0).lstrCodigo, glngCveImpuesto, arrDatosFisc(0).llngCveCiudad, fstrNoLetras(Replace(strFolio, strSerie, "")), strSerie, intUsoCFDI, CDbl(strRetencionISR), 0, intMotivosFacturaListIndex, 0, CDbl(strRetencionIVA), strObservaciones, vlstrRegimenFiscal)
87                Else
88                    lngidfactura = clsFacturaDirecta.flngInsFactura(Trim(strFolio), (CDate(strFecha) + fdtmServerHora), vlRFC, vlNombre, arrDatosFisc(0).strDomicilio, arrDatosFisc(0).strNumExterior, arrDatosFisc(0).strNumInterior, Val(Format(strIVA, cstrCantidad4Decimales)), Val(Format(strDescuentos, cstrCantidad4Decimales)), " ", CLng(strNumCliente), "C", vgintNumeroDepartamento, llngPersonaGraba, llngNumCorte, 0, Val(Format(strTotal, cstrCantidad4Decimales)), IIf(intPesos = 1, 1, 0), Val(Format(strTipoCambio)), arrDatosFisc(0).strTelefono, "C", CLng(strNumCliente), 0, 0, arrDatosFisc(0).lstrCalleNumero, arrDatosFisc(0).lstrColonia, arrDatosFisc(0).lstrCiudad, arrDatosFisc(0).lstrEstado, arrDatosFisc(0).lstrCodigo, glngCveImpuesto, arrDatosFisc(0).llngCveCiudad, fstrNoLetras(Replace(strFolio, strSerie, "")), strSerie, intUsoCFDI, CDbl(strRetencionISR), CDbl(strRetencionIVA), intMotivosFacturaListIndex, intCboTarifaItemData, , strObservaciones, vlstrRegimenFiscal)
89                End If
                  '------------------------------------------------------'
                  ' 2.- Guardar del detalle de la factura
90                pGuardaDetalleFactura lngidfactura, frm, strFolio, vlblnMultiempresa, cstrCantidad4Decimales, intMotivosFacturaListIndex, intRetencionISR, arrTarifas(), intCboTarifaListIndex, intRetencionIVA, strNumCliente, vgintnumemprelacionada, vlblnCuentaIngresoSaldada, vlintBitSaldarCuentas, cstrCantidad, intPesos, vldblTipoCambio, apoliza(), strTotal, strIVA, strRetencionIVA, strRetencionISR, strDescuentos, tipoPago
                  '------------------------------------------------------'
                  
                  'inicializamos el arreglo del corte
91                pAgregarMovArregloCorte 0, 0, "", "", 0, 0, False, "", 0, 0, "", 0, 0, "", ""
                  
                  '------------------------------------------------------'
                  ' 3.- Guardar la factura en el corte
92                pGuardaFacturaCorte lngidfactura, vldblTotalIVACredito, strTotal, cstrCantidad, intPesos, vldblTipoCambio, vlblnPagoForma, llngNumCorte, llngPersonaGraba, strFolio, llngNumFormaCredito, aFormasPago(), strIVA, strSubtotal, strRetencionISR, strRetencionIVA, cstrCantidad4Decimales, strNumCliente, llngNumCtaCliente, strTipoCambio, dblProporcionIVA, vldblComisionIvaBancaria, frm
                  '------------------------------------------------------'
                  ' 4.- Guardar la póliza en el corte
93                pGuardaPolizaCorte apoliza(), llngNumCorte, llngPersonaGraba, strFolio, vlblnPagoForma, cstrCantidad, intPesos, vldblTipoCambio, intMotivosFacturaListIndex, intRetencionISR, strRetencionISR, intRetencionIVA, strRetencionIVA, strTotal, llngNumCtaCliente, vlblnEsCredito, cstrCantidad4Decimales, strIVA, vldblTotalIVACredito
                  '------------------------------------------------------'
                  ' 5.- Guardar el movimiento de crédito
94                pGuardaCredito lngidfactura, frm, vlblnMultiempresa, intMovimiento, cstrCantidad4Decimales, strRetencionIVA, strRetencionISR, strIVA, strSubtotal, strTotal, intPesos, vldblTipoCambio, vlblnPagoForma, strNumCliente, llngNumCtaCliente, strFolio, llngPersonaGraba, strFecha
                  '------------------------------------------------------'
                  'Se agrego para liberar los movimientos en caso de que se hayan agregado créditos para facturar
                  If lblnCreditosaFacturar = True Then
                      For vlIntCont = 1 To vlintAgregarCreditos
                          'Se agregó para que guarde la póliza con el numero de folio
                          vlstrsqlCredFact = "Select chrfolioreferencia from ccmovimientocredito where intnummovimiento = " + CStr(allngAgregarCreditos(vlIntCont))
                          Set rsCredFact = frsRegresaRs(vlstrsqlCredFact, adLockOptimistic, adOpenDynamic)
                          If rsCredFact.RecordCount > 0 Then
                              vlStrMsg = vlStrMsg & Trim(CStr(rsCredFact!chrfolioreferencia)) & ", "
                          End If
                      Next
                      rsCredFact.Close
                      vllngNumPolizaCF = flngInsertarPoliza(CDate(strFecha), "D", "FACTURACIÓN DE CRÉDITO PARA FACTURAR " & vlStrMsg, llngPersonaGraba)
                      pGuardaPolizaCreditosaFacturar vllngNumPolizaCF, strFolio
                  End If
                  '------------------------------------------------------'
                  ' 6.- Registra en el corte
95                vllngCorteUsado = fRegistrarMovArregloCorte(llngNumCorte, True)
                 
96                If vllngCorteUsado = 0 Then
97                   EntornoSIHO.ConeccionSIHO.RollbackTrans
                     'No se pudieron agregar los movimientos de la operación al corte, intente de nuevo.
98                   MsgBox SIHOMsg(1320), vbExclamation, "Mensaje"
99                   Exit Sub
100               Else
101                 If vllngCorteUsado <> llngNumCorte Then
                   'actualizamos el corte en el que se registró la factura, esto es por si hay un cambio de corte al momento de hacer el registro de la información de la factura
102                 pEjecutaSentencia "Update pvfactura set INTNUMCORTE = " & vllngCorteUsado & " where intConsecutivo = " & lngidfactura
103                 End If
104               End If
105           End If
              
             
106           If intFacturaSustitutaDFP = 1 And intFacturaASustituirDFPListCount > 0 Then
107               For i = 0 To UBound(aFoliosPrevios())
108                   If aFoliosPrevios(i).chrfoliofactura <> "" Then
109                       pEjecutaSentencia "INSERT INTO PVREFACTURACION (chrFolioFacturaActivada, chrFolioFacturaCancelada) " & " VALUES ('" & Trim(strFolio) & "', '" & aFoliosPrevios(i).chrfoliofactura & "')"
110                   End If
111               Next i
112           End If
              
              
                                 
              '-------------------------------------------------------------------------------------------------
              'VALIDACIÓN DE LOS DATOS ANTES DE INSERTAR EN GNCOMPROBANTEFISCLADIGITAL EN EL PROCESO DE TIMBRADO
              '-------------------------------------------------------------------------------------------------
113           If intTipoEmisionComprobante = 2 Then
114              If Not fblnValidaDatosCFDCFDi(lngidfactura, "FA", IIf(intTipoCFDFactura = 1, True, False), CInt(strAnoAprobacion), strNumeroAprobacion) Then
115                 EntornoSIHO.ConeccionSIHO.RollbackTrans
116                 Exit Sub
117              End If
118           End If
                    
119           Call pGuardarLogTransaccion(frm.Name, EnmGrabar, llngPersonaGraba, "FACTURACION DIRECTA A CLIENTES", strFolio)
120           EntornoSIHO.ConeccionSIHO.CommitTrans 'cerramos transacción, ya esta lista la factura
                                  
              '*** GENERACIÓN DEL CFD ***
                                       
                  '<Si se realizará una emisión digital>
121           If intTipoEmisionComprobante = 2 Then
                  '|Genera el comprobante fiscal digital para la factura
                  'Barra de progreso CFD
122               frm.pgbBarraCFD.Value = 70
123               frm.freBarraCFD.Top = 3200
124               Screen.MousePointer = vbHourglass
125               If tipoPago.intDirectaMasiva = 0 Then
126                   frm.lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere..."
127                   frm.freBarraCFD.Top = 3200
128               Else
129                   frm.lblTextoBarraCFD.Caption = "Generando el Comprobante Fiscal Digital, por favor espere... (" & frm.grdFacturas.Row & "/" & (frm.grdFacturas.Rows - 1) & ")"
130                   frm.freBarraCFD.Top = 2100
131               End If
132               frm.freBarraCFD.Visible = True
133               frm.freBarraCFD.Refresh
134               frm.Enabled = False
135               If intTipoCFDFactura = 1 Then
136                  pLogTimbrado 2
137                  pMarcarPendienteTimbre lngidfactura, "FA", vgintNumeroDepartamento
138               End If
139               EntornoSIHO.ConeccionSIHO.BeginTrans 'iniciamos transaccion de timbrado
140               If Not fblnGeneraComprobanteDigital(lngidfactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, IIf(intTipoCFDFactura = 1, True, False)) Then
141                   On Error Resume Next

142                   EntornoSIHO.ConeccionSIHO.CommitTrans
143                   If intTipoCFDFactura = 1 Then pLogTimbrado 1
144                   If vgIntBanderaTImbradoPendiente = 1 Then 'timbre pendiente de confirmar
                         'El comprobante se realizó de manera correcta, sin embargo no fue posible confirmar el timbre fiscal
145                       If tipoPago.intDirectaMasiva = 0 Then
146                           MsgBox Replace(SIHOMsg(1306), "El comprobante", "La factura directa"), vbInformation + vbOKOnly, "Mensaje"
147                       Else
148                           frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 13) = "PENDIENTE DE TIMBRE"
149                       End If
150                   ElseIf vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3 Then  'No se realizó el timbrado
                         '1338, 'La factura no pudo ser timbrada, será cancelada en el sistema.
151                      If tipoPago.intDirectaMasiva = 0 Then
152                           MsgBox SIHOMsg(1338), vbCritical + vbOKOnly, "Mensaje"
153                       Else
154                           frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 13) = "CANCELADA"
155                       End If
                         
156                      pCancelarFacturaDirecta frm, CLng(Trim(strNumCliente)), Trim(strFolio), lblnEntraCorte, llngPersonaGraba, strFecha, vllngCorteUsado, llngNumPoliza
                         
                         'Actualiza PDF al cancelar facturas
                         
157                      If Not fblnGeneraComprobanteDigital(lngidfactura, "FA", 1, Val(strAnoAprobacion), strNumeroAprobacion, False, True, -1) Then
158                             On Error Resume Next
159                      End If
                         
160                      If intTipoCFDFactura = 1 Then pEliminaPendientesTimbre lngidfactura, "FA" 'quitamos la factura de pendientes de timbre fiscal
                         'Imprimimos la factura cancelada
161                      If tipoPago.intDirectaMasiva = 0 Then
162                           fblnImprimeComprobanteDigital lngidfactura, "FA", "I", vllngFormatoaUsar, 1
163                       End If
164                      Screen.MousePointer = vbDefault
165                      frmFacturacionDirecta.Enabled = True
166                      frm.freBarraCFD.Visible = False
167                      frm.pReinicia
168                      frm.txtNumCliente.SetFocus
169                      Exit Sub
170                 End If
171               Else


172                  EntornoSIHO.ConeccionSIHO.CommitTrans
173                  If intTipoCFDFactura = 1 Then
174                     pLogTimbrado 1
175                     pEliminaPendientesTimbre lngidfactura, "FA" 'quitamos la factura de pendientes de timbre fiscal
176                  End If
177               End If
178               frm.pgbBarraCFD.Value = 100
179               frm.freBarraCFD.Top = 3200
180               Screen.MousePointer = vbDefault
181               frm.freBarraCFD.Visible = False
182               frm.Enabled = True
183               If tipoPago.intDirectaMasiva = 0 Then
                      
184               Else
                      'Dim RsComprobante As ADODB.Recordset
                      'strSentencia = "SELECT VCHUUID " & _
                      '   "  FROM GNCOMPROBANTEFISCALDIGITAL " & _
                      '   " WHERE INTCOMPROBANTE = " & lngidfactura & _
                      '   "   AND CHRTIPOCOMPROBANTE = 'FA'"
                      'Set RsComprobante = frsRegresaRs(strSentencia)
185                   If Not (vgIntBanderaTImbradoPendiente = 1 Or vgIntBanderaTImbradoPendiente = 2 Or vgIntBanderaTImbradoPendiente = 3) Then
186                       frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 13) = "TIMBRADA"
187                       frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 14) = strFolio
188                       frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 26) = lngidfactura
189                   ElseIf vgIntBanderaTImbradoPendiente = 1 Then
190                       frm.grdFacturas.TextMatrix(frm.grdFacturas.RowSel, 14) = strFolio
191                   End If
192               End If
193           End If
            
              '*** IMPRESIÓN DEL CFD ***
                                      
              '<Si se realizará una emisión digital>
194           If tipoPago.intDirectaMasiva = 0 Then
195               If intTipoEmisionComprobante = 2 Then
196                   If Not fblnImprimeComprobanteDigital(lngidfactura, "FA", "I", vllngFormatoaUsar, 1) Then
197                       Exit Sub
198                   End If
                      
                      'Verifica si debe mostrarse la pantalla de envío de CFDs por correo electrónico
199                   If fblnPermitirEnvio(strNumCliente) And vgIntBanderaTImbradoPendiente = 0 Then
200                       If MsgBox(SIHOMsg(1090), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
201                           pEnviarCFD "FA", lngidfactura, CLng(vgintClaveEmpresaContable), vlRFC, llngPersonaGraba, frm
202                       End If
203                   End If
          
204               Else
                  '<Emisión física>
                      'Asegúrese de que la impresora esté   lista y  presione aceptar.
205                   MsgBox SIHOMsg(343), vbOKOnly + vbInformation, "Mensaje"
206                   If vgintNumeroModulo <> 2 Then
207                       strTotalLetras = fstrNumeroenLetras(CDbl(Format(strTotal, cstrCantidad)), IIf(intPesos = 1, "pesos", "dólares"), IIf(intPesos = 1, "M.N.", " "))
208                       vgstrParametrosSP = Trim(strFolio) & "|" & strTotalLetras
209                       Set rsFactura = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptFactura")
210                       pImpFormato rsFactura, cintTipoFormato, vllngFormatoaUsar
211                   Else
212                       pImprimeFormato vllngFormatoaUsar, lngidfactura
213                   End If
214               End If
                  
215               frm.pReinicia
216               frm.txtNumCliente.SetFocus
217           End If
218       Else
219           EntornoSIHO.ConeccionSIHO.RollbackTrans
220           MsgBox SIHOMsg(intError), vbOKOnly + vbExclamation, "Mensaje"
221       End If
          
222   Exit Sub
NotificaError:
       Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGeneraFacturaDirecta" & " Linea:" & Erl()))
       frm.cmdSave.Enabled = False
       lblnConsulta = False
       Unload frm
End Sub

Private Sub pGuardaDetallePoliza(lngidfactura As Long, ByRef apoliza() As TipoPoliza, ByRef llngNumPoliza As Long, ByRef vldblTotalIVACredito As Double, ByRef vlblnPagoForma As Boolean, strTotal As String, cstrCantidad As String, intPesos As Integer, ByRef llngNumCtaCliente As Long, intMotivosFacturaListIndex As Integer, intRetencionISR As Integer, strRetencionISR As String, ByRef vlblnEsCredito As Boolean, cstrCantidad4Decimales As String, ByRef vldblTipoCambio As Double, intRetencionIVA As Integer, strRetencionIVA As String, strIVA As String, aFormasPago() As FormasPago, frm As Form, strFolio As String, strNumCliente As String, ByRef aPoliza2() As RegistroPoliza, strSubtotal As String, strFecha As String, ByRef llngPersonaGraba As Long, ByRef dblProporcionIVA As Double, ByRef llngNumCorte As Long, ByRef vldblComisionIvaBancaria As Double)
    Dim lngNumDetalle As Long
    Dim intcontador As Integer
    Dim dblTotalCliente As Double
    Dim dblTotalIVA As Double
    Dim vlintContador As Long
    Dim lngCveConcepto As Long
    Dim dblIVA As Double
    Dim dblSubTotal As Double
    Dim dblTotalFactura As Double
    Dim dblPorcentaje As Double                     'Para calcular qué porcentaje es el pago en crédito respecto al total de la factura
    Dim dblSubtotalCredito As Double                'Para guardar el subtotal en el movimiento de crédito
    Dim dblIVACredito As Double                     'Para guardar el IVA en el movimiento de crédito
    Dim vllngMovimientoCredito As Long              'Para la impresión del pagaré y guarda el Movimiento de crédito en caso de que exista uno.
    Dim vldtmFechaHoy As Date                       'Varible con la Fecha actual
    Dim vldtmHoraHoy As Date                        'Varible con la Hora actual
    Dim dblCantidadMovto As Double                  'Cantidad del movimiento en la póliza
    Dim vllngNumDetalleCorte As Long
    
    Dim dblPorcentajeRetencionIVA As Double
    Dim dblPorcentajeRetencionISR As Double
    
    Dim dblRetencionISR As Double
    Dim dblRetencionIVA As Double
    Dim rsPVFacturaFueraCorteForma As New ADODB.Recordset 'Para guardar las formas de pago cuando la factura no entra al corte
      
    Set rsPVFacturaFueraCorteForma = frsRegresaRs("SELECT * FROM PVFacturaFueraCorteForma WHERE INTNUMFACTURA = -1", adLockOptimistic, adOpenDynamic)
    
    intcontador = 0
    Do While intcontador <= UBound(apoliza(), 1)
        lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, apoliza(intcontador).lngnumCuenta, apoliza(intcontador).dblCantidad, apoliza(intcontador).intNaturaleza)
        intcontador = intcontador + 1
    Loop
    
    vldblTotalIVACredito = 0
    
    If vlblnPagoForma = False Then
        '-------------------------------------------
        'Cargo al cliente
        dblTotalCliente = CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
        If dblTotalCliente <> 0 Then
            lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, llngNumCtaCliente, dblTotalCliente, 1)
        End If
        
        'Cuenta de ISR provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionISR = 1 And CDbl(strRetencionISR) <> 0 Then
                lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False)), CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
            End If
        End If
        
        'Cuenta de IVA provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionIVA = 1 And CDbl(strRetencionIVA) <> 0 Then
                If intMotivosFacturaListIndex = 3 Then
                    lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(3, 1, False), fblnCuentaRetencionImpuestos(3, 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
                Else
                    lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
                End If
            End If
        End If
        
        'IVA no cobrado
        dblTotalIVA = CDbl(Format(strIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio)
        If dblTotalIVA <> 0 Then
           lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, glngCtaIVANoCobrado, dblTotalIVA, 0)
        End If
    Else
        dblTotalCliente = CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
        If dblTotalCliente <> 0 Then
            ReDim aPoliza2(0)
            For vlintContador = 0 To UBound(aFormasPago(), 1)
                If aFormasPago(vlintContador).vlbolEsCredito Then
                    '------------------------------'
                    'Crear el movimiento de crédito
                    '------------------------------'
                    lngCveConcepto = 0

                    dblIVA = CDbl(Format(strIVA, cstrCantidad4Decimales))
                    dblSubTotal = CDbl(Format(strSubtotal, cstrCantidad4Decimales))
                    dblTotalFactura = CDbl(Format(strTotal, cstrCantidad4Decimales))
                    dblRetencionISR = CDbl(Format(strRetencionISR, cstrCantidad4Decimales))
                    dblRetencionIVA = CDbl(Format(strRetencionIVA, cstrCantidad4Decimales))

                    dblPorcentaje = Round(aFormasPago(vlintContador).vldblCantidad / ((dblTotalFactura) * IIf(intPesos = 1, 1, vldblTipoCambio)), 2)
                    dblSubtotalCredito = Format(((dblTotalFactura - dblIVA + dblRetencionISR + dblRetencionIVA) * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje, "###############.0000")
                    dblIVACredito = Format(((dblIVA) * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje, "###############.0000")
                    
                    dblRetencionISR = (dblRetencionISR * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje
                    dblRetencionIVA = (dblRetencionIVA * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje
                    
                    vgstrParametrosSP = fstrFechaSQL(strFecha) _
                                        & "|" & strNumCliente _
                                        & "|" & llngNumCtaCliente _
                                        & "|" & Trim(strFolio) _
                                        & "|" & "FA" _
                                        & "|" & aFormasPago(vlintContador).vldblCantidad _
                                        & "|" & Str(vgintNumeroDepartamento) _
                                        & "|" & Str(llngPersonaGraba) _
                                        & "|" & " " & "|" & "0" _
                                        & "|" & (dblSubtotalCredito - dblRetencionIVA) _
                                        & "|" & dblIVACredito & "|" & dblRetencionISR & "|" & dblRetencionIVA
                    vllngMovimientoCredito = 1
                    frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, vllngMovimientoCredito
                    
                    '--------------------------------------------------------------------'
                    ' Se genera un cargo de acuerdo a la forma de pago credito '
                    '--------------------------------------------------------------------'
                    'lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, llngNumCtaCliente, aFormasPago(vlintcontador).vldblCantidad, 1)
                    pIncluyeMovimiento llngNumCtaCliente, aFormasPago(vlintContador).vldblCantidad, 1, aPoliza2()
                    
                    'calculo del IVa que se pudiera ir a IVA no cobrado
                    If dblIVA > 0 Then
                       dblProporcionIVA = Round(aFormasPago(vlintContador).vldblCantidad / (CDbl(Format(dblTotalFactura, "############.0000")) * IIf(intPesos = 1, 1, vldblTipoCambio)), 2)
                       vldblTotalIVACredito = CDbl(Format(dblIVA, "############.0000")) * dblProporcionIVA
                    End If
                
                Else
                    '--------------------------------------------------------------------'
                    ' Se genera un cargo de acuerdo a la forma de pago que NO es credito '
                    '--------------------------------------------------------------------'
                    'lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, aFormasPago(vlintcontador).vllngCuentaContable, aFormasPago(vlintcontador).vldblCantidad, 1)
                    pIncluyeMovimiento aFormasPago(vlintContador).vllngCuentaContable, aFormasPago(vlintContador).vldblCantidad, 1, aPoliza2()
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
                    If aFormasPago(vlintContador).vllngCuentaComisionBancaria <> 0 And aFormasPago(vlintContador).vldblCantidadComisionBancaria <> 0 Then
                         ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                        pIncluyeMovimiento aFormasPago(vlintContador).vllngCuentaComisionBancaria, aFormasPago(vlintContador).vldblCantidadComisionBancaria, 1, aPoliza2()
                        If aFormasPago(vlintContador).vldblIvaComisionBancaria <> 0 Then
                            ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                            pIncluyeMovimiento glngCtaIVAPagado, aFormasPago(vlintContador).vldblIvaComisionBancaria, 1, aPoliza2()
                        End If
                        ' Se genera un abono por la cantidad de la comisión bancaria y su iva que corresponde a la forma de pago
                        pIncluyeMovimiento aFormasPago(vlintContador).vllngCuentaContable, (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria), 0, aPoliza2()
                     End If
                                        
                    '-----------------------------------------------------------------------------------'
                    ' Guardar en el kárdex del banco si hubo pago por medio de transferencias bancarias '
                    '-----------------------------------------------------------------------------------'
                    vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(strFecha) & "|" & aFormasPago(vlintContador).vlintNumFormaPago & "|" & aFormasPago(vlintContador).lngIdBanco & "|" & _
                                        IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, aFormasPago(vlintContador).vldblCantidad, aFormasPago(vlintContador).vldblDolares) & "|" & IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContador).vldblTipoCambio & "|" & _
                                        fstrTipoMovimientoForma(aFormasPago(vlintContador).vlintNumFormaPago, frm) & "|" & "FA" & "|" & lngidfactura & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(strFecha) & "|" & "1" & "|" & cgstrModulo
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                    
                    vllngNumDetalleCorte = flngObtieneIdentity("SEC_PVMOVIMIENTOBANCOFORMA", 0)
                                            
                    If Not aFormasPago(vlintContador).vlbolEsCredito Then
                        If Trim(aFormasPago(vlintContador).vlstrRFC) <> "" And Trim(aFormasPago(vlintContador).vlstrBancoSAT) <> "" Then
                            frsEjecuta_SP llngNumCorte & "|" & vllngNumDetalleCorte * -1 & "|'" & Trim(aFormasPago(vlintContador).vlstrRFC) & "'|'" & Trim(aFormasPago(vlintContador).vlstrBancoSAT) & "'|'" & Trim(aFormasPago(vlintContador).vlstrCuentaBancaria) & "'|'" & IIf(Trim(aFormasPago(vlintContador).vlstrCuentaBancaria) = "", Null, fstrFechaSQL(Trim(aFormasPago(vlintContador).vldtmFecha))) & "'|'" & Trim(aFormasPago(vlintContador).vlstrBancoExtranjero) & "'", "SP_PVINSCORTECHEQUETRANSCTA"
                        End If
                    End If
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                    vldblComisionIvaBancaria = 0
                    If aFormasPago(vlintContador).vllngCuentaComisionBancaria <> 0 And aFormasPago(vlintContador).vldblCantidadComisionBancaria <> 0 Then
                        If aFormasPago(vlintContador).vldblTipoCambio = 0 Then
                             vldblComisionIvaBancaria = (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria) * -1
                        Else
                             vldblComisionIvaBancaria = (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria) / aFormasPago(vlintContador).vldblTipoCambio * -1
                        End If
                        vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(strFecha) & "|" & aFormasPago(vlintContador).vlintNumFormaPago & "|" & aFormasPago(vlintContador).lngIdBanco & "|" & _
                                            vldblComisionIvaBancaria & "|" & IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContador).vldblTipoCambio & "|" & _
                                            "CBA" & "|" & "FA" & "|" & lngidfactura & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(strFecha) & "|" & "1" & "|" & cgstrModulo
                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                    End If
                End If
                
                With rsPVFacturaFueraCorteForma
                    .AddNew
                    !INTNUMFACTURA = lngidfactura
                    !intFormaPago = aFormasPago(vlintContador).vlintNumFormaPago
                    !intFolioForma = IIf(Trim(aFormasPago(vlintContador).vlstrFolio) = "", 0, Trim(aFormasPago(vlintContador).vlstrFolio))
                    !MNYCantidad = aFormasPago(vlintContador).vldblCantidad
                    !mnytipocambio = aFormasPago(vlintContador).vldblTipoCambio
                    !intCveBanco = aFormasPago(vlintContador).lngIdBanco
                    .Update
                End With
            Next vlintContador
            
            vlintContador = 0
            Do While vlintContador <= UBound(aPoliza2(), 1)
                If aPoliza2(vlintContador).vldblCantidadMovimiento <> 0 Then
                    dblCantidadMovto = Format(aPoliza2(vlintContador).vldblCantidadMovimiento, "Fixed")
                    lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, aPoliza2(vlintContador).vllngNumeroCuenta, dblCantidadMovto, aPoliza2(vlintContador).vlintTipoMovimiento)
                End If
                vlintContador = vlintContador + 1
            Loop
        End If
        
        If vldblTotalIVACredito > 0 Then 'Iva no cobrado
           lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, glngCtaIVANoCobrado, vldblTotalIVACredito * IIf(intPesos = 1, 1, vldblTipoCambio), 0)
        End If
        
        'Iva cobrado
        If Format(CDbl(Format(strIVA, cstrCantidad)), "##########.00") - Format(vldblTotalIVACredito, "##########.00") > 0.01 Then
           lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, glngCtaIVACobrado, (Format(CDbl(Format(strIVA, cstrCantidad)), "##########.00") - Format(vldblTotalIVACredito, "##########.00")) * IIf(intPesos = 1, 1, vldblTipoCambio), 0)
        End If
        
        'Cuenta de ISR provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionISR = 1 And CDbl(strRetencionISR) <> 0 Then
                lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False)), CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
            End If
        End If
        
        'Cuenta de IVA provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionIVA = 1 And CDbl(strRetencionIVA) <> 0 Then
                If intMotivosFacturaListIndex = 3 Then
                    lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(3, 1, False), fblnCuentaRetencionImpuestos(3, 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
                Else
                    lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), 1)
                End If
            End If
        End If
    End If
    
End Sub

Private Sub pGuardaFacturaCorte(lngidfactura As Long, ByRef vldblTotalIVACredito As Double, strTotal As String, cstrCantidad As String, intPesos As Integer, vldblTipoCambio As Double, ByRef vlblnPagoForma As Boolean, ByRef llngNumCorte As Long, ByRef llngPersonaGraba As Long, strFolio As String, ByRef llngNumFormaCredito As Long, ByRef aFormasPago() As FormasPago, strIVA As String, strSubtotal As String, strRetencionISR As String, strRetencionIVA As String, cstrCantidad4Decimales As String, strNumCliente As String, llngNumCtaCliente As Long, strTipoCambio As String, ByRef dblProporcionIVA As Double, ByRef vldblComisionIvaBancaria As Double, ByRef frm As Form)
    Dim rs As New ADODB.Recordset
    Dim rsPvDetalleCorte As New ADODB.Recordset     'Aqui añado los registros del detalle del corte
    Dim vlstrSentencia As String
    Dim vlintContador As Long
    Dim vldtmFechaHoy As Date                       'Varible con la Fecha actual
    Dim vldtmHoraHoy As Date                        'Varible con la Hora actual
    Dim lngCveConcepto As Long
    Dim dblIVA As Double
    Dim dblSubTotal As Double
    Dim dblTotalFactura As Double
    Dim dblPorcentaje As Double                     'Para calcular qué porcentaje es el pago en crédito respecto al total de la factura
    Dim dblSubtotalCredito As Double                'Para guardar el subtotal en el movimiento de crédito
    Dim dblIVACredito As Double                     'Para guardar el IVA en el movimiento de crédito
    Dim lngDepartamento As Long
    Dim vllngMovimientoCredito As Long              'Para la impresión del pagaré y guarda el Movimiento de crédito en caso de que exista uno.
    
    Dim dblPorcentajeRetencionIVA As Double
    Dim dblPorcentajeRetencionISR As Double
    
    Dim dblRetencionISR As Double
    Dim dblRetencionIVA As Double
    
    vldblTotalIVACredito = 0
    
    If vlblnPagoForma = False Then
        If CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio) <> 0 Then
'             vgstrParametrosSP = CStr(llngNumCorte) & "|" & _
'                                 fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & "|" & _
'                                 Trim(lblFolio.Caption) _
'                                 & "|" & "FA" _
'                                 & "|" & CStr(llngNumFormaCredito) _
'                                 & "|" & CStr(CDbl(Format(lblTotal.Caption, cstrCantidad)) * IIf(optPesos(0).Value, 1, vldblTipoCambio)) _
'                                 & "|" & "0" _
'                                 & "|" & "0" _
'                                 & "|" & CStr(llngNumCorte)
'             frsEjecuta_SP vgstrParametrosSP, "sp_PvInsDetalleCorte"
             pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", 0, CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio), False, _
             fstrFechaSQL(fdtmServerFecha, fdtmServerHora, True), llngNumFormaCredito, 0, "0", llngNumCorte, 1, Trim(strFolio), "FA"
        End If
    Else
        If CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio) <> 0 Then
            vlstrSentencia = "Select * from PVDetalleCorte where intConsecutivo = -1"
            Set rsPvDetalleCorte = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
        
            vldtmFechaHoy = fdtmServerFecha
            vldtmHoraHoy = fdtmServerHora
        
            For vlintContador = 0 To UBound(aFormasPago(), 1)
'                With rsPvDetalleCorte
'                    .AddNew
'                    !intNumCorte = llngNumCorte
'                    !DTMFECHAHORA = vldtmFechaHoy + vldtmHoraHoy
'                    !CHRFOLIODOCUMENTO = Trim(lblFolio.Caption)
'                    !chrTipoDocumento = "FA"
'                    !intFormaPago = aFormasPago(vlintcontador).vlintNumFormaPago
'                    !MNYCANTIDADPAGADA = IIf(aFormasPago(vlintcontador).vldblTipoCambio = 0, aFormasPago(vlintcontador).vldblCantidad, aFormasPago(vlintcontador).vldblDolares)
'                    !mnyTipoCambio = aFormasPago(vlintcontador).vldblTipoCambio
'                    !INTFOLIOCHEQUE = IIf(Trim(aFormasPago(vlintcontador).vlstrFolio) = "", "0", Trim(aFormasPago(vlintcontador).vlstrFolio))
'                    !intNumCorteDocumento = llngNumCorte
'                    .Update
'                End With
                 pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", 0, IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, aFormasPago(vlintContador).vldblCantidad, aFormasPago(vlintContador).vldblDolares), _
                 False, vldtmFechaHoy + vldtmHoraHoy, CLng(aFormasPago(vlintContador).vlintNumFormaPago), aFormasPago(vlintContador).vldblTipoCambio, IIf(Trim(aFormasPago(vlintContador).vlstrFolio) = "", "0", Trim(aFormasPago(vlintContador).vlstrFolio)), _
                 llngNumCorte, 1, Trim(strFolio), "FA", aFormasPago(vlintContador).vlbolEsCredito, aFormasPago(vlintContador).vlstrRFC, aFormasPago(vlintContador).vlstrBancoSAT, aFormasPago(vlintContador).vlstrBancoExtranjero, aFormasPago(vlintContador).vlstrCuentaBancaria, aFormasPago(vlintContador).vldtmFecha
                
                If aFormasPago(vlintContador).vlbolEsCredito Then
                    '------------------------------'
                    'Crear el movimiento de crédito
                    '------------------------------'
                    lngCveConcepto = 0
                    
                    dblIVA = CDbl(Format(strIVA, cstrCantidad4Decimales))
                    dblSubTotal = CDbl(Format(strSubtotal, cstrCantidad4Decimales))
                    dblTotalFactura = CDbl(Format(strTotal, cstrCantidad4Decimales))
                    dblRetencionISR = CDbl(Format(strRetencionISR, cstrCantidad4Decimales))
                    dblRetencionIVA = CDbl(Format(strRetencionIVA, cstrCantidad4Decimales))
                        
                    dblPorcentaje = Round(aFormasPago(vlintContador).vldblCantidad / ((dblTotalFactura) * IIf(intPesos = 1, 1, vldblTipoCambio)), 2)
                    dblSubtotalCredito = Format(((dblTotalFactura - dblIVA + dblRetencionISR + dblRetencionIVA) * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje, "###############.0000")
                    dblIVACredito = Format(((dblIVA) * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje, "###############.0000")
                    
                    dblRetencionISR = (dblRetencionISR * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje
                    dblRetencionIVA = (dblRetencionIVA * IIf(intPesos = 1, 1, vldblTipoCambio)) * dblPorcentaje
                    
                    vgstrParametrosSP = fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) _
                                        & "|" & strNumCliente _
                                        & "|" & llngNumCtaCliente _
                                        & "|" & Trim(strFolio) _
                                        & "|" & "FA" _
                                        & "|" & aFormasPago(vlintContador).vldblCantidad _
                                        & "|" & Str(vgintNumeroDepartamento) _
                                        & "|" & Str(llngPersonaGraba) _
                                        & "|" & " " & "|" & "0" _
                                        & "|" & (dblSubtotalCredito - dblRetencionIVA) _
                                        & "|" & dblIVACredito & "|" & dblRetencionISR & "|" & dblRetencionIVA
                    vllngMovimientoCredito = 1
                    frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, vllngMovimientoCredito
                    
                    '------------------------------'
                    ' Generar movimiento a crédito '
                    '------------------------------'
                    'pInsCortePoliza llngNumCorte, Trim(lblFolio.Caption), "FA", llngNumCtaCliente, aFormasPago(vlintcontador).vldblCantidad, True
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", llngNumCtaCliente, aFormasPago(vlintContador).vldblCantidad, True, _
                    "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                    
                    'calculo del IVa que se pudiera ir a IVA no cobrado
                    If dblIVA > 0 Then
                        '| Si el importe de la forma de pago es en dólares y la factura se está realizando en pesos
                        '| se convierte a pesos el importe de la forma de pago para calcular la proporción del IVA
                        If aFormasPago(vlintContador).intMoneda <> 0 And intPesos = 1 Then
                            dblProporcionIVA = (aFormasPago(vlintContador).vldblCantidad * Val(strTipoCambio)) / CDbl(Format(dblTotalFactura, "############.0000"))
                        Else
                            '| Si el importe de la forma de pago es en pesos y la factura se está realizando en dólares
                            '| se convierte a dólares el importe de la forma de pago para calcular la proporción del IVA
                            If aFormasPago(vlintContador).intMoneda = 0 And intPesos = 0 Then
                                dblProporcionIVA = (aFormasPago(vlintContador).vldblCantidad / Val(strTipoCambio)) / CDbl(Format(dblTotalFactura, "############.0000"))
                            Else
                                dblProporcionIVA = aFormasPago(vlintContador).vldblCantidad / CDbl(Format(dblTotalFactura, "############.0000"))
                            End If
                        End If
                        vldblTotalIVACredito = CDbl(Format(dblIVA, "############.0000")) * dblProporcionIVA
                    End If
                    
                Else ' Osea que la forma de pago --NO es credito--
                    '--------------------------------------------------------------------'
                    ' Se genera un cargo de acuerdo a la forma de pago que NO es credito '
                    '--------------------------------------------------------------------'
                    'pInsCortePoliza llngNumCorte, Trim(lblFolio.Caption), "FA", aFormasPago(vlintcontador).vllngCuentaContable, aFormasPago(vlintcontador).vldblCantidad, True
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", aFormasPago(vlintContador).vllngCuentaContable, aFormasPago(vlintContador).vldblCantidad, True, _
                    "", 0, 0, "", 0, 2, Trim(strFolio), "FA", aFormasPago(vlintContador).vlbolEsCredito, aFormasPago(vlintContador).vlstrRFC, aFormasPago(vlintContador).vlstrBancoSAT, aFormasPago(vlintContador).vlstrBancoExtranjero, aFormasPago(vlintContador).vlstrCuentaBancaria, aFormasPago(vlintContador).vldtmFecha
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registran los movimientos contables referente a la comision bancaria
                    If aFormasPago(vlintContador).vllngCuentaComisionBancaria <> 0 And aFormasPago(vlintContador).vldblCantidadComisionBancaria <> 0 Then
                         ' Se genera un cargo de acuerdo la comisión bancaria que corresponde a la forma de pago
                        pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", aFormasPago(vlintContador).vllngCuentaComisionBancaria, aFormasPago(vlintContador).vldblCantidadComisionBancaria, True, _
                        "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                        
                        If aFormasPago(vlintContador).vldblIvaComisionBancaria <> 0 Then
                            ' Movimiento contable por el IVA pagado que corresponde de la comisión bancaria
                            pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", glngCtaIVAPagado, aFormasPago(vlintContador).vldblIvaComisionBancaria, True, _
                            "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                        End If
                        ' Se genera un abono por la cantidad de la comisión bancaria y su iva que corresponde a la forma de pago
                        pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", aFormasPago(vlintContador).vllngCuentaContable, (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria), False, _
                        "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                     End If
                                        
                    '-----------------------------------------------------------------------------------'
                    ' Guardar en el kárdex del banco si hubo pago por medio de transferencias bancarias '
                    '-----------------------------------------------------------------------------------'
                    vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(vlintContador).vlintNumFormaPago & "|" & aFormasPago(vlintContador).lngIdBanco & "|" & _
                                        IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, aFormasPago(vlintContador).vldblCantidad, aFormasPago(vlintContador).vldblDolares) & "|" & IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContador).vldblTipoCambio & "|" & _
                                        fstrTipoMovimientoForma(aFormasPago(vlintContador).vlintNumFormaPago, frm) & "|" & "FA" & "|" & lngidfactura & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo
                    frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                    
                    ' Agregado para caso 8741
                    ' Si la forma de pago es de tipo tarjeta se registra la disminución de la cantidad referente a la comision bancaria
                    vldblComisionIvaBancaria = 0
                    If aFormasPago(vlintContador).vllngCuentaComisionBancaria <> 0 And aFormasPago(vlintContador).vldblCantidadComisionBancaria <> 0 Then
                        If aFormasPago(vlintContador).vldblTipoCambio = 0 Then
                             vldblComisionIvaBancaria = (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria) * -1
                        Else
                             vldblComisionIvaBancaria = (aFormasPago(vlintContador).vldblCantidadComisionBancaria + aFormasPago(vlintContador).vldblIvaComisionBancaria) / aFormasPago(vlintContador).vldblTipoCambio * -1
                        End If
                        vgstrParametrosSP = llngNumCorte & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & aFormasPago(vlintContador).vlintNumFormaPago & "|" & aFormasPago(vlintContador).lngIdBanco & "|" & _
                                            vldblComisionIvaBancaria & "|" & IIf(aFormasPago(vlintContador).vldblTipoCambio = 0, 1, 0) & "|" & aFormasPago(vlintContador).vldblTipoCambio & "|" & _
                                            "CBA" & "|" & "FA" & "|" & lngidfactura & "|" & llngPersonaGraba & "|" & vgintNumeroDepartamento & "|" & fstrFechaSQL(Format(vldtmFechaHoy, "dd/mm/yyyy"), Format(vldtmHoraHoy, "hh:mm:ss")) & "|" & "1" & "|" & cgstrModulo
                        frsEjecuta_SP vgstrParametrosSP, "Sp_PvInsMovimientoBancoForma"
                    End If
                    
                End If
            Next vlintContador
        End If
    End If
End Sub

Public Function fblnCuentaProvision(inttipomovimiento As Integer, inttipoimpuesto As Integer, intblnMostrarMensaje As Boolean) As Long
1         On Error GoTo NotificaError
          Dim rs As ADODB.Recordset
          Dim vlstrparametro As String
          Dim vllngMensaje As Integer
          Dim vllngMensajeNoAceptaMovimientos As Integer
          Dim vlblnNoAceptaMovimientos As Boolean
          
          'inttipomovimiento = Tipo del movimiento (1 = Honorarios profesionales, 2 = Recibo de arrendamiento)
          'inttipoimpuesto = Tipo de impuesto (1 = ISR, 2 = IVA)
          'intblnMostrarMensaje = Mostrar mensaje si no está configurada la cuenta

2         fblnCuentaProvision = 0
3         vllngMensaje = 0
4         vllngMensajeNoAceptaMovimientos = 0
5         vlstrparametro = ""

6         If inttipomovimiento = 1 Then
              ' Honorarios profesionales
7             If inttipoimpuesto = 1 Then
                  'ISR
8                 vlstrparametro = "INTCTAPROVISIONISRFACHONORARIO"
                  'No se ha configurado la cuenta contable para provisionar el ISR en facturas de honorarios profesionales, favor de verificar.
9                 vllngMensaje = 1527
                  
                  'La cuenta contable configurada para provisionar el ISR en facturas de honorarios profesionales no acepta movimientos, favor de verificar.
10                vllngMensajeNoAceptaMovimientos = 1535
11            Else
                  'IVA
12                vlstrparametro = "INTCTAPROVISIONIVAFACHONORARIO"
                  'No se ha configurado la cuenta contable para provisionar la retención de IVA en facturas de honorarios profesionales, favor de verificar.
13                vllngMensaje = 1528
                  
                  'La cuenta contable configurada para provisionar la retención de IVA en facturas de honorarios profesionales no acepta movimientos, favor de verificar.
14                vllngMensajeNoAceptaMovimientos = 1536
15            End If
16        ElseIf inttipomovimiento = 2 Then
              ' Recibo de arrendamiento
17            If inttipoimpuesto = 1 Then
                  'ISR
18                vlstrparametro = "INTCTAPROVISIONISRFACARRENDAMIENTO"
                  'No se ha configurado la cuenta contable para provisionar el ISR en facturas de arrendamiento, favor de verificar.
19                vllngMensaje = 1529
                  
                  'La cuenta contable configurada para provisionar el ISR en facturas de arrendamiento no acepta movimientos, favor de verificar.
20                vllngMensajeNoAceptaMovimientos = 1537
21            Else
                  'IVA
22                vlstrparametro = "INTCTAPROVISIONIVAFACARRENDAMIENTO"
                  'No se ha configurado la cuenta contable para provisionar la retención de IVA en facturas de arrendamiento, favor de verificar.
23                vllngMensaje = 1530
                  
                  'La cuenta contable configurada para provisionar la retención de IVA en facturas de arrendamiento no acepta movimientos, favor de verificar.
24                vllngMensajeNoAceptaMovimientos = 1538
25            End If
26        Else
27            vlstrparametro = "INTCTAPROVISIONSERVICIOSCLIENTE"
              
              'No se encuentran registradas las cuentas contables de retención y provisión de servicios de clientes en los parámetros de contabilidad.
28            vllngMensaje = 1573
29        End If
          
30        vlblnNoAceptaMovimientos = False
          
31        If vlstrparametro <> "" Then
32            Set rs = frsSelParametros("CN", vgintClaveEmpresaContable, vlstrparametro)
33            If Not rs.EOF Then
34                fblnCuentaProvision = IIf(IsNull(rs("Valor")), 0, rs("Valor"))
                  
35                If fstrCuentaContable(fblnCuentaProvision) = "" Then
36                    fblnCuentaProvision = 0
37                Else
38                    If Not fblnCuentaAfectable(fstrCuentaContable(fblnCuentaProvision), vgintClaveEmpresaContable) Then
39                        fblnCuentaProvision = 0
40                        vlblnNoAceptaMovimientos = True
41                    End If
42                End If
43            Else
44                fblnCuentaProvision = 0
45            End If
46            rs.Close
              
47            If vlblnNoAceptaMovimientos Then
48                If intblnMostrarMensaje Then
49                    MsgBox SIHOMsg(vllngMensajeNoAceptaMovimientos), vbExclamation, "Mensaje"
50                End If
51            Else
52                If fblnCuentaProvision = 0 Then
53                    If intblnMostrarMensaje Then
54                        MsgBox SIHOMsg(vllngMensaje), vbExclamation, "Mensaje"
55                    End If
56                End If
57            End If
58        End If
          
59    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaProvision" & " Linea:" & Erl()))
End Function

Public Function fblnCuentaRetencionImpuestos(inttipomovimiento As Integer, inttipoimpuesto As Integer, intblnMostrarMensaje As Boolean) As Long
1         On Error GoTo NotificaError
          Dim rs As ADODB.Recordset
          Dim vlstrparametro As String
          Dim vllngMensaje As Integer
          Dim vllngMensajeNoAceptaMovimientos As Integer
          Dim vlblnNoAceptaMovimientos As Boolean
          
          'inttipomovimiento = Tipo del movimiento (1 = Honorarios profesionales, 2 = Recibo de arrendamiento)
          'inttipoimpuesto = Tipo de impuesto (1 = ISR, 2 = IVA)
          'intblnMostrarMensaje = Mostrar mensaje si no está configurada la cuenta

2         fblnCuentaRetencionImpuestos = 0
3         vllngMensaje = 0
4         vllngMensajeNoAceptaMovimientos = 0
5         vlstrparametro = ""

6         If inttipomovimiento = 1 Then
              ' Honorarios profesionales
7             If inttipoimpuesto = 1 Then
                  'ISR
8                 vlstrparametro = "INTCTARETENCIONISRFACHONORARIO"
                  'No se ha configurado la cuenta contable para retener el ISR de facturas de honorarios profesionales, favor de verificar.
9                 vllngMensaje = 1531
                  
                  'La cuenta contable configurada para retener el ISR de facturas de honorarios profesionales no acepta movimientos, favor de verificar.
10                vllngMensajeNoAceptaMovimientos = 1539
11            Else
                  'IVA
12                vlstrparametro = "INTCTARETENCIONIVAFACHONORARIO"
                  'No se ha configurado la cuenta contable para retener el IVA de facturas de honorarios profesionales, favor de verificar.
13                vllngMensaje = 1532
                  
                  'La cuenta contable configurada para retener el IVA de facturas de honorarios profesionales no acepta movimientos, favor de verificar.
14                vllngMensajeNoAceptaMovimientos = 1540
15            End If
16        ElseIf inttipomovimiento = 2 Then
              ' Recibo de arrendamiento
17            If inttipoimpuesto = 1 Then
                  'ISR
18                vlstrparametro = "INTCTARETENCIONISRFACARRENDAMIENTO"
                  'No se ha configurado la cuenta contable para retener el ISR de facturas de arrendamiento, favor de verificar.
19                vllngMensaje = 1533
                  
                  'La cuenta contable configurada para retener el ISR de facturas de arrendamiento no acepta movimientos, favor de verificar.
20                vllngMensajeNoAceptaMovimientos = 1541
21            Else
                  'IVA
22                vlstrparametro = "INTCTARETENCIONIVAFACARRENDAMIENTO"
                  'No se ha configurado la cuenta contable para retener el IVA de facturas de arrendamiento, favor de verificar.
23                vllngMensaje = 1534
                  
                  'La cuenta contable configurada para retener el IVA de facturas de arrendamiento no acepta movimientos, favor de verificar.
24                vllngMensajeNoAceptaMovimientos = 1542
25            End If
26        Else
27            vlstrparametro = "INTCTARETENCIONSERVICIOSCLIENTE"
                  
              'No se encuentran registradas las cuentas contables de retención y provisión de servicios de clientes en los parámetros de contabilidad.
28            vllngMensaje = 1573
29        End If
          
30        vlblnNoAceptaMovimientos = False
          
31        If vlstrparametro <> "" Then
32            Set rs = frsSelParametros("CN", vgintClaveEmpresaContable, vlstrparametro)
33            If Not rs.EOF Then
34                fblnCuentaRetencionImpuestos = IIf(IsNull(rs("Valor")), 0, rs("Valor"))
                  
35                If fstrCuentaContable(fblnCuentaRetencionImpuestos) = "" Then
36                    fblnCuentaRetencionImpuestos = 0
37                Else
38                    If Not fblnCuentaAfectable(fstrCuentaContable(fblnCuentaRetencionImpuestos), vgintClaveEmpresaContable) Then
39                        fblnCuentaRetencionImpuestos = 0
40                        vlblnNoAceptaMovimientos = True
41                    End If
42                End If
43            Else
44                fblnCuentaRetencionImpuestos = 0
45            End If
46            rs.Close
              
47            If vlblnNoAceptaMovimientos Then
48                If intblnMostrarMensaje Then
49                    MsgBox SIHOMsg(vllngMensajeNoAceptaMovimientos), vbExclamation, "Mensaje"
50                End If
51            Else
52                If fblnCuentaRetencionImpuestos = 0 Then
53                    If intblnMostrarMensaje Then
54                        MsgBox SIHOMsg(vllngMensaje), vbExclamation, "Mensaje"
55                    End If
56                End If
57            End If
58        End If
          
59    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuentaRetencionImpuestos" & " Linea:" & Erl()))
End Function

Public Function fintErrorGrabar(ByRef frm As Form, ByRef strFolio As String, ByRef llngNumCorte As Long, strFecha As String, ByRef lblnEntraCorte As Boolean, ByRef llngNumFormaCredito As Long) As Integer
    '---------------------------------------------------------------------------------------------------
    ' Función que regresa un núm de error o mensaje en validaciones necesarias
    ' cuando ya inició la transacción
    '---------------------------------------------------------------------------------------------------
    Dim rsForma As New ADODB.Recordset 'Para cargar la forma de pago crédito

    fintErrorGrabar = 0

    '------------------------------------------------------------------
    '- Folio de la factura
    '------------------------------------------------------------------
'    frm.pCargaFolio 1
'    If Trim(strFolio) = "0" Then
'        fintErrorGrabar = 291 'No existen folios activos para este documento.
'        frm.blnNoFolios = True
'        Exit Function
'    End If
    '------------------------------------------------------------------
    '- Tomar corte para identificar si la factura entra en corte o se
    '- realiza una póliza separada
    '------------------------------------------------------------------
    llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
    If llngNumCorte <> 0 Then
        lblnEntraCorte = CDate(Format(frsEjecuta_SP(Str(llngNumCorte), "sp_GnSelDatosCorte")!dtmFechahora, "dd/mm/yyyy")) <= CDate(strFecha)
    Else
        fintErrorGrabar = 659 'No se encontró un corte abierto.
        Exit Function
    End If
    '------------------------------------------------------------------
    '- Si entra en corte, Validar que se pueda introducir la factura en corte
    '------------------------------------------------------------------
'    If lblnEntraCorte Then
'        fintErrorGrabar = fintErrorBloqueoCorte()
'        If fintErrorGrabar <> 0 Then
'            Exit Function
'        End If
'    End If
    '------------------------------------------------------------------
    '- Si entra en corte, Validar que exista la forma de pago CREDITO para el depto.
    '------------------------------------------------------------------
    If lblnEntraCorte Then
        vgstrParametrosSP = CStr(-1) & "|" & CStr(-1) & "|" & CStr(-1) & "|" & vgintNumeroDepartamento & "|" & CStr(-1) & "|" & "C"
        Set rsForma = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelFormaPago")
        If rsForma.RecordCount <> 0 Then
            llngNumFormaCredito = rsForma!intFormaPago
        Else
            fintErrorGrabar = 622 'No existe una forma de pago tipo crédito para este departamento.
            Exit Function
        End If
    End If
    '------------------------------------------------------------------
    '- Si NO entra en corte, Validar el periodo contable
    '------------------------------------------------------------------
    If Not lblnEntraCorte Then
        fintErrorGrabar = frm.fintErrorContable(strFecha)
    End If
End Function

Private Sub pGuardaDetalleFactura(lngidfactura As Long, ByRef frm As Form, ByRef strFolio As String, ByRef vlblnMultiempresa As Boolean, ByRef cstrCantidad4Decimales As String, intMotivosFacturaListIndex As Integer, intRetencionISR As Integer, ByRef arrTarifas() As typTarifaImpuesto, intCboTarifaListIndex As Integer, intRetencionIVA As Integer, strNumCliente As String, ByRef vgintnumemprelacionada As Integer, ByRef vlblnCuentaIngresoSaldada As Boolean, ByRef vlintBitSaldarCuentas As Long, ByRef cstrCantidad As String, intPesos As Integer, ByRef vldblTipoCambio As Double, ByRef apoliza() As TipoPoliza, strTotal As String, strIVA As String, strRetencionIVA As String, strRetencionISR As String, strDescuentos As String, tipoPago As DirectaMasiva)
    Dim intcontador As Integer
    Dim rs As New ADODB.Recordset
    Dim dblCantidad As Double
    Dim dblDescuento As Double
    Dim dblimportegravado As Double
    Dim dbldescuentogravado As Double
    Dim dblSubTotal As Double
    Dim test As Integer
    Dim dblimporteExento As Double
    Dim dbldescuentoExento As Double
    Dim strSentencia As String
    
    strSentencia = "select * from PvDetalleFactura where chrFolioFactura = '*'"
    Set rs = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
    dblimportegravado = 0
    dbldescuentogravado = 0
    dblimporteExento = 0
    dbldescuentoExento = 0
    
    dblSubTotal = 0
    ReDim apoliza(0)
    apoliza(0).lngnumCuenta = 0
    
    Dim intReduccionNoMulti As Integer
    Dim intReduccionSiMulti As Integer
    If tipoPago.intDirectaMasiva = 0 Then
        intReduccionNoMulti = 2
        intReduccionSiMulti = 1
    Else
        intReduccionSiMulti = 0
        intReduccionNoMulti = 1
    End If
    
    For intcontador = 1 To IIf(vlblnMultiempresa, frm.vsfConcepto.Rows - intReduccionSiMulti, frm.vsfConcepto.Rows - intReduccionNoMulti)
        rs.AddNew
        rs!chrfoliofactura = Trim(strFolio)
        rs!smicveconcepto = Val(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCveConcepto))
        rs!MNYCantidad = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales))
        rs!intUnidades = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCantidad), cstrCantidad4Decimales))
        rs!MNYDESCUENTO = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))
        rs!MNYIVA = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), cstrCantidad4Decimales))
        rs!chrTipo = "NO"
        
        If rs!MNYIVA <> 0 Then
            rs!mnyIVAConcepto = rs!MNYCantidad * (vgdblCantidadIvaGeneral / 100)
        End If
                
        'If vlnblnEmpresaPersonaFisica Then
            If intMotivosFacturaListIndex <> 0 Then
            
                'Retención de ISR
                If intRetencionISR = 1 Then
                    rs!MNYRETENCIONISR = (CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))) * (arrTarifas(intCboTarifaListIndex).dblPorcentaje / 100)
                End If
                
                'Retención de IVA
                If CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), cstrCantidad4Decimales)) > 0 Then
                    If intRetencionIVA = 1 Then
                        If intMotivosFacturaListIndex <> 3 Then
                            rs!MNYRETENCIONIVA = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), cstrCantidad4Decimales)) * (gdblPorcentajeRetIVA / 100)
                        Else
                            rs!MNYRETENCIONIVA = 0
                            rs!MNYRETSERVCONCEPTO = (CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))) * gdblPorcentajeRetIVA
                        End If
                    End If
                End If
            End If
        'End If
        
        rs.Update
        
        If vlblnMultiempresa Then
             vgstrParametrosSP = frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCveCargoMultiEmp) _
                            & "|" & strFolio _
                            & "|" & vgintnumemprelacionada _
                            & "|" & vgintClaveEmpresaContable _
                            & "|" & strNumCliente _
                            & "|" & flngObtieneIdentity("SEC_PVDETALLEFACTURA", 0)
               frsEjecuta_SP vgstrParametrosSP, "SP_CCINSMULTEMPCARGOS"
        End If
        
        '--Calcular el importe gravado y el descuento sobre el importe gravado
        If CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), frm.cstrCantidad)) > 0 Then
            dblimportegravado = dblimportegravado + (CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales)))
            dbldescuentogravado = dbldescuentogravado + CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))
        End If
        
        '--Calcular el importe exento y el descuento sobre el importe exento
        If CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), frm.cstrCantidad)) = 0 And frm.vsfConcepto.TextMatrix(intcontador, frm.cintColBitExento) = 1 Then
            dblimporteExento = dblimporteExento + (CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales)))
            dbldescuentoExento = dbldescuentoExento + CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))
        End If
        
        '-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Cambio para caso 8736
        'Si las cuentas de ingreso y descuento son iguales y el bitSaldarCuentas = 1
        'agrega un sólo movimiento a la póliza con el ingreso menos el descuento
        vlblnCuentaIngresoSaldada = False
        If CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCtaIngreso)) = CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCtaDescuento)) Then
            'Verifica bit pvConceptoFacturacion.bitSaldarCuentas
            vlintBitSaldarCuentas = 1
            frsEjecuta_SP Val(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCveConcepto)), "FN_PVSELBITSALDARCUENTAS", True, vlintBitSaldarCuentas
            If vlintBitSaldarCuentas = 1 Then
                '-----------------------------------'
                ' Abono para el Ingreso - Descuento '
                '-----------------------------------'
                dblCantidad = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
                dblDescuento = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
                If (dblCantidad - dblDescuento) > 0 Then
                    pLlenapoliza CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCtaIngreso)), (dblCantidad - dblDescuento), 0, apoliza()
                    vlblnCuentaIngresoSaldada = True
                ElseIf (dblCantidad - dblDescuento) < 0 Then
                    vlblnCuentaIngresoSaldada = False   'no inserta movimiento porque es mayor el descuento que el ingreso
                ElseIf (dblCantidad - dblDescuento) = 0 Then
                    vlblnCuentaIngresoSaldada = True    'no agrega movimiento en la póliza porque no hay ingreso despues del descuento, por ser iguales las cantidades
                End If
            End If
        End If
        
        If vlblnCuentaIngresoSaldada = False Then
            dblCantidad = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
            pLlenapoliza CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCtaIngreso)), dblCantidad, 0, apoliza()
            
            dblDescuento = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
            If dblDescuento <> 0 Then
                pLlenapoliza CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCtaDescuento)), dblDescuento, 1, apoliza()
            End If
        End If
    Next intcontador
    rs.Close
    
    '--Insertar datos en pvfacturaimporte
    dblSubTotal = Val(Format(strTotal, cstrCantidad4Decimales)) - Val(Format(strIVA, cstrCantidad4Decimales)) + Val(Format(strRetencionIVA, cstrCantidad4Decimales)) + Val(Format(strRetencionISR, cstrCantidad4Decimales))
    vgstrParametrosSP = lngidfactura & "|" & dblimportegravado & "|" & dblSubTotal - dblimportegravado - dblimporteExento & "|" & dbldescuentogravado & "|" & Val(Format(strDescuentos, cstrCantidad)) - dbldescuentogravado - dbldescuentoExento & "|" & dblimporteExento & "|" & dbldescuentoExento
    frsEjecuta_SP vgstrParametrosSP, "sp_PvInsFacturaImporte", True
End Sub

Private Sub pIncluyeMovimiento(vllngxNumeroCuenta As Long, vldblxCantidad As Double, vlintxTipoMovto As Integer, ByRef aPoliza2() As RegistroPoliza)
1     On Error GoTo NotificaError
          
          Dim vlblnEstaCuenta As Boolean
          Dim vllngContador As Long
              
2         If aPoliza2(0).vllngNumeroCuenta = 0 Then
3             aPoliza2(0).vllngNumeroCuenta = vllngxNumeroCuenta
4             aPoliza2(0).vldblCantidadMovimiento = vldblxCantidad
5             aPoliza2(0).vlintTipoMovimiento = vlintxTipoMovto
6         Else
7             vlblnEstaCuenta = False
8             vllngContador = 0
9             Do While vllngContador <= UBound(aPoliza2, 1) And Not vlblnEstaCuenta
10                If aPoliza2(vllngContador).vllngNumeroCuenta = vllngxNumeroCuenta And aPoliza2(vllngContador).vlintTipoMovimiento = vlintxTipoMovto Then
11                    vlblnEstaCuenta = True
12                End If
13                If Not vlblnEstaCuenta Then
14                    vllngContador = vllngContador + 1
15                End If
16            Loop
              
17            If vlblnEstaCuenta Then
18                aPoliza2(vllngContador).vldblCantidadMovimiento = aPoliza2(vllngContador).vldblCantidadMovimiento + vldblxCantidad
19            Else
20                ReDim Preserve aPoliza2(UBound(aPoliza2, 1) + 1)
21                aPoliza2(UBound(aPoliza2, 1)).vllngNumeroCuenta = vllngxNumeroCuenta
22                aPoliza2(UBound(aPoliza2, 1)).vldblCantidadMovimiento = vldblxCantidad
23                aPoliza2(UBound(aPoliza2, 1)).vlintTipoMovimiento = vlintxTipoMovto
24            End If
25        End If

26    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIncluyeMovimiento" & " Linea:" & Erl()))
End Sub

Private Sub pGuardaCredito(lngidfactura As Long, ByRef frm As Form, ByRef vlblnMultiempresa As Boolean, intMovimiento As Integer, cstrCantidad4Decimales As String, strRetencionIVA As String, strRetencionISR As String, strIVA As String, strSubtotal As String, strTotal As String, intPesos As Integer, ByRef vldblTipoCambio As Double, ByRef vlblnPagoForma As Boolean, strNumCliente As String, ByRef llngNumCtaCliente As Long, strFolio As String, ByRef llngPersonaGraba As Long, strFecha As String)
    Dim intcontador As Integer
    Dim dblMontoCredito As Double
    Dim lngDepartamento As Long
    Dim lngNumMovimiento As Long
    Dim dblCantidad As Double
    Dim dblDescuento As Double
    Dim dblIVA As Double
    Dim dblSubTotal As Double
    Dim lngCveConcepto As Long
    Dim dblRetencionISR As Double
    Dim dblRetencionIVA As Double
    Dim dblPorcentajeRetencionIVA As Double
    Dim dblPorcentajeRetencionISR As Double
    
    intcontador = 1
        
    Do While (intcontador <= IIf(vlblnMultiempresa, frm.vsfConcepto.Rows - 1, frm.vsfConcepto.Rows - 2) And intMovimiento = 1) Or (intMovimiento = 0 And intcontador <= 1)
        lngDepartamento = CLng(IIf(intMovimiento = 1, frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDeptoConcepto), vgintNumeroDepartamento))
        If intMovimiento = 1 Then
            lngCveConcepto = CLng(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColCveConcepto))
        
            dblCantidad = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales))
            dblDescuento = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))
                    
            dblIVA = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), cstrCantidad4Decimales))
            If CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) = 0 Then
                dblRetencionIVA = 0
            Else
                dblPorcentajeRetencionIVA = CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) / CDbl(Format(strIVA, cstrCantidad4Decimales))
                dblRetencionIVA = CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColIVA), cstrCantidad4Decimales)) * dblPorcentajeRetencionIVA
            End If
            
            If CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) = 0 Then
                dblRetencionISR = 0
            Else
                dblPorcentajeRetencionIVA = CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) / CDbl(Format(strSubtotal, cstrCantidad4Decimales))
                dblRetencionISR = (CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColImporte), cstrCantidad4Decimales)) - CDbl(Format(frm.vsfConcepto.TextMatrix(intcontador, frm.cintColDescuento), cstrCantidad4Decimales))) * dblPorcentajeRetencionIVA
            End If
            
            dblSubTotal = dblCantidad - dblDescuento
            dblMontoCredito = dblCantidad - dblDescuento + dblIVA - dblRetencionIVA - dblRetencionISR
        Else
            lngCveConcepto = 0
            
            dblIVA = CDbl(Format(strIVA, cstrCantidad4Decimales))
            dblRetencionISR = CDbl(Format(strRetencionISR, cstrCantidad4Decimales))
            dblRetencionIVA = CDbl(Format(strRetencionIVA, cstrCantidad4Decimales))
            dblSubTotal = CDbl(Format(strSubtotal, cstrCantidad4Decimales))
            dblMontoCredito = CDbl(Format(strTotal, cstrCantidad4Decimales))
        End If
        
        dblMontoCredito = dblMontoCredito * IIf(intPesos = 1, 1, vldblTipoCambio)
        dblIVA = dblIVA * IIf(intPesos = 1, 1, vldblTipoCambio)
        dblSubTotal = dblSubTotal * IIf(intPesos = 1, 1, vldblTipoCambio)
        dblRetencionISR = dblRetencionISR * IIf(intPesos = 1, 1, vldblTipoCambio)
        dblRetencionIVA = dblRetencionIVA * IIf(intPesos = 1, 1, vldblTipoCambio)
        
        If dblMontoCredito <> 0 And vlblnPagoForma = False Then
            vgstrParametrosSP = fstrFechaSQL(strFecha) _
                                & "|" & strNumCliente _
                                & "|" & llngNumCtaCliente _
                                & "|" & Trim(strFolio) _
                                & "|" & "FA" _
                                & "|" & dblMontoCredito _
                                & "|" & Str(lngDepartamento) _
                                & "|" & Str(llngPersonaGraba) _
                                & "|" & " " & "|" & "0" & "|" & (dblSubTotal - dblRetencionIVA) & "|" & dblIVA & "|" & dblRetencionISR & "|" & dblRetencionIVA
            lngNumMovimiento = 1
            frsEjecuta_SP vgstrParametrosSP, "SP_GNINSCREDITO", True, lngNumMovimiento
            If lngCveConcepto <> 0 Then
                'Guardar la relación entre los movimientos de créditos creados con los conceptos de la factura
                vgstrParametrosSP = CStr(lngidfactura) & "|" & CStr(lngCveConcepto) & "|" & CStr(lngNumMovimiento)
                frsEjecuta_SP vgstrParametrosSP, "sp_PvInsConceptoCredito", True
            End If
        End If
        intcontador = intcontador + 1
    Loop
End Sub

Private Sub pGuardaPolizaCorte(ByRef apoliza() As TipoPoliza, ByRef llngNumCorte As Long, ByRef llngPersonaGraba As Long, strFolio As String, ByRef vlblnPagoForma As Boolean, cstrCantidad As String, intPesos As Integer, ByRef vldblTipoCambio As Double, intMotivosFacturaListIndex As Integer, intRetencionISR As Integer, strRetencionISR As String, intRetencionIVA As Integer, strRetencionIVA As String, strTotal As String, ByRef llngNumCtaCliente As Long, ByRef vlblnEsCredito As Boolean, cstrCantidad4Decimales As String, strIVA As String, ByRef vldblTotalIVACredito As Double)
    Dim intcontador As Integer
    Dim dblTotalCliente As Double
    Dim dblIVA As Double
    
    intcontador = 0
    Do While intcontador <= UBound(apoliza(), 1)
        'pInsCortePoliza llngNumCorte, lblFolio.Caption,       "FA", aPoliza(intContador).lngNumCuenta, aPoliza(intContador).dblCantidad, aPoliza(intContador).intNaturaleza = 1
         pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", apoliza(intcontador).lngnumCuenta, apoliza(intcontador).dblCantidad, IIf(apoliza(intcontador).intNaturaleza = 1, True, False), _
        "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
        
        intcontador = intcontador + 1
    Loop
    '------------------------------------------------
    'Cargo al cliente
    If vlblnPagoForma = False Then
        dblTotalCliente = CDbl(Format(strTotal, cstrCantidad)) * IIf(intPesos = 1, 1, vldblTipoCambio)
        If dblTotalCliente > 0 Then
            'El total del cliente puede ser cero cuando se descuenta lo mismo del concepto
            'pInsCortePoliza llngNumCorte, lblFolio.Caption,"FA", llngNumCtaCliente, dblTotalCliente, 1
            pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", llngNumCtaCliente, dblTotalCliente, True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
        End If
    
        'Cuenta de ISR provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionISR = 1 And CDbl(strRetencionISR) <> 0 Then
                pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False)), CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
            End If
        End If
        
        'Cuenta de IVA provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionIVA = 1 And CDbl(strRetencionIVA) <> 0 Then
                If intMotivosFacturaListIndex = 3 Then
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(3, 1, False), fblnCuentaRetencionImpuestos(3, 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                Else
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                End If
            End If
        End If
    
        '------------------------------------------------
        'IVA no cobrado
        dblIVA = CDbl(Format(strIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio)
        If dblIVA <> 0 Then
           'pInsCortePoliza llngNumCorte, lblFolio.Caption,"FA", glngCtaIVANoCobrado, dblIVA, 0
           pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", glngCtaIVANoCobrado, dblIVA, False, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
        End If
    Else
        If vldblTotalIVACredito > 0 Then 'Iva no cobrado
           'lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, glngCtaIVANoCobrado, vldblTotalIVACredito * IIf(optPesos(0).Value, 1, vldblTipoCambio), 0)
            pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", glngCtaIVANoCobrado, vldblTotalIVACredito * IIf(intPesos = 1, 1, vldblTipoCambio), False, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
        End If
        
        'Iva cobrado
        If Format(CDbl(Format(strIVA, cstrCantidad)), "##########.00") - Format(vldblTotalIVACredito, "##########.00") > 0.01 Then
           'lngNumDetalle = flngInsertarPolizaDetalle(llngNumPoliza, glngCtaIVACobrado, (Format(CDbl(Format(lblIVA.Caption, cstrCantidad)), "##########.00") - Format(vldblTotalIVACredito, "##########.00")) * IIf(optPesos(0).Value, 1, vldblTipoCambio), 0)
           pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", glngCtaIVACobrado, _
           (Format(CDbl(Format(strIVA, cstrCantidad)), "##########.00") - Format(vldblTotalIVACredito, "##########.00")) * IIf(intPesos = 1, 1, vldblTipoCambio), _
           False, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
        End If
        
        'Cuenta de ISR provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionISR = 1 And CDbl(strRetencionISR) <> 0 Then
                pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, 2), 1, False)), CDbl(Format(strRetencionISR, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
            End If
        End If
        
        'Cuenta de IVA provisionado
        If intMotivosFacturaListIndex <> 0 Then
            If intRetencionIVA = 1 And CDbl(strRetencionIVA) <> 0 Then
                If intMotivosFacturaListIndex = 3 Then
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(3, 1, False), fblnCuentaRetencionImpuestos(3, 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                Else
                    pAgregarMovArregloCorte llngNumCorte, llngPersonaGraba, Trim(strFolio), "FA", IIf(vlblnEsCredito, fblnCuentaProvision(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), IIf(intMotivosFacturaListIndex = 3, 1, 2), False), fblnCuentaRetencionImpuestos(IIf(intMotivosFacturaListIndex = 1, 1, IIf(intMotivosFacturaListIndex = 2, 2, 3)), 2, False)), CDbl(Format(strRetencionIVA, cstrCantidad4Decimales)) * IIf(intPesos = 1, 1, vldblTipoCambio), True, "", 0, 0, "", 0, 2, Trim(strFolio), "FA"
                End If
            End If
        End If
    End If
End Sub

Public Sub pCancelarFacturaDirecta(frm As Form, lngNumCliente As Long, strFolioFactura As String, blnEntroAlCorte As Boolean, vllngPersonaGraba As Long, strFecha As String, Optional lngCorte As Long = 0, Optional lngPoliza As Long = 0)
          Dim llngNumCorte As Long
          Dim vllngCorteGrabando As Long
          Dim intErrorCancelar As Integer
          
          'Variables crédito a facturar
          Dim rsCCDatosPolizaCF As ADODB.Recordset
          Dim vlstrsql As String
              
1     On Error GoTo NotificaError
2         intErrorCancelar = 0
          
3         EntornoSIHO.ConeccionSIHO.BeginTrans
          
          
          
4         If blnEntroAlCorte Then
5            llngNumCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                
6            If llngNumCorte = 0 Then
7                intErrorCancelar = 659
8            Else
                 'Bloqueo del corte
9                intErrorCancelar = fintErrorBloqueoCorte(llngNumCorte)
10           End If
11        Else
12           intErrorCancelar = fintErrorContable(strFecha)
13        End If
           
14        If intErrorCancelar = 0 Then
             '1.- Cancela la factura y registrar en documentos cancelados
15           vgstrParametrosSP = Trim(strFolioFactura) & "|" & Str(vgintNumeroDepartamento) & "|" & Str(vllngPersonaGraba)
16           frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaFactura", True
             
             '2.- Cancelar los créditos
17           vgstrParametrosSP = Trim(strFolioFactura) & "|" & "FA"
18           frsEjecuta_SP vgstrParametrosSP, "Sp_CcUpdCancelaCredito", True
              
19           vgstrParametrosSP = Trim(strFolioFactura) & "|" & vgintClaveEmpresaContable & "|" & CStr(lngNumCliente)
20           frsEjecuta_SP vgstrParametrosSP, "SP_CCDELMULTEMPCARGOS", True
              
             '3.- Si la factura no entró al corte, cancelar la póliza
21           If Not blnEntroAlCorte Then
                 ' Cancelar la póliza
22                pCancelarPoliza CLng(lngPoliza), "CANCELACION DE FACTURA " & Trim(strFolioFactura) & " (REUTILIZAR POLIZA) "
                  
                  ' Liberar para que se realicen cierres
23                pEjecutaSentencia "UPDATE CnEstatusCierre SET vchEstatus = 'Libre' WHERE tnyClaveEmpresa = " + Str(vgintClaveEmpresaContable)
24           Else
                  ' Cancelar el documento en el corte
25                vgstrParametrosSP = Trim(strFolioFactura) & "|" & vllngPersonaGraba & "|" & "FA" & "|" & CStr(lngCorte) & "|" & Str(llngNumCorte)
26                frsEjecuta_SP vgstrParametrosSP, "Sp_PvUpdCancelaDoctoCorte", True
              
                  ' Liberar el corte
27                pLiberaCorte llngNumCorte
28           End If
             '4.- Si se agregaron creditos a facturar, se cancela la poliza
             vlstrsql = "SELECT DISTINCT(INTNUMEROPOLIZACANCELACION) NUMEROPOLIZA FROM CCMOVIMIENTOCREDITOPOLIZA WHERE CHRFOLIOFACTURA = '" + Trim(strFolioFactura) + "'"
             Set rsCCDatosPolizaCF = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
             If rsCCDatosPolizaCF.RecordCount > 0 Then
                  vgstrParametrosSP = Trim(strFolioFactura) & "|" & rsCCDatosPolizaCF!NumeroPoliza
                  frsEjecuta_SP vgstrParametrosSP, "sp_CCUpdCancelaPolCredaFact", True
             End If
             '---------------------------------
 
              
29            pGuardarLogTransaccion frm.Name, EnmGrabar, vllngPersonaGraba, "CANCELACIÓN DE FACTURA DIRECTA", strFolioFactura
30            EntornoSIHO.ConeccionSIHO.CommitTrans
31        Else
32            EntornoSIHO.ConeccionSIHO.RollbackTrans
33            MsgBox SIHOMsg(intErrorCancelar), vbCritical + vbOKOnly, "Mensaje"
34        End If
35    Exit Sub
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCancelarFacturaDirecta" & " Linea:" & Erl()))
        frm.cmdSave.Enabled = False
        frm.lblnConsulta = False
        Unload frm
End Sub

Public Function fblnPermitirEnvio(strNumCliente As String) As Boolean
1     On Error GoTo NotificaError
          
          Dim rs As New ADODB.Recordset
          Dim strSentencia As String
          
2         fblnPermitirEnvio = False
          
          '- Revisar que el parámetro de envío de CFD esté activado -'
3         If fblnRevisaEnvioCorreo(vgintClaveEmpresaContable) Then
4             fblnPermitirEnvio = True
5         Else
6             Exit Function
7         End If

          '- Revisar que el cliente o paciente no pertenezcan a una empresa -'
          'Set rs = frsEjecuta_SP(Trim(txtNumCliente.Text), "SP_PVSELTIPOCLIENTE")
8         strSentencia = "SELECT chrTipoCliente FROM CcCliente WHERE intNumCliente = " & Trim(strNumCliente)
9         Set rs = frsRegresaRs(strSentencia, adLockOptimistic, adOpenDynamic)
10        If rs.RecordCount <> 0 Then
11            If Trim(rs!chrTipoCliente) <> "CO" Then
12                fblnPermitirEnvio = True
13            Else
14                fblnPermitirEnvio = False
15            End If
16        End If
          
17    Exit Function
NotificaError:
        Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnPermitirEnvio" & " Linea:" & Erl()))
End Function

Public Function fblnValidaPrecioMaxPub(StrCveArticulo As String, dblPrecio As Double) As Boolean
    Dim rs As ADODB.Recordset
    If dblPrecio = 0 Then
        Set rs = frsEjecuta_SP(StrCveArticulo, "sp_PVSelCveListaArtIAPMP")
        If Not rs.EOF Then
            fblnValidaPrecioMaxPub = False
        Else
            fblnValidaPrecioMaxPub = True
        End If
        rs.Close
    Else
        fblnValidaPrecioMaxPub = True
    End If
End Function
