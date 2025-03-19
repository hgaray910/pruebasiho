Attribute VB_Name = "modAlmacen"
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Almacén
'| Nombre del Módulo        : modAlmacén.bas
'-------------------------------------------------------------------------------------
'| Objetivo: Permite crear constantes, banderas y variables públicas que serán utilizadas en el
'| proyecto de admisión.
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Nery Lozano - Luis Astudillo
'| Autor                    : Nery Lozano - Luis Astudillo
'| Fecha de Creación        : 11/Noviembre/1999
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------
Option Private Module
Option Explicit

Public vglngPersonaGraba As Long                'Para el procedimiento pPersonaGraba
Public vgstrGrabaMedicoEmpleado As String       'Para el procedimiento pPersonaGraba
Public First_time As Boolean
Public vgblnTerminate As Boolean
'Public vllngPersonaGraba As Long 'Empleado recibe

Public Declare Function GetComputerName Lib "kernel32" _
Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As _
Long) As Long

Global vgstrParametrosSP As String

Public vgblnFlagGrd As Boolean

Public Type departamento
    vlintDepartamento As Integer
End Type

Public Type MenuPermiso
    vlintNumMenu As Integer
    vlstrNombreMenu As String
    vlstrNombreBoton As String
    vlstrPermisos As String
End Type

Public Type Permiso
    vllngNumeroOpcion As Long
    vlstrTipoPermiso As String
End Type

Public Type CtrlPermiso
    vlstrObjeto As String
    vlstrForma As String
    vlIntOrdenTab As Long
    vlstrTipoPermiso As String
    vlintTotalRegistros As Integer
End Type

'Caja
Public Type FormasPago
    vlintNumFormaPago As Integer
    vldblCantidad As Double
    vllngFolio As Double
    vldblComision As Double
    vllngCuentaContable As Long
    vldblTipoCambio As Double
    vlbolEsCredito As Boolean
    vldblDolares As Double
End Type

'Boletines
Public Type DatosBoletines
    Lote As String
    Url As String
    Descripcion As String
    FechaCaduca As Date
End Type
Public aBoletines() As DatosBoletines

Public Enum TipoOperacion
        EnmLogIn = 1
        EnmGrabar = 2
        EnmBorrar = 3
        EnmCambiar = 4
        EnmCancelacion = 5
        EnmLogout = 6
        EnmReImpresion = 7
        EnmConsulta = 8
End Enum

Public vgLngTotalLotesxArt As Long
Public vglngTotalCantLotexArt As Long
Public vglngTotalCantUVLotexArt As Long
Public vglngTotalCantUMLotexArt As Long
Public vgblnManejaLote As Boolean

Global vgstrBaseDatosUtilizada As String  'Variable para distinguir que base de datos es: "MSSQL" y "ORACLE"
Public vgblnEsTallerIV As Boolean 'Variable para saber si el almacén donde funciona el sistema es un taller
Public vglngCveAlmacenGeneral As Long 'Indica el numero del proveedor asignado como Almacen General o Abastecimiento

'--------------------------------------------------------------
'06/Febrero/2003 Con la actualización del catálogo externos y tarjetas de cme
'--------------------------------------------------------------
Public vgOrden As Byte
Public vglngNumeroPaciente As Long
Public vgstrNombrePaciente As String
Public vglngNumeroCuenta As Long
Public vlY As Long              'Variable que contiene el numero de renglon actual en la impresion
Public vYCol As Long            'Almacena el renglon para inicializar la impresion en caso de que sean mas de una columna
Public vSigueCol As Boolean     'Valida si ya termino de imprimir la segunda columna

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' S E G U R I D A D
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public vgintNumeroModulo As Integer  'Numero de módulo de la tabla Modulo
Public vglngNumeroEmpleado As Long  'Número de empleado logueado
Public vgintNumeroDepartamento As Integer  'Numero de departamento con que se logueo el empleado
Public vglngNumeroLogin As Long  'Numero de login que le corresponde en la tabla Login
Public vgstrNombreUsuario As String  'Login del usuario
Public vgstrNombreDepartamento As String  'Nombre del departamento personalizado con que se loguea el empleado
Public aPermisos() As Permiso  'Arreglo para seguridad
Public aControlesModulo() As CtrlPermiso 'Arreglo para cargar todos los controles del modulo.
Public agMenuPermisos() As MenuPermiso 'Arreglo que contiene permisos del usuario por proceso o menu

'Constantes
Public Const cgstrModulo As String = "IV" 'Constante para registrar el nombre del módulo en manejo de errores

Public Const cintNumOpcionListas = 1034

'Banderas
Public vgblnExistioError As Boolean 'Bandera que permite salir del formulario
Public vlblnExistioError As Boolean 'Bandera que permite salir del formulario
Public vgblnHuboCambio As Boolean 'Bandera para controlar si existe algun cambio y confirmar el mismo
Public vgblnErrorIngreso As Boolean 'Bandera para verificar si existe un error al ingreso de datos como no se ha escrito nada en una variable
Public vgblnContrasenaCorr As Boolean 'Bandera para verificar si la contrasena tecleada fue la correcta
Public vgstrEmpleado As String 'Guarda el nombre del encargado del departamento de Administración

'Variables
Public vgintColOrd As Integer 'Contiene el número de la columna a ordenar
Public vgintTipoOrd As Integer 'Contiene el tipo de ordenación
' variables para cambiar mediante drag & drop el orden de las columnas del grid
Public vgblnArrastrarOk As Boolean 'Variable para identificar si se ha realizado un drag & drop
Public vgintColArrastra As Integer 'Variable para identificar que columna se esta desplazando en
Public vgstrNombreForm As String 'Indica el nombre del formulario en el que ocurre el error
Public vgstrNombreProcedimiento As String 'Indica el Nombre del procedimiento donde ocurre el error
Public vgintColLoc As Integer 'Variable que contiene la columna que va a ordenar al buscar dentro del MshFlexGrid
Public vgintMousePosXdn As Integer, vgintMousePosYdn As Integer 'Variables que contienen la posicion de X y Y cuando se utiliza el ratón en el grid
Public vgstrAcumTextoBusqueda As String 'Acumula los caracteres capturados para la búsqueda en el grdHBusqueda
Public vgstrPasarCve As String 'Pasa el valor de una clave de un formulario a otro
Public vgstrCveSubfamilia As String 'Pasa el valor de la subfamilia de la forma de MantoSubfamilia a la forma de MantoArticulo
Public vgintCveDepartamento As Integer 'Pasa el valor del almacen para la localizacion en el mantenimiento de existencias

'Variables globales para trabajar con cualquier formulario, procedimiento, o donde se lo aplique como parametros de entrada o salida
Public vgstrVarIntercam As String 'Variable de intercambio de información entre procedimientos
Public vgstrVarIntercam2 As String 'Variable de intercambio de información entre procedimientos

Public vglngCuentaPaciente As Long

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
Global vgblnNomina As Boolean                   'Si la empresa va a utilizar el modulo de nomina
Public vgdblCantidadIvaGeneral As Double  'Porcentaje IVA general registrada en Parametros
Public vgintDiasRequisicion As Integer          'Número de dias anteriores sobre los cuales se revisaran las requisiciones
Public vgblnCorteEntregaRecepcion As Boolean    'Indica si se manejan cortes de entrega recepción

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
' Variables del modulo de Nomina
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
Public vglngCtaComisiones As Long 'Cuenta para comisiones por uso de ciertas formas de pago
Public vglngCtaIVAacreeditable As Long 'Cuenta para el IVA acreeditable en comisiones por uso de tarjetas
Public vgdblIVAGeneral  As Double 'Porcentaje de IVA general
Public vglngCuentaCajaChica As Long 'Cuenta contable de caja chica
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
'Variables de la Recepcion de cargos directos y de artículos
Public vgintCantOrdenada As Integer 'Pasa el valor de la cantidad de artículos que tiene la orden de cargos directos
Public vgintCantRecibida As Integer 'Pasa el valor de la cantidad de artículos recibidos de cargos directos
Public vgintCostoOrdenado As Currency 'Pasa el valor del costo que se especifica en la orden de cargos directos
Public vgintCostoRecibido As Currency 'Pasa el valor del costo con que se recibió el artículo de cargo directo
Public vgintDesctoOrdenado As Currency  'Pasa el valor del descuento que se especifica en la orden de cargos directos
Public vgintDesctoRecibido As Currency 'Pasa el valor del descuento con que se recibió el artículo de cargo directo
Public vgintmaximopermitido As Integer 'Pasa el valor de la cantidad total de artículos recibidos de cargos directos
Public vgblnConsulta    'Pasa el valor para indicar si se trata de una modificación o de una recepción nueva
Public vgbitFaltante As Byte 'Pasa el valor si es que detectó un faltante
Public vgbitDanado As Byte 'Pasa el valor si es que detectó que se rechazó el artículo por estar danado
Public vgdtmFechaCaducidad As Date 'Guarda el valor de caducidad en la recepcion de articulos
Public vgintPrecioVenta As Currency 'Guarda el valor del precio de venta del articulo recibido
Public vgstrDescripcion As String 'Pasa el valor de la descripción larga del artículo de cargo directo
Public vgblnConfirmaCargo As Boolean 'Confirma que se van a realizar los cargos
Public vgstrCveArtRecep As String  ' Es el arículo que se recibe

Public vgstrNombreEmpresaContable As String 'Nombre empresa
Public vgstrNombreCortoEmpresaContable As String 'Nombre corto empresa
Public vgstrRepresentanteLegalEmpresaContable As String 'Representante legal
Public vgstrRFCEmpresaContable As String 'RFC
Public vgstrDireccionEmpresaContable As String 'Direccion
Public vgstrTelefonoEmpresaContable As String 'Telefono
Public vglngCuentaResultadoEjercicio As Long 'Cuenta para mostrar el resultado del ejercicio
Public vgintEjercicioInicioOperaciones As Integer 'Ejercicio en el cual se comienzan operaciones
Public vgintMesInicioOperaciones As Integer 'Mes en el cual se comienzan operaciones


'Para trabajar con el formulario frmLista
Public vgstrDataMember As String 'Para saber de que objeto comando traer y revisar la informacion
Public vgstrNombreCbo As String 'Para saber cual es el nombre del combo y llenar la lista con el field apropiado
Public vgstrNomFamilia As String
Public vgstrNomSubFamilia As String
Public vgstrTipoArtMedIns As String

'Devoluciones de Almacenes
Public vgstrFechaDevolucion As Date
Public vgintNumDev As Long
Public vgintNumRef As Long
Public vgstrDeptoDevuelve As String
Public vgstrDeptoRecibe As String
Public vgstrEmpleadoRecibe As String

'Parametros de Cuentas por pagar
Public vgintCuentaFlete As Long
Public vgintClaveEmpresaContable As Integer  'Clave de empresa registrada en Parametros
Public vgstrEstructuraCuentaContable As String  'Estructura de la cuenta contable registrada en Parametros

'Para trabajar con el formulario frmBusquedaPaciente con paso de parametros
Public vgStrTipoPaciente As String          'Tipo de paciente <I>Interno  <E>Externo
Public vgdblNumeroCuenta As Double          'Numero de cuenata de paciente, sea I o E
Public vgdblNumeroPaciente As Double        'Numero de de paciente
Public vgdtmFechaIngreso As Date            'Fecha de Ingreso
Public vgblnRecienNacido As Boolean         'Fecha de Ingreso
Public vgintCveTipoRecNac                   'Numero de tipo de paciente recien nacido
Public vgintCveTipoPacCon                   'Numero de tipo de paciente convenio
Public vgstrNumeroCuarto As String          'Número Cuarto
Public vgintTipoPacienteAdm As Integer      'Numero del tipo de paciente Convenio, particulas, accionista,..
Public vgstrTipoPacienteAdm As String       'Descripcion del tipo de paciente Convenio, particular, accionista,..

Public vgbitApricacionMed As Boolean 'Se usa en la forma de salidas de cargos a pacientes
'Public vgstrFormatoClaveArt As String ' Se usa para darle formato a la clave del artículo
Public vgintLongNumArt As Integer ' Se usa pasa saber cuantas posiciones ocupa el consecutivo de artículo
Public vgintTipoConvenioPac As Integer      'Número del convenio del paciente
Public vgstrTipoConvenioPac As String       'Descripcion del convenio del paciente
Public vgintNumeroEmpresa As Long           'Numero de empresa del paciente
Public vgstrNombreEmpresa As String         'Numero de empresa del paciente
Public vgstrEdad As String                  'Edad del paciente
Public vgstrNomArea As String               'Para cargar el nombre del area
Public vgintCveArea As Integer              'Para carga el numero de area
Public vgstrNumeroAfiliacion As String      'Numero de afiliacion del paciente
Public vgstrSexoPaciente As String          'Sexo del paciente "M"asculino o "F"emenino


Public Sub pDescuentaUbicacion(vlintCveDepto As Integer, vlstrCveArticulo As String, vllngCantidad As Long, vlstrUnidad As String)
  '---------------------------------------------------------------------------
  ' Procedimiento para descontar de Ivubicacion en unidad alterna o minima
  ' vlstrUnidad = "A" alterna
  ' vlstrUnidad = "M" minima
  '---------------------------------------------------------------------------
  Dim vlstrx As String
  Dim vllngContenido As Long
  Dim vllngExistenciaUM As Long
  Dim vllngExistenciaUV As Long
  Dim rs As New ADODB.Recordset
  
  If vlstrUnidad = "A" Then
    vlstrx = "" & _
      "update IvUbicacion set " & _
        "intExistenciaDeptoUV = intExistenciaDeptoUV - " & Str(vllngCantidad) & " " & _
       "Where " & _
        "smiCveDepartamento=" & Str(vlintCveDepto) & " and " & _
        "chrCveArticulo='" & Trim(vlstrCveArticulo) & "'"
  Else
    vllngContenido = frsRegresaRs("SELECT intContenido FROM IvArticulo WHERE chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0)
    Set rs = frsRegresaRs("SELECT intExistenciaDeptoUM,intExistenciaDeptoUV FROM IvUbicacion WHERE smiCveDepartamento=" & Str(vlintCveDepto) & " AND chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ")
    vllngExistenciaUM = rs!intexistenciadeptoum
    vllngExistenciaUV = rs!intExistenciaDeptouv
    rs.Close
    vlstrx = "" & _
      "UPDATE IvUbicacion SET " & _
      "  intExistenciaDeptoUM= CAST((" & _
      "  (" & vllngExistenciaUV & " * " & vllngContenido & " + " & vllngExistenciaUM & " - " & Str(vllngCantidad) & ")/" & vllngContenido & " - " & _
      "  CAST((" & vllngExistenciaUV & " * " & vllngContenido & " + " & vllngExistenciaUM & " - " & Str(vllngCantidad) & ")/" & vllngContenido & " AS INT)) * " & vllngContenido & " AS INT)," & _
      "  intExistenciaDeptoUV= CAST((" & vllngExistenciaUV & " * " & vllngContenido & " + " & vllngExistenciaUM & " - " & Str(vllngCantidad) & ")/" & vllngContenido & " AS INT) " & _
      " WHERE " & _
      "  smiCveDepartamento=" & Str(vlintCveDepto) & " AND " & _
      "  chrCveArticulo='" & Trim(vlstrCveArticulo) & "' "
  End If
  pEjecutaSentencia vlstrx
End Sub


Public Function flngExistenciaTotal(vlintDepto As Integer, vlstrCveArticulo As String, vlstrUnidad As String)
  '---------------------------------------------------------------------------------'
  ' Función que regresa la existencia de un departamento en unidad minima o alterna
  ' vlstrUnidad='M' minima
  ' vlstrUnidad='A' alterna
  '---------------------------------------------------------------------------------'
  Dim vlstrx As String
  Dim vldblContenido As Double
  Dim vldblExistenciaUM As Double
  Dim vldblExistenciaUV As Double
  Dim vllngExistenciaTotal As Long

  vldblContenido = 0
  vldblExistenciaUM = 0
  vldblExistenciaUV = 0
  vllngExistenciaTotal = 0
  If Trim(vlstrUnidad) <> "A" Then
    vldblContenido = CDbl(frsRegresaRs("select intContenido from ivArticulo where chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0))
    vldblExistenciaUM = CDbl(frsRegresaRs("select intExistenciaDeptoUM from IvUbicacion where smiCveDepartamento=" & Str(vlintDepto) & " and chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0))
    vldblExistenciaUV = CDbl(frsRegresaRs("select intExistenciaDeptoUV from IvUbicacion where smiCveDepartamento=" & Str(vlintDepto) & " and chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0))
    vllngExistenciaTotal = CLng(vldblExistenciaUV * vldblContenido + vldblExistenciaUM)
  Else
    vllngExistenciaTotal = frsRegresaRs("SELECT intExistenciaDeptoUV FROM IvUbicacion WHERE smiCveDepartamento=" & Str(vlintDepto) & " AND chrCveArticulo='" & Trim(vlstrCveArticulo) & "' ").Fields(0)
  End If
  flngExistenciaTotal = vllngExistenciaTotal
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
    
    frmBusquedaCuentasContables.vlblnTodasCuentas = vlblnTodasCuentas
    frmBusquedaCuentasContables.vlintcveempresa = intclaveempresa
    frmBusquedaCuentasContables.vlintLvl = vlintLvl
    frmBusquedaCuentasContables.Show vbModal
    flngBusquedaCuentasContables = frmBusquedaCuentasContables.vllngNumeroCuenta

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngBusquedaCuentasContables"))
End Function

'Public Sub pCargaLotes(vlintNumMov As Long, vlstrReferencia As String, vlstrTabla As String, vlintDeptLote As Integer, vlblnLimpiarArreglo As Boolean, Optional vlstrFiltros As String)
'' Para manejo de caducidad de articulos
'' vlstrTabla para saber si es la tabla 'O' Original o la 'T' Temporal(para inventario fisico)
'' si vlstrTabla es Temporal, VlintNumMov es contiene clave de departamento
'   Dim rslotes As New ADODB.Recordset
'   Dim vlintCont As Long
'   Dim vlstrx As String
'   Dim vlintContInicial As Integer
'   Dim vlparam As String
'
'    If vlblnLimpiarArreglo Then
'        ReDim agLotes(0)
'        vgLngTotalLotesxMov = 0
'    End If
'
'    vlparam = "'" & Trim(vlstrTabla) & "'|'" & Trim(vlstrReferencia) & "'|" & vlintNumMov
'    If vlstrTabla = "T" Then
'        vlparam = vlparam & "|" & Trim(vlstrFiltros)
'    Else
'        vlparam = vlparam & "|" & vlintDeptLote & "|||||||"
'    End If
'
'    Set rslotes = frsEjecuta_SP(vlparam, "SP_IVSELKARDEXMOVLOTE", , , , True)
'
'    If rslotes.RecordCount <> 0 Then
'
'        If vlblnLimpiarArreglo Then
'            vlintContInicial = 1
'            vgLngTotalLotesxMov = rslotes.RecordCount
'        Else
'            vlintContInicial = vgLngTotalLotesxMov + 1
'            vgLngTotalLotesxMov = vgLngTotalLotesxMov + rslotes.RecordCount
'        End If
'
'        rslotes.MoveFirst
'        For vlintCont = vlintContInicial To vgLngTotalLotesxMov
'            ReDim Preserve agLotes(vlintCont)
'            agLotes(vlintCont).Articulo = rslotes!chrCveArticulo
'            agLotes(vlintCont).CantidadUM = rslotes!RELCANTIDADUM
'            agLotes(vlintCont).CantidadUV = rslotes!RELCANTIDADUV
'            agLotes(vlintCont).CantidadUMInicio = rslotes!RELCANTIDADUM
'            agLotes(vlintCont).CantidadUVInicio = rslotes!RELCANTIDADUV
'            agLotes(vlintCont).FechaCaducidad = rslotes!DTMFECHACADUCIDAD
'            agLotes(vlintCont).Lote = rslotes!chrlote
'            agLotes(vlintCont).TablaRelacion = rslotes!vchtablarelacion
'            agLotes(vlintCont).TipoAccion = rslotes!inttipoaccion
'            rslotes.MoveNext
'        Next vlintCont
'    End If
'End Sub

Public Sub pGrabaLotesTmp(vllngNumMov As Long, vlstrReferencia As String, vlintNumDepto As Integer)
'Para manejo de caducidad de articulos solo en la toma de inventarios fisicos
'vllngNumMov es el departamento
Dim rslotes As New ADODB.Recordset
'Dim rsCatLote As New ADODB.Recordset
Dim vlstrx As String
Dim vllngContad As Long

    vlstrx = "Select * from ivKardexInventarioLotetmp where 1 = 2 "    'Solo para llenarlo
    Set rslotes = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)

    If vgLngTotalLotesxMov > 0 Then
        For vllngContad = 1 To vgLngTotalLotesxMov
            If agLotes(vllngContad).Borrado <> "*" Then
                With rslotes
                    .AddNew
                    !numnumreferencia = vllngNumMov
                    !chrcvearticulo = agLotes(vllngContad).Articulo
                    !chrlote = agLotes(vllngContad).Lote
                    !dtmfechahoracaduce = agLotes(vllngContad).fechaCaducidad
                    !RELCANTIDADUM = agLotes(vllngContad).CantidadUM
                    !RELCANTIDADUV = agLotes(vllngContad).CantidadUV
                    !vchtablarelacion = Trim(vlstrReferencia)
                    !smicvedepartamento = vlintNumDepto
                    !inttipoaccion = agLotes(vllngContad).TipoAccion
                    .Update
                End With
                'GRABA LOTE
'                Set rsCatLote = frsRegresaRs("Select * from IVLOTES where Rtrim(LTRIM(chrlote)) = '" & Trim(agLotes(vllngContad).Lote) & "' AND CHRCVEARTICULO = '" & Trim(agLotes(vllngContad).Articulo) & "'", adLockOptimistic, adOpenDynamic)
'                If rsCatLote.RecordCount <= 0 Then
'                    rsCatLote.AddNew
'                    rsCatLote!chrlote = Trim(agLotes(vllngContad).Lote)
'                    rsCatLote!chrCveArticulo = Trim(agLotes(vllngContad).Articulo)
'                    rsCatLote!DTMFECHACADUCIDAD = agLotes(vllngContad).FechaCaducidad
'                    rsCatLote.Update
'                End If
            End If
        Next vllngContad
    End If
    
End Sub

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
