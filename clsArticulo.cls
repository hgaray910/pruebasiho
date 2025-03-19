VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private lngIdArticulo As Long
Private StrCveArticulo As String
Private strNombreComercial As String
Private lngIdGenerico As Long
Private strNombreGenerico As String
Private dblCostoUltEntrada As Double
Private dtmFechaUltEntrada As Date
Private strEstatus As String
Private strCostoGasto As String
Private lngCveConceptFact As Long
Private lngCveUniAlternaVta As Long
Private lngCveUniMinimaVta As Long
Private lngContenido As Long
Private lngCveConceptFact2 As Long
Private blnVentaPublico As Boolean
Private strCveArtMedicamen As String
Private strCveFamilia As String
Private strCveSubfamilia As String
Private strNumArticulo  As String
Private blnManejaCaducidad As Boolean
Private lngIdTipoControlado As Long
Private dblCostoMasAlto As Double
Private strFechaCostoMasAlto As String
Private dblPrecioMaximoPublico As Double
Private intcveVia As Integer
Private intcveMedida As Integer
Private intFrecuencia As Integer
Private intcveTiempo As Integer
Private intCuadroBasico As Integer
Private strDosisMax As String
Private blnAntibiotico As Boolean
Private blnLASA As Boolean
Private intcveIEPSvl As Integer
Private blnIndicaciones As Boolean
Private intExcluirDescuento As Integer
Private dtmFechaAltaArticulos As String
Private blnUnidosis As Boolean
Public Enum SegmentoArticulo
    saTipoArticulo = 1
    saFamilia = 2
    saSubFamilia = 3
    saConsecutivo = 4
End Enum

'----------------------------------------------------------------------------------------------------
'||           P r o p i e d a d e s     d e     l e c t u r a  /  e s c r i t u r a                ||
'----------------------------------------------------------------------------------------------------

' used when assigning a value to the property, on the left side of an assignment.
' Syntax: X.lngCveArticulo = 5
Public Property Let idArticulo(ByVal vData As Long)
    lngIdArticulo = vData
    pCargaDatosArticulo vData, ""
End Property

' used when retrieving value of a property, on the right side of an assignment.
' Syntax: Debug.Print X.lngCveArticulo
Public Property Get idArticulo() As Long
    idArticulo = lngIdArticulo
End Property

Public Property Let claveArticulo(ByVal vData As String)
    StrCveArticulo = vData
    pCargaDatosArticulo -1, vData
End Property

Public Property Get claveArticulo() As String
    claveArticulo = StrCveArticulo
End Property
'----------------------------------------------------------------------------------------------------
'||               P r o p i e d a d e s     d e     s o l o      l e c t u r a                     ||
'----------------------------------------------------------------------------------------------------
Public Property Get NombreComercial() As String
    NombreComercial = strNombreComercial
End Property

Public Property Get IdGenerico() As Long
    IdGenerico = lngIdGenerico
End Property

Public Property Get IdTipoControlado() As Long
    IdTipoControlado = lngIdTipoControlado
End Property

Public Property Get NombreGenerico() As String
    NombreGenerico = strNombreGenerico
End Property

Public Property Get CostoUltEntrada() As Double
    CostoUltEntrada = dblCostoUltEntrada
End Property

Public Property Get FechaUltEntrada() As Date
    FechaUltEntrada = dtmFechaUltEntrada
End Property
Public Property Get FechaDeAlta() As String
    FechaDeAlta = dtmFechaAltaArticulos
End Property
Public Property Get Status() As String
    Status = strEstatus
End Property

Public Property Get CostoGasto() As String
    CostoGasto = strCostoGasto
End Property

Public Property Get CveConceptFact() As Long
    CveConceptFact = lngCveConceptFact
End Property

Public Property Get CveUniAlternaVta() As Long
    CveUniAlternaVta = lngCveUniAlternaVta
End Property

Public Property Get CveUniMinimaVta() As Long
    CveUniMinimaVta = lngCveUniMinimaVta
End Property

Public Property Get Contenido() As Long
    Contenido = lngContenido
End Property

Public Property Get CveConceptFact2() As Long
    CveConceptFact2 = lngCveConceptFact2
End Property

Public Property Get EsVentaPublico() As Boolean
    EsVentaPublico = blnVentaPublico
End Property

Public Property Get MostarIndicaciones() As Boolean
    MostarIndicaciones = blnIndicaciones
End Property

Public Property Get CveArtMedicamen() As String
    CveArtMedicamen = strCveArtMedicamen
End Property

Public Property Get CveFamilia() As String
    CveFamilia = strCveFamilia
End Property

Public Property Get CveSubFamilia() As String
    CveSubFamilia = strCveSubfamilia
End Property

Public Property Get NumArticulo() As String
    NumArticulo = strNumArticulo
End Property

Public Property Get ManejaCaducidad() As Boolean
    ManejaCaducidad = blnManejaCaducidad
End Property

Public Property Get CodigoBarras() As String
    CodigoBarras = fstrDameCodigoBarras
End Property

Public Property Get CostoMasAlto() As Double
    CostoMasAlto = dblCostoMasAlto
End Property

Public Property Get FechaCostoMasAlto() As String
    FechaCostoMasAlto = Format(strFechaCostoMasAlto, "dd/mm/yyyy")
End Property

Public Property Get PrecioMaximoPublico() As Date
    PrecioMaximoPublico = dblPrecioMaximoPublico
End Property
Public Property Get CveVia() As Integer
    CveVia = intcveVia
End Property
Public Property Get CveUnidadMedida() As Integer
    CveUnidadMedida = intcveMedida
End Property
Public Property Get Frecuencia() As Integer
    Frecuencia = intFrecuencia
End Property
Public Property Get CveUnidadTiempo() As Integer
    CveUnidadTiempo = intcveTiempo
End Property
Public Property Get DosisMax() As String
    DosisMax = Trim(strDosisMax)
End Property
Public Property Get bitCuadroBasico() As Integer
    bitCuadroBasico = intCuadroBasico
End Property
Public Property Get bitAntibiotico() As Integer
    bitAntibiotico = blnAntibiotico
End Property
Public Property Get bitLASA() As Integer
bitLASA = blnLASA
End Property
Public Property Get intcveIEPS() As Integer
    intcveIEPS = intcveIEPSvl
End Property
Public Property Get bitExcluirDescuento() As Integer
    bitExcluirDescuento = intExcluirDescuento
End Property
Public Property Get bitUnidosis() As Integer
bitUnidosis = blnUnidosis
End Property

'----------------------------------------------------------------------------------------------------
'||                                     M é t o d o s                                              ||
'----------------------------------------------------------------------------------------------------
Public Sub pCargaDatosArticulo(lngIdArt As Long, strClaveArticulo As String)
    Dim rsArticulo As New ADODB.Recordset
    
    Set rsArticulo = frsEjecuta_SP(IIf(lngIdArt = -1, "", lngIdArt) & "|" & strClaveArticulo & "|" & vgintClaveEmpresaContable, "sp_IvSelDatosArticuloEmpresa")
    With rsArticulo
        If rsArticulo.RecordCount > 0 Then
             lngIdArticulo = !intIdArticulo
             StrCveArticulo = !chrcvearticulo
             strNombreComercial = !vchNombreComercial
             dblCostoUltEntrada = !mnyCostoUltEntrada
             If Not IsNull(!dtmFechaUltEntrada) Then dtmFechaUltEntrada = !dtmFechaUltEntrada
             strEstatus = !VCHESTATUS
             strCostoGasto = !chrCostoGasto
             lngCveConceptFact = !smiCveConceptFact
             lngCveUniAlternaVta = !INTCVEUNIALTERNAVTA
             lngCveUniMinimaVta = !INTCVEUNIMINIMAVTA
             lngContenido = !intContenido
             lngCveConceptFact2 = !smiCveConceptFact2
             blnVentaPublico = !bitVentaPublico
             blnIndicaciones = !bitIndicaciones
             strCveArtMedicamen = !CHRCVEARTMEDICAMEN
             strCveFamilia = !chrCveFamilia
             strCveSubfamilia = !chrCveSubFamilia
             strNumArticulo = !smiNumArticulo
             blnManejaCaducidad = !bitManejaCaducidad
             lngIdGenerico = !intidgenerico
             strNombreGenerico = !NombreGenerico
             lngIdTipoControlado = !intIdTipo
             dblCostoMasAlto = !CostoMasAlto
             intExcluirDescuento = !bitExcluirDescuento
             If Not IsNull(!FechaCostoMasAlto) Then strFechaCostoMasAlto = !FechaCostoMasAlto
             dblPrecioMaximoPublico = !PrecioMaximoPublico
             If Not IsNull(!intcveVia) Then intcveVia = !intcveVia
             If Not IsNull(!smicveUnidad) Then intcveMedida = !smicveUnidad
             If Not IsNull(!intFrecuencia) Then intFrecuencia = !intFrecuencia
             If Not IsNull(!intcveUnidadTiempo) Then intcveTiempo = !intcveUnidadTiempo
             intCuadroBasico = !bitCuadroBasico
             If Not IsNull(!numDosisMaxima) Then strDosisMax = !numDosisMaxima
             blnAntibiotico = !bitAntibiotico
             blnLASA = !bitLASA
             
            
             If Not IsNull(!intcveIEPS) Then
                intcveIEPSvl = !intcveIEPS
             Else
                intcveIEPSvl = 0
             End If
             If Not IsNull(!dtmFechaDeAlta) Then dtmFechaAltaArticulos = !dtmFechaDeAlta
             If Not IsNull(!bitUnidosis) Then
             blnUnidosis = !bitUnidosis
            End If
        End If
    End With
    rsArticulo.Close
    
End Sub
'-----------------------------------------------------------------------
'||  Obtiene el primer código de barras registrado para un artículo
'-----------------------------------------------------------------------
Public Function fstrDameCodigoBarras() As String
    Dim rsCodigoBarras As New ADODB.Recordset
    
On Error GoTo NotificaError
    fstrDameCodigoBarras = ""
    Set rsCodigoBarras = frsEjecuta_SP(StrCveArticulo, "SP_IVSELCODIGOBARRAS")
    If rsCodigoBarras.RecordCount > 0 Then
        fstrDameCodigoBarras = rsCodigoBarras!vchCodigoBarras
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_fstrDameCodigoBarras"))
End Function
'--------------------------------------------------------------------------------
'||  Obtiene la longuitud de cualquier segmento del formato del artículo
'||  Tipo, Familia, Subfamilia ó Consecutivo
'---------------------------------------------------------------------------------
Private Function fintLonguitudSegmento(saSegmento As SegmentoArticulo) As Integer
    Dim strSentencia As String
    Dim rsFamiliaSubfamilia As New ADODB.Recordset
    Dim strFormatoArticulo As String
    Dim intIndice As Integer
    
On Error GoTo NotificaError
    fintLonguitudSegmento = 0
    Set rsFamiliaSubfamilia = frsSelParametros("SI", -1, "CHRFORMATOCVEARTICULO")
    If rsFamiliaSubfamilia.RecordCount > 0 Then
        If Not IsNull(rsFamiliaSubfamilia("Valor")) Then
            strFormatoArticulo = rsFamiliaSubfamilia("Valor")
            '|  Recorre el formato del artículo, se  interpreta por cada 2 dígitos,
            '|  el primer digito indica el tipo de dato y el segundo la longitud
            For intIndice = 1 To 8 Step 2
                '|  Compara el primer dígito del segmento con el segmento buscado
                If Mid(strFormatoArticulo, intIndice, 1) = saSegmento Then
                    fintLonguitudSegmento = CInt(Mid(strFormatoArticulo, intIndice + 1, 1))
                    Exit For
                End If
            Next
        End If
    End If
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_fintLonguitudSegmento"))
End Function

'-------------------------------------------------------------------------------------
'||  Obtiene la longuitud del segmento "Tipo de artículo" del formato del artículo
'-------------------------------------------------------------------------------------
Public Function fintLonguitudTipo() As Integer
    fintLonguitudTipo = fintLonguitudSegmento(saTipoArticulo)
End Function

'-------------------------------------------------------------------------------------
'||  Obtiene la longuitud del segmento "Familia" del formato del artículo
'-------------------------------------------------------------------------------------
Public Function fintLonguitudFamilia() As Integer
    fintLonguitudFamilia = fintLonguitudSegmento(saFamilia)
End Function

'-------------------------------------------------------------------------------------
'||  Obtiene la longuitud del segmento "Subfamilia" del formato del artículo
'-------------------------------------------------------------------------------------
Public Function fintLonguitudSubfamilia() As Integer
    fintLonguitudSubfamilia = fintLonguitudSegmento(saSubFamilia)
End Function

'-------------------------------------------------------------------------------------
'||  Obtiene la longuitud del segmento "Consecutivo" del formato del artículo
'-------------------------------------------------------------------------------------
Public Function fintLonguitudConsecutivo() As Integer
    fintLonguitudConsecutivo = fintLonguitudSegmento(saConsecutivo)
End Function

'-------------------------------------------------------------------------
'||  Da formato a un tipo de articulo, familia, subfamilia o consecutivo
'||  dependiendo de la configuración de la clave del artículo
'-------------------------------------------------------------------------
Public Function fstrFormateaSegmento(lngClave As Long, _
                                     saSegmento As SegmentoArticulo) As String
'|  lngClave.   Número al que se le dará formato
'|  saSegmento. Indica si se formateará como tipo, familia, subfamilia o consecutivo
    Dim strClave As String
    Dim intLongitud As Integer '(CR) Agregado (caso 6944)
    
On Error GoTo NotificaError
    
    fstrFormateaSegmento = ""
    strClave = CStr(lngClave)
    
    Select Case saSegmento
        Case saTipoArticulo
            intLongitud = fintLonguitudTipo
        Case saFamilia
            intLongitud = fintLonguitudFamilia
        Case saSubFamilia
            intLongitud = fintLonguitudSubfamilia
        Case saConsecutivo
            intLongitud = fintLonguitudConsecutivo
    End Select
    
    ' (CR) - Caso 6944: Si la longitud dada por la función correspondiente era menor a la longitud de la clave marcaba error '
    If intLongitud >= Len(strClave) Then
        fstrFormateaSegmento = String(intLongitud - Len(strClave), "0") & strClave
    End If
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_fstrFormateaSegmento"))
End Function

'-------------------------------------------------------------------------
'||  Insertar un artículo en el catálogo
'-------------------------------------------------------------------------
Public Function flngIdInsertaArticulo( _
strClaveArticulo As String, strNombreComercial As String _
, dblCostoUltimaEntrada As Double _
, strFechaUltimaEntrada As String _
, strEstado As String, strCostoGasto As String _
, lngCveConceptoFact As Long _
, intCveUniAlterna As Integer _
, intCveUniMinima As Integer _
, intContenido As Long _
, lngCveConceptoFact2 As Long _
, bitVentaPublico As Integer _
, intTipoArticulo As Integer _
, intCveFamilia As Integer _
, intCveSubfamilia As Integer _
, lngNumero As Long _
, bitCaducidad As Integer _
, lngCveGenerico As Long _
, intCveTipoControlado As Integer _
, lngIdArticulo As Long _
, dblCostoMasAlto As Double _
, strFechaCostoMasAlto As String _
, intcveVia As Integer _
, intcveMedida As Integer, intExcluirDescuento As Integer _
, intFrecuencia As Integer, intcveTiempo As Integer, intCuadroBasico As Integer, strDosisMax As String _
, blnAntibiotico As Integer, blnLASA As Integer, intcveIEPSvl As Integer, blnEnIndicaciones, dtmFechaAltaArticulos, blnUnidosis As Integer) As Long

On Error GoTo NotificaError

    Dim lngId As Long

    vgstrParametrosSP = _
    strClaveArticulo _
    & "|" & strNombreComercial _
    & "|" & CStr(dblCostoUltimaEntrada) _
    & "|" & strFechaUltimaEntrada _
    & "|" & strEstado _
    & "|" & strCostoGasto _
    & "|" & CStr(lngCveConceptoFact) _
    & "|" & CStr(intCveUniAlterna) _
    & "|" & CStr(intCveUniMinima) _
    & "|" & CStr(intContenido) _
    & "|" & CStr(lngCveConceptoFact2)
    
    vgstrParametrosSP = vgstrParametrosSP _
    & "|" & CStr(bitVentaPublico) _
    & "|" & CStr(intTipoArticulo) _
    & "|" & CStr(intCveFamilia) _
    & "|" & CStr(intCveSubfamilia) _
    & "|" & CStr(lngNumero) _
    & "|" & CStr(bitCaducidad) _
    & "|" & CStr(lngCveGenerico) _
    & "|" & CStr(intCveTipoControlado) _
    & "|" & CStr(lngIdArticulo) _
    & "|" & CStr(dblCostoMasAlto) _
    & "|" & CStr(strFechaCostoMasAlto) _
    & "|" & CStr(intcveVia) _
    & "|" & CStr(intcveMedida) _
    & "|" & CStr(intFrecuencia) _
    & "|" & CStr(intcveTiempo) _
    & "|" & CStr(intCuadroBasico) _
    & "|" & CStr(strDosisMax) _
    & "|" & CStr(blnAntibiotico) _
    & "|" & CStr(blnLASA) _
    & "|" & IIf(intcveIEPSvl = 0, Null, CStr(intcveIEPSvl)) _
    & "|" & blnEnIndicaciones _
    & "|" & CStr(intExcluirDescuento) _
    & "|" & CStr(dtmFechaAltaArticulos) _
    & "|" & CStr(blnUnidosis)
    
    'blnUnidosis As Integer
    
    flngIdInsertaArticulo = 1
    frsEjecuta_SP vgstrParametrosSP, "sp_IvInsArticulo", True, flngIdInsertaArticulo
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_flngInsertaArticulo"))
End Function

'-------------------------------------------------------------------------
'||  Insertar código de barras del artículo
'-------------------------------------------------------------------------
Public Sub pInsertaCodigoBarras(StrCveArticulo As String, strCodigoBarras As String)
    On Error GoTo NotificaError

    vgstrParametrosSP = Trim(StrCveArticulo) & "|" & Trim(strCodigoBarras)
    frsEjecuta_SP vgstrParametrosSP, "SP_IVINSCODIGOBARRAS"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pInsertaCodigoBarras"))
End Sub

'-----------------------------------------------------------------------
'||  Obtiene los códigos de barras de un artículo
'-----------------------------------------------------------------------
Public Function frsCodigoBarras(StrCveArticulo As String) As Recordset
    On Error GoTo NotificaError

    Set frsCodigoBarras = frsEjecuta_SP(StrCveArticulo, "SP_IVSELCODIGOBARRAS")

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_frsCodigoBarras"))
End Function

'-----------------------------------------------------------------------
'||  Elimina un artículo de las listas de precios
'-----------------------------------------------------------------------
Public Sub pEliminaArticuloLista(StrCveArticulo As String)
    On Error GoTo NotificaError

    vgstrParametrosSP = Trim(StrCveArticulo) & "|" & "AR"
    
    frsEjecuta_SP vgstrParametrosSP, "sp_PvDelCargoListaPrecio", True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pEliminaArticuloLista"))
End Sub

'-----------------------------------------------------------------------
'||  Eliminar los códigos de barra de un artículo
'-----------------------------------------------------------------------
Public Sub pEliminaCodigoBarras(StrCveArticulo As String)
    On Error GoTo NotificaError
    
    frsEjecuta_SP StrCveArticulo, "SP_IVDELCODIGOBARRAS", True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pEliminaCodigoBarras"))
End Sub

'-----------------------------------------------------------------------
'||  Actualizar los datos de un artículo
'-----------------------------------------------------------------------
Public Sub pActualizaArticulo(lngIdArticulo As Long, _
strNombreComercial As String _
, lngIdGenerico As Long _
, strClasificacion As String _
, lngCveConceptoBase As Long _
, lngCveConceptoAplicado As Long _
, lngCveUnidadMinima As Long _
, lngCveUnidadAlterna As Long _
, lngContenido As Long _
, intCaduca As Integer _
, intIdTipoControlado As Integer _
, intVentaPublico As Integer _
, strEstado As String _
, dblCostroEntrada As Double _
, strFechaUltimaEntrada As String _
, dblCostoMasAlto As Double _
, strFechaCostoMasAlto As String _
, intcveVia As Integer _
, intcveMedida As Integer _
, intFrecuencia As Integer _
, intcveTiempo As Integer _
, intCuadroBasico _
, strDosisMax As String, blnAntibiotico As Integer, blnLASA As Integer, intcveIEPSvl As Integer, bitIndicaciones, intExcluirDescuento As Integer, blnUnidosis As Integer)

    On Error GoTo NotificaError
    
    vgstrParametrosSP = _
    CStr(lngIdArticulo) _
    & "|" & Trim(strNombreComercial) _
    & "|" & CStr(lngIdGenerico) _
    & "|" & Trim(strClasificacion) _
    & "|" & CStr(lngCveConceptoBase) _
    & "|" & CStr(lngCveConceptoAplicado) _
    & "|" & CStr(lngCveUnidadMinima) _
    & "|" & CStr(lngCveUnidadAlterna) _
    & "|" & CStr(lngContenido) _
    & "|" & CStr(intCaduca) _
    & "|" & CStr(intIdTipoControlado) _
    & "|" & CStr(intVentaPublico) _
    & "|" & Trim(strEstado) _
    & "|" & Trim(dblCostroEntrada) _
    & "|" & strFechaUltimaEntrada _
    & "|" & Trim(dblCostoMasAlto) _
    & "|" & strFechaCostoMasAlto _
    & "|" & CStr(intcveVia) _
    & "|" & CStr(intcveMedida) _
    & "|" & CStr(intFrecuencia) _
    & "|" & CStr(intcveTiempo) _
    & "|" & CStr(intCuadroBasico) & "|" & CStr(strDosisMax) & "|" & CStr(blnAntibiotico) & "|" & CStr(blnLASA) & "|" & IIf(intcveIEPSvl = 0, Null, CStr(intcveIEPSvl)) & "|" & bitIndicaciones & "|" & CStr(intExcluirDescuento) & "|" & CStr(blnUnidosis)
    
    frsEjecuta_SP vgstrParametrosSP, "sp_IvUpdArticulo", True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pActualizaArticulo"))
End Sub

'-----------------------------------------------------------------------
'||  Elimina un artículo
'-----------------------------------------------------------------------
Public Function flngEliminaArticulo(lngIdArticulo As Long) As Long
    'Si el procedimiento fue existoso regresa un cero, si no, el num. de error de integridad de ORACLE
    On Error GoTo NotificaError
    
    flngEliminaArticulo = 1
    frsEjecuta_SP CStr(lngIdArticulo), "sp_IvDelArticulo", False, flngEliminaArticulo
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_flngEliminaArticulo"))
End Function

'-----------------------------------------------------------------------
'||  Inserta los manejos
'-----------------------------------------------------------------------
Public Sub pInsertaManejos(intIdArticulo As Long, intCveManejo As Integer)
    On Error GoTo NotificaError
    
    frsEjecuta_SP CStr(intIdArticulo) & "|" & CStr(intCveManejo), "SP_IVINSARTICULOMANEJO"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pEliminaCodigoBarras"))
End Sub

'-----------------------------------------------------------------------
'||  Elimina los manejos
'-----------------------------------------------------------------------
Public Sub pEliminaManejos(intIdArticulo As Long)
    On Error GoTo NotificaError
    
    frsEjecuta_SP CStr(intIdArticulo), "SP_IVDELARTICULOMANEJO"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":clsArticulo_pEliminaCodigoBarras"))
End Sub

