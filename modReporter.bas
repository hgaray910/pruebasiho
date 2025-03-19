Attribute VB_Name = "modReporter"
Option Explicit

Public vgStrVariablesRep(50) As String
Public Function GetValue(ByVal strData As String) As Variant
    Select Case strData
        ' General
        Case "vgstrNombreHospitalCH"
            GetValue = vgstrNombreHospitalCH
        Case "vgstrNombCortoCH"
            GetValue = vgstrNombCortoCH
        Case "vgstrRfCCH"
            GetValue = vgstrRfCCH
        Case "vgstrIMSSCH"
            GetValue = vgstrIMSSCH
        Case "vgstrSSACH"
            GetValue = vgstrSSACH
        Case "vgstrRepLegalCH"
            GetValue = vgstrRepLegalCH
        Case "vgstrDirGnralCH"
            GetValue = vgstrDirGnralCH
        Case "vgstrDirMedCH"
            GetValue = vgstrDirMedCH
        Case "vgstrAdmGnralCH"
            GetValue = vgstrAdmGnralCH
        Case "vgstrDireccionCH"
            GetValue = vgstrDireccionCH
        Case "vgstrColoniaCH"
            GetValue = vgstrColoniaCH
        Case "vgintCveCiudadCH"
            GetValue = vgintCveCiudadCH
        Case "vgstrCiudadCH"
            GetValue = vgstrCiudadCH
        Case "vgintCveEstadoCH"
            GetValue = vgintCveEstadoCH
        Case "vgstrEstadoCH"
            GetValue = vgstrEstadoCH
        Case "vgintCvePaisCH"
            GetValue = vgintCvePaisCH
        Case "vgstrPaisCH"
            GetValue = vgstrPaisCH
        Case "vgstrTelefonoCH"
            GetValue = vgstrTelefonoCH
        Case "vgstrFaxCH"
            GetValue = vgstrFaxCH
        Case "vgstrEmailCH"
            GetValue = vgstrEmailCH
        Case "vgstrWebCH"
            GetValue = vgstrWebCH
        Case "vgstrCodPostalCH"
            GetValue = vgstrCodPostalCH
        Case "vgstrApartPostalCH"
            GetValue = vgstrApartPostalCH
        Case "vgintTipoCuartoCH"
            GetValue = vgintTipoCuartoCH
        Case "vgintEstadoSaludCH"
            GetValue = vgintEstadoSaludCH
        Case "vgbytDisponibleCH"
            GetValue = vgbytDisponibleCH
        Case "vgbytOcupadoCH"
            GetValue = vgbytOcupadoCH
        ' Seguridad
        Case "vgintNumeroModulo"
            GetValue = vgintNumeroModulo
        Case "vglngNumeroEmpleado"
            GetValue = vglngNumeroEmpleado
        Case "vgintNumeroDepartamento"
            GetValue = vgintNumeroDepartamento
        Case "vglngNumeroLogin"
            GetValue = vglngNumeroLogin
        Case "vgstrNombreUsuario"
            GetValue = vgstrNombreUsuario
        Case "vgstrNombreDepartamento"
            GetValue = vgstrNombreDepartamento
        Case "fdtmServerFecha"
            GetValue = fdtmServerFecha()
        Case Else
            GetValue = ""
    End Select
End Function

Public Sub pObtenerVariables(ByRef cboVariables As ComboBox)
    cboVariables.AddItem "Elija ->"
    cboVariables.AddItem "(Establecer valor)"
    cboVariables.AddItem "Variable que contiene el nombre del hospital"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrNombreHospitalCH"
    cboVariables.AddItem "Nombre corto del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrNombCortoCH"
    cboVariables.AddItem "Rfc del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrRfCCH"
    cboVariables.AddItem "Registro IMSS del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrIMSSCH"
    cboVariables.AddItem "Licencia SSA del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrSSACH"
    cboVariables.AddItem "Nombre del respresentante legal del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrRepLegalCH"
    cboVariables.AddItem "Nombre del director general del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrDirGnralCH"
    cboVariables.AddItem "Nombre del director médico del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrDirMedCH"
    cboVariables.AddItem "Nombre del administrador general del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrAdmGnralCH"
    cboVariables.AddItem "Dirección del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrDireccionCH"
    cboVariables.AddItem "Colonia del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrColoniaCH"
    cboVariables.AddItem "Clave de la ciudad donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintCveCiudadCH"
    cboVariables.AddItem "Ciudad donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrCiudadCH"
    cboVariables.AddItem "Clave del estado donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintCveEstadoCH"
    cboVariables.AddItem "Estado donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrEstadoCH"
    cboVariables.AddItem "Clave del pais donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintCvePaisCH"
    cboVariables.AddItem "País donde se encuentra el centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrPaisCH"
    cboVariables.AddItem "Teléfono del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrTelefonoCH"
    cboVariables.AddItem "Fax del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrFaxCH"
    cboVariables.AddItem "Email del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrEmailCH"
    cboVariables.AddItem "Página Web del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrWebCH"
    cboVariables.AddItem "Código postal del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrCodPostalCH"
    cboVariables.AddItem "Apartado postal del del centro hospitalario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrApartPostalCH"
    cboVariables.AddItem "Clave del Tipo de cuarto predeterminado para mostrarlo en la admisin"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintTipoCuartoCH"
    cboVariables.AddItem "Clave del Estado de Salud que por omisión debe ser ingresado por admisión"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintEstadoSaludCH"
    cboVariables.AddItem "Código del Estado de cuarto disponible"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgbytDisponibleCH"
    cboVariables.AddItem "Código del Estado de cuarto ocupado"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgbytOcupadoCH"
    cboVariables.AddItem "Número de módulo de la tabla Modulo"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintNumeroModulo"
    cboVariables.AddItem "Número de empleado logueado"
    vgStrVariablesRep(cboVariables.NewIndex) = "vglngNumeroEmpleado"
    cboVariables.AddItem "Número de departamento con que se logueo el empleado"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgintNumeroDepartamento"
    cboVariables.AddItem "Número de login que le corresponde en la tabla Login"
    vgStrVariablesRep(cboVariables.NewIndex) = "vglngNumeroLogin"
    cboVariables.AddItem "Login del usuario"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrNombreUsuario"
    cboVariables.AddItem "Nombre del departamento personalizado con que se loguea el empleado"
    vgStrVariablesRep(cboVariables.NewIndex) = "vgstrNombreDepartamento"
    cboVariables.AddItem "Fecha del servidor"
    vgStrVariablesRep(cboVariables.NewIndex) = "fdtmServerFecha"
End Sub



