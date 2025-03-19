VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarNombre As String 'local copy
Private mvarApPaterno As String 'local copy
Private mvarApMaterno As String 'local copy
Private mvarFechaNacimiento As Date 'local copy
Private mvarSexo As String 'local copy
Private mvarIdPersona As Long 'local copy
Private mvarStatusAdPaciente As String 'local copy
Private mvarStatusPadron As String 'local copy
Private mvarNumAfiliacion As String 'local copy
Private mvarDireccion As String 'local copy
Private mvarDependencias As String 'local copy
Private mvarPais As String 'local copy
Private mvarCiudad As String 'local copy
Private mvarEstado As String 'local copy
Private mvarColonia As String 'local copy
Private mvarTelefono As String 'local copy
Private mvarVigencia As String 'local copy

Public Property Let Vigencia(ByVal vData As String)
    mvarVigencia = vData
End Property

Public Property Get Vigencia() As String
    Vigencia = mvarVigencia
End Property

Public Property Let Telefono(ByVal vData As String)
    mvarTelefono = vData
End Property


Public Property Get Telefono() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Telefono
    Telefono = mvarTelefono
End Property



Public Property Let Colonia(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Colonia = 5
    mvarColonia = vData
End Property


Public Property Get Colonia() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Colonia
    Colonia = mvarColonia
End Property



Public Property Let Estado(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Estado = 5
    mvarEstado = vData
End Property


Public Property Get Estado() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Estado
    Estado = mvarEstado
End Property



Public Property Let Ciudad(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Ciudad = 5
    mvarCiudad = vData
End Property


Public Property Get Ciudad() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Ciudad
    Ciudad = mvarCiudad
End Property



Public Property Let Pais(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Pais = 5
    mvarPais = vData
End Property


Public Property Get Pais() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Pais
    Pais = mvarPais
End Property



Public Property Let Dependencias(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Dependencias = 5
    mvarDependencias = vData
End Property


Public Property Get Dependencias() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Dependencias
    Dependencias = mvarDependencias
End Property



Public Property Let Direccion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Direccion = 5
    mvarDireccion = vData
End Property


Public Property Get Direccion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Direccion
    Direccion = mvarDireccion
End Property



Public Property Let NumAfiliacion(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NumAfiliacion = 5
    mvarNumAfiliacion = vData
End Property


Public Property Get NumAfiliacion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumAfiliacion
    NumAfiliacion = mvarNumAfiliacion
End Property



Public Property Let StatusPadron(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusPadron = 5
    mvarStatusPadron = vData
End Property


Public Property Get StatusPadron() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusPadron
    StatusPadron = mvarStatusPadron
End Property



Public Property Let StatusAdPaciente(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.StatusAdPaciente = 5
    mvarStatusAdPaciente = vData
End Property


Public Property Get StatusAdPaciente() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StatusAdPaciente
    StatusAdPaciente = mvarStatusAdPaciente
End Property



Public Property Let IdPersona(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.lngIdPersona = 5
    mvarIdPersona = vData
End Property


Public Property Get IdPersona() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.lngIdPersona
    IdPersona = mvarIdPersona
End Property



Public Property Let Sexo(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Sexo = 5
    mvarSexo = vData
End Property


Public Property Get Sexo() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Sexo
    Sexo = mvarSexo
End Property



Public Property Let FechaNacimiento(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FechaNacimiento = 5
    mvarFechaNacimiento = vData
End Property


Public Property Get FechaNacimiento() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FechaNacimiento
    FechaNacimiento = mvarFechaNacimiento
End Property



Public Property Let ApMaterno(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ApMaterno = 5
    mvarApMaterno = vData
End Property


Public Property Get ApMaterno() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ApMaterno
    ApMaterno = mvarApMaterno
End Property



Public Property Let ApPaterno(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ApPaterno = 5
    mvarApPaterno = vData
End Property


Public Property Get ApPaterno() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ApPaterno
    ApPaterno = mvarApPaterno
End Property



Public Property Let Nombre(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

Public Function GetById(ByVal lngId As Long) As Boolean
    Dim objPersona As Object
    Set objPersona = CreateObject("AfiliacionExterna.Persona")
    Dim rs As ADODB.recordSet
    Set rs = objPersona.GetById(lngId)
    If rs.State <> 0 Then
        If Not rs.EOF Then
            mvarNombre = IIf(IsNull(rs!Nombre), "", rs!Nombre)
            mvarApPaterno = IIf(IsNull(rs!Paterno), "", rs!Paterno)
            mvarApMaterno = IIf(IsNull(rs!Materno), "", rs!Materno)
            mvarFechaNacimiento = IIf(IsNull(rs!Fecha_nacimiento), "", rs!Fecha_nacimiento)
            mvarSexo = IIf(IsNull(rs!Sexo), "", rs!Sexo)
            mvarDireccion = IIf(IsNull(rs!calle), "", rs!calle)
            'mvarPais = IIf(IsNull(rs!vchPais), "", rs!vchPais)
            'mvarEstado = IIf(IsNull(rs!vchEstado), "", rs!vchEstado)
            'mvarCiudad = IIf(IsNull(rs!vchCiudad), "", rs!vchCiudad)
            'mvarColonia = IIf(IsNull(rs!vchColonia), "", rs!vchColonia)
            'mvarTelefono = IIf(IsNull(rs!vchTelefono), "", rs!vchTelefono)
            mvarNumAfiliacion = IIf(IsNull(rs!numafi), "", rs!numafi) & IIf(IsNull(rs!numben), "", "/" & rs!numben)
            mvarIdPersona = rs!persona_id
            mvarDependencias = IIf(IsNull(rs!Dependencias), "", rs!Dependencias)
            mvarVigencia = IIf(IsNull(rs!vigente), "", rs!vigente)
            GetById = True
        Else
            mvarIdPersona = lngId
            GetById = False
        End If
    Else
        MsgBox "No se pudo consultar el padrón central" & vbCrLf & "Se cargará la información local", vbInformation
        mvarIdPersona = lngId
        GetById = False
    End If
End Function

Public Function GetByInfo(ByVal Nombre As String, ByVal Paterno As String, ByVal Materno As String, ByVal FechaNacimiento As Date) As Boolean
    Dim objPersona As Object
    Set objPersona = CreateObject("AfiliacionExterna.Persona")
    Dim frmPadron As frmPersonasPadron
    Dim rs As ADODB.recordSet
    Nombre = IIf(Nombre = "", "%", Nombre)
    Paterno = IIf(Paterno = "", "%", Paterno)
    Materno = IIf(Materno = "", "%", Materno)
    Set rs = objPersona.GetByInfo(Nombre, Paterno, Materno, FechaNacimiento)
    If rs.State <> 0 Then
        If Not rs.EOF Then
            Set frmPadron = New frmPersonasPadron
            frmPadron.grdPacientes.Clear
            pLlenarMshFGrdRs frmPadron.grdPacientes, rs
            frmPadron.pConfigura
            frmPadron.Show vbModal
            If frmPadron.lngIdPersona <> 0 Then
                mvarNombre = frmPadron.strNombre
                mvarApPaterno = frmPadron.strApPaterno
                mvarApMaterno = frmPadron.strApMaterno
                mvarFechaNacimiento = frmPadron.dtmFechaNacimiento
                mvarSexo = frmPadron.strSexo
                mvarNumAfiliacion = frmPadron.strNumAfiliacion
                mvarIdPersona = frmPadron.lngIdPersona
                mvarDireccion = frmPadron.strDireccion
                mvarDependencias = frmPadron.strDependencias
                'mvarPais = frmPadron.strPais
                'mvarEstado = frmPadron.strEstado
                'mvarCiudad = frmPadron.strCiudad
                'mvarColonia = frmPadron.strColonia
                'mvarTelefono = frmPadron.strTelefono
                mvarVigencia = frmPadron.strVigencia
                Unload frmPadron
                GetByInfo = True
            Else
                mvarIdPersona = 0
                GetByInfo = False
            End If
        Else
            mvarIdPersona = 0
            GetByInfo = False
        End If
    Else
        MsgBox "No se pudo conectar al padrón central" & vbCrLf & "El identificador quedará pendiente", vbInformation
        mvarIdPersona = -1
        GetByInfo = False
    End If
End Function

Private Sub Class_Initialize()
    mvarStatusAdPaciente = ""
    mvarStatusPadron = ""
    mvarIdPersona = 0
End Sub

Public Sub UpdatePaciente(ByVal Nombre As String, ByVal ApPaterno As String, ByVal ApMaterno As String, ByVal FechaNacimiento As Date, ByVal Sexo As String, ByVal Direccion As String, ByVal Pais As String, ByVal Estado As String, ByVal Ciudad As String, ByVal Colonia As String, ByVal Telefono As String)
    Dim objPersona As Object
    Set objPersona = CreateObject("AfiliacionExterna.Persona")
    If mvarStatusPadron = "Insertar" Then
        mvarNombre = Nombre
        mvarApPaterno = ApPaterno
        mvarApMaterno = ApMaterno
        mvarFechaNacimiento = FechaNacimiento
        mvarSexo = Sexo
        mvarPais = Pais
        mvarEstado = Estado
        mvarCiudad = Ciudad
        mvarColonia = Colonia
        mvarDireccion = Direccion
        mvarTelefono = Telefono
        If objPersona.AddPerson(mvarNombre, mvarApPaterno, mvarApMaterno, mvarFechaNacimiento, mvarSexo, mvarPais, mvarEstado, mvarCiudad, mvarColonia, mvarDireccion, mvarTelefono, mvarIdPersona, mvarNumAfiliacion) Then
        Else
            MsgBox "No se pudo registrar en el padrón central" & vbCrLf & "El identificador quedará pendiente", vbInformation
            mvarIdPersona = -1
        End If
    End If
End Sub
