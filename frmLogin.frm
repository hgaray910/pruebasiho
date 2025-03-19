VERSION 5.00
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de acceso"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContrasena 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "•"
      TabIndex        =   1
      ToolTipText     =   "Contraseña del usuario"
      Top             =   520
      Width           =   3255
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Nombre del usuario"
      Top             =   120
      Width           =   3255
   End
   Begin HSFlatControls.MyCombo cboEmpresa 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   930
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Style           =   1
      Enabled         =   0   'False
      Text            =   ""
      Sorted          =   -1  'True
      List            =   ""
      ItemData        =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblempresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   610
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Programa para acceso al módulo (Nueva Versión para la seguridad)
' Fecha de programacion: 1 de Diciembre del 2000
'-----------------------------------------------------------------------------
' Ultimas modificaciones:
'-----------------------------------------------------------------------------
' Fecha:
' Descripción del cambio:
'-----------------------------------------------------------------------------
Dim vlintNumeroIntentos As Integer
Public vgblnCargaVariablesGlobales As Boolean
Public vgintLogin As Integer
Public vgstrempresacontable As String
Public strTiempoRestanteTotal As String

Private Sub CboEmpresa_KeyDown(KeyAscii As Integer, Shift As Integer)
    If cboEmpresa.ListIndex > -1 Then
        If KeyAscii = 13 Then
            vlintNumeroIntentos = vlintNumeroIntentos + 1
            If fblnContrasenaValida(txtUsuario.Text, txtContrasena.Text, vgblnCargaVariablesGlobales) Then
                If vgblnCargaVariablesGlobales Then
                    pCargarVarPrmGnrl
                    If Forms.Count = 1 Then frmfondo.Show 'Para que se carge el fondo, solo cuando sea llamada en el inicio del módulo
                End If
                Unload Me
            Else
                Call pEnfocaTextBox(txtUsuario)
                If vlintNumeroIntentos = 3 Then
                    Unload Me
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    vgintLogin = -1
    txtUsuario.Text = ""
    txtContrasena.Text = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
      Unload Me
    Else
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))

End Sub

Private Sub Form_Load()

On Error GoTo NotificaError
    
    pCargaCombo
    vlintNumeroIntentos = 0
    vgblnCargaVariablesGlobales = True
    vgstrempresacontable = ""
    
    'Se valida la vigencia
    Call pValidaVigencia
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))

End Sub

Private Sub pCargaCombo()
    Dim rs As ADODB.Recordset
    Set rs = frsRegresaRs("select tnyClaveEmpresa, vchNombreCorto from CNEmpresaContable where bitActiva <> 0 order by vchNombreCorto")
    If Not rs.EOF Then
        pLlenarCboRs_new cboEmpresa, rs, 0, 1
    End If
    rs.Close
End Sub

Private Sub pValidaVigencia()

Dim rsCnEmpresaContable As New ADODB.Recordset
Dim VigenciaDecoded As String
Dim VigenciaEncoded As String
Dim PalabraSecretaPart As String
Dim VigenciaPart As String
Dim RFCPart As String
Dim strFechaHoraServer As String
Dim strFechaServer As String
Dim strHoraServer As String
Dim strFechaHoraServerLetra As String
Dim lngTiempoRestanteHrs As Long
Dim lngTiempoRestanteMins As Long

On Error GoTo NotificaError

'Se obtienen fecha y hora del servidor
strFechaHoraServer = Format(fdtmServerFechaHora, "DD/MM/YYYY HH:MM:SS")
strFechaServer = Mid(strFechaHoraServer, 1, 10)
strHoraServer = Mid(strFechaHoraServer, 12)
strFechaHoraServerLetra = Mid(strFechaHoraServer, 1, 3) & fstrMesLetra(Mid(strFechaHoraServer, 4, 2), False) & Mid(strFechaHoraServer, 6)


'+++++++++++++++++++++++++ OBTENCIÓN DE LA VIGENCIA +++++++++++++++++++++++++
Set rsCnEmpresaContable = frsRegresaRs("select * from CNEmpresaContable where tnyClaveEmpresa = 1")

If rsCnEmpresaContable.RecordCount > 0 Then

    VigenciaEncoded = IIf(IsNull(Trim(rsCnEmpresaContable!vchVigencia)), "", Trim(rsCnEmpresaContable!vchVigencia))
    
    If Trim(VigenciaEncoded) = "" Then
        ' No se ha encontrado licencia del SiHO
        MsgBox SIHOMsg(1082), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
    'Se eliminan los caracteres basura de la vigencia
    VigenciaEncoded = Mid(VigenciaEncoded, 13, Len(VigenciaEncoded) - 22)

    'Se reemplazan los caracteres especiales ("U"<-"?"   "l"<-"ñ"   "="<-"Ñ"   "=="<-"Ñ?")
    VigenciaEncoded = Replace(VigenciaEncoded, "Ñ?", "==")
    VigenciaEncoded = Replace(VigenciaEncoded, "Ñ", "=")
    VigenciaEncoded = Replace(VigenciaEncoded, "ñ", "l")
    VigenciaEncoded = Replace(VigenciaEncoded, "?", "U")
        
    'Se decodifica la vigencia a partir de VigenciaEncoded (1a vez)
    VigenciaDecoded = Decode(VigenciaEncoded)
    
    'Se decodifica la vigencia a partir de VigenciaEncoded (2a vez)
    VigenciaDecoded = Decode(VigenciaDecoded)
    
    'Validación simple
    If Trim(VigenciaDecoded) = "P@sSWorD" Then
        ' La licencia del SiHO ha sido alterada
        MsgBox SIHOMsg(1081), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
    'Se obtiene la palabra secreta a partir de VigenciaDecoded
    PalabraSecretaPart = Left(VigenciaDecoded, 8)
    
    'Se obtiene la fecha de vigencia a partir de VigenciaDecoded
    VigenciaPart = Right(VigenciaDecoded, 10)
    
    'Se obtiene el RFC a partir de VigenciaDecoded
    RFCPart = Mid(VigenciaDecoded, 9, Len(VigenciaDecoded) - 18)
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    

    '------------------------------------- VALIDACIÓN Y COMPARACIÓN DE VALORES -------------------------------------
    ' Se comentó esta linea para validar el RFC sin espacios o guiones
    If Trim(PalabraSecretaPart) <> "P@sSWorD" Or Trim(Replace(Replace(Replace(RFCPart, "-", ""), "_", ""), " ", "")) <> Trim(Replace(Replace(Replace(rsCnEmpresaContable!vchRFC, "-", ""), "_", ""), " ", "")) Or IsDate(Trim(VigenciaPart)) = False Then
        ' La licencia del SiHO ha sido alterada
        MsgBox SIHOMsg(1081), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
    End If
    
    'Si la vigencia no ha sido alterada, se verifica si el cliente es de arrendamiento (Fecha de vigencia <> "01/01/2099")
    If Trim(VigenciaPart) <> "01/01/2099" Then
    
        'Comparación de la fecha de vigencia con la fecha del server
        If CDate(Trim(VigenciaPart)) < CDate(strFechaServer) Then
            ' La licencia del SiHO ha expirado
            MsgBox SIHOMsg(1080), vbOKOnly + vbCritical, "Mensaje"
            Unload Me
            Exit Sub
            
        ElseIf CDate(Trim(VigenciaPart)) > CDate(strFechaServer) Then 'Se obtienen los días restantes de vigencia
            strTiempoRestanteTotal = CStr(DateDiff("d", Format(CDate(strFechaServer), "DD/MM/YYYY"), Format(CDate(VigenciaPart), "DD/MM/YYYY")))
                        
            'Si es mayor a 5 días, se limpian los días restantes para no mostrar la etiqueta de advertencia en el menú principal
            If Val(strTiempoRestanteTotal) > 5 Then
                strTiempoRestanteTotal = ""
            End If
            
        ElseIf CDate(Trim(VigenciaPart)) = CDate(strFechaServer) Then 'Si es la misma fecha, se compara la hora (por default la hora de vigencia es a las 23:30pm)
            If Format(CDate("23:30:00"), "HH:MM:SS") > Format(CDate(strHoraServer), "HH:MM:SS") Then
                
                'Se calculan los minutos restantes de vigencia
                lngTiempoRestanteHrs = DateDiff("h", Format(CDate(strHoraServer), "HH:MM:SS"), Format(CDate("23:30:00"), "HH:MM:SS"))
                lngTiempoRestanteMins = DateDiff("n", Format(CDate(strHoraServer), "HH:MM:SS"), Format(CDate("23:30:00"), "HH:MM:SS"))
                strTiempoRestanteTotal = CStr(lngTiempoRestanteHrs) & " horas " & CStr(Abs(lngTiempoRestanteMins - (60 * lngTiempoRestanteHrs))) & " minutos"
                                
                'Se muestra mensaje únicamente cuando quedan unas horas para que expire la vigencia (como medida de emergencia)
                MsgBox "¡La licencia del SiHO expirará en " & strTiempoRestanteTotal & "!", vbExclamation, "Mensaje"
                
            Else
                ' La licencia del SiHO ha expirado
                MsgBox SIHOMsg(1080), vbOKOnly + vbCritical, "Mensaje"
                Unload Me
                Exit Sub
            End If
        End If
            
    'Si la vigencia está en blanco
    ElseIf Trim(PalabraSecretaPart) = "" Then
        ' No se ha encontrado licencia del SiHO
        MsgBox SIHOMsg(1082), vbOKOnly + vbCritical, "Mensaje"
        Unload Me
        Exit Sub
        
    End If
    '-------------------------------------------------------------------------------------------------------------------------------

'No se encuentra información en CNEMPRESACONTABLE
Else
    'No se ha configurado la información de la empresa contable
    MsgBox "No se ha configurado la información de la empresa contable", vbOKOnly + vbCritical, "Mensaje"
    Unload Me
    Exit Sub
    
End If

Exit Sub

NotificaError:
'    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaVigencia"))
    ' La licencia del SiHO ha sido alterada
    MsgBox SIHOMsg(1081), vbOKOnly + vbCritical, "Mensaje"
    Unload Me
    Exit Sub
End Sub
Private Sub txtContrasena_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtContrasena

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtContrasena_GotFocus"))

End Sub

Private Sub txtContrasena_KeyPress(KeyAscii As Integer)

Dim vllngNumeroEmpleado As Long
Dim vlintNumeroDepartamento As Integer
Dim vlstrNombreDepartamento As String
Dim vllngNumeroLogin As Long
Dim vlstrNombreUsuario As String

On Error GoTo NotificaError
    If KeyAscii = 13 Then
        If cboEmpresa.Enabled Then
            cboEmpresa.SetFocus
        Else
            vlintNumeroIntentos = vlintNumeroIntentos + 1
            
            
            
            
            If fblnContrasenaValida(txtUsuario.Text, txtContrasena.Text, vgblnCargaVariablesGlobales) Then
                vgstrempresacontable = txtUsuario.Text
                If vgblnCargaVariablesGlobales Then
                    pCargarVarPrmGnrl
                    If Forms.Count = 1 Then frmfondo.Show 'Para que se carge el fondo, solo cuando sea llamada en el inicio del módulo
                End If
                Unload Me
            Else
                Call pEnfocaTextBox(txtUsuario)
                If vlintNumeroIntentos = 3 Then
                    Unload Me
                End If
            End If
            
            
            
            
            
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtContrasena_KeyPress"))
End Sub

Private Sub txtUsuario_Change()
    If UCase(txtUsuario.Text) = "ADMINISTRADOR" Then
        cboEmpresa.Enabled = True
        cboEmpresa.ListIndex = flngLocalizaCbo_new(cboEmpresa, "1")
    Else
        cboEmpresa.Enabled = False
        cboEmpresa.ListIndex = -1
    End If
End Sub

Private Sub txtUsuario_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtUsuario

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtUsuario_GotFocus"))
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 13 Then
        txtContrasena.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtUsuario_KeyPress"))

End Sub

Public Function fblnContrasenaValida(vlstrNombreUsuario As String, vlstrContrasena As String, Optional pblnCargaVariablesGlobales As Boolean = True) As Boolean
'------------------------------------------------------------------------
' Valida una contraseña
'------------------------------------------------------------------------
    Dim rsPassword As New ADODB.Recordset
    Dim vlStrSQL As String
    Dim vlstrContrasenaDada As String
    Dim vlblnVencida As Boolean
    Dim lintx As Integer
    Dim lblnPasswordOK As Integer
    Dim lblnUsuarioSeleccionado As Boolean
    Dim lblnUsuarioNoValido As Boolean
    Dim rsUsuarioActivo As New ADODB.Recordset  'CALV
    Dim vlUsuarioActivo As String               'CALV
    
    fblnContrasenaValida = False
    vlblnVencida = False
    If cboEmpresa.Enabled Then
        Set rsPassword = frsEjecuta_SP(UCase(Trim(vlstrNombreUsuario)) & "|" & CStr(cboEmpresa.ItemData(cboEmpresa.ListIndex)), "sp_GnSelDatosLogin")
    Else
        Set rsPassword = frsEjecuta_SP(UCase(Trim(vlstrNombreUsuario)) & "|1", "sp_GnSelDatosLogin")
    End If
    If rsPassword.RecordCount <> 0 Then
    
        vlstrContrasenaDada = fstrEncrypt(UCase(vlstrContrasena), UCase(vlstrNombreUsuario))
        lblnPasswordOK = IIf(vlstrContrasenaDada = rsPassword!Contraseña, 1, 0)
        
        If rsPassword!tipousuario = "M" And rsPassword!UsuarioCompartido = 1 Or rsPassword!tipousuario = "M" Then ' Restringe el acceso al tipo de usuario M = Médico
            'Usuario no válido
            MsgBox SIHOMsg(1350), vbOKOnly + vbExclamation, "Mensaje"
        Else
            If lblnPasswordOK = 1 Then
                If UCase(vlstrNombreUsuario) <> "ADMINISTRADOR" Then
                    'valida usuario activo
                    Set rsUsuarioActivo = frsRegresaRs("SELECT L.BITACTIVO || E.BITACTIVO BITACTIVO, E.VCHAPELLIDOPATERNO || ' ' || E.VCHAPELLIDOMATERNO || ' ' || E.VCHNOMBRE NOMBRE " & _
                                                       "FROM LOGIN L LEFT JOIN NOEMPLEADO E ON E.INTCVEEMPLEADO = L.INTCVEEMPLEADO " & _
                                                       "Where l.VCHUSUARIO = TRIM('" & UCase(vlstrNombreUsuario) & "')")
                    If rsUsuarioActivo.RecordCount > 0 Then
                        If (rsUsuarioActivo!bitactivo = 1) Or (rsUsuarioActivo!bitactivo = 11) Then
                            If fdtmServerFecha > rsPassword!FechaFinal Or fdtmServerFecha < rsPassword!FechaInicial Then
                                vlblnVencida = True
                            End If
                        Else
                            MsgBox "El empleado relacionado con este usuario está dado de baja: " & Chr(13) & _
                                    rsUsuarioActivo!Nombre & ", " & Chr(13) & _
                                    "no se puede ingresar al módulo.", vbOKOnly + vbExclamation, "Mensaje"
                            rsUsuarioActivo.Close
                            Exit Function
                        End If
                    End If
                End If
                If Not vlblnVencida Then
                    If pblnCargaVariablesGlobales Then
                        lblnUsuarioSeleccionado = True
                        If rsPassword!UsuarioCompartido = 1 Then
                            vllngPersonaGraba = flngPersonaGraba(rsPassword!departamento)
                            If vllngPersonaGraba = Null Or vllngPersonaGraba = 0 Then
                                'No se ha seleccionado ningún empleado
                                MsgBox SIHOMsg(1351), vbOKOnly + vbExclamation, "Mensaje"
                                lblnUsuarioSeleccionado = False
                            Else
                                lblnUsuarioSeleccionado = True
                                vglngNumeroEmpleado = vllngPersonaGraba
                                lblnUsuarioSeleccionado = True
                            End If
                        Else
                            vglngNumeroEmpleado = rsPassword!claveempleado
                            lblnUsuarioSeleccionado = True
                        End If
                        If lblnUsuarioSeleccionado = True Then
                              vgintNumeroDepartamento = rsPassword!departamento
                              vgstrNombreDepartamento = rsPassword!nombreDepartamento
                              vglngNumeroLogin = rsPassword!NumeroLogin
                              vgstrNombreUsuario = UCase(vlstrNombreUsuario)
                              vgintClaveEmpresaContable = rsPassword!empresa
                        End If
                    End If
                    If Not lblnUsuarioSeleccionado = False Then
                        fblnContrasenaValida = True
                    End If
                Else
                    'El acceso del usuario expiró
                    MsgBox SIHOMsg(244), vbOKOnly + vbInformation, "Mensaje"
                End If
            Else
                MsgBox SIHOMsg(11), vbOKOnly + vbExclamation, "Mensaje"
            End If
        End If
   Else
        MsgBox SIHOMsg(11), vbOKOnly + vbExclamation, "Mensaje"
   End If
        
End Function


