VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "hsflatcontrols.ocx"
Begin VB.Form frmEgreso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egreso del paciente"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   75
      TabIndex        =   4
      Top             =   0
      Width           =   8340
      Begin VB.TextBox txtTipoCuartoActual 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Tipo de cuarto actual o sugerido"
         Top             =   300
         Width           =   2205
      End
      Begin VB.TextBox txtAreaActual 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Area actual o sugerida"
         Top             =   700
         Width           =   2205
      End
      Begin VB.TextBox txtCuartoActual 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Cuarto actual o sugerido"
         Top             =   1110
         Width           =   2205
      End
      Begin HSFlatControls.MyCombo cboEstadoCuarto 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Selección del estado del cuarto al egresar el paciente"
         Top             =   1110
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Style           =   1
         Enabled         =   -1  'True
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
      Begin MSMask.MaskEdBox mskHoraEgreso 
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "Hora del egreso"
         Top             =   705
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaEgreso 
         Height          =   375
         Left            =   2040
         TabIndex        =   0
         ToolTipText     =   "Fecha de egreso"
         Top             =   300
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Tipo de cuarto"
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
         Left            =   4320
         TabIndex        =   13
         Top             =   315
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Área"
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
         Left            =   4320
         TabIndex        =   12
         Top             =   765
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Cuarto actual"
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
         Left            =   4320
         TabIndex        =   11
         Top             =   1170
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Fecha de egreso"
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
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Hora de egreso"
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
         Left            =   135
         TabIndex        =   9
         Top             =   760
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Estado del cuarto"
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
         Left            =   135
         TabIndex        =   8
         Top             =   1170
         Width           =   1725
      End
   End
   Begin VB.Frame fraGrabar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   3990
      TabIndex        =   14
      Top             =   1530
      Width           =   720
      Begin MyCommandButton.MyButton cmdEgresar 
         Height          =   600
         Left            =   60
         TabIndex        =   3
         Top             =   200
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         MaskColor       =   16777215
         Picture         =   "frmEgreso.frx":0000
         BackColorDown   =   -2147483643
         BackColorOver   =   -2147483633
         BackColorFocus  =   -2147483633
         BackColorDisabled=   -2147483633
         BorderColor     =   -2147483627
         TransparentColor=   16777215
         Caption         =   ""
         DepthEvent      =   1
         PictureDisabled =   "frmEgreso.frx":0984
         PictureAlignment=   4
         PictureDisabledEffect=   0
         ShowFocus       =   -1  'True
      End
   End
End
Attribute VB_Name = "frmEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Admisión
'| Nombre del Formulario    : frmEgreso
'-------------------------------------------------------------------------------------
'| Objetivo: Permite el egresar al paciente y asignar un estado al cuarto ocupado
'-------------------------------------------------------------------------------------

Public vllngNumeroCuenta As Long        'Número de cuenta del paciente que se desea egresar
Public vlblnEgresoPaciente As Boolean   'Indica si se efectuó el egreso del paciente
Public vlblnesexterno As Boolean 'Indica si el paciente que se selecciono fue externo
Public vllngNumeroCuentaexterno As Long 'Indica numero de cuenta del paciente externo

Dim rs As New ADODB.Recordset
Dim vllngNumeroCuentaRN As Long
Dim vldtmFechaEgreso As Date
Dim vlblnRecienNacidoInternado As Boolean
Dim vllngPersonaGraba As String
Dim vlintMensajeFalla As Integer

Private Sub cboEstadoCuarto_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        If fblnCanFocus(cmdEgresar) Then cmdEgresar.SetFocus
    End If

End Sub

Private Sub pEgresarRecienNacido()
On Error GoTo NotificaError
    Dim vlstrSentencia As String

    If Not vlblnRecienNacidoInternado Then
        '¿ Desea egresar tambien al recien nacido ?
        If MsgBox(SIHOMsg(600), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
            pEgresarPersona vllngNumeroCuentaRN
        Else
            EntornoSIHO.ConeccionSIHO.BeginTrans
                frsEjecuta_SP CStr(vllngNumeroCuentaRN), "SP_EXUPDINGRESORECIENNACIDO"
            EntornoSIHO.ConeccionSIHO.CommitTrans
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEgresarRecienNacido"))
End Sub

Private Function fTieneRecienNacido(llngCuentaMama As Long) As Boolean
    Dim rsAdAdmisionRN As New ADODB.Recordset
    Dim vlstrSentencia As String

    vllngNumeroCuentaRN = 0
    fTieneRecienNacido = False
    
    vlstrSentencia = "SELECT * FROM AdAdmision WHERE numNumCuentaRel = " & llngCuentaMama & " AND CHRESTATUSADMISION = 'A'"
    Set rsAdAdmisionRN = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)

    If rsAdAdmisionRN.RecordCount > 0 Then
        vllngNumeroCuentaRN = rsAdAdmisionRN!numNumCuenta
        fTieneRecienNacido = True
        vlblnRecienNacidoInternado = IIf(rsAdAdmisionRN!chrTipoIngreso = "1", True, False)
    End If
    rsAdAdmisionRN.Close
    
End Function
Private Sub pEgresarPersona(vllngNoCuenta As Long)
On Error GoTo NotificaError
    Dim rsAdAdmision As New ADODB.Recordset
    Dim rsAdCuarto As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim rsregistroexterno As New ADODB.Recordset
       
    vlblnEgresoPaciente = False
    EntornoSIHO.ConeccionSIHO.BeginTrans
    
    vgstrParametrosSP = IIf(vlblnesexterno, "E", "I") & "|" & _
                         IIf(vlblnesexterno, vllngNumeroCuentaexterno, vllngNoCuenta) & "|" & _
                         fstrFechaSQL(mskFechaEgreso.Text, mskHoraEgreso.Text) & "|" & _
                         vllngPersonaGraba & "|" & _
                         cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex)
                         
    frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDEGRESOPACIENTE"
        
    EntornoSIHO.ConeccionSIHO.CommitTrans
    vlblnEgresoPaciente = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEgresarPersona"))
End Sub
Private Sub cmdEgresar_Click()
   If vgintNumeroModulo <> 16 And fblnRevisaPermiso(vglngNumeroLogin, IIf(vgintNumeroModulo = 5, 467, IIf(vgintNumeroModulo = 9, 2446, 753)), "E") Then
        If fblnDatosValidos() Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba <> 0 Then
                pEgresarPersona vllngNumeroCuenta
                If vlblnEgresoPaciente Then
                    If fTieneRecienNacido(vllngNumeroCuenta) Then
                        pEgresarRecienNacido
                    End If
                End If
                If vlblnEgresoPaciente Then
                    'La operación se realizó satisfactoriamente.
                    MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
                    pGuardarLogTransaccion Me.Name, EnmGrabar, frmPersonaGraba.vllngEmpleadoSeleccionado, "EGRESO DESDE PANTALLA (CUENTA)", Str(vllngNumeroCuenta)
                End If
                Unload Me
            End If
        End If
    End If
End Sub

Private Function fblnDatosValidos() As Boolean
    fblnDatosValidos = True
    Dim rs As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If fblnDatosValidos And Not IsDate(mskHoraEgreso.Text) Then
        fblnDatosValidos = False
        '¡Hora no válida!, formato de hora hh:mm
        MsgBox SIHOMsg(41), vbOKOnly + vbInformation, "Mensaje"
        mskHoraEgreso.SetFocus
        Exit Function
    End If
    If IsDate(mskFechaEgreso.Text) Then
       If CDate(Format(((mskFechaEgreso.Text) + " " + (mskHoraEgreso.Text)), "DD/MM/YYYY HH:MM:SS")) > CDate(Format(fdtmServerFechaHora, "DD/MM/YYYY HH:MM:SS")) Then
            '¡Fecha no válida!
            fblnDatosValidos = False
            MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
            mskFechaEgreso.SetFocus
            Exit Function
       Else
            vlstrSentencia = "select * from expacienteingreso where expacienteingreso.intnumcuenta = " & vllngNumeroCuenta
            Set rs = frsRegresaRs(vlstrSentencia)
            If CDate(Format(((mskFechaEgreso.Text) + " " + (mskHoraEgreso.Text)), "DD/MM/YYYY HH:MM:SS")) < CDate(Format(rs!DTMFECHAHORAINGRESO, "DD/MM/YYYY HH:MM:SS")) Then
                '¡Fecha no válida!
                fblnDatosValidos = False
                MsgBox "¡Fecha no válida!", vbOKOnly + vbInformation, "Mensaje"
                mskFechaEgreso.SetFocus
                Exit Function
            Else
                fblnDatosValidos = True
            End If
       End If
    Else
        fblnDatosValidos = False
        '¡Fecha no válida!
        MsgBox SIHOMsg(254), vbOKOnly + vbInformation, "Mensaje"
        mskFechaEgreso.SetFocus
        Exit Function
    End If
     
      
    If fblnDatosValidos And cboEstadoCuarto.ListIndex = -1 Then
        fblnDatosValidos = False
        '¡Dato no válido, seleccione un valor de la lista!
        MsgBox SIHOMsg(3), vbOKOnly + vbInformation, "Mensaje"
        If fblnCanFocus(cboEstadoCuarto) Then cboEstadoCuarto.SetFocus
    End If

End Function

Private Sub Form_Activate()
    If cboEstadoCuarto.ListCount = 0 Then
        'No se encontraron estados de cuarto disponibles.
        MsgBox SIHOMsg(577), vbExclamation + vbOKOnly, "Mensaje"
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim vlstrSentencia  As String
    Dim vlstrCuartoActual As String
    Dim vlstrCuartoActualexterno As String
    Dim vlstrcuarto As String
    Dim rsCuarto As New ADODB.Recordset
    Dim vllngCveEstadoCuartoDisponible As Long
    
    vllngCveEstadoCuartoOcupado = vglngCveEstadoCuartoOcupado
    vllngCveEstadoCuartoDisponible = vglngCveEstadoCuartoDisponible
    
    Me.Icon = frmMenuPrincipal.Icon
    vlblnEgresoPaciente = False
    
    mskFechaEgreso.Mask = ""
    mskFechaEgreso.Text = fdtmServerFecha
    mskFechaEgreso.Mask = "##/##/####"
    
    mskHoraEgreso.Mask = ""
    mskHoraEgreso.Text = IIf(Len(Trim(Str(Hour(fdtmServerHora)))) = 1, "0" + Trim(Str(Hour(fdtmServerHora))), Trim(Str(Hour(fdtmServerHora)))) + ":" + IIf(Len(Trim(Str(Minute(fdtmServerHora)))) = 1, "0" + Trim(Str(Minute(fdtmServerHora))), Trim(Str(Minute(fdtmServerHora))))
    mskHoraEgreso.Mask = "##:##"
    
    'Cargar estados de cuarto activos
    vlstrSentencia = "select tnyCveEstadoCuarto, vchDescripcion from AdEstadoCuarto where bitActivo = 1 and tnyCveEstadoCuarto <> " & Str(vglngCveEstadoCuartoOcupado)
    Set rs = frsRegresaRs(vlstrSentencia)
    
    If rs.RecordCount <> 0 Then
        pLlenarCboRs_new cboEstadoCuarto, rs, 0, 1
        If vlblnesexterno = False Then
             cboEstadoCuarto.ListIndex = 0
        Else
            If vllngCveEstadoCuartoDisponible <> 0 Then
                cboEstadoCuarto.ListIndex = flngLocalizaCbo_new(cboEstadoCuarto, CStr(vllngCveEstadoCuartoDisponible))
            Else
                cboEstadoCuarto.ListIndex = 0
            End If
        End If
    End If
    
    'Cargar datos del cuarto actual del paciente
        If vlblnesexterno = True Then
            
            vlstrCuartoActualexterno = "select isnull(vchNumCuarto,'')nombrecuarto from registroexterno where intNumCuenta =" & vllngNumeroCuentaexterno
            Set rsCuarto = frsRegresaRs(vlstrCuartoActualexterno)
          
            If rsCuarto.RecordCount > 0 Then
                vlstrcuarto = rsCuarto!nombrecuarto
                vlstrSentencia = "" & _
                "select " & _
                    "AdTipoCuarto.vchDescripcion TipoCuarto," & _
                    "AdArea.vchDescripcion Area," & _
                    "AdCuarto.vchDescripcion CuartoActual " & _
                "From " & _
                    "AdCuarto " & _
                    "inner join AdTipoCuarto on " & _
                    "AdCuarto.tnyCveTipoCuarto = AdTipoCuarto.tnyCveTipoCuarto " & _
                    "inner join AdArea on " & _
                    "AdCuarto.tnyCveArea = AdArea.tnyCveArea " & _
                "Where " & _
                    "ltrim(rtrim(AdCuarto.vchNumCuarto)) = '" & Trim(vlstrcuarto) & "'"
                
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount <> 0 Then
                    txtTipoCuartoActual.Text = Trim(rs!TipoCuarto)
                    txtAreaActual.Text = Trim(rs!Area)
                    txtCuartoActual.Text = Trim(rs!CuartoActual)
                End If
            End If
        Else
            vlstrCuartoActual = Trim(frsRegresaRs("select isnull(vchNumCuarto,' ') from AdAdmision where numNumCuenta = " & Str(vllngNumeroCuenta)).Fields(0))
            If vlstrCuartoActual <> "" Then
                vlstrSentencia = "" & _
                "select " & _
                    "AdTipoCuarto.vchDescripcion TipoCuarto," & _
                    "AdArea.vchDescripcion Area," & _
                    "AdCuarto.vchDescripcion CuartoActual " & _
                "From " & _
                    "AdCuarto " & _
                    "inner join AdTipoCuarto on " & _
                    "AdCuarto.tnyCveTipoCuarto = AdTipoCuarto.tnyCveTipoCuarto " & _
                    "inner join AdArea on " & _
                    "AdCuarto.tnyCveArea = AdArea.tnyCveArea " & _
                "Where " & _
                    "ltrim(rtrim(AdCuarto.vchNumCuarto)) = '" & Trim(vlstrCuartoActual) & "'"
                Set rs = frsRegresaRs(vlstrSentencia)
                If rs.RecordCount <> 0 Then
                    txtTipoCuartoActual.Text = Trim(rs!TipoCuarto)
                    txtAreaActual.Text = Trim(rs!Area)
                    txtCuartoActual.Text = Trim(rs!CuartoActual)
                End If
            Else
                'Aqui entran los pacientes ambulatorios los cuales se realiza un egreso pero
                'no se libera cuarto ya que no se asigna en la admisión
                txtTipoCuartoActual.Text = ""
                txtAreaActual.Text = ""
                txtCuartoActual.Text = ""
                cboEstadoCuarto.Enabled = False
            End If
       End If
End Sub

Private Sub mskFechaEgreso_GotFocus()
    pSelMkTexto mskFechaEgreso
End Sub

Private Sub mskFechaEgreso_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        mskHoraEgreso.SetFocus
    End If

End Sub

Private Sub mskHoraEgreso_GotFocus()
    pSelMkTexto mskHoraEgreso
End Sub

Private Sub mskHoraEgreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboEstadoCuarto.Enabled = True Then
            If fblnCanFocus(cboEstadoCuarto) Then cboEstadoCuarto.SetFocus
        Else
            If fblnCanFocus(cmdEgresar) Then cmdEgresar.SetFocus
        End If
    End If
End Sub
