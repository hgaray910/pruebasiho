VERSION 5.00
Begin VB.Form frmPersonalAutorizado 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal autorizado"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5655
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Label lblMensaje 
         BackColor       =   &H80000005&
         Caption         =   "** Mensaje del proceso **"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5175
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   5415
      Begin VB.OptionButton optFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Enfermeras y empleados"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Mostrar enfermeras y empleados"
         Top             =   300
         Value           =   -1  'True
         Width           =   2760
      End
      Begin VB.OptionButton optFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Médicos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Mostrar médicos"
         Top             =   300
         Width           =   1095
      End
      Begin VB.ListBox lstPersonaAutorizada 
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
         Height          =   2070
         ItemData        =   "frmPersonalAutorizado.frx":0000
         Left            =   130
         List            =   "frmPersonalAutorizado.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Médicos/empleados autorizados"
         Top             =   600
         Width           =   5160
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Contraseña"
         Top             =   2710
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Left            =   1800
         TabIndex        =   7
         Top             =   2770
         Width           =   1215
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmPersonalAutorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lintClaveProceso As Integer          'Indica la clave del proceso del cual se requiere autorización
Public lstrTipoPersonaAutoriza As String    'M=médico, E=empleado, Persona que autoriza
Public llngCvePersonaAutoriza As Long       'Clave del médico o del empleado
Public lstrFechaAutorizacion As String      'Fecha de autorización

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError
    Dim rsProceso As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    Set rsProceso = frsEjecuta_SP(CStr(lintClaveProceso), "SP_SISELPROCESOS")
    If rsProceso.RecordCount <> 0 Then
        lblMensaje.Caption = IIf(IsNull(rsProceso!vchmensaje), "", Trim(rsProceso!vchmensaje))
    End If
    llngCvePersonaAutoriza = 0
    lstrTipoPersonaAutoriza = ""
    lstrFechaAutorizacion = ""
    optFiltro(0).Value = True
    optFiltro_Click 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub lstPersonaAutorizada_DblClick()
    txtPassword.SetFocus
End Sub

Private Sub lstPersonaAutorizada_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtPassword.SetFocus
    End If
End Sub

Private Sub optFiltro_Click(Index As Integer)
On Error GoTo NotificaError
    Dim rsPersonas As New ADODB.Recordset
    
    lstPersonaAutorizada.Clear
    vgstrParametrosSP = lintClaveProceso & IIf(Index = 0, "|E|", "|M|") & vgintNumeroDepartamento & "|" & vgintClaveEmpresaContable
    Set rsPersonas = frsEjecuta_SP(vgstrParametrosSP, "SP_SISELPERSONASAUTORIZADAS")
    Do While Not rsPersonas.EOF
        lstPersonaAutorizada.AddItem rsPersonas!Nombre
        lstPersonaAutorizada.ItemData(lstPersonaAutorizada.NewIndex) = rsPersonas!clave
        rsPersonas.MoveNext
    Loop
    If lstPersonaAutorizada.ListCount > 0 Then lstPersonaAutorizada.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optFiltro_Click"))
End Sub

Private Sub optFiltro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstPersonaAutorizada.SetFocus
    End If
End Sub

Private Sub txtPassword_GotFocus()
    pSelTextBox txtPassword
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim rsContrasena As New ADODB.Recordset
    Dim lstrContrasena As String
    Dim lintx As Integer
    Dim lblnPasswordOK As Integer
    Dim lstrsql As String
    
    If KeyCode = vbKeyReturn Then
        If lstPersonaAutorizada.ListIndex <> -1 Then
            If optFiltro(0).Value Then 'Empleados
                lstrsql = "select vchPassword from NoEmpleado where intCveEmpleado = " & lstPersonaAutorizada.ItemData(lstPersonaAutorizada.ListIndex)
            Else
                lstrsql = "select vchPassword from HoMedico where intCveMedico = " & lstPersonaAutorizada.ItemData(lstPersonaAutorizada.ListIndex)
            End If
            Set rsContrasena = frsRegresaRs(lstrsql)
            If rsContrasena.RecordCount <> 0 Then
                lstrContrasena = fstrEncrypt(txtPassword.Text, lstPersonaAutorizada.List(lstPersonaAutorizada.ListIndex))
                lblnPasswordOK = IIf(lstrContrasena = rsContrasena!vchPassword, 1, 0)
                If lblnPasswordOK = 1 Then
                    llngCvePersonaAutoriza = lstPersonaAutorizada.ItemData(lstPersonaAutorizada.ListIndex)
                    lstrTipoPersonaAutoriza = IIf(optFiltro(0).Value, "E", "M")
                    lstrFechaAutorizacion = Format(fdtmServerFechaHora, "YYYY-MM-DD HH:NN:SS")
                    frmPersonalAutorizado.Hide
                Else
                    'La contraseña no coincide, verificar nuevamente
                    MsgBox SIHOMsg(763), vbOKOnly + vbExclamation, "Mensaje"
                    pEnfocaTextBox txtPassword
                End If
            Else
            '    MsgBox SIHOMsg(241) + " Esta persona no está registrada.", vbOKOnly + vbInformation, "Mensaje"
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPassword_KeyPress"))
End Sub

