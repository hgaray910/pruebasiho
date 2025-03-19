VERSION 5.00
Begin VB.Form frmPersonaGraba 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confirmación de contraseña"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8460
   Icon            =   "frmPersonaGraba.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
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
      Height          =   2250
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   8220
      Begin VB.ListBox lstPersonaGraba 
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
         Height          =   1815
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   5160
      End
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
         Left            =   5400
         TabIndex        =   2
         ToolTipText     =   "Cargar enfermeras y empleados del departamento"
         Top             =   255
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
         Left            =   5400
         TabIndex        =   3
         ToolTipText     =   "Cargar médicos"
         Top             =   600
         Width           =   1095
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
         Height          =   1140
         Left            =   5400
         TabIndex        =   5
         Top             =   930
         Width           =   2685
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
            Left            =   120
            PasswordChar    =   "•"
            TabIndex        =   1
            ToolTipText     =   "Contraseña"
            Top             =   600
            Width           =   2445
         End
         Begin VB.Label lblContrasena 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "frmPersonaGraba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------
' Programa para pedir contraseña y password de la persona que desea realizar
' alguna acción, esta forma es llamada por la función flngPersonaGraba
' Fecha de programación: Martes 06 de Febrero de 2001
'----------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
' 21/Octubre/2002
'   Se agrega un parámetro opcional para cargar médicos en la lista
'----------------------------------------------------------------------------
' 28/Julio/2003
'   Se cargan los empleados que tienen asignado este departamento en siDepartamentoEmpleado,
'   no sólo los que son de ese departamento
'----------------------------------------------------------------------------
Option Explicit
Public vlintxDepartamento As Integer
Public vlstrQuienGraba As String            ' "E" = Empleado "M" = Medico "P" = Personas autorizadas a presupuesto
Public vllngEmpleadoSeleccionado As Long    'Esta nos permite regresar cual empleado o médico grabó
Public vlstrPosicionInicialFiltro As String 'Esta nos permite saber en cual de los dos filtros se va ha posicionar (Empleado o Médico)
Public llngCvePersona As Long               'Indica la clave del medico o empleado para que aparezca la persona seleccionada y solo teclear el psw (se utiliza en EX)
Public vlstrFiltro As String                'Filtro adicional para la opción de médicos (Agregado para caso 4889)

Dim rsPersonas As New ADODB.Recordset
Dim vlstrx As String
Dim vlblnCargandoForma As Boolean

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    Dim lintCiclos As Integer
    If fblnAutoVerificacion Then
        pRealizaVerificacion vgintNumeroModulo, True
    End If
    
    If lstPersonaGraba.ListCount = 0 Then
        'No existen empleados activos asignados a este departamento.
        MsgBox SIHOMsg(307), vbOKOnly + vbExclamation, "Mensaje"
        'Se envia 0 para limpiarlo
        vllngEmpleadoSeleccionado = 0
        Unload Me
    Else
        If llngCvePersona <> 0 And vlstrQuienGraba <> "" And vlstrQuienGraba <> "A" Then
            For lintCiclos = 0 To lstPersonaGraba.ListCount - 1
                If llngCvePersona = lstPersonaGraba.ItemData(lintCiclos) Then
                    Exit For
                End If
            Next lintCiclos
            lstPersonaGraba.ListIndex = lintCiclos
            txtPassword.SetFocus
        Else
            lstPersonaGraba.ListIndex = 0
        End If
        optFiltro(0).Visible = vlstrQuienGraba <> ""
        optFiltro(1).Visible = vlstrQuienGraba <> ""
        vllngEmpleadoSeleccionado = 0
        vlstrQuienGraba = ""
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:Form_Activate"))
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 27 Then
        frmPersonaGraba.Hide
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:Form_KeyPress"))
    Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        frmPersonaGraba.Hide
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:Form_QueryUnload"))
    Unload Me
End Sub

Private Sub lstPersonaGraba_DblClick()
    On Error GoTo NotificaError
    
    txtPassword.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:lstPersonaGraba_DblClick"))
    Unload Me
End Sub

Private Sub lstPersonaGraba_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        txtPassword.SetFocus
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:lstPersonaGraba_KeyDown"))
    Unload Me
End Sub

Private Sub optFiltro_Click(Index As Integer)
    On Error GoTo NotificaError
    
    'Iniciacialización del arreglo Auxiliar
    vlstrx = ""
    If optFiltro(0).Value And vlstrQuienGraba <> "P" And cgstrModulo <> "EX" Then
        vgstrParametrosSP = Str(vlintxDepartamento) & "|%|" & 0
        Set rsPersonas = frsEjecuta_SP(vgstrParametrosSP, "SP_EXSELENFERMERASYEMPLEADOS")
                     
    End If
    If optFiltro(0).Value And vlstrQuienGraba <> "P" And cgstrModulo = "EX" Then
         vgstrParametrosSP = Str(vlintxDepartamento) & "|%|" & vgintClaveEmpresaContable
         Set rsPersonas = frsEjecuta_SP(vgstrParametrosSP, "SP_EXSELENFERMERASYEMPLEADOS")
         
     End If
    
    If optFiltro(1).Value Then
        vlstrx = "Select " & _
                    "ltrim(rtrim(vchApellidoPaterno))||' '||ltrim(rtrim(vchApellidoMaterno))||' '||ltrim(rtrim(vchNombre)) Nombre," & _
                    "intCveMedico Clave," & _
                    "'M' Estatus " & _
                 "From HoMedico " & _
                 "Where bitEstaActivo = 1 "
        If Trim(vlstrFiltro) <> "" Then vlstrx = vlstrx & " " & vlstrFiltro & " "
        vlstrx = vlstrx & " order by Nombre"
        Set rsPersonas = frsRegresaRs(vlstrx, adLockReadOnly, adOpenForwardOnly)
    End If
    
        If vlstrQuienGraba = "P" Then
        vlstrx = "select ltrim(rtrim(vchApellidoPaterno))||' '||ltrim(rtrim(vchApellidoMaterno))||' '||ltrim(rtrim(vchNombre)) Nombre," & _
                 "INTCVEEMPLEADO Clave,'E' Estatus  from gnpersonaproceso " & _
                 "inner join noempleado on gnpersonaproceso.INTPERSONA = noempleado.INTCVEEMPLEADO where intproceso = (select intproceso from siproceso where VCHDESCRIPCION  = 'Presupuesto de salidas a departamento') AND INTCVEDEPARTAMENTO = " & vlintxDepartamento & " AND BITPERSONAAUTORIZAR = 1"
        Set rsPersonas = frsRegresaRs(vlstrx, adLockReadOnly, adOpenForwardOnly)
        vlstrQuienGraba = ""
    End If
        
    lstPersonaGraba.Visible = False
    lstPersonaGraba.Clear
    With rsPersonas
        Do While Not .EOF
            lstPersonaGraba.AddItem !Nombre
            lstPersonaGraba.ItemData(lstPersonaGraba.NewIndex) = !clave
            .MoveNext
        Loop
    End With
    lstPersonaGraba.Visible = True
    lstPersonaGraba.Enabled = lstPersonaGraba.ListCount > 0
    If lstPersonaGraba.ListCount > 0 Then
        lstPersonaGraba.ListIndex = 0
    End If
    
    If Not vlblnCargandoForma Then lstPersonaGraba.SetFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:optFiltro_Click"))
    Unload Me
End Sub

Private Sub optFiltro_LostFocus(Index As Integer)
    On Error GoTo NotificaError
    
    If lstPersonaGraba.Enabled Then lstPersonaGraba.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:optFiltro_LostFocus"))
End Sub

Private Sub txtPassword_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtPassword

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:txtPassword_GotFocus"))
    Unload Me
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    Dim rsContrasena As New ADODB.Recordset
    Dim vlstrContrasena As String
    Dim lintx As Integer
    Dim lblnPasswordOK As Integer
    
    If KeyAscii = vbKeyReturn Then
        If lstPersonaGraba.ListIndex <> -1 Then
            If optFiltro(0).Value Then 'Empleados
                vlstrx = "select vchPassword from NoEmpleado where intCveEmpleado=" + Trim(Str(lstPersonaGraba.ItemData(lstPersonaGraba.ListIndex)))
            Else
                vlstrx = "select vchPassword from HoMedico where intCveMedico=" + Trim(Str(lstPersonaGraba.ItemData(lstPersonaGraba.ListIndex)))
            End If
        
            Set rsContrasena = frsRegresaRs(vlstrx)
            
            If rsContrasena.RecordCount <> 0 Then
                vlstrContrasena = fstrEncrypt(txtPassword.Text, lstPersonaGraba.List(lstPersonaGraba.ListIndex))
                lblnPasswordOK = IIf(vlstrContrasena = rsContrasena!vchPassword, 1, 0)
                If lblnPasswordOK = 1 Then
                    vllngEmpleadoSeleccionado = lstPersonaGraba.ItemData(lstPersonaGraba.ListIndex)
                    vlstrQuienGraba = IIf(optFiltro(0).Value, "E", "M")
                    frmPersonaGraba.Hide
                Else
                    'La contraseña no coincide, verificar nuevamente
                    MsgBox SIHOMsg(763), vbOKOnly + vbExclamation, "Mensaje"
                    pEnfocaTextBox txtPassword
                End If
            Else
                MsgBox SIHOMsg(241) + " Esta persona no está registrada.", vbOKOnly + vbInformation, "Mensaje"
            End If
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:txtPassword_KeyPress"))
    Unload Me
End Sub

Public Sub pCargarForma()
    On Error GoTo NotificaError
    vlblnCargandoForma = True
    If vlstrQuienGraba = "" Or vlstrQuienGraba = "E" Or vlstrQuienGraba = "A" Or vlstrQuienGraba = "P" Then
        If vlstrPosicionInicialFiltro <> "M" Then
            optFiltro(0).Value = True
        Else
            optFiltro(1).Value = True
        End If
    ElseIf vlstrQuienGraba = "M" Then
        optFiltro(1).Value = True
    End If
    optFiltro(0).Enabled = Not vlstrQuienGraba = "M"
    optFiltro(1).Enabled = Not vlstrQuienGraba = "E"
    vlblnCargandoForma = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & " - frmPersonaGraba:pCargarForma"))
    Unload Me
End Sub


