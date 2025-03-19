VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCargosEliminados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos eliminados"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   3200
      TabIndex        =   21
      Top             =   3450
      Width           =   1140
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCargosEliminados.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Vista previa"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   570
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCargosEliminados.frx":0403
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango de fechas"
      Height          =   735
      Left            =   1250
      TabIndex        =   18
      Top             =   2640
      Width           =   5055
      Begin MSComCtl2.DTPicker dtpRango 
         Height          =   315
         Index           =   0
         Left            =   980
         TabIndex        =   8
         ToolTipText     =   "Fecha inicial"
         Top             =   260
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   16842755
         CurrentDate     =   40544
      End
      Begin MSComCtl2.DTPicker dtpRango 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   9
         ToolTipText     =   "Fecha final"
         Top             =   260
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   16842755
         CurrentDate     =   40544
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   500
         TabIndex        =   20
         Top             =   320
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   2680
         TabIndex        =   19
         Top             =   320
         Width           =   135
      End
   End
   Begin VB.Frame Frame 
      Height          =   2565
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      Begin VB.CheckBox chkSocios 
         Caption         =   "Socios"
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   1120
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2740
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Nombre del paciente"
         Top             =   1600
         Width           =   4400
      End
      Begin VB.TextBox txtCuenta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cuenta del paciente"
         Top             =   1600
         Width           =   1005
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   1650
         TabIndex        =   3
         ToolTipText     =   "Todos"
         Top             =   1210
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optExterno 
         Caption         =   "Externo"
         Height          =   195
         Left            =   3650
         TabIndex        =   5
         ToolTipText     =   "Externo"
         Top             =   1210
         Width           =   900
      End
      Begin VB.OptionButton optInterno 
         Caption         =   "Interno"
         Height          =   195
         Left            =   2650
         TabIndex        =   4
         ToolTipText     =   "Interno"
         Top             =   1210
         Width           =   900
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   1650
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Tipo de paciente o empresa"
         Top             =   2050
         Width           =   5500
      End
      Begin VB.ComboBox cboEmpleado 
         Height          =   315
         Left            =   1650
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Empleado que eliminó los cargos"
         Top             =   700
         Width           =   5500
      End
      Begin VB.ComboBox cboDepartamento 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Departamento en donde se eliminaron los cargos"
         Top             =   250
         Width           =   5500
      End
      Begin VB.Label lblTipoPaciente 
         Caption         =   "Paciente"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   1210
         Width           =   1335
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta"
         Enabled         =   0   'False
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   1660
         Width           =   510
      End
      Begin VB.Label lblTipoPacienteE 
         Caption         =   "Tipo de paciente / empresa"
         Height          =   480
         Left            =   150
         TabIndex        =   14
         Top             =   2020
         Width           =   1455
      End
      Begin VB.Label lblEmpleado 
         Caption         =   "Empleado"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   760
         Width           =   1095
      End
      Begin VB.Label lblDepartamento 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   310
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCargosEliminados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rptReporte As CRAXDRT.Report
Dim rs As New ADODB.Recordset

Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_KeyPress"))
End Sub

Private Sub cboEmpleado_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_KeyPress"))
End Sub

Private Sub cboTipoPaciente_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoPaciente_KeyPress"))
End Sub

Private Sub chkSocios_Click()
    
On Error GoTo NotificaError
    
    If chkSocios.Value Then
        
        lblTipoPaciente.Enabled = False
        lblTipoPacienteE.Enabled = False
        optTodos.Enabled = False
        optInterno.Enabled = False
        optExterno.Enabled = False
        cboTipoPaciente.Enabled = False
    
    Else
    
        lblTipoPaciente.Enabled = True
        lblTipoPacienteE.Enabled = True
        optTodos.Enabled = True
        optInterno.Enabled = True
        optExterno.Enabled = True
        cboTipoPaciente.Enabled = True
        
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkSocios_Click"))

End Sub

Private Sub cmdPrint_Click()
    pImprime "I"
End Sub

Private Sub cmdPreview_Click()
    pImprime "P"
End Sub

Private Sub dtpRango_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":dtpRango_KeyDown"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
  
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
    
    'Departamentos
    Set rs = frsEjecuta_SP("-1|1|*|" & vgintClaveEmpresaContable, "sp_GnSelDepartamento")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1, 3
        cboDepartamento.ListIndex = 0
        cboDepartamento.ListIndex = flngLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
    End If
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2321, 2320), "C", True) Then
        cboDepartamento.Enabled = True
    End If
    
    'Empleados
    Set rs = frsEjecuta_SP("-1|1|" & vgintClaveEmpresaContable, "Sp_GnSelEmpleado")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboEmpleado, rs, 0, 1, 3
        cboEmpleado.ListIndex = 0
    End If
    
    'Tipos de paciente / Empresa
    Set rs = frsEjecuta_SP("2", "sp_GnSelTipoPacienteEmpresa")
    If rs.RecordCount > 0 Then
        pLlenarCboRs cboTipoPaciente, rs, 1, 0, 3
        cboTipoPaciente.ListIndex = 0
    End If
    rs.Close
    
    pInstanciaReporte rptReporte, "rptCargosEliminados.rpt"
    
    dtpRango(0).Value = DateSerial(Year(Date), Month(Date), 1)
    dtpRango(1).Value = DateSerial(Year(Date), Month(Date), Day(Date))
      
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Sub pImprime(strDestino As String)
On Error GoTo NotificaError
Dim rsReporte As New ADODB.Recordset
Dim alstrParametros(8) As String

    vgstrParametrosSP = cboDepartamento.ItemData(cboDepartamento.ListIndex) & "|" & _
                        cboEmpleado.ItemData(cboEmpleado.ListIndex) & "|" & _
                        IIf(optTodos.Value, "*", IIf(optInterno.Value, "I", "E")) & "|" & _
                        IIf(optTodos.Value, "0", CStr(Val(txtCuenta.Text))) & "|" & _
                        cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) & "|" & _
                        Format(dtpRango(0).Value, "dd/mm/yyyy") & "|" & _
                        Format(dtpRango(1).Value, "dd/mm/yyyy") & "|" & _
                        CStr(chkSocios.Value)
                        
    Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvRptCargosEliminados")
    If rsReporte.RecordCount > 0 Then
    
        rptReporte.DiscardSavedData
        
        alstrParametros(0) = "NombreHospital" & ";" & Trim(vgstrNombreHospitalCH) & ";TRUE"
        alstrParametros(1) = "Departamento" & ";" & IIf(cboDepartamento.ItemData(cboDepartamento.ListIndex) = 0, "TODOS", cboDepartamento.List(cboDepartamento.ListIndex)) & ";STRING"
        alstrParametros(2) = "Empleado" & ";" & IIf(cboEmpleado.ItemData(cboEmpleado.ListIndex) = 0, "TODOS", cboEmpleado.List(cboEmpleado.ListIndex)) & ";STRING"
        alstrParametros(3) = "Paciente" & ";" & IIf(optTodos.Value, "TODOS", IIf(optInterno.Value, "INTERNOS", "EXTERNOS")) & ";STRING"
        alstrParametros(4) = "Cuenta" & ";" & IIf(optTodos.Value, "", txtCuenta.Text) & ";STRING"
        alstrParametros(5) = "NombrePaciente" & ";" & IIf(optTodos.Value, " ", txtNombre.Text) & ";STRING"
        alstrParametros(6) = "TipoPaciente" & ";" & IIf(cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) = 0, "TODOS", cboTipoPaciente.List(cboTipoPaciente.ListIndex)) & ";STRING"
        alstrParametros(7) = "FechaInicial" & ";" & Format(dtpRango(0).Value, "DD/MMM/YYYY") & ";DATE"
        alstrParametros(8) = "FechaFinal" & ";" & Format(dtpRango(1).Value, "DD/MMM/YYYY") & ";DATE"
        
        pCargaParameterFields alstrParametros, rptReporte
        pImprimeReporte rptReporte, rsReporte, strDestino, "Cargos eliminados"
        
    Else
        MsgBox SIHOMsg(13), vbInformation + vbOKOnly, "Mensaje"
    End If
    rsReporte.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub optExterno_Click()
On Error GoTo NotificaError

    lblCuenta.Enabled = True
    txtNombre.Text = ""
    txtCuenta.Text = ""
    txtCuenta.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optExterno_Click"))
End Sub

Private Sub optExterno_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        pEnfocaTextBox txtCuenta
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optExterno_KeyPress"))
End Sub

Private Sub optInterno_Click()
On Error GoTo NotificaError

    lblCuenta.Enabled = True
    txtNombre.Text = ""
    txtCuenta.Text = ""
    txtCuenta.Enabled = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optInterno_Click"))
End Sub

Private Sub optInterno_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        pEnfocaTextBox txtCuenta
    End If
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optInterno_KeyPress"))
End Sub

Private Sub optTodos_Click()
On Error GoTo NotificaError

    lblCuenta.Enabled = False
    txtNombre.Text = ""
    txtCuenta.Text = ""
    txtCuenta.Enabled = False
    cboTipoPaciente.Enabled = True
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTodos_Click"))
End Sub

Private Sub optTodos_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optTodos_KeyPress"))
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
            
        If Trim(txtCuenta.Text) = "" Then
        
            cboTipoPaciente.Enabled = True
            
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = IIf(optInterno.Value, "I", "E")
                .Caption = .Caption & IIf(optInterno.Value, " internos", " externos")
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                .optTodos.Value = True
                .vgstrTamanoCampo = "950,3400,1000,950"
                
                txtCuenta.Text = .flngRegresaPaciente()
                
                If Val(txtCuenta.Text) > 0 Then
                    txtCuenta_KeyDown vbKeyReturn, 0
                Else
                    txtCuenta.Text = ""
                    txtCuenta.SetFocus
                End If
            End With
        
        Else
        
            vgstrParametrosSP = txtCuenta.Text & "|" & "0" & "|" & IIf(optInterno.Value, "I", "E") & "|" & vgintClaveEmpresaContable
            Set rs = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelDatosPaciente")
            If rs.RecordCount <> 0 Then
            
                txtNombre.Text = rs!Nombre
                cboTipoPaciente.ListIndex = 0
                cboTipoPaciente.Enabled = False
                dtpRango(0).SetFocus
                
            Else
            
                txtNombre.Text = ""
                txtCuenta.Text = ""
                
            End If
            rs.Close
        
        End If

   End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuenta_KeyDown"))
    Unload Me
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        
        If UCase(Chr(KeyAscii)) = "E" Or UCase(Chr(KeyAscii)) = "I" Then
            optInterno.Value = UCase(Chr(KeyAscii)) = "I"
            optExterno.Value = UCase(Chr(KeyAscii)) = "E"
        End If
        
        KeyAscii = 7
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuenta_KeyPress"))
End Sub

Private Sub txtCuenta_LostFocus()
On Error GoTo NotificaError

    If Trim(txtCuenta.Text) = "" Then
        txtNombre.Text = ""
    ElseIf txtNombre.Text = "" Then
        txtCuenta.Text = ""
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCuenta_LostFocus"))
End Sub
