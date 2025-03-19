VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de clientes del departamento"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCarga 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   5760
   End
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   -75
      Width           =   8415
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Iniciales"
         Top             =   285
         Width           =   3900
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "&Todos"
         Height          =   210
         Index           =   0
         Left            =   4155
         TabIndex        =   1
         ToolTipText     =   "Tipo de cliente"
         Top             =   195
         Value           =   -1  'True
         Width           =   870
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "Paciente &interno"
         Height          =   210
         Index           =   1
         Left            =   5280
         TabIndex        =   2
         ToolTipText     =   "Tipo de cliente"
         Top             =   195
         Width           =   1515
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "Paciente e&xterno"
         Height          =   210
         Index           =   2
         Left            =   6795
         TabIndex        =   3
         ToolTipText     =   "Tipo de cliente"
         Top             =   195
         Width           =   1515
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "&Empleado"
         Height          =   210
         Index           =   3
         Left            =   4155
         TabIndex        =   4
         ToolTipText     =   "Tipo de cliente"
         Top             =   405
         Width           =   1125
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "&Médico"
         Height          =   210
         Index           =   4
         Left            =   5280
         TabIndex        =   5
         ToolTipText     =   "Tipo de cliente"
         Top             =   405
         Width           =   1035
      End
      Begin VB.OptionButton optTipoConsulta 
         Caption         =   "Em&presa"
         Height          =   210
         Index           =   5
         Left            =   6795
         TabIndex        =   6
         ToolTipText     =   "Tipo de cliente"
         Top             =   405
         Width           =   1020
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdClientes 
      Height          =   5550
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "Clientes encontrados de acuerdo a las iniciales"
      Top             =   675
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9790
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   -2147483633
      FocusRect       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmBusquedaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Búsqueda de clientes del departamento
' Fecha de programacion: 03 de Julio de 2004
'----------------------------------------------------------------------------------

Option Explicit

Public vllngNumCliente As Long
Public lblnTodosClientes As Boolean
Public lIntActivos As Integer
Dim vlstrTipoCliente As String
Public lintNumeroDepartamentoae As Long
Dim lintNumeroDepartamento As Integer
Public lintConfirmaBusquedaReporte As Integer


Private Sub Form_Activate()
    optTipoConsulta(0).Value = True
    pCarga
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = 27 Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
        vllngNumCliente = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdClientes_DblClick()
    If Trim(grdClientes.TextMatrix(1, 1)) <> "" Then
        vllngNumCliente = grdClientes.TextMatrix(grdClientes.Row, 2)
        Unload Me
    End If
End Sub

Private Sub grdClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdClientes_DblClick
End Sub

Private Sub optTipoConsulta_Click(Index As Integer)
    txtBusqueda.SetFocus
End Sub

Private Sub optTipoConsulta_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtBusqueda.SetFocus
End Sub

Private Sub pCarga()
    On Error GoTo NotificaError
    Dim rsClientes As New ADODB.Recordset
    Dim vllngContador As Long
    
    With grdClientes
        .Rows = 2
        .Cols = 4
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    End With
    
    If optTipoConsulta(0).Value Then vlstrTipoCliente = "*"
    If optTipoConsulta(1).Value Then vlstrTipoCliente = "PI"
    If optTipoConsulta(2).Value Then vlstrTipoCliente = "PE"
    If optTipoConsulta(3).Value Then vlstrTipoCliente = "EM"
    If optTipoConsulta(4).Value Then vlstrTipoCliente = "ME"
    If optTipoConsulta(5).Value Then vlstrTipoCliente = "CO"
    
    If cgstrModulo = "SE" Then
        If lintConfirmaBusquedaReporte = 1 Then
            lintNumeroDepartamento = lintNumeroDepartamentoae
        Else
            lintNumeroDepartamento = vgintNumeroDepartamento
        End If
    Else
        lintNumeroDepartamento = vgintNumeroDepartamento
    End If
    
    vgstrParametrosSP = "0|" & Str(IIf(lblnTodosClientes, 0, lintNumeroDepartamento)) & "|" & vlstrTipoCliente & "|" & IIf(Trim(txtBusqueda.Text) = "", "%", Trim(txtBusqueda.Text)) & "|" & CStr(vgintClaveEmpresaContable & "|" & lIntActivos)
    
    Set rsClientes = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelDatosCliente")
    
    If rsClientes.RecordCount <> 0 Then
        Do While Not rsClientes.EOF
            grdClientes.TextMatrix(grdClientes.Rows - 1, 1) = rsClientes!NombreCliente
            grdClientes.TextMatrix(grdClientes.Rows - 1, 2) = rsClientes!INTNUMCLIENTE
            grdClientes.TextMatrix(grdClientes.Rows - 1, 3) = rsClientes!TipoCliente
            
            rsClientes.MoveNext
            
            If Not rsClientes.EOF Then grdClientes.Rows = grdClientes.Rows + 1
            
        Loop
    End If
    rsClientes.Close
    
    With grdClientes
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Cliente|Número|Tipo"
        .ColWidth(0) = 100
        .ColWidth(1) = 5000
        .ColWidth(2) = 1000
        .ColWidth(3) = 2000
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .Col = 1
        .Row = 1
    End With
    lintConfirmaBusquedaReporte = 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCarga"))
End Sub

Private Sub txtBusqueda_GotFocus()
    On Error GoTo NotificaError
        pSelTextBox txtBusqueda
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_GotFocus"))
End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
        Select Case KeyCode
            Case vbKeyReturn
                grdClientes.SetFocus
            Case vbKeyEscape
                Unload Me
            Case vbKeyDown
                grdClientes.SetFocus
        End Select
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyDown"))
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtBusqueda_KeyPress"))
End Sub

Private Sub txtBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    tmrCarga.Enabled = True
End Sub

Private Sub tmrCarga_Timer()
    pCarga
    tmrCarga.Enabled = False
End Sub
