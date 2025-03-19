VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaClienteMedEmp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de clientes del departamento"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6825
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
      Left            =   10
      TabIndex        =   1
      Top             =   -75
      Width           =   6765
      Begin VB.TextBox txtBusqueda 
         Height          =   315
         Left            =   200
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Iniciales"
         Top             =   250
         Width           =   4200
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdClientes 
      Height          =   5550
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Clientes encontrados de acuerdo a las iniciales"
      Top             =   675
      Width           =   6765
      _ExtentX        =   11933
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
Attribute VB_Name = "frmBusquedaClienteMedEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Búsqueda de clientes del departamento
' Solo Empleados o Médicos
'----------------------------------------------------------------------------------
Option Explicit

Public llngNumCliente As Long
Public lstrTipoCliente As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = 27 Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
        vgstrNombreForm = Me.Name
        llngNumCliente = 0
        Me.Caption = "Búsqueda de " & IIf(lstrTipoCliente = "ME", "médicos", "empleados") & " con crédito en el departamento"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdClientes_DblClick()
    If grdClientes.Rows > 0 Then
        If Trim(grdClientes.TextMatrix(1, 1)) <> "" Then
            llngNumCliente = grdClientes.TextMatrix(grdClientes.Row, 2)
            Unload Me
        End If
    Else
        txtBusqueda.SetFocus
    End If
End Sub

Private Sub grdClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then grdClientes_DblClick
End Sub

Private Sub pCarga()
    On Error GoTo NotificaError
    Dim rsClientes As New ADODB.Recordset
    Dim vllngContador As Long
    
    With grdClientes
        .Rows = 2
        .Cols = 3
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    End With
    
    vgstrParametrosSP = "0|" & Str(0) & "|" & lstrTipoCliente & "|" & IIf(Trim(txtBusqueda.Text) = "", "%", Trim(txtBusqueda.Text)) & "|" & CStr(vgintClaveEmpresaContable) & "|1"
    Set rsClientes = frsEjecuta_SP(vgstrParametrosSP, "sp_CcSelDatosCliente")
    If rsClientes.RecordCount <> 0 Then
        Do While Not rsClientes.EOF
            grdClientes.TextMatrix(grdClientes.Rows - 1, 1) = rsClientes!NombreCliente
            grdClientes.TextMatrix(grdClientes.Rows - 1, 2) = rsClientes!INTNUMCLIENTE
            rsClientes.MoveNext
            If Not rsClientes.EOF Then grdClientes.Rows = grdClientes.Rows + 1
        Loop
    End If
    rsClientes.Close
    
    With grdClientes
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Cliente|Número cliente"
        .ColWidth(0) = 150
        .ColWidth(1) = 5000
        .ColWidth(2) = 1200
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .Col = 1
        .Row = 1
    End With
    
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
        tmrCarga.Enabled = True
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
    txtBusqueda.SetFocus
End Sub

Private Sub tmrCarga_Timer()
    pCarga
    tmrCarga.Enabled = False
End Sub
