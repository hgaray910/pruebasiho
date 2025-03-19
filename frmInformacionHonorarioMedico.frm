VERSION 5.00
Begin VB.Form frmInformacionHonorarioMedico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información del honorario médico"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7185
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   3285
      TabIndex        =   10
      ToolTipText     =   "Siguiente pago"
      Top             =   2880
      Width           =   615
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   65
         MaskColor       =   &H80000000&
         Picture         =   "frmInformacionHonorarioMedico.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar "
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtImporteHonorarioMedico 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   3
         Text            =   "$0.00"
         ToolTipText     =   "Importe del honorario médico"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtImporteFacturarConvenio 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Text            =   "$0.00"
         ToolTipText     =   "Importe a facturar"
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox cboMedico 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Médico"
         Top             =   480
         Width           =   6735
      End
      Begin VB.TextBox txtProcedimiento 
         Height          =   855
         Left            =   120
         MaxLength       =   9999
         MultiLine       =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Procedimiento(s) realizado(s)"
         Top             =   1080
         Width           =   6735
      End
      Begin VB.Label lblTotalHonorario 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$0.00"
         Height          =   195
         Left            =   6000
         TabIndex        =   18
         Top             =   2460
         Width           =   405
      End
      Begin VB.Label lblTotalAFacturar 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "$0.00"
         Height          =   195
         Left            =   6000
         TabIndex        =   17
         Top             =   2100
         Width           =   405
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   16
         Top             =   2460
         Width           =   120
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   2100
         Width           =   120
      End
      Begin VB.Label lblCantidadCargo2 
         AutoSize        =   -1  'True
         Caption         =   "99"
         Height          =   195
         Left            =   4680
         TabIndex        =   14
         Top             =   2460
         Width           =   180
      End
      Begin VB.Label lblCantidadCargo1 
         AutoSize        =   -1  'True
         Caption         =   "10"
         Height          =   195
         Left            =   4680
         TabIndex        =   13
         Top             =   2100
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   12
         Top             =   2460
         Width           =   105
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   11
         Top             =   2100
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Importe del honorario médico *"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2460
         Width           =   2145
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Importe a facturar *"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2100
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Médico *"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procedimiento(s) *"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmInformacionHonorarioMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vgblnInformacionValida As Boolean
Public vgintCveMedico As Integer
Public vgstrProcedimiento As String
Public vgdblImporteCargo As Double
Public vgdblImporteHonorario As Double
Public vglngCveCargo As Long
Public vgintCantidadCargo As Integer
Dim vgblnGraboInformacion As Boolean
Dim vgstrEstado As String


Private Sub cmdSave_Click()
    Dim strSentencia As String
    
    If fblnDatosValidos Then
        '  Esta forma solo valida que se hayan capturado todos los datos, la forma frmCargosDirectos es la que graba la información
        vgblnGraboInformacion = True
        vgintCveMedico = frmInformacionHonorarioMedico.cboMedico.ItemData(frmInformacionHonorarioMedico.cboMedico.ListIndex)
        vgstrProcedimiento = frmInformacionHonorarioMedico.txtProcedimiento.Text
        vgdblImporteCargo = Val(Format(frmInformacionHonorarioMedico.txtImporteFacturarConvenio.Text, "############.##"))
        vgdblImporteHonorario = Val(Format(frmInformacionHonorarioMedico.txtImporteHonorarioMedico.Text, "############.##"))
        Unload Me
    End If
End Sub


Private Function fblnDatosValidos() As Boolean
    Dim strSentencia As String
    Dim rsCuentaMedico As New ADODB.Recordset
    
    fblnDatosValidos = False
    If cboMedico.ListIndex = -1 Then
        '¡Falta información!
        MsgBox SIHOMsg(530), vbCritical, "Mensaje"
        If fblnCanFocus(cboMedico) Then cboMedico.SetFocus
        Exit Function
    End If
    If Trim(txtProcedimiento.Text) = "" Then
        '¡Falta información!
        MsgBox SIHOMsg(530), vbCritical, "Mensaje"
        If fblnCanFocus(txtProcedimiento) Then txtProcedimiento.SetFocus
        Exit Function
    End If
    If Str(Val(Format(txtImporteFacturarConvenio.Text, "#############.00"))) = 0 Then
        '¡Falta información!
        MsgBox SIHOMsg(530), vbCritical, "Mensaje"
        If fblnCanFocus(txtImporteFacturarConvenio) Then txtImporteFacturarConvenio.SetFocus
        Exit Function
    End If
    If Str(Val(Format(txtImporteHonorarioMedico.Text, "#############.00"))) = 0 Then
        '¡Falta información!
        MsgBox SIHOMsg(530), vbCritical, "Mensaje"
        If fblnCanFocus(txtImporteHonorarioMedico) Then txtImporteHonorarioMedico.SetFocus
        Exit Function
    End If
    
    'Valida que el médico tenga una cuenta contable configurada
    strSentencia = "Select NVL(HoMedicoEmpresa.INTNUMEROCUENTA, 0) CuentaMedico " & _
                   "  From HoMedicoEmpresa " & _
                   " Where HoMedicoEmpresa.INTCLAVEMEDICO = " & frmInformacionHonorarioMedico.cboMedico.ItemData(frmInformacionHonorarioMedico.cboMedico.ListIndex) & _
                   "   And HoMedicoEmpresa.TNYCLAVEEMPRESA = " & vgintClaveEmpresaContable
    Set rsCuentaMedico = frsRegresaRs(strSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsCuentaMedico.RecordCount <> 0 Then
        If rsCuentaMedico!CuentaMedico = 0 Then
            'El médico no tiene una cuenta contable asignada, verifíquelo.
            MsgBox SIHOMsg(519), vbCritical, "Mensaje"
            If fblnCanFocus(cboMedico) Then cboMedico.SetFocus
            Exit Function
        End If
    Else
        'El médico no tiene una cuenta contable asignada, verifíquelo.
        MsgBox SIHOMsg(519), vbCritical, "Mensaje"
        If fblnCanFocus(cboMedico) Then cboMedico.SetFocus
        Exit Function
    End If
    
    If Val(Format(txtImporteHonorarioMedico.Text, "#############.00")) > Val(Format(txtImporteFacturarConvenio.Text, "#############.00")) Then
        '¡Falta información!
        MsgBox "El importe a facturar debe ser mayor al importe del honorario médico.", vbCritical, "Mensaje"
        If fblnCanFocus(txtImporteFacturarConvenio) Then txtImporteFacturarConvenio.SetFocus
        Exit Function
    End If
    fblnDatosValidos = True
End Function

Private Sub Form_Activate()
    vgblnInformacionValida = False
    vgblnGraboInformacion = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys (vbTab)
        Case vbKeyEscape
            Unload Me
    End Select

End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    lblCantidadCargo1.Caption = vgintCantidadCargo
    lblCantidadCargo2.Caption = vgintCantidadCargo
    pLlenaCombo "Select HoMedico.INTCVEMEDICO Cve, HoMedico.VCHAPELLIDOPATERNO || ' ' || HoMedico.VCHAPELLIDOMATERNO || ' ' || HoMedico.VCHNOMBRE Nombre from HoMedico Where HoMedico.BITESTAACTIVO = 1 Order by Nombre", cboMedico, 0
    cboMedico.ListIndex = -1
    pCargaInformacion
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not vgblnGraboInformacion And vgstrEstado = "NUEVO" Then
        '|No se realizará el cargo. ¿Desea abandonar la operación?
        If MsgBox("No se realizará el cargo." & vbCrLf & SIHOMsg(17), vbExclamation + vbYesNo, "Mensaje") = vbNo Then
            Cancel = True
        End If
    End If
    vgblnInformacionValida = vgblnGraboInformacion

End Sub

Private Sub txtImporteFacturarConvenio_GotFocus()
On Error GoTo NotificaError
    
    txtImporteFacturarConvenio.Text = Str(Val(Format(txtImporteFacturarConvenio.Text, "###############.00")))
    pSelTextBox txtImporteFacturarConvenio
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteFacturarConvenio_GotFocus"))
    Unload Me
End Sub

Private Sub txtImporteFacturarConvenio_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteFacturarConvenio)) Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteFacturarConvenio_KeyPress"))
    Unload Me

End Sub

Private Sub txtImporteFacturarConvenio_LostFocus()
    Dim dblTotal As Double
    
On Error GoTo NotificaError
    
    txtImporteFacturarConvenio.Text = FormatCurrency(Val(txtImporteFacturarConvenio.Text), 2)
    
    lblTotalAFacturar.Caption = Val(Format(txtImporteFacturarConvenio.Text, "###############.00")) * Val(Format(lblCantidadCargo1, "###############.00"))
    lblTotalAFacturar.Caption = FormatCurrency(Val(lblTotalAFacturar.Caption), 2)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteFacturarConvenio_LostFocus"))
    Unload Me

End Sub


Private Sub txtImporteHonorarioMedico_GotFocus()
On Error GoTo NotificaError
    
    txtImporteHonorarioMedico.Text = Str(Val(Format(txtImporteHonorarioMedico.Text, "###############.00")))
    pSelTextBox txtImporteHonorarioMedico
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteHonorarioMedico_GotFocus"))
    Unload Me
End Sub

Private Sub txtImporteHonorarioMedico_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtImporteHonorarioMedico)) Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteHonorarioMedico_KeyPress"))
    Unload Me

End Sub

Private Sub txtImporteHonorarioMedico_LostFocus()
On Error GoTo NotificaError
    
    txtImporteHonorarioMedico.Text = FormatCurrency(Val(txtImporteHonorarioMedico.Text), 2)

    lblTotalHonorario.Caption = Val(Format(txtImporteHonorarioMedico.Text, "###############.00")) * Val(Format(lblCantidadCargo1, "###############.00"))
    lblTotalHonorario.Caption = FormatCurrency(Val(lblTotalHonorario.Caption), 2)


    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtImporteHonorarioMedico_LostFocus"))
    Unload Me
End Sub

Private Sub txtProcedimiento_GotFocus()
    pSelTextBox txtProcedimiento
End Sub

Private Sub txtProcedimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub pCargaInformacion()
    Dim strSentencia As String
    Dim rsInfo As New ADODB.Recordset
    
    vgstrEstado = "NUEVO"
    strSentencia = "SELECT * FROM PVBASEHONORARIOMEDICO WHERE PVBASEHONORARIOMEDICO.INTNUMCARGO = " & vglngCveCargo
    Set rsInfo = frsRegresaRs(strSentencia)
    
    If rsInfo.RecordCount <> 0 Then
        txtProcedimiento.Text = rsInfo!VCHPROCEDIMIENTO
        txtImporteFacturarConvenio.Text = FormatCurrency(rsInfo!NUMIMPORTEAFACTURAR / vgintCantidadCargo)
        txtImporteHonorarioMedico.Text = FormatCurrency(rsInfo!NUMIMPORTEHONORARIO / vgintCantidadCargo)
        lblTotalAFacturar.Caption = FormatCurrency(rsInfo!NUMIMPORTEAFACTURAR)
        lblTotalHonorario.Caption = FormatCurrency(rsInfo!NUMIMPORTEHONORARIO)
        cboMedico.ListIndex = fintLocalizaCbo(cboMedico, rsInfo!intCveMedico)
        vgstrEstado = "EDICION"
    End If
End Sub

