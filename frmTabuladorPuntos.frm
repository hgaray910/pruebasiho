VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTabuladorPuntos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabulador de puntos"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Puntos"
      Height          =   915
      Left            =   3510
      TabIndex        =   13
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtPuntosPaciente 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Cantidad de puntos para el paciente"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtPuntosMedico 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         ToolTipText     =   "Cantidad de puntos para el médico"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Médico"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   525
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Subtotal"
      Height          =   915
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtLimiteInferior 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   120
         MaxLength       =   14
         TabIndex        =   1
         ToolTipText     =   "Límite inferior del subtotal de la cuenta"
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox txtLimiteSuperior 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Límite superior del subtotal de la cuenta"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblLimiteInferior 
         AutoSize        =   -1  'True
         Caption         =   "Límite inferior"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lblLimiteSuperior 
         AutoSize        =   -1  'True
         Caption         =   "Límite superior"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   6900
      MaskColor       =   &H80000014&
      Picture         =   "frmTabuladorPuntos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar"
      Top             =   540
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   4560
      Width           =   1170
      Begin VB.CommandButton cmdDelete 
         Enabled         =   0   'False
         Height          =   480
         Left            =   565
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmTabuladorPuntos.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Borrar el registro"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabar 
         Height          =   480
         Left            =   75
         Picture         =   "frmTabuladorPuntos.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guardar"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7275
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTabulador 
         Height          =   3060
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   5398
         _Version        =   393216
         Cols            =   45
         GridColor       =   12632256
         ScrollBars      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   45
      End
   End
End
Attribute VB_Name = "frmTabuladorPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public llngNumOpcion As Long
Private Enum enmStatus
    stedicion = 1
    stConsulta = 2
    stLimiteInferior = 3
End Enum
Dim stEstado As enmStatus

Dim vgintColumnaLimInferior As Integer            'Para saber cual es la Columna del limite inferior
Dim llngRowActualizar As Long

Private Sub pLlenaGrid()
On Error GoTo NotificaError
    Dim rsTabulador As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    vlstrSentencia = "SELECT * FROM PvTabuladorPuntosLealtad "
    Set rsTabulador = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    rsTabulador.Sort = "mnyLimiteInferior"
    Do Until rsTabulador.EOF
        With grdTabulador
            If Trim(.TextMatrix(1, 1)) = "" Then
                .Row = 1
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        
            .Col = 0
            .TextMatrix(.Row, 1) = Format(rsTabulador!mnyLimiteInferior, "$ ###,###,###,###.00")
            .TextMatrix(.Row, 2) = Format(rsTabulador!mnyLimiteSuperior, "$ ###,###,###,###.00")
            .TextMatrix(.Row, 3) = Format(rsTabulador!intpuntospaciente)
            .TextMatrix(.Row, 4) = Format(rsTabulador!intPuntosMedico)
                        
'            .Sort = flexSortGenericAscending
            rsTabulador.MoveNext
        End With
    Loop
    rsTabulador.Close
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaGrid"))
End Sub

Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
'        For vlbytColumnas = 1 To .Cols - 1
'            .TextMatrix(1, vlbytColumnas) = ""
'        Next vlbytColumnas
'        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
End Sub

Private Function fblnValido() As Boolean
On Error GoTo NotificaError
    Dim llngRow As Long
    
    fblnValido = True
    With grdTabulador
        For llngRow = 1 To .Rows - 1
            If Val(Format(.TextMatrix(llngRow, 1))) = Val(Format(txtLimiteInferior.Text)) And Val(Format(.TextMatrix(llngRow, 2))) = Val(Format(txtLimiteSuperior.Text)) Then
                If Val(Format(.TextMatrix(llngRow, 3))) = Val(Format(txtPuntosPaciente.Text)) And Val(Format(.TextMatrix(llngRow, 4))) = Val(Format(txtPuntosMedico.Text)) Then
                    'Existe información con el mismo contenido.
                    MsgBox SIHOMsg(19), vbOKOnly + vbExclamation, "Mensaje"
                    fblnValido = False
                    Exit For
                Else
                    'Desea actualizar los datos
                    If MsgBox("Ya existe el rango de consumo, ¿Desea actualizar los datos?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        fblnValido = True
                        llngRowActualizar = llngRow
                    Else
                        fblnValido = False
                    End If
                    Exit For
                End If
            End If
        Next llngRow
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValido"))
End Function
Private Sub pCancelar()
    txtLimiteInferior.Text = FormatCurrency(0, 2)
    txtLimiteSuperior.Text = FormatCurrency(0, 2)
    txtPuntosPaciente.Text = 0
    txtPuntosMedico.Text = 0
    pEnfocaTextBox txtLimiteInferior
    cmdDelete.Enabled = False
End Sub


Private Sub pConfiguraGridTabulador()
    With grdTabulador
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedCols = 0
        .FixedRows = 1
        .FormatString = "|Subtotal límite inferior|Subtotal límite superior|Puntos paciente|Puntos médico"
        .ColWidth(0) = 0      'Fix
        .ColWidth(1) = 1800     'Limite inferior
        .ColWidth(2) = 1800     'Límite superior
        .ColWidth(3) = 1500     'Puntos paciente
        .ColWidth(4) = 1500     'Puntos médico
        
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignmentFixed(0) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignRightCenter
        .ColAlignmentFixed(2) = flexAlignRightCenter
        .ColAlignmentFixed(3) = flexAlignRightCenter
        .ColAlignmentFixed(4) = flexAlignRightCenter
        .ScrollBars = flexScrollBarVertical
    End With
End Sub

Private Sub cmdAgregar_Click()
On Error GoTo NotificaError

    If Val(Format(txtLimiteInferior.Text)) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & txtLimiteInferior.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        txtLimiteInferior.SetFocus
        Exit Sub
    End If
    If Val(Format(txtLimiteSuperior.Text)) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & txtLimiteSuperior.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        txtLimiteSuperior.SetFocus
        Exit Sub
    End If
    If Val(txtPuntosPaciente.Text) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & txtPuntosPaciente.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        txtPuntosPaciente.SetFocus
        Exit Sub
    End If
    If Val(txtPuntosMedico.Text) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & txtPuntosMedico.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        txtPuntosMedico.SetFocus
        Exit Sub
    End If
    If Val(Format(txtLimiteInferior.Text)) >= Val(Format(txtLimiteSuperior.Text)) Then
        'El límite superior debe ser mayor que el límite inferior
        MsgBox "El límite superior debe ser mayor que el límite inferior", vbOKOnly + vbExclamation, "Mensaje"
        txtLimiteSuperior.SetFocus
        Exit Sub
    End If

    llngRowActualizar = 0
    If fblnValido Then
        With grdTabulador
            If llngRowActualizar <> 0 Then
                .TextMatrix(llngRowActualizar, 3) = Trim(txtPuntosPaciente.Text)
                .TextMatrix(llngRowActualizar, 4) = Trim(txtPuntosMedico.Text)
            Else
                If Trim(.TextMatrix(1, 1)) = "" Then
                    .Row = 1
                Else
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                End If
            
                .Col = 0
                .TextMatrix(.Row, 1) = Format(txtLimiteInferior.Text, "$ ###,###,###,###.00")
                .TextMatrix(.Row, 2) = Format(txtLimiteSuperior.Text, "$ ###,###,###,###.00")
                .TextMatrix(.Row, 3) = Format(txtPuntosPaciente.Text)
                .TextMatrix(.Row, 4) = Format(txtPuntosMedico.Text)
            
                .Sort = flexSortGenericAscending
            End If
        End With
    End If
    pCancelar
    txtLimiteInferior.SetFocus
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregar_Click"))

End Sub

Private Sub cmdDelete_Click()
    If grdTabulador.Rows - 1 = 1 Then
        pLimpiaGrid grdTabulador
        pConfiguraGridTabulador
        pCancelar
    Else
        grdTabulador.RemoveItem (grdTabulador.Row)
    End If
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo NotificaError
    Dim llngRow As Long
    Dim llngPersonaGraba As Long
    Dim lstrSentencia As String
    Dim rsTabulador As New ADODB.Recordset
    
    
    
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans

        lstrSentencia = "DELETE FROM PvTabuladorPuntosLealtad"
        pEjecutaSentencia lstrSentencia
        
        lstrSentencia = "SELECT * FROM PvTabuladorPuntosLealtad"
        Set rsTabulador = frsRegresaRs(lstrSentencia, adLockOptimistic, adOpenDynamic)
        
        With grdTabulador
            For llngRow = 1 To .Rows - 1
                If Val(Format(grdTabulador.TextMatrix(llngRow, 1), "")) > 0 Then
                    rsTabulador.AddNew
                    rsTabulador!mnyLimiteInferior = Val(Format(.TextMatrix(llngRow, 1), ""))
                    rsTabulador!mnyLimiteSuperior = Val(Format(.TextMatrix(llngRow, 2), ""))
                    rsTabulador!intpuntospaciente = Val(Format(.TextMatrix(llngRow, 3), ""))
                    rsTabulador!intPuntosMedico = Val(Format(.TextMatrix(llngRow, 4), ""))
                    rsTabulador.Update
                End If
            Next llngRow
        End With
        rsTabulador.Close
        
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, llngPersonaGraba, "TABULADOR DE PUNTOS DE LEALTAD", CStr(vgintNumeroDepartamento))
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        pConfiguraGridTabulador
        pLlenaGrid
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        
        txtLimiteInferior.SetFocus
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        pConfiguraGridTabulador
        pLlenaGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabar_Click"))
End Sub

Private Sub Form_Activate()
    pCancelar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    vgintColumnaLimInferior = 1
    
    pLimpiaGrid grdTabulador
    pConfiguraGridTabulador
    pLlenaGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case stEstado
        Case stedicion
            '  ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbExclamation, "Mensaje") = vbYes Then
                Cancel = False
            Else
                Cancel = True
                grdTabulador.SetFocus
            End If
    End Select
End Sub

Private Sub grdTabulador_Click()
    If Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 1))) <> 0 Then
        cmdDelete.Enabled = True
    Else
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub grdTabulador_DblClick()
    If Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 1))) <> 0 Then
        txtLimiteInferior.Text = FormatCurrency(Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 1), "")), 2)
        txtLimiteSuperior.Text = FormatCurrency(Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 2), "")), 2)
        txtPuntosPaciente.Text = Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 3), ""))
        txtPuntosMedico.Text = Val(Format(grdTabulador.TextMatrix(grdTabulador.Row, 4), ""))
    End If
End Sub


Private Sub grdTabulador_GotFocus()
'    If grdTabulador.Col = vgintColumnaLimInferior Then
'        Call pSetCellValueCol(grdTabulador, txtLimiteInferior)
'    End If
End Sub


Private Sub txtLimiteInferior_GotFocus()
    pEnfocaTextBox txtLimiteInferior
End Sub

Private Sub txtLimiteInferior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.ActiveControl.Name = "txtLimiteInferior" Then
            txtLimiteInferior.Text = FormatCurrency(Val(Format(txtLimiteInferior.Text, "")), 2)
            pEnfocaTextBox txtLimiteInferior
        End If
    Else
        If Not fblnFormatoCantidad(txtLimiteInferior, KeyAscii, 2) Then
            KeyAscii = 7
        End If
    End If
End Sub


Private Sub txtLimiteInferior_LostFocus()
    txtLimiteInferior.Text = FormatCurrency(Val(Format(txtLimiteInferior.Text, "")), 2)
End Sub

Private Sub txtLimiteSuperior_GotFocus()
    pEnfocaTextBox txtLimiteSuperior
End Sub

Private Sub txtLimiteSuperior_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.ActiveControl.Name = "txtLimiteSuperior" Then
            txtLimiteSuperior.Text = FormatCurrency(Val(Format(txtLimiteSuperior.Text, "")), 2)
            pEnfocaTextBox txtLimiteSuperior
        End If
    Else
        If Not fblnFormatoCantidad(txtLimiteSuperior, KeyAscii, 2) Then
            KeyAscii = 7
        End If
    End If
End Sub

Private Sub txtLimiteSuperior_LostFocus()
    txtLimiteSuperior.Text = FormatCurrency(Val(Format(txtLimiteSuperior.Text, "")), 2)
End Sub


Private Sub txtPuntosMedico_GotFocus()
    pEnfocaTextBox txtPuntosMedico
End Sub

Private Sub txtPuntosMedico_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub


Private Sub txtPuntosPaciente_GotFocus()
    pEnfocaTextBox txtPuntosPaciente
End Sub

Private Sub txtPuntosPaciente_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub


