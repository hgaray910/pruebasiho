VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAddendaDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos para la addenda"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDescripcion 
      Caption         =   "Descripción"
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Descripción del campo"
      Top             =   3600
      Width           =   9855
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "Descripción del campo"
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.TextBox txtValor 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdGrabar 
      Enabled         =   0   'False
      Height          =   465
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAddendaDatos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Guarda los datos de la addenda"
      Top             =   5160
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAddendaEstructura 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Datos de la addenda"
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5741
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAddendaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAddendaEstructura As New ADODB.Recordset
Public lngAddenda As Long
Public lngCuenta As Long
Public strTipoIngreso As String
Public lngCveEmpresaPaciente As Long
Public lngCveEmpresaContable As Long
Public blngrupocuentas As Integer
Dim vgblnInicio As Boolean

Private Sub cmdGrabar_Click()

Dim strParametrosSP As String
Dim vllngPersonaGraba As Long

On Error GoTo NotificaError
                                
grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4) = txtValor.Text
                                
With grdAddendaEstructura

        If fblnDatosValidos = True Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba <> 0 Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                
                    For i = 1 To .Rows - 1
                        'Graba el detalle de ADDENDADATOS
                        If blngrupocuentas = 1 Then
                            strParametrosSP = CStr(lngCuenta) & "|" & "G" & "|" & Trim(.TextMatrix(i, 1)) & "|" & Trim(.TextMatrix(i, 4))
                            frsEjecuta_SP strParametrosSP, "SP_PVUPDADDENDADATOS", False
                        Else
                            strParametrosSP = CStr(lngCuenta) & "|" & strTipoIngreso & "|" & Trim(.TextMatrix(i, 1)) & "|" & Trim(.TextMatrix(i, 4))
                            frsEjecuta_SP strParametrosSP, "SP_PVUPDADDENDADATOS", False
                        End If
                    Next i
                    
                EntornoSIHO.ConeccionSIHO.CommitTrans
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "CONFIGURACIÓN DE ADDENDA", CStr(lngCuenta) & " " & strTipoIngreso & " - Addenda " & CStr(lngAddenda))
                
                MsgBox SIHOMsg(420), vbInformation, "Mensaje"
                pInicia
            End If
        End If
        
End With
           
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabar_Click"))
End Sub

Private Sub cmdGrabar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

If KeyCode = 27 Then     'ESC
     Unload Me
Else
    If grdAddendaEstructura.Col = 4 Then 'Si la columna a modificar es la columna  "Valor"
        Call pEditarColumna(32, txtValor, grdAddendaEstructura)
    Else
        txtValor.Visible = False
    End If
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabar_KeyDown"))

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
'    If KeyAscii = 27 Then
'        If cmdGrabar.Enabled = True Then
'            '¿Desea abandonar la operación?
'            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
'                pInicia
'            End If
'        Else
'            Unload Me
'        End If
'    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
On Error GoTo NotificaError

    vgstrNombreForm = Me.Name
    Me.Icon = frmMenuPrincipal.Icon

    pInicia
    
    vgblnInicio = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pInicia()
    On Error GoTo NotificaError

    grdAddendaEstructura.Redraw = False
    pLimpiaGrid grdAddendaEstructura
    If blngrupocuentas = 0 Then
    'Primero intenta de cargar el grid con la información previamente cargada
        If lngAddenda <> 0 And lngCuenta <> 0 And strTipoIngreso <> "" Then
            Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "|" & CStr(lngCuenta) & "|" & strTipoIngreso & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
            If rsAddendaEstructura.RecordCount < 1 Then
                Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "||" & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
            End If
    '    End If
    '
    '    'Si no encuentra información, carga la información por default
    '    If rsAddendaEstructura.RecordCount < 1 Then
        Else
            Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "||" & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
        End If
    Else
        If lngAddenda <> 0 And lngCuenta <> 0 And strTipoIngreso <> "" Then
        Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "|" & CStr(lngCuenta) & "|" & "G" & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
        If rsAddendaEstructura.RecordCount < 1 Then
            Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "||" & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
        End If
'    End If
'
'    'Si no encuentra información, carga la información por default
'    If rsAddendaEstructura.RecordCount < 1 Then
        Else
            Set rsAddendaEstructura = frsEjecuta_SP(CStr(lngAddenda) & "||" & "|" & lngCveEmpresaContable & "|" & lngCveEmpresaPaciente, "SP_PVSELADDENDADATOSPANTALLA")
        End If
    End If
    txtValor.Visible = False
    
    pLlenarMshFGrdRs grdAddendaEstructura, rsAddendaEstructura, 0
    pConfiguraGridAddenda
    
    cmdGrabar.Enabled = False
    
    grdAddendaEstructura.Col = 4
    grdAddendaEstructura.Row = 1
        
    txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))
    txtValor.Text = grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4)
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pInicia"))
End Sub
Sub pLimpiaGrid(ObjGrd As MSHFlexGrid)
    On Error GoTo NotificaError

    Dim vlbytColumnas As Byte
    
    With ObjGrd
        .FormatString = ""
        .Row = 1
        .Col = 1
        .Rows = 2
        For vlbytColumnas = 1 To .Cols - 1
            .TextMatrix(1, vlbytColumnas) = ""
        Next vlbytColumnas
        .TextMatrix(1, 1) = ""
        .RowData(1) = -1
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpiaGrid"))
End Sub
Private Sub pConfiguraGridAddenda()
        On Error GoTo NotificaError

    With grdAddendaEstructura
        .Cols = 9
'        .FixedCols = 3
        .FixedRows = 1
        .FormatString = "|Clave|^Dato|Tipo de dato|^Valor|^Obligatorio|Descripción del campo|Longitud|LongMin"
        .ColWidth(0) = 0  'Default fixed
        .ColWidth(1) = 0  'Clave
        .ColWidth(2) = 2000 'Dato
        .ColWidth(3) = 0 'Tipo de dato
        .ColWidth(4) = 6600  'Valor
        .ColWidth(5) = 900 'Obligatorio
        .ColWidth(6) = 0 'Descripción del campo
        .ColWidth(7) = 0 'Longitud del campo
        .ColWidth(8) = 0 'Longitud mínima del campo
                
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        
        If grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "STRING" Then
            .ColAlignment(4) = flexAlignLeftCenter
        ElseIf grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "INTEGER" Then
            .ColAlignment(4) = flexAlignRightCenter
        ElseIf grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "DECIMAL" Then
            .ColAlignment(4) = flexAlignRightCenter
        Else
            .ColAlignment(4) = flexAlignLeftCenter
        End If
        
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter

        .ScrollBars = flexScrollBarBoth
            
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConfiguraGridAddenda"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdGrabar.Enabled = True Then
        Cancel = True
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            pInicia
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub grdAddendaEstructura_Click()
    On Error GoTo NotificaError

If grdAddendaEstructura.Col = 4 Then 'Si la columna a modificar es la columna  "Valor"
    Call pEditarColumna(32, txtValor, grdAddendaEstructura)
Else
    txtValor.Visible = False
End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdAddendaEstructura_Click"))
End Sub
Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid, Optional pintText As Integer = 0)
    On Error GoTo NotificaError

    Dim vlintTexto As Integer
    
    txtEdit.Text = grid.TextMatrix(grid.Row, 4)
    txtEdit.Visible = True
    
    If grid.CellWidth < 0 Then Exit Sub
    txtEdit.Move grid.Left + grid.CellLeft, grid.Top + grid.CellTop, grid.CellWidth - 8, grid.CellHeight - 8
    
    txtEdit.Enabled = True
    
    txtEdit.SetFocus
    
    'Se selecciona el contenido del textBox
    pSelTextBox txtEdit

    'Carga la información de la documentación en el txtDescripcion
    txtDescripcion.Text = grid.TextMatrix(grid.Row, 6)

    cmdGrabar.Enabled = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna"))
End Sub

Private Sub grdAddendaEstructura_EnterCell()
    txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))
    txtValor.Text = grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4)
End Sub

Private Sub grdAddendaEstructura_GotFocus()
    On Error GoTo NotificaError

    Dim vllngCantidadCargosPaquete As Long
    Dim intBitValidaPaquetes As Long
    
    '----------------------------
    'Si esta vacío que se salga
    '----------------------------
    If (grdAddendaEstructura.Row = 0) Or (grdAddendaEstructura.RowData(1) = -1) Then Exit Sub
    '----------------------------
        
    'Copia el valor del textbox al grid y lo esconde
    If grdAddendaEstructura.Col = 4 Then
        Call pSetCellValueCol(grdAddendaEstructura, txtValor)
        txtValor.Visible = False
        
'        txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))
'        txtValor.Text = grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4)

    End If
    
    'Si acaba de iniciar la pantalla mandar un ENTER para habilitar la edición
    If vgblnInicio = True Then
        Call grdAddendaEstructura_KeyDown(13, 0)
        vgblnInicio = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdAddendaEstructura_GotFocus"))
End Sub

Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox)
    On Error GoTo NotificaError

    Dim lngRenglon As Long
    Dim lngColumna As Long
    
    If grid.MouseCol = 4 Then
        grid.Col = 4
    ElseIf grid.MouseCol = 32 Then
        grid.Col = grid.MouseCol
    End If
    If txtEdit.Visible Then
        If txtEdit.Text <> "" Then
            lngRenglon = grid.Row
            lngColumna = grid.Col

            grid.Row = lngRenglon
            grid.Col = lngColumna
            txtEdit.Visible = False
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol"))
End Sub

Private Sub grdAddendaEstructura_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError

If KeyCode = 27 Then     'ESC
     Unload Me
Else
    If grdAddendaEstructura.Col = 4 And (KeyCode <> 13 Or KeyCode <> 38 Or KeyCode <> 40) Then  'Si la columna a modificar es la columna  "Valor"
        Call pEditarColumna(32, txtValor, grdAddendaEstructura)
    Else
        txtValor.Visible = False
    End If
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdAddendaEstructura_KeyDown"))
End Sub

Private Sub grdAddendaEstructura_LeaveCell()
    txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))
    grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4) = txtValor.Text
End Sub

Private Sub txtValor_GotFocus()

    'Se valida la longitud del textBox con el campo a capturar
'    txtValor.Text = grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 4)
'    txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))

End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    'Para verificar que tecla fue presionada en el textbox
    With grdAddendaEstructura
        Select Case KeyCode
            Case 27   'ESC
                 txtValor.Visible = False
                .SetFocus
'                txtValor.Text = ""
            Case 38   'Flecha para arriba
                .SetFocus
                .TextMatrix(.Row, 4) = txtValor.Text
                DoEvents
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                    'Se valida la longitud del textBox con el campo a capturar
                    txtValor.MaxLength = Val(.TextMatrix(.Row, 7))
                    Call pEditarColumna(32, txtValor, grdAddendaEstructura)
                End If
            Case 40, 13  'Flechas y ENTER
                .SetFocus
                .TextMatrix(.Row, 4) = txtValor.Text
                DoEvents
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    'Se valida la longitud del textBox con el campo a capturar
                    txtValor.MaxLength = Val(.TextMatrix(.Row, 7))
                    Call pEditarColumna(32, txtValor, grdAddendaEstructura)
                Else
                    cmdGrabar.SetFocus
                End If
        End Select
    End With
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtValor_KeyDown"))
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

'Se valida la longitud del textBox con el campo a capturar
txtValor.MaxLength = Val(grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 7))

'Se formatea el textBox según el tipo de campo que se va a editar
If grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "STRING" Then
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
ElseIf grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "DECIMAL" Then
    If Not fblnFormatoCantidad(txtValor, KeyAscii, 2) Then KeyAscii = 7
ElseIf grdAddendaEstructura.TextMatrix(grdAddendaEstructura.Row, 3) = "INTEGER" Then
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtValor_KeyPress"))
End Sub

Private Sub txtValor_LostFocus()
    On Error GoTo NotificaError

    Call grdAddendaEstructura_GotFocus
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtValor_LostFocus"))
End Sub

Private Function fblnDatosValidos() As Boolean

    fblnDatosValidos = True

    For i = 1 To grdAddendaEstructura.Rows - 1
        If Trim(grdAddendaEstructura.TextMatrix(i, 4)) = "" And Trim(grdAddendaEstructura.TextMatrix(i, 5)) = "*" Then
            fblnDatosValidos = False
            'Hay datos obligatorios sin definirse
            MsgBox SIHOMsg(1142) & grdAddendaEstructura.TextMatrix(i, 2), vbExclamation, "Mensaje"
            grdAddendaEstructura.Row = i
            grdAddendaEstructura.Col = 4
            Call grdAddendaEstructura_KeyDown(13, 0)
            Exit For
        End If
        
        If Len(Trim(grdAddendaEstructura.TextMatrix(i, 4))) < Val(grdAddendaEstructura.TextMatrix(i, 8)) Then
            fblnDatosValidos = False
            'Este dato no cumple con la longitud requerida
            MsgBox SIHOMsg(1141) & grdAddendaEstructura.TextMatrix(i, 2), vbExclamation, "Mensaje"
            grdAddendaEstructura.Row = i
            grdAddendaEstructura.Col = 4
            Call grdAddendaEstructura_KeyDown(13, 0)
            Exit For
        End If
        
    Next i
    
End Function
