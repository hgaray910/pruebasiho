VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAsignacionPaquetesPaciente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de paquetes"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   1170
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   4695
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   10695
      Begin VB.CommandButton cmdHonorarios 
         Caption         =   "Consultar honorarios médicos"
         Height          =   375
         Left            =   8280
         TabIndex        =   26
         ToolTipText     =   "Consulta honorarios médicos del paquete"
         Top             =   150
         Width           =   2295
      End
      Begin VSFlex7LCtl.VSFlexGrid grdDetallePaquete 
         Height          =   1815
         Left            =   5880
         TabIndex        =   7
         ToolTipText     =   "Porcentaje del precio"
         Top             =   2520
         Width           =   4695
         _cx             =   8281
         _cy             =   3201
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAsignacionPaquetesPaciente.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.ListBox lstPaquetesDisponibles 
         Height          =   3765
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   4710
      End
      Begin VB.ListBox lstPaquetesAsignados 
         Height          =   1815
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   4710
      End
      Begin VB.CommandButton cmdAgrega 
         Caption         =   ">"
         Height          =   495
         Left            =   5085
         TabIndex        =   4
         ToolTipText     =   "Agregar"
         Top             =   1950
         Width           =   495
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "<"
         Height          =   495
         Left            =   5085
         TabIndex        =   5
         ToolTipText     =   "Eliminar"
         Top             =   2460
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Paquetes disponibles"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Paquetes asignados al paciente"
         Height          =   195
         Left            =   5880
         TabIndex        =   24
         Top             =   255
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del paciente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   10695
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   7560
         TabIndex        =   22
         Top             =   0
         Width           =   25
      End
      Begin VB.TextBox txtCuarto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         TabIndex        =   20
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtNoAfiliacion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtArea 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtEmpresaTipoPaciente 
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   4455
      End
      Begin VB.Frame Frame3 
         Height          =   1560
         Left            =   2730
         TabIndex        =   12
         Top             =   0
         Width           =   25
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Externo"
         Height          =   255
         Index           =   1
         Left            =   1185
         TabIndex        =   2
         Top             =   510
         Width           =   975
      End
      Begin VB.OptionButton optTipoPaciente 
         Caption         =   "Interno"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   510
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtCvePaciente 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1545
         TabIndex        =   0
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   25
         Left            =   0
         TabIndex        =   10
         Top             =   915
         Width           =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "Cuarto"
         Height          =   255
         Left            =   7800
         TabIndex        =   21
         Top             =   1120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "No. afiliación"
         Height          =   255
         Left            =   7800
         TabIndex        =   19
         Top             =   400
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Area"
         Height          =   255
         Left            =   7800
         TabIndex        =   17
         Top             =   760
         Width           =   975
      End
      Begin VB.Label lblEmpresaTipoPaciente 
         Caption         =   "Procedencia"
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Número de cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1125
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmAsignacionPaquetesPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vgstrTipoPaciente As String
Dim vgintCvePaciente As Long
Dim vgintTipoPaciente As Long
Dim vgintCveEmpresa As Long
Dim alstrParametrosSalida() As String
Dim intbitpesos As Long
Dim vlblnAceptaCambios As Boolean

Private Sub cmdAgrega_Click()
'AQUI
    pGrabaAsignacion
End Sub

Private Sub cmdElimina_Click()
    Dim strParametrosSP As String
    Dim rsCuenta As New ADODB.Recordset
    
    If vlblnAceptaCambios Then
        '-----------------------------------------------------
        '|  Valida que la cuenta no esté cerrada
        '-----------------------------------------------------
        strParametrosSP = txtCvePaciente.Text & _
                          "|" & "0" & _
                          "|" & IIf(optTipoPaciente(0).Value, "I", "E") & _
                          "|" & vgintClaveEmpresaContable
        Set rsCuenta = frsEjecuta_SP(strParametrosSP, "sp_PvSelDatosPaciente")
        If Not rsCuenta.EOF Then
            If rsCuenta!CuentaCerrada = 1 Or rsCuenta!CuentaCerrada = True Then
                '|  La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
                MsgBox SIHOMsg(596), vbCritical, "Mensaje"
                Exit Sub
            End If
        End If
        pEliminaPaquete True
    End If
End Sub

Private Sub cmdHonorarios_Click()
    Dim rsHonorarios As ADODB.Recordset
    Dim strSentencia As String
    If lstPaquetesAsignados.ListIndex = -1 Then
        'No hay paquete seleccionado
    Else
        strSentencia = "SELECT count(*) numero From PVPAQUETEHONORARIOS  WHERE PVPAQUETEHONORARIOS.INTCVEPAQUETE =" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        Set rsHonorarios = frsRegresaRs(strSentencia)
        If rsHonorarios!numero > 0 Then
            frmHonorariosCirugia.lngNumPaquete = lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            frmHonorariosCirugia.blnSelecciona = False
            frmHonorariosCirugia.Show vbModal
        Else
            MsgBox SIHOMsg(13), vbExclamation, "Mensaje"
        End If
        
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl.Name <> "grdDetallePaquete" Then
        If KeyCode = vbKeyReturn Then SendKeys vbTab
        If KeyCode = vbKeyEscape Then Unload Me
    End If
End Sub

Private Sub Form_Load()
      
    Me.Icon = frmMenuPrincipal.Icon
  
    txtNombre.Locked = True
    txtEmpresaTipoPaciente.Locked = True
  
    If fintEsInterno(vglngNumeroLogin, enmTipoProceso.AsignarPaquetes) > 0 Then
      If fintEsInterno(vglngNumeroLogin, enmTipoProceso.AsignarPaquetes) = 1 Then
        optTipoPaciente(0).Value = True
      Else
        optTipoPaciente(1).Value = True
      End If
    End If

End Sub

Private Sub grdDetallePaquete_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Not IsNumeric(grdDetallePaquete.TextMatrix(Row, Col)) Then
        grdDetallePaquete.TextMatrix(Row, Col) = "100"
    Else
        If CLng(grdDetallePaquete.TextMatrix(Row, Col)) = 0 Or CLng(grdDetallePaquete.TextMatrix(Row, Col)) > 100 Then
            grdDetallePaquete.TextMatrix(Row, Col) = "100"
        Else
            grdDetallePaquete.TextMatrix(Row, Col) = CStr(CLng(grdDetallePaquete.TextMatrix(Row, Col)))
        End If
    End If
    pActualizaPorcentaje CInt(Row), CInt(grdDetallePaquete.TextMatrix(Row, Col))
End Sub

Private Sub grdDetallePaquete_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strParametrosSP As String
    Dim rsCuenta As ADODB.Recordset
    Dim intCantidad As Long
    Dim intPaqueteSinMarcar As Long
    Dim strParametrosPaquetes As String

    If Col = 2 Then
        If vlblnAceptaCambios Then
            strParametrosSP = vgintCvePaciente & "|" & IIf(optTipoPaciente(0).Value, "I", "E")
            Set rsCuenta = frsEjecuta_SP(strParametrosSP, "Sp_EXSelPacienteIngreso")
            If rsCuenta!intcuentacerrada = 1 Or rsCuenta!intcuentacerrada = True Then
                '|  La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
                MsgBox SIHOMsg(596), vbCritical, "Mensaje"
                Cancel = True
            End If
            '------------------------------------------------
            '|  Valida que el paquete no esté facturado
            '------------------------------------------------
            ' Función que regresa paquetes sin facturar
            intCantidad = 1
            strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            frsEjecuta_SP strParametrosPaquetes, "FN_PVSELPAQUETESINFACTURAR", True, intCantidad
        
            If grdDetallePaquete.Rows - 1 - intCantidad >= Row Then
                '|  El paquete está cobrado, no se puede eliminar.
                MsgBox "El paquete está facturado", vbCritical, "Mensaje"
                Cancel = True
            End If
            
            'Valida que el paquete no se haya marcado para facturar al paciente
            intPaqueteSinMarcar = 1
            vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            frsEjecuta_SP strParametrosPaquetes, "Sp_PvSelPaqueteSinMarcar", True, intPaqueteSinMarcar
            If intPaqueteSinMarcar = 0 Then
                'No se puede borrar el cargo, ha sido marcado para facturar
                MsgBox "El paquete está marcado para facturar", vbCritical, "Mensaje"
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub grdDetallePaquete_DblClick()
    grdDetallePaquete.EditCell
End Sub

Private Sub grdDetallePaquete_GotFocus()
    If grdDetallePaquete.Row < 1 And grdDetallePaquete.Rows > 1 Then
        grdDetallePaquete.Row = 1
        grdDetallePaquete.Col = 2
    End If
End Sub

Private Sub grdDetallePaquete_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105) Then
        grdDetallePaquete.EditCell
    End If
    If KeyCode = vbKeyReturn Then
        If grdDetallePaquete.Row < grdDetallePaquete.Rows - 1 Then
            grdDetallePaquete.Row = grdDetallePaquete.Row + 1
        End If
    End If
End Sub

Private Sub grdDetallePaquete_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub grdDetallePaquete_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 27 Then Exit Sub
    If Col = 2 Then
        If KeyAscii > 57 Or KeyAscii < 48 Or (Len(grdDetallePaquete.EditText) = 3 And grdDetallePaquete.EditSelText = "") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub grdDetallePaquete_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Row < grdDetallePaquete.Rows - 1 Then
            grdDetallePaquete.Row = Row + 1
        End If
    End If
End Sub

Private Sub lstPaquetesAsignados_Click()
    pMostrarPaquete
End Sub

Private Sub lstPaquetesAsignados_DblClick()
    Dim vlintIndex As Integer
    Dim vlIntCont As Integer
    Dim vlstrSentencia As String
    Dim strParametrosSP As String
    Dim rsCuenta As New ADODB.Recordset
    
    '-----------------------------------------------------
    '|  Valida que la cuenta no esté cerrada
    '-----------------------------------------------------
    If vlblnAceptaCambios Then
        strParametrosSP = txtCvePaciente.Text & _
                          "|" & "0" & _
                          "|" & IIf(optTipoPaciente(0).Value, "I", "E") & _
                          "|" & vgintClaveEmpresaContable
        Set rsCuenta = frsEjecuta_SP(strParametrosSP, "sp_PvSelDatosPaciente")
        If rsCuenta!CuentaCerrada = 1 Or rsCuenta!CuentaCerrada = True Then
            '|  La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
            MsgBox SIHOMsg(596), vbCritical, "Mensaje"
            Exit Sub
        End If
            
        vlintIndex = lstPaquetesAsignados.ListIndex
        For vlIntCont = 0 To lstPaquetesAsignados.ListCount - 1
            If Mid(lstPaquetesAsignados.List(vlIntCont), 1, 1) = "*" Then
                lstPaquetesAsignados.List(vlIntCont) = Mid(lstPaquetesAsignados.List(vlIntCont), 3, Len(lstPaquetesAsignados.List(vlIntCont)))
            End If
        Next vlIntCont
        lstPaquetesAsignados.List(vlintIndex) = "* " & lstPaquetesAsignados.List(vlintIndex)
        lstPaquetesAsignados.ListIndex = vlintIndex
        vlstrSentencia = ""
        vlstrSentencia = " Update PvPaquetePaciente "
        vlstrSentencia = vlstrSentencia & " Set PvPaquetePaciente.BITPAQUETEDEFAULT = 0 "
        vlstrSentencia = vlstrSentencia & " WHERE PvPaquetePaciente.INTMOVPACIENTE = " & txtCvePaciente.Text & " AND "
        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "'"
        pEjecutaSentencia vlstrSentencia
        vlstrSentencia = ""
        vlstrSentencia = " Update PvPaquetePaciente "
        vlstrSentencia = vlstrSentencia & " Set PvPaquetePaciente.BITPAQUETEDEFAULT = 1 "
        vlstrSentencia = vlstrSentencia & " WHERE PvPaquetePaciente.INTMOVPACIENTE = " & txtCvePaciente.Text & " AND "
        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND "
        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        pEjecutaSentencia vlstrSentencia
        
        'actualiza paquete en el expediente del paciente
         vlstrSentencia = "update expacienteingreso set INTCVEPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & _
            " where intnumcuenta =" & vgintCvePaciente & "                 "
        
        pEjecutaSentencia vlstrSentencia
    End If
End Sub

Private Sub lstPaquetesDisponibles_DblClick()
    cmdAgrega_Click
End Sub

Private Sub OptTipoPaciente_Click(Index As Integer)
    lstPaquetesDisponibles.Clear
    lstPaquetesAsignados.Clear
    
    vlblnAceptaCambios = False
    
    txtNombre.Text = ""
    txtEmpresaTipoPaciente.Text = ""
    vgstrTipoPaciente = ""
    txtNoAfiliacion.Text = ""
    txtArea.Text = ""
    txtCuarto.Text = ""
    vgintCvePaciente = 0
    vgintTipoPaciente = 0
    vgintCveEmpresa = 0
    If txtCvePaciente.Enabled Or txtCvePaciente.Visible Then
       pEnfocaTextBox txtCvePaciente
    End If
End Sub

Private Sub txtCvePaciente_Change()
    lstPaquetesDisponibles.Clear
    lstPaquetesAsignados.Clear
    
    vlblnAceptaCambios = False
    
    txtNombre.Text = ""
    txtEmpresaTipoPaciente.Text = ""
    vgstrTipoPaciente = ""
    txtNoAfiliacion.Text = ""
    txtArea.Text = ""
    txtCuarto.Text = ""
    vgintCvePaciente = 0
    vgintTipoPaciente = 0
    vgintCveEmpresa = 0
    grdDetallePaquete.Rows = 1
End Sub

Private Sub txtCvePaciente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlrsDatos As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vldblValidoDolares As String
    Dim strParametros As String
    Dim vllngCuentaEnGrupo As Long
    
    If KeyCode = vbKeyReturn Then
        If RTrim(txtCvePaciente.Text) = "" Then
            With FrmBusquedaPacientes
                .vgstrTipoPaciente = IIf(optTipoPaciente(1).Value, "E", "I")
                .Caption = .Caption & IIf(optTipoPaciente(1).Value, " externos", " internos")
                .vgblnPideClave = False
                .vgIntMaxRecords = 100
                .vgstrMovCve = "M"
                .optSoloActivos.Enabled = True
                .optSinFacturar.Enabled = True
                .optTodos.Enabled = True
                .optSinFacturar.Value = True
                
                If optTipoPaciente(1).Value Then
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha"", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                Else
                    .vgStrOtrosCampos = ", ExPacienteIngreso.dtmFechaHoraIngreso as ""Fecha ing."", ExPacienteIngreso.dtmFechaHoraEgreso as ""Fecha egr."", isnull(CCempresa.vchDescripcion,adTipoPaciente.vchDescripcion) as Empresa "
                End If
                
                .vgstrTamanoCampo = IIf(optTipoPaciente(1).Value, "800,3400,1700,4100", "800,3400,990,990,4100")
                
                txtCvePaciente.Text = .flngRegresaPaciente()
                
                If txtCvePaciente <> -1 Then
                    txtCvePaciente_KeyDown vbKeyReturn, 0
                Else
                    txtCvePaciente.Text = ""
                End If
            End With
        Else
            If optTipoPaciente(0).Value Then 'Internos
                vlstrSentencia = " SELECT rtrim(AdPaciente.vchApellidoPaterno)||' '||rtrim(AdPaciente.vchApellidoMaterno)||' '||rtrim(AdPaciente.vchNombre) as Nombre, "
                vlstrSentencia = vlstrSentencia & " AdAdmision.intCveEmpresa cveEmpresa, ccEmpresa.vchDescripcion as Empresa, "
                vlstrSentencia = vlstrSentencia & " AdAdmision.tnyCveTipoPaciente cveTipoPaciente, AdTipoPaciente.vchDescripcion as Tipo,  "
                vlstrSentencia = vlstrSentencia & "     Adadmision.vchNumAfiliacion as NumAfiliacion, "
                vlstrSentencia = vlstrSentencia & " AdTipoPaciente.bitUtilizaConvenio as bitUtilizaConvenio, "
                vlstrSentencia = vlstrSentencia & " TRIM(ADPACIENTE.vchCallePart)||' '||TRIM(ADPACIENTE.VCHNUMEROEXTERIOR)||CASE WHEN ADPACIENTE.VCHNUMEROINTERIOR IS NULL THEN '' ELSE ' Int. '|| Trim(AdPaciente.VchNumeroInterior) END as Direccion, AdPaciente.vchColoniaPart as Colonia, "
                vlstrSentencia = vlstrSentencia & " AdPaciente.vchTelefonoPart as Telefono, AdAdmision.dtmFechaIngreso as FechaIngreso, "
                vlstrSentencia = vlstrSentencia & " AdAdmision.intCveExtra, chrRFC as RFC, "
                vlstrSentencia = vlstrSentencia & " AdAdmision.vchNumCuarto Cuarto, "
                vlstrSentencia = vlstrSentencia & " isnull(ccTipoConvenio.bitAseguradora,0) "
                vlstrSentencia = vlstrSentencia & " bitAseguradora,"
                vlstrSentencia = vlstrSentencia & " AdArea.VCHDESCRIPCION Area, "
                vlstrSentencia = vlstrSentencia & " AdAdmision.bitFacturado Facturado "
                vlstrSentencia = vlstrSentencia & " FROM AdAdmision INNER JOIN AdPaciente ON AdAdmision.numCvePaciente = AdPaciente.numCvePaciente "
                vlstrSentencia = vlstrSentencia & "   INNER JOIN AdTipoPaciente ON AdAdmision.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente "
                vlstrSentencia = vlstrSentencia & "   INNER JOIN Nodepartamento ON AdAdmision.intcvedepartamento = nodepartamento.smicvedepartamento "
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER Join CcEmpresa ON AdAdmision.intCveEmpresa = CcEmpresa.intCveEmpresa "
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER Join CcTipoConvenio ON ccEmpresa.tnyCveTipoConvenio = ccTipoConvenio.tnyCveTipoConvenio "
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER JOIN AdCuarto ON (AdAdmision.VCHNUMCUARTO = AdCuarto.VCHNUMCUARTO)"
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER JOIN AdArea ON (AdCuarto.TNYCVEAREA = AdArea.TNYCVEAREA)"
                vlstrSentencia = vlstrSentencia & " Where AdAdmision.numNumCuenta = " & txtCvePaciente.Text & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            Else  'Externos
                vlstrSentencia = " SELECT rtrim(chrApePaterno)||' '||rtrim(chrApeMaterno)||' '||rtrim(chrNombre) as Nombre, "
                vlstrSentencia = vlstrSentencia & "     RegistroExterno.intClaveEmpresa cveEmpresa, "
                vlstrSentencia = vlstrSentencia & "     ccEmpresa.vchDescripcion as Empresa, "
                vlstrSentencia = vlstrSentencia & "     RegistroExterno.tnyCveTipoPaciente cveTipoPaciente, "
                vlstrSentencia = vlstrSentencia & "     RegistroExterno.vchNumAfiliacion as NumAfiliacion, "
                vlstrSentencia = vlstrSentencia & "     AdTipoPaciente.vchDescripcion as Tipo, "
                vlstrSentencia = vlstrSentencia & "     Case "
                vlstrSentencia = vlstrSentencia & "          when RegistroExterno.intClaveEmpresa = 0 then  0 "
                vlstrSentencia = vlstrSentencia & "          else 1 "
                vlstrSentencia = vlstrSentencia & "         end as bitUtilizaConvenio, "
                vlstrSentencia = vlstrSentencia & " TRIM(EXTERNO.CHRCALLE) || ' ' || TRIM(EXTERNO.VCHNUMEROEXTERIOR)||CASE WHEN EXTERNO.VCHNUMEROINTERIOR IS NULL THEN '' ELSE ' Int. '|| Trim(EXTERNO.VchNumeroInterior) END as Direccion, ' ' as Colonia, "
                vlstrSentencia = vlstrSentencia & " Externo.chrTelefono as Telefono, ' ' as FechaIngreso, "
                vlstrSentencia = vlstrSentencia & " ' ' as Medico, "
                vlstrSentencia = vlstrSentencia & "     ' ' as Area,"
                vlstrSentencia = vlstrSentencia & " RegistroExterno.intCveExtra, chrRFC as RFC, "
                vlstrSentencia = vlstrSentencia & " ' ' as Diagnostico, '' as Cuarto, isnull(ccTipoConvenio.bitAseguradora,0) bitAseguradora,RegistroExterno.bitFacturado Facturado "
                vlstrSentencia = vlstrSentencia & " FROM RegistroExterno INNER JOIN Externo ON RegistroExterno.intNumPaciente = Externo.intNumPaciente "
                vlstrSentencia = vlstrSentencia & "   INNER JOIN AdTipoPaciente ON RegistroExterno.tnyCveTipoPaciente = AdTipoPaciente.tnyCveTipoPaciente "
                vlstrSentencia = vlstrSentencia & "   INNER JOIN nodepartamento ON RegistroExterno.intcvedepartamento = nodepartamento.smicvedepartamento "
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER Join CcEmpresa ON RegistroExterno.intClaveEmpresa = CcEmpresa.intCveEmpresa "
                vlstrSentencia = vlstrSentencia & "   LEFT OUTER Join CcTipoConvenio ON ccEmpresa.tnyCveTipoConvenio = ccTipoConvenio.tnyCveTipoConvenio "
                vlstrSentencia = vlstrSentencia & " Where RegistroExterno.intNumCuenta = " & txtCvePaciente.Text & " and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable
            End If
            Set vlrsDatos = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
            
            If vlrsDatos.RecordCount <> 0 Then
                    vlblnAceptaCambios = True
            
                    '-------------------------------
                    'Datos generales del Paciente
                    '-------------------------------
                    txtNombre.Text = vlrsDatos!Nombre
                    txtEmpresaTipoPaciente.Text = IIf(vlrsDatos!bitUtilizaConvenio = 1, IIf(IsNull(vlrsDatos!empresa), "", vlrsDatos!empresa), vlrsDatos!tipo)
                    vgstrTipoPaciente = IIf(optTipoPaciente(0).Value, "I", "E")
                    txtNoAfiliacion.Text = IIf(IsNull(vlrsDatos!NumAfiliacion), "", vlrsDatos!NumAfiliacion)
                    txtArea.Text = IIf(IsNull(vlrsDatos!Area), "", vlrsDatos!Area)
                    txtCuarto.Text = IIf(IsNull(vlrsDatos!Cuarto), "", vlrsDatos!Cuarto)
                    vgintCvePaciente = txtCvePaciente.Text
                    vgintTipoPaciente = vlrsDatos!cveTipoPaciente
                    vgintCveEmpresa = IIf(IsNull(vlrsDatos!cveEmpresa), 0, vlrsDatos!cveEmpresa)
                    pCargaPaquetes
                    
                    lstPaquetesDisponibles.SetFocus
                    
                    vllngCuentaEnGrupo = 1
                    strParametros = vgintCvePaciente & "|" & vgstrTipoPaciente
                    frsEjecuta_SP strParametros, "FN_PVSELCUENTAENGRUPO", True, vllngCuentaEnGrupo
                    If vllngCuentaEnGrupo <> 0 Then
                        'No se puede modificar la cuenta, se encuentra incluida en un grupo de cuentas.
                        MsgBox SIHOMsg(1585), vbCritical, "Mensaje"
                        
                        cmdAgrega.Enabled = False
                        cmdElimina.Enabled = False
                        
                        vlblnAceptaCambios = False
                        Exit Sub
                    End If
            Else
                'La información no existe
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
            End If
        End If
    End If

End Sub

Private Sub pCargaPaquetes()
    pCargaPaquetesDisponibles
    pCargaPaquetesAsignados
    pValidaBotones
End Sub

Private Sub pGrabaAsignacion(Optional pblnContinuar As Boolean)
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim vldblPrecio As Double
    Dim vldblIncremento As Double
    Dim vldblIVA As Double
    Dim lngPersonaGraba As Long
    Dim strParametrosSP As String
    Dim vldblDescuento As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim rsDescuento As New ADODB.Recordset
    Dim rsExpediente As New ADODB.Recordset
    Dim strParametrosPaquetes As String
    Dim intCantidad As Long
    Dim vllngCantPaquetesSinFacturar As Long
    Dim intBitValidaPaquetes As Long
    Dim vlaryParametrosSalida() As String
    Dim vlstrSentenciaexp As String
    Dim Totalnumpaquete As Long
    
    If vlblnAceptaCambios Then
        If lstPaquetesDisponibles.ListIndex > -1 Then
            '-----------------------------------------------------
            '|  Valida que la cuenta no esté cerrada
            '-----------------------------------------------------
            strParametrosSP = txtCvePaciente.Text & "|" & IIf(optTipoPaciente(0).Value, "I", "E")
            Set rsCuenta = frsEjecuta_SP(strParametrosSP, "Sp_ExSelPacienteIngreso")
            If rsCuenta!intcuentacerrada = 1 Or rsCuenta!intcuentacerrada = True Then
                '|  La cuenta del paciente está cerrada, no pueden realizarse modificaciones.
                MsgBox SIHOMsg(596), vbCritical, "Mensaje"
                Exit Sub
            End If
            
            pCargaArreglo alstrParametrosSalida, "|" & adDouble & "||" & adDouble & "||" & adDouble
            frsEjecuta_SP lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|PA|" & vgintTipoPaciente & "|" & vgintCveEmpresa & "|" & IIf(rsCuenta!intCveTipoIngreso = 7, "U", IIf(optTipoPaciente(0).Value, "I", "E")) & "|0|" & fstrFechaSQL("01/01/1900", , False) & "|" & vgintClaveEmpresaContable, "sp_pvselObtenerPreciobit", , , alstrParametrosSalida
            pObtieneValores alstrParametrosSalida, vldblPrecio, vldblIncremento, intbitpesos
            
            If (vldblPrecio > 0) Then ' (vldblPrecio <> -1)
            
                lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                
                If lngPersonaGraba = 0 Then Exit Sub
                EntornoSIHO.ConeccionSIHO.BeginTrans
                
                    vlstrSentencia = " Select pvConceptoFacturacion.SMYIVA "
                    vlstrSentencia = vlstrSentencia & " From PvPaquete "
                    vlstrSentencia = vlstrSentencia & "   Inner Join PvConceptoFacturacion On (pvPaquete.SMICONCEPTOFACTURA = pvConceptoFacturacion.SMICVECONCEPTO)"
                    vlstrSentencia = vlstrSentencia & " Where pvPaquete.INTNUMPAQUETE = " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
                    vldblIVA = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenForwardOnly)!smyIVA
                     
                    '------------------------------
                    'Obtener descuento del paquete
                    '------------------------------
                    vlstrSentencia = "SELECT SP_pvseldescuentopaquete('" & vgstrTipoPaciente & "', " & vgintCvePaciente & ", " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & _
                                                        ", " & vldblPrecio & ", " & vgintNumeroDepartamento & ", " & fstrFechaSQL(fdtmServerFecha) & ") As Descuento " & _
                                     "FROM DUAL"
                    Set rsDescuento = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                    If rsDescuento.RecordCount > 0 Then
                        vldblDescuento = rsDescuento!Descuento
                    Else
                        vldblDescuento = 0
                    End If
                    
                    strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|" & "AGREGAR"
                    
                    intCantidad = 1
                    
                    frsEjecuta_SP strParametrosPaquetes, "FN_PVSELCANTIDADPAQUETES", True, intCantidad
                    
                    If intCantidad = 0 Then
                        vlstrSentencia = "INSERT INTO PvPaquetePaciente (INTMOVPACIENTE, CHRTIPOPACIENTE, INTNUMPAQUETE, mnyPrecioPaquete, MNYIVAPAQUETE, intCveEmpleado, mnydescuento,intcantidad, bitpesos) " & _
                                     "VALUES (" & vgintCvePaciente & " ,'" & vgstrTipoPaciente & "', " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & ", " & vldblPrecio & ", " & (vldblPrecio - vldblDescuento) * (vldblIVA / 100) & ", " & str(lngPersonaGraba) & ", " & vldblDescuento & ",1," & intbitpesos & ")"
                    Else
                        vlstrSentencia = "update pvpaquetepaciente set intcantidad = nvl(intcantidad,0) + 1 " & _
                        "where intmovpaciente =" & vgintCvePaciente & " and chrtipopaciente ='" & vgstrTipoPaciente & "' and intnumpaquete = " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & ""
                    End If
                    
                    pEjecutaSentencia vlstrSentencia
                    
                      ' Ejecutar la consulta
                    Set rsExpediente = frsRegresaRs("SELECT COUNT(*) AS TotalCount FROM pvpaquetepaciente WHERE intmovpaciente = " & vgintCvePaciente & "", adLockOptimistic, adOpenDynamic)
                    
                    ' Obtener el resultado y almacenarlo en la variable
                    If Not rsExpediente.EOF Then
                        Totalnumpaquete = rsExpediente.Fields("TotalCount").Value
                    End If
                   
                    
                    If Totalnumpaquete <= 1 Then
                     'actualiza paquete en el expediente del paciente
                        vlstrSentenciaexp = "update expacienteingreso set INTCVEPAQUETE = " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & _
                        " where intnumcuenta =" & vgintCvePaciente & "                 "
                     pEjecutaSentencia vlstrSentenciaexp
                     
                    End If
                    
                    strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|" & "CANTIDAD"
                    
                    intCantidad = 1
                    
                    ' Busca cantidad de paquetes asignados
                    
                    frsEjecuta_SP strParametrosPaquetes, "FN_PVSELCANTIDADPAQUETES", True, intCantidad
                    
                    ' Valida que el paquete no esté facturado
                    vllngCantPaquetesSinFacturar = 1
                    strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
                    frsEjecuta_SP strParametrosPaquetes, "FN_PVSELPAQUETESINFACTURAR", True, vllngCantPaquetesSinFacturar
                    
                    ' Regresa bit para validar paquetes
                    intBitValidaPaquetes = 1
                    frsEjecuta_SP lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex), "FN_PVSELVALIDACARGOSPAQUETE", True, intBitValidaPaquetes
                    
                    'vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|" & intCantidad & "|" & intBitValidaPaquetes & "|" & -1
                    vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|" & vllngCantPaquetesSinFacturar & "|" & intBitValidaPaquetes & "|" & -1
                    frsEjecuta_SP vgstrParametrosSP, "sp_pvupdcargospaquete"
                    
                    'si se manejan cajas de material, relacionadas al paquete se generan indicaciones medicas
                    frsEjecuta_SP Trim(txtCvePaciente.Text) & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex) & "|1", "SP_EXUPDPAQUETEINDICACIONES"
                    
                    If lstPaquetesAsignados.ListCount = 0 Then
                        vlstrSentencia = ""
                        vlstrSentencia = " Update PvPaquetePaciente "
                        vlstrSentencia = vlstrSentencia & " Set PvPaquetePaciente.BITPAQUETEDEFAULT = 1 "
                        vlstrSentencia = vlstrSentencia & " WHERE PvPaquetePaciente.INTMOVPACIENTE = " & txtCvePaciente.Text & " AND "
                        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND "
                        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.INTNUMPAQUETE = " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
                        
                        pGuardarLogTransaccion Me.Name, EnmGrabar, lngPersonaGraba, "ASIGNACION DE PAQUETE", "Cta. " & CStr(vgintCvePaciente) & " " & vgstrTipoPaciente & " - Paq. * " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
        
                        pEjecutaSentencia vlstrSentencia
                        pPasaPaquete (CStr(vldblPrecio))
                        lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex) = "* " & lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex)
                    Else
                        pGuardarLogTransaccion Me.Name, EnmGrabar, lngPersonaGraba, "ASIGNACION DE PAQUETE", "Cta. " & CStr(vgintCvePaciente) & " " & vgstrTipoPaciente & " - Paq. " & lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
                        pPasaPaquete (CStr(vldblPrecio))
                    End If
                    vlstrSentencia = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
                        "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.AsignarPaquetes
                    pEjecutaSentencia vlstrSentencia
                    
                    vlstrSentencia = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.AsignarPaquetes & "," & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & ")"
                    pEjecutaSentencia vlstrSentencia
                
                    pblnContinuar = False
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                pCargaPaquetesAsignados
            Else
                '301 = El elemento seleccionado no cuenta con un precio capturado.
                '648 = ¡Imposible realizar la asignación con precio cero!
                MsgBox SIHOMsg(IIf(vldblPrecio = -1, 301, 648)), vbInformation, "Mensaje"
                pblnContinuar = True
            End If
        End If
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGrabaAsignacion"))
End Sub

Private Sub pEliminaPaquete(pblnPonPredeterminado As Boolean, Optional pblnContinuar As Boolean)
On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim SQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim rsPQFacturado As New ADODB.Recordset
    Dim vlblnElimina As Boolean
    Dim strParametrosPaquetes As String
    Dim intCantidad As Long
    Dim intBitValidaPaquetes As Long
    Dim intPaqueteSinMarcar As Long
    Dim lngPersonaGraba As Long
    Dim vllongOldPaqueteDefault As Long
    Dim vlintPaquetesFacturados As Integer
    Dim vllgnCargosPaqueteEnGrupo As Long
    
    If lstPaquetesAsignados.ListIndex > -1 Then
        '------------------------------------------------
        '|  Valida que el paquete no esté facturado
        '------------------------------------------------
        ' Función que regresa paquetes sin facturar
        intCantidad = 1
        strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        frsEjecuta_SP strParametrosPaquetes, "FN_PVSELPAQUETESINFACTURAR", True, intCantidad

        If intCantidad <= 0 Then
            '|  El paquete está cobrado, no se puede eliminar.
            MsgBox SIHOMsg(738), vbCritical, "Mensaje"
            Exit Sub
        End If
        
        'Valida que el paquete no se haya marcado para facturar al paciente
        intPaqueteSinMarcar = 1
        vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        frsEjecuta_SP strParametrosPaquetes, "Sp_PvSelPaqueteSinMarcar", True, intPaqueteSinMarcar
        If intPaqueteSinMarcar = 0 Then
            'No se puede borrar el cargo, ha sido marcado para facturar
            MsgBox SIHOMsg(1055), vbCritical, "Mensaje"
            Exit Sub
        End If
        
        ' | Caso 17412
        SQL = "select * from pvpaquete where pvpaquete.intnumpaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
        If rsTemp.RecordCount > 0 Then
            If rsTemp!intOrigen = 1 Then
                MsgBox "El paquete proviene de un presupuesto, no se puede eliminar.", vbCritical, "Mensaje"
                Exit Sub
            End If
        End If
        
'        vllgnCargosPaqueteEnGrupo = 1
'        strParametrosPaquetes = txtCvePaciente.Text & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
'        frsEjecuta_SP strParametrosPaquetes, "FN_PVSELCARGOSPAQUETEENGRUPO", True, vllgnCargosPaqueteEnGrupo
'        If vllgnCargosPaqueteEnGrupo <> 0 Then
'            'No se puede borrar el cargo, ha sido marcado para facturar
'            MsgBox SIHOMsg(1055), vbCritical, "Mensaje"
'            Exit Sub
'        End If
        
        
        strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & "|" & "QUITAR"
        intCantidad = 1
        ' Función para validar la cantidad de paquetes del paciente
        frsEjecuta_SP strParametrosPaquetes, "FN_PVSELCANTIDADPAQUETES", True, intCantidad
        
        lngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            
        If lngPersonaGraba = 0 Then Exit Sub
        EntornoSIHO.ConeccionSIHO.BeginTrans
        If Mid(lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex), 1, 1) = "*" Then
        
            If MsgBox("¿Está seguro que desea eliminar el paquete predeterminado?", vbYesNo + vbExclamation, "Mensaje") = vbYes Then

                '-- Revisa si se ha asignado el paquete a algún cargo --
                vlblnElimina = True
                SQL = "Select intnumpaquete from pvcargo "
                SQL = SQL & " WHERE CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND INTMOVPACIENTE = " & vgintCvePaciente & " AND INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                SQL = SQL & " AND intnumPaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                Set rsTemp = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
                If rsTemp.RecordCount > 0 Then
                    If MsgBox("¿Se perderá la relación de los cargos con el paquete, ¿Desea continuar?", vbYesNo + vbExclamation, "Mensaje") = vbNo Then
                        vlblnElimina = False
                    Else
                                
                        '-- Se elimina la relación al paquete en los cargos --
                        intBitValidaPaquetes = 1
                        frsEjecuta_SP lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex), "FN_PVSELVALIDACARGOSPAQUETE", True, intBitValidaPaquetes
                        
                        strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                        frsEjecuta_SP strParametrosPaquetes, "SP_PVUPDQUITARCARGOSENPAQUETE"

                        vlintPaquetesFacturados = 0
                        SQL = "SELECT SUM(NVL(intcantidadfacturada,0)) Cantidad " & _
                              "FROM PVPAQUETEPACIENTEFACTURADO " & _
                              "WHERE chrestatus = 'F' " & _
                                "AND intmovpaciente = " & vgintCvePaciente & " " & _
                                "AND chrtipopaciente = '" & vgstrTipoPaciente & "' " & _
                                "AND intnumpaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                        Set rsPQFacturado = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
                        If rsPQFacturado.RecordCount > 0 Then
                            vlintPaquetesFacturados = IIf(IsNull(rsPQFacturado!cantidad), 0, rsPQFacturado!cantidad)
                        End If
                        
                        If ((intCantidad - 1) - vlintPaquetesFacturados) >= 1 Then
                            vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & "|" & (intCantidad - 1) - vlintPaquetesFacturados & "|" & intBitValidaPaquetes & "|" & -1
                            frsEjecuta_SP vgstrParametrosSP, "sp_pvupdcargospaquete"
                        End If

                    End If
                End If
                rsTemp.Close
                If vlblnElimina Then
                    lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex) = Mid(lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex), 3, Len(lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex)))
                    
                    If intCantidad = 1 Then
                        vlstrSentencia = "DELETE FROM PvPaquetePaciente WHERE " & _
                        "CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND INTMOVPACIENTE = " & vgintCvePaciente & " AND " & _
                        "INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                    Else
                        vlstrSentencia = "Update PvPaquetePaciente set intcantidad = intcantidad -1 " & _
                                        "where CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and INTMOVPACIENTE = " & vgintCvePaciente & " and " & _
                                        "INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                    End If
                    
                    pEjecutaSentencia "delete from PVPaquetePacienteDetalle " & _
                    "where CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and INTMOVPACIENTE = " & vgintCvePaciente & " and " & _
                    "INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & " and intConsecutivo = " & intCantidad
                    
                    pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "DESASIGNACIÓN DE PAQUETE", "Cta. " & CStr(vgintCvePaciente) & " " & vgstrTipoPaciente & " - Paq. * " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
                    

                    vllongOldPaqueteDefault = lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)

                    pEjecutaSentencia vlstrSentencia
                    
                    If intCantidad = 1 Then
                        'si se manejan cajas de material, relacionadas al paquete se eliminan indicaciones medicas precargadas
                        frsEjecuta_SP vgintCvePaciente & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & "|0", "SP_EXUPDPAQUETEINDICACIONES"
                    End If
                    
                    pRegresaPaquete
                    
                    'Pone un nuevo paquete predeterminado dependiento del parámetro pblnPonPredeterminado y de siexisten todavía paquetes asignados
                    If lstPaquetesAsignados.ListCount > 0 And pblnPonPredeterminado Then
                        vlstrSentencia = ""
                        vlstrSentencia = " Update PvPaquetePaciente "
                        vlstrSentencia = vlstrSentencia & " Set PvPaquetePaciente.BITPAQUETEDEFAULT = 1 "
                        vlstrSentencia = vlstrSentencia & " WHERE PvPaquetePaciente.INTMOVPACIENTE = " & txtCvePaciente.Text & " AND "
                        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND "
                        vlstrSentencia = vlstrSentencia & "    PvPaquetePaciente.INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(0)
                        pEjecutaSentencia vlstrSentencia
                        lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex) = "* " & lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex)
                        pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "CAMBIO DE PAQUETE PREDETERMINADO", "Cta. " & CStr(vgintCvePaciente) & " " & vgstrTipoPaciente & " - Paq. " & vllongOldPaqueteDefault & "->" & lstPaquetesAsignados.ItemData(0)
                    End If
                End If
                pblnContinuar = False
            Else
                pblnContinuar = True
            End If
        Else
            lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex) = Mid(lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex), 1, Len(lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex)))
             '-- Se elimina la relación al paquete en los cargos --
            strParametrosPaquetes = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            frsEjecuta_SP strParametrosPaquetes, "SP_PVUPDQUITARCARGOSENPAQUETE"
                                
            vlintPaquetesFacturados = 0
            SQL = "SELECT SUM(intcantidadfacturada) Cantidad " & _
                  "FROM PVPAQUETEPACIENTEFACTURADO " & _
                  "WHERE chrestatus = 'F' " & _
                    "AND intmovpaciente = " & vgintCvePaciente & " " & _
                    "AND chrtipopaciente = '" & vgstrTipoPaciente & "' " & _
                    "AND intnumpaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            Set rsPQFacturado = frsRegresaRs(SQL, adLockOptimistic, adOpenDynamic)
            If rsPQFacturado.RecordCount > 0 Then
                vlintPaquetesFacturados = IIf(IsNull(rsPQFacturado!cantidad), 0, rsPQFacturado!cantidad)
            End If
            
            ' Regresa bit para validar paquetes
            intBitValidaPaquetes = 1
            frsEjecuta_SP lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex), "FN_PVSELVALIDACARGOSPAQUETE", True, intBitValidaPaquetes
            
            If ((intCantidad - 1) - vlintPaquetesFacturados) >= 1 Then
                vgstrParametrosSP = vgintCvePaciente & "|" & vgstrTipoPaciente & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & "|" & ((intCantidad - 1) - vlintPaquetesFacturados) & "|" & intBitValidaPaquetes & "|" & -1
                frsEjecuta_SP vgstrParametrosSP, "sp_pvupdcargospaquete"
            End If
                                
            If intCantidad = 1 Then
                vlstrSentencia = "DELETE FROM PvPaquetePaciente WHERE " & _
                    "CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' AND rownum < 2 AND " & _
                    "INTMOVPACIENTE = " & vgintCvePaciente & " AND INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            Else
                vlstrSentencia = "UPDATE PvPaquetePaciente set intcantidad = intcantidad -1 " & _
                    "where CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and INTMOVPACIENTE = " & vgintCvePaciente & " and " & _
                    "INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            End If
            pEjecutaSentencia "delete from PVPaquetePacienteDetalle " & _
            "where CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and INTMOVPACIENTE = " & vgintCvePaciente & " and " & _
            "INTNUMPAQUETE = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & " and intConsecutivo = " & intCantidad
               
            If intCantidad = 1 Then
                'si se manejan cajas de material, relacionadas al paquete se eliminan indicaciones medicas precargadas
                frsEjecuta_SP vgintCvePaciente & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|" & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & "|0", "SP_EXUPDPAQUETEINDICACIONES"
            End If
            
            pGuardarLogTransaccion Me.Name, EnmBorrar, lngPersonaGraba, "DESASIGNACIÓN DE PAQUETE", "Cta. " & CStr(vgintCvePaciente) & " " & vgstrTipoPaciente & " - Paq. " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
            pEjecutaSentencia vlstrSentencia
            pRegresaPaquete
            
        End If
        
        SQL = "delete from pvTipoPacienteProceso where pvTipoPacienteProceso.intnumerologin = " & vglngNumeroLogin & _
            "and pvTipoPacienteProceso.intproceso = " & enmTipoProceso.AsignarPaquetes
        pEjecutaSentencia SQL
        
        SQL = "insert into pvTipoPacienteProceso (intnumerologin, intproceso, chrtipopaciente) values(" & vglngNumeroLogin & "," & enmTipoProceso.AsignarPaquetes & "," & IIf(optTipoPaciente(0).Value, "'I'", "'E'") & ")"
        pEjecutaSentencia SQL
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        pCargaPaquetesAsignados
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEliminaPaquete"))
End Sub

Private Sub pPasaPaquete(pstrPrecio As String)
    lstPaquetesAsignados.AddItem lstPaquetesDisponibles.Text & "(" & FormatCurrency(CStr(pstrPrecio), 2) & ")", lstPaquetesAsignados.ListCount
    lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListCount - 1) = lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListIndex)
    lstPaquetesDisponibles.ListIndex = IIf(lstPaquetesDisponibles.ListCount > 0, 0, -1)
    lstPaquetesAsignados.ListIndex = IIf(lstPaquetesAsignados.ListCount > 0, 0, -1)
    pValidaBotones
End Sub

Private Sub pRegresaPaquete()
    lstPaquetesDisponibles.AddItem fstrQuitaPrecio(lstPaquetesAsignados.Text), lstPaquetesDisponibles.ListCount
    lstPaquetesDisponibles.ItemData(lstPaquetesDisponibles.ListCount - 1) = lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
    lstPaquetesAsignados.RemoveItem lstPaquetesAsignados.ListIndex
    lstPaquetesAsignados.ListIndex = IIf(lstPaquetesAsignados.ListCount > 0, 0, -1)
    lstPaquetesDisponibles.ListIndex = IIf(lstPaquetesDisponibles.ListCount > 0, 0, -1)
    pValidaBotones
End Sub

Private Sub pValidaBotones()
    If lstPaquetesAsignados.ListCount > 0 Then cmdElimina.Enabled = True
    If lstPaquetesDisponibles.ListCount > 0 Then cmdAgrega.Enabled = True
End Sub

Private Sub txtCvePaciente_KeyPress(KeyAscii As Integer)
    If UCase(Chr(KeyAscii)) = "E" Then optTipoPaciente(1).Value = True
    If UCase(Chr(KeyAscii)) = "I" Then optTipoPaciente(0).Value = True
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub pCargaPaquetesDisponibles()
    Dim vlrsPaquetes As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlIntCont As Integer

    Set vlrsPaquetes = frsEjecuta_SP(txtCvePaciente.Text & "|" & IIf(optTipoPaciente(0).Value, "I", "E") & "|1", "Sp_PvSelPaqueteNoAsignadoPac")
    pLlenarListRs lstPaquetesDisponibles, vlrsPaquetes, 0, 1
    lstPaquetesDisponibles.ListIndex = IIf(lstPaquetesDisponibles.ListCount > 0, 0, -1)

End Sub

Private Sub pCargaPaquetesAsignados()
    Dim vlrsPaquetes As New ADODB.Recordset
    Dim vlstrSentencia As String
    grdDetallePaquete.Rows = 1
    vlstrSentencia = ""
    vlstrSentencia = " SELECT PvPaquete.INTNUMPAQUETE, RTRIM(PvPaquete.CHRDESCRIPCION) || '($' || cast(PvPaquetePaciente.mnyPrecioPaquete as varchar(20)) || ' '||CASE WHEN PvPaquetePaciente.BITPESOS IS NULL THEN 'PESOS' ELSE CASE WHEN PvPaquetePaciente.BITPESOS = 1 THEN 'PESOS' ELSE 'DÓLARES' END END ||') (' || cast(PvPaquetePaciente.intcantidad as varchar2(10)) || ')' as Descripcion "
    vlstrSentencia = vlstrSentencia & " FROM PvPaquetePaciente INNER JOIN PvPaquete ON (PvPaquetePaciente.INTNUMPAQUETE = PvPaquete.INTNUMPAQUETE) "
    vlstrSentencia = vlstrSentencia & " WHERE PvPaquetePaciente.INTMOVPACIENTE = " & vgintCvePaciente & " " & _
    "AND PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "'"
    
    Set vlrsPaquetes = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    pLlenarListRs lstPaquetesAsignados, vlrsPaquetes, 0, 1
    
    vlstrSentencia = "SELECT PvPaquetePaciente.INTNUMPAQUETE FROM PvPaquetePaciente WHERE PvPaquetePaciente.BITPAQUETEDEFAULT = 1 AND PvPaquetePaciente.INTMOVPACIENTE = " & vgintCvePaciente & " AND PvPaquetePaciente.CHRTIPOPACIENTE = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "'"
    Set vlrsPaquetes = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    If vlrsPaquetes.RecordCount > 0 Then
       lstPaquetesAsignados.ListIndex = fintLocalizaLst(lstPaquetesAsignados, vlrsPaquetes!intnumpaquete)
       lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex) = "* " & lstPaquetesAsignados.List(lstPaquetesAsignados.ListIndex)
    End If

    lstPaquetesAsignados.ListIndex = IIf(lstPaquetesAsignados.ListCount > 0, 0, -1)
End Sub

Private Function fstrQuitaPrecio(pstrText) As String
    Dim vlIntCont As Integer
    
    For vlIntCont = Len(pstrText) - 1 To 1 Step -1
        If Mid(pstrText, vlIntCont, 1) = "(" Then
            fstrQuitaPrecio = Mid(pstrText, 1, vlIntCont - 1)
            Exit For
        End If
    Next
End Function

Private Sub pMostrarPaquete()
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim intCount As Integer
    Dim intCantidad As Integer
    grdDetallePaquete.Rows = 1
    strSQL = "select intCantidad" & _
    " from PVPaquetePaciente" & _
    " where intMovPAciente = " & vgintCvePaciente & " and chrTipoPaciente = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and intNumPaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
    Set rs = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        intCantidad = rs!intCantidad
    End If
    rs.Close
    
    For intCount = 1 To intCantidad
        grdDetallePaquete.AddItem ""
        grdDetallePaquete.TextMatrix(intCount, 1) = "Paquete " & intCount
        strSQL = "select * from PVPaquetePacienteDetalle " & _
        " where intMovPAciente = " & vgintCvePaciente & " and chrTipoPaciente = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and intNumPaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & " and intConsecutivo = " & intCount
        Set rs = frsRegresaRs(strSQL, adLockReadOnly, adOpenForwardOnly)
        If Not rs.EOF Then
            grdDetallePaquete.TextMatrix(intCount, 2) = rs!NUMPORCENTAJE
        Else
            grdDetallePaquete.TextMatrix(intCount, 2) = "100"
        End If
        rs.Close
    Next

End Sub

Private Function pActualizaPorcentaje(intConsecutivo As Integer, intPorcentaje As Integer)
    Dim rs As ADODB.Recordset
    Dim strSQL As String

    strSQL = "select * from PVPaquetePacienteDetalle " & _
    " where intMovPAciente = " & vgintCvePaciente & " and chrTipoPaciente = '" & IIf(optTipoPaciente(0).Value, "I", "E") & "' and intNumPaquete = " & lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex) & " and intConsecutivo = " & intConsecutivo
    Set rs = frsRegresaRs(strSQL, adLockOptimistic, adOpenStatic)
    If rs.EOF Then
        rs.AddNew
        rs!INTMOVPACIENTE = vgintCvePaciente
        rs!CHRTIPOPACIENTE = IIf(optTipoPaciente(0).Value, "I", "E")
        rs!intnumpaquete = lstPaquetesAsignados.ItemData(lstPaquetesAsignados.ListIndex)
        rs!intConsecutivo = intConsecutivo
    End If
    rs!NUMPORCENTAJE = intPorcentaje
    rs.Update
    rs.Close
End Function
