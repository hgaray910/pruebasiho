VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPacientesConvenios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de paciente y convenio para factura de asistencia social"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Height          =   495
      Left            =   5760
      Picture         =   "frmPacientesConvenios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   6450
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11475
      Begin VB.CommandButton cmdAsignar 
         Height          =   495
         Left            =   10710
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPacientesConvenios.frx":01A2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Agregar tipo de paciente o convenio"
         Top             =   375
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.ComboBox cboPacienteConvenio 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Tipos de paciente y convenio"
         Top             =   570
         Width           =   10275
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfgPacienteConvenio 
         Height          =   4515
         Left            =   270
         TabIndex        =   5
         ToolTipText     =   "Listado de tipos de paciente o convenios a los cuales se les podrá generar factura de asistencia social"
         Top             =   1230
         Width           =   10950
         _cx             =   19315
         _cy             =   7964
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
         BackColorBkg    =   -2147483644
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
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPacientesConvenios.frx":0694
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipos de paciente y convenio"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Tipos de paciente y convenio"
         Top             =   330
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmPacientesConvenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public llngNumOpcion As Long
Private Enum enmStatus
    stedicion = 1
    stConsulta = 2
End Enum
Dim stEstado As enmStatus

Private Sub pPonEstado(stNuevoEstado As enmStatus)
    Select Case stNuevoEstado
        Case stConsulta
            stEstado = stConsulta
            cmdGrabar.Enabled = False
        Case stedicion
            stEstado = stedicion
            cmdGrabar.Enabled = True
    End Select
End Sub

Private Sub cmdAsignar_Click()
    Dim vlstrSentencia  As String
    Dim rs As New ADODB.Recordset
    
    If Not fblnExistePacienteConvenio(cboPacienteConvenio.Text) Then
    
        vlstrSentencia = "select intcveempresa cve, trim(vchdescripcion) nombre, 1 PacienteConvenio from ccempresa where trim(vchdescripcion) = " & "'" & Trim(cboPacienteConvenio.Text) & "'"
        Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
        If rs.RecordCount > 0 Then
            vsfgPacienteConvenio.AddItem cboPacienteConvenio.ItemData(cboPacienteConvenio.ListIndex) & vbTab & cboPacienteConvenio.Text & vbTab & "EMPRESA"
        Else
            vsfgPacienteConvenio.AddItem cboPacienteConvenio.ItemData(cboPacienteConvenio.ListIndex) & vbTab & cboPacienteConvenio.Text & vbTab & "TIPO DE PACIENTE"
        End If
        
        vsfgPacienteConvenio.Select 1, 1
        vsfgPacienteConvenio.Sort = flexSortGenericAscending
    Else
        'El tipo de paciente o convenio ya se encuentra en el listado de la factura de asistencia social.
        MsgBox SIHOMsg(1587), vbCritical, "Mensaje"
        cboPacienteConvenio.SetFocus
    End If
    
    pPonEstado stedicion
End Sub

Private Function fblnExistePacienteConvenio(strPacienteConvenio As String) As Boolean
    Dim intCont As Integer
    Dim rsPacienteConvenio As New ADODB.Recordset
    Dim strParametros As String
    
    fblnExistePacienteConvenio = False
   
    'Verifica que el el tipo de paciente o convenio no se encuentre en el listado de la factura de asistencia social
    For intCont = 1 To vsfgPacienteConvenio.Rows - 1
        If vsfgPacienteConvenio.TextMatrix(intCont, 1) = strPacienteConvenio Then
            fblnExistePacienteConvenio = True
            Exit Function
        End If
    Next
End Function

Private Sub cmdGrabar_Click()
    Dim strParametros As String
    Dim intCont As Integer
    Dim llngPersonaGraba As Long
    Dim intPacienteConvenio As Integer
    Dim rsPacienteConvenio As New ADODB.Recordset
    
On Error GoTo NotificaError
        
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        strParametros = -1
        
        Set rsPacienteConvenio = frsEjecuta_SP(strParametros, "SP_PVSELPACIENTECONVENIO")
        If rsPacienteConvenio.RecordCount > 0 Then
            frsEjecuta_SP strParametros, "SP_PVDELPVTIPOPACIENTEEMPRESA"
        End If
        For intCont = 1 To vsfgPacienteConvenio.Rows - 1
            If vsfgPacienteConvenio.TextMatrix(intCont, 2) = "EMPRESA" Then
                intPacienteConvenio = 1
            Else
                intPacienteConvenio = 0
            End If
            strParametros = vsfgPacienteConvenio.TextMatrix(intCont, 0) & "|" & intPacienteConvenio
            frsEjecuta_SP strParametros, "SP_PVINSPVTIPOPACIENTEEMPRESA"
        Next
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, llngPersonaGraba, "TIPOS DE PACIENTE Y CONVENIO PARA FACTURA DE ASISTENCIA SOCIAL", CStr(vgintNumeroDepartamento))
        EntornoSIHO.ConeccionSIHO.CommitTrans
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        cboPacienteConvenio.SetFocus
    End If
    pPonEstado stConsulta
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdGrabar_Click"))
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
    pConfiguraGridPacienteConvenio
    pCargaPacienteConvenio
    pCargaListadoPacienteConvenio
End Sub

Private Sub pCargaPacienteConvenio()
    Dim vlstrSentencia  As String
    Dim rsPacienteConvenio As New ADODB.Recordset
    
On Error GoTo NotificaError

    vlstrSentencia = "select * from " & _
                     "(select tnycvetipopaciente cve, trim(vchdescripcion) nombre, 0 PacienteConvenio from adtipopaciente where bitactivo = 1 and chrtipo <> 'CO' and bitutilizaconvenio = 0" & _
                     "Union All " & _
                     "select intcveempresa cve, trim(vchdescripcion) nombre, 1 PacienteConvenio from ccempresa where bitactivo = 1) info " & _
                     "order by nombre"
    Set rsPacienteConvenio = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboPacienteConvenio, rsPacienteConvenio, 0, 1
    cboPacienteConvenio.ListIndex = 0
    rsPacienteConvenio.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaMedico"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Select Case stEstado
        Case stedicion
            '  ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbExclamation, "Mensaje") = vbYes Then
                Cancel = False
            Else
                Cancel = True
                cboPacienteConvenio.SetFocus
            End If
    End Select
End Sub

Private Sub vsfgPacienteConvenio_DblClick()
    If vsfgPacienteConvenio.Rows >= 2 Then
        vsfgPacienteConvenio.RemoveItem vsfgPacienteConvenio.Row
        pPonEstado stedicion
    End If
End Sub

Public Sub pCargaListadoPacienteConvenio()
    Dim strParametros As String
    Dim rsPacienteConvenio As New ADODB.Recordset
    
    On Error GoTo NotificaError
    
    strParametros = cboPacienteConvenio.ItemData(cboPacienteConvenio.ListIndex)
    Set rsPacienteConvenio = frsEjecuta_SP(strParametros, "SP_PVSELPACIENTECONVENIO")
    If rsPacienteConvenio.RecordCount > 0 Then
        vsfgPacienteConvenio.Rows = 1
        Do Until rsPacienteConvenio.EOF
            vsfgPacienteConvenio.AddItem rsPacienteConvenio!tnycvetipopacempresa & vbTab & rsPacienteConvenio!Descripcion & vbTab & rsPacienteConvenio!TipoPacienteEmpresa
            rsPacienteConvenio.MoveNext
        Loop
    End If
    cboPacienteConvenio.ListIndex = IIf(cboPacienteConvenio.ListCount > 0, 0, -1)
    pPonEstado stConsulta
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaListadoPacienteConvenio"))
End Sub

Public Sub pConfiguraGridPacienteConvenio()

    With vsfgPacienteConvenio
        .Cols = 3
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Nombre|Tipo"
        .ColWidth(1) = 9000 'Nombre
        .ColWidth(2) = 50   'Tipo

        .ColAlignment(1) = flexAlignLeftBottom
        .ColAlignment(2) = flexAlignLeftBottom
 
        .ScrollBars = flexScrollBarBoth
    End With

End Sub
