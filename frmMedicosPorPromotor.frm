VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmMedicosPorPromotor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Médicos por promotor"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAsignar 
      Height          =   495
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMedicosPorPromotor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Agregar médico"
      Top             =   1120
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboMedico 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Médico disponible"
         Top             =   1200
         Width           =   4095
      End
      Begin VB.ComboBox cboPromotor 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Promotor"
         Top             =   480
         Width           =   4695
      End
      Begin VSFlex7LCtl.VSFlexGrid vsfgMedicosAsignados 
         Height          =   2655
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Médicos asignados"
         Top             =   1680
         Width           =   4695
         _cx             =   8281
         _cy             =   4683
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmMedicosPorPromotor.frx":04F2
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
         Caption         =   "Médicos disponibles"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Promotor"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Height          =   495
      Left            =   2475
      Picture         =   "frmMedicosPorPromotor.frx":0543
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   495
   End
End
Attribute VB_Name = "frmMedicosPorPromotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public llngNumOpcion As Long
Private Enum enmStatus
    stEdicion = 1
    stConsulta = 2
End Enum
Dim stEstado As enmStatus

Private Sub cboPromotor_Click()
    Dim strParametros As String
    Dim rsMedicoPromotor As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    strParametros = cboPromotor.ItemData(cboPromotor.ListIndex)
    Set rsMedicoPromotor = frsEjecuta_SP(strParametros, "SP_PVSELMEDICOPROMOTOR")
    vsfgMedicosAsignados.Rows = 1
    Do Until rsMedicoPromotor.EOF
        vsfgMedicosAsignados.AddItem rsMedicoPromotor!ClaveMedico & vbTab & rsMedicoPromotor!NombreMedico
        rsMedicoPromotor.MoveNext
    Loop
    cboMedico.ListIndex = IIf(cboMedico.ListCount > 0, 0, -1)
    pPonEstado stConsulta
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboPromotor_Click"))
End Sub


Private Sub pPonEstado(stNuevoEstado As enmStatus)
    Select Case stNuevoEstado
        Case stConsulta
            stEstado = stConsulta
            cmdGrabar.Enabled = False
        Case stEdicion
            stEstado = stEdicion
            cmdGrabar.Enabled = True
    End Select

End Sub

Private Sub cmdAsignar_Click()
    If Not fblnExisteMedicoAsignado(cboMedico.ItemData(cboMedico.ListIndex)) Then
        vsfgMedicosAsignados.AddItem cboMedico.ItemData(cboMedico.ListIndex) & vbTab & cboMedico.Text
        vsfgMedicosAsignados.Select 1, 1
        vsfgMedicosAsignados.Sort = flexSortGenericAscending
    Else
        '|  El médico ya fue asignado a éste u otro promotor.
        MsgBox SIHOMsg(1007), vbCritical, "Mensaje"
    End If
    
    pPonEstado stEdicion
End Sub


Private Function fblnExisteMedicoAsignado(intCveMedico As Integer) As Boolean
    Dim intCont As Integer
    Dim rsMedicoPromotor As New ADODB.Recordset
    Dim strParametros As String
    
    fblnExisteMedicoAsignado = False
    '|  Verifica que el médico no haya sido asignado a otro promotor
    strParametros = "-1"
    Set rsMedicoPromotor = frsEjecuta_SP(strParametros, "SP_PVSELMEDICOPROMOTOR")
    Do Until rsMedicoPromotor.EOF
        If rsMedicoPromotor!ClaveMedico <> cboPromotor.ItemData(cboPromotor.ListIndex) Then
            If rsMedicoPromotor!ClaveMedico = intCveMedico Then
                fblnExisteMedicoAsignado = True
                Exit Function
            End If
        End If
        rsMedicoPromotor.MoveNext
    Loop
    
    '|  Verifica que el médico no se encuentre actualmente asignado al promotor seleccionado
    For intCont = 1 To vsfgMedicosAsignados.Rows - 1
        If vsfgMedicosAsignados.TextMatrix(intCont, 0) = intCveMedico Then
            fblnExisteMedicoAsignado = True
            Exit Function
        End If
    Next

End Function

Private Sub cmdGrabar_Click()
    Dim strParametros As String
    Dim intCont As Integer
    Dim llngPersonaGraba As Long
    
On Error GoTo NotificaError
        
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        strParametros = cboPromotor.ItemData(cboPromotor.ListIndex) & "|-1"
        frsEjecuta_SP strParametros, "SP_PVDELMEDICOPROMOTOR", True
        For intCont = 1 To vsfgMedicosAsignados.Rows - 1
            strParametros = cboPromotor.ItemData(cboPromotor.ListIndex) & "|" & vsfgMedicosAsignados.TextMatrix(intCont, 0)
            frsEjecuta_SP strParametros, "SP_PVINSMEDICOPROMOTOR"
        Next
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, llngPersonaGraba, "MÉDICOS POR PROMOTOR", CStr(vgintNumeroDepartamento))
        EntornoSIHO.ConeccionSIHO.CommitTrans
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        cboPromotor_Click
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
    pCargaPromotor
    pCargaMedico
End Sub

Private Sub pCargaPromotor()
    Dim strParametros As String
    Dim rsPromotor As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    strParametros = "-1|1|" & vgintClaveEmpresaContable
    Set rsPromotor = frsEjecuta_SP(strParametros, "SP_GNSELEMPLEADO")
    pLlenarCboRs cboPromotor, rsPromotor, 0, 1
    cboPromotor.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaPromotor"))
End Sub

Private Sub pCargaMedico()
    Dim strParametros As String
    Dim rsMedico As New ADODB.Recordset
    
On Error GoTo NotificaError
    
    strParametros = "-1|1"
    Set rsMedico = frsEjecuta_SP(strParametros, "SP_EXSELMEDICO")
    pLlenarCboRs cboMedico, rsMedico, 0, 1
    cboMedico.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCargaMedico"))
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Select Case stEstado
        Case stEdicion
            '  ¿Desea abandonar la operación?
            If MsgBox(SIHOMsg(17), vbYesNo + vbExclamation, "Mensaje") = vbYes Then
                cboPromotor_Click
            End If
            Cancel = True
    End Select
End Sub

Private Sub vsfgMedicosAsignados_DblClick()
    vsfgMedicosAsignados.RemoveItem vsfgMedicosAsignados.Row
    pPonEstado stEdicion
    
End Sub
