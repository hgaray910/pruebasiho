VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmFormatoFacturaPacEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formatos de factura por tipo de paciente y empresa"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgregar 
      Height          =   540
      Left            =   10200
      MaskColor       =   &H80000014&
      Picture         =   "frmFormatoFacturaPacEmpresa.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Agregar"
      Top             =   795
      UseMaskColor    =   -1  'True
      Width           =   570
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   5160
      TabIndex        =   17
      Top             =   720
      Width           =   5010
      Begin VB.ComboBox cboFormato 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Seleccione el formato de factura"
         Top             =   190
         Width           =   3690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Formato"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   250
         Width           =   570
      End
   End
   Begin VB.Frame Frame5 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   5010
      Begin VB.ComboBox cboProcedencia 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Seleccione la procedencia"
         Top             =   190
         Width           =   3690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   250
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   5625
      Begin VB.OptionButton optGrupos 
         Caption         =   "Grupos"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         ToolTipText     =   "Grupos de cuentas"
         Top             =   220
         Width           =   855
      End
      Begin VB.OptionButton optExternos 
         Caption         =   "Externos"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Pacientes externos"
         Top             =   220
         Width           =   975
      End
      Begin VB.OptionButton optInternos 
         Caption         =   "Internos"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         ToolTipText     =   "Pacientes internos"
         Top             =   220
         Width           =   975
      End
      Begin VB.OptionButton optTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Todos los pacientes"
         Top             =   220
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo paciente"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   250
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5010
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione el departamento"
         Top             =   190
         Width           =   3690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   250
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   5167
      TabIndex        =   10
      Top             =   4830
      Width           =   600
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   50
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFormatoFacturaPacEmpresa.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   10695
      Begin VSFlex7LCtl.VSFlexGrid vsfFormatos 
         Height          =   2970
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   10425
         _cx             =   18389
         _cy             =   5239
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483636
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFormatoFacturaPacEmpresa.frx":0834
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
         ComboSearch     =   0
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "frmFormatoFacturaPacEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public llngNumeroOpcionModulo As Long
Dim llngRowActualizar As Long

Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then optTodos.SetFocus
End Sub

Private Sub cboFormato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdAgregar.SetFocus
End Sub

Private Sub cboProcedencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboFormato.SetFocus
End Sub

Private Sub cmdAgregar_Click()
On Error GoTo NotificaError

    If cboFormato.ListIndex = -1 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & cboFormato.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        cboFormato.SetFocus
    Else
        llngRowActualizar = 0
        If fblnConfiValida Then
            With vsfFormatos
                If llngRowActualizar <> 0 Then
                    .TextMatrix(llngRowActualizar, 4) = cboFormato.List(cboFormato.ListIndex)     'Formato de factura
                    .TextMatrix(llngRowActualizar, 7) = cboFormato.ItemData(cboFormato.ListIndex)        'cve formato factura
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = cboDepartamento.List(cboDepartamento.ListIndex)
                    .TextMatrix(.Rows - 1, 2) = IIf(optTodos.Value, "TODOS", IIf(optInternos.Value, "INTERNOS", IIf(optExternos.Value, "EXTERNOS", "GRUPOS")))
                    .TextMatrix(.Rows - 1, 3) = IIf(cboProcedencia.List(cboProcedencia.ListIndex) = "<TODOS>", "TODOS", cboProcedencia.List(cboProcedencia.ListIndex)) 'Procedencia
                    .TextMatrix(.Rows - 1, 4) = cboFormato.List(cboFormato.ListIndex)     'Formato de factura
                    .TextMatrix(.Rows - 1, 5) = cboDepartamento.ItemData(cboDepartamento.ListIndex)
                    .TextMatrix(.Rows - 1, 6) = IIf(cboProcedencia.List(cboProcedencia.ListIndex) = "<TODOS>", 0, cboProcedencia.ItemData(cboProcedencia.ListIndex))    'cve Procedencia/empresa (negativo procedencia)
                    .TextMatrix(.Rows - 1, 7) = cboFormato.ItemData(cboFormato.ListIndex)        'cve formato factura
                    vsfFormatos.Col = 1
                    vsfFormatos.Sort = flexSortGenericAscending
                End If
            End With
            pInicia
        End If
        If cboDepartamento.Enabled Then
            cboDepartamento.SetFocus
        Else
            optTodos.SetFocus
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregar_Click"))
End Sub
Private Function fblnConfiValida() As Boolean
On Error GoTo NotificaError
    Dim llngRow As Long
    
    fblnConfiValida = True
    With vsfFormatos
        For llngRow = 1 To .Rows - 1
            If Val(.TextMatrix(llngRow, 5)) = cboDepartamento.ItemData(cboDepartamento.ListIndex) And _
               .TextMatrix(llngRow, 2) = IIf(optTodos.Value, "TODOS", IIf(optInternos.Value, "INTERNOS", IIf(optExternos.Value, "EXTERNOS", "GRUPOS"))) And _
               Val(.TextMatrix(llngRow, 6)) = cboProcedencia.ItemData(cboProcedencia.ListIndex) Then
                'Igual a una configuración ya existente
                If Val(.TextMatrix(llngRow, 7)) = cboFormato.ItemData(cboFormato.ListIndex) Then
                    'Existe información con el mismo contenido.
                    MsgBox SIHOMsg(19), vbOKOnly + vbExclamation, "Mensaje"
                    fblnConfiValida = False
                Else
                    'Ya existe un formato de factura para ese departamento, tipo de paciente y procedencia.  ¿Desea actualizarlo?
                    If MsgBox(SIHOMsg(989), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        fblnConfiValida = True
                        llngRowActualizar = llngRow
                    Else
                        fblnConfiValida = False
                    End If
                End If
                Exit For
            End If
        Next llngRow
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnConfiValida"))
End Function

Private Sub cmdDelete_Click()
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
    Dim llngRow As Long
    Dim llngPersonaGraba As Long
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumeroOpcionModulo, "E", True) Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans

        frsEjecuta_SP CStr(vgintClaveEmpresaContable), "Sp_Pvdelformatofactura"
        With vsfFormatos
            For llngRow = 1 To .Rows - 1
                vgstrParametrosSP = .TextMatrix(llngRow, 5) _
                        & "|" & .TextMatrix(llngRow, 7) _
                        & "|2|" & IIf(Trim(.TextMatrix(llngRow, 2)) = "TODOS", "T", IIf(Trim(.TextMatrix(llngRow, 2)) = "INTERNOS", "I", IIf(Trim(.TextMatrix(llngRow, 2)) = "EXTERNOS", "E", "G"))) _
                        & "|" & IIf(Val(.TextMatrix(llngRow, 6)) >= 0, "", Val(.TextMatrix(llngRow, 6)) * -1) _
                        & "|" & IIf(Val(.TextMatrix(llngRow, 6)) > 0, Val(.TextMatrix(llngRow, 6)), "")

                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSDOCUMENTODEPARTAMENTO"
            Next llngRow
        End With
        
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, llngPersonaGraba, "FORMATO FACTURA POR TIPO PACIENTE Y EMPRESA", CStr(vgintNumeroDepartamento))
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        pConfiguraGrid
        pLlenaGrid
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        If cboDepartamento.Enabled Then
            cboDepartamento.SetFocus
        Else
            optTodos.SetFocus
        End If
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        pConfiguraGrid
        pLlenaGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    pLlenaCombos
    pConfiguraGrid
    pLlenaGrid
    pInicia
End Sub

Private Sub pLlenaCombos()
On Error GoTo NotificaError

    Dim rsAux As New ADODB.Recordset
    
    '--- Combo departamentos ---
    vgstrParametrosSP = "-1|1|*|" & vgintClaveEmpresaContable
    Set rsAux = frsEjecuta_SP(vgstrParametrosSP, "sp_GnSelDepartamento")
    If rsAux.RecordCount > 0 Then
        pLlenarCboRs cboDepartamento, rsAux, 0, 1
        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento))
        cboDepartamento.Enabled = fblnRevisaPermiso(vglngNumeroLogin, llngNumeroOpcionModulo, "C", True)
    End If
    
    '--- Combo Procedencia ---
    Set rsAux = frsEjecuta_SP("1", "Sp_Gnseltipopacienteempresa")
    If rsAux.RecordCount > 0 Then
        pLlenarCboRs cboProcedencia, rsAux, 1, 0, 3
        cboProcedencia.ListIndex = 0
    End If
    
    '--- Combo formatos ---
    Set rsAux = frsEjecuta_SP("2", "Sp_Gnselformatomaestro")
    If rsAux.RecordCount > 0 Then
        pLlenarCboRs cboFormato, rsAux, 0, 1
        cboFormato.ListIndex = -1
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCombos"))
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError

    With vsfFormatos
        .Clear
        .Rows = 1
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Departamento|Tipo paciente|Procedencia|Formato de factura|||"
        .Cols = 8
        .ColWidth(0) = 100      '
        .ColWidth(1) = 2900     'Departamento
        .ColWidth(2) = 1100     'Tipo paciente
        .ColWidth(3) = 3100     'Procedencia
        .ColWidth(4) = 2900     'Formato de factura
        .ColWidth(5) = 0        'cve departamento
        .ColWidth(6) = 0        'cve tipoPaciente/empresa (negativo tipopaciente)
        .ColWidth(7) = 0        'cve formato factura
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub
Private Sub pLlenaGrid()
On Error GoTo NotificaError
    Dim rsFormatos As New ADODB.Recordset
    
    vgstrParametrosSP = "-1|*|" & vgintClaveEmpresaContable
    Set rsFormatos = frsEjecuta_SP(vgstrParametrosSP, "Sp_Pvselformatosfactura")
    With vsfFormatos
        Do While Not rsFormatos.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = rsFormatos!Depto
            .TextMatrix(.Rows - 1, 2) = rsFormatos!TipoPaciente
            .TextMatrix(.Rows - 1, 3) = rsFormatos!Procedencia
            .TextMatrix(.Rows - 1, 4) = rsFormatos!Formato
            .TextMatrix(.Rows - 1, 5) = rsFormatos!cveDepto
            .TextMatrix(.Rows - 1, 6) = rsFormatos!cveProcedencia
            .TextMatrix(.Rows - 1, 7) = rsFormatos!cveFormato    'negativo cuando es tipoPaciente
            rsFormatos.MoveNext
        Loop
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaGrid"))
End Sub
Private Sub pInicia()
    
    optTodos.Value = True
    optInternos.Value = False
    optExternos.Value = False
    optGrupos.Value = False
    cboFormato.ListIndex = -1
    cboProcedencia.ListIndex = 0

End Sub
Private Sub optExternos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboProcedencia.SetFocus
End Sub
Private Sub optGrupos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboProcedencia.SetFocus
End Sub
Private Sub optInternos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboProcedencia.SetFocus
End Sub
Private Sub optTodos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboProcedencia.SetFocus
End Sub
Private Sub vsfFormatos_DblClick()
On Error GoTo NotificaError

    If vsfFormatos.Row > 0 Then
        If fblnRevisaPermiso(vglngNumeroLogin, llngNumeroOpcionModulo, "C", True) Then
           '¿Está seguro de eliminar los datos?
           If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
               vsfFormatos.RemoveItem (vsfFormatos.Row)
               cmdSave.SetFocus
           End If
        Else
            '¡El usuario debe tener permiso de control total para eliminar los datos!
            MsgBox SIHOMsg(810), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":vsfFormatos_DblClick"))
End Sub
