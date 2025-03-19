VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmConfiguracionFUC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración para facturación con un concepto"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   173
      TabIndex        =   0
      Top             =   180
      Width           =   7545
      Begin VB.CommandButton cmdAgregar 
         Height          =   540
         Left            =   6840
         MaskColor       =   &H80000014&
         Picture         =   "frmConfiguracionFUC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Agregar"
         Top             =   2380
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.ComboBox cboTipoPaciente 
         Height          =   315
         Left            =   2385
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Tipo paciente"
         Top             =   585
         Width           =   5025
      End
      Begin VB.ComboBox cboEmpresaContable 
         Height          =   315
         Left            =   2385
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Empresa contable"
         Top             =   195
         Width           =   5025
      End
      Begin VB.ComboBox cboConcepto 
         Height          =   315
         Left            =   2385
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Concepto de facturación"
         Top             =   1995
         Width           =   5025
      End
      Begin VB.Frame frTasa 
         Height          =   495
         Left            =   2385
         TabIndex        =   13
         Top             =   1380
         Width           =   5025
         Begin VB.OptionButton optTasa 
            Caption         =   "Tasa 0%"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   7
            ToolTipText     =   "Tasa de facturación"
            Top             =   180
            Width           =   1215
         End
         Begin VB.OptionButton optTasa 
            Caption         =   "Tasa "
            Height          =   255
            Index           =   0
            Left            =   1275
            TabIndex        =   6
            ToolTipText     =   "Tasa de facturación"
            Top             =   165
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame frTipoIngreso 
         Height          =   495
         Left            =   2385
         TabIndex        =   12
         Top             =   900
         Width           =   5025
         Begin VB.OptionButton optTipoIngreso 
            Caption         =   "Interno"
            Height          =   255
            Index           =   0
            Left            =   1275
            TabIndex        =   4
            ToolTipText     =   "Paciente interno"
            Top             =   150
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optTipoIngreso 
            Caption         =   "Externo"
            Height          =   255
            Index           =   1
            Left            =   2475
            TabIndex        =   5
            ToolTipText     =   "Paciente externo"
            Top             =   150
            Width           =   1215
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de paciente"
         Height          =   255
         Left            =   105
         TabIndex        =   18
         Top             =   585
         Width           =   2155
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de ingreso"
         Height          =   255
         Left            =   105
         TabIndex        =   17
         Top             =   1035
         Width           =   2155
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   2155
      End
      Begin VB.Label Label4 
         Caption         =   "Tasa de IVA"
         Height          =   255
         Left            =   105
         TabIndex        =   15
         Top             =   1515
         Width           =   2155
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto de facturación"
         Height          =   255
         Left            =   105
         TabIndex        =   14
         Top             =   2025
         Width           =   2155
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdConfFUC 
      Height          =   1575
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Configuraciones de facturación único concepto"
      Top             =   3225
      Width           =   7545
      _cx             =   13309
      _cy             =   2778
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmConfiguracionFUC.frx":04F2
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
   Begin VB.Frame Frame2 
      Height          =   670
      Left            =   3630
      TabIndex        =   11
      Top             =   4820
      Width           =   630
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConfiguracionFUC.frx":058D
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar las configuraciones"
         Top             =   120
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmConfiguracionFUC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rbTasa1 As Double 'variable para almacenar la tasa contenida en el option izquierdo en la seccion tasa
Dim rbTasa2 As Double
'variables para almacenar las columnas
Dim colClaveTipoPaciente As Long
Dim colTipoPaciente As Long
Dim colTipoIngreso As Long
Dim colTasa As Long
Dim colClaveConcepto As Long
Dim colConcepto As Long
Dim vlblnbandera As Boolean 'Variable que indicara cuando agreguen o quiten un concepto
Dim vlblnBanderaInicio As Boolean 'Variable que nos indicara cuando entremos por primera vez a la pantalla para asi hacer el focus al combo
Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdAgregar.SetFocus
End Sub

Private Sub cboEmpresaContable_Click()
    If cboTipoPaciente.ListCount > 0 Then
        pObtenerTasa
    End If
End Sub

Private Sub cboEmpresaContable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       pObtenerTasa
       cboTipoPaciente.SetFocus
    End If
    
End Sub

Private Sub cboTipoPaciente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then optTipoIngreso(0).SetFocus

End Sub

Private Sub cmdAgregar_Click()
    agregaElemento
    vlblnbandera = True
End Sub

Private Sub cmdAplicar_Click()
    agregaElemento
End Sub

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim queryGuarda As String
    Dim vllngPersonaGraba As Long
    Dim rsConfFuc As New ADODB.Recordset
    Dim valorLog As String
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        pEjecutaSentencia ("Delete from PVCONFSOLOCONCEPTO where intempresacontable = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex))
        
        queryGuarda = "SELECT * FROM PVCONFSOLOCONCEPTO WHERE INTCONSECUTIVO = -1"
        Set rsConfFuc = frsRegresaRs(queryGuarda, adLockOptimistic, adOpenDynamic)
        
        For i = 1 To grdConfFUC.Rows - 1
            With rsConfFuc
                .AddNew
                !INTEMPRESACONTABLE = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
                !tnyCveTipoPaciente = grdConfFUC.TextMatrix(i, colClaveTipoPaciente)
                !chrTipoIngreso = IIf(grdConfFUC.TextMatrix(i, colTipoIngreso) = "Interno", "I", "E")
                !INTTASA = Mid(grdConfFUC.TextMatrix(i, colTasa), 1, Len(grdConfFUC.TextMatrix(i, colTasa)) - 1)
                !INTCONCEPTOFACTURACION = grdConfFUC.TextMatrix(i, colClaveConcepto)
                .Update
                valorLog = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) & " - " & grdConfFUC.TextMatrix(i, colClaveTipoPaciente) & " - " & _
                IIf(grdConfFUC.TextMatrix(i, colTipoIngreso) = "Interno", "I", "E") & " - " & grdConfFUC.TextMatrix(i, colTasa) & " - " & _
                grdConfFUC.TextMatrix(i, colClaveConcepto)
                pGuardarLogTransaccion "frmConfiguracionFUC", EnmGrabar, vglngNumeroLogin, "CONFIGURACION PARA FACTURACION CON UN CONCEPTO", valorLog
            End With
        Next i
        
        vlblnbandera = False
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
    End If
End Sub

Private Sub Form_Activate()
    If vlblnBanderaInicio = True Then
        If cboEmpresaContable.ListCount > 0 Then
            If cboEmpresaContable.Enabled = True Then
                cboEmpresaContable.SetFocus
            Else
                cboTipoPaciente.SetFocus
            End If
        End If
    End If
    vlblnBanderaInicio = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo NotificaError
        
    If KeyAscii = 27 Then
        If vlblnbandera = True Then
            '¿Desea abandonar la operación?
            If MsgBox(SIHOMsg("17"), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
        KeyAscii = 0
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strQuery As String
    Dim rsCbo As New ADODB.Recordset
    Dim rsbuscaempresa As New ADODB.Recordset
    Dim vlintnumero As Long
    
    vlblnBanderaInicio = True
    colClaveTipoPaciente = 0
    colTipoPaciente = 1
    colTipoIngreso = 2
    colTasa = 3
    colClaveConcepto = 4
    colConcepto = 5

    Me.Icon = frmMenuPrincipal.Icon
    'cargando el combo de empresa contable
    strQuery = "SELECT * FROM CNEmpresaContable WHERE bitActiva <> 0 ORDER BY vchNombre"
    Set rsCbo = frsRegresaRs(strQuery, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboEmpresaContable, rsCbo, 0, 1
    
    'Si tiene permisos de control total puede cambiar de empresa contable
    vlintnumero = IIf(cgstrModulo = "SI", 7033, 7030)
    
    If fblnRevisaPermiso(vglngNumeroLogin, vlintnumero, "C") Then
        cboEmpresaContable.Enabled = True
    Else
        cboEmpresaContable.Enabled = False
    End If
    
    'cargando el combo de tipo de paciente
    strQuery = "SELECT * FROM ADTIPOPACIENTE WHERE BITACTIVO <> 0 ORDER BY VCHDESCRIPCION"
    Set rsCbo = frsRegresaRs(strQuery, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboTipoPaciente, rsCbo, 0, 1
    cboTipoPaciente.ListIndex = 0
    
    'cargando el listado de conceptos de facturación
    strQuery = "SELECT * FROM PVCONCEPTOFACTURACION WHERE BITACTIVO <> 0 AND INTTIPO <> 1 ORDER BY CHRDESCRIPCION"
    Set rsCbo = frsRegresaRs(strQuery, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboConcepto, rsCbo, 0, 1
    cboConcepto.ListIndex = 0
    
    cboTipoPaciente.Enabled = False
    frTipoIngreso.Enabled = False
    frTasa.Enabled = False
    cboConcepto.Enabled = False
    vlblnbandera = False
    cargaGrd
    cboEmpresaContable.ListIndex = fintLocalizaCbo(cboEmpresaContable, Str(vgintClaveEmpresaContable))
    
End Sub
Public Sub cargaGrd()


    With grdConfFUC
            .Rows = 1
            .Cols = 6
            .FormatString = "Clave tipo paciente|Tipo paciente|Tipo ingreso|Tasa|Clave concepto|Concepto facturación"
            
            .ColWidth(colClaveTipoPaciente) = 0
            .ColWidth(colTipoPaciente) = 1899
            .ColWidth(colTipoIngreso) = 1050
            .ColWidth(colTasa) = 600
            .ColWidth(colClaveConcepto) = 0
            .ColWidth(colConcepto) = 2000
    End With
            
End Sub
Public Sub agregaElemento()
    Dim blnagregadato As Boolean
    Dim queryConcepto As String
    Dim rsDescUni As New ADODB.Recordset
    
            queryConcepto = "SELECT GNCATALOGOSATDETALLE.VCHCLAVE Clave, GNCATALOGOSATDETALLE.VCHDESCRIPCION Descripcion FROM GNCATALOGOSATRELACION " & _
            "INNER JOIN GNCATALOGOSATDETALLE ON GNCATALOGOSATDETALLE.INTIDCATALOGOSAT = 5 AND GNCATALOGOSATRELACION.INTIDREGISTRO = GNCATALOGOSATDETALLE.INTIDREGISTRO " & _
            "WHERE INTCVECONCEPTO = " & cboConcepto.ItemData(cboConcepto.ListIndex) & " AND CHRTIPOCONCEPTO = 'CF' AND INTDIFERENCIADOR = 1"
            Set rsDescUni = frsRegresaRs(queryConcepto, adLockOptimistic, adOpenDynamic)
            If IsNull(rsDescUni!clave) Or IsNull(rsDescUni!Descripcion) Then
                Call MsgBox("El concepto seleccionado no se encuentra relacionado a un concepto válido de el cátalogo del SAT.", vbExclamation, "Mensaje")
                Exit Sub
            End If
            queryConcepto = "SELECT GNCATALOGOSATDETALLE.VCHCLAVE Clave, GNCATALOGOSATDETALLE.VCHDESCRIPCION Descripcion FROM GNCATALOGOSATRELACION " & _
            "INNER JOIN GNCATALOGOSATDETALLE ON GNCATALOGOSATDETALLE.INTIDCATALOGOSAT = 4 AND GNCATALOGOSATRELACION.INTIDREGISTRO = GNCATALOGOSATDETALLE.INTIDREGISTRO " & _
            "WHERE INTCVECONCEPTO = " & cboConcepto.ItemData(cboConcepto.ListIndex) & " AND CHRTIPOCONCEPTO = 'CF' AND INTDIFERENCIADOR = 2"
            Set rsDescUni = frsRegresaRs(queryConcepto, adLockOptimistic, adOpenDynamic)
            If IsNull(rsDescUni!clave) Or IsNull(rsDescUni!Descripcion) Then
                Call MsgBox("El concepto seleccionado no se encuentra relacionado a una clave de unidad de medida válida del cátalogo del SAT.", vbExclamation, "Mensaje")
                Exit Sub
            End If
            
    blnagregadato = False
    With grdConfFUC
        ' verificamos que ya existe el dato en la tabla
        If .Rows > 1 Then
            For i = 1 To grdConfFUC.Rows - 1
                If .TextMatrix(i, colClaveTipoPaciente) = cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex) And .TextMatrix(i, colTipoPaciente) = cboTipoPaciente.Text And .TextMatrix(i, colTipoIngreso) = IIf(optTipoIngreso(0).Value = True, "Interno", "Externo") Then
                    blnagregadato = True
                End If
            Next
            If blnagregadato Then
                Call MsgBox("La configuración seleccionada ya ha sido agregada a la tabla de guardado.", vbExclamation, "Mensaje")
                Exit Sub
            Else ' si no existe el dato lo agregamos
                .AddItem (1)
                .TextMatrix(.Rows - 1, colClaveTipoPaciente) = cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)
                .TextMatrix(.Rows - 1, colTipoPaciente) = cboTipoPaciente.Text
                .TextMatrix(.Rows - 1, colTipoIngreso) = IIf(optTipoIngreso(0).Value = True, "Interno", "Externo")
                .TextMatrix(.Rows - 1, colTasa) = IIf(optTasa(0).Value = True, rbTasa1 & "%", 0 & "%")
                .TextMatrix(.Rows - 1, colClaveConcepto) = cboConcepto.ItemData(cboConcepto.ListIndex)
                .TextMatrix(.Rows - 1, colConcepto) = cboConcepto.Text
                
            End If
        Else
            .AddItem (1)
            .TextMatrix(.Rows - 1, colClaveTipoPaciente) = cboTipoPaciente.ItemData(cboTipoPaciente.ListIndex)
            .TextMatrix(.Rows - 1, colTipoPaciente) = cboTipoPaciente.Text
            .TextMatrix(.Rows - 1, colTipoIngreso) = IIf(optTipoIngreso(0).Value = True, "Interno", "Externo")
            .TextMatrix(.Rows - 1, colTasa) = IIf(optTasa(0).Value = True, rbTasa1 & "%", 0 & "%")
            .TextMatrix(.Rows - 1, colClaveConcepto) = cboConcepto.ItemData(cboConcepto.ListIndex)
            .TextMatrix(.Rows - 1, colConcepto) = cboConcepto.Text
            
        End If
        
    End With

End Sub

Private Sub grdConfFUC_DblClick()
With grdConfFUC
      If .Rows > 1 Then
        If .RowSel <> -1 Then
            .RemoveItem (.RowSel)
            vlblnbandera = True
        End If
    End If
End With
End Sub

Private Sub optTasa_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboConcepto.SetFocus
End Sub

Private Sub optTipoIngreso_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then optTasa(0).SetFocus
End Sub

Private Function pObtenerTasa()
    Dim rsTasa As New ADODB.Recordset
    Dim rsDat As New ADODB.Recordset
    Dim vchTasa As String
    Dim strQuery As String
    If cboEmpresaContable.ListCount <> 0 Then
        strQuery = "SELECT CNIMPUESTO.RELPORCENTAJE TASA FROM SIPARAMETRO INNER JOIN CNIMPUESTO ON SIPARAMETRO.VCHVALOR = CNIMPUESTO.SMICVEIMPUESTO WHERE VCHNOMBRE = 'INTTASAIMPUESTOHOSPITAL' AND INTCVEEMPRESACONTABLE = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
    Else
        Call MsgBox("Por favor seleccionar una empresa contable.", vbExclamation, "Mensaje")
        Exit Function
    End If
    Set rsTasa = frsRegresaRs(strQuery)
    If rsTasa.RecordCount > 0 Then
        vchTasa = rsTasa!TASA
        cboTipoPaciente.Enabled = True
        frTipoIngreso.Enabled = True
        frTasa.Enabled = True
        cboConcepto.Enabled = True
        
        strQuery = "SELECT PVCONFSOLOCONCEPTO.TNYCVETIPOPACIENTE CLAVE_TIPO," & _
        " ADTIPOPACIENTE.VCHDESCRIPCION TIPO_PACIENTE," & _
        " CASE WHEN PVCONFSOLOCONCEPTO.CHRTIPOINGRESO = 'E' THEN 'Externo' ELSE 'Interno' END TIPO_INGRESO," & _
        " PVCONFSOLOCONCEPTO.INTTASA || '%' TASA," & _
        " PVCONFSOLOCONCEPTO.INTCONCEPTOFACTURACION CLAVE_CONCEPTO," & _
         " PVCONCEPTOFACTURACION.CHRDESCRIPCION DESCRIPCION FROM PVCONFSOLOCONCEPTO " & _
         " INNER JOIN ADTIPOPACIENTE ON PVCONFSOLOCONCEPTO.TNYCVETIPOPACIENTE = ADTIPOPACIENTE.TNYCVETIPOPACIENTE " & _
         " INNER JOIN PVCONCEPTOFACTURACION ON PVCONFSOLOCONCEPTO.INTCONCEPTOFACTURACION = PVCONCEPTOFACTURACION.SMICVECONCEPTO " & _
         " WHERE INTEMPRESACONTABLE = " & cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) & " ORDER BY TIPO_PACIENTE, TIPO_INGRESO"
         Set rsDat = frsRegresaRs(strQuery)
         
         grdConfFUC.Rows = 1
         pLlenaVsfGrid grdConfFUC, rsDat, True, False, False
         
         optTasa(0).Caption = "Tasa " & vchTasa & "%"
                
        If vchTasa <> "" Then
            rbTasa1 = CDbl(vchTasa)
        Else
            rbTasa1 = 0
        End If
        
        If rbTasa1 = 0 Then
            optTasa(1).Visible = False
        Else
            optTasa(1).Visible = True
        End If
    Else
        optTasa(0).Caption = ""
        grdConfFUC.Rows = 1
    End If
End Function
