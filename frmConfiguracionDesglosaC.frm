VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfiguracionDesglosaC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos agrupa cargos factura mixta"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freBarra 
      Height          =   860
      Left            =   105
      TabIndex        =   7
      Top             =   2292
      Visible         =   0   'False
      Width           =   8895
      Begin MSComctlLib.ProgressBar pgbBarra 
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Top             =   420
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblTextoBarra 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   45
         TabIndex        =   9
         Top             =   120
         Width           =   8805
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   105
      TabIndex        =   5
      Top             =   120
      Width           =   8880
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   1545
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Empresa"
         Top             =   675
         Width           =   7200
      End
      Begin VB.ComboBox cboEmpresaContable 
         Height          =   315
         Left            =   1545
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa contable"
         Top             =   240
         Width           =   7200
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   705
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa contable"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   1695
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
      Height          =   3735
      Left            =   105
      TabIndex        =   2
      Top             =   1635
      Width           =   8880
      _cx             =   15663
      _cy             =   6588
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
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
      Left            =   4200
      TabIndex        =   4
      Top             =   5400
      Width           =   630
      Begin VB.CommandButton cmdGrabarRegistro 
         Enabled         =   0   'False
         Height          =   480
         Left            =   60
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConfiguracionDesglosaC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Guardar el registro"
         Top             =   140
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmConfiguracionDesglosaC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsConcepto As New ADODB.Recordset
Dim colSeleccion As Integer
Dim colClave As Integer
Dim colDescripcion As Integer
Dim colImpuesto As Integer
Dim colAsterisco As Integer
Dim colAsteriscoCargo As Integer
Dim queryConceptos As String
Dim cont As Integer
Private Type CustomType
    cveConcepto As Long
    bitagrupa As Integer
End Type

Private Sub cboEmpresa_Click()
    If cboEmpresa.ItemData(cboEmpresa.ListIndex) < 0 Then
        Call MsgBox("Por favor selecciona una empresa de convenio para cargar los conceptos de facturación.", vbInformation, "Mensaje")
        If cboEmpresaContable.Enabled = True Then
            cboEmpresaContable.SetFocus
        End If
        Exit Sub
    End If
    CargarConceptosFacturacion
    VSFlexGrid1.FixedRows = 1
    SeleccionarConceptosDesglosables
    VSFlexGrid1.Col = colDescripcion
    VSFlexGrid1.Row = 1
End Sub

Private Sub CboEmpresa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cboEmpresa.ListIndex > -1 Then
            If cboEmpresa.ItemData(cboEmpresa.ListIndex) < 0 Then
                Call MsgBox("Por favor selecciona una empresa de convenio para cargar los conceptos de facturación.", vbInformation, "Mensaje")
                If cboEmpresaContable.Enabled = True Then
                    cboEmpresaContable.SetFocus
                End If
                Exit Sub
            End If
            CargarConceptosFacturacion
            VSFlexGrid1.FixedRows = 1
            SeleccionarConceptosDesglosables
            VSFlexGrid1.Row = 1
            If VSFlexGrid1.Rows > 1 Then
                VSFlexGrid1.Col = colDescripcion
                VSFlexGrid1.SetFocus
            End If
        Else
            Call MsgBox("Por favor selecciona una empresa de convenio para cargar los conceptos de facturación.", vbInformation, "Mensaje")
        End If

    End If
    If KeyCode = 27 Then
        Unload Me
        KeyAscii = 0
    End If
End Sub

Private Sub cboEmpresaContable_Click()
    If cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) < 0 Then
        Call MsgBox("Por favor selecciona una empresa contable para cargar los conceptos de facturación.", vbInformation, "Mensaje")
        Exit Sub
    End If
    If cboEmpresa.ListCount > 0 Then
        cboEmpresa.SetFocus
    End If
    
    VSFlexGrid1.Clear
    VSFlexGrid1.Rows = 0
    VSFlexGrid1.Cols = 0
    pLimpiaGridDesgloa
End Sub

Private Sub cboEmpresaContable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex) < 0 Then
        Call MsgBox("Por favor selecciona una empresa contable para cargar los conceptos de facturación.", vbInformation, "Mensaje")
        Exit Sub
    End If
    cboEmpresa.SetFocus
    VSFlexGrid1.Clear
    VSFlexGrid1.Rows = 0
    VSFlexGrid1.Cols = 0
    pLimpiaGridDesgloa
End If
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub cmdGrabarRegistro_Click()
    agregaConfiguracion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 On Error GoTo NotificaError
    
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    cont = 0
    colSeleccion = 0
    colClave = 1
    colDescripcion = 2
    colImpuesto = 3
    colAsterisco = 4
    colAsteriscoCargo = 5
    cargaEmpresaContable
End Sub
Private Sub CargarConceptosFacturacion()
    
    
    Dim rsConceptos As New ADODB.Recordset
    queryConceptos = "SELECT SMICVECONCEPTO Clave, CHRDESCRIPCION Concepto, SMYIVA Impuesto FROM PVCONCEPTOFACTURACION WHERE INTTIPO = 0 order by CHRDESCRIPCION asc"
    Set rsConceptos = frsRegresaRs(queryConceptos, adLockOptimistic, adOpenStatic)
    
    
    VSFlexGrid1.Rows = rsConceptos.RecordCount + 1
    VSFlexGrid1.Cols = rsConceptos.Fields.Count + 3

    VSFlexGrid1.ColWidth(colSeleccion) = 300
    VSFlexGrid1.ColWidth(colClave) = 0
    VSFlexGrid1.ColWidth(colDescripcion) = 3000
    VSFlexGrid1.ColWidth(colImpuesto) = 800
    VSFlexGrid1.ColWidth(colAsterisco) = 2500
    VSFlexGrid1.ColAlignment(colAsterisco) = flexAlignCenterCenter
    VSFlexGrid1.ColWidth(colAsteriscoCargo) = 800
    VSFlexGrid1.ColAlignment(colAsteriscoCargo) = flexAlignCenterCenter
    
    Dim i As Long
    For i = 0 To rsConceptos.Fields.Count - 1
        VSFlexGrid1.TextMatrix(0, i + 1) = rsConceptos.Fields(i).Name
        
    Next i
    VSFlexGrid1.TextMatrix(0, colClave) = "Clave"
    VSFlexGrid1.TextMatrix(0, colDescripcion) = "Descripción"
    VSFlexGrid1.TextMatrix(0, colImpuesto) = "Impuesto"
    VSFlexGrid1.TextMatrix(0, colAsterisco) = "Agrupar por concepto de factura"
    VSFlexGrid1.TextMatrix(0, colAsteriscoCargo) = "Agrupar por cargo"
    
    i = 1
    While Not rsConceptos.EOF
        VSFlexGrid1.TextMatrix(i, 1) = rsConceptos("Clave").Value
        VSFlexGrid1.TextMatrix(i, 2) = rsConceptos("Concepto").Value
        VSFlexGrid1.TextMatrix(i, 3) = rsConceptos("Impuesto").Value & "%"
        i = i + 1
        rsConceptos.MoveNext
    Wend

    rsConceptos.Close
    Set rsConceptos = Nothing
End Sub
Private Sub SeleccionaColumna()
    Dim filaSeleccionada As Long
    filaSeleccionada = VSFlexGrid1.RowSel
    cmdGrabarRegistro.Enabled = True
    'Verificar si la fila seleccionada ya tiene un asterisco'
    If VSFlexGrid1.Col = colAsterisco Then
        If VSFlexGrid1.TextMatrix(filaSeleccionada, colAsterisco) = "*" Then
            VSFlexGrid1.TextMatrix(filaSeleccionada, colAsterisco) = ""
            VSFlexGrid1.Row = filaSeleccionada
            VSFlexGrid1.Col = colClave
            VSFlexGrid1.CellFontBold = False
            VSFlexGrid1.Col = colDescripcion
            VSFlexGrid1.CellFontBold = False
            VSFlexGrid1.Col = colImpuesto
            VSFlexGrid1.CellFontBold = False
            VSFlexGrid1.Col = colAsterisco
            VSFlexGrid1.SetFocus
            VSFlexGrid1.CellFontBold = False
            
            
        Else
            'Agregar asterisco en la fila seleccionada'
            If VSFlexGrid1.TextMatrix(filaSeleccionada, colAsteriscoCargo) = "*" Then
                VSFlexGrid1.TextMatrix(filaSeleccionada, colAsteriscoCargo) = ""
            End If
            VSFlexGrid1.TextMatrix(filaSeleccionada, colAsterisco) = "*"
            VSFlexGrid1.Row = filaSeleccionada
            VSFlexGrid1.Col = colClave
            VSFlexGrid1.CellFontBold = True
            VSFlexGrid1.Col = colDescripcion
            VSFlexGrid1.CellFontBold = True
            VSFlexGrid1.Col = colImpuesto
            VSFlexGrid1.CellFontBold = True
            
            VSFlexGrid1.Col = colAsterisco
            VSFlexGrid1.CellFontBold = True
            
            
            
        End If
    ElseIf VSFlexGrid1.Col = colAsteriscoCargo Then
        If VSFlexGrid1.TextMatrix(filaSeleccionada, colAsteriscoCargo) = "*" Then
            VSFlexGrid1.TextMatrix(filaSeleccionada, colAsteriscoCargo) = ""
            VSFlexGrid1.Row = filaSeleccionada
            VSFlexGrid1.Col = colClave
            VSFlexGrid1.CellFontBold = False
            VSFlexGrid1.Col = colDescripcion
            VSFlexGrid1.CellFontBold = False
            VSFlexGrid1.Col = colImpuesto
            VSFlexGrid1.CellFontBold = False
            
            VSFlexGrid1.Col = colAsteriscoCargo
            VSFlexGrid1.CellFontBold = False
            
             
        Else
            'Agregar asterisco en la fila seleccionada'
            If VSFlexGrid1.TextMatrix(filaSeleccionada, colAsterisco) = "*" Then
                VSFlexGrid1.TextMatrix(filaSeleccionada, colAsterisco) = ""
            End If
            VSFlexGrid1.TextMatrix(filaSeleccionada, colAsteriscoCargo) = "*"
            VSFlexGrid1.Row = filaSeleccionada
            VSFlexGrid1.Col = colClave
            VSFlexGrid1.CellFontBold = True
            VSFlexGrid1.Col = colDescripcion
            VSFlexGrid1.CellFontBold = True
            VSFlexGrid1.Col = colImpuesto
            VSFlexGrid1.CellFontBold = True
            
            VSFlexGrid1.Col = colAsteriscoCargo
            VSFlexGrid1.CellFontBold = True
            
            
            
        End If
    
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdGrabarRegistro.Enabled Then
        Cancel = True
        ' ¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
          LimpiarFormulario
          Cancel = True
          Exit Sub
        Else
          
        End If
        
    Else
        If cont > 0 Then
            Unload Me
        Else
            cont = cont + 1
            Cancel = True
        End If
    End If
    
    
End Sub

Private Sub VSFlexGrid1_DblClick()
With VSFlexGrid1
    If .Rows > 1 Then
        If .RowSel <> -1 Then
            SeleccionaColumna
        End If
    End If
End With
End Sub
Public Sub cargaEmpresaContable()
    Dim strQuery As String
    Dim rsCbo As New ADODB.Recordset
    Dim vlintnumero As Long
    
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
    
    cboEmpresaContable.ListIndex = fintLocalizaCbo(cboEmpresaContable, CStr(vgintClaveEmpresaContable)) 'se posiciona en la empresa con el que se dio login
    vlintnumero = IIf(cgstrModulo = "SI", 7032, 7031)
    
    If fblnRevisaPermiso(vglngNumeroLogin, vlintnumero, "C") Then
        cboEmpresaContable.Enabled = True
    Else
        cboEmpresaContable.Enabled = False
    End If
    
    strQuery = "SELECT * FROM CCEMPRESA WHERE bitActivo <> 0 ORDER BY vchDescripcion"
    Set rsCbo = frsRegresaRs(strQuery, adLockOptimistic, adOpenDynamic)
    pLlenarCboRs cboEmpresa, rsCbo, 0, 1
    cboEmpresa.ListIndex = -1
End Sub
Public Sub agregaConfiguracion()
    Dim Fila As Integer
    Dim claveEmpresa As Integer
    Dim rs As New ADODB.Recordset
    Dim claveConcepto As Long
    Dim dblAvance As Double
    Dim vllngPersonaGraba As Long
    Dim bitagrupa As Integer
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        If cboEmpresaContable.ListCount > 0 Then
            claveEmpresa = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
            If claveEmpresa = -1 Then
                Exit Sub
            End If
        Else
            Call MsgBox("No se encuentra seleccionada una empresa contable para guardar la configuración seleccionada.", vbExclamation, "Mensaje")
        End If
        lblTextoBarra.Caption = "Gurdando configuración..."
        freBarra.Visible = True
        pgbBarra.Value = 0
        freBarra.Refresh
        pgbBarra.Max = VSFlexGrid1.Rows - 1
        For Fila = 1 To VSFlexGrid1.Rows - 1
            If VSFlexGrid1.TextMatrix(Fila, colAsterisco) = "*" Then
                bitagrupa = 1
            ElseIf VSFlexGrid1.TextMatrix(Fila, colAsteriscoCargo) = "*" Then
                bitagrupa = -1
            Else
                bitagrupa = 0
            End If
            
            If bitagrupa <> 0 Then
                ' Obtener el valor de la clave del concepto seleccionado en el VSFlexGrid1
                claveConcepto = CInt(VSFlexGrid1.TextMatrix(Fila, colClave))
                
                ' Verificar si el registro ya existe en la tabla PVDESGLOSACONCEPTO
                Set rs = frsRegresaRs("SELECT 1 FROM PVDESGLOSACONCEPTO WHERE INTEMPRESACONTABLE = " & claveEmpresa & " AND SMICVECONCEPTO = " & claveConcepto & " AND INTCONVENIO = " & cboEmpresa.ItemData(cboEmpresa.ListIndex))
                If rs.EOF Then
                    ' Insertar el registro en la tabla PVDESGLOSACONCEPTO
                    pEjecutaSentencia "INSERT INTO PVDESGLOSACONCEPTO (INTEMPRESACONTABLE, SMICVECONCEPTO, BITAGRUPA, INTCONVENIO) VALUES (" & claveEmpresa & ", " & claveConcepto & "," & bitagrupa & ", " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & ")"
                    pGuardarLogTransaccion "FRMCONFIGURACIONDESGLOSAC", EnmGrabar, vglngNumeroLogin, "GUARDA CONFIGURACION DESGLOSA CONCEPTOS FACTURA MIXTA", "0 - " & claveEmpresa & " - " & claveConcepto & " - " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
                Else
                    ' Actualizar el valor del campo BITDESGLOSA a 1
                    pEjecutaSentencia "UPDATE PVDESGLOSACONCEPTO SET BITAGRUPA = " & bitagrupa & " WHERE INTEMPRESACONTABLE = " & claveEmpresa & " AND SMICVECONCEPTO = " & claveConcepto & " AND INTCONVENIO = " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
                    pGuardarLogTransaccion "FRMCONFIGURACIONDESGLOSAC", EnmCambiar, vglngNumeroLogin, "ACTUALIZA CONFIGURACION DESGLOSA CONCEPTOS FACTURA MIXTA", "0 - " & claveEmpresa & " - " & claveConcepto & " - " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
                End If
                rs.Close
                            
                Set rs = Nothing
            ElseIf VSFlexGrid1.TextMatrix(Fila, colAsterisco) = "" Then
                ' Obtener el valor de la clave del concepto seleccionado en el VSFlexGrid1
                claveConcepto = CInt(VSFlexGrid1.TextMatrix(Fila, colClave))
                        
                ' Verificar si el registro ya existe en la tabla PVDESGLOSACONCEPTO
                Set rs = frsRegresaRs("SELECT 1 FROM PVDESGLOSACONCEPTO WHERE INTEMPRESACONTABLE = " & claveEmpresa & " AND SMICVECONCEPTO = " & claveConcepto & " AND INTCONVENIO = " & cboEmpresa.ItemData(cboEmpresa.ListIndex))
                If Not rs.EOF Then
                    ' Actualizar el valor del campo BITDESGLOSA a 0
                    pEjecutaSentencia "UPDATE PVDESGLOSACONCEPTO SET BITAGRUPA = 0 WHERE INTEMPRESACONTABLE = " & claveEmpresa & " AND SMICVECONCEPTO = " & claveConcepto & " AND INTCONVENIO = " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
                    pGuardarLogTransaccion "FRMCONFIGURACIONDESGLOSAC", EnmCambiar, vglngNumeroLogin, "ACTUALIZA CONFIGURACION DESGLOSA CONCEPTOS FACTURA MIXTA", "0 - " & claveEmpresa & " - " & claveConcepto & " - " & cboEmpresa.ItemData(cboEmpresa.ListIndex)
                End If
                rs.Close
                                    
                Set rs = Nothing
           End If
            pgbBarra.Value = pgbBarra.Value + 1
        Next Fila
        
        freBarra.Visible = False
        pgbBarra.Max = 100
        pgbBarra.Value = 0
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        LimpiarFormulario
    End If
End Sub
Private Sub SeleccionarConceptosDesglosables()
    Dim rsDesglosa As New ADODB.Recordset
    Dim queryDesglosa As String
    Dim clavesDesglosa() As Long
    Dim clavesPorCargo() As Long
    Dim i As Long, j As Long
    
    
    pLimpiaGridDesgloa
    claveEmpresa = cboEmpresaContable.ItemData(cboEmpresaContable.ListIndex)
    
    ' Obtener los valores de SMICVECONCEPTO que tienen BITDESGLOSA = 1
    queryDesglosa = "SELECT SMICVECONCEPTO,BITAGRUPA FROM PVDESGLOSACONCEPTO WHERE BITAGRUPA <> 0 and INTCONVENIO = " & cboEmpresa.ItemData(cboEmpresa.ListIndex) & " and INTEMPRESACONTABLE =  " & claveEmpresa

    Set rsDesglosa = frsRegresaRs(queryDesglosa, adLockOptimistic, adOpenStatic)
    If rsDesglosa.RecordCount > 0 Then
        ReDim clavesDesglosa(1 To rsDesglosa.RecordCount)
        ReDim clavesPorCargo(1 To rsDesglosa.RecordCount)
        j = 1
        While Not rsDesglosa.EOF
            If rsDesglosa!bitagrupa = 1 Then
                clavesDesglosa(j) = rsDesglosa("SMICVECONCEPTO").Value
            Else
                clavesPorCargo(j) = rsDesglosa("SMICVECONCEPTO").Value
            End If
            j = j + 1
            rsDesglosa.MoveNext
        Wend
        rsDesglosa.Close
        
        ' Recorrer las filas del VSFlexGrid y marcar con asterisco las filas que corresponden a claves desglosables
        For i = 1 To VSFlexGrid1.Rows - 1
            If IsInArray(CInt(VSFlexGrid1.TextMatrix(i, colClave)), clavesDesglosa) Then
                VSFlexGrid1.TextMatrix(i, colAsterisco) = "*"
            End If
            If IsInArray(CInt(VSFlexGrid1.TextMatrix(i, colClave)), clavesPorCargo) Then
                VSFlexGrid1.TextMatrix(i, colAsteriscoCargo) = "*"
            End If
        Next i
        
    End If
End Sub

Private Function IsInArray(ByVal Value As Long, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = Value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
Private Sub LimpiarFormulario()
    ' Limpiar el VSFlexGrid1
    VSFlexGrid1.Clear
    VSFlexGrid1.Rows = 0
    VSFlexGrid1.Cols = 0
    
    ' Establecer el ComboBox cboEmpresaContable a su valor por defecto
    cmdGrabarRegistro.Enabled = False
    If cboEmpresaContable.Enabled = True Then
        cboEmpresaContable.SetFocus
    End If
End Sub


Private Sub VSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If VSFlexGrid1.Rows > 1 Then
            SeleccionaColumna
        End If
    End If
End Sub
Private Function pLimpiaGridDesgloa()
     For i = 1 To VSFlexGrid1.Rows - 1
        VSFlexGrid1.TextMatrix(i, colAsterisco) = ""
        VSFlexGrid1.TextMatrix(i, colAsteriscoCargo) = ""
        VSFlexGrid1.Row = i
        VSFlexGrid1.Col = colAsterisco
        VSFlexGrid1.CellFontBold = False
        VSFlexGrid1.Col = colAsteriscoCargo
        VSFlexGrid1.CellFontBold = False
        VSFlexGrid1.Col = colClave
        VSFlexGrid1.CellFontBold = False
        VSFlexGrid1.Col = colDescripcion
        VSFlexGrid1.CellFontBold = False
        VSFlexGrid1.Col = colImpuesto
        VSFlexGrid1.CellFontBold = False
    Next i
End Function
