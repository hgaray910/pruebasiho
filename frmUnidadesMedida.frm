VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmUnidadesMedida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unidades de medida para conceptos de factura"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1000
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      Begin VB.ComboBox cboTipoConcepto 
         Height          =   315
         ItemData        =   "frmUnidadesMedida.frx":0000
         Left            =   150
         List            =   "frmUnidadesMedida.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   550
         Width           =   2950
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de conceptos"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   250
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6180
      Begin VB.ComboBox cboTipoReferencia 
         Height          =   315
         ItemData        =   "frmUnidadesMedida.frx":0004
         Left            =   150
         List            =   "frmUnidadesMedida.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa o el tipo de paciente."
         Top             =   550
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa / Tipo de paciente"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   250
         Width           =   1980
      End
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   4560
      TabIndex        =   6
      Top             =   6360
      Width           =   600
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   52
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmUnidadesMedida.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Grabar"
         Top             =   150
         Width           =   495
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdUnidades 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9495
      _cx             =   16748
      _cy             =   8705
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      Editable        =   1
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
   Begin VB.Label lblCount 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
End
Attribute VB_Name = "frmUnidadesMedida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------'
'| Nombre del Proyecto      : Caja                                                  |'
'| Nombre del Formulario    : frmUnidadesMedida                                     |'
'------------------------------------------------------------------------------------'
'| Objetivo: Capturar unidades de medida para conceptos de facturación              |'
'------------------------------------------------------------------------------------'
'| Análisis y Diseño        : Claudia I. Ruvalcaba                                  |'
'| Autor                    : Claudia I. Ruvalcaba                                  |'
'| Fecha de Creación        : 14/Nov/2012                                           |'
'| Modificó                 : Claudia I. Ruvalcaba                                  |'
'| Fecha Terminación        : 11/Dic/2012                                           |'
'| Fecha última modificación: 04/Ene/2013                                           |3
'------------------------------------------------------------------------------------'

Option Explicit

Public vllngNumeroOpcion As Long

'- Columnas del grid de unidades -'
Const lintColFixed = 0
Const lintColCveConcepto = 1
Const lintColConcepto = 2
Const lintColCveUnidad = 3
Const lintColUnidad = 4
Const lintColTipo = 5
Const lintColReferencia = 6
Const lintColTipoConcepto = 7
'---------------------------------'

Dim rsTipos As New ADODB.Recordset
Dim rsUnidades As New ADODB.Recordset

Dim vlstrsql As String

'- Variables para revisar si se ha modificado el grid -'
Dim lstrAnterior As String 'Valor anterior de la celda que se ha seleccionado
Dim lblnSalir As Boolean

'- Función para traer el tipo de referencia con la que se está trabajando: -1 = Todos, 0 = Tipos de Pacientes, 1 = Empresas -'
Private Function fintTraeTipo() As Integer
On Error GoTo NotificaError

    fintTraeTipo = -2 'Inicializar en error el valor de la función (-1 está reservado para el tipo Todos)

    If cboTipoReferencia.ListIndex = 0 Then
        fintTraeTipo = -1  'Devuelve >>Todos<< para el primer índice del combo
    Else
        'Localiza el registro en el recordset y devuelve el valor del campo >>Tipo<<
        If fintLocalizaPkRs(rsTipos, 1, cboTipoReferencia.Text) <> 0 Then
            fintTraeTipo = rsTipos.Fields("Tipo").Value
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintTraeTipo"))
    Unload Me
End Function

'- Configura el grid de Unidades -'
Public Sub pConfiguraGrid()
On Error GoTo NotificaError

    Dim lstrTitulo As String
    
    Select Case cboTipoConcepto.ListIndex
        Case 0, 1
            lstrTitulo = "Descripción"
        Case Else
            lstrTitulo = "Nombre"
    End Select
    
    With grdUnidades
        .Clear
        .Rows = 1
        .Cols = 8
        .FixedCols = 1
        .FixedRows = 1
        
        .FormatString = "||" & lstrTitulo & "||Unidad de medida|"
        
        .ColHidden(lintColCveConcepto) = True    'Clave del concepto de facturación, otro concepto, estudio o examen
        .ColHidden(lintColConcepto) = False      'Descripción del concepto de facturación, otro concepto, estudio o examen
        .ColHidden(lintColCveUnidad) = True      'Clave de la unidad de medida
        .ColHidden(lintColUnidad) = False        'Descripción de la unidad de medida
        .ColHidden(lintColTipo) = True           'Tipo de referencia de la unidad de medida
        .ColHidden(lintColReferencia) = True     'Clave de la referencia de la unidad de medida
        .ColHidden(lintColTipoConcepto) = True   'Tipo de concepto: NO, OC, ES, EX, GE.
        
        .ColWidth(lintColFixed) = 100
        .ColWidth(lintColConcepto) = 4950
        .ColWidth(lintColUnidad) = 4110
        
        .ColAlignment(lintColConcepto) = flexAlignLeftCenter
        .ColAlignment(lintColUnidad) = flexAlignLeftCenter
        
        .EditMaxLength = 150 'Longitud máxima de la descripción de la unidad

        .FixedAlignment(lintColConcepto) = flexAlignCenterCenter
        .FixedAlignment(lintColUnidad) = flexAlignCenterCenter
        
        .ScrollBars = flexScrollBarBoth
    End With
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
    Unload Me
End Sub

'- Carga los conceptos de facturación y las unidades de medida relacionadas a estos -'
Private Sub pCargaConceptos()
On Error GoTo NotificaError

    Dim rsConceptos As ADODB.Recordset
    
    If cboTipoReferencia.ListCount = 0 Or cboTipoConcepto.ListIndex = -1 Then Exit Sub

    grdUnidades.Redraw = False
    grdUnidades.Visible = False
    
    vgstrParametrosSP = fintTraeTipo & "|" & _
                        cboTipoReferencia.ItemData(cboTipoReferencia.ListIndex) & "|-1|" & _
                        cboTipoConcepto.ListIndex & "|" & _
                        vgintNumeroDepartamento
    Set rsConceptos = frsEjecuta_SP(vgstrParametrosSP, "Sp_PvSelUnidadMedida")
    If rsConceptos.RecordCount > 0 Then
        pConfiguraGrid
    
        With rsConceptos
            .MoveFirst
            Do While Not .EOF
                grdUnidades.Rows = grdUnidades.Rows + 1
                grdUnidades.Row = grdUnidades.Rows - 1
                
                grdUnidades.TextMatrix(grdUnidades.Row, lintColCveConcepto) = !CveConcepto
                grdUnidades.TextMatrix(grdUnidades.Row, lintColConcepto) = !Concepto
                grdUnidades.TextMatrix(grdUnidades.Row, lintColCveUnidad) = !CveUnidad
                grdUnidades.TextMatrix(grdUnidades.Row, lintColUnidad) = IIf(IsNull(!Unidad), "", !Unidad)
                grdUnidades.TextMatrix(grdUnidades.Row, lintColTipo) = !Tipo
                grdUnidades.TextMatrix(grdUnidades.Row, lintColReferencia) = !Referencia
                grdUnidades.TextMatrix(grdUnidades.Row, lintColTipoConcepto) = !TipoConcepto
                
                .MoveNext
            Loop
        End With
        
        'lblCount.Caption = "Total registros: " & CStr(rsConceptos.RecordCount)
    End If
    rsConceptos.Close
    
    grdUnidades.Redraw = True
    grdUnidades.Visible = True
    
    cmdSave.Enabled = False 'Volver al estado inicial de la forma para que se capturen mas datos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaConceptos"))
    Unload Me
End Sub

'- Carga los listados de Empresas y de Tipos de Pacientes -'
Private Sub pCargaTipos()
    vlstrsql = "SELECT INTCVEEMPRESA AS Clave, VCHDESCRIPCION AS Descripcion, 1 AS Tipo FROM CcEmpresa WHERE bitActivo <> 0 " & _
               " UNION " & _
               "SELECT TNYCVETIPOPACIENTE AS Clave, VCHDESCRIPCION AS Descripcion, 0 AS Tipo FROM ADTIPOPACIENTE WHERE bitActivo <> 0 " & _
               " ORDER BY 3 DESC, 2 ASC"
    Set rsTipos = frsRegresaRs(vlstrsql)
    If Not rsTipos.EOF Then
        pLlenarCboRs cboTipoReferencia, rsTipos, 0, 1, 3
        cboTipoReferencia.ListIndex = 0 'Por defecto en el tipo <TODOS>
        cboTipoConcepto.ListIndex = 0   'Por defecto en conceptos de facturación
    End If
End Sub

Private Sub pCargaTiposConcepto()
    With cboTipoConcepto
        .Clear
        .AddItem "CONCEPTOS DE FACTURACIÓN", 0
        '- Validación para que no aparezcan los otros tipos de concepto en el Hospital San José de Hermosillo -'
        If Replace(Replace(UCase(vgstrRfCCH), "-", ""), " ", "") <> "HSJ040622G70" Then
            .AddItem "OTROS CONCEPTOS", 1
            .AddItem "ESTUDIOS", 2
            .AddItem "EXÁMENES", 3
            .AddItem "GRUPO DE EXÁMENES", 4
        End If
    End With
End Sub

Private Sub cboTipoConcepto_Click()
    pCargaConceptos
End Sub

Private Sub cboTipoConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grdUnidades.Rows > 1 Then
            grdUnidades.Col = lintColUnidad
            grdUnidades.Row = 1
            grdUnidades.SetFocus
        End If
    End If
End Sub

Private Sub cboTipoReferencia_Click()
    pCargaConceptos
End Sub

Private Sub cboTipoReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cboTipoConcepto.SetFocus
End Sub

'- Guardar la información -'
Private Sub cmdSave_Click()
On Error GoTo NotificaError

    Dim vllngContador As Long
    Dim vllngPersonaGraba As Long
    Dim vlstrTipo As String
    
    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcion, "E") Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            'Traer el tipo de Paciente/Empresa
            Select Case fintTraeTipo
                Case 0
                    vlstrTipo = "PACIENTE"
                Case 1
                    vlstrTipo = "EMPRESA"
                Case Else
                    vlstrTipo = "TODOS"
            End Select
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            With grdUnidades
                For vllngContador = 1 To .Rows - 1
                    vlstrsql = "SELECT * FROM PvUnidadMedida WHERE intCveUnidad = " & Val(Trim(.TextMatrix(vllngContador, lintColCveUnidad))) & _
                               " AND intTipo = " & Val(Trim(.TextMatrix(vllngContador, lintColTipo))) & _
                               " AND chrTipoConcepto = '" & Trim(.TextMatrix(vllngContador, lintColTipoConcepto)) & "'"
                    Set rsUnidades = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
                    If Trim(.TextMatrix(vllngContador, lintColUnidad)) <> "" Then 'Existe una descripción capturada
                        If rsUnidades.RecordCount = 0 Then
                            rsUnidades.AddNew 'Es un nuevo registro
                        End If
                        rsUnidades!intTipo = Val(Trim(.TextMatrix(vllngContador, lintColTipo)))
                        rsUnidades!intCveReferencia = Val(Trim(.TextMatrix(vllngContador, lintColReferencia)))
                        rsUnidades!smiCveConcepto = Val(Trim(.TextMatrix(vllngContador, lintColCveConcepto)))
                        rsUnidades!vchDescripcion = Trim(.TextMatrix(vllngContador, lintColUnidad))
                        rsUnidades!chrTipoConcepto = Trim(.TextMatrix(vllngContador, lintColTipoConcepto))
                        rsUnidades.Update
                    Else
                        If rsUnidades.RecordCount <> 0 Then 'Si existe la unidad pero se ha borrado la descripción, se borra el registro
                            rsUnidades.Delete
                        End If
                    End If
                    
                    rsUnidades.Close
                Next
            End With
            
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "UNIDADES DE MEDIDA " & vlstrTipo, cboTipoReferencia.ItemData(cboTipoReferencia.ListIndex))
        
            EntornoSIHO.ConeccionSIHO.CommitTrans

            'La operación se realizó satisfactoriamente
            MsgBox SIHOMsg(420), vbOKOnly + vbInformation, "Mensaje"
            
            pCargaTipos 'Recargar los tipos para actualizar la información
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    cmdSave.Enabled = True
    Unload Me
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon

    pCargaTiposConcepto
    pCargaTipos
    
    lblnSalir = True 'Al iniciar la forma permitir salir sin limpiar los datos
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lblnSalir Then
        If cmdSave.Enabled Then
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                pCargaTipos 'Limpia la forma
            End If
            Cancel = True
            lblnSalir = False
        Else
            If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") <> vbYes Then
                cboTipoReferencia.SetFocus
                Cancel = True
                lblnSalir = False
            End If
        End If
    Else
        Cancel = True
        lblnSalir = True
    End If
End Sub

Private Sub grdUnidades_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim vlintTipo As Long
    Dim vlstrUnidad As String

    If Col = lintColUnidad Then
        vlstrUnidad = Replace(grdUnidades.TextMatrix(Row, Col), vbCrLf, "")
        grdUnidades.TextMatrix(Row, Col) = Trim(vlstrUnidad)
    
        vlintTipo = fintTraeTipo
        If vlintTipo = -2 Then
            grdUnidades.TextMatrix(Row, lintColTipo) = -1
            grdUnidades.TextMatrix(Row, lintColReferencia) = 0
        Else
            grdUnidades.TextMatrix(Row, lintColTipo) = vlintTipo
            grdUnidades.TextMatrix(Row, lintColReferencia) = cboTipoReferencia.ItemData(cboTipoReferencia.ListIndex)
        End If
        
        If lstrAnterior <> Trim(grdUnidades.TextMatrix(Row, Col)) And Not cmdSave.Enabled Then
            cmdSave.Enabled = True 'Permitir guardar si hubo cambios
        End If
    End If
End Sub

Private Sub grdUnidades_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> lintColUnidad Then
        Cancel = True
    Else
        lstrAnterior = Trim(grdUnidades.TextMatrix(Row, Col))
    End If
End Sub
