VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Object = "{FF14BD24-9F8A-41E3-B5B8-7F0D45EE9F16}#15.0#0"; "HSFlatControls.ocx"
Begin VB.Form frmListaPaciente 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda por nombre comercial"
   ClientHeight    =   5355
   ClientLeft      =   5925
   ClientTop       =   5430
   ClientWidth     =   13905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " Nombre comercial completo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   75
      TabIndex        =   3
      Top             =   4300
      Width           =   12600
      Begin VB.TextBox txtDescripcionCompleta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   300
         Width           =   12375
      End
   End
   Begin HSFlatControls.MyCombo cboManejoMedicamentos 
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Enabled         =   -1  'True
      Text            =   ""
      Sorted          =   0   'False
      List            =   ""
      ItemData        =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyCommandButton.MyButton cmdManejos 
      Height          =   375
      Left            =   12720
      TabIndex        =   2
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Manejos"
      CaptionPosition =   4
      DepthEvent      =   1
      ForeColorDisabled=   -2147483629
      PictureAlignment=   4
      ShowFocus       =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHLista 
      DragIcon        =   "frmListaPaciente.frx":0000
      Height          =   4185
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Artículos o medicamentos"
      Top             =   120
      Width           =   13770
      _ExtentX        =   24289
      _ExtentY        =   7382
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   16777215
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      GridColorUnpopulated=   -2147483638
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      MergeCells      =   1
      Appearance      =   0
      FormatString    =   "|tnyCvePiso|vchDescripcion"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "frmListaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Inventarios
'| Nombre del Formulario    : frmLista
'-------------------------------------------------------------------------------------
'| Objetivo: Muestra una en un grid el resultado de una busqueda
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Luis Astudillo - Ines Salais
'| Autor                    : Luis Astudillo - Ines Salais
'| Fecha de Creación        : 14/Abril/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------
Option Explicit
Public vgstrCriterioParam As String
Public vgblnNombreGenerico As Integer
Private objRS As New ADODB.Recordset
Public llngCuentaPaciente As Long
Public lstrTipoPaciente As String
Public lblnCuadroBasico As Boolean
Public lblnAutorizacionMedicamento As Boolean
Dim llngCveAutoriza As Long
Dim lstrTipoAutoriza As String
Dim lstrFecha As String
Dim strColorCuadroBasico As String
Dim strColorExcluido As String
Dim strColorBasicoExcluido As String
Dim strFondoCuadroBasico As String
Dim strFondoExcluido As String
Dim strFondoBasicoExcluido As String

' Variables para configurar el grid '
Dim lintColFixed As Integer           ' Fixed row
Dim lintColCveArticulo As Integer     ' Clave del artículo
Dim lintColNombreComercial As Integer ' Nombre comercial
Dim lintColNombreGenerico As Integer  ' Nombre genérico
Dim lintColExistenciaUV As Integer    ' Existencia UV
Dim lintColExistenciaUM As Integer    ' Existencia UM
Dim lintColPrecioUM As Integer        ' Precio unidad mínima
Dim lintColPrecioUA As Integer        ' Precio unidad alterna
Dim lintColFamilia As Integer         ' Familia
Dim lintColSubfamilia As Integer      ' Subfamilia
Dim lintColCodigoBarras As Integer    ' Código de barras
Dim lintColExcluido As Integer        ' Indica si es excluido 0, 1 es excluido
Dim lintColCuadroBasico As Integer    ' Indica si es un medicamento que pertenece al cuadro básico 0 no pertenece, 1 si pertenece
Dim lintColTipo As Integer            ' indica si es 0=Articulo, 1=Medicamento, 2=Insumo

Dim lintTotalManejos As Integer       'Variable que indica el total de manejos de medicamentos actualmente activos en el sistema

'Estructura para el manejo de medicamentos'
Dim rsManejoMedicamentos As New ADODB.Recordset
Private Type ManejoMedicamentos   'Para los colores del manejo de medicamentos
    intCveManejo As Integer
    strColor As String
    strSimbolo As String
End Type
Dim aManejoMedicamentos() As ManejoMedicamentos

Dim lblnFormaActivada As Boolean 'Variable para evitar que se ejecute el código del Form_Activate otra vez

Private Sub pMostrarDescripcionCompleta()
    txtDescripcionCompleta = grdHLista.TextMatrix(grdHLista.Row, lintColNombreComercial)
End Sub

Private Sub cmdManejos_Click()
On Error GoTo NotificaError
    frmManejoMedicamentos.blnSoloBusqueda = True
    frmManejoMedicamentos.Show vbModal, Me
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdManejo_Click"))
End Sub

Private Sub Form_Activate()

    Dim vlstrTitulos As String
    Dim vlintCnt As Integer
    
    If lblnFormaActivada Then Exit Sub

    Me.Caption = IIf(vgblnNombreGenerico, "Búsqueda por nombre genérico", "Búsqueda por nombre comercial")
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    vlstrTitulos = ""           'Titulos de las columnas de los medicamentos
    lintTotalManejos = 0        'Número por defecto del total de manejo de medicamentos (0 = No se utilizan los manejos)

    Dim vlstrSentencia  As String
    vlstrSentencia = vgstrCriterioParam & "|" & IIf(vgblnNombreGenerico = 1, "1", IIf(vgblnNombreGenerico = 2, "2", "0")) & "|" & vgintCveDeptoCargo & "|" & vgintClaveEmpresaContable & "|" & vgstrVarIntercam & "|" & lstrTipoPaciente & "|" & llngCuentaPaciente
    Set objRS = frsEjecuta_SP(vlstrSentencia, "sp_EXSelListaPaciente")
    If objRS.RecordCount > 0 Then
        lintTotalManejos = fintConfigManejosMedicamentos(cboManejoMedicamentos)
        If lintTotalManejos > 0 Then
            For vlintCnt = 1 To lintTotalManejos
                vlstrTitulos = vlstrTitulos & "|"
            Next
        End If
            
        lintColCveArticulo = 1 + lintTotalManejos
        lintColNombreComercial = 2 + lintTotalManejos
        lintColNombreGenerico = 3 + lintTotalManejos
        lintColExistenciaUV = 4 + lintTotalManejos
        lintColExistenciaUM = 5 + lintTotalManejos
        lintColPrecioUM = 6 + lintTotalManejos
        lintColPrecioUA = 7 + lintTotalManejos
        lintColFamilia = 8 + lintTotalManejos
        lintColSubfamilia = 9 + lintTotalManejos
        lintColCodigoBarras = 10 + lintTotalManejos
        lintColExcluido = 11 + lintTotalManejos
        lintColCuadroBasico = 12 + lintTotalManejos
        lintColTipo = 13 + lintTotalManejos
        
        ''Call pLlenarMshFGrdRs(grdHLista, ObjRS)
        Call pLlenarMshFGrdRsManejos(grdHLista, objRS, lintTotalManejos, lintColCveArticulo)
        Call pConfGrid(vlstrTitulos & "|Clave|Nombre comercial|Nombre genérico|Alterna|Mínima|Precio unidad mínima|Precio unidad alterna|Familia|Subfamilia|Código")
        pIdentificaCargos
    End If
    objRS.Close
        
    If grdHLista.Enabled And grdHLista.Visible Then grdHLista.SetFocus
    Me.Refresh
    
    lblnFormaActivada = True
End Sub

Private Sub pIdentificaCargos()
    Dim llngContador As Long
    Dim llngRow As Long
    
    grdHLista.Redraw = False
    For llngRow = 1 To grdHLista.Rows - 1
        grdHLista.Row = llngRow
        ' si es un cargo excluido o si es un medicamento que pertenece al cuadro básico los identifica marcados con un color
        If lblnCuadroBasico Then
            If Val(grdHLista.TextMatrix(llngRow, lintColExcluido)) = 1 Or Val(grdHLista.TextMatrix(llngRow, lintColCuadroBasico)) = 1 Then
                For llngContador = 1 To grdHLista.Cols - 1
                    grdHLista.Col = llngContador
                    If Val(grdHLista.TextMatrix(llngRow, lintColExcluido)) = 1 And Val(grdHLista.TextMatrix(llngRow, lintColCuadroBasico)) = 1 Then
                    'Excluidos y cuadro básico
                        grdHLista.CellBackColor = strFondoBasicoExcluido
                        grdHLista.CellForeColor = strColorBasicoExcluido
                    ElseIf Val(grdHLista.TextMatrix(llngRow, lintColExcluido)) = 1 Then
                    'Excluido
                        grdHLista.CellBackColor = strFondoExcluido
                        grdHLista.CellForeColor = strColorExcluido
                    ElseIf Val(grdHLista.TextMatrix(llngRow, lintColCuadroBasico)) = 1 Then
                    'Cuadro básico
                        grdHLista.CellBackColor = strFondoCuadroBasico
                        grdHLista.CellForeColor = strColorCuadroBasico
                    End If
                Next
            End If
        Else
            If Val(grdHLista.TextMatrix(llngRow, lintColExcluido)) = "1" Then
                For llngContador = 1 To grdHLista.Cols - 1
                    grdHLista.Col = llngContador
                    grdHLista.CellBackColor = &H99FFFF
                    grdHLista.CellForeColor = &H80000012
                Next
            End If
        End If
        'Manejos de medicamentos
        pColorearManejo grdHLista, cboManejoMedicamentos, lintColFixed + 1, grdHLista.TextMatrix(llngRow, lintColCveArticulo), llngRow, 330
    Next llngRow
    grdHLista.Redraw = True
    grdHLista.Col = 1
    grdHLista.Row = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pSalirForm
    End Select
End Sub

Private Sub Form_Load()
    Dim rsColores As New ADODB.Recordset
    
    Me.Icon = frmMenuPrincipal.Icon
    
    'Color de letra para los medicamentos del cuadro básico
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHCOLORCUADROBASICO")
    If Not rsColores.EOF Then
        strColorCuadroBasico = IIf(IsNull(rsColores!valor), "0", rsColores!valor)
    Else
        strColorCuadroBasico = "0"
    End If
    rsColores.Close
    
    'Color de letra para los cargos excluídos
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHCOLOREXCLUIDO")
    If Not rsColores.EOF Then
        strColorExcluido = IIf(IsNull(rsColores!valor), "0", rsColores!valor)
    Else
        strColorExcluido = "0"
    End If
    rsColores.Close
    
    'Color de letra para los medicamentos del cuadro básico que son excluídos
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHCOLORBASICOEXCLUIDO")
    If Not rsColores.EOF Then
        strColorBasicoExcluido = IIf(IsNull(rsColores!valor), "0", rsColores!valor)
    Else
        strColorBasicoExcluido = "0"
    End If
    rsColores.Close
    
    'Color de fondo para los medicamentos del cuadro básico
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHFONDOCUADROBASICO")
    If Not rsColores.EOF Then
        strFondoCuadroBasico = IIf(IsNull(rsColores!valor), "9420794", rsColores!valor)
    Else
        strFondoCuadroBasico = "9420794"
    End If
    rsColores.Close
    
    'Color de fondo para los cargos excluídos
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHFONDOEXCLUIDO")
    If Not rsColores.EOF Then
        strFondoExcluido = IIf(IsNull(rsColores!valor), "10092543", rsColores!valor)
    Else
        strFondoExcluido = "10092543"
    End If
    rsColores.Close
    
    'Color de fondo para los medicamentos del cuadro básico que son excluídos
    Set rsColores = frsSelParametros("EX", vgintClaveEmpresaContable, "VCHFONDOBASICOEXCLUIDO")
    If Not rsColores.EOF Then
        strFondoBasicoExcluido = IIf(IsNull(rsColores!valor), "16777139", rsColores!valor)
    Else
        strFondoBasicoExcluido = "16777139"
    End If
    rsColores.Close
    
    lblnFormaActivada = False
End Sub

Private Sub grdHLista_Click()
'-------------------------------------------------------------------------------------------
' Refresca el grdEspxMedico y asigna bajo que columna se va a hacer la búsqueda
'-------------------------------------------------------------------------------------------
    If grdHLista.Rows > 0 Then
        pMostrarDescripcionCompleta
        vgintColLoc = grdHLista.Col
        vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
        grdHLista.Col = vgintColLoc
        grdHLista.Refresh
    End If
    
End Sub

Private Sub grdHLista_DblClick()
'-------------------------------------------------------------------------------------------
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    ' Ordena solamente cuando un encabezado de columna es seleccionado con un click
    If grdHLista.Rows > 0 Then
        If (grdHLista.Row <= grdHLista.Rows - 1) Then
            If grdHLista.MouseRow >= grdHLista.FixedRows Then
                grdHLista_KeyDown vbKeyReturn, 0
                Exit Sub
            Else
                vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
                vgintColOrd = grdHLista.Col  'Configura la columna a ordenar
                
                'Escoge el Tipo de Ordenamiento
                If vgintTipoOrd = 1 Then
                     vgintTipoOrd = 2
                Else
                    vgintTipoOrd = 1
                End If
                Call pOrdColMshFGrid(grdHLista, vgintTipoOrd)
                Call pDesSelMshFGrid(grdHLista)
                pIdentificaCargos
                Me.Refresh
            End If
        End If
    End If
    
End Sub

Private Sub grdHLista_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            llngCveAutoriza = 0
            lstrTipoAutoriza = ""
            lstrFecha = ""
            
            'si maneja cuadro básico y  no pertenece al cuadro básico, y si es medicamento
            If lblnCuadroBasico And Val(grdHLista.TextMatrix(grdHLista.Row, lintColCuadroBasico)) = 0 And Val(grdHLista.TextMatrix(grdHLista.Row, lintColTipo)) = 1 Then
                If lblnAutorizacionMedicamento Then
                    If fblnCuadroBasicoContinuar(llngCveAutoriza, lstrTipoAutoriza, lstrFecha) Then
                        frmRequisicionCargoPac.llngCvePersonaAutoriza = llngCveAutoriza
                        frmRequisicionCargoPac.lstrTipoPersonaAutoriza = lstrTipoAutoriza
                        frmRequisicionCargoPac.lstrFechaAutorizacion = lstrFecha
                    Else
                        Exit Sub
                    End If
                Else
                    If MsgBox(SIHOMsg(1047), vbQuestion + vbYesNo, "Mensaje") = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
            vgstrVarIntercam = grdHLista.TextMatrix(grdHLista.Row, lintColCveArticulo)
            vgstrVarIntercam2 = grdHLista.TextMatrix(grdHLista.Row, lintColNombreComercial)
            vgstrvarcodigodebarras = grdHLista.TextMatrix(grdHLista.Row, lintColCodigoBarras)
            frmRequisicionCargoPac.saSeleccion = saSeleccionado
            pSalirForm
    End Select
    
End Sub

Private Function fblnCuadroBasicoContinuar(llngCveAutoriza As Long, lstrTipoAutoriza As String, lstrFecha As String) As Boolean
On Error GoTo NotificaError
    
    fblnCuadroBasicoContinuar = False
    frmPersonalAutorizado.lintClaveProceso = 1
    frmPersonalAutorizado.Show vbModal, Me
    llngCveAutoriza = frmPersonalAutorizado.llngCvePersonaAutoriza
    lstrTipoAutoriza = frmPersonalAutorizado.lstrTipoPersonaAutoriza
    lstrFecha = frmPersonalAutorizado.lstrFechaAutorizacion
    Unload frmPersonalAutorizado
    fblnCuadroBasicoContinuar = llngCveAutoriza <> 0
    
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCuadroBasicoContinuar"))
End Function

Private Sub grdHLista_KeyPress(vlintKeyAscii As Integer)
'-------------------------------------------------------------------------------------------
' Evento que verifica si se presiono una tecla
' de la A-Z, a-z, 0-9, á,é,í,ó,ú,ñ,Ñ, se presiono la barra espaciadora
' Realizando la búsqueda de un criterio dentro del grdHBusqueda
'-------------------------------------------------------------------------------------------
    Dim llngRow As Long
    
    Call pSelCriterioMshFGrid(grdHLista, vgintColLoc, vlintKeyAscii)
    llngRow = grdHLista.Row
    pIdentificaCargos
    grdHLista.Row = llngRow
    grdHLista.Col = 2
    If vlintKeyAscii <> 13 Then grdHLista.SetFocus
    
End Sub

Private Sub pSalirForm()
'-------------------------------------------------------------------------------------------
' Cierra y limpia Recordsets, variables, Grid para el cierre del Form
'-------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub grdHLista_GotFocus()
    If grdHLista.Rows <= 1 Then
        vgstrVarIntercam = ""
        vgstrVarIntercam2 = ""
        vgstrvarcodigodebarras = ""
        pSalirForm
    End If
End Sub

Private Sub grdHLista_RowColChange()
    pMostrarDescripcionCompleta
End Sub

Private Sub pConfGrid(vlstrFormatString As String)

    Dim lintcnt As Integer
    
    With grdHLista
        .Redraw = False
        .FormatString = vlstrFormatString
        
        For lintcnt = lintColFixed + 1 To lintColCveArticulo - 1
            .ColWidth(lintcnt) = 0
        Next lintcnt
        
        .ColWidth(lintColFixed) = 300            '| Fixed row
        .ColWidth(lintColCveArticulo) = 1300     '| Clave del artículo
        .ColWidth(lintColNombreComercial) = 3500 '| Nombre comercial
        .ColWidth(lintColNombreGenerico) = 3500  '| Nombre genérico
        .ColWidth(lintColExistenciaUV) = 800     '| Existencia UV
        .ColWidth(lintColExistenciaUM) = 800     '| Existencia UM
        .ColWidth(lintColPrecioUM) = 1800        '| Precio unidad mínima
        .ColWidth(lintColPrecioUA) = 1800        '| Precio unidad alterna
        .ColWidth(lintColFamilia) = 1800         '| Familia
        .ColWidth(lintColSubfamilia) = 1800      '| Subfamilia
        .ColWidth(lintColCodigoBarras) = 1600    '| Código de barras
        .ColWidth(lintColExcluido) = 0           '| Indica si es excluido 0, 1 es excluido
        .ColWidth(lintColCuadroBasico) = 0       '| Indica si es un medicamento que pertenece al cuadro básico 0 no pertenece, 1 si pertenece
        .ColWidth(lintColTipo) = 0               '| indica si es 0=Articulo, 1=Medicamento, 2=Insumo
        
        .ColAlignment(lintColPrecioUM) = flexAlignRightCenter
        .ColAlignment(lintColPrecioUA) = flexAlignRightCenter
        
        .Redraw = True
    End With
End Sub

' Procedimiento para llenar el grid con las columnas del recordset más las de los manejos de medicamentos '
Private Sub pLlenarMshFGrdRsManejos(ObjGrid As MSHFlexGrid, objRS As Recordset, ByRef vllintTotalManejos As Integer, ByRef vlintColInicial As Integer, Optional vlstrColumnaData As String)
On Error GoTo NotificaError
    
    Dim vlintNumCampos As Long 'Total de Columnas
    Dim vlintNumReg As Long    'Total de Renglones
    Dim vlintSeqFil As Long    'Variable para el seguimiento de los renglones
    Dim vlintSeqCol As Long    'Variable para el seguimiento de las columnas
    Dim vlintSeqReg As Long    'Variable para el seguimiento de los registros del recordset
    
    vlintNumCampos = objRS.Fields.Count
    vlintNumReg = objRS.RecordCount
    
    If vlintNumReg > 0 Then
        With ObjGrid
            .Redraw = False
            .Visible = False
            .Clear
            .ClearStructure
            .Cols = vlintNumCampos + vllintTotalManejos + 1
            .Rows = vlintNumReg + 1
            .FixedCols = 1
            .FixedRows = 1
        
            objRS.MoveFirst
            For vlintSeqFil = 1 To vlintNumReg
                vlintSeqReg = 0
                For vlintSeqCol = vlintColInicial To vlintNumCampos + vllintTotalManejos
                    If IsNull(objRS.Fields(vlintSeqReg).Value) Then
                        .TextMatrix(vlintSeqFil, vlintSeqCol) = ""
                    Else
                        If vlstrColumnaData <> "" Then
                            If vlintSeqCol - 1 = Val(vlstrColumnaData) Then
                                .RowData(vlintSeqFil) = objRS.Fields(vlintSeqCol - 1)
                            End If
                        End If
                        .TextMatrix(vlintSeqFil, vlintSeqCol) = objRS.Fields(vlintSeqReg).Value
                    End If
                    vlintSeqReg = vlintSeqReg + 1
                Next vlintSeqCol
                objRS.MoveNext
            Next vlintSeqFil
            
            .Redraw = True
            .Visible = True
        End With
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarMshFGrdRsManejos"))
End Sub

'- Función agregada para la configuración de manejos de medicamentos -'
Function fintConfigManejosMedicamentos(cboFuente As MyCombo) As Integer
On Error GoTo NotificaError
    Dim lintcnt As Integer
    Dim lstrsql As String
    
    lintcnt = 0
    '- Traer el total de manejos de medicamentos activos -'
    lstrsql = "SELECT MAX(Manejos) as TotalManejos FROM " & _
               "(SELECT IVARTICULO.INTIDARTICULO, COUNT(IvManejoMedicamento.VCHSIMBOLO) Manejos " & _
               "FROM IVARTICULO " & _
               "LEFT JOIN IVARTICULOMANEJO ON IVARTICULO.INTIDARTICULO = IVARTICULOMANEJO.INTIDARTICULO " & _
               "LEFT JOIN IvManejoMedicamento ON IvManejoMedicamento.intCveManejo = IVARTICULOMANEJO.INTCVEMANEJO " & _
               "AND IvManejoMedicamento.BITACTIVO = 1 " & _
               "WHERE IVARTICULO.VCHESTATUS = 'ACTIVO' " & _
               "GROUP BY IVARTICULO.INTIDARTICULO " & _
               "ORDER BY Manejos DESC) "
    Set rsManejoMedicamentos = frsRegresaRs(lstrsql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        lintcnt = rsManejoMedicamentos!TotalManejos
    End If
    rsManejoMedicamentos.Close
    fintConfigManejosMedicamentos = lintcnt
        
    '- Traer los manejos de medicamentos por cada artículo y llenar el combo -'
    lstrsql = "SELECT IVARTICULO.CHRCVEARTICULO, IVARTICULOMANEJO.INTCVEMANEJO " & _
              " FROM IVARTICULOMANEJO " & _
              " LEFT OUTER JOIN IVARTICULO ON IVARTICULOMANEJO.INTIDARTICULO = IVARTICULO.INTIDARTICULO " & _
              " ORDER BY IVARTICULO.CHRCVEARTICULO, IVARTICULOMANEJO.INTCVEMANEJO"
    Set rsManejoMedicamentos = frsRegresaRs(lstrsql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        With cboFuente
            .Clear
            rsManejoMedicamentos.MoveFirst
            For lintcnt = 0 To rsManejoMedicamentos.RecordCount - 1
                .AddItem IIf(IsNull(rsManejoMedicamentos!chrcvearticulo), "", rsManejoMedicamentos!chrcvearticulo)
                .ItemData(lintcnt) = IIf(IsNull(rsManejoMedicamentos!intCveManejo), 0, rsManejoMedicamentos!intCveManejo)
                rsManejoMedicamentos.MoveNext
            Next
        End With
    End If
    rsManejoMedicamentos.Close
    
    '- Traer los manejos de medicamentos activos y llenar el arreglo -'
    lstrsql = "SELECT * FROM IvManejoMedicamento WHERE BITACTIVO = 1 ORDER BY 1"
    Set rsManejoMedicamentos = frsRegresaRs(lstrsql, adLockOptimistic, adOpenForwardOnly)
    If rsManejoMedicamentos.RecordCount > 0 Then
        ReDim aManejoMedicamentos(rsManejoMedicamentos.RecordCount)
        rsManejoMedicamentos.MoveFirst
        For lintcnt = 0 To rsManejoMedicamentos.RecordCount - 1
            aManejoMedicamentos(lintcnt).intCveManejo = rsManejoMedicamentos!intCveManejo
            aManejoMedicamentos(lintcnt).strColor = rsManejoMedicamentos!vchColor
            aManejoMedicamentos(lintcnt).strSimbolo = rsManejoMedicamentos!vchSimbolo
            rsManejoMedicamentos.MoveNext
        Next
    End If
    rsManejoMedicamentos.Close
    
    Exit Function
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fintConfigManejosMedicamentos"))
End Function

'- Procedimiento para el formato visual de los manejos de los medicamentos -'
Private Sub pColorearManejo(grdFuente As MSHFlexGrid, ByVal cboFuente As MyCombo, ByVal intColumna As Integer, ByVal strCveMedicamento As String, ByVal intRenglon As Integer, Optional intColWidth = 320)
On Error GoTo NotificaError
    
    Dim vlstrClave As String
    Dim vlintlista As Integer, vlintCnt As Integer, vlintCol As Integer
    
    If Trim(strCveMedicamento) = "" Then Exit Sub
    
    vlintCol = intColumna
    For vlintlista = 0 To cboFuente.ListCount - 1
        If (cboFuente.List(vlintlista) = strCveMedicamento) Then
            vlstrClave = cboFuente.ItemData(vlintlista)
            vlintCnt = 0
            Do While vlintCnt < UBound(aManejoMedicamentos)
                If aManejoMedicamentos(vlintCnt).intCveManejo = vlstrClave Then
                    With grdFuente
                        .TextMatrix(intRenglon, vlintCol) = aManejoMedicamentos(vlintCnt).strSimbolo
                        .Row = intRenglon
                        .Col = vlintCol
                        .CellFontBold = False
                        .CellFontName = "Wingdings"
                        .CellFontSize = 12
                        .CellForeColor = CLng(aManejoMedicamentos(vlintCnt).strColor)
                        .CellBackColor = .BackColor
                        .ColAlignment(vlintCol) = flexAlignCenterCenter
                        .ColWidth(vlintCol) = intColWidth
                        vlintCol = vlintCol + 1
                    End With
                End If
                vlintCnt = vlintCnt + 1
            Loop
        End If
    Next vlintlista
    
    Exit Sub
    
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pColorearManejo"))
End Sub
