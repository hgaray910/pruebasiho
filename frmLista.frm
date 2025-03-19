VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmLista 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   5925
   ClientTop       =   5430
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11415
   StartUpPosition =   1  'CenterOwner
   Begin MyCommandButton.MyButton cmdManejos 
      Height          =   375
      Left            =   10245
      TabIndex        =   3
      Top             =   4860
      Width           =   975
      _ExtentX        =   1720
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
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHLista 
      DragIcon        =   "frmLista.frx":0000
      Height          =   4530
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Artículos o medicamentos"
      Top             =   45
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7990
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorSel    =   0
      BackColorBkg    =   16777215
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
      Height          =   840
      Left            =   75
      TabIndex        =   1
      Top             =   4560
      Width           =   10140
      Begin VB.TextBox txtnombrecompleto 
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
         ForeColor       =   &H80000011&
         Height          =   525
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   220
         Width           =   9885
      End
   End
End
Attribute VB_Name = "frmLista"
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

Public gintEstatus As Integer
Public gintFamilia As Integer
Public gintSubfamilia As Integer
Public bvchDescripcion As String
Public bytCodigoBarras As Byte
Public glongCveDepartamento As Long 'Clave del departamento en presupuesto a salidas a departamento


Private objRS As New ADODB.Recordset

Dim intColCveArticulo As Integer
Dim intColNombreComercial As Integer
Dim intColCriterio As Integer
Dim intColdescripcion As Integer


Private Sub cmdManejos_Click()
    frmManejoMedicamentos.blnSoloBusqueda = True
    frmManejoMedicamentos.Show vbModal, Me
End Sub

Private Sub Form_Activate()
'----------------------------------------------------------------------------------------------------------
' El formulario permite revisar informacion sobre un comando objeto cualquiera
' Recibe como entrada dos parametros:
' La variable global vgstrVarIntercam de entrada interviene como criterio de busqueda de un campo
' La variable global vgstrVarIntercam2 de entrada interviene como titulo de la pantalla de lista y criterio
' para la busqueda.
' Envia como resultado en dos variables:
' La variable global vgstrVarIntercam contiene el valor del campo 1 cuando se da doble click en el grid
' La variable global vgstrVarIntercam2 contiene el valor del campo 2 cuando se da doble click en el grid
'----------------------------------------------------------------------------------------------------------
Dim vlstrNombreCampo As String
Dim vlstrCampo As String
Dim vlstrCriterio As String
Dim strSentencia As String
Dim intManejos As Integer
Dim intcontador As Integer
Dim strFString As String
Dim strFormatString As String
Dim blnCodigoBarras As Boolean
    
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    intManejos = 0
    strFormatString = ""
    
    grdHLista.Redraw = False
    grdHLista.Visible = False
            
    If Len(vgstrVarIntercam) > 0 Then
        
        Me.Caption = vgstrVarIntercam2
        
        Select Case Me.Tag
        'Tipo de requisicion
            Case "S"
            'Salida a departamento
                vlstrCriterio = "G"
            Case "R"
            'Reubicacion
                vlstrCriterio = "C"
            Case "D"
                vlstrCriterio = "T"
            'Presupuesto salida a departamento
            Case "P"
                vlstrCriterio = "G"
            'Reporte Presupuesto contra consumo de salidas a departamento
            Case "Z"
                vlstrCriterio = "G"
            Case Else
                vlstrCriterio = "T"
        End Select
        
        vgstrParametrosSP = CStr(gintFamilia) & "|" & CStr(gintSubfamilia) & "|" & CStr(gintEstatus) & "|" & vlstrCriterio & "|" & Replace(vgstrVarIntercam, "'", "''")
        blnCodigoBarras = False
        
        Select Case Trim(vgstrVarIntercam2)
            Case "Lista por nombre comercial"
                strFString = "|Clave|Nombre comercial|Código de barras|Nombre genérico"
                If Me.Tag = "D" Then
                    vgstrParametrosSP = vgstrParametrosSP & "|" & vgintNumeroDepartamento
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloComerciaDepto")
                ElseIf Me.Tag = "P" Then 'Presupuesto a salida a departamento por nombre comercial
                    vgstrParametrosSP = vgstrParametrosSP & "|" & glongCveDepartamento
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloComePresuDepto")
                ElseIf Me.Tag = "Z" Then 'Reporte Presupuesto contra consumo de salidas a departamento por nombre comercial
                    vgstrParametrosSP = vgstrParametrosSP & "|" & glongCveDepartamento
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvArticuloComeExPresuDepto")
                Else
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloNombreComercia")
                End If
                
            Case "Lista por nombre genérico"
                strFString = "|Clave|Nombre comercial|Código de barras|Nombre genérico"
                Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloNombreGenerico")
            
             Case "Lista por código de barras"
                strFString = "|Clave|Nombre comercial|Código de barras|Nombre genérico"
                Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloCodigoBarras")
                 
            Case "Lista por clave"
                If Me.Tag = "P" Then
                    strFString = "|Clave|Nombre comercial|Código de barras|Nombre genérico"
                    vgstrParametrosSP = vgstrParametrosSP & "|" & glongCveDepartamento
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloCvePresuDepto")
                Else
                    strFString = "|Clave|Nombre comercial|Código de barras|Nombre genérico"
                    Set objRS = frsEjecuta_SP(vgstrParametrosSP, "Sp_IvSelArticuloClave")
                End If
        End Select
        
        bytCodigoBarras = IIf(blnCodigoBarras, 1, 0)
        
        If objRS.RecordCount > 0 Then
                
            intManejos = objRS!Manejos

            For intcontador = 1 To intManejos
                strFormatString = strFormatString & "|."
            Next intcontador
            strFormatString = strFormatString & strFString
    
            pConfGrid strFormatString, intManejos
            
            With objRS
                Do While Not .EOF
                    If grdHLista.Rows > 2 Or grdHLista.TextMatrix(1, intColCveArticulo) <> "" Then
                        grdHLista.Rows = grdHLista.Rows + 1
                    End If
                    grdHLista.Row = grdHLista.Rows - 1
                    
                    grdHLista.TextMatrix(grdHLista.Row, intColCveArticulo) = IIf(IsNull(!chrcvearticulo), "", !chrcvearticulo)
                    grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial) = IIf(IsNull(!vchNombreComercial), "", !vchNombreComercial)
                    grdHLista.TextMatrix(grdHLista.Row, intColdescripcion) = IIf(IsNull(!VCHDESCRIPCION), "", !VCHDESCRIPCION)
                    grdHLista.TextMatrix(grdHLista.Row, intColCriterio) = IIf(IsNull(!vchGenerico), "", !vchGenerico)
                    
                    'Manejos
                    For intcontador = intManejos To 1 Step -1
                    
                        If Not IsNull(!vchSimbolo) Then
                            grdHLista.Col = intcontador
                            grdHLista.Row = grdHLista.Rows - 1
                            grdHLista.CellFontName = "Wingdings"
                            grdHLista.CellFontSize = 12
                            grdHLista.CellForeColor = CLng(!vchColor)
                            grdHLista.TextMatrix(grdHLista.Row, intcontador) = !vchSimbolo
                        End If
                        .MoveNext
                        
                        If .EOF Then
                            .MovePrevious
                            Exit For
                        Else
                            Select Case Trim(vgstrVarIntercam2)
                            
                            Case "Lista por nombre comercial"
                                If !vchNombreComercial <> grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial) Then
                                    .MovePrevious
                                    Exit For
                                End If
                            Case "Lista por nombre genérico", "Lista por código de barras"
                                If !VCHDESCRIPCION <> grdHLista.TextMatrix(grdHLista.Row, intColCriterio) Then
                                    .MovePrevious
                                    Exit For
                                End If
                            Case "Lista por clave"
                                If !chrcvearticulo <> grdHLista.TextMatrix(grdHLista.Row, intColCveArticulo) Then
                                    .MovePrevious
                                    Exit For
                                End If
                            End Select
                        End If
                        
                    Next intcontador
                    .MoveNext
                Loop
            End With
            
        Else
            
            Me.Hide
            If Me.Tag = "D" Then
                'No existen artículos o medicamentos que inicien con el texto que busca ubicados en este departamento.
                MsgBox SIHOMsg(1498), vbExclamation, "Mensaje"
            Else
                'No existen artículos o medicamentos que inicien con el texto que busca
                MsgBox SIHOMsg(55), vbExclamation, "Mensaje"
            End If
            Unload Me
            
        End If
        objRS.Close
        
        vgstrVarIntercam2 = ""
        vgstrVarIntercam = ""
        Me.Refresh
        
    End If
    
    grdHLista.Redraw = True
    grdHLista.Visible = True
    
    If fblnCanFocus(grdHLista) Then
        grdHLista.Col = intColNombreComercial
        grdHLista.Row = 1
        grdHLista.SetFocus
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pSalirForm
    End Select
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)

    gintEstatus = 1
    gintFamilia = 0
    gintSubfamilia = 0
    
End Sub



Private Sub grdHLista_Click()
    
    If grdHLista.Rows > 0 Then
        vgintColLoc = grdHLista.Col
        vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
        grdHLista.Col = vgintColLoc
        grdHLista.Refresh
    End If
    
End Sub

Private Sub grdHLista_DblClick()
Dim vgintColOrdAnt As Integer
Dim vlintNumero As Integer
    
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    
    If grdHLista.Rows > 0 Then
        If (grdHLista.Row <= grdHLista.Rows - 1) Then
            If grdHLista.MouseRow >= grdHLista.FixedRows Then
                vgstrVarIntercam = grdHLista.TextMatrix(grdHLista.Row, intColCveArticulo)
                vgstrVarIntercam2 = grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial)
                bvchDescripcion = grdHLista.TextMatrix(grdHLista.Row, intColdescripcion)
                pSalirForm
                Exit Sub
            End If
        End If
    End If
    
End Sub

Private Sub grdHLista_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            vgstrVarIntercam = grdHLista.TextMatrix(grdHLista.Row, intColCveArticulo)
            vgstrVarIntercam2 = grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial)
            bvchDescripcion = grdHLista.TextMatrix(grdHLista.Row, intColdescripcion)
            pSalirForm
    End Select
    
End Sub

Private Sub grdHLista_KeyPress(vlintKeyAscii As Integer)
'-------------------------------------------------------------------------------------------
' Evento que verifica si se presiono una tecla
' de la A-Z, a-z, 0-9, á,é,í,ó,ú,ñ,Ñ, se presiono la barra espaciadora
' Realizando la búsqueda de un criterio dentro del grdHBusqueda
'-------------------------------------------------------------------------------------------
    If vgintColLoc >= intColCveArticulo Then
        Call pSelCriterioMshFGrid(grdHLista, vgintColLoc, vlintKeyAscii)
    End If
    
End Sub

Private Sub pSelCriterioMshFGrid(ObjGrid As MSHFlexGrid, vlintColLoc As Integer, vlintCaracter As Integer)
' Realiza la busqueda de un criterio en el grid
On Error GoTo NotificaError

    ObjGrid.Redraw = False

    ObjGrid.FocusRect = flexFocusNone
    
    'Verifica los alfabeticos a-z, A-Z,0-9,ñÑ,áéíóúÁÉÍÓÚ,@/.=
    If ((vlintCaracter >= 65 And vlintCaracter <= 90) Or _
        (vlintCaracter >= 97 And vlintCaracter <= 122) Or _
        (vlintCaracter >= 48 And vlintCaracter <= 57) Or _
        vlintCaracter = 130 Or vlintCaracter = 160 Or vlintCaracter = 161 Or _
        vlintCaracter = 162 Or vlintCaracter = 163 Or vlintCaracter = 164 Or _
        vlintCaracter = 225 Or vlintCaracter = 233 Or vlintCaracter = 237 Or _
        vlintCaracter = 243 Or vlintCaracter = 250 Or vlintCaracter = 241 Or _
        vlintCaracter = 209 Or vlintCaracter = 193 Or vlintCaracter = 201 Or _
        vlintCaracter = 205 Or vlintCaracter = 211 Or vlintCaracter = 218 Or _
        vlintCaracter = 64 Or vlintCaracter = 44 Or vlintCaracter = 47 Or _
        vlintCaracter = 46 Or vlintCaracter = 42 Or vlintCaracter = 32) Then
        
        vgstrAcumTextoBusqueda = vgstrAcumTextoBusqueda & Chr(vlintCaracter)
        Call pLocalizaTxtMshFGrid(ObjGrid, vgstrAcumTextoBusqueda, vlintColLoc)
    Else
        If (vlintCaracter = 8) Then
            If Len(vgstrAcumTextoBusqueda) > 0 Then
                vgstrAcumTextoBusqueda = Left(vgstrAcumTextoBusqueda, (Len(vgstrAcumTextoBusqueda) - 1))
                If Len(vgstrAcumTextoBusqueda) > 0 Then
                    Call pLocalizaTxtMshFGrid(ObjGrid, vgstrAcumTextoBusqueda, vlintColLoc)
                Else
                    Call pDesSelMshFGrid(ObjGrid)
                End If
            Else
                Call pDesSelMshFGrid(ObjGrid)
            End If
        End If
    End If
    ObjGrid.Refresh
    ObjGrid.Col = vlintColLoc
    ObjGrid.FocusRect = flexFocusHeavy
    ObjGrid.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSelCriterioMshFGrid"))
End Sub

Private Sub pLocalizaTxtMshFGrid(ObjGrid As MSHFlexGrid, vlstrCriterio As String, vlintColBus As Integer)
' Realiza la busqueda de un criterio dentro de una de las columnas del grdHBusqueda
' señalando los criterios encontrados
On Error GoTo NotificaError
Dim vlintNumFilas As Integer 'Almacena el número de filas que contiene el grdHBusqueda
Dim vlintseq As Integer 'Contador del número de filas del grdHBusqueda
Dim vlintEFila As Integer 'Fila que se encuentra mediante el criterio de búsqueda
Dim vlintLargo As Integer 'Almacena el largo del criterio de búsqueda
Dim vlstrTexto As String 'Almacena los caracteres obtenidos de las celda del grid, segun el largo de busqueda del criterio, para su comparacion con el criterio de búsqueda
    
    ObjGrid.Redraw = False
    vlintLargo = Len(vlstrCriterio)
    vlintEFila = 0 'Inicializa la búsqueda desde la primera fila
    With ObjGrid
      If .Rows > 0 Then
        vlintNumFilas = .Rows - 1
        If vlintLargo > 0 Then
            For vlintseq = 1 To vlintNumFilas 'Realiza la búsqueda en todo el grid
                vlstrTexto = Left(.TextMatrix(vlintseq, vlintColBus), vlintLargo)
                If UCase(vlstrCriterio) = UCase(vlstrTexto) Then
                    If vlintLargo > 0 Then
                        Call pSelFilaMshFGrid(ObjGrid, vlintseq)
                        vlintEFila = vlintseq
                    End If
                Else
                    If vlintLargo > 0 Then
                        Call pDesSelFilaMshFGrid(ObjGrid, vlintseq)
                    End If
                End If
            Next vlintseq
            .Row = vlintEFila
            .Col = vlintColBus
            If vlintEFila > 0 Then
                .TopRow = vlintEFila
            Else
                vgstrAcumTextoBusqueda = ""
                .Row = 1
                .TopRow = 1
                .Col = 1
            End If
        End If
      End If
      .Redraw = True
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ": pLocalizaTxtMshFGrid"))
End Sub

Private Sub pDesSelFilaMshFGrid(ObjGrid As MSHFlexGrid, vlintNumFila As Integer)
' Quita la selección de una determinada fila dentro del grid
On Error GoTo NotificaError
Dim vlintNumColumnas As Integer 'Almacena el número de columnas del grid
Dim vlintSeqC As Integer 'Contador para el número de columnas del grid

    ObjGrid.Redraw = False
    ObjGrid.FocusRect = flexFocusNone
    vlintNumColumnas = ObjGrid.Cols - 1
    For vlintSeqC = intColCveArticulo To vlintNumColumnas
        ObjGrid.Row = vlintNumFila
        ObjGrid.Col = vlintSeqC
        ObjGrid.CellBackColor = vbWindowBackground
        ObjGrid.CellForeColor = vbWindowText
    Next vlintSeqC
    ObjGrid.FocusRect = flexFocusHeavy
    ObjGrid.Redraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDesSelFilaMshFGrid"))
End Sub

Private Sub pDesSelMshFGrid(ObjGrid As MSHFlexGrid)
' Limpia completamente las selecciones del Grid
On Error GoTo NotificaError
Dim vlintNumFilas As Long 'Almacena el número de filas que tiene el grid
Dim vlintSeqC As Integer 'Contador para el número de columnas
Dim vlintSeqF As Long 'Contador para el número de filas
Dim vlintNumColumnas As Integer 'Almacena el número de columnas que tiene el grid
    
    ObjGrid.Redraw = False
    ObjGrid.FocusRect = flexFocusNone
    vlintNumColumnas = ObjGrid.Cols - 1
    vlintNumFilas = ObjGrid.Rows - 1
    For vlintSeqF = 1 To vlintNumFilas
        For vlintSeqC = intColCveArticulo To vlintNumColumnas
            ObjGrid.Row = vlintSeqF
            ObjGrid.Col = vlintSeqC
            ObjGrid.CellBackColor = vbWindowBackground
            ObjGrid.CellForeColor = vbWindowText
        Next vlintSeqC
    Next vlintSeqF

    ObjGrid.Row = 1
    ObjGrid.Col = intColCveArticulo
    ObjGrid.FocusRect = flexFocusHeavy
    ObjGrid.Redraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDesSelMshFGrid"))
End Sub

Private Sub pSelFilaMshFGrid(ObjGrid As MSHFlexGrid, vlintNumFila As Integer)
' Realiza la seleccion de una fila determinada dentro del grid
On Error GoTo NotificaError
Dim vlintNumColumnas As Integer
Dim vlintSeqC As Integer
    
    ObjGrid.Redraw = False
    ObjGrid.FocusRect = flexFocusNone
    vlintNumColumnas = ObjGrid.Cols - 1
    For vlintSeqC = intColCveArticulo To vlintNumColumnas
        ObjGrid.Row = vlintNumFila
        ObjGrid.Col = vlintSeqC
        ObjGrid.CellBackColor = vbWindowBackground
        ObjGrid.CellForeColor = vbWindowText
        ObjGrid.CellBackColor = vbActiveTitleBar
        ObjGrid.CellForeColor = vbActiveTitleBarText
    Next vlintSeqC
    ObjGrid.FocusRect = flexFocusHeavy
    ObjGrid.Redraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSelFilaMshFGrid"))
End Sub

Private Sub pSalirForm()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub grdHLista_GotFocus()
    
    txtnombrecompleto.Text = ""
    If grdHLista.Row > 0 Then
      txtnombrecompleto.Text = grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial)
    End If
    
    If grdHLista.Rows <= 1 Then
        vgstrVarIntercam = ""
        vgstrVarIntercam2 = ""
        pSalirForm
    End If
    
End Sub
Private Sub pConfGrid(strFormatString As String, intManejos As Integer)
On Error GoTo NotificaError
Dim intcontador As Integer
    
    vgstrNombreProcedimiento = "pConfGrid"
    
    With grdHLista
        .Clear
        .Rows = 2
        .Cols = 5 + intManejos
        .FixedCols = 1
        .FixedRows = 1

        intColCveArticulo = 1 + intManejos
        intColNombreComercial = 2 + intManejos
        intColdescripcion = 3 + intManejos
        intColCriterio = 4 + intManejos

        .FormatString = strFormatString
        
        For intcontador = 1 To intManejos
            .Col = intcontador
            .Row = 0
            .ColWidth(intcontador) = 300
            .ColAlignment(intcontador) = flexAlignCenterCenter
            .CellForeColor = &H8000000F
            
        Next intcontador
    
        .ColWidth(0) = 300
        .ColWidth(intColCveArticulo) = 1300
        .ColWidth(intColNombreComercial) = 4800
        .ColWidth(intColdescripcion) = 1950
        .ColWidth(intColCriterio) = 4500
        
        .ColAlignment(intColdescripcion) = flexAlignRightCenter
        
        .MergeCells = flexMergeRestrictRows
        .MergeRow(0) = True
    
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfArt"))
    Unload Me
End Sub

Private Sub grdHLista_RowColChange()
    txtnombrecompleto.Text = ""
    If grdHLista.Row > 0 Then
      txtnombrecompleto.Text = grdHLista.TextMatrix(grdHLista.Row, intColNombreComercial)
    End If
End Sub
