VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMantoArea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Areas"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   4035
      Left            =   -45
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   -75
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   7117
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   26
      WordWrap        =   0   'False
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "frmMantoArea.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "shp0(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "shp1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescripcion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCveArea"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "shp3(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shp2(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCvePiso"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblMin"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEdadMax"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEdadMin"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCveArea"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtDescripcion"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdPrimerRegistro"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdAnteriorRegistro"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdBuscar"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdSiguienteRegistro"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdUltimoRegistro"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdGrabarRegistro"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboPiso"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cboSexo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cboAsignacion"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Búsqueda"
      TabPicture(1)   =   "frmMantoArea.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTitulo"
      Tab(1).Control(1)=   "grdHBusqueda"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cboAsignacion 
         Height          =   315
         ItemData        =   "frmMantoArea.frx":0038
         Left            =   2940
         List            =   "frmMantoArea.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Clasificación del área: Normal o de Tránsito"
         Top             =   1905
         Width           =   3270
      End
      Begin VB.ComboBox cboSexo 
         Height          =   315
         ItemData        =   "frmMantoArea.frx":005B
         Left            =   2940
         List            =   "frmMantoArea.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Clasificación del área por sexo"
         Top             =   1515
         Width           =   3270
      End
      Begin VB.ComboBox cboPiso 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "frmMantoArea.frx":0088
         Left            =   2940
         List            =   "frmMantoArea.frx":008A
         TabIndex        =   2
         ToolTipText     =   "Piso donde se encuentra el área"
         Top             =   1110
         Width           =   3270
      End
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   5190
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Guardar el registro"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdUltimoRegistro 
         Height          =   495
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":022E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ultimo registro"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguienteRegistro 
         Height          =   495
         Left            =   4170
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":03A0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Siguiente registro"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   495
         Left            =   3660
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":0512
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Búsqueda"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAnteriorRegistro 
         Height          =   495
         Left            =   3150
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Anterior registro"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrimerRegistro 
         Height          =   495
         Left            =   2655
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoArea.frx":07F6
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Primer registro"
         Top             =   3270
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtDescripcion 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   2940
         TabIndex        =   1
         ToolTipText     =   "Descripción del área"
         Top             =   735
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   45
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtCveArea 
         Height          =   315
         Left            =   2940
         TabIndex        =   0
         ToolTipText     =   "Clave del área"
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         DragIcon        =   "frmMantoArea.frx":0968
         Height          =   3120
         Left            =   -74415
         TabIndex        =   22
         Top             =   585
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5503
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Rows            =   16
         Cols            =   3
         BackColorBkg    =   -2147483639
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         MergeCells      =   1
         FormatString    =   "|tnyCvePiso|vchDescripcion"
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
      End
      Begin MSMask.MaskEdBox txtEdadMin 
         Height          =   315
         Left            =   3510
         TabIndex        =   5
         ToolTipText     =   "Edad mínima para los pacientes que ingresen a esta área"
         Top             =   2430
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtEdadMax 
         Height          =   315
         Left            =   4665
         TabIndex        =   6
         ToolTipText     =   "Edad máxima para los pacientes que ingresen a esta área"
         Top             =   2445
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Máximo"
         Height          =   165
         Left            =   4050
         TabIndex        =   20
         Top             =   2490
         Width           =   570
      End
      Begin VB.Label lblMin 
         Caption         =   "Minimo"
         Height          =   165
         Left            =   2940
         TabIndex        =   19
         Top             =   2490
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Clasificación de cuartos por edades"
         Height          =   465
         Left            =   1320
         TabIndex        =   18
         ToolTipText     =   "Número del piso"
         Top             =   2340
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "Clasificación del área"
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         ToolTipText     =   "Número del piso"
         Top             =   1935
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Clasificación por sexo"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "Número del piso"
         Top             =   1545
         Width           =   1665
      End
      Begin VB.Label lblCvePiso 
         Caption         =   "Piso"
         Height          =   225
         Left            =   1320
         TabIndex        =   15
         Top             =   1155
         Width           =   1500
      End
      Begin VB.Shape shp2 
         BorderColor     =   &H8000000C&
         Height          =   2790
         Index           =   2
         Left            =   165
         Top             =   150
         Width           =   8100
      End
      Begin VB.Shape shp3 
         BorderColor     =   &H80000005&
         Height          =   2790
         Index           =   3
         Left            =   180
         Top             =   165
         Width           =   8100
      End
      Begin VB.Label lblCveArea 
         Caption         =   "Clave"
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         ToolTipText     =   "Número del piso"
         Top             =   375
         Width           =   825
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         Height          =   225
         Left            =   1320
         TabIndex        =   14
         Top             =   780
         Width           =   1575
      End
      Begin VB.Shape shp1 
         BorderColor     =   &H80000005&
         Height          =   615
         Index           =   1
         Left            =   2580
         Top             =   3195
         Width           =   3180
      End
      Begin VB.Shape shp0 
         BorderColor     =   &H8000000C&
         Height          =   615
         Index           =   0
         Left            =   2565
         Top             =   3180
         Width           =   3180
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Búsqueda de áreas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74415
         TabIndex        =   23
         Top             =   195
         Width           =   2340
      End
   End
End
Attribute VB_Name = "frmMantoArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Admisión
'| Nombre del Formulario    : frmMantoArea
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de Areas
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Nery Lozano - Luis Astudillo
'| Autor                    : Nery Lozano - Luis Astudillo
'| Fecha de Creación        : 24/Noviembre/1999
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables

Dim vlintCvePiso As Integer 'Guarda la informacion de la clave del piso
Dim vgblnNuevoRegistro As Boolean

Private Sub cboAsignacion_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cboAsignacion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            Call pEnfocaMkTexto(txtEdadMin)
    End Select

End Sub

Private Sub cboPiso_LostFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cboSexo_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cboSexo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            cboAsignacion.SetFocus
    End Select
End Sub

Private Sub cmdAnteriorRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdBuscar_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdBuscar.TabIndex)
End Sub

Private Sub cmdGrabarRegistro_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdGrabarRegistro.TabIndex)
End Sub

Private Sub cmdPrimerRegistro_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdPrimerRegistro.TabIndex)
End Sub

Private Sub cmdSiguienteRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdUltimoRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub Form_Load()
'-------------------------------------------------------------------------------------------
'  Define implicitamente el RS al abrirlo y lo relaciona con el objeto cmd del
'  Data Environment. Relaciona el grdHBusqueda con el DataEnvironment
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    If vgblnNuevoRegistro = True Then
        EntornoSIHO.rscmdPiso.Open 'Abre la conexión con la tabla de piso utilizando un RS
        EntornoSIHO.rscmdArea.Open 'Abre la conexión con la tabla de area utilizando un RS
    End If
    
    Call pIniciaGrid
    
    vgstrVarIntercam = ""
    vgblnErrorIngreso = False
    vgstrNombreForm = Me.Name 'Nombre del formulario que se utiliza actualmente
    vgblnExistioError = False 'Inicia la bandera sin errores
    sstObj.Tab = 0 'Se localiza en el primer tabulador para la alta
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    vgintTipoOrd = 1 'Que tipo de ordenamiento realizará de inicio en el grdHBusqueda
    vgintColLoc = 1 'Localiza la búsqueda de registros para la primera columna del grdHBusqueda
    
    Call pConfMshFGrid(grdHBusqueda, "|Clave|Descripción|Número Piso|Descripción del Piso||||")
    Call pLlenarCboRs(cboPiso, EntornoSIHO.rscmdPiso, 0, 1, 2)
    pAgregarRegistro ("I") 'Permite agregar un registro nuevo
    
    
NotificaError:
    If vgblnExistioError Then
        Unload Me
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'-------------------------------------------------------------------------------------------
' Cierra el RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pSalirForm

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cboPiso_Click()
    If cboPiso.ListIndex = 0 Then
        cboPiso.Clear
        EntornoSIHO.rscmdPiso.Close
        Load frmMantoPiso
        frmMantoPiso.Show vbModal, Me
        EntornoSIHO.rscmdPiso.Open
        Call pLlenarCboRs(cboPiso, EntornoSIHO.rscmdPiso, 0, 1, 2)
        If CInt(vgstrVarIntercam) = 0 Then
            cboPiso.ListIndex = fintLocalizaCbo(cboPiso, CStr(vlintCvePiso))
        Else
            cboPiso.ListIndex = fintLocalizaCbo(cboPiso, vgstrVarIntercam)
        End If
        vlintCvePiso = CInt(cboPiso.ItemData(cboPiso.ListIndex))
    Else
        If cboPiso.ListIndex > 0 Then
            vlintCvePiso = CInt(cboPiso.ItemData(cboPiso.ListIndex))
        End If
    End If
End Sub
Private Sub cboPiso_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación de la tecla presionada en el último campo del formulario, <Esc> cancela lo capturado,
'el <Enter> graba los datos capturados.
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            cboSexo.SetFocus
'            Call pValidaVarCbo(CStr(vlintCvePiso), cboPiso, "N", "", 2, False)
'            cboPiso.ListIndex = fintLocalizaCbo(cboPiso, CStr(vlintCvePiso))
'            cmdGrabarRegistro_Click
    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboPiso_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If

End Sub
Private Sub grdHBusqueda_Click()
'-------------------------------------------------------------------------------------------
' Refresca el GrdHBusqueda y asigna bajo que columna se va a hacer la búsqueda
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    If grdHBusqueda.Rows > 0 Then
        grdHBusqueda.Refresh
        vgintColLoc = grdHBusqueda.Col
        vgstrAcumTextoBusqueda = ""
        grdHBusqueda.Col = vgintColLoc
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub grdHBusqueda_DblClick()
'-------------------------------------------------------------------------------------------
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    If grdHBusqueda.Rows > 0 Then
        vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
        
        ' Ordena solamente cuando un encabezado de columna es seleccionado con un click
        If grdHBusqueda.MouseRow >= grdHBusqueda.FixedRows Then
            sstObj.Tab = 0
            vlintNumero = fintLocalizaPkRs(EntornoSIHO.rscmdArea, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
            pModificaRegistro
            Exit Sub
        End If
        vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
        vgintColOrd = grdHBusqueda.Col  'Configura la columna a ordenar
        
        'Escoge el Tipo de Ordenamiento
        If vgintTipoOrd = 1 Then
             vgintTipoOrd = 2
            Else
                vgintTipoOrd = 1
            End If
        grdHBusqueda.FocusRect = flexFocusNone
        Call pOrdColMshFGrid(grdHBusqueda, vgintTipoOrd)
        Call pDesSelMshFGrid(grdHBusqueda)
        grdHBusqueda.FocusRect = flexFocusHeavy
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub grdhBusqueda_GotFocus()
    Call pVerificaPosTab(grdHBusqueda.TabIndex)
End Sub

Private Sub grdhBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
' Validación del <Escape> para regresar al Tab 0 (Mantenimiento) del sstObj, teniendo el
' enfoque en GrdHBusqueda
'-------------------------------------------------------------------------------------------
    Dim vlintNumero As Integer
    
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyReturn
            If grdHBusqueda.Rows > 0 Then
                sstObj.Tab = 0
                vlintNumero = fintLocalizaPkRs(EntornoSIHO.rscmdArea, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
                pModificaRegistro
            End If
        Case vbKeyEscape
            sstObj.Tab = 0
            pAgregarRegistro ("")
    End Select

    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub grdHBusqueda_KeyPress(vlintKeyAscii As Integer)
'-------------------------------------------------------------------------------------------
' Evento que verifica si se presiono una tecla
' de la A-Z, a-z, 0-9, á,é,í,ó,ú,ñ,Ñ, se presiono la barra espaciadora
' Realizando la búsqueda de un criterio dentro del grdHBusqueda
'-------------------------------------------------------------------------------------------
    If grdHBusqueda.Rows > 0 Then
        Call pSelCriterioMshFGrid(grdHBusqueda, vgintColLoc, vlintKeyAscii)
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_KeyPress"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub sstObj_GotFocus()
    Call pVerificaPosTab(sstObj.TabIndex)
End Sub

Private Sub sstObj_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
' Validación del <Escape> en el segundo Tab del sstObj(Búsqueda) cuando no tiene el enfoque el Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyEscape
            sstObj.Tab = 0
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstObj_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtCveArea_GotFocus()
'-------------------------------------------------------------------------------------------
' Seleccionar el cuadro de texto del primer control del formulario
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pEnfocaMkTexto(txtCveArea)
    Call pVerificaPosTab(txtCveArea.TabIndex)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveArea_GotFocus"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub txtCveArea_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlbytNumero As Byte
    Select Case KeyCode
        Case vbKeyReturn
            'Buscar criterio
            If (Len(txtCveArea.Text) <= 0) Then
                txtCveArea.Text = "0"
            End If
            'Si crea un nuevo registro
            If fintSigNumRs(EntornoSIHO.rscmdArea, 0) = CByte(txtCveArea.Text) Then
                txtCveArea.Enabled = False
                txtDescripcion.Enabled = True
                Call pEnfocaMkTexto(txtDescripcion)
                cboPiso.Enabled = True
                If cboPiso.ListCount = 1 Then
                    cboPiso.ListIndex = 0
                Else
                    cboPiso.ListIndex = 1
                End If
                cboSexo.Enabled = True
                cboSexo.ListIndex = 0
                cboAsignacion.Enabled = True
                cboAsignacion.ListIndex = 0
                txtEdadMin.Enabled = True
                txtEdadMax.Enabled = True
                txtEdadMin.Text = "0"
                txtEdadMax.Text = "0"
                
                vgblnNuevoRegistro = True
            Else 'Si modifica un registro
                vlbytNumero = fintLocalizaPkRs(EntornoSIHO.rscmdArea, 0, txtCveArea.Text)
                If vlbytNumero > 0 Then
                    pModificaRegistro
                Else
                    Call MsgBox(SIHOMsg("12"), vbExclamation, "Mensaje")
                    pAgregarRegistro ("")
                End If
            End If
        Case vbKeyEscape
            Unload Me
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveArea_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pAgregarRegistro(vlstrSelTxt As String)
'-------------------------------------------------------------------------------------------
' Prepara el estado de un alta de registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    txtCveArea.Text = CStr(fintSigNumRs(EntornoSIHO.rscmdArea, 0)) 'Muestra el siguiente consecutivo del campo Clave
    txtCveArea.Enabled = True 'Habilita el ingreso de una clave para su búsqueda
    txtDescripcion.Enabled = False
    txtDescripcion.Text = ""
    vlintCvePiso = 0
    cboPiso.Enabled = False
    cboPiso.Text = ""
    cboSexo.ListIndex = 0
    cboSexo.Enabled = False
    cboAsignacion.ListIndex = 0
    cboAsignacion.Enabled = False
    txtEdadMin.Text = ""
    txtEdadMax.Text = ""
    txtEdadMin.Enabled = False
    txtEdadMax.Enabled = False

    pHabilitaBotonBuscar True
    
    If vlstrSelTxt <> "I" Then
        Call pEnfocaMkTexto(txtCveArea)
    End If
    vgblnNuevoRegistro = True
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregarRegistro"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pModificaRegistro()
'-------------------------------------------------------------------------------------------
' Permite realizar la modificación de la descripción de un registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    vgblnNuevoRegistro = False
    txtCveArea.Text = EntornoSIHO.rscmdArea.Fields(0).Value
    txtDescripcion.Text = EntornoSIHO.rscmdArea.Fields(1).Value
    vlintCvePiso = EntornoSIHO.rscmdArea.Fields(2).Value
    txtDescripcion.Enabled = True
    txtCveArea.Enabled = False
    cboPiso.Enabled = True
    cboPiso.ListIndex = fintLocalizaCbo(cboPiso, CStr(vlintCvePiso))
    
    cboSexo.Enabled = True
    cboSexo.ListIndex = EntornoSIHO.rscmdArea.Fields(4).Value
    cboAsignacion.Enabled = True
    cboAsignacion.ListIndex = EntornoSIHO.rscmdArea.Fields(7).Value
    txtEdadMin.Enabled = True
    txtEdadMax.Enabled = True
    txtEdadMin.Text = EntornoSIHO.rscmdArea.Fields(5).Value
    txtEdadMax.Text = EntornoSIHO.rscmdArea.Fields(6).Value

    
    Call pEnfocaMkTexto(txtDescripcion)
    Call pHabilitaBotonModifica(True)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRegistro"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
    Call pSelMkTexto(txtDescripcion)
    Call pVerificaPosTab(txtDescripcion.TabIndex)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación de la tecla presionada en el último campo del formulario, <Esc> cancela lo capturado,
'el <Enter> graba los datos capturados.
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            If fblnValidaMkText(txtDescripcion, "T", ">", 45, True) = False Then
                If fblnExisteCriterioRs(EntornoSIHO.rscmdArea, 0, 1, txtDescripcion, txtCveArea) = True Then
                    Call MsgBox(SIHOMsg("19") & Chr(13) & "Dato:" & txtDescripcion.ToolTipText, vbExclamation, "Mensaje")
                Else
                    Call pEnfocaCbo(cboPiso)
                End If
            End If
    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdAnteriorRegistro_Click()
'-------------------------------------------------------------------------------------------
' Manda llamar los procedimientos pPosicionaRegRs y pModificaRegistro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pPosicionaRegRs(EntornoSIHO.rscmdArea, "A")
    pModificaRegistro
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdBuscar_Click()
'-------------------------------------------------------------------------------------------
' Manda el enfoque al Tab 1 del sstObj para visualizar la búsqueda y actualizar el Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    sstObj.Tab = 1
    Call pIniciaGrid
    'Call pRefrescaMshFGrid(grdHBusqueda, EntornoSIHO.rscmdArea.RecordCount)
    Call pConfMshFGrid(grdHBusqueda, "|Clave|Descripción|Número Piso|Descripción del Piso||||")
    grdHBusqueda.SetFocus
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdGrabarRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite crear un nuevo registro o actualizar la información de un registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlstrTexto As String
    Dim vlintNum As Integer

If fintBuscaCharStr(vgstrPermisoUsuarioPro, "C", True) > 0 Or fintBuscaCharStr(vgstrPermisoUsuarioPro, "E", True) > 0 Then
    vlintNum = 0
    
    Call pActualizaVar(Nothing)
    
    If vlintNum = 0 Then
        If fblnValidaMkText(txtDescripcion, "T", ">", 45, True) = True Then
            Call pEnfocaMkTexto(txtDescripcion)
        Else
            If fblnExisteCriterioRs(EntornoSIHO.rscmdArea, 0, 1, txtDescripcion, txtCveArea) = True Then
                Call MsgBox(SIHOMsg("19") & Chr(13) & "Dato:" & txtDescripcion.ToolTipText, vbExclamation, "Mensaje")
                vlintNum = vlintNum + 1
                Call pEnfocaMkTexto(txtDescripcion)
            End If
        End If
    End If
    
    If vlintNum = 0 Then
        If cboPiso.ListIndex < 0 Then
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboPiso.ToolTipText, vbExclamation, "Mensaje")
            vlintNum = vlintNum + 1
            Call pEnfocaCbo(cboPiso)
        End If
    End If
    If vlintNum = 0 Then
        If cboSexo.ListIndex < 0 Then
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboSexo.ToolTipText, vbExclamation, "Mensaje")
            vlintNum = vlintNum + 1
            Call pEnfocaCbo(cboSexo)
        End If
    End If
    If vlintNum = 0 Then
        If cboAsignacion.ListIndex < 0 Then
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboAsignacion.ToolTipText, vbExclamation, "Mensaje")
            vlintNum = vlintNum + 1
            Call pEnfocaCbo(cboAsignacion)
        End If
    End If
    
    If CByte(txtEdadMin.Text) >= CByte(txtEdadMax.Text) Then
        txtEdadMin.Text = "0"
        txtEdadMax.Text = "0"
    End If
    
    If vlintNum = 0 Then
        If fblnValidaMkText(txtEdadMin, "N", "", 3, False) = True Then
            vlintNum = vlintNum + 1
            Call pEnfocaMkTexto(txtEdadMin)
        End If
    End If
    
    If vlintNum = 0 Then
        If fblnValidaMkText(txtEdadMax, "N", "", 3, False) = True Then
            vlintNum = vlintNum + 1
            Call pEnfocaMkTexto(txtEdadMax)
        End If
    End If
    
    If vlintNum = 0 Then
        If (Len(txtDescripcion.Text) > 0) And vlintCvePiso > 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            If vgblnNuevoRegistro = True Then
                EntornoSIHO.rscmdArea.AddNew
            End If
            EntornoSIHO.rscmdArea.Fields(1).Value = txtDescripcion.Text
            EntornoSIHO.rscmdArea.Fields(2).Value = vlintCvePiso
            EntornoSIHO.rscmdArea.Fields(4).Value = cboSexo.ListIndex
            EntornoSIHO.rscmdArea.Fields(5).Value = CByte(txtEdadMin.Text)
            EntornoSIHO.rscmdArea.Fields(6).Value = CByte(txtEdadMax.Text)
            EntornoSIHO.rscmdArea.Fields(7).Value = cboAsignacion.ListIndex
    
            EntornoSIHO.rscmdArea.Update
            If vgblnNuevoRegistro = True Then
                vgstrVarIntercam = CStr(EntornoSIHO.rscmdArea.Fields(0).Value)
            End If
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            Set grdHBusqueda.DataSource = Nothing
            grdHBusqueda.DataMember = ""
            
            EntornoSIHO.rscmdArea.Close
            EntornoSIHO.rscmdPiso.Close
            EntornoSIHO.rscmdArea.Open
            EntornoSIHO.rscmdPiso.Open
            pAgregarRegistro ("")
        End If
    End If
Else
    Call MsgBox(SIHOMsg("65"), vbInformation, "Mensaje")
End If
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            EntornoSIHO.ConeccionSIHO.RollbackTrans
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdPrimerRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el primer registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pPosicionaRegRs(EntornoSIHO.rscmdArea, "I")
    pModificaRegistro
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdSiguienteRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el siguiente registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pPosicionaRegRs(EntornoSIHO.rscmdArea, "S")
    pModificaRegistro
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub cmdUltimoRegistro_Click()
'-------------------------------------------------------------------------------------------
' Permite localizarse en el último registro del RS
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pPosicionaRegRs(EntornoSIHO.rscmdArea, "U")
    pModificaRegistro
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pConfMshFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
    'Configura el MSHFlexGrid
    'Configuraciones del grdHBusqueda
    ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
    
    ' Configura el ancho de las columnas del grdHBusqueda
    With ObjGrid
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 2900
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        
        .ColAlignmentFixed(1) = 1
        .ScrollBars = flexScrollBarVertical
    End With
End Sub
Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
    cmdGrabarRegistro.Enabled = vlblnHabilita
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonModifica"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pHabilitaBotonBuscar(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilita el botón de Buscar y deshabilita los demás botones
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = Not vlblnHabilita
    cmdAnteriorRegistro.Enabled = Not vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    cmdSiguienteRegistro.Enabled = Not vlblnHabilita
    cmdUltimoRegistro.Enabled = Not vlblnHabilita
    cmdGrabarRegistro.Enabled = Not vlblnHabilita
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaBotonBuscar"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pSalirForm()
'-------------------------------------------------------------------------------------------
' Cierra y limpia Recordsets, variables, Grid para el cierre del Form
'-------------------------------------------------------------------------------------------
    vgblnNuevoRegistro = False
    Set grdHBusqueda.DataSource = Nothing
    EntornoSIHO.rscmdArea.Close
    EntornoSIHO.rscmdPiso.Close 'Abre la conexión con la tabla de area utilizando un RS
    
End Sub
Private Sub cboPiso_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub pPresionoEscape(KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub cmdAnteriorRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdGrabarRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdPrimerRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdSiguienteRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdUltimoRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub
Private Sub pActualizaVar(ObjTxt As MaskEdBox)
    txtDescripcion.Text = fstrFormatTxt(txtDescripcion.Text, "T", ">", 45, True)
    If Not ObjTxt Is Nothing Then
        Call pEnfocaMkTexto(ObjTxt)
    End If
End Sub

Private Sub txtDescripcion_LostFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub txtEdadMax_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub txtEdadMax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            If CByte(txtEdadMin.Text) >= CByte(txtEdadMax.Text) Then
                Call MsgBox("La Edad mínima debe ser menor o igual a la Edad máxima", vbExclamation, "Mensaje")
                txtEdadMin.Text = "0"
                txtEdadMax.Text = "0"
            Else
                cmdGrabarRegistro_Click
            End If
    End Select
End Sub

Private Sub txtEdadMin_GotFocus()
    Call pActualizaVar(Nothing)
End Sub
Private Sub txtEdadMin_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
        Case vbKeyReturn
            Call pEnfocaMkTexto(txtEdadMax)
    End Select
End Sub
Private Sub pVerificaPosTab(vlintPosIndex As Integer)
    'Procedimiento para verificar la posicion de un control dentro del sstab
    If sstObj.Tab = 0 Then
        Select Case vlintPosIndex
            Case 21 To 23 'En el tab 1
                If txtCveArea.Enabled = True Then
                    Call pEnfocaMkTexto(txtCveArea)
                Else
                    If txtDescripcion.Enabled = True Then
                        Call pEnfocaMkTexto(txtDescripcion)
                    End If
                End If
        End Select
    End If
    If sstObj.Tab = 1 Then
        Select Case vlintPosIndex
            Case 0 To 21 'En el tab 0
                grdHBusqueda.SetFocus
        End Select
    End If
End Sub
Private Sub pIniciaGrid()
    Set grdHBusqueda.DataSource = EntornoSIHO
    grdHBusqueda.DataMember = "cmdArea"
    If grdHBusqueda.Rows = 0 Then
        Call pIniciaMshFGrid(grdHBusqueda)
        Call pLimpiaMshFGrid(grdHBusqueda)
    End If
End Sub
