VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoCuarto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Cuartos"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      CausesValidation=   0   'False
      Height          =   4800
      Left            =   -45
      TabIndex        =   23
      Top             =   -75
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8467
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   26
      TabMaxWidth     =   79
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoCuarto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPrecio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblObservacion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblEstadoCuarto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPiso"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblArea"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTipoCuarto"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Shape1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Shape1(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblClave"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblDescripcion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtObservacion"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPrecio"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPiso"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDescripcion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboEstadoCuarto"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboArea"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboTipoCuarto"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCveCuarto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboConceptoCargo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoCuarto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdHBusqueda"
      Tab(1).Control(1)=   "lblTitulo"
      Tab(1).ControlCount=   2
      Begin VB.ComboBox cboConceptoCargo 
         Height          =   315
         Left            =   2355
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Concncepto de facturación"
         Top             =   2055
         Width           =   5685
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   2595
         TabIndex        =   26
         Top             =   3540
         Width           =   3225
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   495
            Left            =   90
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Primer registro"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   495
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":01AA
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Anterior registro"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   495
            Left            =   1110
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":031C
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Búsqueda"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   495
            Left            =   1620
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":048E
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Siguiente registro"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   495
            Left            =   2130
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":0600
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Ultimo registro"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   495
            Left            =   2640
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoCuarto.frx":0772
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Guardar el registro"
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSMask.MaskEdBox txtCveCuarto 
         Height          =   300
         Left            =   2355
         TabIndex        =   0
         ToolTipText     =   "Número del cuarto"
         Top             =   390
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboTipoCuarto 
         Height          =   315
         ItemData        =   "frmMantoCuarto.frx":0914
         Left            =   2355
         List            =   "frmMantoCuarto.frx":0916
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipo de cuarto al que pertenece el cuarto"
         Top             =   1245
         Width           =   2475
      End
      Begin VB.ComboBox cboArea 
         Height          =   315
         ItemData        =   "frmMantoCuarto.frx":0918
         Left            =   5505
         List            =   "frmMantoCuarto.frx":091A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Area donde se encuentra el cuarto"
         Top             =   1245
         Width           =   2535
      End
      Begin VB.ComboBox cboEstadoCuarto 
         Height          =   315
         ItemData        =   "frmMantoCuarto.frx":091C
         Left            =   2355
         List            =   "frmMantoCuarto.frx":091E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Estado actual del cuarto"
         Top             =   1635
         Width           =   2475
      End
      Begin MSMask.MaskEdBox txtDescripcion 
         Height          =   300
         Left            =   2355
         TabIndex        =   1
         ToolTipText     =   "Descripción del cuarto"
         Top             =   870
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPiso 
         Height          =   300
         Left            =   5505
         TabIndex        =   14
         ToolTipText     =   "Piso donde se encuentra el cuarto"
         Top             =   1635
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPrecio 
         Height          =   315
         Left            =   2355
         TabIndex        =   7
         ToolTipText     =   "Precio del cuarto"
         Top             =   3090
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtObservacion 
         Height          =   555
         Left            =   2355
         TabIndex        =   6
         ToolTipText     =   "Observación a cerca del cuarto"
         Top             =   2460
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   979
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         DragIcon        =   "frmMantoCuarto.frx":0920
         Height          =   3195
         Left            =   -74715
         TabIndex        =   24
         Top             =   540
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   5636
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
         ScrollBars      =   2
         MergeCells      =   1
         FormatString    =   "|tnyCvePiso|vchDescripcion"
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto de cargo"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   2055
         Width           =   1710
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Búsqueda de cuartos"
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
         Left            =   -74715
         TabIndex        =   25
         Top             =   150
         Width           =   4950
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label lblClave 
         Caption         =   "Número de cuarto"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   405
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000005&
         Height          =   3255
         Index           =   1
         Left            =   285
         Top             =   285
         Width           =   7920
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   3255
         Index           =   0
         Left            =   270
         Top             =   270
         Width           =   7920
      End
      Begin VB.Label lblTipoCuarto 
         Caption         =   "Tipo de cuarto"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label lblArea 
         Caption         =   "Area"
         Height          =   255
         Left            =   4965
         TabIndex        =   20
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label lblPiso 
         Caption         =   "Piso"
         Height          =   255
         Left            =   4965
         TabIndex        =   21
         Top             =   1635
         Width           =   495
      End
      Begin VB.Label lblEstadoCuarto 
         Caption         =   "Estado del cuarto"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1635
         Width           =   1575
      End
      Begin VB.Label lblObservacion 
         Caption         =   "Observación"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   2445
         Width           =   1095
      End
      Begin VB.Label lblPrecio 
         Caption         =   "Precio"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   3090
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMantoCuarto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Admisión
'| Nombre del Formulario    : frmMantoCuarto
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de cuartos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        :
'| Autor                    :
'| Fecha de Creación        : 27/Noviembre/1999
'| Modificó                 : EL RODO
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables

Dim vlintTipoCuarto As Integer 'Guarda la informacion de la clave del tipo de cuarto
Dim vlintEstadoCuarto As Integer 'Guarda la informacion de la clave del estado del cuarto
Dim vlintArea As Integer 'Guarda la informacion de la clave del área
Dim vlintPiso As Integer 'Guarda la informacion de la clave del piso
Dim vlintColCve As Integer 'Guarda el numero de columna con la clave
Dim vgblnNuevoRegistro As Boolean
Dim rsCuarto As New ADODB.Recordset 'RS principal pal manto
Dim rsArea As New ADODB.Recordset 'RS para las areas
Dim rsEstadoCuarto As New ADODB.Recordset 'RS para el estado del cuarto
Dim rsTipoCuarto As New ADODB.Recordset 'RS pal tipo de cuarto
Dim rsConceptoCargo As New ADODB.Recordset 'RS para el combo de los conceptos

Private Sub cboArea_Validate(Cancel As Boolean)
    Call pValidaVarCbo(CStr(vlintArea), cboArea, "N", ">", 10, False)
End Sub

Private Sub cboConceptoCargo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call pEnfocaMkTexto(txtObservacion)
    End If
End Sub

Private Sub cboEstadoCuarto_Validate(Cancel As Boolean)
    Call pValidaVarCbo(CStr(vlintEstadoCuarto), cboEstadoCuarto, "N", "", 10, False)
End Sub

Private Sub cmdAnteriorRegistro_Click()
    Call pPosicionaRegRs(rsCuarto, "A")
    pModificaRegistro
End Sub

Private Sub cmdPrimerRegistro_Click()
    Call pPosicionaRegRs(rsCuarto, "I")
    pModificaRegistro
End Sub

Private Sub cmdSiguienteRegistro_Click()
    Call pPosicionaRegRs(rsCuarto, "S")
    pModificaRegistro
End Sub

Private Sub cmdUltimoRegistro_Click()
    Call pPosicionaRegRs(rsCuarto, "U")
    pModificaRegistro
End Sub

Private Sub Form_Load()
'-------------------------------------------------------------------------------------------
'  Define implicitamente el RS al abrirlo y lo relaciona con el objeto cmd del
'  Data Environment. Relaciona el grdHBusqueda con el DataEnvironment
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    pAbrirTablas 'Abre todas la tablas utilizadas en el procedimiento
    vgstrNombreForm = Me.Name 'Nombre del formulario que se utiliza actualmente
    vgblnExistioError = False 'Inicia la bandera sin errores
    vgblnErrorIngreso = False 'Inicia la bandera para verificar si un campo es necesario que sea ingresado

    SSTObj.Tab = 0 'Se localiza en el primer tabulador para la alta
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    vgintTipoOrd = 1 'Que tipo de ordenamiento realizará de inicio en el grdHBusqueda
    vgintColLoc = 1 'Localiza la búsqueda de registros para la primera columna del grdHBusqueda
    vgstrVarIntercam = ""

    'Configurar MSHFlexGrid utilizado
    '    Set grdHBusqueda.DataSource = EntornoSIHO
    '    grdHBusqueda.DataMember = "cmdCuarto"
    vlintColCve = 1
    Call pConfMshFGrid(grdhBusqueda, "|Cuarto|Tipo|Area")
    
    pLlenarCombos 'Llena los combos que se utilzaran el en formulario
    Call pAgregarRegistro("I")  'Permite agregar un registro nuevo
    
    
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

Private Sub pAbrirTablas()
'-------------------------------------------------------------------------------------------
'  Procedimiento para abrir las tablas necesarias para ejectuar el procedimiento
'-------------------------------------------------------------------------------------------
    Dim vlstrSentencia As String
    vlstrSentencia = "select adCuarto.*, adPiso.vchDescripcion Piso from AdCuarto " & _
                    " inner join adPiso ON adCuarto.tnyCvePiso = adPiso.tnyCvePiso"
    Set rsCuarto = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * from AdArea"
    Set rsArea = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * from AdEstadoCuarto"
    Set rsEstadoCuarto = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * from AdTipoCuarto"
    Set rsTipoCuarto = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * from pvOtroConcepto"
    Set rsConceptoCargo = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
End Sub

Private Sub pAgregarRegistro(vlstrSelTxt As String)
'-------------------------------------------------------------------------------------------
' Prepara el estado de un alta de registro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    txtCveCuarto.Enabled = True
    If (vlstrSelTxt <> "I") Then 'Cuando es el inicio no realiza enfoque
        txtCveCuarto.SetFocus
    End If
    txtCveCuarto.Text = ""
    txtDescripcion.Text = ""
    txtDescripcion.Enabled = False
    cboTipoCuarto.ListIndex = -1
    cboTipoCuarto.Enabled = False
    cboEstadoCuarto.ListIndex = -1
    cboEstadoCuarto.Enabled = False
    cboArea.ListIndex = -1
    cboArea.Enabled = False
    cboConceptoCargo.Enabled = False
    cboConceptoCargo.ListIndex = 0
    txtPiso.Text = ""
    txtPiso.Enabled = False
    txtObservacion.Text = ""
    txtObservacion.Enabled = False
    txtPrecio.Text = ""
    txtPrecio.Enabled = False
    vlintArea = 0
    vlintEstadoCuarto = 0
    vlintTipoCuarto = 0
    pHabilitaBotonBuscar True
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
Private Sub pLlenarCombos()
    '-------------------------------------------------------------------------------------------
    ' Procedimiento para Llenar datos en combo box
    '-------------------------------------------------------------------------------------------
    Call pLlenarCboRs(cboTipoCuarto, rsTipoCuarto, 0, 1, 2)
    Call pLlenarCboRs(cboEstadoCuarto, rsEstadoCuarto, 0, 1, 2)
    Call pLlenarCboRs(cboArea, rsArea, 0, 2, 2)
    Call pLlenarCboRs(cboConceptoCargo, rsConceptoCargo, 0, 1, 0)
End Sub
Private Sub pConfMshFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
    'Configura el MSHFlexGrid
    'Configuraciones del grdHBusqueda
    ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
    ' Configura el ancho de las columnas del grdHBusqueda
    With ObjGrid
        .Rows = 2
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .ColWidth(2) = 4000
        .ColWidth(3) = 2370
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        
        .ScrollBars = flexScrollBarVertical
    End With

End Sub
Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdanteriorregistro.Enabled = vlblnHabilita
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
    cmdanteriorregistro.Enabled = Not vlblnHabilita
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

Private Sub Form_Unload(Cancel As Integer)
    pSalirForm
End Sub

Private Sub grdHBusqueda_Click()
'-------------------------------------------------------------------------------------------
' Refresca el GrdHBusqueda y asigna bajo que columna se va a hacer la búsqueda
'-------------------------------------------------------------------------------------------
    If grdhBusqueda.Rows > 0 Then
        grdhBusqueda.Refresh
        vgintColLoc = grdhBusqueda.Col
        vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
        grdhBusqueda.Col = vgintColLoc
    End If
End Sub
Private Sub grdHBusqueda_DblClick()
' Muestra la información del registro encontrado y habilita su posible modificación
'-------------------------------------------------------------------------------------------
    
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Integer
    
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    ' Ordena solamente cuando un encabezado de columna es seleccionado con un click
    SSTObj.Tab = 0
    vlintNumero = fintLocalizaPkRs(rsCuarto, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1))
    pModificaRegistro

End Sub

Private Sub grdHBusqueda_GotFocus()
    Call pVerificaPosTab(grdhBusqueda.TabIndex)
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlintNumero As Integer
    Select Case KeyCode
        Case vbKeyReturn
            If grdhBusqueda.Rows > 0 Then
                SSTObj.Tab = 0
                vlintNumero = fintLocalizaPkRs(rsCuarto, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1))
                pModificaRegistro
            End If
        Case vbKeyEscape
            SSTObj.Tab = 0
            Call pAgregarRegistro("")
    End Select
End Sub
Private Sub grdHBusqueda_KeyPress(vlintKeyAscii As Integer)
'-------------------------------------------------------------------------------------------
' Evento que verifica si se presiono una tecla
' de la A-Z, a-z, 0-9, á,é,í,ó,ú,ñ,Ñ, se presiono la barra espaciadora
' Realizando la búsqueda de un criterio dentro del grdHBusqueda
'-------------------------------------------------------------------------------------------
    If grdhBusqueda.Rows > 0 Then
        grdhBusqueda.FocusRect = flexFocusNone
        Call pSelCriterioMshFGrid(grdhBusqueda, vgintColLoc, vlintKeyAscii)
        grdhBusqueda.FocusRect = flexFocusHeavy
    End If
End Sub



Private Sub sstObj_GotFocus()
    Call pConfMshFGrid(grdhBusqueda, "|Cuarto|Tipo|Area")
    grdhBusqueda.SetFocus
    Call pVerificaPosTab(SSTObj.TabIndex)
End Sub
Private Sub txtCveCuarto_GotFocus()
'-------------------------------------------------------------------------------------------
' Seleccionar el cuadro de texto del primer control del formulario
'-------------------------------------------------------------------------------------------
    Call pActualizaVar(Nothing)
    Call pEnfocaMkTexto(txtCveCuarto)
    Call pVerificaPosTab(txtCveCuarto.TabIndex)
End Sub
Private Sub txtCveCuarto_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    Dim vlbytNumero As Byte
    Dim vlstrMensaje As String
    Select Case KeyCode
        Case vbKeyReturn
            If (Len(txtCveCuarto.Text) <= 0) Then
                vlstrMensaje = SIHOMsg("2") & Chr(13) & "Campo:" & txtCveCuarto.ToolTipText
                Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
                Call pAgregarRegistro("")
            Else
                If fintLocalizaPkRs(rsCuarto, 0, txtCveCuarto.Text) > 0 Then
                    pModificaRegistro
                Else
                    PermNuevoIngreso
                End If
            End If

        Case vbKeyEscape
            Unload Me
    End Select
    
End Sub
Private Sub pModificaRegistro()
'-------------------------------------------------------------------------------------------
' Permite realizar la modificación de la descripción de un registro
'-------------------------------------------------------------------------------------------
    vgblnNuevoRegistro = False
    txtCveCuarto.Enabled = False
    txtDescripcion.Enabled = True
    cboArea.Enabled = True
    cboTipoCuarto.Enabled = True
    cboEstadoCuarto.Enabled = True
    cboConceptoCargo.Enabled = True
    txtPiso.Enabled = False
    txtObservacion.Enabled = True
    txtPrecio.Enabled = True
    
    vgblnHuboCambio = False 'Para comprobar si se ha realizado algun cambio en los campos
    
    txtCveCuarto.Text = rsCuarto!vchNumCuarto
    txtDescripcion.Text = rsCuarto!vchDescripcion
    vlintTipoCuarto = rsCuarto!tnyCveTipoCuarto
    vlintEstadoCuarto = rsCuarto!tnyCveEstadoCuarto
    vlintArea = rsCuarto!tnyCveArea
    vlintPiso = rsCuarto!tnyCvePiso
    txtPiso.Text = rsCuarto!piso
    txtObservacion.Text = IIf(IsNull(rsCuarto!vchObservacion), "", rsCuarto!vchObservacion)
    txtPrecio.Text = IIf(IsNull(rsCuarto!smyPrecio), "", rsCuarto!smyPrecio)
    cboTipoCuarto.ListIndex = fintLocalizaCbo(cboTipoCuarto, CStr(vlintTipoCuarto))
    cboEstadoCuarto.ListIndex = fintLocalizaCbo(cboEstadoCuarto, CStr(vlintEstadoCuarto))
    cboArea.ListIndex = fintLocalizaCbo(cboArea, CStr(vlintArea))
    cboConceptoCargo.ListIndex = fintLocalizaCbo(cboConceptoCargo, rsCuarto!intOtroConcepto)
    Call pEnfocaMkTexto(txtDescripcion)
    Call pHabilitaBotonModifica(True)
End Sub
Private Sub PermNuevoIngreso()
    Dim vlintNumReg As Integer
    Dim vlstrSentencia As String
    Dim rsPiso As New ADODB.Recordset
    
    txtCveCuarto.Enabled = False
    txtDescripcion.Enabled = True
    cboArea.Enabled = True
    cboTipoCuarto.Enabled = True
    cboEstadoCuarto.Enabled = True
    cboConceptoCargo.Enabled = True
    txtPiso.Enabled = False
    txtObservacion.Enabled = True
    txtPrecio.Enabled = True
    
    cboTipoCuarto.ListIndex = 1
    cboEstadoCuarto.ListIndex = 1
    cboArea.ListIndex = 1
    cboConceptoCargo.ListIndex = 0
    vlintArea = CInt(cboArea.ItemData(cboArea.ListIndex))
    vlintTipoCuarto = CInt(cboTipoCuarto.ItemData(cboTipoCuarto.ListIndex))
    vlintEstadoCuarto = CInt(cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex))
    vlintNumReg = fintLocalizaPkRs(rsArea, 0, cboArea.ItemData(cboArea.ListIndex))
    vlintPiso = rsArea!tnyCvePiso
    vlstrSentencia = "Select vchDescripcion Piso from adPiso where tnyCvePiso = " & Trim(Str(vlintPiso))
    Set rsPiso = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    txtPiso.Text = rsPiso!piso
    
    Call pEnfocaMkTexto(txtDescripcion)
End Sub

Private Sub txtDescripcion_GotFocus()
    Call pSelMkTexto(txtDescripcion)
    Call pVerificaPosTab(txtDescripcion.TabIndex)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaMkText(txtDescripcion, "T", ">", 15, True) = False Then
                Call pEnfocaCbo(cboTipoCuarto)
            End If
        Case vbKeyEscape
            Call pAgregarRegistro("")
    End Select
End Sub
Private Sub PresionoTecla(ObjTxt As MaskEdBox, KeyCode As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call pActualizaVar(ObjTxt)
        Case vbKeyEscape
            Call pAgregarRegistro("")
    End Select
End Sub
Private Sub pActualizaVar(ObjTxt As MaskEdBox)
    txtDescripcion.Text = fstrFormatTxt(txtDescripcion.Text, "T", ">", 15, True)
    txtObservacion.Text = fstrFormatTxt(txtObservacion.Text, "T", ">", 60, True)
    
    If Not ObjTxt Is Nothing Then
        Call pEnfocaMkTexto(ObjTxt)
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDescripcion_LostFocus()
    Call pActualizaVar(Nothing)
End Sub
Private Sub txtObservacion_GotFocus()
    Call pSelMkTexto(txtObservacion)
End Sub

Private Sub txtObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call pEnfocaMkTexto(txtPrecio)
        Case vbKeyEscape
            Call pAgregarRegistro("")
    End Select
End Sub
Private Sub txtPrecio_GotFocus()
    Call pActualizaVar(Nothing)
    Call pSelMkTexto(txtPrecio)
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaMkText(txtPrecio, "N", "", 10, False) = False Then
                If CDbl(txtPrecio) > 0 And CDbl(txtPrecio) <= 214748 Then
                    cmdGrabarRegistro_Click
                Else
                    Call MsgBox(SIHOMsg("26") & Chr(13) & "Dato:" & txtPrecio.ToolTipText, vbExclamation, "Mensaje")
                End If
            End If
            
        Case vbKeyEscape
            Call pAgregarRegistro("")
    End Select
End Sub
Private Sub pSalirForm()
'-------------------------------------------------------------------------------------------
' Cierra y limpia Recordsets, variables, Grid para el cierre del Form
'-------------------------------------------------------------------------------------------
    vgblnNuevoRegistro = False
    rsCuarto.Close
    rsArea.Close
    rsEstadoCuarto.Close
    rsTipoCuarto.Close
   
End Sub
Private Sub cboArea_Click()
    Dim vlintNumReg As Integer
    Dim rsPiso As New ADODB.Recordset
    Dim vlstrSentencia As String
    
    If cboArea.ListIndex = 0 Then
        Load frmMantoArea
        frmMantoArea.Show vbModal, Me
        'rsArea.Open
        
        'EntornoSIHO.rscmdPiso.Open 'ya no se requiere
        vlstrSentencia = "Select vchDescripcion Piso from adPiso where tnyCvePiso = " & Trim(Str(vlintPiso))
        Set rsPiso = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
                
        If vgstrVarIntercam = "0" Then
            vgstrVarIntercam = CStr(vlintArea)
        End If
        Call pLlenarCboRs(cboArea, rsArea, 0, 1, 2)
        If Len(vgstrVarIntercam) > 0 Then
            cboArea.ListIndex = fintLocalizaCbo(cboArea, CStr(vgstrVarIntercam))
            vlintArea = CInt(cboArea.ItemData(cboArea.ListIndex))
            vlintNumReg = fintLocalizaPkRs(rsArea, 0, CStr(vlintArea))
            'vlintPiso = rsPiso!Piso
            txtPiso.Text = rsPiso!piso
        Else
            cboArea.ListIndex = fintLocalizaCbo(cboArea, CStr(vlintArea))
        End If
        rsPiso.Close
    Else
        If cboArea.ListIndex > 0 Then
            vlintArea = CInt(cboArea.ItemData(cboArea.ListIndex))
            cboArea.ListIndex = fintLocalizaCbo(cboArea, CStr(vlintArea))
            vlintNumReg = fintLocalizaPkRs(rsArea, 0, CStr(vlintArea))
            vlintPiso = rsArea!tnyCvePiso
            
            vlstrSentencia = "Select vchDescripcion Piso from adPiso where tnyCvePiso = " & Trim(Str(vlintPiso))
            Set rsPiso = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
            txtPiso.Text = rsPiso!piso
        End If
    End If
End Sub
Private Sub cboArea_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vlintNumReg As Integer
    Select Case KeyCode
        Case vbKeyReturn
            cboConceptoCargo.SetFocus
        Case vbKeyEscape
            Call pAgregarRegistro("")
    End Select
End Sub

Private Sub cboEstadoCuarto_Click()
    If cboEstadoCuarto.ListIndex = 0 Then
        cboEstadoCuarto.Clear
        rsEstadoCuarto.Close
        Load frmMantoEstadoCuarto
        frmMantoEstadoCuarto.Show vbModal, Me
        rsEstadoCuarto.Open
        Call pLlenarCboRs(cboEstadoCuarto, rsEstadoCuarto, 0, 1, 2)
        If Len(vgstrVarIntercam) > 0 Then
            cboEstadoCuarto.ListIndex = fintLocalizaCbo(cboEstadoCuarto, CStr(vgstrVarIntercam))
            vlintEstadoCuarto = CInt(cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex))
        End If
    End If

    If cboEstadoCuarto.ListIndex > 0 Then
        vlintEstadoCuarto = CInt(cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex))
    End If

End Sub
Private Sub cboTipoCuarto_Click()
    If cboTipoCuarto.ListIndex = 0 Then
        cboTipoCuarto.Clear
        vgstrVarIntercam = ""
        rsTipoCuarto.Close
        Load frmMantoTipoCuarto
        frmMantoTipoCuarto.Show vbModal, Me
        rsTipoCuarto.Open
        Call pLlenarCboRs(cboTipoCuarto, rsTipoCuarto, 0, 1, 2)
        If Len(vgstrVarIntercam) > 0 Then
            cboTipoCuarto.ListIndex = fintLocalizaCbo(cboTipoCuarto, CStr(vgstrVarIntercam))
            vlintTipoCuarto = CInt(cboTipoCuarto.ItemData(cboTipoCuarto.ListIndex))
        Else
            cboTipoCuarto.ListIndex = fintLocalizaCbo(cboTipoCuarto, CStr(vlintTipoCuarto))
        End If
    End If
End Sub

Private Sub cboEstadoCuarto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call pEnfocaCbo(cboArea)
        Case vbKeyEscape
            Call pAgregarRegistro("")
        Case Else
            If cboEstadoCuarto.ListIndex > 0 Then
                vlintEstadoCuarto = cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex)
            End If
    End Select
End Sub

Private Sub cboTipoCuarto_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call pEnfocaCbo(cboEstadoCuarto)
        Case vbKeyEscape
            Call pAgregarRegistro("")
        Case Else
            If cboTipoCuarto.ListIndex > 0 Then
                vlintTipoCuarto = cboTipoCuarto.ItemData(cboTipoCuarto.ListIndex)
            End If
    End Select
End Sub

Private Sub cmdAnteriorRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdAnteriorRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub pLlenaBusqueda()
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    vlstrSentencia = "select adCuarto.vchNumCuarto Clave, adTipoCuarto.vchDescripcion TipoCuarto, adCuarto.vchDescripcion Cuarto, AdArea.vchDescripcion Area from adCuarto " & _
                    " inner join adArea ON adCuarto.tnyCveArea = adArea.tnyCveArea " & _
                    " inner join adTipoCuarto ON AdCuarto.tnyCveTipoCuarto = adTipoCuarto.tnyCveTipoCuarto "
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    Call pConfMshFGrid(grdhBusqueda, "|Cuarto|Tipo|Area")
    With grdhBusqueda
        .Redraw = False
        .Rows = 2
        Do While Not rs.EOF
            .Row = .Rows - 1
            .TextMatrix(.Row, 1) = rs!Clave
            .TextMatrix(.Row, 2) = rs!TipoCuarto
            .TextMatrix(.Row, 3) = rs!Area
            rs.MoveNext
            If Not rs.EOF Then .Rows = .Rows + 1
        Loop
        .Redraw = True
        If rs.RecordCount = 0 Then .Enabled = False
    End With
    rs.Close
End Sub

Private Sub cmdBuscar_Click()
    SSTObj.Tab = 1
    
    pLlenaBusqueda
    grdhBusqueda.SetFocus
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
End Sub
Private Sub cmdBuscar_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdBuscar.TabIndex)
End Sub
Private Sub cmdBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub cmdGrabarRegistro_Click()
    On Error GoTo NotificaError 'Manejo del error
    
    Dim vlblnGrabar As Boolean
    Dim vlintResul, vlintNum As Integer
'If fintBuscaCharStr(vgstrPermisoUsuarioPro, "C", True) > 0 Or fintBuscaCharStr(vgstrPermisoUsuarioPro, "E", True) > 0 Then
    vlintNum = 0
    Call pActualizaVar(Nothing)
    
    If vlintNum = 0 Then
        If fblnValidaMkText(txtDescripcion, "T", ">", 15, True) = True Then
            vlintNum = vlintNum + 1
            Call pEnfocaMkTexto(txtDescripcion)
        End If
    End If
        
    If vlintNum = 0 Then
        If cboTipoCuarto.ListIndex = -1 Then
            vlintNum = vlintNum + 1
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboTipoCuarto.ToolTipText, vbExclamation, "Mensaje")
            cboTipoCuarto.SetFocus
        End If
    End If
        
    If vlintNum = 0 Then
        If cboEstadoCuarto.ListIndex = -1 Then
            vlintNum = vlintNum + 1
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboEstadoCuarto.ToolTipText, vbExclamation, "Mensaje")
            cboEstadoCuarto.SetFocus
        End If
    End If
        
    If vlintNum = 0 Then
        If cboArea.ListIndex = -1 Then
            vlintNum = vlintNum + 1
            Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboArea.ToolTipText, vbExclamation, "Mensaje")
            cboArea.SetFocus
        End If
    End If
    
    If vlintNum = 0 Then
        If fblnValidaMkText(txtPrecio, "N", "", 10, False) = True Then
            vlintNum = vlintNum + 1
            Call pEnfocaMkTexto(txtPrecio)
        End If
    End If
    
    If cboConceptoCargo.ListIndex = -1 Then
        Call MsgBox(SIHOMsg("2") & Chr(13) & "Dato:" & cboConceptoCargo.ToolTipText, vbExclamation, "Mensaje")
        cboConceptoCargo.SetFocus
        Exit Sub
    End If
    
    If vlintNum = 0 Then
        vlintResul = MsgBox(SIHOMsg("4"), (vbYesNo + vbQuestion), "Mensaje")
        If vlintResul = vbYes Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            If vgblnNuevoRegistro = True Then
                rsCuarto.AddNew
            End If
            rsCuarto!vchNumCuarto = txtCveCuarto.Text
            rsCuarto!vchDescripcion = txtDescripcion.Text
            rsCuarto!tnyCveTipoCuarto = cboTipoCuarto.ItemData(cboTipoCuarto.ListIndex)
            rsCuarto!tnyCveEstadoCuarto = cboEstadoCuarto.ItemData(cboEstadoCuarto.ListIndex)
            rsCuarto!tnyCveArea = cboArea.ItemData(cboArea.ListIndex)
            rsCuarto!tnyCvePiso = vlintPiso
            rsCuarto!intOtroConcepto = cboConceptoCargo.ItemData(cboConceptoCargo.ListIndex)
            rsCuarto!vchObservacion = txtObservacion.Text
            rsCuarto!smyPrecio = Val(txtPrecio.Text)
            rsCuarto.Update
            vgstrVarIntercam = rsCuarto.Fields(0).Value
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            rsCuarto.Requery
'            Set grdHBusqueda.DataSource = Nothing
'            Set grdHBusqueda.DataSource = EntornoSIHO
'            grdHBusqueda.DataMember = "cmdCuarto"
            pLlenaBusqueda
            SSTObj.Tab = 0 'Se localiza en el primer tabulador para la alta
            vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
            vgintTipoOrd = 1    'Que tipo de ordenamiento realizará de inicio en el grdHBusqueda
            vgintColLoc = 1 'Localiza la búsqueda de registros para la primera columna del grdHBusqueda
            vgstrVarIntercam = ""

            vlintColCve = 1
            Call pConfMshFGrid(grdhBusqueda, "|Cuarto|Tipo|Area")
            Call pAgregarRegistro("")
        End If
    End If
'Else
'    Call MsgBox(SIHOMsg("65"), vbInformation, "Mensaje")
'End If
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

Private Sub cmdGrabarRegistro_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdGrabarRegistro.TabIndex)
End Sub

Private Sub cmdGrabarRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    Call pAgregarRegistro("")
End Sub
Private Sub cmdGrabarRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub cmdPrimerRegistro_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdPrimerRegistro.TabIndex)
End Sub

Private Sub cmdPrimerRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub cmdSiguienteRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdSiguienteRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub cmdUltimoRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdUltimoRegistro_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            pAgregarRegistro ("")
    End Select
End Sub

Private Sub txtPrecio_LostFocus()
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

Private Sub cmdPrimerRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdSiguienteRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub

Private Sub cmdUltimoRegistro_KeyDown(KeyCode As Integer, Shift As Integer)
    pPresionoEscape (KeyCode)
End Sub
Private Sub pVerificaPosTab(vlintPosIndex As Integer)
    'Procedimiento para verificar la posicion de un control dentro del sstab
    If SSTObj.Tab = 0 Then
        Select Case vlintPosIndex
            Case 22 To 24 'En el tab 1
                If txtCveCuarto.Enabled = True Then
                    Call pEnfocaMkTexto(txtCveCuarto)
                Else
                    If txtDescripcion.Enabled = True Then
                        Call pEnfocaMkTexto(txtDescripcion)
                    End If
                End If
        End Select
    End If
    If SSTObj.Tab = 1 Then
        Select Case vlintPosIndex
            Case 0 To 22 'En el tab 0
                grdhBusqueda.SetFocus
        End Select
    End If
End Sub
