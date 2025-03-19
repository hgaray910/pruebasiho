VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMantoEstadoCuarto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Estados de Cuarto"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   3240
      Left            =   -30
      TabIndex        =   10
      Top             =   -60
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   5715
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   26
      TabMaxWidth     =   26
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoEstadoCuarto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "shp0(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "shp1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDescripcion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCve"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "shp3(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "shp2(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCveEstadoCuarto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescripcion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdPrimerRegistro"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAnteriorRegistro"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdBuscar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdSiguienteRegistro"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdUltimoRegistro"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdGrabarRegistro"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoEstadoCuarto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTitulo"
      Tab(1).Control(1)=   "grdHBusqueda"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdGrabarRegistro 
         Height          =   495
         Left            =   5190
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guardar el registro"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdUltimoRegistro 
         Height          =   495
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":01DA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ultimo registro"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSiguienteRegistro 
         Height          =   495
         Left            =   4170
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":034C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Siguiente registro"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   495
         Left            =   3660
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":04BE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Búsqueda"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdAnteriorRegistro 
         Height          =   495
         Left            =   3150
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":0630
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Anterior registro"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrimerRegistro 
         Height          =   495
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoEstadoCuarto.frx":07A2
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Primer registro"
         Top             =   2340
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin MSMask.MaskEdBox txtDescripcion 
         Height          =   315
         Left            =   1995
         TabIndex        =   1
         ToolTipText     =   "Descripción del estado del cuarto"
         Top             =   1275
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   30
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtCveEstadoCuarto 
         Height          =   315
         Left            =   1995
         TabIndex        =   0
         ToolTipText     =   "Clave del estado del cuarto"
         Top             =   870
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
         DragIcon        =   "frmMantoEstadoCuarto.frx":0914
         Height          =   2175
         Left            =   -74430
         TabIndex        =   11
         Top             =   600
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   3836
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
      Begin VB.Shape shp2 
         BorderColor     =   &H8000000C&
         Height          =   1485
         Index           =   2
         Left            =   165
         Top             =   480
         Width           =   8145
      End
      Begin VB.Shape shp3 
         BorderColor     =   &H80000005&
         Height          =   1485
         Index           =   3
         Left            =   180
         Top             =   495
         Width           =   8145
      End
      Begin VB.Label lblCve 
         Caption         =   "Clave"
         Height          =   255
         Left            =   405
         TabIndex        =   8
         ToolTipText     =   "Número del piso"
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   405
         TabIndex        =   9
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Shape shp1 
         BorderColor     =   &H80000005&
         Height          =   615
         Index           =   1
         Left            =   2580
         Top             =   2280
         Width           =   3180
      End
      Begin VB.Shape shp0 
         BorderColor     =   &H8000000C&
         Height          =   615
         Index           =   0
         Left            =   2565
         Top             =   2265
         Width           =   3180
      End
      Begin VB.Label lblTitulo 
         Caption         =   "Búsqueda de estados de cuarto"
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
         Left            =   -74430
         TabIndex        =   12
         Top             =   210
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmMantoEstadoCuarto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Admisión
'| Nombre del Formulario    : frmMantoEstadoCuarto
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de estados de cuarto
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Nery Lozano - Luis Astudillo
'| Autor                    : Nery Lozano - Luis Astudillo
'| Fecha de Creación        : 26/Noviembre/1999
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables

Dim vgblnNuevoRegistro As Boolean
Private Sub cmdGrabarRegistro_GotFocus()
    Call pVerificaPosTab(cmdGrabarRegistro.TabIndex)
End Sub

Private Sub Form_Load()
'-------------------------------------------------------------------------------------------
'  Define implicitamente el RS al abrirlo y lo relaciona con el objeto cmd del
'  Data Environment. Relaciona el grdHBusqueda con el DataEnvironment
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    If vgblnNuevoRegistro = True Then
        EntornoSIHO.rscmdEstadoCuarto.Open 'Abre la conexión con la tabla utilizando un RS
    End If
    Set grdHBusqueda.DataSource = EntornoSIHO
    grdHBusqueda.DataMember = "cmdEstadoCuarto"
    
    vgstrNombreForm = Me.Name 'Nombre del formulario que se utiliza actualmente
    vgblnExistioError = False 'Inicia la bandera sin errores
    vgblnErrorIngreso = False

    vgstrVarIntercam = ""
    sstObj.Tab = 0 'Se localiza en el primer tabulador para la alta
    vgstrAcumTextoBusqueda = "" 'Limpia el contenedor de busqueda
    vgintTipoOrd = 1 'Que tipo de ordenamiento realizará de inicio en el grdHBusqueda
    vgintColLoc = 1 'Localiza la búsqueda de registros para la primera columna del grdHBusqueda
    
    Call pConfMshFGrid(grdHBusqueda, "|Clave|Descripción")
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
Private Sub Form_Activate()
'-------------------------------------------------------------------------------------------
' Permite enfocar el primer control del formulario
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    If vgblnNuevoRegistro = False Then
        Call pEnfocaMkTexto(txtCveEstadoCuarto)
    End If
    
NotificaError:
    If vgblnExistioError Then
        vgblnExistioError = False
        Unload Me
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
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
            vlintNumero = fintLocalizaPkRs(EntornoSIHO.rscmdEstadoCuarto, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
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
        Call pOrdColMshFGrid(grdHBusqueda, vgintTipoOrd)
        Call pDesSelMshFGrid(grdHBusqueda)
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
Private Sub grdHBusqueda_GotFocus()
    Call pVerificaPosTab(grdHBusqueda.TabIndex)
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
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
                vlintNumero = fintLocalizaPkRs(EntornoSIHO.rscmdEstadoCuarto, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1))
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
Private Sub sstObj_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
' Validación del <Escape> en el segundo Tab del sstObj(Búsqueda) cuando no tiene el enfoque el Grid
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Select Case KeyCode
        Case vbKeyEscape
            sstObj.Tab = 0
            pPresionoEscape (KeyCode)
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
Private Sub txtCveEstadoCuarto_GotFocus()
'-------------------------------------------------------------------------------------------
' Seleccionar el cuadro de texto del primer control del formulario
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Call pActualizaVar(Nothing)
    Call pEnfocaMkTexto(txtCveEstadoCuarto)
    Call pVerificaPosTab(txtCveEstadoCuarto.TabIndex)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveEstadoCuarto_GotFocus"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub txtCveEstadoCuarto_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim vlbytNumero As Byte
    Select Case KeyCode
        Case vbKeyReturn
            'Buscar criterio
            If (Len(txtCveEstadoCuarto.Text) <= 0) Then
                txtCveEstadoCuarto.Text = "0"
            End If
            If fintSigNumRs(EntornoSIHO.rscmdEstadoCuarto, 0) = CByte(txtCveEstadoCuarto.Text) Then
                txtCveEstadoCuarto.Enabled = False
                txtDescripcion.Enabled = True
                txtDescripcion.SetFocus
            Else
                vlbytNumero = fintLocalizaPkRs(EntornoSIHO.rscmdEstadoCuarto, 0, txtCveEstadoCuarto.Text)
                If vlbytNumero > 0 Then
                    pModificaRegistro
                Else
                    Call MsgBox(SIHOMsg("12"), vbExclamation, "Mensaje")
                    pAgregarRegistro ("")
                    Call pEnfocaMkTexto(txtCveEstadoCuarto)
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
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCveEstadoCuarto_KeyDown"))
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
    
    txtCveEstadoCuarto.Text = CStr(fintSigNumRs(EntornoSIHO.rscmdEstadoCuarto, 0)) 'Muestra el siguiente consecutivo del campo Clave
    txtCveEstadoCuarto.Enabled = True 'Habilita el ingreso de una clave para su búsqueda
    txtDescripcion.Text = ""
    txtDescripcion.Enabled = False
    
    pHabilitaBotonBuscar True
    If vlstrSelTxt <> "I" Then
        Call pEnfocaMkTexto(txtCveEstadoCuarto)
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
    txtCveEstadoCuarto.Text = EntornoSIHO.rscmdEstadoCuarto.Fields(0).Value
    txtDescripcion.Text = EntornoSIHO.rscmdEstadoCuarto.Fields(1).Value
    txtDescripcion.Enabled = True
    txtCveEstadoCuarto.Enabled = False
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
            cmdGrabarRegistro_Click
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
    Call pPosicionaRegRs(EntornoSIHO.rscmdEstadoCuarto, "A")
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
    Call pRefrescaMshFGrid(grdHBusqueda, EntornoSIHO.rscmdEstadoCuarto.RecordCount)
    Call pConfMshFGrid(grdHBusqueda, "|Clave|Descripción")
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
If fintBuscaCharStr(vgstrPermisoUsuarioPro, "C", True) > 0 Or fintBuscaCharStr(vgstrPermisoUsuarioPro, "E", True) > 0 Then
    Call pActualizaVar(Nothing)
    If fblnValidaMkText(txtDescripcion, "T", ">", 30, True) = False Then
        If fblnExisteCriterioRs(EntornoSIHO.rscmdEstadoCuarto, 0, 1, txtDescripcion, txtCveEstadoCuarto) = True Then
            Call MsgBox(SIHOMsg("19") & Chr(13) & "Dato:" & txtDescripcion.ToolTipText, vbExclamation, "Mensaje")
            Call pEnfocaMkTexto(txtDescripcion)
        Else
            If (Len(txtDescripcion.Text) > 0) Then
                EntornoSIHO.ConeccionSIHO.BeginTrans
                If vgblnNuevoRegistro = True Then
                    EntornoSIHO.rscmdEstadoCuarto.AddNew
                End If
                EntornoSIHO.rscmdEstadoCuarto.Fields(1).Value = txtDescripcion.Text
                EntornoSIHO.rscmdEstadoCuarto.Update
                If vgblnNuevoRegistro = True Then
                    vgstrVarIntercam = CStr(EntornoSIHO.rscmdEstadoCuarto.Fields(0).Value)
                End If
                EntornoSIHO.ConeccionSIHO.CommitTrans
                pAgregarRegistro ("")
            End If
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
    Call pPosicionaRegRs(EntornoSIHO.rscmdEstadoCuarto, "I")
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
    Call pPosicionaRegRs(EntornoSIHO.rscmdEstadoCuarto, "S")
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
    Call pPosicionaRegRs(EntornoSIHO.rscmdEstadoCuarto, "U")
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
        .ColWidth(2) = 5850
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
    EntornoSIHO.rscmdEstadoCuarto.Close

End Sub
Private Sub pActualizaVar(ObjTxt As MaskEdBox)
    txtDescripcion.Text = fstrFormatTxt(txtDescripcion.Text, "T", ">", 30, True)
    If Not ObjTxt Is Nothing Then
        Call pEnfocaMkTexto(ObjTxt)
    End If
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

Private Sub cmdAnteriorRegistro_GotFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub cmdBuscar_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(cmdBuscar.TabIndex)
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
Private Sub sstObj_GotFocus()
    Call pActualizaVar(Nothing)
    Call pVerificaPosTab(sstObj.TabIndex)
End Sub


Private Sub txtDescripcion_LostFocus()
    Call pActualizaVar(Nothing)
End Sub

Private Sub pVerificaPosTab(vlintPosIndex As Integer)
    'Procedimiento para verificar la posicion de un control dentro del sstab
    If sstObj.Tab = 0 Then
        Select Case vlintPosIndex
            Case 10 To 12 'En el tab 1
                If txtCveEstadoCuarto.Enabled = True Then
                    Call pEnfocaMkTexto(txtCveEstadoCuarto)
                Else
                    If txtDescripcion.Enabled = True Then
                        Call pEnfocaMkTexto(txtDescripcion)
                    End If
                End If
        End Select
    End If
    If sstObj.Tab = 1 Then
        Select Case vlintPosIndex
            Case 0 To 10 'En el tab 0
                grdHBusqueda.SetFocus
        End Select
    End If
End Sub
