VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConsultaRequisMix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Requisiciones de Artículos"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstObj 
      Height          =   8250
      Left            =   -30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   -345
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   14552
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmConsultaRequisMix.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmArticulosCD"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmRequisicion"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmConsultar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame frmConsultar 
         Caption         =   "Consultar por"
         Height          =   1980
         Left            =   330
         TabIndex        =   15
         Top             =   495
         Width           =   10395
         Begin VB.ComboBox cboTipoRequis 
            Height          =   315
            ItemData        =   "frmConsultaRequisMix.frx":001C
            Left            =   1650
            List            =   "frmConsultaRequisMix.frx":0029
            TabIndex        =   5
            ToolTipText     =   "Tipo de requisición"
            Top             =   1545
            Width           =   5355
         End
         Begin VB.OptionButton optFolio 
            Caption         =   "Folio"
            Height          =   240
            Left            =   2925
            TabIndex        =   1
            ToolTipText     =   "Entradas/salidas por folio"
            Top             =   165
            Width           =   675
         End
         Begin VB.OptionButton optFecha 
            Caption         =   "Fecha"
            Height          =   240
            Left            =   1710
            TabIndex        =   0
            ToolTipText     =   "Entradas/salidas por fecha"
            Top             =   165
            Width           =   855
         End
         Begin VB.ComboBox cboEmpleado 
            Height          =   315
            Left            =   1650
            TabIndex        =   3
            ToolTipText     =   "Empleado que realiza la requisición"
            Top             =   840
            Width           =   5355
         End
         Begin VB.ComboBox cboDepartamento 
            Height          =   315
            Left            =   1650
            TabIndex        =   2
            ToolTipText     =   "Departamento que realiza la requisición"
            Top             =   480
            Width           =   5355
         End
         Begin VB.ComboBox cboEstatus 
            Height          =   315
            Left            =   1650
            TabIndex        =   4
            ToolTipText     =   "Estatus de la requisición"
            Top             =   1185
            Width           =   5355
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   315
            Left            =   9165
            TabIndex        =   7
            ToolTipText     =   "Fecha final "
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFechaInicio 
            Height          =   315
            Left            =   7635
            TabIndex        =   6
            ToolTipText     =   "Fecha inicial "
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            ClipMode        =   1
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFolioFin 
            Height          =   315
            Left            =   9165
            TabIndex        =   9
            ToolTipText     =   "Folio final "
            Top             =   1185
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFolioIni 
            Height          =   315
            Left            =   7635
            TabIndex        =   8
            ToolTipText     =   "Folio inicial "
            Top             =   1185
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label lblTipoRequis 
            Caption         =   "Tipo de requisición"
            Height          =   180
            Left            =   270
            TabIndex        =   25
            Top             =   1605
            Width           =   1395
         End
         Begin VB.Label lblRangoA 
            Caption         =   "A:"
            Height          =   180
            Left            =   8865
            TabIndex        =   24
            Top             =   1245
            Width           =   270
         End
         Begin VB.Label lblRangoDe 
            Caption         =   "De:"
            Height          =   210
            Left            =   7185
            TabIndex        =   23
            Top             =   1230
            Width           =   330
         End
         Begin VB.Label lblRangoFolio 
            Caption         =   "Rango de folio:"
            Height          =   225
            Left            =   7185
            TabIndex        =   22
            Top             =   915
            Width           =   1170
         End
         Begin VB.Label lblRangoFecha 
            Caption         =   "Rango de fecha:"
            Height          =   225
            Left            =   7215
            TabIndex        =   21
            Top             =   165
            Width           =   1290
         End
         Begin VB.Label lblA 
            Caption         =   "A:"
            Height          =   195
            Left            =   8865
            TabIndex        =   20
            Top             =   540
            Width           =   225
         End
         Begin VB.Label lblDe 
            Caption         =   "De:"
            Height          =   195
            Left            =   7215
            TabIndex        =   19
            Top             =   540
            Width           =   285
         End
         Begin VB.Label lblEmpleadoCD 
            Caption         =   "Empleado"
            Height          =   225
            Left            =   270
            TabIndex        =   18
            Top             =   915
            Width           =   780
         End
         Begin VB.Label lblDepartamentoGrid 
            Caption         =   "Departamento"
            Height          =   225
            Left            =   270
            TabIndex        =   17
            Top             =   525
            Width           =   1065
         End
         Begin VB.Label lblEstatusGrid 
            Caption         =   "Estatus"
            Height          =   225
            Left            =   270
            TabIndex        =   16
            Top             =   1290
            Width           =   630
         End
      End
      Begin VB.Frame frmRequisicion 
         Caption         =   "Requisiciones"
         Height          =   2715
         Left            =   330
         TabIndex        =   14
         Top             =   2565
         Width           =   10395
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhBusqueda 
            Height          =   2340
            Left            =   240
            TabIndex        =   10
            ToolTipText     =   "Requisiciones "
            Top             =   225
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   4128
            _Version        =   393216
            Cols            =   11
            BackColorBkg    =   16777215
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            HighLight       =   0
            MergeCells      =   1
            FormatString    =   "|Requisición||Departamento||Empleado|Estatus|Urgente|Fecha|Hora"
            _NumberOfBands  =   1
            _Band(0).Cols   =   11
         End
      End
      Begin VB.Frame frmArticulosCD 
         Caption         =   "Artículos de la requisición"
         Height          =   2715
         Left            =   330
         TabIndex        =   13
         Top             =   5310
         Width           =   10395
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHRequisicion 
            Height          =   2340
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Artículos de la requisición "
            Top             =   240
            Width           =   9930
            _ExtentX        =   17515
            _ExtentY        =   4128
            _Version        =   393216
            Cols            =   6
            BackColorBkg    =   16777215
            GridColor       =   12632256
            AllowBigSelection=   0   'False
            HighLight       =   0
            MergeCells      =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
      End
   End
End
Attribute VB_Name = "frmConsultaRequisMix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Consulta de Requisiciones de artículos para el Módulo de Inventarios'
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : prjInventario
'| Nombre del Formulario    : frmConsultaRequisMix
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza la Consulta de Requisiciones de Artículos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Luis Astudillo - Inés Saláis
'| Autor                    : Luis Astudillo - Inés Saláis
'| Fecha de Creación        : 8/Junio/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables

Private vgintCveDepto As Integer 'Guarda la clave del departamento seleccionado
Private vgintCveEmp As Integer 'Guarda la clave del empleado seleccionado
Private vgstrEstatus As String 'Guarda el estatus seleccionado
Private vgstrTipoRequis As String 'Guarda la abreviatura del tipo de requisición seleccionado
Dim rsIvRequisicionMaestro As New ADODB.recordSet
Dim vlstrSQL As String

Private Sub pAbrirTablas()
'------------------------------------------------------------------------------------------
' Abre las tablas necesarias
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    vlstrSQL = "select * from IVREQUISICIONMAESTRO"
    Set rsIvRequisicionMaestro = frsRegresaRs(vlstrSQL, adLockOptimistic, adOpenDynamic)
    
'   Call EntornoSIHO.cmdIvRequisicionMaestro 'Abre la conexión con la tabla de requisiciones maestro
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAbrirTablas"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pConfFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
'------------------------------------------------------------------------------------------
' Configura el Grid de la requisición
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintseq, vlintLargo As Integer

    If ObjGrid.Rows > 0 Then
        ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas

        With ObjGrid

            .ColWidth(0) = 300 'cabecera
            .ColWidth(1) = 1030 'Clave del articulo
            .ColWidth(2) = 720 'Cantidad
            .ColWidth(3) = 6000 'Nombre Comercial
            .ColWidth(4) = 1600 'Estatus
            .ColWidth(5) = 0 'Unidad de control
            .ScrollBars = flexScrollBarBoth
        End With
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfFGrid"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pConfFGridBus(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
'------------------------------------------------------------------------------------------
' Configura el Grid de búsqueda
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintseq, vlintLargo As Integer

    If ObjGrid.Rows > 0 Then
        ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas

        With ObjGrid

            .ColWidth(0) = 300 'cabecera de fila
            .ColWidth(1) = 950 'requisición
            .ColWidth(2) = 2200 'Tipo de requisicion
            .ColWidth(3) = 0 'clave del departamento
            .ColWidth(4) = 2500 'departamento
            .ColWidth(5) = 0 'clave del empleado
            .ColWidth(6) = 3500 'empleado
            .ColWidth(7) = 2000 'estatus
            .ColWidth(8) = 720 'urgente
            .ColWidth(9) = 1200 'fecha
            .ColWidth(10) = 0 'hora

            For vlintseq = 1 To ObjGrid.Rows - 1
                If .TextMatrix(vlintseq, 8) = True Then
                    .TextMatrix(vlintseq, 8) = "SI"
                Else
                    .TextMatrix(vlintseq, 8) = "NO"
                End If
                
                .TextMatrix(vlintseq, 9) = UCase(Format(.TextMatrix(vlintseq, 9), "DD/MMM/YYYY"))
                
                If .TextMatrix(vlintseq, 2) = "P" Then
                    .TextMatrix(vlintseq, 2) = "CARGO A PACIENTE"
                End If
                If .TextMatrix(vlintseq, 2) = "R" Then
                    .TextMatrix(vlintseq, 2) = "REUBICACION"
                End If
                If .TextMatrix(vlintseq, 2) = "D" Then
                    .TextMatrix(vlintseq, 2) = "SALIDA A DEPARTAMENTO"
                End If
                If .TextMatrix(vlintseq, 2) = "A" Then
                    .TextMatrix(vlintseq, 2) = "ALMACEN ABASTECIMIENTO"
                End If
                
            Next vlintseq


            .ScrollBars = flexScrollBarBoth

        End With
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfFGridBus"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pDeshabilitaStabuno()
'------------------------------------------------------------------------------------------
'Deshabilita los controles del primer stab
'------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    grdhBusqueda.Enabled = False
    grdHRequisicion.Enabled = False
    
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDeshabilitaStabuno"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub


Private Sub cboDepartamento_Click()
'---------------------------------------------------------------------------------------
' Evalúa que hay seleccionado en los combos departamento, empleado, estatus y tipo de requisición para hacer
' el filtrado en el grid
'---------------------------------------------------------------------------------------
    On Error GoTo NotificaError
        
    pLlenarCboEmpleado
    If (txtFechaInicio.Text <> "  /  /    " And txtFechaFin.Text <> "  /  /    ") Or (Len(txtFolioIni) > 0 And Len(txtFolioFin) > 0) Then
        pRevisaParametros
        pVerificaExistencia
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Cbodepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            cboEmpleado.SetFocus

        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboDepartamento_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cboEmpleado_Click()
'---------------------------------------------------------------------------------------
' Evalúa que hay seleccionado en los combos departamento, empleado, estatus y tipo de requisición para hacer
' el filtrado en el grid
'---------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If (txtFechaInicio.Text <> "  /  /    " And txtFechaFin.Text <> "  /  /    ") Or (Len(txtFolioIni) > 0 And Len(txtFolioFin) > 0) Then
        pRevisaParametros
        pVerificaExistencia
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub


Private Sub cboEmpleado_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            cboEstatus.SetFocus

        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEmpleado_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cboEstatus_Click()
'---------------------------------------------------------------------------------------
' Evalúa que hay seleccionado en los combos departamento, empleado, estatus y tipo de requisición para hacer
' el filtrado en el grid
'---------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If (txtFechaInicio.Text <> "  /  /    " And txtFechaFin.Text <> "  /  /    ") Or (Len(txtFolioIni) > 0 And Len(txtFolioFin) > 0) Then
        pRevisaParametros
        pVerificaExistencia
    End If
        
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstatus_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cboEstatus_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            cboTipoRequis.SetFocus
            
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboEstatus_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Function fblnValidaFechaFin() As Boolean
'--------------------------------------------------------------------------
' Función que valida el ingreso de la fecha final
'--------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlstrMensaje As String
    fblnValidaFechaFin = True
    If txtFechaFin = "  /  /    " Then
        vlstrMensaje = SIHOMsg("2") & Chr(13) & "Dato: " & txtFechaFin.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
        fblnValidaFechaFin = False
        txtFechaFin.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
        Call pEnfocaMkTexto(txtFechaFin)
    Else
        If fblnValidaFecha(txtFechaFin) Then
            If CDate(txtFechaFin.Text) > CDate(fdtmServerFecha) Then
                Call MsgBox("La fecha final de consulta debe ser menor o igual a la fecha del sistema", vbExclamation, "Mensaje")
                fblnValidaFechaFin = False
                 txtFechaFin.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
                Call pEnfocaMkTexto(txtFechaFin)
            Else
                If CDate(txtFechaFin.Text) < CDate(txtFechaInicio.Text) Then
                    Call MsgBox("La fecha final de consulta debe ser mayor o igual a la fecha de inicio", vbExclamation, "Mensaje")
                    fblnValidaFechaFin = False
                    txtFechaFin.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
                    Call pEnfocaMkTexto(txtFechaFin)
                Else
                    'La fecha es correcta
                End If
            End If
        Else
            vlstrMensaje = SIHOMsg("29") & Chr(13) '"!Fecha no válida!, formato de fecha dd/mm/aaaa"
            Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
            fblnValidaFechaFin = False
            txtFechaFin.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
            Call pEnfocaMkTexto(txtFechaFin)
        End If
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Function
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaFechaFin"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Function
        End If
    End If
End Function


Private Function fblnValidaFechaInicio()
'--------------------------------------------------------------------------
' Función que valida el ingreso de la fecha inicial
'--------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlstrMensaje As String

    fblnValidaFechaInicio = True
    If txtFechaInicio = "  /  /    " Then
        vlstrMensaje = SIHOMsg("2") & Chr(13) & "Dato: " & txtFechaInicio.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
        fblnValidaFechaInicio = False
        txtFechaInicio.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
        Call pEnfocaMkTexto(txtFechaInicio)
    Else
        If fblnValidaFecha(txtFechaInicio) Then
            If CDate(txtFechaInicio.Text) > CDate(fdtmServerFecha) Then
                Call MsgBox("La fecha inicial de consulta debe ser menor o igual a la fecha del sistema", vbExclamation, "Mensaje")
                fblnValidaFechaInicio = False
                txtFechaInicio.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
                Call pEnfocaMkTexto(txtFechaInicio)
            Else
                'La fecha de inicio es correcta
            End If
        Else
            vlstrMensaje = SIHOMsg("29") & Chr(13)
            Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
             fblnValidaFechaInicio = False
            txtFechaInicio.Text = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
            Call pEnfocaMkTexto(txtFechaInicio)

        End If
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Function
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaFechaInicio"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Function
        End If
    End If
End Function

Private Function fblnValidaFolioFinal()
'--------------------------------------------------------------------------
' Función que valida el ingreso del folio final
'--------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlstrMensaje As String

    fblnValidaFolioFinal = True
    If Len(txtFolioFin) = 0 Then
        vlstrMensaje = SIHOMsg("2") & Chr(13) & "Dato: " & txtFolioFin.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
        fblnValidaFolioFinal = False
        Call pEnfocaMkTexto(txtFolioFin)
    Else
        If txtFolioFin.Text = 0 Then
            txtFolioFin.Text = 1
        End If
        fblnValidaFolioFinal = True
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Function
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaFolioFinal"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Function
        End If
    End If
End Function

Private Function fblnValidaFolioInicio()
'--------------------------------------------------------------------------
' Función que valida el ingreso del folio inicial
'--------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlstrMensaje As String

    fblnValidaFolioInicio = True
    If Len(txtFolioIni) = 0 Then
        vlstrMensaje = SIHOMsg("2") & Chr(13) & "Dato: " & txtFolioIni.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
        fblnValidaFolioInicio = False
        Call pEnfocaMkTexto(txtFolioIni)
    Else
        If txtFolioIni.Text = 0 Then
            txtFolioIni.Text = 1
        End If
        fblnValidaFolioInicio = True
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Function
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValidaFolioInicio"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Function
        End If
    End If
End Function

Private Sub cboTipoRequis_Click()
'---------------------------------------------------------------------------------------
' Evalúa que hay seleccionado en los combos departamento, empleado, estatus y tipo de requisición para hacer
' el filtrado en el grid
'---------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If (txtFechaInicio.Text <> "  /  /    " And txtFechaFin.Text <> "  /  /    ") Or (Len(txtFolioIni) > 0 And Len(txtFolioFin) > 0) Then
        pRevisaParametros
        pVerificaExistencia
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoRequis_Click"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cboTipoRequis_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            If txtFechaInicio.Enabled = True Then
                Call pEnfocaMkTexto(txtFechaInicio)
            Else
                Call pEnfocaMkTexto(txtFolioIni)
            End If

        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboTipoRequis_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Activate()
    optFecha.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim vlstrx As String
    
    vlstrx = "select count(*) from IvRequisicionDepartamento where chrTipoRequisicion = 'AG' "
    If frsRegresaRs(vlstrx).Fields(0) > 0 Then
        cboTipoRequis.AddItem "ALMACEN ABASTECIMIENTO"
    End If
        
    pAbrirTablas
    pLlenarCboDepartamento
    
    pLlenarCboEstatus
    pDeshabilitaStabuno
    Call pIniciaMshFGrid(grdhBusqueda)
    Call pLimpiaMshFGrid(grdhBusqueda)
    Call pIniciaMshFGrid(grdHRequisicion)
    Call pLimpiaMshFGrid(grdHRequisicion)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pLlenarCboDepartamento()
'----------------------------------------------------------------------------------------
' Llena el combo del departamento para poder filtrar las requisiciones
'----------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    Dim rsDepartamento As ADODB.recordSet
    
    vlstrSQL = "select * from NoDepartamento"
    Set rsDepartamento = frsRegresaRs(vlstrSQL, adLockOptimistic, adOpenDynamic)
    
    Call pLlenarCboRs(cboDepartamento, rsDepartamento, 0, 1, -1)
    rsDepartamento.Close
    
    If fblnRevisaPermiso(vglngNumeroLogin, 1020, "C") Then
        cboDepartamento.ListIndex = -1    'se posiciona en el primer dato del combo
        cboDepartamento.Enabled = True
        Call pLlenarCboEmpleado
    Else
        cboDepartamento.ListIndex = fintLocalizaCbo(cboDepartamento, CStr(vgintNumeroDepartamento)) 'se posiciona en el depto con el que se dio login
        cboDepartamento.Enabled = False
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboDepartamento"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pLlenarCboEmpleado()
'-------------------------------------------------------------------------------------------------
' Llena el combo del empleado con todos los empleados que corresponden al departamento
'-------------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintCveDepto As Integer
    
    Dim rsIVSelEmpxDpto As New ADODB.recordSet
    Dim rsIvSelEmpleadosTodos As New ADODB.recordSet
    
    If cboDepartamento.ListIndex <> -1 Then 'Se filtran los empleados por departamento
        vlintCveDepto = cboDepartamento.ItemData(cboDepartamento.ListIndex)
        
        Set rsIVSelEmpxDpto = frsEjecuta_SP(Trim(Str(vlintCveDepto)), "sp_IVSelEmpxDpto")
        
        If rsIVSelEmpxDpto.RecordCount > 0 Then
            Call pLlenarCboRs(cboEmpleado, rsIVSelEmpxDpto, 0, 1, -1)
            cboEmpleado.ListIndex = -1    'se posiciona en el primer dato del combo
        Else
            cboEmpleado.Clear
        End If
        rsIVSelEmpxDpto.Close 'Cierra la conexión con la tabla de empleados
    Else
        Set rsIvSelEmpleadosTodos = frsEjecuta_SP("", "sp_IvSelEmpleadosTodos")

        If rsIvSelEmpleadosTodos.RecordCount > 0 Then
            Call pLlenarCboRs(cboEmpleado, rsIvSelEmpleadosTodos, 0, 1, -1)
            cboEmpleado.ListIndex = -1    'se posiciona en el primer dato del combo
            cboEmpleado.Enabled = True
        Else
            cboEmpleado.Clear
            cboEmpleado.Enabled = False
        End If
        rsIvSelEmpleadosTodos.Close
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboEmpleado"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
Private Sub pLlenarCboEstatus()
'----------------------------------------------------------------------------------------
' Llena el combo del estatus para poder filtrar las requisiciones
'----------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    cboEstatus.AddItem "PENDIENTE", 0
    cboEstatus.AddItem "CANCELADA", 1
    cboEstatus.AddItem "SURTIDA", 2
    cboEstatus.AddItem "SURTIDA PARCIAL", 3

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboEstGrid"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pLimpiaPantalla()
    txtFechaInicio.Text = "  /  /    "
    txtFechaFin.Text = "  /  /    "
    txtFolioIni.Text = " "
    txtFolioFin.Text = " "
    cboEmpleado.ListIndex = -1
    cboEmpleado.Text = ""
    cboEstatus.ListIndex = -1
    cboEstatus.Text = ""
    cboTipoRequis.ListIndex = -1
    cboTipoRequis.Text = ""
    If fblnRevisaPermiso(vglngNumeroLogin, 1020, "C") Then
        cboDepartamento.ListIndex = -1
        cboDepartamento.Text = ""
    End If
    Call pIniciaMshFGrid(grdhBusqueda)
    Call pLimpiaMshFGrid(grdhBusqueda)
    Call pIniciaMshFGrid(grdHRequisicion)
    Call pLimpiaMshFGrid(grdHRequisicion)
End Sub

Private Sub pMuestraRegistro()
'-------------------------------------------------------------------------------------------
' Permite realizar la consulta de la descripción de un registro al teclear el número de
' requisición en el txtNumRequisCD
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumRequis As Long
    
    Dim rsIVSelRequisDetalleDatos As New ADODB.recordSet
    'Datos del Maestro
    vlintNumRequis = rsIvRequisicionMaestro.Fields(0)
    
    'Datos del Detalle
    
    Set rsIVSelRequisDetalleDatos = frsEjecuta_SP(Trim(Str(vlintNumRequis)), "sp_IVSelRequisDetalleDatos")
    
    If rsIVSelRequisDetalleDatos.RecordCount <> 0 Then
        Call pLlenarMshFGrdRs(grdHRequisicion, rsIVSelRequisDetalleDatos)
    End If
    Call pConfFGrid(grdHRequisicion, "|Clave|Cantidad|Nombre Comercial|Estatus|Unidad de control")
    rsIVSelRequisDetalleDatos.Close 'Cierra la conexión con el select del detalle
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraRegistro"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub pRevisaParametros()
'---------------------------------------------------------------------------------------
' Verifica cuales seran los parámetros para el comando que filtra las recepciones
'---------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    'parámetros para el comando que filtra las requisiciones
    If cboEmpleado.ListIndex = -1 Then
        vgintCveEmp = 0
    Else
        vgintCveEmp = cboEmpleado.ItemData(cboEmpleado.ListIndex)
    End If

    If cboDepartamento.ListIndex = -1 Then
        vgintCveDepto = 0
    Else
        vgintCveDepto = cboDepartamento.ItemData(cboDepartamento.ListIndex)
    End If

    If cboEstatus.ListIndex = -1 Then
        vgstrEstatus = ""
    Else
        vgstrEstatus = cboEstatus.Text
    End If


    If cboTipoRequis.ListIndex = -1 Then
        vgstrTipoRequis = ""
    Else
        If cboTipoRequis.ListIndex = 0 Then
            vgstrTipoRequis = "P" 'Cargos a paciente
        End If
        If cboTipoRequis.ListIndex = 1 Then
            vgstrTipoRequis = "R" 'Reubicación
        End If
        If cboTipoRequis.ListIndex = 2 Then
            vgstrTipoRequis = "D" 'Salidas a departamento
        End If
        If cboTipoRequis.ListCount > 2 Then
          If cboTipoRequis.ListIndex = 3 Then
            vgstrTipoRequis = "A" 'Almacén Abastecimiento
          End If
        End If
        
    End If
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRevisaParametros"))
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
    Set grdhBusqueda.DataSource = Nothing

    rsIvRequisicionMaestro.Close 'Cierra la conexión con la tabla de requisiciones maestro

    Screen.MousePointer = vbDefault

    Unload Me
End Sub

Private Sub pVerificaExistencia()
'-------------------------------------------------------------------------------------------
' Abre el comando correspondiente al filtro seleccionado y lo asigna a la consulta
' de requisiciones de artículos
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintFolioInicial As Long
    Dim vlintFolioFinal As Long
    Dim vldtmFechaInicial As Date
    Dim vldtmFechaFinal As Date

    Dim rsIvSelFilRequisMaesxMixtas As New ADODB.recordSet

    If Len(txtFolioIni) > 0 Then
        vlintFolioInicial = txtFolioIni
    Else
        vlintFolioInicial = 0
    End If
    
    If Len(txtFolioFin) > 0 Then
        vlintFolioFinal = txtFolioFin
    Else
        vlintFolioFinal = 0
    End If
    
    If txtFechaInicio.Text <> "  /  /    " Then
        vldtmFechaInicial = txtFechaInicio.Text
    End If
    
    If txtFechaFin.Text <> "  /  /    " Then
        vldtmFechaFinal = txtFechaFin.Text
    End If

        
    'Solo cuando por lo menos una opción tenga valor
    If vgintCveEmp > 0 Or vgintCveDepto > 0 Or Len(vgstrEstatus) > 0 Or Len(vgstrTipoRequis) > 0 Then
        
        Call pIniciaMshFGrid(grdHRequisicion) 'deja especificado 2 renglones,2 columnas
        Call pIniciaMshFGrid(grdhBusqueda) 'deja especificado 2 renglones,2 columnas
        
        vgstrParametrosSP = Trim(Str(vgintCveDepto)) & "|" & Trim(Str(vgintCveEmp)) & "|" & IIf(Trim(vgstrEstatus) = "", "*", Trim(vgstrEstatus)) & "|" & IIf(Trim(vgstrTipoRequis) = "", "*", Trim(vgstrTipoRequis)) & "|" & fstrFechaSQL(Trim(CStr(vldtmFechaInicial)), , True) & "|" & fstrFechaSQL(Trim(CStr(vldtmFechaFinal)), , True) & "|" & Trim(Str(vlintFolioInicial)) & "|" & Trim(Str(vlintFolioFinal))
        
        Set rsIvSelFilRequisMaesxMixtas = frsEjecuta_SP(vgstrParametrosSP, "sp_IvSelFilRequisMaesxMixtas")
        
        If rsIvSelFilRequisMaesxMixtas.RecordCount > 0 Then
            Call pLimpiaMshFGrid(grdHRequisicion) 'limpia completamente el grid
            
            pLlenarMshFGrdRs grdhBusqueda, rsIvSelFilRequisMaesxMixtas
            
            Call pConfFGridBus(grdhBusqueda, "|Requisición|Tipo de requisición||Departamento||Empleado|Estatus|Urgente|Fecha|Hora")
            grdhBusqueda.Enabled = True
            grdHRequisicion.Enabled = True
            grdhBusqueda.SetFocus
        Else
            Call pLimpiaMshFGrid(grdhBusqueda) 'limpia completamente el grid
            Call pLimpiaMshFGrid(grdHRequisicion) 'limpia completamente el grid
            Call MsgBox(SIHOMsg("13"), (vbExclamation), "Mensaje") '!No existe información!
            grdhBusqueda.Enabled = False
            grdHRequisicion.Enabled = False
        End If
        rsIvSelFilRequisMaesxMixtas.Close
    End If

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pVerificaExistencia"))
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

    If grdhBusqueda.Rows > 0 Then
        grdhBusqueda.Refresh
        vgintColLoc = grdhBusqueda.Col
        vgstrAcumTextoBusqueda = ""
        grdhBusqueda.Col = vgintColLoc
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
    
    Dim vlintResul As Integer
    Dim vgintColOrdAnt As Integer
    Dim vlintNumero As Double
    Dim vlstrMensaje As String
    
    vgstrAcumTextoBusqueda = "" 'Inicializa el criterio de búsqueda dentro del gridHBusqueda
    
    ' Ordena solamente cuando un encabezado de columna es seleccionado con un click
    If grdhBusqueda.MouseRow >= grdhBusqueda.FixedRows Then
        vlintNumero = fintLocalizaPkRs(rsIvRequisicionMaestro, 0, grdhBusqueda.TextMatrix(grdhBusqueda.Row, 1))
        pMuestraRegistro
    Else
        vgintColOrdAnt = vgintColOrd 'Guarda la columna de ordenación anterior
        vgintColOrd = grdhBusqueda.Col  'Configura la columna a ordenar

        'Escoge el Tipo de Ordenamiento
        If vgintTipoOrd = 1 Then
            vgintTipoOrd = 2
            Else
                vgintTipoOrd = 1
            End If
        Call pOrdColMshFGrid(grdhBusqueda, vgintTipoOrd)
        Call pDesSelMshFGrid(grdhBusqueda)
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


Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
' Validación del <Escape> para regresar al Tab 0 (Mantenimiento) del sstObj, teniendo el
' enfoque en GrdHBusqueda
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    Dim vlintNumero As Integer
    Dim vlintResul As Integer

    Select Case KeyCode
        Case vbKeyReturn
            grdHBusqueda_DblClick
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
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
    On Error GoTo NotificaError

    grdhBusqueda.FocusRect = flexFocusNone
    Call pSelCriterioMshFGrid(grdhBusqueda, vgintColLoc, vlintKeyAscii)
    grdhBusqueda.FocusRect = flexFocusHeavy

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

Private Sub grdHRequisicion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError 'Manejo del error
    
    Dim vlintResul As Integer
    
    Select Case KeyCode
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
NotificaError:
    If vgblnExistioError Then
        Unload Me
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHRequisicion_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub optFecha_Click()
    pLimpiaPantalla
    grdhBusqueda.Enabled = False
    grdHRequisicion.Enabled = False
    txtFechaInicio.Enabled = True
    txtFechaFin.Enabled = True
    txtFolioIni.Enabled = False
    txtFolioFin.Enabled = False
End Sub

Private Sub optFecha_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError 'Manejo del error

    Dim vlintResul As Integer

    Select Case KeyCode
        Case vbKeyReturn
            If cboDepartamento.Enabled = True Then
                cboDepartamento.SetFocus
            Else
                cboEmpleado.SetFocus
            End If
                
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optFecha_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub optFolio_Click()
    pLimpiaPantalla
    grdhBusqueda.Enabled = False
    grdHRequisicion.Enabled = False
    txtFechaInicio.Enabled = False
    txtFechaFin.Enabled = False
    txtFolioIni.Enabled = True
    txtFolioFin.Enabled = True
End Sub

Private Sub optFolio_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo NotificaError 'Manejo del error

    Dim vlintResul As Integer

    Select Case KeyCode
        Case vbKeyReturn
            If cboDepartamento.Enabled = True Then
                cboDepartamento.SetFocus
            Else
                cboEmpleado.SetFocus
            End If
                
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":optFolio_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFechaFin_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
On Error GoTo NotificaError 'Manejo del error

    txtFechaFin = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
    Call pSelMkTexto(txtFechaFin)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaFin_GotFocus"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------------
' Procedimiento para validar cuando se presiona la tecla <Esc> para salir
'--------------------------------------------------------------------------
On Error GoTo NotificaError 'Manejo del error

    Dim vlintResul As Integer

    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaFechaFin Then
                pRevisaParametros
                pVerificaExistencia
            End If
            
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If

    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaFin_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub


Private Sub txtFechaInicio_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
On Error GoTo NotificaError 'Manejo del error

    txtFechaInicio = Format(CDate(fdtmServerFecha), "dd/mm/yyyy")
    Call pSelMkTexto(txtFechaInicio)
    
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaInicio_GotFocus"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If

End Sub

Private Sub txtFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------------
' Procedimiento para validar cuando se presiona la tecla <Esc> para salir
'--------------------------------------------------------------------------
On Error GoTo NotificaError 'Manejo del error

    Dim vlintResul As Integer

    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaFechaInicio Then
                Call pEnfocaMkTexto(txtFechaFin)
            End If
            
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select

NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaInicio_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFolioFin_GotFocus()
    Call pSelMkTexto(txtFolioFin)
End Sub

Private Sub txtFolioFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    

    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaFolioFinal Then
                pRevisaParametros
                pVerificaExistencia
            End If
            
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
            
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioFin_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub

Private Sub txtFolioIni_GotFocus()
    Call pSelMkTexto(txtFolioIni)
End Sub

Private Sub txtFolioIni_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    Dim vlintResul As Integer
    

    Select Case KeyCode
        Case vbKeyReturn
            If fblnValidaFolioInicio Then
                Call pEnfocaMkTexto(txtFolioFin)
            End If
            
        Case vbKeyEscape
            'Pregunta si desea salir del mantenimiento
            vlintResul = MsgBox(SIHOMsg("18") & " de " & frmConsultaRequisMix.Caption & "?", (vbYesNo + vbQuestion), "Mensaje")
            If vlintResul = vbYes Then
                Unload Me
            End If
    End Select
            
NotificaError:
    If vgblnExistioError Then
        Exit Sub
    Else
        If Err.Number <> 0 Then
            Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioIni_KeyDown"))
            On Error GoTo 0
            vgblnExistioError = True
            Exit Sub
        End If
    End If
End Sub
