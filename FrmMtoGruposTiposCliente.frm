VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMtoGruposTiposCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupos de tipos de clientes"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTobj 
      Height          =   7080
      Left            =   -600
      TabIndex        =   18
      Top             =   -435
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   12488
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmMtoGruposTiposCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmMtoGruposTiposCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GrdHBusqueda"
      Tab(1).Control(1)=   "optOrden(2)"
      Tab(1).Control(2)=   "optOrden(1)"
      Tab(1).Control(3)=   "optOrden(0)"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame1 
         Height          =   4450
         Left            =   705
         TabIndex        =   21
         Top             =   480
         Width           =   7170
         Begin VB.TextBox TxtGrupoTipoCliente 
            Height          =   315
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Clave del grupo"
            Top             =   195
            Width           =   735
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   315
            Left            =   1080
            MaxLength       =   3900
            TabIndex        =   1
            ToolTipText     =   "Descripción"
            Top             =   555
            Width           =   5970
         End
         Begin VB.CommandButton cmdAsignaTodo 
            Height          =   495
            Left            =   3330
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Ultimo registro"
            Top             =   1815
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAsignaUno 
            Height          =   495
            Left            =   3330
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":01AA
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Siguiente registro"
            Top             =   2340
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEliminaUno 
            Height          =   495
            Left            =   3330
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":031C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anterior registro"
            Top             =   2850
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEliminaTodo 
            Height          =   495
            Left            =   3330
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":048E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Primer registro"
            Top             =   3375
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.ListBox lstTiposEmpresasSel 
            Height          =   2985
            Left            =   4005
            TabIndex        =   7
            ToolTipText     =   "Tipos de paciente y empresas asignados"
            Top             =   1320
            Width           =   3030
         End
         Begin VB.ListBox lstTiposEmpresas 
            Height          =   2985
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Tipos de paciente y empresas disponibles"
            Top             =   1320
            Width           =   3030
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   225
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   585
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de paciente / Empresas asignados"
            Height          =   195
            Left            =   4005
            TabIndex        =   23
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipos de paciente / Empresas disponibles"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   2955
         End
      End
      Begin VB.Frame Frame4 
         Height          =   690
         Left            =   2480
         TabIndex        =   20
         Top             =   4960
         Width           =   3615
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   480
            Left            =   2550
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0600
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Guardar el registro"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2055
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0942
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0AB4
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1065
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0C26
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0D98
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":0F0A
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   480
            Left            =   3045
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoGruposTiposCliente.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Borrar el registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Clave"
         Height          =   270
         Index           =   0
         Left            =   -74295
         TabIndex        =   15
         Top             =   540
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Clave"
         Height          =   270
         Index           =   1
         Left            =   -75000
         TabIndex        =   19
         Top             =   0
         Width           =   1080
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Descripción"
         Height          =   270
         Index           =   2
         Left            =   -73320
         TabIndex        =   16
         Top             =   540
         Width           =   1230
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdHBusqueda 
         Height          =   4800
         Left            =   -74280
         TabIndex        =   17
         Top             =   840
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   8467
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmMtoGruposTiposCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Grupo de tipos de clientes
'| Nombre del Formulario    : FrmMtoGruposTiposCliente
'-------------------------------------------------------------------------------------
'| Objetivo: Se contará con un catálogo para crear agrupaciones de tipos de paciente y empresas, que se utilizarán para generar el reporte de ventas cruzadas.
'-------------------------------------------------------------------------------------
'| Autor                    : Jesús Valles Torres
'| Fecha de Creación        : 11/Agosto/2016
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables
Dim vgintCont As Integer
Dim blnEnfocando As Boolean
Dim vlblnActivaraCaptura As Boolean
Dim vlstrx As String
Dim rsPvGruposClientes As New ADODB.Recordset
Dim rsPvGruposClientesDetalle As New ADODB.Recordset
Dim lngSig As Long
Dim vgstrEstadoManto As String
Dim vgstrSentencia As String
Dim vlblnLectura As Boolean

Private Sub pCargarTiposEmpresas()
    Dim rsDatos As New ADODB.Recordset
    
    vgstrSentencia = "select * from " & _
                        "(select tnycvetipopaciente * -1 cve, trim(vchdescripcion) nombre from adtipopaciente where bitactivo = 1 " & _
                        "Union All " & _
                        "select intcveempresa cve, trim(vchdescripcion) nombre from ccempresa where bitactivo = 1) info " & _
                    "order by nombre"
    
    Set rsDatos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    If rsDatos.RecordCount > 0 Then
        With lstTiposEmpresas
            .Clear
            Do While Not rsDatos.EOF
                .AddItem rsDatos!Nombre, .ListCount
                .ItemData(.newIndex) = rsDatos!Cve
                rsDatos.MoveNext
            Loop
        End With
    End If
    rsDatos.Close

End Sub

Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    With GrdHBusqueda
        .FormatString = "|Clave|Descripción"
        .ColWidth(0) = 100 'Fix
        .ColWidth(1) = 700 'Clave
        .ColWidth(2) = 5500 'Descripcion
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub pNuevoRegistro()
    On Error GoTo NotificaError
    
    Dim rsSig As ADODB.Recordset
            
    vgstrEstadoManto = ""
    TxtDescripcion.Text = ""
    lstTiposEmpresasSel.Clear
    
    Set rsSig = frsRegresaRs("select max(INTCVEGRUPO) from PVGRUPOTIPOCLIENTE", adLockReadOnly, adOpenForwardOnly)
    
    If Not rsSig.EOF Then
        lngSig = IIf(IsNull(rsSig.Fields(0).Value), 0, rsSig.Fields(0).Value) + 1
    Else
        lngSig = 1
    End If
    
    TxtGrupoTipoCliente.Text = lngSig
    pHabilitaBotonModifica (True)
    cmdBuscar.Enabled = True
    cmdDelete.Enabled = False
    pHabilitaComponentesCaptura False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNuevoRegistro"))
End Sub

Private Sub pLlenaGrid(vlintOrden As Integer)
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsGruposClientes As New ADODB.Recordset
    Dim vlstrOrden As String
    
    GrdHBusqueda.Clear
    If vlintOrden = 1 Then
        vlstrOrden = " PVGRUPOTIPOCLIENTE.VCHDESCRIPCIONGRUPO"
    Else
        vlstrOrden = " PVGRUPOTIPOCLIENTE.INTCVEGRUPO"
    End If

    vlstrSentencia = "SELECT PVGRUPOTIPOCLIENTE.INTCVEGRUPO, " & _
        "rtrim(PVGRUPOTIPOCLIENTE.VCHDESCRIPCIONGRUPO) " & _
        "FROM PVGRUPOTIPOCLIENTE order by " & vlstrOrden
    
    Set rsGruposClientes = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    pLlenarMshFGrdRs GrdHBusqueda, rsGruposClientes, 0
    
    pConfiguraGrid
    rsGruposClientes.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaGrid"))
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError
    
    SSTobj.Tab = 1
    pConfiguraGrid
    pLlenaGrid 2
    GrdHBusqueda.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
End Sub

Private Sub cmdDelete_Click()
    Dim vllngPersonaGraba As Long

    On Error GoTo NotificaError
    
    If vlblnLectura Then
        'El usuario no tiene permiso para realizar esta operación.
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "GRUPO DE TIPOS DE CLIENTES", TxtGrupoTipoCliente.Text)
        pEjecutaSentencia "Delete from  PVGRUPOTIPOCLIENTEDETALLE where INTCVEGRUPO = " & Trim(TxtGrupoTipoCliente.Text)
        pEjecutaSentencia "Delete from  PVGRUPOTIPOCLIENTE where INTCVEGRUPO = " & Trim(TxtGrupoTipoCliente.Text)
                
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "GRUPOS DE TIPOS DE CLIENTES", TxtGrupoTipoCliente.Text)
                
        EntornoSIHO.ConeccionSIHO.CommitTrans
            
        rsPvGruposClientes.Requery
        rsPvGruposClientesDetalle.Requery
            
        TxtGrupoTipoCliente.SetFocus
        
        vlblnActivaraCaptura = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdGrabarRegistro_Click()
    Dim vllngPersonaGraba As Long
    Dim lngCveGrupoCliente As Long
    Dim vlrsAux As New ADODB.Recordset
    Dim vllngClaveOtroConcepto As Long
    Dim vlintRow As Integer
    Dim lvsqlstr As String
    Dim rsCargoAsignadoACuarto As New ADODB.Recordset 'Valida que el otro concepto no este asignado como cargo a uno o más cuartos
    Dim rsConcEnGrupoCargos As New ADODB.Recordset  'Valida que el otro concepto no esté asignado a uno o más grupos de cargos
    
    On Error GoTo NotificaError
    
    If vlblnLectura Then
        '¡El usuario no tiene permiso para grabar datos!
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        Exit Sub
    End If
    
    If RTrim(TxtDescripcion.Text) = "" Then
        MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
        TxtDescripcion.SetFocus
    Else
    
        If lstTiposEmpresasSel.ListCount = 0 Then
            '¡Dato no válido, seleccione un valor de la lista!
            MsgBox SIHOMsg(3), vbOKOnly + vbCritical, "Mensaje"
            lstTiposEmpresas.SetFocus
            Exit Sub
        End If
            
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            EntornoSIHO.ConeccionSIHO.BeginTrans
            If vgstrEstadoManto = "A" Then
                'Inserta el maestro
                With rsPvGruposClientes
                    .AddNew
                    !VCHDESCRIPCIONGRUPO = Trim(TxtDescripcion.Text)
                    .Update
                    lngCveGrupoCliente = flngObtieneIdentity("SEC_PVGRUPOTIPOCLIENTE", !intCveGrupo)
                    TxtGrupoTipoCliente.Text = lngCveGrupoCliente
                End With
                
                For vlintRow = 0 To lstTiposEmpresasSel.ListCount - 1
                    'Inserta el detalle
                    vgstrSentencia = "Insert into PVGRUPOTIPOCLIENTEDETALLE Values(" & lngCveGrupoCliente & "," & IIf(lstTiposEmpresasSel.ItemData(vlintRow) > 0, lstTiposEmpresasSel.ItemData(vlintRow), lstTiposEmpresasSel.ItemData(vlintRow) * -1) & "," & IIf(lstTiposEmpresasSel.ItemData(vlintRow) > 0, 1, 0) & ")"
                    pEjecutaSentencia vgstrSentencia
                Next vlintRow
            Else
                'Actualiza el maestro
                pEjecutaSentencia "update PVGRUPOTIPOCLIENTE set VCHDESCRIPCIONGRUPO = '" & Trim(TxtDescripcion.Text) & "' WHERE INTCVEGRUPO = " & TxtGrupoTipoCliente.Text
                
                pEjecutaSentencia "Delete from  PVGRUPOTIPOCLIENTEDETALLE Where INTCVEGRUPO = " & TxtGrupoTipoCliente.Text
                
                For vlintRow = 0 To lstTiposEmpresasSel.ListCount - 1
                    'Inserta el detalle
                    vgstrSentencia = "Insert into PVGRUPOTIPOCLIENTEDETALLE Values(" & TxtGrupoTipoCliente.Text & "," & IIf(lstTiposEmpresasSel.ItemData(vlintRow) > 0, lstTiposEmpresasSel.ItemData(vlintRow), lstTiposEmpresasSel.ItemData(vlintRow) * -1) & "," & IIf(lstTiposEmpresasSel.ItemData(vlintRow) > 0, 1, 0) & ")"
                    pEjecutaSentencia vgstrSentencia
                Next vlintRow
            End If
    
            If vgstrEstadoManto = "A" Then
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "GRUPOS DE TIPOS DE CLIENTES", Str(lngCveGrupoCliente))
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "GRUPOS DE TIPOS DE CLIENTES", TxtGrupoTipoCliente.Text)
            End If
    
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            rsPvGruposClientes.Requery
            rsPvGruposClientesDetalle.Requery
            
            TxtGrupoTipoCliente.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
End Sub

Private Sub pAsigna(Asigna As Boolean, Optional Todos As Boolean)
'Procedimiento que asigna o elimina areas

    If Asigna Then
        If lstTiposEmpresas.ListCount > 0 Then
            If Todos Then
                lstTiposEmpresasSel.Clear
                For vgintCont = 0 To lstTiposEmpresas.ListCount - 1
                    lstTiposEmpresasSel.AddItem lstTiposEmpresas.List(vgintCont), lstTiposEmpresasSel.ListCount
                    lstTiposEmpresasSel.ItemData(lstTiposEmpresasSel.newIndex) = lstTiposEmpresas.ItemData(vgintCont)
                Next
            Else
                If lstTiposEmpresas.ListIndex = -1 Then Exit Sub
                If fValida(lstTiposEmpresas.ItemData(lstTiposEmpresas.ListIndex)) = False Then Exit Sub
                lstTiposEmpresasSel.AddItem lstTiposEmpresas.List(lstTiposEmpresas.ListIndex), lstTiposEmpresasSel.ListCount
                lstTiposEmpresasSel.ItemData(lstTiposEmpresasSel.newIndex) = lstTiposEmpresas.ItemData(lstTiposEmpresas.ListIndex)
                lstTiposEmpresas.SetFocus
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
        End If
    Else
        If lstTiposEmpresasSel.ListCount > 0 Then
            If Todos Then
                lstTiposEmpresasSel.Clear
            Else
                If lstTiposEmpresasSel.ListIndex = -1 Then Exit Sub
                lstTiposEmpresasSel.RemoveItem (lstTiposEmpresasSel.ListIndex)
                If lstTiposEmpresasSel.ListCount > 0 Then lstTiposEmpresasSel.SetFocus
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
        End If
    End If
End Sub

Private Function fValida(Cve As Long) As Boolean
'Valida que el elemento no este asignado anteriormente
    
    fValida = True
    
    With lstTiposEmpresasSel
        If .ListCount > 0 Then
            For vgintCont = 0 To .ListCount - 1
                If Cve = .ItemData(vgintCont) Then
                    fValida = False
                    Call Beep
                    Exit For
                End If
            Next
        End If
    End With
End Function

Private Sub cmdAsignaTodo_Click()
    vlblnActivaraCaptura = True
    pAsigna True, True
End Sub

Private Sub cmdAsignaUno_Click()
    vlblnActivaraCaptura = True
    pAsigna True
End Sub

Private Sub cmdEliminaTodo_Click()
    vlblnActivaraCaptura = True
    pAsigna False, True
End Sub

Private Sub cmdEliminaUno_Click()
    vlblnActivaraCaptura = True
    pAsigna False
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    Dim vlstrx As String
    Dim rscmdConceptoFacturacion As New ADODB.Recordset
                   
    If cgstrModulo = "PV" Then
        vlblnLectura = Not (fblnRevisaPermiso(vglngNumeroLogin, 4038, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4038, "C", True))
    Else
        vlblnLectura = Not (fblnRevisaPermiso(vglngNumeroLogin, 4040, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 4040, "C", True))
    End If
            
    Set rsPvGruposClientes = frsRegresaRs("select * from PVGRUPOTIPOCLIENTE", adLockOptimistic, adOpenDynamic)
    Set rsPvGruposClientesDetalle = frsRegresaRs("select * from PVGRUPOTIPOCLIENTEDETALLE", adLockOptimistic, adOpenDynamic)
   
    SSTobj.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    If SSTobj.Tab = 1 Then
        Cancel = True
        optOrden(0).SetFocus
        
        SSTobj.Tab = 0
        TxtGrupoTipoCliente.SetFocus
    Else
        If vgstrEstadoManto <> "" Then
            Cancel = True
            If MsgBox(SIHOMsg(9), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                Me.TxtDescripcion.SetFocus
                TxtGrupoTipoCliente.SetFocus
                lstTiposEmpresasSel.Clear
                optOrden(0).SetFocus
                GrdHBusqueda.Clear
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo NotificaError
    
    rsPvGruposClientes.Close
    rsPvGruposClientesDetalle.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
End Sub

Private Sub grdHBusqueda_DblClick()
    On Error GoTo NotificaError
    
    If GrdHBusqueda.RowData(1) <> 0 Then
        TxtGrupoTipoCliente.Text = GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)
        TxtGrupoTipoCliente_KeyDown vbKeyReturn, 0
        SSTobj.Tab = 0
        optOrden(0).SetFocus
        GrdHBusqueda.Clear
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
End Sub

Private Sub grdHBusqueda_KeyPress(KeyAscii As Integer)
    If GrdHBusqueda.RowData(1) <> 0 Then
        TxtGrupoTipoCliente.Text = GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)
        TxtGrupoTipoCliente_KeyDown vbKeyReturn, 0
        SSTobj.Tab = 0
        optOrden(0).SetFocus
        
        GrdHBusqueda.Clear
    End If
End Sub

Private Sub txtCvePlantilla_GotFocus()
    On Error GoTo NotificaError
    
    pNuevoRegistro
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCvePlantilla_GotFocus"))
End Sub

Private Sub pModificaRegistro()
    On Error GoTo NotificaError
    Dim rsDepartamentos As New ADODB.Recordset
    
    vgstrEstadoManto = "M"
    
    With rsPvGruposClientes
        TxtGrupoTipoCliente.Text = !intCveGrupo
        TxtDescripcion.Text = Trim(!VCHDESCRIPCIONGRUPO)
    End With
    
    lstTiposEmpresasSel.Clear
    
    'Se carga
    vgstrSentencia = "SELECT CASE WHEN NOT ADTIPOPACIENTE.TNYCVETIPOPACIENTE IS NULL THEN ADTIPOPACIENTE.TNYCVETIPOPACIENTE * -1 ELSE CCEMPRESA.INTCVEEMPRESA END CVE, " & _
                            "CASE WHEN NOT ADTIPOPACIENTE.TNYCVETIPOPACIENTE IS NULL THEN TRIM(ADTIPOPACIENTE.VCHDESCRIPCION) ELSE TRIM(CCEMPRESA.VCHDESCRIPCION) END NOMBRE " & _
                        "From PVGRUPOTIPOCLIENTEDETALLE " & _
                            "LEFT JOIN ADTIPOPACIENTE ON PVGRUPOTIPOCLIENTEDETALLE.BITEMPRESA = 0 " & _
                                "AND ADTIPOPACIENTE.TNYCVETIPOPACIENTE = PVGRUPOTIPOCLIENTEDETALLE.TNYCVETIPOPACEMPRESA " & _
                            "LEFT JOIN CCEMPRESA ON PVGRUPOTIPOCLIENTEDETALLE.BITEMPRESA = 1 " & _
                                "AND CCEMPRESA.INTCVEEMPRESA = PVGRUPOTIPOCLIENTEDETALLE.TNYCVETIPOPACEMPRESA " & _
                        "WHERE PVGRUPOTIPOCLIENTEDETALLE.INTCVEGRUPO = " & TxtGrupoTipoCliente.Text & " " & _
                        "ORDER BY NOMBRE"
    Set rsDepartamentos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rsDepartamentos.RecordCount > 0 Then
        With lstTiposEmpresasSel
            Do While Not rsDepartamentos.EOF
                .AddItem rsDepartamentos!Nombre, .ListCount
                .ItemData(.newIndex) = rsDepartamentos!Cve
                rsDepartamentos.MoveNext
            Loop
        End With
    End If
    
    rsDepartamentos.Close
    
    pHabilitaBotonModifica (True)
    pHabilitaComponentesCaptura True
    
    vlblnActivaraCaptura = True
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRegistro"))
End Sub

Private Sub optOrden_Click(Index As Integer)
    pLlenaGrid IIf(optOrden(0).Value, 0, 1)
    GrdHBusqueda.SetFocus
End Sub

Private Sub TxtGrupoTipoCliente_GotFocus()
    On Error GoTo NotificaError
    If Not blnEnfocando Then
        pCargarTiposEmpresas
    
        pNuevoRegistro
        If TxtGrupoTipoCliente.Enabled And TxtGrupoTipoCliente.Visible Then
            pEnfocaTextBox TxtGrupoTipoCliente
        End If
    End If
    blnEnfocando = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtGrupoTipoCliente_GotFocus"))
End Sub

Private Sub TxtGrupoTipoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsBusca As ADODB.Recordset
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If Trim(TxtGrupoTipoCliente.Text) = "" Then
            '¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
            pNuevoRegistro
            TxtGrupoTipoCliente.SetFocus
            Exit Sub
        Else
            If CLng(Trim(TxtGrupoTipoCliente.Text)) = 0 Then
                '¡No ha ingresado datos!
                MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
                TxtGrupoTipoCliente.Text = ""
                pNuevoRegistro
                TxtGrupoTipoCliente.SetFocus
                Exit Sub
            End If
        End If
        
        'Buscar criterio
        TxtDescripcion.Enabled = True
        If lngSig = CLng(TxtGrupoTipoCliente.Text) Then
            vgstrEstadoManto = "A" 'Alta
            Call pEnfocaTextBox(TxtDescripcion)
            pHabilitaComponentesCaptura True
            pHabilitaBotonModifica (False)
            cmdGrabarRegistro.Enabled = True
            cmdBuscar.Enabled = False
        Else
            If fintLocalizaPkRs(rsPvGruposClientes, 0, TxtGrupoTipoCliente.Text) > 0 Then
                pModificaRegistro
                vlblnActivaraCaptura = False
                vgstrEstadoManto = "M" 'Modificacion
                pHabilitaComponentesCaptura True
                pHabilitaBotonModifica (True)
                cmdBuscar.Enabled = True
                Call pEnfocaTextBox(TxtDescripcion)
            Else
'                Set rsBusca = frsRegresaRs("select * from PVGRUPOTIPOCLIENTE where INTCVEGRUPO = " & Me.TxtGrupoTipoCliente.Text, adLockReadOnly, adOpenForwardOnly)
'                If rsBusca.EOF Then
'                    vgstrEstadoManto = "A" 'Alta
'                    Call pEnfocaTextBox(TxtDescripcion)
'                    pHabilitaComponentesCaptura True
'                    pHabilitaBotonModifica (False)
'                    cmdGrabarRegistro.Enabled = True
'                    cmdBuscar.Enabled = False
'                Else
                    pNuevoRegistro
                    MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                    Call pEnfocaTextBox(TxtGrupoTipoCliente)
'                End If
'                rsBusca.Close
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtGrupoTipoCliente_KeyDown"))
End Sub

Private Sub TxtGrupoTipoCliente_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then KeyAscii = 7

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtGrupoTipoCliente_KeyPress"))
End Sub

Private Sub txtDescripcion_GotFocus()
    If vlblnActivaraCaptura Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
    End If
    
    Call pEnfocaTextBox(TxtDescripcion)
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If cmdBuscar.Enabled Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
    cmdGrabarRegistro.Enabled = Not vlblnHabilita
    cmdDelete.Enabled = vlblnHabilita
    
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

Private Sub cmdAnteriorRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGruposClientes, "A")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGruposClientes, "I")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
End Sub

Private Sub cmdSiguienteRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGruposClientes, "S")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
End Sub

Private Sub cmdUltimoRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGruposClientes, "U")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then SendKeys vbTab

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Txtdescripcion_KeyDown"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub lstTiposEmpresas_DblClick()
    pAsigna True
End Sub

Private Sub lstTiposEmpresas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna True
End Sub

Private Sub lstTiposEmpresasSel_DblClick()
    pAsigna False
End Sub

Private Sub lstTiposEmpresasSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna False
End Sub

Private Sub pHabilitaComponentesCaptura(vlblnHabilita As Boolean)
    TxtDescripcion.Enabled = vlblnHabilita
    lstTiposEmpresas.Enabled = vlblnHabilita
    cmdAsignaTodo.Enabled = vlblnHabilita
    cmdAsignaUno.Enabled = vlblnHabilita
    cmdEliminaUno.Enabled = vlblnHabilita
    cmdEliminaTodo.Enabled = vlblnHabilita
    lstTiposEmpresasSel.Enabled = vlblnHabilita
End Sub
