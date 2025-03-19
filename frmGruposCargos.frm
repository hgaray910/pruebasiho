VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGruposCargos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grupos de cargos para paquetes"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTobj 
      Height          =   6960
      Left            =   -10
      TabIndex        =   0
      Top             =   -360
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   12277
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmGruposCargos.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GrdHBusqueda"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmGruposCargos.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "TxtCveGrupo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TxtDescripcion"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "OptOtros"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "OptImagenologia"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "OptLaboratorio"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "OptMedicamentos"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "OptArticulos"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "FrmMovimientos"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "FrmBusqueda"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "GrdCargos"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkActivo"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.CheckBox chkActivo 
         Caption         =   "Activo"
         Height          =   315
         Left            =   6242
         TabIndex        =   22
         ToolTipText     =   "Estado de activo o inactivo para el grupo"
         Top             =   6240
         Value           =   1  'Checked
         Width           =   743
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdCargos 
         Height          =   1980
         Left            =   150
         TabIndex        =   14
         ToolTipText     =   "Elementos incluidos en el grupo de cargos"
         Top             =   4140
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   3493
         _Version        =   393216
         FocusRect       =   0
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame FrmBusqueda 
         Height          =   1695
         Left            =   150
         TabIndex        =   28
         Top             =   1600
         Width           =   6840
         Begin VB.OptionButton optDescripcion 
            Caption         =   "&Descripción"
            Height          =   315
            Left            =   5520
            TabIndex        =   10
            ToolTipText     =   "Buscar por descripción"
            Top             =   205
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optClave 
            Caption         =   "&Clave"
            Height          =   315
            Left            =   4680
            TabIndex        =   9
            ToolTipText     =   "Buscar por clave"
            Top             =   205
            Width           =   735
         End
         Begin VB.TextBox txtSeleArticulo 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Teclee la clave o la descripción del elemento"
            Top             =   205
            Width           =   4455
         End
         Begin VB.ListBox lstCargos 
            DragIcon        =   "frmGruposCargos.frx":0038
            Height          =   1035
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Elementos resultantes de la búsqueda"
            Top             =   540
            Width           =   6615
         End
      End
      Begin VB.Frame FrmMovimientos 
         Height          =   780
         Left            =   2905
         TabIndex        =   27
         Top             =   3290
         Width           =   1320
         Begin VB.CommandButton cmdSelecciona 
            Caption         =   "Excluir"
            Height          =   540
            Index           =   1
            Left            =   660
            MaskColor       =   &H80000014&
            Picture         =   "frmGruposCargos.frx":0482
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Excluir elemento"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   570
         End
         Begin VB.CommandButton cmdSelecciona 
            Caption         =   "Incluir"
            Height          =   540
            Index           =   0
            Left            =   75
            MaskColor       =   &H80000014&
            Picture         =   "frmGruposCargos.frx":05DC
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Incluir elemento"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   570
         End
      End
      Begin VB.OptionButton OptArticulos 
         Caption         =   "Artículos"
         Height          =   315
         Left            =   150
         TabIndex        =   3
         ToolTipText     =   "Sólo artículos"
         Top             =   1320
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptMedicamentos 
         Caption         =   "Medicamentos"
         Height          =   315
         Left            =   1180
         TabIndex        =   4
         ToolTipText     =   "Sólo medicamentos"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton OptLaboratorio 
         Caption         =   "Laboratorio"
         Height          =   315
         Left            =   4330
         TabIndex        =   6
         ToolTipText     =   "Sólo laboratorio"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton OptImagenologia 
         Caption         =   "Servicios auxiliares"
         Height          =   315
         Left            =   2605
         TabIndex        =   5
         ToolTipText     =   "Sólo servicios auxiliares"
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton OptOtros 
         Caption         =   "Otros conceptos"
         Height          =   315
         Left            =   5530
         TabIndex        =   7
         ToolTipText     =   "Sólo otros conceptos"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1210
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "Descripción del grupo"
         Top             =   940
         Width           =   5805
      End
      Begin VB.TextBox TxtCveGrupo 
         Height          =   315
         Left            =   1210
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Clave del grupo"
         Top             =   540
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Height          =   690
         Left            =   1758
         TabIndex        =   24
         Top             =   6130
         Width           =   3615
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   480
            Left            =   2520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":0736
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Guardar el grupo"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2055
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":0A78
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":0BEA
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1065
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":0D5C
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":0ECE
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmGruposCargos.frx":1040
            Style           =   1  'Graphical
            TabIndex        =   15
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
            Picture         =   "frmGruposCargos.frx":11B2
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Borrar el grupo"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdHBusqueda 
         Height          =   6060
         Left            =   -74865
         TabIndex        =   23
         Top             =   600
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   10689
         _Version        =   393216
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   1000
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clave"
         Height          =   195
         Left            =   150
         TabIndex        =   25
         Top             =   600
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmGruposCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja
'| Nombre del Formulario    : frmGruposCargos
'-------------------------------------------------------------------------------------
'| Objetivo: Catalogo de grupos de cargos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Jesus Valles Torres
'| Autor                    : Jesus Valles Torres
'| Fecha de Creación        : 20/Enero/2012
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit

Dim vgstrEstadoManto As String
Dim rsPvGrupoCargo As New ADODB.Recordset
Dim rsPvDetalleGrupoCargo As New ADODB.Recordset
Dim lngSig As Long

Private Sub pConfiguraGridCargos()
' Grid que contiene los cargos incluidos en el grupo
    On Error GoTo NotificaError
    
    With grdCargos
        .FormatString = "|Clave|Descripción|Tipo"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 0    'Clave del cargo
        .ColWidth(2) = 6007 'Descripcion del cargo
        .ColWidth(3) = 450  'Tipo del cargo (guarda AR = Artículos, ME = Medicamentos, OC = otro conceptos, ES = Estudios, EX = Exámenes, GE = Grupos de exámenes)
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
End Sub

Private Sub pConfiguraGrid()
' Grid para busqueda de grupos de cargos
    On Error GoTo NotificaError
    
    With grdHBusqueda
        .Cols = 4
        .FormatString = "|Clave|Descripción|Tipo"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 550  'Clave del grupo
        .ColWidth(2) = 5452 'Descripcion del grupo
        .ColWidth(3) = 450  'Tipo de cargo que se maneja dentro del grupo (guarda AR = Artículos, ME = Medicamentos, OC = otro conceptos, ES = Estudios, EX = Exámenes, GE = Grupos de exámenes)
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub pNuevoRegistro()
' Prepara el catalogo para un nuevo registro
    On Error GoTo NotificaError
    Dim rsSig As ADODB.Recordset
    
    vgstrEstadoManto = ""
    txtDescripcion.Text = ""
    chkActivo.Value = 1
    
    lngSig = fSigConsecutivo("intCveGrupo", "PvGrupoCargo")
    
    TxtCveGrupo.Text = lngSig
    pHabilitaComponentesCaptura False
    grdCargos.Clear
    grdCargos.Rows = 2
    grdCargos.Cols = 4
    pConfiguraGridCargos
    
    pLimpiaBusqueda
    pHabilitaElementos 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNuevoRegistro"))
End Sub

Private Sub pLlenaGrid(vlintOrden As Integer)
' Llena el grid de la busqueda
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsGrupos As New ADODB.Recordset
    Dim vlintcontador As Integer

    grdHBusqueda.Clear
    vlstrSentencia = "SELECT * FROM PvGrupoCargo ORDER BY " & IIf(vlintOrden = 0, "intCveGrupo", IIf(vlintOrden = 1, "vchNombre,intCveGrupo", "chrtipo,vchNombre,intCveGrupo"))

    Set rsGrupos = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    Call pLlenarMshFGrdRs(grdHBusqueda, rsGrupos)
    pConfiguraGrid
    rsGrupos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaGrid"))
End Sub

Private Sub chkActivo_Click()
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError
        SSTObj.Tab = 0
        pLlenaGrid 1
        pConfiguraGrid
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    Dim rsGruposEnUso As New ADODB.Recordset
    Dim vlstrCadenaMsj As String
    Dim vlintcontador As Integer
    Dim vllngPersonaGraba As Long
        
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2440, 2441), "E") Then
        Set rsGruposEnUso = frsRegresaRs("SELECT MPQ.intnumpaquete CvePaquete, TRIM(chrdescripcion) DescPaquete " & _
                                         "FROM PVPAQUETE MPQ " & _
                                            "INNER JOIN PVDETALLEPAQUETE DPQ ON DPQ.intnumpaquete = MPQ.intnumpaquete " & _
                                         "WHERE DPQ.intCveCargo = " & TxtCveGrupo.Text & " " & _
                                            "AND chrTipoCargo = 'GC' ORDER BY CvePaquete", adLockReadOnly)
    
        With rsGruposEnUso
            If .RecordCount > 0 Then
                .MoveFirst
                For vlintcontador = 1 To .RecordCount
                    vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & Format(!cvePaquete, "########") & " " & !DescPaquete
                    .MoveNext
                Next vlintcontador
                MsgBox SIHOMsg(1103) & vlstrCadenaMsj, vbOKOnly + vbInformation, "Mensaje"
                Exit Sub
            Else
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba <> 0 Then
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                        pEjecutaSentencia "DELETE FROM PVDETALLEGRUPOCARGO WHERE IntCveGrupo = " & TxtCveGrupo.Text
                        pEjecutaSentencia "DELETE FROM PVGRUPOCARGO WHERE IntCveGrupo = " & TxtCveGrupo.Text
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "BORRADO DE GRUPOS DE CARGOS", TxtCveGrupo.Text)
                    
                    rsPvGrupoCargo.Requery
                    rsPvDetalleGrupoCargo.Requery
                    
                    TxtCveGrupo.SetFocus
                End If
            End If
            rsGruposEnUso.Close
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Sub cmdGrabarRegistro_Click()
    Dim vlrsAux As New ADODB.Recordset
    Dim vllngClaveOtroConcepto As Long
    Dim vlintRow As Integer
    Dim vlintIdentity As Long
    Dim rsGruposEnUso As New ADODB.Recordset
    Dim rsGrupoExamenInactivo As New ADODB.Recordset
    Dim rsExamenInactivo As New ADODB.Recordset
    Dim vlstrCadenaMsj As String
    Dim vlintcontador As Integer
    Dim vlblnBorrarCargos As Boolean
    Dim vllngPersonaGraba As Long
    Dim rsAux As New ADODB.Recordset
    
    On Error GoTo NotificaError
    
    vlblnBorrarCargos = False
    
    If fblnRevisaPermiso(vglngNumeroLogin, IIf(cgstrModulo = "PV", 2440, 2441), "E") Then
        
        If RTrim(txtDescripcion.Text) = "" Then
            ' ¡No ha ingresado datos!
            MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
            txtDescripcion.SetFocus
        Else
            If grdCargos.Rows > 2 Or grdCargos.TextMatrix(1, 1) <> "" Then
                ' ¿Desea guardar los datos?
                ' ¿Desea guardar los cambios?
                
                If chkActivo.Value = 0 Then
                    Set rsGruposEnUso = frsRegresaRs("SELECT MPQ.intnumpaquete CvePaquete, TRIM(chrdescripcion) DescPaquete " & _
                                                     "FROM PVPAQUETE MPQ " & _
                                                        "INNER JOIN PVDETALLEPAQUETE DPQ ON DPQ.intnumpaquete = MPQ.intnumpaquete " & _
                                                     "WHERE DPQ.intCveCargo = " & TxtCveGrupo.Text & " " & _
                                                        "AND chrTipoCargo = 'GC' ORDER BY CvePaquete", adLockReadOnly)
                    With rsGruposEnUso
                        If .RecordCount > 0 Then
                            .MoveFirst
                            For vlintcontador = 1 To .RecordCount
                                vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & Format(!cvePaquete, "########") & " " & !DescPaquete
                                .MoveNext
                            Next vlintcontador
                            vlblnBorrarCargos = True
                        End If
                    End With
                End If
    
                'If MsgBox(IIf(vlblnBorrarCargos, SIHOMsg(1104) & vlstrCadenaMsj, SIHOMsg(IIf(vgstrEstadoManto = "A", 4, 440))), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                  If chkActivo.Value = 1 Then
                    'Revisa que los articulos o lo que sea no esten inactivos..
                    With grdCargos
                        .Row = 1
                        For vlintRow = .Row To .Rows - 1
                            If Trim(.TextMatrix(vlintRow, 1)) <> "" Then
                                  Select Case Trim(.TextMatrix(vlintRow, 3))
                                    Case "OC" ' Otros conceptos inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT PVOTROCONCEPTO.INTCVECONCEPTO FROM PVOTROCONCEPTO WHERE PVOTROCONCEPTO.BITESTATUS = 0 AND PVOTROCONCEPTO.INTCVECONCEPTO = " & (.TextMatrix(vlintRow, 1)), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El otro concepto " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        End If
                                    Case "AR"  ' Artículos inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT IVARTICULO.INTIDARTICULO FROM IVARTICULO WHERE IVARTICULO.VCHESTATUS = 'INACTIVO' AND IVARTICULO.INTIDARTICULO = " & (.TextMatrix(vlintRow, 1)), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El artículo " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        End If
                                    Case "ME"  ' Medicamentos inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT IVARTICULO.INTIDARTICULO FROM IVARTICULO WHERE IVARTICULO.VCHESTATUS = 'INACTIVO' AND IVARTICULO.INTIDARTICULO = " & (.TextMatrix(vlintRow, 1)), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El medicamento " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        End If
                                    Case "EX" ' Examenes inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT LAEXAMEN.INTCVEEXAMEN FROM LAEXAMEN WHERE LAEXAMEN.BITESTATUSACTIVO = 0 AND LAEXAMEN.INTCVEEXAMEN = " & (.TextMatrix(vlintRow, 1)), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El exámen " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        End If
                                    Case "ES" ' Estudios inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT IMESTUDIO.INTCVEESTUDIO FROM IMESTUDIO WHERE IMESTUDIO.BITSTATUSACTIVO = 0 AND IMESTUDIO.INTCVEESTUDIO = " & (.TextMatrix(vlintRow, 1)), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El estudio " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        End If
                                    Case "GE" ' Grupos de exámenes inactivos
                                        Set rsAux = frsRegresaRs("SELECT DISTINCT LAGRUPOEXAMEN.INTCVEGRUPO FROM LAGRUPOEXAMEN WHERE LAGRUPOEXAMEN.BITESTATUSACTIVO = 0 AND LAGRUPOEXAMEN.INTCVEGRUPO = " & ((.TextMatrix(vlintRow, 1)) / -1), adLockOptimistic, adOpenDynamic)
                                        If rsAux.RecordCount <> 0 Then
                                             Call MsgBox("El grupo de exámenes " & Trim(.TextMatrix(vlintRow, 2)) & " no se encuentra activo.", vbExclamation, "Mensaje")
                                             Exit Sub
                                        Else
                                            Set rsGrupoExamenInactivo = frsRegresaRs("SELECT INTCVEEXAMEN FROM LADETALLEGRUPO WHERE LADETALLEGRUPO.CHRTIPOREGISTRO = 'E' AND LADETALLEGRUPO.INTCVEGRUPO = " & ((.TextMatrix(vlintRow, 1)) / -1), adLockOptimistic, adOpenDynamic)
                                            If rsGrupoExamenInactivo.RecordCount <> 0 Then
                                                rsGrupoExamenInactivo.MoveFirst
                                                Do While Not rsGrupoExamenInactivo.EOF
                                                    Set rsExamenInactivo = frsRegresaRs("SELECT DISTINCT LAEXAMEN.INTCVEEXAMEN, LAEXAMEN.CHRNOMBRE FROM LAEXAMEN WHERE LAEXAMEN.BITESTATUSACTIVO = 0 AND LAEXAMEN.INTCVEEXAMEN = " & Trim(rsGrupoExamenInactivo!IntCveExamen), adLockOptimistic, adOpenDynamic)
                                                    If rsExamenInactivo.RecordCount <> 0 Then
                                                        Call MsgBox("El grupo " & Trim(.TextMatrix(vlintRow, 2)) & " contiene uno o más exámenes que están inactivos.", vbExclamation, "Mensaje")
                                                        Exit Sub
                                                    End If
                                                    rsGrupoExamenInactivo.MoveNext
                                                Loop
                                            End If
                                       End If
                                  End Select
                                End If
                        Next vlintRow
                    End With
                    End If
                     vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba <> 0 Then
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                    
                    pEjecutaSentencia "DELETE FROM PVDETALLEGRUPOCARGO WHERE IntCveGrupo = " & TxtCveGrupo.Text
                    
                    If vgstrEstadoManto = "A" Then rsPvGrupoCargo.AddNew
        
                    rsPvGrupoCargo!vchNombre = Trim(txtDescripcion.Text)
                    rsPvGrupoCargo!CHRTIPO = IIf(OptArticulos.Value, "AR", IIf(OptMedicamentos.Value, "ME", IIf(OptImagenologia.Value, "ES", IIf(OptOtros.Value, "OC", "EX"))))
                    rsPvGrupoCargo!bitactivo = chkActivo.Value
                    rsPvGrupoCargo.Update
                    
                    If vgstrEstadoManto = "A" Then
                        vlintIdentity = flngObtieneIdentity("SEC_PVGRUPOCARGO", rsPvGrupoCargo!intCveGrupo)
                    Else
                        vlintIdentity = CLng(TxtCveGrupo.Text)
                    End If
                    
                    With grdCargos
                        .Row = 1
                        For vlintRow = .Row To .Rows - 1
                            If Trim(.TextMatrix(vlintRow, 1)) <> "" Then
                                rsPvDetalleGrupoCargo.AddNew
                                rsPvDetalleGrupoCargo!intCveGrupo = vlintIdentity
                                rsPvDetalleGrupoCargo!intCveCargo = IIf(Int(.TextMatrix(vlintRow, 1)) < 0, Int(.TextMatrix(vlintRow, 1)) * -1, Int(.TextMatrix(vlintRow, 1)))
                                rsPvDetalleGrupoCargo!chrTipoCargo = Trim(.TextMatrix(vlintRow, 3))
                            End If
                        Next vlintRow
                    End With
                    
                    If vlblnBorrarCargos Then
                        pEjecutaSentencia "DELETE FROM PVDETALLEPAQUETE WHERE IntCveCargo = " & TxtCveGrupo.Text & " AND chrTipoCargo = 'GC'"
                    End If
                    
                    On Error GoTo UpdateErr
                    rsPvGrupoCargo.Update
                    rsPvDetalleGrupoCargo.Update
                    On Error GoTo NotificaError
                    
                    If vgstrEstadoManto = "A" Then
                        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "ALTA DE GRUPOS DE CARGOS", TxtCveGrupo.Text)
                    Else
                        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "MODIFICACION DE GRUPOS DE CARGOS", TxtCveGrupo.Text)
                    End If
            
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    rsPvGrupoCargo.Requery
                    rsPvDetalleGrupoCargo.Requery
                    
                    TxtCveGrupo.SetFocus
                End If
            Else
                ' ¡No se ha seleccionado informacion!
                MsgBox SIHOMsg(873), vbExclamation, "Mensaje"
                grdCargos.SetFocus
            End If
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Exit Sub
UpdateErr:
    ' La clave ya existe, la operación no se realizó.
    MsgBox SIHOMsg(649), , "Mensaje"
    If rsPvGrupoCargo.State = 1 Then
        If Not (rsPvGrupoCargo.BOF Or rsPvGrupoCargo.EOF) Then
            rsPvGrupoCargo.CancelUpdate
        End If
    End If
    If rsPvDetalleGrupoCargo.State = 1 Then
        If Not (rsPvDetalleGrupoCargo.BOF Or rsPvDetalleGrupoCargo.EOF) Then
            rsPvDetalleGrupoCargo.CancelUpdate
        End If
    End If
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    pEnfocaTextBox Me.TxtCveGrupo
End Sub

Private Sub cmdSelecciona_Click(Index As Integer)
    If cmdSelecciona(0) Then
        If lstCargos.SelCount <> 0 Then pAgregaCargo
    Else
        pBorraCargo
    End If
End Sub

Private Sub Form_Activate()
    If vgstrEstadoManto = "" Then TxtCveGrupo.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    Dim rscmdConceptoFacturacion As New ADODB.Recordset
    
    Set rsPvGrupoCargo = frsRegresaRs("SELECT * FROM PVGRUPOCARGO ORDER BY intcvegrupo", adLockOptimistic, adOpenDynamic)
    Set rsPvDetalleGrupoCargo = frsRegresaRs("SELECT * FROM PVDETALLEGRUPOCARGO ORDER BY intcvegrupo, chrtipocargo, intcvecargo", adLockOptimistic, adOpenDynamic)

    SSTObj.Tab = 1
    pNuevoRegistro
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    If SSTObj.Tab = 1 Then
        If vgstrEstadoManto <> "" Then
            Cancel = True
            If MsgBox(SIHOMsg(9), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                TxtCveGrupo.SetFocus
            End If
        End If
    Else
        Cancel = True
        SSTObj.Tab = 1
        TxtCveGrupo.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo NotificaError
        rsPvGrupoCargo.Close
        rsPvDetalleGrupoCargo.Close
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
End Sub

Private Sub GrdCargos_DblClick()
    pBorraCargo
End Sub

Private Sub grdHBusqueda_Click()
    With grdHBusqueda
        If .MouseRow = 0 Then
            pLlenaGrid IIf(.MouseCol = 1, 0, IIf(.MouseCol = 2, 1, 2))
        End If
    End With
End Sub

Private Sub grdHBusqueda_DblClick()
    On Error GoTo NotificaError
        If grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1) <> 0 Then
            TxtCveGrupo.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
            txtCveGrupo_KeyDown vbKeyReturn, 0
            SSTObj.Tab = 1
        End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
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
    Dim vgstrSentencia As String
    Dim rsDetalleGrupo As New ADODB.Recordset
    Dim rsBusquedaCargo As New ADODB.Recordset
    
    vgstrEstadoManto = "M"
    
    With rsPvGrupoCargo
        TxtCveGrupo.Text = !intCveGrupo
        txtDescripcion.Text = Trim(!vchNombre)
        chkActivo.Value = !bitactivo
    End With
    
    vgstrSentencia = "SELECT IntCveCargo, ChrTipoCargo FROM PVDETALLEGRUPOCARGO WHERE IntCveGrupo = " & rsPvGrupoCargo!intCveGrupo
    Set rsDetalleGrupo = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rsDetalleGrupo.RecordCount > 0 Then
        pConfiguraGridCargos
        Do While Not rsDetalleGrupo.EOF
            vgstrSentencia = " SELECT " & _
                                " PVDETALLEGRUPOCARGO.intcvecargo * CASE WHEN PVDETALLEGRUPOCARGO.chrtipocargo = 'GE' THEN -1 ELSE 1 END Clave " & _
                                ",CASE PVDETALLEGRUPOCARGO.chrtipocargo" & _
                                    " WHEN 'AR' THEN (SELECT vchnombrecomercial FROM IVARTICULO WHERE intidarticulo = PVDETALLEGRUPOCARGO.intCveCargo AND chrCostoGasto <> 'G' AND chrCveArtMedicamen <> 1) " & _
                                    " WHEN 'ME' THEN (SELECT vchnombrecomercial FROM IVARTICULO WHERE intidarticulo = PVDETALLEGRUPOCARGO.intCveCargo AND chrCostoGasto <> 'G' AND chrCveArtMedicamen = 1) " & _
                                    " WHEN 'OC' THEN (SELECT chrdescripcion FROM PVOTROCONCEPTO WHERE intcveconcepto = PVDETALLEGRUPOCARGO.intCveCargo) " & _
                                    " WHEN 'ES' THEN (SELECT vchnombre FROM IMESTUDIO WHERE intcveestudio = PVDETALLEGRUPOCARGO.intCveCargo) " & _
                                    " WHEN 'EX' THEN (SELECT chrnombre FROM LAEXAMEN WHERE intcveexamen = PVDETALLEGRUPOCARGO.intCveCargo) " & _
                                    " WHEN 'GE' THEN (SELECT chrnombre FROM LAGRUPOEXAMEN WHERE intcvegrupo = PVDETALLEGRUPOCARGO.intCveCargo) " & _
                                " END Nombre " & _
                                ",PVDETALLEGRUPOCARGO.chrtipocargo Tipo " & _
                             " FROM PVGRUPOCARGO " & _
                             " LEFT JOIN PVDETALLEGRUPOCARGO ON PVGRUPOCARGO.IntcveGrupo = PVDETALLEGRUPOCARGO.IntcveGrupo " & _
                             " WHERE PVGRUPOCARGO.intCveGrupo = " & rsPvGrupoCargo!intCveGrupo
            Set rsBusquedaCargo = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
            Call pLlenarMshFGrdRs(grdCargos, rsBusquedaCargo)
            
            rsDetalleGrupo.MoveNext
        Loop
    End If
    
    pConfiguraGridCargos
    rsDetalleGrupo.Close
    rsBusquedaCargo.Close
    
    pHabilitaComponentesCaptura True
    pHabilitaTipos IIf(rsPvGrupoCargo!CHRTIPO = "AR", 1, IIf(rsPvGrupoCargo!CHRTIPO = "ME", 2, IIf(rsPvGrupoCargo!CHRTIPO = "ES", 3, IIf(rsPvGrupoCargo!CHRTIPO = "EX", 4, 5))))
    pHabilitaElementos 2
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pModificaRegistro"))
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1) <> 0 Then
            TxtCveGrupo.Text = grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)
            txtCveGrupo_KeyDown vbKeyReturn, 0
            SSTObj.Tab = 1
        End If
    End If
End Sub

Private Sub lstCargos_DblClick()
    pAgregaCargo
End Sub

Private Sub lstCargos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And lstCargos.SelCount <> 0 Then pAgregaCargo
End Sub

Private Sub OptArticulos_Click()
    pLimpiaBusqueda
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub OptArticulos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSeleArticulo.SetFocus
End Sub

Private Sub optClave_Click()
    pLimpiaBusqueda
    txtSeleArticulo.SetFocus
End Sub

Private Sub optDescripcion_Click()
    pLimpiaBusqueda
    txtSeleArticulo.SetFocus
End Sub

Private Sub OptImagenologia_Click()
    pLimpiaBusqueda
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub OptImagenologia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSeleArticulo.SetFocus
End Sub

Private Sub OptLaboratorio_Click()
    pLimpiaBusqueda
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub OptLaboratorio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSeleArticulo.SetFocus
End Sub

Private Sub OptMedicamentos_Click()
    pLimpiaBusqueda
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub OptMedicamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSeleArticulo.SetFocus
End Sub

Private Sub OptOtros_Click()
    pLimpiaBusqueda
    If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
End Sub

Private Sub OptOtros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSeleArticulo.SetFocus
End Sub



Private Sub txtCveGrupo_GotFocus()
    On Error GoTo NotificaError
        pNuevoRegistro
        If TxtCveGrupo.Enabled And TxtCveGrupo.Visible Then pEnfocaTextBox TxtCveGrupo
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveGrupo_GotFocus"))
End Sub

Private Sub txtCveGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsBusca As ADODB.Recordset
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        'Buscar criterio
        If TxtCveGrupo.Text = "" Then
            TxtCveGrupo.Text = Trim(Str(lngSig))
        End If
        
        If lngSig = CLng(TxtCveGrupo.Text) Then
            vgstrEstadoManto = "A" 'Alta
            Call pEnfocaTextBox(txtDescripcion)
            pHabilitaComponentesCaptura True
            pHabilitaTipos 0
            pHabilitaElementos 1
            txtDescripcion.SetFocus
        Else
            If fintLocalizaPkRs(rsPvGrupoCargo, 0, TxtCveGrupo.Text) > 0 Then
                pHabilitaComponentesCaptura True
                pModificaRegistro
                vgstrEstadoManto = "M" 'Modificacion
                If SSTObj.Tab = 1 Then
                    Call pEnfocaTextBox(txtDescripcion)
                Else
                    cmdBuscar.SetFocus
                End If
            Else
                MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                TxtCveGrupo.Text = Trim(Str(lngSig))
                Call pEnfocaTextBox(TxtCveGrupo)
                rsPvGrupoCargo.Requery
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveGrupo_KeyDown"))
End Sub

Private Sub txtCveGrupo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveGrupo_KeyPress"))
End Sub

Private Sub txtDescripcion_GotFocus()
    pEnfocaTextBox txtDescripcion
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If OptArticulos.Enabled Then
            OptArticulos.SetFocus
        Else
            If OptMedicamentos.Enabled Then
                OptMedicamentos.SetFocus
            Else
                If OptImagenologia.Enabled Then
                    OptImagenologia.SetFocus
                Else
                    If OptLaboratorio.Enabled Then
                        OptLaboratorio.SetFocus
                    Else
                        If OptOtros.Enabled Then OptOtros.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
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
    
    Call pPosicionaRegRs(rsPvGrupoCargo, "A")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGrupoCargo, "I")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
End Sub

Private Sub cmdSiguienteRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGrupoCargo, "S")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
End Sub

Private Sub cmdUltimoRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvGrupoCargo, "U")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = vbKeyEscape Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub pHabilitaComponentesCaptura(vlblnHabilita As Boolean)
' Habilita o deshabilita los elementos capturables de la pantalla
    txtDescripcion.Enabled = vlblnHabilita
    OptArticulos.Enabled = vlblnHabilita
    OptMedicamentos.Enabled = vlblnHabilita
    OptImagenologia.Enabled = vlblnHabilita
    OptLaboratorio.Enabled = vlblnHabilita
    OptOtros.Enabled = vlblnHabilita
    FrmBusqueda.Enabled = vlblnHabilita
    FrmMovimientos.Enabled = vlblnHabilita
    grdCargos.Enabled = vlblnHabilita
    chkActivo.Enabled = vlblnHabilita
End Sub

Private Sub pHabilitaTipos(vlintTipo As Integer)
' Habilitar o deshabilitar los tipos de cargos
' 0 = Ninguno, 1 = Articulos, 2 = Medicamentos, 3 = Servicios auxiliares, 4 = Laboratorio y 5 = Otros conceptos
    
    On Error GoTo NotificaError

    OptArticulos.Value = IIf(vlintTipo = 1, True, OptArticulos.Value)
    OptMedicamentos.Value = IIf(vlintTipo = 2, True, OptMedicamentos.Value)
    OptImagenologia.Value = IIf(vlintTipo = 3, True, OptImagenologia.Value)
    OptLaboratorio.Value = IIf(vlintTipo = 4, True, OptLaboratorio.Value)
    OptOtros.Value = IIf(vlintTipo = 5, True, OptOtros.Value)
    OptArticulos.Enabled = IIf(vlintTipo <> 0, IIf(vlintTipo = 1, True, False), True)
    OptMedicamentos.Enabled = IIf(vlintTipo <> 0, IIf(vlintTipo = 2, True, False), True)
    OptImagenologia.Enabled = IIf(vlintTipo <> 0, IIf(vlintTipo = 3, True, False), True)
    OptLaboratorio.Enabled = IIf(vlintTipo <> 0, IIf(vlintTipo = 4, True, False), True)
    OptOtros.Enabled = IIf(vlintTipo <> 0, IIf(vlintTipo = 5, True, False), True)
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub txtSeleArticulo_GotFocus()
    pEnfocaTextBox txtSeleArticulo
End Sub

Private Sub txtSeleArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        If lstCargos.Enabled Then lstCargos.SetFocus
    End If
End Sub

Private Sub txtSeleArticulo_KeyPress(KeyAscii As Integer)
    If optClave.Value Then
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtSeleArticulo_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vlstrSentencia As String
    Dim vlstrOtroFiltro As String
    
    vlstrOtroFiltro = " "
    vlstrSentencia = ""
    
    If OptArticulos.Value Then
        vlstrOtroFiltro = " and chrCostoGasto <> 'G' and vchEstatus = 'ACTIVO' and chrCveArtMedicamen <> '1'"
    End If
    
    If OptMedicamentos.Value Then
        vlstrOtroFiltro = " and chrCostoGasto <> 'G' and vchEstatus = 'ACTIVO' and chrCveArtMedicamen = '1'"
    End If
    
    If OptArticulos.Value Or OptMedicamentos.Value Then
        vlstrSentencia = "SELECT intIDArticulo, vchNombreComercial FROM ivarticulo"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstCargos, IIf(optDescripcion.Value, "vchNombreComercial", "chrCveArticulo"), 1000, vlstrOtroFiltro, "vchNombreComercial"
    End If
    
    If OptImagenologia.Value Then
        vlstrSentencia = "SELECT intCveEstudio, vchNombre FROM IMESTUDIO"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstCargos, IIf(optDescripcion.Value, "vchNombre", "intCveEstudio"), 1000, " and bitStatusActivo = 1", "vchNombre"
    End If
    
    If OptLaboratorio.Value Then
        vlstrSentencia = "SELECT * FROM (SELECT intCveExamen Clave, chrNombre FROM LAEXAMEN WHERE bitEstatusActivo = 1 UNION SELECT (intCveGrupo * -1) Clave, chrNombre FROM LAGRUPOEXAMEN WHERE bitEstatusActivo = 1)"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstCargos, IIf(optDescripcion.Value, "chrNombre", "Clave"), 1000, IIf(optDescripcion.Value, vlstrOtroFiltro, " or Clave like '-" & txtSeleArticulo & "%'"), "chrNombre"
    End If
    
    If OptOtros.Value Then
        vlstrSentencia = "SELECT intCveConcepto, chrDescripcion FROM PVOTROCONCEPTO"
        PSuperBusqueda txtSeleArticulo, vlstrSentencia, lstCargos, IIf(optDescripcion.Value, "chrDescripcion", "intCveConcepto"), 1000, " and bitEstatus = 1", "chrDescripcion"
    End If
    
End Sub

Private Sub pLimpiaBusqueda()
' Limpia la busqueda de cargos
    On Error GoTo NotificaError
        txtSeleArticulo.Text = ""
        lstCargos.Clear
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaBusqueda"))
End Sub

Private Sub pAgregaCargo()
' Agrega elemento al grid de cargos
    Dim rsElemenRepetido As ADODB.Recordset
    Dim vlstrCadenaMsj As String
    Dim vlintcontador As Integer
    
    On Error GoTo NotificaError
                                       
    Set rsElemenRepetido = frsRegresaRs("SELECT MPQ.intNumPaquete CvePaquete, TRIM(MPQ.chrDescripcion) DescPaquete, MGC.intcveGrupo CveGrupo, TRIM(MGC.vchNombre) DescGrupo, DGC.intcvecargo Cargo, DGC.chrtipoCargo Tipo " & _
                                        "FROM PVPAQUETE MPQ " & _
                                            "INNER JOIN PVDETALLEPAQUETE DPQ ON DPQ.intNumPaquete = MPQ.intNumPaquete " & _
                                                "AND DPQ.chrtipocargo = 'GC' " & _
                                            "INNER JOIN PVGRUPOCARGO MGC ON MGC.intcvegrupo = DPQ.intcvecargo " & _
                                            "LEFT JOIN PVDETALLEGRUPOCARGO DGC ON DGC.intcvegrupo = MGC.intcvegrupo " & _
                                        "WHERE MPQ.intNumPaquete IN (SELECT intnumpaquete " & _
                                                                    "FROM PVDETALLEPAQUETE " & _
                                                                    "WHERE chrtipocargo = 'GC' " & _
                                                                        "AND intCveCargo = " & CLng(TxtCveGrupo.Text) & ") " & _
                                        "AND MGC.intcveGrupo <> " & CLng(TxtCveGrupo.Text) & " " & _
                                        "AND DGC.intcvecargo = " & lstCargos.ItemData(lstCargos.ListIndex) & " " & _
                                        "AND DGC.chrtipoCargo = '" & IIf(OptArticulos.Value, "AR", IIf(OptMedicamentos.Value, "ME", IIf(OptImagenologia.Value, "ES", IIf(OptOtros.Value, "OC", IIf(lstCargos.ItemData(lstCargos.ListIndex) < 0, "GE", "EX"))))) & "'" & _
                                        "ORDER BY CvePaquete, CveGrupo", adLockReadOnly, adOpenForwardOnly)
     With rsElemenRepetido
        If .RecordCount > 0 Then
            .MoveFirst
            For vlintcontador = 1 To .RecordCount
                vlstrCadenaMsj = vlstrCadenaMsj & Chr(13) & Format(!cvePaquete, "########") & " " & !DescPaquete & Chr(13) & "     " & !cveGrupo & " " & !DescGrupo
                .MoveNext
            Next vlintcontador
            MsgBox SIHOMsg(1102) & vlstrCadenaMsj, vbOKOnly + vbInformation, "Mensaje"
            txtSeleArticulo.SetFocus
            Exit Sub
        End If
    End With
    
    If pRevisaRepetidos = False Then
        With grdCargos
            .Row = .Rows - 1
            If .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .TextMatrix(.Row, 1) = lstCargos.ItemData(lstCargos.ListIndex)
            .TextMatrix(.Row, 2) = lstCargos.Text
            .TextMatrix(.Row, 3) = IIf(OptArticulos.Value, "AR", IIf(OptMedicamentos.Value, "ME", IIf(OptImagenologia.Value, "ES", IIf(OptOtros.Value, "OC", IIf(lstCargos.ItemData(lstCargos.ListIndex) < 0, "GE", "EX")))))
            pRevisaGrid
        End With
        
        If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
        
    Else
        ' ¡El elemento ya está asignado!
        MsgBox SIHOMsg(354), vbOKOnly + vbCritical, "Mensaje"
        txtSeleArticulo.SetFocus
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAgregaCargo"))
End Sub

Private Sub pBorraCargo()
' Procedimiento que manda borrar elementos del grid de cargos
    On Error GoTo NotificaError
        With grdCargos
            .Redraw = False
            grdCargos = fmskBorrarReg(.Row, grdCargos)
            .Redraw = True
            .Refresh
            pRevisaGrid
            If Not cmdGrabarRegistro.Enabled Then pHabilitaElementos 1
        End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pBorraCargo"))
End Sub

Private Function fmskBorrarReg(vllngRenglon As Long, grdNombre As MSHFlexGrid) As MSHFlexGrid
' Borra elementos del grid de cargos
    Dim vllngContador As Long, vllngContador1 As Long
                    
    With grdNombre
        If .Rows > 2 Then
            For vllngContador = .Row + 1 To .Rows - 1
                .RowData(vllngContador - 1) = .RowData(vllngContador)
                For vllngContador1 = 0 To .Cols - 1
                    .TextMatrix(vllngContador - 1, vllngContador1) = .TextMatrix(vllngContador, vllngContador1)
                Next vllngContador1
            Next vllngContador
        Else
            If .Rows = 2 Then
                .RowData(1) = -1
                For vllngContador = 0 To .Cols - 1
                    .TextMatrix(1, vllngContador) = ""
                Next vllngContador
            End If
        End If
        If .Rows > 2 Then
            .Rows = .Rows - 1
            .Row = .Rows - 1
        End If
  End With
  Set fmskBorrarReg = grdNombre

End Function

Private Sub pRevisaGrid()
    On Error GoTo NotificaError
    
    With grdCargos
        If .Rows >= 3 Or (.Rows = 2 And .TextMatrix(.Row, 1) <> "") Then
            OptArticulos.Enabled = IIf(OptArticulos.Value, True, False)
            OptMedicamentos.Enabled = IIf(OptMedicamentos.Value, True, False)
            OptImagenologia.Enabled = IIf(OptImagenologia.Value, True, False)
            OptLaboratorio.Enabled = IIf(OptLaboratorio.Value, True, False)
            OptOtros.Enabled = IIf(OptOtros.Value, True, False)
        Else
            OptArticulos.Enabled = True
            OptMedicamentos.Enabled = True
            OptImagenologia.Enabled = True
            OptLaboratorio.Enabled = True
            OptOtros.Enabled = True
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRevisaGrid"))
End Sub

Private Function pRevisaRepetidos() As Boolean
    On Error GoTo NotificaError
    Dim vlintRenglon As Integer
    Dim vllngContador As Integer
    
    pRevisaRepetidos = False
    With grdCargos
        vlintRenglon = .Row
        .Row = 1
        For vllngContador = .Row To .Rows - 1
            If Trim(.TextMatrix(vllngContador, 1)) = Trim(Str(lstCargos.ItemData(lstCargos.ListIndex))) Then
                pRevisaRepetidos = True
                Exit Function
            End If
        Next vllngContador
        .Row = vlintRenglon
    End With
    
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pRevisaRepetidos"))
End Function

Private Sub pHabilitaElementos(vlintProceso As Integer)
'0 = Inicio (determinara si hay o no datos que buscar), 1 = Captura, 2 = Consulta
    On Error GoTo NotificaError
    
    cmdPrimerRegistro.Enabled = IIf(vlintProceso = 0, IIf(rsPvGrupoCargo.RecordCount > 0, 1, 0), IIf(vlintProceso = 1, 0, 1))
    cmdAnteriorRegistro.Enabled = IIf(vlintProceso = 0, IIf(rsPvGrupoCargo.RecordCount > 0, 1, 0), IIf(vlintProceso = 1, 0, 1))
    cmdBuscar.Enabled = IIf(vlintProceso = 0, IIf(rsPvGrupoCargo.RecordCount > 0, 1, 0), IIf(vlintProceso = 1, 0, 1))
    cmdSiguienteRegistro.Enabled = IIf(vlintProceso = 0, IIf(rsPvGrupoCargo.RecordCount > 0, 1, 0), IIf(vlintProceso = 1, 0, 1))
    cmdUltimoRegistro.Enabled = IIf(vlintProceso = 0, IIf(rsPvGrupoCargo.RecordCount > 0, 1, 0), IIf(vlintProceso = 1, 0, 1))
    cmdGrabarRegistro.Enabled = IIf(vlintProceso = 0, 0, IIf(vlintProceso = 1, 1, 0))
    cmdDelete.Enabled = IIf(vlintProceso = 0, 0, IIf(vlintProceso = 1, 0, 1))
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilitaElementos"))
End Sub
