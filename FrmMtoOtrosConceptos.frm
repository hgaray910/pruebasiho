VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmMtoOtrosConceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Otros conceptos"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTobj 
      Height          =   7680
      Left            =   -600
      TabIndex        =   25
      Top             =   -465
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   13547
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmMtoOtrosConceptos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmMtoOtrosConceptos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optOrden(1)"
      Tab(1).Control(1)=   "optOrden(0)"
      Tab(1).Control(2)=   "GrdHBusqueda"
      Tab(1).ControlCount=   3
      Begin VB.OptionButton optOrden 
         Caption         =   "Descripción"
         Height          =   270
         Index           =   1
         Left            =   -73320
         TabIndex        =   8
         Top             =   585
         Width           =   1230
      End
      Begin VB.OptionButton optOrden 
         Caption         =   "Clave"
         Height          =   270
         Index           =   0
         Left            =   -74190
         TabIndex        =   7
         Top             =   585
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.Frame Frame4 
         Height          =   690
         Left            =   2640
         TabIndex        =   18
         Top             =   6150
         Width           =   3615
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   480
            Left            =   3045
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Borrar el registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":01DA
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Primer registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":034C
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anterior registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1065
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":04BE
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0630
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Siguiente registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2055
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":07A2
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Ultimo registro"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Height          =   480
            Left            =   2550
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0914
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Guardar el registro"
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5610
         Left            =   705
         TabIndex        =   14
         Top             =   480
         Width           =   7400
         Begin VB.TextBox txtFechaAltaOtroConcepto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   6195
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "01/Ene/1999"
            Top             =   1575
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ListBox lstDepartamentos 
            Height          =   2985
            Left            =   240
            TabIndex        =   26
            ToolTipText     =   "Departamentos del hospital"
            Top             =   2280
            Width           =   3030
         End
         Begin VB.ListBox lstDepartamentosSel 
            Height          =   2985
            Left            =   4125
            TabIndex        =   24
            ToolTipText     =   "Departamentos asignados al concepto."
            Top             =   2280
            Width           =   3030
         End
         Begin VB.CommandButton cmdEliminaTodo 
            Height          =   480
            Left            =   3450
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0C56
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Primer registro"
            Top             =   4335
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin VB.CommandButton cmdEliminaUno 
            Height          =   480
            Left            =   3450
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0DC8
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Anterior registro"
            Top             =   3810
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin VB.CommandButton cmdAsignaUno 
            Height          =   480
            Left            =   3450
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":0F3A
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Siguiente registro"
            Top             =   3300
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin VB.CommandButton cmdAsignaTodo 
            Height          =   480
            Left            =   3450
            MaskColor       =   &H00FFFFFF&
            Picture         =   "FrmMtoOtrosConceptos.frx":10AC
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ultimo registro"
            Top             =   2760
            UseMaskColor    =   -1  'True
            Width           =   480
         End
         Begin VB.CheckBox ChkStatus 
            Caption         =   "Activo"
            Height          =   195
            Left            =   2160
            TabIndex        =   3
            ToolTipText     =   "Estado"
            Top             =   1575
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.ComboBox CboConceptoFacturacion 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Concepto de facturación"
            Top             =   1155
            Width           =   5010
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   315
            Left            =   2160
            MaxLength       =   3900
            TabIndex        =   1
            ToolTipText     =   "Descripción"
            Top             =   795
            Width           =   5010
         End
         Begin VB.TextBox TxtCveOConcepto 
            Height          =   315
            Left            =   2160
            MaxLength       =   5
            TabIndex        =   0
            ToolTipText     =   "Clave"
            Top             =   435
            Width           =   735
         End
         Begin VB.Label lblFechaAltaOtroConcepto 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha de alta "
            Height          =   195
            Left            =   5040
            TabIndex        =   29
            Top             =   1575
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Departamentos"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Departamentos asignados"
            Height          =   195
            Left            =   4125
            TabIndex        =   27
            Top             =   2040
            Width           =   1845
         End
         Begin VB.Label Label5 
            Caption         =   "Estado"
            Height          =   255
            Left            =   240
            TabIndex        =   19
            Top             =   1545
            Width           =   1935
         End
         Begin VB.Label Label4 
            Caption         =   "Concepto de facturación"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1185
            Width           =   1935
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   825
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Clave"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   435
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdHBusqueda 
         Height          =   5895
         Left            =   -74250
         TabIndex        =   13
         Top             =   915
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   10398
         _Version        =   393216
         FocusRect       =   0
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmMtoOtrosConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Punto de Venta
'| Nombre del Formulario    : frmOtroConcepto
'-------------------------------------------------------------------------------------
'| Objetivo: Realiza el mantenimiento del catálogo de Otros conceptos
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Jose Torres
'| Fecha de Creación        : 03/Diciembre/2000
'| Modificó                 : Nombre(s)
'| Fecha última modificación: dd/mes/AAAA
'-------------------------------------------------------------------------------------

Option Explicit 'Permite forzar la declaración de las variables
Public vgblnCargarTodosConceptos As Boolean

Dim vgstrEstadoManto As String
Dim rsPvOtroConcepto As New ADODB.Recordset

Dim vlblnActivaraCaptura As Boolean

Dim vlstrx As String
Dim blnClaveManualCatalogo As Boolean
Dim lngSig As Long
Dim vgstrSentencia As String
Dim vgintCont As Integer
Dim blnEnfocando As Boolean

Dim vlstrSentencia As String
Dim rsOtroConceptos As New ADODB.Recordset
Dim vlstrOrden As String
Dim blnLlenagrid As Boolean
Dim strFechaAltaOtroConcepto As String

Private Sub pCargarDepartamentos()
    Dim rsDatos As New ADODB.Recordset
      
    vgstrSentencia = "Select smiCveDepartamento Cve, Rtrim(vchDescripcion) Nombre from NoDepartamento " & _
                     "where BITESTATUS = 1 Order By Nombre"
        
    Set rsDatos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rsDatos.RecordCount > 0 Then
        With lstDepartamentos
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
        .FormatString = "|Clave|Descripción||||||Concepto de facturación"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 700  'Clave
        .ColWidth(2) = 4800 'Descripcion
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 3000 'Concepto de facturacion
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub

Private Sub pNuevoRegistro()
    On Error GoTo NotificaError
    
    Dim rsSig As ADODB.Recordset
    
    pCargarDepartamentos
    
    vgstrEstadoManto = ""
    txtDescripcion.Text = ""
    lstDepartamentosSel.Clear
    
    Set rsSig = frsRegresaRs("select max(intCveConcepto) from PvOtroConcepto", adLockReadOnly, adOpenForwardOnly)
    
    If Not rsSig.EOF Then
        lngSig = rsSig.Fields(0).Value + 1
    Else
        lngSig = 1
    End If
    
    If cboConceptoFacturacion.ListCount > 0 Then
        cboConceptoFacturacion.ListIndex = 0
        TxtCveOConcepto.Text = lngSig
        pHabilitaBotonModifica (True)
        cmdBuscar.Enabled = True
        cmdDelete.Enabled = False
        pHabilitaComponentesCaptura False
    Else
        MsgBox SIHOMsg(13) + Chr(13) + cboConceptoFacturacion.ToolTipText, vbExclamation, "Mensaje"
        Unload Me
        Exit Sub
    End If
    pLimpiaFechaAltaOtroConcepto
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pNuevoRegistro"))
End Sub

Private Sub pLlenaGrid(vlintOrden As Integer)
    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim rsOtroConceptos As New ADODB.Recordset
    Dim vlstrOrden As String
    
    
    GrdHBusqueda.Clear
    If vlintOrden = 1 Then
            vlstrOrden = " PvOtroConcepto.chrDescripcion"
            
    Else
            vlstrOrden = " PvOtroConcepto.intCveConcepto"
          
    End If

    If vgblnCargarTodosConceptos Then
        vlstrSentencia = "SELECT PvOtroConcepto.intCveConcepto, " & _
        "rtrim(PvOtroConcepto.chrDescripcion), " & _
        "rtrim(PvConceptoFacturacion.chrDescripcion) " & _
        "FROM PvConceptoFacturacion INNER JOIN " & _
        "PvOtroConcepto ON " & _
        "PvConceptoFacturacion.smiCveConcepto = PvOtroConcepto.smiConceptoFact " & _
        "WHERE bitActivo = 1 order by " & vlstrOrden
    Else
       
       vlstrSentencia = "SELECT PvOtroConcepto.intCveConcepto intCveConcepto, " & _
        "rtrim(PvOtroConcepto.chrDescripcion) chrDescripcion, " & _
        "rtrim(PvConceptoFacturacion.chrDescripcion) chrDescripcionF, " & _
        "rtrim(PvOtroConcepto.SMICONCEPTOFACT) SMICONCEPTOFACT, " & _
        "PvOtroConcepto.BITESTATUS BITESTATUS " & _
        "FROM PvConceptoFacturacion INNER JOIN " & _
        "PvOtroConcepto ON " & _
        "PvConceptoFacturacion.smiCveConcepto = PvOtroConcepto.smiConceptoFact " & _
        "where bitActivo = 1 and pvOtroConcepto.smiDepartamento = " & Trim(str(vgintNumeroDepartamento)) & _
        " order by " & vlstrOrden
       
    End If
    
    Set rsOtroConceptos = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    pLlenarMshFGrdRs GrdHBusqueda, rsOtroConceptos, 0
    
    pConfiguraGrid
    rsOtroConceptos.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pllenaGrid"))
End Sub

Private Sub CboConceptoFacturacion_GotFocus()
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False
End Sub

Private Sub cboConceptoFacturacion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        lstDepartamentos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":CboConceptoFacturacion_KeyDown"))
End Sub


Private Sub chkStatus_GotFocus()
    pHabilitaBotonModifica (False)
    cmdBuscar.Enabled = False

End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo NotificaError
    
    blnLlenagrid = True
    
    sstObj.Tab = 1
    
    If optOrden(0).Value Then
       optOrden(0).Value = True
       pLlenaGridOtroConcepto 2
    Else
       optOrden(1).Value = True
       pLlenaGridOtroConcepto 1
    End If
    
    GrdHBusqueda.SetFocus
   

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError
    
    If fblnIntegridadValida() Then
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "OTROS CONCEPTOS", TxtCveOConcepto.Text)
       
        vlstrx = " Delete from PvOtroConcepto where PvOtroConcepto.INTCVECONCEPTO = " & Trim(TxtCveOConcepto.Text)
        pEjecutaSentencia vlstrx
        
        rsPvOtroConcepto.Update
        
        vlstrx = "Delete from  PvDetalleLista where chrCveCargo=" + "'" + Trim(TxtCveOConcepto.Text) + "'" + " and chrTipoCargo='OC'"
        pEjecutaSentencia vlstrx
        
        vlstrx = " Delete from PvOtroConceptoDepto where intCveOtroConcepto = " & Val(TxtCveOConcepto.Text)
        
        
        pEjecutaSentencia vlstrx
                
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        TxtCveOConcepto.SetFocus
    Else
        '!No se pueden borrar los datos!
        MsgBox SIHOMsg(257), vbOKOnly + vbCritical, "Mensaje"
    End If
    
    vlblnActivaraCaptura = True
    rsPvOtroConcepto.Requery
    pLimpiaFechaAltaOtroConcepto
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Function fblnIntegridadValida() As Boolean
    On Error GoTo NotificaError
    '-------------------------------------------------------
    ' Si ya se realizó un cargo o el cargo esta contenido en
    ' un paquete, no se puede borrar
    '-------------------------------------------------------
    Dim vlstrx As String
    Dim rsRegistros As New ADODB.Recordset
    Dim vlintAux As Integer
    
    fblnIntegridadValida = True
    
    vlintAux = 0
    vlstrx = "select count(*) " & _
             "From PvCargo " & _
             "where chrCveCargo=" & "'" & Trim(TxtCveOConcepto.Text) & "'" & "and chrTipoCargo='OC'"
    Set rsRegistros = frsRegresaRs(vlstrx)
    vlintAux = rsRegistros.Fields(0).Value

    vlstrx = "select count(*) " & _
             "From PvDetallePaquete " & _
             "where intCveCargo=" & "'" & Trim(TxtCveOConcepto.Text) & "'" & "and chrTipoCargo='OC'"
    Set rsRegistros = frsRegresaRs(vlstrx)
    vlintAux = vlintAux + rsRegistros.Fields(0).Value
    If vlintAux <> 0 Then
        fblnIntegridadValida = False
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnIntegridadValida"))
End Function

Private Sub cmdGrabarRegistro_Click()
    Dim vlrsAux As New ADODB.Recordset
    Dim vllngClaveOtroConcepto As Long
    Dim vlintRow As Integer
    Dim lvsqlstr As String
    Dim rsCargoAsignadoACuarto As New ADODB.Recordset 'Valida que el otro concepto no este asignado como cargo a uno o más cuartos
    Dim rsConcEnGrupoCargos As New ADODB.Recordset  'Valida que el otro concepto no esté asignado a uno o más grupos de cargos
    
    On Error GoTo NotificaError
    If RTrim(txtDescripcion.Text) = "" Then
        MsgBox SIHOMsg(2), vbExclamation, "Mensaje"
        txtDescripcion.SetFocus
    ElseIf cboConceptoFacturacion.ListIndex = -1 Then
        MsgBox "Debe de seleccionar un concepto de facturación.", vbExclamation, "Mensaje"
        cboConceptoFacturacion.SetFocus
    Else
    If vgstrEstadoManto <> "A" And ChkStatus.Value = 0 Then
        lvsqlstr = "SELECT DISTINCT PVOTROCONCEPTO.INTCVECONCEPTO FROM PVDETALLEGRUPOCARGO INNER JOIN PVGRUPOCARGO ON PVGRUPOCARGO.INTCVEGRUPO = PVDETALLEGRUPOCARGO.INTCVEGRUPO INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = PVDETALLEGRUPOCARGO.INTCVECARGO WHERE PVGRUPOCARGO.BITACTIVO = 1 AND PVDETALLEGRUPOCARGO.CHRTIPOCARGO = 'OC' AND PVOTROCONCEPTO.INTCVECONCEPTO = " & Trim(TxtCveOConcepto.Text)
        Set rsConcEnGrupoCargos = frsRegresaRs(lvsqlstr, adLockOptimistic, adOpenDynamic)
        If rsConcEnGrupoCargos.RecordCount <> 0 Then
            Call MsgBox("El otro concepto se encuentra asignado a uno o más grupos de cargos.", vbExclamation, "Mensaje")
            Exit Sub
        End If
        lvsqlstr = "SELECT DISTINCT PVOTROCONCEPTO.INTCVECONCEPTO FROM ADCUARTO INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = ADCUARTO.INTOTROCONCEPTO WHERE PVOTROCONCEPTO.BITESTATUS = 1 AND PVOTROCONCEPTO.INTCVECONCEPTO = " & Trim(TxtCveOConcepto.Text) & " UNION ALL SELECT DISTINCT PVOTROCONCEPTO.INTCVECONCEPTO FROM ADCUARTO INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = ADCUARTO.INTCVECONCEPTOMEDIAESTANCIA WHERE PVOTROCONCEPTO.BITESTATUS = 1 AND PVOTROCONCEPTO.INTCVECONCEPTO = " & Trim(TxtCveOConcepto.Text)
        Set rsCargoAsignadoACuarto = frsRegresaRs(lvsqlstr, adLockOptimistic, adOpenDynamic)
        If rsCargoAsignadoACuarto.RecordCount <> 0 Then
            Call MsgBox("El otro concepto se encuentra asignado a uno o más cuartos.", vbExclamation, "Mensaje")
            Exit Sub
        End If
    End If
    
    EntornoSIHO.ConeccionSIHO.BeginTrans
        If vgstrEstadoManto = "A" Then
            rsPvOtroConcepto.AddNew
        End If
        If rsPvOtroConcepto.Fields("intCveConcepto").Attributes <> 16 And rsPvOtroConcepto.Fields("intCveConcepto").Attributes <> 32784 Then
            rsPvOtroConcepto!intCveConcepto = TxtCveOConcepto.Text
        End If
        rsPvOtroConcepto!chrDescripcion = txtDescripcion.Text
        rsPvOtroConcepto!SMICONCEPTOFACT = cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex)
        Set vlrsAux = frsRegresaRs("Select intCveDepartamento From PvConceptoFacturacionEmpresa Where intCveConceptoFactura = " & CStr(cboConceptoFacturacion.ItemData(cboConceptoFacturacion.ListIndex)) & " And intCveEmpresaContable = " & CStr(vgintClaveEmpresaContable))
        rsPvOtroConcepto!SMIDEPARTAMENTO = vlrsAux.Fields(0).Value
        rsPvOtroConcepto!BITESTATUS = ChkStatus.Value
        If vgstrEstadoManto = "A" Then
            rsPvOtroConcepto!DTMFECHAALTAOTROCONCEPTO = Now
        End If
        On Error GoTo UpdateErr
        rsPvOtroConcepto.Update
        On Error GoTo NotificaError
        
        pEjecutaSentencia ("{call SP_GNINSCARGOLISTASPRECIOS(" & vlrsAux.Fields(0).Value & ", " & Trim(TxtCveOConcepto.Text) & ", 'OC')}")
        
        If vgstrEstadoManto = "A" Then
            vllngClaveOtroConcepto = CDbl(TxtCveOConcepto.Text)
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "OTROS CONCEPTOS", TxtCveOConcepto.Text)
        Else
            vllngClaveOtroConcepto = CDbl(TxtCveOConcepto.Text)
            Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "OTROS CONCEPTOS", TxtCveOConcepto.Text)
        End If

        'Se borra el Detalle
        vgstrSentencia = "Delete from  PvOtroConceptoDepto Where intCveOtroConcepto = " & vllngClaveOtroConcepto
        pEjecutaSentencia vgstrSentencia

        For vlintRow = 0 To lstDepartamentosSel.ListCount - 1
            'Se Graban los datos
            vgstrSentencia = "Insert into PvOtroConceptoDepto Values(" & vllngClaveOtroConcepto & "," & lstDepartamentosSel.ItemData(vlintRow) & ")"
            pEjecutaSentencia vgstrSentencia
        Next vlintRow

        EntornoSIHO.ConeccionSIHO.CommitTrans
        rsPvOtroConcepto.Requery
        
        TxtCveOConcepto.SetFocus
        pLimpiaFechaAltaOtroConcepto
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdGrabarRegistro_Click"))
    Exit Sub
UpdateErr:
    MsgBox SIHOMsg(649), , "Mensaje"
    If rsPvOtroConcepto.State = 1 Then
        If Not (rsPvOtroConcepto.BOF Or rsPvOtroConcepto.EOF) Then
            rsPvOtroConcepto.CancelUpdate
        End If
    End If
    EntornoSIHO.ConeccionSIHO.RollbackTrans
    blnEnfocando = True
    pEnfocaTextBox Me.TxtCveOConcepto
End Sub

Private Sub pAsigna(Asigna As Boolean, Optional Todos As Boolean)
'Procedimiento que asigna o elimina areas

    If Asigna Then
        If lstDepartamentos.ListCount > 0 Then
            If Todos Then
                lstDepartamentosSel.Clear
                For vgintCont = 0 To lstDepartamentos.ListCount - 1
                    lstDepartamentosSel.AddItem lstDepartamentos.List(vgintCont), lstDepartamentosSel.ListCount
                    lstDepartamentosSel.ItemData(lstDepartamentosSel.newIndex) = lstDepartamentos.ItemData(vgintCont)
                Next
            Else
                If lstDepartamentos.ListIndex = -1 Then Exit Sub
                If fValida(lstDepartamentos.ItemData(lstDepartamentos.ListIndex)) = False Then Exit Sub
                lstDepartamentosSel.AddItem lstDepartamentos.List(lstDepartamentos.ListIndex), lstDepartamentosSel.ListCount
                lstDepartamentosSel.ItemData(lstDepartamentosSel.newIndex) = lstDepartamentos.ItemData(lstDepartamentos.ListIndex)
                lstDepartamentos.SetFocus
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
        End If
    Else
        If lstDepartamentosSel.ListCount > 0 Then
            If Todos Then
                lstDepartamentosSel.Clear
            Else
                If lstDepartamentosSel.ListIndex = -1 Then Exit Sub
                lstDepartamentosSel.RemoveItem (lstDepartamentosSel.ListIndex)
                If lstDepartamentosSel.ListCount > 0 Then lstDepartamentosSel.SetFocus
            End If
            
            pHabilitaBotonModifica (False)
            cmdBuscar.Enabled = False
        End If
    End If
End Sub

Private Function fValida(Cve As Long) As Boolean
'Valida que el elemento no este asignado anteriormente
    
    fValida = True
    
    With lstDepartamentosSel
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
    
    blnClaveManualCatalogo = fblnClaveManualCatalogo("OTROS CONCEPTOS")
    
     blnLlenagrid = False
    
    vgblnCargarTodosConceptos = True
'    If cgstrModulo = "SI" Then vgblnCargarTodosConceptos = True
      
    pLlenaGridOtroConcepto 2
        
    If vgblnCargarTodosConceptos Then
        vlstrx = "SELECT smiCveConcepto, chrDescripcion FROM PvConceptoFacturacion WHERE bitactivo = 1"
    Else
        vlstrx = "SELECT PvConceptoFacturacion.smiCveConcepto, PvConceptoFacturacion.chrDescripcion FROM "
        vlstrx = vlstrx & " PvConceptoFacturacion INNER JOIN PvConceptoFacturacionEmpresa ON "
        vlstrx = vlstrx & " PvConceptoFacturacion.smiCveConcepto = PvConceptoFacturacionEmpresa.intCveConceptoFactura "
        vlstrx = vlstrx & " WHERE PvConceptoFacturacion.bitactivo = 1 and PvConceptoFacturacion.intTipo = 0 "
        vlstrx = vlstrx & " and PvConceptoFacturacionEmpresa.intCveDepartamento=" + str(vgintNumeroDepartamento)
    End If
    
    Set rscmdConceptoFacturacion = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    
    pLlenarCboRs cboConceptoFacturacion, rscmdConceptoFacturacion, 0, 1
   
    rscmdConceptoFacturacion.Close
    
    pLimpiaFechaAltaOtroConcepto
   
    sstObj.Tab = 0
   
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo NotificaError
    
    If sstObj.Tab = 1 Then
        Cancel = True
        optOrden(0).SetFocus
        
        
        sstObj.Tab = 0
        
        
        TxtCveOConcepto.SetFocus
        
    Else
        If vgstrEstadoManto <> "" Then
            Cancel = True
            If MsgBox(SIHOMsg(9), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                Me.txtDescripcion.SetFocus
                TxtCveOConcepto.SetFocus
                lstDepartamentosSel.Clear
                optOrden(0).SetFocus
                GrdHBusqueda.Clear
                pLimpiaFechaAltaOtroConcepto
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo NotificaError
    
    rsPvOtroConcepto.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Unload"))
End Sub

Private Sub grdHBusqueda_DblClick()
    On Error GoTo NotificaError
   
    If GrdHBusqueda.RowData(1) <> 0 Then
        TxtCveOConcepto.Text = GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)
        TxtCveOConcepto_KeyDown vbKeyReturn, 0
        sstObj.Tab = 0
        'optOrden(0).SetFocus
        
        GrdHBusqueda.Clear
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdHBusqueda_DblClick"))
End Sub


Private Sub grdHBusqueda_KeyPress(KeyAscii As Integer)
If GrdHBusqueda.RowData(1) <> 0 Then
        TxtCveOConcepto.Text = GrdHBusqueda.TextMatrix(GrdHBusqueda.Row, 1)
        TxtCveOConcepto_KeyDown vbKeyReturn, 0
        sstObj.Tab = 0
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
    
    With rsPvOtroConcepto
        TxtCveOConcepto.Text = !intCveConcepto
        txtDescripcion.Text = Trim(!chrDescripcion)
        cboConceptoFacturacion.ListIndex = fintLocalizaCbo(cboConceptoFacturacion, !SMICONCEPTOFACT)
        '*****Fecha de alta de otro concepto *****************************************
        If Not !DTMFECHAALTAOTROCONCEPTO = vbNullString Then
            txtFechaAltaOtroConcepto.Text = Format(!DTMFECHAALTAOTROCONCEPTO, "dd/mmm/yyyy")
            txtFechaAltaOtroConcepto.Visible = True
            lblFechaAltaOtroConcepto.Visible = True
        Else
            lblFechaAltaOtroConcepto.Visible = False
            txtFechaAltaOtroConcepto.Visible = False
        End If
        '*******************************************************************
        If !BITESTATUS Or !BITESTATUS = 1 Then
            ChkStatus.Value = 1
        Else
            ChkStatus.Value = 0
        End If
    End With
    
    lstDepartamentosSel.Clear
    
    'Se cargan los departamentos asignados al concepto
    vgstrSentencia = "Select PvOtroConceptoDepto.smiCveDepartamento Cve, Rtrim(NoDepartamento.vchDescripcion) Nombre from PvOtroConceptoDepto " & _
                     "Inner Join NoDepartamento On NoDepartamento.smiCveDepartamento = PvOtroConceptoDepto.smiCveDepartamento " & _
                     "Where PvOtroConceptoDepto.intCveOtroConcepto = " & rsPvOtroConcepto!intCveConcepto & _
                     " Order By Nombre"
                     
    Set rsDepartamentos = frsRegresaRs(vgstrSentencia, adLockReadOnly, adOpenForwardOnly)
    
    If rsDepartamentos.RecordCount > 0 Then
        With lstDepartamentosSel
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

    pLlenaGridOtroConcepto IIf(optOrden(0).Value, 2, 1)
    
    GrdHBusqueda.SetFocus
    
End Sub




Private Sub TxtCveOConcepto_GotFocus()
    On Error GoTo NotificaError
    If Not blnEnfocando Then
        pNuevoRegistro
        If TxtCveOConcepto.Enabled And TxtCveOConcepto.Visible Then
            pEnfocaTextBox TxtCveOConcepto
        End If
    End If
    blnEnfocando = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveOConcepto_GotFocus"))
End Sub

Private Sub TxtCveOConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsBusca As ADODB.Recordset
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        'Buscar criterio
        txtDescripcion.Enabled = True
        If lngSig = CLng(TxtCveOConcepto.Text) Then
            vgstrEstadoManto = "A" 'Alta
            Call pEnfocaTextBox(txtDescripcion)
            pHabilitaComponentesCaptura True
            pHabilitaBotonModifica (False)
            cmdGrabarRegistro.Enabled = True
            cmdBuscar.Enabled = False
            ChkStatus.Value = 1
            ChkStatus.Enabled = False
            
        Else
            If fintLocalizaPkRs(rsPvOtroConcepto, 0, TxtCveOConcepto.Text) > 0 Then
                pModificaRegistro
                vlblnActivaraCaptura = False
                vgstrEstadoManto = "M" 'Modificacion
                pHabilitaComponentesCaptura True
                pHabilitaBotonModifica (True)
                cmdBuscar.Enabled = True
                ChkStatus.Enabled = True
                Call pEnfocaTextBox(txtDescripcion)
            Else
                If blnClaveManualCatalogo Then
                    Set rsBusca = frsRegresaRs("select * from PvOtroConcepto where intCveConcepto = " & Me.TxtCveOConcepto.Text, adLockReadOnly, adOpenForwardOnly)
                    If rsBusca.EOF Then
                        vgstrEstadoManto = "A" 'Alta
                        Call pEnfocaTextBox(txtDescripcion)
                        pHabilitaComponentesCaptura True
                        pHabilitaBotonModifica (False)
                        cmdGrabarRegistro.Enabled = True
                        cmdBuscar.Enabled = False
                        ChkStatus.Value = 1
                        ChkStatus.Enabled = False
                    Else
                        MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                        Call pEnfocaTextBox(TxtCveOConcepto)
                    End If
                    rsBusca.Close
                Else
                    MsgBox SIHOMsg(12), vbExclamation, "Mensaje"
                    Call pEnfocaTextBox(TxtCveOConcepto)
                End If
            End If
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveOConcepto_KeyDown"))
End Sub

Private Sub TxtCveOConcepto_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":TxtCveOConcepto_KeyPress"))
End Sub

Private Sub txtDescripcion_GotFocus()
    If vlblnActivaraCaptura Then
        pHabilitaBotonModifica (False)
        cmdBuscar.Enabled = False
        Call pEnfocaTextBox(txtDescripcion)
    End If
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
    
    Call pPosicionaRegRs(rsPvOtroConcepto, "A")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAnteriorRegistro_Click"))
End Sub

Private Sub cmdPrimerRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvOtroConcepto, "I")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrimerRegistro_Click"))
End Sub

Private Sub cmdSiguienteRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvOtroConcepto, "S")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSiguienteRegistro_Click"))
End Sub

Private Sub cmdUltimoRegistro_Click()
    On Error GoTo NotificaError
    
    Call pPosicionaRegRs(rsPvOtroConcepto, "U")
    pModificaRegistro

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdUltimoRegistro_Click"))
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        pEnfocaCbo cboConceptoFacturacion
    End If

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

Private Sub lstDepartamentos_DblClick()
    pAsigna True
End Sub

Private Sub lstDepartamentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna True
End Sub

Private Sub lstDepartamentosSel_DblClick()
    pAsigna False
End Sub

Private Sub lstDepartamentosSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pAsigna False
End Sub

Private Sub pHabilitaComponentesCaptura(vlblnHabilita As Boolean)
    txtDescripcion.Enabled = vlblnHabilita
    cboConceptoFacturacion.Enabled = vlblnHabilita
    ChkStatus.Enabled = vlblnHabilita
    lstDepartamentos.Enabled = vlblnHabilita
    cmdAsignaTodo.Enabled = vlblnHabilita
    cmdAsignaUno.Enabled = vlblnHabilita
    cmdEliminaUno.Enabled = vlblnHabilita
    cmdEliminaTodo.Enabled = vlblnHabilita
    lstDepartamentosSel.Enabled = vlblnHabilita
End Sub

Private Sub pLlenaGridOtroConcepto(vlintOrden As Integer)

    GrdHBusqueda.Clear
    
    If vlintOrden = 1 Then
      vlstrOrden = " PvOtroConcepto.chrDescripcion"
    Else
      vlstrOrden = " PvOtroConcepto.intCveConcepto"
    End If
    
    If vgblnCargarTodosConceptos Then
        vlstrx = "SELECT PvOtroConcepto.*, " & _
        "rtrim(PvConceptoFacturacion.chrDescripcion) chrDescripcionF " & _
        "FROM PvConceptoFacturacion INNER JOIN " & _
        "PvOtroConcepto ON " & _
        "PvConceptoFacturacion.smiCveConcepto = PvOtroConcepto.smiConceptoFact " & _
        "where bitActivo = 1 " & _
        "order by " & vlstrOrden
    Else
        vlstrx = "SELECT PvOtroConcepto.*, " & _
        "rtrim(PvConceptoFacturacion.chrDescripcion) chrDescripcionF " & _
        "FROM PvConceptoFacturacion INNER JOIN " & _
        "PvOtroConcepto ON " & _
        "PvConceptoFacturacion.smiCveConcepto = PvOtroConcepto.smiConceptoFact " & _
        "where bitActivo = 1 and PvConceptoFacturacion.intTipo = 0 and pvOtroConcepto.smiDepartamento = " & Trim(str(vgintNumeroDepartamento)) & _
        " order by " & vlstrOrden
    End If
    
    Set rsPvOtroConcepto = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
    If rsPvOtroConcepto.RecordCount > 0 Then
    End If
    If blnLlenagrid Then
        pLlenarMshFGrdRs GrdHBusqueda, rsPvOtroConcepto, 0
        pConfiguraGrid
    End If
    rsPvOtroConcepto.Requery
End Sub



Private Sub pLimpiaFechaAltaOtroConcepto()
    lblFechaAltaOtroConcepto.Visible = False
    txtFechaAltaOtroConcepto.Visible = False
    txtFechaAltaOtroConcepto.Text = vbNullString
End Sub
