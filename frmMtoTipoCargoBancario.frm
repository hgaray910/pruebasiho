VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMtoTipoCargoBancario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de cargos bancarios"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   568
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTObj 
      Height          =   3540
      Left            =   -45
      TabIndex        =   12
      Top             =   -450
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   6244
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   661
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Mantenimiento"
      TabPicture(0)   =   "frmMtoTipoCargoBancario.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Departamentos que lo utilizan"
      TabPicture(1)   =   "frmMtoTipoCargoBancario.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdHBusqueda"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1880
         Left            =   190
         TabIndex        =   8
         Top             =   600
         Width           =   8190
         Begin VB.TextBox txtCveComision 
            Height          =   315
            Left            =   1635
            MaxLength       =   8
            TabIndex        =   0
            ToolTipText     =   "Clave"
            Top             =   330
            Width           =   1185
         End
         Begin VB.TextBox txtDescripcion 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1635
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   "Descripción "
            Top             =   720
            Width           =   6135
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            BeginProperty DataFormat 
               Type            =   5
               Format          =   ""
               HaveTrueFalseNull=   1
               TrueValue       =   "True"
               FalseValue      =   "False"
               NullValue       =   ""
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   7
            EndProperty
            Enabled         =   0   'False
            Height          =   255
            Left            =   1635
            TabIndex        =   2
            ToolTipText     =   "Estado"
            Top             =   1140
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Left            =   210
            TabIndex        =   16
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblNombre 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   210
            TabIndex        =   14
            Top             =   750
            Width           =   840
         End
         Begin VB.Label lblClave 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   210
            TabIndex        =   13
            Top             =   390
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
         Height          =   720
         Left            =   2520
         TabIndex        =   11
         Top             =   2600
         Width           =   3645
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   480
            Left            =   3060
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Borrar el registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdGrabarRegistro 
            Enabled         =   0   'False
            Height          =   480
            Left            =   2550
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":01DA
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Guardar el registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdUltimoRegistro 
            Height          =   480
            Left            =   2055
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":051C
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ultimo registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSiguienteRegistro 
            Height          =   480
            Left            =   1560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":068E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Siguiente registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBuscar 
            Height          =   480
            Left            =   1065
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":0800
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Búsqueda"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdAnteriorRegistro 
            Height          =   480
            Left            =   570
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":0972
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Anterior registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdPrimerRegistro 
            Height          =   480
            Left            =   75
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMtoTipoCargoBancario.frx":0AE4
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Primer registro"
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdHBusqueda 
         CausesValidation=   0   'False
         DragIcon        =   "frmMtoTipoCargoBancario.frx":0C56
         Height          =   2900
         Left            =   -74895
         TabIndex        =   15
         ToolTipText     =   "Doble click para seleccionar un tipo de cargo bancario"
         Top             =   525
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   5106
         _Version        =   393216
         ForeColor       =   0
         Rows            =   16
         Cols            =   5
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLineWidthBand=   1
         _Band(0).TextStyleBand=   0
      End
   End
End
Attribute VB_Name = "frmMtoTipoCargoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Permite forzar la declaración de las variables
Dim vgblnNuevoRegistro As Boolean
Public vllngNumeroOpcion As Long
Dim rsTipoCargoBancario As New ADODB.Recordset
Dim vlstrsql As String
Dim vlblnConsulta As Boolean
Private Sub chkActivo_GotFocus()
    If vlblnConsulta Then pHabilitaBotonModifica (True)
End Sub

Private Sub pHabilita(intTop As Integer, intBack As Integer, intlocate As Integer, intNext As Integer, intEnd As Integer, intSave As Integer, intDelete As Integer)
    cmdPrimerRegistro.Enabled = intTop = 1
    cmdAnteriorRegistro.Enabled = intBack = 1
    cmdBuscar.Enabled = intlocate = 1
    cmdSiguienteRegistro.Enabled = intNext = 1
    cmdUltimoRegistro.Enabled = intEnd = 1
    cmdGrabarRegistro.Enabled = intSave = 1
    cmdDelete.Enabled = intDelete = 1
End Sub

Private Sub chkPredeterminado_GotFocus()
    If vlblnConsulta Then pHabilitaBotonModifica (True)
End Sub

Private Sub cmdDelete_Click()
    Dim vlstrSentencia As String
    Dim rsComision As New ADODB.Recordset
    Dim vllngPersonaGraba As Long
    
    If fblnRevisaPermiso(vglngNumeroLogin, 3062, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3062, "C", True) _
        Or fblnRevisaPermiso(vglngNumeroLogin, 3063, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3063, "C", True) Then
              
        vlstrSentencia = "Select * from pvformapagotipocargocomision where intcvetipocargo = " & txtCveComision.Text
        Set rsComision = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)

        If rsComision.RecordCount > 0 Then
            'No se puede eliminar la información, ya ha sido utilizada.
            MsgBox SIHOMsg(771), vbOKOnly + vbCritical, "Mensaje"
            pNuevoRegistro True
            pEnfocaTextBox txtCveComision
        Else
            If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                '-----------------------'
                '   Persona que graba   '
                '-----------------------'
                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                If vllngPersonaGraba <> 0 Then
                    EntornoSIHO.ConeccionSIHO.BeginTrans
                                  
                    rsTipoCargoBancario.Delete
                    rsTipoCargoBancario.Requery
                
                    Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "TIPOCARGOBANCARIO", txtCveComision.Text)
                    EntornoSIHO.ConeccionSIHO.CommitTrans
                    
                    pNuevoRegistro True
                    pEnfocaTextBox txtCveComision
                End If
            End If
        End If
    Else
        MsgBox SIHOMsg(635), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub

Private Sub cmdGrabarRegistro_Click()
Dim vlintContador As Integer
Dim vlintclavecomision As Integer
Dim vlintSeqFil As Integer
Dim vllngPersonaGraba As Long
    
    ' 603   CC
    ' 2367  SI (CC)
    ' 348   PV          3060
    ' 1120  SI (PV)     3061
'    If fblnRevisaPermiso(vglngNumeroLogin, 603, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 603, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 2367, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 2367, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3060, "C", True) _
'        Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3061, "C", True) Then
    If fblnRevisaPermiso(vglngNumeroLogin, 3062, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3062, "C", True) _
        Or fblnRevisaPermiso(vglngNumeroLogin, 3063, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, 3063, "C", True) Then
                
        '**********************************************************
        ' Procedimiento para grabar una alta o modificación
        '**********************************************************

        If RTrim(txtDescripcion.Text) = "" Then
            MsgBox SIHOMsg(2) + Chr(13) + txtDescripcion.ToolTipText, vbExclamation, "Mensaje"
            txtDescripcion.SetFocus
            Exit Sub
        End If
        
'        If chkPredeterminado.Value = 1 Then
'            If fblnExisteTipoCargo Then
'                'Ya existe registrado un concepto para generar pagos al cancelar facturas.
'                MsgBox SIHOMsg(1336), vbOKOnly + vbInformation, "Mensaje"
'                chkPredeterminado.SetFocus
'                Exit Sub
'            End If
'        End If
    
        '-----------------------'
        '   Persona que graba   '
        '-----------------------'
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            '--------------------------------------
            ' Grabar la comision
            '--------------------------------------
            EntornoSIHO.ConeccionSIHO.BeginTrans
            
            If Not vlblnConsulta Then
                rsTipoCargoBancario.AddNew
                vlintclavecomision = 0
            Else
                vlintclavecomision = CDbl(txtCveComision.Text)
            End If
            
            rsTipoCargoBancario!chrDescripcion = Trim(txtDescripcion.Text)
            'rsTipoCargoBancario!bitpredeterminado = chkPredeterminado.Value
            rsTipoCargoBancario!bitactivo = chkActivo.Value
            rsTipoCargoBancario.Update
            rsTipoCargoBancario.Requery
            
            If vlintclavecomision = 0 Then
                txtCveComision.Text = CStr(flngObtieneIdentity("sec_pvtipocargobancario", 0))
                vlintclavecomision = CDbl(txtCveComision.Text)
            End If
                       
            If Not vlblnConsulta Then
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "TIPOCARGOBANCARIO", txtCveComision.Text)
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "TIPOCARGOBANCARIO", txtCveComision.Text)
            End If
            
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            Call pNuevoRegistro(True)
            txtCveComision.SetFocus
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If
End Sub


Private Function fblnExisteTipoCargo() As Boolean
    On Error GoTo NotificaError
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    fblnExisteTipoCargo = False
    strSQL = "select count(*) tipoCargobitPredeterminado from pvTipoCargoBancario where bitPredeterminado = 1"
    Set rs = frsRegresaRs(strSQL, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount > 0 Then
        If rs!tipoCargobitPredeterminado > 0 Then
            fblnExisteTipoCargo = True
        End If
    End If
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteTipoCargo"))
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyEscape
            KeyCode = 7
            Unload Me
        Case vbKeyReturn
            SendKeys vbTab
    End Select
    
End Sub

Private Sub Form_Load()
    Dim vlstrSentencia As String
    
    vlstrSentencia = "select * from PvTipoCargoBancario"
    Set rsTipoCargoBancario = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    Me.Icon = frmMenuPrincipal.Icon
    pNuevoRegistro True
    vgblnNuevoRegistro = True
   
    SSTObj.Tab = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If SSTObj.Tab <> 0 Then
       
        SSTObj.Tab = 0
        If txtDescripcion.Enabled Then
            txtDescripcion.SetFocus
        Else
            txtCveComision.SetFocus
        End If
        Cancel = True
    
    Else
        If Not vgblnNuevoRegistro Then
            If MsgBox(SIHOMsg(9), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                Call pNuevoRegistro(True)
                pEnfocaTextBox txtCveComision
            End If
            Cancel = True
        Else
            rsTipoCargoBancario.Close
        End If
    End If
End Sub


Private Sub pNuevoRegistro(vlblnNuevo As Boolean)
    If SSTObj.Tab = 1 Then Exit Sub
    If vlblnNuevo Then txtCveComision.Text = fintSigNumRs(rsTipoCargoBancario, 0)
    txtDescripcion.Text = ""
    txtDescripcion.Enabled = False
    chkActivo.Enabled = False
'    chkPredeterminado.Enabled = False
       
    If rsTipoCargoBancario.RecordCount > 0 Then
        pHabilitaBotonModifica (True)
        cmdGrabarRegistro.Enabled = False
        cmdDelete.Enabled = False
    Else
        pHabilitaBotonModifica (False)
    End If
    chkActivo.Value = 1
'    chkPredeterminado.Value = 0
    vgblnNuevoRegistro = True
    
    vlblnConsulta = False
End Sub
Public Sub pConfFGrid(ObjGrid As MSHFlexGrid, vlstrFormatoTitulo As String)
    'Configura el MSHFlexGrid

    ObjGrid.FormatString = vlstrFormatoTitulo 'Encabezados de columnas
    

     With ObjGrid
       .ColWidth(0) = 150
       .ColWidth(1) = 0
       .ColWidth(2) = 3000
       .ColWidth(3) = 0
       .ColWidth(4) = 1700
       .ColWidth(5) = 4500
       .ColWidth(6) = 0
       .ScrollBars = flexScrollBarHorizontal
    End With


End Sub
Private Sub pLlenaGrid(vlintOrden As Integer)
    Dim vlstrSentencia As String
    Dim PvComisionBancaria As New ADODB.Recordset
    Dim pvTipoCargoBancario As New ADODB.Recordset
    Dim vlintContador As Integer
    grdHBusqueda.Clear
    
    vlstrSentencia = "SELECT pvTipoCargoBancario.intCveTipoCargo, " & _
    "pvTipoCargoBancario.chrDescripcion, " & _
    "pvTipoCargoBancario.bitActivo " & _
    "FROM pvTipoCargoBancario " & _
    "ORDER BY " & vlintOrden
'    "pvTipoCargoBancario.bitPredeterminado, "
    
    Set pvTipoCargoBancario = frsRegresaRs(vlstrSentencia, adLockReadOnly)
    If pvTipoCargoBancario.RecordCount > 0 Then
        Call pLlenarMshFGrdRs(grdHBusqueda, pvTipoCargoBancario)
        pConfiguraGrid
        With grdHBusqueda
            For vlintContador = 1 To .Rows - 1
'                .TextMatrix(vlintContador, 3) = IIf(.TextMatrix(vlintContador, 3) = 1, "X", "")
                .TextMatrix(vlintContador, 3) = IIf(.TextMatrix(vlintContador, 3) = 1, "Activo", "Inactivo")
            Next
        End With
    Else
        SSTObj.Tab = 0
        cmdBuscar.SetFocus
    End If
    pvTipoCargoBancario.Close
End Sub

Private Sub pConfiguraGrid()
    With grdHBusqueda
        .FormatString = "|Clave|Descripción|Estado"
        .ColWidth(0) = 150 'Fix
        .ColWidth(1) = 700 'Clave
        .ColWidth(2) = 5600 'Descripcion
        .ColAlignment(2) = flexAlignLeftCenter
        '.ColWidth(3) = 1250  'Predeterminado
        '.ColAlignment(3) = flexAlignCenterCenter
        .ColWidth(3) = 1500  'Estatus
        .ScrollBars = flexScrollBarVertical
    End With
End Sub


Private Sub grdHBusqueda_Click()
    With grdHBusqueda
        If .MouseRow = 0 And .MouseCol <> 0 Then
            'pLlenaGrid IIf(.MouseCol = 1, 0, IIf(.MouseCol = 2, 1, 2))
            pLlenaGrid (.MouseCol)
        End If
    End With
End Sub

Private Sub grdHBusqueda_DblClick()
    
    If fintLocalizaPkRs(rsTipoCargoBancario, 0, grdHBusqueda.TextMatrix(grdHBusqueda.Row, 1)) > 0 Then
        pModificaRegistro
        Call pEnfocaTextBox(txtDescripcion)
        SSTObj.Tab = 0
        txtCveComision_KeyDown 13, 0
    Else
        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
        Call pEnfocaTextBox(txtCveComision)
    End If
End Sub

Private Sub grdHBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdHBusqueda_DblClick
    End If
End Sub

Private Sub SSTObj_Click(PreviousTab As Integer)
    If SSTObj.Tab = 1 Then
        pLlenaGrid (2)
        grdHBusqueda.Enabled = True
        grdHBusqueda.SetFocus
    End If
End Sub



Private Sub txtCveComision_GotFocus()
    pNuevoRegistro (True)
    pSelTextBox txtCveComision

End Sub

Private Sub txtCveComision_KeyDown(KeyCode As Integer, Shift As Integer)
'-------------------------------------------------------------------------------------------
'Validación para diferenciar cuando es una alta de un registro o cuando se va a consultar o
'modificar uno que ya existe
'-------------------------------------------------------------------------------------------
    Dim vlintNumero As Integer
    Select Case KeyCode
        Case vbKeyReturn
                If fintSigNumRs(rsTipoCargoBancario, 0) = CLng(txtCveComision.Text) Then
                    txtDescripcion.Enabled = True
                    chkActivo.Enabled = True
'                    chkPredeterminado.Enabled = True
                    vgblnNuevoRegistro = False
                    
                    chkActivo.Value = 1
                    chkActivo.Enabled = False
                    
                    pHabilitaBotonModifica False
                    cmdGrabarRegistro.Enabled = True
                    SSTObj.TabEnabled(1) = False
                Else
                    If fintLocalizaPkRs(rsTipoCargoBancario, 0, txtCveComision.Text) > 0 Then
                        pModificaRegistro
                        txtDescripcion.Enabled = True
                        chkActivo.Enabled = True
                        pHabilitaBotonModifica (True)
                        chkActivo.Enabled = True
'                        chkPredeterminado.Enabled = True
                        txtDescripcion.SetFocus
                    Else
                        rsTipoCargoBancario.MoveLast
                        Call MsgBox(SIHOMsg(12), vbExclamation, "Mensaje")
                        Call pEnfocaTextBox(txtCveComision)
                        pNuevoRegistro (True)
                        txtCveComision_GotFocus
                    End If
                End If
    End Select
End Sub

Private Sub pModificaRegistro()
    Dim rscuentasempresas As New ADODB.Recordset
    Dim vlstrSentencia As String
    Dim vlintContador As Integer
    '-------------------------------------------------------------------------------------------
    ' Permite realizar la modificación de la descripción de un registro
    '-------------------------------------------------------------------------------------------
    vgblnNuevoRegistro = False
    
    'txtDescripcion.Enabled = False
    'chkActivo.Enabled = False
    'chkPredeterminado.Enabled = False
    
    '---------------------------------------
    ' Carga las comisiones
    '---------------------------------------
    txtCveComision.Text = rsTipoCargoBancario!intcvetipocargo
    txtDescripcion.Text = rsTipoCargoBancario!chrDescripcion
'    chkPredeterminado.Value = IIf(rsTipoCargoBancario!bitpredeterminado Or rsTipoCargoBancario!bitpredeterminado = 1, 1, 0)
    chkActivo.Value = IIf(rsTipoCargoBancario!bitactivo Or rsTipoCargoBancario!bitactivo = 1, 1, 0)
    SSTObj.TabEnabled(1) = True
    vlblnConsulta = True

End Sub

Private Sub pHabilitaBotonModifica(vlblnHabilita As Boolean)
'-------------------------------------------------------------------------------------------
' Habilitar o deshabilitar la botonera completa cuando se trata de una modficiación
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    cmdPrimerRegistro.Enabled = vlblnHabilita
    cmdAnteriorRegistro.Enabled = vlblnHabilita
    cmdBuscar.Enabled = vlblnHabilita
    SSTObj.TabEnabled(1) = vlblnHabilita
    cmdSiguienteRegistro.Enabled = vlblnHabilita
    cmdUltimoRegistro.Enabled = vlblnHabilita
    cmdDelete.Enabled = vlblnHabilita
    cmdGrabarRegistro.Enabled = vlblnHabilita
  
Exit Sub
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

Private Sub txtCveComision_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    End If
End Sub

Private Sub txtDescripcion_GotFocus()
  pSelTextBox txtDescripcion
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub cmdAnteriorRegistro_Click()
    Call pPosicionaRegRs(rsTipoCargoBancario, "A")
    pModificaRegistro
End Sub

Private Sub cmdPrimerRegistro_Click()
    Call pPosicionaRegRs(rsTipoCargoBancario, "I")
    pModificaRegistro
End Sub

Private Sub cmdSiguienteRegistro_Click()
    Call pPosicionaRegRs(rsTipoCargoBancario, "S")
    pModificaRegistro
End Sub

Private Sub cmdUltimoRegistro_Click()
    Call pPosicionaRegRs(rsTipoCargoBancario, "U")
    pModificaRegistro
End Sub

Private Sub chkActivo_Click()
    If vgblnNuevoRegistro Then
        chkActivo.Value = 1
    End If
End Sub

Private Sub cmdBuscar_Click()
    SSTObj.Tab = 1
End Sub


