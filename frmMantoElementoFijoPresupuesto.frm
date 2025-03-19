VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMantoElementoFijoPresupuesto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elementos fijos para presupuestos"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabConcepto 
      Height          =   4365
      Left            =   -60
      TabIndex        =   10
      Top             =   -555
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   7699
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmMantoElementoFijoPresupuesto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmMantoElementoFijoPresupuesto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdConceptos"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   765
         Left            =   2377
         TabIndex        =   14
         Top             =   2430
         Width           =   3750
         Begin VB.CommandButton cmdTop 
            Height          =   495
            Left            =   90
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Primer registro"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdBack 
            Height          =   495
            Left            =   600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":015A
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Anterior registro"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdLocate 
            Height          =   495
            Left            =   1110
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":02CC
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Búsqueda"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdNext 
            Height          =   495
            Left            =   1620
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":043E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Siguiente registro"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdEnd 
            Height          =   495
            Left            =   2130
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":05B0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Ultimo registro"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton cmdSave 
            Height          =   495
            Left            =   2625
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":0722
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Grabar"
            Top             =   180
            Width           =   495
         End
         Begin VB.CommandButton cmdDelete 
            Height          =   495
            Left            =   3135
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMantoElementoFijoPresupuesto.frx":0A64
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Borrar"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   180
         TabIndex        =   11
         Top             =   570
         Width           =   8280
         Begin VB.TextBox txtIVA 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            MaxLength       =   2
            TabIndex        =   15
            ToolTipText     =   "IVA"
            Top             =   1020
            Width           =   945
         End
         Begin VB.CheckBox chkActivo 
            Caption         =   "Activo"
            Height          =   300
            Left            =   1485
            TabIndex        =   2
            Top             =   1380
            Width           =   990
         End
         Begin VB.TextBox txtClave 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1485
            MaxLength       =   9999
            TabIndex        =   0
            ToolTipText     =   "Clave "
            Top             =   315
            Width           =   945
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   315
            Left            =   1485
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "Descripción "
            Top             =   675
            Width           =   6480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "IVA"
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   2505
            TabIndex        =   16
            Top             =   1095
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Clave"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   375
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   735
            Width           =   840
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptos 
         Height          =   2565
         Left            =   -74880
         TabIndex        =   18
         Top             =   600
         Width           =   8370
         _ExtentX        =   14764
         _ExtentY        =   4524
         _Version        =   393216
         GridLines       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmMantoElementoFijoPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------
' Programa para dar mantenimiento a los conceptos de pago (PvPresupuestoElementoFijo)
' Fecha de programación: Jueves 26 de Septiembre de 2002
'---------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'---------------------------------------------------------------------------
Dim rsPvPresupuestoElementoFijo As New ADODB.Recordset

Dim vlstrX As String
Dim vlstrSentenciaConsulta As String

Dim vlblnConsulta As Boolean


Private Sub chkActivo_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_GotFocus"))
End Sub

Private Sub chkActivo_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkActivo_KeyPress"))
End Sub



Private Sub cmdBack_Click()
    On Error GoTo NotificaError
    
    rsPvPresupuestoElementoFijo.MovePrevious
    If rsPvPresupuestoElementoFijo.BOF Then
        rsPvPresupuestoElementoFijo.MoveNext
    End If
    pMuestraConcepto rsPvPresupuestoElementoFijo!intCveElemento
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBack_Click"))
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo NotificaError

    If Not fblnElementoAsignado() Then
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vglngNumeroLogin, "ELEMENTO FIJO", txtClave.Text)
        rsPvPresupuestoElementoFijo.Delete
        rsPvPresupuestoElementoFijo.Update
        txtClave.SetFocus
    Else
        '!No se pueden borrar los datos!
        MsgBox SIHOMsg(257), vbOKOnly + vbExclamation, "Mensaje"
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdDelete_Click"))
End Sub

Private Function fblnElementoAsignado() As Boolean
    On Error GoTo NotificaError
    
    fblnElementoAsignado = False
    
    vlstrX = "" & _
    "select " & _
        "count(*) " & _
    "From " & _
        "PvDetallePresupuesto " & _
    "Where " & _
        "chrTipoCargo='EF' and " & _
        "intCveCargo = " & Str(rsPvPresupuestoElementoFijo!intCveElemento)
    If frsRegresaRs(vlstrX).Fields(0) <> 0 Then
        fblnElementoAsignado = True
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnElementoAsignado"))
End Function


Private Sub cmdEnd_Click()
    On Error GoTo NotificaError
    
    rsPvPresupuestoElementoFijo.MoveLast
    pMuestraConcepto rsPvPresupuestoElementoFijo!intCveElemento
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdEnd_Click"))
End Sub

Private Sub cmdLocate_Click()
    On Error GoTo NotificaError
    
    sstabConcepto.Tab = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdLocate_Click"))
End Sub

Private Sub cmdNext_Click()
    On Error GoTo NotificaError
    If Not rsPvPresupuestoElementoFijo.EOF Then
        rsPvPresupuestoElementoFijo.MoveNext
    End If
    If rsPvPresupuestoElementoFijo.EOF Then
        rsPvPresupuestoElementoFijo.MovePrevious
    End If
    pMuestraConcepto rsPvPresupuestoElementoFijo!intCveElemento
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdNext_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    If fblnDatosValidos() Then
        With rsPvPresupuestoElementoFijo
            If Not vlblnConsulta Then
                .AddNew
            End If
            !vchDescripcion = Trim(txtDescripcion.Text)
            !smyIVA = Val(txtIva.Text)
            !bitactivo = chkActivo.Value
            .Update
            If Not vlblnConsulta Then
                txtClave.Text = flngObtieneIdentity("SEC_PVPRESUPUESTOELEMENTOFIJO", rsPvPresupuestoElementoFijo!intCveElemento)
                Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngNumeroLogin, "ELEMENTO FIJO", txtClave.Text)
            Else
                Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vglngNumeroLogin, "ELEMENTO FIJO", txtClave.Text)
            End If
        End With
        rsPvPresupuestoElementoFijo.Requery
        txtClave.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Trim(txtDescripcion.Text) = "" Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtDescripcion.SetFocus
    End If
    If fblnDatosValidos And Not vlblnConsulta Then
        If fblnExisteConcepto(txtDescripcion.Text) Then
            fblnDatosValidos = False
            'Existe información con el mismo contenido
            MsgBox SIHOMsg(19), vbOKOnly + vbInformation, "Mensaje"
            txtDescripcion.SetFocus
        End If
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Function fblnExisteConcepto(vlstrxConcepto As String) As Boolean
    On Error GoTo NotificaError
    
    fblnExisteConcepto = False
    If rsPvPresupuestoElementoFijo.RecordCount <> 0 Then
        rsPvPresupuestoElementoFijo.MoveFirst
        Do While Not rsPvPresupuestoElementoFijo.EOF
            If Trim(rsPvPresupuestoElementoFijo!vchDescripcion) = vlstrxConcepto Then
                fblnExisteConcepto = True
            End If
            rsPvPresupuestoElementoFijo.MoveNext
        Loop
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnExisteConcepto"))
End Function

Private Sub cmdTop_Click()
    On Error GoTo NotificaError
    
    rsPvPresupuestoElementoFijo.MoveFirst
    pMuestraConcepto rsPvPresupuestoElementoFijo!intCveElemento
    pHabilita 1, 1, 1, 1, 1, 0, 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTop_Click"))
End Sub


Private Sub Form_Activate()

    vgstrNombreForm = Me.Name
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        If sstabConcepto.Tab = 0 Then
            If cmdSave.Enabled Then
                ' ¿Desea abandonar la operación?
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
                    txtClave.SetFocus
                End If
            Else
                Unload Me
            End If
        Else
            sstabConcepto.Tab = 0
            txtClave.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlstrX = "select * from PvPresupuestoElementoFijo order by intCveElemento"
    Set rsPvPresupuestoElementoFijo = frsRegresaRs(vlstrX, adLockOptimistic, adOpenDynamic)
    

    vlstrSentenciaConsulta = "select intCveElemento, vchDescripcion, case bitActivo when 1 then 'Activo' when 0 then 'Inactivo' end  as estatus FROM PvPresupuestoElementoFijo"
    sstabConcepto.Tab = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub grdConceptos_DblClick()
    On Error GoTo NotificaError
    
    pMuestraConcepto grdConceptos.RowData(grdConceptos.Row)
    pHabilita 1, 1, 1, 1, 1, 0, 1
    sstabConcepto.Tab = 0
    cmdLocate.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConceptos_DblClick"))
End Sub

Private Sub sstabConcepto_Click(PreviousTab As Integer)
    On Error GoTo NotificaError
    
    If sstabConcepto.Tab = 1 Then
        grdConceptos.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":sstabConcepto_Click"))
End Sub

Private Sub txtClave_GotFocus()
    On Error GoTo NotificaError
    If rsPvPresupuestoElementoFijo.RecordCount > 0 Then
        pHabilita 1, 1, 1, 1, 1, 0, 0
    Else
        pHabilita 0, 0, 0, 0, 0, 0, 0
    End If
    pLimpia
    pSelTextBox txtClave

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_GotFocus"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    If rsPvPresupuestoElementoFijo.RecordCount = 0 Then
        txtClave.Text = "1"
    Else
        txtClave.Text = frsRegresaRs("select max(intCveElemento)+1 from PvPresupuestoElementoFijo").Fields(0)
    End If
    txtDescripcion.Text = ""
    txtIva.Text = ""
    chkActivo.Value = 1

    vlblnConsulta = False
    
    If rsPvPresupuestoElementoFijo.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdConceptos, frsRegresaRs(vlstrSentenciaConsulta, adLockOptimistic, adOpenDynamic), 0
        With grdConceptos
            .FormatString = "|Clave|Descripción|Estatus"
            .ColWidth(0) = 100
            .ColWidth(1) = 1000
            .ColWidth(2) = 5000
            .ColWidth(3) = 2000
        End With
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpia"))
End Sub

Private Sub pHabilita(vlb1 As Integer, vlb2 As Integer, vlb3 As Integer, vlb4 As Integer, vlb5 As Integer, vlb6 As Integer, vlb7 As Integer)
    On Error GoTo NotificaError
    
    cmdTop.Enabled = vlb1 = 1
    cmdBack.Enabled = vlb2 = 1
    cmdLocate.Enabled = vlb3 = 1
    cmdNext.Enabled = vlb4 = 1
    cmdEnd.Enabled = vlb5 = 1
    cmdSave.Enabled = vlb6 = 1
    cmdDelete.Enabled = vlb7 = 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pHabilita"))
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtDescripcion.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_KeyPress"))
End Sub

Private Sub txtClave_LostFocus()
    On Error GoTo NotificaError
    
    If Trim(txtClave.Text) = "" Then
        pLimpia
    Else
        pMuestraConcepto Val(txtClave.Text)
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_LostFocus"))
End Sub

Private Sub pMuestraConcepto(vllngxNumero As Long)
    On Error GoTo NotificaError
    
    If fintLocalizaPkRs(rsPvPresupuestoElementoFijo, 0, Str(vllngxNumero)) <> 0 Then
        txtClave.Text = rsPvPresupuestoElementoFijo!intCveElemento
        txtDescripcion.Text = rsPvPresupuestoElementoFijo!vchDescripcion
        txtIva.Text = rsPvPresupuestoElementoFijo!smyIVA
        chkActivo.Value = IIf(rsPvPresupuestoElementoFijo!bitactivo Or rsPvPresupuestoElementoFijo!bitactivo = 1, 1, 0)
        vlblnConsulta = True
    Else
        pLimpia
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pMuestraConcepto"))
End Sub


Private Sub txtDescripcion_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtDescripcion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_GotFocus"))
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        txtIva.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDescripcion_KeyPress"))
End Sub

Private Sub txtIVA_GotFocus()
    On Error GoTo NotificaError
    
    pHabilita 0, 0, 0, 0, 0, 1, 0
    pSelTextBox txtIva

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIVA_GotFocus"))
End Sub

Private Sub txtIVA_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Not vlblnConsulta Then
            chkActivo.SetFocus
        Else
            cmdSave.SetFocus
        End If
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
            KeyAscii = 7
        End If
    End If

End Sub
