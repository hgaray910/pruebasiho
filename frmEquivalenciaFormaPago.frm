VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEquivalenciaFormaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equivalencias de formas de pago"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "frmEquivalenciaFormaPago.frx":0000
      Left            =   3480
      List            =   "frmEquivalenciaFormaPago.frx":0007
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4035
      ScaleWidth      =   6915
      TabIndex        =   6
      Top             =   840
      Width           =   6975
      Begin MSFlexGridLib.MSFlexGrid grdEquivalencias 
         Height          =   4035
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   7117
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         AllowBigSelection=   0   'False
         ScrollBars      =   2
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin VB.ComboBox cboDepartamentoD 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.ComboBox cboDepartamentoF 
      Height          =   315
      ItemData        =   "frmEquivalenciaFormaPago.frx":000E
      Left            =   120
      List            =   "frmEquivalenciaFormaPago.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   3300
      TabIndex        =   7
      Top             =   5040
      Width           =   630
      Begin VB.CommandButton cmdSave 
         Enabled         =   0   'False
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000000&
         Picture         =   "frmEquivalenciaFormaPago.frx":0012
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Enviar transferencia"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Departamento destino"
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Departamento fuente"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmEquivalenciaFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intLastIndexD As Integer
Dim intlastindexF As Integer

Private Sub cboDepartamentoD_Click()
    If Me.cboDepartamentoD.ListIndex <> intLastIndexD Then
        If Me.cboDepartamentoD.ListIndex > -1 Then
            If Me.cmdSave.Enabled Then
                Select Case MsgBox(SIHOMsg(440), vbQuestion + vbYesNoCancel, "Mensaje")
                    Case vbCancel
                        Me.cmdSave.Enabled = False
                        Me.cboDepartamentoD.ListIndex = intLastIndexD
                        Me.cmdSave.Enabled = True
                        Exit Sub
                    Case vbYes
                        pGuardar
                End Select
            End If
            pLlenaList
            If Me.cboDepartamentoF.ListIndex > -1 Then
                pLimpiaColumnaT
                pLlenaGridT Me.cboDepartamentoF.ItemData(Me.cboDepartamentoF.ListIndex), Me.cboDepartamentoD.ItemData(Me.cboDepartamentoD.ListIndex)
            End If
            intLastIndexD = Me.cboDepartamentoD.ListIndex
            Me.cmdSave.Enabled = False
        End If
    End If
End Sub

Private Sub cboDepartamentoF_Click()
    If Me.cboDepartamentoF.ListIndex <> intlastindexF Then
        If Me.cmdSave.Enabled Then
            Select Case MsgBox(SIHOMsg(440), vbQuestion + vbYesNoCancel, "Mensaje")
                Case vbCancel
                    Me.cmdSave.Enabled = False
                    Me.cboDepartamentoF.ListIndex = intlastindexF
                    Me.cmdSave.Enabled = True
                    Exit Sub
                Case vbYes
                    pGuardar
            End Select
        End If
        pFormatGrid
        If Me.cboDepartamentoF.ListIndex > -1 Then
            pLlenaGridS Me.cboDepartamentoF.ItemData(Me.cboDepartamentoF.ListIndex)
            If Me.cboDepartamentoD.ListIndex > -1 Then
                pLlenaGridT Me.cboDepartamentoF.ItemData(Me.cboDepartamentoF.ListIndex), Me.cboDepartamentoD.ItemData(Me.cboDepartamentoD.ListIndex)
            End If
        End If
        intlastindexF = Me.cboDepartamentoF.ListIndex
        Me.cmdSave.Enabled = False
    End If
End Sub

Private Sub cmdSave_Click()
    pGuardar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.ActiveControl.Name <> "grdEquivalencias" And Me.ActiveControl.Name <> "List1" Then
        SendKeys vbTab
    End If
    If KeyAscii = 27 Then
        If Not Me.List1.Visible Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = frmMenuPrincipal.Icon
    
    intLastIndexD = -1
    intlastindexF = -1
    pFormatGrid
    pObtieneDepartamentos

End Sub

Private Sub pFormatGrid()
    Me.grdEquivalencias.Clear
    Me.grdEquivalencias.Rows = 2
    Me.grdEquivalencias.FormatString = "Forma de Pago||Forma de Pago||"
    Me.grdEquivalencias.ColWidth(0) = 3210
    Me.grdEquivalencias.ColWidth(1) = 130
    Me.grdEquivalencias.ColWidth(2) = 3210
    Me.grdEquivalencias.ColWidth(3) = 0
    Me.grdEquivalencias.ColWidth(4) = 0
    Me.grdEquivalencias.RowData(1) = 0
End Sub

Private Sub pObtieneDepartamentos()
    On Error GoTo NotificaError
    Dim rsDepartamentos As ADODB.Recordset
    Dim vlstrx As String
    vlstrx = "select distinct F.smiDepartamento, D.vchDescripcion from PvFormapago F " & _
    " inner join NoDepartamento D on D.smicvedepartamento = F.smiDepartamento " & _
    " where F.bitEstatusActivo = 1 and (f.chrTipo = 'E' or f.chrTipo = 'H')  and D.tnyclaveempresa = " & vgintClaveEmpresaContable & _
    " order by vchDescripcion"
    Set rsDepartamentos = frsRegresaRs(vlstrx)
    Do Until rsDepartamentos.EOF
        Me.cboDepartamentoF.AddItem rsDepartamentos!vchDescripcion
        Me.cboDepartamentoF.ItemData(Me.cboDepartamentoF.NewIndex) = rsDepartamentos!SMIDEPARTAMENTO
        Me.cboDepartamentoD.AddItem rsDepartamentos!vchDescripcion
        Me.cboDepartamentoD.ItemData(Me.cboDepartamentoD.NewIndex) = rsDepartamentos!SMIDEPARTAMENTO
        rsDepartamentos.MoveNext
    Loop
    rsDepartamentos.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pObtieneDepartamentos"))
End Sub

Private Sub pLlenaGridS(ByVal intDepto As Integer)
    Dim rsFormas As ADODB.Recordset
    Dim strX As String
    Dim intCtrl As Integer
    strX = "select * from PvFormaPago where PvFormaPago.bitEstatusActivo = 1 and (PvFormaPago.chrTipo = 'E' or PvFormaPago.chrTipo = 'H') and PvFormaPago.smiDepartamento = " & intDepto & " order  by chrDescripcion"
    Set rsFormas = frsRegresaRs(strX, adLockReadOnly, adOpenForwardOnly)
    intCtrl = 1
    Do Until rsFormas.EOF
        If intCtrl > 1 Then
            Me.grdEquivalencias.AddItem ""
        End If
        Me.grdEquivalencias.TextMatrix(intCtrl, 0) = rsFormas!chrDescripcion
        Me.grdEquivalencias.RowData(intCtrl) = rsFormas!intFormaPago
        intCtrl = intCtrl + 1
        rsFormas.MoveNext
    Loop
    rsFormas.Close
End Sub

Private Sub pLlenaGridT(ByVal intDeptoS As Integer, ByVal intDeptoT As Integer)
    Dim rsFormas As ADODB.Recordset
    Dim strX As String
    Dim intCtrl As Integer
    strX = "select S.intFormaPago as formaPS, T.intFormaPago as formaPT, T.chrDescripcion " & _
    "from PvFormaPago S inner join PvFormaPagoEquivalencia E on S.intFormaPago = E.intFormaPagoFuente " & _
    "inner join PvFormaPago T on T.intFormaPago = E.intFormaPagoDestino " & _
    "where S.smidepartamento = " & intDeptoS & " And T.smidepartamento = " & intDeptoT
    Set rsFormas = frsRegresaRs(strX, adLockReadOnly, adOpenForwardOnly)
    Do Until rsFormas.EOF
        intCtrl = fBuscarForma(rsFormas!formaPS)
        If intCtrl > 0 Then
            Me.grdEquivalencias.TextMatrix(intCtrl, 2) = rsFormas!chrDescripcion
            Me.grdEquivalencias.TextMatrix(intCtrl, 3) = rsFormas!formaPT
        End If
        rsFormas.MoveNext
    Loop
    rsFormas.Close
End Sub

Private Function fBuscarForma(ByVal idForma As Integer) As Integer
    Dim intCtrl As Integer
    For intCtrl = 1 To Me.grdEquivalencias.Rows - 1
        If Me.grdEquivalencias.RowData(intCtrl) = idForma Then
            fBuscarForma = intCtrl
            Exit Function
        End If
    Next
    fBuscarForma = 0
End Function

Private Sub pLimpiaColumnaT()
    Dim intCtrl As Integer
    For intCtrl = 1 To Me.grdEquivalencias.Rows - 1
        Me.grdEquivalencias.TextMatrix(intCtrl, 2) = ""
        Me.grdEquivalencias.TextMatrix(intCtrl, 3) = ""
        Me.grdEquivalencias.TextMatrix(intCtrl, 4) = ""
    Next
End Sub

Private Sub pLlenaList()
    Dim rsFormas As ADODB.Recordset
    Dim strX As String
    strX = "select * from PvFormaPago where PvFormaPago.bitEstatusActivo = 1 and (PvFormaPago.chrTipo = 'E' or PvFormaPago.chrTipo = 'H') and PvFormaPago.smiDepartamento = " & Me.cboDepartamentoD.ItemData(Me.cboDepartamentoD.ListIndex) & " order  by chrDescripcion"
    Set rsFormas = frsRegresaRs(strX, adLockReadOnly, adOpenForwardOnly)
    Me.List1.Clear
    Me.List1.Height = 210
    Me.List1.AddItem ""
    Me.List1.ItemData(Me.List1.NewIndex) = 0
    Do Until rsFormas.EOF
        Me.List1.AddItem rsFormas!chrDescripcion
        Me.List1.ItemData(Me.List1.NewIndex) = rsFormas!intFormaPago
        If Me.List1.Height < 1260 Then
            Me.List1.Height = Me.List1.Height + 210
        End If
        rsFormas.MoveNext
    Loop
    rsFormas.Close
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Me.List1.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Me.cmdSave.Enabled Then
        Select Case MsgBox(SIHOMsg(440), vbQuestion + vbYesNoCancel, "Mensaje")
            Case vbCancel
                Cancel = 1
                Exit Sub
            Case vbYes
                pGuardar
        End Select
    End If
End Sub

Private Sub grdEquivalencias_Click()
    If Me.grdEquivalencias.RowData(Me.grdEquivalencias.Row) <> 0 And Me.cboDepartamentoD.ListIndex <> -1 And Me.cboDepartamentoF.ListIndex <> Me.cboDepartamentoD.ListIndex Then
        If Me.grdEquivalencias.Col = 2 Then
            Me.List1.Top = Me.grdEquivalencias.RowPos(Me.grdEquivalencias.Row) + 850
            Me.List1.ListIndex = fintLocalizaTxtLst(Me.List1, Me.grdEquivalencias.Text)
            Me.List1.Visible = True
            Me.List1.SetFocus
        End If
    End If
End Sub

Private Sub grdEquivalencias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If Me.grdEquivalencias.Col = 2 Then
            Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 2) = ""
            Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 4) = "0"
        End If
    End If
End Sub

Private Sub grdEquivalencias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdEquivalencias_Click
    End If
End Sub

Private Sub grdEquivalencias_Scroll()
    Me.List1.Visible = False
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pPonerDatosEnGrid
    End If
    If KeyAscii = 27 Then
        Me.List1.Visible = False
        Me.grdEquivalencias.SetFocus
    End If
End Sub

Private Sub List1_Lostfocus()
    Me.List1.Visible = False
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pPonerDatosEnGrid
End Sub

Private Sub pPonerDatosEnGrid()
    Dim strNewData As String
    Dim strOldData As String
    strNewData = CStr(Me.List1.ItemData(Me.List1.ListIndex))
    strOldData = Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 3)
    Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 2) = Me.List1.List(Me.List1.ListIndex)
    If strNewData <> strOldData Then
        Me.cmdSave.Enabled = True
        Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 4) = strNewData
    Else
        Me.grdEquivalencias.TextMatrix(Me.grdEquivalencias.Row, 4) = ""
    End If
    Me.List1.Visible = False
    Me.grdEquivalencias.SetFocus
End Sub

Private Sub pGuardar()
    On Error GoTo NotificaError
    Dim intCtrl As Integer
    For intCtrl = 1 To Me.grdEquivalencias.Rows - 1
        If Me.grdEquivalencias.TextMatrix(intCtrl, 3) <> "" Then
            If Me.grdEquivalencias.TextMatrix(intCtrl, 4) <> "" Then
                If Me.grdEquivalencias.TextMatrix(intCtrl, 4) <> "0" Then
                    pEjecutaSentencia "update PvFormaPagoEquivalencia set intFormaPagoDestino = " & Me.grdEquivalencias.TextMatrix(intCtrl, 4) & " where intFormaPagoFuente = " & Me.grdEquivalencias.RowData(intCtrl) & " and intFormaPagoDestino = " & Me.grdEquivalencias.TextMatrix(intCtrl, 3)
                Else
                    pEjecutaSentencia "delete PvFormaPagoEquivalencia where intFormaPagoFuente = " & Me.grdEquivalencias.RowData(intCtrl) & " and intFormaPagoDestino = " & Me.grdEquivalencias.TextMatrix(intCtrl, 3)
                End If
            End If
        Else
            If Me.grdEquivalencias.TextMatrix(intCtrl, 4) <> "" Then
                If Me.grdEquivalencias.TextMatrix(intCtrl, 4) <> "0" Then
                    pEjecutaSentencia "insert into PvFormaPagoEquivalencia(intFormaPagoFuente, intFormaPagoDestino) values(" & Me.grdEquivalencias.RowData(intCtrl) & ", " & Me.grdEquivalencias.TextMatrix(intCtrl, 4) & ")"
                End If
            End If
        End If
    Next
    intLastIndexD = -1
    Me.cmdSave.Enabled = False
    cboDepartamentoD_Click
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pGuardar"))
End Sub

Public Function fintLocalizaTxtLst(ObjCbo As ListBox, vlstrCriterio As String) As Integer
    On Error GoTo NotificaError
    Dim vlintNumReg As Integer
    Dim vlintseq As Integer
    Dim vlintLargo As Integer
    vlintNumReg = ObjCbo.ListCount
    If Len(vlstrCriterio) > 0 Then
        vlintLargo = Len(vlstrCriterio)
        For vlintseq = 0 To vlintNumReg
            If UCase(Left(ObjCbo.List(vlintseq), vlintLargo)) = UCase(vlstrCriterio) Then
                fintLocalizaTxtLst = vlintseq
                Exit For
            Else
                fintLocalizaTxtLst = 0
            End If
        Next vlintseq
    Else
        fintLocalizaTxtLst = 0
    End If
    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fintLocalizaTxtLst"))
End Function

