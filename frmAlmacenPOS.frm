VERSION 5.00
Begin VB.Form frmAlmacenPOS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manejo almacenes venta público"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleMode       =   0  'User
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   3435
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAlmacenPOS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Grabar"
      Top             =   3525
      Width           =   495
   End
   Begin VB.PictureBox grdRelacion 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   6915
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Doble click para eliminar registro."
      Top             =   1320
      Width           =   6975
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox cboAlmacen 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Seleccione el almacén."
         Top             =   600
         Width           =   5175
      End
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         ItemData        =   "frmAlmacenPOS.frx":0342
         Left            =   1560
         List            =   "frmAlmacenPOS.frx":0344
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Seleccione el departamento."
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Departamento"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   3375
      TabIndex        =   7
      Top             =   3375
      Width           =   600
   End
End
Attribute VB_Name = "frmAlmacenPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Option Explicit
   Private bKillKey As Boolean
    Dim vlstrSentencia As String
    Dim rs As New ADODB.Recordset
    Dim rsEstatus As New ADODB.Recordset
Private Sub cboAlmacen_Click()
    If Not cboDepartamento.ListIndex < 0 Then
        If Not cboAlmacen.ListIndex < 0 Then
            cmdSave.Enabled = True
        End If
    End If
End Sub
Private Sub cboAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If Not cboDepartamento.ListIndex < 0 Then
        If Not cboAlmacen.ListIndex < 0 Then
            cmdSave.Enabled = True
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If cmdSave.Enabled = True Then
            cmdSave.SetFocus
        End If
        cboAlmacen.Enabled = False
    End If
End Sub
Private Sub cboDepartamento_Change()
    cboAlmacen.Enabled = True
End Sub
Private Sub cboDepartamento_Click()
    If cboDepartamento.ListIndex > -1 Then
        cboAlmacen.Enabled = True
    End If
End Sub
Private Sub cboDepartamento_GotFocus()
    cboAlmacen.Enabled = False
End Sub
Private Sub cboDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then
            If cboDepartamento.ListIndex > -1 Then
                cboAlmacen.Enabled = True
                SendKeys vbTab
            Else
                SendKeys vbTab
            End If
        End If
End Sub
Private Sub cboDepartamento_KeyPress(KeyAscii As Integer)
    If cboDepartamento.ListIndex > 0 Then
        cboAlmacen.Enabled = True
    End If
End Sub
Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    Dim intCtrl As Integer
    vlstrSentencia = "select * from pvAlmacenes"
    Set rs = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    With rs
        .AddNew
        !intnumDepartamento = cboDepartamento.ItemData(cboDepartamento.ListIndex)
        !vchDescripcion = cboDepartamento.Text
        !intnumalmacen = cboAlmacen.ItemData(cboAlmacen.ListIndex)
        !vchdescalmacen = cboAlmacen.Text
         .Update
            End With
    Call pLlenaGrid
    grdRelacion.Refresh
    Call pCargaCombos
    If cboDepartamento.Enabled = True Then
        cboDepartamento.SetFocus
    End If
    If Not grdRelacion.TextMatrix(1, 1) = "" Then
        grdRelacion.Enabled = True
        grdRelacion.RowSel = 1
    End If
    
   Me.cmdSave.Enabled = False
   Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdsave"))
End Sub

Private Sub cmdSave_LostFocus()
    cboDepartamento.SetFocus
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
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    Call pCargaCombos
    Call pLlenaGrid
    Call pFormatGrid
    Call pObtenerDatos
    cmdSave.Enabled = False
End Sub
Private Sub pObtenerDatos()
    vlstrSentencia = "select * from pvAlmacenes"
    Set rsEstatus = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
End Sub
Private Sub pCargaCombos()
    vlstrSentencia = "Select smicveDepartamento,RTrim (vchDescripcion) From noDepartamento Where chrClasificacion = 'A' and bitEstatus = 1  And tnyClaveEmpresa = " & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboAlmacen, rs, 0, 1
    rs.Close
    vlstrSentencia = "Select  smicveDepartamento, RTrim(vchDescripcion) From nodepartamento where smicveDepartamento not in (select intnumdepartamento from pvalmacenes)" & _
    "and chrClasificacion <> 'A' and bitEstatus = 1  And tnyClaveEmpresa =" & vgintClaveEmpresaContable
    Set rs = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    pLlenarCboRs cboDepartamento, rs, 0, 1
    If rs.RecordCount = 0 Then
        cboDepartamento.Enabled = False
        cboAlmacen.Enabled = False
    End If
    rs.Close
End Sub

Private Sub pLlenaGrid()
    grdRelacion.Rows = 2
    Dim rsFormas As ADODB.Recordset
    Dim strX As String
    Dim intCtrl As Integer
    strX = "Select * from pvalmacenes"
    Set rsFormas = frsRegresaRs(strX)
    intCtrl = 1
    Do Until rsFormas.EOF
        If intCtrl > 1 Then
            Me.grdRelacion.AddItem ""
        End If
        Me.grdRelacion.TextMatrix(intCtrl, 0) = rsFormas!intnumDepartamento
        Me.grdRelacion.TextMatrix(intCtrl, 1) = rsFormas!vchDescripcion
        Me.grdRelacion.TextMatrix(intCtrl, 2) = rsFormas!intnumalmacen
        Me.grdRelacion.TextMatrix(intCtrl, 3) = rsFormas!vchdescalmacen
        intCtrl = intCtrl + 1
        rsFormas.MoveNext
    Loop
    rsFormas.Close
End Sub
Private Sub pFormatGrid()
    With Me.grdRelacion
        .ColWidth(0) = 0
        .ColWidth(1) = 3295
        .ColWidth(2) = 0
        .ColWidth(3) = 3295
        .ColWidth(4) = 0
        .Row = 0
        .Text = ""
        .Col = 1
        .Text = "Departamento"
        .Col = 2
        .Text = ""
        .Col = 3
        .Text = "Almacén"
    End With
End Sub
Private Sub grdRelacion_DblClick()
   If grdRelacion.RowSel = 0 Then grdRelacion.RowSel = 1
   
   Dim vlbflag As Boolean
    If vlbflag = True Then grdRelacion.RemoveItem (3)
    vlbflag = False
    Dim vlstrRowDel As String
    Dim vlintRow As Integer
    Dim vlintCol As Integer
    vlintRow = grdRelacion.RowSel
    vlintCol = grdRelacion.ColSel - 1
    If grdRelacion.Rows = 2 Then
        grdRelacion.Rows = 3
        vlbflag = True
    End If
    If grdRelacion.TextMatrix(1, 1) = "" Then
        grdRelacion.Enabled = False
    Else
        
            If MsgBox("¿Está seguro de eliminar los datos?", vbExclamation + vbOKCancel, "Mensaje") = vbOK Then
               Call pObtenerDatos
               If Not rsEstatus.EOF Then
                    pEjecutaSentencia "delete from pvalmacenes where pvalmacenes.intnumdepartamento=" & grdRelacion.TextMatrix(vlintRow, 0)
                    vlstrRowDel = grdRelacion.RowSel
                    grdRelacion.RemoveItem (vlstrRowDel)
                    Call pCargaCombos
                End If
            End If
      
    End If
    cboDepartamento.Enabled = True
    cmdSave.Enabled = False
    If cboDepartamento.Enabled = True Then
        cboDepartamento.SetFocus
    End If
    
End Sub

Private Sub grdRelacion_SelChange()
    If grdRelacion.Row - grdRelacion.RowSel <> 0 Then
       
       grdRelacion.Row = grdRelacion.RowSel
       
       grdRelacion.SetFocus
    End If
End Sub
