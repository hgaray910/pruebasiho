VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{2EBF261C-FD36-46DA-8D79-010C5D8D7036}#2.1#0"; "MyCommandButton.ocx"
Begin VB.Form frmInformacionFaltantePaciente 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información faltante y notas del paciente"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNoMostrarMas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No mostrar más esta pantalla"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   3480
      TabIndex        =   5
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox txtClavePaciente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   1
      ToolTipText     =   "Número de cuenta del paciente"
      Top             =   160
      Width           =   1245
   End
   Begin VB.ListBox lstInformacionFaltante 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   75
      TabIndex        =   2
      Top             =   980
      Width           =   6660
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   3
      ToolTipText     =   "Nombre del paciente"
      Top             =   570
      Width           =   4450
   End
   Begin MyCommandButton.MyButton cmdAceptar 
      Height          =   375
      Left            =   75
      TabIndex        =   6
      Top             =   5550
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      BackColorOver   =   -2147483633
      BackColorFocus  =   -2147483633
      BackColorDisabled=   -2147483633
      BorderColor     =   -2147483627
      TransparentColor=   16777215
      Caption         =   "Aceptar"
      DepthEvent      =   1
      ShowFocus       =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfNotas 
      Height          =   3210
      Left            =   75
      TabIndex        =   7
      ToolTipText     =   "Notas urgentes"
      Top             =   2310
      Width           =   6660
      _cx             =   11747
      _cy             =   5662
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmInformacionFaltantePaciente.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Número de paciente"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   0
      Top             =   220
      Width           =   2415
   End
   Begin VB.Label lblNombre 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   4
      Top             =   630
      Width           =   795
   End
End
Attribute VB_Name = "frmInformacionFaltantePaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vgintCuentaPaciente As Long

Dim rsCamposMandatorios As New ADODB.Recordset
Dim rsValorCampo As New ADODB.Recordset
Dim rsNombrePaciente As New ADODB.Recordset

Dim vlstrsql As String

Private Sub cmdAceptar_Click()
  Dim iCont As Integer
  Dim bTiene As Boolean
  bTiene = False
  For iCont = 1 To vsfNotas.Rows - 1
    If vsfNotas.TextMatrix(iCont, 4) = 0 And Trim(vsfNotas.TextMatrix(iCont, 3)) <> "" Then
      bTiene = True
    End If
  Next iCont
  
  If bTiene Or chkNoMostrarMas Then
    vgstrGrabaMedicoEmpleado = vgstrMedicoEnfermera
    vglngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento, vgstrGrabaMedicoEmpleado)
    If vglngPersonaGraba = 0 Then
        chkNoMostrarMas.Value = 0
        lstInformacionFaltante.SetFocus
        Exit Sub
    Else
        With vsfNotas
          For iCont = 1 To .Rows - 1
            If .TextMatrix(iCont, 4) = 0 And Trim(.TextMatrix(iCont, 3)) <> "" Then
              vlstrsql = "insert into SINOTAURGENTE (DTMFECHA, INTCVEDEPARTAMENTO, INTCVEEMPLEADO, VCHNOTA, NUMNUMCUENTA) values (" & _
                      fstrFechaSQL(fdtmServerFecha) & "," & _
                      vgintNumeroDepartamento & ", " & vglngPersonaGraba & ", '" & Trim(.TextMatrix(iCont, 3)) & "'," & vgintCuentaPaciente & ")"
              pEjecutaSentencia vlstrsql
              vlstrsql = "delete from ExLogMsgPaciente where intCuentaPaciente = " & vgintCuentaPaciente & " and INTNUMLOGIN = " & vglngNumeroLogin
              pEjecutaSentencia vlstrsql
            End If
          Next iCont
        End With
        If chkNoMostrarMas Then
          vlstrsql = "INSERT INTO ExLogMsgPaciente (intCuentaPaciente,intPersonaGraba,dtmFecha, INTNUMLOGIN) VALUES (" & vgintCuentaPaciente & "," & vglngPersonaGraba & ",GetDate()," & vglngNumeroLogin & ")"
          Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vglngPersonaGraba, "INFORMACION FALTANTE DEL PACIENTE", CStr(vgintCuentaPaciente))
          pEjecutaSentencia vlstrsql
        End If
    End If
  End If
  Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
End Sub

Private Sub Form_Activate()

    cmdAceptar.SetFocus

End Sub

Private Sub Form_Load()
  Dim vlintContador As Integer
  Dim rsNota As New ADODB.Recordset
  Dim SQL As String
  
    vlstrsql = "SELECT vchApellidoPaterno ||' '|| vchApellidoMaterno ||' '|| vchNombre AS Nombre, AdAdmision.numCvePaciente AS CvePaciente FROM AdAdmision " & _
               "INNER JOIN AdPaciente ON AdPaciente.numCvePaciente = AdAdmision.numCvePaciente " & _
               "WHERE numNumCuenta = " & vgintCuentaPaciente
    Set rsNombrePaciente = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)

    txtClavePaciente.Text = rsNombrePaciente!CvePaciente
    txtNombre.Text = rsNombrePaciente!Nombre

    rsNombrePaciente.Close

    vlstrsql = "SELECT vchNombreNemonico,vchNombreReal,vchNombreTabla,vchTipoDato FROM SiCamposMandatoriosPaciente WHERE bitValida = 1"
    Set rsCamposMandatorios = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)

    vlintContador = 0

    Do While Not rsCamposMandatorios.EOF

        vlstrsql = "SELECT " & rsCamposMandatorios!vchNombreReal & " AS Valor FROM " & rsCamposMandatorios!vchNombreTabla & " WHERE numCvePaciente = " & txtClavePaciente.Text
        Set rsValorCampo = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)

        If ((rsCamposMandatorios!vchTipoDato = "nvarchar" _
                Or rsCamposMandatorios!vchTipoDato = "char" _
                Or rsCamposMandatorios!vchTipoDato = "varchar") _
                And (rsValorCampo!valor = "0" _
                Or rsValorCampo!valor = " " _
                Or rsValorCampo!valor = "" _
                Or IsNull(rsValorCampo!valor))) _
            Or ((rsCamposMandatorios!vchTipoDato = "numeric" _
                Or rsCamposMandatorios!vchTipoDato = "decimal" _
                Or rsCamposMandatorios!vchTipoDato = "tinyint" _
                Or rsCamposMandatorios!vchTipoDato = "int" _
                Or rsCamposMandatorios!vchTipoDato = "money") _
                And (rsValorCampo!valor = "0" _
                Or IsNull(rsValorCampo!valor))) Then
            lstInformacionFaltante.AddItem rsCamposMandatorios!vchNombreNemonico, vlintContador
            vlintContador = vlintContador + 1
        End If

        rsCamposMandatorios.MoveNext

    Loop

    If rsCamposMandatorios.RecordCount <> 0 Then
        rsValorCampo.Close
    End If

    rsCamposMandatorios.Close
  
  SQL = "Select DTMFECHA, INTCVEDEPARTAMENTO, INTCVEEMPLEADO, VCHNOTA "
  SQL = SQL & "From SINOTAURGENTE "
  SQL = SQL & "WHERE numNumCuenta = " & vgintCuentaPaciente
  Set rsNota = frsRegresaRs(SQL)
  With rsNota
    If .RecordCount > 0 Then
      Do While Not .EOF
        If vsfNotas.TextMatrix(vsfNotas.Rows - 1, 3) <> "" Then
          vsfNotas.Rows = vsfNotas.Rows + 1
        End If
        vsfNotas.TextMatrix(vsfNotas.Rows - 1, 1) = !intCveDepartamento
        vsfNotas.TextMatrix(vsfNotas.Rows - 1, 2) = !dtmfecha
        vsfNotas.TextMatrix(vsfNotas.Rows - 1, 3) = !vchNota
        vsfNotas.TextMatrix(vsfNotas.Rows - 1, 4) = 1
        vsfNotas.TextMatrix(vsfNotas.Rows - 1, 5) = !intCveEmpleado
        .MoveNext
      Loop
      vsfNotas.Rows = vsfNotas.Rows + 1
    End If
    .Close
    vsfNotas.TextMatrix(vsfNotas.Rows - 1, 1) = vgintNumeroDepartamento
    vsfNotas.TextMatrix(vsfNotas.Rows - 1, 2) = Date
    vsfNotas.TextMatrix(vsfNotas.Rows - 1, 4) = 0
    vsfNotas.TextMatrix(vsfNotas.Rows - 1, 5) = vglngNumeroEmpleado
    vsfNotas.AutoSize 0, 4
    Me.Icon = frmMenuPrincipal.Icon
  
  End With
End Sub
Private Sub vsfNotas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
  With vsfNotas
    If .TextMatrix(.Rows - 1, 3) <> "" Then
'      .AddItem fdtmServerFecha
        .AddItem .Rows & vbTab & vgintNumeroDepartamento & vbTab & Date & vbTab & vbTab & "0"
    End If
    .AutoSize 3
  End With
End Sub
Private Sub vsfNotas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Col = 2 Then
    Cancel = True
  Else
    If dVal(vsfNotas.TextMatrix(Row, 1)) <> vgintNumeroDepartamento Then
      Cancel = True
    ElseIf dVal(vsfNotas.TextMatrix(Row, 4)) <> 0 Then
      Cancel = True
    End If
  End If
End Sub

Private Sub vsfNotas_DblClick()
  With vsfNotas
    If .TextMatrix(.Row, 4) = 0 And Trim(.TextMatrix(.Row, 3)) <> "" Then
      If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo) = vbYes Then
        .RemoveItem .Row
        .AutoSize 3
      End If
    ElseIf .TextMatrix(.Row, 4) = 1 Then
      If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo) = vbYes Then
        vgstrGrabaMedicoEmpleado = vgstrMedicoEnfermera
        vglngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento, vgstrGrabaMedicoEmpleado)
        If vglngPersonaGraba = 0 Then
            chkNoMostrarMas.Value = 0
            lstInformacionFaltante.SetFocus
            Exit Sub
        Else
          If dVal(.TextMatrix(.Row, 5)) = vglngPersonaGraba Then
            .RemoveItem .Row
            .AutoSize 3
          Else
            MsgBox "El usuario no creo la nota, por lo tanto no puede borrar el registro.", vbInformation, "Mensaje"
          End If
        End If
      End If
    End If
  End With
End Sub

Private Sub vsfNotas_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
      vsfNotas_DblClick
  End If
End Sub

Private Sub vsfNotas_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
