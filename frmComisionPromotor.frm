VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmComisionPromotor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones para promotores"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   4575
      Begin VB.TextBox txtPorcentaje 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   5
         EndProperty
         Height          =   285
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   3
         ToolTipText     =   "Porcentaje de comisión"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4200
         TabIndex        =   12
         Top             =   255
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje de comisión"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   255
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   495
      Left            =   4800
      MaskColor       =   &H80000014&
      Picture         =   "frmComisionPromotor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Agregar"
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Tipo paciente / empresa"
         Top             =   600
         Width           =   4815
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Empresa"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         ToolTipText     =   "Empresa"
         Top             =   255
         Width           =   975
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Tipo de paciente"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Tipo de paciente"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   2400
      TabIndex        =   8
      Top             =   4650
      Width           =   600
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   50
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmComisionPromotor.frx":04F2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Grabar"
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5175
      Begin VSFlex7LCtl.VSFlexGrid vsfComisiones 
         Height          =   2325
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Tipo de paciente / empresa"
         Top             =   240
         Width           =   4905
         _cx             =   8652
         _cy             =   4101
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmComisionPromotor.frx":0834
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   0
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
End
Attribute VB_Name = "frmComisionPromotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim llngRowActualizar As Long
Public llngNumOpcion As Long


Private Sub cmdAgregar_Click()
On Error GoTo NotificaError

    If Val(txtPorcentaje.Text) = 0 Then
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2) & Chr(13) & txtPorcentaje.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
        txtPorcentaje.SetFocus
        Exit Sub
    End If
    If Val(Format(txtPorcentaje.Text, "")) > 100 Then
        'Dato incorrecto: El porcentaje debe ser menor a 100%
        MsgBox SIHOMsg(35), vbCritical, "Mensaje"
        txtPorcentaje.SetFocus
        Exit Sub
    End If
    
    llngRowActualizar = 0
    If fblnValido Then
        With vsfComisiones
            If llngRowActualizar <> 0 Then
                .TextMatrix(llngRowActualizar, 2) = Trim(txtPorcentaje.Text)
            Else
                .AddItem "" & vbTab & Trim(cboTipo.List(cboTipo.ListIndex)) & vbTab & Trim(txtPorcentaje.Text) & vbTab & cboTipo.ItemData(cboTipo.ListIndex) & vbTab & IIf(optTipo(0).Value, "P", "E")
                vsfComisiones.Col = 1
                vsfComisiones.Sort = flexSortGenericAscending
            End If
        End With
    End If
    cboTipo.SetFocus
    txtPorcentaje.Text = ""
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAgregar_Click"))
End Sub
Private Function fblnValido() As Boolean
On Error GoTo NotificaError
    Dim llngRow As Long
    
    fblnValido = True
    With vsfComisiones
        For llngRow = 1 To .Rows - 1
            If Val(.TextMatrix(llngRow, 3)) = cboTipo.ItemData(cboTipo.ListIndex) And Trim(.TextMatrix(llngRow, 4)) = IIf(optTipo(0).Value, "P", "E") Then
                If Val(.TextMatrix(llngRow, 2)) = Val(txtPorcentaje.Text) Then
                    'Existe información con el mismo contenido.
                    MsgBox SIHOMsg(19), vbOKOnly + vbExclamation, "Mensaje"
                    fblnValido = False
                Else
                    'Desea actualizar los datos
                    If MsgBox(SIHOMsg(1006), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
                        fblnValido = True
                        llngRowActualizar = llngRow
                    Else
                        fblnValido = False
                    End If
                End If
                Exit For
            End If
        Next llngRow
    End With

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnValido"))
End Function
Private Sub cmdSave_Click()
On Error GoTo NotificaError
    Dim llngRow As Long
    Dim llngPersonaGraba As Long
    Dim lstrSentencia As String
    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Then
        llngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If llngPersonaGraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans

        lstrSentencia = "DELETE FROM PVCOMISIONPROMOTORES"
        pEjecutaSentencia lstrSentencia
        With vsfComisiones
            For llngRow = 1 To .Rows - 1
                vgstrParametrosSP = Trim(.TextMatrix(llngRow, 4)) & "|" & Trim(.TextMatrix(llngRow, 3)) & "|" & Trim(.TextMatrix(llngRow, 2))
                frsEjecuta_SP vgstrParametrosSP, "SP_PVINSCOMISIONPROMOTORES"
            Next llngRow
        End With
        
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, llngPersonaGraba, "COMISIONES PARA PROMOTORES", CStr(vgintNumeroDepartamento))
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        pConfiguraGrid
        pLlenaGrid
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
        
        If cboTipo.ListCount > 0 Then cboTipo.ListIndex = 0
        cboTipo.SetFocus
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
        pConfiguraGrid
        pLlenaGrid
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    pConfiguraGrid
    pLlenaGrid
    optTipo(0).Value = True
    
    txtPorcentaje.Text = Format(txtPorcentaje.Text, "###.##")
    
End Sub

Private Sub optTipo_Click(Index As Integer)
On Error GoTo NotificaError
    Dim rsAux As New ADODB.Recordset
    
    cboTipo.Clear
    If Index = 0 Then
        Set rsAux = frsEjecuta_SP("-1", "SP_ADSELTIPOPACIENTE")
        If rsAux.RecordCount > 0 Then
            Do While Not rsAux.EOF
                If rsAux!bitactivo = 1 Then
                    cboTipo.AddItem Trim(rsAux!vchDescripcion)
                    cboTipo.ItemData(cboTipo.NewIndex) = rsAux!tnyCveTipoPaciente
                End If
                rsAux.MoveNext
            Loop
        End If
    Else
        Set rsAux = frsEjecuta_SP("-1|-1|1", "SP_CCSELEMPRESA")
        If rsAux.RecordCount > 0 Then
            pLlenarCboRs cboTipo, rsAux, 0, 1
        End If
    End If
    If cboTipo.ListCount > 0 Then cboTipo.ListIndex = 0
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub


Private Sub txtPorcentaje_GotFocus()
    pEnfocaTextBox txtPorcentaje
End Sub

Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        
        pFormatoTexto
        
    End If
    
End Sub

Private Sub pConfiguraGrid()
On Error GoTo NotificaError

    With vsfComisiones
        .Clear
        .Rows = 1
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Tipo de paciente/Empresa|Porcentaje|"
        .ColWidth(0) = 100      '
        .ColWidth(1) = 3550     'tipo Paciente/empresa
        .ColWidth(2) = 930     'Porcentaje de comision
        .ColAlignment(2) = flexAlignRightCenter
        '.ColFormat(2) = "%"
        .ColWidth(3) = 0        'cve tipoPaciente o empresa
        .ColWidth(4) = 0        'P= por tipo de paciente, E por empresa
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub
Private Sub pLlenaGrid()
On Error GoTo NotificaError
    Dim rsComisiones As New ADODB.Recordset
    
    vgstrParametrosSP = "-1|'*'"
    Set rsComisiones = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCOMISIONPROMOTORES")
    With vsfComisiones
        Do While Not rsComisiones.EOF
            .AddItem "" & vbTab & rsComisiones!Descripcion & vbTab & rsComisiones!Comision & vbTab & rsComisiones!CveTipo & vbTab & rsComisiones!Tipo
            rsComisiones.MoveNext
        Loop
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaGrid"))
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    
    On Error GoTo NotificaError
        
        If Not fblnFormatoCantidad(txtPorcentaje, KeyAscii, 2) Then
            KeyAscii = 7
        End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtPorcentaje_KeyPress"))
    
End Sub

Private Sub txtPorcentaje_LostFocus()
    
    pFormatoTexto
    
End Sub

Private Sub vsfComisiones_DblClick()
On Error GoTo NotificaError

    If vsfComisiones.Row > 0 Then
        If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C", True) Then
           '¿Está seguro de eliminar los datos?
           If MsgBox(SIHOMsg(6), vbQuestion + vbYesNo, "Mensaje") = vbYes Then
               vsfComisiones.RemoveItem (vsfComisiones.Row)
           End If
        Else
            '¡El usuario debe tener permiso de control total para eliminar los datos!
            MsgBox SIHOMsg(810), vbOKOnly + vbExclamation, "Mensaje"
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":vsfComisiones_DblClick"))
End Sub

Private Sub pFormatoTexto()

On Error GoTo NotificaError

Dim vlstrFormatoPorc As String
    
    vlstrFormatoPorc = "###.00"
    
    txtPorcentaje = Format(IIf(Val(txtPorcentaje.Text) > 100, "100", Val(txtPorcentaje.Text)), vlstrFormatoPorc)

Exit Sub

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoTexto"))
    
End Sub
