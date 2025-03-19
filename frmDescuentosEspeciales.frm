VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDescuentosEspeciales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos especiales"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDescuentos 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Descuentos especiales"
      Top             =   2280
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   4895
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   10935
      Begin VB.CheckBox chkConsideraExcluidos 
         Caption         =   "Realizar el cálculo del descuento especial sin los artículos identificados como excluidos"
         Height          =   435
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Realizar el cálculo del descuento especial sin considerar los artículos excluidos"
         Top             =   1530
         Width           =   3975
      End
      Begin VB.CommandButton cmdAgregar 
         Height          =   495
         Left            =   10080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmDescuentosEspeciales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guardar descuento"
         Top             =   1440
         Width           =   495
      End
      Begin VB.CheckBox chkvigencia 
         Caption         =   "Vigencia"
         Height          =   255
         Left            =   6600
         TabIndex        =   4
         ToolTipText     =   "Marcar para asignar vigencia"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtmonto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Subtotal a partir del cual aplica el descuento especial"
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox txtporcentaje 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3240
         MaxLength       =   5
         TabIndex        =   1
         ToolTipText     =   "Porcentaje de descuento"
         Top             =   720
         Width           =   735
      End
      Begin VB.Frame fraVigencia 
         Height          =   855
         Left            =   6600
         TabIndex        =   13
         Top             =   600
         Width           =   3255
         Begin MSMask.MaskEdBox mskFecIni 
            Height          =   315
            Left            =   480
            TabIndex        =   5
            ToolTipText     =   "Fecha inicial de la vigencia"
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskFecFin 
            Height          =   315
            Left            =   1920
            TabIndex        =   6
            ToolTipText     =   "Fecha final de la vigencia"
            Top             =   345
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mmm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Al"
            Height          =   195
            Left            =   1680
            TabIndex        =   16
            Top             =   405
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Del"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   405
            Width           =   240
         End
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   3240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Empresa"
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "pesos"
         Height          =   195
         Left            =   4650
         TabIndex        =   17
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4080
         TabIndex        =   14
         Top             =   780
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto del subtotal para aplicar descuento"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   2970
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Porcentaje de descuento"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   780
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmDescuentosEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ObjRS As New ADODB.Recordset
Dim objSTR As String
Dim vllngPersonaGraba As Long
Private Sub chkVigencia_Click()
    Me.mskFecFin.Text = "  /  /    "
    Me.mskFecIni.Text = "  /  /    "
    Me.fraVigencia.Enabled = Me.chkVigencia.Value = vbChecked
    If Me.chkVigencia.Value = vbChecked Then Me.mskFecIni.SetFocus
End Sub
Private Sub cmdAgregar_Click()
Dim vllngResultado As Long
Dim strFechaInicial As String
Dim strFechafinal As String
Dim strRowAgregar As String

'Revisar permiso
If Not fblnRevisaPermiso(vglngNumeroLogin, 3059, "E", False) Then Exit Sub

If Not fValidaInformacion Then Exit Sub

If Me.chkVigencia.Value = vbUnchecked Then
   strFechaInicial = fstrFechaSQL(fdtmServerFecha, , True)
   strFechafinal = "3999-12-12"
Else
   strFechaInicial = fstrFechaSQL(Me.mskFecIni.Text, , True)
   strFechafinal = fstrFechaSQL(Me.mskFecFin.Text, , True)
End If

strRowAgregar = Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & Me.txtMonto.Text & "|" & Me.txtporcentaje.Text & "|" & strFechaInicial & "|" & strFechafinal & "|" & Me.chkConsideraExcluidos.Value
objSTR = Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & Format(Me.txtMonto.Text, "") & "|" & Me.txtporcentaje.Text & "|" & strFechaInicial & "|" & strFechafinal & "|" & Me.chkConsideraExcluidos.Value & "|0"

vllngResultado = 1
frsEjecuta_SP objSTR, "FN_PVINSPRECIOESPECIAL", True, vllngResultado

If vllngResultado > 0 Then
   '1331 ¡No se puede guardar la información! Ya existe un descuento en el rango de fechas asignado.
   MsgBox SIHOMsg(1331), vbExclamation + vbOKOnly, "Mensaje"
   If Me.chkVigencia.Value = vbChecked Then
      Me.mskFecIni.SetFocus
   Else
      Me.chkVigencia.SetFocus
   End If
   Exit Sub
End If

vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
If vllngPersonaGraba = 0 Then Exit Sub

objSTR = Me.cboEmpresa.ItemData(cboEmpresa.ListIndex) & "|" & Format(Me.txtMonto.Text, "") & "|" & Me.txtporcentaje.Text & "|" & strFechaInicial & "|" & strFechafinal & "|" & Me.chkConsideraExcluidos.Value & "|1"

vllngResultado = 1
frsEjecuta_SP objSTR, "FN_PVINSPRECIOESPECIAL", True, vllngResultado

If vllngResultado = 0 Then 'se agregó correctamente
   '358 ¡Los datos han sido guardados satisfactoriamente!
    MsgBox SIHOMsg(358), vbOKOnly + vbInformation, "Mensaje"
    pLimpiaForma
    pCargarGrid
    Call pGuardarLogTransaccion(Me.Name, 2, vllngPersonaGraba, "AGREGAR DESCUENTO ESPECIAL", strRowAgregar)
    Me.cboEmpresa.SetFocus
End If
End Sub
Private Sub Form_Activate()
If Me.cboEmpresa.ListCount = 0 Then
  '962 ¡No existen empresas configuradas!
   MsgBox Replace(SIHOMsg(962), "configuradas como clientes!", "activas!"), vbCritical + vbOKOnly, "Mensaje"
   Unload Me
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys vbTab
    ElseIf KeyAscii = 27 Then
            Unload Me
    End If
End Sub
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pLimpiaForma
    chkVigencia_Click
    pCargarGrid
End Sub
Private Sub pLimpiaForma()
    If cboEmpresa.ListCount = 0 Then
       objSTR = "Select vchDescripcion Descripcion, intCveEmpresa Clave From CcEmpresa where ccempresa.bitactivo = 1"
       Set ObjRS = frsRegresaRs(objSTR)
       pLlenarCboRs cboEmpresa, ObjRS, 1, 0, 0
    End If
    cboEmpresa.ListIndex = -1
    Me.mskFecFin.Text = "  /  /    "
    Me.mskFecIni.Text = "  /  /    "
    If chkVigencia.Value = vbChecked Then chkVigencia.Value = vbUnchecked
    Me.txtMonto.Text = ""
    Me.txtporcentaje.Text = ""
    chkConsideraExcluidos.Value = vcUnchecked
End Sub
Private Sub pCargarGrid()
    Dim blnPrimerRow As Boolean
    Dim intRow As Integer
    
    objSTR = "Select VCHDESCRIPCION,IntConsecutivo,Mnymontoaplicarlimite,NumPorcentaje,BitConsideraExcluidos,DtmFechainicial,DtmFechaFinal " & _
             " from PVDESCUENTOESPECIAL inner join ccempresa on ccempresa.intcveempresa = pvdescuentoespecial.intcveempresa order by VCHDESCRIPCION,DtmFechainicial ASC"
    Set ObjRS = frsRegresaRs(objSTR)
    pConfiguraGrid
    If ObjRS.RecordCount > 0 Then
       grdDescuentos.Redraw = False
       blnPrimerRow = True
       ObjRS.MoveFirst
       Do While Not ObjRS.EOF
          If blnPrimerRow Then
             intRow = 1
             blnPrimerRow = False
          Else
             grdDescuentos.Rows = grdDescuentos.Rows + 1
             intRow = grdDescuentos.Rows - 1
          End If
                    
          grdDescuentos.TextMatrix(intRow, 1) = ObjRS!intConsecutivo
          grdDescuentos.TextMatrix(intRow, 2) = ObjRS!VCHDESCRIPCION
          grdDescuentos.TextMatrix(intRow, 3) = ObjRS!NUMPORCENTAJE & " %"
          grdDescuentos.TextMatrix(intRow, 4) = FormatCurrency(ObjRS!MNYMONTOAPLICARLIMITE, 2)
          grdDescuentos.TextMatrix(intRow, 5) = IIf(IsNull(ObjRS!DTMFECHAINICIAL), "", IIf(ObjRS!DTMFECHAFINAL = "12/12/3999", "", Format(ObjRS!DTMFECHAINICIAL, "dd/MMM/YYYY")))
          grdDescuentos.TextMatrix(intRow, 6) = IIf(IsNull(ObjRS!DTMFECHAFINAL), "", IIf(ObjRS!DTMFECHAFINAL = "12/12/3999", "", Format(ObjRS!DTMFECHAFINAL, "dd/MMM/YYYY")))
          grdDescuentos.TextMatrix(intRow, 7) = IIf(ObjRS!BITCONSIDERAEXCLUIDOS = 1, "*", "")
          ObjRS.MoveNext
       Loop
       grdDescuentos.Redraw = True
    End If
End Sub
Private Sub pConfiguraGrid()
    With grdDescuentos
        .Clear
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 8
        .FormatString = "|id|Empresa|Descuento|Monto|Fecha inicio|Fecha fin|No considera excluidos"
        .ColWidth(0) = 100
        .ColWidth(1) = 0
        .ColWidth(2) = 4400
        .ColWidth(3) = 900
        .ColWidth(4) = 1200
        .ColWidth(5) = 1050
        .ColWidth(6) = 1050
        .ColWidth(7) = 1900
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignCenterCenter
    End With
End Sub
Private Function fValidaCantidad(vlintCaracter As Integer, vlintDecimales As Integer, CajaText As TextBox) As Boolean ' procedimiento para validar la cantidad que se introduce a un textbox
    Dim vlintPosicionCursor As Integer
    Dim vlintPosiciones As Integer
    Dim vlintPosicionPunto As Integer
    Dim vlintNumeroDecimales As Integer
    
    fValidaCantidad = True
    If Not IsNumeric(Chr(vlintCaracter)) Then 'no es numero
        If Not vlintCaracter = vbKeyBack Then 'no es retroceso
            If Not vlintCaracter = vbKeyReturn Then 'no es Enter
                If Not vlintCaracter = 46 Then ' no es el punto
                    fValidaCantidad = False ' se anula, estos son los unicos caracteres que se pueden ingresar al texbox
                Else 'es un punto debemos verificar si se tiene un punto ya en el text
                    If fblnValidaPunto(CajaText.Text) Then ' ya hay un punto
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    Else ' se intenta ingresar un caracter numerico, revisar decimales, revisar si se tiene seleccionado el textbox
        If CajaText.SelText <> CajaText.Text Then
            vlintPosicionCursor = CajaText.SelStart
            vlintPosicionPunto = InStr(1, CajaText.Text, ".")
            If vlintPosicionPunto > 0 Then ' si hay punto
                If vlintPosicionCursor > vlintPosicionPunto Then ' si la poscion es mayor entonces debemos de revisar los decimales
                    'contamos la cantidad de decimales
                    For vlintPosiciones = vlintPosicionPunto + 1 To Len(CajaText.Text)
                    vlintNumeroDecimales = vlintNumeroDecimales + 1
                    Next vlintPosiciones
                    'si ya son tantos como vlinDecimales entonces no permite la insercion
                    If vlintNumeroDecimales >= vlintDecimales Then
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    End If
End Function
Private Sub grdDescuentos_DblClick()
    Dim strRowEliminado As String
    'Revisar permiso
    If Not fblnRevisaPermiso(vglngNumeroLogin, 3059, "E", True) Then
       'El usuario no tiene permiso para realizar esta operación
       MsgBox SIHOMsg(635), vbExclamation, "Mensaje"
       Exit Sub
    End If
    If grdDescuentos.Row > 0 And Trim(grdDescuentos.TextMatrix(grdDescuentos.Row, 1)) <> "" Then
       vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
       If vllngPersonaGraba = 0 Then Exit Sub
      
       objSTR = "select * from pvdescuentoespecial where intconsecutivo = " & Me.grdDescuentos.TextMatrix(grdDescuentos.Row, 1)
       Set ObjRS = frsRegresaRs(objSTR, adLockOptimistic)
       
       If ObjRS.RecordCount > 0 Then
          strRowEliminado = ObjRS!intcveempresa & "|" & ObjRS!MNYMONTOAPLICARLIMITE & "|" & ObjRS!NUMPORCENTAJE & "|" & ObjRS!DTMFECHAINICIAL & "|" & ObjRS!DTMFECHAFINAL
          pEjecutaSentencia "Delete from Pvdescuentoespecial where intconsecutivo = " & Me.grdDescuentos.TextMatrix(grdDescuentos.Row, 1)
          Call pGuardarLogTransaccion(Me.Name, 3, vllngPersonaGraba, "ELIMINAR DESCUENTO ESPECIAL", strRowEliminado)
          '501 La información fue eliminada.
          MsgBox SIHOMsg(501), vbOKOnly + vbInformation, "Mensaje"
          pLimpiaForma
          pCargarGrid
          Me.cboEmpresa.SetFocus
       End If
    End If
End Sub
Private Sub mskFecFin_GotFocus()
    pEnfocaMkTexto mskFecFin
End Sub
Private Sub mskFecIni_GotFocus()
    pEnfocaMkTexto mskFecIni
End Sub
Private Sub txtMonto_GotFocus()
    txtMonto.Text = Format(txtMonto.Text, "")
    pEnfocaTextBox txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 2, txtMonto) Then KeyAscii = 0
End Sub
Private Sub txtMonto_LostFocus()
 If Trim(txtMonto.Text) <> "" Then txtMonto.Text = FormatCurrency(Val(txtMonto.Text), 2)
End Sub
Private Sub txtporcentaje_GotFocus()
    pEnfocaTextBox txtporcentaje
End Sub
Private Sub txtporcentaje_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 2, txtporcentaje) Then KeyAscii = 0
End Sub
Private Function fValidaInformacion() As Boolean
        fValidaInformacion = True
        
        If Me.cboEmpresa.ListIndex = -1 Then
           '3 ¡Dato no válido, seleccione un valor de la lista!
           MsgBox SIHOMsg(3), vbExclamation + vbOKOnly, "Mensaje"
           Me.cboEmpresa.SetFocus
           fValidaInformacion = False
           Exit Function
        End If
        If Trim(Me.txtporcentaje.Text) = "" Then
           '406 Dato incorrecto.
           MsgBox Replace(SIHOMsg(406), ".", ": ") & txtporcentaje.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
           Me.txtporcentaje.SetFocus
           fValidaInformacion = False
           Exit Function
        End If
        If Val(Me.txtporcentaje.Text) = 0 Or Val(Me.txtporcentaje.Text) > 100 Then
           '406 Dato incorrecto.
           MsgBox Replace(SIHOMsg(406), ".", ": ") & txtporcentaje.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
           Me.txtporcentaje.SetFocus
           fValidaInformacion = False
           Exit Function
        End If
        If Trim(Me.txtMonto.Text) = "" Then
           '406 Dato incorrecto.
           MsgBox Replace(SIHOMsg(406), ".", ": ") & txtMonto.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
           Me.txtMonto.SetFocus
           fValidaInformacion = False
           Exit Function
        End If
        If Val(Format(Me.txtMonto.Text, "")) = 0 Then
           '406 Dato incorrecto.
           MsgBox Replace(SIHOMsg(406), ".", ": ") & txtMonto.ToolTipText, vbOKOnly + vbExclamation, "Mensaje"
           Me.txtMonto.SetFocus
           fValidaInformacion = False
           Exit Function
        End If
        If Me.chkVigencia.Value = vbChecked Then
           If Not IsDate(Me.mskFecIni.Text) Then
              '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
              MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
              mskFecIni.SetFocus
              fValidaInformacion = False
              Exit Function
           End If
           If CDate(Me.mskFecIni.Text) < fdtmServerFecha Then
              '43 ¡La fecha debe ser mayor o igual a la del sistema!
              MsgBox Replace(SIHOMsg(43), "debe", "inicial debe"), vbExclamation + vbOKOnly, "Mensaje"
              mskFecIni.SetFocus
              fValidaInformacion = False
              Exit Function
           End If
           If Not IsDate(Me.mskFecFin.Text) Then
              '29 ¡Fecha no válida! Formato de fecha dd/mm/aaaa.
              MsgBox SIHOMsg(29), vbExclamation + vbOKOnly, "Mensaje"
              mskFecFin.SetFocus
              fValidaInformacion = False
              Exit Function
           End If
           If CDate(Me.mskFecIni) > CDate(Me.mskFecFin) Then
              '379 ¡La fecha final debe ser mayor a la fecha inicial!
              MsgBox SIHOMsg(379), vbExclamation + vbOKOnly, "Mensaje"
              mskFecFin.SetFocus
              fValidaInformacion = False
              Exit Function
           End If
        End If
End Function
