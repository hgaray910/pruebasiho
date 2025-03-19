VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCargosSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos socios"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCargos 
      Caption         =   "Cargos del socio"
      Height          =   2790
      Left            =   120
      TabIndex        =   15
      Top             =   5745
      Width           =   10230
      Begin VB.TextBox txtMonto 
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCargos 
         Height          =   2415
         Left            =   105
         TabIndex        =   12
         ToolTipText     =   "Lista de cargos relacionados con el socio"
         Top             =   240
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4260
         _Version        =   393216
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraSociosDisponibles 
      Caption         =   "Socios disponibles"
      Height          =   3120
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   10230
      Begin VB.CommandButton cmdAplicarCargos 
         Caption         =   "Aplicar cargos"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8780
         TabIndex        =   11
         ToolTipText     =   "Aplicar cargos"
         Top             =   2685
         Width           =   1350
      End
      Begin VB.CommandButton cmdInvertirSeleccion 
         Caption         =   "Invertir selección"
         Enabled         =   0   'False
         Height          =   375
         Left            =   100
         TabIndex        =   10
         ToolTipText     =   "Invertir selección"
         Top             =   2685
         Width           =   1350
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdSocios 
         Height          =   2415
         Left            =   105
         TabIndex        =   9
         ToolTipText     =   "Lista de socios"
         Top             =   240
         Width           =   10005
         _ExtentX        =   17648
         _ExtentY        =   4260
         _Version        =   393216
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         ScrollBars      =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraOtroConceptos 
      Caption         =   "Otros conceptos a incluir"
      Height          =   2415
      Left            =   4755
      TabIndex        =   13
      Top             =   100
      Width           =   5595
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdConceptos 
         Height          =   1935
         Left            =   100
         TabIndex        =   8
         ToolTipText     =   "Otros conceptos de facturación"
         Top             =   240
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   3413
         _Version        =   393216
         GridLinesFixed  =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame fraBuscar 
      Height          =   2415
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   4575
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   1930
         TabIndex        =   7
         ToolTipText     =   "Buscar socios"
         Top             =   1920
         Width           =   735
      End
      Begin VB.OptionButton optTipoMembresia 
         Caption         =   "Individual"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         ToolTipText     =   "Membresía"
         Top             =   1635
         Width           =   1095
      End
      Begin VB.OptionButton optTipoMembresia 
         Caption         =   "Familiar"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         ToolTipText     =   "Membresía"
         Top             =   1635
         Width           =   855
      End
      Begin VB.OptionButton optTipoMembresia 
         Caption         =   "Ambos"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Membresía"
         Top             =   1635
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtNumeroCredencial 
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Número de credencial"
         Top             =   1140
         Width           =   2535
      End
      Begin VB.TextBox txtClave 
         Height          =   285
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Clave del socio"
         Top             =   810
         Width           =   2535
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Nombre completo"
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lbTipoMembresia 
         Caption         =   "Tipo de membresía"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label lbNumeroCredenciall 
         Caption         =   "Número de credencial"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1185
         Width           =   1815
      End
      Begin VB.Label lbClave 
         Caption         =   "Clave única"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   825
         Width           =   1455
      End
      Begin VB.Label lbNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmCargosSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vllngNumeroOpcion As Long
Dim vllngPersonaGraba As Long
Const cintColMonto = 4
Const cstrCantidad = "###########.00"

Private Sub ActDesactBotonAplicarCargos()

Dim blnAplicar As Boolean

On Error GoTo NotificaError

Dim i As Integer

    With grdConceptos
        
        If .Row > 0 Then
            
            For i = 1 To .Rows - 1
                
                If .TextMatrix(i, 1) <> "" Then
                               
                    blnAplicar = True
                
                End If
                                
            Next i
        
        End If
    
    End With
    
    With grdSocios
        
        If .Row > 0 And blnAplicar Then
        
            blnAplicar = False
            
            For i = 1 To .Rows - 1
                
                If .TextMatrix(i, 1) <> "" Then
                               
                    blnAplicar = True
                
                End If
                                
            Next i
        
        End If
    
    End With
    
    If blnAplicar Then
        
        cmdAplicarCargos.Enabled = True
    
    Else
    
        cmdAplicarCargos.Enabled = False
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":ActDesactBotonAplicarCargos"))
    
End Sub

Private Sub pLimpiaBuscar()

On Error GoTo NotificaError

    txtNombre.Text = ""
    txtClave.Text = ""
    txtNumeroCredencial.Text = ""
    optTipoMembresia(0).Value = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaBuscar"))
        
End Sub

Private Sub pLimpiaGrid(grid)

On Error GoTo NotificaError

    grid.Clear
    grid.Rows = 2
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
    
End Sub

Private Sub cmdAplicarCargos_Click()

Dim lngResultado As Long
Dim lngPrecioOut As Long
Dim strParametrosSP As String
Dim i As Integer
Dim j As Integer
Dim blnCargo As Boolean
Dim vllngNumeroCorte As Long
Dim vllngCorteGrabando  As Long
Dim intNumeroCuenta As Long
Dim rsMonto As ADODB.Recordset
Dim rsPrecio As ADODB.Recordset
Dim rsLista As ADODB.Recordset
Dim dblMonto As Double
Dim aFormasPago() As FormasPago
Dim intContador As Integer
Dim blnFormaPago As Boolean
Dim rsFormaPago As ADODB.Recordset
Dim vlstrsql As String

On Error GoTo NotificaError

    blnCargo = False
    
    If Not fblnRevisaPermiso(vglngNumeroLogin, 2416, "E") Then Exit Sub
    vllngPersonaGraba = 0
    
        If vllngPersonaGraba = 0 Then
          vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        End If
        
        strParametrosSP = -1 & "|" & -1 & "|" & -1 & "|" & CStr(vgintNumeroDepartamento) & "|" & 1 & "|" & "C"
                                        
        Set rsFormaPago = frsEjecuta_SP(strParametrosSP, "sp_PvSelFormaPago")
                        
        If rsFormaPago.RecordCount = 0 Then
        
            'No existen formas de pago
            
            MsgBox Replace(SIHOMsg(293), ".", "") & " a crédito para este departamento", vbExclamation + vbOKOnly, "Mensaje"
                                
            Exit Sub
                        
        End If
        
                
            If vllngPersonaGraba <> 0 Then
                                
                With EntornoSIHO
                
                    If grdSocios.RowData(1) > 0 Then
                
                        .ConeccionSIHO.BeginTrans
                        
                        For i = 1 To grdSocios.Rows - 1
                        
                            If grdSocios.TextMatrix(i, 1) <> "" Then
                            
                                For j = 1 To grdConceptos.Rows - 1
                        
                                        If grdConceptos.TextMatrix(j, 1) <> "" Then
                                                                        
                                        lngResultado = 1
                                        
                                        'Graba en PVCARGO el cargo
                                        strParametrosSP = grdConceptos.RowData(j) & "|" & vgintNumeroDepartamento & "|" & "D" & "|" & 0 & "|" & grdSocios.RowData(i) & "|" & "S" & "|" & "OC" & "|" & 0 & "|" & 1 & "|" & vllngPersonaGraba & "|" & 0 & "|" & "" & "|" & 0 & "|" & 2
                                        frsEjecuta_SP strParametrosSP, "SP_PVUPDCARGOS", True, lngResultado
                                        
                                        blnCargo = True
                                        
                                        'Consulta el precio correcto de la lista de precios (Lista predeterminada del departamento)
                                        Set rsPrecio = frsEjecuta_SP(vgintNumeroDepartamento & "|" & grdConceptos.RowData(j) & "|" & "OC", "SP_PVSELLISTASOCIOS")
                                                            
                                        'Selecciona la lista predeterminada del departamento
                                        vlstrsql = "SELECT intcvelista as Lista " & _
                                                         "From PVLISTAPRECIO " & _
                                                            "INNER JOIN NODEPARTAMENTO " & _
                                                            "ON PVLISTAPRECIO.smidepartamento = NODEPARTAMENTO.smicvedepartamento " & _
                                                        "Where bitpredeterminada = 1 " & _
                                                            "AND smidepartamento = " & vgintNumeroDepartamento
                                                            
                                        Set rsLista = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
                                                            
                                        'Se valida que haya una lista predeterminada para el departamento
                                        If rsLista.RecordCount < 1 Then
                                            .ConeccionSIHO.RollbackTrans
                                            MsgBox SIHOMsg(1110), vbCritical, "Mensaje"
                                            rsLista.Close
                                            Exit Sub

                                        Else
                                        
                                        'Se valida la información de la lista predeterminada
                                        If rsPrecio.RecordCount < 1 Then
                                            .ConeccionSIHO.RollbackTrans
                                            MsgBox SIHOMsg(301), vbExclamation, "Mensaje"
                                            rsPrecio.Close
                                            Exit Sub
                                            
                                        Else
                                        
                                            'Se valida que la lista predeterminada esté activa
                                            If CInt(rsPrecio!Activo) = 0 Then
                                                .ConeccionSIHO.RollbackTrans
                                                MsgBox SIHOMsg(1111), vbCritical, "Mensaje"
                                                rsPrecio.Close
                                                Exit Sub
                                            End If
                                        
                                            'Se valida que tenga un precio capturado en la lista predeterminada del departamento
                                            If CDec(rsPrecio!Precio) <= 0 Then
                                                .ConeccionSIHO.RollbackTrans
                                                MsgBox SIHOMsg(301), vbExclamation, "Mensaje"
                                                rsPrecio.Close
                                                Exit Sub
                                            Else
                                            
                                                If lngResultado < 0 Then
                                    
                                                    .ConeccionSIHO.RollbackTrans
                                                    MsgBox SIHOMsg(lngResultado * -1), vbExclamation, "Mensaje"
                                                    
                                                    Exit Sub
                                                            
                                                End If
                                            
                                                'Modifica el precio del cargo con el obtenido antes
                                                frsEjecuta_SP CStr(lngResultado) & "|" & CDec(rsPrecio!Precio) & "|" & 0, "SP_PVUPDCARGOSOCIOS", False
                                                rsPrecio.Close
                                                
                                                'Regresa el nuevo precio del cargo
                                                Set rsMonto = frsEjecuta_SP(CStr(lngResultado), "SP_PVSELPRECIOCARGOSOCIOS")
                                                                                    
                                                If rsMonto.RecordCount > 0 Then dblMonto = rsMonto!Monto
                                                
                                                rsMonto.Close
                                                
                                            End If
                                        End If
                                    End If
                                        
                                        '--------------------------------------------------------
                                        ' Estatus de "GRABANDO" en el Corte
                                        '--------------------------------------------------------
                                        vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                                        
                                        vllngCorteGrabando = 1
                                        vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                                        frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                                        If vllngCorteGrabando <> 2 Then
                                            EntornoSIHO.ConeccionSIHO.RollbackTrans
                                            'No se puede realizar la operación, inténtelo en unos minutos.
                                            MsgBox SIHOMsg(720), vbExclamation + vbOKOnly, "Mensaje"
                                            Exit Sub
                                        End If

                                        strParametrosSP = CStr(vllngNumeroCorte) _
                                        & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                                        & "|" & CStr(lngResultado) _
                                        & "|" & "SO" _
                                        & "|" & CStr(rsFormaPago!intFormaPago) _
                                        & "|" & CStr(dblMonto) _
                                        & "|" & CStr(0) _
                                        & "|" & CStr(0) _
                                        & "|" & CStr(vllngNumeroCorte)
                                        
                                                                                                                        
                                        frsEjecuta_SP strParametrosSP, "sp_PvInsDetalleCorte"
                                        
                                        
                                        strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                                        
                                        intNumeroCuenta = 1
                                        ' Función que regresa el número de cuenta contable de la cuenta puente de cuentas por cobrar socios
                                        frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
            
                                        pInsCortePoliza vllngNumeroCorte, grdSocios.TextMatrix(i, 2), "SO", intNumeroCuenta, dblMonto, 1 'Cargo a la cuenta puente de cuentas por cobrar socios
                                        
                                        strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                                        
                                        intNumeroCuenta = 1
                                        ' Función que regresa el número de cuenta contable de la cuenta puente de cuotas por devengar socios
                                        frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                                        
                                        pInsCortePoliza vllngNumeroCorte, grdSocios.TextMatrix(i, 2), "SO", intNumeroCuenta, dblMonto, 0 'Abono a la cuenta puente de cuotas por devengar socios
                                        
                                        pLiberaCorte (vllngNumeroCorte)
                                        
                                                                        
                                    End If
                                
                                Next j
                                
                            End If
                            
                        Next i
                            
                        If blnCargo = True Then
                        
                            .ConeccionSIHO.CommitTrans
                            MsgBox SIHOMsg(316), vbInformation, "Mensaje"
                        
                        Else
                        
                            .ConeccionSIHO.RollbackTrans
                            
                        End If
                        
                    End If
                                        
                End With
                                                
                pLlenaCargos
                
                Call pDesmarcaGrid(grdConceptos)
                               
            End If
            
            rsFormaPago.Close
            
'        Else
'            MsgBox SIHOMsg(65), vbExclamation, "Mensaje"
'        End If
                    
Exit Sub
NotificaError:
    
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdAplicarCargos_Click"))
        
End Sub

Private Sub cmdBuscar_Click()

On Error GoTo NotificaError

    pCargaSocios
    pLlenaCargos
    ActDesactBotonAplicarCargos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBuscar_Click"))
        
End Sub

Private Sub cmdBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo NotificaError
    
    If KeyCode = 13 Then cmdBuscar_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdInvertirSeleccion_Click"))
        
End Sub

Private Sub cmdInvertirSeleccion_Click()

Dim i As Long


On Error GoTo NotificaError

    For i = 1 To grdSocios.Rows - 1
    
        With grdSocios
        
            If .TextMatrix(i, 1) = "" Then
            
                .TextMatrix(i, 1) = "*"
            
            Else
            
                .TextMatrix(i, 1) = ""
            
            End If
                
        End With
        
    Next i
    
    ActDesactBotonAplicarCargos
    pLlenaCargos

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdInvertirSeleccion_Click"))

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo NotificaError

    Select Case KeyCode
        
        Case vbKeyEscape
        
            Unload Me
                        
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Form_Load()

On Error GoTo NotificaError


    Me.Icon = frmMenuPrincipal.Icon
        
    pConfiguraGridConceptos
    pConfiguraGridSocios
    pConfiguraGridCargos
    
    pCargaOtrosConceptos
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub pConfiguraGridSocios()
    
On Error GoTo NotificaError

    With grdSocios
        
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||Clave|Nombre del socio|Número de credencial|Tipo de membresía"
        .ColWidth(0) = 0  'Fix
        .ColWidth(1) = 250  'Selección
        .ColWidth(2) = 1600 'Clave
        .ColWidth(3) = 4720  'Nombre del socio
        .ColWidth(4) = 1700 'Número de credencial
        .ColWidth(5) = 1650 'Tipo de membresía
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
                
        .ScrollBars = flexScrollBarVertical
                
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridSocios"))
    
End Sub

Private Sub pConfiguraGridConceptos()

On Error GoTo NotificaError

    With grdConceptos
        
        .Cols = 3
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "||Otro concepto"
        .ColWidth(0) = 0  'Fix
        .ColWidth(1) = 250 'Selección
        .ColWidth(2) = 5050 'Descripción del cargo
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        
        .ScrollBars = flexScrollBarVertical
    
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridConceptos"))
    
End Sub

Private Sub pCargaOtrosConceptos()
    
On Error GoTo NotificaError

    Dim rsSelOtrosConceptos As ADODB.Recordset
        
        
    '------------------------
    ' Limpia el grid
    '------------------------
    Call pLimpiaGrid(grdConceptos)
    
    '------------------------
    ' Configurar el grid
    '------------------------
    pConfiguraGridConceptos
    grdConceptos.RowData(1) = -1
    
    
    vgstrParametrosSP = -1 & "|" & CStr(vgintNumeroDepartamento)
        
    ' SP que regresa otros conceptos del departamento
    
    Set rsSelOtrosConceptos = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELOTROSCONCEPTOSDEPTO")
    
    If rsSelOtrosConceptos.RecordCount > 0 Then
        
        With grdConceptos
                           
            Do While Not rsSelOtrosConceptos.EOF
            
                If .RowData(1) <> -1 Then
                     .Rows = .Rows + 1
                     .Row = .Rows - 1
                End If
                
                .RowData(.Row) = rsSelOtrosConceptos!Clave
                .TextMatrix(.Row, 2) = rsSelOtrosConceptos!Descripcion
                .Col = 1
                .CellFontBold = True
                    
                rsSelOtrosConceptos.MoveNext
                    
            Loop
            
        End With
        
        rsSelOtrosConceptos.Close
    
        If grdConceptos.RowData(1) = -1 Then
            grdConceptos.Enabled = False
        Else
            grdConceptos.Enabled = True
        End If
    
    Else
    
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        
    
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaOtrosConceptos"))
    
End Sub

Private Sub pCargaSocios()
    
On Error GoTo NotificaError

    Dim rsSelSocios As ADODB.Recordset
    Dim intTipoMembresia As Integer
    
        
    '------------------------
    ' Limpia el grid
    '------------------------
    Call pLimpiaGrid(grdSocios)
    
    '------------------------
    ' Configurar el grid
    '------------------------
    pConfiguraGridSocios
    grdSocios.RowData(1) = -1
       
       
    If optTipoMembresia(0) = True Then
        
        intTipoMembresia = -1
    
    ElseIf optTipoMembresia(1) = True Then
    
            intTipoMembresia = 1
        
        Else
        
            intTipoMembresia = 0
            
    End If
        
    vgstrParametrosSP = IIf(txtNombre.Text = "", "*", txtNombre.Text) & "|" & IIf(txtClave.Text = "", "*", txtClave.Text) & "|" & IIf(txtNumeroCredencial.Text = "", "0", txtNumeroCredencial.Text) & _
    "|" & CStr(intTipoMembresia)
        
    ' SP que regresa los socios
    
    Set rsSelSocios = frsEjecuta_SP(vgstrParametrosSP, "SP_SOSELSOCIOS")
    
    If rsSelSocios.RecordCount > 0 Then
        
        With grdSocios
                           
            Do While Not rsSelSocios.EOF
            
                If .RowData(1) <> -1 Then
                     .Rows = .Rows + 1
                     .Row = .Rows - 1
                End If
                
                .RowData(.Row) = rsSelSocios!intcvesocio
                .TextMatrix(.Row, 1) = "*"
                .TextMatrix(.Row, 2) = rsSelSocios!VCHCLAVESOCIO
                .TextMatrix(.Row, 3) = Trim(rsSelSocios!vchApellidoPaterno) & " " & Trim(rsSelSocios!vchApellidoMaterno) & " " & Trim(rsSelSocios!vchNombre)
                .TextMatrix(.Row, 4) = rsSelSocios!intnumerocredencial
                .TextMatrix(.Row, 5) = IIf(rsSelSocios!intCuota = 0, "INDIVIDUAL", "FAMILIAR")
                .Col = 1
                .CellFontBold = True
                    
                rsSelSocios.MoveNext
                    
            Loop
            
        End With
        
        rsSelSocios.Close
    
        If grdSocios.RowData(1) = -1 Then
            grdSocios.Enabled = False
        Else
            grdSocios.Enabled = True
        End If
        
        cmdInvertirSeleccion.Enabled = True
    
    Else
    
        MsgBox SIHOMsg(13), vbOKOnly + vbInformation, "Mensaje"
        
        
        cmdInvertirSeleccion.Enabled = False
        
        txtNombre.SetFocus
            
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaSocios"))
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo NotificaError
    
                If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            
                        If grdSocios.RowData(1) > 0 Then
                            
                            pLimpiaBuscar
                            Call pLimpiaGrid(grdSocios)
                            Call pLimpiaGrid(grdCargos)
                            grdSocios.RowData(1) = 0
                            cmdInvertirSeleccion.Enabled = False
                            cmdAplicarCargos.Enabled = False
                            txtNombre.SetFocus
                            
                            Cancel = 1
                            
                        Else
                        
                            vllngPersonaGraba = 0
                            Unload Me
                            
                        End If
                    
                Else
                
                        Cancel = 1
                
                End If
                
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_QueryUnload"))
                
End Sub

Private Sub grdCargos_Click()

On Error GoTo NotificaError

    With grdCargos
    
        If .TextMatrix(.Row, 1) <> "" Then
            
            If .Col = cintColMonto Then
            
                txtMonto.Visible = True
            
                .RowHeight(.Row) = txtMonto.Height
                txtMonto.Width = .CellWidth
            
                txtMonto.Top = .CellTop + 230
                txtMonto.Left = .CellLeft + 90
                txtMonto.Text = .Text
                txtMonto.Text = Replace(.Text, "$", "")
                txtMonto.SetFocus
                pSelTextBox txtMonto
                
            End If
    
        End If
    
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_Click"))
    
End Sub

Private Sub GrdCargos_DblClick()
    
Dim rsDetalleCorte As ADODB.Recordset
Dim strParametrosSP As String
Dim vllngNumeroCorte As Long
Dim vllngCorteGrabando  As Long
Dim intNumeroCuenta As Long


On Error GoTo NotificaError


    If grdCargos.TextMatrix(grdCargos.Row, 1) = "" Then Exit Sub
    If Not fblnRevisaPermiso(vglngNumeroLogin, 2416, "E") Then Exit Sub
    If MsgBox(Mid(SIHOMsg(327), 1, 40) & "?", vbYesNo + vbQuestion, "Mensaje") = vbYes Then
           
                    
            strParametrosSP = grdCargos.RowData(grdCargos.Row) & "|" & "SO"

            Set rsDetalleCorte = frsEjecuta_SP(strParametrosSP, "SP_PVSELDETALLECORTE")

            If rsDetalleCorte.RecordCount > 0 Then
            
                With EntornoSIHO
            
                    .ConeccionSIHO.BeginTrans
                            
            
                    '--------------------------------------------------------
                    ' Estatus de "GRABANDO" en el Corte
                    '--------------------------------------------------------
                    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                    
                    vllngCorteGrabando = 1
                    vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                    frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                    If vllngCorteGrabando <> 2 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        'No se puede realizar la operación, inténtelo en unos minutos.
                        MsgBox SIHOMsg(720), vbExclamation + vbOKOnly, "Mensaje"
                        Exit Sub
                    End If
                    
                    strParametrosSP = CStr(vllngNumeroCorte) _
                    & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                    & "|" & rsDetalleCorte!CHRFOLIODOCUMENTO _
                    & "|" & "SO" _
                    & "|" & CStr(rsDetalleCorte!intFormaPago) _
                    & "|" & CStr(rsDetalleCorte!MNYCANTIDADPAGADA * -1) _
                    & "|" & CStr(rsDetalleCorte!MNYTIPOCAMBIO) _
                    & "|" & CStr(0) _
                    & "|" & CStr(vllngNumeroCorte)

                    frsEjecuta_SP strParametrosSP, "sp_PvInsDetalleCorte"
                                     
                    strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuentas por cobrar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
            
                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, grdCargos.TextMatrix(grdCargos.Row, 5), 0  'Abono a la cuenta puente de cuentas por cobrar socios
                    
                    strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuotas por devengar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                    
                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, grdCargos.TextMatrix(grdCargos.Row, 5), 1  'Cargo a la cuenta puente de cuotas por devengar socios
                    
                    pLiberaCorte (vllngNumeroCorte)
                    
                    
                    frsEjecuta_SP CStr(grdCargos.RowData(grdCargos.Row)), "SP_PVDELCARGO"
                    
                    
                    If grdCargos.Rows = 2 Then
                        
                        Call pLimpiaGrid(grdCargos)
                        
                    Else
                    
                        grdCargos.RemoveItem (grdCargos.Row)
                    
                    End If
                                       
                    .ConeccionSIHO.CommitTrans
                    
                End With
                
            End If

                        
            rsDetalleCorte.Close
            
            pLlenaCargos
        
    
    End If
   

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_DblClick"))
    

End Sub

Private Sub grdCargos_LeaveCell()
    
On Error GoTo NotificaError

    grdCargos.RowHeight(grdCargos.Row) = 240
    
    txtMonto.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCargos_LeaveCell"))
    
End Sub

Private Sub grdConceptos_Click()
        
On Error GoTo NotificaError


    With grdConceptos
        
        If .Row > 0 Then
            
            If .Col = 1 Then
                
                If .TextMatrix(.Row, 1) = "" Then
                               
                    .TextMatrix(.Row, 1) = "*"
                                    
                Else
                
                    .TextMatrix(.Row, 1) = ""
                
                End If
                                
            End If
        
        End If
    
    End With
    
    ActDesactBotonAplicarCargos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConceptos_Click"))

End Sub

Private Sub grdConceptos_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo NotificaError

    If KeyCode = 13 Then grdConceptos_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdConceptos_KeyDown"))
    
End Sub

Private Sub pConfiguraGridCargos()
    
On Error GoTo NotificaError

    With grdCargos
        
        .Cols = 6
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Clave|Nombre del socio|Otro concepto|Monto"
        .ColWidth(0) = 100  'Fix
        .ColWidth(1) = 1550 'Clave
        .ColWidth(2) = 3900  'Nombre del socio
        .ColWidth(3) = 3150 'Otro concepto
        .ColWidth(4) = 980 'Monto
        .ColWidth(5) = 0 'Monto base
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftBottom
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
                
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
                
        .ScrollBars = flexScrollBarVertical
            
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridCargos"))
    
End Sub

Private Sub grdSocios_Click()
    
On Error GoTo NotificaError

    With grdSocios
        
        If .TextMatrix(.Row, 2) <> "" Then
            
            If .Col = 1 Then
                
                If .TextMatrix(.Row, 1) = "" Then
                               
                    .TextMatrix(.Row, 1) = "*"
                                    
                Else
                
                    .TextMatrix(.Row, 1) = ""
                
                End If
                
                pLlenaCargos
                                
            End If
        
        End If
    
    End With
    
    ActDesactBotonAplicarCargos
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdSocios_Click"))
    
End Sub

Private Sub grdSocios_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error GoTo NotificaError

    If KeyCode = 13 Then grdSocios_Click
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdSocios_KeyDown"))
        
End Sub



Private Sub txtClave_KeyPress(KeyAscii As Integer)
    
On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtClave_KeyPress"))
        
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)

Dim dblMonto As Double
Dim vllngNumPoliza As Long
Dim vllngDetallePoliza As Long
Dim intNumeroCuenta As Long
Dim vlstrFormato As String
Dim strParametrosSP As String
Dim vllngCorteGrabando  As Long
Dim intContador As Integer
Dim rsDetalleCorte As ADODB.Recordset


vlstrFormato = "##########.00"


On Error GoTo NotificaError
    
    If KeyCode = 13 Then
    
        'Valida que el precio capturado sea mayor que 0
        If Val(txtMonto.Text) <= 0 Then
            'No se puede realizar la operación, inténtelo en unos minutos.
            MsgBox SIHOMsg(788), vbExclamation + vbOKOnly, "Mensaje"
            
            With grdCargos
                txtMonto.Text = .Text
                txtMonto.Text = Replace(.Text, "$", "")
                pSelTextBox txtMonto
'                txtMonto.Visible = False
            End With
            
            Exit Sub
        End If
    
        With grdCargos
    
            dblMonto = CDbl(.TextMatrix(.Row, cintColMonto))
            
            
            .Text = FormatCurrency(Val(Format(txtMonto.Text, "##########.00")), 2)
            .RowHeight(.Row) = 240
            txtMonto.Visible = False
            
            If Val(txtMonto.Text) <> .TextMatrix(.Row, 5) Then
            
                If vllngPersonaGraba = 0 Then
                    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                End If
            
                 strParametrosSP = -1 & "|" & -1 & "|" & -1 & "|" & CStr(vgintNumeroDepartamento) & "|" & 1 & "|" & "C"
                                        
                Set rsFormaPago = frsEjecuta_SP(strParametrosSP, "sp_PvSelFormaPago")
                
                
                If rsFormaPago.RecordCount = 0 Then
                
                    'No existen formas de pago
                    
                    MsgBox Replace(SIHOMsg(293), ".", "") & " a crédito para este departamento", vbExclamation + vbOKOnly, "Mensaje"
                                        
                    Exit Sub
                                
                End If
            
                strParametrosSP = .RowData(.Row) & "|" & "SO"

                Set rsDetalleCorte = frsEjecuta_SP(strParametrosSP, "SP_PVSELDETALLECORTE")
                
            
                If vllngPersonaGraba <> 0 And rsDetalleCorte.RecordCount > 0 Then
                                       
                    strParametrosSP = .RowData(.Row) & "|" & CStr(Val(Format(txtMonto.Text, "##########.00"))) & "|" & 1
                    
                    frsEjecuta_SP strParametrosSP, "SP_PVUPDCARGOSOCIOS"
                                       
                                
                    '--------------------------------------------------------
                    ' Estatus de "GRABANDO" en el Corte
                    '--------------------------------------------------------
                    vllngNumeroCorte = flngNumeroCorte(vgintNumeroDepartamento, vglngNumeroEmpleado, "P")
                    
                    vllngCorteGrabando = 1
                    vgstrParametrosSP = vllngNumeroCorte & "|" & "Grabando"
                    frsEjecuta_SP vgstrParametrosSP, "sp_PvUpdEstatusCorte", True, vllngCorteGrabando
                    If vllngCorteGrabando <> 2 Then
                        EntornoSIHO.ConeccionSIHO.RollbackTrans
                        'No se puede realizar la operación, inténtelo en unos minutos.
                        MsgBox SIHOMsg(720), vbExclamation + vbOKOnly, "Mensaje"
                        Exit Sub
                    End If
                    
                    
                    ' Movimiento del monto anterior de los cargos
                    
                    strParametrosSP = CStr(vllngNumeroCorte) _
                    & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                    & "|" & rsDetalleCorte!CHRFOLIODOCUMENTO _
                    & "|" & "SO" _
                    & "|" & CStr(rsDetalleCorte!intFormaPago) _
                    & "|" & CStr(rsDetalleCorte!MNYCANTIDADPAGADA * -1) _
                    & "|" & CStr(rsDetalleCorte!MNYTIPOCAMBIO) _
                    & "|" & CStr(0) _
                    & "|" & CStr(vllngNumeroCorte)

                    frsEjecuta_SP strParametrosSP, "sp_PvInsDetalleCorte"
                    
                    
                    ' Movimiento del monto nuevo de los cargos
                    
                    strParametrosSP = CStr(vllngNumeroCorte) _
                    & "|" & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) _
                    & "|" & rsDetalleCorte!CHRFOLIODOCUMENTO _
                    & "|" & "SO" _
                    & "|" & CStr(rsDetalleCorte!intFormaPago) _
                    & "|" & CStr(Val(txtMonto.Text)) _
                    & "|" & CStr(rsDetalleCorte!MNYTIPOCAMBIO) _
                    & "|" & CStr(0) _
                    & "|" & CStr(vllngNumeroCorte)
                                                                                                                        
                    frsEjecuta_SP strParametrosSP, "sp_PvInsDetalleCorte"
                    
                    
                    strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuentas por cobrar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
            
                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, grdCargos.TextMatrix(grdCargos.Row, 5), 0  'Abono a la cuenta puente de cuentas por cobrar socios
                    
                    strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuotas por devengar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
                    
                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, grdCargos.TextMatrix(grdCargos.Row, 5), 1  'Cargo a la cuenta puente de cuotas por devengar socios
                    
                    
                    ' Movimiento para el nuevo monto de los cargos
                    
'                    vllngNumPoliza = flngInsertarPoliza(FormatDateTime(Now, vbShortDate), "D", "CARGOS SOCIO " & Trim(grdCargos.TextMatrix(grdCargos.Row, 1)), vllngPersonaGraba)
'
'                    strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
'
'                    intNumeroCuenta = 1
'
'                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuentas por cobrar socios
'                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
'
'                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(txtMonto.Text, vlstrFormato)), 1)
                    
'                    strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
'
'                    intNumeroCuenta = 1
'
'                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuotas por devengar socios
'                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta
'
'                    vllngDetallePoliza = flngInsertarPolizaDetalle(vllngNumPoliza, intNumeroCuenta, Val(Format(txtMonto.Text, vlstrFormato)), 0)
                    
                    
                    ' Movimiento para el nuevo monto de los cargos
                    
                    strParametrosSP = "INTNUMCUENTACUOTASOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuentas por cobrar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta

                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, CStr(Val(txtMonto.Text)), 1  'Cargo a la cuenta puente de cuentas por cobrar socios
                    
                    strParametrosSP = "INTNUMCUENTADEVENGARSOCIOS" & "|" & CStr(vgintClaveEmpresaContable)
                    
                    intNumeroCuenta = 1
                    ' Función que regresa el número de cuenta contable de la cuenta puente de cuotas por devengar socios
                    frsEjecuta_SP strParametrosSP, "FN_CCSELNUMEROCUENTAPUENTE", True, intNumeroCuenta

                    pInsCortePoliza vllngNumeroCorte, CStr(grdCargos.TextMatrix(grdCargos.Row, 1)), "SO", intNumeroCuenta, CStr(Val(txtMonto.Text)), 0  'Abono a la cuenta puente de cuotas por devengar socios
                                                            
                    pLiberaCorte (vllngNumeroCorte)
                                        
                                    
                End If
                            
                rsDetalleCorte.Close
                
                pLlenaCargos
                
            End If
            
        End With
    
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMonto_KeyDown"))
    
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    
On Error GoTo NotificaError

    If Not fblnFormatoCantidad(txtMonto, KeyAscii, 2) Then
    
            KeyAscii = 7
            
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMonto_KeyPress"))
    
End Sub

Private Sub txtMonto_LostFocus()
       
   On Error GoTo NotificaError
   
        txtMonto.Visible = False
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtMonto_LostFocus"))
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
On Error GoTo NotificaError
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Then
            
            KeyAscii = 7
    
    End If
    

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNombre_KeyPress"))
        
End Sub

Private Sub txtNumeroCredencial_KeyPress(KeyAscii As Integer)
    
On Error GoTo NotificaError

    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Or KeyAscii = 46 Then
            
            KeyAscii = 7
    
    End If
    
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtNumeroCredencial_KeyPress"))
        
End Sub

Private Sub pLlenaCargos()

On Error GoTo NotificaError

    Dim rsSelCargos As ADODB.Recordset
    Dim strParametrosSP As String
    Dim i As Integer
    Dim intContador As Integer
    
        
    '------------------------
    ' Limpia el grid
    '------------------------
    Call pLimpiaGrid(grdCargos)
    
    '------------------------
    ' Configurar el grid
    '------------------------
    pConfiguraGridCargos
    grdCargos.RowData(1) = -1
    
    
    With grdSocios
    
        If .RowData(1) > 0 Then
        
            For i = 1 To .Rows - 1
                
                If .TextMatrix(i, 1) <> "" Then
            
                    strParametrosSP = CStr(.RowData(i))
                        
                    ' SP que regresa los cargos de los socios
                    
                    Set rsSelCargos = frsEjecuta_SP(strParametrosSP, "SP_PVSELCARGOSDESOCIOS")
                    
                    If rsSelCargos.RecordCount > 0 Then
                        
                        With grdCargos
                                           
                            Do While Not rsSelCargos.EOF
                            
                                If .RowData(1) <> -1 Then
                                     .Rows = .Rows + 1
                                     .Row = .Rows - 1
                                End If
                                
                                .RowData(.Row) = rsSelCargos!ClaveCargo
                                .TextMatrix(.Row, 1) = rsSelCargos!Clave
                                .TextMatrix(.Row, 2) = rsSelCargos!Nombre
                                .TextMatrix(.Row, 3) = rsSelCargos!Concepto
                                .TextMatrix(.Row, 4) = FormatCurrency(rsSelCargos!Monto, 2)
                                .TextMatrix(.Row, 5) = rsSelCargos!Monto
                                
                                If rsSelCargos!PrecioManual = 1 Then
                                
                                    For intContador = 1 To grdCargos.Cols - 1
                                        grdCargos.Col = intContador
                                        grdCargos.CellBackColor = &H80000018
                                    Next
                                
                                End If
                                
                                                    
                                rsSelCargos.MoveNext
                                    
                            Loop
                            
                        End With
                        
                        rsSelCargos.Close
                    
                    End If
                        
                End If
                    
            Next i
        
        End If
        
    
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCargos"))
        
End Sub

Private Sub pDesmarcaGrid(grid As MSHFlexGrid)

Dim i As Integer


On Error GoTo NotificaError

    With grid
    
        For i = 1 To .Rows - 1
        
            .TextMatrix(i, 1) = ""
                    
        Next i
        
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pDesmarcaGrid"))
    

End Sub

