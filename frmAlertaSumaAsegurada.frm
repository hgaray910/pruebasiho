VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAlertaSumaAsegurada 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas de pacientes que exceden la suma asegurada"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   14070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCuentas 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   -2147483633
      FocusRect       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraBotonera 
      Height          =   765
      Left            =   4790
      TabIndex        =   4
      Top             =   4680
      Width           =   4490
      Begin VB.CommandButton cmdCerrarCuentas 
         Caption         =   "Cerrar cuentas"
         Height          =   495
         Left            =   1520
         TabIndex        =   2
         ToolTipText     =   "Cerrar las cuentas seleccionadas"
         Top             =   170
         Width           =   1440
      End
      Begin VB.CommandButton cmdAbrirCuentas 
         Caption         =   "Abrir cuentas"
         Height          =   495
         Left            =   2970
         TabIndex        =   3
         ToolTipText     =   "Abrir las cuentas seleccionadas"
         Top             =   170
         Width           =   1440
      End
      Begin VB.CommandButton cmdInvertir 
         Caption         =   "Invertir selección"
         Height          =   495
         Left            =   70
         TabIndex        =   1
         ToolTipText     =   "Invertir la selección"
         Top             =   170
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmAlertaSumaAsegurada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Facturas previas
'----------------------------------------------------------------------------------
Option Explicit

Public vlintmovpaciente As Long
Public vlchrtipopaciente As String
Public vlchrtipofactura As String

Public vlblnPermiteCerrar As Boolean
Public vlblnPermiteAbrir As Boolean

Private Sub cmdAbrirCuentas_Click()
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    Dim intDias As Integer
    Dim blnPermisoEscritura As Boolean
    Dim blnPermisoCTotal As Boolean
    Dim vllngVar As Long
    Dim vldtmFechaAntes As Date
    Dim vldtmFechaDespues As Date
    Dim rsCargos As New ADODB.Recordset
    Dim vlmsgCargosProgramados As String
        
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
    Dim vlblnSeleccionados As Boolean
    Dim vllngCuentaCerrada As Integer
    Dim rs As ADODB.Recordset
    Dim vlintNumDiasAbrirExt As Integer
    Dim vlintNumDiasAbrirInt As Integer
    
    Set rs = frsRegresaRs("select intDiasAbrirCuentasInternos, intDiasAbrirCuentasExternos from PVParametro where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        vlintNumDiasAbrirExt = IIf(IsNull(rs!intDiasAbrirCuentasExternos), 0, rs!intDiasAbrirCuentasExternos)
        vlintNumDiasAbrirInt = IIf(IsNull(rs!intDiasAbrirCuentasInternos), 0, rs!intDiasAbrirCuentasInternos)
    Else
        vlintNumDiasAbrirExt = 0
        vlintNumDiasAbrirInt = 0
    End If
        
    vllngCuentaCerrada = 1
     
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba <> 0 Then
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        With grdCuentas
            For vllngContador = 1 To .Rows - 1
                If .TextMatrix(vllngContador, 0) = "*" And .TextMatrix(vllngContador, 3) = "Cerrada" Then
                                                                
                    'vllngCuentaCerrada  1 Cuenta cerrada 0 Cuenta abierta
                    vgstrParametrosSP = .TextMatrix(vllngContador, 1) & "|" & _
                                        .TextMatrix(vllngContador, 10) & "|" & _
                                        IIf(vllngCuentaCerrada = 1, "0", "1") & "|" & _
                                        IIf(vllngCuentaCerrada = 0 And .TextMatrix(vllngContador, 10) = "E", fstrFechaSQL(fdtmServerFecha, fdtmServerHora), Null)
                      
                    frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDCERRARABRIRCUENTA"
                    Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, IIf(vllngCuentaCerrada = 1, "APERTURA DE CUENTA", "CIERRE DE CUENTA"), .TextMatrix(vllngContador, 1))
                
                    .TextMatrix(vllngContador, 3) = "Abierta"
                
                    If vllngCuentaCerrada = 1 Then
                        pEjecutaSentencia "delete PVCuentasReabiertas where numNumCuenta = " & .TextMatrix(vllngContador, 1) & " and chrTipoPaciente = '" & .TextMatrix(vllngContador, 10) & "'"
                        pEjecutaSentencia "insert into PVCuentasReabiertas (numNumCuenta, chrTipoPaciente, intCveEmpleado, intDiasReabrirE, intDiasReabierta) values (" & .TextMatrix(vllngContador, 1) & ", '" & .TextMatrix(vllngContador, 10) & "', " & vllngPersonaGraba & ", " & IIf(.TextMatrix(vllngContador, 10) = "I", vlintNumDiasAbrirInt, vlintNumDiasAbrirExt) & ", " & intDias & ")"
                    End If
                End If
            Next vllngContador
        End With
                                                                
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'La información se actualizó satisfactoriamente.
        MsgBox SIHOMsg(284), vbInformation, "Mensaje"
        
        pLimpia
        pHabilitaBotones
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCerrarCuentas_Click"))
End Sub

Private Sub cmdCerrarCuentas_Click()
    On Error GoTo NotificaError
    Dim vlstrSentencia As String
    Dim vllngPersonaGraba As Long
    Dim SQL As String
    Dim intDias As Integer
    Dim blnPermisoEscritura As Boolean
    Dim blnPermisoCTotal As Boolean
    Dim vllngVar As Long
    Dim vldtmFechaAntes As Date
    Dim vldtmFechaDespues As Date
    Dim rsCargos As New ADODB.Recordset
    Dim vlmsgCargosProgramados As String
        
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
    Dim vlblnSeleccionados As Boolean
    Dim vllngCuentaCerrada As Integer
    Dim rs As ADODB.Recordset
    Dim vlintNumDiasAbrirExt As Integer
    Dim vlintNumDiasAbrirInt As Integer
    
    Set rs = frsRegresaRs("select intDiasAbrirCuentasInternos, intDiasAbrirCuentasExternos from PVParametro where tnyclaveempresa = " & vgintClaveEmpresaContable, adLockReadOnly, adOpenForwardOnly)
    If Not rs.EOF Then
        vlintNumDiasAbrirExt = IIf(IsNull(rs!intDiasAbrirCuentasExternos), 0, rs!intDiasAbrirCuentasExternos)
        vlintNumDiasAbrirInt = IIf(IsNull(rs!intDiasAbrirCuentasInternos), 0, rs!intDiasAbrirCuentasInternos)
    Else
        vlintNumDiasAbrirExt = 0
        vlintNumDiasAbrirInt = 0
    End If
        
    vllngCuentaCerrada = 0
     
    vlblnSeleccionados = False
    With grdCuentas
        For vllngContador = 1 To .Rows - 1
            If .TextMatrix(vllngContador, 0) = "*" And .TextMatrix(vllngContador, 3) = "Abierta" Then
                vlblnSeleccionados = True
                
                'si existen requisiciones pendientes
                If fblnRequisicionPaciente(Val(.TextMatrix(vllngContador, 1)), .TextMatrix(vllngContador, 10)) Then
                    frmRequisicionesPendientes.pMostrarRequisiciones Val(.TextMatrix(vllngContador, 1)), .TextMatrix(vllngContador, 10), False
                    If Not frmRequisicionesPendientes.lblnContinuarCerrarCuenta Then
                        Unload frmRequisicionesPendientes
                        Exit Sub
                    End If
                    Unload frmRequisicionesPendientes
                End If
            End If
        Next vllngContador
    End With

    If vlblnSeleccionados Then
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba <> 0 Then
            
            EntornoSIHO.ConeccionSIHO.BeginTrans
            vlmsgCargosProgramados = ""
            
            With grdCuentas
                For vllngContador = 1 To .Rows - 1
                    If .TextMatrix(vllngContador, 0) = "*" And .TextMatrix(vllngContador, 3) = "Abierta" Then
                                                
                        If vllngCuentaCerrada = 0 Then
                            vldtmFechaAntes = fdtmServerFechaHora
                            vllngVar = 1
                            frsEjecuta_SP .TextMatrix(vllngContador, 1) & "|'" & .TextMatrix(vllngContador, 10) & "'|" & fstrFechaSQL(CStr(vldtmFechaAntes), CStr(vldtmFechaAntes)), "FN_PVINSCARGOSPROGRAMADOS", True, vllngVar
                            vldtmFechaDespues = fdtmServerFechaHora
                            
                            Set rsCargos = frsEjecuta_SP(.TextMatrix(vllngContador, 1) & "|'" & .TextMatrix(vllngContador, 10) & "'|" & fstrFechaSQL(CStr(vldtmFechaAntes), CStr(vldtmFechaAntes)) & "|" & fstrFechaSQL(CStr(vldtmFechaDespues), CStr(vldtmFechaDespues)), "SP_PVSELCARGOSPROGRAMADOSBITA")
                            If rsCargos.RecordCount > 0 Then
                                rsCargos.MoveFirst
                                Do While Not rsCargos.EOF
                                    vlmsgCargosProgramados = IIf(Trim(vlmsgCargosProgramados) = "", rsCargos!FechaHora & Chr(9) & rsCargos!Cargo & Chr(9) & rsCargos!Descripcion, vlmsgCargosProgramados & Chr(13) & rsCargos!FechaHora & Chr(9) & rsCargos!Cargo & Chr(9) & rsCargos!Descripcion)
                                    rsCargos.MoveNext
                                Loop
                            End If
                            
                            pEjecutaSentencia "UPDATE PVCARGOPROGRAMADO SET chrestado = 'S', INTPERSONAFINALIZA = " & vllngPersonaGraba & ", CHRMEDICOENFERMERA = 'E' WHERE intnumcuenta = " & .TextMatrix(vllngContador, 1) & " AND chrtipoingreso = '" & .TextMatrix(vllngContador, 10) & "' AND chrestado = 'A'"
                        End If
                        
                        'vllngCuentaCerrada  1 Cuenta cerrada 0 Cuenta abierta
                        vgstrParametrosSP = .TextMatrix(vllngContador, 1) & "|" & _
                                            .TextMatrix(vllngContador, 10) & "|" & _
                                            IIf(vllngCuentaCerrada = 1, "0", "1") & "|" & _
                                            IIf(vllngCuentaCerrada = 0 And .TextMatrix(vllngContador, 10) = "E", fstrFechaSQL(fdtmServerFecha, fdtmServerHora), Null)
                          
                        frsEjecuta_SP vgstrParametrosSP, "SP_EXUPDCERRARABRIRCUENTA"
                        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, IIf(vllngCuentaCerrada = 1, "APERTURA DE CUENTA", "CIERRE DE CUENTA"), .TextMatrix(vllngContador, 1))
                    
                        .TextMatrix(vllngContador, 3) = "Cerrada"
                    
                        If vllngCuentaCerrada = 1 Then
                            pEjecutaSentencia "delete PVCuentasReabiertas where numNumCuenta = " & .TextMatrix(vllngContador, 1) & " and chrTipoPaciente = '" & .TextMatrix(vllngContador, 10) & "'"
                            pEjecutaSentencia "insert into PVCuentasReabiertas (numNumCuenta, chrTipoPaciente, intCveEmpleado, intDiasReabrirE, intDiasReabierta) values (" & .TextMatrix(vllngContador, 1) & ", '" & .TextMatrix(vllngContador, 10) & "', " & vllngPersonaGraba & ", " & IIf(.TextMatrix(vllngContador, 10) = "I", vlintNumDiasAbrirInt, vlintNumDiasAbrirExt) & ", " & intDias & ")"
                        End If
                    End If
                Next vllngContador
            End With
                                                                    
            EntornoSIHO.ConeccionSIHO.CommitTrans
            
            If Trim(vlmsgCargosProgramados) <> "" Then
                vlmsgCargosProgramados = "Se realizaron los siguientes cargos automáticos por cuidados especiales pendientes de aplicarse." & Chr(13) & vlmsgCargosProgramados
                
                'La información se actualizó satisfactoriaments mnas los cargos automaticos realizados
                MsgBox SIHOMsg(284) & Chr(13) & Chr(13) & vlmsgCargosProgramados, vbInformation, "Mensaje"
            Else
                'La información se actualizó satisfactoriamente.
                MsgBox SIHOMsg(284), vbInformation, "Mensaje"
            End If
            
            pLimpia
            pHabilitaBotones
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdCerrarCuentas_Click"))
End Sub

Private Sub cmdInvertir_Click()
    pInvertirSeleccion
    pHabilitaBotones
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = 27 Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    vgstrNombreForm = Me.Name

    pCarga
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = False
End Sub

Private Sub grdCuentas_Click()
    pSeleccionaRenglon
End Sub

Private Sub grdCuentas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pSeleccionaRenglon
End Sub

Private Sub pCarga()
    On Error GoTo NotificaError
    Dim rsCuentas As New ADODB.Recordset
    Dim vllngContador As Long
    Dim vlstrsql As String
    
    With grdCuentas
        .Rows = 2
        .Cols = 12
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    End With
    
    Set rsCuentas = frsEjecuta_SP(Str(vgintClaveEmpresaContable), "sp_PVRPTCUENTAPENDIENTEALERTA")
    If rsCuentas.RecordCount <> 0 Then
        Do While Not rsCuentas.EOF
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 0) = ""
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 1) = rsCuentas!NUMCUENTA
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 2) = IIf(rsCuentas!tipo = "I", "Interno", "Externo")
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 3) = IIf(rsCuentas!BITFACTURADA = 0, "Abierta", "Cerrada") 'Aunque diga bitfacturada en realidad hace referencia al BITCUENTACERRADA
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 4) = rsCuentas!FECHACUENTA
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 5) = Trim(rsCuentas!NOMBREPACIENTE)
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 6) = Trim(rsCuentas!DESCRIPCIONTIPOPACIENTE)
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 7) = rsCuentas!Total
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 8) = rsCuentas!MNYSUMAASEGURADA
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 9) = rsCuentas!Total - rsCuentas!MNYSUMAASEGURADA
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 10) = rsCuentas!tipo
            grdCuentas.TextMatrix(grdCuentas.Rows - 1, 11) = rsCuentas!CVEEMPRESATIPO

            rsCuentas.MoveNext
            If Not rsCuentas.EOF Then grdCuentas.Rows = grdCuentas.Rows + 1
        Loop
    End If
    rsCuentas.Close
    
    With grdCuentas
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Cuenta|Tipo|Estado|Fecha ingreso|Paciente|Empresa|Importe cuenta|Suma asegurada|Excedente"
        .ColWidth(0) = 150
        .ColWidth(1) = 900
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 2800
        .ColWidth(6) = 2800
        .ColWidth(7) = 1400
        .ColWidth(8) = 1400
        .ColWidth(9) = 1400
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(7) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(9) = flexAlignCenterCenter
        
        pFormatoNumeroColumnaGrid grdCuentas, 7, "$ "
        pFormatoNumeroColumnaGrid grdCuentas, 8, "$ "
        pFormatoNumeroColumnaGrid grdCuentas, 9, "$ "
        
        pFormatoFechaLargaColumnaGrid grdCuentas, 4
        
        .ScrollBars = flexScrollBarBoth
        
        .Col = 1
        .Row = 1
    End With
    
    pLimpia
    pHabilitaBotones
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pCarga"))
End Sub

Private Sub pSeleccionaRenglon()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
     
    With grdCuentas
        vlintnumrenglon = .Row
        
        If .TextMatrix(vlintnumrenglon, 0) = "*" Then
            .TextMatrix(vlintnumrenglon, 0) = ""
        Else
            .TextMatrix(vlintnumrenglon, 0) = "*"
        End If
        
        .RowSel = vlintnumrenglon
    End With
    
    pHabilitaBotones
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSeleccionaRenglon"))
End Sub

Private Sub pInvertirSeleccion()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
     
    With grdCuentas
        vlintnumrenglon = .Row
        
        For vllngContador = 1 To .Rows - 1
            If .TextMatrix(vllngContador, 0) = "" Then
                .TextMatrix(vllngContador, 0) = "*"
            Else
                .TextMatrix(vllngContador, 0) = ""
            End If
        Next vllngContador
        
        .RowSel = vlintnumrenglon
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pInvertirSeleccion"))
End Sub

Private Sub pLimpia()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
     
    With grdCuentas
        vlintnumrenglon = .Row
        
        For vllngContador = 1 To .Rows - 1
            .TextMatrix(vllngContador, 0) = ""
        Next vllngContador
        
        .RowSel = vlintnumrenglon
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLimpia"))
End Sub

Private Sub pFormatoFechaLargaColumnaGrid(grdNombre As MSHFlexGrid, vlintxColumna As Integer, Optional vlstrSigno As String)
'----------------------------------------------------------------------
' Procedimiento para dar formato a la columna del grid que son Fechas
'----------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim X As Long
    
    For X = 1 To grdNombre.Rows - 1
        grdNombre.TextMatrix(X, vlintxColumna) = Format(grdNombre.TextMatrix(X, vlintxColumna), "DD/MMM/YYYY")
    Next X

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pFormatoFechaColumnaGrid"))
End Sub

Private Sub pHabilitaBotones()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
    Dim vlblnHabilitaAbrir As Boolean
    Dim vlblnHabilitaCerrar As Boolean
     
    vlblnHabilitaAbrir = False
    vlblnHabilitaCerrar = False
     
    With grdCuentas
        vlintnumrenglon = .Row
        
        For vllngContador = 1 To .Rows - 1
            If .TextMatrix(vllngContador, 0) = "*" Then
                If .TextMatrix(vllngContador, 3) = "Abierta" Then vlblnHabilitaCerrar = True
                If .TextMatrix(vllngContador, 3) = "Cerrada" Then vlblnHabilitaAbrir = True
            End If
        Next vllngContador
        
        .RowSel = vlintnumrenglon
    End With
    
    If vlblnPermiteAbrir Then
        cmdAbrirCuentas.Enabled = vlblnHabilitaAbrir
    Else
        cmdAbrirCuentas.Enabled = False
    End If
    
    If vlblnPermiteCerrar Then
        cmdCerrarCuentas.Enabled = vlblnHabilitaCerrar
    Else
        cmdCerrarCuentas.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaBotones"))
End Sub
