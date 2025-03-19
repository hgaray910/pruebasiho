VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmParametroTipoCorte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de corte por departamento"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   720
      Left            =   3975
      TabIndex        =   1
      Top             =   3855
      Width           =   630
      Begin VB.CommandButton cmdregistrar 
         Height          =   495
         Left            =   60
         Picture         =   "frmParametroTipoCorte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Grabar información"
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTipoCorte 
      Height          =   3735
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "Marque el tipo de corte por departamento"
      Top             =   75
      Width           =   8425
      _ExtentX        =   14870
      _ExtentY        =   6588
      _Version        =   393216
      GridColor       =   12632256
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmParametroTipoCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------
' Programa para dar mantenimiento al tipo de corte por departamento
'---------------------------------------------------------------------------------
' Fecha:            Marzo, 2004
' Autor:            Samantha Delgado Terrazas
'---------------------------------------------------------------------------------
' Fecha modificación:
' Autor:
' Descripción:
'---------------------------------------------------------------------------------
Option Explicit

Dim rs As New ADODB.Recordset                   'Varios usos
Dim ldtmFecha As Date                           'Fecha actual
Dim vlstrSentencia As String

Private Function fblnCortesPendientes() As Boolean
On Error GoTo NotificaError

    Dim X As Integer
    Dim vlblnTieneCortesPendientes As Boolean
    Dim vlstrCortesPendientes As String
    
    vlblnTieneCortesPendientes = False
    vlstrCortesPendientes = "Departamento - Tipo de corte - Número de corte - Persona abrió" & Chr(13)
    vlstrCortesPendientes = vlstrCortesPendientes & "---------------------------------------------------------------------" & Chr(13)
    
    For X = 1 To grdTipoCorte.Rows - 1
        If grdTipoCorte.TextMatrix(X, 6) = "C" Then
            vlstrSentencia = "select pvcorte.* ,noempleado.vchnombre empleado " & _
                             "  from pvcorte ,noempleado " & _
                             " where pvcorte.dtmfecharegistro Is Null and pvcorte.smidepartamento = " & grdTipoCorte.TextMatrix(X, 1) & " and pvcorte.intempleado = noempleado.intcveempleado " & _
                             "order by pvcorte.smidepartamento"
            Set rs = frsRegresaRs(vlstrSentencia)
            
            If rs.RecordCount <> 0 Then
                With rs
                    .MoveFirst
                    Do While Not .EOF
                        '----------------------------------------------------------------------'
                        '    Verifica si hay cuentas liquidadas que no han sido facturadas     '
                        '----------------------------------------------------------------------'
                        vgstrParametrosSP = fstrFechaSQL(Format(ldtmFecha - 1, "dd/mm/yyyy")) & "|" & fstrFechaSQL(Format(ldtmFecha - 1, "dd/mm/yyyy")) & "|" & Str(!SMIDEPARTAMENTO)
                        Set rs = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELCUENTALIQUIDADA")
                        If rs!Total > 0 Then
                            '  Realice la facturación de grupo antes de cerrar el corte.
                            'MsgBox SIHOMsg(790), vbCritical, "Mensaje"
                            MsgBox "Realice la facturación de grupo antes de cerrar el corte del departamento " & Str(!SMIDEPARTAMENTO), vbCritical, "Mensaje"
                            fblnCortesPendientes = True
                            Exit Function
                        Else
                            '-----------------------------------------'
                            '    Verifica si existen movimientos     '
                            '-----------------------------------------'
                            vlstrSentencia = "SELECT NVL(COUNT(*),0) FROM PvCortePoliza WHERE intNumCorte = " & !intnumcorte
                            If frsRegresaRs(vlstrSentencia).Fields(0) <> 0 Then 'Si hay documentos registrados en este corte?
                                vlblnTieneCortesPendientes = True
                                vlstrCortesPendientes = vlstrCortesPendientes & Trim(grdTipoCorte.TextMatrix(X, 2)) & " - " & IIf(!chrtipo = "P", "Caja de ingresos", "Caja chica") & " - " & !intnumcorte & " - " & !Empleado & Chr(13)
                            End If
                        End If
                        .MoveNext
                    Loop
                End With
            End If
            rs.Close
        End If
    Next
    If vlblnTieneCortesPendientes Then
        MsgBox "No es posible cambiar el tipo de corte en los siguientes departamentos:" & Chr(13) & Chr(13) & vlstrCortesPendientes & Chr(13) & "Cierre todos los cortes y vuelva a intentarlo.", vbInformation + vbOKOnly, "Mensaje"
    End If
    fblnCortesPendientes = vlblnTieneCortesPendientes
Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnCortesPendientes"))
End Function

Private Sub pCierraCorte()
    On Error GoTo NotificaError
    Dim X As Integer
    Dim vlblnTieneCortesPendientes As Boolean
    Dim vlstrCortesPendientes As String
    
    For X = 1 To grdTipoCorte.Rows - 1
        If grdTipoCorte.TextMatrix(X, 6) = "C" Then
            vlstrSentencia = "select pvcorte.* ,noempleado.vchapellidopaterno||' '||noempleado.vchnombre empleado " & _
                             "  from pvcorte ,noempleado " & _
                             " where pvcorte.dtmfecharegistro Is Null and pvcorte.smidepartamento = " & grdTipoCorte.TextMatrix(X, 1) & " and pvcorte.intempleado = noempleado.intcveempleado " & _
                             "order by pvcorte.smidepartamento"
            Set rs = frsRegresaRs(vlstrSentencia)
            
            If rs.RecordCount <> 0 Then
                With rs
                    .MoveFirst
                    Do While Not .EOF
                        '-----------------------------------------'
                        '            Cierre de corte              '
                        '-----------------------------------------'
                        vlstrSentencia = "update pvcorte set pvcorte.dtmfecharegistro = " & fstrFechaSQL(fdtmServerFecha, fdtmServerHora) & " where pvcorte.intnumcorte = " & !intnumcorte
                        pEjecutaSentencia vlstrSentencia
                        .MoveNext
                    Loop
                End With
            End If
            rs.Close
        End If
    Next
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCierraCorte"))
End Sub

Private Sub cmdRegistrar_Click()
    On Error GoTo NotificaError
    
    Dim vlstrsql As String
    Dim rsTemp As New ADODB.Recordset
    Dim X As Integer
    Dim vllngPersonaGraba As Long
    
    If MsgBox(SIHOMsg(4), vbQuestion + vbYesNo, "Mensaje") = vbYes Then '¿Desea guardar los datos?
        '--------------------------------------------------------
        ' Verifica si el departamento tiene cortes abiertos
        '--------------------------------------------------------
        If fblnCortesPendientes Then
            pIndicaTipoCorte
            Exit Sub
        End If
        
        '--------------------------------------------------------
        ' Persona que graba
        '--------------------------------------------------------
        vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
        If vllngPersonaGraba = 0 Then Exit Sub
        '-----------------------------------------------------
        
        pCierraCorte  ' Cierra los cortes de los departamentos que han cambiado el tipo de corte
        
        Set rsTemp = frsRegresaRs("Select * From GnParametroTipoCorte", adLockOptimistic, adOpenDynamic)
        With rsTemp
            If .RecordCount > 0 Then
                pEjecutaSentencia "Delete from GnParametroTipoCorte"
            End If
            For X = 1 To grdTipoCorte.Rows - 1
                If grdTipoCorte.TextMatrix(X, 3) = "*" Or grdTipoCorte.TextMatrix(X, 4) = "*" Then
                    .AddNew
                    !intCveDepartamento = grdTipoCorte.RowData(X)
                    !intTipoCorte = IIf(grdTipoCorte.TextMatrix(X, 3) = "*", 1, 2)
                    !intConfirmarCierre = IIf(grdTipoCorte.TextMatrix(X, 5) = "*", 1, 0)
                    !INTDESGLOSAPOLIZACORTE = IIf(grdTipoCorte.TextMatrix(X, 9) = "*", 1, 0)
                    .Update
                End If
            Next
            pIndicaTipoCorte
        End With
        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "PARAMETROS DE TIPOS DE CORTES", CStr(vgintNumeroDepartamento))
        MsgBox SIHOMsg(358), vbInformation, "Mensaje" '¡Los datos han sido guardados satisfactoriamente!
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdRegistrar_Click"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    Dim rsDepartamento As New ADODB.Recordset
    Dim vlstrsql As String
    
    ldtmFecha = fdtmServerFecha

    Me.Icon = frmMenuPrincipal.Icon
    
    vlstrsql = "Select smiCveDepartamento, vchDescripcion From Nodepartamento " & _
               " where nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & " order by vchDescripcion"
    Set rsDepartamento = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rsDepartamento.RecordCount > 0 Then
        pLlenarMshFGrdRs grdTipoCorte, rsDepartamento, 0
    Else
        MsgBox SIHOMsg(239), vbOKCancel, "Mensaje"  'No existen departamentos registrados.
    End If
    
    Call pConfiguraColumnas
    Call pIndicaTipoCorte
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Load"))
End Sub

Private Sub pIndicaTipoCorte()
    On Error GoTo NotificaError
    Dim vlstrsql As String
    Dim rsTemp As New ADODB.Recordset
    Dim X As Integer
    
    For X = 1 To grdTipoCorte.Rows - 1
        grdTipoCorte.TextMatrix(X, 3) = " "
        grdTipoCorte.TextMatrix(X, 4) = " "
        grdTipoCorte.TextMatrix(X, 6) = ""
        grdTipoCorte.TextMatrix(X, 7) = " "
        grdTipoCorte.TextMatrix(X, 8) = " "
        grdTipoCorte.TextMatrix(X, 9) = ""
    Next
    
    vlstrsql = " Select * from GnParametroTipoCorte"
    Set rsTemp = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    With rsTemp
        Do While Not .EOF
            For X = 1 To grdTipoCorte.Rows - 1
                If grdTipoCorte.RowData(X) = !intCveDepartamento Then
                    If !intTipoCorte = 1 Then                   'Marca el tipo de corte por departamento
                        grdTipoCorte.TextMatrix(X, 3) = "*"
                        grdTipoCorte.TextMatrix(X, 4) = " "
                        grdTipoCorte.TextMatrix(X, 7) = "*"
                        grdTipoCorte.TextMatrix(X, 8) = " "
                    ElseIf !intTipoCorte = 2 Then
                        grdTipoCorte.TextMatrix(X, 3) = " "
                        grdTipoCorte.TextMatrix(X, 4) = "*"     'Marca el tipo de corte por empleado
                        grdTipoCorte.TextMatrix(X, 7) = " "
                        grdTipoCorte.TextMatrix(X, 8) = "*"
                    Else
                        grdTipoCorte.TextMatrix(X, 3) = " "
                        grdTipoCorte.TextMatrix(X, 4) = " "
                        grdTipoCorte.TextMatrix(X, 7) = " "
                        grdTipoCorte.TextMatrix(X, 8) = " "
                    End If
                    grdTipoCorte.TextMatrix(X, 5) = " "         'Indica si se muestra mensaje por corte
                    If !intConfirmarCierre = 1 Then grdTipoCorte.TextMatrix(X, 5) = "*"
                    
                    grdTipoCorte.TextMatrix(X, 9) = IIf(!INTDESGLOSAPOLIZACORTE = 1, "*", "")
                End If
            Next
            .MoveNext
        Loop
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pIndicaTipoCorte"))
End Sub

Private Sub pConfiguraColumnas()
    On Error GoTo NotificaError
    Dim vlstr As String
    Dim X As Integer
   
    With grdTipoCorte
        .FormatString = "||Departamento|Por departamento|Por empleado|Cierre diario||||Desglosar póliza"
        .Cols = 10
        .ColWidth(0) = 100
        .ColWidth(1) = 0            'Rowdata
        .ColWidth(2) = 3000         'Descripcion departamento
        .ColWidth(3) = 1400         'Tipo de corte por departamento
        .ColAlignment(3) = 4        'FlexAlignCenterCenter
        .ColWidth(4) = 1100         'Tipo de corte por empleado
        .ColAlignment(4) = 4
        .ColWidth(5) = 1200         'Mensaje de cierre de corte diario
        .ColAlignment(5) = 4
        .ColAlignmentFixed(5) = 4
        .ColWidth(6) = 0            'Bandera de cambio
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 1300         'Desglosar el detalle de la póliza
        .ColAlignment(9) = 4
        For X = 1 To .Rows - 1
            .Row = X
            .Col = 3
            .CellFontSize = 12
            .CellFontBold = True
            .Col = 4
            .CellFontSize = 12
            .CellFontBold = True
            .Col = 5
            .CellFontSize = 12
            .CellFontBold = True
            .Col = 9
            .CellFontSize = 12
            .CellFontBold = True
        Next
        .Redraw = True
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraColumnas"))
End Sub

Private Sub grdTipoCorte_Click()
    On Error GoTo NotificaError
    Dim vlblnIndicador As Boolean       'Para no permitir que se seleccione tipos de corte contrarios (osea Depto y empleado)
    
    With grdTipoCorte
        Select Case .Col
            Case 3
                If .TextMatrix(.Row, 4) = "*" Then .TextMatrix(.Row, 4) = " "
            Case 4
                If .TextMatrix(.Row, 3) = "*" Then .TextMatrix(.Row, 3) = " "
            Case 9
                If Trim(.TextMatrix(.Row, 3)) = "*" Or Trim(.TextMatrix(.Row, 4)) = "*" Then
                    .TextMatrix(.Row, 9) = IIf(.TextMatrix(.Row, 9) = "*", "", "*")
                Else
                    .TextMatrix(.Row, 9) = ""
                End If
        End Select
        
        If (.Col = 4 Or .Col = 3 Or .Col = 5) Then
            .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "*", " ", "*")
            
            If .TextMatrix(.Row, 3) <> .TextMatrix(.Row, 7) Then
                .TextMatrix(.Row, 6) = "C"
            Else
                .TextMatrix(.Row, 6) = ""
                .TextMatrix(.Row, 6) = IIf(.TextMatrix(.Row, 4) <> .TextMatrix(.Row, 8), "C", "")
            End If
        End If
        
        If Trim(.TextMatrix(.Row, 3)) = "" And Trim(.TextMatrix(.Row, 4)) = "" Then
            .TextMatrix(.Row, 5) = " "
            .TextMatrix(.Row, 9) = ""
        End If
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdTipoCorte_Click"))
End Sub

Private Sub grdTipoCorte_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    If KeyCode = vbKeyReturn Then grdTipoCorte_Click
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdTipoCorte_KeyDown"))
End Sub
