VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMantoTipoCambio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de cambio"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDias 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   14
      ToolTipText     =   "Número de dias "
      Top             =   5940
      Width           =   465
   End
   Begin VB.CheckBox chkVariacion 
      Caption         =   "Mostrar variaciones"
      Height          =   285
      Left            =   60
      TabIndex        =   13
      ToolTipText     =   "Mostrar la informacion de las variaciones al tipo de cambio"
      Top             =   5955
      Width           =   1740
   End
   Begin VB.Frame Frame3 
      Height          =   720
      Left            =   3480
      TabIndex        =   8
      Top             =   6330
      Width           =   645
      Begin VB.CommandButton cmdSave 
         Height          =   495
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMantoTipoCambio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Guardar datos"
         Top             =   165
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5115
      Left            =   60
      TabIndex        =   7
      Top             =   765
      Width           =   7440
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdTipoCambio 
         Height          =   4590
         Left            =   165
         TabIndex        =   4
         ToolTipText     =   "Tipos de cambio"
         Top             =   300
         Width           =   7110
         _ExtentX        =   12541
         _ExtentY        =   8096
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   7440
      Begin VB.TextBox txtCompra 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6210
         TabIndex        =   3
         ToolTipText     =   "Tipo de cambio a la compra"
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtVenta 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   2
         ToolTipText     =   "Tipo de cambio a la venta"
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtOficial 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2775
         TabIndex        =   1
         ToolTipText     =   "Tipo de cambio oficial"
         Top             =   240
         Width           =   1000
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   810
         TabIndex        =   0
         ToolTipText     =   "Fecha del tipo de cambio"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Compra"
         Height          =   195
         Left            =   5535
         TabIndex        =   12
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Venta"
         Height          =   195
         Left            =   3915
         TabIndex        =   11
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Oficial"
         Height          =   195
         Left            =   2190
         TabIndex        =   10
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.Label Label5 
      Caption         =   "días anteriores"
      Height          =   210
      Left            =   2430
      TabIndex        =   15
      Top             =   5992
      Width           =   1335
   End
End
Attribute VB_Name = "frmMantoTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '-----------------------------------------------------------------------------
'Forma para dar mantenimiento al tipo de cambio del dia
'Fecha de programacion: Lunes 9 de abril de 2001
'-----------------------------------------------------------------------------
'Ultimas modificaciones, especificar:
'Fecha:
'Descripción del cambio:
'-----------------------------------------------------------------------------

Option Explicit

Dim rsTipoCambio As New ADODB.Recordset
Dim rsTipoCambioVariacion As New ADODB.Recordset

Dim vlstrSentencia As String

Dim vllngPersonaGraba As Long

Dim vlintNumeroDias As Integer
Public vllngNumeroOpcionModulo As Long
    
Private Sub chkVariacion_Click()
    On Error GoTo NotificaError
    
    pCargaTipo

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":chkVariacion_Click"))
End Sub

Private Sub cmdSave_Click()
    On Error GoTo NotificaError
    
    Dim rsTipos As New ADODB.Recordset
    Dim vllngSecuencia As Long

    If fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "E", True) Or fblnRevisaPermiso(vglngNumeroLogin, vllngNumeroOpcionModulo, "C", True) Then
        If fblnDatosValidos() Then
            vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
            If vllngPersonaGraba <> 0 Then
                                
                EntornoSIHO.ConeccionSIHO.BeginTrans
                    vlstrSentencia = "select mnyCantidadOficial," & _
                    "mnyCantidadVenta," & _
                    "mnyCantidadCompra " & _
                    "from TipoCambio " & _
                    "where dtmFecha=" & fstrFechaSQL(MskFecha.Text)
                    Set rsTipos = frsRegresaRs(vlstrSentencia)
                    If rsTipos.RecordCount > 0 Then
                        Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "TIPO DE CAMBIO", rsTipoCambio!intConsecutivo & "--" & MskFecha.Text)
                        If Val(Format(txtOficial.Text, "###.####")) <> rsTipos!mnyCantidadOficial Or Val(Format(txtVenta.Text, "###.####")) <> rsTipos!mnyCantidadVenta Or Val(Format(txtCompra.Text, "###.####")) <> rsTipos!mnyCantidadCompra Then
                            
                            vlstrSentencia = "Update TipoCambio set " & _
                            "mnyCantidadOficial=" & txtOficial.Text & "," & _
                            "mnyCantidadVenta=" & txtVenta.Text & "," & _
                            "mnyCantidadCompra=" & txtCompra.Text & " " & _
                            "where dtmFecha=" & fstrFechaSQL(MskFecha.Text)
                            
                            pEjecutaSentencia vlstrSentencia
                            
                            pGuardaVariacion
                        End If
                    Else
                        With rsTipoCambio
                            .AddNew
                            !dtmfecha = MskFecha.Text
                            !mnyCantidadOficial = Val(Format(txtOficial.Text, "####.####"))
                            !mnyCantidadVenta = Val(Format(txtVenta.Text, "####.####"))
                            !mnyCantidadCompra = Val(Format(txtCompra.Text, "####.####"))
                            .Update
                            vllngSecuencia = rsTipoCambio!intConsecutivo
                            vllngSecuencia = flngObtieneIdentity("sec_TipoCambio", vllngSecuencia)
                            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "TIPO DE CAMBIO", CStr(vllngSecuencia) & "--" & MskFecha.Text)
                            .Requery
                            .Find ("intConsecutivo =" & vllngSecuencia)
                        End With
                        
                        pGuardaVariacion
                    End If
                EntornoSIHO.ConeccionSIHO.CommitTrans
                
                pCargaTipo
                
                MskFecha.SetFocus
            End If
        End If
    Else
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub pGuardaVariacion()
    On Error GoTo NotificaError
    
    With rsTipoCambioVariacion
        .AddNew
        !dtmfecha = MskFecha.Text
        !dtmFechaHoraVariacion = fdtmServerFecha + fdtmServerHora
        !mnyCantidadOficial = Val(Format(txtOficial.Text, "####.####"))
        !mnyCantidadVenta = Val(Format(txtVenta.Text, "####.####"))
        !mnyCantidadCompra = Val(Format(txtCompra.Text, "####.####"))
        !intcveempleado = vllngPersonaGraba
        !smicvedepartamento = vgintNumeroDepartamento
        .Update
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pGuardaVariacion"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    Dim vlrsAux As New ADODB.Recordset
    
    fblnDatosValidos = True
    
    If Not IsDate(MskFecha.Text) Then
        fblnDatosValidos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
        MskFecha.SetFocus
    End If
    If Year(CDate(MskFecha.Text)) > Year(fdtmServerFecha) Then
        fblnDatosValidos = False
        MsgBox "El año de la fecha no debe ser mayor al año en curso.", vbOKOnly + vbInformation, "Mensaje"
        MskFecha.SetFocus
    End If
    If fblnDatosValidos And Val(Format(txtOficial.Text, "####.####")) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtOficial.SetFocus
    End If
    If fblnDatosValidos And Val(Format(txtVenta.Text, "####.####")) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtVenta.SetFocus
    End If
    If fblnDatosValidos And Val(Format(txtCompra.Text, "####.####")) = 0 Then
        fblnDatosValidos = False
        '¡No ha ingresado datos!
        MsgBox SIHOMsg(2), vbOKOnly + vbInformation, "Mensaje"
        txtCompra.SetFocus
    End If

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnDatosValidos"))
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlstrSentencia = "select * from TipoCambio"
    Set rsTipoCambio = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    vlstrSentencia = "select * from TipoCambioVariacion"
    Set rsTipoCambioVariacion = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
    
    vlintNumeroDias = 10
    
    txtDias.Text = vlintNumeroDias
    
    pCargaTipo

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub pCargaTipo()
    On Error GoTo NotificaError
    
    Dim rsTipos As New ADODB.Recordset
    
    grdTipoCambio.Rows = 0
    
    If chkVariacion.Value = 0 Then
        vlstrSentencia = "select dtmFecha," & _
        "mnyCantidadOficial," & _
        "mnyCantidadVenta," & _
        "mnyCantidadCompra " & _
        "from TipoCambio " & _
        "where dtmFecha>=getdate()-" & txtDias.Text & " " & _
        "order by dtmFecha desc"
        
        Set rsTipos = frsRegresaRs(vlstrSentencia)
        If rsTipos.RecordCount <> 0 Then
            pLlenarMshFGrdRs grdTipoCambio, rsTipos
            
            With grdTipoCambio
                .FixedCols = 1
                .FixedRows = 1
                .FormatString = "|Fecha|Oficial|Venta|Compra"
                .ColWidth(0) = 100
                .ColWidth(1) = 1150
                .ColWidth(2) = 1850
                .ColWidth(3) = 1850
                .ColWidth(4) = 1850
                .ColAlignmentFixed(1) = flexAlignCenterCenter
                .ColAlignmentFixed(2) = flexAlignCenterCenter
                .ColAlignmentFixed(3) = flexAlignCenterCenter
                .ColAlignmentFixed(4) = flexAlignCenterCenter
            End With
            
'            pFormatoNumeroColumnaGrid grdTipoCambio, 2
            'pFormatoNumeroColumnaGrid grdTipoCambio, 3
            'pFormatoNumeroColumnaGrid grdTipoCambio, 4
        End If
    Else
        vlstrSentencia = "select dtmFecha," & _
        "dtmFechaHoraVariacion," & _
        "mnyCantidadOficial," & _
        "mnyCantidadVenta," & _
        "mnyCantidadCompra," & _
        "rtrim(NoEmpleado.vchApellidoPaterno)||' '||rtrim(NoEmpleado.vchApellidoMaterno)||' '||rtrim(NoEmpleado.vchNombre) as Empleado," & _
        "NoDepartamento.vchDescripcion " & _
        "From TipoCambioVariacion " & _
        "inner join NoEmpleado on " & _
        "TipoCambioVariacion.intCveEmpleado = NoEmpleado.intCveEmpleado " & _
        "inner join NoDepartamento on " & _
        "TipoCambioVariacion.smiCveDepartamento = NoDepartamento.smiCveDepartamento " & _
        "Where dtmFecha >= getdate()-" & txtDias.Text & " " & _
        "order by dtmFecha desc "
        
        Set rsTipos = frsRegresaRs(vlstrSentencia)
        If rsTipos.RecordCount <> 0 Then
            pLlenarMshFGrdRs grdTipoCambio, rsTipos
            
            With grdTipoCambio
                .FixedCols = 1
                .FixedRows = 1
                .FormatString = "|Fecha|Fecha variación|Oficial|Venta|Compra|Empleado|Departamento"
                .ColWidth(0) = 100
                .ColWidth(1) = 1000
                .ColWidth(2) = 1900
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                .ColWidth(5) = 1000
                .ColWidth(6) = 2500
                .ColWidth(7) = 2000
                .ColAlignmentFixed(1) = flexAlignCenterCenter
                .ColAlignmentFixed(2) = flexAlignCenterCenter
                .ColAlignmentFixed(3) = flexAlignCenterCenter
                .ColAlignmentFixed(4) = flexAlignCenterCenter
                .ColAlignmentFixed(5) = flexAlignCenterCenter
                .ColAlignmentFixed(6) = flexAlignCenterCenter
                .ColAlignmentFixed(7) = flexAlignCenterCenter
            End With
            
            'pFormatoNumeroColumnaGrid grdTipoCambio, 3
            'pFormatoNumeroColumnaGrid grdTipoCambio, 4
            'pFormatoNumeroColumnaGrid grdTipoCambio, 5
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaTipo"))
End Sub

Private Sub mskFecha_GotFocus()
    On Error GoTo NotificaError
    
    MskFecha.Mask = ""
    MskFecha.Text = ""
    MskFecha.Mask = "##/##/####"
    
    txtOficial.Text = ""
    txtVenta.Text = ""
    txtCompra.Text = ""
    
    pSelMkTexto MskFecha

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecha_GotFocus"))
End Sub

Private Sub MskFecha_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    Dim rsTipos As New ADODB.Recordset
    
    If KeyAscii = 13 Then
        If Trim(MskFecha.ClipText) = "" Then
            MskFecha.Text = fdtmServerFecha
        End If
    
        If IsDate(MskFecha.Text) Then
            If Year(CDate(MskFecha.Text)) <= Year(fdtmServerFecha) Then
                vlstrSentencia = "select mnyCantidadOficial," & _
                "mnyCantidadVenta," & _
                "mnyCantidadCompra " & _
                "from TipoCambio " & _
                "where dtmFecha=" + fstrFechaSQL(MskFecha.Text)
                Set rsTipos = frsRegresaRs(vlstrSentencia)
            
                If rsTipos.RecordCount <> 0 Then
                    txtOficial.Text = FormatNumber(Str(rsTipos!mnyCantidadOficial), 4)
                    txtVenta.Text = FormatNumber(Str(rsTipos!mnyCantidadVenta), 4)
                    txtCompra.Text = FormatNumber(Str(rsTipos!mnyCantidadCompra), 4)
        
                End If
                txtOficial.SetFocus
            ElseIf Year(CDate(MskFecha.Text)) > Year(fdtmServerFecha) Then
                MsgBox "El año de la fecha no debe ser mayor al año en curso.", vbOKOnly + vbInformation, "Mensaje"
                MskFecha.SelStart = 0
                MskFecha.Mask = ""
                MskFecha.Text = ""
                MskFecha.Mask = "##/##/####"
            End If
        Else
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbInformation, "Mensaje"
            MskFecha.SelStart = 0
            MskFecha.Mask = ""
            MskFecha.Text = ""
            MskFecha.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":mskFecha_KeyPress"))
End Sub

Private Sub txtCompra_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtCompra

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCompra_GotFocus"))
End Sub

Private Sub txtCompra_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyLeft Then
        txtVenta.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCompra_KeyDown"))
End Sub

Private Sub txtCompra_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(txtCompra.Text) = "" Then
            txtCompra.Text = "0"
        End If
        txtCompra.Text = FormatNumber(txtCompra.Text, 4)
        cmdSave.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtCompra)) Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtCompra_KeyPress"))
End Sub

Private Sub txtDias_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtDias

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDias_GotFocus"))
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn Then
        KeyAscii = 7
    Else
        If KeyAscii = 13 Then
            If Val(txtDias.Text) = 0 Then
                txtDias.Text = vlintNumeroDias
            End If
            pCargaTipo
            MskFecha.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtDias_KeyPress"))
End Sub

Private Sub txtOficial_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtOficial

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtOficial_GotFocus"))
End Sub

Private Sub txtOficial_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyLeft Then
        MskFecha.SetFocus
    Else
        If KeyCode = vbKeyRight Then
            txtVenta.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtOficial_KeyDown"))
End Sub

Private Sub txtOficial_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(txtOficial.Text) = "" Then
            txtOficial.Text = "0"
        End If
        txtOficial.Text = FormatNumber(txtOficial.Text, 4)
        txtVenta.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtOficial)) Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtOficial_KeyPress"))
End Sub

Private Sub txtVenta_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtVenta

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVenta_GotFocus"))
End Sub

Private Sub txtVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyLeft Then
        txtOficial.SetFocus
    Else
        If KeyCode = vbKeyRight Then
            txtCompra.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVenta_KeyDown"))
End Sub

Private Sub txtVenta_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(txtVenta.Text) = "" Then
            txtVenta.Text = "0"
        End If
        txtVenta.Text = FormatNumber(txtVenta.Text, 4)
        txtCompra.SetFocus
    Else
        If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack And Not KeyAscii = vbKeyReturn And Not KeyAscii = 46 Or (KeyAscii = 46 And fblnValidaPunto(txtVenta)) Then
            KeyAscii = 7
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtVenta_KeyPress"))
End Sub
