VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptFoliosAntesVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folios antes de venta"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmDepartamentos 
      Caption         =   "Departamentos"
      Height          =   760
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboDepartamento 
         Height          =   315
         Left            =   130
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione el departamento"
         Top             =   280
         Width           =   4635
      End
   End
   Begin VB.Frame frmFechas 
      Caption         =   "Rango de fechas"
      Height          =   760
      Left            =   660
      TabIndex        =   6
      Top             =   960
      Width           =   3855
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   330
         Left            =   600
         TabIndex        =   1
         ToolTipText     =   "Fecha inicial para el reporte"
         Top             =   285
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   330
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Fecha final para el reporte"
         Top             =   285
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         ClipMode        =   1
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Al"
         Height          =   190
         Left            =   2040
         TabIndex        =   8
         Top             =   355
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   190
         Left            =   130
         TabIndex        =   7
         Top             =   355
         Width           =   735
      End
   End
   Begin VB.Frame frmBotonera 
      Height          =   825
      Left            =   1977
      TabIndex        =   5
      Top             =   1800
      Width           =   1220
      Begin VB.CommandButton cmdImprimir 
         Height          =   495
         Left            =   600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptFoliosAntesVenta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir el reporte"
         Top             =   200
         Width           =   495
      End
      Begin VB.CommandButton cmdVistaPreliminar 
         Height          =   495
         Left            =   100
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptFoliosAntesVenta.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Vista preliminar del reporte"
         Top             =   200
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRptFoliosAntesVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vlstrx As String
Dim rs As New ADODB.Recordset
Dim vglngCuentaPaciente As Long
Public vglngNumeroOpcion As Long
Private vgrptReporte As CRAXDRT.Report
Public vlblnTodosDeptos As Boolean

Private Sub cmdImprimir_Click()
    On Error GoTo NotificaError
    pImprime "I"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdImprimir_Click"))
End Sub

Private Sub cmdVistaPreliminar_Click()
    On Error GoTo NotificaError
    pImprime "P"
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdVistaPreliminar_Click"))
End Sub

Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError

    Dim alstrParametros(11) As String
    Dim rsReporte As New ADODB.Recordset
    Dim vlstrx As String
    
    If fblnVerificaDatos Then
        
        vlstrx = "'" & Format(CDate(txtFechaInicio.Text) & " 00:00:00", "dd/mm/yyyy hh:mm:ss")
        vlstrx = vlstrx & "'|'" & Format(CDate(txtFechaFin.Text) & " 23:59:59", "dd/mm/yyyy hh:mm:ss")
        vlstrx = vlstrx & "'|" & cboDepartamento.ItemData(cboDepartamento.ListIndex)

        Set rsReporte = frsEjecuta_SP(vlstrx, "SP_PVFOLIOANTESVENTA")
        If rsReporte.RecordCount > 0 Then
        
          pInstanciaReporte vgrptReporte, "rptPVFoliosAntesVenta.rpt"
          vgrptReporte.DiscardSavedData
        
          alstrParametros(0) = "p_empresa;" & Trim(vgstrNombreHospitalCH)
          alstrParametros(4) = "p_finicio;" & UCase(Format(txtFechaInicio, "dd/mmm/yyyy"))
          alstrParametros(5) = "p_ffin;" & UCase(Format(txtFechaFin, "dd/mmm/yyyy"))
          alstrParametros(11) = "p_tiporpt;" & "FOLIOS ANTES DE VENTA"
          
          pCargaParameterFields alstrParametros, vgrptReporte
          pImprimeReporte vgrptReporte, rsReporte, vlstrDestino, "Folios antes de venta"
          
        Else
          'No existe información con esos parámetros.
          MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        
        If rsReporte.State <> adStateClosed Then rsReporte.Close
        
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 13 Then
       SendKeys vbTab
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    vgstrNombreForm = Me.Name
    
    Set rs = frsEjecuta_SP("-1|1|*|" & vgintClaveEmpresaContable, "Sp_Gnseldepartamento")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboDepartamento, rs, 0, 1
    End If
    cboDepartamento.AddItem "<TODOS>", 0
    cboDepartamento.ItemData(cboDepartamento.newIndex) = 0
    cboDepartamento.ListIndex = 0
    
    pInicializa
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub txtFechaFin_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    
    pSelMkTexto txtFechaFin
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaFin_GotFocus"))
End Sub

Private Sub txtFechaFin_LostFocus()
    On Error GoTo NotificaError

    If Trim(txtFechaFin.ClipText) = "" Then
        txtFechaFin.Mask = ""
        txtFechaFin.Text = fdtmServerFecha
        txtFechaFin.Mask = "##/##/####"
    Else
        If Not IsDate(txtFechaFin.Text) Then
            txtFechaFin.Mask = ""
            txtFechaFin.Text = fdtmServerFecha
            txtFechaFin.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioIni_KeyPress"))
End Sub

Private Sub txtFechaInicio_GotFocus()
'--------------------------------------------------------------------------
' Procedimiento para que cada vez que tenga el enfoque el control, lo marque
' en azul o seleccionado
'--------------------------------------------------------------------------
    On Error GoTo NotificaError 'Manejo del error
    
    pSelMkTexto txtFechaInicio

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFechaInicio_GotFocus"))
End Sub

Private Sub txtFechaInicio_LostFocus()
    On Error GoTo NotificaError

    If Trim(txtFechaInicio.ClipText) = "" Then
        txtFechaInicio.Mask = ""
        txtFechaInicio.Text = fdtmServerFecha
        txtFechaInicio.Mask = "##/##/####"
    Else
        If Not IsDate(txtFechaInicio.Text) Then
            txtFechaInicio.Mask = ""
            txtFechaInicio.Text = fdtmServerFecha
            txtFechaInicio.Mask = "##/##/####"
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtFolioIni_KeyPress"))
End Sub

Private Sub pInicializa()
    On Error GoTo NotificaError
    
    txtFechaInicio.Mask = ""
    txtFechaInicio.Text = fdtmServerFecha
    txtFechaInicio.Mask = "##/##/####"
    
    txtFechaFin.Mask = ""
    txtFechaFin.Text = fdtmServerFecha
    txtFechaFin.Mask = "##/##/####"
          
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInicializa"))
    Unload Me
End Sub

Private Function fblnVerificaDatos() As Boolean
Dim rsVerifica As New ADODB.Recordset

    On Error GoTo NotificaError
    
    fblnVerificaDatos = True
        
    If Not IsDate(txtFechaInicio.Text) Or CDate(txtFechaInicio.Text) < CDate("01/01/1900") Then ' fechas no menores a 1900
        fblnVerificaDatos = False
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
        txtFechaInicio.SetFocus
    Else
        If Not IsDate(txtFechaFin.Text) Or CDate(txtFechaFin.Text) < CDate("01/01/1900") Then ' fechas no menores a 1900
            fblnVerificaDatos = False
            '¡Fecha no válida!, formato de fecha dd/mm/aaaa
            MsgBox SIHOMsg(29), vbOKOnly + vbExclamation, "Mensaje"
            txtFechaFin.SetFocus
        Else
            If CDate(txtFechaInicio.Text) > fdtmServerFecha Then
                fblnVerificaDatos = False
                '¡La fecha debe ser menor o igual a la del sistema!
                MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
                txtFechaInicio.SetFocus
            Else
                If CDate(txtFechaFin.Text) > fdtmServerFecha Then
                    fblnVerificaDatos = False
                    '¡La fecha debe ser menor o igual a la del sistema!
                    MsgBox SIHOMsg(40), vbOKOnly + vbExclamation, "Mensaje"
                    txtFechaFin.SetFocus
                Else
                    If CDate(txtFechaInicio.Text) > CDate(txtFechaFin.Text) Then
                        fblnVerificaDatos = False
                        '¡Rango de fechas no válido!
                        MsgBox SIHOMsg(64), vbOKOnly + vbExclamation, "Mensaje"
                        txtFechaInicio.SetFocus
                    End If
                End If
            End If
        End If
    End If
    
    Exit Function

NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":fblnVerificaDatos"))
    Unload Me
End Function

