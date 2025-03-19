VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporteCuentasReabiertas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas reabiertas"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   915
      Left            =   6780
      TabIndex        =   17
      Top             =   1530
      Width           =   450
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   4980
      TabIndex        =   16
      Top             =   1530
      Width           =   450
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango de fechas en que se reabrieron las cuentas"
      Height          =   915
      Left            =   105
      TabIndex        =   13
      Top             =   1530
      Width           =   4830
      Begin MSComCtl2.DTPicker dtpFechaI 
         Height          =   315
         Left            =   915
         TabIndex        =   2
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   58064899
         CurrentDate     =   39071
      End
      Begin MSComCtl2.DTPicker dtpFechaF 
         Height          =   315
         Left            =   3075
         TabIndex        =   3
         Top             =   390
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   58064899
         CurrentDate     =   39071
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2505
         TabIndex        =   15
         Top             =   450
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   105
      TabIndex        =   11
      Top             =   0
      Width           =   7125
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la empresa"
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame6 
      Height          =   915
      Left            =   5475
      TabIndex        =   7
      Top             =   1530
      Width           =   1260
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteCuentasReabiertas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Vista previa"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   630
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmReporteCuentasReabiertas.frx":0403
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   105
      TabIndex        =   6
      Top             =   690
      Width           =   7125
      Begin VB.ComboBox cboEmpleados 
         Height          =   315
         Left            =   915
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empleado"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdSigiente 
      Caption         =   "Siguiente"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "frmReporteCuentasReabiertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private lrptReporte As CRAXDRT.Report

Private Sub cboHospital_Click()
    On Error GoTo NotificaError


    If cboHospital.ListIndex <> -1 Then
        pLlenaCboEmpleados
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboHospital_Click"))
End Sub

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        cboEmpleados.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cboHospital_KeyDown"))
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo NotificaError

    pImprime "P"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPreview_Click"))
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo NotificaError

    pImprime "I"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPrint_Click"))
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo NotificaError

    Unload Me
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSalir_Click"))
End Sub

Private Sub cmdSigiente_Click()
    On Error GoTo NotificaError

    SendKeys vbTab
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdSigiente_Click"))
End Sub

Private Sub dtpFechaF_GotFocus()
    On Error GoTo NotificaError

    dtpFechaF.CustomFormat = "dd/MM/yyyy"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":dtpFechaF_GotFocus"))
End Sub

Private Sub dtpFechaF_LostFocus()
    On Error GoTo NotificaError

    dtpFechaF.CustomFormat = "dd/MMM/yyyy"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":dtpFechaF_LostFocus"))
End Sub

Private Sub dtpFechaI_GotFocus()
    On Error GoTo NotificaError

    dtpFechaI.CustomFormat = "dd/MM/yyyy"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":dtpFechaI_GotFocus"))
End Sub

Private Sub dtpFechaI_LostFocus()
    On Error GoTo NotificaError

    dtpFechaI.CustomFormat = "dd/MMM/yyyy"
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":dtpFechaI_LostFocus"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    
    Dim lngNumOpcion As Long

    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 1919
    Case "SE"
         lngNumOpcion = 2001
    End Select
    
    pCargaHospital lngNumOpcion
    
    dtpFechaI.Value = Date
    dtpFechaF.Value = Date
    pInstanciaReporte lrptReporte, "rptCuentasReabiertas.rpt"
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub pLlenaCboEmpleados()
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    
    cboEmpleados.Clear
    vgstrParametrosSP = "-1|1|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    Set rs = frsEjecuta_SP(vgstrParametrosSP, "Sp_Gnselempleado")
    If rs.RecordCount <> 0 Then
        Do Until rs.EOF
            cboEmpleados.AddItem rs!Nombre
            cboEmpleados.ItemData(cboEmpleados.NewIndex) = rs!Clave
            rs.MoveNext
        Loop
    End If
    rs.Close
    cboEmpleados.AddItem "<TODOS>", 0
    cboEmpleados.ItemData(cboEmpleados.NewIndex) = -1
    cboEmpleados.ListIndex = 0
    
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pLlenaCboEmpleados"))
End Sub

Private Sub pImprime(strDestino As String)
    On Error GoTo NotificaError

    Dim rs As ADODB.Recordset
    Dim strParametrosSP As String
    Dim alstrParametros(3) As String
    strParametrosSP = cboEmpleados.ItemData(cboEmpleados.ListIndex)
    strParametrosSP = strParametrosSP & "|" & Format(dtpFechaI.Value, "yyyy-MM-dd") & "|" & Format(dtpFechaF.Value, "yyyy-MM-dd") & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
    Set rs = frsEjecuta_SP(strParametrosSP, "sp_PVRptCuentasReabiertas")
    If Not rs.EOF Then
        alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
        alstrParametros(1) = "Empleado;" & cboEmpleados.Text
        alstrParametros(2) = "FechaInicio;" & UCase(Format(dtpFechaI.Value, "dd/MMM/yyyy"))
        alstrParametros(3) = "FechaFin;" & UCase(Format(dtpFechaF.Value, "dd/MMM/yyyy"))
        pCargaParameterFields alstrParametros, lrptReporte
        pImprimeReporte lrptReporte, rs, strDestino, "Reporte de cuentas reabiertas"
    Else
        MsgBox SIHOMsg(13), vbInformation, "Mensaje"
    End If
    rs.Close
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Sub pCargaHospital(lngNumOpcion As Long)
    On Error GoTo NotificaError
    
    Dim rs As New ADODB.Recordset
    
    Set rs = frsEjecuta_SP("-1", "Sp_Gnselempresascontable")
    If rs.RecordCount <> 0 Then
        pLlenarCboRs cboHospital, rs, 1, 0
        cboHospital.ListIndex = flngLocalizaCbo(cboHospital, Str(vgintClaveEmpresaContable))
    End If
    
    cboHospital.Enabled = fblnRevisaPermiso(vglngNumeroLogin, lngNumOpcion, "C")
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaHospital"))
    Unload Me
End Sub

