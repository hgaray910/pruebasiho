VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReporteSaldoSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de cuotas"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReporte 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5055
      Begin VB.Frame fraImprime 
         Height          =   735
         Left            =   2000
         TabIndex        =   8
         Top             =   1680
         Width           =   1155
         Begin VB.CommandButton cmdPreview 
            Height          =   495
            Left            =   80
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmReporteSaldoSocios.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Vista previa"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   580
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmReporteSaldoSocios.frx":0403
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Imprimir"
            Top             =   150
            Width           =   495
         End
      End
      Begin VB.Frame fraRangoFechas 
         Caption         =   "Rango de fechas"
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4755
         Begin MSComCtl2.DTPicker dtpFechaInicial 
            Height          =   300
            Left            =   825
            TabIndex        =   1
            ToolTipText     =   "Fecha de inicio"
            Top             =   615
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   65011715
            CurrentDate     =   40882
         End
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   300
            Left            =   3105
            TabIndex        =   2
            ToolTipText     =   "Fecha final"
            Top             =   600
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   65011715
            CurrentDate     =   40882
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   645
            Width           =   465
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2445
            TabIndex        =   6
            Top             =   645
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "frmReporteSaldoSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vgrptReporte As CRAXDRT.Report

Private Sub cmdPreview_Click()
    
On Error GoTo NotificaError
    
    pImprime "P"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPreview_Click"))
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()

On Error GoTo NotificaError
    
    pImprime "I"

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdPrint_Click"))
    Unload Me
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo NotificaError

    Select Case KeyCode
        
        Case 27
                
                    Unload Me
                    
        Case 13
            
                SendKeys vbTab
                        
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyDown"))
    Unload Me
    
End Sub

Private Sub Form_Load()

On Error GoTo NotificaError

    Me.Icon = frmMenuPrincipal.Icon
    
    dtpFechaInicial.Value = fdtmServerFecha
    dtpFechaFinal.Value = fdtmServerFecha
         
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSaldoSocio As New ADODB.Recordset
    Dim alstrParametros(3) As String
    Dim strParametros As String
    
        
    pInstanciaReporte vgrptReporte, "rptCuotasSocios.rpt"
        
    vgrptReporte.DiscardSavedData
    

    strParametros = CStr(dtpFechaInicial.Value) & "|" & CStr(dtpFechaFinal.Value)
    
    Set rsSaldoSocio = frsEjecuta_SP(strParametros, "SP_PVRPTCUOTASOCIOS")
    
    If rsSaldoSocio.RecordCount <> 0 Then
            
            alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = "FechaIni;" & CStr(Format(dtpFechaInicial.Value, "dd/MMM/yyyy"))
            alstrParametros(2) = "FechaFin;" & CStr(Format(dtpFechaFinal.Value, "dd/MMM/yyyy"))
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsSaldoSocio, IIf(vlstrTipo = "P", "P", "I"), "CUOTAS SOCIOS"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSaldoSocio.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

