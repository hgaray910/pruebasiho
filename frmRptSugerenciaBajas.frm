VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptSugerenciaBajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sugerencias de baja"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReporte 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   5055
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4335
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   300
            Left            =   2040
            TabIndex        =   0
            ToolTipText     =   "Fecha de aplicación"
            Top             =   210
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   122814467
            CurrentDate     =   40882
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de aplicación"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.Frame fraImprime 
         Height          =   735
         Left            =   1755
         TabIndex        =   4
         Top             =   840
         Width           =   1155
         Begin VB.CommandButton cmdPreview 
            Height          =   495
            Left            =   80
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptSugerenciaBajas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Vista previa"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   580
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptSugerenciaBajas.frx":0403
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir"
            Top             =   150
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmRptSugerenciaBajas"
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
    
    dtpFechaFinal.Value = fdtmServerFecha
         
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSocios As New ADODB.Recordset
    Dim alstrParametros(1) As String
    Dim strParametros As String
    
        
    pInstanciaReporte vgrptReporte, "rptSugBajasSocios.rpt"
        
    vgrptReporte.DiscardSavedData
    

    strParametros = CStr(dtpFechaFinal.Value)
    
    Set rsSocios = frsEjecuta_SP(strParametros, "Sp_SORPTSUGBAJAS")
    
    If rsSocios.RecordCount <> 0 Then
            
            alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = "dtmFecha;" & UCase(CStr(Format(dtpFechaFinal.Value, "dd/MMM/yyyy")))
            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Sugerencias de baja"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSocios.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

