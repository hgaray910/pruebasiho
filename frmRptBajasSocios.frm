VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptBajasSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bajas de socios"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4965
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReporte 
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   5055
      Begin VB.Frame fraRangoFechas 
         Caption         =   "Rango de fechas"
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   4755
         Begin MSComCtl2.DTPicker dtpFechaInicial 
            Height          =   300
            Left            =   825
            TabIndex        =   3
            ToolTipText     =   "Fecha de inicio"
            Top             =   375
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   122880003
            CurrentDate     =   40882
         End
         Begin MSComCtl2.DTPicker dtpFechaFinal 
            Height          =   300
            Left            =   3105
            TabIndex        =   4
            ToolTipText     =   "Fecha final"
            Top             =   360
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "dd/MMM/yyy"
            Format          =   122880003
            CurrentDate     =   40882
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   2445
            TabIndex        =   12
            Top             =   405
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   405
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de socio"
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4755
         Begin VB.OptionButton optTipoSocio 
            Caption         =   "Ambos"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   0
            ToolTipText     =   "Ambos"
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTipoSocio 
            Caption         =   "Titulares"
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   1
            ToolTipText     =   "Titulares"
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton optTipoSocio 
            Caption         =   "Dependientes"
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   2
            ToolTipText     =   "Dependientes"
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame fraImprime 
         Height          =   735
         Left            =   2000
         TabIndex        =   8
         Top             =   2040
         Width           =   1155
         Begin VB.CommandButton cmdPreview 
            Height          =   495
            Left            =   80
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptBajasSocios.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Vista previa"
            Top             =   150
            Width           =   495
         End
         Begin VB.CommandButton cmdPrint 
            Height          =   495
            Left            =   580
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRptBajasSocios.frx":0403
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Imprimir"
            Top             =   150
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmRptBajasSocios"
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
    
    dtpFechaInicial.Value = DateSerial(Year(Date), Month(Date), 1) 'fdtmServerFecha
    dtpFechaFinal.Value = DateSerial(Year(Date), IIf(Month(Date) = 12, 1, Month(Date) + 1), 1) 'fdtmServerFecha
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
    Unload Me
    
End Sub

Private Sub pImprime(vlstrTipo As String)

On Error GoTo NotificaError

    Dim rsSocios As New ADODB.Recordset
    Dim alstrParametros(3) As String
    Dim strParametros As String
    
        
    pInstanciaReporte vgrptReporte, "rptBajasSocios.rpt"
        
    vgrptReporte.DiscardSavedData
    

    strParametros = CStr(dtpFechaInicial.Value) & "|" & CStr(dtpFechaFinal.Value) & "|" & IIf(optTipoSocio(0).Value, "A", IIf(optTipoSocio(1).Value, "T", IIf(optTipoSocio(2).Value, "D", "D")))
    
    Set rsSocios = frsEjecuta_SP(strParametros, "Sp_SORPTBAJAS")
    
    If rsSocios.RecordCount <> 0 Then
            
            alstrParametros(0) = "lblNombreHospital;" & Trim(vgstrNombreHospitalCH)
            alstrParametros(1) = UCase("FechaIni;" & CStr(Format(dtpFechaInicial.Value, "dd/MMM/yyyy")))
            alstrParametros(2) = UCase("FechaFin;" & CStr(Format(dtpFechaFinal.Value, "dd/MMM/yyyy")))
            alstrParametros(3) = UCase("TipoSocio;" & IIf(optTipoSocio(0).Value, "TITULARES Y DEPENDIENTES", IIf(optTipoSocio(1).Value, "TITULARES", IIf(optTipoSocio(2).Value, "DEPENDIENTES", "DEPENDIENTES"))))
            pCargaParameterFields alstrParametros, vgrptReporte

            pCargaParameterFields alstrParametros, vgrptReporte
            pImprimeReporte vgrptReporte, rsSocios, IIf(vlstrTipo = "P", "P", "I"), "Bajas de socios"
    
    Else
        'No existe información con esos parámetros
        MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
    End If
    
    rsSocios.Close

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pImprime"))
End Sub

