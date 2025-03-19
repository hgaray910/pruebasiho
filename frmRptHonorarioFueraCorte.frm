VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRptHonorarioFueraCorte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Honorarios en efectivo fuera del corte"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   675
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   6930
      Begin VB.ComboBox cboHospital 
         Height          =   315
         Left            =   780
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
      Height          =   765
      Left            =   5610
      TabIndex        =   10
      Top             =   675
      Width           =   1380
      Begin VB.CommandButton cmdPrint 
         Height          =   495
         Left            =   675
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptHonorarioFueraCorte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Imprimir"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   510
      End
      Begin VB.CommandButton cmdPreview 
         Height          =   495
         Left            =   165
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRptHonorarioFueraCorte.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Vista previa"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   60
      TabIndex        =   7
      Top             =   675
      Width           =   5535
      Begin MSMask.MaskEdBox mskHoraInicio 
         Height          =   315
         Left            =   2115
         TabIndex        =   2
         ToolTipText     =   "Hora de inicio"
         Top             =   240
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaInicio 
         Height          =   315
         Left            =   780
         TabIndex        =   1
         ToolTipText     =   "Fecha de inicio "
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFechaFin 
         Height          =   315
         Left            =   3480
         TabIndex        =   3
         ToolTipText     =   "Fecha fin"
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mmm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskHoraFin 
         Height          =   315
         Left            =   4770
         TabIndex        =   4
         ToolTipText     =   "Hora de fin"
         Top             =   240
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2985
         TabIndex        =   9
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmRptHonorarioFueraCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
'| Nombre del proyecto      : prjCaja
'| Nombre del formulario    : frmRptHonorarioFueraCorte
'-------------------------------------------------------------------------
'| Objetivo: Listado de los honorarios que se pagan en efectivo en caja
'| y que no entran en el corte
'-------------------------------------------------------------------------

Option Explicit

Private Sub cboHospital_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    If KeyCode = vbKeyReturn Then
        mskFechaInicio.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdPreview_Click"))
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

Private Sub pImprime(vlstrDestino As String)
    On Error GoTo NotificaError

    
    Dim vlrptReporte As CRAXDRT.Report
    Dim vlstrFechaInicio As String
    Dim vlstrFechaFin As String
    Dim rsReporte As New ADODB.Recordset
    Dim alstrParametros(4) As String

    If fblnDatosValidos() Then
        pInstanciaReporte vlrptReporte, "rptpvHonorarioFueraCorte.rpt"
        vlrptReporte.DiscardSavedData
        
        vlstrFechaInicio = fstrFechaSQL(mskFechaInicio.Text, mskHoraInicio.Text & ":00", True)
        vlstrFechaFin = fstrFechaSQL(mskFechaFin.Text, mskHoraFin.Text & ":59", True)
        
        vgstrParametrosSP = vlstrFechaInicio & "|" & vlstrFechaFin & "|" & Str(cboHospital.ItemData(cboHospital.ListIndex))
        
        Me.MousePointer = 11
        
        Set rsReporte = frsEjecuta_SP(vgstrParametrosSP, "sp_PvRptHonorarioFueraCorte")
        Set rsReporte = frsUltimoRecordset(rsReporte)
        If rsReporte.RecordCount > 0 Then
            alstrParametros(0) = "NombreHospital;" & Trim(cboHospital.List(cboHospital.ListIndex))
            alstrParametros(1) = "FechaInicio;" & CDate(mskFechaInicio.Text) & ";DATE"
            alstrParametros(2) = "FechaFin;" & CDate(mskFechaFin.Text) & ";DATE"
            alstrParametros(3) = "HoraInicio;" & CDate(mskHoraInicio.Text) & ";DATE"
            alstrParametros(4) = "HoraFin;" & CDate(mskHoraFin.Text) & ";DATE"
            
            pCargaParameterFields alstrParametros, vlrptReporte
    
            pImprimeReporte vlrptReporte, rsReporte, vlstrDestino, "Honorarios en efectivo fuera del corte"
        Else
            MsgBox SIHOMsg(236), vbOKOnly + vbInformation, "Mensaje"
        End If
        rsReporte.Close

        Me.MousePointer = 0
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pImprime"))
End Sub

Private Function fblnDatosValidos() As Boolean
    On Error GoTo NotificaError
    
    fblnDatosValidos = True
    
    If Not IsDate(mskFechaInicio.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
        mskFechaInicio.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskHoraInicio.Text) Then
        '¡Hora no válida!, formato de hora hh:mm
        MsgBox SIHOMsg(41), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
        mskHoraInicio.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskFechaFin.Text) Then
        '¡Fecha no válida!, formato de fecha dd/mm/aaaa
        MsgBox SIHOMsg(29), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
        mskFechaFin.SetFocus
    End If
    If fblnDatosValidos And Not IsDate(mskHoraFin.Text) Then
        '¡Hora no válida!, formato de hora hh:mm
        MsgBox SIHOMsg(41), vbOKOnly + vbCritical, "Mensaje"
        fblnDatosValidos = False
        mskHoraFin.SetFocus
    End If
    If fblnDatosValidos Then
        If CDate(mskFechaInicio.Text) > CDate(mskFechaFin.Text) Then
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbCritical, "Mensaje"
            fblnDatosValidos = False
            mskFechaInicio.SetFocus
        End If
        If fblnDatosValidos And CDate(mskFechaInicio.Text) = CDate(mskFechaFin.Text) And CDate(mskHoraInicio.Text) > CDate(mskHoraFin.Text) Then
            '¡Rango de fechas no válido!
            MsgBox SIHOMsg(64), vbOKOnly + vbCritical, "Mensaje"
            fblnDatosValidos = False
            mskFechaInicio.SetFocus
        End If
    End If
    
    

    Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":fblnDatosValidos"))
End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_KeyDown"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError

    Dim lngNumOpcion As Long
    Dim dtmfecha As Date

    Me.Icon = frmMenuPrincipal.Icon
    
    Select Case cgstrModulo
    Case "PV"
         lngNumOpcion = 931
    Case "SE"
         lngNumOpcion = 2005
    End Select
    
    pCargaHospital lngNumOpcion
    
    dtmfecha = fdtmServerFecha
    
    mskFechaInicio.Mask = ""
    mskFechaInicio.Text = dtmfecha
    mskFechaInicio.Mask = "##/##/####"
    
    mskFechaFin.Mask = ""
    mskFechaFin.Text = dtmfecha
    mskFechaFin.Mask = "##/##/####"

    mskHoraInicio.Mask = ""
    mskHoraInicio.Text = FormatDateTime(fdtmServerHora, vbShortTime)
    mskHoraInicio.Mask = "##:##"
    mskHoraInicio.Text = "00:00"
    
    mskHoraFin.Mask = ""
    mskHoraFin.Text = FormatDateTime(fdtmServerHora, vbShortTime)
    mskHoraFin.Mask = "##:##"
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":Form_Load"))
End Sub

Private Sub mskFechaFin_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_GotFocus"))
End Sub

Private Sub mskFechaFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        mskHoraFin.SetFocus
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_KeyDown"))
End Sub

Private Sub mskFechaFin_LostFocus()
    On Error GoTo NotificaError


    If Trim(mskFechaFin.ClipText) = "" Then
        mskFechaFin.Mask = ""
        mskFechaFin.Text = fdtmServerFecha
        mskFechaFin.Mask = "##/##/####"
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaFin_LostFocus"))
End Sub

Private Sub mskFechaInicio_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskFechaInicio


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_GotFocus"))
End Sub

Private Sub mskFechaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        mskHoraInicio.SetFocus
    End If
    

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_KeyDown"))
End Sub

Private Sub mskFechaInicio_LostFocus()
    On Error GoTo NotificaError


    If Trim(mskFechaInicio.ClipText) = "" Then
        mskFechaInicio.Mask = ""
        mskFechaInicio.Text = fdtmServerFecha
        mskFechaInicio.Mask = "##/##/####"
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskFechaInicio_LostFocus"))
End Sub

Private Sub mskHoraFin_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskHoraFin

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraFin_GotFocus"))
End Sub

Private Sub mskHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError


    If KeyCode = vbKeyReturn Then
        cmdPreview.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraFin_KeyDown"))
End Sub

Private Sub mskHoraFin_LostFocus()
    On Error GoTo NotificaError


    If Trim(mskHoraFin.ClipText) = "" Then
        mskHoraFin.Mask = ""
        mskHoraFin.Text = FormatDateTime(fdtmServerHora, vbShortTime)
        mskHoraFin.Mask = "##:##"
    End If


    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraFin_LostFocus"))
End Sub

Private Sub mskHoraInicio_GotFocus()
    On Error GoTo NotificaError


    pSelMkTexto mskHoraInicio

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraInicio_GotFocus"))
End Sub

Private Sub mskHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError

    
    If KeyCode = vbKeyReturn Then
        mskFechaFin.SetFocus
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraInicio_KeyDown"))
End Sub

Private Sub mskHoraInicio_LostFocus()
    On Error GoTo NotificaError


    If Trim(mskHoraInicio.ClipText) = "" Then
        mskHoraInicio.Mask = ""
        mskHoraInicio.Text = FormatDateTime(fdtmServerHora, vbShortTime)
        mskHoraInicio.Mask = "##:##"
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":mskHoraInicio_LostFocus"))
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

