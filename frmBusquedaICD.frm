VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaICD 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B�squeda de diagn�sticos ICD"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCarga 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8565
      Top             =   3405
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7680
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   10190
      Begin VB.TextBox txtIniciales 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2460
         TabIndex        =   0
         ToolTipText     =   "Escriba las iniciales"
         Top             =   300
         Width           =   7590
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDiagnosticos 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Lista de diagn�sticos ICD"
         Top             =   705
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   12091
         _Version        =   393216
         ForeColor       =   0
         Cols            =   4
         FixedCols       =   0
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Escriba las iniciales"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmBusquedaICD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vlfrmForma As Form
Public vgstrTipoICD As String 'Indica el tipo de c�digo ICD que se consultar� (US = Urgencia sentida, UR = Urgencia real)
Dim vlblnPrimeraVez As Boolean
Dim rsTemp As New ADODB.Recordset
Dim vlstrExpresion As String

Private Sub grdDiagnosticos_DblClick()
    On Error GoTo NotificaError
    
    If grdDiagnosticos.TextMatrix(grdDiagnosticos.Row, 1) <> "" Then pAsigna

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDiagnosticos_DblClick"))
End Sub

Private Sub pAsigna()
    On Error GoTo NotificaError

     With grdDiagnosticos

         vlfrmForma.vgstrCodigoICD = Trim(.TextMatrix(.Row, 1))
         vlfrmForma.vgstrDescDiagnostico = Trim(.TextMatrix(.Row, 2))

     End With
    
    Unload Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsigna"))
End Sub

Private Sub grdDiagnosticos_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyLeft Then
        txtIniciales.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDiagnosticos_KeyDown"))
End Sub

Private Sub grdDiagnosticos_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(grdDiagnosticos.TextMatrix(1, 1)) <> "" Then
            pAsigna
        End If
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdDiagnosticos_KeyPress"))
End Sub

Private Sub tmrCarga_Timer()
    On Error GoTo NotificaError
    
    PSuperBusqueda (vgstrTipoICD)
    tmrCarga.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrCarga_Timer"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        vlfrmForma.vgstrCodigoICD = 0
        vlfrmForma.vgstrDescDiagnostico = ""
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    vlblnPrimeraVez = True

    pLimpiaGrid
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub txtIniciales_GotFocus()
    On Error GoTo NotificaError
    
    pSelTextBox txtIniciales

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIniciales_GotFocus"))
End Sub

Private Sub txtIniciales_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyReturn Then
        If grdDiagnosticos.TextMatrix(1, 1) <> "" Then
            grdDiagnosticos.SetFocus
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIniciales_KeyDown"))
End Sub

Private Sub txtIniciales_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    tmrCarga.Enabled = False
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIniciales_KeyPress"))
End Sub

Private Sub txtIniciales_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If txtIniciales.Text = "" Then
        pLimpiaGrid
    Else
        tmrCarga.Enabled = True
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":txtIniciales_KeyUp"))
End Sub

Private Sub pLimpiaGrid()
    On Error GoTo NotificaError
    
    Dim x As Long
    
    With grdDiagnosticos
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 1
        
        For x = 1 To grdDiagnosticos.Cols - 1
            .ColAlignmentFixed(x) = flexAlignCenterCenter
            .TextMatrix(1, x) = ""
        Next x
        
        .FormatString = "|C�digo ICD|Descripci�n del diagn�stico|"
        
        .ColWidth(0) = 100
        .ColWidth(1) = 1400     'C�digo ICD
        .ColAlignment(1) = flexAlignLeftCenter
        .ColWidth(2) = 8300     'Descripci�n del diagn�stico
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(3) = 0        'Clave Diagn�stico
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
End Sub

Sub PSuperBusqueda(strTipoICD As String)
    On Error GoTo NotificaError
    
    Dim vlstrInstruccion As String
    Dim vlintRenglones As Integer
    Dim vlintColumnas As Integer
    Dim rsDatos As New ADODB.Recordset
    
    Dim vlstrResultado As String
    Dim vlintInicio As Integer
    Dim vlintCont As Integer
    Dim vlblnBand As Boolean
      
    'Se especifica el query seg�n el tipo de ICD
    If vgstrTipoICD = "UR" Then
        vlstrInstruccion = "SELECT VCHCODIGOICD, VCHDESCRIPCION, INTCVECODIGOICD FROM GNCODIGOICDAXA WHERE BITACTIVO = 1 AND vchDescripcion LIKE '" & Trim(txtIniciales.Text) & "%' AND VCHTIPO = 'UR' ORDER BY GNCODIGOICDAXA.vchDescripcion"
    ElseIf vgstrTipoICD = "US" Then
        vlstrInstruccion = "SELECT VCHCODIGOICD, VCHDESCRIPCION, INTCVECODIGOICD FROM GNCODIGOICDAXA WHERE BITACTIVO = 1 AND vchDescripcion LIKE '" & Trim(txtIniciales.Text) & "%' AND VCHTIPO = 'US' ORDER BY GNCODIGOICDAXA.vchDescripcion"
    End If
    
    
    grdDiagnosticos.Redraw = False
    
    Set rsDatos = frsRegresaRs(vlstrInstruccion, adLockOptimistic, adOpenDynamic)
    
    pLimpiaGrid
    
    grdDiagnosticos.Rows = IIf(rsDatos.RecordCount = 0, 2, rsDatos.RecordCount + 1)
    With grdDiagnosticos
       For vlintRenglones = 1 To rsDatos.RecordCount
            For vlintColumnas = 0 To rsDatos.Fields.Count - 1
                If IsNull(rsDatos.Fields(vlintColumnas).Value) Then
                    .TextMatrix(vlintRenglones, vlintColumnas + 1) = ""
                Else
                    .TextMatrix(vlintRenglones, vlintColumnas + 1) = rsDatos.Fields(vlintColumnas).Value
                End If
            Next vlintColumnas
             rsDatos.MoveNext
       Next vlintRenglones
    End With

    grdDiagnosticos.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":PSuperBusqueda"))
End Sub

