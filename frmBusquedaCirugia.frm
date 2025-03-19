VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaCirugia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Cirugias"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCarga 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5880
      Top             =   3720
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
      Height          =   6480
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   8385
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
         Left            =   2590
         TabIndex        =   0
         ToolTipText     =   "Escriba las iniciales"
         Top             =   300
         Width           =   5655
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCirugias 
         Height          =   5535
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Lista de cirugías"
         Top             =   800
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   9763
         _Version        =   393216
         ForeColor       =   0
         FixedCols       =   0
         ForeColorFixed  =   0
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorUnpopulated=   -2147483643
         GridColor       =   -2147483638
         GridColorFixed  =   -2147483638
         GridColorUnpopulated=   -2147483638
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         GridLinesUnpopulated=   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "|Nombre de la cirugía"
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
         _Band(0).BandIndent=   3
         _Band(0).Cols   =   2
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
         Left            =   180
         TabIndex        =   3
         Top             =   360
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmBusquedaCirugia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Agregué el icono

'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Expediente
'| Nombre del Formulario    : frmBusquedaCirugias
'-------------------------------------------------------------------------------------
'| Objetivo: Tener una SuperBusqueda de Pacientes egresados
'-------------------------------------------------------------------------------------
'| Fecha de Creación        :
'| Modificó                 : Nombre(s)
'| Fecha Terminación        :
'| Fecha última modificación:
'-------------------------------------------------------------------------------------

Public vlfrmForma As Form

Dim vlblnPrimeraVez As Boolean
Dim rsTemp As New ADODB.Recordset
Dim vlstrExpresion As String



Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
    vlblnPrimeraVez = False
End Sub

Private Sub grdCirugias_DblClick()
    On Error GoTo NotificaError
    
    If grdCirugias.TextMatrix(grdCirugias.Row, 1) <> "" Then pAsigna

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCirugias_DblClick"))
End Sub

Private Sub pAsigna()
    On Error GoTo NotificaError

     With grdCirugias

         vlfrmForma.vglngCveCirugia = CLng(.TextMatrix(.Row, 2))
         vlfrmForma.vgstrDescCirugia = .TextMatrix(.Row, 1)

     End With
    
    Unload Me
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pAsigna"))
End Sub

Private Sub grdCirugias_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    
    If KeyCode = vbKeyLeft Then
        txtIniciales.SetFocus
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCirugias_KeyDown"))
End Sub

Private Sub grdCirugias_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        If Trim(grdCirugias.TextMatrix(1, 1)) <> "" Then
            pAsigna
        End If
    End If
        
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdCirugias_KeyPress"))
End Sub

Private Sub tmrCarga_Timer()
    On Error GoTo NotificaError
    
    PSuperBusqueda
    tmrCarga.Enabled = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":tmrCarga_Timer"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        vlfrmForma.vglngCveDiagnostico = 0
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
        If grdCirugias.TextMatrix(1, 1) <> "" Then
            grdCirugias.SetFocus
        Else
            cboClasificacion.SetFocus
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
    
    Dim X As Long
    
    With grdCirugias
        .Rows = 2
        .Cols = 3
        .FixedRows = 1
        .FixedCols = 1
        
        For X = 1 To grdCirugias.Cols - 1
            .ColAlignmentFixed(X) = flexAlignCenterCenter
            .TextMatrix(1, X) = ""
        Next X
        
        .FormatString = "|Nombre de la cirugía"
        
        .ColWidth(0) = 100
        .ColWidth(1) = 7650     'Nombre de la cirugia
        .ColAlignment(1) = flexAlignLeftCenter
'        .ColWidth(2) = 2300     'Clasificación
'        .ColAlignment(2) = flexAlignLeftCenter
'        .ColWidth(3) = 2300     'Tipo de Clasificación (Diagnóstico)
'        .ColAlignment(3) = flexAlignLeftCenter
        .ColWidth(2) = 0        'Clave Diagnóstico
    End With

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pLimpiaGrid"))
End Sub

Sub PSuperBusqueda()
    On Error GoTo NotificaError
    
    Dim vlstrInstruccion As String
    Dim vlintRenglones As Integer
    Dim vlintColumnas As Integer
    Dim rsDatos As New ADODB.Recordset
    
    Dim vlstrResultado As String
    Dim vlintInicio As Integer
    Dim vlIntCont As Integer
    Dim vlblnBand As Boolean
    
    If vlblnPrimeraVez Then
        Exit Sub
    End If
    

        vlstrInstruccion = "SELECT ExCirugia.vchDescripcion ,ExCirugia.INTCVECIRUGIA FROM ExCirugia WHERE ExCirugia.BITACTIVA = 1 AND ExCirugia.vchDescripcion LIKE '" & Trim(txtIniciales.Text) & "%'  ORDER BY ExCirugia.vchDescripcion"
 
    
    grdCirugias.Redraw = False
    
    Set rsDatos = frsRegresaRs(vlstrInstruccion, adLockOptimistic, adOpenDynamic)
    
    pLimpiaGrid
    
    grdCirugias.Rows = IIf(rsDatos.RecordCount = 0, 2, rsDatos.RecordCount + 1)
    With grdCirugias
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

    grdCirugias.Redraw = True

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":PSuperBusqueda"))
End Sub

