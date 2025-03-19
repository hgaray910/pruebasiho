VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSociosBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de socios"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de socio"
      Height          =   630
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   4005
      Begin VB.OptionButton optTipoSoc 
         Caption         =   "Todos"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Búsqueda por todos los tipos de socio"
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optTipoSoc 
         Caption         =   "Dependiente"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Búsqueda por tipo de socio dependiente"
         Top             =   300
         Width           =   1245
      End
      Begin VB.OptionButton optTipoSoc 
         Caption         =   "Titular"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Búsqueda por tipo de socio titular"
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox txtBusqueda 
      Height          =   315
      Left            =   120
      MaxLength       =   30
      TabIndex        =   0
      ToolTipText     =   "Criterios para la búsqueda de socios"
      Top             =   840
      Width           =   9840
   End
   Begin VB.Frame fraParamentrosBusqueda 
      Caption         =   "Filtro de búsqueda"
      Height          =   630
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4605
      Begin VB.OptionButton optClave 
         Caption         =   "Clave"
         Height          =   195
         Index           =   1
         Left            =   3360
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Búsqueda por clave"
         Top             =   270
         Width           =   1080
      End
      Begin VB.OptionButton optNomSoc 
         Caption         =   "Paterno"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Búsqueda por apellido paterno"
         Top             =   270
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optNomSoc 
         Caption         =   "Materno"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Búsqueda por apellido materno"
         Top             =   270
         Width           =   885
      End
      Begin VB.OptionButton optNomSoc 
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Búsqueda por nombre"
         Top             =   270
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdhBuscaSocios 
      Height          =   2760
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Resultado de la búsqueda de socios"
      Top             =   1200
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   4868
      _Version        =   393216
      Cols            =   3
      GridColor       =   -2147483633
      Enabled         =   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
End
Attribute VB_Name = "frmSociosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vgblnAsignaTitular As Boolean
Public vgblnDependiente As Boolean
Public vlblnCambioTipoSocio As Boolean
Public vglngClaveSocio As Long
Public vgstrClaveUnica As String
Public vgstrNombreSocio As String
Public vglngNumeroCuenta As String
Public vgblnEscape As Boolean
Public vgstrCvesSocios As String

'Dim rsSocio As New ADODB.Recordset
Dim rsSocioTemp As New ADODB.Recordset
Dim vlstrsql As String


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
            'KeyCode = 0
            DoEvents
        Case vbKeyEscape
'            KeyCode = 0
            vlblnCambioTipoSocio = False
            'frmSocios.vlblnasignatitular *Checar que hace esta línea*
            vgblnAsignaTitular = False
            vgblnEscape = True
            Unload Me
    End Select
    
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pConfiguraGrid
    
    If vgblnAsignaTitular Then
        optTipoSoc(0).Value = True
        optTipoSoc(0).Enabled = False
        optTipoSoc(1).Enabled = False
        optTipoSoc(2).Enabled = False
    Else
        If vgblnDependiente And Not vgblnAsignaTitular Then
            If vlblnCambioTipoSocio Then
                optTipoSoc(0).Value = True
                optTipoSoc(0).Enabled = False
                optTipoSoc(1).Enabled = False
                optTipoSoc(2).Enabled = False
            Else
                optTipoSoc(1).Value = True
                optTipoSoc(0).Enabled = False
                optTipoSoc(1).Enabled = False
                optTipoSoc(2).Enabled = False
            End If
        Else
            If vlblnCambioTipoSocio Then
                optTipoSoc(1).Value = True
                optTipoSoc(0).Enabled = False
                optTipoSoc(1).Enabled = False
                optTipoSoc(2).Enabled = False
            Else
                optTipoSoc(0).Value = True
                optTipoSoc(0).Enabled = False
                optTipoSoc(1).Enabled = False
                optTipoSoc(2).Enabled = False
            End If
        End If
    End If
    
    vgblnEscape = False
    DoEvents
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
   vlblnCambioTipoSocio = False
   vgblnAsignaTitular = False
   vgblnEscape = True
End If
End Sub

Private Sub optClave_Click(Index As Integer)
pCargaSocios
End Sub

Private Sub optNomSoc_Click(Index As Integer)
pCargaSocios

End Sub

Private Sub txtBusqueda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       If rsSocioTemp.State = adStateOpen Then
          If rsSocioTemp.RecordCount > 0 And grdhBuscaSocios.Cols > 2 Then
              If fblnCanFocus(grdhBuscaSocios) Then grdhBuscaSocios.SetFocus
          Else
              If fblnCanFocus(txtBusqueda) Then txtBusqueda.SetFocus
          End If
       Else
            If fblnCanFocus(txtBusqueda) Then txtBusqueda.SetFocus
       End If
      
    End If
End Sub

Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtBusqueda_KeyUp(KeyCode As Integer, Shift As Integer)
    DoEvents
    pCargaSocios
End Sub
Private Sub pCargaSocios()
    '------------------------------------------------------------------------------
    'PROCEDIMIENTO QUE CARGA LOS Socios EN LA BUSQUEDA POR DIFERENTES PARAMETROS
    '------------------------------------------------------------------------------
    Dim vlstrParamtroSP As String
   ' Dim rsSocioTemp As New ADODB.Recordset
  
    On Error GoTo NotificaError
    
   'Se limpia la estructura del grid
   With grdhBuscaSocios
        .Clear
        .Cols = 2
        .Rows = 2
   End With
   pConfiguraGrid
   
   ' si no se tiene nada en el campo de busqueda no es necesario hacer ninguna busqueda
   If Trim(txtBusqueda.Text) = "" Then Exit Sub
   'Filtros para la búsqueda de los emplados
     
   If frmSocios.vlblnDependiente Then ' primero desde que pantalla se esta buscando( pantalla dependientes)
        vlstrParamtroSP = "'D'|"
   Else  ' desde pantalla de Titulares                                                                                                                                 ' los bisnietos no puede aparecer en los titulares
        vlstrParamtroSP = "'T'|"
   End If
         
   If optTipoSoc(0).Value Then ' ahora se debe identificar el tipo de socio a buscar(titulares)
        vlstrParamtroSP = vlstrParamtroSP & "'T'|"
   ElseIf optTipoSoc(1).Value Then ' Dependinetes
         vlstrParamtroSP = vlstrParamtroSP & "'D'|"
   Else  'siempre vas a buscar dependientes o socios,nunca ambos( por lo pronto)
        
   End If
   ' por ultimo se debe identificar el criterio de busqueda
   If optClave(1).Value Then  'por Clave del empleado
       vlstrParamtroSP = vlstrParamtroSP & "'CLAVE'|" & Trim(txtBusqueda.Text)
   ElseIf optNomSoc(0).Value Then 'por apellido paterno
       vlstrParamtroSP = vlstrParamtroSP & "'PATERNO'|" & Trim(txtBusqueda.Text)
   ElseIf optNomSoc(1).Value Then 'por apellido materno
       vlstrParamtroSP = vlstrParamtroSP & "'MATERNO'|" & Trim(txtBusqueda.Text)
   ElseIf optNomSoc(2).Value Then ' busqueda por nombre del Socio
       vlstrParamtroSP = vlstrParamtroSP & "'NOMBRE'|" & Trim(txtBusqueda.Text)
   End If
      
   ' se ejecuta el SP
   Set rsSocioTemp = frsEjecuta_SP(vlstrParamtroSP, "SP_PVSELBUSQUEDASOCIO")
      
      If rsSocioTemp.RecordCount > 0 Then
         pLlenarMshFGrdRs grdhBuscaSocios, rsSocioTemp
         pConfiguraGrid
         pFormatoFecha
      Else
         grdhBuscaSocios.Enabled = False
         If Trim(txtBusqueda.Text) = "" Then
            pEnfocaTextBox txtBusqueda
         End If
      End If
      
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pConsultaSocio"))
End Sub

Private Sub grdhBuscaSocios_DblClick()
   On Error GoTo NotificaError
 
   '|  No esta vacío el nombre del socio
    If Not grdhBuscaSocios.TextMatrix(grdhBuscaSocios.Row, 1) = "" Then
        vglngClaveSocio = CLng(grdhBuscaSocios.TextMatrix(grdhBuscaSocios.Row, 6))
        vgstrClaveUnica = CStr(grdhBuscaSocios.TextMatrix(grdhBuscaSocios.Row, 2))
        vgstrNombreSocio = CStr(grdhBuscaSocios.TextMatrix(grdhBuscaSocios.Row, 1))
        vglngNumeroCuenta = CLng(grdhBuscaSocios.TextMatrix(grdhBuscaSocios.Row, 7))
        
        vgstrCvesSocios = ""
        If rsSocioTemp.RecordCount > 0 Then
           ' obtenemos el primer registro
           rsSocioTemp.MoveFirst
           vgstrCvesSocios = rsSocioTemp!intcvesocio
           rsSocioTemp.MoveNext
           'recorremos el resto del recordset para obtener las claves de los socios cargados
           Do While Not rsSocioTemp.EOF
              vgstrCvesSocios = vgstrCvesSocios & "," & rsSocioTemp!intcvesocio
               rsSocioTemp.MoveNext
           Loop
        End If
        
        Hide
        vgblnEscape = False
    Else
        pConfiguraGrid
    End If
             
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":DblClick"))
End Sub

Private Sub grdhBuscaSocios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        grdhBuscaSocios_DblClick
        KeyCode = 0
    End If
End Sub
Private Sub pConfiguraGrid()
    On Error GoTo NotificaError
    
    'Configura el grid de la búsqueda de Socios
    Dim vlintseq As Integer
    
    grdhBuscaSocios.Enabled = True
    With grdhBuscaSocios
        .FormatString = "Clave|Nombre|Clave única|Hispanidad|Fecha de ingreso|Fecha de baja||"
        .ColWidth(0) = 100
        .ColWidth(1) = 3600
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        '.Col(5) = Format(.TextMatrix(.Row, 5), "dd mmm yyyy")
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ScrollBars = flexScrollBarBoth
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGrid"))
End Sub
Private Sub pFormatoFecha()
    Dim lRow As Long
    '-- Formatear las fechas 'dd/mmm/yyyy'
    With grdhBuscaSocios
        .Redraw = False
        For lRow = .FixedRows To .Rows - 1
            .TextMatrix(lRow, 4) = Format(.TextMatrix(lRow, 4), "dd/mmm/yyyy")
            .TextMatrix(lRow, 5) = Format(.TextMatrix(lRow, 5), "dd/mmm/yyyy")
        Next lRow
        .Redraw = True
    End With
End Sub

