VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFormatosRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formatos de recibo"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   105
      TabIndex        =   4
      Top             =   60
      Width           =   8655
      Begin VB.CommandButton cmdSelecciona 
         Height          =   540
         Left            =   7770
         MaskColor       =   &H80000014&
         Picture         =   "frmFormatosRecibos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Incluir el recibo capturado"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.TextBox txtNombreFormato 
         Height          =   285
         Left            =   1845
         MaxLength       =   100
         TabIndex        =   0
         ToolTipText     =   "Descripción del formato de recibo"
         Top             =   270
         Width           =   5835
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción del formato"
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   285
         Width           =   1680
      End
   End
   Begin MSComDlg.CommonDialog cmdArchivo 
      Left            =   9105
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formatos dados de alta"
      Height          =   3135
      Left            =   105
      TabIndex        =   3
      Top             =   885
      Width           =   10140
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFormatos 
         Height          =   2655
         Left            =   165
         TabIndex        =   2
         ToolTipText     =   "Listado de recibos dados de alta. DobleClick para eliminar un formato."
         Top             =   285
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   4683
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmFormatosRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmFormatosRecibos                                     -
'-------------------------------------------------------------------------------------
'| Objetivo: parametrizar formatos de impresión con reportes de crystal utilizados
'|           por la pantalla de entradas y salidas de Dinero
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G.
'| Autor                    : Rodolfo Ramos G.
'| Fecha de Creación        : 12/Abr/2006
'| Modificó                 : Nombre(s)
'| Fecha Terminación        : mesmo día...
'| Fecha última modificación:
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
' Ultimas modificaciones, especificar:
'
'------------------------------------------------------------------------------------------

Option Explicit

Private Sub pllenaLista()
    Dim strSentencia As String
    Dim rsFormatos As New ADODB.recordSet
    
    strSentencia = "select * from pvFormatoRecibo order by intIdFormato"
    Set rsFormatos = frsRegresaRs(strSentencia, adLockOptimistic, adOpenForwardOnly)
    
    'Inicialización del Grid
    pLimpiaGrid grdFormatos
    
    With grdFormatos
        .RowData(1) = 0
        .Row = 1
        Do While Not rsFormatos.EOF
            If .RowData(1) > 0 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
            .RowData(.Row) = rsFormatos!intIdFormato
            .TextMatrix(.Row, 1) = rsFormatos!vchDescripcion
            .TextMatrix(.Row, 2) = rsFormatos!vchReporte
            rsFormatos.MoveNext
        Loop
    End With
    
End Sub

Private Sub cmdSelecciona_Click()
    Dim strArchivo As String
    Dim strSentencia As String
    
    If Trim(txtNombreFormato.Text) <> "" Then
        cmdArchivo.FileName = "*.rpt"
        cmdArchivo.ShowOpen
        If cmdArchivo.FileName <> "" And cmdArchivo.FileName <> "*.rpt" Then
            If UCase(Mid(cmdArchivo.FileTitle, Len(cmdArchivo.FileTitle) - 2, 3)) = "RPT" Then
                strSentencia = "insert into pvFormatoRecibo (VCHDESCRIPCION, VCHREPORTE) " & _
                                "Values ('" & Trim(txtNombreFormato.Text) & "','" & cmdArchivo.FileTitle & "')"
                
                pEjecutaSentencia strSentencia
                pllenaLista
                txtNombreFormato.Text = ""
                pEnfocaTextBox txtNombreFormato
            Else
                'Tipo de archivo no válido.
                MsgBox SIHOMsg(716), vbOKOnly + vbCritical, "Mensaje"
            End If
        End If
    End If
End Sub

Private Sub grdFormatos_DblClick()
    Dim strSentencia As String
    
    If grdFormatos.RowData(grdFormatos.Row) > 0 Then
        '¿Está seguro de eliminar los datos?
        If MsgBox(SIHOMsg(6), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            strSentencia = "Delete pvFormatoRecibo where intIdFormato = " & grdFormatos.RowData(grdFormatos.Row)
            pEjecutaSentencia strSentencia
            pllenaLista
        End If
    End If

End Sub



Private Sub Form_Activate()
    Me.Icon = frmMenuPrincipal.Icon
    pllenaLista
End Sub
Private Sub pLimpiaGrid(grdGrid As MSHFlexGrid)
    
    With grdGrid
        'Inicializada...
        .Rows = 2
        .Cols = 3
        .Clear
        
        'Configurada
        .FormatString = "|Nombre|Reporte"
        .ColWidth(0) = 100
        .ColWidth(1) = 3500 'Nombre
        .ColWidth(2) = 6000 'Ruta
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftBottom
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .Redraw = True
        .Visible = True
    End With
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub txtNombreFormato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSelecciona.SetFocus
    End If
End Sub

Private Sub txtNombreFormato_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
