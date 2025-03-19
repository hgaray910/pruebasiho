VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVentaPublicoNoFacturado 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Folios antes de venta"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFacturas 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Folios antes de venta"
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   -2147483633
      FocusRect       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmVentaPublicoNoFacturado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Folios previos a la factura en venta al publico
'----------------------------------------------------------------------------------
Option Explicit

Public vlchrfoliofactura As String
Public vlstrsql As String

Private Sub pSeleccionaRenglon()
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
    Dim vllngConsecutivoFacturas As Long

    If grdFacturas.Rows > 0 Then
    
        ReDim aFolios(0)
        vlchrfoliofactura = ""
        vllngConsecutivoFacturas = 1
        
        With grdFacturas
            vlintnumrenglon = .Row
            
            frmPOS.txtTicketPrevio.Text = Trim(grdFacturas.TextMatrix(vlintnumrenglon, 1))
                                        
            vlchrfoliofactura = Trim(grdFacturas.TextMatrix(vlintnumrenglon, 1))
            
            .RowSel = vlintnumrenglon
        End With
            
        If vlchrfoliofactura <> "" Then
            Unload Me
        Else
            'Seleccione la factura.
            MsgBox SIHOMsg(217), vbOKOnly + vbInformation, "Mensaje"
            grdFacturas.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
        If KeyAscii = 27 Then Unload Me
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
        vgstrNombreForm = Me.Name
        vlchrfoliofactura = ""
        ReDim aFolios(0)
        pCarga
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = False
End Sub

Private Sub grdFacturas_DblClick()
    pSeleccionaRenglon
End Sub

Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then pSeleccionaRenglon
End Sub

Private Sub pCarga()
    On Error GoTo NotificaError
    Dim rsFacturas As New ADODB.Recordset
    Dim vllngContador As Long
      
    With grdFacturas
        .Rows = 2
        .Cols = 5
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    End With

    Set rsFacturas = frsEjecuta_SP(Str(vgintClaveEmpresaContable), "SP_PVSELVENTASPUBLICONOFACT")
    If rsFacturas.RecordCount <> 0 Then
        grdFacturas.Redraw = False
        
        Do While Not rsFacturas.EOF
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 1) = rsFacturas!INTCVEVENTA
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 2) = rsFacturas!DEPTO
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 3) = Format(rsFacturas!Fecha, "dd/mmm/yyyy hh:mm")
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 4) = rsFacturas!Total
        
            rsFacturas.MoveNext
            If Not rsFacturas.EOF Then grdFacturas.Rows = grdFacturas.Rows + 1
        Loop
        
        grdFacturas.Redraw = True
    Else
        Unload Me
    End If
    rsFacturas.Close
    
    grdFacturas.Redraw = False
    
    With grdFacturas
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Folio|Departamento|Fecha|Total"
        .ColWidth(0) = 150
        .ColWidth(1) = 800
        .ColWidth(2) = 2500
        .ColWidth(3) = 1600
        .ColWidth(4) = 1100
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        
        pFormatoNumeroColumnaGrid grdFacturas, 4, "$ "
        
        .ScrollBars = flexScrollBarBoth
        
        .Col = 1
        .Row = 1
    End With
    
    grdFacturas.Redraw = True
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCarga"))
End Sub
