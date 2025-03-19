VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBusquedaFacturasPreviasVP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Facturas previas"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3245
      TabIndex        =   2
      ToolTipText     =   "Confirmar la selección de las facturas"
      Top             =   3840
      Width           =   1560
   End
   Begin VB.CommandButton cmdInvertir 
      Caption         =   "Invertir selección"
      Height          =   495
      Left            =   1670
      TabIndex        =   1
      ToolTipText     =   "Invertir selección"
      Top             =   3840
      Width           =   1560
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdFacturas 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Factura previamente cancelada a la cual sustituye"
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      GridColor       =   -2147483633
      FocusRect       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmBusquedaFacturasPreviasVP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------
' Facturas previas
'----------------------------------------------------------------------------------
Option Explicit

Public vlchrfoliofactura As String
Public vlstrsql As String

Private Sub cmdAceptar_Click()
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
    Dim vllngConsecutivoFacturas As Long

    If grdFacturas.Rows > 0 Then
    
        ReDim aFoliosPrevios(0)
        vlchrfoliofactura = ""
        vllngConsecutivoFacturas = 1
        
        With grdFacturas
            vlintnumrenglon = .Row
            
            For vllngContador = 1 To .Rows - 1
                If .TextMatrix(vllngContador, 0) = "*" Then
                    If vlchrfoliofactura = "" Then
                    
                        ReDim Preserve aFoliosPrevios(vllngConsecutivoFacturas)
                        aFoliosPrevios(vllngConsecutivoFacturas).chrfoliofactura = Trim(grdFacturas.TextMatrix(vllngContador, 1))
                            
                        vlchrfoliofactura = Trim(grdFacturas.TextMatrix(vllngContador, 1))
                    Else
                        vllngConsecutivoFacturas = vllngConsecutivoFacturas + 1
                    
                        ReDim Preserve aFoliosPrevios(vllngConsecutivoFacturas)
                        aFoliosPrevios(vllngConsecutivoFacturas).chrfoliofactura = Trim(grdFacturas.TextMatrix(vllngContador, 1))
                    
                        vlchrfoliofactura = vlchrfoliofactura & ", " & Trim(grdFacturas.TextMatrix(vllngContador, 1))
                    End If
                End If
            Next vllngContador
            
            .RowSel = vlintnumrenglon
        End With
            
        If vlchrfoliofactura <> "" Then
            Unload Me
        Else
            MsgBox SIHOMsg(1626), vbOKOnly + vbInformation, "Mensaje"
            grdFacturas.SetFocus
        End If
    End If
End Sub

Private Sub cmdAceptar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdInvertir_Click()
    pInvertirSeleccion
End Sub

Private Sub cmdInvertir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
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
        ReDim aFoliosPrevios(0)
        pCarga
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = False
End Sub

Private Sub grdFacturas_Click()
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
        .Cols = 6
        For vllngContador = 1 To .Cols - 1
            .TextMatrix(1, vllngContador) = ""
        Next vllngContador
    End With
    
   
    
    Set rsFacturas = frsRegresaRs(vlstrsql, adLockOptimistic, adOpenDynamic)
    If rsFacturas.RecordCount <> 0 Then
        Do While Not rsFacturas.EOF
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 1) = rsFacturas!chrfoliofactura
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 2) = rsFacturas!DTMFECHAHORA
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 3) = rsFacturas!mnyTotalFactura
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 4) = rsFacturas!PESOS
            grdFacturas.TextMatrix(grdFacturas.Rows - 1, 5) = rsFacturas!MNYTIPOCAMBIO
            rsFacturas.MoveNext
            If Not rsFacturas.EOF Then grdFacturas.Rows = grdFacturas.Rows + 1
        Loop
    End If
    rsFacturas.Close
    
    With grdFacturas
        .FixedCols = 1
        .FixedRows = 1
        .FormatString = "|Folio|Fecha|Total|Moneda|Tipo de cambio"
        .ColWidth(0) = 150
        .ColWidth(1) = 1500
        .ColWidth(2) = 2100
        .ColWidth(3) = 1200
        .ColWidth(4) = 800
        .ColWidth(5) = 0
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignRightCenter
        
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(5) = flexAlignCenterCenter
        
        pFormatoNumeroColumnaGrid grdFacturas, 3, "$ "
        
        pFormatoFechaLargaColumnaGrid grdFacturas, 2
        
        .ScrollBars = flexScrollBarBoth
        
        .Col = 1
        .Row = 1
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCarga"))
End Sub

Private Sub pSeleccionaRenglon()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
     
    With grdFacturas
        vlintnumrenglon = .Row
        
        If .TextMatrix(vlintnumrenglon, 0) = "*" Then
            .TextMatrix(vlintnumrenglon, 0) = ""
        Else
            .TextMatrix(vlintnumrenglon, 0) = "*"
        End If
        
        .RowSel = vlintnumrenglon
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaRenglon"))
End Sub

Private Sub pInvertirSeleccion()
    On Error GoTo NotificaError
    
    Dim vlintnumrenglon As Integer
    Dim vllngContador As Integer
     
    With grdFacturas
        vlintnumrenglon = .Row
        
        For vllngContador = 1 To .Rows - 1
            If .TextMatrix(vllngContador, 0) = "" Then
                .TextMatrix(vllngContador, 0) = "*"
            Else
                .TextMatrix(vllngContador, 0) = ""
            End If
        Next vllngContador
        
        .RowSel = vlintnumrenglon
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pInvertirSeleccion"))
End Sub

Private Sub pFormatoFechaLargaColumnaGrid(grdNombre As MSHFlexGrid, vlintxColumna As Integer, Optional vlstrSigno As String)
'----------------------------------------------------------------------
' Procedimiento para dar formato a la columna del grid que son Fechas
'----------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim X As Long
    
    For X = 1 To grdNombre.Rows - 1
        grdNombre.TextMatrix(X, vlintxColumna) = Format(grdNombre.TextMatrix(X, vlintxColumna), "DD/MMM/YYYY")
    Next X

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pFormatoFechaColumnaGrid"))
End Sub
