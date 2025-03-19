VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmListasPreciosConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de incremento en listas de precios"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   350
      Left            =   4448
      TabIndex        =   8
      ToolTipText     =   "Aceptar"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listas de precios"
      Height          =   5015
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10100
      Begin VB.CommandButton cmdInvertirSeleccion 
         Caption         =   "Invertir selección"
         Height          =   350
         Left            =   8000
         TabIndex        =   1
         ToolTipText     =   "Invertir selección"
         Top             =   3770
         Width           =   1830
      End
      Begin VB.TextBox txtEditCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         MaxLength       =   15
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ListBox UpDown1 
         Appearance      =   0  'Flat
         Height          =   615
         ItemData        =   "frmListasPreciosConcepto.frx":0000
         Left            =   240
         List            =   "frmListasPreciosConcepto.frx":000D
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   "Modificar listas"
         Height          =   1185
         Left            =   120
         TabIndex        =   9
         Top             =   3690
         Width           =   7595
         Begin VB.CheckBox chkIncrementoAutomatico 
            Caption         =   "Incremento automático"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Incremento automático"
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdAplicar 
            Caption         =   "Aplicar"
            Height          =   350
            Left            =   5760
            TabIndex        =   7
            ToolTipText     =   "Aplicar cambios a las listas de precios seleccionadas"
            Top             =   705
            Width           =   1560
         End
         Begin VB.CheckBox chkUsarTabulador 
            Caption         =   "Usar tabulador"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            ToolTipText     =   "Usar tabulador"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            TabIndex        =   4
            ToolTipText     =   "Precio inicial"
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtMargenUtilidad 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2880
            TabIndex        =   3
            ToolTipText     =   "Porcentaje de margen de utilidad"
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cboTipoIncremento 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Tipo de incremento que usará"
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label4 
            Caption         =   "Precio"
            Height          =   195
            Left            =   4320
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Margen de utilidad"
            Height          =   195
            Left            =   2880
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de incremento"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1455
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdPrecios 
         Height          =   3385
         Left            =   140
         TabIndex        =   21
         ToolTipText     =   "Captura del tipo de incremento y precio inicial por lista de precios"
         Top             =   240
         Width           =   9800
         _ExtentX        =   17277
         _ExtentY        =   5980
         _Version        =   393216
         GridColor       =   12632256
         WordWrap        =   -1  'True
         HighLight       =   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Precios no asignados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8000
         TabIndex        =   20
         Top             =   4580
         Width           =   1830
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Lista predeterminada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8000
         TabIndex        =   19
         Top             =   4270
         Width           =   1830
      End
   End
   Begin VB.Label lblNombreConcepto 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      ToolTipText     =   "Nombre del concepto"
      Top             =   525
      Width           =   8100
   End
   Begin VB.Label lblClaveConcepto 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1680
      TabIndex        =   15
      ToolTipText     =   "Clave del concepto"
      Top             =   120
      Width           =   1860
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   585
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Clave"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   180
      Width           =   405
   End
End
Attribute VB_Name = "frmListasPreciosConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vgblnNoEditar As Boolean
Dim llngMarcados As Long
Public vlblnCambiaPrecio As Boolean ' permiso que indica si se puede cambiar el precio cuando la politica de precios es precio maximo al publico
Public vgblnCancel As Boolean ' para ver si se requiere actualizar las listas de precios que se tienen en el arreglo
Dim vlblnEscTxtEditCOl As Boolean ' para matar proceso cuando se de escape desde el txt para editar las columnas del grid por que el key preview del form lo cacha

Const cintcolgrdSeleccion = 0   'Fix                    ' seleccion 00
Const cintcolgrdClaveLista = 1                          ' clave de la lista de precios 01
Const cintcolgrdDescripcionLista = 2                       ' lista de precios 02
Const cintcolgrdIncrementoAuto = 3                          ' incremento automatico 03
Const cintcolgrdIncremento = 4                         ' tipo incremento 04
Const cintcolgrdMargenUt = 5                          ' margen de utilidad 05
Const cintcolgrdTabulador = 6                          ' usar tabulador 06
Const cintcolgrdPrecio = 7                         ' precio 08
Const cintcolgrdPredeterminada = 8                           ' lista predeterminada 12
Const cintcolgrdNuevo = 9                           ' nuevo en la lista de precios 13

Private Sub cboTipoIncremento_Click()
    pTxtEditCOlPierdeFoco
End Sub

Private Sub chkIncrementoAutomatico_GotFocus()
    pTxtEditCOlPierdeFoco
End Sub

Private Sub chkUsarTabulador_GotFocus()
    pTxtEditCOlPierdeFoco
End Sub

Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

Private Sub pTxtEditCOlPierdeFoco()
    If txtEditCol.Visible = True Then
        pSetCellValueCol grdPrecios, txtEditCol
    End If
End Sub

Private Sub cmdAceptar_GotFocus()
pTxtEditCOlPierdeFoco
End Sub

Private Sub cmdAplicar_Click()
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 costo base
    ' 08 precio
    ' 09 precio maximo al publico
    ' 10 precio costo mas alto
    ' 11 precio ultima entrada

    Dim lngContador As Long
 
    'If Me.cboTipoIncremento.ListIndex = 2 Then
    '    If MsgBox(SIHOMsg(1220), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            'pPoliticaPrecioMaximo (Replace(Me.txtCostoBase.Text, "$", ""))
    '    Else
    '        Exit Sub
    '    End If
    'End If
 
    For lngContador = 1 To grdPrecios.Rows - 1
        If Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*" Then
            grdPrecios.TextMatrix(lngContador, cintcolgrdIncrementoAuto) = IIf(Me.chkIncrementoAutomatico.Value = 1, "*", "")
            grdPrecios.TextMatrix(lngContador, cintcolgrdIncremento) = IIf(cboTipoIncremento.ListIndex = 0, "ÚLTIMA", IIf(cboTipoIncremento.ListIndex = 1, "COMPRA", "PRECIO"))
            grdPrecios.TextMatrix(lngContador, cintcolgrdMargenUt) = Format(Replace(Me.txtMargenUtilidad.Text, "%", ""), "0.0000") & "%"
            grdPrecios.TextMatrix(lngContador, cintcolgrdTabulador) = IIf(Me.chkUsarTabulador.Value = 1, "*", "")
            grdPrecios.TextMatrix(lngContador, cintcolgrdPrecio) = FormatCurrency(Replace(Me.txtPrecio.Text, "$", ""), 2)

        End If
    Next lngContador
    
End Sub

Private Sub cmdAplicar_GotFocus()
    pTxtEditCOlPierdeFoco
End Sub

Private Sub cmdInvertirSeleccion_Click()
    On Error GoTo NotificaError
    Dim lngContador As Long
    
    For lngContador = 1 To grdPrecios.Rows - 1
        grdPrecios.TextMatrix(lngContador, 0) = IIf(Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*", "", "*")
        
        If Trim(grdPrecios.TextMatrix(lngContador, 0)) = "*" Then
            llngMarcados = llngMarcados + 1
        Else
            llngMarcados = llngMarcados - 1
        End If
    Next lngContador

    pHabilitaModificar
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":cmdInvertirSeleccion_Click"))
End Sub

Private Sub pHabilitaModificar()
    On Error GoTo NotificaError

    cboTipoIncremento.Enabled = llngMarcados <> 0
    txtMargenUtilidad.Enabled = llngMarcados <> 0
    txtPrecio.Enabled = llngMarcados <> 0
    chkIncrementoAutomatico.Enabled = llngMarcados <> 0
    chkUsarTabulador.Enabled = llngMarcados <> 0
    cmdAplicar.Enabled = llngMarcados <> 0
        
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pHabilitaModificar"))
End Sub

Private Sub cmdInvertirSeleccion_GotFocus()
    pTxtEditCOlPierdeFoco
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name = "txtEditCol" Or Me.ActiveControl.Name = "UpDown1" Then
        Exit Sub
    End If
  
    If vlblnEscTxtEditCOl = True Then
        vlblnEscTxtEditCOl = False
        KeyAscii = 0
        Exit Sub
    End If
  
    If KeyAscii = 27 Then
        KeyAscii = 0
        vgblnCancel = True
        Unload Me
    ElseIf KeyAscii = 13 Then
        If Me.ActiveControl.Name <> "grdPrecios" Then
            SendKeys vbTab
        End If
    End If
End Sub
Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    llngMarcados = 0 ' se inicializa las listas marcadas
    pHabilitaModificar
End Sub
Public Sub pEditarColumna(KeyAscii As Integer, txtEdit As TextBox, grid As MSHFlexGrid)
    On Error GoTo NotificaError
    Dim vlintTexto As Integer

    With txtEdit
        .Text = Replace(grid, "%", "") 'Inicialización del Textbox
        Select Case KeyAscii
            Case 0 To 32
                'Edita el texto de la celda en la que está posicionado
                .SelStart = 0
                .SelLength = 1000
            Case 8, 48 To 57
                ' Reemplaza el texto actual solo si se teclean números
                vlintTexto = Chr(KeyAscii)
                .Text = vlintTexto
                .SelStart = 1
            Case 46
                ' Reemplaza el texto actual solo si se teclean números
                .Text = "."
                .SelStart = 1
        End Select
    End With
            
    ' Muestra el textbox en el lugar indicado
    With grid
        If .CellWidth < 0 Then Exit Sub
        txtEdit.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 8, .CellHeight - 8
    End With
    
    vgstrEstadoManto = vgstrEstadoManto & "E"
    txtEdit.Visible = True
    txtEdit.SetFocus
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pEditarColumna"))
End Sub

Private Sub pSetCellValueCol(grid As MSHFlexGrid, txtEdit As TextBox)
    On Error GoTo NotificaError

    ' NOTA:
    '       Este código debe ser  llamado cada vez que
    '       el grid pierde el foco y su contenido puede cambiar.
    '       De otra manera, el nuevo valor de la celda se perdería.
    '       columnas
    ' 00 seleccion
    ' 01 clave de la lista de precios
    ' 02 lista de precios
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 precio
    ' 08 lista predeterminada
    ' 09 nuevo en la lista de precios
    
    If grid.Col = cintcolgrdPrecio Then
        If txtEditCol.Visible Then
            If txtEditCol.Text <> "" Then
                If IsNumeric(txtEditCol.Text) Then
                    grid.Text = FormatCurrency(txtEditCol.Text, 2)
                End If
            End If
            txtEditCol.Visible = False
        End If
    ElseIf grid.Col = cintcolgrdMargenUt Then
        If txtEditCol.Visible Then
            If txtEditCol.Text <> "" Then
                txtEditCol.Text = Replace(txtEditCol.Text, "%", "")
                If IsNumeric(txtEditCol.Text) Then
                    grid.Text = Format(txtEditCol.Text, "0.0000") & "%"
                End If
            End If
            txtEditCol.Visible = False
        End If
    Else
        Exit Sub
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pSetCellValueCol"))
End Sub

Private Sub pPonerUpDown(grid As MSHFlexGrid)
    On Error GoTo NotificaError

    Dim intIndex As Integer
    'If grid.TextMatrix(grid.Row, cintColTipo) = "AR" Then
        UpDown1.ListIndex = -1
        For intIndex = 0 To UpDown1.ListCount - 1
            If UpDown1.List(intIndex) = IIf(grid.Text = "ÚLTIMA", "ÚLTIMA COMPRA", IIf(grid.Text = "COMPRA", "COMPRA MÁS ALTA", "PRECIO MÁXIMO AL PÚBLICO")) Then
                UpDown1.ListIndex = intIndex
                Exit For
            End If
        Next
        With grid
            UpDown1.Move .Left + .CellLeft, .Top + .CellTop, UpDown1.Width, UpDown1.Height
        End With
        UpDown1.Visible = True
        UpDown1.SetFocus
    'End If
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":pPonerUpDown"))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Or (UnloadMode = 1 And vgblnCancel) Then
        Cancel = 1
        '¿Desea abandonar la operación?
        If MsgBox(SIHOMsg(17), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
            vgblnCancel = True
            Me.Hide 'ocultamos el formulario
        Else
            vgblnCancel = False
            grdPrecios.SetFocus
        End If
    End If
End Sub

Private Sub grdPrecios_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'columnas
    ' 00 seleccion
    ' 01 clave de la lista de precios
    ' 02 lista de precios
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 precio
    ' 08 lista predeterminada
    ' 09 nuevo en la lista de precios
    If grdPrecios.Col = cintcolgrdPrecio Or grdPrecios.Col = cintcolgrdMargenUt Then
        If KeyCode = vbKeyF2 And grdPrecios.Row <> 0 Then pEditarColumna 13, txtEditCol, grdPrecios
    ElseIf grdPrecios.Col = cintcolgrdIncremento Then
    Else
        If KeyCode = vbKeyReturn Then
            If grdPrecios.Row - 1 < grdPrecios.Rows Then
                If grdPrecios.Row = grdPrecios.Rows - 1 Then
                    grdPrecios.Row = 1
                Else
                    grdPrecios.Row = grdPrecios.Row + 1
                End If
            End If
        End If
    End If

    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyDown"))
End Sub
Private Sub grdPrecios_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    'columnas
    ' 00 seleccion
    ' 01 clave de la lista de precios
    ' 02 lista de precios
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 precio
    ' 08 lista predeterminada
    ' 09 nuevo en la lista de precios
    If grdPrecios.MouseRow <> 0 Then
        If grdPrecios.Col = cintcolgrdPrecio Or grdPrecios.Col = cintcolgrdMargenUt Then    'Columna que puede ser editada
            pEditarColumna KeyAscii, txtEditCol, grdPrecios
        ElseIf grdPrecios.Col = cintcolgrdIncrementoAuto Or grdPrecios.Col = cintcolgrdTabulador Then
            If KeyAscii = 32 Then grdPrecios_Click
        ElseIf grdPrecios.Col = cintcolgrdIncremento Then
            If KeyAscii = 13 Then pPonerUpDown grdPrecios
        Else
            Exit Sub
        End If
    End If
    
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_KeyPress"))
End Sub
Private Sub grdPrecios_LeaveCell()
    On Error GoTo NotificaError
        If vgblnNoEditar Then Exit Sub
        grdPrecios_GotFocus
    Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_LeaveCell"))
End Sub
Private Sub grdPrecios_GotFocus()
    On Error GoTo NotificaError
        If vgblnNoEditar Then Exit Sub
        'Copia el valor del textbox al grid y lo esconde
        '4 = tipode incremento
        '5 = margen de utilidad
        '7 = precio
        If grdPrecios.Col = cintcolgrdPrecio Or grdPrecios.Col = cintcolgrdMargenUt Then
            pSetCellValueCol grdPrecios, txtEditCol
        ElseIf grdPrecios.Col = cintcolgrdIncremento Then
            UpDown1.Visible = False
        Else
            Exit Sub
        End If
   Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_GotFocus"))
End Sub

Private Sub grdPrecios_Click()
    On Error GoTo NotificaError
    'columnas
    ' 00 seleccion
    ' 01 clave de la lista de precios
    ' 02 lista de precios
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 precio
    ' 08 lista predeterminada
    ' 09 nuevo en la lista de precios
                
    If grdPrecios.MouseRow <> 0 Then ' mientras que no sea el renglon de los titulos
        If grdPrecios.Col = cintcolgrdPrecio Or grdPrecios.Col = cintcolgrdMargenUt Then   'columna de precio(8) o columna de margen de utilidad (5)
            pEditarColumna 32, txtEditCol, grdPrecios
        ElseIf grdPrecios.Col = cintcolgrdIncremento Then ' columna de tipo de incremento
            pPonerUpDown grdPrecios
        ElseIf grdPrecios.Col = cintcolgrdTabulador Then 'columna de usa tabulador
            grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdTabulador) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdTabulador)) = "*", "", "*")
        ElseIf grdPrecios.Col = cintcolgrdIncrementoAuto Then 'columna de incremento automatico
            grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdIncrementoAuto) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdIncrementoAuto)) = "*", "", "*")
        End If
               
        If grdPrecios.MouseCol = 0 Then
            grdPrecios.TextMatrix(grdPrecios.Row, 0) = IIf(Trim(grdPrecios.TextMatrix(grdPrecios.Row, 0)) = "*", "", "*")
            If Trim(grdPrecios.TextMatrix(grdPrecios.Row, 0)) = "*" Then
                llngMarcados = llngMarcados + 1
            Else
                llngMarcados = llngMarcados - 1
            End If
            pHabilitaModificar 'AQUI HABILITAMOS EL FRAME DE MODIFICAR LISTAS
        End If
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_Click"))
End Sub

Private Sub grdPrecios_Scroll()
    On Error GoTo NotificaError
    grdPrecios_GotFocus
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":grdPrecios_Scroll"))
End Sub

Private Sub txtEditCol_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo NotificaError
    'Para verificar que tecla fue presionada en el textbox
    With grdPrecios
        Select Case KeyCode
            Case 27   'ESC
                .SetFocus
                txtEditCol.Visible = False
                KeyCode = 0
                vlblnEscTxtEditCOl = True
            Case 38   'Flecha para arriba
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    vgblnNoEditar = True
                    .Row = .Row - 1
                    vgblnNoEditar = False
                End If
                vlblnEscTxtEditCOl = False
            Case 40, 13
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    vgblnNoEditar = True
                    .Row = .Row + 1
                    vgblnNoEditar = False
                End If
                vlblnEscTxtEditCOl = False
        End Select
    End With
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtEditCol_KeyDown"))
End Sub

Private Sub txtEditCol_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    'fblnFormatoCantidad esta en MODPROCEDIMIENTOS
    ' Solo permite números
    Dim bytNumDecimales As Byte
    If grdPrecios.Col = cintcolgrdMargenUt Then 'margen de utilidad
        bytNumDecimales = 4
    Else
        bytNumDecimales = 2 ' precio
    End If
    If Not fblnFormatoCantidad(txtEditCol, KeyAscii, bytNumDecimales) Then
        KeyAscii = 7
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":txtPrecio_KeyPress"))
End Sub

Private Sub txtMargenUtilidad_GotFocus()
    pTxtEditCOlPierdeFoco
    txtMargenUtilidad.Text = Replace(txtMargenUtilidad.Text, "%", "")
    pSelTextBox txtMargenUtilidad
End Sub

Private Sub txtMargenUtilidad_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 4, txtMargenUtilidad) Then KeyAscii = 0
End Sub

Private Function fValidaCantidad(vlintCaracter As Integer, vlintDecimales As Integer, CajaText As TextBox) As Boolean ' procedimiento para validar la cantidad que se introduce a un textbox
    Dim vlintPosicionCursor As Integer
    Dim vlintPosiciones As Integer
    Dim vlintPosicionPunto As Integer
    Dim vlintNumeroDecimales As Integer
    
    fValidaCantidad = True
    If Not IsNumeric(Chr(vlintCaracter)) Then 'no es numero
        If Not vlintCaracter = vbKeyBack Then 'no es retroceso
            If Not vlintCaracter = vbKeyReturn Then 'no es Enter
                If Not vlintCaracter = 46 Then ' no es el punto
                    fValidaCantidad = False ' se anula, estos son los unicos caracteres que se pueden ingresar al texbox
                Else 'es un punto debemos veriricar si se tiene un punto ya en el text
                    If fblnValidaPunto(CajaText.Text) Then ' ya hay un punto
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    Else ' se intenta ingresar un caracter numerico, revisar decimales, revisar si se tiene seleccionado el textbox
        If CajaText.SelText <> CajaText.Text Then
            vlintPosicionCursor = CajaText.SelStart
            vlintPosicionPunto = InStr(1, CajaText.Text, ".")
            If vlintPosicionPunto > 0 Then ' si hay punto
                If vlintPosicionCursor > vlintPosicionPunto Then ' si la poscion es mayor entonces debemos de revisar los decimales
                    'contamos la cantidad de decimales
                    For vlintPosiciones = vlintPosicionPunto + 1 To Len(CajaText.Text)
                    vlintNumeroDecimales = vlintNumeroDecimales + 1
                    Next vlintPosiciones
                    'si ya son tantos como vlinDecimales entonces no permite la insercion
                    If vlintNumeroDecimales >= vlintDecimales Then
                        fValidaCantidad = False
                    End If
                End If
            End If
        End If
    End If
End Function

Private Sub txtMargenUtilidad_LostFocus()
    If txtMargenUtilidad.Text <> "" Then
        txtMargenUtilidad.Text = Format(txtMargenUtilidad.Text, "0.0000") & "%"
    Else
        txtMargenUtilidad.Text = "0.0000%"
    End If
End Sub

Private Sub txtPrecio_GotFocus()
    pTxtEditCOlPierdeFoco
    txtPrecio.Text = Replace(txtPrecio.Text, "$", "")
    pSelTextBox txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    If Not fValidaCantidad(KeyAscii, 2, txtPrecio) Then KeyAscii = 0
End Sub

Private Sub txtPrecio_LostFocus()
    If txtPrecio.Text <> "" Then
        txtPrecio.Text = FormatCurrency(txtPrecio.Text, 2)
    Else
        txtPrecio.Text = "$0.00"
    End If
End Sub

Private Sub UpDown1_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    If KeyAscii = 13 Then
        UpDown1_MouseUp 0, 0, 0, 0
        If grdPrecios.Row < grdPrecios.Rows - 1 Then
            grdPrecios.Row = grdPrecios.Row + 1
        End If
    End If
    If KeyAscii = 27 Then
        grdPrecios.SetFocus
        UpDown1.Visible = False
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_KeyPress"))
End Sub

Private Sub UpDown1_LostFocus()
    On Error GoTo NotificaError
    UpDown1.Visible = False
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_LostFocus"))
End Sub

Private Sub UpDown1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo NotificaError
    'columnas
    ' 00 seleccion
    ' 01 clave de la lista de precios
    ' 02 lista de precios
    ' 03 incremento automatico
    ' 04 tipo incremento
    ' 05 margen de utilidad
    ' 06 usar tabulador
    ' 07 precio
    ' 08 lista predeterminada
    ' 09 nuevo en la lista de precios

    'If UpDown1.Text = "PRECIO MÁXIMO AL PÚBLICO" And grdPrecios.TextMatrix(grdPrecios.Row, 4) <> "PRECIO" Then
    '    If MsgBox(SIHOMsg(1220), vbYesNo + vbQuestion, "Mensaje") = vbYes Then
    '        pPoliticaPrecioMaximo (Replace(FormatCurrency(grdPrecios.TextMatrix(grdPrecios.Row, 9), 4), "$", ""))
    '    Else
    '        grdPrecios.SetFocus
    '        UpDown1.Visible = False
    '        Exit Sub
    '    End If
    'End If
       
    Select Case UpDown1.Text
        Case "ÚLTIMA COMPRA"
            grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdIncremento) = "ÚLTIMA"
        Case "COMPRA MÁS ALTA"
            grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdIncremento) = "COMPRA"
        Case "PRECIO MÁXIMO AL PÚBLICO"
            grdPrecios.TextMatrix(grdPrecios.Row, cintcolgrdIncremento) = "PRECIO"
    End Select
    grdPrecios.SetFocus
    UpDown1.Visible = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_MouseUp"))
End Sub

Private Sub UpDown1_Validate(Cancel As Boolean)
    On Error GoTo NotificaError
    UpDown1_MouseUp 0, 0, 0, 0
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (Me.Name & ":UpDown1_Validate"))
End Sub

Private Function fblnValidaDecimales(numero As String, decimales As Integer)
    Dim vlintPosicionPunto As Integer
    Dim vlintdecimalesdpunto As Integer
    fblnValidaCuatroDecimales = True
    If numero <> "" Then
        vlintPosicionPunto = InStr(1, numero, ".")
        If vlintPosicionPunto > 0 Then
            For vlintContador = vlintPosicionPunto + 1 To Len(numero)
                vlintdecimalesdpunto = vlintdecimalesdpunto + 1
            Next vlintContador
            If vlintdecimalesdpunto >= decimales Then fblnValidaCuatroDecimales = False
        End If
    End If
End Function
