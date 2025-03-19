VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmHonorariosCirugia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Honorarios de cirugía"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   7725
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   6300
         TabIndex        =   1
         ToolTipText     =   "Aceptar"
         Top             =   3100
         Width           =   1335
      End
      Begin VSFlex7LCtl.VSFlexGrid grdHonorarios 
         Height          =   2655
         Left            =   75
         TabIndex        =   2
         ToolTipText     =   "Honorarios médicos que forman parte del paquete"
         Top             =   360
         Width           =   7575
         _cx             =   13361
         _cy             =   4683
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblMensaje 
         Caption         =   "Seleccione las funciones a incluir al paquete"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmHonorariosCirugia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Listado de funciones de cirugia configuradas para honorarios médicos, en esta pantalla se seleccionarán las asignadas al paciente de ingreso PREVIO
' Fecha de inicio de desarrollo:
' Fecha de término:
'Análisis y diseño        : Mayra Armendáriz
'Autor                    : Mayra Armendáriz
'------------------------------------------------------------------------------
Public lngNumPaquete As Long '
Public blnSelecciona As Boolean 'Variable para indicar que se va a seleccionar solo unos elementos para incluir al paquete
Public vstrFunciones As String
Public strNumCuenta As String 'Cuando sea un paciente ya capturado, para revisar qué paquete si se le asignó
Dim arrFunciones() As Long
Const cintColClaveFuncion = 1
Const cintColDescripcionFuncion = 2
Const cintColClaveConcepto = 3
Const cintColDescripcionConcepto = 4
Const cintColImporteHonorario = 5

Private Sub cmdAceptar_Click()
    ''Agregar al arreglo los que están seleccionados
    Dim intcontador As Integer
    Dim intTamArr As Integer
    vstrFunciones = ""
    intTamArr = 0
    With grdHonorarios
        For intcontador = 1 To .Rows - 1
            If .TextMatrix(intcontador, 0) = "*" Then
                vstrFunciones = vstrFunciones & "-" & .TextMatrix(intcontador, cintColClaveFuncion)
            End If
        Next intcontador
    End With
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 27 Then
        Unload Me
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMenuPrincipal.Icon
    pConfiguraGridHonorarios
    
    If blnSelecciona Then
        lblMensaje.Visible = True
    Else
        lblMensaje.Visible = False
    End If
    cargarHonorarios
End Sub

Private Sub cargarHonorarios()
 Dim rsHonorarios As ADODB.Recordset
 Dim rsHonorariosPaciente As ADODB.Recordset
 Dim strSentencia As String
 Dim contador As Integer
    
    'Cargar los honorarios asignados al paquete
    vlstrSentencia = ""
    'Se cargan los departamentos asignados al paquete
     strSentencia = "SELECT PVPAQUETEHONORARIOS.INTCVEPAQUETE " & _
                        ",PVPAQUETEHONORARIOS.INTCVEFUNCION, TRIM(EXFUNCIONPARTICIPANTECIRUGIA.VCHDESCRIPCION) NOMBREFUNCION " & _
                        " ,PVPAQUETEHONORARIOS.INTCVECONCEPTO, TRIM(PVOTROCONCEPTO.CHRDESCRIPCION) DESCRIPCIONCONCEPTO" & _
                        " , PVPAQUETEHONORARIOS.MNYIMPORTEHONORARIO" & _
                        " From PVPAQUETEHONORARIOS " & _
                        " INNER JOIN EXFUNCIONPARTICIPANTECIRUGIA ON EXFUNCIONPARTICIPANTECIRUGIA.INTCVEFUNCION = PVPAQUETEHONORARIOS.INTCVEFUNCION " & _
                        " INNER JOIN PVOTROCONCEPTO ON PVOTROCONCEPTO.INTCVECONCEPTO = PVPAQUETEHONORARIOS.INTCVECONCEPTO " & _
                     " WHERE PVPAQUETEHONORARIOS.INTCVEPAQUETE = " & lngNumPaquete & _
                     " ORDER BY NOMBREFUNCION"
    Set rsHonorarios = frsRegresaRs(strSentencia)

    With grdHonorarios
        .Rows = 1
        If rsHonorarios.RecordCount <> 0 Then
            Do While Not rsHonorarios.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, cintColDescripcionFuncion) = Trim(rsHonorarios!NOMBREFUNCION)
                .TextMatrix(.Rows - 1, cintColClaveFuncion) = Trim(rsHonorarios!INTCVEFUNCION)
                .TextMatrix(.Rows - 1, cintColDescripcionConcepto) = Trim(rsHonorarios!DescripcionConcepto)
                .TextMatrix(.Rows - 1, cintColClaveConcepto) = Trim(rsHonorarios!intCveConcepto)
                .TextMatrix(.Rows - 1, cintColImporteHonorario) = FormatCurrency(rsHonorarios!MNYIMPORTEHONORARIO, 2)
                rsHonorarios.MoveNext
            Loop
        End If
        
        
    End With
    rsHonorarios.Close
    
    'ver si el paciente se le asignaron solo unos honorarios:
    If strNumCuenta <> "" Then
        Set rsHonorariosPaciente = frsRegresaRs("Select INTCVEFUNCION from PVPACIENTEHONORARIOCIRUGIA where INTCVEPAQUETE= " & lngNumPaquete & " and INTNUMCUENTA = " & strNumCuenta)
        With grdHonorarios
            If rsHonorariosPaciente.RecordCount <> 0 Then
                Do While Not rsHonorariosPaciente.EOF
                    For contador = 1 To .Rows - 1
                        If rsHonorariosPaciente!INTCVEFUNCION = Val(.TextMatrix(contador, cintColClaveFuncion)) Then
                            .TextMatrix(contador, 0) = "*"
                        End If
                    Next contador
                
                    rsHonorariosPaciente.MoveNext
                Loop
            End If
        
        End With
        rsHonorariosPaciente.Close
    End If
    
End Sub
Private Sub pConfiguraGridHonorarios()
On Error GoTo NotificaError
    
    With grdHonorarios
        .FixedCols = 1
        .FixedRows = 1
        .Rows = 2
        .Cols = 6
        .FormatString = "||Funcion de cirugía||Concepto de cargo|Importe honorario"
        .ColWidth(0) = 150
        .ColWidth(cintColClaveFuncion) = 0
        .ColWidth(cintColDescripcionFuncion) = 2900
        .ColWidth(cintColClaveConcepto) = 0
        .ColWidth(cintColDescripcionConcepto) = 2900
        .ColWidth(cintColImporteHonorario) = 1450
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pConfiguraGridHonorarios"))
End Sub
Private Sub grdHonorarios_Click()
'Solo cuando estamos en ADMISION y hay que seleccionar los elementos
    If blnSelecciona Then
        If grdHonorarios.Row > 0 And Trim(grdHonorarios.TextMatrix(grdHonorarios.Row, 1)) <> "" Then
            pSeleccionaDeselecciona
        End If
    End If

End Sub
Private Sub pSeleccionaDeselecciona()

    If Trim(grdHonorarios.TextMatrix(grdHonorarios.Row, 0)) = "" Then
        grdHonorarios.Col = 0
        grdHonorarios.CellFontBold = True
        grdHonorarios.TextMatrix(grdHonorarios.Row, 0) = "*"
    Else
        grdHonorarios.TextMatrix(grdHonorarios.Row, 0) = ""
    End If

End Sub
