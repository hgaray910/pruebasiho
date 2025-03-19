VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmParametroExclusionDescuento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusión de descuentos"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMoveable 
      BackColor       =   &H80000018&
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   6525
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Frame fraBorrar 
      Height          =   720
      Left            =   5175
      TabIndex        =   3
      Top             =   6420
      Width           =   630
      Begin VB.CommandButton cmdBorrar 
         Enabled         =   0   'False
         Height          =   495
         Left            =   60
         MaskColor       =   &H80000005&
         Picture         =   "frmParametroExclusionDescuento.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   165
         UseMaskColor    =   -1  'True
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6420
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   10845
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdParametros 
         Height          =   6150
         Left            =   60
         TabIndex        =   0
         Top             =   165
         Width           =   10710
         _ExtentX        =   18891
         _ExtentY        =   10848
         _Version        =   393216
         Cols            =   5
         FormatString    =   "|Departamento|Concepto|Tipo paciente|Procedencia"
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
End
Attribute VB_Name = "frmParametroExclusionDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmParametroExclusionDescuento
'-------------------------------------------------------------------------------------
'| Objetivo: Registrar los parámetros en PvExclusionDescuento
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rodolfo Ramos G. | Rosenda Hernandez Anaya
'| Autor                    : Rosenda Hernandez Anaya
'| Fecha de Creación        : 22/Noviembre/2002
'| Fecha Terminación        : 25/Noviembre/2002
'| Modificó                 :
'| Fecha última modificación:
'| Descripción de la modificación:
'-------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset
Dim vlstrx As String
Dim vllngSeleccionados As Long
Dim vllngUltimoValorDepto As Long
Dim vllngUltimoValorConcepto As Long
Dim vlstrUltimoValorTipo As String
Dim vllngUltimoValorProcedencia As Long


Private Sub cboMoveable_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    Dim rsPvExclusionDescuento  As New ADODB.Recordset
    Dim vllngPersonaGraba As Long, vllngSecuencia As Long
    
    If KeyAscii = 13 Then
        If cboMoveable.ListIndex <> -1 Then
            cboMoveable.Visible = False
            
            grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col) = cboMoveable.List(cboMoveable.ListIndex)
            If grdParametros.Col <> 6 Then
                grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col + 1) = cboMoveable.ItemData(cboMoveable.ListIndex)
            Else
                grdParametros.TextMatrix(grdParametros.Row, 7) = IIf(Trim(cboMoveable.List(cboMoveable.ListIndex)) = "<TODOS>", "A", IIf(Trim(cboMoveable.List(cboMoveable.ListIndex)) = "INTERNOS", "I", "E"))
            End If
            
            If grdParametros.Col = 8 Then
                If Trim(grdParametros.TextMatrix(grdParametros.Row, 3)) = "" Then
                    grdParametros.Col = 2
                    grdParametros_Click
                Else
                    If Trim(grdParametros.TextMatrix(grdParametros.Row, 5)) = "" Then
                        grdParametros.Col = 4
                        grdParametros_Click
                    Else
                        If Trim(grdParametros.TextMatrix(grdParametros.Row, 7)) = "" Then
                            grdParametros.Col = 6
                            grdParametros_Click
                        Else
                            If flngRepetidos() = 0 Then
                                '--------------------------------------------------------
                                ' Persona que graba
                                '--------------------------------------------------------
                                vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
                                If vllngPersonaGraba = 0 Then Exit Sub
                                'Grabar
                                vlstrx = "select * from PvExclusionDescuento where intConsecutivo=" & Str(Val(grdParametros.TextMatrix(grdParametros.Row, 1)))
                                
                                Set rsPvExclusionDescuento = frsRegresaRs(vlstrx, adLockOptimistic, adOpenDynamic)
                                vllngSecuencia = -1
                                If rsPvExclusionDescuento.RecordCount = 0 Then
                                    rsPvExclusionDescuento.AddNew
                                Else
                                    vllngSecuencia = rsPvExclusionDescuento!intConsecutivo
                                End If
                                rsPvExclusionDescuento!intCveDepartamento = Val(grdParametros.TextMatrix(grdParametros.Row, 3))
                                rsPvExclusionDescuento!intCveConcepto = Val(grdParametros.TextMatrix(grdParametros.Row, 5))
                                rsPvExclusionDescuento!CHRTIPOPACIENTE = grdParametros.TextMatrix(grdParametros.Row, 7)
                                rsPvExclusionDescuento!intCveEmpresaTipoPaciente = grdParametros.TextMatrix(grdParametros.Row, 9)
                                rsPvExclusionDescuento!intClaveEmpresaContable = vgintClaveEmpresaContable
                                rsPvExclusionDescuento.Update
                                If vllngSecuencia = -1 Then
                                  vllngSecuencia = flngObtieneIdentity("sec_PVExclusionDescuento", rsPvExclusionDescuento!intConsecutivo)
                                  Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "EXCLUSION DE DESCUENTOS", CStr(vllngSecuencia))
                                Else
                                  Call pGuardarLogTransaccion(Me.Name, EnmCambiar, vllngPersonaGraba, "EXCLUSION DE DESCUENTOS", CStr(vllngSecuencia))
                                End If
                                rsPvExclusionDescuento.Close
                                
                                vllngUltimoValorDepto = Val(grdParametros.TextMatrix(grdParametros.Row, 3))
                                vllngUltimoValorConcepto = Val(grdParametros.TextMatrix(grdParametros.Row, 5))
                                vlstrUltimoValorTipo = grdParametros.TextMatrix(grdParametros.Row, 7)
                                vllngUltimoValorProcedencia = grdParametros.TextMatrix(grdParametros.Row, 9)
                                'La información se actualizó satisfactoriamente.
                                MsgBox SIHOMsg(284), vbOKOnly + vbInformation, "Mensaje"
                                pCargaExclusion
                            Else
                                'Existe información con el mismo contenido
                                MsgBox SIHOMsg(19), vbOKOnly + vbExclamation, "Mensaje"
                            End If
                            grdParametros.Col = 2
                            grdParametros_Click
                        End If
                    End If
                End If
            Else
                grdParametros.Col = grdParametros.Col + 2
                grdParametros_Click
            End If
        Else
            pEnfocaCbo cboMoveable
        End If
    End If
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboMoveable_KeyPress"))
End Sub

Private Function flngRepetidos() As Long
    On Error GoTo NotificaError
    
    vlstrx = "" & _
    "select " & _
        "count(*) " & _
    "From " & _
        "PvExclusionDescuento " & _
    "Where " & _
        "intConsecutivo<>" & Str(Val(grdParametros.TextMatrix(grdParametros.Row, 1))) & " " & _
        "and intCveDepartamento = " & grdParametros.TextMatrix(grdParametros.Row, 3) & " " & _
        "and intCveConcepto=" & grdParametros.TextMatrix(grdParametros.Row, 5) & " " & _
        "and chrTipoPaciente='" & Trim(grdParametros.TextMatrix(grdParametros.Row, 7)) & "' " & _
        "and intCveEmpresaTipoPaciente=" & grdParametros.TextMatrix(grdParametros.Row, 9) & " " & _
        "and intClaveEmpresaContable =" & vgintClaveEmpresaContable

    flngRepetidos = frsRegresaRs(vlstrx).Fields(0)

Exit Function
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":flngRepetidos"))
End Function

Private Sub cmdBorrar_Click()
    On Error GoTo NotificaError
    
    Dim X As Long, vllngPersonaGraba As Long

    '--------------------------------------------------------
    ' Persona que graba
    '--------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    For X = 1 To grdParametros.Rows - 1
        If Trim(grdParametros.TextMatrix(X, 0)) = "*" Then
            vlstrx = "delete PvExclusionDescuento where intConsecutivo=" & grdParametros.TextMatrix(X, 1)
            pEjecutaSentencia vlstrx
            Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "EXCLUSION DE DESCUENTOS", grdParametros.TextMatrix(X, 1))
        End If
    Next X
    
    pCargaExclusion
    grdParametros.Col = 2
    grdParametros_Click
    
    cmdBorrar.Enabled = False
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdBorrar_Click"))
End Sub

Private Sub Form_Activate()
    On Error GoTo NotificaError
    
    vgstrNombreForm = Me.Name
    grdParametros.Col = 2
    grdParametros_Click

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Activate"))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError

    If KeyAscii = 27 Then
        Unload Me
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_KeyPress"))
End Sub

Private Sub Form_Load()
    On Error GoTo NotificaError
    
    Me.Icon = frmMenuPrincipal.Icon
    
    pCargaExclusion

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pCargaExclusion()
    
    Dim X As Long
    Dim Y As Long
    Dim vlrsPvSelExclusionDescuento As New ADODB.Recordset

On Error GoTo NotificaError
    cboMoveable.Visible = False
    
    vllngSeleccionados = 0
    

    With grdParametros
        .Rows = 2
        .Cols = 10
        .TextMatrix(1, 0) = ""
        .FixedCols = 1
        
        
        For X = 1 To .Cols - 1
            For Y = 1 To .Rows - 1
                .Col = X
                .Row = Y
                .CellBackColor = &H80000014
                .TextMatrix(Y, X) = ""
            Next Y
        Next X
    End With

    vgstrParametrosSP = vgintClaveEmpresaContable
    Set vlrsPvSelExclusionDescuento = frsEjecuta_SP(vgstrParametrosSP, "SP_PVSELEXCLUSIONDESCUENTO")
    If vlrsPvSelExclusionDescuento.RecordCount <> 0 Then
        pLlenarMshFGrdRs grdParametros, vlrsPvSelExclusionDescuento
        
        grdParametros.Redraw = False
        For X = 1 To grdParametros.Cols - 1
            For Y = 1 To grdParametros.Rows - 1
                grdParametros.Col = X
                grdParametros.Row = Y
                grdParametros.CellBackColor = &H80000018
            Next Y
        Next X
        grdParametros.Redraw = True
        grdParametros.Rows = grdParametros.Rows + 1
        
        grdParametros.TopRow = grdParametros.Rows - 1
    End If
    vlrsPvSelExclusionDescuento.Close
    
    With grdParametros
        .RowHeightMin = cboMoveable.Height
        .FormatString = "||Departamento||Concepto||Tipo||Empresa|"
        .ColWidth(0) = 200
        .ColWidth(1) = 0
        .ColWidth(2) = 2100
        .ColWidth(3) = 0
        .ColWidth(4) = 3500
        .ColWidth(5) = 0
        .ColWidth(6) = 1200
        .ColWidth(7) = 0
        .ColWidth(8) = 3200
        .ColWidth(9) = 0
        
        .ColAlignmentFixed(2) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
        .ColAlignmentFixed(6) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignLeftCenter
    End With

    grdParametros.Col = 2
    grdParametros.Row = grdParametros.Rows - 1

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaExclusion"))
End Sub

Private Sub grdParametros_Click()
    On Error GoTo NotificaError
    
    Dim X As Long
    Dim vlblnTermina As Boolean
    
    If grdParametros.Col = 1 Then Exit Sub
    
    cboMoveable.Clear
    
    vlstrx = ""
    'Departamento
    If grdParametros.Col = 2 Then
        vlstrx = "" & _
        "select " & _
            "vchDescripcion," & _
            "smiCveDepartamento " & _
        "From " & _
            "NoDepartamento " & _
        "Where " & _
            "bitEstatus = 1 and nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
        "Order By " & _
            "vchDescripcion "
    End If
    'Concepto de facturacion
    If grdParametros.Col = 4 Then
        vlstrx = "" & _
        "select " & _
            "chrDescripcion," & _
            "smiCveConcepto " & _
        "From " & _
            "PvConceptoFacturacion " & _
        "Where " & _
            "bitActivo = 1 " & _
        "Order By " & _
            "chrDescripcion "
    End If
    If grdParametros.Col = 6 Then
        cboMoveable.AddItem "INTERNOS", 0
        cboMoveable.AddItem "EXTERNOS", 1
    End If
    If grdParametros.Col = 8 Then
        vlstrx = "" & _
        "select " & _
            "vchDescripcion Descripcion," & _
            "intCveEmpresa Clave " & _
        "From " & _
            "CcEmpresa " & _
        "Union " & _
        "select " & _
            "vchDescripcion Descripcion," & _
            "tnyCveTipoPaciente*-1 Clave " & _
        "From " & _
            "AdTipoPaciente " & _
        "Order By " & _
            "Descripcion "
    End If
    
    If Trim(vlstrx) <> "" Then
        Set rs = frsRegresaRs(vlstrx)
        If rs.RecordCount <> 0 Then
            pLlenarCboRs cboMoveable, rs, 1, 0
        End If
    End If
        
    cboMoveable.AddItem "<TODOS>", 0
    cboMoveable.ItemData(cboMoveable.newIndex) = 0
    cboMoveable.ListIndex = 0
    
    If grdParametros.Col <> 6 Then
        If Trim(grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col)) <> "" Then
            X = 0
            vlblnTermina = False
            Do While X <= cboMoveable.ListCount - 1 And Not vlblnTermina
                If cboMoveable.ItemData(X) = Val(grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col + 1)) Then
                    cboMoveable.ListIndex = X
                    vlblnTermina = True
                End If
                X = X + 1
            Loop
        Else
            X = 0
            vlblnTermina = False
            Do While X <= cboMoveable.ListCount - 1 And Not vlblnTermina
                If cboMoveable.ItemData(X) = IIf(grdParametros.Col = 2, vllngUltimoValorDepto, IIf(grdParametros.Col = 4, vllngUltimoValorConcepto, vllngUltimoValorProcedencia)) Then
                    cboMoveable.ListIndex = X
                    vlblnTermina = True
                End If
                X = X + 1
            Loop
        End If
    Else
        If Trim(grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col)) <> "" Then
            If Trim(grdParametros.TextMatrix(grdParametros.Row, 7)) = "A" Then
                cboMoveable.ListIndex = 0
            Else
                If Trim(grdParametros.TextMatrix(grdParametros.Row, 7)) = "I" Then
                    cboMoveable.ListIndex = 1
                Else
                    cboMoveable.ListIndex = 2
                End If
            End If
        Else
            If vlstrUltimoValorTipo = "A" Or Trim(vlstrUltimoValorTipo) = "" Then
                cboMoveable.ListIndex = 0
            Else
                If vlstrUltimoValorTipo = "I" Then
                    cboMoveable.ListIndex = 1
                Else
                    cboMoveable.ListIndex = 2
                End If
            End If
        End If
    End If
    On Error Resume Next
    cboMoveable.Move grdParametros.Left + grdParametros.CellLeft + 20, grdParametros.Top + grdParametros.CellTop - 15, grdParametros.CellWidth
    cboMoveable.Visible = True
    cboMoveable.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_Click"))
End Sub

Private Sub grdParametros_DblClick()
    On Error GoTo NotificaError
    Dim vllngColumnaActual  As Long

    If Val(grdParametros.TextMatrix(grdParametros.Row, grdParametros.Col)) <> 0 Then
        If Trim(grdParametros.TextMatrix(grdParametros.Row, 0)) = "*" Then
            grdParametros.TextMatrix(grdParametros.Row, 0) = ""
            vllngSeleccionados = vllngSeleccionados - 1
        Else
            vllngColumnaActual = grdParametros.Col
            grdParametros.Col = 0
            grdParametros.CellFontBold = True
            grdParametros.TextMatrix(grdParametros.Row, 0) = "*"
            grdParametros.Col = vllngColumnaActual
            vllngSeleccionados = vllngSeleccionados + 1
        End If
    End If
    
    If vllngSeleccionados <> 0 Then
        cmdBorrar.Enabled = True
    Else
        cmdBorrar.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_DblClick"))
End Sub

Private Sub grdParametros_KeyPress(KeyAscii As Integer)
    On Error GoTo NotificaError
    
    If KeyAscii = 13 Then
        grdParametros_Click
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_KeyPress"))
End Sub

Private Sub grdParametros_Scroll()
    On Error GoTo NotificaError
    
    cboMoveable.Visible = False

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":grdParametros_Scroll"))
End Sub
