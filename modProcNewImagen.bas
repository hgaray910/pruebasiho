Attribute VB_Name = "modProcNewImagen"
''Eliminación de la barra de titulo manteniendo el Caption para las pantallas de inicio
''####################################################################
Option Explicit
'DEclaraciones API
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                         (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                         (ByVal hwnd As Long, ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                          ByVal hWndInsertAfter As Long, ByVal X As Long, _
                          ByVal Y As Long, ByVal cx As Long, ByVal cy As Long)
                          
                          
Dim vgstrSentencia As String
  
'El primer parámetro es el Hwnd de la ventana
Public Sub Quitar_Barra_Titulo(ByVal hwnd As Long)
    SetWindowLong hwnd, (-16), GetWindowLong(hwnd, (-16)) And Not &HC00000
    SetWindowPos hwnd, 0, 0, 0, 0, 0
End Sub
''####################################################################


Public Sub pLlenarCboRs_new(ObjCbo As MyCombo, ObjRs As Recordset, vlintNumCampoItmData As Integer, vlintNumCampoList As String, Optional vlintNumCaso As Integer, Optional vlblnMuestraError As Boolean)
'-------------------------------------------------------------------------------------------
' Llena un combobox con datos de un recordset, pidiendo
' ObjCbo Combo Box en donde se llenaran los datos
' ObjRS Recorsed de donde se llenaran los datos
' vlintNumCampoItmData Numero de campo dek RS para guardarlo en la posicion ItemData del ComboBox
' vlintNumCampoList Numero del campo del RS para guardarlos en la posicion List del ComboBox
' vlintNumCaso para llenar con Agregar o Mantenimiento segun el caso
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumCampoOrd As Long
    Dim vlintNumReg As Long
    Dim vlintseq As Long
    Dim vlstrNombreCampo As String
    Dim vlstrNomCampos As String
    Dim vlNum As Long
    Dim vlintIniciaSeq As Long
    Dim vlstrCampo As String
    Dim vlintSeqCampos As Long
    Dim vlintNum As Long
    
    vlintNumReg = ObjRs.RecordCount
    vlNum = 0
    ObjCbo.Clear
    Select Case vlintNumCaso
        Case Is = 1
                ObjCbo.AddItem "<AGREGAR>", 0
        Case Is = 2
            If vlintNumCaso > 0 Then
                Select Case vlintNumReg
                    Case 0
                        ObjCbo.AddItem "<AGREGAR>", 0
                    Case Is > 0
                        ObjCbo.AddItem "<MANTENIMIENTO>", 0
                End Select
            End If
        Case Is = 3
                ObjCbo.AddItem "<TODOS>", 0
        Case Is = 4
                ObjCbo.AddItem "<NINGUNO(A)>", 0
        Case Is = 5
                ObjCbo.AddItem "<MANTENIMIENTO>", 0
    End Select
    
    vlintNumCampoOrd = CInt(fstrFormatTxt(vlintNumCampoList, "N", "", 20, False))
    If (vlintNumReg > 0) Then
        ObjRs.MoveFirst
        vlstrNombreCampo = ObjRs.Fields(vlintNumCampoOrd).Name
        vlstrCampo = vlstrNombreCampo & " Asc"
        ObjRs.Sort = vlstrCampo
        vlintIniciaSeq = ObjCbo.ListCount
        For vlintseq = 1 To vlintNumReg
            Select Case vlintNumCampoList
                Case Is = "*" 'En la lista suma todos los campos
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & ""
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                            End If
                        Else
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & " - "
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Is = "@" 'En la lista suma todos los campos menos el campo clave
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    vlintNum = vlintNumCampoItmData
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos <> vlintNum Then
                            If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & ""
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                                End If
                            Else
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & " - "
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                                End If
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Else 'En la lista va solo en campo que se envio
                    If IsNull(ObjRs.Fields(CInt(vlintNumCampoList))) Then
                        ObjCbo.AddItem "< >", vlintIniciaSeq
                    Else
                        ObjCbo.AddItem UCase(ObjRs.Fields(CInt(vlintNumCampoList))), vlintIniciaSeq
                    End If
            End Select
            
            If vlintNumCampoItmData >= 0 Then
                If IsNumeric(ObjRs.Fields(vlintNumCampoItmData).Value) = True Then
                    ObjCbo.ItemData(vlintIniciaSeq) = CDbl(ObjRs.Fields(vlintNumCampoItmData).Value)
                Else
                    ObjCbo.ItemData(vlintIniciaSeq) = 0
                End If
            Else
                ObjCbo.ItemData(vlintIniciaSeq) = ObjRs.Bookmark
            End If
            vlintIniciaSeq = vlintIniciaSeq + 1
            ObjRs.MoveNext
        Next vlintseq
        ObjRs.MoveFirst
    Else
        If vlblnMuestraError = True Then
            Call MsgBox((SIHOMsg(13) & Chr(13) & ObjCbo.ToolTipText), vbExclamation, "Mensaje") 'Toma un mensaje del módulo de mensajes y lo despliega
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboRs_new"))
End Sub


Public Function flngLocalizaCbo_new(ObjCbo As MyCombo, vlstrCriterio As String) As Long
  '-------------------------------------------------------------------------------------------
  ' Busca un criterio dentro del combobox -1 indica que no encontro
  '-------------------------------------------------------------------------------------------
  On Error GoTo NotificaError
  
  Dim vllngNumReg As Long
  Dim vllngseq As Long
  
  flngLocalizaCbo_new = -1
  vllngNumReg = ObjCbo.ListCount
  
  If Len(vlstrCriterio) > 0 Then
    For vllngseq = 0 To vllngNumReg - 1
      If ObjCbo.ItemData(vllngseq) = vlstrCriterio Then
        flngLocalizaCbo_new = vllngseq
        Exit For
      Else
        flngLocalizaCbo_new = -1
      End If
    Next vllngseq
  End If
    
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":flngLocalizaCbo_new"))
End Function


Public Function fblnComponenteCbo_new(cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   
   On Error GoTo NotificaError
   
   'Llena combos de Ocupaciones
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveHemoComponente, vchNombre From BsHemocomponente WHERE (bitEstatus = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnComponenteCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnComponenteCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnComponenteCbo_new"))
End Function


Public Function fblnTipoSangreCbo_new(cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   On Error GoTo NotificaError
   
   'Llena combos de Ocupaciones
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveTipoSanguineo, chrGrupo ||' '|| CASE WHEN bitPositivo = 1 THEN '+' ELSE '-' END  Grupo FROM BsTipoSanguineo"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnTipoSangreCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnTipoSangreCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnTipoSangreCbo_new"))
End Function


Public Sub pLlenarCboRs2_new(ObjCbo As MyCombo, ObjRs As Recordset, vlintNumCampoItmData As Integer, vlintNumCampoList As String, Optional vlintNumCaso As Integer, Optional vlblnMuestraError As Boolean)
'-------------------------------------------------------------------------------------------
' Llena un combobox con datos de un recordset, pidiendo
' ObjCbo Combo Box en donde se llenaran los datos
' ObjRS Recorsed de donde se llenaran los datos
' vlintNumCampoItmData Numero de campo del RS para guardarlo en la posicion ItemData del ComboBox
' vlintNumCampoList Numero del campo del RS para guardarlos en la posicion List del ComboBox
' vlintNumCaso para llenar con Agregar o Mantenimiento segun el caso
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumCampoOrd As Long
    Dim vlintNumReg As Long
    Dim vlintseq As Long
    Dim vlstrNombreCampo As String
    Dim vlstrNomCampos As String
    Dim vlNum As Long
    Dim vlintIniciaSeq As Long
    Dim vlstrCampo As String
    Dim vlintSeqCampos As Long
    Dim vlintNum As Long
    
    vlintNumReg = ObjRs.RecordCount
    vlNum = 0
    ObjCbo.Clear
    Select Case vlintNumCaso
        Case Is = 1
                ObjCbo.AddItem " ", 0
        Case Is = 2
            If vlintNumCaso > 0 Then
                Select Case vlintNumReg
                    Case 0
                        ObjCbo.AddItem "<AGREGAR>", 0
                    Case Is > 0
                        ObjCbo.AddItem "<MANTENIMIENTO>", 0
                End Select
            End If
        Case Is = 3
                ObjCbo.AddItem "<TODOS>", 0
        Case Is = 4
                ObjCbo.AddItem "<NINGUNO(A)>", 0
        Case Is = 5
                ObjCbo.AddItem "<MANTENIMIENTO>", 0
    End Select
    
    vlintNumCampoOrd = CInt(fstrFormatTxt(vlintNumCampoList, "N", "", 20, False))
    If (vlintNumReg > 0) Then
        ObjRs.MoveFirst
        vlstrNombreCampo = ObjRs.Fields(vlintNumCampoOrd).Name
        vlstrCampo = vlstrNombreCampo & " Asc"
        ObjRs.Sort = vlstrCampo
        vlintIniciaSeq = ObjCbo.ListCount
        For vlintseq = 1 To vlintNumReg
            Select Case vlintNumCampoList
                Case Is = "*" 'En la lista suma todos los campos
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & ""
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                            End If
                        Else
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & " - "
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Is = "@" 'En la lista suma todos los campos menos el campo clave
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    vlintNum = vlintNumCampoItmData
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos <> vlintNum Then
                            If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & ""
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                                End If
                            Else
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & " - "
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                                End If
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Else 'En la lista va solo en campo que se envio
                    If IsNull(ObjRs.Fields(CInt(vlintNumCampoList))) Then
                        ObjCbo.AddItem "< >", vlintIniciaSeq
                    Else
                        ObjCbo.AddItem UCase(ObjRs.Fields(CInt(vlintNumCampoList))), vlintIniciaSeq
                    End If
            End Select
            
            If vlintNumCampoItmData >= 0 Then
                If IsNumeric(ObjRs.Fields(vlintNumCampoItmData).Value) = True Then
                    ObjCbo.ItemData(vlintIniciaSeq) = CDbl(ObjRs.Fields(vlintNumCampoItmData).Value)
                Else
                    ObjCbo.ItemData(vlintIniciaSeq) = 0
                End If
            Else
                ObjCbo.ItemData(vlintIniciaSeq) = ObjRs.Bookmark
            End If
            vlintIniciaSeq = vlintIniciaSeq + 1
            ObjRs.MoveNext
        Next vlintseq
        ObjRs.MoveFirst
    Else
        If vlblnMuestraError = True Then
            Call MsgBox((SIHOMsg(13) & Chr(13) & ObjCbo.ToolTipText), vbExclamation, "Mensaje") 'Toma un mensaje del módulo de mensajes y lo despliega
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboRs2_new"))
End Sub


Public Function fintLocalizaCbo_new(ObjCbo As MyCombo, vlstrCriterio As String) As Integer
'-------------------------------------------------------------------------------------------
' Busca un criterio dentro del combobox -1 indica que no encontro
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumReg As Integer
    Dim vlintseq As Integer
    
    fintLocalizaCbo_new = -1
    vlintNumReg = ObjCbo.ListCount
    If Len(vlstrCriterio) > 0 Then
        For vlintseq = 0 To vlintNumReg - 1
            If ObjCbo.ItemData(vlintseq) = vlstrCriterio Then
                fintLocalizaCbo_new = vlintseq
                Exit For
            Else
                fintLocalizaCbo_new = -1
            End If
        Next vlintseq
    End If

Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaCbo_new"))
End Function


Public Function fintLocalizaCritCbo_new(ObjCbo As MyCombo, vlstrCriterio As String) As Integer
'-------------------------------------------------------------------------------------------
' Busca un criterio dentro del combobox en el campo list
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumReg As Integer
    Dim vlintseq As Integer
    
    vlintNumReg = ObjCbo.ListCount
    If Len(vlstrCriterio) > 0 Then
        For vlintseq = 0 To vlintNumReg
            If ObjCbo.List(vlintseq) = vlstrCriterio Then
                fintLocalizaCritCbo_new = vlintseq
                Exit For
            Else
                fintLocalizaCritCbo_new = -1
            End If
        Next vlintseq
    Else
        fintLocalizaCritCbo_new = -1
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaCritCbo_new"))
End Function


Public Function fintLocalizaTxtCbo_new(ObjCbo As MyCombo, vlstrCriterio As String) As Integer
'-------------------------------------------------------------------------------------------
' Busca un criterio que sea la primera coincidencia del list dentro del combobox
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumReg As Integer
    Dim vlintseq As Integer
    Dim vlintLargo As Integer
    
    vlintNumReg = ObjCbo.ListCount
    If Len(vlstrCriterio) > 0 Then
        vlintLargo = Len(vlstrCriterio)
        For vlintseq = 0 To vlintNumReg
            If UCase(Left(ObjCbo.List(vlintseq), vlintLargo)) = UCase(vlstrCriterio) Then
                fintLocalizaTxtCbo_new = vlintseq
                Exit For
            Else
                fintLocalizaTxtCbo_new = -1
            End If
        Next vlintseq
    Else
        fintLocalizaTxtCbo_new = -1
    End If
    
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fintLocalizaTxtCbo_new"))
End Function


Public Sub pEnfocaCbo_new(ObjCbo As MyCombo)
'-------------------------------------------------------------------------------------------
' Enfoca el Combo box
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError

    If ObjCbo.Enabled And ObjCbo.Visible Then ObjCbo.SetFocus

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pEnfocaCbo_new"))
End Sub


Public Sub pValidaVarCbo_new(vlstrVariable As String, ObjCbo As MyCombo, vlstrTipo, vlstrFormaOrd As String, vlintTamano As Integer, vlblnPermiteEsp As Boolean)
'-------------------------------------------------------------------------------------------------
' Valida que una variable relacionada a un Combo Box este con datos
'-------------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlstrMensaje As String
    vlstrVariable = fstrFormatTxt(vlstrVariable, vlstrTipo, vlstrFormaOrd, vlintTamano, vlblnPermiteEsp)
    If Len(vlstrVariable) = 0 Then
        vgblnErrorIngreso = True
        vlstrMensaje = SIHOMsg(2) & Chr(13) & "Dato:" & ObjCbo.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
    Else
        vgblnErrorIngreso = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaVarCbo_new"))
End Sub


Public Sub pValidaComboBox_new(ObjCbo As MyCombo)
'-------------------------------------------------------------------------------------------------
' Valida que se haya seleccionado un item del combobox
'-------------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlstrMensaje As String

    If ObjCbo.ListIndex = -1 Then
        vgblnErrorIngreso = True
        vlstrMensaje = SIHOMsg(2) & Chr(13) & "Dato:" & ObjCbo.ToolTipText
        Call MsgBox(vlstrMensaje, vbExclamation, "Mensaje")
    Else
        vgblnErrorIngreso = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pValidaComboBox_new"))
End Sub


Public Sub pEnfocaCboL_new(ObjCbo As MyCombo, vlintstrCriterio As String)
'-------------------------------------------------------------------------------------------
'Enfoca el combo box siempre y cuando no se haya seleccionado datos del mismo en una variable
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    If Len(vlintstrCriterio) = 0 Then
        ObjCbo.SetFocus
        If ObjCbo.ListCount > 1 Then
            ObjCbo.ListIndex = 1
        Else
            ObjCbo.ListIndex = 0
        End If
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pEnfocaCboL_new"))
End Sub


Public Sub pIniciaCbo_new(ObjCbo As MyCombo)
'-------------------------------------------------------------------------------------------
'Inicializa un combo box solo con el elemento <AGREGAR> en la posicion 0
'"A" solo <AGREGAR>, "E" solo <ELIMINAR>, "T" ambos <AGREGAR/ELIMINAR>
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    ObjCbo.Clear
    ObjCbo.AddItem "<AGREGAR>", 0

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaCbo_new"))
End Sub


Public Sub pLlenarCboMshFGrd_new(ObjCbo As MyCombo, ObjGrid As MSHFlexGrid, vlintNumCampoCve As Integer, vlstrNCampo As String, Optional vlintNumCaso As Integer)
'-------------------------------------------------------------------------------------------
' Llena un combobox con datos de un recordset
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintDesdeCol As Long
    Dim vlintNumCampo As Long
    Dim vlintNumReg As Long
    Dim vlintseq As Long
    Dim vlstrNombreCampo As String
    Dim vlstrNomCampos As String
    Dim vlstrcaracter As String
    Dim vlNum As Long
    Dim vlintSeqCampos As Long
    Dim vlintNCampo As Long
    
    vlintNumReg = ObjGrid.Rows - 1
    vlNum = 0
    ObjCbo.Clear
    If vlintNumCaso > 0 Then
        Select Case vlintNumReg
            Case Is <= 0
                ObjCbo.AddItem "<AGREGAR>", 0
            Case Is > 0
                ObjCbo.AddItem "<MANTENIMIENTO>", 0
        End Select
    Else
        ObjCbo.AddItem "<AGREGAR>", 0
    End If

    If (vlintNumReg > 0) Then
        vlstrcaracter = Left(vlstrNCampo, 1)
        vlintDesdeCol = CInt(Right(vlstrNCampo, Len(vlstrNCampo) - 1))
        For vlintseq = 1 To vlintNumReg
            Select Case vlstrcaracter
                Case Is = "*" 'En la lista suma todos los campos
                    vlstrNomCampos = ""
                    For vlintSeqCampos = 0 To ObjGrid.Cols - 1
                        If vlintSeqCampos = ObjGrid.Cols - 1 Then
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos)
                        Else
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos) & " - "
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintseq
                Case Is = "@"  'En la lista suma todos los campos menos el campo clave
                    vlstrNomCampos = ""
                    For vlintSeqCampos = vlintDesdeCol To ObjGrid.Cols - 1
                        If vlintSeqCampos = ObjGrid.Cols - 1 Then
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos)
                        Else
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos) & " - "
                        End If
                    Next vlintSeqCampos
                    ObjCbo.AddItem UCase(vlstrNomCampos), vlintseq
                Case Else 'En la lista va solo en campo que se envio
                    ObjCbo.AddItem UCase(ObjGrid.TextMatrix(vlintseq, CInt(vlintNCampo))), vlintseq
            End Select
        Next vlintseq
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboMshFGrd_new"))
End Sub


Public Sub pEditCboMshFGrd_new(cboEdit As MyCombo, vlstrDataMember As String, _
vlintItemData As Integer, vlintList As Integer, vlintCaso As Integer, _
ObjGrid As MSHFlexGrid, ObjForm As Form)
'Procedimiento para Editar un control ComboBox relacionado a un MshFGrid
    
    ObjGrid.RowHeight(ObjGrid.row) = 350
    ObjForm.KeyPreview = False
    With cboEdit
        .Enabled = True
        If Len(vlstrDataMember) > 0 Then
            If EntornoSIHO.Recordsets(vlstrDataMember).State = 0 Then
                EntornoSIHO.Recordsets(vlstrDataMember).Open
                Call pLlenarCboRs_new(cboEdit, EntornoSIHO.Recordsets(vlstrDataMember), vlintItemData, CStr(vlintList), vlintCaso)
                EntornoSIHO.Recordsets(vlstrDataMember).Close
            Else
                Call pLlenarCboRs_new(cboEdit, EntornoSIHO.Recordsets(vlstrDataMember), vlintItemData, CStr(vlintList), vlintCaso)
            End If
        End If
        If cboEdit.ListCount > 0 Then
            
            Select Case vlintCaso
                Case -1 To 0
                    cboEdit.ListIndex = fintLocalizaCritCbo_new(cboEdit, ObjGrid.Text)
                Case Else
                    cboEdit.ListIndex = fintLocalizaCritCbo_new(cboEdit, ObjGrid.Text)
            End Select
            With ObjGrid
                If ObjGrid.CellWidth < 0 Then
                    Exit Sub
                Else
                    cboEdit.Move ObjGrid.Left + ObjGrid.CellLeft, ObjGrid.Top + ObjGrid.CellTop, ObjGrid.CellWidth - 8
                End If
            End With
            .Visible = True
            .SetFocus
        Else
            cboEdit.Clear
            cboEdit.Visible = False
            ObjGrid.SetFocus
        End If
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pEditCboMshFGrd_new"))
End Sub


Public Sub pProcessCboMshFGrd_new(cboEdit As MyCombo, strEvento As String, KeyCode As Integer, KeyAscii As Integer, ObjGrid As MSHFlexGrid, ObjForm As Form)
'Procedimiento para desarrollar procesos relacionados con un Mktext
    Dim vlstrContenido As String
    Dim vlintseq As Integer
    
    vlstrContenido = cboEdit.Text
    
    Select Case strEvento
        Case "Click" 'Cuando se desarrolla un evento Change
            ObjGrid.Text = cboEdit.Text
        Case "Keydown"
            Select Case KeyCode
                Case vbKeyEscape
                    cboEdit.Visible = False
                    If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                        ObjGrid.Text = cboEdit.Text
                        ObjGrid.SetFocus
                    End If
                Case vbKeyUp
                    'DoEvents
                    If cboEdit.ListIndex = 0 Then
                        If (ObjGrid.row - 1) > 0 Then
                            cboEdit.Visible = False
                            ObjGrid.row = ObjGrid.row - 1
                            If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                                ObjGrid.Text = cboEdit.Text
                                ObjGrid.SetFocus
                            End If
                        Else
                            cboEdit.Visible = False
                            If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                                ObjGrid.Text = cboEdit.Text
                                ObjGrid.SetFocus
                            End If
                        End If
                    End If
                Case vbKeyReturn
                If cboEdit.Visible = True Then
                    If (ObjGrid.row + 1) > (ObjGrid.Rows - 1) Then
                        cboEdit.Visible = False
                        If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                            ObjGrid.Text = cboEdit.Text
                            ObjGrid.SetFocus
                        End If
                    Else
                        cboEdit.Visible = False
                        If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                            ObjGrid.Text = cboEdit.Text
                            ObjGrid.SetFocus
                        End If
                        ObjGrid.row = ObjGrid.row + 1
                    End If
                Else
                    If cboEdit.ListIndex = cboEdit.ListCount Then
                        ObjGrid.SetFocus
                        If (ObjGrid.row + 1) > (ObjGrid.Rows - 1) Then
                            cboEdit.Visible = False
                            If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                                ObjGrid.Text = cboEdit.Text
                                ObjGrid.SetFocus
                            End If
                        Else
                            cboEdit.Visible = False
                            If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                                ObjGrid.Text = cboEdit.Text
                                ObjGrid.SetFocus
                            End If
                            ObjGrid.row = ObjGrid.row + 1
                        End If
                    End If
                End If
            End Select
        Case "KeyPress"
            Select Case KeyAscii
                Case vbKeyDown, vbKeyF2
                Case 48 To 57, Asc("a") To Asc("z"), Asc("A") To Asc("Z"), 46
                        If cboEdit.ListCount > 0 Then
                            cboEdit.ListIndex = fintLocalizaTxtCbo_new(cboEdit, Chr(KeyAscii))
                            If Len(ObjGrid.Text) > 0 Then
                                cboEdit.ListIndex = fintLocalizaTxtCbo_new(cboEdit, ObjGrid.Text)
                            Else
                                If cboEdit.ListCount = 0 Then
                                    cboEdit.ListIndex = -1
                                Else
                                    If cboEdit.ListCount > 1 Then
                                        cboEdit.ListIndex = 1
                                    Else
                                        cboEdit.ListIndex = -1
                                    End If
                                End If
                            End If
                        End If
                Case Else
            End Select
        Case "LostFocus"
            cboEdit.Visible = False
            If ObjGrid.Enabled = True And ObjGrid.Rows > 1 Then
                ObjGrid.SetFocus
                cboEdit.Clear
            End If
            ObjGrid.Redraw = False
            For vlintseq = 1 To ObjGrid.Rows - 1
                ObjGrid.RowHeight(vlintseq) = 240
            Next vlintseq
            ObjGrid.Redraw = True
            ObjForm.KeyPreview = True
    End Select
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pProcessCboMshFGrd_new"))
End Sub


Public Function fblnLimpiaForma_new(ByVal frmObj As Form) As Boolean
   On Error GoTo NotificaError
   ' Inicializa los campos de una forma
   Dim vlobjControl As Control
   For Each vlobjControl In frmObj.Controls
      If TypeOf vlobjControl Is TextBox Then
         vlobjControl.Text = ""
      ElseIf TypeOf vlobjControl Is ComboBox Then
        If vlobjControl.Style <> 0 Then
         If vlobjControl.ListCount > 0 Then
            vlobjControl.ListIndex = 0
         Else
            'vlobjControl.ListIndex = -1
         End If
        Else
          vlobjControl.Clear
        End If
      ElseIf TypeOf vlobjControl Is MyCombo Then
        
      ElseIf TypeOf vlobjControl Is CheckBox Then
         vlobjControl.Value = vbUnchecked
      ElseIf TypeOf vlobjControl Is SSTab Then
         vlobjControl.Tab = 0
      End If
   Next
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLimpiaForma_new"))
End Function


Public Sub pFocusNextControl_new(ByRef parForm As Form, parTabIndex As Integer)
  On Error GoTo NotificaError
  
  Dim intTIAux As Integer
  Dim i As Integer
  Dim blnBand As Boolean
  
  blnBand = False
  intTIAux = parTabIndex + 1
  Dim vlobjControl As Control
  
  For Each vlobjControl In parForm.Controls
    If TypeName(vlobjControl) <> "Timer" And TypeName(vlobjControl) <> "CrystalReport" _
      And TypeName(vlobjControl) <> "SysInfo" And TypeName(vlobjControl) <> "Shape" And TypeName(vlobjControl) <> "Line" And TypeName(vlobjControl) <> "Image" Then
        If (vlobjControl.TabIndex = intTIAux) And _
            (TypeOf vlobjControl Is TextBox Or _
             TypeOf vlobjControl Is MyCombo Or _
             TypeOf vlobjControl Is CommandButton Or _
             TypeOf vlobjControl Is MaskEdBox Or _
             TypeOf vlobjControl Is OptionButton Or _
             TypeOf vlobjControl Is CheckBox Or _
             TypeOf vlobjControl Is DTPicker _
             ) Then
             
              If vlobjControl.Enabled And vlobjControl.Visible Then
                vlobjControl.SetFocus
              End If
            
            Exit For
        End If
    End If
  Next
Exit Sub
NotificaError:
  If Abs(err.Number) = 5 Then
    err.Clear
  Else
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pFocusNextControl_new"))
  End If
End Sub


Public Function flngLocalizaCboTxt_new(ObjCbo As MyCombo, vlstrCriterio As String) As Long
  '-------------------------------------------------------------------------------------------
  ' Busca un criterio dentro del combobox -1 indica que no encontro
  '-------------------------------------------------------------------------------------------
  On Error GoTo NotificaError
  
  Dim vllngNumReg As Long
  Dim vllngseq As Long
  
  flngLocalizaCboTxt_new = -1
  vllngNumReg = ObjCbo.ListCount
  
  If Len(vlstrCriterio) > 0 Then
    For vllngseq = 0 To vllngNumReg - 1
        ObjCbo.ListIndex = vllngseq
        If Trim(ObjCbo.Text) = Trim(vlstrCriterio) Then
            flngLocalizaCboTxt_new = vllngseq
            Exit For
        Else
          flngLocalizaCboTxt_new = -1
        End If
    Next vllngseq
  End If
    
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":flngLocalizaCboTxt_new"))
End Function


Public Function fbCboLlenaDepartamento_new(ByVal Cbo As MyCombo) As Boolean
   On Error GoTo NotificaError
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   
   vlstrsql = "SELECT smiCveDepartamento, vchDescripcion descrip " & _
      "From NoDepartamento " & _
      "Where(bitEstatus = 1) " & _
      "ORDER BY vchDescripcion"
      
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new Cbo, rs, 0, 1
      fbCboLlenaDepartamento_new = True
   Else
      fbCboLlenaDepartamento_new = False
   End If
   rs.Close
Exit Function
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ": fbCboLlenaDepartamento_new"))
End Function


Public Function fblnLlenaCiudadesCbo_new(cboObj As MyCombo) As Boolean
   On Error GoTo NotificaError
   
   'Llena combos de ciudades
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveCiudad, vchDescripcion From Ciudad Where (bitActiva = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnLlenaCiudadesCbo_new = True
   Else
      fblnLlenaCiudadesCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLlenaCiudadesCbo_new"))
End Function


Public Function fblnLlenaAntiCuaguloCbo_new(cboObj As MyCombo) As Boolean
On Error GoTo NotificaError

   'Llena combos de Ocupaciones
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intClave Clave, vchDescripcion Descripcion FROM BsAntiCoagulante where (bitStatus  = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnLlenaAntiCuaguloCbo_new = True
   Else
      fblnLlenaAntiCuaguloCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLlenaAntiCuaguloCbo_new"))
End Function


Public Function fblnLlenaEscolaridadCbo_new(cboObj As MyCombo) As Boolean
   On Error GoTo NotificaError

   'Llena combos de Escolaridades
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intClave Clave, vchDescripcion Descripcion From Escolaridad where (bitStatus  = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnLlenaEscolaridadCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnLlenaEscolaridadCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLlenaEscolaridadCbo_new"))
End Function

Public Function fblnLlenaEdoCivilCbo_new(cboObj As MyCombo) As Boolean
   On Error GoTo NotificaError
   
   'Llena combos de Estados civiles
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveEstadoCivil Clave, vchDescripcion Descripcion FROM siEstadoCivil "
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnLlenaEdoCivilCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnLlenaEdoCivilCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLlenaEdoCivilCbo_new"))
End Function


Public Function fblnLlenaOcupacionCbo_new(cboObj As MyCombo) As Boolean
On Error GoTo NotificaError

   'Llena combos de Ocupaciones
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intClave Clave, vchDescripcion Descripcion FROM Ocupacion where (bitStatus  = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnLlenaOcupacionCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnLlenaOcupacionCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnLlenaOcupacionCbo_new"))
End Function


Public Function fblnReligionCbo_new(ByVal cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
On Error GoTo NotificaError

   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "select intClave,vchDescripcion from AdReligion where bitStatus=1"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnReligionCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnReligionCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnReligionCbo_new"))
End Function

Public Function fblnParentescoCbo_new(ByVal cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   On Error GoTo NotificaError
   
   'Llena combos de tipo bolsa
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveParentesco, vchDescripcion From SiParentesco "
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnParentescoCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnParentescoCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnParentescoCbo_new"))
End Function

Public Sub pLlenaCombo_new(Sentence As String, Combo As MyCombo, Optional Indice As Long)
'Procedimioento que llena combos de acuerdo a una instrucción
    On Error GoTo NotificaError
    Dim vlintx As Integer
    Dim rsTemp As New ADODB.Recordset
    
    Combo.Clear     'Se limpia el combo
    
    Set rsTemp = frsRegresaRs(Sentence, adLockReadOnly, adOpenForwardOnly)
    With rsTemp
        If .RecordCount > 0 Then
            Combo.Visible = False
            Do While Not .EOF
                Combo.AddItem !Nombre, Combo.ListCount
                Combo.ItemData(Combo.NewIndex) = !Cve
                .MoveNext
            Loop
            Combo.Visible = True
            vlintx = fintLocalizaCbo_new(Combo, CStr(Indice))
            Combo.ListIndex = IIf(vlintx = -1, 0, vlintx)
        End If
        .Close
    End With

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaCombo_new"))
End Sub

Public Sub pMedicoCargo_new(ObjCbo As MyCombo, vllngNumPaciente As Long, vllngNumCuenta As Long, vlStrTipoPaciente As String, Optional Encabezado As Integer, Optional vlblnNoMostrarMensaje As Boolean)
    
    'Procedimiento que regresa en el combo los médicos asignados al paciente
    'si no los hay, asigna automáticamente el registrado en el movimiento de admisión
    'en el caso de internos o el registro de atención del paciente externo

    On Error GoTo NotificaError
    
    Dim vlstrSentencia As String
    Dim vllngCveMedico As Long
    
    Dim rsDatos As New ADODB.Recordset
    Dim rsExMedicoACargo As New Recordset

    ObjCbo.Clear
    If Encabezado = 0 Then
        ObjCbo.AddItem "<AGREGAR>", 0
    ElseIf Encabezado = 1 Then
        ObjCbo.AddItem "<SELECCIONAR>", 0
    End If
    
    '-*-*-*-*-*-*-*-*-*-*-*-
    'Verificación de la existencia de médicos asignados, asignación en su defecto
    '-*-*-*-*-*-*-*-*-*-*-*-
    
    vlstrSentencia = "select count(*) from ExMedicoaCargo where intNumPaciente = " & Str(vllngNumPaciente) & " and numNumCuenta = " & Str(vllngNumCuenta) & " and chrTipoPaciente = '" & vlStrTipoPaciente & "'"
    Set rsDatos = frsRegresaRs(vlstrSentencia)
    If rsDatos.Fields(0) = 0 Then
        If vlStrTipoPaciente = "I" Then
            vlstrSentencia = "select intCveMedicoCargo ClaveMedico from AdAdmision where numNumCuenta = " & Str(vllngNumCuenta)
        Else
            vlstrSentencia = "select intMedico ClaveMedico from RegistroExterno where intNumCuenta = " & Str(vllngNumCuenta)
        End If
        
        Set rsDatos = frsRegresaRs(vlstrSentencia)
        
        If rsDatos.RecordCount <> 0 Then
            
            vllngCveMedico = IIf(IsNull(rsDatos!ClaveMedico), 0, rsDatos!ClaveMedico)
        
            If vllngCveMedico <> 0 Then
            
                vlstrSentencia = "select * from ExMedicoACargo where numNumCuenta = -1"
                
                Set rsExMedicoACargo = frsRegresaRs(vlstrSentencia, adLockOptimistic, adOpenDynamic)
                With rsExMedicoACargo
                    .AddNew
                    !numNumCuenta = vllngNumCuenta
                    !intnumpaciente = vllngNumPaciente
                    !CHRTIPOPACIENTE = vlStrTipoPaciente
                    !intCveMedico = vllngCveMedico
                    !dtmFechaHoraCargo = fdtmServerFecha
                    !dtmFechaHoraTermino = Null
                    !chrEstatusMedico = "A"
                    .Update
                    .Close
                End With
            End If
        End If
    End If
    
    vlstrSentencia = "Select ExMedicoACargo.intCveMedico Clave, RTrim(HoMedico.vchApellidoPaterno)||' '||RTrim(HoMedico.vchApellidoMaterno)||' '||RTrim(vchNombre) Medico " & _
                     "From ExMedicoACargo " & _
                        "Inner Join HoMedico On " & _
                        "ExMedicoACargo.intCveMedico = HoMedico.intCveMedico " & _
                     "Where " & _
                        " ExMedicoACargo.intNumPaciente = " & vllngNumPaciente & _
                        " And ExMedicoACargo.numNumCuenta = " & vllngNumCuenta & _
                        " and ExMedicoACargo.chrTipoPaciente = '" & vlStrTipoPaciente & "'" & _
                        " And ExMedicoACargo.chrEstatusMedico = 'A' " & _
                     "Order By RTrim(HoMedico.vchApellidoPaterno)||' '||RTrim(HoMedico.vchApellidoMaterno)||' '||RTrim(vchNombre)"

    Set rsDatos = frsRegresaRs(vlstrSentencia, adLockReadOnly, adOpenForwardOnly)
    With rsDatos
        If .RecordCount > 0 Then
            Do While Not .EOF
                ObjCbo.AddItem !Medico, ObjCbo.ListCount
                ObjCbo.ItemData(ObjCbo.NewIndex) = !Clave
                .MoveNext
            Loop
            ObjCbo.ListIndex = IIf(.RecordCount > 1, 1, 0)
        ElseIf Not vlblnNoMostrarMensaje Then
            MsgBox SIHOMsg(467), vbExclamation, "Mensaje"
        End If
        .Close
    End With
    
Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pMedicoCargo_new"))
End Sub

Public Sub pLlenarCboSentencia_new( _
vlcboNombre As MyCombo, _
vlstrSentencia As String, _
vlintColList As Integer, _
vlintColItemData As Integer, _
Optional vlstrRenglonEspecial As String, _
Optional vllngRenglonEspecial As Long)
'vlcboNombre = Nombre del combo
'vlstrSentencia = Sentencia con la cual se cargará el recordset
'vlintColList = Columna que se pondrá en el List
'vlintColItemData = Columna que se pondrá en el ItemData
'Optional vlstrRenglonEspecial = Si se incluye un renglón especial. Ejemplo: <TODOS>
'Optional vllngRenglonEspecial = ItemData para vlstrRenglonEspecial

    On Error GoTo NotificaError

    Dim vllngContador As Long
    Dim rs As New ADODB.Recordset
    
    vlcboNombre.Clear
    Set rs = frsRegresaRs(vlstrSentencia)
    If rs.RecordCount <> 0 Then
        Do While Not rs.EOF
            vlcboNombre.AddItem rs.Fields(vlintColList)
            vlcboNombre.ItemData(vlcboNombre.NewIndex) = rs.Fields(vlintColItemData)
            rs.MoveNext
        Loop
    End If
    If Trim(vlstrRenglonEspecial) <> "" Then
        vlcboNombre.AddItem vlstrRenglonEspecial, 0
        vlcboNombre.ItemData(vlcboNombre.NewIndex) = vllngRenglonEspecial
    End If
    If vlcboNombre.ListCount <> 0 Then
        vlcboNombre.ListIndex = 0
    End If
    rs.Close

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarCboSentencia_new"))
End Sub

Public Function fblnTipoDonacionCbo_new(ByVal cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   On Error GoTo NotificaError
   
   '-----------------------------------
   'Llena combos de Tipo de Donación
   '-----------------------------------
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intClave , vchDescripcion From BsTipoDonacion WHERE bitStatus = 1"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnTipoDonacionCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnTipoDonacionCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnTipoDonacionCbo_new"))
End Function

Public Function fblnTipoBolsaCbo_new(ByVal cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   On Error GoTo NotificaError
   
   '-----------------------------------
   'Llena combos de tipo bolsa
   '-----------------------------------
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveTipoBolsa , vchNombre From BsTipoBolsa WHERE bitStatus = 1"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnTipoBolsaCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnTipoBolsaCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnTipoBolsaCbo_new"))
End Function

Public Function fblnEmpresaSangreCbo_new(cboObj As MyCombo) As Boolean
   On Error GoTo NotificaError
   
   '-----------------------------------
   'Llena combos de Ocupaciones
   '-----------------------------------
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveInstitucion, vchDescripcion FROM Institucion"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1
      fblnEmpresaSangreCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnEmpresaSangreCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnEmpresaSangreCbo_new"))
End Function

Public Function fblnReaccionesCbo_new(ByVal cboObj As MyCombo, Optional vlintNumOpcion As Integer) As Boolean
   On Error GoTo NotificaError
   
   '-----------------------------------
   'Llena combos de Ocupaciones
   '-----------------------------------
   Dim vlstrsql As String
   Dim rs As New ADODB.Recordset
   vlstrsql = "SELECT intCveReaccion, vchDescripcion From BsReaccion WHERE (bitStatus = 1)"
   Set rs = frsRegresaRs(vlstrsql, adLockReadOnly, adOpenForwardOnly)
   If rs.RecordCount > 0 Then
      pLlenarCboRs_new cboObj, rs, 0, 1, vlintNumOpcion
      fblnReaccionesCbo_new = True
      cboObj.ListIndex = 0
   Else
      fblnReaccionesCbo_new = False
   End If
   rs.Close
   
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":fblnReaccionesCbo_new"))
End Function

Public Sub pIniciaList_new(ObjCbo As ListBox)
'-------------------------------------------------------------------------------------------
'Inicializa un ListBox solo con el elemento <AGREGAR> en la posicion 0
'"A" solo <AGREGAR>, "E" solo <ELIMINAR>, "T" ambos <AGREGAR/ELIMINAR>
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    ObjCbo.Clear
    ObjCbo.AddItem "<AGREGAR>", 0

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pIniciaList_new"))
End Sub

Public Sub pLlenarLstMshFGrd_new(ObjLst As ListBox, ObjGrid As MSHFlexGrid, vlintNumCampoCve As Integer, vlstrNCampo As String, Optional vlintNumCaso As Integer)
'-------------------------------------------------------------------------------------------
' Llena un ListBox con datos de un recordset
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintDesdeCol As Long
    Dim vlintNumCampo As Long
    Dim vlintNumReg As Long
    Dim vlintseq As Long
    Dim vlstrNombreCampo As String
    Dim vlstrNomCampos As String
    Dim vlstrcaracter As String
    Dim vlNum As Long
    Dim vlintSeqCampos As Long
    Dim vlintNCampo As Long
    
    vlintNumReg = ObjGrid.Rows - 1
    vlNum = 0
    ObjLst.Clear
    If vlintNumCaso > 0 Then
        Select Case vlintNumReg
            Case Is <= 0
                ObjLst.AddItem "<AGREGAR>", 0
            Case Is > 0
                ObjLst.AddItem "<MANTENIMIENTO>", 0
        End Select
    Else
        ObjLst.AddItem "<AGREGAR>", 0
    End If

    If (vlintNumReg > 0) Then
        vlstrcaracter = Left(vlstrNCampo, 1)
        vlintDesdeCol = CInt(Right(vlstrNCampo, Len(vlstrNCampo) - 1))
        For vlintseq = 1 To vlintNumReg
            Select Case vlstrcaracter
                Case Is = "*" 'En la lista suma todos los campos
                    vlstrNomCampos = ""
                    For vlintSeqCampos = 0 To ObjGrid.Cols - 1
                        If vlintSeqCampos = ObjGrid.Cols - 1 Then
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos)
                        Else
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos) & " - "
                        End If
                    Next vlintSeqCampos
                    ObjLst.AddItem UCase(vlstrNomCampos), vlintseq
                Case Is = "@"  'En la lista suma todos los campos menos el campo clave
                    vlstrNomCampos = ""
                    For vlintSeqCampos = vlintDesdeCol To ObjGrid.Cols - 1
                        If vlintSeqCampos = ObjGrid.Cols - 1 Then
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos)
                        Else
                            vlstrNomCampos = vlstrNomCampos & ObjGrid.TextMatrix(vlintseq, vlintSeqCampos) & " - "
                        End If
                    Next vlintSeqCampos
                    ObjLst.AddItem UCase(vlstrNomCampos), vlintseq
                Case Else 'En la lista va solo en campo que se envio
                    ObjLst.AddItem UCase(ObjGrid.TextMatrix(vlintseq, CInt(vlintNCampo))), vlintseq
            End Select
        Next vlintseq
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarLstMshFGrd_new"))
End Sub

Public Function flngLocalizaLst_new(ObjLst As ListBox, vlstrCriterio As String) As Long
  '-------------------------------------------------------------------------------------------
  ' Busca un criterio dentro del ListBox -1 indica que no encontro
  '-------------------------------------------------------------------------------------------
  On Error GoTo NotificaError
  
  Dim vllngNumReg As Long
  Dim vllngseq As Long
  
  flngLocalizaLst_new = -1
  vllngNumReg = ObjLst.ListCount
  
  If Len(vlstrCriterio) > 0 Then
    For vllngseq = 0 To vllngNumReg - 1
      If ObjLst.ItemData(vllngseq) = vlstrCriterio Then
        flngLocalizaLst_new = vllngseq
        Exit For
      Else
        flngLocalizaLst_new = -1
      End If
    Next vllngseq
  End If
    
Exit Function
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":flngLocalizaLst_new"))
End Function

Public Sub pLlenarLstRs_new(ObjLst As ListBox, ObjRs As Recordset, vlintNumCampoItmData As Integer, vlintNumCampoList As String, Optional vlintNumCaso As Integer, Optional vlblnMuestraError As Boolean)
'-------------------------------------------------------------------------------------------
' Llena un listbox con datos de un recordset, pidiendo
' ObjCbo listbox en donde se llenaran los datos
' ObjRS Recorsed de donde se llenaran los datos
' vlintNumCampoItmData Numero de campo dek RS para guardarlo en la posicion ItemData del listbox
' vlintNumCampoList Numero del campo del RS para guardarlos en la posicion List del listbox
' vlintNumCaso para llenar con Agregar o Mantenimiento segun el caso
'-------------------------------------------------------------------------------------------
    On Error GoTo NotificaError
    
    Dim vlintNumCampoOrd As Long
    Dim vlintNumReg As Long
    Dim vlintseq As Long
    Dim vlstrNombreCampo As String
    Dim vlstrNomCampos As String
    Dim vlNum As Long
    Dim vlintIniciaSeq As Long
    Dim vlstrCampo As String
    Dim vlintSeqCampos As Long
    Dim vlintNum As Long
    
    vlintNumReg = ObjRs.RecordCount
    vlNum = 0
    ObjLst.Clear
    Select Case vlintNumCaso
        Case Is = 1
                ObjLst.AddItem "<AGREGAR>", 0
        Case Is = 2
            If vlintNumCaso > 0 Then
                Select Case vlintNumReg
                    Case 0
                        ObjLst.AddItem "<AGREGAR>", 0
                    Case Is > 0
                        ObjLst.AddItem "<MANTENIMIENTO>", 0
                End Select
            End If
        Case Is = 3
                ObjLst.AddItem "<TODOS>", 0
        Case Is = 4
                ObjLst.AddItem "<NINGUNO(A)>", 0
        Case Is = 5
                ObjLst.AddItem "<MANTENIMIENTO>", 0
    End Select
    
    vlintNumCampoOrd = CInt(fstrFormatTxt(vlintNumCampoList, "N", "", 20, False))
    If (vlintNumReg > 0) Then
        ObjRs.MoveFirst
        vlstrNombreCampo = ObjRs.Fields(vlintNumCampoOrd).Name
        vlstrCampo = vlstrNombreCampo & " Asc"
        ObjRs.Sort = vlstrCampo
        vlintIniciaSeq = ObjLst.ListCount
        For vlintseq = 1 To vlintNumReg
            Select Case vlintNumCampoList
                Case Is = "*" 'En la lista suma todos los campos
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & ""
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                            End If
                        Else
                            If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                vlstrNomCampos = vlstrNomCampos & " - "
                            Else
                                vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjLst.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Is = "@" 'En la lista suma todos los campos menos el campo clave
                    vlstrNombreCampo = ObjRs.Fields(0).Name
                    vlstrCampo = vlstrNombreCampo & " Asc"
                    vlstrNomCampos = ""
                    vlintNum = vlintNumCampoItmData
                    For vlintSeqCampos = 0 To ObjRs.Fields.Count - 1
                        If vlintSeqCampos <> vlintNum Then
                            If vlintSeqCampos = ObjRs.Fields.Count - 1 Then
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & ""
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos)
                                End If
                            Else
                                If IsNull(ObjRs.Fields(vlintSeqCampos)) Then
                                    vlstrNomCampos = vlstrNomCampos & " - "
                                Else
                                    vlstrNomCampos = vlstrNomCampos & ObjRs.Fields(vlintSeqCampos) & " - "
                                End If
                            End If
                        End If
                    Next vlintSeqCampos
                    ObjLst.AddItem UCase(vlstrNomCampos), vlintIniciaSeq
                Case Else 'En la lista va solo en campo que se envio
                    If IsNull(ObjRs.Fields(CInt(vlintNumCampoList))) Then
                        ObjLst.AddItem "< >", vlintIniciaSeq
                    Else
                        ObjLst.AddItem UCase(ObjRs.Fields(CInt(vlintNumCampoList))), vlintIniciaSeq
                    End If
            End Select
            
            If vlintNumCampoItmData >= 0 Then
                If IsNumeric(ObjRs.Fields(vlintNumCampoItmData).Value) = True Then
                    ObjLst.ItemData(vlintIniciaSeq) = CDbl(ObjRs.Fields(vlintNumCampoItmData).Value)
                Else
                    ObjLst.ItemData(vlintIniciaSeq) = 0
                End If
            Else
                ObjLst.ItemData(vlintIniciaSeq) = ObjRs.Bookmark
            End If
            vlintIniciaSeq = vlintIniciaSeq + 1
            ObjRs.MoveNext
        Next vlintseq
        ObjRs.MoveFirst
    Else
        If vlblnMuestraError = True Then
            Call MsgBox((SIHOMsg(13) & Chr(13) & ObjLst.ToolTipText), vbExclamation, "Mensaje") 'Toma un mensaje del módulo de mensajes y lo despliega
        End If
    End If

Exit Sub
NotificaError:
    Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenarLstRs_new"))
End Sub

Public Sub pLlenaVsfGrid_new(ByVal vsfObj As VSFlexGrid, ByRef rs As Recordset, Optional ByVal L_Cargadatos As Boolean = True, Optional ByVal LRegnuevo As Boolean = True, Optional ByVal LTablaNueva As Boolean = True)
On Error GoTo NotificaError
  
  'Procedimiento que llena un vsflexgrid en forma dinamico de un recordset
  'Te llena el encabezado y el cuerpo del grid
  'Encabezado
  Dim lni  As Integer, vlintSeqFil As Integer, vlintSeqCol As Integer
  Dim c_cadena  As String
  If LTablaNueva Then
      vsfObj.Clear
      vsfObj.Cols = rs.Fields.Count
      vsfObj.Rows = 1
      For lni = 1 To rs.Fields.Count - 1
          If rs.Fields(lni).Name = "VOLUMÉN" Then
                vsfObj.TextMatrix(0, lni) = "VOLUMEN"
         Else
                vsfObj.TextMatrix(0, lni) = rs.Fields(lni).Name
         End If
          vsfObj.ColComboList(lni) = ""
          If rs.Fields(lni).Type = adCurrency Or rs.Fields(lni).Type = 10 Or rs.Fields(lni).Type = 3 Then
              vsfObj.ColWidth(lni) = rs.Fields(lni).DefinedSize * vsfObj.FontSize * 18
              vsfObj.ColDataType(lni) = flexDTCurrency
          ElseIf rs.Fields(lni).Type = adDate Or rs.Fields(lni).Type = 135 Then
              vsfObj.ColWidth(lni) = rs.Fields(lni).DefinedSize * vsfObj.FontSize * 7
              vsfObj.ColDataType(lni) = flexDTDate
          Else
              vsfObj.ColWidth(lni) = rs.Fields(lni).DefinedSize * vsfObj.FontSize * 9 + 250
              vsfObj.ColDataType(lni) = rs.Fields(lni).Type
          End If
      Next
  End If
  vsfObj.Col = 0
  'Detalle
    If L_Cargadatos Then
        Do While Not rs.EOF
            c_cadena = ""
            For lni = 0 To rs.Fields.Count - 1
                If rs.Fields(lni).Type = adBoolean Then
                    c_cadena = c_cadena & IIf(rs.Fields(lni).Value, 1, 0) & vbTab
                ElseIf rs.Fields(lni).Type = adCurrency Or rs.Fields(lni).Type = 10 Then
                    c_cadena = c_cadena & Format(rs.Fields(lni).Value, "##,###,##0.00") & vbTab
                Else
                    c_cadena = c_cadena & rs.Fields(lni).Value & vbTab
                End If
            Next
            c_cadena = Left(c_cadena, Len(c_cadena) - 1)
            vsfObj.AddItem c_cadena
            rs.MoveNext
        Loop
    End If
    
Exit Sub
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pLlenaVsfGrid_new"))
End Sub

Public Sub pSeleccionaLista_new(vlintElemento As Integer, vllstOrigen As ListBox, vllstDestino As ListBox, Optional vlcmdSelecciona As MyButton, Optional vlcmdDesSelecciona As MyButton)
    On Error GoTo NotificaError
    
    vllstDestino.AddItem vllstOrigen.List(vlintElemento)
    vllstDestino.ItemData(vllstDestino.NewIndex) = vllstOrigen.ItemData(vlintElemento)
    vllstOrigen.RemoveItem (vlintElemento)
    
    If Not (vlcmdSelecciona Is Nothing) Then
        If vllstOrigen.ListCount = 0 Then
            vlcmdSelecciona.Enabled = False
            vllstOrigen.Enabled = False
        Else
            vllstOrigen.ListIndex = 0
            vlcmdSelecciona.Enabled = True
            vllstOrigen.Enabled = True
        End If
    End If
    If Not (vlcmdDesSelecciona Is Nothing) Then
        If vllstDestino.ListCount = 0 Then
            vlcmdDesSelecciona.Enabled = False
            vllstDestino.Enabled = False
        Else
            vllstDestino.ListIndex = 0
            vlcmdDesSelecciona.Enabled = True
            vllstDestino.Enabled = True
        End If
    End If
    
Exit Sub
NotificaError:
  Call pRegistraError(err.Number, err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccionaLista_new"))
End Sub
