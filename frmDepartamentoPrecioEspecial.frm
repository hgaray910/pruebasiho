VERSION 5.00
Begin VB.Form frmDepartamentoPrecioEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos con precio especial"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDeptosAsignados 
      Height          =   4545
      Left            =   4140
      TabIndex        =   3
      ToolTipText     =   "Departamentos asignados"
      Top             =   150
      Width           =   3135
   End
   Begin VB.ListBox lstDeptos 
      Height          =   4545
      Left            =   135
      TabIndex        =   0
      ToolTipText     =   "Departamentos sin asignar"
      Top             =   150
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Height          =   1785
      Left            =   3375
      TabIndex        =   4
      Top             =   1530
      Width           =   660
      Begin VB.CommandButton cmdTodos 
         Height          =   495
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmDepartamentoPrecioEspecial.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1185
         Width           =   495
      End
      Begin VB.CommandButton cmdSelecciona 
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmDepartamentoPrecioEspecial.frx":02B2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Incluir un cargo al paquete"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdSelecciona 
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   75
         MaskColor       =   &H80000014&
         Picture         =   "frmDepartamentoPrecioEspecial.frx":042C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir un cargo al paquete"
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   45
         X2              =   900
         Y1              =   1665
         Y2              =   1665
      End
   End
End
Attribute VB_Name = "frmDepartamentoPrecioEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'| Nombre del Proyecto      : Caja                                                   -
'| Nombre del Formulario    : frmDepartamentoPrecioEspecial
'-------------------------------------------------------------------------------------
'| Objetivo: Registrar los parámetros en PvHorarioEmpresa
'-------------------------------------------------------------------------------------
'| Análisis y Diseño        : Rosenda Hernandez Anaya
'| Autor                    : Rosenda Hernandez Anaya
'| Fecha de Creación        : 04/Nov/2002
'| Fecha Terminación        : 04/Nov/2002
'| Modificó                 :
'| Fecha última modificación:
'| Descripción de la modificación:
'-------------------------------------------------------------------------------------

Dim vlstrX As String
Dim rs As New ADODB.Recordset
Dim rsPvDepartamentoPrecioEspecial As New ADODB.Recordset

Private Sub chkTodos_Click()
End Sub

Private Sub pSeleccion(Index As Integer, vllngPersonaGraba As Long)
    On Error GoTo NotificaError
    If Index = 0 Then
    
        With rsPvDepartamentoPrecioEspecial
            .AddNew
            !SMICVEDEPARTAMENTO = lstDeptos.ItemData(lstDeptos.ListIndex)
            .Update
        End With
        Call pGuardarLogTransaccion(Me.Name, EnmGrabar, vllngPersonaGraba, "DEPARTAMENTO CON PRECIO ESPECIAL", CStr(lstDeptos.ItemData(lstDeptos.ListIndex)))

    Else
        If fintLocalizaPkRs(rsPvDepartamentoPrecioEspecial, 0, CStr(lstDeptosAsignados.ItemData(lstDeptosAsignados.ListIndex))) <> 0 Then
            With rsPvDepartamentoPrecioEspecial
                .Delete
                .Update
            End With
        End If
        Call pGuardarLogTransaccion(Me.Name, EnmBorrar, vllngPersonaGraba, "DEPARTAMENTO CON PRECIO ESPECIAL", CStr(lstDeptosAsignados.ItemData(lstDeptosAsignados.ListIndex)))
    End If

    pCargaDeptos
    pCargaDeptosAsignados
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pSeleccion"))
End Sub

Private Sub cmdSelecciona_Click(Index As Integer)
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    
    '------------------------------------------------------------------
    ' Persona que graba
    '------------------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    pSeleccion Index, vllngPersonaGraba

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSelecciona_Click"))
End Sub

Private Sub cmdTodos_Click()
    On Error GoTo NotificaError
    Dim vllngPersonaGraba As Long
    
    If lstDeptos.ListCount = 0 Then Exit Sub
    '------------------------------------------------------------------
    ' Persona que graba
    '------------------------------------------------------------------
    vllngPersonaGraba = flngPersonaGraba(vgintNumeroDepartamento)
    If vllngPersonaGraba = 0 Then Exit Sub
    Do While lstDeptos.ListCount > 0
        lstDeptos.ListIndex = 0
        pSeleccion 0, vllngPersonaGraba
    Loop
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdTodos_Click"))
End Sub

Private Sub Form_Activate()
    vgstrNombreForm = Me.Name
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

    vlstrX = "select * from PvDepartamentoPrecioEspecial"
    Set rsPvDepartamentoPrecioEspecial = frsRegresaRs(vlstrX, adLockOptimistic, adOpenDynamic)

    pCargaDeptos
    pCargaDeptosAsignados
    
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub

Private Sub pCargaDeptosAsignados()
    On Error GoTo NotificaError
        
    lstDeptosAsignados.Clear
        
    vlstrX = "" & _
    "select " & _
        "NoDepartamento.vchDescripcion," & _
        "NoDepartamento.smiCveDepartamento " & _
    "From " & _
        "PvDepartamentoPrecioEspecial " & _
        "inner join NoDepartamento on PvDepartamentoPrecioEspecial.smiCveDepartamento=NoDepartamento.smiCveDepartamento " & _
    " Where nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & _
    "Order By " & _
        "NoDepartamento.vchDescripcion"

    Set rs = frsRegresaRs(vlstrX, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        pLlenarListRs lstDeptosAsignados, rs, 1, 0
    
        lstDeptosAsignados.ListIndex = 0
        
        cmdSelecciona(1).Enabled = True
    
    Else
        cmdSelecciona(1).Enabled = False
    End If


Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDeptosAsignados"))
End Sub

Private Sub pCargaDeptos()
    On Error GoTo NotificaError
    
    lstDeptos.Clear
    
    vlstrX = "" & _
    "select " & _
        "NoDepartamento.vchDescripcion," & _
        "NoDepartamento.smiCveDepartamento " & _
    "From " & _
        "NoDepartamento " & _
    "Where " & _
        "NoDepartamento.smiCveDepartamento not in (select smiCveDepartamento from PvDepartamentoPrecioEspecial) " & _
        "and (NoDepartamento.chrClasificacion='A' or NoDepartamento.chrEnfermeria='E') " & _
    "And nodepartamento.tnyclaveempresa = " & vgintClaveEmpresaContable & " Order By " & _
        "NoDepartamento.vchDescripcion"
    
    Set rs = frsRegresaRs(vlstrX, adLockOptimistic, adOpenDynamic)
    If rs.RecordCount <> 0 Then
        pLlenarListRs lstDeptos, rs, 1, 0
        
        lstDeptos.ListIndex = 0
        
        cmdSelecciona(0).Enabled = True
        cmdTodos.Enabled = True
    Else
        cmdSelecciona(0).Enabled = False
        cmdTodos.Enabled = False
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":pCargaDeptos"))
End Sub

Private Sub lstDeptos_DblClick()
    On Error GoTo NotificaError
    
    If lstDeptos.ListCount <> 0 Then
        cmdSelecciona_Click 0
    End If
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstDeptos_DblClick"))
End Sub

Private Sub lstDeptosAsignados_DblClick()
    On Error GoTo NotificaError
    
    If lstDeptosAsignados.ListCount <> 0 Then
        cmdSelecciona_Click 1
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstDeptosAsignados_DblClick"))
End Sub
