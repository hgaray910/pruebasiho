VERSION 5.00
Begin VB.Form frmConceptosExcluidosAreaP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de facturación excluídos por área de productividad"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3800
      Left            =   80
      TabIndex        =   4
      Top             =   0
      Width           =   6450
      Begin VB.Frame Frame2 
         Caption         =   "Conceptos excluídos"
         Height          =   1850
         Left            =   140
         TabIndex        =   8
         Top             =   1800
         Width           =   6150
         Begin VB.ListBox lstConceptosExcluidos 
            Height          =   1425
            Left            =   100
            Sorted          =   -1  'True
            TabIndex        =   3
            ToolTipText     =   "Conceptos excluidos"
            Top             =   300
            Width           =   5920
         End
      End
      Begin VB.ComboBox cboAreasProductividad 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Área de productividad"
         Top             =   540
         Width           =   6200
      End
      Begin VB.ComboBox cboConceptos 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Conceptos disponibles"
         Top             =   1335
         Width           =   5600
      End
      Begin VB.CommandButton cmdExcluirConcepto 
         Height          =   450
         Left            =   5835
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmConceptosExcluidosAreaP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluír concepto"
         Top             =   1190
         UseMaskColor    =   -1  'True
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Área de productividad"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Conceptos disponibles"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1035
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   3060
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmConceptosExcluidosAreaP.frx":04F2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar"
      Top             =   3900
      Width           =   495
   End
End
Attribute VB_Name = "frmConceptosExcluidosAreaP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public llngNumOpcion As Long

Private Sub cboAreasProductividad_Click()
On Error GoTo NotificaError

Dim rsConceptosArea As New ADODB.Recordset
Dim rsConceptosExcluidos As New ADODB.Recordset

    vgstrParametrosSP = cboAreasProductividad.ItemData(cboAreasProductividad.ListIndex) & "|" & vgintClaveEmpresaContable
    Set rsConceptosArea = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelConceptoAreaProd")
    pLlenarCboRs cboConceptos, rsConceptosArea, 0, 1
    cboConceptos.ListIndex = 0
    
    lstConceptosExcluidos.Clear
    vgstrParametrosSP = cboAreasProductividad.ItemData(cboAreasProductividad.ListIndex) & "|" & vgintClaveEmpresaContable
    Set rsConceptosExcluidos = frsEjecuta_SP(vgstrParametrosSP, "sp_PvSelConceptoExclAreaProd")
    
    Do While Not rsConceptosExcluidos.EOF
        lstConceptosExcluidos.AddItem rsConceptosExcluidos!Concepto
        lstConceptosExcluidos.ItemData(lstConceptosExcluidos.NewIndex) = rsConceptosExcluidos!CveConcepto
        rsConceptosExcluidos.MoveNext
    Loop

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cboAreasProductividad_Click"))
End Sub

Private Sub cmdExcluirConcepto_Click()
On Error GoTo NotificaError

    lstConceptosExcluidos.AddItem cboConceptos.List(cboConceptos.ListIndex)
    lstConceptosExcluidos.ItemData(lstConceptosExcluidos.NewIndex) = cboConceptos.ItemData(cboConceptos.ListIndex)

    cboConceptos.RemoveItem (cboConceptos.ListIndex)
    cboConceptos.ListIndex = 0
    
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdExcluirConcepto_Click"))
End Sub

Private Sub cmdSave_Click()
On Error GoTo NotificaError
Dim lngContador As Long
Dim lngPersonagraba As Long


    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "E", True) Then
    
        lngPersonagraba = flngPersonaGraba(vgintNumeroDepartamento)
        If lngPersonagraba = 0 Then Exit Sub
        
        EntornoSIHO.ConeccionSIHO.BeginTrans
        
        'Borra los conceptos excluídos del área, de la empresa contable a la que pertenece el departamento logueado
        vgstrParametrosSP = CStr(cboAreasProductividad.ItemData(cboAreasProductividad.ListIndex)) & "|" & CStr(vgintClaveEmpresaContable)
        frsEjecuta_SP vgstrParametrosSP, "sp_PvDelConceptoExcluidoAreaP"
        
        For lngContador = 0 To lstConceptosExcluidos.ListCount - 1
            'Graba los conceptos excluídos
            vgstrParametrosSP = cboAreasProductividad.ItemData(cboAreasProductividad.ListIndex) & "|" & lstConceptosExcluidos.ItemData(lngContador) & "|" & vgintClaveEmpresaContable
            frsEjecuta_SP vgstrParametrosSP, "sp_PvInsConceptoExcluidoAreaP"
        
            Call pGuardarLogTransaccion(Me.Name, EnmGrabar, lngPersonagraba, "CONCEPTO EXCLUIDO POR AREA", cboAreasProductividad.ItemData(cboAreasProductividad.ListIndex))
        
        Next lngContador
        
        EntornoSIHO.ConeccionSIHO.CommitTrans
        
        'La operación se realizó satisfactoriamente.
        MsgBox SIHOMsg(420), vbInformation, "Mensaje"
    
    Else
        'El usuario no tiene permiso para grabar datos
        MsgBox SIHOMsg(65), vbOKOnly + vbExclamation, "Mensaje"
    End If

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":cmdSave_Click"))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            SendKeys vbTab
        Case vbKeyEscape
            Unload Me
    End Select
    
End Sub



Private Sub Form_Load()
On Error GoTo NotificaError

Me.Icon = frmMenuPrincipal.Icon

cboAreasProductividad.AddItem "PACIENTES EXTERNOS ATENDIDOS", 0
cboAreasProductividad.ItemData(0) = 1
cboAreasProductividad.AddItem "PACIENTES INGRESADOS AL HOSPITAL", 1
cboAreasProductividad.ItemData(1) = 2
cboAreasProductividad.AddItem "PACIENTES REFERIDOS A FARMACIA", 2
cboAreasProductividad.ItemData(2) = 3
cboAreasProductividad.AddItem "PACIENTES REFERIDOS A IMAGENOLOGÍA", 3
cboAreasProductividad.ItemData(3) = 4
cboAreasProductividad.AddItem "PACIENTES REFERIDOS A LABORATORIO", 4
cboAreasProductividad.ItemData(4) = 5
cboAreasProductividad.ListIndex = 0

Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":Form_Load"))
End Sub


Private Sub lstConceptosExcluidos_DblClick()
On Error GoTo NotificaError

    If fblnRevisaPermiso(vglngNumeroLogin, llngNumOpcion, "C", True) Then
        cboConceptos.AddItem lstConceptosExcluidos.List(lstConceptosExcluidos.ListIndex)
        cboConceptos.ItemData(cboConceptos.NewIndex) = lstConceptosExcluidos.ItemData(lstConceptosExcluidos.ListIndex)
        
        lstConceptosExcluidos.RemoveItem (lstConceptosExcluidos.ListIndex)
    Else
        '¡El usuario debe tener permiso de control total para eliminar los datos!
        MsgBox SIHOMsg(810), vbOKOnly + vbExclamation, "Mensaje"
    End If
Exit Sub
NotificaError:
    Call pRegistraError(Err.Number, Err.Description, cgstrModulo, (vgstrNombreForm & ":lstConceptosExcluidos_DblClick"))
End Sub
