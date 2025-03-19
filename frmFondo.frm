VERSION 5.00
Begin VB.Form frmfondo 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   15
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   15
   ControlBox      =   0   'False
   FillColor       =   &H80000001&
   ForeColor       =   &H80000001&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
End
Attribute VB_Name = "frmfondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------------------------------------------
''| Nombre del Proyecto      : Dietología
''| Nombre del Formulario    : frmfondo
''-------------------------------------------------------------------------------------
''| Objetivo: Muestra un fondo que permanece durante la ejecucion del sistema
''-------------------------------------------------------------------------------------
''| Análisis y Diseño        :
''| Autor                    :
''| Fecha de Creación        :
''| Modificó                 : Nombre(s)
''| Fecha última modificación: dd/mes/AAAA
''-------------------------------------------------------------------------------------
Private vlblnPrimeraVez As Boolean

Private Sub Form_Activate()
    On Error Resume Next
    Dim rsPeriodos As New ADODB.Recordset
    Dim frmMenu As Form
    Dim vlblnNuevo As Boolean
    Dim ctrl As Control
    Set frmMenu = frmMenuPrincipal
     
    For Each ctrl In frmMenuPrincipal.Controls
        If ctrl.Name = "lblUsuario" Then
            vlblnNuevo = True
        End If
    Next
    
    If cgstrModulo = "CN" Then
        If vlblnPrimeraVez Then
            frmEmpresasContables.Caption = "Sistema Integral Hospitalario - " + vgstrNombreUsuario + " - " + vgstrNombreDepartamento
            frmEmpresasContables.Show vbModal
        End If
        If Not vlblnNuevo Then
            frmMenuPrincipal.Caption = "Sistema Integral Hospitalario - " + vgstrNombreUsuario + " - " + vgstrNombreDepartamento
        End If
        frmMenuPrincipal.Show
    ElseIf cgstrModulo = "NO" Then
        If Not vgblnMenuNomina Then
            vgstrParametrosSP = "0|0" & "|" & vgintClaveEmpresaContable
            Set rsPeriodos = frsEjecuta_SP(vgstrParametrosSP, "sp_NoSelPeriodosActuales")
            If rsPeriodos.RecordCount > 0 Then
                Do While Not rsPeriodos.EOF
                    Select Case rsPeriodos!intIdTipoPeriodo
                        Case 1: vglngIDPeriodoActualSemanal = rsPeriodos!intIDPeriodo
                        Case 2: vglngIDPeriodoActualDecenal = rsPeriodos!intIDPeriodo
                        Case 3: vglngIDPeriodoActualQuincenal = rsPeriodos!intIDPeriodo
                        Case 4: vglngIDPeriodoActualExtraOrdinario = rsPeriodos!intIDPeriodo
                        Case 5: vglngIDPeriodoActualCatorcenal = rsPeriodos!intIDPeriodo
                    End Select
                    rsPeriodos.MoveNext
                Loop
                rsPeriodos.Close
                frmTipoNomina.vgblnCargarMenuPrincipal = True
                frmTipoNomina.Show
            Else
                rsPeriodos.Close
                frmInstalacionTiposdePeriodo.Show
            End If
        Else
            If Not vlblnNuevo Then
                frmMenuPrincipal.Caption = "Sistema Integral Hospitalario - " + vgstrNombreUsuario + " - " + vgstrNombreDepartamento
            End If
            frmMenuPrincipal.Show
        End If
    ElseIf cgstrModulo = "SI" And vgblnNomina Then
        vgstrParametrosSP = "0|0" & "|" & vgintClaveEmpresaContable
        Set rsPeriodos = frsEjecuta_SP(vgstrParametrosSP, "sp_NoSelPeriodosActuales")
        If rsPeriodos.RecordCount = 0 Then
            rsPeriodos.Close
            frmInstalacionTiposdePeriodo.Show
        Else
            Load frmMenuPrincipal
            If Not vlblnNuevo Then
                frmMenuPrincipal.Caption = "Sistema Integral Hospitalario - " + vgstrNombreUsuario + " - " + vgstrNombreDepartamento
            End If
            frmMenuPrincipal.Show
            If cgstrModulo = "SI" Then
                If First_time Then
                     frmParametrosSistema.Show
                End If
            End If
        End If
    Else
        Load frmMenuPrincipal
        
        If Not vlblnNuevo Then
            frmMenuPrincipal.Caption = "Sistema Integral Hospitalario - " + vgstrNombreUsuario + " - " + vgstrNombreDepartamento
        End If
        frmMenuPrincipal.Show
        If cgstrModulo = "SI" Then
            If First_time Then
                 frmParametrosSistema.Show
            End If
        End If
    End If
    vlblnPrimeraVez = False
    If vlblnNuevo Then
        frmMenu.lblUsuario.Caption = vgstrNombreUsuario + "@" + vgstrNombreDepartamento
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    vlblnPrimeraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    vlblnPrimeraVez = False
End Sub

