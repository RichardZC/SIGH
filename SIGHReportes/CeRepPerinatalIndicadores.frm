VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form CeRepPerinatalIndicadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indicadores Módulo Niño Sano"
   ClientHeight    =   4620
   ClientLeft      =   8910
   ClientTop       =   5520
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   11415
   Begin VB.Frame frOpcionesFiltro 
      Caption         =   "Opciones de Busqueda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   11355
      Begin VB.ComboBox cmbIdDepartamentoDomicilio 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1770
      End
      Begin VB.ComboBox cmbIdProvinciaDomicilio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4140
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   735
         Width           =   2745
      End
      Begin VB.ComboBox cmbIdDistritoDomicilio 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7710
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   735
         Width           =   3495
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7125
         TabIndex        =   17
         Top             =   780
         Width           =   570
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pro&vincia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3435
         TabIndex        =   16
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Depar&tamento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Reporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   810
      End
   End
   Begin VB.Frame fraDatosHistoria 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11355
      Begin Threed.SSOption ssOptIndicadores 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Porcentaje de niñas y niños de 06 a 35 meses de edad con Suplemento de hierro"
         Value           =   -1
      End
      Begin Threed.SSOption ssOptIndicadores 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Porcentaje de niñas y niños menores de 24 meses de edad con vacuna contra rotavirus y neumococo"
      End
      Begin Threed.SSOption ssOptIndicadores 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Porcentaje de niñas y niños menores de 6 meses de edad con Lactancia Materna Exclusiva"
      End
      Begin Threed.SSOption ssOptIndicadores 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   873
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Porcentaje de madres de niñas y niños menores de 12 meses de edad que han participado de 2 sesiones demostrativas"
      End
      Begin Threed.SSOption ssOptIndicadores 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Porcentaje de niñas y niños menores de 36 meses de edad con CRED completo de acuerdo a su edad"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   12
      Top             =   3390
      Width           =   11355
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CeRepPerinatalIndicadores.frx":0000
         DownPicture     =   "CeRepPerinatalIndicadores.frx":04C4
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   5580
         Picture         =   "CeRepPerinatalIndicadores.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CeRepPerinatalIndicadores.frx":0E9C
         DownPicture     =   "CeRepPerinatalIndicadores.frx":12FC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   4080
         Picture         =   "CeRepPerinatalIndicadores.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CeRepPerinatalIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte para el módulo Perinatal
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Cadena As New sighEntidades.Cadena
Dim mo_Formulario As New sighEntidades.Formulario

Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_perinatalIndicadores As New clCePerinatalIndicadores

Dim oRsDptoDomicilio As New Recordset

Dim mo_cmbIdDepartamentoDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdProvinciaDomicilio As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDistritoDomicilio As New sighEntidades.ListaDespleglable


Private Sub btnAceptar_Click()
On Error GoTo miError
    Dim i As Integer
    If ValidarDatosObligaorios() = False Then
        Exit Sub
    End If
    Me.MousePointer = 11
    Dim oProcesaReporte As New clCePerinatal
    Dim sTituloReporte As String, sFiltroAplicados As String
    
    mo_perinatalIndicadores.FechaReporte = txtFdesde.Text
    mo_perinatalIndicadores.IdDepartamento = Val(mo_cmbIdDepartamentoDomicilio.BoundText)
    mo_perinatalIndicadores.IdProvincia = Val(mo_cmbIdProvinciaDomicilio.BoundText)
    mo_perinatalIndicadores.IdDistrito = Val(mo_cmbIdDistritoDomicilio.BoundText)
    
    For i = 0 To ssOptIndicadores.Count - 1
        If ssOptIndicadores(i).Value = True Then
            Exit For
        End If
    Next i
    
    sTituloReporte = ssOptIndicadores(i).Caption
    
    sFiltroAplicados = "Fecha :" & Format(txtFdesde.Text, "dd/mm/yyyy")
    
    If Val(mo_cmbIdDepartamentoDomicilio.BoundText) > 0 And Val(mo_cmbIdProvinciaDomicilio.BoundText) > 0 And Val(mo_cmbIdDistritoDomicilio.BoundText) > 0 Then
        sFiltroAplicados = sFiltroAplicados & ", Dpto :" & Me.cmbIdDepartamentoDomicilio.Text & ", Prov. :" & Me.cmbIdProvinciaDomicilio.Text & ", Dist. :" & Me.cmbIdDistritoDomicilio.Text
    ElseIf Val(mo_cmbIdDepartamentoDomicilio.BoundText) > 0 And Val(mo_cmbIdProvinciaDomicilio.BoundText) > 0 Then
        sFiltroAplicados = sFiltroAplicados & ", Dpto :" & Me.cmbIdDepartamentoDomicilio.Text & ", Prov. :" & Me.cmbIdProvinciaDomicilio.Text
     ElseIf Val(mo_cmbIdDepartamentoDomicilio.BoundText) > 0 Then
        sFiltroAplicados = sFiltroAplicados & ", Dpto :" & Me.cmbIdDepartamentoDomicilio.Text
    End If
    
    Select Case i
        Case 0:
            mo_perinatalIndicadores.reporteDeCREDCompleto sTituloReporte, sFiltroAplicados, Me.hwnd
        Case 1:
            mo_perinatalIndicadores.reporteDeSuplementoDeHierro sTituloReporte, sFiltroAplicados, Me.hwnd
        Case 2:
            mo_perinatalIndicadores.reporteDeLactanciaMaternaExclusiva sTituloReporte, sFiltroAplicados, Me.hwnd
        Case 3:
            mo_perinatalIndicadores.reporteDeVacunaRotavirusNeumococo sTituloReporte, sFiltroAplicados, Me.hwnd
        Case 4:
            mo_perinatalIndicadores.reporteDeSesionesDemostrativas sTituloReporte, sFiltroAplicados, Me.hwnd
        Case Else
            MsgBox "No ha elegido ninguna opción de reporte", vbInformation, "Reportes CRED"
            Exit Sub
    End Select
    Set oProcesaReporte = Nothing
    Me.MousePointer = 1
miError:
    If Err Then
        MsgBox Err.Number & " : " & Err.Description, vbExclamation, "Reporte de Módulo Niño Sano"
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub cmbIdDepartamentoDomicilio_Click()
    If cmbIdDepartamentoDomicilio.ListIndex = -1 Then Exit Sub
       
    mo_cmbIdProvinciaDomicilio.BoundColumn = "IdProvincia"
    mo_cmbIdProvinciaDomicilio.ListField = "Nombre"
    On Error Resume Next
    Set mo_cmbIdProvinciaDomicilio.RowSource = mo_AdminServiciosGeograficos.ProvinciasSeleccionarPorDepartamento(Val(cmbIdDepartamentoDomicilio.ItemData(cmbIdDepartamentoDomicilio.ListIndex)))
         
    mo_cmbIdProvinciaDomicilio.BoundText = ""
    mo_cmbIdDistritoDomicilio.BoundText = ""
    cmbIdProvinciaDomicilio.Enabled = True
End Sub

Private Sub cmbIdDepartamentoDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamentoDomicilio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDistritoDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDistritoDomicilio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdProvinciaDomicilio_Click()
    If cmbIdProvinciaDomicilio.ListIndex = -1 Then Exit Sub
       
    mo_cmbIdDistritoDomicilio.BoundColumn = "IdDistrito"
    mo_cmbIdDistritoDomicilio.ListField = "Nombre"
    Set mo_cmbIdDistritoDomicilio.RowSource = mo_AdminServiciosGeograficos.DistritoSeleccionarPorProvincia(Val(cmbIdProvinciaDomicilio.ItemData(cmbIdProvinciaDomicilio.ListIndex)))
    
    If mo_AdminServiciosGeograficos.MensajeError <> "" Then
         MsgBox mo_AdminServiciosGeograficos.MensajeError, vbInformation, "Datos de paciente"
    End If
    
    mo_cmbIdDistritoDomicilio.BoundText = ""
    cmbIdDistritoDomicilio.Enabled = True
End Sub

Private Sub cmbIdProvinciaDomicilio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdProvinciaDomicilio
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    txtFdesde.Text = Date
    Set mo_cmbIdDepartamentoDomicilio.MiComboBox = Me.cmbIdDepartamentoDomicilio
    Set mo_cmbIdProvinciaDomicilio.MiComboBox = Me.cmbIdProvinciaDomicilio
    Set mo_cmbIdDistritoDomicilio.MiComboBox = Me.cmbIdDistritoDomicilio
    ConfigurarComboBoxes
End Sub


Public Sub ConfigurarComboBoxes()
Dim sMensaje As String
        
        mo_cmbIdDepartamentoDomicilio.BoundColumn = "IdDepartamento"
        mo_cmbIdDepartamentoDomicilio.ListField = "Nombre"
        mo_cmbIdDepartamentoDomicilio.BoundText = Left(lcBuscaParametro.SeleccionaFilaParametro(242), 2)
        sMensaje = sMensaje + mo_AdminServiciosGeograficos.MensajeError
        Set oRsDptoDomicilio = mo_AdminServiciosGeograficos.DepartamentosSeleccionarTodos() 'oRsDptoNacimiento.Clone()
        Set mo_cmbIdDepartamentoDomicilio.RowSource = oRsDptoDomicilio

        If sMensaje <> "" Then
            MsgBox sMensaje, vbInformation, "Datos de paciente"
        End If


End Sub

Private Sub ssOptIndicadores_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, ssOptIndicadores
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtFdesde_LostFocus()
    If txtFdesde.Text <> sighEntidades.FECHA_VACIA_DMY Then
        If Not sighEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde.Text = sighEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub


Public Function ValidarDatosObligaorios() As Boolean
    ValidarDatosObligaorios = False
    If txtFdesde.Text = sighEntidades.FECHA_VACIA_DMY Then
        MsgBox "Ingrese Fecha de Reporte", vbInformation, Me.Caption
        Exit Function
    End If
    ValidarDatosObligaorios = True
End Function

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
