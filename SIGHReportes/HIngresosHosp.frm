VERSION 5.00
Begin VB.Form HIngresosHosp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingresos Hospitalarios por Departamentos y/o Servicios"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   45
      TabIndex        =   2
      Top             =   5880
      Width           =   7185
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HIngresosHosp.frx":0000
         DownPicture     =   "HIngresosHosp.frx":04C4
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
         Left            =   3728
         Picture         =   "HIngresosHosp.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HIngresosHosp.frx":0E9C
         DownPicture     =   "HIngresosHosp.frx":12FC
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
         Left            =   2198
         Picture         =   "HIngresosHosp.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5850
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   7215
      Begin VB.Frame Frame4 
         Caption         =   "Tipos de Número de Historia Clínica"
         Height          =   1125
         Left            =   1215
         TabIndex        =   25
         Top             =   4575
         Width           =   5880
         Begin VB.ComboBox cmbTipoHistoria 
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
            ItemData        =   "HIngresosHosp.frx":1BE6
            Left            =   1065
            List            =   "HIngresosHosp.frx":1BF3
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   230
            Width           =   4620
         End
         Begin VB.ComboBox cmbIdTipoGenHistoriaClinica 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   630
            Visible         =   0   'False
            Width           =   4620
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Considerar"
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
            Left            =   165
            TabIndex        =   28
            Top             =   285
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3450
         Left            =   1215
         TabIndex        =   8
         Top             =   990
         Width           =   5880
         Begin VB.ComboBox cmbIdServicio2 
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
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   2970
            Width           =   4020
         End
         Begin VB.ComboBox cmbIdEspecialidad2 
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
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2580
            Width           =   4020
         End
         Begin VB.ComboBox cmbIdDepartamento2 
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
            Left            =   1695
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2175
            Width           =   4020
         End
         Begin VB.ComboBox cmbIdServicio1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1305
            Width           =   4020
         End
         Begin VB.ComboBox cmbIdEspecialidad1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   900
            Width           =   4020
         End
         Begin VB.ComboBox cmbIdDepartamento1 
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
            Left            =   1710
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   495
            Width           =   4020
         End
         Begin VB.Label lblTitulo2 
            AutoSize        =   -1  'True
            Caption         =   "Servicio2"
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
            Left            =   60
            TabIndex        =   22
            Top             =   1860
            Width           =   720
         End
         Begin VB.Label lblServicio2 
            AutoSize        =   -1  'True
            Caption         =   "Servicio"
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
            Left            =   420
            TabIndex        =   21
            Top             =   3045
            Width           =   615
         End
         Begin VB.Label lblEspecialidad2 
            Caption         =   "Especialidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   20
            Top             =   2625
            Width           =   1395
         End
         Begin VB.Label lblDpto2 
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   420
            TabIndex        =   19
            Top             =   2235
            Width           =   1260
         End
         Begin VB.Label lblTitulo1 
            AutoSize        =   -1  'True
            Caption         =   "Servicio1"
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
            Left            =   105
            TabIndex        =   15
            Top             =   195
            Width           =   720
         End
         Begin VB.Label lblServicio1 
            Caption         =   "Servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   435
            TabIndex        =   14
            Top             =   1365
            Width           =   1275
         End
         Begin VB.Label lblEspecialidad1 
            Caption         =   "Especialidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   435
            TabIndex        =   13
            Top             =   945
            Width           =   1395
         End
         Begin VB.Label lblDpto1 
            Caption         =   "Departamento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   435
            TabIndex        =   12
            Top             =   555
            Width           =   1260
         End
      End
      Begin VB.ComboBox cmbTipoRep 
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
         ItemData        =   "HIngresosHosp.frx":1C5F
         Left            =   1230
         List            =   "HIngresosHosp.frx":1C6F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   615
         Width           =   5850
      End
      Begin VB.ComboBox cmbConsiderar 
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
         ItemData        =   "HIngresosHosp.frx":1CD9
         Left            =   4215
         List            =   "HIngresosHosp.frx":1CE3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   225
         Width           =   2865
      End
      Begin VB.ComboBox cmbAnio 
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
         ItemData        =   "HIngresosHosp.frx":1D04
         Left            =   1245
         List            =   "HIngresosHosp.frx":1D06
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Rep"
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
         Left            =   90
         TabIndex        =   7
         Top             =   705
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Espec"
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
         Left            =   3270
         TabIndex        =   5
         Top             =   285
         Width           =   900
      End
      Begin VB.Label Departamento 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   90
         TabIndex        =   1
         Top             =   285
         Width           =   330
      End
   End
End
Attribute VB_Name = "HIngresosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Ingreso Hospitalario
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbIdDepartamento1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad1 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdDepartamento2 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdServicio2 As New sighEntidades.ListaDespleglable
Dim mo_cmbIdEspecialidad2 As New sighEntidades.ListaDespleglable
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_Titulo As String
Dim ml_TextoDelFiltro As String
Dim mo_cmbIdTipoGenHistoriaClinica As New sighEntidades.ListaDespleglable
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRptIngresosHosp As New RptHIngresosHosp
        oRptIngresosHosp.Anio = Val(cmbAnio.Text)
        oRptIngresosHosp.TipoEspecialidad = cmbConsiderar.ListIndex
        oRptIngresosHosp.TipoReporte = cmbTipoRep.ListIndex
        oRptIngresosHosp.idDepartamento1 = IIf(mo_cmbIdDepartamento1.BoundText = "", 0, mo_cmbIdDepartamento1.BoundText)
        oRptIngresosHosp.idEspecialidad1 = IIf(mo_cmbIdEspecialidad1.BoundText = "", 0, mo_cmbIdEspecialidad1.BoundText)
        oRptIngresosHosp.idServicio1 = IIf(mo_cmbIdServicio1.BoundText = "", 0, mo_cmbIdServicio1.BoundText)
        oRptIngresosHosp.idDepartamento2 = IIf(mo_cmbIdDepartamento2.BoundText = "", 0, mo_cmbIdDepartamento2.BoundText)
        oRptIngresosHosp.idEspecialidad2 = IIf(mo_cmbIdEspecialidad2.BoundText = "", 0, mo_cmbIdEspecialidad2.BoundText)
        oRptIngresosHosp.idServicio2 = IIf(mo_cmbIdServicio2.BoundText = "", 0, mo_cmbIdServicio2.BoundText)
        oRptIngresosHosp.Titulo = ml_Titulo
        oRptIngresosHosp.TextoDelFiltro = ml_TextoDelFiltro
        oRptIngresosHosp.IdTipoNroHistoria = IIf(cmbTipoHistoria.ListIndex = 2, mo_cmbIdTipoGenHistoriaClinica.BoundText, IIf(cmbTipoHistoria.ListIndex = 0, 100, 200))
        oRptIngresosHosp.CrearReporte Me.hwnd
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    Dim sMensaje As String
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:   Año: " & cmbAnio.Text & ",   Tipo Especialidad: " & cmbConsiderar.Text & _
                        "           Tipo de Número de H.C: " & IIf(cmbTipoHistoria.ListIndex = 0, "Todos", IIf(cmbTipoHistoria.ListIndex = 1, "No considerar 'Temporal Alojamiento Conjunto'", cmbIdTipoGenHistoriaClinica.Text))
    
    Select Case cmbTipoRep.ListIndex
    Case 0
        ml_Titulo = "INGRESOS AL HOSPITAL"
    Case 1
        ml_Titulo = "INGRESOS AL HOSPITAL CONSOLIDANDO POR DEPARTAMENTOS"
    Case 2    'por un Servicio
        ml_Titulo = "INGRESOS AL HOSPITAL POR SERVICIO"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     Departamento: " & cmbIdDepartamento1.Text & "     Especialidad: " & cmbIdEspecialidad1.Text
    Case 3    'por 2 SERVICIOS
        ml_Titulo = "INGRESOS AL HOSPITAL CONSOLIDANDO DOS SERVICIOS"
        If cmbIdDepartamento1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdServicio1.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el primer Servicio)" + Chr(13)
        End If
        If cmbIdDepartamento2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Departamento (para  el segundo Servicio)" + Chr(13)
        End If
        If cmbIdEspecialidad2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija la Especialidad (para el segundo Servicio)" + Chr(13)
        End If
        If cmbIdServicio2.Text = "" Then
           sMensaje = sMensaje + "Por favor elija el Servicio (para el segundo Servicio)" + Chr(13)
        End If
        ml_TextoDelFiltro = ml_TextoDelFiltro & ",     SERVICIO1: (" & cmbIdDepartamento1.Text & ")/(" & cmbIdEspecialidad1.Text & ")/(" & cmbIdServicio1.Text & "),     SERVICIO2: (" & cmbIdDepartamento2.Text & ")/(" & cmbIdEspecialidad2.Text & ")/(" & cmbIdServicio2.Text & ")"
    End Select
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function






Private Sub cmbIdDepartamento1_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad1.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad1.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad1.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento1.BoundText))
       mo_cmbIdEspecialidad1.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub


Private Sub cmbIdDepartamento1_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento1
    AdministrarKeyPreview KeyCode
End Sub

Private Sub cmbIdDepartamento2_Click()
       Dim sMensaje As String
       mo_cmbIdEspecialidad2.BoundColumn = "IdEspecialidad"
       mo_cmbIdEspecialidad2.ListField = "DescripcionLarga"
       Set mo_cmbIdEspecialidad2.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(mo_cmbIdDepartamento2.BoundText))
       mo_cmbIdEspecialidad2.BoundText = ""
       If mo_AdminServiciosHosp.MensajeError <> "" Then
          MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
       End If
End Sub

Private Sub cmbIdDepartamento2_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdDepartamento2
    AdministrarKeyPreview KeyCode

End Sub





Private Sub cmbIdEspecialidad1_Click()
    mo_cmbIdServicio1.BoundColumn = "IdServicio"
    mo_cmbIdServicio1.ListField = "DescripcionLarga"
    If cmbConsiderar.ListIndex = 0 Then
       Set mo_cmbIdServicio1.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento1.BoundText), Val(mo_cmbIdEspecialidad1.BoundText))
    Else
       Set mo_cmbIdServicio1.RowSource = mo_ReglasFacturacion.ServiciosSeleccionarEmergenciaPorEspecialidad1(Val(mo_cmbIdEspecialidad1.BoundText))
    End If
End Sub



Private Sub cmbIdEspecialidad2_Click()
    mo_cmbIdServicio2.BoundColumn = "IdServicio"
    mo_cmbIdServicio2.ListField = "DescripcionLarga"
    If cmbConsiderar.ListIndex = 0 Then
       Set mo_cmbIdServicio2.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarPorTipoServicioDptoEspecialidad(3, Val(mo_cmbIdDepartamento2.BoundText), Val(mo_cmbIdEspecialidad2.BoundText))
    Else
       Set mo_cmbIdServicio2.RowSource = mo_ReglasFacturacion.ServiciosSeleccionarEmergenciaPorEspecialidad1(Val(mo_cmbIdEspecialidad2.BoundText))
    End If
End Sub



Private Sub cmbTipoHistoria_Click()
   If cmbTipoHistoria.ListIndex = 2 Then
      cmbIdTipoGenHistoriaClinica.Visible = True
   Else
      cmbIdTipoGenHistoriaClinica.Visible = False
   End If

End Sub

Private Sub cmbTipoRep_Change()
    Select Case cmbTipoRep.ListIndex
    Case 0, 1
        lblTitulo1.Visible = False
        lblDpto1.Visible = False
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = False
        cmbIdDepartamento1.Visible = False
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = False
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
    Case 2
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = False
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = False
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = False
        lblDpto2.Visible = False
        lblServicio2.Visible = False
        lblEspecialidad2.Visible = False
        cmbIdDepartamento2.Visible = False
        cmbIdServicio2.Visible = False
        cmbIdEspecialidad2.Visible = False
        lblTitulo1.Caption = "Elegir el Servicio:"
    Case 3
        lblTitulo1.Visible = True
        lblDpto1.Visible = True
        lblServicio1.Visible = True
        lblEspecialidad1.Visible = True
        cmbIdDepartamento1.Visible = True
        cmbIdServicio1.Visible = True
        cmbIdEspecialidad1.Visible = True
        lblTitulo2.Visible = True
        lblDpto2.Visible = True
        lblServicio2.Visible = True
        lblEspecialidad2.Visible = True
        cmbIdDepartamento2.Visible = True
        cmbIdServicio2.Visible = True
        cmbIdEspecialidad2.Visible = True
        lblTitulo1.Caption = "Elegir la primera Especialidad:"
        lblTitulo2.Caption = "Elegir la segunda Especialidad:"
    End Select
End Sub

Private Sub cmbTipoRep_Click()
    cmbTipoRep_Change
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdDepartamento1.MiComboBox = cmbIdDepartamento1
    Set mo_cmbIdEspecialidad1.MiComboBox = cmbIdEspecialidad1
    Set mo_cmbIdServicio1.MiComboBox = cmbIdServicio1
    Set mo_cmbIdDepartamento2.MiComboBox = cmbIdDepartamento2
    Set mo_cmbIdEspecialidad2.MiComboBox = cmbIdEspecialidad2
    Set mo_cmbIdServicio2.MiComboBox = cmbIdServicio2
    Set mo_cmbIdTipoGenHistoriaClinica.MiComboBox = cmbIdTipoGenHistoriaClinica

End Sub

Private Sub Form_Load()
    mo_Formulario.LlenaComboConAnios cmbAnio
    cmbConsiderar.ListIndex = 0
    cmbTipoRep.ListIndex = 0
    cmbTipoHistoria.ListIndex = 0
    CargaCombos
End Sub

Sub CargaCombos()
       mo_cmbIdDepartamento1.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento1.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento1.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()
       
       mo_cmbIdDepartamento2.BoundColumn = "IdDepartamento"
       mo_cmbIdDepartamento2.ListField = "DescripcionLarga"
       Set mo_cmbIdDepartamento2.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos()

        mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
        mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
        Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarTodos()
        mo_cmbIdTipoGenHistoriaClinica.BoundText = 2

 End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub




