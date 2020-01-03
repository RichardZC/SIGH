VERSION 5.00
Begin VB.Form frmDetalleProgramacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Programación"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
   Icon            =   "frmDetalleProgramacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFechaFinal 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.ComboBox cmbServicio 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtEspecialidad 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   4065
      End
      Begin VB.ComboBox cmbTurno 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3360
         Width           =   4125
      End
      Begin VB.TextBox txtFechaInicial 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbEstablecimiento 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   4080
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Turno"
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
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   1665
      End
      Begin VB.Label lblServicio 
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Final"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblFechaInicial 
         Caption         =   "Fecha Inicial"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Establecimiento"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   3840
      Width           =   4335
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F3)"
         DisabledPicture =   "frmDetalleProgramacion.frx":000C
         DownPicture     =   "frmDetalleProgramacion.frx":046C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   780
         Picture         =   "frmDetalleProgramacion.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleProgramacion.frx":0D56
         DownPicture     =   "frmDetalleProgramacion.frx":121A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2250
         Picture         =   "frmDetalleProgramacion.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmDetalleProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica en donde se hara la programacion del HIS para los responsables de MR.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuario As Long
'Dim ml_IdServicio As Long
Dim ms_IdMedico As String
Dim ml_IdEspecialidad As Long
Dim ms_FechaIncial As String
Dim ms_FechaFinal As String
Dim ms_DescripcionEspecialidad As String
Dim ms_IdHisProgMedEstMR As Long

Dim mo_Formulario As New SIGHEntidades.Formulario
Dim mo_cmbEstablecimiento As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicio As New SIGHEntidades.ListaDespleglable
Dim mo_cmbTurno As New SIGHEntidades.ListaDespleglable
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos
Dim mi_Opcion As sghOpciones

Dim oTablaProgramacionMed As New DOHIS_ProgMedEstMR

'========================================== PROPIEDADES ========================================
Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property
Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property
Property Let FechaInicial(sValue As String)
   ms_FechaIncial = sValue
End Property
Property Let FechaFinal(sValue As String)
   ms_FechaFinal = sValue
End Property
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Let IdMedico(sValue As String)
   ms_IdMedico = sValue
End Property
Property Let IdEspecialidad(lValue As Long)
   ml_IdEspecialidad = lValue
End Property
Property Let DescripcionEspecialidad(sValue As String)
   ms_DescripcionEspecialidad = sValue
End Property
Property Let IdHisProgMedEstMR(sValue As Long)
   ms_IdHisProgMedEstMR = sValue
End Property

'========================================== EVENTOS ========================================
Private Sub Form_Load()
Set mo_cmbEstablecimiento.MiComboBox = Me.cmbEstablecimiento
Set mo_cmbIdServicio.MiComboBox = Me.cmbServicio
Set mo_cmbTurno.MiComboBox = Me.cmbTurno
Me.txtEspecialidad.Text = ms_DescripcionEspecialidad
mo_Formulario.HabilitarDeshabilitar Me.txtEspecialidad, False
mo_Formulario.HabilitarDeshabilitar Me.txtFechaInicial, False
mo_Formulario.HabilitarDeshabilitar Me.txtFechaFinal, False
CargarComboBoxes
Select Case mi_Opcion
    Case sghAgregar
        CargarDatosAlFormulario
        Me.Caption = "Agregar Programación"
    Case sghModificar
        CargarDatosProgramacion
        Me.Caption = "Modificar Programación"
    Case sghConsultar
        mo_Formulario.HabilitarDeshabilitar Me.cmbEstablecimiento, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbServicio, False
        mo_Formulario.HabilitarDeshabilitar Me.cmbTurno, False
        CargarDatosProgramacion
        Me.Caption = "Consultar Programación"
        Me.btnAceptar.Enabled = False
End Select
End Sub

Sub CargarDatosProgramacion()
    Dim mo_ReglasHIS As New ReglasHISGalenos
    Dim DOHIS_ProgMedEstMR As New DOHIS_ProgMedEstMR
    DOHIS_ProgMedEstMR.IdHisProgMedEstMR = ms_IdHisProgMedEstMR
    If mo_ReglasHIS.ConsultarProgramacionMedicaHIS(DOHIS_ProgMedEstMR) Then
        mo_cmbEstablecimiento.BoundText = DOHIS_ProgMedEstMR.IdEstablecimiento
        Me.lblFechaInicial.Caption = "Fecha Programada"
        Me.txtFechaFinal.Visible = False
        Me.txtFechaInicial.Text = DOHIS_ProgMedEstMR.FechaProgramada
        mo_cmbIdServicio.BoundText = DOHIS_ProgMedEstMR.IdServicio
        mo_cmbTurno.BoundText = DOHIS_ProgMedEstMR.IdTurno
    End If
    Set mo_ReglasHIS = Nothing
    Set DOHIS_ProgMedEstMR = Nothing
End Sub

Private Sub cmbEstablecimiento_Click()
Dim oRcs_ServiciosVinculadosAlaEspecialidad As New ADODB.Recordset
mo_cmbIdServicio.BoundColumn = "IdServicio"
mo_cmbIdServicio.ListField = "Nombre"
'verificar ya que los servicios no pueden estar activos en la central, pero en la CS puede estarlo.
'Set mo_cmbIdServicio.RowSource = mo_AdminServiciosHosp.ServiciosSeleccionarConsultoriosPorEspecialidad(ml_IdEspecialidad, sghFiltraSoloActivos)
Set oRcs_ServiciosVinculadosAlaEspecialidad = mo_ReglasHIS.ListaServiciosPorEstablecimientoYEspecialidad(ml_IdEspecialidad, Val(mo_cmbEstablecimiento.BoundText))

If oRcs_ServiciosVinculadosAlaEspecialidad.RecordCount <> 0 Then
    oRcs_ServiciosVinculadosAlaEspecialidad.MoveFirst
    Set mo_cmbIdServicio.RowSource = oRcs_ServiciosVinculadosAlaEspecialidad
Else
    Call MsgBox("No Hay Servicios en el establecimiento vinculados a ala especialidad del Medico.", vbExclamation Or vbSystemModal, Me.Caption)
End If

End Sub

Private Sub cmbEstablecimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbServicio
End Sub

Private Sub cmbTurno_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, btnAceptar
End Sub

Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        If ValidarReglas Then
            GrabarProgramacion
            mi_BotonPresionado = sghAceptar
            Me.Hide
        End If
    End If
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Hide
End Sub

'========================================== METODOS ========================================
Sub CargarComboBoxes()
    mo_cmbEstablecimiento.BoundColumn = "IdEstablecimiento"
    mo_cmbEstablecimiento.ListField = "NombreEstablecimiento"
    Set mo_cmbEstablecimiento.RowSource = mo_ReglasHIS.HisObtenerListaEstablecimientosMRPorEspecialidad(ml_IdEspecialidad)
    
    mo_cmbTurno.BoundColumn = "IdHisTurno"
    mo_cmbTurno.ListField = "Descripcion"
    Set mo_cmbTurno.RowSource = mo_ReglasHIS.ListaTurnos
    
    cmbTurno.ListIndex = 0
End Sub
Sub CargarDatosAlFormulario()
    Me.txtFechaInicial.Text = ms_FechaIncial
    If ms_FechaFinal = "0" Then
        Me.lblFechaInicial.Caption = "Fecha Programada"
        Me.txtFechaFinal.Visible = False
    Else
        Me.lblFechaInicial.Caption = "Fecha Inicial"
        Me.txtFechaFinal.Text = ms_FechaFinal
    End If
    Me.txtEspecialidad.Text = ms_DescripcionEspecialidad
End Sub

Sub GrabarProgramacion()
    Dim mi_FechasInicial As Date
    Dim mi_FechasFinal As Date
    Dim mi_DiasProgramados As New Collection
    
    oTablaProgramacionMed.IdEstablecimiento = Val(mo_cmbEstablecimiento.BoundText)
    oTablaProgramacionMed.IdMedico = ms_IdMedico
    oTablaProgramacionMed.IdServicio = Val(mo_cmbIdServicio.BoundText)
    
    mi_FechasInicial = CDate(Me.txtFechaInicial.Text)
    If Me.txtFechaFinal.Text <> "" Then
        mi_FechasFinal = CDate(Me.txtFechaFinal.Text)
    Else
        mi_FechasFinal = CDate(Me.txtFechaInicial.Text)
    End If
    
    oTablaProgramacionMed.IdTurno = Val(mo_cmbTurno.BoundText)
    oTablaProgramacionMed.IdUsuarioAuditoria = ml_IdUsuario
    
    Dim dia As Date
    For dia = mi_FechasInicial To mi_FechasFinal
        mi_DiasProgramados.Add dia
    Next
    
    Select Case mi_Opcion
        Case sghAgregar
            If mo_ReglasHIS.IngresarRegistroProgramacionMedica(oTablaProgramacionMed, mi_DiasProgramados) Then
                Call MsgBox("Se ingreso la programación satisfactoriamente.", vbInformation, Me.Caption)
            Else
                 Call MsgBox("No se pudo ingresar la programación, Verifique su información.", vbInformation, Me.Caption)
            End If
        Case sghModificar
            oTablaProgramacionMed.IdHisProgMedEstMR = ms_IdHisProgMedEstMR
            If mo_ReglasHIS.ModificarRegistroProgramacionMedica(oTablaProgramacionMed, mi_DiasProgramados) Then
                Call MsgBox("Se actualizó la programación satisfactoriamente.", vbInformation, Me.Caption)
            Else
                 Call MsgBox("No se pudo actualizar la programación, Verifique su información.", vbInformation, Me.Caption)
            End If
    End Select
End Sub

''Verifica los dias programados del responsable a programar
'Private Function ValidarDatos(mo_FechasDuplicadasResultado As Collection) As Boolean
'Dim DiaProgramado As Date
'Dim mo_FechasProgramadasUsuario As New Collection
'Dim mo_FechasProgramadasSistema As New Collection
'Dim mo_FechasDuplicadas As New Collection
'Dim oRcs_DiasProgramados As New ADODB.Recordset
'
''Lee Lista (Dias) ingresada por el Usuario
'DiaProgramado = oTablaProgramacionMed.FechaInicial
'Do While True
'    If DiaProgramado = oTablaProgramacionMed.FechaFinal Then
'        mo_FechasProgramadasUsuario.Add DiaProgramado
'        Exit Do
'    Else
'        mo_FechasProgramadasUsuario.Add DiaProgramado
'        DiaProgramado = DiaProgramado + 1
'    End If
'Loop
'
''Lee Lista (Dias) Programados en el Sistema
' SE CAMBIO A UN BOOLEANO SE RENOMBRARA LISTAR PROGRAMACION MEDICA
''Set oRcs_DiasProgramados = mo_ReglasHIS.ValidarProgramacionMedica_FechasMesActual(oTablaProgramacionMed.IdMedico, _
''                                                                                  oTablaProgramacionMed.IdEstablecimiento, _
''                                                                                  oTablaProgramacionMed.IdServicio, _
''                                                                                  Month(DiaProgramado), _
''                                                                                  Year(DiaProgramado))
'
'Set oRcs_DiasProgramados = mo_ReglasHIS.ValidarProgramacionMedica_FechasMesActual(oTablaProgramacionMed.IdMedico, _
'                                                                                  Month(DiaProgramado), _
'                                                                                  Year(DiaProgramado))
'If oRcs_DiasProgramados.RecordCount <> 0 Then
'    Dim md_FechaInicial As Date
'    Dim md_FechaFinal As Date
'
'    oRcs_DiasProgramados.MoveFirst
'
'    Do While Not oRcs_DiasProgramados.EOF
'        md_FechaInicial = CDate(oRcs_DiasProgramados!FechaInicial)
'        md_FechaFinal = CDate(oRcs_DiasProgramados!FechaFinal)
'        DiaProgramado = md_FechaInicial
'        Do While True
'            If DiaProgramado = md_FechaFinal Then
'                mo_FechasProgramadasSistema.Add DiaProgramado
'                Exit Do
'            Else
'                mo_FechasProgramadasSistema.Add DiaProgramado
'                DiaProgramado = DiaProgramado + 1
'            End If
'        Loop
'
'        oRcs_DiasProgramados.MoveNext
'    Loop
'
'    'Comparacion de las Listas Obtenidas
'    Dim i, x As Integer
'    For i = 1 To mo_FechasProgramadasSistema.Count
'        For x = 1 To mo_FechasProgramadasUsuario.Count
'            If CDate(mo_FechasProgramadasSistema.Item(i)) = CDate(mo_FechasProgramadasUsuario(x)) Then
'                mo_FechasDuplicadas.Add CDate(mo_FechasProgramadasUsuario(x))
'            End If
'        Next
'    Next
'
'    If mo_FechasDuplicadas.Count > 0 Then
'        Set mo_FechasDuplicadasResultado = mo_FechasDuplicadas
'        ValidarDatos = False
'    Else
'        ValidarDatos = True
'    End If
'Else
'    ValidarDatos = True
'End If
'End Function

Private Function ValidaDatosObligatorios() As Boolean
    ValidaDatosObligatorios = True
    
    If Me.cmbEstablecimiento.Text = "" Then
        Call MsgBox("No ha escogido el establecimiento, elija uno.", vbExclamation, Me.Caption)
        ValidaDatosObligatorios = False
        Exit Function
    End If
    
    If Me.cmbServicio.Text = "" Then
        Call MsgBox("No ha escogido el servicio, elija uno.", vbExclamation, Me.Caption)
        ValidaDatosObligatorios = False
        Exit Function
    End If
End Function

Private Function ValidarReglas() As Boolean
ValidarReglas = True
End Function

Sub LimpiarVariablesDeMemoria()
    Set mo_Formulario = Nothing
    Set mo_cmbEstablecimiento = Nothing
    Set mo_cmbIdServicio = Nothing
    Set mo_cmbTurno = Nothing
    Set mo_ReglasHIS = Nothing
    Set oTablaProgramacionMed = Nothing
End Sub


