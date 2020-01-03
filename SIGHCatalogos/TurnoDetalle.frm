VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form TurnoDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "TurnoDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   1080
      Left            =   5535
      TabIndex        =   15
      Top             =   555
      Visible         =   0   'False
      Width           =   4245
      Begin MSDataListLib.DataCombo cmbIdEspecialidad 
         Height          =   315
         Left            =   1470
         TabIndex        =   16
         Top             =   615
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Especialidad"
         Height          =   285
         Left            =   300
         TabIndex        =   17
         Top             =   645
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Valores por defecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   60
      TabIndex        =   12
      Top             =   1170
      Width           =   5265
      Begin VB.ComboBox cmbTipoActividad 
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
         ItemData        =   "TurnoDetalle.frx":000C
         Left            =   1395
         List            =   "TurnoDetalle.frx":001F
         TabIndex        =   5
         Top             =   1065
         Visible         =   0   'False
         Width           =   2835
      End
      Begin MSMask.MaskEdBox txtHoraFin 
         Height          =   315
         Left            =   4290
         TabIndex        =   3
         Top             =   255
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHoraInicio 
         Height          =   315
         Left            =   1395
         TabIndex        =   2
         Top             =   270
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmbIdTipoServicio 
         Height          =   330
         Left            =   1395
         TabIndex        =   4
         Top             =   660
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSuSalud 
         AutoSize        =   -1  'True
         Caption         =   "(SuSalud)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   4260
         TabIndex        =   20
         Top             =   1095
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblTipoActividad 
         Caption         =   "Tipo actividad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1065
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblIdTipoServicio 
         Caption         =   "Tipo servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label lblHoraFin 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora fin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3510
         TabIndex        =   14
         Top             =   300
         Width           =   735
      End
      Begin VB.Label lblHoraInicio 
         Caption         =   "Hora inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   11
      Top             =   2730
      Width           =   5265
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "TurnoDetalle.frx":008E
         DownPicture     =   "TurnoDetalle.frx":0552
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
         Left            =   2685
         Picture         =   "TurnoDetalle.frx":0A3E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "TurnoDetalle.frx":0F2A
         DownPicture     =   "TurnoDetalle.frx":138A
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
         Left            =   1140
         Picture         =   "TurnoDetalle.frx":17FF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   5295
      Begin VB.TextBox txtDescripcion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   3675
      End
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   255
         TabIndex        =   10
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label lblCodigo 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   255
         TabIndex        =   9
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "TurnoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Turnos
'        Programado por: Barrantes D
'        Fecha: Febrero 2007
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdTurno As Long
Dim mo_Turno As New doTurno
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let IdTurno(lValue As Long)
   ml_IdTurno = lValue
End Property
Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property

Private Sub cmbIdTipoServicio_Change()
Dim rsEspecialidad As New Recordset

    cmbIdEspecialidad.BoundColumn = "IdEspecialidad"
    cmbIdEspecialidad.ListField = "Nombre"
    Set rsEspecialidad = mo_AdminServiciosHosp.EspecialidadSeleccionarPorTipoServiciosql2000(Val(cmbIdTipoServicio.BoundText))
    Set cmbIdEspecialidad.RowSource = rsEspecialidad
    
    If rsEspecialidad.RecordCount = 1 Then
            rsEspecialidad.MoveFirst
            cmbIdEspecialidad.BoundText = rsEspecialidad!IdEspecialidad
            cmbIdEspecialidad.Enabled = False
    End If
    
    If mo_AdminServiciosHosp.MensajeError <> "" Then
        MsgBox mo_AdminServiciosHosp.MensajeError, vbInformation, Me.Caption
    End If
    HabilitaSuSalud
End Sub

Sub HabilitaSuSalud()
    If Val(cmbIdTipoServicio.BoundText) <> 99 Then
       lblTipoActividad.Visible = False
       cmbTipoActividad.Visible = False
       lblSuSalud.Visible = False
       cmbTipoActividad.Text = ""
    Else
       lblTipoActividad.Visible = True
       cmbTipoActividad.Visible = True
       lblSuSalud.Visible = True
    End If
End Sub

Private Sub cmbTipoActividad_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbTipoActividad
    AdministrarKeyPreview KeyCode

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCodigo
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCodigo_LostFocus()
   mo_Formulario.MarcarComoVacio txtCodigo
End Sub

Private Sub cmbIdTipoServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdTipoServicio
    AdministrarKeyPreview KeyCode
End Sub


Private Sub cmbIdTipoServicio_LostFocus()
   If cmbIdTipoServicio.Text <> "" Then
       cmbIdTipoServicio.BoundText = Val(Split(cmbIdTipoServicio.Text, " = ")(0))
   End If
   mo_Formulario.MarcarComoVacio cmbIdTipoServicio
End Sub

Private Sub cmbIdTipoServicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraFin_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtHoraFin
   AdministrarKeyPreview KeyCode
End Sub


'A.Yañez 06-11-2014******************************************
Private Sub txtHoraFin_LostFocus()
        If Not SIGHEntidades.EsHora(txtHoraFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            txtHoraFin.Text = SIGHEntidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraFin_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtHoraInicio_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtHoraInicio
    AdministrarKeyPreview KeyCode
End Sub
'A.Yañez 06-11-2014******************************************
Private Sub txtHoraInicio_LostFocus()
        If Not SIGHEntidades.EsHora(txtHoraInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
             txtHoraInicio.Text = SIGHEntidades.HORA_VACIA_HM
        End If
End Sub

Private Sub txtHoraInicio_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsParaHora(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
    AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_LostFocus()
   mo_Formulario.MarcarComoVacio txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetraONumero(KeyAscii) And Not (InStr(1, "!·$%&/()=?¿*-_:,", Chr(KeyAscii)) > 0) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
 
    Select Case mi_Opcion
         Case sghAgregar
         Case sghModificar
         Case sghConsultar
            Me.btnAceptar.Enabled = False
         Case sghEliminar
     End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Turnos"
       Case sghModificar
           Me.Caption = "Modificar Turnos"
       Case sghConsultar
           Me.Caption = "Consultar Turnos"
       Case sghEliminar
           Me.Caption = "Eliminar Turnos"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
            btnAceptar_Click
        End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                    MsgBox "Los datos se eliminaron correctamente", vbInformation, Me.Caption
                    Me.Visible = False
                Else
                    MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   ValidarDatosObligatorios = False
   Dim lcErrores As String
   lcErrores = ""
   If Me.txtCodigo.Text = "" Then
       lcErrores = lcErrores & "Ingrese el código" & Chr(13)
   End If
   If Me.txtDescripcion.Text = "" Then
       lcErrores = lcErrores & "Ingrese la descripción" & Chr(13)
   End If
   If cmbIdTipoServicio.Text = "" Then
       lcErrores = lcErrores & "Elija el Tipo de Servicio" & Chr(13)
   End If
   If lcErrores <> "" Then
      MsgBox lcErrores, vbInformation, "Turnos "
   Else
      ValidarDatosObligatorios = True
   End If

End Function

Function ValidarReglas() As Boolean
   ValidarReglas = False
   Dim lcErrores As String
   lcErrores = ""
   If Me.txtHoraInicio.Text >= Me.txtHoraFin.Text And Val(cmbIdTipoServicio.BoundText) <> sghHospitalizacion And _
                                                  Val(cmbIdTipoServicio.BoundText) <> sghEmergenciaConsultorios Then
      lcErrores = lcErrores & "La Hora de Inicio no puede ser mayor o igual a la hora final" & Chr(13)
   End If
   If lcErrores <> "" Then
      MsgBox lcErrores, vbInformation, "Turnos "
   Else
      ValidarReglas = True
   End If
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()


   With mo_Turno
           .IdTurno = Me.IdTurno
           .Codigo = Me.txtCodigo.Text
           .idTipoServicio = Val(Me.cmbIdTipoServicio.BoundText)
           .HoraFin = Me.txtHoraFin.Text
           .HoraInicio = Me.txtHoraInicio.Text
           .Descripcion = Me.txtDescripcion.Text
           .IdUsuarioAuditoria = Me.idUsuario
           .IdEspecialidad = Val(Me.cmbIdEspecialidad.BoundText)
           If Val(Me.cmbIdTipoServicio.BoundText) = 99 Then
              .idTipoActividades = Me.cmbTipoActividad.ListIndex
           Else
              .idTipoActividades = 0
           End If
   End With
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
    
    CargaDatosAlObjetosDeDatos
    AgregarDatos = mo_AdminProgramacionMedica.TurnosAgregar(mo_Turno, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)
    
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean
    CargaDatosAlObjetosDeDatos
    ModificarDatos = mo_AdminProgramacionMedica.TurnosModificar(mo_Turno, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)
End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    
    CargaDatosAlObjetosDeDatos
    EliminarDatos = mo_AdminProgramacionMedica.TurnosEliminar(mo_Turno, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, txtDescripcion.Text)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Turnos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

       Set mo_Turno = mo_AdminProgramacionMedica.TurnosSeleccionarPorId(Me.IdTurno)
       
       If mo_AdminProgramacionMedica.MensajeError <> "" Then
            MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminProgramacionMedica.MensajeError, vbInformation, Me.Caption
           mb_ExistenDatos = False
           Exit Sub
       End If
       
       If Not mo_Turno Is Nothing Then
            With mo_Turno
                Me.IdTurno = .IdTurno
                Me.txtCodigo.Text = .Codigo
                Me.cmbIdTipoServicio.BoundText = .idTipoServicio
                Me.txtHoraFin.Text = .HoraFin
                Me.txtHoraInicio.Text = .HoraInicio
                Me.txtDescripcion.Text = .Descripcion
                Me.cmbIdEspecialidad.BoundText = .IdEspecialidad
                If .idTipoServicio = 99 Then
                   cmbTipoActividad.ListIndex = .idTipoActividades
                End If
                mb_ExistenDatos = True
            End With
       Else
           mb_ExistenDatos = False
           Exit Sub
       End If
       HabilitaSuSalud
   
End Sub

Sub CargarComboBoxes()
    
    cmbIdTipoServicio.BoundColumn = "IdTipoServicio"
    cmbIdTipoServicio.ListField = "DescripcionLarga"
    Set cmbIdTipoServicio.RowSource = mo_AdminServiciosHosp.TiposServicioSeleccionarAsistenciales

    Dim oRsTmp1 As New Recordset
    Set oRsTmp1 = mo_ReglasComunes.TiposActividadesSeleccionarTodos
    If oRsTmp1.RecordCount > 0 Then
       Me.cmbTipoActividad.Clear
       Me.cmbTipoActividad.AddItem "<<ninguno>>"
       oRsTmp1.MoveFirst
       Do While Not oRsTmp1.EOF
          Me.cmbTipoActividad.AddItem oRsTmp1!actividad
          oRsTmp1.MoveNext
       Loop
    End If
    oRsTmp1.Close
    Set oRsTmp1 = Nothing
End Sub

Sub LimpiarFormulario()
    Me.IdTurno = 0
    Me.txtCodigo.Text = ""
    Me.cmbIdTipoServicio.BoundText = ""
    Me.txtHoraFin.Text = SIGHEntidades.HORA_VACIA_HM
    Me.txtHoraInicio.Text = SIGHEntidades.HORA_VACIA_HM
    Me.txtDescripcion.Text = ""
End Sub
