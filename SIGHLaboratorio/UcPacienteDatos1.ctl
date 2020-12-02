VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl UcPacienteDatos1 
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   ScaleHeight     =   1680
   ScaleWidth      =   7185
   Begin VB.Frame fraDatosPaciente 
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
      Height          =   1640
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   7155
      Begin VB.TextBox txtFRh 
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
         Height          =   330
         Left            =   6555
         MaxLength       =   20
         TabIndex        =   17
         Top             =   210
         Width           =   450
      End
      Begin VB.TextBox txtGs 
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
         Height          =   330
         Left            =   5580
         MaxLength       =   2
         TabIndex        =   16
         Top             =   210
         Width           =   450
      End
      Begin VB.TextBox txtSexo 
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
         Left            =   5580
         TabIndex        =   15
         Top             =   555
         Width           =   1440
      End
      Begin VB.TextBox txtIdNroHistoria 
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
         Left            =   1110
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtEdad 
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
         Left            =   5580
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1230
         Width           =   465
      End
      Begin VB.TextBox txtPrimerNombre 
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
         Left            =   1110
         MaxLength       =   80
         TabIndex        =   2
         Top             =   1230
         Width           =   3075
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   1110
         MaxLength       =   40
         TabIndex        =   0
         Top             =   570
         Width           =   3075
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   1110
         MaxLength       =   40
         TabIndex        =   1
         Top             =   900
         Width           =   3075
      End
      Begin MSMask.MaskEdBox txtFechaNacimiento 
         Height          =   315
         Left            =   5580
         TabIndex        =   3
         Top             =   900
         Width           =   1440
         _ExtentX        =   2540
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
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "F.Rh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   6150
         TabIndex        =   19
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Gs"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   5325
         TabIndex        =   18
         Top             =   270
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Historia "
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
         TabIndex        =   12
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
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
         Left            =   4400
         TabIndex        =   11
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "A. &Paterno"
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
         TabIndex        =   10
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&Nombre"
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
         TabIndex        =   9
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "A. &Materno"
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
         TabIndex        =   8
         Top             =   930
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "&F. Nacimiento"
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
         Left            =   4400
         TabIndex        =   7
         Top             =   930
         Width           =   1110
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Sexo"
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
         Left            =   4400
         TabIndex        =   6
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label7 
         Caption         =   "Datos del paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   0
         Width           =   1785
      End
   End
End
Attribute VB_Name = "UcPacienteDatos1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para mostrar datos personales del paciente en Resultados
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim ml_idUsuario As Long
Dim mb_ExistenDatos As Boolean
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminServiciosComunes As New SIGHNegocios.ReglasComunes
Dim mo_AdminServiciosGeograficos As New SIGHNegocios.ReglasServGeograf
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_AdminProgramacion As New SIGHNegocios.ReglasDeProgMedica
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
Dim mo_AdminHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mrs_Diagnosticos As New ADODB.Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim ml_TipoVistaForm As sghTipoVistaFormAtenciones
Dim mb_PacienteNoIdentificado As Boolean
Public Event SeModificoFechaNacimiento(sFechaNacimiento As String)
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Public Event SeModificoSexo(lIdTipoSexo As Long)

Dim mo_cmbIdTipoGenHistoriaClinica As New sighentidades.ListaDespleglable
Dim mo_CmbIdTipoSexo As New sighentidades.ListaDespleglable

'------------------------------------------------------------------------------------
'                               VARIABLE PARA LA FILIACION
'------------------------------------------------------------------------------------
Dim ml_idPaciente As Long
Dim ms_Autogenerado As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_FechaRegistro As Date
Dim ml_FechaNacimiento As Date
Dim ml_idTipoSexo As Long
Property Let idTipoSexo(lValue As Long)
    If lValue = 1 Then
      txtSexo.Text = "Masculino"
    Else
      txtSexo.Text = "Femenino"
    End If
End Property

Property Let FechaRegistro(lValue As Date)
  ml_FechaRegistro = lValue
End Property

Property Let idPaciente(lValue As Long)
  ml_idPaciente = lValue
End Property

Property Get idPaciente() As Long
  idPaciente = ml_idPaciente
End Property

Property Get Edad() As Long
  Edad = txtEdad.Text
End Property

Property Let Autogenerado(sValue As String)
  ms_Autogenerado = sValue
End Property

Property Get Autogenerado() As String
  Autogenerado = ms_Autogenerado
End Property

Property Let FechaNacimiento(sValue As String)
  txtFechaNacimiento.Text = sValue
End Property

Property Get FechaNacimiento() As String
  FechaNacimiento = txtFechaNacimiento.Text
End Property

Property Let NroHistoriaClinica(lValue As Long)
  txtIdNroHistoria.Text = CStr(lValue)
End Property

Property Get NroHistoriaClinica() As Long
  NroHistoriaClinica = Val(txtIdNroHistoria.Text)
End Property

Property Get ExistePaciente() As Boolean
  ExistePaciente = mb_ExistenDatos
End Property

Property Get APat() As String
  APat = txtApellidoPaterno.Text
End Property

Property Get AMat() As String
  AMat = txtApellidoMaterno.Text
End Property

Property Get Nombre() As String
  Nombre = txtPrimerNombre.Text
End Property

Property Get Sexo() As String
  Sexo = txtSexo.Text
End Property

Private Sub txtFechaNacimiento_Change()
  RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text)
End Sub

Private Sub txtFechaNacimiento_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtFechaNacimiento
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtFechaNacimiento_LostFocus()
  If txtFechaNacimiento <> sighentidades.FECHA_VACIA_DMY Then
    If Not EsFecha(txtFechaNacimiento, "DD/MM/AAAA") Then
      MsgBox "La fecha ingresada no es válida", vbInformation, "Datos de paciente"
      txtFechaNacimiento = sighentidades.FECHA_VACIA_DMY
    Else
      'txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), Now)))
      txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), IIf(ml_FechaRegistro = 0, Date, ml_FechaRegistro))))  'Actualizado 20102014 yamill palomino
    End If
  End If
  mo_Formulario.MarcarComoVacio txtFechaNacimiento
End Sub

Private Sub txtFechaNacimiento_KeyPress(KeyAscii As Integer)
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsParaFecha(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub txtPrimerNombre_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtPrimerNombre
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtPrimerNombre_LostFocus()
  If txtPrimerNombre.Text <> "NN" Then txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
  mo_Formulario.MarcarComoVacio txtPrimerNombre
End Sub

Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtApellidoMaterno_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtApellidoMaterno
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtApellidoMaterno_LostFocus()
  If txtApellidoMaterno.Text <> "NN" Then txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
  mo_Formulario.MarcarComoVacio txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
  mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
  RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub

Private Sub txtApellidoPaterno_LostFocus()
  If txtApellidoPaterno.Text <> "NN" Then txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
  mo_Formulario.MarcarComoVacio txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
  If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
    If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then KeyAscii = 0
  End If
End Sub

Public Sub ConfigurarComboBoxes()
  Dim sMensaje As String
  'CARGA COMBO BOXES DE PACIENTE
  mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
  mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
  Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
  sMensaje = sMensaje + mo_AdminServiciosComunes.MensajeError
End Sub

Public Function ValidarDatosObligatorios() As String
  Dim sMensajeLocal As String
  '---------------------------------------------------------------------------------
  '           VALIDA DATOS DE PACIENTES
  '---------------------------------------------------------------------------------
  If txtApellidoPaterno.Text = "" Then sMensajeLocal = sMensajeLocal + "Ingrese el apellido paterno" + Chr(13)
  If txtApellidoMaterno.Text = "" Then sMensajeLocal = sMensajeLocal + "Ingrese el apellido materno" + Chr(13)
  If txtPrimerNombre.Text = "" Then sMensajeLocal = sMensajeLocal + "Ingrese el primer nombre" + Chr(13)
  If Val(mo_CmbIdTipoSexo.BoundText) = 0 Then sMensajeLocal = sMensajeLocal + "Ingrese el sexo" + Chr(13)
  If txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY Then sMensajeLocal = sMensajeLocal + "Ingrese la Fecha de Nacimiento" + Chr(13)
  ValidarDatosObligatorios = sMensajeLocal
End Function

Public Function ValidarReglas(oDOPaciente As doPaciente) As Boolean
  Dim rspacientes As ADODB.Recordset
  ValidarReglas = False
    
    'Si el paciente aun no existe (IdPaciente = 0) se verifica que no haya duplicados
    If oDOPaciente.idPaciente = 0 Then
        Set rspacientes = mo_AdminAdmision.PacientesObtenerConElMismoNombre(oDOPaciente)
        If Not (rspacientes.EOF And rspacientes.BOF) Then
            rspacientes.MoveFirst
            MsgBox "Existe un paciente con el mismo nombre: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre, vbExclamation, "Datos de paciente"
            rspacientes.Close
            Exit Function
        End If
        rspacientes.Close
         
         Set rspacientes = mo_AdminAdmision.PacientesObtenerConElAutogenerado(oDOPaciente)
         If Not (rspacientes.EOF And rspacientes.BOF) Then
             rspacientes.MoveFirst
             If MsgBox("Existe un paciente con el mismo número autogenerado: " + rspacientes!ApellidoPaterno + " " + rspacientes!ApellidoMaterno + " " + rspacientes!PrimerNombre + Chr(13) + "Desea continuar?", vbQuestion + vbYesNo, "Datos de paciente") = vbNo Then
                 rspacientes.Close
                 Exit Function
             End If
         End If
         rspacientes.Close
        
    End If
    
   If txtFechaNacimiento.Text <> sighentidades.FECHA_VACIA_DMY Then
        If CDate(txtFechaNacimiento.Text) > Date Then
            MsgBox "La fecha de nacimiento no puede ser mayor que la fecha de creación de la historia", vbExclamation, "Registro de pacientes"
            Exit Function
        End If
    End If
   
   ValidarReglas = True

End Function

Public Function CargarDatosAlObjetoDatos(oDOPaciente As doPaciente)
    
    '---------------------------------------------------------------------------------
    '           CARGA DATOS DEL PACIENTE
    '---------------------------------------------------------------------------------
   With oDOPaciente
        .idPaciente = Me.idPaciente
        .ApellidoPaterno = txtApellidoPaterno.Text
        .ApellidoMaterno = txtApellidoMaterno.Text
        .PrimerNombre = txtPrimerNombre.Text
        .FechaNacimiento = IIf(txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY, 0, txtFechaNacimiento.Text)
        .NroHistoriaClinica = Me.NroHistoriaClinica
        .idTipoSexo = Val(mo_CmbIdTipoSexo.BoundText)
        .IdTipoNumeracion = mo_cmbIdTipoGenHistoriaClinica.BoundText
        .Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(oDOPaciente)
         Autogenerado = .Autogenerado
   End With
   Set CargarDatosAlObjetoDatos = oDOPaciente
End Function

Public Sub CargarDatosDePacienteALosControles()
  Dim oPacientes  As New doPaciente
  Dim oConexion As New Connection
  oConexion.Open sighentidades.CadenaConexion
  oConexion.CursorLocation = adUseClient
  'CARGAR DATOS DEL PACIENTE
  Set oPacientes = mo_AdminAdmision.PacientesSeleccionarPorId(ml_idPaciente, oConexion)
  'PacientesSeleccionarPorId
  If mo_AdminAdmision.MensajeError <> "" Then
    MsgBox "No se pudo obtener los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbInformation, "Datos de paciente"
    mb_ExistenDatos = False
    Exit Sub
  End If
  If Not oPacientes Is Nothing Then
    With oPacientes
                txtGs.Text = .grupoSanguineo
                txtFRh.Text = .factorRh
                mo_Formulario.HabilitarDeshabilitar txtGs, False
                mo_Formulario.HabilitarDeshabilitar txtFRh, False
    
                Me.idPaciente = .idPaciente
                Autogenerado = .Autogenerado
                txtApellidoPaterno.Text = Trim(.ApellidoPaterno)
                txtApellidoMaterno.Text = Trim(.ApellidoMaterno)
                txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(Trim(.PrimerNombre) & " " & Trim(.SegundoNombre) & " " & Trim(.TercerNombre))
                txtFechaNacimiento.Text = IIf(.FechaNacimiento = 0, sighentidades.FECHA_VACIA_DMY, _
                                            Format(.FechaNacimiento, sighentidades.DevuelveFechaSoloFormato_DMY)) 'DBB 19 Marzo
                RaiseEvent SeModificoFechaNacimiento(txtFechaNacimiento.Text)
                'mo_CmbIdTipoSexo.BoundText = .IdTipoSexo
                'RaiseEvent SeModificoSexo(.IdTipoSexo)
                If .idTipoSexo = 1 Then
                  txtSexo.Text = "Masculino"
                Else
                  txtSexo.Text = "Femenino"
                End If
                'mo_cmbIdTipoGenHistoriaClinica.BoundText = .IdTipoNumeracion
                'cmbIdTipoGenHistoriaClinica.Tag = .IdTipoNumeracion         'lo guarda para luego comparar
                txtIdNroHistoria.Text = .NroHistoriaClinica          'esto tiene que ir luego del tipo de generacion, por que sino se borra con el change del combo box
                txtIdNroHistoria.Tag = .NroHistoriaClinica
                txtEdad.Text = Trim(Str(EdadActual(.FechaNacimiento, ml_FechaRegistro)))
                mb_ExistenDatos = True
    End With
  Else
    mb_ExistenDatos = False
    Exit Sub
  End If
  oConexion.Close
  Set oConexion = Nothing
End Sub

Public Sub ConfigurarValoresPorDefecto()
  mo_Formulario.HabilitarDeshabilitar txtIdNroHistoria, False
  mo_Formulario.HabilitarDeshabilitar txtEdad, False
  
  mo_cmbIdTipoGenHistoriaClinica.BoundColumn = "IdTipoNumeracion"
  mo_cmbIdTipoGenHistoriaClinica.ListField = "DescripcionLarga"
  Set mo_cmbIdTipoGenHistoriaClinica.RowSource = mo_AdminArchivoClinico.TiposGeneracionHistoriasSeleccionarDeConsultaExterna()
  mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
  mo_CmbIdTipoSexo.BoundColumn = "IdtipoSexo"
  mo_CmbIdTipoSexo.ListField = "DescripcionLarga"
  Set mo_CmbIdTipoSexo.RowSource = mo_AdminServiciosComunes.TiposSexoSeleccionarTodos()
End Sub

Public Sub LimpiarDatosDePaciente()
           
           'LIMPIAR DATOS DEL PACIENTE
           idPaciente = 0
           Autogenerado = 0
           txtApellidoPaterno.Text = ""
           txtApellidoMaterno.Text = ""
           txtPrimerNombre.Text = ""
           txtFechaNacimiento.Text = sighentidades.FECHA_VACIA_DMY
           txtEdad.Text = ""
           mo_cmbIdTipoGenHistoriaClinica.BoundText = sghHistoriaTemporalCOnsultaExterna
           txtIdNroHistoria.Text = ""
           txtGs.Text = ""
           txtFRh.Text = ""

End Sub

Public Sub DeshabilitarFrames(lbDesHabilita As Boolean)
    
    'fraDatosPaciente.Enabled = Not lbDesHabilita
    If lbDesHabilita = True Then
        mo_Formulario.HabilitarDeshabilitar fraDatosPaciente, False
    Else
        mo_Formulario.HabilitarDeshabilitar fraDatosPaciente, True
    End If
End Sub

Public Sub SetFocusOnApellidoPaterno()
         txtApellidoPaterno.SetFocus
End Sub


Public Function Inicializar()
    ConfigurarValoresPorDefecto
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
    Case vbKeyF2
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
     Case vbKeyF8
    Case vbKey9
    End Select
       
End Sub

Public Sub CargaAlgunosDatosDesdeBoleta(lcRazonSocial As String)
    Dim lnLen As Integer
    Dim lnPos As Integer
    lnPos = 1
    txtApellidoPaterno.Text = ""
    txtApellidoMaterno.Text = ""
    txtPrimerNombre.Text = ""
    If lcRazonSocial <> "" Then
        For lnLen = 1 To Len(lcRazonSocial)
            If Mid(lcRazonSocial, lnLen, 1) = " " Then
               lnPos = lnPos + 1
            Else
               Select Case lnPos
               Case 1
                    txtApellidoPaterno.Text = txtApellidoPaterno.Text & Mid(lcRazonSocial, lnLen, 1)
               Case 2
                    txtApellidoMaterno.Text = txtApellidoMaterno.Text & Mid(lcRazonSocial, lnLen, 1)
               Case 3
                    txtPrimerNombre.Text = txtPrimerNombre.Text & Mid(lcRazonSocial, lnLen, 1)
               End Select
            End If
        Next
        txtApellidoPaterno.Text = Left(txtApellidoPaterno.Text, 20)
        txtApellidoMaterno.Text = Left(txtApellidoMaterno.Text, 20)
        txtPrimerNombre.Text = Left(txtPrimerNombre.Text, 20)
        txtApellidoPaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoPaterno.Text)
        txtApellidoMaterno.Text = mo_Teclado.CapitalizarNombres(txtApellidoMaterno.Text)
        txtPrimerNombre.Text = mo_Teclado.CapitalizarNombres(txtPrimerNombre.Text)
        If txtFechaNacimiento.Text <> sighentidades.FECHA_VACIA_DMY Then
           txtEdad.Text = Trim(Str(EdadActual(txtFechaNacimiento, ml_FechaRegistro)))
           'Actualizado30102014 Yamill Palomino
           txtEdad.Text = Trim(Str(EdadActual(CDate(txtFechaNacimiento.Text), IIf(ml_FechaRegistro = 0, Date, ml_FechaRegistro))))
        End If
    End If
End Sub

Public Function DevuelveHistoriaApellidosYnombre() As String
    
    DevuelveHistoriaApellidosYnombre = "(" & HCigualDNI_DevuelveHistoriaConCerosIzquierda(txtIdNroHistoria.Text, False) & ") " & Trim(txtApellidoPaterno.Text) & " " & Trim(txtApellidoMaterno.Text) & " " & txtPrimerNombre.Text
End Function

Public Function DevuelveSexo() As String
    DevuelveSexo = txtSexo.Text
End Function
Public Function DevuelveFechaNacimiento() As String
    DevuelveFechaNacimiento = txtFechaNacimiento.Text
End Function


