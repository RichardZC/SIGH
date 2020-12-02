VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PacienteDetalle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12045
   Icon            =   "PacienteDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabPaciente 
      Height          =   6975
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "1.1) Datos del Paciente"
      TabPicture(0)   =   "PacienteDetalle.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ucPacientesDetalle1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "1.2) SUNASA"
      TabPicture(1)   =   "PacienteDetalle.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdActualizaNroAutomatico"
      Tab(1).Control(1)=   "UcPacientesSunasa1"
      Tab(1).ControlCount=   2
      Begin SISGalenPlus.ucPacientesDetalle ucPacientesDetalle1 
         Height          =   6465
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   11404
      End
      Begin VB.CommandButton cmdActualizaNroAutomatico 
         Caption         =   "..."
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         ToolTipText     =   "Actualiza último N° Automático de HC"
         Top             =   6540
         Width           =   255
      End
      Begin SISGalenPlus.UcPacientesSunasa UcPacientesSunasa1 
         Height          =   5925
         Left            =   -74880
         TabIndex        =   5
         Top             =   450
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   10451
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   0
      Top             =   6990
      Width           =   11835
      Begin VB.CommandButton btnImprimeFiliacion 
         Caption         =   "Filiación Arch.Clínico"
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
         Left            =   180
         Picture         =   "PacienteDetalle.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PacienteDetalle.frx":11DB
         DownPicture     =   "PacienteDetalle.frx":169F
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
         Left            =   6052
         Picture         =   "PacienteDetalle.frx":1B8B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PacienteDetalle.frx":2077
         DownPicture     =   "PacienteDetalle.frx":24D7
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
         Left            =   4507
         Picture         =   "PacienteDetalle.frx":294C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Image pi_imagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   0
      MouseIcon       =   "PacienteDetalle.frx":2DC1
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      ToolTipText     =   "Pulsar Doble Click para ampliar Imagen"
      Top             =   0
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "PacienteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Mantenimiento de Pacientes
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idUsuario As Long
Dim ml_IdPaciente As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_Pacientes  As New doPaciente
Dim ml_TipoServicio As sghTipoServicio
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_Historia As New DOHistoriaClinica
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim mb_AlPulsarClicEnACEPTARdebeSalirDeVentana As Boolean
Dim oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
Dim ldFechaActualServidor As Date
Dim lbCargaUnaVezVEntana As Boolean
Dim mo_DoPacientesDatosAdd As New DoPacienteDatosAdd

Property Let AlPulsarClicEnACEPTARdebeSalirDeVentana(bValue As Boolean)
   mb_AlPulsarClicEnACEPTARdebeSalirDeVentana = bValue
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_IdPaciente
End Property
Property Let TipoServicio(sValue As sghTipoServicio)
   ml_TipoServicio = sValue
End Property
Property Get TipoServicio() As sghTipoServicio
   TipoServicio = ml_TipoServicio
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()
    ucPacientesDetalle1.MarcoCheckPacienteNuevo = False
    '
    Me.ucPacientesDetalle1.HacerVisibleCheckPacienteNoIdentificado Not (ml_TipoServicio = sghConsultaExterna Or ml_TipoServicio = sghHospitalizacion)
    Me.ucPacientesDetalle1.NotaSobreUbicacion = "(*) Última actualización de los datos de ubicación"
    Select Case mi_Opcion
        Case sghAgregar
               ValoresPorDefecto
               Me.ucPacientesDetalle1.TabEnNroHistoria
               ucPacientesDetalle1.MarcoCheckPacienteNuevo = True
        Case sghModificar
            CargarDatosAlosControles
        Case sghConsultar
            CargarDatosAlosControles
        Case sghEliminar
            CargarDatosAlosControles
    End Select
 
     Select Case mi_Opcion
     Case sghAgregar
        
     Case sghModificar
        
     Case sghConsultar
        DeshabilitarControlesParaEdicion
        Me.btnAceptar.Enabled = False
        
    Case sghEliminar
        DeshabilitarControlesParaEdicion
        
    End Select
End Sub
Sub DeshabilitarControlesParaEdicion()

    Me.ucPacientesDetalle1.DeshabilitarFrames

End Sub

Sub ValoresPorDefecto()
    Me.ucPacientesDetalle1.TipoServicio = ml_TipoServicio
    Me.ucPacientesDetalle1.ConfigurarValoresPorDefecto
End Sub
'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Private Sub btnImprimeFiliacion_Click()
    ImprimeFiliacion True
    Me.Visible = False
End Sub
Sub ImprimeFiliacion(lbCargaDatos As Boolean)
    Dim oImprime As New RptHistoriaClinicaCE
    Dim oEdad As Edad
    If lbCargaDatos = True Then
       CargaDatosAlObjetosDeDatos
    End If
    oEdad = sighEntidades.CalcularEdad(mo_Pacientes.FechaNacimiento, mo_Historia.fechacreacion)
    oImprime.ImprimeEnFormatoDeFiliacionParaHistoriaClinica ml_IdPaciente, oEdad.Edad, oEdad.TipoEdad, Me.hwnd
    Set oImprime = Nothing
End Sub
Private Sub cmdActualizaNroAutomatico_Click()
    mo_ReglasArchivoClinico.ActualizaDatosConProblemas False
    Me.Visible = False
End Sub
Sub InicializarParametros()
    wxParametro211 = lcBuscaParametro.SeleccionaFilaParametro(211)
    wxParametro237 = lcBuscaParametro.SeleccionaFilaParametro(237)
    wxParametro242 = lcBuscaParametro.SeleccionaFilaParametro(242)
    wxParametro282 = lcBuscaParametro.SeleccionaFilaParametro(282)
    wxParametro287 = lcBuscaParametro.SeleccionaFilaParametro(287)
    wxParametro333 = lcBuscaParametro.SeleccionaFilaParametro(333)
    wxParametro351 = lcBuscaParametro.SeleccionaFilaParametro(351)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
    
End Sub

Private Sub Form_Deactivate()
    lbCargaUnaVezVEntana = True
End Sub
Sub Form_Load()
        InicializarParametros
        '

        mb_AlPulsarClicEnACEPTARdebeSalirDeVentana = False
        ucPacientesDetalle1.meHwnd = Me.hwnd
        ucPacientesDetalle1.Opcion = mi_Opcion
        ucPacientesDetalle1.Inicializar
        '
        UcPacientesSunasa1.idTipoFinanciamiento = sghTrabajaSeguroSIS
        UcPacientesSunasa1.Opcion = mi_Opcion
        UcPacientesSunasa1.idPaciente = ml_IdPaciente
        UcPacientesSunasa1.Inicializar
        UcPacientesSunasa1.YaNoTieneSeguro
        '
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Pacientes"
       Case sghModificar
           Me.Caption = "Modificar Pacientes"
       Case sghConsultar
           Me.Caption = "Consultar Pacientes"
       Case sghEliminar
           Me.Caption = "Eliminar Pacientes"
       End Select

       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
       '
       lbCargaUnaVezVEntana = True
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   ElseIf lbCargaUnaVezVEntana = True Then
        lbCargaUnaVezVEntana = False
        On Error Resume Next
        Select Case WxDEFAULT_BUSQ_PACIENTE
        Case sghDefaultVentana.sighApellidoPaterno
             Me.ucPacientesDetalle1.SetFocusOnApellidoPaterno
        Case sghDefaultVentana.sighDNI
             Me.ucPacientesDetalle1.SetFocusEnDNI
        Case sghDefaultVentana.sighHistoria
             Me.ucPacientesDetalle1.SetFocusEnHistoria
        End Select
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    'Case vbKeyEscape
    '    btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
     Case vbKeyF7
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 0
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoDomicilio
     Case vbKeyF8
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 1
         On Error Resume Next
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoProcedencia
     Case vbKeyF9
         On Error Resume Next
         Me.ucPacientesDetalle1.SetPestaniaTabPaciente 2
         Me.ucPacientesDetalle1.SetFocusOnDepartamentoNacimiento
    End Select
       
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   'AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   If btnAceptar.Enabled = False Then
      Exit Sub
   End If
  Select Case mi_Opcion
   Case sghAgregar
       GrabaParametro8 "antes validar"
       If ValidarDatosObligatorios() Then
            GrabaParametro8 "antes objeto"
            CargaDatosAlObjetosDeDatos
            GrabaParametro8 "antes validarreglas"
            If ValidarReglas() Then
                
                If AgregarDatos() Then
                     GrabaParametro8 "paso AGREGARDATOS"
                     If mb_AlPulsarClicEnACEPTARdebeSalirDeVentana = True Then
                        ml_IdPaciente = mo_Pacientes.idPaciente
                        Me.Visible = False
                     Else
                        GrabaParametro8 "antes de msgbox"
                        If mo_Pacientes.FichaFamiliar = "" Then
                           MsgBox "Los datos se agregaron correctamente" & Chr(13) & "N° Historia Clínica: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_Pacientes.NroHistoriaClinica)), False), vbInformation, Me.Caption
                        Else
                           MsgBox "Los datos se agregaron correctamente" & Chr(13) & "FichaFamiliar: " & mo_Pacientes.FichaFamiliar, vbInformation, Me.Caption
                        End If
                        '
                        Dim oConexion As New Connection
                        oConexion.CommandTimeout = 300
                        oConexion.CursorLocation = adUseClient
                        oConexion.Open sighEntidades.CadenaConexion
                        mo_ReglasArchivoClinico.generadorNroHistoriaClinicaActualizaNroAutomaticoDeHistoriaClinica oConexion
                        oConexion.Close
                        Set oConexion = Nothing
                        '
                        
                        If MsgBox("¿Imprime Datos del paciente (Filiación)  ?", vbYesNo, Me.Caption) = vbYes Then
                           ml_IdPaciente = mo_Pacientes.idPaciente
                           ImprimeFiliacion False
                        End If
                        ucPacientesDetalle1.LimpiarDatosDePaciente wxParametro211, ldFechaActualServidor
                        UcPacientesSunasa1.LimpiarDatos
                        UcPacientesSunasa1.YaNoTieneSeguro
                        TabPaciente.Tab = 0
                        ucPacientesDetalle1.SetFocusEnDNI
                     End If
                 Else
                     MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
                End If
            End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
            CargaDatosAlObjetosDeDatos
            If ValidarReglas() Then
           
                If ModificarDatos() Then
                    'AGREGAR CODIGO PARA MODIFICAR datos en FUA y pasar parametro mo_Paciente
                    
                    If mo_Pacientes.FichaFamiliar = "" Then
                       MsgBox "Los datos se modificaron correctamente" & Chr(13) & "N° Historia Clínica: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_Pacientes.NroHistoriaClinica)), False), vbInformation, Me.Caption
                    Else
                       MsgBox "Los datos se modificaron correctamente" & Chr(13) & "Ficha Familiar: " & mo_Pacientes.FichaFamiliar, vbInformation, Me.Caption
                    End If
                    Me.Visible = False
                 Else
                     MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
                End If
            End If
       End If
   Case sghEliminar
            CargaDatosAlObjetosDeDatos
            
            If mo_AdminFacturacion.PacienteSePuedeEliminar(Me.idPaciente) Then
                If MsgBox("¿Esta seguro de eliminar el paciente?", vbYesNo, Me.Caption) = vbYes Then
                    If EliminarDatos() Then
                         If mo_Pacientes.FichaFamiliar = "" Then
                            MsgBox "Los datos se eliminaron correctamente" & Chr(13) & "N° Historia Clínica: " & HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(mo_Pacientes.NroHistoriaClinica)), False), vbInformation, Me.Caption
                         Else
                            MsgBox "Los datos se eliminaron correctamente" & Chr(13) & "Ficha Familiar: " & mo_Pacientes.FichaFamiliar, vbInformation, Me.Caption
                         End If
                         Me.Visible = False
                     Else
                         MsgBox "No se pudo eliminar los datos, debe tener Movimientos en CE/Hosp/Emerg/Boleta" + Chr(13) + mo_AdminAdmision.MensajeError, vbExclamation, Me.Caption
                    End If
                End If
            Else
                MsgBox "El paciente no se puede eliminar porque tiene Atenciones registradas", vbInformation, Me.Caption
            End If
   End Select
End Sub
Sub GrabaParametro8(lcValorTExto As String)
    Dim oConexion1 As New Connection
    Dim oRsTmp1 As New Recordset
    Dim lcSql As String
    oConexion1.CommandTimeout = 900
    oConexion1.CursorLocation = adUseClient
    oConexion1.Open sighEntidades.CadenaConexion
    lcSql = "update parametros set valorTexto='" & lcValorTExto & "' where idparametro=8"
    oRsTmp1.Open lcSql, oConexion1, adOpenKeyset, adLockOptimistic
    oConexion1.Close
    Set oRsTmp1 = Nothing
    Set oConexion1 = Nothing
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------
Function AgregarDatos() As Boolean
    GrabaParametro8 "Antes de Grabar"
    AgregarDatos = mo_AdminAdmision.PacientesAgregarPacienteEHistoriaClinica(mo_Pacientes, mo_Historia, _
                   mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                   Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & " " & mo_Pacientes.SegundoNombre, _
                   oDoSunasaPacientesHistoricos, mo_DoPacientesDatosAdd)
    GrabaParametro8 "paso grabar"
    GrabaImagenesEnRutaDelServidor
    
    'mgaray201411f
'    If AgregarDatos = True Then
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes)
'    End If
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------
Function ModificarDatos() As Boolean
    ModificarDatos = mo_AdminAdmision.PacientesModificarYActualizarHistoriaClinicaDefinitiva(mo_Pacientes, mo_Historia, Me.ucPacientesDetalle1.TipoNumeracionAnterior, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(Str(mo_Pacientes.NroHistoriaClinica)) & " : " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & " " & mo_Pacientes.SegundoNombre, oDoSunasaPacientesHistoricos, mo_DoPacientesDatosAdd)
    GrabaImagenesEnRutaDelServidor
'    If ModificarDatos = True Then
'        'mgaray201411f
'        Dim o_ReglasIntegracion As New ReglasIntegracion
'        Call o_ReglasIntegracion.EnviarDatosPacienteRisPacs(mo_Pacientes, False)
'
'    End If
    
End Function

Sub GrabaImagenesEnRutaDelServidor()
    Dim lcArchivoElegido As String
    Dim lcArchivoImagenFinal As String
    
    lcArchivoElegido = ucPacientesDetalle1.ArchivoElegido
    GrabaParametro8 "lee archivo elegido img"
    lcArchivoImagenFinal = lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(mo_Pacientes.NroHistoriaClinica)) & ".JPG"
    GrabaParametro8 "archivo igm"
    If lcArchivoElegido = "DEL" Then
       Kill lcArchivoImagenFinal
    ElseIf lcArchivoElegido <> "" Then
        GrabaParametro8 "antes de grabar img"
        pi_imagen.Picture = LoadPicture(lcArchivoElegido)
        SavePicture pi_imagen, lcArchivoImagenFinal
        GrabaParametro8 "grabo img"
    End If
End Sub

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean
    EliminarDatos = mo_AdminAdmision.PacientesEliminar(mo_Pacientes, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, Trim(Str(mo_Pacientes.NroHistoriaClinica)) & " : " & Trim(mo_Pacientes.ApellidoPaterno) & " " & Trim(mo_Pacientes.ApellidoMaterno) & " " & Trim(mo_Pacientes.PrimerNombre) & " " & mo_Pacientes.SegundoNombre, oDoSunasaPacientesHistoricos)
    Dim lcArchivoImagenFinal As String
    lcArchivoImagenFinal = lcBuscaParametro.SeleccionaFilaParametro(237) & "\" & Trim(Str(mo_Pacientes.NroHistoriaClinica)) & ".JPG"
    On Error Resume Next
    Kill lcArchivoImagenFinal
End Function

Private Sub btnCancelar_Click()
   lbCargaUnaVezVEntana = True
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
Dim sMensaje As String

    ValidarDatosObligatorios = False
       
    sMensaje = sMensaje + ucPacientesDetalle1.ValidarDatosObligatorios(wxParametro282, wxParametro333)
'    If sMensaje <> "" Then
'        If ucPacientesDetalle1.DevuelveEtnia = "" Then
'           ucPacientesDetalle1.SetFocusEnEtnia
'        ElseIf ucPacientesDetalle1.DevuelveIdioma = "" Then
'           ucPacientesDetalle1.SetFocusEnIdioma
'        End If
'    End If
   
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
    End If
   
    ValidarDatosObligatorios = True

End Function
Function ValidarReglas() As Boolean
Dim rspacientes As ADODB.Recordset

   ValidarReglas = False
   
   If Not ucPacientesDetalle1.ValidarReglas(mo_Pacientes) Then
        Exit Function
   End If
   If mo_Pacientes.FechaNacimiento = 0 Then
        MsgBox "Tiene que registrar Fecha de Nacimiento", vbInformation, Me.Caption
        Exit Function
   End If
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()

    ucPacientesDetalle1.idUsuario = ml_idUsuario
    ucPacientesDetalle1.CargarDatosAlObjetoDatos mo_Pacientes, mo_Historia, mo_DoPacientesDatosAdd
    '
    Me.UcPacientesSunasa1.idUsuario = ml_idUsuario
    Me.UcPacientesSunasa1.CargarDatosAlObjetoDatos oDoSunasaPacientesHistoricos
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Pacientes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlosControles()
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighEntidades.CadenaConexion
        
        
        ucPacientesDetalle1.idPaciente = Me.idPaciente
        ucPacientesDetalle1.CargarDatosDePacienteALosControles oConexion, wxParametro242, wxParametro287
        mb_ExistenDatos = ucPacientesDetalle1.ExistePaciente
        '
        Me.UcPacientesSunasa1.idPaciente = Me.idPaciente
        Me.UcPacientesSunasa1.CargarDatosDelUltimoSeguroDelPacienteALosControles oConexion
        oConexion.Close
        Set oConexion = Nothing
   
End Sub
Sub CargarComboBoxes()
        Me.ucPacientesDetalle1.OpcionQueUsaEsteControl = 1      '1->Pacientes, 2->Admision de Emergencia, 3->Admision de Hospitalizacion
        Me.ucPacientesDetalle1.ConfigurarComboBoxes
        
End Sub
Private Sub TabPaciente_Click(PreviousTab As Integer)
    If TabPaciente.Tab = 1 Then
         On Error Resume Next
         UcPacientesSunasa1.SetFocusEnApellidoCasada
         Me.UcPacientesSunasa1.DatosDeCabecera Me.ucPacientesDetalle1.DevuelvePaciente, Me.ucPacientesDetalle1.DevuelveSexo, Me.ucPacientesDetalle1.DevuelveDocumento, Me.ucPacientesDetalle1.DevuelveNroDocumento, Me.ucPacientesDetalle1.DevuelvePaisDomicilio, Me.ucPacientesDetalle1.DevuelveFechaNacimiento, Me.ucPacientesDetalle1.DevuelveUbigeoDomicilio
    End If

End Sub
Private Sub ucPacientesDetalle1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Private Sub btnImprimir_Click()

End Sub
Private Sub UcPacientesSunasa1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub
