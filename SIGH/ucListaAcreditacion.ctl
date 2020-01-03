VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ucAdmisionLista 
   ClientHeight    =   9840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13365
   ScaleHeight     =   8985
   ScaleMode       =   0  'User
   ScaleWidth      =   12360
   Begin TabDlg.SSTab TabCamas 
      Height          =   9225
      Left            =   60
      TabIndex        =   9
      Top             =   540
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   16272
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "."
      TabPicture(0)   =   "ucListaAcreditacion.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblMedico"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grdAdmision"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraBusqueda"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Admisión por Cama disponible"
      TabPicture(1)   =   "ucListaAcreditacion.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdCamasDisponibles"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnRefrescar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdPacientesSinAltaHE"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdPacientesSinAltaHE 
         Caption         =   "Pacientes sin Alta Médica en Hospitalización y Emergencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73140
         TabIndex        =   23
         Top             =   8670
         Width           =   5550
      End
      Begin VB.CommandButton btnRefrescar 
         Caption         =   "Refrescar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74850
         TabIndex        =   11
         Top             =   8670
         Width           =   1515
      End
      Begin VB.Frame fraBusqueda 
         Caption         =   "Búsqueda"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   13035
         Begin VB.CommandButton cmdSinApellidoPaterno 
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
            Left            =   3975
            Picture         =   "ucListaAcreditacion.ctx":0038
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   450
            Width           =   315
         End
         Begin VB.TextBox TxtNroLineasAmostrar 
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
            Left            =   5925
            TabIndex        =   21
            Text            =   "1000"
            Top             =   840
            Width           =   825
         End
         Begin VB.CheckBox chkActivos 
            Alignment       =   1  'Right Justify
            Caption         =   "Solo ACTIVOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   8670
            TabIndex        =   20
            Top             =   870
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.TextBox txtNemergencia 
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
            Left            =   1335
            TabIndex        =   19
            Top             =   825
            Width           =   1260
         End
         Begin VB.CommandButton cmdPacientesSinAltaMedica 
            Caption         =   "Pacientes sin Alta Médica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   11895
            TabIndex        =   16
            ToolTipText     =   "Listado de Pacientes que no tienen Alta Médica"
            Top             =   165
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtDNI 
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
            Left            =   1065
            MaxLength       =   8
            TabIndex        =   1
            Top             =   465
            Width           =   1000
         End
         Begin VB.CommandButton cmdPacientesSinAltaMedica1 
            Height          =   135
            Left            =   12840
            Picture         =   "ucListaAcreditacion.ctx":05C2
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Lista Pacientes que no tienen Alta Médica"
            Top             =   120
            Visible         =   0   'False
            Width           =   165
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
            Left            =   3075
            MaxLength       =   40
            TabIndex        =   3
            Top             =   465
            Width           =   885
         End
         Begin VB.TextBox txtNroHistoria 
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
            Left            =   2070
            MaxLength       =   9
            TabIndex        =   2
            Top             =   465
            Width           =   1000
         End
         Begin VB.TextBox txtNcuenta 
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
            Left            =   60
            MaxLength       =   9
            TabIndex        =   0
            Top             =   465
            Width           =   1000
         End
         Begin VB.CommandButton btnLimpiar 
            Height          =   315
            Left            =   10470
            Picture         =   "ucListaAcreditacion.ctx":0A04
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   525
            Width           =   1275
         End
         Begin VB.CommandButton btnBuscar 
            Height          =   315
            Left            =   10470
            Picture         =   "ucListaAcreditacion.ctx":35E0
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   150
            Width           =   1305
         End
         Begin VB.ComboBox cmbFecha 
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
            Left            =   4395
            TabIndex        =   4
            Text            =   "cmbFecha"
            Top             =   465
            Width           =   1500
         End
         Begin VB.ComboBox cmbIdResponsable 
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
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   465
            Width           =   4560
         End
         Begin VB.Label lblNroLineasAmostrar 
            AutoSize        =   -1  'True
            Caption         =   "N° líneas máximas a mostrar"
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
            Left            =   3630
            TabIndex        =   22
            Top             =   855
            Width           =   2265
         End
         Begin VB.Label lblNemergencia 
            AutoSize        =   -1  'True
            Caption         =   "Emergencia N°"
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
            Left            =   75
            TabIndex        =   18
            Top             =   855
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "N° Cuenta   N° DNI       Hist.Clínica   Apell. Paterno  F.Ingreso            Servicio          "
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
            Left            =   90
            TabIndex        =   13
            Top             =   240
            Width           =   10365
         End
      End
      Begin UltraGrid.SSUltraGrid grdAdmision 
         Height          =   7185
         Left            =   105
         TabIndex        =   8
         Top             =   1590
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   12674
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de admisiones"
      End
      Begin UltraGrid.SSUltraGrid grdCamasDisponibles 
         Height          =   8160
         Left            =   -74850
         TabIndex        =   10
         Top             =   450
         Width           =   13005
         _ExtentX        =   22939
         _ExtentY        =   14393
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lista de disponibilidad de camas"
      End
      Begin VB.Label lblMedico 
         Caption         =   "* El USUARIO es un PROFESIONAL EN SALUD por lo tanto solo se mostrará sus Pacientes"
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
         Height          =   225
         Left            =   150
         TabIndex        =   17
         Top             =   8865
         Visible         =   0   'False
         Width           =   13020
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Admisión"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "ucAdmisionLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Lista de Pacientes con Admisión en Hosp/Emer/Consultorios Externos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_ReglasHoteleria As New SIGHNegocios.ReglasHoteleria
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_Teclado As New sighEntidades.Teclado
Dim ml_idRegistroSeleccionado As Long
Dim ml_IdAtencionSeleccionada As Long
Dim ml_TipoFiltro As sghTipoFiltroAdmision
Dim mo_cmbIdResponsable As New sighEntidades.ListaDespleglable
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mrs_Tmp As New ADODB.Recordset
Dim oRsFiltrados As New Recordset
Dim oRsFiltradosConPagoConsulta As New Recordset
Public Event OnClick(oRecordset As Recordset)
Dim ml_IdServicioConCamaDisponible As Long
Dim ml_idUsuario As Long
Dim mb_NecesitaTriaje As Boolean
Dim ml_NoPagoConsultaEnCaja As Boolean
Dim ml_NoPasoPorTriaje As Boolean
Dim lbElConsultorioUsaModuloPerinatal As Boolean
Dim lbElConsultorioUsaModuloMaterno As Boolean
Dim lbElMedicoNOregistraDatosCE As Boolean
Dim ldFechaActualServidor As Date
Dim mc_HoraEgreso As String 'Actualizado 21102014
Dim lnIdUsuarioMedico As Long

Property Get NoPasoPorTriaje() As Long
    NoPasoPorTriaje = ml_NoPasoPorTriaje
End Property
Property Get NoPagoConsultaEnCaja() As Long
    NoPagoConsultaEnCaja = ml_NoPagoConsultaEnCaja
End Property

Property Let IdServicioConCamaDisponible(lValue As Long)
    ml_IdServicioConCamaDisponible = lValue
End Property
Property Get IdServicioConCamaDisponible() As Long
    IdServicioConCamaDisponible = ml_IdServicioConCamaDisponible
End Property

Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
    idUsuario = ml_idUsuario
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set UserControl.grdAdmision.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = UserControl.grdAdmision.DataSource
End Property
Property Let idRegistroSeleccionado(lValue As Long)
    ml_idRegistroSeleccionado = lValue
End Property
Property Get idRegistroSeleccionado() As Long
    idRegistroSeleccionado = ml_idRegistroSeleccionado
End Property
Property Let IdAtencionSeleccionada(lValue As Long)
    ml_IdAtencionSeleccionada = lValue
End Property
Property Get IdAtencionSeleccionada() As Long
    IdAtencionSeleccionada = ml_IdAtencionSeleccionada
End Property

Property Let Titulo(lValue As String)
    lblNombre = lValue
End Property
Property Get Titulo() As String
    Titulo = lblNombre
End Property
Property Let TipoFiltro(lValue As sghTipoFiltroAdmision)
    ml_TipoFiltro = lValue
End Property
Property Get TipoFiltro() As sghTipoFiltroAdmision
    TipoFiltro = ml_TipoFiltro
End Property

Property Get HoraEgreso() As String 'Actualizado 21102014
    HoraEgreso = mc_HoraEgreso
End Property

Sub BuscaXnEmergencia()
    Dim oRsTmp1 As New Recordset
    Dim oConexion As New Connection
    Dim lnIdAtencion As Long
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighEntidades.CadenaConexion
    Set oRsTmp1 = mo_AdminAdmision.atencionesDatosAdicionalesXfiltro(" dbo.AtencionesDatosAdicionales.emergenciaCorrelativo='" & Trim(txtNemergencia.Text) & "'", oConexion)
    If oRsTmp1.RecordCount > 0 Then
       lnIdAtencion = oRsTmp1!idAtencion
       Set oRsTmp1 = mo_AdminAdmision.AtencionesSeleccionarPorIdAtencion(lnIdAtencion)
       If oRsTmp1.RecordCount > 0 Then
          txtNcuenta.Text = oRsTmp1!idCuentaAtencion
          RealizarBusqueda False
       End If
    End If
    oRsTmp1.Close
    oConexion.Close
    Set oRsTmp1 = Nothing
    Set oConexion = Nothing
End Sub

'Actualizado 14102014
Private Sub btnBuscar_Click()
    Screen.MousePointer = vbHourglass
    If txtNemergencia.Visible = True And Len(txtNemergencia.Text) > 4 Then
       BuscaXnEmergencia
    Else
       RealizarBusqueda False
    End If
    Screen.MousePointer = vbDefault
End Sub

'debb-03/08/2016
Public Sub RealizarBusqueda(lbSoloPacientesSINaltaMedica As Boolean)
Dim oDOPaciente As New doPaciente
Dim oDOAtencion As New DOAtencion
Dim lbSigue As Boolean
Dim lnListIndex As Integer
Dim rsRespuesta As New Recordset
Dim lbServicioVacio As Boolean, lcSql1 As String
Dim lcHistoriaB As String
        grdAdmision.Caption = "Lista de Admisiones"
        If lbSoloPacientesSINaltaMedica = True Then
            If cmbIdResponsable.Text = "" Then
                    Set grdAdmision.DataSource = Nothing
                    MsgBox "Por favor elija el SERVICIO", vbInformation, "Mensaje"
                    Exit Sub
            End If
        Else
            If cmbIdResponsable.Text = "" And UserControl.txtNcuenta.Text = "" And _
               UserControl.txtNroHistoria.Text = "" And txtApellidoPaterno.Text = "" And txtDni.Text = "" Then
                    Set grdAdmision.DataSource = Nothing
                    MsgBox "Por favor elija N°Cuenta, Historia, Ap.Paterno, DNI o Servicio", vbInformation, "Mensaje"
                    Exit Sub
            End If
        End If
        lcHistoriaB = ""
        If mo_Teclado.TextoEsSoloNumeros(UserControl.txtNroHistoria.Text) Then
           UserControl.txtNcuenta.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
           UserControl.txtDni.Text = ""
           UserControl.cmbFecha.ListIndex = 1
           lcHistoriaB = sighEntidades.HCigualDNI_AgregaNUEVEaLaHistoria(UserControl.txtNroHistoria.Text)
        ElseIf mo_Teclado.TextoEsSoloNumeros(UserControl.txtNcuenta.Text) Then
           UserControl.txtNroHistoria.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
           UserControl.txtDni.Text = ""
           UserControl.cmbFecha.ListIndex = 1
        ElseIf UserControl.txtApellidoPaterno.Text <> "" Then
           UserControl.txtNroHistoria.Text = ""
           UserControl.txtNcuenta.Text = ""
           UserControl.txtDni.Text = ""
           UserControl.cmbFecha.ListIndex = 1
        ElseIf txtDni.Text <> "" Then
           UserControl.txtNroHistoria.Text = ""
           UserControl.txtNcuenta.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
           UserControl.cmbFecha.ListIndex = 1
        End If
        lbServicioVacio = IIf(cmbIdResponsable.Text = "", True, False)
        lnListIndex = 0
        lbSigue = True
        Do While lbSigue
            If lbServicioVacio = True Then
               cmbIdResponsable.ListIndex = lnListIndex
               lnListIndex = lnListIndex + 1
               If lnListIndex = cmbIdResponsable.ListCount Then
                  lbSigue = False
               End If
            Else
               lbSigue = False
            End If
            lcSql1 = ""
            Select Case ml_TipoFiltro
            Case sghFiltrarConsultaExterna
                Set oRsFiltrados = mo_AdminAdmision.AtencionesSeleccionarCEPorCuentaPorHistoriaPorApellidosPorServicio(Val(lcHistoriaB), Val(UserControl.txtNcuenta.Text), Trim(UserControl.txtApellidoPaterno.Text), UserControl.cmbFecha.Text, Val(mo_cmbIdResponsable.BoundText), txtDni.Text)
                FiltraCeMarcandoLosQueNoPagaron
            Case sghFiltrarEmergencia
                Set oRsFiltrados = mo_AdminAdmision.AtencionesSeleccionarEmergPorCuentaPorHistoriaPorApellidosPorServicio(Val(lcHistoriaB), Val(UserControl.txtNcuenta.Text), Trim(UserControl.txtApellidoPaterno.Text), UserControl.cmbFecha.Text, Val(mo_cmbIdResponsable.BoundText), txtDni.Text)
                If txtApellidoPaterno.Text = wxSinApellido Then
                   lcSql1 = "ApellidoPaterno='" & wxSinApellido & "'"
                End If
                If chkActivos.Value = 1 Then
                   If lcSql1 = "" Then
                      lcSql1 = "IdEstadoAtencion<>0"
                   Else
                      lcSql1 = lcSql1 & " and IdEstadoAtencion<>0"
                   End If
                End If
                If lcSql1 <> "" Then
                   oRsFiltrados.Filter = lcSql1
                End If
                HospitalizacionYemergencia lbSoloPacientesSINaltaMedica
            Case sghFiltrarHospitalizacion
                Set oRsFiltrados = mo_AdminAdmision.AtencionesSeleccionarHospPorCuentaPorHistoriaPorApellidosPorServicio(Val(lcHistoriaB), Val(UserControl.txtNcuenta.Text), Trim(UserControl.txtApellidoPaterno.Text), UserControl.cmbFecha.Text, Val(mo_cmbIdResponsable.BoundText), txtDni.Text)
                If txtApellidoPaterno.Text = wxSinApellido Then
                   lcSql1 = "ApellidoPaterno='" & wxSinApellido & "'"
                End If
                If chkActivos.Value = 1 Then
                   If lcSql1 = "" Then
                      lcSql1 = "IdEstadoAtencion<>0"
                   Else
                      lcSql1 = lcSql1 & " and IdEstadoAtencion<>0"
                   End If
                End If
                If lcSql1 <> "" Then
                   oRsFiltrados.Filter = lcSql1
                End If
                HospitalizacionYemergencia lbSoloPacientesSINaltaMedica             'Set grdAdmision.DataSource = oRsFiltrados
            End Select
            Set rsRespuesta = grdAdmision.DataSource
            If rsRespuesta.RecordCount > 0 Then
               If ml_TipoFiltro = sghFiltrarConsultaExterna And (Val(UserControl.txtNcuenta.Text) > 0 Or _
                                  Val(UserControl.txtDni.Text) > 0 Or Val(UserControl.txtNroHistoria.Text) > 0) Then
                  mo_cmbIdResponsable.BoundText = Trim(Str(rsRespuesta!IdServicioIngreso))
               End If
               Exit Do
            End If
        Loop
        On Error Resume Next
        If rsRespuesta.RecordCount = 0 Then
            MsgBox "No se encontraron coincidencias", vbInformation, "Búsqueda"
            mo_cmbIdResponsable.BoundText = "" 'Actualizado 15102014
        Else
'            If ml_TipoFiltro = sghFiltrarHospitalizacion Then
'               LimpiarFiltro False
'            End If
            On Error Resume Next
            grdAdmision.SetFocus
        End If
        If mo_AdminAdmision.MensajeError <> "" Then
            MsgBox mo_AdminAdmision.MensajeError, vbInformation, "Filtro Pacientes"
        End If
        'mo_Apariencia.ConfigurarFilasBiColores grdAdmision, SIGHEntidades.GrillaConFilasBicolor
        
End Sub

Sub CreaTemporal()
    If oRsFiltradosConPagoConsulta.State = adStateOpen Then
       Set oRsFiltradosConPagoConsulta = Nothing
    End If
    With oRsFiltradosConPagoConsulta
        .Fields.Append "IdTipoServicio", adInteger
        .Fields.Append "IdPaciente", adInteger
        .Fields.Append "IdAtencion", adInteger
        .Fields.Append "IdCuentaAtencion", adInteger
        .Fields.Append "ApellidoPaterno", adVarChar, 100, adFldIsNullable
        .Fields.Append "ApellidoMaterno", adVarChar, 100, adFldIsNullable
        .Fields.Append "PrimerNombre", adVarChar, 100, adFldIsNullable
        .Fields.Append "SegundoNombre", adVarChar, 100, adFldIsNullable
        .Fields.Append "NroHistoriaClinica", adInteger
        If ml_TipoFiltro = sghFiltrarConsultaExterna Then
           .Fields.Append "FichaFamiliar", adVarChar, 20, adFldIsNullable
        End If
        .Fields.Append "FecNacim", adDate
        .Fields.Append "FechaIngreso", adDate
        .Fields.Append "HoraIngreso", adVarChar, 5, adFldIsNullable
        .Fields.Append "FechaEgreso", adDate, , adFldIsNullable
        .Fields.Append "HoraEgreso", adVarChar, 5, adFldIsNullable
        .Fields.Append "ServicioActual", adVarChar, 100, adFldIsNullable
        .Fields.Append "Edad", adInteger
        .Fields.Append "PagoConsulta", adVarChar, 10
        .Fields.Append "IdServicioIngreso", adInteger
        .Fields.Append "TipoNumeracion", adVarChar, 100, adFldIsNullable
        .Fields.Append "FechaEgresoAdministrativo", adDate, , adFldIsNullable
        .Fields.Append "HoraEgresoAdministrativo", adVarChar, 5, adFldIsNullable
        .Fields.Append "Plan", adVarChar, 100, adFldIsNullable
        .Fields.Append "idEstadoAtencion", adInteger
        .Fields.Append "idCita", adInteger
        .Fields.Append "Usuario", adVarChar, 40, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Sub FiltraCeMarcandoLosQueNoPagaron()
    mb_NecesitaTriaje = False
    CreaTemporal
    If txtApellidoPaterno.Text = wxSinApellido Then
       oRsFiltrados.Filter = "ApellidoPaterno='" & wxSinApellido & "' and idEstadoAtencion<>0"
    Else
       oRsFiltrados.Filter = "idEstadoAtencion<>0"
    End If
    
    If oRsFiltrados.RecordCount > 0 Then
       Dim oRsTmp As New Recordset, oRsTmp1 As New Recordset, oRsTmp2 As New Recordset
       Dim oRsTmp3 As New Recordset
       Dim lcPago As String, lcPasoTriaje As String
       Dim oPaciente As New doPaciente
       Dim oConexion As New Connection
       Dim lcSql As String, lnIdMedico As Long
       Dim lbContinuar As Boolean
       Dim lbElServicioEsCostoCero As Boolean
       Dim lcBuscaParametro As New SIGHDatos.Parametros
       Dim oConexionExterna As New Connection
       Dim lnNroRegistrosTriaje As Long, lnNroRegistrosPagantes As Long, lbEsUnEPS As Boolean
       '
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.CommandTimeout = 150
       oConexionExterna.Open wxParametroJAMO
       '
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open sighEntidades.CadenaConexion

       '
       lbElServicioEsCostoCero = mo_AdminAdmision.EsServicioCostoCero(Val(mo_cmbIdResponsable.BoundText))
       
       Dim lIdServicioAtencion As Long
       Dim dFechaIngresoAtencion As Date
       
       lIdServicioAtencion = 0
       dFechaIngresoAtencion = 0
       '
       oRsFiltrados.MoveFirst
       Do While Not oRsFiltrados.EOF
            'actualizado 20140919
            If lIdServicioAtencion <> oRsFiltrados.Fields!IdServicioIngreso Or dFechaIngresoAtencion <> oRsFiltrados.Fields!FechaIngreso Then
                Set oRsTmp1 = mo_AdminAdmision.atencionesCExServicio(oRsFiltrados.Fields!IdServicioIngreso, oRsFiltrados.Fields!FechaIngreso, oConexionExterna)
                lnNroRegistrosTriaje = oRsTmp1.RecordCount
                '
                Set oRsTmp2 = mo_AdminAdmision.AtencionesParaAtencionPagantesDelMedico(oRsFiltrados.Fields!IdServicioIngreso, oRsFiltrados.Fields!FechaIngreso, oConexion)
                'oRsTmp2.Filter = "idTipoFinanciamiento=1"
                lnNroRegistrosPagantes = oRsTmp2.RecordCount
                
                lIdServicioAtencion = oRsFiltrados.Fields!IdServicioIngreso
                dFechaIngresoAtencion = oRsFiltrados.Fields!FechaIngreso
            End If

            'El usuario es un MEDICO, por lo tanto solo mostrará sus pacientes
            lbContinuar = True
            If lnIdUsuarioMedico > 0 Then
               If lnIdUsuarioMedico <> oRsFiltrados.Fields!IdMedicoIngreso Then
                  lbContinuar = False
               End If
            End If
            'solo se atienden las cuentas con fechas menores o iguales a HOY
            If ldFechaActualServidor < oRsFiltrados.Fields!FechaIngreso Then
               lbContinuar = False
            End If
            '
            If lbContinuar = True Then
                oRsTmp2.Filter = ""
                lbEsUnEPS = False
                If Not IsNull(oRsFiltrados!EpsPorcentaje) Then
                   If oRsFiltrados!EpsPorcentaje > 0 Then
                      lbEsUnEPS = True
                      oRsTmp2.Filter = "idTipoFinanciamiento=1"
                   End If
                End If
                'pago Consulta
                lcPago = ""
                If oRsFiltrados.Fields!generaPago = 1 Or lbEsUnEPS = True Then
                    If lbElServicioEsCostoCero = False Then
                        If lnNroRegistrosPagantes > 0 Then
                          ' If oRsTmp2.RecordCount > 0 Then
                                oRsTmp2.MoveFirst
                                oRsTmp2.Find "idCuentaAtencion=" & oRsFiltrados.Fields!nrocuenta
                                If oRsTmp2.EOF Then
                                   lcPago = "No Pagó"
                                ElseIf oRsTmp2.Fields!idestadofacturacion <> 4 Then
                                   lcPago = "No Pagó"
                                End If
                         '  End If
                        Else
                           lcPago = "No Pagó"
                        End If
                    End If
                End If
                If lcPago <> "" Then
                   Set oRsTmp3 = mo_AdminFacturacion.FacturacionPaquetesCEporIdPuntoCargaNrocuentaIdEspecialidad(oRsFiltrados.Fields!nrocuenta, oRsFiltrados.Fields!IdEspecialidad, 6, oConexion)
                   If oRsTmp3.RecordCount > 0 Then
                      lcPago = ""
                   End If
                   oRsTmp3.Close
                End If
                'pasó por Triaje
                lcPasoTriaje = "No"
                If lnNroRegistrosTriaje > 0 Then
                   oRsTmp1.MoveFirst
                   oRsTmp1.Find "idAtencion=" & oRsFiltrados.Fields!idAtencion
                   If (Not oRsTmp1.EOF) And (Not IsNull(oRsTmp1.Fields!TriajeFecha)) Then
                       lcPasoTriaje = ""
                   End If
                End If
                '
                oRsFiltradosConPagoConsulta.AddNew
                oRsFiltradosConPagoConsulta.Fields!idTipoServicio = 1
                oRsFiltradosConPagoConsulta.Fields!idPaciente = oRsFiltrados.Fields!idPaciente
                oRsFiltradosConPagoConsulta.Fields!idAtencion = oRsFiltrados.Fields!idAtencion
                oRsFiltradosConPagoConsulta.Fields!idCuentaAtencion = oRsFiltrados.Fields!nrocuenta
                oRsFiltradosConPagoConsulta.Fields!ApellidoPaterno = oRsFiltrados.Fields!ApellidoPaterno
                oRsFiltradosConPagoConsulta.Fields!ApellidoMaterno = oRsFiltrados.Fields!ApellidoMaterno
                oRsFiltradosConPagoConsulta.Fields!PrimerNombre = oRsFiltrados.Fields!PrimerNombre
                oRsFiltradosConPagoConsulta.Fields!SegundoNombre = oRsFiltrados.Fields!SegundoNombre
                oRsFiltradosConPagoConsulta.Fields!NroHistoriaClinica = IIf(IsNull(oRsFiltrados.Fields!NroHistoriaClinica), 0, HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsFiltrados.Fields!NroHistoriaClinica)), False))
                oRsFiltradosConPagoConsulta.Fields!FecNacim = IIf(IsNull(oRsFiltrados.Fields!FecNacim), "01/01/1900", oRsFiltrados.Fields!FecNacim) 'Actualizado 22092014
                oRsFiltradosConPagoConsulta.Fields!FechaIngreso = oRsFiltrados.Fields!FechaIngreso
                oRsFiltradosConPagoConsulta.Fields!HoraIngreso = oRsFiltrados.Fields!HoraIngreso
                'oRsFiltradosConPagoConsulta.Fields!FechaEgreso = oRsFiltrados.Fields!FechaEgreso
                oRsFiltradosConPagoConsulta.Fields!HoraEgreso = oRsFiltrados.Fields!HoraEgreso
                oRsFiltradosConPagoConsulta.Fields!ServicioActual = lcPasoTriaje
                oRsFiltradosConPagoConsulta.Fields!Edad = oRsFiltrados.Fields!Edad
                'oRsFiltradosConPagoConsulta.Fields!IdTipoNumeracion = oRsFiltrados.Fields!IdTipoNumeracion
                oRsFiltradosConPagoConsulta.Fields!IdServicioIngreso = oRsFiltrados.Fields!IdServicioIngreso
                oRsFiltradosConPagoConsulta.Fields!TipoNumeracion = oRsFiltrados.Fields!TipoNumeracion
                'oRsFiltradosConPagoConsulta.Fields!FechaEgresoAdministrativo = oRsFiltrados.Fields!FechaEgresoAdministrativo
                'oRsFiltradosConPagoConsulta.Fields!HoraEgresoAdministrativo = oRsFiltrados.Fields!HoraEgresoAdministrativo
                'oRsFiltradosConPagoConsulta.Fields!Plan = oRsFiltrados.Fields!Plan
                oRsFiltradosConPagoConsulta.Fields!IdEstadoAtencion = oRsFiltrados.Fields!IdEstadoAtencion
                oRsFiltradosConPagoConsulta.Fields!PagoConsulta = lcPago
                oRsFiltradosConPagoConsulta.Fields!IdCita = IIf(IsNull(oRsFiltrados.Fields!IdCita), 0, oRsFiltrados.Fields!IdCita)
                If ml_TipoFiltro = sghFiltrarConsultaExterna Then
                   oRsFiltradosConPagoConsulta.Fields!FichaFamiliar = IIf(IsNull(oRsFiltrados.Fields!FichaFamiliar), "", oRsFiltrados.Fields!FichaFamiliar)
                End If
                oRsFiltradosConPagoConsulta.Update
            End If
            oRsFiltrados.MoveNext
       Loop
       On Error Resume Next
       oRsFiltradosConPagoConsulta.MoveFirst
       Set oRsTmp = Nothing
       Set oRsTmp1 = Nothing
       Set oRsTmp2 = Nothing
       Set oRsTmp3 = Nothing
       Set oPaciente = Nothing
       mb_NecesitaTriaje = mo_AdminAdmision.ElServicioNecesitaTriaje(IIf(oRsFiltrados.RecordCount = 1, lIdServicioAtencion, Val(mo_cmbIdResponsable.BoundText)), oConexion, _
                          lbElConsultorioUsaModuloPerinatal, lbElConsultorioUsaModuloMaterno, lbElMedicoNOregistraDatosCE)
       oConexion.Close
       Set oConexion = Nothing
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       Set lcBuscaParametro = Nothing
    End If
    Set grdAdmision.DataSource = oRsFiltradosConPagoConsulta
End Sub



Sub HospitalizacionYemergencia(lbSolPacientesSinAltaMedica As Boolean)
    Dim lcPago As String
    Dim oConexion As New ADODB.Connection
    Dim oRsTmp1 As New Recordset
    Dim oRsTmp2 As New Recordset
    Dim lbContinuar As Boolean, lcUsuario As String, lnLineas As Long
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    oConexion.Open sighEntidades.CadenaConexion
    CreaTemporal
    lnLineas = 1
    If oRsFiltrados.RecordCount > 0 Then
       oRsFiltrados.MoveFirst
       Do While Not oRsFiltrados.EOF
            lbContinuar = True
            'El usuario es un MEDICO, por lo tanto solo mostrará sus pacientes
'            If lnIdUsuarioMedico > 0 Then
'               If IsNull(oRsFiltrados.Fields!FechaEgreso) Then
'                    Set oRsTmp1 = mo_AdminAdmision.EstanciaHospitalariaSeleccionarPorAtencion(oRsFiltrados.Fields!idAtencion, 0, oConexion)
'                    If oRsTmp1.RecordCount > 0 Then
'                       oRsTmp1.MoveLast
'                       If oRsTmp1!IdMedicoOrdena <> lnIdUsuarioMedico Then
'                          lbContinuar = False
'                       End If
'                    Else
'                       lbContinuar = False
'                    End If
'                    oRsTmp1.Close
'               ElseIf lnIdUsuarioMedico <> oRsFiltrados.Fields!IdMedicoEgreso Then
'                   lbContinuar = False
'               End If
'            End If
            '
            If lbSolPacientesSinAltaMedica = True And lbContinuar = True Then
               If Not IsNull(oRsFiltrados.Fields!fechaEgreso) Then
                  lbContinuar = False
               End If
            End If
            '
            If lbContinuar = True And oRsFiltradosConPagoConsulta.RecordCount > 0 Then
               oRsFiltradosConPagoConsulta.MoveFirst
               oRsFiltradosConPagoConsulta.Find "idCuentaAtencion = " & oRsFiltrados.Fields!idCuentaAtencion
               If Not oRsFiltradosConPagoConsulta.EOF Then
                  lbContinuar = False
               End If
            End If
            '
            If lbContinuar = True Then
                lcUsuario = ""
                If lbSolPacientesSinAltaMedica = True Then
                    Set oRsTmp2 = mo_ReglasDeSeguridad.AuditoriaFiltrarCitasPorIdAtencion(oRsFiltrados!idAtencion, _
                                                                         IIf(txtNemergencia.Visible = True, sghAdmisionEmergencia, sghAdmisionHospitalizacion), oConexion)
                    If oRsTmp2.RecordCount > 0 Then
                       lcUsuario = IIf(IsNull(oRsTmp2!dUsuario), "", Left(oRsTmp2!dUsuario, 40))
                    End If
                    oRsTmp2.Close
                
                End If
                '
                lcPago = ""
                If ml_TipoFiltro = sghFiltrarConsultorioEmergencia Then
                    If VerSiTieneServicioAutomaticoPorEstancia(oRsFiltrados.Fields!idAtencion, oConexion) = "" Then
                       lcPago = "Pago"
                    End If
                End If
                oRsFiltradosConPagoConsulta.AddNew
                oRsFiltradosConPagoConsulta.Fields!idTipoServicio = IIf(ml_TipoFiltro = sghFiltrarHospitalizacion, 3, 2)
                oRsFiltradosConPagoConsulta.Fields!idPaciente = oRsFiltrados.Fields!idPaciente
                oRsFiltradosConPagoConsulta.Fields!idAtencion = oRsFiltrados.Fields!idAtencion
                oRsFiltradosConPagoConsulta.Fields!idCuentaAtencion = oRsFiltrados.Fields!idCuentaAtencion
                oRsFiltradosConPagoConsulta.Fields!ApellidoPaterno = oRsFiltrados.Fields!ApellidoPaterno
                oRsFiltradosConPagoConsulta.Fields!ApellidoMaterno = oRsFiltrados.Fields!ApellidoMaterno
                oRsFiltradosConPagoConsulta.Fields!PrimerNombre = oRsFiltrados.Fields!PrimerNombre
                oRsFiltradosConPagoConsulta.Fields!SegundoNombre = oRsFiltrados.Fields!SegundoNombre
                oRsFiltradosConPagoConsulta.Fields!NroHistoriaClinica = IIf(IsNull(oRsFiltrados.Fields!NroHistoriaClinica), 0, HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(Str(oRsFiltrados.Fields!NroHistoriaClinica)), False))
                If Not IsNull(oRsFiltrados.Fields!FecNacim) Then
                   oRsFiltradosConPagoConsulta.Fields!FecNacim = oRsFiltrados.Fields!FecNacim
                End If
                oRsFiltradosConPagoConsulta.Fields!FechaIngreso = oRsFiltrados.Fields!FechaIngreso
                oRsFiltradosConPagoConsulta.Fields!HoraIngreso = oRsFiltrados.Fields!HoraIngreso
                oRsFiltradosConPagoConsulta.Fields!fechaEgreso = oRsFiltrados.Fields!fechaEgreso
                oRsFiltradosConPagoConsulta.Fields!HoraEgreso = oRsFiltrados.Fields!HoraEgreso
                oRsFiltradosConPagoConsulta.Fields!ServicioActual = oRsFiltrados.Fields!ServicioActual
                oRsFiltradosConPagoConsulta.Fields!Edad = oRsFiltrados.Fields!Edad
                oRsFiltradosConPagoConsulta.Fields!IdServicioIngreso = oRsFiltrados.Fields!IdServicioIngreso
                oRsFiltradosConPagoConsulta.Fields!TipoNumeracion = oRsFiltrados.Fields!TipoNumeracion
                oRsFiltradosConPagoConsulta.Fields!FechaEgresoAdministrativo = oRsFiltrados.Fields!FechaEgresoAdministrativo
                oRsFiltradosConPagoConsulta.Fields!HoraEgresoAdministrativo = oRsFiltrados.Fields!HoraEgresoAdministrativo
                oRsFiltradosConPagoConsulta.Fields!Plan = oRsFiltrados.Fields!Plan
                oRsFiltradosConPagoConsulta.Fields!IdEstadoAtencion = oRsFiltrados.Fields!IdEstadoAtencion
                oRsFiltradosConPagoConsulta.Fields!PagoConsulta = lcPago
                oRsFiltradosConPagoConsulta.Fields!Usuario = lcUsuario
                oRsFiltradosConPagoConsulta.Update
                If Val(TxtNroLineasAmostrar) < lnLineas Then
                   Exit Do
                End If
                lnLineas = lnLineas + 1
            End If
            oRsFiltrados.MoveNext
       Loop
       If oRsFiltradosConPagoConsulta.RecordCount > 0 Then
          oRsFiltradosConPagoConsulta.MoveFirst
       End If
    End If
    Set grdAdmision.DataSource = oRsFiltradosConPagoConsulta
    oConexion.Close
    Set oConexion = Nothing
    Set oRsTmp1 = Nothing
    Set oRsTmp2 = Nothing
End Sub


Function VerSiTieneServicioAutomaticoPorEstancia(lnIdAtencion As Long, oConexion As Connection) As String
    Dim lcSql As String
    Dim oRsTmp As New ADODB.Recordset
'    Dim oConexion As New ADODB.Connection
'    oConexion.Open sighEntidades.CadenaConexion
'    oConexion.CursorLocation = adUseClient
'    oConexion.CommandTimeout = 150
    VerSiTieneServicioAutomaticoPorEstancia = ""
    Set oRsTmp = mo_AdminFacturacion.FactOrdenServicioPagosPorIdAtencion(lnIdAtencion, oConexion)
    oRsTmp.Filter = "idPuntoCarga=10"
    If oRsTmp.RecordCount > 0 Then
       VerSiTieneServicioAutomaticoPorEstancia = Chr(13) & "(Ord.Pago)= "
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          VerSiTieneServicioAutomaticoPorEstancia = VerSiTieneServicioAutomaticoPorEstancia & Trim(Str(oRsTmp.Fields!IdOrdenPago)) & " , "
          oRsTmp.MoveNext
       Loop
    End If
    oRsTmp.Close
 '   oConexion.Close
    Set oRsTmp = Nothing
 '   Set oConexion = Nothing
End Function

Private Sub btnLimpiar_Click()
    LimpiarFiltro False
    Set grdAdmision.DataSource = Nothing
End Sub
Public Sub LimpiarFiltro(lbSinAlta As Boolean)
        UserControl.txtApellidoPaterno = ""
        UserControl.txtNroHistoria = ""
        UserControl.txtNcuenta = ""
        If lbSinAlta = False Then cmbFecha.ListIndex = 0
        If lbSinAlta = False Then mo_cmbIdResponsable.BoundText = ""
        txtDni.Text = ""
        mb_NecesitaTriaje = False
        txtNemergencia.Text = Trim(Str(Year(Date)))  'debb-06/07/2016
        On Error Resume Next
        Select Case ml_TipoFiltro
        Case sghFiltrarConsultaExterna
             UserControl.txtNroHistoria.SetFocus
        Case Else
             UserControl.txtNcuenta.SetFocus
        End Select
End Sub

Private Sub btnRefrescar_Click()
    CargaDisponibilidadCamas
End Sub



Private Sub cmbFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbFecha
    AdministrarKeyPreview KeyCode
End Sub





Private Sub cmbFecha_LostFocus()
    If cmbFecha.ListIndex <> 1 Then
        If Not EsFecha(cmbFecha.Text, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, ""
            cmbFecha.Text = Date
            Exit Sub
        End If

    End If

End Sub

Private Sub cmbIdResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsable
    AdministrarKeyPreview KeyCode

End Sub

'Admision 14102014
Public Sub BuscarPacientesSinAltaMedica()
    Dim oRsTmp1 As New Recordset
    Dim lcFecha As String
    lcFecha = ""
    If sighEntidades.EsFecha(cmbFecha.Text, "DD/MM/AAAA") = True Then
        lcFecha = cmbFecha.Text
    End If
    Set oRsTmp1 = mo_AdminAdmision.AdmisionBuscarPacientesAltaMedica(lcFecha, IIf(mo_cmbIdResponsable.BoundText = "", 0, Val(mo_cmbIdResponsable.BoundText)))
    Set grdAdmision.DataSource = oRsTmp1
    
    grdAdmision.Caption = "Lista de Pacientes sin Alta Médica"
    'mo_Apariencia.ConfigurarFilasBiColores grdAdmision, sighentidades.GrillaConFilasBicolor
    
    If oRsTmp1.RecordCount = 0 Then
        MsgBox "No se encontraron coincidencias", vbInformation, "Búsqueda"
        mo_cmbIdResponsable.BoundText = ""
    End If


End Sub



Private Sub cmdPacientesSinAltaHE_Click()
    Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
    Dim oRsTmp96345 As New Recordset
    MousePointer = 11
    Set oRsTmp96345 = mo_AdminAdmision.PacientesSinAltaMedicaEnHospEmerg
    mo_ReglasReportes.ExportarRecordSetAexcel oRsTmp96345, _
                     "Lista de Pacientes sin Alta Médica en Hospitalización y Emergencia", _
                     "", "", 0, False, True
                     
    Set oRsTmp96345 = Nothing
    Set mo_ReglasReportes = Nothing
    MousePointer = 1
End Sub

Private Sub cmdPacientesSinAltaMedica_Click()
           UserControl.txtNcuenta.Text = ""
           UserControl.txtApellidoPaterno.Text = ""
           UserControl.txtDni.Text = ""
           UserControl.txtNroHistoria.Text = ""
           UserControl.cmbFecha.ListIndex = 1
           Screen.MousePointer = vbHourglass
           RealizarBusqueda True
           Screen.MousePointer = vbDefault
'        LimpiarFiltro True
'        BuscarPacientesSinAltaMedica

End Sub

Private Sub cmdSinApellidoPaterno_Click()
    txtApellidoPaterno.Text = wxSinApellido
End Sub



Private Sub grdAdmision_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    ml_NoPasoPorTriaje = False
    ml_NoPagoConsultaEnCaja = False
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
'        If lbElMedicoNOregistraDatosCE = True Then
'           ml_idRegistroSeleccionado = 0
'           ml_IdAtencionSeleccionada = 0
'           MsgBox "El Consultorio se configuró para que no se ingrese DATOS", vbInformation, ""
'        End If
        ml_idRegistroSeleccionado = rsRecordset("IdCita")
        ml_IdAtencionSeleccionada = rsRecordset("IdAtencion")
        If (rsRecordset.Fields!PagoConsulta) <> "" Or (Trim(rsRecordset.Fields!ServicioActual) <> "" And mb_NecesitaTriaje = True) Then
           If rsRecordset.Fields!PagoConsulta <> "" Then
              ml_NoPagoConsultaEnCaja = True
           End If
           If mb_NecesitaTriaje = True And rsRecordset.Fields!ServicioActual <> "" Then
              ml_NoPasoPorTriaje = True
           End If
           ml_idRegistroSeleccionado = 0
           ml_IdAtencionSeleccionada = 0
        End If
    Case sghFiltrarHospitalizacion
        ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
        ml_IdAtencionSeleccionada = rsRecordset("IdAtencion")
    Case sghFiltrarEmergencia
        ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
        ml_IdAtencionSeleccionada = rsRecordset("IdAtencion")
    End Select
    RaiseEvent OnClick(rsRecordset)
End Sub

'Actualizado 15102014
Private Sub grdAdmision_BeforeRowsDeleted(ByVal Rows As UltraGrid.SSSelectedRows, ByVal DisplayPromptMsg As UltraGrid.SSReturnBoolean, ByVal Cancel As UltraGrid.SSReturnBoolean)
    Cancel = True
End Sub

Private Sub grdAdmision_Click()
Dim rsRecordset As ADODB.Recordset

    On Error Resume Next
    ml_idRegistroSeleccionado = -1
    Set rsRecordset = grdAdmision.DataSource
    On Error Resume Next
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
        ml_idRegistroSeleccionado = rsRecordset("IdCita")
        ml_IdAtencionSeleccionada = rsRecordset("IdCuentaAtencion")
        mc_HoraEgreso = IIf(IsNull(rsRecordset.Fields!HoraEgreso), "", rsRecordset.Fields!HoraEgreso) 'Actualizado 21102014
        
        If (rsRecordset.Fields!PagoConsulta) <> "" Or (Trim(rsRecordset.Fields!ServicioActual) <> "" And mb_NecesitaTriaje = True) Then
           ml_idRegistroSeleccionado = 0
           ml_IdAtencionSeleccionada = 0
        End If
        
    Case sghFiltrarHospitalizacion
        ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
        ml_IdAtencionSeleccionada = rsRecordset("IdCuentaAtencion")
    Case sghFiltrarEmergencia
        ml_idRegistroSeleccionado = rsRecordset("IdAtencion")
        ml_IdAtencionSeleccionada = rsRecordset("IdCuentaAtencion")
    End Select
    RaiseEvent OnClick(rsRecordset)
    
End Sub


Private Sub grdAdmision_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    On Error Resume Next
    grdAdmision.Bands(0).Columns("IdPaciente").Hidden = True
    grdAdmision.Bands(0).Columns("IdAtencion").Hidden = True
    
    grdAdmision.Bands(0).Columns("IdTipoNumeracion").Hidden = True
    grdAdmision.Bands(0).Columns("IdServicioIngreso").Hidden = True
    
    grdAdmision.Bands(0).Columns("IdAtencion").Header.Caption = "N° atención"
    grdAdmision.Bands(0).Columns("IdAtencion").Width = 1300
    
    grdAdmision.Bands(0).Columns("FechaIngreso").Header.Caption = "Fecha Ing."
    grdAdmision.Bands(0).Columns("FechaIngreso").Width = 1000
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("FechaIngreso").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("HoraIngreso").Header.Caption = "Hr.Ing"
    grdAdmision.Bands(0).Columns("HoraIngreso").Width = 800
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("HoraIngreso").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "N° Cuenta"
    grdAdmision.Bands(0).Columns("IdCuentaAtencion").Width = 1300
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("IdCuentaAtencion").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("TipoNumeracion").Header.Caption = "Tipo Numeración"
    grdAdmision.Bands(0).Columns("TipoNumeracion").Width = 1500
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("TipoNumeracion").Activation = ssActivationActivateNoEdit
    End If
    
    If ml_TipoFiltro = sghFiltrarConsultaExterna Then
        grdAdmision.Bands(0).Columns("ServicioIngreso").Header.Caption = "Servicio Ing"
        grdAdmision.Bands(0).Columns("ServicioIngreso").Width = 2500
        If wxParametro541 <> "S" Then
            grdAdmision.Bands(0).Columns("ServicioIngreso").Activation = ssActivationActivateNoEdit
            grdAdmision.Bands(0).Columns("ServicioActual").Activation = ssActivationActivateNoEdit
            grdAdmision.Bands(0).Columns("ServicioActual").Activation = ssActivationActivateNoEdit
            grdAdmision.Bands(0).Columns("PagoConsulta").Activation = ssActivationActivateNoEdit
            grdAdmision.Bands(0).Columns("PagoConsulta").Activation = ssActivationActivateNoEdit
        End If
    Else
        grdAdmision.Bands(0).Columns("ServicioActual").Header.Caption = "Servicio Actual"
        grdAdmision.Bands(0).Columns("ServicioActual").Width = 2500
        If wxParametro541 <> "S" Then
            grdAdmision.Bands(0).Columns("ServicioActual").Activation = ssActivationActivateNoEdit
        End If
    End If
    grdAdmision.Bands(0).Columns("Edad").Hidden = True
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("Edad").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap. Paterno"
    grdAdmision.Bands(0).Columns("ApellidoPaterno").Width = 1500
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("ApellidoPaterno").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap. Materno"
    grdAdmision.Bands(0).Columns("ApellidoMaterno").Width = 1500
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("ApellidoMaterno").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("PrimerNombre").Header.Caption = "1er Nombre"
    grdAdmision.Bands(0).Columns("PrimerNombre").Width = 1500
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("PrimerNombre").Activation = ssActivationActivateNoEdit
    End If

'    grdAdmision.Bands(0).Columns("SegundoNombre").Header.Caption = "2do Nombre"
'    grdAdmision.Bands(0).Columns("SegundoNombre").Width = 10
    grdAdmision.Bands(0).Columns("SegundoNombre").Hidden = True
    

    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Header.Caption = "N° HC"
    grdAdmision.Bands(0).Columns("NroHistoriaClinica").Width = 1200
    If wxParametro541 <> "S" Then
        grdAdmision.Bands(0).Columns("NroHistoriaClinica").Activation = ssActivationActivateNoEdit
    End If

    grdAdmision.Bands(0).Columns("IdTipoServicio").Hidden = True
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("FecNacim").Activation = ssActivationActivateNoEdit
    End If
    
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("Plan").Activation = ssActivationActivateNoEdit
    End If
    grdAdmision.Bands(0).Columns("IdEstadoAtencion").Hidden = True
    grdAdmision.Bands(0).Columns("IdCita").Hidden = True
    
    grdAdmision.Bands(0).Columns("PagoConsulta").Header.Caption = "Pagó Consulta"
    grdAdmision.Bands(0).Columns("PagoConsulta").Width = 1500
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("PagoConsulta").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("FechaEgresoAdministrativo").Header.Caption = "F.Egr.Adm."
    grdAdmision.Bands(0).Columns("FechaEgresoAdministrativo").Width = 1750
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("FechaEgresoAdministrativo").Activation = ssActivationActivateNoEdit
    End If
    
    grdAdmision.Bands(0).Columns("HoraEgresoAdministrativo").Header.Caption = "Hr.Egr.Adm."
    grdAdmision.Bands(0).Columns("HoraEgresoAdministrativo").Width = 1750
    If wxParametro541 <> "S" Then
       grdAdmision.Bands(0).Columns("HoraEgresoAdministrativo").Activation = ssActivationActivateNoEdit
    End If
        
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
        grdAdmision.Bands(0).Columns("IdCita").Hidden = True
        grdAdmision.Bands(0).Columns("FechaEgreso").Header.Caption = "F.S.Atención"
        grdAdmision.Bands(0).Columns("FechaEgreso").Width = 1000
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("FechaEgreso").Activation = ssActivationActivateNoEdit
        End If
        
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("FichaFamiliar").Activation = ssActivationActivateNoEdit
        End If
        
        grdAdmision.Bands(0).Columns("HoraEgreso").Header.Caption = "Hr.S.Atención"
        grdAdmision.Bands(0).Columns("HoraEgreso").Width = 800
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("HoraEgreso").Activation = ssActivationActivateNoEdit
        End If
        
        grdAdmision.Bands(0).Columns("ServicioActual").Header.Caption = "Pasó Triaje"
        grdAdmision.Bands(0).Columns("ServicioActual").Width = 1500
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("ServicioActual").Activation = ssActivationActivateNoEdit
        End If
    Case sghFiltrarEmergencia
        grdAdmision.Bands(0).Columns("FechaEgreso").Header.Caption = "F.Egr.Médico"
        grdAdmision.Bands(0).Columns("FechaEgreso").Width = 1000
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("FechaEgreso").Activation = ssActivationActivateNoEdit
        End If
        
        grdAdmision.Bands(0).Columns("HoraEgreso").Header.Caption = "Hr.Egr.Médico"
        grdAdmision.Bands(0).Columns("HoraEgreso").Width = 800
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("HoraEgreso").Activation = ssActivationActivateNoEdit
        End If
        
    Case sghFiltrarHospitalizacion
        grdAdmision.Bands(0).Columns("FechaEgreso").Header.Caption = "F.Egr.Médico"
        grdAdmision.Bands(0).Columns("FechaEgreso").Width = 1000
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("FechaEgreso").Activation = ssActivationActivateNoEdit
        End If
        
        grdAdmision.Bands(0).Columns("HoraEgreso").Header.Caption = "Hr.Egr.Médico"
        grdAdmision.Bands(0).Columns("HoraEgreso").Width = 800
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("HoraEgreso").Activation = ssActivationActivateNoEdit
        End If

        grdAdmision.Bands(0).Columns("DxPrincipal").Header.Caption = "Dx Prin."
        grdAdmision.Bands(0).Columns("DxPrincipal").Width = 1000
        If wxParametro541 <> "S" Then
           grdAdmision.Bands(0).Columns("DxPrincipal").Activation = ssActivationActivateNoEdit
        End If
    
        grdAdmision.Bands(0).Columns("TipoAlta").Header.Caption = "Tipo Alta"
        grdAdmision.Bands(0).Columns("TipoAlta").Width = 2500
        If wxParametro541 <> "S" Then
            grdAdmision.Bands(0).Columns("TipoAlta").Activation = ssActivationActivateNoEdit
        End If
    
        grdAdmision.Bands(0).Columns("CondicionAlta").Header.Caption = "Cond. Alta"
        grdAdmision.Bands(0).Columns("CondicionAlta").Width = 2500
        If wxParametro541 <> "S" Then
          grdAdmision.Bands(0).Columns("CondicionAlta").Activation = ssActivationActivateNoEdit
        End If
    
    End Select


End Sub

Private Sub grdAdmision_InitializeRow(ByVal Context As UltraGrid.Constants_Context, ByVal Row As UltraGrid.SSRow, ByVal ReInitialize As Boolean)
        If Val(Row.Cells("IdEstadoAtencion").GetText()) = 0 Then
            Row.Appearance.ForeColor = vbRed
        ElseIf Row.Cells("HoraEgreso").GetText() <> "" And ml_TipoFiltro = sghFiltrarConsultaExterna Then
            Row.Appearance.ForeColor = vbBlue
        End If
End Sub


Private Sub grdCamasDisponibles_AfterRowActivate()
    ml_IdServicioConCamaDisponible = mrs_Tmp.Fields!IdServicio
End Sub

Private Sub grdCamasDisponibles_Click()
    ml_IdServicioConCamaDisponible = mrs_Tmp.Fields!IdServicio
End Sub

Private Sub grdCamasDisponibles_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdCamasDisponibles.Bands(0).Columns("Dpto").Width = 2500
    grdCamasDisponibles.Bands(0).Columns("Especialidad").Width = 2500
    grdCamasDisponibles.Bands(0).Columns("Servicio").Width = 2500
    grdCamasDisponibles.Bands(0).Columns("Total").Width = 1500
    grdCamasDisponibles.Bands(0).Columns("Llenas").Width = 1500
    grdCamasDisponibles.Bands(0).Columns("Disponible").Width = 1500

End Sub

Private Sub TabCamas_Click(PreviousTab As Integer)
   
      ml_IdServicioConCamaDisponible = 0
      If PreviousTab = 0 Then
         CargaDisponibilidadCamas
      End If
   
End Sub


Private Sub txtApellidoPaterno_KeyDown(KeyCode As Integer, Shift As Integer)
       mo_Teclado.RealizarNavegacion KeyCode, txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsValidoParaNombre(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
   If KeyAscii = 13 And Len(txtApellidoPaterno.Text) > 0 Then
       btnBuscar_Click
   End If
End Sub


Private Sub txtApellidoPaterno_LostFocus()
    If txtApellidoPaterno.Text <> "" Then
       txtNroHistoria.Text = ""
       txtNcuenta.Text = ""
    End If
End Sub



Private Sub txtDNI_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtDni
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
        If KeyAscii = 8 Then
            KeyAscii = 8
        Else
            KeyAscii = 1
        End If
    End If
End Sub

'Actualizado 31102014
Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
'    If txtNcuenta.Text = "" Then
        mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta
'    Else
'        If KeyCode = vbKeyReturn Then
'            btnBuscar_Click
'        End If
'    End If
End Sub

Private Sub txtNcuenta_KeyPress(KeyAscii As Integer)
'    If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
'        If KeyAscii = 8 Then
'            KeyAscii = 8
'        Else
'            KeyAscii = 1
'        End If
'    End If
    If KeyAscii = 13 And Len(txtNcuenta.Text) > 0 Then
        btnBuscar_Click
    End If
End Sub

Private Sub txtNroHistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNroHistoria
End Sub

Private Sub txtNroHistoria_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
'   If KeyAscii = 13 And Len(txtNroHistoria.Text) > 0 Then
'       btnBuscar_Click
'   End If
   
End Sub


Private Sub txtNroHistoria_LostFocus()
    If txtNroHistoria.Text <> "" Then
       txtNcuenta.Text = ""
       txtApellidoPaterno.Text = ""
    End If
End Sub



Private Sub TxtNroLineasAmostrar_KeyPress(KeyAscii As Integer)
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
End Sub

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
         btnBuscar_Click
     Case vbKeyF7
         btnLimpiar_Click
     Case vbKeyF8
    End Select
End Sub

Private Sub UserControl_Resize()
   
    On Error Resume Next
   
    
    lblNombre.Width = UserControl.Width
    
    TabCamas.Width = UserControl.Width - 100
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
        TabCamas.Height = UserControl.Height - 900 - lblNombre.Height
    Case Else
        TabCamas.Height = UserControl.Height - 800 - lblNombre.Height
    End Select
    
    fraBusqueda.Width = TabCamas.Width - 300
    
    grdAdmision.Width = fraBusqueda.Width
    grdAdmision.Height = TabCamas.Height - 700 - lblMedico.Height
    lblMedico.Top = grdAdmision.Top + grdAdmision.Height + 50
    
    grdCamasDisponibles.Width = fraBusqueda.Width
    grdCamasDisponibles.Height = TabCamas.Height - 700
    
    btnRefrescar.Top = TabCamas.Height - 100
    cmdPacientesSinAltaHE.Top = TabCamas.Height - 100
    
End Sub

Sub SkinConfigura()
  On Error GoTo ErrSkin
  If sighEntidades.Parametro282valorInt = "1" Then
        'Skin1.LoadSkin App.Path & "\" & WxSkin
        'Skin1.ApplySkin Me.hwnd
        btnBuscar.Picture = LoadPicture(App.Path & "\Binoculr.ico")
        btnBuscar.Caption = ""
        btnLimpiar.Picture = LoadPicture(App.Path & "\Refresh.ico")
        btnLimpiar.Caption = ""
        mo_Apariencia.ConfigurarFilasBiColores grdAdmision, "99"
        lblNombre.Alignment = 2
        lblNombre.BackColor = vbBlue
  Else
        mo_Apariencia.ConfigurarFilasBiColores grdAdmision, sighEntidades.GrillaConFilasBicolor
  End If
ErrSkin:
End Sub

Public Sub Inicializar()
    SkinConfigura
    
    Dim oRsTmp As New Recordset
    Dim oConexion As New Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 150
    oConexion.Open sighEntidades.CadenaConexion
    'El USUARIO es un Médico
    Set oRsTmp = mo_ReglasDeProgMedica.MedicosXidEmpleado(ml_idUsuario, oConexion)
    lnIdUsuarioMedico = 0
    If oRsTmp.RecordCount > 0 Then
       lnIdUsuarioMedico = oRsTmp!idMedico
       lblMedico.Visible = True
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
    Set oConexion = Nothing
    '
    InicilizarParametros
    '
    cmbFecha.Clear
    cmbFecha.AddItem Date
    cmbFecha.AddItem "Todas"
    cmbFecha.ListIndex = 0
    '
    mb_NecesitaTriaje = False
    '
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    mo_cmbIdResponsable.BoundColumn = "IdServicio"
    mo_cmbIdResponsable.ListField = "DservicioHosp"
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Dim lcEspecialidadesDelUsuario As String
    ml_IdServicioConCamaDisponible = 0
    lcEspecialidadesDelUsuario = ""
    txtNemergencia.Visible = False: lblNemergencia.Visible = False  'debb-06/07/2016
    lblNroLineasAmostrar.Visible = False: TxtNroLineasAmostrar.Visible = False
    
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
        TabCamas.TabVisible(1) = False
        lcEspecialidadesDelUsuario = mo_AdminAdmision.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesCE, ml_idUsuario)
        Set mo_cmbIdResponsable.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(1)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
    Case Else
        lblNroLineasAmostrar.Visible = True: TxtNroLineasAmostrar.Visible = True
        chkActivos.Visible = True
        chkActivos.Value = 1
        cmdPacientesSinAltaMedica.Visible = True
        If ml_TipoFiltro = sghFiltrarHospitalizacion Then
            TabCamas.TabVisible(1) = True
            GenerarRecordsetTemporal
            lcEspecialidadesDelUsuario = mo_AdminAdmision.DevuelveEspecialidadesServicioSegunUsuarioSistema(sghEspecialidadesHosp, ml_idUsuario)
            Set mo_cmbIdResponsable.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(3)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
        Else
            txtNemergencia.Text = Trim(Str(Year(Date))):   txtNemergencia.Visible = True: lblNemergencia.Visible = True 'debb-06/07/2016
            'cmdPacientesSinAltaMedica.Visible = True
            TabCamas.TabVisible(1) = True
            Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghEspecialidadesEmergCons, ml_idUsuario)
            If rsIdAlmacen.RecordCount > 0 Then
               lcEspecialidadesDelUsuario = " and ("
               rsIdAlmacen.MoveFirst
               Do While Not rsIdAlmacen.EOF
                  lcEspecialidadesDelUsuario = lcEspecialidadesDelUsuario & " dbo.Servicios.idEspecialidad=" & Trim(Str(rsIdAlmacen.Fields!idLaboraSubArea)) & " or "
                  rsIdAlmacen.MoveNext
               Loop
               lcEspecialidadesDelUsuario = Left(lcEspecialidadesDelUsuario, Len(lcEspecialidadesDelUsuario) - 4) & ")"
               Set mo_cmbIdResponsable.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(2)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
            Else
               Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghEspecialidadesEmergObs, ml_idUsuario)
               If rsIdAlmacen.RecordCount > 0 Then
                   lcEspecialidadesDelUsuario = " and ("
                   rsIdAlmacen.MoveFirst
                   Do While Not rsIdAlmacen.EOF
                      lcEspecialidadesDelUsuario = lcEspecialidadesDelUsuario & " dbo.Servicios.idEspecialidad=" & Trim(Str(rsIdAlmacen.Fields!idLaboraSubArea)) & " or "
                      rsIdAlmacen.MoveNext
                   Loop
                   lcEspecialidadesDelUsuario = Left(lcEspecialidadesDelUsuario, Len(lcEspecialidadesDelUsuario) - 4) & ")"
                   Set mo_cmbIdResponsable.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(4)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
               Else
                   Set mo_cmbIdResponsable.RowSource = mo_AdminAdmision.DevuelveServiciosDelHospital("(2,4)", lcEspecialidadesDelUsuario, sghFiltraAnuladosYactivos, sghPorDescTipoServicio)
               End If
            End If
        End If
    End Select
    If cmbIdResponsable.ListCount = 1 Then
       cmbIdResponsable.ListIndex = 0
    End If
    Set oBuscaDondeLabora = Nothing
    '
    'mo_Apariencia.ConfigurarFilasBiColores grdAdmision, sighentidades.GrillaConFilasBicolor
    On Error Resume Next
    
    Select Case ml_TipoFiltro
    Case sghFiltrarConsultaExterna
         txtNroHistoria.SetFocus
    Case Else
         txtNcuenta.SetFocus
    End Select
End Sub


Sub GenerarRecordsetTemporal()
    On Error GoTo ErrRsTmp
    With mrs_Tmp
          .Fields.Append "Dpto", adVarChar, 100, adFldIsNullable
          .Fields.Append "Especialidad", adVarChar, 100, adFldIsNullable
          .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
          .Fields.Append "Total", adInteger
          .Fields.Append "Llenas", adInteger
          .Fields.Append "Disponible", adInteger
          .Fields.Append "idServicio", adInteger
          .LockType = adLockOptimistic
          .Open
    End With
ErrRsTmp:
End Sub

Sub CargaDisponibilidadCamas()
    Dim oBuscar As New ADODB.Recordset
    Dim lcSql As String: Dim lcDpto As String: Dim lcEspecialidad As String
    Dim lcServicio As String
    Dim IdServicio As Integer
    Dim lnDisp As Integer: Dim lnLlena As Integer
    On Error GoTo ErrorCargaCama
    If mrs_Tmp.State = 1 Then
        If mrs_Tmp.RecordCount > 0 Then
           mrs_Tmp.MoveFirst
           Do While Not mrs_Tmp.EOF
              mrs_Tmp.Delete
              mrs_Tmp.Update
              mrs_Tmp.MoveNext
           Loop
        End If
    Else
        GenerarRecordsetTemporal
    End If
    Set oBuscar = mo_ReglasHoteleria.cargaCamasDisponiblesXtipoServicio(ml_TipoFiltro, Val(mo_cmbIdResponsable.BoundText))
    If oBuscar.RecordCount > 0 Then
       oBuscar.MoveFirst
       Do While Not oBuscar.EOF
          IdServicio = oBuscar.Fields!IdServicioUbicacionActual
          lcDpto = oBuscar.Fields!dDpto
          lcEspecialidad = oBuscar.Fields!despecialidad
          lcServicio = oBuscar.Fields!DServicio
          lnDisp = 0: lnLlena = 0
          Do While Not oBuscar.EOF And IdServicio = oBuscar.Fields!IdServicioUbicacionActual
             If oBuscar.Fields!IdEstadoCama = 1 Then
                lnDisp = lnDisp + 1
             Else
                lnLlena = lnLlena + 1
             End If
             oBuscar.MoveNext
             If oBuscar.EOF Then
                Exit Do
             End If
          Loop
          mrs_Tmp.AddNew
          mrs_Tmp.Fields!dpto = lcDpto
          mrs_Tmp.Fields!Especialidad = lcEspecialidad
          mrs_Tmp.Fields!Servicio = lcServicio
          mrs_Tmp.Fields!Total = lnDisp + lnLlena
          mrs_Tmp.Fields!Llenas = lnLlena
          mrs_Tmp.Fields!Disponible = lnDisp
          mrs_Tmp.Fields!IdServicio = IdServicio
          mrs_Tmp.Update
       Loop
    End If
    oBuscar.Close
    Set oBuscar = Nothing
    mrs_Tmp.Sort = "dpto,especialidad,servicio"
    Set grdCamasDisponibles.DataSource = mrs_Tmp
    mo_Apariencia.ConfigurarFilasBiColores grdCamasDisponibles, sighEntidades.GrillaConFilasBicolor
    ml_IdServicioConCamaDisponible = 0
    Exit Sub
ErrorCargaCama:
    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub FocusEnNroHistoria()
    txtNroHistoria.SetFocus
End Sub
Public Sub FocusEnGrilla()
    grdAdmision.SetFocus
End Sub


Sub InicilizarParametros()
    wxParametro289 = lcBuscaParametro.SeleccionaFilaParametro(289)
    wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    wxParametro541 = lcBuscaParametro.SeleccionaFilaParametro(541)
    ldFechaActualServidor = lcBuscaParametro.RetornaFechaServidorSQL
End Sub
