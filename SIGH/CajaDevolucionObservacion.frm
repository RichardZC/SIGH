VERSION 5.00
Begin VB.Form PacienteNuevaHistoria 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CajaDevolucionObservacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5385
      Begin VB.TextBox txtNewHistoria 
         Height          =   360
         Left            =   1365
         TabIndex        =   4
         Top             =   195
         Width           =   2490
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Historia"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   5385
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CajaDevolucionObservacion.frx":000C
         DownPicture     =   "CajaDevolucionObservacion.frx":04D0
         Height          =   645
         Left            =   2910
         Picture         =   "CajaDevolucionObservacion.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CajaDevolucionObservacion.frx":0EA8
         DownPicture     =   "CajaDevolucionObservacion.frx":1308
         Height          =   645
         Left            =   1380
         Picture         =   "CajaDevolucionObservacion.frx":177D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "PacienteNuevaHistoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Actualiza clave del Usuario activo
'        Programado por: Barrantes D
'        Fecha: Octubre 2018
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdPaciente As Long
Dim ml_NroHistoriaClinica As Long
Dim ml_DNI As String
Dim ml_idTipoNumeracion As Long
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mi_BotonPresionado As sghBotonDetallePresionado
Property Get BotonPresionado() As sghBotonDetallePresionado
  BotonPresionado = mi_BotonPresionado
End Property
Property Let NewHistoria(lValue As String)
    If lValue <> "" Then
       txtNewHistoria.Text = lValue
       ml_DNI = lValue
    End If
End Property

Property Let idPaciente(lValue As Long)
   ml_IdPaciente = lValue
End Property
Property Let NroHistoriaClinica(lValue As Long)
   ml_NroHistoriaClinica = lValue
End Property
Property Let idTipoNumeracion(lValue As Long)
   ml_idTipoNumeracion = lValue
End Property
Private Sub btnAceptar_Click()
    If txtNewHistoria.Text = "" Then
       MsgBox "Tiene que ingresar el N° Historia Clínica NUEVA", vbInformation, ""
       Exit Sub
    End If
    If Val(txtNewHistoria.Text) <> ml_NroHistoriaClinica Then
       Dim oRsTmp1 As New Recordset
       Dim lbActualizaDNI As Boolean
       If txtNewHistoria.Text = ml_DNI Then
      'GLCC 02/11/20 CAMBIO36 INICIO
       'Quita wxNueve & para que no anteponga el numero 9 a la historia clinica
          'Set oRsTmp1 = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(wxNueve & txtNewHistoria.Text, 2)
           Set oRsTmp1 = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(txtNewHistoria.Text, 2)
      'GLCC 02/11/20 CAMBIO36 FIN
       Else
          Set oRsTmp1 = mo_ReglasAdmision.PacientesXnroHistoriaTipoNumeracion(Val(txtNewHistoria.Text), ml_idTipoNumeracion)
       End If
       If oRsTmp1.RecordCount > 0 Then
          MsgBox "Esa HISTORIA nueva ya existe para : " & oRsTmp1!ApellidoPaterno & " " & oRsTmp1!ApellidoMaterno & " " & _
                 oRsTmp1!PrimerNombre, vbInformation, ""
       Else
          If ml_idTipoNumeracion > 3 Then
            '*** pasa temporal al HC final
            Dim mo_HistoriasClinicas As New DOHistoriaClinica
            With mo_HistoriasClinicas
                .idPaciente = ml_IdPaciente
                .IdEstadoHistoria = 1
                .idTipoHistoria = 1
               ' .FechaPasoAPasivo = IIf(Me.txtFechaPasoAPasivo.Text = sighentidades.FECHA_VACIA_DMY, 0, Me.txtFechaPasoAPasivo.Text)
                .fechacreacion = Now
               'GLCC 02/11/20 CAMBIO36 INICIO
                'Quitar la palabra wxNueve & para que no anteponga el numero 9
               '.NroHistoriaClinica = IIf(txtNewHistoria.Text = ml_DNI, wxNueve & txtNewHistoria.Text, txtNewHistoria.Text)
                .NroHistoriaClinica = IIf(txtNewHistoria.Text = ml_DNI, txtNewHistoria.Text, txtNewHistoria.Text)
               'GLCC 02/11/20 CAMBIO36 FIN
                .IdUsuarioAuditoria = sighEntidades.Usuario
                .idTipoNumeracion = 2
               ' .IdTipoNumeracionAnterior = Val(mo_cmbIdTipoNumeracionHistoriaAnt.BoundText)
                '.NroHistoriaClinicaAnterior = Me.txtIdHistoriaClinicaAnt
                '.FechaUltimoMovimiento = Date
            End With
            If mo_AdminArchivoClinico.HistoriaClinicaAgregar(mo_HistoriasClinicas, sghOpcionGalenHos.sghPacientes, "", _
                   "N° temporal: " & ml_NroHistoriaClinica) Then
            End If
            Set mo_HistoriasClinicas = Nothing
          End If
          If txtNewHistoria.Text <> ml_DNI Then
            Dim oConexion As New Connection
            Dim oConexionExterna As New Connection
            oConexion.CommandTimeout = 900
            oConexion.CursorLocation = adUseClient
            oConexion.Open sighEntidades.CadenaConexion
            oConexionExterna.CommandTimeout = 900
            oConexionExterna.CursorLocation = adUseClient
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
             mo_ReglasAdmision.PasaNuevaHC Trim(Str(ml_NroHistoriaClinica)), txtNewHistoria.Text, oConexion, oConexionExterna
             mi_BotonPresionado = sghAceptar
            oConexion.Close
            oConexionExterna.Close
            Set oConexion = Nothing
            Set oConexionExterna = Nothing
          Else
             mo_ReglasAdmision.PacientesActualizarDNI ml_DNI, ml_IdPaciente, 2
             mo_ReglasAdmision.ActualizaHistoriaIgualDNI Trim(Str(ml_IdPaciente))
             mi_BotonPresionado = sghAceptar
          End If
       End If
       oRsTmp1.Close
       Set oRsTmp1 = Nothing
       Me.Visible = False
    End If
End Sub
Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub
