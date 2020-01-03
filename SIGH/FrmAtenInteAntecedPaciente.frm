VERSION 5.00
Begin VB.Form FrmAtenInteAntecedPaciente 
   Caption         =   "Antecedente de Paciente : Apellidos y Nombres"
   ClientHeight    =   6360
   ClientLeft      =   5295
   ClientTop       =   3060
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10725
   Begin SISGalenPlus.ucHCAntecedentes ucHCAntecedentes1 
      Height          =   5295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10455
      _ExtentX        =   20135
      _ExtentY        =   12726
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Grabar y Cerrar"
      DisabledPicture =   "FrmAtenInteAntecedPaciente.frx":0000
      DownPicture     =   "FrmAtenInteAntecedPaciente.frx":0460
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
      Left            =   4800
      Picture         =   "FrmAtenInteAntecedPaciente.frx":08D5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   1365
   End
End
Attribute VB_Name = "FrmAtenInteAntecedPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Atecedentes del Paciente
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim ml_idUsuario As Long
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim ml_idPaciente As Long
Dim ms_MensajeError As String
Dim mb_generoPlanIntegral As Boolean

Property Get GeneroPlanIntegral() As Boolean
   GeneroPlanIntegral = mb_generoPlanIntegral
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
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

Private Sub btnAceptar_Click()
    Call ucHCAntecedentes1.grabarAntecedentePaciente
    If ucHCAntecedentes1.SeGeneroPlanIntegral = False Then
        If ucHCAntecedentes1.MensajeError <> "" Then
            MsgBox ucHCAntecedentes1.MensajeError
        Else
            If MsgBox("¿Desea Generar Plan Integral Para el Paciente?", vbQuestion + vbYesNo, "Pregunta") = vbYes Then
                If GenerarPlanDeAtencionIntegral = True Then
                    mb_generoPlanIntegral = True
                    MsgBox "Plan de Atencion Integral Generado"
                Else
                    MsgBox "Error al Generar el Plan de Atención Integral : " & ms_MensajeError, vbExclamation, "Advertencia"
                End If
            End If
        End If
    Else
        
'    Else
'        Call GenerarPlanDeAtencionIntegral
    End If
    Unload Me
End Sub

Public Function GenerarPlanDeAtencionIntegral() As Boolean
    Dim oReglasAntecedentesPaciente As New ReglasAntecedentesPaciente
    Dim oReglasAtencionIntegral As New ReglasAtencionIntegral
    Dim oDOAtenIntePlanIntePaciente As New DOAtenIntePlanIntePaciente
    
    oDOAtenIntePlanIntePaciente.IdAtenInteGrupo = sighGrupoEdad.Nino
    oDOAtenIntePlanIntePaciente.idPaciente = ucHCAntecedentes1.idPaciente
    
    GenerarPlanDeAtencionIntegral = oReglasAtencionIntegral.GenerarPlanInmunizacionPaciente(oDOAtenIntePlanIntePaciente)
    If GenerarPlanDeAtencionIntegral = False Then
        ms_MensajeError = oReglasAtencionIntegral.MensajeError
    Else
        ms_MensajeError = ""
    End If
    
    GenerarPlanDeAtencionIntegral = oReglasAtencionIntegral.GenerarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente)
    If GenerarPlanDeAtencionIntegral = False Then
        ms_MensajeError = ms_MensajeError & vbCrLf & oReglasAtencionIntegral.MensajeError
    Else
        ms_MensajeError = ""
    End If
    
    GenerarPlanDeAtencionIntegral = oReglasAtencionIntegral.GenerarPlanCrecimientoPaciente(oDOAtenIntePlanIntePaciente)
    If GenerarPlanDeAtencionIntegral = False Then
        ms_MensajeError = ms_MensajeError & vbCrLf & oReglasAtencionIntegral.MensajeError
    Else
        ms_MensajeError = ""
    End If
    
    GenerarPlanDeAtencionIntegral = oReglasAtencionIntegral.GenerarPlanSuplementoPaciente(oDOAtenIntePlanIntePaciente)
    If GenerarPlanDeAtencionIntegral = False Then
        ms_MensajeError = ms_MensajeError & vbCrLf & oReglasAtencionIntegral.MensajeError
    Else
        ms_MensajeError = ""
    End If
    
    GenerarPlanDeAtencionIntegral = oReglasAtencionIntegral.GenerarPlanTamizajePaciente(oDOAtenIntePlanIntePaciente)
    If GenerarPlanDeAtencionIntegral = False Then
        ms_MensajeError = oReglasAtencionIntegral.MensajeError
    Else
        ms_MensajeError = ""
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
'    FServidor = lcBuscaParametro.RetornaFechaServidorSQL
'    Select Case mi_Opcion
'    Case sghAgregar
'        Me.Caption = "Agregar Triaje"
'    Case sghModificar
'        Me.Caption = "Modificar Triaje"
'    Case sghConsultar
'        Me.Caption = "Consultar Triaje"
'    Case sghEliminar
'        Me.Caption = "Eliminar Triaje"
'    End Select
    
'    Me.ucHCAntecedentes1.idPaciente = ml_idUsuario
    Me.Left = Val((Screen.Width - Me.ScaleWidth) / 2)
    Me.Top = Val((Screen.Height - Me.ScaleHeight) / 2)
    
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
    Select Case KeyCode
       'Case vbKeyEscape
       '    btnCancelar_Click
       Case vbKeyF2
            btnAceptar_Click
    End Select
End Sub

Private Sub ucHCAntecedentes1_SePresionoTeclaEspecial(KeyCode As Integer)
    AdministrarKeyPreview KeyCode
End Sub
