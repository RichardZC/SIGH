VERSION 5.00
Begin VB.Form HerrModificaPacienteAtencionHE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza el Paciente NN de una Atención "
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "HerrModificaPacienteAtencionHE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Paciente que se asignará la Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   60
      TabIndex        =   13
      Top             =   3600
      Width           =   4695
      Begin VB.TextBox txtNombrePacienteNew 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   4455
      End
      Begin VB.TextBox txtNhistoria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   5
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Historia"
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
         Left            =   180
         TabIndex        =   14
         Top             =   345
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuenta Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   60
      TabIndex        =   11
      Top             =   1650
      Width           =   4695
      Begin VB.TextBox txtPlan 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   1380
         Width           =   4515
      End
      Begin VB.TextBox txtNcuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         MaxLength       =   30
         TabIndex        =   1
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtNombrePaciente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4515
      End
      Begin VB.TextBox txtDatosDeCuenta 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   990
         Width           =   4515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "N° Cuenta"
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
         Left            =   150
         TabIndex        =   12
         Top             =   255
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   4710
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   1110
         Left            =   90
         TabIndex        =   0
         Top             =   240
         Width           =   4425
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Top             =   4800
      Width           =   4725
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "HerrModificaPacienteAtencionHE.frx":0CCA
         DownPicture     =   "HerrModificaPacienteAtencionHE.frx":112A
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
         Left            =   968
         Picture         =   "HerrModificaPacienteAtencionHE.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrModificaPacienteAtencionHE.frx":1A14
         DownPicture     =   "HerrModificaPacienteAtencionHE.frx":1ED8
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
         Left            =   2468
         Picture         =   "HerrModificaPacienteAtencionHE.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   225
         Width           =   1335
      End
   End
End
Attribute VB_Name = "HerrModificaPacienteAtencionHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Actualiza CUENTA en otra Historia
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_AdminAdmision As New ReglasAdmision
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim ml_idPaciente As Long
Dim ml_IdPacienteNew As Long
Dim lnIdAtencion As Long
Dim lnIdCama As Long
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_IdUsuario As Long
Dim mo_lcNombrePc  As String


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property

Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property


Private Sub btnAceptar_Click()
    If ml_idPaciente = 0 Then
       MsgBox "Tiene que ingresar el N° Cuenta", vbInformation, "Mensaje"
       Exit Sub
    End If
    If ml_IdPacienteNew = 0 Then
       MsgBox "Tiene que ingresar el N° de Historia Clínica", vbInformation, "Mensaje"
       Exit Sub
    End If
    If ml_idPaciente = ml_IdPacienteNew Then
       MsgBox "El  Paciente Actual y el Nuevo no pueden ser el mismo", vbInformation, "mensaje"
       Exit Sub
    End If
    If MsgBox("Esta seguro", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        Me.MousePointer = 1
        mo_reglasComunes.ActualizaIdPacienteEnTodasLasTablasSegunNroCuenta ml_IdPacienteNew, Val(txtNcuenta.Text), _
                                 lnIdAtencion, lnIdCama, ml_IdUsuario, mo_lcNombrePc, _
                                 "Cambia Pac.NN: " & Trim(txtNombrePaciente.Text) & " por HC: " & txtNhistoria.Text
        Me.MousePointer = 11
        Me.Visible = False
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
  mo_reglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrModificaPacienteAtencionHE"
  mo_Formulario.HabilitarDeshabilitar Me.txtNombrePaciente, False
  mo_Formulario.HabilitarDeshabilitar Me.txtDatosDeCuenta, False
  mo_Formulario.HabilitarDeshabilitar Me.txtPlan, False
  mo_Formulario.HabilitarDeshabilitar Me.txtNombrePacienteNew, False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub txtNcuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNcuenta

End Sub


Private Sub txtNcuenta_LostFocus()
   If Val(txtNcuenta.Text) > 0 Then
       Dim oRsTmp As New Recordset
       Dim lbSigue As Boolean
       Dim oConexion As New Connection
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.CursorLocation = adUseClient
       Set oRsTmp = mo_ReglasFarmacia.AtencionesSelecionarPorCuenta(txtNcuenta.Text, oConexion)
       txtDatosDeCuenta.Text = ""
       txtPlan.Text = ""
       txtNombrePaciente.Text = ""
       ml_idPaciente = 0
       lnIdAtencion = 0
       lnIdCama = 0
       If oRsTmp.RecordCount > 0 Then
            If oRsTmp.Fields!IdEstado <> 1 Then
               MsgBox "Esa cuenta no se  encuentra ABIERTA", vbInformation, "Mensaje"
            Else
                txtDatosDeCuenta.Text = "F.Ing: " & oRsTmp.Fields!FechaIngreso & " - " & IIf(oRsTmp.Fields!idTipoServicio = 1, "Consultorios Externos", IIf(oRsTmp.Fields!idTipoServicio = 3, "Hospitalización", "Emergencia")) & " - (Est: " & Trim(oRsTmp.Fields!estadoCta) & ")"
                txtPlan.Text = "IAFA Act.: " & oRsTmp.Fields!dFuenteFinanciamiento
                txtNombrePaciente.Text = oRsTmp.Fields!NroHistoriaClinica & " " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & oRsTmp.Fields!PrimerNombre
                ml_idPaciente = oRsTmp.Fields!idPaciente
                lnIdAtencion = oRsTmp.Fields!idAtencion
                lnIdCama = IIf(IsNull(oRsTmp.Fields!IdCamaEgreso), 0, oRsTmp.Fields!IdCamaEgreso)
            End If
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       oConexion.Close
       Set oConexion = Nothing
   End If
End Sub


Private Sub txtNhistoria_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtNhistoria

End Sub

Private Sub txtNhistoria_LostFocus()
      If txtNhistoria.Text <> "" Then
        Dim oRsTmp1 As New ADODB.Recordset
        Dim oDOPaciente As New SIGHComun.doPaciente
        oDOPaciente.NroHistoriaClinica = HCigualDNI_AgregaNUEVEaLaHistoria(txtNhistoria.Text)
        Set oRsTmp1 = mo_AdminAdmision.PacientesFiltrar(oDOPaciente, False, False, "")
        If oRsTmp1.RecordCount > 0 Then
           ml_IdPacienteNew = oRsTmp1.Fields!idPaciente
           txtNombrePacienteNew.Text = Trim(oRsTmp1.Fields!ApellidoPaterno) & " " & Trim(oRsTmp1.Fields!ApellidoMaterno) & " " & oRsTmp1.Fields!PrimerNombre
        Else
           ml_IdPacienteNew = 0
           txtNombrePacienteNew.Text = ""
        End If
        Set oRsTmp1 = Nothing
        Set oDOPaciente = Nothing
      End If
End Sub
