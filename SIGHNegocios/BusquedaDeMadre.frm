VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form BusquedaDeMadre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de la Madre para un Recién Nacido"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   Icon            =   "BusquedaDeMadre.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   90
      TabIndex        =   4
      Top             =   2880
      Width           =   8355
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "BusquedaDeMadre.frx":000C
         DownPicture     =   "BusquedaDeMadre.frx":04D0
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
         Left            =   3585
         Picture         =   "BusquedaDeMadre.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   8400
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
         Left            =   1620
         TabIndex        =   7
         Top             =   360
         Width           =   1515
      End
      Begin VB.CommandButton cmdBuscaXapell 
         Caption         =   "..."
         Height          =   315
         Left            =   3210
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtMadre 
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
         Height          =   315
         Left            =   3570
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   4695
      End
      Begin UltraGrid.SSUltraGrid grdNacimientos 
         Height          =   1665
         Left            =   60
         TabIndex        =   6
         Top             =   840
         Width           =   8220
         _ExtentX        =   14499
         _ExtentY        =   2937
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
         Caption         =   "Lista de Nacimientos"
      End
      Begin VB.Label Label2 
         Caption         =   "....Elija y pulse Doble Clic de la lista de Nacimientos....(la madre debe tener EGRESO MEDICO)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   2520
         Width           =   8025
      End
      Begin VB.Label Label1 
         Caption         =   "DNI de la madre"
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
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   1365
      End
   End
End
Attribute VB_Name = "BusquedaDeMadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Busca la madre por DNI
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdPacienteMadre As Long
Dim ml_idAtencionMadre As Long
Dim ml_IdNacimientoSeleccionado As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim ml_texto As String
Dim oRsNacimientos As New Recordset
Dim mo_Apariencia As New sighentidades.GridInfragistic
Dim mo_ReglasAdmision As New ReglasAdmision

Property Get IdNacimientoSeleccionado() As Long
    IdNacimientoSeleccionado = ml_IdNacimientoSeleccionado
End Property
Property Let IdAtencionMadreSeleccionado(lValue As Long)
    ml_idAtencionMadre = lValue
End Property
Property Get IdAtencionMadreSeleccionado() As Long
    IdAtencionMadreSeleccionado = ml_idAtencionMadre
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property


Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False
End Sub

Private Sub btnCancelar_Click()
        mi_BotonPresionado = sghCancelar
        Me.Visible = False
End Sub

Private Sub cmdBuscaXapell_Click()
    Dim oBusqueda As New SIGHNegocios.BuscaPacientes
    Dim oDoPaciente As New doPaciente
    Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
    Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
    Dim oConexion As New Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oBusqueda.TipoFiltro = sghFiltrarTodos
    oBusqueda.MostrarFormulario
    If oBusqueda.BotonPresionado = sghAceptar Then
        Set oDoPaciente = mo_AdminAdmision.PacientesSeleccionarPorId(oBusqueda.IdRegistroSeleccionado, oConexion)
        ml_IdPacienteMadre = 0
        txtMadre.Text = ""
        If Not oDoPaciente Is Nothing Then
            ml_IdPacienteMadre = oDoPaciente.idPaciente
            txtMadre.Text = oDoPaciente.NroHistoriaClinica & " " & Trim(oDoPaciente.ApellidoPaterno) + " " + Trim(oDoPaciente.ApellidoMaterno) + " " + oDoPaciente.PrimerNombre
        End If
        CargaNacimientos
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oBusqueda = Nothing
    Set oDoPaciente = Nothing
    Set mo_ReglasFarmacia = Nothing
    Set mo_AdminAdmision = Nothing
End Sub

Private Sub Form_Load()
    ml_idAtencionMadre = 0
    ml_IdNacimientoSeleccionado = 0
    mo_Apariencia.ConfigurarFilasBiColores Me.grdNacimientos, sighentidades.GrillaConFilasBicolor
End Sub



Private Sub grdNacimientos_DblClick()
    On Error GoTo errNac
    Dim oRsTmp99 As New Recordset
    Set oRsTmp99 = grdNacimientos
    ml_idAtencionMadre = oRsTmp99!idAtencion
    ml_IdNacimientoSeleccionado = oRsTmp99!idNacimiento
    oRsTmp99.Close
    Set oRsTmp99 = Nothing
    btnAceptar_Click
errNac:
End Sub

Private Sub grdNacimientos_KeyPress(KeyAscii As UltraGrid.SSReturnShort)
    If KeyAscii = 13 Then
       grdNacimientos_DblClick
    End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(txtDNI.Text) = 8 Then
       Dim oRsTmp As New Recordset
       Dim mo_ReglasAdmision As New ReglasAdmision
       Dim oConexion As New Connection
            oConexion.CursorLocation = adUseClient
            oConexion.CommandTimeout = 300
            oConexion.Open sighentidades.CadenaConexion
       Set oRsTmp = mo_ReglasAdmision.PacientesXdni(Trim(txtDNI.Text), oConexion)
       ml_IdPacienteMadre = 0
       txtMadre.Text = ""
       If oRsTmp.RecordCount > 0 Then
          ml_IdPacienteMadre = oRsTmp.Fields!idPaciente
          txtMadre.Text = Trim(oRsTmp.Fields!NroHistoriaClinica) & " " & Trim(oRsTmp.Fields!ApellidoPaterno) & " " & Trim(oRsTmp.Fields!ApellidoMaterno) & " " & Trim(oRsTmp.Fields!PrimerNombre)
       End If
       oRsTmp.Close
       Set oRsTmp = Nothing
       CargaNacimientos
    End If
End Sub


Sub CargaNacimientos()
    On Error GoTo errNac
    Set grdNacimientos.DataSource = Nothing
    If ml_IdPacienteMadre > 0 Then
        Set grdNacimientos.DataSource = mo_ReglasAdmision.AtencionesNacimientosXidPacienteDeMadre(ml_IdPacienteMadre)
    End If
    Exit Sub
errNac:
    If Err.Number = 3705 Then
       oRsNacimientos.Close
       Resume
    End If
End Sub
