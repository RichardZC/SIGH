VERSION 5.00
Begin VB.Form MedicosBusqueda 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   Icon            =   "MedicosBusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin SIGHNegocios.ucMedicosLista ucMedicosLista1 
      Height          =   5235
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   9234
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   30
      TabIndex        =   2
      Top             =   5280
      Width           =   10245
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "MedicosBusqueda.frx":0CCA
         DownPicture     =   "MedicosBusqueda.frx":118E
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
         Left            =   5115
         Picture         =   "MedicosBusqueda.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "MedicosBusqueda.frx":1B66
         DownPicture     =   "MedicosBusqueda.frx":1FC6
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
         Left            =   3540
         Picture         =   "MedicosBusqueda.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "MedicosBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mb_Loading As Boolean






Property Let IdEspecialidad(lValue As Long)
    ucMedicosLista1.IdEspecialidad = lValue
End Property

Property Let HoraProgramada(lValue As String)
    ucMedicosLista1.HoraProgramada = lValue
End Property
Property Let FechaProgramada(lValue As Date)
    ucMedicosLista1.FechaProgramada = lValue
End Property

Property Let idTipoServicio(lValue As Long)
    ucMedicosLista1.idTipoServicio = lValue
End Property

Property Let NombreMedico(lValue As String)
    ucMedicosLista1.NombreMedico = lValue
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set ucMedicosLista1.DataSource = oValue
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = ucMedicosLista1.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ucMedicosLista1.IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ucMedicosLista1.IdRegistroSeleccionado
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

Private Sub Form_Activate()
    If mb_Loading Then
        If ucMedicosLista1.IdEspecialidad <> 0 Then
            ucMedicosLista1.RealizarBusqueda
            On Error Resume Next
            Me.ucMedicosLista1.SetFocus
        End If
        mb_Loading = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Load()
    Me.ucMedicosLista1.Titulo = "Búsqueda de médicos"
    mb_Loading = True
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
            ucMedicosLista1.RealizarBusqueda
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub


Private Sub ucMedicosLista1_SeleccionaRegistro(lnIdRegistroSeleccionado As Long)
     If lnIdRegistroSeleccionado > 0 Then
        IdRegistroSeleccionado = lnIdRegistroSeleccionado
        btnAceptar_Click
     End If
End Sub
