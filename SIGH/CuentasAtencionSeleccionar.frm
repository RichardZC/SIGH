VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form CuentasAtencionSeleccionar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar la cuenta de atención"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "CuentasAtencionSeleccionar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   8265
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CuentasAtencionSeleccionar.frx":0CCA
         DownPicture     =   "CuentasAtencionSeleccionar.frx":112A
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
         Left            =   2520
         Picture         =   "CuentasAtencionSeleccionar.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CuentasAtencionSeleccionar.frx":1A14
         DownPicture     =   "CuentasAtencionSeleccionar.frx":1ED8
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
         Left            =   4050
         Picture         =   "CuentasAtencionSeleccionar.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid grdCajeros 
      Height          =   4470
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   7885
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
      Caption         =   "Lista de Cuentas de Atención"
   End
End
Attribute VB_Name = "CuentasAtencionSeleccionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim ml_IdRegistroSeleccionado As Long
Dim ml_TipoFiltro As sghTipoFiltroPacientes
Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Apariencia As New SIGHComun.GridInfragistic
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Property Set DataSource(oValue As ADODB.Recordset)
    Set Me.grdCajeros.DataSource = oValue
    mo_Apariencia.ConfigurarFilasBiColores grdCajeros, SIGHComun.GrillaConFilasBicolor
End Property
Property Get DataSource() As ADODB.Recordset
    Set DataSource = Me.grdCajeros.DataSource
End Property
Property Let IdRegistroSeleccionado(lValue As Long)
    ml_IdRegistroSeleccionado = lValue
End Property
Property Get IdRegistroSeleccionado() As Long
    IdRegistroSeleccionado = ml_IdRegistroSeleccionado
End Property


Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    Me.Visible = False

End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    Me.Visible = False
End Sub

Private Sub Form_Load()
    mi_BotonPresionado = sghBotonDetallePresionado.sghCancelar
End Sub

Private Sub grdCajeros_AfterRowActivate()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCajeros.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCuentaAtencion")

End Sub

Private Sub grdCajeros_Click()
Dim rsRecordset As ADODB.Recordset

    Set rsRecordset = grdCajeros.DataSource
    On Error Resume Next
    ml_IdRegistroSeleccionado = rsRecordset("IdCuentaAtencion")
    
End Sub


Private Sub grdCajeros_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    
    grdCajeros.Bands(0).Columns("IdCuentaAtencion").Header.Caption = "Cuenta"
    grdCajeros.Bands(0).Columns("IdCuentaAtencion").Width = 700
    
    grdCajeros.Bands(0).Columns("FechaApertura").Header.Caption = "Fecha Apertura"
    grdCajeros.Bands(0).Columns("FechaApertura").Width = 1500
    
    grdCajeros.Bands(0).Columns("HoraApertura").Header.Caption = "Hora Apertura"
    grdCajeros.Bands(0).Columns("HoraApertura").Width = 1000
    
    grdCajeros.Bands(0).Columns("ApellidoPaterno").Header.Caption = "Ap.Materno"
    grdCajeros.Bands(0).Columns("ApellidoPaterno").Width = 1500
    
    grdCajeros.Bands(0).Columns("ApellidoMaterno").Header.Caption = "Ap.Materno"
    grdCajeros.Bands(0).Columns("ApellidoMaterno").Width = 1500
    
    grdCajeros.Bands(0).Columns("PrimerNombre").Header.Caption = "Primer Nombre"
    grdCajeros.Bands(0).Columns("PrimerNombre").Width = 1500
    
    grdCajeros.Bands(0).Columns("SegundoNombre").Header.Caption = "Segundo Nombre"
    grdCajeros.Bands(0).Columns("SegundoNombre").Width = 1500
    

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
    
    Select Case KeyCode
    Case vbKeyEscape
        btnCancelar_Click
    Case vbKeyF2
        btnAceptar_Click
    Case vbKeyF3
     Case vbKeyF4
     Case vbKeyF5
     Case vbKeyF6
     Case vbKeyF7
     Case vbKeyF8
    End Select
End Sub











