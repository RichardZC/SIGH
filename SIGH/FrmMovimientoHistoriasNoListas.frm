VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form MovimientoHistoriasNoListas 
   Caption         =   "Historias Clinicas"
   ClientHeight    =   4650
   ClientLeft      =   10815
   ClientTop       =   5430
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8235
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton btnNo 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton btnSI 
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   4080
      Width           =   1095
   End
   Begin UltraGrid.SSUltraGrid grdHistoriasSeleccionadas 
      Height          =   3195
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   5636
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Historias seleccionadas"
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Listado de Historias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7995
   End
End
Attribute VB_Name = "MovimientoHistoriasNoListas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRsHistorias As ADODB.Recordset
Dim mb_Respuesta As Boolean
Dim mo_Apariencia As New sighEntidades.GridInfragistic

Property Set RsHistorias(lValue As ADODB.Recordset)
   Set oRsHistorias = lValue
   Set grdHistoriasSeleccionadas.DataSource = lValue
End Property

Property Get Respuesta() As Boolean
   Respuesta = mb_Respuesta
End Property

Private Sub btnAceptar_Click()
    Me.Visible = False
End Sub

Private Sub btnNo_Click()
    mb_Respuesta = False
    Me.Visible = False
End Sub

Private Sub btnSI_Click()
    mb_Respuesta = True
    Me.Visible = False
End Sub

Private Sub grdHistoriasSeleccionadas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.AllowAddNew = ssAllowAddNewNo
    Layout.Override.AllowDelete = ssAllowDeleteNo
    Layout.Override.AllowUpdate = ssAllowUpdateNo
    
    grdHistoriasSeleccionadas.Bands(0).Columns("IdHistoriaSolicitada").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("EsServicioCostoCero").Hidden = True
    grdHistoriasSeleccionadas.Bands(0).Columns("SeDaraSalida").Hidden = True
        
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Header.Caption = "Nro Historia"
    grdHistoriasSeleccionadas.Bands(0).Columns("HistoriaClinica").Width = 1000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Header.Caption = "Nombres"
    grdHistoriasSeleccionadas.Bands(0).Columns("Nombres").Width = 2000
          
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Header.Caption = "Servicio Origen"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioOrigen").Width = 1500
    
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Header.Caption = "Servicio Destino"
    grdHistoriasSeleccionadas.Bands(0).Columns("NombreServicioDestino").Width = 2000
    
    grdHistoriasSeleccionadas.Bands(0).Columns("Motivo").Header.Caption = "N°Folios"
    grdHistoriasSeleccionadas.Bands(0).Columns("Motivo").Width = 1000
    
    mo_Apariencia.ConfigurarFilasBiColores grdHistoriasSeleccionadas, sighEntidades.GrillaConFilasBicolor
End Sub


Public Function mostrarBotonesSiNO()
    btnSI.Visible = True
    btnNo.Visible = True
    btnAceptar.Visible = False
End Function

Public Function mostrarBotonesSoloAceptar()
    btnSI.Visible = False
    btnNo.Visible = False
    btnAceptar.Visible = True
End Function
