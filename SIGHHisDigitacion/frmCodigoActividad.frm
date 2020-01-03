VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmCodigoActividad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de códigos de actividades"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
   Icon            =   "frmCodigoActividad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigo 
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
      Left            =   30
      MaxLength       =   6
      TabIndex        =   0
      Top             =   885
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcioAct 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   885
      Width           =   6015
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleCodigosActividades 
      Height          =   2535
      Left            =   30
      TabIndex        =   2
      Top             =   1230
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4471
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ugvDetalleCodigosActividades"
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   30
      TabIndex        =   4
      Top             =   3720
      Width           =   7485
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmCodigoActividad.frx":000C
         DownPicture     =   "frmCodigoActividad.frx":046C
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2310
         Picture         =   "frmCodigoActividad.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmCodigoActividad.frx":0D56
         DownPicture     =   "frmCodigoActividad.frx":121A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3870
         Picture         =   "frmCodigoActividad.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   645
      Width           =   675
   End
   Begin VB.Label Label 
      Caption         =   "Descripcion de Codigos de  Actividades"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   7
      Top             =   645
      Width           =   3555
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Códigos de Actividades"
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
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmCodigoActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de Listado de Codigos de Actividades.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_CodigoTipoActividad As Integer                   'Repsenta la descripcion del Codigo del HIS
Dim ml_IdCodigoSeleccionado As Long                     'Representa el
Dim ms_CodigoDevuelto As String
Dim oRcs_DetalleCodigosActividades As New Recordset     'Representa el detalle de codigo HIS de la base de datos
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim mo_Teclado As New SIGHEntidades.Teclado

Property Let IdTipoActividad(lValue As Integer)
   ml_CodigoTipoActividad = lValue
End Property
Property Get IdTipoActividad() As Integer
   IdTipoActividad = ml_CodigoTipoActividad
End Property

Property Let CodigoSeleccionado(sValue As String)
   ms_CodigoDevuelto = sValue
End Property
Property Get CodigoSeleccionado() As String
   CodigoSeleccionado = ms_CodigoDevuelto
End Property

Property Get IdCodigoSeleccionado() As Long
   IdCodigoSeleccionado = ml_IdCodigoSeleccionado
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

Private Sub Form_Load()
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigosActividades, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleCodigosActividades = mr_ReglasHIS.ObtenerListaCodigosActividadesporCodigoyNombre("", "")
    Set Me.ugvDetalleCodigosActividades.DataSource = oRcs_DetalleCodigosActividades
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim orsTemp As Recordset
    If KeyCode = 13 Then
        Set orsTemp = Me.ugvDetalleCodigosActividades.DataSource
        If orsTemp.RecordCount = 1 Then
            btnAceptar_Click
        Else
            If orsTemp.RecordCount = 0 Then
               mo_Teclado.RealizarNavegacion KeyCode, txtDescripcioAct
               AdministrarKeyPreview CInt(KeyCode)
            Else
                ugvDetalleCodigosActividades.SetFocus
            End If
        End If
    Else
        mo_Teclado.RealizarNavegacion KeyCode, Me.txtDescripcioAct
        AdministrarKeyPreview KeyCode
    End If
End Sub

Private Sub txtDescripcioAct_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, ugvDetalleCodigosActividades
    AdministrarKeyPreview KeyCode
End Sub

'MODIFICADO POR YEPE NOVIEMBRE
Private Sub txtDescripcioAct_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigosActividades, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleCodigosActividades = mr_ReglasHIS.ObtenerListaCodigosActividadesporCodigoyNombre(Me.txtCodigo.Text, Me.txtDescripcioAct.Text)
    Set ugvDetalleCodigosActividades.DataSource = oRcs_DetalleCodigosActividades
End Sub
'MODIFICADO POR EYPE NOVIEMBRE
Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigosActividades, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleCodigosActividades = mr_ReglasHIS.ObtenerListaCodigosActividadesporCodigoyNombre(Me.txtCodigo.Text, Me.txtDescripcioAct.Text)
    Set ugvDetalleCodigosActividades.DataSource = oRcs_DetalleCodigosActividades
End Sub

Private Sub ugvDetalleCodigosActividades_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleCodigosActividades_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
'Layout.Override.AllowDelete = ssAllowDeleteNo
'oRcs_DetalleCodigosActividades
With Me.ugvDetalleCodigosActividades.Bands(0)
    .Columns("IdHisCodActvidad").Hidden = True
    .Columns("IdTipoAtencion").Hidden = True
    .Columns("CodigoActividad").Header.Caption = "Codigo"
    .Columns("CodigoActividad").Width = 1000
    .Columns("Descripcion").Header.Caption = "Descripcion"
    .Columns("Descripcion").Width = 5500
End With
End Sub


Private Sub ugvDetalleCodigosActividades_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF2 Then
        btnAceptar_Click
    Else
        AdministrarKeyPreview CInt(KeyCode)
    End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    ml_IdCodigoSeleccionado = CLng(Me.ugvDetalleCodigosActividades.ActiveRow.Cells("IdHisCodActvidad").Value)
    ms_CodigoDevuelto = CStr(Me.ugvDetalleCodigosActividades.ActiveRow.Cells("CodigoActividad").Value)
    ml_CodigoTipoActividad = CStr(Me.ugvDetalleCodigosActividades.ActiveRow.Cells("IdTipoAtencion").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ms_CodigoDevuelto = ""
    Visible = False
End Sub

Public Sub MostrarFormulario()
Me.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

