VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleProductosHis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Código productos His"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   Icon            =   "frmDetalleProductosHis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescripcion 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   5415
   End
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
      Left            =   120
      MaxLength       =   10
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6975
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmDetalleProductosHis.frx":000C
         DownPicture     =   "frmDetalleProductosHis.frx":046C
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
         Left            =   2280
         Picture         =   "frmDetalleProductosHis.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleProductosHis.frx":0D56
         DownPicture     =   "frmDetalleProductosHis.frx":121A
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
         Left            =   3720
         Picture         =   "frmDetalleProductosHis.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleCodigoProductosHis 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6975
      _ExtentX        =   12303
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
      Caption         =   "ugvDetalleCodigoProductosHis"
   End
   Begin VB.Label Label 
      Caption         =   "Nombre Productos His"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   600
      Width           =   2235
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
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   675
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Códigos Productos His"
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
      TabIndex        =   3
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "frmDetalleProductosHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Detalle de productos HIS
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ms_IdDiagCpt As Integer
Dim ms_CodigoDx As String
Dim ms_descripciondiagcpt As String
Dim ms_MasDeUnDiagnosticos As Integer
Dim mo_Teclado As New SIGHEntidades.Teclado
'Repsenta la descripcion del Codigo del HIS
Dim oRcs_DetalleProductosHis As New Recordset                 'Representa el detalle de codigo HIS de la base de datos
'Representa el filtro de codigo HIS
Dim mi_BotonPresionado As sghBotonDetallePresionado

Property Let IdDiagCpt(sValue As Long)
   ms_IdDiagCpt = sValue
End Property
Property Get IdDiagCpt() As Long
   IdDiagCpt = ms_IdDiagCpt
End Property
Property Let CodigoDx(sValue As String)
   ms_CodigoDx = sValue
End Property
Property Get CodigoDx() As String
   CodigoDx = ms_CodigoDx
End Property

Property Let descripciondiagcpt(sValue As String)
   ms_descripciondiagcpt = sValue
End Property
Property Get descripciondiagcpt() As String
   descripciondiagcpt = ms_descripciondiagcpt
End Property

Property Let MasDeUnDiagnosticos(sValue As Integer)
   ms_MasDeUnDiagnosticos = sValue
End Property
Property Get MasDeUnDiagnosticos() As Integer
   MasDeUnDiagnosticos = ms_MasDeUnDiagnosticos
End Property

Property Get BotonPresionado() As sghBotonDetallePresionado
    BotonPresionado = mi_BotonPresionado
End Property

'MODIFICADO POR YEPE NOVIEMBRE
Private Sub Form_Load()
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoProductosHis, SIGHEntidades.GrillaConFilasBicolor
    'Set oRcs_DetalleLABHIS = mr_ReglasHIS.ObtenerListaCodigosLAB '("","")
    Me.txtCodigo.Text = IIf(ms_CodigoDx = "...", "", ms_CodigoDx)
    Set oRcs_DetalleProductosHis = mr_ReglasHIS.ObtenerListaCodigosProductosHisPorNombreYDescripcion(Me.txtCodigo.Text, "")
    Set ugvDetalleCodigoProductosHis.DataSource = oRcs_DetalleProductosHis
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim orsTemp As Recordset
    If KeyCode = 13 Then
        Set orsTemp = ugvDetalleCodigoProductosHis.DataSource
        If orsTemp.RecordCount = 1 Then
            btnAceptar_Click
        Else
            If orsTemp.RecordCount = 0 Then
               mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
               AdministrarKeyPreview CInt(KeyCode)
            Else
                ugvDetalleCodigoProductosHis.SetFocus
            End If
        End If
    Else
        mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleCodigoProductosHis
        AdministrarKeyPreview KeyCode
    End If
    'mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleCodigoProductosHis
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, Me.ugvDetalleCodigoProductosHis
    AdministrarKeyPreview KeyCode
End Sub

'MODIFICADO POR YEPE NOVIEMBRE
Private Sub txtDescripcion_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoProductosHis, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleProductosHis = mr_ReglasHIS.ObtenerListaCodigosProductosHisPorNombreYDescripcion(Me.txtCodigo.Text, Me.txtDescripcion.Text)
    Set ugvDetalleCodigoProductosHis.DataSource = oRcs_DetalleProductosHis
End Sub

'MODIFICADO POR EyPE NOVIEMBRE
Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleCodigoProductosHis, SIGHEntidades.GrillaConFilasBicolor
    Set oRcs_DetalleProductosHis = mr_ReglasHIS.ObtenerListaCodigosProductosHisPorNombreYDescripcion(Me.txtCodigo.Text, Me.txtDescripcion.Text)
    Set ugvDetalleCodigoProductosHis.DataSource = oRcs_DetalleProductosHis
End Sub

Private Sub ugvDetalleCodigoProductosHis_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleCodigoProductosHis_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
With Me.ugvDetalleCodigoProductosHis.Bands(0)
    .Columns("iddiagcpt").Hidden = True
'    .Columns("codigoDiagCpt").Header.Caption = "Codigo"
'    .Columns("codigoDiagCpt").Width = 1200
    .Columns("codigoDiagCptSinPunto").Header.Caption = "Codigo"
    .Columns("codigoDiagCptSinPunto").Width = 1200
    .Columns("descripciondiagcpt").Header.Caption = "Descripcion"
    .Columns("descripciondiagcpt").Width = 5100
    .Columns("EsCpt").Hidden = True
    .Columns("DxSexo").Hidden = True
    .Columns("MasDeUnDiagnosticos").Hidden = True
End With
End Sub

Private Sub ugvDetalleCodigoProductosHis_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF3 Then
        btnAceptar_Click
    End If
    AdministrarKeyPreview CInt(KeyCode)
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    If Not ugvDetalleCodigoProductosHis.ActiveRow Is Nothing Then
    ms_IdDiagCpt = Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("iddiagcpt").Value
'    ms_descripciondiagcpt = "(" & Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("codigoDiagCpt").Value & ") - " & Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("descripciondiagcpt").Value
    ms_descripciondiagcpt = "(" & Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("codigoDiagCptSinPunto").Value & ") - " & Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("descripciondiagcpt").Value
    ms_MasDeUnDiagnosticos = Me.ugvDetalleCodigoProductosHis.ActiveRow.Cells("MasDeUnDiagnosticos").Value
    Visible = False
    End If
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ms_IdDiagCpt = 0
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

