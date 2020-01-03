VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleLotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Lotes"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
   Icon            =   "frmDetalleLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   3090
      Width           =   7335
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmDetalleLotes.frx":000C
         DownPicture     =   "frmDetalleLotes.frx":046C
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
         Left            =   2400
         Picture         =   "frmDetalleLotes.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleLotes.frx":0D56
         DownPicture     =   "frmDetalleLotes.frx":121A
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
         Picture         =   "frmDetalleLotes.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleLotes 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   7335
      _ExtentX        =   12938
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
      Caption         =   "ugvDetalleLotes"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Detalle de Lotes"
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
      TabIndex        =   1
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmDetalleLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de Listado de Lotes Ingresados.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic

Dim ml_IdEstablecimiento As Long
Dim ml_IdUsuario As Long
Dim ml_IdLote As Long
Dim ms_fechaactual As String

Dim ms_Lote As String
Dim nro_Pag As Integer
Dim NumeroPaginasUt As Integer
Dim mi_IdMes As Integer
Dim mi_Mes As String
Dim mi_Anio As Integer
Dim mb_IdCerrado As Boolean

Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRcs_DetalleLotes As New Recordset

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let IdLote(sValue As Long)
   ml_IdLote = sValue
End Property
Property Get IdLote() As Long
   IdLote = ml_IdLote
End Property

Property Let Lote(sValue As String)
   ms_Lote = sValue
End Property
Property Get Lote() As String
   Lote = ms_Lote
End Property

Property Let NumeroPaginas(iValue As Integer)
   nro_Pag = iValue
End Property
Property Get NumeroPaginas() As Integer
   NumeroPaginas = nro_Pag
End Property
'===============================================
Property Let NumeroPaginasUtilizadas(iValue As Integer)
   NumeroPaginasUt = iValue
End Property
Property Get NumeroPaginasUtilizadas() As Integer
   NumeroPaginasUtilizadas = NumeroPaginasUt
End Property
'===============================================
Property Let IdMes(iValue As Integer)
   mi_IdMes = iValue
End Property
Property Get IdMes() As Integer
   IdMes = mi_IdMes
End Property

Property Let Mes(iValue As String)
   mi_Mes = iValue
End Property
Property Get Mes() As String
   Mes = mi_Mes
End Property

Property Let Anio(iValue As Integer)
   mi_Anio = iValue
End Property
Property Get Anio() As Integer
   Anio = mi_Anio
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub Form_Load()
    mb_IdCerrado = False
    Set oRcs_DetalleLotes = mr_ReglasHIS.ConsultarRegistroFiltroLotes(ml_IdEstablecimiento, 0, 0, "", mb_IdCerrado)
    Set ugvDetalleLotes.DataSource = oRcs_DetalleLotes
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleLotes, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub ugvDetalleLotes_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleLotes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With Me.ugvDetalleLotes.Bands(0)
        .Columns("IdHisLote").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("Lote").Header.Caption = "Lote"
        .Columns("Lote").Width = 800
        .Columns("NroHojas").Header.Caption = "Total Paginas"
        .Columns("NroHojas").Width = 1400
        .Columns("idmes").Hidden = True
        .Columns("mes").Header.Caption = "Mes"
        .Columns("mes").Width = 1000
        .Columns("anio").Header.Caption = "Año"
        .Columns("anio").Width = 800
        .Columns("Estado").Header.Caption = "Estado"
        .Columns("Estado").Width = 1100
    End With
End Sub

Private Sub ugvDetalleLotes_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF3 Then
        btnAceptar_Click
    End If
    If KeyCode = vbKeyEscape Then
        btnCancelar_Click
    End If
End Sub

Private Sub btnAceptar_Click()
    If oRcs_DetalleLotes.RecordCount <> 0 Then
        mi_BotonPresionado = sghAceptar
        ml_IdLote = CLng(Me.ugvDetalleLotes.ActiveRow.Cells("IdHisLote").Value)
        ms_Lote = CStr(Me.ugvDetalleLotes.ActiveRow.Cells("Lote").Value)
        nro_Pag = CInt(Me.ugvDetalleLotes.ActiveRow.Cells("NroHojas").Value)
        mi_IdMes = CInt(Me.ugvDetalleLotes.ActiveRow.Cells("idmes").Value)
        mi_Mes = Me.ugvDetalleLotes.ActiveRow.Cells("mes").Value
        mi_Anio = CInt(Me.ugvDetalleLotes.ActiveRow.Cells("anio").Value)
        NumeroPaginasUt = mr_ReglasHIS.ObtenerDatosNumeroHojasUtilizadas(ml_IdLote)
        Visible = False
    End If
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ms_Lote = ""
    nro_Pag = 0
    mi_IdMes = 0
    mi_Anio = 0
    NumeroPaginasUt = 0
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
