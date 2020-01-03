VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleHojasLibres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Hojas Libres"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "frmDetalleHojasLibres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   3975
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "frmDetalleHojasLibres.frx":000C
         DownPicture     =   "frmDetalleHojasLibres.frx":046C
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
         Left            =   600
         Picture         =   "frmDetalleHojasLibres.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleHojasLibres.frx":0D56
         DownPicture     =   "frmDetalleHojasLibres.frx":121A
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
         Left            =   2070
         Picture         =   "frmDetalleHojasLibres.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleHojas 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4683
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
      Caption         =   "ugvDetalleHojas"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Hojas Libres"
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
      Width           =   3975
   End
End
Attribute VB_Name = "frmDetalleHojasLibres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de Listado de Hojas Libres del Lote
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
Dim nro_Pag As Integer
Dim ml_TotalPaginas As Integer
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

Property Let TotalPaginas(iValue As Integer)
   ml_TotalPaginas = iValue
End Property
Property Get TotalPaginas() As Integer
   TotalPaginas = ml_TotalPaginas
End Property

Property Let NumeroHoja(iValue As Integer)
   nro_Pag = iValue
End Property
Property Get NumeroHoja() As Integer
   NumeroHoja = nro_Pag
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub Form_Load()
    Dim lnIndice As Integer
    Dim HojaUsada As Boolean
    Dim oRcs_HojasLibres As New Recordset
    Dim oRcsTemp As Recordset
    'Para cargar los datos de una consulta
    With oRcs_HojasLibres
        .Fields.Append "IdHoja", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Hoja", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    Set oRcsTemp = mr_ReglasHIS.His_ConsultarHojasRegistradas(ml_IdEstablecimiento, ml_IdLote)
    For lnIndice = 1 To ml_TotalPaginas
        HojaUsada = False
        If oRcsTemp.RecordCount <> 0 Then
            oRcsTemp.MoveFirst
            Do While Not oRcsTemp.EOF
                If lnIndice = Int(oRcsTemp!NroHojaHis) Then
                    HojaUsada = True
                End If
                oRcsTemp.MoveNext
            Loop
        End If
        If HojaUsada = False Then
            With oRcs_HojasLibres
                .AddNew
                .Fields!IdHoja = lnIndice
                .Fields!Hoja = "Hoja Nº " & lnIndice
                .Update
            End With
        End If
    Next
    oRcs_HojasLibres.MoveFirst
    Set ugvDetalleHojas.DataSource = oRcs_HojasLibres
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleHojas, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub ugvDetalleHojas_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleHojas_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With Me.ugvDetalleHojas.Bands(0)
        .Columns("IdHoja").Hidden = True
        .Columns("Hoja").Header.Caption = "Hojas Libres"
        .Columns("Hoja").Width = 3300
    End With
End Sub

Private Sub ugvDetalleHojas_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF3 Then
        btnAceptar_Click
    End If
    If KeyCode = vbKeyEscape Then
        btnCancelar_Click
    End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    nro_Pag = CLng(Me.ugvDetalleHojas.ActiveRow.Cells("IdHoja").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    nro_Pag = 0
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
