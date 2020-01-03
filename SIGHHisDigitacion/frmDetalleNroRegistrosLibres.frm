VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form frmDetalleNroRegistrosLibres 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de registros libres"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "frmDetalleNroRegistrosLibres.frx":0000
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
         DisabledPicture =   "frmDetalleNroRegistrosLibres.frx":000C
         DownPicture     =   "frmDetalleNroRegistrosLibres.frx":046C
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
         Picture         =   "frmDetalleNroRegistrosLibres.frx":08E1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "frmDetalleNroRegistrosLibres.frx":0D56
         DownPicture     =   "frmDetalleNroRegistrosLibres.frx":121A
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
         Picture         =   "frmDetalleNroRegistrosLibres.frx":1706
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
   Begin UltraGrid.SSUltraGrid ugvDetalleRegistros 
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
      Caption         =   "ugvDetalleRegistros"
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Registros libres"
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
Attribute VB_Name = "frmDetalleNroRegistrosLibres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Detalla de número de Registros libres
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim ml_IdHisCabecera As Long
Dim ml_IdUsuario As Long
Dim ml_NroRegistros As Integer
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRcs_DetalleRegistros As New Recordset
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property

Property Let IdHisCabecera(lValue As Long)
   ml_IdHisCabecera = lValue
End Property
Property Get IdHisCabecera() As Long
   IdHisCabecera = ml_IdHisCabecera
End Property

Property Let NroRegistros(iValue As Integer)
   ml_NroRegistros = iValue
End Property
Property Get NroRegistros() As Integer
   NroRegistros = ml_NroRegistros
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub Form_Load()
    Dim lnIndice As Integer
    Dim RegistroUsado As Boolean
    Dim oRcs_RegistrosLibres As New Recordset
    Dim oRcsTemp As Recordset
    'Para cargar los datos de una consulta
    With oRcs_RegistrosLibres
        .Fields.Append "IdRegistro", adInteger, , adFldIsNullable + adFldUpdatable
        .Fields.Append "Registro", adVarChar, 30, adFldIsNullable + adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set oRcsTemp = mr_ReglasHIS.ObtenerDatosDetalleAtencion(IdHisCabecera)
    For lnIndice = 1 To Val(lcBuscaParametro.SeleccionaFilaParametro(272))
        RegistroUsado = False
        If oRcsTemp.RecordCount <> 0 Then
            oRcsTemp.MoveFirst
            Do While Not oRcsTemp.EOF
                If lnIndice = Int(oRcsTemp!NroRegistroHoja) Then
                    RegistroUsado = True
                End If
                oRcsTemp.MoveNext
            Loop
        End If
        If RegistroUsado = False Then
            With oRcs_RegistrosLibres
                .AddNew
                .Fields!IdRegistro = lnIndice
                .Fields!Registro = "Registro Nº " & lnIndice
                .Update
            End With
        End If
    Next
    oRcs_RegistrosLibres.MoveFirst
    Set ugvDetalleRegistros.DataSource = oRcs_RegistrosLibres
    mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleRegistros, SIGHEntidades.GrillaConFilasBicolor
End Sub

Private Sub ugvDetalleRegistros_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleRegistros_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With Me.ugvDetalleRegistros.Bands(0)
        .Columns("IdRegistro").Hidden = True
        .Columns("Registro").Header.Caption = "Hojas Libres"
        .Columns("Registro").Width = 3300
    End With
End Sub

Private Sub ugvDetalleRegistros_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF3 Then
        btnAceptar_Click
    End If
    If KeyCode = vbKeyEscape Then
        btnCancelar_Click
    End If
End Sub

Private Sub btnAceptar_Click()
    mi_BotonPresionado = sghAceptar
    ml_NroRegistros = CLng(Me.ugvDetalleRegistros.ActiveRow.Cells("IdRegistro").Value)
    Visible = False
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ml_NroRegistros = 0
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
