VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form rIDI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formato IDI"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rIDI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   30
      TabIndex        =   9
      Top             =   3600
      Width           =   9960
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rIDI.frx":0CCA
         DownPicture     =   "rIDI.frx":118E
         Height          =   700
         Left            =   5108
         Picture         =   "rIDI.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rIDI.frx":1B66
         DownPicture     =   "rIDI.frx":1FC6
         Height          =   700
         Left            =   3578
         Picture         =   "rIDI.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   9945
      Begin VB.CheckBox chkSinMov 
         Caption         =   "Considera aquellos productos sin Movimientos"
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   1080
         Width           =   4305
      End
      Begin VB.Frame fraICI 
         Caption         =   "Formato IDI"
         Height          =   1635
         Left            =   120
         TabIndex        =   15
         Top             =   1770
         Width           =   9735
         Begin VB.Label Label8 
            Caption         =   "*  En la tabla 'FarmAlmacen.codigoSISMED' debe estar definido el Almacén de acuerdo a la codificación SISMEDV2"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1260
            Width           =   9585
         End
         Begin VB.Label Label7 
            Caption         =   "* Debe existir el ODBC: HIS (visual foxpro, tabla libre) que apunte a:   c:\archivos....\galenhos\archivos"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   150
            TabIndex        =   18
            Top             =   270
            Width           =   8595
         End
         Begin VB.Label Label5 
            Caption         =   "* Al imprimir el formato se llena las tablas:    formato.dbf, formdet.dbf, formDetL.dbf, formDetM.dbf"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   150
            TabIndex        =   17
            Top             =   600
            Width           =   8595
         End
         Begin VB.Label Label6 
            Caption         =   "*  Se exporta a la Version del Sismed:  30 de Setiembre del 2011 "
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   930
            Width           =   8595
         End
      End
      Begin VB.CheckBox chkExcel 
         Caption         =   "En Excel"
         Height          =   315
         Left            =   4890
         Picture         =   "rIDI.frx":28B0
         TabIndex        =   12
         Top             =   660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cmbOrden 
         Height          =   330
         ItemData        =   "rIDI.frx":2BC2
         Left            =   7050
         List            =   "rIDI.frx":2BCC
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1020
         Width           =   2745
      End
      Begin VB.ComboBox cmbAlmacen 
         Height          =   330
         Left            =   1350
         TabIndex        =   0
         Top             =   240
         Width           =   8460
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1350
         TabIndex        =   1
         Top             =   660
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   7620
         TabIndex        =   2
         Top             =   630
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2730
         TabIndex        =   13
         Top             =   660
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   9030
         TabIndex        =   14
         Top             =   630
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orden"
         Height          =   210
         Left            =   6480
         TabIndex        =   11
         Top             =   1080
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
         Height          =   210
         Left            =   7110
         TabIndex        =   6
         Top             =   660
         Width           =   435
      End
   End
End
Attribute VB_Name = "rIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Reporte del Formato IDI
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbAlmacen As New SIGHEntidades.ListaDespleglable
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado
Dim ml_TextoDelFiltro As String
Const ml_IdPuntoCarga As Integer = 5
Dim lnIdProducto As Long
Dim mo_Formulario As New SIGHEntidades.Formulario
Dim ml_idUsuario As Long
Dim lcCodigoSismed As String
Dim lbEsDonaciones As Boolean

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
            Dim oRptClaseCry As New rCrystal
            oRptClaseCry.EnArchivoExcel = IIf(chkExcel.Value = 1, True, False)
            oRptClaseCry.IdAlmacen = Val(mo_cmbAlmacen.BoundText)
            oRptClaseCry.FechaInicio = CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.FechaFin = CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", SIGHEntidades.DevuelveFechaSoloFormato_DMY_HMS))
            oRptClaseCry.OrdenadoPor = cmbOrden.ListIndex
            oRptClaseCry.TextoDelFiltro = ml_TextoDelFiltro
            oRptClaseCry.TipoReporte = Me.Name
            oRptClaseCry.CodigoSismed = lcCodigoSismed
            oRptClaseCry.EsDonaciones = lbEsDonaciones
            oRptClaseCry.ConsiderarSinMovimientos = IIf(chkSinMov.Value = 1, True, False)
            oRptClaseCry.Show vbModal
            Set oRptClaseCry = Nothing
             Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    lcCodigoSismed = ""
    ml_TextoDelFiltro = "FILTROS:   Almacén: (" & Trim(cmbAlmacen.Text) & ")      F.Movimiento: (" & txtFdesde.Text & " al " & txtFhasta.Text & ")     Orden: " & cmbOrden.Text & IIf(Me.chkSinMov.Value = 1, "     (" & Me.chkSinMov.Caption & ")", "")
    
    If mo_cmbAlmacen.BoundText = "" Then
        sMensaje = sMensaje + "Por favor elija el Almacén" + Chr(13)
        cmbAlmacen.SetFocus
    Else
        Dim oRsTmp As New Recordset
        Set oRsTmp = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and idAlmacen=" & mo_cmbAlmacen.BoundText)
        If oRsTmp.RecordCount > 0 Then
           lcCodigoSismed = oRsTmp.Fields!CodigoSismed
           lbEsDonaciones = IIf(oRsTmp.Fields!idTipoSuministro = "02", True, False)
        End If
        oRsTmp.Close
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, ""
       Exit Function
    End If
    
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
    LimpiarVariablesDeMemoria
End Sub







Private Sub cmbAlmacen_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbAlmacen

End Sub



Private Sub cmbOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbOrden

End Sub

Private Sub Form_Initialize()
    Set mo_cmbAlmacen.MiComboBox = cmbAlmacen
End Sub

Sub InicializaFechaHora()
    txtFdesde.Text = SIGHEntidades.PrimerFechaDDMMYYDelMesActual
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:01"
    txtHrFin.Text = "23:59"

End Sub
Private Sub Form_Load()
    InicializaFechaHora
    
    mo_cmbAlmacen.BoundColumn = "IdAlmacen"
    mo_cmbAlmacen.ListField = "Descripcion"
    Set mo_cmbAlmacen.RowSource = mo_ReglasFarmacia.FarmAlmacenSeleccionarSegunFiltro("idTipoLocales='A' and (idtipoSuministro='01' or idtipoSuministro='02')")
    cmbOrden.ListIndex = 1
    '
    Dim rsIdAlmacen As Recordset
    Dim oBuscaDondeLabora As New SIGHNegocios.ReglasComunes
    Set rsIdAlmacen = oBuscaDondeLabora.DevuelveSubAreaDondeLaboraElUsuarioDelSistema(sghAlmacenFarmacia, ml_idUsuario)
    Set oBuscaDondeLabora = Nothing
    If rsIdAlmacen.RecordCount > 0 Then
       mo_cmbAlmacen.BoundText = rsIdAlmacen.Fields!idLaboraSubArea
       mo_Formulario.HabilitarDeshabilitar Me.cmbAlmacen, False
    End If
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub





Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            InicializaFechaHora
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_cmbAlmacen = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrFin.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub

Private Sub txtHrInicio_LostFocus()
If Not SIGHEntidades.ValidaHora(txtHrInicio.Text) Then
            MsgBox "La hora ingresada no es correcta", vbInformation, Me.Caption
            InicializaFechaHora
        End If
End Sub
