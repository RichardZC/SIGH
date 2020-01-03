VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form AperturaDecaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura de Caja"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   12
      Top             =   2940
      Width           =   5835
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "AperturaDecaja.frx":0000
         DownPicture     =   "AperturaDecaja.frx":04C4
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
         Left            =   2910
         Picture         =   "AperturaDecaja.frx":09B0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "AperturaDecaja.frx":0E9C
         DownPicture     =   "AperturaDecaja.frx":12FC
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
         Left            =   1410
         Picture         =   "AperturaDecaja.frx":1771
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la sesión actual"
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
      TabIndex        =   8
      Top             =   30
      Width           =   5850
      Begin VB.TextBox txtFecha 
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
         Left            =   1605
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1605
         Width           =   4110
      End
      Begin VB.Frame Frame2 
         Height          =   810
         Left            =   1620
         TabIndex        =   13
         Top             =   1980
         Width           =   4140
         Begin Threed.SSOption optServicios 
            Height          =   255
            Left            =   270
            TabIndex        =   4
            Top             =   330
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Servicios"
            Value           =   -1
         End
         Begin Threed.SSOption optFarmacia 
            Height          =   255
            Left            =   2400
            TabIndex        =   5
            Top             =   360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   450
            _Version        =   262144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Farmacia"
         End
      End
      Begin VB.TextBox txtCajero 
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
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox cmbIdCaja 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   750
         Width           =   4125
      End
      Begin VB.ComboBox cmbIdTurno 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1170
         Width           =   4125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   1695
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "Sólo se emitirán Comprobantes de Pago para"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   165
         TabIndex        =   14
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   1230
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Caja"
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
         TabIndex        =   10
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Cajero"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1365
      End
   End
End
Attribute VB_Name = "AperturaDecaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Apertura de CAJA
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mb_Aceptar As Boolean
Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
Dim mo_cmbIdCaja As New ListaDespleglable
Dim mo_cmbIdTurno As New ListaDespleglable
Dim ml_IdCaja As Long
Dim ml_IdTurno As Long
Dim ml_IdUsuario As Long
Dim lbEmiteSoloServicio As Boolean
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idConfiguracionParaPreventa As Long
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_AperturoCajaOK As Boolean
Dim mo_lcNombrePc As String

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Get AperturoCajaOK() As Boolean
    AperturoCajaOK = mo_AperturoCajaOK
End Property
Property Let EmiteSoloServicio(bValue As Boolean)
    lbEmiteSoloServicio = bValue
End Property
Property Get EmiteSoloServicio() As Boolean
    EmiteSoloServicio = lbEmiteSoloServicio
End Property

Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property

Property Let Aceptar(bValue As Boolean)
    mb_Aceptar = bValue
End Property
Property Get Aceptar() As Boolean
    Aceptar = mb_Aceptar
End Property

Property Let IdCaja(lValue As Long)
    mo_cmbIdCaja.BoundText = lValue
End Property

Property Get IdCaja() As Long
    IdCaja = Val(mo_cmbIdCaja.BoundText)
End Property

Property Let IdTurno(lValue As Long)
    mo_cmbIdTurno.BoundText = lValue
End Property

Property Get IdTurno() As Long
    IdTurno = Val(mo_cmbIdTurno.BoundText)
End Property

Property Let NombreCajero(sValue As String)
    txtCajero = sValue
End Property

Property Get NombreCajero() As String
    NombreCajero = txtCajero.Text
End Property


Public Sub ConfigurarCaja()
    mo_cmbIdCaja.BoundColumn = "IdCaja"
    mo_cmbIdCaja.ListField = "Descripcion"
    Set mo_cmbIdCaja.RowSource = mo_AdminCaja.CajaSeleccionarTodosParaLista()
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_AdminCaja.CajaCajaSeleccionarPorNombrePC(mo_lcNombrePc)
    If oRsTmp.RecordCount > 0 Then
       mo_cmbIdCaja.BoundText = oRsTmp.Fields!IdCaja
       cmbIdCaja.Enabled = False
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
End Sub

Public Sub ConfigurarTurno()
    
    mo_cmbIdTurno.BoundColumn = "IdTurno"
    mo_cmbIdTurno.ListField = "Descripcion"
    Set mo_cmbIdTurno.RowSource = mo_AdminCaja.TurnosSeleccionarTodosParaLista()
    
End Sub

'***************daniel barrantes**************
'***************Busca si YA ESTA APERTURADA LA CAJA/TURNO/CAJERO/DIA, si es asi ya no APERTURARLA OTRA VEZ
'***************Cierra la ULTIMA CAJA/TURNO/CAJERO que es menor al DIA
Private Sub btnAceptar_Click()
Dim oRecordset As New ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oDOCajaCaja1 As New DOCajaCaja
Dim sSQL1 As String
Dim lcFechaServ As String: Dim lcFechaFarm As String: Dim lcFechaAper As String
Dim lbApertura As Boolean
Dim lnIdCajaGestion As Long
        mo_AperturoCajaOK = True
        Me.MousePointer = 11
        optServicios_Click (1)
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        '
        sSQL1 = "idCaja=" & mo_cmbIdCaja.BoundText
        Set oRecordset = mo_AdminCaja.CajaGestionSeleccionarXFiltroOrdenadoXFApertura(sSQL1)
        If oRecordset.RecordCount > 0 Then
            oRecordset.MoveFirst
            If IsNull(oRecordset.Fields!fechaCierre) Then
               If MsgBox("La caja '" & Trim(UCase(cmbIdCaja.Text)) & "' no se ha CERRADO desde el Sistema" & Chr(13) & "o está siendo usado por otra  persona" & Chr(13) & Chr(13) & "¿ desea usarla ?", vbQuestion + vbYesNo, "CAJA") = vbYes Then
                  If mo_AdminCaja.CajaGestionActualizaFechaCierreXidCaja(mo_cmbIdCaja.BoundText, Now) = True Then
                  End If
               Else
                  mo_AperturoCajaOK = False
                  oRecordset.Close
                  Set oConexion = Nothing
                  Set oDOCajaCaja1 = Nothing
                  Me.MousePointer = 1
                  Exit Sub
               End If
            End If
            
        End If
        oRecordset.Close
        '
        sSQL1 = "idcaja=" & mo_cmbIdCaja.BoundText & " and idTurno=" & mo_cmbIdTurno.BoundText & " and idCajero=" & ml_IdUsuario & " and day(FechaApertura)=" & Day(Date) & " and month(FechaApertura)=" & Month(Date) & " and year(FechaApertura)=" & Year(Date)
        Set oRecordset = mo_AdminCaja.CajaGestionSeleccionarXFiltroOrdenadoXFApertura(sSQL1)
        lbApertura = False
        If oRecordset.RecordCount = 0 Then
            lbApertura = True
            oRecordset.Close
            sSQL1 = "(fechaCierre is null) and idcaja=" & mo_cmbIdCaja.BoundText & " and idTurno=" & mo_cmbIdTurno.BoundText & " and idCajero=" & ml_IdUsuario
            Set oRecordset = mo_AdminCaja.CajaGestionSeleccionarXFiltroOrdenadoXFApertura(sSQL1)
            If oRecordset.RecordCount > 0 Then
               oRecordset.MoveFirst
               If IsNull(oRecordset.Fields!fechaCierre) Then
                  lnIdCajaGestion = oRecordset.Fields!IdGestionCaja
                  lcFechaAper = oRecordset.Fields!FechaApertura
                  oRecordset.Close
                  'Busca la ultima boleta  emitida por ese cajero en esa caja
                  Set oRecordset = mo_AdminCaja.CajaComprobantePagoSeleccionarXidCajaGestion(lnIdCajaGestion)
                  lcFechaServ = ""
                  If oRecordset.RecordCount > 0 Then
                     oRecordset.MoveFirst
                     Do While Not oRecordset.EOF
                        If oRecordset.Fields!FechaCobranza >= lcFechaAper And oRecordset.Fields!FechaCobranza <= lcBuscaParametro.RetornaFechaHoraServidorSQL() Then
                           lcFechaServ = oRecordset.Fields!FechaCobranza
                           Exit Do
                        End If
                        oRecordset.MoveNext
                     Loop
                  End If
                  oRecordset.Close
                  'Actualiza la ultima FECHA DE CIERRE vacia con=Fecha ultima Boleta Servicios o Farmacia
                  If lcFechaServ <> "" Then

                     If mo_AdminCaja.CajaGestionActualizaFechaCierre(lnIdCajaGestion, CDate(lcFechaServ)) = True Then
                     End If
                  End If
               End If
            End If
        Else
            oRecordset.Close
            'MsgBox "La Caja/Cajero/Turno ya se aperturó para " & Date, vbInformation, "Resultado"
        End If
        '
        Set oDOCajaCaja1 = mo_AdminCaja.CajaSeleccionarPorId(Val(mo_cmbIdCaja.BoundText))
        wxIdTipoComprobanteDefault = oDOCajaCaja1.IdTipoComprobante
'        CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault
        CargaSetup_Caja App.Path & "\archivos", wxIdTipoComprobanteDefault, False
        '
        If lbApertura Then
           mb_Aceptar = True
           Me.Visible = False
        Else
            mb_Aceptar = False
            Me.Visible = False
        End If
        Set oRecordset = Nothing
        Set oDOCajaCaja1 = Nothing
        Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
    optServicios_Click (1)
    mb_Aceptar = False
    Me.Visible = False

End Sub



Private Sub cmbIdCaja_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdCaja
   AdministrarKeyPreview KeyCode

End Sub



Private Sub cmbIdTurno_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, cmbIdTurno
   AdministrarKeyPreview KeyCode

End Sub



Private Sub Form_Activate()
    Dim oRsTmp As New Recordset
    Set oRsTmp = mo_AdminCaja.CajaCajaSeleccionarPorNombrePC(mo_lcNombrePc)
    If oRsTmp.RecordCount > 0 Then
       mo_cmbIdCaja.BoundText = oRsTmp.Fields!IdCaja
       cmbIdCaja.Enabled = False
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing

End Sub

Private Sub Form_Load()
    
    Set mo_cmbIdCaja.MiComboBox = cmbIdCaja
    Set mo_cmbIdTurno.MiComboBox = cmbIdTurno
    
    ConfigurarTurno
    ConfigurarCaja
    mb_Aceptar = False

    mo_cmbIdCaja.BoundText = 2
    mo_cmbIdTurno.BoundText = 1
    txtFecha.Text = lcBuscaParametro.RetornaFechaHoraServidorSQL()
    
    ml_idConfiguracionParaPreventa = Val(lcBuscaParametro.SeleccionaFilaParametro(229))
    Select Case ml_idConfiguracionParaPreventa
    Case 1    'Se emite en Caja, con N° Serie para FARMACIA/SERVICIO   (como funciona actualmente el HRA
         Frame2.Visible = False
         Label4.Visible = False
    Case 2    'Se emite en CAJA con N° serie solo para FARMACIA    (como MINSA pide)
         Frame2.Visible = True
         Label4.Visible = True
    Case 3    'se emite en FARM con N° Serie solo para FARMACIA    (como MINSA pide)
         Frame2.Visible = True
         Label4.Visible = True
    End Select

End Sub

Private Sub optFarmacia_Click(Value As Integer)
    If optFarmacia.Value Then
       lbEmiteSoloServicio = False
    Else
       lbEmiteSoloServicio = True
    End If
End Sub


    
Private Sub optServicios_Click(Value As Integer)
    If optServicios.Value Then
       lbEmiteSoloServicio = True
    Else
       lbEmiteSoloServicio = False
    End If

End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub



