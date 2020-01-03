VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form InventarioInicial 
   Caption         =   "Inserta Inventario sin CERRAR desde ICI, IDI"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Migrar desde el SISMEDV.2.0"
      TabPicture(0)   =   "InventarioInicial.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Actualiza Saldos GalenHos=Lolcli"
      TabPicture(1)   =   "InventarioInicial.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Actualiza Precios desde SISMEDV2"
         Height          =   1905
         Left            =   8520
         TabIndex        =   51
         Top             =   600
         Width           =   3255
         Begin VB.CommandButton cmdActualizaPrecios 
            Caption         =   "Actualiza Precios GalenHos"
            Height          =   375
            Left            =   90
            TabIndex        =   54
            Top             =   1410
            Width           =   3075
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Left            =   690
            TabIndex        =   52
            Text            =   "C:\barrantes\mproducto.dbf"
            Top             =   330
            Width           =   2445
         End
         Begin VB.Label Label17 
            Caption         =   "Debe existir ODBC: SISMEDV2 que apunte a        c:\barrantes"
            Height          =   435
            Left            =   90
            TabIndex        =   55
            Top             =   750
            Width           =   3105
         End
         Begin VB.Label Label16 
            Caption         =   "Archivo:"
            Height          =   225
            Left            =   60
            TabIndex        =   53
            Top             =   330
            Width           =   675
         End
      End
      Begin VB.Frame Frame6 
         Enabled         =   0   'False
         Height          =   7125
         Left            =   -74850
         TabIndex        =   39
         Top             =   1170
         Visible         =   0   'False
         Width           =   11655
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   345
            Left            =   90
            TabIndex        =   56
            ToolTipText     =   "Pone a todos los MEDICAMENTOS el saldo"
            Top             =   5670
            Width           =   1335
         End
         Begin VB.TextBox txtBuscaMed 
            Height          =   345
            Left            =   1740
            TabIndex        =   50
            Top             =   5670
            Width           =   7635
         End
         Begin VB.CheckBox chkSoloInventario 
            Caption         =   "Solo graba INVENTARIO anulado"
            Height          =   285
            Left            =   7620
            TabIndex        =   49
            Top             =   6660
            Value           =   1  'Checked
            Width           =   3615
         End
         Begin VB.CommandButton cmdSaldosT 
            Caption         =   "..."
            Height          =   345
            Left            =   10770
            TabIndex        =   48
            ToolTipText     =   "Pone a todos los MEDICAMENTOS el saldo"
            Top             =   5670
            Width           =   465
         End
         Begin VB.TextBox txtSaldoT 
            Height          =   315
            Left            =   9630
            TabIndex        =   47
            Text            =   "0"
            Top             =   5670
            Width           =   1065
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Del"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11160
            TabIndex        =   43
            Top             =   450
            Width           =   435
         End
         Begin VB.CommandButton Command4 
            Caption         =   "New"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11160
            TabIndex        =   42
            ToolTipText     =   "New"
            Top             =   180
            Width           =   435
         End
         Begin VB.CommandButton cmdIgualaSaldos 
            Caption         =   $"InventarioInicial.frx":0038
            Height          =   885
            Left            =   60
            TabIndex        =   40
            Top             =   6090
            Width           =   7215
         End
         Begin MSDataGridLib.DataGrid grdSaldosAjuste 
            Height          =   5445
            Left            =   60
            TabIndex        =   41
            Top             =   150
            Width           =   11025
            _ExtentX        =   19447
            _ExtentY        =   9604
            _Version        =   393216
            Enabled         =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "codigo"
               Caption         =   "Código"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "medicamento"
               Caption         =   "Medicamento"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "saldo"
               Caption         =   "saldo a la F.ajuste"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   10250
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   8100.284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1395.213
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   60
         TabIndex        =   10
         Top             =   2550
         Width           =   8325
         Begin VB.Frame Frame7 
            Caption         =   "Solo para FARMACIA"
            Height          =   645
            Left            =   120
            TabIndex        =   44
            Top             =   3990
            Width           =   8055
            Begin VB.CheckBox chkICIfarm1 
               Caption         =   "Si la Division entre 2 es mayor que CERO aumenta 1 en la cantidad"
               Height          =   345
               Left            =   4350
               TabIndex        =   46
               Top             =   180
               Width           =   3555
            End
            Begin VB.CheckBox chkIciFarm 
               Caption         =   "El formato ICI es la SUMA DE 2 FARMACIAS"
               Height          =   345
               Left            =   120
               TabIndex        =   45
               Top             =   210
               Value           =   1  'Checked
               Width           =   3555
            End
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2130
            TabIndex        =   23
            Text            =   "C:\barrantes"
            Top             =   630
            Width           =   4335
         End
         Begin VB.TextBox txtFechaI 
            Height          =   315
            Left            =   2130
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   990
            Width           =   1305
         End
         Begin VB.Frame Frame3 
            Caption         =   "Precios"
            Height          =   1185
            Left            =   60
            TabIndex        =   15
            Top             =   2700
            Width           =   5295
            Begin VB.TextBox txtPorDist 
               Height          =   315
               Left            =   1860
               TabIndex        =   17
               Text            =   "12.50"
               Top             =   240
               Width           =   1305
            End
            Begin VB.TextBox txtPorVta 
               Height          =   315
               Left            =   1860
               TabIndex        =   16
               Text            =   "25"
               Top             =   630
               Width           =   1305
            End
            Begin VB.Label Label5 
               Caption         =   "% de Precio Compra"
               Height          =   225
               Left            =   3360
               TabIndex        =   21
               Top             =   300
               Width           =   1605
            End
            Begin VB.Label Label6 
               Caption         =   "Precio Distribucion"
               Height          =   225
               Left            =   120
               TabIndex        =   20
               Top             =   300
               Width           =   1365
            End
            Begin VB.Label Label7 
               Caption         =   "% de Precio Compra"
               Height          =   225
               Left            =   3360
               TabIndex        =   19
               Top             =   690
               Width           =   1605
            End
            Begin VB.Label Label8 
               Caption         =   "Precio Venta"
               Height          =   225
               Left            =   120
               TabIndex        =   18
               Top             =   690
               Width           =   1365
            End
         End
         Begin VB.TextBox txtNinvent 
            Height          =   315
            Left            =   2100
            TabIndex        =   14
            Text            =   "Text2"
            Top             =   1830
            Width           =   1305
         End
         Begin VB.TextBox txtIdCentroCosto 
            Height          =   315
            Left            =   2100
            TabIndex        =   13
            Text            =   "999"
            Top             =   2250
            Width           =   1305
         End
         Begin VB.TextBox txtPartida 
            Height          =   315
            Left            =   5700
            TabIndex        =   12
            Text            =   "999"
            Top             =   2310
            Width           =   1305
         End
         Begin VB.CommandButton cmdProcesar 
            Caption         =   "Procesar"
            Height          =   585
            Left            =   6240
            TabIndex        =   11
            Top             =   3240
            Width           =   1965
         End
         Begin MSDataListLib.DataCombo cmbAlmacen 
            Height          =   330
            Left            =   2130
            TabIndex        =   24
            Top             =   210
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbUsuario 
            Height          =   330
            Left            =   2100
            TabIndex        =   25
            Top             =   1410
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Almacen Destino:"
            Height          =   225
            Left            =   120
            TabIndex        =   32
            Top             =   330
            Width           =   1365
         End
         Begin VB.Label Label2 
            Caption         =   "Ruta de Archivos ICI/IDI:"
            Height          =   225
            Left            =   120
            TabIndex        =   31
            Top             =   660
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "F.Inventario:"
            Height          =   225
            Left            =   120
            TabIndex        =   30
            Top             =   1050
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Usuario GalenHos:"
            Height          =   225
            Left            =   90
            TabIndex        =   29
            Top             =   1530
            Width           =   1365
         End
         Begin VB.Label Label9 
            Caption         =   "Numero Inventario:"
            Height          =   225
            Left            =   90
            TabIndex        =   28
            Top             =   1890
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Id Centro Costo:"
            Height          =   225
            Left            =   150
            TabIndex        =   27
            Top             =   2310
            Width           =   1725
         End
         Begin VB.Label Label12 
            Caption         =   "Id Partida:"
            Height          =   225
            Left            =   4500
            TabIndex        =   26
            Top             =   2340
            Width           =   1725
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Consideraciones:"
         Height          =   1995
         Left            =   90
         TabIndex        =   8
         Top             =   480
         Width           =   8295
         Begin VB.ListBox List1 
            Height          =   1620
            Left            =   150
            TabIndex        =   9
            Top             =   240
            Width           =   8025
         End
      End
      Begin VB.Frame Frame5 
         Height          =   885
         Left            =   60
         TabIndex        =   4
         Top             =   7350
         Width           =   8295
         Begin VB.CommandButton cmdElimina 
            Caption         =   "Elimina Datos de FARMACIA y correlativos"
            Height          =   585
            Left            =   2820
            TabIndex        =   6
            Top             =   150
            Width           =   4005
         End
         Begin VB.TextBox txtClave 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   840
            PasswordChar    =   "*"
            TabIndex        =   5
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label10 
            Caption         =   "CLAVE:"
            Height          =   225
            Left            =   210
            TabIndex        =   7
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Height          =   705
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   11655
         Begin VB.CommandButton cmdBuscarS 
            Caption         =   "Buscar"
            Height          =   375
            Left            =   10620
            TabIndex        =   38
            Top             =   240
            Width           =   945
         End
         Begin VB.TextBox txtDctoAjuste 
            Height          =   315
            Left            =   5820
            MaxLength       =   4
            TabIndex        =   36
            Text            =   "AJ01"
            Top             =   240
            Width           =   645
         End
         Begin VB.TextBox txtHajuste 
            Height          =   315
            Left            =   9180
            TabIndex        =   35
            Text            =   "19:00"
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtFAjuste 
            Height          =   315
            Left            =   8130
            TabIndex        =   33
            Text            =   "01/11/2009"
            Top             =   240
            Width           =   1035
         End
         Begin MSDataListLib.DataCombo cmbAlmacenA 
            Height          =   345
            Left            =   810
            TabIndex        =   2
            Top             =   210
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   609
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label13 
            Caption         =   "N° Dcto Ajuste:"
            Height          =   255
            Left            =   4620
            TabIndex        =   37
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label Label15 
            Caption         =   "F.Ajuste:"
            Height          =   225
            Left            =   7380
            TabIndex        =   34
            Top             =   270
            Width           =   705
         End
         Begin VB.Label Label14 
            Caption         =   "Almacen Destino:"
            Height          =   465
            Left            =   60
            TabIndex        =   3
            Top             =   210
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "InventarioInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Inventario inicial
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim oConexionFox As New ADODB.Connection
Dim oRsAlmacenes As New ADODB.Recordset
Dim oRsAlmacenes1 As New ADODB.Recordset
Dim oRsUsuario As New ADODB.Recordset
Dim oRsSaldosAjuste As New ADODB.Recordset
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_farmMovimiento As New DoFarmMovimiento
Dim mo_farmMovimientoNotaIngreso As New DOfarmMovimientoNotaIngreso
Dim mo_FarmInventario As New DoFarmInventario
Dim oDoProveedores As New DoProveedores
Dim lcSql As String
Dim lbEsFarmacia As Boolean
Const ml_idUsuario As Long = 738


Private Sub cmbAlmacen_Click(Area As Integer)
    lbEsFarmacia = False
    If cmbAlmacen.BoundText <> "" Then
        Dim oRsTmp As New Recordset
        lcSql = "select * from FarmAlmacen where idAlmacen=" & cmbAlmacen.BoundText
        oRsTmp.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
        If oRsTmp.RecordCount > 0 Then
           If oRsTmp.Fields!idTipoLocales = "F" Then
              lbEsFarmacia = True
           End If
        End If
        oRsTmp.Close
        Set oRsTmp = Nothing
    End If
End Sub

Private Sub cmbAlmacenA_Click(Area As Integer)
    txtDctoAjuste.Text = "AJ0" & cmbAlmacenA.BoundText
End Sub

Private Sub cmdActualizaPrecios_Click()
    Dim oRsTmp11 As New Recordset
    Dim oRsTmp12 As New Recordset
    Dim oRsFox As New ADODB.Recordset
    Dim lnIdProducto As Long
    Dim lcSql As String
    '
    oRsFox.Open "SELECT * from mProducto", oConexionFox, adOpenKeyset, adLockOptimistic
    oRsFox.MoveFirst
    Do While Not oRsFox.EOF
       oRsTmp12.Open "select * from FactCatalogoBienesInsumos where codigo='" & Trim(oRsFox.Fields!medCod) & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
       If oRsTmp12.RecordCount > 0 Then
          lcSql = "update FactCatalogoBienesInsumosHosp set PrecioUnitario=" & oRsFox.Fields!prdPreOpe & " where idProducto=" & oRsTmp12.Fields!idProducto
          oRsTmp11.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       End If
       oRsTmp12.Close
       oRsFox.MoveNext
    Loop
    oRsFox.Close
    Unload Me
End Sub

Private Sub cmdBuscarS_Click()
    If cmbAlmacenA.Text = "" Then
       MsgBox "Elija el Almacen/Farmacia", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtDctoAjuste.Text = "" Then
       MsgBox "Ingrese el Nro Documento de AJUSTE INVENTARIO", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtFAjuste.Text = "" Then
       MsgBox "Ingrese la Fecha (Reporte LolCli)", vbCritical, Me.Caption
       Exit Sub
    End If
    If txtHajuste.Text = "" Then
       MsgBox "Ingrese la Hora (Reporte LolCli)", vbCritical, Me.Caption
       Exit Sub
    End If
    Dim oRsTmp As New Recordset
    lcSql = "SELECT     dbo.farmInventarioCabecera.*, dbo.FactCatalogoBienesInsumos.Codigo, dbo.FactCatalogoBienesInsumos.Nombre, " & _
            "          dbo.farmInventario.numeroInventario , dbo.farmInventario.FechaCreacion" & _
            " FROM         dbo.farmInventarioCabecera LEFT OUTER JOIN" & _
            "          dbo.FactCatalogoBienesInsumos ON dbo.farmInventarioCabecera.idProducto = dbo.FactCatalogoBienesInsumos.IdProducto LEFT OUTER JOIN" & _
            "          dbo.farmInventario ON dbo.farmInventarioCabecera.idInventario = dbo.farmInventario.idInventario" & _
            " WHERE    dbo.farmInventario.numeroInventario='" & Trim(txtDctoAjuste.Text) & "'" & _
            "          and dbo.farmInventario.idEstadoInventario=0 and dbo.farmInventario.idAlmacen=" & cmbAlmacenA.BoundText    'este inventario debe estar ANULADO
    oRsTmp.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
    If oRsTmp.RecordCount > 0 Then
       '************ya esta registrado el Documento
       
       oRsTmp.MoveFirst
       txtFAjuste.Text = Format(oRsTmp.Fields!FechaCreacion, "dd/mm/yyyy")
       txtHajuste.Text = Format(oRsTmp.Fields!FechaCreacion, "hh:mm")
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!saldo = oRsTmp.Fields!Cantidad
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    Else
       '************registran por  primera vez
       oRsTmp.Close
       lcSql = "select * from FactCatalogoBienesInsumos order by Nombre"
       oRsTmp.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       oRsTmp.MoveFirst
       Do While Not oRsTmp.EOF
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp.Fields!nombre
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Fields!saldo = 0
          oRsSaldosAjuste.Update
          oRsTmp.MoveNext
       Loop
    End If
    Set oRsTmp = Nothing
    Frame4.Enabled = False
    Frame6.Enabled = True
End Sub

Private Sub cmdElimina_Click()
        If UCase(txtClave.Text) <> "DEBB" Then
           Exit Sub
        End If
        Dim oRsInvCab As New ADODB.Recordset
        'actualiza correlativos a CERO
        oRsInvCab.Open "update FarmTipoDocumentos set correlativo=0", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "update FarmRelMod set DocumentoUltimoNumero='0'", wxConexionRed, adOpenKeyset, adLockOptimistic
        'ELIMINA DATOS
        oRsInvCab.Open "DELETE FROM FARMINVENTARIOCABECERA", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMINVENTARIODETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMINVENTARIO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTONOTAINGRESO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOPROGRAMAS", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FacturacionBienesPagos", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FactOrdenesBienes", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FacturacionBienesFinanciamientos", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOVENTASDETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTOVENTAS", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTODETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMMOVIMIENTO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMPREVENTADETALLE", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMPREVENTA", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMSALDODETALLADO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "DELETE FROM FARMSALDO", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "delete from CajaComprobantesPago where idTipoOrden>1", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "select * from factCatalogoBienesInsumosHosp", wxConexionRed, adOpenKeyset, adLockOptimistic
        End
End Sub

Private Sub cmdIgualaSaldos_Click()
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
       Dim oRsTmp As New ADODB.Recordset
       Dim oRsTmp1 As New ADODB.Recordset
       Dim oRsTmp2 As New ADODB.Recordset
       Dim lnTotal As Long, lnIdProducto As Long
       Dim lcLote As String, ldFvencimiento As Date, ldFechaProceso As Date
       Dim lnCantAjusEnt As Long, lnCantAjusSal As Long
       Dim lbSigue As Boolean, lnTotalNI As Double, lnTotalNS As Double
       Dim lnSaldoGalenHos As Long, lnPrecioAjuste As Double
       Dim lcErrores As String
       On Error GoTo ErrSald
       ldFechaProceso = CDate(txtFAjuste.Text & " " & txtHajuste.Text)
       
       'Elimina detalle Anterior
       lcSql = "select * from FarmMovimiento where documentoNumero='" & txtDctoAjuste.Text & _
                    "' and documentoIdTipo=10 and idTipoConcepto=20"
       oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
       If oRsTmp1.RecordCount > 0 Then
          oRsTmp1.MoveFirst
          Do While Not oRsTmp1.EOF
            lbSigue = True
            If oRsTmp1.Fields!movTipo = "E" Then
               If oRsTmp1.Fields!idAlmacenDestino <> Val(cmbAlmacenA.BoundText) Then
                  lbSigue = False
               End If
            Else
               If oRsTmp1.Fields!idAlmacenOrigen <> Val(cmbAlmacenA.BoundText) Then
                  lbSigue = False
               End If
            End If
            If lbSigue = True Then
                lcSql = "delete from FarmMovimientoDetalle where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "delete from FarmMovimientoNotaIngreso where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
                lcSql = "delete from FarmMovimiento where MovNumero='" & oRsTmp1.Fields!movNumero & "' and MovTipo='" & oRsTmp1.Fields!movTipo & "'"
                oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
            End If
            oRsTmp1.MoveNext
          Loop
       End If
       oRsTmp1.Close
       'Proceso
       
       Set oRsTmp = mo_ReglasFarmacia.farmRegeneraSaldos
       lnTotal = oRsTmp.RecordCount
       lnTotalNI = 0: lnTotalNS = 0: lcErrores = ""
       If lnTotal > 0 Then
            'Actualiza por cada Producto el AJUSTE DE INGRESO, AJUSTE DE SALIDA
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
               lnIdProducto = oRsTmp.Fields!idProducto
If oRsTmp.Fields!idProducto = 8 Then
   lnSaldoGalenHos = 0
End If
               ldFvencimiento = oRsTmp.Fields!FechaVencimiento
               lcLote = oRsTmp.Fields!Lote
               lnSaldoGalenHos = 0
               lnPrecioAjuste = 0
               Do While Not oRsTmp.EOF And lnIdProducto = oRsTmp.Fields!idProducto
                  If oRsTmp.Fields!FechaCreacion <= ldFechaProceso Then
                        If oRsTmp.Fields!movTipo = "E" Then
                           If oRsTmp.Fields!idAlmacenDestino = Val(cmbAlmacenA.BoundText) Then
                                lnSaldoGalenHos = lnSaldoGalenHos + oRsTmp.Fields!Cantidad
                                If oRsTmp.Fields!FechaVencimiento > ldFvencimiento Then
                                  ldFvencimiento = oRsTmp.Fields!FechaVencimiento
                                  lcLote = oRsTmp.Fields!Lote
                                End If
                                lnPrecioAjuste = oRsTmp.Fields!Precio
                           End If
                        Else
                           If oRsTmp.Fields!idAlmacenOrigen = Val(cmbAlmacenA.BoundText) Then
                                lnSaldoGalenHos = lnSaldoGalenHos - oRsTmp.Fields!Cantidad
                           End If
                        End If
                  End If
                  oRsTmp.MoveNext
                  If oRsTmp.EOF Then
                     Exit Do
                  End If
               Loop
               oRsSaldosAjuste.MoveFirst
               oRsSaldosAjuste.Find "idProducto=" & lnIdProducto
               If Not oRsSaldosAjuste.EOF Then
                  If lnPrecioAjuste = 0 Then
                     lnPrecioAjuste = 1
                  End If
                  If oRsSaldosAjuste.Fields!saldo > lnSaldoGalenHos Then
                     oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo - lnSaldoGalenHos
                     oRsSaldosAjuste.Fields!movTipo = "E"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round((oRsSaldosAjuste.Fields!saldo - lnSaldoGalenHos) * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     oRsSaldosAjuste.Update
                     lnTotalNI = lnTotalNI + oRsSaldosAjuste.Fields!Total
                  ElseIf oRsSaldosAjuste.Fields!saldo <= lnSaldoGalenHos Then
                     oRsSaldosAjuste.Fields!Cantidad = lnSaldoGalenHos - oRsSaldosAjuste.Fields!saldo
                     oRsSaldosAjuste.Fields!movTipo = "S"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round((lnSaldoGalenHos - oRsSaldosAjuste.Fields!saldo) * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     oRsSaldosAjuste.Update
                     lnTotalNS = lnTotalNS + oRsSaldosAjuste.Fields!Total
                  End If
               End If
            Loop
            '
            'Si no  hay LOTE y F.Vencimiento le asigna lote=1234 y FV=01/02/2010
            oRsSaldosAjuste.MoveFirst
            Do While Not oRsSaldosAjuste.EOF
               If IsNull(oRsSaldosAjuste.Fields!Lote) Or oRsSaldosAjuste.Fields!Lote = "" Then
                  oRsSaldosAjuste.Fields!Lote = "1234"
               End If
               If IsNull(oRsSaldosAjuste.Fields!FechaVencimiento) Then
                  oRsSaldosAjuste.Fields!FechaVencimiento = CDate("01/02/2010")
               End If
               If IsNull(oRsSaldosAjuste.Fields!movTipo) Or oRsSaldosAjuste.Fields!movTipo = "" Then
                     'Producto Nuevo
                     lnPrecioAjuste = 1
                     oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo
                     oRsSaldosAjuste.Fields!movTipo = "E"
                     oRsSaldosAjuste.Fields!Precio = lnPrecioAjuste
                     oRsSaldosAjuste.Fields!Total = Round(lnSaldoGalenHos * lnPrecioAjuste, 2)
                     oRsSaldosAjuste.Fields!Lote = lcLote
                     oRsSaldosAjuste.Fields!FechaVencimiento = ldFvencimiento
                     lnTotalNI = lnTotalNI + oRsSaldosAjuste.Fields!Total
               End If
               oRsSaldosAjuste.MoveNext
            Loop
            'Genera Archivos Cabecera/Detalle de NI-Ajuste
            If lnTotalNI > 0 Then
                With mo_farmMovimiento
                    .DocumentoIdtipo = 10    'ajuste inventario
                    .DocumentoNumero = txtDctoAjuste.Text
                    .FechaCreacion = ldFechaProceso
                    .idAlmacenDestino = Val(cmbAlmacenA.BoundText)
                    .idAlmacenOrigen = 0
                    .idEstadoMovimiento = 1   'registrado
                    .idTipoConcepto = 20     'invenario
                    .idUsuario = ml_idUsuario
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "E"
                    .Observaciones = ""
                    .Total = lnTotalNI
                End With
                With mo_farmMovimientoNotaIngreso
                    .DocumentoFechaRecepcion = ldFechaProceso
                    .IdPaciente = 0
                    .IdComprobantePago = 0
                    .idProveedor = 0
                    .idTipoCompra = 0
                    .idTipoProceso = 0
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "E"
                    .NumeroProceso = ""
                    .OrigenFecha = 0
                    .OrigenIdTipo = 22
                    .OrigenNumero = ""
                    .IdCuentaAtencion = 0
                    .idFuenteFinanciamiento = 0
                End With
                oRsSaldosAjuste.Filter = "movTipo='E'"
                If chkSoloInventario.Value = 0 Then
                    If Not mo_ReglasFarmacia.AgregaDatosDeNotaIngreso(mo_farmMovimiento, mo_farmMovimientoNotaIngreso, oDoProveedores, oRsSaldosAjuste, 0, 1304, "Olidata") = True Then
                       MsgBox "Grabo mal Nota Ingreso por Ajuste", vbCritical, Me.Caption
                    End If
                End If
            End If
            'Genera Archivos Cabecera/Detalle de NS-Ajuste
            If lnTotalNS > 0 Then
                With mo_farmMovimiento
                    .DocumentoIdtipo = 10    'ajuste inventario
                    .DocumentoNumero = txtDctoAjuste.Text
                    .FechaCreacion = ldFechaProceso
                    .idAlmacenDestino = 0
                    .idAlmacenOrigen = Val(cmbAlmacenA.BoundText)
                    .idEstadoMovimiento = 1   'registrado
                    .idTipoConcepto = 20     'invenario
                    .idUsuario = ml_idUsuario
                    .IdUsuarioAuditoria = ml_idUsuario
                    .movTipo = "S"
                    .Observaciones = ""
                    .Total = lnTotalNS
                End With
                oRsSaldosAjuste.Filter = "movTipo='S'"
                If chkSoloInventario.Value = 0 Then
                    If Not AgregaDatosDeNotaSalidaAI(mo_farmMovimiento, oRsSaldosAjuste, 1305, "Olidata") = True Then
                       MsgBox "Grabo mal Nota Salida por Ajuste", vbCritical, Me.Caption
                    End If
                 End If
            End If
        End If
        'Graba el Documento registrado como INVENTARIO ANULADO, con la finalidad que se
        'pueda seguir ingresando datos la proxima vez
        lcSql = "select * from FarmInventario where idEstadoInventario=0 and NumeroInventario='" & txtDctoAjuste.Text & "' and idAlmacen=" & cmbAlmacenA.BoundText
        oRsTmp2.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
        If oRsTmp2.RecordCount > 0 Then
           oRsTmp2.MoveFirst
           Do While Not oRsTmp2.EOF
              lcSql = "delete from FarmInventarioCabecera where idInventario=" & oRsTmp2.Fields!idInventario
              oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
              lcSql = "delete from FarmInventarioDetalle where idInventario=" & oRsTmp2.Fields!idInventario
              oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
              oRsTmp2.Delete
              oRsTmp2.Update
              oRsTmp2.MoveNext
           Loop
        End If
        oRsTmp2.Close
        With mo_FarmInventario
            .FechaCreacion = ldFechaProceso
            .numeroInventario = txtDctoAjuste.Text
            .idAlmacen = Val(cmbAlmacenA.BoundText)
            .idEstadoInventario = 0   'registrado
            .idUsuario = ml_idUsuario
            .IdUsuarioAuditoria = ml_idUsuario
        End With
        oRsSaldosAjuste.Filter = ""
        oRsSaldosAjuste.MoveFirst
        Do While Not oRsSaldosAjuste.EOF
           oRsSaldosAjuste.Fields!Cantidad = oRsSaldosAjuste.Fields!saldo
           oRsSaldosAjuste.Update
           oRsSaldosAjuste.MoveNext
        Loop
        If Not mo_ReglasFarmacia.AgregaDatosDeInventario(mo_FarmInventario, oRsSaldosAjuste, oRsSaldosAjuste, 801, "Olidata", False) = True Then
           MsgBox "Grabo mal el DOCUMENTO", vbCritical, Me.Caption
        End If
       
       '
       Unload Me
    End If
    Exit Sub
ErrSald:
    MsgBox Err.Description
    'Resume
End Sub

Private Sub cmdProcesar_Click()
    If cmbAlmacen.Text = "" Then
       MsgBox "Seleccione algun Almacen"
       cmbAlmacen.SetFocus
       Exit Sub
    End If
    If cmbUsuario.Text = "" Then
       MsgBox "Seleccione algun Usuario"
       cmbAlmacen.SetFocus
       Exit Sub
    End If
    If txtIdCentroCosto.Text = "" Then
       MsgBox "Ingrese el Id de la tabla CENTRO DE COSTOS"
       txtIdCentroCosto.SetFocus
       Exit Sub
    End If
    If MsgBox("Esta seguro?", vbQuestion + vbYesNo, "Mensaje") = vbNo Then
       Exit Sub
    End If
    On Error GoTo ErrProInv
    Dim oRsInvCab As New ADODB.Recordset
    Dim oRsInvDet As New ADODB.Recordset
    Dim oRsCatBienes As New ADODB.Recordset
    Dim oRsCatBienesHos As New ADODB.Recordset
    Dim oRsSeguros As New ADODB.Recordset
    Dim oRsFoxFormDet As New ADODB.Recordset
    Dim oRsFoxFormDet1 As New ADODB.Recordset
    Dim oRsFoxProd As New ADODB.Recordset
    Dim oRsFoxSIS As New ADODB.Recordset
    Dim lcInv  As String: Dim lnIdInventario As Long
    Dim lcCodigo As String: Dim lnTipo As Long
    Dim lnIdProducto As Long: Dim lnPv As Double
    Dim lnPc As Double: Dim lnPd As Double
    Dim lcLote As String: Dim ldFVenc As Date
    Dim lcNombre As String: Dim lnCantidad As Long
    Dim lcTipo As String, lcMedTip As String
    Dim lnRestoDiv As Integer, lbSigue As Boolean
    Dim lnDiasVenc  As Integer
    '
    Me.MousePointer = 11


    oRsFoxFormDet.Open "SELECT * from FormDet", oConexionFox, adOpenKeyset, adLockOptimistic
    If oRsFoxFormDet.RecordCount = 0 Then
        MsgBox "El Ici/idi no tiene datos"
        Me.MousePointer = 1
    Else
        oRsInvCab.Open "select * from farmAlmacen where IdAlmacen=" & cmbAlmacen.BoundText, wxConexionRed, adOpenKeyset, adLockOptimistic
        lcTipo = oRsInvCab.Fields!idTipoLocales & oRsInvCab.Fields!idTipoSuministro
        oRsInvCab.Close
        '
        lcInv = Trim(txtNinvent.Text)
        lnIdInventario = 0
        oRsSeguros.Open "select * from TiposFinanciamiento where ((seIngresPrecios=1 or esFuenteFinanciamiento=0 ) and idTipoFinanciamiento<>0 and idTipoFinanciamiento<>1000)", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvCab.Open "select * from farmInventario where NumeroInventario='" & lcInv & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
        If oRsInvCab.RecordCount > 0 Then
           If oRsInvCab.Fields!idEstadoInventario <> 1 Then
              MsgBox "Existe EL inventario"
              Me.MousePointer = 1
              Exit Sub
           End If
           lnIdInventario = oRsInvCab.Fields!idInventario
           oRsInvCab.Close
           oRsInvCab.Open "delete from farmInventarioDetalle where IdInventario=" & lnIdInventario, wxConexionRed, adOpenKeyset, adLockOptimistic
           oRsInvCab.Open "delete from farmInventarioCabecera where IdInventario=" & lnIdInventario, wxConexionRed, adOpenKeyset, adLockOptimistic
        Else
           oRsInvCab.AddNew
           oRsInvCab.Fields!idAlmacen = Val(cmbAlmacen.BoundText)
           oRsInvCab.Fields!numeroInventario = lcInv
           oRsInvCab.Fields!FechaCreacion = CDate(txtFechaI.Text)
           oRsInvCab.Fields!idEstadoInventario = 1
           oRsInvCab.Fields!idUsuario = Val(cmbUsuario.BoundText)
           oRsInvCab.Update
           lnIdInventario = oRsInvCab.Fields!idInventario
           oRsInvCab.Close
        End If
        oRsInvCab.Open "select * from farmInventarioCabecera", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsInvDet.Open "select * from farmInventarioDetalle", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsCatBienes.Open "update factCatalogoBienesInsumos set PrecioCompra=0,precioDistribucion=0,PrecioDonacion=0,precioUltCompra=0", wxConexionRed, adOpenKeyset, adLockOptimistic
        'oRsCatBienes.Open "delete from factCatalogoBienesInsumosHosp", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsCatBienesHos.Open "select * from factCatalogoBienesInsumosHosp", wxConexionRed, adOpenKeyset, adLockOptimistic
        oRsFoxFormDet.MoveFirst
        Do While Not oRsFoxFormDet.EOF
           If oRsFoxFormDet.Fields!stock_fin > 0 Then
                lcCodigo = Left(oRsFoxFormDet.Fields!codigo_med + "       ", 7)
If Trim(lcCodigo) = "16688" Then
   lnTipo = 0
End If
                oRsFoxProd.Open "SELECT * from xprodu where medcod='" & lcCodigo & "'", oConexionFox, adOpenKeyset, adLockOptimistic
                lnTipo = 1: lcNombre = "": lcMedTip = "M"
                If oRsFoxProd.RecordCount = 0 Then
                   oRsFoxProd.Close
                Else
                   lnTipo = IIf(oRsFoxProd.Fields!medEst = "S", 3, IIf(oRsFoxProd.Fields!medEst = "E", 2, 1))
                   lcNombre = Left(Trim(oRsFoxProd.Fields!medNom) & " " & Trim(oRsFoxProd.Fields!medPres) & " " & Trim(oRsFoxProd.Fields!medcnc), 290) & " " & Trim(oRsFoxProd.Fields!medFF)
                   lcMedTip = oRsFoxProd.Fields!medTip
                    oRsFoxProd.Close
                    '
                    lnPv = oRsFoxFormDet.Fields!Precio
                    lnPc = Round((lnPv * 100) / (Val(txtPorVta.Text) + 100), 2)
                    lnPd = Round((Val(txtPorDist.Text) * lnPc / 100) + lnPc, 2)
                    'Actualiza Catalogo productos y precios
                    oRsCatBienes.Open "select * from factCatalogoBienesInsumos where codigo='" & Trim(lcCodigo) & "'", wxConexionRed, adOpenKeyset, adLockOptimistic
                    If oRsCatBienes.RecordCount = 0 Then
                        oRsCatBienes.AddNew
                        oRsCatBienes.Fields!codigo = lcCodigo
                        oRsCatBienes.Fields!NombreComercial = ""
                        oRsCatBienes.Fields!IdGrupoFarmacologico = 999
                        oRsCatBienes.Fields!IdSubGrupoFarmacologico = 999
                        oRsCatBienes.Fields!IdPartida = Val(txtPartida.Text)
                    End If
                    If lcNombre <> "" Then
                        oRsCatBienes.Fields!nombre = lcNombre
                    End If
                    oRsCatBienes.Fields!PrecioCompra = lnPc
                    oRsCatBienes.Fields!PrecioDistribucion = lnPd
                    oRsCatBienes.Fields!idTipoSalidaBienInsumo = lnTipo
                    oRsCatBienes.Fields!TipoProducto = IIf(UCase(lcMedTip) = "M", 0, 1)
                    oRsCatBienes.Fields!IdCentroCosto = Val(txtIdCentroCosto.Text)
                    oRsCatBienes.Update
                    lnIdProducto = oRsCatBienes.Fields!idProducto
                    oRsCatBienes.Close
                    'Actualiza SEGUROS
                    oRsCatBienes.Open "select * from factCatalogoBienesInsumosHosp where idProducto=" & lnIdProducto, wxConexionRed, adOpenKeyset, adLockOptimistic
                    If oRsCatBienes.RecordCount = 0 Then
                        oRsSeguros.MoveFirst
                        Do While Not oRsSeguros.EOF
                            lbSigue = True
                            If oRsSeguros.Fields!IdTipoFinanciamiento = 2 Then  'Sis
                                lcSql = "select * from Farm_SIS WHERE codigo='" & Trim(lcCodigo) & "'"
                                oRsFoxSIS.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
                                If oRsFoxSIS.RecordCount = 0 Then
                                   lbSigue = False
                                End If
                                oRsFoxSIS.Close
                            End If
                            If lbSigue = True Then
                                oRsCatBienesHos.AddNew
                                oRsCatBienesHos.Fields!PrecioUnitario = lnPv
                                oRsCatBienesHos.Fields!idProducto = lnIdProducto
                                oRsCatBienesHos.Fields!IdTipoFinanciamiento = oRsSeguros.Fields!IdTipoFinanciamiento
                                oRsCatBienesHos.Fields!Activo = 1
                                oRsCatBienesHos.Update
                            End If
                            oRsSeguros.MoveNext
                        Loop
                    End If
                    oRsCatBienes.Close
                    'Agrega Cabecera/detalle Inventario
                    lcLote = "..": ldFVenc = CDate("01/01/" & Trim(Str(Year(Date))))
                    If Not IsNull(oRsFoxFormDet.Fields!fec_exp) Then
                       ldFVenc = oRsFoxFormDet.Fields!fec_exp
                    End If
                    oRsFoxFormDet1.Open "SELECT * from FormDet1 where codigo_Med='" & lcCodigo & "' order by fechVto desc", oConexionFox, adOpenKeyset, adLockOptimistic
                    If oRsFoxFormDet1.RecordCount > 0 Then
                       If Not IsNull(oRsFoxFormDet1.Fields!fechVto) Then
                          If oRsFoxFormDet1.Fields!fechVto >= ldFVenc Then
                             ldFVenc = oRsFoxFormDet1.Fields!fechVto
                             lcLote = Trim(oRsFoxFormDet1.Fields!Lote)
                          End If
                       End If
                    End If
                    oRsFoxFormDet1.Close
                    If ldFVenc = CDate("01/01/" & Trim(Str(Year(Date)))) Then
                       ldFVenc = CDate("31/12/" & Trim(Str(Year(Date))))
                    End If
                    lnCantidad = oRsFoxFormDet.Fields!stock_fin
                    If lbEsFarmacia = True Then
                        If chkIciFarm.Value = 1 Then
                           lnRestoDiv = lnCantidad Mod 2
                           lnCantidad = Round(lnCantidad / 2, 0)
                           
                           If chkICIfarm1.Value = 1 And lnRestoDiv > 0 Then
                              lnCantidad = oRsFoxFormDet.Fields!stock_fin - lnCantidad
                           End If
                        End If
                    End If
                    'No debe haber Productos vencidos
                    lnDiasVenc = (CDate(txtFechaI.Text) + 15) - ldFVenc
                    If lnDiasVenc > 0 Then
                       ldFVenc = CDate("31/12/" & Trim(Str(Year(Date))))
                    End If
                    '
                    oRsInvCab.AddNew
                    oRsInvCab.Fields!idInventario = lnIdInventario
                    oRsInvCab.Fields!idProducto = lnIdProducto
                    oRsInvCab.Fields!Cantidad = lnCantidad
                    oRsInvCab.Fields!Precio = lnPv
                    oRsInvCab.Fields!Total = Round(lnCantidad * lnPv, 2)
                    oRsInvCab.Update
                    oRsInvDet.AddNew
                    oRsInvDet.Fields!idInventario = lnIdInventario
                    oRsInvDet.Fields!idProducto = lnIdProducto
                    oRsInvDet.Fields!Lote = lcLote
                    oRsInvDet.Fields!FechaVencimiento = ldFVenc
                    oRsInvDet.Fields!Cantidad = lnCantidad
                    oRsInvDet.Fields!Precio = lnPv
                    oRsInvDet.Update
                End If
           End If
           oRsFoxFormDet.MoveNext
        Loop
        oRsInvCab.Close
        oRsInvDet.Close
       
        End
    End If
    Exit Sub
ErrProInv:
   MsgBox Err.Description
   'resume
End Sub

Private Sub cmdSaldosT_Click()
  If Val(txtSaldoT.Text) > 0 Then
    oRsSaldosAjuste.MoveFirst
    Do While Not oRsSaldosAjuste.EOF
       oRsSaldosAjuste.Fields!saldo = Val(txtSaldoT.Text)
       oRsSaldosAjuste.Update
       oRsSaldosAjuste.MoveNext
    Loop
  End If
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrDel
    oRsSaldosAjuste.Delete
    oRsSaldosAjuste.Update
ErrDel:
End Sub

Private Sub Command2_Click()
    If MsgBox("Elimina toda la Lista de Medicamentos?", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        If oRsSaldosAjuste.RecordCount > 0 Then
          oRsSaldosAjuste.MoveFirst
          Do While Not oRsSaldosAjuste.EOF
             oRsSaldosAjuste.Delete
             oRsSaldosAjuste.Update
             oRsSaldosAjuste.MoveNext
          Loop
        End If
    End If
End Sub

Private Sub Command4_Click()
    Dim lcCodigoS As String
    Dim oRsTmp1 As New Recordset
    lcCodigoS = InputBox("ingrese Codigo del Producto (SISMED): ")
    lcSql = "select * from FactCatalogoBienesInsumos where codigo='" & lcCodigoS & "'"
    oRsTmp1.Open lcSql, wxConexionRed, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       If oRsSaldosAjuste.RecordCount > 0 Then
            oRsSaldosAjuste.MoveFirst
            oRsSaldosAjuste.Find "idProducto=" & oRsTmp1.Fields!idProducto
       End If
       If oRsSaldosAjuste.EOF Then
          oRsSaldosAjuste.AddNew
          oRsSaldosAjuste.Fields!idProducto = oRsTmp1.Fields!idProducto
          oRsSaldosAjuste.Fields!codigo = oRsTmp1.Fields!codigo
          oRsSaldosAjuste.Fields!medicamento = oRsTmp1.Fields!nombre
          oRsSaldosAjuste.Fields!saldo = 0
          oRsSaldosAjuste.Fields!RegistroSanitario = "."
          oRsSaldosAjuste.Update
       Else
          MsgBox "Ese Código ya existe", vbCritical, Me.Caption
       End If
    End If
    Set oRsTmp1 = Nothing
End Sub

Private Sub Form_Load()
    On Error Resume Next
    List1.AddItem "-Debe existir el ODBC  'SISMEDV2' --> Microsoft Visual Foxpro Driver --> que apunte a: "
    List1.AddItem " " & Text1.Text
    List1.AddItem "-Este proceso Limpiará el Inventario del Almacen elegido, "
    List1.AddItem " siempre y cuando no este CERRADO."
    List1.AddItem "-Pone Precios a los SEGUROS y ademas Compra,Distribucion, Venta. Basado en el Pr.VENTA"
    List1.AddItem "-Solo Pone en SIS los Medicamentos que se encuentran en FARM_SIS.DBF"
    List1.AddItem " (debe estar en la ruta del ODBC=SISMEDV2)"
    txtFechaI.Text = Date
    oConexionFox.CommandTimeout = 150
    oConexionFox.CursorLocation = adUseServer
    oConexionFox.Open "dsn=Sismedv2"
    '
    oRsAlmacenes.Open "select * from farmAlmacen where idTipoLocales<>'X' and idEstado=1 order by descripcion", wxConexionRed, adOpenKeyset, adLockOptimistic
    Set cmbAlmacen.RowSource = oRsAlmacenes
    cmbAlmacen.ListField = "descripcion"
    cmbAlmacen.BoundColumn = "idAlmacen"
    '
    oRsUsuario.Open "select * from Empleados order by apellidoPaterno", wxConexionRed, adOpenKeyset, adLockOptimistic
    Set cmbUsuario.RowSource = oRsUsuario
    cmbUsuario.ListField = "apellidoPaterno"
    cmbUsuario.BoundColumn = "idEmpleado"
    '
    oRsAlmacenes1.Open "select * from farmAlmacen where idTipoLocales<>'X' and idEstado=1 order by descripcion", wxConexionRed, adOpenKeyset, adLockOptimistic
    Set cmbAlmacenA.RowSource = oRsAlmacenes1
    cmbAlmacenA.ListField = "descripcion"
    cmbAlmacenA.BoundColumn = "idAlmacen"
    '
    txtNinvent.Text = Right(Date, 2) & "01"
    '
    GenerarRecordsetTemporal
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    oConexionFox.Close
    oRsAlmacenes.Close
    oRsAlmacenes1.Close
    oRsUsuario.Close
    oRsSaldosAjuste.Close
End Sub


Sub GenerarRecordsetTemporal()
    With oRsSaldosAjuste
          .Fields.Append "IdProducto", adInteger
          .Fields.Append "Codigo", adVarChar, 20, adFldIsNullable
          .Fields.Append "Medicamento", adVarChar, 150, adFldIsNullable
          .Fields.Append "Saldo", adInteger
          .Fields.Append "Lote", adChar, 15
          .Fields.Append "FechaVencimiento", adDate, , adFldIsNullable
          .Fields.Append "Cantidad", adInteger
          .Fields.Append "Precio", adDouble
          .Fields.Append "Total", adDouble
          .Fields.Append "MovNumeroS", adChar, 9, adFldIsNullable
          .Fields.Append "MovTipo", adVarChar, 1, adFldIsNullable
          .Fields.Append "RegistroSanitario", adVarChar, 10, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
    End With
    Set grdSaldosAjuste.DataSource = oRsSaldosAjuste
End Sub


Private Sub grdSaldosAjuste_HeadClick(ByVal ColIndex As Integer)
    Select Case ColIndex
    Case 0   'ordenado por codigo
         oRsSaldosAjuste.Sort = "codigo asc"
    Case 1   'ordenado por descripcion
         oRsSaldosAjuste.Sort = "medicamento asc"
    End Select
End Sub




Function farmRegeneraSaldos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    Set farmRegeneraSaldos = Nothing
    ms_MensajeError = ""
    oConexion.Open wxConexionRed
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmRegeneraSaldos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmRegeneraSaldos = oRecordset
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ActualizaSaldosPorProducto(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnCantidad As Long, lnPrecio As Double) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    ActualizaSaldosPorProducto = False
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = wxConexionRed
       .CommandText = "FarmActualizaSaldosPorProducto"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       .Execute
    End With
    ActualizaSaldosPorProducto = True
    Exit Function
ManejadorDeError:

End Function



Function AgregaDatosDeNotaSalidaAI(oDoMovimiento As DoFarmMovimiento, oRsDetalleProductos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oMovimiento As New farmMovimiento
    Dim oMovimientoDetalle As New farmMovimientoDetalle
    Dim oDoMovimientoDetalle As New DoFarmMovimientoDetalle
    Dim lcCorrelativo As String
    Dim lnItem As Long
    Dim bProcesoOK As Boolean
    oConexion.Open sighEntidades.CadenaConexion
    oConexion.BeginTrans
    bProcesoOK = True
    Set oMovimiento.Conexion = oConexion
    Set oMovimientoDetalle.Conexion = oConexion
    '*********  graba tabla correlativos farmTipoDocumentos  ***************
    lcCorrelativo = oMovimiento.DevuelveYactualizaCorrelativosDeDocumentosES(2)
    '*********  graba tabla Movimientos  ***************
    With oDoMovimiento
       .movNumero = lcCorrelativo
    End With
    
    If Not oMovimiento.Insertar(oDoMovimiento) Then
            bProcesoOK = False: GoTo TerminarNS
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDoMovimiento.IdUsuarioAuditoria, "A", 0, "FarmMovimiento/" & oDoMovimiento.movTipo & "/" & oDoMovimiento.movNumero, oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "")            'ListBarItems.idListItem
    '*********  graba tabla farmMovimientosDetalle,farmSaldo,farmSaldoDetalle  ***************
    
    oDoMovimientoDetalle.IdUsuarioAuditoria = oDoMovimiento.IdUsuarioAuditoria
    oDoMovimientoDetalle.movNumero = oDoMovimiento.movNumero
    oDoMovimientoDetalle.movTipo = oDoMovimiento.movTipo
    lnItem = 1
    oRsDetalleProductos.MoveFirst
    Do While Not oRsDetalleProductos.EOF
       oDoMovimientoDetalle.Cantidad = oRsDetalleProductos.Fields!Cantidad
       oDoMovimientoDetalle.FechaVencimiento = oRsDetalleProductos.Fields!FechaVencimiento
       oDoMovimientoDetalle.idProducto = oRsDetalleProductos.Fields!idProducto
       oDoMovimientoDetalle.Item = lnItem
       oDoMovimientoDetalle.Lote = oRsDetalleProductos.Fields!Lote
       oDoMovimientoDetalle.Precio = oRsDetalleProductos.Fields!Precio
       oDoMovimientoDetalle.RegistroSanitario = ""
       oDoMovimientoDetalle.Total = oRsDetalleProductos.Fields!Total
       If Not oMovimientoDetalle.Insertar(oDoMovimientoDetalle) Then
                bProcesoOK = False: GoTo TerminarNS
       End If
       lnItem = lnItem + 1
       oRsDetalleProductos.MoveNext
    Loop
TerminarNS:
    If bProcesoOK Then
        AgregaDatosDeNotaSalidaAI = True
        oConexion.CommitTrans
    Else
        AgregaDatosDeNotaSalidaAI = False
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oMovimiento = Nothing
    Set oMovimientoDetalle = Nothing
    Set oDoMovimientoDetalle = Nothing
End Function


Private Sub txtBuscaMed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtBuscaMed.Text <> "" Then
        Dim lnLen As Integer
        lnLen = Len(Trim(txtBuscaMed.Text))
        oRsSaldosAjuste.MoveFirst
        Do While Not oRsSaldosAjuste.EOF
           If UCase(Left(txtBuscaMed.Text, lnLen)) = UCase(Left(oRsSaldosAjuste.Fields!medicamento, lnLen)) Then
              Exit Do
           End If
           oRsSaldosAjuste.MoveNext
        Loop
        If oRsSaldosAjuste.EOF Then
           oRsSaldosAjuste.MoveFirst
        End If
    End If
End Sub
