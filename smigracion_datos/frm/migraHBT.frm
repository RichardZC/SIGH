VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form migraHBT 
   Caption         =   "Actualiza versión SisgalenPlus"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "migraHBT.frx":0000
   ScaleHeight     =   6585
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSql2000 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "migraHBT.frx":0342
      Top             =   1485
      Visible         =   0   'False
      Width           =   10125
   End
   Begin VB.CommandButton cmdMigraUltimaVersion 
      Caption         =   "Migra última Versión "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   135
      TabIndex        =   4
      Top             =   5400
      Width           =   3225
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "migraHBT.frx":0655
      Top             =   75
      Width           =   10125
   End
   Begin VB.TextBox txtTablaProceso 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3450
      TabIndex        =   2
      Top             =   5460
      Width           =   6780
   End
   Begin VB.TextBox txtSql2008 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "migraHBT.frx":07B2
      Top             =   1695
      Visible         =   0   'False
      Width           =   10125
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   3465
      TabIndex        =   0
      Top             =   6165
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "migraHBT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Actualiza Estructuras de la Base de Datos
'        Programado por: Barrantes D
'        Fecha: Enero 2010
'
'------------------------------------------------------------------------------------
Option Explicit
Dim wrs_Gal As New ADODB.Recordset
Dim oRsFarmMovimientoVentas As New ADODB.Recordset
Dim oRsCajaComprobantePago As New ADODB.Recordset
Dim oRsFacturacionBienesFinanciamiento As New ADODB.Recordset
Dim oRsFactOrdenesBienes As New ADODB.Recordset
Dim oRsFacturacionBienesPagos As New ADODB.Recordset
Dim oRsFactOrdenServicio As New ADODB.Recordset
Dim oRsCajaComprobantePagoS As New ADODB.Recordset
Dim oRsFacturacionServicioFinanciamientos As New ADODB.Recordset
Dim oRsFactOrdenServicioPagos As New ADODB.Recordset
Dim oRsFacturacionServicioPagos As New ADODB.Recordset
Dim oRsPatologia As New Recordset
Dim oRsFarmacia As New Recordset
Dim lcSql As String
Dim oRsUltCodigo As Long
Const lnIdUsuario As Long = 738
Const lnIdTipoFinanciamiento As Long = 1
Const lnIdFuenteFinanciamiento As Long = 1
Const ln2020 As Long = 9999999
Const lcVacio As String = "(VACIO)"
Dim mo_conexion As ADODB.Connection
Dim lnErrCA As Long
Dim ml_Errores As String





Function ConvierteEnDias(lcEdad As String) As Long
    If InStr(lcEdad, "h") > 0 Then
       ConvierteEnDias = 1
    ElseIf InStr(lcEdad, "días") > 0 Or InStr(lcEdad, "d") > 0 Then
       ConvierteEnDias = Val(Left(lcEdad, InStr(lcEdad, "d") - 1))
    ElseIf InStr(lcEdad, "meses") > 0 Then
       ConvierteEnDias = (Val(Left(lcEdad, InStr(lcEdad, "meses") - 1)) * 30) + 29
    Else
       ConvierteEnDias = (Val(Left(lcEdad, InStr(lcEdad, "a") - 1)) * 365) + 364
    End If
End Function


Sub MigraUltimaVersion_TablaSIGH_Parte7(oConexHBT As Connection, oConexODBC As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String

    On Error GoTo errMg

    DoEvents
    ProgressBar1.Value = 191
    Me.Refresh
    txtTablaProceso.Text = "His_Establecimientos"
    lcSql = "CREATE TABLE [dbo].[His_Establecimientos] (" & _
            "    [IdEstablecimiento] [int] NOT NULL " & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[His_Establecimientos] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_His_Establecimientos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "         [IdEstablecimiento]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[His_Establecimientos] ADD " & _
            "    CONSTRAINT [FK_His_Establecimientos_Establecimientos] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 192
    Me.Refresh
    txtTablaProceso.Text = "HIS_Meses"
    lcSql = "CREATE TABLE [dbo].[HIS_Meses] (" & _
            "    [IdMes] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Meses] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_HIS_Meses] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "         [IdMes]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_Meses"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_Meses"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
            lcSql = "select * from HIS_Meses where IdMes=" & oRsTmpOpc1.Fields!IdMes
            If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            If oRsTmpOpc.RecordCount > 0 Then
               lbNuevoRegistro = False
            End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdMes = oRsTmpOpc1.Fields!IdMes
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 193
    Me.Refresh
    txtTablaProceso.Text = "His_EstadosLote"
    lcSql = "CREATE TABLE [dbo].[His_EstadosLote] (" & _
            "    [IdEstado] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[His_EstadosLote] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_His_EstadosLote] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idEstado]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_EstadosLote"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_EstadosLote"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from His_EstadosLote where IdEstado=" & _
                                         oRsTmpOpc1.Fields!idEstado
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstado = oRsTmpOpc1.Fields!idEstado
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 194
    Me.Refresh
    txtTablaProceso.Text = "his_ServEstablecimiento y HIS_Cabecera - Relaciones"
    lcSql = "ALTER TABLE HIS_Cabecera DROP " & _
            "    CONSTRAINT FK_HIS_Cabecera_ServPorEstablec"
            
            
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] DROP CONSTRAINT PK__HIS_ServEstablec__387C9C80"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] DROP Column IdHisServEstablecimiento "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] DROP " & _
            "    CONSTRAINT [HIS_ServEstablecimiento_IdEstablecimiento]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "ALTER TABLE HIS_ServEstablecimiento ALTER COLUMN IdEstablecimiento int not NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE HIS_ServEstablecimiento ALTER COLUMN IdServicio int not NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] ADD " & _
            "    CONSTRAINT [FK_HIS_ServEstablecimiento_His_Establecimientos] FOREIGN KEY " & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[His_Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_HIS_ServEstablecimiento] PRIMARY KEY  CLUSTERED " & _
            "    (" & _
            "        [IdEstablecimiento]," & _
            "        [IdServicio]" & _
            "    )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD " & _
            "        [IdEstablecimiento] [int] NOT NULL default (1867)," & _
            "        [IdServicio] [int] NOT NULL default (31)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD" & _
            "    CONSTRAINT [FK_HIS_Cabecera_HIS_ServEstablecimiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]," & _
            "        [IdServicio]" & _
            "    ) REFERENCES [dbo].[HIS_ServEstablecimiento] (" & _
            "        [IdEstablecimiento]," & _
            "        [IdServicio]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD" & _
            "    CONSTRAINT [FK_HIS_Cabecera_HIS_Turnos] FOREIGN KEY" & _
            "    (" & _
            "        [IdTurno]" & _
            "    ) REFERENCES [dbo].[HIS_Turnos] (" & _
            "        [IdHisTurno]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            
    DoEvents
    ProgressBar1.Value = 195
    Me.Refresh
    txtTablaProceso.Text = "HIS_ProgMedEstMR y HIS_Detalle - Relaciones"
    
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ProgMedEstMR] DROP " & _
            "    CONSTRAINT [IdEstablecimiento_IdEstablecimiento]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ProgMedEstMR] DROP " & _
            "    CONSTRAINT [IdServicio_IdServicio]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ProgMedEstMR] ADD " & _
            "    CONSTRAINT [FK_HIS_ProgMedEstMR_HIS_ServEstablecimiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]," & _
            "        [IdServicio]" & _
            "    ) REFERENCES [dbo].[HIS_ServEstablecimiento] (" & _
            "        [IdEstablecimiento]," & _
            "        [IdServicio]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Detalle] WITH NOCHECK ADD " & _
            "    CONSTRAINT [FK_HIS_Detalle_HIS_Financiador] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoFinanciamiento]" & _
            "    ) REFERENCES [dbo].[HIS_Financiador] (" & _
            "        [IdCodigoFinancHis]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Detalle_HIS_TipoEdad] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoEdad]" & _
            "    ) REFERENCES [dbo].[HIS_TipoEdad] (" & _
            "        [IdHisTipoEdad]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Detalle_TiposCondicionPaciente] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstadoaEstablec]" & _
            "    ) REFERENCES [dbo].[TiposCondicionPaciente] (" & _
            "        [IdTipoCondicionPaciente]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Detalle_TiposCondicionPaciente1] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstadoaServicio]" & _
            "    ) REFERENCES [dbo].[TiposCondicionPaciente] (" & _
            "        [IdTipoCondicionPaciente]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Detalle] ADD " & _
            "    [NroRegistroLote] [int] NULL ," & _
            "    [NroRegistroHoja] [int] NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 196
    Me.Refresh
    txtTablaProceso.Text = "HIS_Lotes - Relaciones"
    lcSql = "ALTER TABLE [dbo].[HIS_Lotes] ADD " & _
            "    CONSTRAINT [FK_HIS_Lotes_His_Establecimientos] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[His_Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Lotes] ADD " & _
            "    CONSTRAINT [FK_HIS_Lotes_His_EstadosLote] FOREIGN KEY" & _
            "    (" & _
            "        [idEstadoLote]" & _
            "    ) REFERENCES [dbo].[His_EstadosLote] (" & _
            "        [IdEstado]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Lotes] ADD " & _
            "     CONSTRAINT [FK_HIS_Lotes_HIS_Meses] FOREIGN KEY" & _
            "    (" & _
            "        [Mes]" & _
            "    ) REFERENCES [dbo].[HIS_Meses] (" & _
            "        [IdMes]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Lotes] ADD " & _
            " [DobleDigitacion] [int] NOT NULL default(0)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 197
    Me.Refresh
    txtTablaProceso.Text = "HIS_Paciente - Relaciones"
    lcSql = "ALTER TABLE [dbo].[HIS_Paciente] ADD" & _
            "    CONSTRAINT [FK_HIS_Paciente_TiposDocIdentidad] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoDocumento]" & _
            "    ) REFERENCES [dbo].[TiposDocIdentidad] (" & _
            "        [IdDocIdentidad]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        lcSql = "ALTER TABLE [dbo].[HIS_Paciente] DROP column NroHc_FF"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 198
    Me.Refresh
    txtTablaProceso.Text = "his_faccatalogoservicios - Relaciones"
    lcSql = "DROP INDEX dbo.HIS_FACTCATALOGOSERVICIOS.IX_HIS_FACTCATALOGOSERVICIOS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "DROP INDEX dbo.HIS_FACTCATALOGOSERVICIOS.Indice_HIS_FACTCATALOGOSERVICIOS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE HIS_FACTCATALOGOSERVICIOS ALTER COLUMN IdDiagCpt int not NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_FACTCATALOGOSERVICIOS] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_HIS_FACTCATALOGOSERVICIOS] PRIMARY KEY  CLUSTERED " & _
            "    (" & _
            "        [IdDiagCpt]" & _
            "    )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DetalleDiagnostico] ADD" & _
            "    CONSTRAINT [FK_HIS_DetalleDiagnostico_HIS_FACTCATALOGOSERVICIOS] FOREIGN KEY" & _
            "    (" & _
            "        [IdCIE]" & _
            "    ) REFERENCES [dbo].[HIS_FACTCATALOGOSERVICIOS] (" & _
            "        [IdDiagCpt]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 199
    Me.Refresh
    txtTablaProceso.Text = "HIS_Detalle_Verifica"
    lcSql = "CREATE TABLE [dbo].[HIS_Detalle_Verifica] (" & _
            "    [IdHisDetalle] [int] NOT NULL ,[IdHisCabecera] [int] NOT NULL ," & _
            "    [IdTipoAtencion] [int] NULL ,[DiaAtencion] [int] NULL ," & _
            "    [Sexo] [int] NULL ,[IdNacionalidad] [int] NULL ," & _
            "    [NroDocIdentidad] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NroHijo] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdEtnia] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdTipoDocumento] [int] NULL ," & _
            "    [NroHC_FF] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [CodigoActividad] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdTipoFinanciamiento] [int] NULL ," & _
            "    [IdDistrito] [int] NULL ," & _
            "    [IdTipoEdad] [int] NULL ," & _
            "    [Edad] [int] NULL ," & _
            "    [Talla] [int] NULL ," & _
            "    [Peso] [char] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdEstadoaEstablec] [int] NULL ," & _
            "    [IdEstadoaServicio] [int] NULL ," & _
            "    [NroRegistroLote] [int] NULL ," & _
            "    [NroRegistroHoja] [int] NULL ," & _
            "    [Registrado] [int] NOT NULL ," & _
            "    [Coincide] [Int] NULL," & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Detalle_Verifica] WITH NOCHECK ADD" & _
            "    CONSTRAINT [PK_HIS_Detalle_Verifica] PRIMARY KEY  CLUSTERED " & _
            "    (" & _
            "        [IdHisDetalle]" & _
            "    )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 200
    Me.Refresh
    txtTablaProceso.Text = "HIS_DetalleDx_Verifica"
    lcSql = "CREATE TABLE [dbo].[HIS_DetalleDx_Verifica] (" & _
            "    [IdHisDetalle] [int] NOT NULL ," & _
            "    [IdCIE] [int] NOT NULL ," & _
            "    [IdSubClasificacionDX] [int] NULL ," & _
            "    [CodLAB] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DetalleDx_Verifica] WITH NOCHECK ADD" & _
            "    CONSTRAINT [PK_HIS_DetalleDx_Verifica] PRIMARY KEY  CLUSTERED " & _
            "    (" & _
            "        [IdHisDetalle]," & _
            "        [IdCIE]" & _
            "    )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DetalleDx_Verifica] ADD" & _
            "    CONSTRAINT [FK_HIS_DetalleDx_Verifica_HIS_Detalle_Verifica] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisDetalle]" & _
            "    ) REFERENCES [dbo].[HIS_Detalle_Verifica] (" & _
            "        [IdHisDetalle]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 201
    Me.Refresh
    txtTablaProceso.Text = "FacturacionCuentasAtencionExon"
    lcSql = "CREATE TABLE [dbo].[FacturacionCuentasAtencionExon] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [NumeroExoneracion]  [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[FacturacionCuentasAtencionExon] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_FacturacionCuentasAtencionExon] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idCuentaAtencion]," & _
            "    [NumeroExoneracion]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    

Exit Sub

errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Private Sub cmdMigraUltimaVersion_Click()
    DepuraColumnasDeTablaAtenciones                    'hasta que se termine el MANTENIMIENTO (VENTANAS Y REPORTES)
    '
    On Error GoTo ErroResMigrar
    Dim oConexHBT As New Connection
    Dim oConexODBC As New Connection
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String
    ml_Errores = ""
    
    
    '***************sigh
    oConexHBT.CommandTimeout = 300
    oConexHBT.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\tablas nuevas galenhos.mdb;"
    oConexODBC.CommandTimeout = 300
    oConexODBC.Open "dsn=GALENHOS"
    '
    ProgressBar1.Min = 0
    ProgressBar1.Max = 320
    
    MigraUltimaVersion_TablaSIGH_Parte1 oConexHBT, oConexODBC 'Barra de Proceso del 1 al 64
    MigraUltimaVersion_TablaSIGH_Parte2 oConexHBT, oConexODBC 'Barra de Proceso del 65 al 133
    MigraUltimaVersion_TablaSIGH_Parte3 oConexHBT, oConexODBC 'Barra de Proceso del 134 al 147
    MigraUltimaVersion_TablaSIGH_Parte4 oConexHBT, oConexODBC 'Barra de Proceso del 148 al 172 *PlanIntegralAtencion MMGV
    MigraUltimaVersion_TablaSIGH_Parte5 oConexHBT, oConexODBC 'Barra de Proceso del 173 al 179 *migraIntegrecionSistemas MMGV
    MigraUltimaVersion_TablaSIGH_Parte6 oConexHBT, oConexODBC 'Barra de Proceso del 180 al 190
    MigraUltimaVersion_TablaSIGH_Parte7 oConexHBT, oConexODBC 'Barra de Proceso del 191 al 201
    MigraUltimaVersion_TablaSIGH_Parte8 oConexHBT, oConexODBC 'Barra de Proceso del 202 al 216
    MigraUltimaVersion_TablaSIGH_Parte9 oConexHBT, oConexODBC 'Barra de Proceso del 217 al 222
    MigraUltimaVersion_TablaSIGH_Parte10 oConexHBT, oConexODBC 'Barra de Proceso del 223 al 228
    MigraUltimaVersion_TablaSIGH_Parte11 oConexHBT, oConexODBC 'Barra de Proceso del 229 al 236 *Cambio Nuevo FUA2015
    
    '*********************sigh_externa
    cmdMigraUltimaVErsionExterna oConexODBC, oConexHBT  'Barra de Proceso del 287 al 320
    '
    'On Error Resume Next
    'mo_AdminArchivoClinico.ActualizaDatosConProblemas True
    'On Error GoTo ErroResMigrar
    '
    EliminaProcedAlmacenados
    '
    oConexHBT.Close
    Set oConexHBT = Nothing
    '********************
    
    '
    If ml_Errores <> "" Then
       MsgBox ml_Errores
    End If
    Unload Me
    Exit Sub
ErroResMigrar:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    ElseIf Err.Number = -2147467259 Then
       oConexHBT.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & App.Path & "\migracion\tablas nuevas galenhos.mdb;"
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub


Sub MigraUltimaVersion_TablaSIGH_Parte1(oConexHBT As Connection, oConexODBC As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String
    Dim lnCodigoProducto As Long
    Dim lnProcedimientoRepetido As Integer
    
    On Error GoTo errMg

    DoEvents
    ProgressBar1.Value = 1
    Me.Refresh
    lcSql = "select * from parametros where idparametro=208"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.RecordCount > 0 Then
       lnCodigoEstablecimiento = Val(oRsTmpOpc.Fields!ValorTexto)
    End If
    
    DoEvents
    ProgressBar1.Value = 2
    Me.Refresh
    txtTablaProceso.Text = "Elimina Tablas no usadas"
    lcSql = "drop table  saldosSi"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    DoEvents
    ProgressBar1.Value = 3
    Me.Refresh
    
    '
    lcSql = "alter table ReservacionCamas drop CONSTRAINT Medicos_ReservacionCamas_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table ReservacionCamas drop CONSTRAINT Pacientes_ReservacionCamas_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table ReservacionCamas drop CONSTRAINT Servicios_ReservacionCamas_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  ReservacionCamas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    DoEvents
    ProgressBar1.Value = 4
    Me.Refresh
    '
    lcSql = "alter table Procedimientos drop CONSTRAINT OPCS_Procedimientos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Procedimientos drop CONSTRAINT TiposSexo_Procedimientos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  Procedimientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 5
    Me.Refresh
    lcSql = "drop table  MotivoAnulacion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 6
    Me.Refresh
    lcSql = "drop table Laboratorios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index IX_Registro on auditoria (IdRegistro)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 7
    Me.Refresh
    lcSql = "alter table InterconsultasDiagnosticos drop CONSTRAINT AtencionesInterconsultas_InterconsultasDiagnosticos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table InterconsultasDiagnosticos drop CONSTRAINT ClasificacionDiagnosticos_InterconsultasDiagnosticos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table InterconsultasDiagnosticos drop CONSTRAINT Diagnosticos_InterconsultasDiagnosticos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table InterconsultasDiagnosticos drop CONSTRAINT SubclasificacionDiagnosticos_InterconsultasDiagnosticos_FK11"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  InterconsultasDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 8
    Me.Refresh
    lcSql = "drop table  HIS_edad_000"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 9
    Me.Refresh
    lcSql = "drop table  FarmaciaRecetas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 10
    Me.Refresh
    lcSql = "alter table FacturacionSeguros drop CONSTRAINT FacturacionCuentasAtencion_FacturacionSeguros_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table FacturacionSeguros drop CONSTRAINT FuentesFinanciamiento_CuentasAtencion_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table FacturacionSeguros drop CONSTRAINT TiposFinanciamiento_CuentasAtencion_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  FacturacionSeguros"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       
    
    
    '****************solo para pruebas (inicio)
'    oConexODBC.Close
'    oConexODBC.CommandTimeout = 300
'    oConexODBC.Open "dsn=GalenhosExterna"
'
'    cmdMigraUltimaVErsionExternaSamuel oConexODBC, oConexHBT
    '****************solo para pruebas (inicio)
    
   
    '
    DoEvents
    ProgressBar1.Value = 11
    Me.Refresh
    lcSql = "drop table  FactCatalogoServiciosMINSA"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 12
    Me.Refresh
    lcSql = "alter table CajaComprobantesPago drop CONSTRAINT CajaTiposPago_CajaComprobantesPago_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  CajaTiposPago"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 13
    Me.Refresh
    lcSql = "alter table CajaFormaPagoComprobante drop CONSTRAINT CajaTiposDocumentoPago_CajaDocumentoPago_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  CajaTiposFormasPago"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 14
    Me.Refresh
    lcSql = "alter table CajaTipoCambio drop CONSTRAINT CajaTiposMoneda_CajaTipoCambio_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  CajaTipoCambio"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 15
    Me.Refresh
    lcSql = "alter table CajaFormaPagoComprobante drop CONSTRAINT CajaDocumento_CajaDocumentoPago_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CajaFormaPagoComprobante drop CONSTRAINT CajaTiposDocumentoPago_CajaDocumentoPago_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CajaFormaPagoComprobante drop CONSTRAINT CajaTiposMoneda_CajaDocumentoPago_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  CajaFormaPagoComprobante"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 16
    Me.Refresh
    lcSql = "drop table  BK_FactCatalogoServicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 17
    Me.Refresh
    lcSql = "drop table  bk_CentrosCosto"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 18
    Me.Refresh
    lcSql = "alter table AtencionesInterconsultas drop CONSTRAINT AtencionesInterconsultas_InterconsultasDiagnosticos_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table AtencionesInterconsultas drop CONSTRAINT Medicos_InterconsultaAtencion_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table AtencionesInterconsultas drop CONSTRAINT Medicos_InterconsultaAtencion_FK2"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table  AtencionesInterconsultas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 19
    Me.Refresh
    txtTablaProceso.Text = "Parametros"
    lcSql = "ALTER TABLE Parametros add  Grupo varchar(30) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Parametros"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Parametros"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idParametro=" & oRsTmpOpc1.Fields!IdParametro
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdParametro = oRsTmpOpc1.Fields!IdParametro
                oRsTmpOpc.Fields!ValorTexto = oRsTmpOpc1.Fields!ValorTexto
                oRsTmpOpc.Fields!ValorInt = oRsTmpOpc1.Fields!ValorInt
                oRsTmpOpc.Fields!ValorFloat = oRsTmpOpc1.Fields!ValorFloat
           ElseIf oRsTmpOpc1.Fields!IdParametro = 283 Or oRsTmpOpc1.Fields!IdParametro = 317 Then     'la Etnia debe estar VACIO para obligar a ingresarlo/Morbilidad DEFAULT debe estar VACIO
                oRsTmpOpc.Fields!ValorTexto = ""
           ElseIf oRsTmpOpc1.Fields!IdParametro = 298 Or oRsTmpOpc1.Fields!IdParametro = 323 Then     'WebReniec/WebSis
                oRsTmpOpc.Fields!ValorTexto = oRsTmpOpc1.Fields!ValorTexto
           ElseIf oRsTmpOpc1.Fields!IdParametro = 318 Or oRsTmpOpc1.Fields!IdParametro = 319 Then     'Contraseña del SIS para Importar y Exportar
                oRsTmpOpc.Fields!ValorTexto = oRsTmpOpc1.Fields!ValorTexto
           End If
           oRsTmpOpc.Fields!Tipo = oRsTmpOpc1.Fields!Tipo
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!Grupo = oRsTmpOpc1.Fields!Grupo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 20
    Me.Refresh
    txtTablaProceso.Text = "ListBarGrupos"
    lcSql = "select * from ListBarGrupos order by IdListGrupo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ListBarGrupos order by IdListGrupo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdListGrupo=" & oRsTmpOpc1.Fields!IdListGrupo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdListGrupo = oRsTmpOpc1.Fields!IdListGrupo
                oRsTmpOpc.Fields!Texto = oRsTmpOpc1.Fields!Texto
           End If
           oRsTmpOpc.Fields!Clave = oRsTmpOpc1.Fields!Clave
           oRsTmpOpc.Fields!Indice = oRsTmpOpc1.Fields!Indice
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 20
    Me.Refresh
    txtTablaProceso.Text = "ListBarItems"
    lcSql = "update ListBarItems set Texto='Tipo Tarifa',clave='TipoTarifa' where idListItem=1337"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "update ListBarItems set Texto='Nota de Ingreso Almacén' where IdListItem=1304"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "update ListBarItems set Texto='Nota de Salida Almacén' where IdListItem=1305"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from ListBarItems order by IdListItem"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ListBarItems order by IdListItem"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdListItem=" & oRsTmpOpc1.Fields!IdListItem
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                lcSql = "DBCC CHECKIDENT (ListBarItems, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!IdListItem - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc1.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!Texto = oRsTmpOpc1.Fields!Texto
           oRsTmpOpc.Fields!Clave = oRsTmpOpc1.Fields!Clave
           oRsTmpOpc.Fields!IdListGrupo = oRsTmpOpc1.Fields!IdListGrupo
           oRsTmpOpc.Fields!Indice = oRsTmpOpc1.Fields!Indice
           oRsTmpOpc.Fields!KeyIcon = oRsTmpOpc1.Fields!KeyIcon
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 21
    Me.Refresh
    txtTablaProceso.Text = "ListBarReporte"
    lcSql = "update ListBarReporte set Reporte ='Tipo Tarifa (CAJA)',id_menuReporte='ID_TipoTarifa',modulo='ECONOMIA' where idReporte=170"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "delete from ListBarReporte where idReporte=14 or idReporte=91"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update ListBarReporte set Reporte='Reembolsos Anuales',id_MenuReporte='ID_ReembolsosAnuales' where idReporte=81"  'debb-31/01/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update ListBarReporte set Reporte='Citados y/o atendidos x Consultorios' where idReporte=94"  'debb-22/05/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update ListBarReporte set Reporte='Egresos Emergencia' where idReporte=92"  'debb-22/05/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update listbarreporte set reporte ='Historias clinicas por tipo de historia ' where idReporte=57"  'debb-20/08/2015
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ListBarReporte order by idReporte"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ListBarReporte order by idReporte"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idReporte=" & oRsTmpOpc1.Fields!idReporte
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                lcSql = "DBCC CHECKIDENT (ListBarReporte, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!idReporte - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc1.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!Reporte = oRsTmpOpc1.Fields!Reporte
           oRsTmpOpc.Fields!id_MenuReporte = oRsTmpOpc1.Fields!id_MenuReporte
           oRsTmpOpc.Fields!Modulo = oRsTmpOpc1.Fields!Modulo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 22
    Me.Refresh
    txtTablaProceso.Text = "TiposDestinoAtencion"
    lcSql = "ALTER TABLE TiposDestinoAtencion add  id_destinoAseguradoSIS varchar(1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE TiposDestinoAtencion add  DestinoSEM varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposDestinoAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposDestinoAtencion"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdDestinoAtencion=" & oRsTmpOpc1.Fields!IdDestinoAtencion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdDestinoAtencion = oRsTmpOpc1.Fields!IdDestinoAtencion
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!IdTipoServicio = oRsTmpOpc1.Fields!IdTipoServicio
           oRsTmpOpc.Fields!TipoServicioHosp = oRsTmpOpc1.Fields!TipoServicioHosp
           oRsTmpOpc.Fields!DestinoSEM = oRsTmpOpc1.Fields!DestinoSEM
           oRsTmpOpc.Fields!id_destinoAseguradoSIS = oRsTmpOpc1.Fields!id_destinoAseguradoSIS
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 23
    Me.Refresh
    txtTablaProceso.Text = "EmergenciaCausaExternaMorbilidad"
    lcSql = "ALTER TABLE EmergenciaCausaExternaMorbilidad add  MotivoSEM varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from EmergenciaCausaExternaMorbilidad"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from EmergenciaCausaExternaMorbilidad"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdCausaExternaMorbilidad=" & oRsTmpOpc1.Fields!IdCausaExternaMorbilidad
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdCausaExternaMorbilidad = oRsTmpOpc1.Fields!IdCausaExternaMorbilidad
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!MotivoSEM = oRsTmpOpc1.Fields!MotivoSEM
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 24
    Me.Refresh
    txtTablaProceso.Text = "FactPUntosCarga"
    lcSql = "update FactPuntosCarga set Descripcion='Tiempo de Internamiento' where idPuntoCArga=9"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update FactPuntosCarga set Descripcion='Consumo en el Servicio' where idPuntoCArga=1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from FactPUntosCarga"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from FactPUntosCarga"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPuntoCarga=" & oRsTmpOpc1.Fields!idPuntoCarga
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idPuntoCarga = oRsTmpOpc1.Fields!idPuntoCarga
                oRsTmpOpc.Fields!IdServicio = oRsTmpOpc1.Fields!IdServicio
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Fields!TipoPunto = oRsTmpOpc1.Fields!TipoPunto
           End If
           oRsTmpOpc.Fields!IdUPS = oRsTmpOpc1.Fields!IdUPS
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 25
    Me.Refresh
    txtTablaProceso.Text = "Permisos"
    lcSql = "select * from Permisos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Permisos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPermiso=" & oRsTmpOpc1.Fields!IdPermiso
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPermiso = oRsTmpOpc1.Fields!IdPermiso
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!Modulo = oRsTmpOpc1.Fields!Modulo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 26
    Me.Refresh
    txtTablaProceso.Text = "TiposCargo"
    lcSql = "select * from TiposCargo order by idTipoCargo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposCargo order by idTipoCargo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idTipoCargo=" & oRsTmpOpc1.Fields!idTipoCargo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                lcSql = "DBCC CHECKIDENT (TiposCargo, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!idTipoCargo - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc1.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!Cargo = oRsTmpOpc1.Fields!Cargo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 27
    Me.Refresh
    txtTablaProceso.Text = "FactCatalogoServicios"
    lcSql = "ALTER TABLE FactCatalogoServicios add  codigoSIS varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoServicios add  idEstado int null"     '1->activo, 0->no activo
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update FactCatalogoServicios set  idEstado=1 where idEstado is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    'ACTUALIZADO POR FCV 23042015
    'Consulta el listado de procedimientos del access de migración
    lcSql = "select * from FactCatalogoServicios order by IdProducto"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    
    'Recorrido de uno en uno y busca en la tabla factcatalogoservicios de la tabla SIGH
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lnProcedimientoRepetido = 1
           lcSql = "select * from FactCatalogoServicios where idProducto=" & oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
           
                lcSql = "select * from FactCatalogoServicios where Codigo='" & Trim(oRsTmpOpc1.Fields!Codigo) & "'"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                If oRsTmpOpc2.RecordCount > 0 Then
                    lnProcedimientoRepetido = 0
                End If
                
                lcSql = "DBCC CHECKIDENT (FactCatalogoServicios, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!IdProducto - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
                oRsTmpOpc.Fields!CodMinsa = oRsTmpOpc1.Fields!CodMinsa
                oRsTmpOpc.Fields!NombreMINSA = oRsTmpOpc1.Fields!NombreMINSA
                oRsTmpOpc.Fields!EsCpt = oRsTmpOpc1.Fields!EsCpt
                
                oRsTmpOpc.Fields!IdServicioGrupo = oRsTmpOpc1.Fields!IdServicioGrupo
                oRsTmpOpc.Fields!IdServicioSubGrupo = oRsTmpOpc1.Fields!IdServicioSubGrupo
                oRsTmpOpc.Fields!IdServicioSeccion = oRsTmpOpc1.Fields!IdServicioSeccion
                oRsTmpOpc.Fields!IdServicioSubSeccion = oRsTmpOpc1.Fields!IdServicioSubSeccion
                If IsNull(oRsTmpOpc.Fields!IdPartida) Then
                   oRsTmpOpc.Fields!IdPartida = oRsTmpOpc1.Fields!IdPartida
                End If
                If IsNull(oRsTmpOpc.Fields!IdCentroCosto) Then
                   oRsTmpOpc.Fields!IdCentroCosto = oRsTmpOpc1.Fields!IdCentroCosto
                End If
                oRsTmpOpc.Fields!codigoSIS = IIf(IsNull(oRsTmpOpc1.Fields!codigoSIS), "", oRsTmpOpc1.Fields!codigoSIS)
           Else
                If oRsTmpOpc.Fields!Codigo <> oRsTmpOpc1.Fields!Codigo Then
                    lcSql = "select * from FactCatalogoServicios where Codigo='" & Trim(oRsTmpOpc1.Fields!Codigo) & "'"
                    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
                    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                    If oRsTmpOpc.RecordCount = 0 Then
                        'Resetea el identificador al idproducto maximo, para que continue la secuencia.
                        lcSql = "select max(IdProducto) as IdProducto from FactCatalogoServicios"
                        If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
                        oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                        lnCodigoProducto = 0
                        If oRsTmpOpc2.RecordCount > 0 Then
                            lnCodigoProducto = Val(oRsTmpOpc2.Fields!IdProducto)
                        End If
                        lcSql = "DBCC CHECKIDENT (FactCatalogoServicios, RESEED, " & CStr(lnCodigoProducto) & ")"
                        If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
                        oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                    
                        oRsTmpOpc.AddNew
                        oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                        oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
                        oRsTmpOpc.Fields!CodMinsa = oRsTmpOpc1.Fields!CodMinsa
                        oRsTmpOpc.Fields!NombreMINSA = oRsTmpOpc1.Fields!NombreMINSA
                        oRsTmpOpc.Fields!EsCpt = oRsTmpOpc1.Fields!EsCpt
                        
                        oRsTmpOpc.Fields!IdServicioGrupo = oRsTmpOpc1.Fields!IdServicioGrupo
                        oRsTmpOpc.Fields!IdServicioSubGrupo = oRsTmpOpc1.Fields!IdServicioSubGrupo
                        oRsTmpOpc.Fields!IdServicioSeccion = oRsTmpOpc1.Fields!IdServicioSeccion
                        oRsTmpOpc.Fields!IdServicioSubSeccion = oRsTmpOpc1.Fields!IdServicioSubSeccion
                        If IsNull(oRsTmpOpc.Fields!IdPartida) Then
                           oRsTmpOpc.Fields!IdPartida = oRsTmpOpc1.Fields!IdPartida
                        End If
                        If IsNull(oRsTmpOpc.Fields!IdCentroCosto) Then
                           oRsTmpOpc.Fields!IdCentroCosto = oRsTmpOpc1.Fields!IdCentroCosto
                        End If
                        oRsTmpOpc.Fields!codigoSIS = IIf(IsNull(oRsTmpOpc1.Fields!codigoSIS), "", oRsTmpOpc1.Fields!codigoSIS)
                    End If
                End If
            End If
            oRsTmpOpc.Fields!CodMINSAnoActualiza = oRsTmpOpc1.Fields!CodMINSAnoActualiza
            oRsTmpOpc.Fields!idOpcs = oRsTmpOpc1.Fields!idOpcs
            If lnProcedimientoRepetido = 0 Then
                oRsTmpOpc.Fields!idEstado = lnProcedimientoRepetido
            Else
                oRsTmpOpc.Fields!idEstado = oRsTmpOpc1.Fields!idEstado
            End If
            oRsTmpOpc.Update
            oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    'Resetea el identificador al idproducto maximo, para que continue la secuencia.
    lcSql = "select max(IdProducto) as IdProducto from FactCatalogoServicios"
    If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
    oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lnCodigoProducto = 0
    If oRsTmpOpc2.RecordCount > 0 Then
        lnCodigoProducto = Val(oRsTmpOpc2.Fields!IdProducto)
    End If
    lcSql = "DBCC CHECKIDENT (FactCatalogoServicios, RESEED, " & CStr(lnCodigoProducto) & ")"
    If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
    oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    'Frank 13112014
    lcSql = "update FactCatalogoServicios set EsCpt = 1 where Codigo='99401'"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 28
    Me.Refresh
    txtTablaProceso.Text = "diagnosticos"
    lcSql = "ALTER TABLE diagnosticos add  DescripcionMINSA varchar(250) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE diagnosticos add  codigoCIEsinPto  varchar(7) null"        '30/05/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    'mgaray09
    lcSql = "ALTER TABLE Diagnosticos ADD FechaInicioVigencia DATETIME NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Diagnosticos ADD EsActivo bit not null DEFAULT 1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from diagnosticos order by IdDiagnostico"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
            If oRsTmpOpc1.Fields!IdDiagnostico = 50036 Then
            lcSql = ""
            End If
           lcSql = "select * from diagnosticos where idDiagnostico='" & oRsTmpOpc1.Fields!IdDiagnostico & "'"
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                lcSql = "DBCC CHECKIDENT (diagnosticos, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!IdDiagnostico - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc1.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!CodigoCIE2004 = oRsTmpOpc1.Fields!CodigoCIE2004
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Fields!DescripcionMINSA = oRsTmpOpc1.Fields!DescripcionMINSA
                oRsTmpOpc.Fields!IdCapitulo = oRsTmpOpc1.Fields!IdCapitulo
                oRsTmpOpc.Fields!idGrupo = oRsTmpOpc1.Fields!idGrupo
                oRsTmpOpc.Fields!IdCategoria = oRsTmpOpc1.Fields!IdCategoria
                oRsTmpOpc.Fields!CodigoExportacion = oRsTmpOpc1.Fields!CodigoExportacion
                oRsTmpOpc.Fields!CodigoCIE9 = oRsTmpOpc1.Fields!CodigoCIE9
                oRsTmpOpc.Fields!Gestacion = oRsTmpOpc1.Fields!Gestacion
                oRsTmpOpc.Fields!Morbilidad = oRsTmpOpc1.Fields!Morbilidad
                oRsTmpOpc.Fields!Intrahospitalario = oRsTmpOpc1.Fields!Intrahospitalario
                oRsTmpOpc.Fields!Restriccion = oRsTmpOpc1.Fields!Restriccion
                oRsTmpOpc.Fields!EdadMaxDias = oRsTmpOpc1.Fields!EdadMaxDias
                oRsTmpOpc.Fields!EdadMinDias = oRsTmpOpc1.Fields!EdadMinDias
                oRsTmpOpc.Fields!idTipoSexo = oRsTmpOpc1.Fields!idTipoSexo
                oRsTmpOpc.Fields!ClaseDxHIS = oRsTmpOpc1.Fields!ClaseDxHIS
                'mgaray09
                oRsTmpOpc.Fields!FechaInicioVigencia = oRsTmpOpc1.Fields!FechaInicioVigencia
           Else
               If oRsTmpOpc.Fields!CodigoCIE2004 = oRsTmpOpc1.Fields!CodigoCIE2004 Then
                   oRsTmpOpc.Fields!DescripcionMINSA = oRsTmpOpc1.Fields!DescripcionMINSA
               End If
           End If
            oRsTmpOpc.Fields!CodigoCIE10 = oRsTmpOpc1.Fields!CodigoCIE2004            '06/03/2013
            oRsTmpOpc.Fields!codigoCIEsinPto = oRsTmpOpc1.Fields!codigoCIEsinPto
            'mgaray09
            oRsTmpOpc.Fields!EsActivo = oRsTmpOpc1.Fields!EsActivo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    
    
    '
    DoEvents
    ProgressBar1.Value = 29
    Me.Refresh
    txtTablaProceso.Text = "provincias"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE provincias add  idReniec int null"           '13/6/12
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from provincias"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           LcTexto1 = "0" & Trim(Str(oRsTmpOpc1.Fields!IdProvincia))
           lcSql = "select * from ReniecU where dptoI=" & Left(LcTexto1, 2) & " and provI=" & Mid(LcTexto1, 3, 2)
           If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
           oRsTmpOpc2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
           lcSql = "select * from provincias where IdProvincia=" & oRsTmpOpc1.Fields!IdProvincia
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdProvincia = oRsTmpOpc1.Fields!IdProvincia
           End If
           oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
           oRsTmpOpc.Fields!IdDepartamento = oRsTmpOpc1.Fields!IdDepartamento
           If oRsTmpOpc2.RecordCount > 0 Then
              oRsTmpOpc.Fields!idReniec = Val(oRsTmpOpc2.Fields!DptoR & oRsTmpOpc2.Fields!ProvR)
           End If
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 30
    Me.Refresh
    txtTablaProceso.Text = "distritos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE distritos add  idReniec int null"           '13/6/12
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE INDEX IX_Distritos  ON Distritos (idReniec)"     '13/6/12
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from distritos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           LcTexto1 = "0" & Trim(Str(oRsTmpOpc1.Fields!IdDistrito))
           lcSql = "select * from ReniecU where dptoI=" & Left(LcTexto1, 2) & " and provI=" & Mid(LcTexto1, 3, 2) & " and distI=" & Right(LcTexto1, 2)
           If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
           oRsTmpOpc2.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
           
           
           lcSql = "select * from distritos where IdDistrito=" & oRsTmpOpc1.Fields!IdDistrito
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdDistrito = oRsTmpOpc1.Fields!IdDistrito
           End If
           oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
           oRsTmpOpc.Fields!IdProvincia = oRsTmpOpc1.Fields!IdProvincia
           If oRsTmpOpc2.RecordCount > 0 Then
              oRsTmpOpc.Fields!idReniec = Val(oRsTmpOpc2.Fields!DptoR & oRsTmpOpc2.Fields!ProvR & oRsTmpOpc2.Fields!distR)
           End If
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 31
    Me.Refresh
    txtTablaProceso.Text = "establecimientos"
    lcSql = "select * from establecimientos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from establecimientos where IdEstablecimiento=" & oRsTmpOpc1.Fields!IdEstablecimiento
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdEstablecimiento = oRsTmpOpc1.Fields!IdEstablecimiento
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
           oRsTmpOpc.Fields!IdDistrito = oRsTmpOpc1.Fields!IdDistrito
           oRsTmpOpc.Fields!IdTipo = oRsTmpOpc1.Fields!IdTipo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 32
    Me.Refresh
    txtTablaProceso.Text = "paises"
    lcSql = "ALTER TABLE paises ALTER COLUMN Codigo CHAR(3) not NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from paises"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from paises"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPais=" & oRsTmpOpc1.Fields!IdPais
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPais = oRsTmpOpc1.Fields!IdPais
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 33
    Me.Refresh
    txtTablaProceso.Text = "farmTipoSalidaSismed"
    lcSql = "update farmTipoSalidaBienInsumo set tipo ='IntervenSanitaria' where idTipoSalidaBienInsumo=2"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update farmTipoSalidaBienInsumo set tipo ='Venta e IntSanitaria' where idTipoSalidaBienInsumo=3"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 33
    Me.Refresh
    txtTablaProceso.Text = "farmTipoProductosSismed"
    lcSql = "CREATE TABLE [dbo].[farmTipoProductosSismed] (" & _
            "    [TipoProductoSismed] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[farmTipoProductosSismed] WITH NOCHECK ADD " & _
    " CONSTRAINT [PK_farmTipoSalidaSismed] PRIMARY KEY  CLUSTERED" & _
    " (" & _
    "     [TipoProductoSismed]" & _
    " )  ON [PRIMARY]"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmTipoProductosSismed"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmTipoProductosSismed"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "TipoProductoSismed='" & oRsTmpOpc1.Fields!TipoProductoSismed & "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!TipoProductoSismed = oRsTmpOpc1.Fields!TipoProductoSismed
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    '
    DoEvents
    ProgressBar1.Value = 33
    Me.Refresh
    txtTablaProceso.Text = "FactCatalogoBienesInsumos"
    lcSql = "alter table FactCatalogoBienesInsumos add Petitorio  bit null"      'debb-08/01/2013
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    lcSql = "alter table FactCatalogoBienesInsumos alter column codigo varchar(7)"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  Denominacion varchar(100) null"         'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  Concentracion varchar(100) null"        'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  Presentacion varchar(100) null"         'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  FormaFarmaceutica varchar(10) null"    'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  MaterialEnvase varchar(100) null"       'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  PresentacionEnvase varchar(100) null"   'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  Fabricante varchar(100) null"           'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  Petitorio bit null"                  'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  IdPaisOrigen int null"                  'debb-20/02/2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    'debb-setiembre2014****inicio
    If wxVersionSQL = sghVersionBD.sighSql2000 Then
       lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  TipoProductoSismed varchar(1) null"
    Else
       lcSql = "ALTER TABLE FactCatalogoBienesInsumos add  TipoProductoSismed varchar(1) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    End If
    'debb-setiembre2014***fin
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from FactCatalogoBienesInsumos order by IdProducto"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from FactCatalogoBienesInsumos where Codigo='" & Left(Trim(oRsTmpOpc1.Fields!Codigo), 7) & "'"
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                'lcSql = "DBCC CHECKIDENT (FactCatalogoBienesInsumos, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!IdProducto - 1)) & ")"
                'If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
                'oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
                oRsTmpOpc.Fields!PrecioCompra = oRsTmpOpc1.Fields!PrecioCompra
                oRsTmpOpc.Fields!PrecioDistribucion = oRsTmpOpc1.Fields!PrecioDistribucion
                oRsTmpOpc.Fields!PrecioDonacion = oRsTmpOpc1.Fields!PrecioDonacion
                oRsTmpOpc.Fields!PrecioUltCompra = oRsTmpOpc1.Fields!PrecioUltCompra
                oRsTmpOpc.Fields!StockMinimo = oRsTmpOpc1.Fields!StockMinimo
                oRsTmpOpc.Fields!idTipoSalidaBienInsumo = oRsTmpOpc1.Fields!idTipoSalidaBienInsumo
                oRsTmpOpc.Fields!TipoProducto = oRsTmpOpc1.Fields!TipoProducto
           Else
                If IsNull(oRsTmpOpc.Fields!idTipoSalidaBienInsumo) Then
                   oRsTmpOpc.Fields!idTipoSalidaBienInsumo = oRsTmpOpc1.Fields!idTipoSalidaBienInsumo
                End If
                If IsNull(oRsTmpOpc.Fields!TipoProducto) Then
                   oRsTmpOpc.Fields!TipoProducto = oRsTmpOpc1.Fields!TipoProducto
                End If
           End If
           oRsTmpOpc.Fields!NombreComercial = oRsTmpOpc1.Fields!NombreComercial
           oRsTmpOpc.Fields!IdGrupoFarmacologico = oRsTmpOpc1.Fields!IdGrupoFarmacologico
           oRsTmpOpc.Fields!IdSubGrupoFarmacologico = oRsTmpOpc1.Fields!IdSubGrupoFarmacologico
           'oRsTmpOpc.Fields!IdPartida = oRsTmpOpc1.Fields!IdPartida
           'oRsTmpOpc.Fields!IdCentroCosto = oRsTmpOpc1.Fields!IdCentroCosto
           If Not IsNull(oRsTmpOpc1.Fields!denominacion) Then
              oRsTmpOpc.Fields!denominacion = Left(oRsTmpOpc1.Fields!denominacion, 100)
           End If
           If Not IsNull(oRsTmpOpc1.Fields!Concentracion) Then
              oRsTmpOpc.Fields!Concentracion = Left(oRsTmpOpc1.Fields!Concentracion, 100)
           End If
           If Not IsNull(oRsTmpOpc1.Fields!Presentacion) Then
              oRsTmpOpc.Fields!Presentacion = Left(oRsTmpOpc1.Fields!Presentacion, 100)
           End If
           If Not IsNull(oRsTmpOpc1.Fields!FormaFarmaceutica) Then
              oRsTmpOpc.Fields!FormaFarmaceutica = Left(oRsTmpOpc1.Fields!FormaFarmaceutica, 100)
           End If
           If Not IsNull(oRsTmpOpc1.Fields!MaterialEnvase) Then
              oRsTmpOpc.Fields!MaterialEnvase = oRsTmpOpc1.Fields!MaterialEnvase
           End If
           If Not IsNull(oRsTmpOpc1.Fields!PresentacionEnvase) Then
              oRsTmpOpc.Fields!PresentacionEnvase = oRsTmpOpc1.Fields!PresentacionEnvase
           End If
           If Not IsNull(oRsTmpOpc1.Fields!Fabricante) Then
              oRsTmpOpc.Fields!Fabricante = oRsTmpOpc1.Fields!Fabricante
           End If
           If Not IsNull(oRsTmpOpc1.Fields!IdPaisOrigen) Then
              oRsTmpOpc.Fields!IdPaisOrigen = oRsTmpOpc1.Fields!IdPaisOrigen
           End If
           oRsTmpOpc.Fields!Petitorio = oRsTmpOpc1.Fields!Petitorio
           If Not IsNull(oRsTmpOpc1.Fields!TipoProductoSismed) Then
              oRsTmpOpc.Fields!TipoProductoSismed = oRsTmpOpc1.Fields!TipoProductoSismed
           End If
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    'Atualizado FCV 05042015 - reseteo de la identidad al ultimo registrado.
    lcSql = "select max(IdProducto) as IdProducto from FactCatalogoBienesInsumos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.RecordCount > 0 Then
        lnCodigoProducto = Val(oRsTmpOpc.Fields!IdProducto)
        lcSql = "DBCC CHECKIDENT (FactCatalogoBienesInsumos, RESEED, " & CStr(lnCodigoProducto) & ")"
        If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
        oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    End If
    
    '
    DoEvents
    ProgressBar1.Value = 34
    Me.Refresh
    txtTablaProceso.Text = "HIS_tabetnia"
    lcSql = "CREATE TABLE [HIS_tabetnia] (" & _
            "    [codetni] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [desetni] [char] (48) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [codgen] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [etnias] [char] (24) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_HIS_tabetnia] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [codetni]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_tabetnia"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_tabetnia"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "codetni='" & oRsTmpOpc1.Fields!codetni & "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!codetni = oRsTmpOpc1.Fields!codetni
           End If
           oRsTmpOpc.Fields!desetni = oRsTmpOpc1.Fields!desetni
           oRsTmpOpc.Fields!codgen = oRsTmpOpc1.Fields!codgen
           oRsTmpOpc.Fields!etnias = oRsTmpOpc1.Fields!etnias
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 35
    Me.Refresh
    txtTablaProceso.Text = "PerinatalListas"
    lcSql = "CREATE TABLE [PerinatalListas] (" & _
            "    [idLista] [int] NOT NULL ," & _
            "    [Lista] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_PerinatalListas] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idLista]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalListas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalListas"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idLista=" & oRsTmpOpc1.Fields!idLista
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idLista = oRsTmpOpc1.Fields!idLista
           End If
           oRsTmpOpc.Fields!Lista = oRsTmpOpc1.Fields!Lista
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 36
    Me.Refresh
    txtTablaProceso.Text = "PerinatalModulos"
    lcSql = "CREATE TABLE [PerinatalModulos] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [Modulo] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [AniosDesde] [int] NULL ," & _
            "    [AniosHasta] [int] NULL ," & _
            "    CONSTRAINT [PK_PerinatalModulos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idModulo]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalModulos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalModulos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idModulo=" & oRsTmpOpc1.Fields!idModulo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
           End If
           oRsTmpOpc.Fields!Modulo = oRsTmpOpc1.Fields!Modulo
           oRsTmpOpc.Fields!AniosDesde = oRsTmpOpc1.Fields!AniosDesde
           oRsTmpOpc.Fields!AniosHasta = oRsTmpOpc1.Fields!AniosHasta
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 37
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoMedicamentos"
    lcSql = "CREATE TABLE [PerinatalCatalogoMedicamentos] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL ," & _
            "    [Cantidad] [int] NOT NULL ," & _
            "    CONSTRAINT [PK_PerinatalCatalogoMedicamentos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idModulo]," & _
            "        [idProducto]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalCatalogoMedicamentos_FactCatalogoBienesInsumos] FOREIGN KEY" & _
            "    (" & _
            "        [idProducto]" & _
            "    ) REFERENCES [FactCatalogoBienesInsumos] (" & _
            "        [idProducto]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoMedicamentos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoMedicamentos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idModulo=" & oRsTmpOpc1.Fields!idModulo
              If Not oRsTmpOpc.EOF Then
                 Do While Not oRsTmpOpc.EOF
                    If oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo And oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto Then
                        lbNuevoRegistro = False
                        Exit Do
                    End If
                    oRsTmpOpc.MoveNext
                 Loop
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           End If
           oRsTmpOpc.Fields!Cantidad = oRsTmpOpc1.Fields!Cantidad
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 38
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoDeProcedimientos"
    lcSql = "alter table PerinatalCatalogoDeProcedimientos drop CONSTRAINT FK_PerinatalCatalogoDeProcedimientos_FactCatalogoServicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencionProcedimientos drop CONSTRAINT FK_PerinatalAtencionProcedimientos_PerinatalCatalogoDeProcedimientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table PerinatalCatalogoDeProcedimientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionProcedimientos ADD  labConfHIS varchar(3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL;"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionProcedimientos ADD ItemProcedimiento INT NOT NULL DEFAULT 1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 39
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoDiagnosticos"
    lcSql = "alter table PerinatalCatalogoDiagnosticos drop CONSTRAINT FK_PerinatalCatalogoDiagnosticos_Diagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencionDiagnosticos drop CONSTRAINT FK_PerinatalAtencionesDiagnosticos_PerinatalCatalogoDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos add  labConfHIS varchar(3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos ADD  IdClasificacionDx INT NOT NULL DEFAULT 1;"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos ADD  IdSubclasificacionDx INT NOT NULL DEFAULT 102;"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    'cambio 06112014 frank y mario
    lcSql = "ALTER TABLE FacturacionServicioDespacho ADD  labConfHIS varchar(3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL;"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    lcSql = "drop table PerinatalCatalogoDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 40
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoCpt"
    lcSql = "CREATE TABLE [PerinatalCatalogoCpt] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [CodigoHIS] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    CONSTRAINT [PK_PerinatalCatalogoDeProcedimientos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idProducto]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalCatalogoDeProcedimientos_FactCatalogoServicios] FOREIGN KEY" & _
            "    (" & _
            "        [idProducto]" & _
            "    ) REFERENCES [FactCatalogoServicios] (" & _
            "        [idProducto]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoCpt"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from PerinatalCatalogoCpt where idModulo =" & oRsTmpOpc1.Fields!idModulo & " And idLista = " & oRsTmpOpc1.Fields!idLista & " And idProducto =" & oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
                oRsTmpOpc.Fields!idLista = oRsTmpOpc1.Fields!idLista
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           End If
           oRsTmpOpc.Fields!CodigoHIS = oRsTmpOpc1.Fields!CodigoHIS
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 41
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencion"
    lcSql = "CREATE TABLE [PerinatalAtencion] (" & _
            "    [idPerinatalAtencion] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [GrafXedadEnMeses] [int] NULL ," & _
            "    [GrafYpercentilTE] [int] NULL ," & _
            "    [GrafYpercentilPT] [int] NULL ," & _
            "    [GrafYpercentilPE] [int] NULL ," & _
            "    [GrafYimc] [money] NULL ," & _
            "    [FechaAtencion] [datetime] NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencion] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencion_1] ON [dbo].[PerinatalAtencion]([idPaciente]) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop index PerinatalAtencion.IX_PerinatalAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column idAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column EstimulacionTemprana"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column AlimentacionComplementaria"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column LactanciaMaterna"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column PersonalSalud"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column DemandaIndividual"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column MujerEdadReproductiva"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion drop column MujerGestante"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalAtencion add  FechaAtencion datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[PerinatalAtencion] ADD " & _
    " CONSTRAINT [FK_PerinatalAtencion_Pacientes] FOREIGN KEY" & _
    " (" & _
    "     [IdPaciente]" & _
    " ) REFERENCES [dbo].[Pacientes] (" & _
    "    [IdPaciente]" & _
    " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 42
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencionCred"
    lcSql = "CREATE TABLE [PerinatalAtencionCred] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [EdadEnAnios] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [CredNumero] [int] NOT NULL ," & _
            "    [CredCheck] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [idAtencion] [int] NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencionCRED] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]," & _
            "        [EdadEnAnios]," & _
            "        [CredNumero]," & _
            "        [CredCheck]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalAtencionCRED_PerinatalAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    ) REFERENCES [PerinatalAtencion] (" & _
            "        [idPerinatalAtencion]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalAtencionCred add  idAtencion int null"    'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencionCred] ON [dbo].[PerinatalAtencionCred]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[PerinatalAtencionCred] ADD " & _
     " CONSTRAINT [FK_PerinatalAtencionCred_Atenciones] FOREIGN KEY" & _
     " (" & _
     "  [IdAtencion]" & _
     " ) REFERENCES [dbo].[Atenciones] (" & _
     "   [IdAtencion]" & _
     " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 43
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencionDiagnosticos"
    lcSql = "CREATE TABLE [PerinatalAtencionDiagnosticos] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idDiagnostico] [int] NOT NULL ," & _
            "    [CodigoHIS] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idAtencion] [int] NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencionDiagnosticos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]," & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idDiagnostico]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalAtencionesDiagnosticos_PerinatalAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    ) REFERENCES [PerinatalAtencion] (" & _
            "        [idPerinatalAtencion]" & _
            "    )," & _
            "    CONSTRAINT [FK_PerinatalAtencionesDiagnosticos_PerinatalListas] FOREIGN KEY" & _
            "    (" & _
            "        [idLista]" & _
            "    ) REFERENCES [PerinatalListas] (" & _
            "        [idLista]"
     lcSql = lcSql & "    )," & _
            "    CONSTRAINT [FK_PerinatalAtencionesDiagnosticos_PerinatalModulos] FOREIGN KEY" & _
            "    (" & _
            "        [idModulo]" & _
            "    ) REFERENCES [PerinatalModulos] (" & _
            "        [idModulo]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos add  idAtencion int null"    'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencionDiagnosticos] ON [dbo].[PerinatalAtencionDiagnosticos]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    'mgaray201410e
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos ADD ItemDiagnostico INT NOT NULL DEFAULT 1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PerinatalAtencionDiagnosticos DROP CONSTRAINT PK_PerinatalAtencionDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[PerinatalAtencionDiagnosticos] WITH NOCHECK ADD " & _
            "     CONSTRAINT [PK_PerinatalAtencionDiagnosticos] PRIMARY KEY  CLUSTERED " & _
            " (" & _
            "    [idPerinatalAtencion]" & _
            "    [idModulo]" & _
            "    [idLista]" & _
            "    [idDiagnostico]" & _
            "    [ItemDiagnostico]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 44
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencionMedicamentos"
    lcSql = "CREATE TABLE [PerinatalAtencionMedicamentos] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [idAtencion] [int] NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencionMedicamentos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]," & _
            "        [idModulo]," & _
            "        [idProducto]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalAtencionMedicamentos_PerinatalAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    ) REFERENCES [PerinatalAtencion] (" & _
            "        [idPerinatalAtencion]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalAtencionMedicamentos add  idAtencion int null"    'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencionMedicamentos] ON [dbo].[PerinatalAtencionMedicamentos]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 45
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencionProcedimientos"
    lcSql = "CREATE TABLE [PerinatalAtencionProcedimientos] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [CptEsAutomatico] [bit] NULL ," & _
            "    [CodigoHIS] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    CONSTRAINT [FK_PerinatalAtencionProcedimientos_PerinatalAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    ) REFERENCES [PerinatalAtencion] (" & _
            "        [idPerinatalAtencion]" & _
            "    )," & _
            "    CONSTRAINT [FK_PerinatalAtencionProcedimientos_PerinatalListas] FOREIGN KEY" & _
            "    (" & _
            "        [idLista]" & _
            "    ) REFERENCES [PerinatalListas] (" & _
            "        [idLista]" & _
            "    )," & _
            "    CONSTRAINT [FK_PerinatalAtencionProcedimientos_PerinatalModulos] FOREIGN KEY" & _
            "    (" & _
            "        [idModulo]" & _
            "    ) REFERENCES [PerinatalModulos] (" & _
            "        [idModulo]" & _
            "    )"
     lcSql = lcSql & " ) ON [PRIMARY]"
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     lcSql = "ALTER TABLE PerinatalAtencionProcedimientos add  idAtencion int null"    'debb-03/02/2011
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     lcSql = "CREATE  INDEX [IX_PerinatalAtencionProcedimientos] ON [dbo].[PerinatalAtencionProcedimientos]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     lcSql = "ALTER TABLE PerinatalAtencionProcedimientos ADD IdOrden INT NULL"    'mgaray-03/02/2011
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     '
    DoEvents
    ProgressBar1.Value = 46
    Me.Refresh
     txtTablaProceso.Text = "PerinatalAtencionCred1"
    lcSql = "CREATE TABLE [PerinatalAtencionCred1] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [idModulo] [int] NULL ," & _
            "    [EstimulacionTemprana] [bit] NULL ," & _
            "    [AlimentacionComplementaria] [bit] NULL ," & _
            "    [LactanciaMaterna] [bit] NULL ," & _
            "    [PersonalSalud] [bit] NULL ," & _
            "    [DemandaIndividual] [bit] NULL ," & _
            "    [MujerEdadReproductiva] [bit] NULL ," & _
            "    [MujerGestante] [bit] NULL ," & _
            "    [idAtencion] [int] NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencionCred1] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencionCred1] ON [dbo].[PerinatalAtencionCred1]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PerinatalAtencion add" & _
            "     CONSTRAINT [FK_PerinatalAtencion_PerinatalAtencionCred1] FOREIGN KEY" & _
            " (" & _
            "    [idPerinatalAtencion]" & _
            " ) REFERENCES [PerinatalAtencionCred1]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table [dbo].[PerinatalAtencion] nocheck constraint [FK_PerinatalAtencion_PerinatalAtencionCred1]    "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 47
    Me.Refresh
    txtTablaProceso.Text = "PerinatalAtencionDiaria"
    lcSql = "CREATE TABLE [PerinatalAtencionDiaria] (" & _
            "    [idPerinatalAtencion] [int] NOT NULL ," & _
            "    [idAtencion] [int] NOT NULL ," & _
            "    CONSTRAINT [PK_PerinatalAtencionDiaria] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPerinatalAtencion]," & _
            "        [idAtencion]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalAtencionDiaria_PerinatalAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [idPerinatalAtencion]" & _
            "    ) REFERENCES [PerinatalAtencion] (" & _
            "        [idPerinatalAtencion]" & _
            "    )" & _
            " ) ON [PRIMARY]"                      'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_PerinatalAtencionDiaria] ON [dbo].[PerinatalAtencionDiaria]([idAtencion]) ON [PRIMARY]"  'debb-03/02/2011
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[PerinatalAtencionDiaria] ADD " & _
     " CONSTRAINT [FK_PerinatalAtencionDiaria_Atenciones] FOREIGN KEY" & _
     " (" & _
     "  [IdAtencion]" & _
     " ) REFERENCES [dbo].[Atenciones] (" & _
     "   [IdAtencion]" & _
     " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     '
    DoEvents
    ProgressBar1.Value = 48
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoCptAutomaticos"
    lcSql = "CREATE TABLE [PerinatalCatalogoCptAutomaticos] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [idProductoAutomatico] [int] NOT NULL ," & _
            "    CONSTRAINT [FK_PerinatalCatalogoDeProcedimientosAutomaticos_PerinatalCatalogoDeProcedimientos] FOREIGN KEY" & _
            "    (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idProducto]" & _
            "    ) REFERENCES [PerinatalCatalogoCpt] (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idProducto]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoCptAutomaticos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from PerinatalCatalogoCptAutomaticos where idModulo =" & oRsTmpOpc1.Fields!idModulo & " And idLista = " & oRsTmpOpc1.Fields!idLista & " And idProducto =" & oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
                oRsTmpOpc.Fields!idLista = oRsTmpOpc1.Fields!idLista
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           End If
           oRsTmpOpc.Fields!idProductoAutomatico = oRsTmpOpc1.Fields!idProductoAutomatico
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 49
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoCptMaximaDosis"
    lcSql = "CREATE TABLE [PerinatalCatalogoCptMaximaDosis] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [Limite] [int] NOT NULL ," & _
            "    [EdadInicial] [int] NOT NULL ," & _
            "    [EdadFinal] [int] NOT NULL ," & _
            "    CONSTRAINT [FK_PerinatalCatalogoDeProcedimientosMaxDosis_PerinatalCatalogoDeProcedimientos] FOREIGN KEY" & _
            "    (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idProducto]" & _
            "    ) REFERENCES [PerinatalCatalogoCpt] (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idProducto]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoCptMaximaDosis"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from PerinatalCatalogoCptMaximaDosis where idModulo =" & oRsTmpOpc1.Fields!idModulo & " And idLista = " & oRsTmpOpc1.Fields!idLista & " And idProducto =" & oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
                oRsTmpOpc.Fields!idLista = oRsTmpOpc1.Fields!idLista
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           End If
           oRsTmpOpc.Fields!Limite = oRsTmpOpc1.Fields!Limite
           oRsTmpOpc.Fields!EdadInicial = oRsTmpOpc1.Fields!EdadInicial
           oRsTmpOpc.Fields!EdadFinal = oRsTmpOpc1.Fields!EdadFinal
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 50
    Me.Refresh
    txtTablaProceso.Text = "PerinatalCatalogoCie10"
    lcSql = "CREATE TABLE [PerinatalCatalogoCie10] (" & _
            "    [idModulo] [int] NOT NULL ," & _
            "    [idLista] [int] NOT NULL ," & _
            "    [idDiagnostico] [int] NOT NULL ," & _
            "    [RangoInicio] [money] NULL ," & _
            "    [RangoFinal] [money] NULL ," & _
            "    [CodigoHIS] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [LabHis] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [TipoDx] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    CONSTRAINT [PK_PerinatalCatalogoDiagnosticos] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idModulo]," & _
            "        [idLista]," & _
            "        [idDiagnostico]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PerinatalCatalogoDiagnosticos_Diagnosticos] FOREIGN KEY" & _
            "    (" & _
            "        [idDiagnostico]" & _
            "    ) REFERENCES [Diagnosticos] (" & _
            "        [idDiagnostico]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalCatalogoCie10 add  CodigoHIS varchar(7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalCatalogoCie10 add  LabHis varchar(20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PerinatalCatalogoCie10 add  TipoDx varchar(6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PerinatalCatalogoCie10"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from PerinatalCatalogoCie10 where idModulo =" & oRsTmpOpc1.Fields!idModulo & " And idLista = " & oRsTmpOpc1.Fields!idLista & " And idDiagnostico =" & oRsTmpOpc1.Fields!IdDiagnostico
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idModulo = oRsTmpOpc1.Fields!idModulo
                oRsTmpOpc.Fields!idLista = oRsTmpOpc1.Fields!idLista
                oRsTmpOpc.Fields!IdDiagnostico = oRsTmpOpc1.Fields!IdDiagnostico
           End If
           oRsTmpOpc.Fields!RangoInicio = oRsTmpOpc1.Fields!RangoInicio
           oRsTmpOpc.Fields!RangoFinal = oRsTmpOpc1.Fields!RangoFinal
           oRsTmpOpc.Fields!CodigoHIS = oRsTmpOpc1.Fields!CodigoHIS
           oRsTmpOpc.Fields!LabHis = oRsTmpOpc1.Fields!LabHis
           oRsTmpOpc.Fields!TipoDx = oRsTmpOpc1.Fields!TipoDx
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 51
    Me.Refresh
    txtTablaProceso.Text = "TiposDestacados"
    lcSql = "CREATE TABLE [TiposDestacados] (" & _
            "    [idDestacado] [int] NOT NULL ," & _
            "    [Destacado] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_TiposDestacados] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idDestacado]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposDestacados"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposDestacados"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idDestacado=" & oRsTmpOpc1.Fields!idDestacado
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idDestacado = oRsTmpOpc1.Fields!idDestacado
           End If
           oRsTmpOpc.Fields!Destacado = oRsTmpOpc1.Fields!Destacado
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 52
    Me.Refresh
    txtTablaProceso.Text = "HIS_colegios"
    lcSql = "CREATE TABLE HIS_colegios (" & _
            "cod_col char(2) NOT NULL CONSTRAINT PK_HIS_colegios PRIMARY KEY," & _
            "des_col char(41) NOT NULL)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_colegios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_colegios"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "cod_col='" & oRsTmpOpc1.Fields!cod_col & "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!cod_col = oRsTmpOpc1.Fields!cod_col
           End If
           oRsTmpOpc.Fields!des_col = oRsTmpOpc1.Fields!des_col
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 53
    Me.Refresh
    txtTablaProceso.Text = "TiposEmpleado"
    lcSql = "ALTER TABLE TiposEmpleado add  EsProgramado bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE TiposEmpleado add  TipoEmpleadoSIS varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE TiposEmpleado add EsColegiatura bit not null DEFAULT 1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposEmpleado"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
       oRsTmpOpc1.MoveFirst
       Do While Not oRsTmpOpc1.EOF
            lcSql = "select * from TiposEmpleado where idTipoEmpleado=" & oRsTmpOpc1.Fields!IdTipoEmpleado
            If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            If oRsTmpOpc.RecordCount > 0 Then
                oRsTmpOpc.Fields!TipoEmpleadoHIS = oRsTmpOpc1.Fields!TipoEmpleadoHIS
                oRsTmpOpc.Fields!Esprogramado = oRsTmpOpc1.Fields!Esprogramado
                oRsTmpOpc.Fields!tipoEmpleadoSis = oRsTmpOpc1.Fields!tipoEmpleadoSis
                oRsTmpOpc.Fields!EsColegiatura = oRsTmpOpc1.Fields!EsColegiatura
                oRsTmpOpc.Update
            Else
                ml_Errores = ml_Errores & "no exite TiposEmpleados.idTipoEmpleado:" & oRsTmpOpc1.Fields!IdTipoEmpleado & Chr(13)
            End If
            oRsTmpOpc1.MoveNext
       Loop
    End If
    '
    DoEvents
    ProgressBar1.Value = 54
    Me.Refresh
    txtTablaProceso.Text = "TiposDocIdentidad"
    lcSql = "ALTER TABLE TiposDocIdentidad add  CodigoSIS varchar(1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE TiposDocIdentidad add  CodigoHIS varchar(1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE TiposDocIdentidad add  CodigoSUNASA int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposDocIdentidad"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
       oRsTmpOpc1.MoveFirst
       Do While Not oRsTmpOpc1.EOF
          lcSql = "select * from TiposDocIdentidad where idDocIdentidad=" & oRsTmpOpc1.Fields!IdDocIdentidad
          If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
          oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
          If oRsTmpOpc.RecordCount = 0 Then
             oRsTmpOpc.AddNew
             oRsTmpOpc.Fields!IdDocIdentidad = oRsTmpOpc1.Fields!IdDocIdentidad
          End If
          oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
          oRsTmpOpc.Fields!CodigoSUNASA = oRsTmpOpc1.Fields!CodigoSUNASA
          oRsTmpOpc.Fields!CodigoHIS = oRsTmpOpc1.Fields!CodigoHIS
          oRsTmpOpc.Fields!codigoSIS = oRsTmpOpc1.Fields!codigoSIS
          oRsTmpOpc.Update
          oRsTmpOpc1.MoveNext
       Loop
    End If
    '
    DoEvents
    ProgressBar1.Value = 55
    Me.Refresh
    txtTablaProceso.Text = "Pacientes"
    lcSql = "drop index Pacientes.ind_autogenerado"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column autogenerado varchar(30)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index ind_autogenerado on Pacientes (autogenerado)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop index pacientes.ind_apellido"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column apellidoPaterno varchar(40)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column apellidoMaterno varchar(40)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column PrimerNombre varchar(40)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column SegundoNombre varchar(40)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Pacientes alter column TercerNombre varchar(40)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index Ind_apellido on Pacientes (ApellidoPaterno,ApellidoMaterno,PrimerNombre)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  IdEtnia varchar(2)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update Pacientes set  IdEtnia='80' where idEtnia is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  GrupoSanguineo varchar(10) null"         'debb-20-02-2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  FactorRh varchar(10) null"               'debb-20-02-2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  UsoWebReniec bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indFichaFamiliar  ON Pacientes (fichaFamiliar)"         'debb-28-08-2012
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  idIdioma int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  Email varchar(50) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madreDocumento varchar(12) null"             'debb-20/11/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madreApellidoPaterno varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madreApellidoMaterno varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madrePrimerNombre varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madreSegundoNombre varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add NroOrdenHijo int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add madreTipoDocumento int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  Sector varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes add  Sectorista int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Pacientes ALTER COLUMN DireccionDomicilio VARCHAR(100) NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    '
    DoEvents
    ProgressBar1.Value = 56
    Me.Refresh
    txtTablaProceso.Text = "Medicos"
    lcSql = "ALTER TABLE Medicos add  idColegioHIS varchar(2)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 58
    Me.Refresh
    txtTablaProceso.Text = "His_situacio"
    lcSql = "drop table his_situacio"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE TABLE [HIS_situacio] (" & _
            "    [IdHisSituacio] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [valores] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [descripcio] [char] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [codigo] [numeric](4, 0) NOT NULL ," & _
            "    [est] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_HISsituacio] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisSituacio]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table His_situacio alter column valores char(5)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table His_situacio alter column codigo [numeric](5, 0) NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_situacio order by IdHisSituacio"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_situacio"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           oRsTmpOpc.AddNew
           oRsTmpOpc.Fields!valores = IIf(IsNull(oRsTmpOpc1.Fields!valores), "", oRsTmpOpc1.Fields!valores)
           oRsTmpOpc.Fields!descripcio = oRsTmpOpc1.Fields!descripcio
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!est = oRsTmpOpc1.Fields!est
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 59
    Me.Refresh
    txtTablaProceso.Text = "TipoFinanciador"
    '****debb-22/06/2015
    lcSql = "CREATE TABLE [dbo].[TipoFinanciador] (" & _
            "    [idTipoFinanciador] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [nombre] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [denominacion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "   [codigo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
            "    CONSTRAINT [PK_TipoFinanciador] PRIMARY KEY  CLUSTERED " & _
            "    (" & _
            "        [idTipoFinanciador]" & _
            "    )  ON [PRIMARY]" & _
            ") ON [PRIMARY]"
'    lcSql = "CREATE TABLE [dbo].[TipoFinanciador] (" & _
'            "    [idTipoFinanciador] [int] NOT NULL ," & _
'            "    [nombre] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
'            "    [denominacion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
'            "   [codigo] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
'            "    CONSTRAINT [PK_TipoFinanciador] PRIMARY KEY  CLUSTERED " & _
'            "    (" & _
'            "        [idTipoFinanciador]" & _
'            "    )  ON [PRIMARY]" & _
'            ") ON [PRIMARY]"
     '
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from TipoFinanciador"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
       oRsTmpOpc1.MoveFirst
       Do While Not oRsTmpOpc1.EOF
          lcSql = "select * from TipoFinanciador where idTipoFinanciador=" & oRsTmpOpc1.Fields!idTipoFinanciador
          If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
          oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
          If oRsTmpOpc.RecordCount = 0 Then
             oRsTmpOpc.AddNew
'             oRsTmpOpc.Fields!idTipoFinanciador = oRsTmpOpc1.Fields!idTipoFinanciador
          End If
          'oRsTmpOpc.Fields!idTipoFinanciador = oRsTmpOpc1.Fields!idTipoFinanciador    'debb-02/07/2015
          oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
          oRsTmpOpc.Fields!denominacion = oRsTmpOpc1.Fields!denominacion
          oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
          oRsTmpOpc.Update
          oRsTmpOpc1.MoveNext
       Loop
    End If
    
    '
    DoEvents
    ProgressBar1.Value = 59
    Me.Refresh
    txtTablaProceso.Text = "FuentesFinanciamiento"
    lcSql = "ALTER TABLE FuentesFinanciamiento add  CodigoHIS varchar(2)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[FuentesFinanciamiento]" & _
            " ADD [idTipoFinanciador] INT CONSTRAINT [DF__FuentesFi__idTip__61189672] DEFAULT ((1)) NOT NULL," & _
            " [codigo] CHAR (11) CONSTRAINT [DF__FuentesFi__codig__620CBAAB] DEFAULT ('00000000000') NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[FuentesFinanciamiento] WITH NOCHECK ADD " & _
            " CONSTRAINT [FK_FuentesFinanciamiento_TipoFinanciador] FOREIGN KEY " & _
            " (" & _
            "     [idTipoFinanciador]" & _
            " ) REFERENCES [dbo].[TipoFinanciador] (" & _
            "    [idTipoFinanciador]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 60
    Me.Refresh
    txtTablaProceso.Text = "FarmSaldoMensualDetallado"
    lcSql = "alter table FarmSaldoMensualDetallado drop CONSTRAINT PK_FarmSaldoMensualDetallado"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[FarmSaldoMensualDetallado] ADD " & _
            "    CONSTRAINT [PK_FarmSaldoMensualDetallado] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idAlmacen]," & _
            "        [idProducto]," & _
            "        [Lote]," & _
            "        [FechaVencimiento]," & _
            "        [IdTipoSalidaBienInsumo]," & _
            "        [SaldoFecha]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub


Sub MigraUltimaVersion_TablaSIGH_Parte2(oConexHBT As Connection, oConexODBC As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String
    On Error GoTo errMg
    '
    
    DoEvents
    ProgressBar1.Value = 61
    Me.Refresh
    txtTablaProceso.Text = "farmRelmod"
'    lcSql = "drop table farmRelmod"
'    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
'    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE TABLE [farmRelMod] (" & _
            "    [TipoAlmacen] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [TipoMov] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [TipoSuministro] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [ConceptoCodigo] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [DocumentoId] [int] NOT NULL ," & _
            "    [DocumentoCodigo] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [DocumentoEsAutomatico] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [DocumentoUltimoNumero] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [NiDocumentoOrigenId] [int] NOT NULL ," & _
            "    [NiDocumentoOrigenCodigo] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [NiDocumentoOrigenEsAutomatico] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [NiDocumentoOrigenUltimoNumero] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [NiFiltroAlmacenOrigen] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NiEsCompra] [bit] NOT NULL ," & _
            "    [NiEsDevolucionPaciente] [bit] NOT NULL ," & _
            "    [NsFiltroAlmacenDestino] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [TipoPrecioParaNiNs] [int] NOT NULL ," & _
            "    [MuestraLoteParaDespachoNS] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NiFiltroAlmacenOrigenCS] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NsFiltroAlmacenDestinoCS] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    CONSTRAINT [PK_farmRelMod1] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [TipoAlmacen],"
lcSql = lcSql & "        [TipoMov]," & _
            "        [TipoSuministro]," & _
            "        [ConceptoCodigo]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
'    lcSql = "update farmRelmod set DocumentoCodigo='14',documentoId=15  WHERE     (TipoSuministro = '02') AND (TipoMov = 'E') AND (ConceptoCodigo = '03')"   '06/07/2012
'    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
'    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
'    lcSql = "select * from farmRelmod order by TipoAlmacen,TipoMov,TipoSuministro,ConceptoCodigo" 'Actualizado 07102014
    lcSql = "select * from farmRelmod"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmRelmod"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
'              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Filter = "TipoAlmacen='" & oRsTmpOpc1.Fields!TipoAlmacen & "' and TipoMov='" & oRsTmpOpc1.Fields!TipoMov & "' and " & _
                              "TipoSuministro='" & oRsTmpOpc1.Fields!TipoSuministro & "' and ConceptoCodigo='" & oRsTmpOpc1.Fields!ConceptoCodigo & "'"
               If Not (oRsTmpOpc.EOF = True And oRsTmpOpc.BOF = True) Then
                 lbNuevoRegistro = False
                 oRsTmpOpc.MoveFirst
              End If
           End If
           
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!TipoAlmacen = oRsTmpOpc1.Fields!TipoAlmacen
                oRsTmpOpc.Fields!TipoMov = oRsTmpOpc1.Fields!TipoMov
                oRsTmpOpc.Fields!TipoSuministro = oRsTmpOpc1.Fields!TipoSuministro
                oRsTmpOpc.Fields!ConceptoCodigo = oRsTmpOpc1.Fields!ConceptoCodigo
                oRsTmpOpc.Fields!DocumentoUltimoNumero = oRsTmpOpc1.Fields!DocumentoUltimoNumero
           End If
                oRsTmpOpc.Fields!DocumentoId = oRsTmpOpc1.Fields!DocumentoId
                oRsTmpOpc.Fields!DocumentoCodigo = oRsTmpOpc1.Fields!DocumentoCodigo
                oRsTmpOpc.Fields!DocumentoEsAutomatico = oRsTmpOpc1.Fields!DocumentoEsAutomatico
                oRsTmpOpc.Fields!NiDocumentoOrigenId = oRsTmpOpc1.Fields!NiDocumentoOrigenId
                oRsTmpOpc.Fields!NiDocumentoOrigenCodigo = oRsTmpOpc1.Fields!NiDocumentoOrigenCodigo
                oRsTmpOpc.Fields!NiDocumentoOrigenEsAutomatico = oRsTmpOpc1.Fields!NiDocumentoOrigenEsAutomatico
                oRsTmpOpc.Fields!NiDocumentoOrigenUltimoNumero = oRsTmpOpc1.Fields!NiDocumentoOrigenUltimoNumero
                oRsTmpOpc.Fields!NiFiltroAlmacenOrigen = oRsTmpOpc1.Fields!NiFiltroAlmacenOrigen
                oRsTmpOpc.Fields!NiEsCompra = oRsTmpOpc1.Fields!NiEsCompra
                oRsTmpOpc.Fields!NiEsDevolucionPaciente = oRsTmpOpc1.Fields!NiEsDevolucionPaciente
                oRsTmpOpc.Fields!NsFiltroAlmacenDestino = oRsTmpOpc1.Fields!NsFiltroAlmacenDestino
                oRsTmpOpc.Fields!TipoPrecioParaNiNs = oRsTmpOpc1.Fields!TipoPrecioParaNiNs
                oRsTmpOpc.Fields!MuestraLoteParaDespachoNS = oRsTmpOpc1.Fields!MuestraLoteParaDespachoNS
                oRsTmpOpc.Fields!NiFiltroAlmacenOrigenCS = oRsTmpOpc1.Fields!NiFiltroAlmacenOrigenCS
                oRsTmpOpc.Fields!NsFiltroAlmacenDestinoCS = oRsTmpOpc1.Fields!NsFiltroAlmacenDestinoCS
                oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 63
    Me.Refresh
    txtTablaProceso.Text = "farmAlmacen"
    lcSql = "ALTER TABLE farmAlmacen add  codigoSISMED varchar(11)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen add  regenerarDias varchar(7) null"       'DEBB2014a
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen add  regenerarHora varchar(5) null"       'DEBB2014a
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen add  regenerarEstado varchar(7) null"       'DEBB2014a
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 64
    Me.Refresh
    txtTablaProceso.Text = "FacturacionCatalogoPaquetes"
    lcSql = "alter table FacturacionCatalogoPaquetes drop CONSTRAINT FK_FacturacionCatalogoPaquetes_FactCatalogoServicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 65
    Me.Refresh
    txtTablaProceso.Text = "Empleados"
    lcSql = "ALTER TABLE Empleados add  FechaNacimiento datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  idTipoDestacado int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  IdEstablecimientoExterno int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  HisCodigoDigitador varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  ReniecAutorizado bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE empleados ALTER COLUMN dni CHAR(20) not NULL"     '14/05/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  idTipoDocumento int null"          '14/05/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update empleados set idTipoDocumento=1 where idTipoDocumento is null"  '14/05/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  idSupervisor int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Empleados add  esActivo bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update empleados set esactivo=1 where esactivo is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 66
    Me.Refresh
    txtTablaProceso.Text = "AtencionesDatosAdicionales"
    lcSql = "CREATE TABLE [AtencionesDatosAdicionales] (" & _
            "    [idAtencion] [int] NOT NULL ," & _
            "    [DireccionDomicilio] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NombreAcompaniante] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Observacion] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ProximaCita] [datetime] NULL ," & _
            "    [NumeroDeHijos] [int] NULL ," & _
            "    CONSTRAINT [PK_AtencionesDatosAdicionales] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idAtencion]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table AtencionesDatosAdicionales alter column NombreAcompaniante varchar(100)"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  ProximaCita datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  NumeroDeHijos int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  idSiasis int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  FuaCodigoPrestacion varchar(3) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  SisCodigo varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN DireccionDomicilio VARCHAR(100) NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ADD SeImprimioFicha BIT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    '
    DoEvents
    ProgressBar1.Value = 67
    Me.Refresh
    txtTablaProceso.Text = "Atenciones"
    lcSql = "alter table atenciones drop CONSTRAINT FK_Atenciones_AtencionesDatosAdicionales"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Atenciones drop column Observacion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Atenciones drop column DireccionDomicilio"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table Atenciones drop column NombreAcompaniante"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 68
    Me.Refresh
    txtTablaProceso.Text = "servicios"
    lcSql = "ALTER TABLE servicios ALTER COLUMN CodigoServicioHIS VARCHAR(6) NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  EsObservacionEmergencia bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  EsObservacionEmergencia bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  UsaModuloNinoSano bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  UsaModuloMaterno bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  UsaGalenHos bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add TipoEdad int null"      'debb-20/11/2013
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Servicios add  UsaFUA bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 69
    Me.Refresh
    txtTablaProceso.Text = "SunasaTiposOperacion"
    lcSql = "drop index SunasaTiposOperacion.PK_TiposSUNASAoperacion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "    ALTER TABLE [dbo].[SunasaTiposOperacion] ADD " & _
             "   CONSTRAINT [PK_TiposSUNASAoperacion] PRIMARY KEY  CLUSTERED" & _
             "   (" & _
             "       [idOperacion]" & _
             "   )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 70
    Me.Refresh
    txtTablaProceso.Text = "SunasaTiposRegimen"
    lcSql = "drop index SunasaTiposRegimen.PK_TiposSUNASAregimen"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "    ALTER TABLE [dbo].[SunasaTiposRegimen] ADD " & _
             "   CONSTRAINT [PK_TiposSUNASAregimen] PRIMARY KEY  CLUSTERED" & _
             "   (" & _
             "       [idRegimen]" & _
             "   )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 71
    Me.Refresh
    txtTablaProceso.Text = "SunasaTiposParentesco"
    lcSql = "drop index SunasaTiposParentesco.PK_TiposSUNASAparentesco"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "    ALTER TABLE [dbo].[SunasaTiposParentesco] ADD " & _
             "   CONSTRAINT [PK_TiposSUNASAparentesco] PRIMARY KEY  CLUSTERED" & _
             "   (" & _
             "       [idParentesco]" & _
             "   )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 72
    Me.Refresh
    txtTablaProceso.Text = "SunasaTiposAfiliacion"
    lcSql = "CREATE TABLE [SunasaTiposAfiliacion] (" & _
            "    [idAfiliacion] [int] NOT NULL ," & _
            "    [Afiliacion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_TiposSUNASAafiliacion] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdAfiliacion]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    '
    DoEvents
    ProgressBar1.Value = 73
    Me.Refresh
    txtTablaProceso.Text = "SunasaPacientesHistoricos"
    lcSql = "CREATE TABLE [SunasaPacientesHistoricos] (" & _
            "    [idSunasaPacienteHistorico] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [CodigoIAFA] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idPaisTitular] [int] NULL ," & _
            "    [idTipoDocumentoTitular] [int] NULL ," & _
            "    [NroDocumentoTitular] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ApellidoCasada] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ValidacionRegIdentidad] [bit] NULL ," & _
            "    [NroCarnetIdentidad] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [EstadoDelSeguro] [int] NULL ," & _
            "    [IdAfiliacion] [int] NULL ," & _
            "    [ProductoYplan] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [FechaInicioAfiliacion] [datetime] NULL ," & _
            "    [FechaFinalAfiliacion] [datetime] NULL ," & _
            "    [idRegimen] [int] NULL ," & _
            "    [CodigoEstablecimientoIAFA] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [CodigoEstablecimientoRENAES] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idParentesco] [int] NULL ," & _
            "    [RUCempleador] [varchar] (11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [AnteriorIdTipoDocumentoAsegurado] [int] NULL ," & _
            "    [AnteriorNroDocumentoAsegurado] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [DNIusarioOperacion] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idOperacion] [int] NULL ,"
    lcSql = lcSql & "[FechaEnvio] [datetime] NULL ," & _
            "    [SisSepelioParienteEncargado] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [SisSepelioDni] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [SisSepelioFnacimiento] [datetime] NULL ," & _
            "    [SisSepelioSexo] [int] NULL ," & _
            "    [SisNroAfiliacion] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [YaNoTieneSeguro] [bit] NULL ," & _
            "    CONSTRAINT [PK_PacienteSunasa] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idSunasaPacienteHistorico]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_PacientesSunasa_Pacientes] FOREIGN KEY" & _
            "    (" & _
            "        [idPaciente]" & _
            "    ) REFERENCES [Pacientes] (" & _
            "        [idPaciente]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposDocIdentidad] FOREIGN KEY" & _
            "    (" & _
            "        [idTipoDocumentoTitular]" & _
            "    ) REFERENCES [TiposDocIdentidad] (" & _
            "        [IdDocIdentidad]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposDocIdentidad1] FOREIGN KEY"
    lcSql = lcSql & "    (" & _
            "        [AnteriorIdTipoDocumentoAsegurado]" & _
            "    ) REFERENCES [TiposDocIdentidad] (" & _
            "        [IdDocIdentidad]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposSUNASAafiliacion] FOREIGN KEY" & _
            "    (" & _
            "        [IdAfiliacion]" & _
            "    ) REFERENCES [SunasaTiposAfiliacion] (" & _
            "        [IdAfiliacion]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposSUNASAoperacion] FOREIGN KEY" & _
            "    (" & _
            "        [idOperacion]" & _
            "    ) REFERENCES [SunasaTiposOperacion] (" & _
            "        [idOperacion]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposSUNASAparentesco] FOREIGN KEY" & _
            "    (" & _
            "        [idParentesco]" & _
            "    ) REFERENCES [SunasaTiposParentesco] (" & _
            "        [idParentesco]" & _
            "    )," & _
            "    CONSTRAINT [FK_PacientesSunasa_TiposSUNASAregimen] FOREIGN KEY" & _
            "    ("
    lcSql = lcSql & "       [idRegimen]" & _
            "    ) REFERENCES [SunasaTiposRegimen] (" & _
            "        [idRegimen]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table SunasaPacientesHistoricos drop CONSTRAINT FK_Atenciones_SunasaPacientesHistoricos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    '
    DoEvents
    ProgressBar1.Value = 74
    Me.Refresh
    txtTablaProceso.Text = "FactOrdenServicioPagos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE FactOrdenServicioPagos add  idUsuarioExonera int null"  'debb-20/01/2012
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    Set oRsTmpOpc = Nothing
    lcSql = "SELECT     dbo.Auditoria.Tabla, dbo.CajaComprobantesPago.IdComprobantePago, dbo.Auditoria.IdEmpleado" & _
            " FROM         dbo.Auditoria FULL OUTER JOIN" & _
            "                      dbo.CajaComprobantesPago ON dbo.Auditoria.IdRegistro = dbo.CajaComprobantesPago.IdCuentaAtencion" & _
            " WHERE     (dbo.Auditoria.Tabla = 'Modificó SERV.SOCIAL') AND (dbo.CajaComprobantesPago.IdComprobantePago > 0)"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.RecordCount > 0 Then
       oRsTmpOpc.MoveFirst
       Do While Not oRsTmpOpc.EOF
            lcSql = "update FactOrdenServicioPagos set idUsuarioExonera=" & oRsTmpOpc.Fields!IdEmpleado & _
                   " where ImporteExonerado>0 and idComprobantePago=" & oRsTmpOpc.Fields!idComprobantePago
            If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
            oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            oRsTmpOpc.MoveNext
       Loop
    End If
    '
    DoEvents
    ProgressBar1.Value = 75
    Me.Refresh
    txtTablaProceso.Text = "FactOrdenesBienes"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE FactOrdenesBienes add  idUsuarioExonera int null"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "SELECT     dbo.Auditoria.Tabla, dbo.CajaComprobantesPago.IdComprobantePago, dbo.Auditoria.IdEmpleado" & _
            " FROM         dbo.Auditoria FULL OUTER JOIN" & _
            "                      dbo.CajaComprobantesPago ON dbo.Auditoria.IdRegistro = dbo.CajaComprobantesPago.IdCuentaAtencion" & _
            " WHERE     (dbo.Auditoria.Tabla = 'Modificó SERV.SOCIAL') AND (dbo.CajaComprobantesPago.IdComprobantePago > 0)"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.RecordCount > 0 Then
       oRsTmpOpc.MoveFirst
       Do While Not oRsTmpOpc.EOF
            lcSql = "update FactOrdenesBienes set idUsuarioExonera=" & oRsTmpOpc.Fields!IdEmpleado & _
                   " where ImporteExonerado>0 and idComprobantePago=" & oRsTmpOpc.Fields!idComprobantePago
            If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
            oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            oRsTmpOpc.MoveNext
       Loop
    End If

    '
    DoEvents
    ProgressBar1.Value = 76
    Me.Refresh
    txtTablaProceso.Text = "RecetaEstados"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE TABLE [RecetaEstados] (" & _
            "    [idEstado] [int] NOT NULL ," & _
            "    [Estado] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    CONSTRAINT [PK_RecetaEstados] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idEstado]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaEstados"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaEstados"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idEstado=" & oRsTmpOpc1.Fields!idEstado
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstado = oRsTmpOpc1.Fields!idEstado
                oRsTmpOpc.Fields!estado = oRsTmpOpc1.Fields!estado
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 77
    Me.Refresh
    txtTablaProceso.Text = "RecetaCabecera"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE TABLE [RecetaCabecera] (" & _
            "    [idReceta] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdPuntoCarga] [int] NOT NULL ," & _
            "    [FechaReceta] [datetime] NOT NULL ," & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [idServicioReceta] [int] NOT NULL ," & _
            "    [idEstado] [int] NOT NULL ," & _
            "    [DocumentoDespacho] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idComprobantePago] [int] NULL ," & _
            "    CONSTRAINT [PK_RecetaCabecera] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdReceta]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_RecetaCabecera_RecetaEstados] FOREIGN KEY" & _
            "    (" & _
            "        [idEstado]" & _
            "    ) REFERENCES [RecetaEstados] (" & _
            "        [idEstado]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE RecetaCabecera add  idMedicoReceta int null"           '23/5/12
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE  INDEX [IX_RecetaCabecera_1] ON [dbo].[RecetaCabecera]([idComprobantePago]) ON [PRIMARY]"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE  INDEX [IX_RecetaCabecera] ON [dbo].[RecetaCabecera]([idCuentaAtencion]) ON [PRIMARY]"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE recetaCabecera add  fechaVigencia datetime null"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 78
    Me.Refresh
    txtTablaProceso.Text = "RecetaDetalle"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "CREATE TABLE [RecetaDetalle] (" & _
            "    [idReceta] [int] NOT NULL ," & _
            "    [idItem] [int] NOT NULL ," & _
            "    [CantidadPedida] [int] NOT NULL ," & _
            "    [Precio] [money] NOT NULL ," & _
            "    [Total] [money] NOT NULL ," & _
            "    [SaldoEnRegistroReceta] [int] NULL ," & _
            "    [SaldoEnDespachoReceta] [int] NULL ," & _
            "    [CantidadDespachada] [int] NULL ," & _
            "    CONSTRAINT [PK_RecetaDetalle] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idReceta]," & _
            "        [idItem]" & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_RecetaDetalle_RecetaCabecera] FOREIGN KEY" & _
            "    (" & _
            "        [IdReceta]" & _
            "    ) REFERENCES [RecetaCabecera] (" & _
            "        [IdReceta]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE RecetaDetalle add  idDosisRecetada int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE RecetaDetalle add  idEstadoDetalle int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE RecetaDetalle add  MotivoAnulacionMedico varchar(300) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE RecetaDetalle add  observaciones varchar(300) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 79
    Me.Refresh
    txtTablaProceso.Text = "HIS_estab"
    lcSql = "CREATE TABLE [dbo].[HIS_ESTAB] (" & _
            "    [COD_ESTAB] [nvarchar] (9) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [DESC_ESTAB] [nvarchar] (44) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_2000] [float] NULL ," & _
            "    [TIPOESTAB] [nvarchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_DPTO] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_PROV] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_DIST] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_DISA] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_RED] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [COD_MIC] [nvarchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_estab"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           
            DoEvents
            Me.Refresh
            txtTablaProceso.Text = oRsTmpOpc1.Fields!cod_estab
           
           lcSql = "select * from HIS_estab where cod_estab='" & oRsTmpOpc1.Fields!cod_estab & "'"
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount > 0 Then
'              oRsTmpOpc.MoveFirst
'              oRsTmpOpc.Find "cod_estab='" & oRsTmpOpc1.Fields!cod_estab & "'"
'              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
'              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!cod_estab = oRsTmpOpc1.Fields!cod_estab
                oRsTmpOpc.Fields!desc_estab = oRsTmpOpc1.Fields!desc_estab
                oRsTmpOpc.Fields!cod_2000 = oRsTmpOpc1.Fields!cod_2000
                oRsTmpOpc.Fields!TIPOESTAB = oRsTmpOpc1.Fields!TIPOESTAB
                oRsTmpOpc.Fields!COD_DPTO = oRsTmpOpc1.Fields!COD_DPTO
                oRsTmpOpc.Fields!COD_PROV = oRsTmpOpc1.Fields!COD_PROV
                oRsTmpOpc.Fields!COD_DIST = oRsTmpOpc1.Fields!COD_DIST
                oRsTmpOpc.Fields!COD_DISA = oRsTmpOpc1.Fields!COD_DISA
                oRsTmpOpc.Fields!COD_RED = oRsTmpOpc1.Fields!COD_RED
                oRsTmpOpc.Fields!COD_MIC = oRsTmpOpc1.Fields!COD_MIC
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 80
    Me.Refresh
    txtTablaProceso.Text = "Departamentos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE Departamentos add  idReniec int null"           '13/6/12
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Departamentos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ReniecU where dptoI=" & oRsTmpOpc1.Fields!IdDepartamento
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount > 0 Then
                oRsTmpOpc1.Fields!idReniec = Val(oRsTmpOpc.Fields!DptoR)
                oRsTmpOpc1.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 81
    Me.Refresh
    txtTablaProceso.Text = "atencionesNacimientos"
    lcSql = "ALTER TABLE atencionesNacimientos add  apgar_1 int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesNacimientos add  apgar_5 int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE atencionesNacimientos add  ClamplajeFecha datetime null"    'debb-20/11/13
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE atencionesNacimientos add  NroOrdenHijoEnParto int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE atencionesNacimientos add  NroOrdenHijo int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 82
    Me.Refresh
    txtTablaProceso.Text = "labResultado"
    lcSql = "ALTER TABLE labResultado add  fecha datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 83
    Me.Refresh
    txtTablaProceso.Text = "TiposEstadoCivil"
    lcSql = "ALTER TABLE TiposEstadoCivil add  sip2000 varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposEstadoCivil"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposEstadoCivil"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdEstadoCivil=" & oRsTmpOpc1.Fields!IdEstadoCivil
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdEstadoCivil = oRsTmpOpc1.Fields!IdEstadoCivil
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Fields!lolcli = oRsTmpOpc1.Fields!lolcli
           End If
           oRsTmpOpc.Fields!sip2000 = oRsTmpOpc1.Fields!sip2000
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 84
    Me.Refresh
    txtTablaProceso.Text = "TiposGradoInstruccion"
    lcSql = "ALTER TABLE TiposGradoInstruccion add  sip2000 varchar(2) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposGradoInstruccion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposGradoInstruccion"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdGradoInstruccion=" & oRsTmpOpc1.Fields!IdGradoInstruccion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdGradoInstruccion = oRsTmpOpc1.Fields!IdGradoInstruccion
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           End If
           oRsTmpOpc.Fields!sip2000 = oRsTmpOpc1.Fields!sip2000
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 85
    Me.Refresh
    txtTablaProceso.Text = "CajaCaja"
    lcSql = "ALTER TABLE CajaCaja add  ImpresoraDefault varchar(50) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE CajaCaja add  Impresora2 varchar(50) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE CajaCaja add  idTipoComprobante int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update CajaCaja set  idTipoComprobante=3 where idTipoComprobante is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    lcSql = "update CajaTiposComprobante set Descripcion='Recibo' where idTipoComprobante=1"   'debb-28-9-12
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update CajaTiposComprobante set Descripcion='Boleta' where idTipoComprobante=3"   'debb-28-9-12
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '**** Programa: se agrego dos campos a la tabla
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014
    lcSql = "ALTER TABLE CajaCaja add  SerieImpresoraDefault varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE CajaCaja add  SerieImpresora2 varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 86
    Me.Refresh
    txtTablaProceso.Text = "FarmComponente"
    lcSql = "select * from FarmComponente"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from FarmComponente"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idComponente=" & oRsTmpOpc1.Fields!idComponente
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!idComponente = oRsTmpOpc1.Fields!idComponente
                oRsTmpOpc.Fields!Componente = oRsTmpOpc1.Fields!Componente
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 87
    Me.Refresh
    txtTablaProceso.Text = "farmComponenteSub"
    lcSql = "ALTER TABLE farmComponenteSub add  componente varchar(2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmComponenteSub add  TipoProductoSismed varchar(1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update farmComponenteSub set  TipoProductoSismed='S' where idSubComponente=9"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmComponenteSub"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
            lcSql = "update farmComponenteSub set componente='" & oRsTmpOpc1.Fields!Componente & "' where idSubcomponente=" & oRsTmpOpc1.Fields!idSubComponente & " and idComponente=" & oRsTmpOpc1.Fields!idComponente
            If oRsTmpOpc.State = 1 Then oRsTmpOpc1.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    lcSql = "select * from FarmComponenteSub"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from FarmComponenteSub"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idSubComponente=" & oRsTmpOpc1.Fields!idSubComponente
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idSubComponente = oRsTmpOpc1.Fields!idSubComponente
                oRsTmpOpc.Fields!idComponente = oRsTmpOpc1.Fields!idComponente
                oRsTmpOpc.Fields!Componente = oRsTmpOpc1.Fields!Componente
                oRsTmpOpc.Fields!subComponente = oRsTmpOpc1.Fields!subComponente
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Fields!Diagnostico = oRsTmpOpc1.Fields!Diagnostico
                oRsTmpOpc.Fields!TipoProductoSismed = oRsTmpOpc1.Fields!TipoProductoSismed
                oRsTmpOpc.Update
           Else
                oRsTmpOpc.Fields!TipoProductoSismed = oRsTmpOpc1.Fields!TipoProductoSismed
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    '
    DoEvents
    ProgressBar1.Value = 88
    Me.Refresh
    txtTablaProceso.Text = "FacturacionCuentasAtencionPtos"
    lcSql = "CREATE TABLE [dbo].[FacturacionCuentasAtencionPtos] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [idPuntoCarga] [int] NOT NULL ," & _
            "    [TotalConsumos] [money] NOT NULL ," & _
            "    [TotalPagos] [money] NOT NULL ," & _
            "    [TotalPagosReembolso]  [Money] NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[FacturacionCuentasAtencionPtos] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_AtencionesFacturacionResumen] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idCuentaAtencion]," & _
            "        [IdPuntoCarga]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[FacturacionCuentasAtencionPtos] ADD " & _
            " CONSTRAINT [DF_FacturacionCuentasAtencionPtos_TotalConsumos] DEFAULT (0) FOR [TotalConsumos]," & _
            " CONSTRAINT [DF_FacturacionCuentasAtencionPtos_TotalPagos0] DEFAULT (0) FOR [TotalPagos]," & _
            " CONSTRAINT [DF_FacturacionCuentasAtencionPtos_TotalPagosReembolso] DEFAULT (0) FOR [TotalPagosReembolso]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 89
    Me.Refresh
    txtTablaProceso.Text = "TiposTarifa"
    lcSql = "CREATE TABLE [dbo].[TiposTarifa] (" & _
            "    [idTipoTarifa] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [codigo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [TipoTarifa] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & _
            "    [EsFarmacia] [bit] NULL " & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[TiposTarifa] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_TiposTarifa] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idTipoTarifa]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposTarifa"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposTarifa"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "TipoTarifa='" & oRsTmpOpc1.Fields!TipoTarifa & "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
'                oRsTmpOpc.Fields!idTipoTarifa = oRsTmpOpc1.Fields!idTipoTarifa
                oRsTmpOpc.Fields!TipoTarifa = oRsTmpOpc1.Fields!TipoTarifa
                oRsTmpOpc.Fields!EsFarmacia = oRsTmpOpc1.Fields!EsFarmacia
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    
'    DoEvents
'    ProgressBar1.Value = 90
'    Me.Refresh
'    txtTablaProceso.Text = "TiposTarifaCpt"
    lcSql = "CREATE TABLE [dbo].[TiposTarifaCpt] (" & _
            "    [idTipoTarifa] [int] NOT NULL ," & _
            "    [idProductoCpt] [int] NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[TiposTarifaCpt] ADD " & _
            "    CONSTRAINT [FK_TiposTarifaCpt_TiposTarifa] FOREIGN KEY" & _
            "    (" & _
            "        [idTipoTarifa]" & _
            "    ) REFERENCES [dbo].[TiposTarifa] (" & _
            "        [idTipoTarifa]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposTarifaCpt"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposTarifaCpt"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
            DoEvents
            txtTablaProceso.Text = oRsTmpOpc1.Fields!idTipoTarifa & " - " & oRsTmpOpc1.Fields!idProductoCpt
            Me.Refresh
    
           lbNuevoRegistro = True
           lcSql = "select * from TiposTarifaCpt where idTipoTarifa=" & oRsTmpOpc1.Fields!idTipoTarifa & " and idProductoCpt=" & oRsTmpOpc1.Fields!idProductoCpt
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idTipoTarifa = oRsTmpOpc1.Fields!idTipoTarifa
                oRsTmpOpc.Fields!idProductoCpt = oRsTmpOpc1.Fields!idProductoCpt
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 91
    Me.Refresh
    txtTablaProceso.Text = "HIS_TipoAtencion"
    lcSql = "CREATE TABLE [dbo].[HIS_TipoAtencion] (" & _
            "    [IdHisTipoAtencion] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [Codigo] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_TipoAtencion] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_TipoAtencion__30DB7AB8] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisTipoAtencion]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_TipoAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_TipoAtencion"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdHisTipoAtencion=" & oRsTmpOpc1.Fields!IdHisTipoAtencion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdHisTipoAtencion = oRsTmpOpc1.Fields!IdHisTipoAtencion
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 92
    Me.Refresh
    txtTablaProceso.Text = "HIS_CodigosActividades"
    lcSql = "CREATE TABLE [dbo].[HIS_CodigosActividades] (" & _
            "    [IdHisCodActvidad] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdTipoAtencion] [int] NULL ," & _
            "    [CodigoActividad] [char] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Descripcion] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_CodigosActividades] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK__HIS_CodigosActiv__3A64E4F2] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [IdHisCodActvidad]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_CodigosActividades] ADD " & _
            "    CONSTRAINT [IdTipoAtencion_IdTipoAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoAtencion]" & _
            "    ) REFERENCES [dbo].[HIS_TipoAtencion] (" & _
            "        [IdHisTipoAtencion]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_CodigosActividades"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_CodigosActividades"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idHisCodActvidad=" & oRsTmpOpc1.Fields!IdHisCodActvidad
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!idHisCodActvidad = oRsTmpOpc1.Fields!idHisCodActvidad
                oRsTmpOpc.Fields!IdTipoAtencion = oRsTmpOpc1.Fields!IdTipoAtencion
                oRsTmpOpc.Fields!CodigoActividad = oRsTmpOpc1.Fields!CodigoActividad
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 93
    Me.Refresh
    txtTablaProceso.Text = "HIS_ServEstablecimiento"
    lcSql = "CREATE TABLE [dbo].[HIS_ServEstablecimiento] (" & _
            "    [IdHisServEstablecimiento] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [IdServicio] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_ServEstablec__387C9C80] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisServEstablecimiento]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ServEstablecimiento] ADD " & _
            "    CONSTRAINT [HIS_ServEstablecimiento_IdEstablecimiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )," & _
            "    CONSTRAINT [HIS_ServEstablecimiento_IdServicio] FOREIGN KEY" & _
            "    (" & _
            "        [IdServicio]" & _
            "    ) REFERENCES [dbo].[Servicios] (" & _
            "        [IdServicio]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 94
    Me.Refresh
    txtTablaProceso.Text = "HIS_Turnos"
    lcSql = "CREATE TABLE [dbo].[HIS_Turnos] (" & _
            "    [IdHisTurno] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [Descripcion] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Turnos] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_Turnos__3E3575D6] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisTurno]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_Turnos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_Turnos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdHisTurno=" & oRsTmpOpc1.Fields!IdHisTurno
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdHisTurno = oRsTmpOpc1.Fields!IdHisTurno
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 95
    Me.Refresh
    txtTablaProceso.Text = "HIS_Cabecera"
    lcSql = "CREATE TABLE [dbo].[HIS_Cabecera] (" & _
            "    [IdHisCabecera] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdHisLote] [int] NULL ," & _
            "    [NroHojaHis] [int] NULL ," & _
            "    [NroFormato] [int] NOT NULL ," & _
            "    [IdTurno] [int] NOT NULL ," & _
            "    [IdUsuario] [int] NULL ," & _
            "    [IdEstadoHis] [int] NULL ," & _
            "    [IdMedico] [int] NULL ," & _
            "    [IdServEstablecimiento] [int] NULL ," & _
            "    [FechaCreacion] [datetime] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_Cabecera__2D0AE9D4] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisCabecera]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD " & _
            " CONSTRAINT [DF__HIS_Cabec__IdEst__37BE4DC8] DEFAULT (1867) FOR [IdEstablecimiento]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD " & _
            " CONSTRAINT [DF__HIS_Cabec__IdSer__38B27201] DEFAULT (31) FOR [IdServicio]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    

    '
    DoEvents
    ProgressBar1.Value = 96
    Me.Refresh
    txtTablaProceso.Text = "HIS_Detalle"
    lcSql = "CREATE TABLE [dbo].[HIS_Detalle] (" & _
            "    [IdHisDetalle] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdHisCabecera] [int] NULL ," & _
            "    [IdTipoAtencion] [int] NULL ," & _
            "    [DiaAtencion] [int] NOT NULL ," & _
            "    [IdHisPaciente] [int] NULL ," & _
            "    [CodigoActividad] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdTipoFinanciamiento] [int] NULL ," & _
            "    [IdDistrito] [int] NULL ," & _
            "    [IdTipoEdad] [int] NULL ," & _
            "    [Edad] [int] NULL ," & _
            "    [Talla] [int] NULL ," & _
            "    [Peso] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdEstadoaEstablec] [int] NULL ," & _
            "    [IdEstadoaServicio] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Detalle] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_Detalle__2EF33246] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisDetalle]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table HIS_Detalle alter column peso char(50)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    '
    DoEvents
    ProgressBar1.Value = 97
    Me.Refresh
    txtTablaProceso.Text = "HIS_DetalleDiagnostico"
    lcSql = "CREATE TABLE [dbo].[HIS_DetalleDiagnostico] (" & _
            "    [IdHisDetalleDiagnostico] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdHisDetalle] [int] NULL ," & _
            "    [IdCIE] [int] NULL ," & _
            "    [IdSubClasificacionDX] [int] NULL ," & _
            "    [CodLAB] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DetalleDiagnostico] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_DetalleDiagn__34AC0B9C] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisDetalleDiagnostico]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DetalleDiagnostico] ADD " & _
            "    CONSTRAINT [HIS_DetalleDiagnostico_IdCIE] FOREIGN KEY" & _
            "    (" & _
            "        [IdCIE]" & _
            "    ) REFERENCES [dbo].[Diagnosticos] (" & _
            "        [IdDiagnostico]" & _
            "    )," & _
            "    CONSTRAINT [HIS_DetalleDiagnostico_IdHisDetalle] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisDetalle]" & _
            "    ) REFERENCES [dbo].[HIS_Detalle] (" & _
            "        [IdHisDetalle]" & _
            "    )," & _
            "    CONSTRAINT [HIS_DetalleDiagnostico_IdSubClasificacionDX] FOREIGN KEY" & _
            "    (" & _
            "        [IdSubclasificacionDx]" & _
            "    ) REFERENCES [dbo].[SubclasificacionDiagnosticos] (" & _
            "        [IdSubclasificacionDx]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 98
    Me.Refresh
    txtTablaProceso.Text = "HIS_Lotes"
    lcSql = "CREATE TABLE [dbo].[HIS_Lotes] (" & _
            "    [IdHisLote] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [Lote] [char] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NroHojas] [int] NULL ," & _
            "    [Mes] [int] NULL ," & _
            "    [Anio] [int] NULL ," & _
            "    [Cerrado] [Int] null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Lotes] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_Lotes__2B22A162] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisLote]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table HIS_Lotes alter column lote char(3)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table HIS_Lotes add idEstadoLote int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic


    '
    DoEvents
    ProgressBar1.Value = 99
    Me.Refresh
    txtTablaProceso.Text = "HIS_Paciente"
    lcSql = "CREATE TABLE [dbo].[HIS_Paciente] (" & _
            "    [IdHisPaciente] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [NroHC_FF] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Sexo] [int] NULL ," & _
            "    [IdNacionalidad] [int] NULL ," & _
            "    [NroDocIdentidad] [varchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [NroHijo] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdEtnia] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdPacienteGalenHos] [int] NULL ," & _
            "    [IdTipoDocumento] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Paciente] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_Paciente__32C3C32A] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisPaciente]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Paciente] ADD " & _
            "    CONSTRAINT [HIS_Paciente_IdEtnia] FOREIGN KEY" & _
            "    (" & _
            "        [IdEtnia]" & _
            "    ) REFERENCES [dbo].[HIS_tabetnia] (" & _
            "        [codetni]" & _
            "    )," & _
            "    CONSTRAINT [HIS_Paciente_IdNacionalidad] FOREIGN KEY" & _
            "    (" & _
            "        [IdNacionalidad]" & _
            "    ) REFERENCES [dbo].[Paises] (" & _
            "        [IdPais]" & _
            "    )," & _
            "    CONSTRAINT [HIS_Paciente_Pacientes] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisPaciente]" & _
            "    ) REFERENCES [dbo].[Pacientes] (" & _
            "        [IdPaciente]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 100
    Me.Refresh
    txtTablaProceso.Text = "HIS_ProgMedEstMR"
    lcSql = "CREATE TABLE [dbo].[HIS_ProgMedEstMR] (" & _
            "    [IdHisProgMedEstMR] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdMedico] [int] NULL ," & _
            "    [IdServicio] [int] NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [FechaProgramada] [datetime] NULL ," & _
            "    [IdTurno] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ProgMedEstMR] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_ProgMedEstMR__3C4D2D64] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisProgMedEstMR]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_ProgMedEstMR] ADD " & _
            "    CONSTRAINT [IdEstablecimiento_IdEstablecimiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )," & _
            "    CONSTRAINT [IdMedico_IdMedico] FOREIGN KEY" & _
            "    (" & _
            "        [IdMedico]" & _
            "    ) REFERENCES [dbo].[Medicos] (" & _
            "        [IdMedico]" & _
            "    )," & _
            "    CONSTRAINT [IdServicio_IdServicio] FOREIGN KEY" & _
            "    (" & _
            "        [IdServicio]" & _
            "    ) REFERENCES [dbo].[Servicios] (" & _
            "        [IdServicio]" & _
            "    )," & _
            "    CONSTRAINT [IdTurno_IIdTurno] FOREIGN KEY" & _
            "    (" & _
            "        [IdTurno]" & _
            "    ) REFERENCES [dbo].[HIS_Turnos] (" & _
            "        [IdHisTurno]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 101
    Me.Refresh
    txtTablaProceso.Text = "HIS Cabecera y Detalle - Relaciones"
    lcSql = "ALTER TABLE [dbo].[HIS_Cabecera] ADD " & _
            "    CONSTRAINT [FK_HIS_Cabecera_IdHisLote] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisLote]" & _
            "    ) REFERENCES [dbo].[HIS_Lotes] (" & _
            "        [IdHisLote]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Cabecera_IdMedico] FOREIGN KEY" & _
            "    (" & _
            "        [IdMedico]" & _
            "    ) REFERENCES [dbo].[Medicos] (" & _
            "        [IdMedico]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Cabecera_IdUsuario] FOREIGN KEY" & _
            "    (" & _
            "        [IdUsuario]" & _
            "    ) REFERENCES [dbo].[Empleados] (" & _
            "        [IdEmpleado]" & _
            "    )," & _
            "    CONSTRAINT [FK_HIS_Cabecera_ServPorEstablec] FOREIGN KEY" & _
            "    (" & _
            "        [IdServEstablecimiento]" & _
            "    ) REFERENCES [dbo].[HIS_ServEstablecimiento] (" & _
            "        [IdHisServEstablecimiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_Detalle] ADD " & _
            "    CONSTRAINT [HIS_Detalle_IdDistrito] FOREIGN KEY" & _
            "    (" & _
            "        [IdDistrito]" & _
            "    ) REFERENCES [dbo].[Distritos] (" & _
            "        [IdDistrito]" & _
            "    )," & _
            "    CONSTRAINT [HIS_Detalle_IdHisCabecera] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisCabecera]" & _
            "    ) REFERENCES [dbo].[HIS_Cabecera] (" & _
            "        [IdHisCabecera]" & _
            "    )," & _
            "    CONSTRAINT [HIS_Detalle_IdHISPaciente] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisPaciente]" & _
            "    ) REFERENCES [dbo].[HIS_Paciente] (" & _
            "        [IdHisPaciente]" & _
            "    )," & _
            "    CONSTRAINT [HIS_Detalle_IdTipoAtencion] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoAtencion]" & _
            "    ) REFERENCES [dbo].[HIS_TipoAtencion] (" & _
            "        [IdHisTipoAtencion]" & _
            "    ),"
    
   lcSql = lcSql & " CONSTRAINT [HIS_Detalle_IdTipoFinanciamiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoFinanciamiento]" & _
            "    ) REFERENCES [dbo].[TiposFinanciamiento] (" & _
            "        [IdTipoFinanciamiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 102
    Me.Refresh
    txtTablaProceso.Text = "HIS_ProgMedEstMR"
    lcSql = "CREATE TABLE [dbo].[HIS_ProgMedEstMR] (" & _
            "    [IdHisProgMedEstMR] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdMedico] [int] NULL ," & _
            "    [IdServicio] [int] NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [FechaProgramada] [datetime] NULL ," & _
            "    [IdTurno] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 103
    Me.Refresh
    txtTablaProceso.Text = "HIS_DatosEstablecimiento"
    lcSql = "CREATE TABLE [dbo].[HIS_DatosEstablecimiento] (" & _
            "    [IdDatoEstablec] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [Color] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Turnos] [int] NULL ," & _
            "    [UltimoNroFormatoHIS] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DatosEstablecimiento] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_DatosEstable__0B74EBDF] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdDatoEstablec]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_DatosEstablecimiento] ADD " & _
            "    CONSTRAINT [HIS_DatosEstablecimiento_IdDatoEstablec] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 103
    Me.Refresh
    txtTablaProceso.Text = "HIS_TipoEdad"
    lcSql = "CREATE TABLE [dbo].[HIS_TipoEdad] (" & _
            "    [IdHisTipoEdad] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [CodigoEdad] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Descripcion] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_TipoEdad] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK__HIS_TipoEdad__3694540E] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdHisTipoEdad]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_TipoEdad"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_TipoEdad"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdHisTipoEdad=" & oRsTmpOpc1.Fields!IdHisTipoEdad
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdHisTipoEdad = oRsTmpOpc1.Fields!IdHisTipoEdad
                oRsTmpOpc.Fields!CodigoEdad = oRsTmpOpc1.Fields!CodigoEdad
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
                
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    DoEvents
    ProgressBar1.Value = 104
    Me.Refresh
    txtTablaProceso.Text = "HIS_EstablecPacienteHIS"
    lcSql = "CREATE TABLE [dbo].[HIS_EstablecPacienteHIS] (" & _
            "    [IdEstablecPacienteHIS] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdEstablecimiento] [int] NULL ," & _
            "    [IdHisPaciente] [int] NULL ," & _
            "    [NroHC_FF] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_EstablecPacienteHIS] WITH NOCHECK ADD " & _
            "     PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdEstablecPacienteHIS]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[HIS_EstablecPacienteHIS] ADD " & _
            "    CONSTRAINT [His_EstablecPacienteHIS_IdEstablecimiento] FOREIGN KEY" & _
            "    (" & _
            "        [IdEstablecimiento]" & _
            "    ) REFERENCES [dbo].[Establecimientos] (" & _
            "        [IdEstablecimiento]" & _
            "    )," & _
            "    CONSTRAINT [His_EstablecPacienteHIS_IdHisPaciente] FOREIGN KEY" & _
            "    (" & _
            "        [IdHisPaciente]" & _
            "    ) REFERENCES [dbo].[HIS_Paciente] (" & _
            "        [IdHisPaciente]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    ProgressBar1.Value = 105
    Me.Refresh
    txtTablaProceso.Text = "PacientesDatosAdicionales"
    lcSql = "CREATE TABLE [dbo].[PacientesDatosAdicionales] (" & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [antecedentes] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[PacientesDatosAdicionales] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_PacientesDatosAdicionales] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idPaciente]" & _
            "   )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PacientesDatosAdicionales add  antecedAlergico varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PacientesDatosAdicionales add  antecedObstetrico varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PacientesDatosAdicionales add  antecedQuirurgico varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PacientesDatosAdicionales add  antecedFamiliar varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE PacientesDatosAdicionales add  antecedPatologico varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE dbo.PacientesDatosAdicionales ADD FNacimientoCalculada bit NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE dbo.PacientesDatosAdicionales WITH NOCHECK ADD " & _
        " CONSTRAINT[DF_PacientesDatosAdicionales_FNacimientoCalculada] DEFAULT(0) FOR FNacimientoCalculada"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    ProgressBar1.Value = 106
    Me.Refresh
    txtTablaProceso.Text = "TiposIdiomas"
    lcSql = "CREATE TABLE [dbo].[TiposIdiomas] (" & _
            "    [IdIdioma] [int] NOT NULL ," & _
            "    [Codigo] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Lengua] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposIdiomas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposIdiomas"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdIdioma=" & oRsTmpOpc1.Fields!IdIdioma
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdIdioma = oRsTmpOpc1.Fields!IdIdioma
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Fields!lengua = oRsTmpOpc1.Fields!lengua
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 107
    Me.Refresh
    txtTablaProceso.Text = "establecimientosNoMinsa"                                   '22/05/2013
    lcSql = "alter table establecimientosNoMinsa drop CONSTRAINT Distritos_EstablecimientosNoMinsa_FK1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE establecimientosNoMinsa add  codigo varchar(10) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table establecimientosNoMinsa alter column nombre varchar(150)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from establecimientosNoMinsa"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from establecimientosNoMinsa where IdEstablecimientoNoMinsa=" & oRsTmpOpc1.Fields!IdEstablecimientoNoMinsa
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           If oRsTmpOpc.RecordCount = 0 Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdEstablecimientoNoMinsa = oRsTmpOpc1.Fields!IdEstablecimientoNoMinsa
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Nombre = oRsTmpOpc1.Fields!Nombre
           oRsTmpOpc.Fields!IdDistrito = oRsTmpOpc1.Fields!IdDistrito
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 108
    Me.Refresh
    txtTablaProceso.Text = "TiposOrigenAtencion"
    lcSql = "ALTER TABLE TiposOrigenAtencion add  id_origenAseguradoSIS varchar(1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposOrigenAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposOrigenAtencion"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdOrigenAtencion=" & oRsTmpOpc1.Fields!IdOrigenAtencion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdOrigenAtencion = oRsTmpOpc1.Fields!IdOrigenAtencion
           End If
           oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!IdTipoServicio = oRsTmpOpc1.Fields!IdTipoServicio
           oRsTmpOpc.Fields!TipoServicioHosp = oRsTmpOpc1.Fields!TipoServicioHosp
           oRsTmpOpc.Fields!id_origenAseguradoSIS = oRsTmpOpc1.Fields!id_origenAseguradoSIS
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 109
    Me.Refresh
    txtTablaProceso.Text = "LabGrupos"
    lcSql = "ALTER TABLE [dbo].[labGrupos] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_labGrupos] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [IdGrupo]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE LabGrupos add  idCargo int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabGrupos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabGrupos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdGrupo=" & oRsTmpOpc1.Fields!idGrupo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idGrupo = oRsTmpOpc1.Fields!idGrupo
                oRsTmpOpc.Fields!NombreGrupo = oRsTmpOpc1.Fields!NombreGrupo
                oRsTmpOpc.Fields!SiglasGrupo = oRsTmpOpc1.Fields!SiglasGrupo
           End If
           oRsTmpOpc.Fields!idCargo = oRsTmpOpc1.Fields!idCargo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 110
    Me.Refresh
    txtTablaProceso.Text = "LabItemsGrupos"
    lcSql = "CREATE TABLE [dbo].[LabItemsGrupos] (" & _
            "    [idItemGrupo] [int] NOT NULL ," & _
            "    [Grupo] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabItemsGrupos] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_LabItemsGrupos] PRIMARY KEY  CLUSTERED" & _
            "(" & _
            "    [idItemGrupo]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItemsGrupos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItemsGrupos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idItemGrupo=" & oRsTmpOpc1.Fields!idItemGrupo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idItemGrupo = oRsTmpOpc1.Fields!idItemGrupo
                oRsTmpOpc.Fields!Grupo = oRsTmpOpc1.Fields!Grupo
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 111
    Me.Refresh
    txtTablaProceso.Text = "LabItems"
    lcSql = "CREATE TABLE [dbo].[LabItems] (" & _
            "    [idItem] [int] NOT NULL ," & _
            "    [Item] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [idProductoCpt] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabItems] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_LabItems] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idItem]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItems"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItems"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idItem=" & oRsTmpOpc1.Fields!idItem
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idItem = oRsTmpOpc1.Fields!idItem
                oRsTmpOpc.Fields!Item = oRsTmpOpc1.Fields!Item
                oRsTmpOpc.Fields!idProductoCpt = oRsTmpOpc1.Fields!idProductoCpt
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 112
    Me.Refresh
    txtTablaProceso.Text = "LabItemsCpt"
    lcSql = "CREATE TABLE [dbo].[LabItemsCpt] (" & _
            "    [idProductoCpt] [int] NOT NULL ," & _
            "    [ordenXresultado] [int] NOT NULL ," & _
            "    [idGrupo] [int] NOT NULL ," & _
            "    [idItemGrupo] [int] NULL ," & _
            "    [idItem] [int] NOT NULL ," & _
            "    [ValorSiEsCombo] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ValorReferencial] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Metodo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [SoloNumero] [bit] NULL ," & _
            "    [SoloTexto] [bit] NULL ," & _
            "    [SoloCombo] [bit] NULL ," & _
            "    [SoloCheck] [bit] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabItemsCpt] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_LabItemsCpt] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idProductoCpt]," & _
            "    [ordenXresultado]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabItemsCpt] ADD " & _
            " CONSTRAINT [FK_LabItemsCpt_FactCatalogoServicios] FOREIGN KEY" & _
            " (" & _
            "    [idProductoCpt]" & _
            " ) REFERENCES [dbo].[FactCatalogoServicios] (" & _
            "    [idProducto]" & _
            " )," & _
            " CONSTRAINT [FK_LabItemsCpt_labGrupos] FOREIGN KEY" & _
            " (" & _
            "    [IdGrupo]" & _
            " ) REFERENCES [dbo].[labGrupos] (" & _
            "    [IdGrupo]" & _
            " )," & _
            " CONSTRAINT [FK_LabItemsCpt_LabItems] FOREIGN KEY" & _
            " (" & _
            "    [idItem]" & _
            " ) REFERENCES [dbo].[LabItems] (" & _
            "    [idItem]" & _
            " )," & _
            " CONSTRAINT [FK_LabItemsCpt_LabItemsGrupos] FOREIGN KEY" & _
            " (" & _
            "    [idItemGrupo]" & _
            " ) REFERENCES [dbo].[LabItemsGrupos] (" & _
            "    [idItemGrupo]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItemsCpt"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from LabItemsCpt"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              Do While Not oRsTmpOpc.EOF
                 If oRsTmpOpc.Fields!idProductoCpt = oRsTmpOpc1.Fields!idProductoCpt And oRsTmpOpc.Fields!ordenXresultado = oRsTmpOpc1.Fields!ordenXresultado Then
                    lbNuevoRegistro = False
                    Exit Do
                 End If
                 oRsTmpOpc.MoveNext
              Loop
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idProductoCpt = oRsTmpOpc1.Fields!idProductoCpt
                oRsTmpOpc.Fields!ordenXresultado = oRsTmpOpc1.Fields!ordenXresultado
'           End If
                oRsTmpOpc.Fields!idGrupo = oRsTmpOpc1.Fields!idGrupo
                oRsTmpOpc.Fields!idItemGrupo = oRsTmpOpc1.Fields!idItemGrupo
                oRsTmpOpc.Fields!idItem = oRsTmpOpc1.Fields!idItem
                oRsTmpOpc.Fields!ValorSiEsCombo = oRsTmpOpc1.Fields!ValorSiEsCombo
                oRsTmpOpc.Fields!ValorReferencial = oRsTmpOpc1.Fields!ValorReferencial
                oRsTmpOpc.Fields!Metodo = oRsTmpOpc1.Fields!Metodo
                oRsTmpOpc.Fields!SoloNumero = oRsTmpOpc1.Fields!SoloNumero
                oRsTmpOpc.Fields!SoloTexto = oRsTmpOpc1.Fields!SoloTexto
                oRsTmpOpc.Fields!SoloCombo = oRsTmpOpc1.Fields!SoloCombo
                oRsTmpOpc.Fields!SoloCheck = oRsTmpOpc1.Fields!SoloCheck
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 113
    Me.Refresh
    txtTablaProceso.Text = "LabResultadoPorItems"
    lcSql = "CREATE TABLE [dbo].[LabResultadoPorItems] (" & _
            "    [idOrden] [int] NOT NULL ," & _
            "    [idProductoCpt] [int] NOT NULL ," & _
            "    [ordenXresultado] [int] NOT NULL ," & _
            "    [ValorNumero] [money] NULL ," & _
            "    [ValorTexto] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ValorCombo] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ValorCheck] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [realizaAnalisis] [int] NOT NULL ," & _
            "    [idUsuario] [int] NOT NULL ," & _
            "    [Fecha] [DateTime] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabResultadoPorItems] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_LabResutadoPorItems] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idOrden]," & _
            "    [idProductoCpt]," & _
            "    [ordenXresultado]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[LabResultadoPorItems] ADD " & _
            " CONSTRAINT [FK_LabResutadoPorItems_LabItemsCpt] FOREIGN KEY" & _
            " (" & _
            "    [idProductoCpt]," & _
            "    [ordenXresultado]" & _
            " ) REFERENCES [dbo].[LabItemsCpt] (" & _
            "    [idProductoCpt]," & _
            "    [ordenXresultado]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 114
    Me.Refresh
    txtTablaProceso.Text = "Citas"
    lcSql = "ALTER TABLE Citas add  EsCitaAdicional bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index IX_Programacion on Citas (idProgramacion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 115
    Me.Refresh
    txtTablaProceso.Text = "FactOrdenServicio"
    lcSql = "ALTER TABLE FactOrdenServicio add  FechaHoraRealizaCpt datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 116
    Me.Refresh
    txtTablaProceso.Text = "UPServicios"
    lcSql = "ALTER TABLE UPServicios add  UPShis varchar(6) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from UPServicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from UPServicios"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdUPS=" & oRsTmpOpc1.Fields!IdUPS
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
              
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdUPS = oRsTmpOpc1.Fields!IdUPS
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!estado = oRsTmpOpc1.Fields!estado
           oRsTmpOpc.Fields!UPShis = oRsTmpOpc1.Fields!UPShis
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 117
    Me.Refresh
    txtTablaProceso.Text = "farmMovimientoProgramas"
    lcSql = "ALTER TABLE farmMovimientoProgramas add  FechaHoraPrescribe datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 118
    Me.Refresh
    txtTablaProceso.Text = "farmMovimientoVentas"
    lcSql = "ALTER TABLE farmMovimientoVentas add  FechaHoraPrescribe datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 119
    Me.Refresh
    txtTablaProceso.Text = "farmPreVenta"
    lcSql = "ALTER TABLE farmPreVenta add  FechaHoraPrescribe datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents                                                   'debb-20/11/2013
    ProgressBar1.Value = 120
    Me.Refresh
    txtTablaProceso.Text = "DiagnosticosSoloPartos"
    lcSql = "CREATE TABLE [dbo].[DiagnosticosSoloPartos] (" & _
            "    [codigoCie10] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from DiagnosticosSoloPartos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from DiagnosticosSoloPartos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "codigoCie10='" & oRsTmpOpc1.Fields!CodigoCIE10 & "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
              
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!CodigoCIE10 = oRsTmpOpc1.Fields!CodigoCIE10
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 121
    Me.Refresh
    txtTablaProceso.Text = "AtencionesEpisodiosCabecera"
    lcSql = "CREATE TABLE [dbo].[AtencionesEpisodiosCabecera] (" & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [idEpisodio] [int] NOT NULL ," & _
            "    [FechaApertura] [datetime] NOT NULL ," & _
            "    [FechaCierre] [DateTime] null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[AtencionesEpisodiosCabecera] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_AtencionesEpisodiosCabecera] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idPaciente]," & _
            "    [idEpisodio]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 122
    Me.Refresh
    txtTablaProceso.Text = "AtencionesEpisodiosDetalle"
    lcSql = "CREATE TABLE [dbo].[AtencionesEpisodiosDetalle] (" & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [idEpisodio] [int] NOT NULL ," & _
            "    [idAtencion] [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[AtencionesEpisodiosDetalle] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_AtencionesEpisodiosDetalle] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idEpisodio]," & _
            "    [idPaciente]," & _
            "    [idAtencion]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[AtencionesEpisodiosDetalle] ADD " & _
            " CONSTRAINT [FK_AtencionesEpisodiosDetalle_Atenciones] FOREIGN KEY" & _
            " (" & _
            "    [IdAtencion]" & _
            " ) REFERENCES [dbo].[Atenciones] (" & _
            "    [IdAtencion]" & _
            " )," & _
            " CONSTRAINT [FK_AtencionesEpisodiosDetalle_AtencionesEpisodiosCabecera] FOREIGN KEY" & _
            " (" & _
            "    [idPaciente]," & _
            "    [idEpisodio]" & _
            " ) REFERENCES [dbo].[AtencionesEpisodiosCabecera] (" & _
            "    [idPaciente]," & _
            "    [idEpisodio]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 123
    Me.Refresh
    txtTablaProceso.Text = "ProgramacionMedica"
    lcSql = "alter table ProgramacionMedica ADD FechaReg datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Update ProgramacionMedica set FechaReg=fecha where fechaReg is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table ProgramacionMedica ADD TiempoPromedioAtencion int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE ProgramacionMedica ADD HoraFinProgramacion char(5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 124
    Me.Refresh
    txtTablaProceso.Text = "HIS_FACTCATALOGOSERVICIOS"
    lcSql = "CREATE TABLE [dbo].[HIS_FACTCATALOGOSERVICIOS] (" & _
    " [IdDiagCpt] [int] NULL ," & _
    " [CodigoDiagCpt] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [DescripcionDiagCpt] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [EsCpt] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [CodCondicion] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
    " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE HIS_FACTCATALOGOSERVICIOS add  DxSexo varchar (4) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE HIS_FACTCATALOGOSERVICIOS add  CodigoDiagCptSinPunto varchar (20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE HIS_FACTCATALOGOSERVICIOS add  MasDeUnDiagnosticos varchar (20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from HIS_FACTCATALOGOSERVICIOS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_FACTCATALOGOSERVICIOS"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
If oRsTmpOpc1.Fields!IdDiagCpt = 14630 Then
lcSql = ""
End If
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
                lcSql = "select * from HIS_FACTCATALOGOSERVICIOS where IdDiagCpt=" & oRsTmpOpc1.Fields!IdDiagCpt
                If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
                oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                If oRsTmpOpc.RecordCount > 0 Then
                   lbNuevoRegistro = False
                End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdDiagCpt = oRsTmpOpc1.Fields!IdDiagCpt
           End If
            oRsTmpOpc.Fields!CodigoDiagCpt = oRsTmpOpc1.Fields!CodigoDiagCpt
            oRsTmpOpc.Fields!DescripcionDiagCpt = oRsTmpOpc1.Fields!DescripcionDiagCpt
            oRsTmpOpc.Fields!EsCpt = oRsTmpOpc1.Fields!EsCpt
            oRsTmpOpc.Fields!CodCondicion = oRsTmpOpc1.Fields!CodCondicion
            oRsTmpOpc.Fields!DxSexo = oRsTmpOpc1.Fields!DxSexo
            oRsTmpOpc.Fields!CodigoDiagCptSinPunto = oRsTmpOpc1.Fields!CodigoDiagCptSinPunto
            oRsTmpOpc.Fields!MasDeUnDiagnosticos = oRsTmpOpc1.Fields!MasDeUnDiagnosticos
            oRsTmpOpc.Update
           
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 125
    Me.Refresh
    txtTablaProceso.Text = "HIS_VALIDACIONES"
    lcSql = "CREATE TABLE [dbo].[HIS_VALIDACIONES] (" & _
    " [CODVALIDACION] [int] NOT NULL ," & _
    " [TXTDESCRIPCION] [varchar] (300) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [TXTVALIDACION] [varchar] (2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [TXTARCHIVO] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [TXTNPAGINA] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [CodCondicion] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [FECACTIVACION] [datetime] NOT NULL ," & _
    " [FECINACTIVACION] [DateTime]  NOT NULL" & _
    " ) ON [PRIMARY]"
    
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE HIS_VALIDACIONES add  CCH varchar (1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from HIS_VALIDACIONES"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from HIS_VALIDACIONES"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           
           lbNuevoRegistro = True
            lcSql = "select * from HIS_VALIDACIONES where CODVALIDACION=" & oRsTmpOpc1.Fields!CODVALIDACION
            If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            If oRsTmpOpc.RecordCount > 0 Then
               lbNuevoRegistro = False
            End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!CODVALIDACION = oRsTmpOpc1.Fields!CODVALIDACION
           End If
           oRsTmpOpc.Fields!TXTDESCRIPCION = oRsTmpOpc1.Fields!TXTDESCRIPCION
           oRsTmpOpc.Fields!TXTVALIDACION = oRsTmpOpc1.Fields!TXTVALIDACION
           oRsTmpOpc.Fields!TXTARCHIVO = oRsTmpOpc1.Fields!TXTARCHIVO
           oRsTmpOpc.Fields!TXTNPAGINA = oRsTmpOpc1.Fields!TXTNPAGINA
           oRsTmpOpc.Fields!CodCondicion = oRsTmpOpc1.Fields!CodCondicion
           oRsTmpOpc.Fields!FECACTIVACION = oRsTmpOpc1.Fields!FECACTIVACION
           oRsTmpOpc.Fields!FECINACTIVACION = oRsTmpOpc1.Fields!FECINACTIVACION
           oRsTmpOpc.Fields!CCH = oRsTmpOpc1.Fields!CCH
           oRsTmpOpc.Update
           txtTablaProceso.Text = oRsTmpOpc1.Fields!CODVALIDACION
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 126
    Me.Refresh
    txtTablaProceso.Text = "ProductosHis"
    lcSql = "CREATE TABLE [dbo].[ProductosHis] (" & _
    " [IdProductoHis] [int] NOT NULL ," & _
    " [EsCpt] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [CodigoProductoHis] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
    " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProductosHis] WITH NOCHECK ADD " & _
    " CONSTRAINT [PK_ProductosHis] PRIMARY KEY  CLUSTERED" & _
    " (" & _
    "    [IdProductoHis]," & _
    "    [EsCpt]" & _
    " )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProductosHis"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProductosHis"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
            lcSql = "select * from ProductosHis where IdProductoHis=" & oRsTmpOpc1.Fields!IdProductoHis
            If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            If oRsTmpOpc.RecordCount > 0 Then
               lbNuevoRegistro = False
            End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdProductoHis = oRsTmpOpc1.Fields!IdProductoHis
                oRsTmpOpc.Fields!EsCpt = oRsTmpOpc1.Fields!EsCpt
                oRsTmpOpc.Fields!CodigoProductoHis = oRsTmpOpc1.Fields!CodigoProductoHis
                oRsTmpOpc.Update
           End If
           
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
     DoEvents
    ProgressBar1.Value = 127
    Me.Refresh
    txtTablaProceso.Text = "HIS_TEMPORAL"
    lcSql = "CREATE TABLE [dbo].[HIS_TEMPORAL] (" & _
    " [Codigo1] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Codigo2] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Codigo3] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Codigo4] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Codigo5] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Codigo6] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [LabConf1] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [LabConf2] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [LabConf3] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [LabConf4] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [LabConf5] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [LabConf6] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    " [Diagnost1] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Diagnost2] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Diagnost3] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Diagnost4] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Diagnost5] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Diagnost6] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Edad] [int] NOT NULL ," & _
    " [TIP_EDAD] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Sexo] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    " [Peso] [money] NOT NULL ," & _
    " [FichaFam] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
    " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE HIS_TEMPORAL add  Establecimiento varchar (1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "ALTER TABLE HIS_TEMPORAL add  Servicio varchar (1) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

Frank1:

     'Modulo Materno _ Frank
    Dim oRsCloneTmp As ADODB.Recordset
    
    
    
    DoEvents
    ProgressBar1.Value = 128
    Me.Refresh
    txtTablaProceso.Text = "Programas"
    lcSql = "CREATE TABLE [dbo].[Programas] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[Programas] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_Programas] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [IdPrograma]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Programas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Programas"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPrograma=" & oRsTmpOpc1.Fields!IdPrograma
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 129
    Me.Refresh
    txtTablaProceso.Text = "ProCabeceraConfigDatos"
    lcSql = "CREATE TABLE [dbo].[ProCabeceraConfigDatos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdCabDato] [int] NOT NULL ," & _
            "    [Cab_Texto] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Cab_Tipo] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Cab_Ancho] [int] NULL ," & _
            "    [Cab_EsDatoObligatorio] [int] NULL ," & _
            "    [Cab_TextoToolTip] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Cab_EsDatoCalculado] [bit] NULL ," & _
            "    [Cab_FormulaCalculaValor] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Cab_EsDatoCalculador] [bit] NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table ProCabeceraConfigDatos ADD Cab_RangoInicial int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table ProCabeceraConfigDatos ADD Cab_RangoFinal int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table ProCabeceraConfigDatos ADD Cab_Fila int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table ProCabeceraConfigDatos ADD Cab_Columna int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[ProCabeceraConfigDatos] ADD" & _
            " CONSTRAINT [FK_ProCabeceraConfigDatos_Programas] FOREIGN KEY" & _
            " ([IdPrograma]) REFERENCES [dbo].[Programas] ([IdPrograma])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCabeceraConfigDatos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCabeceraConfigDatos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdCabDato])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCabeceraConfigDatos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCabeceraConfigDatos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           
           lcSql = "select * from ProCabeceraConfigDatos where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdCabDato=" & _
                                         oRsTmpOpc1.Fields!IdCabDato
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdCabDato = oRsTmpOpc1.Fields!IdCabDato
           End If
           oRsTmpOpc.Fields!Cab_Texto = oRsTmpOpc1.Fields!Cab_Texto
           oRsTmpOpc.Fields!Cab_Tipo = oRsTmpOpc1.Fields!Cab_Tipo
           oRsTmpOpc.Fields!Cab_Ancho = oRsTmpOpc1.Fields!Cab_Ancho
           oRsTmpOpc.Fields!Cab_EsDatoObligatorio = oRsTmpOpc1.Fields!Cab_EsDatoObligatorio
           oRsTmpOpc.Fields!Cab_TextoToolTip = oRsTmpOpc1.Fields!Cab_TextoToolTip
           oRsTmpOpc.Fields!Cab_EsDatoCalculado = oRsTmpOpc1.Fields!Cab_EsDatoCalculado
           oRsTmpOpc.Fields!Cab_FormulaCalculaValor = oRsTmpOpc1.Fields!Cab_FormulaCalculaValor
           oRsTmpOpc.Fields!Cab_EsDatoCalculador = oRsTmpOpc1.Fields!Cab_EsDatoCalculador
           oRsTmpOpc.Fields!Cab_RangoInicial = oRsTmpOpc1.Fields!Cab_RangoInicial
           oRsTmpOpc.Fields!Cab_RangoFinal = oRsTmpOpc1.Fields!Cab_RangoFinal
           oRsTmpOpc.Fields!Cab_Fila = oRsTmpOpc1.Fields!Cab_Fila
           oRsTmpOpc.Fields!Cab_Columna = oRsTmpOpc1.Fields!Cab_Columna
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 130
    Me.Refresh
    txtTablaProceso.Text = "ProControlConfigDatos"
    lcSql = "CREATE TABLE [dbo].[ProControlConfigDatos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdControlDato] [int] NOT NULL ," & _
            "    [Control_Texto] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Control_Tipo] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Control_Ancho] [int] NULL ," & _
            "    [Control_EsDatoObligatorio] [bit] NULL ," & _
            "    [Control_TextoToolTip] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Control_EsPresion] [bit] NULL ," & _
            "    [Control_EsPeso] [bit] NULL ," & _
            "    [Control_EsTalla] [bit] NULL ," & _
            "    [Control_EsDatoCalculado] [bit] NULL ," & _
            "    [Control_FormulaCalculaValor] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [Control_EsDatoGrafico] [bit] NULL ," & _
            "    [Control_EsGraficoEjeX] [bit] NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProControlConfigDatos] ADD" & _
            " CONSTRAINT [FK_ProControlConfigDatos_Programas] FOREIGN KEY" & _
            " ([IdPrograma]) REFERENCES [dbo].[Programas] ([IdPrograma])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProControlConfigDatos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProControlConfigDatos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdControlDato])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " alter table ProControlConfigDatos add Control_Fila int not null default 0"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " alter table ProControlConfigDatos add Control_Columna int not null default 0"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from ProControlConfigDatos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProControlConfigDatos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ProControlConfigDatos where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdControlDato=" & _
                                         oRsTmpOpc1.Fields!IdControlDato
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdControlDato = oRsTmpOpc1.Fields!IdControlDato
           End If
           oRsTmpOpc.Fields!Control_Texto = oRsTmpOpc1.Fields!Control_Texto
           oRsTmpOpc.Fields!Control_Tipo = oRsTmpOpc1.Fields!Control_Tipo
           oRsTmpOpc.Fields!Control_Ancho = oRsTmpOpc1.Fields!Control_Ancho
           oRsTmpOpc.Fields!Control_EsDatoObligatorio = oRsTmpOpc1.Fields!Control_EsDatoObligatorio
           oRsTmpOpc.Fields!Control_TextoToolTip = oRsTmpOpc1.Fields!Control_TextoToolTip
           oRsTmpOpc.Fields!Control_EsPresion = oRsTmpOpc1.Fields!Control_EsPresion
           oRsTmpOpc.Fields!Control_EsPeso = oRsTmpOpc1.Fields!Control_EsPeso
           oRsTmpOpc.Fields!Control_EsTalla = oRsTmpOpc1.Fields!Control_EsTalla
           oRsTmpOpc.Fields!Control_EsDatoCalculado = oRsTmpOpc1.Fields!Control_EsDatoCalculado
           oRsTmpOpc.Fields!Control_FormulaCalculaValor = oRsTmpOpc1.Fields!Control_FormulaCalculaValor
           oRsTmpOpc.Fields!Control_EsDatoGrafico = oRsTmpOpc1.Fields!Control_EsDatoGrafico
           oRsTmpOpc.Fields!Control_EsGraficoEjeX = oRsTmpOpc1.Fields!Control_EsGraficoEjeX
           oRsTmpOpc.Fields!Control_Fila = oRsTmpOpc1.Fields!Control_Fila
           oRsTmpOpc.Fields!Control_Columna = oRsTmpOpc1.Fields!Control_Columna
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
'
    DoEvents
    ProgressBar1.Value = 131
    Me.Refresh
    txtTablaProceso.Text = "ProCatalogoControles"
    lcSql = "CREATE TABLE [dbo].[ProCatalogoControles] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCatalogoControles] ADD" & _
            " CONSTRAINT [FK_ProCatalogoControles_Programas] FOREIGN KEY" & _
            " ([IdPrograma]) REFERENCES [dbo].[Programas] ([IdPrograma])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCatalogoControles] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCatalogoControles] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdControl])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoControles"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoControles"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ProCatalogoControles where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdControl=" & _
                                         oRsTmpOpc1.Fields!IdControl
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdControl = oRsTmpOpc1.Fields!IdControl
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    
    Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Sub MigraUltimaVersion_TablaSIGH_Parte3(oConexHBT As Connection, oConexODBC As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String
    On Error GoTo errMg
    
    '
    DoEvents
    ProgressBar1.Value = 132
    Me.Refresh
    txtTablaProceso.Text = "ProCatalogoDiagnosticos"
    lcSql = "CREATE TABLE [dbo].[ProCatalogoDiagnosticos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdDiagnostico] [int] NOT NULL " & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCatalogoDiagnosticos] ADD" & _
            " CONSTRAINT [FK_ProCatalogoDiagnosticos_Diagnosticos] FOREIGN KEY" & _
            " ([IdDiagnostico]) REFERENCES [dbo].[Diagnosticos] ([IdDiagnostico])," & _
            " CONSTRAINT [FK_ProCatalogoDiagnosticos_Programas1] FOREIGN KEY" & _
            " ([IdPrograma]) REFERENCES [dbo].[Programas] ([IdPrograma])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCatalogoDiagnosticos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCatalogoDiagnosticos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdDiagnostico])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoDiagnosticos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
'           lbNuevoRegistro = True
'           If oRsTmpOpc.RecordCount > 0 Then
''              oRsTmpOpc.MoveFirst
''              oRsTmpOpc.Find "IdPrograma=" & oRsTmpOpc1.Fields!IdPrograma & " and IdDiagnostico=" & oRsTmpOpc1.Fields!idDiagnostico
''              If Not oRsTmpOpc.EOF Then
''                 lbNuevoRegistro = False
''              End If
'
'              Set oRsCloneTmp = oRsTmpOpc.Clone
'              oRsCloneTmp.MoveFirst
'              oRsCloneTmp.Filter = "IdPrograma=" & oRsTmpOpc1.Fields!IdPrograma & " and IdDiagnostico=" & oRsTmpOpc1.Fields!idDiagnostico
'              If Not oRsCloneTmp.EOF Then
'                 lbNuevoRegistro = False
'              End If
'              oRsCloneTmp.Close
'              Set oRsCloneTmp = Nothing
'
'           End If
           lcSql = "select * from ProCatalogoDiagnosticos where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdDiagnostico=" & _
                                         oRsTmpOpc1.Fields!IdDiagnostico
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdDiagnostico = oRsTmpOpc1.Fields!IdDiagnostico
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 133
    Me.Refresh
    txtTablaProceso.Text = "ProCatalogoProcedimientos"
    lcSql = "CREATE TABLE [dbo].[ProCatalogoProcedimientos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdDiagnostico] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCatalogoProcedimientos] ADD " & _
            " CONSTRAINT [FK_ProCatalogoProcedimientos_FactCatalogoServicios1] FOREIGN KEY" & _
            " (" & _
            " [IdProducto]" & _
            " ) REFERENCES [dbo].[FactCatalogoServicios] (" & _
            " [IdProducto]" & _
            " )," & _
            " CONSTRAINT [FK_ProCatalogoProcedimientos_ProCatalogoDiagnosticos] FOREIGN KEY" & _
            " ([IdPrograma],[IdDiagnostico]) REFERENCES [dbo].[ProCatalogoDiagnosticos] (" & _
            " [IdPrograma],[IdDiagnostico])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCatalogoProcedimientos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCatalogoProcedimientos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdDiagnostico],[IdProducto])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoProcedimientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoProcedimientos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ProCatalogoProcedimientos where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdDiagnostico=" & _
                                         oRsTmpOpc1.Fields!IdDiagnostico & _
                                         " and IdProducto=" & oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdDiagnostico = oRsTmpOpc1.Fields!IdDiagnostico
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    '
    DoEvents
    ProgressBar1.Value = 134
    Me.Refresh
    txtTablaProceso.Text = "ProCatalogoTratamientos"
    lcSql = "CREATE TABLE [dbo].[ProCatalogoTratamientos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL " & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCatalogoTratamientos] ADD" & _
            " CONSTRAINT [FK_ProCatalogoTratamientos_FactCatalogoBienesInsumos] FOREIGN KEY" & _
            " ([IdProducto] REFERENCES [dbo].[FactCatalogoBienesInsumos] (" & _
            " [IdProducto])," & _
            " CONSTRAINT [FK_ProCatalogoTratamientos_Programas] FOREIGN KEY" & _
            " ([IdPrograma]) REFERENCES [dbo].[Programas] (" & _
            " [IdPrograma])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCatalogoTratamientos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCatalogoTratamientos] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProducto]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "delete from ProCatalogoTratamientos where idProducto=730 or " & _
                                                      "idProducto=893 or " & _
                                                      "idProducto=1013 or " & _
                                                      "idProducto=1117 or " & _
                                                      "idProducto=8848 or " & _
                                                      "idProducto=10539 or " & _
                                                      "idProducto=10540 or " & _
                                                      "idProducto=10541 or " & _
                                                      "idProducto=10611 or " & _
                                                      "idProducto=10611 or " & _
                                                      "idProducto=10612 or " & _
                                                      "idProducto=10625 or " & _
                                                      "idProducto=10626 or " & _
                                                      "idProducto=10627 "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoTratamientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from ProCatalogoTratamientos"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF

           lcSql = "select * from ProCatalogoTratamientos where IdPrograma=" & _
                                         oRsTmpOpc1.Fields!IdPrograma & _
                                         " and IdProducto=" & _
                                         oRsTmpOpc1.Fields!IdProducto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdPrograma = oRsTmpOpc1.Fields!IdPrograma
                oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
                oRsTmpOpc.Update
           End If
           
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 135
    Me.Refresh
    txtTablaProceso.Text = "TiposResultadosServInterm"
    lcSql = "CREATE TABLE [dbo].[TiposResultadosServInterm] (" & _
            "    [IdResultadoSI] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[TiposResultadosServInterm] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_TiposResultadosServInterm] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdResultadoSI]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposResultadosServInterm"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposResultadosServInterm"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdResultadoSI=" & oRsTmpOpc1.Fields!IdResultadoSI
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdResultadoSI = oRsTmpOpc1.Fields!IdResultadoSI
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 136
    Me.Refresh
    txtTablaProceso.Text = "ProCabecera"
    lcSql = "CREATE TABLE [dbo].[ProCabecera] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProcabecera] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [IdPaciente] [int] NOT NULL " & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCabecera] ADD" & _
            " CONSTRAINT [FK_ProCabecera_Programas] FOREIGN KEY" & _
            " (" & _
            " [IdPrograma]" & _
            " ) REFERENCES [dbo].[Programas] (" & _
            " [IdPrograma]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCabecera] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCabecera] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProcabecera]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 137
    Me.Refresh
    txtTablaProceso.Text = "ProCabeceraDato"
    lcSql = "CREATE TABLE [dbo].[ProCabeceraDato] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdCabDato] [int] NOT NULL ," & _
            "    [CabDato] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProCabeceraDato] ADD" & _
            " CONSTRAINT [FK_ProCabeceraDato_ProCabecera] FOREIGN KEY" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]" & _
            " ) REFERENCES [dbo].[ProCabecera] (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]" & _
            " )," & _
            " CONSTRAINT [FK_ProCabeceraDato_ProCabeceraConfigDatos] FOREIGN KEY" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdCabDato]" & _
            " ) REFERENCES [dbo].[ProCabeceraConfigDatos] (" & _
            " [IdPrograma]," & _
            " [IdCabDato]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProCabeceraDato] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProCabeceraDato] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]," & _
            " [IdCabDato]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 138
    Me.Refresh
    txtTablaProceso.Text = "ProControles"
    lcSql = "CREATE TABLE [dbo].[ProControles] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [IdAtencion] [int] NULL ," & _
            "    [FechaControl] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProControles] ADD" & _
            " CONSTRAINT [FK_ProControles_ProCabecera] FOREIGN KEY" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]" & _
            " ) REFERENCES [dbo].[ProCabecera] (" & _
            " [IdPrograma]," & _
            " [IdProcabecera]" & _
            " )," & _
            " CONSTRAINT [FK_ProControles_ProCatalogoControles] FOREIGN KEY" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdControl]" & _
            " ) REFERENCES [dbo].[ProCatalogoControles] (" & _
            " [IdPrograma]," & _
            " [IdControl]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProControles] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProControles] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]," & _
            " [IdControl]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " alter table ProControles add ControlOtroEESS bit not null default 0 "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " alter table ProControles add IdEstablecimiento int null "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 139
    Me.Refresh
    txtTablaProceso.Text = "ProControlDato"
    lcSql = "CREATE TABLE [dbo].[ProControlDato] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [IdControlDato] [int] NOT NULL ," & _
            "    [ControlDato] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProControlDato] ADD" & _
            " CONSTRAINT [FK_ProControlDato_ProControlConfigDatos] FOREIGN KEY" & _
            " (            [IdPrograma]," & _
            "              [IdControlDato]" & _
            " ) REFERENCES [dbo].[ProControlConfigDatos] (" & _
            "              [IdPrograma]," & _
            "              [IdControlDato]" & _
            " )," & _
            " CONSTRAINT [FK_ProControlDato_ProControles] FOREIGN KEY" & _
            " (            [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " ) REFERENCES [dbo].[ProControles] (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProControlDato] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_ProControlDato] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdProCabecera],[IdControl],[IdControlDato]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 140
    Me.Refresh
    txtTablaProceso.Text = "ProDiagnosticos"
    lcSql = "CREATE TABLE [dbo].[ProDiagnosticos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [IdDiagnostico] [int] NOT NULL ," & _
            "    [Principal] [bit] NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProDiagnosticos] ADD" & _
            " CONSTRAINT [FK_ProDiagnosticos_Diagnosticos] FOREIGN KEY" & _
            " (" & _
            "              [IdDiagnostico]" & _
            " ) REFERENCES [dbo].[Diagnosticos] (" & _
            "              [IdDiagnostico]" & _
            " ) NOT FOR REPLICATION ," & _
            " CONSTRAINT [FK_ProDiagnosticos_ProControles] FOREIGN KEY" & _
            " (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " ) REFERENCES [dbo].[ProControles] (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProDiagnosticos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProDiagnosticos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdProCabecera],[IdControl],[IdDiagnostico]" & _
            " )  ON [PRIMARY]            "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " ALTER TABLE ProDiagnosticos ADD labConfHIS VARCHAR(3) NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " ALTER TABLE ProDiagnosticos ADD IdSubClasificacionDX INT NOT NULL DEFAULT 102"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '12112014
    lcSql = " ALTER TABLE ProDiagnosticos DROP CONSTRAINT PK_ProDiagnosticos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    '
    DoEvents
    ProgressBar1.Value = 141
    Me.Refresh
    txtTablaProceso.Text = "ProProcedimientos"
    lcSql = "CREATE TABLE [dbo].[ProProcedimientos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [IdDiagnostico] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL ," & _
            "    [IdResultado] [int] NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProProcedimientos] ADD" & _
            " CONSTRAINT [FK_ProProcedimientos_ProCatalogoProcedimientos] FOREIGN KEY" & _
            " (" & _
            "              [IdPrograma]," & _
            "              [IdDiagnostico]," & _
            "              [IdProducto]" & _
            " ) REFERENCES [dbo].[ProCatalogoProcedimientos] (" & _
            "              [IdPrograma]," & _
            "              [IdDiagnostico]," & _
            "              [IdProducto]" & _
            " )," & _
            " CONSTRAINT [FK_ProProcedimientos_ProControles] FOREIGN KEY" & _
            " (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " ) REFERENCES [dbo].[ProControles] (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " )," & _
            " CONSTRAINT [FK_ProProcedimientos_TiposResultadosServInterm] FOREIGN KEY" & _
            " ([IdResultado]" & _
            " ) REFERENCES [dbo].[TiposResultadosServInterm] ([IdResultadoSI]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ProProcedimientos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProProcedimientos] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            " [IdPrograma]," & _
            " [IdProCabecera]," & _
            " [IdControl]," & _
            " [IdDiagnostico]," & _
            " [IdProducto]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " ALTER TABLE ProProcedimientos DROP CONSTRAINT FK_ProProcedimientos_ProCatalogoProcedimientos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = " ALTER TABLE ProProcedimientos ADD labConfHIS VARCHAR(3) NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 142
    Me.Refresh
    txtTablaProceso.Text = "ProTratamientos"
    lcSql = "CREATE TABLE [dbo].[ProTratamientos] (" & _
            "    [IdPrograma] [int] NOT NULL ," & _
            "    [IdProCabecera] [int] NOT NULL ," & _
            "    [IdControl] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL " & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProTratamientos] ADD" & _
            " CONSTRAINT [FK_ProTratamientos_ProCatalogoTratamientos] FOREIGN KEY" & _
            " (" & _
            "              [IdPrograma]," & _
            "              [IdProducto]" & _
            " ) REFERENCES [dbo].[ProCatalogoTratamientos] (" & _
            "              [IdPrograma]," & _
            "              [IdProducto]" & _
            " )," & _
            " CONSTRAINT [FK_ProTratamientos_ProControles] FOREIGN KEY" & _
            " (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " ) REFERENCES [dbo].[ProControles] (" & _
            "              [IdPrograma]," & _
            "              [IdProCabecera]," & _
            "              [IdControl]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[ProTratamientos] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_ProTratamientos] PRIMARY KEY  CLUSTERED" & _
            " ([IdPrograma],[IdProCabecera],[IdControl],[IdProducto] " & _
            " )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 143
    Me.Refresh
    txtTablaProceso.Text = "CentrosCosto"
    lcSql = "select * from CentrosCosto"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from CentrosCosto"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from CentrosCosto where idCentroCosto=" & oRsTmpOpc1.Fields!IdCentroCosto
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                lcSql = "DBCC CHECKIDENT (CentrosCosto, RESEED, " & Trim(Str(oRsTmpOpc1.Fields!IdCentroCosto - 1)) & ")"
                If oRsTmpOpc2.State = 1 Then oRsTmpOpc1.Close
                oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
                
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdCentroCosto = oRsTmpOpc1.Fields!IdCentroCosto
                oRsTmpOpc.Fields!Codigo = oRsTmpOpc1.Fields!Codigo
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
   'Actualizado 23092014
    DoEvents
    ProgressBar1.Value = 144
    Me.Refresh
    txtTablaProceso.Text = "AtencionesEstanciaHospitalaria"
    lcSql = "ALTER TABLE AtencionesEstanciaHospitalaria ADD IdMedicoOrdenaOrigen int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[AtencionesEstanciaHospitalaria] ADD " & _
            "    CONSTRAINT [IdMedicoOrdenaOrigen_FK1] FOREIGN KEY" & _
            "    (" & _
            "        [IdMedicoOrdenaOrigen]" & _
            "    ) REFERENCES [dbo].[Medicos] (" & _
            "        [IdMedico]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    'Actualizado 2609
    DoEvents
    ProgressBar1.Value = 145
    Me.Refresh
    txtTablaProceso.Text = "RecetaClasificacionViasAdmin"
    lcSql = "CREATE TABLE [dbo].[RecetaClasificacionViasAdmin] (" & _
            "    [IdCategoria] [int] NOT NULL ," & _
            "    [Categoria] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
   
    lcSql = "ALTER TABLE [dbo].[RecetaClasificacionViasAdmin] WITH NOCHECK ADD " & _
    " CONSTRAINT [PK_RecetaClasificacionViasAdmin] PRIMARY KEY  CLUSTERED" & _
    " (" & _
    "     [IdCategoria]" & _
    " )  ON [PRIMARY]"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from RecetaClasificacionViasAdmin order by IdCategoria"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaClasificacionViasAdmin order by IdCategoria"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdCategoria=" & oRsTmpOpc1.Fields!IdCategoria
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdCategoria = oRsTmpOpc1.Fields!IdCategoria
           End If
           oRsTmpOpc.Fields!Categoria = oRsTmpOpc1.Fields!Categoria
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 146
    Me.Refresh
    txtTablaProceso.Text = "RecetaViaAdministracion"
    lcSql = "CREATE TABLE [dbo].[RecetaViaAdministracion] (" & _
            "    [IdViaAdministracion] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
            "    [IdCategoria] [int] NOT NULL " & _
            "    ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
   
    lcSql = "ALTER TABLE [dbo].[RecetaViaAdministracion] WITH NOCHECK ADD " & _
    " CONSTRAINT [PK_RecetaViaAdministracion] PRIMARY KEY  CLUSTERED" & _
    " (" & _
    "     [IdViaAdministracion]" & _
    " )  ON [PRIMARY]"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[RecetaViaAdministracion] ADD " & _
            "    CONSTRAINT [FK_RecetaViaAdministracion_RecetaClasificacionViasAdmin] FOREIGN KEY" & _
            "    (" & _
            "        [IdCategoria]" & _
            "    ) REFERENCES [dbo].[RecetaClasificacionViasAdmin] (" & _
            "        [IdCategoria]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from RecetaViaAdministracion order by IdCategoria"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaViaAdministracion order by IdCategoria"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdViaAdministracion=" & oRsTmpOpc1.Fields!IdViaAdministracion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdViaAdministracion = oRsTmpOpc1.Fields!IdViaAdministracion
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!IdCategoria = oRsTmpOpc1.Fields!IdCategoria
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    lcSql = "Alter table RecetaDetalle ADD IdViaAdministracion int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic


'Frank 10112014
     DoEvents
     ProgressBar1.Value = 147
     Me.Refresh
     txtTablaProceso.Text = "CptExcepcionesRecetasHIS"
     lcSql = "CREATE TABLE [dbo].[CptExcepcionesRecetasHIS] (" & _
             "    [IdProducto] [int] NOT NULL" & _
             "    ) ON [PRIMARY]"
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
     lcSql = "ALTER TABLE [dbo].[CptExcepcionesRecetasHIS] WITH NOCHECK ADD " & _
     " CONSTRAINT [PK_CptExcepcionesRecetasHIS] PRIMARY KEY  CLUSTERED" & _
     " (" & _
     "     [IdProducto]" & _
     " )  ON [PRIMARY]"
     If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
     oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     
     lcSql = "select * from CptExcepcionesRecetasHIS "
     If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
     oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     lcSql = "select * from CptExcepcionesRecetasHIS "
     If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
     oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
     If oRsTmpOpc1.RecordCount > 0 Then
         oRsTmpOpc1.MoveFirst
         Do While Not oRsTmpOpc1.EOF
            lbNuevoRegistro = True
            If oRsTmpOpc.RecordCount > 0 Then
               oRsTmpOpc.MoveFirst
               oRsTmpOpc.Find "IdProducto=" & oRsTmpOpc1.Fields!IdProducto
               If Not oRsTmpOpc.EOF Then
                  lbNuevoRegistro = False
               End If
            End If
            If lbNuevoRegistro = True Then
                 oRsTmpOpc.AddNew
                 oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
            End If
            oRsTmpOpc.Update
            oRsTmpOpc1.MoveNext
         Loop
     End If
     oRsTmpOpc1.Close
     
    'Frank 13112014
    lcSql = "update FactCatalogoBienesInsumos set IdPartida =1 where IdPartida is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "update FactCatalogoBienesInsumos set IdCentroCosto =999 where IdCentroCosto  is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    'Frank 13112014
    lcSql = "update FactCatalogoServicios  set IdPartida =999 where IdPartida is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "update FactCatalogoServicios  set IdCentroCosto =999 where IdCentroCosto  is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    'Frank 13112014
    lcSql = "Delete from RolesItems where IdListItem=1357 and IdRol=1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "insert into RolesItems (Consultar,Eliminar,Modificar,Agregar,IdRol,IdListItem) values (1,1,1,1,1,1357)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Delete from RolesItems where IdListItem=1358 and IdRol=1"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "insert into RolesItems (Consultar,Eliminar,Modificar,Agregar,IdRol,IdListItem) values (1,1,1,1,1,1358)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
     
    lcSql = "update Diagnosticos set EsActivo = 0 where CodigoCIE2004='e00.6'" 'Casos especiales
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    
    Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Sub MigraUltimaVersion_TablaSIGH_Parte8(oConexHBT As Connection, oConexODBC As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lbNuevoRegistro As Boolean
    Dim lnCodigoEstablecimiento As Long
    Dim LcTexto1 As String
    On Error GoTo errMg
    
    '166
    DoEvents
    ProgressBar1.Value = 202
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_CatalogoDominios"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_CatalogoDominios] (" & _
            "    [IdDominio] [int] NOT NULL ," & _
            "    [CodDominio] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [DominioTexto] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    CONSTRAINT [PK_Enfermeria_CatalogoDominios] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdDominio]" & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoDominios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoDominios"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from Enfermeria_CatalogoDominios where IdDominio=" & _
                                         oRsTmpOpc1.Fields!IdDominio
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdDominio = oRsTmpOpc1.Fields!IdDominio
                oRsTmpOpc.Fields!CodDominio = oRsTmpOpc1.Fields!CodDominio
                oRsTmpOpc.Fields!DominioTexto = oRsTmpOpc1.Fields!DominioTexto
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
       
    
    DoEvents
    ProgressBar1.Value = 203
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_CatalogoVariables"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_CatalogoVariables] (" & _
            "    [IdVariable] [int] NOT NULL ," & _
            "    [IdDominio] [int] NOT NULL ," & _
            "    [OrdernDominio] [int] NOT NULL ," & _
            "    [Texto] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Tipo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Ancho] [int] NOT NULL ," & _
            "    [EsDatoObligatorio] [bit] NOT NULL ," & _
            "    [TextoTooltip] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [EsDatoGrafico] [bit] NOT NULL ," & _
            "    [TieneFormatoMask] [bit] NOT NULL ," & _
            "    [FormatoMask] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [TieneRango] [bit] NOT NULL ," & _
            "    [RangoInicial] [int] NULL ," & _
            "    [RangoFinal] [int] NULL ," & _
            "    [PosicionFila] [int] NULL ," & _
            "    [PosicionColumna] [int] NULL," & _
            "    CONSTRAINT [PK_Enfermeria_CatalogoVariables] PRIMARY KEY  CLUSTERED" & _
            "    ([IdVariable])  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_CatalogoVariables_Enfermeria_CatalogoDominios] FOREIGN KEY" & _
            "    ([IdDominio]" & _
            "    ) REFERENCES [dbo].[Enfermeria_CatalogoDominios] (" & _
            "    [IdDominio])" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoVariables"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoVariables"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from Enfermeria_CatalogoVariables where IdDominio=" & _
                                         oRsTmpOpc1.Fields!IdDominio & " and IdVariable=" & oRsTmpOpc1.Fields!IdVariable
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdVariable = oRsTmpOpc1.Fields!IdVariable
                oRsTmpOpc.Fields!IdDominio = oRsTmpOpc1.Fields!IdDominio
                oRsTmpOpc.Fields!OrdernDominio = oRsTmpOpc1.Fields!OrdernDominio
                oRsTmpOpc.Fields!Texto = oRsTmpOpc1.Fields!Texto
                oRsTmpOpc.Fields!Tipo = oRsTmpOpc1.Fields!Tipo
                oRsTmpOpc.Fields!Ancho = oRsTmpOpc1.Fields!Ancho
                oRsTmpOpc.Fields!EsDatoObligatorio = oRsTmpOpc1.Fields!EsDatoObligatorio
                oRsTmpOpc.Fields!TextoTooltip = oRsTmpOpc1.Fields!TextoTooltip
                oRsTmpOpc.Fields!EsDatoGrafico = oRsTmpOpc1.Fields!EsDatoGrafico
                oRsTmpOpc.Fields!TieneFormatoMask = oRsTmpOpc1.Fields!TieneFormatoMask
                oRsTmpOpc.Fields!FormatoMask = oRsTmpOpc1.Fields!FormatoMask
                oRsTmpOpc.Fields!TieneRango = oRsTmpOpc1.Fields!TieneRango
                oRsTmpOpc.Fields!RangoInicial = oRsTmpOpc1.Fields!RangoInicial
                oRsTmpOpc.Fields!RangoFinal = oRsTmpOpc1.Fields!RangoFinal
                oRsTmpOpc.Fields!PosicionFila = oRsTmpOpc1.Fields!PosicionFila
                oRsTmpOpc.Fields!PosicionColumna = oRsTmpOpc1.Fields!PosicionColumna
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 204
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_CatalogoValoresCombo"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_CatalogoValoresCombo] (" & _
            "    [IdVariable] [int] NOT NULL ," & _
            "    [IdValorCombo] [int] NOT NULL ," & _
            "    [ComboTexto] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & _
            "    CONSTRAINT [PK_Enfermeria_VariablesValoresCombo] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdVariable]," & _
            "        [IdValorCombo]" & _
            "    )  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_CatalogoVariables_Enfermeria_CatalogoValoresCombo] FOREIGN KEY" & _
            "    (" & _
            "        [IdVariable]" & _
            "    ) REFERENCES [dbo].[Enfermeria_CatalogoVariables] (" & _
            "       [IdVariable]" & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoValoresCombo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from Enfermeria_CatalogoValoresCombo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from Enfermeria_CatalogoValoresCombo where IdVariable=" & _
                                         oRsTmpOpc1.Fields!IdVariable & " and IdValorCombo=" & oRsTmpOpc1.Fields!IdValorCombo
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdVariable = oRsTmpOpc1.Fields!IdVariable
                oRsTmpOpc.Fields!IdValorCombo = oRsTmpOpc1.Fields!IdValorCombo
                oRsTmpOpc.Fields!ComboTexto = oRsTmpOpc1.Fields!ComboTexto
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 205
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_Visitas"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_Visitas] (" & _
            "    [IdCuentaAtencion] [int] NOT NULL ," & _
            "    [IdVisita] [int] NOT NULL ," & _
            "    [FechaHoraVisita] [datetime] NOT NULL ," & _
            "    [IdServicio] [int] NOT NULL ," & _
            "    [Observaciones] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [IdCama] [int] NOT NULL ," & _
            "    [IdEmpleadoEnfermera] [int] NOT NULL ," & _
            "    [IngresoValorizacion] [Bit]," & _
            "    CONSTRAINT [PK_Enfermeria_Visitas] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdCuentaAtencion]," & _
            "        [IdVisita]" & _
            "    )  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_Visitas_Camas] FOREIGN KEY" & _
            "    ([IdCama]" & _
            "    ) REFERENCES [dbo].[Camas] (" & _
            "        [IdCama]" & _
            "    )," & _
            "    CONSTRAINT [FK_Enfermeria_Visitas_FacturacionCuentasAtencion] FOREIGN KEY" & _
            "    ([IdCuentaAtencion]" & _
            "    ) REFERENCES [dbo].[FacturacionCuentasAtencion] (" & _
            "        [IdCuentaAtencion])" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "Alter table Enfermeria_Visitas drop CONSTRAINT FK_Enfermeria_Visitas_Camas"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 206
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_Variables"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_Variables] (" & _
            "    [IdCuentaAtencion] [int] NOT NULL ," & _
            "    [IdVisita] [int] NOT NULL ," & _
            "    [IdVariable] [int] NOT NULL ," & _
            "    [VariableDato] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
            "    CONSTRAINT [PK_Enfermeria_Variables] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdCuentaAtencion]," & _
            "        [IdVisita]," & _
            "        [IdVariable]" & _
            "    )  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_Variables_Enfermeria_CatalogoVariables] FOREIGN KEY" & _
            "    ([IdVariable]" & _
            "    ) REFERENCES [dbo].[Enfermeria_CatalogoVariables] (" & _
            "        [IdVariable]" & _
            "    )," & _
            "    CONSTRAINT [FK_Enfermeria_Variables_Enfermeria_Visitas] FOREIGN KEY" & _
            "    ([IdCuentaAtencion]," & _
            "        [IdVisita]" & _
            "    ) REFERENCES [dbo].[Enfermeria_Visitas] (" & _
            "        [IdCuentaAtencion]," & _
            "        [IdVisita]" & _
            "    )" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic


    DoEvents
    ProgressBar1.Value = 207
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_ValoresCombo"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_ValoresCombo] (" & _
            "    [IdCuentaAtencion] [int] NOT NULL ," & _
            "    [IdVisita] [int] NOT NULL ," & _
            "    [IdVariable] [int] NOT NULL ," & _
            "    [IdValorCombo] [Int] NOT NULL ," & _
            "    CONSTRAINT [PK_Enfermeria_ValoresCombo] PRIMARY KEY  CLUSTERED" & _
            "    (   [IdCuentaAtencion]," & _
            "        [IdVisita]," & _
            "        [IdVariable]," & _
            "        [IdValorCombo]" & _
            "    )  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_ValoresCombo_Enfermeria_CatalogoValoresCombo] FOREIGN KEY" & _
            "    (   [IdVariable],[IdValorCombo]" & _
            "    ) REFERENCES [dbo].[Enfermeria_CatalogoValoresCombo] (" & _
            "        [IdVariable],[IdValorCombo]" & _
            "    )," & _
            "    CONSTRAINT [FK_Enfermeria_ValoresCombo_Enfermeria_Variables] FOREIGN KEY" & _
            "    (   [IdCuentaAtencion],[IdVisita],[IdVariable]" & _
            "    ) REFERENCES [dbo].[Enfermeria_Variables] (" & _
            "        [IdCuentaAtencion],[IdVisita],[IdVariable] " & _
            "    ) " & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 208
    Me.Refresh
    txtTablaProceso.Text = "Enfermeria_TratamientoDosis"
    lcSql = "CREATE TABLE [dbo].[Enfermeria_TratamientoDosis] (" & _
            "    [IdCuentaAtencion] [int] NOT NULL ," & _
            "    [IdVisita] [int] NOT NULL ," & _
            "    [IdDiaVisita] [int] NOT NULL ," & _
            "    [IdReceta] [int] NOT NULL ," & _
            "    [IdItem] [int] NOT NULL ," & _
            "    [Dosis] [int] NULL ," & _
            "    [DatoProrenata] [int] NULL," & _
            "    CONSTRAINT [PK_Enfermeria_TratamientoDosis] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdCuentaAtencion]," & _
            "        [IdVisita]," & _
            "        [IdDiaVisita]," & _
            "        [IdReceta]," & _
            "        [idItem]" & _
            "    )  ON [PRIMARY]," & _
            "    CONSTRAINT [FK_Enfermeria_TratamientoDosis_Enfermeria_Visitas] FOREIGN KEY" & _
            "    (   [IdCuentaAtencion],[IdVisita]" & _
            "    ) REFERENCES [dbo].[Enfermeria_Visitas] (" & _
            "        [IdCuentaAtencion],[IdVisita])," & _
            "    CONSTRAINT [FK_Enfermeria_TratamientoDosis_RecetaDetalle] FOREIGN KEY " & _
            "    (   [IdReceta],[idItem]" & _
            "    ) REFERENCES [dbo].[RecetaDetalle] (" & _
            "        [idReceta],[idItem]) " & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
       
    DoEvents
    ProgressBar1.Value = 209
    Me.Refresh
    txtTablaProceso.Text = "TiposCondicionPaciente y HIS_Paciente"
    
    lcSql = "Alter table TiposCondicionPaciente ADD OrdenRegHis int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    lcSql = "select * from TiposCondicionPaciente"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposCondicionPaciente"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdTipoCondicionPaciente=" & oRsTmpOpc1.Fields!IdTipoCondicionPaciente
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!OrdenRegHis = oRsTmpOpc1.Fields!OrdenRegHis
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    lcSql = "Alter table HIS_Paciente drop CONSTRAINT HIS_Paciente_Pacientes"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 210
    Me.Refresh
    txtTablaProceso.Text = "FactPartidasPresupuestalesXMes"
    lcSql = "CREATE TABLE [dbo].[FactPartidasPresupuestalesXMes] (" & _
            "    [Fecha] [datetime] NOT NULL ," & _
            "    [idPartida] [int] NOT NULL ," & _
            "    [IdProducto] [int] NOT NULL ," & _
            "    [ImpAnulado] [money] NOT NULL ," & _
            "    [impExonerado] [money] NOT NULL ," & _
            "    [ImpNormal] [money] NOT NULL ," & _
            "    [ImpCancelado]  [Money] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[FactPartidasPresupuestalesXMes] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_FactPartidasPresupuestalesXMes] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [Fecha]," & _
            "    [idPartida]," & _
            "    [IdProducto]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 211
    Me.Refresh
    txtTablaProceso.Text = "farmInventario"
    lcSql = "Alter table farmInventario ADD idTipoInventario int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update farmInventario set idTipoInventario=1 where idTipoInventario is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 212
    Me.Refresh
    txtTablaProceso.Text = "farmTipoInventario"
    lcSql = "CREATE TABLE [dbo].[farmTipoInventario] (" & _
            "    [idTipoInventario] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [TipoInventarioSismed] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[farmTipoInventario] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_farmTipoInventario] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idTipoInventario]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmTipoInventario"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from farmTipoInventario"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idTipoInventario=" & oRsTmpOpc1.Fields!idTipoInventario
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idTipoInventario = oRsTmpOpc1.Fields!idTipoInventario
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Fields!TipoInventarioSismed = oRsTmpOpc1.Fields!TipoInventarioSismed
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 213
    Me.Refresh
    txtTablaProceso.Text = "farmInventarioCabecera"
    lcSql = "Alter table farmInventarioCabecera ADD CantidadSaldo int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table farmInventarioCabecera ADD CantidadFaltante int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table farmInventarioCabecera ADD CantidadSobrante int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 214
    Me.Refresh
    txtTablaProceso.Text = "farmInventarioDetalle"
    lcSql = "Alter table farmInventarioDetalle ADD CantidadSaldo int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table farmInventarioDetalle ADD CantidadFaltante int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table farmInventarioDetalle ADD CantidadSobrante int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table farmInventarioDetalle ADD EsHistoricoSaldo int null"     'debb2014b
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close                                   'debb2014b
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic              'debb2014b
    '
    DoEvents
    ProgressBar1.Value = 215
    Me.Refresh
    txtTablaProceso.Text = "farmHistPrecio"
    lcSql = "CREATE TABLE [dbo].[farmHistPrecio] (" & _
            "    [idHistPrecio] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [idProducto] [int] NOT NULL ," & _
            "    [fecha] [datetime] NOT NULL ," & _
            "    [PrecioCompra] [money] NOT NULL ," & _
            "    [PrecioDistribucion] [money] NOT NULL ," & _
            "    [PrecioVenta] [money] NOT NULL ," & _
            "    [PrecioDonacion] [money] NOT NULL ," & _
            "    [IdUsuario]  [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[farmHistPrecio] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_farmHistPrecio] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idHistPrecio]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index IX1_farmHistPrecio on farmHistPrecio (idProducto)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 216
    Me.Refresh
    txtTablaProceso.Text = "AtenHospCenso"
    lcSql = "CREATE TABLE [dbo].[AtenHospCenso](" & _
            "    [IdRangoCensoHosp] [int] NOT NULL," & _
            "    [RangoInicial] [money] NOT NULL," & _
            "    [RangoFinal] [money] NOT NULL," & _
            "    [RGBRojo] [int] NULL," & _
            "    [RGBVerde] [int] NULL," & _
            "    [RGBAzul] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenHospCenso"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenHospCenso"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from AtenHospCenso where IdRangoCensoHosp=" & _
                                         oRsTmpOpc1.Fields!IdRangoCensoHosp
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdRangoCensoHosp = oRsTmpOpc1.Fields!IdRangoCensoHosp
           End If
            oRsTmpOpc.Fields!RangoInicial = oRsTmpOpc1.Fields!RangoInicial
            oRsTmpOpc.Fields!RangoFinal = oRsTmpOpc1.Fields!RangoFinal
            oRsTmpOpc.Fields!RGBRojo = oRsTmpOpc1.Fields!RGBRojo
            oRsTmpOpc.Fields!RGBVerde = oRsTmpOpc1.Fields!RGBVerde
            oRsTmpOpc.Fields!RGBAzul = oRsTmpOpc1.Fields!RGBAzul
            oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
       
    Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub


Sub cmdMigraUltimaVErsionExternaSamuel(oConexODBC As Connection, _
                                 oConexHBT As Connection)
    On Error GoTo errMgSS
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim lcSql As String, lbNuevoRegistro As Boolean
    
    '
    DoEvents
    ProgressBar1.Value = 283
    Me.Refresh
    txtTablaProceso.Text = "FuaDefaultsCptFarmacia"
    
    lcSql = " CREATE TABLE [dbo].[FuaDefaultsCptFarmacia](" & _
            "     [codigo] [varchar](20) NULL," & _
            "     [tipo] [char](10) NULL," & _
            "     [idPuntoCarga] [int] NULL," & _
            "     [id] [int] IDENTITY(1,1) NOT NULL," & _
            "     [EsMedicamento] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table FuaDefaultsCptFarmacia ADD id int IDENTITY (1, 1) NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "Alter table FuaDefaultsCptFarmacia ADD EsMedicamento int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 284
    Me.Refresh
    txtTablaProceso.Text = "SisFuaEstadosTrama"
    lcSql = " CREATE TABLE [dbo].[SisFuaEstadosTrama] (" & _
            " [Id] [int] IDENTITY (1, 1) NOT NULL ," & _
            " [Tabla] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [Campo] [varchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [Estado] [bit] NOT NULL ," & _
            " [CampoCondicion] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [Valor] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [Obligatorio] [varchar] (15) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [Formato] [varchar] (50) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [Observaciones] [varchar] (60) COLLATE Modern_Spanish_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "Alter table SisFuaEstadosTrama ADD Orden int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "select * from SisFuaEstadosTrama"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from SisFuaEstadosTrama"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "id=" & oRsTmpOpc1.Fields!ID
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!ID = oRsTmpOpc1.Fields!ID
           End If
           oRsTmpOpc.Fields!tabla = oRsTmpOpc1.Fields!tabla
           oRsTmpOpc.Fields!campo = oRsTmpOpc1.Fields!campo
           oRsTmpOpc.Fields!estado = oRsTmpOpc1.Fields!estado
           oRsTmpOpc.Fields!CampoCondicion = oRsTmpOpc1.Fields!CampoCondicion
           oRsTmpOpc.Fields!valor = oRsTmpOpc1!valor
           oRsTmpOpc.Fields!Obligatorio = oRsTmpOpc1.Fields!Obligatorio
           oRsTmpOpc.Fields!Formato = oRsTmpOpc1.Fields!Formato
           oRsTmpOpc.Fields!Observaciones = oRsTmpOpc1.Fields!Observaciones
           oRsTmpOpc.Fields!Orden = oRsTmpOpc1.Fields!Orden
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    
    DoEvents
    ProgressBar1.Value = 285
    Me.Refresh
    txtTablaProceso.Text = "SisFuaResumen"
    lcSql = "CREATE TABLE [dbo].[SisFuaResumen] (" & _
            " [idResumen] [int] IDENTITY (1, 1) NOT NULL ," & _
            " [Anio] [varchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [Mes] [varchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [NroEnvio] [varchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [NomPaquete] [varchar] (18) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [VersionGTI] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [CantFilATE] [int] NULL ," & _
            " [CantFilSMI] [int] NULL ," & _
            " [CantFilDIA] [int] NULL ," & _
            " [CantFilMED] [int] NULL ," & _
            " [CantFilINS] [int] NULL ," & _
            " [CantFilPRO] [int] NULL ," & _
            " [CantFilUSU] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
     DoEvents
    ProgressBar1.Value = 286
    Me.Refresh
    txtTablaProceso.Text = "SisFuaUsuario"
    lcSql = "CREATE TABLE [dbo].[SisFuaUsuario] (" & _
            " [idUsuario] [int] IDENTITY (1, 1) NOT NULL ," & _
            " [DNI] [varchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            " [TipoDoc] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [ApellidoPat] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [ApellidoMat] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [PrimerNombre] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [SegundoNombre] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [NroEnvio] [int] NULL ," & _
            " [Periodo] [varchar] (4) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [Mes] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            " [CodigoEstablecimiento] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    Exit Sub
errMgSS:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
                                 
End Sub


Sub cmdMigraUltimaVErsionExterna(oConexODBC As Connection, _
                                 oConexHBT As Connection)
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim lbNuevoRegistro As Boolean

    On Error GoTo errMg2
    
    '*********************** aqui empieza SIGH_EXTERNA   ********************************
    oConexODBC.Close
    oConexODBC.CommandTimeout = 300
    oConexODBC.Open "dsn=GalenhosExterna"
    '
    
    cmdMigraUltimaVErsionExternaSamuel oConexODBC, oConexHBT
    
    
    DoEvents
    ProgressBar1.Value = 287
    Me.Refresh
    txtTablaProceso.Text = "CitasWebEstados"
    lcSql = "CREATE TABLE [dbo].[CitasWebEstados] (" & _
            "    [idEstadoCitaWeb] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (30) COLLATE Modern_Spanish_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[CitasWebEstados] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_CitasWebEstados] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idEstadoCitaWeb]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from CitasWebEstados"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from CitasWebEstados"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "idEstadoCitaWeb=" & oRsTmpOpc1.Fields!idEstadoCitaWeb
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstadoCitaWeb = oRsTmpOpc1.Fields!idEstadoCitaWeb
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
                
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 288
    Me.Refresh
    txtTablaProceso.Text = "CitasWebCupos"
    lcSql = "CREATE TABLE [dbo].[CitasWebCupos] (" & _
            "    [Fecha] [datetime] NULL ," & _
            "    [idServicio] [int] NULL ," & _
            "    [idMedico] [int] NULL ," & _
            "    [HoraInicio] [varchar] (5) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [HoraFinal] [varchar] (5) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [idEstadoCitaWeb] [int] NULL ," & _
            "    [idCitaBloqueada] [int] NULL ," & _
            "    [DNI] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ApellidoPaterno] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ApellidoMaterno] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [PrimerNombre] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [SegundoNombre] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [idTipoSexo] [int] NULL ," & _
            "    [FechaNacimiento] [datetime] NULL ," & _
            "    [Ubigeo] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD FechaConfirmacion datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADDD HoraConfirmacion varchar(5) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD idFuenteFinanciamiento int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX IX_Fecha  ON CitasWebCupos (Fecha)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX IX_idCitaBloqueada  ON CitasWebCupos (idCitaBloqueada)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD Email varchar(50) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD Telefono varchar(10) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX IX_Fecha1  ON CitasWebCupos (Fecha,idServicio,idMedico,HoraInicio)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD HoraConfirmacion varchar(5) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD idWeb int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD idTurno int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table CitasWebCupos ADD idPaciente int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 289
    Me.Refresh
    txtTablaProceso.Text = "Elimina tablas SIS"
    lcSql = "drop table Sis_a_categoriaeess"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table Sis_m_servicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table Sis_a_tipodocumento"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table Sis_a_destinoasegurado"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "drop table Sis_a_modalidadatencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 290
    Me.Refresh
    txtTablaProceso.Text = "SisFua"
    lcSql = "CREATE TABLE [dbo].[SisFua] (" & _
            "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNumeroInicial] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNumeroFinal] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaUltimoGenerado] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFua alter column FuaNumeroInicial varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFua alter column FuaNumeroFinal varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFua alter column FuaUltimoGenerado varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 291
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencion"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencion] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [EstablecimientoCodigoSIS] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Reconsideracion] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ReconsideracionCodigoDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ReconsideracionLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ReconsideracionNroFormato] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaComponente] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Situacion] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [AfiliacionDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [AfiliacionTipoFormato] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [AfiliacionNroFormato] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CodigoTipoFormato] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [OrigenAseguradoInstitucion] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [OrigenAseguradoCodigo] [varchar] (16) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Edad] [int] NULL ," & _
            "    [GrupoEtareo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Genero] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaAtencion] [int] NULL ," & _
            "    [FuaCondicionMaterna] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNrohistoria] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,"
    lcSql = lcSql & " [FuaConceptoPr] [int] NULL ," & _
            "    [FuaConceptoPrAutoriz] [varchar] (15) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaConceptoPrMonto] [money] NULL ," & _
            "    [FuaAtencionFecha] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaAtencionHora] [varchar] (5) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaReferidoOrigenCodigoSIS] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaReferidoOrigenNreferencia] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaCodigoPrestacion] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaPersonalQatiende] [int] NULL ," & _
            "    [FuaAtencionLugar] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaDestino] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaHospitalizadoFingreso] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaHospitalizadoFalta] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaReferidoDestinoCodigoRenaes] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaReferidoDestinoNreferencia] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaMedicoDNI] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaMedico] [varchar] (120) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaMedicoTipo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [AfiliacionNroIntegrante] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Codigo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [idSiasis] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaObservaciones] [varchar] (200) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [UltimaFechaAddMod] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,"
    lcSql = lcSql & " [CabEstado] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaFechaParto] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [EstablecimientoDistrito] [varchar] (6) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Anio] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Mes] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CostoTotal] [money] NULL ," & _
            "    [Apaterno] [varchar] (40) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Amaterno] [varchar] (40) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Pnombre] [varchar] (35) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Onombre] [varchar] (35) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [fnacimiento] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Autogenerado] [varchar] (20) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [DocumentoTipo] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [DocumentoNumero] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [EstablecimientoCategoria] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CostoServicio] [money] NULL ," & _
            "    [CostoMedicamento] [money] NULL ," & _
            "    [CostoProcedimiento] [money] NULL ," & _
            "    [CostoInsumo] [money] NULL ," & _
            "    [MedicoDocumentoTipo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [ate_grupoRiesgo] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
            "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ,"
    
    lcSql = lcSql & " [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabIdentificacionPaquete] [int] NULL ," & _
            "    [IdentificacionArfsis] [int] NULL ," & _
            "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [PeriodoOrigen] [varchar] (6) COLLATE Modern_Spanish_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencion] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_SisFuaAtencion] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idCuentaAtencion]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indNumeroFua  ON SisFuaAtencion (FuaDisa,FuaLote,FuaNumero)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencion alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    '
    DoEvents
    ProgressBar1.Value = 292
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionDIA"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionDIA] (" & _
            "    [id] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [DxNumero] [int] NOT NULL ," & _
            "    [DxTipoIE] [varchar] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [DxCodigo] [varchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [DxTipoDPR] [varchar] (1) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabEstado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
            "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabIdentificacionPaquete] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdCuenta  ON SisFuaAtencionDIA (idCuentaAtencion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencionDIA alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 293
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionINS"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionINS] (" & _
            "    [id] [int] IDENTITY (1, 1) NOT NULL ," & _
            "    [idTablaDx] [int] NOT NULL ," & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [DxNumero] [int] NOT NULL ," & _
            "    [Codigo] [varchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [CantidadPrescrita] [int] NOT NULL ," & _
            "    [CantidadEntregada] [int] NOT NULL ," & _
            "    [PrecioUnitario] [money] NULL ," & _
            "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabEstado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
            "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CabIdentificacionPaquete] [int] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdCuenta  ON SisFuaAtencionINS (idCuentaAtencion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencionINS alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 294
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionMED"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionMED] (" & _
                "    [id] [int] IDENTITY (1, 1) NOT NULL ," & _
                "    [idTablaDx] [int] NOT NULL ," & _
                "    [idCuentaAtencion] [int] NOT NULL ," & _
                "    [Codigo] [varchar] (5) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
                "    [DxNumero] [int] NOT NULL ," & _
                "    [CantidadPrescrita] [int] NOT NULL ," & _
                "    [CantidadEntregada] [int] NOT NULL ," & _
                "    [PrecioUnitario] [money] NULL ," & _
                "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabEstado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
                "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabIdentificacionPaquete] [int] NULL" & _
                " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdCuenta  ON SisFuaAtencionMED (idCuentaAtencion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencionMED alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 295
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionPRO"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionPRO] (" & _
                "    [id] [int] IDENTITY (1, 1) NOT NULL ," & _
                "    [idTablaDx] [int] NOT NULL ," & _
                "    [idCuentaAtencion] [int] NOT NULL ," & _
                "    [Codigo] [varchar] (15) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
                "    [DxNumero] [int] NOT NULL ," & _
                "    [CantidadPrescrita] [int] NOT NULL ," & _
                "    [CantidadEjecutada] [int] NOT NULL ," & _
                "    [PrecioUnitario] [money] NULL ," & _
                "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabEstado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [Resultado] [varchar] (15) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
                "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabIdentificacionPaquete] [int] NULL" & _
                " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdCuenta  ON SisFuaAtencionPRO (idCuentaAtencion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencionPRO alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 296
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionSMI"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionSMI] (" & _
                "    [id] [int] IDENTITY (1, 1) NOT NULL ," & _
                "    [idCuentaAtencion] [int] NOT NULL ," & _
                "    [IntervencionesPreventivas] [varchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
                "    [Valor] [varchar] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
                "    [CabDniUsuarioRegistra] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabFechaFuaPrimeraVez] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabEstado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabNroEnvioAlSIS] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabCodigoPuntoDigitacion] [int] NULL ," & _
                "    [CabCodigoUDR] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaLote] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [FuaNumero] [varchar] (8) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabOrigenDelRegistro] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabVersionAplicativo] [varchar] (9) COLLATE Modern_Spanish_CI_AS NULL ," & _
                "    [CabIdentificacionPaquete] [int] NULL" & _
                " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdCuenta  ON SisFuaAtencionSMI (idCuentaAtencion)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFuaAtencionSMI alter column FuaNumero varchar(16)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 297
    Me.Refresh
    txtTablaProceso.Text = "SisFiliaciones"
    lcSql = "CREATE TABLE [dbo].[SisFiliaciones] (" & _
            "    [idSiasis] [int] NOT NULL," & _
            "    [Codigo] [varchar] (2)  NOT NULL," & _
            "    [AfiliacionDisa] [varchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [AfiliacionTipoFormato] [varchar] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [AfiliacionNroFormato] [varchar] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [AfiliacionNroIntegrante] [varchar] (2) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [DocumentoTipo] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [CodigoEstablAdscripcion] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [AfiliacionFecha] [datetime] NULL ," & _
            "    [Paterno] [varchar] (40) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [Materno] [varchar] (40) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [Pnombre] [varchar] (70) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [Onombres] [varchar] (70) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Genero] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Fnacimiento] [datetime] NULL ," & _
            "    [IdDistritoDomicilio] [varchar] (6) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Estado] [varchar] (1) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [Fbaja] [datetime] NULL ," & _
            "    [DocumentoNumero] [varchar] (10) COLLATE Modern_Spanish_CI_AS NULL ," & _
            "    [MotivoBaja] [varchar] (70) COLLATE Modern_Spanish_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[SisFiliaciones] ADD " & _
            "    CONSTRAINT [PK_SisFiliaciones] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idSiaSis]," & _
            "        [Codigo]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table SisFiliaciones alter column Fbaja varchar(10)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE SisFiliaciones add  FbajaOK datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indDcto  ON SisFiliaciones (DocumentoNumero)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indApellidos  ON SisFiliaciones (Paterno,Materno,Pnombre)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indAfiliacion  ON SisFiliaciones (AfiliacionDisa,AfiliacionTipoFormato,AfiliacionNroFormato)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE INDEX indIdSiasis  ON SisFiliaciones (idSiasis)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 298
    Me.Refresh
    txtTablaProceso.Text = "atencionesCE"
    lcSql = "ALTER TABLE atencionesCE add  TriajePulso int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "ALTER TABLE atencionesCE add  TriajeFrecRespiratoria int null"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesCE add  CitaAntecedente varchar(1000) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesCE ADD TriajePerimCefalico money NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesCE ADD TriajeFrecCardiaca int NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesCE ADD TriajeOrigen int NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 299
    Me.Refresh
    
    txtTablaProceso.Text = "TriajeVariable"
    lcSql = "CREATE TABLE [dbo].[TriajeVariable] (" & _
            "    [IdTriajeVariable] [int] IDENTITY (1, 1) NOT NULL," & _
            "    [TriajeVariable] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL," & _
            "    [EsAntropometrica] [bit] NOT NULL ," & _
            "    [TieneLimiteMedicion] [bit] NOT NULL ," & _
            "    [EdadDiaLimiteMinima] [int] NULL ," & _
            "    [EdadDiaLimiteMaxima] [int] NULL ," & _
            "    [EsDatoObligatorio] [bit] NOT NULL ," & _
            "    [EsActivo] [bit] NOT NULL ," & _
            "    CONSTRAINT [PK_VariablesAntropometricas] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdTriajeVariable] " & _
            "    )  ON [PRIMARY]" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    lcSql = "ALTER TABLE [dbo].[TriajeVariable] ADD " & _
            "    CONSTRAINT [DF_VariableTriaje_EsAntropometrica] DEFAULT (0) FOR [EsAntropometrica] , " & _
            "    CONSTRAINT [DF_VariableTriaje_TieneLimiteMedicion] DEFAULT (0) FOR [TieneLimiteMedicion], " & _
            "    CONSTRAINT [DF_VariableTriaje_EsDatoObligatorio] DEFAULT (1) FOR [EsDatoObligatorio], " & _
            "    CONSTRAINT [DF_VariableTriaje_EsActivo] DEFAULT (1) FOR [EsActivo] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from TriajeVariable"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TriajeVariable"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdTriajeVariable=" & oRsTmpOpc1.Fields!IdTriajeVariable
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdTriajeVariable = oRsTmpOpc1.Fields!IdTriajeVariable
           End If
           oRsTmpOpc.Fields!TriajeVariable = oRsTmpOpc1.Fields!TriajeVariable
           oRsTmpOpc.Fields!EsAntropometrica = oRsTmpOpc1.Fields!EsAntropometrica
           oRsTmpOpc.Fields!TieneLimiteMedicion = oRsTmpOpc1.Fields!TieneLimiteMedicion
           oRsTmpOpc.Fields!EdadDiaLimiteMinima = oRsTmpOpc1.Fields!EdadDiaLimiteMinima
           oRsTmpOpc.Fields!EdadDiaLimiteMaxima = oRsTmpOpc1.Fields!EdadDiaLimiteMaxima
           oRsTmpOpc.Fields!EsDatoObligatorio = oRsTmpOpc1.Fields!EsDatoObligatorio
           oRsTmpOpc.Fields!EsActivo = oRsTmpOpc1.Fields!EsActivo
            oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 300
    Me.Refresh
    
    txtTablaProceso.Text = "TriajeValorNormal"
    lcSql = "CREATE TABLE [dbo].[TriajeValorNormal] (" & _
            "    [IdTriajeValorNormal] [int] IDENTITY (1, 1) NOT NULL, " & _
            "    [EdadInicialEnDia] [int] NOT NULL, " & _
            "    [EdadFinalEnDia] [int] NULL ," & _
            "    [ValorNormalMinimo] [money] NULL ," & _
            "    [ValorNormalMaximo] [money] NULL ," & _
            "    [ValorCoherenteMinimo] [money] NULL ," & _
            "    [ValorCoherenteMaximo] [money] NULL ," & _
            "    [IdTriajeVariable] [int] NOT NULL ," & _
            "    [EstadoPaciente] [int] NULL ," & _
            "    [SexoPaciente] [int] NOT NULL ," & _
            "    [FechaVigencia] [datetime] NOT NULL ," & _
            "    CONSTRAINT [PK_TriajeValorNormal] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdTriajeValorNormal] " & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_TriajeValorNormal_TriajeVariable] FOREIGN KEY " & _
            "    (" & _
            "        [IdTriajeVariable] " & _
            "    )  REFERENCES [dbo].[TriajeVariable] (" & _
            "    [IdTriajeVariable] " & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from TriajeValorNormal"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TriajeValorNormal"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdTriajeValorNormal=" & oRsTmpOpc1.Fields!IdTriajeValorNormal
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdTriajeValorNormal = oRsTmpOpc1.Fields!IdTriajeValorNormal
           End If
           oRsTmpOpc.Fields!EdadInicialEnDia = oRsTmpOpc1.Fields!EdadInicialEnDia
           oRsTmpOpc.Fields!EdadFinalEnDia = oRsTmpOpc1.Fields!EdadFinalEnDia
           oRsTmpOpc.Fields!ValorNormalMinimo = oRsTmpOpc1.Fields!ValorNormalMinimo
           oRsTmpOpc.Fields!ValorNormalMaximo = oRsTmpOpc1.Fields!ValorNormalMaximo
           oRsTmpOpc.Fields!ValorCoherenteMinimo = oRsTmpOpc1.Fields!ValorCoherenteMinimo
           oRsTmpOpc.Fields!ValorCoherenteMaximo = oRsTmpOpc1.Fields!ValorCoherenteMaximo
           oRsTmpOpc.Fields!IdTriajeVariable = oRsTmpOpc1.Fields!IdTriajeVariable
           oRsTmpOpc.Fields!EstadoPaciente = oRsTmpOpc1.Fields!EstadoPaciente
           oRsTmpOpc.Fields!SexoPaciente = oRsTmpOpc1.Fields!SexoPaciente
           oRsTmpOpc.Fields!FechaVigencia = oRsTmpOpc1.Fields!FechaVigencia
            oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 300
    Me.Refresh
    
    txtTablaProceso.Text = "TriajeExcepciones"
    lcSql = "CREATE TABLE [dbo].[TriajeExcepciones] (" & _
            "    [IdTriajeExcepciones] [int] IDENTITY (1, 1) NOT NULL, " & _
            "    [IdTriajeVariable] [int] NOT NULL ," & _
            "    [EdadInicialEnDia] [int] NOT NULL ," & _
            "    [EdadFinalEnDia] [int] NOT NULL ," & _
            "    [EsDatoObligatorio] [bit] NOT NULL ," & _
            "    CONSTRAINT [PK_TriajeExcepciones] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdTriajeExcepciones] " & _
            "    )  ON [PRIMARY] ," & _
            "    CONSTRAINT [FK_TriajeExcepciones_TriajeVariable] FOREIGN KEY " & _
            "    (" & _
            "        [IdTriajeVariable] " & _
            "    )  REFERENCES [dbo].[TriajeVariable] (" & _
            "    [IdTriajeVariable] " & _
            "    )" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from TriajeExcepciones"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TriajeExcepciones"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdTriajeExcepciones=" & oRsTmpOpc1.Fields!IdTriajeExcepciones
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!IdTriajeVariable = oRsTmpOpc1.Fields!IdTriajeVariable
           oRsTmpOpc.Fields!EdadInicialEnDia = oRsTmpOpc1.Fields!EdadInicialEnDia
           oRsTmpOpc.Fields!EdadFinalEnDia = oRsTmpOpc1.Fields!EdadFinalEnDia
           oRsTmpOpc.Fields!EsDatoObligatorio = oRsTmpOpc1.Fields!EsDatoObligatorio
            oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    cmdMigraUltimaVErsionExternaFUA_Parte1 oConexODBC, oConexHBT 'Barra de Proceso al 301
    cmdMigraUltimaVErsionExternaFUA_Parte2 oConexODBC, oConexHBT 'Barra de Proceso del 302 al 317 *Cambio Nuevo FUA2015
    
    'Actualiza version de BD en parametros
    oConexODBC.Close
    oConexODBC.Open "dsn=Galenhos"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    lcSql = "update parametros set ValorTexto ='" & wxVersionBDactualizada & "' where idParametro=314"
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    Exit Sub
errMg2:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If

End Sub








































Private Sub Form_Load()
    
    
    '
    '
    
    '
    
    '
    CargaVersionSQL
    If wxVersionSQL = sghVersionBD.sighSql2000 Then
       txtSql2000.Visible = True
       Me.Caption = Me.Caption & " (BD SQL2000)"
    Else
       txtSql2008.Visible = True
       Me.Caption = Me.Caption & " (BD SQL2008)"
    End If
End Sub

 Sub CargaVersionSQL()
    On Error GoTo ErrCarVers
    wxVersionSQL = sghVersionBD.sighSql2000
    Dim oRsTmp1 As New Recordset
    Dim oConexODBC As New Connection
    lcSql = "SELECT @@Version as VersionServidor"
    oConexODBC.CommandTimeout = 300
    oConexODBC.Open "dsn=GALENHOS"
    oRsTmp1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    If oRsTmp1.RecordCount > 0 Then
       If InStr(oRsTmp1!versionServidor, "SQL Server  2000") = 0 Then
          wxVersionSQL = sghVersionBD.sighSql2008
       End If
    End If
ErrCarVers:
    Set oConexODBC = Nothing
    Set oRsTmp1 = Nothing
End Sub

Sub EliminaProcedAlmacenados()
    Dim oProcedimiento As ADOX.Procedure
    Dim oCatalogo As New ADOX.Catalog
    Dim oRsTmp As New Recordset
    Dim lcSql As String, sNombre As String
    Dim oConexODBC1 As New Connection
    On Error GoTo ErrEPA
    'PA de sigh
    oConexODBC1.Open "dsn=GALENHOS"
    oCatalogo.ActiveConnection = oConexODBC1
    For Each oProcedimiento In oCatalogo.Procedures
       If Left(oProcedimiento.Name, 3) <> "dt_" Then
            sNombre = Left(oProcedimiento.Name, InStr(oProcedimiento.Name, ";") - 1)
            lcSql = "DROP PROCEDURE " & sNombre
            oRsTmp.Open lcSql, oConexODBC1, adOpenKeyset, adLockOptimistic
       End If
    Next
    'PA de sigh_externa
    oConexODBC1.Close
    oConexODBC1.Open "dsn=GalenhosExterna"
    oCatalogo.ActiveConnection = oConexODBC1
    For Each oProcedimiento In oCatalogo.Procedures
       If Left(oProcedimiento.Name, 3) <> "dt_" Then
            sNombre = Left(oProcedimiento.Name, InStr(oProcedimiento.Name, ";") - 1)
            lcSql = "DROP PROCEDURE " & sNombre
            oRsTmp.Open lcSql, oConexODBC1, adOpenKeyset, adLockOptimistic
       End If
    Next
    oConexODBC1.Close
    Set oConexODBC1 = Nothing
    '
    Exit Sub
ErrEPA:
    'MsgBox "Error al ELIMINAR PROCEDIMIENTOS ALMACENADOS" & Chr(13) & Err.Number & " - " & Err.Description
    Resume Next
End Sub


Sub DepuraColumnasDeTablaAtenciones()
    On Error GoTo ErrColuAten
    txtTablaProceso.Text = "Elimina COLUMNAS de tabla ATENCIONES y lo pasa a tabla AtencionesDatosAdicionales"
    DoEvents
    Me.Refresh
    Dim oConexODBC As New Connection
    Dim oRsTmpOpc1 As New Recordset
    Dim oRsTmpOpc2 As New Recordset
    Dim lnRegistros As Long
    oConexODBC.CommandTimeout = 300
    oConexODBC.Open "dsn=GALENHOS"
    '
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdTipoReferenciaOrigen int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdTipoReferenciaDestino int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdEstablecimientoOrigen int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdEstablecimientoDestino int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdEstablecimientoNoMinsaOrigen int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdEstablecimientoNoMinsaDestino int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  HuboInfeccionIntraHospitalaria int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  TieneNecropsia bit null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  IdMedicoRespNacimiento int null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  RecienNacido bit null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  NroReferenciaOrigen varchar(20) null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales add  NroReferenciaDestino varchar(20) null"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    lcSql = "select * from atenciones"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lnRegistros = oRsTmpOpc1.RecordCount
    If lnRegistros > 0 Then
       ProgressBar1.Min = 0
       ProgressBar1.Max = lnRegistros + 2
       oRsTmpOpc1.MoveFirst
       Do While Not oRsTmpOpc1.EOF
          DoEvents
          ProgressBar1.Value = ProgressBar1.Value + 1
          Me.Refresh
          If IsNull(oRsTmpOpc1.Fields!IdTipoReferenciaOrigen) And IsNull(oRsTmpOpc1.Fields!IdTipoReferenciaDestino) And _
             IsNull(oRsTmpOpc1.Fields!IdEstablecimientoOrigen) And IsNull(oRsTmpOpc1.Fields!IdEstablecimientoDestino) And _
             IsNull(oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaOrigen) And IsNull(oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaDestino) And _
             IsNull(oRsTmpOpc1.Fields!HuboInfeccionIntraHospitalaria) And IsNull(oRsTmpOpc1.Fields!TieneNecropsia) And _
             IsNull(oRsTmpOpc1.Fields!IdMedicoRespNacimiento) And IsNull(oRsTmpOpc1.Fields!RecienNacido) And _
             IsNull(oRsTmpOpc1.Fields!NroReferenciaOrigen) And IsNull(oRsTmpOpc1.Fields!NroReferenciaDestino) Then
          Else
             lcSql = "select * from AtencionesDatosAdicionales where idAtencion=" & oRsTmpOpc1.Fields!IdAtencion
             If oRsTmpOpc2.State = 1 Then oRsTmpOpc2.Close
             oRsTmpOpc2.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
             If oRsTmpOpc2.RecordCount = 0 Then
                oRsTmpOpc2.AddNew
                oRsTmpOpc2.Fields!IdAtencion = oRsTmpOpc1.Fields!IdAtencion
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdTipoReferenciaOrigen) Then
                oRsTmpOpc2.Fields!IdTipoReferenciaOrigen = oRsTmpOpc1.Fields!IdTipoReferenciaOrigen
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdTipoReferenciaDestino) Then
                oRsTmpOpc2.Fields!IdTipoReferenciaDestino = oRsTmpOpc1.Fields!IdTipoReferenciaDestino
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdEstablecimientoOrigen) Then
                oRsTmpOpc2.Fields!IdEstablecimientoOrigen = oRsTmpOpc1.Fields!IdEstablecimientoOrigen
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdEstablecimientoDestino) Then
                oRsTmpOpc2.Fields!IdEstablecimientoDestino = oRsTmpOpc1.Fields!IdEstablecimientoDestino
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaOrigen) Then
                oRsTmpOpc2.Fields!IdEstablecimientoNoMinsaOrigen = oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaOrigen
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaDestino) Then
                oRsTmpOpc2.Fields!IdEstablecimientoNoMinsaDestino = oRsTmpOpc1.Fields!IdEstablecimientoNoMinsaDestino
             End If
             If Not IsNull(oRsTmpOpc1.Fields!HuboInfeccionIntraHospitalaria) Then
                oRsTmpOpc2.Fields!HuboInfeccionIntraHospitalaria = oRsTmpOpc1.Fields!HuboInfeccionIntraHospitalaria
             End If
             If Not IsNull(oRsTmpOpc1.Fields!TieneNecropsia) Then
                oRsTmpOpc2.Fields!TieneNecropsia = oRsTmpOpc1.Fields!TieneNecropsia
             End If
             If Not IsNull(oRsTmpOpc1.Fields!IdMedicoRespNacimiento) Then
                oRsTmpOpc2.Fields!IdMedicoRespNacimiento = oRsTmpOpc1.Fields!IdMedicoRespNacimiento
             End If
             If Not IsNull(oRsTmpOpc1.Fields!RecienNacido) Then
                oRsTmpOpc2.Fields!RecienNacido = oRsTmpOpc1.Fields!RecienNacido
             End If
             If Not IsNull(oRsTmpOpc1.Fields!NroReferenciaOrigen) Then
                oRsTmpOpc2.Fields!NroReferenciaOrigen = oRsTmpOpc1.Fields!NroReferenciaOrigen
             End If
             If Not IsNull(oRsTmpOpc1.Fields!NroReferenciaDestino) Then
                oRsTmpOpc2.Fields!NroReferenciaDestino = oRsTmpOpc1.Fields!NroReferenciaDestino
             End If
             oRsTmpOpc2.Update
          End If
          oRsTmpOpc1.MoveNext
       Loop
    End If
    oRsTmpOpc1.Close
    '
    lcSql = "ALTER TABLE Atenciones drop column  IdTipoReferenciaOrigen"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdTipoReferenciaDestino"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdEstablecimientoOrigen"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdEstablecimientoDestino"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdEstablecimientoNoMinsaOrigen"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdEstablecimientoNoMinsaDestino"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  HuboInfeccionIntraHospitalaria"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  TieneNecropsia"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  IdMedicoRespNacimiento"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  RecienNacido"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  NroReferenciaOrigen"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE Atenciones drop column  NroReferenciaDestino"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
ErrColuAten:
Resume Next
    oConexODBC.Close
    Set oConexODBC = Nothing
    Set oRsTmpOpc1 = Nothing
    Set oRsTmpOpc2 = Nothing
End Sub


Sub MigraUltimaVersion_TablaSIGH_Parte4(oConexHBT As Connection, oConexODBC As Connection)
    On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean
    
    DoEvents
    ProgressBar1.Value = 148
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePregunta"
    lcSql = "CREATE TABLE [dbo].[AtenIntePregunta] ("
    lcSql = lcSql & "   [IdPregunta] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [Pregunta] [varchar] (70) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [TipoRespuesta] [int] NOT NULL ,"
    lcSql = lcSql & "   [TipoValorRespuesta] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePregunta] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPregunta]"
    lcSql = lcSql & "   )  ON [PRIMARY] "
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenIntePregunta"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenIntePregunta"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPregunta=" & oRsTmpOpc1.Fields!IdPregunta '& "'"
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdPregunta = oRsTmpOpc1.Fields!IdPregunta
           End If
           oRsTmpOpc.Fields!pregunta = oRsTmpOpc1.Fields!pregunta
           oRsTmpOpc.Fields!TipoRespuesta = oRsTmpOpc1.Fields!TipoRespuesta
           oRsTmpOpc.Fields!TipoValorRespuesta = oRsTmpOpc1.Fields!TipoValorRespuesta
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
   
    DoEvents
    ProgressBar1.Value = 149
    Me.Refresh
    txtTablaProceso.Text = "PeriodoTiempo"
    lcSql = "CREATE TABLE [dbo].[PeriodoTiempo] ("
    lcSql = lcSql & "   [IdPeriodoTiempo] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [PeriodoTiempo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_PeriodoTiempo] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPeriodoTiempo]"
    lcSql = lcSql & "   )  ON [PRIMARY] "
    lcSql = lcSql & ") ON [PRIMARY]"
    
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from PeriodoTiempo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PeriodoTiempo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPeriodoTiempo=" & oRsTmpOpc1.Fields!IdPeriodoTiempo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdPeriodoTiempo = oRsTmpOpc1.Fields!IdPeriodoTiempo
           End If
           oRsTmpOpc.Fields!PeriodoTiempo = oRsTmpOpc1.Fields!PeriodoTiempo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 150
    Me.Refresh
    txtTablaProceso.Text = "AtenInteGrupo"
    lcSql = "CREATE TABLE [dbo].[AtenInteGrupo] ("
    lcSql = lcSql & "   [IdAtenInteGrupo] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [AtencionIntegralGrupo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [DesdeAnio] [int] NOT NULL ,"
    lcSql = lcSql & "   [DesdeMes] [int] NOT NULL ,"
    lcSql = lcSql & "   [DesdeDia] [int] NOT NULL ,"
    lcSql = lcSql & "   [HastaAnio] [int] NOT NULL ,"
    lcSql = lcSql & "   [HastaMes] [int] NOT NULL ,"
    lcSql = lcSql & "   [HastaDia] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtencionIntegralGrupos] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   )  ON [PRIMARY] "
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteGrupo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteGrupo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdAtenInteGrupo=" & oRsTmpOpc1.Fields!IdAtenInteGrupo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdAtenInteGrupo = oRsTmpOpc1.Fields!IdAtenInteGrupo
           End If
           oRsTmpOpc.Fields!AtencionIntegralGrupo = oRsTmpOpc1.Fields!AtencionIntegralGrupo
           oRsTmpOpc.Fields!DesdeAnio = oRsTmpOpc1.Fields!DesdeAnio
           oRsTmpOpc.Fields!DesdeMes = oRsTmpOpc1.Fields!DesdeMes
           oRsTmpOpc.Fields!DesdeDia = oRsTmpOpc1.Fields!DesdeDia
           oRsTmpOpc.Fields!HastaAnio = oRsTmpOpc1.Fields!HastaAnio
           oRsTmpOpc.Fields!HastaMes = oRsTmpOpc1.Fields!HastaMes
           oRsTmpOpc.Fields!HastaDia = oRsTmpOpc1.Fields!HastaDia
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 151
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemDesarrollo"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemDesarrollo] ("
    lcSql = lcSql & "   [IdItemDesarrollo] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [ItemDesarrollo] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_PerinatalItemDesarrollo] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   )  ON [PRIMARY] "
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemDesarrollo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemDesarrollo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdItemDesarrollo=" & oRsTmpOpc1.Fields!IdItemDesarrollo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdItemDesarrollo = oRsTmpOpc1.Fields!IdItemDesarrollo
           End If
           oRsTmpOpc.Fields!ItemDesarrollo = oRsTmpOpc1.Fields!ItemDesarrollo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 152
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemPlan"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [ItemPlan] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteItemPlan] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   )  ON [PRIMARY] "
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemPlan"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemPlan"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdAtenInteItemPlan=" & oRsTmpOpc1.Fields!IdAtenInteItemPlan
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdAtenInteItemPlan = oRsTmpOpc1.Fields!IdAtenInteItemPlan
           End If
           oRsTmpOpc.Fields!ItemPlan = oRsTmpOpc1.Fields!ItemPlan
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 153
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanAtencion"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "   [IdPlanAtencion] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteGrupo] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPeriodoTiempo] [int] NOT NULL ,"
    lcSql = lcSql & "   [EdadAnio] [int] NOT NULL ,"
    lcSql = lcSql & "   [EdadMes] [int] NOT NULL ,"
    lcSql = lcSql & "   [EdadDia] [int] NOT NULL ,"
    lcSql = lcSql & "   [Descripcion] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanAtencion] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanAtencion_AtenInteGrupo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteGrupo] ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanAtencion_PeriodoTiempo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPeriodoTiempo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[PeriodoTiempo] ("
    lcSql = lcSql & "       [IdPeriodoTiempo]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenIntePlanAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenIntePlanAtencion"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPlanAtencion=" & oRsTmpOpc1.Fields!IdPlanAtencion
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdPlanAtencion = oRsTmpOpc1.Fields!IdPlanAtencion
           End If
           oRsTmpOpc.Fields!IdAtenInteGrupo = oRsTmpOpc1.Fields!IdAtenInteGrupo
           oRsTmpOpc.Fields!IdPeriodoTiempo = oRsTmpOpc1.Fields!IdPeriodoTiempo
           oRsTmpOpc.Fields!EdadAnio = oRsTmpOpc1.Fields!EdadAnio
           oRsTmpOpc.Fields!EdadMes = oRsTmpOpc1.Fields!EdadMes
           oRsTmpOpc.Fields!EdadDia = oRsTmpOpc1.Fields!EdadDia
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 154
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemPlanCrecimiento"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemPlanCrecimiento] ("
    lcSql = lcSql & "   [IdItemPlanCrecimiento] [bigint] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [NumeroSesion] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteItemPlanCrecimiento] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanCrecimiento]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanCrecimiento_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemPlanCrecimiento"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemPlanCrecimiento"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdItemPlanCrecimiento=" & oRsTmpOpc1.Fields!IdItemPlanCrecimiento
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdItemPlanCrecimiento = oRsTmpOpc1.Fields!IdItemPlanCrecimiento
           End If
           oRsTmpOpc.Fields!IdPlanAtencion = oRsTmpOpc1.Fields!IdPlanAtencion
           oRsTmpOpc.Fields!NumeroSesion = oRsTmpOpc1.Fields!NumeroSesion
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 155
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanCrecDetalle"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanCrecDetalle] ("
    lcSql = lcSql & "   [IdItemPlanCrecimiento] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdTriajeVariable] [int] NOT NULL ,"
    lcSql = lcSql & "   [OrdenItem] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanCrecDetalle] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanCrecimiento],"
    lcSql = lcSql & "       [IdTriajeVariable]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecDetalle_AtenInteItemPlanCrecimiento] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanCrecimiento]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlanCrecimiento] ("
    lcSql = lcSql & "       [IdItemPlanCrecimiento]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenIntePlanCrecDetalle"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenIntePlanCrecDetalle"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
'              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Filter = "IdItemPlanCrecimiento=" & oRsTmpOpc1.Fields!IdItemPlanCrecimiento & _
                                " and IdTriajeVariable=" & oRsTmpOpc1.Fields!IdTriajeVariable
              If Not (oRsTmpOpc.EOF = True And oRsTmpOpc.BOF = True) Then
                 lbNuevoRegistro = False
                 oRsTmpOpc.MoveFirst
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdItemPlanCrecimiento = oRsTmpOpc1.Fields!IdItemPlanCrecimiento
                oRsTmpOpc.Fields!IdTriajeVariable = oRsTmpOpc1.Fields!IdTriajeVariable
           End If
           oRsTmpOpc.Fields!OrdenItem = oRsTmpOpc1.Fields!OrdenItem
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
   
    DoEvents
    ProgressBar1.Value = 156
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemPlanDesarrollo"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemPlanDesarrollo] ("
    lcSql = lcSql & "   [IdItemPlanDesarrollo] [bigint] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [NumeroSesion] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenItemPlanDesarrollo] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanDesarrollo]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenItemPlanDesarrollo_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemPlanDesarrollo"
    Set oRsTmpOpc = Nothing
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemPlanDesarrollo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdItemPlanDesarrollo=" & oRsTmpOpc1.Fields!IdItemPlanDesarrollo
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdItemPlanDesarrollo = oRsTmpOpc1.Fields!IdItemPlanDesarrollo
           End If
           oRsTmpOpc.Fields!IdPlanAtencion = oRsTmpOpc1.Fields!IdPlanAtencion
           oRsTmpOpc.Fields!NumeroSesion = oRsTmpOpc1.Fields!NumeroSesion
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 157
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanDesDetalle"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanDesDetalle] ("
    lcSql = lcSql & "   [IdItemPlanDesarrollo] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdItemDesarrollo] [int] NOT NULL ,"
    lcSql = lcSql & "   [OrdenItem] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanDesDetalle] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanDesarrollo],"
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesDetalle_AtenInteItemDesarrollo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemDesarrollo] ("
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesDetalle_AtenInteItemPlanDesarrollo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanDesarrollo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlanDesarrollo] ("
    lcSql = lcSql & "       [IdItemPlanDesarrollo]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenIntePlanDesDetalle"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenIntePlanDesDetalle"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
'              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Filter = "IdItemPlanDesarrollo=" & oRsTmpOpc1.Fields!IdItemPlanDesarrollo & _
                                " and IdItemDesarrollo=" & oRsTmpOpc1.Fields!IdItemDesarrollo
              If Not (oRsTmpOpc.EOF = True And oRsTmpOpc.BOF = True) Then
                 lbNuevoRegistro = False
                 oRsTmpOpc.MoveFirst
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdItemPlanDesarrollo = oRsTmpOpc1.Fields!IdItemPlanDesarrollo
                oRsTmpOpc.Fields!IdItemDesarrollo = oRsTmpOpc1.Fields!IdItemDesarrollo
           End If
           oRsTmpOpc.Fields!OrdenItem = oRsTmpOpc1.Fields!OrdenItem
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    
    DoEvents
    ProgressBar1.Value = 158
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemPlanProcedimiento"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemPlanProcedimiento] ("
    lcSql = lcSql & "   [IdItemPlanProcedimiento] [bigint] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdProducto] [int] NOT NULL ,"
    lcSql = lcSql & "   [NumeroDosis] [tinyint] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteItemPlanInmunizacion] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemPlanProcedimiento]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanInmunizacion_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanInmunizacion_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanInmunizacion_FactCatalogoServicios] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactCatalogoServicios] ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE AtenInteItemPlanProcedimiento ADD IdDiagnostico int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemPlanProcedimiento"
    Set oRsTmpOpc = Nothing
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemPlanProcedimiento"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdItemPlanProcedimiento=" & oRsTmpOpc1.Fields!IdItemPlanProcedimiento
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdItemPlanProcedimiento = oRsTmpOpc1.Fields!IdItemPlanProcedimiento
           End If
           oRsTmpOpc.Fields!IdPlanAtencion = oRsTmpOpc1.Fields!IdPlanAtencion
           oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           oRsTmpOpc.Fields!NumeroDosis = oRsTmpOpc1.Fields!NumeroDosis
           oRsTmpOpc.Fields!IdAtenInteItemPlan = oRsTmpOpc1.Fields!IdAtenInteItemPlan
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 159
    Me.Refresh
    txtTablaProceso.Text = "AtenInteItemPlanSuplemento"
    lcSql = "CREATE TABLE [dbo].[AtenInteItemPlanSuplemento] ("
    lcSql = lcSql & "   [ItemPlanSuplemento] [bigint] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdProducto] [int] NOT NULL ,"
    lcSql = lcSql & "   [NumeroDosis] [int] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteItemPlanSuplemento] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [ItemPlanSuplemento]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanSuplemento_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteItemPlanSuplemento_FactCatalogoBienesInsumos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactCatalogoBienesInsumos] ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE AtenInteItemPlanSuplemento ADD IdDiagnostico int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteItemPlanSuplemento"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteItemPlanSuplemento"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "ItemPlanSuplemento=" & oRsTmpOpc1.Fields!ItemPlanSuplemento
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
'                oRsTmpOpc.Fields!ItemPlanSuplemento = oRsTmpOpc1.Fields!ItemPlanSuplemento
           End If
           oRsTmpOpc.Fields!IdPlanAtencion = oRsTmpOpc1.Fields!IdPlanAtencion
           oRsTmpOpc.Fields!IdProducto = oRsTmpOpc1.Fields!IdProducto
           oRsTmpOpc.Fields!NumeroDosis = oRsTmpOpc1.Fields!NumeroDosis
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    
    DoEvents
    ProgressBar1.Value = 160
    Me.Refresh
    txtTablaProceso.Text = "AtenInteGrupoHCPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenInteGrupoHCPaciente] ("
    lcSql = lcSql & "   [IdPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdGrupoHCPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteGrupo] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPregunta] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteGrupoHCPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdGrupoHCPaciente],"
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteGrupoHCPaciente_AtenInteGrupo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteGrupo] ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteGrupoHCPaciente_AtenIntePregunta] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPregunta]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePregunta] ("
    lcSql = lcSql & "       [IdPregunta]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteGrupoHCPaciente_Pacientes] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Pacientes] ("
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 161
    Me.Refresh
    txtTablaProceso.Text = "AtenInteGrupoHCRespuestaPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenInteGrupoHCRespuestaPaciente] ("
    lcSql = lcSql & "   [IdGrupoHCPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [ItemRespuesta] [int] NOT NULL ,"
    lcSql = lcSql & "   [ValorTexto] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   [ValorNumero] [money] NULL ,"
    lcSql = lcSql & "   [ValorFecha] [datetime] NULL ,"
    lcSql = lcSql & "   [ValorNumeroFin] [money] NULL ,"
    lcSql = lcSql & "   [ValorFechaFin] [datetime] NULL ,"
    lcSql = lcSql & "   [ValorEspecificacion] [varchar] (1500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteGrupoHCRespuestaPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdGrupoHCPaciente],"
    lcSql = lcSql & "       [IdPaciente],"
    lcSql = lcSql & "       [ItemRespuesta]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    'lcSql = lcSql & "   CONSTRAINT [DF_AtenInteGrupoHCRespuestaPaciente_EsActivo] DEFAULT (1) FOR [EsActivo],"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenInteGrupoHCRespuestaPaciente_AtenInteGrupoHCPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdGrupoHCPaciente],"
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteGrupoHCPaciente] ("
    lcSql = lcSql & "       [IdGrupoHCPaciente],"
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[AtenInteGrupoHCRespuestaPaciente] WITH NOCHECK ADD"
    lcSql = lcSql & " CONSTRAINT [DF_AtenInteGrupoHCRespuestaPaciente_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    
    DoEvents
    ProgressBar1.Value = 162
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanCrecimientoPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanCrecimientoPaciente] ("
    lcSql = lcSql & "   [IdPlanCrecimientoPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   [FechaProgramada] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEjecucion] [datetime] NULL ,"
    lcSql = lcSql & "   [NumeroSesion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtencion] [int] NULL ,"
    lcSql = lcSql & "   [IdEstablecimiento] [int] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanCrecimientoPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanCrecimientoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecimientoPaciente_Atenciones] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Atenciones] ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecimientoPaciente_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecimientoPaciente_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecimientoPaciente_Establecimientos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Establecimientos] ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 163
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanCrecPacienteDet"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanCrecPacienteDet] ("
    lcSql = lcSql & "   [IdPlanCrecimientoPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdTriajeVariable] [int] NOT NULL ,"
    lcSql = lcSql & "   [VariableValor] [money] NOT NULL ,"
    lcSql = lcSql & "   [OrdenItem] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanCrecPacienteDet] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanCrecimientoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente],"
    lcSql = lcSql & "       [IdTriajeVariable]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanCrecPacienteDet_AtenIntePlanCrecimientoPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanCrecimientoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanCrecimientoPaciente] ("
    lcSql = lcSql & "       [IdPlanCrecimientoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 164
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlantillaItemPlan"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlantillaItemPlan] ("
    lcSql = lcSql & "   [IdPlantillaItemPlan] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteGrupo] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlantillaItemPlan] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlantillaItemPlan]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlantillaItemPlan_AtenInteGrupo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteGrupo] ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlantillaItemPlan_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenIntePlantillaItemPlan"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenIntePlantillaItemPlan"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdPlantillaItemPlan=" & oRsTmpOpc1.Fields!IdPlantillaItemPlan
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdPlantillaItemPlan = oRsTmpOpc1.Fields!IdPlantillaItemPlan
           End If
           oRsTmpOpc.Fields!IdAtenInteGrupo = oRsTmpOpc1.Fields!IdAtenInteGrupo
           oRsTmpOpc.Fields!IdAtenInteItemPlan = oRsTmpOpc1.Fields!IdAtenInteItemPlan
           
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    
    DoEvents
    ProgressBar1.Value = 165
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanIntegralPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanIntegralPaciente] ("
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteGrupo] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [FechaElaboracion] [datetime] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanIntegralPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanIntegralPaciente_AtenInteGrupo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteGrupo] ("
    lcSql = lcSql & "       [IdAtenInteGrupo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanIntegralPaciente_Pacientes] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Pacientes] ("
    lcSql = lcSql & "       [IdPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    
    DoEvents
    ProgressBar1.Value = 166
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanDesarrolloPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanDesarrolloPaciente] ("
    lcSql = lcSql & "   [IdPlanDesarrolloPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [Evaluacion] [int] NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   [FechaProgramada] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEjecucion] [datetime] NULL ,"
    lcSql = lcSql & "   [NumeroSesion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtencion] [int] NULL ,"
    lcSql = lcSql & "   [IdEstablecimiento] [int] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanDesarrolloPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanDesarrolloPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesarrolloPaciente_Atenciones] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Atenciones] ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesarrolloPaciente_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesarrolloPaciente_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesarrolloPaciente_AtenIntePlanIntegralPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanIntegralPaciente] ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesarrolloPaciente_Establecimientos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Establecimientos] ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 167
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanDesPacienteDet"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanDesPacienteDet] ("
    lcSql = lcSql & "   [IdPlanDesarrolloPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdItemDesarrollo] [int] NOT NULL ,"
    lcSql = lcSql & "   [OrdenItem] [int] NOT NULL ,"
    lcSql = lcSql & "   [EjecutaAccion] [bit] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanDesPacienteDet] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanDesarrolloPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente],"
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesPacienteDet_AtenInteItemDesarrollo] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemDesarrollo] ("
    lcSql = lcSql & "       [IdItemDesarrollo]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanDesPacienteDet_AtenIntePlanDesarrolloPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanDesarrolloPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanDesarrolloPaciente] ("
    lcSql = lcSql & "       [IdPlanDesarrolloPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 168
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanItemElaborado"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanItemElaborado] ("
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [EsElaborado] [bit] NOT NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanItemElaborado] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanItemElaborado_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanItemElaborado_AtenIntePlanIntegralPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanIntegralPaciente] ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    
    DoEvents
    ProgressBar1.Value = 169
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanProcedimientoPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanProcedimientoPaciente] ("
    lcSql = lcSql & "   [IdPlanProcedimientoPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdProducto] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   [FechaProgramada] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEjecucion] [datetime] NULL ,"
    lcSql = lcSql & "   [NumeroDosis] [tinyint] NOT NULL ,"
    lcSql = lcSql & "   [CodigoHIS] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   [IdAtencion] [int] NULL ,"
    lcSql = lcSql & "   [IdEstablecimiento] [int] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanInmunizacionPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanProcedimientoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanInmunizacionPaciente_Atenciones] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Atenciones] ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanInmunizacionPaciente_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanInmunizacionPaciente_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanInmunizacionPaciente_AtenIntePlanIntegralPaciente] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanIntegralPaciente] ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanInmunizacionPaciente_FactCatalogoServicios] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactCatalogoServicios] ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanProcedimientoPaciente_Establecimientos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Establecimientos] ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE AtenIntePlanProcedimientoPaciente ADD IdDiagnostico int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 170
    Me.Refresh
    txtTablaProceso.Text = "AtenIntePlanSuplementoPaciente"
    lcSql = "CREATE TABLE [dbo].[AtenIntePlanSuplementoPaciente] ("
    lcSql = lcSql & "   [IdPlanSuplementoPaciente] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanIntegralPaciente] [bigint] NOT NULL ,"
    lcSql = lcSql & "   [IdPlanAtencion] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdProducto] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtenInteItemPlan] [int] NOT NULL ,"
    lcSql = lcSql & "   [FechaProgramada] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEjecucion] [datetime] NULL ,"
    lcSql = lcSql & "   [NumeroDosis] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdAtencion] [int] NULL ,"
    lcSql = lcSql & "   [IdEstablecimiento] [int] NULL ,"
    lcSql = lcSql & "   CONSTRAINT [PK_AtenIntePlanSuplementoPaciente] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanSuplementoPaciente],"
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   )  ON [PRIMARY] ,"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_Atenciones] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Atenciones] ("
    lcSql = lcSql & "       [IdAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_AtenInteItemPlan] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenInteItemPlan] ("
    lcSql = lcSql & "       [IdAtenInteItemPlan]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_AtenIntePlanAtencion] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanAtencion] ("
    lcSql = lcSql & "       [IdPlanAtencion]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_AtenIntePlanIntegralPaciente1] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[AtenIntePlanIntegralPaciente] ("
    lcSql = lcSql & "       [IdPlanIntegralPaciente]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_Establecimientos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[Establecimientos] ("
    lcSql = lcSql & "       [IdEstablecimiento]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_AtenIntePlanSuplementoPaciente_FactCatalogoBienesInsumos] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactCatalogoBienesInsumos] ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE AtenIntePlanSuplementoPaciente ADD IdDiagnostico int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 171
    Me.Refresh
    txtTablaProceso.Text = "AtenInteValorTalla"
    lcSql = "CREATE TABLE [dbo].[AtenInteValorTalla] ("
    lcSql = lcSql & "   [IdValorTalla] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdTipoSexo] [int] NOT NULL ,"
    lcSql = lcSql & "   [EdadMeses] [int] NOT NULL ,"
    lcSql = lcSql & "   [NroDesviacion] [int] NOT NULL , "
    lcSql = lcSql & "   [ValorTalla] [money] NOT NULL"
    lcSql = lcSql & "   ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[AtenInteValorTalla] WITH NOCHECK ADD "
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteValorTalla] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "      [IdValorTalla]"
    lcSql = lcSql & "   )  ON [PRIMARY]  "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteValorTalla"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteValorTalla"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdValorTalla=" & oRsTmpOpc1.Fields!IdValorTalla
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                'oRsTmpOpc.Fields!IdValorTalla = oRsTmpOpc1.Fields!IdValorTalla
           End If
           oRsTmpOpc.Fields!idTipoSexo = oRsTmpOpc1.Fields!idTipoSexo
           oRsTmpOpc.Fields!EdadMeses = oRsTmpOpc1.Fields!EdadMeses
           oRsTmpOpc.Fields!NroDesviacion = oRsTmpOpc1.Fields!NroDesviacion
           oRsTmpOpc.Fields!ValorTalla = oRsTmpOpc1.Fields!ValorTalla
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close

    
    DoEvents
    ProgressBar1.Value = 172
    Me.Refresh
    txtTablaProceso.Text = "AtenInteValorPeso"
    lcSql = "CREATE TABLE [dbo].[AtenInteValorPeso] ("
    lcSql = lcSql & "   [IdValorPeso] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdTipoSexo] [int] NOT NULL ,"
    lcSql = lcSql & "   [EdadMeses] [int] NOT NULL ,"
    lcSql = lcSql & "   [NroDesviacion] [int] NOT NULL , "
    lcSql = lcSql & "   [ValorPeso] [money] NOT NULL"
    lcSql = lcSql & "   ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[AtenInteValorPeso] WITH NOCHECK ADD "
    lcSql = lcSql & "   CONSTRAINT [PK_AtenInteValorPeso] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "      [IdValorPeso]"
    lcSql = lcSql & "   )  ON [PRIMARY]  "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from AtenInteValorPeso"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from AtenInteValorPeso"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdValorPeso=" & oRsTmpOpc1.Fields!IdValorPeso
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
           End If
           oRsTmpOpc.Fields!idTipoSexo = oRsTmpOpc1.Fields!idTipoSexo
           oRsTmpOpc.Fields!EdadMeses = oRsTmpOpc1.Fields!EdadMeses
           oRsTmpOpc.Fields!NroDesviacion = oRsTmpOpc1.Fields!NroDesviacion
           oRsTmpOpc.Fields!ValorPeso = oRsTmpOpc1.Fields!ValorPeso
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    '''''''''''
    
    
    
    Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Sub MigraUltimaVersion_TablaSIGH_Parte5(oConexHBT As Connection, oConexODBC As Connection)
On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean

        
    DoEvents
    ProgressBar1.Value = 173
    Me.Refresh
    txtTablaProceso.Text = "ImagCatalgoServicioDuracion"
    lcSql = "CREATE TABLE [dbo].[ImagCatalgoServicioDuracion] ("
    lcSql = lcSql & "   [IdProducto] [int] NOT NULL ,"
    lcSql = lcSql & "   [DuracionEnMin] [money] NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_ImagCatalgoServicioDuracion] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   )  ON [PRIMARY],"
    lcSql = lcSql & "   CONSTRAINT [FK_ImagCatalgoServicioDuracion_FactCatalogoServicios] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactCatalogoServicios] ("
    lcSql = lcSql & "       [IdProducto]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    lcSql = "ALTER TABLE [dbo].[ImagCatalgoServicioDuracion] WITH NOCHECK ADD "
    lcSql = lcSql & " CONSTRAINT [DF_ImagCatalgoServicioDuracion_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
       
    DoEvents
    ProgressBar1.Value = 174
    Me.Refresh
    txtTablaProceso.Text = "ImagTipoModalidadSala"
    lcSql = "CREATE TABLE [dbo].[ImagTipoModalidadSala] ("
    lcSql = lcSql & "   [IdTipoModalidadSala] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [TipoModalidadSala] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [Codigo] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_ImagTipoModalidadSala] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdTipoModalidadSala]"
    lcSql = lcSql & "   )  ON [PRIMARY]"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[ImagTipoModalidadSala] WITH NOCHECK ADD  "
    lcSql = lcSql & " CONSTRAINT [DF_ImagTipoModalidadSala_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 175
    Me.Refresh
    txtTablaProceso.Text = "ImagSala"
    lcSql = "CREATE TABLE [dbo].[ImagSala] ("
    lcSql = lcSql & "   [IdSala] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdTipoModalidadSala] [int] NOT NULL ,"
    lcSql = lcSql & "   [Sala] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [Codigo] [varchar] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_ImagSala] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdSala]"
    lcSql = lcSql & "   )  ON [PRIMARY],"
    lcSql = lcSql & "   CONSTRAINT [FK_ImagSala_ImagTipoModalidadSala] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdTipoModalidadSala]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[ImagTipoModalidadSala] ("
    lcSql = lcSql & "       [IdTipoModalidadSala]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[ImagSala] WITH NOCHECK ADD  "
    lcSql = lcSql & " CONSTRAINT [DF_ImagSala_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    
       
    DoEvents
    ProgressBar1.Value = 176
    Me.Refresh
    txtTablaProceso.Text = "ImagSalaPuntoCarga"
    lcSql = "CREATE TABLE [dbo].[ImagSalaPuntoCarga] ("
    lcSql = lcSql & "   [IdSala] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdPuntoCarga] [int] NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechsCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_ImagSalaPuntoCarga] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "       [IdSala],"
    lcSql = lcSql & "       [IdPuntoCarga]"
    lcSql = lcSql & "   )  ON [PRIMARY],"
    lcSql = lcSql & "   CONSTRAINT [FK_ImagSalaPuntoCarga_FactPuntosCarga] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdPuntoCarga]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[FactPuntosCarga] ("
    lcSql = lcSql & "       [IdPuntoCarga]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_ImagSalaPuntoCarga_ImagSala] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdSala]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[ImagSala] ("
    lcSql = lcSql & "       [IdSala]"
    lcSql = lcSql & "   )"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    

    DoEvents
    ProgressBar1.Value = 177
    Me.Refresh
    txtTablaProceso.Text = "InteoTipoSistema"
    lcSql = "CREATE TABLE [dbo].[InteoTipoSistema] ("
    lcSql = lcSql & "   [IdTipoSistema] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [TipoSistema] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_InteoTipoSistema] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdTipoSistema]"
    lcSql = lcSql & "   )  ON [PRIMARY],"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[InteoTipoSistema] WITH NOCHECK ADD   "
    lcSql = lcSql & " CONSTRAINT [DF_InteoTipoSistema_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from InteoTipoSistema"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from InteoTipoSistema"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdTipoSistema=" & oRsTmpOpc1.Fields!IdTipoSistema
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!FechaCrea = oRsTmpOpc1.Fields!FechaCrea
                oRsTmpOpc.Fields!FechaEdita = oRsTmpOpc1.Fields!FechaEdita
           End If
           oRsTmpOpc.Fields!TipoSistema = oRsTmpOpc1.Fields!TipoSistema
           oRsTmpOpc.Fields!EsActivo = oRsTmpOpc1.Fields!EsActivo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
        
        
    DoEvents
    ProgressBar1.Value = 178
    Me.Refresh
    txtTablaProceso.Text = "InteoProveedorSistema"
    lcSql = "CREATE TABLE [dbo].[InteoProveedorSistema] ("
    lcSql = lcSql & "   [IdProveedorSistema] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [ProveedorSistema] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_IOpProveedorSistema] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProveedorSistema]"
    lcSql = lcSql & "   )  ON [PRIMARY]"
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from InteoProveedorSistema"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from InteoProveedorSistema"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              oRsTmpOpc.MoveFirst
              oRsTmpOpc.Find "IdProveedorSistema=" & oRsTmpOpc1.Fields!IdProveedorSistema
              If Not oRsTmpOpc.EOF Then
                 lbNuevoRegistro = False
              End If
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!FechaCrea = oRsTmpOpc1.Fields!FechaCrea
                oRsTmpOpc.Fields!FechaEdita = oRsTmpOpc1.Fields!FechaEdita
           End If
           oRsTmpOpc.Fields!ProveedorSistema = oRsTmpOpc1.Fields!ProveedorSistema
           oRsTmpOpc.Fields!EsActivo = oRsTmpOpc1.Fields!EsActivo
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
        
    DoEvents
    ProgressBar1.Value = 179
    Me.Refresh
    txtTablaProceso.Text = "InteoIntegracionSistema"
    lcSql = "CREATE TABLE [dbo].[InteoIntegracionSistema] ("
    lcSql = lcSql & "   [IdIntegracionSistema] [int] IDENTITY (1, 1) NOT NULL ,"
    lcSql = lcSql & "   [IdTipoSistema] [int] NOT NULL ,"
    lcSql = lcSql & "   [IdProveedorSistema] [int] NOT NULL ,"
    lcSql = lcSql & "   [NombreUsuario] [varchar] (35) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   [ClaveUsuario] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,"
    lcSql = lcSql & "   [EsProveedorActual] [bit] NOT NULL ,"
    lcSql = lcSql & "   [EsActivo] [bit] NOT NULL ,"
    lcSql = lcSql & "   [FechaCrea] [datetime] NOT NULL ,"
    lcSql = lcSql & "   [FechaEdita] [datetime] NULL,"
    lcSql = lcSql & "   CONSTRAINT [PK_InteoIntegracionSistema] PRIMARY KEY  CLUSTERED "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdIntegracionSistema]"
    lcSql = lcSql & "   )  ON [PRIMARY],"
    lcSql = lcSql & "   CONSTRAINT [FK_InteoIntegracionSistema_InteoProveedorSistema] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdProveedorSistema]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[InteoProveedorSistema] ("
    lcSql = lcSql & "       [IdProveedorSistema]"
    lcSql = lcSql & "   ),"
    lcSql = lcSql & "   CONSTRAINT [FK_InteoIntegracionSistema_InteoTipoSistema] FOREIGN KEY "
    lcSql = lcSql & "   ("
    lcSql = lcSql & "       [IdTipoSistema]"
    lcSql = lcSql & "   ) REFERENCES [dbo].[InteoTipoSistema] ("
    lcSql = lcSql & "       [IdTipoSistema]"
    lcSql = lcSql & "   ) "
    lcSql = lcSql & ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[InteoIntegracionSistema] WITH NOCHECK ADD   "
    lcSql = lcSql & " CONSTRAINT [DF_InteoIntegracionSistema_EsProveedorActual] DEFAULT (0) FOR [EsProveedorActual],"
    lcSql = lcSql & " CONSTRAINT [DF_InteoIntegracionSistema_EsActivo] DEFAULT (1) FOR [EsActivo]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Sub MigraUltimaVersion_TablaSIGH_Parte6(oConexHBT As Connection, oConexODBC As Connection)
On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean

    DoEvents
    ProgressBar1.Value = 180
    Me.Refresh
    txtTablaProceso.Text = "HIS_DxOmitidos"
    lcSql = " CREATE TABLE [dbo].[HIS_DxOmitidos] (" & _
    "[Lote] [varchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[NroPagina] [int] NOT NULL ," & _
    "[Codigo1] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Codigo2] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Codigo3] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Codigo4] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Codigo5] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Codigo6] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[LabConf1] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[LabConf2] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[LabConf3] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[LabConf4] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[LabConf5] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[LabConf6] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[Diagnost1] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Diagnost2] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Diagnost3] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Diagnost4] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Diagnost5] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Diagnost6] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Edad] [int] NOT NULL ,[TIP_EDAD] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[Sexo] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Peso] [money] NOT NULL ,[FichaFam] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[Establecimiento] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,[Servicio] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[CODVALIDACION] [int] NULL) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table HIS_DxOmitidos drop column Lote"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table HIS_DxOmitidos add IdLote int NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table HIS_DxOmitidos add DiaAtencion int NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 181
    Me.Refresh
    txtTablaProceso.Text = "HIS_TemporalSexo"
    lcSql = " CREATE TABLE [dbo].[HIS_TemporalSexo] (" & _
    "[CieCpt] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
    "[DxSexo] [int] NULL ," & _
    "[EsCpt] [int] NULL" & _
    ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from HIS_TemporalSexo"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_TemporalSexo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from HIS_TemporalSexo where CieCpt='" & _
                                         oRsTmpOpc1.Fields!CieCpt & "'"
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!CieCpt = oRsTmpOpc1.Fields!CieCpt
           End If
           oRsTmpOpc.Fields!DxSexo = oRsTmpOpc1.Fields!DxSexo
           oRsTmpOpc.Fields!EsCpt = oRsTmpOpc1.Fields!EsCpt
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
        
    DoEvents
    ProgressBar1.Value = 182
    Me.Refresh
    txtTablaProceso.Text = "HIS_Financiador"
    lcSql = " CREATE TABLE [dbo].[HIS_Financiador] (" & _
        "[IdCodigoFinancHis] [int] NOT NULL , " & _
        "[DescripcionFinancHis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , " & _
        "[IdFuenteFinanciamiento] [int] NULL " & _
        ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = " ALTER TABLE [dbo].[HIS_Financiador] WITH NOCHECK ADD" & _
            " CONSTRAINT [PK_HIS_Financiador] PRIMARY KEY  CLUSTERED" & _
            " ( [IdCodigoFinancHis] " & _
            " )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from HIS_Financiador"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from HIS_Financiador"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from HIS_Financiador where IdCodigoFinancHis=" & _
                                         oRsTmpOpc1.Fields!IdCodigoFinancHis
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdCodigoFinancHis = oRsTmpOpc1.Fields!IdCodigoFinancHis
           End If
           oRsTmpOpc.Fields!DescripcionFinancHis = oRsTmpOpc1.Fields!DescripcionFinancHis
           oRsTmpOpc.Fields!idFuenteFinanciamiento = oRsTmpOpc1.Fields!idFuenteFinanciamiento
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    
    DoEvents
    ProgressBar1.Value = 183
    Me.Refresh
    txtTablaProceso.Text = "His_EstadosCalidad"
    lcSql = "CREATE TABLE [dbo].[His_EstadosCalidad] (" & _
            "    [IdEstado] [int] NOT NULL ," & _
            "    [Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[His_EstadosCalidad] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_His_EstadosCalidad] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idEstado]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_EstadosCalidad"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from His_EstadosCalidad"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from His_EstadosCalidad where IdEstado=" & _
                                         oRsTmpOpc1.Fields!idEstado
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstado = oRsTmpOpc1.Fields!idEstado
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    DoEvents
    ProgressBar1.Value = 184
    Me.Refresh
    txtTablaProceso.Text = "PadronNominal_Detalle"
  
     lcSql = "CREATE TABLE [dbo].[PadronNominal_Detalle] (" & _
     "[IdPaNomDetalle] [int] IDENTITY (1, 1) NOT FOR REPLICATION  NOT NULL ," & _
     "[IdTipoDoc] [int] NULL ," & _
     "[NumDocumento] [int] NULL ," & _
     "[HistClinica] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[ApellidoPaterno] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[ApellidoMaterno] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[Nombres] [varchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[idSexo] [int] NOT NULL ," & _
     "[FecNacimiento] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[IdTipoSeguro] [int] NULL ," & _
     "[NumAfiliacion] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
     "[FecEvaluacion] [varchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[Peso] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[Talla] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
     "[IdDiagNutricional] [int] NULL ," & _
     "[CodRenaes] [int] NULL ," & _
     "[IdDiagPE] [int] NOT NULL ," & _
     "[IdDiagPT] [int] NOT NULL ," & _
     "[IdDiagTE] [int] NOT NULL ," & _
     ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '*****debb-21/04/2015 (inicio)
    lcSql = "ALTER TABLE PadronNominal_Detalle add  hemoglobina int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PadronNominal_Detalle add  heces varchar(2)  COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "alter table PadronNominal_Detalle alter column heces varchar(2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE PadronNominal_Detalle add  renaes varchar(10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '*****debb-21/04/2015 (fin)
    
    DoEvents
    ProgressBar1.Value = 185
    Me.Refresh
    txtTablaProceso.Text = "PadronNominal_DxNutriTemp"
    
    lcSql = " CREATE TABLE [dbo].[PadronNominal_DxNutriTemp] (" & _
    "[IdpadNomFormulario] [int] NOT NULL ," & _
    "[Descripcion] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL " & _
    ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from PadronNominal_DxNutriTemp" 'aCCES
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PadronNominal_DxNutriTemp" ' SQL
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from PadronNominal_DxNutriTemp where IdpadNomFormulario=" & _
                                         oRsTmpOpc1.Fields!IdpadNomFormulario
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdpadNomFormulario = oRsTmpOpc1.Fields!IdpadNomFormulario
           End If
           oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 186
    Me.Refresh
    txtTablaProceso.Text = "padronnominal_dxnutricional"
    lcSql = " CREATE TABLE [dbo].[padronnominal_dxnutricional] (" & _
    "[iddiagnostico] [int] NULL ," & _
    "[LabReferencia] [varchar] (8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[RangoInicial] [money] NULL ," & _
    "[RangoFinal] [money] NULL " & _
    ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from padronnominal_dxnutricional" 'aCCES
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from padronnominal_dxnutricional" ' SQL
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from padronnominal_dxnutricional where iddiagnostico=" & _
                                         oRsTmpOpc1.Fields!IdDiagnostico & _
                                         " and  LabReferencia='" & oRsTmpOpc1.Fields!LabReferencia & "'" & _
                                         " and RangoInicial=" & oRsTmpOpc1.Fields!RangoInicial & _
                                         " and RangoFinal=" & oRsTmpOpc1.Fields!RangoFinal
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdDiagnostico = oRsTmpOpc1.Fields!IdDiagnostico
           End If
           oRsTmpOpc.Fields!LabReferencia = oRsTmpOpc1.Fields!LabReferencia
           oRsTmpOpc.Fields!RangoInicial = oRsTmpOpc1.Fields!RangoInicial
           oRsTmpOpc.Fields!RangoFinal = oRsTmpOpc1.Fields!RangoFinal
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 187
    Me.Refresh
    txtTablaProceso.Text = "his_historicoAtenciones"
    lcSql = "CREATE TABLE [dbo].[his_historicoAtenciones] (" & _
            "    [idPaciente] [int] NOT NULL ," & _
            "    [fecha] [datetime] NOT NULL ," & _
            "    [diagnost] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [cpt] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
            "    [ups] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "create index IX_Paciente on his_historicoAtenciones (idPaciente)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[his_historicoAtenciones] ADD " & _
            "    CONSTRAINT [FK_his_historicoAtenciones_Pacientes] FOREIGN KEY" & _
            "    (" & _
            "        [IdPaciente]" & _
            "    ) REFERENCES [dbo].[Pacientes] (" & _
            "        [IdPaciente]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 188
    Me.Refresh
    txtTablaProceso.Text = "RecetaDetalleItem"
    lcSql = "CREATE TABLE [dbo].[RecetaDetalleItem] (" & _
            "    [idReceta] [int] NOT NULL ," & _
            "    [idItem] [int] NOT NULL ," & _
            "    [DocumentoDespacho] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [CantidadDespachada] [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[RecetaDetalleItem] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_RecetaDetalleItem] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idReceta]," & _
            "    [idItem]," & _
            "    [DocumentoDespacho]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[RecetaDetalleItem] ADD " & _
            " CONSTRAINT [FK_RecetaDetalleItem_RecetaCabecera] FOREIGN KEY" & _
            " (" & _
            "    [idReceta]" & _
            " ) REFERENCES [dbo].[RecetaCabecera] (" & _
            "    [idReceta]" & _
            " )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    '
    DoEvents
    ProgressBar1.Value = 189
    Me.Refresh
    txtTablaProceso.Text = "RecetaEstadosDetalle"
    lcSql = "CREATE TABLE [dbo].[RecetaEstadosDetalle] (" & _
            "    [idEstadoDetalle] [int] NOT NULL ," & _
            "    [Estado] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[RecetaEstadosDetalle] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_RecetaEstadosDetalle] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idEstadoDetalle]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaEstadosDetalle"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from RecetaEstadosDetalle where idEstadoDetalle=" & oRsTmpOpc1.Fields!idEstadoDetalle
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstadoDetalle = oRsTmpOpc1.Fields!idEstadoDetalle
                oRsTmpOpc.Fields!estado = oRsTmpOpc1.Fields!estado
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    '
    DoEvents
    ProgressBar1.Value = 190
    Me.Refresh
    txtTablaProceso.Text = "RecetaDosis"
    lcSql = "CREATE TABLE [dbo].[RecetaDosis] (" & _
            "    [idDosis] [int] NOT NULL ," & _
            "    [NumeroDosis] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[RecetaDosis] WITH NOCHECK ADD " & _
            " CONSTRAINT [PK_RecetaDosis] PRIMARY KEY  CLUSTERED" & _
            " (" & _
            "    [idDosis]" & _
            " )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from RecetaDosis"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from RecetaDosis where idDosis=" & oRsTmpOpc1.Fields!idDosis
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idDosis = oRsTmpOpc1.Fields!idDosis
                oRsTmpOpc.Fields!NumeroDosis = oRsTmpOpc1.Fields!NumeroDosis
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close   '


Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub


Sub cmdMigraUltimaVErsionExternaFUA_Parte1(oConexODBC As Connection, _
                                 oConexHBT As Connection)
    On Error GoTo errMgSS
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim lcSql As String, lbNuevoRegistro As Boolean
    
    '
    DoEvents
    ProgressBar1.Value = 301
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencion"
    
    lcSql = "alter table SisFuaAtencionDIA drop CONSTRAINT FK_SisFuaAtencionDIA_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table SisFuaAtencionINS drop CONSTRAINT FK_SisFuaAtencionINS_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table SisFuaAtencionMED drop CONSTRAINT FK_SisFuaAtencionMED_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table SisFuaAtencionPRO drop CONSTRAINT FK_SisFuaAtencionPRO_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table SisFuaAtencionSMI drop CONSTRAINT FK_SisFuaAtencionSMI_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "alter table SisFuaAtencion drop CONSTRAINT PK_SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "DROP INDEX indNumeroFua ON  SisFuaAtencion"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencion] ALTER COLUMN [FuaDisa] varchar(3) NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencion] ALTER COLUMN [FuaLote] varchar(2) NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencion] ALTER COLUMN [FuaNumero] varchar(16) NOT NULL"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "CREATE INDEX indNumeroFua  ON SisFuaAtencion (FuaDisa,FuaLote,FuaNumero)"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencion] WITH NOCHECK ADD"
    lcSql = lcSql + " CONSTRAINT [PK_SisFuaAtencion] PRIMARY KEY CLUSTERED ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencionDIA] WITH NOCHECK ADD "
    lcSql = lcSql + " CONSTRAINT [FK_SisFuaAtencionDIA_SisFuaAtencion] FOREIGN KEY"
    lcSql = lcSql + " ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero]) REFERENCES [SisFuaAtencion] ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencionINS] WITH NOCHECK ADD "
    lcSql = lcSql + " CONSTRAINT [FK_SisFuaAtencionINS_SisFuaAtencion] FOREIGN KEY"
    lcSql = lcSql + " ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero]) REFERENCES [SisFuaAtencion] ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencionMED] WITH NOCHECK ADD "
    lcSql = lcSql + " CONSTRAINT [FK_SisFuaAtencionMED_SisFuaAtencion] FOREIGN KEY"
    lcSql = lcSql + " ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero]) REFERENCES [SisFuaAtencion] ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencionPRO] WITH NOCHECK ADD  "
    lcSql = lcSql + " CONSTRAINT [FK_SisFuaAtencionPRO_SisFuaAtencion] FOREIGN KEY"
    lcSql = lcSql + " ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero]) REFERENCES [SisFuaAtencion] ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "ALTER TABLE [dbo].[SisFuaAtencionSMI] WITH NOCHECK ADD  "
    lcSql = lcSql + " CONSTRAINT [FK_SisFuaAtencionSMI_SisFuaAtencion] FOREIGN KEY"
    lcSql = lcSql + " ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero]) REFERENCES [SisFuaAtencion] ([idCuentaAtencion],[FuaDisa],[FuaLote],[FuaNumero])"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    Exit Sub
errMgSS:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
                                 
End Sub


Sub MigraUltimaVersion_TablaSIGH_Parte9(oConexHBT As Connection, oConexODBC As Connection)
On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean

    DoEvents
    ProgressBar1.Value = 217
    Me.Refresh
    txtTablaProceso.Text = "CajaDevoluciones"
    lcSql = "create table CajaDevoluciones" & _
    "(" & _
    "    idDevolucion int identity(1,1) primary key," & _
    "    idComprobantePago int not null," & _
    "    montoDevuelto money not null," & _
    "    montoTotal money not null," & _
    "    fechaDevolucion datetime not null," & _
    "    idUsuario int not null," & _
    "    motivo varchar(2000) not null," & _
    "    foreign key (idComprobantePago) references CajaComprobantesPago (idComprobantePago)" & _
    ")"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 218
    Me.Refresh
    txtTablaProceso.Text = "PacientesMovimientos"
    lcSql = "CREATE TABLE [dbo].[PacientesMovimientos] (" & _
    "    [IdCuentaAtencion] [int] NOT NULL ," & _
    "    [Peso] [money] NULL ," & _
    "    [Talla] [money] NULL ," & _
    "    [idDxNutricional] [int] NULL ," & _
    "    [GrafXedadEnMeses] [int] NULL ," & _
    "    [GrafYpercentilTE] [int] NULL ," & _
    "    [GrafYpercentilPT] [int] NULL ," & _
    "    [GrafYpercentilPE] [int] NULL ," & _
    "    [ZetaPT] [money] NULL ," & _
    "    [ZetaTE] [money] NULL ," & _
    "    [ZetaPE] [money] NULL ," & _
    "    [Hemoglobina] [money] NULL ," & _
    "    [Parasitosis] [varchar] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
    ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 219
    Me.Refresh
    txtTablaProceso.Text = "PacientesMovimientosDx"
    lcSql = "CREATE TABLE [dbo].[PacientesMovimientosDx] (" & _
    "    [idDxNutricional] [int] NOT NULL ," & _
    "    [DxNutricional] [varchar] (200) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
    ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 220
    Me.Refresh
    txtTablaProceso.Text = "Alter table PacientesMovimientos"
    lcSql = "ALTER TABLE [dbo].[PacientesMovimientos] WITH NOCHECK ADD" & _
    "    CONSTRAINT [PK_PacientesMovimientos] PRIMARY KEY  CLUSTERED" & _
    "    (" & _
    "        [IdCuentaAtencion]" & _
    "    )  ON [PRIMARY] "
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 221
    Me.Refresh
    txtTablaProceso.Text = "Alter table PacientesMovimientos"
    lcSql = "ALTER TABLE [dbo].[PacientesMovimientosDx] WITH NOCHECK ADD" & _
    "    CONSTRAINT [PK_PacientesMovimientosDx] PRIMARY KEY  CLUSTERED" & _
    "    (" & _
    "        [idDxNutricional]" & _
    "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 222
    Me.Refresh
    txtTablaProceso.Text = "Agregando registros a PacientesMovimientosDx"
    lcSql = "select * from PacientesMovimientosDx"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from PacientesMovimientosDx"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
           lcSql = "select * from PacientesMovimientosDx where idDxNutricional=" & _
                                         oRsTmpOpc1.Fields!idDxNutricional
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idDxNutricional = oRsTmpOpc1.Fields!idDxNutricional
           End If
           oRsTmpOpc.Fields!dxNutricional = oRsTmpOpc1.Fields!dxNutricional
           oRsTmpOpc.Update
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close


Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub

Sub MigraUltimaVersion_TablaSIGH_Parte10(oConexHBT As Connection, oConexODBC As Connection)
On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean

    DoEvents
    ProgressBar1.Value = 223
    Me.Refresh
    txtTablaProceso.Text = "EquivalenciaCPT_SMI"
    lcSql = "CREATE TABLE [dbo].[EquivalenciaCPT_SMI] (" & _
            "    [codigoCPT] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [codigoSMI] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    

    DoEvents
    ProgressBar1.Value = 224
    Me.Refresh
    txtTablaProceso.Text = "ServiciosAtenSimultaneaMov"
    lcSql = "CREATE TABLE [dbo].[ServiciosAtenSimultaneaMov] (" & _
            "    [AScorrelativo] [int] NOT NULL ," & _
            "    [IdAtencion]  [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultaneaMov] WITH NOCHECK ADD" & _
            "    CONSTRAINT [PK_ServiciosAtenSimultaneaMov] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [AScorrelativo]," & _
            "        [IdAtencion]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_ServiciosAtenSimultaneaMov] ON [dbo].[ServiciosAtenSimultaneaMov]([idAtencion]) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    

    DoEvents
    ProgressBar1.Value = 225
    Me.Refresh
    txtTablaProceso.Text = "ServiciosAtenSimultanea"
    lcSql = "CREATE TABLE [dbo].[ServiciosAtenSimultanea] (" & _
            "    [idServicio] [int] NOT NULL ," & _
            "    [idServicioAtencionSimultanea]  [Int] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultanea] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_ServiciosAtencionSimultaneo] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [idServicio]," & _
            "        [idServicioAtencionSimultanea]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultanea] ADD " & _
            "    CONSTRAINT [FK_ServiciosAtenSimultanea_Servicios] FOREIGN KEY" & _
            "    (" & _
            "        [idServicio]" & _
            "    ) REFERENCES [dbo].[Servicios] (" & _
            "        [idServicio]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 226
    txtTablaProceso.Text = "ServiciosAtenSimultaneaFua"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[ServiciosAtenSimultaneaFua] (" & _
            "    [AScorrelativo] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [idAtencion] [int] NOT NULL ," & _
            "    [item] [int] NOT NULL ," & _
            "    [idTipo] [int] NOT NULL ," & _
            "    [idFuaCorrelativo] [int] NOT NULL ," & _
            "    [idFuaIdCuentaAtencion] [int] NULL ," & _
            "    [FuaCodigoPrestacion] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "CREATE  INDEX [IX_ServiciosAtenSimultaneaFua] ON [dbo].[ServiciosAtenSimultaneaFua]([AScorrelativo], [idAtencion]) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic


    DoEvents
    ProgressBar1.Value = 227
    txtTablaProceso.Text = "ServiciosAtenSimultaneaFuaEquiv"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[ServiciosAtenSimultaneaFuaEquiv] (" & _
            "    [Id] [int] NOT NULL ," & _
            "    [ups] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [EdadInicio] [int] NOT NULL ," & _
            "    [EdadFinal] [int] NOT NULL ," & _
            "    [idTipoEdad] [int] NOT NULL ," & _
            "    [FuaCodigoPrestacion] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [DxTipo] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [DxCodigo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [PesoKgMenor] [money] NOT NULL ," & _
            "    [PesoKgMayor] [Money] not null" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultaneaFuaEquiv] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_ServiciosAtenSimultaneaFuaEquiv] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [ID]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultaneaFuaEquiv] ADD " & _
            "    CONSTRAINT [FK_ServiciosAtenSimultaneaFuaEquiv_TiposEdad] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoEdad]" & _
            "    ) REFERENCES [dbo].[TiposEdad] (" & _
            "        [IdTipoEdad]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from ServiciosAtenSimultaneaFuaEquiv"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ServiciosAtenSimultaneaFuaEquiv where id=" & oRsTmpOpc1.Fields!ID
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!ID = oRsTmpOpc1.Fields!ID
                oRsTmpOpc.Fields!ups = oRsTmpOpc1.Fields!ups
                oRsTmpOpc.Fields!EdadInicio = oRsTmpOpc1.Fields!EdadInicio
                oRsTmpOpc.Fields!EdadFinal = oRsTmpOpc1.Fields!EdadFinal
                oRsTmpOpc.Fields!IdTipoEdad = oRsTmpOpc1.Fields!IdTipoEdad
                oRsTmpOpc.Fields!FuaCodigoPrestacion = oRsTmpOpc1.Fields!FuaCodigoPrestacion
                oRsTmpOpc.Fields!DxTipo = oRsTmpOpc1.Fields!DxTipo
                oRsTmpOpc.Fields!DxCodigo = oRsTmpOpc1.Fields!DxCodigo
                oRsTmpOpc.Fields!PesoKgMenor = oRsTmpOpc1.Fields!PesoKgMenor
                oRsTmpOpc.Fields!PesoKgMayor = oRsTmpOpc1.Fields!PesoKgMayor
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 228
    txtTablaProceso.Text = "ServiciosAtenSimultaneaImpHIS"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[ServiciosAtenSimultaneaImpHIS] (" & _
        "    [ups] [varchar] (6) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
        "    [grupo] [int] NOT NULL ," & _
        "    [subgrupo] [int] NOT NULL ," & _
        "    [subgrupoOrden] [int] NULL ," & _
        "    [Lab] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
        "    [cpt_dx] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
        "    [idTipo] [int] NOT NULL ," & _
        "    [EdadInicio] [int] NOT NULL ," & _
        "    [EdadFinal] [int] NOT NULL ," & _
        "    [idTipoEdad] [int] NOT NULL ," & _
        "    [PesoKgMenor] [money] NOT NULL ," & _
        "    [PesoKgMayor]  [Money] not null" & _
        " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "select * from ServiciosAtenSimultaneaImpHIS"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from ServiciosAtenSimultaneaImpHIS where ups='" & oRsTmpOpc1.Fields!ups & "'" & _
                   " and grupo=" & oRsTmpOpc1.Fields!Grupo & " and subgrupoOrden=" & oRsTmpOpc1!subgrupoOrden
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!ups = oRsTmpOpc1.Fields!ups
                oRsTmpOpc.Fields!Grupo = oRsTmpOpc1.Fields!Grupo
                oRsTmpOpc.Fields!subgrupo = oRsTmpOpc1.Fields!subgrupo
                oRsTmpOpc.Fields!subgrupoOrden = oRsTmpOpc1.Fields!subgrupoOrden
                oRsTmpOpc.Fields!cpt_dx = oRsTmpOpc1.Fields!cpt_dx
                oRsTmpOpc.Fields!lab = oRsTmpOpc1.Fields!lab
                oRsTmpOpc.Fields!IdTipo = oRsTmpOpc1.Fields!IdTipo
                oRsTmpOpc.Fields!EdadInicio = oRsTmpOpc1.Fields!EdadInicio
                oRsTmpOpc.Fields!EdadFinal = oRsTmpOpc1.Fields!EdadFinal
                oRsTmpOpc.Fields!IdTipoEdad = oRsTmpOpc1.Fields!IdTipoEdad
                oRsTmpOpc.Fields!PesoKgMenor = oRsTmpOpc1.Fields!PesoKgMenor
                oRsTmpOpc.Fields!PesoKgMayor = oRsTmpOpc1.Fields!PesoKgMayor
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    
    
    DoEvents
    ProgressBar1.Value = 228
    txtTablaProceso.Text = "TiposServiciosIntermedios"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[TiposServiciosIntermedios] (" & _
            "    [idTipo] [int] NOT NULL ," & _
            "    [Tipo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[TiposServiciosIntermedios] WITH NOCHECK ADD " & _
            "    CONSTRAINT [PK_TiposServiciosIntermedios] PRIMARY KEY  CLUSTERED" & _
            "    (" & _
            "        [IdTipo]" & _
            "    )  ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from TiposServiciosIntermedios"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from TiposServiciosIntermedios where idTipo=" & oRsTmpOpc1.Fields!IdTipo
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdTipo = oRsTmpOpc1.Fields!IdTipo
                oRsTmpOpc.Fields!Tipo = oRsTmpOpc1.Fields!Tipo
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultaneaFua] ADD " & _
            "    CONSTRAINT [FK_ServiciosAtenSimultaneaFua_TiposServiciosIntermedios] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipo]" & _
            "    ) REFERENCES [dbo].[TiposServiciosIntermedios] (" & _
            "        [IdTipo]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE [dbo].[ServiciosAtenSimultaneaImpHIS] ADD " & _
            "    CONSTRAINT [FK_ServiciosAtenSimultaneaImpHIS_TiposEdad] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipoEdad]" & _
            "    ) REFERENCES [dbo].[TiposEdad] (" & _
            "        [IdTipoEdad]" & _
            "    )," & _
            "    CONSTRAINT [FK_ServiciosAtenSimultaneaImpHIS_TiposServiciosIntermedios] FOREIGN KEY" & _
            "    (" & _
            "        [IdTipo]" & _
            "    ) REFERENCES [dbo].[TiposServiciosIntermedios] (" & _
            "        [IdTipo]" & _
            "    )"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    
    lcSql = "ALTER TABLE facturacionServicioDespacho add  GrupoHIS int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE facturacionServicioDespacho add  SubGrupoHIS int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesDiagnosticos add  GrupoHIS int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE atencionesDiagnosticos add  SubGrupoHIS int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    lcSql = "ALTER TABLE recetadetalle ALTER COLUMN MotivoAnulacionMedico varchar(300) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE recetadetalle ALTER COLUMN Observaciones varchar(300) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen ALTER COLUMN regenerarDias varchar(7) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen ALTER COLUMN regenerarHora varchar(5) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE farmAlmacen ALTER COLUMN regenerarEstado varchar(7) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN DireccionDomicilio varchar(100) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN NombreAcompaniante varchar(100) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN SisCodigo varchar(2) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN NroReferenciaOrigen varchar(20) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "ALTER TABLE AtencionesDatosAdicionales ALTER COLUMN NroReferenciaDestino varchar(20) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
   'Actualiza 2 columnas con null en la tabla medicamentos/insumos para que se puedan ver y dar precios
    lcSql = "update factcatalogoBienesInsumos set idGrupoFarmacologico=999 where idGrupoFarmacologico is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update factcatalogoBienesInsumos set idSubGrupoFarmacologico=999 where idSubGrupoFarmacologico is null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub
    

Sub MigraUltimaVersion_TablaSIGH_Parte11(oConexHBT As Connection, oConexODBC As Connection)
On Error GoTo errMg
    Dim oRsTmpOpc As New ADODB.Recordset
    Dim oRsTmpOpc1 As New ADODB.Recordset
    Dim lbNuevoRegistro As Boolean

    DoEvents
    ProgressBar1.Value = 229
    Me.Refresh
    txtTablaProceso.Text = "Medicos ADD rne"
    lcSql = "ALTER TABLE  medicos ADD rne varchar(50)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic

    DoEvents
    ProgressBar1.Value = 230
    Me.Refresh
    txtTablaProceso.Text = "Medicos ADD egresado"
    lcSql = "ALTER TABLE  medicos ADD egresado bit null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 231
    Me.Refresh
    txtTablaProceso.Text = "atencionesnacimientos ADD docIdentidad"
    lcSql = "ALTER TABLE  atencionesnacimientos  ADD docIdentidad varchar(20)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 232
    Me.Refresh
    txtTablaProceso.Text = "atencionesnacimientos ADD IdDocIdentidad"
    lcSql = "ALTER TABLE  atencionesnacimientos  ADD IdDocIdentidad int  null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 233
    Me.Refresh
    txtTablaProceso.Text = "SuSalud_ups"
    lcSql = "CREATE TABLE [dbo].[SuSalud_ups] (" & _
            "    [Codigo] [varchar] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ," & _
            "    [Descripcion] [varchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 234
    Me.Refresh
    txtTablaProceso.Text = "servicios ADD codigoServicioSuSalud"
    lcSql = "ALTER TABLE  servicios ADD codigoServicioSuSalud varchar(7) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 235
    Me.Refresh
    txtTablaProceso.Text = "servicios ADD codigoServicioFUA"
    lcSql = "ALTER TABLE  servicios ADD codigoServicioFUA varchar(6) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 236
    Me.Refresh
    txtTablaProceso.Text = "servicios ADD FuaTipoAnexo2015"
    lcSql = "ALTER TABLE  servicios ADD FuaTipoAnexo2015 int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update servicios set FuaTipoAnexo2015=3 WHERE FuaTipoAnexo2015 IS NULL OR FuaTipoAnexo2015=''"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic


    DoEvents
    ProgressBar1.Value = 237
    Me.Refresh
    txtTablaProceso.Text = "ALTER TABLE farmMovimientoDetalle"
    lcSql = "ALTER TABLE farmMovimientoDetalle add DocumentoNumero varchar(20) null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 238
    txtTablaProceso.Text = "NotaCreditoDebitoTipoNota"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[NotaCreditoDebitoTipoNota](" & _
            "    [IdTipoNota] [int] NULL," & _
            "    [TipoNota] [varchar](50) NULL," & _
            "    [NroSerie] [char](3) NULL," & _
            "    [NroDocumento] [char](12) NULL," & _
            "    [NroDocumentoInicial] [char](12) NULL," & _
            "    [NroDocumentoFinal] [char](12) NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from NotaCreditoDebitoTipoNota"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from NotaCreditoDebitoTipoNota where IdTipoNota=" & oRsTmpOpc1.Fields!IdTipoNota
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdTipoNota = oRsTmpOpc1.Fields!IdTipoNota
                oRsTmpOpc.Fields!TipoNota = oRsTmpOpc1.Fields!TipoNota
                oRsTmpOpc.Fields!NroSerie = oRsTmpOpc1.Fields!NroSerie
                oRsTmpOpc.Fields!NroDocumento = oRsTmpOpc1.Fields!NroDocumento
                oRsTmpOpc.Fields!NroDocumentoInicial = oRsTmpOpc1.Fields!NroDocumentoInicial
                oRsTmpOpc.Fields!NroDocumentoFinal = oRsTmpOpc1.Fields!NroDocumentoFinal
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 239
    txtTablaProceso.Text = "NotaCreditoDebito"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[NotaCreditoDebito](" & _
            "    [IdNota] [int] IDENTITY(1,1) NOT NULL," & _
            "    [IdComprobantePago] [int] NULL," & _
            "    [IdTipoNota] [int] NULL," & _
            "    [NroSerie] [char](3) NULL," & _
            "    [NroDocumento] [varchar](12) NULL," & _
            "    [RazonSocial] [varchar](50) NULL," & _
            "    [RUC] [char](11) NULL," & _
            "    [SubTotal] [money] NULL," & _
            "    [IGV] [money] NULL," & _
            "    [Total] [money] NOT NULL," & _
            "    [IdUsuarioAutoriza] [int] NULL," & _
            "    [FechaAprueba] [datetime] NULL," & _
            "    [TipoCambio] [money] NULL," & _
            "    [Observaciones] [varchar](500) NULL," & _
            "    [IdEstadoNota] [int] NULL," & _
            "    [FechaPagado] [datetime] NULL," & _
            "    [IdGestionCaja] [int] NULL," & _
            "    [IdPaciente] [int] NULL," & _
            "    [IdCajero] [int] NULL," & _
            "    [idTurno] [int] NULL,[idCaja] [int] NULL," & _
            "    [idFarmacia] [int] NULL,[idMotivo] [int] NULL," & _
            "    [Direccion] [varchar](50) NULL,[TipoAnulacion] [bit] NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 240
    txtTablaProceso.Text = "NotaCreditoDebitoEstadoNota"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[NotaCreditoDebitoEstadoNota](" & _
            "    [IdEstado] [int] NULL," & _
            "    [EstadoNota] [varchar](50) NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from NotaCreditoDebitoEstadoNota"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from NotaCreditoDebitoEstadoNota where IdEstado=" & oRsTmpOpc1.Fields!idEstado
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!idEstado = oRsTmpOpc1.Fields!idEstado
                oRsTmpOpc.Fields!EstadoNota = oRsTmpOpc1.Fields!EstadoNota
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    DoEvents
    ProgressBar1.Value = 241
    txtTablaProceso.Text = "NotaCreditoDebitoMotivo"
    Me.Refresh
    lcSql = "CREATE TABLE [dbo].[NotaCreditoDebitoMotivo](" & _
            "    [IdMotivo] [int] NULL," & _
            "    [Motivo] [varchar](50) NULL" & _
            " ) ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from NotaCreditoDebitoMotivo"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lcSql = "select * from NotaCreditoDebitoMotivo where IdMotivo=" & oRsTmpOpc1.Fields!IdMotivo
           If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
           oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
           lbNuevoRegistro = True
           If oRsTmpOpc.RecordCount > 0 Then
              lbNuevoRegistro = False
           End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!IdMotivo = oRsTmpOpc1.Fields!IdMotivo
                oRsTmpOpc.Fields!Motivo = oRsTmpOpc1.Fields!Motivo
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
Exit Sub
errMg:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
End Sub


Sub cmdMigraUltimaVErsionExternaFUA_Parte2(oConexODBC As Connection, _
                                 oConexHBT As Connection)
    On Error GoTo errMgSS
    Dim oRsTmpOpc As New Recordset
    Dim oRsTmpOpc1 As New Recordset
    Dim lbNuevoRegistro As Boolean
    
    '
    DoEvents
    ProgressBar1.Value = 303
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionPREST"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionPREST] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [FuaCodigoPrestacion] [varchar] (3) COLLATE Modern_Spanish_CI_AS NOT NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 304
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionFUAS"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionFUAS] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [FuaIdCuentaAtencion] [int] NOT NULL ," & _
            "    [fechainicio] [datetime] NOT NULL ," & _
            "    [fechafinal] [datetime] NOT NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    '
    DoEvents
    ProgressBar1.Value = 305
    Me.Refresh
    txtTablaProceso.Text = "SisFuaAtencionRN"
    lcSql = "CREATE TABLE [dbo].[SisFuaAtencionRN] (" & _
            "    [idCuentaAtencion] [int] NOT NULL ," & _
            "    [FuaDocIdentidad] [varchar] (20) COLLATE Modern_Spanish_CI_AS NOT NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 306
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuacolegioCodigo"
    lcSql = "ALTER TABLE  sisfuaatencion add FuacolegioCodigo varchar(20) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
        
    DoEvents
    ProgressBar1.Value = 307
    Me.Refresh
    txtTablaProceso.Text = "ALTER TABLE  sisfuaatencion  add FuacolegioNivel"
    lcSql = "ALTER TABLE  sisfuaatencion  add FuacolegioNivel varchar(1) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 308
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuacolegioGrado"
    lcSql = "ALTER TABLE  sisfuaatencion  add FuacolegioGrado varchar(1) COLLATE SQL_Latin1_General_CP1_CI_AS"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 309
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuacolegioSeccion"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuacolegioSeccion varchar(5) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    
    DoEvents
    ProgressBar1.Value = 310
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuacolegioTurno"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuacolegioTurno varchar(1) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 311
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD Fuaetnia"
    lcSql = "ALTER TABLE  sisfuaatencion ADD Fuaetnia varchar(2) COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 312
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuafechaFallecimiento"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuafechaFallecimiento datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 313
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaUPS"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuaUPS varchar(6)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 314
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaCodAutorizacion"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuaCodAutorizacion varchar(50)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 315
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaFechaCorteAdm"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuaFechaCorteAdm datetime null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 316
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaVersionFormato"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuaVersionFormato varchar(2)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "update SisFuaAtencion set FuaVersionFormato='A' WHERE FuaVersionFormato IS NULL OR FuaVersionFormato=''"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 317
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaTipoAnexo2015"
    lcSql = "alter table SisFuaAtencion ADD FuaTipoAnexo2015 int null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    DoEvents
    ProgressBar1.Value = 318
    Me.Refresh
    txtTablaProceso.Text = "sisfuaatencion ADD FuaCodOferFlexible"
    lcSql = "ALTER TABLE  sisfuaatencion ADD FuaCodOferFlexible varchar(20)  COLLATE SQL_Latin1_General_CP1_CI_AS null"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
  
    DoEvents
    ProgressBar1.Value = 319
    Me.Refresh
    txtTablaProceso.Text = "SisFuaUPServicios"
    lcSql = "CREATE TABLE [dbo].[SisFuaUPServicios] (" & _
            "    [UPS] [varchar] (6) COLLATE Modern_Spanish_CI_AS NOT NULL ," & _
            "    [descripcion] [varchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL" & _
            ") ON [PRIMARY]"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    
    lcSql = "select * from SisFuaUPServicios"
    If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
    oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
    lcSql = "select * from SisFuaUPServicios"
    If oRsTmpOpc1.State = 1 Then oRsTmpOpc1.Close
    oRsTmpOpc1.Open lcSql, oConexHBT, adOpenKeyset, adLockOptimistic
    If oRsTmpOpc1.RecordCount > 0 Then
        oRsTmpOpc1.MoveFirst
        Do While Not oRsTmpOpc1.EOF
           lbNuevoRegistro = True
            lcSql = "select * from SisFuaUPServicios where UPS='" & Trim(oRsTmpOpc1.Fields!ups) & "'"
            If oRsTmpOpc.State = 1 Then oRsTmpOpc.Close
            oRsTmpOpc.Open lcSql, oConexODBC, adOpenKeyset, adLockOptimistic
            If oRsTmpOpc.RecordCount > 0 Then
               lbNuevoRegistro = False
            End If
           If lbNuevoRegistro = True Then
                oRsTmpOpc.AddNew
                oRsTmpOpc.Fields!ups = oRsTmpOpc1.Fields!ups
                oRsTmpOpc.Fields!Descripcion = oRsTmpOpc1.Fields!Descripcion
                oRsTmpOpc.Update
           End If
           oRsTmpOpc1.MoveNext
        Loop
    End If
    oRsTmpOpc1.Close
    
    
    
    Exit Sub
errMgSS:
    If Err.Number = -2147217900 Or Err.Number = -2147217865 Then
       Resume Next
    Else
       MsgBox Err.Description
       Resume
    End If
                                 
End Sub

