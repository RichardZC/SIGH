VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form HerrVisorFormatoFUA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diseño Fua"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "HerrVisorFormatoFUA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab TabsDominios 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "Frontal"
      TabPicture(0)   =   "HerrVisorFormatoFUA.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgFrontal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Espaldar"
      TabPicture(1)   =   "HerrVisorFormatoFUA.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgRespaldar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Image imgRespaldar 
         Height          =   3495
         Left            =   -74880
         OLEDropMode     =   1  'Manual
         Picture         =   "HerrVisorFormatoFUA.frx":047A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   7035
      End
      Begin VB.Image imgFrontal 
         Height          =   8415
         Left            =   120
         Picture         =   "HerrVisorFormatoFUA.frx":59B2
         Stretch         =   -1  'True
         Top             =   480
         Width           =   6945
      End
   End
End
Attribute VB_Name = "HerrVisorFormatoFUA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: MINSA - OGEI - OIT
'        Aplicativo: SisGalenPlus v.3
'        Programa: Visualizar el Formato FUA seleccionado
'        Programado por: Cachay F
'        Fecha: Setiembre 2015
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mc_FormatoFua As String
Attribute mc_FormatoFua.VB_VarHelpID = -1
Dim mc_TipoAnexo As String

Property Let FormatoFua(lValue As String)
   mc_FormatoFua = lValue
End Property
Property Get FormatoFua() As String
   FormatoFua = mc_FormatoFua
End Property
Property Let TipoAnexo(sValue As String)
   mc_TipoAnexo = sValue
End Property
Property Get TipoAnexo() As String
   TipoAnexo = mc_TipoAnexo
End Property

Public Sub MostrarFormulario()
    Me.Show 1
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    Ruta = App.Path + "\Imagenes\FUA\" + mc_FormatoFua + TipoAnexo + "-1.jpg"
    imgFrontal.Picture = LoadPicture(Ruta)
    imgFrontal.Width = 8025: imgFrontal.Height = 9375
    
    Ruta = App.Path + "\Imagenes\FUA\" + mc_FormatoFua + TipoAnexo + "-2.jpg"
    imgRespaldar.Picture = LoadPicture(Ruta)
    imgRespaldar.Width = 7995: imgRespaldar.Height = 4695
End Sub

