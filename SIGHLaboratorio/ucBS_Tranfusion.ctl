VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ucBS_Tranfusion 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13185
   ScaleHeight     =   7365
   ScaleWidth      =   13185
   Begin TabDlg.SSTab BS000_00 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   11668
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Historial del Paciente"
      TabPicture(0)   =   "ucBS_Tranfusion.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "BS000_01(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "BS000_02(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Conducción de la Transfusion"
      TabPicture(1)   =   "ucBS_Tranfusion.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BS000_01(1)"
      Tab(1).Control(1)=   "BS000_03(19)"
      Tab(1).Control(2)=   "BS000_03(18)"
      Tab(1).Control(3)=   "BS000_02(5)"
      Tab(1).Control(4)=   "BS000_02(3)"
      Tab(1).Control(5)=   "BS000_04(9)"
      Tab(1).Control(6)=   "BS000_04(10)"
      Tab(1).Control(7)=   "BS000_02(4)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Reacción adversa transfucional"
      TabPicture(2)   =   "ucBS_Tranfusion.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BS000_04(35)"
      Tab(2).Control(1)=   "BS000_04(34)"
      Tab(2).Control(2)=   "BS000_02(8)"
      Tab(2).Control(3)=   "BS000_02(7)"
      Tab(2).Control(4)=   "BS000_02(6)"
      Tab(2).Control(5)=   "BS000_04(33)"
      Tab(2).Control(6)=   "BS000_04(32)"
      Tab(2).Control(7)=   "BS000_04(31)"
      Tab(2).Control(8)=   "BS000_04(30)"
      Tab(2).Control(9)=   "BS000_04(29)"
      Tab(2).Control(10)=   "BS000_04(28)"
      Tab(2).Control(11)=   "BS000_04(21)"
      Tab(2).Control(12)=   "BS000_04(20)"
      Tab(2).Control(13)=   "BS000_04(27)"
      Tab(2).Control(14)=   "BS000_04(25)"
      Tab(2).Control(15)=   "BS000_04(24)"
      Tab(2).Control(16)=   "BS000_04(23)"
      Tab(2).Control(17)=   "BS000_04(22)"
      Tab(2).Control(18)=   "BS000_03(60)"
      Tab(2).Control(19)=   "BS000_03(59)"
      Tab(2).Control(20)=   "BS000_03(58)"
      Tab(2).Control(21)=   "BS000_03(57)"
      Tab(2).Control(22)=   "BS000_03(56)"
      Tab(2).Control(23)=   "BS000_03(55)"
      Tab(2).Control(24)=   "BS000_03(54)"
      Tab(2).Control(25)=   "BS000_03(53)"
      Tab(2).Control(26)=   "BS000_03(52)"
      Tab(2).Control(27)=   "BS000_03(51)"
      Tab(2).Control(28)=   "BS000_03(49)"
      Tab(2).Control(29)=   "BS000_03(48)"
      Tab(2).Control(30)=   "BS000_03(47)"
      Tab(2).Control(31)=   "BS000_03(46)"
      Tab(2).Control(32)=   "BS000_03(45)"
      Tab(2).Control(33)=   "BS000_03(44)"
      Tab(2).Control(34)=   "BS000_03(37)"
      Tab(2).Control(35)=   "BS000_03(36)"
      Tab(2).Control(36)=   "BS000_03(42)"
      Tab(2).Control(37)=   "BS000_03(41)"
      Tab(2).Control(38)=   "BS000_03(40)"
      Tab(2).Control(39)=   "BS000_03(39)"
      Tab(2).Control(40)=   "BS000_03(38)"
      Tab(2).Control(41)=   "BS000_01(2)"
      Tab(2).ControlCount=   42
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   35
         Left            =   -71940
         TabIndex        =   171
         Top             =   840
         Width           =   6675
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   34
         Left            =   -71940
         MaxLength       =   35
         TabIndex        =   169
         Top             =   480
         Width           =   1515
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Index           =   8
         Left            =   -74820
         TabIndex        =   157
         Top             =   4320
         Width           =   12615
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   10680
            TabIndex        =   168
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Alergia Urticaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   10680
            TabIndex        =   167
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hepatitis post transfusional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   6360
            TabIndex        =   166
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Transfusión aosciada a enfermedad transmisible"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   6360
            TabIndex        =   165
            Top             =   720
            Width           =   4455
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Contaminación bacteriana"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   3240
            TabIndex        =   164
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reacción Hemofílica no inmune"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   6360
            TabIndex        =   163
            Top             =   240
            Width           =   2775
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reacción hemofílica inmediata"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   120
            TabIndex        =   162
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fiebre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   161
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Anafilaxis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   160
            Top             =   720
            Width           =   3135
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reacción hemofílica tardía"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   3240
            TabIndex        =   159
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sobrecarga circulatoria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   3240
            TabIndex        =   158
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "El paciente se encuentra en:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   7
         Left            =   -68460
         TabIndex        =   151
         Top             =   3240
         Width           =   6255
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quimioterapia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   4080
            TabIndex        =   156
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sepsis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   155
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tratamiento ATB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   154
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CID"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   2280
            TabIndex        =   153
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Uso metildopa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   2280
            TabIndex        =   152
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reacciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Index           =   6
         Left            =   -74820
         TabIndex        =   140
         Top             =   3240
         Width           =   6255
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dolor Toráxico"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   2280
            TabIndex        =   150
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cianosis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   2280
            TabIndex        =   149
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Edema Facial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   148
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Náuseas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   147
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Escalofríos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   146
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cefalea"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   4080
            TabIndex        =   145
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dolor Lumbar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   2280
            TabIndex        =   144
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hemoglobinuria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4080
            TabIndex        =   143
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Prurito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   4080
            TabIndex        =   142
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   5400
            TabIndex        =   141
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   33
         Left            =   -71040
         MaxLength       =   35
         TabIndex        =   131
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   32
         Left            =   -71040
         MaxLength       =   35
         TabIndex        =   130
         Top             =   2520
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   31
         Left            =   -64080
         MaxLength       =   35
         TabIndex        =   129
         Top             =   2520
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   -64080
         MaxLength       =   35
         TabIndex        =   128
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   -73095
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   118
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   28
         Left            =   -73095
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   117
         Top             =   2520
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   -66135
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   116
         Top             =   2520
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   -66135
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   115
         Top             =   2850
         Width           =   1155
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   -71925
         MaxLength       =   35
         TabIndex        =   113
         Top             =   1920
         Width           =   1515
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   25
         Left            =   -71925
         MaxLength       =   35
         TabIndex        =   110
         Top             =   1560
         Width           =   1515
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   24
         Left            =   -66765
         MaxLength       =   35
         TabIndex        =   109
         Top             =   1575
         Width           =   1515
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   23
         Left            =   -71925
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   106
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   -66765
         MaxLength       =   35
         TabIndex        =   105
         Top             =   1215
         Width           =   1515
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado Clínico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Index           =   4
         Left            =   -74520
         TabIndex        =   97
         Top             =   4680
         Width           =   4095
         Begin VB.ComboBox BS000_05 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   2070
            TabIndex        =   102
            Text            =   "Combo1"
            Top             =   960
            Width           =   1875
         End
         Begin VB.ComboBox BS000_05 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   2070
            TabIndex        =   100
            Text            =   "Combo1"
            Top             =   600
            Width           =   1875
         End
         Begin VB.ComboBox BS000_05 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   2070
            TabIndex        =   98
            Text            =   "Combo1"
            Top             =   240
            Width           =   1875
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de Pulmonar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   555
            TabIndex        =   103
            Top             =   990
            Width           =   1425
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de  Cardiovascular"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   105
            TabIndex        =   101
            Top             =   630
            Width           =   1875
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado de Conciencia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   450
            TabIndex        =   99
            Top             =   270
            Width           =   1530
         End
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   10
         Left            =   -67410
         MaxLength       =   35
         TabIndex        =   72
         Top             =   495
         Width           =   1515
      End
      Begin VB.TextBox BS000_04 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   9
         Left            =   -71370
         MaxLength       =   35
         TabIndex        =   71
         Top             =   480
         Width           =   1515
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parámetros de Transfusión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3615
         Index           =   3
         Left            =   -74520
         TabIndex        =   70
         Top             =   840
         Width           =   11655
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   5520
            MaxLength       =   35
            TabIndex        =   84
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   9360
            MaxLength       =   35
            TabIndex        =   83
            Top             =   570
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   9360
            MaxLength       =   35
            TabIndex        =   82
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   9360
            MaxLength       =   35
            TabIndex        =   81
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   1320
            MaxLength       =   35
            TabIndex        =   80
            Top             =   900
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   5520
            MaxLength       =   35
            TabIndex        =   79
            Top             =   240
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   1320
            MaxLength       =   35
            TabIndex        =   78
            Top             =   570
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   5520
            MaxLength       =   35
            TabIndex        =   77
            Top             =   570
            Width           =   1155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   1320
            MaxLength       =   35
            TabIndex        =   76
            Top             =   240
            Width           =   1155
         End
         Begin MSFlexGridLib.MSFlexGrid BS000_10 
            Height          =   1935
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   1560
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   7
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   16777215
            ForeColorSel    =   16777215
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "x minuto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   6720
            TabIndex        =   127
            Top             =   270
            Width           =   615
         End
         Begin VB.Label BS000_03 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSFUSIÓN ACTUAL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   4920
            TabIndex        =   104
            Top             =   1320
            Width           =   1890
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ºC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   10575
            TabIndex        =   96
            Top             =   600
            Width           =   180
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "x minuto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   10575
            TabIndex        =   95
            Top             =   270
            Width           =   615
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "mmHG"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   2565
            TabIndex        =   94
            Top             =   600
            Width           =   450
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sangrado/Plaquetas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   3975
            TabIndex        =   93
            Top             =   930
            Width           =   1455
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Volúmen Sangrado"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   7935
            TabIndex        =   92
            Top             =   930
            Width           =   1335
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Temperatura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   8340
            TabIndex        =   91
            Top             =   600
            Width           =   930
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Respiraciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   8325
            TabIndex        =   90
            Top             =   270
            Width           =   990
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cianosis"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   645
            TabIndex        =   89
            Top             =   930
            Width           =   585
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Frecuencia de Pulso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   4005
            TabIndex        =   88
            Top             =   270
            Width           =   1425
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Palidez/Hematocrito"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   4005
            TabIndex        =   87
            Top             =   600
            Width           =   1425
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Presión Arterial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   135
            TabIndex        =   86
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   930
            TabIndex        =   85
            Top             =   270
            Width           =   345
         End
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reacciones Adversas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   5
         Left            =   -69120
         TabIndex        =   56
         Top             =   4680
         Width           =   6255
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Otros"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   5400
            TabIndex        =   69
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fiebre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   4080
            TabIndex        =   68
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vómitos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   4080
            TabIndex        =   67
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Disnea"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   4080
            TabIndex        =   66
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Escalofríos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   2280
            TabIndex        =   65
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hipotensión"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   2280
            TabIndex        =   64
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Urticaria"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   63
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dolor Subesternal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hemoglobinemia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sangrado en capa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Coombs positivo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   59
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Desasosiego"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   2280
            TabIndex        =   58
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox BS000_09 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dolor Perfusión"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   2055
         End
      End
      Begin VB.Frame BS000_02 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1. Datos del Paciente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6045
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   12555
         Begin VB.ComboBox BS000_05 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   7350
            TabIndex        =   38
            Top             =   225
            Width           =   1875
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   37
            Top             =   600
            Width           =   4155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   36
            Top             =   225
            Width           =   4155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   35
            Top             =   975
            Width           =   4155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   10290
            MaxLength       =   35
            TabIndex        =   34
            Top             =   975
            Width           =   1515
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   7350
            MaxLength       =   35
            TabIndex        =   33
            Top             =   975
            Width           =   1515
         End
         Begin VB.ComboBox BS000_05 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   7350
            TabIndex        =   32
            Text            =   "Combo1"
            Top             =   1365
            Width           =   1875
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   1410
            MaxLength       =   35
            TabIndex        =   31
            Top             =   1350
            Width           =   4155
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   10290
            MaxLength       =   35
            TabIndex        =   30
            Top             =   600
            Width           =   1515
         End
         Begin VB.TextBox BS000_04 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   10290
            MaxLength       =   35
            TabIndex        =   29
            Top             =   225
            Width           =   1515
         End
         Begin VB.Frame BS000_02 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Datos de la Unidad de Sangre"
            ForeColor       =   &H00000000&
            Height          =   1935
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   5655
            Begin VB.ComboBox BS000_05 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   3915
               TabIndex        =   25
               Text            =   "Combo1"
               Top             =   630
               Width           =   1575
            End
            Begin VB.TextBox BS000_04 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   8
               Left            =   1035
               MaxLength       =   35
               TabIndex        =   24
               Top             =   630
               Width           =   915
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ST"
               Height          =   195
               Index           =   0
               Left            =   1560
               TabIndex        =   23
               Top             =   1080
               Width           =   855
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PG"
               Height          =   195
               Index           =   1
               Left            =   2880
               TabIndex        =   22
               Top             =   1080
               Width           =   855
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PFC"
               Height          =   195
               Index           =   2
               Left            =   4440
               TabIndex        =   21
               Top             =   1080
               Width           =   735
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PQ"
               Height          =   195
               Index           =   3
               Left            =   1560
               TabIndex        =   20
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "CRIO"
               Height          =   195
               Index           =   4
               Left            =   2880
               TabIndex        =   19
               Top             =   1320
               Width           =   855
            End
            Begin VB.CheckBox BS000_07 
               BackColor       =   &H00C0C0C0&
               Caption         =   "GRL"
               Height          =   195
               Index           =   5
               Left            =   4440
               TabIndex        =   18
               Top             =   1320
               Width           =   735
            End
            Begin MSMask.MaskEdBox BS000_06 
               Height          =   285
               Index           =   1
               Left            =   1470
               TabIndex        =   54
               Top             =   240
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               ForeColor       =   0
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
               PromptChar      =   " "
            End
            Begin VB.Label BS000_03 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   13
               Left            =   0
               TabIndex        =   55
               Top             =   270
               Width           =   1320
            End
            Begin VB.Label BS000_03 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Grupo Sanguíneo"
               Height          =   195
               Index           =   15
               Left            =   2550
               TabIndex        =   28
               Top             =   660
               Width           =   1275
            End
            Begin VB.Label BS000_03 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nº de Bolsa"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   14
               Left            =   105
               TabIndex        =   27
               Top             =   660
               Width           =   825
            End
            Begin VB.Label BS000_03 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Componentes"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   26
               Top             =   1200
               Width           =   1065
            End
         End
         Begin VB.Frame BS000_02 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Antecedentes"
            ForeColor       =   &H00000000&
            Height          =   1215
            Index           =   2
            Left            =   5880
            TabIndex        =   4
            Top             =   2400
            Width           =   6495
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Consumo de alcohol"
               Height          =   195
               Index           =   9
               Left            =   120
               TabIndex        =   16
               Top             =   960
               Width           =   2055
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Linfogranuloma venérea"
               Height          =   195
               Index           =   4
               Left            =   2640
               TabIndex        =   15
               Top             =   480
               Width           =   2295
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Contacto sexual de riesgo"
               Height          =   195
               Index           =   1
               Left            =   2640
               TabIndex        =   14
               Top             =   240
               Width           =   2295
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Otras alergias"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Alergias a medicamentos"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   2295
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Drogadicción endovenosa"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   2175
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Gonorrea"
               Height          =   195
               Index           =   2
               Left            =   5160
               TabIndex        =   10
               Top             =   240
               Width           =   1215
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Hepatitis"
               Height          =   195
               Index           =   10
               Left            =   2640
               TabIndex        =   9
               Top             =   960
               Width           =   1215
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Diálisis"
               Height          =   195
               Index           =   7
               Left            =   2640
               TabIndex        =   8
               Top             =   720
               Width           =   1215
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Otros"
               Height          =   195
               Index           =   11
               Left            =   5160
               TabIndex        =   7
               Top             =   960
               Width           =   975
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Asma"
               Height          =   195
               Index           =   8
               Left            =   5160
               TabIndex        =   6
               Top             =   720
               Width           =   1215
            End
            Begin VB.CheckBox BS000_08 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Lués"
               Height          =   195
               Index           =   5
               Left            =   5160
               TabIndex        =   5
               Top             =   480
               Width           =   1215
            End
         End
         Begin MSFlexGridLib.MSFlexGrid BS000_10 
            Height          =   1935
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   3960
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   3413
            _Version        =   393216
            Cols            =   7
            BackColor       =   16777215
            ForeColor       =   0
            BackColorFixed  =   12632256
            ForeColorFixed  =   0
            BackColorSel    =   16777215
            ForeColorSel    =   16777215
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox BS000_06 
            Height          =   285
            Index           =   0
            Left            =   7350
            TabIndex        =   39
            Top             =   600
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
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
            PromptChar      =   " "
         End
         Begin VB.Label BS000_03 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSFUSIONES PREVIAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   17
            Left            =   5220
            TabIndex        =   53
            Top             =   3750
            Width           =   2130
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   6885
            TabIndex        =   52
            Top             =   255
            Width           =   360
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Materno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   51
            Top             =   630
            Width           =   1200
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   735
            TabIndex        =   50
            Top             =   1005
            Width           =   630
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido Paterno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   49
            Top             =   255
            Width           =   1170
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   9825
            TabIndex        =   48
            Top             =   1020
            Width           =   405
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Historia Clínica"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   6000
            TabIndex        =   47
            Top             =   1005
            Width           =   1320
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grupo Sanguíneo"
            Height          =   195
            Index           =   12
            Left            =   5880
            TabIndex        =   46
            Top             =   1395
            Width           =   1380
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   780
            TabIndex        =   45
            Top             =   1350
            Width           =   555
         End
         Begin VB.Label BS000_03 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Kg."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   11880
            TabIndex        =   44
            Top             =   315
            Width           =   240
         End
         Begin VB.Label BS000_03 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "mt."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   11880
            TabIndex        =   43
            Top             =   690
            Width           =   240
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Peso"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   9900
            TabIndex        =   42
            Top             =   315
            Width           =   345
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Talla"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   9915
            TabIndex        =   41
            Top             =   690
            Width           =   330
         End
         Begin VB.Label BS000_03 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Nacimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   5880
            TabIndex        =   40
            Top             =   630
            Width           =   1320
         End
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Premedicación previa a transfusión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   60
         Left            =   -74880
         TabIndex        =   172
         Top             =   930
         Width           =   2865
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad transfundida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   59
         Left            =   -73620
         TabIndex        =   170
         Top             =   570
         Width           =   1605
      End
      Begin VB.Label BS000_03 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DESPUES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   58
         Left            =   -64080
         TabIndex        =   139
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label BS000_03 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DESPUES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   57
         Left            =   -71040
         TabIndex        =   138
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label BS000_03 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ANTES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   56
         Left            =   -67080
         TabIndex        =   137
         Top             =   2310
         Width           =   2640
      End
      Begin VB.Label BS000_03 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ANTES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   55
         Left            =   -73095
         TabIndex        =   136
         Top             =   2310
         Width           =   1800
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x minuto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   54
         Left            =   -69825
         TabIndex        =   135
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "mmHG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   53
         Left            =   -69825
         TabIndex        =   134
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x minuto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   52
         Left            =   -62865
         TabIndex        =   133
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ºC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   51
         Left            =   -62865
         TabIndex        =   132
         Top             =   2880
         Width           =   180
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x minuto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   -71880
         TabIndex        =   126
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Presión Arterial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   -74880
         TabIndex        =   125
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Frecuencia de Pulso"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   -74850
         TabIndex        =   124
         Top             =   2550
         Width           =   1425
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Respiraciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   -67170
         TabIndex        =   123
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Temperatura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   -67155
         TabIndex        =   122
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "mmHG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   -71880
         TabIndex        =   121
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "x minuto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   37
         Left            =   -64920
         TabIndex        =   120
         Top             =   2550
         Width           =   615
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ºC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   36
         Left            =   -64920
         TabIndex        =   119
         Top             =   2880
         Width           =   180
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Inicio de recolección de la orina"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   42
         Left            =   -74820
         TabIndex        =   114
         Top             =   2010
         Width           =   2850
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Notificación al Banco de Sangre"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   -69705
         TabIndex        =   112
         Top             =   1665
         Width           =   2865
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Notificación al Médico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   -74820
         TabIndex        =   111
         Top             =   1650
         Width           =   2850
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Término de Transfusión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   -69705
         TabIndex        =   108
         Top             =   1305
         Width           =   2865
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Inicio de Transfusión"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   -74820
         TabIndex        =   107
         Top             =   1290
         Width           =   2850
      End
      Begin VB.Shape BS000_01 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6180
         Index           =   2
         Left            =   -74925
         Top             =   360
         Width           =   12855
      End
      Begin VB.Shape BS000_01 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6180
         Index           =   0
         Left            =   75
         Top             =   360
         Width           =   12855
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Inicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   -72405
         TabIndex        =   74
         Top             =   570
         Width           =   990
      End
      Begin VB.Label BS000_03 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Hora de Término"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   -68640
         TabIndex        =   73
         Top             =   585
         Width           =   1185
      End
      Begin VB.Shape BS000_01 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6180
         Index           =   1
         Left            =   -74925
         Top             =   360
         Width           =   12855
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "BANCO DE SANGRE - CONDUCCIÓN DE TRANSFUSIONES"
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
      TabIndex        =   0
      Top             =   0
      Width           =   13365
   End
End
Attribute VB_Name = "ucBS_Tranfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para Transfusión
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

