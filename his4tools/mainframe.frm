VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form mainframe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "his4tools"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   2805
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton rybr 
      Caption         =   "入院唯一约束"
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton scpbxx 
      Caption         =   "APP上传排班"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton ybkg 
      Caption         =   "医保开关"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton zhaohui 
      Caption         =   "找回病程记录"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton jiesuo 
      Caption         =   "解锁"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Shape ybkglight 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   360
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "版本号:20150213α"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "mainframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim kgzt As String

Private Sub Form_Load()
Dim inter1, inter2, inter3, inter4, inter5 As String
inter1 = canshu.get_ini(App.Path & "\canshu.ini", "interface", "jiesuo", 255)
inter2 = canshu.get_ini(App.Path & "\canshu.ini", "interface", "bingcheng", 255)
inter3 = canshu.get_ini(App.Path & "\canshu.ini", "interface", "weiyiyueshu", 255)
inter4 = canshu.get_ini(App.Path & "\canshu.ini", "interface", "yibao", 255)
inter5 = canshu.get_ini(App.Path & "\canshu.ini", "interface", "app", 255)
If inter1 <> "1" Then jiesuo.Visible = False
If inter2 <> "1" Then zhaohui.Visible = False
If inter3 <> "1" Then rybr.Visible = False
If inter4 <> "1" Then ybkglight.Visible = False
If inter4 <> "1" Then ybkg.Visible = False
If inter5 <> "1" Then scpbxx.Visible = False
kgzt = ""
ybkglight.FillColor = &HFF00&
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
rst1.Open "select canshuzhi from gy_canshu where yingyongid = '0401' and canshuid = '公用_是否启用医保医疗审核控制' ", cnn
kgzt = rst1.Fields("canshuzhi")
If kgzt = "1" Then
ybkglight.FillColor = &HFF00&
ybkg.Caption = "医保开关ON"
Else: ybkglight.FillColor = &HFF&
ybkg.Caption = "医保开关OFF"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub rybr_Click()
ryframe.Show
mainframe.Hide
End Sub

Private Sub scpbxx_Click()
Dim cnn As ADODB.Connection
Dim cmd As New ADODB.Command
Set cnn = New ADODB.Connection
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
cmd.ActiveConnection = cnn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PKG_GY_JOB.PRC_MZ_ZIDONGSCPBXX_HT"
cmd.Parameters.Append cmd.CreateParameter("@PRM_YUANQUID", adVarChar, adParamInput, 2, "3")
cmd.Execute

End Sub

Private Sub ybkg_Click()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
If kgzt = "1" Then
cnn.Execute "update gy_canshu set canshuzhi = '0' where canshuid = '公用_是否启用医保医疗审核控制' and yingyongid in ('0401','1001','1201','3701')"
kgzt = "0"
ybkg.Caption = "医保开关（OFF）"
ybkglight.FillColor = &HFF&
Else: cnn.Execute "update gy_canshu set canshuzhi = '1' where canshuid = '公用_是否启用医保医疗审核控制' and yingyongid in ('0401','1001','1201','3701')"
kgzt = "1"
ybkg.Caption = "医保开关（ON）"
ybkglight.FillColor = &HFF00&
End If
End Sub

Private Sub zhaohui_Click()
zhframe.Show
mainframe.Hide
End Sub

Private Sub jiesuo_Click()
jsframe.Show
mainframe.Hide
End Sub

