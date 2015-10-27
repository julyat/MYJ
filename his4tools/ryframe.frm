VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ryframe 
   Caption         =   "入院病人违反唯一约束条件"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   9960
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1335
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   3240
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
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
   Begin RichTextLib.RichTextBox xingming 
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      _Version        =   393217
      TextRTF         =   $"ryframe.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton return 
      Caption         =   "返回主界面"
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "核对区"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "ryframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim bingrenxm As String
Dim rowid As String
Dim bingrenida As String
Dim bingrenidb As String
Private Sub DataGrid1_DblClick()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst3 = New ADODB.Recordset
rst3.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
cnn.Execute "update zy_bingrenxx set bingrenid = '" & bingrenidb & "' where rowid = '" & rowid & "'"
rst3.Open "select a.rowid,a.xingming as 姓名,a.bingrenid as 错误病人ID,a.yibaokh as 住院医保卡号,b.bingrenid as 正确病人ID,b.yibaokh as 对应医保卡号 from zy_bingrenxx a, gy_bingrenxx b where a.xingming = b.xingming and a.yibaokh = b.yibaokh and a.bingrenid <> b.bingrenid and a.xingming like '%" & bingrenxm & "%'", cnn
If rst3.EOF Or rst3.BOF Then
MsgBox "没有错误信息"
Set DataGrid1.DataSource = Nothing
DataGrid1.Visible = False
bingrenxm = ""
Else: Set DataGrid1.DataSource = rst3
bingrenida = rst1.Fields("错误病人ID")
bingrenidb = rst1.Fields("正确病人ID")
DataGrid1.Columns(0).Width = 0
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 1300
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1300
DataGrid1.Visible = True
DataGrid1.Refresh
End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
rowid = ""
DataGrid1.Col = 0
rowid = DataGrid1.Text
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
bingrenxm = ""
DataGrid2.Col = 0
bingrenxm = DataGrid2.Text
If bingrenxm = "" Then
bingrenxm = ""
DataGrid1.Visible = False
Set DataGrid1.DataSource = Nothing
Else: Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
rst1.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
rst1.Open "select a.rowid,a.xingming as 姓名,a.bingrenid as 错误病人ID,a.yibaokh as 住院医保卡号,b.bingrenid as 正确病人ID,b.yibaokh as 对应医保卡号 from zy_bingrenxx a, gy_bingrenxx b where a.xingming = b.xingming and a.yibaokh = b.yibaokh and a.bingrenid <> b.bingrenid and a.xingming like '%" & bingrenxm & "%'", cnn
If rst1.EOF Or rst1.BOF Then
MsgBox "没有错误信息"
Set DataGrid1.DataSource = Nothing
DataGrid1.Visible = False
bingrenxm = ""
Else: Set DataGrid1.DataSource = rst1
bingrenida = rst1.Fields("错误病人ID")
bingrenidb = rst1.Fields("正确病人ID")
DataGrid1.Columns(0).Width = 0
DataGrid1.Columns(1).Width = 1000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 1300
DataGrid1.Columns(4).Width = 1000
DataGrid1.Columns(5).Width = 1300
DataGrid1.Visible = True
DataGrid1.Refresh
End If
End If
End Sub

Private Sub Form_Load()
bingrenida = ""
bingrenidb = ""
bingrenxm = ""
rowid = ""
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub return_Click()
DataGrid1.Visible = False
DataGrid2.Visible = False
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
bingrenida = ""
bingrenidb = ""
bingrenxm = ""
rowid = ""
xingming.Text = ""
ryframe.Hide
mainframe.Show
End Sub

Private Sub xingming_Change()
DataGrid1.Visible = False
DataGrid2.Visible = False
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
bingrenida = ""
bingrenidb = ""
bingrenxm = ""
rowid = ""
If xingming.Text = "" Then
xingming.Text = ""
DataGrid1.Visible = False
DataGrid2.Visible = False
Else: Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst2 = New ADODB.Recordset
rst2.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
rst2.Open "select xingming,bingrenid,yibaokh,shenfenzh from gy_bingrenxx where xingming like '%" & xingming.Text & "%'", cnn
If rst2.EOF Or rst2.BOF Then
MsgBox "请核实病人姓名"
DataGrid1.Visible = False
DataGrid2.Visible = False
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
bingrenida = ""
bingrenidb = ""
bingrenxm = ""
rowid = ""
Else: Set DataGrid2.DataSource = rst2
DataGrid2.Columns(0).Width = 1000
DataGrid2.Columns(1).Width = 1000
DataGrid2.Columns(2).Width = 1300
DataGrid2.Columns(3).Width = 2000
DataGrid2.Visible = True
DataGrid2.Refresh
End If
End If
End Sub
