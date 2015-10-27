VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form jsframe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "解锁"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9930
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox xingming 
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"unlocker.frx":0000
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2415
      Left            =   2400
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2295
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4048
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
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1720
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
   Begin MSAdodcLib.Adodc his3 
      Height          =   495
      Left            =   3720
      Top             =   6000
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
   Begin VB.CommandButton return 
      Caption         =   "返回主界面"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "双击解锁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "unlocker.frx":009D
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2280
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "mz_chufang1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "mz_yiji1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label labelbingrenid 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label labelxingming 
      Alignment       =   2  'Center
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
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   1215
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
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "jsframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bingrenid As String
Dim bingrenxm As String
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim ztbz2 As String
Dim ztbz3 As String
Dim rowid2 As String
Dim rowid3 As String


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGrid1.Col = 1
bingrenxm = DataGrid1.Text
DataGrid1.Col = 0
bingrenid = DataGrid1.Text
labelxingming.Caption = bingrenxm
labelbingrenid.Caption = bingrenid
ztbz2 = ""
ztbz3 = ""
Set DataGrid2.DataSource = Nothing
Set DataGrid3.DataSource = Nothing
If bingrenid = "" Then DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
Label2.Visible = False
Label3.Visible = False
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst2 = New ADODB.Recordset
Set rst3 = New ADODB.Recordset
rst2.CursorLocation = adUseClient
rst3.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
rst2.Open "select rowid,bingrenxm,xiugairen,xiugaisj,kaidanysxm from mz_yiji1 where xiugaibz <> '0' and bingrenid = '" & bingrenid & "'", cnn, adOpenDynamic, adLockOptimistic
rst3.Open "select rowid,bingrenxm,xiugairen,xiugaisj,kaidanysxm from mz_chufang1 where xiugaibz <> '0' and bingrenid = '" & bingrenid & "'", cnn, adOpenDynamic, adLockOptimistic
If rst2.EOF Or rst2.BOF Then
ztbz2 = "0"
Else: ztbz2 = "1"
End If
If rst3.BOF Or rst3.EOF Then
ztbz3 = "0"
Else: ztbz3 = "1"
End If
If ztbz2 = "0" And ztbz3 = "0" Then
MsgBox "未找到该病人任何纪录"
DataGrid2.Visible = False
DataGrid3.Visible = False
Label2.Visible = False
Label3.Visible = False
ElseIf ztbz2 = "1" And ztbz3 = "0" Then
Set DataGrid2.DataSource = rst2
DataGrid2.Visible = True
Label2.Visible = True
DataGrid3.Visible = False
Label3.Visible = False
DataGrid2.Refresh
ElseIf ztbz2 = "0" And ztbz3 = "1" Then
Set DataGrid3.DataSource = rst3
DataGrid2.Visible = False
Label2.Visible = False
DataGrid3.Visible = True
Label3.Visible = True
DataGrid3.Refresh
Else: Set DataGrid2.DataSource = rst2
Set DataGrid3.DataSource = rst3
DataGrid2.Visible = True
Label2.Visible = True
DataGrid3.Visible = True
Label3.Visible = True
DataGrid2.Refresh
DataGrid3.Refresh
End If
DataGrid2.Columns(0).Width = 0
DataGrid2.Columns(1).Width = 1000
DataGrid3.Columns(0).Width = 0
DataGrid3.Columns(1).Width = 1000
End Sub

Private Sub DataGrid2_DblClick()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst2 = New ADODB.Recordset
rst2.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
cnn.Execute "update mz_yiji1 set xiugaibz = '0' where rowid = '" & rowid2 & "'"
rst2.Open "select rowid,bingrenxm,xiugairen,xiugaisj,kaidanysxm from mz_yiji1 where xiugaibz <> '0' and bingrenid = '" & bingrenid & "'", cnn, adOpenDynamic, adLockOptimistic
If rst2.BOF Or rst2.EOF Then
DataGrid2.Visible = False
Label2.Visible = False
Else: Set DataGrid2.DataSource = rst2
DataGrid2.Visible = True
Label2.Visible = True
DataGrid2.Refresh
End If
DataGrid2.Columns(0).Width = 0
DataGrid2.Columns(1).Width = 1000
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
rowid2 = ""
DataGrid2.Col = 0
rowid2 = DataGrid2.Text
End Sub

Private Sub DataGrid3_DblClick()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst3 = New ADODB.Recordset
rst3.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
cnn.Execute "update mz_chufang1 set xiugaibz = '0' where rowid = '" & rowid3 & "'"
rst3.Open "select rowid,bingrenxm,xiugairen,xiugaisj,kaidanysxm from mz_chufang1 where xiugaibz <> '0' and bingrenid = '" & bingrenid & "'", cnn, adOpenDynamic, adLockOptimistic
If rst3.BOF Or rst2.EOF Then
DataGrid3.Visible = False
Label3.Visible = False
Else: Set DataGrid3.DataSource = rst3
DataGrid3.Visible = True
Label3.Visible = True
DataGrid3.Refresh
End If
DataGrid3.Columns(0).Width = 0
DataGrid3.Columns(1).Width = 1000
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
rowid3 = ""
DataGrid3.Col = 0
rowid3 = DataGrid3.Text
End Sub

Private Sub Form_Load()
bingrenid = ""
bingrenxm = ""
rowid2 = ""
rowid3 = ""
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
Label2.Visible = False
Label3.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub return_Click()
bingrenid = ""
bingrenxm = ""
rowid2 = ""
rowid3 = ""
xingming.Text = ""
labelxingming.Caption = ""
labelbingrenid.Caption = ""
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
Set DataGrid3.DataSource = Nothing
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
Label2.Visible = False
Label3.Visible = False
jsframe.Hide
mainframe.Show
End Sub

Private Sub xingming_Change()
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
Set DataGrid3.DataSource = Nothing
labelxingming.Caption = ""
labelbingrenid.Caption = ""
rowid2 = ""
rowid3 = ""
bingrenid = ""
bingrenxm = ""
ztbz2 = "0"
ztbz3 = "0"
DataGrid2.Visible = False
DataGrid3.Visible = False
Label2.Visible = False
Label3.Visible = False
If xingming.Text = "" Then
xingming.Text = ""
DataGrid1.Visible = False
Else: Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
rst1.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=his3;User ID=his3;Data Source=his4rac;Persist Security Info=True"
rst1.Open "select bingrenid,xingming,jiuzhenkh,xingbie from gy_bingrenxx where xingming like '%" & xingming.Text & "%'", cnn, adOpenDynamic, adLockOptimistic
If rst1.EOF Or rst1.BOF Then
DataGrid1.Visible = False
DataGrid2.Visible = False
DataGrid3.Visible = False
MsgBox "" & xingming.Text & "未找到"
Else: Set DataGrid1.DataSource = rst1
DataGrid1.Visible = True
DataGrid1.Refresh
End If
End If
End Sub

