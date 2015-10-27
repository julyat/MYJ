VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form zhframe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "找回病程记录"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   8385
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   8040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin VB.CommandButton voidit 
      Caption         =   "作废病程记录"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   4560
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5106
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
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
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
      ColumnCount     =   5
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
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
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2459.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton findback 
      Caption         =   "找回病程记录"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton find 
      Caption         =   "查找病人信息"
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox binganhao 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc emr3 
      Height          =   615
      Left            =   1440
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      Connect         =   "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=his4rac;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=his4rac;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "emr3"
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
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "gy_yewudj"
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
      Left            =   3000
      TabIndex        =   12
      Top             =   7560
      Width           =   1935
   End
   Begin VB.Label dangqianbingren 
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
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "非作废区"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "作废区"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   6240
      Picture         =   "zhaohui.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "病案号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "zhframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bingren As String
Dim bingrenid As String
Dim zuofeiren As String
Dim zuofeibz As String
Dim rowid As String
Dim zfrowid As String
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim rst4 As ADODB.Recordset
Dim rst5 As ADODB.Recordset
Dim rst6 As ADODB.Recordset

Private Sub binganhao_Change()
If Len(binganhao) = 8 Then
find.Enabled = True
Else: find.Enabled = False
      findback.Enabled = False
      voidit.Enabled = False
End If
DataGrid1.Visible = False
DataGrid2.Visible = False
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGrid1.Col = 0
rowid = DataGrid1.Text
If rowid = "" Then
findback.Enabled = False
Else: findback.Enabled = True
End If

End Sub


Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
DataGrid2.Col = 0
zfrowid = DataGrid2.Text
If zfrowid = "" Then
voidit.Enabled = False
Else: voidit.Enabled = True
End If

End Sub



Private Sub DataGrid3_GotFocus()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
rst1.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=HIS4RAC;Persist Security Info=True"
rst1.Open "select * from gy_yewudj", cnn
If rst1.EOF Or rst1.BOF Then
MsgBox "无解锁信息"
binganhao.SetFocus
Else
Set DataGrid3.DataSource = rst1
DataGrid3.Refresh
Set rst2 = New ADODB.Recordset
rst2.Open "delete from gy_yewudj", cnn
Set DataGrid3.DataSource = Nothing
End If
End Sub

Private Sub find_Click()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset
rowid = ""
zfrowid = ""
findback.Enabled = False
voidit.Enabled = False
rst1.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=HIS4RAC;Persist Security Info=True"
rst1.Open "select bingrenid,bingrenxm from v_bl_bingrenjbxx where binganhao = '" & binganhao.Text & "'", cnn
If rst1.EOF Or rst1.BOF Then
MsgBox "未找到病人，请核对病案号"
findback.Enabled = False
voidit.Enabled = False
Else: bingren = rst1.Fields("bingrenxm")
bingrenid = rst1.Fields("bingrenid")
rst1.Close
rst1.Open "select rowid,jilumc,xiugaisj,zuofeiren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '1'", cnn, adOpenDynamic, adLockOptimistic
rst2.Open "select rowid,jilumc,xiugaisj,jiluren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '0'", cnn, adOpenDynamic, adLockOptimistic
If rst1.BOF Or rst1.EOF Then
MsgBox "该病人无任何作废病程，请核实"
Set DataGrid2.DataSource = rst2
DataGrid2.Refresh
Else: MsgBox "病人 " & bingren & " 所有作废病程记录已找到，请确认无误后找回病程记录"
dangqianbingren.Caption = bingren
Set DataGrid1.DataSource = rst1
DataGrid1.Columns(0).Width = 0
DataGrid1.Visible = True
Set DataGrid2.DataSource = rst2
DataGrid2.Columns(0).Width = 0
DataGrid2.Visible = True
DataGrid1.Refresh
DataGrid2.Refresh
End If
End If
End Sub

Private Sub findback_Click()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset
rst1.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=HIS4RAC;Persist Security Info=True"
cnn.Execute "update zy_doc_bingchengjl_v4 set zuofeibz = '0' where rowid = '" & rowid & "'"
MsgBox "修改成功"
rst1.Open "select rowid,jilumc,xiugaisj,zuofeiren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '1'", cnn, adOpenDynamic, adLockOptimistic
rst2.Open "select rowid,jilumc,xiugaisj,jiluren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '0'", cnn, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rst1
DataGrid1.Columns(0).Width = 0
Set DataGrid2.DataSource = rst2
DataGrid2.Columns(0).Width = 0
DataGrid1.Refresh
DataGrid2.Refresh
findback.Enabled = False
End Sub

Private Sub Form_Load()
bingren = ""
bingrenid = ""
zuofeiren = ""
zuofeibz = ""
dangqianbingren.Caption = ""
DataGrid1.Visible = False
DataGrid2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub return_Click()
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing
Set DataGrid3.DataSource = Nothing
binganhao.Text = ""
dangqianbingren.Caption = ""
zhframe.Hide
mainframe.Show
End Sub

Private Sub voidit_Click()
Dim cnn As ADODB.Connection
Set cnn = New ADODB.Connection
Set rst1 = New ADODB.Recordset
Set rst2 = New ADODB.Recordset
rst1.CursorLocation = adUseClient
rst2.CursorLocation = adUseClient
cnn.Open "Provider=MSDAORA.1;Password=emr3;User ID=emr3;Data Source=HIS4RAC;Persist Security Info=True"
cnn.Execute "update zy_doc_bingchengjl_v4 set zuofeibz = '1' where rowid = '" & rowid & "'"
MsgBox "修改成功"
rst1.Open "select rowid,jilumc,xiugaisj,zuofeiren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '1'", cnn, adOpenDynamic, adLockOptimistic
rst2.Open "select rowid,jilumc,xiugaisj,jiluren from zy_doc_bingchengjl_v4 where bingrenid = '" & bingrenid & "' and zuofeibz = '0'", cnn, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = rst1
DataGrid1.Columns(0).Width = 0
Set DataGrid2.DataSource = rst2
DataGrid2.Columns(0).Width = 0
DataGrid1.Refresh
DataGrid2.Refresh
voidit.Enabled = False
End Sub
