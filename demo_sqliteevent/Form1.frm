VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18765
   LinkTopic       =   "Form1"
   ScaleHeight     =   10785
   ScaleWidth      =   18765
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btn_addTestData 
      Caption         =   "添加测试数据"
      Height          =   360
      Left            =   15840
      TabIndex        =   28
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btn_delevent 
      Caption         =   "删除事件表记录"
      Height          =   360
      Left            =   12840
      TabIndex        =   27
      Top             =   5160
      Width           =   1815
   End
   Begin VB.CommandButton btn_refresh 
      Caption         =   "刷新显示"
      Height          =   360
      Left            =   10800
      TabIndex        =   22
      Top             =   5160
      Width           =   990
   End
   Begin VB.CommandButton btn_fail 
      Caption         =   "试验失败"
      Height          =   375
      Left            =   13920
      TabIndex        =   21
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton btn_finish 
      Caption         =   "试验完成"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton btn_intrrpt 
      Caption         =   "试验中断"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton btn_pause 
      Caption         =   "暂停继续"
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton btn_ji 
      Caption         =   "进退级"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton btn_yufinisd 
      Caption         =   "预实验完成"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton btn_loadfile 
      Caption         =   "下载所有参数文件"
      Height          =   375
      Left            =   16920
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton btn_start 
      Caption         =   "开始事件记录"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton btn_init 
      Caption         =   "数据库初始化"
      Height          =   375
      Left            =   11760
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid dg_abnormal 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   8400
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   3413
      _Version        =   393216
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
      Caption         =   "非正常结束试验"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.CheckBox ck_flag 
      Caption         =   "数据库记录开启"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dg_sjjlb 
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   3625
      _Version        =   393216
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
      Caption         =   "事件记录表"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSDataGridLib.DataGrid dg_syjlzb 
      Height          =   2295
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   4048
      _Version        =   393216
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
      Caption         =   "试验记录主表"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSDataGridLib.DataGrid dg_cs 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   3201
      _Version        =   393216
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
      Caption         =   "试验参数记录表"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label lb_time_fail 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   11160
      TabIndex        =   26
      Top             =   8040
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "失败记录耗时"
      Height          =   180
      Left            =   9840
      TabIndex        =   25
      Top             =   8040
      Width           =   1080
   End
   Begin VB.Label lb_time_end 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   6840
      TabIndex        =   24
      Top             =   8040
      Width           =   1650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lb_time_process 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   8040
      TabIndex        =   23
      Top             =   5160
      Width           =   1890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lb_time_self 
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   240
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "自检事件记录耗时:"
      Height          =   255
      Left            =   8160
      TabIndex        =   14
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lb_times_init 
      Height          =   255
      Left            =   15240
      TabIndex        =   13
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lb_time_start 
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "结束试验记录耗时"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   9
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "过程事件记录耗时"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "库初始化耗时(ms):"
      Height          =   255
      Left            =   13320
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "开始事件记录耗时:"
      Height          =   255
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsParameters As cRecordset

Private RsTests      As cRecordset

Private RsEvent      As cRecordset

Private Rsabnormal   As cRecordset

Private eventType    As Integer '事件类型  '枚举 0 开始型  1 过程型  2结束型  -1失败型  9初始化

Private testState    As Integer  '按钮互锁状态的Flag

Private maxStartTime As Single




Private Sub btn_addTestData_Click()
'添加测试数据

'循环 从1到1000,000



End Sub

Private Sub btn_delevent_Click()
测试删除事件表数据

Call delDataEvenTb
End Sub

'按钮互锁状态 两种互锁状态
'第一种  试验没开始 按钮状态:  开始试验 真 初始化数据库 真  自检完成 假  过程 假  结束 假 失败假  testState 0
'第二种  自检开始 按钮状态: 开始试验 假 自检完成 真 过程假 结束假  失败真 初始化数据库 假   testState 1
'第三种  试验开始 按钮状态: 开始试验假 自检 假  过程真 结束真  失败真  初始化数据库 假 testState 2

'试验失败 事件表和失败表
Private Sub btn_fail_Click()
    eventType = -1 '-1失败型
    testState = 0
Call Sqlite_Timing(True)
Call FailTest
lb_time_fail.Caption = Sqlite_Timing(False)

Call updateView


'lb_time_test.Caption = New_c.Timing

End Sub

'试验完成 事件表
Private Sub btn_finish_Click()
    eventType = 2 '2结束型
    testState = 0
Call Sqlite_Timing(True)

    Call EndTest("TestFinish")
    
    lb_time_end.Caption = Sqlite_Timing(False)
    
    Call updateView

End Sub

Private Sub btn_intrrpt_Click()
    eventType = 2 '2结束型
    testState = 0
    Call Sqlite_Timing(True)
    Call EndTest("TestInterrupt")
    
    lb_time_end.Caption = Sqlite_Timing(False)
    
    Call updateView


End Sub

'进级
Private Sub btn_ji_Click()
    eventType = 1 '1过程型
    testState = 2
Call Sqlite_Timing(True)
Call TestProcess("NextLevel")

lb_time_process.Caption = Sqlite_Timing(False)

  Call updateView

End Sub

'暂停继续
Private Sub btn_pause_Click()
    eventType = 1 '1过程型
    testState = 2

Call Sqlite_Timing(True)
Call TestProcess("Pause")

lb_time_process.Caption = Sqlite_Timing(False)

Call updateView
End Sub

'仅仅刷新界面的按钮
Private Sub btn_refresh_Click()
    eventType = 9

    update_dg_cs
    update_dg_syjlzb
    update_dg_sjjlb
    update_dg_abnormal

End Sub

'预实验完成 事件
Private Sub btn_yufinisd_Click()
eventType = 1
testState = 2

Call Sqlite_Timing(True)
Call Selfcheckfinish
lb_time_self.Caption = Sqlite_Timing(False)


Call updateView
End Sub

'开始记录数据 初始化 获取MD5 添加 两个表
Private Sub btn_start_Click()
eventType = 0 '开始'
testState = 1

Call Sqlite_Timing(True)

Call StartTest

lb_time_start.Caption = Sqlite_Timing(False)


Call updateView '刷新表格页面

End Sub



Private Sub StartTest()

    Dim data As StartEvent_Data

    data.file_path = GetFilePath()
    data.ampcode = randomInt(1111, 9999)
    data.centring = randomInt(100)
    data.event_name = "StartTest_Maunl"
    data.field_curr = randomSingle(1000)
    data.field_volt = randomSingle(1000)
    data.gain = randomInt(100)
    data.hardcode = randomInt(1111, 9999)
    data.max_acce = randomSingle(100)
    data.max_disp = randomSingle(100)
    data.max_velo = randomSingle(100)
    data.test_name = GetTestName()
    data.wind_pressure = randomSingle(100)
    
    Call startTestRecordData(data)

End Sub


Private Sub Selfcheckfinish()
Dim data As ProcessEvent_Data

data.armature_curr = randomSingle(100)
data.armature_volt = randomSingle(100)
data.control_value = randomSingle(100)
data.driving_volt = randomSingle(100)
data.event_name = "SelfCheck"
data.field_curr = randomSingle(100)
data.field_volt = randomSingle(100)




Call selfCheckRecordData(data)


End Sub


Private Sub TestProcess(name As String)
Dim data As ProcessEvent_Data
data.armature_curr = randomSingle(100)
data.armature_volt = randomSingle(100)
data.control_value = randomSingle(100)
data.driving_volt = randomSingle(100)
data.event_name = name
data.field_curr = randomSingle(100)
data.field_volt = randomSingle(100)
data.random_time = randomInt(100)
data.shock_num = randomInt(1000)
data.sine_freq = randomInt(2000)

Call processRecordData(data)

End Sub


Private Sub EndTest(name As String)
Dim data As ProcessEvent_Data

data.armature_curr = randomSingle(100)
data.armature_volt = randomSingle(100)
data.control_value = randomSingle(100)
data.driving_volt = randomSingle(100)
data.event_name = name
data.field_curr = randomSingle(100)
data.field_volt = randomSingle(100)
data.random_time = randomInt(100)
data.shock_num = randomInt(1000)
data.sine_freq = randomInt(2000)

Call endTestRecordData(data)

End Sub


Private Sub FailTest()


Dim evendata As ProcessEvent_Data
Dim faildata As FailEvent_Data

evendata.armature_curr = randomSingle(100)
evendata.armature_volt = randomSingle(100)
evendata.control_value = randomSingle(100)
evendata.driving_volt = randomSingle(100)
evendata.event_name = "TestFail"
evendata.field_curr = randomSingle(100)
evendata.field_volt = randomSingle(100)
evendata.random_time = randomInt(100)
evendata.shock_num = randomInt(1000)
evendata.sine_freq = randomInt(2000)

faildata.reason = GetFailName()
faildata.linkage = GetFileFromdb()



Call failTestRecordData(evendata, faildata)
End Sub



Private Sub updateView()

    Dim count As Long
    
    '''''''''''''''''''''按钮互锁
    
    Select Case testState
    
        Case 0
            btn_start.Enabled = True
            btn_init.Enabled = True
            btn_yufinisd.Enabled = False
            btn_ji.Enabled = False
            btn_pause.Enabled = False
            btn_intrrpt.Enabled = False
            btn_finish.Enabled = False
            btn_fail.Enabled = False
        
        Case 1
        
            btn_start.Enabled = False
            btn_init.Enabled = False
            btn_yufinisd.Enabled = True
            btn_ji.Enabled = False
            btn_pause.Enabled = False
            btn_intrrpt.Enabled = True
            btn_finish.Enabled = False
            btn_fail.Enabled = True
        
        Case 2
            btn_start.Enabled = False
            btn_init.Enabled = False
            btn_yufinisd.Enabled = False
            btn_ji.Enabled = True
            btn_pause.Enabled = True
            btn_intrrpt.Enabled = True
            btn_finish.Enabled = True
            btn_fail.Enabled = True
    
    End Select

    '''''''''''参数记录''''''''''

    Select Case eventType

        Case 0
            '开始事件
            update_dg_cs
            update_dg_syjlzb
            update_dg_sjjlb
    
        Case 1
            '过程事件

            update_dg_sjjlb

        Case 2
            '结束事件
            update_dg_syjlzb
            update_dg_sjjlb
    
        Case -1
            '失败事件

            'update_dg_cs
            update_dg_syjlzb

            update_dg_sjjlb

            update_dg_abnormal

        Case 9
            update_dg_cs
            update_dg_syjlzb
            update_dg_sjjlb
            update_dg_abnormal

    End Select

End Sub

'参数
Private Sub update_dg_cs()

    If eventType = 9 Then

        Set RsParameters = getDataRecordset("ParameterTable")

        If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
        Set dg_cs.DataSource = RsParameters.DataSource
    Else

        RsParameters.ReQuery

    End If

    dg_cs.Columns(0).Width = 500
    dg_cs.Columns(1).Width = 3000

End Sub

'试验
Private Sub update_dg_syjlzb()

    If eventType = 9 Then
        Set RsTests = getDataRecordset("TesTable")

        If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
        Set dg_syjlzb.DataSource = RsTests.DataSource
    
    Else
        RsTests.ReQuery

    End If
    
    dg_syjlzb.Scroll 0, RsTests.RecordCount - 1
    dg_syjlzb.Columns(0).Width = 500
    dg_syjlzb.Columns(1).Width = 1000
    dg_syjlzb.Columns(2).Width = 3000

End Sub

'事件
Private Sub update_dg_sjjlb()

    If eventType = 9 Then
        Set RsEvent = getDataRecordset("EvenTable")

        If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
        Set dg_sjjlb.DataSource = RsEvent.DataSource
    
    Else
        RsEvent.ReQuery

    End If
    
    dg_sjjlb.Scroll 0, RsEvent.RecordCount - 1
    dg_sjjlb.Columns(0).Width = 500
    dg_sjjlb.Columns(1).Width = 700
    dg_sjjlb.Columns(2).Width = 1000
    dg_sjjlb.Columns(3).Width = 1200
    dg_sjjlb.Columns(4).Width = 1200
    dg_sjjlb.Columns(5).Width = 1200
        
    dg_sjjlb.Columns(6).Width = 1000
    
    dg_sjjlb.Columns(7).Width = 1300
    dg_sjjlb.Columns(8).Width = 1300
    dg_sjjlb.Columns(9).Width = 1300
    dg_sjjlb.Columns(10).Width = 1300
    dg_sjjlb.Columns(11).Width = 1300
    
    dg_sjjlb.Columns(12).Width = 800
    dg_sjjlb.Columns(13).Width = 900
    dg_sjjlb.Columns(14).Width = 1200
    
End Sub

'失败
Private Sub update_dg_abnormal()

    If eventType = 9 Then
        Set Rsabnormal = getDataRecordset("FailTable")

        If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
        Set dg_abnormal.DataSource = Rsabnormal.DataSource
    Else
        Rsabnormal.ReQuery

    End If
    
    dg_abnormal.Scroll 0, Rsabnormal.RecordCount - 1
    dg_abnormal.Columns(0).Width = 500
    dg_abnormal.Columns(1).Width = 800
End Sub




Private Sub Form_Load()

    'Dim s As New clsMD5
    Dim result As String
   
    
    If New_c.FSO.FileExists(DBFileName) Then
        If MsgBox("Database exists. Delete it and start fresh?", vbYesNo) = vbYes Then
            New_c.FSO.DeleteFile DBFileName
        End If
    End If
  
    If Not initDataBase() Then Exit Sub
    testState = 0
    eventType = 9
    updateView
reflashCaption
End Sub


Private Sub setdbnull() '清空显示列表,关闭数据集和连接

    Set dg_syjlzb.DataSource = Nothing
    Set dg_cs.DataSource = Nothing
    Set dg_sjjlb.DataSource = Nothing
    Set dg_abnormal.DataSource = Nothing

    Set RsParameters = Nothing
    Set RsTests = Nothing
    Set RsEvent = Nothing
    Set Rsabnormal = Nothing
    
End Sub

Private Sub initdb()

    Dim flag_init As Boolean
 
    If New_c.FSO.FileExists(DBFileName) Then
        If MsgBox("是否删除当前数据库文件重新初始化?", vbYesNo) = vbYes Then
    
            setdbnull
    
            New_c.FSO.DeleteFile DBFileName
            New_c.Timing True '开始记录 数据库初始化
            flag_init = initDataBase()
            lb_times_init.Caption = New_c.Timing
      
            If (flag_init) Then
      
                MsgBox ("数据库初始完成")
      
            Else
      
                MsgBox ("数据库初始失败")
       
            End If
      
            New_c.Timing False
      
        End If
    End If

End Sub

Private Sub btn_init_Click()
    initdb
End Sub


Private Sub loadFile()

    Dim i             As Integer

    Dim max           As Integer

    Dim filename_temp As String

    Dim flag_temp     As Boolean

    Dim md5codefile   As String

    Dim md5codedb     As String

   

    max = RsParameters.RecordCount - 1
    RsParameters.MoveFirst
    



    flag_temp = False

    For i = 0 To max Step 1

        filename_temp = "E:\SqliteObject\demo_sqliteevent\sss\" & RsParameters.Fields("file_path").value

        New_c.FSO.WriteByteContent filename_temp, RsParameters.Fields("file").value

        md5codedb = RsParameters.Fields("md5").value
        md5codefile = GetFileMD5(filename_temp)

        If (md5codedb = md5codefile) Then '验证文件MD5码
            flag_temp = True
        Else

            flag_temp = False
        End If

        RsParameters.MoveNext '坐标移到下一个 如果到底 直接设置结束循环

        If (RsParameters.EOF) Then
            i = max + 1
        End If

    Next i

    RsParameters.MoveFirst

    If flag_temp Then
        MsgBox "文件导出完成,并完成校验"
    Else
        MsgBox "文件导出完成,校验失败"
    End If

End Sub

Private Sub btn_loadfile_Click()

    loadFile

End Sub

Private Sub reflashCaption(Optional content As String = "")

    If Not content = "" Then
        content = "[ " & content & " ]"
    End If

    Form1.Caption = "数据库测试DEMO  " & content

End Sub

