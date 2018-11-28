Attribute VB_Name = "SqliteDB_ODBC"
Option Explicit

Public conDB As cConnection

'结构体中字符串类数据长度
Public Const Event_name_len = 10
Public Const Param_test_type_len = 10
Public Const Test_md5_len = 32
Public Const Fail_reason_len = 10
Public Const EquipConfig_model_len = 10
Public Const Param_file_path_len = 10

'公共的试验ID 是判断当前试验记录是否是在同一次试验的依据
'当调用开启试验记录函数 会保存 试验记录ID值
'后续记录过程会再通过这个ID值

Private md5code_module      As String '公共数据md5码
Private testMastcode_module As Long '试验记录表ID
Private eventcode_module    As Long '事件记录表ID



Public Type Testb_Data
md5 As String * Test_md5_len
gain As Byte
hardcode As Integer
ampcode As Integer
End Type


Public Type Eventb_Data
test_id As Long
event_name As String * Event_name_len
field_volt As Single
field_curr As Single
centring As Integer
wind_pressure As Single
control_value As Single
driving_volt As Single
armature_volt As Single
armature_curr As Single
sine_freq As Integer
shock_num As Integer
random_time As Integer

End Type


Public Type Paramtb_Data
max_acce As Single
max_velo As Single
max_disp As Single
file_path As String * Param_file_path_len
test_name As String * Param_test_type_len

End Type


Public Type Failtb_Data
fail_test_id As Long
reason As String * Fail_reason_len
linkage As String * 10
profile() As Byte
drive() As Byte

End Type



Public Type EquipConfigtb_Data

table_no As Byte '台体编号
table_model As String * EquipConfig_model_len '设备型号
table_date As Date '出厂日期
serial_no As String * EquipConfig_model_len '设备序列号
max_sine_thrust As Single '最大正弦推力
max_random_thrust As Single '最大随机推力
max_shock_thrust As Single '最大冲击推力
max_acc As Single '最大加速度
max_conti_disp As Single '连续位移
max_transi_disp As Single '瞬态位移
max_load_bear As Single '最大承载
compon_quality As Single '组件总质量
hori_slid_table As Byte '水平滑台数量
verti_expan As Byte '垂直扩展台数量
hori_slid_model As String * EquipConfig_model_len '水平滑台型号
hori_bearing_amount As Byte '水平滑台轴承数量
hori_slid_quality As Single '水平滑台总质量
verti_expan_model As String * EquipConfig_model_len '垂直扩展台型号
auxi_support As Byte '辅助支撑数量
Auxi_guide As Byte '辅助导向数量
verti_expan_quality As Single '扩展台质量
anti_resonance_freq As Single '一阶反共振频率
blower_model As String * EquipConfig_model_len '风机型号
blower_Manufactor As String * EquipConfig_model_len '风机厂家
blower_date As Date '风机出厂日期

End Type




'''''封装到实际调用时传递过来的数据结构体,调用时只需要确定数据,隐藏无关信息'''''

Public Type StartEvent_Data
'14项数据
'开始事件 需要 信息
'当前试验参数文件 目录
'试验限值 最大加速度 速度 位移
'试验名称
'功放增益
'控制器硬件号
'开始事件名称 自动 手动
'励磁电压电流 动圈对中位置 风机风压

max_acce As Single
max_velo As Single
max_disp As Single
file_path As String
test_name As String
gain As Byte

hardcode As Integer
ampcode As Integer

event_name As String

field_volt As Single
field_curr As Single
centring As Integer
wind_pressure As Single

End Type


Public Type ProcessEvent_Data
'自检事件 需要 信息
'事件发生前一刻的
'振动控制值,驱动电压,
'功放励磁电压 , 电流
'动圈电压电流

'过程型事件 和结束型事件 需要 信息
'在自检型事件的基础上还需要
'随机试验已进行时长
'冲击试验冲击次数
'正弦试验当前频率

event_name As String
field_volt As Single
field_curr As Single
control_value As Single
driving_volt As Single
armature_volt As Single
armature_curr As Single

sine_freq As Integer
shock_num As Integer
random_time As Integer
End Type


'失败额外记录 故障原因,功放连锁状态,参考谱,驱动谱
Public Type FailEvent_Data
reason As String
linkage As String
profile(10) As Byte
drive(10) As Byte

End Type
'''''''''''''''''''''''''''''''''''''


Public Function DBFileName() As String
    DBFileName = App.Path & "\testdemodatabase.db3"
End Function


Public Function initDataBase() As Boolean
    Dim filePath  As String
    
    
    On Error GoTo ExitFalse 'Return False if operation fails.
    filePath = DBFileName()
    If New_c.FSO.FileExists(filePath) Then 'normally this is the case
        Set conDB = New_c.Connection(DBFileName, DBOpenFromFile)
  
    Else 'create a new DB, a new Table + a persistent Insert-Command - and then populate the new table with Data
        Set conDB = New_c.Connection(filePath, DBCreateNewFileDB)
        conDB.Execute "Create Table TesTable (_id Integer PRIMARY KEY AUTOINCREMENT, time_stamp INTEGER NOT NULL,md5 TEXT , s_event_id INTEGER,e_event_id INTEGER,gain INTEGER NOT NULL,hardcode INTEGER NOT NULL,ampcode INTEGER NOT NULL)"
        conDB.Execute "Create Table EvenTable (_id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL ,test_id INTEGER ,time_stamp INTEGER NOT NULL,event_name TEXT ,field_volt REAL ,field_curr REAL ,centring INTEGER ,wind_pressure REAL ,control_value REAL ,driving_volt REAL ,armature_volt REAL ,armature_curr REAL, sine_freq INTEGER ,shock_num INTEGER ,random_value INTEGER)"
        conDB.Execute "Create Table ParameterTable (_id INTEGER PRIMARY KEY AUTOINCREMENT ,md5 TEXT ,file BLOB,max_acce REAL ,max_velo REAL ,max_disp REAL,file_path TEXT,test_type TEXT)"
        conDB.Execute "Create Table FailTable (_id INTEGER PRIMARY KEY AUTOINCREMENT , test_id INTEGER , reason TEXT ,linkage TEXT ,profile BLOB ,drive BLOB)"
        
        
        
        
        If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    End If
  
    initDataBase = True

ExitFalse:

End Function


Public Function insertDataParamTb(dbdata As Paramtb_Data) As String
'向参数记录表添加数据
'参数 参数表的记录数据 其中 文件路径先按绝对路径设置
'返回值 添加条目MD5主键,失败返回""

'策略 先按文件路径查询MD5码 然后向记录表查询,如果没有 添加记录 最终返回这个MD5码

''''声明局部变量
    Dim Cmd As cCommand
    Dim md5code As String * 32
    
''''实现体
    md5code = GetFileMD5(dbdata.file_path)
    
    If Not (QueryMd5ParamTb(md5code)) Then
    
    Set Cmd = conDB.CreateCommand("Insert Into ParameterTable(_id,md5,file,max_acce,max_velo,max_disp,file_path,test_type) Values(?,?,?,?,?,?,?,?)")
    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetText 2, md5code  'md5
    Cmd.SetBlob 3, New_c.FSO.ReadByteContent(dbdata.file_path)  'file
    Cmd.SetDouble 4, dbdata.max_acce  'max_acce
    Cmd.SetDouble 5, dbdata.max_velo  'max_velo
    Cmd.SetDouble 6, dbdata.max_disp   'max_disp
    Cmd.SetText 7, Right$(dbdata.file_path, 5) 'file_path
    Cmd.SetText 8, dbdata.test_name
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True
    
    End If
    insertDataParamTb = md5code

End Function



Public Function insertDataParamTb_StartEvent(ByRef dbdata As StartEvent_Data) As String
'向参数记录表添加数据
'参数 参数表的记录数据 其中 文件路径先按绝对路径设置
'返回值 添加条目MD5主键,失败返回""

'策略 先按文件路径查询MD5码 然后向记录表查询,如果没有 添加记录 最终返回这个MD5码

''''声明局部变量
    Dim Cmd As cCommand
    
''''实现体
    md5code_module = GetFileMD5(dbdata.file_path)

    If Not (QueryMd5ParamTb(md5code_module)) Then
    
    Set Cmd = conDB.CreateCommand("Insert Into ParameterTable(_id,md5,file,max_acce,max_velo,max_disp,file_path,test_type) Values(?,?,?,?,?,?,?,?)")
    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetText 2, md5code_module  'md5
    Cmd.SetBlob 3, New_c.FSO.ReadByteContent(dbdata.file_path)  'file
    Cmd.SetDouble 4, dbdata.max_acce   'max_acce
    Cmd.SetDouble 5, dbdata.max_velo   'max_velo
    Cmd.SetDouble 6, dbdata.max_disp    'max_disp
    Cmd.SetText 7, Right$(dbdata.file_path, 5) 'file_path
    Cmd.SetText 8, dbdata.test_name
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True
    
    End If
    insertDataParamTb_StartEvent = md5code_module

End Function



Public Function insertDataTesTb(dbdata As Testb_Data) As Integer
'向试验记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand


    Set Cmd = conDB.CreateCommand("Insert Into TesTable(_id,time_stamp,md5,s_event_id,e_event_id,gain,hardcode) Values(?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans

    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, getNowTimestamp()  'time_stamp
    Cmd.SetText 3, md5code_module 'md5
    Cmd.SetNull 4   's_event_id
    Cmd.SetNull 5   'e_event_id
    Cmd.SetInt32 6, dbdata.gain   'gain
    Cmd.SetInt32 7, dbdata.hardcode   'hardcode

    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True
    
    
    insertDataTesTb = conDB.LastInsertAutoID

End Function


Public Function insertDataTesTb_StartEvent(ByRef dbdata As StartEvent_Data) As Integer
'向试验记录表添加数据 单不封装参数

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand

    Set Cmd = conDB.CreateCommand("Insert Into TesTable(_id,time_stamp,md5,s_event_id,e_event_id,gain,hardcode,ampcode) Values(?,?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans

    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, getNowTimestamp()  'time_stamp
    Cmd.SetText 3, md5code_module 'md5
    Cmd.SetNull 4   's_event_id
    Cmd.SetNull 5   'e_event_id
    Cmd.SetInt32 6, dbdata.gain   'gain
    Cmd.SetInt32 7, dbdata.hardcode 'hardcode
    Cmd.SetInt32 8, dbdata.ampcode 'ampcode
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True
    testMastcode_module = conDB.LastInsertAutoID
    insertDataTesTb_StartEvent = testMastcode_module
    
End Function



Public Function insertDataEvenTb(dbdata As Eventb_Data) As Integer
'向事件记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand
    
    Set Cmd = conDB.CreateCommand("Insert Into EvenTable(_id,test_id,time_stamp,event_name,field_volt,field_curr,centring,wind_pressure,control_value,driving_volt,armature_volt,armature_curr,sine_freq,shock_num,random_value) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, dbdata.test_id   'test_id
    Cmd.SetInt32 3, getNowTimestamp() 'time_stamp
    Cmd.SetText 4, dbdata.event_name   'event_name
    Cmd.SetDouble 5, dbdata.field_volt   'field_volt
    Cmd.SetDouble 6, dbdata.field_curr   'field_curr
    Cmd.SetInt32 7, dbdata.centring   'centring
    Cmd.SetDouble 8, dbdata.wind_pressure   'wind_pressure
    Cmd.SetDouble 9, dbdata.control_value   'control_value
    Cmd.SetDouble 10, dbdata.driving_volt  'driving_volt
    Cmd.SetDouble 11, dbdata.armature_volt   'armature_volt
    Cmd.SetDouble 12, dbdata.armature_curr   'armature_curr
    Cmd.SetInt32 13, dbdata.sine_freq  'sine_freq
    Cmd.SetInt32 14, dbdata.shock_num   'shock_num
    Cmd.SetInt32 15, dbdata.random_time  'random_time
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True

    insertDataEvenTb = conDB.LastInsertAutoID

End Function


Public Function insertDataStartEvenTb(dbdata As StartEvent_Data) As Integer
'向事件记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand
    
    Set Cmd = conDB.CreateCommand("Insert Into EvenTable(_id,test_id,time_stamp,event_name,field_volt,field_curr,centring,wind_pressure,control_value,driving_volt,armature_volt,armature_curr,sine_freq,shock_num,random_value) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, testMastcode_module   'test_id
    Cmd.SetInt32 3, getNowTimestamp() 'time_stamp
    Cmd.SetText 4, dbdata.event_name    'event_name
    Cmd.SetDouble 5, dbdata.field_volt   'field_volt
    Cmd.SetDouble 6, dbdata.field_curr   'field_curr
    Cmd.SetInt32 7, dbdata.centring 'centring
    Cmd.SetDouble 8, dbdata.wind_pressure   'wind_pressure
    Cmd.SetNull 9    'control_value
    Cmd.SetNull 10   'driving_volt
    Cmd.SetNull 11    'armature_volt
    Cmd.SetNull 12    'armature_curr
    Cmd.SetNull 13  'sine_freq
    Cmd.SetNull 14   'shock_num
    Cmd.SetNull 15  'random_time
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True

    insertDataStartEvenTb = conDB.LastInsertAutoID

End Function



Public Function insertDataSelfCheckEvenTb(dbdata As ProcessEvent_Data) As Long
'向事件记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand
    
    Set Cmd = conDB.CreateCommand("Insert Into EvenTable(_id,test_id,time_stamp,event_name,field_volt,field_curr,centring,wind_pressure,control_value,driving_volt,armature_volt,armature_curr,sine_freq,shock_num,random_value) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, testMastcode_module   'test_id
    Cmd.SetInt32 3, getNowTimestamp() 'time_stamp
    Cmd.SetText 4, dbdata.event_name    'event_name
    Cmd.SetDouble 5, dbdata.field_volt   'field_volt
    Cmd.SetDouble 6, dbdata.field_curr   'field_curr
    Cmd.SetNull 7   'centring
    Cmd.SetNull 8   'wind_pressure
    Cmd.SetDouble 9, dbdata.control_value    'control_value
    Cmd.SetDouble 10, dbdata.driving_volt   'driving_volt
    Cmd.SetDouble 11, dbdata.armature_volt    'armature_volt
    Cmd.SetDouble 12, dbdata.armature_curr    'armature_curr
    Cmd.SetNull 13  'sine_freq
    Cmd.SetNull 14   'shock_num
    Cmd.SetNull 15  'random_time
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True

    insertDataSelfCheckEvenTb = conDB.LastInsertAutoID

End Function


Public Function insertDataProcessEvenTb(dbdata As ProcessEvent_Data) As Long
'向事件记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1

    Dim Cmd As cCommand
    
    Set Cmd = conDB.CreateCommand("Insert Into EvenTable(_id,test_id,time_stamp,event_name,field_volt,field_curr,centring,wind_pressure,control_value,driving_volt,armature_volt,armature_curr,sine_freq,shock_num,random_value) Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans
    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, testMastcode_module   'test_id
    Cmd.SetInt32 3, getNowTimestamp() 'time_stamp
    Cmd.SetText 4, dbdata.event_name    'event_name
    Cmd.SetDouble 5, dbdata.field_volt   'field_volt
    Cmd.SetDouble 6, dbdata.field_curr   'field_curr
    Cmd.SetNull 7   'centring
    Cmd.SetNull 8   'wind_pressure
    Cmd.SetDouble 9, dbdata.control_value    'control_value
    Cmd.SetDouble 10, dbdata.driving_volt   'driving_volt
    Cmd.SetDouble 11, dbdata.armature_volt    'armature_volt
    Cmd.SetDouble 12, dbdata.armature_curr    'armature_curr
    Cmd.SetInt32 13, dbdata.sine_freq  'sine_freq
    Cmd.SetInt32 14, dbdata.shock_num   'shock_num
    Cmd.SetInt32 15, dbdata.random_time  'random_time
    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True

    insertDataProcessEvenTb = conDB.LastInsertAutoID

End Function





'Public Function insertDataFailTesTb(dbdata As Failtb_Data)
''向失败记录表添加数据
'
''返回值 添加条目的自增ID值,失败返回-1
'    Dim Cmd As cCommand
'
'    Set Cmd = conDB.CreateCommand("Insert Into FailTable(_id,test_id,reason,linkage,profile,drive) Values(?,?,?,?,?,?)")
'
'    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
'    conDB.Synchronous = False
'    conDB.BeginTrans
'
'    Cmd.SetNull 1   'id
'    Cmd.SetInt32 2, dbdata.test_id   'test_id
'    Cmd.SetText 3, dbdata.reason   'reason
'    Cmd.SetText 4, dbdata.linkage   'linkage
'    Cmd.SetNull 5   'profile
'    Cmd.SetNull 6  'drive
'
'    Cmd.Execute
'    conDB.CommitTrans
'    conDB.Synchronous = True
'
'
''insertDataFailTesTb = conDB.LastInsertAutoID
'
'
'
'End Function


Public Function insertDataFailTesTb(dbdata As FailEvent_Data)
'向失败记录表添加数据

'返回值 添加条目的自增ID值,失败返回-1
    Dim Cmd As cCommand

    Set Cmd = conDB.CreateCommand("Insert Into FailTable(_id,test_id,reason,linkage,profile,drive) Values(?,?,?,?,?,?)")

    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    conDB.Synchronous = False
    conDB.BeginTrans

    Cmd.SetNull 1   'id
    Cmd.SetInt32 2, testMastcode_module   'test_id
    Cmd.SetText 3, dbdata.reason   'reason
    Cmd.SetText 4, dbdata.linkage   'linkage
    Cmd.SetNull 5   'profile
    Cmd.SetNull 6  'drive

    Cmd.Execute
    conDB.CommitTrans
    conDB.Synchronous = True


'insertDataFailTesTb = conDB.LastInsertAutoID



End Function




Public Function UpdateEventIDToTestdb(test_id As Long, event_id As Long, isStart As Boolean) As Integer
'向试验记录表更新事件ID
'参数1 事件ID, 参数2 字段名称 (开始事件/结束事件) 如果开始为true 结束为false
'返回值 更新条目数, 如果失败返回-1
Dim RsTestsLast  As cRecordset

Set RsTestsLast = conDB.OpenRecordset("select * from Testable where _id = '" & test_id & "' order by _id desc limit 0,1")

If (isStart) Then
RsTestsLast!s_event_id.value = event_id
Else
RsTestsLast!e_event_id.value = event_id
End If

RsTestsLast.UpdateBatch
End Function


Public Function getDataRecordset(table As String, Optional sql As String) As cRecordset
If Not (IsMissing(sql)) Then
Set getDataRecordset = conDB.OpenRecordset("Select * From " & table)
Else
Set getDataRecordset = conDB.OpenRecordset("Select" & "From" & table)
End If

'getDataRecordset = Rs
End Function

Public Function delDataParamTb(sql As String) As Integer

End Function

Public Function delDataTesTb(sql As String) As Integer

End Function

Public Function delDataEvenTb() As Integer
Dim RsEventLast  As cRecordset

Set RsEventLast = conDB.OpenRecordset("Select * From EvenTable")
RsEventLast.MoveLast

RsEventLast.Delete

RsEventLast.GetRows


Dim Cmd As cCommand
Set Cmd = conDB.CreateCommand("Delete From EvenTable Where _id <100")
Cmd.Execute
delDataEvenTb = 1

End Function

Public Function delDataFailTb(sql As String) As Integer

End Function


Public Function cleanDatadb(time_stamp As Long) As Long
'清理数据库数据
'参数 给定一个时间戳,该时间戳之前的试验,事件,失败记录均被清除



End Function


Public Function insertDataEquipConfigtb(dbdata As EquipConfigtb_Data) As Integer
'向设备参数表中添加数据
'参数 添加的设备参数结构体
'返回值 添加成功返回添加条目的ID值,添加失败返回-1

End Function

'''''''''''''''''''''''''''''''''''''''''''''''


Public Function startTestRecordData(dbdata As StartEvent_Data) As Long
'封装函数 开始试验事件调用函数 按顺序向三个表中添加数据
'返回值 新记录试验表的ID

'实验开始
'向参数表添加数据 接受MD5码
'向试验表添加数据 接受试验表id
'向事件表添加数据 接受插入事件id
'更新实验表开始事件数据
'
Dim event_id As Long
Dim updateCount As Integer

md5code_module = insertDataParamTb_StartEvent(dbdata)
testMastcode_module = insertDataTesTb_StartEvent(dbdata)
event_id = insertDataStartEvenTb(dbdata)

updateCount = UpdateEventIDToTestdb(testMastcode_module, event_id, True)

End Function


Public Function selfCheckRecordData(dbdata As ProcessEvent_Data) As Long
'自检事件记录
selfCheckRecordData = insertDataSelfCheckEvenTb(dbdata)

End Function


Public Function processRecordData(dbdata As ProcessEvent_Data) As Long
'过程型事件记录
'
processRecordData = insertDataProcessEvenTb(dbdata)

End Function



Public Function endTestRecordData(dbdata As ProcessEvent_Data) As Long
'试验结束记录
Dim event_id As Long
Dim updateCount As Integer

event_id = insertDataProcessEvenTb(dbdata)
updateCount = UpdateEventIDToTestdb(testMastcode_module, event_id, False)
endTestRecordData = 1
End Function



Public Function failTestRecordData(eventdata As ProcessEvent_Data, faildata As FailEvent_Data)
'试验失败记录
Dim count As Long
count = endTestRecordData(eventdata)
Call insertDataFailTesTb(faildata)

End Function




Public Function QueryMd5ParamTb(md5code As String) As Boolean

If (md5code = "") Then Exit Function

Dim RsParametersForMd5 As cRecordset

Set RsParametersForMd5 = conDB.OpenRecordset("Select _id from ParameterTable Where md5='" & md5code & "'")

If (RsParametersForMd5.RecordCount = 0) Then QueryMd5ParamTb = False Else QueryMd5ParamTb = True


End Function

