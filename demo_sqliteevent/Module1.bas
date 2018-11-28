Attribute VB_Name = "Sqlite_Tool"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public countProMs As Currency
Public countProS As Currency
Public temp_time_s As Currency
Public temp_time_e As Currency
Private isInit  As Boolean

Public Function Sqlite_Timing(state As Boolean) As Single
If Not isInit Then
Call QueryPerformanceFrequency(countProS)
countProMs = countProS / 1000
isInit = True
End If


If state Then
Call QueryPerformanceCounter(temp_time_s)

Else
Call QueryPerformanceCounter(temp_time_e)

Sqlite_Timing = CSng((temp_time_e - temp_time_s) / countProS * 1000)

End If

End Function



Public Function getNowTimestamp() As Long

    Dim ToUnixTime  As String

    Dim intTimeZone As Integer

    intTimeZone = 8
    ToUnixTime = DateAdd("h", -intTimeZone, Now())
    getNowTimestamp = CLng(DateDiff("s", "1970-1-1 0:0:0", ToUnixTime))

End Function

'生成[n,m]之间的整数 使用在数组下标时注意m-1

Public Function randomInt(m As Integer, Optional n As Integer = 0) As Integer

    randomInt = Int(Rnd * (Abs(m - n) + 1)) + IIf(m > n, n, m)

End Function

Public Function randomSingle(m As Integer, Optional n As Integer = 0) As Single

    randomSingle = Rnd * (Abs(m - n) + 1) + IIf(m > n, n, m)

End Function


Public Function GetTestName() As String

    Dim names

    names = Array("sine", "fixeSine", "random", "shock", "rnr", "vsa")
    GetTestName = CStr(names(randomInt(5)))
End Function

Public Function GetFailName() As String

    Dim names

    names = Array("Input mutation", "Amplitude go beyond", "Frequency beyond", "Emergency stop", "Max driver", "H-Abort")
    GetFailName = CStr(names(randomInt(5)))
    'GetFailName = CStr(names(1))
    
End Function

Public Function GetFileFromdb() As String

    Dim names

    names = Array("Input mutation", "Amplitude go beyond", "Frequency beyond", "Emergency stop", "Max driver", "H-Abort")
    GetFileFromdb = CStr(names(randomInt(5)))
End Function


Private Function ConvertStringToBytes(ByVal s As String) As Byte()
    ConvertStringToBytes = StrConv(s, vbUnicode)
End Function



'直接在原始库中修改缓存获取文件MD5码的函数,相对速度稍快
Public Function GetFileMD5(ByVal filename As String) As String

    Dim md5 As New clsMD5

    GetFileMD5 = md5.DigestFileToHexStr(filename)

End Function

Public Function GetFilePath() As String

    GetFilePath = App.Path & "\test" & randomInt(5) & ".bin"

End Function




