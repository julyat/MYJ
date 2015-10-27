Attribute VB_Name = "canshu"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long
                                            
'--------------------
'set_ini(文件路径,节点名,关键字,值)
'--------------------
Public Sub set_ini(ByVal fileName As String, ByVal App As String, ByVal key As String, ByVal strValue As String)
    Dim Result As Long
    Result = WritePrivateProfileString(App, key, strValue, fileName)
End Sub
'--------------------
'get_ini(文件路径,节点名,关键字,值的最高字符大小)
'--------------------
Public Function get_ini(ByVal fileName As String, _
                        Section As String, _
                        key As String, _
                        Size As Long) As String
    Dim ReturnStr As String
    Dim ReturnLng As Long
    get_ini = vbNullString
    ReturnStr = Space(Size)
    ReturnLng = GetPrivateProfileString(Section, key, vbNullString, ReturnStr, Size, fileName)
    get_ini = Left(ReturnStr, InStr(1, ReturnStr, Chr(0)) - 1)
End Function


