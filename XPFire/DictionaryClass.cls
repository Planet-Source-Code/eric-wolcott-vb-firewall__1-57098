VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dictCache As New Dictionary
Private intMaxCacheSize As Integer

Public Function CheckDictionary(strIP As String) As String
    Dim strCache As String
        On Error Resume Next
        If dictCache.Exists(strIP) Then
            CheckDictionary = dictCache(strIP)
        Else
            Err.Clear
            CheckDictionary = GetHostNameFromIP(strIP)
            dictCache.Add strIP, CheckDictionary
            While dictCache.Count > intMaxCacheSize
                dictCache.Remove dictCache.Keys(UBound(dictCache.Items))
            Wend
        End If
End Function

Private Sub Class_Initialize()
  Dim fso As New FileSystemObject
  Dim ff As Byte
  Dim strIP As String, strDomain As String
  intMaxCacheSize = Val(GetSetting(App.Title & "Firewall", "Cache", "MaxSize", 100))

If fso.FileExists(GetAppPath & "cache.ini") = False Then fso.CreateTextFile GetAppPath & "cache.ini", False



  'Read in the cache file
  ff = FreeFile
  On Error Resume Next
  Open GetSetting(App.Title & "Firewall", "Cache", "Filename", GetAppPath & "cache.ini") For Input As #ff
    While Not EOF(ff)
        Input #ff, strIP, strDomain
        dictCache.Add strIP, strDomain
    Wend
  Close #ff
      
End Sub

Public Sub WriteCache()
  Dim ff As Byte
  Dim strKey As Variant
    
    'Save the cache to a file
    ff = FreeFile
    Open GetSetting(App.Title & "Firewall", "Cache", "Filename", GetAppPath & "cache.ini") For Output As #ff
        For Each strKey In dictCache.Keys
            Print #ff, strKey & "," & dictCache(strKey)
        Next
    Close #ff
  
End Sub
