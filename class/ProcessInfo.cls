VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal moduleName As String, ByVal nSize As Long) As Long

Const PROCESS_QUERY_INFORMATION = 1024
Const PROCESS_VM_READ = 16

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

'Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

Private qryHandle As Long

Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Type FILETIME: dwLowDateTime As Long: dwHighDateTime As Long: End Type
Private Type SYSTEMTIME
wYear As Integer
wMonth As Integer
wDayOfWeek As Integer
wDay As Integer
wHour As Integer
wMinute As Integer
wSecond As Integer
wMilliseconds As Integer
End Type

Private memFilename As String
Private memStartTime As String
Private memKernalTime As String
Private memMemoryUsage As Long
Private memProcessID As Long

Public Sub newQuery(ProcessID As Long)
 memProcessID = ProcessID
 Dim qryHandle As Long, lngResult As Long
 '>> open our interigating handle
 Const PROCESS_QUERY_INFORMATION = 1024
 Const PROCESS_VM_READ = 16
 qryHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
 '>> get the filename
 Dim moduleName As String * 256
 lngResult = GetModuleFileNameExA(qryHandle, 0, moduleName, Len(moduleName))
 memFilename = Left(moduleName, lngResult)
 memFilename = LCase(memFilename)
 memFilename = Replace(memFilename, "\??\", "")
 memFilename = Replace(memFilename, "\systemroot", Environ("windir"))
 memFilename = Replace(memFilename, "\\", "\")
 '>> get handle start time / cpu time
 Dim start As FILETIME, finish As FILETIME, kernal As FILETIME, user As FILETIME
 Dim sysTime As SYSTEMTIME
 GetProcessTimes qryHandle, start, finish, kernal, user
 FileTimeToLocalFileTime start, start
 FileTimeToSystemTime start, sysTime
 memStartTime = sysTime.wDay & "/" & sysTime.wMonth & "/" & sysTime.wYear & " " & sysTime.wHour & ":" & sysTime.wMinute & ":" & sysTime.wSecond
' FileTimeToLocalFileTime kernal, kernal
' FileTimeToSystemTime kernal, sysTime
 memKernalTime = kernal.dwLowDateTime  ' sysTime.wDay & "/" & sysTime.wMonth & "/" & sysTime.wYear & " " & sysTime.wHour & ":" & sysTime.wMinute & ":" & sysTime.wSecond
 
 '>> get memory usage
 Dim PMC As PROCESS_MEMORY_COUNTERS
 PMC.cb = LenB(PMC)
    
 GetProcessMemoryInfo qryHandle, PMC, PMC.cb
 memMemoryUsage = CLng(PMC.WorkingSetSize / 1024)
    
 '>> close out interigating handle
 CloseHandle qryHandle
End Sub

Public Property Get ProcessID() As Long
 ProcessID = memProcessID
End Property

Public Property Get ImageName() As String
 Dim fn As String
 fn = Filename()
 
 ImageName = Right(fn, Len(fn) - InStrRev(fn, "\"))
End Property

Public Property Get Filename() As String
 Filename = LCase(memFilename)
End Property

'Public Property Get UserName()
'
'End Property

'Public Property Get CPU()
'
'End Property

Public Property Get MemoryUsage() As Long
 MemoryUsage = memMemoryUsage
End Property

Public Property Get StartTime() As String
 StartTime = memStartTime
End Property

Public Property Get KernalTime() As String
 KernalTime = memKernalTime
End Property

