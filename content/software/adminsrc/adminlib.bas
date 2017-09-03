Attribute VB_Name = "Module1"
' ADMINLIB - A few Visual Basic Routines
'
'  Copyright (C) 1999
'
'  This program is free software; you can redistribute it and/or modify
'  it under the terms of version 2 of the GNU General Public License as
'  published by the Free Software Foundation.
'
'  This program is distributed in the hope that it will be useful,
'  but WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'  GNU General Public License for more details.
'
'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software
'  Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'
'  April 4, 1999


Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer
Option Explicit

Public Const INFINITE = &HFFFF
'STARTINFO constants
Private Const STARTF_USESHOWWINDOW = &H1
Public Enum enSW
     SW_HIDE = 0
     SW_NORMAL = 1
     SW_MAXIMIZE = 3
     SW_MINIMIZE = 6
End Enum

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
                 cb As Long
                 lpReserved As String
                 lpDesktop As String
                 lpTitle As String
                 dwX As Long
                 dwY As Long
                 dwXSize As Long
                 dwYSize As Long
                 dwXCountChars As Long
                 dwYCountChars As Long
                 dwFillAttribute As Long
                 dwFlags As Long
                 wShowWindow As Integer
                 cbReserved2 As Integer
                 lpReserved2 As Byte
                 hStdInput As Long
                 hStdOutput As Long
                 hStdError As Long
End Type

Type SECURITY_ATTRIBUTES
                 nLength As Long
                 lpSecurityDescriptor As Long
                 bInheritHandle As Long
End Type
          
Public Enum enPriority_Class
             NORMAL_PRIORITY_CLASS = &H20
             IDLE_PRIORITY_CLASS = &H40
             HIGH_PRIORITY_CLASS = &H80
End Enum
          
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
     lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As _
     SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal _
     dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory _
     As String, lpStartupInfo As STARTUPINFO, lpProcessInformation _
     As PROCESS_INFORMATION) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

' Thanks to http://home.sol.no/~jansh/vb/vbfaq.htm

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
     "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
     String, ByVal lpFile As String, ByVal lpParameters As String, _
     ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub LaunchBrowser(FormName As Form, URL As String)
  Dim ret&
  ret& = ShellExecute(FormName.hwnd, "Open", URL, "", App.Path, 1)
End Sub
 

Sub Split(OldString As String, SplitString() As String, N As Integer)
    
    Dim J As Integer
    
    OldString = Trim(OldString)
    N = 0
    Do While Len(OldString) > 0
      J = InStr(1, OldString, " ")
      N = N + 1
      If J = 0 Then
        SplitString(N) = Trim(OldString)
        OldString = ""
      Else
        SplitString(N) = Left(OldString, J - 1)
        OldString = Trim(Mid(OldString, J + 1))
      End If
'      Debug.Print N; "="; SplitString(N)
    Loop
End Sub

' http://www.global.com.gr/~vbasic/
' http://www.missouri.edu/~finaidtk/supshell.htm

Public Function SuperShell(ByVal App As String, ByVal WorkDir As String, _
    dwMilliseconds As Long, ByVal start_size As enSW, ByVal Priority_Class _
    As enPriority_Class) As Boolean

    Dim pclass As Long
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    'Not used, but needed
    Dim sec1 As SECURITY_ATTRIBUTES
    Dim sec2 As SECURITY_ATTRIBUTES
             
    sec1.nLength = Len(sec1)
    sec2.nLength = Len(sec2)
    sinfo.cb = Len(sinfo)
             
    sinfo.dwFlags = STARTF_USESHOWWINDOW
    sinfo.wShowWindow = start_size
             
    pclass = Priority_Class
             
    If CreateProcess(vbNullString, App, sec1, sec2, False, pclass, _
               0&, WorkDir, sinfo, pinfo) Then
      WaitForSingleObject pinfo.hProcess, dwMilliseconds
      SuperShell = True
    Else
      SuperShell = False
    End If

End Function

' If a drive letter ispart of a path, check to see if it is OK.

Public Function CRS_Chkdrv(DirName As String)
  CRS_Chkdrv = False
  If Mid(DirName, 2, 1) = ":" Then
    On Error Resume Next
    ChDrive Left(DirName, 1)
    If Err.Number = 71 Or Err.Number = 68 Then
      MsgBox ("Error - Drive Not Ready - " & DirName)
      On Error GoTo 0
      Exit Function
    End If
    If Err.Number <> 0 Then
      MsgBox ("Error - Drive Letter wrong (" & Err.Number & ") - " & DirName)
      On Error GoTo 0
      Exit Function
    End If
    CRS_Chkdrv = True
  End If
End Function

'Make a Directory with a prompt

Public Function CRS_Mkdir(DirName As String)
  Dim Well As Integer
  Dim QuitMsg As String
  Dim IPos As Integer
  Dim JPos As Integer
  Dim KPos As Integer
  Dim TmpDir As String
  
  CRS_Mkdir = False
    
  IPos = 1
  If Mid(DirName, 2, 1) = ":" Then
    On Error Resume Next
    ChDrive Left(DirName, 1)
    If Err.Number <> 0 Then
      MsgBox ("Error - Drive Letter wrong - " & DirName)
      On Error GoTo 0
      Exit Function
    End If
    IPos = 3
  End If
    
' Check to see of Directory Exists
  ChDir DirName
  If Err.Number = 76 Then
    Well = MsgBox("OK to create " & DirName, vbOKCancel)
    If Well <> vbOK Then
      QuitMsg = "Unable to create " & DirName
      GoTo WeQuit:
    End If
  End If
    
' Make all of the directories starting from the top down

  Do
    TmpDir = Left(DirName, IPos)
    Debug.Print "CRS_Mkdir Checking "; TmpDir, IPos, DirName
    On Error Resume Next
    ChDir TmpDir
    If Err.Number = 76 Then
      QuitMsg = "Unable to create " & TmpDir
      On Error GoTo WeQuit:
      Debug.Print "CRS_Mkdir Making "; TmpDir, IPos, DirName
      MkDir TmpDir
      QuitMsg = "Unable to enter " & TmpDir
      ChDir TmpDir
    End If
    On Error GoTo 0
    KPos = InStr(IPos + 1, DirName, "\")
    If KPos = 0 Then
      IPos = Len(DirName)
    Else
      IPos = KPos
    End If
  Loop While Len(Trim(TmpDir)) < Len(Trim(DirName))
  CRS_Mkdir = True
  Exit Function
WeQuit:
    MsgBox (QuitMsg)
    CRS_Mkdir = False
End Function

Public Function CRS_ChkDir(DirName As String)
  Dim IsFiles As String
    On Error Resume Next
    IsFiles = Dir(DirName, vbDirectory)
    If Err.Number <> 0 Or IsFiles = "" Then
      CRS_ChkDir = False
    Else
      CRS_ChkDir = True
    End If
    On Error GoTo 0
End Function

Public Function CRS_ChkFile(FileName As String)
  Dim INum As Integer
  INum = FreeFile
  CRS_ChkFile = True
  On Error Resume Next
  Open FileName For Input As #INum
  If Err.Number <> 0 Then
    CRS_ChkFile = False
  End If
  On Error GoTo 0
  Close #INum
End Function

Public Function CRS_MustExist(FileName As String)
  Dim INum As Integer
  INum = FreeFile
  CRS_MustExist = True
  On Error Resume Next
  Open FileName For Input As #INum
  If Err.Number <> 0 Then
    MsgBox (FileName & " does not exist " & Err.Description)
    CRS_MustExist = False
  End If
  Close #INum
  On Error GoTo 0
End Function

Public Function CRS_DelOK(FileName As String)
  Dim INum As Integer
  Dim IRET As Integer
  INum = FreeFile
  CRS_DelOK = True
  On Error Resume Next
  Open FileName For Input As #INum
  If Err.Number <> 0 Then
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0
  Close #INum
  
' File Exists, OK to delete it?

  IRET = MsgBox("Ok to remove  " & FileName, vbOKCancel)
  If IRET <> vbOK Then
    CRS_DelOK = False
    Exit Function
  End If
  On Error Resume Next
  Kill FileName
  If Err.Number <> 0 Then
    MsgBox ("Unable to remove " & FileName)
    CRS_DelOK = False
  End If
  On Error GoTo 0
End Function

Public Function CRS_Wait(Sec As Integer)
  Dim x As Integer
  x = Timer
  Debug.Print "Waiting...", Sec
  Do
    DoEvents
  Loop While (x + Sec) > Timer And x <= Timer
  Debug.Print "Waited..."
End Function


Public Sub RepStr(RepOld As String, RepNew As String, OldString As String, NewString As String)
  Dim Pos As Integer
  Dim TmpString As String
  
' Debug.Print "OLD:", OldString
  NewString = OldString
  TmpString = OldString
  Pos = 1
  Do   'Replacing RepOld With RepNew
    Pos = InStr(Pos, TmpString, RepOld)
    If Pos = 0 Then Exit Do
    NewString = Left(TmpString, Pos - 1) & RepNew & Mid(TmpString, Pos + Len(RepOld))
    Pos = Pos + Len(RepNew) + 1
    TmpString = NewString
  Loop
' Debug.Print "NEW:", NewString

End Sub

