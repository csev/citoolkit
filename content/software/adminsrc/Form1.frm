VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Community Information Adminsitration"
   ClientHeight    =   4860
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelDest 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdDelASP 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "expire.htm"
      Height          =   375
      Left            =   6000
      TabIndex        =   25
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdminPw 
      Caption         =   "Admin PW"
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdLibRemove 
      Caption         =   "Delete Database"
      Height          =   375
      Left            =   4440
      MaskColor       =   &H8000000F&
      TabIndex        =   23
      Top             =   480
      Width           =   1455
   End
   Begin VB.FileListBox filDestPath 
      Height          =   1065
      Left            =   3720
      TabIndex        =   21
      Top             =   3600
      Width           =   2535
   End
   Begin VB.FileListBox filASPPath 
      Height          =   1065
      Left            =   120
      TabIndex        =   20
      Top             =   3600
      Width           =   2535
   End
   Begin VB.FileListBox filOrigPath 
      Height          =   1065
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton cmdNewLib 
      Caption         =   "Add Database"
      Height          =   375
      Left            =   4440
      MaskColor       =   &H8000000F&
      TabIndex        =   18
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Lib DSN"
      Height          =   615
      Left            =   1560
      MaskColor       =   &H8000000F&
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtASPPath 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton cmdASPPath 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdDestPath 
      Caption         =   "Browse"
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox txtDestPath 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Top             =   3240
      Width           =   3375
   End
   Begin VB.TextBox txtShort 
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtLong 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox txtOrigpath 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ListBox lstDSN 
      Height          =   1425
      Left            =   3720
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton cmdOrigBrowse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Lib DSN"
      Height          =   615
      Left            =   480
      MaskColor       =   &H8000000F&
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   6240
      Picture         =   "Form1.frx":030A
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   930
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      X1              =   3600
      X2              =   3600
      Y1              =   960
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label lblCurrentDSN 
      Alignment       =   2  'Center
      Caption         =   "Current System DataSet Names (DSN's)"
      Height          =   255
      Left            =   3840
      TabIndex        =   22
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lblAPSPath 
      Alignment       =   2  'Center
      Caption         =   "ASP library Path"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label lblDestPath 
      Alignment       =   2  'Center
      Caption         =   "Production Database Path"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label lblShort 
      Caption         =   "Short Name"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblLong 
      Caption         =   "Long Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Path to the distribution  mdb files"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuComand 
      Caption         =   "Command"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Database"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Remove Database"
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Set Admin PW"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "Build expire.htm"
      End
      Begin VB.Menu msuReset 
         Caption         =   "Reset Directories"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuDoc 
         Caption         =   "Documentation"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Toolkit Web Site"
      End
      Begin VB.Menu msuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Copyright (C) 1999  Library of Michigan Foundation
'
'  This software program may be used in non-commercial applications
'  without further permission from the copyright owner as long as this
'  copyright notice is maintained and prominently displayed.
'
'  THIS PRODUCT IS DISTRIBUTED WITHOUT WARRANTY
'  OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING
'  IMPLIED WARRANTIES OF MERCHANTABILITY OR FITNESS
'  FOR A PARTICULAR PURPOSE.

Dim SaveCurDir As String

Private Sub Status(Str As String)
  Debug.Print "S:" & Str
  lblStatus.Caption = Str
  DoEvents
  If (chkDebug.Value = 1) Then
    If (MsgBox(Str, vbOKCancel) = vbCancel) Then chkDebug.Value = 0
  End If
End Sub

Private Function FindDSN(ByVal DSNName As String) As Boolean
  Dim Found As Boolean
  Found = False
  Set DSN = CreateObject("odbctool.SystemDSN")
  For I = 1 To DSN.DSN.Count
    If (DSN.DSN.Item(I).DSNName = DSNName) Then
      Found = True
    End If
  Next I
  Set DSN = Nothing
  FindDSN = Found
End Function

Private Function DeleteDSN(ByVal DSNName As String) As Boolean
  Dim Iremove As Integer
  Iremove = -1
  Set DSN = CreateObject("odbctool.SystemDSN")
  For I = 1 To DSN.DSN.Count
    If (DSN.DSN.Item(I).DSNName = DSNName) Then
      MsgStr = "OK to delete Data Name " & DSN.DSN.Item(I).DSNName & "=" & DSN.DSN.Item(I).Attributes
      If MsgBox(MsgStr, vbOKCancel) = vbOK Then
        Iremove = I
      End If
    End If
  Next I
  If Iremove > 0 Then
    MsgStr = "DSN=" & DSN.DSN.Item(Iremove).DSNName
    DSN.RemoveDSN MsgStr, "Microsoft Access Driver (*.mdb)"
    Status (MsgStr & " removed.")
  End If
  Set DSN = Nothing
End Function

Private Function DoSelDir(TextVal As String) As String
  Load SelDir
  SelDir!DirName.Text = TextVal
  SelDir.Show vbModal
  If Right(RTrim(SelDir!DirName.Text), 1) <> "\" Then
    DoSelDir = SelDir!DirName.Text
  Else
    DoSelDir = SelDir!DirName.Text
  End If
  Unload SelDir
End Function


Private Sub cmdAdminPw_Click()
  Dim DestMDB As String
  
  If Len(Trim(txtShort.Text)) <= 0 Then
    MsgBox "Please enter a short library name"
    Exit Sub
  End If
  
  DestMDB = txtDestPath.Text & txtShort.Text & ".mdb"
  On Error Resume Next
  Set conPubs = OpenDatabase(DestMDB)
  Set Records = conPubs.OpenRecordset("SiteInfo", dbOpenDynaset)
    If Err.Number <> 0 Then
    MsgBox "Error Opening " & DestMDB
    Exit Sub
  End If
  
' Set the Admin Password

  Status ("Setting admin password")
  NewStr = InputBox("Please enter admin password")
  Set Records = conPubs.OpenRecordset("Users", dbOpenDynaset)
  Debug.Print Records.Fields("Password").Name
  Debug.Print Records!Password
  Records.MoveFirst
  Records.Edit
  Records!Password = NewStr
  Records!AccountLocked = vbFalse
  Records.Update
  If Err.Number <> 0 Then
    MsgBox "Error Updating " & DestMDB
    Exit Sub
  End If
  Records.Close

  Set conPubs = Nothing
  Refresh_All
  Status ("Unlocked admin account in " & DestMDB)
  Status ("Set admin password in " & DestMDB)
End Sub

Private Sub cmdASPPath_Click()
  txtASPPath.Text = DoSelDir(txtASPPath.Text)
End Sub

Private Sub cmdDefault_Click()
  
  Dim DestMDB As String
  Dim NewASP As String
  Dim MyFile
  Dim FNum As Integer
  Dim FName As String
  Dim TNum As Integer
  
  If Len(Trim(txtShort.Text)) <= 0 Then
    MsgBox "Please enter a short library name"
    Exit Sub
  End If
 
' Create expired.htm

  FNum = FreeFile()
  FName = txtASPPath.Text & "expired.htm"

  On Error Resume Next
  Open FName For Output As #FNum
  Debug.Print Err.Number; Err.Description
  If Err.Number <> 0 Then
    MsgBox ("Error - unable to write to " & FName)
    Exit Sub
  End If
  On Error GoTo 0

' Concatenate all the .inf files into the default.htm file

  MyFile = Dir(txtASPPath.Text & "*.inf")
  First = vbTrue
  While MyFile <> ""
    Debug.Print MyFile
    Status ("Copying - " & MyFile & " to " & FName)
    
    TNum = FreeFile()
    On Error Resume Next
    Open txtASPPath.Text & MyFile For Input As #TNum
    Debug.Print Err.Number; Err.Description
    If Err.Number <> 0 Then
      MsgBox ("Error - unable to copy " & MyFile)
    Else
      If First Then
        Print #FNum, "<h3>Server Variables Expired</h3>"
        Print #FNum, "Your page has expired, please select from the following"
        Print #FNum, "<ul>"
        First = vbFalse
      End If
      Print #FNum, "<li>"
      While Not EOF(TNum) ' Loop until end of file.
        Input #TNum, InpStr
        Debug.Print "InpStr", InpStr
        Print #FNum, InpStr
      Wend
      Close #TNum
    End If
    MyFile = Dir ' Advance to the next file
  Wend
  If Not First Then
    Print #FNum, "</ul>"
  End If
  Close #FNum
  
  Refresh_All



 


End Sub

Private Sub cmdDelASP_Click()
  Dim FName As String
 
  FName = txtASPPath.Text & filASPPath.FileName
  If MsgBox("Ok to Delete " & FName & "?", vbOKCancel) = vbOK Then
    On Error Resume Next
    Kill FName
    On Error GoTo 0
  End If
  Refresh_All
End Sub

Private Sub cmdDelDest_Click()
  Dim FName As String
 
  FName = txtDestPath.Text & filDestPath.FileName
  If MsgBox("Ok to Delete " & FName & "?", vbOKCancel) = vbOK Then
    On Error Resume Next
    Kill FName
    On Error GoTo 0
  End If
  Refresh_All
End Sub

Private Sub cmdDestPath_Click()
  txtDestPath.Text = DoSelDir(txtDestPath.Text)
End Sub

' Note - this is no longer necessary - the lib DSN is not used as of Version Beta 1

Private Sub cmdCreate_Click()
  Dim OrigMDB As String
  Dim DestMDB As String
  
  OrigMDB = txtOrigpath.Text & "lib.mdb"
  DestMDB = txtDestPath.Text & "lib.mdb"

  If Not CRS_ChkFile(OrigMDB) Then
    MsgBox (OrigMDB & " file does not exist")
    Exit Sub
  End If
  
  If Trim(OrigMDB) <> Trim(DestMDB) Then
    Status ("Copying lib.mdb file")
    FileCopy OrigMDB, DestMDB
  End If

  If Not CRS_ChkFile(DestMDB) Then
    MsgBox (DestMDB & " file does not exist")
    Exit Sub
  End If

  DeleteDSN ("lib")
  Status ("Existing DSN deleted")
  
  Set DSN = CreateObject("odbctool.SystemDSN")
  attrib = "DBQ=" & DestMDB & ";DSN=lib"
  DSN.CreateDSN attrib, "Microsoft Access Driver (*.mdb)"
  Set DSN = Nothing
  Refresh_All
  Status ("lib System DSN Created")
End Sub

Private Sub Refresh_DSN()
  Status ("Scanning for System DSN's")
  lstDSN.Clear
  Set DSN = CreateObject("odbctool.SystemDSN")
  Debug.Print DSN.DSN.Count
  For I = 1 To DSN.DSN.Count
    Status ("Found " & I & " " & DSN.DSN.Item(I).DSNName)
    lstDSN.AddItem (DSN.DSN.Item(I).DSNName)
  Next I
  Status ("Found " & DSN.DSN.Count & " Data Set Names")
End Sub

Private Sub Refresh_All()
  Refresh_DSN
  filDestPath.Refresh
  filASPPath.Refresh
  filOrigPath.Refresh
End Sub


Private Sub cmdLibRemove_Click()

  Dim FName As String
  Dim NewDSN As String
  Dim DestMDB As String
  Dim OrigMDB As String
  Dim NewASP As String

  If Len(Trim(txtShort.Text)) <= 0 Then
    MsgBox "Please enter a short library name"
    Exit Sub
  End If
  
  If Not CRS_ChkDir(txtDestPath.Text) Then
    MsgBox txtDestPath.Text & " - Directory does not exist"
    Exit Sub
  End If

  NewDSN = "ctk_" & txtShort.Text
  DestMDB = txtDestPath.Text & txtShort.Text & ".mdb"
    
  ConStr = "DSN=" & NewDSN & ";DBQ=" & DestMDB & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5;"

  On Error Resume Next
  If CRS_ChkFile(DestMDB) Then
    If MsgBox("Ok to delete - " & DestMDB, vbOKCancel) = vbOK Then
      Kill DestMDB
    Else
      Exit Sub
    End If
  End If
  
  NewASP = txtASPPath.Text & txtShort.Text & ".inf"
  Kill NewASP
  
  NewASP = txtASPPath.Text & txtShort.Text & ".asp"
  Kill NewASP

  NewASP = txtASPPath.Text & txtShort.Text & "-admin.asp"
  Kill NewASP
  
  NewASP = txtASPPath.Text & txtShort.Text & "-forum.asp"
  Kill NewASP

  NewASP = txtASPPath.Text & txtShort.Text & "-calendar.asp"
  Kill NewASP

  On Error GoTo 0
  
' Update the default.htm file

  cmdDefault_Click
  
'Delete the DSN

  Status ("Deleting the DSN " & NewDSN)
  DeleteDSN (NewDSN)
  Status ("Existing DSN deleted")

  Refresh_All
End Sub

Private Sub cmdNewLib_Click()

  Dim FName As String
  Dim NewDSN As String
  Dim DestMDB As String
  Dim OrigMDB As String
  Dim NewASP As String
  Dim FNum As Integer
  
  If Len(Trim(txtShort.Text)) <= 0 Then
    MsgBox "Please enter a short library name"
    Exit Sub
  End If
  
  If Not CRS_ChkDir(txtOrigpath.Text) Then
    MsgBox txtOrigpath.Text & " - Directory does not exist"
    Exit Sub
  End If
  
  If Not CRS_ChkDir(txtDestPath.Text) Then
    If CRS_Mkdir(txtDestPath.Text) = False Then Exit Sub
    txtDestPath_Change
  End If
  
  If Not CRS_ChkDir(txtASPPath.Text) Then
    If CRS_Mkdir(txtASPPath.Text) = False Then Exit Sub
    txtASPPath_Change
  End If
  
' Copy the MDB File

  NewDSN = "ctk_" & txtShort.Text
  DestMDB = txtDestPath.Text & txtShort.Text & ".mdb"
  OrigMDB = txtOrigpath.Text & "base.mdb"

  If Not CRS_ChkFile(OrigMDB) Then
    MsgBox (OrigMDB & " file does not exist")
    Exit Sub
  End If

  If CRS_ChkFile(DestMDB) Then
    If MsgBox(DestMDB & " already exists - OK to Overwrite with empty file?", vbOKCancel) = vbCancel Then
      Exit Sub
    End If
  End If
  
  If Trim(OrigMDB) <> Trim(DestMDB) Then
    Status ("Copying base.mdb file to " & DestMDB)
    On Error Resume Next
    FileCopy OrigMDB, DestMDB
    If Err.Number <> 0 Then
      MsgBox ("Error in copy " & OrigMDB & " to " & DestMDB)
      On Error GoTo 0
      Exit Sub
    End If
  Else
    MsgBox ("Error - Source and destination MDB file are the same")
    Exit Sub
  End If

  If Not CRS_ChkFile(DestMDB) Then
    MsgBox (DestMDB & " file does not properly created")
    Exit Sub
  End If
  
' Create the four ASP Files
  
  ConStr = "DSN=" & NewDSN & ";DBQ=" & DestMDB & ";DriverId=25;FIL=MS Access;MaxBufferSize=512;PageTimeout=5;"

  NewASP = txtASPPath.Text & txtShort.Text & ".asp"
  ReDir = "../index.asp"
  MakeASP NewASP, ConStr, ReDir
  
  NewASP = txtASPPath.Text & txtShort.Text & "-admin.asp"
  ReDir = "../Admin/Login_Guts_Frame.asp?mode=1"
  MakeASP NewASP, ConStr, ReDir
  
  NewASP = txtASPPath.Text & txtShort.Text & "-forum.asp"
  ReDir = "../forums/default.htm"
  MakeASP NewASP, ConStr, ReDir

  NewASP = txtASPPath.Text & txtShort.Text & "-calendar.asp"
  ReDir = "../calendar/calendar.asp"
  MakeASP NewASP, ConStr, ReDir
 
' Create the short.htm file

  FNum = FreeFile()
  FName = txtASPPath.Text & txtShort.Text & ".inf"

  On Error Resume Next
  Open FName For Output As #FNum
  Debug.Print Err.Number; Err.Description
  If Err.Number <> 0 Then
    MsgBox ("Error - unable to write to " & FName)
    On Error GoTo 0
  Else
    Status ("Creating " & FName)
    Print #FNum, "<a href=" & txtShort.Text & "-calendar.asp>" & txtLong.Text & " Calendar</a><br>"
    Print #FNum, "<a href=" & txtShort.Text & "-forum.asp>" & txtLong.Text & " Forum</a><br>"
    Print #FNum, "<a href=" & txtShort.Text & "-admin.asp>" & txtLong.Text & " Administration</a><br>"
    Close #FNum
  End If
  On Error GoTo 0

' Update the default.htm file

  cmdDefault_Click

' Create the DSN

  Status ("Deleting any old DSN " & NewDSN)
  DeleteDSN (NewDSN)
  Status ("Existing DSN deleted")

  Status (ConStr)
  Set DSN = CreateObject("odbctool.SystemDSN")
  DSN.CreateDSN ConStr, "Microsoft Access Driver (*.mdb)"
  Set DSN = Nothing
  Status (NewDSN & " System DSN Created")

' Put in the Organization Name
  Status ("Setting Organization to " & txtLong.Text)

  Set conPubs = OpenDatabase(DestMDB)
  Set Records = conPubs.OpenRecordset("SiteInfo", dbOpenDynaset)
  Debug.Print Records.Fields("OrgName").Name
  Debug.Print Records!OrgName
  Records.MoveFirst
  Records.Edit
  Records!OrgName = txtLong.Text
  Records.Update
  Records.Close
  
' Set the Admin Password

  Status ("Setting admin password")
  NewStr = InputBox("Please enter admin password")
  If Len(Trim(NewStr)) > 0 Then
    Set Records = conPubs.OpenRecordset("Users", dbOpenDynaset)
    Debug.Print Records.Fields("Password").Name
    Debug.Print Records!Password
    Records.MoveFirst
    Records.Edit
    Records!Password = NewStr
    Records.Update
    Records.Close
  End If

  Set conPubs = Nothing
  Refresh_All
  Status ("Created library database " & DestMDB)
End Sub

Private Function MakeASP(ByVal FName As String, ByVal ConStr As String, ByVal ReDir As String) As Boolean
  Dim FNum As Integer
  
  FNum = FreeFile()

  On Error Resume Next
  Open FName For Output As #FNum
  Debug.Print Err.Number; Err.Description
  If Err.Number <> 0 Then
    MsgBox ("Error - unable to write to " & FName)
    On Error GoTo 0
    MakeASP = False
    Exit Function
  End If
  On Error GoTo 0
  
  Status ("Creating " & FName)
  Print #FNum, "<%@ Language=VBScript %>"
  Print #FNum, "<% Session(""Connection1_ConnectionString"") = """ & ConStr & """"
  Print #FNum, "Session(""Connection1_ConnectionTimeout"") = 15"
  Print #FNum, "Session(""Connection1_CommandTimeout"") = 30"
  Print #FNum, "Session(""Connection1_CursorLocation"") = 3"
  Print #FNum, "Session(""Connection1_RuntimeUserName"") = """""
  Print #FNum, "Session(""Connection1_RuntimePassword"") = """""
  Print #FNum, "Response.Redirect """ & ReDir & """"
  Print #FNum, "%>"
  Close #FNum
  MakeASP = True
End Function

Private Sub cmdOrigBrowse_Click()
  With txtOrigpath
    Load SelDir
    SelDir!DirName.Text = .Text
    SelDir.Show vbModal
    .Text = SelDir!DirName.Text
    Unload SelDir
  End With
End Sub

' Note - this is no longer necessary - the lib DSN is not used as of Version Beta 1

Private Sub cmdRemove_Click()
  DeleteDSN ("lib")
  Refresh_All
End Sub

Private Sub cmdRemoveDSN_Click()
  Debug.Print
End Sub


Private Sub Form_Load()
  Dim PrefStr As String
  SaveCurDir = CurDir
  Debug.Print SaveCurDir
  
  PrefStr = GetSetting("Caladmin 1.0", "init", "OrigPath")
  If Len(Trim(PrefStr)) > 0 Then
    txtOrigpath.Text = PrefStr
  Else
    txtOrigpath.Text = CurDir & "\"
  End If
  
  PrefStr = GetSetting("Caladmin 1.0", "init", "DestPath")
  If Len(Trim(PrefStr)) > 0 Then
      txtDestPath.Text = PrefStr
  Else
    txtDestPath.Text = "C:\InetPub\databases\community\"
  End If

  PrefStr = GetSetting("Caladmin 1.0", "init", "ASPPath")
  If Len(Trim(PrefStr)) > 0 Then
      txtASPPath.Text = PrefStr
  Else
      txtASPPath.Text = "C:\InetPub\wwwroot\community\start\"
  End If


  Refresh_All
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveSetting "Caladmin 1.0", "init", "OrigPath", txtOrigpath.Text
  SaveSetting "Caladmin 1.0", "init", "DestPath", txtDestPath.Text
  SaveSetting "Caladmin 1.0", "init", "ASPPath", txtASPPath.Text
  Debug.Print "Saved"
End Sub

Private Sub Image1_Click()
    LaunchBrowser Me, "http://www.mel.org/citoolkit/"
End Sub

Private Sub lblStatus_Click()
  If (MsgBox("Would you like to turn on Debugging?", vbYesNoCancel) = vbYes) Then
    chkDebug.Value = 1
  Else
    chkDebug.Value = 0
  End If
End Sub

Private Sub mnuAdd_Click()
  cmdNewLib_Click
End Sub

Private Sub mnuAdmin_Click()
  cmdAdminPw_Click
End Sub

Private Sub mnuDefault_Click()
  cmdDefault_Click
End Sub

Private Sub mnuDel_Click()
  cmdLibRemove_Click
End Sub

Private Sub mnuDoc_Click()
  Dim Tmp As String
  Tmp = SaveCurDir & "\help.htm"
  Debug.Print Tmp
  LaunchBrowser Me, Tmp
End Sub

Private Sub mnuExit_Click()
  SaveSetting "Caladmin 1.0", "init", "OrigPath", txtOrigpath.Text
  SaveSetting "Caladmin 1.0", "init", "DestPath", txtDestPath.Text
  SaveSetting "Caladmin 1.0", "init", "ASPPath", txtASPPath.Text
  Debug.Print "Saved"
  End
End Sub

Private Sub mnuWeb_Click()
    LaunchBrowser Me, "http://www.mel.org/citoolkit/"
End Sub

Private Sub msuAbout_Click()
  Load About
  About.Show vbModal
  Unload About
End Sub

Private Sub msuReset_Click()
  txtOrigpath.Text = SaveCurDir & "\"
  txtDestPath.Text = "C:\InetPub\databases\community\"
  txtASPPath.Text = "C:\InetPub\wwwroot\community\start\"
  Refresh_All
End Sub

Private Sub txtASPPath_Change()
  filASPPath.Pattern = "*.asp"
  If CRS_ChkDir(txtASPPath.Text) Then
    filASPPath.Path = txtASPPath.Text
    Refresh_All
    txtASPPath.BackColor = vbWhite
  Else
    txtASPPath.BackColor = vbYellow
  End If
End Sub

Private Sub txtDestPath_Change()
  filDestPath.Pattern = "*.mdb"
  If CRS_ChkDir(txtDestPath.Text) Then
    filDestPath.Path = txtDestPath.Text
    filDestPath.Visible = True
    Refresh_All
    txtDestPath.BackColor = vbWhite
  Else
    txtDestPath.BackColor = vbYellow
    filDestPath.Visible = False
  End If
End Sub

Private Sub txtOrigpath_Change()
  With txtOrigpath
    .BackColor = vbWhite
    If CRS_ChkDir(.Text) Then
      If CRS_ChkFile(.Text & "base.mdb") Then
        '  Nothing
      Else
        .BackColor = vbYellow
      End If
    Else
      .BackColor = vbYellow
    End If
  End With
  filOrigPath.Pattern = "*.mdb"
  
  If CRS_ChkDir(txtOrigpath.Text) Then
    filOrigPath.Path = txtOrigpath.Text
    filOrigPath.Visible = True
    Refresh_All
  Else
    txtOrigpath.BackColor = vbYellow
    filOrigPath.Visible = False
  End If

End Sub

Private Sub txtShort_Change()
  Dim NewStr As String
  Dim OldStr As String
  
  OldStr = txtShort.Text
  J = 0
  NewStr = ""
  For I = 1 To Len(OldStr)
    ch = Mid(OldStr, I, 1)
    If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-" Then
      NewStr = NewStr & ch
    End If
  Next I
  
  If OldStr <> NewStr Then
    MsgBox ("Please use lower case letters, numbers or - in the short name")
    txtShort.Text = NewStr
  End If
End Sub
