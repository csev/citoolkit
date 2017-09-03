VERSION 5.00
Begin VB.Form SelDir 
   Caption         =   "Select A Directory"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox DirName 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "SelDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sync-O-Matic - Lecture production tool
'
'  Copyright (C) 1997  Michigan State University
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
'  Charles Severance (crs@msu.edu)
'
'  December 17, 1997
'
'  Computer Science Department
'  3115 Engineering Building
'  Michigan State University
'  East Lansing, Michigan  48824
'  USA

Private Sub cmdOK_Click()
  DirName.Text = Trim(DirName.Text)
  If Right(DirName.Text, 1) <> "\" Then
    DirName.Text = DirName.Text & "\"
  End If
  SelDir.Hide
End Sub

Private Sub Dir1_Change()
  DirName.Text = Dir1.Path
  Debug.Print "In SelDir "; DirName.Text
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
  Debug.Print "Acitvating SelDir ", DirName.Text
  If DirName.Text <> "" Then
    On Error Resume Next
    If Mid(DirName.Text, 2, 1) = ":" Then
      On Error Resume Next
      Drive1.Drive = DirName.Text
      If Err.Number <> 0 Then
        MsgBox ("Error - Drive Letter wrong - " & DirName.Text)
        On Error GoTo 0
        Exit Sub
      End If
    End If
    Dir1.Path = DirName.Text
    If Err Then
      Debug.Print "Path Not Found ", DirName.Text
      Dir1.Path = "C:\"
      Drive1.Drive = "C:\"
      DirName.Text = "C:\"
    End If
    On Error GoTo 0
  End If
End Sub

