VERSION 5.00
Begin VB.Form About 
   Caption         =   "About the Admin Tool"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2490
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copyright 1999, Library of Michigan Foundation"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2235
      Left            =   120
      Picture         =   "About.frx":030A
      Top             =   120
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Version 1.0 6/8/99"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "About"
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

Private Sub Command1_Click()
  Hide
End Sub

Private Sub Image1_Click()
    LaunchBrowser Me, "http://www.mel.org/citoolkit/"
End Sub

Private Sub Label2_Click()
   LaunchBrowser Me, "http://www.libofmich.lib.mi.us/foundation/foundation.html"
End Sub
