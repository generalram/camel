VERSION 5.00
Begin VB.Form Frm_Main 
   Caption         =   "Form1"
   ClientHeight    =   6765
   ClientLeft      =   3120
   ClientTop       =   420
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmd_Help 
      Caption         =   "Help"
      Height          =   360
      Left            =   8160
      TabIndex        =   38
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmd_About 
      Caption         =   "About"
      Height          =   375
      Left            =   8160
      TabIndex        =   37
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmd_Search 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   3720
      Width           =   975
   End
   Begin VB.ComboBox cbo_MatterList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   5040
      Width           =   5775
   End
   Begin VB.CommandButton cmd_ClearMatterList 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   4920
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearAll 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox chk_MatterBillingActive 
      Caption         =   "Matter Billing Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   30
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CheckBox chk_MatterSearchActive 
      Caption         =   "Matter Search Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   29
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CheckBox chk_ClientBillingActive 
      Caption         =   "Client Billing Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   28
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chk_ClientSearchActive 
      Caption         =   "Client Search Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmd_ClearAttorney 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox cbo_AttorneyList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   4440
      Width           =   5775
   End
   Begin VB.CommandButton cmd_ClearMatterNumber 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearMatterName 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearClientNumber 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton cmd_ClearClientName 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txt_MatterName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txt_MatterNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox cbo_MatterNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Text            =   "Matter Number List"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ComboBox cbo_MatterName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Text            =   "Matter Name List"
      Top             =   2400
      Width           =   2655
   End
   Begin VB.ComboBox cbo_ClientName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Text            =   "Client Name List"
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox cbo_ClientNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   4
      Text            =   "Client Number List"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txt_ClientNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txt_ClientName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lbl_Title2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Client And Matter Electronic Legal System)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   0
      TabIndex        =   36
      Top             =   360
      Width           =   8115
   End
   Begin VB.Label lbl_SearchStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   35
      Top             =   3840
      Width           =   5655
   End
   Begin VB.Label lbl_MatterList 
      Caption         =   "Matter List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lbl_Copyright 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "  Copyright XORMAD 2003 - 2014,  All rights reserved. Duplication of this product without authorization from XORMAD is prohibited."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   -120
      TabIndex        =   31
      Top             =   6120
      Width           =   9255
   End
   Begin VB.Label lbl_Attorney 
      Caption         =   "Attorney List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   25
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lbl_MatterName 
      Caption         =   "Matter Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbl_MatterNumber 
      Caption         =   "Matter Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lbl_Or1 
      Alignment       =   2  'Center
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lbl_Title 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAMEL Search"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8115
   End
   Begin VB.Label lbl_ClientNumber 
      Caption         =   "Client Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lbl_ClientName 
      Caption         =   "Client Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_About_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub cmd_ClearClientName_Click()
    Frm_Main.txt_ClientName.Text = ""
    Frm_Main.cbo_ClientName.Text = "Client Name List"
End Sub

Private Sub cmd_ClearClientNumber_Click()
    Frm_Main.txt_ClientNumber.Text = ""
    Frm_Main.cbo_ClientNumber.Text = "Client Number List"
End Sub

Private Sub cmd_ClearMatterName_Click()
    Frm_Main.txt_MatterName.Text = ""
    Frm_Main.cbo_MatterName.Text = "Matter Name List"
End Sub

Private Sub cmd_ClearMatterNumber_Click()
    Frm_Main.txt_MatterNumber.Text = ""
    Frm_Main.cbo_MatterNumber.Text = "Matter Number List"
End Sub

Private Sub cmd_Exit_Click()
    ExitProgram
End Sub

Private Sub cmd_Help_Click()
    ' Disabled for now, not visible until help is added to the program.
End Sub

Private Sub cmd_Search_Click()
    Call CMSearch
        
End Sub

Private Sub ExitProgram()
    Unload Frm_Main 'Unload Main form from memory
    End ' End the program
End Sub

Private Sub Form_Load()
    Call InitCMSearch
End Sub

