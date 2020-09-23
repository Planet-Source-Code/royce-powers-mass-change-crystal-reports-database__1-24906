VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Change Report Databases"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5400
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.FileListBox filFiles 
      Height          =   4770
      Left            =   3720
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   11160
      TabIndex        =   13
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "nicholas"
      Top             =   6810
      Width           =   2535
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "rpowers"
      Top             =   6330
      Width           =   2535
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   1080
      TabIndex        =   9
      Text            =   "NEO"
      Top             =   5820
      Width           =   2535
   End
   Begin VB.DirListBox DirDest 
      Height          =   4365
      Left            =   7440
      TabIndex        =   8
      Top             =   1320
      Width           =   3615
   End
   Begin VB.DriveListBox drvDest 
      Height          =   315
      Left            =   7440
      TabIndex        =   7
      Top             =   960
      Width           =   3615
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Make Change"
      Default         =   -1  'True
      Height          =   735
      Left            =   9240
      TabIndex        =   3
      Top             =   6840
      Width           =   1695
   End
   Begin VB.DirListBox DirSource 
      Height          =   4365
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.DriveListBox drvSource 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Destination Folder:"
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "User ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Source report folder:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'    Dim fsoFiles As New FileSystemObject
'    Dim folFolder As Folder
'    Dim sf As Folder
'    Dim strPath As String
'    Dim ts As TextStream
'    Dim crpReport As New CRPEAuto.Report
'    Dim crpApplication As New CRPEAuto.Application
    
Private Sub cmdChange_Click()
On Error GoTo ErrorHandler
    Dim LogInValue As Integer
    
    With CrystalReport1
        .ReportFileName = filFiles.Path & "\" & filFiles.List(0)
        .Connect = "DSN=" & txtServer & ";UID=" & txtUserID & ";PWD=" & txtPassword
        .Destination = crptToFile
        .PrintFileType = crptCrystal
        .PrintFileName = txtPath.Text & filFiles.List(0)
        LogInValue = .LogOnServer("pdssql.dll", txtServer.Text, "cfsReport", txtUserID.Text, txtPassword.Text)
        .RetrieveSQLQuery
        MsgBox .SQLQuery, vbOKOnly, "SQL String"
        .Action = 1
    End With


'    With CrystalReport1
'        .ReportFileName = "C:\Old\sgt02.rpt"
'        .DataFiles(0) = "C:\New\xtreme.mdb"
'        .Destination = crptToFile
'        .PrintFileType = crptCrystal
'        .PrintFileName = "C:\New\sgt02.rpt"
'        .Action = 1
'    End With
'
'
'    With CrystalReport1
'        .ReportFileName = "C:\Old\AccessODBC.rpt"
'        .Connect = "DSN=MDB;UID=Admin;PWD=Password;DBQ=<CRWDC>DBQ=C:\New\xtreme.mdb"
'        .Destination = crptToFile
'        .PrintFileType = crptCrystal
'        .PrintFileName = "C:\New\AccessODBC.rpt"
'        .Action = 1
'    End With

ErrorHandler:
    MsgBox Err.Description, vbCritical, "Error " & Err.Number

End Sub

Private Sub cmdCreate_Click()
    fsoFiles.CreateFolder txtPath
    Call DirDest_Change
End Sub

Private Sub DirDest_Change()
    txtPath.Text = DirDest.Path
    If Right(DirDest.Path, 1) <> "\" Then txtPath = txtPath & "\"
End Sub

Private Sub DirSource_Change()
    filFiles.Path = DirSource.Path
    filFiles.Refresh
End Sub

Private Sub drvDest_Change()
On Error GoTo HandleError
    
    DirDest.Path = drvDest.Drive
    
HandleError:
    MsgBox Err.Number, vbCritical, Err.Description
    
End Sub

Private Sub filFiles_Click()
'    crpApplication.LogOnServer
'    Set crpReport = crpApplication.OpenReport(filFiles.path & "\" & filFiles.List(0))
   
        
        
End Sub

Private Sub filFiles_DblClick()
    With CrystalReport1
        .ReportFileName = filFiles.Path & "\" & filFiles.List(0)
        .Connect = "DSN=" & txtServer.Text & ";UID=" & txtUserID.Text & ";PWD=" & txtPassword.Text
        .PrintFileODBCPassword = txtPassword.Text
        .PrintFileODBCUser = txtUserID.Text
        .UserName = txtUserID.Text
        .Password = txtPassword.Text
        .PrintReport
    End With

End Sub

Private Sub Form_Load()
    Call DirDest_Change
End Sub
