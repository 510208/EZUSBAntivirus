VERSION 5.00
Begin VB.Form MainGUI 
   BackColor       =   &H00D7E8D3&
   Caption         =   "EZUSBAntiVirus"
   ClientHeight    =   4800
   ClientLeft      =   5655
   ClientTop       =   2910
   ClientWidth     =   8745
   Icon            =   "MainGUI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8745
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame3 
      Appearance      =   0  '平面
      BackColor       =   &H00D7E8D3&
      Caption         =   "緊急救援"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "建立AutoRun(&A)"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BackColor       =   &H00D7E8D3&
      Caption         =   "Log"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   7335
      Begin VB.TextBox text1 
         Appearance      =   0  '平面
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   4
         Top             =   480
         Width           =   7095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "Log(&L)："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H00D7E8D3&
      Caption         =   "執行動作"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
      Begin VB.DriveListBox AntivirusAt 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  '平面
         Caption         =   "防毒(&F)"
         Height          =   300
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   120
      X2              =   7440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "MainGUI.frx":10CA
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "MainGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub log(Text)
text1.Text = text1.Text & vbCrLf & Now & ":  " & Text
End Sub


Private Sub Command1_Click()
log ("Scan:If Not UCase(Left(AntivirusAt.Drive, 1)) = UCase('c') Then")
If Not UCase(Left(AntivirusAt.Drive, 1)) = UCase("c") Then
    log ("Start:")
    On Error GoTo Line
        FileCopy "C:\autorun.ini", UCase(Left(AntivirusAt.Drive, 1)) & ":\"
        Exit Sub
Line:
    log ("Error:未知問題,On Error GoTo Line內部,根目錄內是否有autorun.ini")
    MsgBox "錯誤：" & vbCrLf & vbCrLf & "未知問題！" & vbCrLf & "請檢查根目錄內是否有autorun.ini，若無請自行建立(不需輸入任何文字)" & vbCrLf & "否則請向企鵝哥回報(可能為軟體問題)", 16, "EZUSBAntiVirus 錯誤回報員"
Else
    log ("Error:磁碟機位於C")
    MsgBox "錯誤！" & vbCrLf & vbCrLf & "你不能在C磁碟機安裝此防毒，" & vbCrLf & "因為該磁碟機為系統磁碟，無法使用", 16, "EZUSBAntiVirus 錯誤回報員"
End If
End Sub

Private Sub Command2_Click()
log ("Create:Autorun.ini")
Shell "cmd.exe /c start " & "autorun.bat"
End Sub

Private Sub Command3_Click()
MsgBox UCase(Left(AntivirusAt.Drive, 1)) & ":\"
End Sub

Private Sub Form_Load()
Me.Width = Image1.Width
text1.Text = Now & ":  Welcome To EZUSBAntivirus"
log ("Open EZUSBAntiVirus")
End Sub

Private Sub Form_Resize()
Me.Width = Image1.Width
End Sub
