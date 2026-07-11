VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAgent 
   Caption         =   "ExcelAI"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.TextBox.1 txtChat 
      Height          =   4800
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   9120
   End
   Begin MSForms.TextBox.1 txtInput 
      Height          =   600
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   5280
      Width           =   7680
   End
   Begin MSForms.CommandButton.1 cmdSend 
      Caption         =   "Send"
      Height          =   600
      Left            =   8040
      TabIndex        =   2
      Top             =   5280
      Width           =   1320
   End
   Begin MSForms.Label.1 lblStatus 
      Caption         =   "Ready"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   9120
   End
   Begin MSForms.Label.1 lblAutoRun 
      Caption         =   "AutoRun"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6480
      Width           =   4560
   End
   Begin MSForms.Label.1 lblWorkspace 
      Caption         =   "Workspace"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   6840
      Width           =   9120
   End
End
Attribute VB_Name = "frmAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_IsRunning As Boolean

Private Sub UserForm_Initialize()
    InitDefaultSettings
    lblWorkspace.Caption = "Workspace: " & GetWorkspacePath()
    lblAutoRun.Caption = "AutoRun: " & IIf(IsAutoRunEnabled(), "ON", "OFF") & " / ConfirmDanger: " & IIf(IsConfirmDangerEnabled(), "ON", "OFF")
    AppendChat "system", "ExcelAI v" & APP_VERSION
End Sub

Private Sub cmdSend_Click()
    SendMessage
End Sub

Private Sub SendMessage()
    Dim prompt As String

    If m_IsRunning Then
        MsgBox "Busy. Please wait.", vbExclamation
        Exit Sub
    End If

    prompt = Trim$(txtInput.Text)
    If Len(prompt) = 0 Then Exit Sub

    AppendChat "user", prompt
    txtInput.Text = ""
    m_IsRunning = True
    cmdSend.Enabled = False

    Agent_RunPrompt prompt, "frmAgent.Agent_Status"
End Sub

Public Sub Agent_Status(ByVal msg As String)
    lblStatus.Caption = msg
    DoEvents
End Sub

Public Sub Agent_Status_Complete(ByVal msg As String)
    AppendChat "assistant", msg
    lblStatus.Caption = "Ready"
    m_IsRunning = False
    cmdSend.Enabled = True
End Sub

Private Sub AppendChat(ByVal role As String, ByVal text As String)
    Dim prefix As String
    Select Case role
        Case "user": prefix = "You: "
        Case "assistant": prefix = "AI: "
        Case "system": prefix = "[system] "
        Case Else: prefix = role & ": "
    End Select
    txtChat.Text = txtChat.Text & prefix & text & vbCrLf & vbCrLf
End Sub
