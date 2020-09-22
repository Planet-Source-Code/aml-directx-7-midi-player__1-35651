VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMIDI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Midi Player - Aaron Lindsay's First DirectX Program"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop MIDI"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play MIDI"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtMIDI 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   600
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File To Play:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1380
   End
End
Attribute VB_Name = "frmMIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DX As New DirectX7 'This is the main DX object
Public Loader As DirectMusicLoader 'This loads the midis
Public Preformance As DirectMusicPerformance 'This controls the music
Public Segment As DirectMusicSegment 'This Holds the music in memory

Private Sub cmdBrowse_Click()
    With CD
        .DialogTitle = "Please Choose a MIDI File."
        .Filter = "Midi Files (*.MID, *.RMI)|*.mid; *.rmi"
        .FileName = ""
        .ShowOpen
        If .FileName = "" Then Exit Sub
        If Dir(.FileName) = "" Then Exit Sub
        txtMIDI.Text = .FileName
    End With
    LoadMIDI txtMIDI.Text
End Sub

Private Sub cmdPlay_Click()
    If Dir(txtMIDI.Text) = "" Then Exit Sub
    If txtMIDI.Text = "" Then Exit Sub
    PlayMIDI
End Sub

Private Sub cmdStop_Click()
    StopMIDI
End Sub

Private Sub Form_Load()
    On Error Resume Next
    CD.InitDir = App.Path
    MsgBox "This little midi app was created by Aaron Lindsay with DirectX7." & vbLf & "This is my first DirectX app. Please Visit www.pscode.com to vote!", vbOKOnly, "Midi Program" 'Show the message...
    frmLoad.Show
    Call InitProg 'Initialize DX
    Unload frmLoad
    txtMIDI.Text = App.Path & "\MIDI.mid" 'Show the file in the testbox
    LoadMIDI txtMIDI.Text 'Load the midi
    Call PlayMIDI
End Sub

'This will initialize all DX objects
Sub InitProg()
    Set Loader = DX.DirectMusicLoaderCreate 'This creates the loader
    Set Preformance = DX.DirectMusicPerformanceCreate 'This creates the Preformance
    Preformance.Init Nothing, hWnd 'Initialize the preformance
    Preformance.SetPort -1, 1 'Set the port
    Preformance.SetMasterAutoDownload True
    If Err.Number <> DD_OK Then 'If there's an error, exit the program
        MsgBox "Error: Could not load DirectMusic!", vbExclamation, "ERROR!"
        Unload Me 'Quit
    End If
End Sub

Sub LoadMIDI(FileName As String)
    Set Segment = Loader.LoadSegment(FileName) 'Load the midi
    If Err.Number <> DD_OK Then MsgBox "Error: Could not load MIDI file!", vbExclamation, "ERROR!" 'If there's an error, tell the user.
End Sub

Sub PlayMIDI()
    On Error Resume Next
    Preformance.PlaySegment Segment, 0, 0 'Play the segment
End Sub

Sub StopMIDI()
    On Error Resume Next
        Preformance.Stop Segment, Nothing, 0, 0 'Stop music if it's playing
End Sub
'About-------------------------------------------------------
'Created By: Aaron Lindsay
'Created On: 6/09/2002 11:47 AM
'Purpose Of Code: This code plays a midi file with DirectX 7
'I learned DirectX thanks to Simon Price. Visit his site at:
'http://www.vbgames.co.uk/
'Please Vote at pscode.com!
