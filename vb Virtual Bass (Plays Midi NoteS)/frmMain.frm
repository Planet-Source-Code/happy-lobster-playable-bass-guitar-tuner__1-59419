VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Virtual Bassist"
   ClientHeight    =   2640
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   16065
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   176
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1071
   Begin VB.ComboBox cmbInstrument 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmMain.frx":6E95E
      Left            =   60
      List            =   "frmMain.frx":6E97D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2220
      Width           =   2235
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   2220
      Width           =   1215
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Line lneFret 
      Index           =   20
      X1              =   1028
      X2              =   1032
      Y1              =   116
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   19
      X1              =   1004
      X2              =   1012
      Y1              =   120
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   18
      X1              =   976
      X2              =   984
      Y1              =   116
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   17
      X1              =   948
      X2              =   956
      Y1              =   120
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   16
      X1              =   920
      X2              =   924
      Y1              =   116
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   15
      X1              =   888
      X2              =   892
      Y1              =   116
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   14
      X1              =   852
      X2              =   852
      Y1              =   116
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   13
      X1              =   816
      X2              =   820
      Y1              =   112
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   12
      X1              =   776
      X2              =   780
      Y1              =   112
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   11
      X1              =   736
      X2              =   740
      Y1              =   104
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   10
      X1              =   688
      X2              =   692
      Y1              =   108
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   9
      X1              =   640
      X2              =   644
      Y1              =   104
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   8
      X1              =   588
      X2              =   592
      Y1              =   104
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   7
      X1              =   536
      X2              =   540
      Y1              =   104
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   6
      X1              =   478
      X2              =   483
      Y1              =   100
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   5
      X1              =   416
      X2              =   420
      Y1              =   100
      Y2              =   28
   End
   Begin VB.Line lneFret 
      Index           =   4
      X1              =   352
      X2              =   356
      Y1              =   100
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   3
      X1              =   280
      X2              =   284
      Y1              =   96
      Y2              =   24
   End
   Begin VB.Line lneFret 
      Index           =   2
      X1              =   208
      X2              =   208
      Y1              =   104
      Y2              =   16
   End
   Begin VB.Line lneFret 
      Index           =   1
      X1              =   125
      X2              =   125
      Y1              =   16
      Y2              =   92
   End
   Begin VB.Line lneFret 
      Index           =   0
      X1              =   44
      X2              =   44
      Y1              =   96
      Y2              =   12
   End
   Begin VB.Line lneA 
      Tag             =   "A-1"
      X1              =   42
      X2              =   1032
      Y1              =   64
      Y2              =   84
   End
   Begin VB.Line lneG 
      Tag             =   "G-2"
      X1              =   42
      X2              =   1036
      Y1              =   28
      Y2              =   32
   End
   Begin VB.Line lneD 
      Tag             =   "D-2"
      X1              =   42
      X2              =   1032
      Y1              =   48
      Y2              =   60
   End
   Begin VB.Line lneE 
      Tag             =   "E-1"
      X1              =   40
      X2              =   1056
      Y1              =   84
      Y2              =   108
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'http://www.harmony-central.com/MIDI/Doc/table2.html - Midi note numbers

Option Explicit
Private objMidi As New MIDILib ' Midi
Private sStrings(1600, 400) As Byte ' Store where the strings are
Private sFrets(1600, 400) As Byte ' Store where notes are

Private Const iThickness = 10

Private iFret As Integer
Private iStringNote As Integer
Private iNotePlaying As Integer


Private Sub cmbInstrument_Click()
    objMidi.Instrument = Val(Right(cmbInstrument.Text, 2))
End Sub

Private Sub cmdQuit_Click()
    objMidi.CloseMIDI
    End
End Sub


Private Sub Setup()
    StoreString lneE
    StoreString lneA
    StoreString lneD
    StoreString lneG
    StoreFrets
End Sub

Private Sub ToggleLines()
    Dim ctl As Control
    
    For Each ctl In Me
        If (TypeOf ctl Is Line) Then
            ctl.Visible = Not ctl.Visible
        End If
    Next
End Sub

Private Sub StoreString(ByVal BassString As Line)
    Dim dGradient As Double
    Dim X As Integer
    Dim Xn As Integer
    Dim Yn As Integer

    
    With BassString
        dGradient = (.Y1 - .Y2) / (.X1 - .X2)
        If .X2 > .X1 Then
            For X = 0 To .X2 - .X1
                Xn = .X1 + X
                Yn = (X * dGradient) + .Y1
                
                StoreStr Xn, Yn, GetNoteNumber(.Tag)
            Next
        Else
            For X = 0 To .X1 - .X2
                Xn = .X2 + X
                Yn = (X * dGradient) + .Y1
            
                StoreStr Xn, Yn, GetNoteNumber(.Tag)
            Next

        End If
    End With
End Sub

Private Sub StoreFret(ByVal X As Integer, ByVal Y As Integer, ByVal Fret As Integer)
    Dim R As Integer
    
    'Me.PSet (X, Y), vbRed
    
    sFrets(X, Y) = Fret
    
    For R = 1 To iThickness
        sFrets(X - R, Y) = Fret
        sFrets(X + R, Y) = Fret
    Next
End Sub

Private Sub StoreStr(ByVal X As Integer, ByVal Y As Integer, ByVal Note As Integer)
    Dim R As Integer
    'Me.PSet (X, Y), vbRed
    
    sStrings(X, Y) = Note
    
    For R = 1 To iThickness
        sStrings(X, Y - R) = Note
        sStrings(X, Y + R) = Note
    Next
End Sub

Private Function GetNoteNumber(ByVal S As String)
    Select Case S
    Case "E-1"
        GetNoteNumber = 28
    Case "A-1"
        GetNoteNumber = 33
    Case "D-2"
        GetNoteNumber = 38
    Case "G-2"
        GetNoteNumber = 43
    End Select
End Function

Private Function GetNote(ByVal NoteNumber As String)
    Dim iOctave As Integer

    iOctave = Int(NoteNumber / 12) - 1

    If NoteNumber Mod 12 = 0 Then
        'The note is a C
        NoteNumber = 0
    Else
        NoteNumber = NoteNumber - ((iOctave + 1) * 12)
    End If


    Select Case NoteNumber
    Case 0
        GetNote = "C"
    Case 1
        GetNote = "C#"
    Case 2
        GetNote = "D"
    Case 3
        GetNote = "D#"
    Case 4
        GetNote = "E"
    Case 5
        GetNote = "F"
    Case 6
        GetNote = "F#"
    Case 7
        GetNote = "G"
    Case 8
        GetNote = "G#"
    Case 9
        GetNote = "A"
    Case 10
        GetNote = "A#"
    Case 10
        GetNote = "B"
    End Select
    
    If GetNote <> "" Then GetNote = GetNote & "-" & CStr(iOctave)
End Function

Private Sub StoreFrets()
    Dim Fret As Integer
    Dim Y As Integer
    Dim Xn As Integer
    Dim Yn As Integer
    Dim dGradient As Double
    
    For Fret = 0 To 20
        With lneFret(Fret)
            dGradient = (.X2 - .X1) / (.Y2 - .Y1)
        
            If .Y2 > .Y1 Then
                For Y = 0 To .Y2 - .Y1
                    Xn = (Y * dGradient) + .X1
                    Yn = .Y1 + Y
                
                    StoreFret Xn, Yn, Fret + 1
                Next
            Else
                For Y = 0 To .Y1 - .Y2
                    Xn = (Y * dGradient) + .X2
                    Yn = .Y2 + Y
                    
                    StoreFret Xn, Yn, Fret + 1
                Next

            End If
        End With
    Next
End Sub







Private Sub Form_Load()
    Dim blnConnected As Boolean
    
    blnConnected = objMidi.ConnectMIDI
    If blnConnected = False Then
        MsgBox "Error connecting to MIDI Mapper device", vbCritical, "Critical Error"
        Exit Sub
    End If
    
    objMidi.Instrument = 32
    objMidi.BaseNote = 0
    
    Setup
    ToggleLines
    
    MsgBox "Left click to play a note on the fret" & vbCrLf & "Right click plays an open string" & vbCrLf & "Press quit to exit", vbInformation
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub PlayNote(ByVal Note As Integer)
    If Note <> iNotePlaying Then
        objMidi.StopNote iNotePlaying
        objMidi.StartNote Note
        iNotePlaying = Note
    End If
    
   
End Sub

Private Sub StopNote()
    If iNotePlaying <> 0 Then
        objMidi.StopNote iNotePlaying
    End If
    iNotePlaying = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lne As Line
    
    lneE.BorderWidth = 1
    lneA.BorderWidth = 1
    lneD.BorderWidth = 1
    lneG.BorderWidth = 1
    
    lneE.Visible = False
    lneA.Visible = False
    lneD.Visible = False
    lneG.Visible = False
    
    If iFret > 0 Then lneFret(iFret - 1).BorderWidth = 1: lneFret(iFret - 1).Visible = False
    
    If X >= 0 And X <= 1600 And Y >= 0 And Y <= 400 Then
        iFret = sFrets(X, Y)
        iStringNote = sStrings(X, Y)
        
        Select Case GetNote(iStringNote)
        Case "E-1"
            lneE.BorderWidth = iThickness
            lneE.Visible = True
        Case "A-1"
            lneA.BorderWidth = iThickness
            lneA.Visible = True
        Case "D-2"
            lneD.BorderWidth = iThickness
            lneD.Visible = True
        Case "G-2"
            lneG.BorderWidth = iThickness
            lneG.Visible = True
        End Select
        
        If iFret > 0 Then lneFret(iFret - 1).BorderWidth = iThickness: lneFret(iFret - 1).Visible = True
        
    End If
    
    lblNote = GetNote(iStringNote + iFret - 1) & "   fret: " & CStr(iFret - 1)
    
    
    If Button = 1 Then
        If iStringNote > 0 And iFret > 0 Then
            PlayNote iStringNote + iFret - 1
            Debug.Print "play fret"
        End If
    ElseIf Button = 2 Then
        If iStringNote > 0 Then
            PlayNote iStringNote
            Debug.Print "play open"
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StopNote
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objMidi.CloseMIDI
End Sub


