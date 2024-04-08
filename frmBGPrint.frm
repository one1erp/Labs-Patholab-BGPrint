VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBGPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "מדפיסון"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3015
   ControlBox      =   0   'False
   Icon            =   "frmBGPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBGPrint.frx":0442
   RightToLeft     =   -1  'True
   ScaleHeight     =   4950
   ScaleMode       =   0  'User
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1944
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Timer timerBgPrint 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1440
      Top             =   720
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      VariousPropertyBits=   19
      Caption         =   "כיבוי"
      Size            =   "2355;873"
      FontEffects     =   1073741825
      FontHeight      =   156
      FontCharSet     =   177
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdDelPrints 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
      VariousPropertyBits=   19
      Caption         =   "בטל הדפסות"
      Size            =   "2355;873"
      FontEffects     =   1073741825
      FontHeight      =   156
      FontCharSet     =   177
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmBGPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'id of the relevant Mutex:
Public lMutexHandle As Long

'needed as s parameter to the CreateMutex() function
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'uses to operate the Mutex:
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Boolean, ByVal lpName As String) As Long
'Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private Sub cmdDelPrints_Click()
    
    Dim lResult As VbMsgBoxResult
    lResult = MsgBox("האם לבטל הדפסת כל הדרישות הממתינות בתחנה זו ?", vbMsgBoxRight + vbMsgBoxRtlReading + vbQuestion + vbYesNo + vbDefaultButton1, "ביטול הדפסות")
    If lResult = vbYes Then
        Call DelPrints
        MsgBox "כל הדרישות שהמתינו להדפסה בוטלו", vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly + vbExclamation, "ביטול הדפסות"
    End If

End Sub

Private Sub cmdExit_Click()
    
    Dim lResult As VbMsgBoxResult
    Dim lstMsgString As String
    
    lstMsgString = "אוי ... לסגור את המדפיסון ?" & vbCrLf & _
             "ליציאה לחץ 'כן', לביטול לחץ 'לא'"
    lResult = MsgBox(lstMsgString, vbMsgBoxRight + vbMsgBoxRtlReading + vbQuestion + vbYesNo + vbDefaultButton1, "ביטול הדפסות")
    If lResult = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    
    Dim x As Long
    Dim lstInterval As String
    Dim lloInterval As Long
    Dim sa As SECURITY_ATTRIBUTES
    
    lstMsg.Clear
    For x = 1 To 49
        lstMsg.AddItem ""
    Next
    
    lstInterval = Mid(Command(), InStr(1, Command(), "|") + 1)

    If lstInterval = "" Or Not IsNumeric(lstInterval) Then
        lloInterval = 10000
    Else
        lloInterval = CLng(lstInterval)
        Select Case lloInterval
        Case Is < 1
            lloInterval = 1
        Case Is > 60
            lloInterval = 65535
        Case Else
            lloInterval = lloInterval * 1000
        End Select
    End If
    
    lstMsg.AddItem "מדפיסון פעיל כל " & CInt(lloInterval / 1000) & " שניות"
    lstMsg.ListIndex = lstMsg.ListCount - 1

    DBConnect
'    Set gWordReport = New Report
'    Set gWordReport.Connection = gCon
    
    Set gSdgLog = CreateObject("SdgLog.CreateLog")
    Set gSdgLog.con = gCon
    gSdgLog.Session = -2
    
    'create the Mutex / get a handle to it:
    lMutexHandle = CreateMutex(sa, False, "BGPRINT_DOCTOPDF")
    
    timerBgPrint.Interval = lloInterval
    timerBgPrint.Enabled = True
End Sub

Private Sub Form_Terminate()

    gCon.Close
    Set gCon = Nothing
    Set gWordReport = Nothing
    Set gGetProcess = Nothing
    Set gSdgLog = Nothing

End Sub

Private Sub lstMsg_DblClick()
    MsgBox lstMsg, vbMsgBoxRight + vbMsgBoxRtlReading + vbInformation + vbOKOnly, "מדפיסון"
End Sub

' This form has only timer.
'   the Timer execute the print every interval

Private Sub timerBgPrint_Timer()
    
    timerBgPrint.Enabled = False
    Me.MousePointer = vbHourglass
    Me.Caption = "מחפש הדפסות ..."
    Me.Refresh
                 
    Call ExecutePrint
    
    Me.Caption = "מדפיסון"
    Me.MousePointer = vbDefault
    Me.Refresh
    timerBgPrint.Enabled = True

End Sub

Private Sub DBConnect()
    
    Dim lstConStr As String
    ' connect to the db
    
    Set gCon = New ADODB.Connection
    If InStr(1, Command(), "|") > 0 Then
        lstConStr = Mid(Command(), 1, InStr(1, Command(), "|") - 1)
    Else
        lstConStr = Command()
    End If
    gCon.Open lstConStr
    gCon.CursorLocation = adUseClient

End Sub

