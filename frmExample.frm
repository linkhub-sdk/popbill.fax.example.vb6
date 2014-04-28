VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 팩스 SDK 예제"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows 기본값
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팩스 전송 관련 "
      Height          =   6975
      Left            =   120
      TabIndex        =   16
      Top             =   2715
      Width           =   9375
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   450
         Left            =   6810
         TabIndex        =   27
         Top             =   1515
         Width           =   2355
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "전송내역 확인"
         Height          =   450
         Left            =   4320
         TabIndex        =   26
         Top             =   1515
         Width           =   2355
      End
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4725
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   2100
         Width           =   9090
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1200
         TabIndex        =   24
         Top             =   1575
         Width           =   2835
      End
      Begin VB.CommandButton btnSendFax_Multi_Same 
         Caption         =   "다수파일 동보전송"
         Height          =   570
         Left            =   5280
         TabIndex        =   22
         Top             =   720
         Width           =   1875
      End
      Begin VB.CommandButton btnSendFAX_Multi 
         Caption         =   "다수 파일 전송"
         Height          =   570
         Left            =   3600
         TabIndex        =   21
         Top             =   720
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFax_Same 
         Caption         =   "동보 전송"
         Height          =   570
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFAX 
         Caption         =   "전송"
         Height          =   570
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1590
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3420
         TabIndex        =   18
         Top             =   255
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   300
         TabIndex        =   23
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약전송 시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   195
         TabIndex        =   17
         Top             =   330
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   9375
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   1575
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   2160
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "전송 단가 확인"
            Height          =   495
            Left            =   150
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 파트너 관련"
         Height          =   1575
         Left            =   4410
         TabIndex        =   8
         Top             =   360
         Width           =   2535
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   270
            Width           =   1815
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여 포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 기타"
         Height          =   1575
         Left            =   7035
         TabIndex        =   5
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   " 팝빌 기본 URL 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox cboPopbillTOGO 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Text            =   "LOGIN"
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Text            =   "1231212312"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌아이디 : "
      Height          =   180
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "사업자번호 : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

'파트너아이디
Private Const PartnerID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "088b1258aoeMH5OtGjK4zaOlwZGVvSK40ceI8t4j7Hw="

Private FaxService As New PBFAXService


Private Sub btnCancelReserve_Click()
 Dim response As PBResponse
    
    Set response = FaxService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim response As PBResponse
    
    Set response = FaxService.CheckIsMember(txtCorpNum.Text, PartnerID)
    
    If response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (response.message)
End Sub


Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
    
End Sub

Private Sub btnGetFaxDetail_Click()
    Dim sentFaxList As Collection
    
    Set sentFaxList = FaxService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If sentFaxList Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    
    Dim sentFax As PBFaxInfo
    
    
    Dim tmp As String
    tmp = "sendState | convState | sendnum | rcv | rcvnm | T | S | F | R | C | reserveDT | sendDT | resultDT | sendResult" + vbCrLf
    
    For Each sentFax In sentFaxList
    
        tmp = tmp + CStr(sentFax.sendState) + " | "
        tmp = tmp + CStr(sentFax.convState) + " | "
        tmp = tmp + sentFax.sendNum + " | "
        tmp = tmp + sentFax.receiveNum + " | "
        tmp = tmp + sentFax.receiveName + " | "
        
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "
        
        tmp = tmp + sentFax.reserveDT + " | "
        tmp = tmp + sentFax.sendDT + " | "
        tmp = tmp + sentFax.resultDT + " | "
     
        tmp = tmp + CStr(sentFax.sendResult)
        
        tmp = tmp + vbCrLf
    Next
    
    
    txtResult.Text = tmp
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, cboPopbillTOGO.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim response As PBResponse
    
    joinData.PartnerID = PartnerID '파트너 아이디
    joinData.CorpNum = "1231212312" '사업자번호 "-" 제외.
    joinData.CEOName = "대표자성명"
    joinData.CorpName = "회원상호"
    joinData.Addr = "주소"
    joinData.ZipCode = "500-100"
    joinData.BizType = "업태"
    joinData.BizClass = "업종"
    joinData.ID = "userid"      '6자 이상 20자 미만.
    joinData.PWD = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set response = FaxService.JoinMember(joinData)
    
    If response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (response.message)
    
    
End Sub

Private Sub btnSendFAX_Click()
    '전송파일경로 목록
    Dim FilePaths As New Collection
    
    CommonDialog1.FileName = ""
    
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    
    FilePaths.Add CommonDialog1.FileName
    
    '수신자 목록
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    
    receiver.receiverNum = "00001111"
    receiver.receiverName = "수신자 명칭"
    
    receivers.Add receiver
    
    Dim ReceiptNum As String
    
    
    ReceiptNum = FaxService.SendFAX(txtCorpNum.Text, "07075106766", receivers, FilePaths, txtReserveDT.Text, txtUserID.Text)
    
    
     If ReceiptNum = "" Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendFAX_Multi_Click()
    '전송파일경로 목록
    Dim FilePaths As New Collection
    
    Do
        CommonDialog1.FileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.FileName <> "" Then
            FilePaths.Add CommonDialog1.FileName
        End If
    
    Loop While (CommonDialog1.FileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '수신자 목록
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    
    receiver.receiverNum = "00001111"
    receiver.receiverName = "수신자 명칭"
    
    receivers.Add receiver
    
    Dim ReceiptNum As String
    
    ReceiptNum = FaxService.SendFAX(txtCorpNum.Text, "07075106766", receivers, FilePaths, txtReserveDT.Text, txtUserID.Text)
    
    
     If ReceiptNum = "" Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFax_Multi_Same_Click()
    
    '전송파일경로 목록
    Dim FilePaths As New Collection
    
    Do
        CommonDialog1.FileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.FileName <> "" Then
            FilePaths.Add CommonDialog1.FileName
        End If
    
    Loop While (CommonDialog1.FileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '동보 수신자 목록
    Dim receivers As New Collection
    Dim receiver As PBReceiver
    Dim i As Integer
    
    '최대 1000명까지 가능
    For i = 1 To 100
        Set receiver = New PBReceiver
        receiver.receiverNum = "00001111"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = FaxService.SendFAX(txtCorpNum.Text, "07075106766", receivers, FilePaths, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendFax_Same_Click()
    
    '전송파일경로 목록
    Dim FilePaths As New Collection
    
    CommonDialog1.FileName = ""
    
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    
    FilePaths.Add CommonDialog1.FileName
    
    '동보 수신자 목록
    Dim receivers As New Collection
    Dim receiver As PBReceiver
    Dim i As Integer
    
    '최대 1000명까지 가능
    For i = 1 To 100
        Set receiver = New PBReceiver
        receiver.receiverNum = "00001111"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    Dim ReceiptNum As String
    
    
    ReceiptNum = FaxService.SendFAX(txtCorpNum.Text, "07075106766", receivers, FilePaths, txtReserveDT.Text, txtUserID.Text)
    
    
     If ReceiptNum = "" Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = FaxService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "전송 단가 : " + CStr(unitCost)
End Sub

Private Sub Form_Load()
    FaxService.Initialize PartnerID, SecretKey
    FaxService.IsTest = True
    
    cboPopbillTOGO.AddItem "LOGIN"
    cboPopbillTOGO.AddItem "CHRG"
    
End Sub
