VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 팩스 SDK 예제"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   11490
   ScaleWidth      =   11460
   StartUpPosition =   2  '화면 가운데
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팩스 전송 관련 "
      Height          =   7215
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   10695
      Begin VB.CommandButton btnSearchPopUp 
         Caption         =   "전송내역조회 팝업"
         Height          =   465
         Left            =   8505
         TabIndex        =   24
         Top             =   210
         Width           =   1815
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   450
         Left            =   7410
         TabIndex        =   23
         Top             =   1515
         Width           =   2355
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "전송내역 확인"
         Height          =   450
         Left            =   4920
         TabIndex        =   22
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
         Left            =   705
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2100
         Width           =   9090
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   1575
         Width           =   2835
      End
      Begin VB.CommandButton btnSendFax_Multi_Same 
         Caption         =   "다수파일 동보전송"
         Height          =   570
         Left            =   6360
         TabIndex        =   18
         Top             =   720
         Width           =   1875
      End
      Begin VB.CommandButton btnSendFAX_Multi 
         Caption         =   "다수 파일 전송"
         Height          =   570
         Left            =   4680
         TabIndex        =   17
         Top             =   720
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFax_Same 
         Caption         =   "동보 전송"
         Height          =   570
         Left            =   3000
         TabIndex        =   16
         Top             =   720
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFAX 
         Caption         =   "전송"
         Height          =   570
         Left            =   1320
         TabIndex        =   15
         Top             =   720
         Width           =   1590
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   4500
         TabIndex        =   14
         Top             =   255
         Width           =   2835
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   900
         TabIndex        =   19
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약전송 시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   1080
         TabIndex        =   13
         Top             =   330
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2895
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   10695
      Begin VB.Frame Frame7 
         Caption         =   " 회사정보 관련 "
         Height          =   2295
         Left            =   8520
         TabIndex        =   33
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   2295
         Left            =   6480
         TabIndex        =   29
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   32
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   495
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   2295
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   2295
         Left            =   1920
         TabIndex        =   7
         Top             =   360
         Width           =   2505
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   2175
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "전송 단가 확인"
            Height          =   495
            Left            =   150
            TabIndex        =   8
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL "
         Height          =   2295
         Left            =   4560
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6120
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   285
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2415
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   300
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit

'연동아이디
Private Const linkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private FaxService As New PBFAXService


Private Sub btnCancelReserve_Click()
 Dim Response As PBResponse
    
    Set Response = FaxService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
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

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = FaxService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
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

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID '링크 아이디
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
    
    Set Response = FaxService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = FaxService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.ID = "testkorea_20151007"      '담당자 아이디
    joinData.PWD = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = FaxService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnSearchPopUp_Click()
    Dim url As String
    
    url = FaxService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
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

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
    Set Response = FaxService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.CEOName = "대표자"         '대표자명
    CorpInfo.CorpName = "상호_수정"          '상호명
    CorpInfo.Addr = "서울특별시"        '주소
    CorpInfo.BizType = "업태"           '업태
    CorpInfo.BizClass = "업종"          '업종
    
    Set Response = FaxService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub Form_Load()
    FaxService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    FaxService.IsTest = True
        
End Sub
