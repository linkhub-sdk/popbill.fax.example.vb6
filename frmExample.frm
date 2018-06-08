VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 팩스 SDK 예제"
   ClientHeight    =   11910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   11910
   ScaleWidth      =   15795
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
      Height          =   8175
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   13455
      Begin VB.Frame Frame9 
         Caption         =   "발신번호 관리"
         Height          =   1575
         Left            =   10320
         TabIndex        =   38
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetURL_SENDER 
            Caption         =   "발신번호 관리 팝업"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "발신번호 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "전송내역 검색조회"
         Height          =   465
         Left            =   8025
         TabIndex        =   33
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton btnSearchPopUp 
         Caption         =   "전송내역조회 팝업"
         Height          =   465
         Left            =   8025
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "부가기능"
         Height          =   1575
         Left            =   7800
         TabIndex        =   37
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton btnResendFaxSame 
         Caption         =   "동보 재전송"
         Height          =   450
         Left            =   1920
         TabIndex        =   36
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton btnResendFAX 
         Caption         =   "재전송"
         Height          =   450
         Left            =   360
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   450
         Left            =   6120
         TabIndex        =   23
         Top             =   2115
         Width           =   1515
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "전송내역 확인"
         Height          =   450
         Left            =   4440
         TabIndex        =   22
         Top             =   2115
         Width           =   1515
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
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2760
         Width           =   12810
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1320
         TabIndex        =   20
         Top             =   2175
         Width           =   2835
      End
      Begin VB.CommandButton btnSendFax_Multi_Same 
         Caption         =   "다수파일 동보전송"
         Height          =   450
         Left            =   5280
         TabIndex        =   18
         Top             =   840
         Width           =   1875
      End
      Begin VB.CommandButton btnSendFAX_Multi 
         Caption         =   "다수 파일 전송"
         Height          =   450
         Left            =   3600
         TabIndex        =   17
         Top             =   840
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFax_Same 
         Caption         =   "동보 전송"
         Height          =   450
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFAX 
         Caption         =   "전송"
         Height          =   450
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   1470
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3600
         TabIndex        =   14
         Top             =   375
         Width           =   3555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   420
         TabIndex        =   19
         Top             =   2250
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약전송 시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   450
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   15375
      Begin VB.Frame Frame11 
         Caption         =   "파트너과금 포인트"
         Height          =   1935
         Left            =   12840
         TabIndex        =   42
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "연동과금 포인트"
         Height          =   1935
         Left            =   10680
         TabIndex        =   41
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " 회사정보 관련 "
         Height          =   1935
         Left            =   8640
         TabIndex        =   30
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   1935
         Left            =   6600
         TabIndex        =   26
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련 "
         Height          =   1935
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2505
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "전송 단가 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL "
         Height          =   1935
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   410
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
'=========================================================================
'
' 팝빌 팩스 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 : http://blog.linkhub.co.kr/569
' - 업데이트 일자 : 2017-08-30
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'팩스 서비스 객체 생성
Private FaxService As New PBFAXService

'=========================================================================
' 예약전송 팩스요청건을 취소합니다.
' - 예약전송 취소는 예약전송시간 10분전까지 가능합니다.
'=========================================================================

Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 연동회원의 팩스 API 서비스 과금정보를 확인합니다.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = FaxService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (전송단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = FaxService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자성명) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (상호) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팩스 전송요청시 반환받은 접수번호(receiptNum)을 사용하여 팩스전송
' 결과를 확인합니다.
'=========================================================================

Private Sub btnGetFaxDetail_Click()
    Dim sentFaxList As Collection
    Dim i As Integer
    Dim fileName As Variant
    Dim sentFax As PBFaxInfo
    Dim tmp As String
    
    Set sentFaxList = FaxService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
    
    If sentFaxList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state | result | title | sendnum | senderName | rcv | rcvnm | T | S | F | R | C | receiptDT | reserveDT | sendDT | resultDT | filenames" + vbCrLf
    
    For Each sentFax In sentFaxList
    
        tmp = tmp + CStr(sentFax.state) + " | "             '전송상태 코드
        tmp = tmp + CStr(sentFax.result) + " | "            '전송결과 코드
        tmp = tmp + sentFax.title + " | "                   '팩스제목
        tmp = tmp + sentFax.sendNum + " | "                 '발신번호
        tmp = tmp + sentFax.senderName + " | "              '발신자명
        tmp = tmp + sentFax.receiveNum + " | "              '수신번호
        tmp = tmp + sentFax.receiveName + " | "             '수신자명
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "       '전체 페이지수
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "    '성공 페이지수
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "       '실패 페이지수
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "     '환불 페이지수
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "     '취소 페이지수
        
        tmp = tmp + CStr(sentFax.receiptDT) + " | "         '접수일시
        tmp = tmp + sentFax.reserveDT + " | "               '예약일시
        tmp = tmp + sentFax.sendDT + " | "                  '전송일시
        tmp = tmp + sentFax.resultDT + " | "                '전송결과 수신일시
     
        i = 0
        
        For Each fileName In sentFax.fileNames              '팩스전송 파일명
            i = i + 1
            If sentFax.fileNames.Count = i Then
                tmp = tmp + fileName
            Else
                tmp = tmp + fileName + ", "
            End If
        Next
        
        tmp = tmp + vbCrLf
    Next
    
    
    txtResult.Text = tmp
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 파트너 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팩스 발신번호 목록을 조회합니다.
'=========================================================================

Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim SenderNumber As PBFaxSenderNumber
    
    Set SenderNumberList = FaxService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumber In SenderNumberList
        tmp = tmp + "발신번호(number) : " + SenderNumber.number + vbCrLf
        tmp = tmp + "대표번호 지정여부(representYN) : " + CStr(SenderNumber.representYN) + vbCrLf
        tmp = tmp + "등록상태(state) : " + CStr(SenderNumber.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 팩스 발신번호 관리 팝업 URL을 반환합니다.
' 보안정책으로 인해 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetURL_SENDER_Click()
    Dim url As String
    
    url = FaxService.GetURL(txtCorpNum.Text, txtUserID.Text, "SENDER")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 연동회원 가입을 요청합니다.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1231212312"
    
    '대표자성명, 최대 30자
    joinData.CEOName = "대표자성명"
    
    '상호명, 최대 70자
    joinData.CorpName = "회원상호"
    
    '주소, 최대 300자
    joinData.Addr = "주소"
    
    '업태, 최대 40자
    joinData.BizType = "업태"
    
    '종목, 최대 40자
    joinData.BizClass = "종목"
    
    '아이디, 6자이상 20자 미만
    joinData.ID = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.PWD = "pwd_must_be_long_enough"
    
    '담당자명, 최대 30자
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    '담당자 메일, 최대 70자
    joinData.ContactEmail = "test@test.com"
    
    Set Response = FaxService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = FaxService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 20자 미만
    joinData.ID = "testkorea_20161011"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.PWD = "test@test.com"
    
    '담당자명, 최대 30자
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
    
    '담당자 팩스번호
    joinData.fax = "070-1234-1234"
    
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
        
    Set Response = FaxService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 팩스를 재전송합니다.
' - 전송일로부터 180일이 경과되지 않은 건만 재전송할 수 있습니다.
' - 발신자/수신자 정보를 수정하여 전송할 수 있습니다.
'=========================================================================

Private Sub btnResendFAX_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    
    ' 발신번호, 공백처리시 기존발신번호로 재전송
    senderNum = ""
    
    ' 발신자명, 공백처리시 기존발신자명으로 재전송
    senderName = ""
    
    ' 기존수신정보 변경없이 재전송하는 경우, receivers(수신정보) Collection 을 Nothing 으로 선언
    Set receivers = Nothing
    
    
    ' 새로운 수신정보로 재전송하는 경우, 수신번호/수신자명을 기재하여 receivers Collection에 추가
    ' 수신번호
    'receiver.receiverNum = "0700000214"
    
    ' 수신자명
    'receiver.receiverName = "수신자_수정"
    
    ' 수신정보 Collection 추가
    'receivers.Add receiver
    
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 팩스를 재전송합니다.
' - 전송일로부터 180일이 경과되지 않은 건만 재전송할 수 있습니다.
' - 발신자/수신자 정보를 수정하여 전송할 수 있습니다.
'=========================================================================

Private Sub btnResendFaxSame_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim i As Integer
    
    ' 발신번호, 공백처리시 기존발신번호로 재전송
    senderNum = ""
    
    ' 발신자명, 공백처리시 기존발신자명으로 재전송
    senderName = ""
    
    ' 기존수신정보 변경없이 재전송하는 경우, receivers(수신정보) Collection 을 Nothing 으로 선언
    'Set receivers = Nothing
    
    
    ' 새로운 수신정보로 재전송하는 경우, 수신번호/수신자명을 기재하여 receivers Collection에 추가
    ' 수신정보, 최대 1000건
    For i = 1 To 10
        Set receiver = New PBReceiver
        receiver.receiverNum = "010111222"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 검색조건을 사용하여 팩스전송 내역을 조회합니다.
'=========================================================================

Private Sub btnSearch_Click()
    Dim faxSearchList As PBFaxSearchList
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim ReserveYN As Boolean
    Dim SenderOnly As Boolean
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim fileName As Variant
    Dim i As Integer
    Dim tmp As String
    Dim sentFax As PBFaxInfo
    
    '[필수] 시작일자, 형식(yyyyMMdd)
    SDate = "20170601"
    
    '[필수] 종료일자, 형식(yyyyMMdd)
    EDate = "20171231"
    
    '전송상태 배열, 1(대기), 2(성공), 3(실패), 4(취소)
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '예약전송 검색여부, True-예약전송건 조회, False-전체조회
    ReserveYN = False
    
    '개인조회 여부, True-개인조회, False-회사조회
    SenderOnly = False
    
    '페이지 번호, 기본값 1
    Page = 1
    
    '페이지당 목록갯수, 기본값 500
    PerPage = 30
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    Set faxSearchList = FaxService.Search(txtCorpNum.Text, SDate, EDate, state, ReserveYN, SenderOnly, Page, PerPage, Order)
     
    If faxSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (응답코드) : " + CStr(faxSearchList.code) + vbCrLf
    tmp = tmp + "total (총 검색결과 건수) : " + CStr(faxSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 목록개수) : " + CStr(faxSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(faxSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(faxSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + faxSearchList.message + vbCrLf + vbCrLf
    
    MsgBox tmp
    
    tmp = "state | result | title | sendnum | senderName | rcv | rcvnm | T | S | F | R | C | receiptDT | reserveDT | sendDT | resultDT | fileNames" + vbCrLf
    
    For Each sentFax In faxSearchList.list
    
        tmp = tmp + CStr(sentFax.state) + " | "             '전송상태 코드
        tmp = tmp + CStr(sentFax.result) + " | "            '전송결과 코드
        tmp = tmp + sentFax.title + " | "                   '팩스제목
        
        tmp = tmp + sentFax.sendNum + " | "                 '발신번호
        tmp = tmp + sentFax.senderName + " | "              '발신번호
        tmp = tmp + sentFax.receiveNum + " | "              '수신번호
        tmp = tmp + sentFax.receiveName + " | "             '수신자명
        
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "       '페이지수
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "    '성공 페이지수
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "       '실패 페이지수
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "     '환불 페이지수
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "     '취소 페이지수
        
        tmp = tmp + sentFax.receiptDT + " | "               '접수일시
        tmp = tmp + sentFax.reserveDT + " | "               '예약전송일시
        tmp = tmp + sentFax.sendDT + " | "                  '전송일시
        tmp = tmp + sentFax.resultDT + " | "                '전송결과 수신일시
        
        i = 0
        
        For Each fileName In sentFax.fileNames              '전송 파일명
            i = i + 1
            If sentFax.fileNames.Count = i Then
                tmp = tmp + fileName
            Else
                tmp = tmp + fileName + ", "
            End If
        Next
        
        tmp = tmp + vbCrLf
    
    Next
    
    
    txtResult.Text = tmp
    
End Sub

'=========================================================================
' 팩스 전송내역 목록 팝업 URL을 반환합니다.
' 보안정책으로 인해 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnSearchPopUp_Click()
    Dim url As String
    
    url = FaxService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnSendFAX_Click()
    Dim FilePaths As New Collection
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim adsYN As Boolean
    Dim title As String
    
    CommonDialog1.fileName = ""
    
    CommonDialog1.ShowOpen
    
    If CommonDialog1.fileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.fileName
    
    '발신번호
    senderNum = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '수신번호
    receiver.receiverNum = "070111222"
    
    '수신자명
    receiver.receiverName = "수신자 명칭"
    receivers.Add receiver
    
    '광고팩스 전송여부
    adsYN = False
    
    '팩스제목
    title = "팩스 단건전송 제목"
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendFAX_Multi_Click()
    Dim FilePaths As New Collection
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim senderNum As String
    Dim senderName As String
    Dim receiptNum As String
    Dim title As String
    Dim adsYN As Boolean
    
    '전송 파일 개수 최대 20개
    Do
        CommonDialog1.fileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.fileName <> "" Then
            FilePaths.Add CommonDialog1.fileName
        End If
    
    Loop While (CommonDialog1.fileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '발신번호
    senderNum = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '수신번호
    receiver.receiverNum = "070111222"
    
    '수신자명
    receiver.receiverName = "수신자 명칭"
    
    receivers.Add receiver
    
    '광고팩스 전송여부
    adsYN = False
    
    '팩스제목
    title = "팩스 단건 다수파일 팩스제목"
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

Private Sub btnSendFax_Multi_Same_Click()
    Dim FilePaths As New Collection
    Dim receivers As New Collection
    Dim senderNum As String
    Dim senderName As String
    Dim receiptNum As String
    Dim title As String
    Dim receiver As PBReceiver
    Dim i As Integer
    Dim adsYN As Boolean
    
    '전송 파일 개수 최대 20개
    Do
        CommonDialog1.fileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.fileName <> "" Then
            FilePaths.Add CommonDialog1.fileName
        End If
    
    Loop While (CommonDialog1.fileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '발신번호
    senderNum = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '수신정보 최대 1000명까지 가능
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    '광고팩스 전송여부
    adsYN = False
    
    '팩스제목
    title = "팩스 다수파일 동보전송 제목"
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

Private Sub btnSendFax_Same_Click()
    Dim FilePaths As New Collection
    Dim receivers As New Collection
    Dim senderNum As String
    Dim senderName As String
    Dim receiptNum As String
    Dim title As String
    Dim receiver As PBReceiver
    Dim adsYN As Boolean
    Dim i As Integer
    
    
    CommonDialog1.fileName = ""
    
    CommonDialog1.ShowOpen
    
    If CommonDialog1.fileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.fileName
        
    '발신번호
    senderNum = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '수신정보, 최대 1000건
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    '광고팩스 전송여부
    adsYN = True
    
    '팩스제목
    title = "팩스 동보전송 제목"
                
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 팩스 전송단가를 확인합니다.
'=========================================================================

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = FaxService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.ID = txtUserID.Text
    
    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-4304-2991"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, Ture-회사조회, False-개인조
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
                
    Set Response = FaxService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명
    CorpInfo.CEOName = "대표자"
    
    '상호
    CorpInfo.CorpName = "상호"
    
    '주소
    CorpInfo.Addr = "서울특별시"
    
    '업태
    CorpInfo.BizType = "업태"
    
    '종목
    CorpInfo.BizClass = "종목"
    
    Set Response = FaxService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub


Private Sub Form_Load()
    FaxService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    FaxService.IsTest = True
        
End Sub

