VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 팩스 SDK 예제"
   ClientHeight    =   13470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   13470
   ScaleWidth      =   15795
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12120
      TabIndex        =   60
      Top             =   285
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팩스 전송 관련 "
      Height          =   8895
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   13455
      Begin VB.Frame Frame13 
         Caption         =   "요청번호 할당 전송건 처리"
         Height          =   1815
         Left            =   4680
         TabIndex        =   48
         Top             =   1920
         Width           =   4335
         Begin VB.TextBox txtRequestNum 
            Height          =   315
            Left            =   1200
            TabIndex        =   54
            Top             =   240
            Width           =   2835
         End
         Begin VB.CommandButton btnResendFaxRNSame 
            Caption         =   "동보 재전송"
            Height          =   450
            Left            =   2280
            TabIndex        =   53
            Top             =   1200
            Width           =   1875
         End
         Begin VB.CommandButton btnResendFAXRN 
            Caption         =   "재전송"
            Height          =   450
            Left            =   240
            TabIndex        =   52
            Top             =   1200
            Width           =   1875
         End
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "예약전송 취소"
            Height          =   450
            Left            =   2280
            TabIndex        =   51
            Top             =   600
            Width           =   1875
         End
         Begin VB.CommandButton btnGetFaxDetailRN 
            Caption         =   "전송내역 확인"
            Height          =   450
            Left            =   245
            TabIndex        =   50
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label5 
            Caption         =   "요청번호 :"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   295
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "발신번호 관리"
         Height          =   1575
         Left            =   9120
         TabIndex        =   37
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "발신번호 관리 팝업"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "발신번호 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "전송내역 검색조회"
         Height          =   495
         Left            =   11400
         TabIndex        =   32
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton btnGetSentListURL 
         Caption         =   "전송내역조회 팝업"
         Height          =   495
         Left            =   11400
         TabIndex        =   23
         Top             =   720
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "부가기능"
         Height          =   2175
         Left            =   11280
         TabIndex        =   36
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetPreviewURL 
            Caption         =   "팩스 미리보기 팝업"
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnResendFaxSame 
         Caption         =   "동보 재전송"
         Height          =   450
         Left            =   2640
         TabIndex        =   35
         Top             =   3120
         Width           =   1875
      End
      Begin VB.CommandButton btnResendFAX 
         Caption         =   "재전송"
         Height          =   450
         Left            =   600
         TabIndex        =   34
         Top             =   3120
         Width           =   1875
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약전송 취소"
         Height          =   450
         Left            =   2640
         TabIndex        =   22
         Top             =   2520
         Width           =   1875
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "전송내역 확인"
         Height          =   450
         Left            =   600
         TabIndex        =   21
         Top             =   2520
         Width           =   1875
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
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3840
         Width           =   12810
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   2160
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
      Begin VB.Frame Frame12 
         Caption         =   "접수번호 관련 기능 (요청번호 미할당)"
         Height          =   1815
         Left            =   240
         TabIndex        =   46
         Top             =   1920
         Width           =   4335
         Begin VB.Label Label4 
            Caption         =   "접수번호 :"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   280
            Width           =   975
         End
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
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   15375
      Begin VB.Frame Frame11 
         Caption         =   "파트너과금 포인트"
         Height          =   2415
         Left            =   12960
         TabIndex        =   41
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "연동과금 포인트"
         Height          =   2415
         Left            =   10680
         TabIndex        =   40
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " 회사정보 관련 "
         Height          =   2415
         Left            =   8640
         TabIndex        =   29
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 담당자 관련 "
         Height          =   2415
         Left            =   6600
         TabIndex        =   25
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보 "
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   24
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
         Height          =   2415
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2505
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   33
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
         Height          =   2415
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   11520
      TabIndex        =   59
      Top             =   360
      Width           =   525
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
' - 업데이트 일자 : 2022-01-17
' - 연동 기술지원 연락처 : 1600-9854
' - 연동 기술지원 이메일 : code@linkhubcorp.com
' - VB6 SDK 적용방법 안내 : https://docs.popbill.com/fax/tutorial/vb
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
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
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/fax/vb/api#CheckIsMember
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
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#CheckID
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
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/fax/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.CEOName = "대표자성명"
    
    '상호명, 최대 200자
    joinData.CorpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.Addr = "주소"
    
    '업태, 최대 100자
    joinData.BizType = "업태"
    
    '종목, 최대 100자
    joinData.BizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = FaxService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팩스 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetUnitCost
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
' 팝빌 팩스 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetChargeInfo
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
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = FaxService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/fax/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
        
    Set Response = FaxService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = ""
    
    Set info = FaxService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/fax/vb/api#ListContact
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
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/fax/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = FaxService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetCorpInfo
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
' 연동회원의 회사정보를 수정합니다
' - https://docs.popbill.com/fax/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.CEOName = "대표자"
    
    '상호, 최대 200자
    CorpInfo.CorpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.Addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.BizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.BizClass = "종목"
    
    Set Response = FaxService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/fax/vb/api#GetBalance
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
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = FaxService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = FaxService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = FaxService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/fax/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 팩스 1건을 전송합니다. (최대 전송파일 개수: 20개)
' - 팩스전송 문서 파일포맷 안내 : https://docs.popbill.com/fax/format?lang=vb
' - https://docs.popbill.com/fax/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Click()
    Dim FilePaths As New Collection
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim adsYN As Boolean
    Dim title As String
    Dim requestNum As String
    
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 동일한 팩스파일을 다수의 수신자에게 전송하기 위해 팝빌에 접수합니다. (최대 1,000건)
' - 팩스전송 문서 파일포맷 안내 : https://docs.popbill.com/fax/format?lang=vb
' - https://docs.popbill.com/fax/vb/api#SendFAX
'=========================================================================
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
    Dim requestNum As String
    
    
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
                
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 팩스 1건을 전송합니다.(다중파일 전송) (최대 전송파일 개수: 20개)
' - 팩스전송 문서 파일포맷 안내 : https://docs.popbill.com/fax/format?lang=vb
' - https://docs.popbill.com/fax/vb/api#SendFAX
'=========================================================================
Private Sub btnSendFAX_Multi_Click()
    Dim FilePaths As New Collection
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim senderNum As String
    Dim senderName As String
    Dim receiptNum As String
    Dim title As String
    Dim adsYN As Boolean
    Dim requestNum As String
    
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 동일한 팩스파일을 다수의 수신자에게 전송하기 위해 팝빌에 접수합니다.(다중파일 동보전송) (최대 전송파일 개수 : 20개) (최대 1,000건)
' - 팩스전송 문서 파일포맷 안내 : https://docs.popbill.com/fax/format?lang=vb
' - https://docs.popbill.com/fax/vb/api#SendFAX
'=========================================================================
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
    Dim requestNum As String
    
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 팝빌에서 반환 받은 접수번호를 통해 팩스 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetMessages
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
    
    tmp = "state(전송상태 코드) | result(전송결과 코드) | title(팩스제목) | sendNum(발신번호) | senderName(발신자명) | receiveNum(수신번호) | receiveNumType(수신번호 유형)  | receiveName(수신자명) |"
    tmp = tmp + "sendPageCnt(전체 페이지수) | successPageCnt(성공 페이지수) | failPageCnt(실패 페이지수) | refundPageCnt(환불 페이지수) | cancelPageCnt(취소 페이지수) |"
    tmp = tmp + "receiptDT(접수일시) | reserveDT(예약일시) | sendDT(전송일시) | resultDT(전송결과 수신일시) | receiptNum(접수번호) | "
    tmp = tmp + "requestNum(요청번호) | chargePageCnt(과금 페이지수) | tiffFileSize(변환파일용량(단위 : byte)) | fileNames(전송 파일명)" + vbCrLf
    
    For Each sentFax In sentFaxList
            
        '전송상태 코드
        tmp = tmp + CStr(sentFax.state) + " | "
        
        '전송결과 코드
        tmp = tmp + CStr(sentFax.result) + " | "
        
        '팩스제목
        tmp = tmp + sentFax.title + " | "
        
        '발신번호
        tmp = tmp + sentFax.sendNum + " | "
        
        '발신자명
        tmp = tmp + sentFax.senderName + " | "
        
        '수신번호
        tmp = tmp + sentFax.receiveNum + " | "
        
        '수신번호 유형
        tmp = tmp + sentFax.receiveNumType + " | "
        
        '수신자명
        tmp = tmp + sentFax.receiveName + " | "
        
        '전체 페이지수
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "
        
        '성공 페이지수
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "
        
        '실패 페이지수
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "
        
        '환불 페이지수
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "
        
        '취소 페이지수
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "
        
        '접수일시
        tmp = tmp + sentFax.receiptDT + " | "
        
        '예약일시
        tmp = tmp + sentFax.reserveDT + " | "
        
        '전송일시
        tmp = tmp + sentFax.sendDT + " | "
        
        '전송결과 수신일시
        tmp = tmp + sentFax.resultDT + " | "
                
        '접수번호
        tmp = tmp + sentFax.receiptNum + " | "
        
        '요청번호
        tmp = tmp + sentFax.requestNum + " | "
        
        '과금 페이지수
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "
        
        '변환파일용량  (단위 : byte)
        tmp = tmp + sentFax.tiffFileSize + "byte | "
        
        i = 0
        
        '전송 파일명
        For Each fileName In sentFax.fileNames
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
' 팝빌에서 반환받은 접수번호를 통해 예약접수된 팩스 전송을 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/fax/vb/api#CancelReserve
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
' 팝빌에서 반환받은 접수번호를 통해 팩스 1건을 재전송합니다.
' - 발신/수신 정보 미입력시 기존과 동일한 정보로 팩스가 전송되고, 접수일 기준 최대 60일이 경과되지 않는 건만 재전송이 가능합니다.
' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/fax/vb/api#ResendFAX
'=========================================================================
Private Sub btnResendFAX_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim title As String
    Dim requestNum As String
    
    ' 발신번호, 공백처리시 기존발신번호로 재전송
    senderNum = ""
    
    ' 발신자명, 공백처리시 기존발신자명으로 재선송
    senderName = ""
    
    ' 팩스제목
    title = "팩스 재전송 제목"
    
    ' 기존수신정보 변경없이 재전송하는 경우, receivers(수신정보) Collection을 Nothing 으로 선언
    Set receivers = Nothing
    
    ' 새로운 수신정보로 재전송하는 경우, 수신번호/수신자명을 기재하여 receivers Collection에 추가
    ' 수신번호
    'receiver.receiverNum = "0700000214"
    
    ' 수신자명
    'receiver.receiverName = "수신자_수정"
    
    ' 수신정보 Collection 추가
    'receivers.Add receiver
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 팝빌에서 반환받은 접수번호를 통해 다수건의 팩스를 재전송합니다. (최대 전송파일 개수: 20개) (최대 1,000건)
' - 발신/수신 정보 미입력시 기존과 동일한 정보로 팩스가 전송되고, 접수일 기준 최대 60일이 경과되지 않는 건만 재전송이 가능합니다.
' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/fax/vb/api#ResendFAX
'=========================================================================
Private Sub btnResendFAX_Same_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim i As Integer
    Dim title As String
    Dim requestNum As String
    
    ' 발신번호, 공백처리시 기존발신번호로 재전송
    senderNum = ""
    
    ' 발신자명, 공백처리시 기존발신자명으로 재전송
    senderName = ""
    
    ' 팩스제목
    title = "팩스 동보 재전송 제목"
    
    ' 기존수신정보 변경없이 재전송하는 경우, receivers(수신정보) Collection을 Nothing 으로 선언
    'Set receivers = Nothing
    
    ' 새로운 수신정보로 재전송하는 경우, 수신번호/수신자명을 기재하여 receivers Collection에 추가
    ' 수신정보, 최대 1000건
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "수신자 명칭"
        receivers.Add receiver
    Next
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 파트너가 할당한 전송요청 번호를 통해 팩스 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetMessagesRN
'=========================================================================
Private Sub btnGetFaxDetailRN_Click()
Dim sentFaxList As Collection
    Dim i As Integer
    Dim fileName As Variant
    Dim sentFax As PBFaxInfo
    Dim tmp As String
    
    Set sentFaxList = FaxService.GetMessagesRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If sentFaxList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state(전송상태 코드) | result(전송결과 코드) | title(팩스제목) | sendNum(발신번호) | senderName(발신자명) | receiveNum(수신번호) | receiveNumType(수신번호 유형) | receiveName(수신자명) |"
    tmp = tmp + "sendPageCnt(전체 페이지수) | successPageCnt(성공 페이지수) | failPageCnt(실패 페이지수) | refundPageCnt(환불 페이지수) | cancelPageCnt(취소 페이지수) |"
    tmp = tmp + "receiptDT(접수일시) | reserveDT(예약일시) | sendDT(전송일시) | resultDT(전송결과 수신일시) | receiptNum(접수번호) | "
    tmp = tmp + "requestNum(요청번호) | chargePageCnt(과금 페이지수) | tiffFileSize(변환파일용량(단위 : byte)) | fileNames(전송 파일명)" + vbCrLf
    
    For Each sentFax In sentFaxList
        tmp = tmp + CStr(sentFax.state) + " | "             '전송상태 코드
        tmp = tmp + CStr(sentFax.result) + " | "            '전송결과 코드
        tmp = tmp + sentFax.title + " | "                   '팩스제목
        tmp = tmp + sentFax.sendNum + " | "                 '발신번호
        tmp = tmp + sentFax.senderName + " | "              '발신자명
        tmp = tmp + sentFax.receiveNum + " | "              '수신번호
        tmp = tmp + sentFax.receiveNumType + " | "          '수신번호 유형
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
        tmp = tmp + sentFax.receiptNum + " | "              '접수번호
        tmp = tmp + sentFax.requestNum + " | "              '요청번호
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "     '과금 페이지수
        tmp = tmp + sentFax.tiffFileSize + "byte | "        '변환파일용량 (단위 : byte)
        
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
' 파트너가 할당한 전송요청 번호를 통해 예약접수된 팩스 전송을 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/fax/vb/api#CancelReserveRN
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 파트너가 할당한 전송요청 번호를 통해 팩스 1건을 재전송합니다.
' - 발신/수신 정보 미입력시 기존과 동일한 정보로 팩스가 전송되고, 접수일 기준 최대 60일이 경과되지 않는 건만 재전송이 가능합니다.
' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/fax/vb/api#ResendFAXRN
'=========================================================================
Private Sub btnResendFAXRN_Click()
    Dim OrgRequestNum As String
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim requestNum As String
    Dim title As String
    
    '원본 팩스 전송시 할당한 전송요청번호(requestNum)
    OrgRequestNum = ""
    
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
    
    '팩스제목
    title = ""
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.ResendFAXRN(txtCorpNum.Text, OrgRequestNum, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 파트너가 할당한 전송요청 번호를 통해 다수건의 팩스를 재전송합니다. (최대 전송파일 개수: 20개) (최대 1,000건)
' - 발신/수신 정보 미입력시 기존과 동일한 정보로 팩스가 전송되고, 접수일 기준 최대 60일이 경과되지 않는 건만 재전송이 가능합니다.
' - 팩스 재전송 요청시 포인트가 차감됩니다. (전송실패시 환불처리)
' - https://docs.popbill.com/fax/vb/api#ResendFAXRN
'=========================================================================
Private Sub btnResendFAXRN_Same_Click()
    Dim OrgRequestNum As String
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim i As Integer
    Dim requestNum As String
    Dim title As String

    '원본 팩스 전송시 할당한 전송요청번호(requestNum)
    OrgRequestNum = ""

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
    
    '팩스제목
    title = ""
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    receiptNum = FaxService.ResendFAXRN(txtCorpNum.Text, OrgRequestNum, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 팝빌에 등록한 연동회원의 팩스 발신번호 목록을 확인합니다.
' - https://docs.popbill.com/fax/vb/api#GetSenderNumberList
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
' 발신번호를 등록하고 내역을 확인하는 팩스 발신번호 관리 페이지 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetSenderNumberMgtURL
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()
    Dim url As String
    
    url = FaxService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 검색조건에 해당하는 팩스 전송내역 목록을 조회합니다. (조회기간 단위 : 최대 2개월)
' - 팩스 접수일시로부터 2개월 이내 접수건만 조회할 수 있습니다.
' - https://docs.popbill.com/fax/vb/api#Search
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
    Dim QString As String
    
    '[필수] 시작일자, 형식(yyyyMMdd)
    SDate = "20220101"
    
    '[필수] 종료일자, 형식(yyyyMMdd)
    EDate = "20220130"
    
    '전송상태 배열, 1(대기), 2(성공), 3(실패), 4(취소)
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '예약전송 검색여부, True-예약전송건 조회, False-즉시전송건 조회
    ReserveYN = False
    
    '개인조회 여부, True-개인조회, False-회사조회
    SenderOnly = False
    
    '페이지 번호, 기본값 1
    Page = 1
    
    '페이지당 검색개수, 기본값 500, 최대값 1000
    PerPage = 30
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '조회 검색어, 발신자명 또는 수신자명 기재
    QString = ""
    
    Set faxSearchList = FaxService.Search(txtCorpNum.Text, SDate, EDate, state, ReserveYN, SenderOnly, Page, PerPage, Order, QString)
     
    If faxSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (응답코드) : " + CStr(faxSearchList.code) + vbCrLf
    tmp = tmp + "total (총 검색건수) : " + CStr(faxSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(faxSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지번호) : " + CStr(faxSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지개수) : " + CStr(faxSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + faxSearchList.message + vbCrLf + vbCrLf
    
    MsgBox tmp

    tmp = "state(전송상태 코드) | result(전송결과 코드) | title(팩스제목) | sendnum(발신번호) | senderName(발신자명) | receiveNum(수신번호) | receiveNumType(수신번호 유형)  | receiveName(수신자명) |"
    tmp = tmp + "sendPageCnt(전체 페이지수) | successPageCnt(성공 페이지수) | failPageCnt(실패 페이지수) | refundPageCnt(환불 페이지수) | cancelPageCnt(취소 페이지수) |"
    tmp = tmp + "receiptDT(접수일시) | reserveDT(예약일시) | sendDT(전송일시) | resultDT(전송결과 수신일시) | receiptNum(접수번호) | "
    tmp = tmp + "requestNum(요청번호) | chargePageCnt(과금 페이지수) | tiffFileSize(변환파일용량(단위 : byte)) | fileNames(전송 파일명)" + vbCrLf
    
    Dim sentFax As PBFaxInfo
    
    For Each sentFax In faxSearchList.list
    
        '전송상태 코드
        tmp = tmp + CStr(sentFax.state) + " | "
        
        '전송결과 코드
        tmp = tmp + CStr(sentFax.result) + " | "
        
        '팩스제목
        tmp = tmp + sentFax.title + " | "
        
        '발신번호
        tmp = tmp + sentFax.sendNum + " | "
        
        '발신자명
        tmp = tmp + sentFax.senderName + " | "
        
        '수신번호
        tmp = tmp + sentFax.receiveNum + " | "
        
        '수신번호 유형
        tmp = tmp + sentFax.receiveNumType + " | "
        
        '수신자명
        tmp = tmp + sentFax.receiveName + " | "
        
        '전체 페이지수
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "
        
        '성공 페이지수
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "
        
        '실패 페이지수
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "
        
        '환불 페이지수
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "
        
        '취소 페이지수
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "
        
        '접수일시
        tmp = tmp + sentFax.receiptDT + " | "
        
        '예약일시
        tmp = tmp + sentFax.reserveDT + " | "
        
        '전송일시
        tmp = tmp + sentFax.sendDT + " | "
        
        '전송결과 수신일시
        tmp = tmp + sentFax.resultDT + " | "
        
        '접수번호
        tmp = tmp + sentFax.receiptNum + " | "
        
        '요청번호
        tmp = tmp + sentFax.requestNum + " | "
        
        '과금 페이지수
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "
        
        '변환파일용량 (단위 : byte)
        tmp = tmp + sentFax.tiffFileSize + "byte | "
                
        
        i = 0
        
        For Each fileName In sentFax.fileNames              '전송 파일이름
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
' 팝빌 사이트와 동일한 팩스 전송내역 확인 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetSentListURL
'=========================================================================
Private Sub btnGetSentListURL_Click()
    Dim url As String
    
    url = FaxService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
'팩스 미리보기 팝업 URL을 반환하며, 팩스전송을 위한 TIF 포맷 변환 완료 후 호출 할 수 있습니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/fax/vb/api#GetPreviewURL
'=========================================================================
Private Sub btnGetPreviewURL_Click()
    Dim url As String
    
    url = FaxService.GetPreviewURL(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(FaxService.LastErrCode) + vbCrLf + "응답메시지 : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

Private Sub Form_Load()
    
    '팩스서비스 모듈 초기화
    FaxService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    FaxService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
    FaxService.IPRestrictOnOff = True
    
    '로컬시스템 시간 사용여부 True-사용, Fasle-미사용, 기본값(False)
    FaxService.UseLocalTimeYN = False
        
End Sub

