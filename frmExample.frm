VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� �ѽ� SDK ����"
   ClientHeight    =   13470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   13470
   ScaleWidth      =   15795
   StartUpPosition =   2  'ȭ�� ���
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
      Caption         =   " �ѽ� ���� ���� "
      Height          =   8895
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   13455
      Begin VB.Frame Frame13 
         Caption         =   "��û��ȣ �Ҵ� ���۰� ó��"
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
            Caption         =   "���� ������"
            Height          =   450
            Left            =   2280
            TabIndex        =   53
            Top             =   1200
            Width           =   1875
         End
         Begin VB.CommandButton btnResendFAXRN 
            Caption         =   "������"
            Height          =   450
            Left            =   240
            TabIndex        =   52
            Top             =   1200
            Width           =   1875
         End
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "�������� ���"
            Height          =   450
            Left            =   2280
            TabIndex        =   51
            Top             =   600
            Width           =   1875
         End
         Begin VB.CommandButton btnGetFaxDetailRN 
            Caption         =   "���۳��� Ȯ��"
            Height          =   450
            Left            =   245
            TabIndex        =   50
            Top             =   600
            Width           =   1875
         End
         Begin VB.Label Label5 
            Caption         =   "��û��ȣ :"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   295
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "�߽Ź�ȣ ����"
         Height          =   1575
         Left            =   9120
         TabIndex        =   37
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "�߽Ź�ȣ ���� �˾�"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "�߽Ź�ȣ ��� ��ȸ"
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "���۳��� �˻���ȸ"
         Height          =   495
         Left            =   11400
         TabIndex        =   32
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton btnGetSentListURL 
         Caption         =   "���۳�����ȸ �˾�"
         Height          =   495
         Left            =   11400
         TabIndex        =   23
         Top             =   720
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ΰ����"
         Height          =   2175
         Left            =   11280
         TabIndex        =   36
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetPreviewURL 
            Caption         =   "�ѽ� �̸����� �˾�"
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnResendFaxSame 
         Caption         =   "���� ������"
         Height          =   450
         Left            =   2640
         TabIndex        =   35
         Top             =   3120
         Width           =   1875
      End
      Begin VB.CommandButton btnResendFAX 
         Caption         =   "������"
         Height          =   450
         Left            =   600
         TabIndex        =   34
         Top             =   3120
         Width           =   1875
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "�������� ���"
         Height          =   450
         Left            =   2640
         TabIndex        =   22
         Top             =   2520
         Width           =   1875
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "���۳��� Ȯ��"
         Height          =   450
         Left            =   600
         TabIndex        =   21
         Top             =   2520
         Width           =   1875
      End
      Begin VB.TextBox txtResult 
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ټ����� ��������"
         Height          =   450
         Left            =   5280
         TabIndex        =   18
         Top             =   840
         Width           =   1875
      End
      Begin VB.CommandButton btnSendFAX_Multi 
         Caption         =   "�ټ� ���� ����"
         Height          =   450
         Left            =   3600
         TabIndex        =   17
         Top             =   840
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFax_Same 
         Caption         =   "���� ����"
         Height          =   450
         Left            =   1920
         TabIndex        =   16
         Top             =   840
         Width           =   1590
      End
      Begin VB.CommandButton btnSendFAX 
         Caption         =   "����"
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
         Caption         =   "������ȣ ���� ��� (��û��ȣ ���Ҵ�)"
         Height          =   1815
         Left            =   240
         TabIndex        =   46
         Top             =   1920
         Width           =   4335
         Begin VB.Label Label4 
            Caption         =   "������ȣ :"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   280
            Width           =   975
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�������� �ð�(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   450
         Width           =   3210
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   15375
      Begin VB.Frame Frame11 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2415
         Left            =   12960
         TabIndex        =   41
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������� ����Ʈ"
         Height          =   2415
         Left            =   10680
         TabIndex        =   40
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " ȸ������ ���� "
         Height          =   2415
         Left            =   8640
         TabIndex        =   29
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ����� ���� "
         Height          =   2415
         Left            =   6600
         TabIndex        =   25
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "����� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������ "
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ���� "
         Height          =   2415
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2505
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   2175
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "���� �ܰ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL "
         Height          =   2415
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
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
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
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
' �˺� �ѽ� API VB 6.0 SDK Example
'
' - ������Ʈ ���� : 2022-01-17
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB6 SDK ������ �ȳ� : https://docs.popbill.com/fax/tutorial/vb
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'�ѽ� ���� ��ü ����
Private FaxService As New PBFAXService

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
' - https://docs.popbill.com/fax/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/fax/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.CEOName = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.CorpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.Addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.BizType = "����"
    
    '����, �ִ� 100��
    joinData.BizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = FaxService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ѽ� ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = FaxService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� �ѽ� API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = FaxService.GetChargeInfo(txtCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (���۴ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = FaxService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/fax/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� �̸�
    joinData.id = "testkorea"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
    
    '����� �ѽ���,�ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
        
    Set Response = FaxService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = ""
    
    Set info = FaxService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = FaxService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | hp(�޴�����ȣ) |  fax(�ѽ���ȣ) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/fax/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = FaxService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = FaxService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڼ���) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (��ȣ) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
' - https://docs.popbill.com/fax/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.CEOName = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.CorpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.Addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.BizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.BizClass = "����"
    
    Set Response = FaxService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)�� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/fax/vb/api#GetBalance
'=========================================================================
Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = FaxService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = FaxService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = FaxService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)�� �̿��Ͻñ� �ٶ��ϴ�.
' - https://docs.popbill.com/fax/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(FaxService.LastErrCode) + "] " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �ѽ� 1���� �����մϴ�. (�ִ� �������� ����: 20��)
' - �ѽ����� ���� �������� �ȳ� : https://docs.popbill.com/fax/format?lang=vb
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
    
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '�߽��ڸ�
    senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    receiver.receiverNum = "070111222"
    
    '�����ڸ�
    receiver.receiverName = "������ ��Ī"
    receivers.Add receiver
    
    '�����ѽ� ���ۿ���
    adsYN = False
    
    '�ѽ�����
    title = "�ѽ� �ܰ����� ����"
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' ������ �ѽ������� �ټ��� �����ڿ��� �����ϱ� ���� �˺��� �����մϴ�. (�ִ� 1,000��)
' - �ѽ����� ���� �������� �ȳ� : https://docs.popbill.com/fax/format?lang=vb
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
        
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '�߽��ڸ�
    senderName = "�߽��ڸ�"
    
    '��������, �ִ� 1000��
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "������ ��Ī"
        receivers.Add receiver
    Next
    
    '�����ѽ� ���ۿ���
    adsYN = True
    
    '�ѽ�����
    title = "�ѽ� �������� ����"
                
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �ѽ� 1���� �����մϴ�.(�������� ����) (�ִ� �������� ����: 20��)
' - �ѽ����� ���� �������� �ȳ� : https://docs.popbill.com/fax/format?lang=vb
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
    
    '���� ���� ���� �ִ� 20��
    Do
        CommonDialog1.fileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.fileName <> "" Then
            FilePaths.Add CommonDialog1.fileName
        End If
    
    Loop While (CommonDialog1.fileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '�߽��ڸ�
    senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    receiver.receiverNum = "070111222"
    
    '�����ڸ�
    receiver.receiverName = "������ ��Ī"
    
    receivers.Add receiver
    
    '�����ѽ� ���ۿ���
    adsYN = False
    
    '�ѽ�����
    title = "�ѽ� �ܰ� �ټ����� �ѽ�����"
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' ������ �ѽ������� �ټ��� �����ڿ��� �����ϱ� ���� �˺��� �����մϴ�.(�������� ��������) (�ִ� �������� ���� : 20��) (�ִ� 1,000��)
' - �ѽ����� ���� �������� �ȳ� : https://docs.popbill.com/fax/format?lang=vb
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
    
    '���� ���� ���� �ִ� 20��
    Do
        CommonDialog1.fileName = ""
        CommonDialog1.ShowOpen
        
        If CommonDialog1.fileName <> "" Then
            FilePaths.Add CommonDialog1.fileName
        End If
    
    Loop While (CommonDialog1.fileName <> "")
    
    If FilePaths.Count = 0 Then Exit Sub
    
    '�߽Ź�ȣ
    senderNum = "07043042991"
    
    '�߽��ڸ�
    senderName = "�߽��ڸ�"
    
    '�������� �ִ� 1000����� ����
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "������ ��Ī"
        receivers.Add receiver
    Next
    
    '�����ѽ� ���ۿ���
    adsYN = False
    
    '�ѽ�����
    title = "�ѽ� �ټ����� �������� ����"
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' �˺����� ��ȯ ���� ������ȣ�� ���� �ѽ� ���ۻ��� �� ����� Ȯ���մϴ�.
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
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state(���ۻ��� �ڵ�) | result(���۰�� �ڵ�) | title(�ѽ�����) | sendNum(�߽Ź�ȣ) | senderName(�߽��ڸ�) | receiveNum(���Ź�ȣ) | receiveNumType(���Ź�ȣ ����)  | receiveName(�����ڸ�) |"
    tmp = tmp + "sendPageCnt(��ü ��������) | successPageCnt(���� ��������) | failPageCnt(���� ��������) | refundPageCnt(ȯ�� ��������) | cancelPageCnt(��� ��������) |"
    tmp = tmp + "receiptDT(�����Ͻ�) | reserveDT(�����Ͻ�) | sendDT(�����Ͻ�) | resultDT(���۰�� �����Ͻ�) | receiptNum(������ȣ) | "
    tmp = tmp + "requestNum(��û��ȣ) | chargePageCnt(���� ��������) | tiffFileSize(��ȯ���Ͽ뷮(���� : byte)) | fileNames(���� ���ϸ�)" + vbCrLf
    
    For Each sentFax In sentFaxList
            
        '���ۻ��� �ڵ�
        tmp = tmp + CStr(sentFax.state) + " | "
        
        '���۰�� �ڵ�
        tmp = tmp + CStr(sentFax.result) + " | "
        
        '�ѽ�����
        tmp = tmp + sentFax.title + " | "
        
        '�߽Ź�ȣ
        tmp = tmp + sentFax.sendNum + " | "
        
        '�߽��ڸ�
        tmp = tmp + sentFax.senderName + " | "
        
        '���Ź�ȣ
        tmp = tmp + sentFax.receiveNum + " | "
        
        '���Ź�ȣ ����
        tmp = tmp + sentFax.receiveNumType + " | "
        
        '�����ڸ�
        tmp = tmp + sentFax.receiveName + " | "
        
        '��ü ��������
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "
        
        'ȯ�� ��������
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "
        
        '��� ��������
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.receiptDT + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.reserveDT + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.sendDT + " | "
        
        '���۰�� �����Ͻ�
        tmp = tmp + sentFax.resultDT + " | "
                
        '������ȣ
        tmp = tmp + sentFax.receiptNum + " | "
        
        '��û��ȣ
        tmp = tmp + sentFax.requestNum + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "
        
        '��ȯ���Ͽ뷮  (���� : byte)
        tmp = tmp + sentFax.tiffFileSize + "byte | "
        
        i = 0
        
        '���� ���ϸ�
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
' �˺����� ��ȯ���� ������ȣ�� ���� ���������� �ѽ� ������ ����մϴ�. (����ð� 10�� ������ ����)
' - https://docs.popbill.com/fax/vb/api#CancelReserve
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺����� ��ȯ���� ������ȣ�� ���� �ѽ� 1���� �������մϴ�.
' - �߽�/���� ���� ���Է½� ������ ������ ������ �ѽ��� ���۵ǰ�, ������ ���� �ִ� 60���� ������� �ʴ� �Ǹ� �������� �����մϴ�.
' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
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
    
    ' �߽Ź�ȣ, ����ó���� �����߽Ź�ȣ�� ������
    senderNum = ""
    
    ' �߽��ڸ�, ����ó���� �����߽��ڸ����� �缱��
    senderName = ""
    
    ' �ѽ�����
    title = "�ѽ� ������ ����"
    
    ' ������������ ������� �������ϴ� ���, receivers(��������) Collection�� Nothing ���� ����
    Set receivers = Nothing
    
    ' ���ο� ���������� �������ϴ� ���, ���Ź�ȣ/�����ڸ��� �����Ͽ� receivers Collection�� �߰�
    ' ���Ź�ȣ
    'receiver.receiverNum = "0700000214"
    
    ' �����ڸ�
    'receiver.receiverName = "������_����"
    
    ' �������� Collection �߰�
    'receivers.Add receiver
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �˺����� ��ȯ���� ������ȣ�� ���� �ټ����� �ѽ��� �������մϴ�. (�ִ� �������� ����: 20��) (�ִ� 1,000��)
' - �߽�/���� ���� ���Է½� ������ ������ ������ �ѽ��� ���۵ǰ�, ������ ���� �ִ� 60���� ������� �ʴ� �Ǹ� �������� �����մϴ�.
' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
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
    
    ' �߽Ź�ȣ, ����ó���� �����߽Ź�ȣ�� ������
    senderNum = ""
    
    ' �߽��ڸ�, ����ó���� �����߽��ڸ����� ������
    senderName = ""
    
    ' �ѽ�����
    title = "�ѽ� ���� ������ ����"
    
    ' ������������ ������� �������ϴ� ���, receivers(��������) Collection�� Nothing ���� ����
    'Set receivers = Nothing
    
    ' ���ο� ���������� �������ϴ� ���, ���Ź�ȣ/�����ڸ��� �����Ͽ� receivers Collection�� �߰�
    ' ��������, �ִ� 1000��
    For i = 1 To 5
        Set receiver = New PBReceiver
        receiver.receiverNum = "070111222"
        receiver.receiverName = "������ ��Ī"
        receivers.Add receiver
    Next
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �ѽ� ���ۻ��� �� ����� Ȯ���մϴ�.
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
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state(���ۻ��� �ڵ�) | result(���۰�� �ڵ�) | title(�ѽ�����) | sendNum(�߽Ź�ȣ) | senderName(�߽��ڸ�) | receiveNum(���Ź�ȣ) | receiveNumType(���Ź�ȣ ����) | receiveName(�����ڸ�) |"
    tmp = tmp + "sendPageCnt(��ü ��������) | successPageCnt(���� ��������) | failPageCnt(���� ��������) | refundPageCnt(ȯ�� ��������) | cancelPageCnt(��� ��������) |"
    tmp = tmp + "receiptDT(�����Ͻ�) | reserveDT(�����Ͻ�) | sendDT(�����Ͻ�) | resultDT(���۰�� �����Ͻ�) | receiptNum(������ȣ) | "
    tmp = tmp + "requestNum(��û��ȣ) | chargePageCnt(���� ��������) | tiffFileSize(��ȯ���Ͽ뷮(���� : byte)) | fileNames(���� ���ϸ�)" + vbCrLf
    
    For Each sentFax In sentFaxList
        tmp = tmp + CStr(sentFax.state) + " | "             '���ۻ��� �ڵ�
        tmp = tmp + CStr(sentFax.result) + " | "            '���۰�� �ڵ�
        tmp = tmp + sentFax.title + " | "                   '�ѽ�����
        tmp = tmp + sentFax.sendNum + " | "                 '�߽Ź�ȣ
        tmp = tmp + sentFax.senderName + " | "              '�߽��ڸ�
        tmp = tmp + sentFax.receiveNum + " | "              '���Ź�ȣ
        tmp = tmp + sentFax.receiveNumType + " | "          '���Ź�ȣ ����
        tmp = tmp + sentFax.receiveName + " | "             '�����ڸ�
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "       '��ü ��������
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "    '���� ��������
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "       '���� ��������
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "     'ȯ�� ��������
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "     '��� ��������
        tmp = tmp + CStr(sentFax.receiptDT) + " | "         '�����Ͻ�
        tmp = tmp + sentFax.reserveDT + " | "               '�����Ͻ�
        tmp = tmp + sentFax.sendDT + " | "                  '�����Ͻ�
        tmp = tmp + sentFax.resultDT + " | "                '���۰�� �����Ͻ�
        tmp = tmp + sentFax.receiptNum + " | "              '������ȣ
        tmp = tmp + sentFax.requestNum + " | "              '��û��ȣ
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "     '���� ��������
        tmp = tmp + sentFax.tiffFileSize + "byte | "        '��ȯ���Ͽ뷮 (���� : byte)
        
        i = 0
        
        For Each fileName In sentFax.fileNames              '�ѽ����� ���ϸ�
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
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� ���������� �ѽ� ������ ����մϴ�. (����ð� 10�� ������ ����)
' - https://docs.popbill.com/fax/vb/api#CancelReserveRN
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = FaxService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �ѽ� 1���� �������մϴ�.
' - �߽�/���� ���� ���Է½� ������ ������ ������ �ѽ��� ���۵ǰ�, ������ ���� �ִ� 60���� ������� �ʴ� �Ǹ� �������� �����մϴ�.
' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
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
    
    '���� �ѽ� ���۽� �Ҵ��� ���ۿ�û��ȣ(requestNum)
    OrgRequestNum = ""
    
    ' �߽Ź�ȣ, ����ó���� �����߽Ź�ȣ�� ������
    senderNum = ""
    
    ' �߽��ڸ�, ����ó���� �����߽��ڸ����� ������
    senderName = ""
    
    ' ������������ ������� �������ϴ� ���, receivers(��������) Collection �� Nothing ���� ����
    Set receivers = Nothing
    
    ' ���ο� ���������� �������ϴ� ���, ���Ź�ȣ/�����ڸ��� �����Ͽ� receivers Collection�� �߰�
    ' ���Ź�ȣ
    'receiver.receiverNum = "0700000214"
    
    ' �����ڸ�
    'receiver.receiverName = "������_����"
    
    ' �������� Collection �߰�
    'receivers.Add receiver
    
    '�ѽ�����
    title = ""
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.ResendFAXRN(txtCorpNum.Text, OrgRequestNum, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' ��Ʈ�ʰ� �Ҵ��� ���ۿ�û ��ȣ�� ���� �ټ����� �ѽ��� �������մϴ�. (�ִ� �������� ����: 20��) (�ִ� 1,000��)
' - �߽�/���� ���� ���Է½� ������ ������ ������ �ѽ��� ���۵ǰ�, ������ ���� �ִ� 60���� ������� �ʴ� �Ǹ� �������� �����մϴ�.
' - �ѽ� ������ ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
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

    '���� �ѽ� ���۽� �Ҵ��� ���ۿ�û��ȣ(requestNum)
    OrgRequestNum = ""

    ' �߽Ź�ȣ, ����ó���� �����߽Ź�ȣ�� ������
    senderNum = ""
    
    ' �߽��ڸ�, ����ó���� �����߽��ڸ����� ������
    senderName = ""
    
    ' ������������ ������� �������ϴ� ���, receivers(��������) Collection �� Nothing ���� ����
    'Set receivers = Nothing
    
    ' ���ο� ���������� �������ϴ� ���, ���Ź�ȣ/�����ڸ��� �����Ͽ� receivers Collection�� �߰�
    ' ��������, �ִ� 1000��
    For i = 1 To 10
        Set receiver = New PBReceiver
        receiver.receiverNum = "010111222"
        receiver.receiverName = "������ ��Ī"
        receivers.Add receiver
    Next
    
    '�ѽ�����
    title = ""
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    receiptNum = FaxService.ResendFAXRN(txtCorpNum.Text, OrgRequestNum, senderNum, senderName, receivers, txtReserveDT.Text, txtUserID.Text, title, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �˺��� ����� ����ȸ���� �ѽ� �߽Ź�ȣ ����� Ȯ���մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetSenderNumberList
'=========================================================================
Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim SenderNumber As PBFaxSenderNumber
    
    Set SenderNumberList = FaxService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumber In SenderNumberList
        tmp = tmp + "�߽Ź�ȣ(number) : " + SenderNumber.number + vbCrLf
        tmp = tmp + "��ǥ��ȣ ��������(representYN) : " + CStr(SenderNumber.representYN) + vbCrLf
        tmp = tmp + "��ϻ���(state) : " + CStr(SenderNumber.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' �߽Ź�ȣ�� ����ϰ� ������ Ȯ���ϴ� �ѽ� �߽Ź�ȣ ���� ������ �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetSenderNumberMgtURL
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()
    Dim url As String
    
    url = FaxService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' �˻����ǿ� �ش��ϴ� �ѽ� ���۳��� ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 2����)
' - �ѽ� �����Ͻ÷κ��� 2���� �̳� �����Ǹ� ��ȸ�� �� �ֽ��ϴ�.
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
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    SDate = "20220101"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    EDate = "20220130"
    
    '���ۻ��� �迭, 1(���), 2(����), 3(����), 4(���)
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '�������� �˻�����, True-�������۰� ��ȸ, False-������۰� ��ȸ
    ReserveYN = False
    
    '������ȸ ����, True-������ȸ, False-ȸ����ȸ
    SenderOnly = False
    
    '������ ��ȣ, �⺻�� 1
    Page = 1
    
    '�������� �˻�����, �⺻�� 500, �ִ밪 1000
    PerPage = 30
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '��ȸ �˻���, �߽��ڸ� �Ǵ� �����ڸ� ����
    QString = ""
    
    Set faxSearchList = FaxService.Search(txtCorpNum.Text, SDate, EDate, state, ReserveYN, SenderOnly, Page, PerPage, Order, QString)
     
    If faxSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (�����ڵ�) : " + CStr(faxSearchList.code) + vbCrLf
    tmp = tmp + "total (�� �˻��Ǽ�) : " + CStr(faxSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� �˻�����) : " + CStr(faxSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (��������ȣ) : " + CStr(faxSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (����������) : " + CStr(faxSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + faxSearchList.message + vbCrLf + vbCrLf
    
    MsgBox tmp

    tmp = "state(���ۻ��� �ڵ�) | result(���۰�� �ڵ�) | title(�ѽ�����) | sendnum(�߽Ź�ȣ) | senderName(�߽��ڸ�) | receiveNum(���Ź�ȣ) | receiveNumType(���Ź�ȣ ����)  | receiveName(�����ڸ�) |"
    tmp = tmp + "sendPageCnt(��ü ��������) | successPageCnt(���� ��������) | failPageCnt(���� ��������) | refundPageCnt(ȯ�� ��������) | cancelPageCnt(��� ��������) |"
    tmp = tmp + "receiptDT(�����Ͻ�) | reserveDT(�����Ͻ�) | sendDT(�����Ͻ�) | resultDT(���۰�� �����Ͻ�) | receiptNum(������ȣ) | "
    tmp = tmp + "requestNum(��û��ȣ) | chargePageCnt(���� ��������) | tiffFileSize(��ȯ���Ͽ뷮(���� : byte)) | fileNames(���� ���ϸ�)" + vbCrLf
    
    Dim sentFax As PBFaxInfo
    
    For Each sentFax In faxSearchList.list
    
        '���ۻ��� �ڵ�
        tmp = tmp + CStr(sentFax.state) + " | "
        
        '���۰�� �ڵ�
        tmp = tmp + CStr(sentFax.result) + " | "
        
        '�ѽ�����
        tmp = tmp + sentFax.title + " | "
        
        '�߽Ź�ȣ
        tmp = tmp + sentFax.sendNum + " | "
        
        '�߽��ڸ�
        tmp = tmp + sentFax.senderName + " | "
        
        '���Ź�ȣ
        tmp = tmp + sentFax.receiveNum + " | "
        
        '���Ź�ȣ ����
        tmp = tmp + sentFax.receiveNumType + " | "
        
        '�����ڸ�
        tmp = tmp + sentFax.receiveName + " | "
        
        '��ü ��������
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "
        
        'ȯ�� ��������
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "
        
        '��� ��������
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.receiptDT + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.reserveDT + " | "
        
        '�����Ͻ�
        tmp = tmp + sentFax.sendDT + " | "
        
        '���۰�� �����Ͻ�
        tmp = tmp + sentFax.resultDT + " | "
        
        '������ȣ
        tmp = tmp + sentFax.receiptNum + " | "
        
        '��û��ȣ
        tmp = tmp + sentFax.requestNum + " | "
        
        '���� ��������
        tmp = tmp + CStr(sentFax.chargePageCnt) + " | "
        
        '��ȯ���Ͽ뷮 (���� : byte)
        tmp = tmp + sentFax.tiffFileSize + "byte | "
                
        
        i = 0
        
        For Each fileName In sentFax.fileNames              '���� �����̸�
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
' �˺� ����Ʈ�� ������ �ѽ� ���۳��� Ȯ�� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetSentListURL
'=========================================================================
Private Sub btnGetSentListURL_Click()
    Dim url As String
    
    url = FaxService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
'�ѽ� �̸����� �˾� URL�� ��ȯ�ϸ�, �ѽ������� ���� TIF ���� ��ȯ �Ϸ� �� ȣ�� �� �� �ֽ��ϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/fax/vb/api#GetPreviewURL
'=========================================================================
Private Sub btnGetPreviewURL_Click()
    Dim url As String
    
    url = FaxService.GetPreviewURL(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

Private Sub Form_Load()
    
    '�ѽ����� ��� �ʱ�ȭ
    FaxService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    FaxService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    FaxService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, Fasle-�̻��, �⺻��(False)
    FaxService.UseLocalTimeYN = False
        
End Sub

