VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� �ѽ� SDK ����"
   ClientHeight    =   11910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   11910
   ScaleWidth      =   15795
   StartUpPosition =   2  'ȭ�� ���
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   " �ѽ� ���� ���� "
      Height          =   8175
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   13455
      Begin VB.Frame Frame9 
         Caption         =   "�߽Ź�ȣ ����"
         Height          =   1575
         Left            =   10320
         TabIndex        =   38
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetURL_SENDER 
            Caption         =   "�߽Ź�ȣ ���� �˾�"
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "�߽Ź�ȣ ��� ��ȸ"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "���۳��� �˻���ȸ"
         Height          =   465
         Left            =   8025
         TabIndex        =   33
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton btnSearchPopUp 
         Caption         =   "���۳�����ȸ �˾�"
         Height          =   465
         Left            =   8025
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "�ΰ����"
         Height          =   1575
         Left            =   7800
         TabIndex        =   37
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton btnResendFaxSame 
         Caption         =   "���� ������"
         Height          =   450
         Left            =   1920
         TabIndex        =   36
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton btnResendFAX 
         Caption         =   "������"
         Height          =   450
         Left            =   360
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "�������� ���"
         Height          =   450
         Left            =   6120
         TabIndex        =   23
         Top             =   2115
         Width           =   1515
      End
      Begin VB.CommandButton btnGetFaxDetail 
         Caption         =   "���۳��� Ȯ��"
         Height          =   450
         Left            =   4440
         TabIndex        =   22
         Top             =   2115
         Width           =   1515
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "������ȣ : "
         Height          =   180
         Left            =   420
         TabIndex        =   19
         Top             =   2250
         Width           =   900
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
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   15375
      Begin VB.Frame Frame11 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   12840
         TabIndex        =   42
         Top             =   360
         Width           =   2295
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Left            =   10680
         TabIndex        =   41
         Top             =   360
         Width           =   2055
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " ȸ������ ���� "
         Height          =   1935
         Left            =   8640
         TabIndex        =   30
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   32
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ����� ���� "
         Height          =   1935
         Left            =   6600
         TabIndex        =   26
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������ "
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   25
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
         Height          =   1935
         Left            =   2040
         TabIndex        =   7
         Top             =   360
         Width           =   2505
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   34
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
         Height          =   1935
         Left            =   4680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL 
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
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569
' - ������Ʈ ���� : 2017-08-30
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
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
' �������� �ѽ���û���� ����մϴ�.
' - �������� ��Ҵ� �������۽ð� 10�������� �����մϴ�.
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
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
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
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
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
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
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
' ����ȸ���� �ѽ� API ���� ���������� Ȯ���մϴ�.
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
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
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
' �ѽ� ���ۿ�û�� ��ȯ���� ������ȣ(receiptNum)�� ����Ͽ� �ѽ�����
' ����� Ȯ���մϴ�.
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
    
    tmp = "state | result | title | sendnum | senderName | rcv | rcvnm | T | S | F | R | C | receiptDT | reserveDT | sendDT | resultDT | filenames" + vbCrLf
    
    For Each sentFax In sentFaxList
    
        tmp = tmp + CStr(sentFax.state) + " | "             '���ۻ��� �ڵ�
        tmp = tmp + CStr(sentFax.result) + " | "            '���۰�� �ڵ�
        tmp = tmp + sentFax.title + " | "                   '�ѽ�����
        tmp = tmp + sentFax.sendNum + " | "                 '�߽Ź�ȣ
        tmp = tmp + sentFax.senderName + " | "              '�߽��ڸ�
        tmp = tmp + sentFax.receiveNum + " | "              '���Ź�ȣ
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
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = FaxService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = FaxService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �ѽ� �߽Ź�ȣ ����� ��ȸ�մϴ�.
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
' �ѽ� �߽Ź�ȣ ���� �˾� URL�� ��ȯ�մϴ�.
' ������å���� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetURL_SENDER_Click()
    Dim url As String
    
    url = FaxService.GetURL(txtCorpNum.Text, txtUserID.Text, "SENDER")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1231212312"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.CEOName = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.CorpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.Addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.BizType = "����"
    
    '����, �ִ� 40��
    joinData.BizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.ID = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.PWD = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = FaxService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
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
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.ID = "testkorea_20161011"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.PWD = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = FaxService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' �ѽ��� �������մϴ�.
' - �����Ϸκ��� 180���� ������� ���� �Ǹ� �������� �� �ֽ��ϴ�.
' - �߽���/������ ������ �����Ͽ� ������ �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnResendFAX_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    
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
    
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' �ѽ��� �������մϴ�.
' - �����Ϸκ��� 180���� ������� ���� �Ǹ� �������� �� �ֽ��ϴ�.
' - �߽���/������ ������ �����Ͽ� ������ �� �ֽ��ϴ�.
'=========================================================================

Private Sub btnResendFaxSame_Click()
    Dim senderNum As String
    Dim senderName As String
    Dim receivers As New Collection
    Dim receiver As New PBReceiver
    Dim receiptNum As String
    Dim i As Integer
    
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
    
    receiptNum = FaxService.ResendFAX(txtCorpNum.Text, txtReceiptNum.Text, senderNum, senderName, receivers, txtReserveDT.Text)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �˻������� ����Ͽ� �ѽ����� ������ ��ȸ�մϴ�.
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
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    SDate = "20170601"
    
    '[�ʼ�] ��������, ����(yyyyMMdd)
    EDate = "20171231"
    
    '���ۻ��� �迭, 1(���), 2(����), 3(����), 4(���)
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '�������� �˻�����, True-�������۰� ��ȸ, False-��ü��ȸ
    ReserveYN = False
    
    '������ȸ ����, True-������ȸ, False-ȸ����ȸ
    SenderOnly = False
    
    '������ ��ȣ, �⺻�� 1
    Page = 1
    
    '�������� ��ϰ���, �⺻�� 500
    PerPage = 30
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    Set faxSearchList = FaxService.Search(txtCorpNum.Text, SDate, EDate, state, ReserveYN, SenderOnly, Page, PerPage, Order)
     
    If faxSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code (�����ڵ�) : " + CStr(faxSearchList.code) + vbCrLf
    tmp = tmp + "total (�� �˻���� �Ǽ�) : " + CStr(faxSearchList.total) + vbCrLf
    tmp = tmp + "perPage (�������� ��ϰ���) : " + CStr(faxSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (������ ��ȣ) : " + CStr(faxSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (������ ����) : " + CStr(faxSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (����޽���) : " + faxSearchList.message + vbCrLf + vbCrLf
    
    MsgBox tmp
    
    tmp = "state | result | title | sendnum | senderName | rcv | rcvnm | T | S | F | R | C | receiptDT | reserveDT | sendDT | resultDT | fileNames" + vbCrLf
    
    For Each sentFax In faxSearchList.list
    
        tmp = tmp + CStr(sentFax.state) + " | "             '���ۻ��� �ڵ�
        tmp = tmp + CStr(sentFax.result) + " | "            '���۰�� �ڵ�
        tmp = tmp + sentFax.title + " | "                   '�ѽ�����
        
        tmp = tmp + sentFax.sendNum + " | "                 '�߽Ź�ȣ
        tmp = tmp + sentFax.senderName + " | "              '�߽Ź�ȣ
        tmp = tmp + sentFax.receiveNum + " | "              '���Ź�ȣ
        tmp = tmp + sentFax.receiveName + " | "             '�����ڸ�
        
        tmp = tmp + CStr(sentFax.sendPageCnt) + " | "       '��������
        tmp = tmp + CStr(sentFax.successPageCnt) + " | "    '���� ��������
        tmp = tmp + CStr(sentFax.failPageCnt) + " | "       '���� ��������
        tmp = tmp + CStr(sentFax.refundPageCnt) + " | "     'ȯ�� ��������
        tmp = tmp + CStr(sentFax.cancelPageCnt) + " | "     '��� ��������
        
        tmp = tmp + sentFax.receiptDT + " | "               '�����Ͻ�
        tmp = tmp + sentFax.reserveDT + " | "               '���������Ͻ�
        tmp = tmp + sentFax.sendDT + " | "                  '�����Ͻ�
        tmp = tmp + sentFax.resultDT + " | "                '���۰�� �����Ͻ�
        
        i = 0
        
        For Each fileName In sentFax.fileNames              '���� ���ϸ�
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
' �ѽ� ���۳��� ��� �˾� URL�� ��ȯ�մϴ�.
' ������å���� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnSearchPopUp_Click()
    Dim url As String
    
    url = FaxService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
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
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
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
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
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
    
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
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
                
    receiptNum = FaxService.SendFAX(txtCorpNum.Text, senderNum, receivers, FilePaths, txtReserveDT.Text, txtUserID.Text, senderName, adsYN, title)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "������ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �ѽ� ���۴ܰ��� Ȯ���մϴ�.
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
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.ID = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-4304-2991"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
                
    Set Response = FaxService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.CEOName = "��ǥ��"
    
    '��ȣ
    CorpInfo.CorpName = "��ȣ"
    
    '�ּ�
    CorpInfo.Addr = "����Ư����"
    
    '����
    CorpInfo.BizType = "����"
    
    '����
    CorpInfo.BizClass = "����"
    
    Set Response = FaxService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(FaxService.LastErrCode) + vbCrLf + "����޽��� : " + FaxService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub


Private Sub Form_Load()
    FaxService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    FaxService.IsTest = True
        
End Sub

