'�^�p�ɂ������Ă̒��ӎ���
' ���p������()���܂߂āA�ݒ荀�ڂ��㏑�����邱��
' ��F(����) �� �l
' 
' ���M��A�Y�t�t�@�C���A�����A�{����4�ɂ��āA
' �g�p���邽�тɊm�F���s���A�K���u�㏑���ۑ����Ă���v
' ���s�����邱��
'
' �G���[���b�Z�[�W���o�Ă��܂����ꍇ�́A�ł��Ă�����������ƂȂ��A
' ���s�ڂɑ΂���w�E�Ȃ̂����m�F���邱�Ƃ𐄏����܂�

'���[�����M�O�̂��܂��Ȃ��`��������`

Dim result
result = MsgBox ("���[�����M�X�N���v�g�����s���Ă�낵���ł����H", vbYesNo + vbDefaultButton2, "�m�F")
If result = vbNo Then
  WScript.Quit
End If

'���[�����M�O�̂��܂��Ȃ��`�����܂Ł`

Set objMail = CreateObject("CDO.Message")

objMail.From = "(From�ɂ��������[���A�h���X)"

'���� ���M��ݒ�ӏ��I�I ����

'objMail.To = "(To�ɂ��������[���A�h���X)"

'objMail.Cc = "(Cc�ɂ��������[���A�h���X)"


'�Y�t�t�@�C��(�t�@�C���p�X���w�肷��)
'���� ����X�V�E�m�F���邱�� ����
'���� �t�@�C���P��̏ꍇ�ɂ��ẮA�Е��� ' �ŃR�����g�A�E�g���邱�� ����
'objMail.AddAttachment "d:\test.txt"
'objMail.AddAttachment "d:\hoge.txt"

'�����ݒ�ӏ�
'���� �ԈႢ���Ȃ����m�F���邱��
objMail.Subject = "test"

'�{���ݒ�ӏ�
'���� (�c��)�A(���O)�A(�S��) �ɂ��āA
'���� �K�v������΍X�V����
objMail.TextBody = _
 "(�c��) (���O)�@�l" & vbNewLine & _
 " " & vbNewLine & _
 "���������b�ɂȂ��Ă���܂��A" & vbNewLine & _
 "�i�S���j�ł������܂��B" & vbNewLine & _
 " " & vbNewLine & _
 " " & vbNewLine & _
 "����Ƃ���낵�����肢�������܂��B" & vbNewLine & _
 " " & vbNewLine & _
 " " & vbNewLine & _
 "�������������������������������������������������������������������� " & vbNewLine & _
 "    " & vbNewLine & _
 "     (�S��)" & vbNewLine & _
 " " & vbNewLine & _
 "     Tel: ****" & vbNewLine & _
 "     Fax: ****" & vbNewLine & _
 "     Mail: ****" & vbNewLine & _
 " " & vbNewLine & _
 "�������������������������������������������������������������������� " & vbNewLine & _
 " "



'���M���邽�߂̐ݒ�ӏ��A�����Ƃ��Ď�������Ȃ�����
strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
With objMail.Configuration.Fields
'���M���@�@
	'1:���[�J��SMTP�T�[�r�X�̃s�b�N�A�b�v�E�f�B���N�g���Ƀ��[����z�u����
	'2:SMTP�|�[�g�ɐڑ����đ��M 
	'3:OLE DB�𗘗p���ă��[�J����Exchange�ɐڑ�����
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'SMTP�T�[�o���w��(�z�X�g��orIP)
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "(SMTP�T�[�o��)"
	'SMTP�|�[�g
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	'SSL�ʐM������/���Ȃ�
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
	'SMTP�F�� 1(Basic�F��)/2(NTLM�F�؁j
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'SMTP���M���[�U��
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "(SMTP���M���[�U��)"
	'SMTP���M���[�U�p�X���[�h
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "(pass)"
	'�^�C���A�E�g
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	.Update
end With

objMail.Send

Set objMail = Nothing