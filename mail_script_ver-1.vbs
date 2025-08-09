'運用にあたっての注意事項
' 半角かっこ()も含めて、設定項目を上書きすること
' 例：(項目) → 値
' 
' 送信先、添付ファイル、件名、本文の4つについて、
' 使用するたびに確認を行い、必ず「上書き保存してから」
' 実行をすること
'
' エラーメッセージが出てしまった場合は、焦ってそれを消すことなく、
' 何行目に対する指摘なのかを確認することを推奨します

'メール送信前のおまじない～ここから～

Dim result
result = MsgBox ("メール送信スクリプトを実行してよろしいですか？", vbYesNo + vbDefaultButton2, "確認")
If result = vbNo Then
  WScript.Quit
End If

'メール送信前のおまじない～ここまで～

Set objMail = CreateObject("CDO.Message")

objMail.From = "(Fromにしたいメールアドレス)"

'↓↓ 送信先設定箇所！！ ↓↓

'objMail.To = "(Toにしたいメールアドレス)"

'objMail.Cc = "(Ccにしたいメールアドレス)"


'添付ファイル(ファイルパスを指定する)
'↓↓ 逐一更新・確認すること ↓↓
'↓↓ ファイル単一の場合については、片方を ' でコメントアウトすること ↓↓
'objMail.AddAttachment "d:\test.txt"
'objMail.AddAttachment "d:\hoge.txt"

'件名設定箇所
'↓↓ 間違いがないか確認すること
objMail.Subject = "test"

'本文設定箇所
'↓↓ (苗字)、(名前)、(担当) について、
'↓↓ 必要があれば更新する
objMail.TextBody = _
 "(苗字) (名前)　様" & vbNewLine & _
 " " & vbNewLine & _
 "いつもお世話になっております、" & vbNewLine & _
 "（担当）でございます。" & vbNewLine & _
 " " & vbNewLine & _
 " " & vbNewLine & _
 "今後ともよろしくお願いいたします。" & vbNewLine & _
 " " & vbNewLine & _
 " " & vbNewLine & _
 "□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□ " & vbNewLine & _
 "    " & vbNewLine & _
 "     (担当)" & vbNewLine & _
 " " & vbNewLine & _
 "     Tel: ****" & vbNewLine & _
 "     Fax: ****" & vbNewLine & _
 "     Mail: ****" & vbNewLine & _
 " " & vbNewLine & _
 "□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□ " & vbNewLine & _
 " "



'送信するための設定箇所、原則として手を加えないこと
strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
With objMail.Configuration.Fields
'送信方法　
	'1:ローカルSMTPサービスのピックアップ・ディレクトリにメールを配置する
	'2:SMTPポートに接続して送信 
	'3:OLE DBを利用してローカルのExchangeに接続する
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'SMTPサーバを指定(ホスト名orIP)
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "(SMTPサーバ名)"
	'SMTPポート
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	'SSL通信をする/しない
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
	'SMTP認証 1(Basic認証)/2(NTLM認証）
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	'SMTP送信ユーザ名
	.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "(SMTP送信ユーザ名)"
	'SMTP送信ユーザパスワード
	.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "(pass)"
	'タイムアウト
	.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	.Update
end With

objMail.Send

Set objMail = Nothing