Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
        Dim objMe           As Recipient
        Dim bcc_mail        As New Collection
        Dim i               As Variant
        
        ' BCCで送信したいアドレスを指定する。
        ' （編集箇所はここだけ）---------------------------
        bcc_mail.Add "hogehoge1@hoge.com"
        bcc_mail.Add "hogehoge2@hoge.com"
        ' （編集箇所はここだけ）---------------------------
        
        ' BCCを設定してからメール送信をする。
        For Each i In bcc_mail
            Set objMe = Item.Recipients.Add(i)
                objMe.Type = olBCC
                objMe.Resolve
            Set objMe = Nothing
        Next i
        
End Sub
