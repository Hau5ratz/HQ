Sub get_cust_md(EID As String, Password As String, mdx As String)

Dim sts As String

Sheet1.Visible = True
Sheet1.Activate

sts = HypUIConnect("Q", EID, Password, "")
sts = HypExecuteQuery(Empty, mdx)

End Sub
