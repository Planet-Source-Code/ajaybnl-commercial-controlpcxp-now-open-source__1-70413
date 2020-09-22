Attribute VB_Name = "SendLog"
Public Sub Log(ErrorString As String)
On Error GoTo err
If IsConnected Then Call URLDownloadToFile(0, "http://members.lycos.co.uk/uptomoon/controlpcxp.php?user=" & GetSetting(App.Title, "Setup", "User", Text2.Text) & "&error=" & Replace(ErrorString, " ", "_"), "tmp3.tmp", 0, 0)
err:
err.Clear
If Dir("tmp3.tmp") <> "" Then Kill "tmp3.tmp"
End Sub
