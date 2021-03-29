
set zip = CreateObject("Chilkat_9_5_0.Zip")
MsgBox zip.Version

success = zip.UnlockComponent("test")
MsgBox zip.LastErrorText



