On Error Resume Next
Set g_oSB = CreateObject("System.Text.StringBuilder")
Function sprintf(sFmt, aData)
	g_oSB.AppendFormat_4 sFmt, (aData)
   sprintf = g_oSB.ToString()
   g_oSB.Length = 0
End Function

