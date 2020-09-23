Attribute VB_Name = "modAIMSends"
Function AwayMessage(strType As String, strMessage As String)
AwayMessage = Chr(0) & Chr(2) & Chr(0) & Chr(4) & String(5, Chr(0)) & Chr(4) & Chr(0) & Chr(3) & TwoByteLen(strType) & Chr(0) & Chr(4) & TwoByteLen(strMessage)
End Function

Function UserWarning(strType As String, strUserName As String)
UserWarning = Chr(0) & Chr(4) & Chr(0) & Chr(8) & Chr(0) & Chr(0) & strType & Chr(Len(strUserName)) & strUserName
End Function

Function BlockUser(strUserName As String, intType As Integer)
BlockUser = Chr(0) & Chr(19) & Chr(0) & Chr(intType) & Chr(0) & Chr(0) & Chr(0) & Chr(6) & Chr(0) & Chr(intType) & TwoByteLen(strUserName) & Chr(0) & Chr(0) & Chr(11) & Chr(17) & Chr(0) & Chr(3) & Chr(0) & Chr(0)
End Function

Function BuddyIconEdit(strRequestID As String, strBuddyIcon As String, strIconThing)
BuddyIconEdit = ChrA("0 5") & TwoByteLen(ChrA("0 0") & strRequestID & ChrA("9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 39 17 3 136 0 0 131 211 0 0") & IntegerToBase256(Len(strBuddyIcon)) & strIconThing & strBuddyIcon & "AVT1picture.id") & ChrA("0 3 0 0")
End Function

Function FileSendEdit(strRequestID As String, strFileName As String)
FileSendEdit = ChrA("0 5") & TwoByteLen(ChrA("0 0") & strRequestID & ChrA("9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 0 3 0 4 65 13 107 17 0 5 0 2 20 70 39 17") & TwoByteLen(ChrA("0 1 0 1 0 0 3 110") & strFileName & ChrA("0 0 0 0 0 0 0"))) & ChrA("0 3 0 0")
End Function

Function BuddyListEdit(strRequestID As String, strBuddyList As String)
BuddyListEdit = ChrA("0 5") & TwoByteLen(ChrA("0 0") & strRequestID & ChrA("9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 15 0 0 39 17") & TwoByteLen(strBuddyList)) & ChrA("0 3 0 0")
End Function

Function BuddyListForm(strBuddyList)
Dim K1, K2, K3, K4, K5() As String
K5 = Split(strBuddyList, ";")
BuddyListForm = TwoByteLen(K5(0))
BuddyListForm = BuddyListForm & IntegerToBase256(UBound(K5))
For i = 1 To UBound(K5)
    BuddyListForm = BuddyListForm & TwoByteLen(K5(i))
    DoEvents
Next i
End Function

Function BlankAttach(strRequestID As String, strCapa As String)
BlankAttach = ChrA("0 5") & TwoByteLen(ChrA("0 0") & strRequestID & strCapa & ChrA("0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 15 0 0 39 17 0 4 0 0 0 1")) & ChrA("0 3 0 0")
End Function

Function IMConnect(strRequestID As String)
IMConnect = ChrA("0 5") & TwoByteLen(ChrA("0 0") & strRequestID & ChrA("9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0 0 10 0 2 0 1 0 3 0 4 24 16 172 135 0 5 0 2 20 70 0 15 0 0")) & ChrA("0 3 0 0")
End Function

'Capabilities
'Buddy Icon: 9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0
'File Send:  9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0
'Buddy List: 9 70 19 75 76 127 17 209 130 34 68 69 83 84 0 0
'Talk:       9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0
'Chat:       116 143 36 32 98 135 17 209 130 34 68 69 83 84 0 0
'IM Image:   9 70 19 69 76 127 17 209 130 34 68 69 83 84 0 0
