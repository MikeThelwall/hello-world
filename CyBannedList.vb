Option Strict Off
Option Explicit On 
Imports System.IO

Module modBannedList
	Const MAX_BANNED_ARRAY_SIZE As Short = 400
	Public sgBannedArray(MAX_BANNED_ARRAY_SIZE) As String
	Public igBannedArrayCount As Integer
	
	Public sgActiveProject As String
	Public sgLinkFileDirectory As String
	Public sgBannedListFileName As String
	Public sgDomainNamesFileName As String
	Public sgInfoDirectory As String
	
	' Use i_FindIfAnyArrayStringsAreSubstringOfGivenStrings (URL, sgBannedArray, 1, igBannedArrayCount)
	Public Function i_InitialiseBannedArray(ByRef sFileName As String, ByRef sSiteName As String) As Integer
        Static iWarning As Integer
        Dim iTab As Integer, srBannedList As StreamReader
        Dim sLine As String
        Dim bAvoidURL As Boolean

        igBannedArrayCount = 0
        If sFileName = "" Or Not File.Exists(sFileName) Then
            i_InitialiseBannedArray = 0
            If iWarning = 0 Then
                iWarning = MsgBox("No banned list found at " & sFileName & vbCrLf & "Do you wish to abort this indexing?", MsgBoxStyle.YesNoCancel)
                If iWarning <> MsgBoxResult.No Then End
                iWarning = 1
            End If
            Exit Function
        End If

        bAvoidURL = False
        srBannedList = New StreamReader(sFileName)
        sLine = srBannedList.ReadLine
        While Not sLine Is Nothing
            If Left(sLine, 1) = "[" And InStr(sLine, "]") > 0 Then ' allow use of [], [D, [S
                If InStr(Mid(sLine, 2, Len(sLine) - 2), sSiteName & ".") > 0 Then
                    bAvoidURL = True
                Else
                    bAvoidURL = False
                End If
            Else 'blank line or line with URL
                If bAvoidURL Then
                    If Len(sLine) > 10 Then '4/1/05 ALLOW http in banned list
                        If Left(sLine, 11) = "http://www." Then sLine = Mid(sLine, 11)
                    End If
                    If Len(sLine) > 7 Then
                        If Left(sLine, 7) = "http://" Then sLine = Mid(sLine, 8)
                    End If
                    iTab = InStr(sLine, vbTab)
                    If iTab > 0 Then sLine = Left(sLine, iTab - 1)
                    If sLine <> "" Then
                        igBannedArrayCount += 1
                        If igBannedArrayCount > MAX_BANNED_ARRAY_SIZE Then
                            MsgBox("Too many banned URLs! Maximum is " & Str(MAX_BANNED_ARRAY_SIZE))
                            srBannedList.Close()
                            Exit Function
                        End If
                        sgBannedArray(igBannedArrayCount) = sLine
                    End If
                End If
            End If
            sLine = srBannedList.ReadLine
        End While
        srBannedList.Close()
        Return igBannedArrayCount
	End Function 'i_InitialiseBannedArray
End Module