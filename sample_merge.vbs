Private Sub Toggle15_Click()
    	Dim pathMergeTemplate As String
    	Dim sql As String
	stDocName = "invite_letter"
        DoCmd.OpenQuery stDocName

' Link template

            pathMergeTemplate = Application.CurrentProject.Path & "\Mail_Merges\Recruitment_Packet\"

' Base query for merge fields

            sql = "SELECT * FROM invite_letter"

' Create temp query def 
On Error Resume Next

                Dim qd As DAO.QueryDef
                Set qd = New DAO.QueryDef
                    qd.sql = sql
					' Clear stale query defs if Access crashes
					CurrentDb.QueryDefs.Delete "mmexport"
                    qd.Name = "mmexport"
                    
                    CurrentDb.QueryDefs.Append qd
' Export 
                        DoCmd.TransferText _
                            acExportDelim, , _
                            "mmexport", _
                            pathMergeTemplate & "qryMailMergeRec.txt", _
                            True
' Clean
                    CurrentDb.QueryDefs.Delete "mmexport"

                    qd.Close
                Set qd = Nothing

' Data extracted from Access db into raw txt
' Format
' Initial VA DB has limited data integrity, enforce what's needed

Dim sBuf As String
Dim sTemp As String
Dim iFileNum As Integer
Dim sFileName As String

sFileName = pathMergeTemplate & "qryMailMergeRec.txt"

iFileNum = FreeFile
Open sFileName For Output As iFileNum

Do Until EOF(iFileNum)
    Line Input #iFileNum, sBuf
    sTemp = sTemp & sBuf & vbCrLf
Loop
Close iFileNum

sTemp = Replace(sTemp, "Null", "")
sTemp = Replace(sTemp, "null", "")
sTemp = Replace(sTemp, "NULL", "")

iFileNum = FreeFile
Open sFileName For Output As iFileNum
Print #iFileNum, sTemp
Close iFileNum

' Data cleaned, push to Word template

                Dim appWord As Object
                Dim docWord As Object

                Set appWord = CreateObject("Word.Application")

                    appWord.Application.Visible = True

                    Set docWord = appWord.Documents.Add(Template:=pathMergeTemplate & "Recruitment_Packet_Envelope-95x12.docx")

                        docWord.MailMerge.OpenDataSource Name:=pathMergeTemplate & "qryMailMergeRec.txt", LinkToSource:=False

                    Set docWord = Nothing

                Set appWord = Nothing

End Sub