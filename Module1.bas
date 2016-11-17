Attribute VB_Name = "Module1"
Sub convertXML()
   mySheet = ActiveSheet.Name

   Dim sname As String
   Dim frow As Integer
   Dim drow As Integer
   Dim pname As String
   Dim cname As String
   Dim fname As String
   Dim code As String

   Application.ScreenUpdating = False

   sname = Range("sname")
   frow = Range("frow")
   drow = Range("drow")
   pname = Range("parent")
   cname = Range("child")
   fname = Range("fname")
   code = Range("code")
   char = Range("char")

   Sheets(sname).Select
   Cells(drow, 1).Select
   Selection.End(xlDown).Select
   r = Selection.Row
   Range("A1").Select
   Sheets(mySheet).Select


   Counter.Show vbModeless
   Counter.Label1.Caption = "0"
   Counter.Label2.Caption = "/ " & r


   '改行コード選択
   LineFeedCode = vbCrLf
   If code = "\r" Then LineFeedCode = vbCr
   If code = "\n" Then LineFeedCode = vbLf

   '文字コード選択
   charcode = "SJIS"
   xmlcode = "Shift_JIS"
   If char = "UTF-8" Then
      charcode = "UTF-8"
      xmlcode = "UTF-8"
   End If

   'XML形式に変換
   t = "<?xml version=""1.0"" encoding=""" & xmlcode & """?>" & vbCrLf
   t = t & "<" & pname & ">" & LineFeedCode

   r = drow
   cnt = 1
   Do While Sheets(sname).Cells(r, 1) <> ""
      t = t & "  <" & cname & " id=""" & cnt & """>" & LineFeedCode
      c = 1
      Do While Sheets(sname).Cells(frow, c) <> ""
         t = t & "    <" & Sheets(sname).Cells(frow, c) & ">" & Sheets(sname).Cells(r, c) & "</" & Sheets(sname).Cells(frow, c) & ">" & LineFeedCode
         c = c + 1
      Loop
      t = t & "  </" & cname & ">" & LineFeedCode
      
      Counter.Label1.Caption = r
      DoEvents
   
      r = r + 1
      cnt = cnt + 1
   Loop
   
   t = t & "</" & pname & ">" & LineFeedCode


   '保存
   Dim stm As ADODB.Stream
   Set stm = New ADODB.Stream
   
   stm.Charset = charcode
   stm.Open
   
   stm.WriteText t, adWriteLine
   stm.SaveToFile ActiveWorkbook.Path & "\" & fname, adSaveCreateOverWrite
   stm.Close

   Application.ScreenUpdating = True
   
   Unload Counter

End Sub
