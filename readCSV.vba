Option Explicit

Sub readCSV()
　　Dim csvFileName As Variant
　　Dim freeNum As Integer
　　Dim strRec As String
　　Dim strSplit() As String
　　Dim i As Long, j As Long

　　csvFileName = Application.GetOpenFilename(FileFilter:="CSVファイル(*.csv), *.csv", Title:="CSVファイルの選択")
　　If csvFileName = False Then
　　　　Exit Sub
　　End If

    'FreeFile関数で空いてるファイル番号を取得
　　freeNum = FreeFile
　　Open csvFileName For Input As #freeNum
　　
　　i = 0
　　Do Until EOF(freeNum)
　　　　Line Input #freeNum, strRec
　　　　i = i + 1
　　　　strSplit = Split(strRec, ",")
　　　　For j = 0 To UBound(strSplit)
　　　　　　Cells(i, j + 1) = strSplit(j)
　　　　Next
　　Loop
　　
　　Close #freeNum
End Sub
