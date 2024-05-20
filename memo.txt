Dim FSO As New FileSystemObject

Sub Tree()
Dim folderPath As String
Const Msg1_txt As String = "スキャンしたいフォルダを選択してください。"
Const Msg1_title As String = "フォルダ選択"
MsgBox Msg1_txt, vbInformation, Msg1_title
With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    folderPath = .SelectedItems(1)
End With

With Worksheets
    .Add
    .Select
End With
ActiveSheet.Name = Dir(folderPath, vbDirectory)

Call serch(folderPath)
End Sub
Function serch(tPath As String, Optional cntRow As Integer = 1, Optional cntDepth As Integer = 1) As Integer
Dim Rootfolder As Folder '引数で受けたパスのフォルダ
Dim subfolder As Folder 'ルートフォルダのサブフォルダ
Dim subfolderCnt As Integer: subfolderCnt = 1 'サブフォルダ何個目の処理化をカウントする変数
Dim tfile As File 'forEachでフォルダ内のファイルを走るオブジェクト変数
Set Rootfolder = FSO.GetFolder(tPath)
Dim i As Integer

'フォルダ名記述
Cells(cntRow, cntDepth) = Rootfolder.Name
cntRow = cntRow + 1

'ファイル名を羅列していく
For Each tfile In Rootfolder.Files '階層構造をあらわす縦棒を引き継ぐ
    For i = 1 To cntDepth - 1
        If Cells(cntRow - 1, i).Value <> "└" And Cells(cntRow - 1, i).Value <> "" Then
            Cells(cntRow, i).Value = "│"
        End If
    Next i
    If Rootfolder.SubFolders.Count > 0 Then 'もしフォルダにサブフォルダが無ければ縦棒は付けない。
        Cells(cntRow, cntDepth).Value = "│"
    End If
    Cells(cntRow, cntDepth + 1).Value = tfile.Name
    cntRow = cntRow + 1
Next tfile


If Rootfolder.Files.Count > 0 Then 'フォルダごとに空白行を一行入れる
    For i = 1 To cntDepth
        If Cells(cntRow - 1, i).Value <> "└" And Cells(cntRow - 1, i).Value <> "" Then
            Cells(cntRow, i).Value = "│"
        End If
    Next i
    cntRow = cntRow + 1
End If


'フォルダ名を羅列していく
For Each subfolder In Rootfolder.SubFolders
    If cntDepth > 1 Then
        For i = 1 To cntDepth - 1 '階層構造をあらわす縦棒を引き継ぐ
            If Cells(cntRow - 1, i).Value <> "└" And Cells(cntRow - 1, i).Value <> "" Then
                Cells(cntRow, i).Value = "│"
            End If
        Next i
    End If
    If subfolderCnt = Rootfolder.SubFolders.Count Then 'もし親フォルダの中でこのサブフォルダが最後の場合、"├"ではなく"└"を用いる。
        Cells(cntRow, cntDepth) = "└"
    Else
        Cells(cntRow, cntDepth) = "├"
    End If
    
    'サブフォルダの中のサブフォルダやファイルも書き出すために、この関数を再帰的に定義。戻り値はワークシート記入中最終行を意味する最新のcntRow
    cntRow = serch(subfolder.Path, cntRow, cntDepth + 1)

    subfolderCnt = subfolderCnt + 1
Next subfolder
serch = cntRow
End Function
