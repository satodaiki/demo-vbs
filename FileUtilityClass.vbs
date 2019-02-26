Option Explicit

' ファイルに関する便利クラス
Class FileUtility

    Public Function exportFileStr()
        ' 読み込んだファイル文字列
        Dim fileStr

        Dim fso
        Set fso = WScript.CreateObject("Scripting.FileSystemObject")

        ' 読み込みファイルの指定 (相対パスなのでこのスクリプトと同じフォルダに置いておくこと)
        Dim inputFile
        Set inputFile = fso.OpenTextFile("inputText.txt", 1, False, 0)

        ' 読み込みファイルから1行ずつ読み込み、書き出しファイルに書き出すのを最終行まで繰り返す
        Do Until inputFile.AtEndOfStream
          fileStr = fileStr & inputFile.ReadLine
        Loop

        ' バッファを Flush してファイルを閉じる
        inputFile.Close

        exportFileStr = fileStr
    End Function

End Class