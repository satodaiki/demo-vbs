Option Explicit

' JSONオブジェクトに関する便利クラス
Class JsonUtility

    ' json文字列のJSONオブジェクト変換
    Public Function jsonParse(ByVal jsonStr)

        Dim doc
        Dim jsn

        ' HTMLDocumentを取得
        Set doc = CreateObject("HtmlFile")

        ' JSONオブジェクトを使うにはIEの互換表示で8以上(edgeも可)にする。
        doc.write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"

        ' scriptタグを追加
        ' パースはJSON.parseを使うとチェックが厳しいのでevalを使う。
        doc.write "<script>document.JsonParse=function (s) {return eval('(' + s + ')');}</script>"

        ' パース関数でJSONオブジェクトを取得
        jsonParse = doc.JsonParse(jsonStr)

    End Function

    ' jsonオブジェクトの文字列変換
    Public Function jsonStringify(ByVal jsonObj)

        Dim doc
        Dim jsn

        ' HTMLDocumentを取得
        Set doc = CreateObject("HtmlFile")

        ' JSONオブジェクトを使うにはIEの互換表示で8以上(edgeも可)にする。
        doc.write "<meta http-equiv='X-UA-Compatible' content='IE=8' />"

        ' scriptタグを追加
        doc.write "<script>document.JsonStringify=JSON.stringify;</script>"

        jsonStringify = doc.JsonStringify(jsonObj)

    End Function

End Class