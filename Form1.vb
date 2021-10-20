Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Shitei As String



        Shitei = TextBox1.Text

        Call Sample3(Shitei)

        MsgBox("終了しました")


    End Sub

    Sub Sample3(Path As String)
        Dim Path1 As String, Path2 As String, ido1 As String, Path3 As String

        Dim fso = CreateObject("Scripting.FileSystemObject")
        For Each f In fso.GetFolder(Path).subfolders

            Path2 = f.name
            If Mid(Path2, 1, 2) <> "R_" Then


                Dim buf As String

                Path1 = Path & "\" & Path2 & "\ひび割れ\成形結果\BPW\"
                Path3 = Path & "\" & Path2 & "\ひび割れ\成形結果\縮小画像"
                ido1 = Path & "\R_" & Path2

                IO.Directory.CreateDirectory(ido1)

                buf = Dir(Path1 & "\*.pgw")


                Do While buf <> ""

                    IO.File.Copy(Path1 & "\" & buf, ido1 & "\" & buf)

                    buf = Dir()

                Loop

                buf = Dir(Path3 & "\*.png")

                Do While buf <> ""

                    IO.File.Copy(Path3 & "\" & buf, ido1 & "\" & buf)

                    buf = Dir()

                Loop
            End If

        Next

    End Sub

    Function CutRight(s, i)
        Dim iLen As Long

        If VarType(s) <> vbString Then
            Exit Function
        End If

        iLen = Len(s)

        '// 文字列長より指定文字数が大きい場合
        If iLen < i Then
            Exit Function
        End If

        '// 指定文字数を削除して返す
        CutRight = Mid(s, 1, iLen - i)
    End Function


End Class