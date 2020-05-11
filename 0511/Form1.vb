Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'DataGridView を初期値に設定
        DGVClear(DataGridView1)    '初期化のSub プロシージャを Call

        '※ 通常は必要ありませんが、Tips の動作確認のために表示状態を元に戻す場合や
        '　 データファイルを読み込み直す場合等に必要なので
    End Sub

    Private Sub DGVClear(ByVal dgv As DataGridView)
        'DataGridView を初期値に設定するプロシージャ
        With dgv
            '列数が>0なら表示されていると判断し、一旦消去(表示速度には影響なし)
            If .Rows.Count > 0 Then
                .Columns.Clear()                            'コレクションを空にします(行・列削除)
                .DataSource = Nothing                       'DataSource に既定値を設定
                .DefaultCellStyle = Nothing                 'セルスタイルを初期値に設定
                .RowHeadersDefaultCellStyle = Nothing       '行ヘッダーを初期値に設定
                .RowHeadersVisible = True                   '行ヘッダーを表示
                'フォントの高さ＋9 (フォントサイズ 9 の場合、12+9= 21 となる
                .RowTemplate.Height = 21                    'デフォルトの行の高さを設定(表示後では有効にならない)
                .ColumnHeadersDefaultCellStyle = Nothing    '列ヘッダーを初期値に設定
                .ColumnHeadersVisible = True                '列ヘッダーを表示
                .ColumnHeadersHeight = 23                   '列ヘッダーの高さを既定値に設定
                .TopLeftHeaderCell = Nothing                '左端上端のヘッダーを初期値に設定
                '奇数行に適用される既定のセルスタイルを初期値に設定　
                .AlternatingRowsDefaultCellStyle = Nothing
                'セルの境界線スタイルを初期値(一重線の境界線)に設定
                .AdvancedCellBorderStyle.All = DataGridViewAdvancedCellBorderStyle.Single
                .GridColor = SystemColors.ControlDark       'セルを区切るグリッド線の色を初期値に設定
                .Refresh()                                  '再描画
            End If
        End With
        '※ 上記設定は、必要により、追加・削除してください。
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'CSV ファイルを ADO.NET を使って DataGridView に読み込み表示
        DGVClear(DataGridView1)                             '初期化のSub プロシージャを Call
        Using cn As New System.Data.OleDb.OleDbConnection
            'データファイルは、EXE と同じフォルダーに入れてください。
            'データのあるフォルダー(プログラム起動フォルダーのパスを指定)
            Dim FolderPath As String = "C:\dgvdat"       ' Application.StartupPath
            'CSV ファイル名 (フルパスで書かないで下さい)
            Dim dbFileName As String = "dgvtest1.csv"
            cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & FolderPath &
                    ";Extended Properties=""Text;HDR=YES;IMEX=1;FMT=Delimited"""
            Using da As System.Data.OleDb.OleDbDataAdapter =
                New System.Data.OleDb.OleDbDataAdapter("SELECT * FROM " & dbFileName, cn)
                Dim ds As New DataSet
                da.Fill(ds, dbFileName)
                'DataGridView に表示するデータソースを設定
                DataGridView1.DataSource = ds.Tables(dbFileName)
            End Using
        End Using
        'ヘッダーとすべてのセルの内容に合わせて、列の幅を自動調整する
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Excel ファイル(xls)を ADO.NET を使って DataGridView に読み込み表示
        DGVClear(DataGridView1)                             '初期化のSub プロシージャを Call
        Using cn As New System.Data.OleDb.OleDbConnection
            Using cm As New System.Data.OleDb.OleDbCommand
                Using da As New System.Data.OleDb.OleDbDataAdapter
                    'Excelファイルのフルパスを設定
                    Dim dbFileName As String = "C:\dgvdat\dgvtest2.xls"
                    'Excelファイルのシート名を設定
                    Dim SheetName As String = "Sheet1"
                    'データベースに接続するための情報を設定する
                    cn.ConnectionString = "provider=Microsoft.jet.OLEDB.4.0;Data source=" &
                            dbFileName & ";Extended properties=""Excel 8.0;HDR=YES;IMEX=1"""
                    'コネクションの設定
                    cm.Connection = cn
                    'データソースで実行するSQL文の設定
                    cm.CommandText = "select * from [" & SheetName & "$]"
                    '氏名に[子 or 正]の文字が含まれているデータを抽出して表示する場合
                    'cm.CommandText = "Select * from [" & SheetName & "$] WHERE 氏名 LIKE '%子%' or 氏名 LIKE '%正%'"

                    'データソース内のレコードを選択するためのSQLコマンドの設定
                    da.SelectCommand = cm
                    Dim ds As New DataSet
                    da.Fill(ds, SheetName)
                    'DataGridView に表示するデータソースを設定
                    DataGridView1.DataSource = ds.Tables(SheetName)
                End Using
            End Using
        End Using
        'ヘッダーとすべてのセルの内容に合わせて、列の幅を自動調整する
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Excel ファイル(xlsx)を ADO.NET を使って DataGridView に読み込み表示
        DGVClear(DataGridView1)                             '初期化のSub プロシージャを Call
        Using cn As New System.Data.OleDb.OleDbConnection
            Using cm As New System.Data.OleDb.OleDbCommand
                Using da As New System.Data.OleDb.OleDbDataAdapter
                    'Excelファイルのフルパスを設定
                    Dim dbFileName As String = "C:\dgvdat\dgvtest3.xlsx"
                    'Excelファイルのシート名を設定
                    Dim SheetName As String = "Sheet1"
                    'データベースに接続するための情報を設定する
                    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" &
                dbFileName & ";Extended properties=""Excel 8.0;HDR=YES;IMEX=1"""
                    'コネクションの設定
                    cm.Connection = cn
                    'データソースで実行するSQL文の設定
                    cm.CommandText = "select * from [" & SheetName & "$]"
                    '氏名に[子 or 正]の文字が含まれているデータを抽出して表示する場合
                    'cm.CommandText = "Select * from [" & SheetName & "$] WHERE 氏名 LIKE '%子%' or 氏名 LIKE '%正%'"

                    'データソース内のレコードを選択するためのSQLコマンドの設定
                    da.SelectCommand = cm
                    Dim ds As New DataSet
                    da.Fill(ds, SheetName)
                    'DataGridView に表示するデータソースを設定
                    DataGridView1.DataSource = ds.Tables(SheetName)
                End Using
            End Using
        End Using
        'ヘッダーとすべてのセルの内容に合わせて、列の幅を自動調整する
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'mdb ファイルを ADO.NET を使って DataGridView に読み込み表示
        DGVClear(DataGridView1)                             '初期化のSub プロシージャを Call
        Using cn As New System.Data.OleDb.OleDbConnection
            Using cm As New System.Data.OleDb.OleDbCommand
                Using da As New System.Data.OleDb.OleDbDataAdapter
                    Dim ds As New DataSet
                    'EXE と同じフォルダーにデータも入れておく
                    Dim dbFileName As String = "C:\dgvdat\dgvtest4.mdb"
                    Dim TableName As String = "Table1"   '指定のテーブル名(上記ファイル内に存在する事)
                    '接続文字列については、WEB上で、[接続文字列]をキーに検索して見て下さい。
                    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                    "Data Source=" & dbFileName & ";"  'パスワード等があれば続けて記入
                    'コネクションの設定
                    cm.Connection = cn
                    'データソースで実行するSQL文の設定
                    cm.CommandText = "SELECT * from " & TableName
                    'データソース内のレコードを選択するためのSQLコマンドの設定
                    da.SelectCommand = cm
                    'データを取得する
                    da.Fill(ds, TableName)
                    'データグリッドに表示するデータソースを設定
                    DataGridView1.DataSource = ds
                    'グリッドを表示するための、DataSource 内のリストを設定
                    DataGridView1.DataMember = TableName
                    'データソースへの接続を閉る
                End Using
            End Using
        End Using
        'ヘッダーとすべてのセルの内容に合わせて、列の幅を自動調整する
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        'accdb ファイルを ADO.NET を使って DataGridView に読み込み表示
        DGVClear(DataGridView1)                             '初期化のSub プロシージャを Call
        Using cn As New System.Data.OleDb.OleDbConnection
            Using cm As New System.Data.OleDb.OleDbCommand
                Using da As New System.Data.OleDb.OleDbDataAdapter
                    Dim ds As New DataSet
                    'EXE と同じフォルダーにデータも入れておく
                    Dim dbFileName As String = "C:\dgvdat\dgvtest5.accdb"
                    Dim TableName As String = "Table1"   '指定のテーブル名(上記ファイル内に存在する事)
                    '接続文字列については、WEB上で、[接続文字列]をキーに検索して見て下さい。
                    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                                          "Data Source=" & dbFileName & ";"     'パスワード等があれば続けて記入
                    'コネクションの設定
                    cm.Connection = cn
                    'データソースで実行するSQL文の設定
                    cm.CommandText = "SELECT * from " & TableName
                    'データソース内のレコードを選択するためのSQLコマンドの設定
                    da.SelectCommand = cm
                    'データを取得する
                    da.Fill(ds, TableName)
                    'データグリッドに表示するデータソースを設定
                    DataGridView1.DataSource = ds
                    'グリッドを表示するための、DataSource 内のリストを設定
                    DataGridView1.DataMember = TableName
                    'データソースへの接続を閉る
                End Using
            End Using
        End Using
        'ヘッダーとすべてのセルの内容に合わせて、列の幅を自動調整する
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'ヘッダーとすべてのセルの内容に合わせて、行の高さを自動調整する
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        'DataGridView に表示中のデータを CSV 形式で保存
        CsvFileSave("C:\dgvdat\saveTest1.csv")
    End Sub

    Private Sub CsvFileSave(ByVal SaveFileName As String)
        'DataGridView に表示中のデータを CSV 形式で保存用のプロシージャ
        'VB のソースコードのようなデータも保存できるように設定してあり、普通のCSVファイルも保存できます。
        Dim dbFileName As String = SaveFileName
        '現在のファイルに上書き保存
        Using swCsv As New System.IO.StreamWriter(dbFileName, False, System.Text.Encoding.GetEncoding("SHIFT_JIS"))
            Dim sf As String = Chr(34)          'データの前側の括り
            Dim se As String = Chr(34) & ","    'データの後ろの括りとデータの区切りの ","　
            Dim i, j As Integer
            Dim WorkText As String = ""         '1個分のデータ保持用
            Dim LineText As String = ""         '1列分のデータ保持用

            With DataGridView1
                'ヘッダー部分の取得・保存(保存する必要がなければいらない）
                For i = 0 To .Columns.Count - 1
                    WorkText = .Columns.Item(i).HeaderText
                    If WorkText.IndexOf(Chr(34)) > -1 Then                  'データ内に " があるか検索
                        WorkText = WorkText.Replace("""", """""")           'あれば " を "" に置換える
                    End If
                    If i = .Columns.Count - 1 Then                          'ヘッダー行を列分連結
                        LineText &= sf & .Columns.Item(i).HeaderText & sf   '最後の列の場合
                    Else
                        LineText &= sf & .Columns.Item(i).HeaderText & se
                    End If
                Next i
                swCsv.WriteLine(LineText)                               'ヘッダーの部分の書き込み
                '最下部の新しい行（追加オプション）を非表示にする
                DataGridView1.AllowUserToAddRows = False
                '実データ部分の取得・保存処理
                For i = 0 To .RowCount - 1
                    LineText = ""                                       '１行分のデータをクリア
                    For j = 0 To .Columns.Count - 1                     '１行分のデータを取得処理
                        WorkText = .Item(j, i).Value.ToString           '１個セルデータを取得
                        If WorkText.IndexOf(Chr(34)) > -1 Then          'データ内に " があるか検索
                            WorkText = WorkText.Replace("""", """""")   'あれば " を "" に置換える
                        End If
                        If j = .Columns.Count - 1 Then                  '１行分の列データを連結
                            LineText &= sf & WorkText & sf              '最後の列の場合
                        Else
                            LineText &= sf & WorkText & se
                        End If
                    Next j
                    swCsv.WriteLine(LineText)                           '1行分のデータを書き込み
                Next i
            End With
        End Using
        MessageBox.Show("現在表示中のデータを " & dbFileName & " に保存しました。")
    End Sub
End Class
