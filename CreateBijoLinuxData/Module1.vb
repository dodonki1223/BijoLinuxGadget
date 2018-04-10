Module Module1

    ''' <summary>美女Linux定数クラス</summary>
    Private Class cBijoLinuxConst

        ''' <summary>画像の拡張子</summary>
        Public Const ImageExtension As String = "jpg"

        ''' <summary>画像のURL</summary>
        Public Const ImageUrl As String = "http://bijo-linux.com/wp-content/uploads/2015/08/{0}." & ImageExtension

        ''' <summary>美女Linuxテーブル名</summary>
        Public Const BijoLinuxTableName = "BijoLinuxInfo"

        ''' <summary>美女Linux情報を格納したXMLファイル名</summary>
        Public Const XmlFileName = "BijoLinuxInfo.xml"

        ''' <summary>XMLファイルのエンコード</summary>
        Public Shared ReadOnly XmlEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

        ''' <summary>コマンドリスト</summary>
        Public Shared ReadOnly CommandList As New Dictionary(Of String, String) _
        From {
                 {"a2ps", "テキストファイルをPostScript形式に変換する"} _
               , {"adduser", "ユーザーを作成する"} _
               , {"alias", "エイリアスを定義または表示する"} _
               , {"alien", "パッケージの形式を変換する"} _
               , {"apropos", "索引データベースからコマンド名をキーワードで調べるコマンド"} _
               , {"apt-cache", "APTライブラリのキャッシュへアクセス"} _
               , {"aptitude", "apt-getより強力なパッケージ管理機能（ソフトウェアの追加・削除）を提供する"} _
               , {"arping", "ARP要求を送信し、相手とのIPレベルでのネットワーク接続状況の確認を調べたり、MACアドレスを取得したりする" & System.Environment.NewLine & "また，ネットワーク内にあるホストのARPキャッシュの更新もする"} _
               , {"bc", "数値計算を行う"} _
               , {"bg", "バックグラウンドでジョブを実行する"} _
               , {"bmptoppm", "BMPファイルをppmフォーマットに変換する"} _
               , {"bmtoa", "ビットマップからASCII文字列へ変換する"} _
               , {"bunzip2", "bzip2形式アーカイブを伸張する"} _
               , {"bzcat", "bzip2形式アーカイブを標準出力で表示する"} _
               , {"bzip2", "bzip2形式に圧縮する"} _
               , {"cal", "カレンダーを表示する"} _
               , {"cat", "ファイルの内容を表示する"} _
               , {"cd", "作業ディレクトリを変更する"} _
               , {"cdrecord", "光学メディア（CD/DVD)に書き込む"} _
               , {"cfdisk", "chdbursesを使って簡単にパーティションを操作する"} _
               , {"chgrp", "ファイルグループの所有権を変更する"} _
               , {"chkconfig", "サービスの自動起動を設定する"} _
               , {"chown", "ファイルの所有権を変更する"} _
               , {"chroot", "ルートディレクトリを変更してコマンドを実行する"} _
               , {"chsh", "ログインシェルを変更する"} _
               , {"colrm", "指定した桁を削除する"} _
               , {"compress", "ファイルの圧縮、展開"} _
               , {"convert", "画像の回転やフォーマット変換などの処理を行う"} _
               , {"cp", "ファイルのコピー"} _
               , {"cpio", "アーカイブへコピー、アーカイブからコピーする"} _
               , {"cu", "他のシステムを呼び出す"} _
               , {"cut", "テキストファイルから一部を選択して表示する"} _
               , {"df", "ディスクの空き領域を表示"} _
               , {"diff", "２つのテキストファイルの違い（差分）を表示する"} _
               , {"dig", "DNSサーバにドメイン名を問い合わせる"} _
               , {"dir", "ディレクトリの内容をリストで表示する"} _
               , {"echo", "1行テキストを表示する"} _
               , {"eject", "フロッピーディスクやCD-ROM、DVD-ROMなどの媒体の取り出しを行う"} _
               , {"env", "環境変数の変更や表示を行う"} _
               , {"fdformat", "フロッピーディスクを初期化"} _
               , {"file", "ファイル形式を調べるため"} _
               , {"find", "ファイルやディレクトリを検索する"} _
               , {"fsck", "ファイルシステムのチェック/修正を行う"} _
               , {"fuser", "ファイルもしくはソケットを使用しているプロセスを特定/シグナルを送信する"} _
               , {"gpg", "電子メールの暗号化/電子署名を行う"} _
               , {"growisofs", "mkisofsと合わせてDVDに書き込む"} _
               , {"grpck", "グループファイルの整合性を照合する"} _
               , {"gzexe", "実行ファイルを実行できる形式で圧縮する"} _
               , {"halt", "システムを停止する"} _
               , {"hdparm", "ハードディスクパラメータの取得や設定を行う"} _
               , {"iwconfig", "無線LANインタフェースの状態の確認および設定を行う"} _
               , {"jobs", "ジョブの稼働状況を表示する"} _
               , {"join", "テキストファイルを同じ項目同士で結合する"} _
               , {"killall", "プロセスを終了する"} _
               , {"last", "最近ログインしたユーザーリストを表示する"} _
               , {"ldconfig", "/sbin/ldconfig - 動的リンカによる実行時の結合関係を設定する"} _
               , {"ldd", "共有ライブラリへの依存関係を表示する"} _
               , {"ln", "ファイルやディレクトリにリンクを張る"} _
               , {"logger", "ログにloggerのプロセスIDを記録する"} _
               , {"lpq", "印刷ジョブを確認する"} _
               , {"lpr", "印刷する"} _
               , {"ls", "現行ディレクトリの内容を表示"} _
               , {"mailq", "メール・キューに入っているメッセージのリストを出力"} _
               , {"make", "プログラム（C、C++など）のビルド"} _
               , {"mdel", "FTPでファイルを複数削除する"} _
               , {"mkfs", "HDDなどをフォーマットする"} _
               , {"mknod", "スペシャル・ファイル用のディレクトリ・エントリーとそれに対応するiノードを作成する"} _
               , {"mkswap", "スワップ領域をデバイス上またはファイル上に準備する"} _
               , {"modprobe", "カーネルモジュールをロードまたはアンロードする"} _
               , {"mount", "ファイル・システムをマウントする"} _
               , {"mt", "磁気テープドライブの操作を制御する"} _
               , {"newaliases", "/etc/aliases ファイルから別名データベースの新しいコピーを作成"} _
               , {"nkf", "文字コードを変換する"} _
               , {"ntpstat", "ntpd から簡単なステータスレポートを取得"} _
               , {"parted", "パーティションの作成や削除"} _
               , {"pstree", "実行中のプロセスをツリー形式で表示"} _
               , {"pushd", "ディレクトリの移動（cdコマンドと違い履歴がスタックに残る）"} _
               , {"pwconv", "pwconv, pwunconv, grpconv, grpunconv - パスワード・グループの shadow 化と、通常ファイルへの逆変換"} _
               , {"pwd", "現在の作業ディレクトリ名を返す"} _
               , {"pwunconv", "シャドウパスワードの非シャドウ化"} _
               , {"rar", "ファイルの解凍"} _
               , {"readlink", "シンボリックリンクのリンク先を確認"} _
               , {"route", "IPパケットをルーティングするためのルーティングテーブルの内容表示と設定"} _
               , {"screen", "仮想端末を複数使ったり、ウィンドウを分割して使ったり、ターミナルを便利にする"} _
               , {"script", "操作内容（端末の入出力）を記録する"} _
               , {"setup", "デバイスとファイルシステム(file systems)の初期化を行い、 ルートファイルシステム(root filesystem)のマウントを行う"} _
               , {"sftp", "安全なファイル転送を行う"} _
               , {"sg", "別のグループIDでコマンドを実行する"} _
               , {"swapon", "/etc/fstabファイルでスワップとして指定されているデバイスをすべて有効にする"} _
               , {"sync", "システム上にある変更済みのディスク・キャッシュ領域の内容をすべてディスクに書き出す"} _
               , {"time", "実行時間を出力する"} _
               , {"tmpwatch", "作成した一時ファイルや古いログファイルの削除や、使われなくなった古いファイルを自動的に削除する"} _
               , {"uniq", "ファイルから重複する行を削除する"} _
               , {"uuidgen", "UUID値を生成する"} _
               , {"vdir", "ディレクトリの内容を一覧表示"} _
               , {"vi", "vimエディタ"} _
               , {"vim", "vimエディタ"} _
               , {"vimdiff", "左右にdiffを表示させる"} _
               , {"vmstat", "仮想メモリやディスクI/Oの統計情報を表示する"}
             }

        ''' <summary>美女Linuxデータを格納するDataTableのカラム列挙体</summary>
        Public Enum BijoLinuxColumns

            ''' <summary>コマンド名</summary>
            CommandName

            ''' <summary>コマンド説明</summary>
            CommandDescription

            ''' <summary>画像ファイル名</summary>
            ImageFile

            ''' <summary>画像ファイルURL</summary>
            ImageFileUrl

        End Enum

    End Class

#Region "メイン処理"

    ''' <summary>コンソールアプリケーションのメイン処理</summary>
    Sub Main()

        '美女LinuxDataTableのカラムを作成
        Dim mDt As DataTable = _CreateBijoLinuxColumns()
        mDt.TableName = cBijoLinuxConst.BijoLinuxTableName

        '行データを作成
        For Each mCommandListInfo As KeyValuePair(Of String, String) In cBijoLinuxConst.CommandList

            mDt.Rows.Add(AddRowDataToDataTable(mDt, mCommandListInfo.Key, mCommandListInfo.Value))

        Next

        'XMLファイルの出力先フルパスを作成 ※Exe実行パスの親フォルダ＋XMLファイル名
        Dim mExePath As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim mXmlOutputPath As String = mExePath.DirectoryName & "\" & cBijoLinuxConst.XmlFileName

        'XMLファイルの出力処理
        Using mWriter = New System.IO.StreamWriter(mXmlOutputPath, False, cBijoLinuxConst.XmlEncoding)

            mDt.WriteXml(mWriter, XmlWriteMode.WriteSchema, True)

        End Using

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    '''   美女LinuxDataTableのカラムを作成する
    ''' </summary>
    ''' <returns>美女Linuxのカラム作成後のDataTable</returns>
    Private Function _CreateBijoLinuxColumns() As DataTable

        Dim mBijoLinux As New DataTable

        '美女Linuxカラム列挙体分繰り返す
        For Each mColumnName As String In System.Enum.GetNames(GetType(cBijoLinuxConst.BijoLinuxColumns))

            'String型で作成
            mBijoLinux.Columns.Add(mColumnName, Type.GetType("System.String"))

        Next

        Return mBijoLinux

    End Function

    ''' <summary>
    '''   コマンドの行データを作成
    ''' </summary>
    ''' <param name="pCommand">コマンド名</param>
    ''' <param name="pCommandDescription">コマンドの説明</param>
    ''' <returns>コマンドの行データ</returns>
    Private Function AddRowDataToDataTable(ByVal pDataTable As DataTable, ByVal pCommand As String, Optional pCommandDescription As String = "") As DataRow

        Dim mAddDataRow As DataRow = pDataTable.NewRow

        mAddDataRow(cBijoLinuxConst.BijoLinuxColumns.CommandName) = pCommand
        mAddDataRow(cBijoLinuxConst.BijoLinuxColumns.CommandDescription) = pCommandDescription
        mAddDataRow(cBijoLinuxConst.BijoLinuxColumns.ImageFile) = pCommand & "." & cBijoLinuxConst.ImageExtension
        mAddDataRow(cBijoLinuxConst.BijoLinuxColumns.ImageFileUrl) = String.Format(cBijoLinuxConst.ImageUrl, pCommand)

        Return mAddDataRow

    End Function

#End Region

End Module
