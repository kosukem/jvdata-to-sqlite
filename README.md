# JVData To SQLite

JRA-VANデータラボの競馬データJVDataをSQLiteデータベースに変換するExcel VBAプロジェクト


## 1. イントロダクション

JVData To SQLiteは[JRA-VAN データラボ](http://jra-van.jp/dlb/)に登録予定のソフトウェアです。データラボの会員（月会費2,058円）は無料で使うことができます。データラボは1ヶ月の無料トライアルがあるので、データベースを更新しなくていいなら無料で使うことができます。

JVData To SQLiteは競馬データJVDataをSQLiteデータベースに変換するExcel VBAプロジェクト（Excelマクロ有効ブック）です。JVData To SQLiteを使えば、1986年から現在までに開催された競馬レースのSQLiteデータベースをローカルマシン上に持つことができ、SQLやRを使って自在に競馬データを解析することができます。

JRA-VANやデータラボについての詳しい説明は上記リンクをご覧ください。


## 2. 目的

JVData To SQLiteの目的は、ボタンを押すだけで、競馬データJVDataをダウンロードして、読み込み、SQLiteデータベースを構築することです。


## 3. 準備

JVData To SQLiteの実体はExcelマクロ有効ブックです。圧縮ファイルをダウンロードして展開してください。

JVData To SQLiteは次の3つのソフトウェアに依存しているので、それぞれWindows 7/8/10にインストールしてください。

1. Microsoft Excel
2. [JVLink](http://jra-van.jp/dlb/index.html#tab5)
3. [SQLite ODBC Driver](http://www.ch-werner.de/sqliteodbc/)


### 3.1 Microsoft Excelのインストールと設定

JVData To SQLiteはExcel 2010で開発しました。2010以降のバージョンのExcelで動作すると思います。

Excelは32bit版を使ってください。JVLinkは64bit版では動作しません。Excelのインストール時に、ユーザーが自ら64bit版を選ぼうとしなければ、32bit版がインストールされるので、ほとんどの方は32bit版のExcelを使っていると思います。

Excelはマクロ機能を有効にしてください。Excelの*ファイル* > *オプション* > *セキュリティセンター*から次のようにチェックを入れてください。


### 3.2 JVLinkのインストールと設定

[JRA-VANデータラボ](http://jra-van.jp/dlb/)の動作環境タブの説明に従ってインストールし、認証キーの入力を行ってください。JVLinkの実体はActiveXコントロール（dllファイル）で、Windowsでのみ動作します。


### 3.3 SQLite ODBC Driver

SQLite ODBC DriverはSQLiteデータベースを操作する機能を提供するドライバです。VBAから（正確にはADODBオブジェクト）SQLiteデータベースを操作するのに必要なソフトウェアです。

[http://www.ch-werner.de/sqliteodbc/](http://www.ch-werner.de/sqliteodbc/)からダウンロードしてインストールしてください。

リンク先のCurrent versionの下にある`sqliteodbc.exe`（32bit版）を選んでください。`sqliteodbc_w64.exe`（64bit版）もあるので間違えないようにしてください。 


## 4. 使い方

### 4.1 基本的なこと

JVData To SQLiteの実体は`JVDataToSQLite.xlsm`ファイル（Excelマクロ有効ブック）です。ダウンロードした圧縮ファイルを展開するとファイルがあります。

`JVDataToSQLite.xlsm`ファイルをExcelで開くと、Topワークシートが開きます。

基本的にユーザーが行うのはTopワークシートの**Start**ボタンを押すだけです。**Start**ボタンを押せばプログラムが始まります。

途中で止めたい場合は、**Stop**ボタン（**Start**ボタンを押すと**Stop**ボタンに変わります）を押します。

再度スタートすると未取得のデータから取得を再開します。

**レコードの取りこぼしを調べる**チェックボックスにチェックを入れてスタートすると取得済みのデータも読み込みます。取得していないレコード（正確にはjvdファイル）があれば取得しデータベースに追加します。

何か不具合があり、Startボタンが押せなくなった場合はResetボタンを押してください。Startボタンが元に戻ります。


### 4.2 作成するデータベース

プログラムが始まると、データベースファイル`setup.sqlite3`ファイルを`JVDataToSQLite.xlsm`と同じフォルダに新規作成します。

データベースのテーブル設計は、Schemaワークシートか[Wiki](https://github.com/kosukem/jvdata-to-sqlite/wiki)にあります。

または、[DB Browser for SQLite](http://sqlitebrowser.org/)などのビューワーで`setup.sqlite3`を開けば、テーブル設計がわかります。


### 4.3 注意事項

初回実行に24時間から48時間かかります。初回実行では、過去30年分のデータの取得とデータベース追加をするからです。

取得対象期間が短いほどJVLinkは1レコードを高速に読み込むので、一度ストップしてから再スタートするとデータの取得時間が短くなります。これはJVLinkのバグとまではいえないまでもちょっとした不具合です。

JVLinkがフリーズすることがあります。その場合はExcelの自動回復機能が働きます。Excel回復ブック（xlsbファイル）が起動している場合はそのファイルを閉じて、`JVDataToSQLite.xlsm`を開いて再スタートしてください。未取得のデータから取得を再開します。
