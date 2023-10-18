# Excel操作ツール

このツールは、指定されたExcelファイルから選択されたシートを読み込み、数値を削除して新しいExcelファイルとして保存します。

## 使い方

1. スクリプトを実行すると、まずExcelファイルを選択するダイアログが表示されます。
2. 次に、出力する新しいExcelファイルの名前と保存先ディレクトリを指定します。
3. 最後に、処理するシートを選択するダイアログが表示されます。
4. すべての選択が完了すると、新しいExcelファイルが指定した保存先に生成されます。

## 依存関係

- PySimpleGUI
- openpyxl
* * *
* * *
# 在庫管理ツール

このツールは、在庫の追加、減少、現在の残量を自動計算してCSVファイルにて管理を実施します。

### セットアップ

1. `inventory.csv`という名前のCSVファイルを作成して、同じディレクトリに保存します。このファイルは在庫データを保存するために使用されます。
2. Pythonスクリプトを実行してアプリケーションを起動します。

## 使い方
### 商品の追加

1. 「商品名」フィールドに商品名を入力します。
2. 「数量」フィールドに数量を入力します。
3. 「単位」フィールドに単位（例：個、kg、パックなど）を入力します。
4. 「追加」ボタンをクリックします。

### 商品の削除

1. 「商品名」フィールドに商品名を入力します。
2. 「数量」フィールドに削除する数量を入力します。
3. 「削除」ボタンをクリックします。

### 在庫の確認

- アプリケーションの下部にあるリストボックスで現在の在庫を確認できます。

### アプリケーションの終了

- 「終了」ボタンをクリックしてアプリケーションを終了します。

## 注意点

- 在庫が不足している場合、削除操作はできません。
- 無効な入力（例：負の数、空のフィールドなど）は受け付けられません。


## 依存関係

- PySimpleGUI
- csv

#### 追加機能
- 画面に数と単位を表示できます。
- 表形式に表示することで暫定化が決定します。
- 「完全削除」ボタンを追加しています。
- 適切な入力がされていない場合、各ボタンごとにエラーポップアップが表示されます。
    - 商品名が入力されていない場合はエラーを表示します。
    - 商品名に数字や記号が含まれている場合はエラーを表示します。
- 完全削除ボタン押下時
    - 商品名が存在しない場合はエラーを表示します。
- 追加ボタン押下時
    - 数字が入っていない場合はエラーを表示します。
    - 数量が数字出ない場合はエラー文を出力します。
    - 数量が0以下の場合はエラー文を出力します。
    - 単位が入力されていない場合はエラーを表示します。
    - 単位に数字や記号が含まれている場合、エラーポップアップを表示しています。
- 削除ボタン押下時
    - 数量が数字出ない場合はエラー文を出力します。
    - 数量が0以下の場合はエラー文を出力します。
    - 商品名が存在しない場合はエラーを表示します。
    - 商品が0以下になる場合は在庫が足りない旨のエラーを表示します。

- 「完全削除」が押されたときには、警告ポップアップが表示され、OKが押された場合のみ削除が行われます。

* * *
* * *
# 週次販売レポート生成ツールの使用方法

## 概要
このツールは、CSV形式の売上データから週次の販売レポートを生成します。

## インストール
- Python 3.x
- pandas
- PySimpleGUI

## 使用方法

### 1. 売上データをアップロード
- 「売上データをアップロード」欄の「ファイル選択」ボタンをクリックして、売上データのCSVファイルを選択します。

### 2. エクスポート先フォルダを指定
- 「エクスポート先フォルダ」欄の「フォルダ選択」ボタンをクリックして、レポートの保存先となるフォルダを選択します。

### 3. レポート生成
- 「レポート生成」ボタンをクリックすると、選択した売上データからレポートが生成されます。

### 4. 集計結果確認
- 「集計結果確認」ボタンをクリックすると、新しいウィンドウが開き、商品別および日別の集計結果が表示されます。
- 集計結果を確認した後、「閉じる」ボタンをクリックしてウィンドウを閉じます。

### 5. レポートダウンロード
- 「レポートダウンロード」ボタンをクリックすると、生成されたレポートがダウンロードされます。

## 注意点
- 「×」ボタンまたは「閉じる」ボタンをクリックすると、各ウィンドウが閉じます。
- エラーメッセージが表示された場合は、指示に従ってください。