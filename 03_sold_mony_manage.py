import PySimpleGUI as sg
import pandas as pd
import os
from datetime import datetime, timedelta

# GUIレイアウトの設定
layout = [
    [sg.Text("売上データをアップロード:"), sg.Input(), sg.FileBrowse("ファイル選択")],
    [sg.Text("エクスポート先フォルダ:"), sg.Input(), sg.FolderBrowse("フォルダ選択")],
    [sg.Button("レポート生成")],
    [sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')],
    [sg.Button("レポートダウンロード")],
    [sg.Text("", key='error_message')],
]
window = sg.Window("週次販売レポート生成器", layout, size=(900, 600))

while True:
    event, values = window.read()

    if event == "レポート生成":
        # データバリデーションとエラーハンドリング
        try:
            df = pd.read_csv(values[0])  # ここでバリデーションロジックを追加
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

        # データ処理
        try:
            export_folder = values[1]  # ユーザーが選択したフォルダ
            if not export_folder:
                raise Exception("エクスポート先フォルダが指定されていません")
            
            # レポート生成
            # report_dfは生成されたDataFrameと仮定
            report_df = df  # 実際のロジックで置き換え
            report_file = os.path.join(export_folder, "週次販売レポート.csv")
            report_df.to_csv(report_file, index=False)
            
            window['progressbar'].update(100)  # 100% 完了
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

    if event == "レポートダウンロード":
        # レポートをダウンロードするロジック
        try:
            # レポートが存在するか確認
            if not os.path.exists(report_file):
                raise Exception("レポートが生成されていません")
            
            # レポートのダウンロード
            # レポートのダウンロード処理
            # ダウンロードが完了したら、レポートを削除
            os.remove(report_file)
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue
    
    if event == sg.WINDOW_CLOSED:
        break
# ウィンドウを閉じる
window.close()
