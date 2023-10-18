import PySimpleGUI as sg
import pandas as pd
import os
from datetime import datetime, timedelta
import shutil

# データ処理関数
def read_and_filter_data(file_path):
    df = pd.read_csv(file_path)
    df['日付'] = pd.to_datetime(df['日付'])
    one_week_ago = pd.Timestamp(datetime.now() - timedelta(days=7))
    return df[df['日付'] > one_week_ago]

# 集計関数
def aggregate_data(df):
    # 商品別集計
    product_summary = df.groupby('商品ID').agg({
        '販売数': 'sum',
        '単価': 'mean'
    }).reset_index()

    # 日別集計
    daily_summary = df.groupby('日付').agg({
        '販売数': 'sum',
        '単価': 'mean'
    }).reset_index()

    return product_summary, daily_summary

# GUIレイアウトの設定
layout = [
    [sg.Text("売上データをアップロード:"), sg.Input(), sg.FileBrowse("ファイル選択")],
    [sg.Text("エクスポート先フォルダ:"), sg.Input(), sg.FolderBrowse("フォルダ選択")],
    [sg.Button("レポート生成"), sg.Button("集計結果確認")],
    [sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')],
    [sg.Button("レポートダウンロード")],
    [sg.Text("", key='error_message')],
]
# ウィンドウを作成
window = sg.Window("週次販売レポート生成器", layout, size=(900, 600))

while True:
    event, values = window.read()

    if event == "レポート生成":
        # データバリデーションとエラーハンドリング
        try:
            df_filtered = read_and_filter_data(values[0])
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

        # データ処理
        try:
            export_folder = values[1]
            if not export_folder:
                raise Exception("エクスポート先フォルダが指定されていません")

            product_summary, daily_summary = aggregate_data(df_filtered)

            report_file = os.path.join(export_folder, "週次販売レポート.xlsx")
            with pd.ExcelWriter(report_file) as writer:
                product_summary.to_excel(writer, sheet_name='商品別集計', index=False)
                daily_summary.to_excel(writer, sheet_name='日別集計', index=False)

            window['progressbar'].update(100)
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

    elif event == "集計結果確認":
        try:
            # データバリデーションとエラーハンドリング
            df_filtered = read_and_filter_data(values[0])
            # データ処理
            product_summary, daily_summary = aggregate_data(df_filtered)
            # 集計結果の表示
            layout_summary = [
                [sg.Text("商品別集計")],
                [sg.Table(values=product_summary.values.tolist(), headings=list(product_summary.columns), display_row_numbers=False)],
                [sg.Text("日別集計")],
                [sg.Table(values=daily_summary.values.tolist(), headings=list(daily_summary.columns), display_row_numbers=False)],
                [sg.Button("閉じる")]
            ]

            window_summary = sg.Window("集計結果", layout_summary)

            # 集計結果のウィンドウを閉じるまでループ
            while True:
                event_summary, values_summary = window_summary.read()
                if event_summary in ("閉じる", None):
                    window_summary.close()
                    break

        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

    if event == "レポートダウンロード":
        # レポートをダウンロードするロジック
        try:
            if not os.path.exists(report_file):
                raise Exception("レポートが生成されていません")
            
            # ユーザーが保存先を選ぶダイアログを表示
            save_path = sg.popup_get_file('保存先を選んでください', save_as=True)
            if not save_path:
                raise Exception("保存先が指定されていません")

            # ファイルを保存先にコピー
            shutil.copy(report_file, save_path)

            # オリジナルのレポートファイルを削除
            os.remove(report_file)
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue
    
    if event == sg.WINDOW_CLOSED:
        break
# ウィンドウを閉じる
window.close()
