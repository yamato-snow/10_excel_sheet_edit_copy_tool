import PySimpleGUI as sg
import pandas as pd
import os
from datetime import datetime, timedelta

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
            df = pd.read_csv(values[0])  # ここでバリデーションロジックを追加
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

        # データ処理
        try:
            export_folder = values[1]  # ユーザーが選択したフォルダ
            if not export_folder:
                raise Exception("エクスポート先フォルダが指定されていません")
            
            # 過去1週間のデータにフィルタリング
            one_week_ago = pd.Timestamp(datetime.now() - timedelta(days=7))
            df['日付'] = pd.to_datetime(df['日付'])
            df_filtered = df[df['日付'] > one_week_ago]

            # 商品別集計
            product_summary = df_filtered.groupby('商品ID').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # 日別集計
            daily_summary = df_filtered.groupby('日付').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # カテゴリ別集計
            df_filtered['カテゴリ'] = df_filtered['商品ID'].apply(lambda x: '食品' if x <= 1005 else '家電')
            category_summary = df_filtered.groupby('カテゴリ').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # レポート生成
            report_file = os.path.join(export_folder, "週次販売レポート.xlsx")
            with pd.ExcelWriter(report_file) as writer:
                product_summary.to_excel(writer, sheet_name='商品別集計', index=False)
                daily_summary.to_excel(writer, sheet_name='日別集計', index=False)
                category_summary.to_excel(writer, sheet_name='カテゴリ別集計', index=False)
            
            window['progressbar'].update(100)  # 100% 完了
        except Exception as e:
            window['error_message'].update(f"エラー: {e}")
            continue

    elif event == "集計結果確認":
        try:
            # データ読み込み
            df = pd.read_csv(values[0])
            df['日付'] = pd.to_datetime(df['日付'])  # 日付をdatetime型に変換

            # 過去1週間のデータにフィルタリング
            one_week_ago = pd.Timestamp(datetime.now() - timedelta(days=7))
            df_filtered = df[df['日付'] > one_week_ago]

            # 商品別集計
            product_summary = df_filtered.groupby('商品ID').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # 日別集計
            daily_summary = df_filtered.groupby('日付').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # 商品カテゴリ別集計
            df_filtered['カテゴリ'] = df_filtered['商品ID'].apply(lambda x: '食品' if x <= 1005 else '家電')
            category_summary = df_filtered.groupby('カテゴリ').agg({
                '販売数': 'sum',
                '単価': 'mean'
            }).reset_index()

            # 新しいウィンドウで集計結果を表示
            layout_summary = [
                [sg.Text("商品別集計")],
                [sg.Table(values=product_summary.values.tolist(), headings=list(product_summary.columns), display_row_numbers=False)],
                [sg.Text("日別集計")],
                [sg.Table(values=daily_summary.values.tolist(), headings=list(daily_summary.columns), display_row_numbers=False)],
                [sg.Text("カテゴリ別集計")],
                [sg.Table(values=category_summary.values.tolist(), headings=list(category_summary.columns), display_row_numbers=False)],
                [sg.Button("出力確認")]
            ]
            window_summary = sg.Window("集計結果", layout_summary)

            while True:
                event_summary, values_summary = window_summary.read()

                if event_summary == "出力確認":
                    # レポート出力処理
                    report_file = os.path.join(values[1], "週次販売レポート.xlsx")
                    with pd.ExcelWriter(report_file) as writer:
                        product_summary.to_excel(writer, sheet_name='商品別集計', index=False)
                        daily_summary.to_excel(writer, sheet_name='日別集計', index=False)
                        category_summary.to_excel(writer, sheet_name='カテゴリ別集計', index=False)
                    window_summary.close()
                    break
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
