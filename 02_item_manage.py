import PySimpleGUI as sg
import csv

# CSVから在庫データを読み込む関数
def load_data():
    try:
        with open('inventory.csv', mode ='r') as file:
            csvFile = csv.reader(file)
            # CSVの各行から商品名と数量、単位を読み取り、辞書に保存
            inventory = {rows[0]: {'quantity': int(rows[1]), 'unit': rows[2]} for rows in csvFile}
        return inventory
    except FileNotFoundError:
        return {}

# 在庫データをCSVに保存する関数
def save_data(inventory):
    with open('inventory.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        for key, value in inventory.items():
            writer.writerow([key, value['quantity'], value['unit']])

# 既存のデータを読み込む
inventory = load_data()

# GUIのレイアウト設定
layout = [
    [sg.Text('商品名'), sg.Input(key='-PRODUCT-', size=(30, 1))],
    [sg.Text('数量'), sg.Input(key='-QUANTITY-', size=(10, 1))],
    [sg.Text('単位'), sg.Input(key='-UNIT-', size=(10, 1))],
    [sg.Button('追加'), sg.Button('削除')],
    [sg.Table(values=[[k, v['quantity'], v['unit']] for k, v in inventory.items()],
              headings=['商品名', '数量', '単位'],
              display_row_numbers=False,
              auto_size_columns=False,
              col_widths=[60, 20, 10],
              justification='left',
              num_rows=min(25, len(inventory)),
              key='-INVENTORY-')],
    [sg.Button('終了')]
]

# ウィンドウを作成
window = sg.Window('在庫管理', layout, size=(900, 600))

# イベントループ
while True:
    event, values = window.read()

    # 終了条件
    if event == sg.WINDOW_CLOSED or event == '終了':
        break

    # 入力された商品名と数量、単位を取得
    product = values['-PRODUCT-']
    if product:
        quantity = int(values['-QUANTITY-']) if values['-QUANTITY-'].isdigit() else 0
        unit = values['-UNIT-']

        # 「追加」ボタンが押された場合の処理
        if event == '追加':
            if product in inventory:
                inventory[product]['quantity'] += quantity
            else:
                inventory[product] = {'quantity': quantity, 'unit': unit}

        # 「削除」ボタンが押された場合の処理
        if event == '削除':
            if product in inventory:
                if inventory[product]['quantity'] >= quantity:
                    inventory[product]['quantity'] -= quantity
                else:
                    sg.popup_error('在庫が足りません')

        # CSVファイルを更新
        save_data(inventory)
    # 在庫リストを更新
    window['-INVENTORY-'].update(values=[[k, v['quantity'], v['unit']] for k, v in inventory.items()])

# ウィンドウを閉じる
window.close()
