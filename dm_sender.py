import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime
import os
import unicodedata

class DMSenderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("DM送信数入力")
        self.csv_path = None
        self.setup_ui()
        
    def setup_ui(self):
        # CSVファイル選択ボタン
        tk.Button(self.root, text="CSVファイルを選択", command=self.select_excel).pack(pady=10)
        
        # 選択されたファイルパスを表示するラベル
        self.file_label = tk.Label(self.root, text="ファイルが選択されていません", wraplength=300)
        self.file_label.pack(pady=5)
        
        # 入力フィールドの作成
        tk.Label(self.root, text="送信するDM数を入力してください:").pack(pady=10)
        self.count_entry = tk.Entry(self.root)
        self.count_entry.pack(pady=5)
        
        # 実行ボタンの作成
        self.execute_button = tk.Button(self.root, text="実行", command=self.process_dm, state='disabled')
        self.execute_button.pack(pady=10)
        
    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSVファイル", "*.csv")]
        )
        if file_path:
            self.csv_path = file_path
            self.file_label.config(text=f"選択されたファイル:\n{os.path.basename(file_path)}")
            self.execute_button.config(state='normal')  # ファイル選択後に実行ボタンを有効化
    
    def normalize_column_names(self, columns):
        """
        列名の全角数字を半角数字に変換し、先頭と末尾のスペースを削除するヘルパー関数
        """
        return [''.join([unicodedata.normalize('NFKC', c) for c in col]).strip() for col in columns]
    
    def process_dm(self):
        if not self.csv_path:
            messagebox.showerror("エラー", "CSVファイルを選択してください")
            return
            
        try:
            print("1. 入力値の確認")
            count = int(self.count_entry.get())
            if count <= 0:
                messagebox.showerror("エラー", "正の整数を入力してください")
                return
                
            print("2. CSVファイル読み込み開始")
            # 複数のエンコーディングを試す
            encodings = ['cp932', 'shift_jisx0213', 'shift-jis', 'utf-8', 'utf-8-sig']
            for encoding in encodings:
                try:
                    print(f"エンコーディング {encoding} で試行中...")
                    df = pd.read_csv(self.csv_path, encoding=encoding)
                    print(f"成功: {encoding} で読み込みました")
                    break
                except Exception as e:
                    print(f"失敗: {encoding} - {str(e)}")
                    if encoding == encodings[-1]:
                        raise Exception("すべてのエンコーディングで読み込みに失敗しました")
                    continue
                
            print("3. CSVファイル読み込み完了")
            
            # 列名の正規化
            df.columns = self.normalize_column_names(df.columns)
            print("4. 列名の正規化完了")
            
            # デバッグ: 正規化後の列名を表示
            print("現在の DataFrame の列名:")
            print(df.columns.tolist())
            
            print("5. DM停止フラグ確認")
            if 'DM停止' not in df.columns:
                raise KeyError("'DM停止' 列が存在しません")
            df = df[df['DM停止'].fillna('').str.strip() != '停止']
            print(f"DM停止フラグ確認完了。対象行数: {len(df)}")
            
            print("6. 発送回数のソート開始")
            # 発送回数の列名を取得（実際の列名に合わせて変更してください）
            send_count_col = '発送回数'  # 実際の列名に合わせて変更
            if send_count_col not in df.columns:
                raise KeyError(f"'{send_count_col}' 列が存在しません")
            df = df.sort_values(by=send_count_col, ascending=False)
            print("7. ソート完了")
            
            print("8. 送信リスト作成開始")
            send_list = df.head(count).copy()
            print(f"送信リストの作成完了。送信件数: {len(send_list)}")
            
            # 最新の発送回数を取得して新しい発送日列名を作成
            latest_send_count = send_list[send_count_col].max()
            new_send_count = latest_send_count + 1
            latest_send_date = f'第{new_send_count}回発送日'
            print(f"新しい発送日列名: {latest_send_date}")
            
            # 新しい発送日列を追加
            df = self._add_new_send_date_column(df, f'第{latest_send_count}回発送日', latest_send_date)
            print(f"列追加後の DataFrame の列名:")
            print(df.columns.tolist())
            
            print("9. 発送日更新処理開始")
            df = self._update_send_dates(df, send_list.index, latest_send_date)
            print("発送日更新処理完了")
            
            # デバッグ: 更新後の DataFrame の確認
            print("更新後の DataFrame の一部:")
            print(df.head())
            
            print("10. ファイル出力準備")
            # 送信リストの作成（.copyを使用して明示的にコピーを作成）
            required_columns = {'郵便番号', '宛先住所1', '送付先名', '管理者氏名'}
            if required_columns.issubset(send_list.columns):
                output_list = send_list[['郵便番号', '宛先住所1', '送付先名', '管理者氏名']].copy()
                output_list.loc[:, '敬称'] = '様'
            else:
                missing_cols = required_columns - set(send_list.columns)
                raise KeyError(f"必要な列が不足しています: {missing_cols}")
            
            # デバッグ: 送信リストの確認
            print("送信リストの作成完了。送信リストの一部:")
            print(output_list.head())
            
            # ファイル名の生成
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_dir = os.path.dirname(self.csv_path)
            
            # 送信リストの保存
            list_filename = os.path.join(excel_dir, f'送信リスト_{timestamp}.csv')
            output_list.to_csv(list_filename, index=False, encoding='cp932')
            print(f"送信リストを保存しました: {list_filename}")
            
            # 更新したマスターの保存
            master_filename = os.path.join(excel_dir, f'マスター_{timestamp}.csv')
            df.to_csv(master_filename, index=False, encoding='cp932')
            print(f"更新マスターを保存しました: {master_filename}")
            print(f"作成したファイル:")
            print(f"- 送信リスト: {os.path.basename(list_filename)}")
            print(f"- 更新マスター: {os.path.basename(master_filename)}")
            messagebox.showinfo("完了", f"処理が完了しました\n作成ファイル:\n- {os.path.basename(list_filename)}\n- {os.path.basename(master_filename)}")
            
            self.root.destroy()
            
        except ValueError as ve:
            print(f"ValueError発生: {str(ve)}")
            messagebox.showerror("エラー", "有効な数値を入力してください")
        except KeyError as ke:
            print(f"KeyError発生: {str(ke)}")
            messagebox.showerror("エラー", f"必要な列が見つかりません: {str(ke)}")
        except Exception as e:
            print(f"エラー発生箇所の詳細: {str(e)}")
            print(f"エラーの種類: {type(e)}")
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")
    
    def _add_new_send_date_column(self, df, prev_col, new_col):
        try:
            if prev_col in df.columns:
                prev_idx = df.columns.get_loc(prev_col)
                df.insert(prev_idx + 1, new_col, None)
                print(f"新しい列 '{new_col}' を '{prev_col}' の後ろに追加しました")
            else:
                # 前の列が存在しない場合は最後に追加
                df[new_col] = None
                print(f"新しい列 '{new_col}' を最後に追加しました")
            return df
        except Exception as e:
            print(f"新しい発送日列の追加でエラー発生: {str(e)}")
            raise
        
    def _update_send_dates(self, df, target_indices, date_column):
        try:
            # 発送日の列に今日の日付を挿入
            today = datetime.now().strftime('%Y年%m月%d日')
            print(f"今日の日付: {today}")

            # 対象の行のみ日付を更新
            print(f"対象インデックス数: {len(target_indices)}")
            df.loc[target_indices, date_column] = today
            print(f"'{date_column}' 列の更新が完了しました")
            
            # デバッグ: 更新された行の確認
            updated_rows = df.loc[target_indices, date_column]
            print(f"更新された '{date_column}' 列の一部:")
            print(updated_rows.head())
            
            return df  # 更新したDataFrameを返す
                
        except Exception as e:
            print(f"発送日更新処理でエラー発生: {str(e)}")
            raise

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = DMSenderApp()
    app.run() 