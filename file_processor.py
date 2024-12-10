import pandas as pd
import chardet
import io
from datetime import datetime

class FileProcessor:
    @staticmethod
    def detect_encoding(file_bytes):
        result = chardet.detect(file_bytes)
        return result['encoding']

    @staticmethod
    def read_excel(file):
        try:
            df = pd.read_excel(file, engine='openpyxl')
            return df
        except Exception as e:
            print(f"Excelファイルの読み込みエラー: {str(e)}")
            raise

    @staticmethod
    def read_csv(file, file_type='default'):
        try:
            file_bytes = file.getvalue()
            encoding = FileProcessor.detect_encoding(file_bytes)
            
            if file_type == 'inventory':
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, skiprows=7)
            else:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
            
            return df
        except Exception as e:
            print(f"CSVファイルの読み込みエラー: {str(e)}")
            raise

    @staticmethod
    def process_data(purchase_history_df, inventory_df, yj_code_df):
        try:
            print("\n=== エラーデバッグ情報 ===")
            print("処理開始時のデータ型:")
            print("purchase_history_df type:", type(purchase_history_df))
            print("inventory_df type:", type(inventory_df))
            print("yj_code_df type:", type(yj_code_df))

            print("\n=== データ処理開始 ===")
            print("1. データフレームの初期状態:")
            print(f"Inventory shape: {inventory_df.shape}")
            print(f"Purchase history shape: {purchase_history_df.shape}")
            print(f"YJ code shape: {yj_code_df.shape}")
            
            print("\n2. 各データフレームのカラム:")
            print("Inventory columns:", inventory_df.columns.tolist())
            print("Purchase history columns:", purchase_history_df.columns.tolist())
            print("YJ code columns:", yj_code_df.columns.tolist())
            
            print("\n3. データサンプル:")
            print("\nInventory データ:")
            print(inventory_df.head())
            print("\nPurchase history データ:")
            print(purchase_history_df.head())
            print("\nYJ code データ:")
            print(yj_code_df.head())
            
            # データの前処理
            inventory_df = inventory_df.fillna('')
            purchase_history_df = purchase_history_df.fillna('')
            yj_code_df = yj_code_df.fillna('')
            
            # データ型の確認と変換
            for col in inventory_df.columns:
                if inventory_df[col].dtype == 'object':
                    inventory_df[col] = inventory_df[col].astype(str)
                print(f"Column: {col}")
                print(f"Type: {inventory_df[col].dtype}")
                print(f"First 5 values: {inventory_df[col].head().tolist()}")
        
            # 空の薬品名を持つ行を削除
            inventory_df = inventory_df[inventory_df['薬品名'].notna() & (inventory_df['薬品名'] != '')].copy()
            print("空の薬品名を削除後の inventory shape:", inventory_df.shape)
            
            # データフレームの型チェックとNaN値の処理
            for df_name, df in {"inventory": inventory_df, "purchase_history": purchase_history_df, "yj_code": yj_code_df}.items():
                if not isinstance(df, pd.DataFrame):
                    print(f"Warning: {df_name} is not a DataFrame")
                    print(f"Type of {df_name}: {type(df)}")
                    continue
                print(f"\nProcessing {df_name} DataFrame:")
                print(f"Columns: {df.columns.tolist()}")
                for col in df.columns:
                    print(f"Processing column: {col}")
                    df[col] = df[col].fillna('').astype(str)
            
            # 在庫金額CSVから薬品名とＹＪコードのマッピングを作成
            if not yj_code_df.empty:
                yj_mapping = dict(zip(yj_code_df['薬品名'], zip(yj_code_df['ＹＪコード'], yj_code_df['単位'])))
            else:
                print("Warning: YJ code DataFrame is empty")
                yj_mapping = {}
            
            # 不良在庫データに対してＹＪコードと単位を設定
            inventory_df['ＹＪコード'] = inventory_df['薬品名'].map(lambda x: yj_mapping.get(x, (None, None))[0])
            inventory_df['単位'] = inventory_df['薬品名'].map(lambda x: yj_mapping.get(x, (None, None))[1])
            
            # マージ前の状態を確認
            print("\nマージ前のデータ確認:")
            print("Inventory columns:", inventory_df.columns.tolist())
            print("Purchase history columns:", purchase_history_df.columns.tolist())
            
            # 必要なカラムが存在することを確認
            required_columns = ['厚労省CD', '法人名', '院所名', '品名・規格', '新薬品ｺｰﾄﾞ']
            missing_columns = [col for col in required_columns if col not in purchase_history_df.columns]
            
            if missing_columns:
                print(f"Warning: Missing columns in purchase_history_df: {missing_columns}")
                # 不足しているカラムを空の文字列で追加
                for col in missing_columns:
                    purchase_history_df[col] = ''
            
            # ＹＪコードと厚労省CDで紐付け
            try:
                print("\n4. マージ処理:")
                print("マージ前の状態:")
                print("- Inventory df keys:", inventory_df['ＹＪコード'].head())
                print("- Purchase history df keys:", purchase_history_df['厚労省CD'].head())
                print("\nデータ型確認:")
                print("- Inventory df['ＹＪコード'] type:", type(inventory_df['ＹＪコード']))
                print("- Purchase history df['厚労省CD'] type:", type(purchase_history_df['厚労省CD']))
                print("\nデータサンプル:")
                print("- Inventory df sample:\n", inventory_df[['ＹＪコード', '薬品名']].head())
                print("- Purchase history df sample:\n", purchase_history_df[['厚労省CD', '品名・規格']].head())
                
                merged_df = pd.merge(
                    inventory_df,
                    purchase_history_df[required_columns],
                    left_on='ＹＪコード',
                    right_on='厚労省CD',
                    how='left'
                )
                print("\nマージ後の状態:")
                print("- 形状:", merged_df.shape)
                print("- カラム:", merged_df.columns.tolist())
                print("- データサンプル:")
                print(merged_df.head())
            except Exception as e:
                print(f"\nマージ中にエラーが発生:")
                print(f"エラー詳細: {str(e)}")
                print(f"エラータイプ: {type(e)}")
                raise
            
            # 必要なカラムのみを選択
            try:
                result_df = merged_df[[
                    '品名・規格',
                    '在庫量',
                    '単位',
                    '新薬品ｺｰﾄﾞ',
                    '使用期限',
                    'ロット番号',
                    '法人名',
                    '院所名'
                ]].copy()
                
                # 空の値を空文字列に変換
                result_df = result_df.fillna('')
                
                # 空の品名・規格を持つ行を削除
                result_df = result_df[result_df['品名・規格'].notna() & (result_df['品名・規格'] != '')].copy()
                
                # 院所名でソート
                result_df = result_df.sort_values(['法人名', '院所名'])
                
                print("\n最終的なデータフレームの状態:")
                print("Columns:", result_df.columns.tolist())
                print("データ型:")
                print(result_df.dtypes)
                print("\nサンプルデータ:")
                print(result_df.head())
                
                return result_df
            
            except Exception as e:
                print(f"結果データフレーム作成中にエラーが発生: {str(e)}")
                raise
            
        except Exception as e:
            print(f"データ処理中にエラーが発生: {str(e)}")
            raise

    @staticmethod
    def generate_excel(df):
        print("\n=== Excel生成開始 ===")
        print("1. 入力データフレーム情報:")
        print(f"行数: {df.shape[0]}")
        print(f"列名: {df.columns.tolist()}")
        print("\n2. データサンプル:")
        print(df.head())
        
        excel_buffer = io.BytesIO()
        
        # シート名として無効な文字を置換する関数
        def clean_sheet_name(name):
            print(f"\n3. シート名クリーニング:")
            print(f"元の名前: '{name}'")
            print(f"型: {type(name)}")
            
            if not isinstance(name, str) or not name.strip():
                print("警告: 無効なシート名、'Unknown'を使用")
                return 'Unknown'
            
            # 特殊文字を置換
            invalid_chars = ['/', '\\', '?', '*', ':', '[', ']']
            cleaned_name = ''.join('_' if c in invalid_chars else c for c in name)
            # 最大31文字に制限（Excelの制限）
            final_name = cleaned_name[:31].strip()
            print(f"クリーニング後の名前: '{final_name}'")
            return final_name

        try:
            # ExcelWriterを使用して、院所名ごとにシートを作成
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # 院所名ごとにシートを作成（空の値を除外）
                for name in df['院所名'].unique():
                    if pd.notna(name) and str(name).strip():  # 空の値をスキップ
                        sheet_name = clean_sheet_name(str(name))
                        # 該当する院所のデータを抽出
                        sheet_df = df[df['院所名'] == name].copy()
                        
                        if not sheet_df.empty:
                            try:
                                # 法人名と院所名を取得（ヘッダー用）
                                houjin_name = str(sheet_df['法人名'].iloc[0]).strip() if not pd.isna(sheet_df['法人名'].iloc[0]) else ''
                                insho_name = str(sheet_df['院所名'].iloc[0]).strip() if not pd.isna(sheet_df['院所名'].iloc[0]) else ''
                                
                                # デバッグ情報の詳細出力
                                print("\n=== ヘッダー生成処理 ===")
                                print(f"1. シート名: {sheet_name}")
                                print("2. データフレーム情報:")
                                print(f"  - 行数: {sheet_df.shape[0]}")
                                print(f"  - カラム: {sheet_df.columns.tolist()}")
                                print("\n3. 値の確認:")
                                print(f"  法人名（生データ）: {sheet_df['法人名'].iloc[0]}")
                                print(f"  院所名（生データ）: {sheet_df['院所名'].iloc[0]}")
                                print(f"  法人名（変換後）: '{houjin_name}'")
                                print(f"  院所名（変換後）: '{insho_name}'")
                                print(f"  結合後のテキスト: '{f'{houjin_name} {insho_name}'.strip()}'")
                                
                                # ヘッダーデータの作成
                                header_rows = [
                                    ['不良在庫引き取り依頼', None, None],
                                    [None, None, None],
                                    [f"{houjin_name} {insho_name}".strip(), None, '御中'],
                                    [None, None, None],
                                    ['下記の不良在庫につきまして、引き取りのご検討を賜れますと幸いです。どうぞよろしくお願いいたします。', None, None],
                                    [None, None, None]
                                ]
                                header_data = pd.DataFrame(header_rows)
                                
                                print("Debug - ヘッダーデータ:")
                                print(header_data)
                                
                                print("Debug - ヘッダーデータ:")
                                print(header_data)
                                
                                # ヘッダーとデータを書き込み
                                header_data.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                                sheet_df.to_excel(writer, sheet_name=sheet_name, startrow=6, index=False)
                                
                                # シートを取得してフォーマットを設定
                                worksheet = writer.sheets[sheet_name]
                                
                                # 列幅の設定
                                worksheet.column_dimensions['A'].width = 35  # 品名・規格
                                worksheet.column_dimensions['B'].width = 15  # 在庫量
                                worksheet.column_dimensions['C'].width = 10  # 単位
                                worksheet.column_dimensions['D'].width = 15  # 新薬品コード
                                worksheet.column_dimensions['E'].width = 15  # 使用期限
                                worksheet.column_dimensions['F'].width = 15  # ロット番号
                                worksheet.column_dimensions['G'].width = 20  # 引取り可能数
                                
                                # 行の高さを設定（30ピクセル）
                                for row in range(1, worksheet.max_row + 1):
                                    worksheet.row_dimensions[row].height = 30
                                
                                # デフォルトのフォントサイズを14に設定
                                for row in worksheet.rows:
                                    for cell in row:
                                        cell.font = cell.font.copy(size=14)
                                
                                # タイトルのフォント設定（サイズ16）
                                cell_a1 = worksheet['A1']
                                cell_a1.font = cell_a1.font.copy(size=16)
                                
                                # 法人名、院所名、御中のフォント設定（サイズ14、太字）
                                cell_a3 = worksheet['A3']  # 法人名
                                cell_b3 = worksheet['B3']  # 院所名
                                cell_c3 = worksheet['C3']  # 御中
                                font_style = cell_a3.font.copy(size=14, bold=True)
                                cell_a3.font = font_style
                                cell_b3.font = font_style
                                cell_c3.font = font_style
                                
                            except Exception as e:
                                print(f"シート '{sheet_name}' の処理中にエラーが発生: {str(e)}")
                                continue
                            
        except Exception as e:
            print(f"Excel生成中にエラーが発生: {str(e)}")
            return None

        excel_buffer.seek(0)
        return excel_buffer
