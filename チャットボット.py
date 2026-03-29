import streamlit as st
import pandas as pd
import os
import re
import unicodedata # 全角半角変換用
from collections import defaultdict

# --- 設定と定数 ---
# Excelファイルのパス。
# ここにあなたの環境での「契約情報まとめ_単一シート.xlsx」の絶対パスを指定してください。
# 例: r"C:\Users\ts-kanta.yamamura\Documents\03_契約書回り\11_save\契約情報まとめ_単一シート.xlsx"
EXCEL_FILE_PATH = r"C:\Users\ts-kanta.yamamura\Documents\03_契約書回り\11_save\契約情報まとめ_単一シート.xlsx"

# 検索対象とする列
SEARCH_COLUMNS = ['項目', '期間/備考']

# 表記揺れマッピング (ユーザー入力 -> 内部検索キーワード)
# Excel内の表記に合わせて調整してください。
# 例: Excelに「ホークアイ」と日本語で書かれている場合、ユーザーが「ホークアイ」と入力したら「ホークアイ」で検索すべき。
#     Excelに「Hawk-eye」と英語で書かれている場合、ユーザーが「ホークアイ」と入力したら「Hawk-eye」で検索すべき。
#     ここでは、ユーザーが日本語で入力しても、Excel内の表記（日本語/英語）に合わせて検索できるよう、
#     より広い意味での「キーワード」としてマッピングを定義しています。
KEYWORD_MAPPING = {
    # PITCHBASE関連
    'ピッチベース': 'pitchbase',
    'pitchbase': 'pitchbase',
    'pichbase': 'pitchbase', # よくあるtypoも含む
    
    # ホークアイ関連
    'ホークアイ': 'ホークアイ', # Excelに日本語で「ホークアイ」とあるため
    'hawk-eye': 'hawk-eye',   # Excelに英語で「Hawk-eye」とある場合も想定
    'ホークアイトラッキング': 'ホークアイ',
    'tracking system': 'tracking system',
    'トラッキングシステム': 'tracking system',

    # Trajekt Arc関連
    'トラジェクト': 'trajekt',
    'trajekt': 'trajekt',
    'ロボット': 'ロボット',
    'robot': 'robot',

    # 費用・金額関連
    '費用': '金額', 
    '料金': '金額',
    'コスト': '金額',
    '金額': '金額',
    
    # 契約期間関連
    '契約期間': '契約期間',
    '期間': '期間',
    '有効期間': '有効期間',

    # 年度関連 (Excel内の表記に合わせる)
    '2021年': '2021年',
    '2022年': '2022年',
    '2023年': '2023年',
    '2024年': '2024年',
    '2025年': '2025年',
    '2026年': '2026年',
    '2021': '2021年', # 数字のみでもヒットするように
    '2022': '2022年',
    '2023': '2023年',
    '2024': '2024年',
    '2025': '2025年',
    '2026': '2026年',

    # その他
    'ライセンス': 'ライセンス',
    'サービス': 'サービス',
    'npb': 'npb',
    'dmp': 'dmp',
    'cms': 'cms',
    'id発行': 'id発行',
    'rapsodo': 'rapsodo',
    'blast': 'blast',
    '運用保守': '運用保守',
    'ap設置': 'ap設置',
    'シーズン': 'シーズン',
    '購入': '購入',
    'デポジット': 'デポジット',
    '残金': '残金',
}

# カテゴリ検索用のキーワード
CATEGORY_KEYWORDS = ['カテゴリ', 'カテゴリー', '分類', '一覧', '種類']

# --- ContractInfoBot クラスの定義 ---
class ContractInfoBot:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.load_data()

    def load_data(self):
        """Excelファイルを読み込み、前処理を行う"""
        try:
            self.df = pd.read_excel(self.file_path)
            st.success(f"Excelファイル '{os.path.basename(self.file_path)}' を読み込みました。")
            
            # 検索効率のために、指定された列を結合した列を作成
            # NaN値を空文字列に変換し、文字列として結合
            self.df['search_text'] = self.df[SEARCH_COLUMNS].astype(str).agg(' '.join, axis=1)
            # search_textを小文字に変換し、正規化も行う
            self.df['search_text_normalized'] = self.df['search_text'].apply(self._normalize_text)
            
            # 項目も小文字にしておく (カテゴリ判定用)
            self.df['項目_lower'] = self.df['項目'].astype(str).str.lower()

            # デバッグ用に最初の数行をサイドバーに表示
            # st.sidebar.write("--- Loaded Data Head ---")
            # st.sidebar.dataframe(self.df.head())
            # st.sidebar.write("------------------------")

        except FileNotFoundError:
            st.error(f"エラー: Excelファイルが見つかりません。パスを確認してください: {self.file_path}")
            st.info("Excelファイルが正しい絶対パスで指定されているか確認してください。")
        except Exception as e:
            st.error(f"エラー: Excelファイルの読み込み中に問題が発生しました: {e}")
            st.info(f"詳細: {e}")

    def _normalize_text(self, text):
        """テキストを正規化（全角半角、大文字小文字、記号除去）"""
        # NaNが来る可能性があるので文字列に変換
        text = str(text) 
        # 全角英数字を半角に、全角カタカナを半角に
        text = unicodedata.normalize('NFKC', text)
        text = text.lower()
        # 不要な記号を除去（句読点、記号など）
        text = re.sub(r'[!-/:-@\[-`{-~、。！？・]', '', text)
        # 連続するスペースを一つにする
        text = re.sub(r'\s+', ' ', text).strip()
        return text

    def _extract_keywords(self, query):
        """ユーザー入力から検索キーワードを抽出し、マッピングを適用する"""
        normalized_query = self._normalize_text(query)
        
        # 助詞や一般的なストップワードを除去（簡易版）
        # より高度な形態素解析を行う場合はMeCabなどを使用
        stopwords = ['の', 'を', 'は', 'が', 'に', 'で', 'と', 'から', 'まで', 'へ', 'より', 'や', 'も', 'か', 
                     'ください', '教えて', 'いくら', 'なんですか', 'ですか', 'について', 'に関する', 'について教えて', 'について知りたい']
        
        # ストップワードを除去した後にキーワードを抽出
        processed_keywords_str = normalized_query
        for sw in stopwords:
            processed_keywords_str = processed_keywords_str.replace(sw, ' ')
        
        # スペースで区切ってキーワードを抽出し、マッピングを適用
        extracted_raw_keywords = [k for k in processed_keywords_str.split() if k]
        
        mapped_keywords = []
        for raw_k in extracted_raw_keywords:
            # マッピング辞書に存在すれば変換、なければそのまま
            mapped_keywords.append(KEYWORD_MAPPING.get(raw_k, raw_k))
            
        # 重複キーワードを除去
        final_keywords = list(dict.fromkeys(mapped_keywords))

        # デバッグ出力
        st.sidebar.write(f"--- Debug Info ---")
        st.sidebar.write(f"Original Query: {query}")
        st.sidebar.write(f"Normalized Query: {normalized_query}")
        st.sidebar.write(f"Extracted Raw Keywords: {extracted_raw_keywords}")
        st.sidebar.write(f"Mapped Keywords: {final_keywords}")
        st.sidebar.write(f"------------------")

        return final_keywords

    def search_info(self, query):
        """
        ユーザーのクエリに基づいて契約情報を検索し、結果を返します。
        """
        if self.df is None:
            return "データが読み込まれていないため、検索できません。"

        # クエリから検索キーワードを抽出
        keywords = self._extract_keywords(query)

        if not keywords: # キーワードが抽出できなかった場合
            return f"質問の意図を理解できませんでした。もう少し具体的な単語で質問してください。"

        # 全てのキーワードを含む行を検索する（AND検索）
        # search_text_normalized 列に対して検索
        filtered_df = self.df.copy()
        for k in keywords:
            # 各キーワードが search_text_normalized に含まれるか
            filtered_df = filtered_df[
                filtered_df['search_text_normalized'].str.contains(k, na=False)
            ]
        
        results = filtered_df

        # ヘッダー行を除外
        results = results[~results['項目'].astype(str).str.startswith('■')]

        if results.empty:
            # 検索結果が見つからなかった場合のヒント
            hint_message = ""
            if len(keywords) > 1:
                hint_message = "キーワードが多すぎるか、組み合わせが一致しませんでした。キーワードを減らして再試行するか、別の表現をお試しください。"
            else:
                hint_message = "関連情報が見つかりませんでした。別のキーワードや表現をお試しください。"
            return f"'{query}' に関連する情報は見つかりませんでした。\n\n{hint_message}"

        response_parts = []
        # 検索結果が多すぎる場合のヒント
        if len(results) > 10: # 例えば10件以上は多すぎると判断
            response_parts.append(f"'{query}' に関連する情報が多数見つかりました（**{len(results)}件**）。\n\nもう少し質問を絞り込んでいただけますか？（例: 'ホークアイ 2023年'）")
            response_parts.append("---")
            # 上位5件のみ表示
            results = results.head(5) 

        response_parts.append(f"'{query}' に関連する情報が見つかりました:\n")
        
        for index, row in results.iterrows():
            response_parts.append(f"  **項目**: {row['項目']}")
            
            # 金額表示を整形 (通貨と税に関する情報を明示)
            amount = str(row['金額'])
            
            # クエリに「税別」「税込」「USD」「ドル」などが含まれるかチェックし、重複表示を避ける
            query_lower = query.lower()
            if '税別' in amount and '税別' not in query_lower:
                amount += " (税別)"
            elif '税込' in amount and '税込' not in query_lower:
                amount += " (税込)"
            elif 'usd' in amount and ('usd' not in query_lower and 'ドル' not in query_lower):
                amount += " (米ドル)"
            
            response_parts.append(f"  **金額**: {amount}")
            response_parts.append(f"  **期間/備考**: {row['期間/備考']}")
            response_parts.append("  ---")
        return "\n".join(response_parts)

    def get_all_categories(self):
        """
        ファイル内のすべての契約カテゴリ（見出し）を返します。
        """
        if self.df is None:
            return "データが読み込まれていません。"
        
        # 項目が '■' で始まる行をカテゴリとして抽出
        categories = self.df[self.df['項目'].astype(str).str.startswith('■')]['項目'].tolist()
        if categories:
            return "**以下のカテゴリがあります:**\n\n" + "\n".join(f"- {cat}" for cat in categories)
        else:
            return "カテゴリが見つかりませんでした。"

# --- Streamlit アプリケーションの構築 ---

# Streamlit のページ設定
st.set_page_config(page_title="契約情報チャットボット", layout="centered", initial_sidebar_state="expanded")
st.title("📄 契約情報チャットボット")
st.markdown("---")
st.write("契約情報に関する質問を入力してください。（例: 'ホークアイ 費用', 'PITCHBASE 2023年', 'カテゴリ一覧'）")
st.info("左側のサイドバーにデバッグ情報が表示されます。")


# ボットの初期化（一度だけ実行されるように st.session_state を使用）
if 'bot' not in st.session_state:
    st.session_state.bot = ContractInfoBot(EXCEL_FILE_PATH)
    st.session_state.messages = [] # チャット履歴を保持

# Excelファイルが読み込まれていない場合は処理を中断
if st.session_state.bot.df is None:
    st.warning("Excelファイルの読み込みに失敗しているため、チャットボットは利用できません。")
    st.stop() # ここでアプリの実行を停止

# チャット履歴の表示
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"]) # Markdown表示に対応

# ユーザーからの入力
if prompt := st.chat_input("質問を入力してください..."):
    # ユーザーのメッセージを履歴に追加
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt) # Markdown表示に対応

    # ボットの応答を生成
    with st.chat_message("assistant"):
        response = ""
        # カテゴリ検索のトリガー
        if any(cat_word in prompt.lower() for cat_word in CATEGORY_KEYWORDS):
            response = st.session_state.bot.get_all_categories()
        else:
            response = st.session_state.bot.search_info(prompt)
        
        st.markdown(response) # Markdown表示に対応
        # ボットのメッセージを履歴に追加
        st.session_state.messages.append({"role": "assistant", "content": response})

