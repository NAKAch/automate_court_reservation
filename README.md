# 京都市テニスコート抽選予約自動化

## 環境条件 (Requirements)

- **OS**: Windows11
- **Python**: 3.10.11

## 環境構築手順 (Setup Instructions)

1. 仮想環境を作成 (Create a virtual environment):
   ```bash
   python -m venv venv

2. 仮想環境をアクティベート (Activate the virtual environment):
   ```bash
   venv\Scripts\activate

3. 必要なパッケージをインストール (Install the required packages):

   ```bash
   pip install -r requirements.txt

4. コンパイルせずに実行する場合（Run without compile）:
   ```bash
   streamlit run main.py

5. コンパイルを実行する場合 (Compile):
   ```bash
   pyinstaller automate_court_reservation.spec