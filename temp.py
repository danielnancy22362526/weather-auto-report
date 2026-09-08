name: Weekly Weather Report

on:
  schedule:
    # 每週三 15:00 執行 (UTC 時間 07:00)
    - cron: '0 7 * * 3'
  workflow_dispatch: # 支援手動按鈕執行

# 關鍵：賦予 Actions 對存儲庫的寫入權限
permissions:
  contents: write

jobs:
  build-and-run:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4
        with:
          fetch-depth: 0 # 拉取完整 commit 歷史以利 rebase

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'
          cache: 'pip'

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install requests pandas openpyxl matplotlib pillow

      - name: Run temp.py
        run: |
          python temp.py

      - name: Commit and push generated report
        run: |
          git config --global user.name "github-actions[bot]"
          git config --global user.email "github-actions[bot]@users.noreply.github.com"
          git add *.xlsx
          # 檢查是否有檔案變動，有才進行 commit 與 push
          if git diff --staged --quiet; then
            echo "No changes to commit"
          else
            git commit -m "Auto-update weekly weather report [skip ci]"
            git pull --rebase origin main
            git push origin main
          fi

      - name: Upload Excel as Artifact
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: weather-report
          path: "*.xlsx"
