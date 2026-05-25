#!/bin/bash
cd /Users/kelly/Desktop/CB價格更新
/Users/kelly/opt/anaconda3/envs/py311/bin/python fetch_prices.py
/Users/kelly/opt/anaconda3/envs/py311/bin/python cb_inst_trading.py
git add cb_tracker.html prices.json cb_tracker_latest.json cb_list.xlsx cb_inst_latest.json
git commit -m "自動更新 $(date '+%Y-%m-%d %H:%M')"
git push origin main --force
echo done >> run_log.txt
