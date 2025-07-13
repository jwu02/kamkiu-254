# 254

## Todos
- take into account when opening a excel file, displays multiple CPK
- color cells in same column, 1 colour for each value

## Installation
- python -m venv venv
- MAC - source ./venv/bin/activate
- WINDOWS - venv\Scripts\activate
- pip install -r requirements.txt
- python main.py

## Notes
- 用WPS表格更改过的 CSV 没办法上传上去 并报以下错误
    - 'utf-8' codec can't decode byte 0xbf in position 0; invalid start byte
    - 问题在于 WPS 表格 saves files with different encodings
        - but pandas tries to read a CSV file that's not actually encoded in UTF-8
    - 需要用另一个打开方式（记事本）来更改数据
- 发货批次表 的排序优先：地区 -> 客户 -> 型号 -> 炉号 -> 发货数 -> 挤压批号 -> 时效批号