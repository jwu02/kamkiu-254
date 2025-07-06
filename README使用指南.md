# 254_v1_1

## Todos
- color cells in same column, 1 colour for each value
- check a specfic directory, report sum for each model code

## Installation 安装
- python -m venv venv
- MAC - source ./venv/bin/activate
- WINDOWS - venv\Scripts\activate
- pip install -r requirements.txt
    - 下载需要的 package
- python main.py
    - 运行程序

## Notes 笔记
- pip freeze > requirements.txt
- ME → A产品每日发货批次表（493）
    - 更改布局，删除第一列挤压批号
- ME -> 型材时效二维码（新）（507）
    - 复杂查询
        - 型号
        - 生产挤压批：（挤压批号）
- ME -> 流程卡二维码记录（pz230）
    - 复杂查询
        - sfc：（时效批号*）
        - 型号
        - 挤压批号
    - 下载 CSV

## 切记
- 用WPS表格更改过的 CSV 没办法上传上去 并报以下错误
    - 'utf-8' codec can't decode byte 0xbf in position 0; invalid start byte
    - 问题在于 WPS 表格 saves files with different encodings
        - but pandas tries to read a CSV file that's not actually encoded in UTF-8
    - 需要用另一个打开方式（记事本）来更改数据
- 发货批次表 的排序优先：地区 -> 客户 -> 型号 -> 炉号 -> 发货数 -> 挤压批号 -> 时效批号