# 254

## Todos
1. 将性能分为金相（按熔铸炉号）和性能（按时效批号），明天找高工链接访问ME系统获取已经判定过的数据；
2. 若出现任意一个不合格或空值，则生成不出报告；
3. 未有生成报告的给一个标签，回头将这个批次信息根据发货报告紧急度提供实验室优先处理；
4. 当天因数据不齐全需再次生成报告，仅针对原先没有生成报告的批次再次访问ME系统，避免重复生成；
5. 注意CPK的尺寸要适当调整目标范围；

下一阶段延展：
1. 看看如何和出仓单链接，确保出货报告与实物，出仓记录一致，从数量到批次信息；
2. 加上蓝思和富士康的发货报告格式；
3. CPK 数据处理；
4. LAR-OQC数据处理；

- 性能 - 报合不合格的项目
- 警告不合格的报告先不生成？


- take into account when opening a excel file, displays multiple CPK
- color cells in same column, 1 colour for each value

## Installation
- python -m venv venv
- MAC - source ./venv/bin/activate
- WINDOWS - venv\Scripts\activate
- pip install -r requirements.txt
- python main.py

## Version Control
- git remote add origin https://gitee.com/jwu02/kamkiu-254.git
    - git remote add gitee https://gitee.com/jwu02/kamkiu-254.git
- git remote rename origin gitee
- git push -u github
- git push -u gitee

## Notes
- 用WPS表格更改过的 CSV 没办法上传上去 并报以下错误
    - 'utf-8' codec can't decode byte 0xbf in position 0; invalid start byte
    - 问题在于 WPS 表格 saves files with different encodings
        - but pandas tries to read a CSV file that's not actually encoded in UTF-8
    - 需要用另一个打开方式（记事本）来更改数据
- 发货批次表 的排序优先：地区 -> 客户 -> 型号 -> 炉号 -> 发货数 -> 挤压批号 -> 时效批号

## API
- API response data format
```
{
    "msg": "操作成功",
    "code": 0,
    "data": {
        "titleList": [
            {},
            ...
        ],
        "list": [
            {},
            ...
        ]
    }
}
```