
## 如何安装

1. Python

2. 安装依赖

```
pip install keyboard mouse colorama
```

3. 编辑配置文件

示例
```
[技能1]
触发键 = 1
循环 = 是
按下 q
等待 25ms

[基础攻击连招]
触发键 = 1
重复 = 按住时
动作 =
  按下 空格
  按下 q
  按下 w
  按下 e
  按下 r
```

4. 运行脚本
```
python source/macro.py
```
