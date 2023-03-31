




## 使用说明
为快速构建json 结构体，使用excel 可以快速构建想要的数据，故弄个程序快速转换

使用SheetName#{ToName!Row} 封装子对象，可以添加每一个对象，也可以添加到某一行的子对象

使用标签关联数据
- toName 存在必填 关联是那一个sheet页下的子项，即每一行都关联子项
- Row 选填 关联是某一行的子项
- AsName 选填，别名 默认item

| test1 | test2                      |
|------------------------|----------------------------|
| a2                     | 21                         |
|                        |                            |
|                        |                            |
| Sheet1                 | Sheet2#{Sheet!A1} |

**导出为**
```json
[
    [
        {
            "Sheet1": [
                {
                    "A": "test1",
                    "B": "test2",
                    "Sheet3Item": [
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        }
                    ]
                },
                {
                    "A": "ab",
                    "B": 1.0,
                    "Sheet3Item": [
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        }
                    ]
                },
                {
                    "A": "test1",
                    "B": "test2",
                    "Sheet3Item": [
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        }
                    ]
                },
                {
                    "A": "ab",
                    "B": 1.0,
                    "Sheet3Item": [
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        },
                        {
                            "test2": 21.0,
                            "test1": "a2"
                        }
                    ]
                }
            ]
        },
        {
            "Sheet4": [
                {
                    "test2": 21.0,
                    "test1": "a2"
                },
                {
                    "test2": 21.0,
                    "test1": "a2"
                },
                {
                    "test2": 21.0,
                    "test1": "a2"
                },
                {
                    "test2": 21.0,
                    "test1": "a2"
                }
            ]
        }
    ]
]
```