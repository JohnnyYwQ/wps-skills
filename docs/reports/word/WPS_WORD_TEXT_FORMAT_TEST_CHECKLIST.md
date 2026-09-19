# Word 文字与段落复杂测试清单（27–38）

状态：已完成三阶段统一实机测试；本组12/12 Task、49/49 Action成功，12/12清单断言通过，人工粗审待进行。详见 [统一结果与耗时分析](WPS_WORD_THREE_STAGE_RESULTS.md)。默认 Windows 已安装宋体、黑体、楷体、Times New Roman、Arial、Calibri。

本组仅测试标题、文字字体/字号、中西文字体、段落、行距、缩进及文字替换。图片和分页计数的既有差异保留，暂不处理；不加入页眉页脚、Agent、重试或并发。

## 执行与通用成功条件

使用已修复指纹的独立 Word 包。每例新建空白纵向文档、独立Task，保存一份DOCX。将示例根目录 `C:/wps-word-text` 替换为本轮新目录，并给输出文件名加本轮唯一前缀。所有JSON均另存于 `src/test/resources/wps_skills/word/text-format/requests`。每例仅提交一次，失败保留具体Action错误，不重试。

**A/T**：逐项检查document、steps、completion的state=succeeded、response.outcome=succeeded、无error，address/taskId一致，data满足现行result schema。Task为completed/succeeded、stop=null，steps数量/顺序/id与请求一致，document.id=task_document；completion.save.id=task_save，completion.pdf=null。创建返回unsaved/readOnly=false。

**W/R**：写入/替换的revisionBefore/After非空；本组实际文字变化要求不同。range或ranges合法、非空，revision=revisionAfter。相邻变更之间仅有读取时，revision衔接一致。标题、字体和段落格式由Action内部回读验证，不额外要求变更响应含未定义的格式快照。

**S**：saveAs返回本例outputPath、docx、sizeBytes>0、saved，replacedExisting=false，failIfExists不带outputResolution。不额外解析DOCX或截图验收，人工粗审独立记录。

## 耗时统计口径

- **Task端到端耗时**：测试进程启动提交子进程前，到完整stdout/stderr接收且子进程退出。包括Python/Skill加载、JSON读取和预检、Task执行、回读、持久化、清理和响应输出；不包含测试素材准备、人工审核、后置断言与报告生成。
- **每个Action执行耗时**：在测试子进程中围绕正式WordTask.execute调用，用perf_counter_ns测量进入到返回/异常的时间。包括线程派发、桥接通信、WPS操作及内部回读；不含Task层调用前的参数校验/回执发布，也不是纯COM耗时。首个createDocument可能包含惰性启动成本。
- **其他耗时**：端到端耗时减已执行Action耗时之和，仅标“其余开销”，不冒称已精确区分启动与清理。未执行Action耗时为null；异常/超时单独记录，不能补零或混入成功平均值。
- 逐Task记录总耗时、Action合计和其余开销；逐Action记录步骤id、名称、outcome、traceId及毫秒数。另按Action名称汇总成功次数、平均值、中位数、最小和最大值，指出最慢Task及Action。
- 每例一次是功能运行耗时，不代表稳定性能基准，不宣称P95或吞吐率。记录Windows/WPS版本、包版本、桌面会话和WPS是否已运行；本轮前台已有文档等状态也需记录。
- 正式运行时现已内置Task/Action/桥接计时；优先采集正式trace，参见 WPS_WORD_TIMING.md。此前的测试专用包装入口仅为历史备用，不应再叠加使用。contract和响应结构未改变。计时数据单独写JSON/CSV，原始Action/Task响应照常保留。无计时记录时标缺失，不用trace文件首尾时间猜测。

## 用例索引

| 编号 | 用例 | 响应结果 | 耗时 | 人工粗审 |
| --- | --- | --- | --- | --- |
| 27 | [多级标题与正文交错](#case-27) | 通过 | 2197.65 ms | 未审阅 |
| 28 | [单段多字号与强调格式](#case-28) | 通过 | 2085.47 ms | 未审阅 |
| 29 | [同一 run 的中西文字体](#case-29) | 通过 | 2079.09 ms | 未审阅 |
| 30 | [同段多组中西文字体](#case-30) | 通过 | 2161.45 ms | 未审阅 |
| 31 | [连续段落对齐与段间距](#case-31) | 通过 | 2267.37 ms | 未审阅 |
| 32 | [六种行距模式连续使用](#case-32) | 通过 | 2233.38 ms | 未审阅 |
| 33 | [大小字号与行距交叉](#case-33) | 通过 | 2228.25 ms | 未审阅 |
| 34 | [首行、悬挂与左右缩进](#case-34) | 通过 | 2275.18 ms | 未审阅 |
| 35 | [多 Action 追加与格式恢复](#case-35) | 通过 | 2574.21 ms | 未审阅 |
| 36 | [替换完整格式段落并保留邻段](#case-36) | 通过 | 2617.74 ms | 未审阅 |
| 37 | [跨 run 查找与多处格式化替换](#case-37) | 通过 | 2799.61 ms | 未审阅 |
| 38 | [综合文字报告与定点改写](#case-38) | 通过 | 2795.17 ms | 未审阅 |

<a id="case-27"></a>

## 27 · 多级标题与正文交错

**自然语言效果 / 人工粗审**：依次出现一级、二级、三级标题和各自正文；标题分别 24/18/14 pt，一级居中；正文 12 pt、左对齐、黑色、不加粗。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "outline",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "heading",
            "level": 1,
            "runs": [
              {
                "text": "文字报告",
                "format": {
                  "fontSizePt": 24,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "center"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "报告正文 English 123。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          },
          {
            "kind": "heading",
            "level": 2,
            "runs": [
              {
                "text": "第一章",
                "format": {
                  "fontSizePt": 18,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "left"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "章正文。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          },
          {
            "kind": "heading",
            "level": 3,
            "runs": [
              {
                "text": "第一节",
                "format": {
                  "fontSizePt": 14,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "left"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "节正文。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-27.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- outline 按顺序写入六个块；标题 level 与请求一致，格式由写入回读判定。

**预期 Task Response**：通用T；steps[*].id=["outline"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2197.65 ms；外部端到端＝2269.46 ms；各Action＝task_document=1498.22 ms；outline=326.52 ms；task_save=196.16 ms；Task ID＝`task-81c68570f502ec8f32596b0dfbd59f50e67cc588ef274cc81fcbdc88847c5aa4`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/27.stdout.json)；人工粗审＝未审阅。

<a id="case-28"></a>

## 28 · 单段多字号与强调格式

**自然语言效果 / 人工粗审**：一段内依次为普通、18 pt 加粗、10 pt 斜体、14 pt 下划线蓝字、恢复普通；格式界限与各文字片段一致。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "mixed",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "普通 ",
                "format": {
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              },
              {
                "text": "重点 ",
                "format": {
                  "fontSizePt": 18,
                  "bold": true,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              },
              {
                "text": "注释 ",
                "format": {
                  "fontSizePt": 10,
                  "bold": false,
                  "italic": true,
                  "underline": "none",
                  "color": "#000000"
                }
              },
              {
                "text": "链接样式 ",
                "format": {
                  "fontSizePt": 14,
                  "bold": false,
                  "italic": false,
                  "underline": "single",
                  "color": "#1F4E79"
                }
              },
              {
                "text": "恢复普通",
                "format": {
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ]
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-28.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- mixed 非空写入成功；人工按五个片段检查字号与强调边界，不把 run 分割数量当作固定响应要求。

**预期 Task Response**：通用T；steps[*].id=["mixed"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2085.47 ms；外部端到端＝2155.19 ms；各Action＝task_document=1519.65 ms；mixed=223.46 ms；task_save=183.81 ms；Task ID＝`task-1f7e6342f25e49fef3cb18f5b670c9b0b16cb1300eee577e32aee7ed71aa7006`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/28.stdout.json)；人工粗审＝未审阅。

<a id="case-29"></a>

## 29 · 同一 run 的中西文字体

**自然语言效果 / 人工粗审**：三个段落各在一个 run 内混排中文、英文、数字与标点，分别采用宋体/Times New Roman、黑体/Arial、楷体/Calibri。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "bilingual",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "中文字体 English 123，标点 punctuation.",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "westernFontFamily": "Times New Roman",
                  "fontSizePt": 14
                }
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "中文字体 English 123，标点 punctuation.",
                "format": {
                  "eastAsiaFontFamily": "黑体",
                  "westernFontFamily": "Arial",
                  "fontSizePt": 14
                }
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "中文字体 English 123，标点 punctuation.",
                "format": {
                  "eastAsiaFontFamily": "楷体",
                  "westernFontFamily": "Calibri",
                  "fontSizePt": 14
                }
              }
            ]
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-29.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 使用 eastAsiaFontFamily/ westernFontFamily，不同时传入 fontFamily；不对中性标点强行指定其归属字体系。

**预期 Task Response**：通用T；steps[*].id=["bilingual"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2079.09 ms；外部端到端＝2149.41 ms；各Action＝task_document=1497.28 ms；bilingual=212.49 ms；task_save=204.62 ms；Task ID＝`task-2935ed8942fa73504afd92b9bb9973ba66207897c349b8a7b06f65b8fe0d2603`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/29.stdout.json)；人工粗审＝未审阅。

<a id="case-30"></a>

## 30 · 同段多组中西文字体

**自然语言效果 / 人工粗审**：单段依次为宋体/Times New Roman、黑体/Arial、楷体/Calibri，再回宋体/Times New Roman；字号分别 12/16/14/12 pt。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "fonts",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "第一组 Alpha ",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "westernFontFamily": "Times New Roman",
                  "fontSizePt": 12
                }
              },
              {
                "text": "第二组 Beta ",
                "format": {
                  "eastAsiaFontFamily": "黑体",
                  "westernFontFamily": "Arial",
                  "fontSizePt": 16
                }
              },
              {
                "text": "第三组 Gamma ",
                "format": {
                  "eastAsiaFontFamily": "楷体",
                  "westernFontFamily": "Calibri",
                  "fontSizePt": 14
                }
              },
              {
                "text": "恢复 Delta",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "westernFontFamily": "Times New Roman",
                  "fontSizePt": 12
                }
              }
            ]
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-30.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- fonts 满足 W；内部回读负责字体与字号，人工检查相邻片段无串扰。

**预期 Task Response**：通用T；steps[*].id=["fonts"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2161.45 ms；外部端到端＝2231.61 ms；各Action＝task_document=1568.84 ms；fonts=219.92 ms；task_save=205.74 ms；Task ID＝`task-5a38c93c213b220dd1e8a709f3c8c92817bc3114633dff5cf53732f3040fe136`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/30.stdout.json)；人工粗审＝未审阅。

<a id="case-31"></a>

## 31 · 连续段落对齐与段间距

**自然语言效果 / 人工粗审**：四段分别左对齐、居中、右对齐、两端对齐；段前/后依次 0/6、6/12、12/6、0/0 pt。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "alignments",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "alignment": "left",
              "spaceBeforePt": 0,
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "alignment": "center",
              "spaceBeforePt": 6,
              "spaceAfterPt": 12
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "alignment": "right",
              "spaceBeforePt": 12,
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。 对齐示例 English words 用于观察段落布局。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "alignment": "justify",
              "spaceBeforePt": 0,
              "spaceAfterPt": 0
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-31.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 长段落提供自然换行；两端对齐不要求末行填满。

**预期 Task Response**：通用T；steps[*].id=["alignments"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2267.37 ms；外部端到端＝2338.26 ms；各Action＝task_document=1659.19 ms；alignments=261.10 ms；task_save=186.84 ms；Task ID＝`task-1bd21cd41dcc6398c623dc531224d2090ba42902c76f17b7c273894ec1b11d57`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/31.stdout.json)；人工粗审＝未审阅。

<a id="case-32"></a>

## 32 · 六种行距模式连续使用

**自然语言效果 / 人工粗审**：六个长段落依次采用单倍、1.5 倍、双倍、固定 18 pt、最小 18 pt、1.25 倍行距，文字均 12 pt。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "spacing",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。 single 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "single"
              },
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。 oneAndHalf 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "oneAndHalf"
              },
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。 double 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "double"
              },
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。 exact 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "exact",
                "points": 18
              },
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。 atLeast 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "atLeast",
                "points": 18
              },
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。 multiple 行距示例 English words。",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "multiple",
                "value": 1.25
              },
              "spaceAfterPt": 6
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-32.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 必须区分 kind=exact 与 atLeast；不得仅按视觉高度相近判断二者相同。格式成功以 Action 内部回读为准。

**预期 Task Response**：通用T；steps[*].id=["spacing"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2233.38 ms；外部端到端＝2302.95 ms；各Action＝task_document=1575.66 ms；spacing=296.10 ms；task_save=193.22 ms；Task ID＝`task-529135cffa5886affd89023845b0f6516aeb886ce754afa65539dc5fe6f9d6ae`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/32.stdout.json)；人工粗审＝未审阅。

<a id="case-33"></a>

## 33 · 大小字号与行距交叉

**自然语言效果 / 人工粗审**：三个段落都混合 12 pt 正文与 24 pt 大字，分别固定 18 pt、最小 18 pt、1.5 倍行距。固定 18 pt 可能裁切大字，这是预期展示，不要求不裁切。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "size_spacing",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              },
              {
                "text": "大字 BIG ",
                "format": {
                  "fontSizePt": 24
                }
              },
              {
                "text": "后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "exact",
                "points": 18
              }
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              },
              {
                "text": "大字 BIG ",
                "format": {
                  "fontSizePt": 24
                }
              },
              {
                "text": "后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "atLeast",
                "points": 18
              }
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              },
              {
                "text": "大字 BIG ",
                "format": {
                  "fontSizePt": 24
                }
              },
              {
                "text": "后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 后续普通文字 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "multiple",
                "value": 1.5
              }
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-33.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 三个段落均需写入成功；人工观察固定行距与自动扩展行距的差异，不对渲染行高另做数值断言。

**预期 Task Response**：通用T；steps[*].id=["size_spacing"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2228.25 ms；外部端到端＝2298.56 ms；各Action＝task_document=1586.57 ms；size_spacing=289.31 ms；task_save=193.94 ms；Task ID＝`task-3b47ada5623690f6d34829addf52fdf122c88e3438ceae6841c60bdfc9e43549`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/33.stdout.json)；人工粗审＝未审阅。

<a id="case-34"></a>

## 34 · 首行、悬挂与左右缩进

**自然语言效果 / 人工粗审**：三个长段落：首行缩进24 pt；左缩进24 pt且首行-24 pt的悬挂段；左右均36 pt且首行24 pt的窄正文。段前后分别设置。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "indents",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "firstLineIndentPt": 24,
              "leftIndentPt": 0,
              "rightIndentPt": 0,
              "spaceBeforePt": 0,
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "firstLineIndentPt": -24,
              "leftIndentPt": 24,
              "rightIndentPt": 0,
              "spaceBeforePt": 6,
              "spaceAfterPt": 6
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 缩进示例 English long paragraph。 ",
                "format": {
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "firstLineIndentPt": 24,
              "leftIndentPt": 36,
              "rightIndentPt": 36,
              "spaceBeforePt": 12,
              "spaceAfterPt": 12
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-34.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 负 firstLineIndentPt 表示悬挂，不作为非法参数；两侧缩进不等于页面页边距。

**预期 Task Response**：通用T；steps[*].id=["indents"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2275.18 ms；外部端到端＝2345.79 ms；各Action＝task_document=1645.48 ms；indents=260.89 ms；task_save=213.88 ms；Task ID＝`task-fca36da1f78d531ca86d6eda520ca3efb366a1b39c5d020444a349e2cc26a443`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/34.stdout.json)；人工粗审＝未审阅。

<a id="case-35"></a>

## 35 · 多 Action 追加与格式恢复

**自然语言效果 / 人工粗审**：连续追加居中24 pt标题、普通正文、右对齐16 pt红色斜体强调段，再追加普通正文；最后正文明确恢复为12 pt黑色、非强调、左对齐单倍行距。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "title",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "heading",
            "level": 1,
            "runs": [
              {
                "text": "连续追加测试",
                "format": {
                  "fontSizePt": 24,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "center"
            }
          }
        ]
      }
    },
    {
      "id": "body",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "第一段正文 English。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    },
    {
      "id": "emphasis",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "强调内容 Important",
                "format": {
                  "fontSizePt": 16,
                  "bold": true,
                  "italic": true,
                  "underline": "single",
                  "color": "#C00000"
                }
              }
            ],
            "format": {
              "alignment": "right",
              "lineSpacing": {
                "kind": "double"
              }
            }
          }
        ]
      }
    },
    {
      "id": "reset",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "恢复正文 Normal text。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-35.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- 四个 writeContent 均满足 W；后一个 revisionBefore 等于前一个 revisionAfter。显式恢复格式字段以免把未指定格式的继承当成错误。

**预期 Task Response**：通用T；steps[*].id=["title", "body", "emphasis", "reset"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2574.21 ms；外部端到端＝2645.06 ms；各Action＝task_document=1693.96 ms；title=229.00 ms；body=110.46 ms；emphasis=72.40 ms；reset=74.19 ms；task_save=227.89 ms；Task ID＝`task-9891c2070ff8b48068bc34c46c148736a39572e5f6cec4e72aefcbcca2b47424`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/35.stdout.json)；人工粗审＝未审阅。

<a id="case-36"></a>

## 36 · 替换完整格式段落并保留邻段

**自然语言效果 / 人工粗审**：三段中第一段“保留开头”及末段“保留结尾”保留；混合格式中段替换成黑体/Arial、16 pt加粗、居中双倍行距的“新说明 New description”。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "seed",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "保留开头",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "旧说明 ",
                "format": {
                  "fontSizePt": 10,
                  "italic": true
                }
              },
              {
                "text": "Old description",
                "format": {
                  "fontSizePt": 18,
                  "bold": true
                }
              }
            ],
            "format": {
              "alignment": "right"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "保留结尾",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    },
    {
      "id": "locate",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 16000,
          "maxParagraphs": 256,
          "maxRuns": 1024
        }
      }
    },
    {
      "id": "replace_para",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "range",
          "range": {
            "$ref": {
              "step": "locate",
              "path": [
                "data",
                "paragraphs",
                1,
                "range"
              ]
            }
          }
        },
        "replacement": {
          "kind": "blocks",
          "blocks": [
            {
              "kind": "paragraph",
              "runs": [
                {
                  "text": "新说明 New description",
                  "format": {
                    "westernFontFamily": "Arial",
                    "eastAsiaFontFamily": "黑体",
                    "fontSizePt": 16,
                    "bold": true
                  }
                }
              ],
              "format": {
                "alignment": "center",
                "lineSpacing": {
                  "kind": "double"
                }
              }
            }
          ]
        }
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-36.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- locate.paragraphs[*].text=["保留开头\n","旧说明 Old description\n","保留结尾"]，complete=true、truncated=false。
- replace_para.matchedCount=1，ranges长度1，revisionBefore=locate.revision；邻段效果由人工粗审和内部写入范围保证。

**预期 Task Response**：通用T；steps[*].id=["seed", "locate", "replace_para"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2617.74 ms；外部端到端＝2689.22 ms；各Action＝task_document=1704.84 ms；seed=289.88 ms；locate=141.31 ms；replace_para=113.15 ms；task_save=216.01 ms；Task ID＝`task-20a399f6d381430031636fd1562ffcbbca33993c32ed708a1b5ac0dd48eb3387`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/36.stdout.json)；人工粗审＝未审阅。

<a id="case-37"></a>

## 37 · 跨 run 查找与多处格式化替换

**自然语言效果 / 人工粗审**：两段原本分别拆成不同字体/强调 run 的“项目Alpha待确认”被查找到，整串替换为18 pt蓝色加粗“项目Alpha已确认”；第三段“独立项待确认”在一次替换子串后变成“独立项已确认”。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "seed",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "项目",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12
                }
              },
              {
                "text": "Alpha",
                "format": {
                  "westernFontFamily": "Arial",
                  "fontSizePt": 16,
                  "bold": true
                }
              },
              {
                "text": "待确认",
                "format": {
                  "fontSizePt": 12
                }
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "项目Alpha",
                "format": {
                  "fontSizePt": 14,
                  "italic": true
                }
              },
              {
                "text": "待确认",
                "format": {
                  "fontSizePt": 12
                }
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "独立项待确认",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    },
    {
      "id": "find_full",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "项目Alpha待确认",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "replace_two",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "query",
          "query": {
            "scope": {
              "kind": "document"
            },
            "text": "项目Alpha待确认",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 2
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "项目Alpha已确认",
              "format": {
                "fontSizePt": 18,
                "bold": true,
                "color": "#1F4E79"
              }
            }
          ]
        }
      }
    },
    {
      "id": "find_remaining",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "待确认",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "replace_one",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "query",
          "query": {
            "scope": {
              "kind": "document"
            },
            "text": "待确认",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "已确认",
              "format": {
                "fontSizePt": 12,
                "bold": false,
                "color": "#000000"
              }
            }
          ]
        }
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-37.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- find_full.matches长度2，text均为项目Alpha待确认，truncated=false、remainingRange=null。
- replace_two.matchedCount=2，ranges长度2；find_remaining.matches长度1，text=待确认，revision=replace_two.revisionAfter。
- replace_one.matchedCount=1，ranges长度1；所有 match.range 使用各自find.revision。

**预期 Task Response**：通用T；steps[*].id=["seed", "find_full", "replace_two", "find_remaining", "replace_one"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2799.61 ms；外部端到端＝2869.53 ms；各Action＝task_document=1731.76 ms；seed=298.76 ms；find_full=72.63 ms；replace_two=105.03 ms；find_remaining=62.45 ms；replace_one=84.22 ms；task_save=200.19 ms；Task ID＝`task-11c1716eb899230fb649a8b5650c2bed8a9eae96fd4ff4b1627a461d79f741b9`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/37.stdout.json)；人工粗审＝未审阅。

<a id="case-38"></a>

## 38 · 综合文字报告与定点改写

**自然语言效果 / 人工粗审**：一份纯文字报告包含多级标题、中西文混排、强调段、正文缩进和不同段间距；将“状态：草稿”更新为“状态：定稿”，追加结束说明并保存。

**前置条件**：新建空白文档；默认所用字体已安装；不依赖其他用例。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "createDocument"
    },
    "params": {}
  },
  "steps": [
    {
      "id": "opening",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "heading",
            "level": 1,
            "runs": [
              {
                "text": "文字排版综合报告",
                "format": {
                  "fontSizePt": 24,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "center"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "状态：草稿",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          },
          {
            "kind": "heading",
            "level": 2,
            "runs": [
              {
                "text": "一、正文",
                "format": {
                  "fontSizePt": 18,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "left"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 中文正文 English 123。 ",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "westernFontFamily": "Times New Roman",
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "alignment": "justify",
              "firstLineIndentPt": 24,
              "lineSpacing": {
                "kind": "oneAndHalf"
              },
              "spaceAfterPt": 6
            }
          }
        ]
      }
    },
    {
      "id": "details",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "heading",
            "level": 3,
            "runs": [
              {
                "text": "1.1 重点事项",
                "format": {
                  "fontSizePt": 14,
                  "bold": true,
                  "color": "#1F4E79"
                }
              }
            ],
            "format": {
              "alignment": "left"
            }
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "注意：",
                "format": {
                  "bold": true,
                  "color": "#C00000",
                  "fontSizePt": 14
                }
              },
              {
                "text": "请确认字体与行距。",
                "format": {
                  "bold": false,
                  "color": "#000000",
                  "fontSizePt": 12
                }
              }
            ],
            "format": {
              "lineSpacing": {
                "kind": "atLeast",
                "points": 18
              },
              "spaceBeforePt": 6,
              "spaceAfterPt": 12
            }
          }
        ]
      }
    },
    {
      "id": "find_draft",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "草稿",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "finalize",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "query",
          "query": {
            "scope": {
              "kind": "document"
            },
            "text": "草稿",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "定稿",
              "format": {
                "fontSizePt": 12,
                "bold": true
              }
            }
          ]
        }
      }
    },
    {
      "id": "ending",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "报告结束 End of report。",
                "format": {
                  "westernFontFamily": "Times New Roman",
                  "eastAsiaFontFamily": "宋体",
                  "fontSizePt": 12,
                  "bold": false,
                  "italic": false,
                  "underline": "none",
                  "color": "#000000"
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "single"
              }
            }
          }
        ]
      }
    }
  ],
  "completion": [
    {
      "address": {
        "app": "word",
        "action": "saveAs"
      },
      "params": {
        "outputPath": "C:/wps-word-text/outputs/text-38.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用A/W/R/S，另需：

- find_draft 恰好命中草稿一次，truncated=false、remainingRange=null。finalize.matchedCount=1、ranges长度1。
- 各文字变更revision衔接正确，ending之后saveAs返回saved。标题和正文的混合格式均由各自Action回读负责。

**预期 Task Response**：通用T；steps[*].id=["opening", "details", "find_draft", "finalize", "ending"]；仅saveAs，无PDF。

**执行记录**：Task和Action均成功、全部响应断言通过；内部Task耗时＝2795.17 ms；外部端到端＝2869.76 ms；各Action＝task_document=1729.83 ms；opening=310.65 ms；details=102.79 ms；find_draft=77.39 ms；finalize=102.77 ms；ending=77.42 ms；task_save=190.55 ms；Task ID＝`task-3102001aa91d3610c017978e3dae548c4618113d6d0ffa90544e4c1c58e22549`；[完整响应](../../../build/evidence/word-three-stage-20260917-234351/windows/responses/38.stdout.json)；人工粗审＝未审阅。

## 计时采集入口

测试专用入口位于 [run_timed_cases.py](../../../src/test/python/tests/word/timing/run_timed_cases.py)，与同目录 timed_entry.py 一起部署。它依次提交已部署的请求，保存原始响应、原始纳秒计时及 task-timings/action-timings/action-summary 三组 JSON/CSV；逐例业务断言仍按本清单核对。此脚本已通过语法检查，尚未在 Windows 验证。

```text
python run_timed_cases.py --skill "<完整包目录>" --requests "<已部署请求目录>" --output "<全新计时结果目录>"
```

必须在Windows交互桌面运行。准备时另外记录WPS版本、进程/打开文档状态；脚本自动记录Python、会话与包清单摘要。不要在同一个已执行请求上重复运行来“补计时”。


正式计时更新：已实现 [正式执行链路日志](WPS_WORD_TIMING.md)。后续本组测试使用该日志记录Action/通信耗时；外部总耗时可由测试进程另行记录。以上旧测试包装脚本不是正式日志的必需依赖。


本轮使用正式运行时trace，不叠加旧计时包装器。Task耗时为request.total，包含清理与响应输出；外部耗时由提交进程测量。全部已执行Action的耗时均已记录，详见统一分析报告及CSV。
