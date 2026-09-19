# Word Skill 实机功能测试清单

日期：2026-09-17。状态：指纹修复后重跑 18 项，18 项响应断言全部通过；人工审阅待进行。详见 [实机结果](WPS_WORD_FUNCTIONAL_TEST_RESULTS.md)。

## 本轮约定

人工构造 Agent 可能提交的固定 JSON，经正式 Task Client 执行；用预期 Action Response 与 Task Response 判定。每例另列自然语言效果，供用户在 WPS 或 PDF 中粗看。

信任 Action 内部原生回读：不另建最终文档解析、截图识别或磁盘哈希验收。人工审阅独立记录；若响应通过但人工发现明显异常，保留两项结果并标记“响应与人工观察不一致”，不能直接宣称功能完全符合预期。

本轮仅覆盖正常功能。Task 重试、回执恢复、重名重试、并发、身份保护专项、取消、超时、进程丢失、Quarantine、启动器及 Agent 行为均不测。现有综合验收脚本不直接整套执行，以免引入上述范围。

## 准备与执行方式

- 使用最新完整包 `build/archive/runs/word-fingerprint-fix-20260917-223921/wps-word/`，不安装到 WorkBuddy。
- 下文 `C:/wps-word-functional` 是示例测试根目录；运行前统一替换成 Windows 上本轮全新绝对目录。先创建 `fixtures`、`requests`、`outputs`、`responses`，输出文件起初不存在。
- 每个“已有文件”用例使用自己编号的独立、已保存、可写 DOCX，具体内容见对应前置条件；不得复用上一例已被修改的文档。已有文件应在未编辑时正常报告 saved。后续复现已确认本轮素材刚经 Documents.Open 打开时 Saved=true，是 openDocument 内部指纹读取期间变成 false，不能将改用原生素材当作已验证的解决办法。准备素材不是用例通过的证据。
- 新建文档基线：普通空白纵向文档，不启用首页不同或奇偶页不同；示例字体宋体、Times New Roman 可用。
- 图片固定为 `fixtures/sample.png`：200 × 100 像素、蓝底白字 TEST。准备时记录其 SHA-256 和字节数，供响应断言比较；素材及 SHA-256 已在本轮 fixture-metadata.json 留档。
- 下文 JSON 是完整请求，每例保存为 `requests/编号.json`。仅有根目录需部署替换；不需要 Agent 生成或改写请求。
- 每例前台提交一次，保留完整 stdout JSON 与 stderr；不自动重试。请求成功后通常被消费，以回执中原请求作为留档。失败保留响应与现场，记录阻塞，不追加补救动作。

```text
python "<skill-dir>/scripts/word.py" --app word --task-file "<本轮目录>/requests/01.json"
```

## 通用响应断言

各例“预期 Action Response”均指对应条目的 `response`，业务字段若无特别说明均位于 `response.data`。

### A：每个 Action 都适用

- `outcome == "succeeded"`，无错误对象；`address` 与请求相符。
- `taskId` 与外层 Task Response 相同；`data` 符合该 Action 的正式 result schema。
- 新建：`documentState == {"persistenceState":"unsaved","readOnly":false}`，`revision` 非空。
- 打开本清单的已保存素材：`documentState == {"persistenceState":"saved","readOnly":false}`；`artifact.path` 为对应素材路径、`format == "docx"`、`sizeBytes > 0`。
- Revision、Task ID、trace 等动态值不写死；仅检查契约和明确关联，不把不透明 ID 当作数值序号。

### W / R：内容变更

- W（writeContent）：`revisionBefore`、`revisionAfter` 非空；`range` 合法、`range.revision == revisionAfter`。本清单非空写入要求 `start < end`。文字和格式由内部回读保证，不额外安排读操作，也不期待响应含正文或格式快照。
- R（replaceContent、insertTable）：变更前后 Revision 非空；返回 ranges/table.range 使用 `revisionAfter`，坐标合法。替换文字范围非空，删除结果允许空范围。
- Revision 仅按相等关系判断；需要范围或节选择的 `$ref` 均取自本 Task 前序步骤，不手工构造 Revision。

### B：分页符

`break.type == "page"`，`sectionCountBefore == sectionCountAfter == 1`；`break.range` 非空且 revision 等于 `revisionAfter`。分页效果由 Action 回读负责，不额外断言页面渲染数量。

### S：另存 DOCX

`artifact.path` 为本例 `outputs/编号.docx`，`format == "docx"`，`sizeBytes > 0`；`documentState.persistenceState == "saved"`；`replacedExisting == false`；本清单使用 `failIfExists`，响应不含 `outputResolution`（该字段仅适用于 `renameIfExists`）。不测试重名候选。

### P：导出 PDF

`artifact.path` 为本例 `outputs/编号.pdf`，`format == "pdf"`，`sizeBytes > 0`；`replacedExisting == false`；本清单使用 `failIfExists`，响应不含 `outputResolution`；`revisionBefore == revisionAfter`，`documentStateBefore == documentStateAfter`。

### T：每例的 Task Response

- `type == "task.response"`，`app == "word"`，`state == "completed"`，`outcome == "succeeded"`，`stop == null`。
- `document.id == "task_document"`；`document.state == "succeeded"`。
- `steps` 的数量、顺序、id、address 与输入一致，每项 `state == "succeeded"`，每个 `response` 满足 A 及本例断言。
- 选择保存时 `completion.save.id == "task_save"`，选择 PDF 时 `completion.pdf.id == "task_pdf"`，对应条目 state 均为 succeeded；未选择的条目严格为 null。
- 不只检查外层 succeeded，必须检查 document、每个 step 及 completion 的实际响应。
- `cleanup`、`taskFile` 等诊断原样留存，异常另记，不把资源生命周期扩为本轮测试专项，也不把 cleanup 当作文档效果判定。

## 用例索引

| 编号 | 用例 | 自动响应结果 | 人工粗审 |
| --- | --- | --- | --- |
| 01 | [新建文字，不保存](#case-01) | 通过 | 未审阅 |
| 02 | [文字与段落格式](#case-02) | 通过 | 未审阅 |
| 03 | [在指定段落前插入](#case-03) | 通过 | 未审阅 |
| 04 | [读取及查找](#case-04) | 通过 | 未审阅 |
| 05 | [批量替换文字](#case-05) | 通过 | 未审阅 |
| 06 | [替换指定完整段落](#case-06) | 通过 | 未审阅 |
| 07 | [删除指定完整段落](#case-07) | 通过 | 未审阅 |
| 08 | [插入表格](#case-08) | 通过 | 未审阅 |
| 09 | [插入图片](#case-09) | 通过 | 未审阅 |
| 10 | [页面方向与页边距](#case-10) | 通过 | 未审阅 |
| 11 | [页眉页脚](#case-11) | 通过 | 未审阅 |
| 12 | [插入分页符](#case-12) | 通过 | 未审阅 |
| 13 | [分节与第二节横向](#case-13) | 通过 | 未审阅 |
| 14 | [编辑已有文件并保存](#case-14) | 通过 | 未审阅 |
| 15 | [编辑已有文件后另存](#case-15) | 通过 | 未审阅 |
| 16 | [编辑已有文件，不保存](#case-16) | 通过 | 未审阅 |
| 17 | [修改后仅导出 PDF](#case-17) | 通过 | 未审阅 |
| 18 | [完整报告：DOCX 与 PDF](#case-18) | 通过 | 未审阅 |

<a id="case-01"></a>

## 01 · 新建文字，不保存

**自然语言效果 / 人工粗审**：WPS 中出现一份未保存的新文档：标题“项目周报”，下方依次是中文正文和英文正文。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "write",
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
                "text": "项目周报"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "本周完成接口联调。"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "Next week: release 1.0."
              }
            ]
          }
        ]
      }
    }
  ],
  "completion": []
}
```

**预期 Action Response**：通用 A，以及 write：符合 W（写入）断言。

**预期 Task Response**：通用 T；`steps[*].id == ["write"]`；completion.save == null；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-dfebaa64d1b5ecde2a9b37d9be436a066545530c82da8b0604a1237fe1cee483`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/01.stdout.json)。

<a id="case-02"></a>

## 02 · 文字与段落格式

**自然语言效果 / 人工粗审**：一份格式样张：蓝色、18 磅、加粗、居中的一级标题；正文中文宋体、西文 Times New Roman，12 磅，首行缩进 24 磅，1.5 倍行距，段前后各 6 磅。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "write",
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
                "text": "格式样张",
                "format": {
                  "fontSizePt": 18,
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
                "text": "中文字体与 English 123 混排。",
                "format": {
                  "eastAsiaFontFamily": "宋体",
                  "westernFontFamily": "Times New Roman",
                  "fontSizePt": 12,
                  "bold": false
                }
              }
            ],
            "format": {
              "alignment": "left",
              "lineSpacing": {
                "kind": "oneAndHalf"
              },
              "firstLineIndentPt": 24,
              "spaceBeforePt": 6,
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
        "outputPath": "C:/wps-word-functional/outputs/02.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 write：符合 W；具体格式效果由 writeContent 内部回读判定，响应不含格式快照，不能断言不存在的字段。

**预期 Task Response**：通用 T；`steps[*].id == ["write"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-6f417fabf0c43bab0714efed3a9bf5c960bce885ff332bbd13ede26e7fa26d88`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/02.stdout.json)。

<a id="case-03"></a>

## 03 · 在指定段落前插入

**自然语言效果 / 人工粗审**：原来的“甲、乙、丙”三段变成“甲、新增段、乙、丙”，原段落保持原样。

**前置条件**：`fixtures/03.docx`：恰好三段：甲 / 乙 / 丙；无表格或额外空段落。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/03.docx"
    }
  },
  "steps": [
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "insert",
      "address": {
        "app": "word",
        "action": "writeContent"
      },
      "params": {
        "anchor": {
          "kind": "before",
          "range": {
            "$ref": {
              "step": "read",
              "path": [
                "data",
                "paragraphs",
                1,
                "range"
              ]
            }
          }
        },
        "blocks": [
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "新增段"
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
        "outputPath": "C:/wps-word-functional/outputs/03.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 read：paragraphs[*].text == ["甲\n","乙\n","丙"]，truncated == false；insert：符合 W，range.start 等于 read.paragraphs[1].range.start。

**预期 Task Response**：通用 T；`steps[*].id == ["read", "insert"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-1dbfa14b3dc6bb07f40c027976b0ff0036cdb9ae9e9a864131c7f14902c79f5e`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/03.stdout.json)。

<a id="case-04"></a>

## 04 · 读取及查找

**自然语言效果 / 人工粗审**：文档保持原样。返回的正文是“Alpha alpha Alphabet”和“项目：待确认”；区分大小写的完整词 Alpha 有 1 处，不区分大小写有 2 处，“不存在”有 0 处。

**前置条件**：`fixtures/04.docx`：恰好两段：Alpha alpha Alphabet / 项目：待确认。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/04.docx"
    }
  },
  "steps": [
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "exact",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "Alpha",
          "caseSensitive": true,
          "wholeWord": true
        },
        "limit": 50
      }
    },
    {
      "id": "ignore_case",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "Alpha",
          "caseSensitive": false,
          "wholeWord": true
        },
        "limit": 50
      }
    },
    {
      "id": "missing",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "不存在",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 50
      }
    }
  ],
  "completion": []
}
```

**预期 Action Response**：通用 A，以及 read：paragraphs[*].text == ["Alpha alpha Alphabet\n","项目：待确认"]，truncated == false；exact.matches 的 text 列表为 ["Alpha"]、范围 [0,5)；ignore_case 为 ["Alpha","alpha"]、范围依次 [0,5)、[6,11)；missing.matches == []。三个 find 均 truncated == false，remainingRange == null，range.revision 与各自 data.revision 一致。

**预期 Task Response**：通用 T；`steps[*].id == ["read", "exact", "ignore_case", "missing"]`；completion.save == null；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-aa453b64843266c86e7929a1cb85f72251151d141835b59f999b169d38b7449f`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/04.stdout.json)。

<a id="case-05"></a>

## 05 · 批量替换文字

**自然语言效果 / 人工粗审**：“设计：待确认”“开发：待确认”“测试：待确认”三段中的“待确认”全部变成“已确认”。

**前置条件**：`fixtures/05.docx`：恰好三段：设计：待确认 / 开发：待确认 / 测试：待确认。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/05.docx"
    }
  },
  "steps": [
    {
      "id": "replace",
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
          "expectedMatchCount": 3
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "已确认"
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
        "outputPath": "C:/wps-word-functional/outputs/05.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 replace：matchedCount == 3，len(ranges) == 3，各范围符合 R（变更）断言。

**预期 Task Response**：通用 T；`steps[*].id == ["replace"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-684c89cb8490437c180a14712fe2f6a8c4491603eeba80f09a80d7595692c571`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/05.stdout.json)。

<a id="case-06"></a>

## 06 · 替换指定完整段落

**自然语言效果 / 人工粗审**：三段“开头、旧说明、结尾”变成“开头、新说明、结尾”；中间段落加粗、居中。

**前置条件**：`fixtures/06.docx`：恰好三段：开头 / 旧说明 / 结尾。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/06.docx"
    }
  },
  "steps": [
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "replace",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "range",
          "range": {
            "$ref": {
              "step": "read",
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
                  "text": "新说明",
                  "format": {
                    "bold": true
                  }
                }
              ],
              "format": {
                "alignment": "center"
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
        "outputPath": "C:/wps-word-functional/outputs/06.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 read：paragraphs[*].text == ["开头\n","旧说明\n","结尾"]，truncated == false；replace：matchedCount == 1，len(ranges) == 1，符合 R。

**预期 Task Response**：通用 T；`steps[*].id == ["read", "replace"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-533988ae10614dabf137b32e8e891fbd98861b14916dc0be776d4184d173546e`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/06.stdout.json)。

<a id="case-07"></a>

## 07 · 删除指定完整段落

**自然语言效果 / 人工粗审**：三段“保留甲、删除我、保留乙”删除中间段后，仅保留甲、乙两段。

**前置条件**：`fixtures/07.docx`：恰好三段：保留甲 / 删除我 / 保留乙。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/07.docx"
    }
  },
  "steps": [
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "delete",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "range",
          "range": {
            "$ref": {
              "step": "read",
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
          "kind": "delete"
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
        "outputPath": "C:/wps-word-functional/outputs/07.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 read：paragraphs[*].text == ["保留甲\n","删除我\n","保留乙"]，truncated == false；delete：matchedCount == 1，len(ranges) == 1，符合 R；删除后的范围允许 start == end。

**预期 Task Response**：通用 T；`steps[*].id == ["read", "delete"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-1f7c7f2e56a38fdf80642a8a85bee09b8f2674180dc3179bd936af6a3dec19b4`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/07.stdout.json)。

<a id="case-08"></a>

## 08 · 插入表格

**自然语言效果 / 人工粗审**：“进度表”下面是一张 4 行 3 列的表格，首行为表头；表格下方有独立正文“以上为本周进度。”。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
            "kind": "paragraph",
            "runs": [
              {
                "text": "进度表"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "table",
      "address": {
        "app": "word",
        "action": "insertTable"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "data": [
          [
            "项目",
            "负责人",
            "状态"
          ],
          [
            "设计",
            "张三",
            "完成"
          ],
          [
            "开发",
            "李四",
            "进行中"
          ],
          [
            "测试",
            "王五",
            "待开始"
          ]
        ],
        "headerRow": true
      }
    },
    {
      "id": "after",
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
                "text": "以上为本周进度。"
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
        "outputPath": "C:/wps-word-functional/outputs/08.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 title、after：符合 W；table：data.table.rowCount == 4、columnCount == 3、headerRow == true、data 与请求矩阵完全相等，range 符合 R。表头不额外要求请求未指定的颜色或加粗。

**预期 Task Response**：通用 T；`steps[*].id == ["title", "table", "after"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-0763bb08549cd327eecbbc2cacebac81a6154ab3463f56b15c58afaba8bf9f68`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/08.stdout.json)。

<a id="case-09"></a>

## 09 · 插入图片

**自然语言效果 / 人工粗审**：文字“图片示例”之后嵌入蓝底白字 TEST 图片，宽 2 英寸、高 1 英寸，替代文本为“蓝底白字 TEST 图片”。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "intro",
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
                "text": "图片示例"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image",
      "address": {
        "app": "word",
        "action": "insertImage"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "source": {
          "kind": "file",
          "path": "C:/wps-word-functional/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "box",
          "width": {
            "value": 144,
            "unit": "pt"
          },
          "height": {
            "value": 72,
            "unit": "pt"
          },
          "fit": "stretch"
        },
        "alternativeText": {
          "kind": "description",
          "text": "蓝底白字 TEST 图片"
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
        "outputPath": "C:/wps-word-functional/outputs/09.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 intro：符合 W；image：image.kind == "inline"、embedded == true、source.mediaType == "image/png"；source.sha256、byteLength 等于已固定素材的哈希和字节数；size.width == {value:144,unit:"pt"}、height == {value:72,unit:"pt"}（数值容差 0.1 pt）；alternativeText 与请求一致。

**预期 Task Response**：通用 T；`steps[*].id == ["intro", "image"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-ac760cbe79d9291ec4681ec56d77a8cae2aaed045cf488d4530f2e67fcefe983`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/09.stdout.json)。

<a id="case-10"></a>

## 10 · 页面方向与页边距

**自然语言效果 / 人工粗审**：页面横向，四边页边距均为 1 英寸；正文为“横向页面样张”。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "write",
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
                "text": "横向页面样张"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "layout",
      "address": {
        "app": "word",
        "action": "setPageLayout"
      },
      "params": {
        "sections": {
          "kind": "all",
          "revision": {
            "$ref": {
              "step": "read",
              "path": [
                "data",
                "revision"
              ]
            }
          }
        },
        "layout": {
          "orientation": "landscape",
          "margins": {
            "top": {
              "value": 72,
              "unit": "pt"
            },
            "right": {
              "value": 72,
              "unit": "pt"
            },
            "bottom": {
              "value": 72,
              "unit": "pt"
            },
            "left": {
              "value": 72,
              "unit": "pt"
            }
          }
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
        "outputPath": "C:/wps-word-functional/outputs/10.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 write：符合 W；read：structure.sectionCount == 1，paragraphs[0].text == "横向页面样张"；layout：selectedSectionCount == 1，sections[0].index == 0，layout.orientation == "landscape"，四边 margins 的 unit == "pt"、value == 72（容差 0.1 pt）。

**预期 Task Response**：通用 T；`steps[*].id == ["write", "read", "layout"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-489b0287321bf939ff43fc498ea069d54ff81cc1e36108c0f61b8a2403d62009`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/10.stdout.json)。

<a id="case-11"></a>

## 11 · 页眉页脚

**自然语言效果 / 人工粗审**：两页文档，第一页和第二页各有相应正文；每页页眉“项目报告”、页脚“内部资料”。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "first",
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
                "text": "第一页正文"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "break",
      "address": {
        "app": "word",
        "action": "insertBreak"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "type": "page"
      }
    },
    {
      "id": "second",
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
                "text": "第二页正文"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "header_footer",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "all",
          "revision": {
            "$ref": {
              "step": "read",
              "path": [
                "data",
                "revision"
              ]
            }
          }
        },
        "updates": [
          {
            "area": "header",
            "variant": "primary",
            "operation": {
              "kind": "replace",
              "text": "项目报告"
            }
          },
          {
            "area": "footer",
            "variant": "primary",
            "operation": {
              "kind": "replace",
              "text": "内部资料"
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
        "outputPath": "C:/wps-word-functional/outputs/11.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 first、second：符合 W；break：符合 B（分页）；read：structure.sectionCount == 1；header_footer：selectedSectionCount == 1，stories 中 sectionIndex == 0 的 primary header/footer 均 exists == true、variantEnabled == true，text 分别精确等于“项目报告”和“内部资料”。

**预期 Task Response**：通用 T；`steps[*].id == ["first", "break", "second", "read", "header_footer"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-cfd98c1924c81f5fdae9ff22e98839b2c8a6d5fb2377fd011a03944fb9f80f72`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/11.stdout.json)。

<a id="case-12"></a>

## 12 · 插入分页符

**自然语言效果 / 人工粗审**：“第一部分”在第一页，“第二部分”从下一页开始。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "first",
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
                "text": "第一部分"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "break",
      "address": {
        "app": "word",
        "action": "insertBreak"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "type": "page"
      }
    },
    {
      "id": "second",
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
                "text": "第二部分"
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
        "outputPath": "C:/wps-word-functional/outputs/12.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 first、second：符合 W；break：符合 B。

**预期 Task Response**：通用 T；`steps[*].id == ["first", "break", "second"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-6b53852aadda5b59c01e30b98af1dc841338e6f128c32b99df60c7a2a91fdb81`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/12.stdout.json)。

<a id="case-13"></a>

## 13 · 分节与第二节横向

**自然语言效果 / 人工粗审**：第一节纵向；第二节另起一页且横向，四边页边距为 1 英寸。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
      "id": "first",
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
                "text": "第一节：纵向"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "break",
      "address": {
        "app": "word",
        "action": "insertBreak"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "type": "sectionNextPage"
      }
    },
    {
      "id": "second",
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
                "text": "第二节：横向"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "layout",
      "address": {
        "app": "word",
        "action": "setPageLayout"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            1
          ],
          "revision": {
            "$ref": {
              "step": "read",
              "path": [
                "data",
                "revision"
              ]
            }
          }
        },
        "layout": {
          "orientation": "landscape",
          "margins": {
            "top": {
              "value": 72,
              "unit": "pt"
            },
            "right": {
              "value": 72,
              "unit": "pt"
            },
            "bottom": {
              "value": 72,
              "unit": "pt"
            },
            "left": {
              "value": 72,
              "unit": "pt"
            }
          }
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
        "outputPath": "C:/wps-word-functional/outputs/13.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 first、second：符合 W；break：break.type == "sectionNextPage"，sectionCountBefore == 1、sectionCountAfter == 2、followingSectionIndex == 1；read：structure.sectionCount == 2，sections[0].layout.orientation == "portrait"；layout：selectedSectionCount == 1，仅返回 index == 1，方向和页边距同 10。

**预期 Task Response**：通用 T；`steps[*].id == ["first", "break", "second", "read", "layout"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-8dfed0b43671808c8b81e792474777a329c66f7b3410a267e959f33a3813538b`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/13.stdout.json)。

<a id="case-14"></a>

## 14 · 编辑已有文件并保存

**自然语言效果 / 人工粗审**：原文件“版本：旧版本”变成“版本：新版本”，后面增加“新增说明”，保存到原路径。

**前置条件**：`fixtures/14.docx`：恰好一段：版本：旧版本。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/14.docx"
    }
  },
  "steps": [
    {
      "id": "replace",
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
            "text": "旧版本",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "新版本"
            }
          ]
        }
      }
    },
    {
      "id": "append",
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
                "text": "新增说明"
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
        "action": "save"
      },
      "params": {}
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 replace：matchedCount == 1，len(ranges) == 1，符合 R；append：符合 W；save：artifact.path 等于本例输入路径、format == "docx"、sizeBytes > 0，documentState.persistenceState == "saved"。

**预期 Task Response**：通用 T；`steps[*].id == ["replace", "append"]`；completion.save 成功，Action 为 save；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-368b9eaa174e1248cb6c218fd20d50bfb72abab92db1675a4cd49919309e74e5`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/14.stdout.json)。

<a id="case-15"></a>

## 15 · 编辑已有文件后另存

**自然语言效果 / 人工粗审**：生成新的 15.docx，正文为“版本：新版本”；原输入文件仍是“版本：旧版本”。

**前置条件**：`fixtures/15.docx`：恰好一段：版本：旧版本。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/15.docx"
    }
  },
  "steps": [
    {
      "id": "replace",
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
            "text": "旧版本",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "新版本"
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
        "outputPath": "C:/wps-word-functional/outputs/15.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 replace：matchedCount == 1，len(ranges) == 1，符合 R；saveAs：符合 S；原文件未改写的效果信任 Action 契约，本轮不额外做磁盘哈希断言。

**预期 Task Response**：通用 T；`steps[*].id == ["replace"]`；completion.save 成功并满足 S；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-ebce8ecf356f25dc79bb1b879f5d0344b3f817727d4f848744aee20fa464faa8`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/15.stdout.json)。

<a id="case-16"></a>

## 16 · 编辑已有文件，不保存

**自然语言效果 / 人工粗审**：WPS 打开的文档显示“版本：新版本”，处于有未保存修改的状态；磁盘原文件仍是旧版本。

**前置条件**：`fixtures/16.docx`：恰好一段：版本：旧版本。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/16.docx"
    }
  },
  "steps": [
    {
      "id": "replace",
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
            "text": "旧版本",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "新版本"
            }
          ]
        }
      }
    }
  ],
  "completion": []
}
```

**预期 Action Response**：通用 A，以及 replace：matchedCount == 1，len(ranges) == 1，符合 R；不增加额外 inspect，replaceContent 响应不含 documentState，不对不存在的字段作断言。

**预期 Task Response**：通用 T；`steps[*].id == ["replace"]`；completion.save == null；completion.pdf == null。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-0af3759cafed2eb2cb98262a58c9b16d52308853ee5b68c7a226b4fc366c8fdd`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/16.stdout.json)。

<a id="case-17"></a>

## 17 · 修改后仅导出 PDF

**自然语言效果 / 人工粗审**：17.pdf 内容为“版本：新版本”；WPS 文档保留未保存修改，磁盘 DOCX 仍是旧版本。

**前置条件**：`fixtures/17.docx`：恰好一段：版本：旧版本。

**输入 JSON**：

```json
{
  "app": "word",
  "document": {
    "address": {
      "app": "word",
      "action": "openDocument"
    },
    "params": {
      "path": "C:/wps-word-functional/fixtures/17.docx"
    }
  },
  "steps": [
    {
      "id": "replace",
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
            "text": "旧版本",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "新版本"
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
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-functional/outputs/17.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 replace：matchedCount == 1，len(ranges) == 1，符合 R；exportPdf：符合 P，documentStateBefore.persistenceState == "modified"，documentStateAfter 与 Before 相等。

**预期 Task Response**：通用 T；`steps[*].id == ["replace"]`；completion.save == null；completion.pdf 成功并满足 P。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-41a5ccb07d60fb122ddbc03979c97d411dbc0521f00e0901798ad9ca3891a073`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/17.stdout.json)。

<a id="case-18"></a>

## 18 · 完整报告：DOCX 与 PDF

**自然语言效果 / 人工粗审**：一份两页项目周报：第一页标题、正文和进度表；第二页“附图说明”和 TEST 图片；两页都有页眉页脚。同时交付 DOCX 和 PDF。

**前置条件**：新建空白文档；不依赖其他用例输出。

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
                "text": "项目周报"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "本周进度如下。"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "table",
      "address": {
        "app": "word",
        "action": "insertTable"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "data": [
          [
            "项目",
            "负责人",
            "状态"
          ],
          [
            "设计",
            "张三",
            "完成"
          ],
          [
            "开发",
            "李四",
            "进行中"
          ],
          [
            "测试",
            "王五",
            "待开始"
          ]
        ],
        "headerRow": true
      }
    },
    {
      "id": "break",
      "address": {
        "app": "word",
        "action": "insertBreak"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "type": "page"
      }
    },
    {
      "id": "second",
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
                "text": "附图说明"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image",
      "address": {
        "app": "word",
        "action": "insertImage"
      },
      "params": {
        "anchor": {
          "kind": "documentEnd"
        },
        "source": {
          "kind": "file",
          "path": "C:/wps-word-functional/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "box",
          "width": {
            "value": 144,
            "unit": "pt"
          },
          "height": {
            "value": 72,
            "unit": "pt"
          },
          "fit": "stretch"
        },
        "alternativeText": {
          "kind": "description",
          "text": "蓝底白字 TEST 图片"
        }
      }
    },
    {
      "id": "read",
      "address": {
        "app": "word",
        "action": "inspectDocument"
      },
      "params": {
        "scope": {
          "kind": "document"
        },
        "limits": {
          "maxTextCharacters": 4096,
          "maxParagraphs": 128,
          "maxRuns": 512
        }
      }
    },
    {
      "id": "header_footer",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "all",
          "revision": {
            "$ref": {
              "step": "read",
              "path": [
                "data",
                "revision"
              ]
            }
          }
        },
        "updates": [
          {
            "area": "header",
            "variant": "primary",
            "operation": {
              "kind": "replace",
              "text": "项目报告"
            }
          },
          {
            "area": "footer",
            "variant": "primary",
            "operation": {
              "kind": "replace",
              "text": "内部资料"
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
        "outputPath": "C:/wps-word-functional/outputs/18.docx",
        "overwritePolicy": "failIfExists"
      }
    },
    {
      "address": {
        "app": "word",
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-functional/outputs/18.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及 title、second：符合 W；table：同 08；break：符合 B；image：同 09；read：structure.tableCount == 1、inlineImageCount == 1、sectionCount == 1、pageBreakCount == 1；header_footer：同 11；saveAs 符合 S，exportPdf 符合 P，导出前后 persistenceState 均为 "saved"。

**预期 Task Response**：通用 T；`steps[*].id == ["title", "table", "break", "second", "image", "read", "header_footer"]`；completion.save 成功并满足 S；completion.pdf 成功并满足 P。

**执行记录**：本轮响应结果＝通过；人工粗审＝未审阅；Task ID＝`task-91b4f34677ab0d49e81d1238dbc87e1a9ec5cd1e5410b114ceeefe2fcc58bb60`；[完整响应](../../../build/evidence/word-functional-rerun-20260917-225604/windows/responses/18.stdout.json)。

## 本清单检查情况

18 份 JSON 已通过现行请求编排与 Action 参数静态预检（含前序结果引用的部分参数校验）。此检查不启动 Task Executor 或 WPS，不属于实机通过证据。后续已完成 Windows 首轮执行，素材、请求、响应与输出均已归档；静态检查不替代各例实机结果。

## 实机期间修正的清单问题

- `failIfExists` 不应期待 `outputResolution`；该字段仅属于 `renameIfExists`。已修正 S/P，原始响应未改写。
- `inspectDocument.paragraphs[*].text` 中非末段包含规范化段落换行 `\n`；已修正 03/04/06/07 的精确预期，没有通过删字符来掩盖差异。
- 本轮简化 DOCX 素材打开后报告 modified，未满足已保存素材的前置条件。原生素材对照在首次保存时因另一个打开文档同名而停止，尚未得到有效对照结论。后续素材、产物的文件名需要各自独立，不能只依赖目录隔离。

## 后续复现更正

已定位 saved → modified 发生在 openDocument 的 Get-DocumentFingerprint 执行期间，刚打开文件时为 saved。首轮状态不符不能归因于素材未由 WPS 保存；见 [复现记录](WPS_WORD_OPEN_STATE_REPRO.md)。首轮原始记录已归档；指纹读取问题已修复，本轮 18 项响应均通过。


## 指纹修复后全量重跑

2026-09-17 22:56，使用独立素材和唯一文件名重跑全部 18 项，每项提交一次；Action 与 Task 响应全部满足原预期。原首轮清单和结果归档在 `build/evidence/word-functional-rerun-20260917-225604/`，保留当时失败证据。Windows 本轮目录为 `C:/Users/yim/wps-functional-rerun-20260917-225604`；文件名增加轮次及 fixtures/outputs 前缀，JSON 业务操作未变。


## 进阶测试单

多节、多变体页眉页脚、连续范围定位及复杂内容组合见 [进阶 19–26 例](WPS_WORD_ADVANCED_TEST_CHECKLIST.md)。该组已完成首轮实机：6 例通过、2 例响应值差异；见 [进阶结果](WPS_WORD_ADVANCED_TEST_RESULTS.md)，与以上 18 例统计分开。
