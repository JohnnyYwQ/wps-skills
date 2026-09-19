# Word Skill 进阶功能测试清单（19–26）

日期：2026-09-17。状态：已完成 Windows 首轮执行；6 例满足全部预期，20/24 两例响应值存在差异；8 个 Task 及其 Action 均返回成功。人工粗审待进行。详见 [进阶结果](WPS_WORD_ADVANCED_TEST_RESULTS.md)。

延续 [基础 18 例](WPS_WORD_FUNCTIONAL_TEST_CHECKLIST.md)：固定 JSON → 正式 Task Client → Action Response / Task Response。每项另列自然语言效果供人工粗审。仅增加正常操作组合复杂度，不测 Agent、Task 重试、恢复、并发、取消或故障注入。

## 准备与提交

- 使用包含最新指纹修复的独立 Word 包。当前基线 `build/archive/runs/word-fingerprint-fix-20260917-223921/wps-word`。
- `C:/wps-word-advanced` 为示例根目录。执行前部署到全新 Windows 绝对目录，并同时给所有素材和产物文件名加本轮唯一前缀；目录不同但同名仍可能与 WPS 已打开文档冲突。只替换路径，不改变业务 JSON。
- 创建 fixtures、requests、outputs、responses。8 例互相独立，每例提交一次，不自动重试。失败记录具体 step.id、Action 名称、response.error.code/message、Task stop 和未执行动作，保留现场。
- sample.png 沿用基础清单的 200×100 蓝底白字 TEST 素材；SHA-256 为 `76297a773d85d4d6a6d99d4197041deaf9c8d7a190e86eb54c63ce6fa848c2e8`，525 字节。可以从 `build/evidence/word-functional-rerun-20260917-225604/windows/fixtures/sample.png` 复制。图片未包含在本请求目录中。
- 第 25 例复杂素材已在本轮单独创建并保存，复制为独立输入文件，initial 响应已验证约定结构。准备素材不算该用例通过；不修改用户原文件。其他 7 例均从新建文档开始。
- 新建基线为无页眉页脚、普通空白纵向文档；宋体和 Times New Roman 等所需字体可用。准备工作不属于动作重试。
- 本文每份 JSON 同步单独保存于 `src/test/resources/wps_skills/word/advanced/requests/19.json` 至 `26.json`；只有路径部署替换后才能提交。

```text
python "<skill-dir>/scripts/word.py" --app word --task-file "<本轮目录>/requests/19.json"
```

## 通用成功条件

**A · Action**：每项 state=succeeded；response.outcome=succeeded、无 error；address、taskId 与输入/外层一致，data 符合对应正式 result schema。新建返回 unsaved/readOnly=false；25 打开返回 saved/readOnly=false、正确 DOCX 路径和 sizeBytes>0。revision 非空，按相等关系判断，不把它当序号。

**W/R · 变更**：writeContent 的 range、replaceContent 的 ranges、insertTable 的 table.range 均使用 revisionAfter，坐标合法；非空插入/替换 start<end，删除允许相等。revisionBefore/After 均非空；这些有实际内容/结构变化的操作要求不相等。若两次变更间只有只读动作，后者 revisionBefore 应等于最近变更 revisionAfter。格式效果信任 Action 内部回读，不要求变更响应返回未定义的格式快照。

**F · 查找**：matches 数量和 text 符合本例，各 range 使用 find 返回的 revision，按 start 升序且不重叠；truncated=false、remainingRange=null。连续只读步骤 revision 相同。内容改变后需要定位时重新查找/检查，不能沿用旧 range。

**I · 图片**：kind=inline、embedded=true；source.mediaType=image/png、sha256/byteLength 等于固定素材，alternativeText 与请求一致；宽高单位 pt，数值容差 0.1 pt；image.range 合法、使用 revisionAfter。

**H · 页眉页脚**：返回 story 的 sectionIndex/area/variant 与请求组合完全对应，无遗漏或重复；replace/clear 断开该 story 的继承，linkToPrevious 保持相应链接；文字、variantEnabled 和本例明确的字段符合预期。首页/偶数页变体由对应更新启用，不引入 contract 不支持的独立开关参数。

**S/P · 持久化**：saveAs 返回本例 outputPath、docx、sizeBytes>0、saved；exportPdf 返回本例 PDF 路径、pdf、sizeBytes>0。二者 replacedExisting=false，failIfExists 不包含 outputResolution。PDF 前后 revision 和 documentState 相同，本轮都在保存之后导出，故均为 saved。原文件保持不变依赖 saveAs 契约，不额外做磁盘内容比对。

**T · Task**：type=task.response、app=word、state=completed、outcome=succeeded、stop=null。document.id=task_document；steps 的数量、顺序、id、address 与请求完全相同。completion.save.id=task_save、Action=saveAs；选择 PDF 的 completion.pdf.id=task_pdf，未选择严格为 null。document、每个 step 和 completion 均需检查 A，不能仅检查外层 outcome。cleanup 和 taskFile 诊断原样留存，异常单独记录。

inspectDocument 是场景中获取新范围/revision 或读取既有结构的业务动作；不为验收增加终态文档解析或截图识别。较复杂内容的页面数和排版由用户粗审，不把分节符数量直接当成渲染页数。

## 用例索引

| 编号 | 用例 | 响应结果 | 人工粗审 |
| --- | --- | --- | --- |
| 19 | [三节布局与连续 revision 引用](#case-19) | 通过 | 未审阅 |
| 20 | [三节六变体页眉页脚与继承](#case-20) | 响应值与清单预期不符 | 未审阅 |
| 21 | [重复查找、变长替换与相对插入](#case-21) | 通过 | 未审阅 |
| 22 | [混合格式段落替换与删除](#case-22) | 通过 | 未审阅 |
| 23 | [多张表格与正文连续插入](#case-23) | 通过 | 未审阅 |
| 24 | [多图片尺寸模式与分页混排](#case-24) | 响应值与清单预期不符 | 未审阅 |
| 25 | [编辑独立复杂文档并另存](#case-25) | 通过 | 未审阅 |
| 26 | [三节综合报告与较长 Task](#case-26) | 通过 | 未审阅 |

<a id="case-19"></a>

## 19 · 三节布局与连续 revision 引用

**自然语言效果 / 人工粗审**：三节分别为纵向、横向、纵向；第二节页边距 36 pt，第三节 54 pt；各节另起一页。

**前置条件**：新建空白纵向文档。

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
      "id": "s1",
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
                "text": "第一节：概览"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "b1",
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
      "id": "s2",
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
                "text": "第二节：宽表区"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "b2",
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
      "id": "s3",
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
                "text": "第三节：说明"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "before",
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
      "id": "layout2",
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
              "step": "before",
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
              "value": 36,
              "unit": "pt"
            },
            "right": {
              "value": 36,
              "unit": "pt"
            },
            "bottom": {
              "value": 36,
              "unit": "pt"
            },
            "left": {
              "value": 36,
              "unit": "pt"
            }
          }
        }
      }
    },
    {
      "id": "middle",
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
      "id": "layout3",
      "address": {
        "app": "word",
        "action": "setPageLayout"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            2
          ],
          "revision": {
            "$ref": {
              "step": "middle",
              "path": [
                "data",
                "revision"
              ]
            }
          }
        },
        "layout": {
          "orientation": "portrait",
          "margins": {
            "top": {
              "value": 54,
              "unit": "pt"
            },
            "right": {
              "value": 54,
              "unit": "pt"
            },
            "bottom": {
              "value": 54,
              "unit": "pt"
            },
            "left": {
              "value": 54,
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-19.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- s1/s2/s3 满足 W；b1 的 sectionCountBefore/After 为 1/2，followingSectionIndex=1；b2 为 2/3，followingSectionIndex=2。
- before.structure.sectionCount=3，三个 section 的 orientation 均为 portrait。
- layout2.selectedSectionCount=1，仅返回 index=1；方向 landscape，四边均 36 pt。
- middle.revision=layout2.revisionAfter，middle 的第 0/2 节 layout 与 before 相同，第 1 节与 layout2 返回值相同。
- layout3.selectedSectionCount=1，仅返回 index=2；方向 portrait，四边均 54 pt；其 revisionBefore=middle.revision。

**预期 Task Response**：通用 T；`steps[*].id == ["s1", "b1", "s2", "b2", "s3", "before", "layout2", "middle", "layout3"]`；completion.save 成功；completion.pdf 严格为 null。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-f0bb89147b3da479839f8296e5525fdd2e089930bb939c4af7ea1461783c751a`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/19.stdout.json)；人工粗审＝未审阅。

<a id="case-20"></a>

## 20 · 三节六变体页眉页脚与继承

**自然语言效果 / 人工粗审**：三节各有三页正文。首页、偶数页、普通奇数页显示对应 header/footer 变体文字；第二节继承第一节。第三节普通页眉为“第三节独立页眉”，偶数页页脚为空，其余继承。

**前置条件**：新建文档；不要求页码重置。根据实际页码奇偶查看变体，不把“本节第2页”一律当偶数页。

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
      "id": "text1_1",
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
                "text": "第1节第1页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page1_2",
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
      "id": "text1_2",
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
                "text": "第1节第2页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page1_3",
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
      "id": "text1_3",
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
                "text": "第1节第3页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "section2",
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
      "id": "text2_1",
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
                "text": "第2节第1页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page2_2",
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
      "id": "text2_2",
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
                "text": "第2节第2页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page2_3",
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
      "id": "text2_3",
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
                "text": "第2节第3页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "section3",
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
      "id": "text3_1",
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
                "text": "第3节第1页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page3_2",
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
      "id": "text3_2",
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
                "text": "第3节第2页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "page3_3",
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
      "id": "text3_3",
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
                "text": "第3节第3页"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "base",
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
      "id": "base_stories",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            0
          ],
          "revision": {
            "$ref": {
              "step": "base",
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
              "text": "header-primary"
            }
          },
          {
            "area": "header",
            "variant": "firstPage",
            "operation": {
              "kind": "replace",
              "text": "header-firstPage"
            }
          },
          {
            "area": "header",
            "variant": "evenPages",
            "operation": {
              "kind": "replace",
              "text": "header-evenPages"
            }
          },
          {
            "area": "footer",
            "variant": "primary",
            "operation": {
              "kind": "replace",
              "text": "footer-primary"
            }
          },
          {
            "area": "footer",
            "variant": "firstPage",
            "operation": {
              "kind": "replace",
              "text": "footer-firstPage"
            }
          },
          {
            "area": "footer",
            "variant": "evenPages",
            "operation": {
              "kind": "replace",
              "text": "footer-evenPages"
            }
          }
        ]
      }
    },
    {
      "id": "r1",
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
      "id": "inherit2",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            1
          ],
          "revision": {
            "$ref": {
              "step": "r1",
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
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "header",
            "variant": "firstPage",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "header",
            "variant": "evenPages",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "primary",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "firstPage",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "evenPages",
            "operation": {
              "kind": "linkToPrevious"
            }
          }
        ]
      }
    },
    {
      "id": "r2",
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
      "id": "inherit3",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            2
          ],
          "revision": {
            "$ref": {
              "step": "r2",
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
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "header",
            "variant": "firstPage",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "header",
            "variant": "evenPages",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "primary",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "firstPage",
            "operation": {
              "kind": "linkToPrevious"
            }
          },
          {
            "area": "footer",
            "variant": "evenPages",
            "operation": {
              "kind": "linkToPrevious"
            }
          }
        ]
      }
    },
    {
      "id": "r3",
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
      "id": "override3",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            2
          ],
          "revision": {
            "$ref": {
              "step": "r3",
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
              "text": "第三节独立页眉"
            }
          },
          {
            "area": "footer",
            "variant": "evenPages",
            "operation": {
              "kind": "clear"
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-20.docx",
        "overwritePolicy": "failIfExists"
      }
    },
    {
      "address": {
        "app": "word",
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-20.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- 正文写入满足 W；base.structure.sectionCount=3、pageBreakCount=6、sectionBreakCount=2。所有分页符保持当时节数；两个分节符分别由 1→2、2→3。
- base_stories 返回六条 sectionIndex=0 的 story，exists=true、variantEnabled=true、linkToPrevious=false，text 等于各 update.text。
- r1 的第 0 节 firstPageEnabled/evenPagesEnabled 均为 true。inherit2 返回六条 sectionIndex=1 的 story，variantEnabled=true、linkToPrevious=true，text 与第一节对应变体相同。
- inherit3 同理，sectionIndex=2，继承第二节；r3 的三节 flags 均为 true，每节六变体文字一致。
- override3 只返回两条 story：primary/header 的 text=第三节独立页眉，linkToPrevious=false；evenPages/footer 的 text=""，linkToPrevious=false；二者 variantEnabled=true。clear 不额外写死 exists 的实现相关值。

**预期 Task Response**：通用 T；`steps[*].id == ["text1_1", "page1_2", "text1_2", "page1_3", "text1_3", "section2", "text2_1", "page2_2", "text2_2", "page2_3", "text2_3", "section3", "text3_1", "page3_2", "text3_2", "page3_3", "text3_3", "base", "base_stories", "r1", "inherit2", "r2", "inherit3", "r3", "override3"]`；completion.save 成功；completion.pdf 成功。

**执行记录**：响应结果＝响应值与清单预期不符；Action 错误＝无；差异＝base / inspectDocument：pageBreakCount 预期 6，实际 8；Task ID＝`task-47f6d3aa7c0583561d818f70c95f9854a419b02cb93ab0dfc7da5bd7c09d015c`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/20.stdout.json)；人工粗审＝未审阅。

<a id="case-21"></a>

## 21 · 重复查找、变长替换与相对插入

**自然语言效果 / 人工粗审**：“订单A：待处理”“订单B：待处理”均变成“已完成并复核”；“结束标记”前增加“复核说明：”。

**前置条件**：新建空白文档。

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
                "text": "订单A：待处理"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "订单B：待处理"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "结束标记"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "find_old",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "待处理",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "replace_all",
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
            "text": "待处理",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 2
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "已完成并复核"
            }
          ]
        }
      }
    },
    {
      "id": "find_new",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "已完成并复核",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "find_anchor",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "结束标记",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
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
              "step": "find_anchor",
              "path": [
                "data",
                "matches",
                0,
                "range"
              ]
            }
          }
        },
        "blocks": [
          {
            "kind": "text",
            "runs": [
              {
                "text": "复核说明："
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-21.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- find_old.matches 恰好 2 项，text 均为待处理，范围升序、不重叠。
- replace_all.matchedCount=2，ranges 长度=2，均满足 R。
- find_new.matches 恰好 2 项，text 均为已完成并复核；revision=replace_all.revisionAfter。
- find_anchor 恰好命中结束标记一次；revision=find_new.revision。
- insert 满足 W，range.start=find_anchor.matches[0].range.start，revisionBefore=find_anchor.revision；不复用替换前的范围。

**预期 Task Response**：通用 T；`steps[*].id == ["seed", "find_old", "replace_all", "find_new", "find_anchor", "insert"]`；completion.save 成功；completion.pdf 严格为 null。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-318219754da2978d516801cd78035159b4b4b1a53f469550053d6f146f3e8a17`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/21.stdout.json)；人工粗审＝未审阅。

<a id="case-22"></a>

## 22 · 混合格式段落替换与删除

**自然语言效果 / 人工粗审**：保留“开头”和“结尾”；中间“旧说明”改为居中加粗的“新说明”；“删除我”整段消失。

**前置条件**：新建文档。

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
                "text": "开头"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "旧",
                "format": {
                  "italic": true
                }
              },
              {
                "text": "说明",
                "format": {
                  "color": "#1F4E79"
                }
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "删除我"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "结尾"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "before",
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
              "step": "before",
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
                "alignment": "center",
                "lineSpacing": {
                  "kind": "double"
                }
              }
            }
          ]
        }
      }
    },
    {
      "id": "updated",
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
      "id": "delete_para",
      "address": {
        "app": "word",
        "action": "replaceContent"
      },
      "params": {
        "target": {
          "kind": "range",
          "range": {
            "$ref": {
              "step": "updated",
              "path": [
                "data",
                "paragraphs",
                2,
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-22.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- before.paragraphs[*].text=["开头\n","旧说明\n","删除我\n","结尾"]，均 complete=true，truncated=false。
- replace_para.matchedCount=1，ranges 长度=1，满足 R；replacement 的格式由内部回读负责。
- updated.revision=replace_para.revisionAfter；paragraphs[*].text=["开头\n","新说明\n","删除我\n","结尾"]。
- delete_para.matchedCount=1，ranges 长度=1，允许 start=end；revisionBefore=updated.revision。

**预期 Task Response**：通用 T；`steps[*].id == ["seed", "before", "replace_para", "updated", "delete_para"]`；completion.save 成功；completion.pdf 严格为 null。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-59691fab942585f150148fe7c37ede274755cf4f86be6c86c0351644024d2995`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/22.stdout.json)；人工粗审＝未审阅。

<a id="case-23"></a>

## 23 · 多张表格与正文连续插入

**自然语言效果 / 人工粗审**：标题“月度汇总”下是三行三列进度表，接“阶段说明”，再是三行两列费用表，最后“汇总结束”。空单元格保留。

**前置条件**：新建空白文档。

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
                "text": "月度汇总"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "progress",
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
            "任务",
            "责任人",
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
            ""
          ]
        ],
        "headerRow": true
      }
    },
    {
      "id": "between",
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
                "text": "阶段说明"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "cost",
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
            "类别",
            "金额"
          ],
          [
            "开发",
            "1200"
          ],
          [
            "测试",
            "800"
          ]
        ],
        "headerRow": true
      }
    },
    {
      "id": "tail",
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
                "text": "汇总结束"
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-23.docx",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- title/between/tail 满足 W。
- progress.table.rowCount=3、columnCount=3；cost 为 3/2；两者 headerRow=true，data 与各自请求矩阵逐格相等，table.range 满足 R。
- 按步骤顺序追加，后续变更的 revisionBefore 与最近一次变更的 revisionAfter 相同；不比较跨 revision 的旧坐标作为最终位置。

**预期 Task Response**：通用 T；`steps[*].id == ["title", "progress", "between", "cost", "tail"]`；completion.save 成功；completion.pdf 严格为 null。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-d309c5fa261441ad4676b0d96075a8afdc5be7aba25e0b3eefee877fde4c5537`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/23.stdout.json)；人工粗审＝未审阅。

<a id="case-24"></a>

## 24 · 多图片尺寸模式与分页混排

**自然语言效果 / 人工粗审**：第一页“图一：宽度模式”后为 144×72 pt 图片；第二页依次为“图二：适应方框”和 100×50 pt 图片、“图三：拉伸”和 90×90 pt 图片。

**前置条件**：sample.png 固定为 200×100 像素蓝底白字 TEST，与基础清单同一素材。

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
      "id": "caption1",
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
                "text": "图一：宽度模式"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image1",
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
          "path": "C:/wps-word-advanced/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "width",
          "width": {
            "value": 2,
            "unit": "in"
          }
        },
        "alternativeText": {
          "kind": "description",
          "text": "图一 TEST"
        }
      }
    },
    {
      "id": "page",
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
      "id": "caption2",
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
                "text": "图二：适应方框"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image2",
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
          "path": "C:/wps-word-advanced/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "box",
          "width": {
            "value": 100,
            "unit": "pt"
          },
          "height": {
            "value": 100,
            "unit": "pt"
          },
          "fit": "contain"
        },
        "alternativeText": {
          "kind": "description",
          "text": "图二 TEST"
        }
      }
    },
    {
      "id": "caption3",
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
                "text": "图三：拉伸"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image3",
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
          "path": "C:/wps-word-advanced/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "box",
          "width": {
            "value": 90,
            "unit": "pt"
          },
          "height": {
            "value": 90,
            "unit": "pt"
          },
          "fit": "stretch"
        },
        "alternativeText": {
          "kind": "description",
          "text": "图三 TEST"
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-24.docx",
        "overwritePolicy": "failIfExists"
      }
    },
    {
      "address": {
        "app": "word",
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-24.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- caption1/2/3 满足 W；page.type=page、节数 1→1。
- 三个 image.kind=inline、embedded=true，source 的 sha256/byteLength 与固定素材相同；alternativeText 与各自请求一致。
- image1.size=144×72 pt，image2=100×50 pt，image3=90×90 pt；容差均 0.1 pt；image.range.revision=revisionAfter。

**预期 Task Response**：通用 T；`steps[*].id == ["caption1", "image1", "page", "caption2", "image2", "caption3", "image3"]`；completion.save 成功；completion.pdf 成功。

**执行记录**：响应结果＝响应值与清单预期不符；Action 错误＝无；差异＝image2 / insertImage：预期 100×50 pt（±0.1），实际 99.75×49.5 pt；Task ID＝`task-f4d5be9867bbeb2c08d8e23f3df727e68dc864c554e937eb2ad76f05acb0a291`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/24.stdout.json)；人工粗审＝未审阅。

<a id="case-25"></a>

## 25 · 编辑独立复杂文档并另存

**自然语言效果 / 人工粗审**：在已有两节报告中，将唯一“状态：草稿”变为“状态：定稿”，追加“复核完成”；第二节变横向、页边距 54 pt；表格、图片及既有页眉页脚保留。另存 DOCX 并导出 PDF。

**前置条件**：独立、已保存、可写的 advanced-input-25.docx：第一节含一段“状态：草稿”、一张 2×2 表格（项目/值；版本/1）、一张 sample.png 内嵌图片；第二节含“附录”。两节初始纵向，primary 页眉“存量报告”、页脚“内部”，第二节链接第一节；其余变体未启用。正文中“草稿”只出现一次；不得使用本轮其他用例的产物。

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
      "path": "C:/wps-word-advanced/fixtures/advanced-input-25.docx"
    }
  },
  "steps": [
    {
      "id": "initial",
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
              "text": "定稿"
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
                "text": "复核完成"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "for_layout",
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
      "id": "appendix_layout",
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
              "step": "for_layout",
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
              "value": 54,
              "unit": "pt"
            },
            "right": {
              "value": 54,
              "unit": "pt"
            },
            "bottom": {
              "value": 54,
              "unit": "pt"
            },
            "left": {
              "value": 54,
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-25.docx",
        "overwritePolicy": "failIfExists"
      }
    },
    {
      "address": {
        "app": "word",
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-25.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- openDocument 返回 saved/readOnly=false；initial 的 state 和 revision 与 openDocument 相同。
- initial.structure：sectionCount=2、tableCount=1、inlineImageCount=1、floatingImageCount=0；两节方向 portrait；primary 页眉/脚文字符合素材约定。
- finalize.matchedCount=1，ranges 长度=1，满足 R；append 满足 W。
- for_layout.documentState.persistenceState=modified；revision=append.revisionAfter；tableCount/inlineImageCount/sectionCount 与 initial 相同；两节 headerFooter 与 initial 相同。
- appendix_layout 只返回 index=1、方向 landscape、四边 54 pt。另存后 saved；PDF 导出前后均 saved、revision 不变。

**预期 Task Response**：通用 T；`steps[*].id == ["initial", "finalize", "append", "for_layout", "appendix_layout"]`；completion.save 成功；completion.pdf 成功。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-facdecc8561da71781f3d60c962ac1caca4f33ab909a5cfd4da5914908ebaca9`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/25.stdout.json)；人工粗审＝未审阅。

<a id="case-26"></a>

## 26 · 三节综合报告与较长 Task

**自然语言效果 / 人工粗审**：第一节为“交付报告”、项目状态和进度表；第二节为横向图示，包含两张图片；第三节为纵向说明。状态更新为“已交付”，三节统一页眉“交付报告”、页脚“内部资料”，交付 DOCX/PDF。

**前置条件**：新建空白纵向文档；sample.png 与 24 相同。

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
      "id": "cover",
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
                "text": "交付报告",
                "format": {
                  "fontSizePt": 20,
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
                "text": "状态：待交付"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "progress",
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
            "任务",
            "责任人",
            "状态"
          ],
          [
            "开发",
            "李四",
            "完成"
          ],
          [
            "测试",
            "王五",
            "完成"
          ]
        ],
        "headerRow": true
      }
    },
    {
      "id": "section2",
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
      "id": "figures",
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
                "text": "图示"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image1",
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
          "path": "C:/wps-word-advanced/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "width",
          "width": {
            "value": 144,
            "unit": "pt"
          }
        },
        "alternativeText": {
          "kind": "description",
          "text": "示意图一"
        }
      }
    },
    {
      "id": "caption",
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
                "text": "对照图"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "image2",
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
          "path": "C:/wps-word-advanced/fixtures/sample.png"
        },
        "placement": {
          "kind": "inline"
        },
        "size": {
          "kind": "height",
          "height": {
            "value": 54,
            "unit": "pt"
          }
        },
        "alternativeText": {
          "kind": "description",
          "text": "示意图二"
        }
      }
    },
    {
      "id": "section3",
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
      "id": "notes",
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
                "text": "交付说明"
              }
            ]
          },
          {
            "kind": "paragraph",
            "runs": [
              {
                "text": "请按计划验收。"
              }
            ]
          }
        ]
      }
    },
    {
      "id": "find_pending",
      "address": {
        "app": "word",
        "action": "findContent"
      },
      "params": {
        "query": {
          "scope": {
            "kind": "document"
          },
          "text": "待交付",
          "caseSensitive": true,
          "wholeWord": false
        },
        "limit": 20
      }
    },
    {
      "id": "mark_delivered",
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
            "text": "待交付",
            "caseSensitive": true,
            "wholeWord": false
          },
          "expectedMatchCount": 1
        },
        "replacement": {
          "kind": "text",
          "runs": [
            {
              "text": "已交付"
            }
          ]
        }
      }
    },
    {
      "id": "for_layout",
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
      "id": "wide",
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
              "step": "for_layout",
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
              "value": 36,
              "unit": "pt"
            },
            "right": {
              "value": 36,
              "unit": "pt"
            },
            "bottom": {
              "value": 36,
              "unit": "pt"
            },
            "left": {
              "value": 36,
              "unit": "pt"
            }
          }
        }
      }
    },
    {
      "id": "for_headers",
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
      "id": "headers",
      "address": {
        "app": "word",
        "action": "setHeaderFooter"
      },
      "params": {
        "sections": {
          "kind": "indexes",
          "indexes": [
            0,
            1,
            2
          ],
          "revision": {
            "$ref": {
              "step": "for_headers",
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
              "text": "交付报告"
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
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-26.docx",
        "overwritePolicy": "failIfExists"
      }
    },
    {
      "address": {
        "app": "word",
        "action": "exportPdf"
      },
      "params": {
        "outputPath": "C:/wps-word-advanced/outputs/advanced-output-26.pdf",
        "overwritePolicy": "failIfExists"
      }
    }
  ]
}
```

**预期 Action Response**：通用 A，以及适用的 W/R/F/I/H/S/P；额外逐步断言：

- 所有文字和标题写入满足 W；progress.table 为 3×3、headerRow=true、data 与请求相同。
- 两个 sectionBreak 分别由 1→2、2→3；image1=144×72 pt、image2=108×54 pt，图片源及替代文字满足 I。
- find_pending 恰好命中待交付一次；mark_delivered.matchedCount=1、ranges 长度=1。
- for_layout.structure：sectionCount=3、sectionBreakCount=2、tableCount=1、inlineImageCount=2、floatingImageCount=0；revision=mark_delivered.revisionAfter。
- wide 只返回 index=1，landscape、四边 36 pt；for_headers.revision=wide.revisionAfter，第 0/2 节 layout 与 for_layout 一致。
- headers.selectedSectionCount=3，返回六条 story；三节 primary/header 的 text=交付报告、primary/footer 的 text=内部资料，variantEnabled/exists=true、linkToPrevious=false。
- 保存和 PDF 满足 S/P；导出前后 saved，revision 不变。

**预期 Task Response**：通用 T；`steps[*].id == ["cover", "progress", "section2", "figures", "image1", "caption", "image2", "section3", "notes", "find_pending", "mark_delivered", "for_layout", "wide", "for_headers", "headers"]`；completion.save 成功；completion.pdf 成功。

**执行记录**：响应结果＝通过；Action 错误＝无；差异＝无；Task ID＝`task-555d2b3105fd510329beb7bc80a11181ff84ce6a731bb601337aa8b5445821cf`；[完整响应](../../../build/evidence/word-advanced-20260917-231026/windows/responses/26.stdout.json)；人工粗审＝未审阅。

## 静态预检记录

8 份 JSON 通过现行 compile_request 与 Task Client _preflight，检查 Task 编排、Action 名称、参数和前序 $ref 使用。含动态引用的参数仅能静态部分校验；实际引用值、素材、WPS 行为及响应断言须在实机执行时确认。此记录不是实机成功证据。


## 本轮执行说明

Windows 目录为 `C:/Users/yim/wps-advanced-20260917-231026`。每例只提交一次，业务 JSON 未变，仅替换路径及唯一文件名前缀；素材准备为单独 Task，不计入 8 例。原始 20/24 断言保留，没有按实测值放宽为通过。
