# 公式和计算

`setFormulas.formulas` 是与矩形严格匹配的二维公式数组，每项以 `=` 开始，使用 WPS COM `Formula` 接受的英文 A1 形式。不要使用 `FormulaLocal` 或依赖区域设置替换函数名及分隔符。

限定同一工作表的单元格引用、数值/文本/布尔常量、算术与比较，以及 SUM、AVERAGE、COUNT、MIN、MAX、IF、IFERROR、ROUND、ABS；还支持 COUNTIF、COUNTIFS、SUMIF、SUMIFS、COUNTA、COUNTBLANK、AND、OR、NOT、LEFT、RIGHT、MID、LEN、TRIM、UPPER、LOWER、CONCATENATE、VLOOKUP、HLOOKUP、INDEX、MATCH、DATE、YEAR、MONTH、DAY、ROUNDUP、ROUNDDOWN、TEXT、VALUE、SUBSTITUTE、SEARCH、FIND。不支持名称、外部工作簿、跨工作表引用、DDE、链接、宏函数或动态数组公式。

先读取区域获得 token，写公式后检查返回的 formula；WPS 若改写公式且无法与请求一致验证，停止并检查，不能宣称写入成功。再使用写入响应的新 token 调用 `calculateRange`，检查实际 value 和 errorCode。公式存在与结果正确是两个条件。

`calculateRange` 只请求指定区域的计算，不承诺任意跨区域依赖链都已完整刷新。依赖链复杂或计算模式不明确时，先缩小支持范围，不能把旧缓存值报告为最终结果。`writeRange` 用于字面文本和数值，即便文本以 `=` 开头也不应产生公式。
