/**
 * Input: Word 操作参数
 * Output: Word 操作结果
 * Pos: macOS 加载项 Word 处理器。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * Word操作处理器 - Mac版WPS加载项
 * 使用WPS JavaScript API实现Word文档操作
 * @author 老李（参考老王的PowerShell实现）
 */

/**
 * 获取当前活动文档信息
 */
function getActiveDocument(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        return {
            success: true,
            data: {
                name: doc.Name,
                path: doc.FullName,
                paragraphCount: doc.Paragraphs.Count,
                wordCount: doc.Words.Count,
                characterCount: doc.Characters.Count
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 读取文档文本内容
 */
function getDocumentText(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        var text = doc.Content.Text;
        var length = text.length;
        // 限制返回长度，防止内存爆炸
        if (length > 10000) {
            text = text.substring(0, 10000) + '...(truncated)';
        }
        return {
            success: true,
            data: { text: text, length: length }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 插入文本
 * @param {Object} params - { text: string, position?: 'start'|'end'|'cursor' }
 */
function insertText(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        var text = params.text || '';
        var position = params.position || 'cursor';

        switch (position) {
            case 'start':
                var range = doc.Range(0, 0);
                range.InsertBefore(text);
                break;
            case 'end':
                var endPos = doc.Content.End - 1;
                var range = doc.Range(endPos, endPos);
                range.InsertAfter(text);
                break;
            default: // cursor
                app.Selection.TypeText(text);
        }
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 设置字体
 * @param {Object} params - { range?: 'all'|'selection', fontName?, fontSize?, bold?, italic? }
 */
function setFont(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        var range = (params.range === 'all') ? doc.Content : app.Selection.Range;

        if (params.fontName) range.Font.Name = params.fontName;
        if (params.fontSize) range.Font.Size = params.fontSize;
        if (params.bold !== undefined) range.Font.Bold = params.bold;
        if (params.italic !== undefined) range.Font.Italic = params.italic;

        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 查找替换
 * @param {Object} params - { findText: string, replaceText: string, replaceAll?: boolean }
 */
function findReplace(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        var find = doc.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = params.findText;
        find.Replacement.Text = params.replaceText || '';
        // replaceType: 1=单个, 2=全部
        var replaceType = params.replaceAll ? 2 : 1;
        var result = find.Execute(
            params.findText,  // FindText
            false,            // MatchCase
            false,            // MatchWholeWord
            false,            // MatchWildcards
            false,            // MatchSoundsLike
            false,            // MatchAllWordForms
            true,             // Forward
            1,                // Wrap (wdFindContinue)
            false,            // Format
            params.replaceText || '',  // ReplaceWith
            replaceType       // Replace
        );
        return { success: true, data: { replaced: result } };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 插入表格
 * @param {Object} params - { rows: number, cols: number, data?: array[][] }
 */
function insertTable(params) {
    try {
        var app = Application;
        var doc = app.ActiveDocument;
        if (!doc) {
            return { success: false, error: '没有打开的文档' };
        }
        var rows = params.rows || 3;
        var cols = params.cols || 3;
        var range = app.Selection.Range;
        var table = doc.Tables.Add(range, rows, cols);

        // 填充数据
        if (params.data && Array.isArray(params.data)) {
            var maxRows = Math.min(params.data.length, rows);
            for (var r = 0; r < maxRows; r++) {
                var rowData = params.data[r];
                if (Array.isArray(rowData)) {
                    var maxCols = Math.min(rowData.length, cols);
                    for (var c = 0; c < maxCols; c++) {
                        table.Cell(r + 1, c + 1).Range.Text = String(rowData[c]);
                    }
                }
            }
        }
        // 启用边框
        table.Borders.Enable = true;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 应用样式
 * @param {Object} params - { styleName: string }
 */
function applyStyle(params) {
    try {
        var app = Application;
        var range = app.Selection.Range;
        range.Style = params.styleName;
        return { success: true };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 获取文档段落结构
 */
function getDocumentParagraphs(params) {
    try {
        params = params || {};
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };

        var totalCount = Number(doc.Paragraphs.Count) || 0;
        if (totalCount === 0) {
            return { success: true, data: { paragraphs: [], totalCount: 0, returnedCount: 0 } };
        }
        var startIndex = params.startParagraph === undefined ? 1 : Number(params.startParagraph);
        var endIndex = params.endParagraph === undefined
            ? Math.min(totalCount, startIndex + 49)
            : Number(params.endParagraph);
        if (!isFinite(startIndex) || !isFinite(endIndex)
            || startIndex < 1 || endIndex < startIndex
            || startIndex % 1 !== 0 || endIndex % 1 !== 0) {
            return { success: false, error: '段落范围无效' };
        }
        endIndex = Math.min(endIndex, totalCount);

        var paragraphs = [];
        for (var i = startIndex; i <= endIndex; i++) {
            var paragraph = doc.Paragraphs.Item(i) || doc.Paragraphs(i);
            var text = String(paragraph.Range.Text || '').replace(/[\r\n]+$/, '');
            if (text.length > 200) text = text.substr(0, 200) + '...';
            var style = '';
            try {
                var paragraphStyle = paragraph.Range.Style;
                style = paragraphStyle.NameLocal || paragraphStyle.Name || String(paragraphStyle || '');
            } catch (ignored) {}
            paragraphs.push({
                index: i,
                text: text,
                style: style,
                start: paragraph.Range.Start,
                end: paragraph.Range.End
            });
        }
        return {
            success: true,
            data: { paragraphs: paragraphs, totalCount: totalCount, returnedCount: paragraphs.length }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 查找文档中的文本并返回位置，不修改文档
 */
function findInDocument(params) {
    try {
        params = params || {};
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };
        if (!params.findText) return { success: false, error: '查找文本不能为空' };

        var maxResults = params.maxResults === undefined ? 20 : Number(params.maxResults);
        if (!isFinite(maxResults) || maxResults < 1 || maxResults % 1 !== 0) {
            return { success: false, error: '最大返回结果数无效' };
        }

        var results = [];
        var contentEnd = doc.Content.End;
        var searchRange = doc.Content.Duplicate;
        while (results.length < maxResults) {
            searchRange.Find.ClearFormatting();
            var found = searchRange.Find.Execute(
                params.findText,
                !!params.matchCase,
                !!params.matchWholeWord,
                false,
                false,
                false,
                true,
                0,
                false,
                '',
                0
            );
            if (!found) break;

            var matchStart = searchRange.Start;
            var matchEnd = searchRange.End;
            var paragraphIndex = 0;
            for (var i = 1; i <= Number(doc.Paragraphs.Count); i++) {
                var paragraphRange = doc.Paragraphs.Item(i).Range;
                if (matchStart >= paragraphRange.Start && matchStart <= paragraphRange.End) {
                    paragraphIndex = i;
                    break;
                }
            }
            results.push({
                text: String(searchRange.Text || ''),
                start: matchStart,
                end: matchEnd,
                paragraphIndex: paragraphIndex,
                context: String(doc.Range(Math.max(0, matchStart - 50), Math.min(contentEnd, matchEnd + 50)).Text || '')
            });
            if (matchEnd >= contentEnd) break;
            searchRange = doc.Range(matchEnd, contentEnd);
        }

        return {
            success: true,
            data: { results: results, count: results.length, findText: params.findText }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 智能填写模板字段
 */
function smartFillField(params) {
    try {
        params = params || {};
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };
        if (!params.keyword || String(params.keyword).trim() === '') {
            return { success: false, error: '关键字不能为空' };
        }
        if (params.value === undefined || params.value === null) {
            return { success: false, error: '填写值不能为空（空字符串将清除该字段内容）' };
        }

        var searchRange = doc.Content.Duplicate;
        searchRange.Find.ClearFormatting();
        var found = searchRange.Find.Execute(
            params.keyword, false, false, false, false, false, true, 1, false, '', 0
        );
        if (!found) return { success: false, error: '未找到关键字: ' + params.keyword };

        var matchStart = searchRange.Start;
        var matchEnd = searchRange.End;
        var paragraphRange = searchRange.Paragraphs.Item(1).Range;
        var paragraphText = String(paragraphRange.Text || '');
        var fillMode = params.fillMode || 'auto';
        var escapedKeyword = String(params.keyword).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

        if (fillMode === 'auto') {
            if (new RegExp('[{【\\[]' + escapedKeyword + '[}】\\]]').test(paragraphText)) {
                fillMode = 'placeholder';
            } else {
                var afterKeyword = paragraphText.substr(Math.max(0, matchEnd - paragraphRange.Start));
                if (/^[\s]*[：:][\s]*_+/.test(afterKeyword)) {
                    fillMode = 'underline';
                } else if (/^[\s]*[：:]/.test(afterKeyword)) {
                    fillMode = 'afterColon';
                } else {
                    fillMode = 'afterLabel';
                }
            }
        }

        var resultMessage;
        if (fillMode === 'placeholder') {
            var placeholderMatch = paragraphText.match(
                new RegExp('([{【\\[])' + escapedKeyword + '([}】\\]])')
            );
            var placeholder = placeholderMatch ? placeholderMatch[0] : params.keyword;
            var placeholderOffset = paragraphText.indexOf(placeholder);
            doc.Range(
                paragraphRange.Start + placeholderOffset,
                paragraphRange.Start + placeholderOffset + placeholder.length
            ).Text = params.value;
            resultMessage = 'replaced placeholder';
        } else if (fillMode === 'underline') {
            var underlineText = doc.Range(matchEnd, paragraphRange.End).Text || '';
            var underlineMatch = /_+/.exec(underlineText);
            if (underlineMatch) {
                var underlineStart = matchEnd + underlineMatch.index;
                var underlineRange = doc.Range(underlineStart, underlineStart + underlineMatch[0].length);
                underlineRange.Text = params.value;
                try { underlineRange.Font.Underline = 1; } catch (ignored) {}
                resultMessage = 'replaced underline';
            } else {
                doc.Range(matchEnd, matchEnd).InsertAfter(params.value);
                resultMessage = 'inserted after keyword';
            }
        } else if (fillMode === 'afterColon') {
            var suffixRange = doc.Range(matchEnd, paragraphRange.End);
            var suffix = String(suffixRange.Text || '');
            var colonIndex = suffix.search(/[：:]/);
            if (colonIndex < 0) {
                doc.Range(matchEnd, matchEnd).InsertAfter('：' + params.value);
                resultMessage = 'inserted after new colon';
            } else {
                var valueStart = matchEnd + colonIndex + 1;
                var valueEnd = paragraphRange.End;
                while (valueEnd > valueStart && /[\r\n]/.test(doc.Range(valueEnd - 1, valueEnd).Text || '')) {
                    valueEnd--;
                }
                doc.Range(valueStart, valueEnd).Text = params.value;
                resultMessage = 'replaced content after colon';
            }
        } else {
            doc.Range(matchEnd, matchEnd).InsertAfter(params.value);
            resultMessage = 'inserted after label';
        }

        return {
            success: true,
            data: {
                keyword: params.keyword,
                value: params.value,
                fillMode: fillMode,
                result: resultMessage
            }
        };
    } catch (e) {
        return { success: false, error: e.message };
    }
}

/**
 * 替换书签内容并重建书签范围
 */
function replaceBookmarkContent(params) {
    try {
        params = params || {};
        var doc = Application.ActiveDocument;
        if (!doc) return { success: false, error: '没有打开的文档' };
        if (!params.name || String(params.name).trim() === '') {
            return { success: false, error: '书签名称不能为空' };
        }
        if (params.text === undefined || params.text === null) {
            return { success: false, error: '替换文本不能为空（空字符串将清空书签）' };
        }

        var bookmark = doc.Bookmarks.Item(params.name) || doc.Bookmarks(params.name);
        if (!bookmark) return { success: false, error: '未找到书签: ' + params.name };
        var start = bookmark.Start;
        bookmark.Range.Text = params.text;
        var end = start + String(params.text).length;
        doc.Bookmarks.Add(params.name, doc.Range(start, end));
        return {
            success: true,
            data: { name: params.name, text: params.text, start: start, end: end }
        };
    } catch (e) {
        return { success: false, error: '书签替换失败: ' + e.message };
    }
}

// 导出模块
module.exports = {
    getActiveDocument: getActiveDocument,
    getDocumentText: getDocumentText,
    insertText: insertText,
    setFont: setFont,
    findReplace: findReplace,
    insertTable: insertTable,
    applyStyle: applyStyle,
    getDocumentParagraphs: getDocumentParagraphs,
    findInDocument: findInDocument,
    smartFillField: smartFillField,
    replaceBookmarkContent: replaceBookmarkContent
};
