/**
 * Input: 响应数据与错误信息
 * Output: 标准化响应对象
 * Pos: macOS 加载项响应封装工具。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 响应封装工具
 * 统一的响应格式，别TM乱改格式
 */

function success(data) {
    return {
        success: true,
        data: data || {},
        error: null
    };
}

function error(message) {
    return {
        success: false,
        data: null,
        error: message || '未知错误'
    };
}

function timeout(operation) {
    return {
        success: false,
        data: null,
        error: (operation ? operation + '操作超时' : '操作超时') + '，请稍后重试'
    };
}

function paramError(message) {
    return {
        success: false,
        data: null,
        error: '参数错误: ' + (message || '请检查输入参数')
    };
}

// 导出
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { success, error, timeout, paramError };
}
