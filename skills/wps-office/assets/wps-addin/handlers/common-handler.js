/**
 * Input: 通用 action 与参数
 * Output: 通用操作结果
 * Pos: macOS 加载项通用处理器。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * 通用操作处理器
 * 处理跨应用的通用操作
 */

function handleCommon(action, params) {
    // TODO: 实现通用操作
    return { success: false, error: '未实现: ' + action };
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { handleCommon };
}
