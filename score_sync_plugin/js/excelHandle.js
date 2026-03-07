/**
 * 导入学校Excel成绩模板，解析为JSON数组
 * @param {File} file - 上传的Excel文件
 * @returns {Promise<Array>} 解析后的学生成绩数组
 */
async function importExcelTemplate(file) {
    const reader = new FileReader();
    return new Promise((resolve) => {
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0]; // 读取第一个工作表
            const worksheet = workbook.Sheets[sheetName];
            // 解析为JSON，保持原表头
            const studentData = XLSX.utils.sheet_to_json(worksheet);
            resolve(studentData);
        };
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 导出成绩表
 * @param {Array} studentData - 同步后的学生成绩数组
 * @param {String} sheetName - 工作表名称
 */
function exportExcelResult(studentData, sheetName = "成绩表") {
    const worksheet = XLSX.utils.json_to_sheet(studentData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    // 导出为Excel
    XLSX.writeFile(workbook, `成绩录入完成_${new Date().toLocaleDateString()}.xlsx`);
}