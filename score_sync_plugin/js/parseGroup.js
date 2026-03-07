/**
 * 解析PDF文件 - 读取第一页文本内容
 * @param {File} file - 上传的PDF文件对象
 * @returns {Promise<string>} PDF第一页的文本内容
 */
async function parsePdfFile(file) {
    try {
        const arrayBuffer = await file.arrayBuffer();
        // pdfjsLib调用方式
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
        const page = await pdf.getPage(1); // 读取第一页
        const content = await page.getTextContent();
        // 提取文本，拼接内容
        const text = content.items.map(item => item.str).join('\n');
        return text;
    } catch (error) {
        console.error("PDF解析失败：", error);
        alert("PDF文件解析失败，请检查文件是否为有效PDF格式！");
        return "";
    }
}