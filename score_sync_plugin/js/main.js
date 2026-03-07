// 全局变量
let studentIdToName = {}; // 学号到姓名
let studentIdToGroup = {}; // 学号到组
let groupToStudentIds = {}; // 组到学号列表（统计组员）
let excelStudentData = []; // 导入的Excel成绩模板数据
let excelHeader = []; // 存储Excel模板的原始表头
// 动态配置变量
let EXPORT_COLUMNS = []; // 导出列
let EDITABLE_COLUMNS = []; // 可编辑列
let SYNC_COLUMNS = []; // 同步列

// 绑定按钮事件
window.onload = function() {
    // 上传组信息文件按钮事件 (txt/pdf/doc/docx)
    document.getElementById("groupFileInput").addEventListener("change", handleGroupFileUpload);
    // 上传Excel成绩模板按钮事件
    document.getElementById("excelFileInput").addEventListener("change", handleExcelUpload);
    // 导出成绩表按钮事件
    document.getElementById("exportBtn").addEventListener("click", exportExcelResult);
    // 确认列配置按钮事件
    document.getElementById("confirmColumnsBtn").addEventListener("click", confirmColumnConfig);
    // 导出TXT按钮事件
    document.getElementById("exportTxtBtn").addEventListener("click", exportTxtResult);
}

// 功能1：多文件上传
async function handleGroupFileUpload(e) {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    let totalGroups = 0;
    let totalStudents = 0;
    const failedFiles = [];

    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fileName = file.name;
        const fileType = fileName.split('.').pop().toLowerCase();
        let groupText = "";

        try {
            switch (fileType) {
                case "txt":
                    groupText = await readTxtFile(file);
                    break;
                case "pdf":
                    groupText = await parsePdfFile(file);
                    // 检查PDF解析后文本是否为空
                    if (!groupText) {
                        failedFiles.push(`${fileName}（PDF无有效文本）`);
                        continue;
                    }
                    break;
                case "docx":
                    groupText = await readWordFile(file);
                    // 检查Word解析后文本是否为空
                    if (!groupText) {
                        failedFiles.push(`${fileName}（Word无有效文本）`);
                        continue;
                    }
                    break;
                case "doc":
                    failedFiles.push(`${fileName}（不支持.doc格式，请转为.docx）`);
                    continue;
                default:
                    failedFiles.push(`${fileName}（不支持的格式）`);
                    continue;
            }

            // 提取学号姓名
            const members = extractMembersFromGroupFile(groupText);
            if (members.length === 0) {
                failedFiles.push(`${fileName}（未提取到学号/姓名信息）`);
                continue;
            }

            // 生成小组映射
            const groupId = `组${totalGroups + 1}`;
            groupToStudentIds[groupId] = members.map(item => item.id);
            members.forEach(member => {
                studentIdToName[member.id] = member.name;
                studentIdToGroup[member.id] = groupId;
            });

            totalGroups++;
            totalStudents += members.length;
        } catch (error) {
            // 捕获所有类型错误
            failedFiles.push(`${fileName}（${error.message}）`);
            continue;
        }
    }

    // 汇总反馈
    let alertMsg = `批量解析完成！\n 共上传${files.length}个文件\n 成功解析${totalGroups}个小组，${totalStudents}名学生`;
    if (failedFiles.length > 0) {
        alertMsg += `\n\n 解析失败的文件：\n${failedFiles.join('\n')}`;
    }
    alert(alertMsg);
}

// 功能2：解析TXT文件
function readTxtFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.readAsText(file, "utf-8");
    });
}

// 功能3：解析Word(docx)文件
async function readWordFile(file) {
    try {
        // 读取docx文件为ArrayBuffer
        const arrayBuffer = await file.arrayBuffer();
        
        // 将docx转为纯文本，跳过图片/表格/格式
        const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
        let text = result.value; // 提取纯文本内容
        
        // 文本格式化
        text = text.replace(/\s+/g, ' ')           // 合并多个空格
                   .replace(/(\r\n|\r|\n)+/g, '\n')// 统一换行符
                   .replace(/：/g, ':')             // 冒号转换
                   .replace(/，/g, ',')             // 逗号转换
                   .trim();
        
        // 空文本检查
        if (!text) {
            throw new Error("Word文件中无有效文本内容");
        }
        
        return text;
    } catch (error) {
        let errMsg = "Word解析失败：";
        if (error.message.includes("Unsupported file format")) {
            errMsg += "不是标准的docx文件（不支持.doc）";
        } else {
            errMsg += error.message;
        }
        throw new Error(errMsg);
    }
}

// 功能4：从新模板中提取学号姓名
function extractMembersFromGroupFile(text) {
    // 正则匹配格式
    const memberReg = /(?:\[)?([^[\]-]+)(?:\])?\s*-\s*(\d+)\s*-\s*[\w@.]+/g;
    const members = []; // id:学号  name:姓名
    let match;

    while ((match = memberReg.exec(text)) !== null) {
        const name = match[1].trim();
        const studentId = match[2].trim(); // 学号作为唯一标识
        if (name && studentId && !members.some(item => item.id === studentId)) {
            members.push({ id: studentId, name: name });
        }
    }

    return members; // 返回学号姓名列表
}

// 功能5：离线解析PDF文件
async function parsePdfFile(file) {
    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
        const page = await pdf.getPage(1);
        const content = await page.getTextContent();
        let text = content.items.map(item => item.str).join('\n');
        
        // 移除多余空格换行，统一分隔符
        text = text.replace(/\s+/g, ' ') // 合并多个空格
                   .replace(/(\r\n|\r|\n)+/g, '\n') // 统一换行符
                   .replace(/：/g, ':') // 冒号转换
                   .replace(/，/g, ',') // 逗号转换
                   .trim();
        
        return text;
    } catch (error) {
        alert("PDF解析失败：文件格式错误");
        return "";
    }
}

// 功能6：解析TXT成绩模板
function parseTxtTemplate(text) {
    const lines = text.trim().split('\n').filter(line => line);
    if (lines.length < 2) {
        throw new Error("TXT模板格式错误：需包含表头和至少一条数据");
    }
    // 第一行作为表头
    const header = lines[0].split('\t').map(col => col.trim());
    // 后续行作为数据
    const data = lines.slice(1).map(line => {
        const values = line.split('\t').map(val => val.trim());
        const row = {};
        header.forEach((col, index) => {
            row[col] = values[index] || ""; // 空值设为空字符串
        });
        return row;
    });
    return { header, data };
}

// 功能7：生成双向映射
function convertToStandardMap(text) {
    studentToGroup = {};
    groupToStudents = {};
    const reg = /组(\d+)[:：](.*)/g;
    let match;
    while ((match = reg.exec(text)) !== null) {
        const groupId = match[1];
        const students = match[2].split(/[,，、;；]/).map(name => name.trim()).filter(name => name);
        groupToStudents[groupId] = students;
        students.forEach(student => studentToGroup[student] = groupId);
    }
}

// 功能8：导入成绩模板
async function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    const fileName = file.name;
    const fileType = fileName.split('.').pop().toLowerCase();
    let header = [];
    let data = [];

    try {
        // 区分文件类型
        if (["xlsx", "xls"].includes(fileType)) {
            // Excel文件解析
            const reader = new FileReader();
            reader.onload = function(e) {
                const dataBuffer = new Uint8Array(e.target.result);
                const workbook = XLSX.read(dataBuffer, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                header = XLSX.utils.sheet_to_json(worksheet, { header: 1 })[0].filter(col => col);
                data = XLSX.utils.sheet_to_json(worksheet);
                processImportData(header, data);
            };
            reader.readAsArrayBuffer(file);
        } else if (fileType === "txt") {
            // TXT文件解析
            const reader = new FileReader();
            reader.onload = function(e) {
                const text = e.target.result;
                const { header: txtHeader, data: txtData } = parseTxtTemplate(text);
                header = txtHeader;
                data = txtData;
                processImportData(header, data);
            };
            reader.readAsText(file, "utf-8");
        } else {
            alert("不支持的文件格式！");
            return;
        }
    } catch (error) {
        alert(`模板导入失败：${error.message}`);
        return;
    }
}

// 功能9：处理导入后的数据
function processImportData(header, data) {
    excelHeader = header;
    excelStudentData = data;

    // 初始化默认值
    excelStudentData.forEach(student => {
        excelHeader.forEach(col => {
            if (!student[col]) {
                // 数字列设为0，文本列设为空
                student[col] = /^[\d.]+$/.test(String(student[col])) ? 0 : "";
            }
        });
    });

    // 渲染列选择UI
    renderColumnSelectUI();
    document.getElementById("columnSelectContainer").style.display = "block";

    alert(`模板导入成功！\n共读取 ${excelStudentData.length} 名学生信息\n检测到表头：${excelHeader.join('、')}`);
}

// 功能10：渲染列选择UI
function renderColumnSelectUI() {
    const exportColumnsList = document.getElementById("exportColumnsList");
    const editableColumnsList = document.getElementById("editableColumnsList");
    const syncColumnsList = document.getElementById("syncColumnsList");
    
    // 清空原有内容
    exportColumnsList.innerHTML = "";
    editableColumnsList.innerHTML = "";
    syncColumnsList.innerHTML = "";

    // 渲染模板中存在的表头
    excelHeader.forEach(col => {
        // 导出列选择（默认全选）
        const exportItem = document.createElement("div");
        exportItem.className = "column-item";
        exportItem.innerHTML = `
            <input type="checkbox" id="export_${col}" value="${col}" checked>
            <label for="export_${col}">${col}</label>
        `;
        exportColumnsList.appendChild(exportItem);

        // 可编辑列选择
        const isBasicCol = ["序号", "班级", "学号", "姓名"].includes(col);
        const editableItem = document.createElement("div");
        editableItem.className = "column-item";
        editableItem.innerHTML = `
            <input type="checkbox" id="editable_${col}" value="${col}" ${!isBasicCol ? "checked" : ""}>
            <label for="editable_${col}">${col}</label>
        `;
        editableColumnsList.appendChild(editableItem);

        // 同步列选择
        const isSyncCol = ["课堂陈述与讨论(必填)","案例分析(必填)","专题讨论(必填)","期末考试(必填)"].includes(col);
        const syncItem = document.createElement("div");
        syncItem.className = "column-item";
        syncItem.innerHTML = `
            <input type="checkbox" id="sync_${col}" value="${col}" ${isSyncCol ? "checked" : ""}>
            <label for="sync_${col}">${col}</label>
        `;
        syncColumnsList.appendChild(syncItem);
    });
}

// 功能11：确认列配置
function confirmColumnConfig() {
    // 获取导出列
    EXPORT_COLUMNS = Array.from(document.querySelectorAll("#exportColumnsList input:checked"))
        .map(input => input.value);
    
    // 获取可编辑列
    EDITABLE_COLUMNS = Array.from(document.querySelectorAll("#editableColumnsList input:checked"))
        .map(input => input.value);
    
    // 获取同步列
    SYNC_COLUMNS = Array.from(document.querySelectorAll("#syncColumnsList input:checked"))
        .map(input => input.value);

    // 验证必要列
    if (!EXPORT_COLUMNS.includes("学号")) {
        alert("导出列必须包含学号！");
        return;
    }
    if (SYNC_COLUMNS.length === 0) {
        alert("请至少选择一个同步列！");
        return;
    }

    // 渲染表格
    renderExcelTable();
    //// 隐藏列配置区域
    // document.getElementById("columnSelectContainer").style.display = "none";
    alert(`列配置完成！\n导出列：${EXPORT_COLUMNS.join('、')}\n可编辑列：${EDITABLE_COLUMNS.join('、')}\n同步列：${SYNC_COLUMNS.join('、')}`);
}

// 功能12：渲染表格
function renderExcelTable() {
    const tableContainer = document.getElementById("tableContainer");
    tableContainer.innerHTML = "";
    const table = document.createElement("table");
    table.border = 1;
    table.cellPadding = 5;

    // 生成表头
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    EXPORT_COLUMNS.forEach(key => {
        const th = document.createElement("th");
        th.innerText = key;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // 生成表格内容
    const tbody = document.createElement("tbody");
    excelStudentData.forEach((student, index) => {
        const tr = document.createElement("tr");
        EXPORT_COLUMNS.forEach(key => {
            const td = document.createElement("td");
            // 同步列采用数字输入框
            if (SYNC_COLUMNS.includes(key)) {
                td.innerHTML = `<input type="number" step="0.1" min="0" max="100" value="${student[key]}" 
                    style="width:90px;text-align:center;"
                    onblur="syncScore(${index}, '${key}', this.value)">`;
            } 
            // 其他可编辑列采用文本输入框
            else if (EDITABLE_COLUMNS.includes(key)) {
                td.innerHTML = `<input type="text" value="${student[key]}" 
                    style="width:120px;text-align:left;"
                    onblur="updateOtherColumn(${index}, '${key}', this.value)">`;
            }
            // 非可编辑列采用纯文本
            else {
                td.innerText = student[key] ?? "";
            }
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    tableContainer.appendChild(table);
}

// 功能13：更新非同步的可编辑列
function updateOtherColumn(index, column, value) {
    excelStudentData[index][column] = value;
}

// 功能14：成绩同步
function syncScore(index, itemKey, score) {
    // 校验成绩格式
    if (isNaN(score) || score < 0 || score > 100) {
        alert("成绩格式错误！请输入0-100的数字");
        return;
    }

    // 获取当前学生的学号
    const currentStudent = excelStudentData[index];
    const studentId = currentStudent.学号?.toString().trim(); // Excel中的学号转字符串
    const studentName = currentStudent.姓名 || studentIdToName[studentId];

    if (!studentId) {
        alert("未找到学生学号，无法同步成绩");
        return;
    }

    // 用学号查找所属组
    const groupId = studentIdToGroup[studentId];
    if (!groupId) {
        alert(`${studentName}（学号：${studentId}）未分配小组，不进行成绩同步`);
        return;
    }

    // 获取同组所有学生的学号
    const sameGroupStudentIds = groupToStudentIds[groupId];
    if (!sameGroupStudentIds || sameGroupStudentIds.length === 0) {
        alert(`组${groupId}无有效组员，同步失败`);
        return;
    }

    // 同步学生成绩
    excelStudentData.forEach(student => {
        const targetStudentId = student.学号?.toString().trim();
        if (sameGroupStudentIds.includes(targetStudentId)) {
            student[itemKey] = score;
        }
    });

    // 重新渲染表格
    renderExcelTable();
    alert(`成绩同步完成！\n已将【组${groupId}】的【${itemKey}】成绩同步为：${score}分\n同步学生：${sameGroupStudentIds.length}人`);
}

// 功能15：导出成绩表
function exportExcelResult() {
    if (excelStudentData.length === 0) {
        alert("请先导入Excel成绩模板！");
        return;
    }
    // 按指定列顺序重构数据
    const exportData = excelStudentData.map(student => {
        const formattedStudent = {};
        EXPORT_COLUMNS.forEach(col => {
            formattedStudent[col] = student[col] ?? "";
        });
        return formattedStudent;
    });
    // 生成工作表
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "学生成绩表");
    XLSX.writeFile(workbook, `小组成绩输入完成_${new Date().toLocaleDateString()}.xlsx`);
}

// 功能16：导出为制表符分隔的TXT格式
function exportTxtResult() {
    if (excelStudentData.length === 0) {
        alert("请先导入成绩模板！");
        return;
    }
    if (EXPORT_COLUMNS.length === 0) {
        alert("请先配置导出列！");
        return;
    }

    // 表头（制表符分隔）
    const headerLine = EXPORT_COLUMNS.join('\t');
    // 数据（制表符分隔，空值保留为空字符串）
    const dataLines = excelStudentData.map(student => {
        return EXPORT_COLUMNS.map(col => {
            const value = student[col] ?? "";
            // 数字类型去除多余小数点
            return typeof value === "number" ? (value % 1 === 0 ? value : value.toFixed(1)) : value;
        }).join('\t');
    });

    // 拼接完整TXT内容
    const txtContent = [headerLine, ...dataLines].join('\n');
    // 导出TXT文件
    const blob = new Blob([txtContent], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `小组成绩输入完成_${new Date().toLocaleDateString()}.txt`;
    a.click();
    URL.revokeObjectURL(url);

    alert("TXT成绩表导出成功！");
}