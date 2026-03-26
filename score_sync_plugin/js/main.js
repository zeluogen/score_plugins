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
// 导入文件管理
let importedGroupFiles = []; // 存储导入的分组文件信息

// 配置管理默认初始配置
let settings = {
    // 列配置初始值
    defaultEditableColumns: ["课堂陈述与讨论(必填)", "案例分析(必填)", "专题讨论(必填)", "期末考试(必填)", "备注"],
    defaultSyncColumns: ["课堂陈述与讨论(必填)", "案例分析(必填)", "专题讨论(必填)", "期末考试(必填)"],
    
    // 学号解析规则初始值
    studentIdRegex: "\\d{12}",
    
    // 同步列输入规则初始值
    scoreMin: 0,
    scoreMax: 100,
    scoreDecimalPlaces: 1,
    
    // 成绩同步设置初始值
    syncEnabled: true,
    
    // 导出设置初始值
    fileNameTemplate: "小组成绩输入完成_{date}",
    exportPath: ""
};

// 默认配置
let defaultSettings = JSON.parse(JSON.stringify(settings));

// 绑定按钮事件
window.onload = function() {
    // 上传组信息文件按钮事件
    document.getElementById("groupFileInput").addEventListener("change", handleGroupFileUpload);
    // 上传Excel成绩模板按钮事件
    document.getElementById("excelFileInput").addEventListener("change", handleExcelUpload);
    // 确认列配置按钮事件
    document.getElementById("confirmColumnsBtn").addEventListener("click", confirmColumnConfig);
    
    // 配置管理事件
    document.getElementById("saveSettingsBtn").addEventListener("click", saveSettings);
    document.getElementById("resetSettingsBtn").addEventListener("click", resetSettings);
    
    // 成绩统计事件
    document.getElementById("statsColumnSelect").addEventListener("change", updateStatsDisplay);
    
    // 加载配置
    loadSettings();
    
    // 初始化统计部分
    initStatsSection();
    
    // 平滑滚动效果
    addSmoothScroll();
}


/*------------------------------ 平滑滚动相关 ------------------------------*/
// 平滑滚动函数
function addSmoothScroll() {
    const navLinks = document.querySelectorAll('.sidebar nav ul li a');
    
    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            
            const targetId = this.getAttribute('href');
            const targetElement = document.querySelector(targetId);
            
            if (targetElement) {
                smoothScrollTo(targetElement);
            }
        });
    });
}

// 实现平滑滚动
function smoothScrollTo(element) {
    const startPosition = window.pageYOffset;
    const targetPosition = element.getBoundingClientRect().top + window.pageYOffset;
    const distance = targetPosition - startPosition;
    const duration = 800; // 滚动持续时间
    let startTime = null;
    
    function animation(currentTime) {
        if (startTime === null) startTime = currentTime;
        const timeElapsed = currentTime - startTime;
        const progress = Math.min(timeElapsed / duration, 1);
        
        // 使用缓动函数实现高速滚动后减速的效果
        const easeOutQuart = 1 - Math.pow(1 - progress, 4);
        window.scrollTo(0, startPosition + distance * easeOutQuart);
        
        if (timeElapsed < duration) {
            requestAnimationFrame(animation);
        }
    }
    
    requestAnimationFrame(animation);
}
/*-------------------------------------------------------------------------*/


/*------------------------------ 配置管理相关 ------------------------------*/
// 加载配置
function loadSettings() {
    try {
        // 从localStorage加载配置
        const savedSettings = localStorage.getItem('scoreSyncSettings');
        if (savedSettings) {
            settings = JSON.parse(savedSettings);
        }
        updateSettingsUI();
    } catch (error) {
        console.error('加载配置失败：', error);
    }
}

// 保存配置
function saveSettings() {
    try {
        // 从界面获取配置
        settings.defaultEditableColumns = document.getElementById('defaultEditableColumns').value.split(',').map(col => col.trim()).filter(col => col);
        settings.defaultSyncColumns = document.getElementById('defaultSyncColumns').value.split(',').map(col => col.trim()).filter(col => col);
        settings.studentIdRegex = document.getElementById('studentIdRegex').value;
        settings.scoreMin = parseFloat(document.getElementById('scoreMin').value) || 0;
        settings.scoreMax = parseFloat(document.getElementById('scoreMax').value) || 100;
        settings.scoreDecimalPlaces = parseInt(document.getElementById('scoreDecimalPlaces').value) || 1;
        settings.syncEnabled = document.getElementById('syncEnabled').checked;
        settings.fileNameTemplate = document.getElementById('fileNameTemplate').value || '小组成绩输入完成_{date}';
        
        // 保存到localStorage
        localStorage.setItem('scoreSyncSettings', JSON.stringify(settings));
        
        // 如果已经导入了模板，重新渲染列选择界面
        if (excelHeader.length > 0) {
            renderColumnSelectUI();
        }
        
        console.log('配置保存成功');
    } catch (error) {
        console.error('保存配置失败：', error);
    }
}

// 恢复默认配置
function resetSettings() {
    try {
        // 使用默认配置
        settings = JSON.parse(JSON.stringify(defaultSettings));
        updateSettingsUI();
        localStorage.removeItem('scoreSyncSettings');
        
        // 如果已经导入了模板，重新渲染列选择界面
        if (excelHeader.length > 0) {
            renderColumnSelectUI();
        }
        
        console.log('配置已恢复默认');
    } catch (error) {
        console.error('恢复默认配置失败：', error);
    }
}

// 更新配置界面
function updateSettingsUI() {
    try {
        document.getElementById('defaultEditableColumns').value = settings.defaultEditableColumns.join(', ');
        document.getElementById('defaultSyncColumns').value = settings.defaultSyncColumns.join(', ');
        document.getElementById('studentIdRegex').value = settings.studentIdRegex;
        document.getElementById('scoreMin').value = settings.scoreMin;
        document.getElementById('scoreMax').value = settings.scoreMax;
        document.getElementById('scoreDecimalPlaces').value = settings.scoreDecimalPlaces;
        document.getElementById('syncEnabled').checked = settings.syncEnabled;
        document.getElementById('fileNameTemplate').value = settings.fileNameTemplate;
    } catch (error) {
        console.error('更新配置界面失败：', error);
    }
}
/*-------------------------------------------------------------------------*/


/*---------------------------- 分组文件上传相关 ----------------------------*/    
// 多文件上传
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

            // 检查是否已存在同名文件
            const existingFileIndex = importedGroupFiles.findIndex(f => f.name === fileName);
            if (existingFileIndex !== -1) {
                // 覆盖同名文件
                const existingFile = importedGroupFiles[existingFileIndex];
                // 移除旧文件的学生信息
                Object.keys(existingFile.students).forEach(studentId => {
                    delete studentIdToName[studentId];
                    delete studentIdToGroup[studentId];
                });
                // 移除旧文件的组信息
                if (existingFile.groupId && groupToStudentIds[existingFile.groupId]) {
                    delete groupToStudentIds[existingFile.groupId];
                }
            }

            // 生成小组映射
            const groupId = `组${Object.keys(groupToStudentIds).length + 1}`;
            groupToStudentIds[groupId] = members.map(item => item.id);
            
            // 存储学生信息
            const students = {};
            members.forEach(member => {
                studentIdToName[member.id] = member.name;
                studentIdToGroup[member.id] = groupId;
                students[member.id] = member.name;
            });

            // 添加或更新文件信息
            if (existingFileIndex !== -1) {
                importedGroupFiles[existingFileIndex] = {
                    name: fileName,
                    groupId: groupId,
                    students: students,
                    studentCount: members.length
                };
            } else {
                importedGroupFiles.push({
                    name: fileName,
                    groupId: groupId,
                    students: students,
                    studentCount: members.length
                });
            }

            totalGroups++;
            totalStudents += members.length;
        } catch (error) {
            // 捕获所有类型错误
            failedFiles.push(`${fileName}（${error.message}）`);
            continue;
        }
    }

    // 更新文件列表显示
    updateGroupFileList();

    // 汇总反馈（使用console.log代替alert）
    console.log(`批量解析完成！共上传${files.length}个文件，成功解析${totalGroups}个小组，${totalStudents}名学生`);
    if (failedFiles.length > 0) {
        console.log(`解析失败的文件：${failedFiles.join(', ')}`);
    }
}

// 更新分组文件列表显示
function updateGroupFileList() {
    // 检查是否存在文件列表容器，不存在则创建
    let fileListContainer = document.getElementById('groupFileList');
    if (!fileListContainer) {
        const uploadSection = document.getElementById('upload-section');
        if (uploadSection) {
            fileListContainer = document.createElement('div');
            fileListContainer.id = 'groupFileList';
            fileListContainer.className = 'file-list';
            uploadSection.appendChild(fileListContainer);
        }
    }

    if (fileListContainer) {
        if (importedGroupFiles.length === 0) {
            fileListContainer.innerHTML = '<p style="color: #6b7280; font-size: 13px; margin-top: 10px;">暂无导入的分组文件</p>';
            return;
        }

        let html = '<h4 style="margin-top: 20px; margin-bottom: 10px;">已导入的分组文件</h4>';
        html += '<div style="display: flex; flex-direction: column; gap: 8px;">';
        
        importedGroupFiles.forEach((file, index) => {
            html += `
                <div style="display: flex; align-items: center; padding: 10px 12px; background: #f9fafb; border: 1px solid #e5e7eb; border-radius: 6px; font-size: 13px;">
                    <span style="flex: 1;">${file.name} (${file.studentCount}人)</span>
                    <button onclick="removeGroupFile(${index})" style="background: none; border: none; color: #6b7280; font-size: 16px; cursor: pointer; padding: 0; width: 20px; height: 20px; display: flex; align-items: center; justify-content: center;">&times;</button>
                </div>
            `;
        });
        
        html += '</div>';
        fileListContainer.innerHTML = html;
    }
}

// 删除分组文件
function removeGroupFile(index) {
    if (index >= 0 && index < importedGroupFiles.length) {
        const file = importedGroupFiles[index];
        // 移除学生信息
        Object.keys(file.students).forEach(studentId => {
            delete studentIdToName[studentId];
            delete studentIdToGroup[studentId];
        });
        // 移除组信息
        if (file.groupId && groupToStudentIds[file.groupId]) {
            delete groupToStudentIds[file.groupId];
        }
        // 从列表中移除
        importedGroupFiles.splice(index, 1);
        // 更新显示
        updateGroupFileList();
    }
}
/*-------------------------------------------------------------------------*/


/*------------------------------- 解析文件相关 -----------------------------*/
// 解析TXT文件
function readTxtFile(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.readAsText(file, "utf-8");
    });
}

// 解析Word文件
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

// 从新模板中提取学号姓名
function extractMembersFromGroupFile(text) {
    // 使用配置中的正则表达式匹配学号，同时尝试提取姓名
    try {
        const studentIdReg = new RegExp(settings.studentIdRegex, 'g');
        const members = []; // id:学号  name:姓名
        let match;

        // 首先提取所有学号
        const studentIds = [];
        while ((match = studentIdReg.exec(text)) !== null) {
            const studentId = match[0];
            if (!studentIds.includes(studentId)) {
                studentIds.push(studentId);
            }
        }

        // 为每个学号生成一个简单的姓名（实际应用中可能需要更复杂的逻辑）
        studentIds.forEach((studentId, index) => {
            // 尝试从文本中提取姓名
            let name = `学生${index + 1}`;
            // 简单的姓名提取逻辑：查找学号附近的中文姓名
            const nameReg = new RegExp(`([\u4e00-\u9fa5]+)\s*[\-:：]?\s*${studentId}`, 'g');
            const nameMatch = nameReg.exec(text);
            if (nameMatch && nameMatch[1]) {
                name = nameMatch[1].trim();
            }
            
            members.push({ id: studentId, name: name });
        });

        return members; // 返回学号姓名列表
    } catch (error) {
        console.error('学号解析失败：', error);
        // 使用默认正则表达式作为 fallback
        const studentIdReg = /(\d{12})/g;
        const members = [];
        let match;
        const studentIds = [];
        while ((match = studentIdReg.exec(text)) !== null) {
            const studentId = match[1];
            if (!studentIds.includes(studentId)) {
                studentIds.push(studentId);
            }
        }
        studentIds.forEach((studentId, index) => {
            let name = `学生${index + 1}`;
            const nameReg = new RegExp(`([\u4e00-\u9fa5]+)\s*[\-:：]?\s*${studentId}`, 'g');
            const nameMatch = nameReg.exec(text);
            if (nameMatch && nameMatch[1]) {
                name = nameMatch[1].trim();
            }
            members.push({ id: studentId, name: name });
        });
        return members;
    }
}

// 解析PDF文件
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

// 解析TXT成绩模板
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
/*-------------------------------------------------------------------------*/


// 生成双向映射
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

// 导入成绩模板
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
            console.log("不支持的文件格式！");
            return;
        }
    } catch (error) {
        console.log(`模板导入失败：${error.message}`);
        return;
    }
}

// 处理导入后的数据
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

    // 输出导入成功信息到控制台
    console.log(`模板导入成功！共读取 ${excelStudentData.length} 名学生信息，检测到表头：${excelHeader.join('、')}`);
}


/*------------------------------列选择相关 ------------------------------*/
// 渲染列选择UI
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
        // 导出列选择（始终默认全选）
        const exportItem = document.createElement("div");
        exportItem.className = "column-item";
        exportItem.innerHTML = `
            <input type="checkbox" id="export_${col}" value="${col}" checked>
            <label for="export_${col}">${col}</label>
        `;
        exportColumnsList.appendChild(exportItem);

        // 可编辑列选择（使用配置中的默认值）
        const isEditableCol = settings.defaultEditableColumns.includes(col);
        const editableItem = document.createElement("div");
        editableItem.className = "column-item";
        editableItem.innerHTML = `
            <input type="checkbox" id="editable_${col}" value="${col}" ${isEditableCol ? "checked" : ""}>
            <label for="editable_${col}">${col}</label>
        `;
        editableColumnsList.appendChild(editableItem);

        // 同步列选择（使用配置中的默认值）
        const isSyncCol = settings.defaultSyncColumns.includes(col); 
        const syncItem = document.createElement("div");
        syncItem.className = "column-item";
        syncItem.innerHTML = `
            <input type="checkbox" id="sync_${col}" value="${col}" ${isSyncCol ? "checked" : ""}>
            <label for="sync_${col}">${col}</label>
        `;
        syncColumnsList.appendChild(syncItem);
    });
}

// 确认列配置
function confirmColumnConfig() {
    
    // 获取导出列
    const exportCheckboxes = document.querySelectorAll("#exportColumnsList input:checked");
    console.log("导出列复选框数量：", exportCheckboxes.length);
    EXPORT_COLUMNS = Array.from(exportCheckboxes)
        .map(input => input.value);
    console.log("导出列：", EXPORT_COLUMNS);
    
    // 获取可编辑列
    const editableCheckboxes = document.querySelectorAll("#editableColumnsList input:checked");
    console.log("可编辑列复选框数量：", editableCheckboxes.length);
    EDITABLE_COLUMNS = Array.from(editableCheckboxes)
        .map(input => input.value);
    console.log("可编辑列：", EDITABLE_COLUMNS);
    
    // 获取同步列
    const syncCheckboxes = document.querySelectorAll("#syncColumnsList input:checked");
    console.log("同步列复选框数量：", syncCheckboxes.length);
    SYNC_COLUMNS = Array.from(syncCheckboxes)
        .map(input => input.value);
    console.log("同步列：", SYNC_COLUMNS);

    // 验证必要列
    if (!EXPORT_COLUMNS.includes("学号")) {
        console.log("导出列必须包含学号！");
        return;
    }

    console.log("excelStudentData长度：", excelStudentData.length);
    
    // 渲染表格
    renderExcelTable();
    // 更新统计列选择器
    updateStatsColumnSelect();
    //// 隐藏列配置区域
    // document.getElementById("columnSelectContainer").style.display = "none";
    console.log(`列配置完成！导出列：${EXPORT_COLUMNS.join('、')}，可编辑列：${EDITABLE_COLUMNS.join('、')}，同步列：${SYNC_COLUMNS.join('、')}`);
}
/*----------------------------------------------------------------------*/


/*-----------------------------表格编辑相关 -----------------------------*/
// 渲染表格
function renderExcelTable() {
    
    const tableContainer = document.getElementById("tableContainer");
    console.log("tableContainer元素：", tableContainer);
    
    if (!tableContainer) {
        console.error("未找到tableContainer元素");
        return;
    }
    
    tableContainer.innerHTML = "";
    console.log("清空tableContainer");

    // 添加表格顶部导出按钮
    const topExportArea = document.createElement("div");
    topExportArea.className = "export-area";
    topExportArea.style.marginBottom = "15px";
    topExportArea.innerHTML = `
        <button id="topExportBtn">导出Excel成绩表</button>
        <button id="topExportTxtBtn" class="secondary">导出TXT成绩表</button>
    `;
    tableContainer.appendChild(topExportArea);

    // 创建表格
    const table = document.createElement("table");

    // 生成表头
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    EXPORT_COLUMNS.forEach(key => {
        const th = document.createElement("th");
        th.innerText = key;
        // 为同步列设置一致的宽度
        if (SYNC_COLUMNS.includes(key)) {
            th.style.width = "100px";
        }
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    console.log("生成表头");


    // 生成表格内容
    const tbody = document.createElement("tbody");
    console.log("excelStudentData长度：", excelStudentData.length);
    excelStudentData.forEach((student, index) => {
        const tr = document.createElement("tr");
        EXPORT_COLUMNS.forEach(key => {
            const td = document.createElement("td");
            // 同步列采用数字输入框
            if (SYNC_COLUMNS.includes(key)) {
                td.style.width = "100px";
                td.innerHTML = `<input type="number" step="0.1" min="0" max="100" value="${student[key]}" 
                    onblur="syncScore(${index}, '${key}', this.value)">`;
            } 
            // 其他可编辑列采用文本输入框
            else if (EDITABLE_COLUMNS.includes(key)) {
                td.innerHTML = `<input type="text" value="${student[key]}" 
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
    console.log("生成表格内容");

    // 添加表格底部导出按钮
    const bottomExportArea = document.createElement("div");
    bottomExportArea.className = "export-area";
    bottomExportArea.style.marginTop = "15px";
    bottomExportArea.innerHTML = `
        <button id="bottomExportBtn">导出Excel成绩表</button>
        <button id="bottomExportTxtBtn" class="secondary">导出TXT成绩表</button>
    `;
    tableContainer.appendChild(bottomExportArea);
    console.log("添加底部导出按钮");

    // 绑定导出按钮事件
    try {
        document.getElementById("topExportBtn").addEventListener("click", exportExcelResult);
        document.getElementById("topExportTxtBtn").addEventListener("click", exportTxtResult);
        document.getElementById("bottomExportBtn").addEventListener("click", exportExcelResult);
        document.getElementById("bottomExportTxtBtn").addEventListener("click", exportTxtResult);
        console.log("绑定导出按钮事件成功");
    } catch (error) {
        console.error("绑定导出按钮事件失败：", error);
    }
    
    console.log("表格渲染完成");
}

// 更新非同步的可编辑列
function updateOtherColumn(index, column, value) {
    // 尝试将值转换为数字（用于成绩列）
    const numValue = parseFloat(value);
    if (!isNaN(numValue)) {
        // 处理小数位数
        let formattedValue = numValue;
        if (settings.scoreDecimalPlaces !== -1) {
            formattedValue = parseFloat(numValue.toFixed(settings.scoreDecimalPlaces));
        }
        excelStudentData[index][column] = formattedValue;
    } else {
        // 非数字值保持原样
        excelStudentData[index][column] = value;
    }
}

// 成绩同步
function syncScore(index, itemKey, score) {
    // 检查是否开启成绩同步
    if (!settings.syncEnabled) {
        console.log("成绩同步已关闭，仅更新个人成绩");
        const currentStudent = excelStudentData[index];
        currentStudent[itemKey] = score;
        renderExcelTable();
        return;
    }

    // 校验成绩格式
    const scoreNum = parseFloat(score);
    if (isNaN(scoreNum) || scoreNum < settings.scoreMin || scoreNum > settings.scoreMax) {
        console.log(`成绩格式错误！请输入${settings.scoreMin}-${settings.scoreMax}的数字`);
        return;
    }

    // 处理小数位数
    let formattedScore = scoreNum;
    if (settings.scoreDecimalPlaces !== -1) {
        formattedScore = parseFloat(scoreNum.toFixed(settings.scoreDecimalPlaces));
    }

    // 获取当前学生的学号
    const currentStudent = excelStudentData[index];
    const studentId = currentStudent.学号?.toString().trim(); // Excel中的学号转字符串
    const studentName = currentStudent.姓名 || studentIdToName[studentId];

    if (!studentId) {
        console.log("未找到学生学号，无法同步成绩");
        return;
    }

    // 用学号查找所属组
    const groupId = studentIdToGroup[studentId];
    
    // 先更新当前学生的成绩
    currentStudent[itemKey] = formattedScore;
    
    if (!groupId) {
        console.log(`${studentName}（学号：${studentId}）未分配小组，仅更新个人成绩`);
        // 重新渲染表格
        renderExcelTable();
        return;
    }

    // 获取同组所有学生的学号
    const sameGroupStudentIds = groupToStudentIds[groupId];
    if (!sameGroupStudentIds || sameGroupStudentIds.length === 0) {
        console.log(`组${groupId}无有效组员，同步失败`);
        return;
    }

    // 同步学生成绩
    excelStudentData.forEach(student => {
        const targetStudentId = student.学号?.toString().trim();
        if (sameGroupStudentIds.includes(targetStudentId)) {
            student[itemKey] = formattedScore;
        }
    });

    // 重新渲染表格
    renderExcelTable();
    console.log(`成绩同步完成！已将【组${groupId}】的【${itemKey}】成绩同步为：${formattedScore}分，同步学生：${sameGroupStudentIds.length}人`);
}
/*---------------------------------------------------------------------*/


/*------------------------------统计相关 ------------------------------*/
// 初始化统计部分
function initStatsSection() {
    // 初始时清空统计显示
    updateStatsDisplay();
}

// 更新统计列选择器
function updateStatsColumnSelect() {
    const select = document.getElementById('statsColumnSelect');
    select.innerHTML = '<option value="">请选择列</option>';
    
    // 添加同步列到选择器
    SYNC_COLUMNS.forEach(col => {
        const option = document.createElement('option');
        option.value = col;
        option.textContent = col;
        select.appendChild(option);
    });
}

// 更新统计显示
function updateStatsDisplay() {
    const selectedColumn = document.getElementById('statsColumnSelect').value;
    
    if (!selectedColumn || excelStudentData.length === 0) {
        // 清空统计显示
        document.getElementById('avgScore').textContent = '--';
        document.getElementById('maxScore').textContent = '--';
        document.getElementById('minScore').textContent = '--';
        document.getElementById('passRate').textContent = '--';
        
        // 清空图表
        const chartCanvas = document.createElement('canvas');
        chartCanvas.id = 'scoreChartCanvas';
        document.getElementById('scoreChart').innerHTML = '';
        document.getElementById('scoreChart').appendChild(chartCanvas);
        return;
    }
    
    // 计算统计数据
    const stats = calculateStats(selectedColumn);
    
    // 更新统计概览
    document.getElementById('avgScore').textContent = stats.avg.toFixed(2);
    document.getElementById('maxScore').textContent = stats.max;
    document.getElementById('minScore').textContent = stats.min;
    document.getElementById('passRate').textContent = (stats.passRate * 100).toFixed(1) + '%';
    
    // 生成并渲染饼图
    renderScoreChart(selectedColumn, stats.distribution);
}

// 计算统计数据
function calculateStats(column) {
    const scores = [];
    
    excelStudentData.forEach(student => {
        const score = parseFloat(student[column]);
        if (!isNaN(score)) {
            scores.push(score);
        }
    });
    
    if (scores.length === 0) {
        return {
            avg: 0,
            max: 0,
            min: 0,
            passRate: 0,
            distribution: {}
        };
    }
    
    const sum = scores.reduce((acc, score) => acc + score, 0);
    const avg = sum / scores.length;
    const max = Math.max(...scores);
    const min = Math.min(...scores);
    const passCount = scores.filter(score => score >= 60).length;
    const passRate = passCount / scores.length;
    
    // 生成成绩分布
    const distribution = generateScoreDistribution(scores);
    
    return {
        avg,
        max,
        min,
        passRate,
        distribution
    };
}

// 生成成绩分布
function generateScoreDistribution(scores) {
    const distribution = {
        '90-100': 0,
        '80-89': 0,
        '70-79': 0,
        '60-69': 0,
        '0-59': 0
    };
    
    scores.forEach(score => {
        if (score >= 90) {
            distribution['90-100']++;
        } else if (score >= 80) {
            distribution['80-89']++;
        } else if (score >= 70) {
            distribution['70-79']++;
        } else if (score >= 60) {
            distribution['60-69']++;
        } else {
            distribution['0-59']++;
        }
    });
    
    return distribution;
}

// 渲染成绩分布饼图
function renderScoreChart(column, distribution) {
    const chartContainer = document.getElementById('scoreChart');
    chartContainer.innerHTML = '';
    
    const canvas = document.createElement('canvas');
    canvas.id = 'scoreChartCanvas';
    chartContainer.appendChild(canvas);
    
    const ctx = canvas.getContext('2d');
    
    // 准备图表数据
    const labels = Object.keys(distribution);
    const data = Object.values(distribution);
    const backgroundColors = [
        'rgba(54, 162, 235, 0.7)',
        'rgba(75, 192, 192, 0.7)',
        'rgba(255, 206, 86, 0.7)',
        'rgba(153, 102, 255, 0.7)',
        'rgba(255, 99, 132, 0.7)'
    ];
    
    // 创建饼图
    new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: backgroundColors,
                borderColor: [
                    'rgba(54, 162, 235, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 99, 132, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: `${column}成绩分布`,
                    font: {
                        size: 16
                    }
                },
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}
/*-------------------------------------------------------------------*/


/*------------------------------导出相关 -----------------------------*/
// 生成文件名
function generateFileName(extension) {
    let fileName = settings.fileNameTemplate;
    // 替换日期占位符
    const date = new Date().toLocaleDateString();
    fileName = fileName.replace('{date}', date);
    return `${fileName}.${extension}`;
}

// 导出成绩表
function exportExcelResult() {
    if (excelStudentData.length === 0) {
        console.log("请先导入Excel成绩模板！");
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
    XLSX.writeFile(workbook, generateFileName('xlsx'));
}

// 导出为制表符分隔的TXT格式
function exportTxtResult() {
    if (excelStudentData.length === 0) {
        console.log("请先导入成绩模板！");
        return;
    }
    if (EXPORT_COLUMNS.length === 0) {
        console.log("请先配置导出列！");
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
    a.download = generateFileName('txt');
    a.click();
    URL.revokeObjectURL(url);

    console.log("TXT成绩表导出成功！");
}
/*-------------------------------------------------------------------*/