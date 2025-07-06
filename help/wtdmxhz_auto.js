// 要求
// 输入 型号 和 挤压批号 查询资料
// 查表格有没有不合格的（会有粉色背景）
// 导出不合格的数据为 Excel

// ！！！注意
// 有时候行，有时候不行
// 搜索了数据后才 inspect element 后面一定会行

// 选择 型号
function selectModelCode(modelCode) {
    // zbm 是 型号 input 的 id
    const selectElement = document.getElementById('zbm');

    // 检查 型号 是否存在在选项之中
    const optionExists = Array.from(selectElement.options).some(option => option.value === modelCode);
    
    if (optionExists) {
        // Set the value
        selectElement.value = modelCode;
        
        // Trigger change events
        const event = new Event('change', { bubbles: true });
        selectElement.dispatchEvent(event);
        
        console.log(`✅ 成功：选择型号 ${modelCode}`);
    } else {
        console.log(`❌ 错误: 从选项中找不到型号 ${modelCode}`);
    }
}

// 输入 挤压批号
function enterExtrusionBatchCode(extrusionBatchCode) {
    // jyno 是挤压批号的输入元素
    const inputElement = document.getElementById('jyno');
    
    if (inputElement) {
        // 设置数值
        inputElement.value = extrusionBatchCode;

        // 触发变更事件 来更新DOM / 网页
        inputElement.dispatchEvent(new Event('input', { bubbles: true }));
        inputElement.dispatchEvent(new Event('change', { bubbles: true }));

        console.log(`✅ 成功：输入挤压批号 ${extrusionBatchCode}`);
    } else {
        console.error(`❌ 失败: 找不到挤压批号的输入元素`);
    }
}

const ExportTypes = Object.freeze({
    CSV: 'csv',
    TXT: 'txt',
    WORD: 'doc',
    EXCEL: 'excel',
});

// 导出数据
function exportData(type=ExportTypes.EXCEL) {
    // 找到并点击数据导出按键
    const exportButton = document.querySelector('button[title="Export data"]');
    if (!exportButton) {
        console.error('❌ 错误：找不到数据导出按键');
    } else {
        // 点击数据导出按键来打开菜单
        exportButton.click();

        // 从菜单里找出想导出为的选项
        const excelOption = document.querySelector(`li[data-type="${type}"]`);
        
        if (!excelOption) {
            console.error(`❌ 错误：找不到 ${type} 的选项`);
        } else {
            // 点击想导出为的选项
            excelOption.click();
            console.log(`开始导出数据为 ${type}`);
        }
    }
}

// 点击 资料查询的表格
function clickSearchButton() {
    // 在 #formId 内查找带有文本 “查询 ”的准确链接
    const links = document.querySelectorAll('#formId a.btn.btn-primary');
    const targetLink = Array.from(links).find(link => 
        link.textContent.includes('查询')
    );

    if (targetLink) {
        targetLink.click();
        console.log("✅ 成功：点击查询按键，开始查询资料");
    } else {
        console.error("❌ 失败：找不到查询按键");
    }
}

// 找包含不合格性能数值的列名
function findHeadersForPinkSpans() {
    const results = [];
    
    // 寻找全部粉色的span
    const pinkSpans = document.querySelectorAll('span[style="background-color: pink"]');
    
    if (pinkSpans.length === 0) {
        console.log("✅ 全部合格");
        return results;
    }

    // Loop through each matching span
    pinkSpans.forEach(span => {
        const td = span.closest('td');
        if (!td) {
            console.log("Span is not inside a table cell");
            return; // Skip to next span
        }

        const columnIndex = td.cellIndex;
        const table = td.closest('table');
        if (!table) {
            console.log("Table not found for span");
            return;
        }

        // Get the header row (supports both thead and first tr if no thead exists)
        const headerRow = table.querySelector('thead tr') || table.querySelector('tr');
        if (!headerRow) {
            console.log("Header row not found");
            return;
        }

        const headerTh = headerRow.querySelectorAll('th, td')[columnIndex];
        if (!headerTh) {
            console.log("Corresponding header not found");
            return;
        }

        const headerText = headerTh.textContent.trim();
        if (!results.includes(headerText)) { // Avoid duplicates
            results.push(headerText);
        }
    });

    return results;
}

async function processRow(modelCode, extrusionBatchCode) {
    // Implementation depends on how the table updates
    // This example waits for the table to contain more than just headers
    
    const maxWaitTime = 30000; // 30 seconds max
    const checkInterval = 500; // check every 500ms
    const startTime = Date.now();

    try {
        // Execute steps sequentially
        selectModelCode(modelCode);
        enterExtrusionBatchCode(extrusionBatchCode);
        clickSearchButton();

        setTimeout(() => {
            // 检查有没有匹配的记录
            if (noRecordsFound()) {
                console.log(`❌ 没有找到 ${modelCode} ${extrusionBatchCode} 匹配的记录`)
                return;
            }
            
            // Wait for results to load
            const headerNames = findHeadersForPinkSpans();
            
            if (headerNames.length !== 0) {
                console.log('❌ 以下性能包含不合格的数值：', headerNames);
                console.log(`${extrusionBatchCode}_${headerNames.join('_')}`)
                exportData();
            }
        }, 2000);

    } catch (error) {
        console.error(`处理型号 ${modelCode} 时出现错误:`, error);
    }
}

function noRecordsFound() {
    const noRecordsFoundElement = document.querySelectorAll('.no-records-found');
    return noRecordsFoundElement.length > 0;
}

// 运行并打印出结果
async function main() {
    data = prompt("粘贴想处理的 型号 和 内部批号：");

    if (data) {
        const rows = data.split("\r\n");
        for (const row of rows) {
            const cells = row.split("\t");
            const modelCode = cells[0];
            const extrusionBatchCode = cells[1];

            // Process current row and wait for completion
            await processRow(modelCode, extrusionBatchCode);
        }
    } else {
        console.log('❌ 无数据填入，放弃查询')
    }
}

main();
