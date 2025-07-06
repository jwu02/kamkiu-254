// 要求
// 输入 型号 和 挤压批号 查询资料
// 查表格有没有不合格的（会有粉色背景）
// 导出不合格的数据为 Excel

// ！！！注意
// 有时候行，有时候不行
// 搜索了数据后才 inspect element 后面一定会行

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

        const headerText = '_'+headerTh.textContent.trim();
        if (!results.includes(headerText)) { // Avoid duplicates
            results.push(headerText);
        }
    });

    return results;
}

function noRecordsFound() {
    const noRecordsFoundElement = document.querySelectorAll('.no-records-found');
    return noRecordsFoundElement.length > 0;
}

// 运行并打印出结果
function main() {
    // Wait for results to load
    const headerNames = findHeadersForPinkSpans();
    
    if (headerNames.length !== 0) {
        console.log('❌ 以下性能包含不合格的数值：', headerNames);
        console.log(headerNames.join(''))
        exportData();
    }
}

main();
