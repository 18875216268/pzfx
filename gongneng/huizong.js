// 汇总数据模块 - 优化版本

import { getAllDrugInfo, calculateRate, calculateYoYGrowth, calculateRateDifference } from './chuli.js';

// 数据字段配置
const DATA_FIELDS = [
    '含税出库金额', 'P4毛利额', 'P4毛利率', 
    '应收边际利润额(不含税)', '边际利润率', '客户数'
];

/**
 * 创建汇总工作簿
 * @param {Object} aggregatedData - 汇总后的数据
 * @returns {Object} ExcelJS工作簿对象
 */
export async function createSummaryWorkbook(aggregatedData) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('汇总');
    
    // 准备数据
    const allDrugInfo = getAllDrugInfo(aggregatedData);
    const drugIds = Object.keys(allDrugInfo).sort();
    const fileInfo = prepareFileInfo(aggregatedData);
    
    // 创建表头
    createHeaders(worksheet, fileInfo);
    
    // 填充数据
    drugIds.forEach(drugId => {
        const rowData = buildDataRow(drugId, allDrugInfo[drugId], fileInfo);
        worksheet.addRow(rowData);
    });
    
    // 应用样式
    applyStyles(worksheet, fileInfo.files.length);
    
    return workbook;
}

/**
 * 准备文件信息，按同期、上期、当期排序
 */
function prepareFileInfo(aggregatedData) {
    const files = [];
    const typeOrder = { '同期': 0, '上期': 1, '当期': 2 };
    
    // 收集文件信息
    Object.entries(aggregatedData).forEach(([fileName, data]) => {
        // 更严格的类型识别
        let type = null;
        if (fileName.includes('同期')) type = '同期';
        else if (fileName.includes('上期')) type = '上期';
        else if (fileName.includes('当期')) type = '当期';
        
        if (type) {
            files.push({ 
                fileName, 
                data, 
                type, 
                order: typeOrder[type]
            });
        }
    });
    
    // 确保按照同期(0)、上期(1)、当期(2)的顺序排序
    files.sort((a, b) => {
        // 首先按order排序
        if (a.order !== b.order) {
            return a.order - b.order;
        }
        // 如果order相同，按文件名排序（备用）
        return a.fileName.localeCompare(b.fileName);
    });
    
    // 验证排序结果（可选的调试代码）
    console.log('文件排序结果:', files.map(f => `${f.type}(${f.order})`).join(' -> '));
    
    return { 
        files,
        dataByType: Object.fromEntries(files.map(f => [f.type, f.data]))
    };
}

/**
 * 创建表头（两行）
 */
function createHeaders(worksheet, fileInfo) {
    // 第一行：合并单元格标题
    const row1 = ['药品基础信息', '', '', '', ''];
    fileInfo.files.forEach(f => {
        row1.push(f.fileName.replace(/\.[^/.]+$/, ''), ...Array(5).fill(''));
    });
    row1.push('同比', ...Array(5).fill(''), '环比', ...Array(5).fill(''));
    worksheet.addRow(row1);
    
    // 第二行：字段名
    const row2 = ['品种负责人', '药品ID', '商品名称', '商品规格', '生产厂家'];
    const repeatCount = fileInfo.files.length + 2;
    for (let i = 0; i < repeatCount; i++) {
        row2.push(...DATA_FIELDS);
    }
    worksheet.addRow(row2);
}

/**
 * 构建数据行
 */
function buildDataRow(drugId, drugInfo, fileInfo) {
    // 基础信息
    const row = [
        drugInfo['品种负责人'] || '',
        drugId,
        drugInfo['商品名称'] || '',
        drugInfo['商品规格'] || '',
        drugInfo['生产厂家'] || ''
    ];
    
    // 各期数据 - 按照files数组的顺序填充
    const periodData = {};
    fileInfo.files.forEach(file => {
        const data = file.data[drugId] || {};
        periodData[file.type] = data;
        row.push(...formatDrugData(data));
    });
    
    // 同比（当期 vs 同期）
    row.push(...calculateComparison(periodData['当期'], periodData['同期']));
    
    // 环比（当期 vs 上期）
    row.push(...calculateComparison(periodData['当期'], periodData['上期']));
    
    return row;
}

/**
 * 格式化药品数据
 */
function formatDrugData(data) {
    const amount = data['含税出库金额'] || 0;
    const p4Profit = data['P4毛利额'] || 0;
    const marginProfit = data['应收边际利润额(不含税)'] || 0;
    
    return [
        amount,
        p4Profit,
        calculateRate(p4Profit, amount),
        marginProfit,
        calculateRate(marginProfit, amount),
        data['客户数'] || 0
    ];
}

/**
 * 计算比较数据
 */
function calculateComparison(current = {}, previous = {}) {
    if (!current || !previous) return Array(6).fill(0);
    
    const curr = formatDrugData(current);
    const prev = formatDrugData(previous);
    
    return [
        calculateYoYGrowth(curr[0], prev[0]),  // 金额增长率
        calculateYoYGrowth(curr[1], prev[1]),  // P4毛利额增长率
        calculateRateDifference(curr[2], prev[2]),  // P4毛利率差异
        calculateYoYGrowth(curr[3], prev[3]),  // 边际利润额增长率
        calculateRateDifference(curr[4], prev[4]),  // 边际利润率差异
        calculateYoYGrowth(curr[5], prev[5])   // 客户数增长率
    ];
}

/**
 * 应用样式和格式
 */
function applyStyles(worksheet, fileCount) {
    // 设置合并单元格
    worksheet.mergeCells('A1:E1');
    let startCol = 6;
    for (let i = 0; i < fileCount + 2; i++) {
        worksheet.mergeCells(1, startCol, 1, startCol + 5);
        startCol += 6;
    }
    
    // 设置样式
    const headerStyle = {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BBB59' } },
        font: { color: { argb: 'FFFFFFFF' }, bold: true },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        }
    };
    
    // 应用到前两行
    [1, 2].forEach(rowNum => {
        worksheet.getRow(rowNum).eachCell(cell => Object.assign(cell, headerStyle));
    });
    worksheet.getRow(1).height = 25;
    
    // 设置列宽
    worksheet.columns.forEach(col => col.width = 12);
    
    // 设置数字格式
    const totalCols = 5 + (fileCount + 2) * 6;
    for (let row = 3; row <= worksheet.rowCount; row++) {
        for (let col = 6; col <= totalCols; col++) {
            const cell = worksheet.getCell(row, col);
            const header = worksheet.getCell(2, col).value;
            const isComparison = col >= 6 + fileCount * 6;
            
            if (header?.includes('率')) {
                cell.numFmt = '0.00%';
            } else if (isComparison && !header?.includes('率')) {
                cell.numFmt = '0.00%';
            } else if (typeof cell.value === 'number') {
                cell.numFmt = '#,##0.00';
            }
        }
    }
}