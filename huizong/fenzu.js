// 基础数据提取、分组填充模块

// Excel处理类的部分功能
class ExcelProcessor {
    // 读取Excel文件
    static async readFile(file) {
        const workbook = new ExcelJS.Workbook();
        const arrayBuffer = await file.arrayBuffer();
        await workbook.xlsx.load(arrayBuffer);
        return workbook;
    }
    
    // 查找包含ID的标题行
    static findHeaderRow(worksheet, maxRows = 5) {
        for (let rowNum = 1; rowNum <= Math.min(maxRows, worksheet.rowCount); rowNum++) {
            const row = worksheet.getRow(rowNum);
            for (let colNum = 1; colNum <= row.cellCount; colNum++) {
                const cell = row.getCell(colNum);
                if (String(cell.value || '').toUpperCase().includes('ID')) {
                    return {
                        rowNum,
                        idColumn: colNum,
                        headers: row.values.slice(1).map(v => String(v || '').trim())
                    };
                }
            }
        }
        throw new Error('未找到包含"ID"的标题行');
    }
    
    // 处理工作表数据
    static processWorksheet(worksheet, headerRowNum, keepFields) {
        const headerRow = worksheet.getRow(headerRowNum);
        const headers = [];
        const fieldMap = new Map();
        
        // 构建字段映射
        headerRow.eachCell((cell, colNum) => {
            const header = String(cell.value || '').trim();
            headers.push(header);
            if (keepFields.includes(header)) {
                fieldMap.set(header, colNum);
            }
        });
        
        // 提取数据
        const rows = [];
        for (let rowNum = headerRowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const drugId = row.getCell(fieldMap.get('药品ID')).value;
            
            if (drugId) {
                const rowData = {};
                for (const [field, colNum] of fieldMap) {
                    const value = row.getCell(colNum).value;
                    rowData[field] = CONFIG.numericFields.includes(field) 
                        ? utils.parseNumber(value)
                        : String(value || '').trim();
                }
                rows.push(rowData);
            }
        }
        
        return { headers: Array.from(fieldMap.keys()), rows };
    }
    
    // 获取所有药品信息
    static getAllDrugInfo(processedData) {
        const allDrugInfo = {};
        
        Object.values(processedData).forEach(fileData => {
            Object.entries(fileData).forEach(([drugId, data]) => {
                allDrugInfo[drugId] = allDrugInfo[drugId] || {
                    '品种负责人': data['品种负责人'] || '',
                    '商品名称': data['商品名称'] || '',
                    '商品规格': data['商品规格'] || '',
                    '生产厂家': data['生产厂家'] || ''
                };
            });
        });
        
        return allDrugInfo;
    }
    
    // 准备文件信息
    static prepareFileInfo(processedData) {
        const files = [];
        const typeOrder = { '同期': 0, '上期': 1, '当期': 2 };
        
        Object.entries(processedData).forEach(([fileName, data]) => {
            const type = utils.getFileType(fileName);
            if (type) {
                files.push({ fileName, data, type, order: typeOrder[type] });
            }
        });
        
        files.sort((a, b) => a.order - b.order);
        
        return { 
            files,
            dataByType: Object.fromEntries(files.map(f => [f.type, f.data]))
        };
    }
}