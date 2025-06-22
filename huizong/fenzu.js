// 分组模块
const GroupModule = {
    generateGroupKey(rowData, groupFields) {
        if (!groupFields || groupFields.length === 0) {
            return String(rowData[0] || '').trim();
        }
        
        const keyParts = groupFields.map(field => 
            String(rowData[field] || '').trim()
        ).filter(part => part);
        
        return keyParts.length > 0 ? keyParts.join('|') : '';
    },
    
    processWorksheet(worksheet, headerRowNum, allFields, groupFields) {
        const headerRow = worksheet.getRow(headerRowNum);
        const fieldMap = new Map();
        
        headerRow.eachCell((cell, colNum) => {
            const header = String(cell.value || '').trim();
            if (header && allFields.includes(header)) {
                fieldMap.set(header, colNum);
            }
        });
        
        const aggregated = {};
        for (let rowNum = headerRowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            
            const rowData = {};
            for (const [field, colNum] of fieldMap) {
                const value = row.getCell(colNum).value;
                rowData[field] = CONFIG.numericFields.includes(field) 
                    ? utils.parseNumber(value)
                    : String(value || '').trim();
            }
            
            const groupKey = this.generateGroupKey(rowData, groupFields);
            if (groupKey) {
                if (!aggregated[groupKey]) {
                    aggregated[groupKey] = { ...rowData };
                } else {
                    CONFIG.numericFields.forEach(field => {
                        if (rowData[field] !== undefined) {
                            aggregated[groupKey][field] = (aggregated[groupKey][field] || 0) + (rowData[field] || 0);
                        }
                    });
                }
            }
        }
        
        return aggregated;
    }
};