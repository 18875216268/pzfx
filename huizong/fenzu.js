// 分组模块
const GroupModule = {
    generateGroupKey(rowData, groupFields) {
        if (!groupFields.length) {
            const firstField = Object.keys(rowData)[0];
            return String(rowData[firstField] || '').trim();
        }
        return groupFields
            .map(field => String(rowData[field] || '').trim())
            .filter(Boolean)
            .join('|');
    },
    
    processWorksheet(worksheet, headerRowNum, allFields, groupFields, aggregateFields) {
        // 构建字段映射
        const fieldMap = new Map();
        worksheet.getRow(headerRowNum).eachCell((cell, colNum) => {
            const header = String(cell.value || '').trim();
            if (header && allFields.includes(header)) {
                fieldMap.set(header, colNum);
            }
        });
        
        // 聚合数据
        const aggregated = {};
        
        for (let rowNum = headerRowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const rowData = {};
            
            // 读取所有字段数据
            fieldMap.forEach((colNum, field) => {
                rowData[field] = row.getCell(colNum).value;
            });
            
            const groupKey = this.generateGroupKey(rowData, groupFields);
            
            if (groupKey) {
                if (!aggregated[groupKey]) {
                    // 初始化分组
                    aggregated[groupKey] = {
                        basicInfo: {},
                        data: {}
                    };
                    
                    // 保存基础信息
                    state.basicInfoFields.forEach(field => {
                        if (field in rowData) {
                            aggregated[groupKey].basicInfo[field] = String(rowData[field] || '').trim();
                        }
                    });
                    
                    // 初始化聚合字段
                    Object.keys(aggregateFields).forEach(field => {
                        aggregated[groupKey].data[field] = {
                            sum: 0,
                            count: 0,
                            values: []
                        };
                    });
                }
                
                // 聚合数据
                Object.entries(aggregateFields).forEach(([field, method]) => {
                    if (field in rowData) {
                        const value = utils.parseNumber(rowData[field]);
                        const data = aggregated[groupKey].data[field];
                        
                        if (method === 'count') {
                            data.count++;
                        } else {
                            data.sum += value;
                            data.values.push(value);
                        }
                    }
                });
            }
        }
        
        // 计算最终值
        Object.values(aggregated).forEach(group => {
            Object.entries(aggregateFields).forEach(([field, method]) => {
                const data = group.data[field];
                group.data[field] = method === 'sum' ? data.sum : 
                                   method === 'avg' ? (data.values.length ? data.sum / data.values.length : 0) : 
                                   data.count;
            });
        });
        
        return aggregated;
    }
};