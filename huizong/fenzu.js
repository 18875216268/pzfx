// 分组模块 - 重构：直接使用显示格式，简化数据处理
const GroupModule = {
    processWorksheet(worksheet, headerRowNum) {
        const fieldMap = new Map();
        worksheet.getRow(headerRowNum).eachCell((cell, colNum) => {
            const header = String(cell.value || '').trim();
            if (header && state.allFields.includes(header)) {
                fieldMap.set(header, colNum);
            }
        });
        
        const aggregated = {};
        
        // 直接使用显示格式解析字段
        const groupFields = state.groupFields
            .filter(displayName => displayName.endsWith('.分组'))
            .map(displayName => displayName.split('.')[0]);
            
        const infoFields = state.groupFields
            .filter(displayName => displayName.endsWith('.信息'))
            .map(displayName => displayName.split('.')[0]);
        
        for (let rowNum = headerRowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const rowData = {};
            fieldMap.forEach((colNum, field) => {
                rowData[field] = row.getCell(colNum).value;
            });
            
            // 直接生成分组键
            const groupKey = groupFields
                .map(field => String(rowData[field] || '').trim())
                .filter(Boolean)
                .join('|');
            
            if (!groupKey) continue;
            
            if (!aggregated[groupKey]) {
                aggregated[groupKey] = {
                    basicInfo: {},
                    data: {},
                    groupData: {}
                };
                
                // 直接填充基础信息和分组数据
                infoFields.forEach(field => {
                    if (field in rowData) {
                        aggregated[groupKey].basicInfo[field] = String(rowData[field] || '').trim();
                    }
                });
                
                groupFields.forEach(field => {
                    if (field in rowData) {
                        aggregated[groupKey].groupData[field] = String(rowData[field] || '').trim();
                    }
                });
                
                // 初始化聚合数据 - 使用显示名称作为键
                state.aggregateFields.forEach(displayName => {
                    aggregated[groupKey].data[displayName] = { sum: 0, count: 0, values: [] };
                });
            }
            
            // 聚合计算 - 使用显示名称
            state.aggregateFields.forEach(displayName => {
                const field = displayName.split('.')[0];
                const method = state.parseDisplayName(displayName, 'aggregate').method;
                
                if (!(field in rowData)) return;
                
                const value = utils.parseNumber(rowData[field]);
                const data = aggregated[groupKey].data[displayName];
                
                if (method === 'count') {
                    data.count++;
                } else {
                    data.sum += value;
                    data.values.push(value);
                }
            });
        }
        
        // 计算最终值 - 直接使用显示名称
        Object.values(aggregated).forEach(group => {
            state.aggregateFields.forEach(displayName => {
                const method = state.parseDisplayName(displayName, 'aggregate').method;
                const data = group.data[displayName];
                
                switch (method) {
                    case 'sum':
                        group.data[displayName] = data.sum;
                        break;
                    case 'avg':
                        group.data[displayName] = data.values.length ? data.sum / data.values.length : 0;
                        break;
                    case 'count':
                        group.data[displayName] = data.count;
                        break;
                    case 'max':
                        group.data[displayName] = data.values.length ? Math.max(...data.values) : 0;
                        break;
                    case 'min':
                        group.data[displayName] = data.values.length ? Math.min(...data.values) : 0;
                        break;
                }
            });
        });
        
        return aggregated;
    }
};