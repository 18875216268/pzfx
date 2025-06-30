/**
 * 重构后的 FillModule - 直接使用统一的显示格式，大幅简化代码
 */
const FillModule = {
    /**
     * 工作簿创建主函数 - 直接使用显示格式，无需格式转换
     */
    createSummaryWorkbook(processedData) {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('汇总');
        const files = Object.keys(processedData);
        const groups = AggregateModule.getAllGroupKeys(processedData);
        const { basicInfo, groupData } = AggregateModule.extractGroupAndBasicInfo(processedData);
        
        // 直接使用显示格式创建表头
        const groupHeaders = state.groupFields.filter(name => name.endsWith('.分组'));
        const infoHeaders = state.groupFields.filter(name => name.endsWith('.信息'));
        
        const headers = [
            // 基础字段直接使用显示名称
            ...groupHeaders,
            ...infoHeaders,
            // 聚合和计算字段加上工作簿前缀
            ...files.flatMap(fileName => {
                const shortName = fileName.replace(/\.[^/.]+$/, '');
                
                // 聚合字段：工作簿名称.显示名称
                const aggregateHeaders = state.aggregateFields.map(displayName => 
                    `${shortName}.${displayName}`
                );
                
                // 计算字段：工作簿名称.字段名称.计算
                const calcHeaders = state.calcFields.map(field => 
                    `${shortName}.${field.name}.计算`
                );
                
                return [...aggregateHeaders, ...calcHeaders];
            })
        ];
        
        ws.addRow(headers);
        
        // 直接填充数据行
        groups.forEach(groupKey => {
            const row = [
                // 分组和信息字段数据
                ...groupHeaders.map(displayName => {
                    const field = displayName.split('.')[0];
                    return groupData[groupKey]?.[field] || '';
                }),
                ...infoHeaders.map(displayName => {
                    const field = displayName.split('.')[0];
                    return basicInfo[groupKey]?.[field] || '';
                }),
                // 聚合和计算字段数据
                ...files.flatMap(fileName => {
                    const fileData = processedData[fileName]?.[groupKey];
                    if (!fileData) {
                        // 如果没有数据，返回对应数量的0
                        return new Array(state.aggregateFields.length + state.calcFields.length).fill(0);
                    }
                    
                    // 聚合字段数据
                    const aggregateData = state.aggregateFields.map(displayName => 
                        fileData.data[displayName] || 0
                    );
                    
                    // 计算字段数据
                    const calcData = state.calcFields.map(field => 
                        fileData.data[`${field.name}.计算`] || 0
                    );
                    
                    return [...aggregateData, ...calcData];
                })
            ];
            ws.addRow(row);
        });
        
        // 应用样式
        ws.columns.forEach(c => c.width = 12);
        ws.getRow(1).height = 45;
        
        const baseStyle = {
            font: {name: '宋体', size: 10}, 
            border: {
                top: {style: 'thin'}, left: {style: 'thin'}, 
                bottom: {style: 'thin'}, right: {style: 'thin'}
            }
        };
        
        const headerStyle = {
            fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFBFBFBF'}}, 
            alignment: {horizontal: 'center', vertical: 'middle', wrapText: true}
        };
        
        ws.eachRow((row, rowIndex) => {
            row.eachCell(cell => {
                Object.assign(cell, baseStyle);
                
                if (rowIndex === 1) {
                    Object.assign(cell, headerStyle);
                } else if (cell.col > groupHeaders.length + infoHeaders.length && typeof cell.value === 'number') {
                    const headerText = ws.getCell(1, cell.col).value;
                    cell.numFmt = headerText?.includes('率') || headerText?.includes('占比') ? '0.00%' : '#,##0.00';
                }
            });
        });
        
        return wb;
    }
};