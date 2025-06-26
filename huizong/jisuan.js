// 计算模块
const CalculateModule = {
    calculateFields(data, calcFields) {
        const result = { ...data };
        Object.entries(calcFields).forEach(([fieldName, expression]) => {
            result[fieldName] = utils.parseExpression(expression, data);
        });
        return result;
    },
    
    formatRowData(data = {}, aggregateFields, calcFields) {
        return [
            ...Object.keys(aggregateFields).map(field => data[field] || 0),
            ...Object.keys(calcFields).map(field => data[field] || 0)
        ];
    },
    
    calculateComparison(currentData = {}, previousData = {}, aggregateFields, calcFields) {
        const allFields = [...Object.keys(aggregateFields), ...Object.keys(calcFields)];
        
        return allFields.map(field => {
            const currVal = utils.parseNumber(currentData[field] || 0);
            const prevVal = utils.parseNumber(previousData[field] || 0);
            
            // 率类字段：计算差值，其他字段：计算增长率
            return field.includes('率') ? 
                currVal - prevVal : 
                (prevVal ? (currVal - prevVal) / prevVal : 0);
        });
    }
};