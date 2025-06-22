// 计算模块
const CalculateModule = {
    formatDrugData(data = {}) {
        const amount = data['含税出库金额'] || 0;
        const p4Profit = data['P4毛利额'] || 0;
        const marginProfit = data['应收边际利润额(不含税)'] || 0;
        
        return [
            amount,
            p4Profit,
            amount ? p4Profit / amount : 0,
            marginProfit,
            amount ? marginProfit / amount : 0,
            data['客户数'] || 0
        ];
    },
    
    calculateComparison(currentData = {}, previousData = {}) {
        const curr = this.formatDrugData(currentData);
        const prev = this.formatDrugData(previousData);
        
        return [
            prev[0] ? (curr[0] - prev[0]) / prev[0] : 0,
            prev[1] ? (curr[1] - prev[1]) / prev[1] : 0,
            curr[2] - prev[2],
            prev[3] ? (curr[3] - prev[3]) / prev[3] : 0,
            curr[4] - prev[4],
            prev[5] ? (curr[5] - prev[5]) / prev[5] : 0
        ];
    },
    
    buildDataRow(groupKey, drugInfo, periodData, basicInfoFields, matchConfig, matchFileKeyMap) {
        const row = [];
        
        basicInfoFields.forEach(field => {
            row.push(drugInfo[field] || '');
        });
        
        const periods = ['同期', '上期', '当期'];
        periods.forEach(period => {
            const data = periodData[period]?.[groupKey] || {};
            row.push(...this.formatDrugData(data));
        });
        
        const currentData = periodData['当期']?.[groupKey] || {};
        const previousYearData = periodData['同期']?.[groupKey] || {};
        const previousPeriodData = periodData['上期']?.[groupKey] || {};
        
        row.push(...this.calculateComparison(currentData, previousYearData));
        row.push(...this.calculateComparison(currentData, previousPeriodData));
        
        if (matchConfig && matchFileKeyMap && matchFileKeyMap.has(groupKey)) {
            const matchedRow = matchFileKeyMap.get(groupKey);
            matchConfig.fieldsToAdd.forEach(field => {
                row.push(matchedRow[field.index] || '');
            });
        }
        
        return { row, matchKey: groupKey };
    }
};