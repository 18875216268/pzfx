// 聚合模块
const AggregateModule = {
    prepareFileInfo(processedData) {
        const files = [];
        const typeOrder = { '同期': 0, '上期': 1, '当期': 2 };
        
        Object.entries(processedData).forEach(([fileName, data]) => {
            const type = utils.getFileType(fileName);
            if (type) {
                files.push({ fileName, data, type, order: typeOrder[type] });
            }
        });
        
        return files.sort((a, b) => a.order - b.order);
    },
    
    aggregateByPeriod(fileInfo) {
        const periodData = {};
        fileInfo.forEach(file => periodData[file.type] = file.data);
        return periodData;
    },
    
    getAllGroupKeys(processedData) {
        const groupKeys = new Set();
        Object.values(processedData).forEach(fileData => {
            Object.keys(fileData).forEach(groupKey => groupKeys.add(groupKey));
        });
        return Array.from(groupKeys).sort();
    },
    
    extractAllDrugInfo(processedData, basicInfoFields) {
        const drugInfo = {};
        
        Object.values(processedData).forEach(fileData => {
            Object.entries(fileData).forEach(([groupKey, data]) => {
                if (!drugInfo[groupKey]) {
                    drugInfo[groupKey] = {};
                    basicInfoFields.forEach(field => {
                        drugInfo[groupKey][field] = data[field] || '';
                    });
                }
            });
        });
        
        return drugInfo;
    }
};