// 聚合模块
const AggregateModule = {
    prepareFileInfo(processedData) {
        const typeOrder = { '同期': 0, '上期': 1, '当期': 2 };
        
        return Object.entries(processedData)
            .map(([fileName, data]) => {
                const type = utils.getFileType(fileName);
                return type ? { fileName, data, type, order: typeOrder[type] } : null;
            })
            .filter(Boolean)
            .sort((a, b) => a.order - b.order);
    },
    
    aggregateByPeriod(fileInfo) {
        return fileInfo.reduce((acc, file) => {
            acc[file.type] = file.data;
            return acc;
        }, {});
    },
    
    getAllGroupKeys(processedData) {
        const groupKeys = new Set();
        Object.values(processedData).forEach(fileData => {
            Object.keys(fileData).forEach(key => groupKeys.add(key));
        });
        return Array.from(groupKeys).sort();
    },
    
    extractAllBasicInfo(periodData, basicInfoFields) {
        const basicInfo = {};
        
        ['当期', '上期', '同期'].forEach(period => {
            if (periodData[period]) {
                Object.entries(periodData[period]).forEach(([groupKey, data]) => {
                    if (!basicInfo[groupKey] && data.basicInfo) {
                        basicInfo[groupKey] = { ...data.basicInfo };
                    }
                });
            }
        });
        
        return basicInfo;
    }
};