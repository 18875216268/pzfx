// 聚合模块 - 重构：直接使用显示格式，简化数据提取
const AggregateModule = {
    getAllGroupKeys(processedData) {
        const keys = new Set();
        Object.values(processedData).forEach(data => 
            Object.keys(data).forEach(key => keys.add(key))
        );
        return Array.from(keys).sort();
    },
    
    // 提取分组和基础信息 - 直接使用字段名
    extractGroupAndBasicInfo(processedData) {
        const basicInfo = {};
        const groupData = {};
        
        for (const fileData of Object.values(processedData)) {
            for (const [groupKey, data] of Object.entries(fileData)) {
                if (!basicInfo[groupKey] && data.basicInfo) {
                    basicInfo[groupKey] = { ...data.basicInfo };
                }
                if (!groupData[groupKey] && data.groupData) {
                    groupData[groupKey] = { ...data.groupData };
                }
            }
        }
        
        return { basicInfo, groupData };
    }
};