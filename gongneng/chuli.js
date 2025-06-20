// 处理数据模块 - 优化版本

// 数值字段列表
const NUMERIC_FIELDS = [
    '销售数量', '客户数', '含税出库金额', 
    'P4成本金额', 'P4毛利额', '应收边际利润额(不含税)'
];

/**
 * 处理和汇总数据
 * @param {Map} extractedDataMap - 提取后的数据映射
 * @returns {Object} 汇总后的数据对象
 */
export function processAndAggregateData(extractedDataMap) {
    const result = {};
    
    for (const [fileName, fileData] of extractedDataMap) {
        try {
            result[fileName] = aggregateByDrugId(fileData.rows);
            console.log(`${fileName} 汇总完成：${Object.keys(result[fileName]).length} 种药品`);
        } catch (error) {
            throw new Error(`处理文件 ${fileName} 失败：${error.message}`);
        }
    }
    
    return result;
}

/**
 * 按药品ID汇总数据
 * @param {Array} rows - 数据行数组
 * @returns {Object} 按药品ID汇总的数据
 */
function aggregateByDrugId(rows) {
    return rows.reduce((acc, row) => {
        const drugId = String(row['药品ID']);
        if (!drugId) return acc;
        
        if (!acc[drugId]) {
            // 初始化：复制文本字段，数值字段置零
            acc[drugId] = Object.fromEntries(
                Object.entries(row).map(([key, value]) => 
                    [key, NUMERIC_FIELDS.includes(key) ? 0 : value]
                )
            );
        }
        
        // 累加数值字段
        const item = acc[drugId];
        NUMERIC_FIELDS.forEach(field => {
            item[field] += parseFloat(row[field]) || 0;
        });
        
        // 更新文本字段（保持最新值）
        Object.keys(row).forEach(field => {
            if (!NUMERIC_FIELDS.includes(field) && row[field]) {
                item[field] = row[field];
            }
        });
        
        return acc;
    }, {});
}

/**
 * 获取所有药品的基础信息
 * @param {Object} aggregatedData - 汇总后的数据
 * @returns {Object} 所有药品的基础信息
 */
export function getAllDrugInfo(aggregatedData) {
    const allDrugInfo = {};
    
    Object.values(aggregatedData).forEach(fileData => {
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

/**
 * 计算工具函数
 */
export const calculateRate = (numerator, denominator) => 
    denominator ? numerator / denominator : 0;

export const calculateYoYGrowth = (current, previous) => 
    previous ? (current - previous) / previous : 0;

export const calculateMoMGrowth = calculateYoYGrowth;

export const calculateRateDifference = (current, previous) => 
    current - previous;