// 提取数据模块 - 优化版本

// 配置
const CONFIG = {
    keepFields: [
        "药品ID", "商品名称", "商品规格", "生产厂家",
        "销售数量", "客户数", "含税出库金额", "P4成本金额",
        "P4毛利额", "应收边际利润额(不含税)", "品种负责人"
    ],
    numericFields: [
        "销售数量", "客户数", "含税出库金额", 
        "P4成本金额", "P4毛利额", "应收边际利润额(不含税)"
    ]
};

/**
 * 从所有文件中提取需要的字段
 * @param {Map} fileDataMap - 文件数据映射
 * @returns {Map} 提取后的数据
 */
export function extractRequiredFields(fileDataMap) {
    const extractedDataMap = new Map();
    
    for (const [fileName, fileData] of fileDataMap) {
        try {
            const extracted = extractFieldsFromFile(fileData);
            extractedDataMap.set(fileName, extracted);
            console.log(`文件 ${fileName} 提取完成：${extracted.rowCount} 行数据`);
        } catch (error) {
            throw new Error(`提取文件 ${fileName} 失败：${error.message}`);
        }
    }
    
    // 验证数据
    if (!Array.from(extractedDataMap.values()).some(data => data.rowCount > 0)) {
        throw new Error('所有文件都没有有效数据');
    }
    
    return extractedDataMap;
}

/**
 * 从单个文件提取字段
 * @param {Object} fileData - 文件数据
 * @returns {Object} 提取结果
 */
function extractFieldsFromFile(fileData) {
    const { data, headers, headerRowNum, fileName } = fileData;
    
    // 验证数据
    if (!data || data.length <= headerRowNum) {
        throw new Error(`数据不足，需要至少 ${headerRowNum + 1} 行`);
    }
    
    // 创建字段映射
    const fieldMap = new Map();
    headers.forEach((header, index) => {
        if (CONFIG.keepFields.includes(header)) {
            fieldMap.set(header, index);
        }
    });
    
    // 提取数据行
    const rows = data.slice(headerRowNum)
        .map(row => {
            const newRow = {};
            for (const [field, index] of fieldMap) {
                const value = row[index];
                newRow[field] = CONFIG.numericFields.includes(field) 
                    ? (parseFloat(value) || 0)
                    : (value ? String(value).trim() : '');
            }
            return newRow;
        })
        .filter(row => row['药品ID']); // 过滤空行
    
    return {
        fileName,
        headers: Array.from(fieldMap.keys()),
        rows,
        rowCount: rows.length,
        headerRowNum
    };
}