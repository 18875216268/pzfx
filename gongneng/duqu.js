// 读取数据模块 - 优化版本

/**
 * 读取选中的Excel文件
 * @param {FileList} files - 选中的文件列表
 * @param {number} headerRowNum - 标题行行号（从1开始）
 * @returns {Promise<Map>} 返回文件名到工作表数据的映射
 */
export async function readExcelFiles(files, headerRowNum = 1) {
    const fileDataMap = new Map();
    
    for (const file of files) {
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array', cellStyles: true });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const merges = worksheet['!merges'] || [];
            
            // 转换为JSON数组
            const data = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '', 
                raw: false 
            });
            
            // 提取标题行（处理合并单元格）
            const headers = extractHeaders(worksheet, headerRowNum - 1, merges);
            
            fileDataMap.set(file.name, {
                fileName: file.name,
                worksheet,
                data,
                headers,
                headerRowNum,
                merges
            });
            
            console.log(`成功读取文件：${file.name}，标题行在第${headerRowNum}行`);
            
        } catch (error) {
            throw new Error(`读取文件 ${file.name} 失败：${error.message}`);
        }
    }
    
    return fileDataMap;
}

/**
 * 提取标题行，智能处理合并单元格
 * @param {Object} worksheet - 工作表对象
 * @param {number} rowIndex - 行索引（从0开始）
 * @param {Array} merges - 合并单元格信息
 * @returns {Array} 标题数组
 */
function extractHeaders(worksheet, rowIndex, merges) {
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const headers = [];
    
    // 创建合并单元格查找表
    const mergeMap = {};
    merges.forEach(({ s, e }) => {
        for (let r = s.r; r <= e.r; r++) {
            for (let c = s.c; c <= e.c; c++) {
                mergeMap[`${r}_${c}`] = { r: s.r, c: s.c };
            }
        }
    });
    
    // 提取每列的标题
    for (let col = 0; col <= range.e.c; col++) {
        const merge = mergeMap[`${rowIndex}_${col}`];
        const cellRef = XLSX.utils.encode_cell({ 
            r: merge ? merge.r : rowIndex, 
            c: merge ? merge.c : col 
        });
        
        const cell = worksheet[cellRef];
        headers.push(cell ? String(cell.v || '').trim() : '');
    }
    
    return headers;
}

/**
 * 文件验证
 * @param {string} fileName - 文件名
 * @returns {string|null} 文件类型或null
 */
export function getFileType(fileName) {
    const types = ['同期', '上期', '当期'];
    return types.find(type => fileName.includes(type)) || null;
}