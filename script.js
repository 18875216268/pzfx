// 主JavaScript文件 - 初始化和主处理流程

// 主处理流程
async function processFiles() {
    const headerRowNum = parseInt(elements.headerRowSelect.value);
    
    if (state.selectedFiles.length === 0) {
        utils.showError('请先选择包含"同期"、"上期"、"当期"的Excel文件');
        return;
    }
    
    utils.hideError();
    utils.setLoadingState(elements.processBtn, true);
    
    try {
        // 步骤1: 读取和处理文件
        const processedData = {};
        
        for (const file of state.selectedFiles) {
            const workbook = await ExcelProcessor.readFile(file);
            const worksheet = workbook.getWorksheet(1);
            const data = ExcelProcessor.processWorksheet(worksheet, headerRowNum, CONFIG.keepFields);
            
            // 按药品ID汇总
            const aggregated = {};
            data.rows.forEach(row => {
                const drugId = String(row['药品ID']);
                if (!aggregated[drugId]) {
                    aggregated[drugId] = { ...row };
                } else {
                    // 累加数值字段
                    CONFIG.numericFields.forEach(field => {
                        aggregated[drugId][field] += row[field];
                    });
                }
            });
            
            processedData[file.name] = aggregated;
        }
        
        state.processedData = processedData;
        
        // 步骤2: 创建汇总工作簿
        state.summaryWorkbook = ExcelProcessor.createSummaryWorkbook(processedData);
        
        // 完成处理
        const drugCount = new Set(
            Object.values(processedData).flatMap(data => Object.keys(data))
        ).size;
        
        utils.showNotification(`处理完成！共 ${drugCount} 种药品`);
        
        elements.downloadBtn.disabled = false;
        elements.matchBtn.disabled = false;
        
    } catch (error) {
        console.error('处理错误：', error);
        utils.showError(`处理失败：${error.message}`);
        utils.showNotification('处理失败，请查看错误信息', 'error');
    } finally {
        utils.setLoadingState(elements.processBtn, false, '开始处理');
    }
}

// 下载结果
async function downloadResult() {
    if (!state.summaryWorkbook) {
        utils.showNotification('没有可下载的数据', 'error');
        return;
    }
    
    utils.setLoadingState(elements.downloadBtn, true);
    
    try {
        const buffer = await state.summaryWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `品种同环比分析_${utils.formatDate()}.xlsx`;
        link.click();
        URL.revokeObjectURL(link.href);
        
        utils.showNotification('文件下载成功');
    } catch (error) {
        utils.showNotification('下载失败', 'error');
    } finally {
        utils.setLoadingState(elements.downloadBtn, false, '下载结果');
    }
}

// 文件选择处理
function handleFileSelect(event) {
    const files = Array.from(event.target.files);
    state.selectedFiles = files.filter(file => {
        const isExcel = /\.(xlsx?|xls)$/i.test(file.name);
        const hasKeyword = utils.getFileType(file.name) !== null;
        return isExcel && hasKeyword;
    });
    
    if (state.selectedFiles.length > 0) {
        elements.folderPath.value = state.selectedFiles.map(f => f.name).join(', ');
        elements.processBtn.disabled = false;
    } else {
        elements.folderPath.value = '';
        elements.processBtn.disabled = true;
        if (files.length > 0) {
            utils.showError('未找到包含"同期"、"上期"、"当期"的Excel文件');
        }
    }
}

// 初始化函数
function init() {
    // 初始化DOM元素
    initElements();
    
    // 绑定事件
    elements.selectFolderBtn.addEventListener('click', () => elements.folderInput.click());
    elements.folderInput.addEventListener('change', handleFileSelect);
    
    elements.processBtn.addEventListener('click', processFiles);
    elements.matchBtn.addEventListener('click', showMatchModal);
    elements.downloadBtn.addEventListener('click', downloadResult);
    
    elements.selectMatchFileBtn.addEventListener('click', () => elements.matchFileInput.click());
    elements.matchFileInput.addEventListener('change', handleMatchFileSelect);
    elements.cancelMatchBtn.addEventListener('click', closeMatchModal);
    elements.confirmMatchBtn.addEventListener('click', confirmMatch);
}

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', init);