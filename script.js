// 处理文件选择 - 保持不变
async function handleFileSelect(event) {
    const files = Array.from(event.target.files);
    const excelFiles = files.filter(file => /\.(xlsx?|xls)$/i.test(file.name));
    
    if (excelFiles.length > CONFIG.maxFiles) {
        utils.showError(`最多只能选择 ${CONFIG.maxFiles} 个文件`);
        event.target.value = '';
        return;
    }
    
    if (!excelFiles.length && files.length) {
        utils.showError('请选择有效的Excel文件（.xlsx, .xls）');
    }
    
    // 重置状态
    state.reset();
    state.selectedFiles = excelFiles;
    document.getElementById('folderPath').value = excelFiles.map(f => f.name).join(', ');
    
    UIManager.renderAllFields();
    UIManager.renderContainer('fieldsListArea', [], { placeholder: CONFIG.placeholder.needSetup });
    checkButtonStates();
}

// 处理标题行选择 - 保持不变
async function handleHeaderRowSelect() {
    const headerRowValue = document.getElementById('headerRowSelect').value;
    if (!state.selectedFiles.length || !headerRowValue) {
        UIManager.renderContainer('fieldsListArea', [], { placeholder: CONFIG.placeholder.needSetup });
        return;
    }

    UIManager.renderContainer('fieldsListArea', CONFIG.placeholder.loading);

    try {
        const workbook = await ExcelReader.readFile(state.selectedFiles[0]);
        const worksheet = workbook.getWorksheet(1);
        const fields = ExcelReader.extractFields(worksheet, parseInt(headerRowValue));
        
        state.allFields = fields;
        UIManager.renderContainer('fieldsListArea', fields);
    } catch (error) {
        utils.showError('读取字段失败');
        console.error('读取字段失败：', error);
        UIManager.renderContainer('fieldsListArea', [], { placeholder: CONFIG.placeholder.needSetup });
    }
    
    checkButtonStates();
}

// 处理文件 - 增加计算字段处理
async function processFiles() {
    const headerRowNum = parseInt(document.getElementById('headerRowSelect').value);
    const processBtn = document.getElementById('processBtn');
    
    processBtn.disabled = true;
    processBtn.textContent = '正在处理中......';
    
    try {
        const processedData = {};
        
        // 处理每个文件
        for (const file of state.selectedFiles) {
            const workbook = await ExcelReader.readFile(file);
            const worksheet = workbook.getWorksheet(1);
            processedData[file.name] = GroupModule.processWorksheet(worksheet, headerRowNum);
        }
        
        // 应用计算字段
        if (state.calcFields && state.calcFields.length > 0) {
            CalcEngine.calculateForAllGroups(processedData, state.calcFields);
        }
        
        state.processedData = processedData;
        state.summaryWorkbook = FillModule.createSummaryWorkbook(processedData);
        
        const groupCount = AggregateModule.getAllGroupKeys(processedData).length;
        const calcFieldsCount = state.calcFields ? state.calcFields.length : 0;
        notification.show(
            `处理完成！${state.selectedFiles.length} 个文件，${groupCount} 个分组，${state.aggregateFields.length} 个聚合字段${calcFieldsCount > 0 ? `，${calcFieldsCount} 个计算字段` : ''}`
        );
        
        document.getElementById('downloadBtn').disabled = false;
        
    } catch (error) {
        console.error('处理错误：', error);
        utils.showError(`处理失败：${error.message}`);
        notification.show('处理失败，请查看错误信息', 'error');
    } finally {
        processBtn.textContent = '开始处理';
        checkButtonStates();
    }
}

// 下载结果 - 保持不变
async function downloadResult() {
    if (!state.summaryWorkbook) {
        notification.show('没有可下载的数据', 'error');
        return;
    }
    
    const downloadBtn = document.getElementById('downloadBtn');
    const downloadIcon = downloadBtn.querySelector('.download-icon');
    const loadingSpinner = downloadBtn.querySelector('.loading-spinner');
    
    downloadBtn.disabled = true;
    downloadIcon.style.display = 'none';
    loadingSpinner.style.display = 'block';
    
    try {
        const buffer = await state.summaryWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `品种数据汇总_${utils.formatDate()}.xlsx`;
        link.click();
        
        setTimeout(() => URL.revokeObjectURL(link.href), 100);
        notification.show('文件下载成功');
    } catch (error) {
        notification.show('下载失败', 'error');
        console.error('下载失败：', error);
    } finally {
        downloadIcon.style.display = 'block';
        loadingSpinner.style.display = 'none';
        downloadBtn.disabled = false;
    }
}

// 初始化 - 保持不变
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('selectFolderBtn').addEventListener('click', () => {
        document.getElementById('folderInput').click();
    });
    document.getElementById('folderInput').addEventListener('change', handleFileSelect);
    document.getElementById('headerRowSelect').addEventListener('change', handleHeaderRowSelect);
    document.getElementById('processBtn').addEventListener('click', processFiles);
    document.getElementById('downloadBtn').addEventListener('click', downloadResult);
});