// 字段选择器实例
const fieldSelectors = {
    group: null,
    basicInfo: null
};

// 业务逻辑函数
function checkButtonStates() {
    const hasFiles = state.selectedFiles.length > 0;
    const hasHeaderRow = document.getElementById('headerRowSelect').value !== '';
    const hasGroupFields = fieldSelectors.group && fieldSelectors.group.getSelected().length > 0;
    const hasBasicInfoFields = fieldSelectors.basicInfo && fieldSelectors.basicInfo.getSelected().length > 0;
    
    // 修改：添加字段按钮和开始处理按钮使用相同的启用条件
    const canProcess = hasFiles && hasHeaderRow && hasGroupFields && hasBasicInfoFields;
    
    document.getElementById('matchBtn').disabled = !canProcess;
    document.getElementById('processBtn').disabled = !canProcess;
}

async function handleMainFileSelect(event) {
    const validator = file => {
        const isExcel = /\.(xlsx?|xls)$/i.test(file.name);
        const hasKeyword = utils.getFileType(file.name) !== null;
        return isExcel && hasKeyword;
    };

    await FormHandler.handleFileSelect(
        event.target,
        document.getElementById('folderPath'),
        validator,
        (validFiles) => {
            state.selectedFiles = validFiles;
            utils.hideError();
            Object.values(fieldSelectors).forEach(selector => {
                if (selector) selector.reset();
            });
            checkButtonStates();
        }
    );

    if (state.selectedFiles.length === 0 && event.target.files.length > 0) {
        utils.showError('未找到包含"同期"、"上期"、"当期"的Excel文件');
    }
}

async function handleMainHeaderRowSelect() {
    try {
        const result = await FormHandler.handleHeaderRowSelect(
            state.selectedFiles,
            document.getElementById('headerRowSelect'),
            [fieldSelectors.group, fieldSelectors.basicInfo]
        );
        
        if (result) {
            state.allFields = result.fields;
        }
        
        checkButtonStates();
    } catch (error) {
        utils.showError('读取字段失败');
    }
}

async function processFiles() {
    const headerRowNum = parseInt(document.getElementById('headerRowSelect').value);
    const groupFields = fieldSelectors.group.getSelected();
    const basicInfoFields = fieldSelectors.basicInfo.getSelected();
    
    if (state.selectedFiles.length === 0) {
        utils.showError('请先选择包含"同期"、"上期"、"当期"的Excel文件');
        return;
    }
    
    if (!headerRowNum || headerRowNum < 1 || headerRowNum > 5) {
        utils.showError('请选择有效的标题行（1-5）');
        return;
    }

    if (groupFields.length === 0) {
        utils.showError('请选择至少一个分组字段');
        return;
    }

    if (basicInfoFields.length === 0) {
        utils.showError('请选择至少一个基础信息字段');
        return;
    }
    
    utils.hideError();
    utils.setLoadingState(document.getElementById('processBtn'), true, '开始处理');
    document.getElementById('matchBtn').disabled = true;
    
    try {
        const processedData = {};
        
        for (const file of state.selectedFiles) {
            const workbook = await ExcelReader.readFile(file);
            const worksheet = workbook.getWorksheet(1);
            const aggregated = GroupModule.processWorksheet(
                worksheet, headerRowNum, state.allFields, groupFields
            );
            processedData[file.name] = aggregated;
        }
        
        state.processedData = processedData;
        state.summaryWorkbook = FillModule.createSummaryWorkbook(
            processedData, groupFields, basicInfoFields, state.allFields, state.matchConfig
        );
        
        const groupCount = AggregateModule.getAllGroupKeys(processedData).length;
        const groupFieldsText = groupFields.join(' + ');
        const basicInfoFieldsText = basicInfoFields.join(' + ');
        let message = `处理完成！按 ${groupFieldsText} 分组，基础信息：${basicInfoFieldsText}，共 ${groupCount} 个分组`;
        
        if (state.matchConfig) {
            message += `，已匹配 ${state.matchConfig.fieldsToAdd.length} 个字段`;
        }
        
        notification.show(message);
        
        document.getElementById('downloadBtn').disabled = false;
        
    } catch (error) {
        console.error('处理错误：', error);
        utils.showError(`处理失败：${error.message}`);
        notification.show('处理失败，请查看错误信息', 'error');
    } finally {
        utils.setLoadingState(document.getElementById('processBtn'), false, '开始处理');
        checkButtonStates();
    }
}

async function downloadResult() {
    if (!state.summaryWorkbook) {
        notification.show('没有可下载的数据', 'error');
        return;
    }
    
    utils.setLoadingState(document.getElementById('downloadBtn'), true, '下载结果');
    
    try {
        const buffer = await state.summaryWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `品种同环比分析_${utils.formatDate()}.xlsx`;
        link.click();
        
        setTimeout(() => URL.revokeObjectURL(link.href), 100);
        notification.show('文件下载成功');
    } catch (error) {
        notification.show('下载失败', 'error');
    } finally {
        utils.setLoadingState(document.getElementById('downloadBtn'), false, '下载结果');
    }
}

// 初始化
function init() {
    // 初始化字段选择器
    fieldSelectors.group = new FieldSelector('groupFieldsArea', checkButtonStates);
    fieldSelectors.basicInfo = new FieldSelector('basicInfoFieldsArea', checkButtonStates);
    window.fieldSelectors = fieldSelectors;
    
    // 初始化匹配模块
    MatchModule.init();
    
    // 绑定事件
    document.getElementById('selectFolderBtn').addEventListener('click', () => {
        document.getElementById('folderInput').click();
    });
    document.getElementById('folderInput').addEventListener('change', handleMainFileSelect);
    document.getElementById('headerRowSelect').addEventListener('change', handleMainHeaderRowSelect);
    
    document.getElementById('processBtn').addEventListener('click', processFiles);
    document.getElementById('matchBtn').addEventListener('click', () => MatchModule.showModal());
    document.getElementById('downloadBtn').addEventListener('click', downloadResult);
    
    checkButtonStates();
    MatchModule.updateButtonText();
}

document.addEventListener('DOMContentLoaded', init);