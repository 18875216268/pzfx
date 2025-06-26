// 处理文件选择
async function handleFileSelect(event) {
    const files = Array.from(event.target.files);
    state.selectedFiles = files.filter(file => 
        /\.(xlsx?|xls)$/i.test(file.name) && utils.getFileType(file.name)
    );
    
    if (state.selectedFiles.length > 0) {
        document.getElementById('folderPath').value = state.selectedFiles.map(f => f.name).join(', ');
        utils.hideError();
        
        // 重置所有字段
        Object.assign(state, {
            groupFields: [],
            aggregateFields: {},
            basicInfoFields: [],
            calcFields: {}
        });
        
        // 重新渲染所有区域
        ['group', 'aggregate', 'basicInfo'].forEach(type => {
            dragDropManager?.renderFields(type);
        });
        renderCalcFields();
        fieldsListManager?.reset();
    } else {
        document.getElementById('folderPath').value = '';
        if (files.length > 0) {
            utils.showError('未找到包含"同期"、"上期"、"当期"的Excel文件');
        }
    }
    
    checkButtonStates();
}

// 处理标题行选择
async function handleHeaderRowSelect() {
    const headerRowValue = document.getElementById('headerRowSelect').value;
    if (!state.selectedFiles.length || !headerRowValue) {
        fieldsListManager.reset();
        return;
    }

    fieldsListManager.showLoading();

    try {
        const workbook = await ExcelReader.readFile(state.selectedFiles[0]);
        const worksheet = workbook.getWorksheet(1);
        const fields = ExcelReader.extractFields(worksheet, parseInt(headerRowValue));
        
        state.allFields = fields;
        fieldsListManager.setFields(fields);
        checkButtonStates();
    } catch (error) {
        utils.showError('读取字段失败');
        console.error('读取字段失败：', error);
    }
}

// 计算字段相关函数
function showCalcFieldModal() {
    const modalHtml = `
        <div class="modal-overlay" id="calcFieldModal">
            <div class="modal">
                <div class="modal-header">创建计算字段</div>
                <div class="modal-content">
                    <input type="text" id="calcFieldInput" class="modal-input" 
                           placeholder="新字段=表达式 (如: 毛利率=毛利额/销售额)">
                    <div style="margin-top: 8px; font-size: 12px; color: #666;">
                        提示：请使用已添加的聚合字段进行计算
                    </div>
                </div>
                <div class="modal-actions">
                    <button class="modal-btn modal-btn-cancel" onclick="closeCalcFieldModal()">取消</button>
                    <button class="modal-btn modal-btn-confirm" onclick="confirmCalcField()">确认</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    
    const input = document.getElementById('calcFieldInput');
    input.focus();
    input.addEventListener('keypress', e => e.key === 'Enter' && confirmCalcField());
}

function closeCalcFieldModal() {
    document.getElementById('calcFieldModal')?.remove();
}

function confirmCalcField() {
    const input = document.getElementById('calcFieldInput');
    const content = input.value.trim();
    
    if (!content) {
        closeCalcFieldModal();
        return;
    }
    
    const match = content.match(/^([^=]+)=(.+)$/);
    if (!match) {
        utils.showError('公式格式错误，请使用：新字段=表达式');
        return;
    }
    
    const [, fieldName, expression] = match.map(s => s.trim());
    
    if (!fieldName || !expression) {
        utils.showError('字段名和表达式不能为空');
        return;
    }
    
    state.calcFields[fieldName] = expression;
    renderCalcFields();
    closeCalcFieldModal();
    utils.hideError();
}

// 渲染计算字段
function renderCalcFields() {
    const fields = Object.entries(state.calcFields).map(([field, expr]) => ({
        field,
        displayText: `${field} = ${expr}`
    }));
    
    FieldRenderer.renderToContainer('calcFieldsList', fields, {
        itemOptions: { removable: true },
        placeholder: CONFIG.placeholder.default
    });
}

// 处理文件
async function processFiles() {
    const headerRowNum = parseInt(document.getElementById('headerRowSelect').value);
    const processBtn = document.getElementById('processBtn');
    
    const originalText = processBtn.textContent;
    processBtn.disabled = true;
    processBtn.textContent = '正在处理中......';
    
    setTimeout(async () => {
        try {
            const processedData = {};
            
            for (const file of state.selectedFiles) {
                const workbook = await ExcelReader.readFile(file);
                const worksheet = workbook.getWorksheet(1);
                processedData[file.name] = GroupModule.processWorksheet(
                    worksheet, headerRowNum, state.allFields, 
                    state.groupFields, state.aggregateFields
                );
            }
            
            state.processedData = processedData;
            state.summaryWorkbook = FillModule.createSummaryWorkbook(
                processedData, state.groupFields, state.aggregateFields, 
                state.calcFields, state.basicInfoFields, state.allFields
            );
            
            const stats = {
                groups: AggregateModule.getAllGroupKeys(processedData).length,
                aggregates: Object.keys(state.aggregateFields).length,
                calcs: Object.keys(state.calcFields).length
            };
            
            notification.show(
                `处理完成！${stats.groups} 个分组，${stats.aggregates} 个聚合字段，${stats.calcs} 个计算字段`
            );
            
            document.getElementById('downloadBtn').disabled = false;
            
        } catch (error) {
            console.error('处理错误：', error);
            utils.showError(`处理失败：${error.message}`);
            notification.show('处理失败，请查看错误信息', 'error');
        } finally {
            processBtn.textContent = originalText;
            checkButtonStates();
        }
    }, 0);
}

// 下载结果
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
        link.download = `品种同环比分析_${utils.formatDate()}.xlsx`;
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

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    // 初始化管理器
    dragDropManager = new DragDropManager();
    fieldsListManager = new FieldsListManager('fieldsListArea');
    
    // 绑定事件
    const events = {
        'selectFolderBtn': ['click', () => document.getElementById('folderInput').click()],
        'folderInput': ['change', handleFileSelect],
        'headerRowSelect': ['change', handleHeaderRowSelect],
        'processBtn': ['click', processFiles],
        'downloadBtn': ['click', downloadResult],
        'addCalcFieldBtn': ['click', showCalcFieldModal]
    };
    
    Object.entries(events).forEach(([id, [event, handler]]) => {
        document.getElementById(id)?.addEventListener(event, handler);
    });
    
    checkButtonStates();
});