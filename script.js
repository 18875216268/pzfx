// 导入模块
import { readExcelFiles, getFileType } from './gongneng/duqu.js';
import { extractRequiredFields } from './gongneng/tiqu.js';
import { processAndAggregateData } from './gongneng/chuli.js';
import { createSummaryWorkbook } from './gongneng/huizong.js';

// DOM元素
const elements = {
    folderInput: document.getElementById('folderInput'),
    folderPath: document.getElementById('folderPath'),
    selectFolderBtn: document.getElementById('selectFolderBtn'),
    headerRowSelect: document.getElementById('headerRowSelect'),
    processBtn: document.getElementById('processBtn'),
    errorSection: document.getElementById('errorSection'),
    errorMessage: document.getElementById('errorMessage'),
    downloadBtn: document.getElementById('downloadBtn')
};

// 状态管理
let state = {
    selectedFiles: [],
    summaryWorkbook: null,
    processedData: {}
};

// 事件监听 - 点击按钮或输入框都能选择文件
elements.selectFolderBtn.addEventListener('click', () => elements.folderInput.click());
elements.folderPath.addEventListener('click', () => elements.folderInput.click());
elements.folderInput.addEventListener('change', handleFileSelect);
elements.processBtn.addEventListener('click', processFiles);
elements.downloadBtn.addEventListener('click', downloadResult);

// 文件选择处理
function handleFileSelect(event) {
    const files = Array.from(event.target.files);
    
    // 筛选符合条件的文件
    state.selectedFiles = files.filter(file => {
        const isExcel = /\.(xlsx?|xls)$/i.test(file.name);
        const hasKeyword = getFileType(file.name) !== null;
        return isExcel && hasKeyword;
    });
    
    if (state.selectedFiles.length > 0) {
        // 显示文件名，用逗号分隔
        const fileNames = state.selectedFiles.map(f => f.name).join(', ');
        elements.folderPath.value = fileNames;
        elements.processBtn.disabled = false;
        
        console.log(`已选择 ${state.selectedFiles.length} 个符合条件的文件`);
    } else {
        elements.folderPath.value = '';
        elements.processBtn.disabled = true;
        if (files.length > 0) {
            showError('未找到包含"同期"、"上期"、"当期"的Excel文件');
        }
    }
}

// 处理文件
async function processFiles() {
    const headerRowNum = parseInt(elements.headerRowSelect.value);
    
    if (state.selectedFiles.length === 0) {
        return showError('请先选择包含"同期"、"上期"、"当期"的Excel文件');
    }
    
    // 重置状态
    state.processedData = {};
    elements.errorSection.style.display = 'none';
    elements.processBtn.disabled = true;
    elements.downloadBtn.disabled = true;
    
    const startTime = Date.now();
    
    try {
        // 处理步骤
        const steps = [
            { percent: 10, text: '正在读取Excel文件...', 
              action: () => readExcelFiles(state.selectedFiles, headerRowNum) },
            { percent: 30, text: '正在提取数据字段...', 
              action: extractRequiredFields },
            { percent: 50, text: '正在处理和汇总数据...', 
              action: processAndAggregateData },
            { percent: 80, text: '正在生成汇总表...', 
              action: createSummaryWorkbook }
        ];
        
        let result;
        for (const step of steps) {
            updateProgress(step.percent, step.text);
            result = await step.action(result);
            if (step.percent === 50) state.processedData = result;
        }
        
        state.summaryWorkbook = result;
        
        // 完成处理
        const processTime = ((Date.now() - startTime) / 1000).toFixed(2);
        const drugCount = new Set(
            Object.values(state.processedData)
                .flatMap(data => Object.keys(data))
        ).size;
        
        // 恢复按钮文字
        elements.processBtn.textContent = '开始处理';
        
        console.log(`处理完成！共处理 ${state.selectedFiles.length} 个文件，${drugCount} 种药品，用时 ${processTime} 秒`);
        
        // 启用下载按钮
        elements.downloadBtn.disabled = false;
        elements.processBtn.disabled = false;
        
    } catch (error) {
        console.error('处理错误：', error);
        showError(`处理失败：${error.message}`);
        // 恢复按钮文字
        elements.processBtn.textContent = '开始处理';
        elements.processBtn.disabled = false;
    }
}

// 下载结果
async function downloadResult() {
    if (!state.summaryWorkbook) {
        return showError('没有可下载的数据');
    }
    
    try {
        const buffer = await state.summaryWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `品种同环比分析_${new Date().toISOString().slice(0, 10)}.xlsx`;
        link.click();
        URL.revokeObjectURL(link.href);
    } catch (error) {
        showError('下载失败，请重试');
    }
}

// 更新进度 - 修改为更新按钮文字
function updateProgress(percent, text) {
    elements.processBtn.textContent = `正在处理 ${percent}%`;
    console.log(text); // 在控制台记录详细进度信息
}

// 显示错误
function showError(message) {
    elements.errorSection.style.display = 'block';
    elements.errorMessage.textContent = message;
}