// 公共模块 - 配置和工具函数

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
    ],
    dataFields: [
        '含税出库金额', 'P4毛利额', 'P4毛利率', 
        '应收边际利润额(不含税)', '边际利润率', '客户数'
    ]
};

// 状态管理
const state = {
    selectedFiles: [],
    summaryWorkbook: null,
    processedData: {},
    matchFileData: null,
    selectedFieldIndices: []
};

// DOM元素缓存
const $ = id => document.getElementById(id);
const elements = {
    folderInput: null,
    folderPath: null,
    selectFolderBtn: null,
    headerRowSelect: null,
    processBtn: null,
    matchBtn: null,
    errorSection: null,
    errorMessage: null,
    downloadBtn: null,
    matchModal: null,
    matchFileInput: null,
    matchFilePath: null,
    selectMatchFileBtn: null,
    fieldTagsArea: null,
    cancelMatchBtn: null,
    confirmMatchBtn: null
};

// 初始化DOM元素（需要在页面加载后调用）
function initElements() {
    elements.folderInput = $('folderInput');
    elements.folderPath = $('folderPath');
    elements.selectFolderBtn = $('selectFolderBtn');
    elements.headerRowSelect = $('headerRowSelect');
    elements.processBtn = $('processBtn');
    elements.matchBtn = $('matchBtn');
    elements.errorSection = $('errorSection');
    elements.errorMessage = $('errorMessage');
    elements.downloadBtn = $('downloadBtn');
    elements.matchModal = $('matchModal');
    elements.matchFileInput = $('matchFileInput');
    elements.matchFilePath = $('matchFilePath');
    elements.selectMatchFileBtn = $('selectMatchFileBtn');
    elements.fieldTagsArea = $('fieldTagsArea');
    elements.cancelMatchBtn = $('cancelMatchBtn');
    elements.confirmMatchBtn = $('confirmMatchBtn');
}

// 工具函数
const utils = {
    parseNumber: str => {
        if (typeof str === 'number') return str;
        return parseFloat(String(str).replace(/,/g, '')) || 0;
    },
    
    formatDate: (date = new Date()) => {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    },
    
    getFileType: fileName => {
        const types = ['同期', '上期', '当期'];
        return types.find(type => fileName.includes(type)) || null;
    },
    
    setLoadingState: (button, loading, text = '') => {
        button.disabled = loading;
        button.innerHTML = loading ? '<span class="loading"></span>' : (text || button.textContent);
    },
    
    showError: message => {
        elements.errorSection.style.display = 'block';
        elements.errorMessage.textContent = message;
    },
    
    hideError: () => {
        elements.errorSection.style.display = 'none';
    }
};