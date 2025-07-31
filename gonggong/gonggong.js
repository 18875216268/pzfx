// 全局配置
const CONFIG = {
    aggregateMethods: {
        sum: '求和',
        avg: '均值',
        count: '计数',
        max: '最大值',
        min: '最小值'
    },
    groupTypes: {
        group: '分组',
        info: '信息'
    },
    placeholder: {
        default: '暂无数据',
        needSetup: '请先选择文件和标题行......',
        loading: '字段加载中......'
    },
    maxFiles: 10
};

// 状态管理器 - 重构：直接使用显示格式作为内部存储
class StateManager {
    constructor() {
        this.reset();
    }
    
    reset() {
        Object.assign(this, {
            selectedFiles: [],
            summaryWorkbook: null,
            processedData: {},
            allFields: [],
            // 直接存储显示名称
            groupFields: [], // ['品种.分组', '负责人.信息']
            aggregateFields: [], // ['金额.求和', '金额.求和.2', '成本.均值']
            calcFields: [], // [{name: '利润', formula: '金额.求和-成本.均值', id}]
            basicInfoFields: [] // 保留但不使用
        });
    }
    
    // 添加分组字段
    addGroupField(field, type = 'group') {
        const displayName = `${field}.${CONFIG.groupTypes[type]}`;
        if (!this.groupFields.includes(displayName)) {
            this.groupFields.push(displayName);
            return true;
        }
        return false;
    }
    
    // 添加聚合字段
    addAggregateField(field, method = 'sum') {
        const methodDisplay = CONFIG.aggregateMethods[method];
        const baseDisplayName = `${field}.${methodDisplay}`;
        
        // 计算序号
        const sameFieldMethod = this.aggregateFields.filter(name => 
            name.startsWith(`${field}.${methodDisplay}`)
        );
        
        const displayName = sameFieldMethod.length > 0 ? 
            `${baseDisplayName}.${sameFieldMethod.length + 1}` : 
            baseDisplayName;
            
        this.aggregateFields.push(displayName);
        return true;
    }
    
    // 移除字段
    removeField(type, displayName) {
        if (type === 'group') {
            this.groupFields = this.groupFields.filter(name => name !== displayName);
        } else if (type === 'aggregate') {
            this.aggregateFields = this.aggregateFields.filter(name => name !== displayName);
            // 重新计算序号
            this.renumberAggregateFields();
        } else if (type === 'calculated') {
            this.calcFields = this.calcFields.filter(f => f.name !== displayName);
        }
    }
    
    // 重新计算聚合字段序号
    renumberAggregateFields() {
        const fieldMethodGroups = {};
        
        // 按字段和方法分组
        this.aggregateFields.forEach(displayName => {
            const parts = displayName.split('.');
            const field = parts[0];
            const method = parts[1];
            const key = `${field}.${method}`;
            
            if (!fieldMethodGroups[key]) {
                fieldMethodGroups[key] = [];
            }
            fieldMethodGroups[key].push(displayName);
        });
        
        // 重新生成显示名称
        this.aggregateFields = [];
        Object.values(fieldMethodGroups).forEach(group => {
            if (group.length === 1) {
                // 只有一个字段，不需要序号
                const parts = group[0].split('.');
                this.aggregateFields.push(`${parts[0]}.${parts[1]}`);
            } else {
                // 多个字段，添加序号
                group.forEach((_, index) => {
                    const parts = group[0].split('.');
                    this.aggregateFields.push(`${parts[0]}.${parts[1]}.${index + 1}`);
                });
            }
        });
    }
    
    // 更新聚合方法
    updateAggregateMethod(oldDisplayName, newMethod) {
        const index = this.aggregateFields.indexOf(oldDisplayName);
        if (index === -1) return;
        
        const parts = oldDisplayName.split('.');
        const field = parts[0];
        const newMethodDisplay = CONFIG.aggregateMethods[newMethod];
        
        // 移除旧的
        this.aggregateFields.splice(index, 1);
        
        // 添加新的
        this.addAggregateField(field, newMethod);
        
        // 重新计算序号
        this.renumberAggregateFields();
    }
    
    // 更新分组类型
    updateGroupType(oldDisplayName, newType) {
        const index = this.groupFields.indexOf(oldDisplayName);
        if (index === -1) return;
        
        const field = oldDisplayName.split('.')[0];
        const newDisplayName = `${field}.${CONFIG.groupTypes[newType]}`;
        
        this.groupFields[index] = newDisplayName;
    }


    // 添加新方法：移除并返回字段
    removeAndGetField(type, displayName) {
        let removedField = null;
        
        if (type === 'calculated') {
            const index = this.calcFields.findIndex(f => f.name === displayName);
            if (index !== -1) {
                removedField = this.calcFields[index];
                this.calcFields.splice(index, 1);
            }
        }
        
        return removedField;
    }
    
    // 解析显示名称 - 用于向后兼容
    parseDisplayName(displayName, type) {
        const parts = displayName.split('.');
        
        if (type === 'group') {
            return {
                field: parts[0],
                type: Object.keys(CONFIG.groupTypes).find(key => 
                    CONFIG.groupTypes[key] === parts[1]
                ) || 'group'
            };
        } else if (type === 'aggregate') {
            return {
                field: parts[0],
                method: Object.keys(CONFIG.aggregateMethods).find(key => 
                    CONFIG.aggregateMethods[key] === parts[1]
                ) || 'sum'
            };
        }
        
        return { field: parts[0] };
    }
}

// 全局状态实例
const state = new StateManager();

// 工具函数
const utils = {
    parseNumber: str => parseFloat(String(str || 0).replace(/,/g, '')) || 0,
    formatDate: (date = new Date()) => date.toISOString().split('T')[0],
    showError: message => {
        const errorSection = document.getElementById('errorSection');
        const errorMessage = document.getElementById('errorMessage');
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        setTimeout(() => errorSection.style.display = 'none', 5000);
    },
    hideError: () => document.getElementById('errorSection').style.display = 'none'
};

// Excel读取器
class ExcelReader {
    static async readFile(file) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await file.arrayBuffer());
        return workbook;
    }

    static extractFields(worksheet, headerRowNum) {
        return worksheet.getRow(headerRowNum).values
            .slice(1)
            .map(v => String(v || '').trim())
            .filter(Boolean);
    }
}

// 统一的UI管理器
class UIManager {
    static renderContainer(containerId, items, options = {}) {
        const container = document.getElementById(containerId);
        if (!container) return;
        
        if (typeof items === 'string') {
            container.innerHTML = `<div class="placeholder">${items}</div>`;
            return;
        }
        
        if (!items || !items.length) {
            container.innerHTML = `<div class="placeholder">${options.placeholder || CONFIG.placeholder.default}</div>`;
            return;
        }
        
        container.innerHTML = items.map(item => this.createFieldItem(item, options)).join('');
    }
    
    static createFieldItem(item, options = {}) {
        const { type } = options;
        let displayText = item;
        
        // 直接使用显示名称
        return `
            <div class="field-item draggable${type ? ' removable' : ''}" draggable="true" data-field="${item}" data-type="${type || ''}" data-display="${displayText}">
                <span class="content" title="${displayText}">${displayText}</span>
                ${type ? '<span class="remove-btn">×</span>' : ''}
            </div>
        `;
    }
    
    static renderFields(type) {
        // 对于计算字段，调用专门的渲染方法
        if (type === 'calculated') {
            // 直接调用 CalcFieldModule 的渲染方法
            if (typeof CalcFieldModule !== 'undefined' && CalcFieldModule.renderCalcFields) {
                CalcFieldModule.renderCalcFields();
            }
            return;
        }
        
        // 其他类型字段的原有逻辑
        const containerId = { 
            group: 'groupFieldsInput', 
            aggregate: 'aggregateFieldsInput'
        }[type];
        
        let fields = [];
        if (type === 'group') fields = state.groupFields;
        else if (type === 'aggregate') fields = state.aggregateFields;
        
        this.renderContainer(containerId, fields, { type });
    }
    
    static renderAllFields() {
        ['group', 'aggregate', 'calculated'].forEach(type => this.renderFields(type));
    }
}

// 拖拽管理器
class DragManager {
    constructor() {
        this.dragData = null;
        this.indicator = document.createElement('div');
        this.indicator.className = 'drop-indicator';
        this.init();
    }
    
    init() {
        document.addEventListener('dragstart', e => {
            const item = e.target.closest('.draggable');
            if (item) {
                this.dragData = {
                    field: item.dataset.field,
                    displayName: item.dataset.display || item.dataset.field,
                    fromType: item.dataset.type,
                    isFromList: !item.dataset.type
                };
            }
        });
        
        document.addEventListener('dragend', () => {
            this.dragData = null;
            this.indicator.remove();
            document.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
        });
        
        document.querySelectorAll('.field-container[data-type]').forEach(container => {
            container.addEventListener('dragover', e => {
                e.preventDefault();
                
                // 检查是否允许拖入
                const targetType = container.dataset.type;
                const sourceType = this.dragData ? this.dragData.fromType : null;
                
                // 计算字段区域只接受内部拖动
                if (targetType === 'calculated') {
                    if (!this.dragData || this.dragData.fromType !== 'calculated') {
                        e.dataTransfer.effectAllowed = 'none';
                        e.dataTransfer.dropEffect = 'none';
                        return;
                    }
                }
                
                // 其他区域不接受计算字段
                if (sourceType === 'calculated' && targetType !== 'calculated') {
                    e.dataTransfer.effectAllowed = 'none';
                    e.dataTransfer.dropEffect = 'none';
                    return;
                }
                
                container.classList.add('drag-over');
                
                const items = container.querySelectorAll('.field-item');
                if (items.length === 0) {
                    this.indicator.remove();
                    return;
                }
                
                const afterElement = [...container.children]
                    .filter(el => el !== this.indicator && el.classList.contains('field-item'))
                    .find(child => {
                        const box = child.getBoundingClientRect();
                        return e.clientY < box.top + box.height / 2;
                    });
                
                if (afterElement) {
                    container.insertBefore(this.indicator, afterElement);
                } else {
                    container.appendChild(this.indicator);
                }
            });

            container.addEventListener('drop', e => {
                e.preventDefault();
                
                // 检查是否允许放置
                const targetType = container.dataset.type;
                const sourceType = this.dragData ? this.dragData.fromType : null;
                
                // 计算字段区域只接受内部拖动
                if (targetType === 'calculated') {
                    if (!this.dragData || this.dragData.fromType !== 'calculated') {
                        return;
                    }
                }
                
                // 其他区域不接受计算字段
                if (sourceType === 'calculated' && targetType !== 'calculated') {
                    return;
                }
                
                container.classList.remove('drag-over');
                const afterElement = this.indicator.nextElementSibling;
                this.indicator.remove();
                this.handleDrop(container.dataset.type, afterElement);
            });
        });
    }
    
    handleDrop(targetType, beforeElement) {
        if (!this.dragData) return;
        
        const { field, displayName, fromType, isFromList } = this.dragData;
        
        // 特殊处理计算字段
        if (targetType === 'calculated' && fromType === 'calculated') {
            // 先获取并移除字段
            const draggedField = state.removeAndGetField('calculated', displayName);
            if (!draggedField) return;
            
            // 计算插入位置
            const insertIndex = beforeElement ? 
                state.calcFields.findIndex(f => f.name === beforeElement.dataset.display) : 
                state.calcFields.length;
            
            // 插入到新位置
            if (insertIndex >= 0 && insertIndex < state.calcFields.length) {
                state.calcFields.splice(insertIndex, 0, draggedField);
            } else {
                state.calcFields.push(draggedField);
            }
            
            UIManager.renderFields('calculated');
            checkButtonStates();
            return;
        }
        
        // 删除原位置
        if (!isFromList && fromType) {
            state.removeField(fromType, displayName);
        }
        
        // 添加到新位置
        if (targetType === 'aggregate') {
            const parsed = state.parseDisplayName(field, 'group');
            const insertIndex = beforeElement ? 
                state.aggregateFields.indexOf(beforeElement.dataset.display) : 
                state.aggregateFields.length;
            
            // 临时移除以便插入到正确位置
            const newDisplayName = state.addAggregateField(parsed.field);
            const addedName = state.aggregateFields.pop(); // 移除刚添加的
            
            if (insertIndex >= 0 && insertIndex < state.aggregateFields.length) {
                state.aggregateFields.splice(insertIndex, 0, addedName);
            } else {
                state.aggregateFields.push(addedName);
            }
            
        } else if (targetType === 'group') {
            if (!state.groupFields.includes(displayName)) {
                const parsed = state.parseDisplayName(field, 'group');
                const insertIndex = beforeElement ? 
                    state.groupFields.indexOf(beforeElement.dataset.display) : 
                    state.groupFields.length;
                
                const newDisplayName = `${parsed.field}.${CONFIG.groupTypes.group}`;
                
                if (insertIndex >= 0 && insertIndex < state.groupFields.length) {
                    state.groupFields.splice(insertIndex, 0, newDisplayName);
                } else {
                    state.groupFields.push(newDisplayName);
                }
            }
        }
        
        UIManager.renderFields(targetType);
        if (fromType && fromType !== targetType) {
            UIManager.renderFields(fromType);
        }
        checkButtonStates();
    }
}

// 全局事件处理
document.addEventListener('click', e => {
    if (e.target.classList.contains('remove-btn')) {
        const item = e.target.closest('.field-item');
        const type = item.dataset.type;
        const displayName = item.dataset.display || item.dataset.field;
        
        state.removeField(type, displayName);
        UIManager.renderFields(type);
        checkButtonStates();
    }
    
    // 聚合方法菜单
    const aggregateItem = e.target.closest('.field-container[data-type="aggregate"] .field-item');
    if (aggregateItem) {
        showPopupMenu(e, aggregateItem.dataset.display, 'aggregate');
    }
    
    // 分组类型菜单
    const groupItem = e.target.closest('.field-container[data-type="group"] .field-item');
    if (groupItem) {
        showPopupMenu(e, groupItem.dataset.display, 'group');
    }
});

// 统一的弹出菜单函数
function showPopupMenu(event, displayName, type) {
    event.stopPropagation();
    
    document.querySelectorAll('.popup-menu').forEach(m => m.remove());
    
    const menu = document.createElement('div');
    menu.className = 'popup-menu';
    
    let currentValue, options, updateFunction;
    
    if (type === 'aggregate') {
        const parsed = state.parseDisplayName(displayName, 'aggregate');
        currentValue = parsed.method;
        options = CONFIG.aggregateMethods;
        updateFunction = method => {
            state.updateAggregateMethod(displayName, method);
            UIManager.renderFields('aggregate');
        };
    } else {
        const parsed = state.parseDisplayName(displayName, 'group');
        currentValue = parsed.type;
        options = CONFIG.groupTypes;
        updateFunction = groupType => {
            state.updateGroupType(displayName, groupType);
            UIManager.renderFields('group');
        };
    }
    
    Object.entries(options).forEach(([key, label]) => {
        const item = document.createElement('div');
        item.className = `popup-menu-item${currentValue === key ? ' active' : ''}`;
        item.textContent = label;
        item.onclick = () => {
            updateFunction(key);
            menu.remove();
        };
        menu.appendChild(item);
    });
    
    const rect = event.target.getBoundingClientRect();
    Object.assign(menu.style, {
        left: rect.left + 'px',
        top: rect.bottom + 5 + 'px'
    });
    
    document.body.appendChild(menu);
    setTimeout(() => {
        document.addEventListener('click', () => menu.remove(), { once: true });
    }, 0);
}

// 检查按钮状态
function checkButtonStates() {
    const hasGroupFields = state.groupFields.filter(name => 
        name.endsWith('.分组')
    ).length > 0;
    
    const canProcess = state.selectedFiles.length > 0 &&
                      document.getElementById('headerRowSelect').value &&
                      hasGroupFields;
    
    document.getElementById('processBtn').disabled = !canProcess;
}

// 初始化
let dragManager = null;

document.addEventListener('DOMContentLoaded', () => {
    dragManager = new DragManager();
    UIManager.renderAllFields();
    checkButtonStates();
});
