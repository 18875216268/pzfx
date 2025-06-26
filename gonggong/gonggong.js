// 全局配置
const CONFIG = {
    aggregateMethods: {
        sum: '求和',
        avg: '均值',
        count: '计数'
    },
    placeholder: {
        default: '暂无数据',
        needSetup: '请添加文件并选择标题行......'
    }
};

// 全局状态
const state = {
    selectedFiles: [],
    summaryWorkbook: null,
    processedData: {},
    allFields: [],
    groupFields: [],
    aggregateFields: {},
    calcFields: {},
    basicInfoFields: []
};

// 工具函数
const utils = {
    parseNumber: str => parseFloat(String(str || 0).replace(/,/g, '')) || 0,
    
    formatDate: (date = new Date()) => date.toISOString().split('T')[0],
    
    getFileType: fileName => ['同期', '上期', '当期'].find(type => fileName.includes(type)),
    
    showError: message => {
        const errorSection = document.getElementById('errorSection');
        const errorMessage = document.getElementById('errorMessage');
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        setTimeout(() => errorSection.style.display = 'none', 5000);
    },
    
    hideError: () => document.getElementById('errorSection').style.display = 'none',
    
    parseExpression(expression, data) {
        try {
            let safeExpression = expression;
            const fields = Object.keys(data).sort((a, b) => b.length - a.length);
            
            fields.forEach(field => {
                const value = utils.parseNumber(data[field]);
                const regex = new RegExp(field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
                safeExpression = safeExpression.replace(regex, value);
            });
            
            return new Function('return ' + safeExpression)();
        } catch (error) {
            console.error('计算错误:', expression, error);
            return 0;
        }
    }
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

// 统一的字段渲染器
class FieldRenderer {
    static createFieldItem(field, { draggable = false, removable = false, displayText = field } = {}) {
        const classes = ['field-item'];
        if (draggable) classes.push('draggable');
        if (removable) classes.push('removable');

        return `
            <div class="${classes.join(' ')}" ${draggable ? 'draggable="true"' : ''}>
                <span class="content" title="${displayText}">${displayText}</span>
                ${removable ? '<span class="remove-btn">×</span>' : ''}
            </div>
        `;
    }

    static renderToContainer(containerId, fields, options = {}) {
        const container = document.getElementById(containerId);
        if (!container) return;

        if (fields.length === 0) {
            container.innerHTML = `<div class="placeholder">${options.placeholder || CONFIG.placeholder.default}</div>`;
            return;
        }

        container.innerHTML = fields.map(field => 
            typeof field === 'object' ? 
                this.createFieldItem(field.field, { ...options.itemOptions, displayText: field.displayText }) :
                this.createFieldItem(field, options.itemOptions || {})
        ).join('');
    }

    static showLoading(containerId, message = '加载中...') {
        const container = document.getElementById(containerId);
        if (container) {
            container.innerHTML = `<div class="placeholder">${message}</div>`;
        }
    }
}

// 拖拽管理器
// 在 gonggong.js 中替换 DragDropManager 类
class DragDropManager {
    constructor() {
        this.draggedField = null;
        this.draggedElement = null;
        this.dropIndicator = this.createDropIndicator();
        this.init();
    }
    
    createDropIndicator() {
        const indicator = document.createElement('div');
        indicator.className = 'drop-indicator';
        indicator.style.cssText = 'position: absolute; height: 2px; background: #2196F3; pointer-events: none; display: none; z-index: 1000;';
        document.body.appendChild(indicator);
        return indicator;
    }
    
    init() {
        // 从字段列表拖拽
        document.addEventListener('dragstart', e => {
            if (e.target.classList.contains('draggable')) {
                if (e.target.closest('#fieldsListArea')) {
                    // 从字段列表拖拽
                    this.draggedField = e.target.querySelector('.content').textContent;
                    this.draggedElement = null; // 明确标记这不是元素排序
                    e.dataTransfer.effectAllowed = 'copy';
                    e.dataTransfer.setData('text/plain', this.draggedField);
                } else if (e.target.closest('.field-container[data-type]')) {
                    // 区域内排序
                    this.draggedElement = e.target;
                    this.draggedField = null; // 明确标记这不是字段拖拽
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', 'sorting');
                }
            }
        });
        
        // 为所有拖放区域设置事件
        document.querySelectorAll('.field-container[data-type]').forEach(zone => {
            zone.addEventListener('dragover', e => this.handleDragOver(e));
            zone.addEventListener('dragleave', e => this.handleDragLeave(e));
            zone.addEventListener('drop', e => this.handleDrop(e));
        });
        
        document.addEventListener('dragend', () => {
            this.hideDropIndicator();
            document.querySelectorAll('.drag-over').forEach(el => el.classList.remove('drag-over'));
        });
    }
    
    handleDragOver(e) {
        e.preventDefault();
        const container = e.currentTarget;
        
        // 添加拖拽悬停样式
        container.classList.add('drag-over');
        
        // 无论是字段拖拽还是元素排序，都显示插入位置指示器
        const afterElement = this.getDragAfterElement(container, e.clientY);
        
        if (afterElement) {
            const rect = afterElement.getBoundingClientRect();
            this.showDropIndicator(rect.left, rect.top - 3, rect.width);
        } else {
            const items = container.querySelectorAll('.field-item');
            if (items.length > 0) {
                const lastItem = items[items.length - 1];
                const rect = lastItem.getBoundingClientRect();
                this.showDropIndicator(rect.left, rect.bottom + 3, rect.width);
            } else {
                // 如果容器为空，隐藏指示器
                this.hideDropIndicator();
            }
        }
    }
    
    handleDragLeave(e) {
        // 只有当离开整个容器时才移除样式
        if (e.target === e.currentTarget) {
            e.currentTarget.classList.remove('drag-over');
            this.hideDropIndicator();
        }
    }
    
    handleDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
        this.hideDropIndicator();
        
        const container = e.currentTarget;
        const type = container.dataset.type;
        const afterElement = this.getDragAfterElement(container, e.clientY);
        
        if (this.draggedField && !this.draggedElement) {
            // 处理从字段列表拖拽的情况
            this.handleFieldDrop(type, this.draggedField, container, afterElement);
        } else if (this.draggedElement && !this.draggedField) {
            // 处理区域内排序的情况
            if (this.draggedElement.parentElement === container) {
                this.handleSortDrop(container, afterElement);
            }
        }
        
        // 重置状态
        this.draggedField = null;
        this.draggedElement = null;
    }
    
    handleFieldDrop(type, field, container, afterElement) {
        // 检查字段是否已存在
        const exists = this.checkFieldExists(type, field);
        if (exists) return;
        
        // 添加字段到状态
        this.addFieldToState(type, field);
        
        // 如果有 afterElement，需要调整顺序
        if (afterElement || container.querySelector('.field-item')) {
            this.reorderAfterInsert(type, field, container, afterElement);
        }
        
        // 重新渲染
        this.renderFields(type);
        checkButtonStates();
    }
    
    handleSortDrop(container, afterElement) {
        if (afterElement && afterElement !== this.draggedElement) {
            container.insertBefore(this.draggedElement, afterElement);
        } else if (!afterElement && this.draggedElement.parentElement === container) {
            container.appendChild(this.draggedElement);
        }
        
        this.updateFieldOrder(container.dataset.type);
        checkButtonStates();
    }
    
    checkFieldExists(type, field) {
        switch (type) {
            case 'group':
                return state.groupFields.includes(field);
            case 'aggregate':
                return field in state.aggregateFields;
            case 'basicInfo':
                return state.basicInfoFields.includes(field);
            default:
                return false;
        }
    }
    
    addFieldToState(type, field) {
        switch (type) {
            case 'group':
                state.groupFields.push(field);
                break;
            case 'aggregate':
                state.aggregateFields[field] = 'sum';
                break;
            case 'basicInfo':
                state.basicInfoFields.push(field);
                break;
        }
    }
    
    reorderAfterInsert(type, field, container, afterElement) {
        const items = Array.from(container.querySelectorAll('.field-item'));
        const fields = items.map(item => {
            const content = item.querySelector('.content').textContent;
            if (type === 'aggregate') {
                const match = content.match(/\((.+)\)/);
                return match ? match[1] : content;
            }
            return content;
        });
        
        // 确定插入位置
        let insertIndex = fields.length;
        if (afterElement) {
            const afterField = this.getFieldFromElement(afterElement, type);
            insertIndex = fields.indexOf(afterField);
        }
        
        // 根据类型更新状态中的顺序
        switch (type) {
            case 'group':
                state.groupFields = state.groupFields.filter(f => f !== field);
                state.groupFields.splice(insertIndex, 0, field);
                break;
            case 'basicInfo':
                state.basicInfoFields = state.basicInfoFields.filter(f => f !== field);
                state.basicInfoFields.splice(insertIndex, 0, field);
                break;
            case 'aggregate':
                // 聚合字段需要重建对象以保持顺序
                const newAggregateFields = {};
                const method = state.aggregateFields[field];
                delete state.aggregateFields[field];
                
                const entries = Object.entries(state.aggregateFields);
                entries.splice(insertIndex, 0, [field, method]);
                
                entries.forEach(([f, m]) => {
                    newAggregateFields[f] = m;
                });
                state.aggregateFields = newAggregateFields;
                break;
        }
    }
    
    getFieldFromElement(element, type) {
        const content = element.querySelector('.content').textContent;
        if (type === 'aggregate') {
            const match = content.match(/\((.+)\)/);
            return match ? match[1] : content;
        }
        return content;
    }
    
    getDragAfterElement(container, y) {
        const draggableElements = [...container.querySelectorAll('.field-item')]
            .filter(el => el !== this.draggedElement);
        
        return draggableElements.reduce((closest, child) => {
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            
            if (offset < 0 && offset > closest.offset) {
                return { offset: offset, element: child };
            } else {
                return closest;
            }
        }, { offset: Number.NEGATIVE_INFINITY }).element;
    }
    
    showDropIndicator(x, y, width) {
        this.dropIndicator.style.left = x + 'px';
        this.dropIndicator.style.top = y + 'px';
        this.dropIndicator.style.width = width + 'px';
        this.dropIndicator.style.display = 'block';
    }
    
    hideDropIndicator() {
        this.dropIndicator.style.display = 'none';
    }
    
    updateFieldOrder(type) {
        const container = document.querySelector(`.field-container[data-type="${type}"]`);
        const fieldItems = container.querySelectorAll('.field-item');
        
        switch (type) {
            case 'group':
                state.groupFields = Array.from(fieldItems).map(item => 
                    item.querySelector('.content').textContent
                );
                break;
            case 'basicInfo':
                state.basicInfoFields = Array.from(fieldItems).map(item => 
                    item.querySelector('.content').textContent
                );
                break;
            case 'aggregate':
                const newAggregateFields = {};
                fieldItems.forEach(item => {
                    const match = item.querySelector('.content').textContent.match(/\((.+)\)/);
                    const field = match ? match[1] : item.querySelector('.content').textContent;
                    if (state.aggregateFields[field]) {
                        newAggregateFields[field] = state.aggregateFields[field];
                    }
                });
                state.aggregateFields = newAggregateFields;
                break;
        }
    }
    
    renderFields(type) {
        const renderers = {
            group: () => FieldRenderer.renderToContainer('groupFieldsInput', state.groupFields, {
                itemOptions: { removable: true, draggable: true },
                placeholder: CONFIG.placeholder.default
            }),
            aggregate: () => {
                const fields = Object.keys(state.aggregateFields).map(field => ({
                    field,
                    displayText: `${CONFIG.aggregateMethods[state.aggregateFields[field]]}(${field})`
                }));
                FieldRenderer.renderToContainer('aggregateFieldsInput', fields, {
                    itemOptions: { removable: true, draggable: true },
                    placeholder: CONFIG.placeholder.default
                });
            },
            basicInfo: () => FieldRenderer.renderToContainer('basicInfoFieldsInput', state.basicInfoFields, {
                itemOptions: { removable: true, draggable: true },
                placeholder: CONFIG.placeholder.default
            })
        };
        
        renderers[type]?.();
        
        // 重新绑定事件
        this.rebindEvents();
    }
    
    rebindEvents() {
        // 为新渲染的元素重新绑定拖拽事件
        document.querySelectorAll('.field-container[data-type] .field-item[draggable="true"]').forEach(item => {
            item.addEventListener('dragstart', e => {
                this.draggedElement = e.target;
                this.draggedField = null;
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/plain', 'sorting');
            });
        });
    }
}

// 字段列表管理器
class FieldsListManager {
    constructor(elementId) {
        this.elementId = elementId;
        this.fields = [];
    }

    setFields(fields) {
        this.fields = fields;
        this.render();
    }

    showLoading(message = '字段加载中...') {
        FieldRenderer.showLoading(this.elementId, message);
    }

    render() {
        FieldRenderer.renderToContainer(this.elementId, this.fields, {
            placeholder: CONFIG.placeholder.needSetup,
            itemOptions: { draggable: true }
        });
    }

    reset() {
        this.fields = [];
        this.render();
    }
}

// 事件处理
document.addEventListener('click', e => {
    // 删除按钮
    if (e.target.classList.contains('remove-btn')) {
        e.stopPropagation();
        
        const container = e.target.closest('.field-container');
        const type = container?.dataset.type;
        const fieldText = e.target.parentElement.querySelector('.content').textContent;
        
        if (type && fieldText) {
            const field = type === 'calc' ? fieldText.split(' = ')[0] :
                         type === 'aggregate' ? (fieldText.match(/\((.+)\)/) || [, fieldText])[1] :
                         fieldText;
            
            removeField(type, field);
        }
    }
    
    // 聚合字段菜单
    const fieldItem = e.target.closest('.field-item');
    if (fieldItem?.closest('.field-container[data-type="aggregate"]')) {
        const match = fieldItem.querySelector('.content').textContent.match(/\((.+)\)/);
        if (match) showAggregateMenu(e, match[1]);
    }
});

// 聚合菜单
function showAggregateMenu(event, field) {
    event.stopPropagation();
    
    document.querySelectorAll('.popup-menu').forEach(menu => menu.remove());
    
    const menu = document.createElement('div');
    menu.className = 'popup-menu';
    
    Object.entries(CONFIG.aggregateMethods).forEach(([method, label]) => {
        const item = document.createElement('div');
        item.className = 'popup-menu-item';
        if (state.aggregateFields[field] === method) item.classList.add('active');
        
        item.textContent = label;
        item.onclick = () => {
            state.aggregateFields[field] = method;
            dragDropManager.renderFields('aggregate');
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
    setTimeout(() => document.addEventListener('click', () => menu.remove(), { once: true }), 0);
}

// 移除字段
function removeField(type, field) {
    const actions = {
        group: () => state.groupFields = state.groupFields.filter(f => f !== field),
        aggregate: () => delete state.aggregateFields[field],
        basicInfo: () => state.basicInfoFields = state.basicInfoFields.filter(f => f !== field),
        calc: () => delete state.calcFields[field]
    };
    
    if (actions[type]) {
        actions[type]();
        type === 'calc' ? renderCalcFields() : dragDropManager.renderFields(type);
        checkButtonStates();
    }
}

// 检查按钮状态
function checkButtonStates() {
    document.getElementById('processBtn').disabled = !(
        state.selectedFiles.length > 0 &&
        document.getElementById('headerRowSelect').value &&
        state.groupFields.length > 0
    );
}

let dragDropManager = null;
let fieldsListManager = null;