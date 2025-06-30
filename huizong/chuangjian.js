/**
 * 计算字段创建模块 - 重构：使用统一的显示格式
 */
const CalcFieldModule = {
    // 临时存储的计算字段
    tempCalcFields: [],
    
    // 当前编辑的内容
    currentInput: '',
    
    /**
     * 初始化模块
     */
    init() {
        this.bindEvents();
        this.renderCalcFields();
    },
    
    /**
     * 绑定事件
     */
    bindEvents() {
        // 添加按钮点击
        document.getElementById('addCalcFieldBtn').addEventListener('click', () => {
            this.showModal();
        });
        
        // 弹窗关闭按钮
        document.querySelector('.calc-field-modal-close').addEventListener('click', () => {
            this.hideModal();
        });
        
        // 取消按钮
        document.getElementById('calcFieldCancelBtn').addEventListener('click', () => {
            this.hideModal();
        });
        
        // 添加按钮
        document.getElementById('calcFieldAddBtn').addEventListener('click', () => {
            this.addCalcField();
        });
        
        // 确认按钮
        document.getElementById('calcFieldConfirmBtn').addEventListener('click', () => {
            this.confirmCalcFields();
        });
        
        // 输入框事件
        const input = document.getElementById('calcFieldInput');
        
        // 处理Enter键 - 换行
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.processInput();
            }
        });
        
        // 处理输入变化
        input.addEventListener('input', () => {
            this.currentInput = input.innerHTML;
        });
        
        // 点击弹窗背景关闭
        document.getElementById('calcFieldModal').addEventListener('click', (e) => {
            if (e.target.id === 'calcFieldModal') {
                this.hideModal();
            }
        });
        
        // 主界面计算字段删除和拖拽
        this.initMainCalcFieldEvents();
    },
    
    /**
     * 显示弹窗
     */
    showModal() {
        // 初始化临时字段列表
        this.tempCalcFields = [...state.calcFields];
        
        // 渲染字段列表
        this.renderModalFieldsList();
        
        // 渲染已创建字段
        this.renderModalCreatedFields();
        
        // 清空输入框
        document.getElementById('calcFieldInput').innerHTML = '';
        this.currentInput = '';
        
        // 显示弹窗
        document.getElementById('calcFieldModal').style.display = 'flex';
    },
    
    /**
     * 隐藏弹窗
     */
    hideModal() {
        document.getElementById('calcFieldModal').style.display = 'none';
    },
    
    /**
     * 渲染弹窗中的字段列表 - 显示聚合字段
     */
    renderModalFieldsList() {
        const container = document.getElementById('calcModalFieldsList');
        
        if (!state.aggregateFields.length) {
            container.innerHTML = '<div class="placeholder">暂无聚合字段</div>';
            return;
        }
        
        // 直接使用聚合字段的显示名称
        container.innerHTML = state.aggregateFields
            .map(displayName => {
                return `
                    <div class="field-item" data-display="${displayName}">
                        <span class="content">${displayName}</span>
                    </div>
                `;
            }).join('');
        
        // 绑定点击事件 - 插入显示名称
        container.querySelectorAll('.field-item').forEach(item => {
            item.addEventListener('click', () => {
                this.insertFieldTag(item.dataset.display);
            });
        });
    },
    
    /**
     * 渲染弹窗中的已创建字段
     */
    renderModalCreatedFields() {
        const container = document.getElementById('calcModalCreatedFields');
        
        if (!this.tempCalcFields.length) {
            container.innerHTML = '<div class="placeholder">暂无创建的字段</div>';
            return;
        }
        
        container.innerHTML = this.tempCalcFields
            .map((field, index) => {
                return `
                    <div class="field-item calc-created-field-item draggable removable" 
                         draggable="true" 
                         data-index="${index}"
                         data-display="${field.name}">
                        <div class="content">
                            <div class="field-name">${field.name}</div>
                            <div class="field-formula">${field.formula}</div>
                        </div>
                        <span class="remove-btn" data-index="${index}">×</span>
                    </div>
                `;
            }).join('');
        
        // 绑定事件
        this.bindModalCreatedFieldEvents();
    },
    
    /**
     * 绑定已创建字段的事件
     */
    bindModalCreatedFieldEvents() {
        const container = document.getElementById('calcModalCreatedFields');
        
        // 删除按钮
        container.querySelectorAll('.remove-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const index = parseInt(btn.dataset.index);
                this.tempCalcFields.splice(index, 1);
                this.renderModalCreatedFields();
            });
        });
        
        // 点击插入字段
        container.querySelectorAll('.calc-created-field-item').forEach(item => {
            item.addEventListener('click', (e) => {
                if (!e.target.classList.contains('remove-btn')) {
                    this.insertFieldTag(item.dataset.display);
                }
            });
        });
        
        // 拖拽排序
        this.initModalDragSort(container);
    },
    
    /**
     * 初始化弹窗中的拖拽排序
     */
    initModalDragSort(container) {
        let draggedElement = null;
        let draggedIndex = null;
        
        // 创建拖拽指示器
        const indicator = document.createElement('div');
        indicator.className = 'drop-indicator';
        
        container.querySelectorAll('.calc-created-field-item').forEach(item => {
            item.addEventListener('dragstart', (e) => {
                draggedElement = item;
                draggedIndex = parseInt(item.dataset.index);
                item.style.opacity = '0.5';
            });
            
            item.addEventListener('dragend', () => {
                item.style.opacity = '';
                indicator.remove();
            });
            
            item.addEventListener('dragover', (e) => {
                e.preventDefault();
                if (item !== draggedElement) {
                    const box = item.getBoundingClientRect();
                    const afterElement = e.clientY > box.top + box.height / 2;
                    
                    if (afterElement) {
                        item.parentNode.insertBefore(indicator, item.nextSibling);
                    } else {
                        item.parentNode.insertBefore(indicator, item);
                    }
                }
            });
        });
        
        container.addEventListener('dragover', (e) => {
            e.preventDefault();
            
            // 如果容器内没有元素时的处理
            const items = container.querySelectorAll('.calc-created-field-item');
            if (items.length === 0 && draggedElement) {
                container.appendChild(indicator);
            }
        });
        
        container.addEventListener('drop', (e) => {
            e.preventDefault();
            const afterElement = indicator.nextElementSibling;
            indicator.remove();
            
            if (!draggedElement) return;
            
            // 获取拖拽的字段对象
            const draggedField = this.tempCalcFields[draggedIndex];
            
            // 从原数组中移除
            this.tempCalcFields.splice(draggedIndex, 1);
            
            // 计算新位置
            let newIndex = 0;
            if (afterElement) {
                // 找到目标位置
                const allItems = Array.from(container.querySelectorAll('.calc-created-field-item'));
                const targetIndex = allItems.indexOf(afterElement);
                
                // 调整索引（因为已经移除了一个元素）
                newIndex = draggedIndex < targetIndex ? targetIndex - 1 : targetIndex;
            } else {
                // 拖到最后
                newIndex = this.tempCalcFields.length;
            }
            
            // 插入到新位置
            this.tempCalcFields.splice(newIndex, 0, draggedField);
            
            // 重新渲染（这会更新所有的 data-index）
            this.renderModalCreatedFields();
        });
        
        container.addEventListener('dragleave', (e) => {
            if (!container.contains(e.relatedTarget)) {
                indicator.remove();
            }
        });
    },
    
    /**
     * 插入字段标签到输入框
     */
    insertFieldTag(fieldName) {
        const input = document.getElementById('calcFieldInput');
        
        // 确保输入框获得焦点
        input.focus();
        
        const selection = window.getSelection();
        let range;
        
        // 检查当前选区是否在输入框内
        if (selection.rangeCount > 0) {
            const currentRange = selection.getRangeAt(0);
            const container = currentRange.commonAncestorContainer;
            
            // 判断选区是否在输入框内
            if (input.contains(container) || container === input) {
                range = currentRange;
            } else {
                // 如果不在输入框内，将光标移到输入框末尾
                range = document.createRange();
                range.selectNodeContents(input);
                range.collapse(false); // false = 光标在末尾
                selection.removeAllRanges();
                selection.addRange(range);
            }
        } else {
            // 没有选区时，将光标设置到输入框末尾
            range = document.createRange();
            range.selectNodeContents(input);
            range.collapse(false);
            selection.removeAllRanges();
            selection.addRange(range);
        }
        
        // 创建标签元素
        const tag = document.createElement('span');
        tag.className = 'field-tag';
        tag.textContent = fieldName;
        tag.contentEditable = 'false';
        
        // 在光标位置插入
        range.insertNode(tag);
        
        // 移动光标到标签后
        const newRange = document.createRange();
        newRange.setStartAfter(tag);
        newRange.setEndAfter(tag);
        selection.removeAllRanges();
        selection.addRange(newRange);
        
        // 更新当前输入
        this.currentInput = input.innerHTML;
    },
    
    /**
     * 处理输入（Enter键）- 支持多行
     */
    processInput() {
        const input = document.getElementById('calcFieldInput');
        
        // 获取光标位置
        const selection = window.getSelection();
        if (selection.rangeCount === 0) return;
        
        const range = selection.getRangeAt(0);
        
        // 删除选中的内容（如果有）
        range.deleteContents();
        
        // 创建一个包含换行的文档片段
        const fragment = document.createDocumentFragment();
        
        // 插入一个 br 标签
        const br = document.createElement('br');
        fragment.appendChild(br);
        
        // 在 br 后面插入一个零宽空格，确保光标能正确定位
        const textNode = document.createTextNode('\u200B'); // 零宽空格
        fragment.appendChild(textNode);
        
        // 插入片段
        range.insertNode(fragment);
        
        // 将光标移动到零宽空格之后
        range.setStartAfter(textNode);
        range.setEndAfter(textNode);
        selection.removeAllRanges();
        selection.addRange(range);
        
        // 更新当前输入
        this.currentInput = input.innerHTML;
    },
    
    /**
     * 添加计算字段
     */
    addCalcField() {
        const input = document.getElementById('calcFieldInput');
        const content = input.innerHTML;
        
        if (!content.trim()) {
            notification.show('请输入表达式', 'error');
            return;
        }
        
        // 获取所有文本内容，按行分割
        const processNode = (node) => {
            if (node.nodeType === Node.TEXT_NODE) {
                return node.textContent;
            } else if (node.nodeName === 'BR') {
                return '\n';
            } else if (node.classList && node.classList.contains('field-tag')) {
                return node.textContent;
            } else {
                return Array.from(node.childNodes).map(processNode).join('');
            }
        };
        
        const fullText = Array.from(input.childNodes).map(processNode).join('');
        const expressions = fullText.split('\n').filter(line => line.trim());
        
        if (!expressions.length) {
            notification.show('请输入表达式', 'error');
            return;
        }
        
        let successCount = 0;
        let errorMessages = [];
        
        expressions.forEach((expression, index) => {
            try {
                // 解析表达式
                const parsed = CalcEngine.parseExpression(expression.trim());
                
                // 准备可用字段列表
                const availableFields = [
                    ...state.aggregateFields,
                    ...this.tempCalcFields.map(f => f.name)
                ];
                
                // 验证字段
                const validation = CalcEngine.validateCalcField(parsed, availableFields);
                if (!validation.isValid) {
                    errorMessages.push(`第${index + 1}行: ${validation.errors.join(', ')}`);
                    return;
                }
                
                // 检查是否已存在同名字段
                if (this.tempCalcFields.some(f => f.name === parsed.name)) {
                    errorMessages.push(`第${index + 1}行: 字段名"${parsed.name}"已存在`);
                    return;
                }
                
                // 添加到临时列表
                parsed.id = Date.now() + Math.random() + index;
                this.tempCalcFields.push(parsed);
                successCount++;
                
            } catch (error) {
                errorMessages.push(`第${index + 1}行: ${error.message}`);
            }
        });
        
        // 刷新列表
        if (successCount > 0) {
            this.renderModalCreatedFields();
            
            // 清空输入框
            input.innerHTML = '';
            this.currentInput = '';
            
            notification.show(`成功添加 ${successCount} 个字段`);
        }
        
        // 显示错误信息
        if (errorMessages.length > 0) {
            notification.show(errorMessages[0], 'error'); // 只显示第一个错误
        }
    },
    
    /**
     * 确认计算字段
     */
    confirmCalcFields() {
        // 更新状态
        state.calcFields = [...this.tempCalcFields];
        
        // 刷新主界面
        this.renderCalcFields();
        
        // 关闭弹窗
        this.hideModal();
        
        // 检查按钮状态
        checkButtonStates();
        
        notification.show('计算字段已更新');
    },
    
    /**
     * 渲染主界面的计算字段
     */
    /**
     * 渲染主界面的计算字段
     */
    renderCalcFields() {
        const container = document.getElementById('calcFieldsInput');
        
        if (!state.calcFields || !state.calcFields.length) {
            container.innerHTML = '<div class="placeholder">点击右上角"+"创建计算字段</div>';
            return;
        }
        
        container.innerHTML = state.calcFields
            .map(field => {
                return `
                    <div class="field-item calc-created-field-item draggable removable" 
                        draggable="true" 
                        data-field="${field.name}" 
                        data-type="calculated" 
                        data-display="${field.name}"
                        data-id="${field.id}">
                        <div class="content">
                            <div class="field-name">${field.name}</div>
                            <div class="field-formula">${field.formula}</div>
                        </div>
                        <span class="remove-btn">×</span>
                    </div>
                `;
            }).join('');
        
        // 刷新弹窗中的已创建字段（如果弹窗是打开的）
        if (document.getElementById('calcFieldModal').style.display === 'flex') {
            this.renderModalCreatedFields();
        }
    },
    
    /**
     * 初始化主界面计算字段事件
     */
    initMainCalcFieldEvents() {
        const container = document.getElementById('calcFieldsInput');
        
        // 防止其他区域的字段拖入
        container.addEventListener('dragover', (e) => {
            const dragData = dragManager.dragData;
            if (dragData && dragData.fromType !== 'calculated') {
                e.preventDefault();
                e.dataTransfer.effectAllowed = 'none';
                e.dataTransfer.dropEffect = 'none';
            }
        });
        
        container.addEventListener('drop', (e) => {
            const dragData = dragManager.dragData;
            if (dragData && dragData.fromType !== 'calculated') {
                e.preventDefault();
                e.stopPropagation();
            }
        });
        
        // 处理删除事件
        document.addEventListener('click', (e) => {
            if (e.target.classList.contains('remove-btn')) {
                const item = e.target.closest('.calc-created-field-item');
                if (item && item.closest('#calcFieldsInput')) {
                    const fieldName = item.dataset.display;
                    state.calcFields = state.calcFields.filter(f => f.name !== fieldName);
                    this.renderCalcFields();
                    checkButtonStates();
                }
            }
        });
    }
};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
    CalcFieldModule.init();
});