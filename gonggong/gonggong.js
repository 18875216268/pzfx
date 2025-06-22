// 全局配置
const CONFIG = {
    numericFields: [
        "销售数量", "客户数", "含税出库金额", 
        "P4成本金额", "P4毛利额", "应收边际利润额(不含税)"
    ],
    dataFields: [
        '含税出库金额', 'P4毛利额', 'P4毛利率', 
        '应收边际利润额(不含税)', '边际利润率', '客户数'
    ]
};

// 全局状态
const state = {
    selectedFiles: [],
    summaryWorkbook: null,
    processedData: {},
    matchFileData: null,
    allFields: [],
    matchConfig: null
};

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
    
    setLoadingState: (button, loading, originalText = '') => {
        if (loading) {
            if (!originalText) {
                button.dataset.originalText = button.textContent;
            }
            button.classList.add('btn-loading');
            button.innerHTML = '<span class="loading"></span>';
        } else {
            button.classList.remove('btn-loading');
            button.textContent = originalText || button.dataset.originalText || '按钮';
        }
    },
    
    showError: message => {
        const errorSection = document.getElementById('errorSection');
        const errorMessage = document.getElementById('errorMessage');
        errorSection.style.display = 'block';
        errorMessage.textContent = message;
    },
    
    hideError: () => {
        const errorSection = document.getElementById('errorSection');
        errorSection.style.display = 'none';
    }
};

// Excel读取器
class ExcelReader {
    static async readFile(file) {
        const workbook = new ExcelJS.Workbook();
        const arrayBuffer = await file.arrayBuffer();
        await workbook.xlsx.load(arrayBuffer);
        return workbook;
    }

    static extractFields(worksheet, headerRowNum) {
        const headerRow = worksheet.getRow(headerRowNum);
        const fields = [];
        
        headerRow.eachCell((cell, colNum) => {
            const header = String(cell.value || '').trim();
            if (header) {
                fields.push(header);
            }
        });
        
        return fields;
    }

    static buildDataMap(worksheet, headerRowNum, headers) {
        const dataMap = new Map();
        const keyColumnIndex = 0;
        
        for (let rowNum = headerRowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const key = String(row.getCell(keyColumnIndex + 1).value || '').trim();
            
            if (key) {
                const rowData = headers.map((_, index) => 
                    row.getCell(index + 1).value || ''
                );
                dataMap.set(key, rowData);
            }
        }
        
        return dataMap;
    }
}

// 字段选择器
class FieldSelector {
    constructor(areaId, onChange = null) {
        this.areaId = areaId;
        this.fields = [];
        this.selected = [];
        this.onChange = onChange;
        this.useIndices = false;
        this.dragState = {
            dragging: null,
            dragOver: null
        };
    }

    setFields(fields, useIndices = false) {
        this.fields = [...fields];
        this.selected = [];
        this.useIndices = useIndices;
        this.render();
    }

    getSelected() {
        if (this.useIndices) {
            return [...this.selected];
        } else {
            const selectedFields = [];
            this.fields.forEach(field => {
                if (this.selected.includes(field)) {
                    selectedFields.push(field);
                }
            });
            return selectedFields;
        }
    }

    getOrderedFields() {
        return this.fields.map((field, index) => ({
            field,
            identifier: this.useIndices ? index : field,
            isSelected: this.selected.includes(this.useIndices ? index : field)
        }));
    }

    toggle(identifier) {
        const index = this.selected.indexOf(identifier);
        if (index > -1) {
            this.selected.splice(index, 1);
        } else {
            this.selected.push(identifier);
        }
        this.updateTagState(identifier);
        if (this.onChange) this.onChange(this.selected);
    }

    setDefaultSelected(fieldsToSelect) {
        this.selected = [];
        fieldsToSelect.forEach(field => {
            const index = this.fields.indexOf(field);
            if (index !== -1) {
                const identifier = this.useIndices ? index : field;
                this.selected.push(identifier);
            }
        });
        this.render();
    }

    updateTagState(identifier) {
        const area = document.getElementById(this.areaId);
        const tag = area.querySelector(`[data-identifier="${identifier}"]`);
        if (tag) {
            tag.classList.toggle('selected', this.selected.includes(identifier));
        }
    }

    moveField(fromIndex, toIndex) {
        if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || 
            fromIndex >= this.fields.length || toIndex >= this.fields.length) {
            return;
        }

        const [movedField] = this.fields.splice(fromIndex, 1);
        this.fields.splice(toIndex, 0, movedField);

        if (this.useIndices) {
            const newSelected = [];
            this.selected.forEach(selectedIndex => {
                let newIndex = selectedIndex;
                if (selectedIndex === fromIndex) {
                    newIndex = toIndex;
                } else if (selectedIndex > fromIndex && selectedIndex <= toIndex) {
                    newIndex = selectedIndex - 1;
                } else if (selectedIndex < fromIndex && selectedIndex >= toIndex) {
                    newIndex = selectedIndex + 1;
                }
                newSelected.push(newIndex);
            });
            this.selected = newSelected;
        }

        this.render();
        if (this.onChange) this.onChange(this.selected);
    }

    getTagPosition(tag) {
        const rect = tag.getBoundingClientRect();
        const containerRect = document.getElementById(this.areaId).getBoundingClientRect();
        return {
            x: rect.left - containerRect.left,
            y: rect.top - containerRect.top,
            width: rect.width,
            height: rect.height
        };
    }

    render() {
        const area = document.getElementById(this.areaId);
        if (!area) return;

        if (this.fields.length === 0) {
            area.innerHTML = '<div class="placeholder">暂无可用字段</div>';
            return;
        }

        area.innerHTML = '';
        
        this.fields.forEach((field, index) => {
            const identifier = this.useIndices ? index : field;
            const tag = document.createElement('div');
            tag.className = 'tag';
            tag.textContent = field;
            tag.dataset.identifier = identifier;
            tag.dataset.index = index;
            tag.draggable = true;
            
            if (this.selected.includes(identifier)) {
                tag.classList.add('selected');
            }
            
            tag.addEventListener('click', (e) => {
                if (!tag.classList.contains('dragging')) {
                    this.toggle(identifier);
                }
            });

            tag.addEventListener('dragstart', (e) => {
                this.dragState.dragging = tag;
                tag.classList.add('dragging');
                area.classList.add('drag-active');
                e.dataTransfer.effectAllowed = 'move';
                e.dataTransfer.setData('text/html', tag.outerHTML);
            });

            tag.addEventListener('dragend', (e) => {
                tag.classList.remove('dragging');
                area.classList.remove('drag-active');
                this.clearDragStyles();
                this.dragState.dragging = null;
                this.dragState.dragOver = null;
            });

            tag.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                
                if (this.dragState.dragging && this.dragState.dragging !== tag) {
                    this.clearDragStyles();
                    tag.classList.add('drag-over');
                    this.dragState.dragOver = tag;
                }
            });

            tag.addEventListener('drop', (e) => {
                e.preventDefault();
                
                if (this.dragState.dragging && this.dragState.dragging !== tag) {
                    const fromIndex = parseInt(this.dragState.dragging.dataset.index);
                    const toIndex = parseInt(tag.dataset.index);
                    
                    const pos = this.getTagPosition(tag);
                    const mouseX = e.clientX - area.getBoundingClientRect().left;
                    const insertBefore = mouseX < pos.x + pos.width / 2;
                    
                    let targetIndex = insertBefore ? toIndex : toIndex + 1;
                    if (fromIndex < toIndex && !insertBefore) {
                        targetIndex = toIndex;
                    } else if (fromIndex > toIndex && insertBefore) {
                        targetIndex = toIndex;
                    } else if (fromIndex < toIndex && insertBefore) {
                        targetIndex = toIndex - 1;
                    }
                    
                    this.moveField(fromIndex, targetIndex);
                }
            });
            
            area.appendChild(tag);
        });
    }

    clearDragStyles() {
        const area = document.getElementById(this.areaId);
        if (area) {
            area.querySelectorAll('.tag.drag-over').forEach(tag => {
                tag.classList.remove('drag-over');
            });
        }
    }

    showStatus(message) {
        const area = document.getElementById(this.areaId);
        if (area) {
            area.innerHTML = `<div class="placeholder">${message}</div>`;
        }
    }

    reset() {
        this.fields = [];
        this.selected = [];
        this.showStatus('请先选择文件和标题行');
    }
}

// 表单处理器
class FormHandler {
    static async handleFileSelect(fileInput, pathInput, validator, onSuccess) {
        const files = Array.from(fileInput.files);
        const validFiles = validator ? files.filter(validator) : files;
        
        if (validFiles.length > 0) {
            pathInput.value = validFiles.map(f => f.name).join(', ');
            onSuccess(validFiles);
        } else {
            pathInput.value = '';
            if (files.length > 0) {
                utils.showError('文件格式不符合要求');
            }
        }
    }

    static async handleHeaderRowSelect(selectedFiles, headerRowSelect, fieldSelectors) {
        if (!selectedFiles.length || !headerRowSelect.value) {
            fieldSelectors.forEach(selector => selector.reset());
            return;
        }

        fieldSelectors.forEach(selector => selector.showStatus('读取字段中......'));

        try {
            const firstFile = selectedFiles[0];
            const headerRowNum = parseInt(headerRowSelect.value);
            const workbook = await ExcelReader.readFile(firstFile);
            const worksheet = workbook.getWorksheet(1);
            
            const fields = ExcelReader.extractFields(worksheet, headerRowNum);
            
            fieldSelectors.forEach(selector => {
                selector.setFields(fields, selector.useIndices);
            });

            return { workbook, worksheet, headerRowNum, fields };
        } catch (error) {
            fieldSelectors.forEach(selector => selector.showStatus('读取字段失败'));
            console.error('读取字段失败：', error);
            throw error;
        }
    }
}