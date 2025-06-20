// 匹配数据功能模块

// 匹配文件选择处理
async function handleMatchFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        elements.fieldTagsArea.innerHTML = '<div style="color: #999; padding: 20px; text-align: center;">读取字段中......</div>';
        
        const workbook = await ExcelProcessor.readFile(file);
        const worksheet = workbook.getWorksheet(1);
        const headerInfo = ExcelProcessor.findHeaderRow(worksheet);
        
        // 读取数据
        const dataMap = new Map();
        for (let rowNum = headerInfo.rowNum + 1; rowNum <= worksheet.rowCount; rowNum++) {
            const row = worksheet.getRow(rowNum);
            const id = String(row.getCell(headerInfo.idColumn).value || '').trim();
            if (id) {
                dataMap.set(id, row.values.slice(1));
            }
        }
        
        state.matchFileData = {
            headers: headerInfo.headers,
            idColumn: headerInfo.idColumn,
            dataMap,
            rowCount: dataMap.size
        };
        
        elements.matchFilePath.value = file.name;
        
        // 显示字段标签
        elements.fieldTagsArea.innerHTML = '';
        headerInfo.headers.forEach((header, index) => {
            if (index !== headerInfo.idColumn - 1 && header) {
                const tag = document.createElement('div');
                tag.className = 'tag';
                tag.textContent = header;
                tag.dataset.index = index;
                tag.addEventListener('click', toggleFieldTag);
                elements.fieldTagsArea.appendChild(tag);
            }
        });
        
    } catch (error) {
        utils.showError(`读取匹配文件失败：${error.message}`);
        utils.showNotification('读取文件失败', 'error');
    }
}

// 切换字段标签选中状态
function toggleFieldTag(event) {
    const tag = event.target;
    const index = parseInt(tag.dataset.index);
    
    if (tag.classList.contains('selected')) {
        tag.classList.remove('selected');
        state.selectedFieldIndices = state.selectedFieldIndices.filter(i => i !== index);
    } else {
        tag.classList.add('selected');
        state.selectedFieldIndices.push(index);
    }
}

// 确认匹配
async function confirmMatch() {
    if (!state.matchFileData || state.selectedFieldIndices.length === 0) {
        utils.showNotification('请选择文件和字段', 'error');
        return;
    }
    
    elements.matchModal.style.display = 'none';
    utils.setLoadingState(elements.matchBtn, true);
    
    try {
        const worksheet = state.summaryWorkbook.getWorksheet('汇总');
        const lastCol = worksheet.columnCount;
        const fieldsToAdd = state.selectedFieldIndices.map(i => ({
            index: i,
            header: state.matchFileData.headers[i]
        }));
        
        // 添加标题
        const newStartCol = lastCol + 1;
        const newEndCol = lastCol + fieldsToAdd.length;
        
        worksheet.getCell(1, newStartCol).value = '其它信息';
        if (fieldsToAdd.length > 1) {
            worksheet.mergeCells(1, newStartCol, 1, newEndCol);
        }
        
        // 添加字段名和样式
        const headerStyle = {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BBB59' } },
            font: { color: { argb: 'FFFFFFFF' }, bold: true },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
        };
        
        for (let col = newStartCol; col <= newEndCol; col++) {
            Object.assign(worksheet.getCell(1, col), headerStyle);
        }
        
        fieldsToAdd.forEach((field, i) => {
            const cell = worksheet.getCell(2, newStartCol + i);
            cell.value = field.header;
            Object.assign(cell, headerStyle);
        });
        
        // 匹配数据
        for (let row = 3; row <= worksheet.rowCount; row++) {
            const drugId = String(worksheet.getCell(row, 2).value || '').trim();
            if (drugId && state.matchFileData.dataMap.has(drugId)) {
                const matchedRow = state.matchFileData.dataMap.get(drugId);
                fieldsToAdd.forEach((field, i) => {
                    worksheet.getCell(row, newStartCol + i).value = matchedRow[field.index] || '';
                });
            }
        }
        
        // 设置列宽
        for (let col = newStartCol; col <= newEndCol; col++) {
            worksheet.getColumn(col).width = 12;
        }
        
        utils.showNotification('数据匹配成功');
        
    } catch (error) {
        utils.showNotification(`匹配失败：${error.message}`, 'error');
    } finally {
        utils.setLoadingState(elements.matchBtn, false, '匹配数据');
    }
}

// 显示匹配弹窗
function showMatchModal() {
    if (!state.summaryWorkbook) {
        utils.showNotification('请先处理数据', 'error');
        return;
    }
    state.matchFileData = null;
    state.selectedFieldIndices = [];
    elements.matchFilePath.value = '';
    elements.fieldTagsArea.innerHTML = '';
    elements.matchModal.style.display = 'flex';
}

// 关闭匹配弹窗
function closeMatchModal() {
    elements.matchModal.style.display = 'none';
    elements.matchFileInput.value = '';
}