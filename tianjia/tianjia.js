// 添加字段模块
const MatchModule = {
    fieldSelectors: {
        matchFields: null,
        addFields: null
    },
    
    init() {
        this.fieldSelectors.matchFields = new FieldSelector('matchFieldsArea');
        this.fieldSelectors.addFields = new FieldSelector('addFieldsArea');
        this.fieldSelectors.matchFields.useIndices = true;
        this.fieldSelectors.addFields.useIndices = true;
        this.bindEvents();
    },
    
    bindEvents() {
        const selectMatchFileBtn = document.getElementById('selectMatchFileBtn');
        const matchFileInput = document.getElementById('matchFileInput');
        const matchHeaderRowSelect = document.getElementById('matchHeaderRowSelect');
        const cancelMatchBtn = document.getElementById('cancelMatchBtn');
        const confirmMatchBtn = document.getElementById('confirmMatchBtn');
        
        selectMatchFileBtn.addEventListener('click', () => matchFileInput.click());
        matchFileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        matchHeaderRowSelect.addEventListener('change', () => this.handleHeaderRowSelect());
        cancelMatchBtn.addEventListener('click', () => this.closeModal());
        confirmMatchBtn.addEventListener('click', () => this.confirmMatch());
    },
    
    async handleFileSelect(event) {
        await FormHandler.handleFileSelect(
            event.target,
            document.getElementById('matchFilePath'),
            file => /\.(xlsx?|xls)$/i.test(file.name),
            async (validFiles) => {
                try {
                    const workbook = await ExcelReader.readFile(validFiles[0]);
                    const worksheet = workbook.getWorksheet(1);
                    
                    state.matchFileData = { workbook, worksheet, file: validFiles[0] };
                    document.getElementById('matchHeaderRowSelect').value = '';
                    
                    this.fieldSelectors.matchFields.showStatus('请选择标题行');
                    this.fieldSelectors.addFields.showStatus('请选择标题行');
                } catch (error) {
                    utils.showError(`读取匹配文件失败：${error.message}`);
                }
            }
        );
    },
    
    async handleHeaderRowSelect() {
        const matchHeaderRowSelect = document.getElementById('matchHeaderRowSelect');
        
        if (!state.matchFileData || !matchHeaderRowSelect.value) {
            this.fieldSelectors.matchFields.reset();
            this.fieldSelectors.addFields.reset();
            return;
        }

        try {
            this.fieldSelectors.matchFields.showStatus('读取字段中......');
            this.fieldSelectors.addFields.showStatus('读取字段中......');

            const headerRowNum = parseInt(matchHeaderRowSelect.value);
            const { worksheet } = state.matchFileData;
            
            const fields = ExcelReader.extractFields(worksheet, headerRowNum);
            const dataMap = ExcelReader.buildDataMap(worksheet, headerRowNum, fields);

            state.matchFileData = {
                ...state.matchFileData,
                headers: fields,
                headerRowNum,
                dataMap,
                rowCount: dataMap.size
            };

            this.fieldSelectors.matchFields.setFields(fields, true);
            this.fieldSelectors.addFields.setFields(fields, true);
            
            if (state.matchConfig && state.matchConfig.fileName === state.matchFileData.file.name) {
                this.fieldSelectors.matchFields.selected = [...state.matchConfig.selectedMatchIndices];
                this.fieldSelectors.addFields.selected = [...state.matchConfig.selectedAddIndices];
                this.fieldSelectors.matchFields.render();
                this.fieldSelectors.addFields.render();
            } else {
                const groupFields = window.fieldSelectors.group.getSelected();
                this.fieldSelectors.matchFields.setDefaultSelected(groupFields);
            }
            
        } catch (error) {
            this.fieldSelectors.matchFields.showStatus('读取字段失败');
            this.fieldSelectors.addFields.showStatus('读取字段失败');
            console.error('读取匹配字段失败：', error);
        }
    },
    
    confirmMatch() {
        const selectedMatchIndices = this.fieldSelectors.matchFields.getSelected();
        const selectedAddIndices = this.fieldSelectors.addFields.getSelected();
        
        if (!state.matchFileData || selectedMatchIndices.length === 0 || selectedAddIndices.length === 0) {
            notification.show('请选择文件、标题行、匹配字段和添加字段', 'error');
            return;
        }
        
        state.matchConfig = {
            fileName: state.matchFileData.file.name,
            headers: state.matchFileData.headers,
            dataMap: state.matchFileData.dataMap,
            selectedMatchIndices: [...selectedMatchIndices],
            selectedAddIndices: [...selectedAddIndices],
            fieldsToAdd: selectedAddIndices.map(i => ({
                index: i,
                header: state.matchFileData.headers[i]
            }))
        };
        
        document.getElementById('matchModal').style.display = 'none';
        this.updateButtonText();
        notification.show(`匹配配置已保存：${state.matchConfig.fieldsToAdd.length}个字段`);
    },
    
    showModal() {
        const matchFileInput = document.getElementById('matchFileInput');
        matchFileInput.value = '';
        
        if (state.matchConfig) {
            document.getElementById('matchFilePath').value = state.matchConfig.fileName;
        } else {
            state.matchFileData = null;
            this.fieldSelectors.matchFields.reset();
            this.fieldSelectors.addFields.reset();
            document.getElementById('matchFilePath').value = '';
            document.getElementById('matchHeaderRowSelect').value = '';
        }
        
        document.getElementById('matchModal').style.display = 'flex';
    },
    
    closeModal() {
        document.getElementById('matchModal').style.display = 'none';
    },
    
    updateButtonText() {
        const matchBtn = document.getElementById('matchBtn');
        const matchConfigHint = document.getElementById('matchConfigHint');
        
        if (state.matchConfig) {
            matchBtn.textContent = '重新添加';
            matchConfigHint.style.display = 'block';
            matchConfigHint.textContent = `已配置匹配：${state.matchConfig.fileName} (${state.matchConfig.fieldsToAdd.length}个字段)`;
        } else {
            matchBtn.textContent = '添加字段';
            matchConfigHint.style.display = 'none';
        }
    }
};