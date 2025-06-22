// 填充模块
const FillModule = {
    createSummaryWorkbook(processedData, groupFields, basicInfoFields, allFields, matchConfig) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('汇总');
        
        const drugInfo = AggregateModule.extractAllDrugInfo(processedData, basicInfoFields);
        const groupKeys = AggregateModule.getAllGroupKeys(processedData);
        const fileInfo = AggregateModule.prepareFileInfo(processedData);
        const periodData = AggregateModule.aggregateByPeriod(fileInfo);
        
        this.createHeaders(worksheet, fileInfo, basicInfoFields, matchConfig);
        
        const matchKeyMap = new Map();
        
        let matchFileKeyMap = null;
        if (matchConfig) {
            matchFileKeyMap = new Map();
            for (const [key, rowData] of matchConfig.dataMap) {
                const matchKeyParts = matchConfig.selectedMatchIndices.map(index => 
                    String(rowData[index] || '').trim()
                ).filter(part => part);
                
                const matchKey = matchKeyParts.join('|');
                if (matchKey) {
                    matchFileKeyMap.set(matchKey, rowData);
                }
            }
        }
        
        groupKeys.forEach(groupKey => {
            const { row, matchKey } = CalculateModule.buildDataRow(
                groupKey, drugInfo[groupKey], periodData, basicInfoFields, 
                matchConfig, matchFileKeyMap
            );
            const rowIndex = worksheet.addRow(row).number;
            matchKeyMap.set(rowIndex, matchKey);
        });
        
        this.applyStyles(worksheet, fileInfo.length, basicInfoFields.length, matchConfig);
        
        workbook.matchKeyMap = matchKeyMap;
        
        return workbook;
    },
    
    createHeaders(worksheet, fileInfo, basicInfoFields, matchConfig) {
        const row1 = ['药品基础信息', ...Array(basicInfoFields.length - 1).fill('')];
        fileInfo.forEach(f => {
            const fileName = f.fileName.replace(/\.[^/.]+$/, '');
            row1.push(fileName, ...Array(5).fill(''));
        });
        row1.push('同比', ...Array(5).fill(''), '环比', ...Array(5).fill(''));
        
        if (matchConfig) {
            row1.push('其它信息', ...Array(matchConfig.fieldsToAdd.length - 1).fill(''));
        }
        
        worksheet.addRow(row1);
        
        const row2 = [...basicInfoFields];
        const repeatCount = fileInfo.length + 2;
        for (let i = 0; i < repeatCount; i++) {
            row2.push(...CONFIG.dataFields);
        }
        
        if (matchConfig) {
            matchConfig.fieldsToAdd.forEach(field => {
                row2.push(field.header);
            });
        }
        
        worksheet.addRow(row2);
    },
    
    applyStyles(worksheet, fileCount, basicInfoFieldsCount, matchConfig) {
        worksheet.mergeCells(1, 1, 1, basicInfoFieldsCount);
        let startCol = basicInfoFieldsCount + 1;
        for (let i = 0; i < fileCount + 2; i++) {
            worksheet.mergeCells(1, startCol, 1, startCol + 5);
            startCol += 6;
        }
        
        if (matchConfig && matchConfig.fieldsToAdd.length > 1) {
            worksheet.mergeCells(1, startCol, 1, startCol + matchConfig.fieldsToAdd.length - 1);
        }
        
        const headerStyle = {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BBB59' } },
            font: { color: { argb: 'FFFFFFFF' }, bold: true },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: {
                top: { style: 'thin' }, left: { style: 'thin' },
                bottom: { style: 'thin' }, right: { style: 'thin' }
            }
        };
        
        [1, 2].forEach(rowNum => {
            worksheet.getRow(rowNum).eachCell(cell => Object.assign(cell, headerStyle));
        });
        worksheet.getRow(1).height = 25;
        
        worksheet.columns.forEach(col => col.width = 12);
        
        const totalCols = basicInfoFieldsCount + (fileCount + 2) * 6 + (matchConfig ? matchConfig.fieldsToAdd.length : 0);
        for (let row = 3; row <= worksheet.rowCount; row++) {
            for (let col = basicInfoFieldsCount + 1; col <= totalCols; col++) {
                const cell = worksheet.getCell(row, col);
                const header = worksheet.getCell(2, col).value;
                const isComparison = col >= basicInfoFieldsCount + 1 + fileCount * 6 && 
                                   col < basicInfoFieldsCount + 1 + (fileCount + 2) * 6;
                
                if (header?.includes('率')) {
                    cell.numFmt = '0.00%';
                } else if (isComparison && !header?.includes('率')) {
                    cell.numFmt = '0.00%';
                } else if (typeof cell.value === 'number') {
                    cell.numFmt = '#,##0.00';
                }
            }
        }
    }
};