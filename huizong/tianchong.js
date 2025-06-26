// 填充模块
const FillModule = {
    createSummaryWorkbook(processedData, groupFields, aggregateFields, calcFields, basicInfoFields, allFields) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('汇总');
        
        // 准备数据
        const groupKeys = AggregateModule.getAllGroupKeys(processedData);
        const fileInfo = AggregateModule.prepareFileInfo(processedData);
        const periodData = AggregateModule.aggregateByPeriod(fileInfo);
        const basicInfo = AggregateModule.extractAllBasicInfo(periodData, basicInfoFields);
        
        // 为每个期间的数据添加计算字段
        Object.keys(periodData).forEach(period => {
            Object.keys(periodData[period]).forEach(groupKey => {
                const data = periodData[period][groupKey].data;
                periodData[period][groupKey].data = CalculateModule.calculateFields(data, calcFields);
            });
        });
        
        // 创建表头
        this.createHeaders(worksheet, fileInfo, basicInfoFields, aggregateFields, calcFields);
        
        // 填充数据
        groupKeys.forEach(groupKey => {
            const row = this.buildDataRow(
                groupKey, basicInfo[groupKey], periodData, 
                basicInfoFields, aggregateFields, calcFields
            );
            worksheet.addRow(row);
        });
        
        // 应用样式
        this.applyStyles(worksheet, fileInfo.length, basicInfoFields.length, 
                         Object.keys(aggregateFields).length + Object.keys(calcFields).length);
        
        return workbook;
    },
    
    createHeaders(worksheet, fileInfo, basicInfoFields, aggregateFields, calcFields) {
        // 第一行：大标题
        const row1 = new Array(basicInfoFields.length).fill('');
        row1[0] = '基础信息';
        
        const dataFields = [...Object.keys(aggregateFields), ...Object.keys(calcFields)];
        const dataFieldCount = dataFields.length;
        
        // 各期数据标题
        fileInfo.forEach(f => {
            const fileName = f.fileName.replace(/\.[^/.]+$/, '');
            row1.push(fileName, ...new Array(dataFieldCount - 1).fill(''));
        });
        
        // 比较标题
        row1.push('同比', ...new Array(dataFieldCount - 1).fill(''), 
                  '环比', ...new Array(dataFieldCount - 1).fill(''));
        
        worksheet.addRow(row1);
        
        // 第二行：字段名
        const row2 = [...basicInfoFields];
        
        // 重复5次（3个期间 + 2个比较）
        for (let i = 0; i < 5; i++) {
            row2.push(...dataFields);
        }
        
        worksheet.addRow(row2);
    },
    
    buildDataRow(groupKey, basicInfo, periodData, basicInfoFields, aggregateFields, calcFields) {
        const row = [];
        
        // 基础信息
        basicInfoFields.forEach(field => {
            row.push(basicInfo?.[field] || '');
        });
        
        const periods = ['同期', '上期', '当期'];
        const periodValues = {};
        
        // 各期数据
        periods.forEach(period => {
            const data = periodData[period]?.[groupKey]?.data || {};
            periodValues[period] = data;
            row.push(...CalculateModule.formatRowData(data, aggregateFields, calcFields));
        });
        
        // 同比和环比
        row.push(...CalculateModule.calculateComparison(
            periodValues['当期'], periodValues['同期'], aggregateFields, calcFields
        ));
        row.push(...CalculateModule.calculateComparison(
            periodValues['当期'], periodValues['上期'], aggregateFields, calcFields
        ));
        
        return row;
    },
    
    applyStyles(worksheet, fileCount, basicInfoFieldsCount, dataFieldCount) {
        // 合并单元格
        worksheet.mergeCells(1, 1, 1, basicInfoFieldsCount);
        
        let startCol = basicInfoFieldsCount + 1;
        for (let i = 0; i < fileCount + 2; i++) {
            worksheet.mergeCells(1, startCol, 1, startCol + dataFieldCount - 1);
            startCol += dataFieldCount;
        }
        
        // 表头样式
        const headerStyle = {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4CAF50' } },
            font: { color: { argb: 'FFFFFFFF' }, bold: true },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };
        
        // 应用表头样式
        [1, 2].forEach(rowNum => {
            worksheet.getRow(rowNum).eachCell(cell => Object.assign(cell, headerStyle));
        });
        worksheet.getRow(1).height = 25;
        
        // 设置列宽
        worksheet.columns.forEach(col => col.width = 12);
        
        // 数字格式处理
        const totalCols = basicInfoFieldsCount + (fileCount + 2) * dataFieldCount;
        const comparisonStartCol = basicInfoFieldsCount + 1 + fileCount * dataFieldCount;
        
        for (let row = 3; row <= worksheet.rowCount; row++) {
            for (let col = basicInfoFieldsCount + 1; col <= totalCols; col++) {
                const cell = worksheet.getCell(row, col);
                const header = worksheet.getCell(2, col).value;
                const isComparison = col >= comparisonStartCol;
                
                if (typeof cell.value === 'number') {
                    if (header?.includes('率')) {
                        cell.numFmt = '0.00%';
                    } else if (isComparison && !header?.includes('数')) {
                        cell.numFmt = '0.00%';
                    } else {
                        cell.numFmt = '#,##0.00';
                    }
                }
            }
        }
    }
};