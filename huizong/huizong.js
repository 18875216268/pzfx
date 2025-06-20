// 分组后计算、汇总对应字段并填充模块

// 扩展ExcelProcessor类的汇总功能
Object.assign(ExcelProcessor, {
    // 创建汇总工作簿
    createSummaryWorkbook(processedData) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('汇总');
        
        // 准备数据
        const allDrugInfo = this.getAllDrugInfo(processedData);
        const drugIds = Object.keys(allDrugInfo).sort();
        const fileInfo = this.prepareFileInfo(processedData);
        
        // 创建表头
        this.createHeaders(worksheet, fileInfo);
        
        // 填充数据
        drugIds.forEach(drugId => {
            const rowData = this.buildDataRow(drugId, allDrugInfo[drugId], fileInfo);
            worksheet.addRow(rowData);
        });
        
        // 应用样式
        this.applyStyles(worksheet, fileInfo.files.length);
        
        return workbook;
    },
    
    // 创建表头
    createHeaders(worksheet, fileInfo) {
        // 第一行
        const row1 = ['药品基础信息', '', '', '', ''];
        fileInfo.files.forEach(f => {
            row1.push(f.fileName.replace(/\.[^/.]+$/, ''), ...Array(5).fill(''));
        });
        row1.push('同比', ...Array(5).fill(''), '环比', ...Array(5).fill(''));
        worksheet.addRow(row1);
        
        // 第二行
        const row2 = ['品种负责人', '药品ID', '商品名称', '商品规格', '生产厂家'];
        const repeatCount = fileInfo.files.length + 2;
        for (let i = 0; i < repeatCount; i++) {
            row2.push(...CONFIG.dataFields);
        }
        worksheet.addRow(row2);
    },
    
    // 构建数据行
    buildDataRow(drugId, drugInfo, fileInfo) {
        const row = [
            drugInfo['品种负责人'] || '',
            drugId,
            drugInfo['商品名称'] || '',
            drugInfo['商品规格'] || '',
            drugInfo['生产厂家'] || ''
        ];
        
        const periodData = {};
        fileInfo.files.forEach(file => {
            const data = file.data[drugId] || {};
            periodData[file.type] = data;
            row.push(...this.formatDrugData(data));
        });
        
        // 同比和环比
        row.push(...this.calculateComparison(periodData['当期'], periodData['同期']));
        row.push(...this.calculateComparison(periodData['当期'], periodData['上期']));
        
        return row;
    },
    
    // 格式化药品数据
    formatDrugData(data) {
        const amount = data['含税出库金额'] || 0;
        const p4Profit = data['P4毛利额'] || 0;
        const marginProfit = data['应收边际利润额(不含税)'] || 0;
        
        return [
            amount,
            p4Profit,
            amount ? p4Profit / amount : 0,
            marginProfit,
            amount ? marginProfit / amount : 0,
            data['客户数'] || 0
        ];
    },
    
    // 计算比较数据
    calculateComparison(current = {}, previous = {}) {
        const curr = this.formatDrugData(current);
        const prev = this.formatDrugData(previous);
        
        return [
            prev[0] ? (curr[0] - prev[0]) / prev[0] : 0,  // 金额增长率
            prev[1] ? (curr[1] - prev[1]) / prev[1] : 0,  // P4毛利额增长率
            curr[2] - prev[2],  // P4毛利率差异
            prev[3] ? (curr[3] - prev[3]) / prev[3] : 0,  // 边际利润额增长率
            curr[4] - prev[4],  // 边际利润率差异
            prev[5] ? (curr[5] - prev[5]) / prev[5] : 0   // 客户数增长率
        ];
    },
    
    // 应用样式
    applyStyles(worksheet, fileCount) {
        // 合并单元格
        worksheet.mergeCells('A1:E1');
        let startCol = 6;
        for (let i = 0; i < fileCount + 2; i++) {
            worksheet.mergeCells(1, startCol, 1, startCol + 5);
            startCol += 6;
        }
        
        // 样式配置
        const headerStyle = {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BBB59' } },
            font: { color: { argb: 'FFFFFFFF' }, bold: true },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };
        
        // 应用到前两行
        [1, 2].forEach(rowNum => {
            worksheet.getRow(rowNum).eachCell(cell => Object.assign(cell, headerStyle));
        });
        worksheet.getRow(1).height = 25;
        
        // 设置列宽
        worksheet.columns.forEach(col => col.width = 12);
        
        // 设置数字格式
        const totalCols = 5 + (fileCount + 2) * 6;
        for (let row = 3; row <= worksheet.rowCount; row++) {
            for (let col = 6; col <= totalCols; col++) {
                const cell = worksheet.getCell(row, col);
                const header = worksheet.getCell(2, col).value;
                const isComparison = col >= 6 + fileCount * 6;
                
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
});