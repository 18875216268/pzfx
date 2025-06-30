/**
 * 计算字段模块 - 重构：实现Excel风格的数组计算
 */
const CalcEngine = {
    /**
     * 支持的函数
     */
    functions: {
        sum: arr => arr.reduce((a, b) => a + b, 0),
        avg: arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0,
        count: arr => arr.length,
        max: arr => arr.length ? Math.max(...arr) : 0,
        min: arr => arr.length ? Math.min(...arr) : 0
    },

    /**
     * 解析表达式
     * @param {string} expression - 表达式字符串
     * @returns {Object} 解析后的表达式对象
     */
    parseExpression(expression) {
        // 提取字段名和表达式
        const parts = expression.split('=');
        if (parts.length !== 2) {
            throw new Error('表达式格式错误，应为：字段名=表达式');
        }

        const fieldName = parts[0].trim();
        const formula = parts[1].trim();

        if (!fieldName) {
            throw new Error('字段名不能为空');
        }

        // 提取表达式中使用的字段
        const usedFields = this.extractFields(formula);

        return {
            name: fieldName,
            formula: formula,
            usedFields: usedFields
        };
    },

    /**
     * 从表达式中提取字段名 - 直接提取显示格式的字段名
     * @param {string} formula - 表达式
     * @returns {Array} 字段名数组
     */
    extractFields(formula) {
        const fields = new Set();
        
        // 匹配函数参数中的字段
        const funcPattern = /\b(sum|avg|count|max|min)\s*\(\s*([^)]+)\s*\)/gi;
        let match;
        while ((match = funcPattern.exec(formula)) !== null) {
            const fieldName = match[2].trim();
            if (fieldName && !this.isNumber(fieldName)) {
                fields.add(fieldName);
            }
        }

        // 匹配独立的字段名（不在函数内的）
        let cleanFormula = formula.replace(funcPattern, '');
        
        // 移除运算符和括号，替换为空格
        cleanFormula = cleanFormula.replace(/[+\-*/()]/g, ' ');
        
        // 按空格分割，获取所有可能的字段名
        const tokens = cleanFormula.split(/\s+/).filter(token => token.length > 0);
        
        tokens.forEach(token => {
            // 排除纯数字、运算符和函数名
            if (!this.isNumber(token) && !this.isOperator(token) && !this.isFunction(token)) {
                fields.add(token);
            }
        });

        return Array.from(fields);
    },

    /**
     * 检查是否为数字
     */
    isNumber(str) {
        return !isNaN(parseFloat(str)) && isFinite(str);
    },

    /**
     * 检查是否为运算符关键字
     */
    isOperator(str) {
        return ['and', 'or', 'not'].includes(str.toLowerCase());
    },

    /**
     * 检查是否为函数名
     */
    isFunction(str) {
        return Object.keys(this.functions).includes(str.toLowerCase());
    },

    /**
     * 构建数组计算表达式
     * @param {string} formula - 原始表达式
     * @param {Object} fieldArrays - 字段数组映射 {fieldName: [values]}
     * @returns {string} 可执行的JavaScript代码
     */
    buildArrayExpression(formula, fieldArrays) {
        let processedFormula = formula;
        
        // 第一步：处理函数调用，替换为标量值
        Object.keys(this.functions).forEach(funcName => {
            const pattern = new RegExp(`\\b${funcName}\\s*\\(\\s*([^)]+)\\s*\\)`, 'gi');
            processedFormula = processedFormula.replace(pattern, (match, fieldExpr) => {
                const field = fieldExpr.trim();
                const values = fieldArrays[field];
                if (!Array.isArray(values)) {
                    throw new Error(`字段 ${field} 不存在或不是数组`);
                }
                return this.functions[funcName](values);
            });
        });

        // 第二步：将字段名替换为数组引用
        Object.keys(fieldArrays).forEach(field => {
            const escapedField = field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const pattern = new RegExp(`(?<![\\w.])${escapedField}(?![\\w.])`, 'g');
            processedFormula = processedFormula.replace(pattern, `fieldArrays["${field}"][i]`);
        });

        // 构建数组计算的JavaScript代码
        const arrayLength = Math.max(...Object.values(fieldArrays).map(arr => arr.length));
        const code = `
            const result = [];
            for (let i = 0; i < ${arrayLength}; i++) {
                result.push(${processedFormula});
            }
            return result;
        `;

        return code;
    },

    /**
     * 计算表达式 - 返回数组结果
     * @param {string} formula - 表达式
     * @param {Object} fieldArrays - 字段数组数据 {fieldDisplayName: [values]}
     * @returns {Array} 计算结果数组
     */
    calculateArray(formula, fieldArrays) {
        try {
            const code = this.buildArrayExpression(formula, fieldArrays);
            const func = new Function('fieldArrays', code);
            return func(fieldArrays);
        } catch (error) {
            throw new Error(`计算表达式失败: ${error.message}`);
        }
    },

    /**
     * 对每个分组计算新字段
     * @param {Object} processedData - 处理后的数据
     * @param {Array} calcFields - 计算字段配置
     */
    calculateForAllGroups(processedData, calcFields) {
        if (!calcFields || !calcFields.length) return;

        // 对每个文件的数据进行处理
        Object.keys(processedData).forEach(fileName => {
            const fileData = processedData[fileName];
            
            // 获取所有分组（保持顺序）
            const groups = Object.keys(fileData);
            
            // 按顺序处理每个计算字段
            calcFields.forEach(calcField => {
                const { name, formula, usedFields } = calcField;
                
                // 收集所有分组的数据，构建数组
                const fieldArrays = {};
                
                // 为每个使用的字段创建有序数组
                usedFields.forEach(fieldDisplayName => {
                    const fieldData = [];
                    
                    // 按分组顺序收集数据
                    groups.forEach(groupKey => {
                        const group = fileData[groupKey];
                        let foundValue = null;
                        
                        // 查找聚合字段
                        if (group.data[fieldDisplayName] !== undefined) {
                            foundValue = group.data[fieldDisplayName];
                        } else {
                            // 查找计算字段
                            const calcFieldKey = `${fieldDisplayName}.计算`;
                            if (group.data[calcFieldKey] !== undefined) {
                                foundValue = group.data[calcFieldKey];
                            }
                        }
                        
                        // 确保数值类型
                        fieldData.push(foundValue !== null && foundValue !== undefined ? Number(foundValue) : 0);
                    });
                    
                    fieldArrays[fieldDisplayName] = fieldData;
                });

                // 计算得到结果数组
                try {
                    const resultArray = this.calculateArray(formula, fieldArrays);
                    
                    // 将结果分配给每个分组
                    groups.forEach((groupKey, index) => {
                        const group = fileData[groupKey];
                        const calcKey = `${name}.计算`;
                        group.data[calcKey] = resultArray[index] || 0;
                    });
                    
                } catch (error) {
                    console.error(`计算字段 ${name} 失败:`, error);
                    // 出错时给所有分组赋值0
                    groups.forEach(groupKey => {
                        const group = fileData[groupKey];
                        group.data[`${name}.计算`] = 0;
                    });
                }
            });
        });
    },

    /**
     * 验证计算字段配置
     * @param {Object} calcField - 计算字段配置
     * @param {Array} availableFields - 可用字段列表
     * @returns {Object} 验证结果
     */
    validateCalcField(calcField, availableFields) {
        const result = {
            isValid: true,
            errors: []
        };

        // 检查字段名
        if (!calcField.name) {
            result.isValid = false;
            result.errors.push('字段名不能为空');
        }

        // 检查表达式
        if (!calcField.formula) {
            result.isValid = false;
            result.errors.push('表达式不能为空');
        }

        // 检查使用的字段是否存在
        const availableDisplayNames = [
            ...state.aggregateFields,
            ...state.calcFields.filter(f => f.id !== calcField.id).map(f => f.name)
        ];
        
        const missingFields = calcField.usedFields.filter(field => 
            !availableDisplayNames.includes(field)
        );
        
        if (missingFields.length > 0) {
            result.isValid = false;
            result.errors.push(`以下字段不存在: ${missingFields.join(', ')}`);
        }

        return result;
    }
};