
class Spreadsheet {
    constructor() {
        this.sheets = [];
        this.activeSheet = null;
        this.selectedCells = [];
        this.currentEditingCell = null;
        this.isEditing = false;
        this.INITIAL_ROWS = 100; // 100 linhas iniciais
        this.clipboard = [];
        this.undoStack = [];
        this.redoStack = [];
        this.isSelecting = false;
        this.selectionStart = null;
        this.resizeCol = null;
        this.resizeStartX = null;
        this.resizeStartWidth = null;
        
        // Create first sheet
        this.createSheet('Planilha1');
        
        // Bind events
        this.bindEvents();
        this.setupToolbar();
        this.setupExport();
        this.setupContextMenu();
        
        // Atualizar contador de linhas
        this.updateRowCount();
    }
    
    createSheet(name) {
        const sheet = {
            id: Date.now(),
            name: name,
            rows: [],
            cols: 26,
            data: {},
            colWidths: Array(26).fill(100)
        };
        
        // Criar 100 linhas iniciais
        for (let i = 0; i < this.INITIAL_ROWS; i++) {
            this.ensureRow(sheet, i);
        }
        
        this.sheets.push(sheet);
        this.activeSheet = sheet;
        this.render();
        this.updateSheetTabs();
        this.updateRowCount();
    }
    
    ensureRow(sheet, rowIndex) {
        if (!sheet.data[rowIndex]) {
            sheet.data[rowIndex] = {};
        }
        while (sheet.rows.length <= rowIndex) {
            sheet.rows.push(null);
        }
    }
    
    ensureCol(sheet, colIndex) {
        if (colIndex >= sheet.cols) {
            sheet.cols = colIndex + 1;
        }
    }
    
    getColumnName(colIndex) {
        let name = '';
        let n = colIndex;
        while (n >= 0) {
            name = String.fromCharCode(65 + (n % 26)) + name;
            n = Math.floor(n / 26) - 1;
        }
        return name;
    }
    
    getColumnIndex(colName) {
        let result = 0;
        for (let i = 0; i < colName.length; i++) {
            result = result * 26 + (colName.charCodeAt(i) - 64);
        }
        return result - 1;
    }
    
    getCellRef(row, col) {
        return this.getColumnName(col) + (row + 1);
    }
    
    parseCellRef(ref) {
        const match = ref.match(/([A-Z]+)(\d+)/);
        if (match) {
            return {
                row: parseInt(match[2]) - 1,
                col: this.getColumnIndex(match[1])
            };
        }
        return null;
    }
    
    getCellValue(sheet, row, col) {
        if (!sheet.data[row] || !sheet.data[row][col]) {
            return '';
        }
        return sheet.data[row][col].value || '';
    }
    
    getNumericValue(sheet, row, col) {
        const value = this.getCellValue(sheet, row, col);
        if (value === '') return 0;
        
        if (typeof value === 'string' && value.includes('R$')) {
            return parseFloat(value.replace(/[^\d,-]/g, '').replace(',', '.'));
        }
        
        return parseFloat(value) || 0;
    }
    
    setCellValue(sheet, row, col, value) {
        if (!sheet.data[row]) {
            sheet.data[row] = {};
        }
        if (!sheet.data[row][col]) {
            sheet.data[row][col] = {
                value: '',
                type: 'text',
                formula: null,
                hyperlink: null,
                style: {
                    fontFamily: 'Arial',
                    fontSize: '12px',
                    bold: false,
                    italic: false,
                    backgroundColor: '#ffffff',
                    color: '#000000'
                }
            };
        }
        
        const cell = sheet.data[row][col];
        
        // Detectar fórmula
        if (typeof value === 'string' && value.startsWith('=')) {
            cell.formula = value;
            cell.value = this.processFormula(sheet, row, col, value);
        } else {
            cell.formula = null;
            
            // Process value based on type
            if (cell.type === 'number') {
                value = parseFloat(value.replace(',', '.'));
                if (isNaN(value)) {
                    value = 0;
                }
                cell.value = value.toFixed(2);
            } else if (cell.type === 'currency') {
                const numValue = parseFloat(value.replace(/[^\d,-]/g, '').replace(',', '.'));
                if (!isNaN(numValue)) {
                    cell.value = numValue.toLocaleString('pt-BR', {
                        style: 'currency',
                        currency: 'BRL'
                    });
                } else {
                    cell.value = '';
                }
            } else {
                cell.value = value;
            }
        }
    }
    
    processFormula(sheet, row, col, formula) {
        try {
            let expression = formula.substring(1).trim();
            
            const functionMatch = expression.match(/^([A-Z]+\.?[A-Z]*)\s*\((.*)\)$/i);
            if (functionMatch) {
                return this.processFunction(sheet, row, col, functionMatch[1].toUpperCase(), functionMatch[2]);
            }
            
            return this.processArithmetic(sheet, row, col, expression);
            
        } catch (error) {
            console.error('Erro ao processar fórmula:', error);
            return '#ERRO!';
        }
    }
    
    processFunction(sheet, row, col, funcName, args) {
        const argList = this.parseArguments(args);
        
        switch(funcName) {
            case 'SOMA':
                return this.calculateSum(sheet, argList);
            case 'MEDIA':
                return this.calculateAverage(sheet, argList);
            case 'MAX':
                return this.calculateMax(sheet, argList);
            case 'MIN':
                return this.calculateMin(sheet, argList);
            case 'CONT.SE':
                return this.calculateCountIf(sheet, argList);
            case 'SOMA.SE':
                return this.calculateSumIf(sheet, argList);
            case 'SE':
                return this.calculateIf(sheet, row, col, argList);
            case 'CONCATENAR':
                return this.calculateConcatenate(sheet, argList);
            case 'UPPER':
                return this.calculateUpper(sheet, argList);
            case 'LOWER':
                return this.calculateLower(sheet, argList);
            case 'TRIM':
                return this.calculateTrim(sheet, argList);
            case 'E':
                return this.calculateAnd(sheet, argList);
            case 'OU':
                return this.calculateOr(sheet, argList);
            case 'NÃO':
                return this.calculateNot(sheet, argList);
            case 'ARRED':
                return this.calculateRound(sheet, argList);
            case 'ABS':
                return this.calculateAbs(sheet, argList);
            case 'RAIZ':
                return this.calculateSqrt(sheet, argList);
            case 'POTENCIA':
                return this.calculatePower(sheet, argList);
            case 'DESVPAD':
                return this.calculateStdDev(sheet, argList);
            case 'VAR':
                return this.calculateVar(sheet, argList);
            default:
                return '#NOME?';
        }
    }
    
    parseArguments(args) {
        const argList = [];
        let current = '';
        let depth = 0;
        
        for (let i = 0; i < args.length; i++) {
            const char = args[i];
            
            if (char === '(') depth++;
            if (char === ')') depth--;
            
            if ((char === ';' || char === ',') && depth === 0) {
                argList.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        if (current.trim()) {
            argList.push(current.trim());
        }
        
        return argList;
    }
    
    parseRange(sheet, rangeStr) {
        const parts = rangeStr.split(':');
        if (parts.length === 2) {
            const start = this.parseCellRef(parts[0].trim());
            const end = this.parseCellRef(parts[1].trim());
            
            if (start && end) {
                const values = [];
                const minRow = Math.min(start.row, end.row);
                const maxRow = Math.max(start.row, end.row);
                const minCol = Math.min(start.col, end.col);
                const maxCol = Math.max(start.col, end.col);
                
                for (let r = minRow; r <= maxRow; r++) {
                    for (let c = minCol; c <= maxCol; c++) {
                        values.push({
                            row: r,
                            col: c,
                            value: this.getNumericValue(sheet, r, c)
                        });
                    }
                }
                
                return values;
            }
        } else if (parts.length === 1) {
            const cellRef = this.parseCellRef(parts[0].trim());
            if (cellRef) {
                return [{
                    row: cellRef.row,
                    col: cellRef.col,
                    value: this.getNumericValue(sheet, cellRef.row, cellRef.col)
                }];
            }
        }
        
        return [];
    }
    
    getArgumentValue(sheet, arg) {
        if (!isNaN(parseFloat(arg))) {
            return parseFloat(arg);
        }
        
        const cellRef = this.parseCellRef(arg);
        if (cellRef) {
            return this.getNumericValue(sheet, cellRef.row, cellRef.col);
        }
        
        if (arg.startsWith('"') && arg.endsWith('"')) {
            return arg.substring(1, arg.length - 1);
        }
        
        if (arg.includes(':')) {
            const range = this.parseRange(sheet, arg);
            return range[0] ? range[0].value : 0;
        }
        
        return 0;
    }
    
    calculateSum(sheet, args) {
        let total = 0;
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    total += cell.value;
                });
            } else {
                total += this.getArgumentValue(sheet, arg);
            }
        });
        return total;
    }
    
    calculateAverage(sheet, args) {
        let total = 0;
        let count = 0;
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    if (cell.value !== 0) {
                        total += cell.value;
                        count++;
                    }
                });
            } else {
                const value = this.getArgumentValue(sheet, arg);
                if (value !== 0) {
                    total += value;
                    count++;
                }
            }
        });
        return count > 0 ? total / count : 0;
    }
    
    calculateMax(sheet, args) {
        let max = -Infinity;
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    if (cell.value > max) max = cell.value;
                });
            } else {
                const value = this.getArgumentValue(sheet, arg);
                if (value > max) max = value;
            }
        });
        return max === -Infinity ? 0 : max;
    }
    
    calculateMin(sheet, args) {
        let min = Infinity;
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    if (cell.value < min) min = cell.value;
                });
            } else {
                const value = this.getArgumentValue(sheet, arg);
                if (value < min) min = value;
            }
        });
        return min === Infinity ? 0 : min;
    }
    
    calculateCountIf(sheet, args) {
        if (args.length < 2) return 0;
        
        const range = this.parseRange(sheet, args[0]);
        const criteria = args[1].trim();
        
        let count = 0;
        range.forEach(cell => {
            const value = cell.value;
            
            if (criteria.startsWith('>')) {
                const num = parseFloat(criteria.substring(1));
                if (value > num) count++;
            } else if (criteria.startsWith('<')) {
                const num = parseFloat(criteria.substring(1));
                if (value < num) count++;
            } else if (criteria.startsWith('>=')) {
                const num = parseFloat(criteria.substring(2));
                if (value >= num) count++;
            } else if (criteria.startsWith('<=')) {
                const num = parseFloat(criteria.substring(2));
                if (value <= num) count++;
            } else if (criteria.startsWith('=')) {
                const num = parseFloat(criteria.substring(1));
                if (value === num) count++;
            } else {
                if (value === parseFloat(criteria)) count++;
            }
        });
        
        return count;
    }
    
    calculateSumIf(sheet, args) {
        if (args.length < 2) return 0;
        
        const range = this.parseRange(sheet, args[0]);
        const criteria = args[1].trim();
        
        let sum = 0;
        range.forEach(cell => {
            const value = cell.value;
            
            if (criteria.startsWith('>')) {
                const num = parseFloat(criteria.substring(1));
                if (value > num) sum += value;
            } else if (criteria.startsWith('<')) {
                const num = parseFloat(criteria.substring(1));
                if (value < num) sum += value;
            } else if (criteria.startsWith('>=')) {
                const num = parseFloat(criteria.substring(2));
                if (value >= num) sum += value;
            } else if (criteria.startsWith('<=')) {
                const num = parseFloat(criteria.substring(2));
                if (value <= num) sum += value;
            } else if (criteria.startsWith('=')) {
                const num = parseFloat(criteria.substring(1));
                if (value === num) sum += value;
            } else {
                if (value === parseFloat(criteria)) sum += value;
            }
        });
        
        return sum;
    }
    
    calculateIf(sheet, row, col, args) {
        if (args.length < 3) return '';
        
        const condition = args[0].trim();
        const valueIfTrue = args[1].trim();
        const valueIfFalse = args[2].trim();
        
        let result = false;
        
        if (condition.includes('>=')) {
            const parts = condition.split('>=');
            const left = this.getArgumentValue(sheet, parts[0].trim());
            const right = parseFloat(parts[1].trim());
            result = left >= right;
        } else if (condition.includes('<=')) {
            const parts = condition.split('<=');
            const left = this.getArgumentValue(sheet, parts[0].trim());
            const right = parseFloat(parts[1].trim());
            result = left <= right;
        } else if (condition.includes('>')) {
            const parts = condition.split('>');
            const left = this.getArgumentValue(sheet, parts[0].trim());
            const right = parseFloat(parts[1].trim());
            result = left > right;
        } else if (condition.includes('<')) {
            const parts = condition.split('<');
            const left = this.getArgumentValue(sheet, parts[0].trim());
            const right = parseFloat(parts[1].trim());
            result = left < right;
        } else if (condition.includes('=')) {
            const parts = condition.split('=');
            const left = this.getArgumentValue(sheet, parts[0].trim());
            const right = this.getArgumentValue(sheet, parts[1].trim());
            result = left === right;
        }
        
        if (result) {
            return this.getArgumentValue(sheet, valueIfTrue);
        } else {
            return this.getArgumentValue(sheet, valueIfFalse);
        }
    }
    
    calculateConcatenate(sheet, args) {
        let result = '';
        args.forEach(arg => {
            const value = this.getArgumentValue(sheet, arg);
            if (typeof value === 'string') {
                result += value;
            } else if (typeof value === 'number') {
                result += String(value);
            } else if (arg.startsWith('"')) {
                result += arg.substring(1, arg.length - 1);
            }
        });
        return result;
    }
    
    calculateUpper(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return typeof value === 'string' ? value.toUpperCase() : String(value).toUpperCase();
    }
    
    calculateLower(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return typeof value === 'string' ? value.toLowerCase() : String(value).toLowerCase();
    }
    
    calculateTrim(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return typeof value === 'string' ? value.trim() : String(value).trim();
    }
    
    calculateAnd(sheet, args) {
        let result = true;
        args.forEach(arg => {
            const value = this.getArgumentValue(sheet, arg);
            if (value === 0 || value === false || value === '') {
                result = false;
            }
        });
        return result ? 'VERDADEIRO' : 'FALSO';
    }
    
    calculateOr(sheet, args) {
        let result = false;
        args.forEach(arg => {
            const value = this.getArgumentValue(sheet, arg);
            if (value !== 0 && value !== false && value !== '') {
                result = true;
            }
        });
        return result ? 'VERDADEIRO' : 'FALSO';
    }
    
    calculateNot(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return value === 0 || value === false || value === '' ? 'VERDADEIRO' : 'FALSO';
    }
    
    calculateRound(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        const decimals = parseInt(this.getArgumentValue(sheet, args[1] || '0'));
        return Math.round(value * Math.pow(10, decimals)) / Math.pow(10, decimals);
    }
    
    calculateAbs(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return Math.abs(value);
    }
    
    calculateSqrt(sheet, args) {
        const value = this.getArgumentValue(sheet, args[0]);
        return Math.sqrt(value);
    }
    
    calculatePower(sheet, args) {
        const base = this.getArgumentValue(sheet, args[0]);
        const exponent = this.getArgumentValue(sheet, args[1]);
        return Math.pow(base, exponent);
    }
    
    calculateStdDev(sheet, args) {
        const values = [];
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    if (cell.value !== 0) val