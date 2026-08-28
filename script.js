lclass Spreadsheet {
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
        this.setupMobile();
        this.setupTouchEvents();
        this.setupResizeObserver();
        
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
                    if (cell.value !== 0) values.push(cell.value);
                });
            } else {
                const value = this.getArgumentValue(sheet, arg);
                if (value !== 0) values.push(value);
            }
        });
        
        if (values.length < 2) return 0;
        
        const mean = values.reduce((a, b) => a + b) / values.length;
        const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1);
        
        return Math.sqrt(variance);
    }
    
    calculateVar(sheet, args) {
        const values = [];
        args.forEach(arg => {
            if (arg.includes(':')) {
                const range = this.parseRange(sheet, arg);
                range.forEach(cell => {
                    if (cell.value !== 0) values.push(cell.value);
                });
            } else {
                const value = this.getArgumentValue(sheet, arg);
                if (value !== 0) values.push(value);
            }
        });
        
        if (values.length < 2) return 0;
        
        const mean = values.reduce((a, b) => a + b) / values.length;
        return values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length - 1);
    }
    
    processArithmetic(sheet, row, col, expression) {
        // Substituir referências de células por seus valores
        expression = expression.replace(/([A-Z]+)(\d+)/g, (match, colName, rowNum) => {
            const cellRow = parseInt(rowNum) - 1;
            const cellCol = this.getColumnIndex(colName);
            return this.getNumericValue(sheet, cellRow, cellCol);
        });
        
        // Substituir vírgulas por pontos para cálculos
        expression = expression.replace(/,/g, '.');
        
        // Avaliar expressão com eval seguro
        try {
            return Function('"use strict"; return (' + expression + ')')();
        } catch (error) {
            return '#ERRO!';
        }
    }
    
    // Métodos de renderização e interação
    render() {
        const spreadsheet = document.getElementById('spreadsheet');
        spreadsheet.innerHTML = '';
        
        const sheet = this.activeSheet;
        
        // Criar grid
        spreadsheet.style.gridTemplateColumns = `100px ${sheet.colWidths.map(w => `${w}px`).join(' ')}`;
        
        // Cabeçalho de colunas
        const cornerCell = document.createElement('div');
        cornerCell.className = 'cell header corner';
        cornerCell.textContent = '';
        spreadsheet.appendChild(cornerCell);
        
        for (let col = 0; col < sheet.cols; col++) {
            const headerCell = document.createElement('div');
            headerCell.className = 'cell header';
            headerCell.textContent = this.getColumnName(col);
            headerCell.dataset.col = col;
            headerCell.addEventListener('click', () => this.selectColumn(col));
            headerCell.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                this.showContextMenu(e.clientX, e.clientY, { type: 'column', col: col });
            });
            spreadsheet.appendChild(headerCell);
        }
        
        // Linhas
        for (let row = 0; row < sheet.rows.length; row++) {
            // Cabeçalho de linha
            const rowHeader = document.createElement('div');
            rowHeader.className = 'cell header row-header';
            rowHeader.textContent = row + 1;
            rowHeader.dataset.row = row;
            rowHeader.addEventListener('click', () => this.selectRow(row));
            rowHeader.addEventListener('contextmenu', (e) => {
                e.preventDefault();
                this.showContextMenu(e.clientX, e.clientY, { type: 'row', row: row });
            });
            spreadsheet.appendChild(rowHeader);
            
            // Células
            for (let col = 0; col < sheet.cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.dataset.row = row;
                cell.dataset.col = col;
                
                const cellData = sheet.data[row] && sheet.data[row][col];
                if (cellData) {
                    cell.textContent = cellData.value || '';
                    cell.style.fontFamily = cellData.style.fontFamily;
                    cell.style.fontSize = cellData.style.fontSize;
                    cell.style.fontWeight = cellData.style.bold ? 'bold' : 'normal';
                    cell.style.fontStyle = cellData.style.italic ? 'italic' : 'normal';
                    cell.style.backgroundColor = cellData.style.backgroundColor;
                    cell.style.color = cellData.style.color;
                    
                    if (cellData.type === 'currency') {
                        cell.classList.add('currency');
                    } else if (cellData.type === 'number') {
                        cell.classList.add('number');
                    }
                }
                
                cell.addEventListener('click', (e) => this.handleCellClick(e, row, col));
                cell.addEventListener('dblclick', (e) => this.handleCellDoubleClick(e, row, col));
                cell.addEventListener('contextmenu', (e) => {
                    e.preventDefault();
                    this.showContextMenu(e.clientX, e.clientY, { type: 'cell', row: row, col: col });
                });
                cell.addEventListener('touchstart', (e) => {
                    if (window.innerWidth <= 768) {
                        this.handleCellTouch(row, col);
                    }
                }, { passive: true });
                
                spreadsheet.appendChild(cell);
            }
        }
    }
    
    handleCellClick(e, row, col) {
        if (this.isEditing) {
            this.finishEditing();
        }
        
        // Limpar seleção anterior
        document.querySelectorAll('.cell.selected').forEach(cell => {
            cell.classList.remove('selected');
        });
        
        // Selecionar célula atual
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (cellElement) {
            cellElement.classList.add('selected');
        }
        
        this.selectedCells = [this.getCellRef(row, col)];
        this.updateCellRef(row, col);
    }
    
    handleCellDoubleClick(e, row, col) {
        this.startEditing(row, col);
    }
    
    handleCellTouch(row, col) {
        // Melhorar seleção para touch
        if (this.isEditing) {
            this.finishEditing();
        }
        
        // Selecionar célula
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (cellElement) {
            cellElement.classList.add('selected');
        }
        
        // Mostrar opções rápidas no mobile
        if (window.innerWidth <= 768) {
            this.showQuickOptions(row, col);
        }
    }
    
    showQuickOptions(row, col) {
        // Criar menu rápido para mobile
        const existingMenu = document.querySelector('.quick-menu');
        if (existingMenu) existingMenu.remove();
        
        const quickMenu = document.createElement('div');
        quickMenu.className = 'quick-menu';
        quickMenu.style.cssText = `
            position: fixed;
            bottom: 60px;
            left: 50%;
            transform: translateX(-50%);
            background: white;
            border-radius: 8px;
            padding: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            display: flex;
            gap: 5px;
            z-index: 1000;
        `;
        
        const buttons = [
            { icon: '✏️', action: () => this.startEditing(row, col) },
            { icon: '📋', action: () => this.copyCell(row, col) },
            { icon: '📝', action: () => this.pasteCell(row, col) },
            { icon: '🗑️', action: () => this.clearCell(row, col) },
            { icon: '❌', action: () => quickMenu.remove() }
        ];
        
        buttons.forEach(btn => {
            const button = document.createElement('button');
            button.innerHTML = btn.icon;
            button.style.cssText = `
                width: 40px;
                height: 40px;
                border: 1px solid #ccc;
                border-radius: 50%;
                background: white;
                cursor: pointer;
                font-size: 18px;
            `;
            button.addEventListener('click', () => {
                btn.action();
                quickMenu.remove();
            });
            quickMenu.appendChild(button);
        });
        
        document.body.appendChild(quickMenu);
        
        // Fechar ao tocar fora
        setTimeout(() => {
            document.addEventListener('click', function handler(e) {
                if (!quickMenu.contains(e.target)) {
                    quickMenu.remove();
                    document.removeEventListener('click', handler);
                }
            });
        }, 100);
    }
    
    updateCellRef(row, col) {
        document.getElementById('cellRef').value = this.getCellRef(row, col);
        const cellData = this.activeSheet.data[row] && this.activeSheet.data[row][col];
        document.getElementById('formulaInput').value = cellData ? (cellData.formula || cellData.value || '') : '';
    }
    
    startEditing(row, col) {
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        if (!cellElement) return;
        
        this.isEditing = true;
        this.currentEditingCell = { row, col };
        
        cellElement.classList.add('editing');
        cellElement.innerHTML = '';
        
        const input = document.createElement('input');
        const cellData = this.activeSheet.data[row] && this.activeSheet.data[row][col];
        input.value = cellData ? (cellData.formula || cellData.value || '') : '';
        input.type = 'text';
        
        // No mobile, manter o teclado aberto
        if (window.innerWidth <= 768) {
            input.style.fontSize = '16px'; // Prevenir zoom no iOS
            input.style.height = '100%';
        }
        
        input.addEventListener('blur', () => {
            this.finishEditing(input.value);
        });
        
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.finishEditing(input.value);
            } else if (e.key === 'Escape') {
                this.finishEditing(null);
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.finishEditing(input.value);
                this.startEditing(row, col + 1);
            }
        });
        
        // No mobile, focar imediatamente
        cellElement.appendChild(input);
        input.focus();
        if (window.innerWidth <= 768) {
            setTimeout(() => input.select(), 100);
        } else {
            input.select();
        }
    }
    
    finishEditing(value = null) {
        if (!this.currentEditingCell) return;
        
        const { row, col } = this.currentEditingCell;
        const cellElement = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
        
        if (value !== null) {
            this.setCellValue(this.activeSheet, row, col, value);
        }
        
        this.isEditing = false;
        this.currentEditingCell = null;
        
        if (cellElement) {
            cellElement.classList.remove('editing');
        }
        
        this.render();
        this.updateCellRef(row, col);
    }
    
    setupMobile() {
        const mobileMenuBtn = document.getElementById('mobileMenuBtn');
        const toolbarContent = document.getElementById('toolbarContent');
        
        // Adicionar botão de menu mobile se não existir
        if (!mobileMenuBtn) {
            const toolbar = document.querySelector('.toolbar');
            const btn = document.createElement('button');
            btn.className = 'mobile-menu-btn';
            btn.id = 'mobileMenuBtn';
            btn.innerHTML = '☰';
            btn.title = 'Menu';
            toolbar.insertBefore(btn, toolbar.firstChild);
        }
        
        // Evento de clique para mostrar/ocultar menu
        document.getElementById('mobileMenuBtn').addEventListener('click', (e) => {
            e.stopPropagation();
            toolbarContent.classList.toggle('show');
        });
        
        // Fechar menu ao clicar fora
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.toolbar') && toolbarContent.classList.contains('show')) {
                toolbarContent.classList.remove('show');
            }
        });
        
        // Fechar menu ao selecionar opção
        document.querySelectorAll('.toolbar-content select, .toolbar-content button, .toolbar-content input').forEach(el => {
            el.addEventListener('click', () => {
                if (window.innerWidth <= 768) {
                    setTimeout(() => toolbarContent.classList.remove('show'), 200);
                }
            });
        });
    }
    
    setupTouchEvents() {
        // Suporte a gestos de pinça para zoom
        let initialScale = 1;
        let startDistance = 0;
        
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        
        spreadsheetContainer.addEventListener('touchstart', (e) => {
            if (e.touches.length === 2) {
                const touch1 = e.touches[0];
                const touch2 = e.touches[1];
                startDistance = Math.hypot(
                    touch1.clientX - touch2.clientX,
                    touch1.clientY - touch2.clientY
                );
            }
        }, { passive: true });
        
        spreadsheetContainer.addEventListener('touchmove', (e) => {
            if (e.touches.length === 2) {
                e.preventDefault();
                const touch1 = e.touches[0];
                const touch2 = e.touches[1];
                const currentDistance = Math.hypot(
                    touch1.clientX - touch2.clientX,
                    touch1.clientY - touch2.clientY
                );
                
                if (startDistance > 0) {
                    const scale = currentDistance / startDistance;
                    const cells = document.querySelectorAll('.cell');
                    cells.forEach(cell => {
                        const currentFontSize = parseFloat(cell.style.fontSize) || 14;
                        cell.style.fontSize = Math.min(Math.max(currentFontSize * scale, 10), 24) + 'px';
                    });
                }
            }
        }, { passive: false });
    }
    
    setupResizeObserver() {
        // Observar mudanças no tamanho da tela
        window.addEventListener('resize', () => {
            this.handleResize();
        });
    }
    
    handleResize() {
        const isMobile = window.innerWidth <= 768;
        
        // Ajustar quantidade de colunas visíveis
        const spreadsheet = document.getElementById('spreadsheet');
        const container = document.querySelector('.spreadsheet-container');
        
        if (isMobile) {
            // Otimização para mobile: reduzir número de linhas renderizadas
            const maxVisibleRows = Math.floor(container.clientHeight / 40);
            // Implementar virtualização básica se necessário
        }
        
        // Re-renderizar se necessário
        if (this.activeSheet) {
            this.render();
        }
    }
    
    updateRowCount() {
        const rowCount = this.activeSheet.rows.length;
        // Implementar atualização do contador de linhas
        console.log(`Total de linhas: ${rowCount}`);
    }
    
    setupToolbar() {
        document.getElementById('boldBtn').addEventListener('click', () => {
            this.toggleBold();
        });
        
        document.getElementById('italicBtn').addEventListener('click', () => {
            this.toggleItalic();
        });
        
        document.getElementById('fontFamily').addEventListener('change', (e) => {
            this.setFontFamily(e.target.value);
        });
        
        document.getElementById('fontSize').addEventListener('change', (e) => {
            this.setFontSize(e.target.value);
        });
        
        document.getElementById('bgColor').addEventListener('input', (e) => {
            this.setBackgroundColor(e.target.value);
        });
        
        document.getElementById('textColor').addEventListener('input', (e) => {
            this.setTextColor(e.target.value);
        });
        
        document.getElementById('cellType').addEventListener('change', (e) => {
            this.setCellType(e.target.value);
        });
        
        document.getElementById('insertRowBtn').addEventListener('click', () => {
            this.insertRow();
        });
        
        document.getElementById('deleteRowBtn').addEventListener('click', () => {
            this.deleteRow();
        });
        
        document.getElementById('insertColBtn').addEventListener('click', () => {
            this.insertCol();
        });
        
        document.getElementById('deleteColBtn').addEventListener('click', () => {
            this.deleteCol();
        });
        
        document.getElementById('addSheetBtn').addEventListener('click', () => {
            this.createSheet(`Planilha${this.sheets.length + 1}`);
        });
    }
    
    setupExport() {
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.exportToExcel();
        });
    }
    
    setupContextMenu() {
        document.addEventListener('contextmenu', (e) => {
            if (e.target.classList.contains('cell') && !e.target.classList.contains('header')) {
                e.preventDefault();
                const row = parseInt(e.target.dataset.row);
                const col = parseInt(e.target.dataset.col);
                this.showContextMenu(e.clientX, e.clientY, { type: 'cell', row: row, col: col });
            }
        });
        
        document.addEventListener('click', () => {
            const contextMenu = document.querySelector('.context-menu');
            if (contextMenu) {
                contextMenu.remove();
            }
        });
    }
    
    showContextMenu(x, y, target) {
        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = x + 'px';
        menu.style.top = y + 'px';
        
        let items = [];
        
        if (target.type === 'cell') {
            items = [
                { label: 'Copiar', action: () => this.copyCell(target.row, target.col) },
                { label: 'Colar', action: () => this.pasteCell(target.row, target.col) },
                { label: 'Cortar', action: () => this.cutCell(target.row, target.col) },
                { label: '---' },
                { label: 'Inserir linha acima', action: () => this.insertRowAbove(target.row) },
                { label: 'Inserir linha abaixo', action: () => this.insertRowBelow(target.row) },
                { label: 'Excluir linha', action: () => this.deleteRow(target.row) },
                { label: '---' },
                { label: 'Limpar conteúdo', action: () => this.clearCell(target.row, target.col) },
                { label: 'Formatar célula', action: () => this.formatCell(target.row, target.col) }
            ];
        } else if (target.type === 'row') {
            items = [
                { label: 'Inserir linha acima', action: () => this.insertRowAbove(target.row) },
                { label: 'Inserir linha abaixo', action: () => this.insertRowBelow(target.row) },
                { label: 'Excluir linha', action: () => this.deleteRow(target.row) }
            ];
        } else if (target.type === 'column') {
            items = [
                { label: 'Inserir coluna à esquerda', action: () => this.insertColLeft(target.col) },
                { label: 'Inserir coluna à direita', action: () => this.insertColRight(target.col) },
                { label: 'Excluir coluna', action: () => this.deleteCol(target.col) }
            ];
        }
        
        items.forEach(item => {
            if (item.label === '---') {
                const separator = document.createElement('div');
                separator.className = 'context-menu-separator';
                menu.appendChild(separator);
            } else {
                const menuItem = document.createElement('div');
                menuItem.className = 'context-menu-item';
                menuItem.textContent = item.label;
                menuItem.addEventListener('click', () => {
                    item.action();
                    menu.remove();
                });
                menu.appendChild(menuItem);
            }
        });
        
        document.body.appendChild(menu);
    }
    
    // Métodos de formatação
    toggleBold() {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.bold = !sheet.data[row][col].style.bold;
        });
        this.render();
    }
    
    toggleItalic() {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.italic = !sheet.data[row][col].style.italic;
        });
        this.render();
    }
    
    setFontFamily(fontFamily) {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.fontFamily = fontFamily;
        });
        this.render();
    }
    
    setFontSize(fontSize) {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.fontSize = fontSize + 'px';
        });
        this.render();
    }
    
    setBackgroundColor(color) {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.backgroundColor = color;
        });
        this.render();
    }
    
    setTextColor(color) {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].style.color = color;
        });
        this.render();
    }
    
    setCellType(type) {
        const sheet = this.activeSheet;
        this.selectedCells.forEach(cellRef => {
            const { row, col } = this.parseCellRef(cellRef);
            if (!sheet.data[row]) sheet.data[row] = {};
            if (!sheet.data[row][col]) sheet.data[row][col] = { style: {} };
            
            sheet.data[row][col].type = type;
        });
        this.render();
    }
    
    // Métodos de inserção/exclusão
    insertRow() {
        const sheet = this.activeSheet;
        const lastRow = sheet.rows.length;
        
        // Adicionar nova linha
        for (let i = 0; i < 10; i++) {
            this.ensureRow(sheet, lastRow + i);
        }
        
        this.render();
        this.updateRowCount();
    }
    
    deleteRow(rowIndex = null) {
        const sheet = this.activeSheet;
        
        if (rowIndex === null) {
            rowIndex = sheet.rows.length - 1;
        }
        
        // Excluir linha
        if (sheet.data[rowIndex]) {
            delete sheet.data[rowIndex];
        }
        
        // Reorganizar dados
        for (let i = rowIndex; i < sheet.rows.length - 1; i++) {
            sheet.data[i] = sheet.data[i + 1] || {};
        }
        
        if (sheet.data[sheet.rows.length - 1]) {
            delete sheet.data[sheet.rows.length - 1];
        }
        
        sheet.rows.pop();
        
        this.render();
        this.updateRowCount();
    }
    
    insertCol() {
        const sheet = this.activeSheet;
        sheet.cols++;
        sheet.colWidths.push(100);
        
        this.render();
    }
    
    deleteCol(colIndex = null) {
        const sheet = this.activeSheet;
        
        if (colIndex === null) {
            colIndex = sheet.cols - 1;
        }
        
        // Excluir coluna
        Object.keys(sheet.data).forEach(rowKey => {
            if (sheet.data[rowKey][colIndex]) {
                delete sheet.data[rowKey][colIndex];
            }
            // Reorganizar colunas
            for (let i = colIndex; i < sheet.cols - 1; i++) {
                sheet.data[rowKey][i] = sheet.data[rowKey][i + 1] || {};
            }
            if (sheet.data[rowKey][sheet.cols - 1]) {
                delete sheet.data[rowKey][sheet.cols - 1];
            }
        });
        
        sheet.cols--;
        sheet.colWidths.pop();
        
        this.render();
    }
    
    insertRowAbove(rowIndex) {
        const sheet = this.activeSheet;
        
        // Mover dados para baixo
        for (let i = sheet.rows.length - 1; i >= rowIndex; i--) {
            sheet.data[i + 1] = sheet.data[i] || {};
        }
        
        sheet.data[rowIndex] = {};
        sheet.rows.push(null);
        
        this.render();
        this.updateRowCount();
    }
    
    insertRowBelow(rowIndex) {
        const sheet = this.activeSheet;
        
        // Mover dados para baixo
        for (let i = sheet.rows.length - 1; i > rowIndex; i--) {
            sheet.data[i + 1] = sheet.data[i] || {};
        }
        
        sheet.data[rowIndex + 1] = {};
        sheet.rows.push(null);
        
        this.render();
        this.updateRowCount();
    }
    
    insertColLeft(colIndex) {
        const sheet = this.activeSheet;
        
        // Mover colunas para direita
        Object.keys(sheet.data).forEach(rowKey => {
            for (let i = sheet.cols - 1; i >= colIndex; i--) {
                sheet.data[rowKey][i + 1] = sheet.data[rowKey][i] || {};
            }
            sheet.data[rowKey][colIndex] = {};
        });
        
        sheet.cols++;
        sheet.colWidths.splice(colIndex, 0, 100);
        
        this.render();
    }
    
    insertColRight(colIndex) {
        const sheet = this.activeSheet;
        
        // Mover colunas para direita
        Object.keys(sheet.data).forEach(rowKey => {
            for (let i = sheet.cols - 1; i > colIndex; i--) {
                sheet.data[rowKey][i + 1] = sheet.data[rowKey][i] || {};
            }
            sheet.data[rowKey][colIndex + 1] = {};
        });
        
        sheet.cols++;
        sheet.colWidths.splice(colIndex + 1, 0, 100);
        
        this.render();
    }
    
    // Métodos de clipboard
    copyCell(row, col) {
        const sheet = this.activeSheet;
        const cellData = sheet.data[row] && sheet.data[row][col];
        
        if (cellData) {
            this.clipboard = JSON.parse(JSON.stringify(cellData));
        }
    }
    
    pasteCell(row, col) {
        if (this.clipboard.length === 0) return;
        
        const sheet = this.activeSheet;
        if (!sheet.data[row]) sheet.data[row] = {};
        sheet.data[row][col] = JSON.parse(JSON.stringify(this.clipboard));
        
        this.render();
    }
    
    cutCell(row, col) {
        this.copyCell(row, col);
        this.clearCell(row, col);
    }
    
    clearCell(row, col) {
        const sheet = this.activeSheet;
        if (sheet.data[row]) {
            delete sheet.data[row][col];
        }
        this.render();
    }
    
    formatCell(row, col) {
        const sheet = this.activeSheet;
        const cellData = sheet.data[row] && sheet.data[row][col];
        
        if (cellData) {
            // Abrir dialog de formatação
            alert(`Célula ${this.getCellRef(row, col)}:\nValor: ${cellData.value || 'vazio'}\nTipo: ${cellData.type}`);
        }
    }
    
    selectRow(rowIndex) {
        // Selecionar linha inteira
        this.selectedCells = [];
        for (let col = 0; col < this.activeSheet.cols; col++) {
            this.selectedCells.push(this.getCellRef(rowIndex, col));
        }
        this.render();
    }
    
    selectColumn(colIndex) {
        // Selecionar coluna inteira
        this.selectedCells = [];
        for (let row = 0; row < this.activeSheet.rows.length; row++) {
            this.selectedCells.push(this.getCellRef(row, colIndex));
        }
        this.render();
    }
    
    bindEvents() {
        // Eventos do teclado
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'b') {
                e.preventDefault();
                this.toggleBold();
            } else if (e.ctrlKey && e.key === 'i') {
                e.preventDefault();
                this.toggleItalic();
            } else if (e.ctrlKey && e.key === 'c') {
                e.preventDefault();
                if (this.selectedCells.length > 0) {
                    const firstCell = this.parseCellRef(this.selectedCells[0]);
                    this.copyCell(firstCell.row, firstCell.col);
                }
            } else if (e.ctrlKey && e.key === 'v') {
                e.preventDefault();
                if (this.selectedCells.length > 0) {
                    const firstCell = this.parseCellRef(this.selectedCells[0]);
                    this.pasteCell(firstCell.row, firstCell.col);
                }
            } else if (e.ctrlKey && e.key === 'x') {
                e.preventDefault();
                if (this.selectedCells.length > 0) {
                    const firstCell = this.parseCellRef(this.selectedCells[0]);
                    this.cutCell(firstCell.row, firstCell.col);
                }
            } else if (e.key === 'Delete' || e.key === 'Backspace') {
                if (this.selectedCells.length > 0 && !this.isEditing) {
                    const firstCell = this.parseCellRef(this.selectedCells[0]);
                    this.clearCell(firstCell.row, firstCell.col);
                }
            }
        });
        
        // Eventos da barra de fórmula
        const formulaInput = document.getElementById('formulaInput');
        formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (this.selectedCells.length > 0) {
                    const firstCell = this.parseCellRef(this.selectedCells[0]);
                    this.setCellValue(this.activeSheet, firstCell.row, firstCell.col, formulaInput.value);
                    this.render();
                    formulaInput.value = '';
                }
            }
        });
    }
    
    updateSheetTabs() {
        const sheetsContainer = document.getElementById('sheets');
        
        // Remover tabs existentes (exceto botão de adicionar)
        const existingTabs = sheetsContainer.querySelectorAll('.sheet-tab');
        existingTabs.forEach(tab => tab.remove());
        
        // Adicionar tabs
        this.sheets.forEach((sheet, index) => {
            const tab = document.createElement('div');
            tab.className = 'sheet-tab' + (index === this.sheets.indexOf(this.activeSheet) ? ' active' : '');
            
            const input = document.createElement('input');
            input.value = sheet.name;
            input.readOnly = true;
            
            input.addEventListener('dblclick', () => {
                input.readOnly = false;
                input.focus();
                input.select();
            });
            
            input.addEventListener('blur', () => {
                input.readOnly = true;
                sheet.name = input.value;
            });
            
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    input.blur();
                }
            });
            
            tab.appendChild(input);
            
            tab.addEventListener('click', () => {
                this.activeSheet = sheet;
                this.render();
                this.updateSheetTabs();
            });
            
            sheetsContainer.insertBefore(tab, sheetsContainer.lastChild);
        });
    }
    
    exportToExcel() {
        const wb = XLSX.utils.book_new();
        
        this.sheets.forEach((sheet, index) => {
            const data = [];
            
            // Criar cabeçalho
            const header = [];
            for (let col = 0; col < sheet.cols; col++) {
                header.push(this.getColumnName(col));
            }
            data.push(header);
            
            // Adicionar dados
            for (let row = 0; row < sheet.rows.length; row++) {
                const rowData = [];
                for (let col = 0; col < sheet.cols; col++) {
                    const cellData = sheet.data[row] && sheet.data[row][col];
                    rowData.push(cellData ? cellData.value || '' : '');
                }
                data.push(rowData);
            }
            
            const ws = XLSX.utils.aoa_to_sheet(data);
            XLSX.utils.book_append_sheet(wb, ws, sheet.name || `Planilha${index + 1}`);
        });
        
        XLSX.writeFile(wb, 'planilha.xlsx');
    }
}

// Inicializar planilha
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheet = new Spreadsheet();
});