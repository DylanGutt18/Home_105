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
        expression = this.replaceCellReferences(sheet, expression);
        expression = expression.replace(/(\d+),(\d+)/g, '$1.$2');
        expression = expression.replace(/\^/g, '**');
        
        try {
            const result = Function('"use strict"; return (' + expression + ')')();
            return typeof result === 'number' ? parseFloat(result.toFixed(2)) : result;
        } catch (error) {
            return '#ERRO!';
        }
    }
    
    replaceCellReferences(sheet, expression) {
        return expression.replace(/([A-Z]+)(\d+)/g, (match, colName, rowNum) => {
            const row = parseInt(rowNum) - 1;
            const col = this.getColumnIndex(colName);
            const value = this.getNumericValue(sheet, row, col);
            return value;
        });
    }
    
    render() {
        const spreadsheet = document.getElementById('spreadsheet');
        const sheet = this.activeSheet;
        
        if (!sheet) return;
        
        // Limpar planilha
        spreadsheet.innerHTML = '';
        
        // Configurar grid com colunas e linhas corretas
        spreadsheet.style.display = 'grid';
        spreadsheet.style.gridTemplateColumns = `100px repeat(${sheet.cols}, ${sheet.colWidths.slice(0, sheet.cols).map(w => w + 'px').join(' ')})`;
        spreadsheet.style.gridTemplateRows = `35px repeat(${sheet.rows.length}, 35px)`;
        spreadsheet.style.gridAutoFlow = 'row';
        
        // Renderizar cabeçalho das colunas
        const cornerCell = document.createElement('div');
        cornerCell.className = 'cell header corner';
        cornerCell.style.gridColumn = '1';
        cornerCell.style.gridRow = '1';
        cornerCell.innerHTML = '';
        spreadsheet.appendChild(cornerCell);
        
        for (let col = 0; col < sheet.cols; col++) {
            const headerCell = document.createElement('div');
            headerCell.className = 'cell header';
            headerCell.textContent = this.getColumnName(col);
            headerCell.style.gridColumn = (col + 2).toString();
            headerCell.style.gridRow = '1';
            headerCell.style.width = sheet.colWidths[col] + 'px';
            headerCell.style.position = 'sticky';
            headerCell.style.top = '0';
            headerCell.style.zIndex = '5';
            
            // Adicionar handle para redimensionar
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'resize-handle';
            resizeHandle.addEventListener('mousedown', (e) => {
                e.preventDefault();
                this.startColResize(col, e);
            });
            headerCell.appendChild(resizeHandle);
            
            spreadsheet.appendChild(headerCell);
        }
        
        // Renderizar linhas de dados
        for (let row = 0; row < sheet.rows.length; row++) {
            // Número da linha (coluna fixa à esquerda)
            const rowNumber = document.createElement('div');
            rowNumber.className = 'cell header row-header';
            rowNumber.textContent = row + 1;
            rowNumber.style.gridColumn = '1';
            rowNumber.style.gridRow = (row + 2).toString();
            rowNumber.style.position = 'sticky';
            rowNumber.style.left = '0';
            rowNumber.style.zIndex = '5';
            spreadsheet.appendChild(rowNumber);
            
            // Células de dados
            for (let col = 0; col < sheet.cols; col++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.dataset.row = row;
                cell.dataset.col = col;
                cell.tabIndex = 0;
                cell.style.gridColumn = (col + 2).toString();
                cell.style.gridRow = (row + 2).toString();
                
                const cellData = sheet.data[row] && sheet.data[row][col];
                
                if (cellData) {
                    // Adicionar hiperlink se existir
                    if (cellData.hyperlink) {
                        cell.innerHTML = `<a href="${cellData.hyperlink}" target="_blank" style="color: #1a73e8; text-decoration: underline;">${cellData.value}</a>`;
                        cell.style.cursor = 'pointer';
                    } else {
                        cell.textContent = cellData.value;
                    }
                    
                    if (cellData.style) {
                        cell.style.fontFamily = cellData.style.fontFamily;
                        cell.style.fontSize = cellData.style.fontSize;
                        cell.style.fontWeight = cellData.style.bold ? 'bold' : 'normal';
                        cell.style.fontStyle = cellData.style.italic ? 'italic' : 'normal';
                        cell.style.backgroundColor = cellData.style.backgroundColor;
                        cell.style.color = cellData.style.color;
                    }
                    
                    if (cellData.type === 'number' || cellData.type === 'currency') {
                        cell.style.textAlign = 'right';
                    } else {
                        cell.style.textAlign = 'left';
                    }
                }
                
                cell.addEventListener('mousedown', (e) => this.handleCellMouseDown(e, row, col));
                cell.addEventListener('mouseover', (e) => this.handleCellMouseOver(e, row, col));
                cell.addEventListener('dblclick', (e) => this.handleCellDoubleClick(e, row, col));
                cell.addEventListener('keydown', (e) => this.handleCellKeyDown(e, row, col));
                
                spreadsheet.appendChild(cell);
            }
        }
        
        this.updateFormulaBar();
        this.updateRowCount();
        this.syncFormulaBar();
    }
    
    handleCellMouseDown(e, row, col) {
        if (e.button === 2) return; // Botão direito
        
        // Parar edição anterior se existir
        if (this.isEditing) {
            this.stopEditing();
        }
        
        // Iniciar seleção
        this.isSelecting = true;
        this.selectionStart = { row, col };
        
        // Se for Shift+click, adicionar ao intervalo
        if (e.shiftKey) {
            const lastCell = this.selectedCells[this.selectedCells.length - 1];
            if (lastCell) {
                this.selectRange(lastCell.row, lastCell.col, row, col);
            }
        } else if (e.ctrlKey || e.metaKey) {
            // Ctrl+click para seleção múltipla
            const existingIndex = this.selectedCells.findIndex(cell => cell.row === row && cell.col === col);
            if (existingIndex >= 0) {
                this.selectedCells.splice(existingIndex, 1);
            } else {
                this.selectedCells.push({ row, col });
            }
            this.updateSelectionVisual();
        } else {
            // Clique simples
            document.querySelectorAll('.cell.selected').forEach(cell => {
                cell.classList.remove('selected');
            });
            
            e.target.classList.add('selected');
            this.selectedCells = [{ row, col }];
            
            document.getElementById('cellRef').value = this.getCellRef(row, col);
            
            const cellData = this.activeSheet.data[row] && this.activeSheet.data[row][col];
            const formulaInput = document.getElementById('formulaInput');
            
            if (cellData && cellData.formula) {
                formulaInput.value = cellData.formula;
            } else {
                const cellValue = this.getCellValue(this.activeSheet, row, col);
                formulaInput.value = cellValue;
            }
            
            this.updateToolbarState(row, col);
            this.syncFormulaBar();
        }
    }
    
    handleCellMouseOver(e, row, col) {
        if (this.isSelecting && this.selectionStart) {
            this.selectRange(this.selectionStart.row, this.selectionStart.col, row, col);
        }
    }
    
    handleCellDoubleClick(e, row, col) {
        this.startEditing(row, col);
    }
    
    handleCellKeyDown(e, row, col) {
        // Se não está editando e o usuário digita um caractere, iniciar edição
        if (!this.isEditing) {
            const key = e.key;
            
            // Ignorar teclas de navegação e atalhos
            if (key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                this.startEditing(row, col, key);
                e.preventDefault();
            }
        }
    }
    
    startEditing(row, col, initialValue = null) {
        // Se já está editando outra célula, parar a edição anterior
        if (this.isEditing) {
            this.stopEditing();
        }
        
        this.isEditing = true;
        this.currentEditingCell = { row, col };
        
        const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
        const cellData = this.activeSheet.data[row] && this.activeSheet.data[row][col];
        const currentValue = cellData && cellData.formula ? cellData.formula : this.getCellValue(this.activeSheet, row, col);
        
        // Marcar célula como editando
        cell.classList.add('editing');
        
        // Criar input dentro da célula
        const input = document.createElement('input');
        input.value = initialValue || currentValue;
        input.type = 'text';
        input.placeholder = '';
        input.style.width = '100%';
        input.style.height = '100%';
        input.style.border = 'none';
        input.style.outline = 'none';
        input.style.background = 'transparent';
        input.style.fontFamily = 'inherit';
        input.style.fontSize = 'inherit';
        input.style.color = 'inherit';
        input.style.textAlign = 'inherit';
        
        // Eventos do input
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.stopEditing(true);
                this.moveSelection(row + 1, col);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                this.stopEditing(false);
            } else if (e.key === 'Tab') {
                e.preventDefault();
                this.stopEditing(true);
                this.moveSelection(row, col + 1);
            }
        });
        
        input.addEventListener('input', (e) => {
            // Atualizar a caixa de edição (formula bar) em tempo real
            document.getElementById('formulaInput').value = e.target.value;
        });
        
        input.addEventListener('blur', () => {
            this.stopEditing(true);
        });
        
        // Limpar célula e adicionar input
        cell.innerHTML = '';
        cell.appendChild(input);
        
        // Focar e selecionar texto
        input.focus();
        
        if (initialValue) {
            // Se veio de digitação, colocar cursor no final
            input.setSelectionRange(input.value.length, input.value.length);
        } else {
            // Selecionar todo o texto
            input.select();
        }
        
        // Atualizar caixa de edição
        document.getElementById('formulaInput').value = initialValue || currentValue;
    }
    
    stopEditing(save = true) {
        if (!this.isEditing) return;
        
        const input = document.querySelector('.cell.editing input');
        if (input && save) {
            const value = input.value;
            const { row, col } = this.currentEditingCell;
            
            // Se o valor for diferente do original, salvar
            const originalValue = this.getCellValue(this.activeSheet, row, col);
            if (value !== originalValue) {
                this.saveState();
                this.setCellValue(this.activeSheet, row, col, value);
                this.recalculateFormulas();
            }
            
            this.render();
            document.getElementById('formulaInput').value = value;
            document.getElementById('cellRef').value = this.getCellRef(row, col);
        }
        
        this.isEditing = false;
        this.currentEditingCell = null;
    }
    
    recalculateFormulas() {
        const sheet = this.activeSheet;
        for (let row = 0; row < sheet.rows.length; row++) {
            for (let col = 0; col < sheet.cols; col++) {
                const cellData = sheet.data[row] && sheet.data[row][col];
                if (cellData && cellData.formula) {
                    cellData.value = this.processFormula(sheet, row, col, cellData.formula);
                }
            }
        }
    }
    
    updateFormulaBar() {
        const formulaInput = document.getElementById('formulaInput');
        
        const newFormulaInput = formulaInput.cloneNode(true);
        formulaInput.parentNode.replaceChild(newFormulaInput, formulaInput);
        
        newFormulaInput.addEventListener('change', (e) => {
            const value = e.target.value;
            const { row, col } = this.currentEditingCell || this.selectedCells[0] || { row: 0, col: 0 };
            
            if (row !== undefined && col !== undefined) {
                this.saveState();
                this.setCellValue(this.activeSheet, row, col, value);
                this.recalculateFormulas();
                this.render();
                document.getElementById('cellRef').value = this.getCellRef(row, col);
            }
        });
        
        newFormulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const value = e.target.value;
                const { row, col } = this.currentEditingCell || this.selectedCells[0] || { row: 0, col: 0 };
                
                if (row !== undefined && col !== undefined) {
                    this.saveState();
                    this.setCellValue(this.activeSheet, row, col, value);
                    this.recalculateFormulas();
                    this.render();
                    
                    const nextRow = row + 1;
                    if (nextRow < this.activeSheet.rows.length) {
                        this.moveSelection(nextRow, col);
                    }
                }
            }
        });
        
        this.syncFormulaBar();
    }
    
    syncFormulaBar() {
        const formulaInput = document.getElementById('formulaInput');
        const cellRef = document.getElementById('cellRef');
        
        if (this.selectedCells.length > 0) {
            const { row, col } = this.selectedCells[0];
            const cellData = this.activeSheet.data[row] && this.activeSheet.data[row][col];
            
            cellRef.value = this.getCellRef(row, col);
            
            if (cellData && cellData.formula) {
                formulaInput.value = cellData.formula;
            } else {
                const cellValue = this.getCellValue(this.activeSheet, row, col);
                formulaInput.value = cellValue;
            }
        } else {
            cellRef.value = '';
            formulaInput.value = '';
        }
    }
    
    updateToolbarState(row, col) {
        const sheet = this.activeSheet;
        const cellData = sheet.data[row] && sheet.data[row][col];
        
        if (cellData && cellData.style) {
            document.getElementById('boldBtn').classList.toggle('active', cellData.style.bold);
            document.getElementById('italicBtn').classList.toggle('active', cellData.style.italic);
            document.getElementById('fontFamily').value = cellData.style.fontFamily;
            document.getElementById('fontSize').value = cellData.style.fontSize.replace('px', '');
            document.getElementById('bgColor').value = cellData.style.backgroundColor;
            document.getElementById('textColor').value = cellData.style.color;
        }
    }
    
    updateSheetTabs() {
        const sheetsContainer = document.getElementById('sheets');
        const addBtn = document.getElementById('addSheetBtn');
        sheetsContainer.innerHTML = '';
        sheetsContainer.appendChild(addBtn);
        
        this.sheets.forEach((sheet) => {
            const tab = document.createElement('div');
            tab.className = `sheet-tab ${sheet.id === this.activeSheet.id ? 'active' : ''}`;
            tab.dataset.sheetId = sheet.id;
            
            const tabName = document.createElement('input');
            tabName.value = sheet.name;
            tabName.addEventListener('change', (e) => {
                sheet.name = e.target.value;
            });
            
            tab.appendChild(tabName);
            tab.addEventListener('click', () => {
                this.activeSheet = sheet;
                this.render();
                this.updateSheetTabs();
            });
            
            sheetsContainer.appendChild(tab);
        });
    }
    
    setupToolbar() {
        document.getElementById('fontFamily').addEventListener('change', (e) => {
            this.applyStyleToSelected({ fontFamily: e.target.value });
        });
        
        document.getElementById('fontSize').addEventListener('change', (e) => {
            this.applyStyleToSelected({ fontSize: e.target.value + 'px' });
        });
        
        document.getElementById('boldBtn').addEventListener('click', () => {
            const cellData = this.getSelectedCellData();
            this.saveState();
            this.applyStyleToSelected({ bold: !cellData.style.bold });
        });
        
        document.getElementById('italicBtn').addEventListener('click', () => {
            const cellData = this.getSelectedCellData();
            this.saveState();
            this.applyStyleToSelected({ italic: !cellData.style.italic });
        });
        
        document.getElementById('bgColor').addEventListener('change', (e) => {
            this.saveState();
            this.applyStyleToSelected({ backgroundColor: e.target.value });
        });
        
        document.getElementById('textColor').addEventListener('change', (e) => {
            this.saveState();
            this.applyStyleToSelected({ color: e.target.value });
        });
        
        document.getElementById('cellType').addEventListener('change', (e) => {
            this.saveState();
            this.applyTypeToSelected(e.target.value);
        });
        
        document.getElementById('insertRowBtn').addEventListener('click', () => {
            this.saveState();
            this.insertRow();
        });
        
        document.getElementById('deleteRowBtn').addEventListener('click', () => {
            this.saveState();
            this.deleteRow();
        });
        
        document.getElementById('insertColBtn').addEventListener('click', () => {
            this.saveState();
            this.insertColumn();
        });
        
        document.getElementById('deleteColBtn').addEventListener('click', () => {
            this.saveState();
            this.deleteColumn();
        });
        
        document.getElementById('addSheetBtn').addEventListener('click', () => {
            const newName = `Planilha${this.sheets.length + 1}`;
            this.createSheet(newName);
        });
        
        // Adicionar botões de busca e mesclar
        const toolbar = document.querySelector('.toolbar');
        const searchGroup = document.createElement('div');
        searchGroup.className = 'toolbar-group';
        searchGroup.innerHTML = `
            <input type="text" id="searchInput" placeholder="🔍 Buscar..." style="width: 150px;">
            <button class="toolbar-btn" id="searchBtn" title="Buscar">🔍</button>
        `;
        
        const mergeGroup = document.createElement('div');
        mergeGroup.className = 'toolbar-group';
        mergeGroup.innerHTML = `
            <button class="toolbar-btn" id="mergeBtn" title="Mesclar células">🔗</button>
            <button class="toolbar-btn" id="hyperlinkBtn" title="Adicionar hiperlink">🔗</button>
        `;
        
        toolbar.appendChild(searchGroup);
        toolbar.appendChild(mergeGroup);
        
        document.getElementById('searchBtn').addEventListener('click', () => {
            this.searchInSheet();
        });
        
        document.getElementById('searchInput').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                this.searchInSheet();
            }
        });
        
        document.getElementById('mergeBtn').addEventListener('click', () => {
            this.mergeCells();
        });
        
        document.getElementById('hyperlinkBtn').addEventListener('click', () => {
            this.addHyperlink();
        });
    }
    
    searchInSheet() {
        const searchTerm = document.getElementById('searchInput').value.trim();
        if (!searchTerm) return;
        
        const sheet = this.activeSheet;
        let found = false;
        
        for (let row = 0; row < sheet.rows.length; row++) {
            for (let col = 0; col < sheet.cols; col++) {
                const cellData = sheet.data[row] && sheet.data[row][col];
                if (cellData && cellData.value && String(cellData.value).toLowerCase().includes(searchTerm.toLowerCase())) {
                    this.selectRange(row, col, row, col);
                    const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
                    if (cell) {
                        cell.classList.add('found');
                        setTimeout(() => cell.classList.remove('found'), 2000);
                        cell.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }
                    found = true;
                    break;
                }
            }
            if (found) break;
        }
        
        if (!found) {
            alert(`🔍 "${searchTerm}" não encontrado na planilha.`);
        }
    }
    
    mergeCells() {
        if (this.selectedCells.length < 2) {
            alert('Selecione pelo menos 2 células para mesclar!');
            return;
        }
        
        const sheet = this.activeSheet;
        const minRow = Math.min(...this.selectedCells.map(c => c.row));
        const maxRow = Math.max(...this.selectedCells.map(c => c.row));
        const minCol = Math.min(...this.selectedCells.map(c => c.col));
        const maxCol = Math.max(...this.selectedCells.map(c => c.col));
        
        // Obter valor da célula superior esquerda
        const value = this.getCellValue(sheet, minRow, minCol);
        
        // Limpar células mescladas
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                if (!(r === minRow && c === minCol)) {
                    if (sheet.data[r] && sheet.data[r][c]) {
                        sheet.data[r][c].value = '';
                        sheet.data[r][c].merged = true;
                    }
                }
            }
        }
        
        // Marcar célula principal como mesclada
        if (!sheet.data[minRow]) sheet.data[minRow] = {};
        if (!sheet.data[minRow][minCol]) {
            sheet.data[minRow][minCol] = {
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
        
        sheet.data[minRow][minCol].value = value;
        sheet.data[minRow][minCol].merged = true;
        
        this.render();
    }
    
    addHyperlink() {
        if (this.selectedCells.length === 0) {
            alert('Selecione uma célula para adicionar hiperlink!');
            return;
        }
        
        const { row, col } = this.selectedCells[0];
        const url = prompt('Digite a URL do hiperlink (ex: https://www.google.com):');
        
        if (url) {
            const sheet = this.activeSheet;
            if (!sheet.data[row]) sheet.data[row] = {};
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
            
            sheet.data[row][col].hyperlink = url;
            if (!sheet.data[row][col].value) {
                sheet.data[row][col].value = url;
            }
            
            this.render();
        }
    }
    
    setupContextMenu() {
        const spreadsheet = document.getElementById('spreadsheet');
        
        spreadsheet.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            
            const target = e.target;
            if (target.classList.contains('cell') && !target.classList.contains('header')) {
                const row = parseInt(target.dataset.row);
                const col = parseInt(target.dataset.col);
                
                // Selecionar célula
                this.selectedCells = [{ row, col }];
                this.updateSelectionVisual();
                
                // Criar menu de contexto
                const menu = document.createElement('div');
                menu.className = 'context-menu';
                menu.style.left = e.clientX + 'px';
                menu.style.top = e.clientY + 'px';
                
                menu.innerHTML = `
                    <div class="context-menu-item" data-action="copy">📋 Copiar</div>
                    <div class="context-menu-item" data-action="paste">📥 Colar</div>
                    <div class="context-menu-separator"></div>
                    <div class="context-menu-item" data-action="insertRow">➕ Inserir linha</div>
                    <div class="context-menu-item" data-action="insertCol">➕ Inserir coluna</div>
                    <div class="context-menu-separator"></div>
                    <div class="context-menu-item" data-action="hyperlink">🔗 Adicionar hiperlink</div>
                    <div class="context-menu-item" data-action="clear">🗑️ Limpar célula</div>
                    <div class="context-menu-separator"></div>
                    <div class="context-menu-item" data-action="undo">🔄 Desfazer</div>
                    <div class="context-menu-item" data-action="redo">↩️ Refazer</div>
                `;
                
                document.body.appendChild(menu);
                
                // Adicionar eventos
                menu.querySelectorAll('.context-menu-item').forEach(item => {
                    item.addEventListener('click', (e) => {
                        const action = e.target.dataset.action;
                        this.handleContextAction(action);
                        menu.remove();
                    });
                });
                
                // Fechar menu ao clicar fora
                document.addEventListener('click', function closeMenu(e) {
                    if (!menu.contains(e.target)) {
                        menu.remove();
                        document.removeEventListener('click', closeMenu);
                    }
                });
            }
        });
    }
    
    handleContextAction(action) {
        switch(action) {
            case 'copy':
                this.copySelection();
                break;
            case 'paste':
                this.pasteSelection();
                break;
            case 'insertRow':
                this.insertRow();
                break;
            case 'insertCol':
                this.insertColumn();
                break;
            case 'hyperlink':
                this.addHyperlink();
                break;
            case 'clear':
                this.clearSelection();
                break;
            case 'undo':
                this.undo();
                break;
            case 'redo':
                this.redo();
                break;
        }
    }
    
    clearSelection() {
        this.saveState();
        const sheet = this.activeSheet;
        this.selectedCells.forEach(({ row, col }) => {
            if (sheet.data[row] && sheet.data[row][col]) {
                sheet.data[row][col].value = '';
                sheet.data[row][col].formula = null;
                sheet.data[row][col].hyperlink = null;
            }
        });
        this.render();
    }
    
    copySelection() {
        if (this.selectedCells.length === 0) return;
        
        // Determinar dimensões da seleção
        const rows = [...new Set(this.selectedCells.map(c => c.row))];
        const cols = [...new Set(this.selectedCells.map(c => c.col))];
        
        const minRow = Math.min(...rows);
        const maxRow = Math.max(...rows);
        const minCol = Math.min(...cols);
        const maxCol = Math.max(...cols);
        
        this.clipboard = [];
        
        for (let r = minRow; r <= maxRow; r++) {
            const rowData = [];
            for (let c = minCol; c <= maxCol; c++) {
                const cellData = this.activeSheet.data[r] && this.activeSheet.data[r][c];
                rowData.push(cellData ? { ...cellData } : null);
            }
            this.clipboard.push(rowData);
        }
        
        // Feedback visual
        const cell = document.querySelector(`.cell[data-row="${this.selectedCells[0].row}"][data-col="${this.selectedCells[0].col}"]`);
        if (cell) {
            cell.style.outline = '2px dashed #1a73e8';
            setTimeout(() => {
                cell.style.outline = '';
            }, 500);
        }
    }
    
    pasteSelection() {
        if (this.clipboard.length === 0) {
            alert('Área de transferência vazia! Copie células primeiro.');
            return;
        }
        
        this.saveState();
        
        const { row, col } = this.selectedCells[0] || { row: 0, col: 0 };
        const sheet = this.activeSheet;
        
        this.clipboard.forEach((rowData, rIndex) => {
            rowData.forEach((cellData, cIndex) => {
                const targetRow = row + rIndex;
                const targetCol = col + cIndex;
                
                if (cellData) {
                    if (!sheet.data[targetRow]) sheet.data[targetRow] = {};
                    sheet.data[targetRow][targetCol] = { ...cellData };
                }
            });
        });
        
        this.render();
    }
    
    saveState() {
        // Salvar estado para desfazer
        const state = JSON.stringify({
            sheets: this.sheets.map(sheet => ({
                id: sheet.id,
                name: sheet.name,
                rows: [...sheet.rows],
                cols: sheet.cols,
                data: JSON.parse(JSON.stringify(sheet.data)),
                colWidths: [...sheet.colWidths]
            })),
            activeSheetId: this.activeSheet ? this.activeSheet.id : null
        });
        
        this.undoStack.push(state);
        if (this.undoStack.length > 50) {
            this.undoStack.shift();
        }
        this.redoStack = [];
    }
    
    undo() {
        if (this.undoStack.length === 0) {
            alert('Nada para desfazer!');
            return;
        }
        
        const currentState = JSON.stringify({
            sheets: this.sheets.map(sheet => ({
                id: sheet.id,
                name: sheet.name,
                rows: [...sheet.rows],
                cols: sheet.cols,
                data: JSON.parse(JSON.stringify(sheet.data)),
                colWidths: [...sheet.colWidths]
            })),
            activeSheetId: this.activeSheet ? this.activeSheet.id : null
        });
        
        this.redoStack.push(currentState);
        
        const previousState = JSON.parse(this.undoStack.pop());
        this.restoreState(previousState);
    }
    
    redo() {
        if (this.redoStack.length === 0) {
            alert('Nada para refazer!');
            return;
        }
        
        const currentState = JSON.stringify({
            sheets: this.sheets.map(sheet => ({
                id: sheet.id,
                name: sheet.name,
                rows: [...sheet.rows],
                cols: sheet.cols,
                data: JSON.parse(JSON.stringify(sheet.data)),
                colWidths: [...sheet.colWidths]
            })),
            activeSheetId: this.activeSheet ? this.activeSheet.id : null
        });
        
        this.undoStack.push(currentState);
        
        const nextState = JSON.parse(this.redoStack.pop());
        this.restoreState(nextState);
    }
    
    restoreState(state) {
        this.sheets = state.sheets.map(sheet => ({
            ...sheet,
            data: sheet.data
        }));
        
        this.activeSheet = this.sheets.find(sheet => sheet.id === state.activeSheetId) || this.sheets[0];
        
        this.render();
        this.updateSheetTabs();
        this.updateRowCount();
    }
    
    setupExport() {
        document.getElementById('exportBtn').addEventListener('click', () => {
            this.exportToExcel();
        });
    }
    
    exportToExcel() {
        try {
            const wb = XLSX.utils.book_new();
            
            this.sheets.forEach((sheet) => {
                let minRow = Infinity;
                let maxRow = -1;
                let minCol = Infinity;
                let maxCol = -1;
                
                for (let row = 0; row < sheet.rows.length; row++) {
                    for (let col = 0; col < sheet.cols; col++) {
                        const cellData = sheet.data[row] && sheet.data[row][col];
                        if (cellData && cellData.value !== '') {
                            minRow = Math.min(minRow, row);
                            maxRow = Math.max(maxRow, row);
                            minCol = Math.min(minCol, col);
                            maxCol = Math.max(maxCol, col);
                        }
                    }
                }
                
                if (maxRow === -1) {
                    return;
                }
                
                const data = [];
                
                for (let row = minRow; row <= maxRow; row++) {
                    const rowData = [];
                    for (let col = minCol; col <= maxCol; col++) {
                        const cellData = sheet.data[row] && sheet.data[row][col];
                        if (cellData && cellData.value !== '') {
                            rowData.push(cellData.value);
                        } else {
                            rowData.push('');
                        }
                    }
                    data.push(rowData);
                }
                
                const ws = XLSX.utils.aoa_to_sheet(data);
                
                const colWidths = [];
                for (let col = 0; col < (maxCol - minCol + 1); col++) {
                    let maxLength = 0;
                    for (let row = 0; row < data.length; row++) {
                        if (data[row][col]) {
                            maxLength = Math.max(maxLength, String(data[row][col]).length);
                        }
                    }
                    colWidths.push({ wch: Math.min(maxLength + 2, 50) });
                }
                ws['!cols'] = colWidths;
                
                XLSX.utils.book_append_sheet(wb, ws, sheet.name);
            });
            
            if (wb.SheetNames.length === 0) {
                alert('❌ Nenhuma célula preenchida encontrada para exportar!');
                return;
            }
            
            const fileName = `planilha_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(wb, fileName);
            
            let totalCells = 0;
            this.sheets.forEach(sheet => {
                for (let row = 0; row < sheet.rows.length; row++) {
                    for (let col = 0; col < sheet.cols; col++) {
                        const cellData = sheet.data[row] && sheet.data[row][col];
                        if (cellData && cellData.value !== '') {
                            totalCells++;
                        }
                    }
                }
            });
            
            alert(`✅ Planilha exportada com sucesso!\n\nNome do arquivo: ${fileName}\nAbas exportadas: ${wb.SheetNames.length}\nCélulas preenchidas: ${totalCells}`);
            
        } catch (error) {
            console.error('Erro ao exportar:', error);
            alert('❌ Erro ao exportar planilha: ' + error.message);
        }
    }
    
    updateRowCount() {
        const existingCounter = document.querySelector('.row-count');
        if (existingCounter) {
            existingCounter.remove();
        }
        
        const counter = document.createElement('div');
        counter.className = 'row-count';
        const totalRows = this.activeSheet ? this.activeSheet.rows.length : 0;
        const totalCols = this.activeSheet ? this.activeSheet.cols : 0;
        counter.textContent = `${totalRows} linhas × ${totalCols} colunas`;
        
        document.body.appendChild(counter);
    }
    
    getSelectedCellData() {
        const { row, col } = this.selectedCells[0] || { row: 0, col: 0 };
        const sheet = this.activeSheet;
        
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
        
        return sheet.data[row][col];
    }
    
    applyStyleToSelected(style) {
        this.selectedCells.forEach(({ row, col }) => {
            this.setCellStyle(this.activeSheet, row, col, style);
        });
        
        this.render();
    }
    
    applyTypeToSelected(type) {
        this.selectedCells.forEach(({ row, col }) => {
            const sheet = this.activeSheet;
            if (!sheet.data[row]) {
                sheet.data[row] = {};
            }
            if (!sheet.data[row][col]) {
                sheet.data[row][col] = {
                    value: '',
                    type: type,
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
            
            sheet.data[row][col].type = type;
            
            const value = sheet.data[row][col].value;
            if (value !== '') {
                this.setCellValue(sheet, row, col, value);
            }
        });
        
        this.render();
    }
    
    insertRow() {
        const sheet = this.activeSheet;
        const newRowIndex = this.selectedCells[0] ? this.selectedCells[0].row + 1 : sheet.rows.length;
        
        sheet.rows.splice(newRowIndex, 0, null);
        
        const newData = {};
        for (let i = 0; i < sheet.rows.length; i++) {
            if (i < newRowIndex) {
                newData[i] = sheet.data[i];
            } else if (i > newRowIndex) {
                newData[i] = sheet.data[i - 1];
            } else {
                newData[i] = {};
            }
        }
        
        sheet.data = newData;
        
        this.render();
        this.updateRowCount();
    }
    
    deleteRow() {
        const sheet = this.activeSheet;
        
        if (sheet.rows.length <= 1) {
            alert('Não é possível excluir a última linha!');
            return;
        }
        
        const deleteRowIndex = this.selectedCells[0] ? this.selectedCells[0].row : 0;
        
        sheet.rows.splice(deleteRowIndex, 1);
        
        const newData = {};
        const newRows = [];
        
        for (let i = 0; i < sheet.rows.length; i++) {
            const oldIndex = i >= deleteRowIndex ? i + 1 : i;
            newData[i] = sheet.data[oldIndex] || {};
            newRows.push(null);
        }
        
        sheet.rows = newRows;
        sheet.data = newData;
        
        this.selectedCells = [];
        
        this.render();
        this.updateRowCount();
    }
    
    insertColumn() {
        const sheet = this.activeSheet;
        const newColIndex = this.selectedCells[0] ? this.selectedCells[0].col + 1 : sheet.cols;
        
        sheet.cols++;
        
        Object.keys(sheet.data).forEach(rowKey => {
            const row = sheet.data[rowKey];
            const newRow = {};
            
            for (let i = 0; i < sheet.cols; i++) {
                if (i < newColIndex) {
                    newRow[i] = row[i];
                } else if (i > newColIndex) {
                    newRow[i] = row[i - 1];
                }
            }
            
            sheet.data[rowKey] = newRow;
        });
        
        this.render();
        this.updateRowCount();
    }
    
    deleteColumn() {
        const sheet = this.activeSheet;
        
        if (sheet.cols <= 1) {
            alert('Não é possível excluir a última coluna!');
            return;
        }
        
        const deleteColIndex = this.selectedCells[0] ? this.selectedCells[0].col : 0;
        
        sheet.cols--;
        
        Object.keys(sheet.data).forEach(rowKey => {
            const row = sheet.data[rowKey];
            const newRow = {};
            
            for (let i = 0; i < sheet.cols; i++) {
                const oldIndex = i >= deleteColIndex ? i + 1 : i;
                newRow[i] = row[oldIndex];
            }
            
            sheet.data[rowKey] = newRow;
        });
        
        this.selectedCells = [];
        
        this.render();
        this.updateRowCount();
    }
    
    bindEvents() {
        document.addEventListener('mousemove', (e) => {
            if (this.isSelecting) {
                const cell = e.target;
                if (cell.classList.contains('cell') && !cell.classList.contains('header')) {
                    const row = parseInt(cell.dataset.row);
                    const col = parseInt(cell.dataset.col);
                    this.selectRange(this.selectionStart.row, this.selectionStart.col, row, col);
                }
            }
        });
        
        document.addEventListener('mouseup', () => {
            this.isSelecting = false;
            this.selectionStart = null;
        });
        
        document.addEventListener('keydown', (e) => {
            if (this.isEditing) return;
            
            // Atalhos de teclado
            if (e.ctrlKey || e.metaKey) {
                switch(e.key.toLowerCase()) {
                    case 'c':
                        e.preventDefault();
                        this.copySelection();
                        break;
                    case 'v':
                        e.preventDefault();
                        this.pasteSelection();
                        break;
                    case 'z':
                        e.preventDefault();
                        if (e.shiftKey) {
                            this.redo();
                        } else {
                            this.undo();
                        }
                        break;
                    case 'y':
                        e.preventDefault();
                        this.redo();
                        break;
                    case 'f':
                        e.preventDefault();
                        document.getElementById('searchInput').focus();
                        break;
                }
            }
            
            const { row, col } = this.selectedCells[0] || { row: 0, col: 0 };
            
            switch(e.key) {
                case 'ArrowUp':
                    e.preventDefault();
                    this.moveSelection(row - 1, col);
                    break;
                case 'ArrowDown':
                    e.preventDefault();
                    this.moveSelection(row + 1, col);
                    break;
                case 'ArrowLeft':
                    e.preventDefault();
                    this.moveSelection(row, col - 1);
                    break;
                case 'ArrowRight':
                    e.preventDefault();
                    this.moveSelection(row, col + 1);
                    break;
                case 'Enter':
                    e.preventDefault();
                    this.moveSelection(row + 1, col);
                    break;
                case 'Tab':
                    e.preventDefault();
                    this.moveSelection(row, col + 1);
                    break;
                case 'Delete':
                    e.preventDefault();
                    this.clearSelection();
                    break;
            }
        });
    }
    
    moveSelection(row, col) {
        row = Math.max(0, Math.min(row, this.activeSheet.rows.length - 1));
        col = Math.max(0, Math.min(col, this.activeSheet.cols - 1));
        
        const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
        if (cell) {
            cell.click();
            cell.focus(); // Focar na célula para receber eventos de teclado
        }
    }
    
    selectRange(startRow, startCol, endRow, endCol) {
        this.selectedCells = [];
        
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        document.querySelectorAll('.cell.selected').forEach(cell => {
            cell.classList.remove('selected');
        });
        
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                this.selectedCells.push({ row: r, col: c });
                const cell = document.querySelector(`.cell[data-row="${r}"][data-col="${c}"]`);
                if (cell) {
                    cell.classList.add('selected');
                }
            }
        }
        
        this.updateToolbarState(endRow, endCol);
    }
}

// Initialize application
document.addEventListener('DOMContentLoaded', () => {
    const app = new Spreadsheet();
});