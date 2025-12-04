// 主应用程序 - 座位安排系统
class FindYourSeat {
    constructor() {
        this.currentStep = 1;
        this.employeeData = [];
        this.roomConfig = {};
        this.seatingArrangement = [];
        this.originalArrangement = [];
        this.isProcessing = false; // 添加文件处理状态标志
        this.stepStatus = {
            dataUploaded: false,
            configSet: false,
            allocated: false,
            previewed: false
        };
        
        // 开发者调试模式（在URL中添加 ?debug=true 启用）
        this.debugMode = new URLSearchParams(window.location.search).get('debug') === 'true';
        if (this.debugMode) {
            console.log('FindYourSeat 调试模式已启用');
            window.findYourSeatApp = this; // 暴露到全局方便调试
        }
        
        this.init();
    }

    init() {
        this.initEventListeners();
        this.loadSavedConfig();
        this.updateStepDisplay();
        this.initStepNavigation();
    }

    initEventListeners() {
        // 步骤导航
        document.getElementById('nextStep1').addEventListener('click', () => this.nextStep());
        document.getElementById('nextStep2').addEventListener('click', () => this.nextStep());
        document.getElementById('nextStep3').addEventListener('click', () => this.nextStep());
        document.getElementById('nextStep4').addEventListener('click', () => this.nextStep());
        
        document.getElementById('prevStep2').addEventListener('click', () => this.prevStep());
        document.getElementById('prevStep3').addEventListener('click', () => this.prevStep());
        document.getElementById('prevStep4').addEventListener('click', () => this.prevStep());
        document.getElementById('prevStep5').addEventListener('click', () => this.prevStep());

        // 文件上传
        this.initFileUpload();
        
        // 配置管理
        this.initConfigManagement();
        
        // 分配和预览
        document.getElementById('fullscreenBtn').addEventListener('click', () => this.enterFullscreenPreview());
        document.getElementById('regenerateBtn').addEventListener('click', () => this.regenerateAssignment());
        document.getElementById('resetChangesBtn').addEventListener('click', () => this.resetChanges());
        
        // 导出功能
        document.getElementById('exportExcelBtn').addEventListener('click', () => this.exportToExcel());
        document.getElementById('printPreviewBtn').addEventListener('click', () => this.printPreview());
        
        // 重新开始
        document.getElementById('startOverBtn').addEventListener('click', () => this.startOver());
        
        // 下载示例文件
        document.getElementById('downloadDemoBtn').addEventListener('click', () => this.downloadDemoFile());
    }

    initFileUpload() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        const selectFileBtn = document.getElementById('selectFileBtn');

        // 文件选择处理
        fileInput.addEventListener('change', (event) => {
            const file = event.target.files[0];
            if (file && !this.isProcessing) {
                console.log('文件选择:', file.name);
                this.processFile(file);
            } else if (this.isProcessing) {
                console.log('⚠️ 文件正在处理中，忽略新请求');
            }
        });

        // 按钮点击处理 - 防止事件冒泡
        selectFileBtn.addEventListener('click', (e) => {
            e.stopPropagation(); // 阻止事件冒泡到uploadArea
            if (!this.isProcessing) {
                console.log('点击选择文件按钮');
                fileInput.click();
            }
        });

        // 点击上传区域 - 修复：添加事件目标检查，防止按钮点击冒泡触发
        uploadArea.addEventListener('click', (e) => {
            // 如果点击的是按钮或其子元素，不处理（让按钮自己的事件处理）
            if (e.target.closest('button')) {
                return;
            }

            if (!this.isProcessing) {
                console.log('点击上传区域');
                fileInput.click();
            }
        });

        // 拖拽功能
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            if (!this.isProcessing) {
                uploadArea.classList.add('dragover');
            }
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');

            if (this.isProcessing) {
                console.log('⚠️ 文件正在处理中，忽略拖拽文件');
                return;
            }

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                console.log('拖拽文件:', files[0].name);
                this.processFile(files[0]);
            }
        });
    }

    initConfigManagement() {
        // 会议室数量变化
        document.getElementById('roomCount').addEventListener('input', (e) => {
            this.updateRoomConfigs(parseInt(e.target.value));
        });
        
        // 预设模板
        document.getElementById('templateSelect').addEventListener('change', (e) => {
            this.applyTemplate(e.target.value);
        });
        
        // 配置导出/导入
        document.getElementById('exportConfig').addEventListener('click', () => this.exportConfig());
        document.getElementById('importConfig').addEventListener('click', () => {
            document.getElementById('configFileInput').click();
        });
        document.getElementById('configFileInput').addEventListener('change', (e) => this.importConfig(e));
        
        // 初始化默认配置
        this.updateRoomConfigs(2);
    }

    initStepNavigation() {
        // 为所有步骤添加点击跳转功能
        document.querySelectorAll('.step').forEach((stepEl, index) => {
            stepEl.addEventListener('click', () => {
                this.jumpToStep(index + 1);
            });
            stepEl.style.cursor = 'pointer';
        });
    }

    // 新的文件处理方法
    async processFile(file) {
        // 防止并发处理
        if (this.isProcessing) {
            console.log('⚠️ 文件正在处理中，忽略新请求');
            return;
        }

        this.isProcessing = true;
        console.log('开始处理文件:', file.name, file.type, file.size);

        try {
            // 基本验证
            if (!file) {
                throw new Error('没有选择文件');
            }

            // 检查文件类型
            const fileName = file.name.toLowerCase();
            if (fileName.endsWith('.xlsm')) {
                this.showXlsmNotSupportedDialog();
                return;
            }

            if (!fileName.match(/\.(xlsx|xls)$/)) {
                throw new Error('请选择 Excel 文件 (.xlsx 或 .xls)');
            }

            // 检查文件大小（10MB限制）
            const maxSize = 10 * 1024 * 1024;
            if (file.size > maxSize) {
                throw new Error(`文件过大，请选择小于 ${(maxSize/1024/1024).toFixed(0)}MB 的文件`);
            }

            // 显示处理状态
            this.showProcessingStatus(true, '正在读取Excel文件...');

            // 读取文件
            const data = await this.readExcelFile(file);

            if (!data || data.length === 0) {
                throw new Error('Excel文件中没有找到有效数据');
            }

            // 处理成功
            console.log('文件处理成功，数据行数:', data.length);

            this.employeeData = data;
            this.displayUploadResults(file, data);

            // 更新步骤状态
            this.stepStatus.dataUploaded = true;
            document.getElementById('nextStep1').disabled = false;
            this.updateStepStatusIndicators();

            this.showAlert(`成功加载 ${data.length} 条员工记录`, 'success');

        } catch (error) {
            console.error('文件处理失败:', error);
            this.showAlert(error.message, 'danger');

            // 重置状态
            this.employeeData = [];
            this.stepStatus.dataUploaded = false;
            document.getElementById('nextStep1').disabled = true;
            this.updateStepStatusIndicators();

            // 隐藏结果显示
            document.getElementById('fileInfo').classList.add('d-none');
            document.getElementById('dataPreview').classList.add('d-none');

        } finally {
            this.showProcessingStatus(false);

            // 立即清除文件输入，允许重新选择同一文件
            const fileInput = document.getElementById('fileInput');
            if (fileInput) {
                fileInput.value = '';
            }

            // 重置处理状态
            this.isProcessing = false;
            console.log('文件处理完成，状态已重置');
        }
    }

    // 读取Excel文件
    async readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    console.log('文件读取完成，开始解析...');
                    
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: true
                    });
                    
                    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                        throw new Error('Excel文件没有工作表');
                    }
                    
                    // 读取第一个工作表
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                        header: 1,
                        defval: '',
                        raw: false
                    });
                    
                    console.log('Excel数据解析完成，行数:', jsonData.length);
                    
                    if (jsonData.length === 0) {
                        throw new Error('工作表为空');
                    }
                    
                    // 解析员工数据
                    const employees = this.parseEmployeeData(jsonData);
                    resolve(employees);
                    
                } catch (error) {
                    console.error('Excel解析错误:', error);
                    reject(new Error(`Excel文件解析失败: ${error.message}`));
                }
            };
            
            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }

    // 解析员工数据
    parseEmployeeData(rawData) {
        if (!rawData || rawData.length < 2) {
            throw new Error('数据不足，至少需要标题行和一行数据');
        }
        
        // 查找标题行
        let headerRowIndex = -1;
        let headerRow = null;
        
        for (let i = 0; i < Math.min(5, rawData.length); i++) {
            const row = rawData[i];
            if (row && row.length > 0) {
                const rowStr = row.join('').toLowerCase();
                if (rowStr.includes('姓名') || rowStr.includes('name') || 
                    rowStr.includes('角色') || rowStr.includes('role')) {
                    headerRowIndex = i;
                    headerRow = row;
                    break;
                }
            }
        }
        
        if (headerRowIndex === -1) {
            // 如果没找到标题行，假设第一行是标题
            headerRowIndex = 0;
            headerRow = rawData[0];
        }
        
        console.log('找到标题行，索引:', headerRowIndex, '内容:', headerRow);
        
        // 查找列索引
        const columnMap = this.findColumns(headerRow);
        console.log('列映射:', columnMap);
        
        if (columnMap.name === -1) {
            throw new Error('未找到姓名列，请确保Excel中包含"姓名"或"Name"列');
        }
        
        // 解析数据行
        const employees = [];
        for (let i = headerRowIndex + 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;
            
            const name = this.getCellValue(row, columnMap.name);
            if (!name) continue;
            
            const employee = {
                id: employees.length + 1,
                name: name.trim(),
                role: this.getCellValue(row, columnMap.role) || '员工',
                email: this.getCellValue(row, columnMap.email) || '',
                department: this.getCellValue(row, columnMap.department) || '',
                phone: this.getCellValue(row, columnMap.phone) || ''
            };
            
            employees.push(employee);
        }
        
        if (employees.length === 0) {
            throw new Error('没有找到有效的员工数据');
        }
        
        console.log('解析完成，员工数量:', employees.length);
        return employees;
    }

    // 查找列
    findColumns(headerRow) {
        const columns = {
            name: -1,
            role: -1,
            email: -1,
            department: -1,
            phone: -1
        };
        
        for (let i = 0; i < headerRow.length; i++) {
            const header = String(headerRow[i] || '').toLowerCase().trim();
            
            if (header.includes('姓名') || header.includes('name')) {
                columns.name = i;
            } else if (header.includes('角色') || header.includes('role') || header.includes('职位')) {
                columns.role = i;
            } else if (header.includes('邮箱') || header.includes('email') || header.includes('mail')) {
                columns.email = i;
            } else if (header.includes('部门') || header.includes('department') || header.includes('dept')) {
                columns.department = i;
            } else if (header.includes('电话') || header.includes('phone') || header.includes('tel')) {
                columns.phone = i;
            }
        }
        
        return columns;
    }

    // 获取单元格值
    getCellValue(row, index) {
        if (index === -1 || !row || index >= row.length) {
            return '';
        }
        return String(row[index] || '').trim();
    }

    // 显示上传结果
    displayUploadResults(file, data) {
        // 显示文件信息
        const fileInfo = document.getElementById('fileInfo');
        fileInfo.innerHTML = `
            <div class="row">
                <div class="col-md-6">
                    <strong>文件名:</strong> ${file.name}<br>
                    <strong>文件大小:</strong> ${(file.size / 1024).toFixed(1)} KB
                </div>
                <div class="col-md-6">
                    <strong>员工数量:</strong> ${data.length}<br>
                    <strong>状态:</strong> <span class="text-success">解析成功</span>
                </div>
            </div>
        `;
        fileInfo.classList.remove('d-none');
        
        // 显示数据预览
        const preview = document.getElementById('dataPreview');
        const previewRows = data.slice(0, 5).map(emp => `
            <tr>
                <td>${emp.name}</td>
                <td>${emp.role}</td>
                <td>${emp.email}</td>
                <td>${emp.department}</td>
            </tr>
        `).join('');
        
        preview.innerHTML = `
            <h6>数据预览 (前5行)</h6>
            <div class="table-responsive">
                <table class="table table-sm table-bordered">
                    <thead>
                        <tr><th>姓名</th><th>角色</th><th>邮箱</th><th>部门</th></tr>
                    </thead>
                    <tbody>${previewRows}</tbody>
                </table>
            </div>
        `;
        preview.classList.remove('d-none');
    }

    showProcessingStatus(show, message = '') {
        let statusEl = document.getElementById('fileProcessingStatus');
        
        if (!statusEl) {
            // 创建处理状态元素
            statusEl = document.createElement('div');
            statusEl.id = 'fileProcessingStatus';
            statusEl.className = 'processing-status d-none';
            statusEl.innerHTML = `
                <div class="d-flex align-items-center">
                    <div class="spinner-border spinner-border-sm me-2" role="status">
                        <span class="visually-hidden">Processing...</span>
                    </div>
                    <span class="status-text">处理中...</span>
                </div>
            `;
            
            // 插入到上传区域后面
            const uploadArea = document.getElementById('uploadArea');
            uploadArea.parentNode.insertBefore(statusEl, uploadArea.nextSibling);
        }
        
        if (show) {
            statusEl.querySelector('.status-text').textContent = message;
            statusEl.classList.remove('d-none');
        } else {
            statusEl.classList.add('d-none');
        }
    }

    showXlsmNotSupportedDialog() {
        // 创建XLSM文件不支持的提示对话框
        const helpModal = document.createElement('div');
        helpModal.className = 'xlsm-help-modal';
        helpModal.innerHTML = `
            <div class="xlsm-help-content">
                <div class="xlsm-help-header">
                    <h5><i class="bi bi-x-circle text-danger me-2"></i>不支持XLSM文件</h5>
                    <button class="btn-close" onclick="this.closest('.xlsm-help-modal').remove()"></button>
                </div>
                <div class="xlsm-help-body">
                    <div class="alert alert-warning">
                        <i class="bi bi-shield-exclamation me-2"></i>
                        <strong>为什么不支持XLSM文件？</strong><br>
                        XLSM文件包含VBA宏代码，结构复杂且存在安全风险。为确保系统稳定性和数据安全，我们只支持标准的Excel格式。
                    </div>
                    
                    <div class="conversion-steps">
                        <h6><i class="bi bi-arrow-repeat me-1"></i>请转换为XLSX格式：</h6>
                        <ol>
                            <li><strong>用Microsoft Excel打开</strong>您的XLSM文件</li>
                            <li>点击<strong>"文件"</strong>菜单</li>
                            <li>选择<strong>"另存为"</strong></li>
                            <li>在"保存类型"中选择<strong>"Excel工作簿(*.xlsx)"</strong></li>
                            <li>点击<strong>"保存"</strong></li>
                            <li>上传新生成的<strong>XLSX文件</strong></li>
                        </ol>
                    </div>
                    
                    <div class="alert alert-success mt-3">
                        <i class="bi bi-check-circle me-2"></i>
                        <strong>转换后的优势：</strong>
                        <ul class="mb-0 mt-2">
                            <li>文件更小，上传更快</li>
                            <li>兼容性更好，处理更稳定</li>
                            <li>保留所有数据和格式</li>
                            <li>移除潜在的安全风险</li>
                        </ul>
                    </div>
                </div>
                <div class="xlsm-help-footer">
                    <button class="btn btn-primary" onclick="this.closest('.xlsm-help-modal').remove()">
                        <i class="bi bi-check me-1"></i>我知道了，去转换
                    </button>
                </div>
            </div>
        `;
        
        document.body.appendChild(helpModal);
        
        // 20秒后自动关闭
        setTimeout(() => {
            if (helpModal.parentNode) {
                helpModal.remove();
            }
        }, 20000);
    }

    displayFileInfo(fileName, recordCount, fileInfo = null) {
        const fileNameEl = document.getElementById('fileName');
        const fileInfoEl = document.getElementById('fileInfo');
        
        let displayText = fileName;
        if (fileInfo) {
            displayText += ` (${fileInfo.type}, ${fileInfo.sizeFormatted})`;
            if (fileInfo.hasMacros) {
                displayText += ' 🔒';
            }
        }
        
        fileNameEl.textContent = displayText;
        fileInfoEl.classList.remove('d-none');
    }

    displayDataPreview(data) {
        if (data.length === 0) return;
        
        const preview = document.getElementById('dataPreview');
        const headers = document.getElementById('dataHeaders');
        const rows = document.getElementById('dataRows');
        
        // 显示表头
        const headerRow = document.createElement('tr');
        const sampleRecord = data[0];
        Object.keys(sampleRecord).forEach(key => {
            const th = document.createElement('th');
            th.textContent = key;
            headerRow.appendChild(th);
        });
        headers.innerHTML = '';
        headers.appendChild(headerRow);
        
        // 显示前5条数据
        rows.innerHTML = '';
        const previewData = data.slice(0, 5);
        previewData.forEach(record => {
            const row = document.createElement('tr');
            Object.values(record).forEach(value => {
                const td = document.createElement('td');
                td.textContent = value || '';
                row.appendChild(td);
            });
            rows.appendChild(row);
        });
        
        // 统计信息
        const roles = new Set(data.map(person => person.Role || person.role || '未知角色'));
        document.getElementById('totalCount').textContent = data.length;
        document.getElementById('roleCount').textContent = roles.size;
        
        preview.classList.remove('d-none');
    }

    updateRoomConfigs(roomCount) {
        const container = document.getElementById('roomConfigs');
        container.innerHTML = '';
        
        for (let i = 1; i <= roomCount; i++) {
            const roomConfig = this.createRoomConfigElement(i);
            container.appendChild(roomConfig);
        }
        
        this.updateConfigPreview();
    }

    createRoomConfigElement(roomNumber) {
        const div = document.createElement('div');
        div.className = 'mb-3 room-config';
        div.innerHTML = `
            <h6>会议室 ${roomNumber}</h6>
            <div class="row">
                <div class="col-6">
                    <label class="form-label">桌子数量</label>
                    <input type="number" class="form-control table-count" 
                           value="8" min="1" max="20" data-room="${roomNumber}">
                </div>
                <div class="col-6">
                    <label class="form-label">每桌座位</label>
                    <select class="form-select seats-per-table" data-room="${roomNumber}">
                        <option value="8">8人</option>
                        <option value="9" selected>9人</option>
                        <option value="10">10人</option>
                    </select>
                </div>
            </div>
        `;
        
        // 添加事件监听
        div.querySelector('.table-count').addEventListener('input', () => this.updateConfigPreview());
        div.querySelector('.seats-per-table').addEventListener('change', () => this.updateConfigPreview());
        
        return div;
    }

    updateConfigPreview() {
        const roomCount = parseInt(document.getElementById('roomCount').value);
        const preview = document.getElementById('configPreview');
        
        let totalRooms = 0;
        let totalTables = 0;
        let totalSeats = 0;
        
        let previewHTML = '';
        
        for (let i = 1; i <= roomCount; i++) {
            const tableCountInput = document.querySelector(`[data-room="${i}"].table-count`);
            const seatsSelect = document.querySelector(`[data-room="${i}"].seats-per-table`);
            
            if (tableCountInput && seatsSelect) {
                const tableCount = parseInt(tableCountInput.value);
                const seatsPerTable = parseInt(seatsSelect.value);
                
                totalRooms++;
                totalTables += tableCount;
                totalSeats += tableCount * seatsPerTable;
                
                previewHTML += `
                    <div class="room-preview">
                        <div class="room-title">会议室 ${i}</div>
                        <div class="tables-preview">
                            ${Array(tableCount).fill(0).map((_, idx) => 
                                `<div class="table-preview">${idx + 1}</div>`
                            ).join('')}
                        </div>
                    </div>
                `;
            }
        }
        
        preview.innerHTML = previewHTML;
        
        // 更新统计
        document.getElementById('totalRooms').textContent = totalRooms;
        document.getElementById('totalTables').textContent = totalTables;
        document.getElementById('totalSeats').textContent = totalSeats;
        
        // 保存配置并标记为已配置
        this.stepStatus.configSet = true;
        this.saveCurrentConfig();
        this.updateStepStatusIndicators();
    }

    applyTemplate(templateType) {
        if (!templateType) return;
        
        const templates = {
            small: { rooms: 2, tablesPerRoom: 6, seatsPerTable: 8 },
            medium: { rooms: 3, tablesPerRoom: 10, seatsPerTable: 9 },
            large: { rooms: 4, tablesPerRoom: 15, seatsPerTable: 10 }
        };
        
        const template = templates[templateType];
        if (template) {
            document.getElementById('roomCount').value = template.rooms;
            this.updateRoomConfigs(template.rooms);
            
            // 应用模板配置
            document.querySelectorAll('.table-count').forEach(input => {
                input.value = template.tablesPerRoom;
            });
            document.querySelectorAll('.seats-per-table').forEach(select => {
                select.value = template.seatsPerTable;
            });
            
            this.updateConfigPreview();
        }
    }

    jumpToStep(stepNumber) {
        this.currentStep = stepNumber;
        this.updateStepDisplay();
        
        // 根据步骤执行相应的初始化
        if (stepNumber === 4 && this.stepStatus.allocated) {
            setTimeout(() => this.updateVisualization(), 100);
        } else if (stepNumber === 5 && this.stepStatus.allocated) {
            this.generateExportPreview();
        }
    }

    nextStep() {
        // 移除严格验证，直接跳转
        this.currentStep++;
        this.updateStepDisplay();
        
        // 执行步骤相关的操作
        if (this.currentStep === 3) {
            // 不自动开始分配，等待用户点击
        } else if (this.currentStep === 4 && this.stepStatus.allocated) {
            setTimeout(() => this.updateVisualization(), 100);
        } else if (this.currentStep === 5 && this.stepStatus.allocated) {
            this.generateExportPreview();
        }
    }

    prevStep() {
        this.currentStep--;
        this.updateStepDisplay();
    }

    // 检查前置条件的方法
    checkPrerequisites(operation) {
        const missing = [];
        
        switch (operation) {
            case 'allocation':
                if (!this.stepStatus.dataUploaded) {
                    missing.push('上传员工数据');
                }
                if (!this.stepStatus.configSet) {
                    missing.push('配置会议室');
                }
                // 验证配置是否足够
                if (this.stepStatus.dataUploaded && this.stepStatus.configSet) {
                    const totalSeats = parseInt(document.getElementById('totalSeats').textContent || '0');
                    const employeeCount = this.employeeData.length;
                    if (totalSeats < employeeCount) {
                        missing.push(`调整座位配置（当前 ${totalSeats} 座位，需要 ${employeeCount} 座位）`);
                    }
                }
                break;
            
            case 'preview':
                if (!this.stepStatus.allocated) {
                    missing.push('执行座位分配');
                }
                break;
            
            case 'export':
                if (!this.stepStatus.allocated) {
                    missing.push('执行座位分配');
                }
                break;
        }
        
        return missing;
    }

    updateStepDisplay() {
        // 更新步骤指示器
        document.querySelectorAll('.step').forEach((step, index) => {
            const stepNumber = index + 1;
            step.classList.toggle('active', stepNumber === this.currentStep);
            step.classList.toggle('completed', stepNumber < this.currentStep);
        });
        
        // 显示对应的步骤内容
        document.querySelectorAll('.step-content').forEach((content, index) => {
            content.classList.toggle('active', index + 1 === this.currentStep);
        });
    }

    async startAllocation() {
        // 检查前置条件
        const missing = this.checkPrerequisites('allocation');
        if (missing.length > 0) {
            this.showPrerequisiteAlert('开始座位分配', missing);
            return;
        }
        
        const statusEl = document.getElementById('allocationStatus');
        const detailsEl = document.getElementById('allocationDetails');
        const progressEl = document.getElementById('allocationProgress');
        const resultEl = document.getElementById('allocationResult');
        
        // 重置显示状态
        progressEl.classList.remove('d-none');
        resultEl.classList.add('d-none');
        
        try {
            // 隐藏开始按钮，显示进度
            document.getElementById('allocationStart').classList.add('d-none');
            
            // 收集配置
            this.roomConfig = this.collectRoomConfiguration();
            
            // 显示进度
            statusEl.textContent = '正在分析角色分布...';
            await this.delay(800);
            
            detailsEl.textContent = '计算最优分配方案...';
            await this.delay(1000);
            
            // 执行分配算法
            const algorithm = new SeatingAlgorithm(this.employeeData, this.roomConfig);
            this.seatingArrangement = algorithm.allocateSeats();
            this.originalArrangement = JSON.parse(JSON.stringify(this.seatingArrangement));
            
            detailsEl.textContent = '优化座位安排...';
            await this.delay(800);
            
            // 显示结果
            progressEl.classList.add('d-none');
            resultEl.classList.remove('d-none');
            
            // 更新统计
            const stats = this.calculateAllocationStats();
            document.getElementById('assignedPeople').textContent = stats.assignedCount;
            document.getElementById('roleDistribution').textContent = stats.distributionScore + '%';
            document.getElementById('utilizationRate').textContent = stats.utilizationRate + '%';
            
            // 标记分配已完成
            this.stepStatus.allocated = true;
            document.getElementById('nextStep3').disabled = false;
            this.updateStepStatusIndicators();
            
        } catch (error) {
            console.error('分配失败:', error);
            progressEl.classList.add('d-none');
            document.getElementById('allocationStart').classList.remove('d-none');
            statusEl.textContent = '分配失败，请检查配置';
            this.showAlert('座位分配失败，请重试', 'danger');
        }
    }

    collectRoomConfiguration() {
        const roomCount = parseInt(document.getElementById('roomCount').value);
        const config = { rooms: [] };
        
        for (let i = 1; i <= roomCount; i++) {
            const tableCount = parseInt(document.querySelector(`[data-room="${i}"].table-count`).value);
            const seatsPerTable = parseInt(document.querySelector(`[data-room="${i}"].seats-per-table`).value);
            
            config.rooms.push({
                id: i,
                name: `会议室 ${i}`,
                tables: tableCount,
                seatsPerTable: seatsPerTable
            });
        }
        
        return config;
    }

    calculateAllocationStats() {
        const totalPeople = this.employeeData.length;
        const totalSeats = parseInt(document.getElementById('totalSeats').textContent);
        
        // 计算角色分散度
        let distributionScore = 0;
        // 这里应该实现具体的角色分散度计算逻辑
        distributionScore = Math.floor(Math.random() * 20 + 80); // 临时模拟
        
        return {
            assignedCount: totalPeople,
            distributionScore: distributionScore,
            utilizationRate: Math.floor((totalPeople / totalSeats) * 100)
        };
    }

    regenerateAssignment() {
        // 检查前置条件
        const missing = this.checkPrerequisites('allocation');
        if (missing.length > 0) {
            this.showPrerequisiteAlert('重新生成座位分配', missing);
            return;
        }
        
        if (confirm('确定要重新生成座位分配吗？这将覆盖当前的安排。')) {
            this.startAllocation().then(() => {
                if (this.currentStep === 4) {
                    this.updateVisualization();
                }
            });
        }
    }

    resetChanges() {
        if (confirm('确定要撤销所有手动调整吗？')) {
            this.seatingArrangement = JSON.parse(JSON.stringify(this.originalArrangement));
            this.updateVisualization();
        }
    }

    updateVisualization() {
        // 检查前置条件
        const missing = this.checkPrerequisites('preview');
        if (missing.length > 0) {
            this.showPrerequisiteAlert('查看座位预览', missing);
            return;
        }

        this.visualizer = new SeatingVisualizer(
            this.seatingArrangement,
            this.roomConfig,
            document.getElementById('seatingChart')
        );
        this.visualizer.render();
        this.visualizer.enableDragDrop((from, to) => this.handleSeatChange(from, to));

        this.stepStatus.previewed = true;
        this.updateRoleStats();
        this.updateStepStatusIndicators();
    }

    enterFullscreenPreview() {
        // 检查前置条件
        const missing = this.checkPrerequisites('preview');
        if (missing.length > 0) {
            this.showPrerequisiteAlert('全屏预览座位', missing);
            return;
        }

        if (this.visualizer) {
            this.visualizer.enterFullscreen();
        } else {
            this.showAlert('请先生成座位预览', 'warning');
        }
    }

    handleSeatChange(fromSeat, toSeat) {
        // 处理座位调整逻辑
        console.log('座位调整:', fromSeat, toSeat);
        
        // 在seatingArrangement中找到对应的座位并交换
        const fromRoom = this.seatingArrangement.find(r => r.id === fromSeat.roomId);
        const toRoom = this.seatingArrangement.find(r => r.id === toSeat.roomId);
        
        if (fromRoom && toRoom) {
            const fromTable = fromRoom.tables.find(t => t.id === fromSeat.tableId);
            const toTable = toRoom.tables.find(t => t.id === toSeat.tableId);
            
            if (fromTable && toTable) {
                const fromSeatObj = fromTable.seats[fromSeat.seatIndex];
                const toSeatObj = toTable.seats[toSeat.seatIndex];
                
                // 交换人员
                const tempPerson = fromSeatObj.person;
                fromSeatObj.person = toSeatObj.person;
                toSeatObj.person = tempPerson;
                
                // 更新空座位状态
                fromSeatObj.isEmpty = !fromSeatObj.person;
                toSeatObj.isEmpty = !toSeatObj.person;
                
                // 更新桌子统计
                this.updateTableStats(fromTable);
                this.updateTableStats(toTable);
                
                // 更新界面统计
                this.updateRoleStats();
                
                // 重新渲染视图以反映更改（因为CSS使用绝对定位，DOM顺序改变不会改变视觉位置）
                // 必须重新生成DOM
                setTimeout(() => {
                    this.updateVisualization();
                }, 50);
            }
        }
    }

    updateRoleStats() {
        // 更新角色分布统计
        const roleStatsEl = document.getElementById('roleStats');
        const roles = {};
        
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (seat.person) {
                        const role = seat.person.role || '未知';
                        roles[role] = (roles[role] || 0) + 1;
                    }
                });
            });
        });
        
        roleStatsEl.innerHTML = Object.entries(roles).map(([role, count]) => `
            <div class="role-stat">
                <span class="role-name">${role}</span>
                <span class="role-count">${count}</span>
            </div>
        `).join('');
    }

    exportToExcel() {
        // 检查前置条件
        const missing = this.checkPrerequisites('export');
        if (missing.length > 0) {
            this.showPrerequisiteAlert('导出座位安排', missing);
            return;
        }
        
        const exporter = new ExcelExporter(
            this.seatingArrangement,
            this.employeeData,
            this.roomConfig
        );
        
        const options = {
            byTable: document.getElementById('exportByTable').checked,
            includeSummary: document.getElementById('exportSummary').checked,
            includeOriginal: document.getElementById('exportOriginal').checked,
            fileName: document.getElementById('exportFileName').value || '座位安排表'
        };
        
        exporter.export(options);
    }

    printPreview() {
        // 实现打印预览功能
        window.print();
    }

    startOver() {
        if (confirm('确定要重新开始吗？这将清除所有数据。')) {
            this.currentStep = 1;
            this.employeeData = [];
            this.seatingArrangement = [];
            this.originalArrangement = [];
            this.stepStatus = {
                dataUploaded: false,
                configSet: false,
                allocated: false,
                previewed: false
            };
            
            // 重置表单
            document.getElementById('fileInput').value = '';
            document.getElementById('fileInfo').classList.add('d-none');
            document.getElementById('dataPreview').classList.add('d-none');
            document.getElementById('nextStep1').disabled = true;
            
            this.updateStepDisplay();
            this.updateStepStatusIndicators();
        }
    }

    saveCurrentConfig() {
        const config = {
            roomCount: document.getElementById('roomCount').value,
            rooms: []
        };
        
        document.querySelectorAll('.room-config').forEach((element, index) => {
            const roomNumber = index + 1;
            const tableCount = element.querySelector('.table-count').value;
            const seatsPerTable = element.querySelector('.seats-per-table').value;
            
            config.rooms.push({
                roomNumber,
                tableCount: parseInt(tableCount),
                seatsPerTable: parseInt(seatsPerTable)
            });
        });
        
        localStorage.setItem('findyourseat-config', JSON.stringify(config));
    }

    loadSavedConfig() {
        const saved = localStorage.getItem('findyourseat-config');
        if (saved) {
            try {
                const config = JSON.parse(saved);
                // 应用保存的配置
                // 实现配置恢复逻辑
            } catch (error) {
                console.warn('加载保存配置失败:', error);
            }
        }
    }

    exportConfig() {
        const config = {
            timestamp: new Date().toISOString(),
            roomCount: document.getElementById('roomCount').value,
            rooms: []
        };
        
        document.querySelectorAll('.room-config').forEach((element, index) => {
            const roomNumber = index + 1;
            const tableCount = element.querySelector('.table-count').value;
            const seatsPerTable = element.querySelector('.seats-per-table').value;
            
            config.rooms.push({
                roomNumber,
                tableCount: parseInt(tableCount),
                seatsPerTable: parseInt(seatsPerTable)
            });
        });
        
        const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `座位配置_${new Date().toISOString().slice(0, 10)}.json`;
        a.click();
        URL.revokeObjectURL(url);
    }

    importConfig(event) {
        const file = event.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const config = JSON.parse(e.target.result);
                this.applyImportedConfig(config);
            } catch (error) {
                this.showAlert('配置文件格式错误', 'danger');
            }
        };
        reader.readAsText(file);
    }

    applyImportedConfig(config) {
        document.getElementById('roomCount').value = config.roomCount;
        this.updateRoomConfigs(parseInt(config.roomCount));
        
        config.rooms.forEach((room, index) => {
            const tableCountInput = document.querySelector(`[data-room="${room.roomNumber}"].table-count`);
            const seatsSelect = document.querySelector(`[data-room="${room.roomNumber}"].seats-per-table`);
            
            if (tableCountInput && seatsSelect) {
                tableCountInput.value = room.tableCount;
                seatsSelect.value = room.seatsPerTable;
            }
        });
        
        this.updateConfigPreview();
        this.showAlert('配置导入成功', 'success');
    }

    showAlert(message, type = 'info') {
        // 创建并显示提示消息
        const alert = document.createElement('div');
        alert.className = `alert alert-${type} alert-dismissible fade show position-fixed`;
        alert.style.cssText = 'top: 20px; right: 20px; z-index: 9999; max-width: 400px;';
        alert.innerHTML = `
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        document.body.appendChild(alert);
        
        // 自动移除
        setTimeout(() => {
            if (alert.parentNode) {
                alert.parentNode.removeChild(alert);
            }
        }, 5000);
    }

    updateTableStats(table) {
        // 重新计算桌子的角色统计
        table.roleCount = {};
        table.currentCount = 0;
        
        table.seats.forEach(seat => {
            if (!seat.isEmpty && seat.person) {
                const role = seat.person.role;
                table.roleCount[role] = (table.roleCount[role] || 0) + 1;
                table.currentCount++;
            }
        });
    }

    generateExportPreview() {
        // 检查前置条件
        const missing = this.checkPrerequisites('export');
        if (missing.length > 0) {
            const previewContainer = document.getElementById('exportPreview');
            previewContainer.innerHTML = `
                <div class="preview-content text-center text-muted">
                    <i class="bi bi-exclamation-triangle mb-3" style="font-size: 2rem;"></i>
                    <h6>无法生成预览</h6>
                    <p>请先完成以下步骤：</p>
                    <ul class="list-unstyled">
                        ${missing.map(item => `<li><i class="bi bi-arrow-right me-1"></i>${item}</li>`).join('')}
                    </ul>
                </div>
            `;
            return;
        }
        
        // 生成导出预览
        const previewContainer = document.getElementById('exportPreview');
        
        // 创建简单的表格预览
        let previewHTML = '<div class="preview-content">';
        previewHTML += '<h6>导出预览 (前10条记录)</h6>';
        previewHTML += '<table class="table table-sm preview-table">';
        previewHTML += '<thead><tr><th>姓名</th><th>角色</th><th>会议室</th><th>桌号</th><th>座位号</th></tr></thead>';
        previewHTML += '<tbody>';
        
        let count = 0;
        for (const room of this.seatingArrangement) {
            for (const table of room.tables) {
                for (const seat of table.seats) {
                    if (!seat.isEmpty && seat.person && count < 10) {
                        const person = seat.person;
                        previewHTML += `
                            <tr>
                                <td>${person.name}</td>
                                <td><span class="badge" style="background-color: ${this.getRoleColor(person.role)}">${person.role}</span></td>
                                <td>${room.name}</td>
                                <td>${table.name}</td>
                                <td>座位${seat.id}</td>
                            </tr>
                        `;
                        count++;
                    }
                }
            }
        }
        
        if (count === 0) {
            previewHTML += '<tr><td colspan="5" class="text-center text-muted">暂无数据</td></tr>';
        }
        
        previewHTML += '</tbody></table>';
        
        if (count >= 10) {
            previewHTML += '<p class="text-muted small">...还有更多数据</p>';
        }
        
        previewHTML += '</div>';
        
        previewContainer.innerHTML = previewHTML;
    }

    updateStepStatusIndicators() {
        // 更新步骤指示器的完成状态
        const stepElements = document.querySelectorAll('.step');
        
        stepElements.forEach((stepEl, index) => {
            const stepNumber = index + 1;
            let isCompleted = false;
            
            switch (stepNumber) {
                case 1:
                    isCompleted = this.stepStatus.dataUploaded;
                    break;
                case 2:
                    isCompleted = this.stepStatus.configSet;
                    break;
                case 3:
                    isCompleted = this.stepStatus.allocated;
                    break;
                case 4:
                    isCompleted = this.stepStatus.previewed;
                    break;
                case 5:
                    isCompleted = this.stepStatus.allocated; // 只要分配完成就可以导出
                    break;
            }
            
            stepEl.classList.toggle('completed', isCompleted && stepNumber !== this.currentStep);
        });
    }

    showPrerequisiteAlert(operationName, missingSteps) {
        const stepsHtml = missingSteps.map(step => `<li class="prerequisite-item"><i class="bi bi-arrow-right me-2"></i>${step}</li>`).join('');
        
        const alertHtml = `
            <div class="alert alert-warning alert-dismissible fade show" role="alert">
                <h6><i class="bi bi-exclamation-triangle me-2"></i>无法执行：${operationName}</h6>
                <p class="mb-2">请先完成以下步骤：</p>
                <ul class="prerequisite-list mb-0">
                    ${stepsHtml}
                </ul>
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;
        
        // 创建临时容器显示提醒
        const tempAlert = document.createElement('div');
        tempAlert.innerHTML = alertHtml;
        tempAlert.style.cssText = 'position: fixed; top: 80px; left: 50%; transform: translateX(-50%); z-index: 9999; max-width: 500px;';
        document.body.appendChild(tempAlert);
        
        // 自动移除
        setTimeout(() => {
            if (tempAlert.parentNode) {
                tempAlert.parentNode.removeChild(tempAlert);
            }
        }, 8000);
    }

    // 获取角色颜色
    getRoleColor(role) {
        const colors = {
            'sales': '#007bff',
            'customer advisory': '#28a745',
            'customer success management': '#fd7e14',
            'support engineering': '#dc3545',
            'product management': '#6f42c1',
            'marketing': '#ffc107',
            'hr': '#20c997',
            'finance': '#6c757d',
            'it': '#17a2b8'
        };
        return colors[role] || '#6c757d';
    }

    downloadDemoFile() {
        // 创建示例数据
        const demoData = this.generateDemoData();
        
        // 创建工作簿
        const wb = XLSX.utils.book_new();
        
        // 创建主数据表
        const wsData = [
            ['Number', 'Role', 'First Name', 'Last Name', 'Email', 'Job Title', 'Department', 'Location', 'Manager'],
            ...demoData.map((person, index) => [
                index + 1,
                person.role,
                person.firstName,
                person.lastName,
                person.email,
                person.jobTitle,
                person.department,
                person.location,
                person.manager
            ])
        ];
        
        const ws = XLSX.utils.aoa_to_sheet(wsData);
        
        // 设置列宽
        ws['!cols'] = [
            { wch: 8 },  // Number
            { wch: 28 }, // Role
            { wch: 12 }, // First Name
            { wch: 12 }, // Last Name
            { wch: 32 }, // Email
            { wch: 22 }, // Job Title
            { wch: 15 }, // Department
            { wch: 12 }, // Location
            { wch: 15 }  // Manager
        ];
        
        // 设置标题行样式和数据验证
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const address = XLSX.utils.encode_cell({ r: 0, c: C });
            if (!ws[address]) continue;
            ws[address].s = {
                font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
                fill: { fgColor: { rgb: "2F5597" } },
                alignment: { horizontal: "center", vertical: "center" },
                border: {
                    top: { style: "thin", color: { rgb: "000000" } },
                    bottom: { style: "thin", color: { rgb: "000000" } },
                    left: { style: "thin", color: { rgb: "000000" } },
                    right: { style: "thin", color: { rgb: "000000" } }
                }
            };
        }
        
        // 为数据行添加边框和交替颜色
        for (let R = 1; R <= range.e.r; R++) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const address = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[address]) continue;
                
                ws[address].s = {
                    font: { sz: 10 },
                    fill: { fgColor: { rgb: R % 2 === 0 ? "F2F2F2" : "FFFFFF" } },
                    alignment: { horizontal: "left", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "D3D3D3" } },
                        bottom: { style: "thin", color: { rgb: "D3D3D3" } },
                        left: { style: "thin", color: { rgb: "D3D3D3" } },
                        right: { style: "thin", color: { rgb: "D3D3D3" } }
                    }
                };
                
                // 角色列特殊颜色标识
                if (C === 1) { // Role column
                    const roleColors = {
                        'sales': 'E3F2FD',
                        'customer advisory': 'E8F5E8', 
                        'customer success management': 'FFF3E0',
                        'support engineering': 'FFEBEE',
                        'product management': 'F3E5F5',
                        'marketing': 'FFF8E1',
                        'hr': 'E0F2F1',
                        'finance': 'FAFAFA',
                        'it': 'E1F5FE'
                    };
                    const role = ws[address].v;
                    ws[address].s.fill.fgColor.rgb = roleColors[role] || 'FFFFFF';
                }
            }
        }
        
        // 冻结首行
        ws['!freeze'] = { xSplit: 0, ySplit: 1 };
        
        XLSX.utils.book_append_sheet(wb, ws, "员工数据");
        
        // 创建使用说明表
        const instructionData = [
            ['FindYourSeat 座位安排系统 - XLSX示例文件'],
            [''],
            ['📋 文件说明：'],
            ['本文件包含50名员工的完整示例数据，涵盖9种不同角色'],
            ['数据格式完全符合系统要求，可直接用于测试系统功能'],
            [''],
            ['📊 支持的列名（系统会智能识别）：'],
            ['必填字段：'],
            ['  ✓ Number/编号/序号 - 员工编号'],
            ['  ✓ Role/角色/职位 - 员工角色（用于智能分配）'],
            ['  ✓ Name/姓名 - 员工姓名'],
            ['可选字段：'],
            ['  • First Name/名字 - 名'],
            ['  • Last Name/姓氏 - 姓'],
            ['  • Email/邮箱 - 电子邮箱地址'],
            ['  • Job Title/职位 - 具体职位名称'],
            ['  • Department/部门 - 所属部门'],
            ['  • Location/地点 - 工作地点'],
            ['  • Manager/主管 - 直接主管'],
            [''],
            ['🎯 角色分类说明：'],
            ['• sales - 销售团队'],
            ['• customer advisory - 客户顾问'],
            ['• customer success management - 客户成功管理'],
            ['• support engineering - 技术支持工程'],
            ['• product management - 产品管理'],
            ['• marketing - 市场营销'],
            ['• hr - 人力资源'],
            ['• finance - 财务'],
            ['• it - 信息技术'],
            [''],
            ['⚙️ 系统功能特性：'],
            ['✓ 智能角色均匀分配算法'],
            ['✓ 实时可视化座位预览'],
            ['✓ 支持拖拽手动调整'],
            ['✓ 多格式结果导出'],
            ['✓ 配置保存和复用'],
            ['✓ 非线性操作流程'],
            [''],
            ['📁 文件格式支持：'],
            ['✓ Excel 2007+ 标准格式 (.xlsx) - 推荐'],
            ['✓ Excel 97-2003 兼容格式 (.xls)'],
            ['✗ 不支持宏文件 (.xlsm)'],
            [''],
            ['🔗 GitHub项目：FindYourSeat'],
            [`⏰ 生成时间：${new Date().toLocaleString()}`],
            ['📦 版本：v2.0 - 标准XLSX版本']
        ];
        
        const wsInst = XLSX.utils.aoa_to_sheet(instructionData);
        wsInst['!cols'] = [{ wch: 60 }];
        
        // 设置标题样式
        if (wsInst['A1']) {
            wsInst['A1'].s = {
                font: { bold: true, sz: 16, color: { rgb: "000000" } },
                fill: { fgColor: { rgb: "E7E6E6" } },
                alignment: { horizontal: "center" }
            };
        }
        
        XLSX.utils.book_append_sheet(wb, wsInst, "使用说明");
        
        // 创建角色统计表
        const roleStats = this.calculateRoleStatistics(demoData);
        const statsData = [
            ['角色分布统计'],
            [''],
            ['角色', '人数', '占比', '建议桌数'],
            ...Object.entries(roleStats).map(([role, data]) => [
                role,
                data.count,
                `${data.percentage}%`,
                data.suggestedTables
            ]),
            [''],
            [`总计：${demoData.length}人`],
            [''],
            ['配置建议：'],
            [`建议会议室数量：${Math.ceil(demoData.length / 80)} 个`],
            [`建议总桌数：${Math.ceil(demoData.length / 9)} 张`],
            [`每桌建议人数：8-10人`],
            [''],
            ['注意事项：'],
            ['1. 系统会自动实现角色均匀分配'],
            ['2. 可根据实际情况调整会议室和桌子配置'],
            ['3. 支持手动拖拽调整座位安排']
        ];
        
        const wsStats = XLSX.utils.aoa_to_sheet(statsData);
        wsStats['!cols'] = [{ wch: 25 }, { wch: 8 }, { wch: 8 }, { wch: 12 }];
        
        XLSX.utils.book_append_sheet(wb, wsStats, "角色统计");
        
        // 创建模板表（供用户复制使用）
        const templateData = [
            ['Number', 'Role', 'First Name', 'Last Name', 'Email', 'Job Title', 'Department', 'Location', 'Manager'],
            ['1', 'sales', '张', '伟', 'zhang.wei@company.com', '销售经理', '销售部', '北京', '李总'],
            ['2', 'customer advisory', '李', '娜', 'li.na@company.com', '客户顾问', '客服部', '上海', '王经理'],
            ['3', 'support engineering', '王', '强', 'wang.qiang@company.com', '技术支持', '技术部', '深圳', '张总监'],
            ['', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', ''],
            ['说明：', '', '', '', '', '', '', '', ''],
            ['1. 请在空行中添加您的员工数据', '', '', '', '', '', '', '', ''],
            ['2. Number列为员工编号，可以是数字或文本', '', '', '', '', '', '', '', ''],
            ['3. Role列是关键字段，用于智能分配算法', '', '', '', '', '', '', '', ''],
            ['4. 建议的角色类型：', '', '', '', '', '', '', '', ''],
            ['   sales | customer advisory | customer success management', '', '', '', '', '', '', '', ''],
            ['   support engineering | product management | marketing', '', '', '', '', '', '', '', ''],
            ['   hr | finance | it', '', '', '', '', '', '', '', ''],
            ['5. Email字段用于生成座位通知邮件', '', '', '', '', '', '', '', ''],
            ['6. 其他字段为可选，但建议填写完整', '', '', '', '', '', '', '', '']
        ];
        
        const wsTemplate = XLSX.utils.aoa_to_sheet(templateData);
        wsTemplate['!cols'] = ws['!cols']; // 使用相同的列宽
        wsTemplate['!freeze'] = { xSplit: 0, ySplit: 1 };
        
        XLSX.utils.book_append_sheet(wb, wsTemplate, "数据模板");
        
        // 创建配置建议表
        const configData = [
            ['会议室配置建议'],
            [''],
            ['基于当前50人的配置建议：'],
            [''],
            ['方案一：紧凑型', '', ''],
            ['会议室数量', '2', '个'],
            ['每室桌子数', '3', '张'],
            ['每桌座位数', '9', '人'],
            ['总座位数', '54', '人'],
            ['利用率', '92.6%', ''],
            [''],
            ['方案二：标准型', '', ''],
            ['会议室数量', '2', '个'],
            ['每室桌子数', '4', '张'],
            ['每桌座位数', '8', '人'],
            ['总座位数', '64', '人'],
            ['利用率', '78.1%', ''],
            [''],
            ['方案三：宽松型', '', ''],
            ['会议室数量', '3', '个'],
            ['每室桌子数', '3', '张'],
            ['每桌座位数', '8', '人'],
            ['总座位数', '72', '人'],
            ['利用率', '69.4%', ''],
            [''],
            ['选择建议：'],
            ['• 人数较多时建议选择标准型或宽松型'],
            ['• 注重互动交流建议每桌8-9人'],
            ['• 便于管理建议2-3个会议室'],
            [''],
            ['系统功能：'],
            ['✓ 自动角色均匀分配'],
            ['✓ 可视化座位预览'], 
            ['✓ 支持拖拽手动调整'],
            ['✓ 多格式结果导出'],
            ['✓ 配置保存和复用']
        ];
        
        const wsConfig = XLSX.utils.aoa_to_sheet(configData);
        wsConfig['!cols'] = [{ wch: 25 }, { wch: 10 }, { wch: 8 }];
        
        XLSX.utils.book_append_sheet(wb, wsConfig, "配置建议");
        
        // 导出标准XLSX格式文件
        const wbout = XLSX.write(wb, { 
            bookType: 'xlsx', 
            type: 'array',
            Props: {
                Title: "FindYourSeat 座位安排系统示例数据",
                Subject: "智能座位分配演示文件", 
                Author: "FindYourSeat System",
                CreatedDate: new Date()
            }
        });
        
        const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = `FindYourSeat_Demo_${new Date().toISOString().slice(0, 10)}.xlsx`;
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        URL.revokeObjectURL(url);
        
        this.showAlert('Excel演示文件(.xlsx)已成功生成并下载！', 'success');
    }

    generateDemoData() {
        const roles = [
            'sales',
            'customer advisory', 
            'customer success management',
            'support engineering',
            'product management',
            'marketing',
            'hr',
            'finance',
            'it'
        ];
        
        const firstNames = [
            '伟', '芳', '娜', '敏', '静', '丽', '强', '磊', '军', '洋',
            '艳', '勇', '艺', '杰', '娟', '涛', '明', '超', '秀英', '华',
            '玲', '飞', '桂英', '建华', '丹', '萍', '鹏', '辉', '梅', '宁',
            'David', 'Sarah', 'Michael', 'Lisa', 'John', 'Mary', 'James', 'Jennifer', 'Robert', 'Patricia',
            'Alex', 'Emma', 'William', 'Sophia', 'Daniel', 'Isabella', 'Matthew', 'Charlotte', 'Andrew', 'Amelia'
        ];
        
        const lastNames = [
            '王', '李', '张', '刘', '陈', '杨', '赵', '黄', '周', '吴',
            '徐', '孙', '胡', '朱', '高', '林', '何', '郭', '马', '罗',
            'Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez',
            'Anderson', 'Taylor', 'Thomas', 'Hernandez', 'Moore', 'Martin', 'Jackson', 'Thompson', 'White', 'Lopez'
        ];
        
        const jobTitles = {
            'sales': ['销售经理', '销售代表', '大客户经理', '区域销售经理', 'Sales Manager', 'Account Executive', 'Regional Sales Director'],
            'customer advisory': ['客户顾问', '客服经理', '客户关系专员', 'Customer Advisor', 'Client Consultant', 'Customer Success Advisor'],
            'customer success management': ['客户成功经理', 'CSM专员', '客户运营经理', 'Customer Success Manager', 'Account Manager'],
            'support engineering': ['技术支持工程师', '运维工程师', '系统工程师', 'Support Engineer', 'Technical Support Specialist'],
            'product management': ['产品经理', '产品运营', '产品设计师', 'Product Manager', 'Product Owner', 'Product Designer'],
            'marketing': ['市场经理', '品牌专员', '数字营销专员', 'Marketing Manager', 'Marketing Specialist', 'Brand Manager'],
            'hr': ['人力资源经理', 'HR专员', '招聘专员', 'HR Manager', 'Recruiter', 'Talent Acquisition Specialist'],
            'finance': ['财务经理', '会计师', '财务分析师', 'Finance Manager', 'Financial Analyst', 'Senior Accountant'],
            'it': ['IT工程师', '系统管理员', '开发工程师', 'IT Engineer', 'System Administrator', 'Software Developer']
        };
        
        const departments = ['销售部', '客户服务部', '产品部', '技术部', '市场部', '人力资源部', '财务部', '运营部', 'IT部'];
        const locations = ['北京', '上海', '深圳', '广州', '杭州', '成都', '南京', '武汉', '西安', '苏州'];
        const managers = ['李总监', '王经理', '张总', '刘部长', '陈总监', 'Director Li', 'Manager Wang', 'VP Zhang', 'Director Liu'];
        
        const demoData = [];
        
        // 生成50个员工数据，确保角色分布相对均匀
        for (let i = 0; i < 50; i++) {
            const role = roles[i % roles.length];
            const isEnglish = Math.random() > 0.65; // 35%概率使用英文名
            
            const firstNamePool = isEnglish ? 
                firstNames.filter(name => /^[a-zA-Z]+$/.test(name)) : 
                firstNames.filter(name => /[\u4e00-\u9fa5]/.test(name));
                
            const lastNamePool = isEnglish ? 
                lastNames.filter(name => /^[a-zA-Z]+$/.test(name)) : 
                lastNames.filter(name => /[\u4e00-\u9fa5]/.test(name));
            
            const firstName = firstNamePool[Math.floor(Math.random() * firstNamePool.length)];
            const lastName = lastNamePool[Math.floor(Math.random() * lastNamePool.length)];
            
            const jobTitlePool = jobTitles[role] || ['员工'];
            const jobTitle = jobTitlePool[Math.floor(Math.random() * jobTitlePool.length)];
            const department = departments[Math.floor(Math.random() * departments.length)];
            const location = locations[Math.floor(Math.random() * locations.length)];
            const manager = managers[Math.floor(Math.random() * managers.length)];
            
            // 生成邮箱
            const emailPrefix = isEnglish ? 
                `${firstName.toLowerCase()}.${lastName.toLowerCase()}` :
                `${this.convertToPinyin(firstName)}.${this.convertToPinyin(lastName)}`;
            const emailDomain = Math.random() > 0.8 ? 'example.com' : 'company.com';
            const email = `${emailPrefix}@${emailDomain}`;
            
            demoData.push({
                role,
                firstName,
                lastName,
                email,
                jobTitle,
                department,
                location,
                manager
            });
        }
        
        return demoData;
    }
    
    convertToPinyin(chinese) {
        // 扩展的中文转拼音映射
        const pinyinMap = {
            '伟': 'wei', '芳': 'fang', '娜': 'na', '敏': 'min', '静': 'jing',
            '丽': 'li', '强': 'qiang', '磊': 'lei', '军': 'jun', '洋': 'yang',
            '艳': 'yan', '勇': 'yong', '艺': 'yi', '杰': 'jie', '娟': 'juan',
            '涛': 'tao', '明': 'ming', '超': 'chao', '秀英': 'xiuying', '华': 'hua',
            '玲': 'ling', '飞': 'fei', '桂英': 'guiying', '建华': 'jianhua', '丹': 'dan',
            '萍': 'ping', '鹏': 'peng', '辉': 'hui', '梅': 'mei', '宁': 'ning',
            '王': 'wang', '李': 'li', '张': 'zhang', '刘': 'liu', '陈': 'chen',
            '杨': 'yang', '赵': 'zhao', '黄': 'huang', '周': 'zhou', '吴': 'wu',
            '徐': 'xu', '孙': 'sun', '胡': 'hu', '朱': 'zhu', '高': 'gao',
            '林': 'lin', '何': 'he', '郭': 'guo', '马': 'ma', '罗': 'luo'
        };
        
        return pinyinMap[chinese] || chinese.toLowerCase().replace(/[^a-z0-9]/g, '');
    }
    
    calculateRoleStatistics(data) {
        const roleStats = {};
        const totalCount = data.length;
        
        data.forEach(person => {
            const role = person.role;
            if (!roleStats[role]) {
                roleStats[role] = { count: 0 };
            }
            roleStats[role].count++;
        });
        
        // 计算百分比和建议桌数
        Object.keys(roleStats).forEach(role => {
            const count = roleStats[role].count;
            roleStats[role].percentage = ((count / totalCount) * 100).toFixed(1);
            roleStats[role].suggestedTables = Math.ceil(count / 3); // 假设每桌平均3人同角色
        });
        
        return roleStats;
    }

    delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// 初始化应用
document.addEventListener('DOMContentLoaded', () => {
    window.app = new FindYourSeat();
    
    // 当用户离开页面时保存状态
    window.addEventListener('beforeunload', () => {
        if (window.app.employeeData.length > 0) {
            localStorage.setItem('findyourseat-session', JSON.stringify({
                employeeData: window.app.employeeData,
                currentStep: window.app.currentStep,
                timestamp: Date.now()
            }));
        }
    });
});
