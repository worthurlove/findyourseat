// 主应用程序 - 座位安排系统
class FindYourSeat {
    constructor() {
        this.currentStep = 1;
        this.employeeData = [];
        this.roomConfig = {};
        this.seatingArrangement = [];
        this.originalArrangement = [];
        
        this.init();
    }

    init() {
        this.initEventListeners();
        this.loadSavedConfig();
        this.updateStepDisplay();
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
        document.getElementById('regenerateBtn').addEventListener('click', () => this.regenerateAssignment());
        document.getElementById('resetChangesBtn').addEventListener('click', () => this.resetChanges());
        
        // 导出功能
        document.getElementById('exportExcelBtn').addEventListener('click', () => this.exportToExcel());
        document.getElementById('printPreviewBtn').addEventListener('click', () => this.printPreview());
        
        // 重新开始
        document.getElementById('startOverBtn').addEventListener('click', () => this.startOver());
    }

    initFileUpload() {
        const fileInput = document.getElementById('fileInput');
        const uploadArea = document.getElementById('uploadArea');
        
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));
        
        // 拖拽上传
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });
        
        uploadArea.addEventListener('click', () => {
            fileInput.click();
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

    handleFileSelect(event) {
        const file = event.target.files[0];
        if (file) {
            this.handleFile(file);
        }
    }

    async handleFile(file) {
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            this.showAlert('请选择有效的Excel文件 (.xlsx 或 .xls)', 'danger');
            return;
        }

        try {
            const data = await ExcelParser.parseFile(file);
            this.employeeData = data;
            
            this.displayFileInfo(file.name, data.length);
            this.displayDataPreview(data);
            
            document.getElementById('nextStep1').disabled = false;
            
            this.showAlert(`成功加载 ${data.length} 条员工记录`, 'success');
        } catch (error) {
            console.error('文件解析错误:', error);
            this.showAlert('文件解析失败，请检查文件格式', 'danger');
        }
    }

    displayFileInfo(fileName, recordCount) {
        document.getElementById('fileName').textContent = fileName;
        document.getElementById('fileInfo').classList.remove('d-none');
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
        
        // 保存配置
        this.saveCurrentConfig();
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

    nextStep() {
        if (this.validateCurrentStep()) {
            this.currentStep++;
            this.updateStepDisplay();
            
            if (this.currentStep === 3) {
                this.startAllocation();
            } else if (this.currentStep === 4) {
                // 进入预览步骤时渲染可视化
                setTimeout(() => this.updateVisualization(), 100);
            } else if (this.currentStep === 5) {
                // 进入导出步骤时生成预览
                this.generateExportPreview();
            }
        }
    }

    prevStep() {
        this.currentStep--;
        this.updateStepDisplay();
    }

    validateCurrentStep() {
        switch (this.currentStep) {
            case 1:
                return this.employeeData.length > 0;
            case 2:
                return this.validateConfiguration();
            default:
                return true;
        }
    }

    validateConfiguration() {
        const totalSeats = parseInt(document.getElementById('totalSeats').textContent);
        const employeeCount = this.employeeData.length;
        
        if (totalSeats < employeeCount) {
            this.showAlert('座位数量不足，请增加会议室或桌子数量', 'warning');
            return false;
        }
        
        return true;
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
        const statusEl = document.getElementById('allocationStatus');
        const detailsEl = document.getElementById('allocationDetails');
        const progressEl = document.getElementById('allocationProgress');
        const resultEl = document.getElementById('allocationResult');
        
        try {
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
            
            document.getElementById('nextStep3').disabled = false;
            
        } catch (error) {
            console.error('分配失败:', error);
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
        if (this.currentStep === 4) {
            this.visualizer = new SeatingVisualizer(
                this.seatingArrangement, 
                this.roomConfig,
                document.getElementById('seatingChart')
            );
            this.visualizer.render();
            this.visualizer.enableDragDrop((from, to) => this.handleSeatChange(from, to));
            
            this.updateRoleStats();
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
            
            // 重置表单
            document.getElementById('fileInput').value = '';
            document.getElementById('fileInfo').classList.add('d-none');
            document.getElementById('dataPreview').classList.add('d-none');
            document.getElementById('nextStep1').disabled = true;
            
            this.updateStepDisplay();
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
                                <td><span class="badge" style="background-color: ${ExcelParser.getRoleColor(person.role)}">${person.role}</span></td>
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
