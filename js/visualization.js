// 座位可视化和拖拽模块
class SeatingVisualizer {
    constructor(seatingArrangement, roomConfig, container) {
        this.seatingArrangement = seatingArrangement;
        this.roomConfig = roomConfig;
        this.container = container;
        this.dragDropEnabled = false;
        this.dragDropCallback = null;
        this.sortableInstances = [];
        
        this.init();
    }
    
    init() {
        this.container.innerHTML = '';
        this.container.className = 'seating-chart';
    }
    
    /**
     * 渲染座位图表
     */
    render() {
        this.container.innerHTML = '';
        
        this.seatingArrangement.forEach(room => {
            const roomElement = this.createRoomElement(room);
            this.container.appendChild(roomElement);
        });
        
        // 添加图例
        this.addLegend();
    }
    
    /**
     * 创建会议室元素
     * @param {Object} room - 会议室对象
     * @returns {HTMLElement} - 会议室DOM元素
     */
    createRoomElement(room) {
        const roomDiv = document.createElement('div');
        roomDiv.className = 'room-container';
        roomDiv.setAttribute('data-room-id', room.id);
        
        // 会议室标题
        const roomHeader = document.createElement('div');
        roomHeader.className = 'room-header';
        roomHeader.innerHTML = `
            <h5>${room.name}</h5>
            <span class="room-stats">
                ${this.getRoomOccupancy(room)}
            </span>
        `;
        
        // 桌子容器
        const tablesContainer = document.createElement('div');
        tablesContainer.className = 'tables-container';
        
        room.tables.forEach(table => {
            const tableElement = this.createTableElement(table, room.id);
            tablesContainer.appendChild(tableElement);
        });
        
        roomDiv.appendChild(roomHeader);
        roomDiv.appendChild(tablesContainer);
        
        return roomDiv;
    }
    
    /**
     * 获取会议室占用率信息
     * @param {Object} room - 会议室对象
     * @returns {string} - 占用率字符串
     */
    getRoomOccupancy(room) {
        const totalSeats = room.tables.reduce((sum, table) => sum + table.maxSeats, 0);
        const occupiedSeats = room.tables.reduce((sum, table) => sum + table.currentCount, 0);
        const rate = ((occupiedSeats / totalSeats) * 100).toFixed(0);
        
        return `${occupiedSeats}/${totalSeats} (${rate}%)`;
    }
    
    /**
     * 创建桌子元素
     * @param {Object} table - 桌子对象
     * @param {number} roomId - 会议室ID
     * @returns {HTMLElement} - 桌子DOM元素
     */
    createTableElement(table, roomId) {
        const tableContainer = document.createElement('div');
        tableContainer.className = 'table-container';
        tableContainer.setAttribute('data-table-id', table.id);
        tableContainer.setAttribute('data-room-id', roomId);
        
        // 桌子主体
        const tableVisual = document.createElement('div');
        tableVisual.className = 'table-visual';
        
        // 桌子编号
        const tableNumber = document.createElement('div');
        tableNumber.className = 'table-number';
        tableNumber.textContent = table.name;
        
        // 座位容器
        const seatsContainer = document.createElement('div');
        seatsContainer.className = 'seats-container';
        
        // 创建座位
        table.seats.forEach(seat => {
            const seatElement = this.createSeatElement(seat, table.id, roomId);
            seatsContainer.appendChild(seatElement);
        });
        
        // 桌子信息提示
        const tableInfo = this.createTableInfoTooltip(table);
        
        tableVisual.appendChild(tableNumber);
        tableContainer.appendChild(tableVisual);
        tableContainer.appendChild(seatsContainer);
        tableContainer.appendChild(tableInfo);
        
        return tableContainer;
    }
    
    /**
     * 创建座位元素
     * @param {Object} seat - 座位对象
     * @param {number} tableId - 桌子ID
     * @param {number} roomId - 会议室ID
     * @returns {HTMLElement} - 座位DOM元素
     */
    createSeatElement(seat, tableId, roomId) {
        const seatDiv = document.createElement('div');
        seatDiv.className = `seat ${seat.position}`;
        seatDiv.setAttribute('data-seat-id', seat.id);
        seatDiv.setAttribute('data-table-id', tableId);
        seatDiv.setAttribute('data-room-id', roomId);
        
        if (seat.isEmpty) {
            seatDiv.classList.add('empty');
            seatDiv.textContent = seat.id;
        } else {
            const person = seat.person;
            seatDiv.classList.add(`role-${this.normalizeRoleName(person.role)}`);
            
            // 显示姓名的首字母或简称
            const displayText = this.getPersonDisplayText(person);
            seatDiv.textContent = displayText;
            
            // 添加角色标识
            seatDiv.setAttribute('data-role', person.role);
            seatDiv.setAttribute('data-person-id', person.id);
            
            // 添加提示信息
            seatDiv.title = this.getPersonTooltip(person);
            
            // 点击事件
            seatDiv.addEventListener('click', (e) => {
                e.stopPropagation();
                this.showPersonDetails(person, seatDiv);
            });
        }
        
        return seatDiv;
    }
    
    /**
     * 标准化角色名称用于CSS类名
     * @param {string} role - 角色名称
     * @returns {string} - 标准化的角色名称
     */
    normalizeRoleName(role) {
        return role.toLowerCase()
            .replace(/[^a-z0-9]/g, '-')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '');
    }
    
    /**
     * 获取人员显示文本
     * @param {Object} person - 人员对象
     * @returns {string} - 显示文本
     */
    getPersonDisplayText(person) {
        const name = person.name || '';
        
        // 中文姓名取姓氏
        if (/[\u4e00-\u9fa5]/.test(name)) {
            return name.charAt(0);
        }
        
        // 英文姓名取首字母
        const words = name.trim().split(/\s+/);
        if (words.length >= 2) {
            return `${words[0].charAt(0)}${words[1].charAt(0)}`.toUpperCase();
        }
        
        return name.substring(0, 2).toUpperCase();
    }
    
    /**
     * 获取人员提示信息
     * @param {Object} person - 人员对象
     * @returns {string} - 提示信息
     */
    getPersonTooltip(person) {
        return `${person.name}\n角色: ${person.role}\n邮箱: ${person.email || '无'}`;
    }
    
    /**
     * 创建桌子信息提示
     * @param {Object} table - 桌子对象
     * @returns {HTMLElement} - 提示元素
     */
    createTableInfoTooltip(table) {
        const infoDiv = document.createElement('div');
        infoDiv.className = 'table-info-tooltip d-none';
        
        const roleStats = Object.entries(table.roleCount)
            .map(([role, count]) => `${role}: ${count}人`)
            .join('<br>');
            
        infoDiv.innerHTML = `
            <div class="tooltip-content">
                <h6>${table.name}</h6>
                <p>座位: ${table.currentCount}/${table.maxSeats}</p>
                <div class="role-breakdown">
                    ${roleStats || '暂无人员'}
                </div>
            </div>
        `;
        
        return infoDiv;
    }
    
    /**
     * 显示人员详细信息
     * @param {Object} person - 人员对象
     * @param {HTMLElement} seatElement - 座位元素
     */
    showPersonDetails(person, seatElement) {
        // 移除之前的详情面板
        const existingPanel = document.querySelector('.person-detail-panel');
        if (existingPanel) {
            existingPanel.remove();
        }
        
        // 创建详情面板
        const panel = document.createElement('div');
        panel.className = 'person-detail-panel';
        panel.innerHTML = `
            <div class="panel-content">
                <div class="panel-header">
                    <h6>${person.name}</h6>
                    <button class="btn-close" onclick="this.closest('.person-detail-panel').remove()"></button>
                </div>
                <div class="panel-body">
                    <p><strong>角色:</strong> ${person.role}</p>
                    <p><strong>邮箱:</strong> ${person.email || '无'}</p>
                    <p><strong>位置:</strong> ${this.getLocationText(person.seatInfo)}</p>
                    ${person.jobTitle ? `<p><strong>职位:</strong> ${person.jobTitle}</p>` : ''}
                    ${person.department ? `<p><strong>部门:</strong> ${person.department}</p>` : ''}
                </div>
            </div>
        `;
        
        // 定位面板
        this.positionPanel(panel, seatElement);
        document.body.appendChild(panel);
        
        // 点击外部关闭
        setTimeout(() => {
            document.addEventListener('click', function closePanel(e) {
                if (!panel.contains(e.target)) {
                    panel.remove();
                    document.removeEventListener('click', closePanel);
                }
            });
        }, 100);
    }
    
    /**
     * 获取位置文本描述
     * @param {Object} seatInfo - 座位信息
     * @returns {string} - 位置文本
     */
    getLocationText(seatInfo) {
        if (!seatInfo) return '未分配';
        
        const room = this.seatingArrangement.find(r => r.id === seatInfo.roomId);
        const roomName = room ? room.name : `会议室${seatInfo.roomId}`;
        
        return `${roomName} - 桌${seatInfo.tableId} - 座位${seatInfo.seatId}`;
    }
    
    /**
     * 定位详情面板
     * @param {HTMLElement} panel - 面板元素
     * @param {HTMLElement} anchor - 锚点元素
     */
    positionPanel(panel, anchor) {
        const rect = anchor.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        
        panel.style.position = 'absolute';
        panel.style.top = `${rect.bottom + scrollTop + 5}px`;
        panel.style.left = `${rect.left + scrollLeft}px`;
        panel.style.zIndex = '1000';
    }
    
    /**
     * 启用拖拽功能
     * @param {Function} callback - 拖拽回调函数
     */
    enableDragDrop(callback) {
        this.dragDropEnabled = true;
        this.dragDropCallback = callback;
        
        // 为每个座位容器启用拖拽
        const seatsContainers = this.container.querySelectorAll('.seats-container');
        
        seatsContainers.forEach(container => {
            const sortable = new Sortable(container, {
                group: 'seats', // 允许跨容器拖拽
                animation: 150,
                ghostClass: 'seat-ghost',
                chosenClass: 'seat-chosen',
                dragClass: 'seat-dragging',
                filter: '.empty', // 排除空座位
                onStart: (evt) => {
                    this.onDragStart(evt);
                },
                onEnd: (evt) => {
                    this.onDragEnd(evt);
                }
            });
            
            this.sortableInstances.push(sortable);
        });
        
        // 添加拖拽样式
        this.addDragDropStyles();
    }
    
    /**
     * 拖拽开始事件
     * @param {Object} evt - 事件对象
     */
    onDragStart(evt) {
        const seatElement = evt.item;
        seatElement.classList.add('dragging');
        
        // 高亮可放置的区域
        document.querySelectorAll('.seat.empty').forEach(emptySeat => {
            emptySeat.classList.add('drop-target');
        });
    }
    
    /**
     * 拖拽结束事件
     * @param {Object} evt - 事件对象
     */
    onDragEnd(evt) {
        const seatElement = evt.item;
        seatElement.classList.remove('dragging');
        
        // 移除高亮
        document.querySelectorAll('.drop-target').forEach(seat => {
            seat.classList.remove('drop-target');
        });
        
        // 如果位置发生变化，触发回调
        if (evt.from !== evt.to || evt.oldIndex !== evt.newIndex) {
            const fromSeat = this.getSeatInfo(evt.from, evt.oldIndex);
            const toSeat = this.getSeatInfo(evt.to, evt.newIndex);
            
            if (this.dragDropCallback && fromSeat && toSeat) {
                this.dragDropCallback(fromSeat, toSeat);
            }
        }
    }
    
    /**
     * 获取座位信息
     * @param {HTMLElement} container - 座位容器
     * @param {number} index - 座位索引
     * @returns {Object|null} - 座位信息
     */
    getSeatInfo(container, index) {
        const tableContainer = container.closest('.table-container');
        if (!tableContainer) return null;
        
        const roomId = parseInt(tableContainer.getAttribute('data-room-id'));
        const tableId = parseInt(tableContainer.getAttribute('data-table-id'));
        
        return {
            roomId: roomId,
            tableId: tableId,
            seatIndex: index
        };
    }
    
    /**
     * 添加拖拽样式
     */
    addDragDropStyles() {
        const styleId = 'drag-drop-styles';
        if (document.getElementById(styleId)) return;
        
        const style = document.createElement('style');
        style.id = styleId;
        style.textContent = `
            .seat-ghost {
                opacity: 0.4;
            }
            
            .seat-chosen {
                box-shadow: 0 0 0 3px rgba(0, 123, 255, 0.25);
            }
            
            .seat-dragging {
                transform: rotate(5deg);
                opacity: 0.8;
            }
            
            .drop-target {
                animation: highlight 1s infinite;
                border: 2px dashed #28a745 !important;
            }
            
            @keyframes highlight {
                0%, 100% { background-color: rgba(40, 167, 69, 0.1); }
                50% { background-color: rgba(40, 167, 69, 0.3); }
            }
            
            .person-detail-panel {
                background: white;
                border: 1px solid #dee2e6;
                border-radius: 0.5rem;
                box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15);
                min-width: 250px;
                max-width: 350px;
            }
            
            .panel-header {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 0.75rem 1rem;
                border-bottom: 1px solid #dee2e6;
                background-color: #f8f9fa;
                border-radius: 0.5rem 0.5rem 0 0;
            }
            
            .panel-body {
                padding: 1rem;
            }
            
            .panel-body p {
                margin-bottom: 0.5rem;
                font-size: 0.875rem;
            }
        `;
        
        document.head.appendChild(style);
    }
    
    /**
     * 添加图例
     */
    addLegend() {
        const legend = document.createElement('div');
        legend.className = 'seating-legend';
        legend.innerHTML = `
            <div class="legend-title">
                <h6><i class="bi bi-info-circle me-2"></i>角色图例</h6>
            </div>
            <div class="legend-items">
                ${this.generateLegendItems()}
            </div>
        `;
        
        this.container.appendChild(legend);
    }
    
    /**
     * 生成图例项
     * @returns {string} - 图例HTML
     */
    generateLegendItems() {
        const roles = new Set();
        
        // 收集所有角色
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                Object.keys(table.roleCount).forEach(role => {
                    roles.add(role);
                });
            });
        });
        
        return Array.from(roles).map(role => {
            const normalizedRole = this.normalizeRoleName(role);
            const color = ExcelParser.getRoleColor(role);
            
            return `
                <div class="legend-item">
                    <div class="legend-color role-${normalizedRole}" style="background-color: ${color};"></div>
                    <span class="legend-label">${role}</span>
                </div>
            `;
        }).join('');
    }
    
    /**
     * 更新可视化
     */
    updateVisualization() {
        this.render();
        
        if (this.dragDropEnabled && this.dragDropCallback) {
            // 重新启用拖拽
            this.disableDragDrop();
            this.enableDragDrop(this.dragDropCallback);
        }
    }
    
    /**
     * 禁用拖拽功能
     */
    disableDragDrop() {
        this.sortableInstances.forEach(instance => {
            if (instance) {
                instance.destroy();
            }
        });
        this.sortableInstances = [];
        this.dragDropEnabled = false;
    }
    
    /**
     * 高亮特定角色
     * @param {string} role - 要高亮的角色
     */
    highlightRole(role) {
        // 移除之前的高亮
        document.querySelectorAll('.seat.highlighted').forEach(seat => {
            seat.classList.remove('highlighted');
        });
        
        if (role) {
            const normalizedRole = this.normalizeRoleName(role);
            document.querySelectorAll(`.seat.role-${normalizedRole}`).forEach(seat => {
                seat.classList.add('highlighted');
            });
        }
    }
    
    /**
     * 缩放视图
     * @param {number} scale - 缩放比例
     */
    scaleView(scale) {
        this.container.style.transform = `scale(${scale})`;
        this.container.style.transformOrigin = 'top left';
    }
    
    /**
     * 导出为图片
     * @returns {Promise<Blob>} - 图片Blob
     */
    async exportAsImage() {
        // 这里可以使用html2canvas等库来实现截图功能
        // 暂时返回一个占位符
        return new Promise((resolve) => {
            setTimeout(() => {
                resolve(new Blob(['图片导出功能待实现'], { type: 'text/plain' }));
            }, 100);
        });
    }
    
    /**
     * 获取座位统计信息
     * @returns {Object} - 统计信息
     */
    getVisualizationStats() {
        const stats = {
            totalSeats: 0,
            occupiedSeats: 0,
            emptySeats: 0,
            roleDistribution: {},
            roomStats: []
        };
        
        this.seatingArrangement.forEach(room => {
            const roomStat = {
                id: room.id,
                name: room.name,
                totalSeats: 0,
                occupiedSeats: 0,
                tables: room.tables.length
            };
            
            room.tables.forEach(table => {
                roomStat.totalSeats += table.maxSeats;
                roomStat.occupiedSeats += table.currentCount;
                
                Object.entries(table.roleCount).forEach(([role, count]) => {
                    stats.roleDistribution[role] = (stats.roleDistribution[role] || 0) + count;
                });
            });
            
            stats.roomStats.push(roomStat);
            stats.totalSeats += roomStat.totalSeats;
            stats.occupiedSeats += roomStat.occupiedSeats;
        });
        
        stats.emptySeats = stats.totalSeats - stats.occupiedSeats;
        
        return stats;
    }
}

// 添加CSS样式到文档
if (typeof document !== 'undefined') {
    const visualStyles = document.createElement('style');
    visualStyles.textContent = `
        .seating-legend {
            margin-top: 1rem;
            padding: 1rem;
            background-color: white;
            border-radius: 0.5rem;
            border: 1px solid #dee2e6;
        }
        
        .legend-title h6 {
            margin-bottom: 0.5rem;
            color: #495057;
        }
        
        .legend-items {
            display: flex;
            flex-wrap: wrap;
            gap: 1rem;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }
        
        .legend-color {
            width: 20px;
            height: 20px;
            border-radius: 50%;
            border: 2px solid white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .legend-label {
            font-size: 0.875rem;
            color: #495057;
        }
        
        .seat.highlighted {
            animation: glow 1.5s infinite;
            border: 3px solid #ffc107 !important;
        }
        
        @keyframes glow {
            0%, 100% { box-shadow: 0 0 5px rgba(255, 193, 7, 0.5); }
            50% { box-shadow: 0 0 20px rgba(255, 193, 7, 0.8); }
        }
    `;
    
    if (!document.getElementById('visualization-styles')) {
        visualStyles.id = 'visualization-styles';
        document.head.appendChild(visualStyles);
    }
}

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SeatingVisualizer;
}
