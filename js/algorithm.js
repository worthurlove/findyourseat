// 智能座位分配算法
class SeatingAlgorithm {
    constructor(employees, roomConfig) {
        this.employees = employees;
        this.roomConfig = roomConfig;
        this.seatingArrangement = [];
        this.roleStats = {};
        this.totalSeats = 0;
        
        this.analyzeData();
    }
    
    /**
     * 分析员工数据和会议室配置
     */
    analyzeData() {
        // 统计角色分布
        this.employees.forEach(emp => {
            const role = emp.role || '未知角色';
            this.roleStats[role] = (this.roleStats[role] || 0) + 1;
        });
        
        // 计算总座位数
        this.roomConfig.rooms.forEach(room => {
            this.totalSeats += room.tables * room.seatsPerTable;
        });
        
        console.log('数据分析完成:', {
            员工总数: this.employees.length,
            角色分布: this.roleStats,
            总座位数: this.totalSeats
        });
    }
    
    /**
     * 执行座位分配
     * @returns {Array} 座位分配结果
     */
    allocateSeats() {
        // 初始化会议室结构
        this.initializeRoomStructure();
        
        // 计算理想的角色分布
        const idealDistribution = this.calculateIdealDistribution();
        
        // 执行分配算法
        this.distributeEmployees(idealDistribution);
        
        // 优化分配结果
        this.optimizeDistribution();
        
        return this.seatingArrangement;
    }
    
    /**
     * 初始化会议室结构
     */
    initializeRoomStructure() {
        this.seatingArrangement = this.roomConfig.rooms.map(room => ({
            id: room.id,
            name: room.name,
            tables: Array(room.tables).fill(0).map((_, tableIndex) => ({
                id: tableIndex + 1,
                name: `桌${tableIndex + 1}`,
                maxSeats: room.seatsPerTable,
                seats: Array(room.seatsPerTable).fill(0).map((_, seatIndex) => ({
                    id: seatIndex + 1,
                    position: this.getSeatPosition(seatIndex, room.seatsPerTable),
                    person: null,
                    isEmpty: true
                })),
                roleCount: {},
                currentCount: 0
            }))
        }));
    }
    
    /**
     * 获取座位位置样式
     * @param {number} seatIndex - 座位索引
     * @param {number} totalSeats - 桌子总座位数
     * @returns {string} - 位置样式类名
     */
    getSeatPosition(seatIndex, totalSeats) {
        const positions = {
            8: ['top-1', 'top-2', 'right-1', 'right-2', 'bottom-2', 'bottom-1', 'left-2', 'left-1'],
            9: ['top-1', 'top-2', 'top-3', 'right-1', 'right-2', 'bottom-2', 'bottom-1', 'left-2', 'left-1'],
            10: ['top-1', 'top-2', 'top-3', 'right-1', 'right-2', 'bottom-3', 'bottom-2', 'bottom-1', 'left-2', 'left-1']
        };
        
        return positions[totalSeats] ? positions[totalSeats][seatIndex] : `seat-${seatIndex + 1}`;
    }
    
    /**
     * 计算理想的角色分布
     * @returns {Object} 理想分布配置
     */
    calculateIdealDistribution() {
        const roles = Object.keys(this.roleStats);
        const totalEmployees = this.employees.length;
        const totalTables = this.seatingArrangement.reduce((sum, room) => sum + room.tables.length, 0);
        
        // 计算每个角色在每桌的理想人数
        const idealDistribution = {};
        roles.forEach(role => {
            const roleCount = this.roleStats[role];
            const idealPerTable = Math.floor(roleCount / totalTables);
            const remainder = roleCount % totalTables;
            
            idealDistribution[role] = {
                perTable: idealPerTable,
                remainder: remainder,
                total: roleCount
            };
        });
        
        return idealDistribution;
    }
    
    /**
     * 分配员工到座位
     * @param {Object} idealDistribution - 理想分布
     */
    distributeEmployees(idealDistribution) {
        // 创建员工副本并随机打乱
        const shuffledEmployees = [...this.employees];
        this.shuffleArray(shuffledEmployees);
        
        // 按角色分组
        const employeesByRole = this.groupByRole(shuffledEmployees);
        
        // 获取所有桌子的引用
        const allTables = [];
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                allTables.push({ room: room.id, table: table.id, ref: table });
            });
        });
        
        // 第一阶段：均匀分配每个角色
        Object.keys(employeesByRole).forEach(role => {
            const employees = employeesByRole[role];
            const distribution = idealDistribution[role];
            
            this.distributeRoleEmployees(employees, allTables, distribution);
        });
        
        // 第二阶段：填充剩余座位
        this.fillRemainingSeats(allTables);
    }
    
    /**
     * 分配特定角色的员工
     * @param {Array} employees - 该角色的员工列表
     * @param {Array} tables - 所有桌子列表
     * @param {Object} distribution - 该角色的分布配置
     */
    distributeRoleEmployees(employees, tables, distribution) {
        let employeeIndex = 0;
        
        // 先给每桌分配基础人数
        tables.forEach(tableInfo => {
            const table = tableInfo.ref;
            const targetCount = Math.min(distribution.perTable, employees.length - employeeIndex);
            
            for (let i = 0; i < targetCount && employeeIndex < employees.length; i++) {
                const employee = employees[employeeIndex];
                const seat = table.seats.find(s => s.isEmpty);
                
                if (seat) {
                    this.assignEmployeeToSeat(employee, seat, table);
                    employeeIndex++;
                }
            }
        });
        
        // 分配剩余人员到合适的桌子
        while (employeeIndex < employees.length) {
            const employee = employees[employeeIndex];
            const bestTable = this.findBestTableForEmployee(employee, tables);
            
            if (bestTable) {
                const seat = bestTable.ref.seats.find(s => s.isEmpty);
                if (seat) {
                    this.assignEmployeeToSeat(employee, seat, bestTable.ref);
                    employeeIndex++;
                } else {
                    break; // 没有可用座位
                }
            } else {
                break; // 找不到合适的桌子
            }
        }
    }
    
    /**
     * 为员工找到最佳桌子
     * @param {Object} employee - 员工对象
     * @param {Array} tables - 桌子列表
     * @returns {Object|null} - 最佳桌子
     */
    findBestTableForEmployee(employee, tables) {
        const role = employee.role;
        
        // 筛选有空位且角色分布较好的桌子
        const availableTables = tables.filter(tableInfo => {
            const table = tableInfo.ref;
            const hasEmptySeat = table.seats.some(s => s.isEmpty);
            const currentRoleCount = table.roleCount[role] || 0;
            const maxRolePerTable = Math.ceil(table.maxSeats * 0.4); // 最多40%同角色
            
            return hasEmptySeat && currentRoleCount < maxRolePerTable;
        });
        
        if (availableTables.length === 0) {
            // 如果没有理想桌子，选择任意有空位的桌子
            return tables.find(tableInfo => tableInfo.ref.seats.some(s => s.isEmpty));
        }
        
        // 按角色多样性排序，选择角色最分散的桌子
        availableTables.sort((a, b) => {
            const diversityA = Object.keys(a.ref.roleCount).length;
            const diversityB = Object.keys(b.ref.roleCount).length;
            const countA = a.ref.currentCount;
            const countB = b.ref.currentCount;
            
            // 优先选择角色多样性高且人数较少的桌子
            if (diversityA !== diversityB) {
                return diversityB - diversityA; // 多样性高的优先
            }
            return countA - countB; // 人数少的优先
        });
        
        return availableTables[0];
    }
    
    /**
     * 分配员工到座位
     * @param {Object} employee - 员工对象
     * @param {Object} seat - 座位对象
     * @param {Object} table - 桌子对象
     */
    assignEmployeeToSeat(employee, seat, table) {
        seat.person = employee;
        seat.isEmpty = false;
        
        const role = employee.role;
        table.roleCount[role] = (table.roleCount[role] || 0) + 1;
        table.currentCount++;
        
        // 为员工添加位置信息
        employee.seatInfo = {
            roomId: table.roomId || this.findRoomIdByTable(table),
            tableId: table.id,
            seatId: seat.id,
            position: seat.position
        };
    }
    
    /**
     * 根据桌子找到所属会议室ID
     * @param {Object} table - 桌子对象
     * @returns {number} - 会议室ID
     */
    findRoomIdByTable(table) {
        for (const room of this.seatingArrangement) {
            if (room.tables.includes(table)) {
                return room.id;
            }
        }
        return 1; // 默认返回第一个会议室
    }
    
    /**
     * 填充剩余座位（如果有未分配的员工）
     * @param {Array} tables - 桌子列表
     */
    fillRemainingSeats(tables) {
        // 获取未分配的员工
        const unassignedEmployees = this.employees.filter(emp => !emp.seatInfo);
        
        unassignedEmployees.forEach(employee => {
            const availableTable = tables.find(tableInfo => 
                tableInfo.ref.seats.some(s => s.isEmpty)
            );
            
            if (availableTable) {
                const seat = availableTable.ref.seats.find(s => s.isEmpty);
                this.assignEmployeeToSeat(employee, seat, availableTable.ref);
            }
        });
    }
    
    /**
     * 优化分配结果
     */
    optimizeDistribution() {
        const maxIterations = 100;
        let improvements = 0;
        
        for (let i = 0; i < maxIterations; i++) {
            const swapMade = this.attemptBeneficialSwap();
            if (swapMade) {
                improvements++;
            } else {
                break; // 没有找到有益的交换
            }
        }
        
        console.log(`优化完成，进行了 ${improvements} 次有益的座位调整`);
    }
    
    /**
     * 尝试进行有益的座位交换
     * @returns {boolean} - 是否进行了交换
     */
    attemptBeneficialSwap() {
        const allSeats = this.getAllOccupiedSeats();
        
        for (let i = 0; i < allSeats.length - 1; i++) {
            for (let j = i + 1; j < allSeats.length; j++) {
                const seat1 = allSeats[i];
                const seat2 = allSeats[j];
                
                if (this.wouldSwapImproveDistribution(seat1, seat2)) {
                    this.swapSeats(seat1, seat2);
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * 检查交换两个座位是否能改善分布
     * @param {Object} seat1 - 座位1
     * @param {Object} seat2 - 座位2
     * @returns {boolean} - 是否改善
     */
    wouldSwapImproveDistribution(seat1, seat2) {
        if (!seat1.person || !seat2.person) return false;
        
        const table1 = this.findTableBySeat(seat1);
        const table2 = this.findTableBySeat(seat2);
        
        if (!table1 || !table2 || table1 === table2) return false;
        
        const role1 = seat1.person.role;
        const role2 = seat2.person.role;
        
        if (role1 === role2) return false; // 同角色交换无意义
        
        // 计算当前角色集中度
        const currentScore = this.calculateTableRoleScore(table1) + this.calculateTableRoleScore(table2);
        
        // 模拟交换后的得分
        const table1Role1Count = table1.roleCount[role1] || 0;
        const table1Role2Count = table1.roleCount[role2] || 0;
        const table2Role1Count = table2.roleCount[role1] || 0;
        const table2Role2Count = table2.roleCount[role2] || 0;
        
        // 交换后的角色计数
        const newTable1Score = this.calculateScoreWithChanges(table1, role1, -1, role2, 1);
        const newTable2Score = this.calculateScoreWithChanges(table2, role1, 1, role2, -1);
        
        const newScore = newTable1Score + newTable2Score;
        
        return newScore > currentScore; // 得分越高表示分布越均匀
    }
    
    /**
     * 计算桌子的角色分布得分（越高越好）
     * @param {Object} table - 桌子对象
     * @returns {number} - 得分
     */
    calculateTableRoleScore(table) {
        const roleCount = Object.values(table.roleCount);
        const totalPeople = table.currentCount;
        
        if (totalPeople === 0) return 0;
        
        // 使用Shannon多样性指数计算
        let diversity = 0;
        roleCount.forEach(count => {
            if (count > 0) {
                const proportion = count / totalPeople;
                diversity -= proportion * Math.log2(proportion);
            }
        });
        
        return diversity * 100; // 放大得分便于比较
    }
    
    /**
     * 计算变更后的得分
     * @param {Object} table - 桌子对象
     * @param {string} role1 - 角色1
     * @param {number} change1 - 角色1变化
     * @param {string} role2 - 角色2
     * @param {number} change2 - 角色2变化
     * @returns {number} - 新得分
     */
    calculateScoreWithChanges(table, role1, change1, role2, change2) {
        const tempRoleCount = { ...table.roleCount };
        tempRoleCount[role1] = (tempRoleCount[role1] || 0) + change1;
        tempRoleCount[role2] = (tempRoleCount[role2] || 0) + change2;
        
        const roleCount = Object.values(tempRoleCount);
        const totalPeople = table.currentCount;
        
        if (totalPeople === 0) return 0;
        
        let diversity = 0;
        roleCount.forEach(count => {
            if (count > 0) {
                const proportion = count / totalPeople;
                diversity -= proportion * Math.log2(proportion);
            }
        });
        
        return diversity * 100;
    }
    
    /**
     * 交换两个座位的人员
     * @param {Object} seat1 - 座位1
     * @param {Object} seat2 - 座位2
     */
    swapSeats(seat1, seat2) {
        const person1 = seat1.person;
        const person2 = seat2.person;
        const table1 = this.findTableBySeat(seat1);
        const table2 = this.findTableBySeat(seat2);
        
        // 更新座位
        seat1.person = person2;
        seat2.person = person1;
        
        // 更新员工位置信息
        if (person1) {
            person1.seatInfo = {
                roomId: this.findRoomIdByTable(table2),
                tableId: table2.id,
                seatId: seat2.id,
                position: seat2.position
            };
        }
        
        if (person2) {
            person2.seatInfo = {
                roomId: this.findRoomIdByTable(table1),
                tableId: table1.id,
                seatId: seat1.id,
                position: seat1.position
            };
        }
        
        // 更新桌子的角色统计
        if (person1 && person2) {
            const role1 = person1.role;
            const role2 = person2.role;
            
            // 更新table1
            table1.roleCount[role1] = (table1.roleCount[role1] || 1) - 1;
            table1.roleCount[role2] = (table1.roleCount[role2] || 0) + 1;
            
            // 更新table2
            table2.roleCount[role1] = (table2.roleCount[role1] || 0) + 1;
            table2.roleCount[role2] = (table2.roleCount[role2] || 1) - 1;
        }
    }
    
    /**
     * 获取所有已占用的座位
     * @returns {Array} - 座位数组
     */
    getAllOccupiedSeats() {
        const seats = [];
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (!seat.isEmpty && seat.person) {
                        seats.push(seat);
                    }
                });
            });
        });
        return seats;
    }
    
    /**
     * 根据座位查找所属桌子
     * @param {Object} seat - 座位对象
     * @returns {Object|null} - 桌子对象
     */
    findTableBySeat(seat) {
        for (const room of this.seatingArrangement) {
            for (const table of room.tables) {
                if (table.seats.includes(seat)) {
                    return table;
                }
            }
        }
        return null;
    }
    
    /**
     * 按角色分组员工
     * @param {Array} employees - 员工数组
     * @returns {Object} - 按角色分组的员工
     */
    groupByRole(employees) {
        const groups = {};
        employees.forEach(emp => {
            const role = emp.role || '未知角色';
            if (!groups[role]) {
                groups[role] = [];
            }
            groups[role].push(emp);
        });
        return groups;
    }
    
    /**
     * 随机打乱数组
     * @param {Array} array - 要打乱的数组
     */
    shuffleArray(array) {
        for (let i = array.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [array[i], array[j]] = [array[j], array[i]];
        }
    }
    
    /**
     * 获取分配统计信息
     * @returns {Object} - 统计信息
     */
    getStatistics() {
        const stats = {
            totalEmployees: this.employees.length,
            totalSeats: this.totalSeats,
            occupiedSeats: 0,
            utilizationRate: 0,
            roleDistribution: {},
            tableStats: []
        };
        
        this.seatingArrangement.forEach((room, roomIndex) => {
            room.tables.forEach((table, tableIndex) => {
                const tableStats = {
                    roomId: room.id,
                    roomName: room.name,
                    tableId: table.id,
                    tableName: table.name,
                    occupancy: table.currentCount,
                    maxCapacity: table.maxSeats,
                    roles: { ...table.roleCount },
                    diversity: this.calculateTableRoleScore(table)
                };
                
                stats.tableStats.push(tableStats);
                stats.occupiedSeats += table.currentCount;
                
                // 统计角色分布
                Object.entries(table.roleCount).forEach(([role, count]) => {
                    stats.roleDistribution[role] = (stats.roleDistribution[role] || 0) + count;
                });
            });
        });
        
        stats.utilizationRate = ((stats.occupiedSeats / stats.totalSeats) * 100).toFixed(1);
        
        return stats;
    }
    
    /**
     * 验证分配结果的有效性
     * @returns {Object} - 验证结果
     */
    validateAllocation() {
        const validation = {
            valid: true,
            errors: [],
            warnings: [],
            statistics: this.getStatistics()
        };
        
        // 检查是否所有员工都被分配
        const assignedEmployees = new Set();
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (seat.person) {
                        assignedEmployees.add(seat.person.id);
                    }
                });
            });
        });
        
        const unassignedCount = this.employees.length - assignedEmployees.size;
        if (unassignedCount > 0) {
            validation.errors.push(`${unassignedCount} 名员工未被分配座位`);
            validation.valid = false;
        }
        
        // 检查角色分布的均匀性
        const roleDistributionScore = this.calculateOverallDistributionScore();
        if (roleDistributionScore < 60) {
            validation.warnings.push('角色分布不够均匀，建议重新分配');
        }
        
        return validation;
    }
    
    /**
     * 计算整体角色分布得分
     * @returns {number} - 得分 (0-100)
     */
    calculateOverallDistributionScore() {
        const tableScores = [];
        
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                if (table.currentCount > 0) {
                    tableScores.push(this.calculateTableRoleScore(table));
                }
            });
        });
        
        if (tableScores.length === 0) return 0;
        
        const averageScore = tableScores.reduce((sum, score) => sum + score, 0) / tableScores.length;
        return Math.min(100, Math.max(0, averageScore));
    }
}

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SeatingAlgorithm;
}
