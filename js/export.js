// Excel导出功能模块
class ExcelExporter {
    constructor(seatingArrangement, employeeData, roomConfig) {
        this.seatingArrangement = seatingArrangement;
        this.employeeData = employeeData;
        this.roomConfig = roomConfig;
        this.workbook = null;
    }
    
    /**
     * 导出Excel文件
     * @param {Object} options - 导出选项
     */
    export(options = {}) {
        const {
            byTable = true,
            includeSummary = true,
            includeOriginal = true,
            fileName = '座位安排表'
        } = options;
        
        try {
            // 创建工作簿
            this.workbook = XLSX.utils.book_new();
            
            // 根据选项添加工作表
            if (byTable) {
                this.addTableGroupedSheet();
            }
            
            if (includeOriginal) {
                this.addOriginalDataSheet();
            }
            
            if (includeSummary) {
                this.addSummarySheet();
            }
            
            // 添加座位清单
            this.addSeatingListSheet();
            
            // 导出文件
            this.downloadWorkbook(fileName);
            
        } catch (error) {
            console.error('导出失败:', error);
            throw new Error(`导出Excel文件失败: ${error.message}`);
        }
    }
    
    /**
     * 添加按桌分组的工作表
     */
    addTableGroupedSheet() {
        const worksheetData = [];
        
        // 添加标题行
        worksheetData.push(['座位安排表 - 按桌分组']);
        worksheetData.push([`导出时间: ${new Date().toLocaleString()}`]);
        worksheetData.push([]); // 空行
        
        this.seatingArrangement.forEach(room => {
            // 会议室标题
            worksheetData.push([`${room.name} (${room.tables.length}张桌子)`]);
            worksheetData.push([]);
            
            room.tables.forEach(table => {
                // 桌子标题
                worksheetData.push([`${table.name} (${table.currentCount}/${table.maxSeats}人)`]);
                
                // 表头
                worksheetData.push(['座位号', '姓名', '角色', '邮箱', '职位', '部门']);
                
                // 座位数据
                table.seats.forEach(seat => {
                    if (!seat.isEmpty && seat.person) {
                        const person = seat.person;
                        worksheetData.push([
                            seat.id,
                            person.name || '',
                            person.role || '',
                            person.email || '',
                            person.jobTitle || '',
                            person.department || ''
                        ]);
                    } else {
                        // 空座位
                        worksheetData.push([
                            seat.id,
                            '空座位',
                            '',
                            '',
                            '',
                            ''
                        ]);
                    }
                });
                
                // 桌子统计
                worksheetData.push([]);
                const roleStats = Object.entries(table.roleCount);
                if (roleStats.length > 0) {
                    worksheetData.push(['角色分布:']);
                    roleStats.forEach(([role, count]) => {
                        worksheetData.push(['', `${role}: ${count}人`]);
                    });
                }
                
                worksheetData.push([]); // 桌子间隔
            });
            
            worksheetData.push([]); // 会议室间隔
        });
        
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 8 },  // 座位号
            { wch: 15 }, // 姓名
            { wch: 20 }, // 角色
            { wch: 25 }, // 邮箱
            { wch: 15 }, // 职位
            { wch: 15 }  // 部门
        ];
        
        // 添加样式
        this.applyTableGroupedStyles(worksheet, worksheetData.length);
        
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '按桌分组');
    }
    
    /**
     * 添加原始数据工作表
     */
    addOriginalDataSheet() {
        const worksheetData = [];
        
        // 标题行
        worksheetData.push(['原始员工数据']);
        worksheetData.push([`总计: ${this.employeeData.length}人`]);
        worksheetData.push([]);
        
        // 表头
        const headers = ['编号', '姓名', '角色', '邮箱', '职位', '部门'];
        worksheetData.push(headers);
        
        // 数据行
        this.employeeData.forEach((person, index) => {
            worksheetData.push([
                person.number || index + 1,
                person.name || '',
                person.role || '',
                person.email || '',
                person.jobTitle || '',
                person.department || ''
            ]);
        });
        
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 8 },  // 编号
            { wch: 15 }, // 姓名
            { wch: 20 }, // 角色
            { wch: 25 }, // 邮箱
            { wch: 15 }, // 职位
            { wch: 15 }  // 部门
        ];
        
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '原始数据');
    }
    
    /**
     * 添加统计汇总工作表
     */
    addSummarySheet() {
        const worksheetData = [];
        const stats = this.calculateStatistics();
        
        // 标题
        worksheetData.push(['座位分配统计汇总']);
        worksheetData.push([`生成时间: ${new Date().toLocaleString()}`]);
        worksheetData.push([]);
        
        // 基本统计
        worksheetData.push(['基本信息']);
        worksheetData.push(['员工总数', stats.totalEmployees]);
        worksheetData.push(['座位总数', stats.totalSeats]);
        worksheetData.push(['已分配座位', stats.occupiedSeats]);
        worksheetData.push(['空余座位', stats.emptySeats]);
        worksheetData.push(['座位利用率', `${stats.utilizationRate}%`]);
        worksheetData.push([]);
        
        // 会议室统计
        worksheetData.push(['会议室统计']);
        worksheetData.push(['会议室', '桌子数', '总座位', '已占用', '利用率']);
        stats.roomStats.forEach(room => {
            const rate = ((room.occupiedSeats / room.totalSeats) * 100).toFixed(1);
            worksheetData.push([
                room.name,
                room.tables,
                room.totalSeats,
                room.occupiedSeats,
                `${rate}%`
            ]);
        });
        worksheetData.push([]);
        
        // 角色分布
        worksheetData.push(['角色分布统计']);
        worksheetData.push(['角色', '人数', '占比']);
        Object.entries(stats.roleDistribution).forEach(([role, count]) => {
            const percentage = ((count / stats.totalEmployees) * 100).toFixed(1);
            worksheetData.push([role, count, `${percentage}%`]);
        });
        worksheetData.push([]);
        
        // 桌子详细统计
        worksheetData.push(['桌子详细统计']);
        worksheetData.push(['会议室', '桌子', '容量', '人数', '利用率', '角色多样性', '角色分布']);
        
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                const rate = ((table.currentCount / table.maxSeats) * 100).toFixed(1);
                const diversity = this.calculateTableDiversity(table);
                const roleList = Object.entries(table.roleCount)
                    .map(([role, count]) => `${role}(${count})`)
                    .join(', ');
                
                worksheetData.push([
                    room.name,
                    table.name,
                    table.maxSeats,
                    table.currentCount,
                    `${rate}%`,
                    `${diversity.toFixed(1)}`,
                    roleList || '无'
                ]);
            });
        });
        
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 12 }, // 会议室/角色
            { wch: 10 }, // 数值列
            { wch: 8 },  // 数值列
            { wch: 8 },  // 数值列
            { wch: 10 }, // 百分比
            { wch: 12 }, // 多样性
            { wch: 30 }  // 角色分布
        ];
        
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '统计汇总');
    }
    
    /**
     * 添加座位清单工作表
     */
    addSeatingListSheet() {
        const worksheetData = [];
        
        // 标题
        worksheetData.push(['完整座位清单']);
        worksheetData.push([`总计: ${this.employeeData.length}人`]);
        worksheetData.push([]);
        
        // 表头
        worksheetData.push([
            '序号', '姓名', '角色', '邮箱', '会议室', '桌号', '座位号', '位置描述'
        ]);
        
        let index = 1;
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (!seat.isEmpty && seat.person) {
                        const person = seat.person;
                        worksheetData.push([
                            index++,
                            person.name || '',
                            person.role || '',
                            person.email || '',
                            room.name,
                            table.name,
                            `座位${seat.id}`,
                            this.getSeatPositionDescription(seat.position)
                        ]);
                    }
                });
            });
        });
        
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        
        // 设置列宽
        worksheet['!cols'] = [
            { wch: 6 },  // 序号
            { wch: 15 }, // 姓名
            { wch: 20 }, // 角色
            { wch: 25 }, // 邮箱
            { wch: 12 }, // 会议室
            { wch: 8 },  // 桌号
            { wch: 10 }, // 座位号
            { wch: 12 }  // 位置描述
        ];
        
        XLSX.utils.book_append_sheet(this.workbook, worksheet, '座位清单');
    }
    
    /**
     * 计算统计信息
     * @returns {Object} - 统计数据
     */
    calculateStatistics() {
        const stats = {
            totalEmployees: this.employeeData.length,
            totalSeats: 0,
            occupiedSeats: 0,
            emptySeats: 0,
            utilizationRate: 0,
            roleDistribution: {},
            roomStats: []
        };
        
        this.seatingArrangement.forEach(room => {
            const roomStat = {
                name: room.name,
                tables: room.tables.length,
                totalSeats: 0,
                occupiedSeats: 0
            };
            
            room.tables.forEach(table => {
                roomStat.totalSeats += table.maxSeats;
                roomStat.occupiedSeats += table.currentCount;
                
                // 统计角色分布
                Object.entries(table.roleCount).forEach(([role, count]) => {
                    stats.roleDistribution[role] = (stats.roleDistribution[role] || 0) + count;
                });
            });
            
            stats.roomStats.push(roomStat);
            stats.totalSeats += roomStat.totalSeats;
            stats.occupiedSeats += roomStat.occupiedSeats;
        });
        
        stats.emptySeats = stats.totalSeats - stats.occupiedSeats;
        stats.utilizationRate = ((stats.occupiedSeats / stats.totalSeats) * 100).toFixed(1);
        
        return stats;
    }
    
    /**
     * 计算桌子角色多样性
     * @param {Object} table - 桌子对象
     * @returns {number} - 多样性得分
     */
    calculateTableDiversity(table) {
        if (table.currentCount === 0) return 0;
        
        const roleCount = Object.values(table.roleCount);
        let diversity = 0;
        
        roleCount.forEach(count => {
            if (count > 0) {
                const proportion = count / table.currentCount;
                diversity -= proportion * Math.log2(proportion);
            }
        });
        
        return diversity;
    }
    
    /**
     * 获取座位位置描述
     * @param {string} position - 位置类名
     * @returns {string} - 位置描述
     */
    getSeatPositionDescription(position) {
        const positionMap = {
            'top-1': '桌子上方左',
            'top-2': '桌子上方中',
            'top-3': '桌子上方右',
            'right-1': '桌子右侧上',
            'right-2': '桌子右侧下',
            'bottom-1': '桌子下方左',
            'bottom-2': '桌子下方中',
            'bottom-3': '桌子下方右',
            'left-1': '桌子左侧上',
            'left-2': '桌子左侧下'
        };
        
        return positionMap[position] || position;
    }
    
    /**
     * 应用表格样式
     * @param {Object} worksheet - 工作表对象
     * @param {number} rowCount - 行数
     */
    applyTableGroupedStyles(worksheet, rowCount) {
        // 设置行高
        worksheet['!rows'] = [];
        for (let i = 0; i < rowCount; i++) {
            worksheet['!rows'][i] = { hpx: 20 };
        }
        
        // 设置标题行样式（这里简化处理，实际可以使用更复杂的样式）
        const range = XLSX.utils.decode_range(worksheet['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                if (!worksheet[cellAddress]) continue;
                
                // 简单的样式标记（实际使用需要完整的样式对象）
                worksheet[cellAddress].s = {
                    font: { bold: false },
                    alignment: { horizontal: 'left' }
                };
            }
        }
    }
    
    /**
     * 下载工作簿
     * @param {string} fileName - 文件名
     */
    downloadWorkbook(fileName) {
        // 生成文件名
        const timestamp = new Date().toISOString().slice(0, 10);
        const fullFileName = `${fileName}_${timestamp}.xlsx`;
        
        // 写入文件
        const wbout = XLSX.write(this.workbook, {
            bookType: 'xlsx',
            type: 'array'
        });
        
        // 创建下载链接
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = fullFileName;
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // 清理URL
        setTimeout(() => URL.revokeObjectURL(url), 1000);
    }
    
    /**
     * 生成CSV格式的座位表
     * @returns {string} - CSV内容
     */
    generateCSV() {
        const rows = [];
        
        // CSV标题行
        rows.push(['姓名', '角色', '邮箱', '会议室', '桌号', '座位号'].join(','));
        
        // 数据行
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (!seat.isEmpty && seat.person) {
                        const person = seat.person;
                        const row = [
                            `"${person.name || ''}"`,
                            `"${person.role || ''}"`,
                            `"${person.email || ''}"`,
                            `"${room.name}"`,
                            `"${table.name}"`,
                            `"座位${seat.id}"`
                        ].join(',');
                        rows.push(row);
                    }
                });
            });
        });
        
        return rows.join('\n');
    }
    
    /**
     * 下载CSV文件
     * @param {string} fileName - 文件名
     */
    downloadCSV(fileName = '座位安排表') {
        const csv = this.generateCSV();
        const timestamp = new Date().toISOString().slice(0, 10);
        const fullFileName = `${fileName}_${timestamp}.csv`;
        
        // 添加BOM以支持中文
        const BOM = '\uFEFF';
        const blob = new Blob([BOM + csv], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = fullFileName;
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        URL.revokeObjectURL(url);
    }
    
    /**
     * 生成打印友好的HTML
     * @returns {string} - HTML内容
     */
    generatePrintHTML() {
        let html = `
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <title>座位安排表</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    h1 { color: #333; text-align: center; }
                    .room { margin-bottom: 30px; page-break-inside: avoid; }
                    .room-title { font-size: 18px; font-weight: bold; margin-bottom: 10px; 
                                 background-color: #f0f0f0; padding: 10px; border-radius: 5px; }
                    .table { margin-bottom: 20px; border: 1px solid #ddd; border-radius: 5px; }
                    .table-title { background-color: #007bff; color: white; padding: 8px; 
                                  font-weight: bold; border-radius: 5px 5px 0 0; }
                    .seats { padding: 10px; }
                    .seat-row { margin-bottom: 5px; }
                    .seat-number { display: inline-block; width: 60px; font-weight: bold; }
                    .person-name { display: inline-block; width: 120px; }
                    .person-role { display: inline-block; width: 150px; color: #666; }
                    .empty-seat { color: #999; font-style: italic; }
                    @media print {
                        body { margin: 0; }
                        .room { page-break-after: always; }
                        .room:last-child { page-break-after: auto; }
                    }
                </style>
            </head>
            <body>
                <h1>座位安排表</h1>
                <p style="text-align: center; color: #666;">
                    生成时间: ${new Date().toLocaleString()}
                </p>
        `;
        
        this.seatingArrangement.forEach(room => {
            html += `<div class="room">`;
            html += `<div class="room-title">${room.name}</div>`;
            
            room.tables.forEach(table => {
                html += `<div class="table">`;
                html += `<div class="table-title">${table.name} (${table.currentCount}/${table.maxSeats}人)</div>`;
                html += `<div class="seats">`;
                
                table.seats.forEach(seat => {
                    html += `<div class="seat-row">`;
                    html += `<span class="seat-number">座位${seat.id}:</span>`;
                    
                    if (seat.isEmpty) {
                        html += `<span class="empty-seat">空座位</span>`;
                    } else {
                        const person = seat.person;
                        html += `<span class="person-name">${person.name}</span>`;
                        html += `<span class="person-role">${person.role}</span>`;
                    }
                    
                    html += `</div>`;
                });
                
                html += `</div></div>`;
            });
            
            html += `</div>`;
        });
        
        html += `</body></html>`;
        return html;
    }
    
    /**
     * 打印预览
     */
    printPreview() {
        const printHTML = this.generatePrintHTML();
        const printWindow = window.open('', '_blank');
        printWindow.document.write(printHTML);
        printWindow.document.close();
        printWindow.focus();
        
        // 等待内容加载完成后打印
        setTimeout(() => {
            printWindow.print();
        }, 500);
    }
    
    /**
     * 生成邮件模板
     * @returns {Array} - 邮件内容数组
     */
    generateEmailTemplates() {
        const templates = [];
        
        this.seatingArrangement.forEach(room => {
            room.tables.forEach(table => {
                table.seats.forEach(seat => {
                    if (!seat.isEmpty && seat.person) {
                        const person = seat.person;
                        const template = {
                            to: person.email,
                            subject: '培训座位安排通知',
                            body: this.generateEmailBody(person, room, table, seat)
                        };
                        templates.push(template);
                    }
                });
            });
        });
        
        return templates;
    }
    
    /**
     * 生成邮件正文
     * @param {Object} person - 人员信息
     * @param {Object} room - 会议室信息
     * @param {Object} table - 桌子信息
     * @param {Object} seat - 座位信息
     * @returns {string} - 邮件正文
     */
    generateEmailBody(person, room, table, seat) {
        return `
亲爱的 ${person.name}，

您好！

我们即将举行的培训活动座位已安排完毕，您的座位信息如下：

会议室：${room.name}
桌子：${table.name}
座位：座位${seat.id} (${this.getSeatPositionDescription(seat.position)})

请准时参加培训，如有任何问题，请及时联系我们。

谢谢！

培训组织团队
        `.trim();
    }
}

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelExporter;
}
