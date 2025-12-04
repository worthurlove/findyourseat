// Excel文件解析模块
class ExcelParser {
    /**
     * 解析Excel文件并提取员工数据
     * @param {File} file - Excel文件对象
     * @returns {Promise<Array>} - 解析后的员工数据数组
     */
    static async parseFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    // 获取第一个工作表
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // 转换为JSON格式
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: ''
                    });
                    
                    if (jsonData.length === 0) {
                        reject(new Error('Excel文件为空'));
                        return;
                    }
                    
                    // 解析数据结构
                    const parsedData = this.parseEmployeeData(jsonData);
                    resolve(parsedData);
                    
                } catch (error) {
                    reject(new Error(`解析Excel文件失败: ${error.message}`));
                }
            };
            
            reader.onerror = () => {
                reject(new Error('文件读取失败'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }
    
    /**
     * 解析员工数据，智能识别列结构
     * @param {Array} rawData - 原始数据数组
     * @returns {Array} - 标准化的员工数据
     */
    static parseEmployeeData(rawData) {
        if (rawData.length < 2) {
            throw new Error('数据格式不正确，需要至少包含标题行和数据行');
        }
        
        const headers = rawData[0];
        const dataRows = rawData.slice(1);
        
        // 智能映射列名
        const columnMapping = this.detectColumnMapping(headers);
        
        const employees = [];
        
        dataRows.forEach((row, index) => {
            // 跳过空行
            if (row.every(cell => !cell || cell.toString().trim() === '')) {
                return;
            }
            
            const employee = {
                id: this.getCellValue(row, columnMapping.id, `EMP_${index + 1}`),
                number: this.getCellValue(row, columnMapping.number, index + 1),
                role: this.getCellValue(row, columnMapping.role, '未知角色'),
                firstName: this.getCellValue(row, columnMapping.firstName, ''),
                lastName: this.getCellValue(row, columnMapping.lastName, ''),
                email: this.getCellValue(row, columnMapping.email, ''),
                jobTitle: this.getCellValue(row, columnMapping.jobTitle, ''),
                department: this.getCellValue(row, columnMapping.department, ''),
                // 原始数据保留
                rawData: {}
            };
            
            // 生成显示名称
            employee.name = this.generateDisplayName(employee);
            
            // 保存原始数据
            headers.forEach((header, idx) => {
                if (row[idx] !== undefined && row[idx] !== '') {
                    employee.rawData[header] = row[idx];
                }
            });
            
            employees.push(employee);
        });
        
        return employees.filter(emp => emp.name && emp.name.trim() !== '');
    }
    
    /**
     * 智能检测列映射关系
     * @param {Array} headers - 表头数组
     * @returns {Object} - 列映射对象
     */
    static detectColumnMapping(headers) {
        const mapping = {
            id: null,
            number: null,
            role: null,
            firstName: null,
            lastName: null,
            name: null,
            email: null,
            jobTitle: null,
            department: null
        };
        
        headers.forEach((header, index) => {
            const headerLower = header.toString().toLowerCase().trim();
            
            // ID相关
            if (this.matchesPattern(headerLower, ['id', 'userid', 'user id', 'sap user id', '员工id', '工号'])) {
                mapping.id = index;
            }
            
            // 编号相关
            if (this.matchesPattern(headerLower, ['number', 'num', '编号', '序号', '号码'])) {
                mapping.number = index;
            }
            
            // 角色相关
            if (this.matchesPattern(headerLower, ['role', '角色', '职位', '岗位', 'position', 'title'])) {
                mapping.role = index;
            }
            
            // 姓名相关
            if (this.matchesPattern(headerLower, ['first name', 'firstname', '名字', '姓名', 'name'])) {
                if (headerLower.includes('first') || headerLower.includes('名字')) {
                    mapping.firstName = index;
                } else if (headerLower.includes('last') || headerLower.includes('姓')) {
                    mapping.lastName = index;
                } else {
                    mapping.name = index;
                }
            }
            
            if (this.matchesPattern(headerLower, ['last name', 'lastname', '姓氏', '姓'])) {
                mapping.lastName = index;
            }
            
            // 邮箱相关
            if (this.matchesPattern(headerLower, ['email', 'mail', '邮箱', '电子邮件', 'e-mail'])) {
                mapping.email = index;
            }
            
            // 职位标题
            if (this.matchesPattern(headerLower, ['job title', 'jobtitle', '职位', '工作标题'])) {
                mapping.jobTitle = index;
            }
            
            // 部门相关
            if (this.matchesPattern(headerLower, ['department', 'dept', '部门', '科室'])) {
                mapping.department = index;
            }
        });
        
        return mapping;
    }
    
    /**
     * 检查字符串是否匹配模式数组中的任一项
     * @param {string} str - 要检查的字符串
     * @param {Array} patterns - 模式数组
     * @returns {boolean} - 是否匹配
     */
    static matchesPattern(str, patterns) {
        return patterns.some(pattern => 
            str.includes(pattern) || pattern.includes(str)
        );
    }
    
    /**
     * 获取单元格值，支持默认值
     * @param {Array} row - 数据行
     * @param {number|null} columnIndex - 列索引
     * @param {*} defaultValue - 默认值
     * @returns {*} - 单元格值
     */
    static getCellValue(row, columnIndex, defaultValue = '') {
        if (columnIndex === null || columnIndex === undefined || !row[columnIndex]) {
            return defaultValue;
        }
        
        const value = row[columnIndex];
        return value !== null && value !== undefined ? value.toString().trim() : defaultValue;
    }
    
    /**
     * 生成显示名称
     * @param {Object} employee - 员工对象
     * @returns {string} - 显示名称
     */
    static generateDisplayName(employee) {
        // 优先使用完整姓名
        if (employee.name && employee.name.trim()) {
            return employee.name.trim();
        }
        
        // 其次组合姓和名
        if (employee.firstName || employee.lastName) {
            const firstName = employee.firstName || '';
            const lastName = employee.lastName || '';
            return `${lastName} ${firstName}`.trim();
        }
        
        // 使用邮箱前缀
        if (employee.email && employee.email.includes('@')) {
            return employee.email.split('@')[0];
        }
        
        // 使用ID
        if (employee.id) {
            return employee.id.toString();
        }
        
        return `员工${employee.number}`;
    }
    
    /**
     * 验证数据完整性
     * @param {Array} employees - 员工数据数组
     * @returns {Object} - 验证结果
     */
    static validateData(employees) {
        const result = {
            valid: true,
            warnings: [],
            errors: [],
            statistics: {}
        };
        
        if (employees.length === 0) {
            result.valid = false;
            result.errors.push('没有找到有效的员工数据');
            return result;
        }
        
        // 统计角色分布
        const roles = {};
        const missingEmails = [];
        const duplicateIds = [];
        const seenIds = new Set();
        
        employees.forEach((emp, index) => {
            // 角色统计
            const role = emp.role || '未知角色';
            roles[role] = (roles[role] || 0) + 1;
            
            // 检查缺失邮箱
            if (!emp.email || emp.email.trim() === '') {
                missingEmails.push(`第${index + 1}行: ${emp.name}`);
            }
            
            // 检查重复ID
            if (emp.id && seenIds.has(emp.id)) {
                duplicateIds.push(emp.id);
            }
            seenIds.add(emp.id);
        });
        
        // 生成警告
        if (missingEmails.length > 0) {
            result.warnings.push(`${missingEmails.length}名员工缺少邮箱信息`);
        }
        
        if (duplicateIds.length > 0) {
            result.warnings.push(`发现重复的员工ID: ${duplicateIds.join(', ')}`);
        }
        
        // 角色分布检查
        if (Object.keys(roles).length === 1) {
            result.warnings.push('所有员工属于同一角色，可能影响分配效果');
        }
        
        result.statistics = {
            totalEmployees: employees.length,
            roleDistribution: roles,
            missingEmailCount: missingEmails.length,
            duplicateIdCount: duplicateIds.length
        };
        
        return result;
    }
    
    /**
     * 导出示例模板
     * @returns {string} - CSV格式的示例模板
     */
    static generateTemplate() {
        const templateData = [
            ['Number', 'Role', 'First Name', 'Last Name', 'Email', 'Job Title'],
            ['1', 'sales', 'Zhang', 'Wei', 'zhang.wei@company.com', 'Sales Manager'],
            ['2', 'customer advisory', 'Li', 'Ming', 'li.ming@company.com', 'Customer Advisor'],
            ['3', 'customer success management', 'Wang', 'Fang', 'wang.fang@company.com', 'Success Manager'],
            ['4', 'support engineering', 'Liu', 'Gang', 'liu.gang@company.com', 'Support Engineer']
        ];
        
        return templateData.map(row => row.join(',')).join('\n');
    }
    
    /**
     * 下载模板文件
     */
    static downloadTemplate() {
        const template = this.generateTemplate();
        const blob = new Blob([template], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        
        const link = document.createElement('a');
        link.href = url;
        link.download = '员工数据模板.csv';
        link.style.display = 'none';
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        URL.revokeObjectURL(url);
    }
    
    /**
     * 数据预处理 - 清理和标准化
     * @param {Array} employees - 员工数据数组
     * @returns {Array} - 处理后的数据
     */
    static preprocessData(employees) {
        return employees.map(emp => {
            // 角色标准化
            emp.role = this.standardizeRole(emp.role);
            
            // 名称清理
            emp.name = emp.name.replace(/[^\u4e00-\u9fa5a-zA-Z\s]/g, '').trim();
            
            // 邮箱验证和清理
            if (emp.email) {
                emp.email = emp.email.toLowerCase().trim();
                const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
                if (!emailPattern.test(emp.email)) {
                    emp.emailValid = false;
                }
            }
            
            return emp;
        });
    }
    
    /**
     * 标准化角色名称
     * @param {string} role - 原始角色名称
     * @returns {string} - 标准化后的角色名称
     */
    static standardizeRole(role) {
        if (!role || role.trim() === '') {
            return '未知角色';
        }
        
        const roleStr = role.toString().toLowerCase().trim();
        
        // 角色映射表
        const roleMappings = {
            'sales': ['sales', '销售', 'sale'],
            'customer advisory': ['customer advisory', '客户顾问', 'advisory'],
            'customer success management': ['customer success management', '客户成功', 'success', 'csm'],
            'support engineering': ['support engineering', '技术支持', 'support', 'engineer'],
            'product management': ['product management', '产品管理', 'product', 'pm'],
            'marketing': ['marketing', '市场营销', '市场'],
            'hr': ['hr', 'human resources', '人力资源', '人事'],
            'finance': ['finance', '财务', '会计'],
            'it': ['it', 'information technology', '信息技术', '技术']
        };
        
        for (const [standardRole, variations] of Object.entries(roleMappings)) {
            if (variations.some(variation => roleStr.includes(variation))) {
                return standardRole;
            }
        }
        
        return role.trim();
    }
}

// 工具函数：角色颜色映射
ExcelParser.getRoleColor = function(role) {
    const colorMap = {
        'sales': '#007bff',
        'customer advisory': '#28a745',
        'customer success management': '#ffc107',
        'support engineering': '#dc3545',
        'product management': '#6f42c1',
        'marketing': '#fd7e14',
        'hr': '#20c997',
        'finance': '#6c757d',
        'it': '#17a2b8'
    };
    
    return colorMap[role] || '#6c757d';
};

// 导出模块
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelParser;
}
