const XLSX = require('xlsx');
const Mock = require('mockjs');

// 定义表头
const headers = ['编号', '单据名称', '单据状态', '制单人', '制单机构'];

// 创建数据数组
const data = [];
// 添加表头
data.push(headers);

// 定义单据状态列表
const statusList = ['已审核', '未审核', '已驳回'];
// 起始编号
let startNumber = 'G202503130000000000';

// 生成 10 万条数据
for (let i = 1; i <= 100000; i++) {
    // 处理编号递增
    const numberPart = parseInt(startNumber.slice(9));
    const newNumber = numberPart + i;
    const paddedNumber = String(newNumber).padStart(10, '0');
    const id = `G20250313${paddedNumber}`;

    const documentName = Mock.Random.ctitle(2, 4); // 生成 2 到 4 个字符的随机中文标题作为单据名称
    const status = statusList[Math.floor(Math.random() * statusList.length)]; // 随机选择单据状态
    const creator = Mock.Random.cname(); // 生成随机的中文姓名作为制单人
    const organization = Mock.Random.ctitle(2, 6); // 生成 2 到 6 个字符的随机中文标题作为制单机构名称

    data.push([id, documentName, status, creator, organization]);
}

// 创建工作表
const worksheet = XLSX.utils.aoa_to_sheet(data);
// 定义边框样式：黑色实线边框
const borderStyle = {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
};
// 获取工作表的范围
const range = XLSX.utils.decode_range(worksheet['!ref']);
// 为每个单元格应用边框样式
for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = worksheet[cellAddress];
        if (!cell) continue;
        cell.s = cell.s || {};
        cell.s.border = borderStyle;
    }
}

// 创建工作簿
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// 保存为 Excel 文件
XLSX.writeFile(workbook, '导出10万条数据.xlsx');

console.log('Excel 文件生成成功！');