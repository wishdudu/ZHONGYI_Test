// src/components/CReportExporter/excelProcessor.js
import * as XLSX from 'xlsx'

/**
 * 解析上传Excel数据
 * @param {File} file
 * @returns {Promise<Object>} 
 */
export async function processExcelFile(file) {
    try {
        // 读取文件
        const arrayBuffer = await file.arrayBuffer()
        const data = new Uint8Array(arrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })

        // --- Sheet0 处理 (检测说明汇总表) ---
        const sheet0Name = workbook.SheetNames[0]
        const worksheet0 = workbook.Sheets[sheet0Name]
        const range0 = XLSX.utils.decode_range(worksheet0['!ref'])

        let projectName = ""
        let clientName = ""
        let clientAddress = ""

        let reportType = "委托检测"
        let sampleType = ""
        let sampleFlag = true // Flag 用于确保 sample type 只被使用一次
        let samplingDate = ""
        let testingDate = ""
        let samplingAddress = ""
        let testingAddress = ""


        // 基础信息 (项目名称, 委托单位, 委托单位地址)
        const cellB2 = worksheet0['B2']
        const cellB3 = worksheet0['B3']
        const cellB4 = worksheet0['B4']
        projectName = cellB2 && cellB2.v ? cellB2.v.toString() : ""
        clientName = cellB3 && cellB3.v ? cellB3.v.toString() : ""
        clientAddress = cellB4 && cellB4.v ? cellB4.v.toString() : ""

        // 具体信息 (检测说明 & 检测依据/仪器表格) 
        for (let r = range0.s.r; r <= range0.e.r; r++) {
            const cellA = worksheet0[XLSX.utils.encode_cell({ r, c: 0 })]
            if (cellA && cellA.v) {
                const cellValue = cellA.v.toString().trim()

                if (cellValue === "项目名称") {
                    const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })]
                    if (cellB && cellB.v && cellB.v.toString().includes("送样")) {
                        reportType = "送样检测"
                    }
                }

                if (cellValue === "样品类别" && sampleFlag) {
                    sampleFlag = false
                    const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })]
                    if (cellB && cellB.v) {
                        sampleType = cellB.v.toString().replace(/,/g, '、')
                    }
                }

                if (cellValue === "采样日期") {
                    const dates = new Set()
                    // 1. 确定日期字段占据的行数
                    let rowSpan = 1
                    for (let i = r + 1; i <= range0.e.r; i++) {
                        const nextCell = worksheet0[XLSX.utils.encode_cell({ r: i, c: 0 })]
                        if (nextCell.cellValue != "采样日期") break
                        rowSpan++
                    }
                    // 2. 提取右侧列对应行数的数据
                    for (let i = 0; i < rowSpan; i++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: r + i, c: 1 })
                        const cell = worksheet0[cellAddress]
                        if (cell && cell.v) {
                            dates.add(cell.v.toString().trim())
                        }
                    }
                    samplingDate = Array.from(dates).join("、")
                }

                if (cellValue === "检测日期") {
                    const dates = new Set()
                    let rowSpan = 1
                    for (let i = r + 1; i <= range0.e.r; i++) {
                        const nextCell = worksheet0[XLSX.utils.encode_cell({ r: i, c: 0 })]
                        if (nextCell.cellValue != "检测日期") break
                        rowSpan++
                    }
                    for (let i = 0; i < rowSpan; i++) {
                        const cellAddress = XLSX.utils.encode_cell({ r: r + i, c: 1 });
                        const cell = worksheet0[cellAddress];
                        if (cell && cell.v) {
                            dates.add(cell.v.toString().trim())
                        }
                    }
                    testingDate = Array.from(dates).join("、")
                }

                if (cellValue === "采样地址") {
                    const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })]
                    samplingAddress = cellB && cellB.v ? cellB.v.toString() : ""
                }

                if (cellValue === "检测地点") {
                    const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })]
                    testingAddress = cellB && cellB.v ? cellB.v.toString() : ""
                }

                if (cellValue === "样品类别" && !sampleFlag) {
                    let rowIndex = r + 1
                    const basisList = [] // 检测项目（B列）
                    const standardList = [] // 检测方法（D列）
                    const instrumentList = [] // 仪器设备（G列）

                    while (true) {
                        const cellB = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 1 })] // Basis
                        if (!cellB || !cellB.v) break

                        const cellD = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 3 })] // Standard
                        const cellG = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 6 })] // Instrument

                        basisList.push(cellB.v.toString())
                        standardList.push(cellD && cellD.v ? cellD.v.toString() : '')
                        instrumentList.push(cellG && cellG.v ? cellG.v.toString() : '')

                        rowIndex++
                    }

                    window.__TEMP_BASIS_LIST = basisList;
                    window.__TEMP_STANDARD_LIST = standardList;
                    window.__TEMP_INSTRUMENT_LIST = instrumentList;
                }

            }
        }

        // --- Sheet1 处理 (水质汇总表)  注：仅处理了废水--- 
        const sheet1Name = workbook.SheetNames[1]
        const worksheet1 = workbook.Sheets[sheet1Name]
        const range1 = XLSX.utils.decode_range(worksheet1['!ref'])

        let shouldIncludeTable = false
        let wastewaterColIndex = null
        let wastewaterCols = []
        let specialRows = []
        let lastColumnValues = [] // 限值列
        let cColumnValues = [] // 检测项目列
        let fColumnValues = [] // 单位列
        let reportNumber = ""; // 报告编号

        // 报告编号
        const cellG8 = worksheet1['G8'];
        if (cellG8 && cellG8.v !== undefined && cellG8.v !== null) {
            const fullReportNo = cellG8.v.toString();
            const parts = fullReportNo.split('-');
            if (parts.length > 0) {
                reportNumber = parts[0];
            }
        }

        // Collect last column values (for '样品性状')
        if (shouldIncludeTable && wastewaterCols.length > 0) {
            let row = 3
            while (true) {
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: range1.e.c }) // Last column
                const lastColCell = worksheet1[cellAddress]
                if (lastColCell && lastColCell.v !== null && lastColCell.v !== '') {
                    lastColumnValues.push(lastColCell.v.toString())
                } else {
                    lastColumnValues.push('')
                }
                row++
                const nextRowCheckAddress = XLSX.utils.encode_cell({ r: row, c: wastewaterCols[0] })
                const nextRowCheckCell = worksheet1[nextRowCheckAddress]
                if (!nextRowCheckCell || nextRowCheckCell.v === null || nextRowCheckCell.v === '') {
                    break
                }
            }
        }

        // 收集C列、F列和限值列数据（从第9行开始）
        let row = 8 // 第9行
        const lastColIndex = range1.e.c
        while (true) {
            const cellCAddress = XLSX.utils.encode_cell({ r: row, c: 2 }) // C列
            const cellC = worksheet1[cellCAddress]

            if (!cellC || cellC.v === null || cellC.v === '') { break}

            cColumnValues.push(cellC.v.toString());

            // 收集F列
            const cellFAddress = XLSX.utils.encode_cell({ r: row, c: 5 }) // F列
            const cellF = worksheet1[cellFAddress];
            fColumnValues.push(cellF && cellF.v !== null ? cellF.v.toString() : '')

            // 收集限值列（最后一列）
            const lastColCellAddress = XLSX.utils.encode_cell({ r: row, c: lastColIndex })
            const lastColCell = worksheet1[lastColCellAddress];
            lastColumnValues.push(lastColCell && lastColCell.v !== null ? lastColCell.v.toString() : '');

            row++;
        }

        // 检查第三行是否有包含"废水"的单元格
        for (let col = range1.s.c; col <= range1.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 2, c: col })
            const cell = worksheet1[cellAddress]
            if (cell && typeof cell.v === 'string' && cell.v.includes('废水')) {
                shouldIncludeTable = true
                wastewaterColIndex = col
                break
            }
        }

        // 收集所有废水列（从找到的列开始向右，直到第8行没有数据）
        if (shouldIncludeTable && wastewaterColIndex !== null) {
            let col = wastewaterColIndex
            while (col <= range1.e.c) {
                const cellAddress = XLSX.utils.encode_cell({ r: 7, c: col })
                const cell = worksheet1[cellAddress]
                if (cell && cell.v !== null && cell.v !== '') {
                    wastewaterCols.push(col)
                } else {
                    break
                }
                col++
            }
        }


        return {
            projectName,
            clientName,
            clientAddress,
            reportType,
            sampleType,
            samplingDate,
            testingDate,
            samplingAddress,
            testingAddress,
            // 废水表格数据
            shouldIncludeTable,
            wastewaterCols,
            lastColumnValues,
            cColumnValues,
            fColumnValues,
            reportNumber,
            worksheet1,
            // 检测依据/仪器表格
            basisList: window.__TEMP_BASIS_LIST || [],
            standardList: window.__TEMP_STANDARD_LIST || [],
            instrumentList: window.__TEMP_INSTRUMENT_LIST || [],
        }

    } catch (error) {
        console.error("Error processing Excel file:", error)
        throw new Error(`处理Excel文件时出错: ${error.message}`)
    }
}

