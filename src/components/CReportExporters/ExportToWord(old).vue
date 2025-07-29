<script setup lang="ts">
import { ElMessage, type Alignment } from 'element-plus'
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun, VerticalAlign, LineRuleType, PageBreak, SectionType, ImageRun} from 'docx'
import { saveAs } from 'file-saver'
import * as XLSX from 'xlsx'

const props = defineProps({
  originalFile: File
})

// 表格边距样式
const customTableStyle = {
  margins: {
    left: 108, // 0.19cm
    right: 108
  }
};

// --- 辅助函数：创建居中、中英文混合字体的段落 ---
const createCenteredMixedFontParagraph = (text: string, size: number) => {
  const runs: TextRun[] = [];
  for (const char of text) {
    const isAscii = /^[\x00-\x7F]$/.test(char); // 判断是否为ASCII字符
    runs.push(
      new TextRun({
        text: char,
        font: isAscii ? 'Times New Roman' : '宋体',
        size: size * 2,
        bold: false // 内容单元格通常不加粗
      })
    );
  }
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: runs
  });
};

// 设置二级标题格式
const createFormattedParagraph = (text: string, config: any) => {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text,
        bold: config.bold,
        font: config.font,
        size: config.size * 2, // 转换为半磅单位
      }),
    ],
  })
}

// 设置检测说明中的中文段落
const createParagraph0 = (text: string, font: string, bold = false) => {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: {
      line: 400, // 20磅行距 -> 400 twips
      lineRule: LineRuleType.AT_LEAST, // 最小值行距
    },
    children: [
      new TextRun({
        text,
        bold,
        font,
        size: 24, // 小四 = 12磅 -> 半磅单位24
      }),
    ],
  })
}

// 设置检测说明中的英文段落
const createParagraph1 = (text: string, font: string, bold = false) => {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 400 }, // 首行缩进0.02字符（约400 twips）
    spacing: {
      line: 400, // 20磅行距 -> 400 twips
      lineRule: LineRuleType.AT_LEAST, // 最小值行距
    },
    children: [
      new TextRun({
        text,
        bold,
        font,
        size: 24, // 小四 = 12磅 -> 半磅单位24
      }),
    ],
  })
}

// 设置表格单元格格式
type AlignmentTypeEnum = typeof AlignmentType[keyof typeof AlignmentType]
const createFormattedCell0 = (text: string, bold = false, colSpan?: number, alignment: AlignmentTypeEnum = AlignmentType.CENTER) => {
 const runs: TextRun[] = []

  for (const char of text) {
    const isAscii = /^[\x00-\x7F]$/.test(char) // 英文/数字/符号
    runs.push(
      new TextRun({
        text: char,
        bold,
        font: isAscii ? 'Times New Roman' : '宋体',
        size: 21, // 五号字
      })
    )
  }

  return new TableCell({
    margins: customTableStyle.margins,
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: colSpan, 
    children: [
      new Paragraph({
        alignment: alignment,
        children: runs,
      }),
    ],
  })
}

// 设置表格单元格格式（中英文两行）
const createFormattedCell1 = (chineseText: string, englishText: string) => {
  return new TableCell({
    margins: customTableStyle.margins,
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: chineseText,
            bold: true,
            font: '宋体',
            size: 21, 
          }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: englishText,
            font: 'Times New Roman',
            size: 21, 
          }),
        ],
      }),
    ],
  })
}

// 辅助函数：合并项目名称和单位
const mergeItemWithUnit = (itemName: string, unit: string): string => {
  if (!unit) {
    return itemName; // 如果没有单位，只返回项目名称
  }
  // 检查单位是否包含中文字符
  const hasChinese = /[\u4e00-\u9fa5]/.test(unit);
  if (hasChinese) {
    return `${itemName}（${unit}）`; // 中文单位加括号
  } else {
    return `${itemName} ${unit}`; // 英文单位用空格分隔
  }
};

// 智能合并单元格函数
const mergeLocationCells = (values: string[]) => {
  const mergedCells: { text: string; colSpan: number }[] = []
  let i = 0

  while (i < values.length) {
    const currentValue = values[i]
    
    // 如果是非空单元格
    if (currentValue && currentValue.trim() !== '') {
      let j = i + 1
      // 向后查找连续的空单元格
      while (j < values.length && (!values[j] || values[j].trim() === '')) {
        j++
      }
      // 计算需要合并的单元格数量
      const colSpan = j - i
      mergedCells.push({ text: currentValue, colSpan })
      i = j
    } 
    // 如果是空单元格
    else {
      let j = i + 1
      // 向后查找连续的空单元格
      while (j < values.length && (!values[j] || values[j].trim() === '')) {
        j++
      }
      // 合并所有连续的空单元格
      const colSpan = j - i
      mergedCells.push({ text: '', colSpan })
      i = j
    }
  }

  return mergedCells
}

const handleClick = async () => {
  if (!props.originalFile) {
    ElMessage.error('请先上传文件')
    return
  }


  try {

    // 读取Excel文件
    const arrayBuffer = await props.originalFile.arrayBuffer()
    const data = new Uint8Array(arrayBuffer)
    const workbook = XLSX.read(data, { type: 'array' })
    
    // 获取sheet1汇总数据
    const worksheet0 = workbook.Sheets[workbook.SheetNames[0]]
    const range0 = XLSX.utils.decode_range(worksheet0['!ref']!)
    
    let Type = "委托检测"; // 默认值
    let sampleType = ""
    let sampleFlag = true
    let samplingDate = "";
    let testingDate = "";
    let samplingAddress = "";
    let testingAddress = "";
    let basisList: string[] = [];   // 检测项目（B列）
    let standardList: string[] = []; // 检测方法（D列）
    let instrumentList: string[] = []; // 仪器设备（G列）
    let projectName = "";
    let clientName = "";
    let clientAddress = "";
    if (worksheet0) {
      const cellB2 = worksheet0['B2'];
      const cellB3 = worksheet0['B3'];
      const cellB4 = worksheet0['B4'];
      projectName = cellB2 && cellB2.v ? cellB2.v.toString() : "";
      clientName = cellB3 && cellB3.v ? cellB3.v.toString() : "";
      clientAddress = cellB4 && cellB4.v ? cellB4.v.toString() : "";

      for (let r = range0.s.r; r <= range0.e.r; r++) {
        // 获取第一列单元格
        const cellA = worksheet0[XLSX.utils.encode_cell({ r, c: 0 })];
        if (cellA && cellA.v) {
          const cellValue = cellA.v.toString().trim();
          
          // 处理"项目名称"
          if (cellValue === "项目名称") {
            const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })];
            if (cellB && cellB.v && cellB.v.toString().includes("送样")) {
              Type = "送样检测";
            }
          }
          // 处理"样品类别"
          else if (cellValue === "样品类别" && sampleFlag) {
            sampleFlag = false;
            const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })];
            if (cellB && cellB.v) {
              sampleType = cellB.v.toString().replace(/,/g, '、');
            }
          }
          // 处理"采样日期" - 只提取右侧一列的多行数据
          else if (cellValue === "采样日期") {
            const dates = new Set<string>();
            // 1. 确定日期字段占据的行数
            let rowSpan = 1;
            for (let i = r + 1; i <= range0.e.r; i++) {
              const nextCell = worksheet0[XLSX.utils.encode_cell({ r: i, c: 0 })];
              if (nextCell.cellValue != "采样日期") break;
              rowSpan++;
            }
            // 2. 提取右侧列对应行数的数据
            for (let i = 0; i < rowSpan; i++) {
              const cellAddress = XLSX.utils.encode_cell({ r: r + i, c: 1 });
              const cell = worksheet0[cellAddress];
              if (cell && cell.v) {
                dates.add(cell.v.toString().trim());
              }
            }
            samplingDate = Array.from(dates).join("、");
          }
          
          // 处理"检测日期" - 只提取右侧一列的多行数据
          else if (cellValue === "检测日期") {
            const dates = new Set<string>();
            // 1. 确定日期字段占据的行数
            let rowSpan = 1;
            for (let i = r + 1; i <= range0.e.r; i++) {
              const nextCell = worksheet0[XLSX.utils.encode_cell({ r: i, c: 0 })];
              if (nextCell.cellValue != "检测日期") break;
              rowSpan++;
            }
            
            // 2. 提取右侧列对应行数的数据
            for (let i = 0; i < rowSpan; i++) {
              const cellAddress = XLSX.utils.encode_cell({ r: r + i, c: 1 });
              const cell = worksheet0[cellAddress];
              if (cell && cell.v) {
                dates.add(cell.v.toString().trim());
              }
            }
            testingDate = Array.from(dates).join("、");
          }
          
          // 处理"采样地址"
          else if (cellValue === "采样地址") {
            const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })];
            if (cellB && cellB.v) {
              samplingAddress = cellB.v.toString();
            }
          }
          
          // 处理"检测地点"
          else if (cellValue === "检测地点") {
            const cellB = worksheet0[XLSX.utils.encode_cell({ r, c: 1 })];
            if (cellB && cellB.v) {
              testingAddress = cellB.v.toString();
            }
          }
          
          // 处理检测依据/仪器表格
          else if (cellValue === "样品类别" && !sampleFlag) {
            sampleFlag = true;
            let rowIndex = r + 1;         
            while (true) {
              // 获取B列单元格（检测依据）
              const cellB = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 1 })];
              // 如果B列为空，则结束循环
              if (!cellB || !cellB.v) break;
              
              // 获取D列单元格（仪器1）
              const cellD = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 3 })];
              // 获取G列单元格（仪器2）
              const cellG = worksheet0[XLSX.utils.encode_cell({ r: rowIndex, c: 6 })];
              
              // 将值存入数组
              basisList.push(cellB.v.toString());
              standardList.push(cellD && cellD.v ? cellD.v.toString() : '');
              instrumentList.push(cellG && cellG.v ? cellG.v.toString() : '');
              
              rowIndex++;
            }
          }
               

        }
      }
    }

    // 获取sheet2水质数据
    const worksheet1 = workbook.Sheets[workbook.SheetNames[1]]
    const range1 = XLSX.utils.decode_range(worksheet1['!ref']!)
    const cColumnValues: string[] = [];
    const fColumnValues: string[] = [];
    const lastColumnValues: string[] = [];
    let row = 9
    let reportNumber = "";
    const lastColIndex = range1.e.c
    // 项目编号
    const fullReportNo = worksheet1['G8'].v.toString();
    const parts = fullReportNo.split('-');
    if (parts.length > 0) {
      reportNumber = parts[0];
    }
    while (true) {
      const cellAddress = `C${row}`
      const cell = worksheet1[cellAddress]
      // 获取F列对应单元格
      const fCellAddress = XLSX.utils.encode_cell({ r: row - 1, c: 5 });
      const fCell = worksheet1[fCellAddress];
      // 获取最后一列对应单元格
      const lastColCellAddress = XLSX.utils.encode_cell({ r: row - 1, c: lastColIndex })
      const lastColCell = worksheet1[lastColCellAddress]
      
      if (cell && cell.v !== null && cell.v !== '') {
        cColumnValues.push(cell.v.toString())

        // 添加F列数据 (单位)
        if (fCell && fCell.v !== null && fCell.v !== '') {
            fColumnValues.push(fCell.v.toString());
        } else {
            fColumnValues.push(''); // 如果F列单元格为空，推入空字符串
        }

        // 添加最后一列数据
        if (lastColCell && lastColCell.v !== null && lastColCell.v !== '') {
          lastColumnValues.push(lastColCell.v.toString())
        } else {
          lastColumnValues.push('')
        }
        row++
      } else {
        break 
      }
    }
    
    // 检查第三行是否有包含"废水"的单元格
    let shouldIncludeTable = false
    let wastewaterColIndex: number | null = null

    for (let col = range1.s.c; col <= range1.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 2, c: col }) // 第三行 r=2
      const cell = worksheet1[cellAddress]
      if (cell && typeof cell.v === 'string' && cell.v.includes('废水')) {
        shouldIncludeTable = true
        wastewaterColIndex = col
        break
      }
    }
    
    
    // 收集所有废水列（从找到的列开始向右，直到第8行没有数据）
    const wastewaterCols: number[] = []
    if (shouldIncludeTable && wastewaterColIndex !== null) {
      let currentCol = wastewaterColIndex
      while (currentCol <= range1.e.c) {
        const cellAddress8 = XLSX.utils.encode_cell({ r: 7, c: currentCol }) // 第8行 r=7
        const cell8 = worksheet1[cellAddress8]
        if (cell8 && cell8.v !== null && cell8.v !== '') {
          wastewaterCols.push(currentCol)
          currentCol++
        } else {
          break
        }
      }
    }
    
    // 计算表格行列数
    const tableRows = 4 + cColumnValues.length // 4行固定行 + 项目行数
    const tableCols = wastewaterCols.length + 2 // 项目列 + 废水列 + 限值列
    

    // section1Children:检测报告首页
    const section1Children: any[] = [
      new Paragraph({}),
      new Paragraph({}),
      new Paragraph({}),
      new Paragraph({
        spacing: {
          line: 500,
          lineRule: LineRuleType.EXACT // 固定值行距
        }
      }),

      // --- 第一个表格：标题信息 ---
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: { // 隐藏所有边框
          top: { style: 'none' },
          bottom: { style: 'none' },
          left: { style: 'none' },
          right: { style: 'none' },
          insideHorizontal: { style: 'none' },
          insideVertical: { style: 'none' }
        },
        rows: [
          new TableRow({
            height: { value: 777, rule: "exact" }, // 1.37cm
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.TOP,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: '浙江中一检测研究院股份有限公司',
                        font: '黑体',
                        size: 48, // 小一
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 437, rule: "exact" }, // 0.77cm
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.TOP,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: 'ZHEJIANG ZHONGYI TEST INSTITUTE CO.,LTD',
                        font: 'Times New Roman',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 1287, rule: "exact" }, // 2.27cm
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: '检 测 报 告',
                        font: '黑体',
                        size: 72, // 小初
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 425, rule: "exact" }, // 0.75cm
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: 'Test Report',
                        font: 'Times New Roman',
                        size: 32, // 三号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 840, rule: "exact" }, // 0.74cm * 2
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: '报告编号：',
                        font: '宋体',
                        size: 30, // 小三
                        bold: true
                      }),
                      new TextRun({
                        text: reportNumber,
                        font: 'Times New Roman',
                        size: 30, // 小三
                        bold: false
                      })
                    ]
                  })
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 289, rule: "exact" }, // 0.51cm
            children: [
              new TableCell({
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    indent: { firstLine: 13.58 * 240 }, // 13.58字符
                    alignment: AlignmentType.LEFT, // 缩进后左对齐更符合要求
                    children: [
                      new TextRun({
                        text: 'Report No.',
                        font: 'Times New Roman',
                        size: 24, // 小四
                        bold: false
                      })
                    ]
                  })
                ]
              })
            ]
          })
        ]
      }),

      new Paragraph({
        spacing: {
          line: 1100,
          lineRule: LineRuleType.EXACT // 固定值行距
        }
      }),

      // --- 第二个表格：项目、委托单位信息 ---
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [20.5, 79.5], // 设置列宽百分比
        borders: { // 初始隐藏所有边框
          top: { style: 'none' },
          bottom: { style: 'none' },
          left: { style: 'none' },
          right: { style: 'none' },
          insideHorizontal: { style: 'none' },
          insideVertical: { style: 'none' }
        },
        rows: [
          new TableRow({
            height: { value: 573, rule: "exact" }, // 1.01cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    children: [
                      new TextRun({
                        text: '项目名称',
                        font: '宋体',
                        size: 28, // 四号
                        bold: true
                      })
                    ]
                  })
                ],

              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  createCenteredMixedFontParagraph(projectName, 14) // 使用辅助函数处理中英文混合
                ],

              })
            ]
          }),
          new TableRow({
            height: { value: 238, rule: "exact" }, // 0.42cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                verticalAlign: VerticalAlign.TOP,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: 'Project name',
                        font: 'Times New Roman',
                        size: 20, // 10号
                        bold: true
                      })
                    ]
                  })
                ],

              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.TOP,
                children: [
                ],

              })
            ]
          }),
          new TableRow({
            height: { value: 573, rule: "exact" }, // 1.01cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    children: [
                      new TextRun({
                        text: '委托单位',
                        font: '宋体',
                        size: 28, // 四号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  createCenteredMixedFontParagraph(clientName, 14) // 使用辅助函数处理中英文混合
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 272, rule: "exact" }, // 0.48cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: 'Client',
                        font: 'Times New Roman',
                        size: 20, // 10号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.TOP,
                children: [
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 573, rule: "exact" }, // 1.01cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    children: [
                      new TextRun({
                        text: '委托单位地址',
                        font: '宋体',
                        size: 28, // 四号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  createCenteredMixedFontParagraph(clientAddress, 14)
                ]
              })
            ]
          }),
          new TableRow({
            height: { value: 210, rule: "exact" }, // 0.37cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.BOTTOM,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: 'Address',
                        font: 'Times New Roman',
                        size: 20, // 10号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.TOP,
                children: [
                ]
              })
            ]
          })
        ]
      }),

      new Paragraph({
        spacing: {
          line: 500,
          lineRule: LineRuleType.EXACT // 固定值行距
        }
      }), // 空一行

      // --- 第三个表格：签名和日期 ---
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },

        borders: { // 初始隐藏所有边框
          top: { style: 'none' },
          bottom: { style: 'none' },
          left: { style: 'none' },
          right: { style: 'none' },
          insideHorizontal: { style: 'none' },
          insideVertical: { style: 'none' }
        },
        rows: [
          // A1-A6 合并单元格（图片）
          new TableRow({
            height: { value: 323, rule: "exact" }, // 0.57cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                width: {size: 53.1, type: WidthType.PERCENTAGE},
                verticalAlign: VerticalAlign.TOP,
                columnSpan: 1, // 不跨列，因为是第一个单元格
                rowSpan: 6,    // 跨6行
                children: [
               new Paragraph({
                  children: [
                    new ImageRun({
                      data: `data:image/png;base64,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g3dffPHF21taWnTOWKfTCbIsA8B/p5LQglNKEMy8Xq9Op+dyuUCW5TZeKoIs9n/Onz//uwceeOAKf/fhcDjOueSSS34qLy+PRS9ZkiTIyMiAH3/8UYQISpEwwGRyxniRTU1NOfv27buiqqoqfdWqVdNaWlrg8OHDEBsbC8OHD4drrrlm5ZQpU16QZflgpIWCukmExsbGKfv27Zvx3HPPzTt69CiYzWbo1asXDBs2zBETE2OfNm3a6pEjR36XlJS0LppJuTsjTzzxRMXy5cvT0LvkOA5mz579/bPPPnuHyWQqpV+/bNmyd//0pz/d7vF4dLCzWq1wySWXQEZGhh6ixVAskgug94kjvnieh5aWFnC5XOByuaCsrAw8Hg/k5+frryFDt263G3bu3Pk/aWlpX/m7l507dz7y+9//frGiKGhwgclkgq+//vqFs88++3EGmEyYhFk0TUs7fvz4zCVLljyjqqpQVVUVV19fD/X19fqoIp/PB5dffnnl/fffPy8hIWGHIAgsBBt6kauqqma88847L/7973/PsFqtupcSGxsLSUlJcNZZZ1X27t27+oYbbng8Pj5+uyAILHQbQE6ePDl76tSpH9fV1YEsy9g+UvXjjz/e1KdPn++N/mb58uVvL1q0aB7mJbFCddmyZWsuuuii2RzHqT6fT+I4TiX/rvVnDq/XG0//XFGULFVV4xsbG89auHDhn37++WfdC8XPsVgscN555xUvX778XH+0ez6fz7xkyZIVr7322lVYhMTzPPTr1w/y8vLiOY5rYoDJhEkYlHNdXd3UAwcOTFuxYsXsioqKuKKiIj20hNavyWSC7OxsuOmmm5afd95577Eht+EXl8s1+ueff37g7bffno3eJk7LQK9nxIgR0KtXr6b/9//+34rs7Ow1ycnJq4FV2p4il112me/AgQP6EGeXywV33333Z4888sh1/v7mo48+evupp56ah/2PVqsVHA4HfPfdd2+MGDHi3tO9puPHjz/w/vvvv/yPf/wD6uvrITY2Vq/A5TgOPv/880fHjh37or+/dzqdOSNHjixArltBEMDtdsOSJUtWz5gx4yoGmEyYhEhaWlom7d27d+aHH344t6amRm5oaID6+no9hCQIAjgcDhBFEcxmM1x00UXVTz755LVWq7WAhQK7Tnw+n7Wurm7qDz/8MP/VV1+djIUekiTphSQm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C1UfD5OuEa/Xm7Jly5an33333Xn79+/XlQOGlrDaD2cEosLAw0+2hACAHtrB3AgCJoIoSXCAByYxMRFuv/32jTfeeONFbEWYdIMIK1as+PbNN9+cipWY6FkS50T3LMkmfwRaMtKCxiFGY7BPE0GC53lwOp16dEWWZRgwYADcddddyyZNmnSrz+eznk6UxefzWV9++eX89957byjyueLYvDvvvPNf999//0OyLB/1eDwJR44cudlkMjlaWlqSzWZzs9vttgwdOnSFKIol4XjQx44du62lpaWXzWardTqdCSaTySmKovLkk0/+defOndDQ0KCHbZFQfvXq1a+PHz8+IqcTMcA8w6S8vPyWuXPnvl9XV6eDIlp3CJpYDEFafag8yEkOqFjosBX5WvRQyXDs1KlTK//4xz9OkiSpiK0Ik+4Qp9M57o9//OO3X3/9dUpsbKxOaC5Jkq686X1PGo/k91jURrZNYagWQSA2NlafyoPeaWpqKnzwwQc39enTZ/np3o+qqkNvvvnmg5s3b26TClFVFebNm/fZnXfe+XhsbOyhnvDsv/vuu49effXVOaWlpXpFrNlsBpfLBS+88EJEG9IMMM8gqampmXnfffd9tnv3br2MHj1ERVFAlmXdkySLGWjFgd4i+TvS6safo/LBEK+qqpCQkAAffvjhfYzNh0l3S0NDw7SLLrroW2yyxwI3MpyKk3jwe6PWE/QmMRyKgsYn/gwjOEgcIIoipKenQ05OTulTTz2Vebr343A4ct99993333rrrWy8fkVRwGQyQVpaGrz66qsv5+bmPgrd1NPo9XoTli5d+s0nn3wy4dixY/qUItQz77///vLf/OY3z0iSVBype47lMM8MkYuKih657bbbPjt06JBexIOTQrDniSRHx4If2ppGC5kulCAbucnKWEz2a5oGMTEx8OSTT34ycODA99iSMOluSUhI+O6dd95ZEhcXBw6HQwc2sjIcQ7AYrkXPkQRMJCzAlARpdCLDkM/ng5iYmDY5UI/HA5WVlfDFF19kEG0nnRar1br1yiuvXD5mzBhQVbVNoU11dTUsWLDgwd/97neOqqqqmV38qE1Hjhy564477jj62muvTaioqNB1EHrnt912G0Q6WAIACH/84x/ZyevBomla+s8///zyn//854cOHDgAkiSBKIpgt9v1EUFoCWOoiGQDQaG9S3pO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`, // 替换为实际的 Base64 字符串
                      transformation: {
                        width: 105.5,  // 2.79 cm
                        height: 118, // 3.12 cm
                      },
                        type: "png", // 显式指定图片类型
                    })
                  ]
                })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                width: {size: 16.8, type: WidthType.PERCENTAGE},
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '编制人',
                        font: '宋体',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                width: {size: 30.1, type: WidthType.PERCENTAGE},
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 412, rule: "exact" }, // 0.36cm  * 2
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: 'Compiled by',
                        font: 'Times New Roman',
                        size: 14, // 7号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 346, rule: "exact" }, // 0.61cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '审核人',
                        font: '宋体',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 14, // 7号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 301, rule: "exact" }, // 0.53cm  * 2
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: 'Inspected by',
                        font: 'Times New Roman',
                        size: 14, // 7号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 340, rule: "exact" }, // 0.61cm
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '批准人',
                        font: '宋体',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 320, rule: "exact" }, 
            children: [
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: 'Approved by',
                        font: 'Times New Roman',
                        size: 14, // 7号
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 323, rule: "exact" }, // 0.57cm
            children: [
              // A7 单元格
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } }, 
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '检测单位（盖章）',
                        font: '宋体',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.DISTRIBUTE,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '报告日期',
                        font: '宋体',
                        size: 24, // 小四
                        bold: true
                      })
                    ]
                  })
                ]
              }),
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: '2025',
                        font: 'Times New Roman',
                        size: 22, // 11号
                        bold: true
                      }),
                      new TextRun({
                        text: '年',
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      }),
                      new TextRun({
                        text: '   ',
                        font: 'Times New Roman',
                        size: 22, 
                        bold: true
                      }),
                      new TextRun({
                        text: '月',
                        font: '宋体',
                        size: 22,
                        bold: true
                      }),
                      new TextRun({
                        text: '   ',
                        font: 'Times New Roman',
                        size: 22, 
                        bold: true
                      }),
                      new TextRun({
                        text: '日',
                        font: '宋体',
                        size: 22,
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          }),

          new TableRow({
            height: { value: 352, rule: "exact" }, // 0.31cm * 2
            children: [
              // A8 单元格
              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    indent: { firstLine: 0.5 * 1440 / 12 }, // 0.5字符 * 120 twips/字符 (近似)
                    alignment: AlignmentType.LEFT, // 缩进后左对齐
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: 'Detection unit (seal)',
                        font: 'Times New Roman',
                        size: 16, // 8号
                        bold: true
                      })
                    ]
                  })
                ]
              }),

              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        text: 'Report date',
                        font: 'Times New Roman',
                        size: 14, // 7号
                        bold: true
                      })
                    ]
                  })
                ]
              }),

              new TableCell({
                margins: customTableStyle.margins,
                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                verticalAlign: VerticalAlign.CENTER,
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: {
                      line: 120, // 6磅
                      lineRule: LineRuleType.AT_LEAST,
                    },
                    children: [
                      new TextRun({
                        font: '宋体',
                        size: 22, // 11号
                        bold: true
                      })
                    ]
                  })
                ]
              })
            ]
          })
        ]
      }),

      new Paragraph({
        spacing: {
          line: 600,
          lineRule: LineRuleType.EXACT // 固定值行距
        }
      }),

      // --- 第四个表格：中一检测单位信息 ---
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: { // 初始隐藏所有边框
          top: { style: 'none' },
          bottom: { style: 'none' },
          left: { style: 'none' },
          right: { style: 'none' },
          insideHorizontal: { style: 'none' },
          insideVertical: { style: 'none' }
        },
        rows: [
          // A1-B1 合并单元格
          new TableRow({
            children: [
              createFormattedCell0("浙江中一检测研究院股份有限公司  ZHEJIANG ZHONGYI TEST INSTITUTE CO.,LTD", false, 2, AlignmentType.LEFT),       
            ]
          }),
          new TableRow({
            children: [
              createFormattedCell0("地址Address:浙江省宁波市高新区清逸路69号C幢", false, 1, AlignmentType.LEFT),       
              createFormattedCell0("邮编Post Code:315040", false, 1, AlignmentType.LEFT), 
            ]
          }),
          new TableRow({
            children: [
              createFormattedCell0("电话Tel:0574-87908555  87837222  87836111", false, 1, AlignmentType.LEFT),
              createFormattedCell0("传真Fax: 0574-87835222", false, 1, AlignmentType.LEFT),             
            ]
          }),
          new TableRow({
            children: [
              createFormattedCell0("网址Web: www.zynb.com.cn", false, 1, AlignmentType.LEFT),
              createFormattedCell0("Email: zyjc@zynb.com.cn", false, 1, AlignmentType.LEFT),             
            ]
          }),
        ]
      }),
    ];


    // section2Children:检测声明
    const section2Children: any[] = [
      new Paragraph({children: [new PageBreak()]}), // 分页符

      createFormattedParagraph('检 测 声 明', {
        font: '宋体',
        size: 22, // 二号
        bold: true
      }),
      
      createFormattedParagraph('Test report statement', {
        font: 'Times New Roman',
        size: 15, // 小三
        bold: false
      }),
      
      // 空一行
      new Paragraph({}),
      
      // 检测说明文本
      createParagraph0('1、本机构保证检测工作的公正性、独立性和诚实性，对检测的数据负责。', '宋体'),
      createParagraph1('We ensure the testing data impartiality, independence and integrity, and responsible for the testing data.', 'Times New Roman'),
      createParagraph0('2、本报告不得涂改、增删。', '宋体'),
      createParagraph1('The report shall not be altered, added and deleted.', 'Times New Roman'),
      createParagraph0('3、本报告无公司检验检测专用章无效', '宋体'),
      createParagraph1('The report is invalid without “The Special Stamp for Inspection & Test Report”.', 'Times New Roman'),
      createParagraph0('4、本报告无审核人、批准人签名无效。', '宋体'),
      createParagraph1('The report is invalid without the verifier and the approver.', 'Times New Roman'),
      createParagraph0('5、本报告只对采样/送检样品检测结果负责。', '宋体'),
      createParagraph1('The results relate only to the items tested.', 'Times New Roman'),
      createParagraph0('6、对本报告有疑议,请在收到报告15天内与本公司联系。', '宋体'),
      createParagraph1('Please contacts with us within 15 days after you received this report if you have any questions with it .', 'Times New Roman'),
      createParagraph0('7、未经本公司书面允许，对本检测报告局部复印无效，本单位不承担任何法律责任。', '宋体'),
      createParagraph1('The local copy of the report is invalid without prior written permission of our unit, our company will not bear any legal responsibility.', 'Times New Roman'),
      createParagraph0('8、本报告未经同意不得作为商业广告使用。', '宋体'),
      createParagraph1('The reports shall not be published as advertisement without the approval of us.', 'Times New Roman'),
      createParagraph0('9、委托方要求对检测结果进行符合性判定时，如无特殊说明，本公司根据委托方提供的标准限值，采用实测值进行符合性判定，不考虑不确定度所带来的风险，据此判定方式引发的风险由委托方自行承担，本公司不承担连带责任。', '宋体'),
      createParagraph1('When the client requests the conformity judgment of the test results,if there is no special instructions,the company will use the actual measured value to make the conformity judgment according to the evaluation standards provided by the client, and the risk arised by the uncertainty is not considered. The risks caused are borne by the entrusting party, and the company does not bear joint liability.', 'Times New Roman'),
    ]
    

    // section3Children:检测说明
    const section3Children: any[] = [
      new Paragraph({children: [new PageBreak()]}), // 分页符

      createFormattedParagraph('检 测 说 明', {
        font: '宋体',
        size: 22,
        bold: true
      }),

      createFormattedParagraph('Test Description', {
        font: 'Times New Roman',
        size: 15,
        bold: false
      }),
      
      // 空一行
      new Paragraph({}),
      
      // 检测说明表格
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          // 第一行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A1~B1: 样品类别
              createFormattedCell1('样品类别', 'Sample Type'),
              createFormattedCell0(sampleType, false),
              // C1~D1: 检测类别
              createFormattedCell1('检测类别', 'Type'),
              createFormattedCell0(Type, false)
            ]
          }),
          
          // 第二行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A2~B2: 采样日期
              createFormattedCell1('采样日期', 'Sample Date'),
              createFormattedCell0(samplingDate, false),
              // C2~D2: 检测日期
              createFormattedCell1('检测日期', 'Testing date'),
              createFormattedCell0(testingDate, false)
            ]
          }),
          
          // 第三行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A3~D3: 采样地址
              createFormattedCell1('采样地址', 'Sample Address'),
              new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  children: [
                    new TextRun({
                      text: samplingAddress,
                      font: '宋体',
                      size: 21,
                    })
                  ]
                })]
              })
            ]
          }),
          
          // 第四行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A4~D4: 检测地点
              createFormattedCell1('检测地点', 'Testing Address'),
              new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  children: [
                    new TextRun({
                      text: testingAddress,
                      font: '宋体',
                      size: 21,
                    })
                  ]
                })]
              })
            ]
          }),
          
          // 第五行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A5~D5: 采样方法
              createFormattedCell1('采样方法', 'Sample Standard'),
              new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph('')]
              })
            ]
          }),
          
          // 第六行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A6~D6: 评价标准
              createFormattedCell1('评价标准', 'Evaluation Standard'),
              new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph('')]
              })
            ]
          }),
          
          // 第七行
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              // A7~B7: 备注
              createFormattedCell1('备注', 'Note'),
              new TableCell({
                columnSpan: 3,
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph('')]
              })
            ]
          })
        ]
      }),
      
      // 空一行
      new Paragraph({}),
    ]


    // section4Children:检测结果
    const section4Children: any[] = [
      new Paragraph({children: [new PageBreak()]}), // 分页符
      
      // 原有检测结果标题
      createFormattedParagraph('检 测 结 果', {
        font: '宋体',
        size: 22, // 二号字 = 22磅
        bold: true
      }),
      
      // Test Conclusion
      createFormattedParagraph('Test Conclusion', {
        font: 'Times New Roman',
        size: 15, // 小三 = 15磅
        bold: false
      })
    ]
    
    // 添加废水表格（如果存在废水列）
    if (shouldIncludeTable && tableRows > 0 && tableCols > 0 && wastewaterCols.length > 0) {
      // 添加标题
      section4Children.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            line: 360, // 1.5 倍行距：240 × 1.5 = 360（单位：twips）
            lineRule: LineRuleType.AUTO,
          },
          children: [
            new TextRun({
              text: '表1、废水检测结果',
              bold: true,
              font: '宋体',
              size: 21, // 五号字
            }),
          ],
        })
      )
      
      // 创建表格
      const tableRowsData: TableRow[] = []
      
      // 添加前4行特殊内容
      const specialRows = [
        { text: '检测点位', row: 0 },
        { text: '采样日期', row: 1 },
        { text: '采样时间', row: 2 },
        { text: '样品性状', row: 3 }
      ]
      
      // 1. 准备废水列第一行数据（检测点位）用于智能合并
      const locationValues: string[] = []
      wastewaterCols.forEach(col => {
        const cellAddress = XLSX.utils.encode_cell({ r: 3, c: col }) // 第4行 r=3
        const cell = worksheet1[cellAddress]
        locationValues.push(cell && cell.v ? cell.v.toString() : '')
      })
      
      // 智能合并单元格
      const mergedLocationCells = mergeLocationCells(locationValues)
      
      // 2. 创建特殊行（前4行）
      specialRows.forEach((item, rowIndex) => {
        const cells: TableCell[] = []
        
        // 第一列（特殊行标题）
        cells.push(createFormattedCell0(item.text, true))
        
        // 第一行（检测点位行）特殊处理 - 使用合并后的单元格
        if (rowIndex === 0) {
          mergedLocationCells.forEach(cellInfo => {
            cells.push(createFormattedCell0(cellInfo.text, true, cellInfo.colSpan))
          })
        } 
        // 其他行正常处理
        else {
          wastewaterCols.forEach(col => {
            const cellAddress = XLSX.utils.encode_cell({ r: 3 + rowIndex, c: col }) // 第5-7行
            const cell = worksheet1[cellAddress]
            let cellText = ''
            if (cell && cell.v !== null && cell.v !== '') {
              cellText = cell.v.toString()
            }
            if (rowIndex === 3) { // 第四行 "样品性状"在末尾追加“液体”
                cellText += '液体';
            }
            cells.push(createFormattedCell0(cellText, true))
          })
        }
        
        // 最后一列（限值）特殊处理
        if (rowIndex === 0) {
          // 第一行添加跨4行的限值单元格
          cells.push(
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              rowSpan: 4,
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '限值',
                      bold: true,
                      font: '宋体',
                      size: 21,
                    }),
                  ],
                }),
              ],
            })
          )
        } else {
          // 其他行在限值列位置不需要单元格（因为已合并）
        }
        
        tableRowsData.push(new TableRow({ height: { value: 454, rule: "exact" },children: cells }))
      })
      
      // 添加检测项目行
      for (let i = 0; i < cColumnValues.length; i++) {
        const cells: TableCell[] = []
        
        // 第一列（项目名称）
        const itemName = cColumnValues[i] || '';
        const unit = fColumnValues[i] || '';
        const combinedItemText = mergeItemWithUnit(itemName, unit);
        cells.push(createFormattedCell0(combinedItemText)); // 使用合并后的文本
        
        // 废水列数据（从第8行开始对应检测项目）
        wastewaterCols.forEach(col => {
          const cellAddress = XLSX.utils.encode_cell({ r: 8 + i, c: col }) // 第9行 r=8
          const cell = worksheet1[cellAddress]
          let cellText = ''
          if (cell && cell.v !== null && cell.v !== '') {
            cellText = cell.v.toString()
          }
          cells.push(createFormattedCell0(cellText))
        })
        
        // 填入Excel最后一列数据
        cells.push(createFormattedCell0(lastColumnValues[i] || ''))
        
        tableRowsData.push(new TableRow({ height: { value: 567, rule: "exact" },children: cells }))
      } 
      
      // 创建表格
      section4Children.push(
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: tableRowsData,
        })
      )
    }
    

    // section5Children:检测依据/仪器表格
    const section5Children: any[] = []   
    
    if (shouldIncludeTable && tableRows > 0 && tableCols > 0 && wastewaterCols.length > 0) {
      const instrumentTableRows: TableRow[] = [
      // 表头行
        new TableRow({
          height: { value: 567, rule: "exact" },
          children: [
            createFormattedCell1('检测项目', 'Tested Item'),
            createFormattedCell1('检测依据', 'Testing Standard'),
            createFormattedCell1('主要检测仪器', 'Main Instruments')
          ]
        })
      ];

      // 添加数据行
      for (let i = 0; i < basisList.length; i++) {   
        instrumentTableRows.push(
          new TableRow({
            height: { value: 567, rule: "exact" },
            children: [
              createFormattedCell0(basisList[i]), 
              createFormattedCell0(standardList[i]), 
              createFormattedCell0(instrumentList[i]) 
            ]
          })
        );
      }

      section5Children.push(
         new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: {
            line: 360, // 1.5 倍行距：240 × 1.5 = 360（单位：twips）
            lineRule: LineRuleType.AUTO,
          },
          children: [
            new TextRun({
              text: '表5、废水检测项目、检出限、检测依据及主要检测仪器',
              bold: true,
              font: '宋体',
              size: 21, // 五号字
            }),
          ],
        }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: instrumentTableRows
        })
      );   
    }
    // 生成Word
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          ...section1Children,   // 检测报告首页
          ...section2Children,   // 检测声明
          ...section3Children,   // 检测说明
          ...section4Children,   // 检测结果
        ],
      },
      { // 第二个 Section (横向)
        properties: {
          type: SectionType.CONTINUOUS, // 连续分节符，保持页码连续
          page: {
            size: {
              orientation: 'landscape', // 设置为横向
            },
          },
        },
          children: [
            ...section5Children,   // 检测依据/仪器表格
          ],
        }
      ],
    })
    
    // 导出为Word文件
    const blob = await Packer.toBlob(doc)
    saveAs(blob, 'C表.docx')
    ElMessage.success('已导出Word文件')
    
  } catch (error) {
    console.error('导出失败:', error)
    ElMessage.error('导出失败: ' + (error as Error).message)
  }
}
</script>

<template>
  <el-button type="primary" @click="handleClick">
    导出为Word（C表）
  </el-button>
</template>

<style scoped>
</style>