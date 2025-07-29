// src/components/CReportExporter/docxGenerator.js
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, AlignmentType, TextRun, VerticalAlign, LineRuleType, PageBreak, SectionType, ImageRun } from 'docx'
import { createParagraph, createParagraph0, createParagraph1, createFormattedCell0, createFormattedCell1, mergeItemWithUnit, mergeLocationCells } from './docxHelper'
import { imageBase64 } from '../../utils/CReportExporters/images/zhongyi_test';
import * as XLSX from 'xlsx'

// 表格常用样式
export const customTableStyle = {
    margins: { left: 108, right: 108 }, // 单元格边距0.19cm
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { // 隐藏所有边框
        top: { style: 'none' },
        bottom: { style: 'none' },
        left: { style: 'none' },
        right: { style: 'none' },
        insideHorizontal: { style: 'none' },
        insideVertical: { style: 'none' }
    },
}

/**
 * 根据上传Excel数据的解析结果生成C表Word
 * @param {Object} data - 来自excelProcessor.js的结构化数据
 * @returns {Promise<Blob>}
 */
export async function generateDocxDocument(data) {
    try {
        // --- Section 1: 检测报告首页 ---
        const section1Children = [
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
                width: customTableStyle.width,
                borders: customTableStyle.borders,
                rows: [
                    new TableRow({
                        height: { value: 777, rule: "exact" }, // 1.37cm
                        children: [
                            new TableCell({
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
                                verticalAlign: VerticalAlign.TOP,
                                children: [
                                    createParagraph('ZHEJIANG ZHONGYI TEST INSTITUTE CO.,LTD', 12, true) // 小四
                                ]
                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 1287, rule: "exact" }, // 2.27cm
                        children: [
                            new TableCell({
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
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('Test Report', 16, true) // 三号
                                ]
                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 840, rule: "exact" }, // 0.74cm * 2
                        children: [
                            new TableCell({
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
                                                text: data.reportNumber,
                                                font: 'Times New Roman',
                                                size: 30,
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
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    new Paragraph({
                                        indent: { firstLine: 13.58 * 240 }, // 13.58字符
                                        alignment: AlignmentType.LEFT,
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
                width: customTableStyle.width,
                borders: customTableStyle.borders,
                rows: [
                    new TableRow({
                        height: { value: 573, rule: "exact" }, // 1.01cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                width: { size: 20.6, type: WidthType.PERCENTAGE },
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('项目名称', 14, true, AlignmentType.DISTRIBUTE) // 四号
                                ],

                            }),
                            new TableCell({
                                margins: customTableStyle.margins,
                                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph(data.projectName, 14)
                                ],

                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 238, rule: "exact" }, // 0.42cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                verticalAlign: VerticalAlign.TOP,
                                children: [
                                    createParagraph('Project name', 10, true) // 10号
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
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('委托单位', 14, true, AlignmentType.DISTRIBUTE) // 四号
                                ]
                            }),
                            new TableCell({
                                margins: customTableStyle.margins,
                                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph(data.clientName, 14)
                                ]
                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 272, rule: "exact" }, // 0.48cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('Client', 10, true) // 10号
                                ]
                            }),
                            new TableCell({
                                margins: customTableStyle.margins,
                                borders: { top: { style: 'none' }, bottom: { style: 'none' }, left: { style: 'none' }, right: { style: 'none' } },
                                verticalAlign: VerticalAlign.TOP,
                                children: []
                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 573, rule: "exact" }, // 1.01cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('委托单位地址', 14, true, AlignmentType.DISTRIBUTE) // 四号
                                ]
                            }),
                            new TableCell({
                                margins: customTableStyle.margins,
                                borders: { top: { style: 'none' }, bottom: { style: 'single' }, left: { style: 'none' }, right: { style: 'none' } },
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph(data.clientAddress, 14)
                                ]
                            })
                        ]
                    }),
                    new TableRow({
                        height: { value: 210, rule: "exact" }, // 0.37cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                verticalAlign: VerticalAlign.BOTTOM,
                                children: [
                                    createParagraph('Address', 10, true) // 10号
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
            }),

            // --- 第三个表格：签名和日期 ---
            new Table({
                width: customTableStyle.width,

                borders: customTableStyle.borders,
                rows: [
                    // A1-A6 合并单元格（图片）
                    new TableRow({
                        height: { value: 323, rule: "exact" }, // 0.57cm
                        children: [
                            new TableCell({
                                margins: customTableStyle.margins,
                                width: { size: 53.1, type: WidthType.PERCENTAGE },
                                verticalAlign: VerticalAlign.TOP,
                                columnSpan: 1, // 不跨列，因为是第一个单元格
                                rowSpan: 6,    // 跨6行
                                children: [
                                    new Paragraph({
                                        children: [
                                            new ImageRun({
                                                data: imageBase64,
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
                                width: { size: 16.8, type: WidthType.PERCENTAGE },
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
                                width: { size: 30.1, type: WidthType.PERCENTAGE },
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
                width: customTableStyle.width,
                borders: customTableStyle.borders,
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
        ]

        // --- Section 2: 检测声明 ---
        const section2Children = [
            new Paragraph({ children: [new PageBreak()] }), //分页符
            createParagraph('检 测 声 明', 22, true), // 二号
            createParagraph('Test report statement', 15), // 小三
            new Paragraph({}),

            createParagraph0('1、本机构保证检测工作的公正性、独立性和诚实性，对检测的数据负责。', '宋体'),
            createParagraph1('We ensure the testing data impartiality, independence and integrity, and responsible for the testing data.', 'Times New Roman'),
            createParagraph0('2、本报告不得涂改、增删。', '宋体'),
            createParagraph1('The report shall not be altered, added and deleted.', 'Times New Roman'),
            createParagraph0('3、本报告无公司检验检测专用章无效', '宋体'),
            createParagraph1('The report is invalid without “The Special Stamp for Inspection & Test Report”.', 'Times New Roman'),
            createParagraph0('4、本报告无审核人、批准人签名无效。', '宋体'),
            createParagraph1('The report is invalid without signatures of reviewer and approver.', 'Times New Roman'),
            createParagraph0('5、本报告涂改、复制无效。', '宋体'),
            createParagraph1('The report is invalid if altered or copied.', 'Times New Roman'),
            createParagraph0('6、本报告仅对来样负责。', '宋体'),
            createParagraph1('The report is responsible for the sample only.', 'Times New Roman'),
            createParagraph0('7、本报告未经本单位书面同意不得部分复制。', '宋体'),
            createParagraph1('The local copy of the report is invalid without prior written permission of our unit, our company will not bear any legal responsibility.', 'Times New Roman'),
            createParagraph0('8、本报告未经同意不得作为商业广告使用。', '宋体'),
            createParagraph1('The reports shall not be published as advertisement without the approval of us.', 'Times New Roman'),
            createParagraph0('9、委托方要求对检测结果进行符合性判定时，如无特殊说明，本公司根据委托方提供的标准限值，采用实测值进行符合性判定，不考虑不确定度所带来的风险，据此判定方式引发的风险由委托方自行承担，本公司不承担连带责任。', '宋体'),
            createParagraph1('When the client requires conformity judgment of the test results, if there is no special description, our company will make conformity judgment according to the standard limit values provided by the client and the measured values, without considering the risk brought by uncertainty. The risk caused by this judgment method shall be borne by the client, and our company shall not bear joint liability.', 'Times New Roman'),
        ]

        // --- Section 3: 检测说明 ---
        const section3Children = [
            new Paragraph({ children: [new PageBreak()] }), // 分页符
            createParagraph('检 测 说 明', 22, true), // 二号
            createParagraph('Test Description', 15), // 小三
            new Paragraph({}),

            new Table({
                width: customTableStyle.width,
                rows: [
                    // 第一行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A1~B1: 样品类别
                            createFormattedCell1('样品类别', 'Sample Type'),
                            createFormattedCell0(data.sampleType, false),
                            // C1~D1: 检测类别
                            createFormattedCell1('检测类别', 'Type'),
                            createFormattedCell0(data.reportType, false)
                        ]
                    }),

                    // 第二行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A2~B2: 采样日期
                            createFormattedCell1('采样日期', 'Sample Date'),
                            createFormattedCell0(data.samplingDate, false),
                            // C2~D2: 检测日期
                            createFormattedCell1('检测日期', 'Testing date'),
                            createFormattedCell0(data.testingDate, false)
                        ]
                    }),

                    // 第三行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A3~D3: 采样地址
                            createFormattedCell1('采样地址', 'Sample Address'),
                            createFormattedCell0(data.samplingAddress, false, 3, AlignmentType.LEFT)
                        ]
                    }),

                    // 第四行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A4~D4: 检测地点
                            createFormattedCell1('检测地点', 'Testing Address'),
                            createFormattedCell0(data.testingAddress, false, 3, AlignmentType.LEFT)
                        ]
                    }),

                    // 第五行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A5~D5: 采样方法
                            createFormattedCell1('采样方法', 'Sample Standard'),
                            createFormattedCell0('', false, 3)
                        ]
                    }),

                    // 第六行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A6~D6: 评价标准
                            createFormattedCell1('评价标准', 'Evaluation Standard'),
                            createFormattedCell0('', false, 3)
                        ]
                    }),

                    // 第七行
                    new TableRow({
                        height: { value: 567, rule: "exact" },
                        children: [
                            // A7~B7: 备注
                            createFormattedCell1('备注', 'Note'),
                            createFormattedCell0('', false, 3)
                        ]
                    })
                ]
            }),

            new Paragraph({}),
        ]

        // --- Section 4: 检测结果 ---
        const section4Children = [
            new Paragraph({ children: [new PageBreak()] }), // 分页符
            createParagraph('检 测 结 果', 22, true), // 二号
            createParagraph('Test Conclusion', 15), // 小三
        ]
        if (data.shouldIncludeTable && data.wastewaterCols.length > 0) {
            section4Children.push(
                new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        line: 360, // 1.5 倍行距
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
            const tableRowsData = []

            // 添加前4行特殊内容
            const specialRows = [{ text: '检测点位', row: 0 }, { text: '采样日期', row: 1 }, { text: '采样时间', row: 2 }, { text: '样品性状', row: 3 }]

            // 1. 准备废水列第一行数据（检测点位）用于智能合并
            const locationValues = []
            data.wastewaterCols.forEach(col => {
                const cellAddress = XLSX.utils.encode_cell({ r: 3, c: col }) // 第4行 r=3
                const cell = data.worksheet1[cellAddress]
                locationValues.push(cell && cell.v ? cell.v.toString() : '')
            })

            // 智能合并单元格
            const mergedLocationCells = mergeLocationCells(locationValues)

            // 2. 创建特殊行（前4行）
            specialRows.forEach((item, rowIndex) => {
                const cells = []

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
                    data.wastewaterCols.forEach(col => {
                        const cellAddress = XLSX.utils.encode_cell({ r: 3 + rowIndex, c: col }) // 第5-7行
                        const cell = data.worksheet1[cellAddress]
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
                                createParagraph('限值', 10, true) // 10号
                            ],
                        })
                    )
                } else {
                    // 其他行在限值列位置不需要单元格（因为已合并）
                }

                tableRowsData.push(new TableRow({ height: { value: 454, rule: "exact" }, children: cells }))
            })

            // 添加检测项目行
            for (let i = 0; i < data.cColumnValues.length; i++) {
                const cells = []

                // 第一列（项目名称）
                const itemName = data.cColumnValues[i] || '';
                const unit = data.fColumnValues[i] || '';
                const combinedItemText = mergeItemWithUnit(itemName, unit);
                cells.push(createFormattedCell0(combinedItemText)); // 使用合并后的文本

                // 废水列数据（从第8行开始对应检测项目）
                data.wastewaterCols.forEach(col => {
                    const cellAddress = XLSX.utils.encode_cell({ r: 8 + i, c: col }) // 第9行 r=8
                    const cell = data.worksheet1[cellAddress]
                    let cellText = ''
                    if (cell && cell.v !== null && cell.v !== '') {
                        cellText = cell.v.toString()
                    }
                    cells.push(createFormattedCell0(cellText))
                })

                // 填入Excel最后一列数据
                cells.push(createFormattedCell0(data.lastColumnValues[i] || ''))

                tableRowsData.push(new TableRow({ height: { value: 567, rule: "exact" }, children: cells }))
            }

            // 创建表格
            section4Children.push(
                new Table({
                    width: customTableStyle.width,
                    rows: tableRowsData,
                })
            )
        }

        // --- Section 5: 检测依据/仪器表格 ---
        const section5Children = []
        if (data.basisList && data.basisList.length > 0) {
            const instrumentTableRows = [
                new TableRow({
                    height: { value: 567, rule: "exact" },
                    children: [
                        createFormattedCell1('检测项目', 'Tested Item'),
                        createFormattedCell1('检测依据', 'Testing Standard'),
                        createFormattedCell1('主要检测仪器', 'Main Instruments')
                    ]
                })
            ];

            for (let i = 0; i < data.basisList.length; i++) {
                instrumentTableRows.push(new TableRow({
                    height: { value: 567, rule: "exact" },
                    children: [
                        createFormattedCell0(data.basisList[i]),
                        createFormattedCell0(data.standardList[i]),
                        createFormattedCell0(data.instrumentList[i])
                    ]
                }));
            }

            section5Children.push(
                new Paragraph({
                    alignment: AlignmentType.LEFT,
                    spacing: {
                        line: 360, // 1.5倍行距
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
                    width: customTableStyle.width,
                    rows: instrumentTableRows
                })
            );
        }

        // --- 生成Word ---
        const doc = new Document({
            sections: [
                {
                    properties: {},
                    children: [...section1Children, ...section2Children, ...section3Children, ...section4Children],
                },
                {
                    properties: {
                        type: SectionType.CONTINUOUS, // 连续分节符，保持页码连续
                        page: {
                            size: {
                                orientation: 'landscape', // 设置为横向
                            },
                        },
                    },
                    children: [...section5Children], // 检测依据/仪器表格
                },
            ],
        })

        const blob = await Packer.toBlob(doc)
        return blob

    } catch (error) {
        console.error("Error generating DOCX document:", error)
        throw new Error(`生成DOCX文档时出错: ${error.message}`)
    }
}


