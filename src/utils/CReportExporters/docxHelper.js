// src/components/CReportExporter/docxHelpers.js
import { Paragraph, TextRun, TableCell, VerticalAlign, AlignmentType, WidthType, LineRuleType } from 'docx'

// 表格边距样式
export const customTableStyle = {
  margins: { left: 108, right: 108 } // 0.19cm
}

/**
 * 通用段落创建函数
 * @param {string} text - 输入文本
 * @param {number} size - 输入字体大小
 * @param {number} bold - 是否加粗
 * @param {number} alignment - 对齐方式
 * @returns {Paragraph} 输出段落
 */
export const createParagraph = (text, size, bold = false, alignment = AlignmentType.CENTER) => {
  const runs = []
  for (const char of text) {
    const isAscii = /^[\x00-\x7F]$/.test(char)
    runs.push(new TextRun({
      text: char,
      font: isAscii ? 'Times New Roman' : '宋体',
      size: size * 2,
      bold,
    }))
  }
  return new Paragraph({
    alignment,
    children: runs
  })
}

/**
 * 检测说明中的中/英文段落
 * @param {string} text
 * @param {string} font
 * @returns {Paragraph}
 */
export const createParagraph0 = (text, font, bold = false) => {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: {
      line: 400, // 20磅行距
      lineRule: LineRuleType.AT_LEAST,
    },
    children: [
      new TextRun({
        text,
        bold,
        font,
        size: 24, // 小四
      }),
    ],
  })
}
export const createParagraph1 = (text, font, bold = false) => {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    indent: { left: 400 }, // 首行缩进0.02字符
    spacing: {
      line: 400, // 20磅行距
      lineRule: LineRuleType.AT_LEAST,
    },
    children: [
      new TextRun({
        text,
        bold,
        font,
        size: 24, // 小四
      }),
    ],
  })
}

/**
 * 表格单元格格式
 * @param {string} text
 * @param {boolean} bold - 是否加粗
 * @param {number} [colSpan] - 是否跨列
 * @param {AlignmentType} [alignment=AlignmentType.CENTER] - 是否居中
 * @returns {TableCell} 输出单元格
 */
export const createFormattedCell0 = (text, bold = false, colSpan, alignment = AlignmentType.CENTER) => {
  const runs = []
  for (const char of text) {
    const isAscii = /^[\x00-\x7F]$/.test(char) // 英文/数字/符号
    runs.push(new TextRun({
      text: char,
      bold,
      font: isAscii ? 'Times New Roman' : '宋体',
      size: 21, // 五号
    }))
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

/**
 * 设置表格单元格格式（中英文两行）
 * @param {string} chineseText
 * @param {string} englishText
 * @returns {TableCell}
 */
export const createFormattedCell1 = (chineseText, englishText) => {
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
            size: 21, // Size 5
          }),
        ],
      }),
    ],
  })
}

/**
 * 合并项目名称和单位
 * @param {string} itemName - 输入项目名称
 * @param {string} unit - 输入单位
 * @returns {string} 输出合并字符串
 */
export const mergeItemWithUnit = (itemName, unit) => {
  if (!unit) {
    return itemName
  }
  const hasChinese = /[\u4e00-\u9fa5]/.test(unit)
  if (hasChinese) {
    return `${itemName}（${unit}）` // 中文单位加括号
  } else {
    return `${itemName} ${unit}` // 英文单位用空格分隔
  }
}

/**
 * 智能合并单元格函数
 * @param {string[]} values - 输入多个单元格内容组成的数组
 * @returns {Array<{text: string, colSpan: number}>} 
 */
export const mergeLocationCells = (values) => {
  const mergedCells = []
  let i = 0
  while (i < values.length) {
    const currentValue = values[i]
    if (currentValue && currentValue.trim() !== '') {
      let j = i + 1
      while (j < values.length && (!values[j] || values[j].trim() === '')) {
        j++
      }
      const colSpan = j - i
      mergedCells.push({ text: currentValue, colSpan })
      i = j
    } else {
      let j = i + 1
      while (j < values.length && (!values[j] || values[j].trim() === '')) {
        j++
      }
      const colSpan = j - i
      mergedCells.push({ text: '', colSpan })
      i = j
    }
  }
  return mergedCells
}


