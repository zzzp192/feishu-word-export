import $ from 'jquery';
import { bitable, IField, FieldType } from '@lark-base-open/js-sdk';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import './index.scss';

// 定义需要导出的字段名，必须与表格中的表头名称一致
const TARGET_FIELDS = [
  "任务名称",
  "委托单号",
  "期望完成日期",
  "制料阶段",
  "制样阶段",
  "检测阶段",
  "附件"
];

$(async function() {
  $('#exportBtn').on('click', async function() {
    const $status = $('#status');
    $status.text('正在获取数据...');

    try {
      // 1. 获取当前选区（tableId 和 recordId）
      const selection = await bitable.base.getSelection();
      const tableId = selection.tableId;
      const recordId = selection.recordId;

      if (!tableId || !recordId) {
        $status.text('错误：请先在多维表格中选中一行数据！');
        return;
      }

      const table = await bitable.base.getTableById(tableId);
      
      // 2. 获取所有字段元数据，建立 字段名 -> 字段对象 的映射
      const fieldMetaList = await table.getFieldMetaList();
      const fieldMap = new Map<string, string>(); // Name -> FieldId
      
      fieldMetaList.forEach(meta => {
        fieldMap.set(meta.name, meta.id);
      });

      // 3. 提取数据
      const templateData: any = {};
      
      for (const fieldName of TARGET_FIELDS) {
        const fieldId = fieldMap.get(fieldName);
        if (!fieldId) {
          console.warn(`未找到字段: ${fieldName}`);
          templateData[fieldName] = "未找到字段";
          continue;
        }

        const field = await table.getFieldById(fieldId);
        const val = await field.getValue(recordId);

        // 格式化数据 (将不同类型的字段值转为字符串)
        templateData[fieldName] = await formatFieldValue(field, val);
      }

      $status.text('正在生成 Word...');

      // 4. 读取 public/template.docx 模板文件
      // 这里的路径 './template.docx' 对应的是构建后或开发服务器根目录下的文件
      const response = await fetch('./template.docx');
      if (!response.ok) {
        throw new Error('无法加载模板文件，请确保 template.docx 在 public 目录中');
      }
      const content = await response.arrayBuffer();

      // 5. 使用 docxtemplater 渲染模板
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // 填入数据
      doc.render(templateData);

      // 6. 导出文件
      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      // 文件名可以使用 任务名称 + .docx
      const fileName = `${templateData['任务名称'] || '导出文档'}.docx`;
      saveAs(out, fileName);
      
      $status.text(`导出成功：${fileName}`);

    } catch (error: any) {
      console.error(error);
      $status.text(`导出失败：${error.message}`);
    }
  });
});

// 辅助函数：根据字段类型格式化值为字符串
async function formatFieldValue(field: IField, val: any): Promise<string> {
  if (val === null || val === undefined) return "";

  const type = await field.getType();

  // 1. 日期类型 (简单处理，时间戳转 YYYY-MM-DD)
  if (type === FieldType.DateTime) {
     const date = new Date(val);
     return date.toLocaleDateString() + " " + date.toLocaleTimeString(); 
  }

  // 2. 单选/多选 (通常是对象或对象数组)
  if (type === FieldType.SingleSelect) {
    return val.text || "";
  }
  if (type === FieldType.MultiSelect) {
    return val.map((v: any) => v.text).join(", ");
  }

  // 3. 附件 (返回文件名列表)
  if (type === FieldType.Attachment) {
    if (Array.isArray(val)) {
      return val.map((v: any) => v.name).join("\n");
    }
    return "";
  }
  
  // 4. 文本/数字/其他
  if (typeof val === 'object') {
     // 尝试提取 text 属性，这是飞书很多字段的通用结构
     if (Array.isArray(val)) {
        return val.map(v => v.text || JSON.stringify(v)).join("");
     }
     return val.text || JSON.stringify(val);
  }

  return String(val);
}