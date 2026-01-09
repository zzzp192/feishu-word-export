import { bitable, IField, FieldType } from '@lark-base-open/js-sdk';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import './index.scss';

// 定义需要导出的字段名
const TARGET_FIELDS = [
  "任务名称",
  "委托单号",
  "期望完成日期",
  "制料阶段",
  "制样阶段",
  "检测阶段",
  "附件"
];

// 使用原生 DOMContentLoaded 事件替代 $(function() {})
document.addEventListener('DOMContentLoaded', async function () {
  
  const exportBtn = document.getElementById('exportBtn');
  const statusDiv = document.getElementById('status');

  if (!exportBtn || !statusDiv) return;

  // 辅助函数：更新状态文字
  const setStatus = (msg: string) => {
    statusDiv.innerText = msg;
  };

  exportBtn.addEventListener('click', async function (e) {
    // 【关键修复】阻止事件冒泡，防止触发飞书宿主环境的 e.closest 报错
    e.stopPropagation();
    e.preventDefault();

    setStatus('正在获取数据...');

    try {
      // 1. 获取选区
      const selection = await bitable.base.getSelection();
      const tableId = selection.tableId;
      const recordId = selection.recordId;

      if (!tableId || !recordId) {
        setStatus('错误：请先在多维表格中选中一行数据！');
        return;
      }

      const table = await bitable.base.getTableById(tableId);
      const fieldMetaList = await table.getFieldMetaList();
      const fieldMap = new Map<string, string>();
      
      fieldMetaList.forEach(meta => {
        fieldMap.set(meta.name, meta.id);
      });

      // 2. 提取数据
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
        templateData[fieldName] = await formatFieldValue(field, val);
      }

      setStatus('正在读取模板...');

      // 3. 读取模板
      // 注意：这里必须是 './template.docx'，且文件必须在 public 目录下
      const response = await fetch('./template.docx');
      if (!response.ok) {
        throw new Error(`无法加载模板 (Status: ${response.status})。请确认 template.docx 在 public 文件夹下，且 index.html 在根目录。`);
      }
      const content = await response.arrayBuffer();

      setStatus('正在生成 Word...');

      // 4. 生成文档
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      doc.render(templateData);

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      const fileName = `${templateData['任务名称'] || '导出文档'}.docx`;
      saveAs(out, fileName);
      
      setStatus(`导出成功：${fileName}`);

    } catch (error: any) {
      console.error(error);
      setStatus(`导出失败：${error.message}`);
    }
  });
});

// 辅助函数保持不变
async function formatFieldValue(field: IField, val: any): Promise<string> {
  if (val === null || val === undefined) return "";

  const type = await field.getType();

  if (type === FieldType.DateTime) {
     const date = new Date(val);
     // 简单的日期格式化
     return date.toLocaleDateString() + " " + date.toLocaleTimeString(); 
  }

  if (type === FieldType.SingleSelect) {
    return val.text || "";
  }
  if (type === FieldType.MultiSelect) {
    return val.map((v: any) => v.text).join(", ");
  }

  if (type === FieldType.Attachment) {
    if (Array.isArray(val)) {
      return val.map((v: any) => v.name).join("\n");
    }
    return "";
  }
  
  if (typeof val === 'object') {
     if (Array.isArray(val)) {
        return val.map(v => v.text || JSON.stringify(v)).join("");
     }
     return val.text || JSON.stringify(val);
  }

  return String(val);
}
