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

document.addEventListener('DOMContentLoaded', async function () {
  
  const exportBtn = document.getElementById('exportBtn');
  const statusDiv = document.getElementById('status');

  if (!exportBtn || !statusDiv) return;

  const setStatus = (msg: string) => {
    statusDiv.innerText = msg;
  };

  // 【核心修复】定义一个通用的阻止冒泡函数
  const stopEventPropagation = (e: Event) => {
    if (e) {
      e.stopPropagation();
      // 在某些极端情况下，可能还需要阻止立即传播
      e.stopImmediatePropagation();
    }
  };

  // 【核心修复】同时监听 mousedown 和 mouseup，防止报错堆栈中的 _handleMouseUp 触发
  exportBtn.addEventListener('mousedown', stopEventPropagation);
  exportBtn.addEventListener('mouseup', stopEventPropagation);
  
  // 主逻辑
  exportBtn.addEventListener('click', async function (e) {
    // 同样阻止 click 的冒泡
    stopEventPropagation(e);
    
    // 阻止默认行为（比如表单提交）
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
      const response = await fetch('./template.docx');
      if (!response.ok) {
        throw new Error(`无法加载模板 (Status: ${response.status})`);
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

async function formatFieldValue(field: IField, val: any): Promise<string> {
  if (val === null || val === undefined) return "";

  const type = await field.getType();

  if (type === FieldType.DateTime) {
     const date = new Date(val);
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
