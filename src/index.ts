import { bitable, IField, FieldType } from '@lark-base-open/js-sdk';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import './index.scss';

// 1. 注入 CSS 补丁，强制屏蔽按钮内部文字的鼠标事件，解决 e.closest 报错
const style = document.createElement('style');
style.innerHTML = `
  /* 关键：让按钮内的所有子元素（如文字）不响应鼠标，点击事件直接由按钮捕获 */
  #exportBtn * {
    pointer-events: none;
  }
  /* 增加一些点击反馈 */
  #exportBtn:active {
    opacity: 0.8;
    transform: scale(0.98);
  }
`;
document.head.appendChild(style);

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

  // 辅助：阻止事件冒泡的通用函数
  const stopEventPropagation = (e: Event) => {
    if (e) {
      e.stopPropagation();
      e.stopImmediatePropagation();
    }
  };

  // 2. 绑定事件：同时拦截 mousedown, mouseup, click
  // 即使 CSS 修复了 target 问题，阻止冒泡依然是双重保险
  ['mousedown', 'mouseup', 'click'].forEach(eventType => {
    exportBtn.addEventListener(eventType, (e) => {
      // 如果是 click 事件，执行业务逻辑，否则只阻止冒泡
      if (eventType === 'click') {
        handleExport(e);
      } else {
        stopEventPropagation(e);
      }
    });
  });

  // 3. 核心导出逻辑
  async function handleExport(e: Event) {
    stopEventPropagation(e);
    e.preventDefault();

    setStatus('正在获取数据...');

    try {
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

      // 4. 读取模板文件
      const response = await fetch('./template.docx');
      if (!response.ok) {
        throw new Error(`无法加载模板 (Status: ${response.status})。请检查 template.docx 是否在 public 目录下。`);
      }
      const content = await response.arrayBuffer();

      setStatus('正在生成 Word...');

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
  }
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
