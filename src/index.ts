import $ from 'jquery';
import { bitable, FieldType } from '@lark-base-open/js-sdk';
import './index.scss';

declare const docx: any;
declare const saveAs: any;

$(async function() {
  const [tableList, selection] = await Promise.all([bitable.base.getTableMetaList(), bitable.base.getSelection()]);
  const optionsHtml = tableList.map(table => `<option value="${table.id}">${table.name}</option>`).join('');
  $('#tableSelect').append(optionsHtml).val(selection.tableId!);

  $('#exportWord').on('click', async function() {
    const tableId = $('#tableSelect').val() as string;
    if (!tableId) return;

    const table = await bitable.base.getTableById(tableId);
    const sel = await bitable.base.getSelection();
    if (!sel.recordId) {
      alert('请先选中一行记录');
      return;
    }

    const fieldMetaList = await table.getFieldMetaList();
    const fieldMap: Record<string, string> = {};
    for (const f of fieldMetaList) {
      fieldMap[f.name] = f.id;
    }

    const fields = ['任务名称', '委托单号', '期望完成日期', '制料阶段', '制样阶段', '检测阶段', '附件'];
    const data: Record<string, string> = {};

    for (const name of fields) {
      const fieldId = fieldMap[name];
      if (fieldId) {
        const cellValue = await table.getCellValue(fieldId, sel.recordId);
        if (Array.isArray(cellValue)) {
          data[name] = cellValue.map((v: any) => v.text || v.name || String(v)).join(', ');
        } else if (cellValue && typeof cellValue === 'object') {
          data[name] = (cellValue as any).text || (cellValue as any).name || JSON.stringify(cellValue);
        } else {
          data[name] = cellValue ? String(cellValue) : '';
        }
      } else {
        data[name] = '';
      }
    }

    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, Packer } = docx;

    const createCell = (text: string, width: number) => new TableCell({
      children: [new Paragraph({ children: [new TextRun(text)] })],
      width: { size: width, type: WidthType.PERCENTAGE },
    });

    const rows = [
      ['任务名称', data['任务名称']],
      ['委托单号', data['委托单号']],
      ['期望完成日期', data['期望完成日期']],
      ['制料阶段（冶炼、锻压、轧制、罩退、涂镀等）', data['制料阶段']],
      ['制样阶段（热冲压、各类型样品切割、烘烤、点焊等，及其他特殊需求）', data['制样阶段']],
      ['检测阶段（金相制样、各检测项目、报告出具等）', data['检测阶段']],
      ['附件（需要图像说明的）', data['附件']],
    ];

    const tableRows = rows.map(([label, value]) => new TableRow({
      children: [createCell(label, 30), createCell(value, 70)],
    }));

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [new TextRun({ text: '任务委托单', bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: '' }),
          new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }),
        ],
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `任务委托单_${data['委托单号'] || '导出'}.docx`);
  });
});
