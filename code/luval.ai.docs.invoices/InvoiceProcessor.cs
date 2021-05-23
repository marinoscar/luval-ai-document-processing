using Azure.AI.FormRecognizer.Models;
using luval.ai.docs.msft.analyzer;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace luval.ai.docs.invoices
{
    public class InvoiceProcessor
    {
        public InvoiceProcessor(IInvoiceAnalyzer invoiceAnalyzer)
        {
            Analyzer = invoiceAnalyzer;
        }

        protected virtual IInvoiceAnalyzer Analyzer { get; private set; }



        public async Task AnalyzeDocumentToExcelAsync(Stream document, string excelOutputFile, CancellationToken cancellationToken)
        {
            var result = await Analyzer.ExecuteAsync(document, cancellationToken);
            CreateExcelFromResult(result.ToList(), excelOutputFile);
        }

        public void AnalyzeDocumentToExcel(Stream document, string excelOutputFile)
        {
            var res = AnalyzeDocumentToExcelAsync(document, excelOutputFile, CancellationToken.None);
            res.Wait();
        }

        private void CreateExcelFromResult(IList<RecognizedForm> forms, string excelFileName)
        {
            var file = new FileInfo(excelFileName);
            using (var package = new ExcelPackage(file))
            {
                for (int i = 0; i < forms.Count; i++)
                {
                    var sheet = package.Workbook.Worksheets.Add($"Document-{ i + 1 }");
                    CreateExcelSheetFromForm(forms[i], sheet, i + 1);
                }
                // Save to file
                package.Save();
            }
        }

        private void CreateExcelSheetFromForm(RecognizedForm form, ExcelWorksheet sheet, int count)
        {
            //Headers
            sheet.Cells[1, 1].Value = "Form Type";
            sheet.Cells[1, 2].Value = "Form Type Confidence";
            sheet.Cells[1, 3].Value = "Model Id";
            sheet.Cells[1, 4].Value = "Pages";
            sheet.Cells[1, 5].Value = "Key";
            sheet.Cells[1, 6].Value = "Name";
            sheet.Cells[1, 7].Value = "Label Data Text";
            sheet.Cells[1, 8].Value = "Label Data Page";
            sheet.Cells[1, 9].Value = "Label Data Box";
            sheet.Cells[1, 10].Value = "Value Type";
            sheet.Cells[1, 11].Value = "Value";
            sheet.Cells[1, 12].Value = "Value Data Text";
            sheet.Cells[1, 13].Value = "Value Box";
            var row = 2;
            foreach (var field in form.Fields)
            {
                sheet.Cells[row, 1].Value = form.FormType;
                sheet.Cells[row, 2].Value = form.FormTypeConfidence;
                sheet.Cells[row, 3].Value = form.ModelId;
                sheet.Cells[row, 4].Value = form.Pages.Count;
                sheet.Cells[row, 5].Value = field.Key;
                sheet.Cells[row, 6].Value = field.Value.Name;
                sheet.Cells[row, 7].Value = field.Value.LabelData.Text;
                sheet.Cells[row, 8].Value = field.Value.LabelData.PageNumber;
                sheet.Cells[row, 9].Value = field.Value.LabelData.BoundingBox.ToString();
                sheet.Cells[row, 10].Value = field.Value.Value.ValueType.ToString();
                sheet.Cells[row, 11].Value = CastFieldValue(field.Value.Value);
                sheet.Cells[row, 12].Value = field.Value.ValueData.Text;
                sheet.Cells[row, 13].Value = field.Value.ValueData.BoundingBox.ToString();
                row++;
            }
            sheet.Tables.Add(sheet.SelectedRange[1, 13, row - 1, 13], $"DocHeader-{count}");
        }

        private object CastFieldValue(FieldValue value)
        {
            switch (value.ValueType)
            {
                case FieldValueType.Date:
                    return value.AsDate();
                case FieldValueType.Time:
                    return value.AsTime();
                case FieldValueType.Float:
                    return value.AsFloat();
                case FieldValueType.Int64:
                    return value.AsInt64();
                default:
                    return value.AsString();
            }
        }
    }
}
