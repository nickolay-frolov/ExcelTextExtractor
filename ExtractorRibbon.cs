using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTextExtractor
{
    public partial class ExtractorRibbon
    {
        private void ExtractorRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ExtractBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MergeAndCopyText();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MergeAndCopyText()
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection;

            if (selectedRange != null && selectedRange.Cells.Count > 1)
            {
                string mergedText = GetMergedText(selectedRange);
                Excel.Range targetCell = selectedRange.Cells[selectedRange.Cells.Count].Offset[1, 0] as Excel.Range;

                if (!string.IsNullOrEmpty(mergedText))
                {
                    targetCell.Value = mergedText;
                    
                    targetCell.WrapText = true;
                    targetCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    targetCell.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                    Clipboard.SetText(mergedText);
                }
                else
                {
                    MessageBox.Show("Выбранный диапазон не содержит текст.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите более одной ячейки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string GetMergedText(Excel.Range range)
        {
            StringBuilder mergedSB = new StringBuilder();

            foreach (Excel.Range cell in range.Cells)
            {
                string cellText = cell.Text.Trim();

                if (!string.IsNullOrEmpty(cellText))
                {
                    cellText = char.ToUpper(cellText[0]) + cellText.Substring(1);
                }

                if (!string.IsNullOrEmpty(cellText) && !cellText.EndsWith("."))
                {
                    cellText += ".";
                }

                mergedSB.Append(cellText + " ");
            }

            return mergedSB.ToString().TrimEnd();
        }
    }
}
