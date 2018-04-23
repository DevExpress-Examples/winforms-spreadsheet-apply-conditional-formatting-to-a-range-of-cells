using System;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using System.Diagnostics;

namespace ConditionalFormatting_Examples
{
    public partial class Form1 : Form
    {

        IWorkbook workbook;

        public Form1()
        {
            InitializeComponent();

            //  Access a workbook.
            workbook = spreadsheetControl1.Document;
            InitTreeListControl();

        }

        private void InitTreeListControl()
        {
            GroupsOfSpreadsheetExamples examples = new GroupsOfSpreadsheetExamples();
            InitData(examples);
            DataBinding(examples);
        }

        private void InitData(GroupsOfSpreadsheetExamples examples)
        {
            #region GroupNodes
            examples.Add(new SpreadsheetNode("Conditional Formatting Examples"));
            #endregion

            #region ExampleNodes
            // Add nodes to the "Conditional Formatting" group of examples.
            examples[0].Groups.Add(new SpreadsheetExample("Format values that are above or below average", ConditionalFormatting.AddAverageConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells that are not between two specified values", ConditionalFormatting.AddRangeConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format top ranked values", ConditionalFormatting.AddRankConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells that contain the given text", ConditionalFormatting.AddTextConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format only unique values", ConditionalFormatting.AddSpecialConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format today's date", ConditionalFormatting.AddTimePeriodConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format values that are greater than a specified value", ConditionalFormatting.AddExpressionConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Use a formula to determine which cells to format", ConditionalFormatting.AddFormulaExpressionConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells by using a two-color scale", ConditionalFormatting.AddColorScale2ConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells by using a three-color scale", ConditionalFormatting.AddColorScale3ConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells by using data bars", ConditionalFormatting.AddDataBarConditionalFormattingAction));
            examples[0].Groups.Add(new SpreadsheetExample("Format cells by using a custom icon set", ConditionalFormatting.AddIconSetConditionalFormattingAction));
            #endregion
        }

        void DataBinding(GroupsOfSpreadsheetExamples examples)
        {
            treeList1.DataSource = examples;
            treeList1.ExpandAll();
            treeList1.BestFitColumns();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadDocumentFromFile();
            SpreadsheetExample example = treeList1.GetDataRecordByNode(treeList1.FocusedNode) as SpreadsheetExample;
            if (example == null)
                return;
            Action<IWorkbook> action = example.Action;
            action(workbook);
            SaveDocumentToFile();
        }

        // ------------------- Load and Save a Document -------------------
        private void LoadDocumentFromFile() {
            #region #LoadDocumentFromFile
            // Load a workbook from a file.
            workbook.LoadDocument("Documents\\Document.xlsx", DocumentFormat.OpenXml);
            #endregion #LoadDocumentFromFile
        }

        private void SaveDocumentToFile() {
            #region #SaveDocumentToFile
            // Save the modified document to a file.
            workbook.SaveDocument("Documents\\SavedDocument.xlsx", DocumentFormat.OpenXml);
            #endregion #SaveDocumentToFile
        }
    }
}
