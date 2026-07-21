using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace Alarm_to_msgxml
{
    public partial class Form1 : Form
    {
        // 使用者選擇的 Excel 完整路徑
        private string excelPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 選擇 Excel
        /// </summary>
        private void open_file(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "請選擇 Excel";
                dialog.Filter =
                    "Excel 檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    excelPath = dialog.FileName;

                    MessageBox.Show(
                        "已選擇：\n" + excelPath,
                        "完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }

        /// <summary>
        /// 執行 Excel → Comment MSGXML + Symbol MSGXML
        /// </summary>
        private void execute(object sender, EventArgs e)
        {
            // 檢查 Excel 是否已選擇
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                MessageBox.Show(
                    "請先選擇 Excel！",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // 檢查 Excel 是否存在
            if (!File.Exists(excelPath))
            {
                MessageBox.Show(
                    "找不到選擇的 Excel：\n" + excelPath,
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            // 檢查輸出檔名
            if (string.IsNullOrWhiteSpace(file_name.Text))
            {
                MessageBox.Show(
                    "請輸入輸出檔名！",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            string baseFileName = file_name.Text.Trim();

            // 若使用者誤輸入 .msgxml，先移除副檔名
            if (baseFileName.EndsWith(
                    ".msgxml",
                    StringComparison.OrdinalIgnoreCase))
            {
                baseFileName = baseFileName.Substring(
                    0,
                    baseFileName.Length - ".msgxml".Length);
            }

            // 檢查移除副檔名後是否為空
            if (string.IsNullOrWhiteSpace(baseFileName))
            {
                MessageBox.Show(
                    "請輸入有效的輸出檔名！",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // 避免輸入不能作為檔名的字元
            if (baseFileName.IndexOfAny(
                    Path.GetInvalidFileNameChars()) >= 0)
            {
                MessageBox.Show(
                    "輸出檔名包含無效字元：\n" +
                    string.Join(
                        " ",
                        Path.GetInvalidFileNameChars()),
                    "檔名錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            /*
             * 範本路徑
             *
             * comment_example.msgxml
             * symbol_example.msgxml
             *
             * 都必須放在 exe 旁邊
             */
            string commentTemplatePath = Path.Combine(
                Application.StartupPath,
                "comment_example.msgxml");

            string symbolTemplatePath = Path.Combine(
                Application.StartupPath,
                "symbol_example.msgxml");

            // 檢查 Comment 範本
            if (!File.Exists(commentTemplatePath))
            {
                MessageBox.Show(
                    "找不到 Comment MSGXML 範本：\n\n" +
                    commentTemplatePath +
                    "\n\n請將 comment_example.msgxml " +
                    "放在程式執行檔旁邊。",
                    "缺少範本",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            // 檢查 Symbol 範本
            if (!File.Exists(symbolTemplatePath))
            {
                MessageBox.Show(
                    "找不到 Symbol MSGXML 範本：\n\n" +
                    symbolTemplatePath +
                    "\n\n請將 symbol_example.msgxml " +
                    "放在程式執行檔旁邊。",
                    "缺少範本",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return;
            }

            /*
             * 輸出檔名：
             *
             * 使用者輸入：
             * KRV_ALARM_
             *
             * 輸出：
             * KRV_ALARM_comment.msgxml
             * KRV_ALARM_symbol.msgxml
             */
            string commentOutputName =
                baseFileName + "_comment.msgxml";

            string symbolOutputName =
                baseFileName + "_symbol.msgxml";

            // 輸出到 exe 同一個資料夾
            string commentOutputPath = Path.Combine(
                Application.StartupPath,
                commentOutputName);

            string symbolOutputPath = Path.Combine(
                Application.StartupPath,
                symbolOutputName);

            // 避免 Comment 範本和輸出是同一個檔案
            if (IsSamePath(
                    commentTemplatePath,
                    commentOutputPath))
            {
                MessageBox.Show(
                    "Comment 輸出檔名不能和範本相同。\n\n" +
                    "請輸入其他檔名。",
                    "檔名重複",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // 避免 Symbol 範本和輸出是同一個檔案
            if (IsSamePath(
                    symbolTemplatePath,
                    symbolOutputPath))
            {
                MessageBox.Show(
                    "Symbol 輸出檔名不能和範本相同。\n\n" +
                    "請輸入其他檔名。",
                    "檔名重複",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            // 檢查輸出檔案是否已存在
            List<string> existingFiles = new List<string>();

            if (File.Exists(commentOutputPath))
            {
                existingFiles.Add(commentOutputPath);
            }

            if (File.Exists(symbolOutputPath))
            {
                existingFiles.Add(symbolOutputPath);
            }

            // 只詢問一次是否覆蓋
            if (existingFiles.Count > 0)
            {
                string existingFileText =
                    string.Join("\n\n", existingFiles);

                DialogResult answer = MessageBox.Show(
                    "以下輸出檔案已存在：\n\n" +
                    existingFileText +
                    "\n\n是否覆蓋？",
                    "確認覆蓋",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (answer != DialogResult.Yes)
                {
                    return;
                }
            }

            try
            {
                /*
                 * 產生 Comment MSGXML
                 *
                 * 使用 Excel 第 3 欄 Comment
                 */
                ConversionResult commentResult =
                    ConvertExcelToMsgXml(
                        excelPath,
                        commentTemplatePath,
                        commentOutputPath,
                        MessageSource.Comment);

                /*
                 * 產生 Symbol MSGXML
                 *
                 * 使用 Excel 第 2 欄 Symbol
                 */
                ConversionResult symbolResult =
                    ConvertExcelToMsgXml(
                        excelPath,
                        symbolTemplatePath,
                        symbolOutputPath,
                        MessageSource.Symbol);

                MessageBox.Show(
                    "轉換完成！\n\n" +

                    "【Comment】\n" +
                    "繁體中文：" +
                    commentResult.ChineseCount + " 筆\n" +
                    "英文：" +
                    commentResult.EnglishCount + " 筆\n" +
                    "簡體中文：" +
                    commentResult.SimplifiedChineseCount +
                    " 筆\n\n" +

                    "【Symbol】\n" +
                    "英文：" +
                    symbolResult.EnglishCount + " 筆\n" +
                    " 筆\n\n" +
                    "輸出位置：\n\n" +
                    commentOutputPath + "\n\n" +
                    symbolOutputPath,
                    "轉換成功",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "轉換失敗：\n\n" + ex.Message,
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 執行完整轉換
        /// </summary>
        private ConversionResult ConvertExcelToMsgXml(
    string sourceExcelPath,
    string templateMsgXmlPath,
    string outputMsgXmlPath,
    MessageSource messageSource)
        {
            ConversionResult result = new ConversionResult();

            using (XLWorkbook workbook =
                   new XLWorkbook(sourceExcelPath))
            {
                XDocument document = XDocument.Load(
                    templateMsgXmlPath,
                    System.Xml.Linq.LoadOptions.PreserveWhitespace);

                if (messageSource == MessageSource.Symbol)
                {
                    /*
                     * Symbol 是單語言格式：
                     *
                     * Excel：工作表名稱只要包含 EN
                     * XML：Sheet1
                     */
                    IXLWorksheet englishExcelSheet =
                        FindEnglishWorksheet(workbook);

                    result.EnglishCount =
                        UpdateLanguageSheet(
                            englishExcelSheet,
                            document,
                            "Sheet1",
                            MessageSource.Symbol,
                            false);
                }
                else
                {
                    /*
                     * Comment 是多語言格式
                     */
                    IXLWorksheet traditionalChineseExcelSheet =
                        FindTraditionalChineseWorksheet(workbook);

                    IXLWorksheet englishExcelSheet =
                        FindEnglishWorksheet(workbook);

                    IXLWorksheet simplifiedChineseExcelSheet =
                        FindSimplifiedChineseWorksheet(workbook);

                    result.ChineseCount =
                        UpdateLanguageSheet(
                            traditionalChineseExcelSheet,
                            document,
                            "Chinese",
                            MessageSource.Comment,
                            false);

                    result.EnglishCount =
                        UpdateLanguageSheet(
                            englishExcelSheet,
                            document,
                            "EN",
                            MessageSource.Comment,
                            true);

                    result.SimplifiedChineseCount =
                        UpdateLanguageSheet(
                            simplifiedChineseExcelSheet,
                            document,
                            "SimplifiedChinese",
                            MessageSource.Comment,
                            false);
                }

                SaveAsUtf16(document, outputMsgXmlPath);
            }
            return result;
        }

        /// <summary>
        /// 更新其中一個語言 Sheet
        /// </summary>
        private int UpdateLanguageSheet(
            IXLWorksheet excelSheet,
            XDocument document,
            string xmlSheetName,
            MessageSource messageSource,
            bool xmlSheetNameContains)
        {
            XElement xmlSheet = document
                .Descendants()
                .FirstOrDefault(element =>
                {
                    if (element.Name.LocalName != "Sheet")
                    {
                        return false;
                    }

                    string actualSheetName =
                        GetAttributeValue(element, "name");

                    if (xmlSheetNameContains)
                    {
                        return actualSheetName.IndexOf(
                            xmlSheetName,
                            StringComparison.OrdinalIgnoreCase) >= 0;
                    }

                    return string.Equals(
                        actualSheetName,
                        xmlSheetName,
                        StringComparison.OrdinalIgnoreCase);
                });

            if (xmlSheet == null)
            {
                throw new InvalidOperationException(
                    "MSGXML 找不到語言表：" +
                    xmlSheetName);
            }

            // 建立 Bit → Row 快速對照表
            Dictionary<int, XElement> xmlRowsByBit =
                new Dictionary<int, XElement>();

            foreach (XElement row in xmlSheet.Elements()
                         .Where(element =>
                             element.Name.LocalName == "Row"))
            {
                XElement bitElement = row.Elements()
                    .FirstOrDefault(element =>
                        element.Name.LocalName == "Bit");

                if (bitElement == null)
                {
                    continue;
                }

                int bitValue;

                if (!int.TryParse(
                        bitElement.Value,
                        out bitValue))
                {
                    continue;
                }

                if (xmlRowsByBit.ContainsKey(bitValue))
                {
                    throw new InvalidOperationException(
                        "MSGXML 的 " +
                        xmlSheetName +
                        " 出現重複 Bit：" +
                        bitValue);
                }

                xmlRowsByBit.Add(bitValue, row);
            }

            IXLRow lastRow = excelSheet.LastRowUsed();

            if (lastRow == null)
            {
                throw new InvalidOperationException(
                    "Excel 工作表「" +
                    excelSheet.Name +
                    "」沒有資料。");
            }

            int lastRowNumber =
                lastRow.RowNumber();

            int updateCount = 0;

            /*
             * Excel 格式：
             *
             * 第 1 列：標題
             * 第 2 列：AlarmMessage
             * 第 3 列：address / symbol / comment / status
             * 第 4 列開始：警報內容
             */
            for (int rowNumber = 4; rowNumber <= lastRowNumber; rowNumber++)
            {
                string address = excelSheet
                    .Cell(rowNumber, 1)
                    .GetString()
                    .Trim();

                string symbol = excelSheet
                    .Cell(rowNumber, 2)
                    .GetString();

                string comment = excelSheet
                    .Cell(rowNumber, 3)
                    .GetString();

                /*
                 * 根據輸出檔案種類決定要寫入的內容
                 *
                 * Symbol：Excel 第 2 欄
                 * Comment：Excel 第 3 欄
                 */
                string sourceText;

                if (messageSource == MessageSource.Symbol)
                {
                    sourceText = symbol;
                }
                else
                {
                    sourceText = comment;
                }

                // Address 和要輸出的文字都空白時，不處理
                if ((string.IsNullOrWhiteSpace(comment)) || (string.IsNullOrWhiteSpace(address) && string.IsNullOrWhiteSpace(sourceText)))
                {
                    continue;
                }

                // 有文字但是沒有 Address
                if (string.IsNullOrWhiteSpace(address))
                {
                    throw new InvalidOperationException(
                        "工作表「" +
                        excelSheet.Name +
                        "」第 " +
                        rowNumber +
                        " 列缺少 Address。\n\n" +
                        "內容：\n" +
                        sourceText);
                }

                /*
                 * Comment 或 Symbol 空白時，
                 * 保留原本 XML 的 Message，不清除。
                 */
                if (string.IsNullOrWhiteSpace(sourceText))
                {
                    continue;
                }

                int bitValue =
                    ConvertAddressToBit(address);

                XElement xmlRow;

                if (!xmlRowsByBit.TryGetValue(
                        bitValue,
                        out xmlRow))
                {
                    throw new InvalidOperationException(
                        "找不到 XML 對應列。\n\n" +
                        "Excel 工作表：" +
                        excelSheet.Name + "\n" +
                        "Excel 列號：" +
                        rowNumber + "\n" +
                        "Address：" +
                        address + "\n" +
                        "換算 Bit：" +
                        bitValue + "\n" +
                        "內容：" +
                        sourceText);
                }

                XElement messageElement =
                    xmlRow.Elements()
                        .FirstOrDefault(element =>
                            element.Name.LocalName ==
                            "Message");

                if (messageElement == null)
                {
                    throw new InvalidOperationException(
                        "MSGXML 的 Bit " +
                        bitValue +
                        " 缺少 <Message> 節點。");
                }

                string convertedText;

                if (messageSource == MessageSource.Symbol)
                {
                    /*
                     * Symbol 要轉成單行。
                     *
                     * 例如：
                     *
                     * 1002
                     * SPINDLE DRIVER ALARM
                     *
                     * 轉成：
                     *
                     * 1002 SPINDLE DRIVER ALARM
                     */
                    convertedText =
                        ConvertSymbolToSingleLine(sourceText);
                }
                else
                {
                    /*
                     * Comment 保留原本處理方式：
                     * 將 Excel 實際換行轉成字面 \n
                     */
                    convertedText =
                        ConvertExcelLineBreaksToMsgXml(
                            sourceText);
                }

                messageElement.Value = convertedText;

                updateCount++;
            }

            return updateCount;
        }

        /// <summary>
        /// 將 Symbol 內容轉成單行
        /// </summary>
        private string ConvertSymbolToSingleLine(
            string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            /*
             * \s+ 會比對：
             *
             * 換行
             * Tab
             * 多個連續空格
             *
             * 全部轉成一個普通空格
             */
            string singleLine = Regex.Replace(
                text,
                @"\s+",
                " ");

            return $"\"{singleLine}\"";
        }

        /// <summary>
        /// A24.4 → 24 × 8 + 4 → Bit 196
        /// </summary>
        private int ConvertAddressToBit(
            string address)
        {
            Match match = Regex.Match(
                address.Trim(),
                @"^A(\d+)\.(\d+)$",
                RegexOptions.IgnoreCase);

            if (!match.Success)
            {
                throw new FormatException(
                    "Address 格式錯誤：" +
                    address +
                    "\n\n正確格式例如：A1.0、A24.4");
            }

            int addressNumber = int.Parse(
                match.Groups[1].Value);

            int bitNumber = int.Parse(
                match.Groups[2].Value);

            if (bitNumber < 0 || bitNumber > 7)
            {
                throw new FormatException(
                    "Address 的小數點後數字必須介於 0～7：" +
                    address);
            }

            return addressNumber * 8 + bitNumber;
        }

        /// <summary>
        /// 將 Excel 的實際換行轉成 MSGXML 使用的字面 \n，
        /// 並清除每一行前後的多餘空白。
        /// </summary>
        private string ConvertExcelLineBreaksToMsgXml(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string normalized = text
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            /* * "\n" 代表實際換行。 * "\\n" 代表反斜線加英文字母 n。 */
            return normalized.Replace("\n", "\\n");
        }

        /// <summary>
        /// 使用 UTF-16 儲存 MSGXML
        /// </summary>
        private void SaveAsUtf16(
            XDocument document,
            string outputPath)
        {
            XmlWriterSettings settings =
                new XmlWriterSettings();

            // UTF-16 Little Endian
            settings.Encoding = Encoding.Unicode;

            // 不自行美化排版
            settings.Indent = false;

            settings.OmitXmlDeclaration = false;

            // 不修改 Message 中的換行內容
            settings.NewLineHandling =
                NewLineHandling.None;

            using (XmlWriter writer =
                   XmlWriter.Create(
                       outputPath,
                       settings))
            {
                document.Save(writer);
            }
        }

        /// <summary>
        /// 尋找英文工作表：名稱只要包含 EN 即可，不區分大小寫。
        /// 例如 ALARM (EN)、ALARM_EN、English 都能找到。
        /// </summary>
        private IXLWorksheet FindEnglishWorksheet(
            XLWorkbook workbook)
        {
            IXLWorksheet worksheet = workbook.Worksheets
                .FirstOrDefault(item =>
                    item.Name.IndexOf(
                        "EN",
                        StringComparison.OrdinalIgnoreCase) >= 0);

            if (worksheet == null)
            {
                throw new InvalidOperationException(
                    BuildWorksheetNotFoundMessage(
                        workbook,
                        "名稱包含 EN"));
            }

            return worksheet;
        }

        /// <summary>
        /// 尋找繁體中文工作表。
        /// 優先找完全等於 ALARM，避免誤抓英文或簡體工作表。
        /// </summary>
        private IXLWorksheet FindTraditionalChineseWorksheet(
            XLWorkbook workbook)
        {
            IXLWorksheet worksheet = workbook.Worksheets
                .FirstOrDefault(item =>
                    string.Equals(
                        item.Name.Trim(),
                        "ALARM",
                        StringComparison.OrdinalIgnoreCase));

            if (worksheet == null)
            {
                worksheet = workbook.Worksheets
                    .FirstOrDefault(item =>
                        item.Name.IndexOf(
                            "ALARM",
                            StringComparison.OrdinalIgnoreCase) >= 0 &&
                        item.Name.IndexOf(
                            "EN",
                            StringComparison.OrdinalIgnoreCase) < 0 &&
                        !item.Name.Contains("簡"));
            }

            if (worksheet == null)
            {
                throw new InvalidOperationException(
                    BuildWorksheetNotFoundMessage(
                        workbook,
                        "繁體中文 ALARM"));
            }

            return worksheet;
        }

        /// <summary>
        /// 尋找簡體中文工作表：名稱包含「簡」。
        /// </summary>
        private IXLWorksheet FindSimplifiedChineseWorksheet(
            XLWorkbook workbook)
        {
            IXLWorksheet worksheet = workbook.Worksheets
                .FirstOrDefault(item => item.Name.Contains("簡"));

            if (worksheet == null)
            {
                throw new InvalidOperationException(
                    BuildWorksheetNotFoundMessage(
                        workbook,
                        "名稱包含「簡」"));
            }

            return worksheet;
        }

        /// <summary>
        /// 建立工作表找不到時的詳細訊息。
        /// </summary>
        private string BuildWorksheetNotFoundMessage(
            XLWorkbook workbook,
            string condition)
        {
            string worksheetNames = string.Join(
                "\n",
                workbook.Worksheets.Select(item =>
                    "- " + item.Name));

            return
                "Excel 找不到符合條件的工作表：" +
                condition +
                "\n\n目前 Excel 內的工作表：\n" +
                worksheetNames;
        }

        /// <summary>
        /// 不區分大小寫取得 XML 屬性
        /// </summary>
        private string GetAttributeValue(
            XElement element,
            string attributeName)
        {
            XAttribute attribute =
                element.Attributes()
                    .FirstOrDefault(item =>
                        string.Equals(
                            item.Name.LocalName,
                            attributeName,
                            StringComparison.OrdinalIgnoreCase));

            return attribute == null
                ? ""
                : attribute.Value;
        }

        /// <summary>
        /// 比較兩個路徑是否指向同一個檔案
        /// </summary>
        private bool IsSamePath(
            string firstPath,
            string secondPath)
        {
            return string.Equals(
                Path.GetFullPath(firstPath),
                Path.GetFullPath(secondPath),
                StringComparison.OrdinalIgnoreCase);
        }
    }

    /// <summary>
    /// 決定 Message 使用 Excel 的哪一欄
    /// </summary>
    public enum MessageSource
    {
        Comment,
        Symbol
    }

    /// <summary>
    /// 單一 MSGXML 的轉換結果
    /// </summary>
    public class ConversionResult
    {
        public int ChineseCount { get; set; }

        public int EnglishCount { get; set; }

        public int SimplifiedChineseCount { get; set; }
    }
}