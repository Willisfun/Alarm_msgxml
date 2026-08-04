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
        // PICTURE Screen XML 完整路徑
        private string screenXmlPath = "";

        // Alarm Excel 完整路徑
        private string excelPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 選擇 Alarm Excel。
        /// </summary>
        private void open_file(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "請選擇 Alarm Excel";
                dialog.Filter = "Excel 檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                excelPath = dialog.FileName;
            }
        }

        /// <summary>
        /// 選擇 PICTURE Screen XML。
        /// </summary>
        private void open_xml(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "請選擇 PICTURE Screen XML";
                dialog.Filter = "XML 檔案 (*.xml)|*.xml|所有檔案 (*.*)|*.*";
                if (Directory.Exists(@"D:\Fanuc_Picture_project\PICTURE_ORIGIN_V1.0.0.0\KAFO_PICTURE_IHMI"))
                {
                    dialog.InitialDirectory =@"D:\Fanuc_Picture_project\PICTURE_ORIGIN_V1.0.0.0\KAFO_PICTURE_IHMI";
                }

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    screenXmlPath = dialog.FileName;

                    List<string> pairedPaths =
                        GetPairedScreenXmlPaths(screenXmlPath);

                    MessageBox.Show(
                        "已成功配對以下兩個 PICTURE 螢幕：\n\n" +
                        "【螢幕 1】\n" +
                        pairedPaths[0] +
                        "\n\n" +
                        "【螢幕 2】\n" +
                        pairedPaths[1] +
                        "\n\n執行時將同時更新這兩個螢幕。",
                        "Screen XML 配對成功",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    screenXmlPath = "";

                    MessageBox.Show(
                        "Screen XML 配對失敗：\n\n" +
                        ex.Message,
                        "配對失敗",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// 一鍵完成：
        /// 1. 驗證輸入
        /// 2. 找到 Screen XML 同層的 VtsData
        /// 3. Excel 轉 Symbol / Comment MSGXML
        /// 4. 修改 Screen XML 內的 MSGXML 路徑
        /// </summary>
        private void execute(object sender, EventArgs e)
        {
            try
            {
                ValidateInputs();

                string baseFileName = NormalizeBaseFileName(file_name.Text);
                // 使用者只要選 alarm_1 或 alarm_2 任一檔案，
                // 程式會自動找到同資料夾內的另一個配對檔案。
                List<string> screenXmlPaths = GetPairedScreenXmlPaths(screenXmlPath);
                string vtsDataFolder = GetVtsDataFolder(screenXmlPath);
                string commentTemplatePath = Path.Combine(Application.StartupPath, "comment_example.msgxml");

                string symbolTemplatePath = Path.Combine(Application.StartupPath, "symbol_example.msgxml");

                ValidateTemplate(commentTemplatePath, "Comment");
                ValidateTemplate(symbolTemplatePath, "Symbol");

                string commentOutputName = baseFileName + "_comment.msgxml";
                string symbolOutputName = baseFileName + "_symbol.msgxml";

                string commentOutputPath = Path.Combine(
                    vtsDataFolder,
                    commentOutputName);

                string symbolOutputPath = Path.Combine(
                    vtsDataFolder,
                    symbolOutputName);

                ConfirmOverwriteIfNeeded(
                    commentOutputPath,
                    symbolOutputPath);

                // 先產生檔案；兩個都成功後才修改 Screen XML。
                ConversionResult commentResult = ConvertExcelToMsgXml(
                    excelPath,
                    commentTemplatePath,
                    commentOutputPath,
                    MessageSource.Comment);

                ConversionResult symbolResult = ConvertExcelToMsgXml(
                    excelPath,
                    symbolTemplatePath,
                    symbolOutputPath,
                    MessageSource.Symbol);

                List<PictureUpdateResult> pictureResults =
                UpdatePictureXmlFiles(
                    screenXmlPaths,
                    symbolOutputName,
                    commentOutputName);

                PictureUpdateResult alarm1Result = pictureResults[0];
                PictureUpdateResult alarm2Result = pictureResults[1];

                int totalSymbolCount = alarm1Result.SymbolCount + alarm2Result.SymbolCount;

                int totalCommentCount = alarm1Result.CommentCount + alarm2Result.CommentCount;

                MessageBox.Show(
                    "Alarm 建置完成！\n\n" +

                    "【" + Path.GetFileName(screenXmlPaths[0]) + "】\n" +
                    "SYMBOL：修改 " +
                    alarm1Result.SymbolCount +
                    " 個\n" +
                    "COMMENT：修改 " +
                    alarm1Result.CommentCount +
                    " 個\n\n" +

                    "【" + Path.GetFileName(screenXmlPaths[1]) + "】\n" +
                    "SYMBOL：修改 " +
                    alarm2Result.SymbolCount +
                    " 個\n" +
                    "COMMENT：修改 " +
                    alarm2Result.CommentCount +
                    " 個\n\n" +

                    "【修改總計】\n" +
                    "SYMBOL：共修改 " +
                    totalSymbolCount +
                    " 個\n" +
                    "COMMENT：共修改 " +
                    totalCommentCount +
                    " 個\n\n" +

                    "【Comment MSGXML】\n" +
                    "繁體中文：" +
                    commentResult.ChineseCount +
                    " 筆\n" +
                    "英文：" +
                    commentResult.EnglishCount +
                    " 筆\n" +
                    "簡體中文：" +
                    commentResult.SimplifiedChineseCount +
                    " 筆\n\n" +

                    "【Symbol MSGXML】\n" +
                    "英文：" +
                    symbolResult.EnglishCount +
                    " 筆\n\n" +

                    "【輸出資料夾】\n" +
                    vtsDataFolder,

                    "完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                // 使用者取消覆蓋，不顯示錯誤。
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "建置失敗：" + ex.Message,
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(screenXmlPath))
            {
                throw new InvalidOperationException("請先選擇 PICTURE Screen XML。");
            }

            if (!File.Exists(screenXmlPath))
            {
                throw new FileNotFoundException(
                    "找不到 PICTURE Screen XML。",
                    screenXmlPath);
            }

            if (string.IsNullOrWhiteSpace(excelPath))
            {
                throw new InvalidOperationException("請先選擇 Alarm Excel。");
            }

            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException(
                    "找不到 Alarm Excel。",
                    excelPath);
            }
        }

        private string NormalizeBaseFileName(string input)
        {
            string baseFileName = (input ?? "").Trim();

            if (baseFileName.EndsWith(
                    ".msgxml",
                    StringComparison.OrdinalIgnoreCase))
            {
                baseFileName = Path.GetFileNameWithoutExtension(baseFileName);
            }

            if (string.IsNullOrWhiteSpace(baseFileName))
            {
                throw new InvalidOperationException("請輸入有效的輸出檔名。");
            }

            if (baseFileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new InvalidOperationException(
                    "輸出檔名包含 Windows 不允許的字元。");
            }

            return baseFileName;
        }

        /// <summary>
        /// Alarm XML 與 VtsData 資料夾位於同一層。
        /// 例如：Project\Alarm.xml -> Project\VtsData
        /// </summary>
        private string GetVtsDataFolder(string xmlPath)
        {
            string xmlFolder = Path.GetDirectoryName(xmlPath);

            if (string.IsNullOrWhiteSpace(xmlFolder))
            {
                throw new InvalidOperationException(
                    "無法取得 Alarm XML 所在資料夾。");
            }

            return Path.Combine(xmlFolder, "VtsData");
        }

        private void ValidateTemplate(string templatePath, string templateName)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    "找不到 " + templateName + " MSGXML 範本。" +
                    "請將範本放在程式執行檔旁邊。",
                    templatePath);
            }
        }

        private void ConfirmOverwriteIfNeeded(params string[] outputPaths)
        {
            List<string> existingFiles = outputPaths.Where(File.Exists).ToList();

            if (existingFiles.Count == 0)
            {
                return;
            }

            DialogResult answer = MessageBox.Show(
                "以下檔案已存在：" +
                string.Join("", existingFiles) + "是否覆蓋？",
                "確認覆蓋",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (answer != DialogResult.Yes)
            {
                throw new OperationCanceledException();
            }
        }

        /// <summary>
        /// 根據使用者選到的 alarm_1 / alarm_2，
        /// 找到同一機型的另一個 Screen XML。
        ///
        /// 例如：
        /// RVN_alarm_1.xml -> RVN_alarm_2.xml
        /// RVN_alarm_2.xml -> RVN_alarm_1.xml
        /// </summary>
        private List<string> GetPairedScreenXmlPaths(
            string selectedXmlPath)
        {
            string selectedFullPath = Path.GetFullPath(selectedXmlPath);

            string folder = Path.GetDirectoryName(selectedFullPath);

            string fileName = Path.GetFileName(selectedFullPath);

            if (string.IsNullOrWhiteSpace(folder))
            {
                throw new InvalidOperationException(
                    "無法取得 Screen XML 所在資料夾。");
            }

            string pairedFileName;

            if (Regex.IsMatch(fileName, "alarm_1", RegexOptions.IgnoreCase))
            {
                pairedFileName = Regex.Replace(
                    fileName,
                    "alarm_1",
                    "ALARM_2",
                    RegexOptions.IgnoreCase);
            }
            else if (Regex.IsMatch(fileName, "alarm_2", RegexOptions.IgnoreCase))
            {
                pairedFileName = Regex.Replace(
                    fileName,
                    "alarm_2",
                    "ALARM_1",
                    RegexOptions.IgnoreCase);
            }
            else
            {
                throw new InvalidOperationException("選擇的 Screen XML 檔名必須包含 ALARM_1 或 ALARM_2。\n\n" + "目前檔名：" + fileName);
            }

            string pairedXmlPath = Path.Combine(folder, pairedFileName);

            if (!File.Exists(pairedXmlPath))
            {
                throw new FileNotFoundException(
                    "找不到配對的 Screen XML。\n\n" +
                    "已選擇：\n" +
                    selectedFullPath +
                    "\n\n應存在：\n" +
                    pairedXmlPath,
                    pairedXmlPath);
            }

            // 固定依 alarm_1、alarm_2 排序，
            // 讓畫面訊息與執行順序一致。
            return new[]
                {
                    selectedFullPath,
                    Path.GetFullPath(pairedXmlPath)
                }
                .OrderBy(path => Regex.IsMatch(Path.GetFileName(path),"alarm_1",RegexOptions.IgnoreCase)? 1 : 2).ToList();
        }

        /// <summary>
        /// 對同一機型的 alarm_1 與 alarm_2
        /// 套用完全相同的 MSGXML 路徑修改。
        /// </summary>
        private List<PictureUpdateResult> UpdatePictureXmlFiles(
     IEnumerable<string> xmlPaths,
     string symbolFileName,
     string commentFileName)
        {
            List<PictureUpdateResult> results =
                new List<PictureUpdateResult>();

            foreach (string xmlPath in xmlPaths)
            {
                PictureUpdateResult result =
                    UpdatePictureXml(
                        xmlPath,
                        symbolFileName,
                        commentFileName);

                results.Add(result);
            }

            return results;
        }

        private PictureUpdateResult UpdatePictureXml(
            string xmlPath,
            string symbolFileName,
            string commentFileName)
        {
            XDocument document = XDocument.Load(
                xmlPath,
                System.Xml.Linq.LoadOptions.PreserveWhitespace);

            string newSymbolPath = @"VtsData\" + symbolFileName;
            string newCommentPath = @"VtsData\" + commentFileName;

            List<XElement> symbolTargets = document
                .Descendants()
                .Where(element =>
                    element.Name.LocalName == "VTS" &&
                    element.Value.IndexOf(
                        "symbol.msgxml",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            List<XElement> commentTargets = document
                .Descendants()
                .Where(element =>
                    (element.Name.LocalName == "FileName" ||
                     element.Name.LocalName == "VTS") &&
                    element.Value.IndexOf(
                        "comment.msgxml",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            if (symbolTargets.Count == 0 && commentTargets.Count == 0)
            {
                throw new InvalidOperationException(
                    "PICTURE Screen XML 中找不到 SYMBOL 或 COMMENT 的 MSGXML 引用。");
            }

            foreach (XElement element in symbolTargets)
            {
                element.Value = newSymbolPath;
            }

            foreach (XElement element in commentTargets)
            {
                element.Value = newCommentPath;
            }

            document.Save(xmlPath, System.Xml.Linq.SaveOptions.DisableFormatting);

            return new PictureUpdateResult
            {
                SymbolCount = symbolTargets.Count,
                CommentCount = commentTargets.Count
            };
        }

        private ConversionResult ConvertExcelToMsgXml(
    string sourceExcelPath,
    string templateMsgXmlPath,
    string outputMsgXmlPath,
    MessageSource messageSource)
        {
            ConversionResult result = new ConversionResult();

            using (XLWorkbook workbook = new XLWorkbook(sourceExcelPath))
            {
                XDocument document = XDocument.Load(templateMsgXmlPath, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                if (messageSource == MessageSource.Symbol)
                {
                    /*
                     * Symbol 是單語言格式：
                     *
                     * Excel：工作表名稱只要包含 EN
                     * XML：Sheet1
                     */
                    IXLWorksheet englishExcelSheet = FindEnglishWorksheet(workbook);

                    result.EnglishCount =
                        UpdateLanguageSheet(
                            englishExcelSheet,
                            document,
                            "Sheet1",
                            MessageSource.Symbol
                            );
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
                            MessageSource.Comment
                            );

                    result.EnglishCount =
                        UpdateLanguageSheet(
                            englishExcelSheet,
                            document,
                            "English",
                            MessageSource.Comment
                             );

                    result.SimplifiedChineseCount =
                        UpdateLanguageSheet(
                            simplifiedChineseExcelSheet,
                            document,
                            "SimplifiedChinese",
                            MessageSource.Comment
                            );
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
            MessageSource messageSource
            )
        {
            XElement xmlSheet = document
                .Descendants()
                .FirstOrDefault(element =>
                {
                    if (element.Name.LocalName != "Sheet")
                    {
                        return false;
                    }

                    string actualSheetName = GetAttributeValue(element, "name");

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

        private string ConvertExcelLineBreaksToMsgXml(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string normalized = text
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");

            // 如果換行後下一行是空白開頭，代表只是排版，直接接起來
            normalized = System.Text.RegularExpressions.Regex.Replace(
                normalized,
                @"\n\s+",
                "");
            // 剩下真正的換行才轉成 \n
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
            IXLWorksheet worksheet = workbook.Worksheet("ALARM");

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

    public enum MessageSource
    {
        Comment,
        Symbol
    }

    public class ConversionResult
    {
        public int ChineseCount { get; set; }
        public int EnglishCount { get; set; }
        public int SimplifiedChineseCount { get; set; }
    }

    public class PictureUpdateResult
    {
        public int SymbolCount { get; set; }
        public int CommentCount { get; set; }
    }
}