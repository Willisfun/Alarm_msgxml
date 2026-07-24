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
using DocumentFormat.OpenXml;

namespace Alarm_to_msgxml
{
    public partial class Form1 : Form
    {
        // 使用者選擇的 Excel 完整路徑
        private string xmlPath = "";

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 選擇 XML 檔案，Picture 元件檔案
        /// </summary>
        private void open_file(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "請選擇 XML 檔案";
                dialog.Filter =
                    "XML 檔案 (*.xml)|*.xml|所有檔案 (*.*)|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    xmlPath = dialog.FileName;
                    MessageBox.Show(
                        "已選擇：\n" + xmlPath,
                        "完成",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
        }
        /// <summary>
        /// 將 Picture XML 中的 SYMBOL 與 COMMENT msgxml 路徑
        /// 改成使用者輸入的新檔名
        /// </summary>
        private void execute(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(xmlPath))
            {
                MessageBox.Show(
                    "請先選擇 XML 檔案。",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            string fileName = file_name.Text.Trim();

            if (string.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show(
                    "請輸入檔案名稱。",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            try
            {
                XDocument document = XDocument.Load(xmlPath);

                string newSymbolPath =
                    @"VtsData\" + fileName + "_SYMBOL.msgxml";

                string newCommentPath =
                    @"VtsData\" + fileName + "_comment.msgxml";

                var symbolTargets = document
                    .Descendants()
                    .Where(element =>
                        (element.Name.LocalName == "VTS")
                        &&
                        element.Value.IndexOf(
                            "symbol.msgxml",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                var commentTargets = document
                    .Descendants()
                    .Where(element =>
                        (element.Name.LocalName == "FileName"
                         || element.Name.LocalName == "VTS")
                        &&
                        element.Value.IndexOf(
                            "comment.msgxml",
                            StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();

                foreach (XElement element in symbolTargets)
                {
                    element.Value = newSymbolPath;
                }

                foreach (XElement element in commentTargets)
                {
                    element.Value = newCommentPath;
                }

                if (symbolTargets.Count == 0 &&
                    commentTargets.Count == 0)
                {
                    MessageBox.Show(
                        "沒有找到任何 SYMBOL 或 COMMENT 的 msgxml 引用。",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                document.Save(xmlPath, System.Xml.Linq.SaveOptions.DisableFormatting);

                MessageBox.Show(
                    "修改完成！\n\n" +
                    $"SYMBOL：修改 {symbolTargets.Count} 個\n" +
                    $"COMMENT：修改 {commentTargets.Count} 個\n\n" +
                    $"SYMBOL 新路徑：\n{newSymbolPath}\n\n" +
                    $"COMMENT 新路徑：\n{newCommentPath}",
                    "完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "修改失敗：\n" + ex.Message,
                    "錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}