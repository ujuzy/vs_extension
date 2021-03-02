using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using System.Text.RegularExpressions;


namespace VisualStudioExtension
{
    public partial class StatisticsWindowControl : UserControl
    {
        private List<Statistics> mStatisticsList;

        private const string MKeywordsPattern =
            @"\b(alignas|alignof|and|and_eq|asm|auto|bitand|bitor|bool|break|case|catch|char|char16_t|char32_t|class|compl|const|constexpr|const_cast|continue|decltype|default|delete|do|double|dynamic_cast|else|enum|explicit|export|extern|false|float|for|friend|goto|if|inline|int|long|mutable|namespace|new|noexcept|not|not_eq|nullptr|operator|or|or_eq|private|protected|public|register|reinterpret_cast|return|short|signed|sizeof|static|static_assert|static_cast|struct|switch|template|this|thread_local|throw|true|try|typedef|typeid|typename|union|unsigned|using|virtual|void|volatile|wchar_t|while|xor|xor_eq)\b";

        private const string MAsteriskCommentPattern = @"/\*((.*?)|(((.*?)\n)+(.*?)))\*/";
        private const string MSlashCommentPattern = @"//((\r\n)|((.*?)[^\\]\r\n))";
        private const string MQuotationQuotePattern = @"""(("")|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]""))))";
        private const string MApostropheQuotePattern = @"'((')|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]'))))";
        private const string MEmptyLinePattern = @"[\n\r]\s*[\r\n]";

        public StatisticsWindowControl()
        {
            this.InitializeComponent();

            mStatisticsList = new List<Statistics>();
            StatisticsListView.ItemsSource = mStatisticsList;
        }

        private string RemoveComments(string sourceCode) => 
            Regex.Replace(sourceCode, 
                MAsteriskCommentPattern + "|" + MSlashCommentPattern + "|" + MQuotationQuotePattern + "|" + MApostropheQuotePattern,
                match =>
                {
                    if (match.Value.StartsWith("//") || match.Value.StartsWith("/*"))
                    {
                        return Environment.NewLine;
                    }

                    return match.Value;
                }, RegexOptions.Singleline);

        private string RemoveQuotes(string sourceCode) =>
            Regex.Replace(sourceCode,
                MAsteriskCommentPattern + "|" + MSlashCommentPattern + "|" + MQuotationQuotePattern + "|" + MApostropheQuotePattern,
                match =>
                {
                    if (match.Value.StartsWith("//") || match.Value.StartsWith("/*") ||
                        match.Value.StartsWith("\"") || match.Value.StartsWith("\'"))
                    {
                        return Environment.NewLine;
                    }

                    return match.Value;
                }, RegexOptions.Singleline);

        private void FunctionInfo(CodeElement codeElement)
        {
            var funcElement = codeElement as CodeFunction;
            var start = funcElement.GetStartPoint(vsCMPart.vsCMPartHeader);
            var end = funcElement.GetEndPoint(vsCMPart.vsCMPartBodyWithDelimiter);
            var sourceCode = start.CreateEditPoint().GetText(end);
            var openCurlyBracePos = sourceCode.IndexOf('{');

            if (openCurlyBracePos > -1)
            {
                var name = codeElement.FullName + "()";

                var withoutComments = (new Regex(MEmptyLinePattern)).Replace(RemoveComments(sourceCode), "\n");
                var withoutCommentsAndQuotes = (new Regex(MEmptyLinePattern)).Replace(RemoveQuotes(sourceCode), "\n");

                var linesCount = Regex.Matches(sourceCode, @"[\n]").Count + 1;
                var linesWithoutCommentsCount = Regex.Matches(withoutComments, @"[\n]").Count + 1;
                var keywordsCount = Regex.Matches(withoutCommentsAndQuotes, MKeywordsPattern).Count;

                mStatisticsList.Add(new Statistics()
                {
                    Name = name,
                    LinesCount = linesCount,
                    LinesWithoutCommentsCount = linesWithoutCommentsCount,
                    KeywordsCount = keywordsCount
                });
            }
        }

        private void ParseFile(FileCodeModel2 file)
        {
            Dispatcher.VerifyAccess();

            foreach (CodeElement codeElement in file.CodeElements)
            {
                if (codeElement.Kind == vsCMElement.vsCMElementFunction)
                {
                    FunctionInfo(codeElement);
                }
            }
        }

        private void MenuItemCallback()
        {
            Dispatcher.VerifyAccess();
            mStatisticsList.Clear();

            try
            {
                var dte = (DTE2)ServiceProvider.GlobalProvider.GetService(typeof(DTE));
                var project = dte.ActiveDocument.ProjectItem;
                var file = (FileCodeModel2)project.FileCodeModel;

                ParseFile(file);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
        }

        private void RefreshButton_OnClick(object sender, RoutedEventArgs e)
        {
            MenuItemCallback();
            StatisticsListView.Items.Refresh();
        }

        private void StatisticsListView_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateColumnsWidth(sender as ListView);
        }

        private void StatisticsListView_Loaded(object sender, RoutedEventArgs e)
        {
            UpdateColumnsWidth(sender as ListView);
        }

        private void UpdateColumnsWidth(ListView listView)
        {
            var columnWidth = listView.ActualWidth / (listView.View as GridView).Columns.Count;

            foreach (var column in (listView.View as GridView).Columns)
            {
                column.Width = columnWidth;
            }
        }
    }
}