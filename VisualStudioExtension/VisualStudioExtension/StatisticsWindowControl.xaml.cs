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
            @"\b(alignas|alignof|and|and_eq|asm|auto|bitand|bitor|bool|break|case|catch|
                 char|char16_t|char32_t|class|compl|const|constexpr|const_cast|continue|
                 decltype|default|delete|do|double|dynamic_cast|else|enum|explicit|export|
                 extern|false|float|for|friend|goto|if|inline|int|long|mutable|namespace|
                 new|noexcept|not|not_eq|nullptr|operator|or|or_eq|private|protected|public|
                 register|reinterpret_cast|return|short|signed|sizeof|static|static_assert|
                 static_cast|struct|switch|template|this|thread_local|throw|true|try|typedef|
                 typeid|typename|union|unsigned|using|virtual|void|volatile|wchar_t|while|xor|
                 xor_eq)\b";

        public StatisticsWindowControl()
        {
            this.InitializeComponent();

            mStatisticsList = new List<Statistics>();
            Statistics.ItemsSource = mStatisticsList;
        }

        private string RemoveComments(string sourceCode)
        {
            var multiLine = @"/\*(.*?)\*/";
            var singleLine = @"//(.*?)[^\\]\r\n";
            var quotes = @"""(("")|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]""))))";
            var one_quotes = @"'((')|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]'))))";

            return Regex.Replace(sourceCode, multiLine + "|" + singleLine + "|" + quotes + "|" + one_quotes,
                me =>
                {
                    if (me.Value.StartsWith("//"))
                    {
                        return Environment.NewLine;
                    }

                    if (me.Value.StartsWith("/*"))
                    {
                        return "";
                    }

                    return me.Value;
                }, RegexOptions.Singleline);
        }

        private string RemoveMultilineComments(string sourceCode)
        {
            var multiLine = @"/\*((.*?)\n)+(.*?)\*/";
            var quotes = @"""(("")|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]""))))";
            var one_quotes = @"'((')|(\r\n)|((.*?)(([^\\]\r\n)|([^\\]'))))";

            return Regex.Replace(sourceCode, multiLine + "|" + quotes + "|" + one_quotes, me =>
            {
                return me.Value.StartsWith("/*") ? Environment.NewLine : me.Value;
            }, RegexOptions.Singleline);
        }

        private string RemoveQuotes(string sourceCode)
        {
            sourceCode = Regex.Replace(sourceCode, @"""(("")|(\n)|((.*?)(([^\\]\n)|([^\\]""))))", "", RegexOptions.Singleline);
            return Regex.Replace(sourceCode, @"'((')|(\n)|((.*?)(([^\\]\n)|([^\\]'))))", "", RegexOptions.Singleline);
        }

        private string RemoveEmptyLines(string sourceCode)
        {
            var sReg = @"[\n\r]\s*[\r\n]";
            var rgx = new Regex(sReg);
            var emptyLines = rgx.Replace(sourceCode, "\n");

            return emptyLines;
        }

        private void FunctionInfo(CodeElement codeElement)
        {
            var funcElement = codeElement as CodeFunction;
            var start = funcElement.GetStartPoint(vsCMPart.vsCMPartHeader);
            var end = funcElement.GetEndPoint();
            var sourceCode = start.CreateEditPoint().GetText(end);
            var openCurlyBracePos = sourceCode.IndexOf('{');

            if (openCurlyBracePos > -1)
            {
                var name = codeElement.FullName + "()";

                var linesCount = Regex.Matches(sourceCode, @"[\n]").Count + 1;

                sourceCode = RemoveEmptyLines(RemoveMultilineComments(RemoveComments(sourceCode)));

                var linesWithoutCommentsCount = Regex.Matches(sourceCode, @"[\n]").Count + 1;

                sourceCode = RemoveQuotes(sourceCode);

                var keywordsCount = Regex.Matches(sourceCode, MKeywordsPattern).Count;
                
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
            Statistics.Items.Refresh();
        }
    }
}