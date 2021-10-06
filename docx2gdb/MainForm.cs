using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ganjoor;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;


namespace docx2gdb
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            cmbFormat.SelectedIndex = 2;//ریاض العارفین
        }

        enum WaitingFor
        {
            Right,
            Left,
            Middle
        }

        private void ConvertFormatVahdat()
        {
            if (!File.Exists(txtInput.Text))
            {
                MessageBox.Show("فایل ورودی وجود ندارد.");
                return;
            }
            if (File.Exists(txtOutput.Text))
            {
                if (
                    MessageBox.Show("فایل خروجی وجود دارد. پاک شود؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign)
                    == DialogResult.No
                    )
                    return;
                File.Delete(txtOutput.Text);
            }
            DbBrowser newGdbFile = DbBrowser.CreateNewPoemDatabase(
                /*string fileName = */ txtOutput.Text,
                /*bool faileIfExists =*/ true
                );
            if (newGdbFile != null)
            {
                int newPoetID = newGdbFile.NewPoet(txtPoetName.Text);
                GanjoorCat newCategory = newGdbFile.CreateNewCategory(txtCategoryName.Text, /*int ParentCategoryID = */ newGdbFile.GetPoet(newPoetID)._CatID, newPoetID);

                bool emptyMet = false;
                WaitingFor w = WaitingFor.Right;
                int poetNum = 0;
                List<string> RightVerses = new List<string>();
                List<string> LeftVerses = new List<string>();
                using (WordprocessingDocument doc = WordprocessingDocument.Open(txtInput.Text/*path*/, false/*I do not want to edit*/))
                {
                    Body body = doc.MainDocumentPart.Document.Body;
                    foreach (Table table in body.Descendants().OfType<Table>())
                    {
                        foreach (TableRow row in table.Descendants().OfType<TableRow>())
                        {
                            foreach (TableCell cell in row.Descendants().OfType<TableCell>())
                            {
                                List<string> verses = new List<string>();
                                string paragraphText = "";
                                foreach (Paragraph paragraph in cell.Descendants().OfType<Paragraph>())
                                {
                                    foreach (Run run in paragraph.Descendants().OfType<Run>())
                                    {
                                        if (run.ChildElements.OfType<Break>().Any())
                                        {
                                            if (!string.IsNullOrEmpty(paragraphText))
                                                verses.Add(paragraphText);
                                            paragraphText = "";
                                        }
                                        paragraphText += run.InnerText
                                            //optimize for search
                                            .Trim().Replace((char)0x200F, (char)0x200C).Replace("ي", "ی").Replace((char)0xE81D, (char)0x200C);
                                    }
                                }
                                if (!string.IsNullOrEmpty(paragraphText))
                                    verses.Add(paragraphText);

                                if (verses.Count == 0)
                                {
                                    emptyMet = true;
                                }
                                else
                                {
                                    if (!emptyMet && verses.Count != 2)//if after LEFT VERSES there are no BLANK CELLS, do wait for another cell with RIGHT VERSES instead of MIDDLE
                                    {
                                        if (w == WaitingFor.Middle)
                                            w = WaitingFor.Right;
                                    }
                                    emptyMet = false;
                                    switch (w)
                                    {
                                        case WaitingFor.Right:
                                            RightVerses.AddRange(verses);
                                            w = WaitingFor.Left;
                                            break;
                                        case WaitingFor.Left:
                                            LeftVerses.AddRange(verses);
                                            w = WaitingFor.Middle;
                                            break;
                                        case WaitingFor.Middle:
                                            if (RightVerses.Count != LeftVerses.Count)
                                            {
                                                using (CorrectVerses dlg = new CorrectVerses())
                                                {
                                                    dlg.RightVerses = RightVerses.ToArray();
                                                    dlg.LeftVerses = LeftVerses.ToArray();
                                                    DialogResult dlgResult = dlg.ShowDialog(this);
                                                    if (dlgResult == System.Windows.Forms.DialogResult.Yes)
                                                    {
                                                        RightVerses = new List<string>(dlg.RightVerses);
                                                        LeftVerses = new List<string>(dlg.LeftVerses);
                                                    }
                                                    else
                                                        if (dlgResult == System.Windows.Forms.DialogResult.Abort)
                                                    {
                                                        newGdbFile.CloseDb();
                                                        MessageBox.Show("نیمه‌کاره انجام شد.");
                                                        return;
                                                    }
                                                }
                                            }
                                            if (RightVerses.Count == LeftVerses.Count)
                                            {
                                                List<string> RightAndLeftVerses = new List<string>();
                                                for (int i = 0; i < RightVerses.Count/*or LeftVerses.Count*/; i++)
                                                {
                                                    RightAndLeftVerses.Add(RightVerses[i]);
                                                    RightAndLeftVerses.Add(LeftVerses[i]);
                                                }
                                                RightAndLeftVerses.AddRange(verses);
                                                poetNum++;
                                                GanjoorPoem newPoem = newGdbFile.CreateNewPoem(txtPoemTitlePrefix.Text + " " + poetNum.ToString(), newCategory._ID);
                                                newGdbFile.BeginBatchOperation();//speed up batch INSERT sql statements begins
                                                int beforeVerse = 0;
                                                foreach (string VerseText in RightAndLeftVerses)
                                                {
                                                    GanjoorVerse newVerse = newGdbFile.CreateNewVerse(newPoem._ID, beforeVerse, beforeVerse % 2 == 0 ? VersePosition.Right : VersePosition.Left);
                                                    newGdbFile.SetVerseText(newPoem._ID, newVerse._Order, VerseText);
                                                    beforeVerse++;
                                                }

                                                newGdbFile.CommitBatchOperation();//speed up batch INSERT sql statements ends
                                            }
                                            RightVerses.Clear(); LeftVerses.Clear();
                                            w = WaitingFor.Right;
                                            break;
                                    }
                                    verses.Clear();
                                }


                            }
                        }
                    }
                }

                newGdbFile.CloseDb();
                MessageBox.Show("انجام شد.");

            }
        }

        private void ConvertFormatValad()
        {
            if (!File.Exists(txtInput.Text))
            {
                MessageBox.Show("فایل ورودی وجود ندارد.");
                return;
            }
            if (File.Exists(txtOutput.Text))
            {
                if (
                    MessageBox.Show("فایل خروجی وجود دارد. پاک شود؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign)
                    == DialogResult.No
                    )
                    return;
                File.Delete(txtOutput.Text);
            }
            DbBrowser newGdbFile = DbBrowser.CreateNewPoemDatabase(
                /*string fileName = */ txtOutput.Text,
                /*bool faileIfExists =*/ true
                );
            if (newGdbFile != null)
            {
                int newPoetID = newGdbFile.NewPoet(txtPoetName.Text);
                GanjoorCat newCategory = newGdbFile.CreateNewCategory(txtCategoryName.Text, /*int ParentCategoryID = */ newGdbFile.GetPoet(newPoetID)._CatID, newPoetID);

                bool emptyMet = false;
                WaitingFor w = WaitingFor.Right;
                int poemNum = 0;
                int poemId = 0;
                List<string> RightVerses = new List<string>();
                List<string> LeftVerses = new List<string>();
                newGdbFile.BeginBatchOperation();//speed up batch INSERT sql statements begins
                using (WordprocessingDocument doc = WordprocessingDocument.Open(txtInput.Text/*path*/, false/*I do not want to edit*/))
                {
                    Body body = doc.MainDocumentPart.Document.Body;
                    foreach (Table table in body.Descendants().OfType<Table>())
                    {
                        if (table.PreviousSibling<Paragraph>() != null)
                        {
                            poemNum++;
                            GanjoorPoem newPoem = newGdbFile.CreateNewPoem(table.PreviousSibling<Paragraph>().InnerText, newCategory._ID);
                            poemId = newPoem._ID;
                        }
                        foreach (TableRow row in table.Descendants().OfType<TableRow>())
                        {
                            foreach (TableCell cell in row.Descendants().OfType<TableCell>())
                            {
                                List<string> verses = new List<string>();
                                string paragraphText = "";
                                foreach (Paragraph paragraph in cell.Descendants().OfType<Paragraph>())
                                {
                                    foreach (Run run in paragraph.Descendants().OfType<Run>())
                                    {
                                        if (run.ChildElements.OfType<Break>().Any())
                                        {
                                            if (!string.IsNullOrEmpty(paragraphText))
                                                verses.Add(paragraphText);
                                            paragraphText = "";
                                        }
                                        paragraphText += " " + run.InnerText
                                            //optimize for search
                                            .Trim().Replace((char)0x200F, (char)0x200C).Replace("ي", "ی").Replace((char)0xE81D, (char)0x200C);
                                    }
                                }
                                if (!string.IsNullOrEmpty(paragraphText))
                                    verses.Add(paragraphText.Replace("  ", " ").Replace("  ", " ").Trim());

                                if (verses.Count == 0)
                                {
                                    emptyMet = true;
                                }
                                else
                                {
                                    if (!emptyMet && verses.Count != 2)//if after LEFT VERSES there are no BLANK CELLS, do wait for another cell with RIGHT VERSES instead of MIDDLE
                                    {
                                        if (w == WaitingFor.Middle)
                                            w = WaitingFor.Right;
                                    }
                                    emptyMet = false;
                                    switch (w)
                                    {
                                        case WaitingFor.Right:
                                            RightVerses.AddRange(verses);
                                            w = WaitingFor.Left;

                                            break;
                                        case WaitingFor.Left:
                                            LeftVerses.AddRange(verses);
                                            w = WaitingFor.Middle;
                                            if (RightVerses.Count == LeftVerses.Count)
                                            {
                                                List<string> RightAndLeftVerses = new List<string>();
                                                for (int i = 0; i < RightVerses.Count/*or LeftVerses.Count*/; i++)
                                                {
                                                    RightAndLeftVerses.Add(RightVerses[i]);
                                                    RightAndLeftVerses.Add(LeftVerses[i]);
                                                }


                                                int beforeVerse = 0;
                                                foreach (string VerseText in RightAndLeftVerses)
                                                {
                                                    GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, beforeVerse, beforeVerse % 2 == 0 ? VersePosition.Right : VersePosition.Left);
                                                    newGdbFile.SetVerseText(poemId, newVerse._Order, VerseText);
                                                    beforeVerse++;
                                                }


                                                verses.Clear();
                                                RightVerses.Clear(); LeftVerses.Clear();
                                                w = WaitingFor.Right;
                                            }
                                            break;
                                        case WaitingFor.Middle:
                                            if (RightVerses.Count != LeftVerses.Count)
                                            {
                                                using (CorrectVerses dlg = new CorrectVerses())
                                                {
                                                    dlg.RightVerses = RightVerses.ToArray();
                                                    dlg.LeftVerses = LeftVerses.ToArray();
                                                    DialogResult dlgResult = dlg.ShowDialog(this);
                                                    if (dlgResult == System.Windows.Forms.DialogResult.Yes)
                                                    {
                                                        RightVerses = new List<string>(dlg.RightVerses);
                                                        LeftVerses = new List<string>(dlg.LeftVerses);
                                                    }
                                                    else
                                                        if (dlgResult == System.Windows.Forms.DialogResult.Abort)
                                                    {
                                                        newGdbFile.CloseDb();
                                                        MessageBox.Show("نیمه‌کاره انجام شد.");
                                                        return;
                                                    }
                                                }
                                            }


                                            break;
                                    }

                                }


                            }
                        }
                    }
                }
                newGdbFile.CommitBatchOperation();//speed up batch INSERT sql statements ends

                newGdbFile.CloseDb();
                MessageBox.Show("انجام شد.");

            }
        }

        private void ConvertFormatRiaz()
        {
            if (!File.Exists(txtInput.Text))
            {
                MessageBox.Show("فایل ورودی وجود ندارد.");
                return;
            }
            if (File.Exists(txtOutput.Text))
            {
                if (
                    MessageBox.Show("فایل خروجی وجود دارد. پاک شود؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign)
                    == DialogResult.No
                    )
                    return;
                File.Delete(txtOutput.Text);
            }
            DbBrowser newGdbFile = DbBrowser.CreateNewPoemDatabase(
                /*string fileName = */ txtOutput.Text,
                /*bool faileIfExists =*/ true
                );
            if (newGdbFile != null)
            {
                int newPoetID = newGdbFile.NewPoet(txtPoetName.Text);
                GanjoorCat newCategory = newGdbFile.CreateNewCategory(txtCategoryName.Text, /*int ParentCategoryID = */ newGdbFile.GetPoet(newPoetID)._CatID, newPoetID);


                int poemNum = 0;
                int poemId = 0;
                int verseOrder = 0;

                newGdbFile.BeginBatchOperation();//speed up batch INSERT sql statements begins
                using (WordprocessingDocument doc = WordprocessingDocument.Open(txtInput.Text/*path*/, false/*I do not want to edit*/))
                {
                    Body body = doc.MainDocumentPart.Document.Body;
                    foreach (var item in body.Descendants())
                    {
                        if (item is Paragraph)
                        {
                            Paragraph paragraph = item as Paragraph;
                            if (paragraph.Parent is TableCell)
                                continue;
                            if (paragraph.Descendants<ParagraphStyleId>().Any() && paragraph.Descendants().OfType<ParagraphStyleId>().First().Val == "Heading1")
                            {
                                poemNum++;
                                GanjoorPoem newPoem = newGdbFile.CreateNewPoem(paragraph.InnerText, newCategory._ID);
                                poemId = newPoem._ID;
                                verseOrder = 0;
                            }
                            else
                            {
                                List<string> verses = new List<string>();
                                string paragraphText = "";
                                foreach (Run run in paragraph.Descendants().OfType<Run>())
                                {
                                    if (run.ChildElements.OfType<Break>().Any())
                                    {
                                        if (!string.IsNullOrEmpty(paragraphText))
                                            verses.Add(paragraphText);
                                        paragraphText = "";
                                    }
                                    paragraphText += run.InnerText
                                        //optimize for search
                                        .Trim().Replace((char)0x200F, (char)0x200C).Replace("ي", "ی").Replace((char)0xE81D, (char)0x200C);
                                }
                                paragraphText = paragraphText.Replace("  ", " ").Replace("  ", " ").Trim();

                                if (!string.IsNullOrEmpty(paragraphText))
                                    if (paragraph.Descendants().OfType<Run>().Any())
                                    {
                                        if (paragraph.Descendants().OfType<Run>().First().Descendants().OfType<RunProperties>().Any())
                                            if (paragraph.Descendants().OfType<Run>().First().Descendants().OfType<RunProperties>()
                                                .First().Descendants().OfType<Bold>().Any())
                                            {

                                            paragraphText = paragraphText.Replace("  ", " ").Replace("  ", " ").Trim();
                                            if (!string.IsNullOrEmpty(paragraphText))
                                            {
                                                GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, verseOrder, VersePosition.Comment);
                                                newGdbFile.SetVerseText(poemId, newVerse._Order, paragraphText);
                                                verseOrder++;
                                                paragraphText = "";
                                            }
                                        }
                                    }
                                if (!string.IsNullOrEmpty(paragraphText))
                                    verses.Add(paragraphText);

                                foreach (var verse in verses)
                                {
                                    GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, verseOrder, VersePosition.Paragraph);
                                    newGdbFile.SetVerseText(poemId, newVerse._Order, verse);
                                    verseOrder++;
                                }
                            }
                        }
                        else
                        if (item is Table)
                        {
                            bool emptyMet = false;
                            WaitingFor w = WaitingFor.Right;
                            List<string> RightVerses = new List<string>();
                            List<string> LeftVerses = new List<string>();
                            Table table = item as Table;
                            foreach (TableRow row in table.Descendants().OfType<TableRow>())
                                foreach (TableCell cell in row.Descendants().OfType<TableCell>())
                                {
                                    List<string> verses = new List<string>();
                                    string paragraphText = "";
                                    foreach (Paragraph paragraph in cell.Descendants().OfType<Paragraph>())
                                    {

                                        foreach (Run run in paragraph.Descendants().OfType<Run>())
                                        {
                                            if (run.ChildElements.OfType<Break>().Any())
                                            {
                                                paragraphText = paragraphText.Replace("٭٭٭", "").Replace("  ", " ").Replace("  ", " ").Trim();
                                                if (!string.IsNullOrEmpty(paragraphText))
                                                    verses.Add(paragraphText);
                                                paragraphText = "";
                                            }
                                            paragraphText += run.InnerText
                                                //optimize for search
                                                .Replace((char)0x200F, (char)0x200C).Replace("ي", "ی").Replace((char)0xE81D, (char)0x200C).Replace("  ", " ").Replace("  ", " ").Trim();
                                        }

                                        if (!string.IsNullOrEmpty(paragraphText))
                                            if (paragraph.Descendants().OfType<ParagraphProperties>().Any())
                                            {
                                                if (paragraph.Descendants().OfType<ParagraphProperties>().First().Descendants().OfType<Justification>().Any())
                                                {

                                                    paragraphText = paragraphText.Replace("  ", " ").Replace("  ", " ").Trim();
                                                    if (!string.IsNullOrEmpty(paragraphText))
                                                    {
                                                        GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, verseOrder, VersePosition.Comment);
                                                        newGdbFile.SetVerseText(poemId, newVerse._Order, paragraphText);
                                                        verseOrder++;
                                                        paragraphText = "";
                                                    }
                                                }
                                            }
                                    }
                                    paragraphText = paragraphText.Replace("٭٭٭", "").Replace("  ", " ").Replace("  ", " ").Trim();
                                    if (!string.IsNullOrEmpty(paragraphText))
                                        verses.Add(paragraphText);



                                    if (verses.Count == 0)
                                    {
                                        emptyMet = true;
                                    }
                                    else
                                    {
                                        if (!emptyMet && verses.Count != 2)//if after LEFT VERSES there are no BLANK CELLS, do wait for another cell with RIGHT VERSES instead of MIDDLE
                                        {
                                            if (w == WaitingFor.Middle)
                                                w = WaitingFor.Right;
                                        }
                                        emptyMet = false;
                                        switch (w)
                                        {
                                            case WaitingFor.Right:
                                                RightVerses.AddRange(verses);
                                                w = WaitingFor.Left;

                                                break;
                                            case WaitingFor.Left:
                                                LeftVerses.AddRange(verses);
                                                w = WaitingFor.Middle;
                                                if (RightVerses.Count == LeftVerses.Count)
                                                {
                                                    List<string> RightAndLeftVerses = new List<string>();
                                                    for (int i = 0; i < RightVerses.Count/*or LeftVerses.Count*/; i++)
                                                    {
                                                        RightAndLeftVerses.Add(RightVerses[i]);
                                                        RightAndLeftVerses.Add(LeftVerses[i]);
                                                    }


                                                    int beforeVerse = 0;
                                                    foreach (string VerseText in RightAndLeftVerses)
                                                    {
                                                        GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, verseOrder, beforeVerse % 2 == 0 ? VersePosition.Right : VersePosition.Left);
                                                        newGdbFile.SetVerseText(poemId, newVerse._Order, VerseText);
                                                        beforeVerse++;
                                                        verseOrder++;
                                                    }


                                                    verses.Clear();
                                                    RightVerses.Clear(); LeftVerses.Clear();
                                                    w = WaitingFor.Right;
                                                }
                                                break;
                                            case WaitingFor.Middle:
                                                if (RightVerses.Count != LeftVerses.Count)
                                                {
                                                    using (CorrectVerses dlg = new CorrectVerses())
                                                    {
                                                        dlg.RightVerses = RightVerses.ToArray();
                                                        dlg.LeftVerses = LeftVerses.ToArray();
                                                        DialogResult dlgResult = dlg.ShowDialog(this);
                                                        if (dlgResult == System.Windows.Forms.DialogResult.Yes)
                                                        {
                                                            RightVerses = new List<string>(dlg.RightVerses);
                                                            LeftVerses = new List<string>(dlg.LeftVerses);
                                                        }
                                                        else
                                                            if (dlgResult == System.Windows.Forms.DialogResult.Abort)
                                                        {
                                                            newGdbFile.CommitBatchOperation();//speed up batch INSERT sql statements ends
                                                            newGdbFile.CloseDb();
                                                            MessageBox.Show("نیمه‌کاره انجام شد.");
                                                            return;
                                                        }
                                                    }

                                                }
                                                break;
                                        }

                                    }


                                }
                            if (RightVerses.Count != 0 && RightVerses.Count == LeftVerses.Count)
                            {
                                List<string> RightAndLeftVerses = new List<string>();
                                for (int i = 0; i < RightVerses.Count/*or LeftVerses.Count*/; i++)
                                {
                                    RightAndLeftVerses.Add(RightVerses[i]);
                                    RightAndLeftVerses.Add(LeftVerses[i]);
                                }


                                int beforeVerse = 0;
                                foreach (string VerseText in RightAndLeftVerses)
                                {
                                    GanjoorVerse newVerse = newGdbFile.CreateNewVerse(poemId, verseOrder, beforeVerse % 2 == 0 ? VersePosition.Right : VersePosition.Left);
                                    newGdbFile.SetVerseText(poemId, newVerse._Order, VerseText);
                                    beforeVerse++;
                                    verseOrder++;
                                }


                            }
                            else if (RightVerses.Count != 0)
                            {
                                using (CorrectVerses dlg = new CorrectVerses())
                                {
                                    dlg.RightVerses = RightVerses.ToArray();
                                    dlg.LeftVerses = LeftVerses.ToArray();
                                    DialogResult dlgResult = dlg.ShowDialog(this);
                                    if (dlgResult == System.Windows.Forms.DialogResult.Yes)
                                    {
                                        RightVerses = new List<string>(dlg.RightVerses);
                                        LeftVerses = new List<string>(dlg.LeftVerses);
                                    }
                                    else
                                        if (dlgResult == System.Windows.Forms.DialogResult.Abort)
                                    {
                                        newGdbFile.CommitBatchOperation();//speed up batch INSERT sql statements ends
                                        newGdbFile.CloseDb();
                                        MessageBox.Show("نیمه‌کاره انجام شد.");
                                        return;
                                    }
                                }
                            }
                        }
                    }

                }
                newGdbFile.CommitBatchOperation();//speed up batch INSERT sql statements ends

                newGdbFile.CloseDb();
                MessageBox.Show("انجام شد.");

            }
        }
        private void btnConvert_Click(object sender, EventArgs e)
        {
            switch (cmbFormat.SelectedIndex)
            {
                case 1:
                    ConvertFormatValad();
                    break;
                case 2:
                    ConvertFormatRiaz();
                    break;
                default:
                    ConvertFormatVahdat();
                    break;
            }
        }



        private void btnSelectInput_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "docx files (*.docx)|*.docx";
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    txtInput.Text = dlg.FileName;
            }
        }

        private void btnSelectOutput_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "gdb files (*.gdb)|*.gdb";
                if (dlg.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                    txtOutput.Text = dlg.FileName;
            }
        }


    }
}
