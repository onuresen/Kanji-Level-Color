using Microsoft.Office.Interop.Word;

namespace Kanji_Level_Color
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application fileOpen = new Microsoft.Office.Interop.Word.Application();
            //Open a already existing word file into the new document created

            string filePath = "";

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "doc files (*.docx)|*.docx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;

            openFileDialog.ShowDialog();
            if (true == openFileDialog.CheckFileExists)
                filePath = openFileDialog.FileName;

            WordText(openFileDialog.FileName, @"D:\NewFile1");

            Microsoft.Office.Interop.Word._Application word;
            Microsoft.Office.Interop.Word._Document document;
            word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = true;
            document = word.Documents.Open(filePath);  //a .doc or .docx file to open
            document.Activate();

            //object findStr = "ƒXƒgƒŒƒX"; //sonething to find
            //while (word.Selection.Find.Execute(ref findStr))  //found...
            //{
            //    //change font and format of matched words
            //    //word.Selection.Font.Name = "Tahoma"; //change font to Tahoma
            //    word.Selection.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdRed;  //change color to red
            //}

            string fileKanji = @"D:\Kanji.txt";


            if (File.Exists(fileKanji))
            {
                string[] lines = File.ReadAllLines(fileKanji);

                for (int i = 1; i < word.ActiveDocument.Words.Count; i++)
                {
                    string text = word.ActiveDocument.Words[i].Text;

                    for (int j = 0; j < text.Length; j++)
                    {
                        string level = KanjiLevel(lines, text[j].ToString());
                        var color = KanjiColor(level);

                        //word.ActiveDocument.Words[i].Font.ColorIndex = color;
                        var count = word.ActiveDocument.Words[i].Characters.Count;
                        word.ActiveDocument.Words[i].Characters[j+1].Font.ColorIndex = color;
                    }
                }

            }

            //if (File.Exists(fileKanji))
            //{
            //    string[] lines = File.ReadAllLines(fileKanji);

            //    foreach (string line in lines)
            //    {
            //        string[] splits = line.Split(';');
            //        object findStr = splits[1];

            //        WdColorIndex colorIndex = WdColorIndex.wdBlack;

            //        if (splits[0].Contains("5"))
            //            colorIndex = WdColorIndex.wdBrightGreen;
            //        else if (splits[0].Contains("4"))
            //            colorIndex = WdColorIndex.wdGreen;
            //        else if (splits[0].Contains("3"))
            //            colorIndex = WdColorIndex.wdBlue;
            //        else if (splits[0].Contains("2"))
            //            colorIndex = WdColorIndex.wdDarkBlue;
            //        else if (splits[0].Contains("1"))
            //            colorIndex = WdColorIndex.wdDarkRed;

            //        while (word.Selection.Find.Execute(ref findStr))
            //        {
            //            word.Selection.Font.ColorIndex = colorIndex;
            //        }
            //    }
            //}

            Close();
        }

        private void WordText(string fileRead, string fileWrite)
        {
            //Create a new microsoft word file
            Microsoft.Office.Interop.Word.Application fileOpen = new Microsoft.Office.Interop.Word.Application();
            //Open a already existing word file into the new document created
            Microsoft.Office.Interop.Word.Document document = fileOpen.Documents.Open(fileRead, ReadOnly: false);
            //Make the file visible 
            fileOpen.Visible = true;
            document.Activate();
            //The FindAndReplace takes the text to find under any formatting and replaces it with the
            //new text with the same exact formmating (e.g red bold text will be replaced with red bold text)
            FindAndReplace(fileOpen, "useless", "very useful");
            //Save the editted file in a specified location
            //Can use SaveAs instead of SaveAs2 and just give it a name to have it saved by default
            //to the documents folder
            document.SaveAs2(fileWrite);
            //Close the file out
            fileOpen.Quit();

        }

        //Method to find and replace the text in the word document. Replaces all instances of it
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private void FontColor(Microsoft.Office.Interop.Word.Application fileOpen)
        {
        }

        private string KanjiLevel(string[] kanjis, string kanji)
        {
            string level = "0";

            foreach (string line in kanjis)
            {
                string[] splits = line.Split(';');

                if (splits[1] == kanji)
                    return splits[0];
            }

            return level;
        }

        private WdColorIndex KanjiColor(string level)
        {
            WdColorIndex colorIndex = WdColorIndex.wdBlack;

            if (level.Contains("5"))
                colorIndex = WdColorIndex.wdBrightGreen;
            else if (level.Contains("4"))
                colorIndex = WdColorIndex.wdGreen;
            else if (level.Contains("3"))
                colorIndex = WdColorIndex.wdBlue;
            else if (level.Contains("2"))
                colorIndex = WdColorIndex.wdDarkBlue;
            else if (level.Contains("1"))
                colorIndex = WdColorIndex.wdDarkRed;

            return colorIndex;
        }
    }
}