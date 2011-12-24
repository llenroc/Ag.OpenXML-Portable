using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using FiftyNine.Ag.OpenXML.Common.Packaging;
using FiftyNine.Ag.OpenXML.Common.Storage;
using FiftyNine.Ag.OpenXML.Word.Elements;
using FiftyNine.Ag.OpenXML.Word.Parts;
using FiftyNine.Ag.OpenXML.Word.Utilities;
using Style = FiftyNine.Ag.OpenXML.Word.Elements.Style;

namespace FiftyNine.Ag.OpenXML.Word.Sample
{
    public partial class MainPage : UserControl
    {
        Stream _image;

        public MainPage()
        {
            InitializeComponent();
        }

        private void SelectImage(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "Imge Files (*.jpg, *.png)|*.jpg;*.jpeg;*.png";
            if (dlg.ShowDialog() == true)
            {
                _image = dlg.File.OpenRead();
                var img = new BitmapImage();
                img.SetSource(dlg.File.OpenRead());
                MyImage.Source = img;
                ImageBorder.Visibility = Visibility.Visible;
            }
        }
        private void Save(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog();
            dlg.Filter = "Word Document (.docx)|*.docx|Zip Files (.zip)|*.zip";
            dlg.DefaultExt = ".docx";
            if (dlg.ShowDialog() == true)
            {
                var doc = new WordDocument();
                doc.ApplicationName = "SilverWord";
                doc.Creator = "Chris Klug";
                doc.Company = "59NORTH";

                FontReference comicSans;
                FontReference aharoni;
                AddFonts(doc, out comicSans, out aharoni);

                var style = AddStyle(doc, "TitleStyle", "Title Style", aharoni);

                var run = doc.Document.CreateElement<Run>();
                var t = doc.Document.CreateElement<Text>();
                t.Content = Title.Text;
                run.Content.Add(t);
                run.Properties.Style = style;
                doc.Document.Sections[0].Paragraphs[0].Runs.Add(run);

                Paragraph p;
                string[] paragraphs = Text.Text.Split(new string[] { "\r\r" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var paragraph in paragraphs)
                {
                    string[] lines = paragraph.Split(new string[] { "\r" }, StringSplitOptions.RemoveEmptyEntries);
                    p = doc.Document.CreateElement<Paragraph>();
                    run = doc.Document.CreateElement<Run>();
                    run.Properties.Font.HighAnsi = comicSans;
                    run.Properties.Font.ComplexScript = comicSans;
                    run.Properties.Font.ASCII = comicSans;
                    run.Properties.Font.EastAsia = comicSans;
                    for (int i = 0; i < lines.Length; i++)
                    {
                        t = doc.Document.CreateElement<Text>();
                        t.Content = lines[i];
                        run.Content.Add(t);
                        if (i < lines.Length - 1)
                        {
                            run.Content.Add(doc.Document.CreateElement<Break>());
                        }
                    }
                    p.Runs.Add(run);
                    doc.Document.Sections[0].Paragraphs.Add(p);

                }

                if (_image != null)
                {
                    var img = doc.CreatePart<ImagePart>(doc.GetMediaPartName("image1.jpg"));
                    img.Image = _image;

                    var rel = doc.Document.AddRelation(img);
                    p = doc.Document.CreateElement<Paragraph>();
                    run = doc.Document.CreateElement<Run>();
                    var pict = doc.Document.CreateElement<Picture>();
                    pict.Image = rel;
                    pict.Size = new Size(400, 240);
                    run.Content.Add(pict);
                    p.Runs.Add(run);
                    doc.Document.Sections[0].Paragraphs.Insert(0, p);
                }

                using (IStreamProvider storage = new ZipStreamProvider(dlg.OpenFile()))
                {
                    doc.Save(storage);
                }

                _image = null;
                ImageBorder.Visibility = Visibility.Collapsed;
                Title.Text = "";
                Text.Text = "";
            }
        }

        private void AddFonts(WordDocument doc, out FontReference comicSans, out FontReference aharoni)
        {
            var fontDefinition = doc.FontTable.CreateFontDefinition("Comic Sans MS");
            fontDefinition.Panose1 = "030F0702030302020204";
            fontDefinition.CharSet = "00";
            fontDefinition.Family = FontFamilyEnumeration.Script;
            fontDefinition.Pitch = FontPitchEnumeration.Variable;
            fontDefinition.Signature.UnicodeSignature0 = "00000287";
            fontDefinition.Signature.CodePageSignature0 = "0000009F";
            comicSans = doc.FontTable.AddFont(fontDefinition);

            fontDefinition = doc.FontTable.CreateFontDefinition("Aharoni");
            fontDefinition.Panose1 = "02010803020104030203";
            fontDefinition.CharSet = "B1";
            fontDefinition.Family = FontFamilyEnumeration.Auto;
            fontDefinition.Pitch = FontPitchEnumeration.Variable;
            fontDefinition.Signature.UnicodeSignature0 = "00000801";
            fontDefinition.Signature.CodePageSignature0 = "00000020";
            aharoni = doc.FontTable.AddFont(fontDefinition);
        }

        private Style AddStyle(WordDocument doc, string id, string name, FontReference font)
        {
            CharacterStyle style = doc.Styles.AddCharacterStyle(id, name);
            style.RunProperties.FontSize = 30;
            style.RunProperties.IsBold = true;
            style.RunProperties.Font.ComplexScript = font;
            style.RunProperties.Font.HighAnsi = font;
            style.RunProperties.Font.ASCII = font;
            style.RunProperties.Font.EastAsia = font;
            return style;
        }
    }
}
