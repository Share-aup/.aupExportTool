using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Karoterra.AupDotNet;
using Karoterra.AupDotNet.ExEdit;
using Karoterra.AupDotNet.ExEdit.Effects;
using System.IO.Compression;


namespace aup_data
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private const string Pattern = @"<s\d*,([^,>]+)(,[BI]*)?>";
        private List<FileItem> file_item = new();

        //aup file
        String aup_file = "";

        String export_folder = "";


        private void metroSetButton3_Click(object sender, EventArgs e)
        {
            //ファイル選択
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "AviUtlプロジェクトファイル(*.aup)|*.aup";
            ofd.FilterIndex = 2;
            ofd.Title = ".aupファイルの選択";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;
            ofd.CheckPathExists = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                metroSetTextBox1.Text = ofd.FileName;
                aup_file = ofd.FileName;
            }
        }

        private void metroSetButton2_Click(object sender, EventArgs e)
        {
            //項目が全て埋まってるかとか
            if (metroSetTextBox1.Text == "" && metroSetTextBox2.Text == "")
            {
                MessageBox.Show("設定されてない値が存在します。", "Error");
                return;
            }

            //情報取得
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


            var aup = new AviUtlProject();
            using (var reader = new BinaryReader(File.OpenRead(aup_file)))
            {
                aup.Read(reader);
            }
            ExEditProject exedit = null;
            for (int i = 0; i < aup.FilterProjects.Count; i++)
            {
                if (aup.FilterProjects[i].Name == "拡張編集")
                {
                    exedit = new ExEditProject(aup.FilterProjects[i] as RawFilterProject);
                    aup.FilterProjects[i] = exedit;
                    break;
                }
            }

            if (exedit == null)
            {
                MessageBox.Show("拡張編集のデータが確認できません");
            }

            /*================
             フォント一覧取得
            ================*/
            ICollection<string> fonts;
            fonts = new SortedSet<string>();

            var reg = new Regex(Pattern);
            foreach (var obj in exedit.Objects)
            {
                if (obj.Effects[0] is TextEffect te)
                {
                    fonts.Add(te.Font);
                    var matches = reg.Matches(te.Text);
                    foreach (Match m in matches)
                    {
                        fonts.Add(m.Groups[1].Value);
                    }
                }
            }

            String font_list = "";
            foreach (var font in fonts)
            {
                font_list = font_list + font + "\n";
            }
            //MessageBox.Show($"拡張編集{exedit.Version}\n\n{font_list}", "Info");

            /*================
             スクリプト一覧取得
            ================*/
            string script_list = "";
            foreach (var obj in exedit.Objects)
            {
                foreach (var effect in obj.Effects)
                {
                    string scriptName = "";
                    switch (effect)
                    {
                        case AnimationEffect anm:
                            scriptName = anm.Name;
                            if (scriptName != "")
                            {
                                script_list = script_list + scriptName + "\n";
                            }
                            break;
                        case CustomObjectEffect coe:
                            scriptName = coe.Name;
                            if (scriptName != "")
                            {
                                script_list = script_list + scriptName + "\n";
                            }
                            break;
                        case CameraEffect cam:
                            scriptName = cam.Name;
                            if (scriptName != "")
                            {
                                script_list = script_list + scriptName + "\n";
                            }
                            break;
                        case SceneChangeEffect scn:
                            scriptName = scn.Name;
                            if (scriptName != "")
                            {
                                script_list = script_list + scriptName + "\n";
                            }
                            break;
                    }

                }
            }
            //MessageBox.Show(script_list, "Info");

            /*================
             ファイル一覧取得
            ================*/
            file_item.Clear();
            for (int objIdx = 0; objIdx < exedit.Objects.Count; objIdx++)
            {
                var obj = exedit.Objects[objIdx];
                for (int effectIdx = 0; effectIdx < obj.Effects.Count; effectIdx++)
                {
                    try
                    {
                        var effect = obj.Effects[effectIdx];
                        if (effect is VideoFileEffect video && video.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = video.Filename });
                        }
                        else if (effect is ImageFileEffect image && image.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = image.Filename });
                        }
                        else if (effect is AudioFileEffect audio && audio.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = audio.Filename });
                        }
                        else if (effect is WaveformEffect waveform && waveform.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = waveform.Filename });
                        }
                        else if (effect is ShadowEffect shadow && shadow.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = shadow.Filename });
                        }
                        else if (effect is BorderEffect border && border.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = border.Filename });
                        }
                        else if (effect is VideoCompositionEffect videoComp && videoComp.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = videoComp.Filename });
                        }
                        else if (effect is ImageCompositionEffect imageComp && imageComp.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = imageComp.Filename });
                        }
                        else if (effect is FigureEffect figure && figure.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = figure.Filename });
                        }
                        else if (effect is MaskEffect mask && mask.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = mask.Filename });
                        }
                        else if (effect is DisplacementEffect displacement && displacement.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = displacement.Filename });
                        }
                        else if (effect is PartialFilterEffect pf && pf.Filename != "")
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = pf.Filename });
                        }
                        else if (effect is ScriptFileEffect script && script.Params.ContainsKey("file") && script.Params["file"] != "")
                        {
                            var filename = script.Params["file"];
                            filename = filename[1..^1].Replace(@"\\", @"\");
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = filename });
                        }

                    }
                    catch (Exception s)
                    {
                        MessageBox.Show(s.ToString() + "\nOKを押して続行", "Error");
                    }
                }
            }

            if (file_item.Count == 0)
            {
                //依存関係をテキストファイルに出力
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                StreamWriter writer =
                  new StreamWriter($"{export_folder}\\依存関係.txt", true, sjisEnc);
                writer.WriteLine($"【バージョン】\n拡張編集:{exedit.Version}\n\n【スクリプト一覧】\n{script_list}\n【フォント一覧】\n{font_list}");
                writer.Close();
                MessageBox.Show("選択されたプロジェクトファイル内に有効なファイルが確認されなかったため、圧縮アーカイブは作成されず依存関係のみ出力されました。", "Status");
                return;
            }

            String file_list = "";
            foreach (var file in file_item)
            {
                file_list = file_list + file.Filename + "\n";
            }

            //MessageBox.Show(file_list, "Info");

            /*================
             zipファイルに出力
            ================*/
            try
            {
                //プログレスバー
                metroSetProgressBar1.Minimum = 0;
                metroSetProgressBar1.Maximum = file_item.Count;
                metroSetProgressBar1.Value = 0;

                //tmpディレクトリ作成
                Directory.CreateDirectory($"{export_folder}\\tmp");

                //aupファイルのコピー
                File.Copy(aup_file, $"{export_folder}\\tmp\\{Path.GetFileName(aup_file)}", true);
                int count_copy = 0;
                //素材のコピー
                foreach (var file in file_item)
                {
                    if (System.IO.File.Exists(file.Filename))
                    {
                        File.Copy(file.Filename, $"{export_folder}\\tmp\\{Path.GetFileName(file.Filename)}", true);
                    }
                    count_copy++;
                    metroSetProgressBar1.Value = count_copy;
                }

                //zipに圧縮
                ZipFile.CreateFromDirectory($"{export_folder}\\tmp", $"{export_folder}\\{Path.GetFileNameWithoutExtension(aup_file)}.zip");

                //依存関係をテキストファイルに出力
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                StreamWriter writer =
                  new StreamWriter($"{export_folder}\\依存関係.txt", true, sjisEnc);
                writer.WriteLine($"【バージョン】\n拡張編集:{exedit.Version}\n\n【スクリプト一覧】\n{script_list}\n【フォント一覧】\n{font_list}");
                writer.Close();

                //ファイル消去オプションがTrueならtmpファイル消し去る
                if (metroSetSwitch1.Switched == true)
                {
                    DirectoryInfo di = new DirectoryInfo($"{export_folder}\\tmp");
                    di.Delete(true);
                    MessageBox.Show($".aupファイルと依存関係にある素材を\n「{export_folder}\\{Path.GetFileNameWithoutExtension(aup_file)}.zip」\nへ出力しました。\n依存関係にあるフォント及びスクリプト一覧は\n「{export_folder}\\依存関係.txt」\nに保存されています。\n※一時ファイルの自動消去オプションが有効化されていた為、tmpファイルは消去されています。", "Info");
                }
                else
                {
                    MessageBox.Show($".aupファイルと依存関係にある素材を\n「{export_folder}\\{Path.GetFileNameWithoutExtension(aup_file)}.zip」\nへ出力しました。\n依存関係にあるフォント及びスクリプト一覧は\n「{export_folder}\\依存関係.txt」\nに保存されています。", "Info");
                }

            }
            catch (Exception s)
            {
                MessageBox.Show(s.ToString() + "\nOKを押して続行", "Error");
            }
        }



        private void metroSetButton1_Click(object sender, EventArgs e)
        {
            //reset
            metroSetTextBox1.Text = null;
            metroSetTextBox2.Text = null;
        }

        private void metroSetButton4_Click(object sender, EventArgs e)
        {
            //出力フォルダ取得
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "出力フォルダを指定してください。";
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            fbd.ShowNewFolderButton = true;

            //ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                metroSetTextBox2.Text = fbd.SelectedPath;
                export_folder = fbd.SelectedPath;
            }

        }
    }
}
class FileItem
{
    public int ObjectIndex;
    public int EffectIndex;
    public string Filename;
}
