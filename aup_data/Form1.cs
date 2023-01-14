using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            // リセットボタンを初期値無効
            reset1.Enabled = false;
            check_path_status.Start();
        }

        private const string Pattern = @"<s\d*,([^,>]+)(,[BI]*)?>";
        private List<FileItem> file_item = new();

        // aup file
        string aup_file = "";

        string export_folder = "";


        private void Select_aup_click(object sender, EventArgs e)
        {
            // ファイル選択
            OpenFileDialog ofd = new()
            {
                Filter = "AviUtlプロジェクトファイル(*.aup)|*.aup",
                FilterIndex = 2,
                Title = ".aupファイルの選択",
                RestoreDirectory = true,
                CheckFileExists = true,
                CheckPathExists = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                metroSetTextBox1.Text = ofd.FileName;
                aup_file = ofd.FileName;
            }
        }




        private void Select_export_click(object sender, EventArgs e)
        {
            // 出力フォルダ取得
            FolderBrowserDialog fbd = new()
            {
                Description = "出力先フォルダを指定してください。",
                RootFolder = Environment.SpecialFolder.Desktop,
                ShowNewFolderButton = true
            };

            // ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                metroSetTextBox2.Text = fbd.SelectedPath;
                export_folder = fbd.SelectedPath;
            }

        }

        private void run_click(object sender, EventArgs e)
        {
            // 項目が全て埋まってるかとか
            if (metroSetTextBox1.Text == "" || metroSetTextBox2.Text == "")
            {
                MessageBox.Show("未入力の値が存在します。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 情報取得
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
                MessageBox.Show(".aupファイル内に拡張編集のデータが確認できません", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            string font_list = "";
            foreach (var font in fonts)
            {
                font_list = font_list + font + "\n";
            }
            // フォント使用数が0の場合
            if (fonts.Count == 0)
            {
                font_list = "使用されているフォントはありません。";
            }
            // MessageBox.Show($"拡張編集{exedit.Version}\n\n{font_list}", "Info");

            /*================
             スクリプト一覧取得
            ================*/
            var scripts = new SortedSet<string>();
            foreach (var obj in exedit.Objects)
            {
                if (obj.Chain) continue;

                foreach (var effect in obj.Effects)
                {
                    string scriptName = "";
                    switch (effect)
                    {
                        case AnimationEffect anm:
                            scriptName = anm.Name;
                            break;
                        case CustomObjectEffect coe:
                            scriptName = coe.Name;
                            break;
                        case CameraEffect cam:
                            scriptName = cam.Name;
                            break;
                        case SceneChangeEffect scn when scn.Params != null && scn.ScriptId == 2:
                            scriptName = scn.Name;
                            break;
                    }
                    if (!string.IsNullOrEmpty(scriptName))
                    {
                        scripts.Add(scriptName);
                    }

                    foreach (var trackbar in effect.Trackbars)
                    {
                        if (trackbar.Type != TrackbarType.Script) continue;
                        if (trackbar.ScriptIndex < 0 || trackbar.ScriptIndex >= exedit.TrackbarScripts.Count) continue;

                        scriptName = exedit.TrackbarScripts[trackbar.ScriptIndex].Name;
                        if (TrackbarScript.Defaults.Any(x => x.Name == scriptName)) continue;
                        scripts.Add(scriptName);
                    }
                }
            }

            string script_list = "";
            // スクリプト使用数が0の場合
            foreach (var script in scripts)
            {
                script_list = script_list + script + "\n";
            }
            if (scripts.Count == 0)
            {
                script_list = "使用されているスクリプトはありません。";
            }
            // MessageBox.Show(script_list, "Info");



            /*================
             ファイル一覧取得
            ================*/
            file_item.Clear();
            for (int objIdx = 0; objIdx < exedit.Objects.Count; objIdx++)
            {
                var obj = exedit.Objects[objIdx];
                if (obj.Chain) continue;

                for (int effectIdx = 0; effectIdx < obj.Effects.Count; effectIdx++)
                {
                    try
                    {
                        var effect = obj.Effects[effectIdx];
                        if (effect is VideoFileEffect video && !string.IsNullOrEmpty(video.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = video.Filename });
                        }
                        else if (effect is ImageFileEffect image && !string.IsNullOrEmpty(image.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = image.Filename });
                        }
                        else if (effect is AudioFileEffect audio && !string.IsNullOrEmpty(audio.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = audio.Filename });
                        }
                        else if (effect is WaveformEffect waveform && !string.IsNullOrEmpty(waveform.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = waveform.Filename });
                        }
                        else if (effect is ShadowEffect shadow && !string.IsNullOrEmpty(shadow.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = shadow.Filename });
                        }
                        else if (effect is BorderEffect border && !string.IsNullOrEmpty(border.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = border.Filename });
                        }
                        else if (effect is VideoCompositionEffect videoComp && !string.IsNullOrEmpty(videoComp.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = videoComp.Filename });
                        }
                        else if (effect is ImageCompositionEffect imageComp && !string.IsNullOrEmpty(imageComp.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = imageComp.Filename });
                        }
                        else if (effect is FigureEffect figure && !string.IsNullOrEmpty(figure.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = figure.Filename });
                        }
                        else if (effect is MaskEffect mask && !string.IsNullOrEmpty(mask.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = mask.Filename });
                        }
                        else if (effect is DisplacementEffect displacement && !string.IsNullOrEmpty(displacement.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = displacement.Filename });
                        }
                        else if (effect is PartialFilterEffect pf && !string.IsNullOrEmpty(pf.Filename))
                        {
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = pf.Filename });
                        }
                        else if (effect is ScriptFileEffect script && !string.IsNullOrEmpty(script.Params?.GetValueOrDefault("file")))
                        {
                            var filename = script.Params["file"];
                            filename = filename[1..^1].Replace(@"\\", @"\");
                            file_item.Add(new FileItem() { ObjectIndex = objIdx, EffectIndex = effectIdx, Filename = filename });
                        }

                    }
                    catch (Exception s)
                    {
                        MessageBox.Show(s.ToString() + "\nOKを押して続行", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            if (file_item.Count == 0)
            {
                // 依存関係をテキストファイルに出力
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                StreamWriter writer =
                  new StreamWriter($"{export_folder}\\依存関係.txt", true, sjisEnc);
                writer.WriteLine($"拡張編集:{exedit.Version}\n\n【スクリプト】\n{script_list}\n\n【フォント】\n{font_list}");
                writer.Close();
                MessageBox.Show("選択されたプロジェクトファイル内に有効なファイルが確認されなかったため、圧縮アーカイブは作成されず依存関係のみ出力されました。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 競合管理用
            List<string> filename_list = new();
            // コピー先名称保存用
            List<List<string>> copy_list = new();
            // pf編集用
            List<string> pfe_list = new();
            
            string file_list = "";
            foreach (var file in file_item)
            {
                // 競合管理用配列にすでに存在するか確認
                if (filename_list.Contains(Path.GetFileName(file.Filename)))
                {
                    // 存在する

                    // 同名の要素が何個存在するか
                    int count = 0;
                    foreach (var filename in filename_list)
                    {
                        if (filename == file.Filename)
                        {
                            count++;
                        }
                    }
                    // n_ファイル名.ext
                    file_list = file_list + count + "_" + file.Filename + "\n";

                    // ディレクトリの取得
                    var dir = Path.GetDirectoryName(file.Filename) + "\\";

                    // 現在のファイル名を取得
                    string fname = Path.GetFileName(file.Filename);

                    // 元ファイルと移動先ファイル名を対応付ける
                    copy_list.Add(new List<string>() { file.Filename, count + "_" + fname });
                    pfe_list.Add(dir + count + "_" + fname);
                }
                else
                {
                    // 存在しない

                    // 現在のファイル名を取得
                    string fname = Path.GetFileName(file.Filename);

                    // 元ファイルと移動先ファイル名を対応付ける
                    copy_list.Add(new List<string>() { file.Filename, fname });
                }

                // ファイル名を競合間利用配列に追加
                filename_list.Add(Path.GetFileName(file.Filename));
            }


            /*================
             zipファイルに出力
            ================*/
            try
            {
                // プログレスバー
                metroSetProgressBar1.Minimum = 0;
                metroSetProgressBar1.Maximum = file_item.Count;
                metroSetProgressBar1.Value = 0;

                // tmpディレクトリ作成
                Directory.CreateDirectory($"{export_folder}\\tmp");

                string copy_aup_path = $"{export_folder}\\tmp\\_{Path.GetFileName(aup_file)}";

                // aupファイルのコピー
                File.Copy(aup_file, copy_aup_path, true);

                /*=====================================================
                 コピーしたaupファイル内のパスを競合しているもののみ変更先パスに置き換え
                =====================================================*/
                // コピーしたaupファイルを読み込み
                var _aup = new AviUtlProject();
                using (var reader = new BinaryReader(File.OpenRead(copy_aup_path)))
                {
                    _aup.Read(reader);
                }
                ExEditProject _exedit = null;
                for (int i = 0; i < _aup.FilterProjects.Count; i++)
                {
                    if (_aup.FilterProjects[i].Name == "拡張編集")
                    {
                        _exedit = new ExEditProject(_aup.FilterProjects[i] as RawFilterProject);
                        _aup.FilterProjects[i] = _exedit;
                        break;
                    }
                }
                // aupファイル内のファイル名を更新
                aup_rename.Rename(_aup, _exedit, "", file_item, pfe_list);
                //aup_rename.Rename(_aup, _exedit, export_path, file_item, pfe_list);
                
                int count_copy = 0;
                
                // copy_listのファイルを順次コピー
                foreach (var file in copy_list)
                {
                    // 素材元の存在確認
                    if (System.IO.File.Exists(file[0]))
                    {
                        // file[0] -> コピー元フルパス
                        // file[1] -> コピー先ファイル名
                        // コピー
                        File.Copy(file[0], $"{export_folder}\\tmp\\{file[1]}", true);
                        
                    }
                    // プログレスバーの反映
                    count_copy++;
                    metroSetProgressBar1.Value = count_copy;
                }

                // zipに圧縮
                ZipFile.CreateFromDirectory($"{export_folder}\\tmp", $"{export_folder}\\{Path.GetFileNameWithoutExtension(aup_file)}.zip");

                // 依存関係をテキストファイルに出力
                Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
                StreamWriter writer =
                  new StreamWriter($"{export_folder}\\依存関係.txt", true, sjisEnc);
                writer.WriteLine($"【バージョン】\n拡張編集:{exedit.Version}\n\n【スクリプト一覧】\n{script_list}\n\n【フォント一覧】\n{font_list}");
                writer.Close();

                // tmpファイル消し去る
                DirectoryInfo di = new DirectoryInfo($"{export_folder}\\tmp");
                di.Delete(true);
                MessageBox.Show($".aupファイルと依存関係にある素材を\n「{export_folder}\\{Path.GetFileNameWithoutExtension(aup_file)}.zip」\nへ出力しました。\n依存関係にあるフォント及びスクリプト一覧は\n「{export_folder}\\依存関係.txt」\nに保存されています。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // プログレスバーを初期値に
                metroSetProgressBar1.Value = 0;
            }
            catch (Exception s)
            {
                MessageBox.Show(s.ToString() + "\nOKを押して続行", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Reset_click(object sender, EventArgs e)
        {
            // reset
            metroSetTextBox1.Text = "";
            metroSetTextBox2.Text = "";
        }

        private void Check_path_status_Tick(object sender, EventArgs e)
        {
            if(metroSetTextBox1.Text != "" || metroSetTextBox2.Text != "")
            {
                reset1.Enabled = true;
            }
            else
            {
                reset1.Enabled = false;
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