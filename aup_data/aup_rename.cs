using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Windows.Forms;
using Karoterra.AupDotNet;
using Karoterra.AupDotNet.ExEdit;
using Karoterra.AupDotNet.ExEdit.Effects;
using Karoterra.AupDotNet.Extensions;

namespace aup_data
{
    internal class aup_rename
    {
        public static void Rename(AviUtlProject _aup, ExEditProject _exedit, string save_fname, List<FileItem> fileItems, List<string> newNames)
        {
            for (int i = 0; i < fileItems.Count; i++)
            {
                var effect = _exedit.Objects[fileItems[i].ObjectIndex].Effects[fileItems[i].EffectIndex];
                try
                {
                    switch (effect)
                    {
                        case VideoFileEffect video:
                            video.Filename = newNames[i];
                            break;
                        case ImageFileEffect image:
                            image.Filename = newNames[i];
                            break;
                        case AudioFileEffect audio:
                            audio.Filename = newNames[i];
                            break;
                        case WaveformEffect waveform:
                            waveform.Filename = newNames[i];
                            break;
                        case ShadowEffect shadow:
                            shadow.Filename = newNames[i];
                            break;
                        case BorderEffect border:
                            border.Filename = newNames[i];
                            break;
                        case VideoCompositionEffect video:
                            video.Filename = newNames[i];
                            break;
                        case ImageCompositionEffect image:
                            image.Filename = newNames[i];
                            break;
                        case FigureEffect figure:
                            figure.Name = newNames[i];
                            break;
                        case MaskEffect mask:
                            mask.Name = newNames[i];
                            break;
                        case DisplacementEffect d:
                            d.Name = newNames[i];
                            break;
                        case PartialFilterEffect pf:
                            pf.Name = newNames[i];
                            break;
                        case ScriptFileEffect script:
                            script.Params["file"] = '"' + newNames[i].Replace(@"\", @"\\") + '"';
                            if (script.BuildParams().GetSjisByteCount() >= ScriptFileEffect.MaxParamsLength)
                            {
                                throw new MaxByteCountOfStringException();
                            }
                            break;
                    }
                }
                catch (MaxByteCountOfStringException)
                {
                    MessageBox.Show(".aupファイル内に記載されているファイル名が長すぎるため処理に失敗しました。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


            }

            try
            {
                using BinaryWriter writer = new(File.Create(save_fname));
                _aup.Write(writer);
            }
            catch (IOException)
            {
                MessageBox.Show(".aupファイルのパス書き込みに失敗しました。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(".aupファイルのパス書き込みに失敗しました。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
