using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Sony.Vegas;

public class EntryPoint
{
    /// <summary>
    /// 指定されたフォルダから全ての音声ファイル（テキストファイル付き）を読み込み、
    /// そのテキストファイルと一致するタイトルを持つTrackEventを選択していた場合、そこに音声データを代入
    /// </summary>
    private const string VoiceFolderName = @"\Voice";
    private const string VoiceTextFileEx = ".txt";
    private const string VoiceFileEx = ".wav";
    private const string TitleNamePrefix = "Sony タイトルおよびテキスト ";
    private const string Voiceroid2TrackName = "Voiceroid2";

    //VegasだとvarとかC#3.0のLinqのメソッド形式の入力がエラー
    public void FromVegas(Vegas vegas)
    {
        // 音声ファイルのパス取得
        string soundFolderPath = Path.GetDirectoryName(vegas.Project.FilePath) + VoiceFolderName;

        // ファイル名（拡張子無し）を取得
        List<string> filePathsNoEx = new List<string>();
        foreach (string file in Directory.GetFiles(soundFolderPath, "*" + VoiceFileEx))
        {
            filePathsNoEx.Add(Path.GetDirectoryName(file) + @"\" + Path.GetFileNameWithoutExtension(file));
        }
        /*
         List<string> filePathsNoEx = Directory.GetFiles(soundFolderPath, "*" + VoiceFileEx)
            .Select(x => x.Replace(x, System.IO.Path.GetFileNameWithoutExtension(x))).ToList();
            */

        // テキストファイルとそのテキストファイルの名前をDictionary形式で宣言
        Dictionary<string, string> textDataDic = new Dictionary<string, string>();

        // 選択されたテキストイベントを取得(Take名が指定の名称プレフィックスと一致している場合のみ)
        List<TrackEvent> selectedTextEvents = new List<TrackEvent>();
        foreach (TrackEvent ev in FindSelectedEvents(vegas.Project))
        {
            if (ev.ActiveTake.Name.IndexOf(TitleNamePrefix, StringComparison.Ordinal) != -1)
                selectedTextEvents.Add(ev);
        }
        /*
        List<TrackEvent> selectedTextEvents =
            FindSelectedEvents(vegas.Project).Where(x => x.Name.IndexOf(TitleNamePrefix, StringComparison.Ordinal) != -1).ToList();
        */

        // 選択されたトラックを取得
        Track selectedTrack = FindSelectedTrack(vegas.Project);
        if (selectedTrack.MediaType != MediaType.Audio)
        {
            selectedTrack = FindVoiceroid2Track(vegas.Project);
        }
        

        if (!CheckConfiguration(selectedTextEvents, selectedTrack, filePathsNoEx))
            return;

        // テキストファイル読み込み
        foreach (string name in filePathsNoEx)
        {
            StreamReader sr = new StreamReader(
                name + VoiceTextFileEx, Encoding.GetEncoding("Shift_JIS"));
            string text = sr.ReadToEnd();
            sr.Close();

            // 改行コードはタイトルに含まれないため、削除
            textDataDic.Add(text.Replace("\n", ""), name + VoiceFileEx);
        }

        // テキストイベント群と音声ファイル群でテキストが一致しているものをサーチ
        List<TrackEvent> fileExistTextEvents = new List<TrackEvent>();
        foreach (TrackEvent ev in selectedTextEvents)
        {
            if (textDataDic.ContainsKey(ev.ActiveTake.Name.Replace(TitleNamePrefix, "")))
                fileExistTextEvents.Add(ev);
        }
        // List<TrackEvent> fileExistTextEvents = selectedTextEvents.Where(x => textDataDic.ContainsKey(x.Name.Replace(TitleNamePrefix, ""))).ToList();

        // ダイアログのメッセージ生成
        string message = MakeDialogMessage(fileExistTextEvents, selectedTrack, textDataDic);

        // 確認ダイアログ生成
        DialogResult result = MessageBox.Show(message,
            "確認",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button2);

        if (result != DialogResult.OK)
            return;

        // 処理実行
        foreach (TrackEvent textEvent in fileExistTextEvents)
        {
            InsertFileAt(vegas, textDataDic[textEvent.ActiveTake.Name.Replace(TitleNamePrefix, "")], selectedTrack.Index, textEvent.Start);
        }
    }

    /// <summary>
    /// スクリプトを実行する条件を満たしているかを確認する（音声ファイルとともにテキストファイルが存在するかは未確認）
    /// </summary>
    /// <param name="selectedTextEvents">選択されたテキストイベント</param>
    /// <param name="selectedTrack">選択されたトラック</param>
    /// <param name="filePathsNoEx">音声ファイルのファイルパス達</param>
    /// <returns></returns>
    bool CheckConfiguration(List<TrackEvent> selectedTextEvents, Track selectedTrack, List<string> filePathsNoEx)
    {
        if (selectedTextEvents.Count == 0)
        {
            MessageBox.Show("テキストイベントが選択されていません。");
        }
        else if (selectedTrack == null)
        {
            MessageBox.Show("トラックが選択されていません。");
        }
        else if (selectedTrack.MediaType != MediaType.Audio)
        {
            MessageBox.Show("選択されたトラックは、音声トラックではありません。");
        }
        else if (filePathsNoEx.Count == 0)
        {
            MessageBox.Show("音声ファイルが存在しません。");
        }
        else
        {
            return true;
        }

        return false;
    }

    string MakeDialogMessage(List<TrackEvent> teList, Track selectedTrack, Dictionary<string, string> textDataDic)
    {
        string ans = "音声トラック名 : " + selectedTrack.Name + "  番号 : " + (selectedTrack.Index + 1) + " に\n";
        foreach (TrackEvent trackEvent in teList)
        {
            string text = trackEvent.ActiveTake.Name.Replace(TitleNamePrefix, "");
            ans += "音声ファイル名 : " + textDataDic[text] + "  テキスト内容 : " + text + "\n";
        }

        ans += "を挿入します。";
        return ans;
    }

    void InsertFileAt(Vegas vegas, String fileName, int trackIndex, Timecode cursorPosition)
    {
        Media media = new Media(fileName);
        AudioTrack track = (AudioTrack) vegas.Project.Tracks[trackIndex];
        AudioEvent audioEvent = track.AddAudioEvent(cursorPosition,media.Length);
        audioEvent.AddTake(media.GetAudioStreamByIndex(0));
    }

    TrackEvent[] FindSelectedEvents(Project project)
    {
        List<TrackEvent> selectedEvents = new List<TrackEvent>();
        foreach (Track track in project.Tracks)
        {
            foreach (TrackEvent trackEvent in track.Events)
            {
                if (trackEvent.Selected)
                {
                    selectedEvents.Add(trackEvent);
                }
            }
        }
        return selectedEvents.ToArray();
    }

    Track FindSelectedTrack(Project project)
    {
        foreach (Track track in project.Tracks)
        {
            if (track.Selected)
            {
                return track;
            }
        }
        return null;
    }

    Track FindVoiceroid2Track(Project project)
    {
        foreach (Track track in project.Tracks)
        {
            if (track.Name == Voiceroid2TrackName)
            {
                return track;
            }
        }
        return null;
    }
}
