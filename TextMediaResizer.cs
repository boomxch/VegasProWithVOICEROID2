using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Sony.Vegas;

public class EntryPoint
{
    /// <summary>
    /// テキストメディアを文字数に応じてメディアサイズを変える
    /// </summary>
    private const double TimePerWordCountms = 115;
    private const double TimePerCommams = 200;
    private const double ExtraTimems = 500;
    private const string TitleNamePrefix = "Sony タイトルおよびテキスト ";

    //VegasだとvarとかC#3.0のLinqのメソッド形式の入力がエラー
    public void FromVegas(Vegas vegas)
    { 
        // 選択されたテキストイベントを取得(Take名が指定の名称プレフィックスと一致している場合のみ)
        List<TrackEvent> selectedTextEvents = new List<TrackEvent>();
        foreach (TrackEvent ev in FindSelectedEvents(vegas.Project))
        {
            if (ev.ActiveTake.Name.IndexOf(TitleNamePrefix, StringComparison.Ordinal) != -1)
                selectedTextEvents.Add(ev);
        }

        // 条件を満たしているかどうかのチェック
        if (!CheckConfiguration(selectedTextEvents))
            return;

        Dictionary<TrackEvent,double> selectedTextEventsTimesms = new Dictionary<TrackEvent, double>();
        foreach (TrackEvent ev in selectedTextEvents)
        {
            selectedTextEventsTimesms.Add(ev, CalculateEventLengthBasedOnWordCount(ev));
        }

        // ダイアログのメッセージ生成
        string message = MakeDialogMessage(selectedTextEventsTimesms);

        // 確認ダイアログ生成
        DialogResult result = MessageBox.Show(message,
            "確認",
            MessageBoxButtons.OKCancel,
            MessageBoxIcon.Exclamation,
            MessageBoxDefaultButton.Button2);

        if (result != DialogResult.OK)
            return;

        // 処理実行
        foreach (TrackEvent textEvent in selectedTextEvents)
        {
            textEvent.Length = new Timecode(selectedTextEventsTimesms[textEvent]);
        }
    }

    /// <summary>
    /// スクリプトを実行する条件を満たしているかを確認する
    /// </summary>
    /// <param name="selectedTextEvents">選択されたテキストイベント</param>
    /// <returns></returns>
    bool CheckConfiguration(List<TrackEvent> selectedTextEvents)
    {
        if (selectedTextEvents.Count == 0)
        {
            MessageBox.Show("テキストイベントが選択されていません。");
        }
        else
        {
            return true;
        }

        return false;
    }

    string MakeDialogMessage(Dictionary<TrackEvent, double> evTimes)
    {
        string ans = String.Empty;
        foreach (KeyValuePair<TrackEvent,double> evTime in evTimes)
        {
            ans += "テキスト : " + evTime.Key.ActiveTake.Name.Replace(TitleNamePrefix,"") + "  長さ : " + evTime.Value / 1000 + " 秒\n";
        }

        ans += "に変更します。";
        return ans;
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

    /// <summary>
    /// テキストイベントの文字数に応じて、イベントの長さを求めるメソッド
    /// https://www.pine4.net/Memo/Article/Archives/424
    /// </summary>
    /// <param name="textEvent">変更するTrackEvent</param>
    /// <returns>length of event</returns>
    double CalculateEventLengthBasedOnWordCount(TrackEvent textEvent)
    {
        string takeName = textEvent.ActiveTake.Name.Remove(0, TitleNamePrefix.Length);
        IFELanguage ifelang = null;
        string allKanaName=String.Empty;
        try
        {
            ifelang = Activator.CreateInstance(Type.GetTypeFromProgID("MSIME.Japan")) as IFELanguage;
            int hr = ifelang.Open();

            if (hr != 0)
                throw Marshal.GetExceptionForHR(hr);

            hr = ifelang.GetPhonetic(takeName, 1, -1, out allKanaName);
            ifelang.Close();
        }
        catch (COMException ex)
        {
            if (ifelang != null)
                ifelang.Close();
        }
        
        int commaNumber = CountChar(allKanaName, ',') + CountChar(allKanaName, '、');
        int spaceNumber = CountChar(allKanaName, ' ') + CountChar(allKanaName, '　');
        return (allKanaName.Length - commaNumber - spaceNumber) * TimePerWordCountms + commaNumber * TimePerCommams + ExtraTimems;
    }

    int CountChar(string s, char c)
    {
        return s.Length - s.Replace(c.ToString(), "").Length;
    }

    // IFELanguage2 Interface ID
    //[Guid("21164102-C24A-11d1-851A-00C04FCC6B14")]
    [ComImport]
    [Guid("019F7152-E6DB-11d0-83C3-00C04FDDB82E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IFELanguage
    {
        int Open();
        int Close();
        int GetJMorphResult(uint dwRequest, uint dwCMode, int cwchInput, [MarshalAs(UnmanagedType.LPWStr)] string pwchInput, IntPtr pfCInfo, out object ppResult);
        int GetConversionModeCaps(ref uint pdwCaps);
        int GetPhonetic([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
        int GetConversion([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
    }
}
