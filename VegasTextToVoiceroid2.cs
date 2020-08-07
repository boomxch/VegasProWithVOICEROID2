using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Sony.Vegas;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;


public class EntryPoint
{
    /// <summary>
    /// Vegas上で一つだけTextEventを選択し、スクリプトを実行するとVoiceroid2が立ち上がり、テキストの内容を入力するスクリプト
    /// なお、改行や特殊コードは認識しない
    /// </summary>
    private const string TitleNamePrefix = "Sony タイトルおよびテキスト ";

    public void FromVegas(Vegas vegas)
    {
        if (FindSelectedEvents(vegas.Project).Length != 1)
        {
            MessageBox.Show("テキストイベントを一つだけ選択してください。");
            return;
        }

        TrackEvent selectedEvent = FindSelectedEvents(vegas.Project)[0];
        string text = selectedEvent.ActiveTake.Name.Replace(TitleNamePrefix, "");
        StartVoiceroid2(text);
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

    void StartVoiceroid2(string text)
    {
        VRoid2 vr2 = new VRoid2();
        IntPtr vr2hnd = vr2.GetVoiceroid2hWnd();
        if (vr2hnd == IntPtr.Zero)
        {
            MessageBox.Show("VOCICEROID2を起動してください。");
            return;
        }

        vr2.talk(vr2hnd, text);
    }
}

/// <summary>
/// https://hgotoh.jp/wiki/doku.php/documents/voiceroid/tips/tips-003 様より
/// </summary>
public class VRoid2
{
    // VOICEROID2 EDITOR ウインドウハンドル検索
    public IntPtr GetVoiceroid2hWnd()
    {
        IntPtr hWnd = IntPtr.Zero;

        string winTitle1 = "VOICEROID2";
        string winTitle2 = winTitle1 + "*";
        int RetryCount = 3;
        int RetryWaitms = 1000;

        for (int i = 0; i < RetryCount; i++)
        {
            Process[] ps = Process.GetProcesses();

            foreach (Process pitem in ps)
            {
                if ((pitem.MainWindowHandle != IntPtr.Zero) &&
                       ((pitem.MainWindowTitle.Equals(winTitle1)) || (pitem.MainWindowTitle.Equals(winTitle2))))
                {
                    hWnd = pitem.MainWindowHandle;
                }
            }
            if (hWnd != IntPtr.Zero) break;
            if (i < (RetryCount - 1)) Thread.Sleep(RetryWaitms);
        }

        return hWnd;
    }

    // テキスト転記と再生ボタン押下
    public void talk(IntPtr hWnd, string talkText)
    {
        if (hWnd == IntPtr.Zero) return;

        AutomationElement ae = AutomationElement.FromHandle(hWnd);
        TreeScope ts1 = TreeScope.Descendants | TreeScope.Element;
        TreeScope ts2 = TreeScope.Descendants;

        // アプリケーションウインドウ
        AutomationElement editorWindow = ae.FindFirst(ts1, new PropertyCondition(AutomationElement.ClassNameProperty, "Window"));

        // 再生ボタン、テキストボックスが配置されているコンテナの名前は“c”
        AutomationElement customC = ae.FindFirst(ts1, new PropertyCondition(AutomationElement.AutomationIdProperty, "c"));

        // テキストボックスにテキストを転記
        AutomationElement textBox = customC.FindFirst(ts2, new PropertyCondition(AutomationElement.AutomationIdProperty, "TextBox"));
        ValuePattern elem1 = textBox.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
        elem1.SetValue(talkText);

        // 再生ボタンを押す。再生ボタンはボタンのコレクション5番目(Index=4)
        AutomationElementCollection buttons = customC.FindAll(ts2, new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, "ボタン"));
        InvokePattern elem2 = buttons[4].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        elem2.Invoke();
    }
}
