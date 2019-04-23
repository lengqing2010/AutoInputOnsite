using System.Configuration;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualBasic;
//Imports Windows.Forms.Application
public class Com 
{
    public string folder_Hattyuu;
    public string folder_Hattyuu_kanryou;
    public string folder_Nouki;
    public string folder_Nouki_kanryou;
    public string user;
    public string password;
    public string url;
    public string pdfPath;
    public string logFileName;
    public List<string> file_list_hattyuu;
    public List<string> file_list_Nouki;
    public bool IsDebug;


    /// INIT
    public Com(string inLogFileName)
    {
        //発注
        folder_Hattyuu = GetAppSetting("Folder_Hattyuu");
        folder_Hattyuu_kanryou = GetAppSetting("Folder_Hattyuu_kanryou");
        //納期
        folder_Nouki = GetAppSetting("Folder_Nouki");
        folder_Nouki_kanryou = GetAppSetting("Folder_Nouki_kanryou");
        //認証情報
        user = GetUser();
        password = GetPassword();
        //Ie Start URL
        url = GetAppSetting("PasswordNyuuryoku");
        //PDF Save Path
        pdfPath = GetAppSetting("Pdf_Path");
        //Debug
        IsDebug = GetAppSetting("Debug").ToUpper() == "TRUE";

        logFileName = inLogFileName;

        //フォルダ読み
        //発注
        CreateDirectoryCover(folder_Hattyuu);
        CreateDirectoryCover(folder_Hattyuu_kanryou);
        //納期
        CreateDirectoryCover(folder_Nouki);
        CreateDirectoryCover(folder_Nouki_kanryou);
        //PDF
        CreateDirectoryCover(pdfPath);

        file_list_hattyuu = GetAllFiles(folder_Hattyuu, "*.csv");
        file_list_Nouki = GetAllFiles(folder_Nouki, "*.csv");
    }

    #region Windows DLL
        [DllImport("winmm.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern long timeGetTime();
        [DllImport("kernel32", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern void Sleep(long dwMilliseconds);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string lclassName, string windowTitle);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, StringBuilder lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        private enum WM : int
        {
            SETTEXT = 0xC
        }
        private enum BM : int
        {
            CLICK = 0xF5
        }
	#endregion

	#region Sleep Function
	    /// <summary>
	    /// 共通Sleep 現状仕様無し
	    /// </summary>
	    /// <param name="Interval"></param>
	    /// <remarks></remarks>
	    public static void Sleep5(int Interval) 
	    {
		    DateTime __time = DateTime.Now;
		    long __Span = Interval * 10000;
		    while (System.Convert.ToInt32(DateTime.Now.Ticks - __time.Ticks) < __Span) 
		    {
			    Application.DoEvents();
			    System.Threading.Thread.Sleep(0);
		    }
	    }
	#endregion

    #region Config
        //Get Appconfig key value
        public static string GetAppSetting(string key)
        {
            return System.Convert.ToString(ConfigurationManager.AppSettings[key].ToString()); ;
        }

        //Get user
        public static string GetUser()
        {
            string user = GetAppSetting("User");
            if (string.IsNullOrEmpty(user))
            {
                user = @"china\shil2";
            }
            return user;
        }

        //Get Password
        public static string GetPassword()
        {
            string password = GetAppSetting("Password");
            if (string.IsNullOrEmpty(password))
            {
                password = @"qwer@123";
            }
            return password;
        }
    #endregion

    #region Windows Form 自動追加内容
        /// Windows Form 自動追加内容
        public void WindowsInputThread()
        {
            //認証情報
            user = GetUser();
            password = GetPassword();

            int err_kaisuu_close;
            err_kaisuu_close =int.Parse(GetAppSetting("err_kaisuu_close"));

            int errIndex = 0;

            int hatyuu_list_idx = 0;
            //int nouki_list_idx = 0;

            IntPtr hWnd = default(IntPtr);
            while (hWnd == IntPtr.Zero)
            {
                System.Threading.Thread.Sleep(0);

                hWnd = IntPtr.Zero;
                //認証情報
                if (hWnd == IntPtr.Zero)
                {
                    hWnd = FindWindow("#32770", "windows セキュリティ");
                    if (hWnd != IntPtr.Zero)
                    {
                        NewNinnsyou(hWnd);
                    }
                    hWnd = IntPtr.Zero;
                }

                hWnd = IntPtr.Zero;
                //ファイルアップロード
                if (hWnd == IntPtr.Zero)
                {
                    hWnd = FindWindow("#32770", "アップロードするファイルの選択");
                    if (hWnd != IntPtr.Zero)
                    {
                        NewFileSantaku(hWnd, file_list_hattyuu[hatyuu_list_idx]);
                        hatyuu_list_idx++;
                        errIndex = 0;
                    }
                    hWnd = IntPtr.Zero;
                }

                hWnd = IntPtr.Zero;
                //プログラムを終了します
                if (hWnd == IntPtr.Zero)
                {
                    hWnd = FindWindow("#32770", "Internet Explorer");
                    if (hWnd != IntPtr.Zero)
                    {
                        NewStopErrDg(hWnd);
                        Com.Sleep5(2000);
                        NewStopErrDg(hWnd);
                        Com.Sleep5(2000);

                        errIndex = errIndex + 1;

                        if (errIndex > err_kaisuu_close)
                        {
                            MessageBox.Show("帳票Download エラーしました");
                            System.Environment.Exit(0);
                            //if (MessageBox.Show("帳票Download エラーしました、終了ですか？", "Confirm Message", MessageBoxButtons.OK, MessageBoxIcon.Question) == DialogResult.OK)
                            //{
                            //    System.Environment.Exit(0);
                            //}
                        }


                    }
                    hWnd = IntPtr.Zero;
                }

                hWnd = IntPtr.Zero;
                System.Windows.Forms.Application.DoEvents();
                Sleep5(10);
            }
        }

        /// プログラムを終了します
        public static void NewStopErrDg(IntPtr hWnd)
        {
            IntPtr DirectUIHWND = FindWindowEx(hWnd, IntPtr.Zero, "DirectUIHWND", string.Empty);
            IntPtr CtrlNotifySink1 = FindWindowEx(DirectUIHWND, IntPtr.Zero, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink2 = FindWindowEx(DirectUIHWND, CtrlNotifySink1, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink3 = FindWindowEx(DirectUIHWND, CtrlNotifySink2, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink4 = FindWindowEx(DirectUIHWND, CtrlNotifySink3, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink5 = FindWindowEx(DirectUIHWND, CtrlNotifySink4, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink6 = FindWindowEx(DirectUIHWND, CtrlNotifySink5, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink7 = FindWindowEx(DirectUIHWND, CtrlNotifySink6, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink8 = FindWindowEx(DirectUIHWND, CtrlNotifySink7, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink9 = FindWindowEx(DirectUIHWND, CtrlNotifySink8, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink10 = FindWindowEx(DirectUIHWND, CtrlNotifySink9, "CtrlNotifySink", string.Empty);
            IntPtr hEdit2 = FindWindowEx(CtrlNotifySink10, IntPtr.Zero, "Button", "プログラムを終了します");
            SendMessage(hEdit2, (System.Int32)BM.CLICK, 0, null);
        }

        /// 認証 Win7
        public void NewNinnsyou(IntPtr hWnd)
        {
            IntPtr DirectUIHWND = FindWindowEx(hWnd, IntPtr.Zero, "DirectUIHWND", string.Empty);
            IntPtr CtrlNotifySink1 = FindWindowEx(DirectUIHWND, IntPtr.Zero, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink2 = FindWindowEx(DirectUIHWND, CtrlNotifySink1, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink3 = FindWindowEx(DirectUIHWND, CtrlNotifySink2, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink4 = FindWindowEx(DirectUIHWND, CtrlNotifySink3, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink5 = FindWindowEx(DirectUIHWND, CtrlNotifySink4, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink6 = FindWindowEx(DirectUIHWND, CtrlNotifySink5, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink7 = FindWindowEx(DirectUIHWND, CtrlNotifySink6, "CtrlNotifySink", string.Empty);
            IntPtr CtrlNotifySink8 = FindWindowEx(DirectUIHWND, CtrlNotifySink7, "CtrlNotifySink", string.Empty);
            IntPtr hEdit = FindWindowEx(CtrlNotifySink7, IntPtr.Zero, "Edit", string.Empty);
            SendMessage(hEdit, (System.Int32)WM.SETTEXT, 0, new StringBuilder(user));
            IntPtr hEdit1 = FindWindowEx(CtrlNotifySink8, IntPtr.Zero, "Edit", string.Empty);
            SendMessage(hEdit1, (System.Int32)WM.SETTEXT, 0, new StringBuilder(password));
            IntPtr hEdit2 = FindWindowEx(CtrlNotifySink3, IntPtr.Zero, "Button", "OK");
            SendMessage(hEdit2, (System.Int32)BM.CLICK, 0, null);
        }

        /// ファイル選択
        public void NewFileSantaku(IntPtr hWnd, string fileName)
        {
            Sleep5(200);
        reFind:
            //ある場合、違うFORM Finded
            if (FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", string.Empty) == IntPtr.Zero)
            {
                hWnd = FindWindow("#32770", "アップロードするファイルの選択");
            }
            IntPtr hComboBoxEx = FindWindowEx(hWnd, IntPtr.Zero, "ComboBoxEx32", string.Empty);
            IntPtr hComboBox = FindWindowEx(hComboBoxEx, IntPtr.Zero, "ComboBox", string.Empty);
            IntPtr hEdit = FindWindowEx(hComboBox, IntPtr.Zero, "Edit", string.Empty);
            while (!(IsWindowVisible(hEdit)))
            {
                Sleep5(1);
                goto reFind;
            }
            SendMessage(hEdit, (System.Int32)WM.SETTEXT, 0, new StringBuilder(folder_Hattyuu + fileName));
            IntPtr hButton = FindWindowEx(hWnd, IntPtr.Zero, "Button", "開く(&O)");
            SendMessage(hButton, (System.Int32)BM.CLICK, 0, null);
        }
    #endregion

    #region DEBUG
        //MESSAGE 追加
	    public void AddMsg(string msg) 
	    {
		    if (IsDebug) 
		    {
			    WriteLog(DateTime.Now.ToString("yy/MM/dd HH:mm:ss") + "⇒:" + msg, logFileName);
		    }
	    }
        //ログ出力
        private void WriteLog(string Msg, string tmpfileName)
        {
            string varAppPath = "";
            varAppPath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "log";
            System.IO.Directory.CreateDirectory(varAppPath);
            string head = "";
            head = System.DateTime.Now.Hour.ToString() + ":" + System.DateTime.Now.Minute.ToString();
            Msg = head + System.Environment.NewLine + Msg + System.Environment.NewLine;
            string strFile = "";
            strFile = varAppPath + @"\" + tmpfileName + ".log";
            System.IO.StreamWriter SW = default(System.IO.StreamWriter);
            SW = new System.IO.StreamWriter(strFile, true);
            SW.WriteLine(Msg);
            SW.Flush();
            SW.Close();
        }

    #endregion
        
    #region File 

        //Create Directory
        public static void CreateDirectoryCover(string path)
        {
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
        }
        public static void MoveFile(string acPath, string kanryouPath)
        {
            if (System.IO.File.Exists(kanryouPath))
            {
                System.IO.File.Move(kanryouPath, kanryouPath + ".bk." + DateTime.Now.ToString("yyyyMMddHHmmss"));
            }
            Sleep5(500);
            System.IO.File.Move(acPath, kanryouPath);
        }

        //ファイル取得 Order By file name
        private List<string> GetAllFiles(string strDirect, string exName)
        {
            List<string> flList = new List<string>();
            if (!(ReferenceEquals(strDirect, null)))
            {
                System.IO.FileInfo mFileInfo = default(System.IO.FileInfo);
                System.IO.DirectoryInfo mDirInfo = new System.IO.DirectoryInfo(strDirect);
                try
                {
                    System.IO.FileInfo[] files = mDirInfo.GetFiles(exName);
                    Array.Sort<System.IO.FileInfo>(files, delegate(System.IO.FileInfo a, System.IO.FileInfo b)
                    {
                        // ファイルサイズで昇順ソート
                        //return (int)(a.Length - b.Length);
                        // ファイル名でソート
                        return a.Name.CompareTo(b.Name);
                    });
                    foreach (System.IO.FileInfo tempLoopVar_mFileInfo in files)
                    {
                        mFileInfo = tempLoopVar_mFileInfo;
                        flList.Add(mFileInfo.Name);
                    }
                }
                catch (System.IO.DirectoryNotFoundException)
                {
                }
            }
            return flList;
        }
    #endregion

    #region IE Dom



    //Get Frame
    public mshtml.HTMLWindow2 GetFrameByName(ref SHDocVw.InternetExplorerMedium WebOrWindow, string name)
    {
        SleepAndWaitComplete(WebOrWindow);
        try
        {
            mshtml.HTMLDocument Doc = (mshtml.HTMLDocument)WebOrWindow.Document;
            int length = System.Convert.ToInt32(Doc.frames.length);
            mshtml.FramesCollection frames = Doc.frames;
            mshtml.HTMLWindow2 frm = default(mshtml.HTMLWindow2);
            //Object 必要、 frames.item(i) iはInertger時 エラー
            object i = null;
            for (i = 0; (int)i <= length - 1; i = (int)i + 1)
            {
                //frm =
                if (((mshtml.HTMLWindow2)(frames.item(i))).name == name)
                {
                    return ((mshtml.HTMLWindow2)(frames.item(i)));
                }
            }
            object wd = null;
            for (i = 0; (int)i <= length - 1; i = (int)i + 1)
            {
                frm = (mshtml.HTMLWindow2)(frames.item(i));
                wd = GetFrameByName(frm, name);
                if (wd != null)
                {
                    return ((mshtml.HTMLWindow2)wd);
                }
            }
        }
        catch (Exception)
        {
        }
        return null;
    }

    //Get Frame
    public mshtml.HTMLWindow2 GetFrameByName(mshtml.HTMLWindow2 webApp, string name)
    {
        mshtml.HTMLDocument Doc = (mshtml.HTMLDocument)webApp.document;
        int length = System.Convert.ToInt32(Doc.frames.length);
        mshtml.FramesCollection frames = Doc.frames;
        object i = null;
        for (i = 0; (int)i <= length - 1; i = (int)i + 1)
        {
            mshtml.HTMLWindow2 frm = (mshtml.HTMLWindow2)(frames.item(i));
            if (frm.name == name)
            {
                return frm;
            }
        }
        for (i = 0; (int)i <= length - 1; i = (int)i + 1)
        {
            mshtml.HTMLWindow2 frm = (mshtml.HTMLWindow2)(frames.item(i));
            object wd = GetFrameByName(frm, name);
            if (wd != null)
            {
                return ((mshtml.HTMLWindow2)wd);
            }
        }
        return null;
    }

    //Sleep And Wait IE Complete
    public void SleepAndWaitComplete(SHDocVw.InternetExplorerMedium webApp, int tmOut = 100)
    {
        Sleep5(100);
        for (int i = 0; i <= 10; i++)
        {
            while (!(webApp.ReadyState == SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE && !webApp.Busy))
            {
                System.Windows.Forms.Application.DoEvents();
                Sleep5(System.Convert.ToInt32((double)tmOut / 10));
                i = 0;
            }
            while (!(((mshtml.HTMLDocument)webApp.Document).readyState.ToLower() == "complete"))
            {
                System.Windows.Forms.Application.DoEvents();
                Sleep5(System.Convert.ToInt32((double)tmOut / 10));
                i = 0;
            }
            try
            {
                mshtml.HTMLDocument Doc = (mshtml.HTMLDocument)webApp.Document;
                int length = System.Convert.ToInt32(Doc.frames.length);
                mshtml.FramesCollection frames = Doc.frames;
                for (int j = 0; j <= length - 1; j++)
                {
                    mshtml.HTMLWindow2 frm = (mshtml.HTMLWindow2)(frames.item(j));
                    while (!(frm.document.readyState.ToLower() == "complete"))
                    {
                        System.Windows.Forms.Application.DoEvents();
                        Sleep5(1);
                    }
                }
            }
            catch (Exception)
            {
            }
        }
    }

    //Get element and do soming
    public mshtml.IHTMLElement GetElementByDo(ref SHDocVw.InternetExplorerMedium webApp, string fraName, string tagName, string keyName, string keyTxt)
    {
        if (ReferenceEquals(webApp, null))
        {
            return null;
        }
        SleepAndWaitComplete(webApp);
        mshtml.HTMLDocument Doc = (mshtml.HTMLDocument)webApp.Document;
        mshtml.IHTMLElementCollection eles = default(mshtml.IHTMLElementCollection);
        if (fraName == "")
        {
            eles = Doc.getElementsByTagName(tagName);
        }
        else
        {
            mshtml.HTMLWindow2 fra = GetFrameWait(ref webApp, fraName);
            Doc = (mshtml.HTMLDocument)fra.document;
            eles = Doc.getElementsByTagName(tagName);
        }

        SleepAndWaitComplete(webApp);
        foreach (mshtml.IHTMLElement ele in eles)
        {
            try
            {
                if (keyName == "innertext")
                {
                    if (ele.innerText == keyTxt)
                    {
                        return ele;
                    }
                }
                else
                {
                    String myKey;
                    myKey = Convert.ToString(ele.getAttribute(keyName));
                    if (myKey.ToString() == keyTxt)
                    {
                        return ele;
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        return null;
    }

    //Get Frame Wait
    public mshtml.HTMLWindow2 GetFrameWait(ref SHDocVw.InternetExplorerMedium webApp, string name)
    {
        mshtml.HTMLWindow2 fra = default(mshtml.HTMLWindow2);
        for (int i = 1; i <= 10; i++)
        {
            fra = GetFrameByName(ref webApp, name);
            if (fra != null)
            {
                return fra;
            }
            Sleep5(200);
        }
        return null;
    }

    //Get Element
    public mshtml.IHTMLElement GetElementBy(ref SHDocVw.InternetExplorerMedium webApp, string fraName, string tagName, string keyName, string keyTxt)
    {
        SleepAndWaitComplete(webApp);
        mshtml.IHTMLElement tmpHTMLELE;
        tmpHTMLELE = null;
        try
        {
            while (ReferenceEquals(tmpHTMLELE, null))
            {
                tmpHTMLELE = GetElementByDo(ref webApp, fraName, tagName, keyName, keyTxt);
                System.Windows.Forms.Application.DoEvents();
                Sleep5(1);
            }
            return tmpHTMLELE;
            //return GetElementByDo(ref webApp, fraName, tagName, keyName, keyTxt);
        }
        catch (Exception)
        {
            return GetElementByDo(ref webApp, fraName, tagName, keyName, keyTxt);
        }
    }

    // GetElement
    public mshtml.IHTMLElement GetElement(mshtml.HTMLWindow2 fra, string tagName, string keyName, string keyTxt)
    {
        mshtml.IHTMLElementCollection eles = default(mshtml.IHTMLElementCollection);
        mshtml.HTMLDocument Doc = default(mshtml.HTMLDocument);
        Doc = (mshtml.HTMLDocument)fra.document;
        eles = Doc.getElementsByTagName(tagName);
        foreach (mshtml.IHTMLElement ele in eles)
        {
            try
            {
                if (keyName == "innertext")
                {
                    if (ele.innerText == keyTxt)
                    {
                        return ele;
                    }
                }
                else
                {
                    if (ele.getAttribute(keyName).ToString() == keyTxt)
                    {
                        return ele;
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        return null;
    }

    #endregion
}