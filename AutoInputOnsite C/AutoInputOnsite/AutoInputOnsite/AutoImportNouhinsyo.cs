using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;


using System.Runtime.InteropServices;
using System.Threading;

using Microsoft.VisualBasic;
using Microsoft.CSharp;


namespace AutoInputOnsite
{
    public partial class AutoImportNouhinsyo : Form
    {
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
        private static extern bool IsWindowVisible(
            IntPtr hWnd);
        private enum WM : int
        {
            SETTEXT = 0xC
        }
        private enum BM : int
        {
            CLICK = 0xF5
        }
        [DllImport("user32", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern long GetWindowThreadProcessId(long hWnd, long lpdwProcessId);
        [DllImport("kernel32", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern long OpenProcess(long dwDesiredAccess, bool bInheritHandle, long dwProcessId);
        [DllImport("kernel32", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern long TerminateProcess(long hProcess, long uExitCode);
        [DllImport("kernel32", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern long CloseHandle(long hObject);
        #endregion
        public decimal _ProBar = 0;
        public decimal lv1;
        public decimal lv2;
        public bool IeVisible = true;
        public int ProBar
        {
            get
            {
                if (_ProBar < 100)
                {
                    return (int)((_ProBar));
                }
                else
                {
                    return 100;
                }
            }
            set
            {
                _ProBar = value;
            }
        }
        public void AddProBar(decimal x)
        {
            _ProBar += x;
        }
        public bool insatu = false;
        public struct FCW_INFO
        {
            public string fcw_datafile;
            public string fcw_formfile;
            public string Action;
            public string url1;
            public string url2;
        }
        #region 自動実行


        public SHDocVw.InternetExplorerMedium Ie;
        public Com Pub_Com;

        //実行
        private void btnRun_Click(System.Object sender, System.EventArgs e)
        {
            DoAll();
        }

        //実行 MAIN
        public void DoAll()
	    {
		
		    //SavePdf("aaaa.csv")
		    Pub_Com = new Com("納期指定発注" + DateTime.Now.ToString("yyyyMMddHHmmss"));

		    if (Pub_Com.file_list_Nouki.Count == 0)
		    {
			    ProBar = 100;
			    return;
		    }
		
		    lv1 = System.Convert.ToDecimal(90 / Pub_Com.file_list_Nouki.Count);
		    lv2 = lv1 / 15;
		
		    //一回目
		    bool firsOpenKbn = true;
			
		    object authHeader = "Authorization: Basic " + 
			    Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(string.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\\r\\n";
		
		    //＊＊＊ OnSiteパスワード入力画面
            Ie.Navigate(Pub_Com.url, null, null, null, authHeader);		
		    Ie.Silent = true;
		    Ie.Visible = IeVisible;
		
		    ProBar = 5;
		    //＊＊＊ログイン
		    DoStep1_Login();

            //CSV ファイルｓ 取込
		    ProBar = 10;		
		    for (int fileIdx = 0; fileIdx <= Pub_Com.file_list_Nouki.Count - 1; fileIdx++)
		    {
			
			    string csvFileName = System.Convert.ToString(Pub_Com.file_list_Nouki[fileIdx].ToString().Trim());
			    string[] csvNameSplitor = csvFileName.Split('-');

                string 事業所 = csvNameSplitor[0];
                string 得意先 = csvNameSplitor[1];
                string 下店 = csvNameSplitor[2];
                string 現場名 = csvNameSplitor[3];
                string 備考 = csvNameSplitor[4];
                string 日付連番 = csvNameSplitor[5];
		
			
			    //一回目ではなく 実行
			    if (firsOpenKbn == false)
			    {
				    Pub_Com.GetElementBy(ref Ie, "fraHead", "input", "value", "絞込検索").click();
				    Pub_Com.SleepAndWaitComplete(Ie);
			    }
			
			    firsOpenKbn = false;
			    AddProBar(lv2); //1
			    Pub_Com.AddMsg("取込：" + Pub_Com.file_list_Nouki[fileIdx].ToString().Trim());
			
			
			    //見積検索
			    Pub_Com.AddMsg("見積検索");		
			    DoStep1_PoupuSentaku(事業所, 得意先, 下店, 現場名, 備考, 日付連番, csvFileName);
			    Pub_Com.SleepAndWaitComplete(Ie);
			    AddProBar(lv2); //2
			
			    //納期日設定
			    if (!DoStep2_Set())
			    {
				    continue;
			    }
			
			    Pub_Com.SleepAndWaitComplete(Ie);
			    Pub_Com.SleepAndWaitComplete(Ie);
			    Pub_Com.SleepAndWaitComplete(Ie);
			
			
			    //該当データがありません NEXT
			    mshtml.HTMLWindow2 fraTmp = Pub_Com.GetFrameByName(ref Ie, "fraHyou");
			    if (fraTmp != null)
			    {
				    if (fraTmp.document.body.innerText.IndexOf("該当データがありません") >= 0)
				    {
					    continue;
				    }
			    }
			
			
			    //CSVファイル内容取込
			    string[] csvDataLines = System.IO.File.ReadAllLines(Pub_Com.folder_Nouki + csvFileName);
			    string code = "";
			    string nouki = "";
			    AddProBar(lv2); //3			
			
			    mshtml.HTMLWindow2 fra = Pub_Com.GetFrameWait(ref Ie, "fraMitBody");
			    mshtml.HTMLDocument Doc = (mshtml.HTMLDocument) fra.document;
			    mshtml.IHTMLElementCollection eles = Doc.getElementsByTagName("input");			
			    //Radio 明細Key
			    mshtml.IHTMLElementCollection cbEles = Doc.getElementsByName("strMeisaiKey");
			    //指定納期
			    mshtml.IHTMLElementCollection nouhinDateEles = Doc.getElementsByName("strSiteiNouhinDate");
			
			    int csvIdx = 0;
			    int sameCdSuu = 0;
			    int gamenSameCdSuu = 0;
			
			    //CSV LINES
			    for (int csvLinesIdx = 0; csvLinesIdx <= csvDataLines.Length - 1; csvLinesIdx++)
			    {
				
				    if (!string.IsNullOrEmpty(csvDataLines[csvLinesIdx].Trim()))
				    {
					    //コード 納期
					    code = System.Convert.ToString(csvDataLines[csvLinesIdx].Split(',')[1].Trim());
					    nouki = System.Convert.ToString((System.Convert.ToDateTime(csvDataLines[csvLinesIdx].Split(',')[2].Trim())).ToString("yyyy/MM/dd"));
					
					    sameCdSuu = 0;
					    gamenSameCdSuu = 0;
					
					    for (csvIdx = 0; csvIdx <= csvLinesIdx; csvIdx++)
					    {
						    if (csvDataLines[csvIdx].Split(',')[1].Trim() == code)
						    {
							    sameCdSuu++;
						    }
					    }
					
					    for (int i = 0; i <= cbEles.length - 1; i++)
					    {
						
						    mshtml.IHTMLElement cb = (mshtml.IHTMLElement) (cbEles.item(i));
						    mshtml.IHTMLTableRow tr = (mshtml.IHTMLTableRow) cb.parentElement.parentElement;
						    mshtml.HTMLTableCell td = (mshtml.HTMLTableCell) (tr.cells.item(1));
						    mshtml.IHTMLTable table = (mshtml.IHTMLTable) cb.parentElement.parentElement.parentElement.parentElement;
						
						    bool isHaveDate = false;
						
						    if (td.innerText == code)
						    {
							
							    gamenSameCdSuu++;
							
							    if (sameCdSuu == gamenSameCdSuu)
							    {
								
								    mshtml.IHTMLSelectElement sel = (mshtml.IHTMLSelectElement) (nouhinDateEles.item(i));
								
								    for (int j = 0; j <= sel.length - 1; j++)
								    {
									
									    mshtml.IHTMLOptionElement opEle = (mshtml.IHTMLOptionElement) (sel.item(j));
									
									    if (opEle.value.IndexOf(nouki) > 0)
									    {
										    opEle.selected = true;
										    isHaveDate = true;
										    break;
									    }
									
								    }
							    }
							    else
							    {
								    continue;
							    }
							
							    if (!isHaveDate)
							    {
								    MessageBox.Show("コード：[" + code + "] 納品希望日：[" + nouki + "]がありません");
								    return;
								
							    }
						    }
						
					    }
					
				    }
				
			    }
			    AddProBar(lv2); //4			
			    Pub_Com.GetElementBy(ref Ie, "fraMitBody", "select", "name", "strBukkenKbn").setAttribute("value", "01");
			    Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "発　注").click();
			    Pub_Com.SleepAndWaitComplete(Ie);
			    Pub_Com.SleepAndWaitComplete(Ie);
			    Pub_Com.SleepAndWaitComplete(Ie);
			    AddProBar(lv2); //5			
			
			    Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "発注結果照会へ").click();
			    Pub_Com.SleepAndWaitComplete(Ie);
			    AddProBar(lv2); //6
			
                //PDF 印刷
			    if (insatu)
			    {

                    SHDocVw.InternetExplorerMedium childIe= default(SHDocVw.InternetExplorerMedium);
				    int RebackKaisu = -1;
				
    Reback:
				
				    RebackKaisu++;
				    Com.Sleep5(1000);
				
				    //前回印刷画面 Close
				    ClosePrintPage();
				    Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "結果印刷").click();

                    int wait_print;
                    wait_print = int.Parse(Com.GetAppSetting("wait_print"));
                    AutoResetEvent myEvent = new AutoResetEvent(false);
                    myEvent.WaitOne(wait_print*1000);
                    myEvent.Close();
                    
				    try
				    {
					    //印刷画面取得する
					    childIe = GetPrintPage();
				    }
				    catch (Exception)
				    {
					
				    }
				
				    Com.Sleep5(1000);
				
				    //IE エラー判定する
				    if (GetErrCon() == false)
				    {
					    Com.Sleep5(1000);
					    if (RebackKaisu <= 1)
					    {
						    goto Reback;
					    }
					    else
					    {
                            if (MessageBox.Show("帳票Download エラーしました、終了ですか？", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                            {
                                System.Environment.Exit(0);
                            }

					    }
					
				    }
				
				    string flName = Pub_Com.pdfPath + csvFileName;
				
				    try
				    {
					    bool rtv = GetFcwInfo(childIe, flName);
					    ClosePrintPage();
					    if (rtv == false)
					    {
						    if (RebackKaisu <= 1)
						    {
							    goto Reback;
						    }
						    else
						    {
                                if (MessageBox.Show("帳票Download エラーしました、終了ですか？", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                                {
                                    System.Environment.Exit(0);
                                }
						    }
					    }
					
				    }
				    catch (Exception)
				    {
					    ClosePrintPage();
					    if (RebackKaisu <= 1)
					    {
						    goto Reback;
					    }
					    else
					    {
                            if (MessageBox.Show("帳票Download エラーしました、終了ですか？", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                            {
                                System.Environment.Exit(0);
                            }
					    }
				    }
				
			    }
			
			    AddProBar(lv2); //10
			    Pub_Com.AddMsg("移動CSV：" + csvFileName + "→" + Pub_Com.folder_Nouki_kanryou);

                Com.MoveFile(Pub_Com.folder_Nouki + csvFileName, Pub_Com.folder_Nouki_kanryou + csvFileName);
			    AddProBar(lv2); //11
			
			    Pub_Com.GetElementBy(ref Ie, "fraMitMenu", "a", "innertext", "[見積一覧を再表示]").click();
			    Pub_Com.SleepAndWaitComplete(Ie);
			
		    }
		
		
		    ProBar = 100;
		
	    }

        //印刷情報を取得する
        public bool GetFcwInfo(SHDocVw.InternetExplorerMedium cIe, string flName)
        {

            try
            {
                string Action;
                string fcw_formfile = "";
                string fcw_datafile = "";
                string url1 = "";
                string url2 = "";

                url2 = "";
                mshtml.HTMLFormElement form;
                mshtml.HTMLDocument doc;
                doc = (mshtml.HTMLDocument) cIe.Document;
                mshtml.HTMLDocument tmdoc;
                tmdoc = (mshtml.HTMLDocument) doc.body.document;
                form = (mshtml.HTMLFormElement)tmdoc.forms.item(0);
                mshtml.IHTMLElementCollection  forms;
                forms = (mshtml.IHTMLElementCollection)tmdoc.forms;
                form =(mshtml.HTMLFormElement) forms.item(null,0);
                Action = System.Convert.ToString(form.action);

                if (string.IsNullOrEmpty(Action))
                {
                    return false;
                }

                mshtml.HTMLInputElement hid_fcw_formfile = default(mshtml.HTMLInputElement);
                hid_fcw_formfile = (mshtml.HTMLInputElement)(((mshtml.HTMLDocument)form.document).getElementsByName("fcw-formfile").item(0));
                fcw_formfile = System.Convert.ToString(hid_fcw_formfile.value);

                mshtml.HTMLInputElement hid_fcw_datafile = default(mshtml.HTMLInputElement);
                hid_fcw_datafile = (mshtml.HTMLInputElement)(((mshtml.HTMLDocument)form.document).getElementsByName("fcw-datafile").item(0));
                fcw_datafile = System.Convert.ToString(hid_fcw_datafile.value);

                url1 = form.action + "?fcw-driver=FCPC&fcw-formdownload=yes&fcw-newsession=yes&fcw-destination=client&fcw-overlay=3&fcw-endsession=yes&fcw-formfile=" + fcw_formfile + "&fcw-datafile=" + fcw_datafile;

                //Print 画面の情報取得
                object bodyStr;
                string bodyTxt = "";
                MSXML2.XMLHTTP Retrieval = new MSXML2.XMLHTTP();
                Retrieval.open("Get", url1, false, "", "");
                Retrieval.send();
                bodyStr = Retrieval.responseBody;
                bodyTxt = System.Convert.ToString(Retrieval.responseText);
                Retrieval = null;

                //帳票のURL取得
                System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex("src=\"(?<src>[^\"]*)\"");
                System.Text.RegularExpressions.MatchCollection mc = reg.Matches(bodyTxt);
                foreach (System.Text.RegularExpressions.Match m in mc)
                {
                    url2 = System.Convert.ToString(m.Value);
                }
                url2 = System.Convert.ToString(url2.Replace("src", "").Replace("\"", "").Replace("=http", "http"));
                reg = null;

                if (System.IO.File.Exists(flName))
                {
                    FileSystem.Rename(flName.Replace(".csv", ".pdf"), flName.Replace(".csv", "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".bk.pdf"));
                }

                //帳票作成
                GetRemoteFiels(url2, flName.Replace(".csv", ".pdf"));

                return true;

            }
            catch (Exception)
            {
                return false;
            }


        }

        /// <summary>
        /// Window Form 自動追加内容
        /// </summary>
        /// <remarks></remarks>
        public bool GetErrCon()
        {
            bool rtv = true;
            IntPtr hWnd = default(IntPtr);
            hWnd = IntPtr.Zero;

            Com.Sleep5(0);

            //プログラムを終了します
            if (hWnd == IntPtr.Zero)
            {
                hWnd = FindWindow("#32770", "Internet Explorer");
                if (hWnd != IntPtr.Zero)
                {
                    Com.NewStopErrDg(hWnd);
                    Com.Sleep5(1000);
                    rtv = false;
                    Com.NewStopErrDg(hWnd);
                    Com.Sleep5(1000);
                }
                hWnd = IntPtr.Zero;
            }

            System.Windows.Forms.Application.DoEvents();
            Com.Sleep5(10);
            return rtv;
        }


        //public bool pub_print_complete_flg = false;

        //public void DocumentComplete(object pDisp, object URL)
        //{


        //    if (URL.ToString().IndexOf("/HatKekkaPrint.asp") + 1 > 0 && URL.ToString().IndexOf("/Redirect.asp") + 1 == 0)
        //    {
        //        //    If InStr(url, "/HatKekkaPrint.asp") > 0 Then

        //        ((SHDocVw.InternetExplorerMedium)pDisp).Stop();

        //        pub_print_complete_flg = true;
        //    }
        //    else
        //    {
        //        pub_print_complete_flg = false;
        //    }

        //}

        /// <summary>
        /// 印刷画面取得
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public SHDocVw.InternetExplorerMedium GetPrintPage()
        {

            SHDocVw.ShellWindows ShellWindows = new SHDocVw.ShellWindows();

        reloIt:

            foreach (SHDocVw.InternetExplorerMedium childIe in ShellWindows)
            {

                Com.Sleep5(0);
                System.Windows.Forms.Application.DoEvents();

                if ((childIe.FullName.ToLower()).ToString().IndexOf("iexplore.exe") + 1 > 0)
                {

                    string url = System.Convert.ToString(childIe.LocationURL.ToString());

                    if (url.IndexOf("/HatKekkaPrint.asp") + 1 > 0 && url.IndexOf("/Redirect.asp") + 1 == 0)
                    {
                        //    If InStr(url, "/HatKekkaPrint.asp") > 0 Then
                        while (childIe.Busy || (childIe.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE))
                        {
                            Application.DoEvents();
                            Com.Sleep5(0);
                        }
                        childIe.Stop();
                        ShellWindows = null;
                        return childIe;
                    }
                    else if (childIe.LocationURL.Contains("servlet"))
                    {

                        while (childIe.Busy || (childIe.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE))
                        {
                            Application.DoEvents();
                            Com.Sleep5(0);
                        }
                        childIe.Stop();
                        childIe.GoBack();
                        while (childIe.Busy || (childIe.ReadyState != SHDocVw.tagREADYSTATE.READYSTATE_COMPLETE))
                        {
                            Application.DoEvents();
                            Com.Sleep5(0);
                        }
                        childIe.Stop();
                        return childIe;
                    }

                }

            }

            Com.Sleep5(1);

            //Return Nothing
            goto reloIt;


        }

        /// <summary>
        /// 印刷画面Close
        /// </summary>
        /// <remarks></remarks>
        public void ClosePrintPage()
        {

            SHDocVw.ShellWindows ShellWindows = new SHDocVw.ShellWindows();

            try
            {
                foreach (SHDocVw.InternetExplorerMedium childIe in ShellWindows)
                {
                    System.Windows.Forms.Application.DoEvents();
                    string filename = System.Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower());
                    if (filename == "iexplore")
                    {
                        if (childIe.LocationURL.Contains("servlet"))
                        {
                            long whWnd = System.Convert.ToInt64(childIe.HWND);
                            childIe.Quit();
                            KillProcess(whWnd);
                            //childIe.HWND
                            Com.Sleep5(500);
                        }
                        else if (childIe.LocationURL.Contains("HatKekkaPrint.asp"))
                        {
                            long whWnd = System.Convert.ToInt64(childIe.HWND);
                            childIe.Quit();
                            KillProcess(whWnd);
                            Com.Sleep5(500);
                        }
                    }
                }
            }
            catch (Exception)
            {

            }
            ShellWindows = null;

        }

        //Close Popup page
        public void CloseChildPage()
        {

            SHDocVw.ShellWindows ShellWindows = new SHDocVw.ShellWindows();

            try
            {
                foreach (SHDocVw.InternetExplorerMedium childIe in ShellWindows)
                {
                    System.Windows.Forms.Application.DoEvents();
                    string filename = System.Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower());
                    if (filename == "iexplore")
                    {
                        if (childIe.LocationURL.Contains("Sougou_Menu.asp"))
                        {
                        }
                        else
                        {
                            long whWnd = System.Convert.ToInt64(childIe.HWND);
                            childIe.Quit();
                            KillProcess(whWnd);
                            Com.Sleep5(500);
                        }
                    }
                }
            }
            catch (Exception)
            {

            }
            ShellWindows = null;

        }

        //Kill Process
        private void KillProcess(long whWnd)
        {
            long lpdwProcessId = 0;
            long hProcessHandle = 0;
            GetWindowThreadProcessId(whWnd, lpdwProcessId);
            hProcessHandle = System.Convert.ToInt64(OpenProcess(0x1F0FFF, true, lpdwProcessId));
            long success = 0;
            if (hProcessHandle != 0)
            {
                success = System.Convert.ToInt64(TerminateProcess(hProcessHandle, 0));
            }
            if (success == 1)
            {
                CloseHandle(hProcessHandle);
            }
        }


        #region SAVE PDF
            //SAVE PDF
            public bool GetRemoteFiels(object RemotePath, object LocalPath)
            {

                object strBody = null;
                object FilePath = null;
                // On Error Resume Next
                //取得流
                strBody = GetBody(RemotePath);

                if ((string)LocalPath == "")
                {
                    return true;
                }
                FilePath = LocalPath;
                //保存文件
                if (SaveToFile(strBody, FilePath) == true && Information.Err().Number == 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            //GetFileName
            public String GetFileName(object RemotePath, object FileName)
            {
                string[] arrTmp;
                object strFileExt = null;
                arrTmp = Strings.Split(System.Convert.ToString(RemotePath), ".");
                strFileExt = arrTmp[arrTmp.Length - 1];
                return FileName + "." + System.Convert.ToString(strFileExt);
            }

            //Get Body
            public Object GetBody(object url)
            {

                MSXML2.XMLHTTP Retrieval = new MSXML2.XMLHTTPClass();
                Retrieval.open("Get", url.ToString() ,false, null, null);
                Retrieval.send();

                return  Retrieval.responseBody;

            }

            //SaveToFile
            public bool SaveToFile(object Stream, object FilePath)
            {
                Application.DoEvents();
                // Try
                ADODB.StreamClass objStream = new ADODB.StreamClass();

                //objStream = Interaction.CreateObject("ADODB.Stream");
                objStream.Type = ADODB.StreamTypeEnum.adTypeBinary ;

                objStream.Open();
                objStream.Write(Stream);
                objStream.SaveToFile(FilePath.ToString(), ADODB.SaveOptionsEnum.adSaveCreateOverWrite);
                objStream.Close();
                objStream = null;
                if (Information.Err().Number != 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }

            }
        #endregion

        //Step 1 LOGIN IN
        public void DoStep1_Login()
        {


            Pub_Com.AddMsg("OnSiteパスワード入力");
            //＊＊＊ OnSiteパスワード入力
            Pub_Com.GetElementBy(ref Ie, "", "input", "name", "strPassWord").setAttribute("value", ConfigurationManager.AppSettings["OnSitePassword"].ToString());
            Pub_Com.GetElementBy(ref Ie, "", "input", "value", "ログオン").click();

            Pub_Com.AddMsg("業務別総合メニュー");
            //'＊＊＊ 業務別総合メニュー
            Pub_Com.GetElementBy(ref Ie, "SubHeader", "a", "innertext", "[見積]").click();


            CloseChildPage();
            Pub_Com.AddMsg("物販明細");
            //'＊＊＊ 物販明細
            Pub_Com.GetElementBy(ref Ie, "Main", "input", "value", "物販明細").click();
            Pub_Com.SleepAndWaitComplete(Ie, 100);
            Pub_Com.SleepAndWaitComplete(Ie, 100);



        }

        //Step 1 新規見積もり
        public void DoStep1_PoupuSentaku(string 事業所, string 得意先, string 下店, string 現場名, string 備考, string 日付連番, string fl)
        {

            try
            {

                Pub_Com.AddMsg("見積検索 POPUP");
                SHDocVw.InternetExplorerMedium cIe = GetPopupWindow("OnSite", "mitSearch.asp");
                while (ReferenceEquals(cIe, null))
                {
                    Com.Sleep5(100);
                    cIe = GetPopupWindow("OnSite", "mitSearch.asp");
                }

                Pub_Com.SleepAndWaitComplete(cIe);
                Pub_Com.GetElementBy(ref cIe, "", "input", "name", "strGenbaMei").innerText = 現場名;
                Pub_Com.GetElementBy(ref cIe, "", "input", "name", "strUriJgy").click();

                Pub_Com.SleepAndWaitComplete(cIe);
                AddProBar(lv2); //12


                Pub_Com.AddMsg("事業所検索 POPUP");
                SHDocVw.InternetExplorerMedium cIe2 = GetPopupWindow("OnSite", "jgyKensaku.asp");
                while (ReferenceEquals(cIe2, null))
                {
                    Com.Sleep5(100);
                    cIe2 = GetPopupWindow("OnSite", "jgyKensaku.asp");
                }

                Pub_Com.SleepAndWaitComplete(cIe2);
                Pub_Com.GetElementBy(ref cIe2, "", "input", "name", "strJgyCd").innerText = 事業所;
                Pub_Com.GetElementBy(ref cIe2, "", "input", "value", "検　索").click();
                Com.Sleep5(1000);
                AddProBar(lv2); //13
                Pub_Com.SleepAndWaitComplete(cIe);
                Pub_Com.GetElementBy(ref cIe, "", "input", "value", "検　索").click();

                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.SleepAndWaitComplete(Ie);
                Com.Sleep5(1000);
                AddProBar(lv2); //14

                //50件以上の場合
                try
                {
                    Com.Sleep5(1000);

                    mshtml.HTMLDocument Doc = default(mshtml.HTMLDocument);
                    mshtml.IHTMLElementCollection eles = default(mshtml.IHTMLElementCollection);
                    Doc = (mshtml.HTMLDocument)Ie.Document;
                    mshtml.HTMLWindow2 fra = Pub_Com.GetFrameWait(ref Ie, "fraHyou");
                    Doc = (mshtml.HTMLDocument)fra.document;
                    eles = Doc.getElementsByTagName("input");

                    foreach (mshtml.IHTMLElement ele in eles)
                    {
                        try
                        {
                            if (ele.getAttribute("value").ToString() == "継 続")
                            {
                                ele.click();
                                Pub_Com.SleepAndWaitComplete(Ie);
                                goto endOfForLoop;
                            }

                        }
                        catch (Exception)
                        {

                        }

                    }
                endOfForLoop:
                    1.GetHashCode(); //VBConversions note: C# requires an executable line here, so a dummy line was added.
                }
                catch (Exception)
                {

                }
                AddProBar(lv2); //15

                return;
            }
            catch (Exception)
            {
                //DoStep1_SinkiMitumori(事業所, 得意先, 下店, 現場名, 備考, 日付連番, fl)
                return;
            }
        }

        //Step 2 （WHILE）
        public bool DoStep2_Set(int startIdx = 0)
        {

            mshtml.HTMLDocument Doc = default(mshtml.HTMLDocument);
            mshtml.IHTMLElementCollection eles = default(mshtml.IHTMLElementCollection);
            Doc = (mshtml.HTMLDocument)Ie.Document;
            mshtml.HTMLWindow2 fra = Pub_Com.GetFrameWait(ref Ie, "fraHyou");

            while (ReferenceEquals(fra, null))
            {
                Com.Sleep5(1000);
                fra = Pub_Com.GetFrameWait(ref Ie, "fraHyou");
            }
            Pub_Com.SleepAndWaitComplete(Ie);
            Doc = (mshtml.HTMLDocument)fra.document;
            eles = Doc.getElementsByTagName("input");

            Pub_Com.SleepAndWaitComplete(Ie);
            Com.Sleep5(1000);
            if (eles.length == 0)
            {
                MessageBox.Show("明細がありません、多分発注した件数多いです");

                Ie.Visible = true;
                System.Environment.Exit(0);
                return false;
            }

            for (int i = startIdx; i <= eles.length - 1; i++)
            {
                mshtml.IHTMLElement ele = (mshtml.IHTMLElement)(eles.item(i));

                if (ReferenceEquals(ele, null))
                {
                    continue;
                }

                try
                {
                    if (ele.getAttribute("name").ToString() == "strMitKbnHen")
                    {
                        mshtml.IHTMLTableRow tr = (mshtml.IHTMLTableRow)ele.parentElement.parentElement;
                        mshtml.HTMLTableCell td = (mshtml.HTMLTableCell)(tr.cells.item(4));

                        if (td.innerText == "作成中")
                        {
                            ele.click();
                            Pub_Com.GetElementBy(ref Ie, "fraHead", "input", "value", "発注納期非表示").click();
                            Pub_Com.SleepAndWaitComplete(Ie);


                            mshtml.HTMLDocument Doc1 = (mshtml.HTMLDocument)Ie.Document;
                            mshtml.HTMLWindow2 fra1 = Pub_Com.GetFrameWait(ref Ie, "fraMitBody");
                            Doc1 = (mshtml.HTMLDocument)fra1.document;

                            if (Doc1.body.innerText.IndexOf("発注可能な見積ではありません") > 0)
                            {
                                Ie.GoBack();
                                Pub_Com.SleepAndWaitComplete(Ie);
                                return DoStep2_Set(i + 1);
                            }
                            else
                            {
                                return true;
                            }



                        }
                    }
                }
                catch (Exception)
                {

                }


            }

            return false;

        }

        //POPUP Window 取得
        public SHDocVw.InternetExplorerMedium GetPopupWindow(string titleKey, string fileNameKey)
        {
            SHDocVw.ShellWindows ShellWindows = new SHDocVw.ShellWindows();
            foreach (SHDocVw.InternetExplorerMedium childIe in ShellWindows)
            {
                System.Windows.Forms.Application.DoEvents();
                string filename = System.Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(childIe.FullName).ToLower());
                if (filename == "iexplore")
                {
                    if (childIe.LocationURL.Contains(fileNameKey))
                    {
                        if (((mshtml.HTMLDocument)childIe.Document).title == titleKey)
                        {
                            if (((mshtml.HTMLDocument)childIe.Document).url.Contains(fileNameKey))
                            {
                                //childIe.Visible = False
                                Pub_Com.SleepAndWaitComplete(childIe);
                                Com.Sleep5(500);
                                ShellWindows = null;
                                //childIe.Visible = IeVisible
                                return childIe;
                            }
                        }
                    }
                }
            }
            ShellWindows = null;
            return null;
        }


        #endregion
	

        public AutoImportNouhinsyo()
        {
            InitializeComponent();
        }

        private void AutoImportNouhinsyo_Load(object sender, EventArgs e)
        {

        }
    }
}
