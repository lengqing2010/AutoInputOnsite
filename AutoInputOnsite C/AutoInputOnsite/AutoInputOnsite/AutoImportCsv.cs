using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace AutoInputOnsite
{
    public partial class AutoImportCsv : Form
    {

        public double _ProBar = 0;
        public double lv1;
        public double lv2;
        public bool IeVisible = true;
        public int ProBar
        {
            get
            {
                if (_ProBar < 100)
                {
                    return System.Convert.ToInt32(_ProBar);
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
        public void AddProBar(double x)
        {
            _ProBar += x;
        }

        public AutoImportCsv()
        {
            InitializeComponent();
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
		
		    Pub_Com = new Com("見出明細作成" + DateTime.Now.ToString("yyyyMMddHHmmss"));
		    if (Pub_Com.file_list_hattyuu.Count == 0)
		    {
			    ProBar = 100;
			    return;
		    }
		
		    lv1 = System.Convert.ToDouble(90 / Pub_Com.file_list_hattyuu.Count);
		
		    object authHeader = "Authorization: Basic " + 
			    Convert.ToBase64String(System.Text.UnicodeEncoding.UTF8.GetBytes(string.Format("{0}:{1}", Pub_Com.user, Pub_Com.password))) + "\\r\\n";

		    //＊＊＊ OnSiteパスワード入力画面
            Ie.Navigate(Pub_Com.url, null, null, null, authHeader);
		    Ie.Silent = true;
		    Ie.Visible = IeVisible;

            //＊＊＊ログイン １
		    ProBar = 5;
		    DoStep1_Login();

            //＊＊＊CSV 読み ２
		    ProBar = 10;
		    foreach (string fl in Pub_Com.file_list_hattyuu)
		    {
			
			    lv2 = lv1 / 10;
			    Pub_Com.AddMsg("取込：" + fl);

                string 事業所 = System.Convert.ToString(fl.Split('-')[0]);
                string 得意先 = System.Convert.ToString(fl.Split('-')[1]);
                string 下店 = System.Convert.ToString(fl.Split('-')[2]);
                string 現場名 = System.Convert.ToString(fl.Split('-')[3]);
                string 備考 = System.Convert.ToString(fl.Split('-')[4]);
                string 日付連番 = System.Convert.ToString(fl.Split('-')[5]);

			    //While
			    DoStep2_Sinki(事業所, 得意先, 下店, 現場名, 備考, 日付連番, fl);			
			
			    Pub_Com.AddMsg("移動CSV：" + fl + "→" + Pub_Com.folder_Hattyuu_kanryou);

                Com.MoveFile(Pub_Com.folder_Hattyuu + fl, Pub_Com.folder_Hattyuu_kanryou + fl);

                AddProBar(lv2); //10
	
		    }

		    ProBar = 100;
		
	    }

        //Step 1 LOGIN IN
        public void DoStep1_Login()
        {


            Pub_Com.AddMsg("OnSiteパスワード入力");
            //＊＊＊ OnSiteパスワード入力
            Pub_Com.GetElementBy(ref Ie, "", "input", "name", "strPassWord").setAttribute("value", ConfigurationManager.AppSettings["OnSitePassword"].ToString());
            Pub_Com.GetElementBy(ref Ie, "", "input", "value", "ログオン").click();
            Pub_Com.SleepAndWaitComplete(Ie);

            //
            Pub_Com.AddMsg("業務別総合メニュー");
            //'＊＊＊ 業務別総合メニュー
            Pub_Com.GetElementBy(ref Ie, "SubHeader", "a", "innertext", "[見積]").click();
            Pub_Com.SleepAndWaitComplete(Ie);


            Pub_Com.AddMsg("物販明細");


            //'＊＊＊ 物販明細
            Pub_Com.GetElementBy(ref Ie, "Main", "input", "value", "物販明細").click();
            Pub_Com.SleepAndWaitComplete(Ie);

            //新規見積もり
            Pub_Com.AddMsg("新規見積");
            DoStep1_SinkiMitumori();

        }

        //Step 1 新規見積もり
        public void DoStep1_SinkiMitumori()
        {

            SHDocVw.InternetExplorerMedium cIe = GetPopupWindow("OnSite", "mitSearch.asp");
            while (ReferenceEquals(cIe, null))
            {
                Com.Sleep5(1000);
                cIe = GetPopupWindow("OnSite", "mitSearch.asp");
            }

            Pub_Com.GetElementBy(ref cIe, "", "input", "value", "新規見積").click();
            Com.Sleep5(100);

            Pub_Com.SleepAndWaitComplete(Ie);

        }

        //Step 2 （WHILE）
        public void DoStep2_Sinki(string 事業所, string 得意先, string 下店, string 現場名, string 備考, string 日付連番, string fl)
        {
            //POPUP 画面内容入力
            string folder_Hattyuu = Com.GetAppSetting("Folder_Hattyuu");
            Com.Sleep5(500);
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.AddMsg("    事業所：" + 事業所);
            Pub_Com.AddMsg("    得意先：" + 得意先);
            Pub_Com.AddMsg("    下店：" + 下店);
            Pub_Com.AddMsg("    備考：" + 備考);
            Pub_Com.AddMsg("    現場名：" + 現場名);

            Pub_Com.SleepAndWaitComplete(Ie);
            mshtml.HTMLWindow2 fra = Pub_Com.GetFrameWait(ref Ie, "fraMitBody");
            Pub_Com.GetElement(fra, "input", "name", "strJgyCdText").innerText = 事業所;
            Pub_Com.GetElement(fra, "input", "name", "strTokMeiText").innerText = 得意先;
            Pub_Com.GetElement(fra, "input", "name", "strOtdMeiText").innerText = 下店;
            Pub_Com.GetElement(fra, "input", "name", "strBikouMei").innerText = 備考;
            Pub_Com.GetElement(fra, "input", "name", "strGenbaMei").innerText = 現場名;
            Pub_Com.GetElement(fra, "select", "name", "aryKijyunSyouhinBunrui").setAttribute("value", "A0001,サッシ,L90000");

            Pub_Com.AddMsg("    納材店なしで内訳入力へ CLICK");
            Pub_Com.GetElement(fra, "input", "name", "btnUtiwake").click();

            AddProBar(lv2); 

            Com.Sleep5(500);
            Pub_Com.SleepAndWaitComplete(Ie);


            //見積内訳入力
            try
            {
                mshtml.IHTMLElement ele1 = Pub_Com.GetElementBy(ref Ie, "fraMitBody", "DIV", "classname", "ttl");
                if (ele1.innerText != "見積内訳入力")
                {
                    DoStep2_Sinki(事業所, 得意先, 下店, 現場名, 備考, 日付連番, fl);
                    return;
                }
                Pub_Com.AddMsg("    見積内訳入力 CSV取込 CLICK");
                Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "CSV取込").click();
            }
            catch (Exception)
            {
                DoStep2_Sinki(事業所, 得意先, 下店, 現場名, 備考, 日付連番, fl);
                return;
            }

            AddProBar(lv2); //2
            Pub_Com.SleepAndWaitComplete(Ie);
            Com.Sleep5(500);

            Pub_Com.AddMsg("    見積内訳入力 CSV取込 参　照 CLICK");
            SHDocVw.InternetExplorerMedium fra1 = GetPopupWindow("OnSite", "fileYomikomiSiji.asp");
            while (ReferenceEquals(fra1, null))
            {
                fra1 = GetPopupWindow("OnSite", "fileYomikomiSiji.asp");
            }
            Pub_Com.GetElementBy(ref fra1, "", "input", "value", "参　照").click();


            Com.Sleep5(1500);
            try
            {
                while (Pub_Com.GetElementBy(ref fra1, "", "input", "name", "strFilename").getAttribute("value").ToString() == "")
                {
                    Com.Sleep5(1);
                }
            }
            catch (Exception)
            {
            }


            Com.Sleep5(1500);
            AddProBar(lv2); //3

            Pub_Com.AddMsg("    見積内訳入力 CSV取込 取　込 CLICK");
            SHDocVw.InternetExplorerMedium popupPage;
            popupPage =GetPopupWindow("OnSite", "fileYomikomiSiji.asp");
            Pub_Com.GetElementBy(ref popupPage, "", "input", "value", "取　込").click();
            Com.Sleep5(1000);
            AddProBar(lv2); //4

            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.AddMsg("    商品コード複数入力 次　へ CLICK");
            Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.AddMsg("    商品コード複数入力 次　へ CLICK");
            Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.AddMsg("    商品コード複数入力 次　へ CLICK");
            Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
            AddProBar(lv2); //5
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.SleepAndWaitComplete(Ie);

            //寸法入力

            if (ReferenceEquals(Pub_Com.GetElementByDo(ref Ie, "fraMitBody", "input", "value", "見積内訳入力へ"), null))
            {
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.AddMsg("    寸法入力 次　へ CLICK");
                Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.AddMsg("    寸法入力 次　へ CLICK");
                Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.AddMsg("    寸法入力 次　へ CLICK");
                Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "次　へ").click();
                Pub_Com.SleepAndWaitComplete(Ie);
                Pub_Com.SleepAndWaitComplete(Ie);
            }

            AddProBar(lv2); //6
            Pub_Com.AddMsg("    単価入力 見積内訳入力へ CLICK");
            Pub_Com.GetElementBy(ref Ie, "fraMitBody", "input", "value", "見積内訳入力へ").click();
            Pub_Com.SleepAndWaitComplete(Ie);
            AddProBar(lv2); //7
            mshtml.IHTMLElement ele = Pub_Com.GetElementBy(ref Ie, "fraMitBody", "DIV", "classname", "ttl");
            string kekka = System.Convert.ToString(ele.innerText);
            AddProBar(lv2); //8
            Pub_Com.AddMsg("    新規見積 CLICK");
            Pub_Com.GetElementBy(ref Ie, "fraMitMenu", "a", "innertext", "[新規見積]").click();
            Pub_Com.SleepAndWaitComplete(Ie);
            Pub_Com.SleepAndWaitComplete(Ie);
            AddProBar(lv2); //9


        }

        //POPUP Window 取得
        public SHDocVw.InternetExplorerMedium GetPopupWindow(string titleKey, string fileNameKey)
        {
            SHDocVw.ShellWindows ShellWindows = new SHDocVw.ShellWindows();
            foreach (SHDocVw.InternetExplorerMedium childIe in ShellWindows)
            {
                string filename = System.Convert.ToString(System.IO.Path.GetFileNameWithoutExtension(Ie.FullName).ToLower());
                if (filename == "iexplore")
                {
                    if (childIe.LocationURL.Contains(fileNameKey))
                    {

                        if (((mshtml.HTMLDocument)childIe.Document).title.Contains("資格情報が無効"))
                        {
                            MessageBox.Show("資格情報が無効");
                            System.Environment.Exit(0);
                        }
                        if (((mshtml.HTMLDocument)childIe.Document).title == titleKey)
                        {
                            if (((mshtml.HTMLDocument)childIe.Document).url.Contains(fileNameKey))
                            {
                                Pub_Com.SleepAndWaitComplete(childIe);
                                Com.Sleep5(500);

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

        private void AutoImportCsv_Load(object sender, EventArgs e)
        {

        }
    }
}
